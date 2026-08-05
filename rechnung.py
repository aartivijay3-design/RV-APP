"""Confirmation Generator + Kalender tab: parses a Rechnung/Confirmation
document (hotels, flights, client/travel data) and builds the branded
Confirmation Word document. parse_rechnung_with_ai() is shared by both
/generate-confirmation and /parse-calendar-info."""
import json
import re
import zipfile
import io
from pathlib import Path
from typing import Optional

from fastapi import HTTPException, UploadFile

from ai_client import _ai_complete, AI_MODEL
from extraction import _legacy_read_excel, _legacy_read_word, _legacy_read_pdf, _read_pdf_via_vision

CONF_EXTRACT_PROMPT = """Extract the following information from this German travel invoice (Rechnung) and return ONLY valid JSON, no markdown.

{
  "client_names": ["Herr Andreas Claudio Groh", "Frau Bettina Martha Groh"],
  "destination_de": "Indien",
  "destination_en": "India",
  "travel_start": "22.01.2026",
  "travel_end": "29.01.2026",
  "pax": 2,
  "hotels": [
    {
      "check_in": "22.01.2026",
      "check_out": "24.01.2026",
      "nights": 2,
      "hotel_name": "The Imperial",
      "city": "New Delhi",
      "room_type": "Heritage Room",
      "meal_plan_de": "Frühstück",
      "meal_plan_en": "breakfast"
    }
  ]
}

Rules:
- client_names: include salutation (Herr/Frau/Mr/Mrs) and full name
- nights: calculate automatically from check_in to check_out
- meal_plan_en: breakfast | halfboard | fullboard | room only
- destination_en: translate destination to English
- Return every hotel exactly as it appears in the Rechnung
"""


def _nights_between(d1: str, d2: str) -> int:
    """'22.01.2026', '24.01.2026' → 2"""
    from datetime import datetime
    try:
        fmt = "%d.%m.%Y"
        return (datetime.strptime(d2, fmt) - datetime.strptime(d1, fmt)).days
    except Exception:
        return 1


def _meal_en(meal_de: str) -> str:
    m = meal_de.lower()
    if "vollpension" in m or "full" in m:     return "fullboard"
    if "halbpension" in m or "half" in m:     return "halfboard"
    if "frühstück" in m or "breakfast" in m:  return "breakfast"
    if "fruhstuck" in m:                      return "breakfast"
    return "room only"


from datetime import datetime, timedelta

# German month name → number
_DE_MONTHS = {
    "januar":1,"februar":2,"märz":3,"april":4,"mai":5,"juni":6,
    "juli":7,"august":8,"september":9,"oktober":10,"november":11,"dezember":12,
    "jan":1,"feb":2,"mär":3,"apr":4,"jun":6,"jul":7,"aug":8,
    "sep":9,"okt":10,"nov":11,"dez":12,
}

DEST_MAP = {
    "indien":"India","india":"India","japan":"Japan","thailand":"Thailand",
    "sri lanka":"Sri Lanka","vietnam":"Vietnam","indonesien":"Indonesia",
    "china":"China","peru":"Peru","marokko":"Morocco","ägypten":"Egypt",
    "tansania":"Tanzania","kenya":"Kenya","kenia":"Kenya","namibia":"Namibia",
    "jordanien":"Jordan","türkei":"Turkey","griechenland":"Greece",
    "portugal":"Portugal","spanien":"Spain","nepal":"Nepal","bhutan":"Bhutan",
    "myanmar":"Myanmar","kambodscha":"Cambodia","laos":"Laos",
    "malaysia":"Malaysia","singapur":"Singapore","usa":"USA","mexiko":"Mexico",
    "kolumbien":"Colombia","brasilien":"Brazil","argentinien":"Argentina",
    "südafrika":"South Africa","äthiopien":"Ethiopia","tansania":"Tanzania",
    "malediven":"Maldives","seychellen":"Seychelles","mauritius":"Mauritius",
    "kuba":"Cuba","costa rica":"Costa Rica",
}


def _dest_en(dest_de: str) -> str:
    dl = dest_de.lower()
    return next((v for k, v in DEST_MAP.items() if k in dl), dest_de)


def _parse_de_date(date_str: str) -> str:
    """Parse various German date strings → 'DD.MM.YYYY'.
    Handles: '24. Mai 2026', '24.05.2026', '24.05.26'
    """
    date_str = date_str.strip()
    # Already numeric: 24.05.2026 or 24.05.26
    m = re.match(r"(\d{1,2})\.(\d{2})\.(\d{2,4})$", date_str)
    if m:
        d, mo, y = m.group(1), m.group(2), m.group(3)
        if len(y) == 2: y = "20" + y
        return f"{int(d):02d}.{mo}.{y}"
    # German month name: "24. Mai 2026"
    m = re.match(r"(\d{1,2})\.\s*([A-Za-zÄÖÜäöü]+)\.?\s*(\d{4})", date_str)
    if m:
        d = int(m.group(1))
        mo_name = m.group(2).lower()
        mo = _DE_MONTHS.get(mo_name, 1)
        y = m.group(3)
        return f"{d:02d}.{mo:02d}.{y}"
    return date_str


def _add_days(date_str: str, n: int) -> str:
    """'24.05.2026' + 4 → '28.05.2026'"""
    try:
        dt = datetime.strptime(date_str, "%d.%m.%Y")
        return (dt + timedelta(days=n)).strftime("%d.%m.%Y")
    except Exception:
        return date_str


from quote_parser import normalize_quotes as _normalize_quotes
from quote_parser import parse_leistungen_hotels as _parse_leistungen_hotels_impl


def _parse_leistungen_hotels(text: str, travel_start: str) -> list:
    return _parse_leistungen_hotels_impl(text, travel_start)




def _parse_itinerary_hotels(text: str) -> list:
    """
    Parse hotels from itinerary body: 'Übernachtung im [Hotel]' + preceding day dates.
    'Samstag, 24.05.2026' lines give the check-in date.
    """
    # Find all day-date pairs
    day_dates = re.findall(
        r"(?:Montag|Dienstag|Mittwoch|Donnerstag|Freitag|Samstag|Sonntag),?\s+"
        r"(\d{2}\.\d{2}\.\d{4})",
        text
    )
    # Find all hotel overnight lines with their position
    overnight_matches = list(re.finditer(
        r"[Üü]bernachtung\s+im\s+(.+?)(?:\n|$)", text
    ))

    if not day_dates or not overnight_matches:
        return []

    hotels = []
    seen = {}  # hotel name → first check-in

    for match in overnight_matches:
        raw_name = match.group(1).strip()
        # Remove URLs and extra whitespace
        hotel_name = re.split(r"https?://", raw_name)[0].strip()
        if not hotel_name or len(hotel_name) < 3:
            continue

        # Find the day-date that precedes this position
        pos = match.start()
        text_before = text[:pos]
        preceding_dates = re.findall(
            r"(?:Montag|Dienstag|Mittwoch|Donnerstag|Freitag|Samstag|Sonntag),?\s+"
            r"(\d{2}\.\d{2}\.\d{4})",
            text_before
        )
        if not preceding_dates:
            continue
        check_in = preceding_dates[-1]

        if hotel_name not in seen:
            seen[hotel_name] = check_in
            hotels.append({
                "check_in":    check_in,
                "check_out":   "",   # filled in next pass
                "nights":      0,
                "hotel_name":  hotel_name,
                "city":        "",
                "room_type":   "",
                "meal_plan_de": "Frühstück",
                "meal_plan_en": "breakfast",
            })

    # Calculate nights and check-out from consecutive check-ins
    for i, h in enumerate(hotels):
        if i + 1 < len(hotels):
            h["check_out"] = hotels[i + 1]["check_in"]
        else:
            # Last hotel: use travel end if available or check_in + 1
            h["check_out"] = _add_days(h["check_in"], 1)
        h["nights"] = _nights_between(h["check_in"], h["check_out"])

    return hotels


def parse_document(text: str) -> dict:
    """
    Universal parser — handles Rechnung, Reiseverlauf Leistungen, and full itinerary.
    Tries each strategy and returns the first that yields hotels.
    """
    result = {
        "client_names": [],
        "destination_de": "",
        "destination_en": "",
        "travel_start": "",
        "travel_end": "",
        "pax": 2,
        "hotels": [],
    }

    # ── 1. Destination ────────────────────────────────────────────
    # Rechnung: "Reise von DD.MM.YYYY bis DD.MM.YYYY: Destination"
    dest_m = re.search(
        r"Reise\s+von\s+(\d{2}\.\d{2}\.\d{4})\s+bis\s+(\d{2}\.\d{2}\.\d{4}):\s*([^\n\r]+)",
        text, re.IGNORECASE
    )
    if dest_m:
        result["travel_start"]   = dest_m.group(1)
        result["travel_end"]     = dest_m.group(2)
        result["destination_de"] = dest_m.group(3).strip()

    # Reiseverlauf: look for all-caps destination line e.g. "JAPAN", "INDIEN"
    if not result["destination_de"]:
        caps_m = re.search(
            r"\n([A-ZÄÖÜ][A-ZÄÖÜ\s&]{2,30})\n",
            text
        )
        if caps_m:
            candidate = caps_m.group(1).strip()
            if candidate not in ("ENDE DER REISE", "IHR PERSÖNLICHER REISEVERLAUF",
                                  "LEISTUNGSÜBERSICHT", "CONFIRMATION"):
                result["destination_de"] = candidate.title()

    # Reisedatum line: "Reisedatum: 24. Mai – 06. Juni 2026"
    if not result["travel_start"]:
        rd_m = re.search(
            r"Reisedatum[:\t\s]+(.+?)\s*[–\-]\s*(.+?)(?:\n|$)",
            text, re.IGNORECASE
        )
        if rd_m:
            start_str = rd_m.group(1).strip()
            end_str   = rd_m.group(2).strip()
            # If start has no year but end does, add end's year to start
            if not re.search(r"\d{4}", start_str) and re.search(r"(\d{4})", end_str):
                year = re.search(r"(\d{4})", end_str).group(1)
                start_str = start_str + " " + year
            result["travel_start"] = _parse_de_date(start_str)
            result["travel_end"]   = _parse_de_date(end_str)

    result["destination_en"] = _dest_en(result["destination_de"])

    # ── 2. Client names ───────────────────────────────────────────
    # Rechnung style: "Herr   Andreas Claudio Groh"
    client_matches = re.findall(
        r"(Herr|Frau)\s{1,10}([A-ZÄÖÜ][a-zA-ZÄÖÜäöü\-\s]{3,50}?)(?=\n|Herr|Frau|Vorgang|Reise|$)",
        text
    )
    for sal, name in client_matches:
        full = f"{sal} {name.strip()}"
        if full not in result["client_names"] and len(name.strip()) > 3:
            result["client_names"].append(full)

    # Reiseverlauf style: "Reiseteilnehmer:\tName1\nName2\nName3 14J."
    if not result["client_names"]:
        rt_m = re.search(
            r"Reiseteilnehmer[:\t\s]+(.+?)(?=Reisedaten?|Reisepreis|Inkludierte|$)",
            text, re.DOTALL | re.IGNORECASE
        )
        if rt_m:
            block = rt_m.group(1)
            names = [n.strip() for n in re.split(r"[\t\n]", block) if n.strip()]
            for n in names:
                # Strip age suffixes like "14J." or "17J."
                clean = re.sub(r"\s+\d+\s*[Jj]\.?\s*$", "", n).strip()
                # Skip lines that look like dates or labels (contain digits heavily)
                if re.search(r"\d{4}|\d+\.\s*(Januar|Februar|März|April|Mai|Juni|Juli|August|September|Oktober|November|Dezember)", clean, re.IGNORECASE):
                    continue
                if clean and len(clean) > 2:
                    result["client_names"].append(clean)

    result["pax"] = max(len(result["client_names"]), 1)

    # ── 3. Hotels — try all three strategies ─────────────────────

    # Strategy A: Rechnung hotel table (explicit check-in/check-out dates)
    hotel_block_m = re.search(
        r"Hotel\s+[üu]ber\s+Bawa(.+?)(?:Vorgangspreis|Gesamtpreis|$)",
        text, re.DOTALL | re.IGNORECASE
    )
    if hotel_block_m:
        block = hotel_block_m.group(1)
        date_pairs = re.findall(
            r"(\d{2}\.\d{2}\.\d{4})\s*\n\s*(\d{2}\.\d{2}\.\d{4})\s*\n\s*(.+?)(?=\d{2}\.\d{2}\.\d{4}|$)",
            block, re.DOTALL
        )
        for ci, co, rest in date_pairs:
            raw_lines   = [l.strip() for l in rest.strip().split("\n") if l.strip()]
            notes_lines = [l for l in raw_lines if l.lower().startswith("bemerkung")]
            lines       = [l for l in raw_lines if not l.lower().startswith("bemerkung")]
            if not lines:
                continue
            hotel_name = lines[0]
            city       = lines[1].split(",")[0].strip() if len(lines) > 1 else ""
            room_meal  = " ".join(lines[2:]) if len(lines) > 2 else ""
            if "," in room_meal:
                room_type, meal_de = room_meal.split(",", 1)
                room_type = room_type.strip(); meal_de = meal_de.strip()
            else:
                room_type = room_meal; meal_de = "Frühstück"
            notes = re.sub(r"^Bemerkung:\s*", "", notes_lines[0], flags=re.IGNORECASE) if notes_lines else ""
            result["hotels"].append({
                "check_in": ci, "check_out": co,
                "nights": _nights_between(ci, co),
                "hotel_name": hotel_name, "city": city,
                "room_type": room_type, "meal_plan_de": meal_de,
                "meal_plan_en": _meal_en(meal_de),
                "checkin_note": "", "checkout_note": "", "notes": notes,
            })

    # Strategy B: Reiseverlauf Leistungen bullets
    # "X Uebernachtungen in City im [Hotel] in Room inklusive Meal"
    if not result["hotels"]:
        start = result.get("travel_start", "01.01.2026") or "01.01.2026"
        result["hotels"] = _parse_leistungen_hotels(text, start)

    # Strategy C: Itinerary body — "Uebernachtung im Hotel" lines + day headings
    # Always run and use whichever strategy found MORE hotels
    hotels_c = _parse_itinerary_hotels(text)
    if len(hotels_c) > len(result["hotels"]):
        # Enrich C results with room type / meal from B where possible
        b_by_name = {h["hotel_name"]: h for h in result["hotels"]}
        for hc in hotels_c:
            if hc["hotel_name"] in b_by_name:
                hc["room_type"]   = b_by_name[hc["hotel_name"]].get("room_type", "")
                hc["meal_plan_de"] = b_by_name[hc["hotel_name"]].get("meal_plan_de", "Fruehstueck")
                hc["meal_plan_en"] = b_by_name[hc["hotel_name"]].get("meal_plan_en", "breakfast")
        result["hotels"] = hotels_c

    return result


# Keep old name as alias
def parse_rechnung_regex(text: str) -> dict:
    return parse_document(text)


_CONF_PARSE_PROMPT = """You extract hotel and flight booking data from a BAWA Tours travel document (Leistungsübersicht).

Return ONLY a valid JSON object with exactly these fields:
{
  "client_names": ["Full Name 1", "Full Name 2"],
  "destination_en": "Country/Region in English",
  "destination_de": "Country/Region in German",
  "travel_start": "DD.MM.YYYY",
  "travel_end": "DD.MM.YYYY",
  "hotels": [
    {
      "check_in": "DD.MM.YYYY",
      "check_out": "DD.MM.YYYY",
      "nights": 2,
      "hotel_name": "Hotel Name",
      "city": "City",
      "room_type": "Room Type Description",
      "meal_plan_en": "breakfast",
      "checkout_note": "Late check-out until 18:00",
      "checkin_note": "",
      "notes": "Guaranteed upgrade to Corner Suite; USD 100 F&B credit"
    }
  ],
  "flights": [
    {
      "date": "DD.MM.YYYY",
      "arrival_date": "DD.MM.YYYY",
      "airline": "Japan Airlines",
      "flight_number": "JL 407",
      "departure_airport": "Frankfurt (FRA)",
      "arrival_airport": "Tokyo Haneda (HND)",
      "departure_time": "13:20",
      "arrival_time": "08:15",
      "notes": "Business Class; 1 stop via Helsinki"
    }
  ],
  "inclusions": [
    {"text": "All transfers by private air-conditioned vehicle according to itinerary", "hotel": ""},
    {"text": "Private Sightseeing with German Speaking accompanying Guides incl. entrance fees", "hotel": ""},
    {"text": "Example hotel-specific inclusion translated to English", "hotel": "Hotel Name, City"}
  ]
}

Rules:
- If the given text does NOT look like a real BAWA travel/booking document — e.g. it's empty, unrelated content, or too garbled/incomplete to reliably identify a client name and destination — return EXACTLY {"error": "insufficient_data"} and nothing else. Do NOT invent a placeholder client name (e.g. "Mustermann"/"Max Mustermann") or guess a destination when you cannot genuinely determine one from the text — a wrong calendar entry is worse than a clear failure.
- meal_plan_en must be one of: breakfast, halfboard, fullboard, room only
- Derive check_in/check_out from the travel_start date + cumulative nights
- If dates are missing, use 01.01.2026 as travel_start and count forward
- travel_start/travel_end: a document's intro sentence ("für Ihre Reise von X bis Y") can contain a typo'd date that conflicts with the actual flight/hotel dates elsewhere in the same document (seen in practice: an intro sentence read "18.09.2021" while the first outbound flight and every hotel check-in/check-out in the same document consistently used 2026, and the trip actually started on an earlier date matching the first flight). When the intro sentence's date conflicts with the first outbound flight's departure date or the earliest hotel check-in, trust the flight/hotel dates instead — multiple consistent structured entries are more reliable than one prose sentence.
- Characters shown as '?' may be umlauts (ü/ö/ä/Ü) or quote marks — infer from context
- Extract ALL hotels mentioned in the document. A document may have NO hotels at all (flights-only itinerary) — in that case return an empty hotels list, do not invent one.
- Some bookings are NOT a hotel but a boat/yacht charter — recognizable from wording like "ab/bis [Marina]" (from/to a marina, i.e. the charter's start and return point) and a vessel name/model instead of a hotel name (e.g. "Catamaran Lagoon 42"). Treat this exactly like a hotel entry in the `hotels` list: hotel_name = the vessel name, city = the marina/port name (e.g. "ab/bis Marmaris Marina" → city "Marmaris"), nights/check_in/check_out as given. Leave room_type and meal_plan_en empty rather than inventing one if the document gives no room/meal-plan equivalent for the charter.
- checkout_note: if a late check-out time is mentioned for a hotel, write in English (e.g. "Late check-out until 18:00"). Otherwise empty string.
- checkin_note: if early check-in is mentioned for a hotel, write in English. Otherwise empty string.
- notes: any special remark tied to a specific hotel — room upgrades (e.g. "Bemerkung: garantiertes Upgrade Corner Suite"), F&B/dining credits, complimentary amenities. Translate to English. Empty string if none.
- flights: extract EVERY flight leg mentioned (international or domestic) — not local car/airport transfers.
  * date: departure date. arrival_date: only set if different from date (overnight/red-eye flight), otherwise same as date.
  * flight_number and airline: empty string if not stated — never invent one.
  * departure_airport / arrival_airport: airport or city name as given, with IATA code in parentheses if stated.
  * departure_time / arrival_time: 24h "HH:MM" if stated, otherwise empty string.
  * notes: cabin class, layovers/stops, seat info — empty string if none.
  * If the document has no flight details at all, return an empty flights list — do not invent one.
- inclusions: each item has "text" (in English) and "hotel" fields.
  * ALWAYS include the two standard items first with hotel="" (transfers + sightseeing/guides).
  * For every extra service listed AFTER a hotel line in the source (minibar, spa, activities, meet & greet, etc.) — add it with hotel="Hotel Name, City" matching the hotel it belongs to. Translate to clear English.
  * Anything that follows hotel line X and comes before hotel line X+1 belongs to hotel X.
  * Do NOT include room type, meal plan, or price — those go elsewhere.
  * late/early checkout goes in checkout_note/checkin_note on the hotel, NOT in inclusions.
- Return ONLY JSON, no markdown, no explanation"""


def _parse_conf_with_ai(text: str) -> dict:
    """Use Gemini to extract hotel/flight data when regex parsing fails."""
    resp = _ai_complete(
        model=AI_MODEL,
        temperature=0.1,
        max_tokens=8000,
        messages=[
            {"role": "system", "content": _CONF_PARSE_PROMPT},
            {"role": "user",   "content": text[:30000]},
        ],
    )
    raw = resp.choices[0].message.content.strip()
    # Strip markdown fences if present
    raw = re.sub(r"^```[a-z]*\n?", "", raw)
    raw = re.sub(r"\n?```$", "", raw)
    # strict=False tolerates a raw control character (e.g. a literal newline)
    # inside a string value instead of the escaped \n — the AI occasionally
    # copies a multi-line note verbatim like that, which strict JSON parsing
    # otherwise rejects even though the structure itself is valid.
    return json.loads(raw, strict=False)



# A model that can't reliably read the document sometimes invents a
# plausible-looking result instead of admitting it — the standard German
# placeholder name ("Max Mustermann", used on ID card templates and forms)
# is a dead giveaway this happened despite the prompt's explicit
# instruction not to. Checked as a belt-and-suspenders safety net in
# addition to the {"error": "insufficient_data"} escape hatch, since
# prompt compliance isn't 100% guaranteed.
_PLACEHOLDER_NAME_PATTERNS = ("mustermann", "musterfrau", "john doe", "jane doe", "max mustermann")


def _looks_like_placeholder(client_names) -> bool:
    return any(
        p in (name or "").lower()
        for name in (client_names or [])
        for p in _PLACEHOLDER_NAME_PATTERNS
    )


def parse_rechnung_with_ai(text: str, require_hotels: bool = True) -> dict:
    """
    Parse a BAWA Reiseverlauf / Leistungsübersicht / Confirmation document.
    Uses regex-based parse_document first for fallback metadata; Gemini is
    the primary source (handles encoding-broken source docs, flights, remarks).

    require_hotels: True for Confirmation-document generation, which needs a
    hotel table. Pass False (calendar tab) to also accept flights-only trips.
    """
    parsed = parse_document(text)
    hotels = parsed.get("hotels", [])
    parsed.setdefault("flights", [])

    # Always run the AI to get inclusions, flights, checkout notes, and
    # accurate metadata. Regex result is kept as fallback if AI fails —
    # including when the AI "succeeds" but the result looks fabricated
    # (see _looks_like_placeholder), which used to slip through here and
    # produce a wrong calendar entry ("Mustermann - Japan") instead of a
    # clear error when the source document couldn't actually be read.
    try:
        ai_parsed = _parse_conf_with_ai(text)
        if ai_parsed.get("error"):
            raise ValueError("AI reported insufficient data in the source document")
        if _looks_like_placeholder(ai_parsed.get("client_names")):
            raise ValueError("AI returned a placeholder name — likely hallucinated from unreadable input")
        if ai_parsed.get("hotels") or ai_parsed.get("flights"):
            # Merge: fill any AI gaps from regex
            for key in ("client_names", "destination_en", "destination_de",
                        "travel_start", "travel_end"):
                if not ai_parsed.get(key) and parsed.get(key):
                    ai_parsed[key] = parsed[key]
            # If regex found MORE hotels than AI, prefer the regex hotel list
            # but keep AI's inclusions and metadata
            if len(parsed.get("hotels", [])) > len(ai_parsed.get("hotels", [])):
                ai_parsed["hotels"] = parsed["hotels"]
            ai_parsed.setdefault("flights", [])
            return ai_parsed
    except Exception:
        pass  # AI failed — fall through to regex result

    if require_hotels and not hotels:
        raise ValueError(
            "Konnte die Rechnung nicht lesen. Bitte fügen Sie den Text der letzten Seite des "
            "Reiseverlaufs (Leistungsübersicht) ein oder laden Sie eine Confirmation hoch."
        )
    if not require_hotels and not hotels and not parsed["flights"]:
        raise ValueError(
            "Konnte die Rechnung nicht lesen. Bitte fügen Sie den Text der letzten Seite des "
            "Reiseverlaufs (Leistungsübersicht) ein oder laden Sie eine Confirmation hoch."
        )
    return parsed


# ── Confirmation Word document builder (XML-injection approach) ───

CONF_TEMPLATE_PATH = Path("assets/conf_template.docx")


def _xe(text: str) -> str:
    """Escape text for XML."""
    return text.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;").replace('"', "&quot;")


def _month_en(date_str: str) -> str:
    months = ["January","February","March","April","May","June",
              "July","August","September","October","November","December"]
    try:
        parts = date_str.replace(".", "/").split("/")
        m = int(parts[1])
        y = parts[2] if len(parts) > 2 else ""
        return f"{months[m-1]} {y}".strip()
    except Exception:
        return date_str


def _short_date(date_str: str) -> str:
    try:
        parts = date_str.split(".")
        return f"{parts[0]}.{parts[1]}."
    except Exception:
        return date_str


def _short_year(date_str: str) -> str:
    try:
        parts = date_str.split(".")
        return parts[2][-2:] if len(parts) > 2 else ""
    except Exception:
        return ""


# XML helper snippets — Inter font, colors matching reference exactly
_IR  = '<w:rFonts w:ascii="Inter" w:hAnsi="Inter"/>'          # Inter font
_IC  = '<w:color w:val="0B3A43"/>'                             # teal
_GC  = '<w:color w:val="9C8138"/>'                             # gold
_S20 = '<w:sz w:val="20"/><w:szCs w:val="20"/>'               # 10 pt
_B   = '<w:b/><w:bCs/>'                                        # bold

def _rpr_teal()   : return f'<w:rPr>{_IR}{_IC}{_S20}</w:rPr>'
def _rpr_gold_b() : return f'<w:rPr>{_IR}{_B}{_GC}</w:rPr>'
def _rpr_teal_b() : return f'<w:rPr>{_IR}{_B}{_IC}{_S20}</w:rPr>'

def _para_gold_bold(text: str) -> str:
    return (
        f'<w:p><w:pPr><w:rPr>{_IR}{_B}{_GC}</w:rPr></w:pPr>'
        f'<w:r><w:rPr>{_IR}{_B}{_GC}</w:rPr><w:t>{_xe(text)}</w:t></w:r></w:p>'
    )

def _para_teal(text: str, extra_ppr: str = "") -> str:
    return (
        f'<w:p><w:pPr>{extra_ppr}<w:rPr>{_IR}{_IC}{_S20}</w:rPr></w:pPr>'
        f'<w:r><w:rPr>{_IR}{_IC}{_S20}</w:rPr><w:t xml:space="preserve">{_xe(text)}</w:t></w:r></w:p>'
    )

def _blank() -> str:
    return f'<w:p><w:pPr><w:rPr>{_IR}{_IC}{_S20}</w:rPr></w:pPr></w:p>'

def _table_cell(width: int, paragraphs: list[str], center: bool = True) -> str:
    jc = '<w:jc w:val="center"/>' if center else ''
    paras = "".join(paragraphs)
    return (
        f'<w:tc><w:tcPr><w:tcW w:w="{width}" w:type="dxa"/>'
        f'<w:vAlign w:val="center"/></w:tcPr>{paras}</w:tc>'
    )

def _tbl_para(text: str, bold: bool = False) -> str:
    rpr = _rpr_teal_b() if bold else _rpr_teal()
    bld = f'<w:rPr>{_IR}{_B}{_IC}{_S20}</w:rPr>' if bold else f'<w:rPr>{_IR}{_IC}{_S20}</w:rPr>'
    return (
        f'<w:p><w:pPr><w:jc w:val="center"/>{rpr}</w:pPr>'
        f'<w:r>{bld}<w:t xml:space="preserve">{_xe(text)}</w:t></w:r></w:p>'
    )

def _bullet_item(text: str, num_id: str = "20") -> str:
    """Bullet list item matching reference numbering style."""
    rpr = f'<w:rFonts w:ascii="Inter" w:hAnsi="Inter" w:cs="Calibri"/>{_IC}{_S20}'
    return (
        f'<w:p><w:pPr><w:pStyle w:val="Listenabsatz"/>'
        f'<w:numPr><w:ilvl w:val="0"/><w:numId w:val="{num_id}"/></w:numPr>'
        f'<w:rPr>{rpr}</w:rPr></w:pPr>'
        f'<w:r><w:rPr>{rpr}</w:rPr><w:t xml:space="preserve">{_xe(text)}</w:t></w:r></w:p>'
    )


def _build_conf_body(rechnung: dict, dmcs, guide_name: str, guide_phone: str) -> str:
    """Build the <w:body> inner XML for the confirmation document."""
    destination = (rechnung.get("destination_en") or rechnung.get("destination_de") or "").upper()
    start       = rechnung.get("travel_start", "")
    end         = rechnung.get("travel_end", "")
    hotels      = rechnung.get("hotels", [])
    clients     = rechnung.get("client_names", [])

    parts = []

    # ── CONFIRMATION DESTINATION ──────────────────────────────────
    parts.append(_para_gold_bold(f"CONFIRMATION {destination}"))

    # ── Date range + clients ──────────────────────────────────────
    s_parts = start.split(".")
    e_parts = end.split(".")
    s_day   = s_parts[0] if s_parts else ""
    e_day   = e_parts[0] if e_parts else ""
    s_mon   = _month_en(start)
    e_mon   = _month_en(end)
    date_line = f"{s_day}. {s_mon} – {e_day}. {e_mon}"
    clients_line = " & ".join(clients) if clients else ""

    rpr_t = f'{_IR}{_IC}{_S20}'
    date_para = (
        f'<w:p><w:pPr><w:rPr>{rpr_t}</w:rPr></w:pPr>'
        f'<w:r><w:rPr>{rpr_t}</w:rPr><w:t xml:space="preserve">{_xe(date_line)}</w:t></w:r>'
        f'<w:r><w:rPr>{rpr_t}</w:rPr><w:br/></w:r>'
        f'<w:r><w:rPr>{rpr_t}</w:rPr><w:br/></w:r>'
        f'<w:r><w:rPr>{rpr_t}</w:rPr><w:t>{_xe(clients_line)}</w:t></w:r></w:p>'
    )
    parts.append(date_para)
    parts.append(_blank())

    # ── YOUR HOTELS: ──────────────────────────────────────────────
    parts.append(_para_gold_bold("YOUR HOTELS:"))

    # ── Hotel table ───────────────────────────────────────────────
    rows = []
    # Header row
    header_row = (
        '<w:tr><w:trPr><w:trHeight w:val="340"/></w:trPr>'
        + _table_cell(2263, [_tbl_para("DATE",  bold=True)])
        + _table_cell(3402, [_tbl_para("HOTEL", bold=True)])
        + _table_cell(3397, [_tbl_para("ROOM",  bold=True)])
        + '</w:tr>'
    )
    rows.append(header_row)

    for h in hotels:
        ci   = h.get("check_in", "")
        co   = h.get("check_out", "")
        nts  = int(h.get("nights", 1))
        name = h.get("hotel_name", "")
        city = h.get("city", "")
        room = h.get("room_type", "")
        # A boat/yacht charter has no meal-plan equivalent — the AI leaves
        # this genuinely blank rather than guessing "breakfast" for it, so
        # only fall back to the default when the field is truly missing.
        meal = h.get("meal_plan_en") if "meal_plan_en" in h else "breakfast"
        meal = meal or ""

        short_ci = _short_date(ci)
        short_co = _short_date(co)
        yr       = _short_year(co)
        n_word   = "night" if nts == 1 else "nights"

        # DATE cell — date range + line break + "N night(s)"
        rpr_t2 = f'{_IR}{_IC}{_S20}'
        date_cell_para = (
            f'<w:p><w:pPr><w:jc w:val="center"/><w:rPr>{rpr_t2}</w:rPr></w:pPr>'
            f'<w:r><w:rPr>{rpr_t2}</w:rPr><w:t>{_xe(short_ci)} – {_xe(short_co)}{_xe(yr)}</w:t></w:r>'
            f'<w:r><w:rPr>{rpr_t2}</w:rPr><w:br/></w:r>'
            f'<w:r><w:rPr>{rpr_t2}</w:rPr><w:t>{nts} {_xe(n_word)}</w:t></w:r></w:p>'
        )

        # HOTEL cell — booking # + hotel name + city (separate paragraphs)
        hotel_paras = [
            _tbl_para(f"#{name[:4].upper()}", bold=False),   # placeholder booking ref
            _tbl_para(name),
            _tbl_para(city),
        ]

        # ROOM cell — room type + meal + checkout/checkin notes
        checkout_note = (h.get("checkout_note") or "").strip()
        checkin_note  = (h.get("checkin_note")  or "").strip()
        room_paras = []
        if room:
            room_paras.append(_tbl_para(f"1 x {room}"))
        if meal:
            room_paras.append(_tbl_para(f"incl. {meal}"))
        if checkin_note:
            room_paras.append(_tbl_para(checkin_note))
        if checkout_note:
            room_paras.append(_tbl_para(checkout_note))

        data_row = (
            f'<w:tr><w:trPr><w:trHeight w:val="1077"/></w:trPr>'
            + _table_cell(2263, [date_cell_para])
            + _table_cell(3402, hotel_paras)
            + _table_cell(3397, room_paras)
            + '</w:tr>'
        )
        rows.append(data_row)

    table_xml = (
        '<w:tbl>'
        '<w:tblPr>'
        '<w:tblStyle w:val="Tabellenraster"/>'
        '<w:tblW w:w="0" w:type="auto"/>'
        '</w:tblPr>'
        '<w:tblGrid>'
        '<w:gridCol w:w="2263"/><w:gridCol w:w="3402"/><w:gridCol w:w="3397"/>'
        '</w:tblGrid>'
        + "".join(rows)
        + '</w:tbl>'
    )
    parts.append(table_xml)
    parts.append(_blank())

    # ── INCLUSIONS: ───────────────────────────────────────────────
    parts.append(_para_gold_bold("INCLUSIONS:"))
    parts.append(_blank())

    parsed_inclusions = rechnung.get("inclusions", [])

    if parsed_inclusions:
        # Normalise: accept both old format (plain strings) and new format (dicts)
        def _inc_text(item):
            return item.get("text", "") if isinstance(item, dict) else str(item)
        def _inc_hotel(item):
            return item.get("hotel", "") if isinstance(item, dict) else ""

        # 1. General inclusions first (hotel == "")
        for item in parsed_inclusions:
            if not _inc_hotel(item) and _inc_text(item).strip():
                parts.append(_bullet_item(_inc_text(item).strip(), "20"))

        # 2. Hotel-specific inclusions grouped by hotel
        seen_hotels: list = []
        for item in parsed_inclusions:
            hotel_label = _inc_hotel(item).strip()
            if hotel_label and hotel_label not in seen_hotels:
                seen_hotels.append(hotel_label)

        for hotel_label in seen_hotels:
            hotel_items = [
                item for item in parsed_inclusions
                if _inc_hotel(item).strip() == hotel_label and _inc_text(item).strip()
            ]
            if hotel_items:
                # Sub-header: hotel name in gold bold (smaller than main headers)
                parts.append(_blank())
                rpr_sub = f'{_IR}{_B}{_GC}'
                parts.append(
                    f'<w:p><w:pPr><w:rPr>{rpr_sub}</w:rPr></w:pPr>'
                    f'<w:r><w:rPr>{rpr_sub}</w:rPr>'
                    f'<w:t>{_xe(hotel_label)}:</w:t></w:r></w:p>'
                )
                for item in hotel_items:
                    parts.append(_bullet_item(_inc_text(item).strip(), "20"))
    else:
        # Fallback defaults when AI returned nothing
        parts.append(_bullet_item(
            "All transfers by private air-conditioned vehicle according to itinerary", "20"
        ))
        parts.append(_bullet_item(
            "Private Sightseeing with German Speaking accompanying Guides incl. entrance fees", "20"
        ))

    parts.append(f'<w:p><w:pPr><w:ind w:left="360"/><w:rPr>{_IR}{_IC}{_S20}</w:rPr></w:pPr></w:p>')

    # ── GERMAN SPEAKING GUIDE ─────────────────────────────────────
    if guide_name.strip():
        parts.append(_para_gold_bold("GERMAN SPEAKING GUIDE:"))
        guide_text = guide_name.strip()
        if guide_phone.strip():
            guide_text += f"\nContact No: {guide_phone.strip()}"
        # Build as bullet with line break
        rpr_g = f'<w:rFonts w:ascii="Inter" w:hAnsi="Inter" w:cs="Calibri"/>{_IC}{_S20}'
        lines = guide_text.split("\n")
        inner = ""
        for i, ln in enumerate(lines):
            if i == 0:
                inner += f'<w:r><w:rPr>{rpr_g}</w:rPr><w:t>{_xe(ln)}</w:t></w:r>'
            else:
                inner += f'<w:r><w:rPr>{rpr_g}</w:rPr><w:br/><w:t>{_xe(ln)}</w:t></w:r>'
        parts.append(
            f'<w:p><w:pPr><w:pStyle w:val="Listenabsatz"/>'
            f'<w:numPr><w:ilvl w:val="0"/><w:numId w:val="21"/></w:numPr>'
            f'<w:rPr>{rpr_g}</w:rPr></w:pPr>{inner}</w:p>'
        )
        parts.append(_blank())

    # ── CONTACT PERSON ON SITE ────────────────────────────────────
    # A multi-country trip (e.g. Vietnam + Singapore) has a different local
    # partner per country — dmcs can be a single dict (old callers, or a
    # one-DMC trip) or a list of dicts, one per selected DMC.
    if isinstance(dmcs, dict):
        dmcs = [dmcs] if dmcs else []
    dmcs = [d for d in (dmcs or []) if d]

    parts.append(_para_gold_bold("CONTACT PERSON ON SITE / LOCAL TRAVEL AGENCY:"))

    rpr_c  = f'{_IR}{_IC}{_S20}'
    rpr_cb = f'{_IR}{_B}{_IC}{_S20}'
    # Only label each block with its country when there's more than one —
    # a single DMC keeps the original, simpler look.
    show_country_labels = len(dmcs) > 1

    for i, dmc in enumerate(dmcs):
        dmc_name    = dmc.get("name", "")
        dmc_contact = dmc.get("contact_person", "")
        dmc_mobile  = dmc.get("mobile", "")
        dmc_note    = dmc.get("note", "")
        dmc_dest    = dmc.get("_destination", "")
        if not (dmc_name or dmc_contact or dmc_mobile):
            continue

        inner = ""
        if show_country_labels and dmc_dest:
            inner += f'<w:r><w:rPr>{rpr_cb}</w:rPr><w:t>{_xe(dmc_dest.upper())}</w:t></w:r>'
            inner += f'<w:r><w:rPr>{rpr_c}</w:rPr><w:br/></w:r>'
        inner += f'<w:r><w:rPr>{rpr_cb}</w:rPr><w:t>{_xe(dmc_name)}</w:t></w:r>'
        if dmc_contact:
            inner += f'<w:r><w:rPr>{rpr_c}</w:rPr><w:br/></w:r>'
            inner += f'<w:r><w:rPr>{rpr_c}</w:rPr><w:t xml:space="preserve">Ansprechpartner: {_xe(dmc_contact)}</w:t></w:r>'
        if dmc_mobile:
            note_txt = f" ({dmc_note})" if dmc_note else ""
            inner += f'<w:r><w:rPr>{rpr_c}</w:rPr><w:br/></w:r>'
            inner += f'<w:r><w:rPr>{rpr_c}</w:rPr><w:t xml:space="preserve">Mobile: {_xe(dmc_mobile)}{_xe(note_txt)}</w:t></w:r>'
        parts.append(f'<w:p><w:pPr><w:rPr>{rpr_c}</w:rPr></w:pPr>{inner}</w:p>')
        if i < len(dmcs) - 1:
            parts.append(_blank())

    # ── Closing text ──────────────────────────────────────────────
    parts.append(_blank())
    parts.append(_para_teal(
        "Bitte wenden Sie sich bei Fragen während der Reise immer zunächst an Ihren Ansprechpartner vor Ort."
    ))
    parts.append(_para_teal(
        "In dringenden Fällen außerhalb unserer Geschäftszeiten erreichen Sie uns "
        "per E-Mail unter urgent@bawa.de oder der Mobilnummer: 0049 (0) 160-898 7770"
    ))
    parts.append(f'<w:p><w:pPr><w:rPr>{_IR}{_IC}<w:sz w:val="2"/><w:szCs w:val="2"/></w:rPr></w:pPr></w:p>')
    parts.append(_blank())
    parts.append(_para_teal("Wir wünschen Ihnen eine interessante und traumhafte Reise!"))

    # Peter Huber signature in Quentin font (matching reference)
    parts.append(
        f'<w:p><w:pPr><w:spacing w:after="0"/>'
        f'<w:rPr><w:rFonts w:ascii="Quentin" w:hAnsi="Quentin"/><w:i/><w:iCs/>{_IC}<w:sz w:val="36"/><w:szCs w:val="36"/></w:rPr></w:pPr>'
        f'<w:r><w:rPr><w:rFonts w:ascii="Quentin" w:hAnsi="Quentin"/><w:i/><w:iCs/>{_IC}<w:sz w:val="36"/><w:szCs w:val="36"/></w:rPr>'
        f'<w:t>Peter Huber</w:t></w:r></w:p>'
    )
    parts.append(_para_teal("und das BAWA Team"))

    return "".join(parts)


def inject_into_conf_template(body_xml: str) -> bytes:
    """Inject body XML into conf_template.docx, preserving its header/footer/styles."""
    template_bytes = CONF_TEMPLATE_PATH.read_bytes()
    with zipfile.ZipFile(io.BytesIO(template_bytes)) as src:
        doc_xml = src.read("word/document.xml").decode("utf-8")

    body_start = doc_xml.index("<w:body>") + len("<w:body>")
    sect_pr_idx = doc_xml.rindex("<w:sectPr")
    new_doc = doc_xml[:body_start] + body_xml + doc_xml[sect_pr_idx:]

    buf = io.BytesIO()
    with zipfile.ZipFile(io.BytesIO(template_bytes)) as src:
        with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as dst:
            for item in src.infolist():
                if item.filename == "word/document.xml":
                    dst.writestr(item, new_doc.encode("utf-8"))
                else:
                    dst.writestr(item, src.read(item.filename))

    return buf.getvalue()


def build_confirmation_docx(rechnung: dict, dmcs, guide_name: str, guide_phone: str) -> bytes:
    """dmcs: a single DMC dict, a list of DMC dicts (multi-country trip), or
    falsy for none — see _build_conf_body."""
    body_xml = _build_conf_body(rechnung, dmcs or {}, guide_name or "", guide_phone or "")
    return inject_into_conf_template(body_xml)


async def _extract_rechnung_text(file: Optional[UploadFile], pasted_text: str) -> str:
    """Extract plain text from an uploaded Rechnung/Confirmation file or pasted text.
    No truncation — unlike extract_dmc_content, these documents are read in full.
    """
    if pasted_text.strip():
        return pasted_text.strip()
    if file and file.filename:
        fb = await file.read()
        ext = Path(file.filename).suffix.lower()
        if ext == ".pdf":
            text = _legacy_read_pdf(fb)
            if len(text.strip()) < 50:
                # No usable text layer — likely every page is a flattened
                # image (scanned Rechnung, or one exported to image rather
                # than real text) rather than a genuinely empty file.
                vision_text = _read_pdf_via_vision(fb)
                if vision_text.strip():
                    return vision_text
            return text
        elif ext == ".docx":
            return _legacy_read_word(fb, max_chars=None)
        elif ext in (".xlsx", ".xls"):
            return _legacy_read_excel(fb)
        elif ext == ".txt":
            return fb.decode("utf-8", errors="replace")
        else:
            raise HTTPException(400, f"Unsupported format: {ext}")
    raise HTTPException(400, "Bitte Datei hochladen oder Text einfügen.")
