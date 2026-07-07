# -*- coding: utf-8 -*-
"""
Loads and queries data/reference_library.json — the BAWA house-style
reference database built by build_reference_db.py from the Musterreiseverlaeufe
sample itineraries.

Used to find already-written, sent-to-client BAWA paragraphs for a given
hotel or sightseeing day, so call_ai_day() can hand them to the model as a
grounding reference instead of inventing facts from scratch every time.
"""
import json
import re
from functools import lru_cache
from pathlib import Path
from typing import Optional

from paths import DATA_DIR

DB_PATH = DATA_DIR / "reference_library.json"


@lru_cache(maxsize=1)
def _load() -> dict:
    if not DB_PATH.exists():
        return {"entries": []}
    return json.loads(DB_PATH.read_text(encoding="utf-8"))


def _save(data: dict) -> None:
    DB_PATH.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
    _load.cache_clear()


# ── Admin CRUD — used by the "Datenbank" tab so staff can add or correct a
# reference paragraph directly (e.g. after a wrong-hotel mismatch like the
# Aman/JANU Tokyo mixup) without waiting for the next full corpus rebuild. ──

def list_destinations() -> list:
    dests = {e["destination"] for e in _load()["entries"] if e.get("destination")}
    return sorted(dests)


def search_entries(destination: str = "", type_: str = "", query: str = "", limit: int = 200) -> list:
    """Entries matching the given filters, each tagged with its array index
    (`idx`) so the caller can edit/delete it. Re-fetch after any mutation —
    indices shift once an earlier entry is deleted."""
    entries = _load()["entries"]
    query_lower = query.strip().lower()
    out = []
    for i, e in enumerate(entries):
        if destination and e.get("destination", "").lower() != destination.lower():
            continue
        if type_ and e.get("type") != type_:
            continue
        if query_lower:
            haystack = " ".join([
                e.get("hotel_name", ""), e.get("location_heading", ""), e.get("text", ""),
            ]).lower()
            if query_lower not in haystack:
                continue
        out.append({**e, "idx": i})
        if len(out) >= limit:
            break
    return out


def add_entry(destination: str, type_: str, heading: str, text: str) -> dict:
    if type_ not in ("hotel", "sightseeing"):
        raise ValueError("type must be 'hotel' or 'sightseeing'")
    if not destination.strip() or not heading.strip() or not text.strip():
        raise ValueError("destination, heading and text are all required")

    data = _load()
    entry = {
        "type": type_,
        "destination": destination.strip(),
        "source_file": "Manuell hinzugefügt (Datenbank-Tab)",
        "language": "de",
        "text": text.strip(),
    }
    if type_ == "hotel":
        entry["hotel_name"] = heading.strip()
    else:
        entry["location_heading"] = heading.strip()

    data["entries"].append(entry)
    _save(data)
    return {**entry, "idx": len(data["entries"]) - 1}


def update_entry(idx: int, destination: str, type_: str, heading: str, text: str) -> dict:
    if type_ not in ("hotel", "sightseeing"):
        raise ValueError("type must be 'hotel' or 'sightseeing'")
    if not destination.strip() or not heading.strip() or not text.strip():
        raise ValueError("destination, heading and text are all required")

    data = _load()
    entries = data["entries"]
    if not (0 <= idx < len(entries)):
        raise IndexError("entry not found — the list may be stale, please refresh")

    entry = {
        "type": type_,
        "destination": destination.strip(),
        "source_file": entries[idx].get("source_file", "Manuell bearbeitet (Datenbank-Tab)"),
        "language": entries[idx].get("language", "de"),
        "text": text.strip(),
    }
    if type_ == "hotel":
        entry["hotel_name"] = heading.strip()
    else:
        entry["location_heading"] = heading.strip()

    entries[idx] = entry
    _save(data)
    return {**entry, "idx": idx}


def add_entries_bulk(items: list) -> int:
    """Used by the document-upload-and-parse review screen — saves every
    candidate the user kept checked, in a single file write instead of one
    write per entry."""
    data = _load()
    count = 0
    for item in items:
        type_ = item.get("type")
        destination = (item.get("destination") or "").strip()
        heading = (item.get("heading") or "").strip()
        text = (item.get("text") or "").strip()
        if type_ not in ("hotel", "sightseeing") or not destination or not heading or not text:
            continue
        entry = {
            "type": type_,
            "destination": destination,
            "source_file": item.get("source_file") or "Manuell hinzugefügt (Datenbank)",
            "language": item.get("language") or "de",
            "text": text,
        }
        if type_ == "hotel":
            entry["hotel_name"] = heading
        else:
            entry["location_heading"] = heading
        data["entries"].append(entry)
        count += 1
    if count:
        _save(data)
    return count


def delete_entry(idx: int) -> None:
    data = _load()
    entries = data["entries"]
    if not (0 <= idx < len(entries)):
        raise IndexError("entry not found — the list may be stale, please refresh")
    entries.pop(idx)
    _save(data)


def _tokens(name: str):
    return [t for t in re.findall(r"[A-Za-zÀ-ÖØ-öø-ÿ]+", name) if len(t) > 3]


def _contains_word(haystack_lower: str, token_lower: str) -> bool:
    """Whole-word match — plain substring containment would let a short
    token like "Inari" false-match inside an unrelated word such as
    "kulinarische"."""
    return re.search(rf"\b{re.escape(token_lower)}\b", haystack_lower) is not None


def distinguishing_tokens(text: str) -> set:
    """Significant (non-generic) lowercase tokens in `text`. Used both for
    matching and, by callers, to detect when two independently-matched
    references actually describe the same or overlapping sight (see
    app.py's per-bullet retrieval, which resolves each overview bullet to a
    reference separately and needs to avoid pulling in two source paragraphs
    that both happen to cover the same named attraction)."""
    return {t for t in (tok.lower() for tok in _tokens(text)) if not _is_generic(t)}


# Generic hotel-type words that appear across many unrelated properties —
# matching on these alone caused false positives, e.g. "The Tokyo Station
# Hotel" (a real booking) matching "Mandarin Oriental Tokyo" (an unrelated
# Tokyo hotel from a different sample itinerary) purely because both share
# the word "Tokyo", and "Wanosato Iya Onsen Ryokan" matching "Ryokan
# Kurashiki" purely because both are a "Ryokan".
_GENERIC_HOTEL_WORDS = frozenset({
    "hotel", "hotels", "ryokan", "resort", "resorts", "inn", "suite", "suites",
    "lodge", "house", "villa", "villas", "palace", "spa", "onsen", "the", "and",
})


def find_hotel_reference(destination: str, hotel_name: str) -> Optional[str]:
    """Best-matching hotel description text for this destination + hotel name.

    Requires the *distinguishing* (non-generic) words in the two names to
    substantially overlap — a single shared generic word like "Hotel" or
    "Ryokan", or a shared city name alone, is not enough confidence to reuse
    another property's description verbatim.
    """
    if not destination or not hotel_name:
        return None
    want = set(t.lower() for t in _tokens(hotel_name)) - _GENERIC_HOTEL_WORDS
    if not want:
        return None
    best, best_score = None, 0
    for e in _load()["entries"]:
        if e["type"] != "hotel" or e.get("language") != "de":
            continue
        if e["destination"].lower() != destination.lower():
            continue
        have = set(t.lower() for t in _tokens(e["hotel_name"])) - _GENERIC_HOTEL_WORDS
        score = len(want & have)
        # Require every distinguishing word to match when there's only one
        # (it's the sole identifying signal), otherwise at least two —
        # a single shared word among several is too weak to trust verbatim.
        threshold = 1 if len(want) == 1 else 2
        if score >= threshold and score > best_score:
            best, best_score = e, score
    return best["text"] if best else None


def find_sightseeing_references(destination: str, activities_text: str, limit: int = 3):
    """Reference paragraphs whose location heading or text overlaps with the
    place names mentioned in `activities_text` for this destination."""
    if not destination or not activities_text:
        return []
    candidates = [
        e for e in _load()["entries"]
        if e["type"] == "sightseeing"
        and e.get("language") == "de"
        and e["destination"].lower() == destination.lower()
    ]
    if not candidates:
        return []

    want = set(t.lower() for t in re.findall(r"[A-ZÀ-Ö][A-Za-zÀ-ÖØ-öø-ÿ]{3,}", activities_text))
    if not want:
        return []

    scored = []
    for e in candidates:
        haystack = (e.get("location_heading", "") + " " + e["text"]).lower()
        score = sum(1 for t in want if _contains_word(haystack, t))
        if score > 0:
            scored.append((score, e))
    scored.sort(key=lambda pair: -pair[0])

    out, seen = [], set()
    for _, e in scored:
        if e["text"] in seen:
            continue
        seen.add(e["text"])
        out.append(e["text"])
        if len(out) >= limit:
            break
    return out


# Generic word STEMS, matched by prefix — not exact words. German nouns
# decline by case and number (Tempel/Tempels/Tempeln, Schrein/Schreins/
# Schreine, Markt/Marktes, ...), so an exact-string exclusion list always has
# gaps: excluding "tempel" doesn't stop "Kiyomizu-Tempels" (genitive) from
# whole-word-matching "Kotokuin-Tempels" inside a completely unrelated
# reference paragraph. Matching by prefix catches every inflected form at
# once. None of these stems are ever the start of a real place/site name, so
# prefix matching doesn't risk excluding a genuine proper noun.
_GENERIC_STEMS = (
    "visit", "besuch", "golden", "pavillon", "pavilion", "schrein", "shrine",
    "tempel", "temple", "museum", "palast", "palace", "insel", "island",
    "strand", "beach", "garten", "garden", "berg", "mountain", "fluss", "river",
    "stadt", "city", "schloss", "castle", "brücke", "bridge", "markt", "market",
    "straße", "street", "park", "hotel", "ryokan", "flughafen", "airport",
    "zentrum", "center", "centre", "taisha", "jinja",
    # Tour-framing / logistics vocabulary — generic scene-setting bullets like
    # "Ganztägige flexible Tour mit englischsprachigem Guide" name no actual
    # sight, but without these excluded they were treated as having real
    # "distinguishing" content, and semantic search then matched them to
    # whatever unrelated city's "full-day guided tour" text scored closest.
    "tour", "guide", "reiseleiter", "fahrer", "chauffeur", "flexib", "ganztäg",
    "halbtäg", "englischsprachig", "deutschsprachig", "japanischsprachig",
    "sprech", "van", "fahrt", "ankunft", "abreise", "reise", "verfügung",
    "abend", "morgen", "nachmittag", "lobby", "treffen", "privat",
    # German grammar capitalizes ALL nouns, not just proper nouns like in
    # English — so common free-time/activity phrasing ("zur freien
    # Verfügung", "in eigenem Tempo erkunden") also survives as a false
    # "distinguishing" signal unless explicitly excluded here.
    "frei", "tag", "zeit", "eigen", "tempo", "erkund", "entdeck", "erleb",
    "genieß", "geniess", "verbring", "besichtig", "spaziergang", "spazier",
    "programm", "aktivität", "welcome", "arrival", "departure", "transfer",
    "today", "heute", "morning", "afternoon", "evening", "overnight",
    "übernachtung", "season", "walk", "through", "national", "royal", "grand",
    "great", "ancient", "world", "heritage",
    # Generic transport modes — a train or bus itself isn't a distinguishing
    # sight, and matching on it alone caused two different days' "high-speed
    # train to X" bullets to both resolve to the same stored transfer
    # paragraph regardless of route.
    "hochgeschwindigkeitszug", "schnellzug", "shinkansen", "bahnhof", "bahn",
    # Generic hotel-logistics vocabulary — "Ankunft und Check-in im Ryokan"
    # matched a stored paragraph about a totally different hotel purely
    # because both happen to mention "Check-in"/"Check-out".
    "check", "zimmer", "room", "suite",
    # Further generic geography/activity words in the same vein as the
    # transport-mode and hotel-logistics ones above.
    "expresszug", "viertel", "historisch", "schlucht", "durch", "bootsfahrt",
    "boots", "spaziergang",
    # Generic English viewpoint/geography words — DMC source text sometimes
    # keeps an attraction's English name untranslated (e.g. "Hinoji Gorge
    # Observation Deck"), and "Observation Deck"/"Gorge" alone matched an
    # unrelated stored paragraph about Roppongi Hills' observation deck.
    "gorge", "observation", "deck", "viewpoint", "lookout", "overlook",
    "valley", "canyon", "falls", "waterfall",
    # Generic cuisine/food words — a dish name alone isn't a distinguishing
    # sight, and matching on it caused a "Sushi-Kurs" (private sushi-making
    # class) bullet to resolve to an unrelated stored paragraph about Tsukiji
    # fish market, which only mentions sushi in passing as one of many foods
    # sold there.
    "sushi",
)

# Short/irregular words that don't decline predictably enough for prefix
# matching to help — excluded by exact match instead.
_GENERIC_EXACT = frozenset({
    "oder", "eine", "einem", "einen", "einer", "ihre", "ihrem", "ihren", "ihrer",
    "nach", "mit", "dem", "den", "der", "die", "das", "des",
})


def _is_generic(token: str) -> bool:
    return token in _GENERIC_EXACT or any(token.startswith(stem) for stem in _GENERIC_STEMS)


def find_exact_sightseeing_matches(destination: str, activities_text: str, location_heading: str = "", limit: int = 4, min_token_len: int = 5):
    """Sightseeing paragraphs to reuse VERBATIM for this day, in the order their
    subject first appears in `activities_text`.

    Unlike find_sightseeing_references (loose style guidance), this is meant
    for exact reuse: a match requires a distinctive capitalized word (>= 5
    letters, e.g. a place/site name — "Fushimi", "Kinkakuji") that actually
    appears inside the candidate paragraph's text. Generic word stems (see
    _GENERIC_STEMS) are excluded, as is the day's own destination/city name —
    otherwise a city like "Kyoto" matches almost every paragraph about that
    city and crowds out the paragraph actually describing the mentioned sight.
    Paragraphs matching the same set of words (near-duplicates of the same
    site across multiple source files) are collapsed to the single
    longest/most complete version. Paragraphs whose text is a verbatim
    duplicate of a stored *hotel* description (some source files misfile a
    hotel paragraph as sightseeing) are excluded too.
    """
    if not destination or not activities_text:
        return []
    entries = _load()["entries"]
    hotel_texts = {e["text"] for e in entries if e["type"] == "hotel"}
    candidates = [
        e for e in entries
        if e["type"] == "sightseeing"
        and e.get("language") == "de"
        and e["destination"].lower() == destination.lower()
        and e["text"] not in hotel_texts
    ]
    if not candidates:
        return []

    exclude = set(t.lower() for t in _tokens(destination)) | set(t.lower() for t in _tokens(location_heading))
    want_positions = {}
    for m in re.finditer(rf"[A-ZÀ-Ö][A-Za-zÀ-ÖØ-öø-ÿ]{{{min_token_len - 1},}}", activities_text):
        tok = m.group(0).lower()
        if _is_generic(tok) or tok in exclude:
            continue
        if tok not in want_positions:
            want_positions[tok] = m.start()
    if not want_positions:
        return []

    groups: dict = {}
    for e in candidates:
        haystack = e["text"].lower()
        matched = frozenset(t for t in want_positions if _contains_word(haystack, t))
        if not matched:
            continue
        groups.setdefault(matched, []).append(e)

    ranked = []
    for matched_toks, group_entries in groups.items():
        best = max(group_entries, key=lambda e: len(e["text"]))
        pos = min(want_positions[t] for t in matched_toks)
        ranked.append((pos, best["text"]))
    ranked.sort(key=lambda pair: pair[0])

    out, seen = [], set()
    for _, text in ranked:
        if text in seen:
            continue
        seen.add(text)
        out.append(text)
        if len(out) >= limit:
            break
    return out


def build_reference_block(destination: str, hotel_name: str = "", exact_matches: Optional[list] = None, max_chars: int = 2400) -> str:
    """Formats the hotel reference + numbered exact-match placeholders into a
    prompt-ready block, or "" if none found."""
    parts = []
    if hotel_name:
        hotel_ref = find_hotel_reference(destination, hotel_name)
        if hotel_ref:
            parts.append(f"[Hotel – {hotel_name}]\n{hotel_ref}")
    for i, text in enumerate(exact_matches or []):
        parts.append(f"[Platzhalter {{{{SIGHT:{i}}}}}]\n{text}")

    if not parts:
        return ""
    block = "\n\n".join(parts)
    if len(block) > max_chars:
        block = block[:max_chars] + "…"
    return block
