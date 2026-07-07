
"""
Hotel bullet parser — written in a separate file to avoid quote-encoding issues
in the main app.py Edit tool workflow.
"""
import re
from datetime import datetime, timedelta


def normalize_quotes(text):
    """Replace all fancy quote variants with plain ASCII using chr() only."""
    for code in [0x201E, 0x201C, 0x201D, 0x00AB, 0x00BB]:
        text = text.replace(chr(code), chr(0x22))  # -> plain "
    for code in [0x2018, 0x2019, 0x201A, 0x2039, 0x203A]:
        text = text.replace(chr(code), chr(0x27))  # -> plain '
    return text


def join_continuation_lines(text):
    """
    Join hotel bullet lines that were split across paragraphs.

    Two cases handled:
      Case A: line ends with 'inklusive' (meal on next line)
        '6 Übernachtungen in Yogyakarta im "Amanjiwo" in ... inklusive'
        'Frühstück'
        → '6 Übernachtungen in Yogyakarta im "Amanjiwo" in ... inklusive Frühstück'

      Case B: hotel line has room but no 'inklusive', next line is a standalone meal word
        '4 Übernachtungen in Hongkong im Peninsula ... Grand Deluxe Room'
        'Frühstück'
        → '4 Übernachtungen in Hongkong ... Grand Deluxe Room inklusive Frühstück'
    """
    MEAL_WORDS = (
        "breakfast", "fr", "halbpension", "halfboard",
        "vollpension", "fullboard", "dinner", "lunch",
        "abendessen", "room only",
    )
    HOTEL_START = re.compile(r"^\d+\s*.bernachtung", re.IGNORECASE)

    lines = text.splitlines()
    out = []
    i = 0
    while i < len(lines):
        line = lines[i]
        stripped = line.strip()

        # Does this look like a hotel bullet?
        if HOTEL_START.search(stripped):
            # Peek at next non-empty line
            j = i + 1
            while j < len(lines) and not lines[j].strip():
                j += 1
            if j < len(lines):
                next_line = lines[j].strip().lower()
                # Case A: ends with 'inklusive' (with optional trailing whitespace)
                if re.search(r"inklusive\s*$", stripped, re.IGNORECASE):
                    meal_candidate = lines[j].strip()
                    out.append(stripped + " " + meal_candidate)
                    i = j + 1
                    continue
                # Case B: no inklusive at all, next line is a standalone meal word
                elif not re.search(r"inklusive", stripped, re.IGNORECASE):
                    if any(next_line.startswith(w) for w in MEAL_WORDS):
                        meal_candidate = lines[j].strip()
                        out.append(stripped + " inklusive " + meal_candidate)
                        i = j + 1
                        continue

        out.append(line)
        i += 1

    return "\n".join(out)


def nights_between(d1, d2):
    try:
        dt1 = datetime.strptime(d1, "%d.%m.%Y")
        dt2 = datetime.strptime(d2, "%d.%m.%Y")
        return (dt2 - dt1).days
    except Exception:
        return 1


def add_days(date_str, n):
    try:
        dt = datetime.strptime(date_str, "%d.%m.%Y")
        return (dt + timedelta(days=n)).strftime("%d.%m.%Y")
    except Exception:
        return date_str


def meal_en(meal_de):
    m = meal_de.lower()
    if "vollpension" in m or "full" in m:
        return "fullboard"
    if "halbpension" in m or "abendessen" in m or "half" in m:
        return "halfboard"
    if "breakfast" in m or "fr" in m:
        return "breakfast"
    return "room only"


def parse_leistungen_hotels(text, travel_start):
    """
    Parse hotel bullet lines of the form:
      X Uebernachtung(en) in [City] im "[Hotel]" in [Room] inklusive [Meal]

    Handles:
    - Proper Ü/ü (U+00DC / U+00FC)
    - Garbled '?' in place of Ü/ü (source encoding issue)
    - Garbled '?' in place of quote marks
    - Lines split across paragraphs (join_continuation_lines pre-processes these)
    """
    text = normalize_quotes(text)
    text = join_continuation_lines(text)
    hotels = []
    current_date = travel_start

    # Pattern A: proper Ü (U+00DC)
    pat_ue_upper = re.compile(
        r"^(\d+)\s*" + chr(0xDC) + r"bernachtung(?:en)?\s+in\s+(.+?)\s+im\s+"
        + chr(0x22) + r"(.+?)" + chr(0x22)
        + r"\s+in\s+(.+?)\s+inklusive\s+(.+)$",
        re.IGNORECASE
    )
    # Pattern B: lowercase ü (U+00FC)
    pat_ue_lower = re.compile(
        r"^(\d+)\s*" + chr(0xFC) + r"bernachtung(?:en)?\s+in\s+(.+?)\s+im\s+"
        + chr(0x22) + r"(.+?)" + chr(0x22)
        + r"\s+in\s+(.+?)\s+inklusive\s+(.+)$",
        re.IGNORECASE
    )
    # Pattern C: garbled '?' replacing Ü and quote chars (encoding-broken source docs)
    pat_garbled = re.compile(
        r"^(\d+)\s*\?bernachtung(?:en)?\s+in\s+(.+?)\s+im\s+"
        r"\?(.+?)\?"
        r"\s+in\s+(.+?)\s+inklusive\s+(.+)$",
        re.IGNORECASE
    )
    # Pattern D: no quotes around hotel name at all (e.g. Peninsula Hong Kong)
    pat_no_quotes = re.compile(
        r"^(\d+)\s*[" + chr(0xDC) + chr(0xFC) + r"?U]bernachtung(?:en)?\s+in\s+(.+?)\s+im\s+"
        r"(.+?)"
        r"\s+in\s+(?:einem?|einer|dem|den)?\s*(.+?)\s+inklusive\s+(.+)$",
        re.IGNORECASE
    )

    for line in text.splitlines():
        line = line.strip()
        # Collapse non-breaking spaces to regular spaces
        line = line.replace("\xa0", " ").replace("\t", " ")
        line = re.sub(r" +", " ", line).strip()

        m = (pat_ue_upper.match(line)
             or pat_ue_lower.match(line)
             or pat_garbled.match(line)
             or pat_no_quotes.match(line))
        if not m:
            continue

        nights_n = int(m.group(1))
        city     = m.group(2).strip().rstrip(",")
        hotel    = m.group(3).strip()
        room     = re.sub(r"\s+", " ", m.group(4).strip())
        meal_de  = re.sub(r"\s+", " ", m.group(5).strip())
        # Remove trailing "inklusive..." if regex grabs too much
        room = re.split(r"\s+inklusive", room, flags=re.IGNORECASE)[0].strip()

        check_in  = current_date
        check_out = add_days(current_date, nights_n)
        hotels.append({
            "check_in":     check_in,
            "check_out":    check_out,
            "nights":       nights_n,
            "hotel_name":   hotel,
            "city":         city,
            "room_type":    room,
            "meal_plan_de": meal_de,
            "meal_plan_en": meal_en(meal_de),
        })
        current_date = check_out

    return hotels
