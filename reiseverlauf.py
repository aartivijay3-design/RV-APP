"""Reiseverlauf Generator: DMC-offer parsing, per-day AI prose generation
(grounded via reference_db.py + rag_retrieval.py), and Word document assembly."""
import json
import os
import re
import time
import zipfile
import io
from pathlib import Path
from xml.sax.saxutils import escape

import reference_db
from ai_client import _ai_complete, AI_MODEL

# Semantic (embedding-based) retrieval — complements reference_db's keyword
# matching for cases with no shared exact word (e.g. DMC spells a site
# "Kinkaku-ji" while the sample library has it as one word, "Kinkakuji").
# Optional: the app must still work if sentence-transformers isn't installed,
# and must be disableable outright on a memory-constrained host — loading
# the transformer model + PyTorch the first time it's used added enough
# memory to push a 512MB Render free-tier instance into an OOM kill. Set
# DISABLE_RAG=true there; exact-match + the AI's own knowledge still cover
# most cases without it.
_RAG_DISABLED_BY_ENV = os.environ.get("DISABLE_RAG", "").strip().lower() in ("1", "true", "yes")
try:
    if _RAG_DISABLED_BY_ENV:
        raise RuntimeError("RAG disabled via DISABLE_RAG env var")
    from rag_retrieval import retrieve as _rag_retrieve
    _RAG_ENABLED = True
except Exception:
    _RAG_ENABLED = False

TEMPLATE_PATH = Path("assets/template.docx")

# ── XML helpers ──────────────────────────────────────────────────────────────

def x(text: str) -> str:
    """XML-escape a string (handles &, <, >, keeps umlauts as UTF-8).

    None renders as an empty string, never the literal word "None". Any
    field the AI may legitimately return as JSON null can reach here (a
    trip with no stated party size, a stay with no named hotel), and
    str(None) put the word "None" in front of clients — an observed bug
    ("Reiseteilnehmer: None" on a real offer). A blank to fill in by hand
    is always the better failure.
    """
    return "" if text is None else escape(str(text))

RPR_TEAL = '<w:rFonts w:ascii="Inter" w:hAnsi="Inter"/><w:color w:val="0B3A43"/><w:sz w:val="22"/><w:szCs w:val="22"/>'
RPR_TEAL_B = '<w:rFonts w:ascii="Inter" w:hAnsi="Inter"/><w:b/><w:color w:val="0B3A43"/><w:sz w:val="22"/><w:szCs w:val="22"/>'
RPR_TEAL_B_U = '<w:rFonts w:ascii="Inter" w:hAnsi="Inter"/><w:b/><w:iCs/><w:color w:val="0B3A43"/><w:sz w:val="22"/><w:szCs w:val="22"/><w:u w:val="single"/>'
RPR_GOLD_B_U = '<w:rFonts w:ascii="Inter" w:hAnsi="Inter"/><w:b/><w:color w:val="9C8138"/><w:sz w:val="22"/><w:szCs w:val="22"/><w:u w:val="single"/>'
RPR_GOLD_B = '<w:rFonts w:ascii="Inter" w:hAnsi="Inter"/><w:b/><w:color w:val="9C8138"/><w:sz w:val="22"/><w:szCs w:val="22"/>'
RPR_TEAL_I = '<w:rFonts w:ascii="Inter" w:hAnsi="Inter"/><w:i/><w:color w:val="0B3A43"/><w:sz w:val="22"/><w:szCs w:val="22"/>'
RPR_LIST = '<w:rFonts w:ascii="Inter" w:hAnsi="Inter" w:cs="Calibri"/><w:color w:val="0B3A43"/><w:sz w:val="22"/>'

PPR_NUM = '<w:numPr><w:ilvl w:val="12"/><w:numId w:val="0"/></w:numPr><w:suppressAutoHyphens/><w:jc w:val="both"/>'
PPR_BODY = '<w:suppressAutoHyphens/><w:jc w:val="both"/>'


def para(ppr_inner: str, rpr: str, text: str) -> str:
    t = f'<w:t xml:space="preserve">{x(text)}</w:t>' if text else ""
    return (
        f'<w:p><w:pPr>{ppr_inner}<w:rPr>{rpr}</w:rPr></w:pPr>'
        + (f'<w:r><w:rPr>{rpr}</w:rPr>{t}</w:r>' if text else "")
        + "</w:p>"
    )


def ep() -> str:
    """Empty spacer paragraph."""
    return para(PPR_NUM, RPR_TEAL, "")


def day_heading(text: str) -> str:
    return para(PPR_NUM, RPR_TEAL_B_U, text)


def loc_heading(text: str) -> str:
    return para(PPR_NUM, RPR_TEAL_B, text)


def body_para(text: str) -> str:
    return para(PPR_BODY, RPR_TEAL, text)


def hotel_line(name: str) -> str:
    return para(PPR_BODY, RPR_GOLD_B, f"Übernachtung im {name}")


def section_heading(text: str) -> str:
    return (
        f'<w:p><w:pPr><w:spacing w:after="160" w:line="259" w:lineRule="auto"/>'
        f'<w:rPr>{RPR_GOLD_B_U}</w:rPr></w:pPr>'
        f'<w:r><w:rPr>{RPR_GOLD_B_U}</w:rPr><w:t>{x(text)}</w:t></w:r></w:p>'
    )


def page_break() -> str:
    return f'<w:p><w:r><w:rPr>{RPR_TEAL}</w:rPr><w:br w:type="page"/></w:r></w:p>'


def ende_reise() -> str:
    return para(PPR_BODY, RPR_GOLD_B, "ENDE DER REISE")


def truncation_warning() -> str:
    """Red warning paragraph inserted when AI output was truncated."""
    rpr = ('<w:rFonts w:ascii="Inter" w:hAnsi="Inter"/><w:b/>'
           '<w:color w:val="C00000"/><w:sz w:val="22"/><w:szCs w:val="22"/>')
    return (
        f'<w:p><w:pPr><w:jc w:val="both"/><w:rPr>{rpr}</w:rPr></w:pPr>'
        f'<w:r><w:rPr>{rpr}</w:rPr>'
        f'<w:t>⚠ HINWEIS: Die KI-Generierung wurde vorzeitig abgebrochen – '
        f'das Dokument ist möglicherweise unvollständig. '
        f'Bitte erneut generieren oder fehlende Tage manuell ergänzen.</w:t>'
        f'</w:r></w:p>'
    )


def missing_text_placeholder() -> str:
    """Placeholder for a day whose body_paragraphs were cut off."""
    rpr = ('<w:rFonts w:ascii="Inter" w:hAnsi="Inter"/><w:i/>'
           '<w:color w:val="C00000"/><w:sz w:val="22"/><w:szCs w:val="22"/>')
    return (
        f'<w:p><w:pPr><w:jc w:val="both"/><w:rPr>{rpr}</w:rPr></w:pPr>'
        f'<w:r><w:rPr>{rpr}</w:rPr>'
        f'<w:t>[Tagestext fehlt – bitte manuell ergänzen]</w:t>'
        f'</w:r></w:p>'
    )


def label_line(label: str, value: str, bold_label: bool = True, underline: bool = True) -> str:
    u = "<w:u w:val=\"single\"/>" if underline else ""
    rpr_lbl = f'<w:rFonts w:ascii="Inter" w:hAnsi="Inter"/><w:b/><w:color w:val="0B3A43"/><w:sz w:val="22"/><w:szCs w:val="22"/>{u}'
    rpr_val = RPR_TEAL
    return (
        f'<w:p><w:pPr><w:tabs><w:tab w:val="left" w:pos="5670"/></w:tabs>'
        f'<w:rPr>{rpr_lbl}</w:rPr></w:pPr>'
        f'<w:r><w:rPr>{rpr_lbl}</w:rPr><w:t>{x(label)}</w:t></w:r>'
        f'<w:r><w:rPr>{rpr_val}</w:rPr><w:tab/><w:tab/></w:r>'
        f'<w:r><w:rPr>{rpr_val}</w:rPr><w:t xml:space="preserve">{x(value)}</w:t></w:r>'
        f'</w:p>'
    )


def bullet_item(text: str) -> str:
    return (
        f'<w:p><w:pPr><w:pStyle w:val="Listenabsatz"/>'
        f'<w:numPr><w:ilvl w:val="0"/><w:numId w:val="16"/></w:numPr>'
        f'<w:suppressAutoHyphens/>'
        f'<w:rPr>{RPR_LIST}</w:rPr></w:pPr>'
        f'<w:r><w:rPr>{RPR_LIST}</w:rPr>'
        f'<w:t xml:space="preserve">{x(text)}</w:t></w:r></w:p>'
    )


def _ubersicht_photo_xml(rid: str, draw_id: int) -> str:
    """Inline drawing XML for a Reiseübersicht photo: 5.5cm × 4.85cm."""
    # 1 cm = 360000 EMU  →  5.5cm = 1980000, 4.85cm = 1746000
    cx, cy = 1980000, 1746000
    return (
        f'<w:drawing>'
        f'<wp:inline distT="0" distB="0" distL="0" distR="0"'
        f' xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing">'
        f'<wp:extent cx="{cx}" cy="{cy}"/>'
        f'<wp:effectExtent l="0" t="0" r="0" b="0"/>'
        f'<wp:docPr id="{draw_id}" name="Foto {draw_id}"/>'
        f'<wp:cNvGraphicFramePr>'
        f'<a:graphicFrameLocks xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" noChangeAspect="1"/>'
        f'</wp:cNvGraphicFramePr>'
        f'<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">'
        f'<a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">'
        f'<pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">'
        f'<pic:nvPicPr>'
        f'<pic:cNvPr id="{draw_id}" name="Foto {draw_id}"/>'
        f'<pic:cNvPicPr><a:picLocks noChangeAspect="1" noChangeArrowheads="1"/></pic:cNvPicPr>'
        f'</pic:nvPicPr>'
        f'<pic:blipFill>'
        f'<a:blip r:embed="{rid}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" cstate="print"/>'
        f'<a:srcRect/><a:stretch><a:fillRect/></a:stretch>'
        f'</pic:blipFill>'
        f'<pic:spPr bwMode="auto">'
        f'<a:xfrm><a:off x="0" y="0"/><a:ext cx="{cx}" cy="{cy}"/></a:xfrm>'
        f'<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>'
        f'<a:noFill/><a:ln><a:noFill/></a:ln>'
        f'</pic:spPr>'
        f'</pic:pic></a:graphicData></a:graphic>'
        f'</wp:inline></w:drawing>'
    )


def build_reiseubersicht_table(days: list, photo_rids: list = None) -> str:
    """Build the Reiseübersicht: 3-column table, brand colors, all borders, Inter 11pt.
    Consecutive free/transfer days (no bullets) are collapsed into a single merged row.
    """
    TEAL = "0B3A43"
    GOLD = "C4911A"
    SZ   = "22"  # 11pt

    def rpr(bold=False, color=TEAL):
        b = '<w:b/>' if bold else ''
        return (f'<w:rFonts w:ascii="Inter" w:hAnsi="Inter"/>{b}'
                f'<w:color w:val="{color}"/><w:sz w:val="{SZ}"/><w:szCs w:val="{SZ}"/>'
                f'<w:lang w:val="de-DE"/>')

    BS = 'w:val="single" w:sz="4" w:space="0" w:color="auto"'
    ALL_BORDERS = (
        f'<w:tcBorders>'
        f'<w:top {BS}/><w:left {BS}/><w:bottom {BS}/><w:right {BS}/>'
        f'</w:tcBorders>'
    )
    TBL_BORDERS = (
        f'<w:tblBorders>'
        f'<w:top {BS}/><w:left {BS}/><w:bottom {BS}/><w:right {BS}/>'
        f'<w:insideH {BS}/><w:insideV {BS}/>'
        f'</w:tblBorders>'
    )

    # Column widths (DXA): 1600 | 4800 | 2672 = 9072 total
    W1, W2, W3 = 1600, 4800, 2672
    MAR = '<w:tcMar><w:top w:w="80" w:type="dxa"/><w:left w:w="120" w:type="dxa"/><w:bottom w:w="80" w:type="dxa"/><w:right w:w="120" w:type="dxa"/></w:tcMar>'

    def tc(w, content_xml, span=None):
        """Single table cell; span sets gridSpan for column merging."""
        gs = f'<w:gridSpan w:val="{span}"/>' if span else ''
        return (f'<w:tc><w:tcPr><w:tcW w:w="{w}" w:type="dxa"/>{gs}'
                f'{ALL_BORDERS}{MAR}</w:tcPr>'
                f'{content_xml}</w:tc>')

    def is_free_day(day):
        """True when a day has no real activity bullets — just transfer/arrival."""
        bullets = day.get("overview_bullets", [])
        if not bullets:
            return True
        low = [(b or "").strip().lower() for b in bullets]
        return all(b in ("transfer", "anreise", "anreise / flug", "abreise", "flug") for b in low)

    parts = []

    # ── Header row ──────────────────────────────────────────────────────────────
    def hdr_cell(w, text):
        return (f'<w:tc><w:tcPr><w:tcW w:w="{w}" w:type="dxa"/>'
                f'{ALL_BORDERS}{MAR}</w:tcPr>'
                f'<w:p><w:pPr><w:spacing w:after="0"/></w:pPr>'
                f'<w:r><w:rPr>{rpr(bold=True, color=TEAL)}</w:rPr>'
                f'<w:t>{x(text)}</w:t></w:r></w:p></w:tc>')

    parts.append(
        '<w:tbl>'
        f'<w:tblPr><w:tblW w:w="9072" w:type="dxa"/>'
        f'{TBL_BORDERS}'
        f'<w:tblLook w:val="0000"/>'
        f'</w:tblPr>'
        f'<w:tblGrid>'
        f'<w:gridCol w:w="{W1}"/>'
        f'<w:gridCol w:w="{W2}"/>'
        f'<w:gridCol w:w="{W3}"/>'
        f'</w:tblGrid>'
        f'<w:tr>'
        + hdr_cell(W1, "Tag / Datum")
        + hdr_cell(W2, "Highlights")
        + hdr_cell(W3, "Übernachtung")
        + '</w:tr>'
    )

    # ── Group consecutive free days ──────────────────────────────────────────────
    # Each group is either a single normal day, or a run of 2+ free days merged.
    groups = []
    i = 0
    while i < len(days):
        if is_free_day(days[i]):
            run = [days[i]]
            while i + 1 < len(days) and is_free_day(days[i + 1]):
                i += 1
                run.append(days[i])
            groups.append(("free", run))
        else:
            groups.append(("normal", [days[i]]))
        i += 1

    # ── Render rows ──────────────────────────────────────────────────────────────
    for kind, group in groups:
        if kind == "normal":
            day = group[0]
            global_i = days.index(day)
            weekday    = day.get("weekday", "")
            date_str   = day.get("date", "")
            location   = day.get("location_heading", "")
            hotel      = day.get("hotel", {}) if isinstance(day.get("hotel"), dict) else {}
            hotel_name = hotel.get("name", "")
            bullets    = day.get("overview_bullets", [])
            day_num    = day.get("day_number", global_i + 1)
            day_num_end = day.get("day_number_end")

            if day_num_end:
                # A merged multi-day block (see _merge_undifferentiated_days)
                # — show the day/date range instead of a single weekday.
                day_label  = f"Tag {day_num}–{day_num_end}"
                # `or ""` — f-strings stringify None to "None" before x() runs.
                date_label = f"{date_str} – {day.get('date_end') or ''}".strip(" –")
                col1 = (
                    f'<w:p><w:pPr><w:spacing w:after="0"/></w:pPr>'
                    f'<w:r><w:rPr>{rpr(bold=True)}</w:rPr>'
                    f'<w:t xml:space="preserve">{x(day_label)}</w:t></w:r></w:p>'
                    f'<w:p><w:pPr><w:spacing w:after="0"/></w:pPr>'
                    f'<w:r><w:rPr>{rpr()}</w:rPr>'
                    f'<w:t>{x(date_label)}</w:t></w:r></w:p>'
                )
            else:
                col1 = (
                    f'<w:p><w:pPr><w:spacing w:after="0"/></w:pPr>'
                    f'<w:r><w:rPr>{rpr(bold=True)}</w:rPr>'
                    f'<w:t xml:space="preserve">Tag {x(str(day_num))}</w:t></w:r></w:p>'
                    f'<w:p><w:pPr><w:spacing w:after="0"/></w:pPr>'
                    f'<w:r><w:rPr>{rpr()}</w:rPr>'
                    f'<w:t>{x(weekday)}</w:t></w:r></w:p>'
                    f'<w:p><w:pPr><w:spacing w:after="0"/></w:pPr>'
                    f'<w:r><w:rPr>{rpr()}</w:rPr>'
                    f'<w:t>{x(date_str)}</w:t></w:r></w:p>'
                )

            col2 = (
                f'<w:p><w:pPr><w:spacing w:after="40"/></w:pPr>'
                f'<w:r><w:rPr>{rpr(bold=True)}</w:rPr>'
                f'<w:t>{x(location)}</w:t></w:r></w:p>'
            )
            for b in bullets:
                if not b:
                    continue
                col2 += (
                    f'<w:p><w:pPr><w:spacing w:after="20"/>'
                    f'<w:ind w:left="160" w:hanging="160"/></w:pPr>'
                    f'<w:r><w:rPr>{rpr(color=TEAL)}</w:rPr>'
                    f'<w:t xml:space="preserve">· {x(b)}</w:t></w:r></w:p>'
                )

            col3_text = hotel_name if hotel_name else "–"
            col3 = (
                f'<w:p><w:pPr><w:spacing w:after="0"/></w:pPr>'
                f'<w:r><w:rPr>{rpr()}</w:rPr>'
                f'<w:t>{x(col3_text)}</w:t></w:r></w:p>'
            )

            parts.append(f'<w:tr>{tc(W1, col1)}{tc(W2, col2)}{tc(W3, col3)}</w:tr>')

        else:
            # Collapsed free-day row: merge cols 2+3 into one wide cell
            first = group[0]
            last  = group[-1]
            first_num = first.get("day_number", days.index(first) + 1)
            last_num  = last.get("day_number",  days.index(last)  + 1)

            if len(group) == 1:
                day_label = f'Tag {first_num}'
                date_label = first.get("date") or ""
                weekday_label = first.get("weekday") or ""
            else:
                day_label = f'Tag {first_num} – {last_num}'
                # `or ""` — f-strings stringify None to "None" before x() runs.
                date_label = f'{first.get("date") or ""} – {last.get("date") or ""}'.strip(" –")
                weekday_label = ""

            # Determine label from bullets or location
            loc = first.get("location_heading") or ""
            bullets = first.get("overview_bullets") or []
            if bullets:
                activity_label = bullets[0]
            elif loc:
                activity_label = loc
            else:
                activity_label = "Transfer / Anreise"

            col1 = (
                f'<w:p><w:pPr><w:spacing w:after="0"/></w:pPr>'
                f'<w:r><w:rPr>{rpr(bold=True)}</w:rPr>'
                f'<w:t xml:space="preserve">{x(day_label)}</w:t></w:r></w:p>'
                + (f'<w:p><w:pPr><w:spacing w:after="0"/></w:pPr>'
                   f'<w:r><w:rPr>{rpr()}</w:rPr>'
                   f'<w:t>{x(weekday_label)}</w:t></w:r></w:p>' if weekday_label else '')
                + f'<w:p><w:pPr><w:spacing w:after="0"/></w:pPr>'
                  f'<w:r><w:rPr>{rpr()}</w:rPr>'
                  f'<w:t>{x(date_label)}</w:t></w:r></w:p>'
            )

            # Merged col2+col3 spanning 2 grid columns
            merged_content = (
                f'<w:p><w:pPr><w:spacing w:after="0"/></w:pPr>'
                f'<w:r><w:rPr>{rpr(bold=True)}</w:rPr>'
                f'<w:t>{x(activity_label)}</w:t></w:r></w:p>'
            )
            merged_cell = (
                f'<w:tc><w:tcPr><w:tcW w:w="{W2 + W3}" w:type="dxa"/>'
                f'<w:gridSpan w:val="2"/>'
                f'{ALL_BORDERS}{MAR}</w:tcPr>'
                f'{merged_content}</w:tc>'
            )

            parts.append(f'<w:tr>{tc(W1, col1)}{merged_cell}</w:tr>')

    parts.append('</w:tbl>')
    return "".join(parts)


def cover_page(client_name: str, date_line: str, destination: str = "DESTINATION", subtitle: str = "Eine Reise voller Eindrücke") -> str:
    """Build the cover page XML — keeps rId10 (Tokyo photo) from template."""
    inter_rpr    = '<w:rFonts w:ascii="Inter" w:hAnsi="Inter"/><w:color w:val="0B3A43"/><w:sz w:val="28"/><w:szCs w:val="28"/>'
    inter40_rpr  = '<w:rFonts w:ascii="Inter" w:hAnsi="Inter"/><w:b/><w:color w:val="0B3A43"/><w:sz w:val="44"/><w:szCs w:val="44"/>'
    italic_rpr   = '<w:rFonts w:ascii="Inter" w:hAnsi="Inter"/><w:i/><w:color w:val="0B3A43"/><w:sz w:val="28"/><w:szCs w:val="28"/>'
    bold_rpr     = '<w:rFonts w:ascii="Inter" w:hAnsi="Inter"/><w:b/><w:color w:val="0B3A43"/><w:sz w:val="28"/><w:szCs w:val="28"/>'

    # Anchored floating image — 20.32 cm × 13.56 cm, bleeds left (negative h-offset)
    # Template: rId11 = media/image1.jpeg; we reuse rId10 from our template
    photo_xml = (
        '<w:r><w:rPr><w:noProof/></w:rPr><w:drawing>'
        '<wp:anchor distT="0" distB="0" distL="0" distR="0" simplePos="0" '
        '  relativeHeight="251661312" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1">'
        '<wp:simplePos x="0" y="0"/>'
        '<wp:positionH relativeFrom="column"><wp:posOffset>-899795</wp:posOffset></wp:positionH>'
        '<wp:positionV relativeFrom="paragraph"><wp:posOffset>240665</wp:posOffset></wp:positionV>'
        '<wp:extent cx="7559675" cy="5039995"/>'
        '<wp:effectExtent l="0" t="0" r="0" b="0"/>'
        '<wp:wrapTopAndBottom/>'
        '<wp:docPr id="1" name="Cover Photo"/>'
        '<wp:cNvGraphicFramePr>'
        '<a:graphicFrameLocks xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" noChangeAspect="1"/>'
        '</wp:cNvGraphicFramePr>'
        '<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">'
        '<a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">'
        '<pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">'
        '<pic:nvPicPr><pic:cNvPr id="0" name="Cover Photo"/>'
        '<pic:cNvPicPr><a:picLocks noChangeAspect="1" noChangeArrowheads="1"/></pic:cNvPicPr>'
        '</pic:nvPicPr>'
        '<pic:blipFill><a:blip r:embed="rId10" cstate="print"/>'
        '<a:srcRect t="74" b="74"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>'
        '<pic:spPr bwMode="auto"><a:xfrm><a:off x="0" y="0"/>'
        '<a:ext cx="7559675" cy="5039995"/></a:xfrm>'
        '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>'
        '<a:noFill/><a:ln><a:noFill/></a:ln></pic:spPr>'
        '</pic:pic></a:graphicData></a:graphic></wp:anchor></w:drawing></w:r>'
    )

    return (
        # empty top spacer
        f'<w:p><w:pPr><w:rPr>{inter40_rpr}</w:rPr></w:pPr></w:p>'
        # DESTINATION — large centered
        f'<w:p><w:pPr><w:jc w:val="center"/><w:rPr>{inter40_rpr}</w:rPr></w:pPr>'
        f'<w:r><w:rPr>{inter40_rpr}</w:rPr><w:t>{x(destination.upper())}</w:t></w:r></w:p>'
        # empty spacer
        f'<w:p><w:pPr><w:jc w:val="center"/><w:rPr>{inter_rpr}</w:rPr></w:pPr></w:p>'
        # italic subtitle
        f'<w:p><w:pPr><w:jc w:val="center"/><w:rPr>{italic_rpr}</w:rPr></w:pPr>'
        f'<w:r><w:rPr>{italic_rpr}</w:rPr><w:t>{x(subtitle)}</w:t></w:r></w:p>'
        # photo anchor paragraph
        f'<w:p><w:pPr><w:rPr>{inter_rpr}</w:rPr></w:pPr>'
        f'{photo_xml}'
        f'</w:p>'
        # spacer
        f'<w:p><w:pPr><w:jc w:val="center"/><w:rPr>{inter_rpr}</w:rPr></w:pPr></w:p>'
        # "Persönliche Rundreise für"
        f'<w:p><w:pPr><w:jc w:val="center"/><w:rPr>{inter_rpr}</w:rPr></w:pPr>'
        f'<w:r><w:rPr>{inter_rpr}</w:rPr><w:t xml:space="preserve">Persönliche Rundreise für </w:t></w:r></w:p>'
        # client name — bold
        f'<w:p><w:pPr><w:jc w:val="center"/><w:rPr>{bold_rpr}</w:rPr></w:pPr>'
        f'<w:r><w:rPr>{bold_rpr}</w:rPr><w:t>{x(client_name)}</w:t></w:r></w:p>'
        # spacer
        f'<w:p><w:pPr><w:jc w:val="center"/><w:rPr>{inter_rpr}</w:rPr></w:pPr></w:p>'
        # date line + page break
        f'<w:p><w:pPr><w:jc w:val="center"/><w:rPr>{inter_rpr}</w:rPr></w:pPr>'
        f'<w:r><w:rPr>{inter_rpr}</w:rPr><w:t>{x(date_line)}</w:t></w:r>'
        f'<w:r><w:rPr>{inter_rpr}</w:rPr><w:br w:type="page"/></w:r></w:p>'
    )


def map_image_xml(rid: str = "rId11") -> str:
    """Inline map/Landkarte image — full content width, centered."""
    cx, cy = 5760720, 4320540  # 6.30 x 4.72 inches, original template dimensions
    return (
        f'<w:p><w:pPr><w:jc w:val="center"/><w:spacing w:before="0" w:after="0"/></w:pPr>'
        f'<w:r><w:rPr><w:noProof/></w:rPr><w:drawing>'
        f'<wp:inline distT="0" distB="0" distL="0" distR="0">'
        f'<wp:extent cx="{cx}" cy="{cy}"/>'
        f'<wp:effectExtent l="0" t="0" r="0" b="0"/>'
        f'<wp:docPr id="2" name="Landkarte"/>'
        f'<wp:cNvGraphicFramePr>'
        f'<a:graphicFrameLocks xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" noChangeAspect="1"/>'
        f'</wp:cNvGraphicFramePr>'
        f'<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">'
        f'<a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">'
        f'<pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">'
        f'<pic:nvPicPr><pic:cNvPr id="0" name="Landkarte"/>'
        f'<pic:cNvPicPr><a:picLocks noChangeAspect="1" noChangeArrowheads="1"/></pic:cNvPicPr>'
        f'</pic:nvPicPr>'
        f'<pic:blipFill><a:blip r:embed="{rid}" cstate="print"/>'
        f'<a:srcRect/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>'
        f'<pic:spPr bwMode="auto"><a:xfrm><a:off x="0" y="0"/>'
        f'<a:ext cx="{cx}" cy="{cy}"/></a:xfrm>'
        f'<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>'
        f'<a:noFill/><a:ln><a:noFill/></a:ln></pic:spPr>'
        f'</pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>'
    )


# ── Document builder ─────────────────────────────────────────────────────────

def build_body_xml(itinerary: dict, ubersicht_mode: str = "both") -> str:
    # Every route that produces a Word file funnels through here — the AI
    # path, and /build-docx handing back an itinerary the user edited in the
    # browser (which can carry nulls the AI never produced, e.g. a day whose
    # text was cleared, or a re-loaded saved draft). Normalizing the shape at
    # this single choke point means no renderer below has to re-guard against
    # a null list or a null item inside one. See _sanitize_structure.
    itinerary = _sanitize_structure(dict(itinerary))
    parts = []

    # Cover page
    # .get(key, default) only falls back when the key is MISSING — the AI can
    # legitimately return client_name: "" (no client named in the source
    # document), which must still show a generic placeholder, not a blank
    # cover page.
    client_display = itinerary.get("client_name") or "Familie"
    # f-string interpolation stringifies None to the literal "None" BEFORE
    # x() ever sees it, so these need `or ""` rather than a .get default —
    # a document whose dates weren't recognized showed "None – None" across
    # the cover page.
    _start_fmt = itinerary.get("start_date_formatted") or ""
    _end_fmt = itinerary.get("end_date_formatted") or ""
    date_line = f"{_start_fmt} – {_end_fmt}".strip(" –")
    destination = itinerary.get("destination") or "Destination"
    subtitle = itinerary.get("cover_subtitle") or "Eine Reise voller Eindrücke"
    parts.append(cover_page(client_display, date_line, destination, subtitle))

    # Landkarte + Reiseübersicht (cover page already ends with page break)
    if ubersicht_mode != "none":
        parts.append(map_image_xml())  # Landkarte
        parts.append(ep())
        parts.append(build_reiseubersicht_table(itinerary.get("days", [])))
        parts.append(ep())

    if ubersicht_mode == "only":
        return "".join(parts)

    # Section heading
    parts.append(page_break())
    parts.append(section_heading("IHR PERSÖNLICHER REISEVERLAUF"))
    parts.append(ep())

    # Truncation warning banner — shown once at the top if AI output was cut off
    if itinerary.get("_truncated"):
        parts.append(truncation_warning())
        parts.append(ep())

    # Days — merge consecutive free days into one combined section
    def _is_free(day):
        # A day with real generated/entered text is never a free day,
        # regardless of overview_bullets — manually added days (via "+ Tag
        # hinzufügen") never populate overview_bullets at all, which made
        # this treat every one of them as free and silently drop its
        # body_paragraphs (the free-day branch never renders them).
        if any(p.strip() for p in day.get("body_paragraphs", [])):
            return False
        bullets = day.get("overview_bullets", [])
        if not bullets:
            return True
        low = [(b or "").strip().lower() for b in bullets]
        return all(b in ("transfer", "anreise", "anreise / flug", "abreise", "flug") for b in low)

    all_days = itinerary.get("days", [])
    di = 0
    while di < len(all_days):
        day = all_days[di]
        if _is_free(day):
            # Collect the run of consecutive free days
            run = [day]
            while di + 1 < len(all_days) and _is_free(all_days[di + 1]):
                di += 1
                run.append(all_days[di])

            if len(run) == 1:
                # Single free day — render normally but without body text
                parts.append(day_heading(f"{run[0]['weekday']}, {run[0]['date']}"))
                parts.append(loc_heading(run[0].get("location_heading", "Zur freien Verfügung")))
                parts.append(ep())
            else:
                # Multiple consecutive free days — merge into one heading
                first, last = run[0], run[-1]
                dn_first = first.get("day_number", "")
                dn_last  = last.get("day_number", "")
                date_range = f"{first['date']} – {last['date']}"
                parts.append(day_heading(f"Tag {dn_first} – {dn_last} / {date_range}"))
                parts.append(loc_heading("Zur freien Verfügung"))
                parts.append(ep())

            parts.append(ep())
            di += 1
            continue

        # Normal day — or a merged multi-day block (see
        # _merge_undifferentiated_days): show the day/date range instead of
        # a single weekday, since the block spans several calendar days.
        if day.get("day_number_end"):
            heading = f"Tag {day['day_number']}–{day['day_number_end']} · {day['date']} – {day.get('date_end', '')}"
        else:
            heading = f"{day['weekday']}, {day['date']}"
        parts.append(day_heading(heading))
        parts.append(loc_heading(day.get("location_heading", "")))
        parts.append(ep())

        body_paras = [p for p in day.get("body_paragraphs", []) if p.strip()]
        if body_paras:
            for para_text in body_paras:
                parts.append(body_para(para_text))
                parts.append(ep())
        else:
            parts.append(missing_text_placeholder())
            parts.append(ep())

        hotel = day.get("hotel", {})
        if hotel:
            hotel_desc = hotel.get("description", "")
            if hotel.get("is_first_night") and hotel_desc:
                already_in_body = any(hotel_desc.strip()[:40] in p for p in day.get("body_paragraphs", []))
                if not already_in_body:
                    parts.append(body_para(hotel_desc))
                    parts.append(ep())
            # hotel.name is "" (see STRUCTURE_PROMPT's anti-hallucination
            # rule) when the DMC gives no real property name for this stay —
            # only a generic overnight marker like "Ü in Vilnius". Nothing
            # useful to print in that case, so skip the line rather than
            # show either a blank "Übernachtung im" or that marker verbatim.
            hotel_name = (hotel.get("name") or "").strip()
            if hotel_name:
                parts.append(hotel_line(hotel_name))
                parts.append(ep())

        parts.append(ep())
        di += 1

    # ENDE
    parts.append(ende_reise())
    parts.append(ep())

    # Page break → Leistungsübersicht
    parts.append(page_break())

    # Leistungsübersicht header
    parts.append(
        f'<w:p><w:pPr><w:spacing w:after="160"/>'
        f'<w:rPr>{RPR_GOLD_B_U}</w:rPr></w:pPr>'
        f'<w:r><w:rPr>{RPR_GOLD_B_U}</w:rPr>'
        f'<w:t>LEISTUNGSÜBERSICHT</w:t></w:r></w:p>'
    )

    l = itinerary.get("leistungen", {})
    parts.append(label_line("Veranstalter:", "BAWA Tours & Travel"))
    parts.append(
        f'<w:p><w:pPr><w:tabs><w:tab w:val="left" w:pos="5670"/></w:tabs>'
        f'<w:rPr>{RPR_TEAL}</w:rPr></w:pPr></w:p>'
    )
    # leistungen.reiseteilnehmer is legitimately absent/None when the source
    # names no client (see STRUCTURE_PROMPT's anti-hallucination rule) — a
    # missing-key-only .get(..., default) doesn't catch an explicit None.
    parts.append(label_line("Reiseteilnehmer:", l.get("reiseteilnehmer") or client_display))
    parts.append(
        f'<w:p><w:pPr><w:tabs><w:tab w:val="left" w:pos="5670"/></w:tabs>'
        f'<w:rPr>{RPR_TEAL}</w:rPr></w:pPr></w:p>'
    )
    parts.append(label_line("Reisedatum:", l.get("reisedatum", date_line)))
    parts.append(
        f'<w:p><w:pPr><w:rPr>{RPR_TEAL}</w:rPr></w:pPr></w:p>'
    )
    parts.append(label_line("Reisedauer:", l.get("reisedauer", "")))
    parts.append(
        f'<w:p><w:pPr><w:rPr>{RPR_TEAL}</w:rPr></w:pPr></w:p>'
    )
    # pax is legitimately None when the source document names no client/
    # party size (see STRUCTURE_PROMPT's anti-hallucination rule) — that's
    # not a missing key .get(..., "") would catch, so str() on it prints
    # the literal word "None" here instead of a blank to fill in by hand.
    pax_val = itinerary.get("pax")
    pax_text = str(pax_val) if pax_val else "____"
    parts.append(label_line(f"Reisepreis (basierend auf {pax_text} Personen):", "EUR ____________ gesamt"))
    parts.append(
        f'<w:p><w:pPr><w:rPr>{RPR_TEAL}</w:rPr></w:pPr></w:p>'
    )

    # "Inkludierte Leistungen:" heading
    parts.append(
        f'<w:p><w:pPr><w:rPr>{RPR_LIST}<w:u w:val="single"/></w:rPr></w:pPr>'
        f'<w:r><w:rPr>{RPR_LIST}<w:b/><w:u w:val="single"/></w:rPr>'
        f'<w:t>Inkludierte Leistungen:</w:t></w:r></w:p>'
    )
    parts.append(
        f'<w:p><w:pPr><w:pStyle w:val="Listenabsatz"/>'
        f'<w:suppressAutoHyphens/><w:jc w:val="both"/>'
        f'<w:rPr>{RPR_LIST}</w:rPr></w:pPr></w:p>'
    )

    # ── Build bullets in Python — never trust AI to format these ─────────────
    dest = itinerary.get("destination", "")

    # 1. Hotel nights first — one bullet per hotel
    # Fields below can be explicitly None (not just missing) when the DMC
    # source names no real property — e.g. a shorthand overnight marker like
    # "Ü in Vilnius" instead of an actual hotel name (see STRUCTURE_PROMPT's
    # hotel_nights rule) — so .get(key, default) alone doesn't catch it and
    # printed the literal word "None" here.
    for h in l.get("hotel_nights", []):
        nights = h.get("nights") or ""
        city   = h.get("city") or ""
        hotel  = (h.get("hotel") or "").strip()
        room   = h.get("room_type") or ""
        meal   = h.get("meal_plan") or "Frühstück"
        link   = h.get("link", "")
        n_word = "Übernachtung" if nights == 1 else "Übernachtungen"
        hotel_part = f' im „{hotel}“' if hotel else ""
        text   = (
            f"{nights} {n_word} in {city}{hotel_part}"
            + (f" in einem {room}" if room else "")
            + f" inklusive {meal}"
        )
        if link:
            # Inline hyperlink for hotel
            rpr_link = f'{RPR_LIST}<w:rStyle w:val="Hyperlink"/>'
            parts.append(
                f'<w:p><w:pPr><w:pStyle w:val="Listenabsatz"/>'
                f'<w:numPr><w:ilvl w:val="0"/><w:numId w:val="16"/></w:numPr>'
                f'<w:suppressAutoHyphens/>'
                f'<w:rPr>{RPR_LIST}</w:rPr></w:pPr>'
                f'<w:r><w:rPr>{RPR_LIST}</w:rPr><w:t xml:space="preserve">{x(text)} (</w:t></w:r>'
                f'<w:hyperlink r:id="rIdHotel" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
                f'<w:r><w:rPr>{rpr_link}</w:rPr><w:t>{x(link)}</w:t></w:r></w:hyperlink>'
                f'<w:r><w:rPr>{RPR_LIST}</w:rPr><w:t>)</w:t></w:r></w:p>'
            )
        else:
            parts.append(bullet_item(text))

    # 2. Airport transfers
    parts.append(bullet_item(
        "Private Flughafentransfers (Ankunft & Abreise) mit Chauffeur"
    ))

    # 3. Sightseeing programme
    parts.append(bullet_item(
        f"Privates Sightseeing- und Ausflugsprogramm {dest} gemäß Reiseverlauf "
        f"mit eigenem Fahrzeug, Fahrer und englischsprachigem Guide; "
        f"inklusive Eintrittsgebühren für die genannten Sehenswürdigkeiten"
    ))

    # 4. Unique experiences
    for exp in l.get("special_experiences", []):
        if exp and exp.strip():
            clean = exp.strip().strip("'\"")
            if clean:
                parts.append(bullet_item(clean))

    # 5. Train tickets (if applicable)
    if l.get("has_train_tickets", False):
        parts.append(bullet_item(
            "Zugtickets für alle Shinkansen- und Zugverbindungen in der 1. Klasse"
        ))

    # 6. Local contact — always last
    parts.append(bullet_item(
        f"Örtliche Ansprechpartner / Agentur in {dest}"
    ))

    # Footer
    parts.append(f'<w:p><w:pPr><w:rPr>{RPR_TEAL}</w:rPr></w:pPr></w:p>')
    parts.append(
        f'<w:p><w:pPr><w:suppressAutoHyphens/><w:jc w:val="both"/>'
        f'<w:rPr>{RPR_TEAL_I}</w:rPr></w:pPr>'
        f'<w:r><w:rPr>{RPR_TEAL_I}</w:rPr>'
        f'<w:t>*Zwischenverkauf und Preisänderungen vorbehalten</w:t></w:r></w:p>'
    )
    parts.append(f'<w:p><w:pPr><w:rPr>{RPR_TEAL}</w:rPr></w:pPr></w:p>')

    return "".join(parts)


def inject_into_template(body_xml: str, photos: list = None, destination: str = "") -> bytes:
    """Replace document.xml body in template DOCX, return new DOCX bytes.

    photos: list of dicts with keys 'rid' (str), 'filename' (str), 'data' (bytes), 'mime' (str)
    destination: used to update the inner-page header country name
    """
    template_bytes = TEMPLATE_PATH.read_bytes()
    with zipfile.ZipFile(io.BytesIO(template_bytes)) as src:
        doc_xml = src.read("word/document.xml").decode("utf-8")
        orig_rels = src.read("word/_rels/document.xml.rels").decode("utf-8")

    body_start = doc_xml.index("<w:body>") + len("<w:body>")
    sect_pr_idx = doc_xml.rindex("<w:sectPr")
    new_doc = doc_xml[:body_start] + body_xml + doc_xml[sect_pr_idx:]

    # Build updated relationships if photos are provided
    if photos:
        rel_entries = []
        for p in photos:
            ext = Path(p["filename"]).suffix.lstrip(".").lower()
            rel_entries.append(
                f'<Relationship Id="{p["rid"]}" '
                f'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" '
                f'Target="media/{p["filename"]}"/>'
            )
        new_rels = orig_rels.replace(
            "</Relationships>",
            "\n".join(rel_entries) + "\n</Relationships>"
        )
        # Determine mime → content-type extension for [Content_Types].xml
        content_types_additions = set()
        for p in photos:
            ext = Path(p["filename"]).suffix.lstrip(".").lower()
            mime_map = {"jpg": "image/jpeg", "jpeg": "image/jpeg", "png": "image/png",
                        "gif": "image/gif", "bmp": "image/bmp", "webp": "image/webp"}
            content_types_additions.add((ext, mime_map.get(ext, "image/jpeg")))
    else:
        new_rels = orig_rels
        content_types_additions = set()

    buf = io.BytesIO()
    with zipfile.ZipFile(io.BytesIO(template_bytes)) as src:
        # Find header files that contain the destination name
        header_files = [n for n in src.namelist() if n.startswith("word/header") and n.endswith(".xml")]
        with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as dst:
            for item in src.infolist():
                if item.filename == "word/document.xml":
                    dst.writestr(item, new_doc.encode("utf-8"))
                elif item.filename in header_files:
                    hdr_xml = src.read(item.filename).decode("utf-8")
                    import re as _re
                    # Remove wave background image (anchor with behindDoc="1") from all headers
                    hdr_xml = _re.sub(
                        r'<w:r><w:rPr><w:noProof/></w:rPr><w:drawing>'
                        r'<wp:anchor[^>]*behindDoc="1"[^>]*>.*?</wp:anchor>'
                        r'</w:drawing></w:r>',
                        '', hdr_xml, flags=_re.DOTALL
                    )
                    # Replace destination name in header1 (default header has country text)
                    if destination and item.filename.endswith("header1.xml"):
                        # x() escapes &/</> — a destination like "Indien &
                        # Bhutan" inserted raw produces invalid XML
                        # (<w:t>INDIEN & BHUTAN</w:t>) that corrupts this
                        # whole part, making the generated .docx fail to
                        # open at all. Every other injection site in this
                        # file already goes through x(); this one was
                        # missed since it's a regex substitution, not an
                        # f-string built alongside the others.
                        dest_escaped = x(destination.upper())
                        hdr_xml = _re.sub(
                            r'(<w:t[^>]*>)[A-ZÄÖÜ][A-ZÄÖÜ\s&]{2,30}(</w:t>)',
                            lambda m: m.group(1) + dest_escaped + m.group(2),
                            hdr_xml
                        )
                    dst.writestr(item, hdr_xml.encode("utf-8"))
                elif item.filename == "word/_rels/document.xml.rels" and photos:
                    dst.writestr(item, new_rels.encode("utf-8"))
                elif item.filename == "[Content_Types].xml" and content_types_additions:
                    ct_xml = src.read(item.filename).decode("utf-8")
                    for ext, mime in content_types_additions:
                        tag = f'Extension="{ext}"'
                        if tag not in ct_xml:
                            ct_xml = ct_xml.replace(
                                "</Types>",
                                f'<Default Extension="{ext}" ContentType="{mime}"/>\n</Types>'
                            )
                    dst.writestr(item, ct_xml.encode("utf-8"))
                else:
                    dst.writestr(item, src.read(item.filename))

            # Write photo files into word/media/
            if photos:
                for p in photos:
                    dst.writestr(f"word/media/{p['filename']}", p["data"])

    return buf.getvalue()



SYSTEM_PROMPT = """You are the senior travel writer for BAWA Tours & Travel, a German ultra-luxury travel agency. Your output will be printed and handed directly to clients. Every word must be impeccable.

Return ONLY a single valid JSON object. No markdown fences. No explanation. Nothing outside the JSON.

---
LANGUAGE STANDARD — MANDATORY
---
The DMC offer defines the schedule, activities, hotels, and sequence. It is NOT a template to translate word-for-word. Treat it as raw input — extract the facts, then rewrite everything in beautiful, idiomatic German FROM SCRATCH.

Core principle: The finished Reiseverlauf must read as if a native German-speaking luxury travel writer composed it. It must never feel translated.

① No literal translation. Restructure sentences to sound natural in German. English sentence structure rarely maps cleanly — invert, combine, or split sentences as needed.
② Expand sparse DMC entries. A bare entry like "Visit Angkor Wat" becomes 3–4 sentences of atmospheric, factual description — what the place is, why it matters, what the client will experience.
③ Elevate the register. This is a luxury document for high-end clients. Use rich, evocative vocabulary — "Sie erleben", "Sie entdecken", "ein unvergesslicher Anblick", "inmitten", "beeindruckend", "atemberaubend" — but vary it; never repeat the same phrase.
④ German flow over literal accuracy. If a direct translation sounds clunky, rewrite it. "dedicated to Vishnu" becomes "dem Gott Vishnu geweiht" — rearranged for rhythm.
⑤ Vary sentence openings. Never start every paragraph with "Sie besuchen…". Mix: "Im Anschluss…", "Weiter geht es…", "Ein Höhepunkt des Tages…", "Erleben Sie…", "Tauchen Sie ein in…"
⑥ Connect activities narratively. Where the DMC lists bullet points, write flowing prose guiding the client through the day as a journey, not a checklist.
⑦ Correct German typography: ä ö ü ß always — NEVER ae/oe/ue/ss. German quotation marks: „so" not "so". Em dash where appropriate.
⑧ No Anglicisms unless genuinely used in German travel language (Check-in, Transfer = fine; "highlight" → "Höhepunkt").

---
CARDINAL RULE — WRITING QUALITY
---
Standard: GEO SAISON / Condé Nast Traveller Germany. Literate, atmospheric, deeply informed.

!! CRITICAL !! Do NOT paraphrase or translate from the DMC. The DMC is only a list of facts.
You must write entirely from YOUR OWN knowledge of each place — historical depth, cultural meaning,
architectural beauty, sensory atmosphere. The DMC tells you WHAT to visit. YOU explain WHY it matters
and what it feels like to be there.

BANNED phrases (automatic failure if used):
"Sie können", "Es gibt", "Man kann", "ist bekannt für", "erkunden Sie",
"traditionelle Küche", "besuchen Sie", "eine Vielzahl", "wunderschön"

---
SIGHTSEEING PARAGRAPHS — mandatory depth for every named site
---
Each named attraction requires ALL of the following woven into flowing prose:
① Founding year / historical era + the political or spiritual story behind its creation
② Why it matters — its role in Japanese history, religion, or culture
③ What makes it architecturally, artistically, or naturally exceptional
④ UNESCO World Heritage status if it applies
⑤ A vivid sensory moment: the quality of light at dawn, the scent of cedar and incense,
   the sound of gravel underfoot, the weight of 800 years of history in a single gate
⑥ An insight only a deeply informed guide would share — beyond any standard guidebook

EXAMPLE — bad (just paraphrasing DMC):
"Sie besuchen den Kinkaku-ji, der auch als Goldener Pavillon bekannt ist."

EXAMPLE — required quality:
"Der Kinkaku-ji entfaltet seine stärkste Wirkung in den Morgenstunden, wenn das erste Licht
durch die Kiefern bricht und das mit echtem Blattgold verkleidete Obergeschoss im stillen
Spiegelsee darunter flimmert wie ein Traum. Ursprünglich 1397 als Rückzugsvilla des Shōguns
Ashikaga Yoshimitsu erbaut — eines Mannes, der in der Ostasien-Politik seiner Zeit nahezu
kaiserliche Macht besaß — vereint das dreigeschossige Gebäude bewusst drei verschiedene
Architekturstile: Heian-Aristokratie, Samurai-Ästhetik und reines Zen im obersten Geschoss.
Was viele nicht wissen: Das heutige Gebäude ist ein Neubau von 1955, errichtet nach dem
Brandanschlag eines jungen Mönchs, dessen obsessive Liebe zu diesem Bauwerk ins Pathologische
umgeschlagen war — ein Fall, der Mishima zu seinem Roman 'Der Goldene Pavillon' inspirierte."

---
LEISURE / FREE DAYS
---
Never write "freie Zeit zur Verfügung". Always offer 3–4 specific, named recommendations:
- Exact street names, restaurant names (note Michelin stars), neighbourhood names
- State precisely why each is worth the detour — its history, its reputation, its uniqueness
- Include at least one insider recommendation not in mainstream travel guides

---
HOTEL DESCRIPTIONS — first night at each hotel only, minimum 5 sentences
---
Write from genuine knowledge of the hotel. Include:
① Its design concept and architectural philosophy — who designed it, what style
② Exact location and how that shapes the experience (neighbourhood, views, proximity)
③ One specific named signature: a restaurant name, a spa treatment name, a design feature,
   a cultural programme, or an award
④ Its position in the luxury hotel world — Relais & Châteaux, Leading Hotels, Forbes stars,
   Michelin recognition, or other credentials
⑤ Why these specific guests will love it — connect it to their journey

---
JSON SCHEMA
---
{
  "client_name": "Familie Grundler",
  "destination": "Japan",
  "start_date": "22.06.2026",
  "end_date": "02.07.2026",
  "start_date_formatted": "22. Juni",
  "end_date_formatted": "02. Juli 2026",
  "pax": 4,
  "days": [
    {
      "day_number": 1,
      "weekday": "Montag",
      "date": "22.06.2026",
      "location_heading": "Kyoto / Ankunft",
      "overview_bullets": [
        "Ankunft am Kansai International Airport",
        "Privatführung durch den Arashiyama-Bambushain",
        "Besuch des Kinkaku-ji (Goldener Pavillon)"
      ],
      "body_paragraphs": [
        "Minimum 5-sentence paragraph written from your own deep knowledge...",
        "Second paragraph — another attraction or aspect of the day...",
        "Third paragraph if the day has 3+ highlights (omit if only 1-2 activities)"
      ],
      "hotel": {
        "name": "Six Senses Kyoto",
        "is_first_night": true,
        "description": "5-sentence hotel portrait. EMPTY STRING on non-first nights at same hotel.",
        "meal_plan": "Frühstück"
      }
    }
  ],
  "leistungen": {
    "reiseteilnehmer": "Familie Grundler (4 Personen)",
    "reisedatum": "22. Juni bis 02. Juli 2026",
    "reisedauer": "11 Tage / 10 Übernachtungen",
    "special_experiences": [
      "Each unique paid experience as its own short German phrase — e.g. 'Private Rikscha-Fahrt durch den Bambushain Arashiyama'",
      "'Private Teezeremonie mit einem Kyotoer Teemeister'",
      "'Privatbesuch im Atelier des Messerschmieds in Sakai'"
    ],
    "hotel_nights": [
      {
        "nights": 5,
        "city": "Kyoto",
        "hotel": "Six Senses Kyoto",
        "room_type": "Deluxe Junior Suite",
        "meal_plan": "Frühstück"
      },
      {
        "nights": 1,
        "city": "Yamashiro-Onsen",
        "hotel": "Beniya Mukayu",
        "room_type": "Wakamurasaki Suite (120 sqm)",
        "meal_plan": "Halbpension"
      }
    ],
    "has_train_tickets": true
  },
  "duration_label": "11 Tage / 10 Übernachtungen"
}

---
STRUCTURAL RULES
---
1. EVERY day from the DMC must be in the JSON — not one day omitted.
2. overview_bullets: 2–5 short German noun phrases (NOT full sentences) listing the key activities or sights for that day — used in the Reiseübersicht table. Pure transfer or arrival days: just ["Transfer"] or ["Anreise / Flug"]. No full stops. No quotes.
3. body_paragraphs: min. 3 paragraphs for touring days; min. 2 for arrival/departure.
   Each paragraph min. 5 sentences of real, substantive content.
3. hotel.is_first_night = true ONLY on the first night at each hotel.
   All other nights at same hotel: is_first_night = false, description = "".
4. No meal mentions in body paragraphs.
5. hotel_nights: one entry PER hotel, with the exact room type from the DMC and correct night count.
6. Meal translations: BB/Breakfast→Frühstück | HB/Half Board→Halbpension |
   FB/Full Board→Vollpension | Ryokan dinner→Halbpension | No meals→ohne Verpflegung.
7. Weekdays: Montag Dienstag Mittwoch Donnerstag Freitag Samstag Sonntag.
8. Months: Januar Februar März April Mai Juni Juli August September Oktober November Dezember.
9. Dates in headings: DD.MM.YYYY. In leistungen reisedatum: "22. Juni bis 02. Juli 2026".
"""


STRUCTURE_PROMPT = """Extract the structure from this DMC travel offer. Return ONLY valid JSON, no markdown, no explanation.

{
  "client_name": "Familie Grundler",
  "destination": "Japan",
  "start_date": "22.06.2026",
  "end_date": "02.07.2026",
  "start_date_formatted": "22. Juni",
  "end_date_formatted": "02. Juli 2026",
  "pax": 4,
  "duration_label": "11 Tage / 10 Übernachtungen",
  "days": [
    {
      "day_number": 1,
      "weekday": "Montag",
      "date": "22.06.2026",
      "location_heading": "Kyoto / Ankunft",
      "is_free_day": false,
      "overview_bullets": ["Ankunft Kansai International Airport", "Transfer zum Hotel"],
      "hotel": {
        "name": "Six Senses Kyoto",
        "room_type": "Deluxe Junior Suite",
        "meal_plan": "Frühstück",
        "is_first_night": true
      },
      "day_marker": "Arrival Kyoto. Transfer to hotel."
    }
  ],
  "leistungen": {
    "reiseteilnehmer": "Familie Grundler (4 Personen)",
    "reisedatum": "22. Juni bis 02. Juli 2026",
    "reisedauer": "11 Tage / 10 Übernachtungen",
    "special_experiences": ["Private Rikscha-Fahrt durch den Bambushain Arashiyama"],
    "hotel_nights": [
      {"nights": 5, "city": "Kyoto", "hotel": "Six Senses Kyoto", "room_type": "Deluxe Junior Suite", "meal_plan": "Frühstück"}
    ],
    "has_train_tickets": false
  }
}

Rules:
- THE days ARRAY MUST CONTAIN EXACTLY ONE ENTRY PER CALENDAR DAY OF THE TRIP — from start_date through end_date, no gaps, no merging. This is the single most important rule below. DMC offers often describe a multi-night hotel stay as ONE date-range block instead of listing each night separately — for example "13 Jan – 18 Jan (6 Nights) Phu Quoc Island, Regent Phu Quoc" describes 6 calendar days at ONE hotel with no day-by-day breakdown given, and the same applies to a range written as "Day 13 – Day 19 | Enjoy Golf and pamper yourself with spa treatments" — that has a real (if brief) activity sentence, so it is real narrative content. Expand either kind into one entry per calendar date (day_number N through N+5, same hotel repeated, is_first_night true only on the first), not one entry for the whole block. The same applies to a "FREE TIME AT LEISURE" block spanning several nights, or any other stretch with no explicit day-by-day activity list — every one of those nights still gets its own entry, with is_free_day: true and overview_bullets: ["Freizeit"] on the ones with nothing specific listed. Before finishing, count: does the number of entries in `days` equal the number of nights in the trip (or nights + 1 if the departure day also gets its own entry)? If not, you have merged days that must be split apart.
- EXCEPTION to the rule above — do not extract a day entry from a table row that is PURELY logistics: just a day number/date, a destination, and a hotel/room/meal-plan, with NO activity sentence anywhere in the row (e.g. "Day 01 | Colombo | Shangri La Colombo | Horizon Club | Bed & Breakfast Basis" and nothing else). Many DMC documents restate their days a second time in a trailing "Accommodation Summary"/"Hotel Summary" table, after the real day-by-day narrative earlier in the document — extracting from both produced a 20-day trip that came out as 39 days, because that table's dateless rows each got a fabricated placeholder date and never matched up with the real days. This exception applies ONLY to that exact bare shape — never to a row with any activity sentence, however brief. When genuinely unsure which of the two this is, prefer expanding it per the rule above (a duplicate merges away harmlessly later; a dropped real day does not). Never invent a date for a day when none can be determined from context.
- location_heading: the place name ONLY — "Hakone", "Kyoto", "Tokyo / Ankunft", "Kyoto - Kinosaki Onsen". NEVER describe the activity type here ("Ganztägige Tour in Hakone", "Halbtägige Tour in Kyoto") — that belongs in overview_bullets, not the heading. Never append the touring/excursion destination if it differs from the town the hotel is actually in (e.g. touring "Sigiriya" while the hotel is in "Habarana": heading is "Habarana", not "Sigiriya" and not "Sigiriya / Habarana" — the excursion name belongs in overview_bullets). Use "CityA - CityB" (hyphen, starting city to ending city) only on a day that changes which city the hotel is in — never a slash for this. Use "/ Ankunft" only on the actual arrival day, "/ Abreise" only on the actual departure day, appended to whichever city name is correct for that day — if the day both transfers AND departs (e.g. a transfer from the last hotel's city straight to the airport for the international flight home), it's "CityA - CityB / Abreise", not just the airport name or an unrelated city.
- overview_bullets: 2-5 short German noun phrases per day. Transfer/arrival only days: ["Transfer"] or ["Anreise / Flug"].
- day_marker: a SHORT (8-15 words) EXACT, character-for-character excerpt copied verbatim from the very START of this day's section in the source document — the day's header line if there is one (e.g. "DAY SEVEN - MONDAY, 11 JAN 2027"), or the first distinctive sentence describing that day if there's no explicit header. This is used afterward to programmatically locate the day in the original text and slice out its real content, so precision matters more here than for any other field — copy it EXACTLY as it appears (capitalization, punctuation, spacing), never paraphrase, translate, reformat, or summarize it. Pick a phrase that is UNIQUE within the whole document — never a generic word/phrase that also occurs elsewhere ("Breakfast", "Overnight stay", a bare date that repeats). For a multi-night block with no day-by-day breakdown (see the days-array rule above), every expanded day within that same block shares the IDENTICAL day_marker pointing to where the block begins — do not invent different markers for days that have no distinct text of their own.
- hotel.is_first_night = true only on first arrival at each hotel.
- hotel.name must be the ACTUAL named property from the DMC (e.g. "Six Senses Kyoto"), never a generic overnight marker some DMC documents use in place of a real name ("Ü in Vilnius", "Overnight Riga", "Hotel TBD") — that marker is a placeholder, not a hotel name: in that case return name: "". Likewise return room_type: null and meal_plan: null when the DMC states none for that stay — never invent a plausible-looking value.
- Weekdays in German. Dates: DD.MM.YYYY.
- Meal plan: BB→Frühstück, HB→Halbpension, FB→Vollpension, AI→All-inclusive.
- DO NOT write any body_paragraphs — structure and raw activities only.
- is_free_day: true if the DMC gives no specific activity for the day (free/leisure/own arrangements). When true, set overview_bullets: ["Freizeit"]. The prose generator will insert the standard free-day line — do NOT write activities.
- hotel_nights: one entry per hotel (not per day), with correct total nights count.
- hotel_nights.hotel must be the ACTUAL named property from the DMC (e.g. "Six Senses Kyoto"). Some DMC documents give no real property name for a stay — only a generic overnight marker like "Ü in Vilnius", "Overnight Riga", "Hotel TBD" — that marker is a placeholder, not a hotel name: in that case return hotel: "". Likewise return room_type: null and meal_plan: null whenever the DMC states no room type / meal plan for that stay — never invent a plausible-looking value; a fabricated room type or meal plan is worse than a blank one, same reasoning as the client_name/pax rule below.
- client_name: ALWAYS in German format "Familie [Nachname]" (e.g. Familie Schiff, Familie Grundler). Extract the family name and prefix with "Familie". Never use English ("The X family" or "X Family"). If group name, keep it as-is but in German.
- client_name, pax, and leistungen.reiseteilnehmer must come from an ACTUAL name/party-size stated somewhere in THIS document (a cover sheet, "prepared for", a pax count). Many DMC documents (generic activity templates, rate sheets meant for repeat use) name no client at all — in that case return client_name: "", pax: null, and OMIT leistungen.reiseteilnehmer entirely. Never fall back to the example above ("Familie Grundler") or invent any other name/count — the real client name is supplied separately by the person generating this document, and a wrong name on the cover page is worse than a blank one.
"""


# ── Chunked extraction — for documents too long/dense to reliably process
# in one completion (see call_ai_structure's docstring for why). Splits the
# source into overlapping windows small enough that the model handles each
# one reliably (empirically solid up to ~5-6 days), extracts days from each
# independently, then merges by date in Python — sidestepping cross-chunk
# day-numbering consistency entirely, since day_number gets recomputed from
# sorted dates after merging rather than trusted from any one chunk.
STRUCTURE_METADATA_PROMPT = """Extract only the TRIP-LEVEL summary from this DMC travel offer — do NOT extract day-by-day details, only what's below.

Return ONLY valid JSON, no markdown, no explanation:
{
  "client_name": "Familie Grundler",
  "destination": "Japan",
  "start_date": "22.06.2026",
  "expected_days": 11,
  "pax": 4,
  "leistungen": {
    "reiseteilnehmer": "Familie Grundler (4 Personen)",
    "special_experiences": ["Private Rikscha-Fahrt durch den Bambushain Arashiyama"],
    "hotel_nights": [
      {"nights": 5, "city": "Kyoto", "hotel": "Six Senses Kyoto", "room_type": "Deluxe Junior Suite", "meal_plan": "Frühstück"}
    ],
    "has_train_tickets": false
  }
}

Rules:
- client_name: ALWAYS "Familie [Nachname]" in German (e.g. "Familie Schiff"). Never English ("The X family").
- destination: the country/region in German, e.g. "Japan", "Vietnam und Singapur".
- hotel_nights: one entry per hotel (not per day), with the correct total night count and exact room type.
- hotel_nights.hotel must be the ACTUAL named property from the DMC. Some DMC documents give no real property name for a stay — only a generic overnight marker like "Ü in Vilnius", "Overnight Riga", "Hotel TBD" — that marker is a placeholder, not a hotel name: in that case return hotel: "". Likewise return room_type: null and meal_plan: null whenever the DMC states no room type / meal plan for that stay — never invent a plausible-looking value.
- special_experiences: only genuinely distinctive/bookable experiences explicitly named in the offer, not generic sightseeing.
- start_date: the calendar date of Day 1 / arrival, DD.MM.YYYY. Some documents label days only "Day 01", "Day 02"... with no date anywhere near the day-by-day narrative itself — the real date is often only findable elsewhere in the document (a validity/pricing section, "Travelling Date:", a booking confirmation line). Search the WHOLE document for it; this is the one piece of context the day-by-day extraction step (which only sees small excerpts) can't find on its own, so getting it from here matters even when it feels like it belongs to a "pricing" section, not the itinerary. Return "" only if truly no date appears anywhere in the document.
- expected_days: the trip's total length in calendar days, if the document states it anywhere as a number — "7 Days", "19 Nights / 20 Days" (→ 20), "8 Tage / 7 Nächte" (→ 8), a day-by-day list that visibly runs "Day 1" through "Day N" (→ N), etc. This is the ONE independent check against a day-by-day extraction step silently losing days partway through a long document (each excerpt only sees part of the document and has no way to know the true total) — a real, observed failure: a 7-day trip came back with only the first 4 days because the excerpt covering days 4-7 didn't produce anything and nothing caught it, since with no independent count "expected" just gets computed from whatever days a chunk actually returned, which cannot detect days it dropped. Return null only if the document truly never states a total length anywhere.
- client_name, pax, and leistungen.reiseteilnehmer must come from an ACTUAL name/party-size stated somewhere in THIS document. Many DMC documents (generic activity templates, rate sheets meant for repeat use) name no client at all — in that case return client_name: "", pax: null, and OMIT leistungen.reiseteilnehmer entirely. Never fall back to the example above ("Familie Grundler") or invent any other name/count — the real client name is supplied separately by the person generating this document, and a wrong name on the cover page is worse than a blank one.
"""

STRUCTURE_CHUNK_PROMPT = """You are given ONE EXCERPT from a longer multi-day DMC (destination management company) travel offer — not the whole document. This excerpt may begin or end mid-day; overlap with adjacent excerpts covering the same document is expected and fine.

Return ONLY valid JSON, no markdown, no explanation:
{"days": [
  {
    "day_number": 1,
    "weekday": "Montag",
    "date": "22.06.2026",
    "location_heading": "Kyoto / Ankunft",
    "is_free_day": false,
    "overview_bullets": ["Ankunft Kansai International Airport", "Transfer zum Hotel"],
    "hotel": {
      "name": "Six Senses Kyoto",
      "room_type": "Deluxe Junior Suite",
      "meal_plan": "Frühstück",
      "is_first_night": true
    },
    "day_marker": "Arrival Kyoto. Transfer to hotel."
  }
]}

Rules:
- EVERY calendar day in what you're given must become one entry — INCLUDING a multi-day range written as one block, e.g. "13 Jan – 18 Jan (6 Nights)" or "Day 13 – Day 19 | Enjoy Golf and pamper yourself with spa treatments": expand that into one entry per night (same hotel repeated, is_first_night true only on the first, is_free_day: true and overview_bullets: ["Freizeit"] only where truly nothing specific is given). This is the default — do this unless the exception below applies.
- EXCEPTION — do not extract a day entry from a table row that is PURELY logistics: just a day number/date, a destination, and a hotel/room/meal-plan, with NO sentence anywhere describing an activity, e.g. "Day 01 | Colombo | Shangri La Colombo | Horizon Club | Bed & Breakfast Basis" and nothing else. This exception exists ONLY for that exact shape (a bare accommodation-summary/hotel-summary table restating days already narrated earlier in the document) — extracting it too produced a 20-day trip that came out as 39 days. A row like "Day 13 – Day 19 | Enjoy Golf..." has a real activity sentence, so the exception does NOT apply to it — expand it per the rule above. When genuinely unsure which of the two this is, prefer expanding it (a duplicated day merges away harmlessly later; a dropped real day does not).
- Only include a day whose content is FULLY visible in this excerpt. If a day's description is visibly cut off at the very start or end of what you're given (trails off with no clear beginning/end), do NOT include it — an overlapping adjacent excerpt covers it completely elsewhere. It's fine and expected for a day to also appear in an adjacent excerpt; duplicates get merged and de-duplicated afterward by date.
- date: the actual calendar date for this day, DD.MM.YYYY — get this right even when unsure of day_number, since date (not day_number) is what's used to merge and order days across excerpts. If a "Trip start date" is given below and this excerpt only labels days by ordinal ("Day 01", "Day 07"), COMPUTE the date from it (start date + day_number − 1 days) — this is not "inventing", the anchor date makes it a real calculation. Only leave date empty if you can find no day ordinal AND no "Trip start date" was given.
- day_number: your best guess at this day's position in the OVERALL trip if there's a visible ordinal ("DAY SEVEN" → 7) — but this gets recalculated from `date` after merging regardless, so don't worry if you can't tell.
- location_heading: place name ONLY, never the activity type — never append the touring/excursion destination if it differs from the town the hotel is actually in (e.g. touring "Sigiriya" while the hotel itself is in "Habarana": heading is "Habarana", not "Sigiriya" and not "Sigiriya / Habarana" — the excursion name belongs in overview_bullets, not the heading). On a day that changes which city the hotel is in, use "CityA - CityB" (hyphen, the day's starting city to its ending city) — never a slash for this, the slash is reserved for the two cases below. "/ Ankunft" only on the actual arrival day, "/ Abreise" only on the actual departure day — append it to whichever city name is correct for that day (if the day both transfers AND departs, e.g. a transfer from the last hotel's city straight to the airport for an international flight home, it's "CityA - CityB / Abreise", not just the airport name or an unrelated city).
- overview_bullets: 2-5 short German noun phrases. Transfer/arrival only days: ["Transfer"] or ["Anreise / Flug"].
- day_marker: a SHORT (8-15 words) EXACT, character-for-character excerpt copied verbatim from the very START of this day's section — the day's header line if there is one, or the first distinctive sentence if not. Used afterward to locate the day in the original text and slice its real content, so precision matters more here than for any other field — never paraphrase, translate, reformat, or summarize it. Pick a phrase UNIQUE within the document — not a generic word/phrase that also occurs elsewhere. Every expanded day within an undifferentiated multi-night block shares the IDENTICAL day_marker pointing to where that block begins.
- hotel.is_first_night = true only on first arrival at each hotel (within what's visible in this excerpt — a day continuing an already-established hotel stay from before this excerpt should still be false).
- hotel.name must be the ACTUAL named property from the DMC, never a generic overnight marker some DMC documents use in place of a real name ("Ü in Vilnius", "Overnight Riga", "Hotel TBD") — that marker is a placeholder, not a hotel name: in that case return name: "". Likewise return room_type: null and meal_plan: null when the DMC states none for that stay — never invent a plausible-looking value.
- Weekdays in German. Meal plan: BB→Frühstück, HB→Halbpension, FB→Vollpension, AI→All-inclusive.
- Do NOT write body_paragraphs, client_name, destination, or any trip-level field — days only.
"""


def _chunk_text(text: str, chunk_size: int = 9000, overlap: int = 2000) -> list:
    """Splits text into overlapping windows — the overlap ensures a day
    whose text falls near a chunk boundary is still fully contained in at
    least one chunk, so the chunk prompt's "skip if visibly cut off" rule
    doesn't end up dropping it from every chunk."""
    if len(text) <= chunk_size:
        return [text]
    chunks = []
    start = 0
    while start < len(text):
        end = min(start + chunk_size, len(text))
        chunks.append(text[start:end])
        if end == len(text):
            break
        start = end - overlap
    return chunks


def _run_chunk_days(chunk_text: str, trip_start_date: str = "") -> list:
    """Extracts whatever complete days are visible in one chunk. A single
    attempt failing (rate limit, transient network error, malformed JSON)
    used to silently drop that chunk's days entirely with no retry — for a
    document whose chunks aren't fully redundant (a day long enough to
    straddle the overlap window), that meant real days vanishing from the
    final result with nothing to catch it. One quick retry after a bad
    attempt fixes the transient case; only a second consecutive failure
    gives up and skips the chunk (adjacent overlapping chunks likely cover
    the same days anyway).

    trip_start_date: the date of Day 1, if _run_metadata found one anywhere
    in the full document. Some DMC documents label days only "Day 01",
    "Day 02"... with the actual calendar date findable nowhere near the
    day-by-day narrative itself (e.g. only in a pricing/validity section) —
    a chunk covering just the narrative has no way to compute real dates on
    its own without this anchor, and used to either fabricate a placeholder
    or (after that was explicitly disallowed) leave every date blank,
    which drops the day entirely at the merge-by-date step.
    """
    user_msg = chunk_text
    if trip_start_date:
        user_msg = f"Trip start date (Day 1): {trip_start_date}\n\n{chunk_text}"
    for attempt in range(2):
        try:
            response = _ai_complete(
                model=AI_MODEL,
                temperature=0.1,
                max_tokens=16000,
                messages=[
                    {"role": "system", "content": STRUCTURE_CHUNK_PROMPT},
                    {"role": "user",   "content": user_msg},
                ],
            )
            raw = response.choices[0].message.content.strip()
            raw = re.sub(r"^```(?:json)?\s*", "", raw)
            raw = re.sub(r"\s*```$", "", raw)
            try:
                parsed = json.loads(raw, strict=False)
            except json.JSONDecodeError:
                parsed = json.loads(_repair_truncated_json(raw), strict=False)
            # The model occasionally returns the days array directly instead
            # of wrapping it in {"days": [...]}. That used to raise
            # AttributeError on .get and get swallowed by the retry handler
            # below, silently dropping every day in the chunk — accept both
            # shapes rather than losing real days to a formatting quirk.
            days = parsed if isinstance(parsed, list) else parsed.get("days", [])
            return [d for d in days if isinstance(d, dict)]
        except Exception as e:
            if attempt == 0:
                print(f"[structure-chunk] WARNING — a chunk failed, retrying once: {e}", flush=True)
                time.sleep(2)
            else:
                print(f"[structure-chunk] WARNING — chunk failed twice, skipping it: {e}", flush=True)
                return []


def _run_metadata(dmc_content: str) -> dict:
    """Trip-level fields only (client_name, destination, pax, leistungen) —
    a small, cheap, reliable call over the full document regardless of
    length, since its output never scales with the number of days. One
    retry on transient failure, same reasoning as _run_chunk_days."""
    for attempt in range(2):
        try:
            response = _ai_complete(
                model=AI_MODEL,
                temperature=0.1,
                max_tokens=4000,
                messages=[
                    {"role": "system", "content": STRUCTURE_METADATA_PROMPT},
                    {"role": "user",   "content": dmc_content},
                ],
            )
            raw = response.choices[0].message.content.strip()
            raw = re.sub(r"^```(?:json)?\s*", "", raw)
            raw = re.sub(r"\s*```$", "", raw)
            try:
                return json.loads(raw, strict=False)
            except json.JSONDecodeError:
                return json.loads(_repair_truncated_json(raw), strict=False)
        except Exception as e:
            if attempt == 0:
                print(f"[structure-metadata] WARNING — metadata failed, retrying once: {e}", flush=True)
                time.sleep(2)
            else:
                print(f"[structure-metadata] WARNING — metadata extraction failed twice: {e}", flush=True)
                return {}


def _merge_chunk_days(all_days: list) -> list:
    """Deduplicates days collected from multiple overlapping chunks by
    date — not day_number, which each chunk guesses independently and
    can't be trusted to agree across chunks — preferring the most complete
    entry per date (hotel present, more bullets, longer/more specific
    day_marker). Sorts by date and renumbers day_number sequentially
    afterward, so the final numbering is always internally consistent
    regardless of what any individual chunk guessed."""
    from datetime import datetime

    def _score(day: dict) -> tuple:
        return (
            bool((day.get("hotel") or {}).get("name")),
            len(day.get("overview_bullets") or []),
            len(day.get("day_marker") or ""),
        )

    by_date = {}
    for d in all_days:
        date_str = (d.get("date") or "").strip()
        if not date_str:
            continue
        if date_str not in by_date or _score(d) > _score(by_date[date_str]):
            by_date[date_str] = d

    def _parse(date_str):
        try:
            return datetime.strptime(date_str, "%d.%m.%Y")
        except Exception:
            return datetime.max

    merged = sorted(by_date.values(), key=lambda d: _parse(d.get("date", "")))
    for i, d in enumerate(merged):
        d["day_number"] = i + 1
    return merged


_DE_MONTHS_LONG = {
    1: "Januar", 2: "Februar", 3: "März", 4: "April", 5: "Mai", 6: "Juni",
    7: "Juli", 8: "August", 9: "September", 10: "Oktober", 11: "November", 12: "Dezember",
}
_DE_WEEKDAYS = ["Montag", "Dienstag", "Mittwoch", "Donnerstag", "Freitag", "Samstag", "Sonntag"]


def _call_ai_structure_chunked(dmc_content: str) -> dict:
    """The long/dense-document path: splits into overlapping chunks (each
    small enough to extract reliably), merges by date, and computes every
    trip-level date field directly from the merged days themselves — never
    trusted from a single completion that had to hold the whole trip in
    its "working set" at once, which is the failure mode this exists to
    avoid (see call_ai_structure's docstring)."""
    from datetime import datetime

    if len(dmc_content) > 100000:
        dmc_content = dmc_content[:100000] + "\n[...truncated...]"

    chunks = _chunk_text(dmc_content)
    print(f"[structure-chunk] document is {len(dmc_content)} chars — split into {len(chunks)} chunks", flush=True)

    def _attempt():
        metadata = _run_metadata(dmc_content)
        # Some trips ("choose your own start date" templates) genuinely
        # have no calendar date anywhere in the source — metadata correctly
        # returns "" for those rather than inventing one. But every chunk
        # still needs SOME consistent anchor to compute ordinal days
        # ("Day Four") into real dates; without one, a chunk with no
        # in-view date of its own has nothing to compute from at all,
        # which is a real, observed cause of chunks producing nothing for
        # their days. Falling back to a placeholder default (same pattern
        # already used elsewhere for a genuinely dateless source) keeps
        # every chunk's date math consistent — the placeholder gets edited
        # by hand afterward regardless.
        anchor_date = metadata.get("start_date") or "01.01.2026"
        expected_days = metadata.get("expected_days")
        all_days = []
        for chunk in chunks:
            all_days.extend(_run_chunk_days(chunk, trip_start_date=anchor_date))

        merged_days = _merge_chunk_days(all_days)
        unresolved = _slice_activities_by_markers(merged_days, dmc_content)

        result = {
            "client_name": metadata.get("client_name", "Familie"),
            "destination": metadata.get("destination", ""),
            "pax": metadata.get("pax"),
            "days": merged_days,
            "leistungen": metadata.get("leistungen", {}),
        }

        if merged_days:
            first, last = merged_days[0], merged_days[-1]
            result["start_date"] = first.get("date", "")
            result["end_date"] = last.get("date", "")
            try:
                start_dt = datetime.strptime(first["date"], "%d.%m.%Y")
                end_dt = datetime.strptime(last["date"], "%d.%m.%Y")
                result["start_date_formatted"] = f"{start_dt.day:02d}. {_DE_MONTHS_LONG[start_dt.month]}"
                result["end_date_formatted"] = f"{end_dt.day:02d}. {_DE_MONTHS_LONG[end_dt.month]} {end_dt.year}"
                nights = (end_dt - start_dt).days
                result["duration_label"] = f"{nights + 1} Tage / {nights} Übernachtungen"
                result["leistungen"].setdefault(
                    "reisedatum",
                    f"{result['start_date_formatted']} bis {result['end_date_formatted']}",
                )
                result["leistungen"].setdefault("reisedauer", result["duration_label"])
                for d in merged_days:
                    if not d.get("weekday"):
                        dt = datetime.strptime(d["date"], "%d.%m.%Y")
                        d["weekday"] = _DE_WEEKDAYS[dt.weekday()]
            except Exception:
                pass

        return result, _validate_structure(result, dmc_content, unresolved, expected_override=expected_days)

    result, v = _attempt()

    # Unlike the single-pass path, chunked extraction had no retry at all —
    # a day silently dropped by one bad chunk (see _run_chunk_days) stayed
    # dropped. One full re-run, keeping whichever attempt scores better,
    # closes that gap the same way the single-pass path already handles it.
    if v["missing"] or v["unresolved"] or v["bad_hotels"]:
        print(f"[structure-chunk] WARNING — issues after merge: {v} — retrying once", flush=True)
        retry_result, retry_v = _attempt()
        if _issue_score(retry_v) < _issue_score(v):
            result, v = retry_result, retry_v

    if v["missing"]:
        result["_day_count_mismatch"] = {"expected": v["expected"], "actual": v["actual"]}
    if v["unresolved"]:
        result["_hollow_days"] = v["unresolved"]
    if v["bad_hotels"]:
        result["_hallucinated_hotel_days"] = v["bad_hotels"]
    if v["missing"] or v["unresolved"] or v["bad_hotels"]:
        print(f"[structure-chunk] WARNING — issues remain after retry: {v}", flush=True)

    _sanitize_structure(result)
    _backfill_missing_hotels(result.get("days", []))
    _clear_departure_day_hotels(result.get("days", []))
    result["days"] = _merge_undifferentiated_days(result.get("days", []))
    return result


DAY_PROSE_PROMPT = """You are writing one day of a German-language Reiseverlauf for BAWA Tours & Travel, a German luxury travel agency. Your output will be printed and handed directly to clients.

Return ONLY valid JSON — no markdown, no explanation:
{"body_paragraphs": ["...", "..."], "hotel_description": "..."}

---
HOUSE STYLE — based directly on BAWA's own confirmed itineraries
---
The examples below are taken from real, sent-to-client BAWA documents. Match this style exactly.

WRITING STYLE:
- Direct address to the client: "Besuchen Sie...", "Sie erkunden...", "Anschließend spazieren Sie durch..."
- Each sight or activity gets ONE TO THREE sentences. Most get one or two. Do not write a full paragraph of buildup for a single sight.
- Get straight to the fact or the action. Do not open with filler like "bietet eine gute Gelegenheit" or "ist eine gute Möglichkeit, um..." — real BAWA documents never use this.
- No invented framing devices like "Insider-Tipp:" or "Ein interessanter Fakt:" — state the fact as part of the sentence.
- Use real, checkable specifics: a year, a height, a builder's name, a UNESCO designation. Only state a fact if you are genuinely confident it is accurate.
- Vary how sentences open, but don't force literary flourish — confident and direct, not ornate.

REAL EXAMPLES FROM BAWA DOCUMENTS (this is the actual target quality and length):

  "Besuchen Sie Harajuku, das mit seinen angesagten Modeboutiquen und zahlreichen Essensmöglichkeiten begeistert. Probieren Sie lokale Snacks, lassen Sie sich von kreativer Streetwear inspirieren und erleben Sie die jugendliche Energie dieses einzigartigen Viertels."

  "Besuch des Tokyo Sky Tree, dem höchsten Fernsehturm der Welt. Er ist 634m hoch und wurde im Jahr 2012 errichtet. Auf zwei Aussichtsgeschossen genießen Sie einen atemberaubenden Blick auf den Großraum Tokyo."

  "Danach überqueren Sie gemeinsam mit Ihrem Guide die größte Kreuzung der Welt am Shibuya Terminal, die vor allem durch den Film „Lost in Translation" weltberühmt wurde."

  "Der Vulkan ist 3.776 Meter hoch und zählt seit 2013 zum UNESCO Weltkulturerbe."

  "Sie beginnen mit dem goldenen Pavillon Kinkakuji (UNESCO Weltkulturerbe), der Ende des 14. Jahrhunderts als Alterssitz für Shogun Ashikaga Yoshimitsu errichtet wurde."

Notice these are short. A famous landmark can run longer — but even those build from short, concrete sentences, not dense paragraphs.

---
SIGHTSEEING — length and depth
---
- A brief DMC mention of a minor stop with no real history behind it → 1-2 sentences. Do not pad it.
- A landmark with genuine historical, cultural, or architectural significance (a former checkpoint, a centuries-old shrine/temple, a UNESCO site, a bridge tied to a legend, a castle, a major museum) → write 3-5 sentences and cover AT LEAST TWO of:
  - founding/construction year or era, and who built it or why
  - its historical role or the story behind it (what it controlled, protected, commemorated, or is famous for)
  - one specific narrative detail — a legend, a named historical figure, a restoration date, a physical detail (height, material, a named feature)
  - what the client will concretely see or experience there
  A single flat sentence ("X war ein wichtiger Kontrollpunkt...") is too thin for a landmark like this — if you have more genuine knowledge, use it, the way you would for a hotel.
- A famous modern district, crossing, or shopping street (Shibuya Crossing, Ginza, Dotonbori, Harajuku/Takeshita-dori, Times Square-style landmarks) is NOT a "minor stop" just because it has no centuries-old history — these are exactly as well-documented as a temple, just in a different register. Write 2-4 sentences and name at least two concrete, checkable specifics: a named store/department store/landmark building, a pop-culture or film reference that made it famous, a superlative with a real number (busiest crossing — how many people cross at once; a street's length; a district's size), or what specifically the client will do or buy there. "Elegantes Einkaufsviertel mit luxuriösen Boutiquen" names nothing — it could describe any shopping street in the world. Reuse it and the reader learns nothing new.
- Never write a generic scene-setting sentence that doesn't name a real fact. If you don't have a specific, genuine fact about a place, write only what the client will do there — do not invent atmosphere.
- Do not repeat the same descriptive adjective across multiple sights (e.g. do not call several things "atemberaubend" or "einzigartig").
- If the DMC only gives a vague regional description ("full day tour of Hakone", "sightseeing in Nara") without naming specific sights, use your OWN genuine knowledge to name the well-known attractions a guided day trip there would realistically include (e.g. Nikko → Nikko Toshogu Schrein, Shinkyo-Brücke; Nara → Todaiji-Tempel, Nara-Park; Hiroshima/Miyajima → Friedensdenkmal, Itsukushima-Schrein). Name at least one or two real, specific sights — a day must never stay purely generic about the region's reputation with nothing concrete named.

---
HOTEL DESCRIPTIONS — first night at each hotel only
---
Real BAWA hotel descriptions are specific — built from real, checkable facts about that exact property, never generic luxury filler ("bietet eine harmonische Verbindung aus Tradition und Komfort" says nothing — cut it).

For an internationally known brand or any hotel you have real knowledge of (this covers most named international hotels — Andaz, Conrad, Aman, Six Senses, Park Hyatt, Ritz-Carlton and similarly documented properties all qualify), write 4-6 sentences and cover AT LEAST THREE of:
- opening or major renovation year
- architect or interior designer
- one named restaurant, bar, or spa concept
- exact neighbourhood/building and what that location means for the guest (a landmark, a view, a district)
- room count or a specific room category/signature suite
- an award or position in the luxury hotel world (Forbes Travel Guide, Michelin, Relais & Châteaux, Leading Hotels of the World)

  "Seit März 2024 empfängt Janu Tokyo Gäste im kreativen Viertel Azabudai Hills in Minato. Das Hotel bietet 122 elegant gestaltete Zimmer und Suiten des Architekten Jean-Michel Gathy, acht herausragende Restaurants, ein hochmodernes Wellnesscenter sowie einen ruhigen Rückzugsort im Herzen Tokios."

Every sentence must name something SPECIFIC and CHECKABLE: an opening date, an architect, a room count, a named restaurant, a real location detail. A description with only one such sentence is too thin — keep adding real facts until at least three of the categories above are covered.

Close with ONE final sentence stating the property's USP — why BAWA selected this specific hotel for this stay, not generic praise. Ground it in something already named in the description (its exact location relative to what the client will do that day, its design/wellness identity, its scale — intimate vs. grand, its suitability for the traveling party if known from the day's context) rather than inventing a new unrelated claim. Examples of the right register:

  "Damit ist das Hotel der ideale Ausgangspunkt, um die Tempel und Gärten Kyotos zu Fuß zu erkunden."
  "Die kompakte Zimmerzahl und das durchdachte Design machen es zum idealen Rückzugsort nach ereignisreichen Tagen in der Millionenmetropole."
  "Für eine Familienreise bietet das Haus mit seinen großzügigen Suiten und dem ruhigen Innenhof genau die Mischung aus Komfort und Privatsphäre, die diesen Aufenthalt besonders macht."

Only for a genuinely small or undocumented property where you truly have nothing beyond the name and city should you drop to 2 sentences using only what you can verify (and skip the USP sentence if there is nothing genuine to ground it in), or return an empty string rather than inventing generic luxury filler.

hotel_description: only on first night at a hotel (is_first_night=true). Empty string on all other nights.

---
HAUSSTIL-REFERENZEN (IF PRESENT IN THE USER MESSAGE)
---
The user message may include a "HAUSSTIL-REFERENZEN" section: real text BAWA
has already written and sent to clients for this exact hotel or sight.

- Hotel text (no placeholder): use it as the primary source for
  hotel_description — reuse its facts and phrasing, only trimming or lightly
  adapting it to fit this stay. Do not invent competing facts. This reference
  text predates the USP-sentence rule above and will not have one — still
  append your own closing USP sentence per the HOTEL DESCRIPTIONS rules. It
  also predates the "Guide" terminology rule above — if it says "Reiseleiter"
  or "Reiseleitung", change that to "Guide" even though you're otherwise
  reusing this text closely; that one substitution is not a "competing fact".
- Sightseeing text behind a {{SIGHT:n}} placeholder: this exact wording is
  already approved and MUST be reused verbatim — do not paraphrase, shorten,
  or rewrite it. Do not write your own sentences about that same sight.
  Instead place the literal placeholder string (e.g. "{{SIGHT:0}}") as its
  own element in body_paragraphs, positioned where that sight belongs in the
  day's narrative. The exact text gets substituted in afterward — this
  substitution happens outside your control, so an older reference paragraph
  using "Reiseleiter" instead of "Guide" is corrected automatically later,
  not something you need to (or can) fix yourself here.
- A sight marked "bestätigter BAWA-Fakt, aber zu kurz" (no placeholder): treat
  exactly like the semantic-match case below — a confirmed fact to expand on
  with real knowledge, not text to reproduce as-is.

Only fall back to your own knowledge for hotels/sights the reference section
does not cover.

---
GENERAL LANGUAGE RULES
---
- Write entirely in German. Compose original German prose from the facts given — do not translate literally.
- Correct German typography: ä ö ü ß always — never ae/oe/ue/ss substitutions.
- German quotation marks „so" not "so".
- No Anglicisms unless standard in German travel language (Check-in, Transfer are fine).
- No meal mentions in body paragraphs.
- Always call the guide "Guide" — never "Reiseleiter", "Reiseleiterin", "Reiseleitung", "Reiseführer", "Reiseexperte"/"Reiseexpertin" (the DMC source's English "travel expert" translates to "Guide", not "Reiseexperte"), even for a local guide ("lokaler Guide", not "lokaler Reiseleiter"/"lokaler Reiseexperte"). E.g. "Ihr Guide erwartet Sie", "Treffen Sie Ihren Guide", "mit Ihrem deutschsprachigen Guide" — never "Ihre Reiseleitung"/"Ihren Reiseleiter"/"Ihren Reiseexperten".
- Logistics (transfers, trains) stay short and factual: "Treffen Sie Ihren Fahrer für den privaten Transfer zum Bahnhof (ca. 50 Minuten)."
- If the DMC source gives a specific time or time range for an activity (e.g. "Time: 08:30 - 13:00", "10:00 Meet your guide"), include it in the sentence — "Von 08:30 bis 13:00 Uhr...", "Um 10:00 Uhr treffen Sie...". Don't invent a time that isn't in the source, but never drop one that is.
"""


def generate_cover_subtitle(destination: str, overview: str) -> str:
    """Generate a short evocative German subheading for the cover page."""
    try:
        prompt = (
            f"Schreibe einen kurzen, poetischen deutschen Untertitel (max. 6 Wörter) "
            f"für eine Luxusreise nach {destination}. "
            f"Reisehighlights: {overview[:300]}. "
            f"Kein Anführungszeichen, kein Punkt am Ende. Nur den Untertitel, nichts sonst."
        )
        resp = _ai_complete(
            model=AI_MODEL,
            temperature=0.9,
            max_tokens=40,
            messages=[{"role": "user", "content": prompt}],
        )
        return resp.choices[0].message.content.strip().strip('"').strip("'").rstrip(".")
    except Exception:
        return "Eine Reise voller Eindrücke"


def _expected_day_count(start_date: str, end_date: str):
    """Number of calendar days a trip spans, from DD.MM.YYYY start/end dates
    (inclusive of both). None if either date is missing/unparseable."""
    from datetime import datetime
    try:
        fmt = "%d.%m.%Y"
        return (datetime.strptime(end_date, fmt) - datetime.strptime(start_date, fmt)).days + 1
    except Exception:
        return None


def _find_marker(haystack: str, marker: str, start: int = 0) -> int:
    """Locates `marker` in haystack[start:], tolerant of whitespace/case
    differences the AI can introduce even when asked to copy "verbatim" —
    returns the absolute position in haystack, or -1 if not found."""
    marker = (marker or "").strip()
    if not marker:
        return -1
    window = haystack[start:]

    pos = window.find(marker)
    if pos != -1:
        return start + pos

    pos = window.lower().find(marker.lower())
    if pos != -1:
        return start + pos

    # Whitespace-normalized search (collapse all runs of whitespace to a
    # single space on both sides), mapping each normalized index back to
    # its position in the original text.
    norm_chars, orig_positions, prev_space = [], [], False
    for idx, ch in enumerate(window):
        if ch.isspace():
            if not prev_space:
                norm_chars.append(" ")
                orig_positions.append(idx)
            prev_space = True
        else:
            norm_chars.append(ch.lower())
            orig_positions.append(idx)
            prev_space = False
    norm_window = "".join(norm_chars)
    norm_marker = re.sub(r"\s+", " ", marker.lower()).strip()
    npos = norm_window.find(norm_marker)
    if npos != -1:
        return start + orig_positions[npos]

    # Last resort: the AI may have appended or slightly altered the tail of
    # the marker — retry with just its first few words.
    words = marker.split()
    if len(words) > 6:
        return _find_marker(haystack, " ".join(words[:6]), start)
    return -1


def _slice_activities_by_markers(days: list, dmc_content: str) -> list:
    """Populates each day's activities_raw by locating its `day_marker` in
    the ORIGINAL source text and slicing the verbatim text between this
    day's marker and the next resolvable day's marker. This is what
    replaced asking the AI to reproduce activities_raw itself: on a long,
    dense document the model would lose track partway through and either
    truncate or fabricate entirely fictional content for later days (see
    call_ai_structure's docstring) — slicing the real source text in
    Python instead makes that class of error structurally impossible, at
    the cost of only a locate-the-boundary task for the AI, which is far
    lighter than faithfully reproducing paragraphs of prose.

    Markers are resolved in day order with a forward-only search cursor,
    so a generic phrase that happens to repeat elsewhere in the document
    can't match an earlier or later occurrence than the one intended.
    Consecutive days sharing an identical marker (a multi-night block with
    no day-by-day breakdown — every expanded day points at the same block
    start) get the block's full text on the first of them and an empty
    string on the rest, rather than an empty slice for all of them.

    Returns the day_numbers whose marker couldn't be located at all —
    day-prose generation already has its own fallback for thin/missing
    source text, so an unresolved day still degrades gracefully.
    """
    n = len(days)
    positions = [-1] * n
    search_from = 0
    for i, day in enumerate(days):
        marker = (day.get("day_marker") or "").strip()
        prev_marker = (days[i - 1].get("day_marker") or "").strip() if i > 0 else None
        if i > 0 and marker and marker == prev_marker and positions[i - 1] != -1:
            # Identical to the previous day's marker (a multi-night block
            # with no day-by-day breakdown) — reuse its position directly.
            # Re-searching from the cursor would fail here: the cursor has
            # already advanced past the marker's one and only occurrence,
            # so a fresh search from that point always finds nothing,
            # wrongly landing this day in "unresolved" instead of
            # "duplicate of the previous day".
            positions[i] = positions[i - 1]
            continue
        pos = _find_marker(dmc_content, marker, search_from)
        positions[i] = pos
        if pos != -1:
            search_from = pos + 1

    # A run of 3+ consecutive days can share the same marker (e.g. a
    # 4-night "all days free at leisure" block expanded to one entry per
    # calendar day). Comparing only to the immediately preceding day would
    # keep the real text on the first day of the run and blank out every
    # other day in it — give the whole run the same sliced text instead.
    unresolved = []
    i = 0
    while i < n:
        start = positions[i]
        if start == -1:
            days[i]["activities_raw"] = ""
            unresolved.append(days[i].get("day_number"))
            i += 1
            continue
        run_end = i
        while run_end + 1 < n and positions[run_end + 1] == start:
            run_end += 1
        end = len(dmc_content)
        for j in range(run_end + 1, n):
            if positions[j] != -1 and positions[j] > start:
                end = positions[j]
                break
        text = dmc_content[start:end].strip()[:4000]
        for k in range(i, run_end + 1):
            days[k]["activities_raw"] = text
        i = run_end + 1

    return unresolved


# Fields the AI returns as a JSON list, which every downstream renderer
# iterates over and calls string methods on.
_DAY_LIST_FIELDS = ("overview_bullets", "body_paragraphs")
_LEISTUNGEN_LIST_FIELDS = ("special_experiences", "hotel_nights")
# Scalar day fields the renderers read with direct day["..."] indexing, which
# raises KeyError (not a blank) when absent. A day added by hand in the editor
# ("+ Tag hinzufügen") or a partially-filled day from a reloaded draft can be
# missing any of them — an observed 500 on /build-docx was exactly this.
_DAY_SCALAR_FIELDS = ("day_number", "weekday", "date", "location_heading")


def _sanitize_structure(result: dict) -> dict:
    """Normalizes an AI-produced structure so downstream rendering can trust
    its shape.

    JSON `null` and a missing key are different things, and `.get(key,
    default)` only substitutes the default for the SECOND one — so a field
    the model explicitly returned as null sails straight past every
    `.get("body_paragraphs", [])` guard in the codebase and reaches a
    `for p in None` or `None.strip()`. That is a real, repeatedly-observed
    crash source here (it produced the generic "Server-Fehler" users
    reported), and it recurs because the guard *looks* correct at every
    individual call site.

    Rather than patch each site defensively — easy to miss one, and the next
    new renderer starts the cycle over — the shape is fixed once, here,
    immediately after parsing: null lists become empty lists, and null/
    non-string items inside those lists are dropped. Callers downstream can
    then rely on "list of real strings" holding.

    Deliberately does NOT invent content: a null hotel or a null client name
    stays null (those are meaningful — see the prompts' anti-hallucination
    rules); only the container shape is corrected.
    """
    days = result.get("days")
    if not isinstance(days, list):
        days = []
    result["days"] = [d for d in days if isinstance(d, dict)]

    for day in result["days"]:
        for field in _DAY_SCALAR_FIELDS:
            if day.get(field) is None:
                day[field] = ""
        for field in _DAY_LIST_FIELDS:
            value = day.get(field)
            if not isinstance(value, list):
                day[field] = [] if value is None else value
                continue
            day[field] = [item for item in value if isinstance(item, str) and item.strip()]
        # hotel is legitimately None (hotel-less days), but a non-dict truthy
        # value would break every hotel.get(...) call downstream.
        if day.get("hotel") is not None and not isinstance(day.get("hotel"), dict):
            day["hotel"] = None

    leistungen = result.get("leistungen")
    if not isinstance(leistungen, dict):
        leistungen = {}
    result["leistungen"] = leistungen
    for field in _LEISTUNGEN_LIST_FIELDS:
        value = leistungen.get(field)
        if not isinstance(value, list):
            leistungen[field] = []
            continue
        if field == "hotel_nights":
            leistungen[field] = [item for item in value if isinstance(item, dict)]
        else:
            leistungen[field] = [item for item in value if isinstance(item, str) and item.strip()]

    return result


def _strip_location_suffix(location: str) -> str:
    return re.sub(r"\s*/\s*(Ankunft|Abreise)\s*$", "", location or "", flags=re.IGNORECASE).strip().lower()


def _backfill_missing_hotels(days: list) -> None:
    """Carries the previous day's hotel forward onto a day whose hotel is
    missing, when both days are at the same place. The per-day hotel field
    and leistungen.hotel_nights (a separate, cheap, trip-level call) are
    extracted independently — a document can come back with the right
    hotel in the services summary while a specific day's own hotel field
    is empty, silently dropping that day's "Übernachtung im X" line even
    though the itinerary elsewhere shows the stay was correctly identified.

    Deliberately conservative: only fires when the (suffix-stripped)
    location_heading exactly matches the last day that had a real hotel —
    guessing across an actual city change would silently attach the wrong
    hotel, which is worse than a visibly missing one. Also never fires on
    a departure day ("/ Abreise") — that day has no overnight stay by
    design (the client checks out and leaves), so a missing hotel there is
    correct, not a gap to fill. Without this exclusion, a departure day at
    the same city as the previous night incorrectly inherited that night's
    hotel, showing an "Übernachtung im X" line for a night that never
    happens.
    """
    last_hotel = None
    last_location = None
    for day in days:
        hotel = day.get("hotel")
        name = (hotel.get("name") or "").strip() if isinstance(hotel, dict) else ""
        heading = day.get("location_heading") or ""
        location = _strip_location_suffix(heading)
        is_departure = bool(re.search(r"/\s*Abreise\s*$", heading, re.IGNORECASE))
        if name:
            last_hotel = hotel
            last_location = location
        elif last_hotel and location and location == last_location and not is_departure:
            day["hotel"] = {**last_hotel, "is_first_night": False, "description": ""}


def _clear_departure_day_hotels(days: list) -> None:
    """A departure day ("/ Abreise") never has an overnight stay by
    design — the client checks out and leaves that day. This is a
    safety net independent of _backfill_missing_hotels: the AI's own
    direct extraction can still assign a hotel to the departure day
    (e.g. carrying the previous night's hotel over on its own, not via
    the backfill path), which showed up as a wrong "Übernachtung im X"
    line on the last day of the trip. Runs after backfill so it wins
    regardless of where the hotel value came from.
    """
    for day in days:
        heading = day.get("location_heading") or ""
        if re.search(r"/\s*Abreise\s*$", heading, re.IGNORECASE):
            day["hotel"] = None


def _merge_undifferentiated_days(days: list) -> list:
    """Collapses a run of 2+ consecutive non-free days sharing byte-identical
    activities_raw into a single entry spanning the whole range, instead of
    repeating the same content under several near-identical day headings.

    This happens when the DMC source describes a stretch of days as one
    undifferentiated block with no day-by-day breakdown at all — e.g. a
    "choose your own activities" week with a menu of options but no fixed
    schedule ("you'll receive the activity order upon arrival"). Each
    calendar day still gets its own day_marker pointing at the same block
    (correct — every day of the block really is covered by that text), but
    showing the identical content 6 times in a row reads as broken rather
    than as the deliberate, unavoidable ambiguity it actually is.

    Free days (is_free_day) are untouched — those already get their own
    consecutive-run merge at document-build time (see build_body_xml),
    which intentionally keeps every date visible even though the text is
    just the fixed "free time" line.
    """
    merged = []
    i, n = 0, len(days)
    while i < n:
        day = days[i]
        raw = (day.get("activities_raw") or "").strip()
        run_end = i
        if raw and not day.get("is_free_day"):
            while (
                run_end + 1 < n
                and not days[run_end + 1].get("is_free_day")
                and (days[run_end + 1].get("activities_raw") or "").strip() == raw
            ):
                run_end += 1
        if run_end == i:
            merged.append(day)
            i += 1
            continue
        combined = dict(day)
        combined["date_end"] = days[run_end].get("date", "")
        combined["span_days"] = run_end - i + 1
        merged.append(combined)
        i = run_end + 1

    for idx, d in enumerate(merged):
        d["day_number"] = idx + 1
        if d.get("span_days", 1) > 1:
            d["day_number_end"] = d["day_number"] + d["span_days"] - 1

    return merged


def _validate_structure(result: dict, dmc_content: str, unresolved: list = None, expected_override: int = None) -> dict:
    """Cheap, deterministic checks against the result and the original
    source text — catches problems an AI call can introduce that a single
    "did the day count match" check misses:

    - missing: fewer days than expected. Preferably expected_override — a
      day count read from the document's own stated trip length (e.g. "7
      Days"), independent of what the extraction actually returned. Without
      it, expected falls back to computing from the result's own
      start_date/end_date — which is blind to a chunk silently dropping the
      tail end of the trip, since a truncated result's own date range looks
      perfectly self-consistent (a real, observed failure: a 7-day trip's
      last chunk produced nothing, and the result's own start/end dates
      only spanned the 4 days that WERE found, so expected == actual == 4
      and nothing flagged it — this is exactly what expected_override
      exists to catch).
    - unresolved: day_marker couldn't be located anywhere in the source
      text (see _slice_activities_by_markers) — usually because the AI
      paraphrased it instead of copying it verbatim, or invented a marker
      for a day that doesn't really exist in the source.
    - bad_hotels: a day's hotel name doesn't appear anywhere in the source
      text at all — a strong, cheap signal the model fabricated a
      plausible-sounding hotel instead of reading the actual document for
      that day. Confirmed happening in practice: forcing an exact day count
      via a naive retry made the model correctly hit the count but
      hallucinate an entirely different (fictional) hotel/city for several
      of the days it needed to "fill in".
    - stalled_run: a day_marker repeated for MORE consecutive days than a
      legitimate undifferentiated multi-night block realistically would
      (see the threshold below) — a sign the model correctly extracted the
      first several days, then gave up and duplicated the last one it
      resolved for the remainder, rather than fabricating anything new.
      This is the one real failure this validation had to be added for
      last: day count matched, no marker was individually unresolvable
      (they're all "found" — just the wrong day's marker reused), and
      every hotel name genuinely appears in the source (also just for the
      wrong day) — so the missing/unresolved/bad_hotels checks above all
      pass while most of the document silently goes unattributed.
    - low_coverage: total activities_raw length across all days is
      suspiciously small relative to the source document's total length —
      the same underlying failure as stalled_run, caught from a different
      angle in case the duplicated run is broken up rather than one
      contiguous block.
    """
    days = result.get("days", [])
    expected = expected_override or _expected_day_count(result.get("start_date", ""), result.get("end_date", ""))
    actual = len(days)
    missing = (expected - actual) if (expected and actual < expected) else 0

    unresolved = unresolved or []

    # Some PDFs hyphenate a word right at a line wrap (e.g. "Hampton by
    # Hilton Ho-\ntel or similar") — real content, but it breaks a plain
    # substring search since the name isn't contiguous in the extracted
    # text anymore. Rejoining "-\n" before comparing fixes that without
    # weakening the check for a genuinely fabricated hotel name.
    content_lower = re.sub(r"-\s*\n\s*", "", dmc_content.lower())
    bad_hotels = []
    for d in days:
        name = ((d.get("hotel") or {}).get("name") or "").strip()
        if name and name.lower() not in content_lower:
            bad_hotels.append(d.get("day_number"))

    stalled_run = 0
    run = 1
    for i in range(1, len(days)):
        prev = (days[i - 1].get("day_marker") or "").strip()
        cur = (days[i].get("day_marker") or "").strip()
        if cur and cur == prev:
            run += 1
            stalled_run = max(stalled_run, run)
        else:
            run = 1

    low_coverage = False
    if len(dmc_content) > 3000:
        consumed = sum(len(d.get("activities_raw") or "") for d in days)
        low_coverage = consumed < 0.55 * len(dmc_content)

    return {
        "expected": expected, "actual": actual, "missing": missing,
        "unresolved": unresolved, "bad_hotels": bad_hotels,
        "stalled_run": stalled_run, "low_coverage": low_coverage,
    }


# A legitimate undifferentiated multi-night block (e.g. "13 Jan – 18 Jan
# (6 Nights)") realistically runs a handful of nights at most — a run
# longer than this is far more likely the model stalling out and
# duplicating its last resolved day forward than a genuine block this long.
_STALLED_RUN_THRESHOLD = 4


def _issue_score(v: dict) -> int:
    """Lower is better — used to pick the more correct of two candidate
    extraction results (the retry isn't guaranteed to be an improvement).
    stalled_run/low_coverage are weighted heavily since either one means
    most of the document is silently going unattributed — a far worse
    outcome than a handful of individually-bad-hotel days."""
    score = v["missing"] + len(v["unresolved"]) + len(v["bad_hotels"])
    if v["stalled_run"] > _STALLED_RUN_THRESHOLD:
        score += v["stalled_run"]
    if v["low_coverage"]:
        score += 5
    return score


# Above this length, a single completion asked to track every day's
# hotel/location/marker at once becomes unreliable in practice — verified
# against a real 22,000-char, 15-day document that reliably extracted only
# its first ~6 days correctly (or worse, fabricated content for the rest)
# regardless of retries. Below it, the existing single-pass path already
# tests reliably and doesn't need the extra AI calls chunking costs.
_CHUNK_THRESHOLD = 9000


def call_ai_structure(dmc_content: str) -> dict:
    """Fast structure-only extraction — no prose writing. Delegates to the
    chunked path for long/dense documents (see _call_ai_structure_chunked)
    and the single-pass path otherwise."""
    if len(dmc_content) > _CHUNK_THRESHOLD:
        return _call_ai_structure_chunked(dmc_content)
    return _call_ai_structure_single(dmc_content)


def _call_ai_structure_single(dmc_content: str) -> dict:
    """The proven path for documents short enough to extract reliably in
    one pass — see call_ai_structure's module-level threshold.

    A long itinerary (15+ days, each with hotel details and bullets) can
    still produce a JSON response too large for a modest max_tokens budget
    — Gemini's output then gets cut off mid-day, silently dropping every
    day after the cutoff rather than raising an error. _repair_truncated_json()
    salvages whatever completed before the cutoff and flags the result
    ("_truncated": True) — retry once with a much larger budget when that
    flag is set, instead of accepting the shorter itinerary.

    activities_raw is NOT written by the AI at all — asking it to
    faithfully reproduce potentially 2000+ characters of verbatim source
    text per day, for 15+ days, in one long generation, turned out to be
    unreliable in practice: on a long, dense real document the model
    correctly handled the first 5-6 days and then, for the rest, either
    truncated outright or (worse, and easy to miss) fabricated entirely
    fictional hotels/content that never appeared anywhere in the source,
    while still nominally hitting the right day count. Instead, the AI
    only supplies a short `day_marker` — a verbatim excerpt marking where
    each day's section starts — and _slice_activities_by_markers() locates
    those markers in the ORIGINAL source text and slices activities_raw in
    Python. This makes fabricated content structurally impossible (a slice
    is either a real excerpt of the source or empty), and is a far lighter
    task for the model than reproducing long text blocks, since it only
    has to identify boundaries.

    _validate_structure() then checks the result for problems a naive
    day-count check misses (see its docstring), and if any are found,
    retries ONCE with a single message describing every detected issue
    together. A separate retry per issue type was tried first and made
    things worse: correcting one issue at a time gave the model repeated
    opportunities to "fix" the symptom being checked while introducing a
    different failure for the days it had to fill in. Whichever of the
    original or retried result has fewer total issues wins; anything still
    wrong after that one retry is flagged on the result for the frontend
    to warn about rather than silently shipped.
    """
    if len(dmc_content) > 50000:
        dmc_content = dmc_content[:50000] + "\n[...truncated...]"

    def _run(user_content: str) -> tuple:
        result = None
        for max_tok in (32000, 60000):
            response = _ai_complete(
                model=AI_MODEL,
                temperature=0.1,
                max_tokens=max_tok,
                messages=[
                    {"role": "system", "content": STRUCTURE_PROMPT},
                    {"role": "user",   "content": user_content},
                ],
            )
            raw = response.choices[0].message.content.strip()
            raw = re.sub(r"^```(?:json)?\s*", "", raw)
            raw = re.sub(r"\s*```$", "", raw)
            try:
                # strict=False tolerates a raw control character (e.g. a
                # literal newline) inside a string value instead of the
                # escaped \n — Gemini occasionally copies one verbatim like
                # that, which strict JSON parsing rejects outright even
                # though the structure is otherwise valid.
                parsed = json.loads(raw, strict=False)
                truncated = False
            except json.JSONDecodeError:
                parsed = json.loads(_repair_truncated_json(raw), strict=False)
                truncated = bool(parsed.get("_truncated"))
                if not truncated:
                    unresolved = _slice_activities_by_markers(parsed.get("days", []), dmc_content)
                    return parsed, unresolved
                print(
                    f"[structure] WARNING — output truncated at max_tokens={max_tok} "
                    f"({len(parsed.get('days', []))} days recovered) — retrying with a larger budget",
                    flush=True,
                )
                result = parsed
                continue
            unresolved = _slice_activities_by_markers(parsed.get("days", []), dmc_content)
            return parsed, unresolved
        print(
            f"[structure] WARNING — structure extraction still truncated after retry; "
            f"returning {len(result.get('days', []))} days. The source document may be unusually long.",
            flush=True,
        )
        unresolved = _slice_activities_by_markers(result.get("days", []), dmc_content)
        return result, unresolved

    result, unresolved = _run(f"DMC offer:\n\n{dmc_content}")
    v = _validate_structure(result, dmc_content, unresolved)

    stalled = v["stalled_run"] > _STALLED_RUN_THRESHOLD
    if v["missing"] or v["unresolved"] or v["bad_hotels"] or stalled or v["low_coverage"]:
        problems = []
        if v["missing"]:
            problems.append(
                f"Die Analyse ergab nur {v['actual']} von {v['expected']} erwarteten Tagen (Reise von "
                f"{result.get('start_date')} bis {result.get('end_date')}). Jeder Kalendertag braucht "
                f"einen eigenen Eintrag — falls das Angebot einen mehrnächtigen Aufenthalt als einen "
                f"Zeitraum beschreibt (z.B. \"(6 Nights)\" oder \"FROM DAY X TO DAY Y\"), teile ihn in "
                f"einzelne Tage auf, gleiches Hotel auf jedem, is_free_day: true wo keine Aktivität "
                f"genannt ist."
            )
        if v["unresolved"]:
            problems.append(
                f"Die Tage {v['unresolved']} hatten ein day_marker, das im Quelldokument nicht exakt "
                f"gefunden werden konnte — der Marker muss WORTWÖRTLICH (Zeichen für Zeichen) aus dem "
                f"Dokument kopiert sein, nicht paraphrasiert oder übersetzt."
            )
        if v["bad_hotels"]:
            problems.append(
                f"Die Tage {v['bad_hotels']} nennen ein Hotel, das im Quelldokument an keiner Stelle "
                f"vorkommt — vermutlich erfunden statt aus dem Text übernommen."
            )
        if stalled or v["low_coverage"]:
            problems.append(
                f"Ein Großteil des Dokuments wurde offenbar NICHT ausgewertet — vermutlich wurde nach "
                f"einigen Tagen abgebrochen und derselbe day_marker/dasselbe Hotel für alle "
                f"restlichen Tage wiederholt, statt für JEDEN Tag im Dokument nachzusehen, was dort "
                f"wirklich steht. Das gesamte Dokument enthält für jeden Tag eigene, unterscheidbare "
                f"Abschnitte (auch spätere Tage) — diese dürfen nicht übersprungen oder durch den "
                f"letzten erfolgreich erkannten Tag ersetzt werden."
            )
        print(f"[structure] WARNING — issues found: {' | '.join(problems)} — retrying once", flush=True)
        retry_msg = (
            f"DMC offer:\n\n{dmc_content}\n\n"
            f"WICHTIGER HINWEIS zu deiner vorherigen Antwort:\n"
            + "\n".join(f"- {p}" for p in problems)
            + f"\n\nLies das GESAMTE Dokument von Anfang bis Ende noch einmal sorgfältig und "
              f"vollständig durch, bis zum letzten Tag. Für JEDEN Tag: das tatsächliche, im Dokument "
              f"genannte Hotel und Ort für GENAU dieses Datum (niemals ein Hotel/Ort aus dem Dokument "
              f"für ein anderes Datum wiederverwenden oder ein neues erfinden), und ein day_marker, "
              f"der WORTWÖRTLICH und exakt aus dem Dokument kopiert ist und speziell zu DIESEM Tag "
              f"gehört."
        )
        retry_result, retry_unresolved = _run(retry_msg)
        retry_v = _validate_structure(retry_result, dmc_content, retry_unresolved)
        if _issue_score(retry_v) < _issue_score(v):
            result, v = retry_result, retry_v
            stalled = v["stalled_run"] > _STALLED_RUN_THRESHOLD

        if v["missing"]:
            result["_day_count_mismatch"] = {"expected": v["expected"], "actual": v["actual"]}
        if v["unresolved"]:
            result["_hollow_days"] = v["unresolved"]
        if v["bad_hotels"]:
            result["_hallucinated_hotel_days"] = v["bad_hotels"]

        # A stalled_run (the same marker repeated for many consecutive days)
        # is resolved by the merge below — it becomes one legitimate
        # multi-day entry instead of looking like several wrong ones.
        # low_coverage is a separate, still-unresolved signal (most of the
        # document unaccounted for regardless of merging) and still
        # warrants the warning.
        if v["low_coverage"]:
            result["_stalled_extraction"] = True
        if v["missing"] or v["unresolved"] or v["bad_hotels"] or stalled or v["low_coverage"]:
            print(f"[structure] WARNING — issues remain after retry: {v}", flush=True)

    _sanitize_structure(result)
    _backfill_missing_hotels(result.get("days", []))
    _clear_departure_day_hotels(result.get("days", []))
    result["days"] = _merge_undifferentiated_days(result.get("days", []))
    return result


# Pure logistics bullets carry no sightseeing content — nothing to look up a reference for.
_LOGISTICS_BULLETS = {"transfer", "anreise", "abreise", "flug", "anreise / flug", "ankunft"}

# Bullets FRAMED as movement/transit — "Transfer zum Bahnhof Kyoto", "Fahrt mit
# dem Expresszug nach Himeji" — are logistics even though a city or station
# name survives word-filtering. That surviving name is a waypoint being
# passed through, not a sight to visit, but keyword/semantic search can't
# tell the difference: it happily matches "Kyoto" or "Himeji" against ANY
# stored paragraph that mentions the city, including ones about a completely
# different attraction there (e.g. a real Imperial Palace visit paragraph
# getting reused on a day that only passes through Kyoto by train, or a
# Himeji Castle description getting invented for a day that only changes
# trains at Himeji station). Detecting the bullet's own framing — is it about
# going somewhere, or seeing something — catches this regardless of which
# place name follows.
_TRANSIT_PREFIXES = (
    "transfer", "privattransfer", "fahrt mit", "fahrt nach", "weiterfahrt",
    "rückfahrt", "rückreise", "reise nach", "ankunft", "abreise", "abflug",
    "check-in", "check-out", "checkin", "checkout",
)


def _is_transit_bullet(bullet: str) -> bool:
    low = bullet.lower()
    return any(low.startswith(p) for p in _TRANSIT_PREFIXES)


def _map_bullets_to_references(destination: str, location_heading: str, bullets: list) -> list:
    """Resolve grounding for each overview bullet (one atomic sight/activity)
    independently: an exact keyword match (verbatim reuse) first, a semantic
    match (loose style guidance) second, or neither (model uses its own
    knowledge). Retrieving per-bullet instead of over the whole day's activity
    text means a day covering 3 sights gets 3 independently-resolved
    references instead of one fuzzy match for whichever sight's keyword
    happened to score highest against the combined text.

    Source reference paragraphs often cover more than one sight in a single
    flowing paragraph (e.g. one stored text mentions both Kinkakuji and
    Arashiyama together). Resolving bullets independently can then pick that
    same paragraph for two different bullets, or two different paragraphs
    that both happen to mention the same sight — repeating it. To prevent
    that, each bullet's own distinguishing word(s) are tracked, and a
    candidate reference is rejected if it contains another bullet's
    distinguishing word — that reference belongs to the other bullet, so
    this bullet keeps looking (or falls back to the model's own knowledge).
    """
    clean_bullets = [
        b.strip() for b in bullets
        if (b or "").strip() and (b or "").strip().lower() not in _LOGISTICS_BULLETS
    ]
    identity_tokens = {
        b: (set() if _is_transit_bullet(b) else reference_db.distinguishing_tokens(b))
        for b in clean_bullets
    }

    mapped = []
    used_texts = set()
    for bullet in clean_bullets:
        # A bullet with no distinguishing tokens at all (e.g. a generic
        # scene-setter like "Ganztägige flexible Tour mit englischsprachigem
        # Guide") names no actual sight — searching for it risks a semantic
        # match to whatever unrelated city's "guided day tour" text scores
        # closest. Skip retrieval entirely rather than risk that.
        if not identity_tokens[bullet]:
            mapped.append({"bullet": bullet, "exact_text": None, "semantic_text": None})
            continue

        other_tokens = set()
        for other, toks in identity_tokens.items():
            if other != bullet:
                other_tokens |= toks

        candidates = reference_db.find_exact_sightseeing_matches(
            destination, bullet, location_heading=location_heading, limit=4
        )
        exact_text = next(
            (c for c in candidates
             if c not in used_texts and not (reference_db.distinguishing_tokens(c) & other_tokens)),
            None,
        )
        semantic_text = None
        if exact_text:
            used_texts.add(exact_text)
        elif _RAG_ENABLED:
            try:
                results = _rag_retrieve(
                    query=f"{destination} {location_heading} {bullet}",
                    destination=destination,
                    top_k=3,
                )
                semantic_text = next(
                    (r["text"] for r in results
                     if r["text"] not in used_texts
                     and not (reference_db.distinguishing_tokens(r["text"]) & other_tokens)),
                    None,
                )
                if semantic_text:
                    used_texts.add(semantic_text)
            except Exception:
                pass  # Semantic retrieval is a nice-to-have — never block generation on it
        mapped.append({"bullet": bullet, "exact_text": exact_text, "semantic_text": semantic_text})
    return mapped


# A stored exact-match paragraph shorter than this is a genuine BAWA fact,
# but often just one thin sentence carried over from a source document that
# itself under-wrote that sight — locking it in verbatim (as longer matches
# are) would cap every future itinerary mentioning that sight at the same
# one-liner. Below the threshold, the text is handed to the model as a
# factual seed to build on instead of a quote to reproduce untouched.
MIN_VERBATIM_CHARS = 180


def _build_sight_mapping_prompt(mapping: list) -> tuple:
    """Builds the per-sight grounding section of the prompt from
    _map_bullets_to_references' output. Returns (prompt_text, exact_texts) —
    exact_texts is the flat list used afterward to substitute {{SIGHT:n}}
    placeholders with the real verbatim text. Short exact matches (see
    MIN_VERBATIM_CHARS) are inlined as expandable fact seeds instead and
    never enter exact_texts, since there's no placeholder for the model to
    leave alone."""
    if not mapping:
        return "", []
    lines = [
        "SEHENSWÜRDIGKEITEN DIESES TAGES — JEDER Programmpunkt MUSS in body_paragraphs "
        "vorkommen, auch wenn keine Referenz gefunden wurde (dann aus eigenem Wissen "
        "schreiben, siehe SIGHTSEEING-Regeln oben). Lass keinen der folgenden Punkte aus:"
    ]
    exact_texts = []
    for m in mapping:
        if m["exact_text"] and len(m["exact_text"]) >= MIN_VERBATIM_CHARS:
            idx = len(exact_texts)
            exact_texts.append(m["exact_text"])
            lines.append(
                f"- {m['bullet']}: EXAKTE REFERENZ gefunden — MUSS wortwörtlich übernommen werden. "
                f"Füge an passender Stelle in body_paragraphs genau den Platzhalter \"{{{{SIGHT:{idx}}}}}\" "
                f"als eigenständiges Element ein, ohne weiteren Text davor oder danach."
            )
        elif m["exact_text"]:
            lines.append(
                f"- {m['bullet']}: bestätigter BAWA-Fakt, aber zu kurz für einen vollständigen Absatz — "
                f"NICHT wörtlich übernehmen. Nutze ihn als gesicherten Ausgangspunkt und ergänze mit "
                f"weiterem echtem Wissen (siehe SIGHTSEEING-Regeln oben), bis ein vollständiger Absatz "
                f"entsteht:\n  {m['exact_text']}"
            )
        elif m["semantic_text"]:
            lines.append(
                f"- {m['bullet']}: NUR Stil-Vorbild — dieser Text stammt aus einer ANDEREN Reise und "
                f"beschreibt möglicherweise einen ganz anderen Ort, ein anderes Restaurant oder eine "
                f"andere Stadt. Übernimm ausschließlich Ton, Satzrhythmus und Wortwahl. Übernimm "
                f"KEINERLEI Inhalte daraus: keine Eigennamen, Restaurants, Personen, Hotels, "
                f"Auszeichnungen, Orte oder Zahlen. Schreibe den Absatz inhaltlich vollständig aus "
                f"dem oben genannten Programmpunkt und deinem eigenen Wissen:\n  {m['semantic_text']}"
            )
        else:
            lines.append(
                f"- {m['bullet']}: keine Referenz gefunden — schreibe aus eigenem, echtem Wissen "
                f"(siehe SIGHTSEEING-Regeln oben)."
            )
    return "\n".join(lines), exact_texts


def _find_missing_bullets(mapping: list, body_paragraphs: list) -> list:
    """Returns the bullets from `mapping` whose name doesn't appear anywhere
    in the generated prose — a concrete, measurable sign the model dropped a
    programme point instead of writing about it. Used both to log a warning
    and to drive one retry attempt in call_ai_day()."""
    joined = " ".join(body_paragraphs).lower()
    missing = []
    for m in mapping:
        words = [w.lower() for w in re.findall(r"[A-Za-zÀ-ÖØ-öø-ÿ]{4,}", m["bullet"])]
        if words and not any(re.search(rf"\b{re.escape(w)}\b", joined) for w in words):
            missing.append(m["bullet"])
    return missing


_GUIDE_ADJ = r"(?:deutsch|englisch|japanisch|französisch|italienisch|spanisch|chinesisch|koreanisch|portugiesisch)sprachige[nrs]?"


def _normalize_guide_terminology(text: str) -> str:
    """BAWA wants the English loanword "Guide" used consistently instead of
    "Reiseleiter"/"Reiseleiterin"/"Reiseleitung"/"Reiseführer" — including
    in verbatim-reused reference text pulled from the historical corpus,
    which predates this preference and can't be relied on to follow a
    prompt instruction it was never written against. Handled here as a
    deterministic, mechanical substitution applied to every day's final
    text (fresh AI prose and reused reference text alike) rather than
    hoping every generation complies consistently.

    "Guide" is grammatically masculine in German ("der Guide"), so a
    feminine "Reiseleitung" occurrence needs its article corrected too,
    not just the noun swapped ("Ihre Reiseleitung" -> "Ihr Guide", not
    "Ihre Guide"). German's four-case system means the same source phrase
    can require different target grammar depending on sentence role
    (subject vs. object), which plain text substitution can't fully
    resolve — this handles the sentence patterns this app's own prompts
    and reference corpus actually produce; a rarer phrasing falls back to
    a plain word swap that may occasionally leave a mismatched article
    rather than risk a worse mangling.
    """
    def _adj(adj_group: str, ending: str) -> str:
        """Re-ends an adjective (e.g. 'deutschsprachige ') to agree with
        Guide's case/gender instead of the original feminine/plural noun's."""
        if not adj_group:
            return ""
        stem = re.sub(r"(?:e|er|en|em|es)\s*$", "", adj_group.strip())
        return f"{stem}{ending} "

    def _cap_like(article: str, replacement: str) -> str:
        """Preserves sentence-initial capitalization of the matched article."""
        return replacement[0].upper() + replacement[1:] if article[0].isupper() else replacement

    # Reiseexperte/Reiseexpertin: the DMC source's English "travel expert"
    # sometimes gets translated to this instead of "Guide" — same fix, same
    # patterns, different noun stem.
    noun = r"(?:Reiseleit(?:er(?:in)?|ung)|Reiseexpert(?:e|in))\w*"

    # Dative, after a preposition: "mit/von/bei Ihrer/Ihrem ... Reiseleitung/Reiseleiter"
    text = re.sub(
        rf"\b(mit|von|bei)\s+Ihre[rm]\s+({_GUIDE_ADJ}\s+)?(?:lokale[nr]\s+)?{noun}\b",
        lambda m: f"{m.group(1)} Ihrem {_adj(m.group(2), 'en')}Guide",
        text, flags=re.IGNORECASE,
    )
    # Nominative subject: "Ihre/Ihr ... Reiseleitung/Reiseleiter" followed by a 3rd-person verb
    text = re.sub(
        rf"\bIhre?\s+({_GUIDE_ADJ}\s+)?(?:lokale[nr]\s+)?{noun}"
        rf"(?=\s+(?:erwartet|empfängt|begleitet|bringt|wird|holt|führt))",
        lambda m: f"Ihr {_adj(m.group(1), 'er')}Guide",
        text, flags=re.IGNORECASE,
    )
    # Accusative object (everything else with a possessive): "Treffen Sie Ihre/Ihren ... Reiseleitung/Reiseleiter"
    text = re.sub(
        rf"\bIhre[n]?\s+({_GUIDE_ADJ}\s+)?(?:lokale[nr]\s+)?{noun}\b",
        lambda m: f"Ihren {_adj(m.group(1), 'en')}Guide",
        text, flags=re.IGNORECASE,
    )
    # Bare definite/indefinite article: "der/die/ein/eine Reiseleiter/Reiseleitung"
    text = re.sub(
        rf"\b(der|die|ein|eine)\s+({_GUIDE_ADJ}\s+)?(?:lokale[nr]\s+)?{noun}\b",
        lambda m: _cap_like(m.group(1), f"der {_adj(m.group(2), 'e')}Guide"),
        text, flags=re.IGNORECASE,
    )
    # Anything left over — plain word swap.
    text = re.sub(r"\bReiseleiter(?:in)?\b", "Guide", text)
    text = re.sub(r"\bReiseleitung\w*\b", "Guide", text)
    text = re.sub(r"\bReiseführer\b", "Guide", text)
    text = re.sub(r"\bReiseexpert(?:e|in)\w*\b", "Guide", text)
    return re.sub(r"\s{2,}", " ", text)


_OPTIONAL_EXCURSION_RE = re.compile(r"(?im)^\s*option\s*:\s*(.+)$")


def _extract_optional_excursion(activities_raw: str) -> str:
    """Pulls an explicitly-named optional/priced-separately excursion out of
    a free day's source text (e.g. "Option: Half Day Snorkeling..."). Plain
    regex, not an AI call — free days must never get AI-invented sightseeing
    suggestions, but naming an excursion the DMC source itself already
    offers isn't invention, so it's safe to surface deterministically."""
    m = _OPTIONAL_EXCURSION_RE.search(activities_raw or "")
    if not m:
        return ""
    title = m.group(1).strip()
    # Drop a trailing parenthetical (pricing/inclusion notes) — the price
    # card covers cost details separately, keep this to just the name.
    return re.sub(r"\s*\([^)]*\)\s*$", "", title).strip()


def call_ai_day(day: dict, destination: str, day_text_override: str = "") -> dict:
    """Generate prose for a single day. Returns {body_paragraphs, hotel_description}."""
    # Free days get the standard fixed line only — no AI call, no invented
    # suggestions ("Besuchen Sie zum Beispiel Ginza, oder Asakusa..."). Real
    # BAWA documents just say "Genießen Sie die freie Zeit in {city}." and
    # stop there. Skipped if the user typed a manual override in the editor.
    if day.get("is_free_day") and not day_text_override.strip():
        location = (day.get("location_heading") or "").strip()
        line = f"Genießen Sie die freie Zeit in {location}." if location else "Genießen Sie die freie Zeit."
        paragraphs = [line]
        option_title = _extract_optional_excursion(day.get("activities_raw", ""))
        if option_title:
            paragraphs.append(f"Optional (gegen Aufpreis) buchbar: {option_title}.")
        return {"body_paragraphs": paragraphs, "hotel_description": ""}

    activities = day_text_override.strip() or day.get("activities_raw", "")
    # day["hotel"] is explicitly None (not just missing) for hotel-less days
    # (e.g. the departure day) — .get("hotel", {}) doesn't fall back to {} in
    # that case since the key IS present, it just crashed on .get() below.
    hotel = day.get("hotel") or {}
    is_first_night = hotel.get("is_first_night", False)

    # /generate-day receives a day straight from the editor, which can be one
    # the user added by hand and hasn't filled in yet — direct day["..."]
    # indexing raises KeyError on those, so read every field defensively.
    _day_num = day.get("day_number") or ""
    _date = day.get("date") or ""
    _weekday = day.get("weekday") or ""
    _loc = day.get("location_heading") or ""

    day_number_end = day.get("day_number_end")
    if day_number_end:
        day_line = (
            f"Days {_day_num}–{day_number_end}: {_date} – {day.get('date_end') or ''}\n"
            f"NOTE: this is a MULTI-DAY BLOCK, not a single day — the source document describes "
            f"this whole date range as one undifferentiated stretch with no fixed day-by-day "
            f"schedule (e.g. a menu of optional activities guests choose from during their stay). "
            f"Write it as an overview of what's available across these days — NEVER as a single "
            f"day's packed schedule (do not use 'am Vormittag/Nachmittag/Abend' framing implying "
            f"it all happens in one day). Make clear the exact daily order is arranged on-site.\n"
        )
    else:
        day_line = f"Day {_day_num}: {_weekday}, {_date}\n"

    user_msg = (
        f"Destination: {destination}\n"
        f"{day_line}"
        f"Location: {_loc}\n"
        f"Activities (English source): {activities}\n"
        # `or` rather than a .get default — an explicit null would otherwise
        # reach the model as the literal string "None", which it can echo
        # straight back into the client-facing prose.
        f"Hotel: {hotel.get('name') or 'none'} — is_first_night: {is_first_night}\n"
        f"Room: {hotel.get('room_type') or ''}, Meal plan: {hotel.get('meal_plan') or ''}\n"
    )

    ref_hotel_name = hotel.get("name", "") if is_first_night else ""
    hotel_ref = reference_db.find_hotel_reference(destination, ref_hotel_name) if ref_hotel_name else None

    mapping = _map_bullets_to_references(destination, _loc, day.get("overview_bullets") or [])
    sight_block, exact_matches = _build_sight_mapping_prompt(mapping)

    if hotel_ref or sight_block:
        user_msg += (
            "\n\nHAUSSTIL-REFERENZEN — bereits verwendete BAWA-Formulierungen aus echten, "
            "versendeten Reiseverläufen:\n\n"
        )
        if hotel_ref:
            user_msg += (
                f"[Hotel – {ref_hotel_name}]\n{hotel_ref}\n\n"
                "Für den Hotel-Eintrag: nutze ihn als Grundlage für hotel_description, übernimm Fakten "
                "und Tonfall, kürze/passe nur leicht an. Erfinde keine abweichenden Fakten.\n\n"
            )
        if sight_block:
            user_msg += sight_block

    def _generate(msg: str) -> dict:
        for max_tok in (6000, 4000):
            try:
                response = _ai_complete(
                    model=AI_MODEL,
                    temperature=0.45,
                    max_tokens=max_tok,
                    messages=[
                        {"role": "system", "content": DAY_PROSE_PROMPT},
                        {"role": "user",   "content": msg},
                    ],
                )
                break
            except Exception as e:
                if ("413" in str(e) or "rate_limit" in str(e).lower()) and max_tok != 4000:
                    continue
                raise

        raw = response.choices[0].message.content.strip()
        raw = re.sub(r"^```(?:json)?\s*", "", raw)
        raw = re.sub(r"\s*```$", "", raw)
        try:
            parsed = json.loads(raw, strict=False)
        except json.JSONDecodeError:
            parsed = json.loads(_repair_truncated_json(raw), strict=False)

        # Normalize the shape before anything downstream touches it. Valid
        # JSON is not necessarily the shape we asked for — a truncated or
        # off-format response can parse into a list, or into an object whose
        # body_paragraphs is null or holds nulls, and every consumer below
        # (placeholder substitution, bullet checking, guide-terminology
        # rewriting) assumes a list of real strings.
        if not isinstance(parsed, dict):
            parsed = {}
        paragraphs = parsed.get("body_paragraphs")
        parsed["body_paragraphs"] = (
            [p for p in paragraphs if isinstance(p, str) and p.strip()]
            if isinstance(paragraphs, list) else []
        )
        if not isinstance(parsed.get("hotel_description"), str):
            parsed["hotel_description"] = ""

        if exact_matches:
            parsed["body_paragraphs"] = _substitute_sight_placeholders(
                parsed["body_paragraphs"], exact_matches
            )
        return parsed

    day_label = f"Day {day.get('day_number', '?')} ({day.get('date', '')})"
    result = _generate(user_msg)

    # Enforce the "every bullet must appear" instruction, not just assert it:
    # if a programme point is still missing, retry once naming exactly what
    # was dropped. Bounds cost to at most one extra call per day.
    missing = _find_missing_bullets(mapping, result.get("body_paragraphs", []))
    if missing:
        retry_msg = user_msg + (
            "\n\nDeine vorherige Antwort hat folgende Programmpunkte NICHT erwähnt: "
            f"{', '.join(missing)}. Schreibe body_paragraphs erneut und ergänze diese "
            "— alle bereits abgedeckten Punkte bleiben ebenfalls erhalten."
        )
        try:
            retry_result = _generate(retry_msg)
            still_missing = _find_missing_bullets(mapping, retry_result.get("body_paragraphs", []))
            if len(still_missing) < len(missing):
                result = retry_result
                missing = still_missing
        except Exception:
            pass  # keep the first attempt if the retry itself fails

    for bullet in missing:
        print(f"[day-prose] WARNING — {day_label}: bullet {bullet!r} not found in generated text even after retry", flush=True)

    # Deterministic safety net for the "Guide" terminology rule (see
    # _normalize_guide_terminology's docstring) — catches both fresh AI
    # prose and verbatim-reused reference text alike, since it runs after
    # _substitute_sight_placeholders already inserted that text above.
    result["body_paragraphs"] = [_normalize_guide_terminology(p) for p in result.get("body_paragraphs", [])]
    if result.get("hotel_description"):
        result["hotel_description"] = _normalize_guide_terminology(result["hotel_description"])

    return result


_SIGHT_PLACEHOLDER_RE = re.compile(r"\{\{SIGHT:(\d+)\}\}")


def _substitute_sight_placeholders(paragraphs: list, exact_matches: list) -> list:
    """Replace {{SIGHT:n}} placeholders with the exact reference text.

    Matches the model didn't place are dropped, not force-appended: the model
    sees the day's actual activities and is better positioned than a keyword
    match to judge whether a candidate reference is really about this day —
    force-including an unused match risked describing a site the client
    never visits (a keyword false-positive from find_exact_sightseeing_matches).
    """
    final = []
    for p in paragraphs:
        m = _SIGHT_PLACEHOLDER_RE.fullmatch(p.strip())
        if m:
            idx = int(m.group(1))
            if 0 <= idx < len(exact_matches):
                final.append(exact_matches[idx])
                continue
        final.append(p)
    return final


def call_ai(dmc_content: str, day_text: str = "") -> dict:
    if day_text.strip():
        user_msg = (
            "=== DMC STRUCTURE (dates, hotels, sequence) ===\n"
            f"{dmc_content}\n\n"
            "=== ENGLISH DAY TEXT — PRIMARY PROSE SOURCE ===\n"
            "Use this English text as your main source for writing the German body_paragraphs. "
            "Extract every activity, sight, experience, and detail mentioned here. "
            "Do NOT translate it — compose original German prose from these facts.\n\n"
            f"{day_text}"
        )
    else:
        user_msg = f"DMC itinerary content:\n\n{dmc_content}"

    # Free tier rate limits apply. Try progressively smaller max_tokens if rate-limited.
    for max_tok in (8000, 6000, 4000):
        try:
            response = _ai_complete(
                model=AI_MODEL,
                temperature=0.4,
                max_tokens=max_tok,
                messages=[
                    {"role": "system", "content": SYSTEM_PROMPT},
                    {"role": "user",   "content": user_msg},
                ],
            )
            break  # success
        except Exception as e:
            if "413" in str(e) or "rate_limit" in str(e).lower() or "too large" in str(e).lower():
                if max_tok == 4000:
                    raise  # exhausted all retries
                continue   # try next smaller limit
            raise          # non-rate-limit error — propagate immediately

    raw = response.choices[0].message.content.strip()
    # Strip markdown fences if present
    raw = re.sub(r"^```(?:json)?\s*", "", raw)
    raw = re.sub(r"\s*```$", "", raw)
    # If output was truncated mid-JSON, attempt recovery by closing open structures
    try:
        return json.loads(raw, strict=False)
    except json.JSONDecodeError:
        repaired = _repair_truncated_json(raw)
        result = json.loads(repaired, strict=False)
        result["_truncated"] = True
        return result


def _repair_truncated_json(raw: str) -> str:
    """Close any unclosed JSON structures caused by token-limit truncation."""
    # Truncate to last complete top-level value boundary we can find
    # Strategy: count open braces/brackets and close them
    depth_brace   = 0
    depth_bracket = 0
    in_string     = False
    escape_next   = False
    last_safe     = 0

    for i, ch in enumerate(raw):
        if escape_next:
            escape_next = False
            continue
        if ch == '\\' and in_string:
            escape_next = True
            continue
        if ch == '"' and not escape_next:
            in_string = not in_string
            continue
        if in_string:
            continue
        if ch == '{':
            depth_brace += 1
        elif ch == '}':
            depth_brace -= 1
            if depth_brace == 0 and depth_bracket == 0:
                last_safe = i + 1
        elif ch == '[':
            depth_bracket += 1
        elif ch == ']':
            depth_bracket -= 1
            if depth_brace == 0 and depth_bracket == 0:
                last_safe = i + 1

    # Close open arrays then objects
    closing = ']' * depth_bracket + '}' * depth_brace
    if closing:
        # Find a reasonable truncation point: last complete key-value or array item
        # Remove trailing partial value (comma, unfinished string, etc.)
        truncated = raw.rstrip().rstrip(',').rstrip()
        # If we're inside an unclosed string, close it
        open_strings = truncated.count('"') - truncated.count('\\"')
        if open_strings % 2 == 1:
            truncated += '"'
        result = json.loads(truncated + closing, strict=False)
        result["_truncated"] = True
        return json.dumps(result)
    return raw
