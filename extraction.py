"""File/text extraction shared by the Reiseverlauf, Confirmation, and
Kalender features: turns an uploaded Excel/Word/PDF/image file (or pasted
text) into plain text for the AI to parse."""
import base64
import io
from pathlib import Path

import docx as python_docx
import openpyxl
import pypdf
try:
    from markitdown import MarkItDown as _MarkItDown
    _markitdown = _MarkItDown()
    _MARKITDOWN_AVAILABLE = True
except Exception:
    _MARKITDOWN_AVAILABLE = False
from fastapi import HTTPException

from ai_client import _ai_complete, AI_MODEL

# ── Excel reader ─────────────────────────────────────────────────────────────

# Rows containing these phrases are financial/admin — skip entirely
_SKIP_KEYWORDS = (
    "cancellation", "refund", "handling fee", "grand total", "net price",
    "consumption tax", "booking charge", "service charge", "sub total",
    "tel:", "<add>", "regency group", "jingumae",
    "once booking", "days before", "% of the grand", "travel supplier",
    "business hours", "note] a new release", "pip is",
)

# Long boilerplate blocks (luggage notices etc.) — skip if line starts with these
_SKIP_STARTSWITH = (
    "on the bullet and express trains",
    "as there is no designated space",
    "in general, luggage",
    "if you drop off your luggage",
    "please note that due to special",
    "luggage with total dimensions",
    "[luggage allowance",
)

def _legacy_read_excel(file_bytes: bytes) -> str:
    """Extract only itinerary-relevant rows from the DMC Excel file.

    Strategy:
    - Only read columns A–J (skip pricing columns K+)
    - Skip financial, admin and boilerplate rows
    - Truncate long single cells (e.g. luggage notices) to 120 chars
    - Hard cap at 7500 chars to stay within free-tier token limits
    """
    wb = openpyxl.load_workbook(io.BytesIO(file_bytes), data_only=True)
    rows = []

    for sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
        rows.append(f"=== Sheet: {sheet_name} ===")

        for row in ws.iter_rows(values_only=True):
            relevant = row[:10]  # columns A–J only
            vals = [
                str(v).strip()
                for v in relevant
                if v is not None and str(v).strip() not in ("", "None")
            ]
            if not vals:
                continue

            line = " | ".join(vals)
            line_lower = line.lower()

            # Skip financial / admin rows
            if any(kw in line_lower for kw in _SKIP_KEYWORDS):
                continue

            # Skip long boilerplate blocks
            if any(line_lower.startswith(kw) for kw in _SKIP_STARTSWITH):
                continue

            # Truncate very long single values (e.g. multi-sentence notices)
            # but keep them so the day structure is preserved
            if len(line) > 400:
                line = line[:400] + "…"

            rows.append(line)

    content = "\n".join(rows)

    # Sanity cap for pathologically large files — Gemini's context window
    # comfortably covers real itineraries far beyond this (a real 16-day,
    # 2-pax Japan itinerary was ~12,000 chars after filtering).
    if len(content) > 50000:
        content = content[:50000] + "\n[...truncated...]"

    return content


# ── Word (.docx) reader ───────────────────────────────────────────────────────

def _legacy_read_word(file_bytes: bytes, max_chars: int = 8500) -> str:
    doc = python_docx.Document(io.BytesIO(file_bytes))
    lines = []
    for p in doc.paragraphs:
        t = p.text.strip()
        if t:
            lines.append(t)
    for table in doc.tables:
        for row in table.rows:
            cells = [c.text.strip() for c in row.cells if c.text.strip()]
            if cells:
                lines.append(" | ".join(cells))
    content = "\n".join(lines)
    if max_chars:
        content = content[:max_chars]
    return content


# ── PDF reader ────────────────────────────────────────────────────────────────

def _legacy_read_pdf(file_bytes: bytes) -> str:
    reader = pypdf.PdfReader(io.BytesIO(file_bytes))
    pages = []
    for page in reader.pages:
        text = page.extract_text()
        if text:
            pages.append(text.strip())
    return "\n".join(pages)


# ── Image reader (JPG / PNG via Gemini vision) ────────────────────────────────

def read_image(file_bytes: bytes, mime_type: str) -> str:
    b64 = base64.b64encode(file_bytes).decode()
    response = _ai_complete(
        model=AI_MODEL,
        max_tokens=4096,
        messages=[{
            "role": "user",
            "content": [
                {
                    "type": "image_url",
                    "image_url": {"url": f"data:{mime_type};base64,{b64}"},
                },
                {
                    "type": "text",
                    "text": (
                        "This is a DMC travel itinerary document image. "
                        "Extract ALL information: client name, all day-by-day activities, "
                        "dates, cities, hotels with room types and meal plans, transfers, "
                        "and any special experiences. Output as structured plain text."
                    ),
                },
            ],
        }],
    )
    return response.choices[0].message.content.strip()


# ── Smart dispatcher ──────────────────────────────────────────────────────────

def _markitdown_extract(file_bytes: bytes, filename: str) -> str:
    """Try MarkItDown extraction; returns plain text or raises on failure."""
    import tempfile, os
    suffix = Path(filename).suffix.lower()
    with tempfile.NamedTemporaryFile(suffix=suffix, delete=False) as tmp:
        tmp.write(file_bytes)
        tmp_path = tmp.name
    try:
        result = _markitdown.convert(tmp_path)
        text = result.text_content or ""
        if not text.strip():
            raise ValueError("MarkItDown returned empty content")
        return text
    finally:
        os.unlink(tmp_path)


def extract_dmc_content(file_bytes: bytes, filename: str) -> str:
    ext = Path(filename).suffix.lower()

    # Images go straight to vision model — MarkItDown can't help here
    if ext in (".jpg", ".jpeg"):
        return read_image(file_bytes, "image/jpeg")
    if ext == ".png":
        return read_image(file_bytes, "image/png")
    if ext in (".doc",):
        raise HTTPException(400, "Legacy .doc files are not supported. Please save as .docx in Word and re-upload.")
    if ext == ".txt":
        return file_bytes.decode("utf-8", errors="replace")

    # Excel: use legacy reader (custom-filtered, skips pricing columns K+)
    # MarkItDown returns unfiltered 70K+ chars including all financial data
    if ext in (".xlsx", ".xls"):
        return _legacy_read_excel(file_bytes)

    # PDF and DOCX: MarkItDown gives cleaner, fuller output — use it first
    if _MARKITDOWN_AVAILABLE and ext in (".docx", ".pdf"):
        try:
            return _markitdown_extract(file_bytes, filename)
        except Exception as e:
            print(f"MarkItDown failed for {filename} ({e}), using legacy reader")

    if ext in (".docx",):
        return _legacy_read_word(file_bytes)
    elif ext == ".pdf":
        return _legacy_read_pdf(file_bytes)
    else:
        raise HTTPException(400, f"Unsupported file type: {ext}. Supported: .xlsx, .docx, .pdf, .txt, .jpg, .png")
