"""
BAWA Reiseverlauf Generator — FastAPI Backend
Accepts DMC offers as Excel, Word, PDF, or image (JPG/PNG), or pasted text.
Uses Google Gemini (free tier) for AI generation. Returns a branded Bawa .docx.

Split across modules: ai_client.py (Gemini client), extraction.py (file
readers), reiseverlauf.py (itinerary generation), rechnung.py (Confirmation +
Kalender parsing/docx), reference_db.py + rag_retrieval.py (house-style
reference matching). This file just wires up the FastAPI routes.
"""

import base64
import json
import os
import secrets
from pathlib import Path
from typing import Optional

from fastapi import FastAPI, File, Form, HTTPException, Request, UploadFile
from fastapi.responses import FileResponse, HTMLResponse, JSONResponse, Response
from fastapi.staticfiles import StaticFiles

from extraction import extract_dmc_content
from reiseverlauf import (
    TEMPLATE_PATH, call_ai, build_body_xml, inject_into_template,
    call_ai_structure, call_ai_day, generate_cover_subtitle,
)
from rechnung import (
    parse_rechnung_with_ai, build_confirmation_docx, _extract_rechnung_text,
)
import reference_db
import github_store
from paths import OUTPUT_DIR, DMCS_PATH

app = FastAPI(title="BAWA Reiseverlauf Generator")
app.mount("/static", StaticFiles(directory="static"), name="static")


# ── Optional HTTP Basic Auth gate ──────────────────────────────────────────
# Off by default (normal LAN use). Set TUNNEL_AUTH_USER / TUNNEL_AUTH_PASS
# (e.g. when exposing the app via a temporary public tunnel) to require a
# password on every request — this app handles real client travel data and
# should never sit open on the public internet without one.
_TUNNEL_USER = os.environ.get("TUNNEL_AUTH_USER", "")
_TUNNEL_PASS = os.environ.get("TUNNEL_AUTH_PASS", "")


@app.middleware("http")
async def _basic_auth_gate(request: Request, call_next):
    if not (_TUNNEL_USER and _TUNNEL_PASS):
        return await call_next(request)

    auth = request.headers.get("authorization", "")
    if auth.startswith("Basic "):
        try:
            user, pw = base64.b64decode(auth[6:]).decode("utf-8").split(":", 1)
        except Exception:
            user, pw = "", ""
        if secrets.compare_digest(user, _TUNNEL_USER) and secrets.compare_digest(pw, _TUNNEL_PASS):
            return await call_next(request)

    return Response(status_code=401, headers={"WWW-Authenticate": 'Basic realm="BAWA RV App"'})


def _write_output_docx(out_name: str, docx_bytes: bytes) -> Path:
    """Write a generated docx to OUTPUT_DIR under out_name. Falls back to a
    uniquified filename if the original is locked — e.g. still open in Word
    from a previous generation, which raises PermissionError on Windows.
    The download filename the client sees is unaffected either way.
    """
    out_path = OUTPUT_DIR / out_name
    try:
        out_path.write_bytes(docx_bytes)
        return out_path
    except PermissionError:
        stem, suffix = out_path.stem, out_path.suffix
        for i in range(2, 50):
            alt_path = OUTPUT_DIR / f"{stem} ({i}){suffix}"
            try:
                alt_path.write_bytes(docx_bytes)
                return alt_path
            except PermissionError:
                continue


# ── Routes ───────────────────────────────────────────────────────────────────

@app.get("/", response_class=HTMLResponse)
async def index():
    return (Path("static") / "index.html").read_text(encoding="utf-8")


@app.post("/generate")
async def generate(
    file: Optional[UploadFile] = File(default=None),
    pasted_text: str = Form(default=""),
    client_name: str = Form(default=""),
    day_text: str = Form(default=""),
):
    # ── Source: pasted text ───────────────────────────────────────────────────
    if pasted_text.strip():
        dmc_content = pasted_text.strip()
        # Allow longer pasted content — the model can handle ~15K chars comfortably
        if len(dmc_content) > 15000:
            dmc_content = dmc_content[:15000] + "\n[...truncated...]"

    # ── Source: uploaded file ─────────────────────────────────────────────────
    elif file and file.filename:
        file_bytes = await file.read()
        try:
            dmc_content = extract_dmc_content(file_bytes, file.filename)
        except HTTPException:
            raise
        except Exception as e:
            raise HTTPException(400, f"Konnte die Datei nicht lesen: {e}")
    else:
        raise HTTPException(400, "Bitte eine Datei hochladen oder den DMC-Text einfügen.")

    # ── AI generation ─────────────────────────────────────────────────────────
    try:
        dt = day_text.strip()
        if len(dt) > 12000:
            dt = dt[:12000] + "\n[...truncated...]"
        itinerary = call_ai(dmc_content, day_text=dt)
    except json.JSONDecodeError as e:
        raise HTTPException(500, f"KI hat kein gültiges JSON zurückgegeben: {e}")
    except Exception as e:
        raise HTTPException(500, f"AI generation failed: {e}")

    # Override client name if provided manually
    if client_name.strip():
        itinerary["client_name"] = client_name.strip()

    # Build document
    body_xml = build_body_xml(itinerary)
    docx_bytes = inject_into_template(body_xml, destination=itinerary.get("destination", ""))

    # Derive output filename
    last = itinerary.get("client_name", "Client").split()[-1]
    dest = itinerary.get("destination", "Reise").replace(" ", "_")
    sd = itinerary.get("start_date", "").replace(".", "_")
    out_name = f"{last}_{dest}_{sd}.docx"
    out_path = _write_output_docx(out_name, docx_bytes)

    return FileResponse(
        path=str(out_path),
        filename=out_name,
        media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    )


@app.post("/extract-structure")
async def extract_structure(
    file: Optional[UploadFile] = File(default=None),
    pasted_text: str = Form(default=""),
    client_name: str = Form(default=""),
):
    """Step 1 — extract day structure (no prose). Returns JSON for the editor UI."""
    if pasted_text.strip():
        dmc_content = pasted_text.strip()[:12000]
    elif file and file.filename:
        file_bytes = await file.read()
        try:
            dmc_content = extract_dmc_content(file_bytes, file.filename)
        except HTTPException:
            raise
        except Exception as e:
            raise HTTPException(400, f"Konnte die Datei nicht lesen: {e}")
    else:
        raise HTTPException(400, "Bitte eine Datei hochladen oder Text einfügen.")

    try:
        structure = call_ai_structure(dmc_content)
    except json.JSONDecodeError as e:
        raise HTTPException(500, f"Strukturextraktion fehlgeschlagen: {e}")
    except Exception as e:
        raise HTTPException(500, f"Fehler: {e}")

    if client_name.strip():
        structure["client_name"] = client_name.strip()

    # Generate cover page subtitle using AI (destination + trip highlights)
    destination = structure.get("destination", "")
    if destination and not structure.get("cover_subtitle"):
        highlights = ", ".join(
            b for day in structure.get("days", [])
            for b in day.get("overview_bullets", [])
            if b.lower() not in ("transfer", "anreise", "abreise", "flug", "anreise / flug")
        )
        structure["cover_subtitle"] = generate_cover_subtitle(destination, highlights)

    # Store the raw DMC content so /generate-day can reference it
    structure["_dmc_content"] = dmc_content[:8000]
    return JSONResponse(content=structure)


@app.post("/generate-day")
async def generate_day(request: Request):
    """Step 2 — generate prose for one day. Called per-day from the editor UI."""
    body        = await request.json()
    day         = body.get("day", {})
    destination = body.get("destination", "")
    day_text    = body.get("day_text", "")

    try:
        result = call_ai_day(day, destination, day_text_override=day_text)
    except json.JSONDecodeError as e:
        raise HTTPException(500, f"KI-Ausgabe ungültig: {e}")
    except Exception as e:
        raise HTTPException(500, f"Fehler bei Tagesgenerierung: {e}")

    return JSONResponse(content=result)


@app.post("/build-docx")
async def build_docx(request: Request):
    """Step 3 — build the Word file from the final edited itinerary JSON."""
    body = await request.json()
    itinerary = body.get("itinerary", {})
    ubersicht_mode = body.get("ubersicht_mode", "both")
    if not itinerary:
        raise HTTPException(400, "Kein Reiseverlauf übergeben.")

    body_xml   = build_body_xml(itinerary, ubersicht_mode=ubersicht_mode)
    docx_bytes = inject_into_template(body_xml, destination=itinerary.get("destination", ""))

    last     = itinerary.get("client_name", "Client").split()[-1]
    dest     = itinerary.get("destination", "Reise").replace(" ", "_")
    sd       = itinerary.get("start_date", "").replace(".", "_")
    out_name = f"{last}_{dest}_{sd}.docx"
    out_path = _write_output_docx(out_name, docx_bytes)

    return FileResponse(
        path=str(out_path),
        filename=out_name,
        media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    )


@app.get("/health")
async def health():
    return {"status": "ok", "template": TEMPLATE_PATH.exists()}


# ═══════════════════════════════════════════════════════════════════
#  CONFIRMATION GENERATOR
# ═══════════════════════════════════════════════════════════════════

# See reference_db.py for why this pull-on-import / push-on-save pattern
# exists — same reasoning applies here: without it, DMC edits would be
# silently lost on every restart on a host with no persistent disk.
github_store.pull("dmcs.json", DMCS_PATH)


@app.get("/api/dmcs")
async def get_dmcs():
    """Return the full DMC list grouped by destination."""
    if not DMCS_PATH.exists():
        return {}
    return json.loads(DMCS_PATH.read_text(encoding="utf-8"))


@app.post("/api/dmcs")
async def save_dmcs(payload: dict):
    """Overwrite the DMC list (called from the Manage DMCs UI)."""
    DMCS_PATH.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    github_store.push("dmcs.json", DMCS_PATH, "Update dmcs.json via Manage DMCs UI")
    return {"status": "saved"}



# ── Confirmation route ────────────────────────────────────────────

@app.post("/generate-confirmation")
async def generate_confirmation(
    file: Optional[UploadFile] = File(default=None),
    pasted_text: str = Form(default=""),
    dmc_id: str = Form(default=""),
    destination: str = Form(default=""),
    guide_name: str = Form(default=""),
    guide_phone: str = Form(default=""),
):
    rechnung_text = await _extract_rechnung_text(file, pasted_text)

    # Parse with AI
    try:
        rechnung_data = parse_rechnung_with_ai(rechnung_text)
    except Exception as e:
        raise HTTPException(500, f"Rechnung konnte nicht verarbeitet werden: {e}")

    # Find DMC — completely optional; if not selected, dmc stays {}
    dmc: dict = {}
    try:
        dmcs_all = json.loads(DMCS_PATH.read_text(encoding="utf-8")) if DMCS_PATH.exists() else {}
        for dest_dmcs in dmcs_all.values():
            for d in dest_dmcs:
                if d.get("id") == dmc_id:
                    dmc = d
                    break
        if not dmc and destination:
            dmc = (dmcs_all.get(destination) or [{}])[0]
    except Exception:
        pass  # DMC lookup failure should never block document generation

    # Build document
    try:
        docx_bytes = build_confirmation_docx(rechnung_data, dmc, guide_name, guide_phone)
    except Exception as e:
        raise HTTPException(500, f"Dokument konnte nicht erstellt werden: {e}")

    # Filename
    client_last = rechnung_data.get("client_names", ["Client"])[0].split()[-1]
    dest_en     = rechnung_data.get("destination_en", "Trip").replace(" ", "_")
    out_name    = f"Confirmation_{dest_en}_{client_last}.docx"
    out_path    = _write_output_docx(out_name, docx_bytes)

    return FileResponse(
        path=str(out_path),
        filename=out_name,
        media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    )


# ── Calendar tab route ─────────────────────────────────────────────

@app.post("/parse-calendar-info")
async def parse_calendar_info(
    file: Optional[UploadFile] = File(default=None),
    pasted_text: str = Form(default=""),
):
    """Extract client names, travel dates, hotels and flights from a Confirmation
    or Rechnung document, for building 'Add to Google Calendar' links in the frontend.
    Trips without any hotels (flights-only) are valid here.
    """
    text = await _extract_rechnung_text(file, pasted_text)
    try:
        data = parse_rechnung_with_ai(text, require_hotels=False)
    except Exception as e:
        raise HTTPException(500, f"Dokument konnte nicht verarbeitet werden: {e}")
    return JSONResponse(data)


# ═══════════════════════════════════════════════════════════════════
#  REFERENCE DATABASE (Datenbank tab) — add/edit/delete the house-style
#  hotel & sightseeing paragraphs that call_ai_day() reuses verbatim.
# ═══════════════════════════════════════════════════════════════════

@app.get("/api/reference/destinations")
async def reference_destinations():
    return {"destinations": reference_db.list_destinations()}


@app.get("/api/reference/entries")
async def reference_entries(destination: str = "", type: str = "", query: str = ""):
    return {"entries": reference_db.search_entries(destination, type, query)}


@app.post("/api/reference/parse-document")
async def reference_parse_document(
    file: Optional[UploadFile] = File(default=None),
    pasted_text: str = Form(default=""),
    destination: str = Form(default=""),
):
    """Reuses the same rule-based parser that built the corpus
    (build_reference_db.py) so a staff member can upload a past itinerary,
    review the candidate hotel/sightseeing paragraphs it finds, and add the
    good ones to the live database — with a human check in the loop this
    time, unlike the original one-shot batch build."""
    import io
    from build_reference_db import (
        docx_lines_from_stream, pdf_lines_from_stream, parse_lines, detect_language, normalize_ws,
    )

    if not destination.strip():
        raise HTTPException(400, "Bitte ein Land / Destination angeben.")

    if pasted_text.strip():
        lines = [normalize_ws(l) for l in pasted_text.splitlines() if l.strip()]
        source_name = "Manuell eingefügter Text"
    elif file and file.filename:
        data = await file.read()
        ext = file.filename.rsplit(".", 1)[-1].lower()
        try:
            if ext == "docx":
                lines = list(docx_lines_from_stream(io.BytesIO(data)))
            elif ext == "pdf":
                lines = list(pdf_lines_from_stream(io.BytesIO(data)))
            elif ext == "txt":
                lines = [normalize_ws(l) for l in data.decode("utf-8", errors="ignore").splitlines() if l.strip()]
            else:
                raise HTTPException(400, "Nur .docx, .pdf oder .txt werden unterstützt.")
        except HTTPException:
            raise
        except Exception as e:
            raise HTTPException(400, f"Konnte die Datei nicht lesen: {e}")
        source_name = f"Hochgeladen: {file.filename}"
    else:
        raise HTTPException(400, "Bitte eine Datei hochladen oder Text einfügen.")

    if not lines:
        raise HTTPException(400, "Kein verwertbarer Text in der Datei gefunden.")

    language = detect_language(lines)
    candidates = parse_lines(lines, destination.strip(), source_name)
    for c in candidates:
        c["language"] = language
    return {"candidates": candidates}


@app.post("/api/reference/entries/bulk")
async def reference_add_bulk(payload: dict):
    added = reference_db.add_entries_bulk(payload.get("entries", []))
    return {"added": added}


@app.post("/api/reference/entries")
async def reference_add(payload: dict):
    try:
        entry = reference_db.add_entry(
            destination=payload.get("destination", ""),
            type_=payload.get("type", ""),
            heading=payload.get("heading", ""),
            text=payload.get("text", ""),
        )
    except ValueError as e:
        raise HTTPException(400, str(e))
    return entry


@app.put("/api/reference/entries/{idx}")
async def reference_update(idx: int, payload: dict):
    try:
        entry = reference_db.update_entry(
            idx,
            destination=payload.get("destination", ""),
            type_=payload.get("type", ""),
            heading=payload.get("heading", ""),
            text=payload.get("text", ""),
        )
    except ValueError as e:
        raise HTTPException(400, str(e))
    except IndexError as e:
        raise HTTPException(404, str(e))
    return entry


@app.delete("/api/reference/entries/{idx}")
async def reference_delete(idx: int):
    try:
        reference_db.delete_entry(idx)
    except IndexError as e:
        raise HTTPException(404, str(e))
    return {"status": "deleted"}
