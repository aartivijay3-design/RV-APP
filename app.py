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
import uuid
from datetime import datetime, timezone
from pathlib import Path
from typing import List, Optional

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
from paths import OUTPUT_DIR, DMCS_PATH, FEEDBACK_PATH

app = FastAPI(title="BAWA Reiseverlauf Generator")
app.mount("/static", StaticFiles(directory="static"), name="static")


@app.exception_handler(Exception)
async def _unhandled_exception_handler(request: Request, exc: Exception):
    """Safety net for any exception a route doesn't already catch and wrap
    in its own HTTPException (this doesn't touch those — FastAPI's built-in
    HTTPException handling stays in effect for them; this only fires for
    everything else). Without this, an unhandled exception falls through to
    FastAPI's default error page, which isn't JSON — so the frontend's
    `err.detail || 'Server-Fehler'` fallback has nothing to read and just
    shows the generic message with zero information about what broke.
    Every route should still catch what it can and raise a specific
    HTTPException with a helpful message — this exists for whatever slips
    through that, not as a replacement for it."""
    print(f"[unhandled] {request.method} {request.url.path}: {exc}", flush=True)
    return JSONResponse(status_code=500, content={"detail": f"Unerwarteter Fehler: {exc}"})


# ── Optional HTTP Basic Auth gate ──────────────────────────────────────────
# Off by default (normal LAN use). Set TUNNEL_AUTH_USER / TUNNEL_AUTH_PASS to
# require a password on every request — needed whenever the app is reachable
# from outside the office LAN, whether via a temporary tunnel or a permanent
# public host like Render. This app handles real client travel data and
# should never sit on a public URL without one.
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
    docx_bytes = inject_into_template(body_xml, destination=itinerary.get("destination") or "")

    # Derive output filename
    last = (itinerary.get("client_name") or "Client").split()[-1]
    dest = (itinerary.get("destination") or "Reise").replace(" ", "_")
    sd = (itinerary.get("start_date") or "").replace(".", "_")
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
        # No truncation here — call_ai_structure applies its own size caps
        # internally (50000 chars) and switches to a chunked extraction
        # path above 9000 chars. This used to cap pasted text at 12000
        # chars, silently cutting off longer itineraries mid-day (a real
        # 22,000-char document was being truncated mid-Day-7) with no
        # error or warning — the file-upload path below never had this
        # limit, so pasted text was quietly less capable than uploading
        # the same content as a file.
        dmc_content = pasted_text.strip()
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
            if b and b.lower() not in ("transfer", "anreise", "abreise", "flug", "anreise / flug")
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
    docx_bytes = inject_into_template(body_xml, destination=itinerary.get("destination") or "")

    # .get(key, default) only falls back when the key is absent — a
    # DMC-extraction failure can leave it present but explicitly null
    # (e.g. no date recognized), which used to crash the whole request
    # with an AttributeError instead of just producing a plainer filename.
    last     = (itinerary.get("client_name") or "Client").split()[-1]
    dest     = (itinerary.get("destination") or "Reise").replace(" ", "_")
    sd       = (itinerary.get("start_date") or "").replace(".", "_")
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


@app.get("/debug/ai-test")
async def debug_ai_test():
    """Temporary diagnostic route — exercises the exact same Gemini call
    path the app uses in production, and returns the precise exception
    instead of a generic 'Connection error.' Remove once the Render
    deployment issue it was added to diagnose is resolved."""
    import traceback
    from ai_client import _ai_complete, AI_MODEL
    try:
        r = _ai_complete(
            model=AI_MODEL, max_tokens=20,
            messages=[{"role": "user", "content": "Reply with just: OK"}],
        )
        return {"status": "ok", "response": r.choices[0].message.content}
    except Exception as e:
        return {
            "status": "error",
            "error_type": type(e).__name__,
            "error": str(e),
            "traceback": traceback.format_exc(),
        }


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


# ═══════════════════════════════════════════════════════════════════
#  FEEDBACK
# ═══════════════════════════════════════════════════════════════════

github_store.pull("feedback.json", FEEDBACK_PATH)


@app.get("/api/feedback")
async def get_feedback():
    """Return all feedback entries, newest first."""
    if not FEEDBACK_PATH.exists():
        return []
    entries = json.loads(FEEDBACK_PATH.read_text(encoding="utf-8"))
    return sorted(entries, key=lambda e: e.get("timestamp", ""), reverse=True)


@app.post("/api/feedback")
async def submit_feedback(
    area: str = Form(default=""),
    kind: str = Form(default=""),
    message: str = Form(...),
    name: str = Form(default=""),
):
    """Append one feedback entry. Every submission is pulled/pushed
    individually (not just cached in memory) so nothing is lost if the
    process restarts on a host with no persistent disk — same reasoning
    as the DMC list and reference database."""
    message = message.strip()
    if not message:
        raise HTTPException(400, "Bitte eine Nachricht eingeben.")

    github_store.pull("feedback.json", FEEDBACK_PATH)
    entries = json.loads(FEEDBACK_PATH.read_text(encoding="utf-8")) if FEEDBACK_PATH.exists() else []
    entries.append({
        "id": uuid.uuid4().hex[:8],
        "timestamp": datetime.now(timezone.utc).isoformat(timespec="seconds"),
        "name": name.strip(),
        "area": area.strip(),
        "kind": kind.strip(),
        "message": message,
    })
    FEEDBACK_PATH.write_text(json.dumps(entries, ensure_ascii=False, indent=2), encoding="utf-8")
    github_store.push("feedback.json", FEEDBACK_PATH, "New feedback submission")
    return {"status": "saved"}



# ── Confirmation route ────────────────────────────────────────────

@app.post("/generate-confirmation")
async def generate_confirmation(
    file: Optional[UploadFile] = File(default=None),
    pasted_text: str = Form(default=""),
    dmc_ids: List[str] = Form(default=[]),
    guide_name: str = Form(default=""),
    guide_phone: str = Form(default=""),
):
    rechnung_text = await _extract_rechnung_text(file, pasted_text)

    # Parse with AI
    try:
        rechnung_data = parse_rechnung_with_ai(rechnung_text)
    except Exception as e:
        raise HTTPException(500, f"Rechnung konnte nicht verarbeitet werden: {e}")

    # Find DMCs — completely optional; if none selected, dmcs stays [].
    # A trip spanning several countries can have one DMC per destination,
    # so this is a list, not a single lookup — each dmc_id is matched
    # against every destination's contact list and tagged with which
    # destination it belongs to (for the "COUNTRY:" label in the doc when
    # there's more than one).
    dmcs: list = []
    try:
        dmcs_all = json.loads(DMCS_PATH.read_text(encoding="utf-8")) if DMCS_PATH.exists() else {}
        for dmc_id in dmc_ids:
            if not dmc_id:
                continue
            for dest_name, dest_dmcs in dmcs_all.items():
                match = next((d for d in dest_dmcs if d.get("id") == dmc_id), None)
                if match:
                    dmcs.append({**match, "_destination": dest_name})
                    break
    except Exception:
        pass  # DMC lookup failure should never block document generation

    # Build document
    try:
        docx_bytes = build_confirmation_docx(rechnung_data, dmcs, guide_name, guide_phone)
    except Exception as e:
        raise HTTPException(500, f"Dokument konnte nicht erstellt werden: {e}")

    # Filename — .get(key, default) only falls back when the key is absent,
    # not when it's present but null/empty, which a parsing failure can do.
    client_names = rechnung_data.get("client_names") or ["Client"]
    client_last  = (client_names[0] if client_names else "Client").split()[-1]
    dest_en      = (rechnung_data.get("destination_en") or "Trip").replace(" ", "_")
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
    except ValueError as e:
        raise HTTPException(422, str(e))
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
