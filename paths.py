"""Resolves persistent-data paths for reference_db.py, rag_retrieval.py, and
app.py. Locally these are just project-relative folders. On a host with an
ephemeral filesystem (e.g. Railway, where the regular filesystem resets on
every redeploy), set PERSIST_DIR to a mounted persistent Volume's path so
reference_library.json, dmcs.json, and generated output survive redeploys.

On first boot with an empty volume, seeds it once from the versioned
defaults shipped in the repo — after that, the volume is the source of
truth and the repo copies are never touched again (so in-app edits made via
the Datenbank tab or DMC editor persist across deploys).
"""
import os
import shutil
from pathlib import Path

APP_DIR = Path(__file__).parent
_PERSIST_DIR = os.environ.get("PERSIST_DIR", "").strip()

if _PERSIST_DIR:
    _base = Path(_PERSIST_DIR)
    DATA_DIR = _base / "data"
    DMCS_PATH = _base / "dmcs.json"
    OUTPUT_DIR = _base / "output"

    DATA_DIR.mkdir(parents=True, exist_ok=True)
    for _name in ("reference_library.json", "reference_library.npy"):
        _seed = APP_DIR / "data" / _name
        _target = DATA_DIR / _name
        if _seed.exists() and not _target.exists():
            shutil.copy(_seed, _target)
    _seed_dmcs = APP_DIR / "dmcs.json"
    if _seed_dmcs.exists() and not DMCS_PATH.exists():
        shutil.copy(_seed_dmcs, DMCS_PATH)
else:
    DATA_DIR = APP_DIR / "data"
    DMCS_PATH = APP_DIR / "dmcs.json"
    OUTPUT_DIR = APP_DIR / "output"

OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
