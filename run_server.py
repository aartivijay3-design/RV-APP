"""Launcher that guarantees the working directory is this folder before starting uvicorn.
The app uses relative paths (assets/, static/, output/) that only resolve correctly
when the process cwd is this directory, regardless of where it was launched from.
"""
import os
import sys

APP_DIR = os.path.dirname(os.path.abspath(__file__))
os.chdir(APP_DIR)
sys.path.insert(0, APP_DIR)

import uvicorn

if __name__ == "__main__":
    # Railway (and most PaaS hosts) assign a dynamic port via $PORT — the app
    # must listen on that, not a hardcoded one, or the platform's router
    # can't reach it. Falls back to 8000 for local/LAN use where $PORT isn't set.
    port = int(os.environ.get("PORT", 8000))
    uvicorn.run("app:app", host="0.0.0.0", port=port)
