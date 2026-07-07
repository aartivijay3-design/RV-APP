"""Optional GitHub-backed persistence for files that need to survive
restarts on a host with no persistent disk (e.g. Render's free tier, which
resets its local filesystem on every restart/redeploy/sleep-wake cycle).

Set GITHUB_TOKEN (a personal access token with contents:write on the repo)
and GITHUB_REPO ("owner/repo") to enable: pull() fetches the latest version
of a tracked file from the repo before the app reads it, and push() commits
an update back after every write. Both are no-ops (return False) if those
env vars aren't set, so local/LAN use and Railway-volume use are unaffected.
"""
import base64
import os

import httpx

_TOKEN = os.environ.get("GITHUB_TOKEN", "").strip()
_REPO = os.environ.get("GITHUB_REPO", "").strip()       # "owner/repo"
_BRANCH = os.environ.get("GITHUB_BRANCH", "main").strip()
ENABLED = bool(_TOKEN and _REPO)

_API = "https://api.github.com"


def _headers() -> dict:
    return {"Authorization": f"Bearer {_TOKEN}", "Accept": "application/vnd.github+json"}


def pull(repo_path: str, local_path) -> bool:
    """Downloads repo_path from GitHub into local_path. Returns True if a
    remote copy was found and written, False otherwise (including when
    GitHub sync isn't configured, or the file has never been pushed yet —
    in which case the caller's own local/seed copy is used as-is)."""
    if not ENABLED:
        return False
    try:
        r = httpx.get(
            f"{_API}/repos/{_REPO}/contents/{repo_path}",
            headers=_headers(), params={"ref": _BRANCH}, timeout=20,
        )
        if r.status_code != 200:
            return False
        content = base64.b64decode(r.json()["content"])
        local_path.write_bytes(content)
        print(f"[github_store] pulled {repo_path} ({len(content)} bytes)", flush=True)
        return True
    except Exception as e:
        print(f"[github_store] pull failed for {repo_path}: {e}", flush=True)
        return False


def push(repo_path: str, local_path, message: str) -> bool:
    """Commits local_path's current content to repo_path on GitHub."""
    if not ENABLED:
        return False
    url = f"{_API}/repos/{_REPO}/contents/{repo_path}"
    try:
        r = httpx.get(url, headers=_headers(), params={"ref": _BRANCH}, timeout=20)
        sha = r.json().get("sha") if r.status_code == 200 else None

        payload = {
            "message": message,
            "content": base64.b64encode(local_path.read_bytes()).decode("ascii"),
            "branch": _BRANCH,
        }
        if sha:
            payload["sha"] = sha

        r2 = httpx.put(url, headers=_headers(), json=payload, timeout=30)
        if r2.status_code not in (200, 201):
            print(f"[github_store] push failed for {repo_path}: {r2.status_code} {r2.text[:200]}", flush=True)
            return False
        print(f"[github_store] pushed {repo_path}", flush=True)
        return True
    except Exception as e:
        print(f"[github_store] push failed for {repo_path}: {e}", flush=True)
        return False
