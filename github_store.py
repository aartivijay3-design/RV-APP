"""Optional GitHub-backed persistence for files that need to survive
restarts on a host with no persistent disk (e.g. Render's free tier, which
resets its local filesystem on every restart/redeploy/sleep-wake cycle).

Set GITHUB_TOKEN (a personal access token with contents:write on the repo)
and GITHUB_REPO ("owner/repo") to enable: pull() fetches the latest version
of a tracked file from the repo before the app reads it, and push() commits
an update back after every write. Both are no-ops (return False) if those
env vars aren't set, so local/LAN use and Railway-volume use are unaffected.

Uses the Git Data API (blobs/trees/commits) rather than the simpler Contents
API — GitHub only inlines a file's content in the Contents API response for
files under 1MB, and reference_library.json (and its .npy embeddings cache)
are bigger than that. Relying on the Contents API's `content` field for
those silently returns an empty string instead of an error, which is exactly
what happened on first deploy: the pull "succeeded" but wrote a 0-byte file.
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
        meta = httpx.get(
            f"{_API}/repos/{_REPO}/contents/{repo_path}",
            headers=_headers(), params={"ref": _BRANCH}, timeout=20,
        )
        if meta.status_code != 200:
            return False
        blob_sha = meta.json()["sha"]

        blob = httpx.get(f"{_API}/repos/{_REPO}/git/blobs/{blob_sha}", headers=_headers(), timeout=30)
        blob.raise_for_status()
        content = base64.b64decode(blob.json()["content"])

        local_path.parent.mkdir(parents=True, exist_ok=True)
        local_path.write_bytes(content)
        print(f"[github_store] pulled {repo_path} ({len(content)} bytes)", flush=True)
        return True
    except Exception as e:
        print(f"[github_store] pull failed for {repo_path}: {e}", flush=True)
        return False


def push(repo_path: str, local_path, message: str) -> bool:
    """Commits local_path's current content to repo_path on GitHub via a
    manual blob + tree + commit + ref-update sequence — the Contents API's
    single-request PUT has the same 1MB inline-content ceiling as its GET,
    so it can silently fail (or reject) the same large files pull() needs
    this workaround for."""
    if not ENABLED:
        return False
    try:
        ref_url = f"{_API}/repos/{_REPO}/git/refs/heads/{_BRANCH}"
        ref = httpx.get(ref_url, headers=_headers(), timeout=20)
        ref.raise_for_status()
        parent_commit_sha = ref.json()["object"]["sha"]

        commit = httpx.get(f"{_API}/repos/{_REPO}/git/commits/{parent_commit_sha}", headers=_headers(), timeout=20)
        commit.raise_for_status()
        base_tree_sha = commit.json()["tree"]["sha"]

        new_blob = httpx.post(
            f"{_API}/repos/{_REPO}/git/blobs", headers=_headers(), timeout=30,
            json={"content": base64.b64encode(local_path.read_bytes()).decode("ascii"), "encoding": "base64"},
        )
        new_blob.raise_for_status()
        blob_sha = new_blob.json()["sha"]

        new_tree = httpx.post(
            f"{_API}/repos/{_REPO}/git/trees", headers=_headers(), timeout=20,
            json={"base_tree": base_tree_sha, "tree": [
                {"path": repo_path, "mode": "100644", "type": "blob", "sha": blob_sha}
            ]},
        )
        new_tree.raise_for_status()
        tree_sha = new_tree.json()["sha"]

        new_commit = httpx.post(
            f"{_API}/repos/{_REPO}/git/commits", headers=_headers(), timeout=20,
            json={"message": message, "tree": tree_sha, "parents": [parent_commit_sha]},
        )
        new_commit.raise_for_status()
        commit_sha = new_commit.json()["sha"]

        update = httpx.patch(ref_url, headers=_headers(), timeout=20, json={"sha": commit_sha})
        update.raise_for_status()

        print(f"[github_store] pushed {repo_path}", flush=True)
        return True
    except Exception as e:
        print(f"[github_store] push failed for {repo_path}: {e}", flush=True)
        return False
