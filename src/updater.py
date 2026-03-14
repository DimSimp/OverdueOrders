"""
Background update checker.

Queries the GitHub Releases API for the latest release tag and compares it
against the running version.  Designed to be called in a daemon thread so it
never blocks the UI.

Usage (from the main thread):
    import threading
    from src.updater import check_for_update
    from src.version import __version__

    def _on_result(result):
        if result:
            version, page_url, download_url = result
            # schedule UI update on main thread via app.after(0, ...)

    t = threading.Thread(target=lambda: _on_result(check_for_update(__version__)), daemon=True)
    t.start()

NOTE ON PRIVATE REPOS
---------------------
The GitHub releases API returns 404 for private repositories unless the
request includes an Authorization header with a Personal Access Token (PAT)
that has at least "Contents: Read" permission on the repo.

To use a PAT, set GITHUB_PAT below.  Alternatively, make the repository
public (recommended for an internal update checker with no secrets in the
source tree — API keys live in config.json which should be gitignored).
"""
from __future__ import annotations

import json
import logging
import urllib.request
from urllib.error import URLError

log = logging.getLogger("updater")

# Change these if the repo is ever moved or renamed.
GITHUB_OWNER = "DimSimp"
GITHUB_REPO  = "OverdueOrders"

# Optional: set a GitHub Personal Access Token here if the repo is private.
# Fine-grained PAT with "Contents: Read" on this repo is sufficient.
# Leave empty ("") to use the public API (requires the repo to be public).
GITHUB_PAT = ""

RELEASES_API = f"https://api.github.com/repos/{GITHUB_OWNER}/{GITHUB_REPO}/releases/latest"


def _parse_version(tag: str) -> tuple[int, ...]:
    """Convert 'v1.2.3' or '1.2.3' to (1, 2, 3)."""
    tag = tag.lstrip("v").strip()
    try:
        return tuple(int(x) for x in tag.split("."))
    except ValueError:
        return (0,)


def check_for_update(current_version: str) -> tuple[str, str, str] | None:
    """
    Return (latest_version, release_page_url, download_url) if a newer version
    is available, or None if already up to date or the check fails.

    download_url is the direct URL of the first .zip asset attached to the
    release, or empty string if no zip asset was found (fall back to the page).

    This function performs a network request — always call it from a background
    thread, never from the main/UI thread.
    """
    try:
        headers = {
            "Accept": "application/vnd.github+json",
            "User-Agent": f"ScarlettsOverdueOrders/{current_version}",
        }
        if GITHUB_PAT:
            headers["Authorization"] = f"Bearer {GITHUB_PAT}"

        req = urllib.request.Request(RELEASES_API, headers=headers)
        with urllib.request.urlopen(req, timeout=10) as resp:
            data = json.loads(resp.read().decode("utf-8"))

        tag_name: str = data.get("tag_name", "")
        html_url: str = data.get("html_url", "")

        if not tag_name:
            log.debug("No tag_name in GitHub response — skipping update check")
            return None

        latest  = _parse_version(tag_name)
        current = _parse_version(current_version)

        log.debug("Update check: current=%s  latest=%s", current_version, tag_name)

        if latest > current:
            display = tag_name.lstrip("v")
            # Find the first .zip asset (the release build)
            download_url = ""
            for asset in data.get("assets", []):
                if asset.get("name", "").endswith(".zip"):
                    download_url = asset.get("browser_download_url", "")
                    break
            return (display, html_url, download_url)

        return None

    except URLError as exc:
        log.warning("Update check failed: %s", exc)
        return None
    except Exception as exc:
        log.warning("Update check failed: %s", exc)
        return None
