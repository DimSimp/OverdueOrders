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
            version, url = result
            # schedule UI update on main thread via app.after(0, ...)

    t = threading.Thread(target=lambda: _on_result(check_for_update(__version__)), daemon=True)
    t.start()
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

RELEASES_API = f"https://api.github.com/repos/{GITHUB_OWNER}/{GITHUB_REPO}/releases/latest"


def _parse_version(tag: str) -> tuple[int, ...]:
    """Convert 'v1.2.3' or '1.2.3' to (1, 2, 3)."""
    tag = tag.lstrip("v").strip()
    try:
        return tuple(int(x) for x in tag.split("."))
    except ValueError:
        return (0,)


def check_for_update(current_version: str) -> tuple[str, str] | None:
    """
    Return (latest_version, release_url) if a newer version is available,
    or None if already up to date or the check fails.

    This function performs a network request — always call it from a background
    thread, never from the main/UI thread.
    """
    try:
        req = urllib.request.Request(
            RELEASES_API,
            headers={
                "Accept": "application/vnd.github+json",
                "User-Agent": f"ScarlettsOverdueOrders/{current_version}",
            },
        )
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
            # Strip leading 'v' for display
            display = tag_name.lstrip("v")
            return (display, html_url)

        return None

    except URLError as exc:
        log.debug("Update check network error: %s", exc)
        return None
    except Exception as exc:
        log.debug("Update check failed: %s", exc)
        return None
