"""Periodic backup of critical JSON files (users.json, config.json).

Backs up to LOCAL_DATA_DIR/backups/, keeps the 5 most recent copies per file,
and only saves a file if its content parses as valid JSON (no point archiving
a corrupted file).  All errors are logged as warnings — this is best-effort.
"""
from __future__ import annotations

import json
import logging
import threading
from datetime import datetime
from pathlib import Path
from typing import Optional

log = logging.getLogger(__name__)

_MAX_BACKUPS = 5
_INTERVAL_SECONDS = 600   # 10 minutes


def _backup_once(sources: list[Path], backup_dir: Path) -> None:
    """Copy each source file (if it exists and contains valid JSON) to backup_dir."""
    backup_dir.mkdir(parents=True, exist_ok=True)
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")

    for src in sources:
        try:
            if not src.exists():
                continue
            content = src.read_text(encoding="utf-8")
            json.loads(content)          # validate before archiving
        except Exception as exc:
            log.warning("Backup skipped for %s: %s", src.name, exc)
            continue

        dest = backup_dir / f"{src.stem}_{ts}.json"
        try:
            dest.write_text(content, encoding="utf-8")
            log.debug("Backed up %s → %s", src.name, dest.name)
        except Exception as exc:
            log.warning("Backup write failed for %s: %s", src.name, exc)
            continue

        # Prune oldest backups beyond _MAX_BACKUPS
        existing = sorted(backup_dir.glob(f"{src.stem}_????????_??????.json"))
        for old in existing[:-_MAX_BACKUPS]:
            try:
                old.unlink()
            except Exception:
                pass


def schedule_backup(
    sources: list[Path],
    backup_dir: Path,
    interval: int = _INTERVAL_SECONDS,
) -> None:
    """Run a backup immediately (in a daemon thread), then repeat every *interval* seconds."""

    def _run() -> None:
        _backup_once(sources, backup_dir)
        t = threading.Timer(interval, _run)
        t.daemon = True
        t.start()

    t = threading.Timer(0, _run)
    t.daemon = True
    t.start()
