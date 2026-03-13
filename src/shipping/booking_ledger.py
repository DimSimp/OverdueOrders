"""Daily booking ledger — persists today's bookings to a JSON file.

Each day gets its own file (YYYY-MM-DD.json) inside a configurable directory.
Records are appended after each successful courier booking and can be listed
or removed (on cancellation).
"""
from __future__ import annotations

import json
import logging
import os
from datetime import date, datetime
from pathlib import Path
from typing import Optional

log = logging.getLogger(__name__)


def _today_file(directory: str | Path) -> Path:
    return Path(directory) / f"{date.today().isoformat()}.json"


def _read(path: Path) -> list[dict]:
    if not path.exists():
        return []
    try:
        with open(path, "r", encoding="utf-8") as f:
            data = json.load(f)
        return data if isinstance(data, list) else []
    except (json.JSONDecodeError, OSError) as exc:
        log.warning("Failed to read booking ledger %s: %s", path, exc)
        return []


def _write(path: Path, records: list[dict]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with open(path, "w", encoding="utf-8") as f:
        json.dump(records, f, indent=2, ensure_ascii=False)


def add_booking(
    directory: str | Path,
    courier_code: str,
    courier_name: str,
    tracking_number: str,
    order_id: str = "",
    recipient: str = "",
    extras: dict | None = None,
) -> None:
    """Append a booking record to today's ledger file.

    Args:
        extras: Optional courier-specific data (e.g. shipment_id for AusPost).
    """
    path = _today_file(directory)
    records = _read(path)
    record = {
        "courier_code": courier_code,
        "courier_name": courier_name,
        "tracking_number": tracking_number,
        "order_id": order_id,
        "recipient": recipient,
        "booked_at": datetime.now().isoformat(timespec="seconds"),
        "cancelled": False,
    }
    if extras:
        record["extras"] = extras
    records.append(record)
    _write(path, records)
    log.info("Ledger: added %s booking %s", courier_name, tracking_number)


def get_todays_bookings(directory: str | Path) -> list[dict]:
    """Return all non-cancelled bookings from today's ledger."""
    records = _read(_today_file(directory))
    return [r for r in records if not r.get("cancelled")]


def get_all_bookings(directory: str | Path, days: int = 60) -> list[dict]:
    """Return all non-cancelled bookings from the last *days* days, newest first.

    Each record is augmented with a ``"date"`` key (ISO date string from the
    filename) so callers can display when the booking was made.
    """
    from datetime import timedelta
    dir_path = Path(directory)
    if not dir_path.exists():
        return []

    cutoff = date.today() - timedelta(days=days)
    results: list[dict] = []

    for json_file in sorted(dir_path.glob("????-??-??.json"), reverse=True):
        try:
            file_date = date.fromisoformat(json_file.stem)
        except ValueError:
            continue
        if file_date < cutoff:
            break
        for record in _read(json_file):
            if not record.get("cancelled"):
                results.append({**record, "date": json_file.stem})

    return results


def mark_cancelled(
    directory: str | Path,
    tracking_number: str,
    booking_date: str | None = None,
) -> bool:
    """Mark a booking as cancelled in the ledger. Returns True if found.

    If *booking_date* (ISO date string, e.g. ``"2026-03-11"``) is supplied the
    matching day-file is updated directly; otherwise today's file is searched.
    """
    dir_path = Path(directory)
    if booking_date:
        candidates = [dir_path / f"{booking_date}.json"]
    else:
        candidates = [_today_file(directory)]

    for path in candidates:
        records = _read(path)
        for r in records:
            if r.get("tracking_number") == tracking_number and not r.get("cancelled"):
                r["cancelled"] = True
                _write(path, records)
                log.info("Ledger: marked %s as cancelled in %s", tracking_number, path.name)
                return True
    return False
