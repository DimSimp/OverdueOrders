from __future__ import annotations

import csv
from pathlib import Path
from typing import Optional

# Column index → Supabase field name (0-based, no header row)
_COLS: dict[int, str] = {
    0: "musipos_account_code",
    1: "surname",
    2: "first_name",
    3: "business",
    4: "address_1",
    5: "address_2",
    6: "city",
    7: "state",
    8: "postcode",
    # 9, 10 = unknown / skip
    # 11 = Title/Salutation / skip
    12: "email",
    13: "phone_1",
}

_ENCODING = "windows-1252"

_PHONE_FIELDS = ("phone_1", "mobile")


def _fix_phone(value: Optional[str]) -> Optional[str]:
    """Restore leading zero stripped by Excel (e.g. '412345678' → '0412345678').

    Skips values that already start with '0' or '+' (international format).
    """
    if not value:
        return value
    if value.startswith("0") or value.startswith("+"):
        return value
    return "0" + value


def _parse_row(row: list[str]) -> Optional[dict]:
    """Convert a raw CSV row (list of strings) to a customer dict, or None to skip."""
    record: dict = {}
    for idx, field in _COLS.items():
        if idx < len(row):
            val = row[idx].strip()
            record[field] = val if val else None
        else:
            record[field] = None

    # Skip rows with no usable identity at all (no name and no business)
    if not record.get("first_name") and not record.get("surname") and not record.get("business"):
        return None

    # Restore leading zero stripped by Excel on phone fields
    for field in _PHONE_FIELDS:
        if field in record:
            record[field] = _fix_phone(record[field])

    record.setdefault("country", "Australia")
    return record


def load_csv(path: str) -> list[dict]:
    """Parse a Musipos customer CSV (no header, Windows-1252) into a list of dicts."""
    rows: list[dict] = []
    with open(path, encoding=_ENCODING, errors="replace", newline="") as fh:
        reader = csv.reader(fh)
        for raw in reader:
            parsed = _parse_row(raw)
            if parsed is not None:
                rows.append(parsed)
    return rows


def get_preview(path: str, n: int = 5) -> list[dict]:
    """Return the first *n* parsed rows for display in the import dialog."""
    rows: list[dict] = []
    with open(path, encoding=_ENCODING, errors="replace", newline="") as fh:
        reader = csv.reader(fh)
        for raw in reader:
            parsed = _parse_row(raw)
            if parsed is not None:
                rows.append(parsed)
                if len(rows) >= n:
                    break
    return rows


def count_rows(path: str) -> int:
    """Count parseable rows in the CSV (skips blank/name-less rows)."""
    count = 0
    with open(path, encoding=_ENCODING, errors="replace", newline="") as fh:
        reader = csv.reader(fh)
        for raw in reader:
            if _parse_row(raw) is not None:
                count += 1
    return count
