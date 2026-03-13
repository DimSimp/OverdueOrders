from __future__ import annotations

import json
import logging
from pathlib import Path

log = logging.getLogger("label_capture")

# Couriers whose labels we want to capture for tuning
CAPTURE_COURIERS = {"dai_post", "aramex", "auspost", "auspost_express", "allied", "bonds"}

LABELS_DIR = Path("data/shipping")
CAPTURED_FILE = LABELS_DIR / "captured.json"


def _load() -> dict:
    if CAPTURED_FILE.exists():
        try:
            return json.loads(CAPTURED_FILE.read_text(encoding="utf-8"))
        except Exception:
            pass
    return {}


def _save(data: dict) -> None:
    LABELS_DIR.mkdir(parents=True, exist_ok=True)
    CAPTURED_FILE.write_text(json.dumps(data, indent=2), encoding="utf-8")


def needs_capture(courier_code: str) -> bool:
    """Return True if we still need to save a test label for this courier.

    Note: the main app no longer calls this — labels are always saved/overwritten.
    Kept for backward compatibility and potential future use.
    """
    if courier_code not in CAPTURE_COURIERS:
        return False
    return not _load().get(courier_code, False)


def save_label(courier_code: str, pdf_bytes: bytes) -> None:
    """Save pdf_bytes to data/shipping/{courier_code}.pdf, always overwriting."""
    LABELS_DIR.mkdir(parents=True, exist_ok=True)
    dest = LABELS_DIR / f"{courier_code}.pdf"
    dest.write_bytes(pdf_bytes)
    log.info("Label saved: %s (%d bytes)", dest, len(pdf_bytes))
    data = _load()
    data[courier_code] = True
    _save(data)
