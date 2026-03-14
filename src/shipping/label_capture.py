from __future__ import annotations

import json
import logging
import shutil
import sys
from datetime import date, timedelta
from pathlib import Path

log = logging.getLogger("label_capture")

# Couriers whose labels we want to capture for tuning
CAPTURE_COURIERS = {"dai_post", "aramex", "auspost", "auspost_express", "allied", "bonds"}

# Resolve an absolute base directory so labels land in a consistent, findable
# location regardless of the working directory or whether we are running from
# source vs a packaged exe.
if getattr(sys, "frozen", False):
    # Running as packaged exe — save next to the exe
    _BASE = Path(sys.executable).parent
else:
    # Running from source — save at the project root (parent of src/shipping/)
    _BASE = Path(__file__).parent.parent.parent

LABELS_DIR    = _BASE / "data" / "shipping"
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


def save_order_label(bookings_dir: str | Path, order_id: str, pdf_bytes: bytes) -> None:
    """Save the full label PDF to {bookings_dir}/Labels/{today}/{order_id}.pdf.

    Also purges label sub-folders that are older than 7 days to prevent
    unbounded accumulation on the network share.
    """
    labels_root = Path(bookings_dir) / "Labels"
    today_dir = labels_root / date.today().isoformat()
    today_dir.mkdir(parents=True, exist_ok=True)

    dest = today_dir / f"{order_id}.pdf"
    dest.write_bytes(pdf_bytes)
    log.info("Order label saved: %s (%d bytes)", dest, len(pdf_bytes))

    _cleanup_old_labels(labels_root, days=7)


def _cleanup_old_labels(labels_root: Path, days: int = 7) -> None:
    """Delete dated label sub-folders older than *days* days."""
    if not labels_root.exists():
        return
    cutoff = date.today() - timedelta(days=days)
    for subdir in labels_root.iterdir():
        if not subdir.is_dir():
            continue
        try:
            folder_date = date.fromisoformat(subdir.name)
        except ValueError:
            continue
        if folder_date < cutoff:
            try:
                shutil.rmtree(subdir)
                log.info("Cleaned up old label folder: %s", subdir)
            except Exception as exc:
                log.warning("Failed to clean up %s: %s", subdir, exc)
