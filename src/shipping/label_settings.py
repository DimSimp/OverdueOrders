from __future__ import annotations

import json
import sys
from pathlib import Path

# Resolve an absolute path so the settings file is always found next to the exe
# (when frozen) or at the project root (when running from source).
if getattr(sys, "frozen", False):
    _BASE = Path(sys.executable).parent
else:
    _BASE = Path(__file__).parent.parent.parent

SETTINGS_FILE = _BASE / "data" / "shipping" / "label_settings.json"

DEFAULTS: dict = {
    "scale": 0.92,
    "split_ratio": 0.57,
    "label_length_mm": 0.0,
    "no_split": False,
    "rotate_cw": False,
    "skip_initial_rotate": False,
}

# Per-courier overrides applied between global DEFAULTS and any saved settings.
# These ensure correct behaviour even before the user has tuned and saved via label_tuner.
COURIER_DEFAULTS: dict = {
    "allied": {"skip_initial_rotate": True, "rotate_cw": True},
    "auspost_express": {"no_split": True},
}


def load(courier_code: str = "") -> dict:
    """Load print settings for a courier, falling back to defaults.

    Merge order: DEFAULTS → COURIER_DEFAULTS (if any) → saved JSON settings.
    """
    base = dict(DEFAULTS)
    if courier_code and courier_code in COURIER_DEFAULTS:
        base.update(COURIER_DEFAULTS[courier_code])

    if SETTINGS_FILE.exists():
        try:
            all_settings = json.loads(SETTINGS_FILE.read_text(encoding="utf-8"))
            if courier_code and courier_code in all_settings:
                base.update(all_settings[courier_code])
        except Exception:
            pass
    return base


def save(
    courier_code: str,
    scale: float,
    split_ratio: float,
    label_length_mm: float = 0.0,
    no_split: bool = False,
    rotate_cw: bool = False,
    skip_initial_rotate: bool = False,
) -> None:
    """Persist settings for one courier without disturbing others."""
    SETTINGS_FILE.parent.mkdir(parents=True, exist_ok=True)
    existing: dict = {}
    if SETTINGS_FILE.exists():
        try:
            existing = json.loads(SETTINGS_FILE.read_text(encoding="utf-8"))
        except Exception:
            pass
    existing[courier_code] = {
        "scale": round(scale, 4),
        "split_ratio": round(split_ratio, 4),
        "label_length_mm": round(label_length_mm, 1),
        "no_split": no_split,
        "rotate_cw": rotate_cw,
        "skip_initial_rotate": skip_initial_rotate,
    }
    SETTINGS_FILE.write_text(json.dumps(existing, indent=2), encoding="utf-8")
