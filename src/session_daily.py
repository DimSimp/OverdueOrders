from __future__ import annotations

import json
import os
from datetime import datetime
from typing import TYPE_CHECKING

from src.session import (
    _serialize_neto_order,
    _serialize_ebay_order,
    _parse_neto_order,
    _parse_ebay_order,
)

if TYPE_CHECKING:
    from src.neto_client import NetoOrder
    from src.ebay_client import EbayOrder


DAILY_SESSION_DIR = r"\\SERVER\Project Folder\Order-Fulfillment-App\Session\Daily"
DAILY_SESSION_FILE = "CURRENT DAILY SESSION.scar"
DAILY_OVERRIDES_FILE = "DAILY OVERRIDES.json"
DAILY_SESSION_VERSION = 1


def save_daily_session(
    neto_orders: list,
    ebay_orders: list,
    envelope_classifications: dict,
    pick_zones: dict,
    removed_order_ids: set,
) -> None:
    """Overwrite the fixed daily session file on the network share."""
    try:
        os.makedirs(DAILY_SESSION_DIR, exist_ok=True)
        data = {
            "version": DAILY_SESSION_VERSION,
            "timestamp": datetime.now().isoformat(),
            "neto_orders": [_serialize_neto_order(o) for o in neto_orders],
            "ebay_orders": [_serialize_ebay_order(o) for o in ebay_orders],
            "envelope_classifications": envelope_classifications,
            "pick_zones": pick_zones,
            "removed_order_ids": [list(x) for x in removed_order_ids],
        }
        path = os.path.join(DAILY_SESSION_DIR, DAILY_SESSION_FILE)
        with open(path, "w", encoding="utf-8") as f:
            json.dump(data, f, indent=2, ensure_ascii=False, default=str)
    except Exception:
        pass  # network share may be unavailable — silently ignore


def load_daily_session(path: str) -> dict:
    """Read a .scar file and return its raw dict. Caller handles deserialization."""
    with open(path, "r", encoding="utf-8") as f:
        return json.load(f)


def restore_daily_session(data: dict) -> tuple:
    """Deserialize raw session dict.

    Returns (neto_orders, ebay_orders, envelope_classifications, pick_zones, removed_ids).
    """
    neto_orders = [_parse_neto_order(d) for d in data.get("neto_orders", [])]
    ebay_orders = [_parse_ebay_order(d) for d in data.get("ebay_orders", [])]
    envelope_classifications = data.get("envelope_classifications", {})
    pick_zones = data.get("pick_zones", {})
    removed_order_ids = {tuple(x) for x in data.get("removed_order_ids", [])}
    return neto_orders, ebay_orders, envelope_classifications, pick_zones, removed_order_ids


def save_daily_overrides(removed_order_ids: set) -> None:
    """Overwrite the daily overrides JSON with the current removed-order set."""
    try:
        os.makedirs(DAILY_SESSION_DIR, exist_ok=True)
        data = {"removed_order_ids": [list(x) for x in removed_order_ids]}
        path = os.path.join(DAILY_SESSION_DIR, DAILY_OVERRIDES_FILE)
        with open(path, "w", encoding="utf-8") as f:
            json.dump(data, f, indent=2)
    except Exception:
        pass


def load_daily_overrides() -> set:
    """Read daily overrides JSON. Returns set of (platform, order_id) tuples."""
    try:
        path = os.path.join(DAILY_SESSION_DIR, DAILY_OVERRIDES_FILE)
        if not os.path.exists(path):
            return set()
        with open(path, "r", encoding="utf-8") as f:
            data = json.load(f)
        return {tuple(x) for x in data.get("removed_order_ids", [])}
    except Exception:
        return set()
