from __future__ import annotations

import math
from typing import Optional


DISCOUNT_PROFILE_OPTIONS = ["-", "5%", "10%", "15%", "Teacher", "Staff"]


def normalize_discount_profile(value: str | None) -> str | None:
    value = (value or "").strip()
    if not value or value == "-":
        return None
    if value in DISCOUNT_PROFILE_OPTIONS:
        return value
    return None


def discount_percent_for_profile(profile: str | None) -> Optional[float]:
    profile = normalize_discount_profile(profile)
    if profile == "5%":
        return 5.0
    if profile == "10%":
        return 10.0
    if profile in ("15%", "Teacher"):
        return 15.0
    return None


def staff_price_from_cost(cost_ex_gst: float | None) -> Optional[float]:
    """Return a GST-inclusive sell price that preserves at least a 10% margin."""
    if cost_ex_gst is None or cost_ex_gst <= 0:
        return None
    raw_inc_gst = (cost_ex_gst / 0.9) * 1.1
    return math.ceil(raw_inc_gst * 100) / 100
