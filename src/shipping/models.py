from __future__ import annotations

from dataclasses import dataclass, field
from datetime import datetime, timedelta

import holidays


# ── Default sender address (Scarlett Music) ──────────────────────────────────

DEFAULT_SENDER = {
    "name": "Scarlett Music",
    "company": "Scarlett Music",
    "street1": "286-288 Ballarat Rd",
    "street2": "",
    "city": "Footscray",
    "state": "VIC",
    "postcode": "3011",
    "country": "AU",
    "phone": "0393185751",
    "email": "info@scarlettmusic.com.au",
}

# ── Package presets ──────────────────────────────────────────────────────────

PACKAGE_PRESETS: dict[str, dict] = {
    "Bubble Mailer (≤250g, thin)": {
        "weight_kg": 0.25, "length_cm": 21.0, "width_cm": 21.0, "height_cm": 3.0,
    },
    "Bubble Mailer (≤250g)": {
        "weight_kg": 0.25, "length_cm": 18.0, "width_cm": 23.0, "height_cm": 4.0,
    },
    "Bubble Mailer (≤500g)": {
        "weight_kg": 0.50, "length_cm": 18.0, "width_cm": 23.0, "height_cm": 4.0,
    },
    "Custom": {},
}

# ── Fastway/Aramex satchel size classification ──────────────────────────────

SATCHEL_SIZES = [
    # (max_weight_kg, max_cubic, label)
    (0.3, None, "300gm"),   # special: dims must be ≤3×21×21
    (0.5, 0.5, "A5"),
    (1.0, 1.0, "A4"),
    (3.0, 3.0, "A3"),
    (5.0, 5.0, "A2"),
]


def classify_satchel(weight_kg: float, length_cm: float, width_cm: float, height_cm: float) -> str:
    """Determine Fastway satchel size from package dimensions. Returns '' for regular parcel."""
    cubic = (length_cm * width_cm * height_cm) / 1000  # litres
    # 300gm: weight ≤0.3 AND all dims ≤ 3×21×21
    if weight_kg <= 0.3 and height_cm <= 3 and length_cm <= 21 and width_cm <= 21:
        return "300gm"
    if weight_kg <= 0.5 and cubic <= 0.5:
        return "A5"
    if weight_kg <= 1.0 and cubic <= 1.0:
        return "A4"
    if weight_kg <= 3.0 and cubic <= 3.0:
        return "A3"
    if weight_kg <= 5.0 and cubic <= 5.0:
        return "A2"
    return ""  # regular parcel


# ── Core dataclasses ─────────────────────────────────────────────────────────

@dataclass
class Address:
    name: str
    company: str
    street1: str
    street2: str
    city: str           # suburb
    state: str          # e.g. "VIC"
    postcode: str
    country: str = "AU"
    phone: str = ""
    email: str = ""


@dataclass
class Package:
    weight_kg: float
    length_cm: float
    width_cm: float
    height_cm: float

    @property
    def volume_m3(self) -> float:
        return (self.length_cm * self.width_cm * self.height_cm) / 1_000_000

    @property
    def cubic_weight_kg(self) -> float:
        """Volumetric/cubic weight: (L×W×H) / 4000 (industry standard divisor)."""
        return (self.length_cm * self.width_cm * self.height_cm) / 4000

    @property
    def satchel_size(self) -> str:
        return classify_satchel(self.weight_kg, self.length_cm, self.width_cm, self.height_cm)


@dataclass
class ShipmentRequest:
    order_id: str
    platform: str           # "neto" or "ebay"
    sender: Address
    receiver: Address
    packages: list[Package]
    shipping_type: str      # "Express" or "Standard"
    order_value: float
    dry_run: bool = True


@dataclass
class Quote:
    courier_name: str       # Display name
    courier_code: str       # Internal ID
    service_name: str       # e.g. "Standard Parcel"
    price: float
    estimated_days: str
    raw_response: dict = field(default_factory=dict)
    error: str = ""         # Non-empty = failed quote


@dataclass
class BookingResult:
    courier_name: str
    tracking_number: str
    label_pdf: bytes | None
    booking_reference: str
    error: str = ""


# ── Address extraction helpers ───────────────────────────────────────────────

def address_from_neto_order(order) -> Address:
    """Build an Address from a NetoOrder instance."""
    return Address(
        name=f"{order.ship_first_name} {order.ship_last_name}".strip(),
        company=order.ship_company,
        street1=order.ship_street1,
        street2=order.ship_street2,
        city=order.ship_city,
        state=order.ship_state,
        postcode=order.ship_postcode,
        country=order.ship_country or "AU",
        phone=order.ship_phone,
        email=order.email,
    )


def address_from_ebay_order(order) -> Address:
    """Build an Address from an EbayOrder instance."""
    return Address(
        name=order.ship_name,
        company="",
        street1=order.ship_street1,
        street2=order.ship_street2,
        city=order.ship_city,
        state=order.ship_state,
        postcode=order.ship_postcode,
        country=order.ship_country or "AU",
        phone=order.ship_phone,
        email="",
    )


def sender_from_config(cfg) -> Address:
    """Build sender Address from a SenderConfig or dict."""
    if isinstance(cfg, dict):
        return Address(**cfg)
    return Address(
        name=cfg.name,
        company=cfg.company,
        street1=cfg.street1,
        street2=getattr(cfg, "street2", ""),
        city=cfg.city,
        state=cfg.state,
        postcode=cfg.postcode,
        country=getattr(cfg, "country", "AU"),
        phone=getattr(cfg, "phone", ""),
        email=getattr(cfg, "email", ""),
    )


# ── Business day helper ──────────────────────────────────────────────────────

_AU_HOLIDAYS = holidays.Australia(prov="VIC")


def next_business_day(from_date: datetime | None = None) -> datetime:
    """Return the next business day (Mon-Fri, excluding AU/VIC public holidays)."""
    dt = from_date or datetime.now()
    candidate = dt + timedelta(days=1)
    while candidate.weekday() >= 5 or candidate.date() in _AU_HOLIDAYS:
        candidate += timedelta(days=1)
    return candidate
