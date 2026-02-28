from __future__ import annotations

from dataclasses import dataclass, field
from datetime import datetime

import requests

from src.config import NetoConfig

# Only "Pick" status orders are relevant — staff move orders here when items are on PO
UNDISPATCHED_STATUSES = ["Pick"]

# Confirmed field names for this Neto instance.
# StickyNotes = prominent sticky notes; InternalOrderNotes = regular staff notes.
NOTES_FIELDS = ["StickyNotes", "InternalOrderNotes", "DeliveryInstruction"]

OUTPUT_SELECTOR = [
    "OrderID",
    "Username",
    "Email",
    "DatePlaced",
    "DatePaid",
    "OrderStatus",
    "SalesChannel",
    "StickyNotes",
    "InternalOrderNotes",
    "DeliveryInstruction",
    "OrderLine",
]


@dataclass
class NetoLineItem:
    sku: str
    product_name: str
    quantity: int
    unit_price: float


@dataclass
class NetoOrder:
    order_id: str
    customer_name: str
    email: str
    date_placed: datetime | None
    date_paid: datetime | None
    status: str
    notes: str        # Concatenated staff/delivery notes
    sales_channel: str
    line_items: list[NetoLineItem] = field(default_factory=list)


class NetoAPIError(Exception):
    pass


class NetoClient:
    def __init__(self, config: NetoConfig):
        self._config = config
        self._session = requests.Session()

    def get_overdue_orders(
        self,
        date_from: datetime,
        date_to: datetime,
        progress_callback=None,
    ) -> list[NetoOrder]:
        """
        Fetch all paid, undispatched orders within the given date range.
        Handles pagination automatically.
        progress_callback(fetched: int, total: int) is called if provided.
        """
        all_orders = []
        page = 1
        limit = 200
        total = None

        while True:
            body = self._build_filter(date_from, date_to, page, limit)
            data = self._post(body)

            raw_orders = data.get("Order", [])
            if isinstance(raw_orders, dict):
                raw_orders = [raw_orders]

            # Neto returns total count differently across versions
            if total is None:
                total = int(data.get("CurrentPage", {}).get("TotalResults", len(raw_orders)))
                if total == 0 and raw_orders:
                    total = len(raw_orders)

            for raw in raw_orders:
                order = self._parse_order(raw)
                if order:
                    all_orders.append(order)

            if progress_callback:
                progress_callback(len(all_orders), total or len(all_orders))

            if len(raw_orders) < limit:
                break
            page += 1

        return all_orders

    def _build_filter(
        self, date_from: datetime, date_to: datetime, page: int, limit: int
    ) -> dict:
        return {
            "Filter": {
                "OrderStatus": UNDISPATCHED_STATUSES,
                "PaymentStatus": "FullyPaid",
                "DatePaidFrom": date_from.strftime("%Y-%m-%d 00:00:00"),
                "DatePaidTo": date_to.strftime("%Y-%m-%d 23:59:59"),
                "OutputSelector": OUTPUT_SELECTOR,
                "Page": page,
                "Limit": limit,
            }
        }

    def _post(self, body: dict) -> dict:
        url = f"{self._config.store_url}/do/WS/NetoAPI"
        resp = self._session.post(
            url,
            headers={
                "NETOAPI_ACTION": "GetOrder",
                "NETOAPI_KEY": self._config.api_key,
                "NETOAPI_USERNAME": self._config.username,
                "Accept": "application/json",
                "Content-Type": "application/json",
            },
            json=body,
            timeout=30,
        )
        resp.raise_for_status()
        data = resp.json()
        if data.get("Ack") not in ("Success", "Warning"):
            messages = data.get("Messages", {})
            raise NetoAPIError(f"Neto API error: {messages}")
        return data

    def _parse_order(self, raw: dict) -> NetoOrder | None:
        order_id = raw.get("OrderID", "")
        if not order_id:
            return None

        customer_name = raw.get("Username", "") or raw.get("Email", "")

        # Collect all notes fields and concatenate.
        # StickyNotes: single dict or list of dicts {StickyNoteID, Title, Description}.
        # InternalOrderNotes and DeliveryInstruction are plain strings.
        notes_parts = []
        for notes_field in NOTES_FIELDS:
            val = raw.get(notes_field)
            if not val:
                continue
            if isinstance(val, (dict, list)):
                items = [val] if isinstance(val, dict) else val
                # Use Description as the note text; fall back to Title if empty
                text = " | ".join(
                    str(n.get("Description") or n.get("Title") or "").strip()
                    for n in items
                ).strip()
            else:
                text = str(val).strip()
            if text:
                notes_parts.append(text)
        notes = " | ".join(notes_parts)

        line_items = []
        raw_lines = raw.get("OrderLine", [])
        if isinstance(raw_lines, dict):
            raw_lines = [raw_lines]
        for line in raw_lines:
            sku = line.get("SKU", "") or line.get("ProductSKU", "")
            qty_raw = line.get("Quantity", line.get("QuantityOrdered", 1))
            try:
                qty = int(float(str(qty_raw)))
            except (ValueError, TypeError):
                qty = 1
            try:
                price = float(str(line.get("UnitPrice", 0)))
            except (ValueError, TypeError):
                price = 0.0
            line_items.append(NetoLineItem(
                sku=str(sku).strip(),
                product_name=str(line.get("ProductName", "")).strip(),
                quantity=qty,
                unit_price=price,
            ))

        return NetoOrder(
            order_id=str(order_id),
            customer_name=customer_name,
            email=raw.get("Email", ""),
            date_placed=_parse_date(raw.get("DatePlaced")),
            date_paid=_parse_date(raw.get("DatePaid")),
            status=raw.get("OrderStatus", ""),
            notes=notes,
            sales_channel=raw.get("SalesChannel", ""),
            line_items=line_items,
        )


def _parse_date(raw: str | None) -> datetime | None:
    if not raw:
        return None
    for fmt in ("%Y-%m-%d %H:%M:%S", "%Y-%m-%dT%H:%M:%S", "%Y-%m-%d"):
        try:
            return datetime.strptime(raw, fmt)
        except (ValueError, TypeError):
            continue
    return None
