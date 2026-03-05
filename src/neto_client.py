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
    "PurchaseOrderNumber",
    "StickyNotes",
    "InternalOrderNotes",
    "DeliveryInstruction",
    "ShipFirstName",
    "ShipLastName",
    "ShipCompany",
    "ShipStreetLine1",
    "ShipStreetLine2",
    "ShipCity",
    "ShipState",
    "ShipPostCode",
    "ShipCountry",
    "ShipPhone",
    "OrderLine",
    "OrderLine.ProductName",
    "OrderLine.ShortDescription",
    "OrderLine.Name",
    "OrderLine.ThumbURL",
]


@dataclass
class NetoLineItem:
    sku: str
    product_name: str
    quantity: int
    unit_price: float
    image_url: str = ""


@dataclass
class NetoOrder:
    order_id: str
    customer_name: str
    email: str
    date_placed: datetime | None
    date_paid: datetime | None
    status: str
    notes: str             # Concatenated staff/delivery notes
    sales_channel: str
    purchase_order_number: str  # For eBay orders: the eBay order ID (xx-xxxxx-xxxxx)
    line_items: list[NetoLineItem] = field(default_factory=list)
    # Separate note fields for the order detail modal
    sticky_notes: list[dict] = field(default_factory=list)
    internal_notes: str = ""
    delivery_instruction: str = ""
    # Shipping address
    ship_first_name: str = ""
    ship_last_name: str = ""
    ship_company: str = ""
    ship_street1: str = ""
    ship_street2: str = ""
    ship_city: str = ""
    ship_state: str = ""
    ship_postcode: str = ""
    ship_country: str = ""
    ship_phone: str = ""


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
        include_ebay_channel: bool = False,
        progress_callback=None,
    ) -> list[NetoOrder]:
        """
        Fetch all paid, undispatched orders within the given date range.
        By default, eBay-channel orders are excluded (fetched via the eBay API instead).
        Set include_ebay_channel=True to include them (e.g. when eBay direct API is unavailable).
        Handles pagination automatically.
        progress_callback(fetched: int, total: int) is called if provided.
        """
        all_orders = []
        page = 0
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
                if order and (include_ebay_channel or order.sales_channel.lower() != "ebay"):
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
        return self._post_action("GetOrder", body)

    def _post_action(self, action: str, body: dict, timeout: int = 30) -> dict:
        url = f"{self._config.store_url}/do/WS/NetoAPI"
        resp = self._session.post(
            url,
            headers={
                "NETOAPI_ACTION": action,
                "NETOAPI_KEY": self._config.api_key,
                "NETOAPI_USERNAME": self._config.username,
                "Accept": "application/json",
                "Content-Type": "application/json",
            },
            json=body,
            timeout=timeout,
        )
        resp.raise_for_status()
        data = resp.json()
        if data.get("Ack") not in ("Success", "Warning"):
            messages = data.get("Messages", {})
            raise NetoAPIError(f"Neto API error: {messages}")
        return data

    def get_order_status(self, order_id: str) -> str:
        """Lightweight call to check current order status."""
        body = {
            "Filter": {
                "OrderID": order_id,
                "OutputSelector": ["OrderID", "OrderStatus"],
            }
        }
        data = self._post_action("GetOrder", body, timeout=10)
        orders = data.get("Order", [])
        if isinstance(orders, dict):
            orders = [orders]
        if orders:
            return orders[0].get("OrderStatus", "")
        return ""

    def update_order_status(
        self,
        order_id: str,
        new_status: str = "Dispatched",
        tracking_number: str = "",
        carrier: str = "",
        dry_run: bool = True,
    ) -> dict:
        """Mark an order as dispatched (or other status). Returns API response."""
        order_update = {
            "OrderID": order_id,
            "OrderStatus": new_status,
        }
        if tracking_number:
            order_update["ShippingTracking"] = tracking_number
        if carrier:
            order_update["ShippingCarrier"] = carrier

        if dry_run:
            print(f"[DRY RUN] Neto UpdateOrder: {order_update}")
            return {"Ack": "Success", "DryRun": True}

        body = {"Order": [order_update]}
        return self._post_action("UpdateOrder", body)

    def add_sticky_note(
        self,
        order_id: str,
        title: str,
        description: str,
        dry_run: bool = True,
    ) -> dict:
        """Add a sticky note to an order. Returns API response."""
        note = {"Title": title, "Description": description}

        if dry_run:
            print(f"[DRY RUN] Neto AddStickyNote on {order_id}: {note}")
            return {"Ack": "Success", "DryRun": True}

        body = {
            "Order": [{
                "OrderID": order_id,
                "StickyNotes": [note],
            }]
        }
        return self._post_action("UpdateOrder", body)

    def _parse_order(self, raw: dict) -> NetoOrder | None:
        order_id = raw.get("OrderID", "")
        if not order_id:
            return None

        customer_name = raw.get("Username", "") or raw.get("Email", "")

        # Parse separate note fields
        raw_sticky = raw.get("StickyNotes")
        if isinstance(raw_sticky, dict):
            sticky_notes_list = [raw_sticky]
        elif isinstance(raw_sticky, list):
            sticky_notes_list = raw_sticky
        else:
            sticky_notes_list = []
        internal_notes = str(raw.get("InternalOrderNotes") or "").strip()
        delivery_instruction = str(raw.get("DeliveryInstruction") or "").strip()

        # Collect all notes fields and concatenate (backward compat).
        notes_parts = []
        for notes_field in NOTES_FIELDS:
            val = raw.get(notes_field)
            if not val:
                continue
            if isinstance(val, (dict, list)):
                items = [val] if isinstance(val, dict) else val
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
            product_name = (
                line.get("ProductName")
                or line.get("Name")
                or line.get("Title")
                or line.get("ItemDescription")
                or ""
            )
            image_url = str(line.get("ThumbURL") or line.get("DefaultImageURL") or "").strip()
            line_items.append(NetoLineItem(
                sku=str(sku).strip(),
                product_name=str(product_name).strip(),
                quantity=qty,
                unit_price=price,
                image_url=image_url,
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
            purchase_order_number=str(raw.get("PurchaseOrderNumber") or "").strip(),
            line_items=line_items,
            sticky_notes=sticky_notes_list,
            internal_notes=internal_notes,
            delivery_instruction=delivery_instruction,
            ship_first_name=str(raw.get("ShipFirstName") or "").strip(),
            ship_last_name=str(raw.get("ShipLastName") or "").strip(),
            ship_company=str(raw.get("ShipCompany") or "").strip(),
            ship_street1=str(raw.get("ShipStreetLine1") or "").strip(),
            ship_street2=str(raw.get("ShipStreetLine2") or "").strip(),
            ship_city=str(raw.get("ShipCity") or "").strip(),
            ship_state=str(raw.get("ShipState") or "").strip(),
            ship_postcode=str(raw.get("ShipPostCode") or "").strip(),
            ship_country=str(raw.get("ShipCountry") or "").strip(),
            ship_phone=str(raw.get("ShipPhone") or "").strip(),
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
