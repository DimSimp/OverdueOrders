from __future__ import annotations

from dataclasses import dataclass, field
from datetime import datetime
import logging

import requests

log = logging.getLogger(__name__)

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
    "ShipAddress",      # Returns all ShipFirstName/LastName/Company/StreetLine1/2/City/State/PostCode/Country/Phone
    "GrandTotal",
    "ShippingTotal",
    "ShippingOption",
    "OrderLine",
    "OrderLine.ProductName",
    "OrderLine.ShortDescription",
    "OrderLine.Name",
    "OrderLine.ThumbURL",
    "OrderLine.UnitPrice",
    "OrderLine.Misc06",
    "OrderLine.ShippingCategory",
]


@dataclass
class NetoLineItem:
    sku: str
    product_name: str
    quantity: int
    unit_price: float
    image_url: str = ""
    postage_type: str = ""      # "Minilope", "Devilope", "Satchel", or "" if not set
    shipping_category: str = "" # e.g. "Books" — used to filter out certain products


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
    # Pricing & shipping
    grand_total: float = 0.0
    shipping_total: float = 0.0
    shipping_method: str = ""
    shipping_type: str = ""  # "Express", "Regular", "Local Pickup", or ""


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

    def get_orders_by_ids(self, order_ids: list[str]) -> list[NetoOrder]:
        """Fetch specific orders by ID. Returns only Pick-status orders."""
        if not order_ids:
            return []
        body = {
            "Filter": {
                "OrderID": order_ids,
                "OutputSelector": OUTPUT_SELECTOR,
            }
        }
        data = self._post(body)
        raw_orders = data.get("Order", [])
        if isinstance(raw_orders, dict):
            raw_orders = [raw_orders]
        orders = []
        for raw in raw_orders:
            order = self._parse_order(raw)
            if order and order.status in UNDISPATCHED_STATUSES:
                orders.append(order)
        return orders

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
        shipping_method: str = "",
        line_item_skus: list | None = None,
        dry_run: bool = True,
    ) -> dict:
        """
        Mark an order as dispatched (or other status).

        Tracking details go on each OrderLine (Neto API requirement).
        ShippingMethod must match an existing shipping service in Neto.
        If no tracking_number is provided, only the OrderStatus is updated.
        """
        from datetime import datetime

        order_update: dict = {
            "OrderID": order_id,
            "OrderStatus": new_status,
        }

        if tracking_number and line_item_skus:
            date_shipped = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            order_lines = []
            for sku in line_item_skus:
                tracking: dict = {
                    "TrackingNumber": tracking_number,
                    "DateShipped": date_shipped,
                }
                if shipping_method:
                    tracking["ShippingMethod"] = shipping_method
                order_lines.append({"SKU": sku, "TrackingDetails": tracking})
            order_update["OrderLine"] = order_lines

        if dry_run:
            print(f"[DRY RUN] Neto UpdateOrder: {order_update}")
            return {"Ack": "Success", "DryRun": True}

        body = {"Order": [order_update]}
        return self._post_action("UpdateOrder", body)

    def get_item_attributes(self, skus: list[str]) -> dict[str, dict]:
        """
        Return {sku: {"shipping_category": str, "postage_type": str, "pick_zone": str}} for the given SKUs.
        - shipping_category: numeric ID ("4" = Books)
        - postage_type: Misc06 value ("Satchel", "Minilope", "Devilope", or "")
        - pick_zone: e.g. "String Room", "Back Area", "Out Front", "Picks", or ""
        Returns empty dict on error.
        """
        if not skus:
            return {}
        body = {
            "Filter": {
                "SKU": skus,
                "OutputSelector": ["SKU", "ShippingCategory", "Misc06", "PickZone"],
            }
        }
        try:
            data = self._post_action("GetItem", body)
        except NetoAPIError:
            return {}
        items = data.get("Item", [])
        if isinstance(items, dict):
            items = [items]
        result = {}
        for item in items:
            sku = str(item.get("SKU", "")).strip()
            if sku:
                result[sku] = {
                    "shipping_category": str(item.get("ShippingCategory", "")).strip(),
                    "postage_type": str(item.get("Misc06", "")).strip(),
                    "pick_zone": str(item.get("PickZone", "")).strip(),
                }
        return result

    def get_item_name(self, sku: str) -> str | None:
        """
        Fetch the Name for a single product SKU via GetItem.
        Returns the product name string, or None if the SKU is not found.
        """
        if not sku:
            return None
        body = {
            "Filter": {
                "SKU": [sku],
                "OutputSelector": ["SKU", "Name"],
            }
        }
        try:
            data = self._post_action("GetItem", body, timeout=10)
        except NetoAPIError:
            return None
        items = data.get("Item", [])
        if isinstance(items, dict):
            items = [items]
        for item in items:
            if str(item.get("SKU", "")).strip().upper() == sku.strip().upper():
                name = str(item.get("Name", "")).strip()
                return name if name else None
        return None

    def get_item_info(self, sku: str) -> dict | None:
        """Return {sku, name, internal_id} for a product, or None if not found."""
        if not sku:
            return None
        body = {
            "Filter": {
                "SKU": [sku],
                "OutputSelector": ["SKU", "Name", "InternalID"],
            }
        }
        try:
            data = self._post_action("GetItem", body, timeout=10)
        except NetoAPIError:
            return None
        items = data.get("Item", [])
        if isinstance(items, dict):
            items = [items]
        for item in items:
            if str(item.get("SKU", "")).strip().upper() == sku.strip().upper():
                return {
                    "sku": item.get("SKU", "").strip(),
                    "name": str(item.get("Name", "")).strip(),
                    "internal_id": str(item.get("InternalID", "")).strip(),
                }
        return None

    def rename_item_sku(self, old_sku: str, new_sku: str, dry_run: bool = False) -> tuple[bool, str]:
        """
        Rename a product's SKU in Neto.
        Two-step: fetch InternalID via GetItem, then UpdateItem using InternalID as identifier.
        Returns (success, message).
        """
        if dry_run:
            return (True, f"[DRY RUN] Would rename '{old_sku}' → '{new_sku}'")
        info = self.get_item_info(old_sku)
        if not info or not info.get("internal_id"):
            return (False, f"Item '{old_sku}' not found in Neto")
        body = {"Item": [{"InternalID": info["internal_id"], "SKU": new_sku}]}
        try:
            result = self._post_action("UpdateItem", body)
        except NetoAPIError as exc:
            return (False, str(exc))
        if result.get("Ack") == "Success":
            return (True, f"Renamed '{old_sku}' → '{new_sku}' in Neto")
        return (False, str(result.get("Messages", "Unknown error")))

    def get_product_images(self, skus: list[str]) -> dict[str, str]:
        """
        Fetch primary image URLs for a list of product SKUs via GetItem.
        Returns {sku: full_image_url} for products that have an image configured.
        Images are returned under the 'Images' list; we use the 'Main' image or first available.
        """
        if not skus:
            return {}
        body = {
            "Filter": {
                "SKU": skus,
                "OutputSelector": ["SKU", "Images"],
            }
        }
        try:
            data = self._post_action("GetItem", body)
        except NetoAPIError:
            return {}
        items = data.get("Item", [])
        if isinstance(items, dict):
            items = [items]
        result = {}
        for item in items:
            sku = str(item.get("SKU", "")).strip()
            images = item.get("Images", [])
            if isinstance(images, dict):
                images = [images]
            if not images:
                continue
            # Prefer "Main" image; fall back to first
            main = next((i for i in images if i.get("Name") == "Main"), images[0])
            url = str(main.get("URL") or "").strip()
            if sku and url:
                result[sku] = url
        return result

    def get_item_dimensions(self, sku: str, require_satchel: bool = False) -> dict | None:
        """
        Fetch shipping dimensions for a product SKU via GetItem.
        Returns {"weight_kg", "length_cm", "width_cm", "height_cm"} or None.
        Neto stores dimensions in metres; we convert to cm (×100).

        require_satchel=True (default False): only return dimensions if Misc06 is
        'Satchel' or 'e-parcel' and height is non-zero (used for auto-fill detection).
        require_satchel=False: return any non-zero dimensions regardless of Misc06.
        """
        body = {
            "Filter": {
                "SKU": sku,
                "OutputSelector": [
                    "SKU", "ShippingHeight", "ShippingLength",
                    "ShippingWidth", "ShippingWeight", "Misc06",
                ],
            }
        }
        try:
            data = self._post_action("GetItem", body, timeout=10)
        except NetoAPIError as exc:
            log.debug("get_item_dimensions SKU=%s API error: %s", sku, exc)
            return None
        items = data.get("Item", [])
        if isinstance(items, dict):
            items = [items]
        if not items:
            log.debug("get_item_dimensions SKU=%s: no item returned", sku)
            return None
        item = items[0]
        log.debug("get_item_dimensions SKU=%s raw: %s", sku, item)
        misc06 = str(item.get("Misc06") or "").strip()
        if misc06 == "e-parcel":
            misc06 = "Satchel"
        shipping_height = str(item.get("ShippingHeight") or "0").strip()
        if require_satchel:
            if misc06 != "Satchel" or shipping_height in ("0.000", "0"):
                log.debug(
                    "get_item_dimensions SKU=%s: skipped (Misc06=%r, height=%s)",
                    sku, misc06, shipping_height,
                )
                return None
        else:
            if shipping_height in ("0.000", "0", ""):
                log.debug("get_item_dimensions SKU=%s: no height set, returning None", sku)
                return None
        try:
            dims = {
                "weight_kg": round(float(item.get("ShippingWeight", 0)), 2),
                "length_cm": round(float(item.get("ShippingLength", 0)) * 100, 2),
                "width_cm": round(float(item.get("ShippingWidth", 0)) * 100, 2),
                "height_cm": round(float(item.get("ShippingHeight", 0)) * 100, 2),
            }
            log.debug("get_item_dimensions SKU=%s: returning %s", sku, dims)
            return dims
        except (ValueError, TypeError) as exc:
            log.debug("get_item_dimensions SKU=%s: parse error %s", sku, exc)
            return None

    def update_item_dimensions(
        self,
        sku: str,
        weight_kg: float,
        length_cm: float,
        width_cm: float,
        height_cm: float,
        dry_run: bool = True,
    ) -> dict:
        """
        Upload shipping dimensions for a product SKU via UpdateItem.
        Stores dimensions in metres (cm / 100) and sets Misc06='Satchel'.
        """
        item_data = {
            "SKU": sku,
            "Misc06": "Satchel",
            "ShippingWeight": weight_kg,
            "ShippingHeight": height_cm / 100,
            "ShippingLength": length_cm / 100,
            "ShippingWidth": width_cm / 100,
        }

        if dry_run:
            print(f"[DRY RUN] Neto UpdateItem dimensions for {sku}: {item_data}")
            return {"Ack": "Success", "DryRun": True}

        body = {"Item": item_data}
        return self._post_action("UpdateItem", body)

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
                "StickyNotes": {"StickyNote": [note]},
            }]
        }
        return self._post_action("UpdateOrder", body)

    def update_item_postage_type(
        self,
        sku: str,
        postage_type: str,
        dry_run: bool = True,
    ) -> dict:
        """
        Set PostageType ([@misc6@]) on a product SKU via UpdateItem.
        After updating, fetches the item back and logs all Misc fields so we
        can confirm the correct field name.
        """
        if dry_run:
            log.debug("[DRY RUN] UpdateItem PostageType SKU=%s → %s", sku, postage_type)
            return {"Ack": "Success", "DryRun": True}

        log.debug("UpdateItem Misc06 SKU=%s → %s", sku, postage_type)
        body = {"Item": [{"SKU": sku, "Misc06": postage_type}]}
        result = self._post_action("UpdateItem", body)
        log.debug("UpdateItem Misc06 response: %s", result)
        return result

    def update_item_shipping_category(
        self,
        sku: str,
        category_id: str,
        dry_run: bool = True,
    ) -> dict:
        """Set ShippingCategory on a product SKU via UpdateItem (e.g. "4" = Books)."""
        if dry_run:
            log.debug("[DRY RUN] UpdateItem ShippingCategory SKU=%s → %s", sku, category_id)
            return {"Ack": "Success", "DryRun": True}
        log.debug("UpdateItem ShippingCategory SKU=%s → %s", sku, category_id)
        body = {"Item": [{"SKU": sku, "ShippingCategory": category_id}]}
        result = self._post_action("UpdateItem", body)
        log.debug("UpdateItem ShippingCategory response: %s", result)
        return result

    def update_item_pick_zone(
        self,
        sku: str,
        zone: str,
        dry_run: bool = True,
    ) -> dict:
        """Set PickZone on a product SKU via UpdateItem."""
        if dry_run:
            log.debug("[DRY RUN] UpdateItem PickZone SKU=%s → %s", sku, zone)
            return {"Ack": "Success", "DryRun": True}
        log.debug("UpdateItem PickZone SKU=%s → %s", sku, zone)
        body = {"Item": [{"SKU": sku, "PickZone": zone}]}
        result = self._post_action("UpdateItem", body)
        log.debug("UpdateItem PickZone response: %s", result)
        return result

    def update_order_postage_type(
        self,
        order_id: str,
        postage_type: str,
        dry_run: bool = True,
    ) -> dict:
        """Set PostageType on an order — kept for backward compat but likely unused."""
        if dry_run:
            log.debug("[DRY RUN] UpdateOrder PostageType %s → %s", order_id, postage_type)
            return {"Ack": "Success", "DryRun": True}
        log.debug("UpdateOrder PostageType (Misc6) %s → %s", order_id, postage_type)
        body = {"Order": [{"OrderID": order_id, "Misc6": postage_type}]}
        result = self._post_action("UpdateOrder", body)
        log.debug("UpdateOrder PostageType response: %s", result)
        return result

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
            postage_type = str(line.get("Misc06") or "").strip()
            shipping_category = str(line.get("ShippingCategory") or "").strip()
            line_items.append(NetoLineItem(
                sku=str(sku).strip(),
                product_name=str(product_name).strip(),
                quantity=qty,
                unit_price=price,
                image_url=image_url,
                postage_type=postage_type,
                shipping_category=shipping_category,
            ))

        # Pricing & shipping
        try:
            grand_total = float(str(raw.get("GrandTotal", 0) or 0))
        except (ValueError, TypeError):
            grand_total = 0.0
        try:
            shipping_total = float(str(raw.get("ShippingTotal", 0) or 0))
        except (ValueError, TypeError):
            shipping_total = 0.0
        shipping_method = str(raw.get("ShippingOption") or "").strip()
        shipping_type = _classify_shipping(shipping_method)

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
            grand_total=grand_total,
            shipping_total=shipping_total,
            shipping_method=shipping_method,
            shipping_type=shipping_type,
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


def _classify_shipping(method: str) -> str:
    """Classify a shipping method string into Express, Regular, or Local Pickup."""
    m = method.lower()
    if not m:
        return ""
    if "pickup" in m or "collect" in m:
        return "Local Pickup"
    if "express" in m or "priority" in m or "overnight" in m or "next day" in m:
        return "Express"
    if any(kw in m for kw in ("standard", "regular", "economy", "parcel", "post", "flat rate", "shipping")):
        return "Regular"
    # Default: if there's a method but we can't classify, call it Regular
    return "Regular"
