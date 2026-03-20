from __future__ import annotations

from dataclasses import dataclass
from datetime import datetime

from src.ebay_client import EbayOrder
from src.neto_client import NetoOrder
from src.pdf_parser import InvoiceItem


@dataclass
class MatchedOrder:
    platform: str              # Sales channel: "Website", "eBay", "BigW", etc.
    order_id: str              # eBay order ID for eBay orders; Neto ID otherwise
    customer_name: str
    order_date: datetime | None
    sku: str                   # SKU from the order line item
    description: str           # Product name from the order
    quantity: int
    notes: str                 # Staff/buyer notes from the order
    shipping_type: str = ""    # "Express", "Regular", "Local Pickup", or ""
    invoice_sku: str = ""      # Matched invoice SKU; empty if line doesn't match invoice
    invoice_description: str = ""
    invoice_qty: int = 0       # 0 if line doesn't match invoice
    is_invoice_match: bool = False  # True if this line item matched an invoice SKU



def filter_on_po(orders: list, phrases=None) -> list:
    """
    Return only orders whose notes contain at least one phrase (case-insensitive).
    Works with both NetoOrder and EbayOrder.

    phrases: str, list[str], or None (defaults to ["on po"])
    """
    if phrases is None:
        phrases = ["on po"]
    elif isinstance(phrases, str):
        phrases = [phrases]
    phrases_lower = [p.lower() for p in phrases if p]
    if not phrases_lower:
        return list(orders)
    result = []
    for order in orders:
        notes = ""
        if isinstance(order, NetoOrder):
            notes = order.notes or ""
        elif isinstance(order, EbayOrder):
            notes = order.buyer_notes or ""
        notes_lower = notes.lower()
        if any(p in notes_lower for p in phrases_lower):
            result.append(order)
    return result


def exclude_phrases(orders: list, phrases: list) -> list:
    """
    Return orders whose notes do NOT contain any of the given phrases.
    Used in Daily Operations to exclude 'on PO' orders.
    """
    if not phrases:
        return list(orders)
    matched_ids = {id(o) for o in filter_on_po(orders, phrases)}
    return [o for o in orders if id(o) not in matched_ids]


def _apply_supplier_transform(raw_sku: str, supplier_name: str, suppliers: list) -> str:
    """
    Apply a supplier's character substitutions and suffix to a raw invoice SKU,
    producing the final form used as a Neto SKU (and therefore present in invoice_lookup).
    Returns the raw SKU unchanged if supplier_name is empty or not found.
    """
    if not supplier_name or not suppliers:
        return raw_sku
    for s in suppliers:
        if s.name == supplier_name:
            result = raw_sku
            for old, new_char in s.character_substitutions.items():
                result = result.replace(old, new_char)
            if s.suffix:
                if s.suffix_position == "prepend":
                    result = s.suffix + result
                else:
                    result = result + s.suffix
            return result
    return raw_sku


def match_orders_to_invoice(
    invoice_items: list[InvoiceItem],
    neto_orders: list[NetoOrder],
    ebay_orders: list[EbayOrder],
    on_po_phrase: str = "on po",
    sku_alias_manager=None,
    suppliers=None,
) -> tuple[list[MatchedOrder], list[InvoiceItem]]:
    """
    Filter orders for the 'on PO' phrase then match order line SKUs against invoice SKUs.

    When an order has at least one matching SKU, ALL of its line items are included
    in the results. Lines that match the invoice have is_invoice_match=True; others False.

    sku_alias_manager: optional SkuAliasManager — if provided, unmapped SKUs are looked up
    via the alias file before giving up (supports kit mappings and single-item aliases).
    suppliers: list[SupplierConfig] — required for alias suffix application. Each alias mapping
    stores the raw invoice SKU and the supplier name; the suffix is applied here dynamically.

    Returns:
        matched       — list of MatchedOrder (one entry per order line of every matched order)
        unmatched_inv — invoice items that matched no order
    """
    # Build lookup: normalised SKU → InvoiceItem
    invoice_lookup: dict[str, InvoiceItem] = {}
    for item in invoice_items:
        key = item.sku_with_suffix.upper().strip()
        if key:
            invoice_lookup[key] = item

    # Build alias lookup: Neto order SKU (upper) → InvoiceItem (via alias file)
    # Raw invoice SKUs from the alias are transformed using the mapping's supplier config
    # (character substitutions + suffix) before being looked up in invoice_lookup.
    alias_lookup: dict[str, InvoiceItem] = {}
    if sku_alias_manager:
        for neto_sku, mapping in sku_alias_manager.get_all().items():
            supplier_name = mapping.get("supplier", "")
            for raw_inv_sku in mapping["invoice_skus"]:
                final_sku = _apply_supplier_transform(raw_inv_sku, supplier_name, suppliers or [])
                inv = invoice_lookup.get(final_sku.upper().strip())
                if inv:
                    alias_lookup[neto_sku.upper().strip()] = inv
                    break  # use the first alias that matches the current invoice

    def _resolve(key: str):
        """Return InvoiceItem for a line SKU key, checking alias if no direct match."""
        return invoice_lookup.get(key) or alias_lookup.get(key)

    # TODO: "on PO" filter temporarily disabled — using all awaiting-shipment orders instead.
    # Re-enable these two lines (and remove the two below) once notes are consistent.
    # on_po_neto = filter_on_po(neto_orders, on_po_phrase)
    # on_po_ebay = filter_on_po(ebay_orders, on_po_phrase)
    on_po_neto = neto_orders
    on_po_ebay = ebay_orders

    matched: list[MatchedOrder] = []
    matched_invoice_keys: set[str] = set()

    for order in on_po_neto:
        order_date = order.date_paid or order.date_placed

        # Only process orders that have at least one invoice SKU match (direct or alias)
        if not any(_resolve(line.sku.upper().strip()) for line in order.line_items):
            continue

        # Include ALL line items; mark only the invoice-matching ones with is_invoice_match
        # Notes are order-level for Neto — only shown on the first item
        for idx, line in enumerate(order.line_items):
            key = line.sku.upper().strip()
            inv = _resolve(key)
            if inv:
                matched_invoice_keys.add(inv.sku_with_suffix.upper().strip())
            matched.append(MatchedOrder(
                platform=order.sales_channel or "Neto",
                order_id=order.order_id,
                customer_name=order.customer_name,
                order_date=order_date,
                sku=line.sku,
                description=line.product_name,
                quantity=line.quantity,
                notes=order.notes if idx == 0 else "",
                shipping_type=order.shipping_type,
                invoice_sku=inv.sku_with_suffix if inv else "",
                invoice_description=inv.description if inv else "",
                invoice_qty=inv.quantity if inv else 0,
                is_invoice_match=inv is not None,
            ))

    for order in on_po_ebay:
        if not any(_resolve(line.sku.upper().strip()) for line in order.line_items):
            continue

        # Show order-level notes (checkout + PrivateNotes) on the first line item only
        for idx, line in enumerate(order.line_items):
            key = line.sku.upper().strip()
            inv = _resolve(key)
            if inv:
                matched_invoice_keys.add(inv.sku_with_suffix.upper().strip())
            matched.append(MatchedOrder(
                platform="eBay",
                order_id=order.order_id,
                customer_name=order.buyer_name,
                order_date=order.creation_date,
                sku=line.sku,
                description=line.title,
                quantity=line.quantity,
                notes=order.buyer_notes if idx == 0 else "",
                shipping_type=order.shipping_type,
                invoice_sku=inv.sku_with_suffix if inv else "",
                invoice_description=inv.description if inv else "",
                invoice_qty=inv.quantity if inv else 0,
                is_invoice_match=inv is not None,
            ))

    unmatched_inv = [
        item for item in invoice_items
        if item.sku_with_suffix.upper().strip() not in matched_invoice_keys
    ]

    return matched, unmatched_inv
