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
    invoice_sku: str           # Matched invoice SKU; empty if line doesn't match invoice
    invoice_description: str
    invoice_qty: int           # 0 if line doesn't match invoice
    is_invoice_match: bool     # True if this line item matched an invoice SKU



def filter_on_po(orders: list, phrase: str = "on po") -> list:
    """
    Return only orders whose notes contain the phrase (case-insensitive).
    Works with both NetoOrder and EbayOrder.
    """
    phrase_lower = phrase.lower()
    result = []
    for order in orders:
        notes = ""
        if isinstance(order, NetoOrder):
            notes = order.notes or ""
        elif isinstance(order, EbayOrder):
            notes = order.buyer_notes or ""
        if phrase_lower in notes.lower():
            result.append(order)
    return result


def match_orders_to_invoice(
    invoice_items: list[InvoiceItem],
    neto_orders: list[NetoOrder],
    ebay_orders: list[EbayOrder],
    on_po_phrase: str = "on po",
) -> tuple[list[MatchedOrder], list[InvoiceItem]]:
    """
    Filter orders for the 'on PO' phrase then match order line SKUs against invoice SKUs.

    When an order has at least one matching SKU, ALL of its line items are included
    in the results. Lines that match the invoice have is_invoice_match=True; others False.

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

    on_po_neto = filter_on_po(neto_orders, on_po_phrase)
    # eBay buyer_notes are now populated with PrivateNotes from the Trading API,
    # so the "on po" filter works correctly for eBay orders too.
    on_po_ebay = filter_on_po(ebay_orders, on_po_phrase)

    matched: list[MatchedOrder] = []
    matched_invoice_keys: set[str] = set()

    for order in on_po_neto:
        order_date = order.date_paid or order.date_placed

        # Only process orders that have at least one invoice SKU match
        matching_keys = {
            line.sku.upper().strip()
            for line in order.line_items
            if line.sku.upper().strip() in invoice_lookup
        }
        if not matching_keys:
            continue

        matched_invoice_keys.update(matching_keys)

        # Include ALL line items; mark only the invoice-matching ones with is_invoice_match
        # Notes are order-level for Neto — only shown on the first item
        for idx, line in enumerate(order.line_items):
            key = line.sku.upper().strip()
            inv = invoice_lookup.get(key)
            matched.append(MatchedOrder(
                platform=order.sales_channel or "Neto",
                order_id=order.order_id,
                customer_name=order.customer_name,
                order_date=order_date,
                sku=line.sku,
                description=line.product_name,
                quantity=line.quantity,
                notes=order.notes if idx == 0 else "",
                invoice_sku=inv.sku_with_suffix if inv else "",
                invoice_description=inv.description if inv else "",
                invoice_qty=inv.quantity if inv else 0,
                is_invoice_match=inv is not None,
            ))

    for order in on_po_ebay:
        matching_keys = {
            line.sku.upper().strip()
            for line in order.line_items
            if line.sku.upper().strip() in invoice_lookup
        }
        if not matching_keys:
            continue

        matched_invoice_keys.update(matching_keys)

        # eBay notes are item-level (PrivateNotes per listing)
        for line in order.line_items:
            key = line.sku.upper().strip()
            inv = invoice_lookup.get(key)
            matched.append(MatchedOrder(
                platform="eBay",
                order_id=order.order_id,
                customer_name=order.buyer_name,
                order_date=order.creation_date,
                sku=line.sku,
                description=line.title,
                quantity=line.quantity,
                notes=line.notes,
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
