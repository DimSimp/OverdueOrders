from __future__ import annotations

from dataclasses import dataclass
from datetime import datetime

from src.ebay_client import EbayOrder
from src.neto_client import NetoOrder
from src.pdf_parser import InvoiceItem


@dataclass
class MatchedOrder:
    platform: str              # "Neto" or "eBay"
    order_id: str
    customer_name: str
    order_date: datetime | None
    sku: str                   # SKU from the order line item
    description: str           # Product name from the order
    quantity: int
    notes: str                 # Staff/buyer notes from the order
    invoice_sku: str           # Matched invoice SKU (with suffix)
    invoice_description: str
    invoice_qty: int


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
    Filter orders for "on PO" phrase then match order line SKUs
    against invoice item SKUs (case-insensitive exact match).

    Returns:
        matched       — list of MatchedOrder
        unmatched_inv — invoice items that matched no order
    """
    # Build lookup: normalised SKU → InvoiceItem
    invoice_lookup: dict[str, InvoiceItem] = {}
    for item in invoice_items:
        key = item.sku_with_suffix.upper().strip()
        if key:
            invoice_lookup[key] = item

    on_po_neto = filter_on_po(neto_orders, on_po_phrase)
    on_po_ebay = filter_on_po(ebay_orders, on_po_phrase)

    matched: list[MatchedOrder] = []
    matched_invoice_keys: set[str] = set()

    for order in on_po_neto:
        order_date = order.date_paid or order.date_placed
        for line in order.line_items:
            key = line.sku.upper().strip()
            inv = invoice_lookup.get(key)
            if inv:
                matched.append(MatchedOrder(
                    platform=order.sales_channel or "Neto",
                    order_id=order.order_id,
                    customer_name=order.customer_name,
                    order_date=order_date,
                    sku=line.sku,
                    description=line.product_name,
                    quantity=line.quantity,
                    notes=order.notes,
                    invoice_sku=inv.sku_with_suffix,
                    invoice_description=inv.description,
                    invoice_qty=inv.quantity,
                ))
                matched_invoice_keys.add(key)

    for order in on_po_ebay:
        for line in order.line_items:
            key = line.sku.upper().strip()
            inv = invoice_lookup.get(key)
            if inv:
                matched.append(MatchedOrder(
                    platform="eBay",
                    order_id=order.order_id,
                    customer_name=order.buyer_name,
                    order_date=order.creation_date,
                    sku=line.sku,
                    description=line.title,
                    quantity=line.quantity,
                    notes=order.buyer_notes,
                    invoice_sku=inv.sku_with_suffix,
                    invoice_description=inv.description,
                    invoice_qty=inv.quantity,
                ))
                matched_invoice_keys.add(key)

    unmatched_inv = [
        item for item in invoice_items
        if item.sku_with_suffix.upper().strip() not in matched_invoice_keys
    ]

    return matched, unmatched_inv
