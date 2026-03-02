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


def deduplicate_ebay_orders(
    neto_orders: list[NetoOrder],
    ebay_orders: list[EbayOrder],
) -> tuple[list[NetoOrder], list[EbayOrder]]:
    """
    Match eBay API orders to Neto eBay-channel orders using PurchaseOrderNumber.

    For each Neto order with SalesChannel="eBay":
      - If its PurchaseOrderNumber matches an eBay API order:
          copy Neto sticky notes onto the eBay order (enables 'on PO' detection);
          remove the Neto order from results (eBay order replaces it with eBay order ID).
      - If no eBay API match:
          exclude entirely (order is completed/refunded on eBay, not yet synced in Neto).

    Returns:
        remaining_neto  — Neto orders with ALL eBay-channel orders removed
        enriched_ebay   — eBay orders, matched ones having their notes populated from Neto
    """
    # Build lookup: eBay order ID → EbayOrder
    ebay_by_id: dict[str, EbayOrder] = {o.order_id: o for o in ebay_orders}

    # Collect all Neto eBay-channel order IDs (ALL will be removed from Neto results)
    neto_ebay_order_ids: set[str] = set()

    for neto_order in neto_orders:
        if neto_order.sales_channel.lower() != "ebay":
            continue
        neto_ebay_order_ids.add(neto_order.order_id)

        po = neto_order.purchase_order_number
        if not po:
            continue
        ebay_match = ebay_by_id.get(po)
        if ebay_match and neto_order.notes:
            # Copy Neto sticky notes so the 'on PO' filter can detect them on the eBay order
            if ebay_match.buyer_notes:
                ebay_match.buyer_notes = neto_order.notes + " | " + ebay_match.buyer_notes
            else:
                ebay_match.buyer_notes = neto_order.notes

    # Remove ALL Neto eBay-channel orders:
    #   matched ones are replaced by enriched eBay orders (with eBay IDs)
    #   unmatched ones are excluded (completed/refunded on eBay)
    remaining_neto = [o for o in neto_orders if o.order_id not in neto_ebay_order_ids]
    return remaining_neto, ebay_orders


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
        for line in order.line_items:
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
                notes=order.notes,
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
                notes=order.buyer_notes,
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
