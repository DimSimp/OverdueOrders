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


_CMC_SUFFIX   = "CMC"   # Supplier suffix that triggers the P0-prefix special case
_CMC_P0_PREFIX = "P0"   # CMC reschemed their SKUs; older online listings omit this prefix


def _find_supplier_config(name: str, suppliers: list):
    """
    Return the SupplierConfig whose name matches *name*, or None.

    Handles two forms of match:
    - Exact (case-insensitive): "Electric Factory" == "ELECTRIC FACTORY"
    - Prefix (truncated CSV names): "ELECTRI" → "Electric Factory"
      Only used when the candidate is at least 4 characters.
    """
    if not name or not suppliers:
        return None
    name_upper = name.upper().strip()
    for s in suppliers:
        if s.name.upper() == name_upper:
            return s
    if len(name_upper) >= 4:
        for s in suppliers:
            if s.name.upper().startswith(name_upper):
                return s
    return None


def _sku_variants(item: "InvoiceItem", supplier_cfg) -> set:
    """
    Return all normalised uppercase SKU keys that should map to *item* in the
    invoice lookup, implementing three matching passes plus a CMC special case.

    Pass 1 – as-is:
        The sku_with_suffix produced by the parser (character substitutions +
        suffix already applied).  This is the standard Neto SKU form.

    Pass 2 – suffix stripped:
        If sku_with_suffix already ends (or starts, for prepend) with the
        supplier's suffix, the version *without* the suffix is also registered.
        Handles orders where the Neto SKU is stored without a suffix.
        Also catches double-suffix situations (CSV SKU already had the suffix
        before the parser applied it again).

    Pass 3 – suffix appended:
        If sku_with_suffix does *not* already carry the suffix, the version
        *with* the suffix added is also registered.
        Handles orders whose Neto SKU includes the suffix but the CSV row did not.

    CMC special case:
        CMC changed their SKU numbering scheme.  Our online listings were set up
        before the change, so many Neto/eBay SKUs omit the "P0" prefix that
        appears in CMC's own invoices.  For every variant that starts with "P0",
        the P0-stripped form (with and without the CMC suffix) is added.
    """
    base = item.sku_with_suffix.upper().strip()
    variants: set = {base}

    if not supplier_cfg or not supplier_cfg.suffix:
        # No-suffix supplier: also register the raw SKU as a fallback key.
        variants.add(item.sku.upper().strip())
        variants.discard("")
        return variants

    sfx = supplier_cfg.suffix.upper()
    pos = getattr(supplier_cfg, "suffix_position", "append")

    if pos == "prepend":
        if base.startswith(sfx):
            variants.add(base[len(sfx):])       # Pass 2: strip prefix
        else:
            variants.add(f"{sfx}{base}")        # Pass 3: add prefix
    else:
        if base.endswith(sfx):
            variants.add(base[:-len(sfx)])      # Pass 2: strip suffix
        else:
            variants.add(f"{base}{sfx}")        # Pass 3: add suffix

    # CMC special: strip "P0" from any variant to bridge old/new SKU numbering
    if sfx == _CMC_SUFFIX:
        for v in list(variants):
            if v.startswith(_CMC_P0_PREFIX) and len(v) > len(_CMC_P0_PREFIX):
                no_p0 = v[len(_CMC_P0_PREFIX):]
                variants.add(no_p0)
                if not no_p0.endswith(_CMC_SUFFIX):
                    variants.add(f"{no_p0}{_CMC_SUFFIX}")

    variants.discard("")
    return variants


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
    # Each item is registered under all SKU variants (see _sku_variants docstring):
    #   Pass 1 – sku_with_suffix as-is
    #   Pass 2 – suffix stripped (catches orders that store the bare SKU)
    #   Pass 3 – suffix appended (catches orders that store the suffixed SKU when
    #             the CSV row did not include it, or when it was double-applied)
    #   CMC    – P0-prefix stripped variants for old/new CMC SKU numbering
    invoice_lookup: dict[str, InvoiceItem] = {}
    for item in invoice_items:
        sup_cfg = _find_supplier_config(item.supplier_name, suppliers or [])
        for key in _sku_variants(item, sup_cfg):
            if key and key not in invoice_lookup:
                invoice_lookup[key] = item

    # Build alias lookup: Neto order SKU (upper) → InvoiceItem (via alias file)
    # Raw invoice SKUs from the alias are transformed using the mapping's supplier config
    # (character substitutions + suffix) before being looked up in invoice_lookup.
    # Because invoice_lookup now holds multiple variant keys per item, even aliases
    # with bare (unsuffixed) invoice SKUs will resolve correctly.
    alias_lookup: dict[str, InvoiceItem] = {}
    if sku_alias_manager:
        for neto_sku, mapping in sku_alias_manager.get_all().items():
            supplier_name = mapping.get("supplier", "")
            for raw_inv_sku in mapping["invoice_skus"]:
                final_sku = _apply_supplier_transform(raw_inv_sku, supplier_name, suppliers or [])
                inv = invoice_lookup.get(final_sku.upper().strip())
                if not inv:
                    # Also try the raw form — covers aliases whose invoice SKU matches
                    # one of the bare-key variants now registered in invoice_lookup.
                    inv = invoice_lookup.get(raw_inv_sku.upper().strip())
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
