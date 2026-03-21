"""
MUSIPOS Bulk Operations - PO Ordering Engine

Fetches pending orders from eBay + Neto, identifies out-of-stock items,
and creates/updates purchase orders via direct SQL writes.

Replaces the old pyautogui GUI automation approach.
"""

import json
import logging
import os
import re
from collections import defaultdict
from dataclasses import dataclass, field
from datetime import datetime
from decimal import Decimal, ROUND_HALF_UP
from typing import Callable, Dict, List, Optional

from config import (EBAY_CONFIG, NETO_CONFIG, PG_TRACKING_CONFIG,
                    INVENTORY_XLSX_PATH, DEFAULT_WAREHOUSE_ID,
                    DEFAULT_USER_ID, COMPUTER_ID)
from db import get_connection, transaction, dry_run_transaction
from models import ItemInfo
from order_fetch import fetch_all_orders, FetchResult, FetchedOrderLine
from sku_mapping import SkuMappingStore, KitMapping
from validators import resolve_sku, resolve_sku_with_mapping

logger = logging.getLogger(__name__)

GST_RATE = Decimal("1.1")

PROCESSED_ORDERS_PATH = os.path.join(os.path.dirname(__file__), "processed_orders.json")
DISREGARDED_ORDERS_PATH = os.path.join(os.path.dirname(__file__), "disregarded_orders.json")


def _load_processed_orders():
    """Load set of previously processed order line keys."""
    if os.path.exists(PROCESSED_ORDERS_PATH):
        try:
            with open(PROCESSED_ORDERS_PATH, "r") as f:
                data = json.load(f)
            return set(data.get("keys", []))
        except Exception:
            logger.warning("Could not load processed_orders.json, starting fresh")
    return set()


def _save_processed_orders(keys):
    """Save processed order line keys to disk."""
    try:
        with open(PROCESSED_ORDERS_PATH, "w") as f:
            json.dump({"keys": sorted(keys)}, f, indent=2)
    except Exception as e:
        logger.error("Failed to save processed_orders.json: %s", e)


def mark_orders_processed(items):
    """Mark order lines from these items as processed (added to PO).

    Key format: '{order_id}|{sku}' — unique per order line.
    """
    existing = _load_processed_orders()
    for item in items:
        for detail in item.order_line_details:
            key = "{}|{}".format(detail["order_id"], item.cleaned_sku)
            existing.add(key)
    _save_processed_orders(existing)
    logger.info("Marked %d order lines as processed (total: %d)",
                sum(len(i.order_line_details) for i in items), len(existing))


# ──────────────────────────────────────────
# Disregarded orders (manually dismissed)
# ──────────────────────────────────────────

def _load_disregarded_orders():
    """Load dict of disregarded order line keys with metadata."""
    if os.path.exists(DISREGARDED_ORDERS_PATH):
        try:
            with open(DISREGARDED_ORDERS_PATH, "r") as f:
                return json.load(f)
        except Exception:
            logger.warning("Could not load disregarded_orders.json, starting fresh")
    return {"keys": {}}


def _save_disregarded_orders(data):
    """Save disregarded orders to disk."""
    try:
        with open(DISREGARDED_ORDERS_PATH, "w") as f:
            json.dump(data, f, indent=2)
    except Exception as e:
        logger.error("Failed to save disregarded_orders.json: %s", e)


def disregard_order_lines(items):
    """Mark order lines as disregarded (manually dismissed from shortfall).

    Args:
        items: List[OrderDemandItem] to disregard.

    Returns count of new keys added.
    """
    data = _load_disregarded_orders()
    keys = data.get("keys", {})
    added = 0
    for item in items:
        for detail in item.order_line_details:
            key = "{}|{}".format(detail["order_id"], item.cleaned_sku)
            if key not in keys:
                keys[key] = {
                    "sku": item.cleaned_sku,
                    "product_name": item.product_name,
                    "order_id": detail["order_id"],
                    "channel": detail["channel"],
                    "disregarded_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                }
                added += 1
    data["keys"] = keys
    _save_disregarded_orders(data)
    logger.info("Disregarded %d order lines (total: %d)", added, len(keys))
    return added


def get_disregarded_orders():
    """Return list of disregarded order entries for display."""
    data = _load_disregarded_orders()
    return list(data.get("keys", {}).values())


def remove_disregarded(key):
    """Remove a single key from disregarded orders. Returns True if found."""
    data = _load_disregarded_orders()
    keys = data.get("keys", {})
    if key in keys:
        del keys[key]
        data["keys"] = keys
        _save_disregarded_orders(data)
        return True
    return False


# ──────────────────────────────────────────
# Dataclasses
# ──────────────────────────────────────────

@dataclass
class OrderDemandItem:
    """A collated demand item across all orders."""
    cleaned_sku: str
    raw_skus: List[str]
    total_qty: int
    product_name: str
    order_ids: List[str]
    channels: List[str]         # "ebay", "neto"
    is_express: bool
    item: Optional[ItemInfo]    # resolved MUSIPOS item (None if unresolvable)
    qty_on_hand: int
    shortfall: int              # total_qty - qty_on_hand (only positive)
    supplier_id: str
    supplier_name: str
    order_qty: int = 0          # qty to order (editable, defaults to total_qty)
    decision: str = ""          # "po", "dropship", or ""
    earliest_order_date: str = ""  # ISO datetime of earliest order for this SKU
    # Per-line details for note-adding: [{order_id, channel, ebay_olid}]
    order_line_details: List[dict] = field(default_factory=list)
    # Alternate suppliers if SKU is ambiguous: [{supplier_id, supplier_name}]
    alt_suppliers: List[dict] = field(default_factory=list)
    # Sticky notes from orders (deduplicated, non-empty)
    sticky_notes: List[str] = field(default_factory=list)


@dataclass
class POCreationResult:
    """Result of creating purchase orders."""
    success: bool
    message: str
    pos_created: List[dict] = field(default_factory=list)
    dropship_items: List[dict] = field(default_factory=list)
    errors: List[str] = field(default_factory=list)
    details: List[str] = field(default_factory=list)
    dry_run: bool = True


# ──────────────────────────────────────────
# Fetch & Analyze
# ──────────────────────────────────────────

def fetch_and_analyze_orders(progress_cb=None):
    """Fetch orders from eBay + Neto, collate by SKU, check stock.

    Returns (shortfall_items, metadata) where:
        shortfall_items: List[OrderDemandItem] with shortfall > 0
        metadata: {ebay_count, neto_count, total_skus, shortfall_count, errors}
    """
    def log(msg):
        logger.info(msg)
        if progress_cb:
            progress_cb(msg)

    # 1. Fetch orders
    log("Fetching orders from eBay + Neto...")
    fetch_result = fetch_all_orders(
        ebay_config=EBAY_CONFIG,
        neto_config=NETO_CONFIG,
        pg_config=PG_TRACKING_CONFIG,
        inventory_path=INVENTORY_XLSX_PATH,
        progress_cb=progress_cb,
        filter_by_tracking=False,
    )

    metadata = {
        "ebay_count": fetch_result.ebay_count,
        "neto_count": fetch_result.neto_count,
        "total_lines": len(fetch_result.active_lines),
        "errors": fetch_result.errors,
    }

    if not fetch_result.active_lines:
        log("No active order lines found")
        metadata["total_skus"] = 0
        metadata["shortfall_count"] = 0
        metadata["books_items"] = []
        metadata["books_count"] = 0
        return [], metadata

    # 2. Load kit mapping store
    mapping_store = SkuMappingStore()

    # 3. Collate active lines by cleaned_sku
    log("Collating order lines by SKU...")
    collated = defaultdict(lambda: {
        "raw_skus": set(),
        "total_qty": 0,
        "product_name": "",
        "order_ids": set(),
        "channels": set(),
        "is_express": False,
        "order_line_details": [],
        "min_order_date": "",
        "sticky_notes": set(),
    })

    # Detect express from eBay ExpeditedService flag or shipping name
    def _is_express(line):
        if getattr(line, 'is_expedited', False):
            return True
        s = (line.shipping_option or "").lower()
        return "express" in s or "priority" in s

    # Expand kits first, then collate
    expanded_lines = []
    for line in fetch_result.active_lines:
        sku = line.cleaned_sku.strip()
        if not sku:
            continue

        express = _is_express(line)
        mapping = mapping_store.lookup(sku)
        addr_fields = {
            "customer_name": line.customer_name,
            "customer_address": line.customer_address,
            "customer_city": line.customer_city,
            "customer_state": line.customer_state,
            "customer_postcode": line.customer_postcode,
            "customer_country": line.customer_country,
            "customer_phone": line.customer_phone,
        }

        if isinstance(mapping, KitMapping):
            # Expand kit into component lines
            for comp in mapping.components:
                expanded_lines.append({
                    "cleaned_sku": comp.itm_iid,
                    "raw_sku": line.raw_sku,
                    "qty": line.qty * comp.qty,
                    "product_name": line.item_title,
                    "order_id": line.order_number,
                    "channel": line.source,
                    "is_express": express,
                    "ebay_olid": line.ebay_olid,
                    "order_date": line.order_date,
                    "sticky_note": line.sticky_note,
                    **addr_fields,
                })
        else:
            # Auto-detect SKU-N bundle pattern (e.g. 135JTT-20 = 20x 135JTT)
            # Only applies when:
            #   1. No explicit mapping exists
            #   2. The listing title confirms it's a bundle (starts with "N x" or "Nx")
            # This avoids splitting real SKUs like 1SKB-4 which aren't bundles
            bundle_match = re.match(r'^(.+)-(\d+)$', sku)
            if bundle_match and not mapping:
                base_sku = bundle_match.group(1)
                multiplier = int(bundle_match.group(2))
                title = line.item_title.strip()
                title_confirms = bool(
                    multiplier > 1 and
                    re.match(r'^\s*' + str(multiplier) + r'\s*[xX×]\s', title)
                )
                if title_confirms:
                    expanded_lines.append({
                        "cleaned_sku": base_sku,
                        "raw_sku": line.raw_sku,
                        "qty": line.qty * multiplier,
                        "product_name": line.item_title,
                        "order_id": line.order_number,
                        "channel": line.source,
                        "is_express": express,
                        "ebay_olid": line.ebay_olid,
                        "order_date": line.order_date,
                        "sticky_note": line.sticky_note,
                        **addr_fields,
                    })
                    continue

            expanded_lines.append({
                "cleaned_sku": sku,
                "raw_sku": line.raw_sku,
                "qty": line.qty,
                "product_name": line.item_title,
                "order_id": line.order_number,
                "channel": line.source,
                "is_express": express,
                "ebay_olid": line.ebay_olid,
                "order_date": line.order_date,
                "sticky_note": line.sticky_note,
                **addr_fields,
            })

    # Filter out order lines that were previously processed (added to PO)
    processed_keys = _load_processed_orders()
    if processed_keys:
        before_count = len(expanded_lines)
        expanded_lines = [
            el for el in expanded_lines
            if "{}|{}".format(el["order_id"], el["cleaned_sku"]) not in processed_keys
        ]
        filtered = before_count - len(expanded_lines)
        if filtered:
            log("Filtered {} previously-processed order lines".format(filtered))

    # Filter out order lines that were manually disregarded
    disregarded_data = _load_disregarded_orders()
    disregarded_keys = set(disregarded_data.get("keys", {}).keys())
    if disregarded_keys:
        before_count = len(expanded_lines)
        expanded_lines = [
            el for el in expanded_lines
            if "{}|{}".format(el["order_id"], el["cleaned_sku"]) not in disregarded_keys
        ]
        filtered = before_count - len(expanded_lines)
        if filtered:
            log("Filtered {} disregarded order lines".format(filtered))

    for el in expanded_lines:
        sku = el["cleaned_sku"]
        c = collated[sku]
        c["raw_skus"].add(el["raw_sku"])
        c["total_qty"] += el["qty"]
        if not c["product_name"]:
            c["product_name"] = el["product_name"]
        c["order_ids"].add(el["order_id"])
        c["channels"].add(el["channel"])
        if el["is_express"]:
            c["is_express"] = True
        c["order_line_details"].append({
            "order_id": el["order_id"],
            "channel": el["channel"],
            "ebay_olid": el.get("ebay_olid", ""),
            "qty": el["qty"],
            "sku": el["cleaned_sku"],
            "customer_name": el.get("customer_name", ""),
            "customer_address": el.get("customer_address", ""),
            "customer_city": el.get("customer_city", ""),
            "customer_state": el.get("customer_state", ""),
            "customer_postcode": el.get("customer_postcode", ""),
            "customer_country": el.get("customer_country", ""),
            "customer_phone": el.get("customer_phone", ""),
        })
        # Collect sticky notes
        sn = (el.get("sticky_note") or "").strip()
        if sn:
            c["sticky_notes"].add(sn)
        # Track earliest order date for chronological sorting
        od = el.get("order_date", "")
        if od and (not c["min_order_date"] or od < c["min_order_date"]):
            c["min_order_date"] = od

    metadata["total_skus"] = len(collated)

    # 4. Resolve each SKU and check stock
    log("Resolving SKUs and checking stock...")
    BOOKS_SUPPLIER_IDS = {"ALFRED", "PRINT M", "ENCORE"}
    conn = get_connection()
    shortfall_items = []
    books_items = []  # Items from book suppliers (exclusively dropshipped)
    try:
        cursor = conn.cursor()
        # Sort by earliest order date (chronological), then SKU as tiebreaker
        sorted_skus = sorted(collated.items(),
                             key=lambda x: (x[1]["min_order_date"] or "9999", x[0]))
        for sku, info in sorted_skus:
            item = resolve_sku(cursor, sku, DEFAULT_WAREHOUSE_ID)

            # Ernie Ball fallback: titles starting with "Ernie Ball" use P0-prefixed SKUs
            if not item and not sku.startswith("P0"):
                # Check if any order title for this SKU starts with "Ernie Ball"
                if info["product_name"].lower().startswith("ernie ball"):
                    item = resolve_sku(cursor, "P0" + sku, DEFAULT_WAREHOUSE_ID)
                    if item:
                        logger.info("Ernie Ball P0 prefix: %s -> P0%s", sku, sku)

            qty_on_hand = int(item.qpc_qty_on_hand) if item else 0
            shortfall = info["total_qty"] - qty_on_hand

            if shortfall <= 0:
                continue

            # Look up supplier name
            supplier_id = item.itm_supplier_id if item else ""
            supplier_name = ""
            if supplier_id:
                cursor.execute("""
                    SELECT rsp_name FROM ap4rsp
                    WHERE rsp_supplier_id = ?
                """, supplier_id)
                row = cursor.fetchone()
                if row:
                    supplier_name = (row.rsp_name or "").strip()

            # Check for ambiguous suppliers (same SKU from multiple suppliers)
            SKIP_SUPPLIER_IDS = {"HALLEONAR", "ALFRED"}
            SKIP_SUPPLIER_NAMES = {"HAL LEONARD AUSTRALIA PTY LTD"}

            def _supplier_excluded(sid, sname):
                return (sid.strip() in SKIP_SUPPLIER_IDS
                        or sname.strip().upper() in SKIP_SUPPLIER_NAMES)

            alt_suppliers = []
            from validators import check_sku_ambiguous
            matches = check_sku_ambiguous(cursor, sku)
            if len(matches) > 1:
                for alt_sid, alt_iid in matches:
                    alt_sid = alt_sid.strip()
                    cursor.execute("""
                        SELECT rsp_name FROM ap4rsp
                        WHERE rsp_supplier_id = ?
                    """, alt_sid)
                    alt_row = cursor.fetchone()
                    alt_name = (alt_row.rsp_name or "").strip() if alt_row else ""
                    if _supplier_excluded(alt_sid, alt_name):
                        continue
                    alt_suppliers.append({
                        "supplier_id": alt_sid,
                        "supplier_name": alt_name,
                    })

            # If default supplier is excluded, try to fall back to an alternative
            if _supplier_excluded(supplier_id, supplier_name):
                if alt_suppliers:
                    supplier_id = alt_suppliers[0]["supplier_id"]
                    supplier_name = alt_suppliers[0]["supplier_name"]
                    item = resolve_sku(cursor, sku, DEFAULT_WAREHOUSE_ID,
                                       supplier_id=supplier_id)
                    qty_on_hand = int(item.qpc_qty_on_hand) if item else 0
                    shortfall = info["total_qty"] - qty_on_hand
                    if shortfall <= 0:
                        continue
                else:
                    # Book suppliers: collect ALL in-print variants for Books tab
                    # Skip if sticky notes indicate already handled
                    notes_lower = " ".join(info["sticky_notes"]).lower()
                    if any(kw in notes_lower for kw in
                           ("dropship", "invoiced", "on po", "messaged customer",
                            "sent encore", "emailed", "checking eta", "getting eta", "anna")):
                        logger.info("Books: %s — skipped (sticky note: %s)",
                                    sku, notes_lower[:60])
                        continue

                    # Gather all book supplier variants for this SKU
                    book_variants = []
                    all_matches = matches if len(matches) > 1 else [(supplier_id, sku)]
                    for match_sid, match_iid in all_matches:
                        msid = match_sid.strip()
                        if msid not in BOOKS_SUPPLIER_IDS and not _supplier_excluded(msid, ""):
                            continue  # Not a book supplier
                        # Resolve this supplier's version
                        match_item = resolve_sku(cursor, sku, DEFAULT_WAREHOUSE_ID,
                                                 supplier_id=msid)
                        if not match_item:
                            match_item = item  # fallback to default
                        # Filter OOP (itm_status = "C")
                        if match_item and match_item.itm_status.strip().upper() == "C":
                            logger.info("Books: %s — %s is OOP, skipping", sku, msid)
                            continue
                        # Look up supplier name
                        cursor.execute("SELECT rsp_name FROM ap4rsp WHERE rsp_supplier_id = ?", msid)
                        srow = cursor.fetchone()
                        sname = (srow.rsp_name or "").strip() if srow else msid
                        book_variants.append((msid, sname, match_item))

                    # Also check Encore CDN for this SKU (might not be in MUSIPOS at all)
                    from encore_ordering import search_product as encore_search
                    encore_product = encore_search(sku)
                    if encore_product and encore_product.get("stock_status") != "out_of_stock":
                        # Only add if no MUSIPOS variant already covers Encore
                        encore_covered = any(v[0] in ("ENCORE", "PRINT M") for v in book_variants)
                        if not encore_covered:
                            book_variants.append(("ENCORE", "ENCORE MUSIC (via catalogue)", None))

                    if not book_variants and supplier_id.strip() not in BOOKS_SUPPLIER_IDS:
                        logger.info("Skipping %s — supplier %s (%s) excluded, no alternatives",
                                    sku, supplier_id, supplier_name)
                        continue

                    for bsid, bsname, bitem in book_variants:
                        books_items.append(OrderDemandItem(
                            cleaned_sku=sku,
                            raw_skus=sorted(info["raw_skus"]),
                            total_qty=info["total_qty"],
                            product_name=info["product_name"],
                            order_ids=sorted(info["order_ids"]),
                            channels=sorted(info["channels"]),
                            is_express=info["is_express"],
                            item=bitem,
                            qty_on_hand=int(bitem.qpc_qty_on_hand) if bitem else 0,
                            shortfall=info["total_qty"],
                            supplier_id=bsid,
                            supplier_name=bsname,
                            order_qty=info["total_qty"],
                            decision="dropship",
                            earliest_order_date=info["min_order_date"],
                            order_line_details=info["order_line_details"],
                            sticky_notes=sorted(info["sticky_notes"]),
                        ))
                        logger.info("Books: %s — supplier %s (%s)", sku, bsid, bsname)
                    continue

            # Skip items with no resolved supplier — but check Encore catalogue first
            if not supplier_id.strip():
                from encore_ordering import search_product as encore_search
                encore_product = encore_search(sku)
                if encore_product:
                    # Check sticky notes before adding
                    notes_lower = " ".join(info["sticky_notes"]).lower()
                    if not any(kw in notes_lower for kw in
                               ("dropship", "invoiced", "on po", "messaged customer",
                                "sent encore", "emailed", "checking eta", "getting eta", "anna")):
                        books_items.append(OrderDemandItem(
                            cleaned_sku=sku,
                            raw_skus=sorted(info["raw_skus"]),
                            total_qty=info["total_qty"],
                            product_name=encore_product.get("name") or info["product_name"],
                            order_ids=sorted(info["order_ids"]),
                            channels=sorted(info["channels"]),
                            is_express=info["is_express"],
                            item=None,
                            qty_on_hand=0,
                            shortfall=info["total_qty"],
                            supplier_id="ENCORE",
                            supplier_name="ENCORE MUSIC (via catalogue)",
                            order_qty=info["total_qty"],
                            decision="dropship",
                            earliest_order_date=info["min_order_date"],
                            order_line_details=info["order_line_details"],
                            sticky_notes=sorted(info["sticky_notes"]),
                        ))
                        logger.info("Books (Encore fallback): %s — %s",
                                    sku, encore_product.get("name", ""))
                    continue
                logger.info("Skipping %s — no supplier resolved", sku)
                continue

            shortfall_items.append(OrderDemandItem(
                cleaned_sku=sku,
                raw_skus=sorted(info["raw_skus"]),
                total_qty=info["total_qty"],
                product_name=info["product_name"],
                order_ids=sorted(info["order_ids"]),
                channels=sorted(info["channels"]),
                is_express=info["is_express"],
                item=item,
                qty_on_hand=qty_on_hand,
                shortfall=shortfall,
                supplier_id=supplier_id,
                supplier_name=supplier_name,
                order_qty=info["total_qty"],
                earliest_order_date=info["min_order_date"],
                order_line_details=info["order_line_details"],
                alt_suppliers=alt_suppliers,
                sticky_notes=sorted(info["sticky_notes"]),
            ))
    finally:
        conn.close()

    metadata["shortfall_count"] = len(shortfall_items)
    metadata["books_items"] = books_items
    metadata["books_count"] = len(books_items)
    log("Found {} SKUs with shortfall out of {} total ({} books)".format(
        len(shortfall_items), len(collated), len(books_items)))

    return shortfall_items, metadata


# ──────────────────────────────────────────
# Create Purchase Orders
# ──────────────────────────────────────────

def create_purchase_orders(items, dry_run=True, progress_cb=None):
    """Create or update POs for items with decision == 'po'.

    Groups items by supplier. For each supplier:
    - If a CURRENT PO exists, add lines to it (or update qty if item already on PO)
    - If no CURRENT PO, allocate a new one from ap4rsp.rsp_curr_po_no

    Items with decision == 'dropship' are collected but not processed.

    Returns POCreationResult.
    """
    def log(msg):
        logger.info(msg)
        if progress_cb:
            progress_cb(msg)

    po_items = [i for i in items if i.decision == "po" and i.item is not None]
    dropship_items = [i for i in items if i.decision == "dropship"]

    if not po_items:
        return POCreationResult(
            success=True,
            message="No items marked for PO",
            dropship_items=[_item_to_dict(i) for i in dropship_items],
            dry_run=dry_run,
        )

    ctx = dry_run_transaction if dry_run else transaction
    pos_created = []
    details = []
    errors = []

    log("Mode: {}".format("DRY RUN" if dry_run else "LIVE"))
    details.append("MODE: {}".format("DRY RUN (no DB changes)" if dry_run else "LIVE"))

    # Filter out items with order_qty <= 0
    skipped = []
    valid_po_items = []
    for item in po_items:
        if item.order_qty <= 0:
            skipped.append(item)
            details.append("SKIP: {} order_qty={} (must be > 0)".format(
                item.cleaned_sku, item.order_qty))
        else:
            valid_po_items.append(item)
    po_items = valid_po_items

    if not po_items:
        msg = "No items to process"
        if skipped:
            msg += " ({} skipped with order_qty=0)".format(len(skipped))
        return POCreationResult(
            success=True,
            message=msg,
            dropship_items=[_item_to_dict(i) for i in dropship_items],
            errors=errors,
            details=details,
            dry_run=dry_run,
        )

    # Group by supplier
    by_supplier = defaultdict(list)
    for item in po_items:
        by_supplier[item.supplier_id].append(item)

    try:
        with ctx() as conn:
            cursor = conn.cursor()

            for supplier_id, supplier_items in sorted(by_supplier.items()):
                supplier_name = supplier_items[0].supplier_name
                log("Processing supplier: {} ({})...".format(supplier_id, supplier_name))

                # Check for existing CURRENT PO
                cursor.execute("""
                    SELECT phd_po_no FROM sp4phd
                    WHERE phd_supplier_id = ? AND phd_po_status = 'CURRENT'
                    ORDER BY phd_po_no DESC
                """, supplier_id)
                existing_row = cursor.fetchone()

                if existing_row:
                    po_no = int(existing_row.phd_po_no)
                    po_no_str = str(po_no)
                    details.append("EXISTING PO: {} for {} ({})".format(
                        po_no_str, supplier_id, supplier_name))

                    # Get current max line number
                    cursor.execute("""
                        SELECT ISNULL(MAX(pop_lno), 0) FROM sp4pop
                        WHERE pop_supplier_id = ? AND pop_po_no = ?
                    """, supplier_id, po_no)
                    max_lno = cursor.fetchone()[0]

                    added = 0
                    updated = 0
                    for item in supplier_items:
                        # Check if item already on this PO
                        cursor.execute("""
                            SELECT pop_qor FROM sp4pop
                            WHERE pop_supplier_id = ? AND pop_po_no = ? AND pop_iid = ?
                        """, supplier_id, po_no, item.item.itm_iid)
                        existing_line = cursor.fetchone()

                        if existing_line:
                            # Update existing line — add to qty ordered
                            old_qor = int(existing_line.pop_qor)
                            new_qor = old_qor + item.order_qty
                            cursor.execute("""
                                UPDATE sp4pop SET pop_qor = ?
                                WHERE pop_supplier_id = ? AND pop_po_no = ? AND pop_iid = ?
                            """, new_qor, supplier_id, po_no, item.item.itm_iid)
                            if cursor.rowcount == 0:
                                err = "UPDATE sp4pop affected 0 rows for {} on PO {}".format(
                                    item.item.itm_iid, po_no_str)
                                errors.append(err)
                                details.append("  ERROR: " + err)
                            else:
                                details.append(
                                    "  UPDATE sp4pop: {} qty {} -> {} on PO {} (rowcount={})".format(
                                        item.item.itm_iid, old_qor, new_qor, po_no_str,
                                        cursor.rowcount))
                                updated += 1
                                # Update qty on order in stock record
                                cursor.execute("""
                                    UPDATE sp4qpc
                                    SET qpc_qor = ISNULL(qpc_qor, 0) + ?
                                    WHERE qpc_iid = ? AND qpc_warehouse_id = ?
                                """, item.order_qty, item.item.itm_iid,
                                    DEFAULT_WAREHOUSE_ID)
                                details.append(
                                    "  UPDATE sp4qpc: {} qpc_qor += {}".format(
                                        item.item.itm_iid, item.order_qty))
                        else:
                            # Insert new line
                            max_lno += 1
                            cost = item.item.qpc_cost or Decimal("0")
                            cursor.execute("""
                                INSERT INTO sp4pop (
                                    pop_iid, pop_supplier_id, pop_po_no, pop_cid, pop_lno,
                                    pop_po_date, pop_supplier_iid, pop_qor,
                                    pop_curr_cost, pop_curr_order_flag,
                                    pop_qty_rcv, pop_curr_price, pop_total_cost,
                                    pop_itm_status, pop_user_id, pop_computer_id,
                                    pop_comment, pop_qbo, pop_tax_amt,
                                    pop_order_no, pop_process_status
                                ) VALUES (?, ?, ?, 'STOCK001', ?, GETDATE(), ?, ?,
                                          ?, 'N',
                                          0, ?, 0.00,
                                          'C', ?, ?,
                                          'Web ordering', 0, 0.00,
                                          '', 'N')
                            """,
                                item.item.itm_iid, supplier_id, po_no,
                                max_lno,
                                item.item.itm_supplier_iid, item.order_qty,
                                cost,
                                item.item.itm_new_retail_price,
                                DEFAULT_USER_ID, COMPUTER_ID)
                            if cursor.rowcount == 0:
                                err = "INSERT sp4pop affected 0 rows for {} on PO {}".format(
                                    item.item.itm_iid, po_no_str)
                                errors.append(err)
                                details.append("  ERROR: " + err)
                            else:
                                details.append(
                                    "  INSERT sp4pop: {} qty={} cost={} lno={} on PO {} (rowcount={})".format(
                                        item.item.itm_iid, item.order_qty, cost, max_lno,
                                        po_no_str, cursor.rowcount))
                                added += 1
                                # Update qty on order in stock record
                                cursor.execute("""
                                    UPDATE sp4qpc
                                    SET qpc_qor = ISNULL(qpc_qor, 0) + ?
                                    WHERE qpc_iid = ? AND qpc_warehouse_id = ?
                                """, item.order_qty, item.item.itm_iid,
                                    DEFAULT_WAREHOUSE_ID)
                                details.append(
                                    "  UPDATE sp4qpc: {} qpc_qor += {}".format(
                                        item.item.itm_iid, item.order_qty))

                    summary = "Added {} items, updated {} items on existing PO {} for {}".format(
                        added, updated, po_no_str, supplier_name)
                    log(summary)
                    pos_created.append({
                        "supplier_id": supplier_id,
                        "supplier_name": supplier_name,
                        "po_no": po_no_str,
                        "line_count": len(supplier_items),
                        "new_po": False,
                        "summary": summary,
                    })

                else:
                    # Allocate new PO number
                    cursor.execute("""
                        SELECT rsp_curr_po_no
                        FROM ap4rsp WITH (UPDLOCK, HOLDLOCK)
                        WHERE rsp_supplier_id = ?
                    """, supplier_id)
                    rsp_row = cursor.fetchone()
                    if not rsp_row or rsp_row.rsp_curr_po_no is None:
                        err = "Cannot allocate PO: supplier '{}' has no rsp_curr_po_no".format(
                            supplier_id)
                        errors.append(err)
                        details.append("ERROR: " + err)
                        continue

                    new_po_no = rsp_row.rsp_curr_po_no + 1
                    cursor.execute("""
                        UPDATE ap4rsp SET rsp_curr_po_no = ?
                        WHERE rsp_supplier_id = ?
                    """, new_po_no, supplier_id)
                    po_no_str = str(new_po_no)
                    details.append("ALLOCATE PO: {} from ap4rsp (was {})".format(
                        po_no_str, new_po_no - 1))

                    # Insert PO header
                    cursor.execute("""
                        INSERT INTO sp4phd (
                            phd_supplier_id, phd_po_no, phd_po_date, phd_po_status,
                            phd_memo, phd_po_dest
                        ) VALUES (?, ?, GETDATE(), 'CURRENT', 'Web ordering', 'S')
                    """, supplier_id, new_po_no)
                    details.append("INSERT sp4phd: supplier={}, po={}, status=CURRENT (rowcount={})".format(
                        supplier_id, po_no_str, cursor.rowcount))

                    # Insert PO lines
                    for idx, item in enumerate(supplier_items, start=1):
                        cost = item.item.qpc_cost or Decimal("0")
                        cursor.execute("""
                            INSERT INTO sp4pop (
                                pop_iid, pop_supplier_id, pop_po_no, pop_cid, pop_lno,
                                pop_po_date, pop_supplier_iid, pop_qor,
                                pop_curr_cost, pop_curr_order_flag,
                                pop_qty_rcv, pop_curr_price, pop_total_cost,
                                pop_itm_status, pop_user_id, pop_computer_id,
                                pop_comment, pop_qbo, pop_tax_amt,
                                pop_order_no, pop_process_status
                            ) VALUES (?, ?, ?, 'STOCK001', ?, GETDATE(), ?, ?,
                                      ?, 'N',
                                      0, ?, 0.00,
                                      'C', ?, ?,
                                      'Web ordering', 0, 0.00,
                                      '', 'N')
                        """,
                            item.item.itm_iid, supplier_id, new_po_no,
                            idx,
                            item.item.itm_supplier_iid, item.order_qty,
                            cost,
                            item.item.itm_new_retail_price,
                            DEFAULT_USER_ID, COMPUTER_ID)
                        details.append(
                            "  INSERT sp4pop: line {} {} qty={} cost={} (rowcount={})".format(
                                idx, item.item.itm_iid, item.order_qty, cost, cursor.rowcount))
                        # Update qty on order in stock record
                        cursor.execute("""
                            UPDATE sp4qpc
                            SET qpc_qor = ISNULL(qpc_qor, 0) + ?
                            WHERE qpc_iid = ? AND qpc_warehouse_id = ?
                        """, item.order_qty, item.item.itm_iid,
                            DEFAULT_WAREHOUSE_ID)
                        details.append(
                            "  UPDATE sp4qpc: {} qpc_qor += {}".format(
                                item.item.itm_iid, item.order_qty))

                    summary = "Created new PO {} for {} — {} items".format(
                        po_no_str, supplier_name, len(supplier_items))
                    log(summary)
                    pos_created.append({
                        "supplier_id": supplier_id,
                        "supplier_name": supplier_name,
                        "po_no": po_no_str,
                        "line_count": len(supplier_items),
                        "new_po": True,
                        "summary": summary,
                    })

        # Post-commit verification: confirm data actually persisted
        if not dry_run and pos_created:
            details.append("--- POST-COMMIT VERIFICATION ---")
            verify_conn = get_connection()
            try:
                verify_cursor = verify_conn.cursor()
                for po_info in pos_created:
                    po_no_int = int(po_info["po_no"])
                    sid = po_info["supplier_id"]
                    # Show all lines on this PO with full detail
                    verify_cursor.execute("""
                        SELECT pop_iid, pop_cid, pop_lno, pop_qor, pop_qty_rcv,
                               pop_curr_cost, pop_curr_order_flag, pop_po_date,
                               pop_supplier_iid, pop_itm_title
                        FROM sp4pop
                        WHERE pop_supplier_id = ? AND pop_po_no = ?
                        ORDER BY pop_lno
                    """, sid, po_no_int)
                    rows = verify_cursor.fetchall()
                    details.append("  PO {}: {} total lines in DB".format(
                        po_info["po_no"], len(rows)))
                    # Show the items we just modified
                    modified_iids = set()
                    for si in by_supplier.get(sid, []):
                        modified_iids.add(si.item.itm_iid)
                    for row in rows:
                        iid = (row.pop_iid or "").strip()
                        if iid in modified_iids:
                            details.append(
                                "  >> {} cid='{}' lno={} qor={} rcv={} cost={} flag='{}' date={}".format(
                                    iid, (row.pop_cid or "").strip(),
                                    row.pop_lno, row.pop_qor, row.pop_qty_rcv,
                                    row.pop_curr_cost, (row.pop_curr_order_flag or "").strip(),
                                    row.pop_po_date))
                    if not rows:
                        errors.append(
                            "VERIFY FAIL: PO {} has 0 lines after commit!".format(
                                po_info["po_no"]))
                    # Also check PO header status
                    verify_cursor.execute("""
                        SELECT phd_po_status, phd_po_date, phd_memo
                        FROM sp4phd
                        WHERE phd_supplier_id = ? AND phd_po_no = ?
                    """, sid, po_no_int)
                    hdr = verify_cursor.fetchone()
                    if hdr:
                        details.append("  PO header: status='{}' date={} memo='{}'".format(
                            (hdr.phd_po_status or "").strip(),
                            hdr.phd_po_date,
                            (hdr.phd_memo or "").strip()[:60]))
            except Exception as ve:
                details.append("  Verification query failed: {}".format(ve))
            finally:
                verify_conn.close()

        prefix = "[DRY RUN] " if dry_run else ""
        msg = "{}Created/updated {} PO(s) with {} items total".format(
            prefix, len(pos_created), len(po_items))
        if skipped:
            msg += " ({} skipped qty=0)".format(len(skipped))
        if errors:
            msg += " ({} errors)".format(len(errors))

        return POCreationResult(
            success=len(errors) == 0,
            message=msg,
            pos_created=pos_created,
            dropship_items=[_item_to_dict(i) for i in dropship_items],
            errors=errors,
            details=details,
            dry_run=dry_run,
        )

    except Exception as e:
        logger.exception("PO creation failed")
        return POCreationResult(
            success=False,
            message="PO creation failed: {}".format(e),
            errors=[str(e)],
            details=details,
            dry_run=dry_run,
        )


def _item_to_dict(item):
    """Convert OrderDemandItem to a JSON-serializable dict."""
    return {
        "cleaned_sku": item.cleaned_sku,
        "product_name": item.product_name,
        "total_qty": item.total_qty,
        "qty_on_hand": item.qty_on_hand,
        "shortfall": item.shortfall,
        "order_qty": item.order_qty,
        "supplier_id": item.supplier_id,
        "supplier_name": item.supplier_name,
        "channels": item.channels,
        "order_ids": item.order_ids,
        "decision": item.decision,
        "order_line_details": item.order_line_details,
    }


# ──────────────────────────────────────────
# Order Notes (eBay + Neto)
# ──────────────────────────────────────────

def add_order_notes(items, progress_cb=None):
    """Add 'ON PO' notes to eBay and Neto orders for processed items.

    For eBay orders: SetUserNotes API (per OrderLineItemID)
    For Neto orders: UpdateOrder API with StickyNotes (batched)
    eBay orders also get a Neto note if they synced to Neto.

    Args:
        items: List[OrderDemandItem] that were added to POs
        progress_cb: Optional progress callback

    Returns:
        dict with {ebay_ok, ebay_fail, neto_ok, neto_fail,
                    neto_ebay_ok, neto_ebay_fail, errors}
    """
    import requests
    import json

    def log(msg):
        logger.info(msg)
        if progress_cb:
            progress_cb(msg)

    result = {"ebay_ok": 0, "ebay_fail": 0, "neto_ok": 0, "neto_fail": 0,
              "neto_ebay_ok": 0, "neto_ebay_fail": 0,
              "errors": [], "details": []}

    # Collect unique order lines by channel
    ebay_lines = []  # [{order_id, ebay_olid}]
    neto_order_ids = set()
    ebay_order_skus = defaultdict(list)  # {ebay_order_id: [sku1, sku2, ...]}

    for item in items:
        sku = item.cleaned_sku
        note_text = "{} ON PO".format(sku)

        for detail in item.order_line_details:
            if detail["channel"] == "ebay" and detail.get("ebay_olid"):
                ebay_lines.append({
                    "olid": detail["ebay_olid"],
                    "note": note_text,
                })
                ebay_order_skus[detail["order_id"]].append(sku)
            if detail["channel"] == "neto":
                neto_order_ids.add(detail["order_id"])

    # --- eBay: SetUserNotes ---
    if ebay_lines:
        log("Adding notes to {} eBay order lines...".format(len(ebay_lines)))
        try:
            from ebaysdk.trading import Connection as Trading
            api = Trading(
                appid=EBAY_CONFIG["appid"],
                devid=EBAY_CONFIG["devid"],
                certid=EBAY_CONFIG["certid"],
                token=EBAY_CONFIG["token"],
                timeout=60,
                config_file=None,
            )

            for el in ebay_lines:
                olid = el["olid"]
                timestamp = datetime.now().strftime("%m/%d %H:%M")
                final_note = "{}: {}".format(timestamp, el["note"])
                if len(final_note) > 255:
                    final_note = final_note[:252] + "..."

                # Extract ItemID from OrderLineItemID (format: ItemID-TransactionID)
                item_id = olid.split("-")[0] if "-" in olid else ""

                request_data = {
                    "Action": "AddOrUpdate",
                    "NoteText": final_note,
                    "Version": "1149",
                    "OrderLineItemID": olid,
                }
                if item_id:
                    request_data["ItemID"] = item_id

                try:
                    response = api.execute("SetUserNotes", request_data)
                    resp_dict = response.dict()
                    if resp_dict.get("Ack") in ("Success", "Warning"):
                        result["ebay_ok"] += 1
                        result["details"].append(
                            "eBay note OK: {} -> '{}'".format(olid, final_note))
                    else:
                        result["ebay_fail"] += 1
                        err_msg = ""
                        errors = resp_dict.get("Errors", [])
                        if isinstance(errors, dict):
                            errors = [errors]
                        for e in errors:
                            err_msg += e.get("LongMessage", e.get("ShortMessage", ""))
                        result["details"].append(
                            "eBay note FAIL: {} -> {}".format(olid, err_msg))
                except Exception as e:
                    result["ebay_fail"] += 1
                    result["details"].append(
                        "eBay note ERROR: {} -> {}".format(olid, e))

        except ImportError:
            result["errors"].append("ebaysdk not installed — cannot add eBay notes")
        except Exception as e:
            result["errors"].append("eBay API init failed: {}".format(e))

    # --- Neto counterpart for eBay orders ---
    if ebay_order_skus:
        log("Looking up Neto counterparts for {} eBay orders...".format(
            len(ebay_order_skus)))
        try:
            lookup_headers = {
                "Content-Type": "application/json",
                "Accept": "application/json",
                "NETOAPI_ACTION": "GetOrder",
                "NETOAPI_USERNAME": NETO_CONFIG["username"],
                "NETOAPI_KEY": NETO_CONFIG["api_key"],
            }
            lookup_data = {
                "Filter": {
                    "PurchaseOrderNumber": list(ebay_order_skus.keys()),
                    "OutputSelector": ["ID", "PurchaseOrderNumber"],
                }
            }
            r = requests.post(NETO_CONFIG["url"], headers=lookup_headers,
                              json=lookup_data, timeout=30)
            resp = r.json()

            # Build mapping: {ebay_order_id: neto_order_id}
            ebay_to_neto = {}
            orders = resp.get("Order", [])
            if isinstance(orders, dict):
                orders = [orders]
            for order in orders:
                neto_id = order.get("ID") or order.get("OrderID")
                po_num = order.get("PurchaseOrderNumber")
                if neto_id and po_num:
                    ebay_to_neto[po_num] = neto_id

            if ebay_to_neto:
                log("Found {} Neto counterparts, adding notes...".format(
                    len(ebay_to_neto)))

                timestamp = datetime.now().strftime("%d/%m/%Y %H:%M")
                counterpart_data = []
                for ebay_oid, neto_oid in sorted(ebay_to_neto.items()):
                    unique_skus = sorted(set(ebay_order_skus[ebay_oid]))
                    note_text = "{}: {} ON PO".format(
                        timestamp, ", ".join(unique_skus))
                    counterpart_data.append({
                        "OrderID": neto_oid,
                        "StickyNotes": {
                            "StickyNote": [{
                                "Title": "Order Update",
                                "Description": note_text,
                            }]
                        }
                    })

                update_headers = {
                    "Content-Type": "application/json",
                    "Accept": "application/json",
                    "NETOAPI_ACTION": "UpdateOrder",
                    "NETOAPI_USERNAME": NETO_CONFIG["username"],
                    "NETOAPI_KEY": NETO_CONFIG["api_key"],
                }
                r2 = requests.post(NETO_CONFIG["url"], headers=update_headers,
                                   json={"Order": counterpart_data}, timeout=30)
                resp2 = r2.json()
                if resp2.get("Ack") == "Success":
                    result["neto_ebay_ok"] = len(counterpart_data)
                    result["details"].append(
                        "Neto(eBay) notes OK: {} orders updated".format(
                            len(counterpart_data)))
                else:
                    result["neto_ebay_fail"] = len(counterpart_data)
                    err_msg = resp2.get("Messages", str(resp2))
                    result["errors"].append(
                        "Neto(eBay) API error: {}".format(err_msg))
                    result["details"].append(
                        "Neto(eBay) notes FAIL: {}".format(err_msg))
            else:
                log("No Neto counterparts found for eBay orders")
        except Exception as e:
            result["neto_ebay_fail"] = len(ebay_order_skus)
            result["errors"].append(
                "Neto(eBay) counterpart lookup failed: {}".format(e))

    # --- Neto: UpdateOrder with StickyNotes ---
    if neto_order_ids:
        log("Adding notes to {} Neto orders...".format(len(neto_order_ids)))

        # Build per-order note text (combine all SKUs for that order)
        order_notes = defaultdict(list)
        for item in items:
            for detail in item.order_line_details:
                if detail["channel"] == "neto" and detail["order_id"] in neto_order_ids:
                    order_notes[detail["order_id"]].append(item.cleaned_sku)

        timestamp = datetime.now().strftime("%d/%m/%Y %H:%M")
        order_data = []
        for order_id, skus in sorted(order_notes.items()):
            unique_skus = sorted(set(skus))
            note_text = "{}: {} ON PO".format(timestamp, ", ".join(unique_skus))
            order_data.append({
                "OrderID": order_id,
                "StickyNotes": {
                    "StickyNote": [{
                        "Title": "Order Update",
                        "Description": note_text,
                    }]
                }
            })

        try:
            headers = {
                "Content-Type": "application/json",
                "Accept": "application/json",
                "NETOAPI_ACTION": "UpdateOrder",
                "NETOAPI_USERNAME": NETO_CONFIG["username"],
                "NETOAPI_KEY": NETO_CONFIG["api_key"],
            }
            data = {"Order": order_data}
            r = requests.post(NETO_CONFIG["url"], headers=headers,
                              json=data, timeout=30)
            resp = r.json()
            if resp.get("Ack") == "Success":
                result["neto_ok"] = len(order_data)
                result["details"].append(
                    "Neto notes OK: {} orders updated".format(len(order_data)))
            else:
                result["neto_fail"] = len(order_data)
                err_msg = resp.get("Messages", str(resp))
                result["errors"].append("Neto API error: {}".format(err_msg))
                result["details"].append("Neto notes FAIL: {}".format(err_msg))
        except Exception as e:
            result["neto_fail"] = len(order_data)
            result["errors"].append("Neto note request failed: {}".format(e))

    log("Notes done: eBay {}/{}, Neto {}/{}, Neto(eBay) {}/{}".format(
        result["ebay_ok"], result["ebay_ok"] + result["ebay_fail"],
        result["neto_ok"], result["neto_ok"] + result["neto_fail"],
        result["neto_ebay_ok"],
        result["neto_ebay_ok"] + result["neto_ebay_fail"]))

    return result
