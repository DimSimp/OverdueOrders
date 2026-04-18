from __future__ import annotations

from typing import Optional

BATCH_SIZE = 500

# Columns fetched for the search grid (lighter than SELECT *)
_GRID_COLS = (
    "id, sku, web_sku, title, brand, instrument, sub_instrument, supplier_id, "
    "supplier_rrp, online_sale_price, qty_on_hand, qty_allocated_online, "
    "qty_allocated_customer, qty_on_order, active, pick_zone, minimum_sell, is_serialised, "
    "average_cost_exc_gst, last_purchase_cost"
)

_SORT_WHITELIST = {
    "sku", "title", "brand", "supplier_rrp", "online_sale_price",
    "qty_on_hand", "qty_on_order",
}


def search_items(
    query: str = "",
    filters: Optional[dict] = None,
    sort_col: str = "title",
    sort_asc: bool = True,
    page: int = 0,
    page_size: int = 100,
) -> tuple[list[dict], int]:
    """Search items. Returns (rows, total_count).

    Returns empty results if both query and filters are empty/default — the
    inventory tab shows a prompt rather than loading the entire catalogue.
    """
    from src.supabase_client import get_client
    db = get_client()
    filters = filters or {}

    has_query = len(query.strip()) >= 2
    active_filter = filters.get("active", "active")
    has_filter = any([
        filters.get("brand"),
        filters.get("instrument"),
        filters.get("supplier_id"),
        active_filter != "active",
        filters.get("stock") not in (None, "all"),
    ])

    if not has_query and not has_filter:
        return [], 0

    q = db.table("items").select(_GRID_COLS, count="exact")

    if has_query:
        # Sanitise to avoid breaking the OR filter string
        safe = query.replace(",", " ").replace(".", " ").strip()
        q = q.or_(
            f"sku.ilike.%{safe}%,"
            f"title.ilike.%{safe}%,"
            f"brand.ilike.%{safe}%,"
            f"internal_barcode.ilike.%{safe}%"
        )

    if active_filter == "active":
        q = q.eq("active", True)
    elif active_filter == "inactive":
        q = q.eq("active", False)

    if filters.get("brand"):
        q = q.eq("brand", filters["brand"])
    if filters.get("instrument"):
        q = q.eq("instrument", filters["instrument"])
    if filters.get("supplier_id"):
        q = q.eq("supplier_id", filters["supplier_id"])

    stock_filter = filters.get("stock", "all")
    if stock_filter == "in_stock":
        q = q.gt("qty_on_hand", 0)
    elif stock_filter == "out_of_stock":
        q = q.eq("qty_on_hand", 0)

    col = sort_col if sort_col in _SORT_WHITELIST else "title"
    q = q.order(col, desc=not sort_asc)

    offset = page * page_size
    q = q.range(offset, offset + page_size - 1)

    result = q.execute()
    total = result.count if result.count is not None else len(result.data)
    return result.data, total


def get_item_by_id(item_id: str) -> Optional[dict]:
    """Fetch full item record by UUID."""
    from src.supabase_client import get_client
    result = get_client().table("items").select("*").eq("id", item_id).limit(1).execute()
    return result.data[0] if result.data else None


def get_item_by_sku(sku: str) -> Optional[dict]:
    """Fetch item by SKU, web SKU, internal barcode, or product barcode (POS lookup)."""
    from src.supabase_client import get_client
    db = get_client()
    for col in ("sku", "web_sku", "internal_barcode", "product_barcode"):
        result = db.table("items").select("*").eq(col, sku).limit(1).execute()
        if result.data:
            return result.data[0]
    return None


def batch_upsert_items(
    rows: list[dict],
    on_progress=None,
    stop_event=None,
) -> dict:
    """Upsert items in batches of BATCH_SIZE. Returns {done, errors, first_error}."""
    from src.supabase_client import get_client
    db = get_client()
    total = len(rows)
    done = 0
    error_count = 0
    first_error: str = ""

    for i in range(0, total, BATCH_SIZE):
        if stop_event and stop_event.is_set():
            break
        batch = rows[i: i + BATCH_SIZE]
        try:
            db.table("items").upsert(batch, on_conflict="sku").execute()
        except Exception as exc:
            error_count += len(batch)
            if not first_error:
                first_error = str(exc)
            batch_num = i // BATCH_SIZE + 1
            print(f"[IMPORT] Batch {batch_num} failed "
                  f"(rows {i}–{i + len(batch) - 1}): {exc}")
            # Print the first SKU in the failing batch to help identify the problem
            if batch:
                print(f"[IMPORT]   First SKU in batch: {batch[0].get('sku')} "
                      f"| title: {batch[0].get('title')!r}")
        done += len(batch)
        if on_progress:
            on_progress(done, total)

    return {"done": done, "errors": error_count, "first_error": first_error}


def ensure_suppliers(supplier_ids: list[str]) -> None:
    """Create stub supplier records for any IDs not already in the suppliers table."""
    from src.supabase_client import get_client
    db = get_client()
    existing = {r["id"] for r in db.table("suppliers").select("id").execute().data}
    new_stubs = [
        {"id": sid, "name": sid}
        for sid in supplier_ids
        if sid and sid not in existing
    ]
    if new_stubs:
        db.table("suppliers").insert(new_stubs).execute()


def lookup_exact(query: str) -> list[dict]:
    """Exact-match lookup across all barcode/SKU columns for Till entry.

    Checks columns in priority order: sku → web_sku → internal_barcode → product_barcode.
    Returns the results from the FIRST column that finds a match, so that a valid SKU
    entry is never contaminated by a coincidental barcode match on a different item.

    Uses ilike (case-insensitive) so manually typed SKUs match regardless of casing.
    ilike wildcards (% and _) are escaped so the match remains exact.
    """
    from src.supabase_client import get_client
    db = get_client()
    _COLS = (
        "id, sku, title, online_sale_price, supplier_rrp, minimum_sell, "
        "qty_on_hand, qty_allocated_online, qty_allocated_customer, "
        "is_serialised, active, average_cost_exc_gst, last_purchase_cost"
    )
    # Escape ilike special chars so the match is always exact (just case-insensitive)
    safe = query.replace("\\", "\\\\").replace("%", "\\%").replace("_", "\\_")
    for col in ("sku", "web_sku", "internal_barcode", "product_barcode"):
        rows = db.table("items").select(_COLS).ilike(col, safe).eq("active", True).execute().data
        if rows:
            return rows
    return []


def get_distinct_col(col: str, active_only: bool = True) -> list[str]:
    """Return sorted unique non-empty values for a single column (for filter dropdowns)."""
    from src.supabase_client import get_client
    db = get_client()
    q = db.table("items").select(col)
    if active_only:
        q = q.eq("active", True)
    result = q.execute()
    return sorted({r[col] for r in result.data if r.get(col)})
