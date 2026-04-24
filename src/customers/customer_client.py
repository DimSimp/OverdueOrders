from __future__ import annotations

import logging
import threading
from typing import Optional

log = logging.getLogger(__name__)

_SEARCH_COLS = (
    "id,customer_id,customer_code,first_name,surname,business,city,mobile,phone_1,email,active"
)

_FULL_COLS = "*"

_PAGE_SIZE = 100


def search_customers(
    query: str,
    active_filter: str = "active",
    sort_col: str = "customer_id",
    sort_asc: bool = True,
    page: int = 0,
    page_size: int = _PAGE_SIZE,
) -> tuple[list, int]:
    """Search customers by name / ID / phone / email.

    Returns (rows, total_count). Returns empty lists when query is blank
    and active_filter == 'active' (search-first behaviour).
    """
    query = (query or "").strip()
    if not query and active_filter == "active":
        return [], 0

    from src.supabase_client import get_client
    db = get_client()

    # Use RPC so the ilike patterns go in the POST body rather than the URL,
    # bypassing Cloudflare's WAF which blocks or=(field.ilike.%pattern%) queries.
    result = db.rpc("search_customers_fn", {
        "p_query":    query,
        "p_active":   active_filter,
        "p_sort_col": sort_col,
        "p_sort_asc": sort_asc,
        "p_offset":   page * page_size,
        "p_limit":    page_size + 1,   # +1 to detect next page
    }).execute()

    rows = result.data or []
    has_next = len(rows) > page_size
    rows = rows[:page_size]
    total = page * page_size + len(rows) + (1 if has_next else 0)
    return rows, total


def get_customer(uuid: str) -> Optional[dict]:
    """Return full customer record by UUID, or None."""
    from src.supabase_client import get_client
    db = get_client()
    result = (
        db.table("customers")
        .select(_FULL_COLS)
        .eq("id", uuid)
        .limit(1)
        .execute()
    )
    return result.data[0] if result.data else None


def create_customer(data: dict) -> dict:
    """Create a new customer. Auto-assigns customer_id and customer_barcode."""
    from src.supabase_client import get_client
    db = get_client()

    # Assign next sequential customer_id
    max_result = (
        db.table("customers")
        .select("customer_id")
        .order("customer_id", desc=True)
        .limit(1)
        .execute()
    )
    if max_result.data:
        next_id = (max_result.data[0].get("customer_id") or 0) + 1
    else:
        next_id = 1

    payload = {**data}
    payload["customer_id"] = next_id
    payload["customer_barcode"] = str(next_id).zfill(8)
    payload.pop("id", None)  # never pass id on insert

    result = db.table("customers").insert(payload).execute()
    return result.data[0]


def update_customer(uuid: str, data: dict) -> dict:
    """Update an existing customer by UUID."""
    from src.supabase_client import get_client
    db = get_client()

    payload = {k: v for k, v in data.items() if k not in ("id", "customer_id", "customer_barcode")}
    result = (
        db.table("customers")
        .update(payload)
        .eq("id", uuid)
        .execute()
    )
    return result.data[0]


def get_customer_transactions(
    uuid: str,
    page: int = 0,
    page_size: int = 50,
) -> tuple[list, int]:
    """Return completed transactions linked to a customer, newest first.

    Returns (rows, total) where total encodes has-next-page: if there is a
    next page, total is set just above the current page end so the caller's
    (page+1)*page_size < total check enables the Next button.
    """
    from src.supabase_client import get_client
    db = get_client()

    cols = "id,transaction_number,sale_type,total,completed_at"
    # Fetch one extra row to detect whether a next page exists (avoids
    # count="exact" which fails when PostgREST returns content-range: */*).
    result = (
        db.table("transactions")
        .select(cols)
        .eq("customer_id", uuid)
        .eq("sale_status", "completed")
        .order("completed_at", desc=True)
        .range(page * page_size, page * page_size + page_size)  # +1 extra
        .execute()
    )
    rows = result.data or []
    has_next = len(rows) > page_size
    rows = rows[:page_size]
    # Encode has_next into total so existing pagination logic still works
    total = page * page_size + len(rows) + (1 if has_next else 0)
    return rows, total


def lookup_for_till(query: str) -> list:
    """Quick customer lookup for the Till tab.

    Tries an exact ``customer_barcode`` match first (numeric scan), then
    falls back to the ``search_customers_fn`` RPC for name/mobile/email.
    Returns at most 6 rows (caller detects "> 5 matches" from len == 6).
    """
    query = (query or "").strip()
    if not query:
        return []

    from src.supabase_client import get_client
    db = get_client()

    _cols = (
        "id,customer_id,customer_barcode,first_name,surname,business,"
        "mobile,phone_1,email,active,discount_profile"
    )

    # Barcode scan: all digits ≤ 10 chars → try zero-padded exact match first
    if query.isdigit() and len(query) <= 10:
        padded = query.zfill(8)
        result = (
            db.table("customers")
            .select(_cols)
            .eq("customer_barcode", padded)
            .eq("active", True)
            .limit(1)
            .execute()
        )
        if result.data:
            return result.data

    # Text search via RPC
    result = db.rpc("search_customers_fn", {
        "p_query":    query,
        "p_active":   "active",
        "p_sort_col": "customer_id",
        "p_sort_asc": True,
        "p_offset":   0,
        "p_limit":    6,
    }).execute()
    rows = (result.data or [])[:6]
    ids = [r.get("id") for r in rows if r.get("id")]
    if not ids:
        return rows

    full_rows = (
        db.table("customers")
        .select("*")
        .in_("id", ids)
        .execute()
        .data
        or []
    )
    by_id = {r["id"]: r for r in full_rows if r.get("id")}
    return [by_id.get(r["id"], r) for r in rows if r.get("id")]


def batch_upsert_customers(
    rows: list[dict],
    on_progress=None,
    stop_event: Optional[threading.Event] = None,
) -> dict:
    """Upsert customers in batches of 200, conflicting on musipos_account_code.

    Returns {done, errors, first_error}.
    """
    from src.supabase_client import get_client
    db = get_client()

    # Deduplicate by musipos_account_code so a single batch never tries to
    # upsert the same conflict key twice (PostgreSQL rejects that).
    # Prefer the row with the most name data so a name-less duplicate doesn't
    # overwrite a good record.
    def _name_score(row: dict) -> int:
        score = 0
        if row.get("first_name"): score += 2
        if row.get("surname"): score += 1
        return score

    seen_codes: dict = {}
    no_code_rows: list = []
    for row in rows:
        code = row.get("musipos_account_code")
        if code is not None:
            existing = seen_codes.get(code)
            if existing is None or _name_score(row) > _name_score(existing):
                seen_codes[code] = row
        else:
            no_code_rows.append(row)
    rows = no_code_rows + list(seen_codes.values())

    BATCH = 200
    done = 0
    errors = 0
    first_error: Optional[str] = None

    for i in range(0, len(rows), BATCH):
        if stop_event and stop_event.is_set():
            break
        batch = rows[i: i + BATCH]
        try:
            db.table("customers").upsert(
                batch, on_conflict="musipos_account_code"
            ).execute()
            done += len(batch)
        except Exception as exc:
            log.exception("batch_upsert_customers — batch starting at row %d failed", i)
            errors += len(batch)
            if first_error is None:
                first_error = str(exc)
        if on_progress:
            on_progress(done + errors, len(rows))

    return {"done": done, "errors": errors, "first_error": first_error}
