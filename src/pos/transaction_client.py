from __future__ import annotations

from datetime import datetime, timezone
from typing import Optional


def confirm_standard_sale(
    cart_items: dict,        # item_id (str) → {sku, title, qty, unit_price, disc_pct, cost_price}
    subtotal: float,         # sum of per-line totals before cart-level discount
    cart_disc_pct: float,    # cart-level discount percentage (0 if none)
    total: float,            # final charged amount
    payment_cash: float,     # cash applied to sale (= cash_required, not tendered amount)
    payment_eft: list[dict], # [{"amount": X.XX}, ...]
    payment_online: float,   # pre-paid online amount
    cash_tendered: float,    # raw amount physically handed over
    change_given: float,     # cash_tendered − payment_cash
    performed_by: str,            # staff username for audit trail
    sale_type: str = "standard",  # "standard" or "refund"
    source_tx_id: Optional[str] = None,  # for linked refunds: ID of the original sale
) -> dict:
    """Write a completed Standard or Refund sale to Supabase.

    Inserts:
      - one ``transactions`` row  (trigger auto-assigns ``transaction_number``)
      - N ``transaction_lines`` rows
      - one ``stock_movements`` row per line (``sale_instore`` or ``return_instore``)

    For linked refunds, also marks the source transaction ``is_refunded = True``
    so it cannot be refunded a second time.

    Atomically adjusts ``items.qty_on_hand`` for each line via the
    ``adjust_item_qty`` RPC function defined in migration 02.

    Returns the inserted transaction record (includes ``transaction_number``).
    Raises on any Supabase error — the caller must NOT clear the cart on error.
    """
    from src.supabase_client import get_client
    db = get_client()

    now = datetime.now(timezone.utc).isoformat()

    cart_disc_total: Optional[float] = (
        round(subtotal * cart_disc_pct / 100, 2) if cart_disc_pct else None
    )
    total_cost = sum(
        line["qty"] * (line.get("cost_price") or 0)
        for line in cart_items.values()
    )

    movement_type = "return_instore" if sale_type.lower() == "refund" else "sale_instore"

    tx_payload: dict = {
        "sale_type":           sale_type.lower(),
        "sale_status":         "completed",
        "subtotal":            round(subtotal, 2),
        "cart_discount_pct":   round(cart_disc_pct, 2) if cart_disc_pct else None,
        "cart_discount_total": cart_disc_total,
        "total":               round(total, 2),
        "total_cost":          round(total_cost, 2) if total_cost else None,
        "payment_cash":        round(payment_cash, 2) if payment_cash else None,
        "payment_eft":         payment_eft if payment_eft else None,
        "payment_online":      round(payment_online, 2) if payment_online else None,
        "cash_tendered":       round(cash_tendered, 2) if cash_tendered else None,
        "change_given":        round(change_given, 2) if change_given else None,
        "completed_at":        now,
        "performed_by":        performed_by or None,
        "source_tx_id":        source_tx_id or None,
    }

    # Insert transaction — trigger auto-assigns transaction_number
    tx_result = db.table("transactions").insert(tx_payload).execute()
    tx = tx_result.data[0]
    tx_id: str = tx["id"]
    tx_number: str = tx["transaction_number"]

    # Build transaction lines
    lines = []
    for item_id, line in cart_items.items():
        unit_price: float = line["unit_price"]
        disc_pct: float = line["disc_pct"]
        qty: float = line["qty"]
        line_total = round(qty * unit_price * (1 - disc_pct / 100), 2)
        cost: Optional[float] = line.get("cost_price")

        margin_pct: Optional[float] = None
        if cost and cost > 0 and unit_price > 0:
            sell_ex_gst = unit_price / 1.1          # prices stored inc-GST
            if sell_ex_gst > 0:
                margin_pct = round((sell_ex_gst - cost) / sell_ex_gst * 100, 2)

        lines.append({
            "transaction_id": tx_id,
            "item_id":        item_id,
            "sku":            line["sku"],
            "description":    line["title"],
            "qty":            qty,
            "unit_price":     round(unit_price, 2),
            "cost_price":     round(cost, 2) if cost else None,
            "discount_pct":   round(disc_pct, 2) if disc_pct else None,
            "line_total":     line_total,
            "line_margin_pct": margin_pct,
        })

    db.table("transaction_lines").insert(lines).execute()

    # Atomically adjust qty_on_hand and write audit records
    for item_id, line in cart_items.items():
        qty_int = int(line["qty"])

        db.rpc("adjust_item_qty", {
            "p_item_id":    item_id,
            "p_qty_change": -qty_int,
        }).execute()

        db.table("stock_movements").insert({
            "item_id":       item_id,
            "movement_type": movement_type,
            "qty_change":    -qty_int,
            "reference_id":  tx_number,
            "performed_by":  performed_by,
        }).execute()

    # Mark the original sale as refunded so it cannot be refunded again
    if source_tx_id and sale_type.lower() == "refund":
        db.table("transactions").update({"is_refunded": True}).eq("id", source_tx_id).execute()

    return tx


def park_transaction(
    cart_items: dict,
    subtotal: float,
    cart_disc_pct: float,
    total: float,
    sale_type: str,
    customer_name: str,
    performed_by: str,
) -> dict:
    """Snapshot the current cart as a parked transaction row.

    No transaction_lines, no stock movements.  The trigger auto-assigns
    ``transaction_number``.  Returns the inserted transaction record.
    """
    from src.supabase_client import get_client
    db = get_client()

    cart_disc_total: Optional[float] = (
        round(subtotal * cart_disc_pct / 100, 2) if cart_disc_pct else None
    )
    total_cost = sum(
        line["qty"] * (line.get("cost_price") or 0)
        for line in cart_items.values()
    )

    snapshot = {
        "cart_items":    cart_items,
        "cart_disc_pct": cart_disc_pct,
        "customer_name": customer_name,
        "sale_type":     sale_type,
    }

    tx_payload = {
        "sale_type":           sale_type.lower(),
        "sale_status":         "parked",
        "park_name":           customer_name or None,
        "subtotal":            round(subtotal, 2),
        "cart_discount_pct":   round(cart_disc_pct, 2) if cart_disc_pct else None,
        "cart_discount_total": cart_disc_total,
        "total":               round(total, 2),
        "total_cost":          round(total_cost, 2) if total_cost else None,
        "cart_snapshot":       snapshot,
        "performed_by":        performed_by or None,
    }

    result = db.table("transactions").insert(tx_payload).execute()
    return result.data[0]


def complete_parked_sale(
    parked_tx_id: str,
    cart_items: dict,
    subtotal: float,
    cart_disc_pct: float,
    total: float,
    payment_cash: float,
    payment_eft: list,
    payment_online: float,
    cash_tendered: float,
    change_given: float,
    performed_by: str,
) -> dict:
    """Complete a previously-parked transaction.

    UPDATEs the existing ``transactions`` row (preserving its
    ``transaction_number``), then inserts ``transaction_lines`` and
    ``stock_movements`` exactly as ``confirm_standard_sale`` would.

    Returns the updated transaction record.
    """
    from src.supabase_client import get_client
    db = get_client()

    now = datetime.now(timezone.utc).isoformat()

    cart_disc_total: Optional[float] = (
        round(subtotal * cart_disc_pct / 100, 2) if cart_disc_pct else None
    )
    total_cost = sum(
        line["qty"] * (line.get("cost_price") or 0)
        for line in cart_items.values()
    )

    update_payload: dict = {
        "sale_status":         "completed",
        "completed_at":        now,
        "subtotal":            round(subtotal, 2),
        "cart_discount_pct":   round(cart_disc_pct, 2) if cart_disc_pct else None,
        "cart_discount_total": cart_disc_total,
        "total":               round(total, 2),
        "total_cost":          round(total_cost, 2) if total_cost else None,
        "payment_cash":        round(payment_cash, 2) if payment_cash else None,
        "payment_eft":         payment_eft if payment_eft else None,
        "payment_online":      round(payment_online, 2) if payment_online else None,
        "cash_tendered":       round(cash_tendered, 2) if cash_tendered else None,
        "change_given":        round(change_given, 2) if change_given else None,
        "performed_by":        performed_by or None,
    }

    update_result = (
        db.table("transactions")
        .update(update_payload)
        .eq("id", parked_tx_id)
        .execute()
    )
    tx = update_result.data[0]
    tx_id: str = tx["id"]
    tx_number: str = tx["transaction_number"]

    # Insert transaction lines
    lines = []
    for item_id, line in cart_items.items():
        unit_price: float = line["unit_price"]
        disc_pct: float = line["disc_pct"]
        qty: float = line["qty"]
        line_total = round(qty * unit_price * (1 - disc_pct / 100), 2)
        cost: Optional[float] = line.get("cost_price")

        margin_pct: Optional[float] = None
        if cost and cost > 0 and unit_price > 0:
            sell_ex_gst = unit_price / 1.1
            if sell_ex_gst > 0:
                margin_pct = round((sell_ex_gst - cost) / sell_ex_gst * 100, 2)

        lines.append({
            "transaction_id":  tx_id,
            "item_id":         item_id,
            "sku":             line["sku"],
            "description":     line["title"],
            "qty":             qty,
            "unit_price":      round(unit_price, 2),
            "cost_price":      round(cost, 2) if cost else None,
            "discount_pct":    round(disc_pct, 2) if disc_pct else None,
            "line_total":      line_total,
            "line_margin_pct": margin_pct,
        })

    db.table("transaction_lines").insert(lines).execute()

    # Decrement qty_on_hand and write audit records
    for item_id, line in cart_items.items():
        qty_int = int(line["qty"])
        db.rpc("adjust_item_qty", {
            "p_item_id":    item_id,
            "p_qty_change": -qty_int,
        }).execute()
        db.table("stock_movements").insert({
            "item_id":       item_id,
            "movement_type": "sale_instore",
            "qty_change":    -qty_int,
            "reference_id":  tx_number,
            "performed_by":  performed_by,
        }).execute()

    return tx


def delete_parked_transaction(tx_id: str) -> None:
    """Permanently delete a parked transaction by ID."""
    from src.supabase_client import get_client
    db = get_client()
    db.table("transactions").delete().eq("id", tx_id).execute()


def get_daily_transactions(date_str: str = None) -> list:
    """Return all completed transactions for a Melbourne calendar date.

    Parameters
    ----------
    date_str : str, optional
        Date in ``"YYYY-MM-DD"`` format (Melbourne local date).
        Defaults to today in Melbourne time.

    Each returned dict includes a nested ``"transaction_lines"`` list.
    """
    from datetime import date as _date
    from zoneinfo import ZoneInfo
    from src.supabase_client import get_client

    _MELB = ZoneInfo("Australia/Melbourne")

    if date_str is None:
        today_melb = datetime.now(_MELB).date()
    else:
        today_melb = _date.fromisoformat(date_str)

    day_start_utc = (
        datetime(today_melb.year, today_melb.month, today_melb.day,
                 0, 0, 0, tzinfo=_MELB)
        .astimezone(timezone.utc)
        .isoformat()
    )
    day_end_utc = (
        datetime(today_melb.year, today_melb.month, today_melb.day,
                 23, 59, 59, 999999, tzinfo=_MELB)
        .astimezone(timezone.utc)
        .isoformat()
    )

    db = get_client()
    result = (
        db.table("transactions")
        .select("*, transaction_lines(*)")
        .eq("sale_status", "completed")
        .gte("completed_at", day_start_utc)
        .lte("completed_at", day_end_utc)
        .order("completed_at", desc=False)
        .execute()
    )
    return result.data


def get_transaction_by_number(tx_number: str) -> Optional[dict]:
    """Return a completed transaction matching tx_number (with lines), or None."""
    from src.supabase_client import get_client
    db = get_client()
    result = (
        db.table("transactions")
        .select("*, transaction_lines(*)")
        .eq("sale_status", "completed")
        .eq("transaction_number", tx_number.upper().strip())
        .limit(1)
        .execute()
    )
    if result.data:
        return result.data[0]
    return None


def get_parked_transactions() -> list:
    """Return all parked transactions, newest first."""
    from src.supabase_client import get_client
    db = get_client()
    result = (
        db.table("transactions")
        .select("id, transaction_number, park_name, total, created_at, sale_type, cart_snapshot")
        .eq("sale_status", "parked")
        .order("created_at", desc=True)
        .execute()
    )
    return result.data
