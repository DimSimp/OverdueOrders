from __future__ import annotations

import json
import os
from dataclasses import asdict, dataclass, field
from datetime import datetime
from pathlib import Path

from src.data_processor import MatchedOrder
from src.ebay_client import EbayLineItem, EbayOrder
from src.neto_client import NetoLineItem, NetoOrder
from src.pdf_parser import InvoiceItem


SNAPSHOT_VERSION = 1


@dataclass
class SessionSnapshot:
    invoice_items: list[InvoiceItem]
    neto_orders: list[NetoOrder]
    ebay_orders: list[EbayOrder]
    matched_orders: list[MatchedOrder]
    unmatched_inv: list[InvoiceItem]
    excluded_order_ids: list[tuple[str, str]]
    force_matched_order_ids: list[tuple[str, str]]


SESSION_FILENAME = "Incoming_orders_session.scar"


def save_snapshot(
    save_dir: str,
    invoice_items: list[InvoiceItem],
    neto_orders: list[NetoOrder],
    ebay_orders: list[EbayOrder],
    matched_orders: list[MatchedOrder],
    unmatched_inv: list[InvoiceItem],
    excluded_ids: set[tuple[str, str]],
    force_matched_ids: set[tuple[str, str]],
) -> str:
    """Save a session snapshot to a .scar file (overwrites previous). Returns the file path."""
    os.makedirs(save_dir, exist_ok=True)

    filepath = os.path.join(save_dir, SESSION_FILENAME)

    data = {
        "version": SNAPSHOT_VERSION,
        "timestamp": datetime.now().isoformat(),
        "invoice_items": [asdict(i) for i in invoice_items],
        "neto_orders": [_serialize_neto_order(o) for o in neto_orders],
        "ebay_orders": [_serialize_ebay_order(o) for o in ebay_orders],
        "matched_orders": [_serialize_matched(m) for m in matched_orders],
        "unmatched_inv": [asdict(i) for i in unmatched_inv],
        "excluded_order_ids": list(excluded_ids),
        "force_matched_order_ids": list(force_matched_ids),
    }

    with open(filepath, "w", encoding="utf-8") as f:
        json.dump(data, f, indent=2, ensure_ascii=False, default=str)

    return filepath


def load_snapshot(path: str) -> SessionSnapshot:
    """Load a session snapshot from a JSON file."""
    with open(path, "r", encoding="utf-8") as f:
        data = json.load(f)

    return SessionSnapshot(
        invoice_items=[_parse_invoice_item(d) for d in data.get("invoice_items", [])],
        neto_orders=[_parse_neto_order(d) for d in data.get("neto_orders", [])],
        ebay_orders=[_parse_ebay_order(d) for d in data.get("ebay_orders", [])],
        matched_orders=[_parse_matched(d) for d in data.get("matched_orders", [])],
        unmatched_inv=[_parse_invoice_item(d) for d in data.get("unmatched_inv", [])],
        excluded_order_ids=[tuple(x) for x in data.get("excluded_order_ids", [])],
        force_matched_order_ids=[tuple(x) for x in data.get("force_matched_order_ids", [])],
    )


# ── Serialization helpers ─────────────────────────────────────────────

def _serialize_neto_order(o: NetoOrder) -> dict:
    d = asdict(o)
    d["date_placed"] = o.date_placed.isoformat() if o.date_placed else None
    d["date_paid"] = o.date_paid.isoformat() if o.date_paid else None
    return d


def _serialize_ebay_order(o: EbayOrder) -> dict:
    d = asdict(o)
    d["creation_date"] = o.creation_date.isoformat() if o.creation_date else None
    return d


def _serialize_matched(m: MatchedOrder) -> dict:
    d = asdict(m)
    d["order_date"] = m.order_date.isoformat() if m.order_date else None
    return d


# ── Deserialization helpers ───────────────────────────────────────────

def _parse_date(val) -> datetime | None:
    if not val:
        return None
    if isinstance(val, datetime):
        return val
    try:
        return datetime.fromisoformat(str(val))
    except (ValueError, TypeError):
        return None


def _parse_invoice_item(d: dict) -> InvoiceItem:
    return InvoiceItem(
        sku=d.get("sku", ""),
        sku_with_suffix=d.get("sku_with_suffix", ""),
        description=d.get("description", ""),
        quantity=int(d.get("quantity", 0)),
        source_page=int(d.get("source_page", 0)),
        qty_flagged=bool(d.get("qty_flagged", False)),
    )


def _parse_neto_order(d: dict) -> NetoOrder:
    line_items = [
        NetoLineItem(
            sku=li.get("sku", ""),
            product_name=li.get("product_name", ""),
            quantity=int(li.get("quantity", 0)),
            unit_price=float(li.get("unit_price", 0)),
            image_url=li.get("image_url", ""),
        )
        for li in d.get("line_items", [])
    ]
    return NetoOrder(
        order_id=d.get("order_id", ""),
        customer_name=d.get("customer_name", ""),
        email=d.get("email", ""),
        date_placed=_parse_date(d.get("date_placed")),
        date_paid=_parse_date(d.get("date_paid")),
        status=d.get("status", ""),
        notes=d.get("notes", ""),
        sales_channel=d.get("sales_channel", ""),
        purchase_order_number=d.get("purchase_order_number", ""),
        line_items=line_items,
        sticky_notes=d.get("sticky_notes", []),
        internal_notes=d.get("internal_notes", ""),
        delivery_instruction=d.get("delivery_instruction", ""),
        ship_first_name=d.get("ship_first_name", ""),
        ship_last_name=d.get("ship_last_name", ""),
        ship_company=d.get("ship_company", ""),
        ship_street1=d.get("ship_street1", ""),
        ship_street2=d.get("ship_street2", ""),
        ship_city=d.get("ship_city", ""),
        ship_state=d.get("ship_state", ""),
        ship_postcode=d.get("ship_postcode", ""),
        ship_country=d.get("ship_country", ""),
        ship_phone=d.get("ship_phone", ""),
        grand_total=float(d.get("grand_total", 0)),
        shipping_total=float(d.get("shipping_total", 0)),
        shipping_method=d.get("shipping_method", ""),
        shipping_type=d.get("shipping_type", ""),
    )


def _parse_ebay_order(d: dict) -> EbayOrder:
    line_items = [
        EbayLineItem(
            line_item_id=li.get("line_item_id", ""),
            sku=li.get("sku", ""),
            title=li.get("title", ""),
            quantity=int(li.get("quantity", 0)),
            legacy_item_id=li.get("legacy_item_id", ""),
            legacy_transaction_id=li.get("legacy_transaction_id", ""),
            notes=li.get("notes", ""),
            image_url=li.get("image_url", ""),
            unit_price=float(li.get("unit_price", 0)),
        )
        for li in d.get("line_items", [])
    ]
    return EbayOrder(
        order_id=d.get("order_id", ""),
        buyer_name=d.get("buyer_name", ""),
        buyer_notes=d.get("buyer_notes", ""),
        creation_date=_parse_date(d.get("creation_date")),
        order_status=d.get("order_status", ""),
        payment_status=d.get("payment_status", ""),
        line_items=line_items,
        ship_name=d.get("ship_name", ""),
        ship_street1=d.get("ship_street1", ""),
        ship_street2=d.get("ship_street2", ""),
        ship_city=d.get("ship_city", ""),
        ship_state=d.get("ship_state", ""),
        ship_postcode=d.get("ship_postcode", ""),
        ship_country=d.get("ship_country", ""),
        ship_phone=d.get("ship_phone", ""),
        order_total=float(d.get("order_total", 0)),
        shipping_cost=float(d.get("shipping_cost", 0)),
        shipping_method=d.get("shipping_method", ""),
        shipping_type=d.get("shipping_type", ""),
    )


def _parse_matched(d: dict) -> MatchedOrder:
    return MatchedOrder(
        platform=d.get("platform", ""),
        order_id=d.get("order_id", ""),
        customer_name=d.get("customer_name", ""),
        order_date=_parse_date(d.get("order_date")),
        sku=d.get("sku", ""),
        description=d.get("description", ""),
        quantity=int(d.get("quantity", 0)),
        notes=d.get("notes", ""),
        shipping_type=d.get("shipping_type", ""),
        invoice_sku=d.get("invoice_sku", ""),
        invoice_description=d.get("invoice_description", ""),
        invoice_qty=int(d.get("invoice_qty", 0)),
        is_invoice_match=bool(d.get("is_invoice_match", False)),
    )
