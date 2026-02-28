"""
Search ALL undispatched statuses (not just Pick) for any 'on po' / 'po' / 'purchase order' notes.
Also lists all non-empty InternalOrderNotes to see what wording staff actually use.
Run: py test_neto_onpo2.py
"""
from __future__ import annotations
import json, sys, requests
from src.config import config

try:
    config.load()
except Exception as e:
    print(f"[FAIL] {e}"); sys.exit(1)

ALL_STATUSES = ["New", "New Backorder", "Backorder Approved", "Pick", "Pack", "Pending Dispatch"]

def post(body):
    resp = requests.post(
        f"{config.neto.store_url}/do/WS/NetoAPI",
        headers={
            "NETOAPI_ACTION":   "GetOrder",
            "NETOAPI_KEY":      config.neto.api_key,
            "NETOAPI_USERNAME": config.neto.username,
            "Accept": "application/json", "Content-Type": "application/json",
        },
        json=body, timeout=30,
    )
    resp.raise_for_status()
    data = resp.json()
    if data.get("Ack") not in ("Success", "Warning"):
        raise Exception(f"Neto error: {data.get('Messages')}")
    return data

print(f"Scanning all statuses: {ALL_STATUSES}")
print("Looking for 'on po' AND listing all orders with any InternalOrderNotes...\n")

page = 1
total_seen = 0
on_po_orders = []
has_notes = []

while True:
    body = {
        "Filter": {
            "OrderStatus": ALL_STATUSES,
            "OutputSelector": [
                "OrderID", "Username", "DatePaid", "OrderStatus",
                "InternalOrderNotes", "DeliveryInstruction", "OrderLine",
            ],
            "Page": page,
            "Limit": 200,
        }
    }
    data = post(body)
    orders = data.get("Order", [])
    if isinstance(orders, dict):
        orders = [orders]
    if not orders:
        break

    total_seen += len(orders)
    for o in orders:
        notes = (o.get("InternalOrderNotes") or "").strip()
        delivery = (o.get("DeliveryInstruction") or "").strip()
        combined = (notes + " " + delivery).lower()
        if notes or delivery:
            has_notes.append(o)
        if "on po" in combined or "purchase order" in combined or " po " in combined:
            on_po_orders.append(o)

    print(f"  Page {page}: {len(orders)} orders | cumulative={total_seen} | with-notes={len(has_notes)} | on-PO={len(on_po_orders)}")
    if len(orders) < 200:
        break
    page += 1

print(f"\nTotal: {total_seen} orders | {len(has_notes)} with notes | {len(on_po_orders)} with 'on po'\n")

if on_po_orders:
    print("=== Orders with 'on po' ===")
    for o in on_po_orders:
        lines = o.get("OrderLine", [])
        if isinstance(lines, dict): lines = [lines]
        skus = [str(l.get("SKU") or "").strip() for l in lines]
        print(f"  {o.get('OrderID')} | {o.get('OrderStatus')} | {o.get('DatePaid')} | SKUs: {skus}")
        print(f"  Notes: {(o.get('InternalOrderNotes') or '').strip()!r}")
        print()
else:
    print("No 'on po' found. Showing ALL notes that exist:\n")
    for o in has_notes[:30]:
        lines = o.get("OrderLine", [])
        if isinstance(lines, dict): lines = [lines]
        skus = [str(l.get("SKU") or "").strip() for l in lines]
        notes = (o.get("InternalOrderNotes") or "").strip()
        print(f"  {o.get('OrderID')} | {o.get('OrderStatus')} | SKUs: {skus}")
        print(f"  Notes: {notes[:120]!r}")
        print()
    if len(has_notes) > 30:
        print(f"  ... and {len(has_notes)-30} more with notes")
