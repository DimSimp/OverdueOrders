"""
Search ALL Pick orders (no date filter, all pages) for 'on po' in InternalOrderNotes.
Run: py test_neto_onpo.py
"""
from __future__ import annotations
import json, sys, requests
from src.config import config

try:
    config.load()
except Exception as e:
    print(f"[FAIL] {e}"); sys.exit(1)

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

print("Scanning ALL Pick orders for 'on po' in InternalOrderNotes...")
page = 1
total_seen = 0
on_po_orders = []

while True:
    body = {
        "Filter": {
            "OrderStatus": ["Pick"],
            "OutputSelector": [
                "OrderID", "Username", "Email",
                "DatePlaced", "DatePaid", "OrderStatus",
                "InternalOrderNotes", "DeliveryInstruction",
                "SalesChannel", "OrderLine",
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
        notes = (o.get("InternalOrderNotes") or "") + " " + (o.get("DeliveryInstruction") or "")
        if "on po" in notes.lower():
            on_po_orders.append(o)

    print(f"  Page {page}: {len(orders)} orders | cumulative={total_seen} | on-PO found={len(on_po_orders)}")
    if len(orders) < 200:
        break
    page += 1

print(f"\nTotal: {total_seen} Pick orders scanned, {len(on_po_orders)} with 'on po' in notes.\n")

for o in on_po_orders:
    lines = o.get("OrderLine", [])
    if isinstance(lines, dict):
        lines = [lines]
    skus = [str(l.get("SKU") or "").strip() for l in lines]
    notes = (o.get("InternalOrderNotes") or "").strip()
    print(f"  Order {o.get('OrderID')} | {o.get('DatePaid')} | SKUs: {skus}")
    print(f"  Notes: {notes!r}")
    print()
