"""
Search for target SKUs with no date filter - paginate all pages.
Run: py test_neto_notes3.py
"""
from __future__ import annotations
import json, sys, requests
from src.config import config

try:
    config.load()
except Exception as e:
    print(f"[FAIL] {e}"); sys.exit(1)

TARGET_SKUS = {
    "488110AUSTRALIS",
    "208135AUSTRALIS",
    "1200269AUSTRALIS",
    "450711AUSTRALIS",
    "450781AUSTRALIS",
}

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

print("Searching ALL Pick orders (no date filter, paginating)...")
page = 1
total_seen = 0
found = []

while True:
    body = {
        "Filter": {
            "OrderStatus": ["Pick"],
            "OutputSelector": [
                "OrderID", "Username", "Email", "DatePlaced", "DatePaid",
                "OrderStatus", "InternalOrderNotes", "DeliveryInstruction",
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
        lines = o.get("OrderLine", [])
        if isinstance(lines, dict):
            lines = [lines]
        order_skus = {str(l.get("SKU") or "").strip() for l in lines}
        hit = TARGET_SKUS & order_skus
        if hit:
            found.append((o, hit))

    print(f"  Page {page}: {len(orders)} orders, cumulative={total_seen}, matches so far={len(found)}")
    if len(orders) < 200:
        break
    page += 1

print(f"\nDone. {total_seen} total Pick orders scanned, {len(found)} with target SKUs.\n")
for o, hit in found:
    notes = o.get("InternalOrderNotes", "") or o.get("DeliveryInstruction", "") or ""
    print(f"  Order {o.get('OrderID')} | {o.get('DatePaid')} | match={hit}")
    print(f"  InternalOrderNotes: {notes!r}")
    print(json.dumps(o, indent=2, default=str))
    print()
