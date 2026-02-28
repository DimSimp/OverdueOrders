"""
Verify StickyNotes field is returned and check all Pick orders for 'on po'.
Run: py test_neto_sticky.py
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

# Check one known order first to see if StickyNotes appears
print("Checking order N476399 for StickyNotes field...")
data = post({"Filter": {
    "OrderID": ["N476399"],
    "OutputSelector": ["OrderID", "StickyNotes", "InternalOrderNotes", "DeliveryInstruction"],
}})
orders = data.get("Order", [])
if isinstance(orders, dict): orders = [orders]
if orders:
    print(f"Fields returned: {sorted(orders[0].keys())}")
    print(json.dumps(orders[0], indent=2, default=str))
print()

# Scan all Pick orders for 'on po' in any notes field
print("Scanning ALL Pick orders for 'on po' in StickyNotes/InternalOrderNotes...")
page = 1
total = 0
on_po = []
has_sticky = []

while True:
    body = {"Filter": {
        "OrderStatus": ["Pick"],
        "OutputSelector": [
            "OrderID", "Username", "DatePaid", "OrderStatus",
            "StickyNotes", "InternalOrderNotes", "DeliveryInstruction", "OrderLine",
        ],
        "Page": page, "Limit": 200,
    }}
    data = post(body)
    orders = data.get("Order", [])
    if isinstance(orders, dict): orders = [orders]
    if not orders: break

    total += len(orders)
    for o in orders:
        sn = o.get("StickyNotes")
        if isinstance(sn, dict):
            sn = [sn]
        if isinstance(sn, list):
            sticky = " ".join(
                (str(n.get("Title") or "") + " " + str(n.get("Description") or "")).strip()
                for n in sn
            ).strip()
        else:
            sticky = (sn or "").strip()
        internal = (o.get("InternalOrderNotes") or "").strip()
        delivery = (o.get("DeliveryInstruction") or "").strip()
        if sticky:
            has_sticky.append(o)
        combined = (sticky + " " + internal + " " + delivery).lower()
        if "on po" in combined:
            on_po.append(o)

    print(f"  Page {page}: {len(orders)} orders | total={total} | with-sticky={len(has_sticky)} | on-po={len(on_po)}")
    if len(orders) < 200: break
    page += 1

print(f"\nResult: {total} Pick orders | {len(has_sticky)} with StickyNotes | {len(on_po)} with 'on po'\n")

if on_po:
    print("=== Orders with 'on po' ===")
    for o in on_po:
        lines = o.get("OrderLine", [])
        if isinstance(lines, dict): lines = [lines]
        skus = [l.get("SKU","") for l in lines]
        print(f"  {o.get('OrderID')} | {o.get('DatePaid')} | SKUs: {skus}")
        print(f"  StickyNotes: {(o.get('StickyNotes') or '').strip()!r}")
        print(f"  InternalOrderNotes: {(o.get('InternalOrderNotes') or '').strip()[:80]!r}")
        print()

if has_sticky and not on_po:
    print("Orders with StickyNotes (first 10):")
    for o in has_sticky[:10]:
        lines = o.get("OrderLine", [])
        if isinstance(lines, dict): lines = [lines]
        skus = [l.get("SKU","") for l in lines]
        print(f"  {o.get('OrderID')} | SKUs: {skus}")
        print(f"  Sticky: {(o.get('StickyNotes') or '').strip()[:120]!r}")
        print()
