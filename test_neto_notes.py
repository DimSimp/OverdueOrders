"""
Targeted Neto diagnostic: fetch "Pick" status orders and dump raw JSON
to find the sticky note field name. Run from the project root.

Usage:
    python test_neto_notes.py
"""
from __future__ import annotations

import json
import sys
from datetime import datetime, timedelta

from src.config import config

try:
    config.load()
    print("[OK] config.json loaded")
except Exception as e:
    print(f"[FAIL] {e}")
    sys.exit(1)

import requests

TARGET_SKUS = {
    "488110AUSTRALIS",
    "208135AUSTRALIS",
    "1200269AUSTRALIS",
    "450711AUSTRALIS",
    "450781AUSTRALIS",
}

# Wide date range — 365 days back
date_to   = datetime.now()
date_from = date_to - timedelta(days=365)
print(f"Date range: {date_from.date()} to {date_to.date()}")
print(f"Looking for SKUs: {TARGET_SKUS}")
print()

OUTPUT_SELECTOR = [
    "OrderID",
    "Username",
    "Email",
    "ShipFirstName",
    "ShipLastName",
    "DatePlaced",
    "DatePaid",
    "OrderStatus",
    "SalesChannel",
    "DeliveryInstruction",
    "StaffNote",
    "StaffNotes",
    "OrderNotes",
    "CustomerNotes",
    "OrderLine",
]

body = {
    "Filter": {
        "OrderStatus": ["Pick"],
        "DatePaidFrom": date_from.strftime("%Y-%m-%d 00:00:00"),
        "DatePaidTo":   date_to.strftime("%Y-%m-%d 23:59:59"),
        "OutputSelector": OUTPUT_SELECTOR,
        "Page": 1,
        "Limit": 200,
    }
}

url = f"{config.neto.store_url}/do/WS/NetoAPI"
resp = requests.post(
    url,
    headers={
        "NETOAPI_ACTION":   "GetOrder",
        "NETOAPI_KEY":      config.neto.api_key,
        "NETOAPI_USERNAME": config.neto.username,
        "Accept":           "application/json",
        "Content-Type":     "application/json",
    },
    json=body,
    timeout=30,
)
resp.raise_for_status()
data = resp.json()

if data.get("Ack") not in ("Success", "Warning"):
    print(f"[FAIL] Neto API error: {data.get('Messages')}")
    sys.exit(1)

orders = data.get("Order", [])
if isinstance(orders, dict):
    orders = [orders]

print(f"[OK] {len(orders)} 'Pick' status order(s) returned\n")

# --- Print ALL keys present across all orders (to find notes field name) ---
all_keys: set = set()
for o in orders:
    all_keys.update(o.keys())

print("All top-level keys seen across returned orders:")
for k in sorted(all_keys):
    print(f"  {k}")
print()

# --- Print note-related fields for every order ---
NOTE_CANDIDATES = [k for k in all_keys if any(
    word in k.lower() for word in ("note", "staff", "delivery", "instruction", "comment", "memo", "internal")
)]
print(f"Note-candidate fields: {NOTE_CANDIDATES}\n")

# --- Look for SKU matches ---
matched_orders = []
for o in orders:
    lines = o.get("OrderLine", [])
    if isinstance(lines, dict):
        lines = [lines]
    order_skus = {str(l.get("SKU") or l.get("ProductSKU") or "").strip() for l in lines}
    hit = TARGET_SKUS & order_skus
    if hit:
        matched_orders.append((o, hit))

if matched_orders:
    print(f"[OK] Found {len(matched_orders)} order(s) with target SKUs!\n")
    for o, hit in matched_orders:
        print(f"  Order {o.get('OrderID')} — matched SKUs: {hit}")
        # Print every note-candidate field
        for k in NOTE_CANDIDATES:
            v = o.get(k)
            if v:
                print(f"    {k}: {v!r}")
        print("  Full raw JSON:")
        print(json.dumps(o, indent=2, default=str))
        print()
else:
    print("[!] No target SKUs found in Pick orders. Showing note fields for first 5 orders:\n")
    for o in orders[:5]:
        oid = o.get("OrderID", "?")
        lines = o.get("OrderLine", [])
        if isinstance(lines, dict):
            lines = [lines]
        skus = [str(l.get("SKU") or "").strip() for l in lines]
        print(f"  Order {oid} | SKUs: {skus}")
        for k in NOTE_CANDIDATES:
            v = o.get(k)
            print(f"    {k}: {v!r}")
        print()

    # Also dump raw JSON of first order so we can see every field
    if orders:
        print("Full raw JSON of first order:")
        print(json.dumps(orders[0], indent=2, default=str))
