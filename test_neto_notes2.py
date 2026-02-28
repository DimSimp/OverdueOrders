"""
Second Neto diagnostic:
1. Search ALL undispatched statuses for the target SKUs (not just Pick)
2. Try every plausible note field name in OutputSelector
3. Dump full raw JSON of any matching order

Run: py test_neto_notes2.py
"""
from __future__ import annotations

import json
import sys
from datetime import datetime, timedelta

from src.config import config

try:
    config.load()
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

ALL_UNDISPATCHED = [
    "New",
    "New Backorder",
    "Backorder Approved",
    "Pick",
    "Pack",
    "Pending Dispatch",
]

# Wide date range
date_to   = datetime.now()
date_from = date_to - timedelta(days=365)

# Try as many plausible note field names as possible
WIDE_OUTPUT_SELECTOR = [
    "OrderID", "Username", "Email",
    "ShipFirstName", "ShipLastName",
    "BillingFirstName", "BillingLastName",
    "DatePlaced", "DatePaid", "OrderStatus", "SalesChannel",
    # Note field candidates
    "DeliveryInstruction",
    "StaffNote",
    "StaffNotes",
    "OrderNotes",
    "CustomerNotes",
    "InternalNotes",
    "Note",
    "Notes",
    "GiftMessage",
    "SpecialInstructions",
    "SpecialInstruction",
    "InternalOrderNotes",
    "OrderComments",
    "Comments",
    "PrivateNote",
    "PrivateNotes",
    # OrderLine sub-fields
    "OrderLine",
]

def post(body):
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
        raise Exception(f"Neto error: {data.get('Messages')}")
    return data

# --- Pass 1: Search all undispatched statuses ---
print("Pass 1: Searching ALL undispatched statuses for target SKUs...")
body = {
    "Filter": {
        "OrderStatus": ALL_UNDISPATCHED,
        "DatePaidFrom": date_from.strftime("%Y-%m-%d 00:00:00"),
        "DatePaidTo":   date_to.strftime("%Y-%m-%d 23:59:59"),
        "OutputSelector": WIDE_OUTPUT_SELECTOR,
        "Page": 1,
        "Limit": 200,
    }
}

data = post(body)
orders = data.get("Order", [])
if isinstance(orders, dict):
    orders = [orders]

print(f"  {len(orders)} order(s) returned\n")

all_keys: set = set()
for o in orders:
    all_keys.update(o.keys())

print("All top-level keys seen:")
for k in sorted(all_keys):
    print(f"  {k}")
print()

# Find target SKU matches
found_any = False
for o in orders:
    lines = o.get("OrderLine", [])
    if isinstance(lines, dict):
        lines = [lines]
    order_skus = {str(l.get("SKU") or "").strip() for l in lines}
    hit = TARGET_SKUS & order_skus
    if hit:
        found_any = True
        print(f"[MATCH] Order {o.get('OrderID')} status={o.get('OrderStatus')} SKUs={hit}")
        print(json.dumps(o, indent=2, default=str))
        print()

if not found_any:
    print("No target SKUs found in any undispatched status.\n")
    print("First 3 orders (full raw JSON) to see all available fields:\n")
    for o in orders[:3]:
        print(json.dumps(o, indent=2, default=str))
        print()

# --- Pass 2: Search without status or date filter (just OrderID lookup for N476399) ---
print("\nPass 2: Fetch order N476399 directly with wide OutputSelector...")
body2 = {
    "Filter": {
        "OrderID": ["N476399"],
        "OutputSelector": WIDE_OUTPUT_SELECTOR,
    }
}
data2 = post(body2)
orders2 = data2.get("Order", [])
if isinstance(orders2, dict):
    orders2 = [orders2]
if orders2:
    print(f"Raw JSON for N476399:")
    print(json.dumps(orders2[0], indent=2, default=str))
else:
    print("No result for N476399")
