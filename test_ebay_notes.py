"""
Dump raw eBay order JSON to see all available fields and find where
internal/seller notes are stored.
Run: py test_ebay_notes.py
"""
from __future__ import annotations
import json, sys
from datetime import datetime, timedelta

from src.config import config
try:
    config.load()
except Exception as e:
    print(f"[FAIL] {e}"); sys.exit(1)

from src.ebay_client import EbayClient, EbayAuthError, EbayAPIError, _to_ebay_datetime

def _save(at, exp, rt=None):
    config.save_ebay_tokens(at, exp, rt)

ebay = EbayClient(config.ebay, _save)
if not ebay.is_authenticated():
    print("[FAIL] eBay not authenticated"); sys.exit(1)

token = ebay._ensure_valid_token()

date_to   = datetime.now()
date_from = date_to - timedelta(days=30)
from_str = _to_ebay_datetime(date_from)
to_str   = _to_ebay_datetime(date_to)

import requests
resp = requests.get(
    f"{ebay._api_base}/sell/fulfillment/v1/order",
    headers={
        "Authorization": f"Bearer {token}",
        "X-EBAY-C-MARKETPLACE-ID": "EBAY_AU",
    },
    params={
        "filter": f"creationdate:[{from_str}..{to_str}],orderfulfillmentstatus:{{NOT_STARTED|IN_PROGRESS}}",
        "limit": 5,
        "offset": 0,
    },
    timeout=30,
)
resp.raise_for_status()
data = resp.json()

orders = data.get("orders", [])
print(f"Got {len(orders)} orders (sample of 5)\n")

# Print all top-level keys across all orders
all_keys: set = set()
for o in orders:
    all_keys.update(o.keys())
print("Top-level keys on eBay order object:")
for k in sorted(all_keys):
    print(f"  {k}")
print()

# Look for any note-related keys
note_keys = [k for k in all_keys if any(w in k.lower() for w in
    ("note", "comment", "annotation", "memo", "seller", "internal", "message", "checkout"))]
print(f"Note-candidate keys: {note_keys}\n")

# Print note-candidate fields + buyerCheckoutNotes for first 3 orders
for o in orders[:3]:
    paid = o.get("orderPaymentStatus","")
    status = o.get("orderFulfillmentStatus","")
    oid = o.get("orderId","")
    skus = [li.get("sku","") or li.get("title","") for li in o.get("lineItems",[])]
    print(f"Order {oid[:30]} | paid={paid} | status={status}")
    print(f"  SKUs: {skus}")
    for k in note_keys:
        v = o.get(k)
        if v:
            print(f"  {k}: {v!r}")
    print()

# Also dump full JSON of first order for complete inspection
if orders:
    print("Full JSON of first order:")
    print(json.dumps(orders[0], indent=2, default=str))
