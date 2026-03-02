"""
Find the Neto API field name for "Purchase Order #" on eBay-channel orders.
Run: py test_neto_po_field.py
"""
from __future__ import annotations
import json, sys, requests
from src.config import config

try:
    config.load()
except Exception as e:
    print(f"[FAIL] {e}"); sys.exit(1)

CANDIDATES = [
    "OrderID", "Username", "SalesChannel", "DatePaid",
    "StickyNotes", "InternalOrderNotes",
    # PO # candidates
    "PurchaseOrderNumber", "ExternalOrderID", "ExternalID",
    "OrderReference", "PONumber", "PurchaseOrder",
    "ReferenceNumber", "Reference", "SalesOrderNumber",
    "ExternalSource", "ExternalSourceID", "ExternalSourceOrderID",
    "OrderLine",
]

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
    d = resp.json()
    if d.get("Ack") not in ("Success", "Warning"):
        raise Exception(f"{d.get('Messages')}")
    return d

# Fetch a sample of eBay-channel Pick orders
data = post({"Filter": {
    "OrderStatus": ["Pick"],
    "OutputSelector": CANDIDATES,
    "Page": 1, "Limit": 20,
}})
orders = data.get("Order", [])
if isinstance(orders, dict): orders = [orders]

# Find eBay-channel orders
ebay_orders = [o for o in orders if (o.get("SalesChannel") or "").lower() == "ebay"]
print(f"Fetched {len(orders)} Pick orders, {len(ebay_orders)} with SalesChannel=eBay\n")

if not ebay_orders:
    print("No eBay-channel orders in first 20. Expanding search...")
    data2 = post({"Filter": {
        "OrderStatus": ["Pick"],
        "OutputSelector": CANDIDATES,
        "Page": 2, "Limit": 200,
    }})
    orders2 = data2.get("Order", [])
    if isinstance(orders2, dict): orders2 = [orders2]
    ebay_orders = [o for o in orders2 if (o.get("SalesChannel") or "").lower() == "ebay"]
    print(f"Found {len(ebay_orders)} eBay-channel orders in page 2\n")

# Show all returned keys and values for the first few eBay orders
all_keys: set = set()
for o in ebay_orders:
    all_keys.update(o.keys())

print("All keys returned for eBay-channel orders:")
for k in sorted(all_keys):
    print(f"  {k}")
print()

print("First 3 eBay orders — all fields:")
for o in ebay_orders[:3]:
    print(json.dumps({k: v for k, v in o.items() if k != "OrderLine"}, indent=2, default=str))
    print()
