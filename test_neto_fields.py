"""
Try every plausible note/comment field for a specific order.
Also: if GetNote works with OrderID as array, test that too.
Run: py test_neto_fields.py
"""
from __future__ import annotations
import json, sys, requests
from src.config import config

try:
    config.load()
except Exception as e:
    print(f"[FAIL] {e}"); sys.exit(1)

TARGET_ORDER = "N476399"  # order we know exists

def post(action, body):
    resp = requests.post(
        f"{config.neto.store_url}/do/WS/NetoAPI",
        headers={
            "NETOAPI_ACTION":   action,
            "NETOAPI_KEY":      config.neto.api_key,
            "NETOAPI_USERNAME": config.neto.username,
            "Accept": "application/json", "Content-Type": "application/json",
        },
        json=body, timeout=30,
    )
    resp.raise_for_status()
    return resp.json()

# GetOrder with every note-related field we can think of
print(f"GetOrder for {TARGET_ORDER} with extended OutputSelector:")
all_fields = [
    "OrderID", "Username", "Email", "OrderStatus",
    "InternalOrderNotes", "DeliveryInstruction",
    "GiftMessage", "SpecialInstruction", "SpecialInstructions",
    "Note", "Notes", "NoteText",
    "StaffNote", "StaffNotes", "StaffComment",
    "OrderNote", "OrderNotes", "OrderComment", "OrderComments",
    "CustomerNotes", "CustomerNote",
    "PrivateNote", "PrivateNotes",
    "InternalNotes", "InternalNote",
    "Memo", "Comment", "Comments",
    "ShippingNote", "PackingNote",
    "OrderLine",
]
body = {
    "Filter": {
        "OrderID": [TARGET_ORDER],
        "OutputSelector": all_fields,
    }
}
data = post("GetOrder", body)
orders = data.get("Order", [])
if isinstance(orders, dict):
    orders = [orders]
if orders:
    o = orders[0]
    print("All returned fields:")
    for k, v in sorted(o.items()):
        if k != "OrderLine":
            print(f"  {k}: {v!r}")
    print(f"  OrderLine: {o.get('OrderLine')}")
else:
    print("No order returned")
    print(json.dumps(data, indent=2, default=str))

print()

# Try GetNote with OrderID as array
print(f"GetNote with OrderID as array [{TARGET_ORDER}]:")
try:
    body2 = {
        "Filter": {
            "OrderID": [TARGET_ORDER],
            "Page": 1,
            "Limit": 50,
        }
    }
    data2 = post("GetNote", body2)
    print(f"Ack: {data2.get('Ack')}")
    print(json.dumps(data2, indent=2, default=str)[:2000])
except Exception as e:
    print(f"  Error: {e}")

print()

# Try GetNote with no filter at all (just empty body)
print("GetNote with minimal filter body:")
try:
    body3 = {"Filter": {"Page": 1, "Limit": 5}}
    data3 = post("GetNote", body3)
    print(f"Ack: {data3.get('Ack')}")
    notes = data3.get("Note", [])
    if isinstance(notes, dict):
        notes = [notes]
    print(f"Notes returned: {len(notes)}")
    if notes:
        print(f"Keys: {sorted(notes[0].keys())}")
        print(json.dumps(notes[0], indent=2, default=str))
    else:
        print(json.dumps(data3, indent=2, default=str)[:1000])
except Exception as e:
    print(f"  Error: {e}")
