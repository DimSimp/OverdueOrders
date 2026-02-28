"""
Test Neto GetNote action — sticky notes on orders are stored separately.
Run: py test_neto_getnote.py
"""
from __future__ import annotations
import json, sys, requests
from src.config import config

try:
    config.load()
except Exception as e:
    print(f"[FAIL] {e}"); sys.exit(1)

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

# --- Try GetNote action ---
print("Trying GetNote action...")
try:
    body = {
        "Filter": {
            "OutputSelector": ["NoteID", "OrderID", "Note", "NoteText", "NoteType",
                               "StaffNote", "IsSticky", "DateAdded", "AddedBy"],
            "Page": 1,
            "Limit": 50,
        }
    }
    data = post("GetNote", body)
    print(f"Ack: {data.get('Ack')}")
    if data.get("Ack") in ("Success", "Warning"):
        notes = data.get("Note", [])
        if isinstance(notes, dict):
            notes = [notes]
        print(f"Got {len(notes)} note(s)")
        # Show all keys
        all_keys = set()
        for n in notes:
            all_keys.update(n.keys())
        print(f"Keys: {sorted(all_keys)}")
        print()
        for n in notes[:10]:
            print(json.dumps(n, indent=2, default=str))
            print()
    else:
        print(f"Error: {data.get('Messages')}")
        print(json.dumps(data, indent=2, default=str))
except Exception as e:
    print(f"GetNote failed: {type(e).__name__}: {e}")

print()

# --- Try GetOrderNote action ---
print("Trying GetOrderNote action...")
try:
    body = {"Filter": {"Page": 1, "Limit": 10}}
    data = post("GetOrderNote", body)
    print(f"Ack: {data.get('Ack')}")
    print(json.dumps(data, indent=2, default=str)[:1000])
except Exception as e:
    print(f"GetOrderNote failed: {type(e).__name__}: {e}")

print()

# --- Try GetNote filtered to a specific order (N476399 - the one we know exists) ---
print("Trying GetNote for order N476399...")
try:
    body = {
        "Filter": {
            "OrderID": "N476399",
            "OutputSelector": ["NoteID", "OrderID", "Note", "NoteText", "NoteType",
                               "StaffNote", "IsSticky", "DateAdded"],
            "Page": 1,
            "Limit": 50,
        }
    }
    data = post("GetNote", body)
    print(f"Ack: {data.get('Ack')}")
    print(json.dumps(data, indent=2, default=str)[:2000])
except Exception as e:
    print(f"GetNote (order filter) failed: {type(e).__name__}: {e}")
