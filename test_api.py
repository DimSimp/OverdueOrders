"""
Quick API test script — run this from the project root to diagnose Neto and eBay
connectivity without launching the full GUI.

Usage:
    python test_api.py
"""
from __future__ import annotations

import json
import sys
from datetime import datetime, timedelta

# ── Load config ──────────────────────────────────────────────────────────────
from src.config import config

try:
    config.load()
    print("[OK] config.json loaded")
except FileNotFoundError:
    print("[FAIL] config.json not found — run from the project root directory")
    sys.exit(1)
except Exception as e:
    print(f"[FAIL] config.json error: {e}")
    sys.exit(1)

date_to   = datetime.now()
date_from = date_to - timedelta(days=config.app.order_lookback_days)
print(f"  Date range: {date_from.date()} to {date_to.date()}")
print()

# ══════════════════════════════════════════════════════════════════════════════
# NETO
# ══════════════════════════════════════════════════════════════════════════════
print("=" * 60)
print("NETO")
print("=" * 60)
print(f"  Store URL : {config.neto.store_url}")
print(f"  API key   : {config.neto.api_key[:8]}...")

from src.neto_client import NetoClient, NetoAPIError

neto = NetoClient(config.neto)
try:
    orders = neto.get_overdue_orders(date_from, date_to)
    print(f"\n[OK] Neto OK — {len(orders)} undispatched paid order(s) in range\n")

    on_po = [o for o in orders if o.notes]
    print(f"  Orders with any notes: {len(on_po)}")
    for o in orders[:10]:          # show up to 10
        skus = ", ".join(l.sku for l in o.line_items if l.sku) or "(no SKUs)"
        notes_preview = o.notes[:80] + "..." if len(o.notes) > 80 else o.notes
        print(f"  [{o.order_id}] {o.customer_name:<25} | {o.date_paid and o.date_paid.date()} | {o.status}")
        print(f"    SKUs  : {skus}")
        if o.notes:
            print(f"    Notes : {notes_preview}")

    if len(orders) > 10:
        print(f"  ... and {len(orders) - 10} more")

except NetoAPIError as e:
    print(f"\n[FAIL] Neto API error:\n  {e}")
except Exception as e:
    print(f"\n[FAIL] Neto request failed:\n  {type(e).__name__}: {e}")

print()

# ══════════════════════════════════════════════════════════════════════════════
# EBAY
# ══════════════════════════════════════════════════════════════════════════════
print("=" * 60)
print("EBAY")
print("=" * 60)
print(f"  Environment  : {config.ebay.environment}")
print(f"  Client ID    : {config.ebay.client_id[:20]}...")
print(f"  Refresh token: {'SET' if config.ebay.refresh_token else 'NOT SET'}")

from src.ebay_client import EbayClient, EbayAuthError, EbayAPIError as EbayAPIErr

def _save(access_token, expires_at, refresh_token=None):
    config.save_ebay_tokens(access_token, expires_at, refresh_token)
    print(f"  (tokens saved to config.json)")

ebay = EbayClient(config.ebay, _save)

if not ebay.is_authenticated():
    print("\n  eBay is NOT authenticated yet.")
    print("  Run the app (python main.py), go to the Orders tab, and click 'Authenticate eBay'.")
    print("  Or run the auth flow below:\n")
    auth_url = ebay.get_auth_url()
    print(f"  Auth URL:\n  {auth_url}\n")
    print("  Paste the full redirect URL here (or press Enter to skip):")
    try:
        redirect = input("  > ").strip()
    except (EOFError, KeyboardInterrupt):
        redirect = ""
    if redirect:
        try:
            ebay.exchange_code(redirect)
            print("  [OK] eBay tokens saved — re-run this script to fetch orders")
        except Exception as e:
            print(f"  [FAIL] Auth error: {e}")
    sys.exit(0)

print("\n  eBay is authenticated — fetching orders...")
try:
    orders = ebay.get_overdue_orders(date_from, date_to)
    print(f"\n[OK] eBay OK — {len(orders)} unfulfilled paid order(s) in range\n")

    for o in orders[:10]:
        skus = ", ".join(l.sku for l in o.line_items if l.sku) or "(no SKUs)"
        notes_preview = o.buyer_notes[:60] + "..." if len(o.buyer_notes) > 60 else o.buyer_notes
        print(f"  [{o.order_id[:20]}] {o.buyer_name:<25} | {o.creation_date and o.creation_date.date()}")
        print(f"    SKUs  : {skus}")
        if o.buyer_notes:
            print(f"    Notes : {notes_preview}")

    if len(orders) > 10:
        print(f"  ... and {len(orders) - 10} more")

except EbayAuthError as e:
    print(f"\n[FAIL] eBay auth error:\n  {e}")
except EbayAPIErr as e:
    print(f"\n[FAIL] eBay API error:\n  {e}")
except Exception as e:
    print(f"\n[FAIL] eBay request failed:\n  {type(e).__name__}: {e}")
