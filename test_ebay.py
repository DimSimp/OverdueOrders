r"""
eBay API diagnostic script.
Run from the project root:
  C:\Users\lorel\AppData\Local\Programs\Python\Python39\python.exe test_ebay.py

Checks (in order):
  1. Config file loads and eBay section is present
  2. Credentials look non-empty
  3. OAuth token refresh works
  4. Fulfillment API (list orders, last 7 days)
  5. Trading API — GetMyeBaySelling (1 day, page 1)
  6. Trading API — SetUserNotes dry-run (no write)
"""
from __future__ import annotations

import sys
import time
import traceback
from datetime import datetime, timedelta
from pathlib import Path

# ── bootstrap path ──────────────────────────────────────────────────────────
ROOT = Path(__file__).parent
sys.path.insert(0, str(ROOT))

from src.config import ConfigManager, EbayConfig, CONFIG_PATH  # noqa: E402
from src.ebay_client import EbayClient, EbayAPIError, EbayAuthError  # noqa: E402


# ── helpers ─────────────────────────────────────────────────────────────────

def ok(msg: str) -> None:
    print(f"  [OK]  {msg}")

def fail(msg: str) -> None:
    print(f"  [FAIL] {msg}")

def warn(msg: str) -> None:
    print(f"  [WARN] {msg}")

def section(title: str) -> None:
    print(f"\n{'─'*60}\n{title}\n{'─'*60}")


# ── 1. Config ────────────────────────────────────────────────────────────────
section("1. Config")
try:
    cfg = ConfigManager()
    cfg.load()
    ebay: EbayConfig = cfg.ebay
    ok(f"Config loaded from {CONFIG_PATH}")
    ok(f"environment = {ebay.environment!r}")
except Exception as exc:
    fail(f"Config load failed: {exc}")
    traceback.print_exc()
    sys.exit(1)


# ── 2. Credentials ───────────────────────────────────────────────────────────
section("2. Credentials")

fields = {
    "client_id": ebay.client_id,
    "client_secret": ebay.client_secret,
    "ru_name": ebay.ru_name,
}
all_ok = True
for name, val in fields.items():
    if val:
        ok(f"{name} is set ({val[:6]}…)")
    else:
        fail(f"{name} is EMPTY")
        all_ok = False

if ebay.refresh_token:
    ok(f"refresh_token is set ({ebay.refresh_token[:10]}…)")
else:
    fail("refresh_token is EMPTY — re-authenticate via the app first")
    all_ok = False

if ebay.dev_id:
    ok(f"dev_id is set — Trading API enabled")
else:
    warn("dev_id is empty — Trading API (PrivateNotes) will be skipped")

if ebay.user_token:
    ok(f"user_token is set ({ebay.user_token[:10]}…)")
else:
    warn("user_token is empty — will fall back to OAuth for Trading API")

if not all_ok:
    fail("Fix credential issues above before continuing.")
    sys.exit(1)


# ── token save stub ──────────────────────────────────────────────────────────
def _save_tokens(access_token: str, expires_at: float, refresh_token: str = "") -> None:
    """Write refreshed tokens back to config so they persist."""
    cfg.ebay.access_token = access_token
    cfg.ebay.access_token_expires_at = expires_at
    if refresh_token:
        cfg.ebay.refresh_token = refresh_token
    cfg.save()


client = EbayClient(ebay, _save_tokens)


# ── 3. OAuth token refresh ───────────────────────────────────────────────────
section("3. OAuth token refresh")
try:
    token = client._ensure_valid_token()
    exp = datetime.fromtimestamp(ebay.access_token_expires_at)
    ok(f"Access token OK (expires ~{exp.strftime('%Y-%m-%d %H:%M:%S')})")
    ok(f"Token prefix: {token[:20]}…")
except EbayAuthError as exc:
    fail(f"Auth error: {exc}")
    sys.exit(1)
except EbayAPIError as exc:
    fail(f"Token refresh failed: {exc}")
    traceback.print_exc()
    sys.exit(1)
except Exception as exc:
    fail(f"Unexpected error: {exc}")
    traceback.print_exc()
    sys.exit(1)


# ── 4. Fulfillment API — list orders ─────────────────────────────────────────
section("4. Fulfillment API — list orders (last 7 days)")
try:
    date_to = datetime.now()
    date_from = date_to - timedelta(days=7)
    orders = client.search_orders(date_from, date_to, enrich_notes=False)
    ok(f"Returned {len(orders)} orders")
    if orders:
        o = orders[0]
        ok(f"  First order: {o.order_id}  buyer={o.buyer_name!r}  status={o.order_status}")
except EbayAPIError as exc:
    fail(f"Fulfillment API error: {exc}")
    traceback.print_exc()
except Exception as exc:
    fail(f"Unexpected error: {exc}")
    traceback.print_exc()


# ── 5. Trading API — GetMyeBaySelling ────────────────────────────────────────
section("5. Trading API — GetMyeBaySelling (last 1 day)")
if not ebay.dev_id:
    warn("Skipping — dev_id not configured")
else:
    try:
        xml = client._build_sold_list_xml(duration_days=1, page=1)
        root = client._call_trading_api(xml, "GetMyeBaySelling")
        import xml.etree.ElementTree as ET
        ns = {"e": "urn:ebay:apis:eBLBaseComponents"}
        ack = root.find("e:Ack", ns)
        ok(f"Ack = {ack.text if ack is not None else '?'}")
        txns = root.findall(".//e:Transaction", ns)
        ok(f"Transactions found: {len(txns)}")
        if txns:
            t = txns[0]
            oli = t.find("e:OrderLineItemID", ns)
            ok(f"  First txn OLI: {oli.text if oli is not None else '?'}")
        total_pages_el = root.find(".//e:SoldList/e:PaginationResult/e:TotalNumberOfPages", ns)
        if total_pages_el is not None:
            ok(f"  Total pages: {total_pages_el.text}")
    except EbayAPIError as exc:
        fail(f"Trading API error: {exc}")
        traceback.print_exc()
    except Exception as exc:
        fail(f"Unexpected error: {exc}")
        traceback.print_exc()


# ── 6. Trading API — SetUserNotes dry-run ────────────────────────────────────
section("6. Trading API — SetUserNotes (DRY RUN only — no write)")
if not ebay.dev_id:
    warn("Skipping — dev_id not configured")
else:
    try:
        client.set_private_notes(
            item_id="000000000000",
            transaction_id="0",
            note_text="diagnostic test",
            dry_run=True,   # NEVER actually writes
        )
        ok("set_private_notes dry-run completed without error")
    except Exception as exc:
        fail(f"Unexpected error in dry-run: {exc}")
        traceback.print_exc()


# ── Summary ──────────────────────────────────────────────────────────────────
section("Done")
print("If any [FAIL] lines appear above, that is where the issue is.")
print("Share the full output (including any tracebacks) for further diagnosis.")