"""Quick connection test — run from the project root to verify Supabase is reachable
and all expected tables were created successfully.

Usage:
    py test_supabase.py  (or full python path)
"""
from src.config import config
from src.supabase_client import init_client, get_client

config.load()

if not config.supabase:
    print("FAIL: No supabase credentials found in config.json")
    raise SystemExit(1)

print(f"Connecting to {config.supabase.url} ...")
init_client(config.supabase.url, config.supabase.key)
db = get_client()

EXPECTED_TABLES = [
    "users",
    "suppliers",
    "supplier_contacts",
    "items",
    "stock_movements",
    "discounts",
    "customers",
    "transactions",
    "transaction_lines",
]

all_ok = True
for table in EXPECTED_TABLES:
    try:
        result = db.table(table).select("*").limit(1).execute()
        print(f"  OK  {table}")
    except Exception as e:
        print(f"  FAIL {table}: {e}")
        all_ok = False

# Verify the discount presets were seeded
try:
    result = db.table("discounts").select("name, percentage").eq("is_system", True).execute()
    presets = [(r["name"], r["percentage"]) for r in result.data]
    print(f"\nDiscount presets seeded: {presets}")
except Exception as e:
    print(f"\nFAIL reading discounts: {e}")
    all_ok = False

print()
if all_ok:
    print("All tables OK — Supabase connection is working.")
else:
    print("One or more tables had errors — check output above.")

# ── Recent sale verification ───────────────────────────────────────────────
print("\n── Last 3 transactions ──────────────────────────────────────────────")
try:
    txs = (
        db.table("transactions")
        .select("transaction_number, sale_type, sale_status, total, performed_by, completed_at")
        .order("completed_at", desc=True)
        .limit(3)
        .execute()
    )
    if not txs.data:
        print("  (none found)")
    for tx in txs.data:
        print(f"  {tx['transaction_number']}  {tx['sale_type']:<10}  "
              f"{tx['sale_status']:<12}  ${tx['total']:.2f}  "
              f"by={tx['performed_by'] or '—'}  at={tx['completed_at'][:19]}")
except Exception as e:
    print(f"  FAIL: {e}")

print("\n── Lines for most recent transaction ────────────────────────────────")
try:
    latest = db.table("transactions").select("id, transaction_number").order("completed_at", desc=True).limit(1).execute()
    if latest.data:
        tx_id = latest.data[0]["id"]
        tx_num = latest.data[0]["transaction_number"]
        lines = db.table("transaction_lines").select("sku, description, qty, unit_price, discount_pct, line_total").eq("transaction_id", tx_id).execute()
        print(f"  {tx_num} — {len(lines.data)} line(s):")
        for ln in lines.data:
            disc = f"  disc={ln['discount_pct']}%" if ln['discount_pct'] else ""
            print(f"    {ln['sku']:<20}  qty={ln['qty']}  @${ln['unit_price']:.2f}{disc}  total=${ln['line_total']:.2f}")
except Exception as e:
    print(f"  FAIL: {e}")

print("\n── Last 5 stock movements ───────────────────────────────────────────")
try:
    mvts = (
        db.table("stock_movements")
        .select("item_id, movement_type, qty_change, reference_id, performed_by, created_at")
        .order("created_at", desc=True)
        .limit(5)
        .execute()
    )
    if not mvts.data:
        print("  (none found)")
    for m in mvts.data:
        print(f"  {m['movement_type']:<16}  qty={m['qty_change']:+d}  ref={m['reference_id'] or '—'}"
              f"  by={m['performed_by'] or '—'}  at={m['created_at'][:19]}")
except Exception as e:
    print(f"  FAIL: {e}")
