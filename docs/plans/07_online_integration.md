# Plan 07 — Online Order Integration

> **Part of**: [Master Plan](00_overview.md)
> **Status**: 🔲 Not started
> **Phase**: 2 — Online Bridge

---

## Overview

The online integration keeps Supabase inventory in sync with Neto and eBay order activity — automatically and without staff action. It has two distinct parts:

1. **Sync Script** — a lightweight Python script hosted on GitHub Actions that runs every 5 minutes. It fetches active orders from both platforms and creates or updates allocation records in Supabase, without touching `qty_on_hand`.
2. **Dispatch Hook** — a code addition to the existing desktop app. When staff mark an order as dispatched in the existing daily workflow, the hook commits the stock movement to Supabase: `qty_on_hand` decrements, the allocation is cleared, and the sale is logged.

Stock flows in one direction until dispatch:

```
New online order arrives
        │
        ▼
Sync script allocates stock
qty_allocated_online +qty  ←— On Hand unchanged
qty_available effectively −qty
        │
        ▼
Staff dispatches order in existing app
        │
        ▼
Dispatch hook fires
qty_on_hand −qty
qty_allocated_online −qty  (allocation cleared)
sale written to sales log
```

---

## Key Concepts

| Concept | Description |
|---------|-------------|
| **Sync Script** | Standalone Python script run by GitHub Actions every 5 min. Allocates/de-allocates stock based on order status. Idempotent — safe to run multiple times. |
| **Dispatch Hook** | Code in existing app (`src/session.py` / `src/session_daily.py`) that calls Supabase when an order is dispatched. This is the only path that decrements `qty_on_hand`. |
| **`online_allocations`** | One row per order line item currently allocated. Sync creates these; dispatch hook deletes them. |
| **Web SKU** | The SKU as listed on Neto/eBay (with supplier suffix/prefix applied). Must be resolved back to internal `sku` before any Supabase operation. |
| **Idempotency** | The sync always checks for an existing allocation record before creating one. Running the same sync twice has no effect. |
| **Oversell** | When an order is placed on the platform for more units than are available in Supabase. The allocation is created anyway (the order is already placed); a warning flag is set for staff to resolve. |
| **Dropship** | An order fulfilled directly by the supplier — stock never passes through the store. Neither the sync script nor the dispatch hook touches inventory for dropship orders. |
| **Replacement** | A second dispatch for the same order, sent at no charge (e.g. lost in transit). Logged as `sale_type = 'replacement'` with `sale_price = 0` so it appears in the stock movement audit without inflating revenue. |

---

## Database Schema

### `online_allocations` — Active stock commitments for unfulfilled online orders

| Column | Type | Notes |
|--------|------|-------|
| `id` | uuid PK | |
| `order_id` | text NOT NULL | Platform order identifier (Neto order number, eBay order ID) |
| `platform` | text NOT NULL | `neto` or `ebay` |
| `item_id` | uuid FK → `items.id` | Resolved internal item; nullable if SKU not matched |
| `web_sku` | text NOT NULL | SKU as it appears on the platform order |
| `sku` | text | Resolved internal SKU (null if resolution failed) |
| `qty` | integer NOT NULL | Units allocated for this order line |
| `order_status` | text | Last known status from the platform (`pick`, `dispatched`, `cancelled`, etc.) |
| `customer_name` | text | For display/diagnostics only |
| `oversell_flag` | boolean DEFAULT false | True if allocation caused Available to go negative |
| `is_dropship` | boolean DEFAULT false | True if fulfilled directly by supplier — no stock movement |
| `allocated_at` | timestamptz DEFAULT now() | When this allocation was first created |
| `last_synced_at` | timestamptz | Updated every sync run to confirm this order is still active |

**Constraint**: `UNIQUE (order_id, web_sku)` — prevents duplicate allocation rows for the same order line.

---

### `online_sales` — Dispatched online order lines (for reporting)

Created by the dispatch hook when an order is finalised. Feeds plan 08 reporting.

| Column | Type | Notes |
|--------|------|-------|
| `id` | uuid PK | |
| `order_id` | text NOT NULL | Platform order ID |
| `platform` | text NOT NULL | `neto` or `ebay` |
| `item_id` | uuid FK → `items.id` | |
| `sku` | text | Internal SKU |
| `web_sku` | text | Platform SKU |
| `qty` | integer NOT NULL | Units dispatched |
| `sale_price` | numeric(10,2) | Unit sale price at time of dispatch |
| `cost_at_dispatch` | numeric(10,2) | `last_purchase_cost` at time of dispatch (for margin reporting) |
| `sale_type` | text DEFAULT `'standard'` | `standard` or `replacement` (free re-send for lost/damaged parcel) |
| `is_dropship` | boolean DEFAULT false | True if fulfilled directly by supplier |
| `dispatched_at` | timestamptz | |
| `dispatched_by` | text | Staff user who ran the dispatch |
| `returned` | boolean DEFAULT false | True if stock was physically received back |
| `returned_at` | timestamptz | |

---

## The Sync Script

### Location and Hosting

The sync script lives in the existing repo as `sync/order_sync.py` and is executed by GitHub Actions on a schedule. Because the GitHub repo is already public (required for the update checker), no additional hosting cost is incurred.

```
.github/
└── workflows/
    └── order_sync.yml   — scheduled workflow: runs every 5 min
sync/
└── order_sync.py        — main sync entrypoint
```

The script imports shared client modules from `src/` (`neto_client.py`, `ebay_client.py`) and a new `src/supabase_client.py`. Secrets (API keys, Supabase URL/key) are stored as GitHub Actions repository secrets.

> **Minimum schedule**: GitHub Actions enforces a minimum of 5 minutes for scheduled workflows. This is acceptable — a 5-minute allocation window is fine for a music instrument retailer where same-minute stock clashes are unlikely.

> **Local alternative**: If internet access is unreliable or GitHub Actions is unsuitable, the same `order_sync.py` can be run as a Windows Scheduled Task on the store PC (every 5 minutes, pointing to the same Supabase instance). Both approaches produce identical outcomes.

---

### Sync Logic (per run)

```
1. Fetch all active orders from Neto (status: "Pick")
2. Fetch all active orders from eBay (status: "IN_PROGRESS")
3. For each order line across both platforms:
   a. Check if online_allocations row already exists (order_id + web_sku)
   b. If NOT exists → create allocation (see: New Order flow)
   c. If EXISTS and order is still active → update last_synced_at only
4. Fetch all online_allocations rows where last_synced_at < (now − 10 min)
   — these are orders that have disappeared from the active list
5. For each stale allocation:
   a. Re-fetch the specific order from the platform to confirm its current status
   b. If CANCELLED → run Cancellation flow (see below)
   c. If DISPATCHED → the dispatch hook should have already handled this;
      if allocation still exists it means hook hasn't fired yet — leave it alone
   d. If still ACTIVE but missed the sync window → update last_synced_at, leave allocation
```

---

### New Order Flow

When a new order line is detected (no existing allocation row):

1. Resolve `web_sku` → internal `item_id` and `sku`
   - Query `items` where `web_sku = ?`
   - If no match: create the allocation with `item_id = null`, `sku = null` — logs an unresolved SKU warning
2. Insert row into `online_allocations`
3. If `item_id` resolved:
   - Increment `items.qty_allocated_online` by `qty`
   - Check if `qty_on_hand − qty_allocated_online − qty_allocated_customer < 0`
   - If yes: set `oversell_flag = true` on the allocation row
4. Write `stock_movements` record of type `allocate_online` (qty_change = +qty for audit purposes)

---

### SKU Resolution

Online platforms use `web_sku` (e.g. `2221AUSTRALIS`). Supabase `items` stores both `sku` (`2221`) and `web_sku` (`2221AUSTRALIS`). Resolution is a direct lookup:

```sql
SELECT id, sku FROM items WHERE web_sku = :web_sku
```

If `web_sku = sku` (no suffix supplier), this still works because `web_sku` defaults to `sku`.

**Resolution failure** (no match found):
- Allocation row is created with `item_id = null`
- A warning is queued: the desktop app shows an "Unresolved SKU" badge on the Orders tab
- Staff can manually link the SKU from within the app (maps it in a persistent `sku_aliases` table)
- Once resolved, the sync will match correctly on the next run

---

## The Dispatch Hook

### Where It Lives

The existing app dispatches orders in `src/session.py` and/or `src/session_daily.py`. After the existing dispatch logic runs (marking the order as sent on Neto/eBay), the hook calls:

```python
from src.supabase_client import SupabaseClient

supabase = SupabaseClient()
supabase.on_order_dispatched(order_id, platform, lines)
```

Where `lines` is a list of `{web_sku, qty, sale_price}` from the order.

### What the Hook Does

For each line in the dispatched order:

1. Resolve `web_sku` → `item_id` (same lookup as sync script)
2. Decrement `items.qty_on_hand` by `qty`
3. Decrement `items.qty_allocated_online` by `qty` (clearing the allocation)
4. Delete the matching `online_allocations` row
5. Write `stock_movements` record of type `dispatch` (qty_change = −qty)
6. Write `online_sales` record (for reporting)
7. If the item is serialised: prompt staff to select the serial number dispatched (same as plan 01 dispatch flow)

If no matching `online_allocations` row exists (e.g. order was placed before the inventory system went live):
- Still decrement `qty_on_hand` and write `stock_movements`
- Log a note: "No allocation found for this order — direct stock decrement applied"

---

## Cancellations

### Pre-Dispatch Cancellation — Platform-Detected

Order cancelled on the platform before staff dispatch it:

1. Sync script detects the order no longer appears in the active order list
2. Re-fetches the order → confirms status is `cancelled` / `refunded`
3. For each allocation row for this order:
   - Decrement `items.qty_allocated_online` by `qty`
   - Delete the `online_allocations` row
   - Write `stock_movements` record of type `allocate_online` with negative qty_change (allocation released)
4. `qty_on_hand` is **not changed** — stock never left the store
5. The desktop app shows a "Cancelled" badge on the order the next time staff open it

### Pre-Dispatch Cancellation — Staff-Initiated

A customer phones to cancel before the platform order status has updated, or staff need to cancel on the store's side (e.g. item found to be defective before dispatch). This is handled directly from the **Daily Operations order detail view**:

1. Staff finds the order (by order number, customer name, etc.) in the existing search/filter interface
2. Right-click the order row → **"Cancel Order"**
3. Confirmation prompt: *"Cancel order [ID]? This will release the stock allocation and update the order status on [platform]."*
4. On confirm:
   - The order status is updated to `cancelled` on the platform (Neto or eBay API call)
   - `items.qty_allocated_online` decremented for each line
   - `online_allocations` rows deleted
   - `stock_movements` record of type `allocate_online` written (negative qty_change)
5. `qty_on_hand` unchanged

If the platform API call to cancel the order fails, the inventory update is still applied and the failure is shown to staff with a note to cancel manually on the platform portal.

### Partial Line Cancellation

A customer removes one item from a multi-item order before dispatch (more common on Neto). Only the allocation for that one SKU is released:

1. Sync script detects a qty change or line removal on an existing `online_allocations` row
   - Full line removed: treat as cancellation for that line (release allocation for that SKU)
   - Qty reduced: apply delta as per the Order Qty Edit flow
2. Allocations for other lines in the same order are unaffected
3. The remaining allocation rows continue normally through to dispatch

Staff can also trigger a partial line cancellation manually from the order detail view: right-click a specific order line → **"Cancel This Line"** → same confirmation and inventory update flow as above, scoped to that SKU only.

### Post-Dispatch Cancellation (parcel returned undelivered)

Order was dispatched (stock decremented by hook), but the courier returns the parcel undelivered. The platform order may still show as "Dispatched" — the return is a physical event, not a platform status change.

**Staff action required** — this cannot be detected automatically:

1. Staff opens the relevant order in the existing app
2. Right-click → **"Receive Return"**
3. Confirm which items are being returned and their condition:
   - **Resaleable** → `qty_on_hand` +qty, `stock_movements` of type `return`
   - **Damaged / not resaleable** → no stock increase; `stock_movements` of type `adjustment` (reason: `Damaged — returned delivery`)
4. The original `online_sales` record is marked `returned = true`

### Post-Dispatch Customer Return (customer-initiated refund)

Customer receives the item but initiates a return through eBay/Neto. The return is processed in the platform's portal (refund issued, return tracking, etc.). The physical stock may or may not come back to the store.

**Same flow as post-dispatch cancellation** — staff trigger "Receive Return" when the physical item arrives back at the store. If no stock is received (e.g. customer kept the item as part of partial refund), no inventory change is made; the return is recorded in notes only.

> **Out of scope**: Automated return detection from platform APIs. eBay and Neto return flows vary and are complex. Manual staff confirmation is the appropriate approach for a store of this scale — staff are already handling the physical goods and the portal refund.

### Lost in Transit — Free Replacement

Parcel is confirmed lost by the courier, and the store agrees to re-send the item at no charge. The original dispatch stands (stock was decremented correctly). A replacement shipment is a second stock movement.

1. Staff opens the original order in the Daily Ops order detail view
2. Right-click → **"Send Replacement"**
3. Confirmation prompt shows the original order lines; staff confirm which items to re-send
4. On confirm:
   - `items.qty_on_hand` decremented again for each re-sent item
   - `stock_movements` record of type `dispatch` written (qty_change = −qty)
   - `online_sales` record written with `sale_type = 'replacement'`, `sale_price = 0`
   - A new courier booking dialog opens so staff can re-book the shipment immediately
5. The original `online_sales` record is unchanged (revenue was real; the replacement is a separate cost)

This ensures the stock audit correctly shows two outbound movements for the one order, and reporting can distinguish replacement costs from genuine revenue.

### Wrong Item Dispatched

Staff picked and dispatched the wrong SKU. Two separate stock corrections are needed.

1. Staff opens the original order in the Daily Ops order detail view
2. Right-click → **"Report Wrong Item"**
3. Staff specify:
   - **SKU that was incorrectly sent** (the wrong item that left the store)
   - **SKU that should have been sent** (the correct item)
4. On confirm — two stock movements are written:
   - Wrong SKU: `stock_movements` of type `return` (qty_change = +qty) — *when the item physically comes back*
   - Correct SKU: `stock_movements` of type `dispatch` (qty_change = −qty) — when the replacement is sent
5. The correct SKU re-dispatch triggers a new courier booking
6. Staff must physically receive the wrong item back before its On Hand is incremented — the flow prompts: *"Has the incorrect item been returned to the store?"*
   - Yes → stock incremented immediately
   - No → a pending return flag is set; stock incremented when "Receive Return" is triggered later

---

## Edge Cases

### Oversell Detection

If an order is placed on the platform for more stock than Available:

- Allocation is created anyway (the order is confirmed — we can't un-take it)
- `oversell_flag = true` on the allocation row
- `items.qty_allocated_online` may exceed `qty_on_hand`, making Available negative
- Desktop app shows an **"Oversell Warning"** banner on startup: *"SKU [X] is oversold — [N] units short. View orders?"*
- Staff must resolve: procure more stock, cancel the order, or manually adjust

### Order Qty Edit (customer changes qty before dispatch)

Some platforms allow customers to edit order quantities after placing. If the sync detects a qty change on an existing `online_allocations` row:

1. Compute delta: `new_qty − existing_qty`
2. If delta > 0 (qty increased): increment `qty_allocated_online` by delta, check for oversell
3. If delta < 0 (qty decreased): decrement `qty_allocated_online` by |delta|
4. Update `online_allocations.qty` to new_qty

### Unresolved SKU at Dispatch

If dispatch hook can't resolve the `web_sku`:
- Still log the dispatch attempt
- Show error in the existing app's dispatch confirmation: *"Could not update inventory for SKU [X] — item not found in Supabase. Please reconcile manually."*
- Write a `stock_movements` record of type `dispatch` with `notes = "auto-resolution failed"`

### Sync Script Downtime

If the sync script fails to run for a period (GitHub Actions outage, etc.):
- `last_synced_at` will be stale for all active allocations
- On recovery, the sync re-fetches all active orders and re-confirms every allocation
- New orders placed during downtime are picked up immediately on recovery
- No allocations are wrongly deleted during downtime (stale check threshold is 10 min — well above GitHub Actions 5-min window, so a single missed run doesn't trigger false cancellations)

### Dropship Orders

Some online orders are fulfilled directly by the supplier — the item never enters or leaves the store's physical stock. For these orders, no inventory movement should occur.

**Detection**: The existing app and Web Portal already flag dropship orders via `is_dropship = true` (from the parsed invoice data). This flag should also be set on the Neto/eBay order record when the order contains only dropship items.

**Sync script behaviour**: When a new order line is detected and the order is flagged as dropship:
- Insert the `online_allocations` row with `is_dropship = true`
- Do **not** increment `qty_allocated_online`
- The row exists only for tracking/audit purposes

**Dispatch hook behaviour**: When a dropship order is dispatched:
- Do **not** decrement `qty_on_hand`
- Write `online_sales` record with `is_dropship = true` (for revenue reporting — the sale is real even if no physical stock moved)
- Do **not** write a `stock_movements` record
- Delete the `online_allocations` row

**Schema addition**: Add `is_dropship boolean DEFAULT false` to both `online_allocations` and `online_sales`.

> If only some lines in a multi-item order are dropship, each line is handled independently — dropship lines skip stock movements; non-dropship lines follow the normal flow.

### Orders Placed Before Inventory System Goes Live

On cutover, there will be existing active orders in Neto/eBay with no corresponding `online_allocations`. On the first sync run, these are treated as new orders and allocations are created. If `qty_on_hand` was set correctly in the initial stocktake, these allocations will accurately reflect what's committed.

---

## Error Handling and Monitoring

### Sync Script

- All errors are logged to a Supabase `sync_log` table (run_at, platform, orders_processed, allocations_created, allocations_released, errors)
- On fatal error (API timeout, Supabase unreachable), the script exits with a non-zero code → GitHub Actions marks the run as failed → email notification (GitHub's built-in failure alerts)
- Non-fatal errors (single SKU unresolved, single order fetch failed) are logged but don't abort the run

### Desktop App

- On Supabase unavailable during dispatch hook: warn staff with a dialog, but **do not block the dispatch**. Log the failed Supabase call to a local pending queue (`data/pending_sync.json`). On next app start, the queue is replayed.
- Pending sync queue: `[{order_id, platform, lines, dispatched_at, action: "dispatch"}]`

---

### `sync_log` Table

| Column | Type | Notes |
|--------|------|-------|
| `id` | uuid PK | |
| `run_at` | timestamptz DEFAULT now() | |
| `platform` | text | `neto`, `ebay`, or `all` |
| `orders_checked` | integer | |
| `allocations_created` | integer | |
| `allocations_updated` | integer | |
| `allocations_released` | integer | |
| `unresolved_skus` | integer | |
| `oversells_detected` | integer | |
| `error_message` | text | null if clean run |
| `duration_ms` | integer | |

---

## Implementation Checklist

### Database
- [ ] Create `online_allocations` table in Supabase with `UNIQUE (order_id, web_sku)` constraint
- [ ] Create `online_sales` table in Supabase
- [ ] Create `sync_log` table in Supabase
- [ ] Optional: `sku_aliases` table for persistent manual SKU resolutions

### Sync Script (`sync/order_sync.py`)
- [ ] Scaffold `sync/order_sync.py` — imports Neto and eBay clients, Supabase client
- [ ] New Order flow: resolve SKU, insert `online_allocations`, increment `qty_allocated_online`
- [ ] Dropship detection: if `is_dropship = true`, insert allocation row but skip `qty_allocated_online` increment
- [ ] Oversell detection: check Available < 0, set `oversell_flag`, log to `sync_log`
- [ ] Stale allocation detection: query rows where `last_synced_at < now − 10 min`
- [ ] Cancellation flow: confirm with platform API, decrement `qty_allocated_online`, delete row
- [ ] Partial line cancellation: detect removed/zeroed lines, release only that line's allocation
- [ ] Order qty edit detection: compare fetched qty vs stored qty, apply delta
- [ ] Write `stock_movements` records for all allocation changes (type `allocate_online`)
- [ ] Write `sync_log` record after every run (success or failure)
- [ ] Unresolved SKU: create allocation with `item_id = null`, increment unresolved counter

### GitHub Actions Workflow (`.github/workflows/order_sync.yml`)
- [ ] Schedule: `*/5 * * * *` (every 5 minutes)
- [ ] Set up Python environment, install deps
- [ ] Pass secrets: `NETO_API_KEY`, `EBAY_TOKEN`, `SUPABASE_URL`, `SUPABASE_KEY`
- [ ] Call `sync/order_sync.py`
- [ ] On failure: GitHub's native email alert to repo owner

### Dispatch Hook (existing app)
- [ ] Add `supabase_client.on_order_dispatched(order_id, platform, lines)` call in `src/session.py` / `src/session_daily.py` after successful dispatch
- [ ] Hook: for non-dropship lines — decrement `qty_on_hand`, decrement `qty_allocated_online`, delete `online_allocations` row, write `stock_movements`
- [ ] Hook: for dropship lines — write `online_sales` with `is_dropship = true`, skip all stock movements, delete allocation row
- [ ] Hook: write `online_sales` record (with `sale_price`, `cost_at_dispatch`, `sale_type = 'standard'`)
- [ ] Graceful failure: if Supabase unreachable, log to `data/pending_sync.json` and continue
- [ ] Pending sync queue: replay on next app start

### Desktop App — Order Detail Actions
- [ ] "Cancel Order" right-click action in Daily Ops order list (pre-dispatch, staff-initiated)
  - [ ] Platform API cancel call (Neto / eBay)
  - [ ] Release full-order allocation on confirm; graceful failure if API call fails
- [ ] "Cancel This Line" right-click action on individual order line (partial line cancellation)
  - [ ] Release allocation for that SKU only; leave remaining lines untouched
- [ ] "Receive Return" right-click action on dispatched orders
  - [ ] Resaleable path: `qty_on_hand` +qty, `stock_movements` of type `return`
  - [ ] Damaged path: `stock_movements` of type `adjustment` with reason `Damaged — return`
  - [ ] Mark `online_sales.returned = true`, set `returned_at`
- [ ] "Send Replacement" right-click action on dispatched orders (lost in transit)
  - [ ] Decrement `qty_on_hand` for re-sent items
  - [ ] Write `online_sales` with `sale_type = 'replacement'`, `sale_price = 0`
  - [ ] Open courier booking dialog for re-shipment
- [ ] "Report Wrong Item" right-click action (wrong SKU dispatched)
  - [ ] Prompt for wrong SKU and correct SKU
  - [ ] Write return movement for wrong SKU (conditional on physical receipt)
  - [ ] Write dispatch movement for correct SKU; open courier booking
  - [ ] Pending return flag if wrong item not yet back in store

### Desktop App Notifications
- [ ] Oversell warning banner on app startup (check `online_allocations` where `oversell_flag = true`)
- [ ] Unresolved SKU badge in Orders tab
- [ ] Manual SKU link tool (maps `web_sku` → `item_id` for unresolved allocations)
- [ ] Pending sync queue indicator in status bar (shows count of queued Supabase writes)

### Supabase Client (`src/supabase_client.py`)
- [ ] `SupabaseClient` class using `supabase-py` library
- [ ] `on_order_dispatched(order_id, platform, lines)` — dispatch hook method
- [ ] `create_allocation(order_id, platform, web_sku, qty)` — used by sync script
- [ ] `release_allocation(order_id, web_sku)` — cancellation
- [ ] `get_oversold_items()` — for startup warning
- [ ] Read Supabase URL and anon key from `config.json`

---

## Open Questions / Future Considerations

- **Returns automation**: eBay's Post-Order API and Neto's returns API could theoretically allow the sync to detect return requests automatically and flag them for staff confirmation. This would close the loop on the manual "Receive Return" step. Deferred — manual is fine at current scale.
- **Order edit detection on eBay**: eBay rarely allows post-placement qty edits; Neto sometimes does. The delta logic handles it, but edge cases around partial cancellations on a multi-line order haven't been stress-tested. Monitor after go-live.
- **Real-time webhook alternative**: Both Neto and eBay support webhooks/notifications for order events. A webhook listener (e.g. a small cloud function) could replace the polling script for near-instant allocation. Higher complexity, higher reliability. Deferred.
- **Multi-store**: Single-store assumption throughout. If a second location is ever added, allocations would need a `location_id`. Schema is not blocked by this — `online_allocations` could gain a `location_id` column later.

---

*Last updated: 2026-04-13 — added staff-initiated cancellation, partial line cancel, lost-in-transit replacement, wrong item dispatched, dropship handling*
