# Module: Online Order Integration

> **Detail plan**: [docs/plans/07_online_integration.md](../plans/07_online_integration.md)
> **Build phase**: 2 — Online Bridge
> **Tables owned**: `online_allocations`, `online_sales`, `sync_log`, `sku_aliases`

---

## Overview

Two independent components that together keep Supabase inventory in sync with Neto and eBay:

1. **Sync Script** (`sync/order_sync.py`) — GitHub Actions cron every 5 min. Allocates/de-allocates
   stock based on active order status. Never touches `qty_on_hand`.
2. **Dispatch Hook** — added to `src/session.py` / `src/session_daily.py`. The **only** path that
   decrements `qty_on_hand` for online orders. Fires after the existing dispatch logic runs.

---

## Stock Flow

```
New online order detected by sync script
→ online_allocations row created
→ items.qty_allocated_online +qty  (qty_on_hand unchanged)

Staff dispatches order in existing app
→ dispatch hook fires
→ items.qty_on_hand −qty
→ items.qty_allocated_online −qty
→ online_allocations row deleted
→ online_sales row created
```

---

## Sync Script Logic (per run)

1. Fetch all active Neto orders (status: "Pick") + eBay orders (status: "IN_PROGRESS")
2. For each order line:
   - No existing allocation → create new (resolve SKU, set oversell flag if needed)
   - Existing allocation → update `last_synced_at`
3. Fetch allocations where `last_synced_at < (now − 10 min)` — stale
4. For each stale allocation → re-fetch from platform:
   - Cancelled → release allocation (`qty_allocated_online −qty`, delete row)
   - Dispatched → leave (dispatch hook handles it)
   - Still active → update `last_synced_at`

**Idempotent**: Running the sync twice has no effect.

---

## SKU Resolution

Online platforms use `web_sku` (e.g. `2221AUSTRALIS`). Resolved via:
```sql
SELECT id, sku FROM items WHERE web_sku = :web_sku
```
Failure → allocation created with `item_id = null`; "Unresolved SKU" badge shown in Orders tab;
staff can manually link via `sku_aliases` table.

---

## Cancellation Types

| Type | Detected by | Action |
|------|-------------|--------|
| Platform-detected pre-dispatch | Sync script (stale allocation + confirmed cancelled) | Release allocation; no `qty_on_hand` change |
| Staff-initiated pre-dispatch | Right-click in Daily Ops → "Cancel Order" | API cancel on platform + release allocation |
| Partial line cancellation | Sync detects removed/zeroed line | Release that line's allocation only |
| Post-dispatch parcel return | Staff → "Receive Return" | `qty_on_hand +qty` if resaleable; write `return` movement |
| Lost in transit | Staff → "Send Replacement" | `qty_on_hand −qty` again; write `online_sales` with `sale_type = 'replacement'`, `sale_price = 0` |
| Wrong item dispatched | Staff → "Report Wrong Item" | Two movements: `return` for wrong SKU + `dispatch` for correct SKU |

---

## Dropship Orders

Flag: `is_dropship = true` on allocation and sale rows.

- **Sync**: allocation row created but `qty_allocated_online` NOT incremented
- **Dispatch hook**: `qty_on_hand` NOT decremented; `online_sales` written with `is_dropship = true`

---

## Cross-Module Connections

| Connection | Direction |
|-----------|----------|
| `items.qty_allocated_online` | +qty by sync; −qty by dispatch hook or cancellation |
| `items.qty_on_hand` | −qty by dispatch hook only |
| `stock_movements` | `allocate_online` and `dispatch` records written |
| POS "Online" payment | Optionally releases `qty_allocated_online` when manually invoiced |
| Reporting Plan 08 | `online_sales` feeds channel performance and daily summary |
| Daily CSV | NOT modified by this module — the existing dispatch flow writes the CSV |

---

## Error Handling

- **Supabase unreachable during dispatch**: hook logs to `data/pending_sync.json`; replayed on next app start
- **GitHub Actions failure**: email alert via GitHub native failure notifications
- **Oversell**: `oversell_flag = true`; startup banner warns staff with oversold SKUs
