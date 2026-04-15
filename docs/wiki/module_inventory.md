# Module: Inventory System

> **Detail plan**: [docs/plans/01_inventory_system.md](../plans/01_inventory_system.md)
> **Build phase**: 1 — Foundation
> **Tables owned**: `items`, `serial_numbers`, `kit_components`, `item_images`, `stock_movements`

---

## Overview

The inventory system is the single source of truth for all product data and stock levels across
in-store and online channels. All other modules that touch stock read from and write to `items`.
No module other than the dispatch hook and POS confirm-sale should decrement `qty_on_hand`
directly — all other changes go through defined hooks and write a corresponding `stock_movements`
record.

---

## Stock Level Fields (on `items`)

| Field | Meaning |
|-------|---------|
| `qty_on_hand` | Physical units in store right now |
| `qty_allocated_online` | Units committed to unfulfilled online orders |
| `qty_allocated_customer` | Units committed to CSOs, deposits, customer invoices |
| `qty_on_order` | Units on active POs not yet received |
| `qty_available` | `qty_on_hand − qty_allocated_online − qty_allocated_customer` (computed) |

---

## Who Changes What

| Field | Who increments | Who decrements |
|-------|---------------|----------------|
| `qty_on_hand` | Invoice receive (Plan 04), In-store refund (Plan 02) | Dispatch hook (Plan 07), POS sale (Plan 02), Manual adjustment |
| `qty_allocated_online` | Sync script new order | Dispatch hook, Cancellation (sync or staff) |
| `qty_allocated_customer` | CSO creation (Plan 06), Deposit taken (Plan 02), Customer invoice open/sent (Plan 02) | CSO collected/cancelled, Deposit completed/cancelled, Invoice completed/cancelled |
| `qty_on_order` | Add item to PO (Plan 04) | Invoice receive (Plan 04), CSO cancellation (Plan 06) |

**Rule**: Every change to any of these fields **must** write a `stock_movements` record.

---

## UI

Split-panel layout: search/filter grid (top), item detail panel with 5 tabs (bottom).

**Detail tabs**: Details · Customer Allocations · Sale History · Order/Receiving History · Specs/Images

Right-click context menu on item row exposes:
- **Add to PO** (Plan 04)
- **Add to Customer Order** (Plan 06)
- **Add to POS** (Plan 02 — when POS is open)

**Admin-only actions**: Edit item, Adjust Stock, Stocktake zero-out.

---

## Cross-Module Connections

| Event | Triggered by | What inventory does |
|-------|-------------|---------------------|
| New item sold at POS | Plan 02 | `qty_on_hand −qty`, write `sale_instore` movement |
| Invoice received | Plan 04 | `qty_on_hand +qty`, update costs, write `receive` movement |
| Online order synced | Plan 07 sync | `qty_allocated_online +qty`, write `allocate_online` movement |
| Online order dispatched | Plan 07 hook | `qty_on_hand −qty`, `qty_allocated_online −qty`, write `dispatch` movement |
| CSO raised | Plan 06 | `qty_on_order +qty`, `qty_allocated_customer +qty` |
| CSO collected | Plan 06 / 02 | `qty_allocated_customer −qty` |
| Deposit taken | Plan 02 | `qty_allocated_customer +qty` |
| Deposit completed | Plan 02 | `qty_allocated_customer −qty`, `qty_on_hand −qty` |
| In-store refund | Plan 02 | `qty_on_hand +qty`, write `return` movement |

---

## Stocktake Flow

1. Admin triggers "Prepare for Stocktake" — bulk zeros all `qty_on_hand`, writes `stocktake_zero` movements
2. "Stocktake in progress" banner shown in inventory window
3. Staff enter counts per-item (Adjust Stock button, reason = `Count Correction`) or import CSV (SKU + qty)
4. Admin marks stocktake complete — writes `stocktake_count` records for changed items, clears banner

---

## Key Constraints

- `items.sku` is UNIQUE — duplicate SKUs rejected on import and manual creation
- `qty_allocated_online` and `qty_allocated_customer` must never go negative — the app enforces this in logic (Supabase does not enforce via constraint since concurrent updates could race)
- `minimum_sell = null` means no minimum; `0` is treated as null on import
