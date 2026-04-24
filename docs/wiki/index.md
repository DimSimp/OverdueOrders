# Scarlett AIO — POS & Inventory System Wiki

> **Canonical reference for all planning, schema, and cross-module behaviour.**
> Individual plan files in `docs/plans/` remain for historical context and detailed implementation
> checklists. This wiki supersedes them on schema and cross-module definitions.

---

## Contents

| Page | Description |
|------|-------------|
| **[This page]** | Overview, module map, build phases |
| [database_schema.md](database_schema.md) | All Supabase tables — unified, conflict-resolved |
| [module_inventory.md](module_inventory.md) | Inventory system |
| [module_pos.md](module_pos.md) | POS / Till |
| [module_suppliers.md](module_suppliers.md) | Supplier management |
| [module_purchasing.md](module_purchasing.md) | Purchase orders & receiving |
| [module_customers.md](module_customers.md) | Customer management |
| [module_cso.md](module_cso.md) | Customer special orders |
| [module_online.md](module_online.md) | Online order integration |
| [module_reporting.md](module_reporting.md) | Reporting & daily close |
| [module_import.md](module_import.md) | Musipos / APIC import |
| [module_users.md](module_users.md) | Staff & user management |

---

## Project Overview

Scarlett AIO is a Windows desktop application built in Python (CustomTkinter) for Scarlett Music.
This POS and inventory expansion replaces Musipos as the primary stock system while keeping the
existing Neto/eBay dispatch workflow intact. All new persistent data lives in Supabase (cloud
PostgreSQL), making it accessible from the desktop app, the GitHub Actions sync script, and any
future integrations.

Daily Operations and Afternoon Operations remain the live workflows for overdue orders, dispatch,
freight, and invoice-import tasks, while the newer POS, Customers, and Inventory modules continue
to expand inside the tabbed POS window.

**Guiding principles:**
- Single source of truth for stock across all channels (in-store, Neto, eBay)
- Every stock movement is traceable via `stock_movements`
- Search-first UI — no bulk-load screens; all results are paginated
- Musipos continues running for its APIC pipeline; the new system imports from it once and runs alongside

---

## Technology Stack

| Layer | Technology |
|-------|-----------|
| Desktop app | Python 3.9, CustomTkinter |
| Database | Supabase (PostgreSQL) — `supabase-py` client |
| File storage | Supabase Storage (item images) |
| PDF generation | reportlab (receipts, PO PDFs, reports, daily summary) |
| Background sync | `sync/order_sync.py` on GitHub Actions (every 5 min) |
| SMS notifications | TextMagic REST API v2 |
| Invoice parsing | Claude API (`claude-sonnet-4-20250514`) — ported from Web Portal |
| Online platforms | Neto (`POST /do/WS/NetoAPI`) + eBay Fulfillment API |
| Build | PyInstaller (onedir, windowed) via `build.bat` |

---

## Module Map

```
                    ┌─────────────────┐
                    │  Users (Plan 10)│ — auth + permissions
                    └────────┬────────┘
                             │ used by all modules
          ┌──────────────────┼──────────────────────┐
          ▼                  ▼                       ▼
┌──────────────────┐  ┌────────────┐  ┌─────────────────────┐
│ Suppliers (P03)  │  │ Customers  │  │  Inventory (P01)    │
│ + Purchasing(P04)│  │  (P05)     │  │  stock master        │
└────────┬─────────┘  └─────┬──────┘  └──────────┬──────────┘
         │                  │                      │
         │      ┌───────────┼──────────────────────┤
         │      ▼           ▼                      ▼
         │  ┌───────────────────────┐  ┌───────────────────┐
         │  │   POS / Till (P02)    │  │  Online Orders    │
         │  │   transactions        │  │  (P07) — sync     │
         │  └──────────┬────────────┘  │  script + hook    │
         │             │               └────────┬──────────┘
         │      ┌──────┘                        │
         ▼      ▼                               ▼
    ┌────────────────────────────────────────────────────┐
    │                  Reporting (P08)                   │
    │  daily_summaries · Z-report · 13 report catalogue  │
    └────────────────────────────────────────────────────┘

 CSO (P06) ─── depends on: Inventory + Customers + Purchasing
 Import (P09) ─ one-time: populates Suppliers + Inventory + Customers
```

---

## Module Dependencies

| Module | Depends On | Required By |
|--------|-----------|-------------|
| **Inventory** | Suppliers | POS, Purchasing, Online, Reporting |
| **POS / Till** | Inventory, Customers, Users | Reporting |
| **Suppliers** | — | Inventory, Purchasing, Import |
| **Purchasing** | Inventory, Suppliers, CSO | Reporting |
| **Customers** | — | POS, CSO, Reporting |
| **CSO** | Customers, Inventory, Purchasing | POS, Reporting |
| **Online Integration** | Inventory | Reporting |
| **Reporting** | POS, Online, Inventory, Customers, Purchasing | — |
| **Import** | Inventory, Suppliers, Customers | — (utility) |
| **Users** | — | All modules |

---

## Home Screen

After login, the home screen presents workflow buttons in this order:

| # | Button | Who sees it | Launches |
|---|--------|------------|---------|
| 1 | **POS** | Dev admin only during construction; all staff once released | POS window (tab-based, full-screen) |
| 2 | **Daily Operations** | All logged-in users | `DailyOpsWindow` |
| 3 | **Afternoon Operations** | All logged-in users | Afternoon ops workflow |

**Dev-only flag**: `config.json → pos_dev_user` (username string). When set, the POS button is only
rendered for that user. Set to `null` or remove the key to release to all staff.
See [module_pos.md — Development Access](module_pos.md#development-access).

---

## Build Phases

### Phase 1 — Foundation
Supabase infrastructure, initial data.

- [x] Supabase project setup — project `scarlett-aio` created (ap-southeast-2); credentials in `config.json`
- [x] Create all Phase 1 tables — `docs/sql/01_create_tables.sql` run and verified; also includes Phase 3 POS core tables
- [x] `supabase-py` added to `requirements.txt`; `src/supabase_client.py` created (lazy singleton)
- [~] `users` table migration — **deferred indefinitely**. Existing JSON-based auth (`src/user_manager.py`) is kept as-is. POS records staff as `performed_by text` (username string), same pattern as `stock_movements`. `transactions.staff_id` left nullable.
- [ ] Inventory system UI (Plan 01)
- [ ] Supplier management UI (Plan 03) — schema needed before inventory
- [ ] Musipos import wizard (Plan 09) — populates initial data

### Phase 2 — Online Bridge
Keep existing dispatch workflow alive while adding stock sync.

- [ ] `online_allocations` + `online_sales` + `sync_log` tables
- [ ] `sync/order_sync.py` + GitHub Actions workflow
- [ ] Dispatch hook in `src/session.py` / `src/session_daily.py`

### Phase 3 — POS Core
The in-store transaction system.

- [x] `discounts` table — created and seeded from the original schema design; the current POS/customer discount flow now uses hardcoded `discount_profile` values on customer records plus a Till-side selector
- [x] `transactions` + `transaction_lines` tables — created (repairs + deposits deferred)
- [x] POS button added to home screen (`pos_dev_user` gated); home screen updated to order: POS → Daily → Afternoon
- [~] POS/Till (Plan 02) — `PosWindow` + `TillTab` (`src/gui/pos/`). Cart fully functional: SKU/barcode lookup → Supabase `items` (exact + case-insensitive; 1 match → add to cart, 0/many → inventory fuzzy search); spinner animation on Add button during lookup; rescan increments qty; inline editing of Qty/Unit Price/Disc %/Line Total (single-click overlay); bidirectional Line Total ↔ Disc % back-calculation; cart-level TOTAL override; Margin column (green > 10% / orange = 10% / red < 10%) per line + Cart Margin summary; out-of-stock modal (Allow/Remove); serialised item warning. Right-click context menu on cart rows: Show in Inventory (auto-selects item in inventory detail panel), Remove from Cart. Cross-tab wiring: Inventory tab → Add to Cart; "Show in Inventory" → auto-selects via `_auto_select` flag in `InventoryTab.search_and_focus()`. Centralised `ttk.Style` in `src/gui/styles.py`. **Right panel layout** now includes linked customer details plus transaction notes/receipt-print controls above the cart breakdown. **Payment flow complete**: always-visible 3-column inline panel (EFT | Cash | Online); EFT capped at 3 rows; Cash shows live Change hint; live "Remaining / Change due / Paid in full" status label. `confirm_standard_sale()` in `src/pos/transaction_client.py` writes `transactions` + `transaction_lines`, atomically decrements `qty_on_hand` via `adjust_item_qty` RPC, writes `sale_instore` stock movements. Transaction numbers auto-assigned by `BEFORE INSERT` trigger (`T-2026-NNNN`, concurrent-safe). Dev data reset: `docs/sql/reset_dev_transactions.sql`. **Park & Recall**: `park_transaction()` snapshots cart → `parked` row with `cart_snapshot JSONB`; `RecallDialog` (700×460) lists parked transactions with Delete button; recalling a transaction auto-deletes it from parked list; completes as a new standard sale. **Receipt PDF**: `src/pos/receipt_generator.py` (reportlab, 80mm thermal, Code128 barcode, PyMuPDF content trim); post-sale `_ReceiptDialog` prompt; reprint available from Daily Sales. **Customer integration now live**: customer lookup/attach in Till, customer details and transaction notes on receipts, right-click `Load in Till` from the Customers tab, customer `discount_profile` auto-applied on attach, and a manual Till discount dropdown beside Sale Type that overrides the linked customer discount while selected. `Teacher` currently uses a placeholder 15% discount; `Staff` prices from cost with a rounded-up 10% margin target.
- [x] Inventory tab — `InventoryTab` live in POS window (`src/gui/inventory/`). Search/filter/paginate grid; full detail panel on row select. Musipos CSV importer (`src/inventory/importer.py`) with cp1252 encoding, date/price parsing, cross-supplier SKU deduplication; bulk upserted 105K+ items to Supabase with 0 errors.
- [~] Customer management UI (Plan 05) — Customers tab is live in the POS window with search, pagination, customer detail panel, sale history, CSV import, create/edit modal, invoice/shipping addresses, profile discount selection, and right-click `Load in Till`. Quotes, Invoices, Repairs, and Deposits tabs are present as placeholders and still need full implementation.
- [ ] `repairs` + `deposits` tables
- [x] Receipt PDF generation — `src/pos/receipt_generator.py`; see POS/Till entry above.

### Phase 4 — Operations
Purchasing, special orders, notifications.

- [ ] Purchase orders & receiving (Plan 04)
- [ ] `purchase_orders` + `purchase_order_lines` + `invoices` + `credit_notes` tables
- [ ] Customer special orders (Plan 06) + `customer_allocations`
- [ ] TextMagic SMS (`src/sms_client.py`)
- [ ] Admin override mechanism (Plan 10)

### Phase 5 — Reporting
Daily close and report catalogue.

- [ ] `daily_summaries` table
- [ ] Z-report / Daily Sales close process
- [ ] All 13 reports (Plan 08)

---

## Key Cross-Module Flows

### Stock Decrements — where and when `qty_on_hand` changes

| Event | Module | How |
|-------|--------|-----|
| Order dispatched (online) | Online (Plan 07) | Dispatch hook |
| Item sold in-store | POS (Plan 02) | `confirm_sale()` for Standard, Invoice (complete), Deposit (collect) |
| Refund processed | POS (Plan 02) | Refund flow increments `qty_on_hand` |
| Manual stock adjustment | Inventory (Plan 01) | Admin-only Adjust Stock |
| Stocktake zero-out | Inventory (Plan 01) | Admin-only bulk set to 0 |

### Customer Allocation Lifecycle (`qty_allocated_customer`)

| Event | Change | Module |
|-------|--------|--------|
| CSO raised (item added to PO) | +qty | CSO (Plan 06) |
| Deposit taken on item | +qty | POS/Deposits |
| Customer invoice Open/Sent | +qty | POS (Invoice type) |
| CSO collected / deposit completed | −qty | POS/CSO |
| Customer invoice Completed | −qty (and −qty_on_hand) | POS |
| CSO / deposit / invoice Cancelled | −qty | POS/CSO |

### Online Allocation Lifecycle (`qty_allocated_online`)

| Event | Change | Module |
|-------|--------|--------|
| New order detected | +qty | Online sync script |
| Order cancelled (pre-dispatch) | −qty | Online sync / Daily Ops |
| Order dispatched | −qty (and −qty_on_hand) | Dispatch hook |

### `stock_movements` record types

All stock changes must write a `stock_movements` record:

| type | Description |
|------|-------------|
| `receive` | Stock received via supplier invoice |
| `sale_instore` | In-store POS sale |
| `sale_online` | Online order dispatched |
| `allocate_online` | Online allocation created or released (+ or −) |
| `allocate_customer` | Customer allocation created or released (+ or −) |
| `dispatch` | Online order dispatched (duplicate of sale_online? No — use `dispatch` for online; `sale_instore` for POS) |
| `adjustment` | Manual stock adjustment (Adjust Stock, damaged, lost) |
| `return` | Stock returned (online or in-store refund) |
| `stocktake_zero` | Initial bulk zero-out |
| `stocktake_count` | Counted value recorded during stocktake |

---

## Schema Conflict Resolutions

When this wiki was written, the following conflicts between plan files were resolved. The
[database_schema.md](database_schema.md) reflects these resolutions and supersedes the individual
plans.

| Conflict | Resolution |
|----------|-----------|
| Plan 02 `repair_jobs` vs Plan 05 `repairs`/`repair_lines` | Merged into single `repairs` + `repair_lines` tables. Status values use Plan 02's detailed set. |
| Plan 02 `deposits` vs Plan 05 `deposits` (different columns) | Merged into single `deposits` table with all columns from both. |
| Plan 02 `preset_discounts` vs Plan 05 `discounts` | Same concept — one `discounts` table. Column: `percentage` (Plan 05 name). |
| Plan 05 separate `quotes`/`customer_invoices` tables vs Plan 02 `sale_type` in `transactions` | **`transactions` is canonical.** Customer module views query `transactions WHERE sale_type = 'quote'`. Plan 05's dedicated tables are not built. |
| Plan 06 `customer_allocations.customer_id integer` | Changed to `uuid FK → customers.id` — consistent with all other FKs. |
| Plan 08 references `pos_transactions` | Table is `transactions` — Plan 08 reference corrected. |

---

*Last updated: 2026-04-23 — customer profile integration is live in Till; receipts can include linked customer details and transaction notes; Customers tab can load a customer directly into the active Till transaction; customer profile discounts and manual Till discounts are implemented; customer create/edit modal now includes separate invoice and shipping addresses plus profile discount selection.*
