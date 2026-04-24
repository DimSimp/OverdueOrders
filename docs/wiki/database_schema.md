# Database Schema — Unified Reference

> **This is the canonical schema.** Where it differs from individual plan files, this document
> wins. See [index.md](index.md) for a list of conflict resolutions.
>
> All tables live in Supabase (PostgreSQL). Use `supabase-py` client in the desktop app.
> UUIDs are generated with `gen_random_uuid()`. Timestamps use `TIMESTAMPTZ`.

---

## Table of Contents

1. [Users & Auth](#1-users--auth)
2. [Suppliers](#2-suppliers)
3. [Inventory](#3-inventory)
4. [Purchase Orders & Invoices](#4-purchase-orders--invoices)
5. [Customers & Discounts](#5-customers--discounts)
6. [POS Transactions](#6-pos-transactions)
7. [Repairs](#7-repairs)
8. [Deposits](#8-deposits)
9. [Customer Special Orders](#9-customer-special-orders)
10. [Online Order Integration](#10-online-order-integration)
11. [Reporting](#11-reporting)
12. [Sequences & Counters](#12-sequences--counters)
13. [Entity Relationship Summary](#13-entity-relationship-summary)

---

## 1. Users & Auth

### `users`

> Migrated from `data/users.json`. Source of truth for authentication and role checks.
> Session heartbeats remain in-memory only — not written to Supabase.

| Column | Type | Notes |
|--------|------|-------|
| `id` | uuid PK DEFAULT gen_random_uuid() | |
| `username` | text UNIQUE NOT NULL | Login name |
| `first_name` | text NOT NULL | |
| `last_name` | text NOT NULL | |
| `password_hash` | text NOT NULL | PBKDF2-HMAC-SHA256, 200K iterations |
| `role` | text NOT NULL DEFAULT `'user'` | `'admin'` or `'user'` |
| `is_active` | boolean DEFAULT true | |
| `created_at` | timestamptz DEFAULT now() | |
| `created_by` | text | Username of admin who created the account |
| `last_login_at` | timestamptz | |

---

### `admin_overrides`

> Logged whenever a non-admin staff member has an action authorised by an admin's password.

| Column | Type | Notes |
|--------|------|-------|
| `id` | uuid PK | |
| `action` | text | e.g. `below_minimum_sell`, `excess_discount`, `cancel_order` |
| `description` | text | Human-readable context |
| `authorised_by` | text | Admin username |
| `requested_by` | text | Staff username who triggered the prompt |
| `reference_id` | text | Order ID, item SKU, transaction number, etc. |
| `created_at` | timestamptz DEFAULT now() | |

---

## 2. Suppliers

### `suppliers`

> `id` is the short code (e.g. `DUNLOP`), matching Musipos for compatibility.
> SKU suffix/prefix rules migrated here from `config.json`.

| Column | Type | Notes |
|--------|------|-------|
| `id` | text PK | Short supplier code e.g. `DUNLOP`, `ERNIE` |
| `name` | text NOT NULL | Full supplier name |
| `abn` | text | Australian Business Number |
| `account_number` | text | Our account number with this supplier |
| `payment_terms_days` | integer DEFAULT 30 | Days after invoice date that payment is due |
| `sku_suffix` | text | Appended to raw SKU to form `web_sku` |
| `sku_prefix` | text | Prepended to raw SKU |
| `character_substitutions` | jsonb | e.g. `{"/": ""}` — applied before suffix |
| `address_line1` | text | |
| `address_line2` | text | |
| `city` | text | |
| `state` | text | |
| `postcode` | text | |
| `notes` | text | Internal notes |
| `active` | boolean DEFAULT true | Inactive hidden from new PO creation |

---

### `supplier_contacts`

| Column | Type | Notes |
|--------|------|-------|
| `id` | uuid PK | |
| `supplier_id` | text FK → `suppliers.id` | |
| `role` | text | e.g. `Sales Rep`, `Support`, `Accounts` |
| `name` | text | |
| `email` | text | |
| `phone` | text | |
| `is_primary` | boolean DEFAULT false | |

---

## 3. Inventory

### `items` — Product master record

| Column | Type | Notes |
|--------|------|-------|
| `id` | uuid PK | |
| `sku` | text UNIQUE NOT NULL | Internal SKU (was `Supplier_Item_ID` in Musipos) |
| `web_sku` | text | SKU on Neto/eBay after suffix/prefix; defaults to `sku` |
| `title` | text NOT NULL | |
| `brand` | text | Was `Publisher_Brand` |
| `series` | text | Was `Artist_Composer_Series` |
| `instrument` | text | Top-level category |
| `sub_instrument` | text | e.g. "Guitar Strings" |
| `supplier_id` | text FK → `suppliers.id` | |
| `supplier_rrp` | numeric(10,2) | RRP from supplier |
| `last_purchase_cost` | numeric(10,2) | Cost on most recent invoice |
| `average_cost_exc_gst` | numeric(10,2) | Rolling weighted average, ex-GST |
| `average_cost_inc_gst` | numeric(10,2) | Rolling weighted average, inc-GST |
| `gst_amount` | numeric(10,2) | GST component |
| `minimum_sell` | numeric(10,2) | Floor price; `null` = no minimum |
| `online_sale_price` | numeric(10,2) | Price on Neto/eBay |
| `internal_barcode` | text | Musipos-generated barcode (`Barcode` column) |
| `product_barcode` | text | Manufacturer EAN/UPC |
| `pick_zone` | text | Shelf/location code |
| `qty_on_hand` | integer DEFAULT 0 | Physical units in store |
| `qty_allocated_online` | integer DEFAULT 0 | Committed to unfulfilled online orders |
| `qty_allocated_customer` | integer DEFAULT 0 | Reserved for named customers (CSO + deposits) |
| `qty_on_order` | integer DEFAULT 0 | On active POs not yet received |
| `min_order_level` | integer | Reorder trigger |
| `max_order_level` | integer | Maximum desired stock |
| `is_kit` | boolean DEFAULT false | Bundle of other items |
| `is_serialised` | boolean DEFAULT false | Track individual serial numbers |
| `stock_availability_from_supplier` | boolean | Supplier-reported availability |
| `description` | text | Long-form description |
| `weight_kg` | numeric(8,3) | |
| `length_cm` | numeric(8,2) | |
| `width_cm` | numeric(8,2) | |
| `height_cm` | numeric(8,2) | |
| `last_purchase_date` | date | |
| `last_sold_date` | date | |
| `created_date` | date DEFAULT CURRENT_DATE | |
| `active` | boolean DEFAULT true | Inactive = discontinued |

**Computed:**
```sql
-- Either a generated column or a view:
qty_available = qty_on_hand - qty_allocated_online - qty_allocated_customer
```

---

### `serial_numbers`

| Column | Type | Notes |
|--------|------|-------|
| `id` | uuid PK | |
| `item_id` | uuid FK → `items.id` | |
| `serial_number` | text NOT NULL | |
| `status` | text | `available`, `sold`, `allocated_online`, `allocated_customer` |
| `online_order_id` | text | Platform order ID if allocated/sold online |
| `customer_allocation_id` | uuid FK → `customer_allocations.id` | |
| `added_date` | date | Date received |
| `sold_date` | date | |

---

### `kit_components`

| Column | Type | Notes |
|--------|------|-------|
| `kit_id` | uuid FK → `items.id` | The bundle/kit |
| `component_id` | uuid FK → `items.id` | Individual item in the kit |
| `qty` | integer | Units of this component per kit |

Primary key: `(kit_id, component_id)`

---

### `item_images`

| Column | Type | Notes |
|--------|------|-------|
| `id` | uuid PK | |
| `item_id` | uuid FK → `items.id` | |
| `storage_path` | text | Supabase Storage path |
| `sort_order` | integer | 0 = primary image |
| `uploaded_at` | timestamptz | |

---

### `stock_movements` — Audit log of every stock change

| Column | Type | Notes |
|--------|------|-------|
| `id` | uuid PK | |
| `item_id` | uuid FK → `items.id` | |
| `movement_type` | text | See movement types below |
| `qty_change` | integer | Positive = stock in, negative = stock out |
| `reference_id` | text | Order ID, invoice number, transaction number, etc. |
| `notes` | text | |
| `performed_by` | text | Staff username |
| `performed_at` | timestamptz DEFAULT now() | |

**`movement_type` values:**

| Value | Meaning |
|-------|---------|
| `receive` | Stock received via supplier invoice |
| `sale_instore` | In-store POS sale (Standard, Invoice complete, Deposit collect) |
| `dispatch` | Online order dispatched |
| `allocate_online` | Online allocation created (+) or released (−) |
| `allocate_customer` | Customer allocation created (+) or released (−) |
| `adjustment` | Manual stock adjustment (damaged, lost, found) |
| `return` | Stock returned (online return, in-store refund) |
| `stocktake_zero` | Initial bulk zero-out |
| `stocktake_count` | Counted value recorded during stocktake |

---

## 4. Purchase Orders & Invoices

### `purchase_orders`

| Column | Type | Notes |
|--------|------|-------|
| `id` | uuid PK | |
| `po_number` | integer NOT NULL | Sequential per supplier |
| `supplier_id` | text FK → `suppliers.id` | |
| `status` | text | `open`, `pending`, `sent`, `complete` |
| `created_at` | timestamptz DEFAULT now() | When auto-created (first item added) |
| `finalised_at` | timestamptz | When moved to Pending |
| `sent_at` | timestamptz | When emailed to supplier |
| `pdf_path` | text | Local path to generated PO PDF |
| `notes` | text | |

**Constraint:** Only one `open` PO per `supplier_id` at a time (partial unique index).

---

### `purchase_order_lines`

| Column | Type | Notes |
|--------|------|-------|
| `id` | uuid PK | |
| `po_id` | uuid FK → `purchase_orders.id` | |
| `item_id` | uuid FK → `items.id` | |
| `sku` | text | Snapshot at time of ordering |
| `title` | text | Snapshot |
| `rrp` | numeric(10,2) | Snapshot of RRP |
| `min_sell` | numeric(10,2) | Snapshot of minimum sell |
| `qty_ordered` | integer | Qty requested |
| `qty_received` | integer DEFAULT 0 | Filled on invoice receipt |
| `qty_backordered` | integer DEFAULT 0 | Supplier could not supply |
| `unit_cost_inc_gst` | numeric(10,2) | Filled on invoice receipt |
| `unit_cost_exc_gst` | numeric(10,2) | `unit_cost_inc_gst / 1.1` |
| `line_total` | numeric(10,2) | `unit_cost_inc_gst × qty_received` |

---

### `invoices` — Supplier invoices (distinct from customer invoices)

| Column | Type | Notes |
|--------|------|-------|
| `id` | uuid PK | |
| `invoice_number` | text NOT NULL | Supplier-generated |
| `supplier_id` | text FK → `suppliers.id` | |
| `po_id` | uuid FK → `purchase_orders.id` | Nullable if no PO reference |
| `invoice_date` | date | |
| `due_date` | date | Default: `invoice_date + supplier.payment_terms_days`; editable |
| `subtotal_exc_gst` | numeric(10,2) | Product lines only |
| `gst` | numeric(10,2) | |
| `freight` | numeric(10,2) DEFAULT 0 | |
| `insurance` | numeric(10,2) DEFAULT 0 | |
| `total_inc_gst` | numeric(10,2) | |
| `status` | text DEFAULT `'unpaid'` | `unpaid`, `paid`, `overdue` |
| `received_at` | timestamptz DEFAULT now() | |
| `received_by` | text | Staff username |
| `entry_method` | text | `manual` or `ai` |

---

### `credit_notes` — Supplier-issued credits

| Column | Type | Notes |
|--------|------|-------|
| `id` | uuid PK | |
| `credit_note_number` | text NOT NULL | Supplier-issued reference |
| `supplier_id` | text FK → `suppliers.id` | |
| `invoice_id` | uuid FK → `invoices.id` | Original invoice being credited (nullable) |
| `credit_date` | date | |
| `amount_inc_gst` | numeric(10,2) | |
| `reason` | text | `Return`, `Pricing Error`, `Short Delivery` |
| `status` | text DEFAULT `'outstanding'` | `outstanding`, `applied` |
| `notes` | text | |
| `received_at` | timestamptz DEFAULT now() | |
| `received_by` | text | |

---

## 5. Customers & Discounts

### `discounts`

> Legacy/table-driven discount definition table from the original POS schema.
> It still exists in Supabase and is seeded with system presets, but the current customer/Till
> discount flow does not depend on it. The live implementation uses `customers.discount_profile`
> plus a hardcoded Till-side selector with the same values.

| Column | Type | Notes |
|--------|------|-------|
| `id` | uuid PK | |
| `name` | text NOT NULL | e.g. `"Teacher Discount"`, `"10%"` |
| `percentage` | numeric(5,2) NOT NULL | e.g. `15.00` |
| `is_system` | boolean DEFAULT false | System presets (10%–50%) cannot be deleted |
| `is_active` | boolean DEFAULT true | |
| `created_at` | timestamptz DEFAULT now() | |

**System presets** (seeded on first run, `is_system = true`): 10%, 20%, 30%, 40%, 50%.
These are retained for compatibility with the original schema, not because the current Till UI
reads from this table.

---

### `customers`

> Includes `musipos_account_code` and `musipos_barcode_ref` from Plan 09 import.
> `customer_id` is the human-readable sequential reference number — **not** used as FK target.
> All FKs to customers use `customers.id` (UUID).
>
> Current form validation requires `first_name` plus at least one of `mobile` or `phone_1`.

| Column | Type | Notes |
|--------|------|-------|
| `id` | uuid PK | Database FK target |
| `customer_id` | integer UNIQUE | Sequential display reference (e.g. `00142`) |
| `customer_barcode` | text UNIQUE | Code 128 barcode value from `customer_id` |
| `first_name` | text NOT NULL | |
| `surname` | text | |
| `business` | text | Company/school name |
| `mobile` | text | One of `mobile` / `phone_1` should be present |
| `phone_1` | text | Additional phone |
| `fax` | text | |
| `email` | text | |
| `website` | text | |
| `address_1` | text | |
| `address_2` | text | |
| `city` | text | |
| `state` | text | |
| `postcode` | text | |
| `country` | text DEFAULT `'Australia'` | |
| `ship_same_as_invoice` | boolean DEFAULT true | |
| `ship_address_1` | text | |
| `ship_address_2` | text | |
| `ship_city` | text | |
| `ship_state` | text | |
| `ship_postcode` | text | |
| `ship_country` | text | |
| `tax_exemption_number` | text | |
| `discount_id` | uuid FK → `discounts.id` | Legacy/table-driven discount link; not the active POS/customer discount path |
| `discount_profile` | text | Hardcoded profile name used by current POS flow: `5%`, `10%`, `15%`, `Teacher`, `Staff` |
| `terms_days` | integer | Payment terms for customer invoices |
| `credit_limit` | numeric(10,2) | |
| `stop_credit` | boolean DEFAULT false | |
| `is_local` | boolean DEFAULT false | |
| `abn` | text | |
| `newsletter_opt_in` | boolean DEFAULT false | |
| `private_comment` | text | Internal only |
| `statement_comment` | text | Printed on invoices |
| `active` | boolean DEFAULT true | |
| `created_at` | timestamptz DEFAULT now() | |
| `created_by` | text | Staff username |
| `musipos_account_code` | text | Original Musipos account code (import reference) |
| `musipos_barcode_ref` | text | Musipos internal barcode number (import reference) |

---

## 6. POS Transactions

> **Conflict resolved**: Plan 02 `transactions` is the canonical record for ALL POS sale types
> including Quote and Invoice. Plan 05's separate `quotes`, `quote_lines`, `customer_invoices`,
> and `customer_invoice_lines` tables are **not built** — the Customer Management module queries
> `transactions` filtered by `sale_type` and `customer_id` instead.
>
> Quote/Invoice lifecycle is tracked via `sale_status`. A `due_date` column is added for
> Invoice-type payment terms. `quote_number` and `invoice_number` sequences provide human-readable
> document numbers.

### `transactions`

| Column | Type | Notes |
|--------|------|-------|
| `id` | uuid PK | |
| `transaction_number` | text UNIQUE NOT NULL | e.g. `T-2026-0001` — from sequence |
| `quote_number` | integer | Set if `sale_type = 'quote'` |
| `invoice_number` | integer | Set if `sale_type = 'invoice'` |
| `sale_type` | text NOT NULL | `standard`, `quote`, `invoice`, `repair`, `deposit`, `refund` |
| `sale_status` | text NOT NULL DEFAULT `'completed'` | `draft`, `parked`, `pending_payment`, `completed`, `cancelled`, `voided` |
| `customer_id` | uuid FK → `customers.id` | Nullable for anonymous sales |
| `staff_id` | uuid FK → `users.id` | |
| `subtotal` | numeric(10,2) NOT NULL DEFAULT 0 | Before cart-level discount |
| `cart_discount_pct` | numeric(5,2) | Manual % applied to total |
| `cart_discount_total` | numeric(10,2) | Calculated discount amount |
| `override_total` | numeric(10,2) | "Total Sale Price" override — replaces subtotal |
| `total` | numeric(10,2) NOT NULL DEFAULT 0 | Final amount charged |
| `total_cost` | numeric(10,2) | Snapshot sum of cost prices at time of sale |
| `payment_cash` | numeric(10,2) DEFAULT 0 | |
| `payment_eft` | jsonb | Array of EFT entries: `[{"amount": 50.00}, ...]` |
| `payment_online` | numeric(10,2) DEFAULT 0 | Manually-invoiced online orders |
| `cash_tendered` | numeric(10,2) | Entered by cashier for cash payment |
| `change_given` | numeric(10,2) | `cash_tendered − (total − payment_eft − payment_online)` |
| `discount_id` | uuid FK → `discounts.id` | Legacy/table-driven preset link; current Till discount profiles do not rely on this field |
| `notes` | text | Transaction notes |
| `print_notes` | boolean DEFAULT false | Print notes on receipt |
| `due_date` | date | Invoice-type only: `created_at + customer.terms_days` |
| `payment_terms_days` | integer | Snapshot of terms at invoice creation |
| `linked_transaction_id` | uuid FK → `transactions.id` | Refund → original transaction |
| `park_name` | text | Name of parked transaction (while `sale_status = 'parked'`) |
| `cart_snapshot` | jsonb | Full cart state while parked |
| `created_at` | timestamptz DEFAULT now() | |
| `completed_at` | timestamptz | |

**`sale_status` lifecycle:**

| Status | Meaning |
|--------|---------|
| `draft` | Being built — not yet confirmed (Quote/Invoice not yet sent) |
| `parked` | Cart saved; `cart_snapshot` populated; screen cleared |
| `pending_payment` | Invoice sent but not yet paid |
| `completed` | Sale finalised and paid |
| `cancelled` | Cancelled before completion |
| `voided` | Admin-voided after completion (rare) |

---

### `transaction_lines`

| Column | Type | Notes |
|--------|------|-------|
| `id` | uuid PK | |
| `transaction_id` | uuid FK → `transactions.id` ON DELETE CASCADE | |
| `item_id` | uuid FK → `items.id` | Nullable for non-inventory lines |
| `sku` | text | Snapshot |
| `description` | text NOT NULL | Snapshot of item title |
| `qty` | numeric(10,3) NOT NULL | |
| `unit_price` | numeric(10,2) NOT NULL | Price at time of sale (may differ from current RRP) |
| `cost_price` | numeric(10,2) | Snapshot of `last_purchase_cost` at time of sale |
| `discount_pct` | numeric(5,2) DEFAULT 0 | Per-line discount |
| `line_total` | numeric(10,2) NOT NULL | `unit_price × qty × (1 − discount_pct/100)` |
| `line_margin_pct` | numeric(5,2) | `(unit_price − cost_price) / unit_price × 100` |
| `is_refunded` | boolean DEFAULT false | |
| `refunded_qty` | numeric(10,3) DEFAULT 0 | |

---

## 7. Repairs

> **Conflict resolved**: Plan 02 `repair_jobs` and Plan 05 `repairs`/`repair_lines` merged into
> a single `repairs` + `repair_lines` pair. All fields from both plans are included.
>
> Status values use Plan 02's detailed set. Plan 05's `ongoing` → `in_progress`,
> Plan 05's `complete` (done, not collected) → `ready`.

### `repairs`

| Column | Type | Notes |
|--------|------|-------|
| `id` | uuid PK | |
| `repair_number` | integer UNIQUE | Sequential, auto-assigned |
| `transaction_id` | uuid FK → `transactions.id` | The POS transaction that created this repair (nullable if raised outside POS) |
| `customer_id` | uuid NOT NULL FK → `customers.id` | |
| `instrument_description` | text NOT NULL | Free-text description of the instrument |
| `instrument_brand` | text | |
| `instrument_serial` | text | |
| `fault_description` | text | What is wrong with it |
| `estimated_cost` | numeric(10,2) | Repair estimate |
| `quote_approved` | boolean | Customer has approved the estimate |
| `assigned_to` | uuid FK → `users.id` | Technician |
| `intake_date` | date DEFAULT CURRENT_DATE | |
| `due_date` | date | Expected completion |
| `deposit_paid` | numeric(10,2) DEFAULT 0 | Deposit collected via POS on intake |
| `labour_charge` | numeric(10,2) DEFAULT 0 | Final labour charge |
| `status` | text NOT NULL DEFAULT `'intake'` | See status values below |
| `completion_notes` | text | Staff notes on what was done |
| `notes` | text | General internal notes |
| `created_at` | timestamptz DEFAULT now() | |
| `completed_at` | timestamptz | When repair work finished |
| `collected_at` | timestamptz | When customer collected |
| `created_by` | text | Staff username |

**Status values:**

| Status | Display colour | Meaning |
|--------|---------------|---------|
| `intake` | white | Just received, not yet started |
| `in_progress` | white | Being worked on |
| `awaiting_parts` | orange | On hold for parts |
| `ready` | yellow | Complete — awaiting customer collection |
| `collected` | green | Customer collected |
| `cancelled` | red | Cancelled |

---

### `repair_lines` — Parts and labour used in a repair

| Column | Type | Notes |
|--------|------|-------|
| `id` | uuid PK | |
| `repair_id` | uuid FK → `repairs.id` ON DELETE CASCADE | |
| `item_id` | uuid FK → `items.id` | Nullable for labour/misc lines |
| `sku` | text | |
| `title` | text | |
| `qty` | integer NOT NULL | |
| `unit_price` | numeric(10,2) NOT NULL | |
| `is_labour` | boolean DEFAULT false | |

---

## 8. Deposits

> **Conflict resolved**: Plan 02 and Plan 05 each defined a `deposits` table with overlapping
> but inconsistent columns. This merged definition includes all fields from both.
>
> Column name: `balance_due` (Plan 05 name; Plan 02 used `balance_owed`).
> Both `transaction_id` (Plan 02) and `po_id` (Plan 05) are included.

### `deposits`

| Column | Type | Notes |
|--------|------|-------|
| `id` | uuid PK | |
| `transaction_id` | uuid FK → `transactions.id` | The deposit payment transaction |
| `customer_id` | uuid NOT NULL FK → `customers.id` | |
| `item_id` | uuid FK → `items.id` | |
| `sku` | text | Snapshot |
| `title` | text | Snapshot of item name |
| `qty` | integer DEFAULT 1 | |
| `agreed_price` | numeric(10,2) | Total sale price agreed at deposit time |
| `deposit_amount` | numeric(10,2) NOT NULL | Amount paid upfront |
| `balance_due` | numeric(10,2) NOT NULL | `agreed_price − deposit_amount` |
| `deposit_type` | text NOT NULL | `layby` or `cso` |
| `allocation_id` | uuid FK → `customer_allocations.id` | Set if this deposit is for a CSO |
| `po_id` | uuid FK → `purchase_orders.id` | Set if item needs to be ordered for this customer |
| `status` | text DEFAULT `'active'` | `active`, `completed`, `cancelled` |
| `notes` | text | |
| `created_at` | timestamptz DEFAULT now() | |
| `completed_at` | timestamptz | When balance paid and item collected |
| `created_by` | text | Staff username |

---

## 9. Customer Special Orders

### `customer_allocations`

> **FK correction**: Plan 06 incorrectly typed `customer_id` as `integer FK → customers.customer_id`.
> Corrected to `uuid FK → customers.id` to match all other FKs.

| Column | Type | Notes |
|--------|------|-------|
| `id` | uuid PK | |
| `customer_id` | uuid NOT NULL FK → `customers.id` | ← corrected from Plan 06 |
| `item_id` | uuid FK → `items.id` | |
| `sku` | text NOT NULL | Denormalised for display |
| `description` | text | Item title at time of order |
| `qty` | integer NOT NULL | |
| `po_line_id` | uuid FK → `purchase_order_lines.id` | Linked PO line |
| `deposit_id` | uuid FK → `deposits.id` | If a deposit was paid |
| `status` | text NOT NULL DEFAULT `'on_order'` | See status lifecycle below |
| `notes` | text | |
| `created_at` | timestamptz DEFAULT now() | |
| `created_by` | text | |
| `notified_at` | timestamptz | When SMS was sent (null = not yet sent) |
| `collected_at` | timestamptz | |
| `collected_by` | text | |

**Status lifecycle:** `on_order` → `in_stock` → `collected` (or `cancelled`, `waiting_order`)

---

### `sms_log` (optional)

| Column | Type | Notes |
|--------|------|-------|
| `id` | uuid PK | |
| `customer_id` | uuid FK → `customers.id` | |
| `mobile` | text | Number dialled |
| `message` | text | Full message text |
| `sent_at` | timestamptz | |
| `status` | text | `sent`, `failed` |
| `error` | text | Error message if failed |

---

## 10. Online Order Integration

### `online_allocations`

| Column | Type | Notes |
|--------|------|-------|
| `id` | uuid PK | |
| `order_id` | text NOT NULL | Platform order identifier |
| `platform` | text NOT NULL | `neto` or `ebay` |
| `item_id` | uuid FK → `items.id` | Nullable if SKU not resolved |
| `web_sku` | text NOT NULL | SKU as on the platform |
| `sku` | text | Resolved internal SKU |
| `qty` | integer NOT NULL | |
| `order_status` | text | Last known platform status |
| `customer_name` | text | For display only |
| `oversell_flag` | boolean DEFAULT false | `qty_available < 0` after this allocation |
| `is_dropship` | boolean DEFAULT false | Supplier fulfils directly — no stock movement |
| `allocated_at` | timestamptz DEFAULT now() | |
| `last_synced_at` | timestamptz | Updated every sync run |

**Constraint:** `UNIQUE (order_id, web_sku)`

---

### `online_sales`

| Column | Type | Notes |
|--------|------|-------|
| `id` | uuid PK | |
| `order_id` | text NOT NULL | |
| `platform` | text NOT NULL | `neto` or `ebay` |
| `item_id` | uuid FK → `items.id` | |
| `sku` | text | |
| `web_sku` | text | |
| `qty` | integer NOT NULL | |
| `sale_price` | numeric(10,2) | Unit price at dispatch |
| `cost_at_dispatch` | numeric(10,2) | `last_purchase_cost` snapshot |
| `sale_type` | text DEFAULT `'standard'` | `standard` or `replacement` |
| `is_dropship` | boolean DEFAULT false | |
| `dispatched_at` | timestamptz | |
| `dispatched_by` | text | Staff username |
| `returned` | boolean DEFAULT false | |
| `returned_at` | timestamptz | |

---

### `sync_log`

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
| `error_message` | text | Null if clean run |
| `duration_ms` | integer | |

---

### `sku_aliases` (optional — for persistent manual SKU mappings)

| Column | Type | Notes |
|--------|------|-------|
| `web_sku` | text PK | Platform SKU that failed auto-resolution |
| `item_id` | uuid FK → `items.id` | Manually linked item |
| `created_at` | timestamptz | |
| `created_by` | text | |

---

## 11. Reporting

### `daily_summaries`

| Column | Type | Notes |
|--------|------|-------|
| `id` | uuid PK | |
| `summary_date` | date UNIQUE | Business date this summary covers |
| `opening_float` | numeric(10,2) | Cash at start of day |
| `closing_float_expected` | numeric(10,2) | `opening_float + cash_sales` |
| `closing_float_actual` | numeric(10,2) | Staff-counted cash |
| `float_variance` | numeric(10,2) | `actual − expected` |
| `total_sales_instore` | numeric(10,2) | |
| `total_sales_neto` | numeric(10,2) | |
| `total_sales_ebay` | numeric(10,2) | |
| `total_sales_all` | numeric(10,2) | |
| `total_gst_collected` | numeric(10,2) | |
| `total_transactions` | integer | POS transaction count |
| `generated_by` | text | Staff username |
| `generated_at` | timestamptz | |
| `pdf_path` | text | Local path to the locked PDF |
| `notes` | text | Variance explanation |

---

## 12. Sequences & Counters

All sequential human-readable numbers use PostgreSQL sequences:

```sql
-- Transaction numbers: T-2026-0001
CREATE SEQUENCE transaction_number_seq START 1;
-- Usage: 'T-' || to_char(NOW(), 'YYYY') || '-' || LPAD(nextval(...)::TEXT, 4, '0')

-- Quote numbers: Q-0001
CREATE SEQUENCE quote_number_seq START 1;

-- Invoice numbers (customer): INV-0001
CREATE SEQUENCE invoice_number_seq START 1;

-- Repair numbers: REP-0001
CREATE SEQUENCE repair_number_seq START 1;

-- Deposit reference numbers: DEP-0001
CREATE SEQUENCE deposit_number_seq START 1;

-- Customer IDs: 00001 (human-readable only — DB PK is UUID)
CREATE SEQUENCE customer_id_seq START 1;

-- PO numbers are per-supplier (derived from MAX query, not a global sequence)
```

---

## 13. Entity Relationship Summary

```
users ─────────────────── admin_overrides
  │                       transactions.staff_id
  │
suppliers ─────────────── supplier_contacts
  │                       purchase_orders.supplier_id
  │                       invoices.supplier_id
  │                       items.supplier_id
  │
items ─────────────────── serial_numbers
  │                   ├── kit_components
  │                   ├── item_images
  │                   └── stock_movements
  │
purchase_orders ────────── purchase_order_lines.po_id
  │                        invoices.po_id
  │                        credit_notes.invoice_id (via invoices)
  │
customers ─────────────── transactions.customer_id
  │                   ├── repairs.customer_id
  │                   ├── deposits.customer_id
  │                   └── customer_allocations.customer_id
  │
discounts ─────────────── customers.discount_id
  │                       transactions.discount_id
  │
transactions ──────────── transaction_lines.transaction_id
  │                   ├── repairs.transaction_id
  │                   └── deposits.transaction_id
  │
customer_allocations ───── deposits.allocation_id
  │                        purchase_order_lines.id (via po_line_id)
  │
online_allocations ─────── (items.qty_allocated_online)
online_sales ──────────── (stock_movements ref)
```

---

*Last updated: 2026-04-23 — documented current customer/Till discount flow (`discount_profile`), invoice/shipping customer address fields, and the present role of legacy `discounts` / `discount_id` columns in the schema.*
