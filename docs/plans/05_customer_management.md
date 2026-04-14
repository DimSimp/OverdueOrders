# Plan 05 — Customer Management

> **Part of**: [Master Plan](00_overview.md)
> **Status**: 🔲 Not started
> **Phase**: 3 — POS Core

---

## Overview

Manages all customer records and their associated history. Follows the same search-first, paginated pattern as the inventory module. The customer profile is the central hub linking quotes, customer invoices, repairs, supplier PO allocations, audit history, and deposits — all accessible via a tabbed detail panel.

> **Naming note**: This module uses `customer_invoices` to distinguish from the supplier `invoices` table in plan 04. These are two entirely separate concepts.

---

## Key Concepts

| Concept | Description |
|---------|-------------|
| **Customer ID** | Sequential integer, auto-assigned on profile creation. Human-readable, printed on documents, and encoded as a scannable barcode. |
| **Customer Barcode** | Barcode generated from the Customer ID. Printed on invoices and receipts for quick POS lookup via scanner. |
| **Discount** | A named percentage discount defined in app settings. Attached to a customer profile and auto-applied at POS checkout. |
| **Customer Invoice** | A sale document sent to a customer with payment terms and bank details. Primarily used for business accounts. Distinct from supplier invoices. |
| **Quote** | A price proposal generated at the POS without completing a sale. Can be converted to a customer invoice or completed as a regular sale. |
| **Repair** | A customer instrument left in-store for servicing. Tracked separately with its own status lifecycle. |
| **Deposit** | A partial payment against an item being held or special ordered. Allocates the item in inventory until the sale is completed. |

---

## Database Schema

### `discounts` — Named discount definitions (managed in Settings)

| Column | Type | Notes |
|--------|------|-------|
| `id` | uuid PK | |
| `name` | text NOT NULL | e.g. `"Teacher Discount"`, `"10%"` |
| `percentage` | numeric(5,2) NOT NULL | e.g. `15.00` |
| `is_system` | boolean DEFAULT false | System presets (10%, 20%, etc.) cannot be deleted |
| `created_at` | timestamptz DEFAULT now() | |

**System presets** (created on first run, not deletable): 10%, 20%, 30%, 40%, 50%.
Custom discounts are added via app settings and appear in the same dropdown everywhere discounts are used.

---

### `customers` — Customer master record

| Column | Type | Notes |
|--------|------|-------|
| `id` | uuid PK | |
| `customer_id` | integer UNIQUE | Sequential, auto-assigned. Human-readable reference number. |
| `customer_barcode` | text UNIQUE | Barcode value derived from `customer_id`. Scannable at POS. |
| `first_name` | text NOT NULL | Required |
| `surname` | text | |
| `business` | text | Company/school name |
| `mobile` | text NOT NULL | Required |
| `phone_1` | text | Additional phone |
| `fax` | text | |
| `email` | text | |
| `website` | text | |
| `address_1` | text | Invoice/billing address |
| `address_2` | text | |
| `city` | text | |
| `state` | text | |
| `postcode` | text | |
| `country` | text DEFAULT `'Australia'` | |
| `ship_same_as_invoice` | boolean DEFAULT true | If true, ship-to mirrors invoice address |
| `ship_address_1` | text | |
| `ship_address_2` | text | |
| `ship_city` | text | |
| `ship_state` | text | |
| `ship_postcode` | text | |
| `ship_country` | text | |
| `tax_exemption_number` | text | |
| `discount_id` | uuid FK → `discounts.id` | Auto-applied at POS when customer is loaded |
| `terms_days` | integer | Payment terms: 7, 14, 30, or 60 days |
| `credit_limit` | numeric(10,2) | Maximum outstanding balance allowed |
| `stop_credit` | boolean DEFAULT false | Prevents further credit sales if true |
| `is_local` | boolean DEFAULT false | |
| `abn` | text | Australian Business Number |
| `newsletter_opt_in` | boolean DEFAULT false | |
| `private_comment` | text | Internal only — never appears on customer-facing documents |
| `statement_comment` | text | Printed on every invoice sent to this customer |
| `active` | boolean DEFAULT true | |
| `created_at` | timestamptz DEFAULT now() | |
| `created_by` | text | Staff user |

---

### `quotes` — Price proposals from the POS

| Column | Type | Notes |
|--------|------|-------|
| `id` | uuid PK | |
| `quote_number` | integer UNIQUE | Sequential, auto-assigned |
| `customer_id` | uuid FK → `customers.id` | Nullable — quotes can exist without a customer |
| `status` | text DEFAULT `'open'` | `open`, `sent`, `converted`, `expired` |
| `subtotal_exc_gst` | numeric(10,2) | |
| `gst` | numeric(10,2) | |
| `total_inc_gst` | numeric(10,2) | |
| `total_cost` | numeric(10,2) | Used to calculate margin |
| `notes` | text | |
| `created_at` | timestamptz DEFAULT now() | |
| `expires_at` | timestamptz | Optional expiry |
| `created_by` | text | |

### `quote_lines`

| Column | Type | Notes |
|--------|------|-------|
| `id` | uuid PK | |
| `quote_id` | uuid FK → `quotes.id` | |
| `item_id` | uuid FK → `items.id` | |
| `sku` | text | Snapshot |
| `title` | text | Snapshot |
| `qty` | integer | |
| `unit_price` | numeric(10,2) | Price at time of quote |
| `unit_cost` | numeric(10,2) | Cost at time of quote (for margin) |
| `discount_pct` | numeric(5,2) DEFAULT 0 | |
| `line_total` | numeric(10,2) | `unit_price × qty × (1 − discount_pct/100)` |

---

### `customer_invoices` — Invoices sent to customers (distinct from supplier invoices)

| Column | Type | Notes |
|--------|------|-------|
| `id` | uuid PK | |
| `invoice_number` | integer UNIQUE | Sequential, auto-assigned |
| `customer_id` | uuid FK → `customers.id` | |
| `quote_id` | uuid FK → `quotes.id` | Nullable — set if generated from a quote |
| `status` | text DEFAULT `'open'` | `open`, `sent`, `complete`, `cancelled` |
| `payment_terms_days` | integer | Snapshot of customer's terms at time of creation |
| `due_date` | date | `created_at + payment_terms_days` |
| `subtotal_exc_gst` | numeric(10,2) | |
| `gst` | numeric(10,2) | |
| `total_inc_gst` | numeric(10,2) | |
| `total_cost` | numeric(10,2) | For margin |
| `notes` | text | |
| `created_at` | timestamptz DEFAULT now() | |
| `sent_at` | timestamptz | |
| `completed_at` | timestamptz | |
| `created_by` | text | |

### `customer_invoice_lines`

| Column | Type | Notes |
|--------|------|-------|
| `id` | uuid PK | |
| `invoice_id` | uuid FK → `customer_invoices.id` | |
| `item_id` | uuid FK → `items.id` | |
| `sku` | text | Snapshot |
| `title` | text | Snapshot |
| `qty` | integer | |
| `unit_price` | numeric(10,2) | |
| `unit_cost` | numeric(10,2) | |
| `discount_pct` | numeric(5,2) DEFAULT 0 | |
| `line_total` | numeric(10,2) | |

**Inventory allocation**: Items on an `open` or `sent` customer invoice are allocated in inventory (`qty_allocated_customer` incremented). They are not removed from `qty_on_hand` until the invoice is marked `complete`.

---

### `repairs` — Customer instrument repairs

| Column | Type | Notes |
|--------|------|-------|
| `id` | uuid PK | |
| `repair_number` | integer UNIQUE | Sequential, auto-assigned |
| `customer_id` | uuid FK → `customers.id` | |
| `instrument_brand` | text | |
| `instrument_serial` | text | |
| `instrument_description` | text | Free-text description of the instrument |
| `status` | text DEFAULT `'ongoing'` | `ongoing`, `complete`, `collected`, `cancelled` |
| `labour_charge` | numeric(10,2) DEFAULT 0 | |
| `notes` | text | |
| `created_at` | timestamptz DEFAULT now() | |
| `completed_at` | timestamptz | |
| `collected_at` | timestamptz | |
| `created_by` | text | |

### `repair_lines` — Parts and labour used in a repair

| Column | Type | Notes |
|--------|------|-------|
| `id` | uuid PK | |
| `repair_id` | uuid FK → `repairs.id` | |
| `item_id` | uuid FK → `items.id` | Nullable for labour lines |
| `sku` | text | |
| `title` | text | |
| `qty` | integer | |
| `unit_price` | numeric(10,2) | |
| `is_labour` | boolean DEFAULT false | True for the labour line |

---

### `deposits` — Partial payments holding items for a customer

| Column | Type | Notes |
|--------|------|-------|
| `id` | uuid PK | |
| `customer_id` | uuid FK → `customers.id` | |
| `item_id` | uuid FK → `items.id` | |
| `sku` | text | Snapshot |
| `title` | text | Snapshot |
| `qty` | integer | |
| `agreed_price` | numeric(10,2) | Total sale price agreed at time of deposit |
| `deposit_amount` | numeric(10,2) | Amount paid upfront |
| `balance_due` | numeric(10,2) | `agreed_price − deposit_amount` |
| `status` | text DEFAULT `'active'` | `active`, `completed`, `cancelled` |
| `po_id` | uuid FK → `purchase_orders.id` | Nullable — set if item needs to be ordered |
| `notes` | text | |
| `created_at` | timestamptz DEFAULT now() | |
| `completed_at` | timestamptz | When balance paid and item collected |
| `created_by` | text | |

**Inventory allocation**: Creating a deposit increments `items.qty_allocated_customer`. Completing or cancelling a deposit decrements it (completing also decrements `qty_on_hand`).

---

## UI Design

### Customer List View

```
┌──────────────────────────────────────────────────────────────┐
│  [Search bar]                              [+ New Customer]  │
├──────────┬───────────────────┬────────────┬──────────┬───────┤
│ Cust. ID │ Company           │ First Name │ Surname  │ City  │ Phone |
├──────────┼───────────────────┼────────────┼──────────┼───────┤
│ 00142    │ Riverside Primary │ Sarah      │ Nguyen   │ ...   │ ...   │
│ 00143    │ —                 │ James      │ Porter   │ ...   │ ...   │
└──────────┴───────────────────┴────────────┴──────────┴───────┘
```

- Search-first — no results until a query is entered (min 2 chars, 300ms debounce)
- Searches across: Customer ID, First Name, Surname, Business, Mobile, Phone, Email
- Pagination: 100 rows per page
- Clicking a row shows the detail panel below
- `[+ New Customer]` opens the New Customer form

---

### New Customer Form

Opens as a resizable dialog. Two sections separated by a divider:

**Section 1 — General Details** *(minimum required: First Name + Mobile)*
- First Name *, Surname, Business
- Mobile *, Phone 1, Email, Website, Fax
- Invoice Address: Address 1, Address 2, City, State, Postcode, Country
- Ship-to Address: [✓ Same as invoice address] toggle — if unchecked, shows duplicate address fields

**Section 2 — Account Details** *(all optional)*
- Discount (dropdown of all defined discounts — system presets + custom)
- Terms (dropdown: 7 / 14 / 30 / 60 days)
- Credit Limit, Stop Credit (toggle)
- Tax Exemption #, ABN
- Local (toggle), Newsletter Opt-in (toggle)
- Private Comment, Statement Comment
- Fax (if not entered above)

On save:
- `customer_id` auto-assigned (next sequential integer)
- `customer_barcode` generated from `customer_id`
- Record written to Supabase

---

### Customer Detail Panel (Bottom Sub-Window)

Seven tabs. The selected customer's name and ID are shown in the panel header.

---

#### Tab 1: Customer Info
All customer fields displayed in a two-column read-only layout. **[Edit]** button opens the same form used for creation, pre-populated. Shows the customer barcode as a scannable image with the numeric value below it.

---

#### Tab 2: Quotes

```
┌──────────┬───────┬──────────┬──────────┬────────┬──────────┐
│ Quote #  │  QTY  │  Total   │  Cost    │ Margin │  Status  │
├──────────┼───────┼──────────┼──────────┼────────┼──────────┤
│ Q-00841  │   3   │ $245.00  │ $145.00  │  41%   │  Sent    │
└──────────┴───────┴──────────┴──────────┴────────┴──────────┘
```

- Right-click → context menu:
  - **Email Quote** — sends PDF quote to customer's email
  - **Open in POS** — loads quote items into active cart to complete as a sale
  - **Convert to Invoice** — creates a `customer_invoice` from this quote
  - **Mark Expired**
- Cost and Margin columns are role-restricted (manager/admin only)

---

#### Tab 3: Invoices

```
┌───────────┬───────┬──────────┬────────────┬──────────────┬──────────┐
│ Invoice # │  QTY  │  Total   │  Due Date  │ Terms        │ Status   │
├───────────┼───────┼──────────┼────────────┼──────────────┼──────────┤
│ INV-0412  │   5   │ $620.00  │ 2026-04-30 │ 30 days      │ Sent     │
└───────────┴───────┴──────────┴────────────┴──────────────┴──────────┘
```

- Status colours: Open (white), Sent (blue), Complete (green), Cancelled (red), Overdue (orange)
- Items on Open/Sent invoices are **allocated** in inventory; removed from `qty_on_hand` only on Complete
- Overdue invoices (due date passed, status not Complete) flagged with same daily startup popup as supplier invoices
- Right-click → context menu:
  - **Email Invoice** — sends PDF to customer email
  - **Edit** — only available while `open`
  - **Mark Sent** / **Mark Complete** / **Cancel**
  - **View PDF**
- Cancelling sets status to `cancelled` — record is never deleted

---

#### Tab 4: Repair

```
┌───────────┬──────────────────────┬────────────┬──────────────┐
│ Repair #  │ Instrument           │  Created   │   Status     │
├───────────┼──────────────────────┼────────────┼──────────────┤
│ REP-0091  │ Fender Stratocaster  │ 2026-03-10 │  Complete    │ ← yellow text
│ REP-0104  │ Gibson Les Paul      │ 2026-04-01 │  Ongoing     │ ← white text
└───────────┴──────────────────────┴────────────┴──────────────┘
```

**Status colours:**
- `Ongoing` — white
- `Complete` — yellow (repair done, awaiting collection)
- `Collected` — green
- `Cancelled` — red

Right-click → context menu:
- **Mark Complete** — sets status to `complete`
- **Open in POS** — loads repair into POS to process payment; on completion sets status to `collected`
- **Cancel** — sets status to `cancelled`
- **Print Receipt** — reprints the original repair receipt

> **Cross-reference**: A dedicated Repairs module (plan 11) will provide a global view of all repairs across all customers, with workflow management for staff. The repair tab here is customer-scoped only.

---

#### Tab 5: PO

Items on active supplier purchase orders that have been assigned to this customer. Data sourced from plan 06 (Customer Special Orders).

```
┌──────────┬──────────────┬──────────────────────────┬──────────┐
│  PO #    │   SKU        │  Description             │ Arrived  │
├──────────┼──────────────┼──────────────────────────┼──────────┤
│  4068    │ D10           │ D'Addario XL .010 Set   │   No     │
│  4055    │ GCDPCK        │ Guitar Pack Complete    │   Yes    │
└──────────┴──────────────┴──────────────────────────┴──────────┘
```

- "Arrived" = Yes when the PO's invoice has been received and the SKU was on it
- Right-click → **Remove Allocation** (removes customer tag from PO item) | **Go to PO** (opens the PO in Suppliers module)

---

#### Tab 6: Audit

Full transaction history for this customer, combining all activity types.

```
┌────────────┬──────────────┬──────────────────────────┬────────────┐
│  Date      │  Type        │  Reference               │  Amount    │
├────────────┼──────────────┼──────────────────────────┼────────────┤
│ 2026-04-01 │ Invoice      │ INV-0412                 │ $620.00    │
│ 2026-03-28 │ Sale         │ #10841 (in-store)        │ $45.00     │
│ 2026-03-10 │ Repair       │ REP-0091                 │ $80.00     │
│ 2026-03-01 │ Deposit      │ DEP-0021 — GCDPCK        │ $50.00     │
└────────────┴──────────────┴────────────┴─────────────┴────────────┘
```

- **Filters**: Date range, Type (Sale / Return / Invoice / Quote / Repair / Deposit / All)
- **Export PDF** button — generates a report of the filtered view, useful for customer tax/insurance requests
- Clicking a row navigates to the relevant record (opens the invoice, repair, etc.)

---

#### Tab 7: Deposit

```
┌──────────┬──────────────┬──────────────────────┬──────────┬──────────┬──────────┐
│ Dep. #   │  SKU         │ Description          │  Price   │ Deposit  │ Balance  │ Status  │
├──────────┼──────────────┼──────────────────────┼──────────┼──────────┼──────────┤
│ DEP-0021 │ GCDPCK       │ Guitar Pack Complete │ $199.00  │ $50.00   │ $149.00  │ Active  │
└──────────┴──────────────┴──────────────────────┴──────────┴──────────┴──────────┘
```

- Item allocated in inventory until deposit is `completed` or `cancelled`
- Right-click → **Complete Sale** (opens POS with item + balance due pre-loaded) | **Cancel** | **View PO** (if item was ordered)
- Status colours: Active (white), Completed (green), Cancelled (red)

---

### Deposits — POS Flow

At POS checkout, "Deposit" is a sale type option alongside "Sale", "Quote", "Invoice", "Repair":

1. Staff selects **Deposit** as sale type
2. A deposit percentage dropdown appears: **10% / 25% / 50% / 100%** — or staff enters a custom dollar amount
3. The calculated deposit amount is shown; staff processes payment for that amount only
4. On completion:
   - If no customer is loaded → prompt: "Add a customer to continue with deposit"
   - Deposit record written to Supabase; `items.qty_allocated_customer` incremented
   - Prompt: "Would you like to add this item to a Purchase Order?" (Yes → opens Add to PO flow; No → skip)
   - Receipt printed with repair number, item details, deposit paid, balance due

---

## Search, Filter & Sort

- **Search**: Customer ID (exact), First Name, Surname, Business, Mobile, Phone 1, Email
- **Filter**: Active / Inactive / All (default: Active)
- **Sort**: Any column header (Customer ID default, ascending)

---

## Implementation Checklist

### Infrastructure
- [ ] Create `discounts` table in Supabase with system presets
- [ ] Create `customers` table in Supabase
- [ ] Create `quotes` + `quote_lines` tables
- [ ] Create `customer_invoices` + `customer_invoice_lines` tables
- [ ] Create `repairs` + `repair_lines` tables
- [ ] Create `deposits` table
- [ ] Auto-increment trigger or MAX query for `customer_id`
- [ ] Barcode value generation on customer creation (e.g. Code 128 from `customer_id`)

### Data Layer (`src/customers/`)
- [ ] `customer_client.py`
  - [ ] `search_customers(query, filters, page)`
  - [ ] `get_customer(customer_id)`
  - [ ] `create_customer(data)` → auto-assign ID + barcode
  - [ ] `update_customer(id, data)`
  - [ ] `get_quotes(customer_id)`
  - [ ] `get_customer_invoices(customer_id)`
  - [ ] `get_repairs(customer_id)`
  - [ ] `get_po_items(customer_id)` → from customer_allocations (plan 06)
  - [ ] `get_audit_log(customer_id, filters)`
  - [ ] `get_deposits(customer_id)`
- [ ] `discount_client.py`
  - [ ] `get_all_discounts()`
  - [ ] `create_discount(name, percentage)`
  - [ ] `delete_discount(id)` — blocked if `is_system = true`
- [ ] `deposit_client.py`
  - [ ] `create_deposit(data)` → increments `qty_allocated_customer`
  - [ ] `complete_deposit(id)` → decrements `qty_allocated_customer` + `qty_on_hand`
  - [ ] `cancel_deposit(id)` → decrements `qty_allocated_customer`

### GUI (`src/gui/customers/`)
- [ ] `customer_list_view.py` — search-first paginated list
- [ ] `customer_form.py` — new/edit form (two sections, required field validation)
- [ ] `customer_detail_panel.py` — bottom panel with 7-tab bar
- [ ] `tabs/customer_info_tab.py` — read-only display + barcode image + Edit button
- [ ] `tabs/quotes_tab.py` — quote list + right-click menu
- [ ] `tabs/invoices_tab.py` — invoice list + status colours + right-click menu
- [ ] `tabs/repairs_tab.py` — repair list + status colours + right-click menu
- [ ] `tabs/po_tab.py` — PO items list (data from plan 06)
- [ ] `tabs/audit_tab.py` — filterable transaction history + PDF export
- [ ] `tabs/deposits_tab.py` — deposit list + status colours + right-click menu

### Overdue Customer Invoice Flagging
- [ ] Add customer invoices to the daily startup overdue check (plan 04 background task)
- [ ] Overdue = `due_date < today` and `status = 'sent'`
- [ ] Include in the startup popup: *"X customer invoice(s) overdue — $Y outstanding"*

### Inventory Integration
- [ ] Customer invoice Open/Sent → increment `qty_allocated_customer`
- [ ] Customer invoice Complete → decrement `qty_allocated_customer`, decrement `qty_on_hand`
- [ ] Customer invoice Cancelled → decrement `qty_allocated_customer`
- [ ] Deposit created → increment `qty_allocated_customer`
- [ ] Deposit completed → decrement `qty_allocated_customer`, decrement `qty_on_hand`
- [ ] Deposit cancelled → decrement `qty_allocated_customer`

### Settings Integration
- [ ] Discounts management screen in app settings: list, add, delete custom discounts
- [ ] System presets shown as read-only (not deletable)

---

## Open Questions / Future Considerations

- **Repairs module (plan 11)**: A global repairs view across all customers is planned as a separate module — queue management, status board, staff assignment. Deferred until customer management is live.
- **Customer statement generation**: Some business customers may require a periodic statement showing all outstanding invoices. A PDF statement (filter: unpaid invoices for a customer within a date range) could be generated from the Invoices tab. Deferred.
- **Loyalty / points system**: Not in scope currently. The discount system covers the most common case.
- **Customer barcode format**: Code 128 is recommended (compact, widely supported). The exact format (e.g. zero-padded to 8 digits: `00000142`) should be confirmed before implementation to ensure scanner compatibility.
- **Duplicate detection**: When creating a customer, the system should warn if a similar name + phone already exists, to reduce duplicate records. Simple fuzzy match on first name + mobile before save.

---

*Last updated: 2026-04-02*
