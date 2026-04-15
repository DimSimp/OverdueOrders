# Plan 02 — POS / Till

## Overview

The POS/Till is the primary day-to-day interface for all in-store transactions. It handles
standard sales, quotes, invoices, repairs, deposits/laybys, and refunds. It integrates with
inventory, customer profiles, the CSO system (Plan 06), reporting (Plan 08), and the online
order workflow (Plan 07).

Every transaction produces a printable receipt with a unique transaction number and scannable
barcode, and optionally emails the receipt to the customer.

---

## Screen Layout

```
┌──────────────────────────────────────────────────────────────────────────┐
│  TOP BAR                                                                  │
│  [Customer: None ▼]  [Sale Type: Standard ▼]  T-2026-0001  [Notes]       │
│  [Park]  [Recall ▼]  [Sales History]                                      │
├─────────────────────────────────────────────────────┬────────────────────┤
│  CART                                               │  PAYMENT PANEL     │
│  SKU/Barcode │ Desc │ QTY │ Price │ Disc% │ Total │ Margin │             │
│  ──────────────────────────────────────────────     │  Running Total      │
│  [scan/type] │      │  1  │       │       │       │        │  $0.00      │
│  ...                                                │                    │
│                                                     │  Discounts         │
│                                                     │  [__ %] [Override] │
│                                                     │  [Preset ▼]        │
│                                                     │                    │
│                                                     │  Cart Margin: --%  │
│                                                     │                    │
│                                                     │  ── Payment ──     │
│                                                     │  [Cash] [EFT]      │
│                                                     │  [Online]          │
│                                                     │                    │
│                                                     │  Remaining: $0.00  │
│                                                     │  Change:    $0.00  │
│                                                     │                    │
│                                                     │  [CONFIRM SALE]    │
└─────────────────────────────────────────────────────┴────────────────────┘
```

The payment panel sits on the **right side**. This keeps the cart area wide and matches the
standard modern POS layout; it also works better on widescreen monitors at the counter.

---

## Top Bar

| Element | Behaviour |
|---|---|
| Customer button | Opens Customer Management tab; selecting a customer attaches them to the transaction without clearing the cart. Displays customer name once attached. |
| Sale Type dropdown | Standard Sale / Quote / Invoice / Repair / Deposit / Refund. Changing type never clears the cart. Quote, Invoice, Repair, Deposit require a customer before completion. |
| Transaction number | Auto-assigned sequential ID (e.g. `T-2026-0001`). Shown read-only. Printed on receipt. |
| Notes button | Opens a small text editor. Toggle: "Print on receipt" / "Internal only". |
| Park button | Saves current cart state and clears the screen for a new transaction. |
| Recall dropdown | Lists all parked transactions by name/number. Selecting one restores full cart state. |
| Sales History button | Opens the Sales History panel (see below). |

---

## Cart — Item Entry

### Columns (left to right)

| Column | Notes |
|---|---|
| SKU / Barcode | Primary input. Accepts product barcode or SKU. |
| Description | Auto-filled from inventory. Read-only. |
| QTY | Auto-filled to 1. Editable. Re-scanning same barcode increments QTY. |
| Unit Price | Auto-filled from inventory. Editable (overrides for this transaction only — does not update inventory). |
| Disc % | Manual entry. Updates line total live. |
| Total | `unit_price × qty × (1 − disc/100)`. Read-only, calculated. |
| Margin % | `(unit_price − cost_price) / unit_price × 100`. Staff-only display. Never appears on receipts. |

### Input behaviour

- **Barcode scan**: Triggers immediate exact-match lookup. On match, populates the row. On
  re-scan of same barcode on the active line, increments QTY. On no match, shows an error
  notification on that row.
- **Manual SKU entry + Return**: Triggers exact-match lookup. On no match, shows an error
  popup with an **OK** button. Clicking OK navigates to the Inventory module with the unmatched
  SKU pre-loaded as a fuzzy search term. From Inventory, right-clicking an item exposes
  **"Add to POS"**, which appends the item to the active cart.
- **Return key** (when not in SKU field of an existing row): Creates a new blank row.
- **Disc % edit**: Updates Total live.
- **Unit Price edit**: Updates Total and Margin live.

### Row actions

Right-clicking a line item exposes: Remove Line, Duplicate Line.

---

## Payment Panel

### Running Total

Recalculates on every cart or discount change. This is the canonical amount owed.

### Discounts Section

Three independent controls — only one should be applied at a time; applying a second clears
the others:

| Control | Behaviour |
|---|---|
| Manual Disc % | Entered as a percentage. Reduces Running Total proportionally. |
| Total Sale Price | Override field. Typing a dollar amount replaces the Running Total. |
| Preset dropdown | Selects a pre-configured discount (e.g. Teacher, Student, 5%, 10%). Populated from `preset_discounts` table. Applies a percentage to the Running Total. |

Discounts apply to the **cart total** (not individual lines). Per-line discounts are applied
in the cart grid and are separate.

### Cart Margin Display

Shows blended margin across all lines, recalculated live using snapshot cost prices:

```
Cart Margin: 34.2%  ($42.80 profit)
```

Not printed on any document.

### Payment Methods

Clicking a payment method button adds a payment entry. Multiple payment entries can coexist
(split payments).

#### Cash

Clicking **Cash** opens a modal:

```
Cash Tendered: [ $_____ ]
[Confirm]
```

Confirming records `payment_cash` and calculates **Change Given**:
`change = cash_tendered − remaining_amount`

The **Change Given** field in the panel updates to show the change to hand back. This field
only displays a value when cash is used.

#### EFT

Clicking **EFT** appends an EFT entry field. Staff enters the amount for that tap/swipe.
Multiple EFT entries are tracked individually in `payment_eft` (JSON array). The panel shows:

```
EFT:  $50.00  [×]
EFT:  $50.00  [×]
EFT total: $100.00
```

Each `[×]` removes that entry and recalculates. This supports customers making multiple
card payments in one transaction.

#### Online

Clicking **Online** opens a modal:

```
Online Payment Amount: [ $_____ ]

This will record a manually invoiced online order.
Release allocated stock for this amount? [Yes] [No]
```

- **Yes**: Decrements `qty_allocated_online` for any matching online allocation row. Prevents
  double-deduction if the order was already dispatched via the app.
- **No**: Records the sale without touching allocations.

#### Amount Remaining

Live display:
```
Remaining: $23.50
```

Updates as payment entries are added. Turns green when Remaining = $0.00. Confirm Sale is
disabled until Remaining ≤ $0.00 (overpayment is allowed for cash change scenarios).

---

## Sale Types

### Standard Sale

Default flow as described above. Customer optional. Stock decremented on Confirm.

---

### Quote

- Customer required before confirming.
- Items, prices, and discounts recorded but **no stock movement**.
- Saved with `status = 'quote'`, `quote_status = 'draft'`.
- Appears in Sales History with a **[Promote to Invoice]** button.
- Quote PDF layout uses "Quote" header and omits payment section.
- Quotes can be declined (soft-archived) or converted to Invoice.

**Promotion flow**: Quote → Invoice does not require re-entry; the existing cart is carried over.

---

### Invoice

- Customer required.
- Represents goods taken on credit — stock **is decremented** on confirmation.
- Saved with `invoice_status = 'unpaid'`.
- Appears in **Invoices Outstanding** report (Plan 08) with aging.
- Staff can recall the Invoice from Sales History and process payment against it, which updates
  `invoice_status = 'paid'` and records the payment breakdown.
- Invoice PDF layout uses "Tax Invoice" header and shows "Balance Due: $X.XX".

**Promotion flow**: Invoice → paid when payment is processed against it.

---

### Repair

- Customer required.
- Opening a Repair sale shows a **Repair Details modal** alongside the cart (or as a first
  step before the cart):

| Field | Type | Notes |
|---|---|---|
| Item Description | Text | What was brought in (e.g. "Yamaha P-125 keyboard") |
| Fault Description | Text | What is wrong with it |
| Estimated Cost | Currency | Staff estimate; shown on quote approval prompt |
| Quote Approved | Checkbox | Customer has approved the estimated cost |
| Assigned Technician | Dropdown | From user list (role = any) |
| Intake Date | Date | Auto-filled to today |
| Due Date | Date | Expected completion |

- Creates a `repair_jobs` record on save.
- Optional: charge a deposit in the same transaction (amount goes to `payment_cash`/EFT; the
  repair job tracks `deposit_paid`).
- Appears in **Outstanding Repairs** report until status = `collected` or `cancelled`.
- Status flow: `intake` → `in_progress` → `awaiting_parts` → `ready` → `collected`.
- Staff updates status from the Outstanding Repairs report — not the POS.

Cart lines can include repair parts/labour as line items. These are for receipt purposes; the
repair job record is the canonical source of truth for the repairs report.

---

### Deposit / Layby

- Customer required.
- Used for both laybys (in-stock item reserved) and CSOs (item on order, Plan 06).
- Opening a Deposit sale shows a **Deposit Details section**:

| Field | Notes |
|---|---|
| Deposit Type | Layby / Customer Special Order |
| Linked Allocation | (CSO only) Dropdown to select an open `customer_allocations` record for this customer |
| Deposit Amount | Amount being paid now |
| Balance Owed | Auto-calculated: `item_total − deposit_amount` |

- On confirm, creates a `deposits` record and sets `qty_on_hold += qty` for the item (layby).
  For CSO, updates the `customer_allocations.deposit_id`.
- Balance is collected when the customer returns: recall the deposit from Sales History →
  **[Complete Layby / Collect]** → processes remaining payment → creates a Standard Sale
  record → decrements `qty_on_hold`.

---

### Refund

Selecting Refund shows an additional header row in the cart:

```
[Scan receipt barcode]  OR  Transaction #: [______]  [Load]
```

**Loading a transaction**: Populates the cart with all lines from the original transaction.
Staff can remove lines that are not being refunded. All amounts display as negative.

**Manual refund**: Staff types items directly — amounts are negative, stock is added back.

On confirm:
- Line totals are negative (reduces the running total).
- `qty_on_hand` is incremented for each returned line (unless it was a non-inventory item).
- `payment_cash`/`payment_eft` records a negative (refund) payment.
- If loaded from a transaction, the original transaction's `linked_refund_id` is updated.

---

## Sales History

Accessible via the **Sales History** button in the top bar. Opens as a panel or modal.

**Columns**: Transaction #, Date/Time, Customer, Sale Type, Total, Status, Staff

**Filters**: Date range picker (defaults to today).

**Row actions** (right-click or buttons):
- **Re-print Receipt** — regenerates and sends to printer.
- **Email Receipt** — prompts for email if not already on customer profile.
- **Start Refund** — pre-loads the transaction into a new Refund sale.
- **View Details** — read-only expanded view of the transaction.

Sales History updates after every completed transaction (no manual refresh needed).

---

## Parked Transactions

- Multiple transactions can be parked simultaneously.
- On park, staff optionally names the park (e.g. "John - keyboard"), otherwise auto-names
  as "Park 1", "Park 2", etc.
- Parked state is stored as `status = 'parked'` in `transactions` with a `cart_snapshot`
  JSON column preserving all line items, discounts, customer, and payment entries.
- Recalling restores the entire cart state exactly.
- Parked transactions appear in the **Recall** dropdown in the top bar.
- A parked transaction can be discarded via the Recall panel.

---

## Confirm Sale Flow

1. Staff clicks **Confirm Sale**.
2. If the sale type requires a customer and none is attached: block with a prompt.
3. If Remaining > $0.00: block with a prompt.
4. **Receipt prompt**:
   ```
   Transaction complete — T-2026-0001
   [Print Receipt]  [Email Receipt]  [Both]  [Skip]
   ```
5. If **Email Receipt** or **Both** is selected:
   - If customer has an email: use it.
   - If customer has no email: "No email on file. Enter email: [______]". Confirming saves
     the email to the customer profile.
   - If no customer on the transaction: "Enter email to send receipt: [______]". Does not
     create a customer profile.
6. Stock decrements, transaction record written to Supabase.
7. Cart clears, transaction number increments.

---

## Receipt Design

```
SCARLETT MUSIC
123 Example Street, City
Ph: (03) 1234 5678

Tax Invoice                    T-2026-0001
Date: 15/04/2026 14:32         Staff: Jane

Customer: John Smith
         john@example.com

─────────────────────────────────────────
ITEM                     QTY  PRICE  TOTAL
Yamaha P-125              1  $999.00  $999.00
Roland FP-30 (10% disc)   1  $765.00  $765.00
─────────────────────────────────────────
                              SUBTOTAL $1764.00
                              DISCOUNT   $85.00
                              TOTAL    $1679.00

Payment: Cash $100.00 | EFT $1579.00
Change Given: $0.00

Notes: Includes sustain pedal as per quote T-2026-0998

─────────────────────────────────────────
         [BARCODE: T-2026-0001]
     Thank you for shopping with us!
```

Notes section only prints if "Print on receipt" is toggled.
Margin column is **never** printed.
Barcode encodes the transaction number (Code 128 format) for refund scanning.

---

## Database Schema

### `transactions`

```sql
CREATE TABLE transactions (
    id                  UUID PRIMARY KEY DEFAULT gen_random_uuid(),
    transaction_number  TEXT UNIQUE NOT NULL,  -- e.g. T-2026-0001
    sale_type           TEXT NOT NULL,         -- standard/quote/invoice/repair/deposit/refund
    status              TEXT NOT NULL DEFAULT 'completed',
                                               -- draft/parked/completed/voided
    customer_id         UUID REFERENCES customers(id),
    staff_id            UUID REFERENCES users(id),

    -- Financials
    subtotal            NUMERIC(10,2) NOT NULL DEFAULT 0,
    cart_discount_pct   NUMERIC(5,2),
    cart_discount_total NUMERIC(10,2),
    override_total      NUMERIC(10,2),         -- Total Sale Price override
    total               NUMERIC(10,2) NOT NULL DEFAULT 0,
    total_cost          NUMERIC(10,2),         -- snapshot sum of cost prices

    -- Payment
    payment_cash        NUMERIC(10,2) DEFAULT 0,
    payment_eft         JSONB,                 -- [{amount: 50.00}, {amount: 50.00}]
    payment_online      NUMERIC(10,2) DEFAULT 0,
    cash_tendered       NUMERIC(10,2),
    change_given        NUMERIC(10,2),

    -- Preset discount
    preset_discount_id  UUID REFERENCES preset_discounts(id),

    -- Notes
    notes               TEXT,
    print_notes         BOOLEAN DEFAULT FALSE,

    -- Quote/Invoice lifecycle
    quote_status        TEXT,                  -- draft/approved/declined/converted
    invoice_status      TEXT,                  -- unpaid/paid/partial

    -- Refund linkage
    linked_transaction_id UUID REFERENCES transactions(id),

    -- Parked state
    park_name           TEXT,
    cart_snapshot       JSONB,                 -- full cart state while parked

    created_at          TIMESTAMPTZ DEFAULT NOW(),
    completed_at        TIMESTAMPTZ
);
```

### `transaction_lines`

```sql
CREATE TABLE transaction_lines (
    id              UUID PRIMARY KEY DEFAULT gen_random_uuid(),
    transaction_id  UUID NOT NULL REFERENCES transactions(id) ON DELETE CASCADE,
    item_id         UUID REFERENCES items(id),  -- nullable for non-inventory lines
    sku             TEXT,
    description     TEXT NOT NULL,
    qty             NUMERIC(10,3) NOT NULL,
    unit_price      NUMERIC(10,2) NOT NULL,
    cost_price      NUMERIC(10,2),              -- snapshot at time of sale
    discount_pct    NUMERIC(5,2) DEFAULT 0,
    line_total      NUMERIC(10,2) NOT NULL,
    line_margin_pct NUMERIC(5,2),
    is_refunded     BOOLEAN DEFAULT FALSE,
    refunded_qty    NUMERIC(10,3) DEFAULT 0
);
```

### `repair_jobs`

```sql
CREATE TABLE repair_jobs (
    id                  UUID PRIMARY KEY DEFAULT gen_random_uuid(),
    transaction_id      UUID REFERENCES transactions(id),
    customer_id         UUID NOT NULL REFERENCES customers(id),
    item_description    TEXT NOT NULL,
    fault_description   TEXT,
    estimated_cost      NUMERIC(10,2),
    quote_approved      BOOLEAN,
    assigned_to         UUID REFERENCES users(id),
    intake_date         DATE NOT NULL DEFAULT CURRENT_DATE,
    due_date            DATE,
    deposit_paid        NUMERIC(10,2) DEFAULT 0,
    status              TEXT NOT NULL DEFAULT 'intake',
                        -- intake/in_progress/awaiting_parts/ready/collected/cancelled
    completion_notes    TEXT,
    created_at          TIMESTAMPTZ DEFAULT NOW(),
    updated_at          TIMESTAMPTZ
);
```

### `deposits`

```sql
CREATE TABLE deposits (
    id              UUID PRIMARY KEY DEFAULT gen_random_uuid(),
    transaction_id  UUID NOT NULL REFERENCES transactions(id),
    customer_id     UUID NOT NULL REFERENCES customers(id),
    item_id         UUID REFERENCES items(id),
    sku             TEXT,
    description     TEXT,
    deposit_type    TEXT NOT NULL,              -- layby/cso
    allocation_id   UUID REFERENCES customer_allocations(id),
    deposit_amount  NUMERIC(10,2) NOT NULL,
    balance_owed    NUMERIC(10,2) NOT NULL,
    status          TEXT NOT NULL DEFAULT 'active',
                                                -- active/collected/cancelled/expired
    created_at      TIMESTAMPTZ DEFAULT NOW(),
    collected_at    TIMESTAMPTZ
);
```

### `preset_discounts`

```sql
CREATE TABLE preset_discounts (
    id           UUID PRIMARY KEY DEFAULT gen_random_uuid(),
    name         TEXT NOT NULL,        -- e.g. "Teacher", "Staff", "10%"
    discount_pct NUMERIC(5,2) NOT NULL,
    is_active    BOOLEAN DEFAULT TRUE
);
```

### `transaction_number_seq`

Use a Supabase/PostgreSQL sequence to guarantee unique, sequential transaction numbers:

```sql
CREATE SEQUENCE transaction_number_seq START 1;

-- On insert, generate transaction_number:
-- 'T-' || to_char(NOW(), 'YYYY') || '-' || LPAD(nextval('transaction_number_seq')::TEXT, 4, '0')
```

---

## Key Integrations

| System | Integration |
|---|---|
| Inventory | SKU/barcode lookup; `qty_on_hand` decremented on Standard Sale, Invoice, Deposit completion; incremented on Refund |
| Customer Profiles | Attached to transaction; email stored/updated on receipt send; transaction history saved |
| Customer Allocations (Plan 06) | Deposit sale links to `customer_allocations.deposit_id` |
| Online Allocations (Plan 07) | Online payment method optionally releases `qty_allocated_online` |
| Outstanding Repairs (Plan 08) | Repair jobs surfaced in reporting |
| Invoices Outstanding (Plan 08) | Invoice-type transactions appear in AR aging report |
| Z-Report / Daily Sales (Plan 08) | `transactions` table is the source; completed transactions feed daily totals |
| TextMagic SMS (Plan 06) | Not triggered from POS directly — triggered by invoice receipt in supplier workflow |

---

## New Source Files

```
src/gui/pos/
    pos_window.py         — Main POS window; top bar, cart, payment panel layout
    cart_widget.py        — Editable item grid with scan/type input handling
    payment_panel.py      — Running total, discount controls, payment method buttons
    cash_modal.py         — Cash tendered modal
    eft_panel.py          — EFT entry list with running total
    confirm_modal.py      — Post-sale receipt prompt
    sales_history.py      — Sales history panel with filters and row actions
    park_manager.py       — Park/recall logic
    refund_loader.py      — Load transaction by number or barcode scan
    repair_modal.py       — Repair job detail capture
    deposit_panel.py      — Deposit/layby detail section

src/pos_manager.py        — Business logic: SKU lookup, stock movement, transaction writes
src/receipt_generator.py  — PDF receipt generation (reportlab); barcode rendering
```

---

## Implementation Checklist

### Phase 1 — Core Sale
- [ ] `transactions` + `transaction_lines` tables + sequence in Supabase
- [ ] `pos_manager.py`: SKU/barcode exact lookup, row population
- [ ] `cart_widget.py`: editable grid, scan input, QTY increment on re-scan, Return key = new row
- [ ] Running total + live margin calculation
- [ ] Payment panel: Cash modal + Change Given
- [ ] EFT panel: multiple entries, running total, Amount Remaining
- [ ] Confirm Sale flow: validation, stock decrement, DB write
- [ ] Receipt PDF generation + print
- [ ] Transaction number sequence

### Phase 2 — Discount + Customer
- [ ] `preset_discounts` table + settings management
- [ ] Cart discount % / Total Sale Price override / Preset dropdown
- [ ] Customer attachment mid-transaction (no cart clear)
- [ ] Email receipt flow: lookup → prompt if missing → save to profile

### Phase 3 — Sale Types
- [ ] Quote: save, recall, promote to Invoice
- [ ] Invoice: save with AR status, process payment against existing invoice
- [ ] Repair: `repair_jobs` table, Repair Details modal, status flow
- [ ] Deposit/Layby: `deposits` table, CSO link, layby completion recall
- [ ] Refund: load by transaction number/barcode, manual, stock restoration

### Phase 4 — Supporting Features
- [ ] Notes field: text editor, print toggle
- [ ] Park / Recall: cart snapshot, multiple parks, named parks
- [ ] Sales History panel: date filter, re-print, email, start refund
- [ ] Online payment type: amount entry + allocation release prompt
- [ ] Barcode scanning on receipt (Code 128 decode → load transaction)

### Phase 5 — Polish
- [ ] Keyboard navigation (Tab between fields, Return = new line, F-key shortcuts for payment methods)
- [ ] Offline resilience: queue failed Supabase writes to `data/pending_pos.json`
- [ ] Admin override for price edits below cost (optional; links to Plan 10 admin override)
- [ ] Margin warning: configurable threshold (e.g. flash red if line margin < 10%)
