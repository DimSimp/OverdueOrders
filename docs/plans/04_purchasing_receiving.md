# Plan 04 — Purchasing & Receiving

> **Part of**: [Master Plan](00_overview.md)
> **Status**: 🔲 Not started
> **Phase**: 4 — Operations (schema needed earlier to back PO references in inventory)

---

## Overview

Manages the full lifecycle of purchase orders — from quick item-level ordering via the inventory screen, through finalisation and sending to the supplier, to receiving stock against the matching invoice. Accounts payable (supplier bills and due dates) live here. No explicit "create PO" action exists; POs are auto-created the moment an item is added and none is currently open.

---

## Key Concepts

| Concept | Description |
|---------|-------------|
| **Open PO** | The active PO accumulating items. Only one can exist per supplier at a time. Auto-created on first item add. |
| **PO Number** | Sequential integer per supplier. Auto-assigned on PO creation (`last PO for this supplier + 1`). |
| **Finalise** | Staff action that locks the PO (no more items), generates the PDF, and marks it `Pending`. |
| **Send** | Emails the PO PDF to the supplier's sales rep. Marks it `Sent`. |
| **Receive** | Process of matching an incoming invoice to a PO — manual or AI-assisted. Marks PO `Complete`. |
| **Supplier Bill** | The financial record created when an invoice is received: total, due date, payment status. |
| **Running Countdown** | Live total shown while entering invoice lines: `Invoice Total − Σ(line totals) − freight − insurance`. Should reach $0 on a complete entry. |

---

## PO Status Lifecycle

```
[Item added, no Open PO] → Open → Pending → Sent → Complete
                           ↑         ↑         ↑        ↑
                       Auto-created  Finalise  Email   Receive
```

| Status | Description |
|--------|-------------|
| `Open` | Accepting items. Auto-created with next PO number when needed. |
| `Pending` | Locked — no more items. PDF generated. Awaiting email to supplier. |
| `Sent` | Emailed to supplier. Awaiting stock delivery and invoice. |
| `Complete` | Invoice received and processed. Stock updated. |

---

## Database Schema

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
| `pdf_path` | text | Path to generated PO PDF (local) |
| `notes` | text | Optional internal notes |

**Constraint**: Only one `open` PO per `supplier_id` at a time (enforced in app logic and ideally a partial unique index).

---

### `purchase_order_lines`

| Column | Type | Notes |
|--------|------|-------|
| `id` | uuid PK | |
| `po_id` | uuid FK → `purchase_orders.id` | |
| `item_id` | uuid FK → `items.id` | |
| `sku` | text | Snapshot of SKU at time of ordering |
| `title` | text | Snapshot of item title |
| `rrp` | numeric(10,2) | Snapshot of RRP at time of ordering |
| `min_sell` | numeric(10,2) | Snapshot of minimum sell price |
| `qty_ordered` | integer | Qty staff requested |
| `qty_received` | integer DEFAULT 0 | Filled in when invoice is received |
| `qty_backordered` | integer DEFAULT 0 | Supplier could not supply; on backorder |
| `unit_cost_inc_gst` | numeric(10,2) | Filled in on invoice receipt |
| `unit_cost_exc_gst` | numeric(10,2) | Calculated: `unit_cost_inc_gst / 1.1` |
| `line_total` | numeric(10,2) | `unit_cost_inc_gst × qty_received` |

---

### `invoices`

| Column | Type | Notes |
|--------|------|-------|
| `id` | uuid PK | |
| `invoice_number` | text NOT NULL | Supplier-generated; not known until invoice arrives |
| `supplier_id` | text FK → `suppliers.id` | |
| `po_id` | uuid FK → `purchase_orders.id` | Matched PO (nullable if no PO reference) |
| `invoice_date` | date | Written on the invoice |
| `due_date` | date | Default: `invoice_date + supplier.payment_terms_days`; editable |
| `subtotal_exc_gst` | numeric(10,2) | Product lines only |
| `gst` | numeric(10,2) | |
| `freight` | numeric(10,2) DEFAULT 0 | Freight charge from invoice |
| `insurance` | numeric(10,2) DEFAULT 0 | Insurance charge from invoice |
| `total_inc_gst` | numeric(10,2) | As entered by staff (with inc/exc GST toggle) |
| `status` | text DEFAULT `unpaid` | `unpaid`, `paid`, `overdue` |
| `received_at` | timestamptz DEFAULT now() | |
| `received_by` | text | Staff user |
| `entry_method` | text | `manual` or `ai` |

---

## UI Design

### PO List (in Supplier Card → Purchase Orders tab)

Covered in [03_supplier_management.md](03_supplier_management.md). The tab displays:

```
┌─────────┬────────────┬──────────┬─────────┬──────────┬──────────────┐
│  PO #   │  PO Date   │  Status  │ Ordered │ Received │ Back Ordered │
├─────────┼────────────┼──────────┼─────────┼──────────┼──────────────┤
│  4065   │ 2026-03-28 │ Complete │   45    │  45/45   │      0       │
│  4066   │ 2026-03-30 │ Complete │   40    │  32/40   │      8       │
│  4067   │ 2026-04-01 │ Sent     │    6    │    —     │      —       │
│  4068   │ —          │ Open     │    3    │    —     │      —       │
└─────────┴────────────┴──────────┴─────────┴──────────┴──────────────┘
```

- **PO Date** = `finalised_at` date (blank while Open)
- **Ordered** = total `qty_ordered` across all lines
- **Received** = `total_qty_received / total_qty_ordered` — shown as `x/y`; blank until invoice received
- **Back Ordered** = total `qty_backordered` across all lines; blank until invoice received
- A PO always moves to `Complete` when an invoice is received, regardless of whether it was fully fulfilled — the Received column shows the actual result (e.g. `32/40`)
- Right-click a row → context menu (see below)
- The Open PO row is always at the top and highlighted

**Context menu options by status:**

| Status | Options |
|--------|---------|
| Open | View Items, Finalise PO |
| Pending | View Items, View PDF, Send PO, Edit (re-open to Open) |
| Sent | View Items, View PDF, Receive (manual), Auto-Receive (AI) |
| Complete | View Items, View PDF, View Invoice |

---

### Adding Items to a PO (from Inventory Screen)

1. User selects an item in the inventory grid
2. **Right-click → "Add to PO"** or **press Return** on selected row
3. The first column (`QTY on Order`) becomes an editable text field
4. User types desired quantity and presses **Return** to confirm
5. System checks: does an `Open` PO exist for this item's supplier?
   - **Yes** → add line to existing Open PO
   - **No** → auto-create new PO (next sequential number for this supplier), then add line
6. `items.qty_on_order` incremented by entered qty
7. The `QTY on Order` column updates immediately in the grid

---

### Finalise PO (Open → Pending)

1. Right-click Open PO → "Finalise PO"
2. Confirmation dialog: "This will lock PO #4067. No more items can be added."
3. On confirm:
   - Status → `Pending`; `finalised_at` = now
   - PO PDF generated and saved locally (see PDF Format below)
4. PO row updates in the list

---

### Send PO (Pending → Sent)

1. Right-click Pending PO → "Send PO"
2. Email compose dialog:
   - **To**: pre-filled with supplier's primary/sales rep email
   - **Subject**: `Purchase Order #4067 — Scarlett Music`
   - **Body**: brief template with PO number and expected items count
   - **Attachment**: generated PO PDF
3. Staff can edit any field before sending
4. On send: status → `Sent`; `sent_at` = now

---

### Receive Invoice — Manual

Triggered by right-clicking a `Sent` PO → "Receive".

```
┌────────────────────────────────────────────────────────────┐
│  Receive Invoice — Pro Music PO #4066                      │
├────────────────────────────────────────────────────────────┤
│  Invoice No: [__________]  Invoice Date: [__/__/____]      │
│  Due Date:   [__/__/____]  (auto: invoice date + 30 days)  │
│  Invoice Total: [$_______]  [✓ Includes GST]               │
├──────┬────────────────────────┬────────┬────────┬──────────┤
│ SKU  │ Title                  │ Ord'd  │ Rec'd  │Unit Cost │
├──────┼────────────────────────┼────────┼────────┼──────────┤
│ auto │ auto-filled            │ auto   │ [  ]   │ [      ] │
│ ...  │ ...                    │ ...    │ [  ]   │ [      ] │
├──────┴────────────────────────┴────────┴────────┴──────────┤
│                              Freight: [$_____]             │
│                            Insurance: [$_____]             │
│                            ─────────────────────────       │
│                            Remaining: $12.50  ← live       │
│                                                            │
│                      [Cancel]   [Confirm Receipt]          │
└────────────────────────────────────────────────────────────┘
```

**Behaviour:**
- SKU, Title, and Qty Ordered are read-only (from PO lines)
- Tab key moves focus through Qty Received → Unit Cost → next row
- Unit Cost field label shows "(inc GST)" or "(exc GST)" matching the toggle
- Line Total = `unit_cost_inc_gst × qty_received` (calculated, not shown in grid but used for countdown)
- **Remaining** = `Invoice Total − Σ(line totals) − freight − insurance` — updates live
- "Confirm Receipt" is enabled when Remaining is between -$0.10 and $0.10 (small tolerance for rounding)
- If Remaining > $0 after all lines are entered, staff can adjust freight/insurance to account for it

**On confirm:**
1. Write `invoices` record to Supabase
2. Update each `purchase_order_lines` row: `qty_received`, `qty_backordered`, `unit_cost_inc_gst`, `unit_cost_exc_gst`, `line_total`
3. Update `purchase_orders.status` → `Complete`
4. Trigger inventory receive hook (plan 01): update `qty_on_hand`, `last_purchase_cost`, average costs
5. Write `stock_movements` records of type `receive`
6. Append received lines to `daily_reports/received_YYYY-MM-DD.csv`
7. Check customer allocations → prompt for any SKUs with waiting special orders (see plan 06)

---

### Receive Invoice — AI Assisted

Triggered by right-clicking a `Sent` PO → "Auto-Receive".

1. File picker opens — staff selects scanned invoice PDF
2. AI parsing runs (ported from Web Portal `pdf_parser.py`)
3. Parsed result populates the same Receive window pre-filled:
   - Invoice number, invoice date, total — from parsed header
   - Qty received, qty backordered, unit cost — from parsed line items
   - Freight/insurance — from parsed non-product lines
4. Staff reviews: any unresolved SKUs flagged in orange with correction prompt (ported SKU mapping UX)
5. If PO number parsed from invoice doesn't match the current PO → warning shown
6. Staff confirms → same commit flow as manual receipt

---

### PO PDF Format

Generated when a PO is finalised. Stored locally as `data/pos/{supplier_id}_{po_number}.pdf`.

```
┌─────────────────────────────────────────────────────────┐
│  SCARLETT MUSIC                                         │
│  [Address]  |  [Phone]  |  [Email]  |  ABN: [ABN]      │
│                                        PO #: 4067       │
│                                        Date: 2026-04-01 │
├─────────────────────────────────────────────────────────┤
│  To: [Supplier Name]                                    │
│  Account No: [Our account number with this supplier]    │
├──────────────┬───────┬──────────────────────┬──────────┤
│     SKU      │  QTY  │     Description      │  Brand   │
├──────────────┼───────┼──────────────────────┼──────────┤
│  ...         │  ...  │  ...                 │  ...     │
├──────────────┴───────┴──────────────────────┴──────────┤
│                                  Total QTY: 12          │
└─────────────────────────────────────────────────────────┘
```

**Fields included**: Store name, address, phone, email, ABN; supplier name; our account number with this supplier; PO number; date; line items (SKU, QTY, Description, Brand); total QTY at footer.

**Prices are deliberately excluded** — the supplier sets their own prices and will include them on the invoice they send back. Prices at the time of ordering are not relevant to the PO.

---

### Accounts Payable View

Accessible from the main app (not per-supplier — shows all outstanding bills across all suppliers). Displayed as a simple table:

```
┌────────────────┬──────────────┬────────────┬──────────────┬─────────┐
│ Supplier       │ Invoice No.  │ Inv. Date  │ Due Date     │ Total   │
├────────────────┼──────────────┼────────────┼──────────────┼─────────┤
│ Pro Music      │ PM-29041     │ 2026-03-15 │ 2026-04-14   │ $842.10 │
│ D'Addario      │ DA-10827     │ 2026-03-20 │ 2026-04-19   │ $315.40 │
│ Dunlop         │ DL-88321     │ 2026-03-28 │ 2026-04-27   │ $127.60 │ ← overdue shown in red
└────────────────┴──────────────┴────────────┴──────────────┴─────────┘
Total outstanding: $1,285.10
```

- Overdue invoices (due date < today) highlighted in red
- "Mark Paid" button per row (admin/manager only) → sets `invoices.status = paid`
- Filter: All / Unpaid / Overdue / Paid
- Can be exported to CSV for accounting purposes

**Overdue auto-flag**: A daily background task runs on app startup and checks all `unpaid` invoices. Any where `due_date < today` are automatically set to `overdue`. On app startup, if any overdue invoices exist, a toggleable popup notifies the user: *"You have X overdue invoice(s) totalling $Y. View now?"*. The popup can be disabled per-user in settings.

---

### Receiving Without a Matching PO

Occasionally stock arrives without a corresponding PO — a backorder from a long-completed PO, an unsolicited sample, or a standing order placed outside the app. In this case:

1. Staff navigates to the supplier in the Suppliers module
2. Right-click in the PO list → **"Receive Without PO"**
3. A new PO is automatically created with the next sequential number
4. The Receive Invoice window opens against this new PO (no pre-filled line items — staff enters everything manually or uses AI)
5. On confirm, the PO status is immediately set to `Complete` (it skips `Open`, `Pending`, `Sent` entirely)
6. The auto-generated PO is marked with `entry_method = 'ad_hoc'` on the invoice record for reporting purposes

> **Also noted in**: [03_supplier_management.md](03_supplier_management.md) — the "Receive Without PO" option appears in the supplier card PO tab context menu.

---

### Credit Notes

Suppliers occasionally issue credit notes — for returned goods, pricing errors, or short deliveries billed incorrectly. Credits must be tracked to offset outstanding bills.

**`credit_notes` table:**

| Column | Type | Notes |
|--------|------|-------|
| `id` | uuid PK | |
| `credit_note_number` | text NOT NULL | Supplier-issued reference |
| `supplier_id` | text FK → `suppliers.id` | |
| `invoice_id` | uuid FK → `invoices.id` | Original invoice being credited (nullable) |
| `credit_date` | date | |
| `amount_inc_gst` | numeric(10,2) | Credit value |
| `reason` | text | e.g. `Return`, `Pricing Error`, `Short Delivery` |
| `status` | text DEFAULT `outstanding` | `outstanding`, `applied` |
| `notes` | text | |
| `received_at` | timestamptz DEFAULT now() | |
| `received_by` | text | |

**UI**: Credit notes appear in a sub-tab within the supplier Invoices tab. They are also surfaced in the Accounts Payable view as negative amounts, reducing the total outstanding for that supplier.

**Stock reversal**: If a credit relates to returned goods, a corresponding `stock_movements` record of type `return` is written and `qty_on_hand` is decremented. If it is a pricing correction only, no stock movement is made.

---

## Implementation Checklist

### Infrastructure
- [ ] Create `purchase_orders` table in Supabase
- [ ] Create `purchase_order_lines` table in Supabase
- [ ] Create `invoices` table in Supabase
- [ ] Partial unique index: enforce single `open` PO per `supplier_id`
- [ ] Add `last_po_number` column to `suppliers` table (or derive from MAX query)

### Data Layer (`src/purchasing/`)
- [ ] `po_client.py` — CRUD for POs and lines
  - [ ] `get_open_po(supplier_id)` → Open PO or None
  - [ ] `create_po(supplier_id)` → auto-assigns next PO number
  - [ ] `add_item_to_po(supplier_id, item_id, qty)` → creates PO if needed
  - [ ] `remove_item_from_po(po_line_id)`
  - [ ] `update_po_line_qty(po_line_id, qty)`
  - [ ] `finalise_po(po_id)` → status → pending, generate PDF
  - [ ] `send_po(po_id, email_details)` → send email, status → sent
  - [ ] `get_po_lines(po_id)`
  - [ ] `get_pos_for_supplier(supplier_id)`
- [ ] `invoice_client.py` — invoice and accounts payable operations
  - [ ] `receive_invoice_manual(po_id, invoice_data, lines)` → full commit
  - [ ] `receive_invoice_ai(po_id, pdf_path)` → parse + return pre-fill data
  - [ ] `get_unpaid_invoices()` → accounts payable list
  - [ ] `mark_invoice_paid(invoice_id)`
  - [ ] `get_invoice(invoice_number)`

### PDF Generation (`src/purchasing/po_pdf.py`)
- [ ] Generate PO PDF from PO + lines data
- [ ] Save to `data/pos/{supplier_id}_{po_number}.pdf`
- [ ] Open/view PDF (system default viewer)

### Email (`src/purchasing/po_email.py`)
- [ ] Compose email with PO PDF attachment
- [ ] Pre-fill To (supplier sales rep), Subject, Body template
- [ ] Send via SMTP (reuse existing email config)
- [ ] Editable before sending

### GUI (`src/gui/purchasing/`)
- [ ] `receive_invoice_window.py` — shared window for manual + AI receipt
  - [ ] Header fields: Invoice No., Invoice Date, Due Date, Invoice Total + GST toggle
  - [ ] Line items grid: read-only SKU/Title/Qty Ordered + editable Qty Received/Unit Cost
  - [ ] Freight + Insurance fields
  - [ ] Live remaining countdown
  - [ ] Confirm button (enabled near $0)
  - [ ] Unknown SKU correction prompt (AI path)
- [ ] `po_send_dialog.py` — email compose + send
- [ ] `accounts_payable_view.py` — outstanding bills table with Mark Paid + export
- [ ] Inventory grid changes (plan 01 integration):
  - [ ] Right-click context menu item: "Add to PO"
  - [ ] Return key shortcut on selected row
  - [ ] Inline qty entry field in QTY on Order column
  - [ ] Visual feedback on successful add (brief highlight or toast)

### Inventory Integration
- [ ] `add_item_to_po()` increments `items.qty_on_order`
- [ ] Invoice receipt decrements `items.qty_on_order`, increments `items.qty_on_hand`
- [ ] PO line records linked from item's Order/Receiving History tab (plan 01 detail panel)

### Daily CSV Output
- [ ] Append received product lines to `daily_reports/received_YYYY-MM-DD.csv` on invoice commit
- [ ] Format: `supplier_id, supplier_invoice_no, po_no, sku, qty_received, unit_cost, description`

### Overdue Invoice Flagging
- [ ] Daily background task on app startup: set `invoices.status = overdue` for unpaid invoices past due date
- [ ] Startup popup: count + total of overdue invoices with "View Now" button
- [ ] User setting to toggle the startup popup on/off

### Receiving Without a PO
- [ ] "Receive Without PO" option in supplier card PO tab context menu
- [ ] Auto-create PO with next sequential number, `entry_method = 'ad_hoc'`
- [ ] Open Receive Invoice window with no pre-filled lines
- [ ] On confirm: PO immediately set to `Complete`

### Credit Notes
- [ ] Create `credit_notes` table in Supabase
- [ ] `create_credit_note(supplier_id, data)` in `invoice_client.py`
- [ ] Credit notes sub-tab in supplier Invoices tab
- [ ] Surface credits as negative amounts in Accounts Payable view
- [ ] Stock reversal path: if return-related, write `stock_movements` record of type `return` and decrement `qty_on_hand`

---

## Open Questions / Future Considerations

- **Scarlett Music store details for PO PDF**: Phone number, email address, and full address needed to populate the PO header. These should be stored in app settings (config) rather than hardcoded.
- **PDF library choice**: ReportLab is the most capable option for Python PDF generation; WeasyPrint (HTML → PDF) may be easier to template. Decision deferred until build phase.
- **Partial backorder follow-up**: If a PO completes with `32/40` (8 backordered), there is currently no automated mechanism to expect and receive those 8 remaining units later. A future "backorder tracking" feature could create a new PO line entry for backordered items automatically. Deferred.

---

*Last updated: 2026-04-02 — resolved partial delivery, PO PDF fields, overdue flagging, ad-hoc receive, credit notes*
