# Plan 08 — Reporting

> **Part of**: [Master Plan](00_overview.md)
> **Status**: 🔲 Not started
> **Phase**: 5 — Reporting

---

## Overview

The reporting module provides a catalogue of configurable reports covering sales, inventory, customers, suppliers, and compliance. Most reports share a common pattern: select filters → preview results in-app → export to CSV and/or PDF.

The **Daily Sales Report** (end-of-day close) is the exception — it is a structured operational process with float management, a locked printed summary, and a daily record written to Supabase. It is treated separately from the other reports.

---

## Key Concepts

| Concept | Description |
|---------|-------------|
| **Z-Report** | End-of-day close process. Staff enter actual cash in drawer; the system calculates expected vs actual; a locked PDF summary is produced. One per day. |
| **Date Range** | All reports accept a `from` / `to` date range. Default is the current day or current month depending on the report type. |
| **Export** | All reports can export to CSV (for Excel/accounting) and PDF (for printing or filing). |
| **Float** | The cash kept in the till. Opening float is entered at start of day; closing count entered at end of day before Z-report is run. |
| **BAS** | Business Activity Statement — quarterly GST report lodged with the ATO. The GST/BAS report must match the figures in Xero/MYOB, so it must include in-store POS sales, online sales, and supplier purchases. |

---

## Database Schema

### `daily_summaries` — End-of-day close records

One row per day the Z-report is run.

| Column | Type | Notes |
|--------|------|-------|
| `id` | uuid PK | |
| `summary_date` | date UNIQUE | The business date this summary covers |
| `opening_float` | numeric(10,2) | Cash in drawer at start of day |
| `closing_float_expected` | numeric(10,2) | Calculated: opening_float + cash sales |
| `closing_float_actual` | numeric(10,2) | Staff-counted cash at close |
| `float_variance` | numeric(10,2) | `actual − expected` (negative = short) |
| `total_sales_instore` | numeric(10,2) | All POS sales for the day (inc. GST) |
| `total_sales_neto` | numeric(10,2) | Neto dispatches for the day |
| `total_sales_ebay` | numeric(10,2) | eBay dispatches for the day |
| `total_sales_all` | numeric(10,2) | Sum of all channels |
| `total_gst_collected` | numeric(10,2) | |
| `total_transactions` | integer | Number of POS transactions |
| `generated_by` | text | Staff user |
| `generated_at` | timestamptz | |
| `pdf_path` | text | Local path to the saved PDF (for reprinting) |
| `notes` | text | Optional staff notes (e.g. reason for variance) |

---

## The Daily Sales Report (Z-Report)

### Purpose

Performed at end of each trading day. Records total takings, reconciles cash, and produces a printed/PDF summary for filing. This is the source of truth for daily cash and should be run before closing.

### Process Flow

1. Staff opens the Daily Sales Report from the main menu
2. The report automatically loads all transactions for today:
   - POS in-store sales (from plan 02 `pos_transactions` table)
   - Online dispatches (from `online_sales` where `dispatched_at = today`)
3. **Float Entry** panel:
   - `Opening float` — pre-filled from yesterday's closing float (or manually entered if first run)
   - `Actual closing cash count` — staff enter denomination breakdown or total
4. System calculates:
   - Expected cash = opening float + today's cash sales
   - Variance = actual − expected
   - Variance displayed in green (within ±$1 tolerance) or red (outside tolerance)
5. Staff add optional note (explain variance if any)
6. **[Close Day]** button:
   - Writes `daily_summaries` record
   - Generates PDF summary (see layout below)
   - Saves PDF to `data/daily_summaries/{YYYY-MM-DD}.pdf`
   - Locks the record (cannot be re-run for the same date without manager override)

### PDF Summary Layout

```
═══════════════════════════════════════════════
       SCARLETT MUSIC — DAILY SUMMARY
              {Day, Date}
═══════════════════════════════════════════════
SALES SUMMARY
  In-Store Sales        $  1,240.00
  Online — Neto         $    380.00
  Online — eBay         $    215.00
  ─────────────────────────────────
  TOTAL SALES           $  1,835.00
  GST Included          $    166.82

TRANSACTIONS
  In-Store              12
  Online                 7

FLOAT RECONCILIATION
  Opening Float         $    200.00
  + Cash Sales          $    650.00
  Expected Close        $    850.00
  Actual Count          $    848.00
  Variance              −$    2.00  ⚠
  Notes: Customer paid $48 + $2 short

Generated: {timestamp}  Staff: {name}
═══════════════════════════════════════════════
```

> **Reprint**: Staff can reprint any past daily summary by date from the report catalogue. The original locked PDF is used.

---

## Report Catalogue

### Common Controls (all reports)

- **Date range**: From / To date pickers (most default to current month)
- **[Generate]**: Loads results into preview table
- **[Export CSV]**: Downloads a flat CSV for Excel
- **[Export PDF]**: Generates a formatted printable PDF
- **[Print]**: Sends to default printer

---

### 1. Paid Invoices

Shows all customer invoices with `status = 'complete'` in the date range.

**Filters**: Date range, Customer (search), Payment method

**Columns**: Invoice #, Date, Customer, Items (count), Total (inc. GST), Payment Method, Paid By

**Totals row**: Sum of all invoice totals for the period

**Use case**: Bookkeeping reconciliation, providing customer statements

---

### 2. Invoices Outstanding (Aged Debtors)

Shows all customer invoices with `status = 'open'` or `'sent'` — money owed to the store.

**Filters**: Customer (search), Overdue only (checkbox), Aging bracket

**Columns**: Invoice #, Invoice Date, Customer, Due Date, Days Overdue, Amount (inc. GST)

**Aging breakdown** (summary at top of report):

```
Current (not yet due)    $  2,400.00
1–30 days overdue        $    850.00
31–60 days overdue       $    200.00
61–90 days overdue       $      0.00
90+ days overdue         $    125.00
─────────────────────────────────────
TOTAL OUTSTANDING        $  3,575.00
```

**Overdue rows highlighted red** in the preview table

---

### 3. Inventory Report

A snapshot of current stock levels.

**Filters**: Supplier, Brand, Instrument category, Stock status (All / In Stock / Out of Stock / Low Stock), Include zero-qty items (checkbox)

**Columns**: SKU, Title, Brand, Supplier, On Hand, Allocated, Available, On Order, RRP, Cost

**Sort**: Any column; defaults to SKU ascending

**Notes row**: Items where Available < `min_order_level` are flagged with a yellow highlight

---

### 4. Reorder Report

Items that need restocking — where `qty_available < min_order_level`. Grouped by supplier for easy PO creation.

**Filters**: Supplier, Brand, Instrument category

**Columns**: SKU, Title, Brand, On Hand, Available, Min Level, Max Level, Suggested Order Qty (`max_order_level − qty_on_hand`), Supplier

**Grouped by supplier** with subtotal of suggested order quantities per supplier

**[Create PO]** button next to each supplier group — opens a draft PO in the Purchasing module pre-filled with the suggested lines

---

### 5. Stock Valuation

Total value of current inventory, useful for insurance and accounting.

**Filters**: Supplier, Brand, Instrument category, Valuation method (Cost / RRP)

**Columns**: SKU, Title, Brand, Supplier, On Hand, Unit Cost, Total Cost Value, Unit RRP, Total RRP Value

**Summary at top**:

```
Total SKUs:              1,842
Total Units:             4,103
Total Value at Cost:     $  248,500.00
Total Value at RRP:      $  497,200.00
Potential Margin:        $  248,700.00  (50.0%)
```

**Date-stamped** on export (not historical — reflects current stock levels at time of generation)

---

### 6. Stock Movement Audit

Every `stock_movements` record for a given SKU or period. Used to investigate inventory discrepancies.

**Filters**: SKU (required or date range required — must provide at least one), Date range, Movement type (All / receive / sale / dispatch / adjustment / return / stocktake / allocate)

**Columns**: Date/Time, SKU, Title, Movement Type, Qty Change, Running Total (On Hand), Reference (order ID, invoice #, etc.), Performed By, Notes

**Use case**: "Why does the system show 8 units but I can only find 6?" — trace every movement to find the discrepancy

---

### 7. Customer Data Export

Exports customer profile data — primarily for marketing lists.

**Filters**: Date created (range), State, Has email (checkbox), Has mobile (checkbox)

**Columns**: Customer ID, First Name, Surname, Business, Email, Mobile, Address, City, State, Postcode, Created Date

**Privacy note**: This report contains personal information. Access requires manager or admin role. Exports are logged in the audit trail.

**Use case**: Import into Mailchimp, generate mailing labels, marketing campaigns

---

### 8. Items on Hold

All active customer special orders — items that have been ordered or have arrived and are waiting for customer collection.

**Filters**: Status (All / On Order / In Stock — arrived, awaiting pickup), Supplier, Date range

**Columns**: Customer Name, Customer Mobile, SKU, Title, Qty, PO #, Status, Days Waiting, Deposit Paid ($), Notified (SMS sent date)

**"In Stock" rows highlighted yellow** (arrived but not yet collected — action needed)

**Sort**: Defaults to Status (In Stock first), then Days Waiting descending

---

### 9. Outstanding Repairs

All repairs that have not been collected.

**Filters**: Status (Ongoing / Complete — not yet collected / All active), Date range (date lodged)

**Columns**: Repair #, Customer Name, Customer Mobile, Description, Status, Date Lodged, Date Completed, Days in Store, Notes

**Rows colour-coded** matching repair status colours (plan 05): Ongoing = white, Complete = yellow

**Sort**: Defaults to Complete first (needs attention), then Days in Store descending

---

### 10. Supplier Report

Purchase history and outstanding balances for a supplier.

**Filters**: Supplier (required), Date range

**Tabs**:
- **Purchase History**: All received invoices in the period — Invoice #, Date, Lines (count), Total (exc. GST), Total (inc. GST), GST
- **Outstanding POs**: All active POs (Open / Sent) — PO #, Date, Lines, Total Value, Status
- **Accounts Payable**: Unpaid invoices — Invoice #, Due Date, Days Overdue, Amount — with aging breakdown matching Invoices Outstanding format

**Summary row**: Total spend for period, Total outstanding

---

### 11. Dispatch / Shipping Summary

All orders dispatched in a period, broken down by courier.

**Filters**: Date range, Platform (All / Neto / eBay / In-store), Courier

**Columns**: Date, Order ID, Platform, Customer Name, Courier, Tracking Number, Items (count), Weight, Freight Cost

**Courier subtotals**: Number of parcels and total freight spend per courier

**Use case**: Reconcile courier invoice against bookings; identify freight cost trends

---

### 12. Online Channel Performance

Side-by-side comparison of online sales channels vs in-store over a period.

**Filters**: Date range (defaults to current month), Granularity (Daily / Weekly / Monthly)

**Summary table**:

```
                  In-Store    Neto        eBay        TOTAL
Revenue           $12,400     $4,200      $1,800      $18,400
Units Sold            182         54          21          257
Avg Sale Price     $68.13     $77.78      $85.71       $71.59
Returns / Replacmt    $0       $120         $45         $165
Net Revenue       $12,400     $4,080      $1,755      $18,235
```

**Trend chart**: Optional bar chart output in PDF (daily/weekly revenue by channel)

---

### 13. GST / BAS Summary

All taxable transactions for a period — designed to match the figures needed for BAS lodgement.

**Filters**: Date range (defaults to current quarter), Include purchases (checkbox)

**Sales section**:
- Total sales (inc. GST)
- GST on sales (= total ÷ 11)
- GST-free sales (if any)

**Purchases section** (if checked):
- Total supplier invoices (inc. GST)
- GST on purchases (input tax credits)

**Net GST payable** = GST on sales − GST on purchases

> **Disclaimer note on report**: *"This report is a guide only. Confirm figures with your accountant before lodging your BAS."*

---

## UI Layout

The reporting module is a separate tab or section in the main app. Layout:

```
┌────────────────────────────────────────────────────────────────┐
│  REPORTS                                                       │
├─────────────────┬──────────────────────────────────────────────┤
│  Daily Close    │                                              │
│  ─────────────  │   [Selected report configuration panel]      │
│  Sales          │                                              │
│    Paid Invoices│   Date From: [__/__/____]  To: [__/__/____]  │
│    Outstanding  │   Filter:   [____________]  [✓] Overdue only │
│    Online Perf. │                                              │
│    GST / BAS    │   [Generate]  [Export CSV]  [Export PDF]     │
│  ─────────────  │                                              │
│  Inventory      │  ┌──────────────────────────────────────┐    │
│    Inventory    │  │ Invoice # │ Customer │ Due │ Amount  │    │
│    Reorder      │  ├──────────────────────────────────────┤    │
│    Valuation    │  │ INV-1042  │ J. Smith │ ... │ $240.00 │    │
│    Stock Audit  │  │ INV-1039  │ ABC Pty  │ ... │ $880.00 │    │
│  ─────────────  │  └──────────────────────────────────────┘    │
│  Customers      │                                              │
│    Data Export  │   Showing 24 results — Total: $3,575.00      │
│    Items on Hold│                                              │
│  ─────────────  │                                              │
│  Suppliers      │                                              │
│    Supplier Rpt │                                              │
│    Dispatch     │                                              │
│  ─────────────  │                                              │
│  Repairs        │                                              │
│    Outstanding  │                                              │
└─────────────────┴──────────────────────────────────────────────┘
```

---

## Implementation Checklist

### Database
- [ ] Create `daily_summaries` table in Supabase
- [ ] Ensure `stock_movements` has all required fields for audit report (type, qty_change, reference_id, performed_by, notes)
- [ ] Ensure `online_sales` has `sale_price`, `cost_at_dispatch`, `sale_type`, `is_dropship`, `returned`
- [ ] Ensure `customer_invoices` has `status`, `total`, `payment_method`, `paid_at`

### Reporting Module UI
- [ ] Scaffold reporting tab/section with left nav and right configuration + results panel
- [ ] Shared date range picker component
- [ ] Shared export helpers: CSV writer, PDF generator (using reportlab or similar)
- [ ] Shared print helper

### Daily Sales Report (Z-Report)
- [ ] Load today's POS and online transactions automatically
- [ ] Float entry panel (opening float pre-fill from yesterday's record)
- [ ] Expected vs actual variance calculation with colour coding
- [ ] [Close Day] writes `daily_summaries` record and locks it
- [ ] PDF generation with layout matching the summary format above
- [ ] Save PDF to `data/daily_summaries/{YYYY-MM-DD}.pdf`
- [ ] Reprint past summaries from the catalogue
- [ ] Manager override to re-run a locked day

### Individual Reports
- [ ] Paid Invoices
- [ ] Invoices Outstanding (with aging breakdown)
- [ ] Inventory Report (with low-stock highlighting)
- [ ] Reorder Report (with [Create PO] shortcut)
- [ ] Stock Valuation (with summary totals)
- [ ] Stock Movement Audit (SKU or date range required)
- [ ] Customer Data Export (access restricted to manager+; export logged)
- [ ] Items on Hold (In Stock rows highlighted yellow)
- [ ] Outstanding Repairs (colour coding matching repair status)
- [ ] Supplier Report (3-tab: Purchase History, Outstanding POs, Accounts Payable)
- [ ] Dispatch / Shipping Summary (courier subtotals)
- [ ] Online Channel Performance (summary table + optional chart)
- [ ] GST / BAS Summary (with disclaimer note)

### Shared Infrastructure
- [ ] PDF report template (header: store name, report title, date range, generated timestamp)
- [ ] CSV export function (utf-8-sig encoding for Excel on Windows)
- [ ] Role check on Customer Data Export (manager/admin only)
- [ ] Audit log entry when Customer Data Export is generated

---

## Open Questions / Future Considerations

- **Saved report configurations**: Allow staff to save a named filter preset (e.g. "Monthly Neto report") and re-run it in one click. Deferred — straightforward to add later.
- **Scheduled reports**: Auto-generate and email certain reports on a schedule (e.g. weekly stock valuation to management). Deferred — requires email config.
- **Xero/MYOB integration**: Direct export of the GST/BAS figures in a format that imports into the accounting package. Deferred — CSV export is sufficient for now.
- **Online Channel Performance chart**: A bar chart in the PDF adds value but requires a charting library (matplotlib or similar). The table format is sufficient for now; chart can be added later.
- **Repair profitability**: A report showing revenue per repair vs parts cost. Deferred until repair billing is fully implemented in plan 11.

---

*Last updated: 2026-04-14*
