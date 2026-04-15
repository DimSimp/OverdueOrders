# Module: Reporting & Daily Close

> **Detail plan**: [docs/plans/08_reporting.md](../plans/08_reporting.md)
> **Build phase**: 5 — Reporting
> **Tables owned**: `daily_summaries`
> **Data sources**: `transactions`, `online_sales`, `items`, `customers`, `invoices`, `repairs`, `customer_allocations`, `stock_movements`

---

## Overview

Two components:
1. **Z-Report / Daily Close** — structured end-of-day process with float reconciliation and a
   locked PDF. One per day. Source of truth for daily cash.
2. **Report Catalogue** — 13 configurable reports covering sales, inventory, customers,
   suppliers, and compliance. Each follows the pattern: set filters → preview → export CSV/PDF.

> **Table name note**: Plan 08 referenced `pos_transactions` — the correct table name is
> `transactions` (Plan 02). All queries against that plan's `pos_transactions` reference use
> `transactions`.

---

## Z-Report (Daily Close) Process

1. Staff opens Daily Sales Report
2. Today's transactions (in-store + online dispatches) load automatically
3. Staff enter actual closing cash count
4. System calculates expected vs actual variance (green within ±$1, red outside)
5. Staff add optional variance note
6. **[Close Day]** writes `daily_summaries`, generates locked PDF, saves to `data/daily_summaries/{YYYY-MM-DD}.pdf`
7. Record locked — requires manager override to re-run for same date

Next day's opening float pre-filled from yesterday's `closing_float_actual`.

---

## Report Catalogue

| # | Report | Key data source | Admin only? |
|---|--------|----------------|-------------|
| 1 | Paid Invoices | `transactions WHERE sale_type IN ('standard','invoice') AND sale_status = 'completed'` | No |
| 2 | Invoices Outstanding (aged) | `transactions WHERE sale_type = 'invoice' AND sale_status = 'pending_payment'` | No |
| 3 | Inventory Report | `items` | No (hides cost for `user` role) |
| 4 | Reorder Report | `items WHERE qty_available < min_order_level` | No |
| 5 | Stock Valuation | `items` with cost/RRP calc | Yes |
| 6 | Stock Movement Audit | `stock_movements` | Yes |
| 7 | Customer Data Export | `customers` | Yes (logged) |
| 8 | Items on Hold | `customer_allocations WHERE status IN ('on_order','in_stock')` | No |
| 9 | Outstanding Repairs | `repairs WHERE status NOT IN ('collected','cancelled')` | No |
| 10 | Supplier Report | `invoices`, `purchase_orders`, supplier AP | Yes |
| 11 | Dispatch / Shipping Summary | `online_sales` + booking data | No |
| 12 | Online Channel Performance | `transactions` + `online_sales` | Yes |
| 13 | GST / BAS Summary | `transactions` + `invoices` | Yes |

---

## Role Permissions

| Feature | `user` | `admin` |
|---------|--------|---------|
| Daily close (Z-report) | ✓ | ✓ |
| Items on Hold, Outstanding Repairs | ✓ | ✓ |
| Inventory Report (no costs) | ✓ | ✓ |
| Reorder Report | ✓ | ✓ |
| All financial reports (Invoices, Valuation, GST, etc.) | ✗ | ✓ |
| Customer Data Export | ✗ | ✓ (logged) |
| Stock Movement Audit | ✗ | ✓ |
| Re-run locked daily close | ✗ | Override |

---

## Reorder Report → Create PO

The Reorder Report includes a **[Create PO]** button next to each supplier group. Clicking it
opens a draft PO in the Purchasing module pre-filled with the suggested reorder lines
(`max_order_level − qty_on_hand` per item).

---

## Output Formats

- **CSV**: `utf-8-sig` encoding for Excel compatibility on Windows (consistent with existing `src/exporter.py`)
- **PDF**: reportlab; standard header (store name, report title, date range, timestamp)
- **Print**: send to default system printer
