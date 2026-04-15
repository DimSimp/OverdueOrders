# Module: POS / Till

> **Detail plan**: [docs/plans/02_pos_till.md](../plans/02_pos_till.md)
> **Build phase**: 3 — POS Core
> **Tables owned**: `transactions`, `transaction_lines`
> **Tables written to**: `repairs` (Repair type), `deposits` (Deposit type), `items` (stock), `stock_movements`, `customers` (email update), `daily_summaries` (via reporting)

---

## Overview

The primary day-to-day transaction screen. Handles all in-store sales, quotes, invoices, repairs,
deposits/laybys, and refunds. Every completed transaction writes to `transactions` and
`transaction_lines`. Stock movements are written for applicable sale types at confirm-sale time.

> **Schema note**: `transactions` is the canonical record for ALL sale types including Quote and
> Invoice. Plan 05's separate `quotes`/`customer_invoices` tables are not built — customer module
> views query `transactions` filtered by `sale_type`.

---

## UI Structure

The POS system uses a **tab-based layout**. Top-level navigation tabs (Till, Inventory, Customers,
Purchasing, Reporting) are always visible within the POS window. Switching tabs preserves state —
an in-progress cart is not cleared when navigating to Inventory to look up a product and then
returning to the Till.

---

## Development Access

During development the POS button on the home screen is **hidden for all users except one
designated admin**. The authorised username is stored in `config.json` under the key `pos_dev_user`.
Any logged-in user whose username does not match this value will not see the POS button at all.

When the POS system is ready for general release, set `pos_dev_user` to `null` (or remove the
key) — the button becomes visible to all logged-in staff.

```python
# home_window.py — POS button visibility
show_pos = (
    self._current_user is not None and
    (pos_dev_user is None or self._current_user["username"] == pos_dev_user)
)
```

---

## Sale Types

| Type | Customer required? | Stock movement | Notes |
|------|-------------------|----------------|-------|
| Standard Sale | No | `qty_on_hand −qty` on confirm | Default |
| Quote | Yes (before confirm) | None | `sale_status = 'draft'`; promotable to Invoice |
| Invoice | Yes | `qty_on_hand −qty` on confirm | `sale_status = 'pending_payment'`; collects payment later |
| Repair | Yes | None at creation; charged on collection | Also writes to `repairs` table |
| Deposit | Yes | `qty_allocated_customer +qty` on confirm | Also writes to `deposits` table |
| Refund | No | `qty_on_hand +qty` on confirm | Negative totals; loads from prior transaction |

---

## Transaction Lifecycle

```
Standard Sale:   [cart built] → [payment entered] → [confirm] → stock moves, receipt
Quote:           [cart built] → [confirm draft] → [recall → promote to Invoice]
Invoice:         [cart built] → [confirm, stock moves] → [recall → mark paid]
Deposit:         [cart built] → [deposit amount entered] → [confirm] → allocation created
Repair:          [repair details entered] → [confirm intake] → [collect via POS later]
Refund:          [load original transaction OR enter manually] → [confirm] → stock restored
```

---

## Payment Methods

| Method | Behaviour |
|--------|-----------|
| Cash | Modal for cash tendered → calculates Change Given |
| EFT | Multiple entries tracked individually in `payment_eft` JSONB array |
| Online | Records pre-paid online order pickup; prompts to release `qty_allocated_online` |

Split payments: cash + EFT simultaneously. "Amount Remaining" display updates live.

---

## Discounts

Three controls, only one applied at a time:
1. Manual cart discount % → reduces total proportionally
2. Total Sale Price override → replaces running total
3. Preset dropdown → selects from `discounts` table (merged from Plan 02 `preset_discounts` + Plan 05 `discounts`)

Customer's `discount_id` auto-applied when customer is loaded onto the transaction.

---

## Parked Transactions

`sale_status = 'parked'` with `cart_snapshot JSONB` preserves full cart state. Multiple parks
allowed. Recalled via dropdown in top bar.

---

## Cross-Module Connections

| Action | What POS does |
|--------|--------------|
| Load customer | Queries `customers`; auto-applies their `discount_id` |
| Scan barcode / enter SKU | Queries `items` for exact match; populates row |
| Confirm Standard/Invoice sale | Decrements `qty_on_hand`, writes `sale_instore` movement |
| Confirm Deposit | Increments `qty_allocated_customer`, creates `deposits` record |
| Confirm Repair | Creates `repairs` record; charges deposit if collected |
| Confirm Refund | Increments `qty_on_hand`, writes `return` movement |
| CSO detection | On cart load with customer: checks `customer_allocations` for `in_stock` items |
| Online payment method | Optionally releases `qty_allocated_online` (user prompted) |
| Email receipt | Updates `customers.email` if blank and staff enter one |
| SKU no match | Shows error; fuzzy-searches inventory module with unmatched term |

---

## Receipt

Printed via reportlab. Contains:
- Store name, address, phone
- Transaction number (e.g. `T-2026-0001`) + Code 128 barcode (for refund scanning)
- Customer details (if attached)
- Line items: SKU, description, qty, unit price, discount, total — **no margin column**
- Payment method breakdown + change given
- Notes (if "Print on receipt" toggled)
- GST included line

---

## Role Permissions

| Action | `user` | `admin` |
|--------|--------|---------|
| Process sales | ✓ | ✓ |
| Sell below `minimum_sell` | Override | ✓ |
| Apply discount > threshold | Override | ✓ |
| Issue refund | ✗ | ✓ |
| View cost / margin columns | ✗ | ✓ |
| Void a completed transaction | ✗ | ✓ |
