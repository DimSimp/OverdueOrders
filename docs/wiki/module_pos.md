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

## Right Panel Layout

The right panel (≈30% width) displays all cart summary and payment UI, bottom-anchored so content
sits just above the Confirm Sale button. The top of this panel now hosts the active customer card
and transaction-note controls used by the receipt workflow:

```
[linked customer details + transaction notes]
══════════════════════════════════  ← section divider
Cart discount %: [___] [Apply]
Subtotal          $X.XX
Discount             —
Cart Margin          —
──────────────────────────────────
TOTAL             $X.XX
══════════════════════════════════  ← section divider
Payment Method
EFT      |  Cash      |  Online
[  ]+     |  [      ]  |  [      ]
          |  Change $X |
──────────────────────────────────
Remaining $X.XX  /  Paid in full ✓
[        Confirm Sale        ]
```

---

## Payment Methods

Payment is entered via an **always-visible 3-column inline panel** (EFT | Cash | Online) — no modals.

| Method | Behaviour |
|--------|-----------|
| Cash | Inline entry field; live "Change $X.XX" hint shown below as soon as cash exceeds amount due |
| EFT | Inline entry; `[+]` button adds rows below the last entry; `[−]` removes extras (min 1, max 3 rows); stored as `payment_eft` JSONB array |
| Online | Inline entry; records pre-paid amount; future: prompts to release `qty_allocated_online` |

Any combination of methods can be used simultaneously (split payment). A live status label below the
panel shows **Remaining $X.XX** (red), **Change due $X.XX** (green), or **Paid in full ✓** (green)
as amounts are typed. Confirm Sale is blocked until the full amount is covered.

---

## Discounts

The Till currently supports three discount mechanisms:
1. Manual cart discount % → reduces the whole total proportionally
2. Total Sale Price override → replaces the running total
3. Discount profile selector beside Sale Type → applies a predefined line-level pricing rule

The active discount profile options are hardcoded for now:
- `5%`
- `10%`
- `15%`
- `Teacher`
- `Staff`

Current behaviour:
- The same profile list is used by customer records (`customers.discount_profile`) and by the Till-side manual selector.
- The Till selector is shown for every sale type except `Refund`, where the refund lookup UI uses that same space.
- Only one profile source is active at a time. If staff choose a Till discount, it overrides any linked customer profile discount for that transaction. Clearing the Till selector back to `-` hands control back to the linked customer profile.
- The selected profile is applied to current cart lines and to new items added afterward.
- `Teacher` is currently a placeholder 15% discount until online-price-aware pricing is implemented.
- `Staff` calculates a sell price from item cost so margin lands at or just above 10%, rounding up when an exact 10.00% result is not possible.

The legacy `discounts` / `discount_id` schema is still present in Supabase, but it is not the
active POS/customer discount path today.

---

## Customer Integration

Customer linking is now a working part of the Till flow:
- Staff can search and attach a customer from the Till itself.
- Staff can also right-click a customer in the Customers tab and choose `Load in Till`, which switches back to the Till and attaches that profile to the active transaction.
- When a customer is linked, the Till uses the same attach path regardless of where that customer came from, so the customer card, receipt output, and profile discount behaviour stay consistent.
- Historical recall/refund flows re-link customers without repricing old transactions.

---

## Parked Transactions

`sale_status = 'parked'` with `cart_snapshot JSONB` preserves full cart state. The trigger
auto-assigns a `transaction_number` (`T-YYYY-NNNN`) when a row is parked.

**Park flow**: Staff click **Park** in the top bar → confirm dialog → `park_transaction()` INSERTs
a `transactions` row with `sale_status='parked'`, `cart_snapshot={cart_items, cart_disc_pct,
customer_name, sale_type}` — no `transaction_lines` or stock movements. Cart is then cleared.

**Recall flow**: Staff click **Recall** → `RecallDialog` (700×460 modal) loads all parked rows via
`get_parked_transactions()`. The list shows date parked, transaction number, customer name, and
total. Double-click or "Select for POS" restores the cart from `cart_snapshot` and **immediately
deletes the parked DB row** (fire-and-forget background thread). The recalled sale then completes
as a new standard transaction via `confirm_standard_sale()`.

**Delete**: The recall modal also has a red **Delete** button to permanently remove a parked
transaction without recalling it.

**Files**: `src/pos/transaction_client.py` — `park_transaction`, `delete_parked_transaction`,
`get_parked_transactions`; `src/gui/pos/recall_dialog.py` — `RecallDialog`.

---

## Daily Sales

A non-modal **Daily Sales** window (`DailySalesDialog`, `src/gui/pos/daily_sales_dialog.py`)
is accessible via the **Sales** button in the Till top bar. It lists all completed transactions
for today (Melbourne time).

**Layout**:
- Each transaction appears as a **collapsible parent row** (TX #, date/time, customer, user,
  payment methods, total). Expanding it reveals per-line child rows (SKU, description, qty, RRP,
  disc $, line total, cost, margin $, margin %) and a green **TOTAL** footer row.
- A fixed **Day Summary** panel at the bottom aggregates: transactions count, items sold, total
  RRP, total discount, revenue, cost, margin $, margin %, and Cash/EFT/Online payment breakdown.

**Interactions**:
- Click the **#0** (arrow) column header → expand / collapse all transactions at once.
- Click the **Date & Time** column header → toggle sort order (▼ newest-first / ▲ oldest-first).
- Expanding a transaction inserts a subtle **divider row** (`#2a2a2a` stripe) after it via
  `<<TreeviewOpen>>` / `<<TreeviewClose>>` events; dividers are removed when collapsed.
- Click any row → enables **Reprint Receipt** button in the header.
- Right-click any row → context menu with **Reprint Receipt**.
- **Refresh** button re-fetches from Supabase without closing the window.

**Reprint**: reconstructs `cart_items` from stored `transaction_lines`, calls `generate_receipt()`
then `print_pdf()` on a background thread.

---

## Cart UX Details

- **SKU/Barcode entry**: "Add" button animates as a rotating spinner (`◐◓◑◒`) during lookup; disabled until result returns. On 1 match → adds to cart; on 0 or multiple → switches to Inventory tab with search pre-filled.
- **Right-click context menu** on any cart row: **Show in Inventory** (switches tab, searches by SKU, and auto-selects the item so its detail panel opens); **Remove from Cart**.
- **Inline cell editing**: single-click on Qty, Unit Price, Disc %, or Line Total opens an overlay entry; editing Line Total back-calculates Disc %.
- **Cart TOTAL override**: click the TOTAL figure to type a new total directly; Disc % is back-calculated.
- **Margin column**: per-line gross margin shown in green (>10%), orange (=10%), or red (<10%). Cart Margin summary shown in breakdown panel (staff-visible only; never printed on receipts).

---

## Cross-Module Connections

| Action | What POS does |
|--------|--------------|
| Load customer in Till | Queries `customers`; attaches the profile to the active transaction; auto-applies `discount_profile` unless a Till-side manual discount is currently selected |
| Load in Till from Customers tab | Uses the same customer-attach path as the Till search pane, then returns focus to the Till tab |
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

**Implemented** — `src/pos/receipt_generator.py`. Generated via reportlab as an 80mm thermal PDF.
Content trimmed to actual height using PyMuPDF (`_trim_to_content`) to avoid blank paper feed.

Contains:
- Store name, address, phone, ABN (from `config.shipping.sender`)
- Transaction number (e.g. `T-2026-0001`) + Code 128 barcode (for refund scanning)
- Date/time (Melbourne local), staff name
- Linked customer summary when present
- Line items: SKU, description (truncated to 18 chars), qty, unit price, line total
- Per-line discount sub-rows (italic, grey) if `disc_pct > 0`
- Subtotal, item discounts, cart discount, **TOTAL** (large bold)
- Payment breakdown: cash (+ tendered / change), EFT (multiple entries), online
- Optional transaction notes when print-notes is enabled
- GST included line

Printed via `src/printer_utils.print_pdf()` using the configured `receipt_printer` device.
Post-sale: `_ReceiptDialog` (CTkToplevel in `till_tab.py`) prompts staff to print.
Reprint: available from the Daily Sales dialog for any past transaction.

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
