# Module: Purchasing & Receiving

> **Detail plan**: [docs/plans/04_purchasing_receiving.md](../plans/04_purchasing_receiving.md)
> **Build phase**: 4 — Operations
> **Tables owned**: `purchase_orders`, `purchase_order_lines`, `invoices`, `credit_notes`

---

## Overview

Manages the full purchase order lifecycle from creation to stock receipt. POs are auto-created
when the first item is added (no manual "create PO" step). Receiving an invoice updates inventory
costs, increments `qty_on_hand`, triggers customer allocation notifications, and creates a supplier
bill for accounts payable.

---

## PO Status Lifecycle

```
[item added, no open PO] → Open → Pending → Sent → Complete
                           ↑         ↑         ↑        ↑
                        Auto-created Finalise  Email   Receive
```

Only one `open` PO per supplier at a time (partial unique index enforced).

---

## Adding Items to a PO

From the **Inventory screen**: right-click item → "Add to PO", or press Return on selected row.
System finds or creates an Open PO for that item's supplier, increments `items.qty_on_order`.

From the **CSO flow** (Plan 06): item is added to the Open PO automatically when a Customer
Special Order is raised.

---

## Invoice Receiving — what happens on confirm

1. Write `invoices` record
2. Update `purchase_order_lines`: `qty_received`, `qty_backordered`, `unit_cost_inc/exc_gst`, `line_total`
3. Set `purchase_orders.status → 'complete'`
4. **Inventory receive hook** (Plan 01): increment `qty_on_hand`, update `last_purchase_cost`, recalculate average costs
5. Write `stock_movements` record of type `receive`
6. Decrement `items.qty_on_order`
7. **Append to** `daily_reports/received_YYYY-MM-DD.csv` (preserves existing dispatch comparison)
8. **Check customer allocations** (Plan 06): for each received SKU with a waiting CSO → set status to `in_stock`, send TextMagic SMS

---

## Cross-Module Connections

| Connection | Direction |
|-----------|----------|
| `items.qty_on_order` | +qty on PO line add; −qty on invoice receive |
| `items.qty_on_hand` | +qty on invoice receive (via inventory hook) |
| `stock_movements` | `receive` record written on invoice commit |
| `customer_allocations` | Checked per received SKU; SMS triggered if waiting |
| `daily_reports/received_YYYY-MM-DD.csv` | Appended on every invoice receive |
| Inventory screen | "Add to PO" right-click triggers PO line creation |
| Supplier card tabs 3 & 4 | Displays POs and invoices for that supplier |
| Accounts Payable view | Unpaid `invoices` shown here; overdue auto-flagged |

---

## Accounts Payable

Separate view (not per-supplier) showing all `invoices WHERE status != 'paid'`. Overdue invoices
(due date past, status = `unpaid`) auto-flagged on app startup. Startup popup shows count + total.

"Mark Paid" button (admin/manager) sets `invoices.status = 'paid'`.

Credit notes appear as negative amounts in accounts payable, reducing total outstanding.

---

## Role Permissions

| Action | `user` | `admin` |
|--------|--------|---------|
| View / create POs | ✓ | ✓ |
| Receive invoice (manual or AI) | ✓ | ✓ |
| Delete / void a PO | ✗ | ✓ |
| Mark invoice paid | ✓ | ✓ |
| Create / edit credit notes | ✓ | ✓ |
