# Module: Customer Special Orders (CSO)

> **Detail plan**: [docs/plans/06_customer_special_orders.md](../plans/06_customer_special_orders.md)
> **Build phase**: 4 — Operations
> **Tables owned**: `customer_allocations`, `sms_log`

---

## Overview

A Customer Special Order (CSO) is raised when a specific item is being sourced for a named
customer. The item goes onto the supplier's Open PO and `qty_allocated_customer` is incremented.
When the invoice is received and stock arrives, the customer is automatically notified via
TextMagic SMS. When collected, the allocation is cleared.

---

## Status Lifecycle

```
on_order → in_stock → collected
             ↑
    (set by invoice receive hook, Plan 04)
```

| Status | `qty_allocated_customer` | `qty_on_order` |
|--------|--------------------------|----------------|
| `on_order` | +qty (set on creation) | +qty (via PO line) |
| `waiting_order` | +qty | unchanged (no PO yet) |
| `in_stock` | unchanged | −qty (handled by invoice receive) |
| `collected` | −qty | unchanged |
| `cancelled` | −qty | −qty (PO line qty decremented) |

---

## How a CSO is Raised

1. Inventory right-click → **"Add to Customer Order"**
2. Customer Selection dialog (same search-as-you-type as Customer Management)
3. Confirm Order dialog: item (read-only), qty, notes, optional deposit checkbox
4. On confirm:
   - Item added to supplier's Open PO (auto-creating if needed — same logic as Plan 04)
   - `customer_allocations` record created (`status = 'on_order'`)
   - `items.qty_on_order +qty`, `items.qty_allocated_customer +qty`

---

## SMS Notification (TextMagic)

Triggered when Plan 04 invoice receive hook processes a received SKU that has a waiting CSO:

```
Hi [FirstName], your order for [Description] has arrived at Scarlett Music.
Give us a call on [StorePhone] or pop in to collect. Thanks!
```

Config: `textmagic_username`, `textmagic_api_key`, `sms_sender_name`, `store_phone` (in Settings → SMS).
If unconfigured: notification step silently skipped; staff shown manual-contact warning in invoice receipt screen.
All outbound SMS logged to `sms_log`.

**`src/sms_client.py`**:
```
SmsClient
  ├── send_arrival_notification(customer, item_description) → bool
  ├── send_custom_message(mobile, text) → bool
  └── _is_configured() → bool
```

---

## Collection Flow

**Via POS** (preferred): Load customer at till → add the CSO item → POS detects active `in_stock`
allocation and shows a banner: *"This item has a customer allocation — fulfilling special order."*
On confirm-sale: `status → 'collected'`, `qty_allocated_customer −qty`, `qty_on_hand −qty`.

**Manual**: Customer PO tab → right-click `in_stock` row → "Mark as Collected".
Prompts whether sale was through the till; if not, manually decrements stock.

---

## Cross-Module Connections

| Connection | Direction |
|-----------|----------|
| `purchase_order_lines.id` | CSO creation links to PO line (`po_line_id`) |
| `items.qty_allocated_customer` | +qty on creation; −qty on collection/cancellation |
| `items.qty_on_order` | +qty via PO line add; −qty on cancellation |
| Plan 04 invoice receive | Checks `customer_allocations` per received SKU; updates to `in_stock`; fires SMS |
| `deposits` | Optional deposit linked via `deposits.allocation_id` |
| Customer PO tab (Plan 05) | Displays CSOs for this customer |
| Inventory Customers tab (Plan 01) | Displays CSOs for this item |
| Items on Hold report (Plan 08) | Pulls from `customer_allocations` |
