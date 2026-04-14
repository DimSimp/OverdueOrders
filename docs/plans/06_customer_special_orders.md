# Plan 06 — Customer Special Orders

> **Part of**: [Master Plan](00_overview.md)
> **Status**: 🔲 Not started
> **Phase**: 4 — Operations

---

## Overview

A customer special order (CSO) is raised when a specific item is being sourced for a named customer — whether or not a deposit has been paid. The item is added to the supplier's Open PO and `qty_on_order` increments just as it would for a normal stock order. The customer record reflects the outstanding order under their **PO tab**. When the invoice is received and stock arrives, the system automatically sends an SMS to the customer via TextMagic notifying them their item is ready to collect. Once collected and sold, the allocation is cleared.

---

## Key Concepts

| Concept | Description |
|---------|-------------|
| **Customer Special Order (CSO)** | An item being ordered specifically for a named customer |
| **`qty_allocated_customer`** | Units on `items` reserved for named customers; part of the total Allocated figure |
| **`customer_allocations`** | Join table linking a customer to a specific item/PO line with status tracking |
| **TextMagic** | Third-party SMS gateway used to notify customers when their item arrives |
| **Collection** | When the customer comes in and the item is sold/handed over — converts the allocation to a completed sale |

---

## Database Schema

### `customer_allocations` — Tracks each CSO

| Column | Type | Notes |
|--------|------|-------|
| `id` | uuid PK | |
| `customer_id` | integer FK → `customers.customer_id` | |
| `item_id` | uuid FK → `items.id` | |
| `sku` | text NOT NULL | Denormalised for display |
| `description` | text | Item title at time of order |
| `qty` | integer NOT NULL | Qty ordered for this customer |
| `po_line_id` | uuid FK → `purchase_order_lines.id` | Linked PO line (set on creation if auto-added to PO) |
| `deposit_id` | uuid FK → `deposits.id` | Optional — if a deposit was paid |
| `status` | text NOT NULL DEFAULT `'on_order'` | See status lifecycle below |
| `notes` | text | Staff notes |
| `created_at` | timestamptz DEFAULT now() | |
| `created_by` | text | Staff user who raised the order |
| `notified_at` | timestamptz | When SMS notification was sent (null = not yet sent) |
| `collected_at` | timestamptz | When customer collected (null = outstanding) |
| `collected_by` | text | Staff user who marked collection |

**Status lifecycle:**

| Status | Meaning |
|--------|---------|
| `on_order` | Linked to a PO line; `qty_on_order` has been incremented |
| `waiting_order` | Not yet on a PO (e.g. no supplier chosen, or staff chose to defer PO) |
| `in_stock` | Invoice received — stock has arrived; SMS sent |
| `collected` | Customer collected; sale completed; allocation cleared |
| `cancelled` | Cancelled; deposit refund may be required |

In normal flow, status goes `on_order` → `in_stock` → `collected`. The `waiting_order` state is only used if the CSO is raised before a PO is created (edge case).

---

## Workflow

### 1. Raising a Customer Special Order

1. Staff right-clicks an item in the **Inventory module** → context menu → **"Add to Customer Order"**
2. The **Customer Selection** dialog opens — same search-as-you-type interface as the Customer Management module
   - Columns: Customer ID, Name, Business, Mobile
   - Double-click a row, or right-click → **"Select for Order"**, to choose the customer
3. **Confirm Order** dialog appears:
   - Item (SKU + description) — read-only
   - Customer name — read-only
   - **Qty** — editable integer field (default: 1)
   - **Notes** — optional free-text
   - **Add deposit now?** — checkbox; if checked, opens the deposit prompt inline (see plan 05 Deposit flow)
   - **[Confirm]** / **[Cancel]**
4. On confirm:
   - The item is added to the supplier's **Open PO** (auto-creating one if none exists, exactly as per plan 04)
   - A `customer_allocations` record is created with `status = 'on_order'`, linked to the new PO line
   - `items.qty_on_order` is incremented by qty
   - `items.qty_allocated_customer` is incremented by qty
   - The item's detail panel (Customer Allocations tab) refreshes to show the new entry
   - The customer's **PO tab** is updated (visible next time the customer record is opened)

> If the item already has an open CSO for this customer, the system warns staff before creating a duplicate.

---

### 2. Customer PO Tab

Accessible in the customer detail panel (see plan 05). Shows all active and historical special orders for this customer.

```
┌──────────────────┬──────────┬──────────────┬────────────┬────────────┬──────────────┐
│   SKU / Title    │   Qty    │   PO #       │  Status    │  Ordered   │  Notes       │
├──────────────────┼──────────┼──────────────┼────────────┼────────────┼──────────────┤
│  GTRSTR01        │    2     │  4068        │  On Order  │ 2026-04-10 │              │
│  Elixir Light    │          │              │            │            │              │
├──────────────────┼──────────┼──────────────┼────────────┼────────────┼──────────────┤
│  VIOLINBOW       │    1     │  4055        │  In Stock  │ 2026-03-20 │  Called x2   │
├──────────────────┼──────────┼──────────────┼────────────┼────────────┼──────────────┤
│  DRUMKIT-JR      │    1     │  4041        │  Collected │ 2026-02-15 │              │
└──────────────────┴──────────┴──────────────┴────────────┴────────────┴──────────────┘
```

- **Status colour coding**: On Order (white), In Stock (yellow/orange — needs attention), Collected (grey — historical)
- Right-click → **"Cancel Order"** (asks for confirmation; handles deposit refund prompt if deposit exists)
- Right-click → **"Mark as Collected"** (for manual marking if sale handled outside POS; see section 4)
- Click PO # → jumps to the PO in the Purchasing module
- Filter: All / Active (on_order + in_stock) / Collected / Cancelled

---

### 3. Inventory Item — Customer Allocations Tab

Already defined in plan 01 as Tab 2 of the item detail panel. Relevant columns:

```
┌──────────────────────────┬─────┬────────────────┬────────┬────────────┬──────────────┐
│   Customer               │ Qty │ Deposit Paid   │ Status │  Date      │  Notes       │
├──────────────────────────┼─────┼────────────────┼────────┼────────────┼──────────────┤
│  Jane Smith (#1042)      │  2  │  No            │On Order│ 2026-04-10 │              │
└──────────────────────────┴─────┴────────────────┴────────┴────────────┴──────────────┘
```

- Click customer name → opens the customer record
- "Add Special Order" button → triggers the same Confirm Order flow (customer selection first)

---

### 4. Stock Arrival — Notification Trigger

When a supplier invoice is received and committed (plan 04 invoice receiving flow), the importer checks for customer allocations against each received SKU:

1. For each committed `purchase_order_line`, query `customer_allocations` where `po_line_id` matches and `status = 'on_order'`
2. For each matching allocation:
   - Set `customer_allocations.status = 'in_stock'`
   - **Do not** decrement `qty_on_order` for this item — already handled by the normal invoice receive hook
   - `qty_allocated_customer` remains incremented (stock is still held for the customer)
3. Send TextMagic SMS to the customer's `mobile` number (see SMS section below)
4. Set `customer_allocations.notified_at = now()`
5. The customer's PO tab immediately shows **In Stock** (yellow) on next open
6. In the Purchasing / Invoice view, a notification badge or line annotation indicates: *"This item has a customer allocation — customer notified"*

> **If SMS fails**: Log the error, set `notified_at = null`, surface a visible warning in the invoice receipt confirmation screen: *"SMS to [Customer Name] failed — please contact them manually."* Staff can retry from the customer's PO tab (right-click → **"Resend Notification"**).

---

### 5. SMS Notification (TextMagic)

**Configuration** (stored in app settings / config):
- `textmagic_username` — TextMagic account username
- `textmagic_api_key` — TextMagic API key
- `sms_sender_name` — Display name shown on SMS (e.g. `"Scarlett Music"`, max 11 chars)
- `store_phone` — Included in the SMS body

**Default SMS template** (configurable in Settings):
```
Hi [FirstName], your order for [Description] has arrived at Scarlett Music.
Give us a call on [StorePhone] or pop in to collect. Thanks!
```

Variables:
| Variable | Source |
|----------|--------|
| `[FirstName]` | `customers.first_name` |
| `[Description]` | `customer_allocations.description` (item title at order time) |
| `[StorePhone]` | `config.store_phone` |

**API**: TextMagic REST API v2 — `POST https://rest.textmagic.com/api/v2/messages`

```json
{
  "text": "Hi Jane, your order for...",
  "phones": "0412345678"
}
```

Headers: `X-TM-Username`, `X-TM-Key`

The SMS module lives at `src/sms_client.py`. If TextMagic credentials are not configured, the notification step is skipped silently (no error — staff are still shown the "customer has an allocation" notice in the invoice receipt screen so they can call manually).

---

### 6. Collecting the Order

When the customer comes in to collect their item, the transaction is completed through the **POS/Till** (plan 02) or directly from the customer record:

**Via POS**:
- Load the customer at the till
- Add the item to the cart — the till detects an active `in_stock` CSO for this item and this customer
- A banner appears: *"This item has a customer allocation for this customer — fulfilling special order"*
- On sale completion:
  - `customer_allocations.status` → `collected`, `collected_at` and `collected_by` set
  - `qty_allocated_customer` decremented by qty
  - `qty_on_hand` decremented by qty (via normal sale flow)
  - `stock_movements` record of type `sale_instore` written

**Via customer record (manual)**:
- Open customer → PO tab → right-click `in_stock` row → **"Mark as Collected"**
- Prompts: *"Was the sale processed through the till? If not, stock will be adjusted manually."*
  - Yes → no stock change (assumes till already handled it)
  - No → decrements `qty_on_hand` and writes a `stock_movements` record of type `sale_instore` with note `"Manual collection via customer record"`
- Sets `status = 'collected'`

---

### 7. Cancellation

Right-click any active allocation (status `on_order` or `in_stock`) → **"Cancel Order"**:

1. Confirm dialog: *"Cancel special order for [Customer] — [Item]? This cannot be undone."*
2. If a deposit exists → additional prompt: *"A deposit of $[Amount] was paid. Refund required?"*
   - Yes → marks `deposits.status = 'refund_pending'`; refund is handled by staff manually (cash/EFTPOS)
   - No → deposit remains on record (e.g. customer forfeited it)
3. On confirm:
   - `customer_allocations.status` → `cancelled`
   - `qty_allocated_customer` decremented by qty
   - `qty_on_order` decremented by qty (item removed from or reduced on the PO line)
   - If the PO line qty drops to 0, the line is removed from the PO

---

## TextMagic Module (`src/sms_client.py`)

```
SmsClient
  ├── send_arrival_notification(customer, item_description) → bool
  ├── send_custom_message(mobile, text) → bool
  └── _is_configured() → bool  (checks for credentials in config)
```

- `send_arrival_notification()` builds the message from the template, substitutes variables, calls the API
- Returns `True` on success, `False` on failure (logged)
- All outgoing SMS are logged to a `sms_log` table in Supabase for reference (optional but recommended for disputes)

---

## Implementation Checklist

### Database
- [ ] Create `customer_allocations` table in Supabase
- [ ] Partial unique index or constraint: warn on duplicate active CSO for same customer + item
- [ ] Optional `sms_log` table: `id`, `customer_id`, `mobile`, `message`, `sent_at`, `status`, `error`

### Inventory Context Menu
- [ ] Add "Add to Customer Order" to inventory item right-click context menu
- [ ] Customer Selection dialog (search-as-you-type, double-click to select)
- [ ] Confirm Order dialog (qty, notes, optional deposit checkbox)
- [ ] On confirm: auto-add to supplier's Open PO (reuse plan 04 PO logic), create `customer_allocations` record
- [ ] Duplicate CSO warning if active allocation already exists for same customer + item

### Inventory Updates
- [ ] `qty_on_order` increment on CSO creation (via PO line add)
- [ ] `qty_allocated_customer` increment on CSO creation
- [ ] `qty_allocated_customer` decrement on collection or cancellation

### Customer PO Tab
- [ ] Display `customer_allocations` rows for this customer
- [ ] Status colour coding (On Order / In Stock / Collected / Cancelled)
- [ ] Right-click: Cancel Order, Mark as Collected, Resend Notification
- [ ] Click PO# → navigate to PO in Purchasing module

### Stock Arrival Notification
- [ ] Invoice receipt hook (plan 04) calls `check_customer_allocations(po_line_id)` after each line committed
- [ ] Set allocation status to `in_stock`, set `notified_at`
- [ ] Send TextMagic SMS via `SmsClient.send_arrival_notification()`
- [ ] On SMS failure: log error, surface warning in invoice receipt confirmation, allow retry
- [ ] "Customer allocation" annotation on invoice receipt confirmation screen

### SMS Client (`src/sms_client.py`)
- [ ] `SmsClient` class with `send_arrival_notification()` and `send_custom_message()`
- [ ] Read TextMagic credentials from config (username, api_key, sender_name, store_phone)
- [ ] Graceful no-op if credentials not configured
- [ ] SMS template configurable in app Settings
- [ ] Optional `sms_log` Supabase writes

### Collection Flow
- [ ] POS till detects active `in_stock` CSO when customer + item loaded
- [ ] CSO fulfilment banner in till cart
- [ ] On POS sale complete: set `collected`, decrement `qty_allocated_customer`
- [ ] Manual "Mark as Collected" from customer PO tab (with optional manual stock decrement)

### Cancellation Flow
- [ ] Cancel dialog with deposit refund prompt if deposit exists
- [ ] `deposits.status = 'refund_pending'` on Yes
- [ ] Decrement `qty_allocated_customer` and `qty_on_order`
- [ ] Remove PO line if qty drops to 0

---

## Open Questions / Future Considerations

- **Multiple arrivals on one PO line**: If 2 units were ordered for a customer on one PO line but only 1 arrives (partial delivery), the current model sets the whole allocation to `in_stock`. A future enhancement could support partial notification (e.g. "1 of 2 units has arrived"). Deferred for now — partial deliveries are uncommon for CSOs.
- **Customer preferred contact method**: Some customers may prefer a phone call to an SMS. A `contact_preference` field on the customer record could gate whether TextMagic fires. Deferred.
- **Multi-item CSOs**: Currently each allocation is per-item. A customer ordering 3 different items creates 3 separate CSO records. A future "grouped CSO" concept (like a mini customer PO) could bundle them under one reference. Deferred.

---

*Last updated: 2026-04-13*
