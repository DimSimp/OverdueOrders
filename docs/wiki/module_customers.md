# Module: Customer Management

> **Detail plan**: [docs/plans/05_customer_management.md](../plans/05_customer_management.md)
> **Build phase**: 3 — POS Core
> **Tables owned**: `customers`
> **Tables displayed (not owned)**: `transactions` (filtered by customer); placeholder tabs exist for `repairs` and `deposits`

---

## Overview

Manages customer records, contact details, addresses, profile discounts, and linked sale history.
The customer profile is intended to become the hub for quotes, invoices, repairs, deposits, and
special orders; the current detail panel already has that layout in place, with the customer info
and sale history pieces implemented first. Search-first UI, consistent with the inventory pattern.

> **Schema note**: There are no separate `quotes` or `customer_invoices` tables. The Customer
> module's Quotes tab queries `transactions WHERE sale_type = 'quote' AND customer_id = X`.
> The Invoices tab queries `transactions WHERE sale_type = 'invoice' AND customer_id = X`.

---

## Current UI State

The Customers tab is already usable inside the POS window:
- Search-first list with pagination and active/all/inactive filtering
- Detail panel with working `Customer Info` and `Sale History` tabs
- `Sale History` defaults to newest-first completed transactions, paginated 100 at a time
- `Sale History` has `From` / `To` date range filters with popup calendar pickers
- Placeholder tabs for `Quotes`, `Invoices`, `Repairs`, `Deposits`, and `Audit`
- CSV import button
- Create/edit customer modal
- Right-click `Load in Till` on a customer row
- Sale History can start a refund in the Till; the Till preloads the original transaction number,
  cart lines, customer, notes, and original payment method split

The unfinished parts are the deeper workflow tabs and any linked repair/deposit/quote editing from
inside the Customers module.

---

## Profile Discounts

Customer discounts are currently implemented as a simple hardcoded profile field stored on the
customer record as `customers.discount_profile`.

Available options:
- `5%`
- `10%`
- `15%`
- `Teacher`
- `Staff`

Current behaviour:
- When the customer is attached to a Till transaction, their `discount_profile` is auto-applied.
- The Till also has its own manual discount selector with the same options.
- If staff choose a Till-side discount manually, it overrides the customer profile discount for that transaction until cleared.
- `Teacher` is currently a placeholder 15% discount.
- `Staff` uses cost-based pricing that targets a 10% margin and rounds up when needed.

The older `discounts` table and `discount_id` fields still exist in the schema, but they are not
the active customer/POS discount mechanism today.

---

## Customer Detail Panel

| Tab | Data source |
|-----|-------------|
| Customer Info | `customers` record |
| Sale History | Completed transactions linked to the customer; newest-first, 100 per page, optional date range |
| Quotes | Placeholder tab |
| Invoices | Placeholder tab |
| Repairs | Placeholder tab |
| Deposits | Placeholder tab |
| Audit | Placeholder tab |

Only `Customer Info` and `Sale History` are currently implemented. The other tabs are visible so
the final layout is in place, but they still show coming-soon placeholders. The `Audit` tab is
present as a reminder for the future all-actions customer audit trail; it does not query audit data yet.

---

## Overdue Invoice Flagging

Same daily startup check as supplier invoices (Plan 04). Customer invoices where
`sale_status = 'pending_payment'` and `due_date < today` are flagged. Startup popup includes:
*"X customer invoice(s) overdue — $Y outstanding."*

---

## Cross-Module Connections

| Connection | Direction |
|-----------|----------|
| POS checkout | Loads customer; applies `discount_profile` unless Till manual discount is selected; attaches customer to transaction |
| Customers tab → Till | Right-click `Load in Till` attaches the selected customer to the active Till transaction |
| Sale History to Refund in Till | Loads the selected original sale into the Till as a linked refund, preserving transaction number and original payment split |
| POS receipt email | Updates `customers.email` if missing and staff provide one |
| Repair intake | Creates `repairs` record linked to this customer |
| Deposit taken | Creates `deposits` record linked to this customer |
| CSO raised | Creates `customer_allocations` record linked to this customer |
| SMS notification | `customers.mobile` used by `SmsClient` on stock arrival |
| Plan 09 import | Musipos customers bulk-imported; `musipos_account_code` + `musipos_barcode_ref` preserved |

---

## Required Fields

Minimum for new customer: `first_name` plus at least one of `mobile` or `phone_1`.

`customer_id` (sequential integer) and `customer_barcode` (Code 128 from `customer_id`) are
auto-generated on save. These are display/scan references — database FK is always `customers.id` (UUID).

---

## Customer Form

The create/edit customer modal currently uses a two-column layout:
- Left column: `General Details` and `Account Details`
- Right column: `Invoice Address` and `Shipping Address`

Implemented fields and behaviour:
- General details include first name, surname, business/school, mobile, phone, and email
- Shipping address has a `Copy Invoice -> Shipping` button
- Profile Discount dropdown is part of `Account Details`
- Both invoice and shipping addresses are saved to Supabase

---

## Role Permissions

| Action | `user` | `admin` |
|--------|--------|---------|
| View / create / edit customers | ✓ | ✓ |
| View Audit tab | ✗ | ✓ |
| Customer Data Export (report) | ✗ | ✓ |
| Manage discounts (Settings) | ✗ | ✓ |

---

*Last updated: 2026-04-27 — Sale History date range/calendar filtering and 100-row paging documented; Audit placeholder tab documented; refund handoff to Till now preserves original transaction/payment details.*
