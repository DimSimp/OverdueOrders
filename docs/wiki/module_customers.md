# Module: Customer Management

> **Detail plan**: [docs/plans/05_customer_management.md](../plans/05_customer_management.md)
> **Build phase**: 3 — POS Core
> **Tables owned**: `customers`, `discounts`
> **Tables displayed (not owned)**: `transactions` (filtered by customer), `repairs`, `deposits`, `customer_allocations`

---

## Overview

Manages all customer records and their associated history. The customer profile is the hub for
quotes, invoices, repairs, deposits, and special orders — all surfaced as tabs in a detail panel.
Search-first UI, consistent with the inventory pattern.

> **Schema note**: There are no separate `quotes` or `customer_invoices` tables. The Customer
> module's Quotes tab queries `transactions WHERE sale_type = 'quote' AND customer_id = X`.
> The Invoices tab queries `transactions WHERE sale_type = 'invoice' AND customer_id = X`.

---

## `discounts` table

> **Conflict resolved**: Merged from Plan 02 `preset_discounts` + Plan 05 `discounts`. One table,
> used in two contexts:
> 1. Attached to a customer profile (`customers.discount_id`) → auto-applied at POS checkout
> 2. Manual POS preset dropdown → staff picks a discount from the list

System presets (10%, 20%, 30%, 40%, 50%) are seeded with `is_system = true` and cannot be deleted.
Custom discounts are added in Settings → Discounts.

---

## Customer Detail Panel — 7 Tabs

| Tab | Data source |
|-----|-------------|
| Customer Info | `customers` record |
| Quotes | `transactions WHERE sale_type = 'quote'` |
| Invoices | `transactions WHERE sale_type = 'invoice'` |
| Repair | `repairs WHERE customer_id = X` |
| PO | `customer_allocations WHERE customer_id = X` |
| Audit | All activity across all tables for this customer |
| Deposit | `deposits WHERE customer_id = X` |

---

## Overdue Invoice Flagging

Same daily startup check as supplier invoices (Plan 04). Customer invoices where
`sale_status = 'pending_payment'` and `due_date < today` are flagged. Startup popup includes:
*"X customer invoice(s) overdue — $Y outstanding."*

---

## Cross-Module Connections

| Connection | Direction |
|-----------|----------|
| POS checkout | Loads customer; applies `discount_id`; attaches customer to transaction |
| POS receipt email | Updates `customers.email` if missing and staff provide one |
| Repair intake | Creates `repairs` record linked to this customer |
| Deposit taken | Creates `deposits` record linked to this customer |
| CSO raised | Creates `customer_allocations` record linked to this customer |
| SMS notification | `customers.mobile` used by `SmsClient` on stock arrival |
| Plan 09 import | Musipos customers bulk-imported; `musipos_account_code` + `musipos_barcode_ref` preserved |

---

## Required Fields

Minimum for new customer: `first_name` + `mobile`.

`customer_id` (sequential integer) and `customer_barcode` (Code 128 from `customer_id`) are
auto-generated on save. These are display/scan references — database FK is always `customers.id` (UUID).

---

## Role Permissions

| Action | `user` | `admin` |
|--------|--------|---------|
| View / create / edit customers | ✓ | ✓ |
| View Audit tab | ✗ | ✓ |
| Customer Data Export (report) | ✗ | ✓ |
| Manage discounts (Settings) | ✗ | ✓ |
