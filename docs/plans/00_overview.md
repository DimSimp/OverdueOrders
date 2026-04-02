# Scarlett AIO — POS & Inventory System: Master Plan

> **Purpose**: This document is the top-level reference for the full POS and inventory management system being built into Scarlett AIO. Each major feature links to its own detailed plan file. Use this document to track phase-level progress and understand how systems relate to each other.

---

## Guiding Principles

- **Musipos is not replaced outright** — it remains running for its APIC import pipeline and legacy data. The new system imports from it and runs alongside it.
- **Single source of truth for stock** — all channels (in-store POS, Neto, eBay) read from and write to the same inventory database.
- **Cloud-first inventory** — inventory lives in Supabase (PostgreSQL) so it is accessible by the desktop app, the background sync script, and any future integrations.
- **Record keeping is paramount** — every stock movement (sale, receive, allocate, dispatch, return) must be traceable.
- **Search-first UI** — inventory screens do not bulk-load data; everything is paginated and driven by search/filter.

---

## System Architecture Overview

```
┌─────────────────────────────────────────────────────────┐
│                  Scarlett AIO (Desktop App)              │
│   Inventory UI │ POS / Till │ Purchasing │ Reporting     │
└────────────────────────┬────────────────────────────────┘
                         │ REST (Supabase client)
                         ▼
              ┌──────────────────────┐
              │   Supabase (Cloud)   │
              │   PostgreSQL DB      │
              │   + Auth + Storage   │
              └──────┬───────────────┘
                     │
        ┌────────────┴────────────┐
        ▼                         ▼
 GitHub Actions             Existing App
 Sync Script                Dispatch Hook
 (every 5 min)              (marks dispatched
 Neto + eBay orders →       → adjusts stock)
 allocate stock
```

---

## Feature Modules

| # | Module | Plan File | Status |
|---|--------|-----------|--------|
| 1 | **Inventory System** | [01_inventory_system.md](01_inventory_system.md) | 🔲 Not started |
| 2 | **Point of Sale (Till)** | [02_pos_till.md](02_pos_till.md) | 🔲 Not started |
| 3 | **Supplier Management** | [03_supplier_management.md](03_supplier_management.md) | 🔲 Not started |
| 4 | **Purchase Orders & Receiving** | [04_purchasing_receiving.md](04_purchasing_receiving.md) | 🔲 Not started |
| 5 | **Customer Management** | [05_customer_management.md](05_customer_management.md) | 🔲 Not started |
| 6 | **Customer Special Orders** | [06_customer_special_orders.md](06_customer_special_orders.md) | 🔲 Not started |
| 7 | **Online Order Integration** | [07_online_integration.md](07_online_integration.md) | 🔲 Not started |
| 8 | **Reporting & Daily Sales Log** | [08_reporting.md](08_reporting.md) | 🔲 Not started |
| 9 | **APIC / Musipos Import** | [09_apic_import.md](09_apic_import.md) | 🔲 Not started |
| 10 | **Staff & User Management** | [10_staff_users.md](10_staff_users.md) | 🔲 Not started |

---

## Module Summaries

### 1. Inventory System
The foundation of everything else. Manages all product records, stock levels (On Hand / Allocated / Available), serial numbers, kit/bundle definitions, and item metadata. Cloud-hosted in Supabase with a background sync script keeping online allocations current. See [01_inventory_system.md](01_inventory_system.md).

**Depends on**: Supabase setup, Supplier Management (for supplier_id references)
**Required by**: POS, Purchasing, Online Integration, Reporting

---

### 2. Point of Sale (Till)
Full in-store till replacement. Barcode scan or SKU search → cart → payment (cash/EFTPOS) → receipt → stock decrement. In-store sales appear on the same daily sales log as online orders. See [02_pos_till.md](02_pos_till.md).

**Depends on**: Inventory System, Customer Management, Staff & User Management
**Required by**: Reporting, Customer Special Orders

---

### 3. Supplier Management
Manage supplier records (codes, names, contact details, lead times, suffix/prefix rules). Supplier codes map directly from Musipos for compatibility. See [03_supplier_management.md](03_supplier_management.md).

**Depends on**: Nothing
**Required by**: Inventory System, Purchasing, APIC Import

---

### 4. Purchase Orders & Receiving
Create purchase orders for suppliers, track expected delivery, and receive stock into inventory. Receiving stock increments On Hand and triggers customer special order notifications if applicable. See [04_purchasing_receiving.md](04_purchasing_receiving.md).

**Depends on**: Inventory System, Supplier Management, Customer Special Orders
**Required by**: Reporting

---

### 5. Customer Management
Customer database: name, contact details, purchase history, special order history. Foundation for the POS customer lookup and special order notifications. See [05_customer_management.md](05_customer_management.md).

**Depends on**: Nothing
**Required by**: POS, Customer Special Orders, Reporting

---

### 6. Customer Special Orders
Manage orders where stock is being sourced specifically for a named customer (paid or unpaid deposit). When stock is received against a special order, the customer is automatically notified via TextMagic SMS. See [06_customer_special_orders.md](06_customer_special_orders.md).

**Depends on**: Customer Management, Inventory System, Purchasing
**Required by**: POS (to flag allocated stock at till)

---

### 7. Online Order Integration
Background GitHub Actions sync script that fetches Neto and eBay orders every 5 minutes and allocates stock in Supabase. When an order is marked dispatched in the existing app, the hook decrements On Hand and clears the allocation. See [07_online_integration.md](07_online_integration.md).

**Depends on**: Inventory System
**Required by**: Reporting, Daily Sales Log

---

### 8. Reporting & Daily Sales Log
Unified daily sales view combining in-store POS transactions and online orders. End-of-day summary, stock movement log, revenue by channel, and staff performance. See [08_reporting.md](08_reporting.md).

**Depends on**: POS, Online Integration, Inventory System
**Required by**: Nothing (terminal module)

---

### 9. APIC / Musipos Import
Tools for importing product data from Musipos CSV exports and from the APIC shared supplier/retailer database. Includes a field mapping layer (e.g. `Supplier_Item_ID` → `sku`, `Publisher_Brand` → `brand`). See [09_apic_import.md](09_apic_import.md).

**Depends on**: Inventory System, Supplier Management
**Required by**: Nothing (utility module)

---

### 10. Staff & User Management
Role-based access: administrator, manager, staff. Controls who can override minimum sell price, access cost prices, issue refunds, etc. See [10_staff_users.md](10_staff_users.md).

**Depends on**: Nothing
**Required by**: POS, Reporting

---

## Build Phases

### Phase 1 — Foundation (Current)
- [ ] Supabase project setup + schema design
- [ ] Inventory system (full — see plan 01)
- [ ] APIC/Musipos import (to populate initial data — see plan 09)

### Phase 2 — Online Bridge
- [ ] GitHub Actions sync script (online allocation)
- [ ] Dispatch hook integration into existing app
- [ ] Online Integration module (see plan 07)

### Phase 3 — POS Core
- [ ] Supplier Management (see plan 03)
- [ ] Customer Management (see plan 05)
- [ ] Point of Sale / Till (see plan 02)

### Phase 4 — Operations
- [ ] Purchase Orders & Receiving (see plan 04)
- [ ] Customer Special Orders + TextMagic SMS (see plan 06)
- [ ] Staff & User Management (see plan 10)

### Phase 5 — Reporting
- [ ] Unified Reporting & Daily Sales Log (see plan 08)

---

*Last updated: 2026-04-01*
