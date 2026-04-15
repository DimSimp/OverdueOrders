# Module: Supplier Management

> **Detail plan**: [docs/plans/03_supplier_management.md](../plans/03_supplier_management.md)
> **Build phase**: 1 — Foundation (schema needed before inventory)
> **Tables owned**: `suppliers`, `supplier_contacts`

---

## Overview

Manages all supplier records. The `suppliers` table is a dependency of `items` (FK), so it must
be populated before the inventory import runs. SKU suffix/prefix rules (currently in `config.json`)
are migrated here and become the source of truth for `web_sku` derivation.

---

## SKU Rules Migration

Currently `config.json` stores per-supplier suffix/prefix rules. Once `suppliers` is live:
- `suppliers.sku_suffix`, `sku_prefix`, `character_substitutions` become authoritative
- `src/config.py` reads these from Supabase instead of the local file
- `config.json` entries kept as read-only fallback during transition

---

## Cross-Module Connections

| Connection | Direction |
|-----------|----------|
| `items.supplier_id` | Every item FK → `suppliers.id` |
| `purchase_orders.supplier_id` | Every PO FK → `suppliers.id` |
| `invoices.supplier_id` | Every supplier invoice FK → `suppliers.id` |
| SKU suffix rules | Used by sync script and web_sku derivation in `items` |
| AI invoice parsing | Supplier card has "Import Invoice" button → triggers Plan 03 invoice import flow |
| `config.json` migration | After go-live, `src/config.py` reads suffix rules from `suppliers` table |

---

## UI: Supplier Card Tabs

1. **Details** — id, name, ABN, account number, terms, address, SKU rules, notes
2. **Contacts** — named contacts per supplier (sales rep, support, accounts)
3. **Purchase Orders** — list of POs (links to Plan 04 detail)
4. **Invoices** — supplier invoices + "Import Invoice" button + "Receive Without PO"

---

## Role Permissions

| Action | `user` | `admin` |
|--------|--------|---------|
| View suppliers | ✓ | ✓ |
| Edit supplier details | ✗ | ✓ |
| Create / delete suppliers | ✗ | ✓ |
| View POs and invoices | ✓ | ✓ |
| Import invoice | ✓ | ✓ |

---

## AI Invoice Import Integration

The invoice import pipeline is ported from the Web Portal (`C:\VB\Web Portal`) into
`src/invoice/`. Modules to port:

| Web Portal module | Destination |
|-------------------|------------|
| `pdf_parser.py` | `src/invoice/pdf_parser.py` |
| `models.py` | `src/invoice/models.py` |
| `validators.py` | `src/invoice/validators.py` |
| `sku_mapping.py` | `src/invoice/sku_mapping.py` |
| `supplier_import.py` | `src/invoice/importer.py` |

DB lookups updated from Musipos SQL → Supabase `items` table. Existing `src/pdf_parser.py` is
replaced by the ported version.
