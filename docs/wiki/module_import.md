# Module: Musipos / APIC Import

> **Detail plan**: [docs/plans/09_apic_import.md](../plans/09_apic_import.md)
> **Build phase**: 1 — Foundation (runs after initial schema creation, before go-live)
> **Location**: Settings → Data Import (admin-only, one-time wizard)

---

## Overview

A one-time migration wizard that populates Supabase from three Musipos data sources. Import order
is fixed: Suppliers → Inventory → Customers (inventory needs supplier UUIDs as FKs).

| Source | File | Rows (approx) |
|--------|------|--------------|
| Inventory | `musipos_inventory.CSV` | ~8,400 active |
| Suppliers | `Suppliers.PDF` | ~18 relevant of 44 total |
| Customers | `Customer_List.CSV` | ~9,200 of 12,800 after filtering |

---

## Key Import Decisions

| Decision | Details |
|----------|---------|
| `qty_on_hand` | **Ignored** — all items imported at 0. Full stocktake required before go-live. |
| `Minimum_Sell = 0` | Stored as `null` (no minimum set) |
| Inactive items (`Active = N`) | Skipped by default; optional checkbox to include as inactive |
| Supplier PDF page split | Two-page PDF aligned by row index; fuzzy match on IDs (strip trailing 'S', then Levenshtein ≤ 2) |
| Customer garbage records | Filtered by regex: `^[-*\s#!@$%^&()\[\]{}]{3,}$` on col 2 (name field) |
| Customer name split | If col 3 non-empty → use as surname; else rsplit col 2 on last space |
| `01/01/1900` dates | Treated as null throughout |

---

## Schema Impact

Two columns added to `customers` (included in the main schema):
- `musipos_account_code text` — original Musipos customer ID
- `musipos_barcode_ref text` — Musipos internal barcode reference

These are read-only import reference fields, not shown in the regular customer management UI.

---

## Import Wizard UI (4 steps)

1. **Select Files** — browse for each source file
2. **Preview & Warnings** — counts + warnings per source (unresolved suppliers, no-email customers, etc.)
3. **Dry Import** (optional) — exports preview CSVs for review before committing
4. **Commit** — batch insert with progress bars; summary + import log on completion

---

## Post-Import Note

After commit, a persistent banner appears:
> *"Import complete. All items have been set to Qty On Hand = 0. A full stocktake is required
> before the system goes live."*

---

## APIC (Deferred)

APIC (Australian Publishers and Importers Catalogue) — access method and export format unknown.
Investigation needed before this can be planned. See Plan 09 for open questions.

---

## Source Files

```
src/importer/
    supplier_importer.py   — PDF parser, fuzzy match
    inventory_importer.py  — CSV reader, field mapping, FK resolution
    customer_importer.py   — headerless CSV, name split, garbage filter
    import_wizard.py       — step-by-step UI dialog
```
