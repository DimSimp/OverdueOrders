# Plan 09 — APIC / Musipos Import

> **Part of**: [Master Plan](00_overview.md)
> **Status**: 🔲 Not started
> **Phase**: 1 — Foundation

---

## Overview

A one-time migration tool to populate the new Supabase database from Musipos exports. Three data sets need importing:

1. **Inventory** — from `musipos_inventory.CSV` (CSV, clean structure, full export)
2. **Suppliers** — from `Suppliers.PDF` (PDF only, awkward format, all suppliers not just active ones)
3. **Customers** — from `Customer_List.CSV` (CSV, no header row, inconsistent data quality)

APIC (the shared Australian music industry product catalogue) is noted as a future ongoing import source, but the format and access method are currently unknown and deferred for investigation.

The import tool is a **one-time migration wizard** built into the app's Settings/Admin area. It is not exposed to regular staff. It previews mappings and shows a summary of warnings before committing anything to Supabase, so it can be run safely in review mode first.

---

## Scope

| Source | What we import | What we skip |
|--------|---------------|-------------|
| Inventory CSV | Active items (`Active = Y`), all useful columns | `Active = N` items (optionally importable as inactive); `Product_Type`, `Category`, `Department`, `Quantity_Availability` always ignored |
| Supplier PDF | Name + address for suppliers referenced in inventory CSV | All suppliers with no matching `Supplier_ID` in inventory |
| Customer CSV | Active customers with a usable name and phone | Inactive (`Active = N`); garbage records (name is only dashes, stars, or other special characters) |
| APIC | Deferred — import methodology unknown | — |

---

## Data Source Analysis

### Inventory CSV

**File**: `musipos_inventory.CSV`
**Format**: CSV with header row, one item per row
**Encoding**: UTF-8 (verify on export — Musipos may use Windows-1252)
**Notable issues**:
- `Quantity_on_hand` values are **not imported** — Musipos stock counts are known to be inaccurate. All items are imported with `qty_on_hand = 0`. A full stocktake (plan 01) is required before the system goes live.
- `Minimum_Sell = 0` on some items — treat 0 as "no minimum set" (null), not a literal $0 floor.
- `Barcode` column (e.g. `930035012465`) is Musipos's internal barcode. `Product_Barcode` is the manufacturer EAN/UPC and is often empty.
- `Supplier_ID` values like `PAYTONS` and `JACAR` must be resolved to the new `suppliers.id` in Supabase. This lookup is seeded by the supplier import step — **inventory must be imported after suppliers**.

**Column mapping**:

| Musipos Column | Supabase Field | Notes |
|---------------|----------------|-------|
| `Supplier_Item_ID` | `items.sku` | Primary key equivalent |
| `Title` | `items.title` | |
| `Supplier_RRP` | `items.supplier_rrp` | |
| `Publisher_Brand` | `items.brand` | |
| `Artist_Composer_Series` | `items.series` | |
| `Instrument` | `items.instrument` | |
| `Sub_Instrument` | `items.sub_instrument` | |
| `Product_Type` | — | Always `MUSICAL` — ignored |
| `Quantity_on_hand` | — | **Ignored.** Musipos stock counts are known to be inaccurate. All items imported with `qty_on_hand = 0`. Full stocktake required before go-live. |
| `Last_Purchase_Cost` | `items.last_purchase_cost` | Also seed `average_cost_exc_gst` |
| `Barcode` | `items.internal_barcode` | Musipos-generated barcode |
| `Minimum_Sell` | `items.minimum_sell` | 0 → null (no minimum set) |
| `Category` | — | Always `NONE` — ignored |
| `Department` | — | Redundant with `Instrument` — ignored |
| `Last_Purchase_Date` | — | Not stored; used only for import stats |
| `Last_Sold_Date` | — | Not stored; used only for import stats |
| `Created_Date` | — | Not stored |
| `Stock_Availability_from_Supplier` | — | Ignored |
| `Supplier_ID` | `items.supplier_id` | Resolved via supplier lookup table built during supplier import |
| `Quantity_Availability` | — | Ignored |
| `Product_Barcode` | `items.product_barcode` | Manufacturer EAN/UPC; often empty |
| `Active` | Filter only | `N` → skip (or import as inactive if option selected) |

**Fields left empty after import** (set manually or from subsequent workflows):
`web_sku`, `pick_zone`, `online_sale_price`, `gst_amount`, `average_cost_inc_gst`, `min_order_level`, `max_order_level`, `is_kit`, `is_serialised`, shipping dimensions, images

---

### Supplier PDF

**File**: `Suppliers.PDF`
**Format**: Two-page PDF table — data split across pages (name/address on page 1, state/postcode on page 2)
**Notable issues**:
- **All suppliers are included** — not just ones the store deals with. Approximately 44 rows visible across both pages.
- Column text is truncated in the PDF rendering (e.g. `MUSIP` instead of `MUSIPOS`, `NATIONA` instead of `NATIONAL MUSIC SUPPLIES`).
- No phone, email, ABN, payment terms, or account number data — these must be entered manually after import.
- The two-page split means row 1 of page 1 and row 1 of page 2 are the same supplier. The pdfplumber extraction must pair rows by position.
- Some `Supplier_Id` values in the PDF appear slightly truncated compared to the inventory CSV (e.g. PDF shows `PAYTON`, inventory uses `PAYTONS`). A fuzzy match or manual confirmation step handles this.

**Column mapping**:

| PDF Column | Supabase Field | Notes |
|-----------|----------------|-------|
| `Supplier Id` | `suppliers.id` (as the lookup key) | Used to match inventory `Supplier_ID` values |
| `Supplier Name` | `suppliers.name` | May be truncated; correct manually if needed |
| `Address1` | `suppliers.address_1` | |
| `Address2` | `suppliers.address_2` | |
| `City` | `suppliers.city` | |
| `State` (page 2) | `suppliers.state` | |
| `Post Code` (page 2) | `suppliers.postcode` | |
| `Distribution Address*` (page 2) | — | Ignored — not in our schema |

**Fields requiring manual entry after import**:
`phone`, `email`, `abn`, `account_number`, `payment_terms_days`, `sku_suffix`, `sku_prefix`, `character_substitutions`, contacts

**Import strategy**: Only import suppliers whose `Supplier Id` appears in the unique set of `Supplier_ID` values from the inventory CSV. Approximately 15–20 suppliers will be relevant. The remaining ~25 are skipped.

**Known active suppliers** (from inventory sample — cross-reference to confirm full list):
`PAYTONS`, `JACAR`, `AMS`, `PRO`, `AUSTRA`, `ELECTR`, `CMI`, `ROLAND`, `YAMAHA`, `FENDER`, `CASIO`, `MATON`, `KAWAI`

---

### Customer CSV

**File**: `Customer_List.CSV`
**Format**: CSV with **no header row**, 26 columns
**Encoding**: Likely Windows-1252 (test on load)
**Notable issues**:
- No header row — column positions must be hard-coded based on the mapping below.
- Many garbage records where the name fields contain only dashes (`---`), asterisks (`***`), or single special characters. These should be filtered out automatically.
- Dates stored as `01/01/1900` are Musipos null placeholders — treat as null.
- Phone number is a single field — no distinction between mobile and landline. Import to `mobile`; staff can correct if needed.
- The "category" field (col 17, e.g. `GENERAL CUSTOMER`, `GUITARIST`, `PIANO`) is not meaningful enough to map to a new field — used only as an import filter note.
- Column 19 (`110000088274`) is Musipos's internal customer barcode. Store in `musipos_barcode_ref` for cross-referencing historical data.

**Column mapping** (0-indexed):

| Col # | Sample Value | Supabase Field | Notes |
|-------|-------------|----------------|-------|
| 0 | `(UF00001` | `customers.musipos_account_code` | Musipos account code; not used in new system but preserved for reference |
| 1 | `(UFO ARCADIA PTYLTD)` | `customers.business` | Often a business or account name; sometimes garbage |
| 2 | `TONY VALENT` | Name field (see split logic below) | May be first name only, or first + last |
| 3 | (empty) | `customers.surname` | Secondary name field; often empty |
| 4 | `2 PEARCE RIDGE` | `customers.address_1` | May be `NA`, `-`, or empty — treat as null |
| 5 | (empty) | `customers.address_2` | |
| 6 | `WINTHROP` | `customers.city` | |
| 7 | `WA` | `customers.state` | |
| 8 | `6150` | `customers.postcode` | |
| 9 | (empty) | — | Unknown; ignored |
| 10 | (empty) | — | Unknown; ignored |
| 11 | `MR` | — | Title/salutation; not in our schema — ignored |
| 12 | `tvalentc@hotmail.com` | `customers.email` | |
| 13 | `0412987111` | `customers.mobile` | Single phone field; import to mobile |
| 14 | (empty) | — | Unknown; ignored |
| 15 | (empty) | — | Unknown; ignored |
| 16 | `GUITARIST` | — | Customer category from Musipos; ignored |
| 17 | `28/06/2018` | — | Last transaction date; not stored |
| 18 | `110000088274` | `customers.musipos_barcode_ref` | Musipos internal barcode number |
| 19 | (empty) | — | Ignored |
| 20 | `01/01/1900` | — | Null placeholder date; ignored |
| 21 | (empty) | — | Ignored |
| 22 | (empty) | — | Ignored |
| 23 | `Y` | Filter only | Active flag — `N` rows skipped |
| 24 | `Y` | — | Unknown flag; ignored |
| 25 | `01/05/2018` | `customers.created_at` | `01/01/1900` → null |

**Name split logic** (col 2 → `first_name` / `surname`):
- If col 3 is not empty → `first_name = col 2`, `surname = col 3`
- If col 3 is empty and col 2 contains a space → split on last space: everything before = `first_name`, last word = `surname`
- If col 3 is empty and col 2 has no space → `first_name = col 2`, `surname = null`

**Garbage record filter** — skip rows where col 2 (the name field) matches:
- Starts with 3+ consecutive dashes (`---`)
- Starts with 3+ consecutive asterisks (`***`)
- Contains only special characters / whitespace after stripping alphanumerics
- Is empty

**Address normalisation**: Col 4–8 values of `NA`, `-`, `*`, `**`, `NA NA` → set to null

**Schema additions** (two extra columns on `customers` for import cross-reference):
- `musipos_account_code` text — the original Musipos customer ID (e.g. `(UF00001`)
- `musipos_barcode_ref` text — the Musipos internal barcode number

These are read-only reference fields, not exposed in the regular UI.

---

## Import Tool UI

The import wizard lives in **Settings → Data Import** (admin-only). It is a stepped dialog:

```
Step 1: Select Files
  [Browse] Inventory CSV:   musipos_inventory.CSV   ✓ loaded (8,432 rows)
  [Browse] Supplier PDF:    Suppliers.PDF           ✓ loaded (44 rows)
  [Browse] Customer CSV:    Customer_List.CSV       ✓ loaded (12,800 rows)

  [Next →]

─────────────────────────────────────────────────────────────

Step 2: Preview & Warnings
  SUPPLIERS (will import 18 of 44)
    ✓ 18 suppliers matched to inventory Supplier_ID values
    ⚠ 4 supplier names appear truncated in PDF — review recommended
    ⚠ 26 suppliers not referenced in inventory — skipped

  INVENTORY (will import 7,840 of 8,432)
    ✓ 18 supplier IDs resolved
    ⚠ 592 items marked Active = N — skipped
                (☐ Also import inactive items as inactive)
    ℹ All items will be imported with Qty On Hand = 0 — stocktake required before go-live
    ⚠ 14 items have unresolvable Supplier_ID — will import with supplier = null
    ℹ 341 items have Minimum_Sell = 0 → stored as null (no minimum set)

  CUSTOMERS (will import 9,241 of 12,800)
    ✓ Active filter applied
    ⚠ 1,847 records skipped — inactive (Active = N)
    ⚠ 1,712 records skipped — garbage name detected
    ⚠ 3,890 records have no email address
    ⚠ 623 records have no phone number

  [← Back]  [Run Dry Import]  [Commit to Supabase →]

─────────────────────────────────────────────────────────────

Step 3: Dry Import Review (optional)
  Exports a preview CSV for each table showing exactly what will be written.
  Review these before committing.
  [Download suppliers_preview.csv]
  [Download inventory_preview.csv]
  [Download customers_preview.csv]

  [← Back]  [Commit to Supabase →]

─────────────────────────────────────────────────────────────

Step 4: Committing...
  [████████████░░░░░░░░] Suppliers: 18 / 18
  [████░░░░░░░░░░░░░░░░] Inventory: 1,240 / 7,840
  [░░░░░░░░░░░░░░░░░░░░] Customers: 0 / 9,241

  ✓ Commit complete
     Suppliers: 18 imported, 4 need manual review
     Inventory: 7,840 imported, 14 with null supplier
     Customers: 9,241 imported
  [View Import Log]  [Done]
```

---

## Import Logic

### Supplier Import

```python
def import_suppliers(pdf_path, inventory_csv_path):
    # 1. Get unique Supplier_ID values from inventory CSV
    active_supplier_ids = get_unique_supplier_ids(inventory_csv_path)

    # 2. Parse PDF with pdfplumber — extract table rows across both pages
    # Pages are aligned by row index; pair page 1 row N with page 2 row N
    page1_rows, page2_rows = parse_supplier_pdf(pdf_path)
    suppliers = merge_pdf_pages(page1_rows, page2_rows)

    # 3. Match each PDF row to an active_supplier_id (fuzzy match on Supplier_Id column)
    matched, unmatched = match_suppliers(suppliers, active_supplier_ids)

    # 4. Return matched suppliers for preview / commit
    return matched, unmatched
```

**Fuzzy matching**: Strip trailing 'S' from PDF Supplier_Id before comparing to inventory value (handles PAYTON vs PAYTONS). Fall back to Levenshtein distance ≤ 2 for any remaining mismatches.

### Inventory Import

```python
def import_inventory(csv_path, supplier_lookup):
    # supplier_lookup: {musipos_supplier_id: supabase_supplier_uuid}
    for row in read_csv(csv_path):
        if row['Active'] == 'N' and not import_inactive:
            continue
        item = {
            'sku': row['Supplier_Item_ID'],
            'title': row['Title'],
            'supplier_rrp': decimal(row['Supplier_RRP']),
            'brand': row['Publisher_Brand'] or None,
            'series': row['Artist_Composer_Series'] or None,
            'instrument': row['Instrument'] or None,
            'sub_instrument': row['Sub_Instrument'] or None,
            'qty_on_hand': int(row['Quantity_on_hand']),
            'last_purchase_cost': decimal(row['Last_Purchase_Cost']),
            'average_cost_exc_gst': decimal(row['Last_Purchase_Cost']),
            'internal_barcode': row['Barcode'] or None,
            'minimum_sell': decimal(row['Minimum_Sell']) or None,  # 0 → None
            'product_barcode': row['Product_Barcode'] or None,
            'supplier_id': supplier_lookup.get(row['Supplier_ID']),  # None if unresolved
            'is_active': row['Active'] == 'Y',
        }
        yield item
```

**Duplicate handling**: If a SKU already exists in Supabase (re-running import), prompt: Update existing / Skip / Abort.

### Customer Import

```python
GARBAGE_NAME_PATTERN = re.compile(r'^[-*\s#!@$%^&()\[\]{}]{3,}$')
NULL_DATE = '01/01/1900'

def import_customers(csv_path):
    for row in read_csv_no_header(csv_path):
        if row[23] == 'N':  # Active flag
            continue
        name_raw = row[2].strip()
        if GARBAGE_NAME_PATTERN.match(name_raw) or not name_raw:
            continue

        # Name split
        surname_col = row[3].strip()
        if surname_col:
            first_name, surname = name_raw, surname_col
        elif ' ' in name_raw:
            parts = name_raw.rsplit(' ', 1)
            first_name, surname = parts[0], parts[1]
        else:
            first_name, surname = name_raw, None

        created_at = parse_date(row[25]) if row[25] != NULL_DATE else None

        def null_if_placeholder(val):
            return None if val.strip() in ('', 'NA', '-', '*', '**', 'NA NA', 'N/A') else val.strip()

        customer = {
            'first_name': first_name.title(),
            'surname': surname.title() if surname else None,
            'business': null_if_placeholder(row[1]),
            'address_1': null_if_placeholder(row[4]),
            'address_2': null_if_placeholder(row[5]),
            'city': null_if_placeholder(row[6]),
            'state': null_if_placeholder(row[7]),
            'postcode': null_if_placeholder(row[8]),
            'email': row[12].strip() or None,
            'mobile': row[13].strip() or None,
            'created_at': created_at,
            'musipos_account_code': row[0].strip(),
            'musipos_barcode_ref': row[18].strip() or None,
        }
        yield customer
```

---

## Recommended Import Order

1. **Suppliers first** — inventory needs supplier UUIDs to resolve FKs
2. **Inventory second** — uses supplier lookup built in step 1
3. **Customers third** — no dependencies, can run independently but easier after the others are confirmed

---

## APIC Import (Deferred)

APIC (Australian Publishers and Importers Catalogue) is a shared music industry product database used for adding new stock items without manual data entry. The access method and export format are currently unknown — needs investigation with the APIC system administrator or supplier contacts.

**Questions to answer before this can be planned:**
- Does APIC provide a direct data feed (API, FTP, or scheduled email)?
- What file format does an APIC export use (CSV, XML, EDI)?
- Is data exported per-supplier or as a full catalogue?
- How are updates delivered (full refresh vs delta)?
- What fields does the APIC catalogue include vs what's in Musipos inventory?

**Likely approach once investigated**: A periodic import that adds new SKUs from APIC without overwriting existing Supabase data. Existing items are matched by SKU and optionally updated if RRP or title has changed. New items are added in a "pending review" state so staff can confirm before they go live in the inventory.

---

## Schema Additions

Two reference columns to add to `customers` (not exposed in regular UI):

| Column | Type | Notes |
|--------|------|-------|
| `musipos_account_code` | text | Original Musipos account code e.g. `(UF00001` |
| `musipos_barcode_ref` | text | Musipos internal barcode number e.g. `110000088274` |

---

## Implementation Checklist

### Schema
- [ ] Add `musipos_account_code` and `musipos_barcode_ref` columns to `customers` table

### Import Tool (`src/importer/`)
- [ ] `supplier_importer.py` — PDF parser, fuzzy matching against inventory Supplier_IDs
- [ ] `inventory_importer.py` — CSV reader, column mapping, supplier FK resolution, duplicate handling
- [ ] `customer_importer.py` — headerless CSV reader, name split logic, garbage filter, placeholder nullification
- [ ] `import_wizard.py` — step-by-step UI dialog (Settings → Data Import, admin only)

### Supplier Import
- [ ] Parse PDF with pdfplumber — align page 1 and page 2 rows by index
- [ ] Extract unique Supplier_IDs from inventory CSV
- [ ] Fuzzy match PDF supplier IDs to inventory IDs (strip trailing S, Levenshtein ≤ 2)
- [ ] Preview step: show matched / unmatched / skipped with warnings for truncated names
- [ ] Dry run: export `suppliers_preview.csv`
- [ ] Commit: insert to Supabase `suppliers` table
- [ ] Post-import: flag suppliers with truncated names for manual review in the app

### Inventory Import
- [ ] CSV reader with encoding detection (UTF-8 / Windows-1252)
- [ ] Active filter (`Active = Y` default; checkbox to include inactive)
- [ ] Supplier FK lookup: `{musipos_id: supabase_uuid}`
- [ ] Zero Minimum_Sell → null
- [ ] Preview step: show counts and warnings (unresolved suppliers, zero qty count)
- [ ] Dry run: export `inventory_preview.csv`
- [ ] Commit: batch insert to Supabase `items` table (chunks of 500 for performance)
- [ ] Duplicate SKU handling: Update / Skip / Abort prompt

### Customer Import
- [ ] Headerless CSV reader with encoding detection
- [ ] Garbage name filter (regex)
- [ ] Active filter (col 23 = Y)
- [ ] Name split logic (col 3 present → use; else split col 2 on last space)
- [ ] Address placeholder normalisation (NA, -, * → null)
- [ ] `01/01/1900` date → null
- [ ] Preview step: show imported / skipped counts with reasons
- [ ] Dry run: export `customers_preview.csv`
- [ ] Commit: batch insert to Supabase `customers` table
- [ ] Sequential `customer_id` assignment starting from 1 (or from max+1 if partial re-import)

### Post-Import
- [ ] Import log: write summary to `data/import_logs/YYYY-MM-DD_HH-MM.json`
- [ ] In-app confirmation screen with counts and download links for preview CSVs
- [ ] Note in app after commit: *"Import complete. All items have been set to Qty On Hand = 0. A full stocktake is required before the system goes live."*

---

## Open Questions

- **Inactive items**: Import them as `is_active = false` to preserve history, or skip entirely? The preview step will offer this as a checkbox option.
- **Qty on hand**: Confirmed — Musipos stock counts are inaccurate and will not be imported. All items start at 0. Stocktake required before go-live. ✓ Resolved.
- **APIC access**: Who to contact, and what format it provides. Investigation needed before plan 09 can be considered fully complete.
- **Phone number type**: The customer CSV has one phone field. After import, are there known customers whose number is a landline rather than mobile? If so, a post-import cleanup step may be needed.

---

*Last updated: 2026-04-14*
