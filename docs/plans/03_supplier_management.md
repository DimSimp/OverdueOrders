# Plan 03 — Supplier Management

> **Part of**: [Master Plan](00_overview.md)
> **Status**: 🔲 Not started
> **Phase**: 3 — POS Core (but schema needed in Phase 1 to back the `supplier_id` FK on `items`)

---

## Overview

Manages all supplier records — contact details, trading terms, SKU suffix/prefix rules, and a full history of purchase orders and invoices per supplier. Also serves as the entry point for processing incoming stock via the AI-assisted invoice import pipeline.

---

## Key Concepts

| Concept | Description |
|---------|-------------|
| **Supplier ID** | Short code identifying the supplier (e.g. `DUNLOP`, `ERNIE`). Matches Musipos supplier codes for compatibility. Used as the primary key. |
| **Sales Rep** | The individual contact at the supplier for ordering and account queries |
| **Support Contact** | General support/internal sales email and phone for the supplier company |
| **Payment Terms** | Number of days after invoice date that payment is due. Default: 30 days. |
| **SKU Suffix / Prefix** | String appended/prepended to a raw SKU to form the Web SKU used on Neto/eBay. Currently in `config.json` — to be migrated to this table. |
| **AI Invoice Parser** | Coworker-maintained tool that reads scanned/uploaded supplier PDF invoices and outputs structured data as a CSV and writes to Musipos SQL |

---

## Database Schema

### `suppliers` — Supplier master record

| Column | Type | Notes |
|--------|------|-------|
| `id` | text PK | Supplier code (e.g. `DUNLOP`). Matches Musipos. |
| `name` | text NOT NULL | Full supplier name |
| `abn` | text | Australian Business Number |
| `account_number` | text | Our account number with this supplier |
| `payment_terms_days` | integer DEFAULT 30 | Days after invoice date that payment is due |
| `sku_suffix` | text | Appended to raw SKU to form Web SKU. Migrated from `config.json`. |
| `sku_prefix` | text | Prepended to raw SKU. Migrated from `config.json`. |
| `character_substitutions` | jsonb | Character replacement rules before suffix (e.g. remove `/`). Migrated from `config.json`. |
| `address_line1` | text | |
| `address_line2` | text | |
| `city` | text | |
| `state` | text | |
| `postcode` | text | |
| `notes` | text | Free-form internal notes about this supplier |
| `active` | boolean DEFAULT true | Inactive suppliers are hidden from new PO creation |

---

### `supplier_contacts` — Named contacts per supplier

Supports multiple contacts per supplier. Typical setup: one Sales Rep + one general Support entry.

| Column | Type | Notes |
|--------|------|-------|
| `id` | uuid PK | |
| `supplier_id` | text FK → `suppliers.id` | |
| `role` | text | e.g. `Sales Rep`, `Support`, `Accounts` |
| `name` | text | |
| `email` | text | |
| `phone` | text | |
| `is_primary` | boolean DEFAULT false | Primary contact shown prominently in the supplier card |

---

> **Note on Purchase Orders and Invoices**: These are stored in the `purchase_orders` and `invoices` tables defined in [04_purchasing_receiving.md](04_purchasing_receiving.md). The Supplier Management module displays them in tabs on the supplier card but does not own the schema.

---

## UI Design

### Supplier List View

A simple searchable list — not a split-panel, since there are relatively few suppliers and each one opens a full detail card.

```
┌──────────────────────────────────────────────────────┐
│  [Search: ID, Name, Account No.]     [+ New Supplier]│
├────────────┬──────────────────────┬──────────┬───────┤
│ Supplier ID│ Name                 │ Acct No. │ Terms │
├────────────┼──────────────────────┼──────────┼───────┤
│ DUNLOP     │ Dunlop Manufacturing │ SC-04821 │ 30d   │
│ ERNIE      │ Ernie Ball           │ 10042    │ 30d   │
│ ...        │ ...                  │ ...      │ ...   │
└────────────┴──────────────────────┴──────────┴───────┘
```

- Clicking a row opens the Supplier Card (separate dialog window)
- "New Supplier" button opens a blank Supplier Card in edit mode
- Inactive suppliers hidden by default; toggle to show all

---

### Supplier Card (Detail Window)

Opens as a resizable dialog. Header shows Supplier ID + Name prominently. Contains four tabs:

```
┌─────────────────────────────────────────────────────┐
│  DUNLOP — Dunlop Manufacturing                      │
│  ┌─────────┬──────────┬───────────────┬──────────┐  │
│  │ Details │ Contacts │ Purchase Orders│ Invoices │  │
│  └─────────┴──────────┴───────────────┴──────────┘  │
│  [Tab content]                                      │
│                                          [Edit] [✕] │
└─────────────────────────────────────────────────────┘
```

---

#### Tab 1: Details
- Supplier ID, Full Name, ABN, Account Number
- Payment Terms (e.g. "30 days")
- Address (full)
- SKU Suffix, SKU Prefix, Character Substitutions (read-only display; edit via Edit button)
- Notes field
- Active toggle (admin only)

---

#### Tab 2: Contacts
Displays all contacts for this supplier in a small table:

| Role | Name | Email | Phone | Primary |
|------|------|-------|-------|---------|

- "Add Contact" button → inline row or small form
- Edit / Delete per row
- Primary contact flagged with an indicator

---

#### Tab 3: Purchase Orders
List of all POs raised for this supplier, most recent first.

| PO Number | Date Raised | Status | Items | Total Value | Expected Delivery |
|-----------|-------------|--------|-------|-------------|-------------------|

- Status values: `Draft`, `Sent`, `Partially Received`, `Received`, `Cancelled`
- Click a PO row → opens the PO detail view (see plan 04)
- "New PO" button → opens a new PO for this supplier (see plan 04)
- Filter by status

---

#### Tab 4: Invoices
List of all invoices received from this supplier, most recent first.

| Invoice No. | Date | PO Number | Total (inc GST) | Due Date | Status |
|-------------|------|-----------|-----------------|----------|--------|

- Status values: `Unpaid`, `Paid`, `Overdue`
- Click an invoice row → opens Invoice detail view (see plan 04)
- "Import Invoice" button → triggers the AI invoice import flow (see below)
- **"Receive Without PO"** button → creates an ad-hoc PO and opens the Receive Invoice window directly (see [04_purchasing_receiving.md](04_purchasing_receiving.md))
- Filter by status (Unpaid / Overdue surface at the top automatically)

---

### Global Invoice / PO Search

Accessible from the main app toolbar (not just within a supplier card), so staff can jump directly to a known invoice or PO number without navigating through the supplier list first.

- Single search field accepts either an invoice number or PO number
- Returns matching record(s) with supplier name — click to open the relevant detail view

---

## AI Invoice Import Integration

### What the Web Portal System Is
The coworker's tool (`C:\VB\Web Portal`) is a full procurement system, not just a simple CSV exporter. It is:

- **Backend**: Python FastAPI server (`web_portal.py`) running locally, accessible via ngrok on mobile
- **Desktop GUI**: CustomTkinter app (`musipos_bulk.py`) for bulk processing at the counter
- **AI Engine**: Claude API (`claude-sonnet-4-20250514`) for PDF and photo parsing
- **Multi-channel ingestion**: Phone photos (via ngrok), PDF upload, IMAP email auto-fetch

When an invoice is processed, the system:
1. Parses the PDF/image via Claude — extracts supplier, invoice no., PO no., and all line items
2. Resolves SKUs against the Musipos SQL database (with fuzzy matching + manual override dialogs)
3. Runs a **dry-run preview** for staff confirmation
4. On confirmation: writes to Musipos SQL (PO lines, inventory, accounts payable) and outputs a daily received CSV

The current OverdueOrders app already consumes the daily received CSV (`daily_reports/received_YYYY-MM-DD.csv`) for dispatch comparison. **That integration is preserved unchanged.**

---

### Parsed Invoice Data Structure
Claude extracts the following per invoice (from `models.py: ParsedPDFInvoice`):

**Header**: `supplier_name`, `invoice_number`, `invoice_date`, `po_reference`, `subtotal`, `gst`, `total`, `ship_to_address`, `is_dropship`, `tracking_reference`, `carrier`

**Per line item**:
| Field | Description |
|-------|-------------|
| `sku` | Item code as it appears on the invoice |
| `description` | Item description |
| `qty_supplied` | Quantity actually shipped on this delivery |
| `qty_backordered` | Quantity placed on backorder (0 if none) |
| `unit_cost` | Wholesale/dealer cost (NOT RRP) |
| `line_total` | `unit_cost × qty_supplied` |
| `is_product_line` | False for shipping/freight/surcharge lines |
| `category` | `product`, `shipping`, `freight`, `packing`, `handling`, `surcharge`, `other` |

> **Note**: Cost is supplied as a single `unit_cost` (ex-GST dealer price). GST is an invoice-level field, not per-line. RRP is **not** included in the parsed output — it is not extracted from invoices.

---

### Integration Strategy

The Web Portal exposes a full FastAPI at `http://localhost:8050` with endpoints including:

| Endpoint | Purpose |
|----------|---------|
| `POST /api/parse-pdf` | Submit a PDF for Claude parsing |
| `POST /api/validate-skus` | Resolve SKUs against Musipos DB |
| `POST /api/import` | Execute invoice import (supports `dry_run=true`) |
| `GET /api/suppliers` | List known suppliers |

This gives us **two integration options**:

**Option A — CSV consumption (current, simple)**
After the Web Portal commits an invoice, it writes `daily_reports/received_YYYY-MM-DD.csv`. OverdueOrders reads this file and uses it to update Supabase inventory (qty, costs). No changes needed to the Web Portal.

**Option B — Direct API call (future, tighter)**
OverdueOrders calls `/api/import?dry_run=true` to get a preview, shows it to staff, then calls `/api/import` for real. On success, also commits the same data to Supabase. Eliminates the CSV middleman.

**Recommended approach**: Port the core parsing and SKU resolution logic from the Web Portal directly into OverdueOrders (see below), replacing the existing inferior parser in `src/pdf_parser.py`. The Web Portal remains a separate project — we borrow its logic, not its server infrastructure.

---

### What to Port from the Web Portal

OverdueOrders already has a basic invoice parser (`src/pdf_parser.py`) that is significantly less capable than the Web Portal's. The Web Portal will be the source of the following logic, adapted and integrated into OverdueOrders:

| Web Portal Module | What to Port | What to Skip |
|-------------------|-------------|--------------|
| `pdf_parser.py` | Claude-based PDF/image parsing logic, multi-page batching, multi-invoice detection | Phone photo two-pass flow (ngrok-specific), image preprocessing |
| `validators.py` | SKU resolution against inventory DB, fuzzy matching | Musipos SQL lookups (replace with Supabase) |
| `sku_mapping.py` | Persistent SKU correction store (JSON), user prompt for unknown SKUs | As-is — reuse directly |
| `supplier_import.py` | Invoice import core logic, line item processing, backorder detection | Musipos SQL writes (replace with Supabase), accounts payable (plan 04) |
| `models.py` | `ParsedPDFInvoice`, `SupplierInvoiceLine` dataclasses | Sales import models (out of scope) |

**Not ported** (Web Portal-specific infrastructure):
- `web_portal.py` — FastAPI server, ngrok, session management
- `email_fetch.py` — IMAP email invoice fetching
- `po_sender.py` / `portal_ordering.py` — PO sending and supplier portal automation
- `musipos_bulk.py` — their desktop GUI (we build our own)
- `db.py` — Musipos SQL connection (we use Supabase instead)

The existing `src/pdf_parser.py` in OverdueOrders will be **replaced** by the ported logic.

---

### Supabase Update After Invoice Receipt
When an invoice is processed and confirmed by staff:

1. Parsed line items matched to Supabase inventory by SKU (using ported SKU resolution logic)
2. Update inventory: `qty_on_hand` +qty, `last_purchase_cost`, `last_purchase_date`, recalculate average costs
3. Update `qty_on_order` on the matched PO (decremented by qty received)
4. Create a supplier bill record in Supabase with `due_date = invoice_date + payment_terms_days`
5. Write `stock_movements` record of type `receive`
6. Trigger customer special order notifications for any SKUs with waiting customer allocations (see plan 06)
7. Append received lines to `daily_reports/received_YYYY-MM-DD.csv` (preserves existing dispatch comparison workflow)

**Daily CSV columns** (unchanged, for compatibility with existing dispatch flow):
```
supplier_id, supplier_invoice_no, po_no, sku, qty_received, unit_cost, description
```

---

### Overlap Note
The Web Portal's `order_fetch.py` also fetches eBay and Neto orders — overlapping with the existing OverdueOrders functionality. These are separate code paths serving different purposes (the Web Portal uses order data for dropship PO creation; OverdueOrders uses it for dispatch management). **No consolidation is needed now**, but this is worth coordinating with the coworker to avoid conflicting writes if both systems ever touch the same Neto/eBay data simultaneously.

---

### config.json Migration (SKU Suffix/Prefix)

Currently `config.json` stores per-supplier suffix/prefix rules used by the existing app. Once the suppliers table is live in Supabase:

1. The `sku_suffix`, `sku_prefix`, and `character_substitutions` columns on `suppliers` become the source of truth
2. The existing app's `src/config.py` will be updated to read these values from Supabase instead of the local file
3. `config.json` supplier suffix/prefix entries will be deprecated (but left in place as fallback during transition)

**This migration should happen at the same time the inventory system goes live**, since `web_sku` on the `items` table is derived from these rules.

---

## Implementation Checklist

### Infrastructure
- [ ] Create `suppliers` table in Supabase
- [ ] Create `supplier_contacts` table in Supabase
- [ ] Populate initial supplier records from Musipos export / `config.json`
- [ ] Populate `sku_suffix`, `sku_prefix`, `character_substitutions` from existing `config.json` entries

### Data Layer (`src/suppliers/`)
- [ ] `supplier_client.py` — CRUD for suppliers and contacts
  - [ ] `get_all_suppliers(include_inactive=False)`
  - [ ] `get_supplier(supplier_id)`
  - [ ] `create_supplier(data)`
  - [ ] `update_supplier(supplier_id, data)`
  - [ ] `get_contacts(supplier_id)`
  - [ ] `add_contact(supplier_id, contact_data)`
  - [ ] `update_contact(contact_id, data)`
  - [ ] `delete_contact(contact_id)`
  - [ ] `get_sku_rules(supplier_id)` → suffix, prefix, substitutions (for use by inventory web_sku logic)

### Invoice Parsing (ported from Web Portal)
- [ ] Port `pdf_parser.py` Claude-based parsing logic into `src/invoice/pdf_parser.py`
  - [ ] Short PDF path (≤3 pages, single Claude call)
  - [ ] Long PDF batching (3-page chunks, merge results)
  - [ ] Multi-invoice detection (discovery prompt + per-invoice extraction)
  - [ ] Replace existing `src/pdf_parser.py` (inferior version)
- [ ] Port `models.py` dataclasses (`ParsedPDFInvoice`, `SupplierInvoiceLine`) into `src/invoice/models.py`
- [ ] Port `sku_mapping.py` persistent SKU correction store into `src/invoice/sku_mapping.py`
  - [ ] Adapt DB lookups: replace Musipos SQL with Supabase inventory queries
- [ ] Port `validators.py` SKU resolution + fuzzy matching into `src/invoice/validators.py`
  - [ ] Adapt to resolve against Supabase `items` table instead of Musipos SQL

### Invoice Import (`src/invoice/importer.py`)
- [ ] Port core import logic from Web Portal `supplier_import.py`
  - [ ] Line item processing (product lines vs freight/surcharge)
  - [ ] Backorder detection
  - [ ] Unknown SKU flagging + user correction prompt (preserve Web Portal UX)
  - [ ] Replace Musipos SQL writes with Supabase inventory receive hook (plan 01)
- [ ] PO matching (look up PO in Supabase by `po_no`)
- [ ] Discrepancy detection (PO qty vs received qty)
- [ ] Deduplication: don't re-import an already-processed invoice number
- [ ] Commit: write invoice record to Supabase, trigger stock/cost updates, create supplier bill
- [ ] Append received lines to `daily_reports/received_YYYY-MM-DD.csv` (preserves dispatch comparison)

### GUI (`src/gui/suppliers/`)
- [ ] `supplier_list_view.py` — searchable supplier list
- [ ] `supplier_card.py` — detail dialog with 4 tabs
- [ ] `supplier_details_tab.py`
- [ ] `supplier_contacts_tab.py`
- [ ] `supplier_po_tab.py` — PO list (view only; create/edit in plan 04)
- [ ] `supplier_invoices_tab.py` — invoice list + Import Invoice button
- [ ] `invoice_import_dialog.py` — file picker → preview table → confirm
- [ ] Global invoice/PO search widget (toolbar-level)

### config.json Migration
- [ ] Update `src/config.py` to fall back to Supabase for suffix/prefix rules
- [ ] Update existing app's SKU suffix logic to prefer DB over config
- [ ] Test existing dispatch/matching flow still works after migration
- [ ] Deprecate supplier suffix/prefix entries in `config.json`

---

## Open Questions / Future Considerations

- **Web Portal logic sync**: As the coworker continues improving the Web Portal's parsing and SKU resolution, we should periodically review for improvements worth back-porting into OverdueOrders. Keep the ported modules clearly separated so updates are straightforward to apply.
- **Supplier price lists**: Some suppliers provide periodic price list updates (CSV or APIC feed) that update RRP and cost across all their items. A bulk price update tool may be needed — deferred to the APIC import module (plan 09).
- **Multi-currency**: All suppliers are currently AUD. If this changes, a `currency` field on the supplier record and exchange rate handling will be needed.
- **Supplier portal access**: Some suppliers (e.g. larger distributors) have online portals for placing orders. Integration with these (e.g. auto-submit a PO via their API) is possible long-term but out of scope for now.

---

*Last updated: 2026-04-01*
