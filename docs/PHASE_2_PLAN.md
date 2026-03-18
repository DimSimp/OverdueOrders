# Phase 2 — Daily Operations: Plan

## Context

The app currently handles "afternoon operations" — matching received inventory against overdue orders, then booking freight. Phase 2 adds a **home screen launcher** and a full **Daily Operations** workflow that covers morning processing: fetching all pending orders, classifying envelope orders, interactively assigning pick zones, generating a printable picking list, and then handing the full order list off to the existing freight-booking results view.

Phase 2 is broken into sub-phases implemented incrementally:

| Sub-phase | Status | Description |
|-----------|--------|-------------|
| 2a — Foundation | ✅ Done | Home screen, DailyOpsWindow skeleton, Options, Fetch |
| 2b — Envelopes | 🔲 Pending | Envelope classification + PDF printing |
| 2c — Picking | 🔲 Pending | Pick zone assignment + picking list CSV |
| 2d — Results | 🔲 Pending | Full order list with freight booking |

---

## Part A — Home Screen Launcher

### Architecture Decision
Home screen is embedded in the existing `App` window. On normal launch (no `.scar` arg), the window starts at 500×400 showing a home frame with two large buttons. Selecting a mode either transitions the same window (Afternoon Operations → resizes to 1150×720, builds 3-tab UI) or opens a new `DailyOpsWindow` (CTkToplevel).

- If a `.scar` file is passed as argument → skip home screen, open directly to Afternoon Operations
- `DailyOpsWindow` shares `neto_client` and `ebay_client` from the parent `App`

### Files (implemented)
| File | Purpose |
|------|---------|
| `src/gui/home_window.py` | `HomeFrame` widget — two mode-selection buttons |
| `src/gui/app.py` | Modified: `_build_home_screen()`, `_enter_afternoon_mode()`, `_enter_daily_mode()` |

---

## Part B — Daily Operations Window

### Window Structure
`DailyOpsWindow(CTkToplevel)` — 1150×720, opens on top of home screen.
Uses a **stacked-frame navigation model** — each step is a `CTkFrame` placed at the same grid position; `_show_step()` raises the target frame to the top.

### Step Flow (6 Screens)

```
Step 1: Options
  → Step 2: Fetch Orders (background thread)
    → Step 3: Envelope Classification (interactive)
      → Step 4: Envelope PDF Preview/Print
        → Step 5: Pick Zone Classification (interactive)
          → Step 6: Picking List (CSV preview + export)
            → Step 7: Results (order list, freight booking)
```

Navigation: `_show_step(frame, step_text)` raises frames. A persistent header shows the current step name and step counter (e.g. "Step 2 of 6").

---

## Phase 2a — Foundation (✅ Complete)

### Step 1 — Options (`src/gui/daily_ops/options_view.py`)

Controls:
- **Platform toggles** (CTkSwitch): Website, eBay (via Neto), BigW, Kogan, Amazon, eBay (direct)
- **Filter toggles**: "Include Express orders", "Include Click & Collect orders"
- **Date range**: defaults to today only (vs afternoon ops which uses 90-day lookback)
- **Note filter phrases**: scrollable add/remove list; saved to `config.app.note_filter_phrases`
- **"Generate List" button** → proceeds to Step 2

All toggle states persist to `config.app.daily_ops_toggles`.

### Step 2 — Fetch (`src/gui/daily_ops/fetch_view.py`)

- Parallel threads: Neto + eBay workers (mirrors `OrdersTab` logic)
- Progress bar + per-platform status labels
- Post-fetch filters: express, click & collect, note-phrase exclusion (`exclude_phrases()`)
- Results stored in `DailyOpsWindow.neto_orders` / `ebay_orders`

### Supporting changes (all complete)
| File | Change |
|------|--------|
| `src/config.py` | Added `note_filter_phrases: list` + `daily_ops_toggles: dict` to `AppConfig`; backward-compat loading |
| `src/data_processor.py` | `filter_on_po()` now accepts `str \| list`; new `exclude_phrases()` helper |
| `config.json` | Added `"note_filter_phrases": ["on po"]` and `"daily_ops_toggles": {}` |
| `src/gui/daily_ops/__init__.py` | Package init |
| `src/gui/daily_ops/daily_ops_window.py` | Main window + step navigation + placeholder for steps 3–7 |

---

## Phase 2b — Envelope Workflow (🔲 Pending)

### Step 3 — Envelope Classification (`src/gui/daily_ops/envelope_view.py`)

**Automatic pre-classification:**
- Only single-line-item orders are envelope candidates
- Neto: check `PostageType` field on order/line for "Minilope" / "Devilope" → auto-classify
- eBay: check `PrivateNotes` for "Mini." or "Devil." → auto-classify
- All others → added to interactive queue
- Multi-item orders → skip, go straight to freight list

**Interactive UI (one order at a time):**
- Order card: item image (thumb), SKU, description, customer name, destination state
- Existing PostageType shown as greyed note if present
- Three buttons: **Minilope** / **Devilope** / **Neither**
- On selection: updates `PostageType` in Neto via API; "Neither" → sets to "Satchel"
- Progress counter: "3 of 12 remaining"
- "Skip for now" defers the order
- Auto-advances to Step 4 when queue is empty

> **Open question**: Is `PostageType` on the Neto order line (add to `GetOrder` OutputSelector) or on the product item (requires separate `GetItem` call)? Test by adding `"PostageType"` to the OutputSelector and checking the response.

### Step 4 — Envelope PDF Generation (`src/envelope_pdf.py` + `envelope_view.py`)

**PDF layout:**
- Page size: A5 portrait (148 × 210 mm)
- Text rotated 90° clockwise (envelope fed landscape through printer)
- Content:
  ```
  [large]  Recipient Name
           Company (if any)
           Street Line 1
           Street Line 2
           City  STATE  Postcode

  [small, corner]  Order: N123456
  ```
- Library: **reportlab** (add to `requirements.txt` if not already present)
- One multi-page PDF per type: `Minilopes_{YYYY-MM-DD}.pdf`, `Devilopes_{YYYY-MM-DD}.pdf`
- Saved to `{config.app.lists_dir}/{date}/`

**Step 4 UI:**
- Count label per type
- "Open Minilopes PDF" / "Open Devilopes PDF" → `os.startfile(path)`
- "Print Minilopes" / "Print Devilopes" → `os.startfile(path, 'print')`
- "Next: Pick Zones →"

> **Open question**: Minilopes and Devilopes — same A5 page size, or does Devilope need a different size (e.g. A4)? Page size is parameterised so easy to adjust.

**New Neto method needed:**
```python
def update_item_postage_type(self, order_id_or_sku: str, postage_type: str, dry_run: bool = False) -> None:
    # UpdateOrder or UpdateItem depending on where PostageType lives
```

---

## Phase 2c — Picking Workflow (🔲 Pending)

### Step 5 — Pick Zone Classification (`src/gui/daily_ops/pick_zone_view.py`)

**Background pre-pass (runs while step loads):**
- Collect all unique SKUs from ALL orders (including envelope orders — staff still pick them)
- Call `neto_client.get_items_pick_zones(skus)` → `{sku: zone_or_None}`
- Valid zones: "String Room", "Back Area", "Out Front", "Picks"
- SKUs already assigned a valid zone → pre-classified
- Others → interactive queue

**Interactive UI (one SKU at a time):**
- Shows: SKU, description, current zone value (if any)
- Four zone buttons: **String Room** / **Back Area** / **Out Front** / **Picks**
- On click: calls `neto_client.update_item_pick_zone(sku, zone)` (UpdateItem API)
- "Skip" → leaves zone blank (appears in "Undefined" section of picking list)
- Progress: "4 of 9 remaining"
- Auto-advances to Step 6 when queue is empty or all skipped

**New Neto methods needed:**
```python
def get_items_pick_zones(self, skus: list[str]) -> dict[str, str | None]:
    # GetItem with Filter SKU in list + OutputSelector ["SKU", "PickZone"]
    # Returns {sku: zone_str_or_None}

def update_item_pick_zone(self, sku: str, zone: str, dry_run: bool = False) -> None:
    # UpdateItem — sets PickZone field on product
```

### Step 6 — Picking List (`src/picking_list.py` + `src/gui/daily_ops/pick_list_view.py`)

**Generation logic (`src/picking_list.py`):**
```python
def generate_picking_list(
    orders: list,               # NetoOrder + EbayOrder (all orders, including envelopes)
    pick_zones: dict[str, str], # {sku: zone}
    zone_order: list[str],      # ["String Room", "Back Area", "Out Front", "Picks"]
) -> list[dict]:
    # Groups by SKU (sum qty across all orders/listings)
    # Sort: zone order first, then SKU alphabetically within zone
    # Returns [{sku, description, qty, zone}, ...]
```

Note: same SKU can appear across multiple listings (e.g. "2221" = single pack AND "2221" × 5). Quantities are merged — one pick-list row per unique SKU.

**Step 6 UI:**
- Scrollable preview treeview (SKU / Description / QTY / Zone)
- Sections separated by zone heading rows
- "Export CSV" → saves to `{lists_dir}/{date}/Picking List {YYYY-MM-DD}.csv`, opens folder
- "Next: Orders →"

CSV columns: `SKU, Description, QTY, Zone`

---

## Phase 2d — Results View (🔲 Pending)

### Step 7 — Orders List (`src/gui/daily_ops/results_view.py`)

**Architecture:**
New `DailyOpsResultsView(CTkFrame)` — simpler than `ResultsTab` (no invoice matching concept). All fetched orders (including envelopes) go into one unified list.

Reused components (no changes needed):
- `OrderTreeview` — the existing scrollable order list widget
- `OrderDetailView` — stacked detail frame (click order to open)
- `FreightBookingView` — stacked freight booking frame

Additional features:
- Envelope orders tagged visually (e.g. "ENV" badge or different row colour in Shipping column)
- "Refresh All Orders" → re-fetches by order IDs
- "Export CSV" → exports current order list
- "Save Session" → saves `{date} daily ops session.scar` with `"mode": "daily_ops"` field
- Shared overrides sync (same `{date} overrides.json` mechanism as afternoon ops)

**Session restore:**
When a `.scar` file with `"mode": "daily_ops"` is opened, restore directly to Step 7 (skip Steps 1–6). Add `mode` field support to `src/session.py`.

---

## Config Changes Summary

### config.json (already added)
```json
"app": {
  "note_filter_phrases": ["on po"],
  "daily_ops_toggles": {}
}
```

### src/config.py (already done)
- `AppConfig.note_filter_phrases: list` — falls back to `on_po_filter_phrase` if absent
- `AppConfig.daily_ops_toggles: dict`

---

## Full File Inventory

### New files
| File | Phase | Purpose |
|------|-------|---------|
| `src/gui/home_window.py` | 2a ✅ | HomeFrame widget |
| `src/gui/daily_ops/__init__.py` | 2a ✅ | Package init |
| `src/gui/daily_ops/daily_ops_window.py` | 2a ✅ | Main window + step navigation |
| `src/gui/daily_ops/options_view.py` | 2a ✅ | Step 1 — options |
| `src/gui/daily_ops/fetch_view.py` | 2a ✅ | Step 2 — fetch |
| `src/gui/daily_ops/envelope_view.py` | 2b | Steps 3 & 4 — classify + print |
| `src/envelope_pdf.py` | 2b | PDF generation for envelopes |
| `src/gui/daily_ops/pick_zone_view.py` | 2c | Step 5 — pick zone assignment |
| `src/picking_list.py` | 2c | Picking list generation logic |
| `src/gui/daily_ops/pick_list_view.py` | 2c | Step 6 — picking list preview + export |
| `src/gui/daily_ops/results_view.py` | 2d | Step 7 — orders list + freight |

### Modified files
| File | Phase | Changes |
|------|-------|---------|
| `src/gui/app.py` | 2a ✅ | Home screen support |
| `src/config.py` | 2a ✅ | `note_filter_phrases`, `daily_ops_toggles` |
| `src/data_processor.py` | 2a ✅ | `filter_on_po()` phrase list; `exclude_phrases()` |
| `config.json` | 2a ✅ | New app keys |
| `src/neto_client.py` | 2b/2c | `update_item_postage_type()`, `get_items_pick_zones()`, `update_item_pick_zone()` |
| `src/session.py` | 2d | `mode` field for daily ops sessions |

---

## Verification Checklist

### Phase 2a
- [x] Home screen shows on normal launch
- [x] "Afternoon Operations" enters 3-tab UI; direct `.scar` open skips home
- [x] "Daily Operations" opens `DailyOpsWindow`
- [x] Options toggles persist across restarts
- [x] Fetch respects platform/express/C&C toggles
- [x] "on PO" orders are excluded from fetch results

### Phase 2b
- [ ] Single-item orders with existing PostageType auto-classify correctly
- [ ] Interactive classification updates Neto PostageType (verify in backend)
- [ ] eBay PrivateNotes "Mini."/"Devil." patterns are detected
- [ ] Minilope + Devilope PDFs generate with correct layout
- [ ] PDFs open and print from the app

### Phase 2c
- [ ] Pick zones loaded from Neto product catalogue
- [ ] Interactive zone assignment updates Neto (verify in backend)
- [ ] Picking list CSV: correct SKU merging, correct zone sorting, correct quantities
- [ ] Undefined items appear in their own section

### Phase 2d
- [ ] All orders (including envelopes) appear in results view
- [ ] Envelope orders visually tagged
- [ ] Freight booking works from Daily Ops results
- [ ] Session file saves/restores results state (Step 7 only)
- [ ] Shared overrides sync works across two instances
