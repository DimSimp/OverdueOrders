# Phase 2 — Daily Operations: Plan

## Context

The app currently handles "afternoon operations" — matching received inventory against overdue orders, then booking freight. Phase 2 adds a **home screen launcher** and a full **Daily Operations** workflow that covers morning processing: fetching all pending orders, classifying envelope orders, interactively assigning pick zones, and then processing the full order list in a unified results/dispatch screen.

Phase 2 is broken into sub-phases implemented incrementally:

| Sub-phase | Status | Description |
|-----------|--------|-------------|
| 2a — Foundation | ✅ Done | Home screen, DailyOpsWindow skeleton, Options, Fetch |
| 2b — Envelopes | ✅ Done | Envelope classification + PDF printing |
| 2c — Pick Zones | ✅ Done | Pick zone assignment (Step 5) |
| 2cd — Results | ✅ Done | Combined results/dispatch screen (Step 6) |

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
          → Step 6: Results & Dispatch
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
- **Neto product attribute lookup** (`get_item_attributes(skus)`): fetches `Misc06` (PostageType), `ShippingCategory`, and `PickZone` for all unique SKUs across both Neto and eBay orders in a single `GetItem` call
  - Populates `li.postage_type` on Neto line items from Misc06
  - Filters out Books orders (ShippingCategory="4") from both platforms
  - Stores result in `window.sku_attr_map: dict[str, dict]` for use in Steps 3 and 5
- Results stored in `DailyOpsWindow.neto_orders` / `ebay_orders`

### Supporting changes (all complete)
| File | Change |
|------|--------|
| `src/config.py` | Added `note_filter_phrases: list` + `daily_ops_toggles: dict` to `AppConfig`; backward-compat loading |
| `src/data_processor.py` | `filter_on_po()` now accepts `str \| list`; new `exclude_phrases()` helper |
| `config.json` | Added `"note_filter_phrases": ["on po"]` and `"daily_ops_toggles": {}` |
| `src/gui/daily_ops/__init__.py` | Package init |
| `src/gui/daily_ops/daily_ops_window.py` | Main window + step navigation |

---

## Phase 2b — Envelope Workflow (✅ Complete)

### Step 3 — Envelope Classification (`src/gui/daily_ops/envelope_view.py`)

**Automatic pre-classification:**
- Only single-line-item orders are envelope candidates
- For all orders (Neto + eBay): check `sku_attr_map` for Misc06 PostageType → auto-classify Minilope/Devilope/Satchel
- eBay orders: PrivateNotes keywords ("Mini."/"Devil.") are a lower-priority fallback
- Books category (ShippingCategory="4") orders already filtered at Step 2 — will not appear
- Multi-item orders → skip (go straight to freight list at Step 6)

**Interactive UI (one order at a time):**
- Order card: 200×200 product image, SKU, description, platform + order ID + price, customer + state
- Neto: shows existing PostageType from product catalogue; eBay: shows Neto product PostageType if available
- Four buttons: **Minilope** / **Devilope** / **Neither (Satchel)** / **Books**
- On selection:
  - Minilope/Devilope/Satchel: `UpdateItem` sets `Misc06` on Neto product (persists for future sessions)
  - Books: `UpdateItem` sets `ShippingCategory = "4"` on Neto product; order is removed from `window.neto_orders`/`ebay_orders` entirely — does not appear in pick list or results
- Progress counter; "Skip for now" defers the order

**Key design note:** Write-back on classification targets the **Neto product catalogue** (not the order), so eBay orders benefit equally — the next time the same SKU appears it will auto-classify.

### Step 4 — Envelope PDF Generation (`src/envelope_pdf.py` + `envelope_view.py`)

**PDF layout (implemented):**
- Page size: A5 portrait (148 × 210 mm), content rotated 90° CW via canvas transform
- Address block: left-aligned, vertically centred
- Postcode: 28pt bold, bottom-right
- Order ID: 9pt grey, bottom-right (no "Order:" prefix)
- Library: **reportlab**
- Two PDFs: `Minilopes.pdf`, `Devilopes.pdf` (overwritten each run)
- Saved to `\\SERVER\Project Folder\Order-Fulfillment-App\Envelopes\`

**Step 4 UI:**
- Count label per type
- "Open" + "Print" buttons per type
- "Open Envelopes Folder" button → opens network share in Explorer
- Print via **SumatraPDF** (`-print-to-default -print-settings "1x,mono" -silent`); fallback to `start /print`
- SumatraPDF search paths: PATH, Program Files, Program Files (x86), `%LOCALAPPDATA%\SumatraPDF\`

---

## Phase 2c — Pick Zone Assignment (✅ Complete)

### Step 5 — Pick Zone Classification (`src/gui/daily_ops/pick_zone_view.py`)

**Pre-classification (no additional API call):**
- Reads `window.sku_attr_map` (already populated in Step 2) for `pick_zone` on each unique SKU
- Valid zones: `"String Room"`, `"Back Area"`, `"Out Front"`, `"Picks"`
- SKUs already in a valid zone → silently pre-classified into `window.pick_zones`
- Others (no zone or unrecognised value) → interactive queue

**Interactive UI (one unique SKU at a time):**
- 200×200 product image (fetched via `get_product_images()`), SKU, description, order count, current zone value
- Four zone buttons (colour-coded): **String Room** (blue) / **Back Area** (grey) / **Out Front** (light grey) / **Picks** (green)
- On click: `UpdateItem` sets `PickZone` on Neto product (background thread); stores in `window.pick_zones`
- "Skip for now" → SKU appears in "Unassigned" section of picking list
- Progress counter; auto-advances to Step 6 when queue empty or all skipped

**Neto methods (implemented in `src/neto_client.py`):**
- `get_item_attributes(skus)` — now returns `pick_zone` in addition to `postage_type` + `shipping_category`
- `update_item_pick_zone(sku, zone, dry_run)` — `UpdateItem` with `PickZone` field

---

## Phase 2cd — Results & Dispatch (🔲 Pending)

Steps 6 and 7 from the original plan are **combined into a single Step 6** that replicates the afternoon operations results screen with daily-ops-specific additions.

### Step 6 — Results & Dispatch (`src/gui/daily_ops/results_view.py`)

#### Overall architecture

`DailyOpsResultsView(CTkFrame)` contains:
- A **two-tab CTkTabview**: "Active Orders" | "Removed"
- `OrderDetailView` overlaid at grid (0,0) when an order is clicked (same stacking pattern as afternoon ops)
- `FreightBookingView` overlaid on top of `OrderDetailView` when freight is opened

Both tabs reuse `OrderTreeview` from `src/gui/results_tab.py` with no changes to that widget.

#### Column spec (both tabs)

| Column | Content |
|--------|---------|
| Order ID | Order ID |
| Platform | Neto / eBay |
| Customer | Ship name |
| State | Ship state |
| Shipping | Express / Regular / Local Pickup |
| Type | Minilope / Devilope / Satchel / — (from `window.envelope_classifications`) |
| Zone | Pick zone (from `window.pick_zones`; "—" if unassigned) |
| Total | Grand total |

#### Filters

All existing afternoon ops filters (Platform, Shipping type), plus:
- **Envelopes only** — shows only orders where envelope_classifications value is `"minilope"` or `"devilope"`

#### Context menu

- **Active Orders tab**: right-click order → "Move to Removed"
- **Removed tab**: right-click order → "Move back to Active"

Both moves: update `_removed_order_ids` set, save overrides to disk, refresh both tabs.

#### Toolbar / action buttons

| Button | Behaviour |
|--------|-----------|
| Refresh | Re-fetches all listed orders by ID via `get_orders_by_ids()`; merges shared overrides; redraws |
| Cancel / Reprint Booking | Opens the same cancel/reprint dialog from afternoon ops |
| Export Picking List | Generates XLSX picking list (see below); opens save folder |
| Print Pick Labels | Sends pick labels to Brother QL printer for all "Picks" zone items (see below) |

#### Reused components (no changes required)

| Component | Source file |
|-----------|-------------|
| `OrderTreeview` | `src/gui/results_tab.py` |
| `OrderDetailView` | `src/gui/order_detail_view.py` |
| `FreightBookingView` | `src/gui/freight_booking_view.py` |
| Cancel/reprint dialog | `src/gui/results_tab.py` (extract or call directly) |

---

### Session & Overrides (new module: `src/session_daily.py`)

#### Paths (fixed, always overwrite — no date prefix)

| File | Path |
|------|------|
| Session | `\\SERVER\Project Folder\Order-Fulfillment-App\Session\Daily\CURRENT DAILY SESSION.scar` |
| Overrides | `\\SERVER\Project Folder\Order-Fulfillment-App\Session\Daily\DAILY OVERRIDES.json` |

Both paths are hardcoded (not config-driven) since only one daily session exists at a time.

#### Auto-save triggers

- On first load of Step 6 (creates/overwrites session)
- After every "Move to Removed" / "Move back to Active"
- After every fulfilled / dispatched order (triggered by `on_fulfilled` callback)

#### Session file format (`.scar`)

```json
{
  "version": 1,
  "timestamp": "2026-03-20T09:15:00.000000",
  "neto_orders": [ /* serialised NetoOrder objects */ ],
  "ebay_orders":  [ /* serialised EbayOrder objects */ ],
  "envelope_classifications": { "NTO-12345": "minilope" },
  "pick_zones": { "SKU001": "String Room", "SKU002": "Picks" },
  "removed_order_ids": [ ["Neto", "NTO-99999"], ["eBay", "EBY-88888"] ]
}
```

#### Overrides file format (`.json`)

```json
{
  "removed_order_ids": [
    ["Neto", "NTO-99999"],
    ["eBay", "EBY-88888"]
  ]
}
```

**Multi-user sync**: On every Refresh, the overrides file is read from the network share and merged with local state before redrawing — same mechanism as afternoon ops. Changes from one workstation propagate to another on the next refresh.

#### `src/session_daily.py` functions

```python
DAILY_SESSION_DIR = r"\\SERVER\Project Folder\Order-Fulfillment-App\Session\Daily"
DAILY_SESSION_FILE = "CURRENT DAILY SESSION.scar"
DAILY_OVERRIDES_FILE = "DAILY OVERRIDES.json"

def save_daily_session(window: DailyOpsWindow) -> None:
    # Serialise all window state → JSON → write to DAILY_SESSION_DIR/DAILY_SESSION_FILE

def load_daily_session(path: str) -> dict:
    # Read and parse a .scar file; returns raw dict for caller to restore state

def save_daily_overrides(removed_ids: set[tuple[str, str]]) -> None:
    # Write DAILY_OVERRIDES.json (overwrites)

def load_daily_overrides() -> set[tuple[str, str]]:
    # Read DAILY_OVERRIDES.json; returns set of (platform, order_id) tuples
    # Returns empty set if file missing or malformed
```

---

### Export: Picking List XLSX

**Format decision**: XLSX (not plain CSV) — allows programmatic page breaks between zones so each zone section starts on a new printed page.

**Trigger**: "Export Picking List" button → save dialog defaults to:
`\\SERVER\Project Folder\Order-Fulfillment-App\Picking Lists\Picking List {YYYY-MM-DD}.xlsx`

**Generation logic** (`src/picking_list.py`):
```python
ZONE_ORDER = ["String Room", "Back Area", "Out Front", "Picks"]  # then Unassigned

def generate_picking_list(
    orders: list,               # Active (non-removed) NetoOrder + EbayOrder
    pick_zones: dict[str, str], # {sku: zone}
) -> list[dict]:
    """
    Groups line items by SKU across all active orders.
    Sums quantities for duplicate SKUs.
    Sorts: ZONE_ORDER first, then SKU alphabetically within zone.
    Returns [{sku, description, qty, zone}, ...]
    """
```

**XLSX structure (per zone)**:
- Bold section-header row: zone name (e.g. "String Room")
- Data rows: SKU | Description | QTY
- A **page break** is inserted before each new zone section (except the first)
  - Implemented via `openpyxl` `ws.row_breaks.append(Break(id=row_idx))`
- "Unassigned" section at the end for SKUs with no pick zone

**Note on duplicate SKUs**: Same SKU can appear across multiple orders (e.g. "2221" ordered as single AND as × 5). Quantities are summed — one row per unique SKU in the output.

---

### Pick Labels — Brother QL (`src/pick_labels.py`)

**Purpose**: Print a small sticky label for every item in the "Picks" pick zone (guitar picks, plectrums, etc.) so staff can attach them to the plastic baggies used for packaging.

**Trigger**: "Print Pick Labels (N)" button in results toolbar (N = total label count including qty expansion). Button disabled if no "Picks" zone items exist in active orders.

**Input data**: Derived from active orders + `window.pick_zones`:
```python
# Build label list
for order in active_orders:
    for li in order.line_items:
        if window.pick_zones.get(li.sku) == "Picks":
            for _ in range(li.quantity):
                labels.append((li.sku, li.product_name or li.title))
```
Result: one label entry per individual item (e.g. 2 packs of 6 = 2 label entries).

**Label content**: `[SKU: {sku}] {description}` — identical format to existing script.

**Image generation** (adapted from `Currently_Used_Code/Pick_labels.py`):
- PIL: `Image.new("RGB", (240, height), (255, 255, 255))`
- Font: `Calibri 18pt` (path: `C:\Windows\Fonts\calibri.ttf`)
- Text wrapped at 30 chars per line (`re.sub("(.{30})", "\\1\n", ...)`)
- Height calculated from line count: `22 * (nlines + 1)`
- 62mm tape width, `brother_ql` library

**Printer detection** (same auto-discover logic):
- `brother_ql.backends.helpers.discover('pyusb')[0]` → identifier
- Supports QL-700 (`usb://0x04F9:0x2042`) and QL-570 (`usb://0x04F9:0x2028`)
- Falls back to QL-570 on any error

**`src/pick_labels.py` interface**:
```python
def build_label_list(
    orders: list,
    pick_zones: dict[str, str],
) -> list[tuple[str, str]]:
    """Returns [(sku, description), ...] — one entry per quantity unit."""

def print_pick_labels(
    labels: list[tuple[str, str]],
    progress_callback=None,
) -> None:
    """
    Generates and prints one Brother QL label per entry.
    progress_callback(printed: int, total: int) called after each label.
    Raises PrintLabelError on printer discovery failure.
    """

class PrintLabelError(Exception):
    pass
```

**UI flow**:
1. Button click → compute label list → show confirmation: "Print N pick labels?" (Yes/No)
2. On Yes → run `print_pick_labels()` in background thread
3. Button text changes to "Printing… (N remaining)" during print
4. On completion: button resets; toast/status message shown
5. On error: messagebox with error text; button resets

**Dependency**: `brother_ql` — add to `requirements.txt` if not already present. The PyInstaller spec may also need `brother_ql` hidden imports added.

---

## Navigation & Step Counter Updates

Step 6 is now the final step. All step counters in `daily_ops_window.py` need updating from "X of 7" to "X of 6".

After wiring Step 6:
- `_on_pick_zone_complete()` → `_show_results()`
- Back button on Step 6 → `_show_pick_zone()`
- No "Next" button — Step 6 is the terminal step

---

## Full File Inventory

### New files
| File | Phase | Status | Purpose |
|------|-------|--------|---------|
| `src/gui/home_window.py` | 2a | ✅ | HomeFrame widget |
| `src/gui/daily_ops/__init__.py` | 2a | ✅ | Package init |
| `src/gui/daily_ops/daily_ops_window.py` | 2a | ✅ | Main window + step navigation |
| `src/gui/daily_ops/options_view.py` | 2a | ✅ | Step 1 — options |
| `src/gui/daily_ops/fetch_view.py` | 2a | ✅ | Step 2 — fetch + Books filter + attr map |
| `src/gui/daily_ops/envelope_view.py` | 2b | ✅ | Steps 3 & 4 — classify + print |
| `src/envelope_pdf.py` | 2b | ✅ | PDF generation for envelopes |
| `src/gui/daily_ops/pick_zone_view.py` | 2c | ✅ | Step 5 — pick zone assignment |
| `src/gui/daily_ops/results_view.py` | 2cd | ✅ | Step 6 — results + dispatch |
| `src/session_daily.py` | 2cd | ✅ | Daily ops session + overrides helpers |
| `src/picking_list.py` | 2cd | ✅ | XLSX picking list generation |
| `src/pick_labels.py` | 2cd | ✅ | Brother QL pick label printing |

### Modified files
| File | Phase | Status | Changes |
|------|-------|--------|---------|
| `src/gui/app.py` | 2a | ✅ | Home screen support |
| `src/gui/freight_booking_view.py` | 2b | ✅ | Order Items section; PO Box → AusPost-only |
| `src/config.py` | 2a | ✅ | `note_filter_phrases`, `daily_ops_toggles` |
| `src/data_processor.py` | 2a | ✅ | `filter_on_po()` phrase list; `exclude_phrases()` |
| `config.json` | 2a | ✅ | New app keys |
| `src/neto_client.py` | 2b+2c | ✅ | `get_item_attributes()` (PostageType + ShippingCategory + PickZone); `update_item_postage_type()`; `update_item_shipping_category()`; `update_item_pick_zone()` |

---

## Verification Checklist

### Phase 2a ✅
- [x] Home screen shows on normal launch
- [x] "Afternoon Operations" enters 3-tab UI; direct `.scar` open skips home
- [x] "Daily Operations" opens `DailyOpsWindow`
- [x] Options toggles persist across restarts
- [x] Fetch respects platform/express/C&C toggles
- [x] "on PO" orders are excluded from fetch results

### Phase 2b ✅
- [x] Books orders (ShippingCategory=4) filtered at Step 2 for both Neto and eBay
- [x] eBay orders checked against Neto product catalogue for auto-classification
- [x] Write-back on classification targets Neto product (works for both platforms)
- [x] Books button in Step 3 removes order from all subsequent workflow steps
- [x] Minilope + Devilope PDFs generate with correct layout (A5 portrait, rotated 90° CW)
- [x] PDFs saved to `\\SERVER\...\Envelopes\` as `Minilopes.pdf` / `Devilopes.pdf`
- [x] "Open Envelopes Folder" button works
- [x] PDFs print directly via SumatraPDF (B&W, A5, no dialog)

### Phase 2c ✅
- [x] Pick zones loaded from `sku_attr_map` (no extra API call at Step 5)
- [x] Interactive zone assignment updates Neto `PickZone` via `UpdateItem`
- [ ] Verify `PickZone` is the correct Neto API field name (confirm during testing)

### Phase 2cd (Pending)
- [ ] Step 6 loads all active orders; auto-saves session on arrival
- [ ] "Active Orders" and "Removed" tabs both functional
- [ ] Right-click "Move to Removed" / "Move back to Active" updates both tabs + overrides file
- [ ] Overrides sync across two workstations (test with both open)
- [ ] OrderDetailView opens correctly from both tabs (including Removed)
- [ ] FreightBookingView works from detail view
- [ ] "Refresh" re-fetches by order ID and merges shared overrides
- [ ] Cancel/Reprint dialog works
- [ ] "Envelopes only" filter shows only Minilope/Devilope orders
- [ ] XLSX picking list: correct SKU merging, correct zone sort order, page breaks between zones
- [ ] XLSX saved to correct server path
- [ ] Pick label count on button is accurate (qty-expanded)
- [ ] Pick labels print correctly to Brother QL printer
- [ ] One label per quantity unit (2 packs = 2 labels)
- [ ] Session file overwrites on each save (no date-accumulation)
- [ ] All step counters show "X of 6"
