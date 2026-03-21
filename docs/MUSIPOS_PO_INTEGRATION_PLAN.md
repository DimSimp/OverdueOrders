# Musipos PO Integration Plan

## Overview

Add two per-line-item action buttons to `OrderDetailView` (both Daily Ops and Afternoon Ops):

| Button | Action |
|---|---|
| **Left in Waiting Area** | Adds order note: "SKU left in waiting area" (sticky note for Neto, private note for eBay) |
| **On PO** | Resolves item in Musipos DB → confirms qty + supplier → adds to PO → adds order note "SKU on PO" |

These buttons appear in a sub-row below each line item row, keeping the existing row layout unchanged.

---

## User Decisions

| Decision | Answer |
|---|---|
| SKU resolution strategy | Strip supplier suffix → cascade (itm_iid → itm_supplier_iid → barcode) → invoice_skus from sku_mappings.csv → musipos_sku_map.csv → manual entry prompt |
| When item not found | Prompt for manual Musipos SKU; if confirmed, save to musipos_sku_map.csv |
| PO quantity | Let user confirm in dialog (spinner, default = order line qty) |
| Supplier when ambiguous | Use primary supplier by default; show dropdown if item has multiple suppliers |
| Note text format | SKU only (e.g. "6091478 on PO", "6091478 left in waiting area") |
| PO number in note | No — note is just "SKU on PO" |

---

## Files

### New
| File | Purpose |
|---|---|
| `src/musipos_client.py` | DB connection, item resolution, PO creation |
| `src/gui/musipos_po_dialog.py` | Multi-step PO confirmation dialog |

### Modified
| File | Change |
|---|---|
| `src/config.py` | Add `MusiposConfig` dataclass; add `musipos` field to `ConfigManager` |
| `config.json` | Add `musipos` section |
| `requirements.txt` | Add `pyodbc>=4.0.35` |
| `src/gui/order_detail_view.py` | Accept `musipos_client`; add action sub-rows to line items |
| `src/gui/results_tab.py` | Pass `musipos_client` when constructing `OrderDetailView` |
| `src/gui/daily_ops/results_view.py` | Pass `musipos_client` when constructing `OrderDetailView` |
| `src/gui/app.py` | Create `MusiposClient` from config; pass to `ResultsTab` |

---

## 1. `requirements.txt`

Add:
```
pyodbc>=4.0.35
```

**Installation note**: Also requires SQL Server Native Client 10.0 ODBC driver (already installed per DB reference, as this driver is used by the existing po_ordering.py).

---
## 2. `src/config.py` — `MusiposConfig`

```python
@dataclass
class MusiposConfig:
    server: str               # e.g. "SERVER\\MUSIPOSSQLSRV08"
    database: str             # "musipos"
    user: str                 # "sa"
    password: str
    kit_mappings_path: str    # path to existing sku_mappings.csv (kit/invoice expansion map)
    musipos_map_path: str     # path to new musipos_sku_map.csv (neto_sku → musipos_sku)
    driver: str = "SQL Server Native Client 10.0"
    enabled: bool = True
    default_user_id: str = "MAN"  # written to pop_user_id on new PO lines
    computer_id: str = "19"       # written to pop_computer_id
    warehouse_id: str = "00001"   # always "00001"
```

Parsed in `ConfigManager._parse()` from `config.json["musipos"]`. If the key is absent, `self.musipos = None`.

---

## 3. `config.json` — `musipos` section

```json
"musipos": {
  "server": "SERVER\\MUSIPOSSQLSRV08",
  "database": "musipos",
  "user": "sa",
  "password": "YOUR_PASSWORD_HERE",
  "kit_mappings_path": "\\\\SERVER\\Project Folder\\Order-Fulfillment-App\\Inventory_Reports\\sku_mappings.csv",
  "musipos_map_path": "\\\\SERVER\\Project Folder\\Order-Fulfillment-App\\Inventory_Reports\\musipos_sku_map.csv",
  "enabled": true
}
```

---
## 4. `src/musipos_client.py`

### Class: `MusiposClient`

```python
class MusiposClient:
    def __init__(self, config: MusiposConfig): ...
```

#### `get_connection() -> pyodbc.Connection`
Returns a new pyodbc connection. Caller must close it.
```
Driver={SQL Server Native Client 10.0};Server=...;Database=...;UID=...;PWD=...
```

#### `test_connection() -> tuple[bool, str]`
Executes `SELECT 1`; returns `(True, "")` or `(False, error_message)`.

#### `resolve_item(neto_sku: str, suppliers: list[SupplierConfig]) -> dict | None`

Multi-strategy SKU resolution. Tries each strategy in order and returns the first match.

**Strategy 1 — Strip suffix + cascade DB lookup**

For each `SupplierConfig` with a non-empty suffix, check if `neto_sku` ends with (or starts with, if `suffix_position == "prepend"`) that suffix. If matched, `base_sku = stripped value`. If no suffix matches, `base_sku = neto_sku`.

Runs the cascade query with three column targets in order: `itm_iid`, `itm_supplier_iid`, `itm_barcode`:
```sql
SELECT TOP 1
    ap4itm.itm_iid, ap4itm.itm_title, ap4itm.itm_supplier_id,
    ap4itm.itm_supplier_iid, sp4qpc.qpc_qty_on_hand,
    sp4qpc.qpc_last_purchase_cost, ap4itm.itm_new_retail_price
FROM ap4itm
JOIN sp4qpc ON RTRIM(ap4itm.itm_iid) = RTRIM(sp4qpc.qpc_iid)
WHERE ap4itm.itm_lno = '0000'
  AND sp4qpc.qpc_warehouse_id = '00001'
  AND UPPER(RTRIM(ap4itm.<COLUMN>)) = UPPER(?)
```

**Strategy 2 — Kit/invoice mapping fallback (`sku_mappings.csv`)**

If Strategy 1 fails, look up `neto_sku` in the kit mappings CSV (`kit_mappings_path`) by the `neto_sku` column. If found, try each value in the `invoice_skus` column against `itm_iid` and `itm_supplier_iid`.

The CSV columns are: `neto_sku`, `is_kit`, `invoice_skus`, `qty_per_alias`, `supplier`. For kits, `invoice_skus` is pipe-delimited (`|`); single value otherwise.

**Strategy 3 — Musipos SKU map (`musipos_sku_map.csv`)**

If Strategies 1–2 fail, look up `neto_sku` in the new Musipos mapping CSV (`musipos_map_path`). Columns: `neto_sku`, `musipos_sku`. Try the `musipos_sku` value against `itm_iid`.

The `musipos_sku_map.csv` starts empty and grows as users confirm manual lookups. It is created at `musipos_map_path` with headers on first write.

Returns `{itm_iid, title, supplier_id, supplier_iid, qty_on_hand, last_cost, retail_price}` on first match, or `None` if all strategies fail.

**Known limitation**: Electric Factory Neto SKUs (e.g. `05MS201BPLUSEF`) have the `/` removed. After stripping the `EF` suffix, `base_sku = 05MS201BPLUS`. The Musipos `itm_iid` is `05/MS201BPLUS`. Strategy 1 cascade checks `itm_supplier_iid` as a fallback. Strategy 2 may find the item via invoice SKU. If neither works, the item falls through to Strategy 3 (manual map) or the manual entry prompt.

#### `resolve_item_by_musipos_sku(musipos_sku: str) -> dict | None`
Direct lookup by `itm_iid` only. Used for manual entry and Strategy 3 fallback.

#### `get_suppliers_for_item(itm_iid: str) -> list[dict]`
Returns `[{supplier_id, supplier_name}, ...]`. Queries:
- Primary: `ap4itm.itm_supplier_id`
- Alternate: `sp4qpc.qpc_alt_supplier_id` (if set)

Joins `ap4rsp` to get `rsp_name`.

#### `get_current_po(supplier_id: str) -> int | None`
```sql
SELECT TOP 1 phd_po_no FROM sp4phd
WHERE phd_supplier_id = ? AND phd_po_status = 'CURRENT'
ORDER BY phd_po_no DESC
```
Returns PO number or `None`.

#### `add_to_po(itm_iid: str, supplier_id: str, qty: int, dry_run: bool = False) -> dict`

Implements the PO creation workflow (DB reference §4.3). Uses a transaction.

Steps:
1. Find existing CURRENT PO, OR allocate new PO number from `ap4rsp.rsp_curr_po_no` with `UPDLOCK + HOLDLOCK`.
2. If new PO: `INSERT INTO sp4phd` with `phd_po_status = 'CURRENT'`, `phd_po_dest = 'S'`, `phd_memo = 'AIO ordering'`.
3. Check if item already on PO: `SELECT pop_qor FROM sp4pop WHERE pop_iid = ? AND pop_supplier_id = ? AND pop_po_no = ?`
4. If yes: `UPDATE sp4pop SET pop_qor = pop_qor + ? WHERE ...`
5. If no: `INSERT INTO sp4pop (pop_iid, pop_supplier_id, pop_po_no, pop_cid, pop_lno, ...)` with `pop_cid = 'STOCK001'`, `pop_curr_order_flag = 'N'`, `pop_qty_rcv = 0`, `pop_comment = 'AIO ordering'`.
6. `UPDATE sp4qpc SET qpc_qor = ISNULL(qpc_qor, 0) + ? WHERE qpc_iid = ? AND qpc_warehouse_id = '00001'`

If `dry_run=True`: wrap in a manual rollback (execute all SQL but rollback before commit).

Returns:
```python
{
    "po_no": int,
    "new_po": bool,
    "supplier_id": str,
    "supplier_name": str,
    "action": "added" | "updated",   # new line vs existing line updated
    "qty_added": int,
}
```

#### `load_kit_mappings() -> dict[str, list[str]]`
Reads `kit_mappings_path` (the existing `sku_mappings.csv`). Columns: `neto_sku`, `is_kit`, `invoice_skus`, `qty_per_alias`, `supplier`.

For kit items (`is_kit == "TRUE"`), `invoice_skus` is pipe-delimited (e.g. `SKU1|SKU2|SKU3`). For non-kit items it is a single value. Split on `|` in all cases.

Returns `{neto_sku_upper: [invoice_sku1, invoice_sku2, ...]}` — the component/invoice SKUs to try against Musipos.

If file not found or unreadable, returns `{}` and logs a warning.

#### `load_musipos_map() -> dict[str, str]`
Reads `musipos_map_path` (`musipos_sku_map.csv`). Columns: `neto_sku`, `musipos_sku`.
Returns `{neto_sku_upper: musipos_sku}`.

#### `save_musipos_alias(neto_sku: str, musipos_sku: str) -> None`
Appends a new row to `musipos_sku_map.csv`. Creates the file with a `neto_sku,musipos_sku` header if it doesn't exist.

---
## 5. `src/gui/musipos_po_dialog.py`

### Class: `MusiposPODialog(CTkToplevel)`

A multi-step dialog for the "On PO" workflow.

#### Constructor
```python
MusiposPODialog(
    parent,
    *,
    neto_sku: str,
    product_name: str,
    order_qty: int,
    musipos_client: MusiposClient,
    suppliers_config: list[SupplierConfig],
    dry_run: bool = False,
    on_success: callable | None = None,  # called with po_result dict
    on_note_only: callable | None = None,  # called if user cancels PO but still wants note
)
```

Geometry: `500×350`, title: `"Add to Purchase Order"`

#### Steps / States

**RESOLVING** (default on open)
- Indeterminate progress bar + "Resolving [SKU] in Musipos…"
- Runs background thread: calls `musipos_client.resolve_item(neto_sku, suppliers_config)`
- On result → transitions to CONFIRM or NOT_FOUND

**CONFIRM** (item found, single supplier)
```
Item:         [itm_title]                (label)
Musipos ID:   [itm_iid]                  (label)
Supplier:     [supplier_name]            (label, or dropdown if multiple)
Stock:        [qty_on_hand] on hand      (label, coloured red if 0)

Quantity to order: [−] [3] [+]          (stepper, default = order_qty)

[Add to PO]          [Cancel]
```

If multiple suppliers: replace "Supplier:" label with a `CTkOptionMenu` dropdown listing each.

**NOT_FOUND** (cascade lookup + alias CSV both failed)
```
"SKU not found in Musipos."
"You can enter the Musipos SKU manually:"

[  entry field  ]    [Search]

[Cancel — add note only]
```
- Search button triggers another background lookup via `resolve_item_by_musipos_sku()`
- On result → transitions to CONFIRM_ALIAS
- "Cancel — add note only" calls `on_note_only` callback (note still gets added, no PO)

**CONFIRM_ALIAS** (manual search returned a result)
```
Found:  [itm_iid] — [itm_title]
Supplier: [supplier_name]

Is this the correct item?

[Yes — add to PO and save to aliases]    [Try again]    [Cancel]
```
"Yes" saves the mapping via `musipos_client.save_sku_alias()`, then proceeds to add to PO.

**WORKING** (after user clicks "Add to PO")
- Progress bar + "Adding to PO…"

**DONE**
```
✓  Added [itm_iid] to PO #[po_no] for [supplier_name]
   (or: Updated existing PO #[po_no] — qty +N)
   [DRY RUN] label if dry_run=True

[Close]
```
Calls `on_success(po_result)` before showing this state.

**ERROR**
```
✗  [error_message]
[Close]
```

---
## 6. `src/gui/order_detail_view.py` — changes

### Constructor: add `musipos_client=None`

```python
def __init__(
    self,
    master,
    *,
    ...
    musipos_client=None,  # NEW
):
    ...
    self._musipos_client = musipos_client
```

### `_build_neto_line_items()` / `_build_ebay_line_items()` — pass line_item object

```python
# Neto
self._build_line_item_row(items_frame, li.sku, li.product_name,
                           li.quantity, li.unit_price, li.sku, line_item=li)

# eBay — note_row already exists below; action row also below
self._build_line_item_row(items_frame, li.sku, li.title,
                           li.quantity, li.unit_price, li.legacy_item_id, line_item=li)
```

### `_build_line_item_row()` — add `line_item=None` parameter, add action sub-row

After the existing row widget, if `line_item is not None`, build an action sub-row:

```python
action_row = ctk.CTkFrame(items_frame, fg_color="transparent")
action_row.pack(fill="x", padx=(54, 0), pady=(0, 4))

wait_btn = ctk.CTkButton(action_row, text="Left in Waiting Area", ...)
wait_btn.pack(side="left", padx=(0, 6))

if self._musipos_client is not None:
    po_btn = ctk.CTkButton(action_row, text="On PO", ...)
    po_btn.pack(side="left")

# status label for this row (shows "✓ Done" after action)
row_status = ctk.CTkLabel(action_row, text="", ...)
row_status.pack(side="left", padx=(8, 0))
```

Button sizing: `width=140, height=26, font=size(11)` for "Left in Waiting Area"; `width=70, height=26` for "On PO".

Button colour scheme: gray by default; green hover; disabled+green after success.
### `_on_waiting_area_clicked(line_item, btn, status_lbl)`

1. Disable button to prevent double-click.
2. Compose note text: `f"{line_item.sku} left in waiting area"`
3. Run in background thread:
   - **Neto**: `self._neto_client.add_sticky_note(self._order_id, title="Item Status", description=note_text, dry_run=self._dry_run)`
   - **eBay**: append to existing note: `new_text = ((line_item.notes + "
") if line_item.notes else "") + note_text`. Truncate to 255. Call `set_private_notes(item_id, transaction_id, new_text, dry_run)`.
4. On success: `status_lbl.configure(text="✓ Note added", text_color="green")`, keep button disabled.
5. On error: re-enable button, `status_lbl.configure(text=error, text_color="red")`.
6. Also call `self._rebuild_notes_content()` (Neto) or refresh eBay note widget (eBay) after success.

### `_on_add_to_po_clicked(line_item, btn, status_lbl)`

1. Disable button.
2. Open `MusiposPODialog(...)` with:
   - `on_success=lambda result: self._on_po_added(line_item, result, btn, status_lbl)`
   - `on_note_only=lambda: self._on_po_note_only(line_item, btn, status_lbl)`
3. If dialog is cancelled entirely: re-enable button.

### `_on_po_added(line_item, po_result, btn, status_lbl)`

1. Compose note: `f"{line_item.sku} on PO"`
2. Add note (same API calls as waiting area — background thread)
3. `status_lbl.configure(text=f"✓ Added to PO #{po_result['po_no']}", text_color="green")`
4. Keep button disabled (greyed, text "On PO ✓")

### `_on_po_note_only(line_item, btn, status_lbl)`

User cancelled PO addition but still wants the note.
Same note flow as above but just adds note, no PO. Re-enables button after.

---
## 7. Wiring — pass `musipos_client` through the call chain

### `src/gui/app.py`

In `App.__init__` or wherever clients are created:
```python
from src.musipos_client import MusiposClient
_musipos_cfg = config.musipos
self._musipos_client = (
    MusiposClient(_musipos_cfg) if _musipos_cfg and _musipos_cfg.enabled else None
)
```

Pass to `ResultsTab`:
```python
self._results_tab = ResultsTab(..., musipos_client=self._musipos_client)
```

### `src/gui/results_tab.py`

Accept `musipos_client=None`, store as `self._musipos_client`, pass to `OrderDetailView`.

### `src/gui/daily_ops/results_view.py`

Same: accept and pass through.

### `src/gui/daily_ops/daily_ops_window.py`

Create `MusiposClient` and pass to `DailyOpsResultsView` (if results_view is instantiated here).

---
## Known Limitations / Verification Items

1. **sku_mappings.csv is a kit/bundle map, not a Neto→Musipos alias**: Columns: `neto_sku`, `is_kit`, `invoice_skus`, `qty_per_alias`, `supplier`. Delimiter: `|` (pipe) for kits. Used as Strategy 2 in the resolution chain.

2. **musipos_sku_map.csv is a new file**: Does not exist yet. Created automatically on first manual save. Stored at `musipos_map_path` from config.

3. **invoice_skus delimiter confirmed as `|` (pipe)**: For kit items (`is_kit = TRUE`), both `invoice_skus` and `qty_per_alias` are pipe-delimited (e.g. `SKU1|SKU2|SKU3` and `1|1|1`). For non-kit items both are single values. `load_kit_mappings()` must split on `|`.

4. **Electric Factory SKUs**: After stripping the `EF` suffix, `05MS201BPLUS` won't match `itm_iid = 05/MS201BPLUS`. The `itm_supplier_iid` may or may not contain the catalog number. Strategy 2 (kit mappings) may find the item via its invoice SKU. If not, falls to manual entry.

5. **D'Addario SKUs**: `J81244M` (Neto) vs `J812 4/4M` (Musipos itm_iid). D'Addario has no suffix so base_sku = the full Neto SKU. The cascade should find it via `itm_supplier_iid` if D'Addario's supplier SKU column stores `J81244M`. Needs testing.

6. **Musipos password**: Stored in `config.json` on the network share. Acceptable for current deployment model.

7. **pop_itm_title**: The `INSERT INTO sp4pop` should write `pop_itm_title = itm_title` (denormalized title field).

8. **Cost fields on new PO lines**: `pop_curr_cost = qpc_last_purchase_cost`, `pop_curr_price = itm_new_retail_price`. Both fetched in `resolve_item()` and passed to `add_to_po()` via the returned item dict.

---

## Implementation Order

1. `requirements.txt` — add pyodbc
2. `src/config.py` + `config.json` — add MusiposConfig
3. `src/musipos_client.py` — full implementation (test independently with a Python script)
4. `src/gui/musipos_po_dialog.py` — dialog implementation
5. `src/gui/order_detail_view.py` — add buttons + handlers
6. Wiring: `app.py`, `results_tab.py`, `daily_ops/results_view.py`, `daily_ops_window.py`
7. Testing with dry_run=True first

---

## Verification Checklist

- [ ] `pyodbc` installs without errors on this machine
- [ ] `MusiposClient.test_connection()` returns True
- [ ] Standard SKU (e.g. a D'Addario or AMS item) resolves correctly
- [ ] EF SKU: Strategy 1 fails, Strategy 2 (invoice_skus) attempted, if still not found → manual entry → saved to musipos_sku_map.csv
- [ ] "Left in Waiting Area" adds sticky note to Neto order
- [ ] "Left in Waiting Area" appends to eBay private note
- [ ] "On PO" dialog opens, shows correct item info
- [ ] Dry run: no DB writes, status shows "[DRY RUN]"
- [ ] Live run: PO updated in Musipos (verify in PowerBuilder), `qpc_qor` incremented
- [ ] "SKU on PO" note added to order after PO creation
- [ ] Both Daily Ops and Afternoon Ops detail views have the buttons