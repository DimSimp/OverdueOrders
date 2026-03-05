# Order Fulfillment Workflow — Implementation Reference

This document covers all features added in the fulfillment workflow update. Use it for debugging and future development.

---

## Table of Contents

1. [Architecture Overview](#architecture-overview)
2. [Data Flow](#data-flow)
3. [Files Changed / Created](#files-changed--created)
4. [Phase 1: Data Model & API Extensions](#phase-1-data-model--api-extensions)
5. [Phase 2: Order Detail Modal](#phase-2-order-detail-modal)
6. [Phase 3: Results Tab Integration](#phase-3-results-tab-integration)
7. [Phase 4: Session Snapshots](#phase-4-session-snapshots)
8. [Phase 5: Polish](#phase-5-polish)
9. [Configuration Reference](#configuration-reference)
10. [Troubleshooting](#troubleshooting)

---

## Architecture Overview

```
main.py
  └─ App (src/gui/app.py)
       ├─ InvoiceTab (src/gui/invoice_tab.py)
       │    └─ "Load Session" button → src/session.py
       ├─ OrdersTab (src/gui/orders_tab.py)
       │    ├─ NetoClient (src/neto_client.py) → Neto REST API
       │    └─ EbayClient (src/ebay_client.py) → eBay Fulfillment + Trading APIs
       └─ ResultsTab (src/gui/results_tab.py)
            ├─ ReadOnlyTable (clickable rows, hover highlight)
            ├─ OrderDetailModal (src/gui/order_detail_modal.py) ← opens on row click
            ├─ "Refresh Orders" button → re-fetches from APIs
            ├─ "Save Session As" button → src/session.py
            └─ Auto-save snapshot on load_results()
```

**Threading model**: All API calls run in `threading.Thread(daemon=True)`. Results are passed back to the UI thread via `widget.after(0, callback)`. This prevents the GUI from freezing during network requests.

**Safety**: All API write operations respect the `dry_run` flag in `config.json`. When `dry_run: true` (the default), writes print to console instead of calling the real API.

---

## Data Flow

### Normal workflow
1. User imports invoice PDFs on the Invoice tab
2. User fetches orders on the Orders tab (Neto + eBay APIs)
3. App matches invoice SKUs against order line items (`src/data_processor.py`)
4. Results tab displays matched/unmatched orders in grouped tables
5. Session auto-saves to `snapshot_dir` or `output_dir`

### Clicking an order (new)
1. User clicks an order row in the Matched Orders table
2. Background thread calls `get_order_status()` on the relevant API
3. If the order is already dispatched/fulfilled → show "already completed" message → refresh
4. Otherwise → open `OrderDetailModal` with full order details
5. User can view address, add notes, enter tracking, mark as sent
6. On "Mark as Sent" → calls `update_order_status()` (Neto) or `create_shipping_fulfillment()` (eBay)
7. On modal close with `completed=True` → triggers full order refresh

### Session snapshot workflow
1. Auto-saves on every `load_results()` call
2. "Save Session As" lets user pick a custom directory
3. "Load Session" on Invoice tab restores all state and jumps to Results

---

## Files Changed / Created

| File | Status | Purpose |
|------|--------|---------|
| `src/neto_client.py` | **Modified** | Extended dataclasses, added write + status methods |
| `src/ebay_client.py` | **Modified** | Extended dataclasses, added write + status methods, scope change |
| `src/config.py` | **Modified** | Added `dry_run`, `snapshot_dir` to AppConfig |
| `config.json` | **Modified** | Added `dry_run`, `snapshot_dir` fields |
| `src/gui/order_detail_modal.py` | **Created** | Order detail popup window |
| `src/session.py` | **Created** | Session snapshot save/load |
| `src/gui/results_tab.py` | **Modified** | Clickable rows, hover, modal integration, refresh, save |
| `src/gui/invoice_tab.py` | **Modified** | "Load Session" button |
| `src/gui/app.py` | **Modified** | Dry-run banner |

---

## Phase 1: Data Model & API Extensions

### NetoOrder / NetoLineItem (`src/neto_client.py`)

**New fields on `NetoOrder`** (all default to empty string/list):
- `sticky_notes: list[dict]` — raw StickyNotes from API (each has `Title`, `Description`, `StickyNoteID`)
- `internal_notes: str` — InternalOrderNotes field
- `delivery_instruction: str` — DeliveryInstruction field
- `ship_first_name`, `ship_last_name`, `ship_company` — recipient info
- `ship_street1`, `ship_street2`, `ship_city`, `ship_state`, `ship_postcode`, `ship_country`, `ship_phone`

**New field on `NetoLineItem`**:
- `image_url: str` — from `OrderLine.ThumbURL` in the API response

**Backward compatibility**: The existing `notes` field (concatenated string) is still populated and used by the matching logic and results table display. The new separate fields are only used by the order detail modal.

**New API fields in `OUTPUT_SELECTOR`**:
- Shipping: `ShipFirstName`, `ShipLastName`, `ShipCompany`, `ShipStreetLine1`, `ShipStreetLine2`, `ShipCity`, `ShipState`, `ShipPostCode`, `ShipCountry`, `ShipPhone`
- Images: `OrderLine.ThumbURL`

**New methods on `NetoClient`**:

| Method | Purpose | Dry-run behavior |
|--------|---------|-----------------|
| `_post_action(action, body)` | Generalized POST with configurable `NETOAPI_ACTION` header. `_post()` now delegates to this. | N/A (internal) |
| `get_order_status(order_id) → str` | Lightweight GetOrder requesting only `OrderID` + `OrderStatus`. Used before opening the modal. | Always runs (read-only) |
| `update_order_status(order_id, new_status, tracking_number, carrier, dry_run)` | Sets order to "Dispatched" with optional tracking. Uses `UpdateOrder` action. | Prints payload to console, returns `{"Ack": "Success", "DryRun": True}` |
| `add_sticky_note(order_id, title, description, dry_run)` | Adds a sticky note via `UpdateOrder` action with `StickyNotes` array. | Prints payload to console |

### EbayOrder / EbayLineItem (`src/ebay_client.py`)

**New fields on `EbayOrder`**:
- `ship_name`, `ship_street1`, `ship_street2`, `ship_city`, `ship_state`, `ship_postcode`, `ship_country`, `ship_phone`
- Extracted from `fulfillmentStartInstructions[0].shippingStep.shipTo` (was already in the API response but not parsed)

**New field on `EbayLineItem`**:
- `image_url: str` — from `lineItems[].image.imageUrl`

**Scope change**: `EBAY_SCOPE` changed from `sell.fulfillment.readonly` to `sell.fulfillment`. This is required for the `create_shipping_fulfillment` endpoint. **The user will need to re-authenticate eBay once** (existing refresh token won't work with the new scope).

**New methods on `EbayClient`**:

| Method | Purpose | Dry-run behavior |
|--------|---------|-----------------|
| `get_order_status(order_id) → str` | GET single order, return `orderFulfillmentStatus`. Used before opening modal. | Always runs (read-only) |
| `create_shipping_fulfillment(order_id, line_items, tracking_number, carrier, dry_run)` | POST to `/sell/fulfillment/v1/order/{orderId}/shipping_fulfillment`. Sends all line items with their quantities. | Prints payload to console, returns `{"fulfillmentId": "DRY_RUN", "DryRun": True}` |

### Config (`src/config.py`)

**New fields on `AppConfig`**:
- `dry_run: bool = True` — controls whether API write methods actually call the API
- `snapshot_dir: str = ""` — default directory for session snapshots. Falls back to `output_dir` if empty.

---

## Phase 2: Order Detail Modal

### File: `src/gui/order_detail_modal.py`

**Class**: `OrderDetailModal(ctk.CTkToplevel)` — a modal popup window (750x780, resizable, min 600x500).

**Constructor parameters**:
```python
OrderDetailModal(
    master,              # parent widget
    order_id: str,
    platform: str,       # "eBay", "Website", etc.
    neto_order: NetoOrder | None,
    ebay_order: EbayOrder | None,
    matched_skus: list[str],   # SKUs that matched the invoice (shown with "*")
    neto_client: NetoClient | None,
    ebay_client: EbayClient | None,
    dry_run: bool = True,
    on_close_callback=None,    # called with (completed: bool) on close
)
```

**Layout (top to bottom inside a scrollable frame)**:

1. **Header** — order number (bold 18pt), platform badge (blue pill), customer name, date

2. **Shipping Address** (bordered frame) — each non-empty address line shown with a "Copy" button beside it. "Copy All" button at bottom copies the full address block to clipboard.
   - Neto: `{first} {last}`, company, street1, street2, `{city} {state} {postcode}`, country, phone
   - eBay: name, street1, street2, `{city} {state} {postcode}`, country, phone

3. **Line Items** (bordered frame) — header row + one row per line item:
   - **Image** (50x50px) — loaded asynchronously via `threading.Thread`. Uses `requests.get()` to fetch the URL, `PIL.Image` to resize, `ctk.CTkImage` to display. References stored in `self._image_refs` to prevent garbage collection. Failures silently leave a blank placeholder.
   - **SKU**, **Description** (with wraplength), **Qty**, **Arrived** (green `*` if SKU is in `matched_skus`)

4. **Notes** (bordered frame) — platform-specific:
   - **Neto**: delivery instructions (read-only), existing sticky notes (read-only, gray background), internal notes (read-only), new sticky note textbox + "Add Note" button
   - **eBay**: buyer checkout notes (read-only), per-item PrivateNotes (read-only, gray background), placeholder for future note editing

5. **Tracking** — two entry fields: tracking number (width 220) + carrier (width 160)

6. **Freight Placeholder** — disabled "Book Freight" button + "Coming soon" label

7. **Action Bar** (outside the scrollable frame, pinned to bottom):
   - **"Mark as Sent"** (green button) — prompts if no tracking number. Calls the appropriate API method. On success: disables all inputs, changes button to "COMPLETED", changes Close to "Back to Orders", shows status message (orange for dry-run, green for real).
   - **"Close"** / **"Back to Orders"** — calls `on_close_callback(completed)` then destroys window.

**Key implementation details**:
- `self.transient(master)` + `self.grab_set()` makes it a true modal (blocks interaction with parent)
- `self.protocol("WM_DELETE_WINDOW", self._close)` ensures callback fires on window X button
- Image loading uses `label.after(0, ...)` to update the UI from the background thread

---

## Phase 3: Results Tab Integration

### File: `src/gui/results_tab.py`

**Changes to `ReadOnlyTable`**:

1. **New class variable**: `_HOVER_COLORS` — slightly lighter/darker variants of `_GROUP_COLORS` for hover effect

2. **New parameter on `load_rows()`**: `on_row_click: callable | None`
   - When provided, each order group frame gets:
     - `<Enter>` / `<Leave>` bindings that swap `fg_color` between normal and hover colors
     - Each cell label gets `cursor="hand2"` and `<Button-1>` binding
   - The click callback receives `(group_key, platform)` — group_key is the order ID

3. **Platform detection**: The platform string is extracted from the first row of each group. For matched orders (which have a Remove button, so `col_offset=1`), platform is at column index 1. For unmatched orders (col_offset=0), it's at column index 0.

**New methods on `ResultsTab`**:

| Method | Purpose |
|--------|---------|
| `_open_order_detail(order_id, platform)` | Entry point when a row is clicked. Shows "Checking order status..." then spawns a background thread to call `get_order_status()`. |
| `_handle_status_check(order_id, platform, is_completed)` | Runs on UI thread. If completed → messagebox + refresh. Otherwise → open modal. |
| `_show_order_modal(order_id, platform)` | Finds the raw NetoOrder/EbayOrder, gathers matched SKUs, creates `OrderDetailModal`. |
| `_on_modal_close(completed)` | If the order was marked as sent, triggers `_refresh_orders()`. |
| `_refresh_orders()` | Disables refresh button, spawns background thread to re-fetch all orders from both APIs, then re-runs matching. |
| `_apply_refreshed_orders(neto, ebay)` | Updates app state, re-matches, clears manual overrides, re-renders tables. |
| `_save_session_as()` | Opens directory chooser, calls `save_snapshot()` to the selected directory. |

**Status check logic**:
- eBay: completed if `orderFulfillmentStatus == "FULFILLED"`
- Neto: completed if `OrderStatus` is `"dispatched"`, `"shipped"`, or `"completed"` (case-insensitive)
- If the status check fails (network error, etc.), the modal opens anyway with a warning in the error label

**New buttons in bottom bar**:
- "Refresh Orders" (blue) — re-fetches from APIs and re-matches
- "Save Session As" (gray) — opens directory chooser for snapshot save

**Auto-save**: At the end of `load_results()`, after `_refresh_tables()`, a snapshot is saved silently. Wrapped in try/except — failures are ignored (non-critical).

---

## Phase 4: Session Snapshots

### File: `src/session.py`

**Snapshot format** (JSON):
```json
{
  "version": 1,
  "timestamp": "2026-03-05T14:30:00.123456",
  "invoice_items": [ { "sku": "...", "sku_with_suffix": "...", "description": "...", "quantity": 1, "source_page": 0, "qty_flagged": false } ],
  "neto_orders": [ { ...all NetoOrder fields... } ],
  "ebay_orders": [ { ...all EbayOrder fields... } ],
  "matched_orders": [ { ...all MatchedOrder fields... } ],
  "unmatched_inv": [ { ...InvoiceItem fields... } ],
  "excluded_order_ids": [ ["platform", "order_id"], ... ],
  "force_matched_order_ids": [ ["platform", "order_id"], ... ]
}
```

**Filename**: `session_YYYYMMDD_HHMMSS.json`

**Serialization**:
- Uses `dataclasses.asdict()` for most fields
- `datetime` fields are explicitly converted to ISO 8601 strings (`.isoformat()`)
- `default=str` in `json.dump` as a safety net for any remaining non-serializable types

**Deserialization**:
- Each dataclass type has a dedicated `_parse_*` function that reconstructs the object from a dict
- Dates are parsed with `datetime.fromisoformat()`
- Missing fields default to empty strings/zeros/empty lists (forward-compatible)

**Key functions**:

| Function | Purpose |
|----------|---------|
| `save_snapshot(save_dir, ...) → filepath` | Creates directory if needed, serializes all state, writes JSON file |
| `load_snapshot(path) → SessionSnapshot` | Reads JSON, reconstructs all dataclass instances |

### Integration points:

**Auto-save** (`results_tab.py` → `load_results()`):
```python
# After _refresh_tables() succeeds
save_dir = config.app.snapshot_dir or config.app.output_dir
save_snapshot(save_dir, ...)  # wrapped in try/except, non-critical
```

**Save Session As** (`results_tab.py` → `_save_session_as()`):
- Opens `filedialog.askdirectory()` starting at `snapshot_dir` or `output_dir`
- Calls `save_snapshot()` with the chosen directory

**Load Session** (`invoice_tab.py` → `_load_session()`):
- Opens `filedialog.askopenfilename()` for `.json` files
- Calls `load_snapshot()`, then:
  1. Sets app state: `neto_orders`, `ebay_orders`, `matched_orders`
  2. Populates invoice table with snapshot's invoice items
  3. Sets results tab internal state (matched, unmatched, overrides)
  4. Calls `results._refresh_tables()` to re-render
  5. Jumps to "3. Results" tab

---

## Phase 5: Polish

### Dry-run banner (`src/gui/app.py`)

When `config.app.dry_run` is `True`, a red banner appears below the header bar:

> **DRY RUN MODE — API writes are simulated (change dry_run in config.json to disable)**

The banner is a `CTkFrame(height=28, fg_color=("red3", "red4"))` with `pack_propagate(False)` to enforce height.

---

## Configuration Reference

### `config.json` — `app` section

```json
{
  "app": {
    "order_lookback_days": 90,
    "on_po_filter_phrase": "on po",
    "output_dir": "output",
    "dry_run": true,
    "snapshot_dir": ""
  }
}
```

| Field | Type | Default | Description |
|-------|------|---------|-------------|
| `dry_run` | bool | `true` | When true, API write methods print to console instead of calling the real API. Set to `false` only when ready to process real orders. |
| `snapshot_dir` | string | `""` | Default save location for session snapshots. If empty, falls back to `output_dir`. Can be a network path like `\\server\share\snapshots`. |

---

## Troubleshooting

### "eBay API error (401)" after the update
The eBay OAuth scope changed from `sell.fulfillment.readonly` to `sell.fulfillment`. You need to re-authenticate:
1. Go to the Orders tab
2. Click "Authenticate eBay"
3. Complete the OAuth flow in the browser
4. Paste the redirect URL back into the app

### Order detail modal shows blank shipping address
- **Neto**: Check that the order has `ShipFirstName` etc. populated in Neto. Some old orders may not have shipping data.
- **eBay**: The shipping address comes from `fulfillmentStartInstructions[0].shippingStep.shipTo`. If the order has no shipping instructions, this will be empty.

### Images not loading in the order detail modal
- Images are loaded asynchronously from URLs. If the URL is empty or the request fails, the placeholder stays blank (this is intentional — no error is shown).
- **Neto**: Uses `OrderLine.ThumbURL`. If your Neto instance doesn't return this field, images will be blank. You could try `OrderLine.DefaultImageURL` instead (already checked as fallback in `_parse_order`).
- **eBay**: Uses `lineItems[].image.imageUrl` from the Fulfillment API response.

### "Mark as Sent" does nothing / shows dry run message
- Check `config.json` → `app.dry_run`. If it's `true`, all writes are simulated.
- The console (terminal where you launched the app) will show `[DRY RUN]` messages with the payload that would have been sent.

### Session snapshot fails to save
- Check that the `snapshot_dir` path exists and is writable. The app will try to create it with `os.makedirs(exist_ok=True)`.
- Auto-save failures are silently ignored (non-critical). Manual "Save Session As" failures show an error in the results tab.

### Clicking an order row opens nothing / shows "Status check failed"
- This means the API call to check order status failed (network error, auth expired, etc.).
- The modal will still open despite the error — it just couldn't verify whether the order was already completed.
- Check the error message in the bottom bar of the results tab.

### Hover highlight doesn't clear after clicking Remove/Add
- The `_refresh_tables()` call destroys and recreates all group frames, which should clear any lingering hover state.
- If you see a stale highlight, it may be a timing issue with `<Leave>` not firing before the frame is destroyed. This is cosmetic and resolves on the next interaction.

### Snapshot load shows wrong data
- Snapshots capture the state at the time of save, including manual overrides (excluded/force-matched orders).
- If you load a snapshot from yesterday, it reflects yesterday's orders. Use "Refresh Orders" after loading to get current API data while keeping the invoice items.
