# FTP Inventory Integration — Plan & Progress Checklist

## Overview

Add an **FTP Inventory** mode to the Invoice tab (Tab 1) as an alternative to importing PDF supplier invoices.

### How It Works

Every morning, when our POS is launched, an inventory report is exported and uploaded to an FTP server (`ftp.drivehq.com`). This is the **Morning Inventory**. Throughout the day, new stock is received and entered into the system. In the afternoon, another inventory export is uploaded. This is the **Afternoon Inventory**.

By comparing the two files, we can identify every SKU whose quantity increased — meaning it was physically received today. This list of received SKUs becomes our “invoice items” for the matching workflow.

**The key insight:** the inventory files use Neto-native SKUs (the store’s own part numbers, already in the format Neto expects). No supplier suffix transformation is needed. Steps 2 (Orders) and 3 (Results) remain unchanged.

---

## What Changes

### New file: `src/ftp_inventory.py`
Handles FTP connection, file download, and inventory comparison. Isolated from the GUI.

### Modified: `src/config.py`
Add `FTPConfig` dataclass and parse it from `config.json["ftp"]`.

### Modified: `config.json`
Add `"ftp"` section with credentials and filenames.

### Modified: `src/gui/invoice_tab.py`
Add a mode switcher (`CTkSegmentedButton`: "PDF Invoice" / "FTP Inventory") at the top of the tab. The PDF controls (supplier dropdown, Import PDF, Scan) are only visible in PDF mode. In FTP mode, a single "Load from FTP" button is shown. The editable table and Next button work the same in both modes.

Also remove the leftover "Test Modal" debug button.

---

## Excel Column Layout (0-indexed)

| Column | Content |
|--------|----------|
| 0 | SKU |
| 8 | Quantity on hand |
| 18 | Supplier name |

Items are included only if `afternoon_qty − morning_qty > 0`.

---

## FTP Credentials (move to config.json)

| Field | Value |
|-------|-------|
| Host | `ftp.drivehq.com` |
| Username | `kyaldabomb` |
| Password | `D5es4stu!` |
| Morning file | `Morning_Inventory_Report.xlsx` |
| Afternoon file | `Afternoon_Inventory_Report.xlsx` |

---

## Data Flow

```
FTP server
  └─ Download Morning_Inventory_Report.xlsx
  └─ Download Afternoon_Inventory_Report.xlsx
      │
      ▼
Compare (pandas): afternoon_qty - morning_qty > 0
      │
      ▼
ReceivedItem list: [(sku, qty_delta, supplier_name), ...]
      │
      ▼
Convert to InvoiceItem objects:
  sku             = raw SKU (already Neto-native, no suffix needed)
  sku_with_suffix = sku  (same — no transformation)
  description     = ""   (not available from inventory)
  quantity        = int(qty_delta)
  source_page     = 0
  supplier_name   = from column 18
      │
      ▼
Load into EditableTable (user can still edit before proceeding)
      │
      ▼
Next button → SKIP SKU validation (Neto-native SKUs don’t need it)
      │
      ▼
Orders Tab → Results Tab  (unchanged)
```

---

## Detailed Implementation Steps

### Step 1 — `src/ftp_inventory.py` (new file)

```python
from __future__ import annotations

import ftplib
import tempfile
import os
from dataclasses import dataclass
from typing import List, Tuple

import pandas as pd


@dataclass
class FTPConfig:
    host: str
    username: str
    password: str
    morning_filename: str = "Morning_Inventory_Report.xlsx"
    afternoon_filename: str = "Afternoon_Inventory_Report.xlsx"


@dataclass
class ReceivedItem:
    sku: str
    quantity: float
    supplier: str


def download_and_compare(cfg: FTPConfig) -> list[ReceivedItem]:
    """
    1. Connect to FTP and download both inventory reports to temp files.
    2. Compare them: SKUs with increased quantity = received today.
    3. Return a list of ReceivedItem.
    Raises ftplib.all_errors or OSError on failure.
    """
    with tempfile.TemporaryDirectory() as tmpdir:
        morning_path = os.path.join(tmpdir, cfg.morning_filename)
        afternoon_path = os.path.join(tmpdir, cfg.afternoon_filename)

        ftp = ftplib.FTP(cfg.host)
        ftp.login(cfg.username, cfg.password)
        with open(morning_path, "wb") as f:
            ftp.retrbinary(f"RETR {cfg.morning_filename}", f.write)
        with open(afternoon_path, "wb") as f:
            ftp.retrbinary(f"RETR {cfg.afternoon_filename}", f.write)
        ftp.quit()

        return _compare_reports(morning_path, afternoon_path)


def _compare_reports(morning_path: str, afternoon_path: str) -> list[ReceivedItem]:
    df1 = pd.read_excel(morning_path, header=None)
    df2 = pd.read_excel(afternoon_path, header=None)

    morning_qty = df1.set_index(0)[8].to_dict()
    afternoon_qty = df2.set_index(0)[8].to_dict()
    supplier_map = df2.set_index(0)[18].to_dict()

    results = []
    all_skus = set(morning_qty) | set(afternoon_qty)
    for sku in all_skus:
        try:
            delta = float(afternoon_qty.get(sku, 0)) - float(morning_qty.get(sku, 0))
        except (ValueError, TypeError):
            continue
        if delta > 0:
            results.append(ReceivedItem(
                sku=str(sku).strip(),
                quantity=delta,
                supplier=str(supplier_map.get(sku, "")).strip(),
            ))
    return results
```

### Step 2 — `src/config.py`

Add `FTPConfig` dataclass:
```python
@dataclass
class FTPConfig:
    host: str
    username: str
    password: str
    morning_filename: str = "Morning_Inventory_Report.xlsx"
    afternoon_filename: str = "Afternoon_Inventory_Report.xlsx"
```

Add to `ConfigManager.__init__`:
```python
self.ftp: Optional[FTPConfig] = None
```

Add to `ConfigManager._parse()`:
```python
ftp_raw = self._raw.get("ftp", {})
if ftp_raw.get("host"):
    self.ftp = FTPConfig(
        host=ftp_raw["host"],
        username=ftp_raw.get("username", ""),
        password=ftp_raw.get("password", ""),
        morning_filename=ftp_raw.get("morning_filename", "Morning_Inventory_Report.xlsx"),
        afternoon_filename=ftp_raw.get("afternoon_filename", "Afternoon_Inventory_Report.xlsx"),
    )
```

### Step 3 — `config.json`

Add:
```json
"ftp": {
  "host": "ftp.drivehq.com",
  "username": "kyaldabomb",
  "password": "D5es4stu!",
  "morning_filename": "Morning_Inventory_Report.xlsx",
  "afternoon_filename": "Afternoon_Inventory_Report.xlsx"
}
```

### Step 4 — `src/gui/invoice_tab.py`

#### 4a — Add mode switcher at top of `_build_ui()`
```python
self._mode_var = ctk.StringVar(value="PDF Invoice")
self._mode_switcher = ctk.CTkSegmentedButton(
    self,
    values=["PDF Invoice", "FTP Inventory"],
    variable=self._mode_var,
    command=self._on_mode_change,
)
self._mode_switcher.pack(fill="x", padx=12, pady=(12, 4))
```

#### 4b — Wrap existing controls row in `self._pdf_frame`
All current controls (supplier dropdown, Import PDF, Scan, Clear, Load Session) move into a `ctk.CTkFrame` that is shown/hidden based on mode.

#### 4c — Add `self._ftp_frame` (initially hidden)
```python
self._ftp_frame = ctk.CTkFrame(self, fg_color="transparent")
# Contains:
#   - CTkLabel: "Downloads morning and afternoon inventory from FTP, finds received items."
#   - CTkButton: "📥 Load from FTP"  → self._load_from_ftp()
#   - self._ftp_status_label (separate from _status_label)
```

#### 4d — `_on_mode_change(mode: str)`
```python
def _on_mode_change(self, mode: str):
    if mode == "PDF Invoice":
        self._ftp_frame.pack_forget()
        self._pdf_frame.pack(fill="x", padx=12, pady=(0, 6))
    else:
        self._pdf_frame.pack_forget()
        self._ftp_frame.pack(fill="x", padx=12, pady=(0, 6))
    self._clear()  # Reset table when switching modes
```

#### 4e — `_load_from_ftp()`
```python
def _load_from_ftp(self):
    ftp_cfg = self._app.config.ftp
    if ftp_cfg is None:
        self._set_error("FTP not configured in config.json.")
        return
    self._ftp_btn.configure(state="disabled")
    self._set_status("Connecting to FTP…", color="gray60")

    def _worker():
        from src.ftp_inventory import download_and_compare
        try:
            received = download_and_compare(ftp_cfg)
            self.after(0, lambda: self._on_ftp_success(received))
        except Exception as exc:
            self.after(0, lambda: self._on_ftp_error(str(exc)))

    threading.Thread(target=_worker, daemon=True).start()
```

#### 4f — `_on_ftp_success(received: list[ReceivedItem])`
```python
def _on_ftp_success(self, received):
    from src.pdf_parser import InvoiceItem
    items = [
        InvoiceItem(
            sku=r.sku,
            sku_with_suffix=r.sku,   # already Neto-native
            description="",
            quantity=max(1, int(r.quantity)),
            source_page=0,
            supplier_name=r.supplier,
        )
        for r in received
        if r.sku
    ]
    self._invoice_items = items
    self._table.load_items(items, append=False)
    self._loaded_filenames = ["FTP Inventory (Morning vs Afternoon)"]
    self._update_files_box()
    count = self._table.row_count()
    self._set_status(f"{count} received item{'s' if count != 1 else ''} found.", color="green")
    self._next_btn.configure(state="normal" if count > 0 else "disabled")
    self._ftp_btn.configure(state="normal")
```

#### 4g — `_on_next_clicked()` — skip validation for FTP mode
```python
def _on_next_clicked(self):
    items = self.get_invoice_items()
    if not items:
        return
    # Skip SKU validation for FTP items (already Neto-native SKUs)
    if self._mode_var.get() == "FTP Inventory":
        self._on_complete()
        return
    # ... existing PDF validation logic unchanged ...
```

#### 4h — Remove "Test Modal" button
The `_open_test_modal()` method and its button were debug artefacts. Remove both.

---

## Notes & Constraints

- **No suffix logic for FTP items**: The inventory files contain Neto’s own SKUs. The supplier suffix system only applies to PDF invoices where the supplier uses their own part numbers. FTP items set `sku_with_suffix = sku`.
- **Description field is blank**: The inventory comparison does not provide product descriptions. The editable table will show an empty Description column. Users can type descriptions manually if needed, but it has no effect on matching.
- **Kit component expansion**: The original script expanded kit parent SKUs into their component child SKUs (via Neto `GetItem` + `KitComponents`). This is **not** included in this plan. It can be added as a future enhancement once the basic flow is confirmed working.
- **FTP timing**: This mode only makes sense after the afternoon inventory upload has occurred. If run in the morning before the afternoon upload exists on the server, the comparison will show no changes. This is a workflow constraint, not a bug — consider adding a note in the UI.
- **`openpyxl` dependency**: `pandas.read_excel()` with `.xlsx` files requires `openpyxl`. Check `requirements.txt` — it may already be present.

---

## Checklist

### Setup — **Do this before any coding**
- [ ] **Verify Excel column layout**: Download a sample inventory file from the FTP server and open it to confirm that column 0 = SKU, column 8 = quantity, and column 18 = supplier. These indices come from the original script and must be validated before we hardcode them. If they differ, update the plan accordingly.
- [ ] Confirm `openpyxl` is in `requirements.txt` (needed for `pd.read_excel`)

### Implementation
- [ ] **Step 1**: Create `src/ftp_inventory.py` with `ReceivedItem`, `download_and_compare()`, `_compare_reports()`
- [ ] **Step 2**: Add `FTPConfig` to `src/config.py` and parse it in `ConfigManager._parse()`
- [ ] **Step 3**: Add `"ftp"` section to `config.json`
- [ ] **Step 4a**: Add `CTkSegmentedButton` mode switcher to `InvoiceTab._build_ui()`
- [ ] **Step 4b**: Wrap existing PDF controls in `self._pdf_frame`
- [ ] **Step 4c**: Add `self._ftp_frame` with Load button and status label
- [ ] **Step 4d**: Implement `_on_mode_change()`
- [ ] **Step 4e**: Implement `_load_from_ftp()` background thread
- [ ] **Step 4f**: Implement `_on_ftp_success()` and `_on_ftp_error()`
- [ ] **Step 4g**: Update `_on_next_clicked()` to skip validation in FTP mode
- [ ] **Step 4h**: Remove "Test Modal" button and `_open_test_modal()` method

### Verification
- [ ] Launch app in PDF mode — existing import flow works unchanged
- [ ] Switch to FTP mode — PDF controls disappear, FTP button appears
- [ ] Click "Load from FTP" — downloads files, populates table with received items
- [ ] Quantity and supplier columns populated correctly
- [ ] Next button proceeds to Orders tab without SKU validation dialog
- [ ] Orders and Results tabs work identically to PDF workflow
- [ ] Switching back to PDF mode clears the table and re-enables PDF controls
- [ ] If FTP is not configured in config.json, a clear error message appears

### Future Enhancements (out of scope)
- [ ] Kit component expansion (parent SKU → child SKUs via Neto GetItem)
- [ ] Description lookup from Neto after FTP load
- [ ] "Afternoon upload not yet available" detection with user-friendly warning
