# Results / Fulfillment Section — UI Improvements Plan

## Scope

All changes are confined to two files:
- `src/gui/results_tab.py` — the list view (Treeview)
- `src/gui/order_detail_view.py` — the per-order fulfillment detail view

---

## Checklist

### 1. Treeview dark background fix (`results_tab.py`)
- [ ] In `OrderTreeview._apply_style()`, call `style.theme_use("clam")` before
      `style.configure(...)`. On Windows, the default `"vista"` theme ignores
      `fieldbackground`, leaving the Treeview interior white regardless of what
      the style says. Switching to the cross-platform `"clam"` theme lets our
      colour settings take effect.

---

### 2. Copy button feedback (`order_detail_view.py`)
- [ ] Modify `_copy_to_clipboard(text, btn=None, original_text="Copy")` to:
  - Copy text to clipboard (unchanged)
  - If `btn` is supplied, change its text to `"Copied!"`, wait 2 s, revert to
    `original_text` via `self.after(2000, ...)`.
- [ ] Update every `_copy_to_clipboard(...)` call in `_build_shipping()` to pass
      the button reference, e.g.:
  ```python
  btn = ctk.CTkButton(row, text="Copy", ...)
  btn.configure(command=lambda t=line, b=btn: self._copy_to_clipboard(t, b))
  btn.pack(...)
  ```
- [ ] Same for the "Copy All" button.

---

### 3. Copy button for the order number (`order_detail_view.py`)
- [ ] In `_build_header()`, add a small "Copy" button to the right of the
      order-ID label that copies `self._order_id` to the clipboard (with the
      same "Copied!" feedback as above).

---

### 4. Item image click-to-enlarge (`order_detail_view.py`)
- [ ] Current state: thumbnails (50×50) already load asynchronously via
      `_load_image_async()` and display in `img_label`.
- [ ] Store the original image bytes/object in `_load_image_async()` alongside
      the thumbnail, so we can resize it to a larger size on demand (e.g. up to
      400×400, maintaining aspect ratio).
- [ ] Bind `<Button-1>` on each `img_label` to `_open_image_large(url)`.
- [ ] Implement `_open_image_large(url)`:
  - Opens a plain `tk.Toplevel` (not `CTkToplevel` — avoids the canvas
    rendering bug) with title "Image Preview".
  - Downloads/re-uses the image, scales it to fit within 600×600, shows it in
    a `tk.Label`.
  - Clicking the toplevel or pressing `<Escape>` closes it.
  - If the image is still loading (no URL or download error), the click is a
    no-op (or shows a brief "Loading…" tooltip).

---

### 5. Yellow text for existing order notes (`order_detail_view.py`)
- [ ] In `_build_neto_notes()`:
  - Delivery instructions label → `text_color="#f5c518"` (gold/yellow)
  - Each sticky note label → `text_color="#f5c518"`
  - Internal notes label → `text_color="#f5c518"`
- [ ] In `_build_ebay_notes()`:
  - Buyer notes label → `text_color="#f5c518"`
  - Per-line private notes labels → `text_color="#f5c518"`
- Note: the section heading "Notes" and the "Add Sticky Note:" sub-label keep
  their default colours — only the *content* of existing notes turns yellow.

---

### 6. Text selectability (`order_detail_view.py`)
- [ ] **Shipping address lines**: Replace each `ctk.CTkLabel` (not selectable)
      with a `tk.Entry` in `state="readonly"` mode:
  ```python
  e = tk.Entry(row, font=("", 13), readonlybackground="transparent",
               relief="flat", bd=0, fg="white", bg="#2b2b2b", state="readonly")
  e.insert(0, line)
  e.pack(side="left", fill="x", expand=True)
  ```
  (Use a helper `_selectable_label(parent, text)` to avoid repetition.)
  > The exact background/fg colours should match the CTk dark theme; consider
  > reading `ctk.ThemeManager` values or hard-coding both light/dark variants.
- [ ] **Notes text blocks**: Replace multi-line `CTkLabel` with
      `ctk.CTkTextbox(state="disabled")` — CTkTextbox supports text selection
      and cursor display even when read-only.
- [ ] **Order ID and customer name in header**: Wrap in `tk.Entry(state="readonly")`
      or keep as-is if size is too small to warrant it (header labels are short).
- [ ] Add `import tkinter as tk` at the top of `order_detail_view.py` (not yet
      present).

---

### 7. Right-click context menu on text entry fields (`order_detail_view.py`)
- [ ] Add a helper method `_bind_context_menu(widget)` that binds `<Button-3>`
      to show a `tk.Menu` with Cut / Copy / Paste commands using the widget's
      own clipboard methods:
  ```python
  def _bind_context_menu(self, widget):
      menu = tk.Menu(widget, tearoff=0)
      menu.add_command(label="Cut",   command=lambda: widget.event_generate("<<Cut>>"))
      menu.add_command(label="Copy",  command=lambda: widget.event_generate("<<Copy>>"))
      menu.add_command(label="Paste", command=lambda: widget.event_generate("<<Paste>>"))
      widget.bind("<Button-3>", lambda e: menu.tk_popup(e.x_root, e.y_root))
  ```
- [ ] Call `_bind_context_menu(self._tracking_entry)` after creating the
      tracking entry.
- [ ] If the "Add Sticky Note" textbox (`self._note_textbox`) doesn't already
      have native right-click support, bind the context menu to it as well.

---

### 8. Shipping method dropdown (`order_detail_view.py`)
- [ ] In `_build_tracking()`, replace `self._carrier_entry = ctk.CTkEntry(...)` with:
  ```python
  _SHIPPING_METHODS = [
      "Allied Express", "Aramex", "Australia Post",
      "Bonds Couriers", "Courier's Please", "DAI Post", "Toll",
  ]
  self._carrier_combo = ctk.CTkComboBox(
      row, values=_SHIPPING_METHODS, width=180, font=ctk.CTkFont(size=12)
  )
  self._carrier_combo.set("")          # blank default
  self._carrier_combo.pack(side="left")
  ```
- [ ] Update `_mark_as_sent()`: replace `self._carrier_entry.get()` with
      `self._carrier_combo.get()`.
- [ ] Update the `_mark_as_sent()` disable block after sending:
      `self._carrier_combo.configure(state="disabled")`.
- [ ] The constant `_SHIPPING_METHODS` can live at module level (top of
      `order_detail_view.py`) so it's easy to extend later.

---

## Execution Order

1. `results_tab.py` — Treeview theme fix (small, isolated)
2. `order_detail_view.py` — add `import tkinter as tk` at top
3. `order_detail_view.py` — add `_SHIPPING_METHODS` constant at module level
4. `order_detail_view.py` — add `_bind_context_menu()` helper
5. `order_detail_view.py` — modify `_copy_to_clipboard()` signature + logic
6. `order_detail_view.py` — `_build_header()`: add order-number Copy button
7. `order_detail_view.py` — `_build_shipping()`: Copy feedback + selectable address lines
8. `order_detail_view.py` — `_build_neto_notes()`: yellow text + selectable textboxes
9. `order_detail_view.py` — `_build_ebay_notes()`: yellow text
10. `order_detail_view.py` — `_build_line_items()`: image click binding
11. `order_detail_view.py` — `_load_image_async()`: cache full-res image bytes
12. `order_detail_view.py` — add `_open_image_large()` method
13. `order_detail_view.py` — `_build_tracking()`: ComboBox + context menu on tracking entry
14. `order_detail_view.py` — `_mark_as_sent()`: use ComboBox `.get()` + disable it on send

---

## Open Questions / Notes

- **Selectable address lines background colour**: `tk.Entry` needs explicit
  `bg`/`fg` values. Will hard-code `("#1a1a1a", "#2b2b2b")` for dark and
  `("#f5f5f5", "#1a1a1a")` for light and pick based on
  `ctk.get_appearance_mode()` at build time. If the user switches theme at
  runtime, the colour won't update — this is acceptable for now.
- **Image enlarger**: Using plain `tk.Toplevel` (not `ctk.CTkToplevel`) to avoid
  the canvas rendering bug confirmed in previous sessions.
- **CTkTextbox for notes**: `CTkTextbox` in `state="disabled"` allows selection
  but greys out in some themes. Consider `text_color_disabled` override if the
  yellow tint is lost when disabled.
- **`_carrier_entry` rename**: After switching to ComboBox, any existing
  `self._carrier_entry` reference in `_mark_as_sent()` must be updated to
  `self._carrier_combo`.
