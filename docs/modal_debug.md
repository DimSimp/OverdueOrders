# Order Detail Modal — Debugging Notes

## Status
**Option A applied (2026-03-07)** — Switched from `ctk.CTkToplevel` to plain `tk.Toplevel` to bypass all CTkToplevel init issues. Awaiting user test.

---

## The Problem
Clicking an order row in the Results tab opens `OrderDetailModal` (a `CTkToplevel` subclass), but the window appears **completely blank** — no buttons, labels, or input fields are visible. All `_build_*` methods complete without error (confirmed via logging).

---

## Architecture

- **`src/gui/order_detail_modal.py`** — `CTkToplevel` subclass
- **`src/gui/results_tab.py`** — opens modal via `OrderDetailModal(self, order_id=..., ...)`
- **`src/gui/invoice_tab.py`** — has a working "Test Modal" button (identical minimal CTkToplevel, used for comparison)

---

## What We Know

### ✅ Confirmed working
- All `_build_*` methods in `OrderDetailModal` complete successfully
- All order data (order_id, customer, SKUs) is correctly passed to the modal
- The **Invoice tab** test modal (`CTkToplevel(self)` + `withdraw()` + `deiconify()`) renders content perfectly

### ❌ What fails
- `CTkToplevel` created from the **Results tab** context produces a blank window
- The Results tab has a more complex widget hierarchy: inner `CTkTabview` + multiple `CTkScrollableFrame` (ReadOnlyTable) instances
- When using `withdraw()` before adding widgets: `winfo_ismapped=0` (window never shows)
- When using `CTkToplevel(self._app)` as master (root CTk window): window shows but content invisible
- Without `withdraw()`: window shows, but content still invisible

---

## Debugging Timeline

### Stage 1 — Confirmed data is correct
Added `[MODAL]` print logging throughout. Result: all data is present, all methods complete. Problem is purely visual.

### Stage 2 — Pack order bug (partially fixed)
The action bar was packed AFTER the `CTkScrollableFrame` container. Since `CTkScrollableFrame` uses `fill="both", expand=True`, it consumed all remaining space and the action bar got zero height. **Fixed** by packing action bar first with `side="bottom"`.

### Stage 3 — `withdraw()` attempt
Added `self.withdraw()` before building UI, then `deiconify()` after — `winfo_ismapped` dropped from 1 to 0. Removed `withdraw()`.

### Stage 4 — CTkScrollableFrame hypothesis
Replaced `CTkScrollableFrame` with plain `ctk.CTkFrame` for the content container. Still blank. The canvas is not the root cause.

### Stage 5 — `deiconify()` + `update()` attempts
Added explicit `self.deiconify()` + `self.update()` — still `winfo_ismapped=0` when triggered from Results tab.

### Stage 6 — Test modal comparison (key finding)
Added minimal test modal buttons to both Invoice and Results tabs. Results:
- Invoice tab test modal: `winfo_ismapped=1` ✅ renders correctly
- Results tab test modal: `winfo_ismapped=0` ✗ even without any grab_set()

Also found: iconbitmap `TclError` fires 200ms after creating any CTkToplevel from the Results tab. The old instance-level iconbitmap workaround was **broken** — the `after(200, self.iconbitmap, path)` call in `CTkToplevel.__init__` captures the bound method reference before we replaced it.

### Stage 7 — Fixed iconbitmap workaround
Replaced instance-level workaround with **class-level method override** in `OrderDetailModal`. This correctly intercepts `after(200, self.iconbitmap, ...)` since Python MRO resolves the method at call time.

### Stage 8 — Changed master to `self._app` (root CTk)
Changed `CTkToplevel(self)` → `CTkToplevel(self._app)`. Result: window now **appears** (was the original issue), but content is **invisible**. Progress!

Hypothesis: with root CTk as master, DPI/scaling is computed differently → child widget sizes may be 0.

### Stage 9 — Reverted master to `self`, removed `withdraw()`
Reverted to `CTkToplevel(self)` (tab frame as master), removed `win.withdraw()` from the Results tab test modal. Current state: window appears but content still invisible.

---

## Current Code State

**`order_detail_modal.py`**:
```python
class OrderDetailModal(ctk.CTkToplevel):
    def iconbitmap(self, *args, **kwargs):
        try:
            super().iconbitmap(*args, **kwargs)
        except Exception:
            pass

    def __init__(self, master, ...):
        super().__init__(master)  # master = ResultsTab frame (self)
        # ... setup ...
        self._build_ui()
        self.update_idletasks()
        # No withdraw(), no explicit deiconify()
        self.after(150, self._activate)

    def _build_action_bar(self):
        # Packed FIRST with side="bottom" — correct

    def _build_ui(self):
        self._build_action_bar()  # FIRST
        container = ctk.CTkFrame(self, fg_color="transparent")  # plain frame, not CTkScrollableFrame
        container.pack(fill="both", expand=True, ...)
        # ... _build_header, _build_shipping, etc.
```

**`results_tab.py`**:
```python
# Test modal (no withdraw()):
win = ctk.CTkToplevel(self)
win.title(...)
win.geometry("400x250")
win.transient(self.winfo_toplevel())
# add widgets...
win.update_idletasks()
win.deiconify()
win.lift()
win.focus_force()

# OrderDetailModal:
OrderDetailModal(self, order_id=order_id, ...)
```

---

## Hypotheses Not Yet Tested

1. **Widget hierarchy issue**: The Results tab's inner `CTkTabview` or `CTkScrollableFrame` widgets may interfere with how tkinter handles geometry/rendering for child `CTkToplevel` windows. The Invoice tab is structurally simpler (no inner CTkTabview).

2. **DPI scaling via `winfo_fpixels()`**: CTkScalingBaseClass calls `winfo_fpixels('1i')` during `__init__`. For a Toplevel whose master is a widget inside a `CTkTabview`, this might return wrong values, causing all child widgets to have 0 height.

3. **`CTkTabview` focus/visibility side effects**: When CTkTabview switches tabs, it may call configure/focus methods on its content frames. If any of these fire async events that interfere with the new Toplevel's initial rendering pass, content could be invisible.

4. **`ctk.CTkFrame(fg_color="transparent")` rendering**: The content container uses `fg_color="transparent"`. This may not render correctly if the parent Toplevel's background isn't properly initialised when master = complex frame hierarchy.

5. **`grab_set()` leftover**: If a previous (invisible) modal called `grab_set()`, it could prevent subsequent windows from being interacted with. However, test modals (which never use grab_set) also fail, so this is unlikely the primary cause.

---

## Stage 10 — Root cause identified via CTkToplevel source (2026-03-07)

Reading `ctk_toplevel.py` revealed the key mechanism:

`_windows_set_titlebar_color()` is called synchronously during `CTkToplevel.__init__`. It calls `super().withdraw()` + `super().update()` to apply the Windows dark titlebar, then schedules `after(5, _revert_withdraw_after_windows_set_titlebar_color)`. The revert callback calls `self.deiconify()` 5ms later.

This means:
1. At the time `update_idletasks()` is called in `__init__`, the window is **withdrawn** (`winfo_ismapped=0`)
2. The geometry manager runs, widgets are sized and positioned, but the window is not visible
3. At 5ms, the window is deiconified — now visible, but no WM_PAINT has been processed yet
4. `_activate()` fires at 150ms with `self.geometry(geo)` — this just reapplies the same size, no repaint

**The fix:** In `_activate()`, replaced the ineffective `self.geometry(geo)` nudge with `self.update()`. This processes ALL pending events including expose/WM_PAINT, forcing the OS to repaint all child widgets.

Also restored `CTkScrollableFrame` as the content container (it was changed to plain `CTkFrame` as a diagnostic), and removed all `[MODAL]` debug print statements.

---

## Next Steps to Try (if self.update() fix doesn't work)

1. **Print child widget sizes after deiconify**: Add `print(label.winfo_width(), label.winfo_height())` for children of the test modal to confirm if they are sized 0×0.

2. **Try `ctk.CTkFrame(fg_color=("gray90","gray10"))` instead of `fg_color="transparent"`** for the container — to confirm transparent rendering isn't the cause.

3. **Try `tk.Toplevel` directly** (not `CTkToplevel`) with manual customtkinter background color. If plain `tk.Toplevel` works, the issue is in CTkToplevel's subclass logic.

4. **Try creating the modal on a 200ms delay** using `self.after(200, lambda: OrderDetailModal(...))`. This would let the Results tab's event loop settle before the modal is created, potentially bypassing any async callbacks that interfere.

5. **Check CTkTabview tab visibility**: Verify that `self._inner_tabs.winfo_ismapped()` returns 1 when the Results tab is active. If any CTkTabview frame is not properly mapped, child Toplevels may inherit wrong state.

---

## Files Involved
- [src/gui/order_detail_modal.py](../src/gui/order_detail_modal.py)
- [src/gui/results_tab.py](../src/gui/results_tab.py)
- [src/gui/invoice_tab.py](../src/gui/invoice_tab.py)
