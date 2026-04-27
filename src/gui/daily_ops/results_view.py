from __future__ import annotations

import logging
import os
import subprocess
import threading
import tkinter as tk
from datetime import datetime
from pathlib import Path
from tkinter import messagebox
import tkinter.ttk as ttk

import customtkinter as ctk

from src.gui.results_tab import OrderTreeview, _shipping_display

log = logging.getLogger(__name__)

# ── Column spec ────────────────────────────────────────────────────────────────
# Reuses the same column IDs as the afternoon ops (so OrderTreeview renders
# correctly), but with daily-ops-appropriate headings.
#   #0          → Order No.
#   platform    → Platform
#   customer    → Customer
#   date        → State (ship state, not order date)
#   shipping    → Shipping
#   sku / description / qty  → line-item columns (child rows)
#   notes       → Type / Zone  (shown in the parent/order row)

_DAILY_COL_SPEC = {
    "#0":          ("Order No.",   120),
    "platform":    ("Platform",     80),
    "customer":    ("Customer",    170),
    "date":        ("State",        60),
    "shipping":    ("Shipping",     90),
    "assigned":    ("Assigned",     80),
    "sku":         ("SKU",         130),
    "description": ("Description", 180),
    "qty":         ("Qty",          40),
    "notes":       ("Type / Zone", 140),
    "order_notes": ("Notes",       150),
}

# Fixed picking-list output directory and filename (overwritten on every export)
from src.config import SERVER_AVAILABLE, LOCAL_DATA_DIR  # noqa: E402
_SERVER_PICKLIST_DIR = r"\\SERVER\Project Folder\Order-Fulfillment-App\Picking Lists\Daily"
_PICKLIST_DIR = (
    _SERVER_PICKLIST_DIR if SERVER_AVAILABLE
    else str(LOCAL_DATA_DIR / "picking_lists" / "daily")
)
_PICKLIST_FILENAME = "DAILY PICKING LIST.xlsx"


class DailyOpsResultsView(ctk.CTkFrame):
    """
    Step 6 — Results & Dispatch.

    Replicates the afternoon operations results screen with daily-ops specifics:
    - Two-tab layout: Active Orders | Removed
    - Right-click to move orders between tabs (two-way)
    - Refresh re-fetches all order IDs and merges shared overrides
    - Export Picking List → XLSX with zone section headers + page breaks
    - Print Pick Labels → Brother QL labels for every 'Picks' zone line item
    - "Envelopes only" filter checkbox
    - Auto-saves session to fixed network path on arrival and after every move
    - Cancel / Reprint dialog (same logic as afternoon ops)
    """

    def __init__(self, master, window, on_back, **kwargs):
        super().__init__(master, fg_color="transparent", **kwargs)
        self._window = window
        self._on_back = on_back
        self._assign_filter = "All"

        # Local copies of the order lists (refreshed independently of window)
        self._neto_orders: list = list(window.neto_orders)
        self._ebay_orders: list = list(window.ebay_orders)

        # (platform, order_id) tuples for removed orders
        self._removed_order_ids: set[tuple] = set()

        # Detail / freight overlays
        self._detail_frame = None
        self._freight_frame = None
        self._last_clicked_order_id: str | None = None
        self._last_clicked_platform: str | None = None

        self._loaded = False

        # Collation state
        self._ungrouped_order_ids: set[str] = set()
        self._collated_groups: dict = {}   # synthetic_id → CollatedGroup
        self._collated_frame = None
        self._book_freight_for_all: bool = False

        self.grid_rowconfigure(0, weight=1)
        self.grid_columnconfigure(0, weight=1)

        self._build_ui()

    # ── UI construction ────────────────────────────────────────────────────────

    def _build_ui(self):
        # List frame — raised normally; sits at (0,0) so detail/freight overlay it
        self._list_frame = ctk.CTkFrame(self, fg_color="transparent")
        self._list_frame.grid(row=0, column=0, sticky="nsew")
        self._list_frame.grid_rowconfigure(1, weight=1)
        self._list_frame.grid_rowconfigure(2, weight=0)
        self._list_frame.grid_columnconfigure(0, weight=1)

        self._build_list_page(self._list_frame)

    def _build_list_page(self, parent):
        # ── Toolbar ─────────────────────────────────────────────────────────
        toolbar = ctk.CTkFrame(parent, fg_color="transparent")
        toolbar.grid(row=0, column=0, sticky="ew", padx=12, pady=(10, 4))

        self._refresh_btn = ctk.CTkButton(
            toolbar, text="Refresh", width=90,
            fg_color=("dodgerblue3", "dodgerblue4"),
            command=self._refresh_all_orders,
        )
        self._refresh_btn.pack(side="left", padx=(0, 6))

        self._cancel_btn = ctk.CTkButton(
            toolbar, text="Cancel / Reprint", width=140,
            fg_color=("firebrick3", "firebrick4"),
            hover_color=("firebrick4", "firebrick"),
            command=self._open_cancel_dialog,
        )
        self._cancel_btn.pack(side="left", padx=(0, 6))

        self._export_btn = ctk.CTkButton(
            toolbar, text="Export Picking List", width=150,
            command=self._export_picking_list,
        )
        self._export_btn.pack(side="left", padx=(0, 6))

        self._labels_btn = ctk.CTkButton(
            toolbar, text="Print Pick Labels (0)", width=180,
            fg_color=("#2E7D32", "#1B5E20"),
            hover_color=("#256528", "#164A18"),
            state="disabled",
            command=self._print_pick_labels,
        )
        self._labels_btn.pack(side="left", padx=(0, 10))

        self._status_label = ctk.CTkLabel(
            toolbar, text="", font=ctk.CTkFont(size=12),
            text_color=("gray50", "gray60"),
        )
        self._status_label.pack(side="right")

        self._error_label = ctk.CTkLabel(
            toolbar, text="", font=ctk.CTkFont(size=12), text_color="red",
        )
        self._error_label.pack(side="right", padx=(0, 8))

        self._assign_seg = ctk.CTkSegmentedButton(
            toolbar, values=["All", "Mine", "Unassigned"],
            command=self._on_assign_filter_change,
            height=28, font=ctk.CTkFont(size=12),
        )
        self._assign_seg.set("All")
        self._assign_seg.pack(side="right", padx=(0, 12))

        ctk.CTkLabel(
            toolbar, text="Assign:", font=ctk.CTkFont(size=12),
            text_color=("gray50", "gray60"),
        ).pack(side="right", padx=(0, 4))

        # ── Tab view ────────────────────────────────────────────────────────
        def _on_tab_change():
            self._active_tree.reload_column_preferences()
            self._removed_tree.reload_column_preferences()
            self._on_selection_change()

        self._tabs = ctk.CTkTabview(parent, corner_radius=6, command=_on_tab_change)
        self._tabs.grid(row=1, column=0, sticky="nsew", padx=12, pady=(0, 4))
        self._tabs.add("Active Orders")
        self._tabs.add("Removed")

        # Active Orders tab
        _active_c = ctk.CTkFrame(self._tabs.tab("Active Orders"), fg_color="transparent")
        _active_c.pack(fill="both", expand=True)
        _active_c.grid_rowconfigure(0, weight=1)
        _active_c.grid_columnconfigure(0, weight=1)

        self._active_tree = OrderTreeview(
            _active_c,
            col_spec=_DAILY_COL_SPEC,
            table_id="daily_ops_results",
            user_manager=getattr(self._window, "user_manager", None),
            current_user=getattr(self._window, "current_user", None),
            on_row_click=self._open_detail_view,
            on_context_action=self._move_to_removed,
            context_label="Move to Removed",
            selectable=True,
            on_selection_change=self._on_selection_change,
            on_assign_request=self._handle_assign_request if self._is_admin() else None,
        )
        self._active_tree.grid(row=0, column=0, sticky="nsew")

        # Removed tab
        _removed_c = ctk.CTkFrame(self._tabs.tab("Removed"), fg_color="transparent")
        _removed_c.pack(fill="both", expand=True)
        _removed_c.grid_rowconfigure(0, weight=1)
        _removed_c.grid_columnconfigure(0, weight=1)

        self._removed_tree = OrderTreeview(
            _removed_c,
            col_spec=_DAILY_COL_SPEC,
            table_id="daily_ops_results",
            user_manager=getattr(self._window, "user_manager", None),
            current_user=getattr(self._window, "current_user", None),
            on_row_click=self._open_detail_view,
            on_context_action=self._move_to_active,
            context_label="Move back to Active",
            selectable=True,
            on_selection_change=self._on_selection_change,
            on_assign_request=self._handle_assign_request if self._is_admin() else None,
        )
        self._removed_tree.grid(row=0, column=0, sticky="nsew")

        # ── Action bar (row 2, hidden by default) ────────────────────────
        self._build_action_bar(parent)


    def _build_action_bar(self, parent):
        """Build the bulk-action bar (row 2). Hidden until ≥1 order is checked."""
        self._action_bar = ctk.CTkFrame(
            parent, fg_color=("gray85", "gray20"), corner_radius=8,
        )
        self._action_bar.grid(row=2, column=0, sticky="ew", padx=12, pady=(0, 8))
        self._action_bar.grid_remove()

        self._action_clear_btn = ctk.CTkButton(
            self._action_bar, text="✕ Clear", width=70, height=28,
            fg_color="gray50", hover_color="gray40",
            command=self._clear_all_checks,
        )
        self._action_clear_btn.pack(side="left", padx=(8, 4), pady=6)

        self._action_count_lbl = ctk.CTkLabel(
            self._action_bar, text="", font=ctk.CTkFont(size=12),
        )
        self._action_count_lbl.pack(side="left", padx=(4, 12))

        self._action_move_btn = ctk.CTkButton(
            self._action_bar, text="Move", width=140, height=28,
            command=self._bulk_move,
        )
        self._action_move_btn.pack(side="left", padx=(0, 6))

        self._action_sent_btn = ctk.CTkButton(
            self._action_bar, text="Mark as Sent", width=120, height=28,
            fg_color=("#2E7D32", "#1B5E20"), hover_color=("#256528", "#164A18"),
            command=self._bulk_mark_as_sent,
        )
        self._action_sent_btn.pack(side="left", padx=(0, 6))

        self._action_po_btn = ctk.CTkButton(
            self._action_bar, text="Add to PO", width=100, height=28,
            command=self._bulk_add_to_po,
        )
        self._action_po_btn.pack(side="left", padx=(0, 6))

        self._action_assign_btn = ctk.CTkButton(
            self._action_bar, text="Assign to User", width=130, height=28,
            command=self._bulk_assign,
        )
        self._action_assign_btn.pack(side="left", padx=(0, 6))

    def _current_tree(self):
        """Return the OrderTreeview for the currently visible tab."""
        if self._tabs.get() == "Active Orders":
            return self._active_tree
        return self._removed_tree

    def _on_selection_change(self):
        """Update action bar visibility and labels based on current tab's selection."""
        tree = self._current_tree()
        checked = tree.get_checked_orders()
        if not checked:
            self._action_bar.grid_remove()
            return

        n = len(checked)
        self._action_count_lbl.configure(text=f"{n} order{'s' if n != 1 else ''} selected")
        self._action_bar.grid()

        tab = self._tabs.get()
        musipos = getattr(self._window, "musipos_client", None)
        if tab == "Active Orders":
            self._action_move_btn.configure(text="Move to Removed", state="normal")
            self._action_sent_btn.configure(state="normal")
            self._action_po_btn.configure(state="normal" if musipos else "disabled")
        else:  # Removed
            self._action_move_btn.configure(text="Move to Active", state="normal")
            self._action_sent_btn.configure(state="disabled")
            self._action_po_btn.configure(state="disabled")
        self._action_assign_btn.configure(state="normal" if self._is_admin() else "disabled")

    def _is_admin(self) -> bool:
        cu = getattr(self._window, "current_user", None)
        return bool(cu and cu.get("role") == "admin")

    def _clear_all_checks(self):
        self._current_tree().clear_checks()

    # ── Bulk actions ──────────────────────────────────────────────────────────

    def _find_order_for_bulk(self, order_id: str, platform: str):
        if platform.lower() == "ebay":
            for o in self._ebay_orders:
                if o.order_id == order_id:
                    return o
        else:
            for o in self._neto_orders:
                if o.order_id == order_id:
                    return o
        return None

    def _bulk_move(self):
        tree = self._current_tree()
        checked = tree.get_checked_orders()
        if not checked:
            return
        tab = self._tabs.get()
        for c in checked:
            if tab == "Active Orders":
                self._move_to_removed(c["order_id"], c["platform"])
            else:
                self._move_to_active(c["order_id"], c["platform"])
        tree.clear_checks()

    def _bulk_mark_as_sent(self):
        import threading
        from tkinter import messagebox
        from src.cross_dispatch import sync_ebay_to_neto, sync_neto_to_ebay
        tree = self._current_tree()
        checked = tree.get_checked_orders()
        if not checked:
            return
        n = len(checked)
        if not messagebox.askyesno(
            "Mark as Sent",
            f"Mark {n} order{'s' if n != 1 else ''} as sent without tracking numbers?\n\nThis cannot be undone.",
            parent=self,
        ):
            return

        dry_run = self._window.config.app.dry_run
        neto_client = self._window.neto_client
        ebay_client = self._window.ebay_client
        results: list = []
        failed_ebay_syncs: list = []  # (ebay_order_id, error_msg) for dialog

        def _work():
            for c in checked:
                order = self._find_order_for_bulk(c["order_id"], c["platform"])
                if order is None:
                    results.append(f"⚠ {c['order_id']}: not found")
                    continue
                try:
                    if c["platform"].lower() == "ebay":
                        ebay_client.create_shipping_fulfillment(
                            c["order_id"],
                            line_items=order.line_items,
                            tracking_number="",
                            carrier="",
                            dry_run=dry_run,
                        )
                    else:
                        skus = [li.sku for li in order.line_items]
                        neto_client.update_order_status(
                            c["order_id"],
                            new_status="Dispatched",
                            tracking_number="",
                            shipping_method="",
                            line_item_skus=skus,
                            dry_run=dry_run,
                        )
                    label = "[DRY RUN] " if dry_run else ""
                    results.append(f"✓ {label}{c['order_id']}")
                except Exception as exc:
                    results.append(f"✗ {c['order_id']}: {exc}")
                    continue

                # Cross-platform sync (best-effort, non-fatal)
                try:
                    if c["platform"].lower() == "ebay" and neto_client:
                        ok, msg = sync_ebay_to_neto(
                            c["order_id"], neto_client,
                            ebay_created_at=getattr(order, "creation_date", None),
                            dry_run=dry_run,
                        )
                        results.append(f"  {'✓' if ok else '⚠'} Neto sync: {msg}")
                    elif c["platform"].lower() != "ebay" and ebay_client:
                        ebay_id = getattr(order, "purchase_order_number", "")
                        if getattr(order, "sales_channel", "").lower() == "ebay" and ebay_id:
                            ok, msg = sync_neto_to_ebay(
                                order, ebay_client,
                                dry_run=dry_run,
                            )
                            results.append(f"  {'✓' if ok else '⚠'} eBay sync: {msg}")
                            if not ok:
                                failed_ebay_syncs.append((ebay_id, msg))
                except Exception as exc:
                    results.append(f"  ⚠ sync error: {exc}")

            self.after(0, lambda: _done())

        def _done():
            from src.gui.dialogs import show_ebay_sync_failed_dialog
            tree.clear_checks()
            messagebox.showinfo("Mark as Sent — Results", "\n".join(results), parent=self)
            if failed_ebay_syncs:
                show_ebay_sync_failed_dialog(self, failed_ebay_syncs)
            self._refresh_all_orders()

        threading.Thread(target=_work, daemon=True).start()

    def _bulk_add_to_po(self):
        from tkinter import messagebox
        musipos_client = getattr(self._window, "musipos_client", None)
        if musipos_client is None:
            messagebox.showinfo("Not configured", "Musipos is not configured.", parent=self)
            return
        tree = self._current_tree()
        checked = tree.get_checked_orders()
        if not checked:
            return

        from collections import deque
        items_queue: deque = deque()
        for c in checked:
            order = self._find_order_for_bulk(c["order_id"], c["platform"])
            if order is None:
                continue
            for li in (order.line_items or []):
                items_queue.append((c["order_id"], c["platform"], li))

        if not items_queue:
            messagebox.showinfo("Add to PO", "No line items found in selected orders.", parent=self)
            return

        self._bulk_po_next(items_queue)

    def _bulk_po_next(self, queue):
        from tkinter import messagebox
        if not queue:
            messagebox.showinfo("Add to PO", "All items processed.", parent=self.winfo_toplevel())
            self._refresh_all_orders()
            return

        from collections import deque
        from src.gui.musipos_po_dialog import MusiposPODialog

        order_id, platform, line_item = queue.popleft()
        musipos_client = getattr(self._window, "musipos_client", None)
        if musipos_client is None:
            return

        qty = getattr(line_item, "quantity", 1) or 1
        product_name = (
            getattr(line_item, "product_name", None)
            or getattr(line_item, "title", "")
            or ""
        )

        advanced = [False]

        def _advance():
            if not advanced[0]:
                advanced[0] = True
                self.after(0, lambda: self._bulk_po_next(queue))

        def _on_success(po_result):
            sku = po_result.get("resolved_sku") or line_item.sku
            self._bulk_po_write_note(order_id, platform, sku, line_item)

        def _on_note_only(resolved_sku=None):
            sku = resolved_sku or line_item.sku
            self._bulk_po_write_note(order_id, platform, sku, line_item)

        def _on_cancel():
            pass

        dialog = MusiposPODialog(
            self.winfo_toplevel(),
            neto_sku=line_item.sku,
            product_name=product_name,
            order_qty=qty,
            musipos_client=musipos_client,
            suppliers_config=self._window.config.suppliers,
            dry_run=self._window.config.app.dry_run,
            on_success=_on_success,
            on_note_only=_on_note_only,
            on_cancel=_on_cancel,
        )
        dialog.protocol("WM_DELETE_WINDOW", lambda: (dialog.destroy(), _advance()))
        dialog.bind("<Destroy>", lambda e: _advance() if e.widget is dialog else None)

    def _bulk_po_write_note(self, order_id: str, platform: str, sku: str, line_item=None):
        """Fire-and-forget: add '{sku} on PO' note to a Neto or eBay order."""
        import threading
        from datetime import date
        dry_run = self._window.config.app.dry_run
        note_text = f"{sku} on PO"

        if platform.lower() == "ebay":
            ebay_client = self._window.ebay_client
            if not ebay_client or line_item is None:
                return
            item_id = getattr(line_item, "legacy_item_id", "")
            txn_id = getattr(line_item, "legacy_transaction_id", "")
            if not item_id:
                return

            def _ebay_work():
                try:
                    ebay_client.set_private_notes(
                        item_id=item_id,
                        transaction_id=txn_id,
                        note_text=note_text[:255],
                        dry_run=dry_run,
                    )
                except Exception:
                    pass
            threading.Thread(target=_ebay_work, daemon=True).start()
            return

        neto_client = self._window.neto_client
        if not neto_client:
            return
        dated_note = f"[{date.today().strftime('%d/%m/%Y')}] {note_text}"

        def _neto_work():
            try:
                neto_client.add_sticky_note(order_id, title="Item Status", description=dated_note, dry_run=dry_run)
            except Exception:
                pass
        threading.Thread(target=_neto_work, daemon=True).start()

    # ── Assignment ─────────────────────────────────────────────────────────────

    def _on_assign_filter_change(self, value: str) -> None:
        self._assign_filter = value
        self._refresh_tables()

    def _handle_assign_request(self, order_ids: list, platform: str,
                               single: bool = False, remove: bool = False) -> None:
        """Handle right-click assign request (single order)."""
        if remove:
            am = getattr(self._window, "assignment_manager", None)
            if am:
                try:
                    am.unassign(order_ids)
                except Exception:
                    pass
            self._refresh_all_orders()
            return
        self._open_assign_popup(order_ids)

    def _bulk_assign(self) -> None:
        """Called by 'Assign to User' action bar button."""
        checked = self._current_tree().get_checked_orders()
        if not checked:
            return
        order_ids = [c["order_id"] for c in checked]
        self._open_assign_popup(order_ids)

    def _open_assign_popup(self, order_ids: list) -> None:
        """Show a small popup to pick a user (or unassign) for the given order IDs."""
        am = getattr(self._window, "assignment_manager", None)
        um = getattr(self._window, "user_manager", None)
        cu = getattr(self._window, "current_user", None)
        if not am or not um or not cu:
            return
        try:
            users = um.get_active_users()
        except Exception:
            return

        popup = ctk.CTkToplevel(self.winfo_toplevel())
        popup.title("Assign Orders")
        popup.geometry("320x180")
        popup.resizable(False, False)
        popup.transient(self.winfo_toplevel())
        popup.grab_set()

        body = ctk.CTkFrame(popup, fg_color="transparent")
        body.pack(fill="both", expand=True, padx=20, pady=16)

        n = len(order_ids)
        ctk.CTkLabel(
            body,
            text=f"Assign {n} order{'s' if n != 1 else ''} to:",
            font=ctk.CTkFont(size=13),
        ).pack(anchor="w", pady=(0, 10))

        user_options = ["— Unassign —"] + [
            f"{u['first_name']} {u['last_name']}".strip() or u["username"]
            for u in users
        ]
        user_map = {"— Unassign —": ""} | {
            (f"{u['first_name']} {u['last_name']}".strip() or u["username"]): u["username"]
            for u in users
        }
        sel_var = ctk.StringVar(value=user_options[0])
        ctk.CTkComboBox(body, variable=sel_var, values=user_options,
                        state="readonly", height=34).pack(fill="x", pady=(0, 14))

        btn_row = ctk.CTkFrame(body, fg_color="transparent")
        btn_row.pack(fill="x")

        def _confirm():
            target_username = user_map.get(sel_var.get(), "")
            try:
                if target_username:
                    am.assign(order_ids, target_username, cu["username"])
                else:
                    am.unassign(order_ids)
            except Exception:
                pass
            self._current_tree().clear_checks()
            self._refresh_all_orders()
            popup.destroy()

        ctk.CTkButton(btn_row, text="Assign", command=_confirm).pack(side="left", padx=(0, 8))
        ctk.CTkButton(btn_row, text="Cancel", fg_color="transparent",
                      border_width=1, command=popup.destroy).pack(side="left")
        popup.after(50, popup.lift)

    # ── Entry point ────────────────────────────────────────────────────────────

    def show(self, initial_removed_ids=None, initial_ungrouped_ids=None):
        """Called by DailyOpsWindow each time this step is shown. Loads only once.

        initial_removed_ids: optional set of (platform, order_id) tuples restored
        from a saved session file (passed when loading rather than generating).
        initial_ungrouped_ids: optional set of order_id strings for orders that
        have been explicitly ungrouped and should not be collated.
        """
        if self._loaded:
            return
        self._loaded = True
        # Snapshot orders from window (may have been restored from a session file)
        self._neto_orders = list(self._window.neto_orders)
        self._ebay_orders = list(self._window.ebay_orders)
        if initial_ungrouped_ids is not None:
            self._ungrouped_order_ids = set(initial_ungrouped_ids)
        if initial_removed_ids is not None:
            # Loading an existing session — restore removed IDs then merge any
            # cross-workstation overrides added since the session was last saved.
            self._removed_order_ids = set(initial_removed_ids)
            self._merge_overrides()
        else:
            # Fresh session — clear the overrides file so this and other
            # workstations start with an empty Removed tab.
            self._clear_overrides()
        # Auto-save session on first arrival
        self._save_session()
        # Populate tables with current data, then auto-refresh from API so
        # the most up-to-date details (notes, status, etc.) are shown.
        self._refresh_tables()
        self._refresh_all_orders()

    # ── Data helpers ───────────────────────────────────────────────────────────

    def _get_active_orders(self) -> list:
        """Return non-removed orders."""
        active = []
        for o in self._neto_orders:
            key = (o.sales_channel or "Neto", o.order_id)
            if key not in self._removed_order_ids:
                active.append(o)
        for o in self._ebay_orders:
            if ("eBay", o.order_id) not in self._removed_order_ids:
                active.append(o)
        return active

    def _get_removed_orders(self) -> list:
        removed = []
        for o in self._neto_orders:
            if (o.sales_channel or "Neto", o.order_id) in self._removed_order_ids:
                removed.append(o)
        for o in self._ebay_orders:
            if ("eBay", o.order_id) in self._removed_order_ids:
                removed.append(o)
        return removed

    def _build_groups(self, orders: list) -> list[dict]:
        """Convert a list of Neto/eBay order objects to OrderTreeview group dicts.

        Runs collation on the supplied orders before building rows. Collated
        groups get a synthetic order_id; single orders use the normal path.
        """
        from src.order_collator import collate_orders

        neto_orders = [o for o in orders if hasattr(o, "date_placed")]
        ebay_orders = [o for o in orders if not hasattr(o, "date_placed")]
        coll_groups, neto_singles, ebay_singles = collate_orders(
            neto_orders, ebay_orders, self._ungrouped_order_ids
        )
        # Register collated groups so _open_detail_view can look them up
        self._collated_groups.update({g.synthetic_id: g for g in coll_groups})

        # Load assignment data once for this build pass
        _assignments: dict = {}
        _users_cache: dict = {}
        _am = getattr(self._window, "assignment_manager", None)
        _um = getattr(self._window, "user_manager", None)
        if _am:
            try:
                _assignments = _am.get_all_assignments()
            except Exception:
                pass
        if _um:
            try:
                _users_cache = {u["username"]: u for u in _um.get_active_users()}
            except Exception:
                pass

        def _assigned_label(order_id: str) -> str:
            a = _assignments.get(order_id, {})
            u = _users_cache.get(a.get("assigned_to", ""), {})
            return u.get("first_name", "") or a.get("assigned_to", "")

        # Apply assignment filter
        _filt = getattr(self, "_assign_filter", "All")
        if _filt != "All":
            cu = getattr(self._window, "current_user", None)
            current_username = cu.get("username", "") if cu else ""
            def _passes(order_id: str) -> bool:
                assigned_to = _assignments.get(order_id, {}).get("assigned_to", "")
                if _filt == "Mine":
                    return assigned_to == current_username
                return not assigned_to  # Unassigned
            neto_singles = [o for o in neto_singles if _passes(o.order_id)]
            ebay_singles = [o for o in ebay_singles if _passes(o.order_id)]
            coll_groups = []  # collated groups skipped when filtering by assignment

        result = []
        for g in coll_groups:
            d = self._group_dict_for_collated(g)
            d["assigned"] = ""  # collated groups don't have a single assignment
            result.append(d)
        for o in neto_singles + ebay_singles:
            d = self._group_dict_for_single(o)
            d["assigned"] = _assigned_label(o.order_id)
            result.append(d)
        return result

    def _group_dict_for_collated(self, g) -> dict:
        """Build a treeview row dict for a CollatedGroup."""
        env = self._window.envelope_classifications
        zones = self._window.pick_zones

        first = g.orders[0]
        is_neto = hasattr(first, "date_placed")
        if is_neto:
            ship_name = (
                f"{first.ship_first_name} {first.ship_last_name}".strip()
                or first.customer_name
            )
        else:
            ship_name = first.ship_name or first.buyer_name
        ship_state = getattr(first, "ship_state", "") or ""

        # Combined line items for child rows
        all_line_items = []
        for o in g.orders:
            o_is_neto = hasattr(o, "date_placed")
            for li in o.line_items:
                all_line_items.append({
                    "sku": li.sku,
                    "description": li.product_name if o_is_neto else li.title,
                    "qty": str(li.quantity),
                    "is_matched": False,
                })

        # Zone label across all sub-orders
        all_skus = [li["sku"] for li in all_line_items if li["sku"]]
        grp_zones = {zones.get(s, "") for s in all_skus}
        grp_zones.discard("")
        if not grp_zones:
            zone_label = "—"
        elif len(grp_zones) == 1:
            zone_label = next(iter(grp_zones))
        else:
            zone_label = "Mixed"

        # Envelope types across sub-orders
        env_types = {env.get(o.order_id, "") for o in g.orders}
        env_types.discard("")
        if not env_types:
            type_label = "—"
        elif len(env_types) == 1:
            type_label = next(iter(env_types)).capitalize()
        else:
            type_label = "Mixed"

        # Combined order notes
        notes_parts = []
        for o in g.orders:
            n = (o.notes if hasattr(o, "date_placed") else getattr(o, "buyer_notes", "")) or ""
            if n:
                notes_parts.append(f"[{o.order_id}] {n}")
        order_notes = "  |  ".join(notes_parts)

        notes = f"Collated  ·  {type_label}  ·  {zone_label}"

        return {
            "order_id": g.synthetic_id,
            "platform": g.platform,
            "customer": ship_name,
            "date": ship_state,
            "shipping": f"{len(g.orders)} orders",
            "notes": notes,
            "order_notes": order_notes,
            "postage_type": "Mixed",
            "line_items": all_line_items,
        }

    def _group_dict_for_single(self, order) -> dict:
        """Build a treeview row dict for a single (non-collated) order."""
        env = self._window.envelope_classifications
        zones = self._window.pick_zones

        is_neto = hasattr(order, "date_placed")
        if is_neto:
            platform = order.sales_channel or "Neto"
            ship_name = (
                f"{order.ship_first_name} {order.ship_last_name}".strip()
                or order.customer_name
            )
            ship_state = getattr(order, "ship_state", "") or ""
            shipping = order.shipping_type
            total_val = getattr(order, "grand_total", 0.0) or 0.0
            line_items = [
                {
                    "sku": li.sku,
                    "description": li.product_name,
                    "qty": str(li.quantity),
                    "is_matched": False,
                }
                for li in order.line_items
            ]
        else:
            platform = "eBay"
            ship_name = order.ship_name or order.buyer_name
            ship_state = getattr(order, "ship_state", "") or ""
            shipping = order.shipping_type
            total_val = getattr(order, "order_total", 0.0) or 0.0
            line_items = [
                {
                    "sku": li.sku,
                    "description": li.title,
                    "qty": str(li.quantity),
                    "is_matched": False,
                }
                for li in order.line_items
            ]

        oid = order.order_id
        env_type = env.get(oid, "")
        type_label = env_type.capitalize() if env_type else "—"

        if env_type in ("minilope", "devilope"):
            postage_type = "Envelopes"
        elif env_type == "satchel":
            postage_type = "Satchel"
        else:
            postage_type = "Mixed"

        item_skus = [li.sku for li in order.line_items if li.sku]
        order_zones = {zones.get(s, "") for s in item_skus}
        order_zones.discard("")
        if not order_zones:
            zone_label = "—"
        elif len(order_zones) == 1:
            zone_label = next(iter(order_zones))
        else:
            zone_label = "Mixed"

        total_str = f"${total_val:.2f}" if total_val else "—"
        notes = f"{type_label}  ·  {zone_label}  ·  {total_str}"

        if is_neto:
            order_notes = order.notes or ""
        else:
            order_notes = getattr(order, "buyer_notes", "") or ""

        return {
            "order_id": oid,
            "platform": platform,
            "customer": ship_name,
            "date": ship_state,
            "shipping": shipping,
            "notes": notes,
            "order_notes": order_notes,
            "postage_type": postage_type,
            "line_items": line_items,
        }

    def _find_order_data(self, order_id: str, platform: str):
        """Return (neto_order, ebay_order, matched_skus) for the given order."""
        if platform.lower() == "ebay":
            for o in self._ebay_orders:
                if o.order_id == order_id:
                    return None, o, []
        else:
            for o in self._neto_orders:
                if o.order_id == order_id:
                    return o, None, []
        return None, None, []

    # ── Table population ───────────────────────────────────────────────────────

    def _refresh_tables(self):
        self._collated_groups.clear()
        active = self._get_active_orders()
        removed = self._get_removed_orders()
        self._active_tree.load_groups(self._build_groups(active))
        self._removed_tree.load_groups(self._build_groups(removed))
        self._update_counts(active, removed)
        self._update_labels_btn()

    def _refresh_active_tab(self):
        """Refresh only the active tab (used by Envelopes only checkbox)."""
        active = self._get_active_orders()
        self._active_tree.load_groups(self._build_groups(active))
        self._update_counts(active, self._get_removed_orders())
        self._update_labels_btn()

    def _update_counts(self, active: list, removed: list):
        total = len(active) + len(removed)
        msg = (
            f"{len(active)} active  ·  {len(removed)} removed  ·  {total} total"
        )
        self._status_label.configure(text=msg, text_color=("gray50", "gray60"))

    # ── Move between tabs ──────────────────────────────────────────────────────

    def _move_to_removed(self, order_id: str, platform: str):
        if order_id in self._collated_groups:
            g = self._collated_groups[order_id]
            for oid in g.order_ids:
                self._removed_order_ids.add((g.platform, oid))
        else:
            self._removed_order_ids.add((platform, order_id))
        self._save_overrides()
        self._save_session()
        self._refresh_all_orders()

    def _move_to_active(self, order_id: str, platform: str):
        if order_id in self._collated_groups:
            g = self._collated_groups[order_id]
            for oid in g.order_ids:
                self._removed_order_ids.discard((g.platform, oid))
        else:
            self._removed_order_ids.discard((platform, order_id))
        self._save_overrides()
        self._save_session()
        self._refresh_all_orders()

    # ── Session & overrides ────────────────────────────────────────────────────

    def _save_session(self):
        from src.session_daily import save_daily_session
        save_daily_session(
            neto_orders=self._neto_orders,
            ebay_orders=self._ebay_orders,
            envelope_classifications=self._window.envelope_classifications,
            pick_zones=self._window.pick_zones,
            removed_order_ids=self._removed_order_ids,
            ungrouped_order_ids=self._ungrouped_order_ids,
        )

    def _save_overrides(self):
        from src.session_daily import save_daily_overrides
        save_daily_overrides(self._removed_order_ids)

    def _clear_overrides(self):
        from src.session_daily import save_daily_overrides
        save_daily_overrides(set())

    def _merge_overrides(self):
        from src.session_daily import load_daily_overrides
        self._removed_order_ids |= load_daily_overrides()

    # ── Refresh ────────────────────────────────────────────────────────────────

    def _refresh_all_orders(self):
        """Re-fetch all orders by ID in a background thread."""
        neto_ids = [o.order_id for o in self._neto_orders]
        ebay_ids = [o.order_id for o in self._ebay_orders]
        # Snapshot current eBay orders so we can fall back to them if the
        # eBay API is unavailable (e.g. during a Neto-only session).
        cached_ebay = list(self._ebay_orders)

        self._refresh_btn.configure(state="disabled")
        self._error_label.configure(text="Refreshing…")

        def _fetch():
            # Neto is always required — fail the whole refresh if it's unreachable.
            try:
                fresh_neto = (
                    self._window.neto_client.get_orders_by_ids(neto_ids)
                    if neto_ids else []
                )
            except Exception as exc:
                msg = str(exc)
                self.after(0, lambda m=msg: self._on_refresh_error(m))
                return

            # eBay is best-effort — if the API is unavailable, keep the cached
            # orders so that Neto-only actions (e.g. moving an order between
            # lists) still complete successfully.
            fresh_ebay = cached_ebay
            ebay_warn = ""
            if ebay_ids and self._window.ebay_client.is_authenticated():
                try:
                    fresh_ebay = self._window.ebay_client.get_orders_by_ids(ebay_ids)
                except Exception as exc:
                    fresh_ebay = cached_ebay
                    ebay_warn = "eBay unavailable — showing cached data"
                    log.warning("eBay refresh skipped (API unavailable): %s", exc)

            self.after(0, lambda n=fresh_neto, e=fresh_ebay, w=ebay_warn:
                       self._on_refresh_done(n, e, neto_ids, ebay_ids, w))

        threading.Thread(target=_fetch, daemon=True).start()

    def _on_refresh_done(self, fresh_neto: list, fresh_ebay: list, old_neto_ids: list, old_ebay_ids: list, ebay_warn: str = ""):
        old_neto_set = set(old_neto_ids)
        old_ebay_set = set(old_ebay_ids)
        fresh_neto_map = {o.order_id: o for o in fresh_neto}
        fresh_ebay_map = {o.order_id: o for o in fresh_ebay}

        # Replace orders in-place; drop any that were requested but not returned
        self._neto_orders = [
            fresh_neto_map[o.order_id] if o.order_id in old_neto_set else o
            for o in self._neto_orders
            if o.order_id not in old_neto_set or o.order_id in fresh_neto_map
        ]
        self._ebay_orders = [
            fresh_ebay_map[o.order_id] if o.order_id in old_ebay_set else o
            for o in self._ebay_orders
            if o.order_id not in old_ebay_set or o.order_id in fresh_ebay_map
        ]

        # Merge shared overrides from other workstations
        self._merge_overrides()
        # Reset assignment filter on each API refresh
        self._assign_filter = "All"
        self._assign_seg.set("All")
        self._refresh_tables()
        self._save_session()

        self._refresh_btn.configure(state="normal")
        self._error_label.configure(text=ebay_warn, text_color="orange" if ebay_warn else "")
        self._status_label.configure(
            text=f"Refreshed {datetime.now().strftime('%H:%M')}",
            text_color=("gray50", "gray60"),
        )

    def _on_refresh_error(self, msg: str):
        self._refresh_btn.configure(state="normal")
        self._error_label.configure(text=f"Refresh failed: {msg}")

    # ── Order Detail navigation ────────────────────────────────────────────────

    def _open_detail_view(self, order_id: str, platform: str):
        # Route collated groups to the dedicated CollatedDetailView
        if order_id in self._collated_groups:
            self._open_collated_view(self._collated_groups[order_id])
            return

        from src.gui.order_detail_view import OrderDetailView

        self._last_clicked_order_id = order_id
        self._last_clicked_platform = platform

        neto_order, ebay_order, _ = self._find_order_data(order_id, platform)
        if neto_order is None and ebay_order is None:
            self._error_label.configure(text=f"Order {order_id} not found")
            return

        # ── Order conflict check ──────────────────────────────────────────────
        um = getattr(self._window, "user_manager", None)
        cu = getattr(self._window, "current_user", None)
        if um and cu:
            other = um.get_processing_user(order_id)
            if other and other != f"{cu.get('first_name','')} {cu.get('last_name','')}".strip():
                from tkinter import messagebox
                if not messagebox.askyesno(
                    "Order In Progress",
                    f"{other} is currently processing order #{order_id}.\n\nTake over?",
                    parent=self.winfo_toplevel(),
                ):
                    return
            um.set_processing_order(cu["username"], order_id)
            # Also track on the root App for heartbeat polling
            app = getattr(self._window, "master", None)
            if app and hasattr(app, "_processing_order_id"):
                app._processing_order_id = order_id
                app._order_close_callback = self._clear_processing_flag

        if self._detail_frame is not None:
            self._detail_frame.destroy()

        book_freight_cb = None
        if self._window.config.shipping is not None:
            book_freight_cb = self._open_freight_view

        # Determine which tab this order is in — use it for the "move" context label
        key = (platform, order_id)
        if key in self._removed_order_ids:
            move_label_cb = lambda: self._move_to_active(order_id, platform)
        else:
            move_label_cb = lambda: self._move_to_removed(order_id, platform)

        self._detail_frame = OrderDetailView(
            self,
            order_id=order_id,
            platform=platform,
            neto_order=neto_order,
            ebay_order=ebay_order,
            matched_skus=[],
            neto_client=self._window.neto_client,
            ebay_client=self._window.ebay_client,
            dry_run=self._window.config.app.dry_run,
            on_back=self._close_detail_view,
            on_fulfilled=self._on_fulfilled,
            on_move_to_unmatched=move_label_cb,
            move_to_unmatched_label="Remove from List",
            on_book_freight=book_freight_cb,
            sku_alias_manager=self._window.sku_alias_manager,
            suppliers=self._window.config.suppliers,
            musipos_client=getattr(self._window, "musipos_client", None),
            variation_manager=getattr(self._window, "ebay_variation_manager", None),
        )
        self._detail_frame.grid(row=0, column=0, sticky="nsew")
        self._detail_frame.tkraise()

    def _clear_processing_flag(self):
        """Clear the order-processing state in UserManager and on the App."""
        um = getattr(self._window, "user_manager", None)
        cu = getattr(self._window, "current_user", None)
        if um and cu:
            try:
                um.clear_processing_order(cu["username"])
            except Exception:
                pass
        app = getattr(self._window, "master", None)
        if app and hasattr(app, "_processing_order_id"):
            app._processing_order_id = None
            app._order_close_callback = None

    def _close_detail_view(self):
        self._clear_processing_flag()
        if self._detail_frame is not None:
            self._detail_frame.destroy()
            self._detail_frame = None
        self._list_frame.tkraise()
        if self._last_clicked_order_id:
            self._active_tree.scroll_to(self._last_clicked_order_id)
        self._refresh_all_orders()

    def _on_fulfilled(self):
        """Called when an order is marked dispatched in the detail view."""
        if self._last_clicked_order_id:
            am = getattr(self._window, "assignment_manager", None)
            if am:
                try:
                    am.clear_on_dispatch(self._last_clicked_order_id)
                except Exception:
                    pass
        self._close_detail_view()  # already triggers _refresh_all_orders

    # ── Collated detail view ───────────────────────────────────────────────────

    def _open_collated_view(self, group):
        from src.gui.daily_ops.collated_detail_view import CollatedDetailView

        if self._collated_frame is not None:
            self._collated_frame.destroy()

        self._collated_frame = CollatedDetailView(
            self,
            group=group,
            window=self._window,
            dry_run=self._window.config.app.dry_run,
            on_back=self._close_collated_view,
            on_ungroup=self._ungroup_orders,
            on_book_freight=self._on_collated_book_freight,
        )
        self._collated_frame.grid(row=0, column=0, sticky="nsew")
        self._collated_frame.tkraise()

    def _close_collated_view(self):
        if self._collated_frame is not None:
            self._collated_frame.destroy()
            self._collated_frame = None
        self._list_frame.tkraise()
        self._refresh_all_orders()

    def _ungroup_orders(self, order_ids: list):
        """Move the given order IDs to the ungrouped set so they display individually."""
        self._ungrouped_order_ids.update(order_ids)
        self._save_session()
        self._refresh_tables()

    def _on_collated_book_freight(self, order_id: str, platform: str, for_all: bool = False):
        """Open FreightBookingView for a collated sub-order (or the first order when for_all)."""
        self._book_freight_for_all = for_all
        self._open_freight_view(order_id, platform)

    # ── Freight booking ────────────────────────────────────────────────────────

    def _open_freight_view(self, order_id: str, platform: str):
        from src.gui.freight_booking_view import FreightBookingView

        # Track which sub-order triggered the booking (needed for per-card tracking fill)
        self._last_clicked_order_id = order_id
        self._last_clicked_platform = platform

        neto_order, ebay_order, _ = self._find_order_data(order_id, platform)

        if self._freight_frame is not None:
            self._freight_frame.destroy()

        self._freight_frame = FreightBookingView(
            self,
            order_id=order_id,
            platform=platform,
            neto_order=neto_order,
            ebay_order=ebay_order,
            neto_client=self._window.neto_client,
            ebay_client=self._window.ebay_client,
            shipping_config=self._window.config.shipping,
            dry_run=self._window.config.app.dry_run,
            on_back=self._close_freight_view,
            on_courier_selected=lambda name, tracking="": self._on_courier_selected(name, tracking),
        )
        self._freight_frame.grid(row=0, column=0, sticky="nsew")
        self._freight_frame.tkraise()

    def _close_freight_view(self):
        if self._freight_frame is not None:
            self._freight_frame.destroy()
            self._freight_frame = None
        if self._collated_frame is not None:
            self._collated_frame.tkraise()
        elif self._detail_frame is not None:
            self._detail_frame.tkraise()

    def _on_courier_selected(self, courier_name: str, tracking_number: str = ""):
        self._close_freight_view()
        if self._book_freight_for_all and self._collated_frame is not None:
            # Apply tracking to all sub-order cards and bulk-dispatch
            if tracking_number:
                self._collated_frame.set_tracking_all(tracking_number, courier_name)
                self._collated_frame._mark_all_as_sent()
            self._book_freight_for_all = False
            return
        self._book_freight_for_all = False
        if self._collated_frame is not None:
            # Per-sub-order booking from inside the collated view
            # The FreightBookingView was opened for a specific order_id; find its card.
            # We pass tracking back to that card only.
            if tracking_number and self._last_clicked_order_id:
                self._collated_frame.set_tracking_for(
                    self._last_clicked_order_id, tracking_number, courier_name
                )
            return
        if self._detail_frame is not None:
            self._detail_frame.set_tracking(tracking=tracking_number, carrier=courier_name)
            if tracking_number:
                self._detail_frame._mark_as_sent()

    # ── Export picking list ────────────────────────────────────────────────────

    def _export_picking_list(self):
        from src.picking_list import generate_picking_list, export_picking_list_xlsx

        active = self._get_active_orders()
        items = generate_picking_list(active, self._window.pick_zones)
        if not items:
            messagebox.showinfo("Picking List", "No items to export.", parent=self)
            return

        try:
            os.makedirs(_PICKLIST_DIR, exist_ok=True)
            save_dir = _PICKLIST_DIR
        except Exception:
            save_dir = os.path.expanduser("~")

        path = os.path.join(save_dir, _PICKLIST_FILENAME)

        try:
            export_picking_list_xlsx(items, path)
            self._status_label.configure(
                text=f"Picking list exported: {_PICKLIST_FILENAME}",
                text_color=("gray50", "gray60"),
            )
            os.startfile(path)
        except Exception as exc:
            messagebox.showerror("Export Error", str(exc), parent=self)

    # ── Pick labels ────────────────────────────────────────────────────────────

    def _update_labels_btn(self):
        from src.pick_labels import build_label_list
        labels = build_label_list(self._get_active_orders(), self._window.pick_zones)
        n = len(labels)
        self._labels_btn.configure(
            text=f"Print Pick Labels ({n})",
            state="normal" if n > 0 else "disabled",
        )

    def _print_pick_labels(self):
        from src.pick_labels import PrintLabelError, build_label_list, print_pick_labels

        active = self._get_active_orders()
        labels = build_label_list(active, self._window.pick_zones)
        if not labels:
            return

        n = len(labels)
        if not messagebox.askyesno(
            "Print Pick Labels",
            f"Print {n} pick label{'s' if n != 1 else ''}?",
            parent=self,
        ):
            return

        self._labels_btn.configure(state="disabled", text=f"Printing… ({n} remaining)")

        def _run():
            try:
                def _progress(done: int, total: int):
                    remaining = total - done
                    self.after(
                        0,
                        lambda r=remaining: self._labels_btn.configure(
                            text=f"Printing… ({r} remaining)"
                        ),
                    )

                print_pick_labels(labels, progress_callback=_progress)
                self.after(0, self._on_labels_done)
            except PrintLabelError as exc:
                msg = str(exc)
                self.after(0, lambda m=msg: self._on_labels_error(m))

        threading.Thread(target=_run, daemon=True).start()

    def _on_labels_done(self):
        self._update_labels_btn()
        self._status_label.configure(
            text="Pick labels printed.", text_color=("gray50", "gray60")
        )

    def _on_labels_error(self, msg: str):
        self._update_labels_btn()
        messagebox.showerror("Print Error", msg, parent=self)

    # ── Cancel / Reprint shipment dialog ──────────────────────────────────────

    def _open_cancel_dialog(self):
        """Open the Manage Shipments dialog (cancel or reprint a booking)."""
        shipping = self._window.config.shipping
        if shipping is None:
            messagebox.showerror("Not configured", "Shipping is not configured.", parent=self)
            return

        bookings_dir = shipping.bookings_dir
        if not bookings_dir:
            messagebox.showerror("Not configured", "Bookings directory is not configured.", parent=self)
            return

        from src.shipping.booking_ledger import get_all_bookings, mark_cancelled
        from src.shipping.couriers.allied import AlliedCourier
        from src.shipping.couriers.aramex import AramexCourier
        from src.shipping.couriers.auspost import AusPostCourier
        from src.shipping.couriers.bonds import BondsCourier
        from src.shipping.couriers.dai_post import DaiPostCourier

        all_bookings = get_all_bookings(bookings_dir, days=1)

        courier_registry = {
            "auspost": AusPostCourier,
            "aramex": AramexCourier,
            "bonds": BondsCourier,
            "allied": AlliedCourier,
            "dai_post": DaiPostCourier,
        }
        couriers_by_code = {}
        for code, cls in courier_registry.items():
            cfg = shipping.couriers.get(code, {})
            if cfg.get("enabled", False):
                couriers_by_code[code] = cls(cfg)

        _sort_col = ["date"]
        _sort_asc = [False]

        win = tk.Toplevel(self)
        win.title("Manage Shipments")
        win.resizable(True, False)
        win.grab_set()

        tk.Label(
            win, text="Select a booking to cancel or reprint its label:",
            font=("Segoe UI", 11, "bold"),
        ).pack(padx=16, pady=(12, 6), anchor="w")

        tree_frame = tk.Frame(win)
        tree_frame.pack(fill="x", padx=16, pady=(0, 8))

        columns = ("date", "time", "courier", "order", "recipient", "tracking")
        col_labels = {
            "date": "Date", "time": "Time", "courier": "Courier",
            "order": "Order", "recipient": "Recipient", "tracking": "Tracking #",
        }
        col_widths = {
            "date": 90, "time": 55, "courier": 120, "order": 90,
            "recipient": 140, "tracking": 160,
        }
        col_stretch = {
            "date": False, "time": False, "courier": False,
            "order": False, "recipient": True, "tracking": True,
        }

        tree = ttk.Treeview(
            tree_frame, columns=columns, show="headings",
            height=min(max(len(all_bookings), 1), 12),
            selectmode="browse",
        )
        for col in columns:
            tree.heading(col, text=col_labels[col], command=lambda c=col: _sort_by(c))
            tree.column(col, width=col_widths[col], stretch=col_stretch[col])

        _iid_to_booking: dict[str, dict] = {}

        def _col_val(b: dict, col: str) -> str:
            if col == "time":
                t = b.get("booked_at", "")
                return t.split("T")[1][:5] if "T" in t else ""
            key_map = {
                "date": "date", "courier": "courier_name",
                "order": "order_id", "recipient": "recipient",
                "tracking": "tracking_number",
            }
            return b.get(key_map.get(col, col), "")

        def _populate():
            tree.delete(*tree.get_children())
            _iid_to_booking.clear()
            if not all_bookings:
                tree.insert("", "end", values=("", "", "No bookings found", "", "", ""))
                return
            col, asc = _sort_col[0], _sort_asc[0]
            sort_key = (
                (lambda b: b.get("booked_at", ""))
                if col in ("date", "time")
                else (lambda b: _col_val(b, col))
            )
            for b in sorted(all_bookings, key=sort_key, reverse=not asc):
                iid = tree.insert("", "end", values=tuple(_col_val(b, c) for c in columns))
                _iid_to_booking[iid] = b

        def _sort_by(col: str):
            if _sort_col[0] == col:
                _sort_asc[0] = not _sort_asc[0]
            else:
                _sort_col[0] = col
                _sort_asc[0] = True
            for c in columns:
                arrow = (" ▲" if _sort_asc[0] else " ▼") if c == _sort_col[0] else ""
                tree.heading(c, text=col_labels[c] + arrow)
            _populate()

        tree.heading("date", text="Date ▼")
        _populate()

        vsb = ttk.Scrollbar(tree_frame, orient="vertical", command=tree.yview)
        tree.configure(yscrollcommand=vsb.set)
        tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        tree_frame.grid_columnconfigure(0, weight=1)

        btn_frame = tk.Frame(win)
        btn_frame.pack(pady=(4, 4))

        cancel_btn = tk.Button(
            btn_frame, text="Cancel Shipment", font=("Segoe UI", 10),
            bg="#b22222", fg="white", activebackground="#8b0000",
            width=20, state="disabled", command=lambda: _confirm_cancel(),
        )
        cancel_btn.pack(side="left", padx=(0, 8))

        reprint_btn = tk.Button(
            btn_frame, text="Reprint Label", font=("Segoe UI", 10),
            bg="#1a6b1a", fg="white", activebackground="#0f4a0f",
            width=16, state="disabled", command=lambda: _reprint_label(),
        )
        reprint_btn.pack(side="left", padx=(0, 8))

        tk.Button(
            btn_frame, text="Close", font=("Segoe UI", 10),
            width=10, command=win.destroy,
        ).pack(side="left")

        status_lbl = tk.Label(win, text="", font=("Segoe UI", 10), wraplength=560, fg="gray40")
        status_lbl.pack(padx=16, pady=(4, 12))

        def _on_select(_event=None):
            state = "normal" if tree.selection() and all_bookings else "disabled"
            cancel_btn.configure(state=state)
            reprint_btn.configure(state=state)

        tree.bind("<<TreeviewSelect>>", _on_select)
        tree.bind("<Double-1>", lambda _e: _reprint_label())

        def _reprint_label():
            sel = tree.selection()
            if not sel:
                return
            booking = _iid_to_booking.get(sel[0])
            if not booking:
                return
            order_id = booking.get("order_id", "")
            booking_date = booking.get("date", "")
            courier_code = booking.get("print_courier_code") or booking.get("courier_code", "")
            label_path = Path(bookings_dir) / "Labels" / booking_date / f"{order_id}.pdf"
            if not label_path.exists():
                status_lbl.configure(text=f"Label not found:\n{label_path}", fg="red")
                return
            pdf_bytes = label_path.read_bytes()
            reprint_btn.configure(state="disabled")
            status_lbl.configure(text="Printing…", fg="gray40")
            win.update_idletasks()

            def _run():
                from src.shipping.label_printer import print_label
                err = print_label(pdf_bytes, courier_code=courier_code)
                win.after(0, lambda: _on_print_done(err, order_id))

            def _on_print_done(err: str, oid: str):
                reprint_btn.configure(state="normal")
                if err:
                    status_lbl.configure(text=f"Print failed: {err}", fg="red")
                else:
                    status_lbl.configure(text=f"Label for {oid} sent to printer.", fg="green")

            threading.Thread(target=_run, daemon=True).start()

        def _confirm_cancel():
            sel = tree.selection()
            if not sel:
                return
            iid = sel[0]
            booking = _iid_to_booking.get(iid)
            if not booking:
                return

            courier_code = booking.get("courier_code", "")
            courier_name = booking.get("courier_name", "")
            tracking = booking.get("tracking_number", "")
            booking_date = booking.get("date", "")

            courier = couriers_by_code.get(courier_code)
            if courier is None:
                status_lbl.configure(
                    text=f"Courier '{courier_name}' is not enabled — cannot cancel via API.",
                    fg="red",
                )
                return

            confirmed = messagebox.askyesno(
                "Confirm Cancellation",
                f"Cancel {courier_name} shipment?\n\n"
                f"Order:    {booking.get('order_id', '')}\n"
                f"Tracking: {tracking}\n"
                f"Date:     {booking_date}\n\n"
                "This cannot be undone.",
                parent=win,
            )
            if not confirmed:
                return

            cancel_btn.configure(state="disabled")
            reprint_btn.configure(state="disabled")
            status_lbl.configure(text="Cancelling…", fg="gray40")
            win.update_idletasks()

            cancel_kwargs = {}
            extras = booking.get("extras", {})
            if extras.get("booking_reference"):
                cancel_kwargs["shipment_id"] = extras["booking_reference"]
            if extras.get("postcode"):
                cancel_kwargs["postcode"] = extras["postcode"]

            def _run():
                ok, msg = courier.cancel_shipment(tracking, **cancel_kwargs)
                win.after(0, lambda: _on_cancel_done(ok, msg, iid, tracking, booking_date))

            def _on_cancel_done(ok: bool, msg: str, row_iid: str, trk: str, bdate: str):
                if ok:
                    status_lbl.configure(text=f"Cancelled.  {msg}", fg="green")
                    try:
                        mark_cancelled(bookings_dir, trk, booking_date=bdate)
                    except Exception:
                        pass
                    tree.delete(row_iid)
                    del _iid_to_booking[row_iid]
                else:
                    status_lbl.configure(text=f"Cancellation failed:\n{msg}", fg="red")
                    cancel_btn.configure(state="normal")
                    reprint_btn.configure(state="normal")

            threading.Thread(target=_run, daemon=True).start()
