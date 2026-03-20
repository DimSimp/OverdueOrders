from __future__ import annotations

import logging
import os
import subprocess
import threading
import tkinter as tk
from datetime import datetime
from pathlib import Path
from tkinter import BooleanVar, messagebox
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
    "sku":         ("SKU",         130),
    "description": ("Description", 180),
    "qty":         ("Qty",          40),
    "notes":       ("Type / Zone", 140),
    "order_notes": ("Notes",       150),
}

# Fixed picking-list output directory and filename (overwritten on every export)
_PICKLIST_DIR = r"\\SERVER\Project Folder\Order-Fulfillment-App\Picking Lists\Daily"
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

        self.grid_rowconfigure(0, weight=1)
        self.grid_columnconfigure(0, weight=1)

        self._build_ui()

    # ── UI construction ────────────────────────────────────────────────────────

    def _build_ui(self):
        # List frame — raised normally; sits at (0,0) so detail/freight overlay it
        self._list_frame = ctk.CTkFrame(self, fg_color="transparent")
        self._list_frame.grid(row=0, column=0, sticky="nsew")
        self._list_frame.grid_rowconfigure(1, weight=1)
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

        self._env_only_var = BooleanVar(value=False)
        ctk.CTkCheckBox(
            toolbar, text="Envelopes only",
            variable=self._env_only_var,
            font=ctk.CTkFont(size=13),
            command=self._refresh_active_tab,
        ).pack(side="left", padx=(0, 0))

        self._status_label = ctk.CTkLabel(
            toolbar, text="", font=ctk.CTkFont(size=12),
            text_color=("gray50", "gray60"),
        )
        self._status_label.pack(side="right")

        self._error_label = ctk.CTkLabel(
            toolbar, text="", font=ctk.CTkFont(size=12), text_color="red",
        )
        self._error_label.pack(side="right", padx=(0, 8))

        # ── Tab view ────────────────────────────────────────────────────────
        self._tabs = ctk.CTkTabview(parent, corner_radius=6)
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
            on_row_click=self._open_detail_view,
            on_context_action=self._move_to_removed,
            context_label="Move to Removed",
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
            on_row_click=self._open_detail_view,
            on_context_action=self._move_to_active,
            context_label="Move back to Active",
        )
        self._removed_tree.grid(row=0, column=0, sticky="nsew")

        # ── Bottom nav ───────────────────────────────────────────────────────
        bottom = ctk.CTkFrame(parent, fg_color="transparent")
        bottom.grid(row=2, column=0, sticky="ew", padx=12, pady=(4, 12))

        ctk.CTkButton(
            bottom, text="← Back",
            width=120,
            fg_color=("gray70", "gray30"), hover_color=("gray60", "gray25"),
            command=self._on_back,
        ).pack(side="left")

    # ── Entry point ────────────────────────────────────────────────────────────

    def show(self, initial_removed_ids=None):
        """Called by DailyOpsWindow each time this step is shown. Loads only once.

        initial_removed_ids: optional set of (platform, order_id) tuples restored
        from a saved session file (passed when loading rather than generating).
        """
        if self._loaded:
            return
        self._loaded = True
        # Snapshot orders from window (may have been restored from a session file)
        self._neto_orders = list(self._window.neto_orders)
        self._ebay_orders = list(self._window.ebay_orders)
        # Pre-load removed IDs from session restore (before merging overrides)
        if initial_removed_ids:
            self._removed_order_ids = set(initial_removed_ids)
        # Merge shared overrides from other workstations
        self._merge_overrides()
        # Auto-save session on first arrival
        self._save_session()
        # Populate tables
        self._refresh_tables()

    # ── Data helpers ───────────────────────────────────────────────────────────

    def _get_active_orders(self) -> list:
        """Return non-removed orders, optionally filtered to envelopes only."""
        active = []
        for o in self._neto_orders:
            key = (o.sales_channel or "Neto", o.order_id)
            if key not in self._removed_order_ids:
                active.append(o)
        for o in self._ebay_orders:
            if ("eBay", o.order_id) not in self._removed_order_ids:
                active.append(o)
        if self._env_only_var.get():
            env = self._window.envelope_classifications
            active = [
                o for o in active
                if env.get(o.order_id, "") in ("minilope", "devilope")
            ]
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
        """Convert a list of Neto/eBay order objects to OrderTreeview group dicts."""
        env = self._window.envelope_classifications
        zones = self._window.pick_zones
        groups = []

        for order in orders:
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
            # Envelope type label
            env_type = env.get(oid, "")
            type_label = env_type.capitalize() if env_type else "—"

            # Zone label — consolidate across all line items
            item_skus = [li.sku for li in order.line_items if li.sku]
            order_zones = {zones.get(s, "") for s in item_skus}
            order_zones.discard("")
            if not order_zones:
                zone_label = "—"
            elif len(order_zones) == 1:
                zone_label = next(iter(order_zones))
            else:
                zone_label = "Mixed"

            # Total string
            total_str = f"${total_val:.2f}" if total_val else "—"

            # Pack type/zone/total into the notes column for the parent row.
            # Line items will naturally show their sku/description/qty in those columns.
            notes = f"{type_label}  ·  {zone_label}  ·  {total_str}"

            if is_neto:
                order_notes = order.notes or ""
            else:
                order_notes = getattr(order, "buyer_notes", "") or ""

            groups.append({
                "order_id": oid,
                "platform": platform,
                "customer": ship_name,
                "date": ship_state,
                "shipping": shipping,
                "notes": notes,
                "order_notes": order_notes,
                "line_items": line_items,
            })

        return groups

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
        self._removed_order_ids.add((platform, order_id))
        self._save_overrides()
        self._save_session()
        self._refresh_tables()

    def _move_to_active(self, order_id: str, platform: str):
        self._removed_order_ids.discard((platform, order_id))
        self._save_overrides()
        self._save_session()
        self._refresh_tables()

    # ── Session & overrides ────────────────────────────────────────────────────

    def _save_session(self):
        from src.session_daily import save_daily_session
        save_daily_session(
            neto_orders=self._neto_orders,
            ebay_orders=self._ebay_orders,
            envelope_classifications=self._window.envelope_classifications,
            pick_zones=self._window.pick_zones,
            removed_order_ids=self._removed_order_ids,
        )

    def _save_overrides(self):
        from src.session_daily import save_daily_overrides
        save_daily_overrides(self._removed_order_ids)

    def _merge_overrides(self):
        from src.session_daily import load_daily_overrides
        self._removed_order_ids |= load_daily_overrides()

    # ── Refresh ────────────────────────────────────────────────────────────────

    def _refresh_all_orders(self):
        """Re-fetch all orders by ID in a background thread."""
        neto_ids = [o.order_id for o in self._neto_orders]
        ebay_ids = [o.order_id for o in self._ebay_orders]

        self._refresh_btn.configure(state="disabled")
        self._error_label.configure(text="Refreshing…")

        def _fetch():
            try:
                fresh_neto = (
                    self._window.neto_client.get_orders_by_ids(neto_ids)
                    if neto_ids else []
                )
                fresh_ebay = []
                if ebay_ids and self._window.ebay_client.is_authenticated():
                    fresh_ebay = self._window.ebay_client.get_orders_by_ids(ebay_ids)
                self.after(0, lambda n=fresh_neto, e=fresh_ebay: self._on_refresh_done(n, e, neto_ids, ebay_ids))
            except Exception as exc:
                msg = str(exc)
                self.after(0, lambda m=msg: self._on_refresh_error(m))

        threading.Thread(target=_fetch, daemon=True).start()

    def _on_refresh_done(self, fresh_neto: list, fresh_ebay: list, old_neto_ids: list, old_ebay_ids: list):
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
        self._refresh_tables()
        self._save_session()

        self._refresh_btn.configure(state="normal")
        self._error_label.configure(text="")
        self._status_label.configure(
            text=f"Refreshed {datetime.now().strftime('%H:%M')}",
            text_color=("gray50", "gray60"),
        )

    def _on_refresh_error(self, msg: str):
        self._refresh_btn.configure(state="normal")
        self._error_label.configure(text=f"Refresh failed: {msg}")

    # ── Order Detail navigation ────────────────────────────────────────────────

    def _open_detail_view(self, order_id: str, platform: str):
        from src.gui.order_detail_view import OrderDetailView

        self._last_clicked_order_id = order_id
        self._last_clicked_platform = platform

        neto_order, ebay_order, _ = self._find_order_data(order_id, platform)
        if neto_order is None and ebay_order is None:
            self._error_label.configure(text=f"Order {order_id} not found")
            return

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
            on_book_freight=book_freight_cb,
            sku_alias_manager=self._window.sku_alias_manager,
            suppliers=self._window.config.suppliers,
        )
        self._detail_frame.grid(row=0, column=0, sticky="nsew")
        self._detail_frame.tkraise()

    def _close_detail_view(self):
        if self._detail_frame is not None:
            self._detail_frame.destroy()
            self._detail_frame = None
        self._list_frame.tkraise()
        if self._last_clicked_order_id:
            self._active_tree.scroll_to(self._last_clicked_order_id)

    def _on_fulfilled(self):
        """Called when an order is marked dispatched in the detail view."""
        self._close_detail_view()
        self._refresh_all_orders()

    # ── Freight booking ────────────────────────────────────────────────────────

    def _open_freight_view(self, order_id: str, platform: str):
        from src.gui.freight_booking_view import FreightBookingView

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
        if self._detail_frame is not None:
            self._detail_frame.tkraise()

    def _on_courier_selected(self, courier_name: str, tracking_number: str = ""):
        self._close_freight_view()
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
            courier_code = booking.get("courier_code", "")
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
