from __future__ import annotations

import subprocess
import sys
import threading
from tkinter import messagebox, filedialog

import customtkinter as ctk

from src.data_processor import MatchedOrder, filter_on_po, match_orders_to_invoice
from src.exporter import export_to_xlsx
from src.gui.order_detail_modal import OrderDetailModal
from src.pdf_parser import InvoiceItem


class ReadOnlyTable(ctk.CTkScrollableFrame):
    """Scrollable read-only table using CTkLabel cells."""

    # Alternating background colors per order group: (light_mode, dark_mode)
    _GROUP_COLORS = [
        ("gray92", "gray20"),
        ("gray84", "gray28"),
    ]
    _BORDER_COLOR = ("gray65", "gray45")
    _HOVER_COLORS = [
        ("gray88", "gray24"),
        ("gray80", "gray32"),
    ]

    def __init__(self, master, columns: list[str], col_widths: list[int], **kwargs):
        super().__init__(master, **kwargs)
        self._columns = columns
        self._col_widths = col_widths
        self._header_frame: ctk.CTkFrame | None = None
        self._content_frames: list[ctk.CTkFrame] = []
        self._render_headers()

    def _configure_columns(self, frame: ctk.CTkFrame, col_offset: int, btn_width: int) -> None:
        """Set fixed widths on every grid column so widths stay consistent across frames."""
        if btn_width:
            frame.grid_columnconfigure(0, minsize=btn_width + 12, weight=0)
        for col, width in enumerate(self._col_widths):
            frame.grid_columnconfigure(col + col_offset, minsize=width + 12, weight=0)

    def _make_cell(self, parent, text: str, width: int, **kwargs) -> ctk.CTkLabel:
        """Create a label cell constrained to a fixed width with text wrapping."""
        lbl = ctk.CTkLabel(
            parent,
            text=str(text),
            width=width,
            wraplength=width,
            anchor="w",
            justify="left",
            **kwargs,
        )
        return lbl

    def _render_headers(self, col_offset: int = 0, btn_width: int = 0):
        """Render the header row in its own packed frame."""
        if self._header_frame is not None:
            self._header_frame.destroy()
        self._header_frame = ctk.CTkFrame(self, fg_color="transparent")
        self._header_frame.pack(fill="x", padx=4, pady=(4, 2))
        self._configure_columns(self._header_frame, col_offset, btn_width)

        if btn_width:
            ctk.CTkLabel(
                self._header_frame, text="", width=btn_width
            ).grid(row=0, column=0, padx=(4, 8), pady=(2, 4), sticky="w")

        for col, (header, width) in enumerate(zip(self._columns, self._col_widths)):
            self._make_cell(
                self._header_frame, header, width,
                font=ctk.CTkFont(weight="bold"),
            ).grid(row=0, column=col + col_offset, padx=(4, 8), pady=(2, 4), sticky="ew")

    def load_rows(
        self,
        rows: list[list[str]],
        group_key_col: int | None = None,
        group_button: dict | None = None,
        on_row_click: callable | None = None,
    ):
        """
        Load rows into the table.

        Each order group is rendered inside a bordered CTkFrame so orders are
        visually separated with an outline. The action button (if any) appears
        at the left of the first row of each group.

        on_row_click(group_key, platform) is called when a row is clicked (not the button).
        """
        for frame in self._content_frames:
            frame.destroy()
        self._content_frames = []

        btn_width = group_button.get("width", 60) if group_button else 0
        col_offset = 1 if group_button else 0

        self._render_headers(col_offset=col_offset, btn_width=btn_width)

        if group_key_col is None:
            # No grouping — simple alternating rows, no border
            for row_idx, row_data in enumerate(rows):
                colors = self._GROUP_COLORS[row_idx % 2]
                row_frame = ctk.CTkFrame(self, fg_color=colors, corner_radius=2)
                row_frame.pack(fill="x", padx=4, pady=1)
                self._configure_columns(row_frame, 0, 0)
                self._content_frames.append(row_frame)
                for col, (val, width) in enumerate(zip(row_data, self._col_widths)):
                    self._make_cell(
                        row_frame, val, width, fg_color="transparent",
                    ).grid(row=0, column=col, padx=(4, 8), pady=2, sticky="ew")
            return

        # Pre-group rows by the group key column
        groups: list[tuple[str, list]] = []
        for row_data in rows:
            key = row_data[group_key_col] if group_key_col < len(row_data) else ""
            if groups and groups[-1][0] == key:
                groups[-1][1].append(row_data)
            else:
                groups.append((key, [row_data]))

        for group_idx, (key, group_rows) in enumerate(groups):
            colors = self._GROUP_COLORS[group_idx % 2]
            hover_colors = self._HOVER_COLORS[group_idx % 2]

            # One bordered frame per order group
            group_frame = ctk.CTkFrame(
                self,
                border_width=1,
                border_color=self._BORDER_COLOR,
                corner_radius=4,
                fg_color=colors,
            )
            group_frame.pack(fill="x", padx=4, pady=3)
            self._configure_columns(group_frame, col_offset, btn_width)
            self._content_frames.append(group_frame)

            # Determine platform from first row (column index 1 for matched, 0 for unmatched)
            platform_col = 1 if col_offset else 0
            platform = group_rows[0][platform_col] if platform_col < len(group_rows[0]) else ""

            # Hover highlight bindings
            if on_row_click:
                def _on_enter(e, f=group_frame, hc=hover_colors):
                    f.configure(fg_color=hc)
                def _on_leave(e, f=group_frame, c=colors):
                    f.configure(fg_color=c)
                group_frame.bind("<Enter>", _on_enter)
                group_frame.bind("<Leave>", _on_leave)

            for row_idx, row_data in enumerate(group_rows):
                # Action button on the first row only, at column 0
                if group_button and row_idx == 0:
                    ctk.CTkButton(
                        group_frame,
                        text=group_button["text"],
                        width=btn_width,
                        height=24,
                        font=ctk.CTkFont(size=11),
                        fg_color="gray50",
                        hover_color="gray40",
                        command=lambda k=key: group_button["callback"](k),
                    ).grid(row=row_idx, column=0, padx=(4, 8), pady=2, sticky="w")

                for col, (val, width) in enumerate(zip(row_data, self._col_widths)):
                    cell = self._make_cell(
                        group_frame, val, width, fg_color="transparent",
                    )
                    cell.grid(row=row_idx, column=col + col_offset, padx=(4, 8), pady=1, sticky="ew")

                    # Make cells clickable
                    if on_row_click:
                        cell.configure(cursor="hand2")
                        cell.bind("<Button-1>", lambda e, k=key, p=platform: on_row_click(k, p))


class ResultsTab(ctk.CTkFrame):
    def __init__(self, master, app, **kwargs):
        super().__init__(master, fg_color="transparent", **kwargs)
        self._app = app
        self._matched: list[MatchedOrder] = []
        self._unmatched_inv: list[InvoiceItem] = []
        # Deduped order lists — set during load_results for use by summary/unmatched views
        self._neto_orders = []
        self._ebay_orders = []
        # Manual overrides: orders moved between matched ↔ unmatched by the user.
        # Keys are (platform, order_id) tuples.
        self._excluded_order_ids: set[tuple[str, str]] = set()
        self._force_matched_order_ids: set[tuple[str, str]] = set()
        self._build_ui()

    def _build_ui(self):
        # ── Summary row ───────────────────────────────────────────────────
        summary = ctk.CTkFrame(self, fg_color="transparent")
        summary.pack(fill="x", padx=12, pady=(12, 6))

        self._matched_lbl = ctk.CTkLabel(
            summary, text="Matched: —", font=ctk.CTkFont(size=13, weight="bold")
        )
        self._matched_lbl.pack(side="left", padx=(0, 24))

        self._unmatched_inv_lbl = ctk.CTkLabel(
            summary, text="Unmatched invoice items: —", font=ctk.CTkFont(size=13)
        )
        self._unmatched_inv_lbl.pack(side="left", padx=(0, 24))

        self._unmatched_orders_lbl = ctk.CTkLabel(
            summary, text="'On PO' orders with no match: —", font=ctk.CTkFont(size=13)
        )
        self._unmatched_orders_lbl.pack(side="left")

        # ── Inner tab view ────────────────────────────────────────────────
        self._inner_tabs = ctk.CTkTabview(self, corner_radius=6)
        self._inner_tabs.pack(fill="both", expand=True, padx=12, pady=4)

        for name in ("Matched Orders", "Unmatched Invoice Items", "Unmatched Orders"):
            self._inner_tabs.add(name)

        # Matched Orders table — includes "*" column to flag items that arrived with invoice
        MATCHED_COLS = ["*", "Platform", "Order No.", "Customer", "Date", "SKU", "Description", "Qty", "Notes"]
        MATCHED_WIDTHS = [20, 80, 110, 130, 90, 130, 200, 40, 230]
        self._matched_table = ReadOnlyTable(
            self._inner_tabs.tab("Matched Orders"),
            columns=MATCHED_COLS,
            col_widths=MATCHED_WIDTHS,
            corner_radius=4,
        )
        self._matched_table.pack(fill="both", expand=True)

        # Unmatched invoice items
        INV_COLS = ["SKU (with suffix)", "Description", "Qty"]
        INV_WIDTHS = [200, 450, 60]
        self._inv_table = ReadOnlyTable(
            self._inner_tabs.tab("Unmatched Invoice Items"),
            columns=INV_COLS,
            col_widths=INV_WIDTHS,
            corner_radius=4,
        )
        self._inv_table.pack(fill="both", expand=True)

        # Unmatched orders
        self._unmatched_orders_note = ctk.CTkLabel(
            self._inner_tabs.tab("Unmatched Orders"),
            text=(
                "Neto: 'on PO' orders (paid, undispatched) whose SKUs did not match the invoice.\n"
                "eBay: all paid, unfulfilled orders whose SKUs did not match the invoice.\n\n"
                "This may indicate the ordered stock is arriving in a future delivery,\n"
                "or was purchased via phone/counter (not via an online channel)."
            ),
            font=ctk.CTkFont(size=13),
            justify="left",
        )
        self._unmatched_orders_note.pack(padx=20, pady=20, anchor="nw")

        UNMATCHED_COLS = ["Platform", "Order No.", "Customer", "Date", "SKU", "Description", "Qty", "Notes"]
        UNMATCHED_WIDTHS = [80, 110, 130, 90, 130, 200, 40, 230]
        self._unmatched_orders_table = ReadOnlyTable(
            self._inner_tabs.tab("Unmatched Orders"),
            columns=UNMATCHED_COLS,
            col_widths=UNMATCHED_WIDTHS,
            corner_radius=4,
        )
        self._unmatched_orders_table.pack(fill="both", expand=True, padx=0, pady=(0, 0))

        # ── Bottom row: export + refresh + save ─────────────────────────────
        bottom = ctk.CTkFrame(self, fg_color="transparent")
        bottom.pack(fill="x", padx=12, pady=(4, 12))

        self._export_btn = ctk.CTkButton(
            bottom,
            text="Export to Excel",
            width=140,
            command=self._export_csv,
        )
        self._export_btn.pack(side="left")

        self._refresh_btn = ctk.CTkButton(
            bottom,
            text="Refresh Orders",
            width=130,
            fg_color=("dodgerblue3", "dodgerblue4"),
            command=self._refresh_orders,
        )
        self._refresh_btn.pack(side="left", padx=(12, 0))

        self._save_session_btn = ctk.CTkButton(
            bottom,
            text="Save Session As",
            width=130,
            fg_color="gray50",
            hover_color="gray40",
            command=self._save_session_as,
        )
        self._save_session_btn.pack(side="left", padx=(12, 0))

        self._export_label = ctk.CTkLabel(
            bottom, text="", font=ctk.CTkFont(size=12), text_color="gray60"
        )
        self._export_label.pack(side="left", padx=(12, 0))

        self._error_label = ctk.CTkLabel(
            bottom, text="", font=ctk.CTkFont(size=12), text_color="red"
        )
        self._error_label.pack(side="left", padx=(12, 0))

    # ── Public ────────────────────────────────────────────────────────────

    def load_results(self):
        """Called by App when switching to this tab. Runs matching and updates UI."""
        try:
            self._error_label.configure(text="Loading…")
            self.update_idletasks()

            invoice_items = self._app.invoice_tab.get_invoice_items()

            # Neto orders already have eBay channel excluded; eBay orders come directly from eBay API
            self._neto_orders = self._app.neto_orders
            self._ebay_orders = self._app.ebay_orders

            matched, unmatched_inv = match_orders_to_invoice(
                invoice_items,
                self._neto_orders,
                self._ebay_orders,
                on_po_phrase=self._app.config.app.on_po_filter_phrase,
            )

            # Reset manual overrides on fresh data load
            self._excluded_order_ids.clear()
            self._force_matched_order_ids.clear()

            self._matched = matched
            self._unmatched_inv = unmatched_inv
            self._app.matched_orders = matched

            self._refresh_tables()
            self._error_label.configure(text="")

            # Auto-save session snapshot
            try:
                from src.session import save_snapshot
                save_dir = self._app.config.app.snapshot_dir or self._app.config.app.output_dir
                save_snapshot(
                    save_dir=save_dir,
                    invoice_items=invoice_items,
                    neto_orders=self._neto_orders,
                    ebay_orders=self._ebay_orders,
                    matched_orders=matched,
                    unmatched_inv=unmatched_inv,
                    excluded_ids=self._excluded_order_ids,
                    force_matched_ids=self._force_matched_order_ids,
                )
            except Exception:
                pass  # Auto-save is non-critical
        except Exception as exc:
            import traceback, sys
            traceback.print_exc(file=sys.stderr)
            self._error_label.configure(text=f"Error loading results: {exc}")

    def _refresh_tables(self):
        """Re-render matched and unmatched tables with current overrides applied."""
        # Effective matched: original matched minus excluded, plus force-matched
        effective_matched = [
            m for m in self._matched
            if (m.platform, m.order_id) not in self._excluded_order_ids
        ]

        # Build force-matched MatchedOrder entries from the raw order data
        force_matched = self._build_force_matched()
        effective_matched.extend(force_matched)

        # Update the app's matched_orders for export
        self._app.matched_orders = effective_matched

        self._populate_matched(effective_matched)
        self._populate_unmatched_inv(self._unmatched_inv)
        self._populate_unmatched_orders(effective_matched)
        self._update_summary(effective_matched, self._unmatched_inv)

    def _build_force_matched(self) -> list[MatchedOrder]:
        """Create MatchedOrder entries for orders manually moved to matched."""
        if not self._force_matched_order_ids:
            return []

        result = []
        # TODO: filter disabled — iterate all orders
        # for order in filter_on_po(self._neto_orders, self._app.config.app.on_po_filter_phrase):
        for order in self._neto_orders:
            channel = order.sales_channel or "Neto"
            key = (channel, order.order_id)
            if key not in self._force_matched_order_ids:
                continue
            order_date = order.date_paid or order.date_placed
            for idx, line in enumerate(order.line_items):
                result.append(MatchedOrder(
                    platform=channel,
                    order_id=order.order_id,
                    customer_name=order.customer_name,
                    order_date=order_date,
                    sku=line.sku,
                    description=line.product_name,
                    quantity=line.quantity,
                    notes=order.notes if idx == 0 else "",
                    invoice_sku="",
                    invoice_description="",
                    invoice_qty=0,
                    is_invoice_match=False,
                ))

        for order in self._ebay_orders:
            key = ("eBay", order.order_id)
            if key not in self._force_matched_order_ids:
                continue
            for line in order.line_items:
                result.append(MatchedOrder(
                    platform="eBay",
                    order_id=order.order_id,
                    customer_name=order.buyer_name,
                    order_date=order.creation_date,
                    sku=line.sku,
                    description=line.title,
                    quantity=line.quantity,
                    notes=line.notes,
                    invoice_sku="",
                    invoice_description="",
                    invoice_qty=0,
                    is_invoice_match=False,
                ))

        return result

    def _exclude_order(self, platform_and_order_id: str):
        """Move an order from matched → unmatched."""
        # Find the platform for this order_id from the matched list
        for m in self._matched:
            if m.order_id == platform_and_order_id:
                key = (m.platform, m.order_id)
                self._excluded_order_ids.add(key)
                # Also remove from force-matched if it was there
                self._force_matched_order_ids.discard(key)
                break
        else:
            # Check effective matched (could be a force-matched order)
            for key in list(self._force_matched_order_ids):
                if key[1] == platform_and_order_id:
                    self._force_matched_order_ids.discard(key)
                    break
        self._refresh_tables()

    def _include_order(self, order_id: str):
        """Move an order from unmatched → matched."""
        # First check if this was originally matched but excluded
        for m in self._matched:
            if m.order_id == order_id:
                key = (m.platform, m.order_id)
                if key in self._excluded_order_ids:
                    self._excluded_order_ids.discard(key)
                    self._refresh_tables()
                    return

        # Otherwise, force-add it from the raw order lists
        for order in self._neto_orders:
            channel = order.sales_channel or "Neto"
            if order.order_id == order_id:
                self._force_matched_order_ids.add((channel, order.order_id))
                self._refresh_tables()
                return

        for order in self._ebay_orders:
            if order.order_id == order_id:
                self._force_matched_order_ids.add(("eBay", order.order_id))
                self._refresh_tables()
                return

    def _update_summary(self, matched, unmatched_inv):
        # TODO: filter disabled — count all awaiting-shipment orders
        # on_po_neto = filter_on_po(self._neto_orders, self._app.config.app.on_po_filter_phrase)
        # on_po_ebay = filter_on_po(self._ebay_orders, self._app.config.app.on_po_filter_phrase)
        candidate_count = len(self._neto_orders) + len(self._ebay_orders)

        matched_order_ids = {(m.platform, m.order_id) for m in matched}
        unmatched_order_count = max(0, candidate_count - len(matched_order_ids))

        # Count only lines that are actual invoice matches
        match_count = sum(1 for m in matched if m.is_invoice_match)

        self._matched_lbl.configure(
            text=f"Matched: {match_count} invoice line{'s' if match_count != 1 else ''}",
            text_color=("green" if matched else "gray50"),
        )
        self._unmatched_inv_lbl.configure(
            text=f"Unmatched invoice items: {len(unmatched_inv)}"
        )
        self._unmatched_orders_lbl.configure(
            text=f"Unmatched orders: {unmatched_order_count}"
        )

    def _populate_matched(self, matched: list[MatchedOrder]):
        def _platform_key(m: MatchedOrder) -> tuple:
            pl = m.platform.lower()
            if pl == "website":
                return (0, m.platform, m.order_id)
            if pl == "ebay":
                return (2, m.platform, m.order_id)
            return (1, m.platform, m.order_id)

        rows = []
        for m in sorted(matched, key=_platform_key):
            date_str = m.order_date.strftime("%Y-%m-%d") if m.order_date else ""
            arrived = "*" if m.is_invoice_match else ""
            rows.append([
                arrived,
                m.platform,
                m.order_id,
                m.customer_name,
                date_str,
                m.sku,
                m.description,
                str(m.quantity),
                m.notes,
            ])
        self._matched_table.load_rows(
            rows,
            group_key_col=2,
            group_button={
                "text": "Remove",
                "callback": self._exclude_order,
                "width": 65,
            },
            on_row_click=self._open_order_detail,
        )

    def _populate_unmatched_inv(self, items: list[InvoiceItem]):
        rows = [[item.sku_with_suffix, item.description, str(item.quantity)] for item in items]
        self._inv_table.load_rows(rows)

    def _populate_unmatched_orders(self, matched):
        phrase = self._app.config.app.on_po_filter_phrase
        matched_ids = {(m.platform, m.order_id) for m in matched}

        def _platform_key(row: list) -> tuple:
            pl = row[0].lower()  # row[0] is the channel/platform string
            if pl == "website":
                return (0, row[0], row[1])
            if pl == "ebay":
                return (2, row[0], row[1])
            return (1, row[0], row[1])

        rows = []
        # TODO: filter disabled — showing all awaiting-shipment orders
        # for order in filter_on_po(self._neto_orders, phrase):
        for order in self._neto_orders:
            channel = order.sales_channel or "Neto"
            if (channel, order.order_id) not in matched_ids:
                date_str = order.date_paid.strftime("%Y-%m-%d") if order.date_paid else ""
                # One row per line item; notes are order-level so only on first item
                for idx, line in enumerate(order.line_items):
                    rows.append([
                        channel, order.order_id, order.customer_name, date_str,
                        line.sku, line.product_name, str(line.quantity),
                        order.notes if idx == 0 else "",
                    ])

        # All eBay orders are candidates — show any that didn't match the invoice
        for order in self._ebay_orders:
            if ("eBay", order.order_id) not in matched_ids:
                date_str = order.creation_date.strftime("%Y-%m-%d") if order.creation_date else ""
                # One row per line item; eBay notes are item-level
                for line in order.line_items:
                    rows.append([
                        "eBay", order.order_id, order.buyer_name, date_str,
                        line.sku, line.title, str(line.quantity), line.notes,
                    ])

        rows.sort(key=_platform_key)
        self._unmatched_orders_table.load_rows(
            rows,
            group_key_col=1,
            group_button={
                "text": "Add",
                "callback": self._include_order,
                "width": 55,
            },
        )

    # ── Order Detail Modal ───────────────────────────────────────────────

    def _open_order_detail(self, order_id: str, platform: str):
        """Check order status via API, then open detail modal if still active."""
        self._error_label.configure(text="Checking order status...")
        self.update_idletasks()

        def _check_and_open():
            try:
                status = ""
                if platform.lower() == "ebay":
                    status = self._app.ebay_client.get_order_status(order_id)
                    is_completed = status in ("FULFILLED",)
                else:
                    status = self._app.neto_client.get_order_status(order_id)
                    is_completed = status.lower() in ("dispatched", "shipped", "completed")

                self.after(0, lambda: self._handle_status_check(order_id, platform, is_completed))
            except Exception as e:
                # If status check fails (e.g. network error), open the modal anyway
                self.after(0, lambda: self._show_order_modal(order_id, platform))
                self.after(0, lambda: self._error_label.configure(text=f"Status check failed: {e}"))

        threading.Thread(target=_check_and_open, daemon=True).start()

    def _handle_status_check(self, order_id: str, platform: str, is_completed: bool):
        self._error_label.configure(text="")
        if is_completed:
            messagebox.showinfo(
                "Order Completed",
                f"Order {order_id} has already been completed by another user.\n\n"
                "The orders list will now refresh.",
                parent=self,
            )
            self._refresh_orders()
        else:
            self._show_order_modal(order_id, platform)

    def _show_order_modal(self, order_id: str, platform: str):
        neto_order = None
        ebay_order = None
        matched_skus = []

        # Find the raw order object
        if platform.lower() == "ebay":
            for o in self._ebay_orders:
                if o.order_id == order_id:
                    ebay_order = o
                    break
        else:
            for o in self._neto_orders:
                if o.order_id == order_id:
                    neto_order = o
                    break

        # Gather matched SKUs for this order
        for m in self._matched:
            if m.order_id == order_id and m.is_invoice_match:
                matched_skus.append(m.sku)

        OrderDetailModal(
            self,
            order_id=order_id,
            platform=platform,
            neto_order=neto_order,
            ebay_order=ebay_order,
            matched_skus=matched_skus,
            neto_client=self._app.neto_client,
            ebay_client=self._app.ebay_client,
            dry_run=self._app.config.app.dry_run,
            on_close_callback=self._on_modal_close,
        )

    def _on_modal_close(self, completed: bool):
        if completed:
            self._refresh_orders()

    def _refresh_orders(self):
        """Re-fetch orders from APIs and re-run matching."""
        self._error_label.configure(text="Refreshing orders...")
        self._refresh_btn.configure(state="disabled")
        self.update_idletasks()

        def _fetch():
            try:
                from datetime import datetime, timedelta
                lookback = self._app.config.app.order_lookback_days
                date_to = datetime.now()
                date_from = date_to - timedelta(days=lookback)

                neto_orders = self._app.neto_client.get_overdue_orders(date_from, date_to)
                ebay_orders = []
                if self._app.ebay_client.is_authenticated():
                    ebay_orders = self._app.ebay_client.get_overdue_orders(date_from, date_to)

                self.after(0, lambda: self._apply_refreshed_orders(neto_orders, ebay_orders))
            except Exception as e:
                self.after(0, lambda: self._error_label.configure(text=f"Refresh failed: {e}"))
                self.after(0, lambda: self._refresh_btn.configure(state="normal"))

        threading.Thread(target=_fetch, daemon=True).start()

    def _apply_refreshed_orders(self, neto_orders, ebay_orders):
        self._app.neto_orders = neto_orders
        self._app.ebay_orders = ebay_orders
        self._neto_orders = neto_orders
        self._ebay_orders = ebay_orders

        invoice_items = self._app.invoice_tab.get_invoice_items()
        matched, unmatched_inv = match_orders_to_invoice(
            invoice_items,
            self._neto_orders,
            self._ebay_orders,
            on_po_phrase=self._app.config.app.on_po_filter_phrase,
        )
        self._matched = matched
        self._unmatched_inv = unmatched_inv
        self._excluded_order_ids.clear()
        self._force_matched_order_ids.clear()
        self._refresh_tables()
        self._refresh_btn.configure(state="normal")
        self._error_label.configure(text="")

    def _save_session_as(self):
        """Let the user choose a directory and save the session snapshot there."""
        from src.session import save_snapshot
        initial_dir = self._app.config.app.snapshot_dir or self._app.config.app.output_dir
        save_dir = filedialog.askdirectory(
            title="Choose snapshot save location",
            initialdir=initial_dir,
            parent=self,
        )
        if not save_dir:
            return
        try:
            invoice_items = self._app.invoice_tab.get_invoice_items()
            path = save_snapshot(
                save_dir=save_dir,
                invoice_items=invoice_items,
                neto_orders=self._neto_orders,
                ebay_orders=self._ebay_orders,
                matched_orders=self._app.matched_orders,
                unmatched_inv=self._unmatched_inv,
                excluded_ids=self._excluded_order_ids,
                force_matched_ids=self._force_matched_order_ids,
            )
            self._export_label.configure(text=f"Session saved: {path}", text_color="green")
        except Exception as e:
            self._error_label.configure(text=f"Save failed: {e}")

    # ── Export ────────────────────────────────────────────────────────────

    def _export_csv(self):
        if not self._app.matched_orders:
            messagebox.showinfo("No Data", "There are no matched orders to export.")
            return
        try:
            path = export_to_xlsx(
                self._app.matched_orders,
                output_dir=self._app.config.app.output_dir,
            )
            self._export_label.configure(text=f"Saved: {path}", text_color="green")
            self._error_label.configure(text="")
            # Open the output folder in Explorer
            subprocess.Popen(["explorer", "/select,", path])
        except Exception as e:
            self._error_label.configure(text=f"Export failed: {e}")
            self._export_label.configure(text="")
