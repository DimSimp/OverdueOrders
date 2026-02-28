import subprocess
import sys
from tkinter import messagebox

import customtkinter as ctk

from src.data_processor import MatchedOrder, match_orders_to_invoice
from src.exporter import export_to_csv
from src.pdf_parser import InvoiceItem


class ReadOnlyTable(ctk.CTkScrollableFrame):
    """Scrollable read-only table using CTkLabel cells."""

    def __init__(self, master, columns: list[str], col_widths: list[int], **kwargs):
        super().__init__(master, **kwargs)
        self._columns = columns
        self._col_widths = col_widths
        self._render_headers()

    def _render_headers(self):
        for col, (header, width) in enumerate(zip(self._columns, self._col_widths)):
            lbl = ctk.CTkLabel(
                self,
                text=header,
                font=ctk.CTkFont(weight="bold"),
                width=width,
                anchor="w",
            )
            lbl.grid(row=0, column=col, padx=(4, 8), pady=(4, 6), sticky="w")

    def load_rows(self, rows: list[list[str]]):
        # Clear existing data rows (not header)
        for widget in self.winfo_children():
            info = widget.grid_info()
            if info and int(info.get("row", 0)) > 0:
                widget.destroy()

        for row_idx, row_data in enumerate(rows, start=1):
            bg = ("gray92", "gray22") if row_idx % 2 == 0 else ("gray96", "gray18")
            for col, (val, width) in enumerate(zip(row_data, self._col_widths)):
                lbl = ctk.CTkLabel(
                    self,
                    text=str(val),
                    width=width,
                    anchor="w",
                    fg_color=bg,
                    corner_radius=2,
                )
                lbl.grid(row=row_idx, column=col, padx=(4, 8), pady=1, sticky="w")


class ResultsTab(ctk.CTkFrame):
    def __init__(self, master, app, **kwargs):
        super().__init__(master, fg_color="transparent", **kwargs)
        self._app = app
        self._matched: list[MatchedOrder] = []
        self._unmatched_inv: list[InvoiceItem] = []
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

        # Matched Orders table
        MATCHED_COLS = ["Platform", "Order No.", "Customer", "Date", "SKU", "Description", "Qty", "Notes"]
        MATCHED_WIDTHS = [70, 100, 130, 90, 130, 250, 40, 280]
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

        # Unmatched orders label (simple, since we don't track the full list currently)
        self._unmatched_orders_note = ctk.CTkLabel(
            self._inner_tabs.tab("Unmatched Orders"),
            text=(
                "These are 'on PO' orders (paid, undispatched, within date range)\n"
                "whose SKUs did not match any item in the imported invoices.\n\n"
                "This may indicate the ordered stock is arriving in a future delivery,\n"
                "or was purchased via phone/counter (not via an online channel)."
            ),
            font=ctk.CTkFont(size=13),
            justify="left",
        )
        self._unmatched_orders_note.pack(padx=20, pady=20, anchor="nw")

        self._unmatched_orders_table = ReadOnlyTable(
            self._inner_tabs.tab("Unmatched Orders"),
            columns=["Platform", "Order No.", "Customer", "Date", "SKU", "Notes"],
            col_widths=[70, 100, 130, 90, 130, 350],
            corner_radius=4,
        )
        self._unmatched_orders_table.pack(fill="both", expand=True, padx=0, pady=(0, 0))

        # ── Bottom row: export ────────────────────────────────────────────
        bottom = ctk.CTkFrame(self, fg_color="transparent")
        bottom.pack(fill="x", padx=12, pady=(4, 12))

        self._export_btn = ctk.CTkButton(
            bottom,
            text="Export to CSV",
            width=140,
            command=self._export_csv,
        )
        self._export_btn.pack(side="left")

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
        invoice_items = self._app.invoice_tab.get_invoice_items()
        neto_orders = self._app.neto_orders
        ebay_orders = self._app.ebay_orders

        matched, unmatched_inv = match_orders_to_invoice(
            invoice_items,
            neto_orders,
            ebay_orders,
            on_po_phrase=self._app.config.app.on_po_filter_phrase,
        )

        self._matched = matched
        self._unmatched_inv = unmatched_inv
        self._app.matched_orders = matched

        self._populate_matched(matched)
        self._populate_unmatched_inv(unmatched_inv)
        self._populate_unmatched_orders(neto_orders, ebay_orders, matched)
        self._update_summary(matched, unmatched_inv)

    def _update_summary(self, matched, unmatched_inv):
        on_po_count = self._count_on_po_orders()
        unmatched_order_count = max(0, on_po_count - len({m.order_id + m.platform for m in matched}))

        self._matched_lbl.configure(
            text=f"Matched: {len(matched)} order line{'s' if len(matched) != 1 else ''}",
            text_color=("green" if matched else "gray50"),
        )
        self._unmatched_inv_lbl.configure(
            text=f"Unmatched invoice items: {len(unmatched_inv)}"
        )
        self._unmatched_orders_lbl.configure(
            text=f"'On PO' orders with no match: {unmatched_order_count}"
        )

    def _count_on_po_orders(self) -> int:
        from src.data_processor import filter_on_po
        phrase = self._app.config.app.on_po_filter_phrase
        return len(filter_on_po(self._app.neto_orders, phrase)) + \
               len(filter_on_po(self._app.ebay_orders, phrase))

    def _populate_matched(self, matched: list[MatchedOrder]):
        rows = []
        for m in matched:
            date_str = m.order_date.strftime("%Y-%m-%d") if m.order_date else ""
            rows.append([
                m.platform,
                m.order_id,
                m.customer_name,
                date_str,
                m.sku,
                m.description,
                str(m.quantity),
                m.notes,
            ])
        self._matched_table.load_rows(rows)

    def _populate_unmatched_inv(self, items: list[InvoiceItem]):
        rows = [[item.sku_with_suffix, item.description, str(item.quantity)] for item in items]
        self._inv_table.load_rows(rows)

    def _populate_unmatched_orders(self, neto_orders, ebay_orders, matched):
        from src.data_processor import filter_on_po
        from src.neto_client import NetoOrder
        from src.ebay_client import EbayOrder

        phrase = self._app.config.app.on_po_filter_phrase
        matched_ids = {(m.platform, m.order_id) for m in matched}

        rows = []
        for order in filter_on_po(neto_orders, phrase):
            channel = order.sales_channel or "Neto"
            if (channel, order.order_id) not in matched_ids:
                date_str = ""
                if order.date_paid:
                    date_str = order.date_paid.strftime("%Y-%m-%d")
                skus = ", ".join(l.sku for l in order.line_items if l.sku)
                rows.append([channel, order.order_id, order.customer_name, date_str, skus, order.notes])

        for order in filter_on_po(ebay_orders, phrase):
            if ("eBay", order.order_id) not in matched_ids:
                date_str = order.creation_date.strftime("%Y-%m-%d") if order.creation_date else ""
                skus = ", ".join(l.sku for l in order.line_items if l.sku)
                rows.append(["eBay", order.order_id, order.buyer_name, date_str, skus, order.buyer_notes])

        self._unmatched_orders_table.load_rows(rows)

    # ── Export ────────────────────────────────────────────────────────────

    def _export_csv(self):
        if not self._matched:
            messagebox.showinfo("No Data", "There are no matched orders to export.")
            return
        try:
            path = export_to_csv(
                self._matched,
                output_dir=self._app.config.app.output_dir,
            )
            self._export_label.configure(text=f"Saved: {path}", text_color="green")
            self._error_label.configure(text="")
            # Open the output folder in Explorer
            subprocess.Popen(["explorer", "/select,", path])
        except Exception as e:
            self._error_label.configure(text=f"Export failed: {e}")
            self._export_label.configure(text="")
