from __future__ import annotations

import threading
from tkinter import ttk
from typing import Optional

import customtkinter as ctk


class SaleHistoryTab(ctk.CTkFrame):
    """Shows completed transactions linked to a customer."""

    _PAGE_SIZE = 50

    def __init__(self, parent, **kwargs):
        super().__init__(parent, fg_color="transparent", **kwargs)
        self._customer_id: Optional[str] = None
        self._page = 0
        self._total = 0
        self._loaded = False

        self._build_ui()

    # ── Build ────────────────────────────────────────────────────────────

    def _build_ui(self):
        # Treeview
        tree_frame = ctk.CTkFrame(self, fg_color="transparent")
        tree_frame.pack(fill="both", expand=True, padx=8, pady=(8, 0))

        columns = ("date", "tx_num", "type", "total")
        self._tree = ttk.Treeview(
            tree_frame,
            columns=columns,
            show="headings",
            style="Cust.Treeview",
            selectmode="browse",
        )
        vsb = ttk.Scrollbar(tree_frame, orient="vertical", command=self._tree.yview)
        self._tree.configure(yscrollcommand=vsb.set)

        self._tree.heading("date",   text="Date",    anchor="w")
        self._tree.heading("tx_num", text="Tx #",    anchor="w")
        self._tree.heading("type",   text="Type",    anchor="w")
        self._tree.heading("total",  text="Total",   anchor="e")

        self._tree.column("date",   width=120, stretch=False, anchor="w")
        self._tree.column("tx_num", width=130, stretch=False, anchor="w")
        self._tree.column("type",   width=100, stretch=False, anchor="w")
        self._tree.column("total",  width=90,  stretch=False, anchor="e")

        self._tree.pack(side="left", fill="both", expand=True)
        vsb.pack(side="right", fill="y")

        # Pagination bar
        pag = ctk.CTkFrame(self, fg_color="transparent")
        pag.pack(fill="x", padx=8, pady=4)

        self._btn_prev = ctk.CTkButton(
            pag, text="◀ Prev", width=80, height=26,
            font=ctk.CTkFont(size=11),
            command=self._prev_page,
        )
        self._btn_prev.pack(side="left")

        self._lbl_page = ctk.CTkLabel(
            pag, text="", font=ctk.CTkFont(size=11),
        )
        self._lbl_page.pack(side="left", padx=12)

        self._btn_next = ctk.CTkButton(
            pag, text="Next ▶", width=80, height=26,
            font=ctk.CTkFont(size=11),
            command=self._next_page,
        )
        self._btn_next.pack(side="left")

        self._lbl_status = ctk.CTkLabel(
            pag, text="", font=ctk.CTkFont(size=11),
            text_color=("gray50", "gray60"),
        )
        self._lbl_status.pack(side="right", padx=8)

    # ── Public API ───────────────────────────────────────────────────────

    def load_for_customer(self, customer_uuid: str):
        """Set the customer and load the first page."""
        self._customer_id = customer_uuid
        self._page = 0
        self._loaded = True
        self._fetch()

    def on_tab_selected(self):
        """Called when the tab is first selected — lazy load."""
        if not self._loaded and self._customer_id:
            self._fetch()

    def clear(self):
        self._customer_id = None
        self._loaded = False
        self._page = 0
        self._total = 0
        self._tree.delete(*self._tree.get_children())
        self._lbl_page.configure(text="")
        self._lbl_status.configure(text="")
        self._btn_prev.configure(state="disabled")
        self._btn_next.configure(state="disabled")

    # ── Internal ─────────────────────────────────────────────────────────

    def _fetch(self):
        if not self._customer_id:
            return
        self._lbl_status.configure(text="Loading…")
        uuid = self._customer_id
        page = self._page

        def _thread():
            from src.customers.customer_client import get_customer_transactions
            try:
                rows, total = get_customer_transactions(
                    uuid, page=page, page_size=self._PAGE_SIZE
                )
                self.after(0, lambda: self._populate(rows, total))
            except Exception as exc:
                err = str(exc)
                self.after(0, lambda: self._lbl_status.configure(
                    text=f"Error: {err}", text_color=("red", "#e74c3c")
                ))

        threading.Thread(target=_thread, daemon=True).start()

    def _populate(self, rows: list, total: int):
        self._total = total
        self._tree.delete(*self._tree.get_children())

        for r in rows:
            date_raw = r.get("completed_at") or ""
            date_str = date_raw[:10] if date_raw else "—"
            sale_type = (r.get("sale_type") or "standard").capitalize()
            total_val = r.get("total")
            total_str = f"${total_val:,.2f}" if total_val is not None else "—"
            self._tree.insert("", "end", values=(
                date_str,
                r.get("transaction_number") or "—",
                sale_type,
                total_str,
            ))

        self._lbl_page.configure(
            text=f"Page {self._page + 1}" if self._page > 0 or total > 0 else ""
        )
        self._lbl_status.configure(text=f"{len(rows)} transaction(s)" if rows else "No transactions")
        self._btn_prev.configure(state="normal" if self._page > 0 else "disabled")
        self._btn_next.configure(
            state="normal" if (self._page + 1) * self._PAGE_SIZE < total else "disabled"
        )

    def _prev_page(self):
        if self._page > 0:
            self._page -= 1
            self._fetch()

    def _next_page(self):
        if (self._page + 1) * self._PAGE_SIZE < self._total:
            self._page += 1
            self._fetch()
