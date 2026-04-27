from __future__ import annotations

import threading
from datetime import datetime
from tkinter import messagebox, ttk
from typing import Optional
from zoneinfo import ZoneInfo

import customtkinter as ctk


class QuotesTab(ctk.CTkFrame):
    """Customer-scoped quote transaction list."""

    _PAGE_SIZE = 100

    def __init__(self, parent, **kwargs):
        super().__init__(parent, fg_color="transparent", **kwargs)
        self._customer: Optional[dict] = None
        self._customer_id: Optional[str] = None
        self._page = 0
        self._total = 0
        self._rows: list[dict] = []
        self._selected_tx: Optional[dict] = None
        self._loaded = False
        self._build_ui()

    def _build_ui(self) -> None:
        toolbar = ctk.CTkFrame(self, fg_color="transparent")
        toolbar.pack(fill="x", padx=8, pady=(8, 2))

        self._btn_preview = ctk.CTkButton(
            toolbar,
            text="Preview Quote",
            width=120,
            height=28,
            font=ctk.CTkFont(size=12),
            state="disabled",
            command=self._preview_quote,
        )
        self._btn_preview.pack(side="left")

        self._btn_refresh = ctk.CTkButton(
            toolbar,
            text="Refresh",
            width=80,
            height=28,
            font=ctk.CTkFont(size=12),
            command=self.refresh,
        )
        self._btn_refresh.pack(side="left", padx=(8, 0))

        self._status_lbl = ctk.CTkLabel(
            toolbar,
            text="No customer selected",
            font=ctk.CTkFont(size=11),
            text_color=("gray45", "gray60"),
        )
        self._status_lbl.pack(side="left", padx=(12, 0))

        nav = ctk.CTkFrame(toolbar, fg_color="transparent")
        nav.pack(side="right")
        self._prev_btn = ctk.CTkButton(nav, text="< Prev", width=70, height=26, command=self._prev_page)
        self._prev_btn.pack(side="left", padx=(0, 6))
        self._page_lbl = ctk.CTkLabel(nav, text="Page 1", font=ctk.CTkFont(size=11))
        self._page_lbl.pack(side="left", padx=4)
        self._next_btn = ctk.CTkButton(nav, text="Next >", width=70, height=26, command=self._next_page)
        self._next_btn.pack(side="left", padx=(6, 0))

        frame = ctk.CTkFrame(self, fg_color="transparent")
        frame.pack(fill="both", expand=True, padx=8, pady=(4, 8))

        cols = ("quote", "date", "items", "total", "status", "user")
        self._tree = ttk.Treeview(frame, columns=cols, show="headings", selectmode="browse")
        self._tree.heading("quote", text="Quote #")
        self._tree.heading("date", text="Date")
        self._tree.heading("items", text="Items")
        self._tree.heading("total", text="Total")
        self._tree.heading("status", text="Status")
        self._tree.heading("user", text="User")

        self._tree.column("quote", width=130, anchor="w", stretch=False)
        self._tree.column("date", width=140, anchor="w", stretch=False)
        self._tree.column("items", width=70, anchor="e", stretch=False)
        self._tree.column("total", width=110, anchor="e", stretch=False)
        self._tree.column("status", width=110, anchor="w", stretch=False)
        self._tree.column("user", width=120, anchor="w", stretch=True)

        vsb = ttk.Scrollbar(frame, orient="vertical", command=self._tree.yview)
        self._tree.configure(yscrollcommand=vsb.set)
        self._tree.pack(side="left", fill="both", expand=True)
        vsb.pack(side="right", fill="y")

        self._tree.bind("<<TreeviewSelect>>", self._on_select)
        self._tree.bind("<Double-1>", lambda _e: self._preview_quote())

        self._update_nav()

    def load_for_customer(self, customer: dict | str | None) -> None:
        if isinstance(customer, dict):
            customer_id = customer.get("id")
            if customer_id == self._customer_id:
                self._customer = customer
                return
            self._customer = customer
            self._customer_id = customer_id
        else:
            if customer == self._customer_id:
                return
            self._customer = None
            self._customer_id = customer
        self._page = 0
        self._loaded = False
        self._rows = []
        self._selected_tx = None
        self._tree.delete(*self._tree.get_children())
        self._btn_preview.configure(state="disabled")
        self._status_lbl.configure(text="Quotes not loaded")
        self._update_nav()

    def on_tab_selected(self) -> None:
        if not self._loaded:
            self.refresh()

    def refresh(self) -> None:
        if not self._customer_id:
            self._status_lbl.configure(text="No customer selected")
            return
        self._status_lbl.configure(text="Loading quotes...")
        self._btn_refresh.configure(state="disabled")
        self._btn_preview.configure(state="disabled")
        threading.Thread(target=self._load_thread, daemon=True).start()

    def _load_thread(self) -> None:
        from src.customers.customer_client import get_customer_quotes

        try:
            rows, total = get_customer_quotes(
                self._customer_id,
                page=self._page,
                page_size=self._PAGE_SIZE,
            )
            self.after(0, lambda: self._show_rows(rows, total))
        except Exception as exc:
            err = str(exc)
            self.after(0, lambda: self._show_error(err))

    def _show_rows(self, rows: list[dict], total: int) -> None:
        self._rows = rows
        self._total = total
        self._loaded = True
        self._selected_tx = None
        self._tree.delete(*self._tree.get_children())

        for tx in rows:
            lines = tx.get("transaction_lines") or []
            self._tree.insert("", "end", iid=tx.get("id"), values=(
                _quote_number(tx),
                _fmt_datetime(tx.get("created_at")),
                _fmt_qty(sum(float(line.get("qty") or 0) for line in lines)),
                _fmt_money(float(tx.get("total") or 0)),
                _fmt_status(tx.get("sale_status")),
                tx.get("performed_by") or "",
            ))

        if rows:
            start = self._page * self._PAGE_SIZE + 1
            end = start + len(rows) - 1
            self._status_lbl.configure(text=f"Showing {start}-{end}")
        else:
            self._status_lbl.configure(text="No quotes")
        self._btn_refresh.configure(state="normal")
        self._btn_preview.configure(state="disabled")
        self._update_nav()

    def _show_error(self, err: str) -> None:
        self._btn_refresh.configure(state="normal")
        self._status_lbl.configure(text="Could not load quotes")
        messagebox.showerror(
            "Quotes Error",
            f"Could not load quotes:\n\n{err}",
            parent=self.winfo_toplevel(),
        )

    def _on_select(self, _event=None) -> None:
        sel = self._tree.selection()
        if not sel:
            self._selected_tx = None
            self._btn_preview.configure(state="disabled")
            return
        tx_id = sel[0]
        self._selected_tx = next((row for row in self._rows if row.get("id") == tx_id), None)
        self._btn_preview.configure(state="normal" if self._selected_tx else "disabled")

    def _preview_quote(self) -> None:
        tx = self._selected_tx
        if not tx:
            return
        try:
            from src.pos.customer_document_generator import generate_quote_document
            from src.gui.pos.document_preview_dialog import DocumentPreviewDialog

            cart_items = _lines_to_cart(tx.get("transaction_lines") or [])
            pdf_path = generate_quote_document(tx, cart_items, self._customer, tx.get("notes"))
            DocumentPreviewDialog(
                self.winfo_toplevel(),
                pdf_path=pdf_path,
                title=f"Quote Preview - {_quote_number(tx)}",
                customer=self._customer,
                email_subject=f"Quote {_quote_number(tx)} from Scarlett Music",
            )
        except Exception as exc:
            messagebox.showerror(
                "Preview Failed",
                f"Could not preview quote:\n{exc}",
                parent=self.winfo_toplevel(),
            )

    def _prev_page(self) -> None:
        if self._page <= 0:
            return
        self._page -= 1
        self.refresh()

    def _next_page(self) -> None:
        if (self._page + 1) * self._PAGE_SIZE >= self._total:
            return
        self._page += 1
        self.refresh()

    def _update_nav(self) -> None:
        self._page_lbl.configure(text=f"Page {self._page + 1}")
        self._prev_btn.configure(state="normal" if self._page > 0 else "disabled")
        has_next = (self._page + 1) * self._PAGE_SIZE < self._total
        self._next_btn.configure(state="normal" if has_next else "disabled")


def _quote_number(tx: dict) -> str:
    if tx.get("quote_number"):
        return f"Q-{int(tx['quote_number']):05d}"
    return tx.get("transaction_number") or ""


def _fmt_datetime(raw: str | None) -> str:
    if not raw:
        return ""
    try:
        dt = datetime.fromisoformat(str(raw).replace("Z", "+00:00"))
        if dt.tzinfo is not None:
            dt = dt.astimezone(ZoneInfo("Australia/Melbourne"))
        hour = dt.hour % 12 or 12
        am_pm = "AM" if dt.hour < 12 else "PM"
        return f"{dt.day:02d}/{dt.month:02d}/{dt.year} {hour}:{dt.minute:02d} {am_pm}"
    except Exception:
        return str(raw)[:16]


def _fmt_money(amount: float) -> str:
    return f"${amount:,.2f}"


def _fmt_qty(qty: float) -> str:
    return str(int(qty)) if qty == int(qty) else f"{qty:.3f}".rstrip("0").rstrip(".")


def _fmt_status(status: str | None) -> str:
    status = (status or "").replace("_", " ").strip()
    return status.title() if status else ""


def _lines_to_cart(lines: list[dict]) -> dict:
    cart = {}
    for idx, line in enumerate(lines):
        item_id = line.get("item_id") or str(idx)
        cart[item_id] = {
            "sku": line.get("sku") or "",
            "title": line.get("description") or "",
            "qty": float(line.get("qty") or 0),
            "unit_price": float(line.get("unit_price") or 0),
            "disc_pct": float(line.get("discount_pct") or 0),
            "cost_price": float(line.get("cost_price") or 0) if line.get("cost_price") else None,
        }
    return cart
