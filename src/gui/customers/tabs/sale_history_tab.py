from __future__ import annotations

import calendar
import threading
import tkinter as tk
from datetime import date, datetime
from tkinter import messagebox, ttk
from typing import Callable, Optional

import customtkinter as ctk


_COL_DEFS = [
    ("tx", "TX # / SKU", 130, "w", False),
    ("detail", "Date & Time / Desc", 220, "w", True),
    ("customer", "Customer / Qty", 120, "w", False),
    ("user", "User / RRP", 120, "w", False),
    ("payment", "Payment / Disc $", 170, "w", False),
    ("lt", "Line Total", 110, "e", False),
    ("cost", "Cost", 95, "e", False),
    ("marg_d", "Margin $", 95, "e", False),
    ("marg_p", "Margin %", 85, "e", False),
    ("total", "Total", 100, "e", False),
]
_COLS = tuple(name for name, *_ in _COL_DEFS)
_BLANK_VALUES = ("",) * len(_COLS)


class SaleHistoryTab(ctk.CTkFrame):
    """Shows completed customer transactions with expandable line detail."""

    _PAGE_SIZE = 100

    def __init__(self, parent, on_refund: Optional[Callable[[dict], None]] = None, **kwargs):
        super().__init__(parent, fg_color="transparent", **kwargs)
        self._customer_id: Optional[str] = None
        self._page = 0
        self._total = 0
        self._loaded = False
        self._rows: list[dict] = []
        self._selected_tx: Optional[dict] = None
        self._all_expanded = False
        self._on_refund_cb = on_refund
        self._date_from: Optional[date] = None
        self._date_to: Optional[date] = None

        self._build_ui()

    def _build_ui(self):
        toolbar = ctk.CTkFrame(self, fg_color="transparent")
        toolbar.pack(fill="x", padx=8, pady=(8, 2))

        self._btn_reprint = ctk.CTkButton(
            toolbar,
            text="Reprint Receipt",
            width=130,
            height=28,
            font=ctk.CTkFont(size=12),
            state="disabled",
            command=self._reprint_receipt,
        )
        self._btn_reprint.pack(side="left")

        self._btn_refund = ctk.CTkButton(
            toolbar,
            text="Refund in Till",
            width=120,
            height=28,
            font=ctk.CTkFont(size=12),
            state="disabled",
            fg_color=("#9b1c1c", "#7f1d1d"),
            hover_color=("#7f1d1d", "#5c1212"),
            command=self._do_refund,
        )
        self._btn_refund.pack(side="left", padx=(8, 0))

        ctk.CTkLabel(
            toolbar,
            text="From",
            font=ctk.CTkFont(size=11),
            text_color=("gray35", "gray70"),
        ).pack(side="left", padx=(16, 4))

        self._from_var = ctk.StringVar()
        self._from_entry = ctk.CTkEntry(
            toolbar,
            textvariable=self._from_var,
            placeholder_text="DD/MM/YYYY",
            width=105,
            height=28,
            font=ctk.CTkFont(size=11),
        )
        self._from_entry.pack(side="left")
        self._from_entry.bind("<Return>", lambda _event: self._apply_date_filter())

        self._btn_from_calendar = ctk.CTkButton(
            toolbar,
            text="...",
            width=30,
            height=28,
            font=ctk.CTkFont(size=11),
            command=lambda: self._open_calendar(self._from_entry, self._from_var),
        )
        self._btn_from_calendar.pack(side="left", padx=(3, 0))

        ctk.CTkLabel(
            toolbar,
            text="To",
            font=ctk.CTkFont(size=11),
            text_color=("gray35", "gray70"),
        ).pack(side="left", padx=(8, 4))

        self._to_var = ctk.StringVar()
        self._to_entry = ctk.CTkEntry(
            toolbar,
            textvariable=self._to_var,
            placeholder_text="DD/MM/YYYY",
            width=105,
            height=28,
            font=ctk.CTkFont(size=11),
        )
        self._to_entry.pack(side="left")
        self._to_entry.bind("<Return>", lambda _event: self._apply_date_filter())

        self._btn_to_calendar = ctk.CTkButton(
            toolbar,
            text="...",
            width=30,
            height=28,
            font=ctk.CTkFont(size=11),
            command=lambda: self._open_calendar(self._to_entry, self._to_var),
        )
        self._btn_to_calendar.pack(side="left", padx=(3, 0))

        self._btn_apply_dates = ctk.CTkButton(
            toolbar,
            text="Apply",
            width=64,
            height=28,
            font=ctk.CTkFont(size=11),
            command=self._apply_date_filter,
        )
        self._btn_apply_dates.pack(side="left", padx=(8, 0))

        self._btn_clear_dates = ctk.CTkButton(
            toolbar,
            text="Clear",
            width=64,
            height=28,
            font=ctk.CTkFont(size=11),
            fg_color="transparent",
            border_width=1,
            border_color=("gray60", "gray45"),
            text_color=("gray30", "gray70"),
            hover_color=("gray85", "gray25"),
            command=self._clear_date_filter,
        )
        self._btn_clear_dates.pack(side="left", padx=(6, 0))

        self._lbl_status = ctk.CTkLabel(
            toolbar,
            text="",
            font=ctk.CTkFont(size=11),
            text_color=("gray50", "gray60"),
        )
        self._lbl_status.pack(side="right")

        tree_frame = ctk.CTkFrame(self, fg_color="transparent")
        tree_frame.pack(fill="both", expand=True, padx=8, pady=(0, 0))
        tree_frame.grid_rowconfigure(0, weight=1)
        tree_frame.grid_columnconfigure(0, weight=1)

        self._tree = ttk.Treeview(
            tree_frame,
            columns=_COLS,
            show="tree headings",
            selectmode="browse",
            style="Sales.Treeview",
        )

        self._tree.column("#0", width=16, minwidth=16, stretch=False)
        self._tree.heading("#0", text="", command=self._toggle_expand_all)

        for name, heading, width, anchor, stretch in _COL_DEFS:
            self._tree.column(name, width=width, minwidth=60, anchor=anchor, stretch=stretch)
            self._tree.heading(name, text=heading, anchor=anchor)

        vsb = ttk.Scrollbar(tree_frame, orient="vertical", command=self._tree.yview)
        hsb = ttk.Scrollbar(tree_frame, orient="horizontal", command=self._tree.xview)
        self._tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        self._tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")

        self._tree.tag_configure("tx_row", font=("Segoe UI", 10, "bold"))
        self._tree.tag_configure(
            "refund_tx_row",
            font=("Segoe UI", 10, "bold"),
            foreground="#e74c3c",
        )
        self._tree.tag_configure("line_row", font=("Segoe UI", 10))
        self._tree.tag_configure(
            "sum_row",
            font=("Segoe UI", 10, "bold"),
            foreground="#22c55e",
        )
        self._tree.tag_configure(
            "div_row",
            background="#2a2a2a",
            foreground="#2a2a2a",
        )

        self._tree.bind("<<TreeviewSelect>>", self._on_tree_select)
        self._tree.bind("<Button-3>", self._on_right_click)
        self._tree.bind("<<TreeviewOpen>>", self._on_tree_open)
        self._tree.bind("<<TreeviewClose>>", self._on_tree_close)

        pag = ctk.CTkFrame(self, fg_color="transparent")
        pag.pack(fill="x", padx=8, pady=4)

        self._btn_prev = ctk.CTkButton(
            pag,
            text="Prev",
            width=80,
            height=26,
            font=ctk.CTkFont(size=11),
            command=self._prev_page,
        )
        self._btn_prev.pack(side="left")

        self._lbl_page = ctk.CTkLabel(pag, text="", font=ctk.CTkFont(size=11))
        self._lbl_page.pack(side="left", padx=12)

        self._btn_next = ctk.CTkButton(
            pag,
            text="Next",
            width=80,
            height=26,
            font=ctk.CTkFont(size=11),
            command=self._next_page,
        )
        self._btn_next.pack(side="left")

    def load_for_customer(self, customer_uuid: str):
        self._customer_id = customer_uuid
        self._page = 0
        self._date_from = None
        self._date_to = None
        self._from_var.set("")
        self._to_var.set("")
        self._loaded = True
        self._fetch()

    def on_tab_selected(self):
        if not self._loaded and self._customer_id:
            self._fetch()

    def clear(self):
        self._customer_id = None
        self._loaded = False
        self._page = 0
        self._total = 0
        self._date_from = None
        self._date_to = None
        self._from_var.set("")
        self._to_var.set("")
        self._rows = []
        self._selected_tx = None
        self._all_expanded = False
        self._tree.delete(*self._tree.get_children())
        self._lbl_page.configure(text="")
        self._lbl_status.configure(text="")
        self._btn_prev.configure(state="disabled")
        self._btn_next.configure(state="disabled")
        self._btn_reprint.configure(state="disabled")
        self._btn_refund.configure(state="disabled")

    def _fetch(self):
        if not self._customer_id:
            return
        self._lbl_status.configure(text="Loading...", text_color=("gray50", "gray60"))
        uuid = self._customer_id
        page = self._page
        date_from = self._date_from
        date_to = self._date_to

        def _thread():
            from src.customers.customer_client import get_customer_transactions
            try:
                rows, total = get_customer_transactions(
                    uuid,
                    page=page,
                    page_size=self._PAGE_SIZE,
                    date_from=date_from,
                    date_to=date_to,
                )
                self.after(0, lambda: self._populate(rows, total))
            except Exception as exc:
                err = str(exc)
                self.after(
                    0,
                    lambda: self._lbl_status.configure(
                        text=f"Error: {err}",
                        text_color=("red", "#e74c3c"),
                    ),
                )

        threading.Thread(target=_thread, daemon=True).start()

    def _populate(self, rows: list, total: int):
        self._rows = rows
        self._total = total
        self._selected_tx = None
        self._all_expanded = False
        self._btn_reprint.configure(state="disabled")
        self._btn_refund.configure(state="disabled")
        self._tree.delete(*self._tree.get_children())

        for tx in rows:
            self._insert_transaction(tx)

        self._lbl_page.configure(
            text=f"Page {self._page + 1}" if self._page > 0 or total > 0 else ""
        )
        self._lbl_status.configure(
            text=f"{len(rows)} transaction(s)" if rows else "No transactions",
            text_color=("gray50", "gray60"),
        )
        self._btn_prev.configure(state="normal" if self._page > 0 else "disabled")
        self._btn_next.configure(
            state="normal" if (self._page + 1) * self._PAGE_SIZE < total else "disabled"
        )

    def _apply_date_filter(self):
        try:
            date_from = _parse_date(self._from_var.get())
            date_to = _parse_date(self._to_var.get())
        except ValueError as exc:
            messagebox.showwarning(
                "Invalid Date",
                str(exc),
                parent=self.winfo_toplevel(),
            )
            return

        if date_from and date_to and date_from > date_to:
            messagebox.showwarning(
                "Invalid Date Range",
                "The From date must be on or before the To date.",
                parent=self.winfo_toplevel(),
            )
            return

        self._date_from = date_from
        self._date_to = date_to
        self._page = 0
        self._fetch()

    def _clear_date_filter(self):
        self._from_var.set("")
        self._to_var.set("")
        self._date_from = None
        self._date_to = None
        self._page = 0
        self._fetch()

    def _open_calendar(self, anchor: tk.Widget, target_var: tk.StringVar):
        try:
            selected = _parse_date(target_var.get())
        except ValueError:
            selected = date.today()
        if selected is None:
            selected = date.today()

        _CalendarPopup(
            self,
            anchor=anchor,
            initial_date=selected,
            on_select=lambda chosen: target_var.set(_format_date(chosen)),
        )

    def _insert_transaction(self, tx: dict):
        tx_id = tx["id"]
        tx_num_raw = tx.get("transaction_number") or ""
        total = float(tx.get("total") or 0)
        lines = tx.get("transaction_lines") or []
        is_refund = (tx.get("sale_type") or "").lower() == "refund"
        is_refunded = bool(tx.get("is_refunded"))
        parent_tag = "refund_tx_row" if is_refund else "tx_row"
        tx_num = f"{tx_num_raw}  [Refunded]" if is_refunded else tx_num_raw

        parts: list[str] = []
        cash = float(tx.get("payment_cash") or 0)
        if cash:
            parts.append(f"Cash ${cash:,.2f}")
        eft_total = sum(float(e.get("amount") or 0) for e in (tx.get("payment_eft") or []))
        if eft_total:
            parts.append(f"EFT ${eft_total:,.2f}")
        online = float(tx.get("payment_online") or 0)
        if online:
            parts.append(f"Online ${online:,.2f}")
        payment_str = "  /  ".join(parts) if parts else "-"

        self._tree.insert(
            "",
            "end",
            iid=tx_id,
            text="",
            values=(
                tx_num,
                _fmt_dt(tx.get("completed_at", "")),
                tx.get("customer_name") or tx.get("park_name") or "",
                tx.get("performed_by") or "",
                payment_str,
                "",
                "",
                "",
                "",
                f"${total:,.2f}",
            ),
            tags=(parent_tag,),
            open=False,
        )

        if not lines:
            return

        total_qty = total_rrp = total_cost = 0.0
        for line in lines:
            qty = float(line.get("qty") or 0)
            unit_price = float(line.get("unit_price") or 0)
            cost_price = float(line.get("cost_price") or 0)
            line_total = float(line.get("line_total") or 0)

            rrp_line = unit_price * qty
            disc_line = rrp_line - line_total
            cost_line = cost_price * qty
            sell_ex = line_total / 1.1
            margin_d = (sell_ex - cost_line) if cost_price else None
            margin_p = (margin_d / sell_ex * 100) if (margin_d is not None and sell_ex > 0) else None

            total_qty += qty
            total_rrp += rrp_line
            total_cost += cost_line

            self._tree.insert(
                tx_id,
                "end",
                text="",
                values=(
                    line.get("sku") or "",
                    line.get("description") or "",
                    f"{qty:g}",
                    f"${unit_price:,.2f}",
                    f"${disc_line:,.2f}" if disc_line > 0.005 else "-",
                    f"${line_total:,.2f}",
                    f"${cost_line:,.2f}" if cost_price else "-",
                    f"${margin_d:,.2f}" if margin_d is not None else "-",
                    f"{margin_p:.1f}%" if margin_p is not None else "-",
                    "",
                ),
                tags=("line_row",),
            )

        total_disc = total_rrp - total
        total_ex_gst = total / 1.1
        total_margin_d = (total_ex_gst - total_cost) if total_cost else None
        total_margin_p = (
            total_margin_d / total_ex_gst * 100
            if (total_margin_d is not None and total_ex_gst > 0)
            else None
        )

        self._tree.insert(
            tx_id,
            "end",
            text="",
            values=(
                "TOTAL",
                "",
                f"{total_qty:g}",
                "",
                f"${total_disc:,.2f}" if total_disc > 0.005 else "-",
                f"${total:,.2f}",
                f"${total_cost:,.2f}" if total_cost else "-",
                f"${total_margin_d:,.2f}" if total_margin_d is not None else "-",
                f"{total_margin_p:.1f}%" if total_margin_p is not None else "-",
                "",
            ),
            tags=("sum_row",),
        )

    def _toggle_expand_all(self):
        tx_ids = {r["id"] for r in self._rows}
        if not self._all_expanded:
            top = [iid for iid in self._tree.get_children("") if iid in tx_ids]
            for iid in top:
                self._tree.item(iid, open=True)
            for iid in reversed(top):
                self._insert_divider_after(iid, tx_ids)
            self._all_expanded = True
        else:
            for iid in list(self._tree.get_children("")):
                if _is_divider(iid):
                    self._tree.delete(iid)
            for iid in self._tree.get_children(""):
                self._tree.item(iid, open=False)
            self._all_expanded = False

    def _on_tree_open(self, _event=None):
        iid = self._tree.focus()
        if not iid:
            return
        tx_ids = {r["id"] for r in self._rows}
        if iid not in tx_ids:
            return
        self._insert_divider_after(iid, tx_ids)

    def _on_tree_close(self, _event=None):
        iid = self._tree.focus()
        if not iid:
            return
        sep_iid = _sep_id(iid)
        if self._tree.exists(sep_iid):
            self._tree.delete(sep_iid)

    def _insert_divider_after(self, tx_iid: str, tx_ids: set):
        sep_iid = _sep_id(tx_iid)
        if self._tree.exists(sep_iid):
            return
        children = list(self._tree.get_children(""))
        try:
            idx = children.index(tx_iid)
        except ValueError:
            return
        self._tree.insert("", idx + 1, iid=sep_iid, text="", values=_BLANK_VALUES, tags=("div_row",))

    def _on_tree_select(self, _event=None):
        sel = self._tree.selection()
        if not sel:
            self._selected_tx = None
            self._btn_reprint.configure(state="disabled")
            self._btn_refund.configure(state="disabled")
            return
        tx = self._resolve_tx(sel[0])
        self._selected_tx = tx
        self._btn_reprint.configure(state="normal" if tx else "disabled")
        self._btn_refund.configure(state="normal" if self._can_refund(tx) else "disabled")

    def _on_right_click(self, event):
        iid = self._tree.identify_row(event.y)
        if not iid:
            return
        tx = self._resolve_tx(iid)
        if tx is None:
            return
        self._tree.selection_set(iid)
        self._selected_tx = tx
        self._btn_reprint.configure(state="normal")
        can_refund = self._can_refund(tx)
        self._btn_refund.configure(state="normal" if can_refund else "disabled")

        menu = tk.Menu(self, tearoff=0)
        menu.add_command(label="Reprint Receipt", command=self._reprint_receipt)
        if can_refund:
            menu.add_command(label="Refund in Till", command=self._do_refund)
        menu.tk_popup(event.x_root, event.y_root)

    def _resolve_tx(self, iid: str) -> Optional[dict]:
        if _is_divider(iid):
            return None
        tx_ids = {r["id"] for r in self._rows}
        if iid in tx_ids:
            return next((r for r in self._rows if r["id"] == iid), None)
        parent_iid = self._tree.parent(iid)
        if parent_iid in tx_ids:
            return next((r for r in self._rows if r["id"] == parent_iid), None)
        return None

    def _can_refund(self, tx: Optional[dict]) -> bool:
        if not tx or not self._on_refund_cb:
            return False
        if (tx.get("sale_type") or "").lower() == "refund":
            return False
        if tx.get("is_refunded"):
            return False
        return True

    def _do_refund(self):
        tx = self._selected_tx
        if not self._can_refund(tx):
            return
        self._on_refund_cb(tx)

    def _reprint_receipt(self):
        tx = self._selected_tx
        if not tx:
            return

        from src.config import config

        printer = getattr(getattr(config, "device", None), "receipt_printer", None)
        if not printer:
            messagebox.showwarning(
                "No Printer",
                "No receipt printer is configured.\nGo to Settings -> Printers to set one up.",
                parent=self.winfo_toplevel(),
            )
            return

        lines = tx.get("transaction_lines") or []
        cart_items = {
            (line.get("item_id") or str(i)): {
                "sku": line.get("sku") or "",
                "title": line.get("description") or "",
                "qty": float(line.get("qty") or 1),
                "unit_price": float(line.get("unit_price") or 0),
                "disc_pct": float(line.get("discount_pct") or 0),
                "cost_price": float(line.get("cost_price") or 0),
            }
            for i, line in enumerate(lines)
        }

        def _thread():
            try:
                from src.pos.receipt_generator import generate_receipt
                from src.printer_utils import print_pdf

                customer = None
                cust_uuid = tx.get("customer_id")
                if cust_uuid:
                    from src.customers.customer_client import get_customer

                    customer = get_customer(cust_uuid)
                pdf_path = generate_receipt(tx, cart_items, customer)
                print_pdf(pdf_path, printer)
            except Exception as exc:
                err = str(exc)
                self.after(
                    0,
                    lambda: messagebox.showerror(
                        "Print Failed",
                        f"Could not print receipt:\n{err}",
                        parent=self.winfo_toplevel(),
                    ),
                )

        threading.Thread(target=_thread, daemon=True).start()

    def _prev_page(self):
        if self._page > 0:
            self._page -= 1
            self._fetch()

    def _next_page(self):
        if (self._page + 1) * self._PAGE_SIZE < self._total:
            self._page += 1
            self._fetch()


def _sep_id(tx_iid: str) -> str:
    return f"__sep_{tx_iid}"


def _is_divider(iid: str) -> bool:
    return iid.startswith("__sep_")


def _fmt_dt(iso: str) -> str:
    try:
        import re
        from datetime import datetime
        from zoneinfo import ZoneInfo

        melb = ZoneInfo("Australia/Melbourne")
        s = iso.replace("Z", "+00:00")
        s = re.sub(r"\.(\d{1,6})", lambda m: "." + m.group(1).ljust(6, "0"), s)
        dt = datetime.fromisoformat(s)
        dt_local = dt.astimezone(melb)
        hour = dt_local.hour % 12 or 12
        am_pm = "AM" if dt_local.hour < 12 else "PM"
        return (
            f"{dt_local.day:02d}-{dt_local.month:02d}-{dt_local.year}  "
            f"{hour}:{dt_local.minute:02d} {am_pm}"
        )
    except Exception:
        return iso[:16] if len(iso) >= 16 else iso


class _CalendarPopup(ctk.CTkToplevel):
    def __init__(
        self,
        parent,
        anchor: tk.Widget,
        initial_date: date,
        on_select: Callable[[date], None],
    ):
        super().__init__(parent.winfo_toplevel())
        self.withdraw()
        self.title("Select Date")
        self.resizable(False, False)
        self.transient(parent.winfo_toplevel())

        self._anchor = anchor
        self._selected = initial_date
        self._display_year = initial_date.year
        self._display_month = initial_date.month
        self._on_select = on_select
        self._day_buttons: list[ctk.CTkButton] = []

        self._build()
        self._render_month()
        self._position()

        self.bind("<Escape>", lambda _event: self.destroy())
        self.protocol("WM_DELETE_WINDOW", self.destroy)
        self.deiconify()
        self.lift()
        self.focus_force()
        try:
            self.grab_set()
        except tk.TclError:
            pass

    def _build(self):
        outer = ctk.CTkFrame(self, fg_color=("white", "gray16"), corner_radius=8)
        outer.pack(fill="both", expand=True, padx=8, pady=8)

        header = ctk.CTkFrame(outer, fg_color="transparent")
        header.pack(fill="x", padx=6, pady=(6, 4))

        ctk.CTkButton(
            header,
            text="<",
            width=32,
            height=28,
            font=ctk.CTkFont(size=13, weight="bold"),
            command=self._prev_month,
        ).pack(side="left")

        self._lbl_month = ctk.CTkLabel(
            header,
            text="",
            width=170,
            font=ctk.CTkFont(size=13, weight="bold"),
        )
        self._lbl_month.pack(side="left", padx=6)

        ctk.CTkButton(
            header,
            text=">",
            width=32,
            height=28,
            font=ctk.CTkFont(size=13, weight="bold"),
            command=self._next_month,
        ).pack(side="left")

        weekday_row = ctk.CTkFrame(outer, fg_color="transparent")
        weekday_row.pack(fill="x", padx=6, pady=(2, 0))
        for col, name in enumerate(calendar.day_abbr):
            label = ctk.CTkLabel(
                weekday_row,
                text=name,
                width=34,
                font=ctk.CTkFont(size=10, weight="bold"),
                text_color=("gray45", "gray60"),
            )
            label.grid(row=0, column=col, padx=1)

        self._days_frame = ctk.CTkFrame(outer, fg_color="transparent")
        self._days_frame.pack(fill="x", padx=6, pady=(2, 6))

        footer = ctk.CTkFrame(outer, fg_color="transparent")
        footer.pack(fill="x", padx=6, pady=(0, 6))
        ctk.CTkButton(
            footer,
            text="Today",
            width=72,
            height=28,
            font=ctk.CTkFont(size=11),
            command=lambda: self._select(date.today()),
        ).pack(side="left")
        ctk.CTkButton(
            footer,
            text="Cancel",
            width=72,
            height=28,
            font=ctk.CTkFont(size=11),
            fg_color="transparent",
            border_width=1,
            border_color=("gray60", "gray45"),
            text_color=("gray30", "gray70"),
            hover_color=("gray85", "gray25"),
            command=self.destroy,
        ).pack(side="right")

    def _position(self):
        self.update_idletasks()
        x = self._anchor.winfo_rootx()
        y = self._anchor.winfo_rooty() + self._anchor.winfo_height() + 2
        self.geometry(f"+{x}+{y}")

    def _render_month(self):
        self._lbl_month.configure(
            text=f"{calendar.month_name[self._display_month]} {self._display_year}"
        )

        for widget in self._days_frame.winfo_children():
            widget.destroy()
        self._day_buttons = []

        month = calendar.Calendar(firstweekday=0)
        for row, week in enumerate(month.monthdayscalendar(self._display_year, self._display_month)):
            for col, day_num in enumerate(week):
                if day_num == 0:
                    ctk.CTkLabel(self._days_frame, text="", width=34, height=28).grid(
                        row=row,
                        column=col,
                        padx=1,
                        pady=1,
                    )
                    continue

                day = date(self._display_year, self._display_month, day_num)
                selected = day == self._selected
                button = ctk.CTkButton(
                    self._days_frame,
                    text=str(day_num),
                    width=34,
                    height=28,
                    font=ctk.CTkFont(size=11),
                    fg_color=("#1f6aa5", "#1f6aa5") if selected else "transparent",
                    border_width=0 if selected else 1,
                    border_color=("gray75", "gray35"),
                    text_color="white" if selected else ("gray20", "gray80"),
                    hover_color=("#d9e8ff", "#26374f"),
                    command=lambda chosen=day: self._select(chosen),
                )
                button.grid(row=row, column=col, padx=1, pady=1)
                self._day_buttons.append(button)

    def _prev_month(self):
        if self._display_month == 1:
            self._display_month = 12
            self._display_year -= 1
        else:
            self._display_month -= 1
        self._render_month()

    def _next_month(self):
        if self._display_month == 12:
            self._display_month = 1
            self._display_year += 1
        else:
            self._display_month += 1
        self._render_month()

    def _select(self, chosen: date):
        self._on_select(chosen)
        self.destroy()


def _parse_date(value: str) -> Optional[date]:
    value = (value or "").strip()
    if not value:
        return None

    for fmt in ("%d/%m/%Y", "%d-%m-%Y", "%Y-%m-%d"):
        try:
            return datetime.strptime(value, fmt).date()
        except ValueError:
            pass

    raise ValueError("Enter dates as DD/MM/YYYY.")


def _format_date(value: date) -> str:
    return value.strftime("%d/%m/%Y")
