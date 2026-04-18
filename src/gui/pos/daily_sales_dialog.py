"""Daily Sales dialog — expandable transaction list for today's completed sales."""
from __future__ import annotations

import threading
import tkinter as tk
from tkinter import messagebox, ttk
from typing import Optional

import customtkinter as ctk


# ── Treeview column definitions ───────────────────────────────────────────────
# Each column is dual-purpose: parent (tx summary) rows use the first meaning,
# child (line item) rows use the second.
#
# name       heading                   width  anchor  stretch
_COL_DEFS = [
    ("tx",       "TX # / SKU",          130,   "w",    False),
    ("detail",   "Date & Time / Desc",  200,   "w",    True),   # stretchable
    ("customer", "Customer / Qty",      120,   "w",    False),
    ("user",     "User / RRP",          120,   "w",    False),
    ("payment",  "Payment / Disc $",    165,   "w",    False),
    ("lt",       "Line Total",          110,   "e",    False),
    ("cost",     "Cost",                 95,   "e",    False),
    ("marg_d",   "Margin $",             95,   "e",    False),
    ("marg_p",   "Margin %",             85,   "e",    False),
    ("total",    "Total",               100,   "e",    False),
]
_COLS = tuple(name for name, *_ in _COL_DEFS)
_BLANK_VALUES = ("",) * len(_COLS)


class DailySalesDialog(ctk.CTkToplevel):
    """Non-modal report window showing all completed sales for today.

    Transactions appear as collapsible rows; expanding one reveals per-line
    detail (SKU, description, qty, RRP, disc $, line total, cost, margin)
    plus a TOTAL footer row.  A fixed summary panel at the bottom aggregates
    the entire day.

    Interaction
    -----------
    - Click the #0 (arrow) column header to expand / collapse all transactions.
    - Click the "Date & Time" column header to toggle sort order (▼ / ▲).
    - Click any transaction row (or child row) to enable "Reprint Receipt".
    - Right-click any row to open a context menu with "Reprint Receipt".
    """

    def __init__(self, master, **kwargs) -> None:
        super().__init__(master, **kwargs)
        self._rows: list[dict] = []
        self._selected_tx: Optional[dict] = None
        self._sort_desc: bool = True        # True = newest first (▼)
        self._all_expanded: bool = False    # False = next heading click expands all

        from datetime import datetime
        from zoneinfo import ZoneInfo
        _MELB = ZoneInfo("Australia/Melbourne")
        self._today_label = datetime.now(_MELB).strftime("%A, %d %B %Y")

        self.title(f"Daily Sales — {self._today_label}")
        self.geometry("1380x820")
        self.minsize(900, 500)
        self.resizable(True, True)

        # Force to front on Windows (topmost trick — removed after 300 ms)
        def _raise():
            self.lift()
            self.focus_force()
            self.attributes("-topmost", True)
            self.after(300, lambda: self.attributes("-topmost", False))
        self.after(50, _raise)

        self._build_ui()
        self._load()

    # ── Build ──────────────────────────────────────────────────────────────

    def _build_ui(self) -> None:
        # ── Header bar ────────────────────────────────────────────────────
        header = ctk.CTkFrame(
            self, height=44, corner_radius=0, fg_color=("gray88", "gray18"),
        )
        header.pack(fill="x", side="top")
        header.pack_propagate(False)

        ctk.CTkLabel(
            header,
            text=f"Daily Sales  —  {self._today_label}",
            font=ctk.CTkFont(size=14, weight="bold"),
        ).pack(side="left", padx=16, pady=10)

        ctk.CTkButton(
            header, text="Refresh", width=80, font=ctk.CTkFont(size=12),
            fg_color="transparent", border_width=1,
            border_color=("gray60", "gray45"),
            text_color=("gray40", "gray70"),
            hover_color=("gray85", "gray25"),
            command=self._load,
        ).pack(side="right", padx=16, pady=8)

        self._reprint_btn = ctk.CTkButton(
            header, text="Reprint Receipt", width=130, font=ctk.CTkFont(size=12),
            state="disabled",
            command=self._reprint_receipt,
        )
        self._reprint_btn.pack(side="right", padx=(0, 8), pady=8)

        self._status_lbl = ctk.CTkLabel(
            header, text="Loading…",
            font=ctk.CTkFont(size=11),
            text_color=("gray50", "gray60"),
        )
        self._status_lbl.pack(side="right", padx=(0, 12), pady=10)

        # ── Summary panel (bottom — packed before tree so it stays fixed) ──
        self._build_summary_panel()

        # ── Treeview area ─────────────────────────────────────────────────
        tree_outer = ctk.CTkFrame(self, fg_color=("gray92", "gray14"))
        tree_outer.pack(fill="both", expand=True, padx=12, pady=(8, 4))
        tree_outer.grid_rowconfigure(0, weight=1)
        tree_outer.grid_columnconfigure(0, weight=1)

        self._tree = ttk.Treeview(
            tree_outer,
            columns=_COLS,
            show="tree headings",
            selectmode="browse",
            style="Sales.Treeview",
        )

        # #0 column: just the expand arrow — clicking the heading toggles all
        self._tree.column("#0", width=16, minwidth=16, stretch=False)
        self._tree.heading("#0", text="", command=self._toggle_expand_all)

        for name, heading, width, anchor, stretch in _COL_DEFS:
            self._tree.column(name, width=width, minwidth=60,
                              anchor=anchor, stretch=stretch)
            self._tree.heading(name, text=heading, anchor=anchor)

        # Date heading gets sort command + initial arrow
        self._update_sort_heading()

        vsb = ttk.Scrollbar(tree_outer, orient="vertical",   command=self._tree.yview)
        hsb = ttk.Scrollbar(tree_outer, orient="horizontal", command=self._tree.xview)
        self._tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        self._tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")

        # Row tags
        self._tree.tag_configure("tx_row",   font=("Segoe UI", 10, "bold"))
        self._tree.tag_configure("line_row", font=("Segoe UI", 10))
        self._tree.tag_configure("sum_row",  font=("Segoe UI", 10, "bold"),
                                             foreground="#22c55e")
        # Divider row — a subtle colour-band shown between open transactions
        self._tree.tag_configure("div_row",
                                 background="#2a2a2a", foreground="#2a2a2a")

        # Event bindings
        self._tree.bind("<<TreeviewSelect>>", self._on_tree_select)
        self._tree.bind("<Button-3>",         self._on_right_click)
        self._tree.bind("<<TreeviewOpen>>",   self._on_tree_open)
        self._tree.bind("<<TreeviewClose>>",  self._on_tree_close)

    def _build_summary_panel(self) -> None:
        """Fixed panel at the bottom showing day-level aggregates."""
        panel = ctk.CTkFrame(
            self, fg_color=("gray86", "gray20"), corner_radius=0,
        )
        panel.pack(fill="x", side="bottom", padx=0, pady=0)

        inner = ctk.CTkFrame(panel, fg_color="transparent")
        inner.pack(fill="x", padx=16, pady=(8, 10))

        ctk.CTkLabel(
            inner, text="Day Summary",
            font=ctk.CTkFont(size=11, weight="bold"),
        ).pack(anchor="w", pady=(0, 6))

        # Metric columns
        metrics_row = ctk.CTkFrame(inner, fg_color="transparent")
        metrics_row.pack(fill="x", pady=(0, 6))

        for label, attr in [
            ("Transactions",  "_s_count"),
            ("Items Sold",    "_s_qty"),
            ("Total RRP",     "_s_rrp"),
            ("Discount",      "_s_disc"),
            ("Revenue",       "_s_revenue"),
            ("Total Cost",    "_s_cost"),
            ("Margin $",      "_s_marg_d"),
            ("Margin %",      "_s_marg_p"),
        ]:
            col = ctk.CTkFrame(metrics_row, fg_color="transparent")
            col.pack(side="left", padx=(0, 28))
            ctk.CTkLabel(
                col, text=label, font=ctk.CTkFont(size=10),
                text_color=("gray50", "gray55"),
            ).pack(anchor="w")
            lbl = ctk.CTkLabel(
                col, text="—", font=ctk.CTkFont(size=12, weight="bold"),
            )
            lbl.pack(anchor="w")
            setattr(self, attr, lbl)

        ctk.CTkFrame(inner, height=1, fg_color=("gray70", "gray35")).pack(
            fill="x", pady=(2, 6),
        )

        pay_row = ctk.CTkFrame(inner, fg_color="transparent")
        pay_row.pack(fill="x")

        ctk.CTkLabel(
            pay_row, text="Payment Breakdown:",
            font=ctk.CTkFont(size=10), text_color=("gray50", "gray55"),
        ).pack(side="left", padx=(0, 16))

        for method, attr in [
            ("Cash",   "_s_pay_cash"),
            ("EFT",    "_s_pay_eft"),
            ("Online", "_s_pay_online"),
        ]:
            ctk.CTkLabel(
                pay_row, text=f"{method}:",
                font=ctk.CTkFont(size=10), text_color=("gray50", "gray55"),
            ).pack(side="left", padx=(0, 4))
            lbl = ctk.CTkLabel(
                pay_row, text="—", font=ctk.CTkFont(size=12, weight="bold"),
            )
            lbl.pack(side="left", padx=(0, 24))
            setattr(self, attr, lbl)

    # ── Data loading ───────────────────────────────────────────────────────

    def _load(self) -> None:
        self._status_lbl.configure(text="Loading…", text_color=("gray50", "gray60"))
        threading.Thread(target=self._fetch, daemon=True).start()

    def _fetch(self) -> None:
        try:
            from src.pos.transaction_client import get_daily_transactions
            rows = get_daily_transactions()
            self.after(0, lambda: self._apply(rows))
        except Exception as exc:
            err = str(exc)
            self.after(0, lambda: self._status_lbl.configure(
                text=f"Error: {err}",
                text_color=("#b91c1c", "#f87171"),
            ))

    def _apply(self, rows: list[dict]) -> None:
        self._rows = sorted(
            rows,
            key=lambda tx: tx.get("completed_at", ""),
            reverse=self._sort_desc,
        )
        self._all_expanded = False
        self._selected_tx = None
        self._reprint_btn.configure(state="disabled")
        self._tree.delete(*self._tree.get_children())
        for tx in self._rows:
            self._insert_transaction(tx)
        count = len(self._rows)
        self._status_lbl.configure(
            text=f"{count} transaction{'s' if count != 1 else ''} today.",
            text_color=("gray50", "gray60"),
        )
        self._update_summary()

    # ── Treeview population ────────────────────────────────────────────────

    def _insert_transaction(self, tx: dict) -> None:
        """Insert one transaction parent row + child line rows + TOTAL footer."""
        tx_id   = tx["id"]
        tx_num  = tx.get("transaction_number") or ""
        total   = float(tx.get("total") or 0)
        lines   = tx.get("transaction_lines") or []

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
        payment_str = "  /  ".join(parts) if parts else "—"

        self._tree.insert(
            "", "end", iid=tx_id, text="",
            values=(
                tx_num,
                _fmt_dt(tx.get("completed_at", "")),
                tx.get("park_name") or "",
                tx.get("performed_by") or "",
                payment_str,
                "", "", "", "",
                f"${total:,.2f}",
            ),
            tags=("tx_row",),
            open=False,
        )

        if not lines:
            return

        t_qty = t_rrp = t_cost = 0.0

        for line in lines:
            qty        = float(line.get("qty") or 0)
            unit_price = float(line.get("unit_price") or 0)
            cost_price = float(line.get("cost_price") or 0)
            line_total = float(line.get("line_total") or 0)

            rrp_line  = unit_price * qty
            disc_line = rrp_line - line_total
            cost_line = cost_price * qty
            sell_ex   = line_total / 1.1
            marg_d    = (sell_ex - cost_line) if cost_price else None
            marg_p    = (marg_d / sell_ex * 100) if (marg_d is not None and sell_ex > 0) else None

            t_qty  += qty
            t_rrp  += rrp_line
            t_cost += cost_line

            self._tree.insert(
                tx_id, "end", text="",
                values=(
                    line.get("sku") or "",
                    line.get("description") or "",
                    f"{qty:g}",
                    f"${unit_price:,.2f}",
                    f"${disc_line:,.2f}" if disc_line > 0.005 else "—",
                    f"${line_total:,.2f}",
                    f"${cost_line:,.2f}" if cost_price else "—",
                    f"${marg_d:,.2f}"    if marg_d is not None else "—",
                    f"{marg_p:.1f}%"     if marg_p is not None else "—",
                    "",
                ),
                tags=("line_row",),
            )

        t_disc   = t_rrp - total
        t_ex_gst = total / 1.1
        t_marg_d = (t_ex_gst - t_cost) if t_cost else None
        t_marg_p = (t_marg_d / t_ex_gst * 100) if (t_marg_d is not None and t_ex_gst > 0) else None

        self._tree.insert(
            tx_id, "end", text="",
            values=(
                "TOTAL", "",
                f"{t_qty:g}", "",
                f"${t_disc:,.2f}" if t_disc > 0.005 else "—",
                f"${total:,.2f}",
                f"${t_cost:,.2f}"    if t_cost else "—",
                f"${t_marg_d:,.2f}" if t_marg_d is not None else "—",
                f"{t_marg_p:.1f}%"  if t_marg_p is not None else "—",
                "",
            ),
            tags=("sum_row",),
        )

    # ── Day summary panel ──────────────────────────────────────────────────

    def _update_summary(self) -> None:
        _blank = [
            "_s_count", "_s_qty", "_s_rrp", "_s_disc", "_s_revenue",
            "_s_cost", "_s_marg_d", "_s_marg_p",
            "_s_pay_cash", "_s_pay_eft", "_s_pay_online",
        ]
        if not self._rows:
            for attr in _blank:
                getattr(self, attr).configure(text="—")
            return

        d_count = len(self._rows)
        d_qty = d_rrp = d_revenue = d_cost = d_cash = d_eft = d_online = 0.0

        for tx in self._rows:
            d_revenue += float(tx.get("total") or 0)
            d_cash    += float(tx.get("payment_cash")   or 0)
            d_online  += float(tx.get("payment_online") or 0)
            for e in (tx.get("payment_eft") or []):
                d_eft += float(e.get("amount") or 0)
            for line in (tx.get("transaction_lines") or []):
                qty  = float(line.get("qty") or 0)
                up   = float(line.get("unit_price") or 0)
                cost = float(line.get("cost_price") or 0)
                d_qty  += qty
                d_rrp  += up * qty
                d_cost += cost * qty

        d_disc   = d_rrp - d_revenue
        d_ex_gst = d_revenue / 1.1 if d_revenue else 0
        d_marg_d = (d_ex_gst - d_cost) if d_cost else None
        d_marg_p = (d_marg_d / d_ex_gst * 100) if (d_marg_d is not None and d_ex_gst > 0) else None

        self._s_count.configure(text=str(d_count))
        self._s_qty.configure(text=f"{d_qty:g}")
        self._s_rrp.configure(text=f"${d_rrp:,.2f}")
        self._s_disc.configure(text=f"${d_disc:,.2f}" if d_disc > 0.005 else "—")
        self._s_revenue.configure(text=f"${d_revenue:,.2f}")
        self._s_cost.configure(text=f"${d_cost:,.2f}" if d_cost else "—")
        self._s_marg_d.configure(text=f"${d_marg_d:,.2f}" if d_marg_d is not None else "—")
        self._s_marg_p.configure(text=f"{d_marg_p:.1f}%" if d_marg_p is not None else "—")
        self._s_pay_cash.configure(text=f"${d_cash:,.2f}"   if d_cash   else "—")
        self._s_pay_eft.configure(text=f"${d_eft:,.2f}"     if d_eft    else "—")
        self._s_pay_online.configure(text=f"${d_online:,.2f}" if d_online else "—")

    # ── Expand / collapse ──────────────────────────────────────────────────

    def _toggle_expand_all(self) -> None:
        """Heading click on #0: expand all if currently collapsed, else collapse all."""
        tx_ids = {r["id"] for r in self._rows}
        if not self._all_expanded:
            # Expand all tx rows, then add dividers in reverse to preserve indices
            top = [iid for iid in self._tree.get_children("") if iid in tx_ids]
            for iid in top:
                self._tree.item(iid, open=True)
            for iid in reversed(top):
                self._insert_divider_after(iid, tx_ids)
            self._all_expanded = True
        else:
            # Remove all dividers, then collapse all tx rows
            for iid in list(self._tree.get_children("")):
                if _is_divider(iid):
                    self._tree.delete(iid)
            for iid in self._tree.get_children(""):
                self._tree.item(iid, open=False)
            self._all_expanded = False

    def _on_tree_open(self, _event=None) -> None:
        """User manually expanded a transaction row — insert divider after it."""
        iid = self._tree.focus()
        if not iid:
            return
        tx_ids = {r["id"] for r in self._rows}
        if iid not in tx_ids:
            return
        self._insert_divider_after(iid, tx_ids)

    def _on_tree_close(self, _event=None) -> None:
        """User manually collapsed a transaction row — remove its divider."""
        iid = self._tree.focus()
        if not iid:
            return
        sep_iid = _sep_id(iid)
        if self._tree.exists(sep_iid):
            self._tree.delete(sep_iid)

    def _insert_divider_after(self, tx_iid: str, tx_ids: set) -> None:
        """Insert a divider row immediately after tx_iid (if not already present)."""
        sep_iid = _sep_id(tx_iid)
        if self._tree.exists(sep_iid):
            return
        children = list(self._tree.get_children(""))
        try:
            idx = children.index(tx_iid)
        except ValueError:
            return
        self._tree.insert("", idx + 1, iid=sep_iid, text="",
                          values=_BLANK_VALUES, tags=("div_row",))

    # ── Sort ───────────────────────────────────────────────────────────────

    def _toggle_sort(self) -> None:
        """Heading click on 'detail': reverse the date sort order."""
        self._sort_desc = not self._sort_desc
        self._update_sort_heading()
        self._rows.sort(
            key=lambda tx: tx.get("completed_at", ""),
            reverse=self._sort_desc,
        )
        # Re-render collapsed (simplest; expand state is discarded on sort)
        self._all_expanded = False
        self._selected_tx = None
        self._reprint_btn.configure(state="disabled")
        self._tree.delete(*self._tree.get_children())
        for tx in self._rows:
            self._insert_transaction(tx)

    def _update_sort_heading(self) -> None:
        arrow = " ▼" if self._sort_desc else " ▲"
        self._tree.heading(
            "detail",
            text=f"Date & Time / Desc{arrow}",
            command=self._toggle_sort,
        )

    # ── Selection & interaction ────────────────────────────────────────────

    def _on_tree_select(self, _event=None) -> None:
        sel = self._tree.selection()
        if not sel:
            self._selected_tx = None
            self._reprint_btn.configure(state="disabled")
            return
        tx = self._resolve_tx(sel[0])
        self._selected_tx = tx
        self._reprint_btn.configure(state="normal" if tx else "disabled")

    def _on_right_click(self, event) -> None:
        iid = self._tree.identify_row(event.y)
        if not iid:
            return
        tx = self._resolve_tx(iid)
        if tx is None:
            return
        self._tree.selection_set(iid)
        self._selected_tx = tx
        self._reprint_btn.configure(state="normal")
        menu = tk.Menu(self, tearoff=0)
        menu.add_command(label="Reprint Receipt", command=self._reprint_receipt)
        menu.tk_popup(event.x_root, event.y_root)

    def _resolve_tx(self, iid: str) -> Optional[dict]:
        """Return the transaction dict for a treeview iid (parent or child row)."""
        if _is_divider(iid):
            return None
        tx_ids = {r["id"] for r in self._rows}
        if iid in tx_ids:
            return next((r for r in self._rows if r["id"] == iid), None)
        parent_iid = self._tree.parent(iid)
        if parent_iid in tx_ids:
            return next((r for r in self._rows if r["id"] == parent_iid), None)
        return None

    # ── Receipt reprint ────────────────────────────────────────────────────

    def _reprint_receipt(self) -> None:
        tx = self._selected_tx
        if not tx:
            return

        from src.config import config
        printer = getattr(getattr(config, "device", None), "receipt_printer", None)
        if not printer:
            messagebox.showwarning(
                "No Printer",
                "No receipt printer is configured.\n"
                "Go to Settings → Printers to set one up.",
                parent=self,
            )
            return

        lines = tx.get("transaction_lines") or []
        cart_items = {
            (line.get("item_id") or str(i)): {
                "sku":        line.get("sku") or "",
                "title":      line.get("description") or "",
                "qty":        float(line.get("qty") or 1),
                "unit_price": float(line.get("unit_price") or 0),
                "disc_pct":   float(line.get("discount_pct") or 0),
                "cost_price": float(line.get("cost_price") or 0),
            }
            for i, line in enumerate(lines)
        }

        def _thread():
            try:
                from src.pos.receipt_generator import generate_receipt
                from src.printer_utils import print_pdf
                pdf_path = generate_receipt(tx, cart_items)
                print_pdf(pdf_path, printer)
            except Exception as exc:
                err = str(exc)
                self.after(0, lambda: messagebox.showerror(
                    "Print Failed", f"Could not print receipt:\n{err}", parent=self,
                ))

        threading.Thread(target=_thread, daemon=True).start()


# ── Module helpers ─────────────────────────────────────────────────────────────

def _sep_id(tx_iid: str) -> str:
    return f"__sep_{tx_iid}"


def _is_divider(iid: str) -> bool:
    return iid.startswith("__sep_")


def _fmt_dt(iso: str) -> str:
    """Format an ISO UTC datetime string as 'DD-MM-YYYY  H:MM AM/PM' (Melbourne time)."""
    try:
        from datetime import datetime
        from zoneinfo import ZoneInfo
        _MELB = ZoneInfo("Australia/Melbourne")
        dt = datetime.fromisoformat(iso.replace("Z", "+00:00"))
        dt_local = dt.astimezone(_MELB)
        hour = dt_local.hour % 12 or 12
        am_pm = "AM" if dt_local.hour < 12 else "PM"
        return (
            f"{dt_local.day:02d}-{dt_local.month:02d}-{dt_local.year}  "
            f"{hour}:{dt_local.minute:02d} {am_pm}"
        )
    except Exception:
        return iso[:16] if len(iso) >= 16 else iso
