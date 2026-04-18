"""Recall dialog — modal list of parked transactions for POS recall."""
from __future__ import annotations

import threading

import customtkinter as ctk
from tkinter import messagebox, ttk


class RecallDialog(ctk.CTkToplevel):
    """Modal that lists parked transactions and lets staff restore one to the till."""

    def __init__(self, master, on_select: callable, **kwargs) -> None:
        super().__init__(master, **kwargs)
        self._on_select = on_select
        self._rows: list[dict] = []

        self.title("Parked Transactions")
        self.geometry("700x460")
        self.minsize(600, 360)
        self.resizable(True, True)
        self.grab_set()
        self.transient(master)
        self.after(50, self.lift)

        self._build_ui()
        self._load()

    # ── Build ──────────────────────────────────────────────────────────────

    def _build_ui(self) -> None:
        ctk.CTkLabel(
            self,
            text="Parked Transactions",
            font=ctk.CTkFont(size=15, weight="bold"),
        ).pack(anchor="w", padx=16, pady=(14, 2))

        ctk.CTkLabel(
            self,
            text="Select a transaction to restore the cart, or delete one permanently.",
            font=ctk.CTkFont(size=11),
            text_color=("gray50", "gray60"),
        ).pack(anchor="w", padx=16, pady=(0, 10))

        # Treeview
        tree_frame = ctk.CTkFrame(self, fg_color=("gray92", "gray14"))
        tree_frame.pack(fill="both", expand=True, padx=16, pady=(0, 8))

        cols = ("date", "tx_num", "customer", "total")
        self._tree = ttk.Treeview(
            tree_frame, columns=cols, show="headings", selectmode="browse",
        )
        self._tree.heading("date",     text="Date Parked")
        self._tree.heading("tx_num",   text="Transaction #")
        self._tree.heading("customer", text="Customer")
        self._tree.heading("total",    text="Total", anchor="e")

        self._tree.column("date",     width=180, minwidth=140, anchor="w", stretch=False)
        self._tree.column("tx_num",   width=140, minwidth=110, anchor="w", stretch=False)
        self._tree.column("customer", width=220, minwidth=120, anchor="w")
        self._tree.column("total",    width=110, minwidth=80,  anchor="e", stretch=False)

        vsb = ttk.Scrollbar(tree_frame, orient="vertical", command=self._tree.yview)
        self._tree.configure(yscrollcommand=vsb.set)
        self._tree.pack(side="left", fill="both", expand=True)
        vsb.pack(side="right", fill="y")

        self._tree.bind("<<TreeviewSelect>>", self._on_row_select)
        self._tree.bind("<Double-1>", self._confirm_select)

        # Status label
        self._status_lbl = ctk.CTkLabel(
            self,
            text="Loading…",
            font=ctk.CTkFont(size=11),
            text_color=("gray50", "gray60"),
        )
        self._status_lbl.pack(anchor="w", padx=16, pady=(0, 4))

        # Button row — delete on left, cancel + select on right
        btn_row = ctk.CTkFrame(self, fg_color="transparent")
        btn_row.pack(fill="x", padx=16, pady=(0, 14))

        self._delete_btn = ctk.CTkButton(
            btn_row, text="Delete", width=90,
            state="disabled",
            fg_color=("#b91c1c", "#7f1d1d"),
            hover_color=("#991b1b", "#6b1a1a"),
            command=self._delete_selected,
        )
        self._delete_btn.pack(side="left")

        ctk.CTkButton(
            btn_row, text="Cancel", width=90,
            fg_color="transparent", border_width=1,
            border_color=("gray60", "gray45"),
            text_color=("gray20", "gray90"),
            hover_color=("gray85", "gray25"),
            command=self.destroy,
        ).pack(side="right", padx=(8, 0))

        self._select_btn = ctk.CTkButton(
            btn_row, text="Select for POS", width=140,
            state="disabled",
            command=self._confirm_select,
        )
        self._select_btn.pack(side="right")

    # ── Data loading ───────────────────────────────────────────────────────

    def _load(self) -> None:
        threading.Thread(target=self._fetch, daemon=True).start()

    def _fetch(self) -> None:
        try:
            from src.pos.transaction_client import get_parked_transactions
            rows = get_parked_transactions()
            self.after(0, lambda: self._apply(rows))
        except Exception as exc:
            err = str(exc)
            self.after(0, lambda: self._status_lbl.configure(
                text=f"Error: {err}",
                text_color=("#b91c1c", "#f87171"),
            ))

    def _apply(self, rows: list[dict]) -> None:
        self._rows = rows
        self._tree.delete(*self._tree.get_children())
        for row in rows:
            self._tree.insert(
                "", "end", iid=row["id"],
                values=(
                    _fmt_dt(row.get("created_at", "")),
                    row.get("transaction_number", ""),
                    row.get("park_name") or "",
                    f"${float(row.get('total') or 0):,.2f}",
                ),
            )
        self._update_status()

    def _update_status(self) -> None:
        count = len(self._rows)
        self._status_lbl.configure(
            text=(
                f"{count} parked transaction{'s' if count != 1 else ''}."
                if count > 0 else "No parked transactions."
            ),
            text_color=("gray50", "gray60"),
        )

    # ── Interaction ────────────────────────────────────────────────────────

    def _on_row_select(self, _event=None) -> None:
        sel = self._tree.selection()
        state = "normal" if sel else "disabled"
        self._select_btn.configure(state=state)
        self._delete_btn.configure(state=state)

    def _confirm_select(self, _event=None) -> None:
        sel = self._tree.selection()
        if not sel:
            return
        tx_id = sel[0]
        tx = next((r for r in self._rows if r["id"] == tx_id), None)
        if tx is None:
            return
        self.destroy()
        self._on_select(tx)

    def _delete_selected(self) -> None:
        sel = self._tree.selection()
        if not sel:
            return
        tx_id = sel[0]
        tx = next((r for r in self._rows if r["id"] == tx_id), None)
        if tx is None:
            return

        tx_num = tx.get("transaction_number", "")
        if not messagebox.askyesno(
            "Delete Parked Transaction",
            f"Permanently delete parked transaction {tx_num}?\n"
            "This cannot be undone.",
            parent=self,
        ):
            return

        self._delete_btn.configure(state="disabled")
        self._select_btn.configure(state="disabled")

        def _do_delete():
            try:
                from src.pos.transaction_client import delete_parked_transaction
                delete_parked_transaction(tx_id)
                self.after(0, lambda: self._on_deleted(tx_id))
            except Exception as exc:
                err = str(exc)
                self.after(0, lambda: self._status_lbl.configure(
                    text=f"Delete failed: {err}",
                    text_color=("#b91c1c", "#f87171"),
                ))

        threading.Thread(target=_do_delete, daemon=True).start()

    def _on_deleted(self, tx_id: str) -> None:
        self._rows = [r for r in self._rows if r["id"] != tx_id]
        self._tree.delete(tx_id)
        self._update_status()


# ── Helpers ───────────────────────────────────────────────────────────────────

def _fmt_dt(iso: str) -> str:
    """Format an ISO datetime string as 'DD-MM-YYYY  H:MM AM/PM' (Melbourne time)."""
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
