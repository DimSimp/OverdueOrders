from __future__ import annotations

import threading
from tkinter import ttk
from typing import Optional

import customtkinter as ctk


_PAGE_SIZE = 100
_SORT_COLS = {
    "#0": None,
    "cust_id":    "customer_id",
    "first_name": "first_name",
    "surname":    "surname",
    "business":   "business",
    "city":       "city",
    "phone":      "phone_1",
}


class CustomerTab(ctk.CTkFrame):
    """Customers module — search list + detail panel."""

    def __init__(self, parent, current_user=None, on_load_in_till=None, **kwargs):
        super().__init__(parent, fg_color="transparent", **kwargs)
        self.current_user = current_user
        self._on_load_in_till = on_load_in_till

        self._page = 0
        self._total = 0
        self._sort_col = "customer_id"
        self._sort_asc = True
        self._debounce_id: Optional[str] = None
        self._selected_customer: Optional[dict] = None
        self._detail_loaded = False

        self._build_ui()

    # ── Build ────────────────────────────────────────────────────────────

    def _build_ui(self):
        # ── Top bar ──────────────────────────────────────────────────────
        top_bar = ctk.CTkFrame(self, fg_color=("gray88", "gray18"), corner_radius=0)
        top_bar.pack(fill="x", side="top")

        ctk.CTkLabel(
            top_bar, text="Search:",
            font=ctk.CTkFont(size=12),
        ).pack(side="left", padx=(16, 4), pady=10)

        self._search_var = ctk.StringVar()
        self._search_var.trace_add("write", self._on_search_change)
        self._search_entry = ctk.CTkEntry(
            top_bar,
            textvariable=self._search_var,
            placeholder_text="Name, ID, phone, email…",
            width=260,
            font=ctk.CTkFont(size=12),
        )
        self._search_entry.pack(side="left", pady=10)

        ctk.CTkLabel(
            top_bar, text="Show:",
            font=ctk.CTkFont(size=12),
        ).pack(side="left", padx=(16, 4), pady=10)

        self._active_var = ctk.StringVar(value="Active only")
        ctk.CTkOptionMenu(
            top_bar,
            values=["Active only", "All customers", "Inactive only"],
            variable=self._active_var,
            width=140,
            font=ctk.CTkFont(size=12),
            command=lambda _: self._trigger_search(),
        ).pack(side="left", pady=10)

        ctk.CTkButton(
            top_bar, text="Import CSV", width=100, height=30,
            font=ctk.CTkFont(size=12),
            fg_color="transparent", border_width=1,
            border_color=("gray60", "gray45"),
            text_color=("gray30", "gray70"),
            hover_color=("gray85", "gray25"),
            command=self._open_import,
        ).pack(side="right", padx=(0, 12), pady=10)

        ctk.CTkButton(
            top_bar, text="+ New Customer", width=130, height=30,
            font=ctk.CTkFont(size=12, weight="bold"),
            command=self._open_new,
        ).pack(side="right", padx=(0, 8), pady=10)

        # ── Main area: list + detail panel ───────────────────────────────
        main = ctk.CTkFrame(self, fg_color="transparent")
        main.pack(fill="both", expand=True)

        # List pane
        list_pane = ctk.CTkFrame(main, fg_color="transparent")
        list_pane.pack(fill="both", expand=True)

        # Treeview
        tree_frame = ctk.CTkFrame(list_pane, fg_color="transparent")
        tree_frame.pack(fill="both", expand=True, padx=8, pady=(8, 0))

        cols = ("cust_id", "first_name", "surname", "business", "city", "phone")
        self._tree = ttk.Treeview(
            tree_frame,
            columns=cols,
            show="headings",
            style="Cust.Treeview",
            selectmode="browse",
        )
        vsb = ttk.Scrollbar(tree_frame, orient="vertical", command=self._tree.yview)
        self._tree.configure(yscrollcommand=vsb.set)

        headings = {
            "cust_id":    ("ID",         70,  "w"),
            "first_name": ("First Name", 130, "w"),
            "surname":    ("Surname",    130, "w"),
            "business":   ("Business",  200, "w"),
            "city":       ("City",      110, "w"),
            "phone":      ("Phone",     120, "w"),
        }
        for col, (text, width, anchor) in headings.items():
            self._tree.heading(
                col, text=text, anchor=anchor,
                command=lambda c=col: self._on_heading_click(c),
            )
            self._tree.column(col, width=width, anchor=anchor, stretch=(col == "business"))

        self._tree.pack(side="left", fill="both", expand=True)
        vsb.pack(side="right", fill="y")

        self._tree.bind("<<TreeviewSelect>>", self._on_row_select)
        self._tree.bind("<Button-3>", self._on_right_click)
        self._tree.bind("<Double-1>", self._on_double_click)

        # Pagination bar
        pag = ctk.CTkFrame(list_pane, fg_color="transparent")
        pag.pack(fill="x", padx=8, pady=4)

        self._btn_prev = ctk.CTkButton(
            pag, text="◀ Prev", width=80, height=26,
            font=ctk.CTkFont(size=11), command=self._prev_page,
        )
        self._btn_prev.pack(side="left")

        self._lbl_page = ctk.CTkLabel(pag, text="", font=ctk.CTkFont(size=11))
        self._lbl_page.pack(side="left", padx=12)

        self._btn_next = ctk.CTkButton(
            pag, text="Next ▶", width=80, height=26,
            font=ctk.CTkFont(size=11), command=self._next_page,
        )
        self._btn_next.pack(side="left")

        self._lbl_count = ctk.CTkLabel(
            pag, text="", font=ctk.CTkFont(size=11),
            text_color=("gray50", "gray60"),
        )
        self._lbl_count.pack(side="right", padx=8)

        # Thin separator
        ctk.CTkFrame(self, fg_color=("gray70", "gray35"), height=1).pack(fill="x")

        # ── Detail panel (hidden until a customer is selected) ────────────
        self._detail_frame = ctk.CTkFrame(self, fg_color="transparent", height=320)
        self._detail_frame.pack_propagate(False)
        # Not packed yet — shown on first row select

        self._detail_tabs = ctk.CTkTabview(self._detail_frame, anchor="nw", height=300)
        self._detail_tabs.pack(fill="both", expand=True, padx=0, pady=0)

        for name in ("Customer Info", "Sale History", "Quotes", "Invoices", "Repairs", "Deposits"):
            self._detail_tabs.add(name)

        # Real tabs
        from src.gui.customers.tabs.customer_info_tab import CustomerInfoTab
        self._info_tab = CustomerInfoTab(
            self._detail_tabs.tab("Customer Info"),
            on_edit=self._open_edit,
        )
        self._info_tab.pack(fill="both", expand=True)

        from src.gui.customers.tabs.sale_history_tab import SaleHistoryTab
        self._history_tab = SaleHistoryTab(self._detail_tabs.tab("Sale History"))
        self._history_tab.pack(fill="both", expand=True)

        # Stub tabs
        for name in ("Quotes", "Invoices", "Repairs", "Deposits"):
            ctk.CTkLabel(
                self._detail_tabs.tab(name),
                text=f"{name} — coming soon",
                font=ctk.CTkFont(size=13),
                text_color=("gray50", "gray60"),
            ).pack(expand=True)

        # Wire sale history lazy load
        self._detail_tabs.configure(command=self._on_detail_tab_change)

        # Right-click menu
        self._ctx_menu = None

    # ── Search ───────────────────────────────────────────────────────────

    def _on_search_change(self, *_):
        if self._debounce_id:
            self.after_cancel(self._debounce_id)
        self._debounce_id = self.after(300, self._trigger_search)

    def _trigger_search(self):
        self._page = 0
        self._run_search()

    def _run_search(self):
        query = self._search_var.get().strip()
        active_raw = self._active_var.get()
        active_filter = {
            "Active only": "active",
            "Inactive only": "inactive",
            "All customers": "all",
        }.get(active_raw, "active")

        sort_col = self._sort_col
        sort_asc = self._sort_asc
        page = self._page

        self._lbl_count.configure(text="Searching…")

        def _thread():
            from src.customers.customer_client import search_customers
            try:
                rows, total = search_customers(
                    query, active_filter, sort_col, sort_asc, page, _PAGE_SIZE
                )
                self.after(0, lambda: self._populate(rows, total))
            except Exception as exc:
                err = str(exc)
                self.after(0, lambda: self._lbl_count.configure(
                    text=f"Error: {err}", text_color=("red", "#e74c3c")
                ))

        threading.Thread(target=_thread, daemon=True).start()

    def _populate(self, rows: list, total: int):
        self._total = total
        self._tree.delete(*self._tree.get_children())

        for r in rows:
            self._tree.insert("", "end", iid=r["id"], values=(
                r.get("customer_code") or "",
                r.get("first_name") or "",
                r.get("surname") or "",
                r.get("business") or "",
                r.get("city") or "",
                r.get("mobile") or r.get("phone_1") or "",
            ))

        self._lbl_page.configure(
            text=f"Page {self._page + 1}" if self._page > 0 or rows else ""
        )
        self._lbl_count.configure(
            text=f"{len(rows)} customer(s)" if rows else "No results",
            text_color=("gray50", "gray60"),
        )
        self._btn_prev.configure(state="normal" if self._page > 0 else "disabled")
        self._btn_next.configure(
            state="normal" if (self._page + 1) * _PAGE_SIZE < total else "disabled"
        )

    # ── Sort ─────────────────────────────────────────────────────────────

    def _on_heading_click(self, col: str):
        db_col = _SORT_COLS.get(col)
        if not db_col:
            return
        if self._sort_col == db_col:
            self._sort_asc = not self._sort_asc
        else:
            self._sort_col = db_col
            self._sort_asc = True
        self._page = 0
        self._run_search()

    # ── Pagination ───────────────────────────────────────────────────────

    def _prev_page(self):
        if self._page > 0:
            self._page -= 1
            self._run_search()

    def _next_page(self):
        if (self._page + 1) * _PAGE_SIZE < self._total:
            self._page += 1
            self._run_search()

    # ── Row selection ────────────────────────────────────────────────────

    def _on_row_select(self, _event=None):
        sel = self._tree.selection()
        if not sel:
            return
        uuid = sel[0]
        self._load_customer_detail(uuid)

    def _on_double_click(self, _event=None):
        sel = self._tree.selection()
        if sel:
            self._open_edit_by_id(sel[0])

    def _load_customer_detail(self, uuid: str):
        def _thread():
            from src.customers.customer_client import get_customer
            try:
                customer = get_customer(uuid)
                self.after(0, lambda: self._show_detail(customer))
            except Exception as exc:
                self.after(0, lambda: None)  # silently fail for now

        threading.Thread(target=_thread, daemon=True).start()

    def _show_detail(self, customer: Optional[dict]):
        if not customer:
            return
        self._selected_customer = customer

        # Show detail panel if hidden
        if not self._detail_frame.winfo_ismapped():
            self._detail_frame.pack(fill="x", side="bottom")

        self._info_tab.load_customer(customer)
        self._history_tab.load_for_customer(customer["id"])
        self._detail_tabs.set("Customer Info")

    # ── Detail tab change ─────────────────────────────────────────────────

    def _on_detail_tab_change(self):
        tab = self._detail_tabs.get()
        if tab == "Sale History":
            self._history_tab.on_tab_selected()

    # ── Right-click ──────────────────────────────────────────────────────

    def _on_right_click(self, event):
        row = self._tree.identify_row(event.y)
        if not row:
            return
        self._tree.selection_set(row)

        from tkinter import Menu
        menu = Menu(self, tearoff=0)
        menu.add_command(label="Edit Customer", command=lambda: self._open_edit_by_id(row))
        menu.add_separator()
        if self._on_load_in_till:
            menu.add_command(label="Load in Till", command=lambda: self._load_in_till_by_id(row))
        else:
            menu.add_command(label="Load in Till", state="disabled")
        try:
            menu.tk_popup(event.x_root, event.y_root)
        finally:
            menu.grab_release()

    def _load_in_till_by_id(self, uuid: str) -> None:
        def _thread():
            from src.customers.customer_client import get_customer
            try:
                customer = get_customer(uuid)
                self.after(0, lambda: self._finish_load_in_till(customer))
            except Exception as exc:
                err = str(exc)
                self.after(0, lambda: self._show_load_in_till_error(err))

        threading.Thread(target=_thread, daemon=True).start()

    def _finish_load_in_till(self, customer: Optional[dict]) -> None:
        from tkinter import messagebox
        if not customer:
            messagebox.showerror(
                "Load in Till",
                "That customer could not be loaded.",
                parent=self.winfo_toplevel(),
            )
            return
        if self._on_load_in_till:
            self._on_load_in_till(customer)

    def _show_load_in_till_error(self, err: str) -> None:
        from tkinter import messagebox
        messagebox.showerror(
            "Load in Till",
            f"Could not load that customer into the till:\n{err}",
            parent=self.winfo_toplevel(),
        )

    # ── Forms ────────────────────────────────────────────────────────────

    def _open_new(self):
        from src.gui.customers.customer_form import CustomerForm
        CustomerForm(self.winfo_toplevel(), on_saved=self._on_customer_saved)

    def _open_edit(self, customer: dict):
        from src.gui.customers.customer_form import CustomerForm
        CustomerForm(
            self.winfo_toplevel(),
            customer=customer,
            on_saved=self._on_customer_saved,
        )

    def _open_edit_by_id(self, uuid: str):
        def _thread():
            from src.customers.customer_client import get_customer
            try:
                customer = get_customer(uuid)
                if customer:
                    self.after(0, lambda: self._open_edit(customer))
            except Exception:
                pass

        threading.Thread(target=_thread, daemon=True).start()

    def _on_customer_saved(self, customer: dict):
        # Refresh the detail panel
        self._show_detail(customer)
        # Re-run search to update list
        self._run_search()

    def show_customer(self, uuid: str) -> None:
        """Navigate to and display the detail panel for the given customer UUID.

        Called from the Till tab's 'Profile →' button.
        """
        self._load_customer_detail(uuid)

    # ── Import ───────────────────────────────────────────────────────────

    def _open_import(self):
        from src.gui.customers.import_dialog import CustomerImportDialog
        CustomerImportDialog(
            self.winfo_toplevel(),
            on_done=self._run_search,
        )
