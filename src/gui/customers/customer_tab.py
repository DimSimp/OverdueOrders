from __future__ import annotations

import threading
from tkinter import ttk
from typing import Optional

import customtkinter as ctk


_PAGE_SIZE = 100
_SORT_COLS = {
    "#0": None,
    "cust_id": "customer_id",
    "first_name": "first_name",
    "surname": "surname",
    "business": "business",
    "city": "city",
    "phone": "phone_1",
}


class CustomerTab(ctk.CTkFrame):
    """Customers module - search list + detail panel."""

    def __init__(self, parent, current_user=None, on_load_in_till=None, on_refund_in_till=None, **kwargs):
        super().__init__(parent, fg_color="transparent", **kwargs)
        self.current_user = current_user
        self._on_load_in_till = on_load_in_till
        self._on_refund_in_till = on_refund_in_till

        self._page = 0
        self._total = 0
        self._sort_col = "customer_id"
        self._sort_asc = True
        self._debounce_id: Optional[str] = None
        self._selected_customer: Optional[dict] = None
        self._detail_loaded = False
        self._merge_source_customer: Optional[dict] = None

        self._build_ui()

    def _build_ui(self):
        top_bar = ctk.CTkFrame(self, fg_color=("gray88", "gray18"), corner_radius=0)
        top_bar.pack(fill="x", side="top")

        ctk.CTkLabel(
            top_bar,
            text="Search:",
            font=ctk.CTkFont(size=12),
        ).pack(side="left", padx=(16, 4), pady=10)

        self._search_var = ctk.StringVar()
        self._search_var.trace_add("write", self._on_search_change)
        self._search_entry = ctk.CTkEntry(
            top_bar,
            textvariable=self._search_var,
            placeholder_text="Name, ID, phone, email...",
            width=260,
            font=ctk.CTkFont(size=12),
        )
        self._search_entry.pack(side="left", pady=10)

        self._btn_prev = ctk.CTkButton(
            top_bar,
            text="Prev",
            width=72,
            height=30,
            font=ctk.CTkFont(size=11),
            command=self._prev_page,
        )
        self._btn_prev.pack(side="left", padx=(16, 0), pady=10)

        self._lbl_page = ctk.CTkLabel(top_bar, text="", font=ctk.CTkFont(size=11))
        self._lbl_page.pack(side="left", padx=(10, 10), pady=10)

        self._btn_next = ctk.CTkButton(
            top_bar,
            text="Next",
            width=72,
            height=30,
            font=ctk.CTkFont(size=11),
            command=self._next_page,
        )
        self._btn_next.pack(side="left", pady=10)

        self._lbl_count = ctk.CTkLabel(
            top_bar,
            text="",
            font=ctk.CTkFont(size=11),
            text_color=("gray50", "gray60"),
        )
        self._lbl_count.pack(side="left", padx=(12, 0), pady=10)

        self._merge_banner = ctk.CTkFrame(
            top_bar,
            fg_color=("#e8f1ff", "#20314d"),
            corner_radius=8,
            border_width=1,
            border_color=("#8fb2ff", "#38558c"),
        )
        self._lbl_merge = ctk.CTkLabel(
            self._merge_banner,
            text="",
            font=ctk.CTkFont(size=11, weight="bold"),
            text_color=("#24406f", "#d8e6ff"),
        )
        self._lbl_merge.pack(side="left", padx=(10, 8), pady=6)
        ctk.CTkButton(
            self._merge_banner,
            text="Cancel",
            width=64,
            height=24,
            font=ctk.CTkFont(size=11),
            fg_color="transparent",
            border_width=1,
            border_color=("#7b9be0", "#4d6da8"),
            text_color=("#24406f", "#d8e6ff"),
            hover_color=("#d7e5ff", "#2a4063"),
            command=self._clear_merge_source,
        ).pack(side="right", padx=(0, 8), pady=5)

        ctk.CTkButton(
            top_bar,
            text="Import CSV",
            width=100,
            height=30,
            font=ctk.CTkFont(size=12),
            fg_color="transparent",
            border_width=1,
            border_color=("gray60", "gray45"),
            text_color=("gray30", "gray70"),
            hover_color=("gray85", "gray25"),
            command=self._open_import,
        ).pack(side="right", padx=(0, 12), pady=10)

        ctk.CTkButton(
            top_bar,
            text="+ New Customer",
            width=130,
            height=30,
            font=ctk.CTkFont(size=12, weight="bold"),
            command=self._open_new,
        ).pack(side="right", padx=(0, 8), pady=10)

        main = ctk.CTkFrame(self, fg_color="transparent")
        main.pack(fill="both", expand=True)

        list_pane = ctk.CTkFrame(main, fg_color="transparent")
        list_pane.pack(fill="both", expand=True)

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
            "cust_id": ("ID", 70, "w"),
            "first_name": ("First Name", 130, "w"),
            "surname": ("Surname", 130, "w"),
            "business": ("Business", 200, "w"),
            "city": ("City", 110, "w"),
            "phone": ("Phone", 120, "w"),
        }
        for col, (text, width, anchor) in headings.items():
            self._tree.heading(
                col,
                text=text,
                anchor=anchor,
                command=lambda c=col: self._on_heading_click(c),
            )
            self._tree.column(col, width=width, anchor=anchor, stretch=(col == "business"))

        self._tree.pack(side="left", fill="both", expand=True)
        vsb.pack(side="right", fill="y")

        self._tree.bind("<<TreeviewSelect>>", self._on_row_select)
        self._tree.bind("<ButtonRelease-1>", self._on_left_click_release)
        self._tree.bind("<Button-3>", self._on_right_click)
        self._tree.bind("<Double-1>", self._on_double_click)

        ctk.CTkFrame(self, fg_color=("gray70", "gray35"), height=1).pack(fill="x")

        self._detail_frame = ctk.CTkFrame(self, fg_color="transparent", height=400)
        self._detail_frame.pack_propagate(False)

        self._detail_tabs = ctk.CTkTabview(self._detail_frame, anchor="nw", height=380)
        self._detail_tabs.pack(fill="both", expand=True, padx=0, pady=0)

        for name in (
            "Customer Info",
            "Sale History",
            "Quotes",
            "Invoices",
            "Repairs",
            "Deposits",
            "Audit",
        ):
            self._detail_tabs.add(name)

        from src.gui.customers.tabs.customer_info_tab import CustomerInfoTab

        self._info_tab = CustomerInfoTab(
            self._detail_tabs.tab("Customer Info"),
            on_edit=self._open_edit,
        )
        self._info_tab.pack(fill="both", expand=True)

        from src.gui.customers.tabs.sale_history_tab import SaleHistoryTab

        self._history_tab = SaleHistoryTab(
            self._detail_tabs.tab("Sale History"),
            on_refund=self._on_refund_in_till,
        )
        self._history_tab.pack(fill="both", expand=True)

        for name in ("Quotes", "Invoices", "Repairs", "Deposits", "Audit"):
            ctk.CTkLabel(
                self._detail_tabs.tab(name),
                text=f"{name} - coming soon",
                font=ctk.CTkFont(size=13),
                text_color=("gray50", "gray60"),
            ).pack(expand=True)

        self._detail_tabs.configure(command=self._on_detail_tab_change)
        self._ctx_menu = None

    def _on_search_change(self, *_):
        if self._debounce_id:
            self.after_cancel(self._debounce_id)
        self._debounce_id = self.after(300, self._trigger_search)

    def _trigger_search(self):
        self._page = 0
        self._run_search()

    def _run_search(self):
        query = self._search_var.get().strip()
        active_filter = "all"

        sort_col = self._sort_col
        sort_asc = self._sort_asc
        page = self._page

        self._lbl_count.configure(text="Searching...")

        def _thread():
            from src.customers.customer_client import search_customers
            try:
                rows, total = search_customers(
                    query, active_filter, sort_col, sort_asc, page, _PAGE_SIZE
                )
                self.after(0, lambda: self._populate(rows, total))
            except Exception as exc:
                err = str(exc)
                self.after(
                    0,
                    lambda: self._lbl_count.configure(
                        text=f"Error: {err}",
                        text_color=("red", "#e74c3c"),
                    ),
                )

        threading.Thread(target=_thread, daemon=True).start()

    def _populate(self, rows: list, total: int):
        self._total = total
        self._tree.delete(*self._tree.get_children())

        for r in rows:
            self._tree.insert(
                "",
                "end",
                iid=r["id"],
                values=(
                    r.get("customer_code") or r.get("customer_id") or "",
                    r.get("first_name") or "",
                    r.get("surname") or "",
                    r.get("business") or "",
                    r.get("city") or "",
                    r.get("mobile") or r.get("phone_1") or "",
                ),
            )

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

    def _prev_page(self):
        if self._page > 0:
            self._page -= 1
            self._run_search()

    def _next_page(self):
        if (self._page + 1) * _PAGE_SIZE < self._total:
            self._page += 1
            self._run_search()

    def _on_row_select(self, _event=None):
        sel = self._tree.selection()
        if not sel:
            return
        uuid = sel[0]
        self._load_customer_detail(uuid)

    def _on_left_click_release(self, _event=None):
        if not self._merge_source_customer:
            return
        sel = self._tree.selection()
        if not sel:
            return
        target_uuid = sel[0]
        if target_uuid == self._merge_source_customer.get("id"):
            return
        self.after(0, lambda: self._confirm_merge_pair_by_id(target_uuid))

    def _on_double_click(self, _event=None):
        sel = self._tree.selection()
        if not sel:
            return
        if self._merge_source_customer and sel[0] != self._merge_source_customer.get("id"):
            return
        self._open_edit_by_id(sel[0])

    def _load_customer_detail(self, uuid: str):
        def _thread():
            from src.customers.customer_client import get_customer
            try:
                customer = get_customer(uuid)
                self.after(0, lambda: self._show_detail(customer))
            except Exception:
                self.after(0, lambda: None)

        threading.Thread(target=_thread, daemon=True).start()

    def _show_detail(self, customer: Optional[dict]):
        if not customer:
            return
        self._selected_customer = customer

        if not self._detail_frame.winfo_ismapped():
            self._detail_frame.pack(fill="x", side="bottom")

        self._info_tab.load_customer(customer)
        self._history_tab.load_for_customer(customer["id"])
        self._detail_tabs.set("Customer Info")

    def _on_detail_tab_change(self):
        tab = self._detail_tabs.get()
        if tab == "Sale History":
            self._history_tab.on_tab_selected()

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
        menu.add_separator()

        merge_source = self._merge_source_customer
        if not merge_source:
            menu.add_command(label="Merge", command=lambda: self._begin_merge_by_id(row))
        elif merge_source.get("id") == row:
            menu.add_command(label="Merge: choose second profile", state="disabled")
            menu.add_command(label="Cancel Merge", command=self._clear_merge_source)
        else:
            menu.add_command(
                label=f"Merge with {self._customer_ref(merge_source)}",
                command=lambda: self._confirm_merge_pair_by_id(row),
            )
            menu.add_command(label="Cancel Merge", command=self._clear_merge_source)

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

    def _begin_merge_by_id(self, uuid: str) -> None:
        def _thread():
            from src.customers.customer_client import get_customer
            try:
                customer = get_customer(uuid)
                self.after(0, lambda: self._activate_merge_source(customer))
            except Exception as exc:
                err = str(exc)
                self.after(0, lambda: self._show_merge_error(err))

        threading.Thread(target=_thread, daemon=True).start()

    def _activate_merge_source(self, customer: Optional[dict]) -> None:
        from tkinter import messagebox

        if not customer:
            messagebox.showerror(
                "Merge Customers",
                "That customer could not be loaded for merging.",
                parent=self.winfo_toplevel(),
            )
            return

        self._merge_source_customer = customer
        self._lbl_merge.configure(
            text=(
                f"Merge mode: {self._customer_ref(customer)} selected. "
                "Choose the second profile."
            )
        )
        if not self._merge_banner.winfo_ismapped():
            self._merge_banner.pack(side="left", padx=(14, 0), pady=7)

        if customer.get("id"):
            self._tree.selection_set(customer["id"])

    def _clear_merge_source(self) -> None:
        self._merge_source_customer = None
        if self._merge_banner.winfo_ismapped():
            self._merge_banner.pack_forget()

    def _confirm_merge_pair_by_id(self, target_uuid: str) -> None:
        source = self._merge_source_customer
        if not source or target_uuid == source.get("id"):
            return

        source_uuid = source.get("id")

        def _thread():
            from src.customers.customer_client import get_customer
            try:
                target = get_customer(target_uuid)
                self.after(0, lambda: self._confirm_merge_pair(source, target))
            except Exception as exc:
                err = str(exc)
                self.after(0, lambda: self._show_merge_error(err))

        if not source_uuid:
            self._clear_merge_source()
            return

        threading.Thread(target=_thread, daemon=True).start()

    def _confirm_merge_pair(self, customer_a: dict, customer_b: Optional[dict]) -> None:
        from tkinter import messagebox

        if not customer_b:
            messagebox.showerror(
                "Merge Customers",
                "The second customer could not be loaded.",
                parent=self.winfo_toplevel(),
            )
            self._clear_merge_source()
            return

        confirmed = messagebox.askyesno(
            "Merge Customer Profiles",
            (
                f"Merge Customer Profiles {self._customer_ref(customer_a)} "
                f"and {self._customer_ref(customer_b)}?"
            ),
            parent=self.winfo_toplevel(),
        )
        if not confirmed:
            self._clear_merge_source()
            return

        self._clear_merge_source()
        self._open_merge_modal(customer_a, customer_b)

    def _open_merge_modal(self, customer_a: dict, customer_b: dict) -> None:
        from src.gui.customers.customer_merge_modal import CustomerMergeModal

        CustomerMergeModal(
            self.winfo_toplevel(),
            customer_a=customer_a,
            customer_b=customer_b,
            merged_by=self._current_username(),
            on_merged=self._on_customer_merged,
        )

    def _show_merge_error(self, err: str) -> None:
        from tkinter import messagebox

        messagebox.showerror(
            "Merge Customers",
            f"Could not prepare the customer merge:\n{err}",
            parent=self.winfo_toplevel(),
        )

    def _on_customer_merged(self, result: dict) -> None:
        from tkinter import messagebox

        customer = result.get("customer")
        moved_transactions = result.get("moved_transaction_count") or 0
        moved_parked = result.get("moved_parked_count") or 0

        if customer:
            self._show_detail(customer)
        self._run_search()

        if customer:
            messagebox.showinfo(
                "Merge Complete",
                (
                    f"Created merged profile {self._customer_ref(customer)}.\n\n"
                    f"Transactions moved: {moved_transactions}\n"
                    f"Parked transactions updated: {moved_parked}"
                ),
                parent=self.winfo_toplevel(),
            )

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
        self._show_detail(customer)
        self._run_search()

    def show_customer(self, uuid: str) -> None:
        """Navigate to and display the detail panel for the given customer UUID."""
        self._load_customer_detail(uuid)

    def _open_import(self):
        from src.gui.customers.import_dialog import CustomerImportDialog

        CustomerImportDialog(
            self.winfo_toplevel(),
            on_done=self._run_search,
        )

    def _customer_ref(self, customer: dict) -> str:
        return str(customer.get("customer_code") or customer.get("customer_id") or "Unknown")

    def _current_username(self) -> str | None:
        if not self.current_user:
            return None
        return self.current_user.get("username")
