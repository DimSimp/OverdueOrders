from __future__ import annotations

import threading
import tkinter as tk
from tkinter import ttk
from typing import Optional

import customtkinter as ctk


class InventoryTab(ctk.CTkFrame):
    """Inventory browse/search tab for the POS window.

    Layout
    ------
    Top bar:    [Search ________] [Active ▾] [Stock ▾]  [Import CSV]
    ─────────────────────────────────────────────────────────────────
    Grid:       SKU | Title | Brand | On Hand | Available | On Order | Price
                (paginated, sortable column headers)
    Pagination: Showing X-Y of N  [← Prev]  [Page N of M]  [Next →]
    ─────────────────────────────────────────────────────────────────
    Detail:     [Details] [Cust. Orders] [Sale History] [PO History] [Specs]
                (visible when a row is selected; hidden otherwise)
    """

    _PAGE_SIZE = 100
    _TREE_COLS = ("sku", "title", "brand", "on_hand", "available", "on_order", "price")
    _SORT_MAP  = {
        "sku":   "sku",
        "title": "title",
        "brand": "brand",
        "on_hand":   "qty_on_hand",
        "on_order":  "qty_on_order",
        "price": "supplier_rrp",
    }

    def __init__(self, master, current_user=None, **kwargs):
        super().__init__(master, fg_color="transparent", **kwargs)
        self._current_user = current_user
        self._page = 0
        self._total = 0
        self._sort_col = "title"
        self._sort_asc = True
        self._debounce_id = None
        self._selected_id: str = ""
        self._row_data: dict[str, dict] = {}       # item_id → full row (for cart)
        self._auto_select: bool = False            # select first result after next search
        self.on_add_to_cart: Optional[callable] = None  # set by PosWindow
        self._build_ui()

    # ── Build UI ─────────────────────────────────────────────────────────

    def _build_ui(self):
        # ── Top search/filter bar ────────────────────────────────────────
        bar = ctk.CTkFrame(self, fg_color=("gray90", "gray18"), corner_radius=0)
        bar.pack(fill="x", side="top")

        ctk.CTkLabel(bar, text="Search:", font=ctk.CTkFont(size=12)).pack(
            side="left", padx=(14, 4), pady=10)

        self._search_var = ctk.StringVar()
        self._search_var.trace_add("write", self._on_search_changed)
        ctk.CTkEntry(
            bar, textvariable=self._search_var,
            placeholder_text="SKU, title, brand, barcode…",
            width=280, font=ctk.CTkFont(size=12),
        ).pack(side="left", pady=10)

        ctk.CTkLabel(bar, text="Show:", font=ctk.CTkFont(size=12)).pack(
            side="left", padx=(16, 4), pady=10)
        self._active_var = ctk.StringVar(value="Active only")
        ctk.CTkOptionMenu(
            bar,
            variable=self._active_var,
            values=["Active only", "All items", "Inactive only"],
            width=120, font=ctk.CTkFont(size=12),
            command=self._on_filter_changed,
        ).pack(side="left", pady=10)

        ctk.CTkLabel(bar, text="Stock:", font=ctk.CTkFont(size=12)).pack(
            side="left", padx=(12, 4), pady=10)
        self._stock_var = ctk.StringVar(value="All")
        ctk.CTkOptionMenu(
            bar,
            variable=self._stock_var,
            values=["All", "In stock", "Out of stock"],
            width=110, font=ctk.CTkFont(size=12),
            command=self._on_filter_changed,
        ).pack(side="left", pady=10)

        ctk.CTkButton(
            bar, text="Import CSV", width=100, font=ctk.CTkFont(size=12),
            command=self._open_import_dialog,
        ).pack(side="right", padx=14, pady=10)

        self._status_lbl = ctk.CTkLabel(
            bar, text="Search or filter to browse inventory.",
            font=ctk.CTkFont(size=11), text_color=("gray50", "gray60"),
        )
        self._status_lbl.pack(side="right", padx=(0, 8))

        # ── Grid area ────────────────────────────────────────────────────
        grid_frame = ctk.CTkFrame(self, fg_color="transparent")
        grid_frame.pack(fill="both", expand=True)
        grid_frame.grid_rowconfigure(0, weight=1)
        grid_frame.grid_columnconfigure(0, weight=1)

        tree_frame = ctk.CTkFrame(grid_frame, fg_color=("gray95", "gray15"))
        tree_frame.grid(row=0, column=0, sticky="nsew", padx=8, pady=(8, 0))

        self._tree = ttk.Treeview(
            tree_frame, columns=self._TREE_COLS,
            show="headings", selectmode="browse", style="Inv.Treeview",
        )

        col_cfg = {
            "sku":       ("SKU",        120, "w", False),
            "title":     ("Title",      280, "w", True),
            "brand":     ("Brand",      120, "w", False),
            "on_hand":   ("On Hand",     65, "e", False),
            "available": ("Available",   75, "e", False),
            "on_order":  ("On Order",    70, "e", False),
            "price":     ("Price",       80, "e", False),
        }
        for col, (heading, width, anchor, stretch) in col_cfg.items():
            self._tree.heading(col, text=heading,
                               command=lambda c=col: self._on_header_click(c))
            self._tree.column(col, width=width, minwidth=40,
                              anchor=anchor, stretch=stretch)

        vsb = ttk.Scrollbar(tree_frame, orient="vertical", command=self._tree.yview)
        self._tree.configure(yscrollcommand=vsb.set)
        self._tree.pack(side="left", fill="both", expand=True)
        vsb.pack(side="right", fill="y")

        self._tree.bind("<<TreeviewSelect>>", self._on_row_select)
        self._tree.bind("<Button-3>", self._on_right_click)

        # ── Pagination bar ───────────────────────────────────────────────
        pag = ctk.CTkFrame(grid_frame, fg_color=("gray88", "gray20"), height=34,
                           corner_radius=0)
        pag.grid(row=1, column=0, sticky="ew", padx=8)
        pag.pack_propagate(False)

        self._prev_btn = ctk.CTkButton(
            pag, text="← Prev", width=80, height=26, font=ctk.CTkFont(size=12),
            fg_color="transparent", border_width=1,
            border_color=("gray60", "gray45"),
            text_color=("gray40", "gray70"),
            hover_color=("gray85", "gray25"),
            command=self._prev_page, state="disabled",
        )
        self._prev_btn.pack(side="left", padx=(8, 4), pady=4)

        self._page_lbl = ctk.CTkLabel(
            pag, text="", font=ctk.CTkFont(size=11),
            text_color=("gray50", "gray60"),
        )
        self._page_lbl.pack(side="left", padx=4)

        self._next_btn = ctk.CTkButton(
            pag, text="Next →", width=80, height=26, font=ctk.CTkFont(size=12),
            fg_color="transparent", border_width=1,
            border_color=("gray60", "gray45"),
            text_color=("gray40", "gray70"),
            hover_color=("gray85", "gray25"),
            command=self._next_page, state="disabled",
        )
        self._next_btn.pack(side="left", padx=(4, 0), pady=4)

        self._count_lbl = ctk.CTkLabel(
            pag, text="", font=ctk.CTkFont(size=11),
            text_color=("gray50", "gray60"),
        )
        self._count_lbl.pack(side="right", padx=12)

        # ── Detail panel ─────────────────────────────────────────────────
        sep = ctk.CTkFrame(self, height=1, fg_color=("gray70", "gray35"),
                           corner_radius=0)
        sep.pack(fill="x", padx=8)

        self._detail_frame = ctk.CTkFrame(self, fg_color=("gray92", "gray14"),
                                          corner_radius=0)
        self._detail_frame.pack(fill="x", padx=8, pady=(0, 8))

        self._no_sel_lbl = ctk.CTkLabel(
            self._detail_frame,
            text="Select an item to see details.",
            font=ctk.CTkFont(size=12), text_color=("gray50", "gray60"),
        )
        self._no_sel_lbl.pack(pady=12)

        self._detail_tabs: ctk.CTkTabview | None = None

    # ── Search / filter ──────────────────────────────────────────────────

    def _on_search_changed(self, *_):
        if self._debounce_id:
            self.after_cancel(self._debounce_id)
        self._debounce_id = self.after(300, self._reset_and_search)

    def _on_filter_changed(self, *_):
        self._reset_and_search()

    def _reset_and_search(self):
        self._page = 0
        self._do_search()

    def _do_search(self):
        query = self._search_var.get().strip()
        filters = self._get_filters()
        self._status_lbl.configure(text="Searching…", text_color=("gray50", "gray60"))
        threading.Thread(
            target=self._search_thread, args=(query, filters, self._page,
                                              self._sort_col, self._sort_asc),
            daemon=True,
        ).start()

    def _get_filters(self) -> dict:
        active_map = {
            "Active only":   "active",
            "All items":     "all",
            "Inactive only": "inactive",
        }
        stock_map = {
            "All":          "all",
            "In stock":     "in_stock",
            "Out of stock": "out_of_stock",
        }
        return {
            "active": active_map.get(self._active_var.get(), "active"),
            "stock":  stock_map.get(self._stock_var.get(), "all"),
        }

    def _search_thread(self, query, filters, page, sort_col, sort_asc):
        try:
            from src.inventory.inventory_client import search_items
            rows, total = search_items(query, filters, sort_col, sort_asc, page,
                                       self._PAGE_SIZE)
            self.after(0, lambda: self._show_results(rows, total, page))
        except Exception as exc:
            err = str(exc)
            self.after(0, lambda: self._status_lbl.configure(
                text=f"Error: {err}", text_color="red",
            ))

    def _show_results(self, rows: list[dict], total: int, page: int):
        self._total = total
        self._page = page
        self._row_data.clear()

        # Populate treeview
        self._tree.delete(*self._tree.get_children())
        for r in rows:
            self._row_data[r["id"]] = r
            on_hand = r.get("qty_on_hand") or 0
            avail = (on_hand
                     - (r.get("qty_allocated_online") or 0)
                     - (r.get("qty_allocated_customer") or 0))
            on_order = r.get("qty_on_order") or 0
            price = r.get("supplier_rrp")
            price_str = f"${price:.2f}" if price else "—"
            self._tree.insert(
                "", "end",
                iid=r["id"],
                values=(
                    r.get("sku") or "",
                    r.get("title") or "",
                    r.get("brand") or "",
                    on_hand,
                    avail,
                    on_order if on_order else "—",
                    price_str,
                ),
            )

        # Pagination
        total_pages = max(1, -(-total // self._PAGE_SIZE))  # ceiling div
        start = page * self._PAGE_SIZE + 1
        end = min((page + 1) * self._PAGE_SIZE, total)

        if total == 0:
            self._status_lbl.configure(text="No items found.", text_color=("gray50", "gray60"))
            self._page_lbl.configure(text="")
            self._count_lbl.configure(text="")
        else:
            self._status_lbl.configure(
                text=f"{total:,} item{'s' if total != 1 else ''}",
                text_color=("gray50", "gray60"),
            )
            self._page_lbl.configure(text=f"Page {page + 1} of {total_pages}")
            self._count_lbl.configure(text=f"Showing {start:,}–{end:,} of {total:,}")

        self._prev_btn.configure(state="normal" if page > 0 else "disabled")
        self._next_btn.configure(state="normal" if page + 1 < total_pages else "disabled")

        # Clear detail panel, then auto-select first row if requested
        self._selected_id = ""
        self._hide_detail()
        if self._auto_select and rows:
            self._auto_select = False
            first_id = rows[0]["id"]
            self._tree.selection_set(first_id)
            self._tree.see(first_id)
            self._selected_id = first_id
            threading.Thread(target=self._load_detail, args=(first_id,), daemon=True).start()

    # ── Sorting ──────────────────────────────────────────────────────────

    def _on_header_click(self, col: str):
        db_col = self._SORT_MAP.get(col)
        if not db_col:
            return
        if self._sort_col == db_col:
            self._sort_asc = not self._sort_asc
        else:
            self._sort_col = db_col
            self._sort_asc = True
        self._page = 0
        self._do_search()

    # ── Pagination ───────────────────────────────────────────────────────

    def _prev_page(self):
        if self._page > 0:
            self._page -= 1
            self._do_search()

    def _next_page(self):
        total_pages = max(1, -(-self._total // self._PAGE_SIZE))
        if self._page + 1 < total_pages:
            self._page += 1
            self._do_search()

    # ── Detail panel ─────────────────────────────────────────────────────

    def _on_row_select(self, _event):
        sel = self._tree.selection()
        if not sel:
            self._hide_detail()
            return
        item_id = sel[0]
        if item_id == self._selected_id:
            return
        self._selected_id = item_id
        threading.Thread(target=self._load_detail, args=(item_id,), daemon=True).start()

    def _load_detail(self, item_id: str):
        from src.inventory.inventory_client import get_item_by_id
        try:
            item = get_item_by_id(item_id)
            if item:
                self.after(0, lambda: self._show_detail(item))
        except Exception:
            pass

    def _hide_detail(self):
        if self._detail_tabs:
            self._detail_tabs.pack_forget()
        self._no_sel_lbl.pack(pady=12)

    def _show_detail(self, item: dict):
        self._no_sel_lbl.pack_forget()

        # Rebuild tab view each time (simple; acceptable at this stage)
        if self._detail_tabs:
            self._detail_tabs.destroy()

        self._detail_tabs = ctk.CTkTabview(self._detail_frame, anchor="nw", height=220)
        self._detail_tabs.pack(fill="x", padx=4, pady=4)

        for name in ("Details", "Cust. Orders", "Sale History", "PO History", "Specs"):
            self._detail_tabs.add(name)

        # Stub tabs
        for name in ("Cust. Orders", "Sale History", "PO History", "Specs"):
            ctk.CTkLabel(
                self._detail_tabs.tab(name),
                text=f"{name} — coming soon",
                font=ctk.CTkFont(size=12), text_color=("gray50", "gray60"),
            ).pack(expand=True)

        self._build_details_tab(self._detail_tabs.tab("Details"), item)

    def _build_details_tab(self, parent, item: dict):
        """Two-column key/value grid of item fields."""
        scroll = ctk.CTkScrollableFrame(parent, fg_color="transparent")
        scroll.pack(fill="both", expand=True)

        def _val(v):
            if v is None:
                return "—"
            if isinstance(v, bool):
                return "Yes" if v else "No"
            return str(v)

        fields = [
            ("SKU",              item.get("sku")),
            ("Web SKU",          item.get("web_sku")),
            ("Title",            item.get("title")),
            ("Brand",            item.get("brand")),
            ("Series",           item.get("series")),
            ("Instrument",       item.get("instrument")),
            ("Sub-Instrument",   item.get("sub_instrument")),
            ("Supplier",         item.get("supplier_id")),
            ("Pick Zone",        item.get("pick_zone")),
            ("Active",           item.get("active")),
            ("Supplier RRP",     f"${item['supplier_rrp']:.2f}" if item.get("supplier_rrp") else None),
            ("Last Cost",        f"${item['last_purchase_cost']:.2f}" if item.get("last_purchase_cost") else None),
            ("Min Sell",         f"${item['minimum_sell']:.2f}" if item.get("minimum_sell") else None),
            ("On Hand",          item.get("qty_on_hand")),
            ("Alloc. Online",    item.get("qty_allocated_online")),
            ("Alloc. Customer",  item.get("qty_allocated_customer")),
            ("On Order",         item.get("qty_on_order")),
            ("Internal Barcode", item.get("internal_barcode")),
            ("Product Barcode",  item.get("product_barcode")),
            ("Last Purchased",   item.get("last_purchase_date")),
            ("Last Sold",        item.get("last_sold_date")),
            ("Created",          item.get("created_date")),
        ]

        for i, (label, value) in enumerate(fields):
            col = (i % 2) * 2
            row_n = i // 2
            ctk.CTkLabel(
                scroll, text=f"{label}:",
                font=ctk.CTkFont(size=11), text_color=("gray45", "gray65"),
                anchor="e", width=120,
            ).grid(row=row_n, column=col, sticky="e", padx=(8, 4), pady=1)
            ctk.CTkLabel(
                scroll, text=_val(value),
                font=ctk.CTkFont(size=11), anchor="w",
            ).grid(row=row_n, column=col + 1, sticky="w", padx=(0, 20), pady=1)

    # ── Right-click context menu ─────────────────────────────────────────

    def _on_right_click(self, event):
        if not self.on_add_to_cart:
            return
        item_id = self._tree.identify_row(event.y)
        if not item_id:
            return
        self._tree.selection_set(item_id)
        row_data = self._row_data.get(item_id)
        if not row_data:
            return
        menu = tk.Menu(self, tearoff=False)
        menu.add_command(
            label="Add to Cart",
            command=lambda: self.on_add_to_cart(row_data),
        )
        try:
            menu.tk_popup(event.x_root, event.y_root)
        finally:
            menu.grab_release()

    # ── Public API ────────────────────────────────────────────────────────

    def refresh_search(self):
        """Re-run the current search in place (called after a sale to update qty display)."""
        self._do_search()

    def search_and_focus(self, query: str, auto_select: bool = False):
        """Pre-populate the search box and trigger a search (called by Till tab).

        If auto_select is True, the first result will be selected automatically
        so its detail panel opens without the user having to click.
        """
        self._auto_select = auto_select
        self._search_var.set(query)
        # The StringVar trace fires _on_search_changed → debounced _do_search

    # ── Import ────────────────────────────────────────────────────────────

    def _open_import_dialog(self):
        from src.gui.inventory.import_dialog import ImportDialog
        ImportDialog(self.winfo_toplevel(), on_complete=self._on_import_complete)

    def _on_import_complete(self):
        # Re-run current search to reflect newly imported data
        self._do_search()
