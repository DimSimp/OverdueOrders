from __future__ import annotations

import threading
import tkinter as tk
from tkinter import messagebox, ttk
from typing import Optional

import customtkinter as ctk


class TillTab(ctk.CTkFrame):
    """
    POS Till tab — cart entry, totals, and payment.

    Layout
    ------
    Top bar:  [Sale Type ▾]  [Customer ________]          [Park] [Recall]
    ─────────────────────────────────────────────────────────────────────
    Left (70%):   SKU / Barcode: [______________] [Add]
                  Cart treeview (SKU, Description, Qty, Unit Price, Disc%, Line Total)
                  [Remove Line]  [Clear Cart]

    Right (30%):  Subtotal   $0.00
                  Discount   —
                  ───────────────
                  TOTAL      $0.00

                  Cart discount %: [___] [Apply]

                  Payment:
                  [       Cash       ]
                  [       EFT        ]
                  [      Online      ]

                  [    Confirm Sale   ]
    """

    def __init__(self, master, current_user=None, **kwargs):
        super().__init__(master, fg_color="transparent", **kwargs)
        self.current_user = current_user
        # Set by PosWindow after wiring both tabs
        self.navigate_to_inventory: Optional[callable] = None
        self.refresh_inventory: Optional[callable] = None
        # Cart state
        self._cart_items: dict[str, dict] = {}  # item_id → line data
        self._cart_disc_pct: float = 0.0
        self._parked_tx_id: Optional[str] = None
        # Payment state
        self._payment_cash: float = 0.0    # cash tendered (0 = not entered)
        self._payment_eft: list[dict] = [] # [{"amount": X.XX}, ...]
        self._payment_online: float = 0.0
        # Inline payment field widgets (populated in _build_payment_panel)
        self._eft_vars: list[tk.StringVar] = []
        self._eft_row_widgets: list = []
        self._build_ui()

    # ── Build UI ─────────────────────────────────────────────────────────

    def _build_ui(self):
        # ── Top control bar ──────────────────────────────────────────────
        top_bar = ctk.CTkFrame(
            self, height=46, corner_radius=0,
            fg_color=("gray90", "gray18"),
        )
        top_bar.pack(fill="x", side="top")
        top_bar.pack_propagate(False)

        ctk.CTkLabel(
            top_bar, text="Sale type:", font=ctk.CTkFont(size=12),
        ).pack(side="left", padx=(14, 4), pady=10)

        self._sale_type_var = ctk.StringVar(value="Standard")
        ctk.CTkOptionMenu(
            top_bar,
            variable=self._sale_type_var,
            values=["Standard", "Quote", "Invoice", "Repair", "Deposit", "Refund"],
            width=120,
            font=ctk.CTkFont(size=12),
        ).pack(side="left", pady=10)

        ctk.CTkLabel(
            top_bar, text="Customer:", font=ctk.CTkFont(size=12),
        ).pack(side="left", padx=(20, 4), pady=10)

        self._customer_entry = ctk.CTkEntry(
            top_bar,
            placeholder_text="Name / mobile / barcode...",
            width=200,
            font=ctk.CTkFont(size=12),
        )
        self._customer_entry.pack(side="left", pady=10)

        ctk.CTkButton(
            top_bar, text="Recall", width=80, font=ctk.CTkFont(size=12),
            fg_color="transparent", border_width=1,
            border_color=("gray60", "gray45"),
            text_color=("gray40", "gray70"),
            hover_color=("gray85", "gray25"),
            command=self._on_recall,
        ).pack(side="right", padx=(0, 14), pady=10)

        ctk.CTkButton(
            top_bar, text="Park", width=80, font=ctk.CTkFont(size=12),
            fg_color="transparent", border_width=1,
            border_color=("gray60", "gray45"),
            text_color=("gray40", "gray70"),
            hover_color=("gray85", "gray25"),
            command=self._on_park,
        ).pack(side="right", padx=(0, 6), pady=10)

        ctk.CTkButton(
            top_bar, text="Sales", width=80, font=ctk.CTkFont(size=12),
            fg_color="transparent", border_width=1,
            border_color=("gray60", "gray45"),
            text_color=("gray40", "gray70"),
            hover_color=("gray85", "gray25"),
            command=self._on_sales,
        ).pack(side="right", padx=(0, 6), pady=10)

        # ── Main area ────────────────────────────────────────────────────
        main = ctk.CTkFrame(self, fg_color="transparent")
        main.pack(fill="both", expand=True)
        main.grid_columnconfigure(0, weight=7)
        main.grid_columnconfigure(1, weight=3)
        main.grid_rowconfigure(0, weight=1)

        self._build_cart_panel(main)
        self._build_payment_panel(main)

    def _build_cart_panel(self, parent):
        left = ctk.CTkFrame(parent, fg_color="transparent")
        left.grid(row=0, column=0, sticky="nsew", padx=(8, 4), pady=8)
        left.grid_rowconfigure(1, weight=1)
        left.grid_columnconfigure(0, weight=1)

        # SKU search bar
        search_row = ctk.CTkFrame(left, fg_color="transparent")
        search_row.grid(row=0, column=0, sticky="ew", pady=(0, 6))
        search_row.grid_columnconfigure(1, weight=1)

        ctk.CTkLabel(
            search_row, text="SKU / Barcode:", font=ctk.CTkFont(size=13),
        ).grid(row=0, column=0, padx=(0, 8))

        self._sku_entry = ctk.CTkEntry(
            search_row,
            placeholder_text="Scan or type...",
            font=ctk.CTkFont(size=13),
        )
        self._sku_entry.grid(row=0, column=1, sticky="ew")
        self._sku_entry.bind("<Return>", lambda _e: self._add_item_from_entry())

        self._add_btn = ctk.CTkButton(
            search_row, text="Add", width=64,
            command=self._add_item_from_entry,
        )
        self._add_btn.grid(row=0, column=2, padx=(6, 0))

        # Cart treeview
        cart_frame = ctk.CTkFrame(left, fg_color=("gray95", "gray15"))
        cart_frame.grid(row=1, column=0, sticky="nsew")

        cols = ("sku", "description", "qty", "unit_price", "disc_pct", "line_total", "margin")
        self._tree = ttk.Treeview(
            cart_frame, columns=cols, show="headings",
            selectmode="browse", style="POS.Treeview",
        )

        self._tree.heading("sku",         text="SKU")
        self._tree.heading("description", text="Description")
        self._tree.heading("qty",         text="Qty",        anchor="e")
        self._tree.heading("unit_price",  text="Unit Price", anchor="e")
        self._tree.heading("disc_pct",    text="Disc %",     anchor="e")
        self._tree.heading("line_total",  text="Line Total", anchor="e")
        self._tree.heading("margin",      text="Margin",     anchor="e")

        self._tree.column("sku",         width=110, minwidth=80,  anchor="w", stretch=False)
        self._tree.column("description", width=260, minwidth=120, anchor="w")
        self._tree.column("qty",         width=60,  minwidth=40,  anchor="e", stretch=False)
        self._tree.column("unit_price",  width=90,  minwidth=70,  anchor="e", stretch=False)
        self._tree.column("disc_pct",    width=60,  minwidth=40,  anchor="e", stretch=False)
        self._tree.column("line_total",  width=100, minwidth=70,  anchor="e", stretch=False)
        self._tree.column("margin",      width=70,  minwidth=50,  anchor="e", stretch=False)

        vsb = ttk.Scrollbar(cart_frame, orient="vertical", command=self._tree.yview)
        self._tree.configure(yscrollcommand=vsb.set)
        self._tree.tag_configure("margin_green",  foreground="#22c55e")
        self._tree.tag_configure("margin_orange", foreground="#f59e0b")
        self._tree.tag_configure("margin_red",    foreground="#ef4444")
        self._tree.pack(side="left", fill="both", expand=True)
        vsb.pack(side="right", fill="y")
        self._tree.bind("<ButtonRelease-1>", self._on_cell_click)
        self._tree.bind("<Button-3>", self._show_context_menu)

        # Cart action buttons
        actions = ctk.CTkFrame(left, fg_color="transparent")
        actions.grid(row=2, column=0, sticky="ew", pady=(6, 0))

        ctk.CTkButton(
            actions, text="Remove Line", width=110,
            fg_color=("gray75", "gray30"),
            command=self._remove_selected_line,
        ).pack(side="left")

        ctk.CTkButton(
            actions, text="Clear Cart", width=100,
            fg_color=("gray75", "gray30"),
            command=self._clear_cart,
        ).pack(side="left", padx=(6, 0))

    def _build_payment_panel(self, parent):
        right = ctk.CTkFrame(parent, fg_color=("gray90", "gray15"), corner_radius=8)
        right.grid(row=0, column=1, sticky="nsew", padx=(4, 8), pady=8)

        # Push content to the bottom — spacer absorbs all free space at the top
        ctk.CTkFrame(right, fg_color="transparent").pack(expand=True)

        # ── Cart breakdown section cap ────────────────────────────────────
        ctk.CTkFrame(right, fg_color=("gray60", "gray35"), height=2).pack(
            fill="x", padx=0, pady=(0, 0),
        )

        # Cart discount % control — top of breakdown section
        disc_row = ctk.CTkFrame(right, fg_color="transparent")
        disc_row.pack(fill="x", padx=16, pady=(10, 0))

        ctk.CTkLabel(
            disc_row, text="Cart discount %:", font=ctk.CTkFont(size=12),
            text_color=("gray40", "gray70"),
        ).pack(side="left")

        self._disc_entry = ctk.CTkEntry(disc_row, width=64, font=ctk.CTkFont(size=12),
                                         placeholder_text="0")
        self._disc_entry.pack(side="left", padx=(6, 0))
        self._disc_entry.bind("<Return>", lambda _e: self._apply_discount())

        ctk.CTkButton(
            disc_row, text="Apply", width=60, font=ctk.CTkFont(size=12),
            command=self._apply_discount,
        ).pack(side="left", padx=(4, 0))

        # Totals — Subtotal → Discount → Cart Margin → divider → TOTAL
        totals = ctk.CTkFrame(right, fg_color="transparent")
        totals.pack(fill="x", padx=16, pady=(10, 0))

        def _total_row(label_text: str, attr: str):
            row = ctk.CTkFrame(totals, fg_color="transparent")
            row.pack(fill="x", pady=2)
            ctk.CTkLabel(
                row, text=label_text, font=ctk.CTkFont(size=13),
                text_color=("gray40", "gray70"), anchor="w",
            ).pack(side="left")
            lbl = ctk.CTkLabel(row, text="$0.00", font=ctk.CTkFont(size=13), anchor="e")
            lbl.pack(side="right")
            setattr(self, attr, lbl)

        _total_row("Subtotal", "_lbl_subtotal")
        _total_row("Discount", "_lbl_discount")

        # Cart margin row (staff-only — never shown on customer receipts)
        margin_row = ctk.CTkFrame(totals, fg_color="transparent")
        margin_row.pack(fill="x", pady=(2, 0))
        ctk.CTkLabel(
            margin_row, text="Cart Margin",
            font=ctk.CTkFont(size=12),
            text_color=("gray40", "gray70"), anchor="w",
        ).pack(side="left")
        self._lbl_margin = ctk.CTkLabel(
            margin_row, text="—",
            font=ctk.CTkFont(size=13, weight="bold"),
            text_color=("gray50", "gray60"),
            anchor="e",
        )
        self._lbl_margin.pack(side="right")

        ctk.CTkFrame(totals, height=1, fg_color=("gray70", "gray40")).pack(fill="x", pady=8)

        # TOTAL — largest element, at the bottom of the breakdown
        total_row = ctk.CTkFrame(totals, fg_color="transparent")
        total_row.pack(fill="x")
        ctk.CTkLabel(
            total_row, text="TOTAL",
            font=ctk.CTkFont(size=18, weight="bold"), anchor="w",
        ).pack(side="left")
        self._lbl_total = ctk.CTkLabel(
            total_row, text="$0.00",
            font=ctk.CTkFont(size=22, weight="bold"), anchor="e",
        )
        self._lbl_total.pack(side="right")
        self._lbl_total.bind("<Button-1>", self._on_total_click)
        self._lbl_total.configure(cursor="hand2")

        # Hidden entry — swaps in over _lbl_total when the user clicks it
        self._entry_total = ctk.CTkEntry(
            total_row,
            font=ctk.CTkFont(size=20, weight="bold"),
            width=130, justify="right",
        )
        self._entry_total.bind("<Return>", self._commit_total_edit)
        self._entry_total.bind("<Tab>",    self._commit_total_edit)
        self._entry_total.bind("<FocusOut>", self._commit_total_edit)
        self._entry_total.bind("<Escape>", self._cancel_total_edit)

        # ── Payment section ───────────────────────────────────────────────
        ctk.CTkFrame(right, fg_color=("gray60", "gray35"), height=2).pack(
            fill="x", padx=0, pady=(20, 0),
        )
        ctk.CTkLabel(
            right, text="Payment Method",
            font=ctk.CTkFont(size=13, weight="bold"),
            text_color=("gray30", "gray80"), anchor="w",
        ).pack(fill="x", padx=16, pady=(8, 8))

        # 3-column payment grid
        cols = ctk.CTkFrame(right, fg_color="transparent")
        cols.pack(fill="x", padx=16, pady=(0, 4))
        cols.grid_columnconfigure(0, weight=1)
        cols.grid_columnconfigure(1, weight=1)
        cols.grid_columnconfigure(2, weight=1)

        # ── EFT column ────────────────────────────────────────────────────
        eft_col = ctk.CTkFrame(cols, fg_color="transparent")
        eft_col.grid(row=0, column=0, sticky="nsew", padx=(0, 4))

        ctk.CTkLabel(
            eft_col, text="EFT",
            font=ctk.CTkFont(size=11, weight="bold"),
            text_color=("gray40", "gray70"), anchor="w",
        ).pack(anchor="w")

        self._eft_frame = ctk.CTkFrame(eft_col, fg_color="transparent")
        self._eft_frame.pack(fill="x", pady=(4, 0))
        self._add_eft_row()  # start with one empty row

        # [+] button always lives below the last EFT row
        self._btn_add_eft = ctk.CTkButton(
            self._eft_frame, text="+", width=22, height=22,
            font=ctk.CTkFont(size=13, weight="bold"),
            fg_color="transparent", border_width=1,
            border_color=("gray60", "gray45"),
            text_color=("gray40", "gray70"),
            hover_color=("gray85", "gray25"),
            command=self._add_eft_row,
        )
        self._btn_add_eft.pack(anchor="w", pady=(3, 0))

        # ── Cash column ───────────────────────────────────────────────────
        cash_col = ctk.CTkFrame(cols, fg_color="transparent")
        cash_col.grid(row=0, column=1, sticky="nsew", padx=(4, 4))

        ctk.CTkLabel(
            cash_col, text="Cash",
            font=ctk.CTkFont(size=11, weight="bold"),
            text_color=("gray40", "gray70"), anchor="w",
        ).pack(anchor="w")

        self._cash_var = tk.StringVar()
        ctk.CTkEntry(
            cash_col, textvariable=self._cash_var,
            width=70, font=ctk.CTkFont(size=12), justify="right",
            placeholder_text="0.00",
        ).pack(anchor="w", pady=(4, 0))
        self._cash_change_lbl = ctk.CTkLabel(
            cash_col, text="",
            font=ctk.CTkFont(size=10),
            text_color=("#1a6b2e", "#22c55e"), anchor="w",
        )
        self._cash_change_lbl.pack(anchor="w", pady=(2, 0))
        self._cash_var.trace_add("write", self._recalc_payments)

        # ── Online column ─────────────────────────────────────────────────
        online_col = ctk.CTkFrame(cols, fg_color="transparent")
        online_col.grid(row=0, column=2, sticky="nsew", padx=(4, 0))

        ctk.CTkLabel(
            online_col, text="Online",
            font=ctk.CTkFont(size=11, weight="bold"),
            text_color=("gray40", "gray70"), anchor="w",
        ).pack(anchor="w")

        self._online_var = tk.StringVar()
        ctk.CTkEntry(
            online_col, textvariable=self._online_var,
            width=70, font=ctk.CTkFont(size=12), justify="right",
            placeholder_text="0.00",
        ).pack(anchor="w", pady=(4, 0))
        self._online_var.trace_add("write", self._recalc_payments)

        # Remaining / Change / Paid status
        ctk.CTkFrame(right, fg_color=("gray70", "gray40"), height=1).pack(
            fill="x", padx=16, pady=(10, 4),
        )
        self._remaining_lbl = ctk.CTkLabel(
            right, text="",
            font=ctk.CTkFont(size=13, weight="bold"),
            anchor="e",
        )
        self._remaining_lbl.pack(fill="x", padx=16, pady=(0, 2))

        self._btn_confirm = ctk.CTkButton(
            right,
            text="Confirm Sale",
            height=56,
            font=ctk.CTkFont(size=16, weight="bold"),
            fg_color=("#1a6b2e", "#145722"),
            hover_color=("#145722", "#0d3d18"),
            command=self._confirm_sale,
        )
        self._btn_confirm.pack(fill="x", padx=16, pady=(0, 16))

    # ── SKU lookup ────────────────────────────────────────────────────────

    _SPINNER_FRAMES = ("◐", "◓", "◑", "◒")

    def _add_item_from_entry(self):
        query = self._sku_entry.get().strip()
        if not query:
            return
        self._sku_entry.configure(state="disabled")
        self._start_spinner()
        threading.Thread(target=self._lookup_thread, args=(query,), daemon=True).start()

    def _start_spinner(self):
        self._spinner_idx = 0
        self._spinner_after_id: Optional[str] = None
        self._add_btn.configure(text=self._SPINNER_FRAMES[0], state="disabled")
        self._tick_spinner()

    def _tick_spinner(self):
        self._spinner_idx = (self._spinner_idx + 1) % len(self._SPINNER_FRAMES)
        self._add_btn.configure(text=self._SPINNER_FRAMES[self._spinner_idx])
        self._spinner_after_id = self.after(100, self._tick_spinner)

    def _stop_spinner(self):
        if getattr(self, "_spinner_after_id", None):
            self.after_cancel(self._spinner_after_id)
            self._spinner_after_id = None
        if hasattr(self, "_add_btn"):
            self._add_btn.configure(text="Add", state="normal")

    def _lookup_thread(self, query: str):
        from src.inventory.inventory_client import lookup_exact
        try:
            matches = lookup_exact(query)
            self.after(0, lambda: self._handle_lookup_result(query, matches))
        except Exception as exc:
            err = str(exc)
            self.after(0, lambda: self._on_lookup_error(err))

    def _handle_lookup_result(self, query: str, matches: list[dict]):
        self._stop_spinner()
        self._sku_entry.configure(state="normal")
        self._sku_entry.delete(0, "end")
        if len(matches) == 1:
            self.add_item_to_cart(matches[0])
            self._sku_entry.focus()
        else:
            # 0 matches → fuzzy search; multiple matches → let user pick
            if self.navigate_to_inventory:
                self.navigate_to_inventory(query)

    def _on_lookup_error(self, err: str):
        self._stop_spinner()
        self._sku_entry.configure(state="normal")
        messagebox.showerror("Lookup Error", f"Item lookup failed:\n{err}",
                             parent=self.winfo_toplevel())

    # ── Cart management ───────────────────────────────────────────────────

    def add_item_to_cart(self, item: dict):
        """Public API — add an item to the cart (from lookup or Inventory tab)."""
        item_id = item.get("id")
        if not item_id:
            return

        # Out-of-stock check
        qty_avail = (
            (item.get("qty_on_hand") or 0)
            - (item.get("qty_allocated_online") or 0)
            - (item.get("qty_allocated_customer") or 0)
        )
        if qty_avail <= 0:
            if not self._warn_out_of_stock(item):
                return

        # Serialised warning (informational — still adds)
        if item.get("is_serialised"):
            self._warn_serialised(item)

        unit_price = float(item.get("online_sale_price") or item.get("supplier_rrp") or 0)
        _cost = item.get("average_cost_exc_gst") or item.get("last_purchase_cost")
        cost_price = float(_cost) if _cost else None

        if item_id in self._cart_items:
            self._cart_items[item_id]["qty"] += 1
            self._refresh_tree_row(item_id)
        else:
            self._cart_items[item_id] = {
                "sku":        item.get("sku") or "",
                "title":      item.get("title") or "",
                "qty":        1,
                "unit_price": unit_price,
                "disc_pct":   0.0,
                "cost_price": cost_price,
            }
            self._insert_tree_row(item_id)

        self._update_totals()

    def _on_cell_click(self, event):
        """Inline-edit Qty (#3), Unit Price (#4), or Disc % (#5) on single click."""
        row_id = self._tree.identify_row(event.y)
        col = self._tree.identify_column(event.x)
        if not row_id or col not in ("#3", "#4", "#5", "#6"):
            return
        bbox = self._tree.bbox(row_id, col)
        if not bbox:
            return
        x, y, width, height = bbox

        line = self._cart_items.get(row_id)
        if not line:
            return

        if col == "#3":
            current = str(line["qty"])
        elif col == "#4":
            current = f"{line['unit_price']:.2f}"
        elif col == "#5":
            current = f"{line['disc_pct']:.1f}"
        else:  # "#6" — line total
            lt = line["qty"] * line["unit_price"] * (1 - line["disc_pct"] / 100)
            current = f"{lt:.2f}"

        var = tk.StringVar(value=current)
        entry = tk.Entry(
            self._tree, textvariable=var,
            justify="center", font=("Segoe UI", 11),
            relief="flat", bd=0,
            bg="#2a2a2a", fg="#e0e0e0",
            insertbackground="#e0e0e0",
            selectbackground="#1f538d", selectforeground="white",
        )
        entry.place(x=x, y=y, width=width, height=height)
        entry.select_range(0, "end")
        entry.focus_set()

        _done = [False]

        def _commit(_event=None):
            if _done[0]:
                return
            _done[0] = True
            try:
                if col == "#3":
                    new_val = int(var.get().strip())
                    if new_val < 1:
                        raise ValueError
                    line["qty"] = new_val
                elif col == "#4":
                    new_val = float(var.get().strip().lstrip("$"))
                    if new_val < 0:
                        raise ValueError
                    line["unit_price"] = round(new_val, 2)
                elif col == "#5":
                    new_val = float(var.get().strip().rstrip("%"))
                    if not (0 <= new_val <= 100):
                        raise ValueError
                    line["disc_pct"] = round(new_val, 2)
                else:  # "#6" — line total → back-calculate disc_pct
                    new_lt = float(var.get().strip().lstrip("$"))
                    if new_lt < 0:
                        raise ValueError
                    full = line["qty"] * line["unit_price"]
                    if full > 0:
                        disc = (1 - new_lt / full) * 100
                        line["disc_pct"] = round(max(0.0, min(100.0, disc)), 2)
                self._refresh_tree_row(row_id)
                self._update_totals()
            except ValueError:
                pass    # invalid input — leave cart unchanged
            finally:
                entry.destroy()

        def _cancel(_event=None):
            if _done[0]:
                return
            _done[0] = True
            entry.destroy()

        entry.bind("<Return>", _commit)
        entry.bind("<Tab>", _commit)
        entry.bind("<FocusOut>", _commit)
        entry.bind("<Escape>", _cancel)

    def _calc_margin(self, line: dict) -> Optional[float]:
        """Gross margin % = (sell ex-GST − cost ex-GST) / sell ex-GST × 100.

        Returns None if cost data is unavailable or sell price is zero.
        Prices in the system are GST-inclusive; dividing by 1.1 gives ex-GST.
        """
        cost = line.get("cost_price")
        if not cost or cost <= 0:
            return None
        line_total = line["qty"] * line["unit_price"] * (1 - line["disc_pct"] / 100)
        sell_ex_gst = line_total / 1.1
        if sell_ex_gst <= 0:
            return None
        total_cost = line["qty"] * cost
        return (sell_ex_gst - total_cost) / sell_ex_gst * 100

    def _margin_tag(self, margin: Optional[float]) -> str:
        """Return treeview tag name for the margin %, or '' if unavailable."""
        if margin is None:
            return ""
        m = round(margin, 1)
        if m > 10.0:
            return "margin_green"
        elif m == 10.0:
            return "margin_orange"
        else:
            return "margin_red"

    def _insert_tree_row(self, item_id: str):
        line = self._cart_items[item_id]
        line_total = line["qty"] * line["unit_price"]
        margin = self._calc_margin(line)
        margin_str = f"{margin:.1f}%" if margin is not None else "—"
        tag = self._margin_tag(margin)
        self._tree.insert(
            "", "end", iid=item_id,
            values=(
                line["sku"],
                line["title"],
                line["qty"],
                f"${line['unit_price']:.2f}",
                "—",
                f"${line_total:.2f}",
                margin_str,
            ),
            tags=(tag,) if tag else (),
        )

    def _refresh_tree_row(self, item_id: str):
        line = self._cart_items[item_id]
        line_total = line["qty"] * line["unit_price"] * (1 - line["disc_pct"] / 100)
        disc_str = f"{line['disc_pct']:.0f}%" if line["disc_pct"] else "—"
        margin = self._calc_margin(line)
        margin_str = f"{margin:.1f}%" if margin is not None else "—"
        tag = self._margin_tag(margin)
        self._tree.item(
            item_id,
            values=(
                line["sku"],
                line["title"],
                line["qty"],
                f"${line['unit_price']:.2f}",
                disc_str,
                f"${line_total:.2f}",
                margin_str,
            ),
            tags=(tag,) if tag else (),
        )

    def _remove_selected_line(self):
        sel = self._tree.selection()
        if sel:
            item_id = sel[0]
            self._tree.delete(item_id)
            self._cart_items.pop(item_id, None)
            self._update_totals()

    def _show_context_menu(self, event):
        """Right-click context menu for a cart row."""
        row_id = self._tree.identify_row(event.y)
        if not row_id or row_id not in self._cart_items:
            return
        self._tree.selection_set(row_id)

        menu = tk.Menu(self, tearoff=0)
        menu.add_command(
            label="Show in Inventory",
            command=lambda: self._ctx_show_in_inventory(row_id),
        )
        menu.add_command(
            label="Remove from Cart",
            command=lambda: self._ctx_remove_line(row_id),
        )
        try:
            menu.tk_popup(event.x_root, event.y_root)
        finally:
            menu.grab_release()

    def _ctx_show_in_inventory(self, row_id: str):
        sku = self._cart_items[row_id]["sku"]
        if self.navigate_to_inventory:
            self.navigate_to_inventory(sku, auto_select=True)

    def _ctx_remove_line(self, row_id: str):
        self._tree.delete(row_id)
        self._cart_items.pop(row_id, None)
        self._update_totals()

    def _clear_cart(self):
        self._tree.delete(*self._tree.get_children())
        self._cart_items.clear()
        self._cart_disc_pct = 0.0
        self._disc_entry.delete(0, "end")
        self._payment_cash = 0.0
        self._payment_eft = []
        self._payment_online = 0.0
        # Clear inline payment fields
        if hasattr(self, "_cash_var"):
            self._cash_var.set("")
        if hasattr(self, "_online_var"):
            self._online_var.set("")
        # Reset EFT to one empty row
        if hasattr(self, "_eft_frame"):
            for row in list(self._eft_row_widgets):
                row.destroy()
            self._eft_vars.clear()
            self._eft_row_widgets.clear()
            self._add_eft_row()
        self._parked_tx_id = None
        if hasattr(self, "_customer_entry"):
            self._customer_entry.delete(0, "end")
        self._update_totals()

    def _apply_discount(self):
        val = self._disc_entry.get().strip().rstrip("%")
        try:
            pct = float(val)
            if not (0 <= pct <= 100):
                raise ValueError
            self._cart_disc_pct = pct
            self._update_totals()
        except ValueError:
            pass

    def _on_total_click(self, _event=None):
        """Swap TOTAL label → editable entry for a direct total override."""
        if not self._cart_items:
            return
        subtotal = sum(
            line["qty"] * line["unit_price"] * (1 - line["disc_pct"] / 100)
            for line in self._cart_items.values()
        )
        if subtotal <= 0:
            return
        current_total = subtotal * (1 - self._cart_disc_pct / 100)
        self._lbl_total.pack_forget()
        self._entry_total.delete(0, "end")
        self._entry_total.insert(0, f"{current_total:.2f}")
        self._entry_total.pack(side="right")
        self._entry_total.select_range(0, "end")
        self._entry_total.focus_set()
        self._total_edit_active = True

    def _commit_total_edit(self, _event=None):
        if not getattr(self, "_total_edit_active", False):
            return
        self._total_edit_active = False
        raw = self._entry_total.get().strip().lstrip("$")
        self._entry_total.pack_forget()
        self._lbl_total.pack(side="right")
        try:
            new_total = float(raw)
            if new_total < 0:
                raise ValueError
            subtotal = sum(
                line["qty"] * line["unit_price"] * (1 - line["disc_pct"] / 100)
                for line in self._cart_items.values()
            )
            if subtotal > 0:
                disc = (1 - new_total / subtotal) * 100
                self._cart_disc_pct = round(max(0.0, min(100.0, disc)), 2)
                # Keep the Cart Discount % field in sync
                self._disc_entry.delete(0, "end")
                if self._cart_disc_pct:
                    self._disc_entry.insert(0, f"{self._cart_disc_pct:.1f}")
                self._update_totals()
        except ValueError:
            pass

    def _cancel_total_edit(self, _event=None):
        if not getattr(self, "_total_edit_active", False):
            return
        self._total_edit_active = False
        self._entry_total.pack_forget()
        self._lbl_total.pack(side="right")

    def _update_totals(self):
        subtotal = sum(
            line["qty"] * line["unit_price"] * (1 - line["disc_pct"] / 100)
            for line in self._cart_items.values()
        )
        discount_amt = subtotal * (self._cart_disc_pct / 100)
        total = subtotal - discount_amt

        self._lbl_subtotal.configure(text=f"${subtotal:,.2f}")
        if self._cart_disc_pct:
            self._lbl_discount.configure(
                text=f"-${discount_amt:,.2f} ({self._cart_disc_pct:.0f}%)"
            )
        else:
            self._lbl_discount.configure(text="—")
        self._lbl_total.configure(text=f"${total:,.2f}")

        # Cart margin (staff-only)
        total_cost = sum(
            line["qty"] * (line.get("cost_price") or 0)
            for line in self._cart_items.values()
        )
        has_cost = any(line.get("cost_price") for line in self._cart_items.values())
        total_ex_gst = total / 1.1 if total > 0 else 0
        if has_cost and total_ex_gst > 0:
            cart_margin = (total_ex_gst - total_cost) / total_ex_gst * 100
            m = round(cart_margin, 1)
            if m > 10.0:
                color = "#22c55e"
            elif m == 10.0:
                color = "#f59e0b"
            else:
                color = "#ef4444"
            self._lbl_margin.configure(text=f"{cart_margin:.1f}%", text_color=color)
        else:
            self._lbl_margin.configure(text="—", text_color=("gray50", "gray60"))

        self._update_remaining_display()

    # ── Warnings ─────────────────────────────────────────────────────────

    def _warn_out_of_stock(self, item: dict) -> bool:
        """Modal out-of-stock warning. Returns True if user chose Allow."""
        result = [False]
        dlg = ctk.CTkToplevel(self.winfo_toplevel())
        dlg.title("Out of Stock")
        dlg.geometry("380x160")
        dlg.resizable(False, False)
        dlg.grab_set()
        dlg.transient(self.winfo_toplevel())

        name = (item.get("title") or item.get("sku") or "This item")[:60]
        ctk.CTkLabel(
            dlg,
            text=f"{name}\nhas no available stock.",
            font=ctk.CTkFont(size=13),
            justify="center",
        ).pack(expand=True, pady=(20, 8))

        btn_row = ctk.CTkFrame(dlg, fg_color="transparent")
        btn_row.pack(pady=(0, 16))

        def _allow():
            result[0] = True
            dlg.destroy()

        ctk.CTkButton(
            btn_row, text="Allow", width=110, command=_allow,
        ).pack(side="left", padx=6)
        ctk.CTkButton(
            btn_row, text="Remove", width=110,
            fg_color=("gray70", "gray30"),
            command=dlg.destroy,
        ).pack(side="left", padx=6)

        dlg.wait_window()
        return result[0]

    def _warn_serialised(self, item: dict):
        """Informational dialog for serialised items."""
        name = (item.get("title") or item.get("sku") or "This item")[:60]
        messagebox.showinfo(
            "Serialised Item",
            f"{name}\n\nThis is a serialised item — please record the serial number separately.",
            parent=self.winfo_toplevel(),
        )

    # ── Payment helpers ────────────────────────────────────────────────────

    def _get_subtotal(self) -> float:
        return sum(
            line["qty"] * line["unit_price"] * (1 - line["disc_pct"] / 100)
            for line in self._cart_items.values()
        )

    def _get_total(self) -> float:
        return self._get_subtotal() * (1 - self._cart_disc_pct / 100)

    def _amount_remaining(self) -> float:
        total = self._get_total()
        paid = (
            self._payment_cash
            + sum(e["amount"] for e in self._payment_eft)
            + self._payment_online
        )
        return max(0.0, total - paid)

    _MAX_EFT_ROWS = 3

    def _add_eft_row(self):
        """Append a new empty EFT amount entry row (max 3)."""
        if len(self._eft_row_widgets) >= self._MAX_EFT_ROWS:
            return
        var = tk.StringVar()
        var.trace_add("write", self._recalc_payments)
        self._eft_vars.append(var)

        row = ctk.CTkFrame(self._eft_frame, fg_color="transparent")
        if hasattr(self, "_btn_add_eft"):
            row.pack(fill="x", pady=(0, 3), before=self._btn_add_eft)
        else:
            row.pack(fill="x", pady=(0, 3))
        self._eft_row_widgets.append(row)

        ctk.CTkEntry(
            row, textvariable=var,
            width=52, font=ctk.CTkFont(size=12), justify="right",
            placeholder_text="0.00",
        ).pack(side="left")

        ctk.CTkButton(
            row, text="−", width=20, height=20,
            font=ctk.CTkFont(size=12),
            fg_color=("gray70", "gray30"),
            hover_color=("gray60", "gray40"),
            command=lambda r=row: self._remove_eft_row(r),
        ).pack(side="left", padx=(3, 0))

        # Hide [+] when the limit is reached
        if len(self._eft_row_widgets) >= self._MAX_EFT_ROWS and hasattr(self, "_btn_add_eft"):
            self._btn_add_eft.pack_forget()

        self._refresh_eft_minus_buttons()

    def _remove_eft_row(self, row: ctk.CTkFrame):
        """Remove an EFT row (minimum one row always remains)."""
        if len(self._eft_row_widgets) <= 1:
            return
        idx = self._eft_row_widgets.index(row)
        self._eft_vars.pop(idx)
        self._eft_row_widgets.pop(idx)
        row.destroy()
        # Re-show [+] if we're now below the limit
        if len(self._eft_row_widgets) < self._MAX_EFT_ROWS and hasattr(self, "_btn_add_eft"):
            self._btn_add_eft.pack(anchor="w", pady=(3, 0))
        self._refresh_eft_minus_buttons()
        self._recalc_payments()

    def _refresh_eft_minus_buttons(self):
        """Show [−] buttons only when more than one EFT row exists."""
        show = len(self._eft_row_widgets) > 1
        for row in self._eft_row_widgets:
            children = row.winfo_children()
            if len(children) >= 2:
                btn = children[1]
                if show:
                    btn.pack(side="left", padx=(4, 0))
                else:
                    btn.pack_forget()

    def _recalc_payments(self, *_):
        """Read inline field values → update payment state vars → refresh display."""
        try:
            raw = self._cash_var.get().strip()
            self._payment_cash = float(raw) if raw else 0.0
        except ValueError:
            self._payment_cash = 0.0

        self._payment_eft = []
        for var in self._eft_vars:
            raw = var.get().strip()
            if raw:
                try:
                    amt = float(raw)
                    if amt > 0:
                        self._payment_eft.append({"amount": amt})
                except ValueError:
                    pass

        try:
            raw = self._online_var.get().strip()
            self._payment_online = float(raw) if raw else 0.0
        except ValueError:
            self._payment_online = 0.0

        self._update_remaining_display()

    def _update_remaining_display(self):
        """Update the cash change hint and the remaining/change/paid status label."""
        if not hasattr(self, "_remaining_lbl"):
            return
        total = self._get_total()

        # Cash change hint (next to the cash field)
        if hasattr(self, "_cash_change_lbl"):
            if self._payment_cash > 0 and total > 0:
                eft_total = sum(e["amount"] for e in self._payment_eft)
                cash_needed = max(0.0, total - eft_total - self._payment_online)
                change = self._payment_cash - cash_needed
                if change > 0.005:
                    self._cash_change_lbl.configure(
                        text=f"Change ${change:.2f}",
                        text_color=("#1a6b2e", "#22c55e"),
                    )
                else:
                    self._cash_change_lbl.configure(text="")
            else:
                self._cash_change_lbl.configure(text="")

        if total <= 0:
            self._remaining_lbl.configure(text="")
            return

        paid = (
            self._payment_cash
            + sum(e["amount"] for e in self._payment_eft)
            + self._payment_online
        )
        remaining = total - paid

        if remaining > 0.005:
            self._remaining_lbl.configure(
                text=f"Remaining  ${remaining:.2f}",
                text_color=("#c0392b", "#e74c3c"),
            )
        elif remaining < -0.005:
            self._remaining_lbl.configure(
                text=f"Change due  ${-remaining:.2f}",
                text_color=("#1a6b2e", "#22c55e"),
            )
        else:
            self._remaining_lbl.configure(
                text="Paid in full  ✓",
                text_color=("#1a6b2e", "#22c55e"),
            )

    # ── Confirm Sale ──────────────────────────────────────────────────────

    def _confirm_sale(self):
        if not self._cart_items:
            messagebox.showerror("Empty Cart", "No items in cart.", parent=self.winfo_toplevel())
            return

        remaining = self._amount_remaining()
        if remaining > 0.005:
            messagebox.showerror(
                "Payment Incomplete",
                f"${remaining:.2f} still needs to be paid.",
                parent=self.winfo_toplevel(),
            )
            return

        total = self._get_total()
        subtotal = self._get_subtotal()
        eft_total = sum(e["amount"] for e in self._payment_eft)
        cash_required = max(0.0, total - eft_total - self._payment_online)
        change_given = max(0.0, self._payment_cash - cash_required)

        performed_by = ""
        if self.current_user:
            performed_by = (
                self.current_user.get("username", "")
                or f"{self.current_user.get('first_name', '')} "
                   f"{self.current_user.get('last_name', '')}".strip()
            )

        # Snapshot state before handing off to the thread
        cart_snapshot = {k: dict(v) for k, v in self._cart_items.items()}
        eft_snapshot = list(self._payment_eft)
        parked_tx_id = self._parked_tx_id

        self._btn_confirm.configure(state="disabled", text="Processing…")

        def _thread():
            try:
                if parked_tx_id:
                    from src.pos.transaction_client import complete_parked_sale
                    tx = complete_parked_sale(
                        parked_tx_id=parked_tx_id,
                        cart_items=cart_snapshot,
                        subtotal=subtotal,
                        cart_disc_pct=self._cart_disc_pct,
                        total=total,
                        payment_cash=cash_required if self._payment_cash else 0.0,
                        payment_eft=eft_snapshot,
                        payment_online=self._payment_online,
                        cash_tendered=self._payment_cash,
                        change_given=change_given,
                        performed_by=performed_by,
                    )
                else:
                    from src.pos.transaction_client import confirm_standard_sale
                    tx = confirm_standard_sale(
                        cart_items=cart_snapshot,
                        subtotal=subtotal,
                        cart_disc_pct=self._cart_disc_pct,
                        total=total,
                        payment_cash=cash_required if self._payment_cash else 0.0,
                        payment_eft=eft_snapshot,
                        payment_online=self._payment_online,
                        cash_tendered=self._payment_cash,
                        change_given=change_given,
                        performed_by=performed_by,
                    )
                self.after(0, lambda: _on_success(tx))
            except Exception as exc:
                err = str(exc)
                self.after(0, lambda: _on_error(err))

        def _on_success(tx: dict):
            cart_snapshot = dict(self._cart_items)   # capture before clear
            self._btn_confirm.configure(state="normal", text="Confirm Sale")
            self._clear_cart()
            if self.refresh_inventory:
                self.refresh_inventory()
            _show_sale_complete_dialog(self, tx, cart_snapshot)

        def _on_error(err: str):
            self._btn_confirm.configure(state="normal", text="Confirm Sale")
            messagebox.showerror(
                "Sale Failed",
                f"Could not save the transaction:\n\n{err}",
                parent=self.winfo_toplevel(),
            )

        threading.Thread(target=_thread, daemon=True).start()

    # ── Transaction stubs ─────────────────────────────────────────────────

    def _on_park(self):
        if not self._cart_items:
            messagebox.showwarning(
                "Empty Cart", "Nothing to park.", parent=self.winfo_toplevel(),
            )
            return
        if not messagebox.askyesno(
            "Park Transaction",
            "Save the current cart as a parked transaction?\n"
            "You can recall it later to complete the sale.",
            parent=self.winfo_toplevel(),
        ):
            return

        cart_items   = {k: dict(v) for k, v in self._cart_items.items()}
        subtotal     = self._get_subtotal()
        total        = self._get_total()
        sale_type    = self._sale_type_var.get()
        customer     = self._customer_entry.get().strip()
        performed_by = (self.current_user or {}).get("username", "")

        def _thread():
            try:
                from src.pos.transaction_client import park_transaction
                tx = park_transaction(
                    cart_items=cart_items,
                    subtotal=subtotal,
                    cart_disc_pct=self._cart_disc_pct,
                    total=total,
                    sale_type=sale_type,
                    customer_name=customer,
                    performed_by=performed_by,
                )
                self.after(0, lambda: _done(tx))
            except Exception as exc:
                err = str(exc)
                self.after(0, lambda: _fail(err))

        def _done(tx: dict):
            tx_num = tx.get("transaction_number", "")
            self._clear_cart()
            messagebox.showinfo(
                "Transaction Parked",
                f"Transaction {tx_num} has been parked.\n"
                "Recall it later to complete the sale.",
                parent=self.winfo_toplevel(),
            )

        def _fail(err: str):
            messagebox.showerror(
                "Park Failed", f"Could not park transaction:\n{err}",
                parent=self.winfo_toplevel(),
            )

        threading.Thread(target=_thread, daemon=True).start()

    def _on_recall(self):
        from src.gui.pos.recall_dialog import RecallDialog
        RecallDialog(self.winfo_toplevel(), on_select=self._load_parked_cart)

    def _on_sales(self):
        from src.gui.pos.daily_sales_dialog import DailySalesDialog
        DailySalesDialog(self.winfo_toplevel())

    def _load_parked_cart(self, tx: dict):
        """Restore the till from a recalled parked transaction snapshot."""
        snapshot   = tx.get("cart_snapshot") or {}
        cart_items = snapshot.get("cart_items") or {}
        disc_pct   = float(snapshot.get("cart_disc_pct") or 0.0)
        customer   = snapshot.get("customer_name") or ""
        sale_type  = snapshot.get("sale_type") or "Standard"

        if self._cart_items:
            if not messagebox.askyesno(
                "Replace Cart?",
                "The current cart will be replaced by the recalled transaction. Continue?",
                parent=self.winfo_toplevel(),
            ):
                return

        self._clear_cart()  # resets _parked_tx_id to None

        # Delete the parked row from the DB so it no longer appears in the list
        recalled_id = tx["id"]
        threading.Thread(
            target=lambda: _bg_delete_parked(recalled_id), daemon=True,
        ).start()

        self._cart_items = {iid: dict(line) for iid, line in cart_items.items()}
        for item_id in self._cart_items:
            self._insert_tree_row(item_id)
            self._refresh_tree_row(item_id)

        self._cart_disc_pct = disc_pct
        if disc_pct:
            self._disc_entry.delete(0, "end")
            self._disc_entry.insert(0, str(disc_pct))

        if customer:
            self._customer_entry.delete(0, "end")
            self._customer_entry.insert(0, customer)

        if sale_type in ("Standard", "Quote", "Invoice", "Repair", "Deposit", "Refund"):
            self._sale_type_var.set(sale_type)

        self._update_totals()


# ── Module-level helpers ──────────────────────────────────────────────────────

def _bg_delete_parked(tx_id: str) -> None:
    """Fire-and-forget: delete a parked transaction row from the DB."""
    try:
        from src.pos.transaction_client import delete_parked_transaction
        delete_parked_transaction(tx_id)
    except Exception:
        pass  # non-critical; the row will just linger until manually cleaned


def _show_sale_complete_dialog(parent: ctk.CTkFrame, tx: dict, cart_snapshot: dict) -> None:
    """CTkToplevel shown after a successful sale — offers to print a receipt."""
    tx_num = tx.get("transaction_number", "")
    dlg = ctk.CTkToplevel(parent.winfo_toplevel())
    dlg.title("Sale Complete")
    dlg.geometry("320x170")
    dlg.resizable(False, False)
    dlg.grab_set()
    dlg.after(50, dlg.lift)

    ctk.CTkLabel(
        dlg,
        text=f"Transaction {tx_num} confirmed.",
        font=ctk.CTkFont(size=13, weight="bold"),
    ).pack(pady=(28, 6))

    ctk.CTkLabel(
        dlg,
        text="Print receipt?",
        font=ctk.CTkFont(size=12),
        text_color=("gray40", "gray70"),
    ).pack(pady=(0, 20))

    btn_row = ctk.CTkFrame(dlg, fg_color="transparent")
    btn_row.pack()

    def _yes():
        from src.config import config
        printer = config.device.receipt_printer
        if not printer:
            dlg.destroy()
            messagebox.showwarning(
                "No Printer Configured",
                "No receipt printer is configured.\nGo to Settings → Printers to set one up.",
                parent=parent.winfo_toplevel(),
            )
            return
        dlg.destroy()
        try:
            from src.pos.receipt_generator import generate_receipt
            from src.printer_utils import print_pdf
            pdf_path = generate_receipt(tx, cart_snapshot)
            print_pdf(pdf_path, printer)
        except Exception as exc:
            messagebox.showerror(
                "Print Failed",
                f"Could not print receipt:\n{exc}",
                parent=parent.winfo_toplevel(),
            )

    def _no():
        dlg.destroy()

    ctk.CTkButton(btn_row, text="Yes", width=100, command=_yes).pack(side="left", padx=(0, 8))
    ctk.CTkButton(
        btn_row,
        text="No",
        width=100,
        fg_color="transparent",
        border_width=1,
        border_color=("gray60", "gray45"),
        text_color=("gray20", "gray90"),
        hover_color=("gray85", "gray25"),
        command=_no,
    ).pack(side="left")
