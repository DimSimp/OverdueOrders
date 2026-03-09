from __future__ import annotations

import subprocess
import threading
import re
from tkinter import messagebox, filedialog, Menu, StringVar, BooleanVar
import tkinter.ttk as ttk

import customtkinter as ctk

from src.data_processor import MatchedOrder, match_orders_to_invoice
from src.exporter import export_to_xlsx
from src.gui.order_detail_view import OrderDetailView
from src.pdf_parser import InvoiceItem


def _numeric_sort_key(value: str) -> tuple:
    """Sort key that puts numeric values first, then alphabetic."""
    match = re.match(r"^(\d+)", value.strip())
    if match:
        return (0, int(match.group(1)), value.lower())
    return (1, 0, value.lower())


def _shipping_display(shipping_type: str) -> str:
    """Return shipping text with a distinct Unicode shape prefix."""
    st = shipping_type.lower()
    if "express" in st:
        return "▲ Express"
    if "pickup" in st:
        return "■ Local Pickup"
    if st:
        return "● Regular"
    return ""


# ── OrderTreeview ──────────────────────────────────────────────────────────────

class OrderTreeview(ctk.CTkFrame):
    """
    Treeview-based order list.

    Parent rows = orders (bold).  Child rows = line items (indented).
    All parents start expanded.  Right-click shows a context menu.
    """

    def __init__(
        self,
        master,
        col_spec: dict,
        on_row_click=None,
        on_context_action=None,
        context_label: str = "Move to Unmatched",
        **kwargs,
    ):
        super().__init__(master, fg_color="transparent", **kwargs)
        self._on_row_click = on_row_click
        self._on_context_action = on_context_action
        self._context_label = context_label
        self._col_spec = col_spec
        self._group_meta: dict[str, dict] = {}  # iid → {order_id, platform}
        self._all_groups: list[dict] = []
        self._all_flat_rows: list[list[str]] = []
        self._search_var = StringVar()
        self._hovered_group: str | None = None  # parent iid of currently hovered group
        # Sorting state
        self._sort_col: str | None = None
        self._sort_reverse: bool = False
        self._shipping_cycle: int = 0  # cycles Express→Regular→Pickup priority
        self._col_headings: dict[str, str] = {}  # col_id → original heading text
        # Filter state
        self._filter_visible: bool = False
        self._filter_frame: ctk.CTkFrame | None = None
        self._platform_filters: dict[str, BooleanVar] = {}
        self._shipping_filters: dict[str, BooleanVar] = {}
        self._apply_style()
        self._build_tree()

    # ── Styling ───────────────────────────────────────────────────────────

    def _apply_style(self):
        dark = ctk.get_appearance_mode() == "Dark"
        style = ttk.Style()
        # Switch to "clam" theme so fieldbackground is respected on Windows.
        # The default "vista" theme ignores fieldbackground, leaving the
        # treeview interior white regardless of what we configure.
        style.theme_use("clam")
        style.configure(
            "Orders.Treeview",
            background="#2b2b2b" if dark else "#f5f5f5",
            foreground="#ffffff" if dark else "#1a1a1a",
            fieldbackground="#2b2b2b" if dark else "#f5f5f5",
            rowheight=28,
            font=("", 12),
            borderwidth=0,
        )
        style.configure(
            "Orders.Treeview.Heading",
            background="#1f1f1f" if dark else "#d8d8d8",
            foreground="#cccccc" if dark else "#333333",
            font=("", 12, "bold"),
            relief="flat",
        )
        style.map(
            "Orders.Treeview",
            background=[("selected", "#3a6ea5")],
            foreground=[("selected", "#ffffff")],
        )

    def _configure_tags(self):
        dark = ctk.get_appearance_mode() == "Dark"
        self._bg_a = "#303030" if dark else "#f8f8f8"
        self._bg_b = "#252525" if dark else "#ebebeb"
        self._bg_hover = "#3d5a80" if dark else "#cde0f5"
        self._tree.tag_configure("order_hdr", font=("", 12, "bold"))
        self._tree.tag_configure("matched_sku", foreground="#4fc3f7")

    # ── Widget build ──────────────────────────────────────────────────────

    def _build_tree(self):
        cols = [c for c in self._col_spec if c != "#0"]
        h0, w0 = self._col_spec["#0"]
        # Use "headings" (no tree column) when the #0 width is 0 (flat list)
        show = "tree headings" if w0 > 0 else "headings"
        self._tree = ttk.Treeview(
            self,
            style="Orders.Treeview",
            columns=cols,
            show=show,
            selectmode="browse",
        )

        # Tree (#0) column — only when show="tree headings"
        if w0 > 0:
            self._col_headings["#0"] = h0
            self._tree.heading("#0", text=h0, anchor="w",
                               command=lambda: self._on_header_click("#0"))
            self._tree.column("#0", width=w0, minwidth=60, stretch=False)

        # Data columns
        for col_id, (heading, width) in self._col_spec.items():
            if col_id == "#0":
                continue
            self._col_headings[col_id] = heading
            stretch = col_id in ("notes", "description")
            self._tree.heading(col_id, text=heading, anchor="w",
                               command=lambda c=col_id: self._on_header_click(c))
            self._tree.column(col_id, width=width, minwidth=30, stretch=stretch)

        # ── Search bar (row 0) ────────────────────────────────────────────
        search_bar = ctk.CTkFrame(self, fg_color="transparent")
        search_bar.grid(row=0, column=0, columnspan=2, sticky="ew", padx=2, pady=(2, 4))

        ctk.CTkLabel(
            search_bar, text="Search:", font=ctk.CTkFont(size=12),
        ).pack(side="left", padx=(4, 6))

        ctk.CTkEntry(
            search_bar, textvariable=self._search_var,
            placeholder_text="Filter by order, customer, SKU…",
            font=ctk.CTkFont(size=12), height=28,
        ).pack(side="left", fill="x", expand=True)

        ctk.CTkButton(
            search_bar, text="✕", width=28, height=28,
            fg_color="gray50", hover_color="gray40",
            command=lambda: self._search_var.set(""),
        ).pack(side="left", padx=(4, 4))

        self._filter_btn = ctk.CTkButton(
            search_bar, text="Filter", width=60, height=28,
            fg_color="gray50", hover_color="gray40",
            command=self._toggle_filter_panel,
        )
        self._filter_btn.pack(side="left", padx=(4, 4))

        self._search_var.trace_add("write", self._apply_filter)

        # ── Filter panel (row 1, hidden by default) ──────────────────────
        # Built on demand in _toggle_filter_panel / _rebuild_filter_panel

        # ── Treeview + scrollbar (row 2) ─────────────────────────────────
        vsb = ttk.Scrollbar(self, orient="vertical", command=self._tree.yview)
        self._tree.configure(yscrollcommand=vsb.set)
        self._tree.grid(row=2, column=0, sticky="nsew")
        vsb.grid(row=2, column=1, sticky="ns")
        self.grid_rowconfigure(0, weight=0)
        self.grid_rowconfigure(1, weight=0)
        self.grid_rowconfigure(2, weight=1)
        self.grid_columnconfigure(0, weight=1)

        self._configure_tags()

        self._tree.bind("<ButtonRelease-1>", self._on_click)
        self._tree.bind("<Button-3>", self._on_right_click)
        self._tree.bind("<Motion>", self._on_hover)
        self._tree.bind("<Leave>", self._on_leave)
        self._tree.configure(cursor="hand2")

    # ── Data loading ──────────────────────────────────────────────────────

    def load_groups(self, groups: list[dict]):
        """
        Store groups and (re-)render, respecting any active search filter.

        groups: list of {
            order_id: str, platform: str, customer: str, date: str,
            shipping: str, notes: str,
            line_items: list of {sku, description, qty, is_matched}
        }
        """
        self._all_groups = groups
        self._all_flat_rows = []
        self._rebuild_filter_options()
        self._apply_filter()

    def load_flat(self, rows: list[list[str]]):
        """Store flat rows and (re-)render, respecting any active search filter."""
        self._all_flat_rows = rows
        self._all_groups = []
        self._apply_filter()

    # ── Filtering ─────────────────────────────────────────────────────────

    def _apply_filter(self, *_):
        """Re-render the tree showing only rows that match the search query and filters."""
        query = self._search_var.get().lower().strip()
        self._tree.delete(*self._tree.get_children())
        self._group_meta.clear()

        if self._all_groups:
            visible = [
                g for g in self._all_groups
                if self._group_passes_filters(g, query)
            ]
            for i, g in enumerate(visible):
                bg = self._bg_a if i % 2 == 0 else self._bg_b
                tag_bg = f"bg_{i}"
                self._tree.tag_configure(tag_bg, background=bg)
                shipping = g.get("shipping", "")
                ship_display = _shipping_display(shipping)
                row_tags = [tag_bg, "order_hdr"]
                piid = self._tree.insert(
                    "", "end",
                    text=g["order_id"],
                    values=(
                        g["platform"], g["customer"], g["date"],
                        ship_display, "", "", "", g.get("notes", ""),
                    ),
                    tags=row_tags,
                    open=True,
                )
                self._group_meta[piid] = {"order_id": g["order_id"], "platform": g["platform"], "bg_tag": tag_bg, "bg": bg}
                for item in g["line_items"]:
                    tags = [tag_bg] + (["matched_sku"] if item["is_matched"] else [])
                    ciid = self._tree.insert(
                        piid, "end", text="",
                        values=("", "", "", "", item["sku"], item["description"], item["qty"], ""),
                        tags=tags,
                    )
                    self._group_meta[ciid] = {"order_id": g["order_id"], "platform": g["platform"], "bg_tag": tag_bg, "bg": bg}

        elif self._all_flat_rows:
            visible = [
                r for r in self._all_flat_rows
                if not query or any(query in str(v).lower() for v in r)
            ]
            for i, row in enumerate(visible):
                bg = self._bg_a if i % 2 == 0 else self._bg_b
                tag_bg = f"bg_{i}"
                self._tree.tag_configure(tag_bg, background=bg)
                self._tree.insert("", "end", text="", values=row, tags=[tag_bg])

    def _group_passes_filters(self, g: dict, query: str) -> bool:
        """Return True if group passes search query + checkbox filters."""
        # Platform filter
        if self._platform_filters:
            platform = g.get("platform", "")
            var = self._platform_filters.get(platform)
            if var is not None and not var.get():
                return False

        # Shipping filter
        if self._shipping_filters:
            shipping = g.get("shipping", "") or "Regular"
            var = self._shipping_filters.get(shipping)
            if var is not None and not var.get():
                return False

        # Search query
        if query:
            return self._group_matches_query(g, query)
        return True

    def _group_matches_query(self, g: dict, query: str) -> bool:
        """Return True if query appears in any field of the order or its line items."""
        if any(
            query in str(v).lower()
            for v in (g.get("order_id", ""), g.get("platform", ""),
                      g.get("customer", ""), g.get("date", ""),
                      g.get("shipping", ""), g.get("notes", ""))
        ):
            return True
        for item in g.get("line_items", []):
            if any(
                query in str(v).lower()
                for v in (item.get("sku", ""), item.get("description", ""),
                          str(item.get("qty", "")))
            ):
                return True
        return False

    # ── Sorting ──────────────────────────────────────────────────────────

    def _on_header_click(self, col_id: str):
        """Sort groups by the clicked column header."""
        if not self._all_groups:
            return

        if col_id == "shipping":
            # Shipping cycles through priority order instead of asc/desc
            if self._sort_col == "shipping":
                self._shipping_cycle = (self._shipping_cycle + 1) % 3
            else:
                self._sort_col = "shipping"
                self._shipping_cycle = 0
            self._sort_reverse = False
        elif col_id == self._sort_col:
            self._sort_reverse = not self._sort_reverse
        else:
            self._sort_col = col_id
            self._sort_reverse = False

        self._sort_groups()
        self._update_header_indicators()
        self._apply_filter()

    def _sort_groups(self):
        """Sort self._all_groups in place based on current sort state."""
        col = self._sort_col
        if not col:
            return

        def _sort_key(g: dict):
            if col == "#0":
                return _numeric_sort_key(g.get("order_id", ""))
            elif col == "platform":
                return g.get("platform", "").lower()
            elif col == "customer":
                return g.get("customer", "").lower()
            elif col == "date":
                return g.get("date", "")
            elif col == "shipping":
                shipping = (g.get("shipping", "") or "Regular").lower()
                # Cycle priority: 0=Express first, 1=Regular first, 2=Local Pickup first
                priority_maps = [
                    {"express": 0, "regular": 1, "local pickup": 2},
                    {"regular": 0, "express": 1, "local pickup": 2},
                    {"local pickup": 0, "express": 1, "regular": 2},
                ]
                pmap = priority_maps[self._shipping_cycle]
                return pmap.get(shipping, 3)
            elif col == "sku":
                items = g.get("line_items", [])
                return _numeric_sort_key(items[0]["sku"]) if items else ("", "")
            elif col == "qty":
                items = g.get("line_items", [])
                try:
                    return sum(int(i.get("qty", 0)) for i in items)
                except (ValueError, TypeError):
                    return 0
            elif col == "description":
                items = g.get("line_items", [])
                return items[0].get("description", "").lower() if items else ""
            elif col == "notes":
                return g.get("notes", "").lower()
            return ""

        self._all_groups.sort(key=_sort_key, reverse=self._sort_reverse)

    def _update_header_indicators(self):
        """Update column heading text with ▲/▼ sort indicator."""
        for col_id, original in self._col_headings.items():
            if col_id == self._sort_col:
                if col_id == "shipping":
                    labels = ["Express", "Regular", "Pickup"]
                    indicator = f" ({labels[self._shipping_cycle]})"
                else:
                    indicator = " ▼" if self._sort_reverse else " ▲"
                self._tree.heading(col_id, text=original + indicator)
            else:
                self._tree.heading(col_id, text=original)

    # ── Filter panel ─────────────────────────────────────────────────────

    def _toggle_filter_panel(self):
        """Show or hide the filter panel."""
        if self._filter_visible:
            if self._filter_frame:
                self._filter_frame.grid_remove()
            self._filter_visible = False
            self._filter_btn.configure(fg_color="gray50")
        else:
            self._rebuild_filter_panel()
            self._filter_visible = True
            self._filter_btn.configure(fg_color=("dodgerblue3", "dodgerblue4"))

    def _rebuild_filter_options(self):
        """Scan groups for unique platforms and update filter variables."""
        platforms = sorted({g.get("platform", "") for g in self._all_groups if g.get("platform")})
        # Preserve existing check states
        old_states = {k: v.get() for k, v in self._platform_filters.items()}
        self._platform_filters.clear()
        for p in platforms:
            var = BooleanVar(value=old_states.get(p, True))
            self._platform_filters[p] = var

        # Shipping types are fixed
        if not self._shipping_filters:
            for st in ("Express", "Regular", "Local Pickup"):
                self._shipping_filters[st] = BooleanVar(value=True)

        # Rebuild panel if visible
        if self._filter_visible and self._filter_frame:
            self._rebuild_filter_panel()

    def _rebuild_filter_panel(self):
        """Build or rebuild the filter panel UI."""
        if self._filter_frame:
            self._filter_frame.destroy()

        self._filter_frame = ctk.CTkFrame(self, fg_color=("gray88", "gray20"), corner_radius=6)
        self._filter_frame.grid(row=1, column=0, columnspan=2, sticky="ew", padx=2, pady=(0, 4))

        # Platform section
        plat_frame = ctk.CTkFrame(self._filter_frame, fg_color="transparent")
        plat_frame.pack(side="left", padx=(10, 20), pady=6)
        ctk.CTkLabel(plat_frame, text="Platform:", font=ctk.CTkFont(size=11, weight="bold")).pack(side="left", padx=(0, 6))
        for platform, var in self._platform_filters.items():
            ctk.CTkCheckBox(
                plat_frame, text=platform, variable=var,
                font=ctk.CTkFont(size=11), height=24, checkbox_width=18, checkbox_height=18,
                command=self._apply_filter,
            ).pack(side="left", padx=(0, 8))

        # Separator
        ctk.CTkLabel(self._filter_frame, text="|", text_color="gray50").pack(side="left", padx=(0, 10))

        # Shipping section
        ship_frame = ctk.CTkFrame(self._filter_frame, fg_color="transparent")
        ship_frame.pack(side="left", padx=(0, 10), pady=6)
        ctk.CTkLabel(ship_frame, text="Shipping:", font=ctk.CTkFont(size=11, weight="bold")).pack(side="left", padx=(0, 6))
        for ship_type, var in self._shipping_filters.items():
            ctk.CTkCheckBox(
                ship_frame, text=ship_type, variable=var,
                font=ctk.CTkFont(size=11), height=24, checkbox_width=18, checkbox_height=18,
                command=self._apply_filter,
            ).pack(side="left", padx=(0, 8))

    # ── Events ────────────────────────────────────────────────────────────

    def _on_click(self, event):
        if not self._on_row_click:
            return
        # Ignore clicks on the expand/collapse indicator
        element = self._tree.identify_element(event.x, event.y)
        if element == "Treeitem.indicator":
            return
        iid = self._tree.identify_row(event.y)
        if iid and iid in self._group_meta:
            meta = self._group_meta[iid]
            self._on_row_click(meta["order_id"], meta["platform"])

    def _on_right_click(self, event):
        if not self._on_context_action:
            return
        iid = self._tree.identify_row(event.y)
        if not iid or iid not in self._group_meta:
            return
        self._tree.selection_set(iid)
        meta = self._group_meta[iid]
        menu = Menu(self._tree, tearoff=0)
        menu.add_command(
            label=self._context_label,
            command=lambda: self._on_context_action(meta["order_id"], meta["platform"]),
        )
        menu.tk_popup(event.x_root, event.y_root)

    def _on_hover(self, event):
        iid = self._tree.identify_row(event.y)
        if not iid or iid not in self._group_meta:
            self._clear_hover()
            return
        parent = self._tree.parent(iid)
        group_root = parent if parent else iid
        if group_root == self._hovered_group:
            return
        self._clear_hover()
        self._hovered_group = group_root
        meta = self._group_meta[group_root]
        self._tree.tag_configure(meta["bg_tag"], background=self._bg_hover)

    def _on_leave(self, _event):
        self._clear_hover()

    def _clear_hover(self):
        if self._hovered_group is None:
            return
        meta = self._group_meta.get(self._hovered_group)
        self._hovered_group = None
        if meta:
            self._tree.tag_configure(meta["bg_tag"], background=meta["bg"])

    def scroll_to(self, order_id: str):
        """Scroll to and select the parent row for the given order_id."""
        for iid, meta in self._group_meta.items():
            if meta["order_id"] == order_id and self._tree.parent(iid) == "":
                self._tree.see(iid)
                self._tree.selection_set(iid)
                break


# ── Column specs ───────────────────────────────────────────────────────────────

# Matched orders: tree col = Order No., data cols = Platform … Notes
_MATCHED_COL_SPEC = {
    "#0":          ("Order No.",   120),
    "platform":    ("Platform",     80),
    "customer":    ("Customer",    140),
    "date":        ("Date",         90),
    "shipping":    ("Shipping",     90),
    "sku":         ("SKU",         140),
    "description": ("Description", 200),
    "qty":         ("Qty",          40),
    "notes":       ("Notes",       120),
}

# Unmatched orders: same columns but no detail-view on click
_UNMATCHED_ORD_COL_SPEC = {
    "#0":          ("Order No.",   120),
    "platform":    ("Platform",     80),
    "customer":    ("Customer",    140),
    "date":        ("Date",         90),
    "shipping":    ("Shipping",     90),
    "sku":         ("SKU",         140),
    "description": ("Description", 200),
    "qty":         ("Qty",          40),
    "notes":       ("Notes",       120),
}

# Unmatched invoice items: flat list
_INV_COL_SPEC = {
    "#0":          ("",             0),
    "sku":         ("SKU (suffix)", 220),
    "description": ("Description",  450),
    "qty":         ("Qty",           60),
}


# ── ResultsTab ─────────────────────────────────────────────────────────────────

class ResultsTab(ctk.CTkFrame):
    def __init__(self, master, app, **kwargs):
        super().__init__(master, fg_color="transparent", **kwargs)
        self._app = app
        self._matched: list[MatchedOrder] = []
        self._unmatched_inv: list[InvoiceItem] = []
        self._neto_orders = []
        self._ebay_orders = []
        # Manual overrides: (platform, order_id) tuples
        self._excluded_order_ids: set[tuple[str, str]] = set()
        self._force_matched_order_ids: set[tuple[str, str]] = set()
        # Navigation state
        self._detail_frame: OrderDetailView | None = None
        self._freight_frame = None
        self._last_clicked_order_id: str | None = None
        self._build_ui()

    # ── UI construction ───────────────────────────────────────────────────

    def _build_ui(self):
        self.grid_rowconfigure(0, weight=1)
        self.grid_columnconfigure(0, weight=1)

        # List page — always exists; sits at (0,0)
        self._list_frame = ctk.CTkFrame(self, fg_color="transparent")
        self._list_frame.grid(row=0, column=0, sticky="nsew")
        self._build_list_page(self._list_frame)

        # Detail frame is built on demand and placed at (0,0) on top of list_frame

    def _build_list_page(self, parent):
        parent.grid_rowconfigure(1, weight=1)
        parent.grid_columnconfigure(0, weight=1)

        # ── Summary row ───────────────────────────────────────────────────
        summary = ctk.CTkFrame(parent, fg_color="transparent")
        summary.grid(row=0, column=0, sticky="ew", padx=12, pady=(12, 6))

        self._matched_lbl = ctk.CTkLabel(
            summary, text="Matched: —", font=ctk.CTkFont(size=13, weight="bold")
        )
        self._matched_lbl.pack(side="left", padx=(0, 24))

        self._unmatched_inv_lbl = ctk.CTkLabel(
            summary, text="Unmatched invoice items: —", font=ctk.CTkFont(size=13)
        )
        self._unmatched_inv_lbl.pack(side="left", padx=(0, 24))

        self._unmatched_orders_lbl = ctk.CTkLabel(
            summary, text="Unmatched orders: —", font=ctk.CTkFont(size=13)
        )
        self._unmatched_orders_lbl.pack(side="left")

        # ── Inner tab view ────────────────────────────────────────────────
        self._inner_tabs = ctk.CTkTabview(parent, corner_radius=6)
        self._inner_tabs.grid(row=1, column=0, sticky="nsew", padx=12, pady=4)

        for name in ("Matched Orders", "Unmatched Invoice Items", "Unmatched Orders"):
            self._inner_tabs.add(name)

        # Matched Orders tree
        self._matched_tree = OrderTreeview(
            self._inner_tabs.tab("Matched Orders"),
            col_spec=_MATCHED_COL_SPEC,
            on_row_click=self._open_detail_view,
            on_context_action=self._exclude_order,
            context_label="Move to Unmatched",
        )
        self._matched_tree.pack(fill="both", expand=True)

        # Unmatched Invoice Items tree (flat)
        self._inv_tree = OrderTreeview(
            self._inner_tabs.tab("Unmatched Invoice Items"),
            col_spec=_INV_COL_SPEC,
        )
        self._inv_tree.pack(fill="both", expand=True)

        # Unmatched Orders tree
        ctk.CTkLabel(
            self._inner_tabs.tab("Unmatched Orders"),
            text=(
                "Neto: 'on PO' orders (paid, undispatched) whose SKUs did not match the invoice.\n"
                "eBay: all paid, unfulfilled orders whose SKUs did not match the invoice.\n\n"
                "Right-click an order to move it to the Matched list."
            ),
            font=ctk.CTkFont(size=13),
            justify="left",
        ).pack(padx=20, pady=(16, 4), anchor="nw")

        self._unmatched_orders_tree = OrderTreeview(
            self._inner_tabs.tab("Unmatched Orders"),
            col_spec=_UNMATCHED_ORD_COL_SPEC,
            on_context_action=self._include_order,
            context_label="Move to Matched",
        )
        self._unmatched_orders_tree.pack(fill="both", expand=True)

        # ── Bottom row ────────────────────────────────────────────────────
        bottom = ctk.CTkFrame(parent, fg_color="transparent")
        bottom.grid(row=2, column=0, sticky="ew", padx=12, pady=(4, 12))

        self._export_btn = ctk.CTkButton(
            bottom, text="Export to Excel", width=140, command=self._export_csv,
        )
        self._export_btn.pack(side="left")

        self._refresh_btn = ctk.CTkButton(
            bottom, text="Refresh Orders", width=130,
            fg_color=("dodgerblue3", "dodgerblue4"),
            command=self._refresh_orders,
        )
        self._refresh_btn.pack(side="left", padx=(12, 0))

        self._save_session_btn = ctk.CTkButton(
            bottom, text="Save Session As", width=130,
            fg_color="gray50", hover_color="gray40",
            command=self._save_session_as,
        )
        self._save_session_btn.pack(side="left", padx=(12, 0))

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
        try:
            self._error_label.configure(text="Loading…")
            self.update_idletasks()

            invoice_items = self._app.invoice_tab.get_invoice_items()
            self._neto_orders = self._app.neto_orders
            self._ebay_orders = self._app.ebay_orders

            matched, unmatched_inv = match_orders_to_invoice(
                invoice_items,
                self._neto_orders,
                self._ebay_orders,
                on_po_phrase=self._app.config.app.on_po_filter_phrase,
            )

            self._excluded_order_ids.clear()
            self._force_matched_order_ids.clear()
            self._matched = matched
            self._unmatched_inv = unmatched_inv
            self._app.matched_orders = matched

            self._refresh_tables()
            self._error_label.configure(text="")

        except Exception as exc:
            import traceback, sys
            traceback.print_exc(file=sys.stderr)
            self._error_label.configure(text=f"Error loading results: {exc}")

    # ── Table population ──────────────────────────────────────────────────

    def _refresh_tables(self):
        effective_matched = [
            m for m in self._matched
            if (m.platform, m.order_id) not in self._excluded_order_ids
        ]
        force_matched = self._build_force_matched()
        effective_matched.extend(force_matched)

        self._app.matched_orders = effective_matched
        self._populate_matched(effective_matched)
        self._populate_unmatched_inv(self._unmatched_inv)
        self._populate_unmatched_orders(effective_matched)
        self._update_summary(effective_matched, self._unmatched_inv)
        self._auto_save_session()

    def _build_force_matched(self) -> list[MatchedOrder]:
        if not self._force_matched_order_ids:
            return []

        result = []
        for order in self._neto_orders:
            channel = order.sales_channel or "Neto"
            key = (channel, order.order_id)
            if key not in self._force_matched_order_ids:
                continue
            order_date = order.date_paid or order.date_placed
            for idx, line in enumerate(order.line_items):
                result.append(MatchedOrder(
                    platform=channel,
                    order_id=order.order_id,
                    customer_name=order.customer_name,
                    order_date=order_date,
                    sku=line.sku,
                    description=line.product_name,
                    quantity=line.quantity,
                    notes=order.notes if idx == 0 else "",
                    shipping_type=order.shipping_type,
                    invoice_sku="",
                    invoice_description="",
                    invoice_qty=0,
                    is_invoice_match=False,
                ))

        for order in self._ebay_orders:
            key = ("eBay", order.order_id)
            if key not in self._force_matched_order_ids:
                continue
            for line in order.line_items:
                result.append(MatchedOrder(
                    platform="eBay",
                    order_id=order.order_id,
                    customer_name=order.buyer_name,
                    order_date=order.creation_date,
                    sku=line.sku,
                    description=line.title,
                    quantity=line.quantity,
                    notes=line.notes,
                    shipping_type=order.shipping_type,
                    invoice_sku="",
                    invoice_description="",
                    invoice_qty=0,
                    is_invoice_match=False,
                ))

        return result

    def _populate_matched(self, matched: list[MatchedOrder]):
        def _platform_key(m: MatchedOrder) -> tuple:
            pl = m.platform.lower()
            if pl == "website":
                return (0, m.platform, m.order_id)
            if pl == "ebay":
                return (2, m.platform, m.order_id)
            return (1, m.platform, m.order_id)

        # Build ordered groups preserving sort order
        groups: list[dict] = []
        seen: dict[tuple[str, str], int] = {}

        for m in sorted(matched, key=_platform_key):
            key = (m.platform, m.order_id)
            if key not in seen:
                date_str = m.order_date.strftime("%Y-%m-%d") if m.order_date else ""
                seen[key] = len(groups)
                groups.append({
                    "order_id": m.order_id,
                    "platform": m.platform,
                    "customer": m.customer_name,
                    "date": date_str,
                    "shipping": m.shipping_type,
                    "notes": m.notes,
                    "line_items": [],
                })
            groups[seen[key]]["line_items"].append({
                "sku": m.sku,
                "description": m.description,
                "qty": str(m.quantity),
                "is_matched": m.is_invoice_match,
            })

        self._matched_tree.load_groups(groups)

    def _populate_unmatched_inv(self, items: list[InvoiceItem]):
        rows = [[item.sku_with_suffix, item.description, str(item.quantity)] for item in items]
        self._inv_tree.load_flat(rows)

    def _populate_unmatched_orders(self, matched):
        matched_ids = {(m.platform, m.order_id) for m in matched}

        def _platform_key(g: dict) -> tuple:
            pl = g["platform"].lower()
            if pl == "website":
                return (0, g["platform"], g["order_id"])
            if pl == "ebay":
                return (2, g["platform"], g["order_id"])
            return (1, g["platform"], g["order_id"])

        groups: list[dict] = []
        seen: dict[tuple[str, str], int] = {}

        for order in self._neto_orders:
            channel = order.sales_channel or "Neto"
            if (channel, order.order_id) in matched_ids:
                continue
            date_str = order.date_paid.strftime("%Y-%m-%d") if order.date_paid else ""
            key = (channel, order.order_id)
            if key not in seen:
                seen[key] = len(groups)
                groups.append({
                    "order_id": order.order_id,
                    "platform": channel,
                    "customer": order.customer_name,
                    "date": date_str,
                    "shipping": order.shipping_type,
                    "notes": order.notes,
                    "line_items": [],
                })
            for line in order.line_items:
                groups[seen[key]]["line_items"].append({
                    "sku": line.sku,
                    "description": line.product_name,
                    "qty": str(line.quantity),
                    "is_matched": False,
                })

        for order in self._ebay_orders:
            if ("eBay", order.order_id) in matched_ids:
                continue
            date_str = order.creation_date.strftime("%Y-%m-%d") if order.creation_date else ""
            key = ("eBay", order.order_id)
            if key not in seen:
                seen[key] = len(groups)
                groups.append({
                    "order_id": order.order_id,
                    "platform": "eBay",
                    "customer": order.buyer_name,
                    "date": date_str,
                    "shipping": order.shipping_type,
                    "notes": "",
                    "line_items": [],
                })
            for line in order.line_items:
                groups[seen[key]]["line_items"].append({
                    "sku": line.sku,
                    "description": line.title,
                    "qty": str(line.quantity),
                    "is_matched": False,
                })

        groups.sort(key=_platform_key)
        self._unmatched_orders_tree.load_groups(groups)

    def _update_summary(self, matched, unmatched_inv):
        candidate_count = len(self._neto_orders) + len(self._ebay_orders)
        matched_order_ids = {(m.platform, m.order_id) for m in matched}
        unmatched_order_count = max(0, candidate_count - len(matched_order_ids))
        match_count = sum(1 for m in matched if m.is_invoice_match)

        self._matched_lbl.configure(
            text=f"Matched: {match_count} invoice line{'s' if match_count != 1 else ''}",
            text_color=("green" if matched else "gray50"),
        )
        self._unmatched_inv_lbl.configure(text=f"Unmatched invoice items: {len(unmatched_inv)}")
        self._unmatched_orders_lbl.configure(text=f"Unmatched orders: {unmatched_order_count}")

    # ── Override management ───────────────────────────────────────────────

    def _exclude_order(self, order_id: str, platform: str = ""):
        """Move an order from matched → unmatched."""
        # Find platform if not provided (from matched list)
        if not platform:
            for m in self._matched:
                if m.order_id == order_id:
                    platform = m.platform
                    break

        key = (platform, order_id)
        self._excluded_order_ids.add(key)
        self._force_matched_order_ids.discard(key)
        self._refresh_tables()

    def _include_order(self, order_id: str, platform: str = ""):
        """Move an order from unmatched → matched."""
        # If originally matched but excluded, just un-exclude
        for m in self._matched:
            if m.order_id == order_id:
                key = (m.platform, m.order_id)
                if key in self._excluded_order_ids:
                    self._excluded_order_ids.discard(key)
                    self._refresh_tables()
                    return

        # Otherwise, force-add from raw order lists
        for order in self._neto_orders:
            channel = order.sales_channel or "Neto"
            if order.order_id == order_id:
                self._force_matched_order_ids.add((channel, order.order_id))
                self._refresh_tables()
                return

        for order in self._ebay_orders:
            if order.order_id == order_id:
                self._force_matched_order_ids.add(("eBay", order.order_id))
                self._refresh_tables()
                return

    # ── Order Detail (browser-style navigation) ───────────────────────────

    def _find_order_data(self, order_id: str, platform: str):
        """Return (neto_order, ebay_order, matched_skus) for the given order."""
        neto_order = None
        ebay_order = None

        if platform.lower() == "ebay":
            for o in self._ebay_orders:
                if o.order_id == order_id:
                    ebay_order = o
                    break
        else:
            for o in self._neto_orders:
                if o.order_id == order_id:
                    neto_order = o
                    break

        matched_skus = [m.sku for m in self._matched if m.order_id == order_id and m.is_invoice_match]
        return neto_order, ebay_order, matched_skus

    def _open_detail_view(self, order_id: str, platform: str):
        """Navigate from the list to an order's detail page."""
        self._last_clicked_order_id = order_id

        neto_order, ebay_order, matched_skus = self._find_order_data(order_id, platform)
        if neto_order is None and ebay_order is None:
            self._error_label.configure(text=f"Order {order_id} not found in loaded data")
            return

        # Destroy previous detail frame if any
        if self._detail_frame is not None:
            self._detail_frame.destroy()

        # Only offer "Book Freight" if shipping is configured
        book_freight_cb = None
        if self._app.config.shipping is not None:
            book_freight_cb = self._open_freight_view

        self._detail_frame = OrderDetailView(
            self,
            order_id=order_id,
            platform=platform,
            neto_order=neto_order,
            ebay_order=ebay_order,
            matched_skus=matched_skus,
            neto_client=self._app.neto_client,
            ebay_client=self._app.ebay_client,
            dry_run=self._app.config.app.dry_run,
            on_back=self._close_detail_view,
            on_fulfilled=self._on_fulfilled,
            on_move_to_unmatched=lambda: self._exclude_order(order_id, platform),
            on_book_freight=book_freight_cb,
        )
        self._detail_frame.grid(row=0, column=0, sticky="nsew")
        self._detail_frame.tkraise()

        # Background status check — warn in detail view if already completed
        self._check_status_background(order_id, platform)

    def _close_detail_view(self):
        if self._detail_frame is not None:
            self._detail_frame.destroy()
            self._detail_frame = None
        self._list_frame.tkraise()
        if self._last_clicked_order_id:
            self._matched_tree.scroll_to(self._last_clicked_order_id)

    def _on_fulfilled(self):
        self._close_detail_view()
        self._refresh_orders()

    # ── Freight Booking View ─────────────────────────────────────────────

    def _open_freight_view(self, order_id: str, platform: str):
        """Open the freight booking view, stacking it on top of the detail view."""
        from src.gui.freight_booking_view import FreightBookingView

        neto_order, ebay_order, _ = self._find_order_data(order_id, platform)

        # Destroy previous freight frame if any
        if hasattr(self, "_freight_frame") and self._freight_frame is not None:
            self._freight_frame.destroy()

        self._freight_frame = FreightBookingView(
            self,
            order_id=order_id,
            platform=platform,
            neto_order=neto_order,
            ebay_order=ebay_order,
            neto_client=self._app.neto_client,
            shipping_config=self._app.config.shipping,
            dry_run=self._app.config.app.dry_run,
            on_back=self._close_freight_view,
            on_courier_selected=lambda name: self._on_courier_selected(name),
        )
        self._freight_frame.grid(row=0, column=0, sticky="nsew")
        self._freight_frame.tkraise()

    def _close_freight_view(self):
        """Close freight view, return to order detail view."""
        if hasattr(self, "_freight_frame") and self._freight_frame is not None:
            self._freight_frame.destroy()
            self._freight_frame = None
        if self._detail_frame is not None:
            self._detail_frame.tkraise()

    def _on_courier_selected(self, courier_name: str):
        """Called when user picks a courier from the freight view."""
        self._close_freight_view()
        if self._detail_frame is not None:
            self._detail_frame.set_tracking(carrier=courier_name)

    def _check_status_background(self, order_id: str, platform: str):
        """Check if an order is already completed; warn in the detail view if so."""
        if platform.lower() == "ebay" and not self._app.ebay_client.is_authenticated():
            return

        def _check():
            try:
                if platform.lower() == "ebay":
                    status = self._app.ebay_client.get_order_status(order_id)
                    is_completed = status in ("FULFILLED",)
                else:
                    status = self._app.neto_client.get_order_status(order_id)
                    is_completed = status.lower() in ("dispatched", "shipped", "completed")

                if is_completed:
                    self.after(0, lambda: self._handle_already_completed(order_id, platform))
            except Exception:
                pass  # status check is informational; don't block the user

        threading.Thread(target=_check, daemon=True).start()

    def _handle_already_completed(self, order_id: str, platform: str):
        # Only warn if the detail view for this order is still open
        if (
            self._detail_frame is not None
            and self._last_clicked_order_id == order_id
        ):
            self._detail_frame.show_completed_warning()

    # ── Orders refresh ────────────────────────────────────────────────────

    def _refresh_orders(self):
        self._error_label.configure(text="Refreshing orders...")
        self._refresh_btn.configure(state="disabled")
        self.update_idletasks()

        def _fetch():
            try:
                from datetime import datetime, timedelta
                lookback = self._app.config.app.order_lookback_days
                date_to = datetime.now()
                date_from = date_to - timedelta(days=lookback)

                neto_orders = self._app.neto_client.get_overdue_orders(date_from, date_to)
                ebay_orders = []
                if self._app.ebay_client.is_authenticated():
                    ebay_orders = self._app.ebay_client.get_overdue_orders(date_from, date_to)

                self.after(0, lambda: self._apply_refreshed_orders(neto_orders, ebay_orders))
            except Exception as e:
                err_msg = f"Refresh failed: {e}"
                self.after(0, lambda m=err_msg: self._error_label.configure(text=m))
                self.after(0, lambda: self._refresh_btn.configure(state="normal"))

        threading.Thread(target=_fetch, daemon=True).start()

    def _apply_refreshed_orders(self, neto_orders, ebay_orders):
        self._app.neto_orders = neto_orders
        self._app.ebay_orders = ebay_orders
        self._neto_orders = neto_orders
        self._ebay_orders = ebay_orders

        invoice_items = self._app.invoice_tab.get_invoice_items()
        matched, unmatched_inv = match_orders_to_invoice(
            invoice_items,
            self._neto_orders,
            self._ebay_orders,
            on_po_phrase=self._app.config.app.on_po_filter_phrase,
        )
        self._matched = matched
        self._unmatched_inv = unmatched_inv
        self._excluded_order_ids.clear()
        self._force_matched_order_ids.clear()
        self._refresh_tables()
        self._refresh_btn.configure(state="normal")
        self._error_label.configure(text="")

    # ── Save session ──────────────────────────────────────────────────────

    def _auto_save_session(self):
        """Silently save session to the default location on every refresh."""
        try:
            from src.session import save_snapshot
            save_dir = self._app.config.app.snapshot_dir or self._app.config.app.output_dir
            invoice_items = self._app.invoice_tab.get_invoice_items()
            save_snapshot(
                save_dir=save_dir,
                invoice_items=invoice_items,
                neto_orders=self._neto_orders,
                ebay_orders=self._ebay_orders,
                matched_orders=self._app.matched_orders,
                unmatched_inv=self._unmatched_inv,
                excluded_ids=self._excluded_order_ids,
                force_matched_ids=self._force_matched_order_ids,
            )
        except Exception:
            pass

    def _save_session_as(self):
        from src.session import save_snapshot
        initial_dir = self._app.config.app.snapshot_dir or self._app.config.app.output_dir
        save_dir = filedialog.askdirectory(
            title="Choose snapshot save location",
            initialdir=initial_dir,
            parent=self,
        )
        if not save_dir:
            return
        try:
            invoice_items = self._app.invoice_tab.get_invoice_items()
            path = save_snapshot(
                save_dir=save_dir,
                invoice_items=invoice_items,
                neto_orders=self._neto_orders,
                ebay_orders=self._ebay_orders,
                matched_orders=self._app.matched_orders,
                unmatched_inv=self._unmatched_inv,
                excluded_ids=self._excluded_order_ids,
                force_matched_ids=self._force_matched_order_ids,
            )
            self._export_label.configure(text=f"Session saved: {path}", text_color="green")
        except Exception as e:
            self._error_label.configure(text=f"Save failed: {e}")

    # ── Export ────────────────────────────────────────────────────────────

    def _export_csv(self):
        if not self._app.matched_orders:
            messagebox.showinfo("No Data", "There are no matched orders to export.")
            return
        try:
            path = export_to_xlsx(
                self._app.matched_orders,
                output_dir=self._app.config.app.output_dir,
            )
            self._export_label.configure(text=f"Saved: {path}", text_color="green")
            self._error_label.configure(text="")
            subprocess.Popen(["explorer", "/select,", path])
        except Exception as e:
            self._error_label.configure(text=f"Export failed: {e}")
            self._export_label.configure(text="")
