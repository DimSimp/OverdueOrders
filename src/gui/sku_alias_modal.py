from __future__ import annotations

import logging
import threading
from tkinter import messagebox

import customtkinter as ctk

log = logging.getLogger(__name__)


class SkuAliasModal(ctk.CTkToplevel):
    """
    Modal for SKU management — two tabs:

    Tab 1 "SKU Aliases"
        Create / edit / remove invoice-SKU alias mappings (CSV).

    Tab 2 "Rename SKU"
        Rename a product's SKU directly in Neto via the API.

    Two modes
    ---------
    "line_items"
        Opens with a pre-populated list of (sku, description) tuples from an order.
        Tab 1 shows a card for each line item.
        Tab 2 shows a dropdown to select which line item to rename.

    "search"
        Opens with an empty search bar.
        Tab 1: user types a Neto SKU, clicks Look up, a card appears.
        Tab 2: user types a SKU, clicks Look up, the rename form fills.
    """

    def __init__(
        self,
        master,
        *,
        sku_alias_manager,
        mode: str = "line_items",
        line_items: list[tuple[str, str]] | None = None,   # [(sku, description), ...]
        neto_client=None,
        suppliers=None,   # list[SupplierConfig] from config
        on_sku_renamed=None,   # callable(old_sku, new_sku) or None
        dry_run: bool = False,
    ):
        super().__init__(master)
        self._manager = sku_alias_manager
        self._mode = mode
        self._line_items = line_items or []
        self._neto_client = neto_client
        self._on_sku_renamed = on_sku_renamed
        self._dry_run = dry_run

        # Build supplier option list: display label → supplier name (empty = no supplier)
        self._supplier_options: list[str] = ["— No supplier —"]
        self._supplier_name_map: dict[str, str] = {"— No supplier —": ""}
        for s in (suppliers or []):
            if s.suffix:
                if s.suffix_position == "prepend":
                    label = f"{s.name}  ({s.suffix}+)"
                else:
                    label = f"{s.name}  (+{s.suffix})"
            else:
                label = s.name
            self._supplier_options.append(label)
            self._supplier_name_map[label] = s.name

        # Tracks per-card state: {sku_upper: _CardState}
        self._cards: dict[str, _CardState] = {}

        # Rename tab state
        self._rename_current_sku: str | None = None
        self._rename_current_name: str | None = None
        self._rename_new_var = ctk.StringVar()
        self._rename_update_alias_var = ctk.BooleanVar(value=False)
        self._rename_status_lbl: ctk.CTkLabel | None = None
        self._rename_apply_btn: ctk.CTkButton | None = None
        self._rename_form_frame: ctk.CTkFrame | None = None
        self._rename_current_sku_lbl: ctk.CTkLabel | None = None
        self._rename_current_name_lbl: ctk.CTkLabel | None = None

        self.title("SKU Management")
        self.geometry("700x600")
        self.minsize(580, 420)
        self.resizable(True, True)
        self.grab_set()
        self.lift()
        self.focus_force()

        self._build_ui()

        # In line_items mode, populate alias cards immediately
        if self._mode == "line_items":
            for sku, desc in self._line_items:
                self._add_card(sku, desc)

    # ── Layout ────────────────────────────────────────────────────────────

    def _build_ui(self):
        self.grid_rowconfigure(0, weight=1)
        self.grid_columnconfigure(0, weight=1)

        tabview = ctk.CTkTabview(self)
        tabview.grid(row=0, column=0, sticky="nsew", padx=8, pady=8)

        tab_aliases = tabview.add("SKU Aliases")
        tab_rename = tabview.add("Rename SKU")

        self._build_aliases_tab(tab_aliases)
        self._build_rename_tab(tab_rename)

    # ── Tab 1: SKU Aliases ────────────────────────────────────────────────

    def _build_aliases_tab(self, tab):
        tab.grid_rowconfigure(1, weight=1)
        tab.grid_columnconfigure(0, weight=1)

        # ── Search bar (search mode only) ─────────────────────────────────
        if self._mode == "search":
            search_frame = ctk.CTkFrame(tab, fg_color=("gray90", "gray20"), corner_radius=0)
            search_frame.grid(row=0, column=0, sticky="ew", padx=0, pady=(0, 4))
            search_frame.grid_columnconfigure(1, weight=1)

            ctk.CTkLabel(
                search_frame,
                text="Neto / eBay SKU:",
                font=ctk.CTkFont(size=13),
            ).grid(row=0, column=0, padx=(12, 6), pady=10)

            self._search_var = ctk.StringVar()
            self._search_entry = ctk.CTkEntry(
                search_frame,
                textvariable=self._search_var,
                placeholder_text="e.g. TORMIX",
                font=ctk.CTkFont(size=13),
            )
            self._search_entry.grid(row=0, column=1, sticky="ew", padx=(0, 6), pady=10)
            self._search_entry.bind("<Return>", lambda _e: self._do_lookup())

            self._lookup_btn = ctk.CTkButton(
                search_frame,
                text="Look up",
                width=90,
                command=self._do_lookup,
            )
            self._lookup_btn.grid(row=0, column=2, padx=(0, 12), pady=10)

            self._lookup_status = ctk.CTkLabel(
                search_frame,
                text="",
                font=ctk.CTkFont(size=11),
                text_color=("gray50", "gray60"),
            )
            self._lookup_status.grid(row=1, column=0, columnspan=3, padx=12, pady=(0, 8), sticky="w")
        else:
            ctk.CTkFrame(tab, height=0, fg_color="transparent").grid(row=0, column=0)

        # ── Scrollable card area ──────────────────────────────────────────
        self._scroll = ctk.CTkScrollableFrame(tab, fg_color="transparent")
        self._scroll.grid(row=1, column=0, sticky="nsew", padx=0, pady=(0, 0))
        self._scroll.grid_columnconfigure(0, weight=1)

        # ── Footer ────────────────────────────────────────────────────────
        footer = ctk.CTkFrame(tab, fg_color="transparent")
        footer.grid(row=2, column=0, sticky="ew", padx=0, pady=(8, 4))

        ctk.CTkButton(
            footer,
            text="Cancel",
            width=100,
            fg_color=("gray70", "gray30"),
            hover_color=("gray60", "gray25"),
            command=self.destroy,
        ).pack(side="right", padx=(6, 0))

        ctk.CTkButton(
            footer,
            text="Save",
            width=100,
            command=self._save_all,
        ).pack(side="right")

    # ── Tab 2: Rename SKU ─────────────────────────────────────────────────

    def _build_rename_tab(self, tab):
        tab.grid_rowconfigure(2, weight=1)
        tab.grid_columnconfigure(0, weight=1)

        # ── Top section (mode-dependent) ──────────────────────────────────
        top_frame = ctk.CTkFrame(tab, fg_color=("gray90", "gray20"), corner_radius=6)
        top_frame.grid(row=0, column=0, sticky="ew", padx=0, pady=(0, 10))
        top_frame.grid_columnconfigure(1, weight=1)

        if self._mode == "line_items" and self._line_items:
            ctk.CTkLabel(
                top_frame,
                text="Select item:",
                font=ctk.CTkFont(size=13),
            ).grid(row=0, column=0, padx=(12, 6), pady=12)

            options = [f"{sku}  —  {desc}" if desc else sku for sku, desc in self._line_items]
            self._rename_select_var = ctk.StringVar(value=options[0] if options else "")
            menu = ctk.CTkOptionMenu(
                top_frame,
                variable=self._rename_select_var,
                values=options,
                font=ctk.CTkFont(size=12),
                dynamic_resizing=False,
                command=self._on_rename_item_selected,
            )
            menu.grid(row=0, column=1, sticky="ew", padx=(0, 12), pady=12)
        else:
            # search mode: SKU entry + Look up
            ctk.CTkLabel(
                top_frame,
                text="Neto SKU:",
                font=ctk.CTkFont(size=13),
            ).grid(row=0, column=0, padx=(12, 6), pady=10)

            self._rename_search_var = ctk.StringVar()
            rename_entry = ctk.CTkEntry(
                top_frame,
                textvariable=self._rename_search_var,
                placeholder_text="e.g. TORMIX",
                font=ctk.CTkFont(size=13),
            )
            rename_entry.grid(row=0, column=1, sticky="ew", padx=(0, 6), pady=10)
            rename_entry.bind("<Return>", lambda _e: self._do_rename_lookup())

            self._rename_lookup_btn = ctk.CTkButton(
                top_frame,
                text="Look up",
                width=90,
                command=self._do_rename_lookup,
            )
            self._rename_lookup_btn.grid(row=0, column=2, padx=(0, 12), pady=10)

            self._rename_lookup_status = ctk.CTkLabel(
                top_frame,
                text="",
                font=ctk.CTkFont(size=11),
                text_color=("gray50", "gray60"),
            )
            self._rename_lookup_status.grid(row=1, column=0, columnspan=3, padx=12, pady=(0, 8), sticky="w")

        # ── Rename form ────────────────────────────────────────────────────
        self._rename_form_frame = ctk.CTkFrame(tab, fg_color="transparent")
        self._rename_form_frame.grid(row=1, column=0, sticky="ew", padx=4, pady=0)
        self._rename_form_frame.grid_columnconfigure(1, weight=1)

        ctk.CTkLabel(
            self._rename_form_frame,
            text="Current SKU:",
            font=ctk.CTkFont(size=12),
            width=110,
            anchor="w",
        ).grid(row=0, column=0, padx=(4, 8), pady=(8, 2), sticky="w")

        self._rename_current_sku_lbl = ctk.CTkLabel(
            self._rename_form_frame,
            text="—",
            font=ctk.CTkFont(size=12, weight="bold"),
            anchor="w",
        )
        self._rename_current_sku_lbl.grid(row=0, column=1, padx=0, pady=(8, 2), sticky="w")

        ctk.CTkLabel(
            self._rename_form_frame,
            text="Description:",
            font=ctk.CTkFont(size=12),
            width=110,
            anchor="w",
        ).grid(row=1, column=0, padx=(4, 8), pady=(2, 6), sticky="w")

        self._rename_current_name_lbl = ctk.CTkLabel(
            self._rename_form_frame,
            text="",
            font=ctk.CTkFont(size=12),
            text_color=("gray45", "gray65"),
            anchor="w",
        )
        self._rename_current_name_lbl.grid(row=1, column=1, padx=0, pady=(2, 6), sticky="w")

        ctk.CTkLabel(
            self._rename_form_frame,
            text="New SKU:",
            font=ctk.CTkFont(size=12),
            width=110,
            anchor="w",
        ).grid(row=2, column=0, padx=(4, 8), pady=(0, 6), sticky="w")

        new_sku_entry = ctk.CTkEntry(
            self._rename_form_frame,
            textvariable=self._rename_new_var,
            placeholder_text="Enter new SKU",
            font=ctk.CTkFont(size=12),
        )
        new_sku_entry.grid(row=2, column=1, padx=(0, 8), pady=(0, 6), sticky="ew")
        new_sku_entry.bind("<Return>", lambda _e: self._do_rename())

        ctk.CTkCheckBox(
            self._rename_form_frame,
            text="Update alias mapping if one exists",
            variable=self._rename_update_alias_var,
            font=ctk.CTkFont(size=12),
        ).grid(row=3, column=0, columnspan=2, padx=(4, 0), pady=(0, 10), sticky="w")

        self._rename_apply_btn = ctk.CTkButton(
            self._rename_form_frame,
            text="Rename in Neto",
            width=160,
            command=self._do_rename,
        )
        self._rename_apply_btn.grid(row=4, column=0, columnspan=2, pady=(0, 8))

        self._rename_status_lbl = ctk.CTkLabel(
            self._rename_form_frame,
            text="",
            font=ctk.CTkFont(size=12),
            text_color=("gray50", "gray60"),
            wraplength=480,
        )
        self._rename_status_lbl.grid(row=5, column=0, columnspan=2, padx=4, pady=(0, 4), sticky="w")

        # Spacer
        ctk.CTkFrame(tab, fg_color="transparent").grid(row=2, column=0, sticky="nsew")

        # Pre-select first item now that all form widgets are built
        if self._mode == "line_items" and self._line_items:
            options = [f"{sku}  —  {desc}" if desc else sku for sku, desc in self._line_items]
            if options:
                self._on_rename_item_selected(options[0])

    def _on_rename_item_selected(self, option: str):
        """Called when user picks an item from the line_items dropdown."""
        # Extract the SKU from "SKU  —  description" or just "SKU"
        sku = option.split("  —  ")[0].strip()
        name = option.split("  —  ", 1)[1].strip() if "  —  " in option else ""
        self._set_rename_subject(sku, name)

    def _do_rename_lookup(self):
        """Called in search mode: look up the SKU via Neto API."""
        sku = self._rename_search_var.get().strip()
        if not sku:
            return
        self._rename_lookup_btn.configure(state="disabled")
        self._rename_lookup_status.configure(text="Looking up…", text_color=("gray50", "gray60"))

        def _fetch():
            info = None
            if self._neto_client:
                try:
                    info = self._neto_client.get_item_info(sku)
                except Exception as exc:
                    log.warning("SkuAliasModal rename lookup failed: %s", exc)
            self.after(0, lambda: self._on_rename_lookup_done(sku, info))

        threading.Thread(target=_fetch, daemon=True).start()

    def _on_rename_lookup_done(self, sku: str, info: dict | None):
        self._rename_lookup_btn.configure(state="normal")
        if info is None:
            self._rename_lookup_status.configure(
                text=f"'{sku}' not found in Neto.",
                text_color="orange",
            )
            self._set_rename_subject(sku, "")
        else:
            self._rename_lookup_status.configure(text="", text_color=("gray50", "gray60"))
            self._set_rename_subject(info["sku"], info["name"])

    def _set_rename_subject(self, sku: str, name: str):
        self._rename_current_sku = sku.upper().strip()
        self._rename_current_name = name
        self._rename_new_var.set("")
        if self._rename_status_lbl:
            self._rename_status_lbl.configure(text="", text_color=("gray50", "gray60"))
        if self._rename_current_sku_lbl:
            self._rename_current_sku_lbl.configure(text=self._rename_current_sku or "—")
        if self._rename_current_name_lbl:
            self._rename_current_name_lbl.configure(text=name or "")
        # Auto-check alias checkbox if a mapping exists for this SKU
        if self._rename_current_sku and self._manager.has(self._rename_current_sku):
            self._rename_update_alias_var.set(True)
        else:
            self._rename_update_alias_var.set(False)

    def _do_rename(self):
        old_sku = self._rename_current_sku
        new_sku = self._rename_new_var.get().strip().upper()
        if not old_sku:
            self._rename_status_lbl.configure(text="Select a SKU first.", text_color="orange")
            return
        if not new_sku:
            self._rename_status_lbl.configure(text="Enter a new SKU.", text_color="orange")
            return
        if new_sku == old_sku:
            self._rename_status_lbl.configure(text="New SKU is the same as current SKU.", text_color="orange")
            return
        if not self._neto_client and not self._dry_run:
            self._rename_status_lbl.configure(text="No Neto connection available.", text_color="orange")
            return

        self._rename_apply_btn.configure(state="disabled")
        self._rename_status_lbl.configure(text="Renaming…", text_color=("gray50", "gray60"))

        def _work():
            if self._neto_client:
                ok, msg = self._neto_client.rename_item_sku(old_sku, new_sku, dry_run=self._dry_run)
            else:
                ok, msg = (True, f"[DRY RUN] Would rename '{old_sku}' → '{new_sku}'")
            self.after(0, lambda: self._on_rename_done(old_sku, new_sku, ok, msg))

        threading.Thread(target=_work, daemon=True).start()

    def _on_rename_done(self, old_sku: str, new_sku: str, success: bool, msg: str):
        self._rename_apply_btn.configure(state="normal")
        if success:
            self._rename_status_lbl.configure(text=f"✓ {msg}", text_color="green")
            if self._rename_update_alias_var.get():
                self._manager.rename_key(old_sku, new_sku)
            if self._on_sku_renamed:
                self._on_sku_renamed(old_sku, new_sku)
        else:
            self._rename_status_lbl.configure(text=f"✗ {msg}", text_color="red")

    # ── Search (alias tab, search mode) ───────────────────────────────────

    def _do_lookup(self):
        sku = self._search_var.get().strip()
        if not sku:
            return
        sku_upper = sku.upper()
        if sku_upper in self._cards:
            self._lookup_status.configure(text=f"'{sku}' is already shown below.", text_color="gray50")
            return

        self._lookup_btn.configure(state="disabled")
        self._lookup_status.configure(text="Looking up…", text_color=("gray50", "gray60"))

        def _fetch():
            name = None
            if self._neto_client:
                try:
                    name = self._neto_client.get_item_name(sku)
                except Exception as exc:
                    log.warning("SkuAliasModal: get_item_name failed: %s", exc)
            self.after(0, lambda: self._on_lookup_done(sku, name))

        threading.Thread(target=_fetch, daemon=True).start()

    def _on_lookup_done(self, sku: str, name: str | None):
        self._lookup_btn.configure(state="normal")
        if name is None:
            self._lookup_status.configure(
                text=f"'{sku}' not found in Neto — you can still create an alias.",
                text_color="orange",
            )
        else:
            self._lookup_status.configure(text="", text_color="gray50")
        self._add_card(sku, name or "")

    # ── Card management ───────────────────────────────────────────────────

    def _add_card(self, sku: str, description: str):
        sku_upper = sku.upper().strip()
        existing = self._manager.get_all().get(sku_upper)

        card = _CardState(
            sku=sku_upper,
            description=description,
            existing_mapping=existing,
        )
        self._cards[sku_upper] = card
        card.frame = self._build_card(self._scroll, card)
        card.frame.pack(fill="x", padx=4, pady=(0, 8))

    def _build_card(self, parent, card: _CardState) -> ctk.CTkFrame:
        outer = ctk.CTkFrame(parent, border_width=1, border_color=("gray65", "gray45"), corner_radius=6)
        outer.grid_columnconfigure(0, weight=1)

        # ── Card header ───────────────────────────────────────────────────
        hdr = ctk.CTkFrame(outer, fg_color="transparent")
        hdr.pack(fill="x", padx=10, pady=(8, 4))

        ctk.CTkLabel(
            hdr,
            text=card.sku,
            font=ctk.CTkFont(size=13, weight="bold"),
            anchor="w",
        ).pack(side="left")

        if card.description:
            ctk.CTkLabel(
                hdr,
                text=f"  —  {card.description}",
                font=ctk.CTkFont(size=12),
                text_color=("gray45", "gray65"),
                anchor="w",
            ).pack(side="left")

        # ── Kit toggle ────────────────────────────────────────────────────
        kit_frame = ctk.CTkFrame(outer, fg_color="transparent")
        kit_frame.pack(fill="x", padx=10, pady=(0, 4))

        card.kit_var = ctk.BooleanVar(value=card.existing_mapping["is_kit"] if card.existing_mapping else False)

        ctk.CTkCheckBox(
            kit_frame,
            text="Kit item  (maps to multiple supplier SKUs)",
            variable=card.kit_var,
            font=ctk.CTkFont(size=12),
            command=lambda c=card: self._on_kit_toggle(c),
        ).pack(side="left")

        # ── Supplier dropdown ─────────────────────────────────────────────
        sup_frame = ctk.CTkFrame(outer, fg_color="transparent")
        sup_frame.pack(fill="x", padx=10, pady=(0, 4))

        ctk.CTkLabel(
            sup_frame,
            text="Supplier:",
            font=ctk.CTkFont(size=12),
            width=70,
            anchor="w",
        ).pack(side="left")

        existing_supplier_name = (card.existing_mapping or {}).get("supplier", "")
        initial_sup_label = "— No supplier —"
        for lbl, name in self._supplier_name_map.items():
            if name == existing_supplier_name:
                initial_sup_label = lbl
                break
        card.supplier_var = ctk.StringVar(value=initial_sup_label)
        ctk.CTkOptionMenu(
            sup_frame,
            variable=card.supplier_var,
            values=self._supplier_options,
            font=ctk.CTkFont(size=12),
            width=320,
            dynamic_resizing=False,
        ).pack(side="left")

        # ── SKU inputs area ───────────────────────────────────────────────
        card.inputs_frame = ctk.CTkFrame(outer, fg_color="transparent")
        card.inputs_frame.pack(fill="x", padx=10, pady=(0, 4))

        card.prompt_label = ctk.CTkLabel(
            card.inputs_frame,
            text=self._prompt_text(card.kit_var.get()),
            font=ctk.CTkFont(size=12),
            anchor="w",
        )
        card.prompt_label.pack(fill="x", pady=(0, 4))

        if card.existing_mapping:
            initial_skus = card.existing_mapping["invoice_skus"]
            initial_qtys = card.existing_mapping.get("qty_per_alias") or [1] * len(initial_skus)
        else:
            initial_skus = []
            initial_qtys = []

        card.sku_rows_frame = ctk.CTkFrame(card.inputs_frame, fg_color="transparent")
        card.sku_rows_frame.pack(fill="x")

        card.entry_vars = []
        card.qty_vars = []
        card.entry_frames = []
        if initial_skus:
            for i, s in enumerate(initial_skus):
                q = initial_qtys[i] if i < len(initial_qtys) else 1
                self._add_sku_row(card, initial_value=s, initial_qty=q)
        else:
            self._add_sku_row(card, initial_value="", initial_qty=1)

        card.add_btn = ctk.CTkButton(
            card.inputs_frame,
            text="＋ Add SKU",
            width=100,
            height=26,
            font=ctk.CTkFont(size=12),
            fg_color=("gray70", "gray30"),
            hover_color=("gray60", "gray25"),
            command=lambda c=card: self._add_sku_row(c),
        )
        if card.kit_var.get():
            card.add_btn.pack(anchor="w", pady=(4, 0))

        # ── Remove mapping button ─────────────────────────────────────────
        if card.existing_mapping:
            btn_frame = ctk.CTkFrame(outer, fg_color="transparent")
            btn_frame.pack(fill="x", padx=10, pady=(4, 8))
            ctk.CTkButton(
                btn_frame,
                text="Remove mapping",
                width=140,
                height=26,
                font=ctk.CTkFont(size=12),
                fg_color=("gray70", "gray30"),
                hover_color=("firebrick3", "firebrick4"),
                command=lambda c=card: self._remove_mapping(c),
            ).pack(side="left")
        else:
            ctk.CTkFrame(outer, height=4, fg_color="transparent").pack()

        return outer

    def _add_sku_row(self, card: _CardState, initial_value: str = "", initial_qty: int = 1):
        row = ctk.CTkFrame(card.sku_rows_frame, fg_color="transparent")
        row.pack(fill="x", pady=2)

        var = ctk.StringVar(value=initial_value)
        qty_var = ctk.StringVar(value=str(initial_qty))
        card.entry_vars.append(var)
        card.qty_vars.append(qty_var)
        card.entry_frames.append(row)

        entry = ctk.CTkEntry(
            row,
            textvariable=var,
            placeholder_text="Supplier / invoice SKU",
            font=ctk.CTkFont(size=12),
            width=220,
        )
        entry.pack(side="left", padx=(0, 6))

        ctk.CTkLabel(
            row,
            text="×",
            font=ctk.CTkFont(size=13),
            text_color=("gray50", "gray60"),
        ).pack(side="left", padx=(0, 4))

        def _decrement(qv=qty_var):
            try:
                qv.set(str(max(1, int(qv.get()) - 1)))
            except ValueError:
                qv.set("1")

        def _increment(qv=qty_var):
            try:
                qv.set(str(int(qv.get()) + 1))
            except ValueError:
                qv.set("1")

        ctk.CTkButton(
            row,
            text="−",
            width=26,
            height=26,
            font=ctk.CTkFont(size=14),
            fg_color=("gray70", "gray30"),
            hover_color=("gray60", "gray25"),
            command=_decrement,
        ).pack(side="left", padx=(0, 2))

        ctk.CTkEntry(
            row,
            textvariable=qty_var,
            font=ctk.CTkFont(size=12),
            width=46,
            justify="center",
        ).pack(side="left", padx=(0, 2))

        ctk.CTkButton(
            row,
            text="+",
            width=26,
            height=26,
            font=ctk.CTkFont(size=14),
            fg_color=("gray70", "gray30"),
            hover_color=("gray60", "gray25"),
            command=_increment,
        ).pack(side="left", padx=(0, 8))

        delete_btn = ctk.CTkButton(
            row,
            text="🗑",
            width=28,
            height=28,
            font=ctk.CTkFont(size=13),
            fg_color=("gray70", "gray30"),
            hover_color=("firebrick3", "firebrick4"),
        )
        delete_btn.configure(
            command=lambda r=row, v=var, q=qty_var, c=card: self._remove_sku_row(c, r, v, q)
        )
        delete_btn.pack(side="left")

    def _remove_sku_row(self, card: _CardState, row_frame, var: ctk.StringVar, qty_var: ctk.StringVar):
        if len(card.entry_vars) <= 1:
            var.set("")
            qty_var.set("1")
            return
        idx = card.entry_vars.index(var)
        card.entry_vars.pop(idx)
        card.qty_vars.pop(idx)
        card.entry_frames.pop(idx)
        row_frame.destroy()

    def _on_kit_toggle(self, card: _CardState):
        is_kit = card.kit_var.get()
        card.prompt_label.configure(text=self._prompt_text(is_kit))
        if is_kit:
            card.add_btn.pack(anchor="w", pady=(4, 0))
        else:
            card.add_btn.pack_forget()
            while len(card.entry_vars) > 1:
                self._remove_sku_row(card, card.entry_frames[-1], card.entry_vars[-1])

    @staticmethod
    def _prompt_text(is_kit: bool) -> str:
        if is_kit:
            return "What are the SKUs of the items in this kit?"
        return "What is the supplier SKU for this item?"

    # ── Remove mapping ────────────────────────────────────────────────────

    def _remove_mapping(self, card: _CardState):
        if not messagebox.askyesno(
            "Remove Alias",
            f"Remove the alias mapping for '{card.sku}'?",
            parent=self,
        ):
            return
        try:
            self._manager.remove(card.sku)
        except Exception as exc:
            messagebox.showerror("Error", f"Failed to remove mapping:\n{exc}", parent=self)
            return
        card.existing_mapping = None
        while len(card.entry_vars) > 1:
            self._remove_sku_row(card, card.entry_frames[-1], card.entry_vars[-1])
        card.entry_vars[0].set("")
        if card.qty_vars:
            card.qty_vars[0].set("1")
        card.kit_var.set(False)
        if card.supplier_var:
            card.supplier_var.set("— No supplier —")
        self._on_kit_toggle(card)
        card.frame.destroy()
        card.frame = self._build_card(self._scroll, card)
        card.frame.pack(fill="x", padx=4, pady=(0, 8))

    # ── Save ─────────────────────────────────────────────────────────────

    def _save_all(self):
        errors: list[str] = []
        saved = 0

        for sku_upper, card in self._cards.items():
            pairs = [
                (v.get().strip(), q.get().strip())
                for v, q in zip(card.entry_vars, card.qty_vars)
                if v.get().strip()
            ]
            if not pairs:
                continue
            invoice_skus = [sku for sku, _ in pairs]
            qty_per_alias = []
            for _, q in pairs:
                try:
                    qty_per_alias.append(max(1, int(q)))
                except (ValueError, TypeError):
                    qty_per_alias.append(1)
            try:
                sup_label = card.supplier_var.get() if card.supplier_var else "— No supplier —"
                supplier_name = self._supplier_name_map.get(sup_label, "")
                self._manager.save(
                    sku_upper, invoice_skus, card.kit_var.get(), supplier_name,
                    qty_per_alias=qty_per_alias,
                )
                saved += 1
            except Exception as exc:
                errors.append(f"{sku_upper}: {exc}")

        if errors:
            messagebox.showerror(
                "Save Error",
                "Some mappings could not be saved:\n\n" + "\n".join(errors),
                parent=self,
            )
        else:
            self.destroy()


class _CardState:
    """Internal state for a single alias card in the modal."""
    def __init__(self, sku: str, description: str, existing_mapping: dict | None):
        self.sku = sku
        self.description = description
        self.existing_mapping = existing_mapping

        self.frame: ctk.CTkFrame | None = None
        self.kit_var: ctk.BooleanVar | None = None
        self.supplier_var: ctk.StringVar | None = None
        self.prompt_label: ctk.CTkLabel | None = None
        self.inputs_frame: ctk.CTkFrame | None = None
        self.sku_rows_frame: ctk.CTkFrame | None = None
        self.add_btn: ctk.CTkButton | None = None
        self.entry_vars: list[ctk.StringVar] = []
        self.qty_vars: list[ctk.StringVar] = []
        self.entry_frames: list[ctk.CTkFrame] = []
