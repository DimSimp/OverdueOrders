from __future__ import annotations

import logging
import threading
from tkinter import messagebox

import customtkinter as ctk

log = logging.getLogger(__name__)


class SkuAliasModal(ctk.CTkToplevel):
    """
    Modal for creating / editing / removing SKU alias mappings.

    Two modes
    ---------
    "line_items"
        Opens with a pre-populated list of (sku, description) tuples from an order.
        Shows a card for each line item.

    "search"
        Opens with an empty search bar. User types a Neto SKU, clicks Look up (or
        presses Enter), and the app calls neto_client.get_item_name() in a background
        thread. A single card appears for the looked-up SKU.
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
    ):
        super().__init__(master)
        self._manager = sku_alias_manager
        self._mode = mode
        self._line_items = line_items or []
        self._neto_client = neto_client

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

        self.title("Manage SKU Aliases")
        self.geometry("640x560")
        self.minsize(540, 400)
        self.resizable(True, True)
        self.grab_set()
        self.lift()
        self.focus_force()

        self._build_ui()

        # In line_items mode, populate cards immediately
        if self._mode == "line_items":
            for sku, desc in self._line_items:
                self._add_card(sku, desc)

    # ── Layout ────────────────────────────────────────────────────────────

    def _build_ui(self):
        self.grid_rowconfigure(1, weight=1)
        self.grid_columnconfigure(0, weight=1)

        # ── Search bar (search mode only) ─────────────────────────────────
        if self._mode == "search":
            search_frame = ctk.CTkFrame(self, fg_color=("gray90", "gray20"), corner_radius=0)
            search_frame.grid(row=0, column=0, sticky="ew", padx=0, pady=0)
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
            # Spacer row for line_items mode
            ctk.CTkFrame(self, height=0, fg_color="transparent").grid(row=0, column=0)

        # ── Scrollable card area ──────────────────────────────────────────
        self._scroll = ctk.CTkScrollableFrame(self, fg_color="transparent")
        self._scroll.grid(row=1, column=0, sticky="nsew", padx=8, pady=(8, 0))
        self._scroll.grid_columnconfigure(0, weight=1)

        # ── Footer ────────────────────────────────────────────────────────
        footer = ctk.CTkFrame(self, fg_color="transparent")
        footer.grid(row=2, column=0, sticky="ew", padx=12, pady=(8, 12))

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

    # ── Search (search mode) ──────────────────────────────────────────────

    def _do_lookup(self):
        sku = self._search_var.get().strip()
        if not sku:
            return
        sku_upper = sku.upper()
        if sku_upper in self._cards:
            # Already showing a card for this SKU
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

        # Determine initial supplier label
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

        # Determine initial invoice SKUs and quantities
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

        # "＋ Add SKU" button (only visible in kit mode)
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

        ctk.CTkEntry(
            row,
            textvariable=qty_var,
            font=ctk.CTkFont(size=12),
            width=52,
            justify="center",
        ).pack(side="left", padx=(0, 6))

        # "−" remove button
        remove_btn = ctk.CTkButton(
            row,
            text="−",
            width=28,
            height=28,
            font=ctk.CTkFont(size=14),
            fg_color=("gray70", "gray30"),
            hover_color=("firebrick3", "firebrick4"),
        )
        remove_btn.configure(
            command=lambda r=row, v=var, q=qty_var, c=card: self._remove_sku_row(c, r, v, q)
        )
        remove_btn.pack(side="left")

    def _remove_sku_row(self, card: _CardState, row_frame, var: ctk.StringVar, qty_var: ctk.StringVar):
        if len(card.entry_vars) <= 1:
            # Don't remove the last field — just clear it
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
            # Collapse to a single entry row
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
        # Update card state
        card.existing_mapping = None
        # Clear entries
        while len(card.entry_vars) > 1:
            self._remove_sku_row(card, card.entry_frames[-1], card.entry_vars[-1])
        card.entry_vars[0].set("")
        if card.qty_vars:
            card.qty_vars[0].set("1")
        card.kit_var.set(False)
        if card.supplier_var:
            card.supplier_var.set("— No supplier —")
        self._on_kit_toggle(card)
        # Rebuild the card (simplest way to remove the Remove button)
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
                # No input — skip (don't save an empty mapping)
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

        # Set by _build_card
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
