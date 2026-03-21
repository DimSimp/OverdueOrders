from __future__ import annotations

import threading
import tkinter as tk

import customtkinter as ctk


# Dialog state constants
_RESOLVING = "resolving"
_DISAMBIGUATE = "disambiguate"
_CONFIRM = "confirm"
_NOT_FOUND = "not_found"
_CONFIRM_ALIAS = "confirm_alias"
_WORKING = "working"
_DONE = "done"
_ERROR = "error"


class MusiposPODialog(ctk.CTkToplevel):
    """
    Multi-step dialog for the "On PO" workflow.

    Steps:
        RESOLVING       → background SKU resolution
        CONFIRM         → show item info, qty stepper, [Add to PO] / [Cancel]
        NOT_FOUND       → manual Musipos SKU entry
        CONFIRM_ALIAS   → confirm manually-found item, offer to save alias
        WORKING         → adding to PO
        DONE            → success display
        ERROR           → failure display
    """

    def __init__(
        self,
        parent,
        *,
        neto_sku: str,
        product_name: str,
        order_qty: int,
        musipos_client,
        suppliers_config: list,
        dry_run: bool = False,
        on_success=None,   # called with po_result dict
        on_note_only=None, # called if user cancels PO but still wants note
    ):
        super().__init__(parent)
        self.title("Add to Purchase Order")
        self.geometry("520x420")
        self.resizable(False, False)
        self.grab_set()

        self._neto_sku = neto_sku
        self._product_name = product_name
        self._order_qty = max(1, order_qty)
        self._client = musipos_client
        self._suppliers_config = suppliers_config
        self._dry_run = dry_run
        self._on_success = on_success
        self._on_note_only = on_note_only

        # State
        self._item: dict | None = None              # resolved item dict
        self._suppliers: list[dict] = []             # [{supplier_id, supplier_name}]
        self._manual_item: dict | None = None       # result from manual search
        self._disambiguate_items: list[dict] = []   # multiple matches pending selection
        self._disambiguate_source: str = "auto"     # "auto" or "manual"
        self._qty_var = ctk.StringVar(value=str(self._order_qty))
        self._supplier_var = ctk.StringVar()
        self._manual_sku_var = ctk.StringVar()

        # Build a single content frame that we swap out per state
        self._content = ctk.CTkFrame(self, fg_color="transparent")
        self._content.pack(fill="both", expand=True, padx=16, pady=12)

        self._transition(_RESOLVING)

        # Centre on parent
        self.after(50, self._centre_on_parent)

    # ── State machine ─────────────────────────────────────────────────────

    def _transition(self, state: str):
        for w in self._content.winfo_children():
            w.destroy()
        if state == _RESOLVING:
            self._build_resolving()
        elif state == _DISAMBIGUATE:
            self._build_disambiguate()
        elif state == _CONFIRM:
            self._build_confirm()
        elif state == _NOT_FOUND:
            self._build_not_found()
        elif state == _CONFIRM_ALIAS:
            self._build_confirm_alias()
        elif state == _WORKING:
            self._build_working()
        elif state == _DONE:
            self._build_done()
        elif state == _ERROR:
            self._build_error()

    # ── RESOLVING ─────────────────────────────────────────────────────────

    def _build_resolving(self):
        ctk.CTkLabel(
            self._content,
            text=f"Resolving  {self._neto_sku}  in Musipos…",
            font=ctk.CTkFont(size=13),
        ).pack(pady=(30, 12))
        pb = ctk.CTkProgressBar(self._content, mode="indeterminate", width=300)
        pb.pack(pady=4)
        pb.start()
        threading.Thread(target=self._do_resolve, daemon=True).start()

    def _do_resolve(self):
        try:
            items = self._client.resolve_item_multi(self._neto_sku, self._suppliers_config)
        except Exception as exc:
            self.after(0, lambda: self._on_resolve_done([], str(exc)))
            return
        self.after(0, lambda: self._on_resolve_done(items, None))

    def _on_resolve_done(self, items, error):
        if error:
            self._error_msg = f"Error resolving SKU: {error}"
            self._transition(_ERROR)
            return
        if len(items) == 1:
            item = items[0]
            self._item = item
            self._suppliers = self._get_suppliers_for(item)
            if self._suppliers:
                self._supplier_var.set(self._suppliers[0]["supplier_id"])
            self._transition(_CONFIRM)
        elif len(items) > 1:
            self._disambiguate_items = items
            self._disambiguate_source = "auto"
            self._transition(_DISAMBIGUATE)
        else:
            self._transition(_NOT_FOUND)

    def _get_suppliers_for(self, item: dict) -> list[dict]:
        """Fetch supplier list for an item, falling back to item data."""
        try:
            suppliers = self._client.get_suppliers_for_item(item["itm_iid"])
        except Exception:
            suppliers = []
        if not suppliers and item.get("supplier_id"):
            sup_name = item.get("supplier_name") or item["supplier_id"]
            suppliers = [{"supplier_id": item["supplier_id"], "supplier_name": sup_name}]
        return suppliers

    # ── DISAMBIGUATE ──────────────────────────────────────────────────────

    def _build_disambiguate(self):
        items = self._disambiguate_items
        f = self._content

        ctk.CTkLabel(
            f, text=f"Multiple items match  '{self._neto_sku}'",
            font=ctk.CTkFont(size=13, weight="bold"),
        ).pack(anchor="w", pady=(4, 0))
        ctk.CTkLabel(
            f, text="Select the correct item:",
            font=ctk.CTkFont(size=12), text_color="gray60",
        ).pack(anchor="w", pady=(0, 6))

        scroll = ctk.CTkScrollableFrame(f, height=230)
        scroll.pack(fill="x", expand=False, pady=(0, 4))

        for item in items:
            sup_sku = item.get("supplier_iid") or item["itm_iid"]
            title = item.get("title") or item["itm_iid"]
            sup_name = item.get("supplier_name") or item.get("supplier_id", "")
            stock = item.get("qty_on_hand", 0)

            row_frame = ctk.CTkFrame(scroll, fg_color=("gray88", "gray22"), corner_radius=6)
            row_frame.pack(fill="x", pady=3, padx=2)
            row_frame.columnconfigure(1, weight=1)

            ctk.CTkLabel(
                row_frame, text=sup_sku,
                font=ctk.CTkFont(size=12, weight="bold"),
                width=90, anchor="w",
            ).grid(row=0, column=0, sticky="w", padx=(8, 4), pady=(6, 2))

            ctk.CTkLabel(
                row_frame, text=title,
                font=ctk.CTkFont(size=12), anchor="w", wraplength=260,
            ).grid(row=0, column=1, sticky="w", padx=4, pady=(6, 2))

            ctk.CTkLabel(
                row_frame, text=f"{sup_name}  ·  {stock} in stock",
                font=ctk.CTkFont(size=11), anchor="w", text_color="gray60",
            ).grid(row=1, column=0, columnspan=2, sticky="w", padx=8, pady=(0, 6))

            def _pick(i=item):
                self._select_from_disambiguate(i)

            ctk.CTkButton(
                row_frame, text="Select", width=70,
                command=_pick,
            ).grid(row=0, column=2, rowspan=2, padx=8, pady=4)

        btn_row = ctk.CTkFrame(f, fg_color="transparent")
        btn_row.pack(fill="x", pady=(4, 0))
        ctk.CTkButton(
            btn_row, text="Search manually", width=140,
            fg_color=("gray70", "gray30"), hover_color=("gray60", "gray25"),
            command=lambda: self._transition(_NOT_FOUND),
        ).pack(side="left", padx=(0, 8))
        ctk.CTkButton(
            btn_row, text="Cancel", width=80,
            fg_color="gray50", hover_color="gray40",
            command=self._cancel,
        ).pack(side="left")

    def _select_from_disambiguate(self, item: dict):
        self._suppliers = self._get_suppliers_for(item)
        if self._suppliers:
            self._supplier_var.set(self._suppliers[0]["supplier_id"])
        if self._disambiguate_source == "manual":
            self._manual_item = item
            self._transition(_CONFIRM_ALIAS)
        else:
            self._item = item
            self._transition(_CONFIRM)

    # ── CONFIRM ───────────────────────────────────────────────────────────

    def _build_confirm(self):
        item = self._item
        f = self._content

        # Item title
        ctk.CTkLabel(f, text=item["title"] or item["itm_iid"],
                     font=ctk.CTkFont(size=14, weight="bold"),
                     wraplength=480).pack(anchor="w", pady=(4, 0))

        # Info grid
        grid = ctk.CTkFrame(f, fg_color="transparent")
        grid.pack(fill="x", pady=6)
        grid.columnconfigure(1, weight=1)

        on_hand = item["qty_on_hand"]
        on_order = item.get("qty_on_order", 0)
        stock_color = "red" if on_hand == 0 else ("gray20", "gray80")
        on_order_color = ("darkorange3", "orange") if on_order > 0 else ("gray20", "gray80")

        musipos_sku = item.get("supplier_iid") or item["itm_iid"]
        rows = [
            ("Musipos SKU:", musipos_sku),
            ("Stock:", f"{on_hand} on hand"),
            ("On Order:", f"{on_order} currently on PO"),
        ]
        if self._dry_run:
            rows.append(("Mode:", "[DRY RUN — no DB writes]"))

        color_map = {"Stock:": stock_color, "On Order:": on_order_color}

        for r, (lbl, val) in enumerate(rows):
            ctk.CTkLabel(grid, text=lbl, font=ctk.CTkFont(size=12, weight="bold"),
                         width=100, anchor="w").grid(row=r, column=0, sticky="w", pady=1)
            tc = color_map.get(lbl, ("gray20", "gray80"))
            ctk.CTkLabel(grid, text=val, font=ctk.CTkFont(size=12),
                         anchor="w", text_color=tc).grid(row=r, column=1, sticky="w", pady=1)

        # Supplier row
        sup_row = ctk.CTkFrame(f, fg_color="transparent")
        sup_row.pack(fill="x", pady=2)
        ctk.CTkLabel(sup_row, text="Supplier:", font=ctk.CTkFont(size=12, weight="bold"),
                     width=100, anchor="w").pack(side="left")
        if len(self._suppliers) > 1:
            sup_ids = [s["supplier_id"] for s in self._suppliers]
            ctk.CTkOptionMenu(sup_row, values=sup_ids, variable=self._supplier_var,
                              width=200, font=ctk.CTkFont(size=12)).pack(side="left")
        else:
            lbl = self._suppliers[0]["supplier_name"] if self._suppliers else "Unknown"
            ctk.CTkLabel(sup_row, text=lbl, font=ctk.CTkFont(size=12)).pack(side="left")

        # Qty stepper
        qty_row = ctk.CTkFrame(f, fg_color="transparent")
        qty_row.pack(fill="x", pady=6)
        ctk.CTkLabel(qty_row, text="Quantity to order:",
                     font=ctk.CTkFont(size=12, weight="bold")).pack(side="left", padx=(0, 10))
        ctk.CTkButton(qty_row, text="−", width=28, height=28, font=ctk.CTkFont(size=14),
                      fg_color=("gray70", "gray30"), hover_color=("gray60", "gray25"),
                      command=self._decrement_qty).pack(side="left", padx=(0, 2))
        ctk.CTkEntry(qty_row, textvariable=self._qty_var, width=52,
                     justify="center", font=ctk.CTkFont(size=13)).pack(side="left", padx=(0, 2))
        ctk.CTkButton(qty_row, text="+", width=28, height=28, font=ctk.CTkFont(size=14),
                      fg_color=("gray70", "gray30"), hover_color=("gray60", "gray25"),
                      command=self._increment_qty).pack(side="left")

        # Check whether a kit/alias mapping exists for this Neto SKU (supplier IIDs)
        alias_map = self._client.load_kit_mappings()
        _alias_skus = alias_map.get(self._neto_sku.upper().strip(), [])

        # Buttons
        btn_row = ctk.CTkFrame(f, fg_color="transparent")
        btn_row.pack(fill="x", pady=(10, 0))
        self._add_btn = ctk.CTkButton(
            btn_row, text="Add to PO", width=120,
            fg_color=("green3", "green4"), hover_color=("green4", "green3"),
            command=self._do_add_to_po,
        )
        self._add_btn.pack(side="left", padx=(0, 10))
        ctk.CTkButton(btn_row, text="Cancel", width=80,
                      fg_color="gray50", hover_color="gray40",
                      command=self._cancel).pack(side="left")
        # Right-side escape hatches (packed right-to-left)
        ctk.CTkButton(btn_row, text="Not the correct item?", width=150,
                      fg_color="transparent", border_width=1,
                      hover_color=("gray80", "gray25"),
                      command=lambda: self._transition(_NOT_FOUND)).pack(side="right")
        if _alias_skus:
            ctk.CTkButton(btn_row, text="Use SKU alias", width=110,
                          fg_color="transparent", border_width=1,
                          hover_color=("gray80", "gray25"),
                          command=lambda a=_alias_skus: self._try_alias(a)).pack(side="right", padx=(0, 6))

    def _decrement_qty(self):
        try:
            self._qty_var.set(str(max(1, int(self._qty_var.get()) - 1)))
        except ValueError:
            self._qty_var.set("1")

    def _increment_qty(self):
        try:
            self._qty_var.set(str(int(self._qty_var.get()) + 1))
        except ValueError:
            self._qty_var.set("1")

    # ── NOT_FOUND ─────────────────────────────────────────────────────────

    def _build_not_found(self):
        f = self._content
        ctk.CTkLabel(f, text=f"'{self._neto_sku}' was not found in Musipos.",
                     font=ctk.CTkFont(size=13, weight="bold"),
                     text_color=("firebrick3", "firebrick1")).pack(anchor="w", pady=(4, 2))
        ctk.CTkLabel(f, text="Search by Supplier SKU, Musipos internal ID, or barcode:",
                     font=ctk.CTkFont(size=12)).pack(anchor="w", pady=(0, 6))

        entry_row = ctk.CTkFrame(f, fg_color="transparent")
        entry_row.pack(fill="x", pady=2)
        entry = ctk.CTkEntry(entry_row, textvariable=self._manual_sku_var,
                             width=260, font=ctk.CTkFont(size=13),
                             placeholder_text="e.g. J812 4/4M")
        entry.pack(side="left", padx=(0, 8))
        entry.focus_set()
        self._manual_search_btn = ctk.CTkButton(
            entry_row, text="Search", width=80,
            command=self._do_manual_search,
        )
        self._manual_search_btn.pack(side="left")
        entry.bind("<Return>", lambda e: self._do_manual_search())

        self._manual_status = ctk.CTkLabel(f, text="", font=ctk.CTkFont(size=11),
                                           text_color=("firebrick3", "firebrick1"))
        self._manual_status.pack(anchor="w", pady=(2, 0))

        ctk.CTkButton(
            f, text="Cancel — add note only", width=180,
            fg_color="gray50", hover_color="gray40",
            command=self._cancel_note_only,
        ).pack(anchor="w", pady=(14, 0))

    def _do_manual_search(self):
        sku = self._manual_sku_var.get().strip()
        if not sku:
            return
        self._manual_search_btn.configure(state="disabled", text="Searching…")
        self._manual_status.configure(text="")

        def _work():
            try:
                items = self._client.resolve_manual_multi(sku)
            except Exception as exc:
                self.after(0, lambda: self._on_manual_done([], str(exc)))
                return
            self.after(0, lambda: self._on_manual_done(items, None))

        threading.Thread(target=_work, daemon=True).start()

    def _on_manual_done(self, items, error):
        self._manual_search_btn.configure(state="normal", text="Search")
        if error:
            self._manual_status.configure(text=f"Error: {error}")
            return
        if not items:
            self._manual_status.configure(text="Not found. Try a different SKU.")
            return
        if len(items) == 1:
            item = items[0]
            self._manual_item = item
            self._suppliers = self._get_suppliers_for(item)
            if self._suppliers:
                self._supplier_var.set(self._suppliers[0]["supplier_id"])
            self._transition(_CONFIRM_ALIAS)
        else:
            self._disambiguate_items = items
            self._disambiguate_source = "manual"
            self._transition(_DISAMBIGUATE)

    # ── CONFIRM_ALIAS ─────────────────────────────────────────────────────

    def _build_confirm_alias(self):
        item = self._manual_item
        f = self._content
        ctk.CTkLabel(f, text="Found a match. Is this the correct item?",
                     font=ctk.CTkFont(size=13, weight="bold")).pack(anchor="w", pady=(4, 8))

        info = ctk.CTkFrame(f, fg_color="transparent")
        info.pack(fill="x", pady=2)
        info.columnconfigure(1, weight=1)
        for r, (lbl, val) in enumerate([
            ("Musipos ID:", item["itm_iid"]),
            ("Description:", item["title"]),
            ("Supplier:", self._suppliers[0]["supplier_name"] if self._suppliers else ""),
        ]):
            ctk.CTkLabel(info, text=lbl, font=ctk.CTkFont(size=12, weight="bold"),
                         width=110, anchor="w").grid(row=r, column=0, sticky="w", pady=1)
            ctk.CTkLabel(info, text=val, font=ctk.CTkFont(size=12),
                         anchor="w", wraplength=330).grid(row=r, column=1, sticky="w", pady=1)

        btn_row = ctk.CTkFrame(f, fg_color="transparent")
        btn_row.pack(fill="x", pady=(14, 0))
        ctk.CTkButton(
            btn_row, text="Yes — add to PO & save alias", width=210,
            fg_color=("green3", "green4"), hover_color=("green4", "green3"),
            command=self._confirm_alias_and_proceed,
        ).pack(side="left", padx=(0, 8))
        ctk.CTkButton(
            btn_row, text="Try again", width=90,
            fg_color=("gray70", "gray30"), hover_color=("gray60", "gray25"),
            command=lambda: self._transition(_NOT_FOUND),
        ).pack(side="left", padx=(0, 8))
        ctk.CTkButton(
            btn_row, text="Cancel", width=80,
            fg_color="gray50", hover_color="gray40",
            command=self._cancel,
        ).pack(side="left")

    def _confirm_alias_and_proceed(self):
        # Save alias mapping
        try:
            self._client.save_musipos_alias(self._neto_sku, self._manual_item["itm_iid"])
        except Exception as exc:
            print(f"[MusiposDialog] Failed to save alias: {exc}")
        # Promote manual item to resolved item and go to confirm
        self._item = self._manual_item
        self._transition(_CONFIRM)

    def _try_alias(self, alias_skus: list):
        """Re-resolve using supplier IID aliases from sku_mappings.csv."""
        for w in self._content.winfo_children():
            w.destroy()
        ctk.CTkLabel(self._content, text="Looking up SKU alias…",
                     font=ctk.CTkFont(size=13)).pack(pady=(30, 12))
        pb = ctk.CTkProgressBar(self._content, mode="indeterminate", width=300)
        pb.pack(pady=4)
        pb.start()

        def _work():
            found = {}  # itm_iid → item dict (deduplicates across multiple aliases)
            for sku in alias_skus:
                try:
                    items = self._client.resolve_manual_multi(sku)
                except Exception:
                    items = []
                for item in items:
                    found[item["itm_iid"]] = item
            self.after(0, lambda: self._on_alias_done(list(found.values())))

        threading.Thread(target=_work, daemon=True).start()

    def _on_alias_done(self, items: list):
        if not items:
            self._transition(_NOT_FOUND)
            return
        if len(items) == 1:
            item = items[0]
            self._item = item
            self._suppliers = self._get_suppliers_for(item)
            if self._suppliers:
                self._supplier_var.set(self._suppliers[0]["supplier_id"])
            self._transition(_CONFIRM)
        else:
            self._disambiguate_items = items
            self._disambiguate_source = "auto"
            self._transition(_DISAMBIGUATE)

    # ── WORKING ───────────────────────────────────────────────────────────

    def _build_working(self):
        ctk.CTkLabel(self._content, text="Adding to PO…",
                     font=ctk.CTkFont(size=13)).pack(pady=(30, 12))
        pb = ctk.CTkProgressBar(self._content, mode="indeterminate", width=300)
        pb.pack(pady=4)
        pb.start()

    def _do_add_to_po(self):
        try:
            qty = max(1, int(self._qty_var.get()))
        except ValueError:
            qty = 1
        supplier_id = self._supplier_var.get()
        if not supplier_id and self._suppliers:
            supplier_id = self._suppliers[0]["supplier_id"]

        self._transition(_WORKING)

        def _work():
            try:
                result = self._client.add_to_po(
                    itm_iid=self._item["itm_iid"],
                    supplier_id=supplier_id,
                    qty=qty,
                    item_dict=self._item,
                    dry_run=self._dry_run,
                )
                self.after(0, lambda: self._on_add_done(result, None))
            except Exception as exc:
                self.after(0, lambda: self._on_add_done(None, str(exc)))

        threading.Thread(target=_work, daemon=True).start()

    def _on_add_done(self, result, error):
        if error:
            self._error_msg = f"Failed to add to PO:\n{error}"
            self._transition(_ERROR)
            return
        self._po_result = result
        if self._on_success:
            try:
                self._on_success(result)
            except Exception:
                pass
        self._transition(_DONE)

    # ── DONE ──────────────────────────────────────────────────────────────

    def _build_done(self):
        r = self._po_result
        f = self._content
        ctk.CTkLabel(f, text="✓", font=ctk.CTkFont(size=28),
                     text_color="green").pack(pady=(20, 4))
        if r["action"] == "added":
            msg = f"Added {self._item['itm_iid']} to PO #{r['po_no']}  ({r['supplier_name']})"
        else:
            msg = f"Updated PO #{r['po_no']}  —  qty +{r['qty_added']}  ({r['supplier_name']})"
        if self._dry_run:
            msg += "\n[DRY RUN — no changes were written]"
        ctk.CTkLabel(f, text=msg, font=ctk.CTkFont(size=13),
                     wraplength=460, justify="center").pack(pady=4)
        ctk.CTkButton(f, text="Close", width=90,
                      command=self.destroy).pack(pady=(16, 0))

    # ── ERROR ─────────────────────────────────────────────────────────────

    def _build_error(self):
        f = self._content
        ctk.CTkLabel(f, text="✗", font=ctk.CTkFont(size=28),
                     text_color="red").pack(pady=(20, 4))
        ctk.CTkLabel(f, text=getattr(self, "_error_msg", "An error occurred."),
                     font=ctk.CTkFont(size=12), text_color="red",
                     wraplength=460, justify="left").pack(pady=4)
        ctk.CTkButton(f, text="Close", width=90,
                      command=self.destroy).pack(pady=(16, 0))

    # ── Shared helpers ────────────────────────────────────────────────────

    def _cancel(self):
        self.destroy()

    def _cancel_note_only(self):
        """User declined PO but still wants an order note added."""
        if self._on_note_only:
            try:
                self._on_note_only()
            except Exception:
                pass
        self.destroy()

    def _centre_on_parent(self):
        try:
            pw = self.master.winfo_width()
            ph = self.master.winfo_height()
            px = self.master.winfo_rootx()
            py = self.master.winfo_rooty()
            w, h = 520, 420
            x = px + (pw - w) // 2
            y = py + (ph - h) // 2
            self.geometry(f"{w}x{h}+{x}+{y}")
        except Exception:
            pass
