import threading
from tkinter import filedialog

import customtkinter as ctk

from src.pdf_parser import InvoiceItem, ParseError, parse_invoice


class EditableTable(ctk.CTkScrollableFrame):
    """Scrollable grid of editable entry cells for invoice items."""

    HEADERS = ["SKU (with suffix)", "Description", "Qty"]
    COL_WIDTHS = [200, 430, 60]
    COL_WEIGHTS = [1, 3, 0]

    def __init__(self, master, **kwargs):
        super().__init__(master, **kwargs)
        self._rows: list[list[ctk.CTkEntry]] = []
        self._render_headers()

    def _render_headers(self):
        for col, (header, width) in enumerate(zip(self.HEADERS, self.COL_WIDTHS)):
            lbl = ctk.CTkLabel(
                self,
                text=header,
                font=ctk.CTkFont(weight="bold"),
                width=width,
                anchor="w",
            )
            lbl.grid(row=0, column=col, padx=(4, 8), pady=(4, 6), sticky="w")
        self.grid_columnconfigure(1, weight=1)

    def load_items(self, items: list[InvoiceItem], append: bool = True):
        """Add items to the table. If append=True, adds to existing rows."""
        start_row = len(self._rows) + 1
        for i, item in enumerate(items):
            flag = " ⚠" if item.qty_flagged else ""
            self._add_row(
                start_row + i,
                item.sku_with_suffix,
                item.description,
                str(item.quantity) + flag,
            )

    def get_items(self) -> list[dict]:
        """Read current cell values. Returns list of {sku, description, qty}."""
        result = []
        for row_entries in self._rows:
            sku = row_entries[0].get().strip()
            desc = row_entries[1].get().strip()
            qty_str = row_entries[2].get().strip().replace("⚠", "").strip()
            if not sku:
                continue
            try:
                qty = max(1, int(float(qty_str)))
            except ValueError:
                qty = 1
            result.append({"sku": sku, "description": desc, "qty": qty})
        return result

    def clear(self):
        for row_entries in self._rows:
            for entry in row_entries:
                entry.destroy()
        self._rows = []

    def row_count(self) -> int:
        return len(self._rows)

    def _add_row(self, row_idx: int, sku: str, desc: str, qty: str):
        entries = []
        for col, (val, width) in enumerate(zip([sku, desc, qty], self.COL_WIDTHS)):
            entry = ctk.CTkEntry(self, width=width)
            entry.insert(0, val)
            entry.grid(row=row_idx, column=col, padx=(4, 8), pady=2, sticky="ew" if col == 1 else "w")
            entries.append(entry)
        self._rows.append(entries)


class InvoiceTab(ctk.CTkFrame):
    def __init__(self, master, app, on_complete, **kwargs):
        super().__init__(master, fg_color="transparent", **kwargs)
        self._app = app
        self._on_complete = on_complete
        self._build_ui()

    def _build_ui(self):
        # ── Top controls row ──────────────────────────────────────────────
        controls = ctk.CTkFrame(self, fg_color="transparent")
        controls.pack(fill="x", padx=12, pady=(12, 6))

        ctk.CTkLabel(controls, text="Supplier:", font=ctk.CTkFont(size=13)).pack(side="left")

        supplier_names = self._app.config.supplier_names()
        self._supplier_var = ctk.StringVar(value=supplier_names[0] if supplier_names else "")
        self._supplier_menu = ctk.CTkOptionMenu(
            controls,
            values=supplier_names,
            variable=self._supplier_var,
            width=220,
        )
        self._supplier_menu.pack(side="left", padx=(6, 16))

        self._import_btn = ctk.CTkButton(
            controls,
            text="Import PDF",
            width=120,
            command=self._import_pdf,
        )
        self._import_btn.pack(side="left")

        self._clear_btn = ctk.CTkButton(
            controls,
            text="Clear All",
            width=90,
            fg_color="gray50",
            hover_color="gray40",
            command=self._clear,
        )
        self._clear_btn.pack(side="left", padx=(8, 0))

        self._status_label = ctk.CTkLabel(
            controls,
            text="No items loaded.",
            text_color="gray60",
            font=ctk.CTkFont(size=12),
        )
        self._status_label.pack(side="left", padx=(16, 0))

        # ── Editable table ────────────────────────────────────────────────
        self._table = EditableTable(self, corner_radius=6)
        self._table.pack(fill="both", expand=True, padx=12, pady=4)

        # ── Bottom row: error + next button ───────────────────────────────
        bottom = ctk.CTkFrame(self, fg_color="transparent")
        bottom.pack(fill="x", padx=12, pady=(4, 12))

        self._error_label = ctk.CTkLabel(
            bottom,
            text="",
            text_color="red",
            font=ctk.CTkFont(size=12),
            wraplength=700,
            justify="left",
        )
        self._error_label.pack(side="left", fill="x", expand=True)

        self._next_btn = ctk.CTkButton(
            bottom,
            text="Next: Fetch Orders →",
            width=180,
            state="disabled",
            command=self._on_complete,
        )
        self._next_btn.pack(side="right")

    # ── Actions ───────────────────────────────────────────────────────────

    def _import_pdf(self):
        path = filedialog.askopenfilename(
            title="Select Supplier Invoice PDF",
            filetypes=[("PDF files", "*.pdf"), ("All files", "*.*")],
        )
        if not path:
            return

        supplier_name = self._supplier_var.get()
        supplier = self._app.config.get_supplier_by_name(supplier_name)
        if supplier is None:
            self._set_error(f"Unknown supplier: '{supplier_name}'. Check config.json.")
            return

        self._import_btn.configure(state="disabled")
        self._set_status("Parsing PDF…", color="gray60")
        self._set_error("")

        def _worker():
            try:
                items = parse_invoice(path, supplier)
                self.after(0, lambda: self._on_parse_success(items))
            except ParseError as exc:
                self.after(0, lambda e=str(exc): self._on_parse_error(e))
            except Exception as exc:
                self.after(0, lambda e=str(exc): self._on_parse_error(f"Unexpected error: {e}"))

        threading.Thread(target=_worker, daemon=True).start()

    def _on_parse_success(self, items: list[InvoiceItem]):
        self._table.load_items(items, append=True)
        count = self._table.row_count()
        self._set_status(f"{count} item{'s' if count != 1 else ''} loaded.", color="green")
        self._next_btn.configure(state="normal")
        self._import_btn.configure(state="normal")

    def _on_parse_error(self, message: str):
        self._set_error(message)
        self._set_status("Parse failed.", color="orange")
        self._import_btn.configure(state="normal")

    def _clear(self):
        self._table.clear()
        self._set_status("No items loaded.", color="gray60")
        self._set_error("")
        self._next_btn.configure(state="disabled")

    def _set_status(self, text: str, color: str = "gray60"):
        self._status_label.configure(text=text, text_color=color)

    def _set_error(self, text: str):
        self._error_label.configure(text=text)

    # ── Public API (used by results_tab) ──────────────────────────────────

    def get_invoice_items(self):
        """Convert editable table rows back to InvoiceItem objects."""
        from src.pdf_parser import InvoiceItem as _Item
        return [
            _Item(
                sku=d["sku"],
                sku_with_suffix=d["sku"],
                description=d["description"],
                quantity=d["qty"],
                source_page=0,
            )
            for d in self._table.get_items()
        ]
