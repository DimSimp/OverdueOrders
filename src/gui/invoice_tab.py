import os
import tempfile
import threading
from collections import OrderedDict
from tkinter import filedialog

import customtkinter as ctk

from src.pdf_parser import InvoiceItem, ParseError, build_neto_sku, parse_invoice


def _image_to_temp_pdf(image_path: str) -> str:
    """
    Convert an image file (JPG, PNG, TIFF, etc.) to a single-page temporary PDF.

    The PDF contains the image as a full-page raster (no embedded text), so the
    standard parse pipeline will naturally fall through to the AI Vision fallback.

    Returns the path to a temporary .pdf file. The caller is responsible for
    deleting it after use.
    """
    try:
        import fitz
    except ImportError:
        raise ParseError(
            "Image import requires PyMuPDF.\n\nInstall it with: pip install pymupdf"
        )

    img_bytes = open(image_path, "rb").read()
    doc = fitz.open()
    page = doc.new_page()
    page.insert_image(page.rect, stream=img_bytes)

    tmp = tempfile.NamedTemporaryFile(suffix=".pdf", delete=False)
    tmp.close()
    doc.save(tmp.name)
    doc.close()
    return tmp.name


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
        self._loaded_filenames: list[str] = []
        self._invoice_items: list[InvoiceItem] = []
        self._build_ui()

    def _build_ui(self):
        # ── Mode switcher ─────────────────────────────────────────────────
        self._mode_var = ctk.StringVar(value="PDF Invoice")
        self._mode_switcher = ctk.CTkSegmentedButton(
            self,
            values=["PDF Invoice", "FTP Inventory"],
            variable=self._mode_var,
            command=self._on_mode_change,
        )
        self._mode_switcher.pack(fill="x", padx=12, pady=(12, 4))

        # ── PDF controls row ─────────────────────────────────────────────
        self._pdf_frame = ctk.CTkFrame(self, fg_color="transparent")
        self._pdf_frame.pack(fill="x", padx=12, pady=(0, 6))

        ctk.CTkLabel(self._pdf_frame, text="Supplier:", font=ctk.CTkFont(size=13)).pack(side="left")

        supplier_names = self._app.config.supplier_names()
        self._supplier_var = ctk.StringVar(value=supplier_names[0] if supplier_names else "")
        self._supplier_menu = ctk.CTkOptionMenu(
            self._pdf_frame,
            values=supplier_names,
            variable=self._supplier_var,
            width=220,
        )
        self._supplier_menu.pack(side="left", padx=(6, 16))

        self._import_btn = ctk.CTkButton(
            self._pdf_frame,
            text="Import PDF",
            width=120,
            command=self._import_pdf,
        )
        self._import_btn.pack(side="left")

        self._scan_btn = ctk.CTkButton(
            self._pdf_frame,
            text="Scan",
            width=90,
            fg_color=("dodgerblue3", "dodgerblue4"),
            command=self._scan_with_phone,
        )
        self._scan_btn.pack(side="left", padx=(8, 0))

        self._clear_btn = ctk.CTkButton(
            self._pdf_frame,
            text="Clear All",
            width=90,
            fg_color="gray50",
            hover_color="gray40",
            command=self._clear,
        )
        self._clear_btn.pack(side="left", padx=(8, 0))

        self._load_session_btn = ctk.CTkButton(
            self._pdf_frame,
            text="Load Session",
            width=110,
            fg_color=("dodgerblue3", "dodgerblue4"),
            command=self._load_session,
        )
        self._load_session_btn.pack(side="left", padx=(8, 0))

        self._status_label = ctk.CTkLabel(
            self._pdf_frame,
            text="No items loaded.",
            text_color="gray60",
            font=ctk.CTkFont(size=12),
        )
        self._status_label.pack(side="left", padx=(16, 0))

        # ── FTP controls row (initially hidden) ──────────────────────────
        self._ftp_frame = ctk.CTkFrame(self, fg_color="transparent")
        # Not packed yet — only shown when mode switches to FTP

        ctk.CTkLabel(
            self._ftp_frame,
            text="Downloads morning & afternoon inventory from FTP, finds received items.",
            font=ctk.CTkFont(size=12),
            text_color="gray60",
        ).pack(side="left", padx=(0, 12))

        self._ftp_btn = ctk.CTkButton(
            self._ftp_frame,
            text="Load from FTP",
            width=140,
            command=self._load_from_ftp,
        )
        self._ftp_btn.pack(side="left")

        self._ftp_status_label = ctk.CTkLabel(
            self._ftp_frame,
            text="",
            font=ctk.CTkFont(size=12),
            text_color="gray60",
        )
        self._ftp_status_label.pack(side="left", padx=(12, 0))

        # ── Loaded invoices list ──────────────────────────────────────────
        self._files_row = ctk.CTkFrame(self, fg_color="transparent")
        self._files_row.pack(fill="x", padx=12, pady=(0, 2))
        ctk.CTkLabel(
            self._files_row, text="Loaded invoices:",
            font=ctk.CTkFont(size=12), text_color="gray60",
        ).pack(side="left")
        self._files_box = ctk.CTkTextbox(
            self._files_row, height=52, state="disabled",
            font=ctk.CTkFont(size=12), wrap="none",
            border_width=1, corner_radius=4,
        )
        self._files_box.pack(side="left", fill="x", expand=True, padx=(8, 0))

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
            command=self._on_next_clicked,
        )
        self._next_btn.pack(side="right")

    # ── Actions ───────────────────────────────────────────────────────────

    _IMAGE_EXTENSIONS = {".jpg", ".jpeg", ".png", ".tiff", ".tif", ".bmp", ".gif", ".webp"}

    def _import_pdf(self):
        paths = filedialog.askopenfilenames(
            title="Select Supplier Invoice PDF(s) or Images",
            filetypes=[
                ("Supported files", "*.pdf *.jpg *.jpeg *.png *.tiff *.tif *.bmp *.gif *.webp"),
                ("PDF files", "*.pdf"),
                ("Image files", "*.jpg *.jpeg *.png *.tiff *.tif *.bmp *.gif *.webp"),
                ("All files", "*.*"),
            ],
        )
        if not paths:
            return

        supplier_name = self._supplier_var.get()
        supplier = self._app.config.get_supplier_by_name(supplier_name)
        if supplier is None:
            self._set_error(f"Unknown supplier: '{supplier_name}'. Check config.json.")
            return

        self._import_btn.configure(state="disabled")
        self._set_status(f"Parsing {len(paths)} file(s)…", color="gray60")
        self._set_error("")

        openai_config = getattr(self._app.config, "openai", None)

        def _notify_ai_fallback(filename: str):
            self.after(0, lambda: self._set_status(f"AI parsing {filename}…", color="gray60"))

        def _worker():
            errors = []
            temp_files: list[str] = []
            try:
                for idx, path in enumerate(paths):
                    fname = os.path.basename(path)
                    self.after(0, lambda f=fname, i=idx, n=len(paths):
                        self._set_status(f"Parsing {f} ({i+1}/{n})…", color="gray60"))
                    try:
                        # Convert images to a temporary PDF so the standard parse
                        # pipeline (including AI fallback) can handle them.
                        parse_path = path
                        if os.path.splitext(path)[1].lower() in self._IMAGE_EXTENSIONS:
                            self.after(0, lambda f=fname:
                                self._set_status(f"Converting {f} to PDF…", color="gray60"))
                            parse_path = _image_to_temp_pdf(path)
                            temp_files.append(parse_path)

                        items = parse_invoice(
                            parse_path, supplier,
                            openai_config=openai_config,
                            on_ai_fallback=lambda f=fname: _notify_ai_fallback(f),
                        )
                        self.after(0, lambda it=items, f=fname: self._on_parse_success(it, f))
                    except ParseError as exc:
                        errors.append(f"{fname}: {exc}")
                    except Exception as exc:
                        errors.append(f"{fname}: Unexpected error: {exc}")
            finally:
                for tmp in temp_files:
                    try:
                        os.unlink(tmp)
                    except OSError:
                        pass

            if errors:
                self.after(0, lambda e="\n".join(errors): self._on_parse_error(e))
            else:
                self.after(0, lambda: self._import_btn.configure(state="normal"))

        threading.Thread(target=_worker, daemon=True).start()

    def _on_parse_success(self, items: list[InvoiceItem], filename: str):
        self._invoice_items.extend(items)
        self._table.load_items(items, append=True)
        self._loaded_filenames.append(filename)
        self._update_files_box()
        count = self._table.row_count()
        self._set_status(f"{count} item{'s' if count != 1 else ''} loaded.", color="green")
        self._next_btn.configure(state="normal")
        self._import_btn.configure(state="normal")

    def _update_files_box(self):
        self._files_box.configure(state="normal")
        self._files_box.delete("1.0", "end")
        self._files_box.insert("1.0", "  |  ".join(self._loaded_filenames))
        self._files_box.configure(state="disabled")

    def _on_parse_error(self, message: str):
        self._set_error(message)
        self._set_status("Parse failed.", color="orange")
        self._import_btn.configure(state="normal")

    def _clear(self):
        self._invoice_items.clear()
        self._table.clear()
        self._loaded_filenames.clear()
        self._update_files_box()
        self._set_status("No items loaded.", color="gray60")
        self._set_error("")
        self._next_btn.configure(state="disabled")

    def _set_status(self, text: str, color: str = "gray60"):
        self._status_label.configure(text=text, text_color=color)

    def _set_error(self, text: str):
        self._error_label.configure(text=text)

    def _load_session(self):
        """Load a previously saved session snapshot."""
        initial_dir = self._app.config.app.snapshot_dir or self._app.config.app.output_dir
        path = filedialog.askopenfilename(
            title="Load Session Snapshot",
            initialdir=initial_dir,
            filetypes=[("Session files", "*.scar"), ("JSON files", "*.json"), ("All files", "*.*")],
        )
        if not path:
            return

        try:
            from src.session import load_snapshot
            snapshot = load_snapshot(path)

            # Populate app state
            self._app.neto_orders = snapshot.neto_orders
            self._app.ebay_orders = snapshot.ebay_orders
            self._app.matched_orders = snapshot.matched_orders

            # Populate invoice table
            self._table.clear()
            self._invoice_items = list(snapshot.invoice_items)
            self._table.load_items(snapshot.invoice_items, append=True)
            self._loaded_filenames = [os.path.basename(path)]
            self._update_files_box()
            count = self._table.row_count()
            self._set_status(f"Session loaded: {count} item{'s' if count != 1 else ''}.", color="green")
            self._next_btn.configure(state="normal")

            # Set results tab state and jump to it
            results = self._app.results_tab
            results._neto_orders = snapshot.neto_orders
            results._ebay_orders = snapshot.ebay_orders
            results._matched = snapshot.matched_orders
            results._unmatched_inv = snapshot.unmatched_inv
            results._excluded_order_ids = set(snapshot.excluded_order_ids)
            results._force_matched_order_ids = set(snapshot.force_matched_order_ids)
            results._refresh_tables()

            # Jump to results tab
            self._app.tabview.set("3. Results")

        except Exception as e:
            self._set_error(f"Failed to load session: {e}")

    def _scan_with_phone(self):
        """Open the phone scan dialog, then parse any received images."""
        supplier_name = self._supplier_var.get()
        supplier = self._app.config.get_supplier_by_name(supplier_name)
        if supplier is None:
            self._set_error(f"Unknown supplier: '{supplier_name}'. Check config.json.")
            return

        openai_config = getattr(self._app.config, "openai", None)

        def _on_images_done(image_paths: list[str]):
            if not image_paths:
                return
            self._import_btn.configure(state="disabled")
            self._scan_btn.configure(state="disabled")
            self._set_status(f"Processing {len(image_paths)} scanned image(s)…", color="gray60")
            self._set_error("")

            def _notify_ai_fallback(filename: str):
                self.after(0, lambda: self._set_status(f"AI parsing {filename}…", color="gray60"))

            def _worker():
                errors = []
                temp_pdfs: list[str] = []
                try:
                    for idx, img_path in enumerate(image_paths):
                        fname = f"scan_{idx + 1}.jpg"
                        self.after(0, lambda i=idx, n=len(image_paths):
                            self._set_status(f"Parsing scan {i+1}/{n}…", color="gray60"))
                        try:
                            pdf_path = _image_to_temp_pdf(img_path)
                            temp_pdfs.append(pdf_path)
                            items = parse_invoice(
                                pdf_path, supplier,
                                openai_config=openai_config,
                                on_ai_fallback=lambda f=fname: _notify_ai_fallback(f),
                            )
                            self.after(0, lambda it=items, f=fname:
                                self._on_parse_success(it, f))
                        except ParseError as exc:
                            errors.append(f"Scan {idx + 1}: {exc}")
                        except Exception as exc:
                            errors.append(f"Scan {idx + 1}: Unexpected error: {exc}")
                        finally:
                            try:
                                os.unlink(img_path)
                            except OSError:
                                pass
                finally:
                    for tmp in temp_pdfs:
                        try:
                            os.unlink(tmp)
                        except OSError:
                            pass

                if errors:
                    self.after(0, lambda e="\n".join(errors): self._on_parse_error(e))
                else:
                    self.after(0, lambda: (
                        self._import_btn.configure(state="normal"),
                        self._scan_btn.configure(state="normal"),
                    ))

            threading.Thread(target=_worker, daemon=True).start()

        from src.gui.phone_scan_dialog import PhoneScanDialog
        PhoneScanDialog(self, on_done=_on_images_done, supplier_name=supplier_name)

    # ── FTP mode ──────────────────────────────────────────────────────────

    def _on_mode_change(self, mode: str):
        if mode == "PDF Invoice":
            self._ftp_frame.pack_forget()
            self._pdf_frame.pack(fill="x", padx=12, pady=(0, 6),
                                 before=self._files_row)
        else:
            self._pdf_frame.pack_forget()
            self._ftp_frame.pack(fill="x", padx=12, pady=(0, 6),
                                 before=self._files_row)
        self._clear()

    def _load_from_ftp(self):
        ftp_cfg = self._app.config.ftp
        if ftp_cfg is None:
            self._set_error("FTP not configured in config.json.")
            return
        self._ftp_btn.configure(state="disabled")
        self._ftp_status_label.configure(text="Connecting to FTP...", text_color="gray60")
        self._set_error("")

        def _worker():
            from src.ftp_inventory import download_and_compare
            try:
                received = download_and_compare(
                    ftp_cfg.host, ftp_cfg.username, ftp_cfg.password,
                    ftp_cfg.morning_filename, ftp_cfg.afternoon_filename,
                )
                self.after(0, lambda: self._on_ftp_success(received))
            except Exception as exc:
                err_msg = str(exc)
                self.after(0, lambda m=err_msg: self._on_ftp_error(m))

        threading.Thread(target=_worker, daemon=True).start()

    def _on_ftp_success(self, received):
        from src.pdf_parser import InvoiceItem
        items = [
            InvoiceItem(
                sku=r.sku,
                sku_with_suffix=r.sku,
                description="",
                quantity=max(1, int(r.quantity)),
                source_page=0,
                supplier_name=r.supplier,
            )
            for r in received
            if r.sku
        ]
        self._invoice_items = items
        self._table.clear()
        self._table.load_items(items, append=False)
        self._loaded_filenames = ["FTP Inventory (Morning vs Afternoon)"]
        self._update_files_box()
        count = self._table.row_count()
        self._ftp_status_label.configure(
            text=f"{count} received item{'s' if count != 1 else ''} found.",
            text_color="green",
        )
        self._next_btn.configure(state="normal" if count > 0 else "disabled")
        self._ftp_btn.configure(state="normal")

    def _on_ftp_error(self, message: str):
        self._set_error(f"FTP error: {message}")
        self._ftp_status_label.configure(text="Load failed.", text_color="orange")
        self._ftp_btn.configure(state="normal")

    # ── Public API (used by results_tab) ──────────────────────────────────

    def get_invoice_items(self) -> list[InvoiceItem]:
        """
        Merge editable table values (description, qty) with stored item metadata
        (sku, sku_with_suffix, supplier_name).
        """
        table_rows = self._table.get_items()
        if not self._invoice_items:
            # Fallback: table only (e.g. session loaded before this feature)
            return [
                InvoiceItem(
                    sku=r["sku"],
                    sku_with_suffix=r["sku"],
                    description=r["description"],
                    quantity=r["qty"],
                    source_page=0,
                )
                for r in table_rows
            ]
        return [
            InvoiceItem(
                sku=item.sku,
                sku_with_suffix=item.sku_with_suffix,
                description=row["description"],
                quantity=row["qty"],
                source_page=item.source_page,
                qty_flagged=item.qty_flagged,
                supplier_name=item.supplier_name,
            )
            for item, row in zip(self._invoice_items, table_rows)
        ]

    # ── SKU validation flow ────────────────────────────────────────────────

    def _on_next_clicked(self):
        """Validate SKUs against inventory.CSV, show correction dialogs, then proceed."""
        items = self.get_invoice_items()
        if not items:
            return

        # FTP items already use Neto-native SKUs — skip validation
        if self._mode_var.get() == "FTP Inventory":
            self._on_complete()
            return

        app_cfg = self._app.config.app
        self._set_status("Validating SKUs…", color="gray60")
        self._next_btn.configure(state="disabled")

        def _worker():
            from src.sku_validator import load_corrections, load_inventory, validate_items
            inventory = load_inventory(app_cfg.inventory_csv)
            corrections = load_corrections(app_cfg.sku_corrections_file)
            results = validate_items(items, inventory, corrections)
            self.after(0, lambda: self._on_validation_done(results, app_cfg.sku_corrections_file))

        threading.Thread(target=_worker, daemon=True).start()

    def _on_validation_done(self, results, corrections_path: str):
        unconfirmed = [r for r in results if not r.is_confirmed]
        if not unconfirmed:
            self._set_status(f"{len(results)} item(s) validated.", color="green")
            self._apply_corrections_and_proceed(results, corrections_path)
            return

        # Group by supplier name, preserving encounter order
        groups: OrderedDict[str, list] = OrderedDict()
        for r in unconfirmed:
            groups.setdefault(r.item.supplier_name, []).append(r)

        self._set_status(
            f"{len(unconfirmed)} unrecognised SKU(s) — review required.",
            color="orange",
        )
        self._process_supplier_groups(list(groups.items()), results, corrections_path)

    def _process_supplier_groups(self, remaining_groups, all_results, corrections_path: str):
        """Show one SkuCorrectionDialog per supplier group, sequentially."""
        if not remaining_groups:
            self._apply_corrections_and_proceed(all_results, corrections_path)
            return

        supplier_name, group = remaining_groups[0]
        rest = remaining_groups[1:]

        from src.gui.sku_correction_dialog import SkuCorrectionDialog

        def on_accepted(new_corrections):
            if new_corrections:
                from src.sku_validator import save_corrections
                save_corrections(corrections_path, new_corrections)
            self._process_supplier_groups(rest, all_results, corrections_path)

        def on_skip():
            self._process_supplier_groups(rest, all_results, corrections_path)

        def on_cancel():
            self._set_status("Validation cancelled.", color="orange")
            self._next_btn.configure(state="normal")

        SkuCorrectionDialog(
            self,
            unconfirmed=group,
            supplier_name=supplier_name,
            on_accepted=on_accepted,
            on_skip=on_skip,
            on_cancel=on_cancel,
        )

    def _apply_corrections_and_proceed(self, all_results, corrections_path: str):
        """Apply all saved corrections to stored items, then move to orders tab."""
        from src.sku_validator import load_corrections
        corrections = load_corrections(corrections_path)

        for item in self._invoice_items:
            key = (item.supplier_name.lower(), item.sku.upper())
            if key in corrections:
                corrected = corrections[key]
                supplier = self._app.config.get_supplier_by_name(item.supplier_name)
                item.sku = corrected
                item.sku_with_suffix = (
                    build_neto_sku(corrected, supplier)
                    if supplier is not None
                    else corrected
                )

        self._next_btn.configure(state="normal")
        self._on_complete()
