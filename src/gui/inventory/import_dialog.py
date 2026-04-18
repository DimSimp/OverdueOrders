from __future__ import annotations

import threading
from tkinter import filedialog, ttk

import customtkinter as ctk


class ImportDialog(ctk.CTkToplevel):
    """Modal dialog for importing a Musipos inventory CSV into Supabase.

    Flow:
        1. User browses to the .CSV file → row count + 5-row preview shown
        2. Optional: check "Import stock quantities from CSV"
        3. Click "Start Import" → progress bar + status label update live
        4. On complete: result summary; Close button dismisses
    """

    def __init__(self, master, on_complete=None, **kwargs):
        super().__init__(master, **kwargs)
        self._on_complete = on_complete
        self._csv_path: str = ""
        self._stop_event = threading.Event()
        self._running = False

        self.title("Import Musipos Inventory CSV")
        self.geometry("720x540")
        self.resizable(False, False)
        self.grab_set()
        self._build_ui()

        try:
            self.iconbitmap(master.iconbitmap())
        except Exception:
            pass

    # ── Layout ────────────────────────────────────────────────────────────

    def _build_ui(self):
        # ── File picker ──────────────────────────────────────────────────
        file_row = ctk.CTkFrame(self, fg_color="transparent")
        file_row.pack(fill="x", padx=20, pady=(20, 0))

        ctk.CTkLabel(file_row, text="CSV File:", font=ctk.CTkFont(size=13)).pack(side="left")
        self._path_label = ctk.CTkLabel(
            file_row, text="No file selected",
            font=ctk.CTkFont(size=12), text_color=("gray50", "gray60"), anchor="w",
        )
        self._path_label.pack(side="left", padx=(8, 0), expand=True, fill="x")
        ctk.CTkButton(file_row, text="Browse...", width=90, command=self._browse).pack(side="right")

        self._file_info = ctk.CTkLabel(
            self, text="", font=ctk.CTkFont(size=12), text_color=("gray50", "gray60"),
        )
        self._file_info.pack(anchor="w", padx=20, pady=(4, 0))

        # ── Options ──────────────────────────────────────────────────────
        opts = ctk.CTkFrame(self, fg_color="transparent")
        opts.pack(fill="x", padx=20, pady=(12, 0))
        self._qty_var = ctk.BooleanVar(value=False)
        ctk.CTkCheckBox(
            opts,
            text="Import stock quantities from CSV  (leave unchecked if quantities are stale)",
            variable=self._qty_var,
            font=ctk.CTkFont(size=12),
        ).pack(anchor="w")

        # ── Preview ──────────────────────────────────────────────────────
        ctk.CTkLabel(
            self, text="Preview (first 5 rows):",
            font=ctk.CTkFont(size=12, weight="bold"),
        ).pack(anchor="w", padx=20, pady=(16, 4))

        self._preview_frame = ctk.CTkFrame(self, fg_color=("gray95", "gray15"), corner_radius=6)
        self._preview_frame.pack(fill="x", padx=20)

        self._preview_msg = ctk.CTkLabel(
            self._preview_frame, text="Select a file to see a preview.",
            text_color=("gray50", "gray60"), font=ctk.CTkFont(size=12),
        )
        self._preview_msg.pack(pady=16)

        # ── Progress ─────────────────────────────────────────────────────
        self._progress_bar = ctk.CTkProgressBar(self)
        self._progress_bar.set(0)

        self._status_label = ctk.CTkLabel(
            self, text="", font=ctk.CTkFont(size=12), text_color=("gray50", "gray60"),
        )
        self._status_label.pack(anchor="w", padx=20, pady=(8, 0))

        # ── Buttons ──────────────────────────────────────────────────────
        btn_row = ctk.CTkFrame(self, fg_color="transparent")
        btn_row.pack(fill="x", padx=20, pady=(12, 20), side="bottom")

        self._action_btn = ctk.CTkButton(
            btn_row, text="Cancel", width=90,
            fg_color=("gray75", "gray30"), command=self._on_action,
        )
        self._action_btn.pack(side="right", padx=(6, 0))

        self._import_btn = ctk.CTkButton(
            btn_row, text="Start Import", width=120,
            command=self._start_import, state="disabled",
        )
        self._import_btn.pack(side="right")

    # ── File / Preview ────────────────────────────────────────────────────

    def _browse(self):
        path = filedialog.askopenfilename(
            parent=self,
            title="Select Musipos Inventory CSV",
            filetypes=[("CSV files", "*.csv *.CSV"), ("All files", "*.*")],
        )
        if not path:
            return
        self._csv_path = path
        self._path_label.configure(text=path, text_color=("gray80", "gray80"))
        self._file_info.configure(text="Reading…", text_color=("gray50", "gray60"))
        self._import_btn.configure(state="disabled")
        threading.Thread(target=self._do_preview, args=(path,), daemon=True).start()

    def _do_preview(self, path: str):
        from src.inventory.importer import get_preview, count_rows
        try:
            n = count_rows(path)
            headers, rows = get_preview(path, n=5)
            self.after(0, lambda: self._show_preview(headers, rows, n))
        except Exception as exc:
            self.after(0, lambda: self._file_info.configure(
                text=f"Error reading file: {exc}", text_color="red",
            ))

    def _show_preview(self, headers: list[str], rows: list[list], total: int):
        self._file_info.configure(
            text=f"{total:,} items found in file",
            text_color=("gray50", "gray60"),
        )

        for w in self._preview_frame.winfo_children():
            w.destroy()

        if not rows:
            ctk.CTkLabel(
                self._preview_frame, text="No data rows found.",
                text_color="orange", font=ctk.CTkFont(size=12),
            ).pack(pady=12)
            return

        # Show a readable subset of columns
        show_cols = ["sku", "title", "brand", "supplier_rrp", "last_purchase_cost",
                     "qty_on_hand", "supplier_id", "active"]
        idxs = [headers.index(c) for c in show_cols if c in headers]
        disp_h = [headers[i] for i in idxs]

        tree = ttk.Treeview(
            self._preview_frame, columns=disp_h, show="headings",
            height=5, style="Prev.Treeview",
        )
        col_widths = {"sku": 90, "title": 200, "brand": 80, "supplier_id": 70}
        for h in disp_h:
            tree.heading(h, text=h)
            tree.column(h, width=col_widths.get(h, 70), minwidth=40)
        for row in rows:
            tree.insert("", "end", values=[row[i] for i in idxs])
        tree.pack(fill="x", padx=4, pady=4)

        self._import_btn.configure(state="normal")

    # ── Import ────────────────────────────────────────────────────────────

    def _start_import(self):
        if not self._csv_path or self._running:
            return
        self._running = True
        self._stop_event.clear()
        self._import_btn.configure(state="disabled")
        self._action_btn.configure(text="Stop")
        self._progress_bar.set(0)
        self._progress_bar.pack(fill="x", padx=20, pady=(8, 0))
        self._status_label.configure(text="Starting import…")
        threading.Thread(target=self._do_import, daemon=True).start()

    def _do_import(self):
        from src.inventory.importer import load_csv, get_unique_supplier_ids
        from src.inventory.inventory_client import ensure_suppliers, batch_upsert_items
        try:
            self.after(0, lambda: self._status_label.configure(text="Reading CSV…"))
            rows = load_csv(self._csv_path, import_quantities=self._qty_var.get())
            total = len(rows)

            self.after(0, lambda: self._status_label.configure(text="Checking suppliers…"))
            supplier_ids = get_unique_supplier_ids(self._csv_path)
            ensure_suppliers(supplier_ids)

            def _progress(done, _total):
                frac = done / _total if _total else 1.0
                self.after(0, lambda: self._progress_bar.set(frac))
                self.after(0, lambda: self._status_label.configure(
                    text=f"Importing…  {done:,} / {_total:,}",
                ))

            result = batch_upsert_items(rows, on_progress=_progress,
                                        stop_event=self._stop_event)
            print(f"[IMPORT] Complete — {result['done']:,}/{total:,} rows, "
                  f"{result['errors']:,} batch errors.")
            if result.get("first_error"):
                print(f"[IMPORT] First error: {result['first_error']}")
            self.after(0, lambda: self._finish(result, total))

        except Exception as exc:
            err = str(exc)
            self.after(0, lambda: self._error(err))

    def _finish(self, result: dict, total: int):
        self._running = False
        self._progress_bar.set(1)
        stopped = self._stop_event.is_set()
        done = result["done"]
        errors = result["errors"]
        first_error = result.get("first_error", "")
        msg = f"{'Stopped early.  ' if stopped else ''}Processed {done:,} of {total:,} rows."
        if errors:
            msg += f"  {errors:,} batch error(s)."
            if first_error:
                msg += f"\nFirst error: {first_error}"
        self._status_label.configure(
            text=msg,
            text_color="orange" if errors else ("gray50", "gray60"),
        )
        self._action_btn.configure(text="Close")
        if self._on_complete:
            self._on_complete()

    def _error(self, err: str):
        self._running = False
        self._status_label.configure(text=f"Error: {err}", text_color="red")
        self._action_btn.configure(text="Close")
        self._import_btn.configure(state="normal")

    def _on_action(self):
        if self._running:
            self._stop_event.set()
        else:
            self.destroy()
