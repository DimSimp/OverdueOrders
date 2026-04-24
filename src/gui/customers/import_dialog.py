from __future__ import annotations

import logging
import threading
from tkinter import filedialog, ttk
from typing import Optional

import customtkinter as ctk

log = logging.getLogger(__name__)


class CustomerImportDialog(ctk.CTkToplevel):
    """Import customers from a Musipos CSV export (no header row)."""

    def __init__(self, master, on_done: Optional[callable] = None):
        super().__init__(master)
        self._on_done = on_done
        self._csv_path: Optional[str] = None
        self._stop_event = threading.Event()
        self._import_thread: Optional[threading.Thread] = None

        self.title("Import Customers — Musipos CSV")
        self.geometry("620x520")
        self.minsize(520, 420)
        self.resizable(True, True)
        self.transient(master)
        self.grab_set()

        self._build_ui()

    # ── Build ────────────────────────────────────────────────────────────

    def _build_ui(self):
        body = ctk.CTkFrame(self, fg_color="transparent")
        body.pack(fill="both", expand=True, padx=16, pady=16)

        # File picker row
        file_row = ctk.CTkFrame(body, fg_color="transparent")
        file_row.pack(fill="x", pady=(0, 8))

        ctk.CTkButton(
            file_row, text="Browse…", width=90, height=30,
            font=ctk.CTkFont(size=12),
            command=self._browse,
        ).pack(side="left")

        self._lbl_path = ctk.CTkLabel(
            file_row, text="No file selected",
            font=ctk.CTkFont(size=11),
            text_color=("gray50", "gray60"),
            anchor="w",
        )
        self._lbl_path.pack(side="left", padx=(10, 0), fill="x", expand=True)

        # Warning note
        ctk.CTkLabel(
            body,
            text=(
                "Note: Mobile numbers are not included in the Musipos customer export.\n"
                "They will need to be added to each customer profile manually after import."
            ),
            font=ctk.CTkFont(size=11),
            text_color=("orange", "#e67e22"),
            justify="left",
            anchor="w",
        ).pack(fill="x", pady=(0, 10))

        # Preview table
        ctk.CTkLabel(
            body, text="Preview (first 5 rows):",
            font=ctk.CTkFont(size=11, weight="bold"),
            anchor="w",
        ).pack(fill="x", pady=(0, 4))

        preview_frame = ctk.CTkFrame(body, fg_color="transparent")
        preview_frame.pack(fill="x", pady=(0, 12))

        prev_cols = ("musipos_id", "first_name", "surname", "business", "email", "phone_1")
        self._preview_tree = ttk.Treeview(
            preview_frame,
            columns=prev_cols,
            show="headings",
            style="Prev.Treeview",
            height=5,
        )
        self._preview_tree.heading("musipos_id",  text="Musipos ID", anchor="w")
        self._preview_tree.heading("first_name",  text="First Name", anchor="w")
        self._preview_tree.heading("surname",     text="Surname",    anchor="w")
        self._preview_tree.heading("business",    text="Business",   anchor="w")
        self._preview_tree.heading("email",       text="Email",      anchor="w")
        self._preview_tree.heading("phone_1",     text="Phone",      anchor="w")

        self._preview_tree.column("musipos_id", width=90,  stretch=False)
        self._preview_tree.column("first_name", width=100, stretch=False)
        self._preview_tree.column("surname",    width=100, stretch=False)
        self._preview_tree.column("business",   width=130, stretch=True)
        self._preview_tree.column("email",      width=140, stretch=True)
        self._preview_tree.column("phone_1",    width=90,  stretch=False)

        self._preview_tree.pack(fill="x")

        self._lbl_count = ctk.CTkLabel(
            body, text="", font=ctk.CTkFont(size=11),
            text_color=("gray50", "gray60"), anchor="w",
        )
        self._lbl_count.pack(fill="x", pady=(2, 0))

        # Progress bar (hidden until import starts)
        self._progress = ctk.CTkProgressBar(body, mode="determinate")
        self._progress.set(0)

        # Status label
        self._lbl_status = ctk.CTkLabel(
            body, text="", font=ctk.CTkFont(size=11), anchor="w",
        )
        self._lbl_status.pack(fill="x")

        # Footer
        footer = ctk.CTkFrame(self, fg_color=("gray85", "gray20"), height=52)
        footer.pack(fill="x", side="bottom")
        footer.pack_propagate(False)

        ctk.CTkButton(
            footer, text="Close", width=80, height=32,
            fg_color="transparent", border_width=1,
            border_color=("gray60", "gray45"),
            text_color=("gray30", "gray70"),
            hover_color=("gray85", "gray25"),
            command=self._on_close,
        ).pack(side="right", padx=(0, 12), pady=10)

        self._btn_import = ctk.CTkButton(
            footer, text="Start Import", width=120, height=32,
            font=ctk.CTkFont(size=12, weight="bold"),
            state="disabled",
            command=self._start_import,
        )
        self._btn_import.pack(side="right", padx=(0, 8), pady=10)

    # ── Browse / Preview ─────────────────────────────────────────────────

    def _browse(self):
        path = filedialog.askopenfilename(
            parent=self,
            title="Select Musipos Customer Export CSV",
            filetypes=[("CSV files", "*.csv"), ("All files", "*.*")],
        )
        if not path:
            return
        self._csv_path = path
        self._lbl_path.configure(text=path, text_color=("gray20", "gray80"))
        self._btn_import.configure(state="disabled")
        self._lbl_count.configure(text="Loading preview…")
        self._preview_tree.delete(*self._preview_tree.get_children())

        threading.Thread(target=self._load_preview, daemon=True).start()

    def _load_preview(self):
        from src.customers.importer import get_preview, count_rows
        try:
            rows = get_preview(self._csv_path)
            total = count_rows(self._csv_path)
            self.after(0, lambda: self._show_preview(rows, total))
        except Exception as exc:
            log.exception("Customer import — failed to read CSV preview")
            err = str(exc)
            self.after(0, lambda: self._lbl_count.configure(
                text=f"Error reading file: {err}",
                text_color=("red", "#e74c3c"),
            ))

    def _show_preview(self, rows: list, total: int):
        self._preview_tree.delete(*self._preview_tree.get_children())
        for r in rows:
            self._preview_tree.insert("", "end", values=(
                r.get("musipos_account_code") or "",
                r.get("first_name") or "",
                r.get("surname") or "",
                r.get("business") or "",
                r.get("email") or "",
                r.get("phone_1") or "",
            ))
        self._lbl_count.configure(
            text=f"{total} customer rows found.",
            text_color=("gray50", "gray60"),
        )
        self._btn_import.configure(state="normal")

    # ── Import ───────────────────────────────────────────────────────────

    def _start_import(self):
        if not self._csv_path:
            return
        self._btn_import.configure(state="disabled", text="Importing…")
        self._progress.pack(fill="x", before=self._lbl_status)
        self._progress.set(0)
        self._lbl_status.configure(
            text="Reading CSV…", text_color=("gray20", "gray80")
        )
        self._stop_event.clear()
        self._import_thread = threading.Thread(
            target=self._import_thread_fn, daemon=True
        )
        self._import_thread.start()

    def _import_thread_fn(self):
        from src.customers.importer import load_csv
        from src.customers.customer_client import batch_upsert_customers
        try:
            rows = load_csv(self._csv_path)
            total = len(rows)
            self.after(0, lambda: self._lbl_status.configure(
                text=f"Upserting {total} records…"
            ))

            result = batch_upsert_customers(
                rows,
                on_progress=self._on_progress,
                stop_event=self._stop_event,
            )
            self.after(0, lambda: self._on_done_import(result, total))
        except Exception as exc:
            log.exception("Customer import — unexpected error during import")
            err = str(exc)
            self.after(0, lambda: self._lbl_status.configure(
                text=f"Error: {err}", text_color=("red", "#e74c3c")
            ))
            self.after(0, lambda: self._btn_import.configure(
                state="normal", text="Start Import"
            ))

    def _on_progress(self, done: int, total: int):
        frac = done / total if total else 0
        self.after(0, lambda: self._progress.set(frac))
        self.after(0, lambda: self._lbl_status.configure(
            text=f"Processed {done} of {total}…"
        ))

    def _on_done_import(self, result: dict, total: int):
        self._progress.set(1)
        done = result["done"]
        errors = result["errors"]
        first_error = result.get("first_error")

        msg = f"Imported {done} of {total} customers."
        if errors:
            msg += f"  {errors} error(s)."
        color = ("red", "#e74c3c") if errors else ("green", "#2ecc71")
        self._lbl_status.configure(text=msg, text_color=color)

        if first_error:
            log.error("Customer import — first batch error: %s", first_error)
            ctk.CTkLabel(
                self, text=f"First error: {first_error[:120]}",
                font=ctk.CTkFont(size=10),
                text_color=("red", "#e74c3c"),
                wraplength=560,
                anchor="w",
            ).pack(fill="x", padx=16, pady=(0, 4))

        self._btn_import.configure(state="disabled", text="Done")
        if self._on_done:
            self._on_done()

    # ── Close ─────────────────────────────────────────────────────────────

    def _on_close(self):
        self._stop_event.set()
        self.destroy()
