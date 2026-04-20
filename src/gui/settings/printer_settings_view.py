from __future__ import annotations

import threading

import customtkinter as ctk

from src.config import config


_NONE_LABEL = "(none)"


class PrinterSettingsView(ctk.CTkFrame):
    """Printer settings panel — select receipt and A4 printers per device."""

    def __init__(self, master, **kwargs) -> None:
        super().__init__(master, fg_color="transparent", **kwargs)
        self._printers: list[str] = []
        self._receipt_var = ctk.StringVar(value=config.device.receipt_printer or _NONE_LABEL)
        self._a4_var = ctk.StringVar(value=config.device.a4_printer or _NONE_LABEL)
        self._envelope_var = ctk.StringVar(value=config.device.envelope_printer or _NONE_LABEL)
        self._build_ui()
        self._load_printers()

    # ── Build ──────────────────────────────────────────────────────────────

    def _build_ui(self) -> None:
        ctk.CTkLabel(
            self,
            text="Printers",
            font=ctk.CTkFont(size=16, weight="bold"),
        ).pack(anchor="w", pady=(0, 4))

        ctk.CTkLabel(
            self,
            text="These settings are saved to this device only.",
            font=ctk.CTkFont(size=11),
            text_color=("gray50", "gray60"),
        ).pack(anchor="w", pady=(0, 20))

        # ── Receipt printer ───────────────────────────────────────────────
        ctk.CTkLabel(
            self,
            text="Receipt printer  (thermal)",
            font=ctk.CTkFont(size=12, weight="bold"),
        ).pack(anchor="w")
        ctk.CTkLabel(
            self,
            text="Used for standard in-store receipts.",
            font=ctk.CTkFont(size=11),
            text_color=("gray50", "gray60"),
        ).pack(anchor="w", pady=(0, 4))

        self._receipt_menu = ctk.CTkOptionMenu(
            self,
            variable=self._receipt_var,
            values=[_NONE_LABEL],
            width=320,
            font=ctk.CTkFont(size=12),
        )
        self._receipt_menu.pack(anchor="w", pady=(0, 20))

        # ── A4 printer ────────────────────────────────────────────────────
        ctk.CTkLabel(
            self,
            text="A4 printer",
            font=ctk.CTkFont(size=12, weight="bold"),
        ).pack(anchor="w")
        ctk.CTkLabel(
            self,
            text="Used for formal A4 invoices when requested by a customer.",
            font=ctk.CTkFont(size=11),
            text_color=("gray50", "gray60"),
        ).pack(anchor="w", pady=(0, 4))

        self._a4_menu = ctk.CTkOptionMenu(
            self,
            variable=self._a4_var,
            values=[_NONE_LABEL],
            width=320,
            font=ctk.CTkFont(size=12),
        )
        self._a4_menu.pack(anchor="w", pady=(0, 20))

        # ── Envelope printer ──────────────────────────────────────────────
        ctk.CTkLabel(
            self,
            text="Envelope printer  (A5)",
            font=ctk.CTkFont(size=12, weight="bold"),
        ).pack(anchor="w")
        ctk.CTkLabel(
            self,
            text="Used for printing Minilope and Devilope address labels.",
            font=ctk.CTkFont(size=11),
            text_color=("gray50", "gray60"),
        ).pack(anchor="w", pady=(0, 4))

        self._envelope_menu = ctk.CTkOptionMenu(
            self,
            variable=self._envelope_var,
            values=[_NONE_LABEL],
            width=320,
            font=ctk.CTkFont(size=12),
        )
        self._envelope_menu.pack(anchor="w", pady=(0, 24))

        # ── Buttons ───────────────────────────────────────────────────────
        btn_row = ctk.CTkFrame(self, fg_color="transparent")
        btn_row.pack(anchor="w")

        ctk.CTkButton(
            btn_row,
            text="Save",
            width=100,
            command=self._save,
        ).pack(side="left", padx=(0, 8))

        ctk.CTkButton(
            btn_row,
            text="Refresh list",
            width=110,
            fg_color="transparent",
            border_width=1,
            border_color=("gray60", "gray45"),
            text_color=("gray20", "gray90"),
            hover_color=("gray85", "gray25"),
            command=self._load_printers,
        ).pack(side="left")

        self._status_lbl = ctk.CTkLabel(
            self,
            text="",
            font=ctk.CTkFont(size=11),
            text_color=("gray50", "gray60"),
        )
        self._status_lbl.pack(anchor="w", pady=(10, 0))

    # ── Printer discovery ─────────────────────────────────────────────────

    def _load_printers(self) -> None:
        self._status_lbl.configure(text="Loading printers...", text_color=("gray50", "gray60"))
        threading.Thread(target=self._fetch_printers, daemon=True).start()

    def _fetch_printers(self) -> None:
        from src.printer_utils import list_printers
        printers = list_printers()
        self.after(0, lambda: self._apply_printers(printers))

    def _apply_printers(self, printers: list[str]) -> None:
        self._printers = printers
        options = [_NONE_LABEL] + printers

        # Re-insert saved values that aren't in the discovered list, then
        # explicitly restore the StringVar — CTkOptionMenu.configure(values=…)
        # can reset the variable to the first item in some customtkinter builds.
        for var, menu in ((self._receipt_var, self._receipt_menu),
                          (self._a4_var, self._a4_menu),
                          (self._envelope_var, self._envelope_menu)):
            # Strip any stale "(not found)" suffix carried from a previous scan
            raw = var.get()
            if raw.endswith(" (not found)"):
                raw = raw[: -len(" (not found)")]

            if raw and raw != _NONE_LABEL and raw not in printers:
                opts = [_NONE_LABEL, f"{raw} (not found)"] + printers
                menu.configure(values=opts)
                var.set(f"{raw} (not found)")
            else:
                menu.configure(values=options)
                var.set(raw if raw in printers else _NONE_LABEL)

        count = len(printers)
        self._status_lbl.configure(
            text=f"{count} printer{'s' if count != 1 else ''} found.",
            text_color=("gray50", "gray60"),
        )

    # ── Save ──────────────────────────────────────────────────────────────

    def _save(self) -> None:
        receipt = self._receipt_var.get()
        a4 = self._a4_var.get()
        envelope = self._envelope_var.get()

        # Strip "(not found)" suffix if present
        if receipt.endswith(" (not found)"):
            receipt = receipt[: -len(" (not found)")]
        if a4.endswith(" (not found)"):
            a4 = a4[: -len(" (not found)")]
        if envelope.endswith(" (not found)"):
            envelope = envelope[: -len(" (not found)")]

        config.device.receipt_printer = "" if receipt == _NONE_LABEL else receipt
        config.device.a4_printer = "" if a4 == _NONE_LABEL else a4
        config.device.envelope_printer = "" if envelope == _NONE_LABEL else envelope

        try:
            config.save_local()
            self._status_lbl.configure(
                text="Saved.", text_color=("#1a6b2e", "#22c55e"),
            )
            self.after(3000, lambda: self._status_lbl.configure(
                text=f"{len(self._printers)} printer{'s' if len(self._printers) != 1 else ''} found.",
                text_color=("gray50", "gray60"),
            ))
        except Exception as exc:
            self._status_lbl.configure(
                text=f"Save failed: {exc}",
                text_color=("#b91c1c", "#f87171"),
            )
