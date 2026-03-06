"""
Dialog for reviewing and correcting unrecognised SKUs after AI parsing.

Shown once per supplier group (sequentially) when validate_items() returns
unconfirmed items.  The user can correct each SKU via a text entry pre-filled
with the best suggestion, or choose from a drop-down of up to 5 suggestions.
"""

from __future__ import annotations

from typing import Callable, Optional

import customtkinter as ctk

from src.sku_validator import ValidationResult


class SkuCorrectionDialog(ctk.CTkToplevel):
    """
    Modal dialog showing unrecognised SKUs for a single supplier group.

    Callbacks (all called on the main thread after destroy()):
      on_accepted(corrections)   — list of (supplier_name, raw_sku, corrected_sku)
                                   only for rows where corrected != original
      on_skip()                  — user skipped validation for this supplier
      on_cancel()                — user cancelled the entire validation flow
    """

    # Suppress CTkToplevel's deferred iconbitmap call
    def iconbitmap(self, *args, **kwargs):
        try:
            super().iconbitmap(*args, **kwargs)
        except Exception:
            pass

    def __init__(
        self,
        master,
        unconfirmed: list[ValidationResult],
        supplier_name: str,
        on_accepted: Callable[[list[tuple[str, str, str]]], None],
        on_skip: Callable[[], None],
        on_cancel: Callable[[], None],
    ):
        super().__init__(master)
        self._unconfirmed = unconfirmed
        self._supplier_name = supplier_name
        self._on_accepted = on_accepted
        self._on_skip = on_skip
        self._on_cancel = on_cancel

        # Per-row widgets: list of (entry_var, original_sku)
        self._rows: list[tuple[ctk.StringVar, str]] = []

        self.title(f"SKU Corrections — {supplier_name}")
        self.resizable(True, False)
        self.transient(master.winfo_toplevel())
        self.protocol("WM_DELETE_WINDOW", self._cancel)

        self._build_ui()
        self.update_idletasks()
        self._size_window()
        self.after(150, self._activate)

    # ------------------------------------------------------------------
    # UI construction
    # ------------------------------------------------------------------

    def _build_ui(self):
        n = len(self._unconfirmed)

        # ── Heading ──────────────────────────────────────────────────────
        ctk.CTkLabel(
            self,
            text=f"Unrecognised SKUs — {self._supplier_name}",
            font=ctk.CTkFont(size=17, weight="bold"),
        ).pack(padx=24, pady=(20, 4))

        ctk.CTkLabel(
            self,
            text=(
                f"{n} SKU{'s' if n != 1 else ''} could not be found in inventory.\n"
                "Review each SKU below — edit directly or choose a suggestion."
            ),
            font=ctk.CTkFont(size=13),
            text_color="gray60",
            justify="center",
        ).pack(padx=24, pady=(0, 12))

        # ── Column headers ────────────────────────────────────────────────
        hdr = ctk.CTkFrame(self, fg_color="transparent")
        hdr.pack(fill="x", padx=24)
        for col_text, col_w in [
            ("AI SKU", 130),
            ("Description", 240),
            ("Qty", 45),
            ("Corrected SKU", 160),
            ("Suggestions", 175),
        ]:
            ctk.CTkLabel(
                hdr,
                text=col_text,
                font=ctk.CTkFont(size=12, weight="bold"),
                width=col_w,
                anchor="w",
            ).pack(side="left", padx=(0, 6))

        # ── Scrollable item rows ──────────────────────────────────────────
        scroll = ctk.CTkScrollableFrame(self, height=min(n * 46 + 10, 340))
        scroll.pack(fill="x", padx=24, pady=(4, 0))

        for vr in self._unconfirmed:
            self._add_row(scroll, vr)

        # ── Buttons ───────────────────────────────────────────────────────
        btn_row = ctk.CTkFrame(self, fg_color="transparent")
        btn_row.pack(padx=24, pady=(14, 20), fill="x")

        ctk.CTkButton(
            btn_row,
            text="Accept All",
            width=130,
            fg_color=("green3", "green4"),
            hover_color=("green4", "green3"),
            command=self._accept,
        ).pack(side="left", padx=(0, 10))

        ctk.CTkButton(
            btn_row,
            text="Skip Validation",
            width=130,
            fg_color="gray50",
            hover_color="gray40",
            command=self._skip,
        ).pack(side="left", padx=(0, 10))

        ctk.CTkButton(
            btn_row,
            text="Cancel",
            width=90,
            fg_color=("firebrick3", "firebrick4"),
            hover_color=("firebrick4", "firebrick3"),
            command=self._cancel,
        ).pack(side="right")

    def _add_row(self, parent, vr: ValidationResult):
        item = vr.item
        row_frame = ctk.CTkFrame(parent, fg_color="transparent")
        row_frame.pack(fill="x", pady=3)

        # AI SKU (red — unrecognised)
        ctk.CTkLabel(
            row_frame,
            text=item.sku,
            width=130,
            anchor="w",
            font=ctk.CTkFont(size=12),
            text_color=("firebrick3", "tomato"),
        ).pack(side="left", padx=(0, 6))

        # Description (truncated)
        desc = item.description[:40] + "…" if len(item.description) > 40 else item.description
        ctk.CTkLabel(
            row_frame,
            text=desc,
            width=240,
            anchor="w",
            font=ctk.CTkFont(size=12),
            text_color="gray60",
        ).pack(side="left", padx=(0, 6))

        # Qty
        ctk.CTkLabel(
            row_frame,
            text=str(item.quantity),
            width=45,
            anchor="e",
            font=ctk.CTkFont(size=12),
        ).pack(side="left", padx=(0, 6))

        # Corrected SKU entry — pre-fill with best suggestion or original
        best = vr.suggestions[0] if vr.suggestions else item.sku.upper()
        var = ctk.StringVar(value=best)
        entry = ctk.CTkEntry(
            row_frame,
            textvariable=var,
            width=160,
            font=ctk.CTkFont(size=12),
        )
        entry.pack(side="left", padx=(0, 6))

        # Suggestions drop-down
        if vr.suggestions:
            options = vr.suggestions[:5]

            def _on_select(choice: str, v: ctk.StringVar = var) -> None:
                v.set(choice)

            ctk.CTkOptionMenu(
                row_frame,
                values=options,
                width=175,
                font=ctk.CTkFont(size=12),
                command=_on_select,
            ).pack(side="left")
        else:
            ctk.CTkLabel(
                row_frame,
                text="No suggestions",
                width=175,
                anchor="w",
                font=ctk.CTkFont(size=12),
                text_color="gray50",
            ).pack(side="left")

        self._rows.append((var, item.sku.upper()))

    # ------------------------------------------------------------------
    # Window sizing
    # ------------------------------------------------------------------

    def _size_window(self):
        self.update_idletasks()
        w = max(self.winfo_reqwidth(), 820)
        h = self.winfo_reqheight()
        self.geometry(f"{w}x{h}")

    # ------------------------------------------------------------------
    # Actions
    # ------------------------------------------------------------------

    def _accept(self):
        corrections: list[tuple[str, str, str]] = []
        for (var, original_upper), vr in zip(self._rows, self._unconfirmed):
            corrected = var.get().strip().upper()
            if corrected and corrected != original_upper:
                corrections.append((self._supplier_name, original_upper, corrected))
        self.destroy()
        self._on_accepted(corrections)

    def _skip(self):
        self.destroy()
        self._on_skip()

    def _cancel(self):
        self.destroy()
        self._on_cancel()

    def _activate(self):
        self.lift()
        self.focus_force()
        try:
            self.grab_set()
        except Exception:
            pass
