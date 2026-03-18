from __future__ import annotations

import os

import customtkinter as ctk

from src.version import __version__

_ROOT = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
_APP_ICON = os.path.join(_ROOT, "AIO.ico")


class HomeFrame(ctk.CTkFrame):
    """
    Home screen embedded in the App window.
    Shows mode-selection buttons before the user enters a workflow.
    """

    def __init__(self, master, on_afternoon, on_daily, **kwargs):
        super().__init__(master, fg_color="transparent", **kwargs)
        self._on_afternoon = on_afternoon
        self._on_daily = on_daily
        self._build_ui()

    def _build_ui(self):
        # ── Title banner ─────────────────────────────────────────────────
        banner = ctk.CTkFrame(self, height=70, corner_radius=0, fg_color=("gray85", "gray20"))
        banner.pack(fill="x", side="top")
        banner.pack_propagate(False)

        ctk.CTkLabel(
            banner,
            text="Scarlett Music  —  Overdue Orders",
            font=ctk.CTkFont(size=18, weight="bold"),
        ).pack(side="left", padx=24, pady=10)

        ctk.CTkLabel(
            banner,
            text=f"v{__version__}",
            font=ctk.CTkFont(size=11),
            text_color=("gray50", "gray60"),
        ).pack(side="right", padx=20)

        # ── Button area ───────────────────────────────────────────────────
        center = ctk.CTkFrame(self, fg_color="transparent")
        center.pack(expand=True, fill="both", padx=50, pady=30)

        ctk.CTkLabel(
            center,
            text="Select a workflow to begin:",
            font=ctk.CTkFont(size=13),
            text_color=("gray40", "gray70"),
        ).pack(pady=(0, 20))

        # Daily Operations button
        daily_btn = ctk.CTkButton(
            center,
            text="Daily Operations",
            font=ctk.CTkFont(size=16, weight="bold"),
            height=64,
            command=self._on_daily,
        )
        daily_btn.pack(fill="x", pady=(0, 6))
        ctk.CTkLabel(
            center,
            text="Morning workflow  —  picking list, envelopes, order management",
            font=ctk.CTkFont(size=11),
            text_color=("gray50", "gray60"),
        ).pack(pady=(0, 16))

        # Afternoon Operations button
        aft_btn = ctk.CTkButton(
            center,
            text="Afternoon Operations",
            font=ctk.CTkFont(size=16, weight="bold"),
            height=64,
            fg_color=("gray70", "gray30"),
            hover_color=("gray60", "gray25"),
            command=self._on_afternoon,
        )
        aft_btn.pack(fill="x", pady=(0, 6))
        ctk.CTkLabel(
            center,
            text="Match received inventory to overdue orders, book freight",
            font=ctk.CTkFont(size=11),
            text_color=("gray50", "gray60"),
        ).pack()
