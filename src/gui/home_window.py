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

    def __init__(self, master, on_afternoon, on_daily, on_settings=None, on_switch_user=None,
                 current_user=None, on_pos=None, pos_dev_user=None, **kwargs):
        super().__init__(master, fg_color="transparent", **kwargs)
        self._on_afternoon = on_afternoon
        self._on_daily = on_daily
        self._on_settings = on_settings
        self._on_switch_user = on_switch_user
        self._current_user = current_user
        self._on_pos = on_pos
        self._pos_dev_user = pos_dev_user
        self._build_ui()

    def _open_dev_console(self):
        from src.gui.dev_console import open_dev_console
        open_dev_console(self.winfo_toplevel())

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

        ver_label = ctk.CTkLabel(
            banner,
            text=f"v{__version__}",
            font=ctk.CTkFont(size=11),
            text_color=("gray50", "gray60"),
        )
        ver_label.pack(side="right", padx=20)
        ver_label.bind("<Button-3>", lambda e: self._open_dev_console())

        # Logged-in user name + switch button (right of version label)
        if self._current_user:
            first = self._current_user.get("first_name", "")
            last = self._current_user.get("last_name", "")
            display = f"{first} {last}".strip() or self._current_user.get("username", "")
            if self._on_switch_user:
                ctk.CTkButton(
                    banner,
                    text="Switch User",
                    font=ctk.CTkFont(size=11),
                    height=26, width=96,
                    fg_color="transparent",
                    border_width=1,
                    border_color=("gray60", "gray45"),
                    text_color=("gray40", "gray70"),
                    hover_color=("gray85", "gray25"),
                    command=self._on_switch_user,
                ).pack(side="right", padx=(0, 12))
            ctk.CTkLabel(
                banner,
                text=f"👤 {display}",
                font=ctk.CTkFont(size=11),
                text_color=("gray50", "gray60"),
            ).pack(side="right", padx=(0, 4))

        # ── Button area ───────────────────────────────────────────────────
        center = ctk.CTkFrame(self, fg_color="transparent")
        center.pack(expand=True, fill="both", padx=50, pady=(20, 10))

        ctk.CTkLabel(
            center,
            text="Select a workflow to begin:",
            font=ctk.CTkFont(size=13),
            text_color=("gray40", "gray70"),
        ).pack(pady=(0, 20))

        # POS button — only shown for pos_dev_user (or all users when pos_dev_user is None)
        _show_pos = (
            self._current_user is not None
            and self._on_pos is not None
            and (
                self._pos_dev_user is None
                or self._current_user.get("username") == self._pos_dev_user
            )
        )
        if _show_pos:
            ctk.CTkButton(
                center,
                text="POS / Till",
                font=ctk.CTkFont(size=16, weight="bold"),
                height=64,
                command=self._on_pos,
            ).pack(fill="x", pady=(0, 6))
            ctk.CTkLabel(
                center,
                text="In-store sales, quotes, invoices, repairs, deposits",
                font=ctk.CTkFont(size=11),
                text_color=("gray50", "gray60"),
            ).pack(pady=(0, 16))

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
            command=self._on_afternoon,
        )
        aft_btn.pack(fill="x", pady=(0, 6))
        ctk.CTkLabel(
            center,
            text="Match received inventory to overdue orders, book freight",
            font=ctk.CTkFont(size=11),
            text_color=("gray50", "gray60"),
        ).pack()

        # ── Settings button ───────────────────────────────────────────────
        if self._on_settings:
            ctk.CTkButton(
                center,
                text="⚙  Settings",
                font=ctk.CTkFont(size=12),
                height=32,
                fg_color="transparent",
                border_width=1,
                border_color=("gray60", "gray45"),
                text_color=("gray40", "gray70"),
                hover_color=("gray85", "gray25"),
                command=self._on_settings,
            ).pack(pady=(14, 0))
