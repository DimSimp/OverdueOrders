"""Settings window — CTkToplevel with sidebar navigation."""
from __future__ import annotations

import os
from typing import TYPE_CHECKING

import customtkinter as ctk

if TYPE_CHECKING:
    from src.user_manager import UserManager

_ROOT = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
_APP_ICON = os.path.join(_ROOT, "AIO.ico")


class SettingsWindow(ctk.CTkToplevel):
    """Settings window with a sidebar and swappable content area.

    Currently has one section: Users.  More sections can be added to
    ``_NAV_ITEMS`` and the corresponding ``_build_*_view()`` methods.
    """

    _NAV_ITEMS = [
        ("Users", "_build_users_view"),
    ]

    def __init__(self, master, user_manager: "UserManager", current_user: dict, **kwargs) -> None:
        super().__init__(master, **kwargs)
        self.user_manager = user_manager
        self.current_user = current_user

        self.title("Settings")
        self.geometry("720x520")
        self.minsize(600, 420)
        self.resizable(True, True)
        if os.path.exists(_APP_ICON):
            try:
                self.iconbitmap(_APP_ICON)
            except Exception:
                pass

        self._content_frame: ctk.CTkFrame | None = None
        self._nav_buttons: dict[str, ctk.CTkButton] = {}
        self._build_ui()
        self._select_section("Users")
        self.after(50, self.lift)

    # ── Layout ────────────────────────────────────────────────────────────────

    def _build_ui(self) -> None:
        # Header
        header = ctk.CTkFrame(self, height=48, corner_radius=0, fg_color=("gray85", "gray20"))
        header.pack(fill="x", side="top")
        header.pack_propagate(False)
        ctk.CTkLabel(
            header,
            text="⚙  Settings",
            font=ctk.CTkFont(size=15, weight="bold"),
        ).pack(side="left", padx=20, pady=10)
        if self.current_user:
            _u = self.current_user
            _display = (f"{_u.get('first_name', '')} {_u.get('last_name', '')}".strip()
                        or _u.get("username", ""))
            ctk.CTkLabel(
                header,
                text=f"👤 {_display}",
                font=ctk.CTkFont(size=11),
                text_color=("gray50", "gray60"),
            ).pack(side="right", padx=16, pady=10)

        # Body — sidebar + content
        body = ctk.CTkFrame(self, fg_color="transparent")
        body.pack(fill="both", expand=True)

        # Sidebar
        sidebar = ctk.CTkFrame(body, width=160, corner_radius=0, fg_color=("gray88", "gray18"))
        sidebar.pack(fill="y", side="left")
        sidebar.pack_propagate(False)

        for label, _ in self._NAV_ITEMS:
            btn = ctk.CTkButton(
                sidebar,
                text=label,
                anchor="w",
                font=ctk.CTkFont(size=13),
                fg_color="transparent",
                text_color=("gray20", "gray90"),
                hover_color=("gray80", "gray28"),
                height=40,
                corner_radius=0,
                command=lambda lbl=label: self._select_section(lbl),
            )
            btn.pack(fill="x", pady=(4, 0), padx=4)
            self._nav_buttons[label] = btn

        # Content area
        self._content_area = ctk.CTkFrame(body, fg_color="transparent")
        self._content_area.pack(fill="both", expand=True, padx=16, pady=12)

    def _select_section(self, label: str) -> None:
        # Highlight selected nav button
        for lbl, btn in self._nav_buttons.items():
            btn.configure(
                fg_color=("gray75", "gray32") if lbl == label else "transparent"
            )
        # Find and call the builder
        for nav_label, method_name in self._NAV_ITEMS:
            if nav_label == label:
                if self._content_frame:
                    self._content_frame.destroy()
                builder = getattr(self, method_name)
                self._content_frame = builder()
                self._content_frame.pack(fill="both", expand=True)
                return

    # ── Section builders ──────────────────────────────────────────────────────

    def _build_users_view(self) -> ctk.CTkFrame:
        from src.gui.settings.user_management_view import UserManagementView
        return UserManagementView(
            self._content_area,
            user_manager=self.user_manager,
            current_user=self.current_user,
        )
