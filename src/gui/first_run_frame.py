"""First-run setup frame — shown when no users exist in users.json."""
from __future__ import annotations

from typing import TYPE_CHECKING, Callable

import customtkinter as ctk

from src.version import __version__

if TYPE_CHECKING:
    from src.user_manager import UserManager


class FirstRunFrame(ctk.CTkFrame):
    """Guides the user through creating the first admin account.

    Shown in the 500×400 home window when users.json has no users yet.
    On success, calls *on_complete(user_dict)*.
    """

    def __init__(
        self,
        master,
        user_manager: "UserManager",
        on_complete: Callable[[dict], None],
        **kwargs,
    ) -> None:
        super().__init__(master, fg_color="transparent", **kwargs)
        self._um = user_manager
        self._on_complete = on_complete
        self._build_ui()

    # ── UI ────────────────────────────────────────────────────────────────────

    def _build_ui(self) -> None:
        # Banner
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

        # Form
        form = ctk.CTkFrame(self, fg_color="transparent")
        form.pack(expand=True, fill="both", padx=50, pady=8)

        ctk.CTkLabel(
            form,
            text="Create Admin Account",
            font=ctk.CTkFont(size=15, weight="bold"),
        ).pack(pady=(0, 2))
        ctk.CTkLabel(
            form,
            text="No users found. Set up the first admin account to continue.",
            font=ctk.CTkFont(size=11),
            text_color=("gray40", "gray65"),
            wraplength=360,
        ).pack(pady=(0, 10))

        def _field(label: str, show: str = "") -> ctk.CTkEntry:
            ctk.CTkLabel(form, text=label, anchor="w").pack(fill="x")
            var = ctk.StringVar()
            entry = ctk.CTkEntry(form, textvariable=var, height=32, show=show)
            entry.pack(fill="x", pady=(2, 6))
            return entry, var

        self._first_entry, self._first_var = _field("First Name")
        self._last_entry, self._last_var = _field("Last Name")
        self._user_entry, self._user_var = _field("Username")
        self._pass_entry, self._pass_var = _field("Password", show="•")
        self._conf_entry, self._conf_var = _field("Confirm Password", show="•")

        # Chain Enter keys
        self._first_entry.bind("<Return>", lambda _: self._last_entry.focus_set())
        self._last_entry.bind("<Return>", lambda _: self._user_entry.focus_set())
        self._user_entry.bind("<Return>", lambda _: self._pass_entry.focus_set())
        self._pass_entry.bind("<Return>", lambda _: self._conf_entry.focus_set())
        self._conf_entry.bind("<Return>", lambda _: self._submit())

        self._create_btn = ctk.CTkButton(
            form,
            text="Create Account",
            font=ctk.CTkFont(size=14, weight="bold"),
            height=40,
            command=self._submit,
        )
        self._create_btn.pack(fill="x", pady=(4, 0))

        self._status_var = ctk.StringVar()
        self._status_lbl = ctk.CTkLabel(
            form,
            textvariable=self._status_var,
            text_color=("red", "#FF6B6B"),
            font=ctk.CTkFont(size=12),
            wraplength=400,
            justify="left",
        )
        self._status_lbl.pack(fill="x", pady=(8, 0))

        self.after(100, self._first_entry.focus_set)

    # ── Submit ────────────────────────────────────────────────────────────────

    def _submit(self) -> None:
        first = self._first_var.get().strip()
        last = self._last_var.get().strip()
        username = self._user_var.get().strip()
        password = self._pass_var.get()
        confirm = self._conf_var.get()

        if not first or not last:
            self._status_var.set("Please enter a first and last name.")
            return
        if not username:
            self._status_var.set("Please enter a username.")
            return
        if not password:
            self._status_var.set("Please enter a password.")
            return
        if password != confirm:
            self._status_var.set("Passwords do not match.")
            return

        self._create_btn.configure(state="disabled")
        self._status_var.set("Creating account…")

        def _run() -> None:
            try:
                user = self._um.create_user(
                    first_name=first,
                    last_name=last,
                    username=username,
                    password=password,
                    role="admin",
                )
                result = self._um.login(username)
                final_user = result.user if result.success else user
                self.after(0, lambda: self._on_complete(final_user))
            except Exception as exc:
                self.after(0, lambda e=exc: self._on_error(str(e)))

        import threading
        threading.Thread(target=_run, daemon=True).start()

    def _on_error(self, msg: str) -> None:
        self._status_var.set(msg)
        self._create_btn.configure(state="normal")
