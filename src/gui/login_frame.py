"""Login frame — shown on app startup in the 500×400 home window."""
from __future__ import annotations

import threading
from typing import TYPE_CHECKING, Callable

import customtkinter as ctk

from src.version import __version__

if TYPE_CHECKING:
    from src.user_manager import UserManager


class LoginFrame(ctk.CTkFrame):
    """Full-window login frame.

    Shown in the 500×400 home window before any workflow is accessible.
    On successful login, calls *on_login_success(user_dict)*.
    """

    def __init__(
        self,
        master,
        user_manager: "UserManager",
        on_login_success: Callable[[dict], None],
        **kwargs,
    ) -> None:
        super().__init__(master, fg_color="transparent", **kwargs)
        self._um = user_manager
        self._on_success = on_login_success
        self._build_ui()

    # ── UI ────────────────────────────────────────────────────────────────────

    def _build_ui(self) -> None:
        # Banner (matches HomeFrame)
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

        # Centre form
        form = ctk.CTkFrame(self, fg_color="transparent")
        form.pack(expand=True, fill="both", padx=60, pady=20)

        ctk.CTkLabel(
            form,
            text="Welcome back",
            font=ctk.CTkFont(size=15, weight="bold"),
        ).pack(pady=(0, 20))

        # Username dropdown
        ctk.CTkLabel(form, text="Username", anchor="w").pack(fill="x")
        self._username_var = ctk.StringVar()
        usernames = self._load_usernames()
        _PLACEHOLDER = "Select user…"
        self._username_combo = ctk.CTkComboBox(
            form,
            variable=self._username_var,
            values=[_PLACEHOLDER] + usernames,
            state="readonly",
            height=36,
        )
        self._username_var.set(_PLACEHOLDER)
        self._username_combo.pack(fill="x", pady=(2, 12))

        # Password
        ctk.CTkLabel(form, text="Password", anchor="w").pack(fill="x")
        self._password_var = ctk.StringVar()
        self._password_entry = ctk.CTkEntry(
            form, textvariable=self._password_var, show="•", height=36
        )
        self._password_entry.pack(fill="x", pady=(2, 20))
        self._password_entry.bind("<Return>", lambda _e: self._do_login())

        # Log In button
        self._login_btn = ctk.CTkButton(
            form,
            text="Log In",
            font=ctk.CTkFont(size=14, weight="bold"),
            height=40,
            command=self._do_login,
        )
        self._login_btn.pack(fill="x")

        # Error / status label
        self._status_var = ctk.StringVar()
        self._status_lbl = ctk.CTkLabel(
            form,
            textvariable=self._status_var,
            text_color=("red", "#FF6B6B"),
            font=ctk.CTkFont(size=12),
            wraplength=340,
        )
        self._status_lbl.pack(pady=(12, 0))

        # Focus password field (user is already selected)
        self.after(100, self._password_entry.focus_set)

    # ── Helpers ───────────────────────────────────────────────────────────────

    def _load_usernames(self) -> list:
        """Return list of active usernames from users.json."""
        try:
            data = self._um.load()
            return [
                u["username"]
                for u in data.get("users", [])
                if u.get("active", True)
            ]
        except Exception:
            return []

    # ── Login logic ───────────────────────────────────────────────────────────

    def _set_busy(self, busy: bool) -> None:
        state = "disabled" if busy else "normal"
        self._login_btn.configure(state=state)

    def _do_login(self) -> None:
        username = self._username_var.get().strip()
        password = self._password_var.get()
        if not username or username == "Select user…":
            self._status_var.set("Please select a username.")
            return
        if not password:
            self._status_var.set("Please enter your password.")
            return

        self._set_busy(True)
        self._status_var.set("Checking…")

        def _run() -> None:
            # Step 1: verify credentials
            user = self._um.authenticate(username, password)
            if user is None:
                self.after(0, lambda: self._on_auth_failed())
                return
            # Step 2: attempt login (checks multi-device conflict)
            from src.user_manager import LoginResult
            result: LoginResult = self._um.login(user["username"])
            self.after(0, lambda: self._on_login_result(result, user))

        threading.Thread(target=_run, daemon=True).start()

    def _on_auth_failed(self) -> None:
        self._status_var.set("Incorrect username or password.")
        self._set_busy(False)
        self._password_var.set("")
        self._password_entry.focus_set()

    def _on_login_result(self, result, user: dict) -> None:
        from src.user_manager import LoginResult
        from tkinter import messagebox

        if result.success:
            self._on_success(result.user or user)
            return

        if result.conflict_device:
            name = f"{user.get('first_name', '')} {user.get('last_name', '')}".strip() or user["username"]
            proceed = messagebox.askyesno(
                "Already Logged In",
                f"{name} is still logged in on {result.conflict_device}.\n\nLog in here instead?",
                parent=self.winfo_toplevel(),
            )
            if proceed:
                # Force login
                result2 = self._um.login(user["username"], force=True)
                if result2.success:
                    self._on_success(result2.user or user)
                    return
                self._status_var.set(result2.error or "Login failed.")
            self._set_busy(False)
            return

        self._status_var.set(result.error or "Login failed.")
        self._set_busy(False)
