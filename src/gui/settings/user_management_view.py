"""User management view — shown inside the Settings window."""
from __future__ import annotations

import tkinter as tk
from tkinter import messagebox, ttk
from typing import TYPE_CHECKING

import customtkinter as ctk

if TYPE_CHECKING:
    from src.user_manager import UserManager


class UserManagementView(ctk.CTkFrame):
    """Displays the user list and provides add / edit / deactivate controls.

    Admins see all users and all controls.
    Non-admins see only their own profile with a limited edit form.
    """

    def __init__(self, master, user_manager: "UserManager", current_user: dict, **kwargs) -> None:
        super().__init__(master, fg_color="transparent", **kwargs)
        self._um = user_manager
        self._current_user = current_user
        self._is_admin = current_user.get("role") == "admin"
        self._build_ui()
        self._reload()

    # ── UI ────────────────────────────────────────────────────────────────────

    def _build_ui(self) -> None:
        ctk.CTkLabel(
            self,
            text="Users",
            font=ctk.CTkFont(size=15, weight="bold"),
        ).pack(anchor="w", pady=(0, 10))

        if self._is_admin:
            self._build_admin_ui()
        else:
            self._build_profile_ui()

    # ── Admin view ────────────────────────────────────────────────────────────

    def _build_admin_ui(self) -> None:
        # Toolbar
        toolbar = ctk.CTkFrame(self, fg_color="transparent")
        toolbar.pack(fill="x", pady=(0, 8))

        self._new_btn = ctk.CTkButton(toolbar, text="New User", width=100, command=self._new_user)
        self._new_btn.pack(side="left", padx=(0, 6))
        self._edit_btn = ctk.CTkButton(
            toolbar, text="Edit", width=80, state="disabled", command=self._edit_user
        )
        self._edit_btn.pack(side="left", padx=(0, 6))
        self._deactivate_btn = ctk.CTkButton(
            toolbar, text="Deactivate", width=100,
            fg_color=("gray60", "gray35"), hover_color=("gray50", "gray28"),
            state="disabled", command=self._deactivate_user,
        )
        self._deactivate_btn.pack(side="left", padx=(0, 6))
        self._reset_pwd_btn = ctk.CTkButton(
            toolbar, text="Reset Password", width=130,
            fg_color="transparent", border_width=1,
            state="disabled", command=self._reset_password,
        )
        self._reset_pwd_btn.pack(side="left")

        # Treeview
        cols = ("name", "username", "role", "status")
        self._tree = ttk.Treeview(self, columns=cols, show="headings", height=12)
        for col, heading, width in (
            ("name",     "Name",     180),
            ("username", "Username", 140),
            ("role",     "Role",      80),
            ("status",   "Status",    80),
        ):
            self._tree.heading(col, text=heading)
            self._tree.column(col, width=width, anchor="w", stretch=False)
        self._tree.pack(fill="both", expand=True)
        self._tree.bind("<<TreeviewSelect>>", self._on_select)
        self._tree.bind("<Double-1>", lambda _e: self._edit_user())

    # ── Non-admin profile view ────────────────────────────────────────────────

    def _build_profile_ui(self) -> None:
        u = self._current_user
        form = ctk.CTkFrame(self, fg_color="transparent")
        form.pack(fill="x")

        def _row(label: str, value: str) -> None:
            row = ctk.CTkFrame(form, fg_color="transparent")
            row.pack(fill="x", pady=3)
            ctk.CTkLabel(row, text=label, width=120, anchor="w",
                         text_color=("gray40", "gray65")).pack(side="left")
            ctk.CTkLabel(row, text=value, anchor="w").pack(side="left")

        first = u.get("first_name", "")
        last = u.get("last_name", "")
        _row("Name:", f"{first} {last}".strip())
        _row("Username:", u.get("username", ""))
        _row("Role:", u.get("role", "").capitalize())

        ctk.CTkButton(
            self,
            text="Edit Profile",
            width=140,
            command=lambda: self._edit_user(user_id=u["user_id"]),
        ).pack(anchor="w", pady=(16, 0))

    # ── Data ──────────────────────────────────────────────────────────────────

    def _reload(self) -> None:
        if not self._is_admin:
            return
        self._tree.delete(*self._tree.get_children())
        try:
            data = self._um.load()
            users = data.get("users", [])
        except Exception as exc:
            messagebox.showerror("Error", f"Could not load users:\n{exc}", parent=self.winfo_toplevel())
            return
        for u in users:
            first = u.get("first_name", "")
            last = u.get("last_name", "")
            status = "Active" if u.get("active", True) else "Inactive"
            iid = u["user_id"]
            self._tree.insert("", "end", iid=iid, values=(
                f"{first} {last}".strip(),
                u.get("username", ""),
                u.get("role", "").capitalize(),
                status,
            ))
        self._update_toolbar()

    def _selected_user_id(self) -> str | None:
        sel = self._tree.selection()
        return sel[0] if sel else None

    def _selected_user_dict(self) -> dict | None:
        uid = self._selected_user_id()
        if not uid:
            return None
        try:
            data = self._um.load()
            return next((u for u in data["users"] if u["user_id"] == uid), None)
        except Exception:
            return None

    def _on_select(self, _event=None) -> None:
        self._update_toolbar()

    def _update_toolbar(self) -> None:
        sel = self._selected_user_id()
        state = "normal" if sel else "disabled"
        self._edit_btn.configure(state=state)
        self._deactivate_btn.configure(state=state)
        self._reset_pwd_btn.configure(state=state)

    # ── Actions ───────────────────────────────────────────────────────────────

    def _new_user(self) -> None:
        UserEditDialog(
            self.winfo_toplevel(),
            user_manager=self._um,
            current_user=self._current_user,
            on_save=self._reload,
        )

    def _edit_user(self, user_id: str | None = None) -> None:
        uid = user_id or self._selected_user_id()
        if not uid:
            return
        try:
            data = self._um.load()
            user = next((u for u in data["users"] if u["user_id"] == uid), None)
        except Exception:
            user = None
        if not user:
            return
        UserEditDialog(
            self.winfo_toplevel(),
            user_manager=self._um,
            current_user=self._current_user,
            existing_user=user,
            on_save=self._reload,
        )

    def _deactivate_user(self) -> None:
        uid = self._selected_user_id()
        if not uid:
            return
        user = self._selected_user_dict()
        if not user:
            return
        name = f"{user.get('first_name', '')} {user.get('last_name', '')}".strip() or user["username"]
        if not messagebox.askyesno(
            "Deactivate User",
            f"Deactivate {name}?\n\nThey will no longer be able to log in.",
            parent=self.winfo_toplevel(),
        ):
            return
        try:
            self._um.deactivate_user(uid)
            self._reload()
        except Exception as exc:
            messagebox.showerror("Error", str(exc), parent=self.winfo_toplevel())

    def _reset_password(self) -> None:
        uid = self._selected_user_id()
        if not uid:
            return
        user = self._selected_user_dict()
        if not user:
            return
        ResetPasswordDialog(
            self.winfo_toplevel(),
            user_manager=self._um,
            user=user,
        )


# ── User edit dialog ──────────────────────────────────────────────────────────

class UserEditDialog(ctk.CTkToplevel):
    """Add or edit a user."""

    def __init__(self, master, user_manager: "UserManager", current_user: dict,
                 existing_user: dict | None = None, on_save=None, **kwargs) -> None:
        super().__init__(master, **kwargs)
        self._um = user_manager
        self._current_user = current_user
        self._existing = existing_user
        self._on_save = on_save
        self._is_admin = current_user.get("role") == "admin"
        is_edit = existing_user is not None

        self.title("Edit User" if is_edit else "New User")
        self.geometry("480x760")
        self.resizable(True, True)
        self.transient(master)
        self.grab_set()

        self._build_ui(is_edit)
        self.after(50, self.lift)

    def _build_ui(self, is_edit: bool) -> None:
        form = ctk.CTkFrame(self, fg_color="transparent")
        form.pack(fill="both", expand=True, padx=24, pady=20)

        ctk.CTkLabel(
            form,
            text="Edit User" if is_edit else "New User",
            font=ctk.CTkFont(size=14, weight="bold"),
        ).pack(anchor="w", pady=(0, 14))

        def _field(label: str, initial: str = "", show: str = "") -> tuple:
            ctk.CTkLabel(form, text=label, anchor="w").pack(fill="x")
            var = ctk.StringVar(value=initial)
            entry = ctk.CTkEntry(form, textvariable=var, height=34, show=show)
            entry.pack(fill="x", pady=(2, 10))
            return entry, var

        u = self._existing or {}
        _, self._first_var = _field("First Name", u.get("first_name", ""))
        _, self._last_var = _field("Last Name", u.get("last_name", ""))
        _, self._user_var = _field("Username", u.get("username", ""))

        if self._is_admin:
            ctk.CTkLabel(form, text="Role", anchor="w").pack(fill="x")
            self._role_var = ctk.StringVar(value=u.get("role", "user").capitalize())
            ctk.CTkComboBox(
                form,
                variable=self._role_var,
                values=["User", "Admin"],
                state="readonly",
                height=34,
            ).pack(fill="x", pady=(2, 10))

            ctk.CTkLabel(form, text="Neto API Key (optional — for future use)",
                         anchor="w", text_color=("gray40", "gray65")).pack(fill="x")
            self._neto_var = ctk.StringVar(value=u.get("neto_api", ""))
            ctk.CTkEntry(form, textvariable=self._neto_var, height=34).pack(fill="x", pady=(2, 10))

        pwd_hint = "Password (leave blank to keep current)" if is_edit else "Password"
        _, self._pwd_var = _field(pwd_hint, show="•")
        _, self._conf_var = _field("Confirm Password", show="•")

        self._status_var = ctk.StringVar()
        ctk.CTkLabel(
            form,
            textvariable=self._status_var,
            text_color=("red", "#FF6B6B"),
            font=ctk.CTkFont(size=11),
            wraplength=360,
        ).pack(pady=(0, 6))

        btn_row = ctk.CTkFrame(form, fg_color="transparent")
        btn_row.pack(fill="x")
        ctk.CTkButton(btn_row, text="Save", command=self._save).pack(side="left", padx=(0, 8))
        ctk.CTkButton(
            btn_row, text="Cancel",
            fg_color="transparent", border_width=1,
            command=self.destroy,
        ).pack(side="left")

    def _save(self) -> None:
        first = self._first_var.get().strip()
        last = self._last_var.get().strip()
        username = self._user_var.get().strip()
        password = self._pwd_var.get()
        confirm = self._conf_var.get()

        if not first or not last:
            self._status_var.set("First and last name are required.")
            return
        if not username:
            self._status_var.set("Username is required.")
            return
        if password and password != confirm:
            self._status_var.set("Passwords do not match.")
            return
        if not self._existing and not password:
            self._status_var.set("Password is required for new users.")
            return

        role = "user"
        neto_api = ""
        if self._is_admin:
            role = self._role_var.get().lower()
            neto_api = self._neto_var.get().strip()

        try:
            if self._existing:
                self._um.update_user(
                    self._existing["user_id"],
                    first_name=first,
                    last_name=last,
                    role=role,
                    neto_api=neto_api,
                )
                if password:
                    self._um.change_password(self._existing["user_id"], password)
            else:
                self._um.create_user(
                    first_name=first,
                    last_name=last,
                    username=username,
                    password=password,
                    role=role,
                    neto_api=neto_api,
                )
        except Exception as exc:
            self._status_var.set(str(exc))
            return

        if self._on_save:
            self._on_save()
        self.destroy()


# ── Reset password dialog ─────────────────────────────────────────────────────

class ResetPasswordDialog(ctk.CTkToplevel):
    """Admin resets a user's password."""

    def __init__(self, master, user_manager: "UserManager", user: dict, **kwargs) -> None:
        super().__init__(master, **kwargs)
        self._um = user_manager
        self._user = user
        name = f"{user.get('first_name', '')} {user.get('last_name', '')}".strip() or user["username"]

        self.title(f"Reset Password — {name}")
        self.geometry("360x260")
        self.resizable(False, False)
        self.transient(master)
        self.grab_set()

        form = ctk.CTkFrame(self, fg_color="transparent")
        form.pack(fill="both", expand=True, padx=24, pady=20)

        ctk.CTkLabel(form, text=f"Reset password for {name}",
                     font=ctk.CTkFont(size=13, weight="bold")).pack(anchor="w", pady=(0, 14))

        ctk.CTkLabel(form, text="New Password", anchor="w").pack(fill="x")
        self._pwd_var = ctk.StringVar()
        ctk.CTkEntry(form, textvariable=self._pwd_var, show="•", height=34).pack(fill="x", pady=(2, 10))
        ctk.CTkLabel(form, text="Confirm Password", anchor="w").pack(fill="x")
        self._conf_var = ctk.StringVar()
        ctk.CTkEntry(form, textvariable=self._conf_var, show="•", height=34).pack(fill="x", pady=(2, 10))

        self._status_var = ctk.StringVar()
        ctk.CTkLabel(form, textvariable=self._status_var,
                     text_color=("red", "#FF6B6B"), font=ctk.CTkFont(size=11)).pack(pady=(0, 8))

        btn_row = ctk.CTkFrame(form, fg_color="transparent")
        btn_row.pack(fill="x")
        ctk.CTkButton(btn_row, text="Save", command=self._save).pack(side="left", padx=(0, 8))
        ctk.CTkButton(btn_row, text="Cancel", fg_color="transparent",
                      border_width=1, command=self.destroy).pack(side="left")

        self.after(50, self.lift)

    def _save(self) -> None:
        pwd = self._pwd_var.get()
        conf = self._conf_var.get()
        if not pwd:
            self._status_var.set("Password cannot be empty.")
            return
        if pwd != conf:
            self._status_var.set("Passwords do not match.")
            return
        try:
            self._um.change_password(self._user["user_id"], pwd)
            self.destroy()
        except Exception as exc:
            self._status_var.set(str(exc))
