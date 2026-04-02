"""User management and session tracking for the Scarlett AIO app.

Handles user CRUD, authentication, login/logout, heartbeat, and order-processing
state.  All data is persisted to a shared JSON file on the network.

No GUI code in this module.
"""
from __future__ import annotations

import base64
import hashlib
import hmac
import json
import logging
import os
import socket
import time
import uuid
from dataclasses import dataclass, field
from datetime import datetime, timedelta
from pathlib import Path
from typing import Optional

log = logging.getLogger(__name__)

# Session is considered stale if last_heartbeat is older than this
_STALE_MINUTES = 3

from src.config import SERVER_AVAILABLE, LOCAL_DATA_DIR  # noqa: E402

_SERVER_USERS_FILE = Path(r"\\SERVER\Project Folder\Order-Fulfillment-App\Config\users.json")
USERS_FILE = _SERVER_USERS_FILE if SERVER_AVAILABLE else LOCAL_DATA_DIR / "users.json"


# ── Data helpers ──────────────────────────────────────────────────────────────

@dataclass
class LoginResult:
    success: bool
    user: Optional[dict] = None
    conflict_device: Optional[str] = None  # set if active session exists on another device
    error: str = ""


# ── Password hashing (stdlib only, no extra deps) ─────────────────────────────

def _hash_password(password: str) -> str:
    """Return a storable hash string for *password*."""
    salt = os.urandom(16)
    dk = hashlib.pbkdf2_hmac("sha256", password.encode("utf-8"), salt, 200_000)
    return (
        "pbkdf2:sha256:"
        + base64.b64encode(salt).decode()
        + ":"
        + base64.b64encode(dk).decode()
    )


def _check_password(password: str, stored: str) -> bool:
    """Return True if *password* matches *stored* hash."""
    try:
        _, algo, b64_salt, b64_dk = stored.split(":")
        if algo != "sha256":
            return False
        salt = base64.b64decode(b64_salt)
        expected = base64.b64decode(b64_dk)
        dk = hashlib.pbkdf2_hmac("sha256", password.encode("utf-8"), salt, 200_000)
        # Constant-time comparison
        return hmac.compare_digest(dk, expected)
    except Exception:
        return False


# ── UserManager ───────────────────────────────────────────────────────────────

class UserManager:
    """Manage users, sessions, and order-processing state via a shared JSON file."""

    def __init__(self, path: Optional[Path] = None) -> None:
        self._path = Path(path) if path else USERS_FILE

    # ── Low-level I/O ─────────────────────────────────────────────────────────

    def load(self) -> dict:
        """Read and return the full data dict.  Returns empty structure if file doesn't exist yet."""
        try:
            data = json.loads(self._path.read_text(encoding="utf-8"))
        except FileNotFoundError:
            data = {}
        if "users" not in data:
            data["users"] = []
        if "sessions" not in data:
            data["sessions"] = {}
        return data

    def save(self, data: dict) -> None:
        """Write *data* atomically.  Retries up to 3 times on OSError."""
        tmp = self._path.with_suffix(".tmp")
        payload = json.dumps(data, indent=2, ensure_ascii=False)
        for attempt in range(3):
            try:
                self._path.parent.mkdir(parents=True, exist_ok=True)
                tmp.write_text(payload, encoding="utf-8")
                os.replace(tmp, self._path)
                return
            except OSError as exc:
                if attempt == 2:
                    raise
                log.warning("users.json write attempt %d failed: %s", attempt + 1, exc)
                time.sleep(0.2)

    # ── Existence check ───────────────────────────────────────────────────────

    def has_any_users(self) -> bool:
        """Return True if at least one user record exists."""
        try:
            data = self.load()
            return bool(data.get("users"))
        except (OSError, ValueError):
            return False

    def is_available(self) -> bool:
        """Return True if the network file is reachable."""
        try:
            # Just check the parent directory exists
            return self._path.parent.exists()
        except OSError:
            return False

    # ── User CRUD ─────────────────────────────────────────────────────────────

    def get_active_users(self) -> list[dict]:
        """Return all active user records (without password_hash)."""
        data = self.load()
        return [_strip_hash(u) for u in data["users"] if u.get("active", True)]

    def authenticate(self, username: str, password: str) -> Optional[dict]:
        """Return the user dict (without password_hash) if credentials are valid, else None."""
        try:
            data = self.load()
        except Exception:
            return None
        for user in data.get("users", []):
            if (
                user.get("username", "").lower() == username.lower()
                and user.get("active", True)
                and _check_password(password, user.get("password_hash", ""))
            ):
                return _strip_hash(user)
        return None

    def create_user(
        self,
        first_name: str,
        last_name: str,
        username: str,
        password: str,
        role: str = "user",
        neto_api: str = "",
    ) -> dict:
        """Create a new user.  Raises ValueError on duplicate username."""
        data = self.load()
        if any(u["username"].lower() == username.lower() for u in data["users"]):
            raise ValueError(f"Username '{username}' is already taken.")
        user: dict = {
            "user_id": str(uuid.uuid4()),
            "username": username,
            "first_name": first_name,
            "last_name": last_name,
            "password_hash": _hash_password(password),
            "role": role,
            "neto_api": neto_api,
            "created_at": datetime.now().isoformat(timespec="seconds"),
            "active": True,
        }
        data["users"].append(user)
        self.save(data)
        log.info("Created user '%s' (role=%s)", username, role)
        return _strip_hash(user)

    def update_user(self, user_id: str, **fields) -> dict:
        """Update allowed fields for a user.  Returns updated user (no hash)."""
        _ALLOWED = {"first_name", "last_name", "role", "neto_api"}
        data = self.load()
        for user in data["users"]:
            if user["user_id"] == user_id:
                for k, v in fields.items():
                    if k in _ALLOWED:
                        user[k] = v
                self.save(data)
                return _strip_hash(user)
        raise ValueError(f"User ID '{user_id}' not found.")

    def change_password(self, user_id: str, new_password: str) -> None:
        """Set a new password for the given user."""
        data = self.load()
        for user in data["users"]:
            if user["user_id"] == user_id:
                user["password_hash"] = _hash_password(new_password)
                self.save(data)
                log.info("Password changed for user_id=%s", user_id)
                return
        raise ValueError(f"User ID '{user_id}' not found.")

    def deactivate_user(self, user_id: str) -> None:
        """Soft-delete a user (sets active=False)."""
        data = self.load()
        for user in data["users"]:
            if user["user_id"] == user_id:
                user["active"] = False
                self.save(data)
                # Clear any active session
                username = user.get("username", "")
                if username in data.get("sessions", {}):
                    data["sessions"].pop(username, None)
                    self.save(data)
                log.info("Deactivated user_id=%s (%s)", user_id, username)
                return
        raise ValueError(f"User ID '{user_id}' not found.")

    def get_user_by_id(self, user_id: str) -> Optional[dict]:
        """Return user dict (no hash) or None."""
        try:
            data = self.load()
        except Exception:
            return None
        for u in data.get("users", []):
            if u.get("user_id") == user_id:
                return _strip_hash(u)
        return None

    # ── Login / Logout ────────────────────────────────────────────────────────

    def login(self, username: str, device_name: str = "", force: bool = False) -> LoginResult:
        """Record a login for *username*.

        Returns LoginResult.  If another fresh session exists on a different device,
        returns ``LoginResult(success=False, conflict_device=<name>)`` unless
        *force* is True (in which case the old session is cleared and login proceeds).
        """
        if not device_name:
            try:
                device_name = socket.gethostname()
            except Exception:
                device_name = "unknown"
        try:
            data = self.load()
        except Exception as exc:
            return LoginResult(success=False, error=f"Cannot read users file: {exc}")

        # Find the user record
        user = next(
            (u for u in data["users"] if u.get("username", "").lower() == username.lower()),
            None,
        )
        if not user:
            return LoginResult(success=False, error="User not found.")
        if not user.get("active", True):
            return LoginResult(success=False, error="Account is inactive.")

        sessions = data.setdefault("sessions", {})
        existing = sessions.get(username, {})

        # Check for active session on a different device
        if (
            not force
            and existing.get("logged_in")
            and existing.get("device_name", "") != device_name
            and not self.is_session_stale(existing)
        ):
            return LoginResult(
                success=False,
                conflict_device=existing.get("device_name", "another device"),
            )

        # Record login
        sessions[username] = {
            "logged_in": True,
            "device_name": device_name,
            "last_heartbeat": datetime.now().isoformat(timespec="seconds"),
            "processing_order_id": None,
            "processing_since": None,
        }
        self.save(data)
        log.info("Login: %s on %s", username, device_name)
        return LoginResult(success=True, user=_strip_hash(user))

    def logout(self, username: str) -> None:
        """Clear login and processing state for *username*."""
        try:
            data = self.load()
            sessions = data.setdefault("sessions", {})
            if username in sessions:
                sessions[username]["logged_in"] = False
                sessions[username]["processing_order_id"] = None
                sessions[username]["processing_since"] = None
                self.save(data)
                log.info("Logout: %s", username)
        except Exception as exc:
            log.warning("logout(%s) failed: %s", username, exc)

    # ── Heartbeat ─────────────────────────────────────────────────────────────

    def heartbeat(self, username: str) -> None:
        """Update last_heartbeat for *username*.  Silent on error."""
        try:
            data = self.load()
            sessions = data.setdefault("sessions", {})
            if username not in sessions:
                return
            sessions[username]["last_heartbeat"] = datetime.now().isoformat(timespec="seconds")
            self.save(data)
        except Exception as exc:
            log.debug("heartbeat(%s) failed: %s", username, exc)

    def is_session_stale(self, session: dict) -> bool:
        """Return True if *session*'s last_heartbeat is older than STALE_MINUTES."""
        hb = session.get("last_heartbeat")
        if not hb:
            return True
        try:
            ts = datetime.fromisoformat(hb)
            return datetime.now() - ts > timedelta(minutes=_STALE_MINUTES)
        except Exception:
            return True

    def get_session(self, username: str) -> dict:
        """Return current session dict for *username*, or empty dict."""
        try:
            data = self.load()
            return data.get("sessions", {}).get(username, {})
        except Exception:
            return {}

    # ── Order processing state ────────────────────────────────────────────────

    def set_processing_order(self, username: str, order_id: str) -> None:
        """Mark *username* as actively processing *order_id*."""
        try:
            data = self.load()
            sessions = data.setdefault("sessions", {})
            if username in sessions:
                sessions[username]["processing_order_id"] = order_id
                sessions[username]["processing_since"] = datetime.now().isoformat(timespec="seconds")
                self.save(data)
        except Exception as exc:
            log.warning("set_processing_order failed: %s", exc)

    def clear_processing_order(self, username: str) -> None:
        """Clear the processing flag for *username*."""
        try:
            data = self.load()
            sessions = data.setdefault("sessions", {})
            if username in sessions:
                sessions[username]["processing_order_id"] = None
                sessions[username]["processing_since"] = None
                self.save(data)
        except Exception as exc:
            log.warning("clear_processing_order failed: %s", exc)

    def get_processing_user(self, order_id: str) -> Optional[str]:
        """Return the display name (First Last) of whoever is processing *order_id*.

        Returns None if no active (non-stale) session is processing it.
        """
        try:
            data = self.load()
            sessions = data.get("sessions", {})
            users_by_username = {u["username"]: u for u in data.get("users", [])}
            for username, session in sessions.items():
                if (
                    session.get("processing_order_id") == order_id
                    and session.get("logged_in")
                    and not self.is_session_stale(session)
                ):
                    user = users_by_username.get(username, {})
                    first = user.get("first_name", "")
                    last = user.get("last_name", "")
                    return f"{first} {last}".strip() or username
        except Exception as exc:
            log.warning("get_processing_user(%s) failed: %s", order_id, exc)
        return None

    def get_all_sessions(self) -> dict:
        """Return the full sessions dict.  Returns {} on error."""
        try:
            return self.load().get("sessions", {})
        except Exception:
            return {}


# ── Helpers ───────────────────────────────────────────────────────────────────

def _strip_hash(user: dict) -> dict:
    """Return a copy of *user* without the password_hash key."""
    return {k: v for k, v in user.items() if k != "password_hash"}
