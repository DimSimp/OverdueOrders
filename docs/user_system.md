# User System — Design & Reference Document

> **Status**: In Progress
> **Last updated**: 2026-03-27
> **Scope**: Phase 1 — core user system, login, settings window, order conflict detection, order assignment

---

## Context

The app is currently single-user and untracked — anyone with the exe can open any order.
The immediate need is preventing duplicate processing (two staff accidentally working the same
order at the same time) and assigning orders to specific users. The user system is also
foundational for the future POS/inventory management direction, so user accounts and roles are
being established now even though some fields (e.g. Neto API key per user) won't be used yet.

---

## 1. Data Storage

### Users & Sessions
**File**: `\\SERVER\Project Folder\Order-Fulfillment-App\Config\users.json`

```json
{
  "users": [
    {
      "user_id": "uuid4-string",
      "username": "jsmith",
      "first_name": "John",
      "last_name": "Smith",
      "password_hash": "pbkdf2:sha256:<b64-salt>:<b64-hash>",
      "role": "admin",
      "neto_api": "",
      "created_at": "2026-03-27T10:00:00",
      "active": true
    }
  ],
  "sessions": {
    "jsmith": {
      "logged_in": true,
      "device_name": "SCARLETT-PC1",
      "last_heartbeat": "2026-03-27T14:30:00",
      "processing_order_id": null,
      "processing_since": null
    }
  }
}
```

- **`users`** — persistent user definitions; never hard-deleted, only deactivated (`active: false`)
- **`sessions`** — volatile runtime state keyed by username; written frequently
- **`password_hash`** — `pbkdf2_hmac('sha256', ...)` via stdlib `hashlib` + `os.urandom(16)` salt (no extra deps)
- **`role`** — `"admin"` or `"user"`

### Order Assignments
**File**: `\\SERVER\Project Folder\Order-Fulfillment-App\Config\order_assignments.json`

```json
{
  "assignments": {
    "N1234": {
      "assigned_to": "jsmith",
      "assigned_by": "admin_username",
      "assigned_at": "2026-03-27T14:00:00"
    }
  }
}
```

- Keyed by order ID (Neto order number or eBay order ID)
- Assignments persist until explicitly removed or the order is marked as Sent/Dispatched

### File I/O Rules
- **Atomic writes**: write to `<file>.tmp`, then `os.replace()` → prevents corrupt reads
- **Retry loop**: up to 3 attempts with 200 ms sleep on `OSError`
- Full read-modify-write on every operation (files are small, ~a few KB)

---

## 2. Core Modules

### `src/user_manager.py` — `UserManager`

```
UserManager(network_path: Path | str)
  ├── load() → dict                         read-parse users.json; raises on unavailable
  ├── save(data: dict) → None               atomic write with retry
  │
  ├── has_any_users() → bool                False if file missing or users list empty
  ├── authenticate(username, password) → dict | None
  ├── create_user(admin, first, last, username, password, role) → dict
  ├── update_user(admin, user_id, **fields) → dict
  ├── deactivate_user(admin, user_id) → None
  ├── change_password(user_id, new_password) → None
  ├── get_active_users() → list[dict]
  │
  ├── login(username, device_name, force=False) → LoginResult
  │     checks session staleness, multi-device conflict, sets logged_in + heartbeat
  ├── logout(username) → None
  │     clears logged_in, clears processing state
  │
  ├── heartbeat(username) → None            update last_heartbeat timestamp
  ├── is_session_stale(session: dict) → bool   True if last_heartbeat > 3 min ago
  │
  ├── set_processing_order(username, order_id) → None
  ├── clear_processing_order(username) → None
  └── get_processing_user(order_id) → str | None
        returns display name of user processing this order (if session is fresh)
```

**`LoginResult` dataclass**:
```python
@dataclass
class LoginResult:
    success: bool
    user: dict | None = None
    conflict_device: str | None = None   # set if logged in elsewhere on a fresh session
    error: str = ""
```

### `src/assignment_manager.py` — `AssignmentManager`

```
AssignmentManager(network_path: Path | str)
  ├── load() → dict
  ├── save(data: dict) → None
  │
  ├── assign(order_ids: list[str], assigned_to: str, assigned_by: str) → None
  ├── unassign(order_ids: list[str]) → None
  ├── get_assignment(order_id: str) → dict | None
  ├── get_all_assignments() → dict[str, dict]
  └── clear_on_dispatch(order_id: str) → None
```

---

## 3. Application Flow

### Startup / Login

1. `App.__init__` instantiates `UserManager` and checks `has_any_users()`
2. **No users** → show `FirstRunFrame` (create initial admin account)
3. **Users exist** → show `LoginFrame`
4. On successful login → `pack_forget` login frame → `pack` `HomeFrame`
5. `App` stores `self._current_user: dict` and `self._user_manager: UserManager`
6. Heartbeat + poll thread starts (10-second interval)

**Server unavailable**: login frame shows error label + "Retry" button. No offline fallback.

### Login Frame (`src/gui/login_frame.py`)

- `LoginFrame(master, user_manager, on_login_success)`
- Shows within the existing 500×400 home window
- Username + password fields, "Log In" button, error label
- **Multi-device conflict**: `messagebox.askyesno` → "still logged in on [device]. Log in here instead?" → `login(force=True)`

### First Run Frame (`src/gui/first_run_frame.py`)

- Shown only when `users.json` has no users
- Collects: First Name, Last Name, Username, Password (+ confirm)
- Creates the first admin account, then proceeds to `HomeFrame`

### Home Screen (`src/gui/home_window.py`)

- `HomeFrame` gains `current_user: dict` and `on_settings` parameters
- Logged-in user's name shown (small label, banner or bottom)
- **Settings button** added below workflow buttons (small, secondary style, height=32)

---

## 4. Heartbeat & Background Polling

A single background daemon thread in `App`, started after login. **10-second interval.**

Each tick:
1. `user_manager.heartbeat(username)` — writes `last_heartbeat`
2. If `self._processing_order_id` is set: read own session — if `processing_order_id` was cleared by someone else → `root.after(0, self._on_processing_taken_over)`
3. Check own `logged_in` flag — if `False` (forced logout from another device) → `root.after(0, self._force_relogin)`

**`_on_processing_taken_over()`**: shows popup, calls registered `_order_close_callback()`, clears `_processing_order_id`.

**`_force_relogin()`**: closes child windows, shows `LoginFrame` again.

---

## 5. Session Cleanup on Close

**`App.protocol("WM_DELETE_WINDOW", self._logout_and_exit)`**:
- Calls `user_manager.logout(username)` — clears `logged_in` + `processing_order_id`
- Then `self.destroy()`

**`DailyOpsWindow.protocol("WM_DELETE_WINDOW", self._on_close)`**:
- Clears `processing_order_id` only (does NOT log out)
- Destroys window

Same pattern for the Afternoon Ops CTkToplevel.

**Crash handling**: heartbeat stops when process dies. Other instances see `last_heartbeat` > 3 min → treat as stale → auto-recover. Crash recovery is automatic within ~3 minutes.

---

## 6. Order Conflict Detection

When an order detail is opened (in `results_tab.py` or `results_view.py`), before showing `OrderDetailView`:

```python
other_user = user_manager.get_processing_user(order_id)
if other_user and other_user != current_user["username"]:
    answer = messagebox.askyesno(
        "Order In Progress",
        f"{other_user} is currently processing order #{order_id}.\nTake over?"
    )
    if not answer:
        return  # abort
# Force takeover (or fresh open)
user_manager.set_processing_order(current_user["username"], order_id)
app._processing_order_id = order_id
```

`OrderDetailView` gains `on_processing_end` callback → called on back/fulfilled/destroy to clear the flag.

When displaced user's poll detects the takeover → popup → their order detail window is disabled/closed.

---

## 7. Settings Window (`src/gui/settings_window.py`)

- `SettingsWindow(master, user_manager, current_user)` — `CTkToplevel`, ~700×500px
- Sidebar nav on left, content area on right (frame-swap pattern)
- First section: **Users** (admin sees full list + management; non-admin sees own profile only)
- More sections added over time

---

## 8. User Management View (`src/gui/settings/user_management_view.py`)

- `ttk.Treeview` — columns: Name, Username, Role, Status
- Toolbar: **New User | Edit | Deactivate | Reset Password**
- New/Edit dialog fields: First Name, Last Name, Username, Password (+ Confirm), Role, Neto API Key (optional)
- Deactivate: sets `active: false`; deactivated users cannot log in
- Non-admins: edit own profile only (name + password)

---

## 9. Order Assignment

### Bulk assignment (primary)
1. User checks ≥1 orders via the existing ☐/☑ checkbox system
2. Action bar appears → **"Assign to…"** button (admin only; already a placeholder in the action bar)
3. Popup: dropdown of active users + "Unassign" option
4. Confirm → `AssignmentManager.assign(order_ids, assigned_to, assigned_by)`
5. Treeview rows refresh; checkboxes clear

Uses `OrderTreeview.get_checked_orders()` (already exists) → feeds into `AssignmentManager`.

### Single-order assignment (secondary)
Right-click context menu → **"Assign to…"** / **"Remove Assignment"** (admin only)

### Display in treeview
- New **"Assigned"** column (narrow) — shows assignee first name, or empty if unassigned
- Assigned rows get a treeview tag (light blue background)

### Filtering
Filter dropdown near the search bar — `All Orders | My Orders | Unassigned`
- Applied client-side; resets to `All Orders` on each fetch

### Auto-clearing
- Marked as Sent/Dispatched → `AssignmentManager.clear_on_dispatch(order_id)`

---

## 10. Files Created / Modified

### New files
| File | Purpose |
|---|---|
| `src/user_manager.py` | All user/session logic — no GUI |
| `src/assignment_manager.py` | Order assignment CRUD — no GUI |
| `src/gui/login_frame.py` | Login UI (CTkFrame) |
| `src/gui/first_run_frame.py` | First-admin setup UI (CTkFrame) |
| `src/gui/settings_window.py` | Settings window (CTkToplevel) |
| `src/gui/settings/__init__.py` | Package init |
| `src/gui/settings/user_management_view.py` | User list + add/edit UI |

### Modified files
| File | Change |
|---|---|
| `src/gui/app.py` | Login flow, `_current_user`, heartbeat thread, `WM_DELETE_WINDOW`, pass user context |
| `src/gui/home_window.py` | Settings button, `on_settings` + `current_user` params |
| `src/gui/daily_ops/daily_ops_window.py` | Accept user context, `WM_DELETE_WINDOW` clears processing flag |
| `src/gui/daily_ops/results_view.py` | Conflict check, processing flag, assignment column, filter, bulk assign button |
| `src/gui/results_tab.py` | Same as above |
| `src/gui/order_detail_view.py` | `on_processing_end` callback |

---

## 11. Deferred / Future

- **Neto API key per user**: stored in schema now; wired up when Neto action tracking is built
- **Order assignment by admin to specific user before dispatch**: tracked, not in scope yet
- **Password complexity rules**: none for now
- **Session expiry / idle timeout**: not in scope for this phase
- **Audit log**: who did what, when — future phase

---

## 12. Implementation Progress

- [ ] `docs/user_system.md` created ✓
- [ ] `src/user_manager.py`
- [ ] `src/assignment_manager.py`
- [ ] `src/gui/login_frame.py`
- [ ] `src/gui/first_run_frame.py`
- [ ] Login wired into `src/gui/app.py`
- [ ] `src/gui/settings_window.py`
- [ ] `src/gui/settings/user_management_view.py`
- [ ] Order conflict detection in results screens
- [ ] Order assignment UI (bulk + single, filter, display)
