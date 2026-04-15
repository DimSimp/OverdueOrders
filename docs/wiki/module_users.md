# Module: Staff & User Management

> **Detail plan**: [docs/plans/10_staff_users.md](../plans/10_staff_users.md)
> **Build phase**: Any (independent of other modules)
> **Tables owned**: `users`, `admin_overrides`
> **Existing code**: `src/user_manager.py`, `src/gui/login_frame.py`, `src/gui/settings_window.py`, `src/gui/settings/user_management_view.py`

---

## Overview

A production-ready staff management system already exists in the app. This module documents what
exists and the three extension tasks needed to support new modules.

---

## What Already Exists

| Component | Status |
|-----------|--------|
| PBKDF2-SHA256 password hashing (200K iterations) | ✓ Done |
| Multi-device session heartbeats (10s interval, 3-min stale) | ✓ Done |
| Username dropdown + password login screen | ✓ Done |
| First-run admin account creation wizard | ✓ Done |
| Full user CRUD in Settings (admin only) | ✓ Done |
| Two roles: `admin` and `user` | ✓ Done |
| JSON storage on server (`\\SERVER\...`) with local fallback | ✓ Done |

**Role name is `"user"` (not `"staff"`)** — kept to avoid renaming the entire existing codebase.

---

## Extension Tasks

### 1. Migrate storage to Supabase

Add `users` table (schema in [database_schema.md](database_schema.md)). Update `UserManager` to
read/write Supabase. Keep `data/users.json` as read-only emergency fallback at login.
Session heartbeats remain in-memory only — not written to Supabase.

### 2. Admin Override Mechanism

When a `user`-role staff member attempts an Override-gated action, an inline modal appears:

```
Admin authorisation required

Selling GTRSTR01 at $45.00
(below minimum sell of $55.00)

Admin password:  [____________]

[Cancel]              [Authorise]
```

- Any active admin's username + password accepted
- 3 consecutive failures → action cancelled (count resets after 60 seconds)
- Success → logs to `admin_overrides` table, action proceeds

Helper method in `UserManager`:
```python
def require_admin(self, action: str, description: str, reference_id: str = "") -> bool:
    if self.is_admin(): return True
    return self._show_admin_override_dialog(action, description, reference_id)
```

### 3. Settings Window Extensions

Existing `_NAV_ITEMS` pattern in `src/gui/settings_window.py`. Add:

| Section | Visible to | Plan |
|---------|-----------|------|
| Suppliers (SKU rules) | Admin | P03 |
| Discounts | Admin | P05 |
| SMS / TextMagic | Admin | P06 |
| Store Details (name, address, phone, ABN, email) | Admin | P04 + P02 receipts |
| App Preferences | All | — |
| Data Import | Admin | P09 |

---

## Permission Matrix (new modules)

| Feature | `user` | `admin` |
|---------|--------|---------|
| View inventory | ✓ | ✓ |
| View cost prices | ✗ | ✓ |
| Edit items / adjust stock / zero-out | ✗ | ✓ |
| View/create/edit customers | ✓ | ✓ |
| Customer audit + data export | ✗ | ✓ |
| Dispatch online orders | ✓ | ✓ |
| Cancel order / process return | Override | ✓ |
| View/create POs | ✓ | ✓ |
| Void a PO / edit suppliers | ✗ | ✓ |
| POS — sell below minimum sell | Override | ✓ |
| POS — apply excess discount | Override | ✓ |
| POS — issue refund | ✗ | ✓ |
| Z-report + operational reports | ✓ | ✓ |
| Financial reports + stock audit | ✗ | ✓ |
| Settings / Data import | ✗ | ✓ |

**Key**: ✓ = allowed, ✗ = blocked, Override = `require_admin()` inline prompt

---

## What Does NOT Need to Change

- Login flow (username dropdown + password)
- First-run wizard
- Heartbeat / multi-device session management
- Password hashing algorithm
- User CRUD UI in Settings
- Order assignment tracking (`src/assignment_manager.py`)
