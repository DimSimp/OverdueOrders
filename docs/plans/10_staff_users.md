# Plan 10 — Staff & User Management

> **Part of**: [Master Plan](00_overview.md)
> **Status**: 🔲 Not started (extension of existing system)
> **Phase**: Any (can be built independently)

---

## Overview

A staff management system already exists in the app. This plan describes what is already implemented, and what needs to be extended to support the new modules.

**Existing implementation is production-ready** — no rebuild needed. The work here is:
1. Migrating user storage from JSON files to Supabase
2. Adding a permission enforcement layer across all new modules
3. Adding the admin override mechanism for sensitive POS actions

---

## What Already Exists

| Component | File | Status |
|-----------|------|--------|
| User CRUD, auth, sessions | `src/user_manager.py` | ✓ Done |
| Order assignment tracking | `src/assignment_manager.py` | ✓ Done |
| Login screen (username + password) | `src/gui/login_frame.py` | ✓ Done |
| First-run admin account setup | `src/gui/first_run_frame.py` | ✓ Done |
| Home screen (workflow selection) | `src/gui/home_window.py` | ✓ Done |
| Settings window with sidebar | `src/gui/settings_window.py` | ✓ Done |
| User management view (admin) | `src/gui/settings/user_management_view.py` | ✓ Done |

### Existing User Record Structure

```json
{
  "user_id": "uuid",
  "username": "john.doe",
  "first_name": "John",
  "last_name": "Doe",
  "password_hash": "pbkdf2:sha256:...",
  "role": "admin" | "user",
  "active": true,
  "created_at": "2024-01-15T10:30:45"
}
```

**Storage**: `\\SERVER\Project Folder\Order-Fulfillment-App\Config\users.json` with local fallback at `data/users.json`.

### Existing Auth & Session Features

- **Password hashing**: PBKDF2-HMAC-SHA256, 200,000 iterations, constant-time comparison
- **Multi-device conflict**: Detects if user is already logged in on another device; offers force-login
- **Heartbeat**: Background thread every 10 seconds; sessions stale after 3 minutes without heartbeat
- **Session tracking**: `sessions` block in `users.json` tracks device, heartbeat, and currently-processing order
- **Roles**: `admin` (full access) and `user` (limited) — already two-role model matching plan requirements

### Existing UI Features

- Username dropdown + password field login
- First-run admin account creation wizard
- Settings window with sidebar nav (currently "Users" section only)
- Admin: full user list with New / Edit / Deactivate / Reset Password
- Non-admin: read-only profile view, can edit own name and password

---

## What Needs to Be Added

### 1. Migrate User Storage to Supabase

Currently stored as a JSON file on the server. As the rest of the system moves to Supabase, users should live there too — so the sync script, new modules, and any future integrations can reference a single `performed_by` user source.

**Migration steps**:
- Add `users` table to Supabase (schema matches existing JSON structure)
- Update `UserManager` to read/write Supabase instead of the JSON file
- Keep local JSON as an emergency fallback (read-only, used if Supabase is unreachable at login)
- Session heartbeats continue to use local state (in-memory) — no need to write heartbeats to Supabase

**`users` table** (Supabase):

| Column | Type | Notes |
|--------|------|-------|
| `id` | uuid PK | Replaces `user_id` |
| `username` | text UNIQUE NOT NULL | |
| `first_name` | text NOT NULL | |
| `last_name` | text NOT NULL | |
| `password_hash` | text NOT NULL | PBKDF2-SHA256, existing format preserved |
| `role` | text NOT NULL DEFAULT `'user'` | `admin` or `user` |
| `is_active` | boolean DEFAULT true | Replaces `active` |
| `created_at` | timestamptz DEFAULT now() | |
| `created_by` | text | Username of admin who created the account |
| `last_login_at` | timestamptz | |

> **Role name**: The existing codebase uses `"user"` (not `"staff"`). Keep `"user"` to avoid a rename refactor — it means the same thing.

---

### 2. Permission Enforcement Across New Modules

The existing system has role-based UI filtering for the Settings window only. All new modules need the same pattern applied systematically.

The existing `_is_admin` check pattern (`current_user.get("role") == "admin"`) should be wrapped in a shared helper in `src/user_manager.py`:

```python
def is_admin(self) -> bool:
    return self.current_user.get("role") == "admin"

def require_admin(self, action: str, description: str, reference_id: str = "") -> bool:
    """
    Shows an inline admin password prompt if current user is not admin.
    Logs the override to admin_overrides table on success.
    Returns True if authorised, False if denied/cancelled.
    """
    if self.is_admin():
        return True
    return self._show_admin_override_dialog(action, description, reference_id)
```

**Permission matrix** (new modules — existing modules already handled):

| Feature | user | admin |
|---------|------|-------|
| View inventory (qty/availability) | ✓ | ✓ |
| View cost prices | ✗ | ✓ |
| Edit item details | ✗ | ✓ |
| Manual stock adjustment | ✗ | ✓ |
| Stocktake (count entry) | ✓ | ✓ |
| Stocktake zero-out (destructive) | ✗ | ✓ |
| View/create/edit customers | ✓ | ✓ |
| Customer audit trail | ✗ | ✓ |
| Customer data export | ✗ | ✓ |
| Dispatch online orders | ✓ | ✓ |
| Cancel order / process return | Override | ✓ |
| Send replacement / wrong item | Override | ✓ |
| View/create purchase orders | ✓ | ✓ |
| Delete / void a PO | ✗ | ✓ |
| Edit supplier details | ✗ | ✓ |
| Create/delete suppliers | ✗ | ✓ |
| Create/update repairs | ✓ | ✓ |
| Delete a repair record | ✗ | ✓ |
| POS — process sales | ✓ | ✓ |
| POS — sell below minimum sell | Override | ✓ |
| POS — apply excess discount | Override | ✓ |
| POS — issue refund | ✗ | ✓ |
| Daily Z-report | ✓ | ✓ |
| Operational reports (repairs, on hold, etc.) | ✓ | ✓ |
| Financial reports (invoices, GST, etc.) | ✗ | ✓ |
| Settings | ✗ | ✓ |
| Data import wizard | ✗ | ✓ |

**Key**: ✓ = allowed, ✗ = blocked, Override = allowed via admin password prompt

---

### 3. Admin Override Mechanism

When a `user`-role staff member attempts an Override action, a small modal appears inline (without navigating away):

```
┌──────────────────────────────────────┐
│  Admin authorisation required        │
│                                      │
│  Selling GTRSTR01 at $45.00          │
│  (below minimum sell of $55.00)      │
│                                      │
│  Admin password:  [____________]     │
│                                      │
│  [Cancel]              [Authorise]   │
└──────────────────────────────────────┘
```

- Any active admin's username + password is accepted
- On success: the action proceeds; an `admin_overrides` log entry is written
- On 3 consecutive failures: access denied, action cancelled (count resets after 60 seconds)
- The non-admin user's session continues unchanged

### `admin_overrides` Table (Supabase)

| Column | Type | Notes |
|--------|------|-------|
| `id` | uuid PK | |
| `action` | text | e.g. `below_minimum_sell`, `excess_discount`, `cancel_order` |
| `description` | text | Human-readable description |
| `authorised_by` | text | Admin username |
| `requested_by` | text | Staff username who triggered the prompt |
| `reference_id` | text | Order ID, item SKU, etc. |
| `created_at` | timestamptz DEFAULT now() | |

---

### 4. Settings Window — New Sections

The existing Settings window already has the sidebar nav pattern. New sections will be added as new modules are built:

| Section | Who can see it | Already exists? |
|---------|---------------|-----------------|
| Users | Admin only | ✓ Already built |
| Suppliers | Admin only | New (plan 03) |
| Discounts | Admin only | New (plan 05) |
| SMS (TextMagic) | Admin only | New (plan 06) |
| Store Details | Admin only | New |
| App Preferences | All | New |
| Data Import | Admin only | New (plan 09) |

No structural changes to the Settings window needed — just add new `_NAV_ITEMS` entries and corresponding view classes following the existing pattern.

---

## Implementation Checklist

### Supabase Migration
- [ ] Create `users` table in Supabase (schema above)
- [ ] Create `admin_overrides` table in Supabase
- [ ] Update `UserManager` to use Supabase as primary storage
- [ ] Keep `data/users.json` as read-only emergency fallback at login
- [ ] One-time migration: read existing `users.json` → insert to Supabase

### Permission Helper
- [ ] Add `is_admin()` method to `UserManager` (or a shared `Session` helper)
- [ ] Add `require_admin(action, description, reference_id)` method
- [ ] Admin override dialog UI (inline modal, username + password entry, 3-attempt lockout)
- [ ] Write `admin_overrides` record on successful override

### Permission Enforcement in New Modules
- [ ] Inventory: hide cost price columns from `user` role
- [ ] Inventory: require admin for edit, manual adjustment, zero-out
- [ ] Customers: hide audit trail tab from `user` role; block data export
- [ ] Daily Ops: require admin override for cancel order, process return, replacement, wrong item
- [ ] Purchasing: require admin for PO void/delete
- [ ] Suppliers: require admin for edit/create/delete
- [ ] POS: require admin override for below-minimum sell and excess discount; block refund for `user`
- [ ] Reporting: hide financial reports from `user` role
- [ ] Settings: already admin-only; add new sections as modules are built
- [ ] Data import wizard: admin-only

### Settings Window Extensions
- [ ] Add Suppliers settings section (plan 03)
- [ ] Add Discounts settings section (plan 05)
- [ ] Add SMS/TextMagic settings section (plan 06)
- [ ] Add Store Details section (name, address, phone, ABN, email — used in PO PDFs and Z-report)
- [ ] Add App Preferences section (visible to all — display, date format, etc.)

---

## What Does NOT Need to Change

- Login flow (username dropdown + password) — keep as-is
- First-run wizard — keep as-is
- Heartbeat / multi-device session management — keep as-is
- Password hashing algorithm — keep PBKDF2-SHA256
- User CRUD UI in Settings — keep as-is (already complete)
- Order assignment tracking (`assignment_manager.py`) — keep as-is

---

*Last updated: 2026-04-14 — revised to reflect existing implementation; reduced scope to migration + extensions only*
