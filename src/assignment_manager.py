"""Order assignment manager for the Scarlett AIO app.

Persists order-to-user assignments in a shared JSON file on the network.
No GUI code in this module.
"""
from __future__ import annotations

import json
import logging
import os
import time
from datetime import datetime
from pathlib import Path
from typing import Optional

log = logging.getLogger(__name__)

ASSIGNMENTS_FILE = Path(
    r"\\SERVER\Project Folder\Order-Fulfillment-App\Config\order_assignments.json"
)


class AssignmentManager:
    """Manage order-to-user assignments via a shared JSON file."""

    def __init__(self, path: Optional[Path] = None) -> None:
        self._path = Path(path) if path else ASSIGNMENTS_FILE

    # ── Low-level I/O ─────────────────────────────────────────────────────────

    def load(self) -> dict:
        """Read and return the full assignments data dict."""
        if not self._path.exists():
            return {"assignments": {}}
        data = json.loads(self._path.read_text(encoding="utf-8"))
        if "assignments" not in data:
            data["assignments"] = {}
        return data

    def save(self, data: dict) -> None:
        """Write *data* atomically with up to 3 retries."""
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
                log.warning("assignments write attempt %d failed: %s", attempt + 1, exc)
                time.sleep(0.2)

    # ── Assignment CRUD ───────────────────────────────────────────────────────

    def assign(
        self,
        order_ids: list[str],
        assigned_to: str,
        assigned_by: str,
    ) -> None:
        """Assign *order_ids* to *assigned_to* (username).

        Passing assigned_to="" is equivalent to calling unassign().
        """
        if not assigned_to:
            self.unassign(order_ids)
            return
        data = self.load()
        assignments = data["assignments"]
        now = datetime.now().isoformat(timespec="seconds")
        for oid in order_ids:
            assignments[oid] = {
                "assigned_to": assigned_to,
                "assigned_by": assigned_by,
                "assigned_at": now,
            }
        self.save(data)
        log.info(
            "Assigned %d order(s) to '%s' by '%s'",
            len(order_ids), assigned_to, assigned_by,
        )

    def unassign(self, order_ids: list[str]) -> None:
        """Remove assignments for *order_ids*."""
        data = self.load()
        assignments = data["assignments"]
        removed = 0
        for oid in order_ids:
            if oid in assignments:
                del assignments[oid]
                removed += 1
        if removed:
            self.save(data)
            log.info("Removed assignment for %d order(s)", removed)

    def get_assignment(self, order_id: str) -> Optional[dict]:
        """Return assignment dict for *order_id*, or None if unassigned."""
        try:
            return self.load()["assignments"].get(order_id)
        except Exception:
            return None

    def get_all_assignments(self) -> dict:
        """Return the full assignments dict.  Returns {} on error."""
        try:
            return self.load().get("assignments", {})
        except Exception:
            return {}

    def clear_on_dispatch(self, order_id: str) -> None:
        """Remove the assignment for *order_id* when it is marked as sent/dispatched."""
        try:
            self.unassign([order_id])
        except Exception as exc:
            log.warning("clear_on_dispatch(%s) failed: %s", order_id, exc)

    def get_assigned_to_username(self, order_id: str) -> Optional[str]:
        """Return the username assigned to *order_id*, or None."""
        a = self.get_assignment(order_id)
        return a["assigned_to"] if a else None
