from __future__ import annotations

import csv
import logging
from pathlib import Path
from typing import Optional

log = logging.getLogger(__name__)


class MusiposClient:
    """
    SQL Server client for the Musipos POS database.

    Handles SKU resolution (3-strategy cascade), PO creation/update,
    and the growing musipos_sku_map.csv alias file.
    """

    def __init__(self, config):
        """config: MusiposConfig instance"""
        self._cfg = config
        self._conn_str = (
            f"Driver={{{config.driver}}};"
            f"Server={config.server};"
            f"Database={config.database};"
            f"UID={config.user};"
            f"PWD={config.password};"
        )

    # ── Connection ────────────────────────────────────────────────────────

    def get_connection(self):
        """Return a new pyodbc connection. Caller must close it."""
        import pyodbc
        return pyodbc.connect(self._conn_str, timeout=10)

    def test_connection(self) -> tuple[bool, str]:
        """Execute SELECT 1; returns (True, '') or (False, error_message)."""
        try:
            conn = self.get_connection()
            conn.execute("SELECT 1")
            conn.close()
            return (True, "")
        except Exception as exc:
            return (False, str(exc))

    # ── SKU Resolution ────────────────────────────────────────────────────

    def resolve_item(self, neto_sku: str, suppliers: list) -> Optional[dict]:
        """
        Multi-strategy SKU resolution.

        Strategy 1: Strip supplier suffix → cascade DB lookup
                    (itm_iid → itm_supplier_iid → itm_barcode)
        Strategy 2: Kit/invoice mapping from sku_mappings.csv
        Strategy 3: Manual alias from musipos_sku_map.csv

        Returns {itm_iid, title, supplier_id, supplier_iid, qty_on_hand,
                 last_cost, retail_price} or None.
        """
        # Strategy 1
        base_sku = self._strip_suffix(neto_sku, suppliers)
        result = self._cascade_lookup(base_sku)
        if result:
            return result

        # Strategy 2 — kit/invoice expansion
        kit_map = self.load_kit_mappings()
        invoice_skus = kit_map.get(neto_sku.upper().strip(), [])
        for inv_sku in invoice_skus:
            result = self._cascade_lookup(inv_sku, columns=("itm_iid", "itm_supplier_iid"))
            if result:
                return result

        # Strategy 3 — manual alias CSV
        alias_map = self.load_musipos_map()
        musipos_sku = alias_map.get(neto_sku.upper().strip())
        if musipos_sku:
            return self.resolve_item_by_musipos_sku(musipos_sku)

        return None

    def resolve_item_by_musipos_sku(self, musipos_sku: str) -> Optional[dict]:
        """Direct lookup by itm_iid. Used for manual entry and Strategy 3."""
        return self._cascade_lookup(musipos_sku, columns=("itm_iid",))

    def _strip_suffix(self, neto_sku: str, suppliers: list) -> str:
        """Return the base SKU with any matching supplier suffix removed."""
        for supplier in suppliers:
            suffix = supplier.suffix
            if not suffix:
                continue
            pos = getattr(supplier, "suffix_position", "append")
            if pos == "append" and neto_sku.upper().endswith(suffix.upper()):
                base = neto_sku[: -len(suffix)]
                # Re-apply any character substitutions in reverse (best-effort)
                return base
            elif pos == "prepend" and neto_sku.upper().startswith(suffix.upper()):
                return neto_sku[len(suffix):]
        return neto_sku

    _ITEM_SELECT = (
        "SELECT TOP 1 "
        "ap4itm.itm_iid, ap4itm.itm_title, ap4itm.itm_supplier_id, "
        "ap4itm.itm_supplier_iid, sp4qpc.qpc_qty_on_hand, "
        "sp4qpc.qpc_last_purchase_cost, ap4itm.itm_new_retail_price, "
        "sp4qpc.qpc_qor "
        "FROM ap4itm "
        "JOIN sp4qpc ON RTRIM(ap4itm.itm_iid) = RTRIM(sp4qpc.qpc_iid) "
        "WHERE ap4itm.itm_lno = '0000' "
        "  AND sp4qpc.qpc_warehouse_id = '00001' "
        "  AND UPPER(RTRIM(ap4itm.{col})) = UPPER(?)"
    )

    # Same query without TOP 1, plus supplier name for disambiguation display.
    _ITEM_SELECT_MULTI = (
        "SELECT "
        "ap4itm.itm_iid, ap4itm.itm_title, ap4itm.itm_supplier_id, "
        "ap4itm.itm_supplier_iid, sp4qpc.qpc_qty_on_hand, "
        "sp4qpc.qpc_last_purchase_cost, ap4itm.itm_new_retail_price, "
        "sp4qpc.qpc_qor, "
        "ISNULL(RTRIM(ap4rsp.rsp_name), ap4itm.itm_supplier_id) "
        "FROM ap4itm "
        "JOIN sp4qpc ON RTRIM(ap4itm.itm_iid) = RTRIM(sp4qpc.qpc_iid) "
        "LEFT JOIN ap4rsp ON ap4rsp.rsp_supplier_id = ap4itm.itm_supplier_id "
        "WHERE ap4itm.itm_lno = '0000' "
        "  AND sp4qpc.qpc_warehouse_id = '00001' "
        "  AND UPPER(RTRIM(ap4itm.{col})) = UPPER(?)"
    )

    def _cascade_lookup(
        self, sku: str, columns: tuple = ("itm_iid", "itm_supplier_iid", "itm_barcode")
    ) -> Optional[dict]:
        if not sku:
            return None
        try:
            conn = self.get_connection()
            try:
                cur = conn.cursor()
                for col in columns:
                    sql = self._ITEM_SELECT.format(col=col)
                    cur.execute(sql, (sku.strip(),))
                    row = cur.fetchone()
                    if row:
                        return {
                            "itm_iid": (row[0] or "").strip(),
                            "title": (row[1] or "").strip(),
                            "supplier_id": (row[2] or "").strip(),
                            "supplier_iid": (row[3] or "").strip(),
                            "qty_on_hand": int(row[4] or 0),
                            "last_cost": float(row[5] or 0),
                            "retail_price": float(row[6] or 0),
                            "qty_on_order": int(row[7] or 0),
                        }
            finally:
                conn.close()
        except Exception as exc:
            log.warning("Musipos cascade lookup failed for %r: %s", sku, exc)
        return None

    def _multi_lookup(self, sku: str, col: str) -> list[dict]:
        """Return ALL matching items for a column/value (no TOP 1 limit)."""
        if not sku:
            return []
        results = []
        try:
            conn = self.get_connection()
            try:
                cur = conn.cursor()
                sql = self._ITEM_SELECT_MULTI.format(col=col)
                cur.execute(sql, (sku.strip(),))
                for row in cur.fetchall():
                    results.append({
                        "itm_iid": (row[0] or "").strip(),
                        "title": (row[1] or "").strip(),
                        "supplier_id": (row[2] or "").strip(),
                        "supplier_iid": (row[3] or "").strip(),
                        "qty_on_hand": int(row[4] or 0),
                        "last_cost": float(row[5] or 0),
                        "retail_price": float(row[6] or 0),
                        "qty_on_order": int(row[7] or 0),
                        "supplier_name": (row[8] or "").strip(),
                    })
            finally:
                conn.close()
        except Exception as exc:
            log.warning("Musipos multi lookup failed for %r col=%s: %s", sku, col, exc)
        return results

    def resolve_item_multi(self, neto_sku: str, suppliers: list) -> list[dict]:
        """
        Multi-strategy SKU resolution returning ALL matches.

        Returns [] (not found), [item] (unique), or [item, ...] (ambiguous).
        Strategy 1: itm_iid (TOP 1 — unique), then itm_supplier_iid / itm_barcode (multi)
        Strategy 2: Kit/invoice map → itm_iid lookup
        Strategy 3: Manual alias CSV → itm_iid lookup
        """
        base_sku = self._strip_suffix(neto_sku, suppliers)

        # itm_iid is a primary key — always unique; keep TOP 1 lookup
        result = self._cascade_lookup(base_sku, columns=("itm_iid",))
        if result:
            return [result]

        # itm_supplier_iid may match multiple items from different suppliers
        matches = self._multi_lookup(base_sku, "itm_supplier_iid")
        if matches:
            return matches

        matches = self._multi_lookup(base_sku, "itm_barcode")
        if matches:
            return matches

        # Strategy 2 — kit/invoice expansion (resolves to itm_iid, always unique)
        kit_map = self.load_kit_mappings()
        invoice_skus = kit_map.get(neto_sku.upper().strip(), [])
        for inv_sku in invoice_skus:
            result = self._cascade_lookup(inv_sku, columns=("itm_iid", "itm_supplier_iid"))
            if result:
                return [result]

        # Strategy 3 — manual alias CSV
        alias_map = self.load_musipos_map()
        musipos_sku = alias_map.get(neto_sku.upper().strip())
        if musipos_sku:
            result = self.resolve_item_by_musipos_sku(musipos_sku)
            if result:
                return [result]

        return []

    def resolve_manual_multi(self, sku: str) -> list[dict]:
        """
        Manual SKU search returning ALL matches.

        Tries itm_iid first (unique primary key), then itm_supplier_iid,
        then itm_barcode — each may return multiple rows.
        """
        result = self._cascade_lookup(sku, columns=("itm_iid",))
        if result:
            return [result]

        matches = self._multi_lookup(sku, "itm_supplier_iid")
        if matches:
            return matches

        matches = self._multi_lookup(sku, "itm_barcode")
        if matches:
            return matches

        return []

    # ── Supplier info ─────────────────────────────────────────────────────

    def get_suppliers_for_item(self, itm_iid: str) -> list[dict]:
        """
        Return [{supplier_id, supplier_name}, ...].
        Includes primary and alternate suppliers.
        """
        sql = (
            "SELECT DISTINCT s.rsp_supplier_id, RTRIM(s.rsp_name) "
            "FROM ap4itm i "
            "JOIN ap4rsp s ON s.rsp_supplier_id = i.itm_supplier_id "
            "WHERE RTRIM(i.itm_iid) = ? "
            "UNION "
            "SELECT DISTINCT s2.rsp_supplier_id, RTRIM(s2.rsp_name) "
            "FROM sp4qpc q "
            "JOIN ap4rsp s2 ON s2.rsp_supplier_id = q.qpc_alt_supplier_id "
            "WHERE RTRIM(q.qpc_iid) = ? AND q.qpc_warehouse_id = '00001' "
            "  AND q.qpc_alt_supplier_id IS NOT NULL AND RTRIM(q.qpc_alt_supplier_id) <> ''"
        )
        results = []
        try:
            conn = self.get_connection()
            try:
                cur = conn.cursor()
                cur.execute(sql, (itm_iid.strip(), itm_iid.strip()))
                for row in cur.fetchall():
                    results.append({"supplier_id": (row[0] or "").strip(),
                                    "supplier_name": (row[1] or "").strip()})
            finally:
                conn.close()
        except Exception as exc:
            log.warning("get_suppliers_for_item failed for %r: %s", itm_iid, exc)
        return results

    # ── PO lookup ─────────────────────────────────────────────────────────

    def get_current_po(self, supplier_id: str) -> Optional[int]:
        """Return the PO number of the current open PO for supplier, or None."""
        sql = (
            "SELECT TOP 1 phd_po_no FROM sp4phd "
            "WHERE phd_supplier_id = ? AND phd_po_status = 'CURRENT' "
            "ORDER BY phd_po_no DESC"
        )
        try:
            conn = self.get_connection()
            try:
                cur = conn.cursor()
                cur.execute(sql, (supplier_id,))
                row = cur.fetchone()
                return int(row[0]) if row else None
            finally:
                conn.close()
        except Exception as exc:
            log.warning("get_current_po failed for %r: %s", supplier_id, exc)
            return None

    # ── PO creation ───────────────────────────────────────────────────────

    def add_to_po(
        self,
        itm_iid: str,
        supplier_id: str,
        qty: int,
        item_dict: Optional[dict] = None,
        dry_run: bool = True,
    ) -> dict:
        """
        Add or update a line in the supplier's CURRENT PO.

        Returns {po_no, new_po, supplier_id, supplier_name, action, qty_added}.
        Raises on DB error.
        """
        import pyodbc

        item_dict = item_dict or {}
        last_cost = item_dict.get("last_cost", 0.0)
        retail_price = item_dict.get("retail_price", 0.0)
        title = item_dict.get("title", "")

        conn = self.get_connection()
        try:
            conn.autocommit = False

            cur = conn.cursor()

            # --- 1. Find or allocate PO number ---
            cur.execute(
                "SELECT TOP 1 phd_po_no FROM sp4phd "
                "WHERE phd_supplier_id = ? AND phd_po_status = 'CURRENT' "
                "ORDER BY phd_po_no DESC",
                (supplier_id,),
            )
            row = cur.fetchone()
            new_po = row is None
            if new_po:
                # Allocate new PO number from supplier record (with lock)
                cur.execute(
                    "SELECT rsp_curr_po_no FROM ap4rsp WITH (UPDLOCK, HOLDLOCK) "
                    "WHERE rsp_supplier_id = ?",
                    (supplier_id,),
                )
                rsp_row = cur.fetchone()
                if not rsp_row:
                    raise ValueError(f"Supplier '{supplier_id}' not found in ap4rsp")
                po_no = int(rsp_row[0] or 0) + 1
                cur.execute(
                    "UPDATE ap4rsp SET rsp_curr_po_no = ? WHERE rsp_supplier_id = ?",
                    (po_no, supplier_id),
                )
                cur.execute(
                    "INSERT INTO sp4phd "
                    "(phd_po_no, phd_supplier_id, phd_po_status, phd_po_dest, phd_memo) "
                    "VALUES (?, ?, 'CURRENT', 'S', 'AIO ordering')",
                    (po_no, supplier_id),
                )
            else:
                po_no = int(row[0])

            # --- 2. Check if item already on this PO ---
            cur.execute(
                "SELECT pop_qor FROM sp4pop "
                "WHERE RTRIM(pop_iid) = ? AND pop_supplier_id = ? AND pop_po_no = ?",
                (itm_iid.strip(), supplier_id, po_no),
            )
            existing = cur.fetchone()

            if existing:
                # Update quantity on existing line
                cur.execute(
                    "UPDATE sp4pop SET pop_qor = pop_qor + ? "
                    "WHERE RTRIM(pop_iid) = ? AND pop_supplier_id = ? AND pop_po_no = ?",
                    (qty, itm_iid.strip(), supplier_id, po_no),
                )
                action = "updated"
            else:
                # Determine next line number for this PO
                cur.execute(
                    "SELECT ISNULL(MAX(pop_lno), 0) FROM sp4pop "
                    "WHERE pop_supplier_id = ? AND pop_po_no = ?",
                    (supplier_id, po_no),
                )
                max_lno_row = cur.fetchone()
                next_lno = int(max_lno_row[0] or 0) + 1

                supplier_iid = item_dict.get("supplier_iid", "") or ""

                cur.execute(
                    "INSERT INTO sp4pop "
                    "(pop_iid, pop_supplier_id, pop_po_no, pop_cid, pop_lno, "
                    " pop_po_date, pop_supplier_iid, pop_qor, "
                    " pop_curr_cost, pop_curr_order_flag, "
                    " pop_qty_rcv, pop_curr_price, pop_total_cost, "
                    " pop_itm_status, pop_user_id, pop_computer_id, "
                    " pop_comment, pop_qbo, pop_tax_amt, "
                    " pop_order_no, pop_process_status, pop_itm_title) "
                    "VALUES (?, ?, ?, 'STOCK001', ?, GETDATE(), ?, ?, "
                    "        ?, 'N', "
                    "        0, ?, 0.00, "
                    "        'C', ?, ?, "
                    "        'AIO ordering', 0, 0.00, "
                    "        '', 'N', ?)",
                    (
                        itm_iid.strip(), supplier_id, po_no, next_lno,
                        supplier_iid, qty,
                        last_cost,
                        retail_price,
                        self._cfg.default_user_id, self._cfg.computer_id,
                        title,
                    ),
                )
                action = "added"

            # --- 3. Update qty on order in inventory ---
            cur.execute(
                "UPDATE sp4qpc SET qpc_qor = ISNULL(qpc_qor, 0) + ? "
                "WHERE qpc_iid = ? AND qpc_warehouse_id = ?",
                (qty, itm_iid.strip(), self._cfg.warehouse_id),
            )

            # Get supplier name for return value
            cur.execute(
                "SELECT RTRIM(rsp_name) FROM ap4rsp WHERE rsp_supplier_id = ?", (supplier_id,)
            )
            name_row = cur.fetchone()
            supplier_name = (name_row[0] if name_row else supplier_id) or supplier_id

            if dry_run:
                conn.rollback()
            else:
                conn.commit()

            return {
                "po_no": po_no,
                "new_po": new_po,
                "supplier_id": supplier_id,
                "supplier_name": supplier_name,
                "action": action,
                "qty_added": qty,
            }
        except Exception:
            try:
                conn.rollback()
            except Exception:
                pass
            raise
        finally:
            conn.close()

    # ── CSV helpers ───────────────────────────────────────────────────────

    def load_kit_mappings(self) -> dict[str, list[str]]:
        """
        Read sku_mappings.csv (kit/invoice expansion map).
        Returns {neto_sku_upper: [invoice_sku, ...]}.
        """
        path = Path(self._cfg.kit_mappings_path)
        if not path.exists():
            log.warning("kit_mappings_path not found: %s", path)
            return {}
        result: dict[str, list[str]] = {}
        try:
            with open(path, newline="", encoding="utf-8-sig") as f:
                reader = csv.DictReader(f)
                for row in reader:
                    neto = (row.get("neto_sku") or "").strip().upper()
                    if not neto:
                        continue
                    raw_inv = (row.get("invoice_skus") or "").strip()
                    invoice_skus = [s.strip() for s in raw_inv.split("|") if s.strip()]
                    if invoice_skus:
                        result[neto] = invoice_skus
        except Exception as exc:
            log.warning("Failed to load kit mappings from %s: %s", path, exc)
        return result

    def load_musipos_map(self) -> dict[str, str]:
        """
        Read musipos_sku_map.csv.
        Returns {neto_sku_upper: musipos_sku}.
        """
        path = Path(self._cfg.musipos_map_path)
        if not path.exists():
            return {}
        result: dict[str, str] = {}
        try:
            with open(path, newline="", encoding="utf-8-sig") as f:
                reader = csv.DictReader(f)
                for row in reader:
                    neto = (row.get("neto_sku") or "").strip().upper()
                    musipos = (row.get("musipos_sku") or "").strip()
                    if neto and musipos:
                        result[neto] = musipos
        except Exception as exc:
            log.warning("Failed to load musipos map from %s: %s", path, exc)
        return result

    def save_musipos_alias(self, neto_sku: str, musipos_sku: str) -> None:
        """Append a new row to musipos_sku_map.csv, creating it if needed."""
        path = Path(self._cfg.musipos_map_path)
        write_header = not path.exists()
        try:
            path.parent.mkdir(parents=True, exist_ok=True)
            with open(path, "a", newline="", encoding="utf-8-sig") as f:
                writer = csv.writer(f)
                if write_header:
                    writer.writerow(["neto_sku", "musipos_sku"])
                writer.writerow([neto_sku.upper().strip(), musipos_sku.strip()])
        except Exception as exc:
            log.error("Failed to save musipos alias (%s → %s): %s", neto_sku, musipos_sku, exc)
            raise
