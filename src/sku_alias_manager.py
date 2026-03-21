from __future__ import annotations

import csv
import logging
from pathlib import Path

log = logging.getLogger(__name__)

_FIELDNAMES = ["neto_sku", "is_kit", "invoice_skus", "qty_per_alias", "supplier"]


class SkuAliasManager:
    """
    Manages SKU alias mappings stored in a CSV file.

    CSV columns:
        neto_sku     — the SKU as it appears on a Neto/eBay order line item (normalised upper)
        is_kit       — "True" / "False"
        invoice_skus — pipe-separated list of RAW supplier invoice SKUs (no suffix), e.g. "50TOR|60TOR"
        supplier     — supplier name as in config (used to apply suffix/substitutions at match time)

    The suffix and character substitutions for the supplier are NOT stored here — they are
    applied dynamically by match_orders_to_invoice() using the live SupplierConfig list.

    All reads load fresh from disk (file is tiny; avoids stale state across multiple
    processes/users on the same network share).
    """

    def __init__(self, csv_path: str):
        self._path = Path(csv_path) if csv_path else None

    # ── Internal helpers ──────────────────────────────────────────────────

    def _load(self) -> dict[str, dict]:
        """Return {neto_sku_upper: {is_kit, invoice_skus, supplier}}."""
        if not self._path:
            return {}
        try:
            if not self._path.exists():
                return {}
            with open(self._path, newline="", encoding="utf-8") as f:
                reader = csv.DictReader(f)
                result: dict[str, dict] = {}
                for row in reader:
                    key = row.get("neto_sku", "").upper().strip()
                    if not key:
                        continue
                    raw_skus = row.get("invoice_skus", "")
                    inv_skus = [s.strip() for s in raw_skus.split("|") if s.strip()]
                    raw_qtys = row.get("qty_per_alias", "")
                    qty_parts = [q.strip() for q in raw_qtys.split("|") if q.strip()]
                    qty_list = []
                    for i in range(len(inv_skus)):
                        try:
                            qty_list.append(max(1, int(qty_parts[i])))
                        except (IndexError, ValueError):
                            qty_list.append(1)
                    result[key] = {
                        "is_kit": row.get("is_kit", "False").strip() == "True",
                        "invoice_skus": inv_skus,
                        "qty_per_alias": qty_list,
                        "supplier": row.get("supplier", "").strip(),
                    }
            return result
        except Exception as exc:
            log.warning("SkuAliasManager: failed to load %s: %s", self._path, exc)
            return {}

    def _write(self, data: dict[str, dict]) -> None:
        """Write the full mapping dict back to the CSV file."""
        if not self._path:
            return
        try:
            self._path.parent.mkdir(parents=True, exist_ok=True)
            with open(self._path, "w", newline="", encoding="utf-8") as f:
                writer = csv.DictWriter(f, fieldnames=_FIELDNAMES)
                writer.writeheader()
                for neto_sku, mapping in sorted(data.items()):
                    skus = mapping["invoice_skus"]
                    qtys = mapping.get("qty_per_alias") or [1] * len(skus)
                    writer.writerow({
                        "neto_sku": neto_sku,
                        "is_kit": str(mapping["is_kit"]),
                        "invoice_skus": "|".join(skus),
                        "qty_per_alias": "|".join(str(q) for q in qtys),
                        "supplier": mapping.get("supplier", ""),
                    })
        except Exception as exc:
            log.error("SkuAliasManager: failed to write %s: %s", self._path, exc)
            raise

    # ── Public API ────────────────────────────────────────────────────────

    def get_aliases(self, neto_sku: str) -> list[str]:
        """Return list of raw invoice SKUs for a given Neto SKU. Empty list if not mapped."""
        data = self._load()
        mapping = data.get(neto_sku.upper().strip())
        return mapping["invoice_skus"] if mapping else []

    def get_all(self) -> dict[str, dict]:
        """Return all mappings as {neto_sku_upper: {is_kit, invoice_skus, supplier}}."""
        return self._load()

    def has(self, neto_sku: str) -> bool:
        """Return True if a mapping exists for this Neto SKU."""
        return neto_sku.upper().strip() in self._load()

    def save(
        self,
        neto_sku: str,
        invoice_skus: list[str],
        is_kit: bool,
        supplier: str = "",
        qty_per_alias: list[int] | None = None,
    ) -> None:
        """Add or update the mapping for neto_sku. Writes the full CSV."""
        key = neto_sku.upper().strip()
        if not key:
            return
        clean_skus = [s.strip() for s in invoice_skus if s.strip()]
        qtys = qty_per_alias or [1] * len(clean_skus)
        # Ensure qtys aligns with clean_skus length
        qtys = [max(1, q) for q in qtys[:len(clean_skus)]]
        while len(qtys) < len(clean_skus):
            qtys.append(1)
        data = self._load()
        data[key] = {
            "is_kit": is_kit,
            "invoice_skus": clean_skus,
            "qty_per_alias": qtys,
            "supplier": supplier.strip(),
        }
        self._write(data)
        log.debug("SkuAliasManager: saved %s → %s qty=%s (kit=%s, supplier=%s)", key, clean_skus, qtys, is_kit, supplier)

    def remove(self, neto_sku: str) -> None:
        """Remove the mapping for neto_sku. No-op if not present."""
        key = neto_sku.upper().strip()
        data = self._load()
        if key in data:
            del data[key]
            self._write(data)
            log.debug("SkuAliasManager: removed %s", key)

    def rename_key(self, old_sku: str, new_sku: str) -> bool:
        """Move alias mapping from old_sku key to new_sku. Returns True if a mapping was found."""
        old_key = old_sku.upper().strip()
        new_key = new_sku.upper().strip()
        data = self._load()
        if old_key not in data:
            return False
        data[new_key] = data.pop(old_key)
        self._write(data)
        log.debug("SkuAliasManager: renamed key %s → %s", old_key, new_key)
        return True
