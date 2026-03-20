from __future__ import annotations

import csv
import logging
from pathlib import Path

log = logging.getLogger(__name__)

_FIELDNAMES = ["neto_sku", "is_kit", "invoice_skus", "supplier"]


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
                    result[key] = {
                        "is_kit": row.get("is_kit", "False").strip() == "True",
                        "invoice_skus": inv_skus,
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
                    writer.writerow({
                        "neto_sku": neto_sku,
                        "is_kit": str(mapping["is_kit"]),
                        "invoice_skus": "|".join(mapping["invoice_skus"]),
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

    def save(self, neto_sku: str, invoice_skus: list[str], is_kit: bool, supplier: str = "") -> None:
        """Add or update the mapping for neto_sku. Writes the full CSV."""
        key = neto_sku.upper().strip()
        if not key:
            return
        data = self._load()
        data[key] = {
            "is_kit": is_kit,
            "invoice_skus": [s.strip() for s in invoice_skus if s.strip()],
            "supplier": supplier.strip(),
        }
        self._write(data)
        log.debug("SkuAliasManager: saved %s → %s (kit=%s, supplier=%s)", key, invoice_skus, is_kit, supplier)

    def remove(self, neto_sku: str) -> None:
        """Remove the mapping for neto_sku. No-op if not present."""
        key = neto_sku.upper().strip()
        data = self._load()
        if key in data:
            del data[key]
            self._write(data)
            log.debug("SkuAliasManager: removed %s", key)
