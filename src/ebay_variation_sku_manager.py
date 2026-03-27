from __future__ import annotations

import csv
import logging
from pathlib import Path

log = logging.getLogger(__name__)

_FIELDNAMES = ["item_id", "variation_specifics", "neto_sku"]


class EbayVariationSkuManager:
    """
    Manages a CSV cache that maps eBay variation specifics strings to real Neto SKUs.

    CSV columns:
        item_id             — eBay legacy item ID (e.g. "382608589068")
        variation_specifics — raw string from Fulfillment API (e.g. "Size: Medium Light ·String: 3G - .022")
        neto_sku            — resolved Neto/eBay SKU, or "" if tried but none found in the listing

    Lookup keys are normalised (stripped + lowercased) for comparison so minor
    whitespace differences don't cause cache misses.

    All reads load fresh from disk so changes made by another process are picked up
    without restarting the app.
    """

    def __init__(self, csv_path: str):
        self._path = Path(csv_path) if csv_path else None

    # ── Internal helpers ──────────────────────────────────────────────────

    @staticmethod
    def _key(item_id: str, specifics: str) -> tuple[str, str]:
        return item_id.strip(), specifics.strip().lower()

    def _load(self) -> dict[tuple[str, str], str]:
        """Return {(item_id, specifics_lower): neto_sku}."""
        if not self._path:
            return {}
        try:
            if not self._path.exists():
                return {}
            with open(self._path, newline="", encoding="utf-8") as f:
                reader = csv.DictReader(f)
                result: dict[tuple[str, str], str] = {}
                for row in reader:
                    item_id = row.get("item_id", "").strip()
                    specs = row.get("variation_specifics", "").strip()
                    sku = row.get("neto_sku", "").strip()
                    if item_id and specs:
                        result[(item_id, specs.lower())] = sku
            return result
        except Exception as exc:
            log.warning("EbayVariationSkuManager: failed to load %s: %s", self._path, exc)
            return {}

    def _write(self, data: dict[tuple[str, str], str]) -> None:
        if not self._path:
            return
        try:
            self._path.parent.mkdir(parents=True, exist_ok=True)
            with open(self._path, "w", newline="", encoding="utf-8") as f:
                writer = csv.DictWriter(f, fieldnames=_FIELDNAMES)
                writer.writeheader()
                for (item_id, specs_lower), sku in sorted(data.items()):
                    writer.writerow({
                        "item_id": item_id,
                        "variation_specifics": specs_lower,
                        "neto_sku": sku,
                    })
        except Exception as exc:
            log.error("EbayVariationSkuManager: failed to write %s: %s", self._path, exc)
            raise

    # ── Public API ────────────────────────────────────────────────────────

    def lookup(self, item_id: str, variation_specifics: str) -> str | None:
        """
        Return the mapped neto_sku, or None if this pair has never been cached.

        A return value of "" means the pair IS in the cache but no SKU was found
        in the eBay listing (prevents repeated futile API calls).
        """
        key = self._key(item_id, variation_specifics)
        data = self._load()
        return data.get(key)  # None if not present

    def save(self, item_id: str, variation_specifics: str, neto_sku: str) -> None:
        """
        Cache the resolved SKU for this (item_id, variation_specifics) pair.
        Pass neto_sku="" to record that the listing has no SKU set.
        """
        key = self._key(item_id, variation_specifics)
        data = self._load()
        data[key] = neto_sku.strip()
        self._write(data)
        log.debug(
            "EbayVariationSkuManager: saved item=%s specifics=%r → %r",
            item_id, variation_specifics, neto_sku,
        )

    def get_all(self) -> list[tuple[str, str, str]]:
        """Return all entries as [(item_id, variation_specifics, neto_sku), ...]."""
        data = self._load()
        return [
            (item_id, specs, sku)
            for (item_id, specs), sku in sorted(data.items())
        ]