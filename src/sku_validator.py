"""
SKU validation against inventory.CSV.

Compares raw (pre-suffix) SKUs extracted from invoices against a local inventory
CSV to catch OCR errors. Provides fuzzy/OCR-aware suggestions and persists
corrections to a CSV file for reuse in future sessions.
"""

from __future__ import annotations

import csv
from dataclasses import dataclass, field
from pathlib import Path
from typing import Optional

from src.pdf_parser import InvoiceItem


# ---------------------------------------------------------------------------
# Supplier name → Supplier_ID mapping (matches inventory.CSV Supplier_ID col)
# ---------------------------------------------------------------------------
_SUPPLIER_ID_MAP: dict[str, str] = {
    "australis music": "AUSTRAL",
    "paytons": "PAYTONS",
    "electric factory": "ELECTRI",
    "cmi": "CMI",
    "cmc music": "CMCMUS",
    "ams (australasian music supplies)": "AMS",
    "pro music australia": "PRO",
    "amber technology": "AMBER",
    "d'addario australia": "DADDAUS",
    "shriro": "",        # not in inventory
    "jade": "JADE",
}

# OCR confusion pairs — each tuple is bidirectionally substitutable
_OCR_PAIRS: list[tuple[str, str]] = [
    ("0", "O"),
    ("1", "I"),
    ("1", "l"),
    ("I", "l"),
    ("5", "S"),
    ("W", "V"),
    ("B", "8"),
    ("Z", "2"),
    ("6", "G"),
    ("4", "A"),
]

# Build lookup: char → all chars it could be confused with
_OCR_MAP: dict[str, list[str]] = {}
for _a, _b in _OCR_PAIRS:
    _OCR_MAP.setdefault(_a, []).append(_b)
    _OCR_MAP.setdefault(_b, []).append(_a)


# ---------------------------------------------------------------------------
# Inventory loading
# ---------------------------------------------------------------------------

def load_inventory(csv_path: str) -> dict[str, set[str]]:
    """
    Load inventory.CSV and return a dict of upper-cased SKU sets keyed by
    Supplier_ID (and "" for the full all-suppliers set).

    Returns {} if the file does not exist (causes validation to be skipped).
    """
    path = Path(csv_path)
    if not path.exists():
        return {}

    all_skus: set[str] = set()
    by_supplier: dict[str, set[str]] = {}

    try:
        with open(path, encoding="cp1252", newline="") as f:
            reader = csv.DictReader(f)
            for row in reader:
                sku = (row.get("Supplier_Item_ID") or "").strip().upper()
                sup_id = (row.get("Supplier_ID") or "").strip().upper()
                if not sku:
                    continue
                all_skus.add(sku)
                by_supplier.setdefault(sup_id, set()).add(sku)
    except Exception:
        return {}

    result: dict[str, set[str]] = {"": all_skus}
    result.update(by_supplier)
    return result


# ---------------------------------------------------------------------------
# Corrections CSV
# ---------------------------------------------------------------------------

def load_corrections(path: str) -> dict[tuple[str, str], str]:
    """
    Load sku_corrections.csv.

    Returns {(supplier_name_lower, raw_sku_upper): corrected_sku_upper}.
    Returns {} if the file does not exist.
    """
    p = Path(path)
    if not p.exists():
        return {}

    corrections: dict[tuple[str, str], str] = {}
    try:
        with open(p, encoding="utf-8-sig", newline="") as f:
            reader = csv.DictReader(f)
            for row in reader:
                supplier = (row.get("supplier_name") or "").strip().lower()
                raw = (row.get("raw_sku") or "").strip().upper()
                corrected = (row.get("corrected_sku") or "").strip().upper()
                if supplier and raw and corrected:
                    corrections[(supplier, raw)] = corrected
    except Exception:
        return {}

    return corrections


def save_corrections(
    path: str,
    new_corrections: list[tuple[str, str, str]],
) -> None:
    """
    Merge new_corrections into the existing corrections file and rewrite it.

    new_corrections: [(supplier_name, raw_sku, corrected_sku), ...]
    """
    existing = load_corrections(path)

    for supplier_name, raw_sku, corrected_sku in new_corrections:
        key = (supplier_name.lower(), raw_sku.upper())
        existing[key] = corrected_sku.upper()

    p = Path(path)
    with open(p, "w", encoding="utf-8-sig", newline="") as f:
        writer = csv.DictWriter(
            f, fieldnames=["supplier_name", "raw_sku", "corrected_sku"]
        )
        writer.writeheader()
        for (supplier_lower, raw_upper), corrected_upper in sorted(existing.items()):
            writer.writerow(
                {
                    "supplier_name": supplier_lower,
                    "raw_sku": raw_upper,
                    "corrected_sku": corrected_upper,
                }
            )


# ---------------------------------------------------------------------------
# Fuzzy / OCR suggestion helpers
# ---------------------------------------------------------------------------

def _levenshtein(a: str, b: str) -> int:
    """Compute Levenshtein edit distance between two strings (stdlib only)."""
    if a == b:
        return 0
    if not a:
        return len(b)
    if not b:
        return len(a)

    prev = list(range(len(b) + 1))
    for i, ca in enumerate(a, 1):
        curr = [i]
        for j, cb in enumerate(b, 1):
            curr.append(
                min(
                    prev[j] + 1,           # deletion
                    curr[j - 1] + 1,       # insertion
                    prev[j - 1] + (ca != cb),  # substitution
                )
            )
        prev = curr
    return prev[-1]


def _ocr_variants(sku: str) -> set[str]:
    """
    Generate all single-character OCR substitutions of `sku`.
    Returns a set of upper-cased variant strings (not including the original).
    """
    sku_upper = sku.upper()
    variants: set[str] = set()
    for i, ch in enumerate(sku_upper):
        for alt in _OCR_MAP.get(ch, []):
            variant = sku_upper[:i] + alt + sku_upper[i + 1:]
            if variant != sku_upper:
                variants.add(variant)
    return variants


def suggest_skus(
    raw_sku: str,
    inventory_all: set[str],
    inventory_supplier: set[str],
    max_results: int = 5,
) -> list[str]:
    """
    Return up to max_results inventory SKUs that are likely corrections for raw_sku.

    Stage 1: OCR single-char variants — checked against supplier set first, then all.
    Stage 2: Levenshtein distance against supplier set to fill remaining slots.
    """
    upper = raw_sku.upper()
    suggestions: list[str] = []
    seen: set[str] = set()

    # Stage 1: OCR variants
    variants = _ocr_variants(upper)
    for pool in (inventory_supplier, inventory_all):
        for v in variants:
            if v in pool and v not in seen:
                suggestions.append(v)
                seen.add(v)
                if len(suggestions) >= max_results:
                    return suggestions

    # Stage 2: Levenshtein against supplier pool (capped to avoid slow searches)
    if len(suggestions) < max_results and inventory_supplier:
        scored = [
            (s, _levenshtein(upper, s))
            for s in inventory_supplier
            if s not in seen
        ]
        scored.sort(key=lambda x: x[1])
        for s, dist in scored:
            if dist > 4:  # ignore very distant matches
                break
            suggestions.append(s)
            seen.add(s)
            if len(suggestions) >= max_results:
                break

    return suggestions


# ---------------------------------------------------------------------------
# Validation result
# ---------------------------------------------------------------------------

@dataclass
class ValidationResult:
    item: InvoiceItem
    is_confirmed: bool          # True if SKU found in inventory (or inventory missing)
    applied_correction: Optional[str]  # non-None if a saved correction was applied
    suggestions: list[str] = field(default_factory=list)


def validate_items(
    items: list[InvoiceItem],
    inventory_data: dict[str, set[str]],
    corrections: dict[tuple[str, str], str],
) -> list[ValidationResult]:
    """
    Validate each item's raw SKU against inventory_data.

    If inventory_data is empty (file missing) all items are marked confirmed.
    Saved corrections are applied before lookup.
    """
    results: list[ValidationResult] = []
    all_skus = inventory_data.get("", set())
    missing_inventory = not inventory_data

    for item in items:
        # Resolve supplier inventory subset
        sup_id = _SUPPLIER_ID_MAP.get(item.supplier_name.lower(), "")
        supplier_skus = inventory_data.get(sup_id, set()) if not missing_inventory else set()

        raw_upper = item.sku.upper()
        correction_key = (item.supplier_name.lower(), raw_upper)
        applied: Optional[str] = None

        # Apply saved correction if present
        if correction_key in corrections:
            applied = corrections[correction_key]
            lookup_sku = applied
        else:
            lookup_sku = raw_upper

        if missing_inventory:
            results.append(ValidationResult(
                item=item,
                is_confirmed=True,
                applied_correction=applied,
            ))
            continue

        if lookup_sku in all_skus:
            results.append(ValidationResult(
                item=item,
                is_confirmed=True,
                applied_correction=applied,
            ))
        else:
            suggs = suggest_skus(raw_upper, all_skus, supplier_skus)
            results.append(ValidationResult(
                item=item,
                is_confirmed=False,
                applied_correction=applied,
                suggestions=suggs,
            ))

    return results
