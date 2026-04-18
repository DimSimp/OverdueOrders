from __future__ import annotations

import csv
from datetime import datetime
from typing import Optional

# Musipos CSV column → Supabase items column
FIELD_MAP: dict[str, str] = {
    "Supplier_Item_ID":                 "sku",
    "Title":                            "title",
    "Supplier_RRP":                     "supplier_rrp",
    "Publisher_Brand":                  "brand",
    "Artist_Composer_Series":           "series",
    "Instrument":                       "instrument",
    "Sub_Instrument":                   "sub_instrument",
    "Quantity_on_hand":                 "qty_on_hand",
    "Last_Purchase_Cost":               "last_purchase_cost",
    "Barcode":                          "internal_barcode",
    "Minimum_Sell":                     "minimum_sell",
    "Last_Purchase_Date":               "last_purchase_date",
    "Last_Sold_Date":                   "last_sold_date",
    "Created_Date":                     "created_date",
    "Stock_Availability_from_Supplier": "stock_availability_from_supplier",
    "Supplier_ID":                      "supplier_id",
    "Product_Barcode":                  "product_barcode",
    "Active":                           "active",
    # Ignored: Product_Type, Category, Department, Quantity_Availability
}

_DATE_COLS  = {"last_purchase_date", "last_sold_date", "created_date"}
_BOOL_COLS  = {"active", "stock_availability_from_supplier"}
_PRICE_COLS = {"supplier_rrp", "last_purchase_cost"}

# numeric(10,2) max is 99,999,999.99 — anything above this limit is bad data
# (e.g. a barcode accidentally in a price column). Use a conservative cap.
_PRICE_MAX = 999_999.99

# Musipos exports are Windows-1252 (contains ® and similar chars)
_ENCODING = "cp1252"


def _parse_date(s: str) -> Optional[str]:
    for fmt in ("%d/%m/%Y", "%Y-%m-%d"):
        try:
            return datetime.strptime(s.strip(), fmt).strftime("%Y-%m-%d")
        except (ValueError, AttributeError):
            pass
    return None


def _parse_decimal(s: str) -> Optional[float]:
    try:
        v = float(s.strip())
        return v if v != 0.0 else None
    except (ValueError, AttributeError):
        return None


def load_csv(path: str, import_quantities: bool = False) -> list[dict]:
    """Parse the Musipos inventory CSV and return cleaned rows ready for Supabase upsert.

    Args:
        path:              Path to the Musipos .CSV export.
        import_quantities: If True, populate qty_on_hand from the CSV value.
                           Default False — leaves qty_on_hand = 0 since Musipos
                           export stock counts are typically out of date.
    """
    rows: list[dict] = []

    with open(path, encoding=_ENCODING, newline="") as f:
        reader = csv.DictReader(f)
        for raw in reader:
            sku = (raw.get("Supplier_Item_ID") or "").strip()
            if not sku:
                continue

            row: dict = {}
            for csv_col, db_col in FIELD_MAP.items():
                val = (raw.get(csv_col) or "").strip()

                if db_col in _DATE_COLS:
                    row[db_col] = _parse_date(val) if val else None

                elif db_col in _PRICE_COLS:
                    try:
                        v = float(val) if val else None
                        if v is not None and v > _PRICE_MAX:
                            print(f"[IMPORT] SKU {sku!r}: {db_col}={v} exceeds limit — set to NULL")
                            v = None
                        row[db_col] = v
                    except ValueError:
                        row[db_col] = None

                elif db_col == "minimum_sell":
                    # 0 means "no minimum" in Musipos → store as NULL
                    try:
                        v = float(val) if val else None
                        if v is not None and v > _PRICE_MAX:
                            print(f"[IMPORT] SKU {sku!r}: minimum_sell={v} exceeds limit — set to NULL")
                            v = None
                        row[db_col] = v if (v and v > 0) else None
                    except ValueError:
                        row[db_col] = None

                elif db_col in _BOOL_COLS:
                    row[db_col] = val.upper() == "Y"

                elif db_col == "qty_on_hand":
                    if import_quantities:
                        try:
                            row[db_col] = int(float(val)) if val else 0
                        except ValueError:
                            row[db_col] = 0
                    else:
                        row[db_col] = 0

                else:
                    row[db_col] = val or None

            # title NOT NULL — fall back to SKU if the CSV has a blank title
            if not row.get("title"):
                row["title"] = sku

            rows.append(row)

    # Deduplicate / disambiguate rows that share the same SKU.
    #
    # Two cases:
    #   • Same SKU, same supplier  → true duplicate (data entry); keep the last row.
    #   • Same SKU, diff suppliers → genuinely different products; rename every row
    #     in the group to  "{sku}_{supplier_id}"  so nothing is lost.
    from collections import defaultdict

    sku_groups: dict = defaultdict(list)
    for row in rows:
        sku_groups[row["sku"]].append(row)

    deduped: list[dict] = []
    same_supplier_dupes = 0
    cross_supplier_renamed = 0

    for sku, group in sku_groups.items():
        if len(group) == 1:
            deduped.append(group[0])
            continue

        suppliers = {r.get("supplier_id") for r in group}
        if len(suppliers) == 1:
            # True duplicate — keep last occurrence
            deduped.append(group[-1])
            same_supplier_dupes += len(group) - 1
        else:
            # Cross-supplier conflict — append supplier ID to every row in the group
            for row in group:
                sid = row.get("supplier_id") or "UNKNOWN"
                row = dict(row)          # don't mutate original
                row["sku"] = f"{sku}_{sid}"
                deduped.append(row)
            cross_supplier_renamed += len(group)

    if same_supplier_dupes or cross_supplier_renamed:
        print(f"[IMPORT] Deduplication: {same_supplier_dupes} same-supplier duplicate(s) removed, "
              f"{cross_supplier_renamed} cross-supplier SKU(s) renamed (suffix added).")

    # Final safety pass — a renamed SKU (e.g. "ABC_PAYTONS") may collide with an
    # existing standalone SKU that was already called "ABC_PAYTONS". Keep last.
    final_seen: dict[str, dict] = {}
    for row in deduped:
        final_seen[row["sku"]] = row
    deduped = list(final_seen.values())

    print(f"[IMPORT] {len(rows)} raw rows → {len(deduped)} unique items to upsert.")

    return deduped


def get_preview(path: str, n: int = 5) -> tuple[list[str], list[list]]:
    """Return (column_headers, first_n_data_rows) for the preview table."""
    headers = list(FIELD_MAP.values())
    rows = load_csv(path, import_quantities=True)[:n]
    return headers, [[str(r.get(h) or "") for h in headers] for r in rows]


def get_unique_supplier_ids(path: str) -> list[str]:
    """Return all distinct Supplier_ID values from the CSV (for stub creation)."""
    ids: set[str] = set()
    with open(path, encoding=_ENCODING, newline="") as f:
        for row in csv.DictReader(f):
            sid = (row.get("Supplier_ID") or "").strip()
            if sid:
                ids.add(sid)
    return sorted(ids)


def count_rows(path: str) -> int:
    """Count data rows (excluding header)."""
    with open(path, encoding=_ENCODING, newline="") as f:
        return sum(1 for _ in csv.DictReader(f))
