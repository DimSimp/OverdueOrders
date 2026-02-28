from __future__ import annotations

from datetime import datetime
from pathlib import Path

import pandas as pd

from src.data_processor import MatchedOrder

COLUMNS = [
    "Platform",
    "Order Number",
    "Customer",
    "Order Date",
    "SKU",
    "Description",
    "Quantity",
    "Invoice SKU",
    "Invoice Description",
    "Invoice Qty",
    "Notes",
]


def export_to_csv(
    matched_orders: list[MatchedOrder],
    output_dir: str,
    filename_prefix: str = "overdue_matches",
) -> str:
    """
    Export matched orders to a CSV file sorted by Platform (Neto first) then Order Date.
    Returns the absolute path of the created file.
    Raises ValueError if no orders to export.
    """
    if not matched_orders:
        raise ValueError("No matched orders to export.")

    rows = []
    for m in matched_orders:
        date_str = ""
        if m.order_date:
            try:
                date_str = m.order_date.strftime("%Y-%m-%d")
            except Exception:
                date_str = str(m.order_date)

        rows.append({
            "Platform": m.platform,
            "Order Number": m.order_id,
            "Customer": m.customer_name,
            "Order Date": date_str,
            "SKU": m.sku,
            "Description": m.description,
            "Quantity": m.quantity,
            "Invoice SKU": m.invoice_sku,
            "Invoice Description": m.invoice_description,
            "Invoice Qty": m.invoice_qty,
            "Notes": m.notes,
        })

    df = pd.DataFrame(rows, columns=COLUMNS)

    # Sort: Website first, direct eBay last, all other channels alphabetically in between
    def _platform_sort_key(p: str) -> tuple:
        pl = p.lower()
        if pl == "website":
            return (0, p)
        if pl == "ebay":
            return (2, p)
        return (1, p)

    df["_sort_platform"] = df["Platform"].apply(_platform_sort_key)
    df = df.sort_values(["_sort_platform", "Order Date"]).drop(columns=["_sort_platform"])

    out_path = Path(output_dir)
    out_path.mkdir(parents=True, exist_ok=True)

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    filepath = out_path / f"{filename_prefix}_{timestamp}.csv"

    # utf-8-sig BOM so Excel on Windows opens it correctly
    df.to_csv(filepath, index=False, encoding="utf-8-sig")

    return str(filepath.resolve())
