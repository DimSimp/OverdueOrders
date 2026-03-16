from __future__ import annotations

import ftplib
import os
import tempfile
from dataclasses import dataclass

import pandas as pd


@dataclass
class ReceivedItem:
    sku: str
    quantity: float
    supplier: str


def download_and_compare(host: str, username: str, password: str,
                         morning_filename: str, afternoon_filename: str) -> list[ReceivedItem]:
    """
    1. Connect to FTP and download both inventory reports to temp files.
    2. Compare them: SKUs with increased quantity = received today.
    3. Return a list of ReceivedItem.
    Raises ftplib.all_errors or OSError on failure.
    """
    with tempfile.TemporaryDirectory() as tmpdir:
        morning_path = os.path.join(tmpdir, morning_filename)
        afternoon_path = os.path.join(tmpdir, afternoon_filename)

        ftp = ftplib.FTP(host)
        ftp.login(username, password)
        with open(morning_path, "wb") as f:
            ftp.retrbinary(f"RETR {morning_filename}", f.write)
        with open(afternoon_path, "wb") as f:
            ftp.retrbinary(f"RETR {afternoon_filename}", f.write)
        ftp.quit()

        return _compare_reports(morning_path, afternoon_path)


def compare_local_files(morning_path: str, afternoon_path: str) -> list[ReceivedItem]:
    """Compare two local inventory Excel files directly (no FTP needed).

    Useful when the inventory reports are accessible on a local or network drive.
    Uses the same delta logic as the FTP method.
    """
    return _compare_reports(morning_path, afternoon_path)


def _compare_reports(morning_path: str, afternoon_path: str) -> list[ReceivedItem]:
    df1 = pd.read_excel(morning_path, header=None)
    df2 = pd.read_excel(afternoon_path, header=None)

    morning_qty = df1.set_index(0)[8].to_dict()
    afternoon_qty = df2.set_index(0)[8].to_dict()
    supplier_map = df2.set_index(0)[18].to_dict()

    results = []
    all_skus = set(morning_qty) | set(afternoon_qty)
    for sku in all_skus:
        try:
            delta = float(afternoon_qty.get(sku, 0)) - float(morning_qty.get(sku, 0))
        except (ValueError, TypeError):
            continue
        if delta > 0:
            results.append(ReceivedItem(
                sku=str(sku).strip(),
                quantity=delta,
                supplier=str(supplier_map.get(sku, "")).strip(),
            ))
    return results
