from __future__ import annotations

import os
import shutil
import subprocess
import sys
from pathlib import Path


def list_printers() -> list[str]:
    """Return installed printer display names via PowerShell Get-Printer."""
    try:
        out = subprocess.check_output(
            [
                "powershell", "-NoProfile", "-Command",
                "Get-Printer | Select-Object -ExpandProperty Name",
            ],
            text=True,
            timeout=8,
            stderr=subprocess.DEVNULL,
        )
        return [ln.strip() for ln in out.splitlines() if ln.strip()]
    except Exception:
        return []


def print_pdf(pdf_path: str, printer_name: str) -> None:
    """Send a PDF to the named printer via SumatraPDF (silent, non-blocking).

    Raises RuntimeError if SumatraPDF is not found.
    """
    sumatra = _find_sumatra()
    if not sumatra:
        raise RuntimeError(
            "SumatraPDF not found.\n"
            "Install it from https://www.sumatrapdfreader.org/ or place "
            "SumatraPDF.exe next to the app."
        )
    subprocess.Popen([
        sumatra,
        "-print-to", printer_name,
        "-print-settings", "1x",
        "-silent",
        pdf_path,
    ])


def _find_sumatra() -> str | None:
    """Return the path to SumatraPDF.exe, or None if not found.

    Search order:
      1. PATH (system-wide install)
      2. App folder — staff can place SumatraPDF.exe alongside the exe
      3. Common install locations
    """
    if getattr(sys, "frozen", False):
        app_dir = Path(sys.executable).parent
    else:
        app_dir = Path(__file__).parent.parent

    app_dir_matches = [str(p) for p in sorted(app_dir.glob("SumatraPDF*.exe"))]

    candidates = [
        shutil.which("SumatraPDF"),
        *app_dir_matches,
        r"C:\Program Files\SumatraPDF\SumatraPDF.exe",
        r"C:\Program Files (x86)\SumatraPDF\SumatraPDF.exe",
        os.path.expanduser(r"~\AppData\Local\SumatraPDF\SumatraPDF.exe"),
    ]
    for p in candidates:
        if p and os.path.isfile(p):
            return p
    return None
