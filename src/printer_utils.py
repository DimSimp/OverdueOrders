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


def get_printable_width_mm(printer_name: str) -> float:
    """Return the printable area width of the named printer in mm.

    Queries the GDI HORZRES / LOGPIXELSX device capabilities via ctypes.
    Returns 0.0 if the printer cannot be opened or the call fails.
    """
    try:
        import ctypes
        gdi = ctypes.windll.gdi32
        hdc = gdi.CreateDCW("WINSPOOL", printer_name, None, None)
        if not hdc:
            return 0.0
        HORZRES    = 8
        LOGPIXELSX = 88
        pixels = gdi.GetDeviceCaps(hdc, HORZRES)
        dpi    = gdi.GetDeviceCaps(hdc, LOGPIXELSX)
        gdi.DeleteDC(hdc)
        if dpi <= 0:
            return 0.0
        return pixels / dpi * 25.4
    except Exception:
        return 0.0


def _resize_pdf_width(pdf_path: str, target_w_mm: float) -> None:
    """Scale a PDF page to target_w_mm wide (proportionally), in-place.

    Skips silently if the page is already within 0.5 mm of the target or if
    PyMuPDF is unavailable.
    """
    try:
        import fitz
        MM_TO_PT = 72.0 / 25.4
        target_pt = target_w_mm * MM_TO_PT

        src = fitz.open(pdf_path)
        orig_w = src[0].rect.width
        orig_h = src[0].rect.height
        if abs(orig_w - target_pt) < (0.5 * MM_TO_PT):
            src.close()
            return  # already close enough

        scale  = target_pt / orig_w
        new_w  = target_pt
        new_h  = orig_h * scale

        out      = fitz.open()
        new_page = out.new_page(width=new_w, height=new_h)
        new_page.show_pdf_page(new_page.rect, src, 0)
        src.close()

        tmp = pdf_path + ".sz"
        out.save(tmp, garbage=4, deflate=True)
        out.close()
        os.replace(tmp, pdf_path)
    except Exception:
        pass  # Print the original if resize fails


def print_pdf(pdf_path: str, printer_name: str) -> None:
    """Send a PDF to the named printer via SumatraPDF (silent, non-blocking).

    Queries the printer's actual printable-area width via GDI and pre-scales
    the PDF to that width before printing, so that noscale printing is
    edge-accurate regardless of the hardware margin the driver exposes.

    Raises RuntimeError if SumatraPDF is not found.
    """
    sumatra = _find_sumatra()
    if not sumatra:
        raise RuntimeError(
            "SumatraPDF not found.\n"
            "Install it from https://www.sumatrapdfreader.org/ or place "
            "SumatraPDF.exe next to the app."
        )

    pw = get_printable_width_mm(printer_name)
    if pw > 0:
        _resize_pdf_width(pdf_path, pw)

    subprocess.Popen([
        sumatra,
        "-print-to", printer_name,
        "-print-settings", "noscale,portrait,1x",
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
