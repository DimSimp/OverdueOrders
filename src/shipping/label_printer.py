from __future__ import annotations

import logging
import re
import tempfile
from pathlib import Path

log = logging.getLogger("label_printer")

# Known Poppler binary paths on Windows (tried in order if not in PATH)
_POPPLER_PATHS = [
    r"C:\Program Files\poppler-0.68.0_x86\poppler-0.68.0\bin",
    r"C:\Program Files\poppler\bin",
    r"C:\poppler\bin",
    r"C:\Program Files (x86)\poppler\bin",
]


def _normalise_printer_id(raw: str) -> str:
    """Normalise a brother_ql discover() identifier for use with send().

    discover() on Windows may return identifiers like usb://0x04f9:0x2042_Љ
    where the product ID has a device-instance suffix.  send() (pyusb backend)
    parses vendor and product by splitting on ':' then calling int(..., 16),
    which fails if a suffix is present.  The legacy script explicitly maps
    these to usb://0x04F9:0x2042 (uppercase hex, no suffix).
    """
    m = re.match(r"^(usb://)(0x[0-9a-fA-F]+):(0x[0-9a-fA-F]+)", raw, re.IGNORECASE)
    if m:
        return f"{m.group(1)}{m.group(2).upper()}:{m.group(3).upper()}"
    # Fallback for unexpected formats: just uppercase any hex literals
    return re.sub(r"0x([0-9a-fA-F]+)", lambda m: f"0x{m.group(1).upper()}", raw)


def print_label(label_pdf: bytes, printer_model: str = "QL-700") -> str:
    """
    Print a 4×6 courier label PDF on a Brother QL-700 (62mm tape).

    Mirrors the working legacy approach:
      - PDF → PIL images via pdf2image
      - Split portrait label into left/right halves (each ~2"×6" strip fits 62mm tape)
      - Save halves as JPEG temp files
      - Print via brother_ql.brother_ql_create.convert + send

    Returns empty string on success, or an error message string on failure.
    """
    try:
        from pdf2image import convert_from_bytes
    except ImportError:
        return "pdf2image is not installed (pip install pdf2image)"

    try:
        from PIL import Image
    except ImportError:
        return "Pillow is not installed (pip install Pillow)"

    try:
        import brother_ql
        from brother_ql.raster import BrotherQLRaster
        from brother_ql.backends.helpers import send, discover
        # Use the legacy brother_ql_create API — same as the working legacy script
        _convert = brother_ql.brother_ql_create.convert
    except ImportError:
        return "brother_ql is not installed (pip install brother-ql)"
    except AttributeError:
        return "brother_ql.brother_ql_create module not found — unexpected library version"

    # Discover USB printer and normalise the identifier (lowercase → uppercase hex)
    try:
        found = discover("pyusb")
        log.debug("Brother QL printers found: %s", found)
        if not found:
            return "No Brother QL printer found via USB"
        raw_id = found[0]["identifier"]
        printer_id = _normalise_printer_id(raw_id)
        log.info("Using printer: %s (raw: %s)", printer_id, raw_id)
    except Exception as exc:
        return f"Printer discovery failed: {exc}"

    # Convert PDF bytes → PIL images, trying Poppler paths in order
    log.debug("Converting %d-byte PDF at 300 dpi", len(label_pdf))
    pages = None
    last_exc: Exception | None = None
    for poppler_path in [None] + _POPPLER_PATHS:
        try:
            kwargs: dict = {"dpi": 300}
            if poppler_path:
                kwargs["poppler_path"] = poppler_path
            pages = convert_from_bytes(label_pdf, **kwargs)
            if poppler_path:
                log.debug("PDF conversion succeeded with poppler_path=%s", poppler_path)
            break
        except Exception as exc:
            last_exc = exc
            log.debug("PDF conversion failed (poppler_path=%s): %s", poppler_path, exc)
    if not pages:
        return f"PDF conversion failed: {last_exc}"
    log.debug("PDF converted to %d page(s)", len(pages))

    # Process and print each page using the legacy approach
    with tempfile.TemporaryDirectory() as tmpdir:
        tmp = Path(tmpdir)
        for page_num, img in enumerate(pages):
            w, h = img.size
            log.debug("Page %d raw size: %dx%d (%s)", page_num + 1, w, h,
                      "landscape" if w > h else "portrait")

            # Horizontal split: top half and bottom half.
            # For a portrait 4×6 label (e.g. 1200×1800px) this gives two 1200×900
            # landscape pieces.  Rotate each 90° so they become portrait (900×1200)
            # strips that fit the 62mm tape width without excessive scaling.
            # For landscape source labels, rotate to portrait first.
            if w > h:
                img = img.transpose(Image.ROTATE_90)
                w, h = img.size
                log.debug("Page %d rotated to portrait: %dx%d", page_num + 1, w, h)

            # Scale down slightly so labels are less tight and the split doesn't
            # clip barcodes/QR codes near the midpoint.
            scale = 0.92
            new_w, new_h = int(w * scale), int(h * scale)
            img = img.resize((new_w, new_h), Image.LANCZOS)
            w, h = new_w, new_h
            log.debug("Page %d scaled to %dx%d", page_num + 1, w, h)

            # Split slightly below halfway so any QR/barcode near the centre
            # stays fully within the top half.  Pad the (shorter) bottom half
            # with white so both strips are the same height → equal printed size.
            mid = int(h * 0.57)
            top_half = img.crop((0, 0, w, mid)).convert("RGB")
            bot_half_raw = img.crop((0, mid, w, h)).convert("RGB")
            if bot_half_raw.height < mid:
                bot_half = Image.new("RGB", (w, mid), (255, 255, 255))
                bot_half.paste(bot_half_raw, (0, 0))
            else:
                bot_half = bot_half_raw

            # Rotate each landscape half to portrait so brother_ql fits it neatly
            # on the 62mm tape (e.g. 1200×900 → ROTATE_270 → 900×1200).
            top_strip = top_half.transpose(Image.ROTATE_270)
            bot_strip = bot_half.transpose(Image.ROTATE_270)

            log.debug(
                "Page %d: split at y=%d  top_strip=%s  bot_strip=%s",
                page_num + 1, mid, top_strip.size, bot_strip.size,
            )

            # Save as JPEG temp files — legacy code passes file paths to convert()
            top_path = str(tmp / f"label_p{page_num}_top.jpg")
            bot_path = str(tmp / f"label_p{page_num}_bot.jpg")
            top_strip.save(top_path, "JPEG")
            bot_strip.save(bot_path, "JPEG")

            # Print using legacy API: brother_ql_create.convert + send (positional)
            try:
                printer = BrotherQLRaster(printer_model)
                print_data = _convert(printer, [top_path, bot_path], "62", dither=True)
                send(print_data, printer_id)
                log.info("Page %d printed successfully to %s", page_num + 1, printer_id)
            except Exception as exc:
                log.error("Print failed for page %d: %s", page_num + 1, exc, exc_info=True)
                return f"Print failed: {exc}"

    return ""
