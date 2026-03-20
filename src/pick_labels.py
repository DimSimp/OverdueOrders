from __future__ import annotations

import os
import re
import tempfile
from typing import Callable, Optional


class PrintLabelError(Exception):
    pass


def build_label_list(
    orders: list,
    pick_zones: dict,
) -> list[tuple]:
    """Return [(sku, description), ...] — one entry per quantity unit for 'Picks' zone."""
    labels = []
    for order in orders:
        for li in order.line_items:
            if pick_zones.get(li.sku) == "Picks":
                name = (
                    getattr(li, "product_name", None)
                    or getattr(li, "title", None)
                    or li.sku
                )
                for _ in range(li.quantity):
                    labels.append((li.sku, name))
    return labels


def print_pick_labels(
    labels: list[tuple],
    progress_callback: Optional[Callable[[int, int], None]] = None,
) -> None:
    """Generate and print one Brother QL label per entry in *labels*.

    progress_callback(printed, total) is called after each label is sent.
    Raises PrintLabelError on printer or dependency failure.
    """
    try:
        from PIL import Image, ImageDraw, ImageFont
        import brother_ql
        import brother_ql.brother_ql_create
        from brother_ql.raster import BrotherQLRaster
        from brother_ql.backends.helpers import send, discover
    except ImportError as exc:
        raise PrintLabelError(f"Missing dependency: {exc}") from exc

    # Discover USB printer
    try:
        discovered = discover("pyusb")
        if not discovered:
            raise PrintLabelError("No Brother QL printer found via USB")
        identifier_raw: str = discovered[0]["identifier"]
    except PrintLabelError:
        raise
    except Exception as exc:
        raise PrintLabelError(f"Printer discovery failed: {exc}") from exc

    # Normalise identifier and choose model
    id_lower = identifier_raw.lower()
    if "0x2042" in id_lower:
        PRINTER_IDENTIFIER = "usb://0x04F9:0x2042"
        printer_model = "QL-700"
    elif "0x2028" in id_lower:
        PRINTER_IDENTIFIER = "usb://0x04F9:0x2028"
        printer_model = "QL-570"
    else:
        PRINTER_IDENTIFIER = identifier_raw
        printer_model = "QL-700"

    font_path = r"C:\Windows\Fonts\calibri.ttf"
    tmpfile = os.path.join(tempfile.gettempdir(), "scarlett_pick_label.png")
    total = len(labels)

    def _send_one(sku: str, description: str) -> None:
        text = f"[SKU: {sku}] {description}"
        text = re.sub(r"(.{30})", r"\1\n", text, 0, re.DOTALL)
        nlines = text.count("\n")
        image_height = 22 * (nlines + 1)

        img = Image.new("RGB", (240, image_height), color=(255, 255, 255))
        try:
            fnt = ImageFont.truetype(font_path, 18, encoding="unic")
        except OSError:
            fnt = ImageFont.load_default()

        d = ImageDraw.Draw(img)
        d.multiline_text((10, 10), text, font=fnt, fill=(0, 0, 0))
        img.save(tmpfile)

        printer = BrotherQLRaster(printer_model)
        print_data = brother_ql.brother_ql_create.convert(printer, [tmpfile], "62", dither=True)
        try:
            send(print_data, PRINTER_IDENTIFIER)
        except Exception:
            # Fallback to QL-570 if primary fails
            fallback = BrotherQLRaster("QL-570")
            fallback_data = brother_ql.brother_ql_create.convert(
                fallback, [tmpfile], "62", dither=True
            )
            send(fallback_data, "usb://0x04F9:0x2028")

    try:
        for i, (sku, description) in enumerate(labels):
            _send_one(sku, description)
            if progress_callback:
                progress_callback(i + 1, total)
    finally:
        try:
            os.remove(tmpfile)
        except OSError:
            pass
