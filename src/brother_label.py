from __future__ import annotations

import os
import tempfile


def print_shipping_label(
    order_id: str,
    name: str,
    company: str,
    street1: str,
    street2: str,
    city: str,
    state: str,
    postcode: str,
) -> None:
    """Print a shipping label on the connected Brother QL-700 (or QL-570) printer.

    Label format:
        {order_id}
        {name}
        {company}          (omitted if blank)
        {street1}
        {street2}          (omitted if blank)
        {city}
        {state}\t{postcode}

    Raises RuntimeError if no printer is found, or any exception from PIL /
    brother_ql if the print fails.
    """
    from PIL import Image, ImageDraw, ImageFont
    import brother_ql
    from brother_ql.raster import BrotherQLRaster
    from brother_ql.backends.helpers import send, discover

    # ── Build label text ──────────────────────────────────────────────────
    parts = [order_id]
    for field in (name, company, street1, street2):
        if field and field.strip():
            parts.append(field.strip())
    if city and city.strip():
        parts.append(city.strip())
    state_post = "    ".join(f for f in (state.strip(), postcode.strip()) if f)
    if state_post:
        parts.append(state_post)
    text = "\n".join(parts)

    # ── Render to image ───────────────────────────────────────────────────
    img = Image.new("RGB", (400, 240), color=(255, 255, 255))
    try:
        fnt = ImageFont.truetype(r"C:\Windows\Fonts\calibri.ttf", 23, encoding="unic")
    except Exception:
        fnt = ImageFont.load_default()

    ImageDraw.Draw(img).multiline_text((10, 10), text, font=fnt, fill=(0, 0, 0))

    # ── Save, rotate, save ────────────────────────────────────────────────
    tmp = tempfile.NamedTemporaryFile(suffix=".png", delete=False)
    tmp.close()
    filename = tmp.name

    try:
        img.save(filename)
        Image.open(filename).transpose(Image.ROTATE_90).save(filename)

        # ── Discover printer ──────────────────────────────────────────────
        printers = discover("pyusb")
        if not printers:
            raise RuntimeError("No Brother printer found via USB.")

        raw_id: str = printers[0]["identifier"]
        raw_lower = raw_id.lower()

        if "0x2042" in raw_lower:
            printer_id = "usb://0x04F9:0x2042"
            printer = BrotherQLRaster("QL-700")
        elif "0x2028" in raw_lower:
            printer_id = "usb://0x04F9:0x2028"
            printer = BrotherQLRaster("QL-570")
        else:
            printer_id = raw_id
            printer = BrotherQLRaster("QL-700")

        # ── Send to printer ───────────────────────────────────────────────
        try:
            print_data = brother_ql.brother_ql_create.convert(
                printer, [filename], "62", dither=True
            )
            send(print_data, printer_id)
        except Exception:
            # Fallback to QL-570 identifier (matches legacy try/except)
            printer_id = "usb://0x04F9:0x2028"
            printer = BrotherQLRaster("QL-570")
            print_data = brother_ql.brother_ql_create.convert(
                printer, [filename], "62", dither=True
            )
            send(print_data, printer_id)

    finally:
        try:
            os.unlink(filename)
        except Exception:
            pass
