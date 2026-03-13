from __future__ import annotations

import logging
import re
import tempfile
from pathlib import Path
from typing import TYPE_CHECKING

if TYPE_CHECKING:
    from PIL import Image as _Image

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
    return re.sub(r"0x([0-9a-fA-F]+)", lambda m: f"0x{m.group(1).upper()}", raw)


def process_label_pdf(
    pdf_bytes: bytes,
    scale: float = 0.92,
    split_ratio: float = 0.57,
    no_split: bool = False,
    label_length_mm: float = 0.0,
    rotate_cw: bool = False,
    skip_initial_rotate: bool = False,
) -> list[tuple]:
    """Convert a courier label PDF into printable strip pairs (or a single strip).

    Returns a list of ``(top_strip, bot_strip)`` PIL Image tuples — one per
    PDF page.  Each strip is an RGB portrait image ready to print on 62mm tape.
    When *no_split* is True, ``bot_strip`` is ``None`` and the full label is
    returned as a single portrait strip (use for labels with vertical barcodes,
    e.g. AusPost Express, where any horizontal cut would bisect the barcode).

    When *label_length_mm* > 0 the strip height is resized so that — given
    the printer always stretches strip width to fill 62mm tape — the printed
    label will be exactly *label_length_mm* millimetres long (tape feed direction).

    *rotate_cw* controls the direction of the landscape→portrait rotation in step 1,
    and the strip rotation direction in step 4.
    False (default): initial rotate CCW (ROTATE_90), strip rotate CCW (ROTATE_270).
    True: initial rotate CW (ROTATE_270), strip rotate CW (ROTATE_90) — use for
    Allied Express where strips must be rotated CW to read correctly.

    *skip_initial_rotate* — when True, skip step 2 entirely (no landscape→portrait
    conversion). Use for Allied Express labels which arrive in landscape orientation
    and should be split horizontally as-is; the per-half strip rotation in step 4
    handles the final portrait orientation. When False (default), the initial
    rotation is applied only when the page is already landscape (w > h).

    Normal split pipeline:
      1. PDF page → PIL image at 300 dpi
      2. If landscape (w > h) AND NOT skip_initial_rotate: rotate to portrait
      3. Scale by *scale*
      4. Split horizontally at *split_ratio* of the image height
      5. Pad the shorter bottom half with white so both strips are equal height
      6. Rotate each half to portrait strip (direction controlled by rotate_cw)
      7. If label_length_mm > 0, resize strip height to achieve target length

    No-split pipeline (no_split=True):
      Steps 1–3 as above, then return the full portrait image as the sole
      element of the tuple (bot_strip=None). No rotation — the portrait image
      is already the correct orientation for 62mm tape (width→62mm, height→length).
      Step 7 applied if label_length_mm > 0.
    """
    from pdf2image import convert_from_bytes
    from PIL import Image

    # ── PDF → PIL images ──────────────────────────────────────────────────
    pages = None
    last_exc: Exception | None = None
    for poppler_path in [None] + _POPPLER_PATHS:
        try:
            kwargs: dict = {"dpi": 300}
            if poppler_path:
                kwargs["poppler_path"] = poppler_path
            pages = convert_from_bytes(pdf_bytes, **kwargs)
            break
        except Exception as exc:
            last_exc = exc
    if not pages:
        raise RuntimeError(f"PDF conversion failed: {last_exc}")

    # rotation constants — flipped depending on rotate_cw
    _portrait_rot = Image.ROTATE_270 if rotate_cw else Image.ROTATE_90
    _strip_rot = Image.ROTATE_90 if rotate_cw else Image.ROTATE_270

    result = []
    for page_num, img in enumerate(pages):
        w, h = img.size

        # ── 1. Ensure portrait (skipped for landscape labels like Allied) ──
        if w > h and not skip_initial_rotate:
            img = img.transpose(_portrait_rot)
            w, h = img.size

        # ── 2. Scale ────────────────────────────────────────────────────
        new_w, new_h = int(w * scale), int(h * scale)
        img = img.resize((new_w, new_h), Image.LANCZOS)
        w, h = new_w, new_h

        if no_split:
            # ── 3a. No-split: full label as a single portrait strip ──────
            full_strip = img.convert("RGB")
            if label_length_mm > 0:
                target_h = round(label_length_mm / 62.0 * full_strip.width)
                full_strip = full_strip.resize((full_strip.width, target_h), Image.LANCZOS)
            log.debug(
                "process_label_pdf page %d (no_split): raw=%dx%d scale=%.2f "
                "label_length_mm=%.1f full_strip=%s",
                page_num + 1, w, h, scale, label_length_mm, full_strip.size,
            )
            result.append((full_strip, None))
        else:
            # ── 3. Split + pad ───────────────────────────────────────────
            mid = int(h * split_ratio)
            top_half = img.crop((0, 0, w, mid)).convert("RGB")
            bot_half_raw = img.crop((0, mid, w, h)).convert("RGB")
            if bot_half_raw.height < mid:
                bot_half = Image.new("RGB", (w, mid), (255, 255, 255))
                bot_half.paste(bot_half_raw, (0, 0))
            else:
                bot_half = bot_half_raw

            # ── 4. Rotate halves to portrait strips ──────────────────────
            top_strip = top_half.transpose(_strip_rot)
            bot_strip = bot_half.transpose(_strip_rot)

            # ── 5. Target label length ───────────────────────────────────
            if label_length_mm > 0:
                target_h = round(label_length_mm / 62.0 * top_strip.width)
                top_strip = top_strip.resize((top_strip.width, target_h), Image.LANCZOS)
                bot_strip = bot_strip.resize((bot_strip.width, target_h), Image.LANCZOS)

            log.debug(
                "process_label_pdf page %d: raw=%dx%d scale=%.2f split=%.2f "
                "label_length_mm=%.1f rotate_cw=%s skip_initial_rotate=%s "
                "top_strip=%s bot_strip=%s",
                page_num + 1, w, h, scale, split_ratio, label_length_mm,
                rotate_cw, skip_initial_rotate, top_strip.size, bot_strip.size,
            )
            result.append((top_strip, bot_strip))

    return result


def print_label(
    label_pdf: bytes,
    printer_model: str = "QL-700",
    courier_code: str = "",
    scale: float | None = None,
    split_ratio: float | None = None,
    no_split: bool | None = None,
    label_length_mm: float | None = None,
    rotate_cw: bool | None = None,
    skip_initial_rotate: bool | None = None,
) -> str:
    """Print a 4×6 courier label PDF on a Brother QL-700 (62mm tape).

    Per-courier scale, split_ratio, label_length_mm, no_split, and rotate_cw
    are loaded from ``Test Labels/label_settings.json`` when not supplied directly.

    Returns empty string on success, or an error message string on failure.
    """
    # ── Resolve settings from file for any None params ─────────────────────
    if any(v is None for v in (scale, split_ratio, label_length_mm, no_split, rotate_cw, skip_initial_rotate)):
        from src.shipping.label_settings import load as _load_settings
        settings = _load_settings(courier_code)
        if scale is None:
            scale = settings["scale"]
        if split_ratio is None:
            split_ratio = settings["split_ratio"]
        if label_length_mm is None:
            label_length_mm = settings.get("label_length_mm", 0.0)
        if no_split is None:
            no_split = settings.get("no_split", False)
        if rotate_cw is None:
            rotate_cw = settings.get("rotate_cw", False)
        if skip_initial_rotate is None:
            skip_initial_rotate = settings.get("skip_initial_rotate", False)

    # ── Lazy imports ──────────────────────────────────────────────────────
    try:
        from pdf2image import convert_from_bytes  # noqa: F401 (triggers ImportError early)
    except ImportError:
        return "pdf2image is not installed (pip install pdf2image)"

    try:
        from PIL import Image  # noqa: F401
    except ImportError:
        return "Pillow is not installed (pip install Pillow)"

    try:
        import brother_ql
        from brother_ql.raster import BrotherQLRaster
        from brother_ql.backends.helpers import send, discover
        _convert = brother_ql.brother_ql_create.convert
    except ImportError:
        return "brother_ql is not installed (pip install brother-ql)"
    except AttributeError:
        return "brother_ql.brother_ql_create module not found — unexpected library version"

    # ── Discover printer ──────────────────────────────────────────────────
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

    # ── Process PDF into strip pairs ──────────────────────────────────────
    log.debug(
        "Converting %d-byte PDF (scale=%.2f split=%.2f no_split=%s "
        "label_length_mm=%.1f rotate_cw=%s skip_initial_rotate=%s)",
        len(label_pdf), scale, split_ratio, no_split, label_length_mm,
        rotate_cw, skip_initial_rotate,
    )
    try:
        strip_pairs = process_label_pdf(
            label_pdf, scale=scale, split_ratio=split_ratio, no_split=no_split,
            label_length_mm=label_length_mm, rotate_cw=rotate_cw,
            skip_initial_rotate=skip_initial_rotate,
        )
    except Exception as exc:
        return f"PDF processing failed: {exc}"

    # ── Print each page ───────────────────────────────────────────────────
    with tempfile.TemporaryDirectory() as tmpdir:
        tmp = Path(tmpdir)
        for page_num, (top_strip, bot_strip) in enumerate(strip_pairs):
            top_path = str(tmp / f"label_p{page_num}_top.jpg")
            top_strip.save(top_path, "JPEG")
            strip_paths = [top_path]

            if bot_strip is not None:
                bot_path = str(tmp / f"label_p{page_num}_bot.jpg")
                bot_strip.save(bot_path, "JPEG")
                strip_paths.append(bot_path)

            try:
                printer = BrotherQLRaster(printer_model)
                print_data = _convert(printer, strip_paths, "62", dither=True)
                send(print_data, printer_id)
                log.info(
                    "Page %d printed successfully to %s (%d strip(s))",
                    page_num + 1, printer_id, len(strip_paths),
                )
            except Exception as exc:
                log.error("Print failed for page %d: %s", page_num + 1, exc, exc_info=True)
                return f"Print failed: {exc}"

    return ""
