from __future__ import annotations

"""Thermal receipt PDF generator for in-store POS sales.

Generates an 80mm-wide receipt PDF to LOCAL_DATA_DIR/receipts/receipt_last.pdf.
Call generate_receipt(tx, cart_items) and then print the returned path.
"""

from datetime import datetime, timezone
from zoneinfo import ZoneInfo
from pathlib import Path
from typing import Optional

from reportlab.lib import colors
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.units import mm
from reportlab.platypus import (
    Flowable,
    HRFlowable,
    Paragraph,
    SimpleDocTemplate,
    Spacer,
    Table,
    TableStyle,
)
from reportlab.graphics.barcode.code128 import Code128
from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_RIGHT

from src.config import LOCAL_DATA_DIR, config


# ── Paper dimensions ──────────────────────────────────────────────────────────
# These are computed at render time from config so the user can change paper
# width in Settings → Printers without restarting the app.
# Common roll widths: 58 mm (desktop) · 72 mm · 76 mm · 80 mm (standard POS)

_MARGIN_MM = 3  # fixed narrow margin on each side


# ── Styles ────────────────────────────────────────────────────────────────────

def _styles():
    base = dict(fontName="Helvetica", leading=11)
    bold = dict(fontName="Helvetica-Bold", leading=11)
    return {
        "store_name": ParagraphStyle("store_name", fontSize=12, alignment=TA_CENTER,
                                     spaceAfter=1, **bold),
        "store_sub":  ParagraphStyle("store_sub",  fontSize=8,  alignment=TA_CENTER,
                                     spaceAfter=1, **base),
        "tx_info":    ParagraphStyle("tx_info",    fontSize=8,  alignment=TA_LEFT,
                                     spaceAfter=1, **base),
        "item_normal":ParagraphStyle("item_normal",fontSize=8,  alignment=TA_LEFT,
                                     spaceAfter=0, **base),
        "item_disc":  ParagraphStyle("item_disc",  fontSize=7,  alignment=TA_LEFT,
                                     spaceAfter=0, fontName="Helvetica-Oblique", leading=9),
        "label":      ParagraphStyle("label",      fontSize=8,  alignment=TA_LEFT,  **base),
        "label_bold": ParagraphStyle("label_bold", fontSize=8,  alignment=TA_LEFT,  **bold),
        "amount":     ParagraphStyle("amount",     fontSize=8,  alignment=TA_RIGHT, **base),
        "amount_bold":ParagraphStyle("amount_bold",fontSize=8,  alignment=TA_RIGHT, **bold),
        "total_label":ParagraphStyle("total_label",fontSize=10, alignment=TA_LEFT,  **bold),
        "total_amt":  ParagraphStyle("total_amt",  fontSize=10, alignment=TA_RIGHT, **bold),
        "gst":        ParagraphStyle("gst",        fontSize=7,  alignment=TA_LEFT,  **base),
        "footer":     ParagraphStyle("footer",     fontSize=8,  alignment=TA_CENTER,
                                     fontName="Helvetica-Oblique", leading=10),
        "notes":      ParagraphStyle("notes",      fontSize=8,  alignment=TA_LEFT,
                                     fontName="Helvetica-Oblique", leading=10,
                                     spaceAfter=1),
    }


# ── Helpers ───────────────────────────────────────────────────────────────────

def _fmt_phone(raw: str) -> str:
    digits = "".join(c for c in raw if c.isdigit())
    if len(digits) == 10 and digits.startswith("0"):
        return f"{digits[:2]} {digits[2:6]} {digits[6:]}"
    return raw


def _fmt_money(amount: Optional[float], prefix: str = "$") -> str:
    if amount is None:
        return "—"
    return f"{prefix}{amount:,.2f}"


_MELB_TZ = ZoneInfo("Australia/Melbourne")


def _fmt_datetime(iso: str) -> str:
    """Convert UTC ISO string to Melbourne local time, formatted as DD-MM-YYYY H:MM AM/PM."""
    try:
        import re
        s = iso.replace("Z", "+00:00")
        s = re.sub(r'\.(\d{1,6})', lambda m: "." + m.group(1).ljust(6, "0"), s)
        dt = datetime.fromisoformat(s)
        dt_local = dt.astimezone(_MELB_TZ)
        hour = dt_local.hour % 12 or 12          # 0 → 12, 13 → 1, etc.
        am_pm = "AM" if dt_local.hour < 12 else "PM"
        return (f"{dt_local.day:02d}-{dt_local.month:02d}-{dt_local.year}  "
                f"{hour}:{dt_local.minute:02d} {am_pm}")
    except Exception:
        return iso[:16] if len(iso) >= 16 else iso


def _divider(thick: bool = False) -> HRFlowable:
    return HRFlowable(
        width="100%",
        thickness=1.5 if thick else 0.75,
        color=colors.black,
        spaceAfter=3,
        spaceBefore=3,
    )


class _BarcodeFlowable(Flowable):
    """Centres a Code128 barcode on the thermal paper."""

    def __init__(self, value: str, bar_width: float = 0.75, bar_height: float = 28):
        super().__init__()
        self._value = value
        self._bar_width = bar_width
        self._bar_height = bar_height
        self._bc = Code128(value, barWidth=bar_width, barHeight=bar_height, humanReadable=True,
                           fontSize=7, fontName="Helvetica")

    def wrap(self, avail_w, avail_h):
        self.width  = avail_w
        self.height = self._bc.height
        return avail_w, self.height

    def draw(self):
        bc_w = self._bc.width
        x = (self.width - bc_w) / 2
        self._bc.drawOn(self.canv, x, 0)


# ── Main entry point ──────────────────────────────────────────────────────────

def generate_receipt(tx: dict, cart_items: dict, customer: Optional[dict] = None,
                     notes: Optional[str] = None) -> str:
    """Generate an 80mm thermal receipt PDF and return its absolute path.

    Parameters
    ----------
    tx:
        The transaction dict returned by ``confirm_standard_sale()``.
    cart_items:
        A snapshot of ``TillTab._cart_items`` captured before the cart was cleared.
        Keys are item_id strings; values are dicts with keys:
        sku, title, qty, unit_price, disc_pct.

    Returns
    -------
    str
        Absolute path to the generated PDF.
    """
    out_dir = Path(LOCAL_DATA_DIR) / "receipts"
    out_dir.mkdir(parents=True, exist_ok=True)
    out_path = str(out_dir / "receipt_last.pdf")

    # ── Paper dimensions (read from device config each call) ──────────────────
    paper_w_mm = (config.device.receipt_paper_width_mm or 80.0)
    page_w  = paper_w_mm * mm
    margin  = _MARGIN_MM * mm
    body_w  = page_w - 2 * margin
    # Line-items column widths — proportional to body width
    # SKU/desc ≈44% | Qty ≈10.5% | Unit ≈21% | Total ≈24%
    col_widths = [
        round(body_w * 0.442),
        round(body_w * 0.106),
        round(body_w * 0.212),
        round(body_w * 0.240),
    ]

    s = _styles()
    story: list = []

    # ── Store header ──────────────────────────────────────────────────────────
    sender = config.shipping.sender if config.shipping else None
    store_name = (sender.name if sender else "Scarlett Music")
    addr1 = sender.street1 if sender else ""
    city_line = ""
    if sender:
        parts = [p for p in [sender.city, sender.state, sender.postcode] if p]
        city_line = " ".join(parts)
    phone = _fmt_phone(sender.phone if sender else "")
    email = sender.email if sender else ""

    story.append(Paragraph(f"── {store_name.upper()} ──", s["store_name"]))
    if addr1:
        story.append(Paragraph(addr1, s["store_sub"]))
    if city_line:
        story.append(Paragraph(city_line, s["store_sub"]))
    if phone:
        story.append(Paragraph(phone, s["store_sub"]))
    if email:
        story.append(Paragraph(email, s["store_sub"]))
    story.append(Paragraph("ABN: 90 094 665 723", s["store_sub"]))
    story.append(Spacer(1, 3))

    # ── Transaction info ──────────────────────────────────────────────────────
    story.append(_divider(thick=True))
    tx_num = tx.get("transaction_number", "")
    completed_at = _fmt_datetime(tx.get("completed_at", ""))
    performed_by = tx.get("performed_by") or ""

    story.append(Paragraph(f"<b>{tx_num}</b>", s["tx_info"]))
    story.append(Paragraph(completed_at, s["tx_info"]))
    if performed_by:
        story.append(Paragraph(f"Staff: {performed_by}", s["tx_info"]))
    story.append(Spacer(1, 3))

    # ── Customer details (when linked) ────────────────────────────────────────
    if customer:
        first = customer.get("first_name") or ""
        last  = customer.get("surname") or ""
        cname = f"{first} {last}".strip()
        cid   = customer.get("customer_id")
        mob   = customer.get("mobile") or ""

        story.append(_divider(thick=False))
        if cname:
            story.append(Paragraph(f"Customer: <b>{cname}</b>", s["tx_info"]))
        if cid is not None:
            story.append(Paragraph(f"ID{cid}", s["tx_info"]))
        if mob:
            story.append(Paragraph(f"Ph: {_fmt_phone(mob)}", s["tx_info"]))
        story.append(Spacer(1, 3))

    # ── Transaction notes ─────────────────────────────────────────────────────
    _notes = (notes or tx.get("notes") or "").strip()
    if _notes:
        story.append(_divider(thick=False))
        story.append(Paragraph("Notes:", s["label_bold"]))
        for _line in _notes.splitlines():
            story.append(Paragraph(_line or " ", s["notes"]))
        story.append(Spacer(1, 3))

    # ── Line items ────────────────────────────────────────────────────────────
    story.append(_divider(thick=True))

    header_style = TableStyle([
        ("FONTNAME",    (0, 0), (-1, 0), "Helvetica-Bold"),
        ("FONTSIZE",    (0, 0), (-1, 0), 7),
        ("BOTTOMPADDING", (0, 0), (-1, 0), 2),
        ("TOPPADDING",  (0, 0), (-1, 0), 0),
        ("ALIGN",       (1, 0), (-1, 0), "RIGHT"),
        ("LINEBELOW",   (0, 0), (-1, 0), 0.5, colors.black),
    ])
    row_style_base = TableStyle([
        ("FONTNAME",    (0, 0), (-1, -1), "Helvetica"),
        ("FONTSIZE",    (0, 0), (-1, -1), 8),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 1),
        ("TOPPADDING",  (0, 0), (-1, -1), 1),
        ("ALIGN",       (1, 0), (-1, -1), "RIGHT"),
        ("VALIGN",      (0, 0), (-1, -1), "TOP"),
    ])

    # Header row
    tbl_data = [["Item", "Qty", "Unit", "Total"]]
    tbl_styles = [header_style]

    row_idx = 1
    for line in cart_items.values():
        sku: str   = line.get("sku", "")
        desc: str  = line.get("title", "")
        qty: float = line.get("qty", 1)
        price: float = line.get("unit_price", 0.0)
        disc_pct: float = line.get("disc_pct", 0.0)

        # Truncate description to fit column
        display = f"{sku}"
        if desc:
            max_chars = 18
            short_desc = desc[:max_chars] + ("…" if len(desc) > max_chars else "")
            display = f"{sku}\n{short_desc}"

        line_total = round(qty * price * (1 - disc_pct / 100), 2)

        # Main item row
        qty_display = str(int(qty)) if qty == int(qty) else f"{qty:.2f}"
        tbl_data.append([
            display,
            qty_display,
            f"{price:.2f}",
            f"{line_total:.2f}",
        ])
        row_idx += 1

        # Discount sub-row
        if disc_pct:
            disc_amt = round(qty * price * disc_pct / 100, 2)
            tbl_data.append([f"  Disc {disc_pct:.0f}%", "", "", f"-{disc_amt:.2f}"])
            tbl_styles.append(TableStyle([
                ("FONTNAME",  (0, row_idx), (-1, row_idx), "Helvetica-Oblique"),
                ("FONTSIZE",  (0, row_idx), (-1, row_idx), 7),
                ("TEXTCOLOR", (0, row_idx), (-1, row_idx), colors.gray),
                ("TOPPADDING", (0, row_idx), (-1, row_idx), 0),
            ]))
            row_idx += 1

    combined_style = TableStyle([
        ("FONTNAME",    (0, 0), (-1, -1), "Helvetica"),
        ("FONTSIZE",    (0, 0), (-1, -1), 8),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 1),
        ("TOPPADDING",  (0, 0), (-1, -1), 1),
        ("ALIGN",       (1, 0), (-1, -1), "RIGHT"),
        ("VALIGN",      (0, 0), (-1, -1), "TOP"),
        ("FONTNAME",    (0, 0), (-1, 0), "Helvetica-Bold"),
        ("FONTSIZE",    (0, 0), (-1, 0), 7),
        ("LINEBELOW",   (0, 0), (-1, 0), 0.5, colors.black),
    ])
    for extra in tbl_styles[1:]:
        for cmd in extra._cmds:
            combined_style.add(*cmd)

    items_table = Table(tbl_data, colWidths=col_widths, style=combined_style)
    story.append(items_table)
    story.append(Spacer(1, 4))

    # ── Totals ────────────────────────────────────────────────────────────────
    story.append(_divider(thick=True))
    subtotal = tx.get("subtotal") or 0.0
    cart_disc_total = tx.get("cart_discount_total")
    total = tx.get("total") or 0.0

    totals_style = TableStyle([
        ("FONTNAME",    (0, 0), (-1, -1), "Helvetica"),
        ("FONTSIZE",    (0, 0), (-1, -1), 8),
        ("ALIGN",       (0, 0), (0, -1), "LEFT"),
        ("ALIGN",       (1, 0), (1, -1), "RIGHT"),
        ("TOPPADDING",  (0, 0), (-1, -1), 1),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 1),
    ])

    totals_data: list[list] = []
    totals_data.append(["Subtotal", _fmt_money(subtotal)])
    if cart_disc_total:
        totals_data.append(["Discount", _fmt_money(-cart_disc_total)])

    # Add line discounts total if any
    line_disc_sum = sum(
        round(ln.get("qty", 1) * ln.get("unit_price", 0) * ln.get("disc_pct", 0) / 100, 2)
        for ln in cart_items.values()
        if ln.get("disc_pct")
    )
    if line_disc_sum and not cart_disc_total:
        totals_data.append(["Item discounts", _fmt_money(-line_disc_sum)])

    # TOTAL row — larger bold
    totals_table = Table(totals_data, colWidths=[body_w - 60, 60], style=totals_style)
    story.append(totals_table)

    story.append(_divider(thick=False))

    total_row = Table(
        [["TOTAL", _fmt_money(total)]],
        colWidths=[body_w - 60, 60],
        style=TableStyle([
            ("FONTNAME",  (0, 0), (-1, -1), "Helvetica-Bold"),
            ("FONTSIZE",  (0, 0), (-1, -1), 11),
            ("ALIGN",     (0, 0), (0, 0), "LEFT"),
            ("ALIGN",     (1, 0), (1, 0), "RIGHT"),
            ("TOPPADDING", (0, 0), (-1, -1), 2),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 2),
        ]),
    )
    story.append(total_row)
    story.append(Spacer(1, 4))

    # ── Payment breakdown ─────────────────────────────────────────────────────
    pay_data: list[list] = []

    cash = tx.get("payment_cash")
    if cash:
        pay_data.append(["Cash", _fmt_money(cash)])
        tendered = tx.get("cash_tendered")
        change = tx.get("change_given")
        if tendered and tendered != cash:
            pay_data.append(["  Tendered", _fmt_money(tendered)])
        if change and change > 0.001:
            pay_data.append(["  Change", f"-{_fmt_money(change)}"])

    eft_list = tx.get("payment_eft") or []
    if isinstance(eft_list, list):
        for i, entry in enumerate(eft_list, 1):
            amt = entry.get("amount") if isinstance(entry, dict) else entry
            label = f"EFT" if len(eft_list) == 1 else f"EFT #{i}"
            pay_data.append([label, _fmt_money(amt)])

    online = tx.get("payment_online")
    if online:
        pay_data.append(["Online", _fmt_money(online)])

    if pay_data:
        pay_style = TableStyle([
            ("FONTNAME",  (0, 0), (-1, -1), "Helvetica"),
            ("FONTSIZE",  (0, 0), (-1, -1), 8),
            ("ALIGN",     (0, 0), (0, -1), "LEFT"),
            ("ALIGN",     (1, 0), (1, -1), "RIGHT"),
            ("TOPPADDING", (0, 0), (-1, -1), 1),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 1),
        ])
        pay_table = Table(pay_data, colWidths=[body_w - 60, 60], style=pay_style)
        story.append(pay_table)
        story.append(Spacer(1, 4))

    # ── GST ───────────────────────────────────────────────────────────────────
    story.append(_divider(thick=True))
    gst_amount = round(total / 11, 2)
    gst_row = Table(
        [["GST Included:", _fmt_money(gst_amount)]],
        colWidths=[body_w - 60, 60],
        style=TableStyle([
            ("FONTNAME",  (0, 0), (-1, -1), "Helvetica"),
            ("FONTSIZE",  (0, 0), (-1, -1), 7),
            ("ALIGN",     (0, 0), (0, 0), "LEFT"),
            ("ALIGN",     (1, 0), (1, 0), "RIGHT"),
            ("TOPPADDING", (0, 0), (-1, -1), 1),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 1),
        ]),
    )
    story.append(gst_row)
    story.append(_divider(thick=True))
    story.append(Spacer(1, 6))

    # ── Barcode ───────────────────────────────────────────────────────────────
    if tx_num:
        story.append(_BarcodeFlowable(tx_num))
        story.append(Spacer(1, 8))

    # ── Footer ────────────────────────────────────────────────────────────────
    story.append(Paragraph("Thank you for shopping at", s["footer"]))
    store_display = store_name.title() if store_name.isupper() else store_name
    story.append(Paragraph(f"{store_display}!", s["footer"]))
    story.append(Spacer(1, 6))

    # ── Build document ────────────────────────────────────────────────────────
    # Use a generous fixed height, then trim to actual content with PyMuPDF
    # so no blank space appears at the end of the printed receipt.
    _TALL = 2000  # plenty for any realistic receipt

    doc = SimpleDocTemplate(
        out_path,
        pagesize=(page_w, _TALL),
        leftMargin=margin,
        rightMargin=margin,
        topMargin=margin,
        bottomMargin=margin,
        title=f"Receipt {tx_num}",
    )
    doc.build(story)

    _trim_to_content(out_path)
    return out_path


def _trim_to_content(pdf_path: str, pad: float = 8.0) -> None:
    """Resize the PDF page to the height of its actual content.

    Reads the built PDF, finds the lowest drawn element, and saves a new copy
    with the page height set to content_height + pad points.  This removes the
    blank space that would otherwise feed through the printer as blank paper.
    """
    import fitz  # PyMuPDF — already in requirements

    src = fitz.open(pdf_path)
    page = src[0]
    orig_w = page.rect.width
    orig_h = page.rect.height

    # PyMuPDF uses top-left coordinates (y increases downward).
    # Collect the lowest y1 of every text block and vector drawing.
    max_y: float = 0.0
    for block in page.get_text("blocks"):
        max_y = max(max_y, float(block[3]))
    for drw in page.get_drawings():
        max_y = max(max_y, float(drw["rect"].y1))

    if max_y <= 0 or max_y >= orig_h - pad:
        src.close()
        return  # Nothing useful to trim

    trimmed_h = max_y + pad

    # Build a new single-page document with the exact trimmed dimensions,
    # copying only the top portion (0 … trimmed_h) of the original page.
    out = fitz.open()
    new_page = out.new_page(width=orig_w, height=trimmed_h)
    new_page.show_pdf_page(
        new_page.rect,
        src,
        0,
        clip=fitz.Rect(0, 0, orig_w, trimmed_h),
    )
    src.close()  # release Windows file lock on the source before writing

    # Save to a sibling .tmp file, then atomically replace the original.
    # Saving directly to pdf_path while src had it open causes a
    # "Permission Denied / cannot remove file" error on Windows.
    import os
    tmp_path = pdf_path + ".tmp"
    out.save(tmp_path, garbage=4, deflate=True)
    out.close()
    os.replace(tmp_path, pdf_path)
