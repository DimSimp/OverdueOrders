from __future__ import annotations

"""A4 customer document PDF generator for quotes, invoices, repairs, and deposits.

The quote document is the first consumer. The layout is intentionally reusable:
store/customer header, document metadata, line table, totals, notes, and terms.
"""

from datetime import datetime
from pathlib import Path
from typing import Optional
from xml.sax.saxutils import escape
import re
from zoneinfo import ZoneInfo

from reportlab.lib import colors
from reportlab.lib.enums import TA_LEFT, TA_RIGHT
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.units import mm
from reportlab.platypus import (
    Paragraph,
    SimpleDocTemplate,
    Spacer,
    Table,
    TableStyle,
)

from src.config import LOCAL_DATA_DIR, config


_PAGE_W, _PAGE_H = A4
_MARGIN = 16 * mm
_BODY_W = _PAGE_W - 2 * _MARGIN


def generate_quote_document(
    tx: dict,
    cart_items: dict,
    customer: Optional[dict],
    notes: Optional[str] = None,
) -> str:
    """Generate an A4 quote PDF and return its absolute path."""
    tx_num = tx.get("transaction_number") or "quote"
    safe_num = re.sub(r"[^A-Za-z0-9_.-]+", "_", tx_num)
    out_dir = Path(LOCAL_DATA_DIR) / "customer_documents" / "quotes"
    out_dir.mkdir(parents=True, exist_ok=True)
    out_path = str(out_dir / f"{safe_num}.pdf")

    styles = _styles()
    story: list = []

    sender = config.shipping.sender if config.shipping else None
    store_name = sender.name if sender else "Scarlett Music"
    company = sender.company if sender and sender.company else store_name
    store_lines = []
    if sender:
        store_lines.extend([
            sender.street1,
            sender.street2,
            _join_parts(sender.city, sender.state, sender.postcode),
            sender.phone,
            sender.email,
        ])
    store_lines = [line for line in store_lines if line]

    doc_title = "QUOTE"
    doc_no = _quote_display_number(tx)
    doc_date = _fmt_date(tx.get("created_at") or tx.get("completed_at"))

    header = Table(
        [
            [
                Paragraph(f"<b>{escape(store_name)}</b>", styles["brand"]),
                Paragraph(f"<b>{doc_title}</b><br/><font size='10'>{escape(doc_no)}</font>", styles["doc_title"]),
            ],
            [
                Paragraph("<br/>".join(escape(line) for line in store_lines), styles["small"]),
                _metadata_table(tx, doc_date, styles),
            ],
        ],
        colWidths=[_BODY_W * 0.58, _BODY_W * 0.42],
    )
    header.setStyle(TableStyle([
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 8),
    ]))
    story.append(header)
    story.append(Spacer(1, 10))

    story.append(Table(
        [[
            _party_block("From", company, store_lines, styles),
            _party_block("Quote To", _customer_name(customer), _customer_lines(customer), styles),
        ]],
        colWidths=[_BODY_W * 0.48, _BODY_W * 0.52],
        style=[
            ("VALIGN", (0, 0), (-1, -1), "TOP"),
            ("BOX", (0, 0), (-1, -1), 0.5, colors.HexColor("#d0d5dd")),
            ("INNERGRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#d0d5dd")),
            ("BACKGROUND", (0, 0), (-1, -1), colors.HexColor("#f8fafc")),
            ("PADDING", (0, 0), (-1, -1), 8),
        ],
    ))
    story.append(Spacer(1, 14))

    line_rows = [[
        Paragraph("SKU", styles["table_header"]),
        Paragraph("Description", styles["table_header"]),
        Paragraph("Qty", styles["table_header_right"]),
        Paragraph("Unit", styles["table_header_right"]),
        Paragraph("Disc", styles["table_header_right"]),
        Paragraph("Line Total", styles["table_header_right"]),
    ]]
    for line in cart_items.values():
        qty = float(line.get("qty") or 0)
        unit_price = float(line.get("unit_price") or 0)
        disc_pct = float(line.get("disc_pct") or 0)
        line_total = qty * unit_price * (1 - disc_pct / 100)
        line_rows.append([
            Paragraph(escape(str(line.get("sku") or "")), styles["cell"]),
            Paragraph(escape(str(line.get("title") or "")), styles["cell"]),
            Paragraph(_fmt_qty(qty), styles["cell_right"]),
            Paragraph(_fmt_money(unit_price), styles["cell_right"]),
            Paragraph(f"{disc_pct:.2f}%" if disc_pct else "-", styles["cell_right"]),
            Paragraph(_fmt_money(line_total), styles["cell_right"]),
        ])

    line_table = Table(
        line_rows,
        colWidths=[
            _BODY_W * 0.15,
            _BODY_W * 0.38,
            _BODY_W * 0.08,
            _BODY_W * 0.12,
            _BODY_W * 0.10,
            _BODY_W * 0.17,
        ],
        repeatRows=1,
    )
    line_table.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#1f2937")),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("GRID", (0, 0), (-1, -1), 0.35, colors.HexColor("#d0d5dd")),
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#f8fafc")]),
        ("TOPPADDING", (0, 0), (-1, -1), 5),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
    ]))
    story.append(line_table)
    story.append(Spacer(1, 12))

    subtotal = float(tx.get("subtotal") or 0)
    cart_discount = float(tx.get("cart_discount_total") or 0)
    total = float(tx.get("total") or 0)
    gst = round(total / 11, 2) if total else 0.0

    totals_rows = [
        ["Subtotal", _fmt_money(subtotal)],
    ]
    if cart_discount:
        totals_rows.append(["Cart discount", f"-{_fmt_money(cart_discount)}"])
    totals_rows.extend([
        ["GST included", _fmt_money(gst)],
        ["Total", _fmt_money(total)],
    ])
    totals = Table(
        [[Paragraph(label, styles["total_label"]), Paragraph(value, styles["total_value"])]
         for label, value in totals_rows],
        colWidths=[_BODY_W * 0.20, _BODY_W * 0.18],
        hAlign="RIGHT",
    )
    totals.setStyle(TableStyle([
        ("LINEABOVE", (0, -1), (-1, -1), 1.0, colors.black),
        ("TOPPADDING", (0, 0), (-1, -1), 4),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
    ]))
    story.append(totals)
    story.append(Spacer(1, 16))

    doc_notes = (notes or tx.get("notes") or "").strip()
    if doc_notes:
        story.append(Paragraph("Notes", styles["section"]))
        story.append(Paragraph(escape(doc_notes).replace("\n", "<br/>"), styles["body"]))
        story.append(Spacer(1, 10))

    story.append(Paragraph("Quote Terms", styles["section"]))
    story.append(Paragraph(
        "This quote is valid for 14 days unless otherwise stated. Prices include GST. "
        "Stock availability and pricing may change until the quote is accepted.",
        styles["body"],
    ))

    doc = SimpleDocTemplate(
        out_path,
        pagesize=A4,
        leftMargin=_MARGIN,
        rightMargin=_MARGIN,
        topMargin=14 * mm,
        bottomMargin=14 * mm,
        title=doc_no,
    )
    doc.build(story)
    return out_path


def _styles() -> dict[str, ParagraphStyle]:
    base = dict(fontName="Helvetica")
    bold = dict(fontName="Helvetica-Bold")
    return {
        "brand": ParagraphStyle("brand", fontSize=18, textColor=colors.HexColor("#111827"), **bold),
        "doc_title": ParagraphStyle("doc_title", fontSize=24, alignment=TA_RIGHT, textColor=colors.HexColor("#111827"), leading=28, **bold),
        "small": ParagraphStyle("small", fontSize=9, textColor=colors.HexColor("#475467"), **base),
        "label": ParagraphStyle("label", fontSize=8, textColor=colors.HexColor("#667085"), leading=10, **bold),
        "body": ParagraphStyle("body", fontSize=9, textColor=colors.HexColor("#344054"), **base),
        "section": ParagraphStyle("section", fontSize=10, textColor=colors.HexColor("#111827"), spaceAfter=4, **bold),
        "table_header": ParagraphStyle("table_header", fontSize=8, textColor=colors.white, **bold),
        "table_header_right": ParagraphStyle("table_header_right", fontSize=8, alignment=TA_RIGHT, textColor=colors.white, **bold),
        "cell": ParagraphStyle("cell", fontSize=8, textColor=colors.HexColor("#111827"), **base),
        "cell_right": ParagraphStyle("cell_right", fontSize=8, alignment=TA_RIGHT, textColor=colors.HexColor("#111827"), **base),
        "total_label": ParagraphStyle("total_label", fontSize=9, alignment=TA_LEFT, **base),
        "total_value": ParagraphStyle("total_value", fontSize=9, alignment=TA_RIGHT, **bold),
    }


def _metadata_table(tx: dict, doc_date: str, styles: dict[str, ParagraphStyle]) -> Table:
    rows = [
        ["Quote date", doc_date],
        ["Reference", tx.get("transaction_number") or ""],
    ]
    if tx.get("performed_by"):
        rows.append(["Prepared by", tx.get("performed_by") or ""])
    table = Table(
        [[Paragraph(escape(str(a)), styles["label"]), Paragraph(escape(str(b)), styles["body"])]
         for a, b in rows],
        colWidths=[34 * mm, 44 * mm],
        hAlign="RIGHT",
    )
    table.setStyle(TableStyle([
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
    ]))
    return table


def _party_block(
    label: str,
    name: str,
    lines: list[str],
    styles: dict[str, ParagraphStyle],
) -> Paragraph:
    body = [f"<font size='8' color='#667085'><b>{escape(label.upper())}</b></font>"]
    if name:
        body.append(f"<b>{escape(name)}</b>")
    body.extend(escape(line) for line in lines if line)
    return Paragraph("<br/>".join(body), styles["body"])


def _customer_name(customer: Optional[dict]) -> str:
    if not customer:
        return ""
    personal = f"{customer.get('first_name') or ''} {customer.get('surname') or ''}".strip()
    business = customer.get("business") or ""
    return business or personal


def _customer_lines(customer: Optional[dict]) -> list[str]:
    if not customer:
        return []
    lines = []
    personal = f"{customer.get('first_name') or ''} {customer.get('surname') or ''}".strip()
    business = customer.get("business") or ""
    if business and personal:
        lines.append(personal)
    for key in ("address_1", "address_2"):
        if customer.get(key):
            lines.append(customer[key])
    city = _join_parts(customer.get("city"), customer.get("state"), customer.get("postcode"))
    if city:
        lines.append(city)
    phone = customer.get("mobile") or customer.get("phone_1")
    if phone:
        lines.append(phone)
    if customer.get("email"):
        lines.append(customer["email"])
    cid = customer.get("customer_id")
    if cid is not None:
        lines.append(f"Customer ID: {cid}")
    return lines


def _join_parts(*parts: Optional[str]) -> str:
    return " ".join(str(p).strip() for p in parts if p)


def _quote_display_number(tx: dict) -> str:
    if tx.get("quote_number"):
        return f"Q-{int(tx['quote_number']):05d}"
    return tx.get("transaction_number") or "Quote"


def _fmt_date(raw: Optional[str]) -> str:
    if not raw:
        return ""
    try:
        dt = datetime.fromisoformat(str(raw).replace("Z", "+00:00"))
        if dt.tzinfo is not None:
            dt = dt.astimezone(ZoneInfo("Australia/Melbourne"))
        return f"{dt.day:02d}/{dt.month:02d}/{dt.year}"
    except Exception:
        return str(raw)[:10]


def _fmt_money(amount: float) -> str:
    return f"${amount:,.2f}"


def _fmt_qty(qty: float) -> str:
    return str(int(qty)) if qty == int(qty) else f"{qty:.3f}".rstrip("0").rstrip(".")
