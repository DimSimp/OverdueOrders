from __future__ import annotations

import os


def generate_envelope_pdfs(
    orders: list,
    classifications: dict[str, str],
    output_dir: str,
    date_str: str,
) -> dict[str, str | None]:
    """
    Generate Minilopes and Devilopes address PDFs.

    orders           — list of NetoOrder / EbayOrder objects
    classifications  — {order_id: "minilope" | "devilope"}  (satchel entries ignored)
    output_dir       — directory to write PDFs into (created if absent)
    date_str         — "YYYY-MM-DD" used in filename

    Returns {"minilope": path_or_None, "devilope": path_or_None}.
    """
    try:
        from reportlab.lib.pagesizes import A5, landscape
        from reportlab.lib.units import mm
        from reportlab.pdfgen.canvas import Canvas
    except ImportError:
        raise RuntimeError(
            "reportlab is required for envelope PDF generation. "
            "Run: pip install reportlab"
        )

    os.makedirs(output_dir, exist_ok=True)

    # A5 landscape: 210 mm × 148 mm
    PAGE_W, PAGE_H = landscape(A5)
    MARGIN_X = 18 * mm
    MARGIN_Y = 12 * mm
    NAME_SIZE = 20
    ADDR_SIZE = 16
    SMALL_SIZE = 9
    LINE_SPACING = 1.45  # multiplier of font size

    def _address_lines(order) -> tuple[list[str], str]:
        """Return (address_lines, order_id). Imports inline to avoid circular deps."""
        from src.neto_client import NetoOrder
        from src.ebay_client import EbayOrder

        if isinstance(order, NetoOrder):
            name = f"{order.ship_first_name} {order.ship_last_name}".strip()
            company = order.ship_company or ""
            street1 = order.ship_street1 or ""
            street2 = order.ship_street2 or ""
            city_line = "  ".join(
                filter(None, [order.ship_city, order.ship_state, order.ship_postcode])
            )
            order_id = order.order_id
        else:  # EbayOrder
            name = order.ship_name or ""
            company = ""
            street1 = order.ship_street1 or ""
            street2 = order.ship_street2 or ""
            city_line = "  ".join(
                filter(None, [order.ship_city, order.ship_state, order.ship_postcode])
            )
            order_id = order.order_id

        lines = [l for l in [name, company, street1, street2, city_line] if l]
        return lines, order_id

    def _draw_page(c: Canvas, order) -> None:
        lines, order_id = _address_lines(order)

        # Calculate block height so we can vertically centre it
        total_h = 0.0
        for i, _ in enumerate(lines):
            sz = NAME_SIZE if i == 0 else ADDR_SIZE
            total_h += sz * (LINE_SPACING if i < len(lines) - 1 else 1.0)

        # y of first line's baseline (centred vertically with slight upward bias)
        y = (PAGE_H / 2) + (total_h / 2) - NAME_SIZE * 0.1

        for i, line in enumerate(lines):
            sz = NAME_SIZE if i == 0 else ADDR_SIZE
            font = "Helvetica-Bold" if i == 0 else "Helvetica"
            c.setFont(font, sz)
            c.drawString(MARGIN_X, y - sz, line)
            y -= sz * LINE_SPACING

        # Order number — small text in bottom-right corner
        c.setFont("Helvetica", SMALL_SIZE)
        c.drawRightString(PAGE_W - MARGIN_X, MARGIN_Y, f"Order: {order_id}")

    # Group orders by envelope type
    order_lookup = {o.order_id: o for o in orders}
    groups: dict[str, list] = {"minilope": [], "devilope": []}
    for oid, label in classifications.items():
        if label in groups and oid in order_lookup:
            groups[label].append(order_lookup[oid])

    paths: dict[str, str | None] = {"minilope": None, "devilope": None}

    for label, label_orders in groups.items():
        if not label_orders:
            continue
        type_name = "Minilopes" if label == "minilope" else "Devilopes"
        path = os.path.join(output_dir, f"{type_name}_{date_str}.pdf")
        c = Canvas(path, pagesize=(PAGE_W, PAGE_H))
        for i, order in enumerate(label_orders):
            _draw_page(c, order)
            if i < len(label_orders) - 1:
                c.showPage()
        c.save()
        paths[label] = path

    return paths
