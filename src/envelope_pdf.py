from __future__ import annotations

import os


def generate_envelope_pdfs(
    orders: list,
    classifications: dict[str, str],
    output_dir: str,
) -> dict[str, str | None]:
    """
    Generate Minilopes and Devilopes address PDFs.

    orders           — list of NetoOrder / EbayOrder objects
    classifications  — {order_id: "minilope" | "devilope"}  (satchel entries ignored)
    output_dir       — directory to write PDFs into (created if absent)

    Returns {"minilope": path_or_None, "devilope": path_or_None}.
    """
    try:
        from reportlab.lib.pagesizes import A5
        from reportlab.lib.units import mm
        from reportlab.pdfgen.canvas import Canvas
    except ImportError:
        raise RuntimeError(
            "reportlab is required for envelope PDF generation. "
            "Run: pip install reportlab"
        )

    os.makedirs(output_dir, exist_ok=True)

    # A5 portrait page (148 mm × 210 mm).  Content is drawn rotated 90° CW
    # so that when the envelope is fed portrait into the printer the address
    # reads correctly (i.e. the text is landscape relative to the page).
    PAGE_W, PAGE_H = A5          # 148 mm × 210 mm  (portrait)
    # Drawing coords after the CW-90° transform act as a landscape surface:
    DRAW_W = PAGE_H              # 210 mm — horizontal drawing extent
    DRAW_H = PAGE_W              # 148 mm — vertical drawing extent
    MARGIN_X = 18 * mm
    MARGIN_Y = 12 * mm
    NAME_SIZE = 20
    ADDR_SIZE = 16
    POSTCODE_SIZE = 28
    SMALL_SIZE = 9
    LINE_SPACING = 1.45  # multiplier of font size

    def _address_lines(order) -> tuple[list[str], str, str]:
        """Return (address_lines_without_postcode, postcode, order_id)."""
        from src.neto_client import NetoOrder

        if isinstance(order, NetoOrder):
            name = f"{order.ship_first_name} {order.ship_last_name}".strip()
            company = order.ship_company or ""
            street1 = order.ship_street1 or ""
            street2 = order.ship_street2 or ""
            postcode = order.ship_postcode or ""
            city_state = "  ".join(filter(None, [order.ship_city, order.ship_state]))
            order_id = order.order_id
        else:  # EbayOrder
            name = order.ship_name or ""
            company = ""
            street1 = order.ship_street1 or ""
            street2 = order.ship_street2 or ""
            postcode = order.ship_postcode or ""
            city_state = "  ".join(filter(None, [order.ship_city, order.ship_state]))
            order_id = order.order_id

        lines = [l for l in [name, company, street1, street2, city_state] if l]
        return lines, postcode, order_id

    def _draw_page(c: Canvas, order) -> None:
        lines, postcode, order_id = _address_lines(order)

        # Rotate coordinate system 90° CW so we can draw as if the page is
        # landscape (DRAW_W × DRAW_H).
        c.saveState()
        c.translate(0, PAGE_H)
        c.rotate(-90)

        # Now draw within DRAW_W × DRAW_H landscape coordinates -----------

        # Fixed y positions for the bottom-right annotation block.
        # ORDER_Y ≈ "slightly higher than where the postcode used to sit"
        # POSTCODE_Y ≈ above order number, moved well up from the very bottom
        ORDER_Y    = MARGIN_Y + SMALL_SIZE + 30      # bumped up slightly more
        POSTCODE_Y = MARGIN_Y + SMALL_SIZE + 65      # bumped up slightly

        # Reserve space so the address block doesn't overlap the postcode.
        usable_bottom = POSTCODE_Y + POSTCODE_SIZE + 6
        usable_top    = DRAW_H - MARGIN_Y - 20       # bias the band downward

        # Calculate block height for vertical centering within usable band
        total_h = 0.0
        for i, _ in enumerate(lines):
            sz = NAME_SIZE if i == 0 else ADDR_SIZE
            total_h += sz * (LINE_SPACING if i < len(lines) - 1 else 1.0)

        usable_mid = (usable_top + usable_bottom) / 2
        y = usable_mid + (total_h / 2)

        for i, line in enumerate(lines):
            sz = NAME_SIZE if i == 0 else ADDR_SIZE
            font = "Helvetica-Bold" if i == 0 else "Helvetica"
            c.setFont(font, sz)
            c.drawString(MARGIN_X, y - sz, line)   # left-aligned
            y -= sz * LINE_SPACING

        # Postcode — large bold, bottom-right
        if postcode:
            c.setFont("Helvetica-Bold", POSTCODE_SIZE)
            c.drawRightString(DRAW_W - MARGIN_X, POSTCODE_Y, postcode)

        # Order number — small grey, bottom-right below postcode
        c.setFont("Helvetica", SMALL_SIZE)
        c.setFillColorRGB(0.6, 0.6, 0.6)
        c.drawRightString(DRAW_W - MARGIN_X, ORDER_Y, order_id)
        c.setFillColorRGB(0, 0, 0)  # reset fill colour

        c.restoreState()

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
        path = os.path.join(output_dir, f"{type_name}.pdf")
        c = Canvas(path, pagesize=A5)  # portrait 148 mm × 210 mm
        for i, order in enumerate(label_orders):
            _draw_page(c, order)
            if i < len(label_orders) - 1:
                c.showPage()
        c.save()
        paths[label] = path

    return paths
