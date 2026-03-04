"""
OpenAI Vision fallback for extracting invoice items from scanned/image PDF pages.

Used when the standard pdfplumber/regex parsing returns no items — typically
because the PDF is a scan of a printed invoice rather than a digital PDF.
"""

from __future__ import annotations

import base64
import io
import json

import fitz  # PyMuPDF

from src.config import OpenAIConfig, SupplierConfig
from src.pdf_parser import InvoiceItem, ParseError


_SYSTEM_PROMPT = """\
You are an invoice data extraction assistant. You will be shown an image of a \
supplier invoice page. Extract every line item from the invoice and return them \
as a JSON array.

Each item must have these fields:
- "sku": the product SKU / item code / part number (string)
- "description": the product description (string)
- "quantity": the quantity shipped/supplied as an integer (NOT qty ordered or backordered)

Rules:
- Return ONLY the JSON array, no markdown fences or commentary.
- If a line item has both "Qty Ordered" and "Qty Shipped", use Qty Shipped.
- Skip backordered items (quantity 0 or amount $0.00).
- Skip subtotal, total, freight, and tax lines.
- If you cannot read a SKU or quantity clearly, make your best effort.
- Return an empty array [] if no line items are found.
"""


def _build_user_prompt(supplier: SupplierConfig) -> str:
    """Build a supplier-specific hint to include with the image."""
    parts = [f"This is an invoice from {supplier.name}."]

    if supplier.pdf_format == "daddario":
        parts.append(
            "Column layout: Item Number, Description, U/M, QtyOrdered, "
            "QtyShipped, [QtyBackOrdered], RRP, Disc%, UnitPrice, Amount. "
            "Use the 'Item Number' column as the SKU. "
            "Use QtyShipped as the quantity. "
            "Skip items where Amount is .00 or QtyShipped is 0."
        )
    elif supplier.name.startswith("AMS"):
        parts.append(
            "Column layout: Model No, Description, SUPP, B/O (and possibly other columns). "
            "'Model No' is the SKU. "
            "'SUPP' is the quantity supplied (received) — use this as the quantity. "
            "'B/O' is the backordered quantity — ignore it entirely. "
            "Stop extracting items when you encounter a line starting with "
            "'Packed by' or 'Collect' — those mark the end of the invoice."
        )

    parts.append("Extract all line items from this invoice page.")
    return " ".join(parts)


def _render_page_to_png(pdf_path: str, page_index: int) -> bytes:
    """Render a single PDF page to a PNG image at 300 DPI."""
    doc = fitz.open(pdf_path)
    page = doc.load_page(page_index)
    mat = fitz.Matrix(300 / 72, 300 / 72)
    pix = page.get_pixmap(matrix=mat)
    return pix.tobytes("png")


def extract_items_with_ai(
    pdf_path: str,
    page_index: int,
    supplier: SupplierConfig,
    openai_config: OpenAIConfig,
) -> list[InvoiceItem]:
    """
    Send a PDF page image to OpenAI Vision and parse the response into
    InvoiceItem objects.

    Raises ParseError if the API call or response parsing fails.
    """
    try:
        from openai import OpenAI
    except ImportError:
        raise ParseError(
            "OpenAI fallback requires the openai package.\n\n"
            "Install it with: pip install openai"
        )

    if not openai_config.is_configured:
        raise ParseError(
            "OpenAI API key is not configured.\n\n"
            "Add your API key to the 'openai' section in config.json."
        )

    # Render the page to a PNG image
    try:
        png_bytes = _render_page_to_png(pdf_path, page_index)
    except Exception as e:
        raise ParseError(f"Failed to render PDF page {page_index + 1} for AI parsing: {e}") from e

    image_b64 = base64.b64encode(png_bytes).decode("utf-8")

    # Call OpenAI Vision API
    try:
        client = OpenAI(api_key=openai_config.api_key)
        response = client.chat.completions.create(
            model=openai_config.model,
            messages=[
                {"role": "system", "content": _SYSTEM_PROMPT},
                {
                    "role": "user",
                    "content": [
                        {"type": "text", "text": _build_user_prompt(supplier)},
                        {
                            "type": "image_url",
                            "image_url": {
                                "url": f"data:image/png;base64,{image_b64}",
                            },
                        },
                    ],
                },
            ],
            max_tokens=4096,
            temperature=0,
        )
    except Exception as e:
        raise ParseError(f"OpenAI API call failed: {e}") from e

    # Parse the response
    raw_text = response.choices[0].message.content.strip()

    # Strip markdown fences if the model included them despite instructions
    if raw_text.startswith("```"):
        lines = raw_text.split("\n")
        # Remove first line (```json) and last line (```)
        lines = [l for l in lines if not l.strip().startswith("```")]
        raw_text = "\n".join(lines)

    try:
        data = json.loads(raw_text)
    except json.JSONDecodeError as e:
        raise ParseError(
            f"OpenAI returned invalid JSON for page {page_index + 1}.\n\n"
            f"Response: {raw_text[:200]}"
        ) from e

    if not isinstance(data, list):
        raise ParseError(
            f"OpenAI returned unexpected format for page {page_index + 1}.\n\n"
            f"Expected a JSON array, got: {type(data).__name__}"
        )

    page_num = page_index + 1
    items = []
    for entry in data:
        sku = str(entry.get("sku", "")).strip()
        description = str(entry.get("description", "")).strip()
        qty_raw = entry.get("quantity", 1)

        if not sku:
            continue

        try:
            qty = max(1, int(qty_raw))
        except (ValueError, TypeError):
            qty = 1

        items.append(InvoiceItem(
            sku=sku,
            sku_with_suffix=sku,  # will be set by _build_neto_sku() in parse_invoice()
            description=description or sku,
            quantity=qty,
            source_page=page_num,
            qty_flagged=False,
        ))

    return items
