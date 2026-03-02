from __future__ import annotations

import re
from dataclasses import dataclass, field

import pdfplumber

from src.config import SupplierConfig

# Default regex patterns used when supplier-specific ones are not configured
_DEFAULT_SKU_PATTERN = r"\b([A-Z][A-Z0-9\-]{2,19})\b"
_DEFAULT_QTY_PATTERN = r"(?:Qty|QTY|Quantity|Units)[:\s]+(\d+)"
_FALLBACK_QTY_IN_LINE = r"(?<!\$)(?<!\d\.)\b([1-9][0-9]{0,2})\b"

# Matches a decimal number like "74.00" or "317.55" — used to identify price tokens
_FLOAT_RE = re.compile(r"^\d+\.\d+$")

# D'Addario fixed-format line item regex.
# Format: SKU Description U/M QtyOrdered QtyShipped [QtyBackOrdered] RRP Disc% UnitPrice Amount
# U/M is strictly uppercase letters (EA, BX, PR, SET, etc.).
# Amount is ".00" for backordered items (those lines are skipped).
_DADDARIO_ITEM_RE = re.compile(
    r"^(\S+(?:\s+\d+/\S+)?)\s+"  # group 1: SKU (with optional fractional size like "1/2", "3/4", "1/4M")
    r"(.+?)\s+"           # group 2: Description (non-greedy — stops at first valid U/M)
    r"([A-Za-z]{1,6})\s+"  # group 3: U/M (case-insensitive — tolerates OCR misreads like "cA" for "EA")
    r"(\d+)\s+"           # group 4: QtyOrdered
    r"(\d+)\s+"           # group 5: QtyShipped
    r"(?:\d+\s+)?"        # optional QtyBackOrdered (uncaptured)
    r"\d+\.\d+\s+"        # RRP (uncaptured)
    r"\d+\.\d+%\s+"       # Disc% (uncaptured)
    r"\d+\.\d+\s+"        # UnitPrice (uncaptured)
    r"(\.00|\d+\.\d+)$"   # group 6: Amount (".00" = backordered)
)


@dataclass
class InvoiceItem:
    sku: str               # Raw SKU from PDF before suffix applied
    sku_with_suffix: str   # SKU after character substitutions + suffix
    description: str
    quantity: int
    source_page: int
    qty_flagged: bool = False  # True if qty could not be parsed (defaulted to 1)


class ParseError(Exception):
    """Raised when a PDF cannot be parsed for the selected supplier."""
    pass


def parse_invoice(pdf_path: str, supplier: SupplierConfig) -> list[InvoiceItem]:
    """
    Main entry point. Validates supplier, extracts items, applies suffix.

    Raises ParseError with a user-friendly message on any failure.
    """
    try:
        with pdfplumber.open(pdf_path) as pdf:
            full_text = "\n".join(
                page.extract_text() or "" for page in pdf.pages
            )

            # Validate supplier marker only when the PDF has embedded text.
            # Scanned PDFs (full_text empty) cannot be validated this way — trust
            # the user's supplier selection and let extraction succeed or fail.
            if supplier.validation_marker and full_text.strip():
                if not _validate_supplier(full_text, supplier.validation_marker):
                    raise ParseError(
                        f"This PDF does not appear to be a {supplier.name} invoice.\n\n"
                        f"Expected to find '{supplier.validation_marker}' in the document "
                        f"but it was not present.\n\n"
                        "Please check that you selected the correct supplier."
                    )

            items: list[InvoiceItem] = []

            if supplier.pdf_format == "marker":
                # Marker mode: work from the combined full-page text
                items.extend(_extract_by_markers(full_text, supplier, page_num=1))
            elif supplier.pdf_format == "daddario":
                for page_num, page in enumerate(pdf.pages, start=1):
                    text = page.extract_text() or ""
                    if not text.strip():
                        # Scanned page — fall back to Tesseract OCR
                        text = _ocr_page_to_text(pdf_path, page_num - 1)
                    items.extend(_extract_daddario(text, page_num))
            elif supplier.pdf_format == "table":
                for page_num, page in enumerate(pdf.pages, start=1):
                    page_items = _extract_from_table(page, supplier, page_num)
                    if not page_items:
                        text = page.extract_text() or ""
                        page_items = _extract_from_text(text, supplier, page_num)
                    items.extend(page_items)
            else:
                for page_num, page in enumerate(pdf.pages, start=1):
                    text = page.extract_text() or ""
                    items.extend(_extract_from_text(text, supplier, page_num))

    except ParseError:
        raise
    except Exception as e:
        raise ParseError(f"Failed to read PDF: {e}") from e

    if not items:
        raise ParseError(
            f"No items could be extracted from this PDF for supplier '{supplier.name}'.\n\n"
            "The PDF may use an unsupported format. Try adjusting the column hints in config.json, "
            "or contact support."
        )

    # Apply character substitutions and suffix to produce sku_with_suffix
    for item in items:
        item.sku_with_suffix = _build_neto_sku(item.sku, supplier)

    return items


def _validate_supplier(full_text: str, marker: str) -> bool:
    return marker.lower() in full_text.lower()


def _build_neto_sku(raw_sku: str, supplier: SupplierConfig) -> str:
    """Apply character substitutions then append/prepend suffix."""
    sku = raw_sku
    for char, replacement in supplier.character_substitutions.items():
        sku = sku.replace(char, replacement)
    if not supplier.suffix:
        return sku
    if supplier.suffix_position == "prepend":
        return f"{supplier.suffix}{sku}"
    return f"{sku}{supplier.suffix}"


def _extract_from_table(
    page, supplier: SupplierConfig, page_num: int
) -> list[InvoiceItem]:
    """
    Use pdfplumber table extraction. Finds header row by matching
    the supplier's column hints (case-insensitive substring).
    """
    tables = page.extract_tables()
    if not tables:
        return []

    items = []
    for table in tables:
        result = _parse_table(table, supplier, page_num)
        items.extend(result)

    return items


def _parse_table(
    table: list[list], supplier: SupplierConfig, page_num: int
) -> list[InvoiceItem]:
    """Find the header row in a table and extract items from data rows."""
    header_result = _find_header_row(table, supplier)
    if header_result is None:
        return []

    header_idx, col_map = header_result
    items = []

    for row in table[header_idx + 1:]:
        if row is None:
            continue
        # Clean up cells: pdfplumber may return None for empty cells
        cells = [str(c).strip() if c is not None else "" for c in row]

        sku_idx = col_map.get("sku")
        desc_idx = col_map.get("desc")
        qty_idx = col_map.get("qty")

        sku = cells[sku_idx].strip() if sku_idx is not None and sku_idx < len(cells) else ""
        desc = cells[desc_idx].strip() if desc_idx is not None and desc_idx < len(cells) else ""
        qty_raw = cells[qty_idx].strip() if qty_idx is not None and qty_idx < len(cells) else ""

        # Skip blank or header-looking rows
        if not sku or sku.lower() in ("sku", "code", "part no", "part", "item", "product"):
            continue

        # Skip rows that look like totals or section headers (no alphanumeric SKU)
        if not re.search(r"[A-Za-z0-9]", sku):
            continue

        qty, qty_flagged = _parse_qty(qty_raw)

        items.append(InvoiceItem(
            sku=sku,
            sku_with_suffix=sku,  # will be set in parse_invoice()
            description=desc,
            quantity=qty,
            source_page=page_num,
            qty_flagged=qty_flagged,
        ))

    return items


def _find_header_row(
    table: list[list], supplier: SupplierConfig
) -> tuple[int, dict] | None:
    """
    Scan table rows for a header row matching the supplier's column hints.
    Returns (row_index, {sku: col_idx, desc: col_idx, qty: col_idx}) or None.
    """
    sku_hint = supplier.sku_column_hint.lower()
    qty_hint = supplier.qty_column_hint.lower()
    desc_hint = supplier.desc_column_hint.lower()

    for row_idx, row in enumerate(table):
        if row is None:
            continue
        cells = [str(c).lower().strip() if c is not None else "" for c in row]

        col_map = {}
        for col_idx, cell in enumerate(cells):
            if sku_hint and sku_hint in cell:
                col_map.setdefault("sku", col_idx)
            if desc_hint and desc_hint in cell:
                col_map.setdefault("desc", col_idx)
            if qty_hint and qty_hint in cell:
                col_map.setdefault("qty", col_idx)

        if "sku" in col_map:
            return row_idx, col_map

    # Fallback: look for generic SKU column headers if specific hints not found
    generic_sku_hints = ["sku", "code", "part no", "part", "item code", "product code"]
    generic_qty_hints = ["qty", "quantity", "units", "ordered"]
    generic_desc_hints = ["description", "product", "item", "name"]

    for row_idx, row in enumerate(table):
        if row is None:
            continue
        cells = [str(c).lower().strip() if c is not None else "" for c in row]

        col_map = {}
        for col_idx, cell in enumerate(cells):
            if not cell:
                continue
            if any(h in cell for h in generic_sku_hints):
                col_map.setdefault("sku", col_idx)
            if any(h in cell for h in generic_desc_hints):
                col_map.setdefault("desc", col_idx)
            if any(h in cell for h in generic_qty_hints):
                col_map.setdefault("qty", col_idx)

        if "sku" in col_map:
            return row_idx, col_map

    return None


def _extract_from_text(
    text: str, supplier: SupplierConfig, page_num: int
) -> list[InvoiceItem]:
    """
    Regex-based extraction for text-only (non-table) PDFs.
    Scans line by line, matching SKU pattern and qty pattern.
    """
    sku_pattern = supplier.sku_pattern or _DEFAULT_SKU_PATTERN
    qty_pattern = supplier.qty_pattern or _DEFAULT_QTY_PATTERN

    lines = [line.strip() for line in text.splitlines() if line.strip()]
    items = []

    i = 0
    while i < len(lines):
        line = lines[i]
        sku_match = re.search(sku_pattern, line)
        if not sku_match:
            i += 1
            continue

        sku = sku_match.group(1) if sku_match.lastindex else sku_match.group(0)

        # Try to find qty on this line or the next 2
        qty = 1
        qty_flagged = True
        context = " ".join(lines[i: i + 3])
        qty_match = re.search(qty_pattern, context, re.IGNORECASE)
        if qty_match:
            qty, qty_flagged = _parse_qty(qty_match.group(1))

        # Description: everything on the sku line after the SKU match
        sku_end = sku_match.end()
        desc_raw = line[sku_end:].strip()
        # Clean up leading punctuation/separators
        desc_raw = re.sub(r"^[\s\-:,|]+", "", desc_raw)
        description = desc_raw or sku  # fallback to SKU if no description found

        items.append(InvoiceItem(
            sku=sku,
            sku_with_suffix=sku,
            description=description,
            quantity=qty,
            source_page=page_num,
            qty_flagged=qty_flagged,
        ))
        i += 1

    return items


def _strip_page_breaks(text: str, start_marker: str) -> str:
    """
    Remove inter-page content from multi-page marker-format invoices.

    pdfplumber produces per-page text in this order:
        [header block]
        [items — column header "No. Description ..." only on page 1]
        Continued...  [running totals]
        [footer: AUSTRALIS MUSIC GROUP, address, ABN, page N of M]

    On pages 2+, the header block is:
        TAX INVOICE
        [invoice number]
        [Sell To / Deliver To addresses]
        Customer No. Customer PO Document Date Order No. Salesperson Payment Terms Due Date
        301111        7031        28/04/25      SO604042   VIC3        30 DAYS FROM EOM  31/05/25

    The values row (ending with the due date in dd/mm/yy format) is the LAST line
    of inter-page content. Items start on the very next line.

    Strategy: after each "Continued...", find the values row by matching a line that
    ends with a date token (dd/mm/yy or dd/mm/yyyy). Skip to the line after it.
    If no date line is found, fall back to looking for start_marker (for suppliers
    whose subsequent pages repeat the column headers), then truncate.

    Handles 3+ page invoices by looping.
    """
    # Matches a line whose last non-whitespace token is a date like 31/05/25 or 31/05/2025
    _due_date_line_re = re.compile(r"\d{2}/\d{2}/\d{2,4}\s*$", re.MULTILINE)
    marker_lower = start_marker.lower()
    parts = []
    remaining = text

    while True:
        cont_match = re.search(r"(?im)^Continued", remaining)
        if not cont_match:
            parts.append(remaining)
            break

        # Keep everything before "Continued..."
        parts.append(remaining[:cont_match.start()])
        after_cont = remaining[cont_match.start():]

        # Primary: find the header values row (ends with the due date)
        date_match = _due_date_line_re.search(after_cont)
        if date_match is not None:
            # Skip to the start of the line after the due-date line
            line_end = after_cont.find("\n", date_match.end())
            skip_to = (line_end + 1) if line_end != -1 else len(after_cont)
        else:
            # Fallback: look for a repeated column-header line (start_marker)
            next_marker_pos = after_cont.lower().find(marker_lower)
            if next_marker_pos == -1:
                # No more item sections — discard the rest
                break
            newline_pos = after_cont.find("\n", next_marker_pos)
            skip_to = (newline_pos + 1) if newline_pos != -1 else (next_marker_pos + len(start_marker))

        remaining = after_cont[skip_to:]

    return "".join(parts)


def _extract_by_markers(
    text: str, supplier: SupplierConfig, page_num: int
) -> list[InvoiceItem]:
    """
    Marker-based extraction for invoices whose text follows a column-based pattern.

    Handles three complications that arise from pdfplumber's left-to-right reading:

    1. Variable trailing numeric columns: items may have 3 or 4 price columns
       depending on whether a discount applies. We peel trailing float-tokens from
       the right of each SKU line rather than counting a fixed number of fields.

    2. Multi-line descriptions: when a description is too long to fit, pdfplumber
       extracts it split across lines with the SKU line in the middle, e.g.:
           "GHS ML7200 (44-102) PRESSUREWOUND"   ← prefix
           "751207 2 89.00 56.64 20.00 90.62"    ← SKU + data
           "BASS STRINGS"                          ← suffix
       We reassemble these by classifying each line as SKU line or description
       line, then applying a rule to assign description lines as prefix/suffix.

    3. Multi-page invoices: a "Continued..." line marks the end of items on a page,
       followed by footers and a repeated column header on the next page. We strip
       this inter-page content before line classification.

    Assignment rule for non-SKU lines between two consecutive SKU items A and B:
    - If B has NO inline description (all description is in adjacent lines):
        the LAST non-SKU line before B is B's prefix; remaining lines are A's suffix.
    - If B HAS an inline description:
        all non-SKU lines in the block are A's suffix.
    """
    start_marker = supplier.item_start_marker
    end_marker = supplier.item_end_marker
    min_digits = supplier.sku_min_digits

    # --- Slice to the items section ---
    start_idx = 0
    if start_marker:
        pos = text.find(start_marker)
        if pos == -1:
            pos = text.lower().find(start_marker.lower())
        if pos != -1:
            newline_pos = text.find("\n", pos)
            start_idx = newline_pos + 1 if newline_pos != -1 else pos + len(start_marker)

    end_idx = len(text)
    if end_marker:
        pos = text.find(end_marker, start_idx)
        if pos == -1:
            pos = text.lower().find(end_marker.lower(), start_idx)
        if pos != -1:
            end_idx = pos

    item_section = text[start_idx:end_idx]
    if not item_section.strip():
        return []

    # --- Strip inter-page content ("Continued..." + footer + repeated header) ---
    if start_marker and re.search(r"(?im)^Continued", item_section):
        item_section = _strip_page_breaks(item_section, start_marker)

    # --- Classify lines as SKU lines or description lines ---
    def is_sku_line(line: str) -> bool:
        tokens = line.split()
        return bool(tokens) and tokens[0].isdigit() and len(tokens[0]) >= min_digits

    parsed = []  # ("sku", sku, middle_desc, qty, flagged) | ("desc", text)
    for line in (l.strip() for l in item_section.splitlines() if l.strip()):
        if is_sku_line(line):
            tokens = line.split()
            sku = tokens[0]
            middle_tokens, qty_token = _parse_rest_tokens(tokens[1:])
            middle_desc = " ".join(middle_tokens)
            qty, flagged = _parse_qty(qty_token)
            parsed.append(("sku", sku, middle_desc, qty, flagged))
        else:
            parsed.append(("desc", line))

    # Extract ordered list of SKU items: (parsed_index, sku, middle_desc, qty, flagged)
    sku_items = [
        (i, p[1], p[2], p[3], p[4])
        for i, p in enumerate(parsed)
        if p[0] == "sku"
    ]
    if not sku_items:
        return []

    # --- Assign prefix / suffix lines to each SKU item ---
    prefix_for: dict[int, list[str]] = {pos: [] for pos, *_ in sku_items}
    suffix_for: dict[int, list[str]] = {pos: [] for pos, *_ in sku_items}

    # Non-SKU lines before the very first SKU line
    first_pos, _, first_middle, _, _ = sku_items[0]
    before_first = [parsed[j][1] for j in range(0, first_pos) if parsed[j][0] == "desc"]
    if not first_middle:
        prefix_for[first_pos] = before_first

    # Process the block of non-SKU lines that falls between each pair of SKU lines
    for idx, (sku_pos, _sku, middle_desc, _qty, _flag) in enumerate(sku_items):
        if idx + 1 < len(sku_items):
            next_pos, _, next_middle, _, _ = sku_items[idx + 1]
        else:
            next_pos = len(parsed)
            next_middle = "_"  # sentinel: treat all trailing lines as suffix

        after_block = [
            parsed[j][1] for j in range(sku_pos + 1, next_pos)
            if parsed[j][0] == "desc"
        ]

        if not next_middle and after_block:
            # Next SKU has no inline description → its last preceding line is its prefix
            prefix_for[sku_items[idx + 1][0]] = [after_block[-1]]
            suffix_for[sku_pos] = after_block[:-1]
        else:
            # Next SKU has its own inline description → all lines belong to current suffix
            suffix_for[sku_pos] = after_block

    # --- Build InvoiceItem list ---
    items = []
    for sku_pos, sku, middle_desc, qty, flagged in sku_items:
        if middle_desc:
            description = middle_desc
        else:
            parts = prefix_for.get(sku_pos, []) + suffix_for.get(sku_pos, [])
            description = " ".join(parts).strip() or sku

        items.append(InvoiceItem(
            sku=sku,
            sku_with_suffix=sku,
            description=description,
            quantity=qty,
            source_page=page_num,
            qty_flagged=flagged,
        ))

    return items


def _parse_rest_tokens(tokens: list[str]) -> tuple[list[str], str]:
    """
    Given the tokens after the SKU on a SKU line, separate the inline description
    from the trailing price columns and qty.

    Strategy: peel float-formatted tokens (e.g. "74.00") from the right — these
    are the price columns (RRP, unit price, discount %, amount). The last remaining
    non-float token is the quantity integer. Everything before that is the inline
    description (may be empty if the description is on adjacent lines).

    Returns (middle_desc_tokens, qty_token_string).
    """
    if not tokens:
        return [], "1"

    i = len(tokens) - 1
    while i >= 0 and _FLOAT_RE.match(tokens[i]):
        i -= 1

    if i < 0:
        # All tokens were floats — no qty found
        return [], "1"

    return tokens[:i], tokens[i]


def _parse_qty(raw: str) -> tuple[int, bool]:
    """
    Parse a quantity string to int.
    Returns (qty, flagged) where flagged=True means qty defaulted to 1.
    """
    if not raw:
        return 1, True
    # Extract first integer from the string
    match = re.search(r"\d+", raw.replace(",", ""))
    if match:
        val = int(match.group())
        return (val if val > 0 else 1), False
    return 1, True


def _ocr_page_to_text(pdf_path: str, page_index: int) -> str:
    """
    Render a single PDF page to a high-resolution image and extract text via
    Tesseract OCR. Used as a fallback when a page contains no embedded text
    (i.e. it is a scanned image rather than a digital PDF).

    Requires:
        pip install pymupdf pytesseract
        Tesseract OCR installed from https://github.com/UB-Mannheim/tesseract/wiki
    """
    try:
        import fitz  # PyMuPDF
    except ImportError:
        raise ParseError(
            "This PDF appears to be scanned (no embedded text) and requires OCR.\n\n"
            "Please install the required packages:\n"
            "  pip install pymupdf pytesseract\n\n"
            "Tesseract OCR must also be installed from:\n"
            "  https://github.com/UB-Mannheim/tesseract/wiki"
        )

    try:
        import io
        import pytesseract
        from PIL import Image
    except ImportError:
        raise ParseError(
            "OCR support requires: pip install pytesseract Pillow"
        )

    # On Windows, Tesseract is often installed but not on PATH.
    # Try common install locations so users don't need to edit PATH manually.
    if pytesseract.pytesseract.tesseract_cmd == "tesseract":
        import os
        _common_paths = [
            r"C:\Program Files\Tesseract-OCR\tesseract.exe",
            r"C:\Program Files (x86)\Tesseract-OCR\tesseract.exe",
            os.path.expanduser(r"~\AppData\Local\Programs\Tesseract-OCR\tesseract.exe"),
        ]
        for candidate in _common_paths:
            if os.path.isfile(candidate):
                pytesseract.pytesseract.tesseract_cmd = candidate
                break

    try:
        doc = fitz.open(pdf_path)
        page = doc.load_page(page_index)
        # 300 DPI gives reliable text recognition for printed invoices
        mat = fitz.Matrix(300 / 72, 300 / 72)
        pix = page.get_pixmap(matrix=mat)
        img = Image.open(io.BytesIO(pix.tobytes("png")))
        # PSM 6: assume a uniform block of text — suits invoice page layout
        return pytesseract.image_to_string(img, config="--psm 6")
    except ParseError:
        raise
    except pytesseract.TesseractNotFoundError:
        raise ParseError(
            "Tesseract OCR is not installed or could not be found.\n\n"
            "Please install it from:\n"
            "  https://github.com/UB-Mannheim/tesseract/wiki\n\n"
            "After installing, restart the application."
        )
    except Exception as e:
        raise ParseError(f"OCR failed for page {page_index + 1}: {e}") from e


def _extract_daddario(text: str, page_num: int) -> list[InvoiceItem]:
    """
    D'Addario-specific parser for their fixed-column text invoices.

    Column order per line:
        SKU  Description  U/M  QtyOrdered  QtyShipped  [QtyBackOrdered]  RRP  Disc%  UnitPrice  Amount

    Items start on the line after the "Item Number" column header and end
    before the legal disclaimer ("It is expressly agreed").
    Backordered lines (Amount == ".00") and zero-shipped lines are skipped.
    QtyShipped is used as the item quantity.
    """
    # Normalise OCR artefacts that appear in scanned D'Addario invoices:
    #   • Pipe characters (Tesseract reads table column rules as "|") → space
    #   • Bracket characters ("[", "]", "{", "}") next to letters → removed
    text = re.sub(r"[ \t]*\|[ \t]*", " ", text)
    text = re.sub(r"[\[\]{}]", "", text)
    text = re.sub(r"[ \t]{2,}", " ", text)  # compact runs of spaces

    lines = text.splitlines()

    # Find the column header line — items begin on the very next line.
    # Falls back to scanning from line 0 when the header is garbled (OCR'd pages).
    start_idx = 0
    for i, line in enumerate(lines):
        if "Item Number" in line:
            start_idx = i + 1
            break

    # Items end at the legal disclaimer or totals line
    end_idx = len(lines)
    for i in range(start_idx, len(lines)):
        stripped = lines[i].strip()
        if stripped.startswith("It is expressly") or stripped.startswith("Total Include GST"):
            end_idx = i
            break

    items = []
    for line in lines[start_idx:end_idx]:
        line = line.strip()
        if not line:
            continue
        m = _DADDARIO_ITEM_RE.match(line)
        if not m:
            continue

        sku = m.group(1)
        description = m.group(2)
        qty_shipped = int(m.group(5))
        amount = m.group(6)

        # Skip backordered / unshipped items
        if amount == ".00" or qty_shipped == 0:
            continue

        items.append(InvoiceItem(
            sku=sku,
            sku_with_suffix=sku,  # set in parse_invoice() via _build_neto_sku()
            description=description,
            quantity=qty_shipped,
            source_page=page_num,
            qty_flagged=False,
        ))

    return items
