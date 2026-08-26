"""Purchase-order text -> typed records.

Pure: text in, dataclasses out. No file I/O, no openpyxl.

The item table's column geometry is *derived from the document* and not hardcoded. 

*See `derive_layout` for the rule and why the obvious rules don't work.*
"""

from __future__ import annotations

import re
from dataclasses import dataclass, field
from datetime import date
from decimal import Decimal, InvalidOperation

__all__ = [
    "ParseError",
    "Column",
    "Layout",
    "LineItem",
    "Header",
    "PurchaseOrder",
    "parse_po",
    "derive_layout",
    "parse_decimal",
    "parse_date",
]


class ParseError(ValueError):
    """Raised when a line cannot be understood. Names file and line."""

    def __init__(self, message: str, *, filename: str, line_no: int | None = None):
        where = filename if line_no is None else f"{filename}:{line_no}"
        super().__init__(f"{where}: {message}")
        self.filename = filename
        self.line_no = line_no


# --------------------------------------------------------------------------
# Column geometry
# --------------------------------------------------------------------------

#: Printed label -> canonical field name. Keys are normalised by `_norm_label`.
#: Synonyms exist so a PO that spells a column differently still lands correctly
COLUMN_ALIASES: dict[str, str] = {
    "DPT": "dept", "DEPT": "dept", "DEPARTMENT": "dept", "DIV": "dept",
    "SKU": "sku", "SKU/UPC": "sku", "ITEM": "sku", "ITEM#": "sku",
    "ITEMNO": "sku", "STYLE": "sku",
    # For a PO printing the UPC in its own column, not on a continuation line.
    "UPC": "upc", "UPC#": "upc", "UPCCODE": "upc", "EAN": "upc", "GTIN": "upc",
    "VENDORPART#": "vendor_part", "VENDORPART": "vendor_part",
    "VENDPART#": "vendor_part", "VENDORSTYLE": "vendor_part",
    "DESCRIPTION": "description", "DESC": "description",
    "ITEMDESCRIPTION": "description",
    "RETAIL": "retail", "RETAILPRICE": "retail", "SRP": "retail",
    "COST": "cost", "UNITCOST": "cost",
    "EXTCOST": "ext_cost", "EXTENDEDCOST": "ext_cost", "EXTCOST$": "ext_cost",
    "CTNS": "ctns", "CTN": "ctns", "CARTONS": "ctns", "CASES": "ctns",
    "CSPK": "cspk", "CASEPACK": "cspk", "PACK": "cspk", "PK": "cspk",
    "EXTQTY": "ext_qty", "EXTENDEDQTY": "ext_qty", "TOTALQTY": "ext_qty",
    "QTY": "ext_qty",
    "CUBE": "cube", "CUBEFT": "cube", "CBM": "cube",
    "KILOGRAMS": "kilograms", "KILOS": "kilograms", "KGS": "kilograms",
    "KG": "kilograms", "WEIGHT": "kilograms",
}

#: Canonical fields parsed as Decimal.
NUMERIC_COLUMNS = frozenset(
    {"retail", "cost", "ext_cost", "ctns", "cspk", "ext_qty", "cube", "kilograms"}
)

#: The fields a printed label can resolve to.
CANONICAL_COLUMNS = frozenset(COLUMN_ALIASES.values())

#: How many recognised columns a dashed line must underline before we accept it
#: as the item-table ruler rather than an address-block underline.
_MIN_ITEM_COLUMNS = 5


def _norm_label(text: str) -> str:
    """Uppercase, drop all whitespace. Makes label matching case- and
    spacing-insensitive. 
    Example: 'Vendor Part #' == 'VENDOR PART#'"""
    return re.sub(r"\s+", "", text).upper()


@dataclass(frozen=True)
class Column:
    name: str          # canonical name, or the normalised label not found
    label: str         # exactly as printed
    start: int
    end: int | None    # None = to end of line

    def slice(self, line: str) -> str:
        return line[self.start : self.end].strip() if self.start < len(line) else ""


@dataclass(frozen=True)
class Layout:
    """Derived column geometry for one fixed-width table."""

    columns: tuple[Column, ...]

    def by_name(self, name: str) -> Column | None:
        return next((c for c in self.columns if c.name == name), None)

    def get(self, line: str, name: str) -> str:
        col = self.by_name(name)
        return col.slice(line) if col else ""

    @property
    def names(self) -> frozenset[str]:
        return frozenset(c.name for c in self.columns)


def _runs(ruler: str) -> list[tuple[int, int]]:
    """(start, end) of each dash run. `end` is exclusive."""
    return [(m.start(), m.end()) for m in re.finditer(r"-+", ruler)]


def _is_blank_at(rows: list[str], col: int) -> bool:
    return all(len(r) <= col or r[col] == " " for r in rows)


def _cut_point(rows: list[str], gap_start: int, gap_end: int) -> int | None:
    """Pick the boundary between two adjacent columns.

    `gap_start`/`gap_end` bracket the whitespace between two dash runs
    (inclusive on both ends).

    Why not something simpler:

      * "cut at the next run's start" breaks on right-aligned numerics.
      * "cut at the gap midpoint" breaks on left-aligned text.

    Taking the *last* blank run rather than the first also protects text
    columns whose rows happen to share an interior space at the same offset.
    """
    blanks = [c for c in range(gap_start, gap_end + 1) if _is_blank_at(rows, c)]
    if not blanks:
        return None
    cut = blanks[-1]
    while cut - 1 in blanks:
        cut -= 1
    return cut


def derive_layout(header_line: str, ruler_line: str, rows: list[str]) -> Layout:
    """Build a Layout from a ruler line, its label line, and the block's rows.

    `rows` must include every physical row of the block.
    """
    runs = _runs(ruler_line)
    if not runs:
        raise ValueError("ruler line contains no dash runs")

    bounds = [runs[0][0]]
    for (_, prev_end), (next_start, _) in zip(runs, runs[1:]):
        cut = _cut_point(rows, prev_end, next_start)
        # No blank column anywhere in the gap means two columns are printed
        # flush against each other. Fall back to the ruler and let validation
        # surface whatever damage that does.
        bounds.append(cut if cut is not None else next_start)

    columns = []
    for start, end in zip(bounds, bounds[1:] + [None]):
        label = header_line[start:end].strip() if start < len(header_line) else ""
        key = _norm_label(label)
        columns.append(Column(COLUMN_ALIASES.get(key, key), label, start, end))

    # Two headings for one field leave no way to choose but a silent guess.
    seen: dict[str, str] = {}
    for col in columns:
        if col.name not in CANONICAL_COLUMNS:
            continue
        if col.name in seen:
            raise ValueError(
                f"columns {seen[col.name]!r} and {col.label!r} both mean "
                f"{col.name!r}; cannot tell which one the output should carry"
            )
        seen[col.name] = col.label
    return Layout(tuple(columns))


# --------------------------------------------------------------------------
# Scalar parsing
# --------------------------------------------------------------------------

_TRAILING_MINUS = re.compile(r"^(.*?)-$")


def parse_decimal(raw: str) -> Decimal | None:
    """Decimal, never float. Blank -> None.

    Handles thousands separators and both negative conventions
    (`(1.00)` and `1.00-`) that ERP reports use.
    """
    text = raw.strip().replace(",", "")
    if not text:
        return None
    negative = False
    if text.startswith("(") and text.endswith(")"):
        text, negative = text[1:-1].strip(), True
    elif (m := _TRAILING_MINUS.match(text)) and m.group(1):
        text, negative = m.group(1), True
    try:
        value = Decimal(text)
    except InvalidOperation:
        raise ValueError(f"not a number: {raw.strip()!r}") from None
    if not value.is_finite():
        # 'nan' and 'inf' are valid Decimal literals but never valid quantities.
        raise ValueError(f"not a finite number: {raw.strip()!r}")
    return -value if negative else value


#: (pattern, field order, whether a month above 12 may be read as day-first).
#: ISO never reorders: `2026-14-03` is malformed, not day-first.
_DATE_PATTERNS = (
    (re.compile(r"^(\d{4})-(\d{1,2})-(\d{1,2})$"), ("y", "m", "d"), False),
    (re.compile(r"^(\d{1,2})[/.-](\d{1,2})[/.-](\d{2}|\d{4})$"), ("m", "d", "y"), True),
)


def parse_date(raw: str) -> str | None:
    """Normalise to ISO `YYYY-MM-DD`. Blank -> None.

    Slash dates are read US-style (month first). A first component above 12 is
    treated as day-first instead.
    """
    text = raw.strip()
    if not text:
        return None
    for pattern, order, may_swap in _DATE_PATTERNS:
        m = pattern.match(text)
        if not m:
            continue
        parts = dict(zip(order, (int(g) for g in m.groups())))
        year, month, day = parts["y"], parts["m"], parts["d"]
        if may_swap and month > 12 and day <= 12:
            month, day = day, month
        if len(m.group(order.index("y") + 1)) == 2:
            year += 2000
        try:
            # A real calendar date, so 02/31 and 2026-14-03 are rejected.
            return date(year, month, day).isoformat()
        except ValueError:
            raise ValueError(f"not a valid date: {text!r}") from None
    raise ValueError(f"unrecognised date format: {text!r}")


# --------------------------------------------------------------------------
# Records
# --------------------------------------------------------------------------


@dataclass
class LineItem:
    line_no: int                       # 1-based line of the item's first row
    dept: str = ""
    sku: str = ""
    upc: str = ""
    vendor_part: str = ""
    description: str = ""
    retail: Decimal | None = None
    cost: Decimal | None = None
    ext_cost: Decimal | None = None
    ctns: Decimal | None = None
    cspk: Decimal | None = None
    ext_qty: Decimal | None = None
    cube: Decimal | None = None
    kilograms: Decimal | None = None


@dataclass
class Header:
    buyer: str = ""
    ship_terms: str = ""
    po_number: str = ""
    ref_master_po: str = ""
    ship_date: str = ""          # ISO, or "" when absent
    vendor: str = ""
    ship_to: str = ""
    bill_to: str = ""
    total_invoice_value: Decimal | None = None
    #: Every `LABEL:` pair found outside the address blocks, normalised key ->
    #: raw value. Kept so unexpected fields are inspectable rather than lost.
    labels: dict[str, str] = field(default_factory=dict)


@dataclass
class PurchaseOrder:
    filename: str
    header: Header
    items: list[LineItem]
    layout: Layout
    #: Numbers printed on the TOTALS line, left to right, with their line number.
    totals: list[Decimal] = field(default_factory=list)
    totals_line_no: int | None = None
    #: Item tables read, cross-checked against the count the document claims.
    pages_seen: int = 1
    pages_declared: int | None = None


# --------------------------------------------------------------------------
# Document structure
# --------------------------------------------------------------------------

#: Separator used to flatten a multi-line address block into one cell.
#: Chosen over "," because addresses already contain commas, and over "\n"
#: because a single-line cell survives CSV/clipboard round trips.
BLOCK_JOINER = " | "

_TOTALS_RE = re.compile(r"^\s*(?:GRAND\s+)?TOTALS?\s*[:.]?\s", re.IGNORECASE)
_GRAND_TOTALS_RE = re.compile(r"^\s*GRAND\s+TOTALS?\s*[:.]?\s", re.IGNORECASE)
_RULER_RE = re.compile(r"^[ \t]*-[- \t]*$")
_PAGE_RE = re.compile(r"^\s*(?:PAGE|PG)\s*[:.]?\s*\d", re.IGNORECASE)
#: The total page count out of a `PAGE : 1 of 3` marker, which prints at a line
#: start or after a gutter, never mid-sentence inside a description.
_PAGE_OF_RE = re.compile(
    r"(?:^|\s{2,})(?:PAGE|PG)\s*[:.]?\s*\d+\s*(?:OF|/)\s*(\d+)", re.IGNORECASE
)
_LABEL_RE = re.compile(
    r"(?:^|\s{2,})"                                  # line start or a gutter
    r"([A-Za-z][A-Za-z0-9#/&.'\-]*(?:[ ][A-Za-z0-9#/&.'\-]+)*)"  # label words
    r"\s*:"                                          # its colon
)
_VALUE_END_RE = re.compile(r"\s{2,}")
_NUMBER_RE = re.compile(r"\(?[\d,]*\.?\d+\)?-?")


def _is_ruler(line: str) -> bool:
    return bool(line.strip()) and bool(_RULER_RE.match(line))


def _labels_in(line: str) -> list[tuple[str, str, int]]:
    """Every `(label, value, column)` triple on one physical line."""
    found = []
    matches = list(_LABEL_RE.finditer(line))
    for i, m in enumerate(matches):
        stop = matches[i + 1].start() if i + 1 < len(matches) else len(line)
        rest = line[m.end() : stop].lstrip()
        found.append((
            _norm_label(m.group(1)),
            _VALUE_END_RE.split(rest, 1)[0].strip(),
            m.start(1),
        ))
    return found


@dataclass
class _Block:
    """A titled, dash-underlined block: VENDOR, or SHIP TO / BILL TO."""

    start: int          # index of the title line
    end: int            # exclusive; first line past the block
    cells: dict[str, str]


def _read_block(lines: list[str], ruler_idx: int) -> _Block | None:
    """Read a title + underline + indented body block ending at a blank line.
    """
    if ruler_idx == 0:
        return None
    title_line = lines[ruler_idx - 1]
    runs = _runs(lines[ruler_idx])
    starts = [s for s, _ in runs]
    titles = [
        title_line[s:e].strip()
        for s, e in zip(starts, starts[1:] + [None])
    ]
    if not any(titles) or any(not t for t in titles):
        return None

    body_end = ruler_idx + 1
    while body_end < len(lines) and lines[body_end].strip():
        body_end += 1
    body = lines[ruler_idx + 1 : body_end]
    if not body:
        return None

    cells = {}
    for title, (s, e) in zip(titles, zip(starts, starts[1:] + [None])):
        parts = [seg.strip() for line in body if (seg := line[s:e].strip())]
        cells[_norm_label(title)] = BLOCK_JOINER.join(parts)
    return _Block(ruler_idx - 1, body_end, cells)


def _find_item_rulers(lines: list[str]) -> list[int]:
    """Indices of every ruler that underlines the item table (one per page)."""
    hits = []
    for i, line in enumerate(lines):
        if i == 0 or not _is_ruler(line):
            continue
        runs = _runs(line)
        if len(runs) < _MIN_ITEM_COLUMNS:
            continue
        header = lines[i - 1]
        labels = {
            COLUMN_ALIASES.get(_norm_label(header[s:e]), "")
            for s, e in zip([r[0] for r in runs], [r[0] for r in runs][1:] + [None])
        }
        if len(labels - {""}) >= _MIN_ITEM_COLUMNS:
            hits.append(i)
    return hits


@dataclass(frozen=True)
class _PageFurniture:
    """Text this document's own header printed, and the column it printed at.

    A page break reprints its header verbatim at the same offsets, so the
    column tells a reprint apart from an item that merely contains a colon.
    """

    #: Whole normalised lines: block titles and block bodies.
    lines: frozenset[str]
    #: `(label, column)` for every `LABEL:` the header printed.
    labels: frozenset[tuple[str, int]]


def _is_furniture(line: str, header_line: str, furniture: _PageFurniture) -> bool:
    """Page headers, footers, rulers and subtotals printed inside the table."""
    if not line.strip() or "\f" in line:
        return True
    if _PAGE_RE.match(line) or _is_ruler(line):
        return True
    # A TOTALS line inside the table is a per-page subtotal.
    if _TOTALS_RE.match(line):
        return True
    if _norm_label(line) == _norm_label(header_line):
        return True
    if _norm_label(line) in furniture.lines:
        return True
    return any((label, col) in furniture.labels for label, _, col in _labels_in(line))


# --------------------------------------------------------------------------
# Entry point
# --------------------------------------------------------------------------


def parse_po(text: str, filename: str) -> PurchaseOrder:
    """Parse one purchase-order document into a header and its line items."""
    lines = text.replace("\r\n", "\n").replace("\r", "\n").split("\n")

    rulers = _find_item_rulers(lines)
    if not rulers:
        raise ParseError(
            "no item table found (expected a dashed ruler under recognisable "
            "column headings such as DESCRIPTION / QTY / COST)",
            filename=filename,
        )

    _reject_layout_drift(lines, rulers, filename)

    # Items stop at the first TOTALS after the last table, but the numbers to
    # reconcile against come from GRAND TOTALS when the document prints one.
    candidates = [
        i for i in range(rulers[-1], len(lines)) if _TOTALS_RE.match(lines[i])
    ]
    end_idx = candidates[0] if candidates else len(lines)
    totals_idx = next(
        (i for i in candidates if _GRAND_TOTALS_RE.match(lines[i])),
        candidates[0] if candidates else None,
    )

    # The header is read first so that page breaks reprinting it mid-table can
    # be recognised and skipped rather than parsed as items.
    header, furniture = _read_header(lines, rulers[0], end_idx, filename)

    header_line = lines[rulers[0] - 1]
    rows = _collect_rows(lines, rulers, end_idx, header_line, furniture)
    _reject_tabs(lines, rulers, rows, filename)
    layout = _derive_item_layout(lines, rulers[0], rows, filename)
    items = _build_items(rows, layout, filename)

    totals: list[Decimal] = []
    if totals_idx is not None:
        totals = [
            value
            for m in _NUMBER_RE.finditer(lines[totals_idx])
            if (value := parse_decimal(m.group())) is not None
        ]

    return PurchaseOrder(
        filename=filename,
        header=header,
        items=items,
        layout=layout,
        totals=totals,
        totals_line_no=None if totals_idx is None else totals_idx + 1,
        pages_seen=len(rulers),
        pages_declared=_declared_pages(lines),
    )


def _reject_layout_drift(lines: list[str], rulers: list[int], filename: str) -> None:
    """Refuse a document whose pages do not share one column geometry.

    Every page is sliced at offsets derived from page one's ruler. A page that
    moved or resized a column would still yield rows.
    """
    first = _runs(lines[rulers[0]])
    for ruler in rulers[1:]:
        if _runs(lines[ruler]) != first:
            raise ParseError(
                "this page's column ruler differs from page one's, so its rows "
                "cannot be sliced at the same offsets",
                filename=filename,
                line_no=ruler + 1,
            )


def _declared_pages(lines: list[str]) -> int | None:
    """The largest page count the document claims, or None if it never says."""
    declared = [
        int(m.group(1)) for line in lines if (m := _PAGE_OF_RE.search(line))
    ]
    return max(declared) if declared else None


def _reject_tabs(
    lines: list[str], rulers: list[int], rows: list[tuple[int, str]], filename: str
) -> None:
    """Refuse a table containing tabs rather than guessing the tab stop.
    """
    frame = [(i, lines[i - 1]) for i in rulers] + [(i + 1, lines[i]) for i in rulers]
    for line_no, text in frame + rows:
        if "\t" in text:
            raise ParseError(
                "tab character in the item table; columns are sliced by "
                "character offset and the tab stop is not recoverable",
                filename=filename,
                line_no=line_no,
            )


def _derive_item_layout(
    lines: list[str], ruler_idx: int, rows: list[tuple[int, str]], filename: str
) -> Layout:
    """`derive_layout` against the document, with the ruler line named on error."""
    try:
        return derive_layout(
            lines[ruler_idx - 1], lines[ruler_idx], [text for _, text in rows]
        )
    except ValueError as exc:
        raise ParseError(str(exc), filename=filename, line_no=ruler_idx + 1) from None


def _collect_rows(
    lines: list[str],
    rulers: list[int],
    end_idx: int,
    header_line: str,
    furniture: _PageFurniture,
) -> list[tuple[int, str]]:
    """(line_no, text) for every real table row, across every page."""
    rows = []
    for page, ruler in enumerate(rulers):
        # A page's table runs to the next page's column headings, or to TOTALS.
        stop = rulers[page + 1] - 1 if page + 1 < len(rulers) else end_idx
        for i in range(ruler + 1, min(stop, end_idx)):
            if not _is_furniture(lines[i], header_line, furniture):
                rows.append((i + 1, lines[i]))
    return rows


def _build_items(
    rows: list[tuple[int, str]], layout: Layout, filename: str
) -> list[LineItem]:
    """Fold physical rows into records, attaching each continuation upward."""
    # The continuation line prints under whichever column carries the UPC.
    id_cols = [c for c in (layout.by_name("upc"), layout.by_name("sku")) if c]
    # What a row has to fill to be a line item at all.
    substantive = id_cols + [c for c in layout.columns if c.name in NUMERIC_COLUMNS]
    items: list[LineItem] = []

    for line_no, row in rows:
        filled = [c for c in layout.columns if c.slice(row)]
        if not filled:
            raise ParseError(
                "row sits inside the item table but every column is empty",
                filename=filename,
                line_no=line_no,
            )
        # A continuation carries a value in the SKU/UPC column and nothing else.
        if len(filled) == 1 and filled[0] in id_cols:
            value = filled[0].slice(row)
            if not items:
                raise ParseError(
                    f"continuation row {value!r} has no preceding "
                    "item line to attach to",
                    filename=filename,
                    line_no=line_no,
                )
            if items[-1].upc:
                raise ParseError(
                    f"item {items[-1].sku!r} already carries UPC "
                    f"{items[-1].upc!r}; a second continuation "
                    f"{value!r} is unexpected",
                    filename=filename,
                    line_no=line_no,
                )
            items[-1].upc = value
            continue

        # A text-only row is a wrapped description
        if not any(c in filled for c in substantive):
            printed = "; ".join(f"{c.label}={c.slice(row)!r}" for c in filled)
            raise ParseError(
                "row fills only free-text columns -- no SKU, quantity or "
                f"money value, so it cannot be a line item ({printed})",
                filename=filename,
                line_no=line_no,
            )

        item = LineItem(line_no=line_no)
        for col in layout.columns:
            if not hasattr(item, col.name):
                continue                  # a column we do not model: nothing to fill
            raw = col.slice(row)
            if col.name not in NUMERIC_COLUMNS:
                setattr(item, col.name, raw)
                continue
            try:
                setattr(item, col.name, parse_decimal(raw))
            except ValueError as exc:
                raise ParseError(
                    f"column {col.label!r}: {exc}", filename=filename, line_no=line_no
                ) from None
        items.append(item)
    return items


def _read_header(
    lines: list[str], first_ruler: int, end_idx: int, filename: str
) -> tuple[Header, _PageFurniture]:
    """Read the block above the table plus the footer below TOTALS.

    Returns the header and the furniture a page break may reprint mid-table.
    """
    header = Header()
    reprinted_lines: set[str] = set()
    reprinted_labels: set[tuple[str, int]] = set()

    consumed: set[int] = set()
    for i in range(first_ruler - 1):
        if not _is_ruler(lines[i]):
            continue
        block = _read_block(lines, i)
        if block is None:
            continue
        consumed.update(range(block.start, block.end))
        reprinted_lines.update(block.cells)
        for title, value in block.cells.items():
            if title == "VENDOR":
                header.vendor = value
            elif title == "SHIPTO":
                header.ship_to = value
            elif title == "BILLTO":
                header.bill_to = value

    # A page break can reprint an entire VENDOR or SHIP TO block.
    reprinted_lines.update(
        norm for i in consumed if (norm := _norm_label(lines[i]))
    )

    above = [i for i in range(first_ruler - 1) if i not in consumed]
    for i in above:
        for label, value, col in _labels_in(lines[i]):
            header.labels.setdefault(label, value)
            reprinted_labels.add((label, col))   # a page break reprints this
    for i in range(end_idx, len(lines)):
        for label, value, _ in _labels_in(lines[i]):
            header.labels.setdefault(label, value)

    header.buyer = header.labels.get("BUYER", "")
    header.ship_terms = header.labels.get("SHIPTERMS", "")
    header.po_number = _first(header.labels, "PO#", "PONUMBER", "PO")
    header.ref_master_po = _first(header.labels, "REFMASTERPO#", "MASTERPO#")

    header.ship_date = _scalar(
        parse_date, header.labels.get("SHIPDATE", ""), "SHIP DATE", filename
    ) or ""
    header.total_invoice_value = _scalar(
        parse_decimal,
        header.labels.get("TOTALINVOICEVALUE", ""),
        "Total Invoice Value",
        filename,
    )
    return header, _PageFurniture(
        frozenset(reprinted_lines), frozenset(reprinted_labels)
    )


def _first(labels: dict[str, str], *keys: str) -> str:
    """First non-empty value among `keys`. Absent optional fields stay blank."""
    return next((labels[k] for k in keys if labels.get(k)), "")


def _scalar(convert, raw: str, field_name: str, filename: str):
    try:
        return convert(raw)
    except ValueError as exc:
        raise ParseError(f"{field_name}: {exc}", filename=filename) from None
