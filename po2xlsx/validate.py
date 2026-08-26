"""Arithmetic and consistency checks.

The document carries the totals to check them against the parsed rows
and report what does not reconcile.
"""

from __future__ import annotations

from dataclasses import dataclass
from decimal import Decimal

from .parse import NUMERIC_COLUMNS, PurchaseOrder

__all__ = ["Issue", "TotalsMatch", "validate", "match_totals", "column_sums"]

ERROR = "error"
WARNING = "warning"


@dataclass(frozen=True)
class Issue:
    severity: str
    filename: str
    line_no: int | None
    message: str

    def __str__(self) -> str:
        where = self.filename if self.line_no is None else f"{self.filename}:{self.line_no}"
        return f"{self.severity}: {where}: {self.message}"


@dataclass(frozen=True)
class TotalsMatch:
    """Which printed TOTALS number belongs to which column."""

    sums: dict[str, Decimal]
    matched: dict[str, Decimal]      # column name -> printed total
    unmatched: list[Decimal]         # printed totals matching no column sum


def column_sums(po: PurchaseOrder) -> dict[str, Decimal]:
    """Sum each numeric column the document prints, in printed order."""
    names = [c.name for c in po.layout.columns if c.name in NUMERIC_COLUMNS]
    return {
        name: sum(
            (v for it in po.items if (v := getattr(it, name, None)) is not None),
            Decimal(0),
        )
        for name in names
    }


def match_totals(po: PurchaseOrder) -> TotalsMatch:
    """Assign each number on the TOTALS line to the column it totals.

    Each printed number is matched by value against the column
    sums, left to right, with the constraint that assignments stay in column
    order.
    """
    sums = column_sums(po)
    names = list(sums)
    matched: dict[str, Decimal] = {}
    unmatched: list[Decimal] = []

    cursor = 0
    for printed in po.totals:
        for index in range(cursor, len(names)):
            if sums[names[index]] == printed:
                matched[names[index]] = printed
                cursor = index + 1
                break
        else:
            unmatched.append(printed)
    return TotalsMatch(sums=sums, matched=matched, unmatched=unmatched)


def validate(po: PurchaseOrder) -> list[Issue]:
    """Every check, most specific first. Empty list means the document ties out."""
    issues: list[Issue] = []

    def add(severity: str, message: str, line_no: int | None = None) -> None:
        issues.append(Issue(severity, po.filename, line_no, message))

    if not po.items:
        add(ERROR, "no line items were found between the ruler and TOTALS")
        return issues

    _check_row_arithmetic(po, add)
    _check_totals(po, add)
    _check_completeness(po, add)
    return issues


def _check_row_arithmetic(po: PurchaseOrder, add) -> None:
    for item in po.items:
        if None not in (item.ctns, item.cspk, item.ext_qty):
            expected = item.ctns * item.cspk
            if expected != item.ext_qty:
                add(
                    ERROR,
                    f"EXT QTY {item.ext_qty} != CTNS {item.ctns} x CSPK "
                    f"{item.cspk} ({expected})",
                    item.line_no,
                )
        if None not in (item.cost, item.ext_qty, item.ext_cost):
            expected = item.cost * item.ext_qty
            if expected != item.ext_cost:
                add(
                    ERROR,
                    f"EXT COST {item.ext_cost} != COST {item.cost} x EXT QTY "
                    f"{item.ext_qty} ({expected})",
                    item.line_no,
                )
        if None not in (item.retail, item.cost) and item.retail < item.cost:
            add(WARNING, f"RETAIL {item.retail} is below COST {item.cost}", item.line_no)
        for name in ("ctns", "cspk", "ext_qty"):
            value = getattr(item, name)
            if value is not None and value <= 0:
                add(WARNING, f"{name.upper()} is {value}", item.line_no)


def _check_totals(po: PurchaseOrder, add) -> None:
    if po.totals_line_no is None:
        add(WARNING, "no TOTALS line found! Document totals were not cross-checked")
        return

    match = match_totals(po)
    for printed in match.unmatched:
        add(
            ERROR,
            f"TOTALS value {printed} does not equal any column sum "
            f"({_render(match.sums)})",
            po.totals_line_no,
        )
    if not match.matched:
        add(WARNING, "no TOTALS value could be tied to a column", po.totals_line_no)

    invoice = po.header.total_invoice_value
    if invoice is None:
        add(WARNING, "no Total Invoice Value on the document to cross-check")
        return
    reconciled = match.matched.get("ext_cost", match.sums.get("ext_cost"))
    if reconciled is not None and reconciled != invoice:
        add(
            ERROR,
            f"Total Invoice Value {invoice} != EXT COST total {reconciled}",
            po.totals_line_no,
        )


#: Header fields worth reporting when they arrive blank. They are optional --
#: a missing one is a blank cell, never a crash -- but silence would hide a
#: label this parser failed to recognise on an unfamiliar PO.
_EXPECTED_HEADER_FIELDS = (
    ("po_number", "PO#"),
    ("buyer", "BUYER"),
    ("ship_terms", "SHIP TERMS"),
    ("ref_master_po", "REF MASTER PO#"),
    ("ship_date", "SHIP DATE"),
    ("vendor", "VENDOR"),
    ("ship_to", "SHIP TO"),
    ("bill_to", "BILL TO"),
)


def _check_completeness(po: PurchaseOrder, add) -> None:
    blank = [
        label for attr, label in _EXPECTED_HEADER_FIELDS
        if not getattr(po.header, attr)
    ]
    if blank:
        add(WARNING, f"header fields left blank: {', '.join(blank)}")

    seen: dict[str, int] = {}
    for item in po.items:
        if not item.upc:
            add(WARNING, f"item {item.sku or '?'} has no UPC continuation line",
                item.line_no)
        elif (first := seen.get(item.upc)) is not None:
            add(WARNING, f"UPC {item.upc} already appeared on line {first}",
                item.line_no)
        else:
            seen[item.upc] = item.line_no


def _render(sums: dict[str, Decimal]) -> str:
    return ", ".join(f"{name}={value}" for name, value in sums.items())
