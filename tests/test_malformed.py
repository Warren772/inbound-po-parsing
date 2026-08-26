"""Every test here is driven by a file in `fixtures/malformed/`.

Each fixture is a whole purchase order that is wrong in exactly one way, so a
failure names the shape of document that broke rather than a line of code.
"""

from decimal import Decimal

import openpyxl
import pytest

from po2xlsx.parse import ParseError, parse_po
from po2xlsx.validate import ERROR, WARNING, validate
from po2xlsx.write import write_workbook

from conftest import MALFORMED

#: Fixtures that must raise, mapped to a phrase the message has to contain.
REJECTED = {
    "duplicate_qty_columns.txt": "both mean 'ext_qty'",
    "impossible_ship_date.txt": "not a valid date",
    "no_item_table.txt": "no item table found",
    "orphan_upc.txt": "no preceding item line",
    "tabbed_columns.txt": "tab character",
    "text_in_numeric_column.txt": "not a number",
    "wrapped_description.txt": "only free-text columns",
}
ACCEPTED = {
    "credit_lines.txt",
    "formula_injection.txt",
    "label_in_description.txt",
    "missing_page.txt",
    "page_subtotals.txt",
    "upc_in_own_column.txt",
}


def load(name):
    path = MALFORMED / name
    return parse_po(path.read_text(), path.name)


def rejects(name):
    with pytest.raises(ParseError) as excinfo:
        load(name)
    return str(excinfo.value)


def test_the_corpus_is_completely_classified():
    """A new fixture must be declared accepted or rejected, not left to drift
    into whichever bucket the code happens to put it in."""
    on_disk = {p.name for p in MALFORMED.glob("*.txt")}
    assert on_disk == set(REJECTED) | ACCEPTED


# --- what must be refused -------------------------------------------------


@pytest.mark.parametrize("name, phrase", sorted(REJECTED.items()))
def test_rejected_file_names_itself_and_says_why(name, phrase):
    message = rejects(name)
    assert name in message
    assert phrase in message


@pytest.mark.parametrize(
    "name, line_no",
    [
        ("duplicate_qty_columns.txt", 5),   # the ruler line
        ("orphan_upc.txt", 6),
        ("tabbed_columns.txt", 6),
        ("text_in_numeric_column.txt", 6),
        ("wrapped_description.txt", 7),
    ],
)
def test_rejected_file_names_the_offending_line(name, line_no):
    """'An unparseable line raises with the filename and line number.'"""
    assert f":{line_no}:" in rejects(name)


def test_tabs_are_refused_rather_than_guessed():
    """A tab is one character and prints as up to eight. Slicing it would
    misfile every field to its right as plausible-looking data."""
    assert "tab stop is not recoverable" in rejects("tabbed_columns.txt")


def test_two_columns_meaning_one_field_name_both_labels():
    """QTY beside EXT QTY: whichever this parser picked would be a silent guess
    at which number reaches the output."""
    message = rejects("duplicate_qty_columns.txt")
    assert "'QTY'" in message and "'EXT QTY'" in message


def test_a_wrapped_description_is_not_a_phantom_line_item():
    """The wrap row fills DESCRIPTION alone. Emitting it would put a
    blank-money row in the workbook and hand it the next item's UPC."""
    message = rejects("wrapped_description.txt")
    assert "WITH A CONTINUED NAME" in message


def test_bad_column_value_names_the_printed_column_label():
    assert "'COST'" in rejects("text_in_numeric_column.txt")


# --- what must be read correctly ------------------------------------------


def test_page_subtotals_do_not_truncate_the_document():
    """The regression this corpus exists for. Stopping at the first TOTALS
    ends the block at page one's subtotal, and the rows that survive still tie
    out against it -- so no arithmetic check can see the missing page."""
    po = load("page_subtotals.txt")
    assert [item.sku for item in po.items] == ["5551", "5552", "5553"]
    assert po.pages_seen == 2


def test_the_document_totals_are_the_grand_totals_not_a_subtotal():
    po = load("page_subtotals.txt")
    assert po.totals == [
        Decimal("142.000"), Decimal("10"), Decimal("28"), Decimal("8.15")
    ]
    assert [i for i in validate(po) if i.severity == ERROR] == []


def test_a_page_subtotal_never_becomes_a_line_item():
    po = load("page_subtotals.txt")
    assert all(item.description for item in po.items)
    assert Decimal("82.000") not in [item.ext_cost for item in po.items]


def test_upc_printed_in_its_own_column_is_kept():
    """The template has a UPC column; a PO that fills one inline rather than on
    a continuation line must not lose it."""
    po = load("upc_in_own_column.txt")
    assert [item.upc for item in po.items] == ["0001234567890", "0009876543210"]
    assert [item.sku for item in po.items] == ["5551", "5552"]
    assert not [i for i in validate(po) if "no UPC" in i.message]


def test_a_description_containing_a_header_label_is_still_an_item():
    """Page furniture is matched with its column offset. Without that anchor an
    item reading 'PO#: SEE MASTER' is mistaken for a reprinted header line and
    silently dropped."""
    po = load("label_in_description.txt")
    assert [item.description for item in po.items] == [
        "PO#: SEE MASTER", "SHIP DATE: ON FILE"
    ]
    assert po.header.po_number == "AB1"          # the real label still won


def test_a_page_the_parser_never_saw_is_reported():
    po = load("missing_page.txt")
    issues = [i for i in validate(po) if "declares 3 pages" in i.message]
    assert [i.severity for i in issues] == [WARNING]


def test_credits_keep_their_sign_and_still_reconcile():
    """`(55.000)` and `5-` are both negatives in ERP print."""
    po = load("credit_lines.txt")
    item = po.items[0]
    assert (item.ctns, item.ext_qty, item.ext_cost) == (
        Decimal("-5"), Decimal("-10"), Decimal("-55.000")
    )
    assert item.cost * item.ext_qty == item.ext_cost
    assert [i for i in validate(po) if i.severity == ERROR] == []


# --- the workbook side ----------------------------------------------------


@pytest.mark.parametrize("name", sorted(ACCEPTED))
def test_every_accepted_fixture_writes_a_workbook(name, template, tmp_path):
    """Whatever a readable-but-odd PO does to the records, it must not take the
    writer down."""
    out = tmp_path / f"{name}.xlsx"
    po = load(name)
    write_workbook([po], template, out)
    assert openpyxl.load_workbook(out).active.max_row == 1 + len(po.items)
