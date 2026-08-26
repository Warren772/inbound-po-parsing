from decimal import Decimal
from pathlib import Path

import pytest

from po2xlsx.parse import ParseError, derive_layout, parse_date, parse_decimal, parse_po


# --- header block ---------------------------------------------------------


def test_header_fields_from_sample(sample):
    header = sample.header
    assert header.buyer == "610 JORDAN MILLER"      # code + name, kept verbatim
    assert header.ship_terms == "FOB Ningbo,CHN"
    assert header.po_number == "TX9K2QP"
    assert header.ref_master_po == "TX9H1LM"
    assert header.ship_date == "2026-05-02"
    assert header.total_invoice_value == Decimal("447455.12")


def test_buyer_stops_at_the_gutter(sample):
    """The company name and order date share BUYER's physical line."""
    assert "Meridian" not in sample.header.buyer
    assert "3/14/2026" not in sample.header.buyer


def test_vendor_block_is_the_whole_block(sample):
    vendor = sample.header.vendor
    assert vendor.startswith("50201 HARBOR POINT TRADING LTD")
    assert "FAX : 9085551099" in vendor          # address and contact retained


def test_ship_to_and_bill_to_split_by_column(sample):
    """Side-by-side blocks must separate by offset"""
    assert sample.header.ship_to == (
        "331 DISTRIBUTION/NORTHEAST | C/O SWIFTLINE EXPRESS | "
        "1450 CROSSROADS BLVD | EDISON, NJ 08817 | PHONE : 9085551088"
    )
    assert sample.header.bill_to == (
        "331 MERIDIAN SOURCING CO. INC. | "
        "C/O NORTHSTAR RETAIL GROUP - DEPT OPS | PO BOX 4021 | EDISON, NJ 08837"
    )
    # The SHIP TO phone line has no BILL TO counterpart.
    assert "9085551088" not in sample.header.bill_to


def test_missing_optional_header_field_is_blank_not_an_error(mini):
    assert mini.header.ref_master_po == ""
    assert mini.header.po_number == "AB1"


# --- item rows ------------------------------------------------------------


def test_upc_continuation_attaches_to_the_line_above(sample):
    assert len(sample.items) == 12
    first = sample.items[0]
    assert first.sku == "71097188"
    assert first.upc == "0091120032129"
    assert sample.items[-1].upc == "0091120032426"
    assert all(item.upc for item in sample.items)


def test_upc_is_wider_than_the_sku_column_it_shares(sample):
    """The UPC overruns the SKU's dashes."""
    assert len(sample.items[0].upc) == 13
    assert len(sample.items[0].sku) == 8


def test_description_keeps_its_internal_spaces(sample):
    descriptions = [item.description for item in sample.items]
    assert "NORDPEAK ASHER STRIPE F/Q DVT GRAY" in descriptions


def test_money_keeps_source_scale(sample):
    """82913.600 must not collapse to 82913.6"""
    item = sample.items[1]
    assert item.ext_cost == Decimal("82913.600")
    assert str(item.ext_cost) == "82913.600"
    assert isinstance(item.cost, Decimal)


def test_orphan_continuation_names_file_and_line(tmp_path, fixtures_dir):
    lines = (fixtures_dir / "mini_po.txt").read_text().splitlines()
    ruler = next(i for i, line in enumerate(lines) if set(line.strip()) == {"-", " "}
                 and line.count("-") > 40)
    orphan = lines[:ruler + 1] + ["     0001234567890"] + lines[ruler + 1:]
    path = tmp_path / "orphan.txt"
    path.write_text("\n".join(orphan))

    with pytest.raises(ParseError) as excinfo:
        parse_po(path.read_text(), path.name)
    assert "orphan.txt" in str(excinfo.value)
    assert f":{ruler + 2}:" in str(excinfo.value)
    assert "no preceding item line" in str(excinfo.value)


# --- derived geometry -----------------------------------------------------


def test_layout_derived_from_a_different_ruler(mini):
    """Same code, different widths"""
    assert {c.name for c in mini.layout.columns} >= {
        "dept", "sku", "vendor_part", "description", "retail", "cost",
        "ext_cost", "ctns", "cspk", "ext_qty", "cube", "kilograms",
    }
    assert mini.items[0].description == "WIDE LOAD / BIG BOX"
    assert mini.items[0].cube == Decimal("1.2345")
    assert mini.items[0].kilograms == Decimal("3.10")


def test_right_aligned_value_wider_than_its_dashes():
    """COST prints 6 characters under a 4-dash heading, spilling left into the
    gap."""
    header = "ITEM    COST"
    ruler = "----    ----"
    rows = ["A     12.500"]
    layout = derive_layout(header, ruler, rows)
    assert layout.get(rows[0], "sku") == "A"
    assert layout.get(rows[0], "cost") == "12.500"


def test_left_aligned_value_wider_than_its_dashes():
    """DESCRIPTION prints far past its dashes"""
    header = "DESCRIPTION        QTY"
    ruler = "-----------        ---"
    rows = ["A VERY LONG NAME     7"]
    layout = derive_layout(header, ruler, rows)
    assert layout.get(rows[0], "description") == "A VERY LONG NAME"
    assert layout.get(rows[0], "ext_qty") == "7"


def test_blank_fields_stay_blank(mini):
    assert mini.items[1].dept == ""            # blank DEPT, still an item row
    assert mini.items[1].sku == "5552"
    assert mini.items[2].vendor_part == ""


def test_page_furniture_between_pages_is_skipped(mini):
    """Repeated headings, rulers and a PAGE marker must not become items."""
    assert len(mini.items) == 3
    assert [item.sku for item in mini.items] == ["5551", "5552", "5553"]


def test_item_block_ends_at_totals(mini):
    assert mini.totals == [
        Decimal("142.000"), Decimal("10"), Decimal("28"), Decimal("8.15")
    ]
    assert all(item.sku != "142.000" for item in mini.items)


# --- scalars --------------------------------------------------------------


@pytest.mark.parametrize(
    "raw, expected",
    [
        ("447,455.12", Decimal("447455.12")),   # thousands separator
        ("82913.600", Decimal("82913.600")),    # trailing zeros preserved
        ("(1.50)", Decimal("-1.50")),           # accounting negative
        ("1.50-", Decimal("-1.50")),            # trailing-minus negative
        ("   ", None),                          # blank stays blank
    ],
)
def test_parse_decimal(raw, expected):
    assert parse_decimal(raw) == expected


def test_parse_decimal_scale_survives():
    assert str(parse_decimal("82913.600")) == "82913.600"


@pytest.mark.parametrize("raw", ["abc", "1.2.3", "nan", "inf"])
def test_parse_decimal_rejects_non_quantities(raw):
    with pytest.raises(ValueError):
        parse_decimal(raw)


@pytest.mark.parametrize(
    "raw, expected",
    [
        ("3/14/2026", "2026-03-14"),   # four-digit year
        ("5/02/26", "2026-05-02"),     # two-digit year resolves to 20xx
        ("12/31/27", "2027-12-31"),
        ("2026-05-02", "2026-05-02"),  # already ISO
        ("", None),
    ],
)
def test_parse_date(raw, expected):
    assert parse_date(raw) == expected


def test_parse_date_reads_day_first_when_month_cannot_be(raw="14/3/2026"):
    assert parse_date(raw) == "2026-03-14"


def test_parse_date_rejects_nonsense():
    with pytest.raises(ValueError):
        parse_date("31/31/2026")


# --- regressions from PR#1 review -------------------------------------------


@pytest.mark.parametrize("raw", ["2026-14-03", "2026-02-30"])
def test_iso_dates_are_never_reordered(raw):
    """A malformed ISO date is malformed, not day-first: 2026-14-03 must not
    become 2026-03-14."""
    with pytest.raises(ValueError):
        parse_date(raw)


@pytest.mark.parametrize("raw", ["02/31/2026", "4/31/2026", "2/30/26"])
def test_impossible_calendar_dates_are_rejected(raw):
    """A day inside 1-31 is not enough; the date has to actually exist."""
    with pytest.raises(ValueError):
        parse_date(raw)


def test_reprinted_address_block_is_not_parsed_as_items(fixtures_dir):
    """A page break can reprint the whole VENDOR block, not just its title.
    Its body lines must be skipped rather than becoming bogus line items."""
    lines = (fixtures_dir / "mini_po.txt").read_text().split("\n")
    at = lines.index("PAGE : 1 of 2")
    reprinted = ["VENDOR", "------", " 90001 SMALL SUPPLIER GMBH", " 12 KURZ STR"]
    text = "\n".join(lines[: at + 1] + reprinted + lines[at + 1 :])

    po = parse_po(text, "reprinted.txt")
    assert [item.sku for item in po.items] == ["5551", "5552", "5553"]
