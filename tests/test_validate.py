from decimal import Decimal

from po2xlsx.validate import ERROR, WARNING, column_sums, match_totals, validate


def test_sample_reconciles_completely(sample):
    assert validate(sample) == []


def test_totals_are_matched_by_value_not_by_column_position(sample):
    """The TOTALS line is printed on its own grid: 447455.120 sits at columns
    68-78 while EXT COST occupies 85-96."""
    match = match_totals(sample)
    assert match.matched == {
        "ext_cost": Decimal("447455.120"),
        "ctns": Decimal("3475"),
        "ext_qty": Decimal("6950"),
        "kilograms": Decimal("119.18"),
    }
    assert match.unmatched == []


def test_the_fourth_total_is_kilograms_not_cube(sample):
    """Worth pinning: the totals row looks like it ends with Cube, but the
    arithmetic says Kilograms."""
    sums = column_sums(sample)
    assert sums["kilograms"] == Decimal("119.18")
    assert sums["cube"] == Decimal("0.9050")
    assert "cube" not in match_totals(sample).matched


def test_totals_ext_cost_matches_total_invoice_value(sample):
    assert column_sums(sample)["ext_cost"] == sample.header.total_invoice_value


def test_row_arithmetic(sample):
    for item in sample.items:
        assert item.ext_qty == item.ctns * item.cspk
        assert item.ext_cost == item.cost * item.ext_qty


def test_ext_qty_mismatch_is_reported_with_its_line(broken):
    issues = [i for i in validate(broken) if i.message.startswith("EXT QTY")]
    assert len(issues) == 1
    assert issues[0].severity == ERROR
    assert issues[0].line_no == broken.items[0].line_no
    assert "5 x CSPK 2" in issues[0].message


def test_ext_cost_mismatch_is_reported_per_row(broken):
    lines = {i.line_no for i in validate(broken)
             if i.message.startswith("EXT COST")}
    assert lines == {item.line_no for item in broken.items}


def test_unmatched_totals_value_is_reported(broken):
    messages = [i.message for i in validate(broken) if "TOTALS value" in i.message]
    assert any("777.000" in m for m in messages)
    assert all(i.severity == ERROR for i in validate(broken)
               if "TOTALS value" in i.message)


def test_missing_upc_is_a_warning_not_an_error(broken):
    issues = [i for i in validate(broken) if "no UPC" in i.message]
    assert [i.severity for i in issues] == [WARNING]


def test_mini_fixture_reconciles(mini):
    assert [i for i in validate(mini) if i.severity == ERROR] == []


def test_blank_header_fields_are_surfaced(mini):
    """A missing optional field is a blank cell, but the user is told."""
    messages = [i.message for i in validate(mini) if "header fields" in i.message]
    assert messages == ["header fields left blank: REF MASTER PO#"]
    assert all(i.severity == WARNING for i in validate(mini)
               if "header fields" in i.message)
