from pathlib import Path

import openpyxl
import pytest

from po2xlsx.parse import parse_po

ROOT = Path(__file__).resolve().parent.parent
FIXTURES = Path(__file__).parent / "fixtures"
SAMPLE = ROOT / "samples" / "purchase_order_sample.txt"
TEMPLATE = ROOT / "templates" / "output_template_blank.xlsx"

TEMPLATE_COLUMNS = [
    "FILENAME", "BUYER", "SHIP TERMS", "PO#", "REF MASTER PO#", "SHIP DATE",
    "VENDOR", "SHIP TO", "BILL TO", "Dept", "SKU", "UPC", "Vendor Part #",
    "Description", "Retail", "Cost", "EXT Cost", "CTNS", "CSPK", "EXT QTY",
    "Cube", "Kilograms", "Total Cost",
]


def load(path: Path):
    return parse_po(path.read_text(), path.name)


@pytest.fixture
def fixtures_dir():
    return FIXTURES


@pytest.fixture
def sample():
    return load(SAMPLE)


@pytest.fixture
def mini():
    """A trimmed PO with deliberately different column widths to the sample."""
    return load(FIXTURES / "mini_po.txt")


@pytest.fixture
def broken():
    """The same layout with seeded arithmetic and totals errors."""
    return load(FIXTURES / "broken_po.txt")


@pytest.fixture
def template(tmp_path):
    """A stand-in for the provided template, which is not committed."""
    workbook = openpyxl.Workbook()
    workbook.active.append(TEMPLATE_COLUMNS)
    path = tmp_path / "template.xlsx"
    workbook.save(path)
    return path
