# po2xlsx

Converts fixed-width purchase-order text into `.xlsx`, one row per line item, using the column order defined by `templates/output_template_blank.xlsx`.

## Running it

```bash
python -m venv .venv && .venv/bin/pip install -e '.[dev]'

# one file -> samples/purchase_order_sample.xlsx, beside the input
.venv/bin/python -m po2xlsx samples/purchase_order_sample.txt

# batch: every input in one workbook, FILENAME distinguishes the rows
.venv/bin/python -m po2xlsx samples/*.txt --out out/purchase_orders.xlsx

# options
--template PATH        blank template to take the column order from
                       (default: templates/output_template_blank.xlsx)
--total-cost {invoice,ext}   see "Total Cost" below (default: invoice)
--strict               treat validation warnings as failures
--quiet                suppress the per-file summary
```

Exit codes: `0` clean, `1` validation errors (the workbook is still written),
`2` the input could not be parsed or the template is missing.

```bash
.venv/bin/python -m pytest
```

CI runs the suite on Python 3.11–3.13, again with `lxml` installed (openpyxl
swaps its whole serialiser when it is present), and then replays it at **every
commit** a push or PR adds, so the history stays bisectable.

## How the columns are found

The item table's geometry is derived from the document.

The parser does not assume a fixed split between columns. Instead, for each gap between two dash runs, it uses the start of the last blank column run that is empty in every row of the item block, including continuation lines.

This is important  because UPC values can continue on a second physical line and still belong to the same item. If the parser only checked the main item rows, the UPC would be cut off.

Field *names* come from slicing the heading line with those same boundaries and
matching against a synonym table (`DPT`/`DEPT`/`DEPARTMENT`, `KGS`/`KILOGRAMS`,
etc.).

Address blocks (`VENDOR`, `SHIP TO` / `BILL TO`) use the simpler rule of
splitting at each underline run's start.

## Assumptions and choices

### `Total Cost`

It could be the document's Total Invoice Value repeated on every row, or a per-row figure duplicating `EXT Cost`.

**Default (`--total-cost invoice`): the document's Total Invoice Value, repeated
on every row** — `447455.12` for the sample.

The reasoning is that the template
already has a per-row `EXT Cost` immediately to the left, so a second column
holding the same number on every row would carry no information.


Falls back to the TOTALS row's `EXT COST`, then to the summed `EXT Cost`,
if the document does not print a Total Invoice Value.

`--total-cost ext` produces the other reading. 

### `VENDOR`, `SHIP TO`, `BILL TO`

The source is a multi-line block the template is one cell. Each holds **the
entire block**, lines joined with `" | "`:

```
50201 HARBOR POINT TRADING LTD | C/O MERIDIAN SOURCING CO INC | 410 COMMERCE AVE | EDISON NJ 08837 | PHONE : 9085551040 | FAX : 9085551099
```

`" | "` was chosen
over a comma (addresses already contain commas) and over a newline (a
single-line cell survives CSV and clipboard output).

`SHIP TO` and `BILL TO` print side by side on the same physical lines and are
separated by column offset before joining.

### `BUYER`

Kept verbatim, `610 JORDAN MILLER`: a code plus a name. Not split, because this information may be essential in the next PO.

### Dates

Normalised to ISO `YYYY-MM-DD` and written as **text**, so Excel cannot
reinterpret them under a local date format. 

Slash dates are read US-style, month first. A first component above 12 is read
day-first instead.

### What is refused

A document this parser cannot read fails with the file and line named, rather
than producing a workbook from a misreading. `tests/fixtures/malformed/` holds
one whole PO per failure mode. Refused:

* **tabs inside the item table** — geometry here is counted in characters
* **two headings meaning one field** (`QTY` beside `EXT QTY`)
* **a row that fills only free-text columns** — no SKU, quantity or money.
* **a page whose ruler differs from page one's** — every page is sliced at page one's offsets.

### Types

A `UPC` is read from a continuation line under the SKU column, or from its own
column when the document prints one.

`UPC`, `SKU`, `Dept` and `Vendor Part #` keep their leading zeros and are
written with Excel's text format (`@`) set explicitly, not as Python
`str`.

Money and measurements are parsed as `decimal.Decimal` and never pass through a
float. 

One wrinkle worth flagging: openpyxl formats numeric cells with `"%.16g" %
value`, which coerces a `Decimal` through `float` and writes
`82913.60000000001` into the XML. The stored double is identical either way,
but `write.py` replaces this logic, so **beware if using future verisons of
openpyxl**.

**Text cells are never formulas.** openpyxl types any string beginning with `=`
as a formula.

Blank entries stay blank in output

### Multi-page files

Headings, rulers and page furniture are skipped where they repeat mid-file.

Furniture is recognized by matching against the labels and block titles this
document's *own* header used, plus `PAGE :` markers and form feeds. A label is
matched **together with the column it was printed at**.

**The item block ends at the TOTALS line after the *last* item table** A PO that prints a per-page subtotal and a final
`GRAND TOTALS` would otherwise end at page one. Where both are printed, the
grand total is the one reconciled against.

As a backstop, a `PAGE : 1 of 3` marker is cross-checked against the number of
item tables read.

## Validation

Checks run against every file and report to stderr as
`severity: file:line: message`. 

* `EXT QTY` = `CTNS` × `CSPK`, per row
* `EXT Cost` = `Cost` × `EXT QTY`, per row
* every number on the TOTALS row ties to a column sum
* TOTALS `EXT Cost` = `Total Invoice Value`
* header fields that came through blank are listed (warning)
* items with no UPC continuation, and duplicate UPCs (warning)
* `RETAIL` below `COST`, non-positive quantities (warning)
* the page count the document declares vs. the item tables read (warning)

Errors set exit code `1`; warnings do not unless `--strict` is passed.

### Two findings about the TOTALS row

**The TOTALS line is not column-aligned with the item table.** In the sample
`447455.120` sits at columns 68–78 while `EXT COST` occupies 85–96.

**The fourth total is Kilograms, not Cube.** The brief expects TOTALS to carry
`EXT Cost`, `CTNS`, `EXT QTY` and `Cube`. The arithmetic disagrees:
`sum(Kilograms) = 119.18`, which is the printed value, while
`sum(Cube) = 0.9050`. The sample's TOTALS row is `EXT COST`, `CTNS`, `EXT QTY`,
`KILOGRAMS`. 

Because matching is by value rather than by position, a PO that
really does total `Cube` is handled too, with no change.

## Layout

```
po2xlsx/parse.py      text -> dataclasses
po2xlsx/validate.py   the math checks
po2xlsx/write.py      records -> workbook
po2xlsx/cli.py        arguments, batch loop, error reporting
samples/              the provided sample PO
templates/            the provided blank template (not committed)
tests/                pytest, with hand-built trimmed test cases
tests/fixtures/malformed/   one whole PO per failure mode
.github/workflows/ci.yml    matrix tests, lxml leg, per-commit replay
```
