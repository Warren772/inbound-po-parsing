# CLAUDE.md

## What this is

Input is a plain-text PO of the kind an ERP prints or a vendor emails.
fixed-width columns, a header block, and a table of line items. Output is
`.xlsx` with one row per line item, using the exact column order in
`templates/output_template_blank.xlsx`.

Two constraints drive every decision below:

1. **Clean code pattern**
2. **The sample is one example, not the spec.** The graders may run this
   against POs I have not seen. Do not overfit to the sample file, but do not make TOO generic. 

## Stack

- Python 3.11+
- `openpyxl` for writing the workbook
- `decimal.Decimal` for all money and quantity fields
- Standard library for everything else

No pandas. Pandas will silently coerce the exact types this exercise is about preserving. No parsing
frameworks. No dependencies beyond openpyxl unless we discuss it first.

## Output contract

The template's header row is authoritative. Column order, spelling, and casing
come from the template file. 
Columns are:

```
FILENAME, BUYER, SHIP TERMS, PO#, REF MASTER PO#, SHIP DATE, VENDOR,
SHIP TO, BILL TO, Dept, SKU, UPC, Vendor Part #, Description, Retail,
Cost, EXT Cost, CTNS, CSPK, EXT QTY, Cube, Kilograms, Total Cost
```

The first nine are header-level fields repeated on every line-item row. The
rest are per-line-item, except `Total Cost`.

`FILENAME` is the source file's name.

## Parsing rules

**Derive column boundaries from the ruler line.** The item table is preceded by
a row of dashes whose runs mark each field's start and width. Parse that to get
slice offsets rather than hardcoding magic numbers. This is the single most
important decision in the repo — it's what makes the parser survive a PO with
different column widths.

**Slice by column, never split on whitespace.** Descriptions contain spaces
(`NORDPEAK ASHER STRIPE F/Q DVT GRAY`), and several fields can be blank. A
`.split()` approach appears to work on the sample and breaks on the second file.

**A line item spans two physical lines.** The first carries Dept through
Kilograms; the second is an indented UPC belonging to the line above it. Attach
it to the preceding record. A UPC with no preceding item line is a parse error,
not something to skip silently.

**SHIP TO and BILL TO are side-by-side blocks on the same physical lines.**
They must be separated by column offset, then each collapsed into a single
cell. Preserve internal order; join with a consistent separator and document
which one in the README.

**The item block ends at the TOTALS line.** Don't scan to EOF.

**Multi-page files exist.** The sample says `PAGE : 1 of 1`. Assume headers,
column rulers, and page furniture can repeat mid-file, and skip repeats rather
than emitting them as items.

## Type handling

- **UPC, SKU, Dept, and Vendor Part # are text, not numbers.** They have
  leading zeros (`0091120032129`, `04212`) and Excel will destroy them on
  numeric write. Set the cell's number format to text explicitly — don't rely
  on the value being a `str`.
- **Money and measurements use `Decimal`**, parsed from the string. Never
  `float`. `EXT Cost` values like `82913.600` carry trailing-zero precision
  that matters for tie breakers.
- **Strip thousands separators before parsing** (`447,455.12`), but only in
  numeric fields.
- **Dates normalize to ISO `YYYY-MM-DD`.** The sample mixes `3/14/2026` and
  `5/02/26` two-digit years resolve to 20xx. State this in the README.
- **Blank means blank.** Empty cells stay empty, no placeholder value. 

## Validation

The PO carries its own checks to cross-reference. Validate these and probe for more.

- `EXT QTY` should equal `CTNS × CSPK`
- `EXT Cost` should equal `Cost × EXT QTY`
- Summed `EXT Cost`, `CTNS`, `EXT QTY`, and `Cube` should match the TOTALS row
- The TOTALS `EXT Cost` should match `Total Invoice Value`

Report mismatches to stderr with the file and line or in client UI. 

## Generalization

Do not tune to the sample. Specifically:

- No hardcoded slice indices.
- No matching on the literal vendor, buyer, or brand names in the sample.
- Label matching should be case-insensitive aware of extra whitespace.
- A missing optional header field yields a blank cell, not a crash. Also, extra logging or UI signal (if we extend in the future) to surface this to the user. 
- An unparseable line raises with the filename and line number in the message.

## Code organization

- `parse.py` — text -> list of records
- `write.py` — records -> workbook based on template
- `validate.py` — the math checks
- `cli.py` — argument handling, batch loop, error reporting
- `templates/` — the provided blank template
- `samples/` — the provided sample PO (commit)

Rules:

- Parsing is pure: text in, dataclasses out. No file I/O, no openpyxl imports.
- One dataclass for the header block, one for a line item. Typed fields.
- No abstraction until the second use. No plugin architecture, no format
  registry, no config files.
- Regex only where a regex is clearer than slicing.

## Testing

`pytest`, small and targeted. Example tests:

1. Header field extraction from the sample
2. The two-line item + UPC pairing
3. SHIP TO / BILL TO column split
4. Leading zeros survive a write/read round trip
5. The validation arithmetic

Hand-write one trimmed fixture rather than depending on the full sample for
every test. Do not reward hack for coverage wins. Apply to edge-cases, not obvious
malformed errors. 

## README requirements

The brief rewards stated assumptions over silent guesses. The README must
cover, at minimum:

- How to run it, including the batch case
- `Total Cost` — which value it holds and why (see below)
- The SHIP TO / BILL TO join separator
- Two-digit year resolution
- What happens on a validation mismatch
- Approximate time spent

## Assumptions to flag, not resolve silently

- **`Total Cost`** is ambiguous: it may be the document's Total Invoice Value
  repeated on each row, or a per-row figure duplicating `EXT Cost`. 
- **`VENDOR`** in the template is one cell but the source is a multi-line block
  including a vendor number, `C/O` line, address, phone, and fax. Document
  whether it holds the full block or just the name and number.
- **`BUYER`** in the sample is `610 JORDAN MILLER` — a code plus a name.
  Keep it verbatim rather than splitting; note the choice.

## Working style

- Small, reviewable commits. Feature-sized PRs.
- Explain trade-offs before implementing anything structural. I want the
  decision, not just the diff.
- If the brief is ambiguous, state the assumption in the README rather than
  guessing silently.
- Push back if I ask for something that contradicts the rules above.
