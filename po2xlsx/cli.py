"""Command line entry point: argument handling, batch loop, error reporting.

    python -m po2xlsx samples/*.txt --out out/purchase_orders.xlsx

`Total Cost` is ambiguous in the brief. The default (`--total-cost invoice`)
puts the document's Total Invoice Value on every row, on the reading that a
column named "Total" next to a per-row "EXT Cost" is the document total. Pass
`--total-cost ext` to duplicate each row's EXT Cost instead.
"""

from __future__ import annotations

import argparse
import sys
from pathlib import Path

from .parse import ParseError, parse_po
from .validate import ERROR, match_totals, validate
from .write import TOTAL_COST_MODES, write_workbook

__all__ = ["main", "build_parser"]

EXIT_OK = 0
EXIT_VALIDATION = 1
EXIT_PARSE = 2


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="po2xlsx",
        description="Convert fixed-width purchase-order text files to .xlsx.",
    )
    parser.add_argument("inputs", nargs="+", type=Path, help="purchase-order text files")
    parser.add_argument(
        "--template",
        type=Path,
        default=Path("templates/output_template_blank.xlsx"),
        help="blank template whose header row defines the output columns",
    )
    parser.add_argument(
        "--out",
        type=Path,
        help="write every input into this one workbook; "
        "omit to write <input>.xlsx beside each input",
    )
    parser.add_argument(
        "--total-cost",
        choices=TOTAL_COST_MODES,
        default="invoice",
        help="'invoice' repeats the document's Total Invoice Value on every row; ",
    )
    parser.add_argument(
        "--strict",
        action="store_true",
        help="treat validation warnings as failures",
    )
    parser.add_argument(
        "--quiet", action="store_true", help="suppress the per-file summary on stdout"
    )
    return parser


def main(argv: list[str] | None = None) -> int:
    args = build_parser().parse_args(argv)

    if not args.template.exists():
        print(f"error: template not found: {args.template}", file=sys.stderr)
        return EXIT_PARSE

    parsed, issues = [], []
    for path in args.inputs:
        try:
            text, note = _read_source(path)
            po = parse_po(text, path.name)
        except (ParseError, ValueError) as exc:
            print(f"error: {exc}", file=sys.stderr)
            return EXIT_PARSE
        except OSError as exc:
            print(f"error: {path}: {exc}", file=sys.stderr)
            return EXIT_PARSE
        if note:
            print(f"warning: {note}", file=sys.stderr)
        parsed.append((path, po))
        issues.extend(validate(po))

    for issue in issues:
        print(issue, file=sys.stderr)

    warnings = _write(parsed, args, issues)
    for warning in warnings:
        print(f"warning: {warning}", file=sys.stderr)

    if not args.quiet:
        for path, po in parsed:
            checked = ", ".join(match_totals(po).matched) or "nothing"
            print(f"{path.name}: {len(po.items)} line items; "
                  f"TOTALS cross-checked {checked}")

    failed = any(i.severity == ERROR for i in issues) or (
        args.strict and (issues or warnings)
    )
    return EXIT_VALIDATION if failed else EXIT_OK


def _read_source(path: Path) -> tuple[str, str | None]:
    """Decode one PO file. Returns (text, warning-or-None).

    Strict UTF-8 first. Replacing undecodable bytes would corrupt a UPC or a
    description while still reporting success.
    """
    raw = path.read_bytes()
    try:
        return raw.decode("utf-8"), None
    except UnicodeDecodeError:
        pass
    try:
        return raw.decode("cp1252"), f"{path.name}: not valid UTF-8, decoded as cp1252"
    except UnicodeDecodeError as exc:
        raise ValueError(f"{path}: could not decode as UTF-8 or cp1252: {exc}") from None


def _write(parsed, args, issues) -> list[str]:
    """Write the workbooks, carrying each file's findings into its own output."""
    if args.out:
        return write_workbook(
            [po for _, po in parsed], args.template, args.out, args.total_cost,
            issues=issues,
        )
    warnings = []
    for path, po in parsed:
        warnings += write_workbook(
            [po], args.template, path.with_suffix(".xlsx"), args.total_cost,
            issues=[i for i in issues if i.filename == po.filename],
        )
    return warnings


if __name__ == "__main__":
    sys.exit(main())
