#!/usr/bin/env python3
"""Compare workbook properties not currently exposed by the Rust dump schema."""
from __future__ import annotations

import argparse
import json
from pathlib import Path

from openpyxl import load_workbook


def workbook_snapshot(path: Path):
    workbook = load_workbook(path, data_only=False)
    sheets = []
    for worksheet in workbook.worksheets:
        cells = {
            cell.coordinate: tuple(cell._style) if cell.has_style else ()
            for row in worksheet.iter_rows()
            for cell in row
            if cell.value is not None or cell.has_style
        }
        comments = {
            cell.coordinate: (cell.comment.text, cell.comment.author)
            for row in worksheet.iter_rows()
            for cell in row
            if cell.comment is not None
        }
        rows = {
            str(index): (dimension.height, dimension.hidden, dimension.outlineLevel)
            for index, dimension in worksheet.row_dimensions.items()
            if dimension.has_style or dimension.height is not None or dimension.hidden
        }
        columns = {
            key: (
                dimension.min,
                dimension.max,
                dimension.width,
                dimension.hidden,
                dimension.outlineLevel,
            )
            for key, dimension in worksheet.column_dimensions.items()
        }
        sheets.append(
            {
                "name": worksheet.title,
                "state": worksheet.sheet_state,
                "cells_style": cells,
                "comments": comments,
                "merges": sorted(str(item) for item in worksheet.merged_cells.ranges),
                "row_dimensions": rows,
                "column_dimensions": columns,
                "page_setup": {
                    "orientation": worksheet.page_setup.orientation,
                    "paperSize": worksheet.page_setup.paperSize,
                    "fitToWidth": worksheet.page_setup.fitToWidth,
                    "fitToHeight": worksheet.page_setup.fitToHeight,
                    "scale": worksheet.page_setup.scale,
                },
                "print_area": worksheet.print_area,
                "print_options": {
                    "horizontalCentered": worksheet.print_options.horizontalCentered,
                    "verticalCentered": worksheet.print_options.verticalCentered,
                },
                "header_footer": (
                    worksheet.oddHeader.left.text,
                    worksheet.oddHeader.center.text,
                    worksheet.oddHeader.right.text,
                    worksheet.oddFooter.left.text,
                    worksheet.oddFooter.center.text,
                    worksheet.oddFooter.right.text,
                ),
            }
        )
    workbook.close()
    return sheets


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("before", type=Path)
    parser.add_argument("after", type=Path)
    parser.add_argument("--output", type=Path)
    args = parser.parse_args()
    before = workbook_snapshot(args.before)
    after = workbook_snapshot(args.after)
    failures = []
    if len(before) != len(after):
        failures.append("sheet_count")
    for index, (left, right) in enumerate(zip(before, after)):
        for field in left:
            if left[field] != right[field]:
                failures.append(
                    f"sheet[{index}] {left['name']} {field}: "
                    f"before={left[field]!r}; after={right[field]!r}"
                )
    report = {
        "schema_version": 1,
        "suite": "pytools.institution_normalize_preservation",
        "before": args.before.name,
        "after": args.after.name,
        "status": "保留" if not failures else "损坏",
        "matched_sheets": len(before) if not failures else len(before) - len(failures),
        "total_sheets": len(before),
        "failures": failures,
        "checked": [
            "sheet order and visibility",
            "cell styles",
            "comments and merged ranges",
            "row heights and hidden state",
            "column widths and hidden state",
            "page setup, print area, centering, and header/footer",
        ],
    }
    encoded = json.dumps(report, ensure_ascii=False, indent=2) + "\n"
    if args.output:
        args.output.parent.mkdir(parents=True, exist_ok=True)
        args.output.write_text(encoded, encoding="utf-8")
    print(encoded, end="")
    return 1 if failures else 0


if __name__ == "__main__":
    raise SystemExit(main())
