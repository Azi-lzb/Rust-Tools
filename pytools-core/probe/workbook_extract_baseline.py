"""Compare the original Python workbook extraction with the Rust adapter vectors."""

from __future__ import annotations

import argparse
import json
import sys
from dataclasses import asdict, is_dataclass
from datetime import date, datetime
from pathlib import Path

from openpyxl import Workbook


def json_value(value):
    if is_dataclass(value):
        return json_value(asdict(value))
    if isinstance(value, dict):
        return {key: json_value(item) for key, item in value.items()}
    if isinstance(value, (list, tuple)):
        return [json_value(item) for item in value]
    if isinstance(value, (datetime, date)):
        return value.isoformat()
    if hasattr(value, "item"):
        return json_value(value.item())
    return value


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--source-root", required=True, type=Path)
    parser.add_argument(
        "--cases",
        type=Path,
        default=Path(__file__).parents[1] / "tests/fixtures/workbook_extract_cases.json",
    )
    parser.add_argument(
        "--fixture-dir",
        type=Path,
        default=Path(__file__).parents[1] / "tests/fixtures",
    )
    parser.add_argument(
        "--output",
        type=Path,
        default=Path(__file__).parents[1] / "probe/reports/workbook_extract_baseline.json",
    )
    args = parser.parse_args()
    sys.path.insert(0, str(args.source_root.resolve()))
    from PyCode.core.config_xlsx import PathMapRule, TimelineRule  # noqa: PLC0415
    from PyCode.core.rule_engine import extract_cells  # noqa: PLC0415

    cases = json.loads(args.cases.read_text(encoding="utf-8"))
    fixture = args.fixture_dir.resolve() / cases["filename"]
    fixture.parent.mkdir(parents=True, exist_ok=True)
    workbook = Workbook()
    worksheet = workbook.active
    worksheet.title = cases["sheet_name"]
    for row_index, row in enumerate(cases["sheet"], start=1):
        for column_index, value in enumerate(row, start=1):
            if value is not None:
                worksheet.cell(row_index, column_index, value)
    for merged_range in cases["merges"]:
        worksheet.merge_cells(merged_range)
    workbook.save(fixture)

    rule_case = cases["rule"]
    rule = TimelineRule(
        enabled=True,
        name=rule_case["name"],
        wb_keyword=rule_case["wb_keyword"],
        sheet_keyword=rule_case["sheet_keyword"],
        row_header_col=rule_case["row_header_col"],
        col_header_row=rule_case["col_header_rows"][0],
        row_header_cols=rule_case["row_header_cols"],
        row_header_col_specified=rule_case["row_header_col_specified"],
        col_header_rows=rule_case["col_header_rows"],
        required_col_headers=rule_case["required_col_headers"],
        required_row_headers=rule_case["required_row_headers"],
        data_row_start=rule_case["data_row_start"],
        data_row_end=rule_case["data_row_end"],
        data_col_start=rule_case["data_col_start"],
        data_col_end=rule_case["data_col_end"],
        skip_keywords=rule_case["skip_keywords"],
    )
    path_maps = [PathMapRule(enabled=True, **path_map) for path_map in cases["path_maps"]]
    extracted = extract_cells(fixture, rule, cases["sheet_name"], path_maps, True)
    report = {
        "schema_version": 1,
        "suite": "pytools.workbook_extract",
        "source": "PyCode/core/rule_engine.py::extract_cells",
        "shared_vectors": "tests/fixtures/workbook_extract_cases.json",
        "workbook_fixture": f"tests/fixtures/{fixture.name}",
        "python_baseline": {"status": "保留", "passed": len(extracted), "total": len(extracted)},
        "result": json_value(extracted),
        "limitations": [
            "Covers one synthetic xlsx worksheet, one merge, and one extraction rule.",
            "Does not verify batch discovery, CSV, output workbook writing, target append, or real templates.",
        ],
    }
    args.output.parent.mkdir(parents=True, exist_ok=True)
    args.output.write_text(json.dumps(report, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    print(json.dumps(report, ensure_ascii=False, indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
