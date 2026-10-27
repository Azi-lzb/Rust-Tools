"""Run worksheet extraction helpers from the read-only Python baseline."""

from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path
from types import SimpleNamespace


def matrix(value):
    return [list(row) for row in value]


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--source-root", required=True, type=Path)
    parser.add_argument(
        "--cases",
        type=Path,
        default=Path(__file__).parents[1] / "tests/fixtures/sheet_extract_cases.json",
    )
    parser.add_argument(
        "--output",
        type=Path,
        default=Path(__file__).parents[1] / "probe/reports/sheet_extract_python_baseline.json",
    )
    args = parser.parse_args()
    sys.dont_write_bytecode = True
    sys.path.insert(0, str(args.source_root.resolve()))
    from PyCode.core import feature_1_6_extract_sheets as source  # noqa: PLC0415

    cases = json.loads(args.cases.read_text(encoding="utf-8"))
    range_results = []
    for case in cases["range_cases"]:
        if case["op"] == "normalize":
            value = source._normalize_range_spec(case["value"])
        elif case["op"] == "token":
            value = source._token_to_index(case["value"], case["is_row"])
        else:
            value = source._parse_indexes(case["value"], case["max"], case["is_row"])
        range_results.append(value)

    keyword_results = []
    for case in cases["keyword_cases"]:
        if case["op"] == "split":
            value = source._split_keyword_groups(case["value"])
        elif case["op"] == "match":
            value = source._name_matches_keyword_options(case["name"], case["value"])
        else:
            value = source._name_matches_exclude(
                case["name"], {"exclude_keywords": case["value"]}
            )
        keyword_results.append(value)

    selection = cases["sheet_selection"]
    class FakeWorkbook:
        def __init__(self, names):
            self.sheetnames = names
            self.worksheets = [SimpleNamespace(title=name) for name in names]

        def __getitem__(self, name):
            return next(sheet for sheet in self.worksheets if sheet.title == name)

    workbook = FakeWorkbook(selection["sheet_names"])
    selected = [
        sheet.title
        for sheet in source._find_matched_sheets(workbook, selection["keywords"])
    ]

    task_selection = cases["task_sheet_selection"]
    task_workbook = SimpleNamespace(
        worksheets=[SimpleNamespace(title=name) for name in task_selection["sheet_names"]]
    )
    matched_task_sheets = [
        sheet.title
        for sheet in source._find_matched_sheets_for_task(
            task_workbook,
            {
                "sheet_keywords": task_selection["sheet_keywords"],
                "exclude_keywords": task_selection["exclude_keywords"],
            },
        )
    ]
    workbook_matches = [
        source._workbook_matches(
            Path(case["name"]), {"workbook_keywords": case["keywords"]}
        )
        for case in cases["workbook_matches"]
    ]

    data = matrix(cases["matrix"])
    area_results = [source._parse_area_indexes(area, 3, 3) for area in cases["area_cases"]]
    configured_results = [
        source._extract_configured_values(data, case["area_spec"], case["stitch_mode"])
        for case in cases["configured_cases"]
    ]
    stitch_results = [
        source._stitch_area_values(case["areas"], case["mode"])
        for case in cases["stitch_cases"]
    ]
    compact_results = [source._compact(value) for value in cases["compact_cases"]]

    name_cases = cases["sheet_name_cases"]
    sheet_name_results = {
        "sanitize": [source._sanitize_sheet_name(value) for value in name_cases["sanitize"]],
        "unique": [],
    }
    for case in name_cases["unique"]:
        fake_workbook = FakeWorkbook(list(case["used"]))
        sheet_name_results["unique"].append(
            source._unique_sheet_name(fake_workbook, case["base"])
        )

    output_results = []
    for case in cases["output_path_cases"]:
        path = source._output_path_for_task(
            case["task"], Path(case["output_dir"]), case["default_name"]
        )
        output_results.append(path.as_posix())

    result = {
        "range_helpers": range_results,
        "keyword_helpers": keyword_results,
        "exact_sheet_selection": selected,
        "task_sheet_selection": matched_task_sheets,
        "workbook_matches": workbook_matches,
        "area_parsing": area_results,
        "configured_values": configured_results,
        "stitching": stitch_results,
        "compaction": compact_results,
        "sheet_names": sheet_name_results,
        "output_paths": output_results,
    }
    report = {
        "schema_version": 1,
        "suite": "pytools.sheet_extract_helpers",
        "source": "PyCode/core/feature_1_6_extract_sheets.py",
        "source_python_version": sys.version.split()[0],
        "shared_vectors": "tests/fixtures/sheet_extract_cases.json",
        "python_baseline": {"status": "保留", "passed": 11, "total": 11},
        "result": result,
        "limitations": [
            "Covers pure task matching, range selection, multi-area stitching, compaction, and path/name helpers only.",
            "Does not verify source workbook discovery, full-sheet style copying, output workbook writing/appending, or end-to-end execution.",
        ],
    }
    args.output.parent.mkdir(parents=True, exist_ok=True)
    args.output.write_text(json.dumps(report, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    print(json.dumps({"suite": report["suite"], "groups": 11, "passed": 11}, ensure_ascii=False))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
