"""Emit a comparison baseline by calling the read-only PyCode implementation."""

from __future__ import annotations

import argparse
import json
import sys
import tempfile
from pathlib import Path

from openpyxl import Workbook


WORKSPACE = Path(__file__).resolve().parents[2]
CASES_PATH = WORKSPACE / "pytools-core/tests/fixtures/file_sort_prefix_cases.json"


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument(
        "--source-root",
        type=Path,
        default=Path(r"C:\Users\AZI\Desktop\VibeCode\4.pytools-main"),
    )
    parser.add_argument(
        "--output",
        type=Path,
        default=WORKSPACE / "pytools-core/probe/reports/file_sort_prefix_python_baseline.json",
    )
    args = parser.parse_args()

    sys.dont_write_bytecode = True
    sys.path.insert(0, str(args.source_root))
    from PyCode.core import config_xlsx, convert_33_34_35_37 as source

    cases = json.loads(CASES_PATH.read_text(encoding="utf-8"))
    rows = cases["config_rows"]

    digits = [source._sort_prefix_digits(value) for value in cases["digit_cases"]]
    widths = [source._sort_prefix_width(value) for value in cases["width_cases"]]
    normalized = [
        source._normalize_sort_prefix(case["value"], case["width"])
        for case in cases["normalize_cases"]
    ]
    stripped = [source._strip_sort_prefix(value) for value in cases["strip_cases"]]

    workbook = Workbook()
    sheet = workbook.active
    sheet.title = config_xlsx.SHEET_FILE_SORT_CONFIG
    sheet.append(["是否启用", "排序前缀", "匹配关键词", "备注"])
    for row in rows:
        sheet.append([row["enabled"], row["sort_prefix"], row["keywords"], row["remark"]])

    with tempfile.TemporaryDirectory(prefix="pytools-sort-prefix-baseline-") as temp_dir:
        config_path = Path(temp_dir) / "file-sort-config.xlsx"
        workbook.save(config_path)
        workbook.close()
        rules = source._load_file_sort_rules(config_path)

    matches = []
    for name in cases["match_names"]:
        matched = source._match_file_sort_rule(name, rules)
        matches.append(
            None
            if matched is None
            else {"row_no": matched["row_no"], "prefix": matched["prefix"]}
        )

    prefixed_names = [
        f"{case['prefix']}_{source._strip_sort_prefix(case['file_name'])}"
        for case in cases["prefix_names"]
    ]

    output = {
        "source": {
            "project": str(args.source_root),
            "module": "PyCode/core/convert_33_34_35_37.py",
            "operation": "_sort_prefix_digits/_sort_prefix_width/_normalize_sort_prefix/_load_file_sort_rules/_match_file_sort_rule/_strip_sort_prefix",
        },
        "sort_prefix_digits": digits,
        "sort_prefix_width": widths,
        "normalize_sort_prefix": normalized,
        "strip_sort_prefix": stripped,
        "rules": rules,
        "matches": matches,
        "prefixed_names": prefixed_names,
    }
    args.output.parent.mkdir(parents=True, exist_ok=True)
    args.output.write_text(json.dumps(output, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    print(args.output)


if __name__ == "__main__":
    main()
