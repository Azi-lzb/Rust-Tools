"""Run the read-only Python operation 4-10 on an isolated fixture set."""

from __future__ import annotations

import argparse
import hashlib
import json
import logging
import sys
import tempfile
from pathlib import Path

from openpyxl import Workbook, load_workbook


WORKSPACE = Path(__file__).resolve().parents[2]
CASES_PATH = WORKSPACE / "pytools-core/tests/fixtures/file_sort_prefix_cases.json"
SOURCE_RELATIVE = Path("PyCode/core/convert_33_34_35_37.py")
SORT_COLUMNS = [
    ("是否启用", "enabled"),
    ("排序前缀", "sort_prefix"),
    ("匹配关键词", "keywords"),
    ("备注", "remark"),
]


def make_config(path: Path, cases: dict[str, object]) -> None:
    workbook = Workbook()
    rename_sheet = workbook.active
    rename_sheet.title = "重命名配置"
    rename_sheet.cell(1, 15, "内容重命名启用")
    rename_sheet.cell(1, 16, "内容重命名取值规则")
    rename_sheet.cell(1, 17, "备注")
    rename_sheet.cell(2, 15, "Y")
    rename_sheet.cell(2, 16, cases["operation410"]["content_spec"])
    rename_sheet.cell(2, 17, "内容前缀对拍")

    sort_sheet = workbook.create_sheet("文件排序配置")
    sort_sheet.append([header for header, _ in SORT_COLUMNS])
    for row in cases["config_rows"]:
        sort_sheet.append([row.get(key, "") for _, key in SORT_COLUMNS])

    workbook.save(path)
    workbook.close()


def build_input_files(input_dir: Path, operation: dict[str, object]) -> list[Path]:
    for name in operation["existing_names"]:
        (input_dir / name).write_bytes(b"occupied target")

    missing_names = set(operation["missing_names"])
    directory_names = set(operation["directory_names"])
    corrupt_names = set(operation["corrupt_names"])
    workbook_values = operation["workbook_values"]
    paths = []
    for name in operation["input_names"]:
        path = input_dir / name
        paths.append(path)
        if name in missing_names:
            continue
        if name in directory_names:
            path.mkdir()
            (path / "marker.txt").write_text("directory", encoding="utf-8")
        elif name in corrupt_names:
            path.write_bytes(b"not an xlsx package")
        elif name in workbook_values:
            workbook = Workbook()
            sheet = workbook.active
            sheet.title = "sheet1"
            sheet["A1"] = workbook_values[name]
            workbook.save(path)
            workbook.close()
        else:
            path.write_bytes(b"fixture payload")
    return paths


def normalize_cell(value: object) -> object:
    if value is None:
        return ""
    if isinstance(value, float) and value.is_integer():
        return int(value)
    return value


def close_source_loggers() -> None:
    for name, value in list(logging.root.manager.loggerDict.items()):
        if not name.endswith("4_10_remove_file_sort_prefix"):
            continue
        if not isinstance(value, logging.Logger):
            continue
        for handler in list(value.handlers):
            handler.close()
            value.removeHandler(handler)


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
        default=WORKSPACE
        / "pytools-core/probe/reports/file_sort_prefix_remove_python_baseline.json",
    )
    args = parser.parse_args()

    sys.dont_write_bytecode = True
    sys.path.insert(0, str(args.source_root))
    from PyCode.core.convert_33_34_35_37 import run_remove_file_sort_prefix_410

    cases = json.loads(CASES_PATH.read_text(encoding="utf-8"))
    operation = cases["operation410"]
    with tempfile.TemporaryDirectory(prefix="pytools-sort-prefix-remove-") as raw_root:
        root = Path(raw_root)
        input_dir = root / "input"
        output_dir = root / "output"
        log_dir = root / "logs"
        input_dir.mkdir()
        config_path = root / "config.xlsx"
        make_config(config_path, cases)
        files = build_input_files(input_dir, operation)

        try:
            stats = run_remove_file_sort_prefix_410(
                config_path, files, output_dir, log_dir
            )
        finally:
            close_source_loggers()

        empty_output_dir = root / "empty-output"
        empty_log_dir = root / "empty-logs"
        try:
            empty_stats = run_remove_file_sort_prefix_410(
                config_path, [], empty_output_dir, empty_log_dir
            )
        finally:
            close_source_loggers()

        workbook = load_workbook(stats["output"], read_only=True, data_only=True)
        try:
            sheet = workbook["取消文件排序前缀结果"]
            report_rows = []
            for values in sheet.iter_rows(min_row=2, values_only=True):
                if all(value is None for value in values):
                    continue
                report_rows.append(
                    {
                        "source_file": Path(str(values[0])).name,
                        "original_name": normalize_cell(values[1]),
                        "new_name": normalize_cell(values[2]),
                        "status": normalize_cell(values[3]),
                        "remark": normalize_cell(values[4]),
                    }
                )
        finally:
            workbook.close()

        empty_workbook = load_workbook(
            empty_stats["output"], read_only=True, data_only=True
        )
        try:
            empty_sheet = empty_workbook["取消文件排序前缀结果"]
            empty_report = {
                "sheet_name": empty_sheet.title,
                "max_row": empty_sheet.max_row,
                "max_column": empty_sheet.max_column,
                "cells": [list(row) for row in empty_sheet.values],
            }
        finally:
            empty_workbook.close()

        normalized_stats = {
            key: stats[key]
            for key in ("rules", "files", "renamed", "skipped", "failed")
        }
        final_files = sorted(path.name for path in input_dir.iterdir() if path.is_file())

    source_file = args.source_root / SOURCE_RELATIVE
    payload = {
        "source": {
            "project": str(args.source_root),
            "module": SOURCE_RELATIVE.as_posix(),
            "sha256": hashlib.sha256(source_file.read_bytes()).hexdigest().upper(),
            "operation": "run_remove_file_sort_prefix_410",
        },
        "stats": normalized_stats,
        "rows": report_rows,
        "final_files": final_files,
        "empty_input": {
            "stats": {
                key: empty_stats[key]
                for key in ("rules", "files", "renamed", "skipped", "failed")
            },
            "report": empty_report,
        },
    }
    args.output.parent.mkdir(parents=True, exist_ok=True)
    args.output.write_text(
        json.dumps(payload, ensure_ascii=False, indent=2) + "\n", encoding="utf-8"
    )
    print(args.output)


if __name__ == "__main__":
    main()
