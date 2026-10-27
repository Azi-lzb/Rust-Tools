"""Run the read-only Python operation 4-9 on an isolated temporary directory."""

from __future__ import annotations

import argparse
import hashlib
import json
import logging
import sys
import tempfile
from pathlib import Path

from openpyxl import load_workbook


WORKSPACE = Path(__file__).resolve().parents[2]
CASES_PATH = WORKSPACE / "pytools-core/tests/fixtures/file_sort_prefix_cases.json"
CONFIG_PATH = WORKSPACE / "pytools-core/tests/fixtures/file_sort_prefix_config.xlsx"
SOURCE_RELATIVE = Path("PyCode/core/convert_33_34_35_37.py")


def normalize_cell(value: object) -> object:
    if value is None:
        return ""
    if isinstance(value, float) and value.is_integer():
        return int(value)
    return value


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
        / "pytools-core/probe/reports/file_sort_prefix_operation_python_baseline.json",
    )
    args = parser.parse_args()

    sys.dont_write_bytecode = True
    sys.path.insert(0, str(args.source_root))
    from PyCode.core.convert_33_34_35_37 import run_file_sort_prefix_49

    cases = json.loads(CASES_PATH.read_text(encoding="utf-8"))["operation49"]
    with tempfile.TemporaryDirectory(prefix="pytools-sort-prefix-operation-") as raw_root:
        root = Path(raw_root)
        input_dir = root / "input"
        output_dir = root / "output"
        log_dir = root / "logs"
        input_dir.mkdir()

        for name in cases["existing_names"]:
            (input_dir / name).write_bytes(b"occupied target")
        missing_names = set(cases["missing_names"])
        directory_names = set(cases["directory_names"])
        files = []
        for name in cases["input_names"]:
            path = input_dir / name
            files.append(path)
            if name in missing_names:
                continue
            if name in directory_names:
                path.mkdir()
                (path / "marker.txt").write_text("directory", encoding="utf-8")
            else:
                path.write_bytes(b"fixture payload")

        try:
            stats = run_file_sort_prefix_49(CONFIG_PATH, files, output_dir, log_dir)
        finally:
            logger = logging.getLogger("PyCode.4_9_file_sort_prefix")
            for handler in list(logger.handlers):
                handler.close()
                logger.removeHandler(handler)
        workbook = load_workbook(stats["output"], read_only=True, data_only=True)
        try:
            sheet = workbook["文件排序前缀结果"]
            report_rows = []
            for values in sheet.iter_rows(min_row=2, values_only=True):
                if all(value is None for value in values):
                    continue
                report_rows.append(
                    {
                        "source_file": Path(str(values[0])).name,
                        "original_name": normalize_cell(values[1]),
                        "new_name": normalize_cell(values[2]),
                        "config_row": normalize_cell(values[3]),
                        "sort_prefix": normalize_cell(values[4]),
                        "matched_keywords": normalize_cell(values[5]),
                        "status": normalize_cell(values[6]),
                        "remark": normalize_cell(values[7]),
                    }
                )
        finally:
            workbook.close()

        stats = {key: stats[key] for key in (
            "rules", "files", "renamed", "unmatched", "skipped", "failed"
        )}
        final_files = sorted(
            path.name for path in input_dir.iterdir() if path.is_file()
        )

    source_file = args.source_root / SOURCE_RELATIVE
    payload = {
        "source": {
            "project": str(args.source_root),
            "module": SOURCE_RELATIVE.as_posix(),
            "sha256": hashlib.sha256(source_file.read_bytes()).hexdigest().upper(),
            "operation": "run_file_sort_prefix_49",
        },
        "stats": stats,
        "rows": report_rows,
        "final_files": final_files,
    }
    args.output.parent.mkdir(parents=True, exist_ok=True)
    args.output.write_text(
        json.dumps(payload, ensure_ascii=False, indent=2) + "\n", encoding="utf-8"
    )
    print(args.output)


if __name__ == "__main__":
    main()
