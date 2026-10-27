"""Run read-only pytools operation 4-6 against isolated synthetic workbooks."""

from __future__ import annotations

import argparse
import hashlib
import json
import logging
import sys
import tempfile
from collections import Counter
from pathlib import Path

from openpyxl import Workbook


WORKSPACE = Path(__file__).resolve().parents[2]
CASES_PATH = WORKSPACE / "pytools-core/tests/fixtures/content_rename_46_cases.json"
SOURCE_RELATIVE = Path("PyCode/core/convert_33_34_35_37.py")


def sha256(data: bytes) -> str:
    return hashlib.sha256(data).hexdigest().upper()


def make_config(path: Path, case: dict[str, object]) -> None:
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "重命名配置"
    sheet.cell(1, 15, "内容重命名启用")
    sheet.cell(1, 16, "内容重命名取值规则")
    sheet.cell(1, 17, "备注")
    sheet.cell(2, 15, "Y")
    sheet.cell(2, 16, case["content_spec"])
    sheet.cell(2, 17, case["remark"])
    workbook.save(path)
    workbook.close()


def address_to_row(address: str) -> int:
    digits = next(index for index, char in enumerate(address) if char.isdigit())
    return int(address[digits:])


def address_to_column(address: str) -> int:
    digits = next(index for index, char in enumerate(address) if char.isdigit())
    column = 0
    for char in address[:digits]:
        column = column * 26 + ord(char.upper()) - ord("A") + 1
    return column


def build_inputs(input_dir: Path, case: dict[str, object]) -> list[Path]:
    for name in case.get("existing_names", []):
        (input_dir / name).write_bytes(b"occupied target")

    missing = set(case.get("missing_names", []))
    directories = set(case.get("directory_names", []))
    corrupt = set(case.get("corrupt_names", []))
    workbooks = case["workbooks"]
    paths: list[Path] = []
    for name in case["input_names"]:
        path = input_dir / name
        paths.append(path)
        if name in missing:
            continue
        if name in directories:
            path.mkdir()
            (path / "marker.txt").write_text("directory", encoding="utf-8")
        elif name in corrupt:
            path.write_bytes(b"not a workbook package")
        else:
            workbook_data = workbooks.get(name)
            if workbook_data is None:
                path.write_bytes(b"not a workbook package")
                continue
            workbook = Workbook()
            default_sheet = workbook.active
            for sheet_index, (sheet_name, cells) in enumerate(workbook_data.items()):
                sheet = default_sheet if sheet_index == 0 else workbook.create_sheet(sheet_name)
                sheet.title = sheet_name
                for address, value in cells.items():
                    sheet[address] = value
            workbook.save(path)
            workbook.close()
    return paths


def close_source_loggers() -> None:
    for name, value in list(logging.root.manager.loggerDict.items()):
        if not name.endswith("4_6_batch_rename_by_content"):
            continue
        if not isinstance(value, logging.Logger):
            continue
        for handler in list(value.handlers):
            handler.close()
            value.removeHandler(handler)


def run_case(function, root: Path, case: dict[str, object]) -> dict[str, object]:
    input_dir = root / "input"
    input_dir.mkdir(parents=True)
    config_path = root / "config.xlsx"
    make_config(config_path, case)
    files = build_inputs(input_dir, case)
    before_hashes = {
        path.name: sha256(path.read_bytes())
        for path in files
        if path.is_file()
    }
    try:
        result = function(config_path, files, root / "logs")
    finally:
        close_source_loggers()

    final_entries = sorted(path.name for path in input_dir.iterdir())
    after_hashes = {
        path.name: sha256(path.read_bytes())
        for path in input_dir.iterdir()
        if path.is_file()
    }
    return {
        "stats": {
            key: result[key]
            for key in ("ok", "skip", "spec", "config_row", "config_remark")
        },
        "final_entries": final_entries,
        "input_file_sha256_by_name": dict(sorted(before_hashes.items())),
        "final_file_sha256_by_name": dict(sorted(after_hashes.items())),
        "all_input_file_bytes_preserved": all(
            Counter(after_hashes.values())[digest] >= count
            for digest, count in Counter(before_hashes.values()).items()
        ),
    }


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
        default=WORKSPACE / "pytools-core/probe/reports/content_rename_46_python_baseline.json",
    )
    args = parser.parse_args()

    sys.dont_write_bytecode = True
    sys.path.insert(0, str(args.source_root))
    from PyCode.core.convert_33_34_35_37 import run_batch_rename_by_content_46

    cases = json.loads(CASES_PATH.read_text(encoding="utf-8"))
    with tempfile.TemporaryDirectory(prefix="pytools-content-rename-46-") as raw_root:
        result = run_case(
            run_batch_rename_by_content_46,
            Path(raw_root) / "main-case",
            cases["operation46"],
        )

    source_file = args.source_root / SOURCE_RELATIVE
    payload = {
        "source": {
            "project": str(args.source_root),
            "module": SOURCE_RELATIVE.as_posix(),
            "sha256": sha256(source_file.read_bytes()),
            "operation": "run_batch_rename_by_content_46",
        },
        "operation46": result,
    }
    args.output.parent.mkdir(parents=True, exist_ok=True)
    args.output.write_text(json.dumps(payload, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    print(args.output)


if __name__ == "__main__":
    main()
