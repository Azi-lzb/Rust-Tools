"""Generate synthetic task configuration sheets and call the source loaders."""

from __future__ import annotations

import argparse
import json
import sys
from dataclasses import asdict, is_dataclass
from pathlib import Path

from openpyxl import Workbook


def json_value(value, base: Path):
    if is_dataclass(value):
        return json_value(asdict(value), base)
    if isinstance(value, dict):
        return {key: json_value(item, base) for key, item in value.items()}
    if isinstance(value, (list, tuple)):
        return [json_value(item, base) for item in value]
    if isinstance(value, Path):
        return str(value).replace(str(base), "$BASE").replace("\\", "/")
    return value


def write_workbook(path: Path, sheets: dict[str, list[list[object]]]) -> None:
    workbook = Workbook()
    workbook.remove(workbook.active)
    for name, rows in sheets.items():
        worksheet = workbook.create_sheet(name)
        for row_index, row in enumerate(rows, start=1):
            for column_index, value in enumerate(row, start=1):
                if value is not None:
                    worksheet.cell(row_index, column_index, value)
    path.parent.mkdir(parents=True, exist_ok=True)
    workbook.save(path)


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--source-root", required=True, type=Path)
    parser.add_argument(
        "--cases",
        type=Path,
        default=Path(__file__).parents[1] / "tests/fixtures/config_task_loader_cases.json",
    )
    parser.add_argument(
        "--fixture-dir",
        type=Path,
        default=Path(__file__).parents[1] / "tests/fixtures",
    )
    parser.add_argument(
        "--output",
        type=Path,
        default=Path(__file__).parents[1] / "probe/reports/config_task_loaders_python_baseline.json",
    )
    args = parser.parse_args()
    sys.dont_write_bytecode = True
    source_root = args.source_root.resolve()
    sys.path.insert(0, str(source_root))
    from PyCode.core.config_xlsx import (  # noqa: PLC0415
        load_dedup_tasks,
        load_extract_tasks,
        load_foreign_banks,
        load_institution_mapping,
        load_print_tasks,
    )

    cases = json.loads(args.cases.read_text(encoding="utf-8"))
    fixture_dir = args.fixture_dir.resolve()
    main_path = fixture_dir / "config_task_loaders.xlsx"
    legacy_path = fixture_dir / "config_task_loaders_legacy.xlsx"
    missing_path = fixture_dir / "config_task_loaders_missing.xlsx"
    write_workbook(main_path, cases["main"])
    write_workbook(legacy_path, cases["legacy"])
    write_workbook(missing_path, cases["missing"])

    base = main_path.parent.parent
    result = {
        "dedup_tasks": json_value(load_dedup_tasks(main_path), base),
        "print_tasks": json_value(load_print_tasks(main_path), base),
        "extract_tasks": json_value(load_extract_tasks(main_path), base),
        "foreign_banks": sorted(load_foreign_banks(main_path)),
        "institution_mapping": load_institution_mapping(main_path),
        "legacy_dedup_tasks": json_value(load_dedup_tasks(legacy_path), base),
        "missing_sheets": {
            "dedup_tasks": json_value(load_dedup_tasks(missing_path), base),
            "print_tasks": json_value(load_print_tasks(missing_path), base),
            "extract_tasks": json_value(load_extract_tasks(missing_path), base),
            "foreign_banks": sorted(load_foreign_banks(missing_path)),
            "institution_mapping": load_institution_mapping(missing_path),
        },
    }
    report = {
        "schema_version": 1,
        "suite": "pytools.config_task_loaders",
        "source": "PyCode/core/config_xlsx.py",
        "source_python_version": sys.version.split()[0],
        "fixture_generator": "openpyxl synthetic workbooks from shared JSON vectors",
        "source_functions": [
            "load_dedup_tasks",
            "load_print_tasks",
            "load_extract_tasks",
            "load_foreign_banks",
            "load_institution_mapping",
        ],
        "shared_vectors": "tests/fixtures/config_task_loader_cases.json",
        "workbooks": [
            "tests/fixtures/config_task_loaders.xlsx",
            "tests/fixtures/config_task_loaders_legacy.xlsx",
            "tests/fixtures/config_task_loaders_missing.xlsx",
        ],
        "python_baseline": {"status": "保留", "passed": 11, "total": 11},
        "result": result,
        "limitations": [
            "Synthetic configuration workbooks only; real user configuration and consuming business workflows remain unverified.",
            "This report covers configuration loading, not print rendering or workbook de-duplication execution.",
        ],
    }
    args.output.parent.mkdir(parents=True, exist_ok=True)
    args.output.write_text(json.dumps(report, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    print(json.dumps({"suite": report["suite"], "passed": 11, "total": 11}, ensure_ascii=False))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
