"""Compare a synthetic BIFF XLS elapsed-time wide summary with PyCode."""

from __future__ import annotations

import hashlib
import json
import os
import shutil
import sys
import tempfile
from datetime import date
from pathlib import Path

import xlrd

WORKSPACE = Path(__file__).resolve().parents[2]
FIXTURE = WORKSPACE / "pytools-core" / "tests" / "fixtures" / "timeline" / "wide_summary_elapsed_duration.xls"
REPORT = WORKSPACE / "pytools-core" / "probe" / "reports" / "wide_summary_biff_elapsed_duration_parity_20261026.json"
RUST_CLI = os.environ.get("PYTOOLS_RUST_CLI")

sys.path.insert(0, str(WORKSPACE / "pytools-core" / "probe"))
import compare_wide_summary as parity  # noqa: E402


def sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for chunk in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest().upper()


def inspect_biff_formats(path: Path) -> list[dict[str, object]]:
    workbook = xlrd.open_workbook(path, formatting_info=True)
    sheet = workbook.sheet_by_name("Report")
    cells = []
    for row, column, address, expected_format in (
        (4, 1, "B5", "[h]:mm:ss.000"),
        (4, 2, "C5", "yyyy/m/d h:mm:ss.000"),
        (5, 1, "B6", "[h]:mm:ss.000"),
        (5, 2, "C6", "yyyy/m/d h:mm:ss.000"),
    ):
        cell = sheet.cell(row, column)
        xf = workbook.xf_list[cell.xf_index]
        number_format = workbook.format_map[xf.format_key].format_str
        if number_format != expected_format:
            raise AssertionError(
                f"{address} BIFF number format {number_format!r} != {expected_format!r}"
            )
        cells.append(
            {
                "cell": address,
                "xlrd_cell_type": cell.ctype,
                "number_format": number_format,
            }
        )
    return cells


def run_case(root: Path, row_suffix: bool) -> dict[str, object]:
    mode = "row_suffix" if row_suffix else "no_suffix"
    case = root / mode
    input_dir = case / "input"
    input_dir.mkdir(parents=True)
    source = input_dir / "sample_2024-06-30.xls"
    shutil.copyfile(FIXTURE, source)

    python_output = case / "python"
    rust_output = case / "rust"
    config = case / "config.xlsx"
    target = rust_output / "target.xlsx"
    parity.make_config(config, rust_output, input_dir, target, [".xls"])
    source_before = source.read_bytes()
    config_before = config.read_bytes()

    baseline, _ = parity.python_baseline(source, python_output, row_suffix)
    actual, plans = parity.run_rust(config, source, row_suffix)
    if baseline is None or actual is None:
        raise AssertionError(f"expected summary output for BIFF XLS mode {mode}")
    parity.assert_parity(baseline, actual)
    if len(plans) != 1:
        raise AssertionError(f"expected one append plan, got {len(plans)}")
    if source.read_bytes() != source_before:
        raise AssertionError("Python/Rust workflow modified its BIFF input")
    if config.read_bytes() != config_before:
        raise AssertionError("Python/Rust workflow modified its synthetic configuration")

    normalized = parity.normalized(baseline)
    return {
        "mode": mode,
        "status": "passed",
        "summary_rows": len(normalized["rows"][0]),
        "summary_columns": len(normalized["rows"][0][0]),
        "python_rust_value_type_format_style_projection": "exact",
        "append_plan_count": len(plans),
        "input_unchanged": True,
        "config_unchanged": True,
    }


def main() -> int:
    if not FIXTURE.is_file():
        raise FileNotFoundError(FIXTURE)
    if not RUST_CLI or not Path(RUST_CLI).is_file():
        raise FileNotFoundError("set PYTOOLS_RUST_CLI to the native PyTools executable")

    format_cells = inspect_biff_formats(FIXTURE)
    with tempfile.TemporaryDirectory(prefix="pytools-wide-biff-duration-") as temporary:
        root = Path(temporary)
        cases = [run_case(root, row_suffix) for row_suffix in (False, True)]

    report = {
        "schema_version": 1,
        "suite": "pytools.timeline.wide_summary.biff_elapsed_duration_parity",
        "date": date.today().isoformat(),
        "status": "passed_synthetic_python_rust_parity",
        "baseline": "PyCode/core/wide_397_398.py::run_wide_summary",
        "python_version": sys.version.split()[0],
        "xlrd_version": xlrd.__version__,
        "rust_entry": "Windows-native pytools CLI wide-summary / wide-summary-row-suffix",
        "fixture": "pytools-core/tests/fixtures/timeline/wide_summary_elapsed_duration.xls",
        "fixture_sha256": sha256(FIXTURE),
        "biff_cells": format_cells,
        "cases": cases,
        "limitations": [
            "The fixture is synthetic; no real business XLS workbook was available.",
            "The result covers this BIFF XLS fixture and Python xlrd semantics, not every XLS number format or input origin.",
            "This does not claim general BIFF number-format metadata exposure from the shared reader.",
        ],
    }
    REPORT.write_text(json.dumps(report, ensure_ascii=True, indent=2) + "\n", encoding="utf-8")
    print(json.dumps({"suite": report["suite"], "status": report["status"], "cases": len(cases)}))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
