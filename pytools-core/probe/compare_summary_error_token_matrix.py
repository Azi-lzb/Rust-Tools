"""Compare PyCode and the Rust PT-6-1 CLI on standard and extended XLSX errors."""

from __future__ import annotations

import datetime as dt
import json
import sys
import tempfile
from pathlib import Path

import compare_summary_tools as parity


TOKENS = (
    "#DIV/0!",
    "#N/A",
    "#NAME?",
    "#NULL!",
    "#NUM!",
    "#REF!",
    "#VALUE!",
    "#DATA!",
    "#SPILL!",
    "#CALC!",
    "#FIELD!",
    "#BLOCKED!",
    "#UNKNOWN!",
    "#CONNECT!",
    "#BUSY!",
)
OUTPUT_CELLS = (
    "D2",
    "E2",
    "F2",
    "G2",
    "H2",
    "I2",
    "J2",
    "K2",
    "L2",
    "M2",
    "N2",
    "O2",
    "P2",
    "Q2",
    "R2",
)


def main() -> int:
    fixture = (
        parity.WORKSPACE
        / "pytools-core/tests/fixtures/summary_used_range/error_token_matrix.xlsx"
    )
    if not parity.PYTOOLS_SOURCE.is_dir():
        raise FileNotFoundError(parity.PYTOOLS_SOURCE)
    if not fixture.is_file():
        raise FileNotFoundError(fixture)

    fingerprints_before = parity.source_fingerprints([fixture])
    with tempfile.TemporaryDirectory(
        prefix="pytools-summary-error-token-matrix-", ignore_cleanup_errors=True
    ) as temp:
        root = Path(temp)
        python_output = parity.python_used_range_only([fixture], root / "python")
        rust_output, rust_report = parity.run_rust(
            "summary-used-range", [fixture], root / "rust"
        )
        if python_output is None or rust_output is None:
            raise AssertionError("PyCode or Rust did not save the error-token matrix summary")

        differences = parity.projection_difference_fields(python_output, rust_output)
        python_sheets, python_dimensions = parity.workbook_projection(python_output)
        rust_sheets, rust_dimensions = parity.workbook_projection(rust_output)
        if differences:
            raise AssertionError(
                f"error-token matrix PyCode/Rust projection differs: {differences}"
            )
        if python_dimensions != [{"name": "汇总", "rows": 2, "columns": 18}]:
            raise AssertionError(f"unexpected PyCode matrix dimensions: {python_dimensions}")

        rust_execution = rust_report.get("execution", {})
        if rust_report.get("rows") != 1 or rust_report.get("files") != 1:
            raise AssertionError(f"unexpected Rust summary result: {rust_report}")
        if rust_execution.get("ooxml_error_cells_patched") != 7:
            raise AssertionError(f"expected seven OOXML-preserved errors: {rust_execution}")
        if rust_execution.get("unknown_error_token_cells_normalized") != 8:
            raise AssertionError(f"expected eight normalized errors: {rust_execution}")

        expected_values = {
            coordinate: token for coordinate, token in zip(OUTPUT_CELLS, TOKENS)
        }
        expected_types = {
            coordinate: ("e" if index < 7 else "s")
            for index, coordinate in enumerate(OUTPUT_CELLS)
        }
        for label, sheets in (("PyCode", python_sheets), ("Rust", rust_sheets)):
            values = sheets[0]["values"]
            types = sheets[0]["types"]
            if any(values.get(cell) != token for cell, token in expected_values.items()):
                raise AssertionError(f"{label} matrix output values do not match the 15 tokens")
            if any(types.get(cell) != kind for cell, kind in expected_types.items()):
                raise AssertionError(f"{label} matrix output types do not match PyCode semantics")

        fingerprints_after = parity.source_fingerprints([fixture])
        if fingerprints_before != fingerprints_after:
            raise AssertionError("the error-token matrix fixture changed during parity probe")

        log = rust_report.get("progress_log", {})
        host, executable = parity.rust_cli_runtime_metadata()
        report = {
            "schema_version": 1,
            "suite": "pytools.summary_used_range_error_token_matrix_parity",
            "date": dt.date.today().isoformat(),
            "status": "passed_python_rust_cli_parity",
            "source_python_file": "PyCode/core/summary_321_322.py::run_summary_by_usedrange",
            "entrypoint": "pytools summary-used-range",
            "fixture": "tests/fixtures/summary_used_range/error_token_matrix.xlsx",
            "fixture_sha256_before_after_equal": True,
            "coverage": {
                "input_error_cells": len(TOKENS),
                "common_excel_error_tokens": list(TOKENS[:7]),
                "extended_error_tokens": list(TOKENS[7:]),
                "python_rust_full_workbook_projection": "exact",
                "output_rows": rust_report["rows"],
                "output_columns": rust_dimensions[0]["columns"],
                "common_errors_preserved_as_error_cells": 7,
                "extended_errors_output_as_text": 8,
            },
            "rust_cli": {
                "runtime": host,
                "executable": executable,
                "reader_engine": rust_execution.get("reader_engine"),
                "writer_engine": rust_execution.get("writer_engine"),
                "ooxml_error_cells_patched": rust_execution.get(
                    "ooxml_error_cells_patched"
                ),
                "unknown_error_token_cells_normalized": rust_execution.get(
                    "unknown_error_token_cells_normalized"
                ),
                "read_elapsed_ms": rust_execution.get("read_elapsed_ms"),
                "write_elapsed_ms": rust_execution.get("write_elapsed_ms"),
                "total_elapsed_ms": rust_execution.get("total_elapsed_ms"),
                "progress_log": {
                    "view": log.get("view"),
                    "status": log.get("status"),
                    "event_count": log.get("event_count"),
                    "event_stream_matches_saved_history": log.get(
                        "event_stream_matches_saved_history"
                    ),
                    "actual_engines": log.get("actual_engines"),
                },
            },
            "limitations": [
                "Synthetic XLSX only; this verifies literal t=e token handling for the listed values.",
                "No behavior is inferred for other unknown tokens, XLS/XLSB, or non-XLSX containers.",
                "This probe does not establish Kylin/UOS, deb, or production-default acceptance.",
            ],
        }
        report_path = (
            parity.WORKSPACE
            / "pytools-core/probe/reports"
            / f"summary_used_range_error_token_matrix_{dt.date.today():%Y%m%d}.json"
        )
        previous_report = (
            json.loads(report_path.read_text(encoding="utf-8"))
            if report_path.is_file()
            else {}
        )
        platform_runs = previous_report.get("platform_runs", {})
        previous_run = previous_report.get("rust_cli")
        if not platform_runs and previous_run:
            previous_key = (
                "windows_gnu_debug"
                if "Windows" in previous_run.get("runtime", "")
                else "wsl_cargo_debug"
            )
            platform_runs[previous_key] = previous_run
        current_key = (
            "windows_gnu_debug"
            if "Windows" in host
            else "wsl_cargo_debug"
            if "WSL" in host
            else "configured_cli"
        )
        platform_runs[current_key] = report["rust_cli"]
        report["platform_runs"] = platform_runs
        report_path.parent.mkdir(parents=True, exist_ok=True)
        report_path.write_text(
            json.dumps(report, ensure_ascii=False, indent=2) + "\n", encoding="utf-8"
        )
        print(
            json.dumps(
                {
                    "status": report["status"],
                    "tokens": len(TOKENS),
                    "common_error_cells_preserved": 7,
                    "extended_tokens_as_text": 8,
                    "reader_engine": rust_execution.get("reader_engine"),
                    "writer_engine": rust_execution.get("writer_engine"),
                    "runtime_key": current_key,
                    "report": str(report_path),
                },
                ensure_ascii=False,
            )
        )
    return 0


if __name__ == "__main__":
    sys.exit(main())
