from __future__ import annotations

import datetime as dt
import json
import subprocess
import sys
import tempfile
from pathlib import Path

sys.dont_write_bytecode = True
from compare_summary_tools import assert_equivalent, python_comments_only, workbook_projection

WORKSPACE = Path(__file__).resolve().parents[2]
FIXTURES = WORKSPACE / "pytools-core/tests/fixtures/summary_comments"


def main() -> int:
    template = FIXTURES / "error_tokens_template.xlsx"
    source = FIXTURES / "error_tokens_source.xlsx"
    before = {path: path.read_bytes() for path in (template, source)}
    native_cli = WORKSPACE / "target/debug/pytools.exe"
    if not native_cli.is_file():
        raise FileNotFoundError(native_cli)

    with tempfile.TemporaryDirectory(prefix="pytools-pt62-error-token-parity-") as temporary:
        root = Path(temporary)
        python_output = python_comments_only(template, [source], root / "python")
        rust_output_dir = root / "rust/output"
        rust_log_dir = root / "rust/logs"
        completed = subprocess.run(
            [
                str(native_cli),
                "summary-comments",
                "--template",
                str(template),
                "--output-dir",
                str(rust_output_dir),
                "--log-dir",
                str(rust_log_dir),
                "--log-view",
                "developer",
                str(source),
            ],
            check=True,
            capture_output=True,
            text=True,
            encoding="utf-8",
        )
        rust_report = json.loads(completed.stdout)
        rust_output = Path(rust_report["path"])
        assert_equivalent(python_output, rust_output, "PT-6-2 error-token summary")

        python_sheets, python_dimensions = workbook_projection(python_output)
        rust_sheets, rust_dimensions = workbook_projection(rust_output)
        if python_dimensions != rust_dimensions or [s["name"] for s in python_sheets] != [
            "模板",
            "Branch_B",
        ]:
            raise AssertionError("PT-6-2 error-token summary worksheet shape changed")
        for sheet in rust_sheets:
            if sheet["values"].get("D2") != "#DIV/0!" or sheet["types"].get("D2") != "e":
                raise AssertionError(f"known Excel error type was not preserved on {sheet['name']}")
            if sheet["values"].get("E2") != "#DATA!" or sheet["types"].get("E2") != "s":
                raise AssertionError(f"unknown Excel error token was not written as text on {sheet['name']}")

        history_path = Path(rust_report["run_log_path"])
        history = json.loads(history_path.read_text(encoding="utf-8"))
        if history.get("status") != "succeeded":
            raise AssertionError("PT-6-2 developer run history did not succeed")
        phases = {(event["phase"], event["status"]) for event in history["events"]}
        if ("source_read_compatibility", "succeeded") not in phases:
            raise AssertionError("unknown-token compatibility was not recorded")
        if ("ooxml_error_patch", "succeeded") not in phases:
            raise AssertionError("known error-type preservation was not recorded")
        if any(token in history_path.read_text(encoding="utf-8") for token in ("#DATA!", "#DIV/0!")):
            raise AssertionError("developer run history contains raw cell values")

        if any(path.read_bytes() != payload for path, payload in before.items()):
            raise AssertionError("a PT-6-2 input fixture changed")

        report = {
            "schema_version": 1,
            "suite": "pytools.summary_comments_error_token_parity",
            "status": "passed_synthetic_windows_native_python_rust_parity",
            "recorded_at": dt.datetime.now().astimezone().isoformat(),
            "feature_id": "PT-6-2",
            "entrypoint": "pytools summary-comments",
            "fixtures": [
                "tests/fixtures/summary_comments/error_tokens_template.xlsx",
                "tests/fixtures/summary_comments/error_tokens_source.xlsx",
            ],
            "output_worksheets": [sheet["name"] for sheet in rust_sheets],
            "compared_projection": "exact, including values, Excel cell types, styles, dimensions and page settings",
            "error_value_behavior": {
                "known_division_by_zero": "Excel error type preserved on both output worksheets",
                "unknown_data_token": "written as text on both output worksheets, matching PyCode",
            },
            "rust_execution": rust_report.get("execution"),
            "developer_history": {
                "status": history.get("status"),
                "compatibility_read_recorded": True,
                "OOXML_error_patch_recorded": True,
                "raw_cell_values_logged": False,
            },
            "input_fixtures_unchanged": True,
            "scope_limits": [
                "synthetic XLSX fixtures only; no authorized real PT-6-2 template was available",
                "target Linux/deb acceptance remains unverified",
            ],
        }
        report_path = (
            WORKSPACE
            / f"pytools-core/probe/reports/summary_comments_error_tokens_parity_{dt.date.today():%Y%m%d}.json"
        )
        if report_path.exists():
            raise FileExistsError(report_path)
        report_path.write_text(json.dumps(report, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
        print(report_path)
        print(json.dumps(report, ensure_ascii=False, indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
