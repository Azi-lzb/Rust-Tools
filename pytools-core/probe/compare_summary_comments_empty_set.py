"""Focused PyCode/Rust CLI parity for empty set() comment markers."""

from __future__ import annotations

import contextlib
import importlib.util
import json
import logging
import os
import shlex
import subprocess
import sys
import tempfile
from pathlib import Path

sys.dont_write_bytecode = True

from openpyxl import Workbook
from openpyxl.comments import Comment

WORKSPACE = Path(__file__).resolve().parents[2]
PYTOOLS_SOURCE = Path(
    os.environ.get(
        "PYTOOLS_SOURCE_ROOT",
        str(Path(__file__).resolve().parents[3] / "4.pytools-main"),
    )
)
REPORT = WORKSPACE / "pytools-core/probe/reports/summary_comments_empty_set_marker_parity_20261024.json"


def create_template(path: Path) -> None:
    workbook = Workbook()
    worksheet = workbook.active
    worksheet.title = "模板"
    markers = {
        "A2": ("行开始", "行区域1"),
        "A3": ("行结束", "行区域#1"),
        "B1": ("余额", "列区域1"),
        "C1": ("余额", "列区域#1"),
        "D1": ("空集合", "set()"),
    }
    for address, (value, marker) in markers.items():
        worksheet[address] = value
        worksheet[address].comment = Comment(marker, "probe")
    workbook.save(path)
    workbook.close()


def create_source(path: Path) -> None:
    workbook = Workbook()
    worksheet = workbook.active
    worksheet.title = "Branch_A"
    worksheet["B2"] = 10
    worksheet["C2"] = "甲"
    worksheet["B3"] = 20
    worksheet["C3"] = "乙"
    workbook.save(path)
    workbook.close()


def run_rust_cli(
    template: Path,
    source: Path,
    output: Path,
    probe_module: object,
) -> tuple[Path | None, dict[str, object]]:
    native_cli = os.environ.get("PYTOOLS_RUST_CLI")
    if native_cli:
        invocation = [
            native_cli,
            "summary-comments",
            "--template",
            str(template),
            "--output-dir",
            str(output),
            str(source),
        ]
    else:
        command = (
            f"cd {shlex.quote(probe_module.wsl_path(WORKSPACE))} && "
            "cargo run --quiet --offline --locked -p pytools-core --bin pytools -- "
            f"summary-comments --template {shlex.quote(probe_module.wsl_path(template))} "
            f"--output-dir {shlex.quote(probe_module.wsl_path(output))} "
            f"{shlex.quote(probe_module.wsl_path(source))}"
        )
        invocation = ["wsl.exe", "-d", "Ubuntu", "--exec", "bash", "-lc", command]
    completed = subprocess.run(invocation, check=True, capture_output=True)
    stdout = completed.stdout.decode("utf-8-sig", errors="replace").replace("\x00", "")
    for index, character in enumerate(stdout):
        if character == "{":
            try:
                report = json.loads(stdout[index:])
            except json.JSONDecodeError:
                continue
            path = report.get("path")
            resolved_path = (
                Path(path)
                if native_cli and path
                else probe_module.from_wsl_path(path) if path else None
            )
            return resolved_path, report
    stderr = completed.stderr.decode("utf-8-sig", errors="replace").replace("\x00", "")
    raise RuntimeError(f"Rust summary-comments CLI emitted no JSON: {stdout}\n{stderr}")


def main() -> int:
    if not PYTOOLS_SOURCE.is_dir():
        raise FileNotFoundError(PYTOOLS_SOURCE)
    sys.path.insert(0, str(PYTOOLS_SOURCE / "PyCode"))
    from core.summary_321_322 import run_summary_by_comment

    probe_path = WORKSPACE / "pytools-core/probe/compare_summary_tools.py"
    spec = importlib.util.spec_from_file_location("summary_tools_probe", probe_path)
    if spec is None or spec.loader is None:
        raise RuntimeError("could not load the existing summary parity helpers")
    summary_probe = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(summary_probe)

    try:
        with tempfile.TemporaryDirectory(
            prefix="pytools-summary-empty-set-", ignore_cleanup_errors=True
        ) as temp:
            root = Path(temp)
            template = root / "template.xlsx"
            source = root / "source.xlsx"
            create_template(template)
            create_source(source)
            fingerprints_before = [
                (path.stat().st_size, path.stat().st_mtime_ns)
                for path in (template, source)
            ]

            python_output_dir = root / "python"
            python_output_dir.mkdir()
            with contextlib.redirect_stdout(None):
                python_report = run_summary_by_comment(
                    template,
                    [source],
                    python_output_dir,
                    root / "python_logs",
                )
            python_output = Path(python_report["path"])
            rust_output, rust_report = run_rust_cli(
                template, source, root / "rust", summary_probe
            )
            if rust_output is None:
                raise AssertionError("Rust did not save the comment summary")
            summary_probe.assert_equivalent(
                python_output,
                rust_output,
                "PT-6-2 empty set() marker",
            )
            rust_sheets, dimensions = summary_probe.workbook_projection(rust_output)
            headers = rust_sheets[0]["rows"][0]
            expected_headers = ["工作簿", "工作表", "余额", "余额_2", "行号"]
            if headers != expected_headers:
                raise AssertionError(f"unexpected output headers: {headers}")
            if rust_report.get("rows") != 2 or dimensions[0]["columns"] != 5:
                raise AssertionError(
                    "empty set() marker changed the PyCode output shape: "
                    f"rows={rust_report.get('rows')}, dimensions={dimensions}"
                )
            fingerprints_after = [
                (path.stat().st_size, path.stat().st_mtime_ns)
                for path in (template, source)
            ]
            if fingerprints_before != fingerprints_after:
                raise AssertionError("template or source workbook changed")
            execution = rust_report.get("execution", {})
            native_cli = os.environ.get("PYTOOLS_RUST_CLI")
            rust_runtime = {
                "host": os.environ.get(
                    "PYTOOLS_RUST_CLI_HOST",
                    "Windows host" if native_cli else "Ubuntu WSL debug CLI",
                ),
                "entrypoint": "pytools summary-comments",
                "reader_engine": execution.get("reader_engine"),
                "writer_engine": execution.get("writer_engine"),
            }
            if not native_cli:
                rust_runtime["cargo_flags"] = ["--offline", "--locked"]
            result = {
                "feature_id": "PT-6-2",
                "status": "passed_python_rust_cli_parity",
                "scenario": "empty set() comment is ignored; non-empty column markers remain ordered",
                "comparison": "exact full workbook projection, including values, types, styles, dimensions, freeze panes, and page settings",
                "rows": rust_report["rows"],
                "columns": dimensions[0]["columns"],
                "header_contract": "duplicate column names receive _2 suffix and no blank set column is added",
                "template_and_source_unchanged": True,
                "rust_runtime": rust_runtime,
                "python_runtime": (
                    f"{sys.platform} Python {sys.version_info.major}.{sys.version_info.minor}"
                ),
            }
            REPORT.parent.mkdir(parents=True, exist_ok=True)
            REPORT.write_text(
                json.dumps(result, ensure_ascii=False, indent=2) + "\n",
                encoding="utf-8",
            )
            print(json.dumps({"status": result["status"], "rows": result["rows"], "columns": result["columns"], "report": str(REPORT)}, ensure_ascii=False))
            logging.shutdown()
    finally:
        logging.shutdown()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
