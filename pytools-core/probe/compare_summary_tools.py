"""Python/Rust parity probe for PyTools summary workflows 6-1 and 6-2."""

from __future__ import annotations

import contextlib
import datetime as dt
import hashlib
import json
import logging
import math
import os
import posixpath
import shlex
import subprocess
import sys
import tempfile
import time
import xml.etree.ElementTree as ET
from pathlib import Path
from zipfile import ZipFile

from openpyxl import Workbook, load_workbook
from openpyxl.comments import Comment
from openpyxl.styles import PatternFill

sys.dont_write_bytecode = True
WORKSPACE = Path(__file__).resolve().parents[2]
PYTOOLS_SOURCE = Path(r"C:\Users\AZI\Desktop\VibeCode\4.pytools-main")
XLSXWRITER_XLSM_COMMIT = "a880ecc15c1573449dba051e5c21025f00c7c9fc"


def wsl_path(path: Path) -> str:
    return f"/mnt/{path.drive.rstrip(':').lower()}/{path.as_posix()[2:].lstrip('/')}"


def from_wsl_path(value: str) -> Path:
    normalized = value.replace("\\", "/")
    if not normalized.startswith("/mnt/"):
        return Path(value)
    return Path(f"{normalized[5].upper()}:\\{normalized[7:].replace('/', chr(92))}")


def make_source(path: Path) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "Branch_A"
    ws["A1"] = "机构"
    ws["B1"] = "余额"
    ws["C1"] = "余额"
    ws["D1"] = "华东"
    ws["A2"] = "甲"
    ws["B2"] = 10.5
    ws["C2"] = 20
    ws["A3"] = "乙"
    ws["B3"] = dt.date(2026, 6, 30)
    ws["C3"] = True
    ws["A4"] = "丙"
    ws["B4"] = "text"
    ws["C4"] = 0
    for column, error in enumerate(
        ("#DIV/0!", "#N/A", "#NAME?", "#NULL!", "#NUM!", "#REF!", "#VALUE!"), start=5
    ):
        ws.cell(row=4, column=column).value = error
    ws["F6"] = "merged value"
    ws.merge_cells("F6:G6")
    ws["E8"] = "=1+1"
    ws["F8"] = "=1+2"
    ws["H10"] = "dimension tail"
    ws["H10"].fill = PatternFill(fill_type="solid", fgColor="FFFF00")

    matching = wb.create_sheet("模板")
    matching["A1"] = "sheet"
    matching["B2"] = 5
    matching["C2"] = 6
    matching["D1"] = "华南"
    wb.save(path)



def patch_formula_cache(path: Path, address: str, cached_value: str) -> None:
    main_ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
    ET.register_namespace("", main_ns)
    with ZipFile(path, "r") as source_archive:
        entries = [
            (info, source_archive.read(info.filename))
            for info in source_archive.infolist()
        ]

    worksheet_found = False
    patched_entries = []
    for info, payload in entries:
        if info.filename == "xl/worksheets/sheet1.xml":
            root = ET.fromstring(payload)
            cell = root.find(f".//{{{main_ns}}}c[@r='{address}']")
            if cell is None:
                raise AssertionError(f"formula cell {address} is missing from source")
            formula = cell.find(f"{{{main_ns}}}f")
            if formula is None:
                raise AssertionError(f"formula text is missing from {address}")
            cached = cell.find(f"{{{main_ns}}}v")
            if cached is None:
                cached = ET.SubElement(cell, f"{{{main_ns}}}v")
            cached.text = cached_value
            payload = ET.tostring(root, encoding="utf-8", xml_declaration=True)
            worksheet_found = True
        patched_entries.append((info, payload))

    if not worksheet_found:
        raise AssertionError("source worksheet part is missing")
    with ZipFile(path, "w") as output_archive:
        for info, payload in patched_entries:
            output_archive.writestr(info, payload)

    formula_book = load_workbook(path, data_only=False, read_only=False)
    value_book = load_workbook(path, data_only=True, read_only=False)
    try:
        formula_ws = formula_book["Branch_A"]
        value_ws = value_book["Branch_A"]
        if formula_ws[address].value != "=1+1" or value_ws[address].value != 2:
            raise AssertionError("cached formula fixture did not round trip as expected")
        if formula_ws["F8"].value != "=1+2" or value_ws["F8"].value is not None:
            raise AssertionError("missing formula cache fixture did not round trip as expected")
    finally:
        formula_book.close()
        value_book.close()

def make_second_source(path: Path) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "Extra"
    ws["A2"] = "第二个工作簿"
    ws["C2"] = 42
    wb.create_sheet("Empty")
    hidden = wb.create_sheet("Hidden")
    hidden.sheet_state = "hidden"
    hidden["A2"] = "隐藏工作表数据"
    hidden["B3"] = 99
    very_hidden = wb.create_sheet("VeryHidden")
    very_hidden.sheet_state = "veryHidden"
    very_hidden["A4"] = "深度隐藏工作表数据"
    offset = wb.create_sheet("Offset")
    offset["D5"] = "偏移区域起点"
    offset["F6"] = 3.5
    wb.save(path)
    if path.suffix.lower() == ".xlsm":
        mark_macro_enabled(path)


def mark_macro_enabled(path: Path) -> None:
    content_types_ns = "http://schemas.openxmlformats.org/package/2006/content-types"
    with ZipFile(path, "r") as archive:
        entries = [(info, archive.read(info.filename)) for info in archive.infolist()]
    content_types_found = False
    updated_entries = []
    ET.register_namespace("", content_types_ns)
    for info, payload in entries:
        if info.filename == "[Content_Types].xml":
            root = ET.fromstring(payload)
            workbook_type = root.find(
                f"{{{content_types_ns}}}Override[@PartName='/xl/workbook.xml']"
            )
            if workbook_type is None:
                raise AssertionError("workbook content type is missing")
            workbook_type.set(
                "ContentType", "application/vnd.ms-excel.sheet.macroEnabled.main+xml"
            )
            payload = ET.tostring(root, encoding="utf-8", xml_declaration=True)
            content_types_found = True
        updated_entries.append((info, payload))
    if not content_types_found:
        raise AssertionError("package content types are missing")
    with ZipFile(path, "w") as archive:
        for info, payload in updated_entries:
            archive.writestr(info, payload)


def make_comment_template(path: Path, duplicate_row_boundaries: bool = False) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "模板"
    ws["A1"] = "单位"
    ws["A2"] = "区域起"
    ws["A4"] = "区域止"
    ws["A2"].comment = Comment("行区域1", "probe")
    if duplicate_row_boundaries:
        ws["A3"].comment = Comment("行区域1", "probe")
    ws["A4"].comment = Comment("行区域#1", "probe")
    if duplicate_row_boundaries:
        ws["A5"].comment = Comment("行区域#1", "probe")
    ws["B1"] = "余额"
    ws["C1"] = "余额"
    ws["B1"].comment = Comment("列区域1", "probe")
    ws["C1"].comment = Comment("列区域#1", "probe")
    ws["D1"] = "区域"
    ws["D1"].comment = Comment("set(区域)", "probe")
    wb.save(path)


def python_baselines(
    sources: list[Path], comment_source: Path, template: Path, output: Path
) -> tuple[Path | None, Path | None]:
    sys.path.insert(0, str(PYTOOLS_SOURCE / "PyCode"))
    from core.summary_321_322 import run_summary_by_comment, run_summary_by_usedrange

    output.mkdir(parents=True, exist_ok=True)
    with contextlib.redirect_stdout(None):
        used = run_summary_by_usedrange(sources, output / "python_used", output / "python_logs")
        comments = run_summary_by_comment(
            template, [comment_source], output / "python_comments", output / "python_logs"
        )
    logging.shutdown()
    return (
        Path(used["path"]) if used.get("saved") else None,
        Path(comments["path"]) if comments.get("saved") else None,
    )


def python_used_range_only(sources: list[Path], output: Path) -> Path | None:
    sys.path.insert(0, str(PYTOOLS_SOURCE / "PyCode"))
    from core.summary_321_322 import run_summary_by_usedrange

    output.mkdir(parents=True, exist_ok=True)
    with contextlib.redirect_stdout(None):
        result = run_summary_by_usedrange(sources, output / "used", output / "logs")
    logging.shutdown()
    return Path(result["path"]) if result.get("saved") else None


def compare_vba_enabled_xlsm_sample(output: Path) -> dict[str, object]:
    fixture = WORKSPACE / "pytools-core/tests/fixtures/summary_used_range/vba_enabled_xlsxwriter_macro01.xlsm"
    with ZipFile(fixture) as package:
        try:
            vba_project_bytes = package.getinfo("xl/vbaProject.bin").file_size
        except KeyError as error:
            raise AssertionError("VBA-enabled XLSM fixture has no xl/vbaProject.bin") from error
    if vba_project_bytes == 0:
        raise AssertionError("VBA-enabled XLSM fixture has an empty VBA project part")

    source_sha256 = hashlib.sha256(fixture.read_bytes()).hexdigest()
    python_output = python_used_range_only([fixture], output / "python")
    rust_output, rust_report = run_rust(
        "summary-used-range", [fixture], output / "rust"
    )
    if python_output is None or rust_output is None:
        raise AssertionError("PyCode or Rust did not save the VBA-enabled XLSM summary")
    assert_equivalent(python_output, rust_output, "PT-6-1 VBA-enabled XLSM summary")
    if python_output.suffix.lower() != ".xlsx" or rust_output.suffix.lower() != ".xlsx":
        raise AssertionError("summary outputs for XLSM input must be new XLSX workbooks")
    for generated in (python_output, rust_output):
        with ZipFile(generated) as package:
            if "xl/vbaProject.bin" in package.namelist():
                raise AssertionError("new data-only summary unexpectedly copied source VBA")
    if hashlib.sha256(fixture.read_bytes()).hexdigest() != source_sha256:
        raise AssertionError("VBA-enabled XLSM source changed during parity probe")

    progress_log = rust_report.get("progress_log", {})
    if (
        progress_log.get("status") != "succeeded"
        or not progress_log.get("event_stream_matches_saved_history")
        or progress_log.get("completed_stage_elapsed_ms", {}).get("output_workbook") is None
    ):
        raise AssertionError("VBA-enabled XLSM CLI progress log is incomplete")
    execution = rust_report.get("execution", {})
    if rust_report.get("rows") != 1 or execution.get("reader_engine") != "calamine":
        raise AssertionError(f"unexpected Rust result for VBA-enabled XLSM fixture: {rust_report}")

    return {
        "status": "passed_python_rust_cli_parity",
        "fixture": "tests/fixtures/summary_used_range/vba_enabled_xlsxwriter_macro01.xlsm",
        "fixture_origin": {
            "repository": "jmcnamara/XlsxWriter",
            "commit": XLSXWRITER_XLSM_COMMIT,
            "path": "xlsxwriter/test/comparison/xlsx_files/macro01.xlsm",
            "license": "BSD-2-Clause; notice is adjacent to fixture",
            "macro_execution": "not performed",
        },
        "vba_project_bytes": vba_project_bytes,
        "summary_rows": rust_report["rows"],
        "python_rust_output": "exact",
        "new_output_format": "xlsx",
        "source_vba_copied_to_new_summary": False,
        "source_sha256_before_after_equal": True,
        "rust_execution": {
            key: execution.get(key)
            for key in (
                "reader_engine",
                "writer_engine",
                "read_elapsed_ms",
                "write_elapsed_ms",
                "total_elapsed_ms",
            )
        },
        "progress_log": progress_log,
    }


def _real_sample_limitations(rust_cli_executable: str) -> list[str]:
    limitations = [
        "Original-project XLSX assets only; no production default or broad format acceptance inferred.",
        "VBA-bearing XLSM, XLS/XLSB and target Linux/deb remain unvalidated.",
        "Report contains no source cell values or output workbook; both outputs were temporary.",
    ]
    if not (
        sys.platform == "win32"
        and "windows-msvc" in rust_cli_executable.lower()
    ):
        limitations.insert(1, "Windows-native Cargo build remains unvalidated.")
    return limitations


def compare_real_used_range_sample() -> dict[str, object]:
    sample_batches = [
        [
            "2.报表不分口径.xlsx",
            "2.报表校验关系.xlsx",
            "2.报表校验关系(仅公式校验).xlsx",
            "2.报表校验关系((仅公式校验1).xlsx",
        ],
        ["1.报表分口径.xlsx"],
        [
            "4.校验公式v1.xlsx",
            "4.校验公式v2.xlsx",
            "4.校验公式v3.xlsx",
            "4.校验公式（仅公式校验-AI分析版）.xlsx",
        ],
        [
            "4.校验公式（程序生成）.xlsx",
            "4.校验公式（对比）.xlsx",
            "4.校验公式（AI分析版）.xlsx",
            "5.校验关系-规则定位索引.xlsx",
        ],
    ]
    samples = [
        PYTOOLS_SOURCE / "xlsx" / name
        for batch in sample_batches
        for name in batch
    ]
    missing = [sample for sample in samples if not sample.is_file()]
    if missing:
        raise FileNotFoundError(missing[0])
    before = source_fingerprints(samples)
    sys.path.insert(0, str(PYTOOLS_SOURCE / "PyCode"))
    from core.summary_321_322 import _used_bounds as python_used_bounds

    source_summaries = []
    for sample in samples:
        source_workbook = load_workbook(
            sample,
            data_only=True,
            read_only=False,
            keep_vba=sample.suffix.lower() == ".xlsm",
        )
        try:
            bounds = [python_used_bounds(sheet) for sheet in source_workbook.worksheets]
            source_summaries.append({
                "relative_path": sample.relative_to(PYTOOLS_SOURCE).as_posix(),
                "sheet_count": len(source_workbook.worksheets),
                "worksheet_dimensions_missing": missing_worksheet_dimension_count(sample),
                "used_range_rows": sum(
                    max(0, end_row - start_row + 1)
                    for start_row, _, end_row, _ in bounds
                ),
                "max_used_rows": max(
                    (end_row for _, _, end_row, _ in bounds), default=0
                ),
                "max_used_columns": max(
                    (end_column for _, _, _, end_column in bounds), default=0
                ),
            })
        finally:
            source_workbook.close()
    with tempfile.TemporaryDirectory(
        prefix="pytools-real-summary-parity-", ignore_cleanup_errors=True
    ) as temporary:
        root = Path(temporary)
        summaries_by_path = {
            item["relative_path"]: item for item in source_summaries
        }
        batch_reports = []
        for batch_index, batch_names in enumerate(sample_batches, start=1):
            batch_samples = [PYTOOLS_SOURCE / "xlsx" / name for name in batch_names]
            python_started = time.perf_counter()
            python_output = python_used_range_only(
                batch_samples, root / f"python_{batch_index}"
            )
            python_elapsed_ms = round((time.perf_counter() - python_started) * 1000)
            if python_output is None:
                raise AssertionError("PyCode did not save the real-sample used-range batch")

            rust_started = time.perf_counter()
            rust_output, rust_result = run_rust(
                "summary-used-range", batch_samples, root / f"rust_{batch_index}"
            )
            rust_cli_wall_elapsed_ms = round(
                (time.perf_counter() - rust_started) * 1000
            )
            if rust_output is None:
                raise AssertionError("Rust CLI did not save the real-sample used-range batch")
            python_dimensions, rust_dimensions = assert_equivalent(
                python_output,
                rust_output,
                f"PT-6-1 real source-project batch {batch_index}",
            )
            summary_rows = python_dimensions[0]["rows"] - 1
            source_used_range_rows = sum(
                summaries_by_path[
                    sample.relative_to(PYTOOLS_SOURCE).as_posix()
                ]["used_range_rows"]
                for sample in batch_samples
            )
            if (
                rust_result.get("rows") != summary_rows
                or rust_result.get("files") != len(batch_samples)
                or summary_rows != source_used_range_rows
            ):
                raise AssertionError("real-sample used-range row/file counts differ")
            batch_reports.append({
                "source_sample_relative_paths": [
                    sample.relative_to(PYTOOLS_SOURCE).as_posix()
                    for sample in batch_samples
                ],
                "source_workbooks": len(batch_samples),
                "source_used_range_rows": source_used_range_rows,
                "summary_rows": summary_rows,
                "pycode_dimensions": python_dimensions,
                "rust_dimensions": rust_dimensions,
                "python_elapsed_ms": python_elapsed_ms,
                "rust_cli_wall_elapsed_ms": rust_cli_wall_elapsed_ms,
                "rust_execution": rust_result.get("execution"),
                "progress_log": rust_result.get("progress_log"),
            })

        after = source_fingerprints(samples)
        if before != after:
            raise AssertionError("the original PyTools source workbook changed")
        report_day = dt.date.today().strftime("%Y%m%d")
        source_sheet_count = sum(item["sheet_count"] for item in source_summaries)
        source_max_used_rows = max(item["max_used_rows"] for item in source_summaries)
        source_max_used_columns = max(item["max_used_columns"] for item in source_summaries)
        source_used_range_rows = sum(item["used_range_rows"] for item in source_summaries)
        summary_rows = sum(batch["summary_rows"] for batch in batch_reports)
        python_elapsed_ms = sum(batch["python_elapsed_ms"] for batch in batch_reports)
        rust_cli_wall_elapsed_ms = sum(
            batch["rust_cli_wall_elapsed_ms"] for batch in batch_reports
        )
        rust_cli_host, rust_cli_executable = rust_cli_runtime_metadata()
        limitations = _real_sample_limitations(rust_cli_executable)
        report = {
            "schema_version": 1,
            "suite": "pytools.summary_used_range_real_sample_parity",
            "date": dt.date.today().isoformat(),
            "status": "passed_real_sample_cli_parity",
            "data_origin": "read-only workbook assets in the original PyTools project",
            "source_samples": source_summaries,
            "input_format": "xlsx",
            "source_workbooks": len(samples),
            "source_sha256_before_after_equal": True,
            "source_sheet_count": source_sheet_count,
            "source_used_range_rows": source_used_range_rows,
            "source_max_used_rows": source_max_used_rows,
            "source_max_used_columns": source_max_used_columns,
            "summary_rows": summary_rows,
            "batches": batch_reports,
            "progress_logging": {
                "view": "developer",
                "status": "validated",
                "batch_count": len(batch_reports),
                "event_stream_matches_saved_history": True,
                "batches": [batch["progress_log"] for batch in batch_reports],
            },
            "verified_edge_cases": [
                {
                    "kind": "worksheet_dimension_elements_missing",
                    "sample": "xlsx/2.报表校验关系.xlsx",
                    "worksheets_missing_dimension": next(
                        item["worksheet_dimensions_missing"]
                        for item in source_summaries
                        if item["relative_path"] == "xlsx/2.报表校验关系.xlsx"
                    ),
                    "source_used_range_rows": next(
                        item["used_range_rows"]
                        for item in source_summaries
                        if item["relative_path"] == "xlsx/2.报表校验关系.xlsx"
                    ),
                    "max_used_rows": next(
                        item["max_used_rows"]
                        for item in source_summaries
                        if item["relative_path"] == "xlsx/2.报表校验关系.xlsx"
                    ),
                    "max_used_columns": next(
                        item["max_used_columns"]
                        for item in source_summaries
                        if item["relative_path"] == "xlsx/2.报表校验关系.xlsx"
                    ),
                    "comparison": "exact in PyCode/Rust batch output",
                },
                {
                    "kind": "merged_non_anchor_cell",
                    "sample": "xlsx/1.报表分口径.xlsx",
                    "source_cell": "B1",
                    "merged_range": "A1:AG1",
                    "python_output_semantics": "empty",
                    "rust_output_semantics": "empty",
                    "comparison": "exact after filtering non-anchor merged values",
                },
            ],
            "comparison": {
                "cell_values_and_types": "exact",
                "populated_cell_styles": "exact",
                "worksheet_names_and_dimensions": "exact",
                "freeze_panes_and_header_bold": "exact",
                "formula_merge_name_and_page_properties": "exact",
            },
            "python_elapsed_ms": python_elapsed_ms,
            "rust_cli_wall_elapsed_ms": rust_cli_wall_elapsed_ms,
            "execution_environment": {
                "python_sys_platform": sys.platform,
                "rust_cli_host": rust_cli_host,
                "rust_cli_executable": rust_cli_executable,
                "rust_cli_wall_includes_cargo_run": not bool(
                    os.environ.get("PYTOOLS_RUST_CLI")
                ),
            },
            "timing_note": "Single execution per batch for diagnostics; not a performance comparison.",
            "limitations": limitations,
        }
        report_path = (
            Path(__file__).resolve().parent
            / "reports"
            / f"summary_used_range_real_parity_{report_day}.json"
        )
        previous_report = (
            json.loads(report_path.read_text(encoding="utf-8"))
            if report_path.is_file()
            else {}
        )
        for key, value in previous_report.items():
            if key not in report:
                report[key] = value
        report_path.parent.mkdir(parents=True, exist_ok=True)
        report_path.write_text(
            json.dumps(report, ensure_ascii=False, indent=2) + "\n", encoding="utf-8"
        )
        return {
            "status": report["status"],
            "source_workbooks": report["source_workbooks"],
            "source_sheet_count": report["source_sheet_count"],
            "source_used_range_rows": report["source_used_range_rows"],
            "source_max_used_columns": report["source_max_used_columns"],
            "summary_rows": report["summary_rows"],
            "source_sha256_before_after_equal": report["source_sha256_before_after_equal"],
            "python_elapsed_ms": python_elapsed_ms,
            "rust_cli_wall_elapsed_ms": rust_cli_wall_elapsed_ms,
            "execution_environment": report["execution_environment"],
            "batch_count": len(batch_reports),
            "rust_execution": [batch["rust_execution"] for batch in batch_reports],
            "report": str(report_path),
        }


def missing_worksheet_dimension_count(path: Path) -> int:
    main_ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
    office_rel_ns = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
    package_rel_ns = "http://schemas.openxmlformats.org/package/2006/relationships"
    with ZipFile(path) as archive:
        workbook = ET.fromstring(archive.read("xl/workbook.xml"))
        relationships = ET.fromstring(archive.read("xl/_rels/workbook.xml.rels"))
        targets = {
            relation.get("Id"): relation.get("Target")
            for relation in relationships.findall(f"{{{package_rel_ns}}}Relationship")
        }
        missing = 0
        for sheet in workbook.findall(f".//{{{main_ns}}}sheet"):
            relationship_id = sheet.get(f"{{{office_rel_ns}}}id")
            target = targets.get(relationship_id)
            if target is None:
                raise AssertionError("worksheet relationship target is missing")
            part = (
                target.lstrip("/")
                if target.startswith("/")
                else posixpath.normpath(posixpath.join("xl", target))
            )
            worksheet = ET.fromstring(archive.read(part))
            if worksheet.find(f"{{{main_ns}}}dimension") is None:
                missing += 1
        return missing


def python_comments_only(template: Path, sources: list[Path], output: Path) -> Path | None:
    sys.path.insert(0, str(PYTOOLS_SOURCE / "PyCode"))
    from core.summary_321_322 import run_summary_by_comment

    output.mkdir(parents=True, exist_ok=True)
    with contextlib.redirect_stdout(None):
        result = run_summary_by_comment(
            template, sources, output / "comments", output / "logs"
        )
    logging.shutdown()
    return Path(result["path"]) if result.get("saved") else None


def python_nocom_used_range(sources: list[Path], output: Path) -> Path | None:
    sys.path.insert(0, str(PYTOOLS_SOURCE / "NoComPyTools"))
    from nocompytools.core.summary_321_322 import run_summary_by_usedrange

    with contextlib.redirect_stdout(None):
        result = run_summary_by_usedrange(sources, output / "nocom_used", output / "nocom_logs")
    logging.shutdown()
    return Path(result["path"]) if result.get("saved") else None


def run_rust(command_name: str, args: list[Path], output: Path) -> tuple[Path | None, dict[str, object]]:
    native_cli = os.environ.get("PYTOOLS_RUST_CLI")
    if native_cli:
        command = [native_cli, command_name, "--output-dir", str(output)]
        if command_name == "summary-used-range":
            command.extend(
                ["--log-dir", str(output / "logs"), "--log-view", "developer"]
            )
        command.extend(str(path) for path in args)
        invocation = command
    else:
        cli = " ".join(shlex.quote(wsl_path(path)) for path in args)
        log_options = ""
        if command_name == "summary-used-range":
            log_options = (
                f"--log-dir {shlex.quote(wsl_path(output / 'logs'))} "
                "--log-view developer "
            )
        command = (
            f"cd {shlex.quote(wsl_path(WORKSPACE))} && cargo run --quiet -p pytools-core --bin pytools -- "
            f"{command_name} --output-dir {shlex.quote(wsl_path(output))} {log_options}{cli}"
        )
        invocation = ["wsl", "-e", "sh", "-lc", command]
    try:
        completed = subprocess.run(invocation, check=True, capture_output=True)
    except subprocess.CalledProcessError as error:
        stderr = error.stderr.decode("utf-8-sig", errors="replace") if error.stderr else ""
        raise RuntimeError(
            "Rust summary CLI failed: " + stderr[-4000:]
        ) from error
    stdout = completed.stdout.decode("utf-8-sig", errors="replace").replace("\x00", "")
    stderr = completed.stderr.decode("utf-8-sig", errors="replace").replace("\x00", "")
    for index, character in enumerate(stdout):
        if character != "{":
            continue
        try:
            result = json.loads(stdout[index:])
            path = result.get("path")
            if command_name == "summary-used-range":
                events = []
                for line in stderr.splitlines():
                    try:
                        event = json.loads(line)
                    except json.JSONDecodeError:
                        continue
                    if isinstance(event, dict) and "phase" in event:
                        events.append(event)
                if not events:
                    raise RuntimeError("Rust CLI emitted no developer progress events")
                if [event.get("sequence") for event in events] != list(
                    range(1, len(events) + 1)
                ):
                    raise RuntimeError("Rust CLI progress event sequence is not contiguous")
                log_path_value = result.get("run_log_path")
                if not log_path_value:
                    raise RuntimeError("Rust CLI omitted its PT-6-1 run log path")
                log_path = from_wsl_path(log_path_value)
                with log_path.open("r", encoding="utf-8") as log_file:
                    run_log = json.load(log_file)
                if run_log.get("status") != "succeeded":
                    raise RuntimeError("Rust CLI PT-6-1 run log did not record success")
                if run_log.get("events", []) != events:
                    raise RuntimeError("Rust CLI progress stream and saved run log differ")
                input_filter_events = [
                    event for event in events if event.get("phase") == "input_filter"
                ]
                result["progress_log"] = {
                    "view": "developer",
                    "status": run_log["status"],
                    "event_stream_matches_saved_history": True,
                    "event_count": len(events),
                    "phases": sorted({event["phase"] for event in events}),
                    "actual_engines": sorted(
                        {
                            event["actual_engine"]
                            for event in events
                            if event.get("actual_engine")
                        }
                    ),
                    "completed_stage_elapsed_ms": {
                        event["phase"]: event["stage_elapsed_ms"]
                        for event in events
                        if event.get("status") == "succeeded"
                        and event.get("stage_elapsed_ms") is not None
                    },
                    "skipped_input_names": [
                        event.get("file_name")
                        for event in input_filter_events
                        if event.get("status") == "skipped"
                    ],
                    "input_filter_events_have_no_engine": all(
                        event.get("actual_engine") is None
                        for event in input_filter_events
                    ),
                }
            return (from_wsl_path(path) if path else None), result
        except json.JSONDecodeError:
            continue
    raise RuntimeError(f"Rust summary CLI emitted no JSON: {stdout}\n{stderr}")


def compare_used_range_extension_filter(
    sources: list[Path], root: Path
) -> dict[str, object]:
    root.mkdir(parents=True, exist_ok=True)
    unsupported = []
    for extension in (".xls", ".xlsb"):
        path = root / f"ignored_source{extension}"
        path.write_bytes(b"not a workbook; expected to be skipped by the PyCode contract")
        unsupported.append(path)

    all_sources = [*sources, *unsupported]
    before = source_fingerprints(all_sources)
    python_output = python_used_range_only(all_sources, root / "python_extension_filter")
    rust_output, rust_report = run_rust(
        "summary-used-range", all_sources, root / "rust_extension_filter"
    )
    assert_equivalent(python_output, rust_output, "6-1 input extension filter")
    if rust_report.get("files") != len(sources):
        raise AssertionError("Rust summary did not skip unsupported .xls/.xlsb inputs")
    if rust_report.get("rows") != 19:
        raise AssertionError("unsupported inputs changed the summary row count")
    progress_log = rust_report.get("progress_log", {})
    expected_skipped = sorted(path.name for path in unsupported)
    if progress_log.get("skipped_input_names") != expected_skipped:
        raise AssertionError("Rust CLI did not report both unsupported input extensions")
    if progress_log.get("input_filter_events_have_no_engine") is not True:
        raise AssertionError("unsupported inputs were associated with a reader engine")
    if progress_log.get("completed_stage_elapsed_ms", {}).get("output_workbook") is None:
        raise AssertionError("Rust CLI did not record the completed workbook-write duration")
    if before != source_fingerprints(all_sources):
        raise AssertionError("an extension-filter input was modified")
    return {
        "python_supported_extensions": [".xlsx", ".xlsm"],
        "unsupported_extensions_skipped": [".xls", ".xlsb"],
        "python_rust_output": "exact",
        "files_summarized": rust_report["files"],
        "summary_rows": rust_report["rows"],
        "skipped_input_names": expected_skipped,
        "skipped_inputs_used_no_engine": True,
        "progress_log": progress_log,
    }


def run_rust_comments(template: Path, source: Path, output: Path) -> tuple[Path | None, dict[str, object]]:
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
            f"cd {shlex.quote(wsl_path(WORKSPACE))} && cargo run --quiet -p pytools-core --bin pytools -- "
            f"summary-comments --template {shlex.quote(wsl_path(template))} "
            f"--output-dir {shlex.quote(wsl_path(output))} {shlex.quote(wsl_path(source))}"
        )
        invocation = ["wsl", "-e", "sh", "-lc", command]
    completed = subprocess.run(invocation, check=True, capture_output=True)
    stdout = completed.stdout.decode("utf-8-sig", errors="replace").replace("\x00", "")
    stderr = completed.stderr.decode("utf-8-sig", errors="replace").replace("\x00", "")
    for index, character in enumerate(stdout):
        if character != "{":
            continue
        try:
            result = json.loads(stdout[index:])
            path = result.get("path")
            return (from_wsl_path(path) if path else None), result
        except json.JSONDecodeError:
            continue
    raise RuntimeError(f"Rust comment-summary CLI emitted no JSON: {stdout}\n{stderr}")


def normalize(value: object) -> object:
    if isinstance(value, (dt.datetime, dt.date, dt.time)):
        return value.isoformat()
    if isinstance(value, float) and not math.isfinite(value):
        return str(value)
    return value


def workbook_projection(path: Path) -> tuple[list[dict[str, object]], list[dict[str, object]]]:
    wb = load_workbook(path, data_only=False)
    sheets = []
    for ws in wb.worksheets:
        cells = [cell for row in ws.iter_rows() for cell in row]
        values = {
            cell.coordinate: normalize(cell.value)
            for cell in cells
            if cell.value is not None
        }
        types = {
            cell.coordinate: cell.data_type
            for cell in cells
            if cell.value is not None
        }
        styles = {
            cell.coordinate: {
                "number_format": cell.number_format,
                "bold": bool(cell.font.bold),
                "fill_type": cell.fill.fill_type,
                "fill_color": cell.fill.fgColor.rgb if cell.fill.fgColor.type == "rgb" else None,
            }
            for cell in cells
            if cell.value is not None
        }
        setup_properties = ws.sheet_properties.pageSetUpPr
        rows = [
            [normalize(cell.value) for cell in row]
            for row in ws.iter_rows()
        ]
        sheets.append({
            "name": ws.title,
            "rows": rows,
            "values": values,
            "types": types,
            "styles": styles,
            "freeze_panes": ws.freeze_panes,
            "header_bold": [ws.cell(1, column).font.bold for column in range(1, ws.max_column + 1)],
            "merged_ranges": sorted(str(item) for item in ws.merged_cells.ranges),
            "formula_cells": {
                cell.coordinate: cell.value
                for cell in cells
                if cell.data_type == "f"
            },
            "defined_names": sorted(wb.defined_names.keys()),
            "print_area": str(ws.print_area) if ws.print_area else None,
            "page_setup": {
                "orientation": ws.page_setup.orientation or "portrait",
                "fitToPage": setup_properties.fitToPage if setup_properties else None,
                "fitToWidth": ws.page_setup.fitToWidth,
                "fitToHeight": ws.page_setup.fitToHeight,
                "paperSize": ws.page_setup.paperSize,
                "scale": ws.page_setup.scale,
            },
            "page_margins": {
                "left": ws.page_margins.left,
                "right": ws.page_margins.right,
                "top": ws.page_margins.top,
                "bottom": ws.page_margins.bottom,
                "header": ws.page_margins.header,
                "footer": ws.page_margins.footer,
            },
            "horizontal_centered": ws.print_options.horizontalCentered,
            "vertical_centered": ws.print_options.verticalCentered,
        })
    wb.close()
    return sheets, [
        {"name": sheet["name"], "rows": len(sheet["rows"]), "columns": max((len(row) for row in sheet["rows"]), default=0)}
        for sheet in sheets
    ]


def assert_equivalent(
    expected: Path | None,
    actual: Path | None,
    label: str,
    compare_page_settings: bool = True,
) -> tuple[list[dict[str, object]], list[dict[str, object]]]:
    if expected is None or actual is None:
        if expected is not actual:
            raise AssertionError(f"{label}: saved status differs Python={expected}, Rust={actual}")
        return [], []
    expected_sheets, expected_dimensions = workbook_projection(expected)
    actual_sheets, actual_dimensions = workbook_projection(actual)
    if not compare_page_settings:
        omitted_fields = {
            "page_setup",
            "page_margins",
            "horizontal_centered",
            "vertical_centered",
        }
        expected_sheets = [
            {key: value for key, value in sheet.items() if key not in omitted_fields}
            for sheet in expected_sheets
        ]
        actual_sheets = [
            {key: value for key, value in sheet.items() if key not in omitted_fields}
            for sheet in actual_sheets
        ]
    if expected_sheets != actual_sheets or expected_dimensions != actual_dimensions:
        field_differences = []
        for expected_sheet, actual_sheet in zip(expected_sheets, actual_sheets):
            field_differences.extend(
                f"{expected_sheet['name']}.{field}"
                for field in expected_sheet
                if expected_sheet.get(field) != actual_sheet.get(field)
            )
            if expected_sheet.get("styles") != actual_sheet.get("styles"):
                expected_styles = expected_sheet.get("styles", {})
                actual_styles = actual_sheet.get("styles", {})
                style_cells = sorted(set(expected_styles) | set(actual_styles))
                style_differences = {
                    cell: {"python": expected_styles.get(cell), "rust": actual_styles.get(cell)}
                    for cell in style_cells
                    if expected_styles.get(cell) != actual_styles.get(cell)
                }
            else:
                style_differences = {}
        if expected_dimensions != actual_dimensions:
            field_differences.append("worksheet_dimensions")
        raise AssertionError(json.dumps({
            "label": label,
            "different_fields": field_differences,
            "style_differences": style_differences,
            "python_page_setup": expected_sheets[0].get("page_setup"),
            "rust_page_setup": actual_sheets[0].get("page_setup"),
            "python_page_margins": expected_sheets[0].get("page_margins"),
            "rust_page_margins": actual_sheets[0].get("page_margins"),
            "python_dimensions": expected_dimensions,
            "rust_dimensions": actual_dimensions,
        }, ensure_ascii=False, default=str, indent=2))
    return expected_dimensions, actual_dimensions


def source_fingerprints(paths: list[Path]) -> dict[str, str]:
    return {
        path.name: hashlib.sha256(path.read_bytes()).hexdigest()
        for path in paths
    }


def projection_difference_fields(expected: Path, actual: Path) -> list[str]:
    expected_sheets, expected_dimensions = workbook_projection(expected)
    actual_sheets, actual_dimensions = workbook_projection(actual)
    differences = []
    for expected_sheet, actual_sheet in zip(expected_sheets, actual_sheets):
        differences.extend(
            f"{expected_sheet['name']}.{field}"
            for field in expected_sheet
            if expected_sheet.get(field) != actual_sheet.get(field)
        )
    if expected_dimensions != actual_dimensions:
        differences.append("worksheet_dimensions")
    return sorted(differences)




def projection_difference_summary(expected: Path, actual: Path) -> dict[str, object]:
    expected_sheets, expected_dimensions = workbook_projection(expected)
    actual_sheets, actual_dimensions = workbook_projection(actual)
    cell_differences: dict[str, dict[str, object]] = {}
    for expected_sheet, actual_sheet in zip(expected_sheets, actual_sheets):
        for field in ("values", "types", "styles"):
            expected_cells = expected_sheet.get(field, {})
            actual_cells = actual_sheet.get(field, {})
            coordinates = sorted(set(expected_cells) | set(actual_cells))
            changed = [
                coordinate
                for coordinate in coordinates
                if expected_cells.get(coordinate) != actual_cells.get(coordinate)
            ]
            if changed:
                cell_differences[f"{expected_sheet['name']}.{field}"] = {
                    "count": len(changed),
                    "coordinates_sample": changed[:12],
                    "examples": [
                        {
                            "cell": coordinate,
                            "pycode": expected_cells.get(coordinate),
                            "nocom_pytools": actual_cells.get(coordinate),
                        }
                        for coordinate in changed[:6]
                    ],
                }
    return {
        "difference_fields": projection_difference_fields(expected, actual),
        "cell_differences": cell_differences,
        "pycode_dimensions": expected_dimensions,
        "nocom_pytools_dimensions": actual_dimensions,
    }


def rust_cli_runtime_metadata() -> tuple[str, str]:
    native_cli = os.environ.get("PYTOOLS_RUST_CLI")
    host = os.environ.get("PYTOOLS_RUST_CLI_HOST")
    executable = os.environ.get("PYTOOLS_RUST_CLI_LABEL")
    if host is None:
        if not native_cli:
            host = "WSL Linux"
        elif sys.platform == "win32":
            host = "Windows host"
        elif sys.platform == "linux":
            host = "Linux host"
        else:
            host = sys.platform
    if executable is None:
        if not native_cli:
            executable = "cargo run"
        elif sys.platform == "win32":
            executable = "configured Windows GNU executable"
        elif sys.platform == "linux":
            executable = "configured Linux executable"
        else:
            executable = "configured executable"
    return host, executable


def main() -> int:
    if not PYTOOLS_SOURCE.is_dir():
        raise FileNotFoundError(PYTOOLS_SOURCE)
    with tempfile.TemporaryDirectory(prefix="pytools-summary-parity-", ignore_cleanup_errors=True) as temp:
        root = Path(temp)
        source = root / "sample_2026-06-30.xlsx"
        second_source = root / "sample_2026-07-31.xlsm"
        template = root / "comment_template.xlsx"
        duplicate_marker_template = root / "comment_template_duplicate_marker.xlsx"
        merged_comment_source = (
            WORKSPACE / "pytools-core/tests/fixtures/summary_comments/merged_children_source.xlsx"
        )
        unsupported_error_source = WORKSPACE / "pytools-core/tests/fixtures/summary_used_range/getting_data_error.xlsx"
        make_source(source)
        patch_formula_cache(source, "E8", "2")
        make_second_source(second_source)
        make_comment_template(template)
        make_comment_template(duplicate_marker_template, duplicate_row_boundaries=True)
        sources = [source, second_source]
        fingerprints_before = source_fingerprints(
            [
                *sources,
                template,
                duplicate_marker_template,
                merged_comment_source,
                unsupported_error_source,
            ]
        )
        python_used, python_comments = python_baselines(sources, source, template, root)
        nocom_used = python_nocom_used_range(sources, root)
        rust_used, used_report = run_rust("summary-used-range", sources, root / "rust_used")
        extension_filter_report = compare_used_range_extension_filter(
            sources, root / "extension_filter_probe"
        )
        vba_enabled_xlsm_report = compare_vba_enabled_xlsm_sample(
            root / "vba_enabled_xlsm_probe"
        )
        rust_comments, comments_report = run_rust_comments(template, source, root / "rust_comments")
        assert_equivalent(python_used, rust_used, "6-1 used-range summary")
        used_sheets, _ = workbook_projection(rust_used)
        offset_values = {
            address: used_sheets[0]["values"].get(address)
            for address in ("A19", "B19", "C19", "G19", "A20", "B20", "C20", "I20")
        }
        expected_offset_values = {
            "A19": "sample_2026-07-31",
            "B19": "Offset",
            "C19": 5,
            "G19": "偏移区域起点",
            "A20": "sample_2026-07-31",
            "B20": "Offset",
            "C20": 6,
            "I20": 3.5,
        }
        if offset_values != expected_offset_values:
            raise AssertionError(
                f"offset used-range coordinates changed: {offset_values}"
            )
        assert_equivalent(
            python_comments,
            rust_comments,
            "6-2 comment summary",
        )
        python_merged_comments = python_comments_only(
            template, [merged_comment_source], root / "python_merged_comments"
        )
        rust_merged_comments, merged_comments_report = run_rust_comments(
            template, merged_comment_source, root / "rust_merged_comments"
        )
        if python_merged_comments is None or rust_merged_comments is None:
            raise AssertionError("PyCode or Rust did not save the merged-cell comment summary")
        assert_equivalent(
            python_merged_comments,
            rust_merged_comments,
            "PT-6-2 merged non-anchor cell behavior",
        )
        merged_sheets, _ = workbook_projection(rust_merged_comments)
        merged_values = {
            address: merged_sheets[0]["values"].get(address)
            for address in ("D2", "E2", "D3", "E3")
        }
        if (
            merged_comments_report.get("rows") != 3
            or set(merged_values.values()) != {"merged anchor"}
        ):
            raise AssertionError(
                "merged source child cells must resolve to the PyCode anchor value: "
                f"rows={merged_comments_report.get('rows')}, cells={merged_values}"
            )
        python_duplicate_comments = python_comments_only(
            duplicate_marker_template, [source], root / "python_duplicate_comments"
        )
        rust_duplicate_comments, duplicate_comments_report = run_rust_comments(
            duplicate_marker_template, source, root / "rust_duplicate_comments"
        )
        if python_duplicate_comments is None:
            raise AssertionError("PyCode did not save the duplicate-marker comment summary")
        assert_equivalent(
            python_duplicate_comments,
            rust_duplicate_comments,
            "6-2 duplicate row-region marker behavior",
        )
        duplicate_sheets, _ = workbook_projection(rust_duplicate_comments)
        duplicate_source_rows = [
            duplicate_sheets[0]["values"].get(f"F{row}") for row in range(2, 8)
        ]
        if duplicate_comments_report.get("rows") != 6 or duplicate_source_rows != [3, 4, 5, 3, 4, 5]:
            raise AssertionError(
                "duplicate row-region markers should select the final start marker: "
                f"rows={duplicate_comments_report.get('rows')}, values={duplicate_source_rows}"
            )
        python_unknown_error = python_used_range_only(
            [unsupported_error_source], root / "python_unknown_error"
        )
        if python_unknown_error is None:
            raise AssertionError("PyCode did not save the unknown-error summary")
        unknown_error_sheets, _ = workbook_projection(python_unknown_error)
        python_unknown_cell = unknown_error_sheets[0]
        if (
            python_unknown_cell["values"].get("D2") != "#DATA!"
            or python_unknown_cell["types"].get("D2") != "s"
        ):
            raise AssertionError("PyCode unknown-error behavior changed from string output")
        rust_unknown_error, rust_unknown_report = run_rust(
            "summary-used-range",
            [unsupported_error_source],
            root / "rust_unknown_error",
        )
        if rust_unknown_error is None or rust_unknown_report.get("saved") is not True:
            raise AssertionError("Rust CLI did not save the unknown-error summary")
        assert_equivalent(
            python_unknown_error,
            rust_unknown_error,
            "PT-6-1 unknown error-token summary",
        )
        rust_unknown_sheets, _ = workbook_projection(rust_unknown_error)
        rust_unknown_cell = rust_unknown_sheets[0]
        if (
            rust_unknown_cell["values"].get("D2") != "#DATA!"
            or rust_unknown_cell["types"].get("D2") != "s"
        ):
            raise AssertionError("Rust unknown-error output must preserve #DATA! as string")
        rust_unknown_execution = rust_unknown_report.get("execution", {})
        if (
            rust_unknown_execution.get("reader_engine")
            != "calamine + OOXML error-token normalization"
            or rust_unknown_execution.get("unknown_error_token_cells_normalized") != 1
            or rust_unknown_execution.get("ooxml_error_cells_patched") != 0
        ):
            raise AssertionError(
                "Rust unknown-error route did not use the narrow shared reader API: "
                + str(rust_unknown_execution)
            )
        if not rust_unknown_report.get("progress_log", {}).get(
            "event_stream_matches_saved_history"
        ):
            raise AssertionError("unknown-error developer events do not match history")
        rust_unknown_log_path = from_wsl_path(rust_unknown_report["run_log_path"])
        rust_unknown_log_bytes = rust_unknown_log_path.read_bytes()
        if b"#DATA!" in rust_unknown_log_bytes:
            raise AssertionError("unknown-error history contains the cell token")
        rust_unknown_log = json.loads(rust_unknown_log_bytes)
        compatibility_events = [
            event
            for event in rust_unknown_log.get("events", [])
            if event.get("phase") == "value_read_compatibility"
            and event.get("status") == "succeeded"
        ]
        if (
            len(compatibility_events) != 1
            or compatibility_events[0].get("actual_engine")
            != "calamine + OOXML error-token normalization"
            or compatibility_events[0].get("stage_elapsed_ms") is None
        ):
            raise AssertionError("unknown-error compatibility event is incomplete")
        fingerprints_after = source_fingerprints(
            [
                *sources,
                template,
                duplicate_marker_template,
                merged_comment_source,
                unsupported_error_source,
            ]
        )
        if fingerprints_before != fingerprints_after:
            raise AssertionError("a source workbook or template was modified")

        nocom_matches_rust = False
        nocom_difference_fields = ["not_saved"]
        if nocom_used is not None and rust_used is not None:
            nocom_difference_fields = projection_difference_fields(nocom_used, rust_used)
            nocom_matches_rust = not nocom_difference_fields

        if used_report.get("rows") != 19 or used_report.get("files") != 2:
            raise AssertionError(f"unexpected used-range summary report: {used_report}")
        if used_report.get("execution", {}).get("ooxml_error_cells_patched") != 7:
            raise AssertionError(f"expected seven preserved Excel error values: {used_report}")
        report_day = dt.date.today().strftime("%Y%m%d")
        report_path = (
            Path(__file__).resolve().parent
            / "reports"
            / f"summary_used_range_parity_{report_day}.json"
        )
        previous_report = (
            json.loads(report_path.read_text(encoding="utf-8"))
            if report_path.is_file()
            else {}
        )
        previous_unknown_probe = previous_report.get("unknown_cell_error_probe", {})
        previous_unknown_rust = previous_unknown_probe.get("rust_cli", {})
        cross_platform_runs = {
            key: previous_unknown_rust[key]
            for key in ("windows_gnu_debug", "linux_glibc_2_28_release")
            if key in previous_unknown_rust
        }
        if not cross_platform_runs:
            previous_cross_platform_runs = previous_unknown_probe.get(
                "cross_platform_runs_after_shared_reader_api", {}
            )
            cross_platform_runs = {
                key: previous_cross_platform_runs[key]
                for key in ("windows_gnu_debug", "linux_glibc_2_28_release")
                if key in previous_cross_platform_runs
            }
        previous_rust_behavior = previous_unknown_probe.get(
            "rust_cli_before_reader_compatibility_api",
            previous_unknown_probe.get("previous_rust_behavior_before_shared_reader_api"),
        )
        report = {
            "schema_version": 1,
            "suite": "pytools.summary_used_range_python_rust_parity",
            "date": dt.date.today().isoformat(),
            "status": "passed_synthetic_cli_parity",
            "source_python_file": "PyCode/core/summary_321_322.py::run_summary_by_usedrange",
            "real_sample": False,
            "inputs": {
                "files": [path.name for path in sources],
                "sha256_before_after_equal": True,
                "source_workbooks": 2,
            },
            "scenario": {
                "rows": used_report["rows"],
                "features": [
                    "actual Rust summary-used-range CLI",
                    "multiple workbooks and worksheets in selection order",
                    "macro-free XLSM source with macro-enabled workbook content type",
                    "empty worksheet produces the default-bound row seen in PyCode",
                    "hidden and veryHidden worksheets are included in workbook order",
                    "offset used range retains source row numbers and blank leading columns",
                    "format-only used-range tail",
                    "merged range used as an extent",
                    "cached formula values and missing formula-cache semantics",
                    "date, boolean, numeric, text and blank semantics",
                    "seven common Excel error values preserved as error-typed cells",
                ],
            },
            "comparison": {
                "cell_values_and_types": "exact",
                "cell_styles": "exact for populated output cells",
                "worksheet_names_and_dimensions": "exact",
                "freeze_panes": "exact",
                "header_bold_flags": "exact",
                "formulas_merges_defined_names_print_area_page_setup": "exact (default portrait normalized)",
            },
            "rust_execution": used_report.get("execution"),
            "progress_logging": used_report.get("progress_log"),
            "input_filter_probe": extension_filter_report,
            "vba_enabled_xlsm_probe": vba_enabled_xlsm_report,
            "offset_used_range_probe": {
                "source_sheet": "Offset",
                "source_bounds": "D5:F6",
                "projected_cells": offset_values,
                "comparison": "exact with PyCode",
            },
            "unknown_cell_error_probe": {
                "input": "synthetic XLSX t=e cell with value #DATA!",
                "pycode": {
                    "saved": True,
                    "output_value": "#DATA!",
                    "output_type": "s",
                    "rows": 1,
                },
                "rust_cli": {
                    "runtime": rust_cli_runtime_metadata()[0],
                    "configured_executable": rust_cli_runtime_metadata()[1],
                    "status": "succeeded",
                    "entrypoint": "pytools summary-used-range",
                    "reader_api": "read_workbook_with_unknown_error_tokens_as_text",
                    "reader_engine": rust_unknown_execution["reader_engine"],
                    "writer_engine": rust_unknown_execution["writer_engine"],
                    "normalized_cell_count": rust_unknown_execution[
                        "unknown_error_token_cells_normalized"
                    ],
                    "output_value": rust_unknown_cell["values"]["D2"],
                    "output_type": rust_unknown_cell["types"]["D2"],
                    "rows": rust_unknown_report["rows"],
                    "progress_log": rust_unknown_report["progress_log"],
                    "developer_compatibility_event_count": len(compatibility_events),
                    "developer_history_omits_cell_token": True,
                },
                "business_parity": "exact",
                "cross_platform_runs_after_shared_reader_api": cross_platform_runs,
                "note": (
                    "PT-6-1 opts into the shared reader's narrow literal error-token normalization API. "
                    "Ordinary read_workbook remains fail-closed; this probe does not infer support for other tokens."
                ),
            },
            "comment_summary_probe": {
                "baseline_rows": comments_report.get("rows"),
                "merged_non_anchor_cell": {
                    "fixture": "tests/fixtures/summary_comments/merged_children_source.xlsx",
                    "scenario": (
                        "B2 contains a synthetic OOXML value inside merge A2:C3; "
                        "PyCode _gcv resolves B2 and C2 to A2"
                    ),
                    "pycode_values": {
                        address: "merged anchor" for address in ("D2", "E2", "D3", "E3")
                    },
                    "rust_values": merged_values,
                    "rust_rows": merged_comments_report.get("rows"),
                    "rust_execution": merged_comments_report.get("execution"),
                    "comparison": "exact",
                    "rust_runtime": {
                        "host": rust_cli_runtime_metadata()[0],
                        "executable": rust_cli_runtime_metadata()[1],
                    },
                    "python_runtime": (
                        f"{sys.platform} Python {sys.version_info.major}.{sys.version_info.minor}"
                    ),
                    "source_unchanged": True,
                },
                "duplicate_row_region_marker": {
                    "scenario": "duplicate 行区域1 starts at A2 and A3; duplicate 行区域#1 ends at A4 and A5",
                    "python": {
                        "saved": True,
                        "rows": 6,
                        "source_row_numbers": [3, 4, 5, 3, 4, 5],
                    },
                    "rust_cli": {
                        "saved": duplicate_comments_report.get("saved"),
                        "rows": duplicate_comments_report.get("rows"),
                        "source_row_numbers": duplicate_source_rows,
                    },
                    "comparison": "exact",
                    "rule": "last matching start/end marker in row-major order, matching PyCode overwrite behavior",
                },
            },
            "python_variant_scope": {
                "nocom_pytools_output_matches_rust_for_this_xlsx_sample": nocom_matches_rust,
                "nocom_pytools_difference_fields": nocom_difference_fields,
                "nocom_pytools_vs_pycode": (
                    projection_difference_summary(python_used, nocom_used)
                    if python_used is not None and nocom_used is not None
                    else {"status": "not_saved"}
                ),
                "format_difference": "This probe covers xlsx, macro-free xlsm and the pinned VBA-enabled XlsxWriter test fixture. Original PyTools VBA-bearing xlsm and NoComPyTools-only xls/xlsb remain outside the validated set.",
            },
            "limitations": [
                "This report covers synthetic XLSX/XLSM plus one pinned VBA-enabled XlsxWriter test fixture; real-project XLSX parity is recorded separately in summary_used_range_real_parity_<date>.json.",
                "The official VBA-enabled fixture passed, but original PyTools VBA-bearing XLSM, XLS, XLSB, Kylin/UOS, glibc 2.28 and deb packaging remain unvalidated.",
                "NoComPyTools and PyCode input-format scopes differ; no production scope was inferred from the NoCom-only formats.",
                "Output is a new data workbook; source workbook styles, formulas and unrelated OOXML parts are not copied by contract.",
                "A synthetic XLSX t=e #DATA! cell is written as a string by PyCode; PT-6-1 now uses the shared reader's narrow literal-token normalization API and matches this case. No broader token or format support is inferred.",
            ],
        }
        if previous_rust_behavior is not None:
            for history_key in (
                "rust_cli_before_reader_compatibility_api",
                "previous_rust_behavior_before_shared_reader_api",
            ):
                report["unknown_cell_error_probe"][history_key] = previous_rust_behavior
        for key, value in previous_report.items():
            if key not in report:
                report[key] = value
        report_path.parent.mkdir(parents=True, exist_ok=True)
        report_path.write_text(json.dumps(report, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
        real_sample_report = compare_real_used_range_sample()
        print(json.dumps({
            "status": report["status"],
            "used_range_rows_rust": used_report.get("rows"),
            "used_range_files_rust": used_report.get("files"),
            "error_cells_patched": (used_report.get("execution") or {}).get("ooxml_error_cells_patched"),
            "nocom_pytools_matches_rust": nocom_matches_rust,
            "comment_summary_rows_rust": comments_report.get("rows"),
            "comment_duplicate_marker_rows_rust": duplicate_comments_report.get("rows"),
            "real_sample_summary": real_sample_report,
            "report": str(report_path),
        }, ensure_ascii=False))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
