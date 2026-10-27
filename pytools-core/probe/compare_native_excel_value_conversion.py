from __future__ import annotations

import json
import math
import platform
import shutil
import subprocess
import tempfile
from datetime import date
from datetime import datetime, time, timedelta
from pathlib import Path
import zipfile

import openpyxl
import xlrd
from openpyxl import Workbook, load_workbook
from openpyxl.utils.datetime import MAC_EPOCH, to_excel

REPO = Path(__file__).resolve().parents[2]
FIXTURE = REPO / "pytools-core" / "tests" / "fixtures" / "timeline" / "wide_summary_legacy.xls"
RUST_CLI = REPO / "target" / "debug" / "pytools.exe"
REPORT = (
    REPO
    / "pytools-core"
    / "probe"
    / "reports"
    / "native_excel_value_conversion_parity_20261026.json"
)


def make_config(path: Path, root: Path) -> None:
    workbook = Workbook()
    global_sheet = workbook.active
    global_sheet.title = "全局配置"
    global_sheet.append(["配置项", "值"])
    for key, value in (
        ("输入目录", str(root)),
        ("输出目录", str(root / "output")),
        ("日志目录", str(root / "logs")),
        ("excel目的格式", "xlsx"),
        ("word目的格式", "docx"),
    ):
        global_sheet.append([key, value])
    workbook.save(path)


def compare_values(source_path: Path, output_path: Path) -> dict[str, object]:
    source = xlrd.open_workbook(source_path, on_demand=True)
    output = load_workbook(output_path, data_only=False, read_only=False)
    try:
        source_names = source.sheet_names()
        if source_names != output.sheetnames:
            raise AssertionError(
                f"worksheet names/order differ: Python={source_names}, Rust={output.sheetnames}"
            )

        nonempty_cells = 0
        numeric_cells = 0
        text_cells = 0
        for source_sheet, output_sheet in zip(source.sheets(), output.worksheets):
            expected: dict[str, tuple[int, object]] = {}
            for row_index in range(source_sheet.nrows):
                for column_index in range(source_sheet.ncols):
                    cell = source_sheet.cell(row_index, column_index)
                    if cell.ctype in (xlrd.XL_CELL_EMPTY, xlrd.XL_CELL_BLANK):
                        continue
                    coordinate = f"{xlrd.colname(column_index)}{row_index + 1}"
                    expected[coordinate] = (cell.ctype, cell.value)

            actual = {
                cell.coordinate: cell
                for row in output_sheet.iter_rows()
                for cell in row
                if cell.value is not None
            }
            if set(expected) != set(actual):
                raise AssertionError(
                    f"non-empty cell coordinates differ on {source_sheet.name}: "
                    f"Python={sorted(expected)}, Rust={sorted(actual)}"
                )

            for coordinate, (source_type, source_value) in expected.items():
                target_cell = actual[coordinate]
                if source_type == xlrd.XL_CELL_TEXT:
                    if not isinstance(target_cell.value, str) or target_cell.value != source_value:
                        raise AssertionError(
                            f"text differs at {source_sheet.name}!{coordinate}: "
                            f"Python={source_value!r}, Rust={target_cell.value!r}"
                        )
                    text_cells += 1
                elif source_type == xlrd.XL_CELL_NUMBER:
                    if isinstance(target_cell.value, bool) or not isinstance(
                        target_cell.value, (int, float)
                    ):
                        raise AssertionError(
                            f"numeric type differs at {source_sheet.name}!{coordinate}: "
                            f"Rust={type(target_cell.value).__name__}"
                        )
                    if not math.isclose(
                        float(source_value), float(target_cell.value), rel_tol=0.0, abs_tol=0.0
                    ):
                        raise AssertionError(
                            f"numeric value differs at {source_sheet.name}!{coordinate}: "
                            f"Python={source_value!r}, Rust={target_cell.value!r}"
                        )
                    numeric_cells += 1
                else:
                    raise AssertionError(
                        f"fixture requires an explicit comparison for xlrd cell type "
                        f"{source_type} at {source_sheet.name}!{coordinate}"
                    )
                nonempty_cells += 1

        return {
            "scope": "Python xlrd source values compared with the Rust-generated XLSX reopened by openpyxl; this does not compare the legacy PyCode fidelity-conversion result.",
            "worksheet_names_and_order_equal": True,
            "nonempty_cell_coordinates_equal": True,
            "text_values_equal": True,
            "numeric_values_equal": True,
            "worksheets": len(source_names),
            "nonempty_cells": nonempty_cells,
            "text_cells": text_cells,
            "numeric_cells": numeric_cells,
            "numeric_type_projection": "XLS integer storage is read by xlrd as NUMBER and emitted by Rust as an XLSX numeric cell; values compare exactly.",
        }
    finally:
        source.release_resources()
        output.close()


def main() -> int:
    for required in (FIXTURE, RUST_CLI):
        if not required.is_file():
            raise FileNotFoundError(required)

    source_before = FIXTURE.read_bytes()
    with tempfile.TemporaryDirectory(prefix="pytools-native-xls-parity-") as temp_name:
        output_dir = Path(temp_name)
        config = output_dir / "config.xlsx"
        make_config(config, output_dir)
        result = subprocess.run(
            [
                str(RUST_CLI),
                "convert-excel",
                "--config",
                str(config),
                "--to",
                "xlsx",
                "--engine",
                "native",
                "--output-dir",
                str(output_dir),
                "--log-dir",
                str(output_dir / "logs"),
                "--log-view",
                "developer",
                str(FIXTURE),
            ],
            check=False,
            capture_output=True,
        )
        stdout = result.stdout.decode("utf-8-sig", errors="replace").replace("\x00", "")
        stderr = result.stderr.decode("utf-8-sig", errors="replace").replace("\x00", "")
        if result.returncode != 0:
            raise RuntimeError(f"native conversion failed ({result.returncode}): {stderr[-2000:]}")
        rust_report = json.loads(stdout)
        if rust_report.get("converted") != 1 or rust_report.get("failed") != 0:
            raise AssertionError(f"unexpected Rust conversion result: {rust_report}")
        history_path = Path(rust_report["run_history_path"])
        history = json.loads(history_path.read_text(encoding="utf-8"))
        conversion_event = next(
            (
                event
                for event in history.get("events", [])
                if event.get("phase") == "convert_file" and event.get("status") == "completed"
            ),
            None,
        )
        if not conversion_event or conversion_event.get("actual_engine") != "calamine + rust_xlsxwriter":
            raise AssertionError(f"unexpected conversion engine history: {history}")
        history_text = history_path.read_text(encoding="utf-8")
        if FIXTURE.name.lower() in history_text.lower():
            raise AssertionError("developer run history contains the input file name")
        output_path = Path(rust_report["items"][0]["output"])
        comparison = compare_values(FIXTURE, output_path)
        source_unchanged = FIXTURE.read_bytes() == source_before
        if not source_unchanged:
            raise AssertionError("the synthetic XLS source was modified")

        iso_source = output_dir / "iso-date.xlsm"
        iso_workbook = Workbook()
        iso_workbook.iso_dates = True
        iso_sheet = iso_workbook.active
        iso_sheet.title = "日期"
        iso_sheet["A1"] = datetime(2024, 2, 29, 12, 30)
        iso_workbook.save(iso_source)
        iso_source_before = iso_source.read_bytes()
        iso_result = subprocess.run(
            [
                str(RUST_CLI),
                "convert-excel",
                "--config",
                str(config),
                "--to",
                "xlsx",
                "--engine",
                "native",
                "--output-dir",
                str(output_dir),
                "--log-dir",
                str(output_dir / "logs"),
                "--log-view",
                "developer",
                str(iso_source),
            ],
            check=False,
            capture_output=True,
        )
        iso_stdout = iso_result.stdout.decode("utf-8-sig", errors="replace").replace("\x00", "")
        if iso_result.returncode != 2:
            raise AssertionError(
                f"ISO-typed date input must fail closed; exit={iso_result.returncode}, "
                f"stdout={iso_stdout[-1000:]!r}"
            )
        iso_report = json.loads(iso_stdout)
        if iso_report.get("converted") != 0 or iso_report.get("failed") != 1:
            raise AssertionError(f"unexpected ISO date refusal result: {iso_report}")
        iso_output = output_dir / "iso-date.xlsx"
        if iso_output.exists():
            raise AssertionError("ISO-typed date refusal published an output workbook")
        if iso_source.read_bytes() != iso_source_before:
            raise AssertionError("the ISO date source workbook was modified")
        iso_failure_detail = str(iso_report["items"][0].get("detail", ""))
        if "ISO-typed date/time" not in iso_failure_detail or "cell type" not in iso_failure_detail:
            raise AssertionError(f"ISO date refusal detail is incomplete: {iso_failure_detail!r}")

        error_source = output_dir / "error-cell.xlsm"
        error_workbook = Workbook()
        error_sheet = error_workbook.active
        error_sheet.title = "错误"
        error_sheet["A1"] = "#N/A"
        error_sheet["A1"].data_type = "e"
        error_workbook.save(error_source)
        error_source_before = error_source.read_bytes()
        error_result = subprocess.run(
            [
                str(RUST_CLI),
                "convert-excel",
                "--config",
                str(config),
                "--to",
                "xlsx",
                "--engine",
                "native",
                "--output-dir",
                str(output_dir),
                "--log-dir",
                str(output_dir / "logs"),
                "--log-view",
                "developer",
                str(error_source),
            ],
            check=False,
            capture_output=True,
        )
        error_stdout = error_result.stdout.decode("utf-8-sig", errors="replace").replace(
            "\x00", ""
        )
        if error_result.returncode != 2:
            raise AssertionError(
                f"typed error cell must fail closed; exit={error_result.returncode}, "
                f"stdout={error_stdout[-1000:]!r}"
            )
        error_report = json.loads(error_stdout)
        if error_report.get("converted") != 0 or error_report.get("failed") != 1:
            raise AssertionError(f"unexpected typed error refusal result: {error_report}")
        error_output = output_dir / "error-cell.xlsx"
        if error_output.exists():
            raise AssertionError("typed error refusal published an output workbook")
        if error_source.read_bytes() != error_source_before:
            raise AssertionError("the typed error source workbook was modified")
        error_failure_detail = str(error_report["items"][0].get("detail", ""))
        if "error cell" not in error_failure_detail or "错误!A1" not in error_failure_detail:
            raise AssertionError(f"typed error refusal detail is incomplete: {error_failure_detail!r}")

        serial_source = output_dir / "serial-date-time.xlsm"
        serial_workbook = Workbook()
        serial_sheet = serial_workbook.active
        serial_sheet.title = "日期时间"
        serial_sheet["A1"] = datetime(2024, 2, 29, 12, 30)
        serial_sheet["A1"].number_format = "yyyy-mm-dd hh:mm:ss"
        serial_sheet["A2"] = time(12, 30)
        serial_sheet["A2"].number_format = "hh:mm"
        serial_sheet["A3"] = timedelta(hours=26, minutes=15)
        serial_sheet["A3"].number_format = "[h]:mm:ss"
        serial_sheet["B1"] = True
        serial_sheet["B2"] = False
        serial_sheet["C1"] = "=SUM(A1,A2)"
        serial_workbook.save(serial_source)
        serial_source_before = serial_source.read_bytes()
        serial_result = subprocess.run(
            [
                str(RUST_CLI),
                "convert-excel",
                "--config",
                str(config),
                "--to",
                "xlsx",
                "--engine",
                "native",
                "--output-dir",
                str(output_dir),
                "--log-dir",
                str(output_dir / "logs"),
                "--log-view",
                "developer",
                str(serial_source),
            ],
            check=False,
            capture_output=True,
        )
        serial_stdout = (
            serial_result.stdout.decode("utf-8-sig", errors="replace").replace("\x00", "")
        )
        if serial_result.returncode != 0:
            raise AssertionError(
                f"numeric serial date/time conversion failed: {serial_stdout[-1000:]!r}"
            )
        serial_report = json.loads(serial_stdout)
        if serial_report.get("converted") != 1 or serial_report.get("failed") != 0:
            raise AssertionError(f"unexpected numeric serial conversion result: {serial_report}")
        serial_output = load_workbook(serial_report["items"][0]["output"], data_only=False)
        try:
            serial_output_sheet = serial_output["日期时间"]
            serial_expected = {
                "A1": to_excel(datetime(2024, 2, 29, 12, 30)),
                "A2": to_excel(time(12, 30)),
                "A3": timedelta(hours=26, minutes=15).total_seconds() / 86_400,
            }
            serial_actual = {}
            for address, expected in serial_expected.items():
                actual = serial_output_sheet[address].value
                if isinstance(actual, bool) or not isinstance(actual, (int, float)):
                    raise AssertionError(
                        f"numeric serial cell type changed at {address}: {type(actual).__name__}"
                    )
                if not math.isclose(float(actual), expected, rel_tol=0.0, abs_tol=1e-10):
                    raise AssertionError(
                        f"numeric serial value differs at {address}: Python={expected!r}, "
                        f"Rust={actual!r}"
                    )
                serial_actual[address] = float(actual)
            for address, expected in {"B1": True, "B2": False}.items():
                actual = serial_output_sheet[address]
                if actual.data_type != "b" or actual.value is not expected:
                    raise AssertionError(
                        f"boolean cell type/value changed at {address}: "
                        f"Python={expected!r}, Rust={actual.value!r}/{actual.data_type!r}"
                    )
            formula = serial_output_sheet["C1"]
            if formula.data_type != "f" or formula.value != "=SUM(A1,A2)":
                raise AssertionError(
                    f"formula text/type changed at C1: {formula.value!r}/{formula.data_type!r}"
                )
        finally:
            serial_output.close()
        if serial_source.read_bytes() != serial_source_before:
            raise AssertionError("the numeric serial date/time source workbook was modified")

        epoch_source = output_dir / "epoch-1904.xlsm"
        epoch_workbook = Workbook()
        epoch_workbook.epoch = MAC_EPOCH
        epoch_sheet = epoch_workbook.active
        epoch_sheet.title = "日期"
        epoch_sheet["A1"] = datetime(2024, 2, 29)
        epoch_sheet["A1"].number_format = "yyyy-mm-dd"
        epoch_sheet["B1"] = "=DATE(2024,2,29)"
        epoch_workbook.save(epoch_source)
        epoch_source_before = epoch_source.read_bytes()
        epoch_result = subprocess.run(
            [
                str(RUST_CLI),
                "convert-excel",
                "--config",
                str(config),
                "--to",
                "xlsx",
                "--engine",
                "native",
                "--output-dir",
                str(output_dir),
                "--log-dir",
                str(output_dir / "logs"),
                "--log-view",
                "developer",
                str(epoch_source),
            ],
            check=False,
            capture_output=True,
        )
        epoch_stdout = (
            epoch_result.stdout.decode("utf-8-sig", errors="replace").replace("\x00", "")
        )
        if epoch_result.returncode != 2:
            raise AssertionError(
                f"1904 date-system XLSM must fail closed; exit={epoch_result.returncode}, "
                f"stdout={epoch_stdout[-1000:]!r}"
            )
        epoch_report = json.loads(epoch_stdout)
        if epoch_report.get("converted") != 0 or epoch_report.get("failed") != 1:
            raise AssertionError(f"unexpected 1904 date-system result: {epoch_report}")
        epoch_output = output_dir / "epoch-1904.xlsx"
        if epoch_output.exists():
            raise AssertionError("1904 date-system refusal published an output workbook")
        if epoch_source.read_bytes() != epoch_source_before:
            raise AssertionError("the 1904 date-system source workbook was modified")
        epoch_failure_detail = str(epoch_report["items"][0].get("detail", ""))
        if "1904 date system" not in epoch_failure_detail:
            raise AssertionError(f"1904 date-system refusal detail is incomplete: {epoch_failure_detail!r}")

        vba_source = REPO / "pytools-core" / "tests" / "fixtures" / "vba_config" / "VbaV2.4.2.xlsm"
        vba_source_before = vba_source.read_bytes()
        vba_result = subprocess.run(
            [
                str(RUST_CLI),
                "convert-excel",
                "--config",
                str(config),
                "--to",
                "xlsx",
                "--engine",
                "native",
                "--output-dir",
                str(output_dir),
                "--log-dir",
                str(output_dir / "logs"),
                "--log-view",
                "developer",
                str(vba_source),
            ],
            check=False,
            capture_output=True,
        )
        vba_stdout = vba_result.stdout.decode("utf-8-sig", errors="replace").replace("\x00", "")
        if vba_result.returncode != 0:
            raise AssertionError(
                f"native VBA-config fixture conversion failed: {vba_stdout[-1000:]!r}"
            )
        vba_report = json.loads(vba_stdout)
        if vba_report.get("converted") != 1 or vba_report.get("failed") != 0:
            raise AssertionError(f"unexpected VBA-config fixture result: {vba_report}")
        vba_history_path = Path(vba_report["run_history_path"])
        vba_history = json.loads(vba_history_path.read_text(encoding="utf-8"))
        vba_conversion_event = next(
            (
                event
                for event in vba_history.get("events", [])
                if event.get("phase") == "convert_file" and event.get("status") == "completed"
            ),
            None,
        )
        if (
            not vba_conversion_event
            or vba_conversion_event.get("actual_engine") != "calamine + rust_xlsxwriter"
        ):
            raise AssertionError(f"unexpected VBA-config conversion engine: {vba_history}")
        if vba_source.name.lower() in vba_history_path.read_text(encoding="utf-8").lower():
            raise AssertionError("VBA-config run history contains the source file name")
        vba_output_path = Path(vba_report["items"][0]["output"])
        vba_source_workbook = load_workbook(
            vba_source, data_only=False, read_only=False, keep_vba=True
        )
        vba_output_workbook = load_workbook(vba_output_path, data_only=False, read_only=False)
        vba_counts = {
            "worksheets": len(vba_source_workbook.sheetnames),
            "source_nonempty_cells": 0,
            "output_nonempty_cells": 0,
            "date_cells": 0,
            "formula_cells": 0,
            "boolean_cells": 0,
            "numeric_cells": 0,
            "text_cells": 0,
            "max_numeric_abs_diff": 0.0,
            "mismatches": 0,
            "unexpected_output_cells": 0,
        }
        if vba_source_workbook.sheetnames != vba_output_workbook.sheetnames:
            raise AssertionError("VBA-config conversion changed worksheet names or order")
        for source_sheet, output_sheet in zip(
            vba_source_workbook.worksheets, vba_output_workbook.worksheets
        ):
            max_row = max(source_sheet.max_row or 1, output_sheet.max_row or 1)
            max_column = max(source_sheet.max_column or 1, output_sheet.max_column or 1)
            source_rows = source_sheet.iter_rows(
                min_row=1, max_row=max_row, min_col=1, max_col=max_column
            )
            output_rows = output_sheet.iter_rows(
                min_row=1, max_row=max_row, min_col=1, max_col=max_column
            )
            for source_row, output_row in zip(source_rows, output_rows):
                for source_cell, output_cell in zip(source_row, output_row):
                    if output_cell.value is not None:
                        vba_counts["output_nonempty_cells"] += 1
                    value = source_cell.value
                    if value is None:
                        if output_cell.value is not None:
                            vba_counts["unexpected_output_cells"] += 1
                            vba_counts["mismatches"] += 1
                        continue
                    vba_counts["source_nonempty_cells"] += 1
                    if source_cell.data_type == "f":
                        vba_counts["formula_cells"] += 1
                        equal = output_cell.data_type == "f" and output_cell.value == value
                    elif source_cell.data_type == "e":
                        vba_counts["mismatches"] += 1
                        continue
                    elif source_cell.is_date:
                        vba_counts["date_cells"] += 1
                        expected = to_excel(value, epoch=vba_source_workbook.epoch)
                        actual = output_cell.value
                        equal = (
                            isinstance(actual, (int, float))
                            and not isinstance(actual, bool)
                            and math.isclose(float(actual), expected, rel_tol=0.0, abs_tol=1e-10)
                        )
                        if equal:
                            vba_counts["max_numeric_abs_diff"] = max(
                                vba_counts["max_numeric_abs_diff"], abs(float(actual) - expected)
                            )
                    elif isinstance(value, bool):
                        vba_counts["boolean_cells"] += 1
                        equal = output_cell.value is value
                    elif isinstance(value, (int, float)):
                        vba_counts["numeric_cells"] += 1
                        equal = (
                            isinstance(output_cell.value, (int, float))
                            and not isinstance(output_cell.value, bool)
                            and math.isclose(
                                float(output_cell.value), float(value), rel_tol=0.0, abs_tol=1e-10
                            )
                        )
                        if equal:
                            vba_counts["max_numeric_abs_diff"] = max(
                                vba_counts["max_numeric_abs_diff"],
                                abs(float(output_cell.value) - float(value)),
                            )
                    else:
                        vba_counts["text_cells"] += 1
                        equal = isinstance(output_cell.value, str) and output_cell.value == value
                    if not equal:
                        vba_counts["mismatches"] += 1
        vba_source_workbook.close()
        vba_output_workbook.close()
        with zipfile.ZipFile(vba_source) as source_archive, zipfile.ZipFile(
            vba_output_path
        ) as output_archive:
            source_has_vba = "xl/vbaProject.bin" in source_archive.namelist()
            output_has_vba = "xl/vbaProject.bin" in output_archive.namelist()
        if not source_has_vba or output_has_vba:
            raise AssertionError("native .xlsm conversion did not remove VBA as documented")
        if vba_source.read_bytes() != vba_source_before:
            raise AssertionError("the repository VBA-config fixture was modified")
        if (
            vba_counts["source_nonempty_cells"] != vba_counts["output_nonempty_cells"]
            or vba_counts["mismatches"] != 0
        ):
            raise AssertionError(f"VBA-config fixture projection differs: {vba_counts}")

    report = {
        "schema_version": 1,
        "suite": "pytools.conversion.native_excel_value_only",
        "date": date.today().isoformat(),
        "status": "passed_synthetic_and_repository_fixture_python_rust_value_projection",
        "environment": {
            "os": platform.platform(),
            "rust_entry": "Windows-native pytools convert-excel --engine native",
            "python_baseline": f"xlrd {xlrd.__version__} source reader + openpyxl {openpyxl.__version__} output reader",
        },
        "input": {
            "format": "xls",
            "kind": "synthetic smoke workbook",
            "fixture": "pytools-core/tests/fixtures/timeline/wide_summary_legacy.xls",
        },
        "output": {
            "format": "xlsx",
            "requested_engine": "native",
            "actual_engine": conversion_event["actual_engine"],
            "actual_components": conversion_event.get("components"),
            "run_history_redacted": True,
            "source_unchanged": source_unchanged,
        },
        "comparison": comparison,
        "type_guard": {
            "input": "synthetic .xlsm containing an ISO-typed Excel date generated with openpyxl iso_dates=True",
            "rust_result": "conversion refused before publishing because rust_xlsxwriter output would change the cell type to text",
            "converted": iso_report["converted"],
            "failed": iso_report["failed"],
            "output_published": False,
            "source_unchanged": True,
        },
        "error_value_guard": {
            "input": "synthetic .xlsm containing a typed #N/A error cell at 错误!A1",
            "rust_result": "conversion refused before publishing because the writer cannot preserve a typed Excel error",
            "converted": error_report["converted"],
            "failed": error_report["failed"],
            "error_location_reported": "错误!A1" in error_failure_detail,
            "output_published": False,
            "source_unchanged": True,
        },
        "numeric_serial_values": {
            "date_system": "1900",
            "cells": {
                address: {
                    "python_expected_serial": expected,
                    "rust_output_number": serial_actual[address],
                    "output_type": "number",
                }
                for address, expected in serial_expected.items()
            },
            "boolean_cells": {
                "B1": {"python_value": True, "rust_value": True, "output_type": "boolean"},
                "B2": {"python_value": False, "rust_value": False, "output_type": "boolean"},
            },
            "formula_cells": {
                "C1": {
                    "python_formula_text": "=SUM(A1,A2)",
                    "rust_formula_text": "=SUM(A1,A2)",
                    "output_type": "formula",
                    "cache_parity": "not tested; native value-only conversion does not preserve formula caches",
                }
            },
            "formatting_dropped": True,
            "source_unchanged": True,
        },
        "date_epoch_guard": {
            "input": "synthetic XLSM with 1904 epoch, typed date serial 43889 and formula text",
            "rust_result": "conversion refused because the new-workbook writer cannot preserve the source epoch",
            "converted": epoch_report["converted"],
            "failed": epoch_report["failed"],
            "output_published": False,
            "source_unchanged": True,
        },
        "repository_vba_fixture_projection": {
            "fixture": "pytools-core/tests/fixtures/vba_config/VbaV2.4.2.xlsm",
            "kind": "repository VBA config fixture; not a real user business template",
            "actual_engine": vba_conversion_event["actual_engine"],
            "actual_components": vba_conversion_event.get("components"),
            "source_has_vba": source_has_vba,
            "output_has_vba": output_has_vba,
            "formatting_dropped": True,
            "source_unchanged": True,
            **vba_counts,
        },
        "unverified": [
            "Real business XLS workbooks and configurations.",
            "Formula cache parity against legacy XLS conversion, 1904 epoch detection for legacy .xls/.xlsb inputs, and empty-string cells.",
            "Styles, comments, merged regions, dimensions, names, print settings, validations, conditional formatting, drawings and macros are intentionally not preserved by this value-only route.",
            "Windows desktop file dialog, glibc 2.28, Kylin V10/UOS V20 and offline deb delivery.",
        ],
    }
    REPORT.parent.mkdir(parents=True, exist_ok=True)
    REPORT.write_text(json.dumps(report, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    print(
        json.dumps(
            {"status": report["status"], "report": str(REPORT), "comparison": comparison},
            ensure_ascii=False,
        )
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
