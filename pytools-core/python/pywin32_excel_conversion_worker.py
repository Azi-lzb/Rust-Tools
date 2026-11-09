#!/usr/bin/env python
"""Batch-scoped Excel/WPS COM worker used by the Windows PyTools package.

The Rust process owns file discovery, output paths, logging, and final publish.
This worker drives Excel/WPS through pywin32 for conversion and disposable-copy
validation recalculation. Conversion candidates are verified before Rust accepts
them. Requests and responses use one UTF-8 JSON object per line.
"""

from __future__ import annotations

import importlib.metadata
import json
import math
import os
import re
import shutil
import sys
import tempfile
import traceback
from pathlib import Path
from typing import Any

PROTOCOL_VERSION = 3
XLSX_FILE_FORMAT = 51  # xlOpenXMLWorkbook
# wdFormatXMLDocument. `wdFormatDocumentDefault` (16) is accepted by Microsoft
# Word but WPS Writer ignores it and writes a binary .doc under a .docx name,
# so use the classic OOXML constant and verify the saved package afterwards.
DOCX_FILE_FORMAT = 12
# WPS-native spreadsheet containers. `.et` / `.ett` are CFB files that Excel
# would happily open as a renamed XLS, so the native FileFormat is verified
# before anything is published. WPS reports its own formats in a private range
# starting at 65521 (ET workbook); any value below that is an Excel container.
WPS_NATIVE_FILE_FORMAT_MIN = 65521
# BIFF workbook formats an `.et`/`.ett` file may legitimately carry: xlExcel8
# (56) and the older BIFF5/95 variants. Content decides the route.
BIFF_FILE_FORMATS = {56, 8, 16}  # xlExcel8, xlExcel4/5, xlExcel2/3/95
ET_FILE_FORMAT = 65521
SPREADSHEET_SUFFIXES = (".xls", ".xlsb", ".xlsm", ".et", ".ett")
WORD_SUFFIXES = (".doc", ".wps", ".wpt")
WPS_WORD_SUFFIXES = (".wps", ".wpt")
WPS_SPREADSHEET_SUFFIXES = (".et", ".ett")


def _write_json(value: dict[str, Any]) -> None:
    sys.stdout.write(json.dumps(value, ensure_ascii=False, separators=(",", ":")) + "\n")
    sys.stdout.flush()


def _absolute(value: str) -> Path:
    """Office/WPS COM resolves relative paths against its own working
    directory, so every path crossing the COM boundary must be absolute."""
    return Path(value).expanduser().resolve()


# Containers that are ZIP packages. A truncated/damaged package makes Excel
# fail `Workbooks.Open` with the opaque "类 Workbooks 的 Open 方法无效"
# (error 1004), so we check the container locally first and report the real
# reason instead of an Excel-internal generic message.
_ZIP_SPREADSHEET_SUFFIXES = (".xlsx", ".xlsm", ".xlsb", ".xltx", ".xltm")
_ZIP_WORD_SUFFIXES = (".docx", ".docm", ".dotx", ".dotm")


def _validate_container(path: Path, label: str) -> None:
    """Reject an unreadable package before Excel/WPS sees it."""
    import zipfile

    if not path.is_file():
        raise FileNotFoundError(f"{label}不存在：{path}")
    size = path.stat().st_size
    if size == 0:
        raise ValueError(f"{label}是空文件（0 字节）：{path}")
    suffix = path.suffix.lower()
    if suffix not in _ZIP_SPREADSHEET_SUFFIXES + _ZIP_WORD_SUFFIXES:
        return
    with path.open("rb") as handle:
        if handle.read(4) != b"PK\x03\x04":
            # 4-4 also normalizes legacy Excel files that were given an
            # .xlsx suffix without being saved in OOXML format. Let Excel/WPS
            # identify the real workbook type, then rely on the COM snapshot
            # comparison after SaveAs before Rust publishes the candidate.
            if suffix == ".xlsx":
                return
            raise ValueError(
                f"{label}不是 OOXML/ZIP 包（扩展名 {suffix}，文件头不是 PK）；"
                f"文件可能已损坏或只是被改名：{path}"
            )
    try:
        with zipfile.ZipFile(path) as archive:
            broken = archive.testzip()
            if broken is not None:
                raise ValueError(f"压缩包成员 {broken!r} 校验失败")
            names = set(archive.namelist())
    except zipfile.BadZipFile as error:
        raise ValueError(
            f"{label}的 OOXML 包已损坏（{error}）；文件大小 {size} 字节：{path}"
        ) from error
    # A ZIP can be intact yet still unusable: Excel then fails the same opaque
    # way. Check the parts every OOXML Office package must carry.
    required = ["[Content_Types].xml", "_rels/.rels"]
    if suffix in _ZIP_WORD_SUFFIXES:
        required.append("word/document.xml")
    elif suffix == ".xlsb":
        # Binary workbook: the main part is `xl/workbook.bin`, not .xml.
        required.append("xl/workbook.bin")
    else:
        required.append("xl/workbook.xml")
    missing = [name for name in required if name not in names]
    if missing:
        raise ValueError(
            f"{label}缺少 OOXML 必需部件 {missing}；文件可能未保存完整或被裁剪：{path}"
        )


def _force_recalc(application: Any) -> tuple[str | None, list[str]]:
    """Force every formula to be recalculated before SaveAs.

    `Calculate` is deliberately absent: measured on a freshly opened workbook
    whose cached values were stale (999 instead of 34), `Application.Calculate`
    left the cache untouched because nothing was marked dirty. It reports no
    error, so treating it as success would publish stale numbers that
    Calamine-style readers consume. `CalculateFull` still forces a full
    calculation of all formulas; only the dependency-tree rebuild is skipped.

    Returns `(method_name, errors)`; `method_name` is None when every method
    failed. These are Application-level methods - Workbook exposes none of them.
    """
    errors: list[str] = []
    for name in ("CalculateFullRebuild", "CalculateFull"):
        method = getattr(application, name, None)
        if method is None:
            errors.append(f"{name}: 该 COM 宿主不提供此方法")
            continue
        try:
            method()
            return name, errors
        except Exception as error:
            errors.append(f"{name}: {type(error).__name__}: {error}")
    return None, errors


def _process_is_alive(process_id: int) -> bool:
    """Whether a PID is still running, without requiring extra privileges."""
    try:
        import win32api
        import win32con
        import win32process

        handle = win32api.OpenProcess(win32con.PROCESS_QUERY_LIMITED_INFORMATION, False, process_id)
        try:
            return win32process.GetExitCodeProcess(handle) == win32con.STILL_ACTIVE
        finally:
            win32api.CloseHandle(handle)
    except Exception:
        return False


def _matrix(value: Any, rows: int, columns: int) -> list[list[Any]]:
    """Normalize a Range.Formula/Value2 payload into rows x columns.

    pywin32 is inconsistent here: a multi-cell range comes back as a tuple of
    row-tuples even when it has a single column, so `(('base',), ('10',))`
    must become `[['base'], ['10']]` instead of a list of 1-tuples.
    """
    if rows == 1 and columns == 1:
        return [[value]]
    if rows == 1:
        if isinstance(value, tuple) and len(value) == 1 and isinstance(value[0], tuple):
            value = value[0]
        return [list(value)]
    if columns == 1:
        return [list(item) if isinstance(item, tuple) else [item] for item in value]
    return [list(row) if isinstance(row, tuple) else [row] for row in value]


# xlCellTypeFormulas — SpecialCells selector for formula cells.
_XL_CELL_TYPE_FORMULAS = -4123


def _formula_map(worksheet: Any, used: Any) -> dict[tuple[int, int], str]:
    """Formula text per cell, keyed by absolute (row, column).

    Reading `UsedRange.Formula` in one go fails on some workbooks (observed on a
    154x40 sheet of a 24-sheet legacy .xls: `Value2` reads fine while
    `.Formula` raises com_error). Fall back to asking Excel for the formula
    cells only, then read each block individually.
    """
    try:
        matrix = _matrix(used.Formula, int(used.Rows.Count), int(used.Columns.Count))
        first_row = int(used.Row)
        first_column = int(used.Column)
        result: dict[tuple[int, int], str] = {}
        for row_offset, row in enumerate(matrix):
            for column_offset, formula in enumerate(row):
                if isinstance(formula, str):
                    result[(first_row + row_offset, first_column + column_offset)] = formula
        return result
    except Exception:
        pass
    # Fallback: enumerate formula cells through SpecialCells, one area at a time.
    result = {}
    try:
        areas = worksheet.Cells.SpecialCells(_XL_CELL_TYPE_FORMULAS)
    except Exception:
        return result
    for area_index in range(1, int(areas.Areas.Count) + 1):
        area = areas.Areas(area_index)
        rows = int(area.Rows.Count)
        columns = int(area.Columns.Count)
        first_row = int(area.Row)
        first_column = int(area.Column)
        block = _matrix(area.Formula, rows, columns)
        for row_offset, row in enumerate(block):
            for column_offset, formula in enumerate(row):
                if isinstance(formula, str):
                    result[(first_row + row_offset, first_column + column_offset)] = formula
    return result


def _snapshot(workbook: Any) -> list[dict[str, Any]]:
    sheets: list[dict[str, Any]] = []
    for index in range(1, int(workbook.Worksheets.Count) + 1):
        worksheet = workbook.Worksheets(index)
        used = worksheet.UsedRange
        rows = int(used.Rows.Count)
        columns = int(used.Columns.Count)
        first_row = int(used.Row)
        first_column = int(used.Column)
        # `Value2` tolerates regions where the bulk `.Formula` read fails.
        values = _matrix(used.Value2, rows, columns)
        formula_by_cell = _formula_map(worksheet, used)
        formula_cells: dict[tuple[int, int], str] = {}
        formula_values: dict[tuple[int, int], Any] = {}
        value_cells: dict[tuple[int, int], Any] = {}

        for row_offset in range(rows):
            for column_offset in range(columns):
                address = (first_row + row_offset, first_column + column_offset)
                value = values[row_offset][column_offset]
                formula = formula_by_cell.get(address)
                if formula is not None and formula.startswith("="):
                    formula_cells[address] = formula
                    formula_values[address] = value
                elif value is not None:
                    value_cells[address] = value

        snapshot = {
            "name": str(worksheet.Name),
            "formulas": formula_cells,
            "formula_values": formula_values,
            "values": value_cells,
            "used_cells": rows * columns,
        }
        sheets.append(snapshot)
    return sheets


def _column_name(column: int) -> str:
    result = ""
    while column:
        column, remainder = divmod(column - 1, 26)
        result = chr(65 + remainder) + result
    return result


def _cell_name(address: tuple[int, int]) -> str:
    return f"{_column_name(address[1])}{address[0]}"


def _values_equal(left: Any, right: Any) -> bool:
    if isinstance(left, bool) or isinstance(right, bool):
        return type(left) is type(right) and left == right
    if isinstance(left, (int, float)) and isinstance(right, (int, float)):
        if left == right:
            return True
        return (
            math.isfinite(left)
            and math.isfinite(right)
            and left != 0
            and right != 0
            and abs(left - right) <= 1e-15 * max(abs(left), abs(right))
        )
    return type(left) is type(right) and left == right


_ROW_CEILING_BIFF = 65_536
_ROW_CEILING_OOXML = 1_048_576


def _canonicalize_formula(formula: str) -> str:
    """Normalize the formula rewrites a SaveAs may legitimately perform.

    Mirrors Rust `excel_reader::canonicalize_formula_for_recalculation`:

    * BIFF caps a sheet at 65,536 rows, OOXML at 1,048,576; Excel's SaveAs
      rewrites a whole-column span (`$1:$65536`, `$D$1:$D$65536`) to the OOXML
      ceiling. Both name the same column.
    * `TRUE()`/`FALSE()` versus `TRUE`/`FALSE`.
    * Optional quotes around known internal sheet names.
    * A leading `=` and surrounding whitespace.

    Cell references, ranges, functions and external workbook links stay
    significant; a quoted string is copied verbatim.
    """
    text = formula.strip()
    if text.startswith("="):
        text = text[1:].lstrip()
    text = _canonicalize_row_ceilings(text)
    text = _canonicalize_boolean_literals(text)
    return text


def _canonicalize_row_ceilings(formula: str) -> str:
    out: list[str] = []
    index = 0
    in_string = False
    length = len(formula)
    while index < length:
        current = formula[index]
        if current == '"':
            out.append(current)
            in_string = not in_string
            index += 1
            continue
        if not in_string and current == "$":
            start = index + 1
            end = start
            while end < length and formula[end].isdigit():
                end += 1
            if end > start:
                number = formula[start:end]
                at_reference_end = end >= length or not (
                    formula[end].isalnum() or formula[end] == "_"
                )
                preceded_by_colon = index > 0 and formula[index - 1] == ":"
                preceded_by_column = False
                if index > 0 and formula[index - 1].isalpha():
                    cursor = index - 1
                    while cursor > 0 and formula[cursor - 1].isalpha():
                        cursor -= 1
                    probe = cursor
                    if probe > 0 and formula[probe - 1] == "$":
                        probe -= 1
                    preceded_by_column = probe > 0 and formula[probe - 1] == ":"
                if (
                    at_reference_end
                    and (preceded_by_colon or preceded_by_column)
                    and number.isdigit()
                    and int(number) == _ROW_CEILING_BIFF
                ):
                    out.append(f"${_ROW_CEILING_OOXML}")
                    index = end
                    continue
        out.append(current)
        index += 1
    return "".join(out)


def _canonicalize_boolean_literals(formula: str) -> str:
    out: list[str] = []
    index = 0
    in_string = False
    in_sheet_name = False
    length = len(formula)
    while index < length:
        current = formula[index]
        if current == '"' and not in_sheet_name:
            out.append(current)
            if in_string and formula[index + 1 : index + 2] == '"':
                out.append('"')
                index += 2
                continue
            in_string = not in_string
            index += 1
            continue
        if current == "'" and not in_string:
            out.append(current)
            if in_sheet_name and formula[index + 1 : index + 2] == "'":
                out.append("'")
                index += 2
                continue
            in_sheet_name = not in_sheet_name
            index += 1
            continue
        if not in_string and not in_sheet_name:
            token_len = 0
            upper = formula[index : index + 5].upper()
            if upper.startswith("TRUE"):
                token_len = 4
            elif upper.startswith("FALSE"):
                token_len = 5
            if token_len:
                previous_is_identifier = index > 0 and (
                    formula[index - 1].isalnum() or formula[index - 1] in "_."
                )
                next_char = formula[index + token_len : index + token_len + 1]
                next_is_identifier = bool(next_char) and (
                    next_char.isalnum() or next_char in "_."
                )
                followed_by_empty_call = (
                    formula[index + token_len : index + token_len + 2] == "()"
                )
                if (
                    not previous_is_identifier
                    and not next_is_identifier
                    and followed_by_empty_call
                ):
                    out.append(formula[index : index + token_len])
                    index += token_len + 2
                    continue
        out.append(current)
        index += 1
    return "".join(out)


def _compare(source: list[dict[str, Any]], output: list[dict[str, Any]]) -> dict[str, Any]:
    source_names = [sheet["name"] for sheet in source]
    output_names = [sheet["name"] for sheet in output]
    if source_names != output_names:
        raise ValueError("工作表数量、名称或顺序发生变化")

    formula_count = 0
    value_count = 0
    used_cell_count = 0
    for source_sheet, output_sheet in zip(source, output):
        sheet_name = source_sheet["name"]
        for kind, label, equals in (
            (
                "formulas",
                "公式",
                lambda a, b: _canonicalize_formula(a) == _canonicalize_formula(b),
            ),
            ("values", "非公式数据", _values_equal),
        ):
            before = source_sheet[kind]
            after = output_sheet[kind]
            for address in sorted(set(before) | set(after)):
                if address not in before or address not in after:
                    raise ValueError(
                        f"工作表 {sheet_name!r} 的{label}单元格 {_cell_name(address)} 有增删"
                    )
                if not equals(before[address], after[address]):
                    raise ValueError(
                        f"工作表 {sheet_name!r} 的{label}在 {_cell_name(address)} 不一致："
                        f"{before[address]!r} -> {after[address]!r}"
                    )
        formula_count += len(source_sheet["formulas"])
        value_count += len(source_sheet["values"])
        used_cell_count += source_sheet["used_cells"]

    return {
        "verified": True,
        "level": "excel_com_sheet_order_formula_text_and_cell_values",
        "worksheet_count": len(source),
        "formula_cells": formula_count,
        "value_cells": value_count,
        "used_cells": used_cell_count,
        "numeric_tolerance": 1e-15,
    }


def _verify_institution_normalize_com_snapshots(
    before: list[dict[str, Any]],
    after: list[dict[str, Any]],
    updates: dict[tuple[str, tuple[int, int]], str],
) -> dict[str, int]:
    before_names = [sheet["name"] for sheet in before]
    after_names = [sheet["name"] for sheet in after]
    if before_names != after_names:
        raise ValueError("机构规范化保存后工作表数量、名称或顺序发生变化")
    before_by_name = {sheet["name"]: sheet for sheet in before}
    after_by_name = {sheet["name"]: sheet for sheet in after}
    for name in before_names:
        original = before_by_name[name]
        saved = after_by_name[name]
        if original["formulas"] != saved["formulas"]:
            raise ValueError(f"工作表 {name!r} 的公式文本在 COM 保存后发生变化")
        if set(original["formula_values"]) != set(saved["formula_values"]):
            raise ValueError(f"工作表 {name!r} 的公式结果单元格在 COM 保存后发生变化")
        for address, value in original["formula_values"].items():
            if not _values_equal(value, saved["formula_values"][address]):
                raise ValueError(
                    f"工作表 {name!r} 的公式结果 {_cell_name(address)} 在 COM 保存后发生变化"
                )
        if original["used_cells"] != saved["used_cells"]:
            raise ValueError(f"工作表 {name!r} 的使用区域尺寸在 COM 保存后发生变化")
        expected_updates = {
            address: value
            for (sheet_name, address), value in updates.items()
            if sheet_name == name
        }
        for address in set(original["values"]) | set(saved["values"]):
            if address in expected_updates:
                continue
            if address not in original["values"] or address not in saved["values"]:
                raise ValueError(
                    f"工作表 {name!r} 的非目标单元格 {_cell_name(address)} 有增删"
                )
            if type(original["values"][address]) is not type(saved["values"][address]) or original["values"][address] != saved["values"][address]:
                raise ValueError(
                    f"工作表 {name!r} 的非目标单元格 {_cell_name(address)} 值发生变化"
                )
        for address, expected in expected_updates.items():
            if saved["values"].get(address) != expected:
                raise ValueError(
                    f"工作表 {name!r} 的映射单元格 {_cell_name(address)} 未保存预期文本"
                )
    return {
        "worksheet_count": len(before),
        "formula_cells": sum(len(sheet["formulas"]) for sheet in before),
    }


class ConversionSession:
    def __init__(self) -> None:
        import pythoncom

        pythoncom.CoInitialize()
        self._pythoncom = pythoncom
        self._application: Any | None = None
        self._application_name: str | None = None
        self._process_id: int | None = None
        self._word_application: Any | None = None
        self._word_application_name: str | None = None
        self._word_process_id: int | None = None

    # ProgID preference per host family. WPS-native spreadsheet formats (.et /
    # .ett) must be opened by WPS Spreadsheets (KET): Excel would open them as a
    # renamed XLS and report a plausible FileFormat for the wrong container.
    _SPREADSHEET_PROG_IDS = ("Excel.Application", "ket.Application", "KET.Application")
    _SPREADSHEET_PROG_IDS_WPS_FIRST = ("ket.Application", "KET.Application", "Excel.Application")
    _WORD_PROG_IDS = (
        "Word.Application",
        "kwps.Application",
        "KWPS.Application",
        "wps.Application",
    )
    _WORD_PROG_IDS_WPS_FIRST = (
        "kwps.Application",
        "KWPS.Application",
        "wps.Application",
        "Word.Application",
    )

    def _is_wps_spreadsheet(self) -> bool:
        return (self._application_name or "").lower().startswith("ket.")

    def _is_wps_writer(self) -> bool:
        name = (self._word_application_name or "").lower()
        return name.startswith("kwps.") or name.startswith("wps.")

    def _ensure_application(self, prefer_wps: bool | None = None) -> None:
        if prefer_wps is None:
            prefer_wps = os.environ.get("PYTOOLS_COM_PREFER_WPS", "").strip().lower() in {
                "1",
                "true",
                "yes",
                "y",
            }
        if self._application is not None:
            # A batch may mix .et with .xls; WPS-native files need a KET host,
            # so replace an already-created Excel host instead of reusing it.
            if not prefer_wps or self._is_wps_spreadsheet():
                return
            self._reset_application()
        application, name, process_id = self._dispatch_empty_application(
            self._SPREADSHEET_PROG_IDS_WPS_FIRST
            if prefer_wps
            else self._SPREADSHEET_PROG_IDS,
            "Workbooks",
            document_alerts=False,
        )
        self._application = application
        self._application_name = name
        self._process_id = process_id

    def _ensure_word_application(self, prefer_wps: bool = False) -> None:
        if self._word_application is not None:
            # `.wps` / `.wpt` are WPS Writer formats; Word may not open them.
            if not prefer_wps or self._is_wps_writer():
                return
            self._reset_word_application()
        application, name, process_id = self._dispatch_empty_application(
            self._WORD_PROG_IDS_WPS_FIRST if prefer_wps else self._WORD_PROG_IDS,
            "Documents",
            document_alerts=True,
        )
        self._word_application = application
        self._word_application_name = name
        self._word_process_id = process_id

    def _dispatch_empty_application(
        self,
        prog_ids: tuple[str, ...],
        collection_name: str,
        document_alerts: bool,
    ) -> tuple[Any, str, int | None]:
        """Create a COM host only when it can be proven to be a fresh, empty
        instance owned by this worker. Never attach to a session that already
        holds user documents."""
        import win32com.client

        errors: list[str] = []
        for prog_id in prog_ids:
            try:
                application = win32com.client.DispatchEx(prog_id)
            except Exception as error:
                errors.append(f"{prog_id}: {error}")
                continue
            try:
                collection = getattr(application, collection_name)
                open_documents = int(collection.Count)
            except Exception as error:
                # We cannot prove this is an empty instance, so release only
                # our COM reference; never call Quit on an unknown session.
                errors.append(f"{prog_id}: 无法确认实例中的文档数量：{error}")
                continue
            if open_documents != 0:
                errors.append(
                    f"{prog_id}: 返回了含 {open_documents} 个已打开文档的实例；为保护用户数据已跳过"
                )
                continue
            try:
                application.Visible = False
                if document_alerts:
                    # Word expects the WdAlertLevel enum (wdAlertsNone = 0).
                    application.DisplayAlerts = 0
                else:
                    application.DisplayAlerts = False
                    application.ScreenUpdating = False
                    application.EnableEvents = False
                try:
                    application.AutomationSecurity = 3  # msoAutomationSecurityForceDisable
                except Exception:
                    pass
                process_id: int | None = None
                try:
                    import win32process

                    process_id = int(
                        win32process.GetWindowThreadProcessId(int(application.Hwnd))[1]
                    )
                except Exception:
                    process_id = None
                return application, prog_id, process_id
            except Exception as error:
                try:
                    application.Quit()
                except Exception:
                    pass
                errors.append(f"{prog_id}: 初始化空 COM 实例失败：{error}")
        raise RuntimeError("无法启动 Office/WPS COM：" + "；".join(errors))

    def _quit_owned_instance(self, application: Any, process_id: int | None) -> None:
        """Quit our own automation instance, terminating it by exact PID if
        Quit leaves the process behind.

        Automation hosts frequently ignore Quit (a modal state or a reference
        we lost), and every leaked EXCEL.EXE keeps file handles and can make
        the next `Workbooks.Open` fail. Only the PID recorded for the instance
        this worker created is ever terminated - never a process picked by
        image name, which could hit a user's open Excel.
        """
        try:
            application.Quit()
        except Exception:
            pass
        if process_id is None:
            return
        import time

        deadline = time.monotonic() + 5.0
        while time.monotonic() < deadline:
            if not _process_is_alive(process_id):
                return
            time.sleep(0.1)
        try:
            import win32api
            import win32con

            handle = win32api.OpenProcess(
                win32con.PROCESS_TERMINATE | win32con.PROCESS_QUERY_LIMITED_INFORMATION,
                False,
                process_id,
            )
            try:
                win32api.TerminateProcess(handle, 0)
            finally:
                win32api.CloseHandle(handle)
        except Exception:
            pass

    def _reset_application(self) -> None:
        application = self._application
        process_id = self._process_id
        self._application = None
        self._application_name = None
        self._process_id = None
        if application is not None:
            self._quit_owned_instance(application, process_id)

    def _reset_word_application(self) -> None:
        application = self._word_application
        process_id = self._word_process_id
        self._word_application = None
        self._word_application_name = None
        self._word_process_id = None
        if application is not None:
            self._quit_owned_instance(application, process_id)

    def convert(self, source_text: str, output_text: str) -> dict[str, Any]:
        source = _absolute(source_text)
        output = _absolute(output_text)
        stage = "检查转换路径"
        workbook: Any | None = None
        reopened: Any | None = None
        source_file_format: int | None = None
        source_link_sources: list[str] = []
        try:
            if not source.is_file():
                raise FileNotFoundError(f"输入文件不存在：{source}")
            if output.exists():
                raise FileExistsError(f"拒绝覆盖已存在的候选文件：{output}")
            stage = "检查源文件容器"
            _validate_container(source, "源工作簿")
            suffix = source.suffix.lower()
            # WPS-native spreadsheet formats must be opened by WPS itself.
            self._ensure_application(prefer_wps=suffix in WPS_SPREADSHEET_SUFFIXES)
            application = self._application
            if application is None:
                raise RuntimeError("Excel/WPS COM 尚未启动")

            stage = "读取源工作簿"
            workbook = application.Workbooks.Open(str(source), 0, True)
            # Record what Office/WPS recognized from the container itself,
            # not only from ET/ETT suffix checks. Rust uses this evidence for
            # an OLE/BIFF workbook whose filename was misleadingly changed to
            # .xlsx, and rejects non-Excel or WPS-native containers there.
            source_file_format = int(workbook.FileFormat)
            source_link_sources = _excel_link_sources(workbook)
            # `.et` / `.ett` come in two shapes:
            #   * WPS-native ET/ETT (FileFormat >= 65521) — the original target,
            #     still validated strictly before publishing;
            #   * a BIFF workbook carrying an .et suffix, which is what WPS 2019
            #     actually writes when the user picks "save as .et" through the
            #     COM automation layer (measured 2026-11-09 on four samples,
            #     all FileFormat=56). Content decides: a BIFF payload is
            #     converted like .xls, and the evidence records that it was not
            #     native WPS so Rust can gate on it.
            if suffix in WPS_SPREADSHEET_SUFFIXES:
                stage = "校验源文件格式（原生 ET 或 BIFF 内容）"
                if source_file_format < WPS_NATIVE_FILE_FORMAT_MIN:
                    # Accept BIFF only; anything else (e.g. an OOXML workbook
                    # renamed to .et) is still refused.
                    if source_file_format not in BIFF_FILE_FORMATS:
                        raise ValueError(
                            f"Input has {suffix} extension but the application opened "
                            f"FileFormat {source_file_format}; expected a WPS-native "
                            f"FileFormat >= {WPS_NATIVE_FILE_FORMAT_MIN} "
                            f"(native ET is {ET_FILE_FORMAT}) or a BIFF workbook "
                            f"({sorted(BIFF_FILE_FORMATS)})"
                        )
            stage = "调用 Excel/WPS 完整重算工作簿"
            recalc_method, recalc_errors = _force_recalc(application)
            stage = "读取源工作簿公式和值"
            source_snapshot = _snapshot(workbook)
            formula_cells = sum(len(sheet["formulas"]) for sheet in source_snapshot)
            if recalc_method is None:
                if formula_cells:
                    # Both methods failed and the workbook has formulas: the
                    # cached values may be stale, and the source/candidate
                    # comparison below cannot detect it (it would compare the
                    # same stale numbers on both sides).
                    raise RuntimeError(
                        f"工作簿无法重算（{formula_cells} 个公式单元格）；"
                        f"为避免发布陈旧公式缓存，已取消转换："
                        + "；".join(recalc_errors)
                    )
                # No formulas: nothing to recalculate, so the failure is inert.
                # Keep the reason in the evidence instead of hiding it.
            stage = "通过 pywin32 写入 XLSX 候选"
            workbook.SaveAs(str(output), XLSX_FILE_FORMAT)
            workbook.Close(SaveChanges=False)
            workbook = None

            stage = "保留外部工作簿链接目标"
            _restore_external_link_targets(output, source_link_sources)

            stage = "重新打开候选工作簿"
            reopened = application.Workbooks.Open(str(output), 0, True)
            stage = "读取候选工作簿并验证公式和值"
            output_snapshot = _snapshot(reopened)
            validation = _compare(source_snapshot, output_snapshot)
            output_link_sources = _excel_link_sources(reopened)
            if _normalized_link_sources(source_link_sources) != _normalized_link_sources(
                output_link_sources
            ):
                raise ValueError(
                    "另存后外部工作簿链接目标发生变化："
                    f"源链接 {source_link_sources!r}，转出链接 {output_link_sources!r}"
                )
            validation["external_link_count"] = len(source_link_sources)
            validation["external_links_preserved"] = True
            if recalc_method is not None:
                validation["recalc_method"] = recalc_method
                validation["calculation"] = f"{recalc_method} before SaveAs"
                if recalc_errors:
                    validation["recalc_fallback_errors"] = recalc_errors
            else:
                validation["recalc_failed"] = True
                validation["recalc_errors"] = recalc_errors
                validation["calculation"] = "no formulas; recalc not required"
            if source_file_format is not None:
                validation["source_file_format"] = source_file_format
            if suffix in WPS_SPREADSHEET_SUFFIXES:
                validation["source_extension"] = suffix.lstrip(".")
                # Native ET/ETT keeps the original meaning of this flag; a BIFF
                # payload records that it was verified *as BIFF* instead, so
                # Rust can tell the two apart.
                validation["source_format_verified"] = (
                    source_file_format is not None
                    and source_file_format >= WPS_NATIVE_FILE_FORMAT_MIN
                )
                validation["source_content_kind"] = (
                    "wps_native"
                    if validation["source_format_verified"]
                    else "biff"
                )
            reopened.Close(SaveChanges=False)
            reopened = None
            return {
                "ok": True,
                "application": self._application_name,
                "process_id": self._process_id,
                "validation": validation,
            }
        except Exception as error:
            for opened in (reopened, workbook):
                if opened is not None:
                    try:
                        opened.Close(SaveChanges=False)
                    except Exception:
                        pass
            # Reset only the DispatchEx instance owned by this worker.
            self._reset_application()
            return {
                "ok": False,
                "stage": stage,
                "error": f"{type(error).__name__}: {error}",
                "traceback": traceback.format_exc(limit=6),
            }

    def _open_workbook(self, path: Path, *, read_only: bool, label: str) -> Any:
        """Open one workbook, rebuilding the host once if Excel rejects it."""
        first_error: Exception | None = None
        for attempt in (1, 2):
            assert self._application is not None
            try:
                return self._application.Workbooks.Open(str(path), 0, read_only)
            except Exception as error:
                if attempt == 2:
                    raise RuntimeError(
                        f"{label}无法打开：{path}\n"
                        f"第一次错误：{type(first_error).__name__}: {first_error}\n"
                        f"重建 Excel/WPS 实例后重试仍失败：{type(error).__name__}: {error}"
                    ) from error
                first_error = error
                self._reset_application()
                self._ensure_application()
        raise AssertionError("unreachable")

    @staticmethod
    def _legacy_xls_path(value: str, label: str) -> Path:
        path = _absolute(value)
        if path.suffix.lower() != ".xls":
            raise ValueError(f"{label}的 COM 业务路由仅接受 .xls：{path}")
        if not path.is_file():
            raise FileNotFoundError(f"{label}不存在：{path}")
        return path

    @staticmethod
    def _legacy_number(value: Any) -> float:
        if value is None:
            return 0.0
        try:
            return float(value)
        except (ValueError, TypeError):
            try:
                return float(str(value).strip().replace(",", ""))
            except (ValueError, TypeError):
                return 0.0

    @staticmethod
    def _legacy_normalize(value: Any) -> str:
        if value is None:
            return ""
        text = re.sub(r"[\r\n\t　\u00a0]", " ", str(value)).strip()
        while "  " in text:
            text = text.replace("  ", " ")
        return text

    def institution_normalize_com(
        self, candidate_text: str, updates: list[dict[str, str]]
    ) -> dict[str, Any]:
        """Apply Rust-planned institution-name updates to an isolated `.xls` copy."""
        stage = "检查 5-2 COM 候选"
        candidate = self._legacy_xls_path(candidate_text, "规范化候选工作簿")
        workbook = None
        try:
            if not updates:
                raise ValueError("没有可写入的机构名称更新")
            self._ensure_application()
            if self._application is not None:
                try:
                    self._application.Calculation = -4135  # xlCalculationManual, matching PyCode.
                except Exception:
                    pass
            stage = "通过 Excel/WPS 打开 5-2 候选工作簿"
            workbook = self._open_workbook(candidate, read_only=False, label="规范化候选工作簿")
            sheets = {str(sheet.Name): sheet for sheet in workbook.Worksheets}
            stage = "记录 5-2 候选原始公式、数据和合并区域"
            before_snapshot = _snapshot(workbook)
            applied = 0
            expected_updates: dict[tuple[str, tuple[int, int]], str] = {}
            seen: set[tuple[str, tuple[int, int]]] = set()
            stage = "写入机构映射结果"
            for update in updates:
                sheet_name = str(update.get("sheet", ""))
                address = str(update.get("cell", ""))
                value = update.get("value")
                if not sheet_name or not address or not isinstance(value, str):
                    raise ValueError("机构名称更新缺少工作表、单元格或文本值")
                key = (sheet_name, address.upper())
                if key in seen:
                    raise ValueError(f"同一候选包含重复单元格更新：{sheet_name}!{address}")
                seen.add(key)
                if sheet_name not in sheets:
                    raise ValueError(f"候选工作簿缺少工作表：{sheet_name}")
                if value.startswith("="):
                    raise ValueError(f"机构映射结果不能写成公式：{sheet_name}!{address}")
                target = sheets[sheet_name].Range(address)
                coordinate = (int(target.Row), int(target.Column))
                key = (sheet_name, coordinate)
                if key in seen:
                    raise ValueError(f"同一候选包含重复单元格更新：{sheet_name}!{address}")
                seen.add(key)
                expected_updates[key] = value
                target.Value = value
                applied += 1
            stage = "保存 5-2 COM 候选工作簿"
            workbook.Save()
            stage = "核对 5-2 COM 保存结果"
            after_snapshot = _snapshot(workbook)
            verification = _verify_institution_normalize_com_snapshots(
                before_snapshot, after_snapshot, expected_updates
            )
            workbook.Close(SaveChanges=False)
            workbook = None
            return {
                "ok": True,
                "operation": "institution_normalize_com",
                "candidate": str(candidate),
                "application": self._application_name,
                "process_id": self._process_id,
                "applied_count": applied,
                "verified": True,
                "verification": verification,
            }
        except Exception as error:
            if workbook is not None:
                try:
                    workbook.Close(SaveChanges=False)
                except Exception:
                    pass
            return {
                "ok": False,
                "stage": stage,
                "error": f"{type(error).__name__}: {error}",
                "traceback": traceback.format_exc(limit=6),
            }

    def village_split_com(self, current_text: str, previous_text: str, candidate_text: str) -> dict[str, Any]:
        """Apply PyCode's 5-1 COM algorithm to an isolated `.xls` candidate."""
        stage = "检查 5-1 COM 输入"
        current = self._legacy_xls_path(current_text, "本期工作簿")
        previous = self._legacy_xls_path(previous_text, "去年同期工作簿")
        candidate = self._legacy_xls_path(candidate_text, "目标候选工作簿")
        current_wb = previous_wb = target_wb = None
        try:
            self._ensure_application()
            application = self._application
            if application is None:
                raise RuntimeError("Excel/WPS COM 尚未启动")
            try:
                application.Calculation = -4135  # xlCalculationManual; matches PyCode.
            except Exception:
                pass

            stage = "读取本期和去年同期村镇银行数据"
            current_wb = self._open_workbook(current, read_only=True, label="本期工作簿")
            previous_wb = self._open_workbook(previous, read_only=True, label="去年同期工作簿")
            matrix: list[list[Any]] = [[0] * 8 for _ in range(6)]
            indicators = ("村镇银行本外币", "村镇银行人民币", "村镇银行外汇")
            for index, indicator in enumerate(indicators):
                current_sheet = next(
                    (sheet for sheet in current_wb.Worksheets if indicator in str(sheet.Name)),
                    None,
                )
                if current_sheet is None:
                    continue
                previous_sheet = next(
                    (sheet for sheet in previous_wb.Worksheets if str(sheet.Name) == str(current_sheet.Name)),
                    None,
                )
                if previous_sheet is None:
                    continue
                for row_offset, base_row in enumerate((5, 6)):
                    target_row = index + row_offset * 3
                    current_deposit = self._legacy_number(current_sheet.Cells(base_row, 3).Value)
                    previous_deposit = self._legacy_number(previous_sheet.Cells(base_row, 3).Value)
                    current_loan = self._legacy_number(current_sheet.Cells(base_row, 87).Value)
                    previous_loan = self._legacy_number(previous_sheet.Cells(base_row, 87).Value)
                    matrix[target_row][0] = current_deposit
                    matrix[target_row][1] = current_sheet.Cells(base_row, 4).Value
                    matrix[target_row][2] = current_sheet.Cells(base_row, 5).Value
                    matrix[target_row][3] = (
                        0 if current_deposit == 0 else "∞"
                    ) if previous_deposit == 0 else round((current_deposit / previous_deposit - 1) * 100, 2)
                    matrix[target_row][4] = current_loan
                    matrix[target_row][5] = current_sheet.Cells(base_row, 88).Value
                    matrix[target_row][6] = current_sheet.Cells(base_row, 89).Value
                    matrix[target_row][7] = (
                        0 if current_loan == 0 else "∞"
                    ) if previous_loan == 0 else round((current_loan / previous_loan - 1) * 100, 2)
            if all(matrix[0][column] in (0, None) for column in (0, 4)):
                return {
                    "ok": True,
                    "matrix_empty": True,
                    "application": self._application_name,
                    "process_id": self._process_id,
                    "candidate": str(candidate),
                    "branch_sheets": 0,
                    "filled_sheets": 0,
                    "skipped_sheets": 0,
                }
            current_wb.Close(SaveChanges=False)
            current_wb = None
            previous_wb.Close(SaveChanges=False)
            previous_wb = None

            stage = "通过 Excel/WPS 修改目标候选工作簿"
            target_wb = self._open_workbook(candidate, read_only=False, label="目标候选工作簿")
            counts = {"branch_sheets": 0, "filled_sheets": 0, "skipped_sheets": 0}
            row_names = ("博罗长江村镇银行", "惠东惠民村镇银行")
            for sheet in target_wb.Worksheets:
                title = str(sheet.Name)
                if "分机构" not in title:
                    continue
                counts["branch_sheets"] += 1
                if "本外币" in title:
                    kind = 0
                elif "人民币" in title:
                    kind = 1
                elif "外汇" in title:
                    kind = 2
                else:
                    counts["skipped_sheets"] += 1
                    continue
                used = sheet.UsedRange
                last_row = int(used.Row) + int(used.Rows.Count) - 1
                village_row = next(
                    (
                        row
                        for row in range(1, last_row + 1)
                        if sheet.Cells(row, 1).Value is not None
                        and str(sheet.Cells(row, 1).Value).strip() == "村镇银行"
                    ),
                    None,
                )
                if village_row is None:
                    counts["skipped_sheets"] += 1
                    continue
                sheet.Rows(village_row).Delete()
                sheet.Rows(village_row).Insert()
                sheet.Rows(village_row).Insert()
                sheet.Cells(village_row, 1).Value = row_names[0]
                sheet.Cells(village_row + 1, 1).Value = row_names[1]
                for column in range(8):
                    sheet.Cells(village_row, column + 2).Value = matrix[kind][column]
                    sheet.Cells(village_row + 1, column + 2).Value = matrix[kind + 3][column]
                counts["filled_sheets"] += 1
            stage = "保存 5-1 COM 候选工作簿"
            target_wb.Save()
            target_wb.Close(SaveChanges=False)
            target_wb = None
            return {
                "ok": True,
                "matrix_empty": False,
                "application": self._application_name,
                "process_id": self._process_id,
                "candidate": str(candidate),
                **counts,
            }
        except Exception as error:
            for workbook in (target_wb, previous_wb, current_wb):
                if workbook is not None:
                    try:
                        workbook.Close(SaveChanges=False)
                    except Exception:
                        pass
            return {"ok": False, "stage": stage, "error": f"{type(error).__name__}: {error}"}

    def rural_loan_com(self, current_candidate_text: str, previous_text: str) -> dict[str, Any]:
        """Apply PyCode's 5-7 COM algorithm to an isolated current `.xls` candidate."""
        stage = "检查 5-7 COM 输入"
        current = self._legacy_xls_path(current_candidate_text, "本期候选工作簿")
        previous = self._legacy_xls_path(previous_text, "上期工作簿")
        current_wb = previous_wb = None
        try:
            self._ensure_application()
            application = self._application
            if application is None:
                raise RuntimeError("Excel/WPS COM 尚未启动")
            try:
                application.Calculation = -4135  # xlCalculationManual; matches PyCode.
            except Exception:
                pass
            stage = "打开本期候选与上期工作簿"
            current_wb = self._open_workbook(current, read_only=False, label="本期候选工作簿")
            previous_wb = self._open_workbook(previous, read_only=True, label="上期工作簿")

            def unique_sheet(workbook: Any, keys: tuple[str, ...], label: str) -> Any:
                matches = [sheet for sheet in workbook.Worksheets if all(key in str(sheet.Name) for key in keys)]
                if len(matches) > 1:
                    raise ValueError(f"{label}匹配到多个工作表（关键词 {keys}）")
                return matches[0] if matches else None

            def last_row(sheet: Any) -> int:
                used = sheet.UsedRange
                return int(used.Row) + int(used.Rows.Count) - 1

            def last_col(sheet: Any) -> int:
                used = sheet.UsedRange
                return int(used.Column) + int(used.Columns.Count) - 1

            def institution_map(sheet: Any) -> tuple[dict[str, float], dict[str, int], list[str]]:
                values: dict[str, float] = {}
                rows: dict[str, int] = {}
                duplicates: list[str] = []
                for row in range(3, last_row(sheet) + 1):
                    name = self._legacy_normalize(sheet.Cells(row, 2).Value)
                    if not name:
                        continue
                    if name in values:
                        duplicates.append(name)
                        continue
                    values[name] = self._legacy_number(sheet.Cells(row, 3).Value)
                    rows[name] = row
                return values, rows, duplicates

            stage = "定位涉农贷款分机构与分地区工作表"
            current_inst = unique_sheet(current_wb, ("惠州市", "涉农贷款分机构"), "本期分机构")
            previous_inst = unique_sheet(previous_wb, ("惠州市", "涉农贷款分机构"), "上期分机构")
            current_area = unique_sheet(current_wb, ("本外币", "涉农贷款分地区"), "本期分地区")
            if current_inst is None:
                raise ValueError("本期文件未找到包含“惠州市”和“涉农贷款分机构”的工作表")
            if previous_inst is None:
                raise ValueError("上期文件未找到包含“惠州市”和“涉农贷款分机构”的工作表")
            if current_area is None:
                raise ValueError("本期文件未找到包含“本外币”和“涉农贷款分地区”的工作表")

            current_values, current_rows, current_duplicates = institution_map(current_inst)
            previous_values, _, previous_duplicates = institution_map(previous_inst)
            if current_duplicates:
                raise ValueError("本期分机构表存在重复机构名称：" + "、".join(current_duplicates))
            if previous_duplicates:
                raise ValueError("上期分机构表存在重复机构名称：" + "、".join(previous_duplicates))

            stage = "计算涉农贷款比上月并写入本期候选"
            fill_total = 0.0
            unmatched: list[str] = []
            changed_cells = 0
            for name, row in current_rows.items():
                if name in previous_values:
                    difference = current_values[name] - previous_values[name]
                    current_inst.Cells(row, 4).Value = difference
                    fill_total += difference
                else:
                    current_inst.Cells(row, 4).Value = None
                    unmatched.append(name)
                changed_cells += 1
            previous_only = [name for name in previous_values if name not in current_values]

            header_col = 0
            for row in range(1, min(last_row(current_area), 10) + 1):
                for column in range(1, last_col(current_area) + 1):
                    if self._legacy_normalize(current_area.Cells(row, column).Value) == "涉农贷款比上月":
                        header_col = column
                        break
                if header_col:
                    break
            if header_col == 0:
                raise ValueError('未在“本外币涉农贷款分地区”表中找到表头“涉农贷款比上月”')
            city_row = 0
            for row in range(1, last_row(current_area) + 1):
                if any(
                    self._legacy_normalize(current_area.Cells(row, column).Value) == "惠州市"
                    for column in range(1, 4)
                ):
                    city_row = row
                    break
            if city_row == 0:
                raise ValueError('未在“本外币涉农贷款分地区”表中找到“惠州市”所在行')
            area_total = self._legacy_number(current_area.Cells(city_row, header_col).Value)

            stage = "保存 5-7 COM 候选工作簿"
            current_wb.Save()
            current_wb.Close(SaveChanges=False)
            current_wb = None
            previous_wb.Close(SaveChanges=False)
            previous_wb = None
            return {
                "ok": True,
                "application": self._application_name,
                "process_id": self._process_id,
                "changed_cells": changed_cells,
                "summary": {
                    "matched": len(current_rows) - len(unmatched),
                    "fill_total": fill_total,
                    "area_total": area_total,
                    "consistent": abs(fill_total - area_total) <= 0.01,
                    "unmatched_current": unmatched,
                    "previous_only": previous_only,
                    "error": None,
                },
            }
        except Exception as error:
            for workbook in (previous_wb, current_wb):
                if workbook is not None:
                    try:
                        workbook.Close(SaveChanges=False)
                    except Exception:
                        pass
            return {"ok": False, "stage": stage, "error": f"{type(error).__name__}: {error}"}

    def duplicate_mark_com(
        self, candidate_text: str, marks: list[dict[str, str]]
    ) -> dict[str, Any]:
        """Apply Rust-planned 2-1 red-fill marks to an isolated `.xls` copy."""
        stage = "检查 2-1 COM 候选"
        candidate = self._legacy_xls_path(candidate_text, "标记候选工作簿")
        workbook = None
        try:
            if not marks:
                raise ValueError("没有可写入的重复标记")
            self._ensure_application()
            if self._application is not None:
                try:
                    self._application.Calculation = -4135  # xlCalculationManual
                except Exception:
                    pass
            stage = "通过 Excel/WPS 打开 2-1 候选工作簿"
            workbook = self._open_workbook(candidate, read_only=False, label="标记候选工作簿")
            sheets = {str(sheet.Name): sheet for sheet in workbook.Worksheets}
            stage = "写入重复行红色标记"
            verified_marks = []
            for mark in marks:
                sheet_name = str(mark.get("sheet", ""))
                address = str(mark.get("cell", "")).upper()
                if not sheet_name or not address:
                    raise ValueError("重复标记缺少工作表或单元格")
                if sheet_name not in sheets:
                    raise ValueError(f"候选工作簿缺少工作表：{sheet_name}")
                target = sheets[sheet_name].Range(address)
                target.Interior.Color = 255  # RGB(255, 0, 0)
                verified_marks.append((sheet_name, address))
            stage = "保存 2-1 COM 候选工作簿"
            workbook.Save()
            stage = "核对 2-1 COM 标记结果"
            applied = 0
            mismatches: list[str] = []
            for sheet_name, address in verified_marks:
                color = sheets[sheet_name].Range(address).Interior.Color
                if color == 255:
                    applied += 1
                else:
                    mismatches.append(f"{sheet_name}!{address} color={color}")
            if mismatches:
                raise ValueError("标记回读不一致：" + "、".join(mismatches[:5]))
            workbook.Close(SaveChanges=False)
            workbook = None
            return {
                "ok": True,
                "operation": "duplicate_mark_com",
                "candidate": str(candidate),
                "application": self._application_name,
                "process_id": self._process_id,
                "applied_count": applied,
                "verified": True,
            }
        except Exception as error:
            if workbook is not None:
                try:
                    workbook.Close(SaveChanges=False)
                except Exception:
                    pass
            return {
                "ok": False,
                "stage": stage,
                "error": f"{type(error).__name__}: {error}",
                "traceback": traceback.format_exc(limit=6),
            }

    def delete_rows_com(
        self, candidate_text: str, deletions: list[dict[str, Any]]
    ) -> dict[str, Any]:
        """Delete Rust-planned whole rows on an isolated `.xls` copy (2-2/5-3)."""
        stage = "检查删行 COM 候选"
        candidate = self._legacy_xls_path(candidate_text, "删行候选工作簿")
        workbook = None
        try:
            if not deletions:
                raise ValueError("没有可执行的删行计划")
            self._ensure_application()
            if self._application is not None:
                try:
                    self._application.Calculation = -4135  # xlCalculationManual
                except Exception:
                    pass
            stage = "通过 Excel/WPS 打开删行候选工作簿"
            workbook = self._open_workbook(candidate, read_only=False, label="删行候选工作簿")
            sheets = {str(sheet.Name): sheet for sheet in workbook.Worksheets}

            def used_rows(sheet: Any) -> int:
                used = sheet.UsedRange
                return int(used.Row) + int(used.Rows.Count) - 1

            stage = "记录删行前行数"
            before_rows = {name: used_rows(sheet) for name, sheet in sheets.items()}
            expected: dict[str, int] = {}
            grouped: dict[str, list[tuple[int, int]]] = {}
            for deletion in deletions:
                sheet_name = str(deletion.get("sheet", ""))
                first_row = int(deletion.get("first_row", 0))
                count = int(deletion.get("count", 0))
                if not sheet_name or first_row < 1 or count < 1:
                    raise ValueError("删行计划缺少工作表、起始行或行数")
                if sheet_name not in sheets:
                    raise ValueError(f"候选工作簿缺少工作表：{sheet_name}")
                grouped.setdefault(sheet_name, []).append((first_row, count))
            stage = "按工作表倒序删除整行"
            deleted_rows = 0
            for sheet_name, runs in grouped.items():
                sheet = sheets[sheet_name]
                # 倒序删除，保持 Rust 计划里的原始行号语义。
                for first_row, count in sorted(runs, key=lambda item: item[0], reverse=True):
                    sheet.Rows(f"{first_row}:{first_row + count - 1}").Delete()
                    deleted_rows += count
                expected[sheet_name] = (
                    before_rows[sheet_name]
                    - sum(count for _, count in grouped[sheet_name])
                )
            stage = "保存删行候选工作簿"
            workbook.Save()
            stage = "核对删行后行数"
            mismatches = []
            for sheet_name, expect in expected.items():
                actual = used_rows(sheets[sheet_name])
                if actual != expect:
                    mismatches.append(f"{sheet_name}: 行数 {actual} ≠ 预期 {expect}")
            if mismatches:
                raise ValueError("删行后行数校验不一致：" + "；".join(mismatches[:5]))
            workbook.Close(SaveChanges=False)
            workbook = None
            return {
                "ok": True,
                "operation": "delete_rows_com",
                "candidate": str(candidate),
                "application": self._application_name,
                "process_id": self._process_id,
                "deleted_rows": deleted_rows,
                "verified": True,
            }
        except Exception as error:
            if workbook is not None:
                try:
                    workbook.Close(SaveChanges=False)
                except Exception:
                    pass
            return {
                "ok": False,
                "stage": stage,
                "error": f"{type(error).__name__}: {error}",
                "traceback": traceback.format_exc(limit=6),
            }

    def fx_header_com(self, candidate_text: str) -> dict[str, Any]:
        """Apply PyCode's 5-4 COM algorithm to an isolated `.xls` copy."""
        stage = "检查 5-4 COM 候选"
        candidate = self._legacy_xls_path(candidate_text, "外汇页眉候选工作簿")
        workbook = None
        old_unit = "万元"
        new_unit = "万美元"

        def replace_in_default(ps) -> bool:
            changed = False
            for attr in ("LeftHeader", "CenterHeader", "RightHeader"):
                try:
                    text = getattr(ps, attr) or ""
                    if old_unit in text:
                        setattr(ps, attr, text.replace(old_unit, new_unit))
                        changed = True
                except Exception:
                    pass
            return changed

        def replace_in_subpage(sub_obj) -> bool:
            changed = False
            if sub_obj is None:
                return False
            for attr in ("LeftHeader", "CenterHeader", "RightHeader"):
                try:
                    hf = getattr(sub_obj, attr)
                    if hf is None:
                        continue
                    text = hf.Text or ""
                    if old_unit in text:
                        hf.Text = text.replace(old_unit, new_unit)
                        changed = True
                except Exception:
                    pass
            return changed

        try:
            self._ensure_application()
            if self._application is None:
                raise RuntimeError("Excel/WPS COM 尚未启动")
            stage = "通过 Excel/WPS 打开 5-4 候选工作簿"
            workbook = self._open_workbook(candidate, read_only=False, label="外汇页眉候选工作簿")
            fx_sheets = 0
            modified_sheets = 0
            for sheet in workbook.Worksheets:
                if "外汇" not in str(sheet.Name):
                    continue
                fx_sheets += 1
                ps = sheet.PageSetup
                changed = replace_in_default(ps)
                try:
                    if ps.OddAndEvenPagesHeaderFooter:
                        if replace_in_subpage(ps.EvenPage):
                            changed = True
                except Exception:
                    pass
                try:
                    if ps.DifferentFirstPageHeaderFooter:
                        if replace_in_subpage(ps.FirstPage):
                            changed = True
                except Exception:
                    pass
                if changed:
                    modified_sheets += 1
            if modified_sheets:
                stage = "保存 5-4 COM 候选工作簿"
                workbook.Save()
            stage = "核对 5-4 COM 页眉结果"
            residual = []
            for sheet in workbook.Worksheets:
                if "外汇" not in str(sheet.Name):
                    continue
                for attr in ("LeftHeader", "CenterHeader", "RightHeader"):
                    text = getattr(sheet.PageSetup, attr) or ""
                    if old_unit in text:
                        residual.append(f"{sheet.Name}.{attr}")
            if residual:
                raise ValueError("外汇表页眉仍残留旧单位：" + "、".join(residual[:5]))
            workbook.Close(SaveChanges=False)
            workbook = None
            return {
                "ok": True,
                "operation": "fx_header_com",
                "candidate": str(candidate),
                "application": self._application_name,
                "process_id": self._process_id,
                "fx_sheets": fx_sheets,
                "modified_sheets": modified_sheets,
                "verified": True,
            }
        except Exception as error:
            if workbook is not None:
                try:
                    workbook.Close(SaveChanges=False)
                except Exception:
                    pass
            return {
                "ok": False,
                "stage": stage,
                "error": f"{type(error).__name__}: {error}",
                "traceback": traceback.format_exc(limit=6),
            }

    def append_rows_com(
        self,
        candidate_text: str,
        sheet_name: str,
        start_row: int,
        rows: list[list[Any]],
    ) -> dict[str, Any]:
        """Append typed 2-D rows to a sheet in an isolated `.xls` target copy.

        `rows` is a list of lists; each inner value may be a str (text),
        int/float (number), bool, or None (blank). Written in bulk via
        `Range.Resize` for speed; the Rust caller owns dedup filtering.
        """
        stage = "检查追加 COM 候选"
        candidate = self._legacy_xls_path(candidate_text, "追加候选工作簿")
        workbook = None
        try:
            if not rows:
                raise ValueError("没有可追加的数据行")
            self._ensure_application()
            if self._application is not None:
                try:
                    self._application.Calculation = -4135
                except Exception:
                    pass
            stage = "通过 Excel/WPS 打开追加候选工作簿"
            workbook = self._open_workbook(candidate, read_only=False, label="追加候选工作簿")
            sheet_names = [str(s.Name) for s in workbook.Worksheets]
            if sheet_name in sheet_names:
                sheet = workbook.Worksheets(sheet_names.index(sheet_name) + 1)
            else:
                sheet = workbook.Worksheets.Add(After=workbook.Worksheets(workbook.Worksheets.Count))
                sheet.Name = sheet_name
            stage = "写入追加数据行"
            num_rows = len(rows)
            num_cols = max(len(r) for r in rows) if rows else 0
            # COM Range.Value 接受 2-D tuple；None → Empty。
            range_obj = sheet.Range(
                sheet.Cells(start_row, 1),
                sheet.Cells(start_row + num_rows - 1, num_cols),
            )
            range_obj.Value = tuple(
                tuple(v if v is not None else "" for v in row) for row in rows
            )
            stage = "保存追加候选工作簿"
            workbook.Save()
            stage = "核对追加结果"
            spot_checks = min(num_rows, 3)
            for i in range(spot_checks):
                row_idx = start_row + i
                col_idx = 1
                for expected in rows[i]:
                    actual = sheet.Cells(row_idx, col_idx).Value
                    expected_str = str(expected) if expected is not None else ""
                    actual_str = str(actual) if actual is not None else ""
                    if expected is not None and expected_str != actual_str \
                        and str(float(expected)) != str(float(actual_str or 0)):
                        raise ValueError(
                            f"追加回读不一致：{sheet_name}!{row_idx},{col_idx} "
                            f"预期 {expected!r}，实际 {actual!r}"
                        )
                    col_idx += 1
            workbook.Close(SaveChanges=False)
            workbook = None
            return {
                "ok": True,
                "operation": "append_rows_com",
                "candidate": str(candidate),
                "application": self._application_name,
                "process_id": self._process_id,
                "appended_rows": num_rows,
                "start_row": start_row,
                "sheet": sheet_name,
                "verified": True,
            }
        except Exception as error:
            if workbook is not None:
                try:
                    workbook.Close(SaveChanges=False)
                except Exception:
                    pass
            return {
                "ok": False,
                "stage": stage,
                "error": f"{type(error).__name__}: {error}",
                "traceback": traceback.format_exc(limit=6),
            }

    def rename_sheets(self, candidate_text: str, changes: list[dict[str, str]]) -> dict[str, Any]:
        """复用 pytools 的 openpyxl 改名方案，并改写跨表引用与命名区域。

        openpyxl 的 title 赋值不会改写公式字符串，因此这里在同一遍里把
        「'旧表'!」与「旧表!」两种引用形式替换为新表名；字符串字面量
        （双引号内）不改写。不经过 COM，无需本机 Excel。
        """
        # COM/文件接口都会拒绝 Windows 扩展长度路径，先剥掉前缀。
        if candidate_text.startswith("\\\\?\\UNC\\"):
            candidate_text = "\\\\" + candidate_text[8:]
        elif candidate_text.startswith("\\\\?\\"):
            candidate_text = candidate_text[4:]
        candidate = _absolute(candidate_text)
        stage = "检查改名候选工作簿"
        workbook = None
        reloaded = None
        try:
            _validate_container(candidate, "改名候选工作簿")
            if candidate.suffix.lower() != ".xlsx" or not changes:
                raise ValueError("工作表改名仅接受非空改名计划和 XLSX 临时候选")
            stage = "读取候选工作簿"
            import openpyxl

            workbook = openpyxl.load_workbook(candidate)
            stage = "校验改名计划"
            renames = {}
            sheetnames = list(workbook.sheetnames)
            # 按计划顺序逐个放行：A→RenamedA、B→A 这类链式/换名在 Excel 里
            # 合法，只要目标名在轮到它时已经空出来。不能拿原始表名集一次性
            # 判重，否则会把旧 pytools 能处理的计划误判为重名。
            seen = set(sheetnames)
            for change in changes:
                old, new = str(change["old_name"]), str(change["new_name"])
                if old not in sheetnames:
                    raise ValueError(f"工作表 {old!r} 不存在")
                if re.search(r"[:\\/?*\[\]]", new) or len(new) > 31:
                    raise ValueError(f"新表名 {new!r} 含非法字符或超过 31 字符")
                if old == new:
                    continue
                names_without_old = {name for name in seen if name != old}
                if any(name.lower() == new.lower() for name in names_without_old):
                    raise ValueError(f"新表名 {new!r} 与现有工作表重名")
                seen.discard(old)
                seen.add(new)
                renames[old] = new
            if not renames:
                raise ValueError("改名计划中没有实际改名")
            stage = "读取改名前快照"
            before = _openpyxl_snapshot(workbook)
            formula_caches = _formula_cache_snapshot(candidate, before, renames)
            stage = "执行改名并改写引用"
            for old, new in renames.items():
                workbook[old].title = new
            for ws in workbook.worksheets:
                for row in ws.iter_rows():
                    for cell in row:
                        value = cell.value
                        if isinstance(value, str) and value.startswith("="):
                            cell.value = "=" + _rewrite_sheet_refs(value[1:], renames)
            for names in [workbook.defined_names] + [
                ws.defined_names for ws in workbook.worksheets
            ]:
                for name in list(names):
                    defined = names[name]
                    defined.attr_text = _rewrite_sheet_refs(defined.attr_text or "", renames)
            stage = "保存改名候选"
            workbook.save(candidate)
            workbook.close()
            workbook = None
            stage = "恢复空字符串单元格"
            _restore_empty_string_cells(candidate, before, renames)
            stage = "恢复公式缓存值"
            _restore_formula_caches(candidate, formula_caches)
            stage = "重新读取并校验"
            reloaded = openpyxl.load_workbook(candidate)
            after = _openpyxl_snapshot(reloaded)
            expected_names = [
                renames.get(name, name) for name in [sheet["name"] for sheet in before]
            ]
            if [sheet["name"] for sheet in after] != expected_names:
                raise ValueError("改名后工作表顺序或名称与计划不一致")
            for source, result in zip(before, after):
                if set(source["formulas"]) != set(result["formulas"]):
                    raise ValueError(f"工作表 {source['name']!r} 的公式单元格发生增删")
                for address, formula in source["formulas"].items():
                    expected = "=" + _rewrite_sheet_refs(formula, renames)
                    if result["formulas"].get(address) != expected[1:]:
                        raise ValueError(
                            f"工作表 {source['name']!r} 公式 {address} 引用改写不一致"
                        )
                for address, value in source["values"].items():
                    if not _values_equal(value, result["values"].get(address)):
                        raise ValueError(
                            f"工作表 {source['name']!r} 非公式单元格 {address} 发生变化"
                        )
            reloaded.close()
            reloaded = None
            if _count_ref_errors(after) > _count_ref_errors(before):
                raise ValueError("改名后出现新的 #REF! 公式引用")
            validation = {
                "verified": True,
                "level": "openpyxl_rename_with_reference_rewrite",
                "renamed_sheets": len(renames),
            }
            return {"ok": True, "validation": validation}
        except Exception as error:
            for opened in (reloaded, workbook):
                if opened is not None:
                    try:
                        opened.close()
                    except Exception:
                        pass
            return {"ok": False, "stage": stage, "error": f"{type(error).__name__}: {error}"}

    def rename_sheets_com(self, candidate_text: str, changes: list[dict[str, str]]) -> dict[str, Any]:
        """Rename a legacy .xls candidate with Excel/WPS, as PyCode does.

        Excel updates formula references as part of Worksheet.Name assignment.
        Rust supplies an isolated candidate and publishes it only after this
        method has reopened it and checked the sheet order, formulas and values.
        """
        if candidate_text.startswith("\\\\?\\UNC\\"):
            candidate_text = "\\\\" + candidate_text[8:]
        elif candidate_text.startswith("\\\\?\\"):
            candidate_text = candidate_text[4:]
        candidate = _absolute(candidate_text)
        stage = "检查 COM 改名候选工作簿"
        workbook = None
        reopened = None
        try:
            if candidate.suffix.lower() != ".xls":
                raise ValueError("Excel COM 改名路径仅接受 .xls 候选")
            if not changes:
                raise ValueError("Excel COM 改名计划为空")
            self._ensure_application()
            stage = "通过 Excel/WPS 打开候选工作簿"
            workbook = self._open_workbook(candidate, read_only=False, label="改名候选工作簿")
            before = _snapshot(workbook)
            before_names = [sheet["name"] for sheet in before]
            current_names = set(before_names)
            renames: dict[str, str] = {}
            stage = "按配置修改工作表名称"
            for change in changes:
                old = str(change["old_name"])
                new = str(change["new_name"])
                if old not in current_names:
                    raise ValueError(f"工作表 {old!r} 不存在")
                if not new.strip() or re.search(r"[:\\/?*\[\]]", new) or len(new) > 31:
                    raise ValueError(f"新工作表名 {new!r} 含非法字符或超过 31 个字符")
                if old == new:
                    continue
                if any(name.lower() == new.lower() for name in current_names if name != old):
                    raise ValueError(f"新工作表名 {new!r} 与现有工作表重名")
                workbook.Worksheets(old).Name = new
                current_names.remove(old)
                current_names.add(new)
                renames[old] = new
            if not renames:
                raise ValueError("改名计划中没有实际改名")

            stage = "保存改名候选"
            workbook.Save()
            workbook.Close(SaveChanges=False)
            workbook = None
            stage = "重新打开并校验改名候选"
            reopened = self._open_workbook(candidate, read_only=True, label="改名候选工作簿")
            after = _snapshot(reopened)
            expected_names = [renames.get(name, name) for name in before_names]
            if [sheet["name"] for sheet in after] != expected_names:
                raise ValueError("改名后工作表名称或顺序与计划不一致")
            formula_count = 0
            value_count = 0
            for old_sheet, new_sheet in zip(before, after):
                old_formulas = old_sheet["formulas"]
                new_formulas = new_sheet["formulas"]
                if set(old_formulas) != set(new_formulas):
                    raise ValueError(f"工作表 {new_sheet['name']!r} 的公式单元格有增删")
                for address, formula in old_formulas.items():
                    expected = "=" + _rewrite_sheet_refs(formula[1:], renames)
                    actual = new_formulas[address]
                    if expected.strip() != actual.strip():
                        raise ValueError(
                            f"工作表 {new_sheet['name']!r} 公式 {_cell_name(address)} "
                            f"引用改写不一致：{expected!r} -> {actual!r}"
                        )
                old_values = old_sheet["values"]
                new_values = new_sheet["values"]
                if set(old_values) != set(new_values):
                    raise ValueError(f"工作表 {new_sheet['name']!r} 的非公式数据单元格有增删")
                for address, value in old_values.items():
                    if not _values_equal(value, new_values[address]):
                        raise ValueError(
                            f"工作表 {new_sheet['name']!r} 非公式单元格 {_cell_name(address)} "
                            f"值发生变化：{value!r} -> {new_values[address]!r}"
                        )
                formula_count += len(old_formulas)
                value_count += len(old_values)
            return {
                "ok": True,
                "application": self._application_name,
                "process_id": self._process_id,
                "validation": {
                    "verified": True,
                    "level": "excel_com_sheet_order_formula_text_and_cell_values",
                    "worksheet_count": len(after),
                    "formula_cells": formula_count,
                    "value_cells": value_count,
                },
            }
        except Exception as error:
            for opened in (reopened, workbook):
                if opened is not None:
                    try:
                        opened.Close(SaveChanges=False)
                    except Exception:
                        pass
            return {
                "ok": False,
                "stage": stage,
                "error": f"{type(error).__name__}: {error}",
                "traceback": traceback.format_exc(limit=6),
            }

    def _open_validation_pair(
        self, template_path: Path, source_copy_path: Path
    ) -> tuple[Any, Any]:
        """Open the template (read-only) and the writable copy in one host.

        Excel may reject either Open with the same opaque 1004 error when the
        automation instance is wedged, so on failure both workbooks are
        discarded together and the instance is rebuilt before one retry -
        reusing the template handle from a reset instance would be invalid.
        """
        first_error = None
        for attempt in (1, 2):
            assert self._application is not None
            template_workbook = None
            try:
                template_workbook = self._application.Workbooks.Open(
                    str(template_path), 0, True, AddToMru=False
                )
                source_workbook = self._application.Workbooks.Open(
                    str(source_copy_path), 0, False, AddToMru=False
                )
                return template_workbook, source_workbook
            except Exception as error:
                if template_workbook is not None:
                    try:
                        template_workbook.Close(SaveChanges=False)
                    except Exception:
                        pass
                if attempt == 2:
                    diagnostics = self._host_diagnostics(first_error, error)
                    raise RuntimeError(
                        f"无法打开校验模板或工作簿副本：\n"
                        f"模板：{template_path}\n副本：{source_copy_path}\n"
                        f"第一次错误：{type(first_error).__name__}: {first_error}\n"
                        f"重建 Excel/WPS 实例后重试仍失败：{type(error).__name__}: {error}\n"
                        f"{diagnostics}"
                    ) from error
                first_error = error
                self._reset_application()
                self._ensure_application()
        raise AssertionError("unreachable")

    def _host_diagnostics(self, first_error: Any, second_error: Any) -> str:
        """State of the COM host at the moment both Open attempts failed.

        `Application.Ready` distinguishes "Excel is busy/blocked by a dialog"
        from "Excel rejected this file", which is otherwise indistinguishable
        from the single generic 1004 message.
        """
        parts = [f"宿主应用：{self._application_name or '未知'}"]
        process_id = self._process_id
        parts.append(f"宿主进程 PID：{process_id if process_id is not None else '未取得'}")
        if process_id is not None:
            parts.append(f"进程仍在运行：{'是' if _process_is_alive(process_id) else '否'}")
        application = self._application
        if application is not None:
            try:
                parts.append(f"Application.Ready：{bool(application.Ready)}")
            except Exception as error:
                parts.append(f"Application.Ready：读取失败（{type(error).__name__}）")
            try:
                parts.append(f"已打开的文档数：{int(application.Workbooks.Count)}")
            except Exception:
                parts.append("已打开的文档数：读取失败")
        same_error = type(first_error) is type(second_error) and str(first_error) == str(second_error)
        parts.append(
            "两次错误完全相同（更像文件或宿主状态问题，而非偶发）"
            if same_error
            else "两次错误不同（更像宿主实例的偶发状态）"
        )
        return "；".join(parts)

    def recalculate_named_validation(
        self,
        source_copy_text: str,
        regions: list[dict[str, str]],
        template_text: str | None = None,
    ) -> dict[str, Any]:
        """Recalculate validation formulas in one working copy.

        The template is normally NOT opened here: Rust copies its formulas into
        the working copy over OOXML (`ooxml_patch::copy_cell_region`) so this
        host only ever opens a single workbook. Passing `template_text` keeps
        the older two-workbook behaviour available for callers that still need
        it.
        """
        source_copy_path = _absolute(source_copy_text)
        template_path = _absolute(template_text) if template_text else None
        stage = "检查校验区域临时副本"
        template_workbook: Any | None = None
        source_workbook: Any | None = None
        staging: tempfile.TemporaryDirectory[str] | None = None
        try:
            if not source_copy_path.is_file():
                raise FileNotFoundError(f"校验工作簿副本不存在：{source_copy_path}")
            if template_path is not None and not template_path.is_file():
                raise FileNotFoundError(f"校验区域模板不存在：{template_path}")
            if not regions:
                raise ValueError("没有可复制的校验区域")
            # Check the container locally first: Excel reports a damaged
            # package as an unhelpful "Workbooks.Open invalid" error.
            stage = "检查校验工作簿副本文件"
            _validate_container(source_copy_path, "校验工作簿副本")
            if template_path is not None:
                _validate_container(template_path, "校验区域模板")

            # Never hand Excel a path the user may have open. A workbook that
            # is already open (or that would trigger a "file in use" prompt)
            # makes `Workbooks.Open` fail with the same opaque 1004 error even
            # though the file is perfectly valid, because DisplayAlerts is off.
            # Working on a private staged copy removes that whole failure class.
            stage = "为 COM 准备私有临时副本"
            staging = tempfile.TemporaryDirectory(prefix="pytools-validation-")
            staged_copy = Path(staging.name) / f"copy{source_copy_path.suffix}"
            shutil.copy2(source_copy_path, staged_copy)

            self._ensure_application()
            application = self._application
            if application is None:
                raise RuntimeError("Excel/WPS COM 尚未启动")

            if template_path is None:
                stage = "通过 COM 打开工作簿副本"
                source_workbook = self._open_workbook(
                    staged_copy, read_only=False, label="校验工作簿副本"
                )
            else:
                staged_template = (
                    Path(staging.name) / f"template{template_path.suffix}"
                )
                shutil.copy2(template_path, staged_template)
                stage = "通过 COM 打开校验模板和工作簿副本"
                template_workbook, source_workbook = self._open_validation_pair(
                    staged_template, staged_copy
                )

                stage = "复制模板校验公式到工作簿副本"
                for region in regions:
                    template_sheet = str(region["template_sheet"])
                    source_sheet = str(region["source_sheet"])
                    address = str(region["address"])
                    source_range = template_workbook.Worksheets(template_sheet).Range(address)
                    target_range = source_workbook.Worksheets(source_sheet).Range(address)
                    source_range.Copy(Destination=target_range)
                application.CutCopyMode = False

            stage = "调用 Excel/WPS 完整重算校验公式"
            # 1-4 always copies formula cells in, so a failed recalculation
            # would report stale/blank results as if they were computed.
            recalc_method, recalc_errors = _force_recalc(application)
            if recalc_method is None:
                raise RuntimeError(
                    "校验区域公式无法重算，已取消提取：" + "；".join(recalc_errors)
                )
            source_workbook.Save()

            stage = "读取校验区域公式计算结果"
            cells: list[dict[str, str]] = []
            for region in regions:
                sheet_name = str(region["source_sheet"])
                address = str(region["address"])
                worksheet = source_workbook.Worksheets(sheet_name)
                selected = worksheet.Range(address)
                row_count = int(selected.Rows.Count)
                column_count = int(selected.Columns.Count)
                first_row = int(selected.Row)
                first_column = int(selected.Column)
                formulas = _matrix(selected.Formula, row_count, column_count)
                values = _matrix(selected.Value2, row_count, column_count)
                for row_offset in range(row_count):
                    for column_offset in range(column_count):
                        formula = formulas[row_offset][column_offset]
                        value = values[row_offset][column_offset]
                        if not (isinstance(formula, str) and formula.startswith("=")):
                            continue
                        if not isinstance(value, str):
                            continue
                        if not any(keyword in value for keyword in ("错", "硬性", "软性", "校验失败")):
                            continue
                        cells.append(
                            {
                                "sheet": sheet_name,
                                "address": _cell_name(
                                    (first_row + row_offset, first_column + column_offset)
                                ),
                                "value": value,
                            }
                        )

            source_workbook.Close(SaveChanges=False)
            source_workbook = None
            if template_workbook is not None:
                template_workbook.Close(SaveChanges=False)
                template_workbook = None

            # Publish the recalculated workbook back to the retained 副本 path
            # so the file the user inspects is exactly what was calculated.
            stage = "回写已重算的校验副本"
            shutil.copy2(staged_copy, source_copy_path)
            return {
                "ok": True,
                "application": self._application_name,
                "process_id": self._process_id,
                "recalc_method": recalc_method,
                "cells": cells,
            }
        except Exception as error:
            for opened in (source_workbook, template_workbook):
                if opened is not None:
                    try:
                        opened.Close(SaveChanges=False)
                    except Exception:
                        pass
            self._reset_application()
            return {
                "ok": False,
                "stage": stage,
                "error": f"{type(error).__name__}: {error}",
                "traceback": traceback.format_exc(limit=6),
            }
        finally:
            if staging is not None:
                staging.cleanup()

    def convert_word(self, source_text: str, output_text: str) -> dict[str, Any]:
        """Convert a Word/WPS Writer document to .docx through Word COM."""
        source = _absolute(source_text)
        output = _absolute(output_text)
        stage = "检查转换路径"
        document: Any | None = None
        try:
            if not source.is_file():
                raise FileNotFoundError(f"输入文件不存在：{source}")
            if output.exists():
                raise FileExistsError(f"拒绝覆盖已存在的候选文件：{output}")
            # `.wps` / `.wpt` are WPS Writer formats; Word cannot open them.
            self._ensure_word_application(
                prefer_wps=source.suffix.lower() in WPS_WORD_SUFFIXES
            )
            application = self._word_application
            if application is None:
                raise RuntimeError("Word/WPS Writer COM 尚未启动")

            stage = "读取源文档"
            document = application.Documents.Open(str(source), False, True)
            stage = "保存为 DOCX 候选"
            try:
                document.SaveAs2(str(output), DOCX_FILE_FORMAT)
            except Exception:
                # Older WPS Writer builds lack SaveAs2.
                document.SaveAs(str(output), DOCX_FILE_FORMAT)
            document.Close(False)
            document = None
            stage = "校验 DOCX 候选文件"
            if not output.is_file():
                raise RuntimeError(f"Word/WPS Writer 未生成候选文件：{output}")
            with output.open("rb") as candidate:
                if candidate.read(4) != b"PK\x03\x04":
                    raise ValueError(
                        "Word/WPS Writer 没有写出 OOXML 包（候选文件不是 ZIP），"
                        "请确认本机 Word 或 WPS 文字支持 DOCX 另存"
                    )
            return {
                "ok": True,
                "application": self._word_application_name,
                "process_id": self._word_process_id,
                "validation": {
                    "level": "docx_package_and_text_content",
                    "source_extension": source.suffix.lower().lstrip("."),
                },
            }
        except Exception as error:
            if document is not None:
                try:
                    document.Close(False)
                except Exception:
                    pass
            self._reset_word_application()
            return {
                "ok": False,
                "stage": stage,
                "error": f"{type(error).__name__}: {error}",
                "traceback": traceback.format_exc(limit=6),
            }

    def extract_word_text(self, input_text: str) -> dict[str, Any]:
        """Plain text of a document, for fidelity validation after conversion."""
        source = _absolute(input_text)
        stage = "检查文档路径"
        document: Any | None = None
        try:
            if not source.is_file():
                raise FileNotFoundError(f"文档不存在：{source}")
            self._ensure_word_application(
                prefer_wps=source.suffix.lower() in WPS_WORD_SUFFIXES
            )
            application = self._word_application
            if application is None:
                raise RuntimeError("Word/WPS Writer COM 尚未启动")

            stage = "读取文档正文"
            document = application.Documents.Open(str(source), False, True)
            text = document.Content.Text
            if not isinstance(text, str):
                text = "" if text is None else str(text)
            document.Close(False)
            document = None
            return {
                "ok": True,
                "application": self._word_application_name,
                "process_id": self._word_process_id,
                "text": text,
            }
        except Exception as error:
            if document is not None:
                try:
                    document.Close(False)
                except Exception:
                    pass
            self._reset_word_application()
            return {
                "ok": False,
                "stage": stage,
                "error": f"{type(error).__name__}: {error}",
                "traceback": traceback.format_exc(limit=6),
            }

    def close(self) -> None:
        self._reset_application()
        self._reset_word_application()
        try:
            self._pythoncom.CoUninitialize()
        except Exception:
            pass


def _needs_quote(name: str) -> bool:
    return bool(re.search(r"[\s:\/?*\[\]]", name)) or name[:1].isdigit()


_QUOTED_SHEET_REF = re.compile(r"'((?:[^']|'')*)'!")
_BARE_SHEET_REF = re.compile(r"([^\W\d][\w.]*)!")


def _split_string_literals(text: str):
    """把公式切成「字符串字面量 / 其他」交替片段。

    公式里的双引号文本（含 "" 转义）不是引用，改写必须跳过。
    """
    segments = []
    start = 0
    i = 0
    length = len(text)
    while i < length:
        if text[i] != '"':
            i += 1
            continue
        segments.append((False, text[start:i]))
        j = i + 1
        while j < length:
            if text[j] == '"':
                if text[j + 1 : j + 2] == '"':
                    j += 2
                    continue
                break
            j += 1
        segments.append((True, text[i : j + 1]))
        i = j + 1
        start = i
    segments.append((False, text[start:]))
    return segments


def _apply_sheet_rename(text: str, renames: dict) -> str:
    def quoted(match):
        name = match.group(1).replace("''", "'")
        new = renames.get(name)
        if new is None:
            return match.group(0)
        rendered = ("'" + new.replace("'", "''") + "'") if _needs_quote(new) else new
        return rendered + "!"

    def bare(match):
        # [1]Sheet! 是外部工作簿引用，不属于当前工作簿改名范围。
        if match.start() > 0 and text[match.start() - 1] == "]":
            return match.group(0)
        name = match.group(1)
        new = renames.get(name)
        if new is None:
            return match.group(0)
        rendered = ("'" + new.replace("'", "''") + "'") if _needs_quote(new) else new
        return rendered + "!"

    text = _QUOTED_SHEET_REF.sub(quoted, text)
    return _BARE_SHEET_REF.sub(bare, text)


def _rewrite_sheet_refs(text: str, renames: dict) -> str:
    """改写公式/命名区域引用中的旧表名（引号、裸名、Unicode 表名）。"""
    if not renames:
        return text
    rewritten = [
        piece
        if is_literal
        else _apply_sheet_rename(piece, renames)
        for is_literal, piece in _split_string_literals(text)
    ]
    return "".join(rewritten)


def _openpyxl_snapshot(workbook):
    sheets = []
    for ws in workbook.worksheets:
        formulas = {}
        values = {}
        for row in ws.iter_rows():
            for cell in row:
                value = cell.value
                if value is None:
                    continue
                address = cell.coordinate
                if isinstance(value, str) and value.startswith("="):
                    formulas[address] = value[1:]
                else:
                    values[address] = value
        sheets.append({"name": ws.title, "formulas": formulas, "values": values})
    return sheets


def _worksheet_parts_by_name(archive) -> dict[str, str]:
    import xml.etree.ElementTree as ET

    spreadsheet = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
    office_relationships = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
    package_relationships = "http://schemas.openxmlformats.org/package/2006/relationships"
    workbook = ET.fromstring(archive.read("xl/workbook.xml"))
    relationships_root = ET.fromstring(archive.read("xl/_rels/workbook.xml.rels"))
    relationships = {
        relationship.attrib["Id"]: relationship.attrib
        for relationship in relationships_root.findall(
            f"{{{package_relationships}}}Relationship"
        )
    }
    parts: dict[str, str] = {}
    for sheet in workbook.findall(f"{{{spreadsheet}}}sheets/{{{spreadsheet}}}sheet"):
        relationship_id = sheet.attrib[f"{{{office_relationships}}}id"]
        relationship = relationships.get(relationship_id)
        if relationship is None or not relationship["Type"].endswith("/worksheet"):
            continue
        target = relationship["Target"]
        part = target.lstrip("/") if target.startswith("/") else f"xl/{target}"
        parts[sheet.attrib["name"]] = part
    return parts


def _replace_cell_xml(xml_data: bytes, address: str, transform) -> bytes:
    coordinate = re.escape(address.encode("ascii"))
    pattern = re.compile(
        rb"<c\b(?=[^>]*\br=(?:\""
        + coordinate
        + rb"\"|'"
        + coordinate
        + rb"'))[^>]*?(?:/>|>.*?</c>)",
        re.DOTALL,
    )
    matches = list(pattern.finditer(xml_data))
    if len(matches) != 1:
        raise ValueError(f"无法唯一定位工作表单元格：{address}")
    match = matches[0]
    return xml_data[: match.start()] + transform(match.group(0)) + xml_data[match.end() :]


def _cell_opening_with_type(opening: bytes, cell_type: str | None) -> bytes:
    if cell_type is None:
        return re.sub(rb"\s+t=(?:\"[^\"]*\"|'[^']*')", b"", opening, count=1)
    if re.search(rb"\bt=(?:\"[^\"]*\"|'[^']*')", opening):
        return re.sub(
            rb"\bt=(?:\"[^\"]*\"|'[^']*')",
            f't="{cell_type}"'.encode("ascii"),
            opening,
            count=1,
        )
    return opening[:-1] + f' t="{cell_type}">'.encode("ascii")


def _formula_cache_snapshot(
    candidate: Path,
    before: list[dict[str, Any]],
    renames: dict[str, str],
) -> dict[tuple[str, str], tuple[str | None, bool, str | None]]:
    """Capture formula cached values before openpyxl rewrites the package."""
    import xml.etree.ElementTree as ET
    import zipfile

    namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
    caches: dict[tuple[str, str], tuple[str | None, bool, str | None]] = {}
    with zipfile.ZipFile(candidate, "r") as source_zip:
        worksheet_parts = _worksheet_parts_by_name(source_zip)
        for sheet in before:
            if not sheet["formulas"]:
                continue
            name = sheet["name"]
            part = worksheet_parts.get(name)
            if part is None:
                raise ValueError(f"改名前公式工作表部件缺失：{name}")
            root = ET.fromstring(source_zip.read(part))
            cells = {
                cell.attrib.get("r"): cell
                for cell in root.iter(f"{{{namespace}}}c")
            }
            final_name = renames.get(name, name)
            for address in sheet["formulas"]:
                cell = cells.get(address)
                formula = (
                    cell.find(f"{{{namespace}}}f") if cell is not None else None
                )
                if cell is None or formula is None:
                    raise ValueError(f"改名前公式单元格缺失：{name}!{address}")
                value = cell.find(f"{{{namespace}}}v")
                caches[(final_name, address)] = (
                    cell.attrib.get("t"),
                    value is not None,
                    value.text if value is not None else None,
                )
    return caches


def _restore_formula_caches(
    candidate: Path,
    caches: dict[tuple[str, str], tuple[str | None, bool, str | None]],
) -> None:
    """Restore formula results that openpyxl does not retain when saving."""
    if not caches:
        return

    import xml.etree.ElementTree as ET
    import zipfile
    from xml.sax.saxutils import escape

    namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
    by_sheet: dict[str, list[tuple[str, tuple[str | None, bool, str | None]]]] = {}
    for (sheet_name, address), cache in caches.items():
        by_sheet.setdefault(sheet_name, []).append((address, cache))

    patched_parts: dict[str, bytes] = {}
    with zipfile.ZipFile(candidate, "r") as source_zip:
        worksheet_parts = _worksheet_parts_by_name(source_zip)
        for sheet_name, cell_caches in by_sheet.items():
            part = worksheet_parts.get(sheet_name)
            if part is None:
                raise ValueError(f"改名后公式工作表部件缺失：{sheet_name}")
            xml_data = source_zip.read(part)
            root = ET.fromstring(xml_data)
            cells = {
                cell.attrib.get("r"): cell
                for cell in root.iter(f"{{{namespace}}}c")
            }
            for address, (cell_type, has_value, value_text) in cell_caches:
                cell = cells.get(address)
                if cell is None or cell.find(f"{{{namespace}}}f") is None:
                    raise ValueError(f"改名后公式单元格缺失：{sheet_name}!{address}")

                def restore(element: bytes) -> bytes:
                    opening_end = element.find(b">") + 1
                    if opening_end == 0:
                        raise ValueError(f"公式单元格 XML 不完整：{sheet_name}!{address}")
                    opening = _cell_opening_with_type(element[:opening_end], cell_type)
                    self_closing = opening.rstrip().endswith(b"/>")
                    if self_closing:
                        raise ValueError(f"改名后公式内容缺失：{sheet_name}!{address}")
                    content = element[opening_end:-4]
                    content = re.sub(
                        rb"<(?:v)\b[^>]*/>|<v\b[^>]*>.*?</v>",
                        b"",
                        content,
                        flags=re.DOTALL,
                    )
                    formula = re.search(
                        rb"<f\b[^>]*/>|<f\b[^>]*>.*?</f>",
                        content,
                        re.DOTALL,
                    )
                    if formula is None:
                        raise ValueError(
                            f"改名后公式 XML 缺失：{sheet_name}!{address}"
                        )
                    if has_value:
                        escaped_value = (
                            escape(value_text).encode("utf-8")
                            if value_text is not None
                            else b""
                        )
                        content = (
                            content[: formula.end()]
                            + b"<v>"
                            + escaped_value
                            + b"</v>"
                            + content[formula.end() :]
                        )
                    return opening + content + b"</c>"

                try:
                    xml_data = _replace_cell_xml(xml_data, address, restore)
                except ValueError as error:
                    raise ValueError(f"{sheet_name}!{address}：{error}") from error
            patched_parts[part] = xml_data

    _rewrite_zip_parts(candidate, patched_parts)


def _rewrite_zip_parts(candidate: Path, patched_parts: dict[str, bytes]) -> None:
    import zipfile

    with tempfile.NamedTemporaryFile(
        prefix=f".{candidate.name}.",
        suffix=".xlsx",
        dir=candidate.parent,
        delete=False,
    ) as temporary:
        temporary_path = Path(temporary.name)
    try:
        with zipfile.ZipFile(candidate, "r") as source_zip, zipfile.ZipFile(
            temporary_path, "w"
        ) as output_zip:
            for item in source_zip.infolist():
                data = patched_parts.get(item.filename)
                if data is None:
                    data = source_zip.read(item.filename)
                output_zip.writestr(item, data)
        temporary_path.replace(candidate)
    finally:
        if temporary_path.exists():
            temporary_path.unlink()


def _excel_link_sources(workbook: Any) -> list[str]:
    """Read Excel workbook-link sources as absolute paths before/after SaveAs."""
    try:
        sources = workbook.LinkSources(1)  # xlExcelLinks
    except Exception as error:
        raise RuntimeError(f"读取 Excel 外部工作簿链接失败：{error}") from error
    if sources is None:
        return []
    if isinstance(sources, (str, bytes)):
        return [str(sources)]
    return [str(source) for source in sources]


def _link_source_path(source: str) -> Path | None:
    """Convert a local Excel link source to a path, leaving URL links intact."""
    import re
    from urllib.parse import unquote, urlsplit

    source = str(source)
    if re.match(r"^[A-Za-z]:[\\/]", source) or source.startswith("\\\\"):
        return Path(source).resolve()

    parsed = urlsplit(source)
    if parsed.scheme.lower() != "file":
        return None if parsed.scheme else Path(source).resolve()

    path = unquote(parsed.path)
    if parsed.netloc:
        path = "\\\\" + parsed.netloc + path.replace("/", "\\")
    elif re.match(r"^/[A-Za-z]:/", path):
        path = path[1:].replace("/", "\\")
    return Path(path).resolve()


def _normalized_link_sources(sources: list[str]) -> list[str]:
    import os

    normalized = []
    for source in sources:
        path = _link_source_path(source)
        normalized.append(
            os.path.normcase(os.path.normpath(str(path))) if path is not None else source
        )
    return normalized


def _external_link_target(source: str, output_directory: Path) -> str:
    """Make an OOXML external-link target resolve to the original link source."""
    import os

    source_path = _link_source_path(source)
    if source_path is None:
        return source
    try:
        relative = os.path.relpath(str(source_path), str(output_directory.resolve()))
    except ValueError:
        # A different Windows volume cannot be represented by a relative URI.
        return source_path.as_uri()
    return relative.replace("\\", "/")


def _restore_external_link_targets(candidate: Path, source_links: list[str]) -> None:
    """Restore Excel's external-link relationships to their pre-SaveAs targets.

    Excel can rebase a relative external workbook link to the output directory
    during SaveAs. The formula text and cached values then look unchanged while
    the published workbook points at a different (often nonexistent) file.
    Keep the COM-calculated workbook intact and patch only externalLinkPath
    relationship targets, then let Excel reopen the package for verification.
    """
    if not source_links:
        return

    import re
    import xml.etree.ElementTree as ET
    import zipfile

    package_relationships = "http://schemas.openxmlformats.org/package/2006/relationships"
    office_relationships = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
    spreadsheet_main = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
    ET.register_namespace("", package_relationships)
    relationship_tag = f"{{{package_relationships}}}Relationship"
    external_book_tag = f"{{{spreadsheet_main}}}externalBook"
    link_parts: list[tuple[int, str]] = []
    with zipfile.ZipFile(candidate, "r") as archive:
        for name in archive.namelist():
            match = re.fullmatch(
                r"xl/externalLinks/_rels/externalLink(\d+)\.xml\.rels", name
            )
            if match:
                link_parts.append((int(match.group(1)), name))
        link_parts.sort()
        if len(link_parts) != len(source_links):
            raise ValueError(
                "另存候选中的外部工作簿链接数量与源工作簿不一致："
                f"源 {len(source_links)} 个，候选 {len(link_parts)} 个"
            )

        patched_parts: dict[str, bytes] = {}
        for (_, part), source in zip(link_parts, source_links):
            root = ET.fromstring(archive.read(part))
            external_link_part = part.replace("/_rels/", "/")[: -len(".rels")]
            external_link_root = ET.fromstring(archive.read(external_link_part))
            external_book = external_link_root.find(external_book_tag)
            relationship_id = (
                external_book.attrib.get(f"{{{office_relationships}}}id")
                if external_book is not None
                else None
            )
            relationship = next(
                (
                    relation
                    for relation in root.findall(relationship_tag)
                    if relation.attrib.get("Id") == relationship_id
                    and relation.attrib.get("Type", "").endswith("/externalLinkPath")
                ),
                None,
            )
            if relationship is None:
                raise ValueError(
                    f"外部链接关系部件 {part} 找不到 externalBook 引用的路径关系："
                    f"r:id={relationship_id!r}"
                )
            relationship.set(
                "Target", _external_link_target(source, candidate.parent)
            )
            relationship.set("TargetMode", "External")
            patched_parts[part] = ET.tostring(
                root, encoding="utf-8", xml_declaration=True
            )

    _rewrite_zip_parts(candidate, patched_parts)


def _restore_empty_string_cells(
    candidate: Path,
    before: list[dict[str, Any]],
    renames: dict[str, str],
) -> None:
    """Restore explicit empty strings that openpyxl serializes as blank cells.

    In a source XLSX an empty string may be represented by a shared-string
    entry whose text is empty. openpyxl writes that cell as an empty
    ``inlineStr`` element, which reloads as ``None`` and changes Excel
    semantics (for example, ``COUNTA`` / ``ISBLANK``). Reinsert the explicit
    empty inline string in the saved worksheet XML so the rename does not
    silently erase it.
    """
    import xml.etree.ElementTree as ET
    import zipfile

    namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
    empty_cells_by_name = {
        renames.get(sheet["name"], sheet["name"]): [
            address
            for address, value in sheet["values"].items()
            if isinstance(value, str) and value == ""
        ]
        for sheet in before
    }
    empty_cells_by_name = {
        name: cells for name, cells in empty_cells_by_name.items() if cells
    }
    if not empty_cells_by_name:
        return

    patched_parts: dict[str, bytes] = {}
    with zipfile.ZipFile(candidate, "r") as source_zip:
        worksheet_parts = _worksheet_parts_by_name(source_zip)

        empty_by_part = {
            worksheet_parts[name]: cells
            for name, cells in empty_cells_by_name.items()
            if name in worksheet_parts
        }
        missing_sheets = sorted(set(empty_cells_by_name) - set(worksheet_parts))
        if missing_sheets:
            raise ValueError(
                f"改名后包含显式空字符串的工作表部件缺失：{', '.join(missing_sheets)}"
            )
        for part, addresses in empty_by_part.items():
            try:
                xml_data = source_zip.read(part)
            except KeyError as error:
                raise ValueError(f"改名后工作表部件缺失：{part}") from error
            root = ET.fromstring(xml_data)
            cells = {cell.attrib.get("r"): cell for cell in root.iter(f"{{{namespace}}}c")}
            for address in addresses:
                cell = cells.get(address)
                if cell is None:
                    raise ValueError(f"改名后显式空字符串单元格缺失：{part}!{address}")
                if cell.find(f"{{{namespace}}}f") is not None:
                    raise ValueError(f"显式空字符串单元格意外变成公式：{part}!{address}")
                def restore(element: bytes) -> bytes:
                    opening_end = element.find(b">") + 1
                    if opening_end == 0:
                        raise ValueError(f"工作表单元格 XML 不完整：{part}!{address}")
                    opening = element[:opening_end]
                    self_closing = opening.rstrip().endswith(b"/>")
                    if self_closing:
                        opening = opening.rstrip()[:-2] + b">"
                    opening = _cell_opening_with_type(opening, "inlineStr")
                    if self_closing:
                        return opening + b"<is><t></t></is></c>"
                    content = element[opening_end:-4]
                    content = re.sub(
                        rb"<(?:v|is)\b[^>]*/>|<(?:v|is)\b[^>]*>.*?</(?:v|is)>",
                        b"",
                        content,
                        flags=re.DOTALL,
                    )
                    return opening + content + b"<is><t></t></is></c>"

                xml_data = _replace_cell_xml(xml_data, address, restore)
            patched_parts[part] = xml_data

    _rewrite_zip_parts(candidate, patched_parts)


def _count_ref_errors(sheets) -> int:
    total = 0
    for sheet in sheets:
        for formula in sheet["formulas"].values():
            total += formula.upper().count("#REF!") + formula.upper().count("#NAME?")
    return total


def _serve() -> int:
    try:
        session = ConversionSession()
    except Exception as error:
        _write_json({"ready": False, "error": f"COM worker 初始化失败：{error}"})
        return 2
    _write_json({"ready": True, "protocol": PROTOCOL_VERSION})
    try:
        for line in sys.stdin:
            try:
                request = json.loads(line)
                operation = request.get("op")
                if operation == "shutdown":
                    _write_json({"ok": True})
                    break
                if operation == "rename_sheets":
                    _write_json(session.rename_sheets(str(request["candidate"]), request["changes"]))
                    continue
                if operation == "rename_sheets_com":
                    _write_json(
                        session.rename_sheets_com(
                            str(request["candidate"]), request["changes"]
                        )
                    )
                    continue
                if operation == "recalculate_named_validation":
                    _write_json(
                        session.recalculate_named_validation(
                            str(request["source_copy"]),
                            request["regions"],
                            request.get("template"),
                        )
                    )
                    continue
                if operation == "village_split_com":
                    _write_json(
                        session.village_split_com(
                            str(request["current"]),
                            str(request["previous"]),
                            str(request["candidate"]),
                        )
                    )
                    continue
                if operation == "institution_normalize_com":
                    _write_json(
                        session.institution_normalize_com(
                            str(request["candidate"]), request["updates"]
                        )
                    )
                    continue
                if operation == "rural_loan_com":
                    _write_json(
                        session.rural_loan_com(
                            str(request["current_candidate"]),
                            str(request["previous"]),
                        )
                    )
                    continue
                if operation == "duplicate_mark_com":
                    _write_json(session.duplicate_mark_com(
                        str(request["candidate"]), request["marks"]))
                    continue
                if operation == "delete_rows_com":
                    _write_json(session.delete_rows_com(
                        str(request["candidate"]), request["deletions"]))
                    continue
                if operation == "fx_header_com":
                    _write_json(session.fx_header_com(str(request["candidate"])))
                    continue
                if operation == "append_rows_com":
                    _write_json(session.append_rows_com(
                        str(request["candidate"]),
                        str(request["sheet"]),
                        int(request["start_row"]),
                        request["rows"],
                    ))
                    continue
                if operation == "convert_word":
                    _write_json(
                        session.convert_word(
                            str(request["input"]), str(request["output"])
                        )
                    )
                    continue
                if operation == "extract_word_text":
                    _write_json(session.extract_word_text(str(request["input"])))
                    continue
                if operation != "convert":
                    _write_json({"ok": False, "error": f"未知操作：{operation!r}"})
                    continue
                _write_json(session.convert(str(request["input"]), str(request["output"])))
            except Exception as error:
                _write_json({"ok": False, "error": f"请求处理失败：{type(error).__name__}: {error}"})
    finally:
        session.close()
    return 0


def main() -> int:
    sys.stdin.reconfigure(encoding="utf-8", errors="strict")
    sys.stdout.reconfigure(encoding="utf-8", errors="strict", line_buffering=True)
    if "--version" in sys.argv:
        try:
            version = importlib.metadata.version("pywin32")
            import pythoncom  # noqa: F401
            import win32com.client  # noqa: F401

            print(f"protocol={PROTOCOL_VERSION} pywin32={version}")
            return 0
        except Exception as error:
            print(f"protocol={PROTOCOL_VERSION} pywin32=missing error={error}")
            return 2
    if "--serve" in sys.argv:
        return _serve()
    print("usage: pywin32_excel_conversion_worker.py --version|--serve", file=sys.stderr)
    return 2


if __name__ == "__main__":
    raise SystemExit(main())
