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
import re
import shutil
import sys
import tempfile
import traceback
from pathlib import Path
from typing import Any

PROTOCOL_VERSION = 2
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


def _snapshot(workbook: Any) -> list[dict[str, Any]]:
    sheets: list[dict[str, Any]] = []
    for index in range(1, int(workbook.Worksheets.Count) + 1):
        worksheet = workbook.Worksheets(index)
        used = worksheet.UsedRange
        rows = int(used.Rows.Count)
        columns = int(used.Columns.Count)
        first_row = int(used.Row)
        first_column = int(used.Column)
        formulas = _matrix(used.Formula, rows, columns)
        values = _matrix(used.Value2, rows, columns)
        formula_cells: dict[tuple[int, int], str] = {}
        value_cells: dict[tuple[int, int], Any] = {}

        for row_offset in range(rows):
            for column_offset in range(columns):
                address = (first_row + row_offset, first_column + column_offset)
                formula = formulas[row_offset][column_offset]
                value = values[row_offset][column_offset]
                if isinstance(formula, str) and formula.startswith("="):
                    formula_cells[address] = formula
                elif value is not None:
                    value_cells[address] = value

        sheets.append(
            {
                "name": str(worksheet.Name),
                "formulas": formula_cells,
                "values": value_cells,
                "used_cells": rows * columns,
            }
        )
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
            ("formulas", "公式", lambda a, b: a.strip() == b.strip()),
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

    def _ensure_application(self, prefer_wps: bool = False) -> None:
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
            # `.et` / `.ett` are CFB containers: a rename to `.xls` would still
            # open with a plausible FileFormat, so require a WPS-native format
            # before publishing anything.
            if suffix in WPS_SPREADSHEET_SUFFIXES:
                stage = "校验源文件确为原生 ET/ETT 工作簿"
                source_file_format = int(workbook.FileFormat)
                if source_file_format < WPS_NATIVE_FILE_FORMAT_MIN:
                    raise ValueError(
                        f"Input has {suffix} extension but the application opened "
                        f"FileFormat {source_file_format}; expected a WPS-native "
                        f"FileFormat >= {WPS_NATIVE_FILE_FORMAT_MIN} "
                        f"(native ET is {ET_FILE_FORMAT})"
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

            stage = "重新打开候选工作簿"
            reopened = application.Workbooks.Open(str(output), 0, True)
            stage = "读取候选工作簿并验证公式和值"
            output_snapshot = _snapshot(reopened)
            validation = _compare(source_snapshot, output_snapshot)
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
                validation["source_extension"] = suffix.lstrip(".")
                validation["source_file_format"] = source_file_format
                validation["source_format_verified"] = True
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
            workbook = None
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
                    if result["values"].get(address) != value:
                        raise ValueError(
                            f"工作表 {source['name']!r} 非公式单元格 {address} 发生变化"
                        )
            if _count_ref_errors(after) > _count_ref_errors(before):
                raise ValueError("改名后出现新的 #REF! 公式引用")
            validation = {
                "verified": True,
                "level": "openpyxl_rename_with_reference_rewrite",
                "renamed_sheets": len(renames),
            }
            return {"ok": True, "validation": validation}
        except Exception as error:
            return {"ok": False, "stage": stage, "error": f"{type(error).__name__}: {error}"}

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
                if operation == "recalculate_named_validation":
                    _write_json(
                        session.recalculate_named_validation(
                            str(request["source_copy"]),
                            request["regions"],
                            request.get("template"),
                        )
                    )
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
