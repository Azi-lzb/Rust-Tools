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
import sys
import traceback
from pathlib import Path
from typing import Any

PROTOCOL_VERSION = 1
XLSX_FILE_FORMAT = 51  # xlOpenXMLWorkbook


def _write_json(value: dict[str, Any]) -> None:
    sys.stdout.write(json.dumps(value, ensure_ascii=False, separators=(",", ":")) + "\n")
    sys.stdout.flush()


def _matrix(value: Any, rows: int, columns: int) -> list[list[Any]]:
    if rows == 1 and columns == 1:
        return [[value]]
    if rows == 1:
        return [list(value)]
    if columns == 1:
        return [[item] for item in value]
    return [list(row) for row in value]


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

    def _ensure_application(self) -> None:
        if self._application is not None:
            return
        import win32com.client

        errors: list[str] = []
        for prog_id in ("Excel.Application", "KET.Application"):
            try:
                application = win32com.client.DispatchEx(prog_id)
            except Exception as error:
                errors.append(f"{prog_id}: {error}")
                continue
            try:
                open_workbooks = int(application.Workbooks.Count)
            except Exception as error:
                # We cannot prove this is an empty instance, so release only
                # our COM reference; never call Quit on an unknown session.
                errors.append(f"{prog_id}: 无法确认实例中的工作簿数量：{error}")
                continue
            if open_workbooks != 0:
                errors.append(
                    f"{prog_id}: 返回了含 {open_workbooks} 个已打开工作簿的实例；为保护用户数据已跳过"
                )
                continue
            try:
                application.Visible = False
                application.DisplayAlerts = False
                application.ScreenUpdating = False
                application.EnableEvents = False
                try:
                    application.AutomationSecurity = 3  # msoAutomationSecurityForceDisable
                except Exception:
                    pass
                self._application = application
                self._application_name = prog_id
                try:
                    import win32process

                    self._process_id = int(win32process.GetWindowThreadProcessId(int(application.Hwnd))[1])
                except Exception:
                    self._process_id = None
                return
            except Exception as error:
                try:
                    application.Quit()
                except Exception:
                    pass
                errors.append(f"{prog_id}: 初始化空 COM 实例失败：{error}")
        raise RuntimeError("无法启动 Excel/WPS COM：" + "；".join(errors))

    def _reset_application(self) -> None:
        application = self._application
        self._application = None
        self._application_name = None
        self._process_id = None
        if application is not None:
            try:
                application.Quit()
            except Exception:
                pass

    def convert(self, source_text: str, output_text: str) -> dict[str, Any]:
        source = Path(source_text)
        output = Path(output_text)
        stage = "检查转换路径"
        workbook: Any | None = None
        reopened: Any | None = None
        try:
            if not source.is_file():
                raise FileNotFoundError(f"输入文件不存在：{source}")
            if output.exists():
                raise FileExistsError(f"拒绝覆盖已存在的候选文件：{output}")
            self._ensure_application()
            application = self._application
            if application is None:
                raise RuntimeError("Excel/WPS COM 尚未启动")

            stage = "读取源工作簿并记录公式和值"
            workbook = application.Workbooks.Open(str(source), 0, True)
            stage = "调用 Excel/WPS 完整重算工作簿"
            application.CalculateFullRebuild()
            stage = "读取完整重算后的源工作簿公式和值"
            source_snapshot = _snapshot(workbook)
            stage = "通过 pywin32 写入 XLSX 候选"
            workbook.SaveAs(str(output), XLSX_FILE_FORMAT)
            workbook.Close(SaveChanges=False)
            workbook = None

            stage = "重新打开候选工作簿"
            reopened = application.Workbooks.Open(str(output), 0, True)
            stage = "读取候选工作簿并验证公式和值"
            output_snapshot = _snapshot(reopened)
            validation = _compare(source_snapshot, output_snapshot)
            validation["calculation"] = "CalculateFullRebuild before SaveAs"
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

    def recalculate_named_validation(
        self,
        template_text: str,
        source_copy_text: str,
        regions: list[dict[str, str]],
    ) -> dict[str, Any]:
        template_path = Path(template_text)
        source_copy_path = Path(source_copy_text)
        stage = "检查校验区域临时副本"
        template_workbook: Any | None = None
        source_workbook: Any | None = None
        try:
            if not template_path.is_file():
                raise FileNotFoundError(f"校验区域模板不存在：{template_path}")
            if not source_copy_path.is_file():
                raise FileNotFoundError(f"校验区域工作簿副本不存在：{source_copy_path}")
            if not regions:
                raise ValueError("没有可复制的校验区域")
            self._ensure_application()
            application = self._application
            if application is None:
                raise RuntimeError("Excel/WPS COM 尚未启动")

            stage = "通过 COM 打开校验模板和工作簿副本"
            template_workbook = application.Workbooks.Open(
                str(template_path), 0, True, AddToMru=False
            )
            source_workbook = application.Workbooks.Open(
                str(source_copy_path), 0, False, AddToMru=False
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
            application.CalculateFullRebuild()
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
            template_workbook.Close(SaveChanges=False)
            template_workbook = None
            return {
                "ok": True,
                "application": self._application_name,
                "process_id": self._process_id,
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

    def close(self) -> None:
        self._reset_application()
        try:
            self._pythoncom.CoUninitialize()
        except Exception:
            pass


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
                if operation == "recalculate_named_validation":
                    _write_json(
                        session.recalculate_named_validation(
                            str(request["template"]),
                            str(request["source_copy"]),
                            request["regions"],
                        )
                    )
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
