"""Read successful real-workbook candidates through Excel COM without saving."""
import json
import math
from pathlib import Path

import openpyxl
import win32com.client
from openpyxl.utils.datetime import to_excel


ROOT = Path(__file__).resolve().parents[3]
EVIDENCE = ROOT / "artifacts" / "excel-edit-phase2"


def records(name):
    data = json.loads((EVIDENCE / name).read_text(encoding="utf-8"))
    return data.get("runs", []) if isinstance(data, dict) else data


def successful_outputs():
    selected = []
    for run in records("dedup-results.json"):
        if run.get("business_correct"):
            selected.append((run["menu"], run["mode"], run["output_copy"]))
    for run in records("remaining-results.json"):
        if run.get("menu") in {"5-2", "5-3"} and run.get("com"):
            if any(item.get("status") == "opened_read_only" for item in run["com"].values()):
                path = next(iter(run["com"]))
                selected.append((run["menu"], run["mode"], path))
    for run in records("special-results.json"):
        if run.get("menu") in {"5-1", "5-7"} and run.get("changed"):
            selected.append((run["menu"], run["mode"], run["output_copy"]))
    for run in records("special-results-final.json"):
        if run.get("menu") == "5-4" and run.get("changed"):
            selected.append((run["menu"], run["mode"], run["output_copy"]))
    for run in records("import-results.json"):
        if run.get("menu") == "5-6" and run.get("changed"):
            selected.append((run["menu"], run["mode"], run["target"]))
    unique = {}
    for menu, mode, path in selected:
        unique[(menu, mode, str(Path(path).resolve()))] = (menu, mode, Path(path))
    return list(unique.values())


def matrix(value, rows, columns):
    if rows == columns == 1:
        return ((value,),)
    if rows == 1:
        if len(value) == 1 and isinstance(value[0], (tuple, list)):
            return (tuple(value[0]),)
        return (tuple(value),)
    if columns == 1:
        return tuple(
            (row[0] if isinstance(row, (tuple, list)) else row,) for row in value
        )
    return tuple(tuple(row) for row in value)


def formula_text(value):
    if isinstance(value, str) and value.startswith("="):
        return value
    return getattr(value, "text", None)


def excel_matches(expected, actual, epoch):
    if isinstance(expected, bool):
        return actual is expected
    if isinstance(expected, (int, float)) and not isinstance(expected, bool):
        return isinstance(actual, (int, float)) and math.isclose(
            float(expected), float(actual), rel_tol=0.0, abs_tol=1e-8
        )
    if hasattr(expected, "isoformat"):
        return isinstance(actual, (int, float)) and math.isclose(
            to_excel(expected, epoch), float(actual), rel_tol=0.0, abs_tol=1e-8
        )
    return expected == actual


def verify_with_com(path):
    workbook = openpyxl.load_workbook(path, data_only=False, read_only=False)
    excel = win32com.client.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    excel.AskToUpdateLinks = False
    excel.AutomationSecurity = 3
    try:
        try:
            excel.Calculation = -4135  # xlCalculationManual
        except Exception:
            pass
        com_book = excel.Workbooks.Open(
            Filename=str(path),
            UpdateLinks=0,
            ReadOnly=True,
            IgnoreReadOnlyRecommended=True,
            AddToMru=False,
        )
        expected_sheets = {sheet.title: sheet for sheet in workbook.worksheets}
        actual_sheets = {com_book.Worksheets(index).Name: com_book.Worksheets(index)
                         for index in range(1, com_book.Worksheets.Count + 1)}
        mismatches = []
        checked_values = 0
        checked_formulas = 0
        for sheet_name, sheet in expected_sheets.items():
            if sheet_name not in actual_sheets:
                mismatches.append({"sheet": sheet_name, "kind": "missing_sheet"})
                continue
            if not sheet._cells:
                continue
            max_row = max((cell.row for cell in sheet._cells.values() if cell.value is not None), default=0)
            max_col = max((cell.column for cell in sheet._cells.values() if cell.value is not None), default=0)
            if max_row == 0 or max_col == 0:
                continue
            com_sheet = actual_sheets[sheet_name]
            cell_range = com_sheet.Range(com_sheet.Cells(1, 1), com_sheet.Cells(max_row, max_col))
            values = matrix(cell_range.Value2, max_row, max_col)
            formulas = matrix(cell_range.Formula, max_row, max_col)
            for cell in sheet._cells.values():
                expected = cell.value
                if expected is None:
                    continue
                row, col = cell.row - 1, cell.column - 1
                if cell.data_type == "f" or formula_text(expected):
                    checked_formulas += 1
                    wanted = formula_text(expected)
                    actual = formulas[row][col]
                    if wanted and actual != wanted:
                        mismatches.append({"sheet": sheet_name, "cell": cell.coordinate,
                                           "kind": "formula", "expected": wanted,
                                           "actual": actual})
                else:
                    checked_values += 1
                    actual = values[row][col]
                    if not excel_matches(expected, actual, workbook.epoch):
                        mismatches.append({"sheet": sheet_name, "cell": cell.coordinate,
                                           "kind": "value", "expected": str(expected),
                                           "actual": str(actual)})
        result = {
            "status": "passed" if not mismatches else "failed",
            "com_opened_read_only": bool(com_book.ReadOnly),
            "saved": False,
            "sheet_count": com_book.Worksheets.Count,
            "checked_constant_cells": checked_values,
            "checked_formula_cells": checked_formulas,
            "mismatch_count": len(mismatches),
            "mismatches": mismatches[:50],
        }
        com_book.Close(SaveChanges=False)
        return result
    finally:
        workbook.close()
        excel.Quit()


def main():
    results = []
    for menu, mode, path in successful_outputs():
        try:
            result = verify_with_com(path)
        except Exception as error:
            result = {"status": "error", "error": repr(error), "saved": False}
        row = {"menu": menu, "mode": mode, "output": str(path), **result}
        results.append(row)
        print(menu, mode, result["status"],
              "values", result.get("checked_constant_cells", 0),
              "formulas", result.get("checked_formula_cells", 0),
              "mismatches", result.get("mismatch_count", "n/a"), flush=True)
    (EVIDENCE / "com-readback-results.json").write_text(
        json.dumps(results, ensure_ascii=False, indent=2), encoding="utf-8"
    )
    if any(row["status"] != "passed" for row in results):
        raise SystemExit(1)


if __name__ == "__main__":
    main()
