# -*- coding: utf-8 -*-
"""Office compatibility smoke via Excel COM (pywin32).

Opens each candidate output in real Excel, checks it loads without repair
errors, worksheet count matches expectation, then closes WITHOUT saving.
Writes JSONL results. Run with the bundled python runtime (has pywin32).
"""
import json
import sys
from pathlib import Path

import openpyxl
import win32com.client


def expected_sheets(path: Path) -> int:
    book = openpyxl.load_workbook(path, read_only=True, keep_vba=False)
    n = len(book.sheetnames)
    book.close()
    return n


def check_one(excel, path: Path) -> dict:
    record = {"file": str(path), "engine_context": path.parent.name}
    try:
        expected = expected_sheets(path)
    except Exception as error:
        record.update({"verdict": "FAIL", "reason": f"fixture unreadable by openpyxl: {error}"})
        return record
    try:
        book = excel.Workbooks.Open(str(path.resolve()), UpdateLinks=0, ReadOnly=True)
    except Exception as error:
        record.update({"verdict": "FAIL", "reason": f"Excel COM open failed: {error}"})
        return record
    try:
        loaded = book.Worksheets.Count
        if loaded != expected:
            record.update({
                "verdict": "FAIL",
                "reason": f"worksheet count changed in Excel: expected {expected}, loaded {loaded}",
            })
        else:
            record["verdict"] = "PASS"
        record["worksheets_loaded"] = loaded
    finally:
        book.Close(SaveChanges=False)
    return record


def main():
    out_dir = Path(sys.argv[1])
    results_path = Path(sys.argv[2])
    outputs = sorted(out_dir.glob("*.xlsx"))
    if not outputs:
        print("no candidate outputs found")
        return
    try:
        excel = win32com.client.DispatchEx("Excel.Application")
        excel.DisplayAlerts = False
        excel.Visible = False
    except Exception as error:
        for path in outputs:
            print(json.dumps({
                "file": str(path), "verdict": "NOT_TESTED",
                "reason": f"Excel COM unavailable: {error}",
            }, ensure_ascii=False))
        return
    try:
        with open(results_path, "a", encoding="utf-8") as sink:
            for path in outputs:
                record = check_one(excel, path)
                sink.write(json.dumps(record, ensure_ascii=False) + "\n")
                print(record["verdict"], path.name)
    finally:
        excel.Quit()


if __name__ == "__main__":
    main()
