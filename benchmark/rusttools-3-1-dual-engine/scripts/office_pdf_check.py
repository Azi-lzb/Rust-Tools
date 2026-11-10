# -*- coding: utf-8 -*-
"""Excel COM PDF export + print settings validation for dual-engine outputs."""
import json
import sys
from pathlib import Path

import openpyxl
import win32com.client


def check_pdf(excel, xlsx_path: Path, pdf_path: Path) -> dict:
    record = {"xlsx": str(xlsx_path)}
    expected_sheets = None
    wb = openpyxl.load_workbook(xlsx_path, keep_vba=False)
    expected_sheets = len(wb.sheetnames)
    wb.close()

    book = excel.Workbooks.Open(str(xlsx_path.resolve()), UpdateLinks=0, ReadOnly=True)
    try:
        loaded = book.Worksheets.Count
        record["worksheets"] = loaded
        if loaded != expected_sheets:
            record["verdict"] = "FAIL"
            record["reason"] = f"sheet count {loaded} != {expected_sheets}"
            return record
        # Export to PDF
        book.ExportAsFixedFormat(0, str(pdf_path.resolve()))  # 0 = xlTypePDF
        if not pdf_path.is_file():
            record["verdict"] = "FAIL"
            record["reason"] = "PDF not created"
            return record
        record["pdf_bytes"] = pdf_path.stat().st_size
        # 检查导出的工作表打印区域（COM 侧）
        areas = []
        for i in range(1, loaded + 1):
            ws = book.Worksheets(i)
            area = None
            try:
                pa = ws.PageSetup.PrintArea
                if pa:
                    area = str(pa)
            except Exception:
                pass
            areas.append(area)
        record["print_areas"] = areas
        record["orientation"] = str(book.Worksheets(1).PageSetup.Orientation)
        record["verdict"] = "PASS"
        return record
    finally:
        book.Close(SaveChanges=False)


def main():
    out_dir = Path(sys.argv[1])
    results_path = Path(sys.argv[2])
    pdf_dir = out_dir / "pdf"
    pdf_dir.mkdir(parents=True, exist_ok=True)
    outputs = sorted(out_dir.glob("*.xlsx"))
    if not outputs:
        print("no outputs")
        return
    try:
        excel = win32com.client.DispatchEx("Excel.Application")
        excel.DisplayAlerts = False
        excel.Visible = False
    except Exception as error:
        for path in outputs:
            print(json.dumps({"file": str(path), "verdict": "NOT_TESTED",
                              "reason": f"Excel COM unavailable: {error}"}, ensure_ascii=False))
        return
    try:
        with open(results_path, "a", encoding="utf-8") as sink:
            for path in outputs:
                pdf_path = pdf_dir / (path.stem + ".pdf")
                record = check_pdf(excel, path, pdf_path)
                record["pdf_file"] = str(pdf_path)
                sink.write(json.dumps(record, ensure_ascii=False) + "\n")
                print(record["verdict"], path.name,
                      f"pdf={record.get('pdf_bytes', 0)}B")
    finally:
        excel.Quit()


if __name__ == "__main__":
    main()
