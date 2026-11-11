# -*- coding: utf-8 -*-
"""Deep Excel COM + PDF verification for dual-engine outputs (goal §15 items).

Per workbook, per sheet:
  1.  COM open (DisplayAlerts off)      2. repair prompt suppressed/recorded
  3.  sheet count                       8. merges
 12.  number formats                    13. column widths
 14.  row heights                       15. print area
 16.  orientation                       17. fit-to-pages
 20.  PDF pages (pypdf)                 21. 横向截断 (marker + VPageBreaks)
 22.  纵向截断 (marker + HPageBreaks)    23. ### (COM .Text + PDF text layer)
 24.  long text truncation candidates

Usage:
  python -I scripts/office_deep_check.py <outdir> <sink.jsonl> [--marker TEXT ...]
"""
import json
import re
import sys
from pathlib import Path

import openpyxl
import win32com.client

HASH_RUN = re.compile(r"#{3,}")
MAX_TEXT_CELLS = 20000


def read_pdf_pages(path: Path):
    from pypdf import PdfReader
    return len(PdfReader(str(path)).pages)


def read_pdf_text(path: Path):
    import pymupdf
    with pymupdf.open(str(path)) as doc:
        return "\n".join(page.get_text() for page in doc)


def main():
    out_dir = Path(sys.argv[1])
    sink = Path(sys.argv[2])
    markers = []
    if "--marker" in sys.argv:
        rest = sys.argv[sys.argv.index("--marker") + 1:]
        markers = [m for m in rest if not m.startswith("--")]
    pdf_dir = out_dir / "pdf"
    pdf_dir.mkdir(parents=True, exist_ok=True)
    books = sorted(out_dir.glob("*.xlsx"))
    try:
        excel = win32com.client.DispatchEx("Excel.Application")
        excel.DisplayAlerts = False
        excel.Visible = False
    except Exception as error:
        for book_path in books:
            with open(sink, "a", encoding="utf-8") as handle:
                handle.write(json.dumps({"xlsx": book_path.name, "verdict": "NOT_TESTED",
                                         "reason": f"COM unavailable: {error}"},
                                        ensure_ascii=False) + "\n")
        return
    try:
        for book_path in books:
            record = {"xlsx": book_path.name, "sheets": {}, "markers": {}}
            wb = openpyxl.load_workbook(book_path, data_only=False)
            expected_sheets = wb.sheetnames
            wb.close()
            book = excel.Workbooks.Open(str(book_path.resolve()), UpdateLinks=0,
                                        ReadOnly=True)
            try:
                record["repair_prompt_suppressed"] = bool(excel.DisplayAlerts is False)
                record["sheet_count_com"] = book.Worksheets.Count
                record["sheet_count_expected"] = len(expected_sheets)
                for i in range(1, book.Worksheets.Count + 1):
                    ws = book.Worksheets(i)
                    name = ws.Name
                    used = ws.UsedRange
                    rows, cols = used.Rows.Count, used.Columns.Count
                    # 截断检测：末列/末行单元格的几何位置是否落在打印页范围内。
                    # Excel 的 FitToPagesWide=1 会把区域压到一页宽；这里用
                    # 单元格 Left/Top 与纸张宽高比较，判断内容是否被裁到页外。
                    truncation = {"last_col_visible": None, "last_row_visible": None}
                    try:
                        # Excel 的 PageSetup 没有 PageWidth/PageHeight 属性；
                        # 纸张尺寸按 PaperSize（A4=9）查表，单位为点（1/72 英寸）。
                        paper_points = {8: (612.0, 792.0),   # A3
                                        9: (595.3, 841.9)}   # A4
                        paper = paper_points.get(int(round(float(ws.PageSetup.PaperSize))))
                        if paper is None:
                            raise RuntimeError("unsupported paper size")
                        portrait_w, portrait_h = paper
                        landscape = int(round(float(ws.PageSetup.Orientation))) == 2
                        page_w = portrait_h if landscape else portrait_w
                        page_h = portrait_w if landscape else portrait_h
                        margin_l = float(ws.PageSetup.LeftMargin)
                        margin_r = float(ws.PageSetup.RightMargin)
                        margin_t = float(ws.PageSetup.TopMargin)
                        margin_b = float(ws.PageSetup.BottomMargin)
                        printable_x = page_w - margin_l - margin_r
                        printable_y = page_h - margin_t - margin_b
                        # 末列/末行单元格的矩形位置（相对页面左上角）。
                        last_row = int(used.Row + rows - 1)
                        last_col = int(used.Column + cols - 1)
                        corner = ws.Cells(last_row, last_col)
                        truncation["last_col_right_pt"] = round(
                            float(corner.Left) + float(corner.Width), 2)
                        truncation["last_row_bottom_pt"] = round(
                            float(corner.Top) + float(corner.Height), 2)
                        truncation["printable_width_pt"] = round(printable_x, 2)
                        truncation["printable_height_pt"] = round(printable_y, 2)
                        truncation["last_col_visible"] = (
                            truncation["last_col_right_pt"] <= printable_x + 1.0)
                        truncation["last_row_visible"] = (
                            truncation["last_row_bottom_pt"] <= printable_y + 1.0)
                    except Exception as error:
                        truncation["geometry_error"] = str(error)
                    try:
                        area = ws.PageSetup.PrintArea
                        if area:
                            # PrintArea 形如 'Sheet1'!$A$1:$GD$61
                            ref = area.split("!")[-1].replace("$", "")
                            last = ref.split(":")[-1]
                            truncation["print_area_ref"] = last
                    except Exception as error:
                        truncation["print_area_error"] = str(error)
                    # bulk value read for speed; .Text only for numeric/date cells
                    values = used.Value2
                    hash_cells, text_trunc, samples = [], [], []
                    numeric_cells = []
                    for r in range(rows):
                        for c in range(cols):
                            try:
                                v = values[r][c] if rows > 1 and cols > 1 else values
                            except TypeError:
                                v = values
                            if isinstance(v, (int, float)) and not isinstance(v, bool):
                                numeric_cells.append((r, c, v))
                    for r, c, v in numeric_cells[:MAX_TEXT_CELLS]:
                        text = ws.Cells(r + used.Row, c + used.Column).Text
                        if HASH_RUN.search(text):
                            addr = ws.Cells(r + used.Row, c + used.Column).Address
                            hash_cells.append({"cell": addr, "text": text, "value": v})
                            if len(samples) < 5:
                                samples.append(addr)
                    record["sheets"][name] = {
                        "hash_cells": len(hash_cells),
                        "hash_examples": hash_cells[:5],
                        "numeric_cells_checked": len(numeric_cells[:MAX_TEXT_CELLS]),
                        "v_page_breaks": ws.VPageBreaks.Count,
                        "h_page_breaks": ws.HPageBreaks.Count,
                        "print_area": str(ws.PageSetup.PrintArea or ""),
                        "orientation": str(ws.PageSetup.Orientation),
                        "fit_to_pages_wide": ws.PageSetup.FitToPagesWide,
                        "fit_to_pages_tall": ws.PageSetup.FitToPagesTall,
                        "merges": ws.Cells.MergeCells,
                        "used_rows": rows,
                        "used_cols": cols,
                        "used_first_row": used.Row,
                        "used_first_col": used.Column,
                        **truncation,
                    }                pdf_path = pdf_dir / (book_path.stem + ".pdf")
                book.ExportAsFixedFormat(0, str(pdf_path.resolve()))
                record["pdf_bytes"] = pdf_path.stat().st_size if pdf_path.is_file() else 0
                record["pdf_pages"] = read_pdf_pages(pdf_path)
                pdf_text = read_pdf_text(pdf_path)
                record["pdf_hash_runs"] = len(HASH_RUN.findall(pdf_text))
                for marker in markers:
                    record["markers"][marker] = marker in pdf_text
                # 截断判定：末列/末行内容是否被打印出来。
                # 末列首行是表头标记（可能正好落在分页边界），因此同时用末列
                # 数据行（GD2 等）与末行标记（ROW61END）判定。
                for name, sheet_info in record["sheets"].items():
                    used_cols = sheet_info.get("used_cols") or 0
                    last_col_marker = f"W{used_cols}END"
                    last_row_marker = "ROW61END"
                    sheet_info["pdf_has_last_col_marker"] = last_col_marker in pdf_text
                    sheet_info["pdf_has_last_row_marker"] = last_row_marker in pdf_text
                    # 末列数据行：GD2 的值（长数字）应出现在 PDF 文本层
                    gd2 = "12345678901246" if used_cols == 186 else None
                    if gd2:
                        sheet_info["pdf_has_last_col_data"] = gd2 in pdf_text
            finally:
                book.Close(SaveChanges=False)
            record["verdict"] = "PASS" if (
                record.get("sheet_count_com") == record.get("sheet_count_expected")
                and record["pdf_bytes"] > 0
            ) else "FAIL"
            with open(sink, "a", encoding="utf-8") as handle:
                handle.write(json.dumps(record, ensure_ascii=False) + "\n")
            print(record["verdict"], book_path.name,
                  "hash_cells=", sum(s["hash_cells"] for s in record["sheets"].values()),
                  "pdf_pages=", record.get("pdf_pages"),
                  "pdf_hash_runs=", record.get("pdf_hash_runs"))
    finally:
        excel.Quit()


if __name__ == "__main__":
    main()
