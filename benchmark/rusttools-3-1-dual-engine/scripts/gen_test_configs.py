# -*- coding: utf-8 -*-
"""Dual-engine test configs: mode 1/2/3, multi-range, multi-target, zero/comments."""
import openpyxl
from pathlib import Path

OUT = Path(__file__).resolve().parent.parent / "fixtures"
OUT.mkdir(parents=True, exist_ok=True)


def source_book(path, rows=12):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Main"
    ws["A1"] = "表头"
    ws.merge_cells("A1:C1")
    ws["A1"].font = openpyxl.styles.Font(bold=True, size=14)
    for r in range(2, rows + 1):
        ws.cell(row=r, column=1, value=r * 1.5)
        ws.cell(row=r, column=2, value="text-%d" % r)
        c = ws.cell(row=r, column=3, value=0 if r % 3 == 0 else r)
        c.number_format = "0.00%"
        if r % 4 == 0:
            ws.cell(row=r, column=1).comment = openpyxl.comments.Comment("note-%d" % r, "t")
    ws.column_dimensions["A"].width = 12
    ws.column_dimensions["B"].width = 15
    wb.save(path)
    return path


def config_book(path, tasks):
    wb = openpyxl.Workbook()
    g = wb.active
    g.title = "全局配置"
    g["A1"] = "键"
    g["B1"] = "值"
    root = str(path.parent)
    for i, (k, v) in enumerate([("输入目录", root), ("输出目录", root), ("日志目录", root + "/logs")], start=2):
        g.cell(row=i, column=1, value=k)
        g.cell(row=i, column=2, value=v)
    ws = wb.create_sheet("打印配置")
    headers = ["是否启用", "打印模式", "源工作簿", "源工作表", "源工作表打印区域", "目标工作簿",
               "目标工作表", "FitToPagesWide", "FitToPagesTall", "打印方向", "水平居中",
               "垂直居中", "不输出批注", "零值不输出", "备注"]
    for c, h in enumerate(headers, 1):
        ws.cell(row=1, column=c, value=h)
    for i, row in enumerate(tasks, start=2):
        for c, v in enumerate(row, 1):
            if v is not None:
                ws.cell(row=i, column=c, value=v)
    wb.save(path)
    return path


def main():
    src = source_book(OUT / "src_mode.xlsx")
    src2 = source_book(OUT / "src_second.xlsx")
    root = OUT
    cfgs = {
        "cfg_mode1.xlsx": [("Y", 1, str(src), "Main", "", str(root / "out_m1.xlsx"), "Printed", 1, 1, "纵向", "Y", "Y", "Y", "Y", "")],
        "cfg_mode2.xlsx": [("Y", 2, str(src), "Main", "", str(root / "out_m2.xlsx"), "Printed", 1, 1, "横向", "Y", "Y", "N", "N", "")],
        "cfg_mode3.xlsx": [("Y", 3, str(src), "Main", "", str(root / "out_m3.xlsx"), "Printed", 1, 1, "纵向", "Y", "Y", "Y", "Y", "")],
        "cfg_multirange.xlsx": [("Y", 1, str(src), "Main", "A2:C5;A8:C10", str(root / "out_mr.xlsx"), "Printed", 1, 1, "纵向", "Y", "Y", "Y", "Y", "")],
        "cfg_emptyrange.xlsx": [("Y", 1, str(src), "Main", "", str(root / "out_er.xlsx"), "Printed", 1, 1, "纵向", "Y", "Y", "Y", "Y", "")],
        "cfg_multitarget.xlsx": [
            ("Y", 1, str(src), "Main", "A2:C4", str(root / "out_mt1.xlsx"), "One", 1, 1, "纵向", "Y", "Y", "Y", "Y", ""),
            ("Y", 1, str(src2), "Main", "A2:C4", str(root / "out_mt1.xlsx"), "Two", 1, 1, "纵向", "Y", "Y", "Y", "Y", ""),
            ("Y", 1, str(src), "Main", "A5:C7", str(root / "out_mt2.xlsx"), "Other", 1, 1, "纵向", "Y", "Y", "Y", "Y", ""),
        ],
        "cfg_repeatsheet.xlsx": [
            ("Y", 1, str(src), "Main", "A2:C4", str(root / "out_rs.xlsx"), "Printed", 1, 1, "纵向", "Y", "Y", "Y", "Y", ""),
            ("Y", 1, str(src2), "Main", "A2:C4", str(root / "out_rs.xlsx"), "Printed", 1, 1, "纵向", "Y", "Y", "Y", "Y", ""),
        ],
    }
    for name, tasks in cfgs.items():
        config_book(root / name, tasks)
        print("  ", name, f"({len(tasks)} tasks)")
    print("done:", OUT)


if __name__ == "__main__":
    main()
