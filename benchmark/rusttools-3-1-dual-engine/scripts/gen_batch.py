# -*- coding: utf-8 -*-
"""Generate N source workbooks + config for batch benchmark."""
import sys
import openpyxl
from pathlib import Path

OUT = Path(__file__).resolve().parent.parent / "fixtures"

def gen(count):
    batch_dir = OUT / f"batch_{count}"
    batch_dir.mkdir(parents=True, exist_ok=True)
    tasks = []
    for i in range(count):
        path = batch_dir / f"source_{i}.xlsx"
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Main"
        ws["A1"] = f"报表标题 {i}"
        ws.merge_cells("A1:D1")
        ws["A1"].font = openpyxl.styles.Font(bold=True, size=13)
        for r in range(2, 32):
            ws.cell(row=r, column=1, value=r * 2.5)
            ws.cell(row=r, column=2, value=f"项目-{r}")
            c = ws.cell(row=r, column=3, value=r * 100)
            c.number_format = "#,##0"
            ws.cell(row=r, column=4, value="备注" if r % 5 == 0 else None)
        wb.save(path)
        tasks.append(("Y", 1, str(path), "Main", "", str(OUT / f"batch_{count}_Rust.xlsx"), f"Sheet{i}", 1, 1, "纵向", "Y", "Y", "Y", "Y", ""))
    # config with tasks → single _Rust target
    wb = openpyxl.Workbook()
    g = wb.active
    g.title = "全局配置"
    g["A1"] = "键"; g["B1"] = "值"
    root = str(OUT)
    for i2, (k, v) in enumerate([("输入目录", root), ("输出目录", root), ("日志目录", root + "/logs")], start=2):
        g.cell(row=i2, column=1, value=k); g.cell(row=i2, column=2, value=v)
    ws = wb.create_sheet("打印配置")
    headers = ["是否启用","打印模式","源工作簿","源工作表","源工作表打印区域","目标工作簿","目标工作表",
               "FitToPagesWide","FitToPagesTall","打印方向","水平居中","垂直居中","不输出批注","零值不输出","备注"]
    for c, h in enumerate(headers, 1):
        ws.cell(row=1, column=c, value=h)
    for i, row in enumerate(tasks, start=2):
        for c, v in enumerate(row, 1):
            if v is not None:
                ws.cell(row=i, column=c, value=v)
    cfg = OUT / f"cfg_batch_{count}.xlsx"
    wb.save(cfg)
    print("  ", cfg.name, f"({count} tasks)")

if __name__ == "__main__":
    for count in (10, 50, 100):
        gen(count)
    print("done")
