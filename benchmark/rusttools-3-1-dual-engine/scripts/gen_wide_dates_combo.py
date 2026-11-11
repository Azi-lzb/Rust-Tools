# -*- coding: utf-8 -*-
"""Wide-table (186/210/222 cols), date-system (1900/1904) fixtures + combo70 config.

Completes goal §12 (### 专项) / §10.3 (日期系统) / §14.4 (70 条组合) fixtures.
Run with -I:  python -I scripts/gen_wide_dates_combo.py
"""
import sys
from pathlib import Path
from datetime import datetime, timedelta, time

import openpyxl
from openpyxl.utils.datetime import MAC_EPOCH as EPOCH_1904
from openpyxl.utils import get_column_letter

OUT = Path(__file__).resolve().parent.parent / "fixtures"

# 186 -> GD, 210 -> HB, 222 -> HN (matches real config's widest areas)
WIDE_WIDTHS = {186: "GD", 210: "HB", 222: "HN"}

# column type cycle, 1-based position within cycle
TYPES = ["DATE", "PCT", "THOUS", "LONGNUM", "TEXT", "LEADZERO", "PLAIN", "DATETIME"]
# narrow widths deliberately trigger ### under mode 1 (source width inherited)
NARROW = {"DATE": 5, "PCT": 5, "THOUS": 8, "LONGNUM": 9, "LEADZERO": 8, "PLAIN": 8}
NORMAL = {"DATE": 11, "PCT": 9, "THOUS": 12, "LONGNUM": 15, "TEXT": 12,
          "LEADZERO": 8, "PLAIN": 8, "DATETIME": 16}


def wide_book(path: Path, width: int):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Wide"
    for c in range(1, width + 1):
        ctype = TYPES[(c - 1) % len(TYPES)]
        ws.cell(row=1, column=c, value=f"C{c:03d}{ctype}")
        ws.cell(row=1, column=c).font = openpyxl.styles.Font(bold=True, size=11)
        # alternate narrow / normal widths on numeric columns
        base = NARROW[ctype] if (c % 2 == 0 and ctype in NARROW) else NORMAL[ctype]
        ws.column_dimensions[get_column_letter(c)].width = base
        for r in range(2, 62):
            i = r * 100 + c
            cell = ws.cell(row=r, column=c)
            if ctype == "DATE":
                cell.value = datetime(2024, 1, 1) + timedelta(days=(i % 900))
                cell.number_format = "yyyy-mm-dd"
            elif ctype == "PCT":
                cell.value = (i % 1000) / 97.0
                cell.number_format = "0.00%"
            elif ctype == "THOUS":
                cell.value = 1234567.89 + (i % 500)
                cell.number_format = "#,##0.00"
            elif ctype == "LONGNUM":
                cell.value = 12345678901234 + i
                cell.number_format = "0"
            elif ctype == "TEXT":
                cell.value = f"长文本截断检查-{i:05d}-abcdefghijklmnopqrstuvwxyz"
            elif ctype == "LEADZERO":
                cell.value = f"00{i % 10000:04d}"
            elif ctype == "PLAIN":
                cell.value = i * 1.5
            else:  # DATETIME
                d = datetime(2024, 1, 1) + timedelta(days=(i % 900))
                cell.value = datetime(d.year, d.month, d.day, 13, 45, 30)
                cell.number_format = "yyyy-mm-dd h:mm:ss"
    ws.cell(row=1, column=width, value=f"W{width}END")  # PDF marker: last column
    ws.cell(row=61, column=1, value="ROW61END")          # PDF marker: last row
    ws.merge_cells("A59:C59")
    ws["A59"] = "MERGE59-WIDE"
    wb.save(path)
    return path


def dates_book(path: Path, epoch_1904: bool):
    wb = openpyxl.Workbook()
    if epoch_1904:
        wb.epoch = EPOCH_1904
    ws = wb.active
    ws.title = "Dates"
    headers = ["类型", "值", "公式缓存", "公式+1", "年份公式"]
    for c, h in enumerate(headers, 1):
        ws.cell(row=1, column=c, value=h).font = openpyxl.styles.Font(bold=True)
    d = datetime(2024, 3, 15)
    rows = [
        ("普通日期", d, "yyyy-mm-dd"),
        ("时间", time(13, 45, 30), "hh:mm:ss"),
        ("日期时间", datetime(2024, 3, 15, 13, 45, 30), "yyyy-mm-dd h:mm:ss"),
        ("自定义格式", d, "dd/mm/yyyy"),
        ("纯数字串行值", 45366, "0"),
    ]
    for r, (label, value, fmt) in enumerate(rows, start=2):
        ws.cell(row=r, column=1, value=label)
        b = ws.cell(row=r, column=2, value=value)
        b.number_format = fmt
        # 公式缓存日期：openpyxl 只写公式；缓存值由 Excel COM 烘焙。
        ws.cell(row=r, column=3, value=f"=B{r}")
        ws.cell(row=r, column=4, value=f"=B{r}+1")
        ws.cell(row=r, column=5, value=f"=YEAR(B{r})")
        ws.cell(row=r, column=3).number_format = fmt
        ws.cell(row=r, column=4).number_format = fmt
    ws.column_dimensions["A"].width = 14
    ws.column_dimensions["B"].width = 20
    ws.column_dimensions["C"].width = 20
    ws.column_dimensions["D"].width = 20
    wb.save(path)
    # verify epoch round-trip
    check = openpyxl.load_workbook(path)
    assert (check.epoch == EPOCH_1904) == epoch_1904, f"epoch mismatch for {path}"
    return path


def config_book(path: Path, tasks):
    wb = openpyxl.Workbook()
    g = wb.active
    g.title = "全局配置"
    g["A1"], g["B1"] = "键", "值"
    root = str(OUT)
    for i, (k, v) in enumerate([("输入目录", root), ("输出目录", root),
                                ("日志目录", root + "/logs")], start=2):
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
    OUT.mkdir(parents=True, exist_ok=True)
    print("== wide sources ==")
    for width in (186, 210, 222):
        p = wide_book(OUT / f"src_wide{width}.xlsx", width)
        print("  ", p.name, f"({WIDE_WIDTHS[width]}1)")
    print("== date-system sources ==")
    print("  ", dates_book(OUT / "src_dates1900.xlsx", epoch_1904=False).name)
    print("  ", dates_book(OUT / "src_dates1904.xlsx", epoch_1904=True).name)

    print("== wide configs (mode1 inherit narrow widths / mode2 adaptive) ==")
    for width in (186, 210, 222):
        src = str(OUT / f"src_wide{width}.xlsx")
        for mode in (1, 2):
            name = f"cfg_wide{width}_m{mode}.xlsx"
            config_book(OUT / name, [
                ("Y", mode, src, "Wide", "", str(OUT / f"out_wide{width}_m{mode}.xlsx"),
                 "Printed", 1, 1, "", "Y", "Y", "Y", "Y", "")])
            print("  ", name)

    print("== dates config ==")
    config_book(OUT / "cfg_dates.xlsx", [
        ("Y", 1, str(OUT / "src_dates1900.xlsx"), "Dates", "", str(OUT / "out_dates.xlsx"),
         "D1900", 1, 1, "纵向", "Y", "Y", "Y", "Y", ""),
        ("Y", 1, str(OUT / "src_dates1904.xlsx"), "Dates", "", str(OUT / "out_dates.xlsx"),
         "D1904", 1, 1, "纵向", "Y", "Y", "Y", "Y", ""),
    ])
    print("   cfg_dates.xlsx (2 tasks → out_dates.xlsx)")

    print("== combo70 config (shape of real 70-row config) ==")
    batch = [str(OUT / "batch_100" / f"source_{i}.xlsx") for i in range(7)]
    tasks = []

    def add(mode, src_i, rng, tgt_i, sheet):
        tasks.append(("Y", mode, batch[src_i], "Main", rng,
                      str(OUT / "combo" / f"combo_target_{tgt_i:02d}.xlsx"),
                      sheet, 1, 1, "纵向", "Y", "Y", "Y", "Y", ""))

    # target_01: 29 tasks (sheets S0..S26 + repeatsheet SR written twice)
    for i in range(27):
        add(1, i % 7, "" if i >= 13 else "A2:D15", 1, f"S{i}")
    add(1, 1, "A2:D15", 1, "SR")
    add(1, 2, "A2:D15", 1, "SR")          # repeatsheet: clear-rewrite semantics
    # target_02: 15 tasks, first two are mode 2, includes multi-range rows
    for i in range(15):
        mode = 2 if i < 2 else 1
        rng = "A2:D5;A8:D11" if i < 4 else ("" if i == 4 else "A2:D15")
        add(mode, i % 7, rng, 2, f"T{i}")
    # targets 03..12: remaining 26 tasks, two mode 3 samples
    dist = [5, 4, 3, 3, 3, 2, 2, 2, 1, 1]
    n = 2  # running index into targets 3..12
    for tgt_off, count in enumerate(dist, start=3):
        for j in range(count):
            mode = 3 if (tgt_off == 3 and j < 2) else 1
            add(mode, n % 7, "A2:D20", tgt_off, f"U{tgt_off}_{j}")
            n += 1
    assert len(tasks) == 70, f"expected 70 tasks, got {len(tasks)}"
    assert sum(1 for t in tasks if t[4] == "") == 15, "expect 15 blank ranges"
    (OUT / "combo").mkdir(exist_ok=True)
    config_book(OUT / "cfg_combo70.xlsx", tasks)
    print("   cfg_combo70.xlsx (70 tasks, 7 sources, 12 targets; m1=66 m2=2 m3=2)")
    print("done:", OUT)


if __name__ == "__main__":
    sys.exit(main())
