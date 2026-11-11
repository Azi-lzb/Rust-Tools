# -*- coding: utf-8 -*-
"""Fixture for user-feedback fixes: blank styled cells, error cells, center
alignment, theme colors (Office 2013 palette), high-precision float caches."""
import sys
from pathlib import Path

import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.styles.colors import Color

OUT = Path(__file__).resolve().parent.parent / "fixtures"


def main():
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "F"
    thin = Side(style="thin")
    box = Border(left=thin, right=thin, top=thin, bottom=thin)
    # 1) 空样式格：有边框+填充+日期格式，无值（对齐真实模板 C1 s=14）
    for c in range(1, 6):
        cell = ws.cell(row=2, column=c)
        cell.border = box
        cell.number_format = "mm-dd-yy"
        if c == 1:
            cell.fill = PatternFill("solid", fgColor="FFFF00")
    # 2) center / centerContinuous / vertical 对齐格
    ws["A4"] = "居中"
    ws["A4"].alignment = Alignment(horizontal="center", vertical="center")
    ws["B4"] = "跨列居中"
    ws["B4"].alignment = Alignment(horizontal="centerContinuous")
    ws["C4"] = 42
    ws["C4"].alignment = Alignment(horizontal="center")
    # 3) theme 颜色（对齐源模板 A1 真身）：填充 theme=8（accent5）、字体 indexed=9
    ws["A6"] = "主题色标题"
    ws["A6"].font = Font(bold=True, size=14, color=Color(indexed=9))
    ws["A6"].fill = PatternFill("solid", fgColor=Color(theme=8))
    ws["B6"] = "普通"
    # 4) 错误值公式（COM 烘焙后成真错误格）
    ws["A8"] = "=1/0"
    ws["B8"] = "=1/0+1"
    ws["C8"] = "=NA()"
    # 5) 高精度浮点（COM 重算后缓存为完整精度）
    ws["A10"] = "=1/3"
    ws["B10"] = "=2/7"
    # 6) 有值 + 边框（确认值格样式不受空格修复影响）
    ws["A12"] = 3.5
    ws["A12"].border = box
    ws["B12"] = "文本"
    ws["B12"].border = box
    wb.save(OUT / "src_feedback.xlsx")
    print("fixture:", OUT / "src_feedback.xlsx")

    # 测试配置
    cfg = openpyxl.Workbook()
    g = cfg.active
    g.title = "全局配置"
    g["A1"], g["B1"] = "键", "值"
    for i, (k, v) in enumerate([("输入目录", str(OUT)), ("输出目录", str(OUT)),
                                ("日志目录", str(OUT / "logs"))], start=2):
        g.cell(row=i, column=1, value=k)
        g.cell(row=i, column=2, value=v)
    ws2 = cfg.create_sheet("打印配置")
    headers = ["是否启用", "打印模式", "源工作簿", "源工作表", "源工作表打印区域", "目标工作簿",
               "目标工作表", "FitToPagesWide", "FitToPagesTall", "打印方向", "水平居中",
               "垂直居中", "不输出批注", "零值不输出", "备注"]
    for c, h in enumerate(headers, 1):
        ws2.cell(row=1, column=c, value=h)
    task = ("Y", 1, str(OUT / "src_feedback.xlsx"), "F", "", str(OUT / "out_feedback.xlsx"),
            "Printed", 1, 1, "", "Y", "Y", "Y", "Y", "")
    for c, v in enumerate(task, 1):
        ws2.cell(row=2, column=c, value=v)
    cfg.save(OUT / "cfg_feedback.xlsx")
    print("config:", OUT / "cfg_feedback.xlsx")


if __name__ == "__main__":
    sys.exit(main())
