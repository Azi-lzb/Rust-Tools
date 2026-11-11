# -*- coding: utf-8 -*-
"""Cell-level dual-engine output comparison (values, merges, widths, print setup).

Usage:
  python -I scripts/compare_xlsx.py <openpyxl_out> <rust_out> <sink.json> [style_limit]

Emits one JSON object: {openpyxl, rust, value_diffs, style_diffs, setup_diffs,
value_diff_count, style_diff_count}.
"""
import json
import re
import sys
import zipfile
from pathlib import Path

import openpyxl

# ECMA-376 标准 indexed 颜色表（0-63；Excel 无覆盖时的渲染基准）。
INDEXED_COLORS = {
    0: "000000", 1: "FFFFFF", 2: "FF0000", 3: "00FF00", 4: "0000FF",
    5: "FFFF00", 6: "FF00FF", 7: "00FFFF", 8: "800000", 9: "008000",
    10: "000080", 11: "808000", 12: "800080", 13: "008080", 14: "C0C0C0",
    15: "808080", 16: "9999FF", 17: "993366", 18: "FFFFCC", 19: "CCFFFF",
    20: "660066", 21: "FF8080", 22: "0066CC", 23: "CCCCFF", 24: "000080",
    25: "FF00FF", 26: "FFFF00", 27: "00FFFF", 28: "800080", 29: "800000",
    30: "008080", 31: "0000FF", 32: "00CCFF", 33: "CCFFFF", 34: "CCFFCC",
    35: "FFFF99", 36: "99CCFF", 37: "FF99CC", 38: "CC99FF", 39: "FFCC99",
    40: "3366FF", 41: "33CCCC", 42: "99CC00", 43: "FFCC00", 44: "FF9900",
    45: "FF6600", 46: "666699", 47: "969696", 48: "003366", 49: "339966",
    50: "003300", 51: "333300", 52: "993300", 53: "993366", 54: "333399",
    55: "333333", 56: "800000", 57: "FF0000", 58: "00FF00", 59: "0000FF",
    60: "FFFF00", 61: "FF00FF", 62: "00FFFF", 63: "FFFFFF",
}


def theme_slots(path: Path):
    """主题槽位（Excel 语义：0=lt1, 1=dk1, 2=lt2, 3=dk2, 4-9=accent1-6）。"""
    try:
        with zipfile.ZipFile(path) as archive:
            theme = archive.read("xl/theme/theme1.xml").decode("utf-8", "replace")
    except (OSError, KeyError):
        return []
    scheme = re.search(r"<a:clrScheme.*?>(.*?)</a:clrScheme>", theme, re.S)
    if not scheme:
        return []
    slots = []
    for name in ("lt1", "dk1", "lt2", "dk2", "accent1", "accent2", "accent3",
                 "accent4", "accent5", "accent6"):
        m = re.search(rf"<a:{name}\b[^>]*>(.*?)</a:{name}>", scheme.group(1), re.S)
        if not m:
            slots.append(None)
            continue
        inner = m.group(1)
        last = re.search(r'lastClr="([^"]+)"', inner)
        val = re.search(r'val="([^"]+)"', inner)
        color = last.group(1) if last else (val.group(1) if val else None)
        slots.append(color)
    return slots


def default_font_name(path: Path) -> str:
    """工作簿默认字体名（styles.xml 第一个 <font> 的 <name>）。

    源字体记录可能没有 <name> 元素（openpyxl 写 bold/size 时如此），
    此时 openpyxl 读回 None，Excel 按工作簿默认字体渲染；
    rust_xlsxwriter 则总是写出显式字体名。比较时统一成“有效字体名”。
    """
    try:
        with zipfile.ZipFile(path) as archive:
            styles = archive.read("xl/styles.xml").decode("utf-8", "replace")
    except (OSError, KeyError):
        return "Calibri"
    fonts = re.search(r"<fonts[^>]*>(.*?)</fonts>", styles, re.S)
    if not fonts:
        return "Calibri"
    name = re.search(r"<name[^>]*val=\"([^\"]+)\"", fonts.group(1))
    return name.group(1) if name else "Calibri"


def excel_column_pixels(width: float) -> int:
    """Excel 的 char→px 渲染公式（MDW=7，padding 18/256 修正项）。

    存盘 char 宽不同的两个文件只要渲染像素相同，Excel 中显示完全一致
    （rust_xlsxwriter 像素反算的 18.5703 与 openpyxl 的 18.625 同为 130px）。
    """
    return int(((256.0 * width + 18.0) / 256.0) * 7.0)


def column_widths_by_letter(path: Path):
    """按 XML <col min=.. max=..> 展开每列的有效宽度。

    openpyxl 读取分组列时只把 min 列登记到 column_dimensions，
    其余列读回为 None（读取器简化，不代表文件缺失宽度）。
    为公平比较宽度，这里直接解析 XML，把分组展开到每列。
    """
    widths = {}
    try:
        with zipfile.ZipFile(path) as archive:
            for part in archive.namelist():
                if not part.startswith("xl/worksheets/sheet"):
                    continue
                xml = archive.read(part).decode("utf-8", "replace")
                cols = re.search(r"<cols>(.*?)</cols>", xml, re.S)
                if not cols:
                    continue
                for col in re.finditer(r"<col\b[^>]*/?>", cols.group(1)):
                    attrs = dict(re.findall(r'(\w+)="([^"]*)"', col.group(0)))
                    if "width" not in attrs:
                        continue
                    low = int(attrs.get("min", 0))
                    high = int(attrs.get("max", low))
                    for index in range(low, max(low, high) + 1):
                        widths[index] = round(float(attrs["width"]), 4)
                break
    except (OSError, KeyError, ValueError):
        return widths
    return widths


def snapshot(path: Path):
    wb = openpyxl.load_workbook(path, data_only=True)
    default_name = default_font_name(path)
    slots = theme_slots(path)

    def color_of(color_obj):
        # 颜色统一归一到“有效 RGB”比较：theme/indexed/rgb 三种表示在 Excel
        # 中渲染相同的颜色应判等（openpyxl 输出保留 theme/indexed 引用，
        # Rust 输出为解析后的 RGB）。解析失败才保留原始表示。
        if color_obj is None:
            # 字体无颜色 = 自动色 = 黑（openpyxl 输出省略 color 元素，
            # rust_xlsxwriter 写出 theme=1/rgb 000000，渲染等价）。
            return ("resolved", "000000")
        if color_obj.type == "rgb":
            return ("resolved", str(color_obj.rgb)[-6:].upper())
        if color_obj.type == "theme":
            idx = color_obj.theme
            base = slots[idx] if idx is not None and idx < len(slots) else None
            tint = color_obj.tint or 0.0
            if not base or not re.fullmatch(r"[0-9A-Fa-f]{6}", base):
                return ("unresolved-theme", idx, tint)
            r, g, b = (int(base[i:i + 2], 16) for i in (0, 2, 4))
            if tint > 0:
                r, g, b = (round(c * (1 - tint) + 255 * tint) for c in (r, g, b))
            elif tint < 0:
                r, g, b = (round(c * (1 + tint)) for c in (r, g, b))
            return ("resolved", f"{r:02X}{g:02X}{b:02X}")
        if color_obj.type == "indexed":
            table = INDEXED_COLORS
            hex_value = table.get(color_obj.indexed)
            if hex_value:
                return ("resolved", hex_value[-6:].upper())
            return ("unresolved-indexed", color_obj.indexed)
        return (color_obj.type,)
    data = {
        "sheets": wb.sheetnames,
        "values": {},
        "merges": {},
        "col_widths": {},
        "row_heights": {},
        "print_area": {},
        "orientation": {},
        "footer": {},
        "fonts": {},
        "fills": {},
        "numfmts": {},
        "aligns": {},
        "borders": {},
    }
    for name in wb.sheetnames:
        ws = wb[name]
        values = {}
        for row in ws.iter_rows():
            for cell in row:
                if cell.value is not None:
                    if isinstance(cell.value, (int, float)) and not isinstance(cell.value, bool):
                        values[cell.coordinate] = f"num:{float(cell.value):.10g}"
                    else:
                        values[cell.coordinate] = repr(cell.value)
        data["values"][name] = values
        data["merges"][name] = sorted(str(rng) for rng in ws.merged_cells.ranges)
        data["col_widths"][name] = {
            k: round(v.width, 4) for k, v in ws.column_dimensions.items() if v.width
        }
        # 列宽按 XML 分组展开后逐列比较（openpyxl 读取器会漏读分组内的列）。
        data["col_widths_xml"] = column_widths_by_letter(path)
        # 行高按“相对默认行高”比较：openpyxl 输出所有行显式 ht=15（其默认），
        # rust 输出（默认=源 13.5）只写非默认行；Excel 渲染等价。
        default_height = None
        try:
            with zipfile.ZipFile(path) as archive:
                for part in archive.namelist():
                    if part.startswith("xl/worksheets/sheet"):
                        sheet_xml = archive.read(part).decode("utf-8", "replace")
                        m = re.search(r'defaultRowHeight="([0-9.]+)"', sheet_xml)
                        if m:
                            default_height = float(m.group(1))
                        break
        except (OSError, KeyError):
            pass
        heights = {}
        for k, v in ws.row_dimensions.items():
            if v.height is None:
                continue
            if default_height is not None and abs(v.height - default_height) < 0.01:
                continue
            heights[k] = round(v.height, 4)
        data["row_heights"][name] = heights
        data["print_area"][name] = str(ws.print_area) if ws.print_area else ""
        data["orientation"][name] = ws.page_setup.orientation or ""
        data["footer"][name] = str(ws.oddFooter.center.text or "")
        fonts, fills, numfmts, aligns, borders = {}, {}, {}, {}, {}
        for row in ws.iter_rows():
            for cell in row:
                if cell.value is None and not cell.has_style:
                    continue
                c = cell.coordinate
                fonts[c] = ((cell.font.name or default_name), cell.font.size,
                            bool(cell.font.bold), bool(cell.font.italic),
                            color_of(cell.font.color))
                if cell.fill and cell.fill.patternType:
                    fills[c] = (cell.fill.patternType, color_of(cell.fill.fgColor))
                numfmts[c] = cell.number_format
                aligns[c] = (cell.alignment.horizontal, cell.alignment.vertical,
                             bool(cell.alignment.wrap_text))
                b = cell.border
                borders[c] = tuple(
                    getattr(b, side).style or ""
                    for side in ("left", "right", "top", "bottom"))
        data["fonts"][name] = fonts
        data["fills"][name] = fills
        data["numfmts"][name] = numfmts
        data["aligns"][name] = aligns
        data["borders"][name] = borders
    wb.close()
    return data


def diff_section(a, b, limit):
    diffs = []
    for key in sorted(set(a) | set(b)):
        va, vb = a.get(key), b.get(key)
        if va != vb:
            if isinstance(va, dict) and isinstance(vb, dict):
                for sub in sorted(set(va) | set(vb)):
                    sa, sb = va.get(sub), vb.get(sub)
                    if sa != sb:
                        diffs.append({"key": f"{key}!{sub}", "a": sa, "b": sb})
                        if len(diffs) >= limit:
                            return diffs
            else:
                diffs.append({"key": key, "a": va, "b": vb})
                if len(diffs) >= limit:
                    break
    return diffs


def main():
    openpyxl_out, rust_out = Path(sys.argv[1]), Path(sys.argv[2])
    sink, style_limit = Path(sys.argv[3]), int(sys.argv[4]) if len(sys.argv) > 4 else 20
    a, b = snapshot(openpyxl_out), snapshot(rust_out)
    value_diffs = diff_section(a["values"], b["values"], style_limit)
    style_diffs = []
    for section in ("fonts", "fills", "numfmts", "aligns", "borders",
                    "row_heights", "merges"):
        style_diffs.extend(
            {"section": section, **item}
            for item in diff_section(a[section], b[section], style_limit))
    # 列宽按 XML 展开后逐列比较。
    for index in sorted(set(a["col_widths_xml"]) | set(b["col_widths_xml"])):
        left = a["col_widths_xml"].get(index)
        right = b["col_widths_xml"].get(index)
        if left is None or right is None:
            equal = left == right
        else:
            equal = excel_column_pixels(left) == excel_column_pixels(right)
        if not equal:
            style_diffs.append({"section": "col_widths",
                                "key": f"col{index}", "a": left, "b": right})
            if len(style_diffs) >= style_limit:
                break
    setup_diffs = diff_section(a["print_area"], b["print_area"], 10)
    setup_diffs += diff_section(a["orientation"], b["orientation"], 10)
    setup_diffs += diff_section(a["footer"], b["footer"], 10)
    result = {
        "openpyxl": str(openpyxl_out),
        "rust": str(rust_out),
        "value_diff_count": len(value_diffs),
        "style_diff_count": len(style_diffs),
        "value_diffs": value_diffs[:style_limit],
        "style_diffs": style_diffs[:style_limit],
        "setup_diffs": setup_diffs,
    }
    Path(sink).write_text(json.dumps(result, ensure_ascii=False, indent=1),
                          encoding="utf-8")
    print(f"{openpyxl_out.name} vs {rust_out.name}:",
          "value_diffs =", result["value_diff_count"],
          "style_diffs =", result["style_diff_count"],
          "setup_diffs =", len(setup_diffs))


if __name__ == "__main__":
    main()
