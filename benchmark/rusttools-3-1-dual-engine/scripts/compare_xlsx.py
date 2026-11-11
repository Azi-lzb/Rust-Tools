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
                    values[cell.coordinate] = repr(cell.value)
        data["values"][name] = values
        data["merges"][name] = sorted(str(rng) for rng in ws.merged_cells.ranges)
        data["col_widths"][name] = {
            k: round(v.width, 4) for k, v in ws.column_dimensions.items() if v.width
        }
        # 列宽按 XML 分组展开后逐列比较（openpyxl 读取器会漏读分组内的列）。
        data["col_widths_xml"] = column_widths_by_letter(path)
        data["row_heights"][name] = {
            k: round(v.height, 4) for k, v in ws.row_dimensions.items() if v.height
        }
        data["print_area"][name] = str(ws.print_area) if ws.print_area else ""
        data["orientation"][name] = ws.page_setup.orientation or ""
        data["footer"][name] = str(ws.oddFooter.center.text or "")
        fonts, fills, numfmts, aligns, borders = {}, {}, {}, {}, {}
        for row in ws.iter_rows():
            for cell in row:
                if cell.value is None and not cell.has_style:
                    continue
                c = cell.coordinate
                fonts[c] = (cell.font.name or default_name, cell.font.size,
                            bool(cell.font.bold), bool(cell.font.italic))
                if cell.fill and cell.fill.patternType:
                    fills[c] = (cell.fill.patternType, str(cell.fill.fgColor.rgb))
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
        if left != right:
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
