# -*- coding: utf-8 -*-
"""Benchmark fixture generator (openpyxl 3.1.5, bundled runtime).

Generates deterministic fixtures under benchmark/excel-modify-engine/fixtures/.
Real business samples are COPIED from repo test fixtures (originals untouched).
Run: dist/python/python.exe -I scripts/gen_fixtures.py
"""
import shutil
import sys
import zipfile
from datetime import datetime
from pathlib import Path
from xml.etree import ElementTree as ET

import openpyxl
from openpyxl.comments import Comment
from openpyxl.formatting.rule import CellIsRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.worksheet.datavalidation import DataValidation

ROOT = Path(__file__).resolve().parents[3]
OUT = Path(__file__).resolve().parents[1] / "fixtures"
MAIN_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"


def perf_book(rows: int) -> openpyxl.Workbook:
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Main"
    ws.append(["col%02d" % c for c in range(1, 21)])
    for r in range(2, rows + 2):
        for c in range(1, 21):
            if c % 3 == 0:
                ws.cell(row=r, column=c, value="s%d_%d" % (r, c))
            elif c % 3 == 1:
                ws.cell(row=r, column=c, value=float(r * 20 + c) + 0.5)
            else:
                ws.cell(row=r, column=c, value=r * 20 + c)
    return wb


def gen_perf():
    for name, rows in [("perf_small_1k", 1000), ("perf_mid_10k", 10000),
                       ("perf_large_50k", 50000), ("perf_huge_100k", 100000)]:
        wb = perf_book(rows)
        path = OUT / (name + ".xlsx")
        wb.save(path)
        wb.close()
        print("  ", path.name, "%.1f KB" % (path.stat().st_size / 1024))


def style_demo(ws):
    bold = Font(bold=True, size=12)
    fill = PatternFill("solid", fgColor="FFF2CC")
    thin = Side(style="thin", color="999999")
    ws["B2"].font = bold
    ws["B3"].fill = fill
    ws["B4"].border = Border(left=thin, right=thin, top=thin, bottom=thin)
    ws["B5"].alignment = Alignment(horizontal="center", vertical="center")
    ws["C2"].number_format = "0.000"
    ws["C3"].number_format = "#,##0"
    ws["C4"].number_format = "yyyy-mm-dd"
    ws.row_dimensions[6].height = 30.0
    ws.column_dimensions["D"].width = 25.0


def gen_capability_fixtures():
    # basic values
    wb = openpyxl.Workbook(); ws = wb.active; ws.title = "Main"
    for r in range(1, 21):
        for c in range(1, 6):
            ws.cell(row=r, column=c, value=r * 10 + c)
    style_demo(ws)
    wb.save(OUT / "f_basic.xlsx"); wb.close()

    # formulas (caches injected afterwards)
    wb = openpyxl.Workbook(); ws = wb.active; ws.title = "Main"
    ws2 = wb.create_sheet("Aux")
    for r in range(1, 11):
        ws2.cell(row=r, column=1, value=r)
    for r in range(1, 11):
        ws.cell(row=r, column=1, value="=Aux!A%d*2" % r)
    ws.cell(row=1, column=3, value="=SUM(Aux!A1:A10)")
    ws.cell(row=2, column=3, value='=IF(A1>0,"pos","neg")')
    wb.save(OUT / "f_formulas.xlsx"); wb.close()
    inject_caches(OUT / "f_formulas.xlsx", {
        ("Main", "A1"): "2", ("Main", "A2"): "4", ("Main", "A3"): "6",
        ("Main", "A4"): "8", ("Main", "A5"): "10", ("Main", "A6"): "12",
        ("Main", "A7"): "14", ("Main", "A8"): "16", ("Main", "A9"): "18",
        ("Main", "A10"): "20", ("Main", "C1"): "55", ("Main", "C2"): "pos",
    })

    # multi-sheet
    wb = openpyxl.Workbook()
    for i in range(5):
        ws = wb.active if i == 0 else wb.create_sheet()
        ws.title = "S%d" % (i + 1)
        for r in range(1, 31):
            for c in range(1, 6):
                ws.cell(row=r, column=c, value=(i + 1) * 1000 + r * 10 + c)
    wb.save(OUT / "f_multisheet.xlsx"); wb.close()

    # styles (already covered inside basic; dedicated too)
    wb = openpyxl.Workbook(); ws = wb.active; ws.title = "Main"
    style_demo(ws)
    for c in range(1, 8):
        ws.cell(row=8, column=c, value=c).font = Font(italic=True)
    wb.save(OUT / "f_styles.xlsx"); wb.close()

    # merged
    wb = openpyxl.Workbook(); ws = wb.active; ws.title = "Main"
    ws["A1"] = "title"; ws.merge_cells("A1:C1")
    ws["B3"] = "m3"; ws.merge_cells("B3:C4")
    ws["E5"] = "m5"; ws.merge_cells("E5:F6")
    wb.save(OUT / "f_merged.xlsx"); wb.close()

    # comments
    wb = openpyxl.Workbook(); ws = wb.active; ws.title = "Main"
    ws["A1"] = "with note"
    ws["A1"].comment = Comment("author line one\nsecond line", "bench")
    ws["B2"] = 1; ws["B2"].comment = Comment("note on b2", "bench")
    wb.save(OUT / "f_comments.xlsx"); wb.close()

    # defined names (global + local scope)
    wb = openpyxl.Workbook(); ws = wb.active; ws.title = "Main"
    ws["A1"] = 1; ws["A2"] = 2
    wb.defined_names.add(DefinedName("G_GlobalName", attr_text="Main!$A$1:$A$2"))
    wb.defined_names.add(DefinedName("G_Point", attr_text="Main!$B$2"))
    ws.defined_names.add(DefinedName("L_LocalName", attr_text="Main!$A$2"))
    wb.save(OUT / "f_definednames.xlsx"); wb.close()

    # data validation
    wb = openpyxl.Workbook(); ws = wb.active; ws.title = "Main"
    dv = DataValidation(type="whole", operator="between", formula1=1, formula2=100,
                        allow_blank=True)
    dv.error = "must be 1-100"
    ws.add_data_validation(dv)
    dv.add("B1:B10")
    for r in range(1, 11):
        ws.cell(row=r, column=1, value=r)
    wb.save(OUT / "f_datavalidation.xlsx"); wb.close()

    # conditional formatting
    wb = openpyxl.Workbook(); ws = wb.active; ws.title = "Main"
    for r in range(1, 16):
        ws.cell(row=r, column=1, value=r)
    ws.conditional_formatting.add("A1:A15",
        CellIsRule(operator="greaterThan", formula=["10"], fill=PatternFill("solid", fgColor="FFC7CE")))
    wb.save(OUT / "f_condformat.xlsx"); wb.close()

    # print settings
    wb = openpyxl.Workbook(); ws = wb.active; ws.title = "Main"
    for r in range(1, 41):
        for c in range(1, 6):
            ws.cell(row=r, column=c, value=r * 5 + c)
    ws.print_area = "A1:D20"
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth = 1
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.print_options.horizontalCentered = True
    wb.save(OUT / "f_print.xlsx"); wb.close()

    # complex template (10 sheets, mixed features)
    wb = openpyxl.Workbook()
    ws = wb.active; ws.title = "汇总"
    ws["A1"] = "汇总表"; ws.merge_cells("A1:D1")
    for i in range(1, 6):
        src = wb.create_sheet("明细%d" % i)
        for r in range(1, 51):
            for c in range(1, 6):
                src.cell(row=r, column=c, value=i * 1000 + r * 10 + c)
        src["G1"] = "=A1+B1"
        src["G1"].comment = Comment("check col G", "bench")
        col = get_column_letter(i)
        ws.cell(row=i + 1, column=1, value="明细%d" % i)
        ws.cell(row=i + 1, column=2, value="=明细%d!A1" % i)
        src.column_dimensions["A"].width = 18.0
        src.page_setup.orientation = "portrait"
    wb.defined_names.add(DefinedName("G_范围", attr_text="明细1!$A$1:$A$5"))
    ws.print_area = "A1:D10"
    wb.save(OUT / "f_complex.xlsx"); wb.close()


def inject_caches(path: Path, caches: dict) -> None:
    """Zip-level injection of cached formula values (openpyxl cannot write them)."""
    zin = zipfile.ZipFile(path, "r")
    # map sheet name -> part
    wb_xml = ET.fromstring(zin.read("xl/workbook.xml"))
    rels_xml = ET.fromstring(zin.read("xl/_rels/workbook.xml.rels"))
    rel_map = {r.get("Id"): r.get("Target") for r in rels_xml}
    sheet_part = {}
    for sh in wb_xml.find("{%s}sheets" % MAIN_NS):
        rid = sh.get("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id")
        target = rel_map[rid]
        if target.startswith("/"):
            sheet_part[sh.get("name")] = target.lstrip("/")
        else:
            sheet_part[sh.get("name")] = "xl/" + target
    cached = {}
    for (sheet, addr), value in caches.items():
        cached.setdefault(sheet_part[sheet], {})[addr] = value
    zout = zipfile.ZipFile(str(path) + ".tmp", "w", zipfile.ZIP_DEFLATED)
    for item in zin.infolist():
        data = zin.read(item.filename)
        if item.filename in cached:
            root = ET.fromstring(data)
            ns = "{%s}" % MAIN_NS
            for cell in root.iter(ns + "c"):
                addr = cell.get("r")
                if addr in cached[item.filename] and cell.find(ns + "f") is not None:
                    v = cell.find(ns + "v")
                    if v is None:
                        v = ET.SubElement(cell, ns + "v")
                    text = str(cached[item.filename][addr])
                    v.text = text
                    if not text.replace(".", "", 1).replace("-", "", 1).isdigit():
                        cell.set("t", "str")
            ET.register_namespace("", MAIN_NS)
            data = ET.tostring(root, encoding="utf-8", xml_declaration=True)
        zout.writestr(item, data)
    zin.close(); zout.close()
    Path(str(path) + ".tmp").replace(path)


def gen_external_link():
    """Hand-built externalLink part (openpyxl cannot author external links)."""
    src = OUT / "f_basic.xlsx"
    dst = OUT / "f_external_link.xlsx"
    zin = zipfile.ZipFile(src, "r")
    wb_xml = zin.read("xl/workbook.xml").decode("utf-8")
    wb_rels = zin.read("xl/_rels/workbook.xml.rels").decode("utf-8")
    ctypes = zin.read("[Content_Types].xml").decode("utf-8")
    # find a free rel id
    rid = "rIdExt1"
    ext_xml = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
               '<externalLink xmlns="%s">'
               '<externalBook><sheetNames><sheetName val="Sheet1"/></sheetNames>'
               '<sheetDataSet><sheetData sheetId="0">'
               '<cell r="C3"><v>12345</v></cell>'
               '</sheetData></sheetDataSet></externalBook></externalLink>' % MAIN_NS)
    wb_xml = wb_xml.replace("</workbook>",
        '<externalReferences><externalReference r:id="%s"/></externalReferences></workbook>' % rid)
    wb_xml = wb_xml.replace('<workbook xmlns="%s"' % MAIN_NS,
        '<workbook xmlns="%s" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"' % MAIN_NS, 1)
    wb_rels = wb_rels.replace("</Relationships>",
        '<Relationship Id="%s" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLink" Target="externalLinks/externalLink1.xml"/></Relationships>' % rid)
    ctypes = ctypes.replace("</Types>",
        '<Override PartName="/xl/externalLinks/externalLink1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.externalLink+xml"/></Types>')
    ext_rels = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLinkPath" '
                'Target="file:///C:/bench-nonexistent-target.xlsx" TargetMode="External"/></Relationships>')
    replaced = {"xl/workbook.xml", "xl/_rels/workbook.xml.rels", "[Content_Types].xml"}
    zout = zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED)
    for item in zin.infolist():
        if item.filename in replaced:
            continue
        zout.writestr(item, zin.read(item.filename))
    zout.writestr("xl/externalLinks/externalLink1.xml", ext_xml)
    zout.writestr("xl/externalLinks/_rels/externalLink1.xml.rels", ext_rels)
    zout.writestr("xl/workbook.xml", wb_xml)
    zout.writestr("xl/_rels/workbook.xml.rels", wb_rels)
    zout.writestr("[Content_Types].xml", ctypes)
    zin.close(); zout.close()
    print("  f_external_link.xlsx (hand-built part)")


def gen_real_copies():
    pairs = [
        (ROOT / "测试用例/2/目标用例.xlsx", "real_target.xlsx"),
        (ROOT / "测试用例/2/检查重复去重用例.xlsx", "real_dedup.xlsx"),
        (ROOT / "pytools-core/tests/fixtures/vba_config/VbaV2.4.2.xlsm", "real_vba.xlsm"),
        (ROOT / "pytools-core/tests/fixtures/summary_used_range/vba_enabled_xlsxwriter_macro01.xlsm",
         "real_vba_xlsxwriter.xlsm"),
    ]
    for src, name in pairs:
        if src.is_file():
            shutil.copy2(src, OUT / name)
            print("  copied real sample:", name, "(source untouched)")
        else:
            print("  MISSING real sample:", src)


def main():
    OUT.mkdir(parents=True, exist_ok=True)
    print("openpyxl", openpyxl.__version__, "->", OUT)
    gen_perf()
    gen_capability_fixtures()
    gen_external_link()
    gen_real_copies()
    print("done:", len(list(OUT.glob('*'))), "files")


if __name__ == "__main__":
    sys.exit(main())
