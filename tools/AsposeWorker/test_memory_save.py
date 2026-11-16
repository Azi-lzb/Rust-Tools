"""Independent synthetic XLS fixture and saved-file contract verification."""
import hashlib
import json
import struct
import sys
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

import openpyxl
import xlrd
import xlrd.compdoc
import xlwt


def fixture(root):
    root.mkdir(exist_ok=False, parents=True)
    spec_path = Path(__file__).resolve().parents[2] / "benchmark/excel-engine-evaluation/baselines/results/four-dim-20261114/fixtures/calc-spec.json"
    spec = json.loads(spec_path.read_text("utf-8"))
    wb = xlwt.Workbook()
    data, calc = wb.add_sheet("Data"), wb.add_sheet("Calc")
    style = xlwt.easyxf("font: bold on; pattern: pattern solid, fore_colour yellow;")
    for r, row in enumerate(spec["data"]):
        for c, value in enumerate(row):
            data.write(r, c, value, style if r == 0 else xlwt.Style.default_style)
    for r, row in enumerate(spec["formulas"]):
        for c, formula in enumerate(row):
            calc.write(r, c, xlwt.Formula(formula.lstrip("=")))
    data.col(3).width = 4500
    data.row(11).height = 500
    calc.selected = True
    data.selected = False
    wb.active_sheet = 1
    path = root / "synthetic-complex.xls"
    wb.save(str(path))
    book = xlrd.open_workbook(path, formatting_info=True)
    assert book.sheet_names() == ["Data", "Calc"]
    # Verify actual BIFF WINDOW1 active tab independently of Aspose.
    stream = xlrd.compdoc.CompDoc(path.read_bytes()).get_named_stream("Workbook")
    pos, active = 0, None
    while pos + 4 <= len(stream):
        kind, size = struct.unpack_from("<HH", stream, pos)
        if kind == 0x003D:
            active = struct.unpack_from("<H", stream, pos + 4 + 10)[0]
            break
        pos += 4 + size
    assert active == 1, active
    manifest = dict(input=str(path.resolve()), sha256=hashlib.sha256(path.read_bytes()).hexdigest(),
                    sheets=book.sheet_names(), active_tab=active, formula_count=2048,
                    expected=spec["expected"] if "expected" in spec else None)
    (root / "fixture.json").write_text(json.dumps(manifest, indent=2), "utf-8")
    print(json.dumps({k: v for k, v in manifest.items() if k != "expected"}))


def verify(root):
    manifest = json.loads((root / "fixture.json").read_text("utf-8"))
    assert hashlib.sha256(Path(manifest["input"]).read_bytes()).hexdigest() == manifest["sha256"]
    reports = []
    for op in ("calculate", "modify", "calculate-reloaded", "modify-reloaded"):
        path = root / "outputs" / (op + ".xlsx")
        ns = {"s": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
        with zipfile.ZipFile(path) as z:
            assert z.testzip() is None
            xml = ET.fromstring(z.read("xl/workbook.xml"))
            names = [s.attrib["name"] for s in xml.findall("s:sheets/s:sheet", ns)]
            active = int(xml.find("s:bookViews/s:workbookView", ns).attrib.get("activeTab", "0"))
        wb = openpyxl.load_workbook(path, data_only=False)
        caches = openpyxl.load_workbook(path, data_only=True)
        try:
            formulas = sum(c.data_type == "f" for row in wb["Calc"] for c in row)
            target = not op.startswith("modify") or (wb["Data"]["D12"].value == 9999 and wb["Data"]["D12"].font.bold)
            populated = sum(c.value is not None for row in caches["Calc"] for c in row)
            actual = [c.value for row in caches["Calc"] for c in row]
            expected = manifest["expected"]
            calc_matches = None if not op.startswith("calculate") else (len(actual) == len(expected) and all(
                isinstance(a, (int, float)) and not isinstance(a, bool) and abs(a - b) <= 1e-8
                for a, b in zip(actual, expected)))
            reports.append(dict(operation=op, zip_valid=True, sheets=names, active_tab=active,
                                formula_count=formulas, nonempty_formula_results=populated,
                                calculation_matches_expected=calc_matches,
                                target_correct=bool(target), clean_output=names == manifest["sheets"] and active == 1))
        finally:
            wb.close()
            caches.close()
    (root / "independent-verification.json").write_text(json.dumps(reports, indent=2), "utf-8")
    print(json.dumps(reports, indent=2))
    return 0 if all(r["clean_output"] and r["target_correct"] and r["calculation_matches_expected"] is not False
                    and r["formula_count"] == 2048 for r in reports) else 2


if __name__ == "__main__":
    path = Path(sys.argv[2]).resolve()
    if sys.argv[1] == "fixture":
        fixture(path)
    else:
        sys.exit(verify(path))
