"""Generate the fixed-size synthetic workload for the gated phase-3 measurements."""
from __future__ import annotations

import hashlib
import json
from pathlib import Path

import xlwt
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill

ROOT = Path(__file__).resolve().parents[1] / "fixtures" / "performance"
ROOT.mkdir(parents=True, exist_ok=True)


def make_xlsx(path: Path, rows: int, columns: int, *, write_only: bool = False) -> None:
    wb = Workbook(write_only=write_only)
    ws = wb.create_sheet("Data") if write_only else wb.active
    ws.title = "Data"
    headers = ["Item", "Value", "Double", "Group", "Date", "Pct", "Code", "Notes", "Extra", "Calc"][:columns]
    ws.append(headers)
    if not write_only:
        for cell in ws[1]:
            cell.font = Font(bold=True, color="FFFFFF")
            cell.fill = PatternFill("solid", fgColor="17365D")
    for row in range(2, rows + 2):
        values = [f"row-{row - 1}", row, f"=B{row}*2", f"g{row % 7}", "2026-01-01", 0.125,
                  f"C{row:07d}", "synthetic", row % 11, f"=SUM(B{row}:C{row})"][:columns]
        ws.append(values)
    if not write_only:
        ws.freeze_panes = "A2"
        ws.print_area = f"A1:{chr(64 + columns)}{rows + 1}"
        ws.print_title_rows = "1:1"
        ws.page_setup.paperSize = ws.PAPERSIZE_A4
        ws.page_setup.orientation = "landscape"
        ws.oddFooter.right.text = "Page &P of &N"
    wb.save(path)


def make_xls(path: Path, rows: int) -> None:
    wb = xlwt.Workbook(encoding="utf-8")
    ws = wb.add_sheet("Data")
    header = xlwt.easyxf("font: bold on; pattern: pattern solid, fore_colour dark_blue;")
    for col, value in enumerate(["Item", "Value", "Double", "Group", "Date", "Pct", "Code", "Notes"]):
        ws.write(0, col, value, header)
    for row in range(1, rows + 1):
        ws.write(row, 0, f"row-{row}")
        ws.write(row, 1, row)
        ws.write(row, 2, xlwt.Formula(f"B{row + 1}*2"))
        ws.write(row, 3, f"g{row % 7}")
        ws.write(row, 4, xlwt.Formula("DATE(2026,1,1)"))
        ws.write(row, 5, 0.125)
        ws.write(row, 6, f"C{row:07d}")
        ws.write(row, 7, "synthetic")
    wb.save(str(path))


for i in range(1, 9):
    make_xls(ROOT / f"batch-{i:02d}.xls", 100)
for i in range(9, 25):
    make_xlsx(ROOT / f"batch-{i:02d}.xlsx", 100, 8)
make_xlsx(ROOT / "large-150k.xlsx", 150_000, 10)
make_xlsx(ROOT / "large-1m.xlsx", 1_000_000, 8, write_only=True)

manifest = []
for path in sorted(ROOT.iterdir()):
    digest = hashlib.sha256(path.read_bytes()).hexdigest()
    manifest.append({"name": path.name, "bytes": path.stat().st_size, "sha256": digest})
(ROOT / "manifest.json").write_text(json.dumps(manifest, indent=2), encoding="utf-8")
print(f"generated {len(manifest)} files in {ROOT}")
