"""Generate a fixed, synthetic BIFF8 XLS workload using local Excel COM."""

from __future__ import annotations

import argparse
import atexit
import gc
import hashlib
import json
import os
import shutil
import sys
import tempfile
import time
from pathlib import Path

import openpyxl
import pythoncom
import win32api
import win32con
import win32event
import win32process
import win32com.client
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.workbook.properties import CalcProperties


HERE = Path(__file__).resolve().parent
DATASET = HERE / "generated"
OLE_SIGNATURE = bytes.fromhex("D0CF11E0A1B11AE1")
XL_FILE_FORMAT_XLS = 56
XL_FILE_FORMAT_XLSX = 51
XL_CALCULATION_MANUAL = -4135


def sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for block in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(block)
    return digest.hexdigest()


def remove_generated_temp_tree(path: Path) -> None:
    if not path.exists():
        return
    parent = Path(tempfile.gettempdir()).resolve()
    target = path.resolve()
    if (
        target.parent != parent
        or not target.name.startswith("rust-tools-xls-com-fixtures-")
    ):
        raise RuntimeError(f"Refusing to remove an unexpected temporary path: {target}")
    shutil.rmtree(target)


def profile(group: str, ordinal: int) -> tuple[int, int, int]:
    if group == "small":
        return 24 + (ordinal * 17) % 73, 5 + (ordinal * 3) % 8, 1 + ordinal % 3
    if group == "medium":
        return 180 + (ordinal * 97) % 501, 10 + (ordinal * 5) % 15, 2 + ordinal % 3
    return 1200 + (ordinal * 137) % 1401, 16 + (ordinal * 7) % 25, 3 + ordinal % 3


def make_xlsx(path: Path, group: str, ordinal: int, file_index: int) -> dict:
    rows, columns, sheet_count = profile(group, ordinal)
    workbook = Workbook()
    workbook.calculation = CalcProperties(
        calcMode="manual", fullCalcOnLoad=False, forceFullCalc=False
    )
    workbook.remove(workbook.active)

    sheet_expectations = []
    column_widths = {
        get_column_letter(column): 11 + column % 7
        for column in range(1, columns + 1)
    }
    for sheet_index in range(sheet_count):
        name = f"Data{sheet_index + 1:02d}"
        sheet = workbook.create_sheet(name)
        sheet.merge_cells("A1:C1")
        sheet["A1"] = f"SYNTHETIC {group.upper()} {file_index:03d} / {name}"
        sheet["A1"].font = Font(name="Arial", size=13, bold=True, color="FFFFFF")
        sheet["A1"].fill = PatternFill("solid", fgColor="1F4E78")
        sheet["A1"].alignment = Alignment(horizontal="center", vertical="center")
        sheet.row_dimensions[1].height = 28

        headers = [f"Field {column:02d}" for column in range(1, columns + 1)]
        for column, header in enumerate(headers, 1):
            cell = sheet.cell(2, column, header)
            cell.font = Font(name="Arial", bold=True, color="FFFFFF")
            cell.fill = PatternFill("solid", fgColor="4472C4")
            cell.alignment = Alignment(horizontal="center", wrap_text=True)
            cell.border = Border(bottom=Side(style="thin", color="1F1F1F"))
            sheet.column_dimensions[get_column_letter(column)].width = column_widths[
                get_column_letter(column)
            ]
        sheet.row_dimensions[2].height = 23

        sheet.cell(3, 1, "Formula total")
        sheet.cell(3, 1).font = Font(bold=True)
        sheet.cell(3, 4, f"=SUM(D4:D{rows + 3})")
        sheet.cell(3, 4).number_format = '#,##0.00;[Red](#,##0.00)'
        sheet.row_dimensions[3].height = 21

        for row in range(4, rows + 4):
            source_row = row - 3
            sheet.cell(row, 1, f"record-{file_index:03d}-{sheet_index + 1}-{source_row:05d}")
            sheet.cell(row, 2, file_index * 10000 + sheet_index * 1000 + source_row)
            sheet.cell(row, 3, f"category-{source_row % 11:02d}")
            sheet.cell(row, 4, round((source_row * 1.25 + file_index + sheet_index) / 3, 2))
            for column in range(5, columns + 1):
                if column % 3 == 0:
                    sheet.cell(row, column, f"text-{file_index:03d}-{row}-{column}")
                else:
                    sheet.cell(row, column, round((row * column + file_index) / 7, 3))
            sheet.cell(row, 2).number_format = "0"
            sheet.cell(row, 4).number_format = '#,##0.00;[Red](#,##0.00)'
            if row == 4:
                sheet.row_dimensions[row].height = 19
        sheet.freeze_panes = "A4"
        if sheet_index > 0 and (file_index + sheet_index) % 4 == 0:
            sheet.sheet_state = "veryHidden" if sheet_index % 2 else "hidden"

        sheet_expectations.append(
            {
                "name": name,
                "rows": rows + 3,
                "columns": columns,
                "merge": "A1:C1",
                "formula_cell": "D3",
                "formula": f"=SUM(D4:D{rows + 3})",
                "hidden": sheet.sheet_state,
                "row_heights": {"1": 28, "2": 23, "3": 21, "4": 19},
                "column_widths": column_widths,
                "key_values": {
                    "A1": sheet["A1"].value,
                    "A4": sheet["A4"].value,
                    "B4": sheet["B4"].value,
                    "C4": sheet["C4"].value,
                },
                "style_checks": {
                    "A1": {
                        "font_bold": bool(sheet["A1"].font.bold),
                        "font_color": sheet["A1"].font.color.rgb,
                        "fill_color": sheet["A1"].fill.fgColor.rgb,
                        "horizontal_alignment": sheet["A1"].alignment.horizontal,
                    },
                    "A2": {
                        "font_bold": bool(sheet["A2"].font.bold),
                        "fill_color": sheet["A2"].fill.fgColor.rgb,
                    },
                    "B4": {"number_format": sheet["B4"].number_format},
                    "D4": {"number_format": sheet["D4"].number_format},
                },
            }
        )

    name = f"SyntheticArea_{file_index:03d}"
    first_sheet = workbook.worksheets[0].title
    workbook.defined_names.add(
        DefinedName(name, attr_text=f"'{first_sheet}'!$A$4:$D$9")
    )
    workbook.save(path)
    workbook.close()
    return {
        "group": group,
        "sheets": sheet_expectations,
        "defined_name": name,
        "key_values": {
            "A1": f"SYNTHETIC {group.upper()} {file_index:03d} / Data01",
            "A4": f"record-{file_index:03d}-1-00001",
            "B4": file_index * 10000 + 1,
            "C4": "category-01",
        },
        "rows_per_sheet": rows + 3,
        "columns_per_sheet": columns,
    }


def assert_excel_instance_is_owned(app, before: set[int]) -> tuple[int, object]:
    hwnd = int(app.Hwnd)
    _, pid = win32process.GetWindowThreadProcessId(hwnd)
    if not pid or pid in before or app.Workbooks.Count != 0:
        raise RuntimeError(
            f"Excel PID {pid} is not a newly-created empty instance; refusing to use or close it"
        )
    handle = win32api.OpenProcess(win32con.SYNCHRONIZE, False, pid)
    return pid, handle


def convert_to_xls(app, xlsx_path: Path, xls_path: Path, expected: dict) -> None:
    book = None
    try:
        book = app.Workbooks.Open(str(xlsx_path), 0, True)
        book.SaveAs(str(xls_path), FileFormat=XL_FILE_FORMAT_XLS)
        if int(book.FileFormat) != XL_FILE_FORMAT_XLS:
            raise RuntimeError(f"{xls_path.name}: Excel did not save FileFormat 56")
        book.Close(False)
        book = None
        book = app.Workbooks.Open(str(xls_path), 0, True)
        if int(book.FileFormat) != XL_FILE_FORMAT_XLS:
            raise RuntimeError(f"{xls_path.name}: reopened file is not FileFormat 56")
        if [book.Worksheets(i).Name for i in range(1, book.Worksheets.Count + 1)] != [
            item["name"] for item in expected["sheets"]
        ]:
            raise RuntimeError(f"{xls_path.name}: sheet names changed during XLS creation")
        name_ref = str(book.Names.Item(expected["defined_name"]).RefersTo)
        if "Data01" not in name_ref or "$A$4:$D$9" not in name_ref:
            raise RuntimeError(f"{xls_path.name}: named range was lost")
        expected_visibility = {"visible": -1, "hidden": 0, "veryHidden": 2}
        for index, sheet_expected in enumerate(expected["sheets"], 1):
            sheet = book.Worksheets(index)
            if int(sheet.Visible) != expected_visibility[sheet_expected["hidden"]]:
                raise RuntimeError(f"{xls_path.name}/{sheet.Name}: hidden state changed")
            merged_range = str(sheet.Range("A1").MergeArea.Address).replace("$", "")
            if merged_range != "A1:C1":
                raise RuntimeError(f"{xls_path.name}/{sheet.Name}: merged title was lost")
            if sheet.Range(sheet_expected["formula_cell"]).Formula != sheet_expected["formula"]:
                raise RuntimeError(f"{xls_path.name}/{sheet.Name}: formula text changed")
            if not bool(sheet.Range("A1").Font.Bold):
                raise RuntimeError(f"{xls_path.name}/{sheet.Name}: title font formatting was lost")
            if int(sheet.Range("A1").Interior.Color) == -4142:
                raise RuntimeError(f"{xls_path.name}/{sheet.Name}: title fill formatting was lost")
            if not bool(sheet.Range("A2").Font.Bold):
                raise RuntimeError(f"{xls_path.name}/{sheet.Name}: header font formatting was lost")
            if str(sheet.Range("B4").NumberFormat) != "0":
                raise RuntimeError(f"{xls_path.name}/{sheet.Name}: integer number format changed")
            for cell, style in sheet_expected["style_checks"].items():
                if "number_format" in style:
                    if cell == "D4":
                        style.pop("number_format", None)
                        style["custom_number_format"] = True
                    else:
                        style["number_format"] = str(sheet.Range(cell).NumberFormat)
            xls_row_heights = {}
            for row in sheet_expected["row_heights"]:
                actual = float(sheet.Rows(int(row)).RowHeight)
                if actual <= 0:
                    raise RuntimeError(
                        f"{xls_path.name}/{sheet.Name}: invalid row {row} height ({actual})"
                    )
                xls_row_heights[row] = actual
            sheet_expected["row_heights"] = xls_row_heights
            xls_column_widths = {}
            for column in sheet_expected["column_widths"]:
                actual = float(sheet.Columns(column).ColumnWidth)
                if actual <= 0:
                    raise RuntimeError(
                        f"{xls_path.name}/{sheet.Name}: invalid column {column} width ({actual})"
                    )
                xls_column_widths[column] = actual
            sheet_expected["column_widths"] = xls_column_widths
        for cell, value in expected["key_values"].items():
            actual = book.Worksheets(1).Range(cell).Value2
            if actual != value:
                raise RuntimeError(f"{xls_path.name}: {cell} value changed ({actual!r})")
        if xls_path.read_bytes()[:8] != OLE_SIGNATURE:
            raise RuntimeError(f"{xls_path.name}: output is not an OLE/BIFF XLS file")
    finally:
        if book is not None:
            book.Close(False)
        book = None
        gc.collect()


def build_batches(items: list[dict]) -> dict[str, list[dict]]:
    by_group = {
        group: sorted(
            (item for item in items if item["group"] == group),
            key=lambda item: item["name"],
        )
        for group in ("small", "medium", "large")
    }
    quotas = {
        "B10": (5, 4, 1),
        "B25": (13, 9, 3),
        "B50": (25, 17, 8),
        "B100": (50, 35, 15),
    }
    batches = {}
    for batch, (small, medium, large) in quotas.items():
        selected = (
            by_group["small"][:small]
            + by_group["medium"][:medium]
            + by_group["large"][:large]
        )
        batches[batch] = sorted(selected, key=lambda item: item["name"])
    return batches


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--excel-progid", default="Excel.Application")
    args = parser.parse_args()
    if DATASET.exists() and any(DATASET.iterdir()):
        raise SystemExit(
            f"Refusing to overwrite existing fixed dataset: {DATASET}; move it manually before regenerating"
        )
    DATASET.mkdir(parents=True, exist_ok=True)
    temporary = Path(tempfile.mkdtemp(prefix="rust-tools-xls-com-fixtures-"))
    atexit.register(remove_generated_temp_tree, temporary)
    xlsx_dir = temporary / "xlsx-seeds"
    xlsx_dir.mkdir()
    expected_source_hashes = {}
    items = []
    file_index = 0
    for group, count in (("small", 50), ("medium", 35), ("large", 15)):
        for ordinal in range(1, count + 1):
            file_index += 1
            rows, columns, sheet_count = profile(group, ordinal)
            name = f"{group.upper()}-{ordinal:03d}"
            xlsx_path = xlsx_dir / f"{name}.xlsx"
            xls_path = DATASET / f"{name}.xls"
            expected = make_xlsx(xlsx_path, group, ordinal, file_index)
            expected["name"] = f"{name}.xls"
            expected["size_profile"] = {
                "data_rows": rows,
                "columns": columns,
                "sheets": sheet_count,
            }
            expected_source_hashes[name] = sha256(xlsx_path)
            items.append(expected)
            if file_index % 10 == 0:
                print(f"Prepared {file_index}/100 synthetic XLSX seeds", flush=True)

    before = set(win32process.EnumProcesses())
    pythoncom.CoInitialize()
    app = None
    owned = False
    excel_pid = None
    excel_handle = None
    quit_error = None
    created = time.perf_counter()
    try:
        app = win32com.client.DispatchEx(args.excel_progid)
        excel_pid, excel_handle = assert_excel_instance_is_owned(app, before)
        owned = True
        app.Visible = False
        app.DisplayAlerts = False
        app.EnableEvents = False
        app.ScreenUpdating = False
        configuration_book = None
        try:
            configuration_book = app.Workbooks.Add()
            app.Calculation = XL_CALCULATION_MANUAL
            app.CalculateBeforeSave = False
            if int(app.Calculation) != XL_CALCULATION_MANUAL:
                raise RuntimeError("Could not set Excel calculation mode to manual")
        finally:
            if configuration_book is not None:
                configuration_book.Close(False)
            configuration_book = None
        for item in items:
            stem = Path(item["name"]).stem
            convert_to_xls(app, xlsx_dir / f"{stem}.xlsx", DATASET / item["name"], item)
            if len(list(DATASET.glob("*.xls"))) % 10 == 0:
                print(
                    f"Saved and verified {len(list(DATASET.glob('*.xls')))}/100 BIFF8 XLS workbooks",
                    flush=True,
                )
    finally:
        if app is not None and owned:
            try:
                app.Quit()
            except Exception as error:
                quit_error = error
        app = None
        gc.collect()
        pythoncom.CoUninitialize()
        if excel_handle is not None:
            exit_status = win32event.WaitForSingleObject(excel_handle, 30000)
            win32api.CloseHandle(excel_handle)
            if exit_status != win32event.WAIT_OBJECT_0:
                raise RuntimeError(
                    f"Generator's owned Excel PID {excel_pid} did not exit after Quit; it was not force-terminated"
                )
        if quit_error is not None:
            raise RuntimeError(f"Could not quit the generator-owned Excel instance: {quit_error}")

    if len(list(DATASET.glob("*.xls"))) != 100:
        raise RuntimeError("Fixed dataset must contain exactly 100 XLS files")
    batches = build_batches(items)
    batch_dir = DATASET / "batches"
    batch_dir.mkdir()
    for batch, batch_items in batches.items():
        payload = {
            "batch": batch,
            "count": len(batch_items),
            "files": [
                {"name": item["name"], "sha256": sha256(DATASET / item["name"])}
                for item in batch_items
            ],
        }
        (batch_dir / f"{batch}.json").write_text(
            json.dumps(payload, indent=2, ensure_ascii=False) + "\n", encoding="utf-8"
        )

    manifest = {
        "schema_version": 1,
        "created_at_local": time.strftime("%Y-%m-%d %H:%M:%S"),
        "source_generator": "benchmarks/xls-to-xlsx-com/fixtures/generate_dataset.py",
        "generator_excel_progid": args.excel_progid,
        "generator_excel_startup_elapsed_seconds": round(time.perf_counter() - created, 3),
        "python_version": sys.version.split()[0],
        "openpyxl_version": openpyxl.__version__,
        "groups": {"small": 50, "medium": 35, "large": 15},
        "total_files": 100,
        "signature": "D0 CF 11 E0 A1 B1 1A E1 (OLE compound XLS container)",
        "xl_file_format": XL_FILE_FORMAT_XLS,
        "items": [
            {
                **item,
                "sha256": sha256(DATASET / item["name"]),
                "bytes": (DATASET / item["name"]).stat().st_size,
            }
            for item in items
        ],
        "batches": {
            batch: {
                "count": len(batch_items),
                "manifest": f"batches/{batch}.json",
                "sha256": sha256(batch_dir / f"{batch}.json"),
            }
            for batch, batch_items in batches.items()
        },
        "xlsx_seed_hashes": expected_source_hashes,
    }
    (DATASET / "manifest.json").write_text(
        json.dumps(manifest, indent=2, ensure_ascii=False) + "\n", encoding="utf-8"
    )
    print(
        json.dumps(
            {
                "dataset": str(DATASET),
                "files": len(items),
                "batches": {name: len(group) for name, group in batches.items()},
                "excel_pid": excel_pid,
                "sha256_manifest": sha256(DATASET / "manifest.json"),
            },
            ensure_ascii=False,
        ),
        flush=True,
    )


if __name__ == "__main__":
    main()
