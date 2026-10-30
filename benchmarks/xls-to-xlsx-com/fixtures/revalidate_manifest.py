"""Re-read the generated BIFF8 files and anchor expected metadata to saved XLS."""

from __future__ import annotations

import argparse
import gc
import hashlib
import json
import os
import time
from pathlib import Path

import pythoncom
import win32api
import win32con
import win32event
import win32process
import win32com.client


HERE = Path(__file__).resolve().parent
DATASET = HERE / "generated"
MANIFEST = DATASET / "manifest.json"
XL_CALCULATION_MANUAL = -4135
XL_FILE_FORMAT_XLS = 56
OLE_SIGNATURE = bytes.fromhex("D0CF11E0A1B11AE1")


def sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for block in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(block)
    return digest.hexdigest()


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--excel-progid", default="Excel.Application")
    args = parser.parse_args()
    manifest = json.loads(MANIFEST.read_text(encoding="utf-8"))
    if manifest.get("total_files") != 100 or len(manifest.get("items", [])) != 100:
        raise RuntimeError("Expected the fixed 100-file synthetic dataset")

    original_hashes = {}
    for item in manifest["items"]:
        path = DATASET / item["name"]
        if path.parent.resolve() != DATASET.resolve() or path.suffix.lower() != ".xls":
            raise RuntimeError(f"Unsafe or unexpected dataset entry: {item['name']}")
        actual_hash = sha256(path)
        if actual_hash != item["sha256"]:
            raise RuntimeError(f"Source XLS hash changed before metadata revalidation: {item['name']}")
        if path.read_bytes()[:8] != OLE_SIGNATURE:
            raise RuntimeError(f"Source is not an OLE/BIFF XLS: {item['name']}")
        original_hashes[item["name"]] = actual_hash

    before = set(win32process.EnumProcesses())
    pythoncom.CoInitialize()
    app = None
    excel_pid = None
    excel_handle = None
    owned = False
    quit_error = None
    started = time.perf_counter()
    try:
        app = win32com.client.DispatchEx(args.excel_progid)
        hwnd = int(app.Hwnd)
        _, pid = win32process.GetWindowThreadProcessId(hwnd)
        if not pid or pid in before or app.Workbooks.Count != 0:
            raise RuntimeError(
                f"Excel PID {pid} is not a new empty instance; refusing to use or close it"
            )
        excel_pid = int(pid)
        excel_handle = win32api.OpenProcess(win32con.SYNCHRONIZE, False, pid)
        owned = True
        app.Visible = False
        app.DisplayAlerts = False
        app.EnableEvents = False
        app.ScreenUpdating = False
        configuration_book = app.Workbooks.Add()
        try:
            app.Calculation = XL_CALCULATION_MANUAL
            app.CalculateBeforeSave = False
            if int(app.Calculation) != XL_CALCULATION_MANUAL:
                raise RuntimeError("Excel did not enter manual calculation mode")
        finally:
            configuration_book.Close(False)
            configuration_book = None

        for file_index, item in enumerate(manifest["items"], 1):
            path = DATASET / item["name"]
            book = None
            try:
                book = app.Workbooks.Open(str(path), 0, True)
                if int(book.FileFormat) != XL_FILE_FORMAT_XLS:
                    raise RuntimeError(f"{item['name']}: unexpected FileFormat {book.FileFormat}")
                expected_sheets = item["sheets"]
                names = [book.Worksheets(i).Name for i in range(1, book.Worksheets.Count + 1)]
                if names != [sheet["name"] for sheet in expected_sheets]:
                    raise RuntimeError(f"{item['name']}: saved XLS sheet names/order changed")
                if str(book.Names.Item(item["defined_name"]).RefersTo).find("$A$4:$D$9") < 0:
                    raise RuntimeError(f"{item['name']}: saved XLS named range changed")
                visibility = {"visible": -1, "hidden": 0, "veryHidden": 2}
                for sheet_index, sheet_expected in enumerate(expected_sheets, 1):
                    sheet = book.Worksheets(sheet_index)
                    if int(sheet.Visible) != visibility[sheet_expected["hidden"]]:
                        raise RuntimeError(f"{item['name']}/{sheet.Name}: visibility changed")
                    if str(sheet.Range("A1").MergeArea.Address).replace("$", "") != "A1:C1":
                        raise RuntimeError(f"{item['name']}/{sheet.Name}: merged title changed")
                    formula = str(sheet.Range(sheet_expected["formula_cell"]).Formula)
                    if formula != sheet_expected["formula"]:
                        raise RuntimeError(f"{item['name']}/{sheet.Name}: formula text changed")
                    if not bool(sheet.Range("A1").Font.Bold) or int(sheet.Range("A1").Interior.Color) == -4142:
                        raise RuntimeError(f"{item['name']}/{sheet.Name}: title style missing")
                    if not bool(sheet.Range("A2").Font.Bold):
                        raise RuntimeError(f"{item['name']}/{sheet.Name}: header style missing")
                    for cell, value in sheet_expected["key_values"].items():
                        actual = sheet.Range(cell).Value2
                        if actual != value:
                            raise RuntimeError(f"{item['name']}/{sheet.Name}: key value {cell} changed")

                    sheet_expected["row_heights"] = {
                        row: float(sheet.Rows(int(row)).RowHeight)
                        for row in sheet_expected["row_heights"]
                    }
                    sheet_expected["column_widths"] = {
                        column: float(sheet.Columns(column).ColumnWidth)
                        for column in sheet_expected["column_widths"]
                    }
                    for cell, style in sheet_expected["style_checks"].items():
                        if "number_format" in style:
                            if cell == "D4":
                                style.pop("number_format", None)
                                style["custom_number_format"] = True
                            else:
                                style["number_format"] = str(sheet.Range(cell).NumberFormat)
                book.Close(False)
                book = None
            finally:
                if book is not None:
                    book.Close(False)
                    book = None
                gc.collect()
            if file_index % 10 == 0:
                print(f"Re-read saved XLS metadata {file_index}/100", flush=True)
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
                    f"Metadata reader Excel PID {excel_pid} did not exit after Quit; not force-terminated"
                )
        if quit_error is not None:
            raise RuntimeError(f"Could not quit metadata reader Excel PID {excel_pid}: {quit_error}")

    for item in manifest["items"]:
        if sha256(DATASET / item["name"]) != original_hashes[item["name"]]:
            raise RuntimeError(f"Source XLS was modified during read-only metadata validation: {item['name']}")
    manifest["metadata_reference"] = "reopened saved BIFF8 XLS through Excel COM, read-only"
    manifest["metadata_revalidated_at_local"] = time.strftime("%Y-%m-%d %H:%M:%S")
    manifest["metadata_revalidation_excel_pid"] = excel_pid
    manifest["metadata_revalidation_elapsed_seconds"] = round(time.perf_counter() - started, 3)
    temporary = MANIFEST.with_suffix(".json.tmp")
    temporary.write_text(json.dumps(manifest, indent=2, ensure_ascii=False) + "\n", encoding="utf-8")
    os.replace(temporary, MANIFEST)
    print(json.dumps({"files": 100, "excel_pid": excel_pid, "manifest": str(MANIFEST), "manifest_sha256": sha256(MANIFEST)}, ensure_ascii=False), flush=True)


if __name__ == "__main__":
    main()
