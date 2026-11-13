"""Explicit Excel COM checks on synthetic files, using the current debug CLI.

Set RUSTTOOLS_OFFICE_PROGID=Excel.Application. WPS uses test_wps_readonly.py
instead: its server lifecycle must not close a possibly reused application.
"""
import hashlib
import importlib.util
import os
from pathlib import Path
import shutil
import sys
import tempfile
import unittest

import openpyxl

ROOT = Path(__file__).resolve().parents[3]
spec = importlib.util.spec_from_file_location(
    "boundaries", Path(__file__).with_name("test_general_boundaries.py")
)
boundaries = importlib.util.module_from_spec(spec)
spec.loader.exec_module(boundaries)


def digest(path):
    return hashlib.sha256(path.read_bytes()).hexdigest()


@unittest.skipUnless(sys.platform == "win32" and os.environ.get("RUSTTOOLS_OFFICE_PROGID") == "Excel.Application",
                     "requires explicit Excel COM authorization; WPS must use the read-only harness")
class OfficeValidation(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        import win32com.client
        cls.app = win32com.client.DispatchEx(os.environ["RUSTTOOLS_OFFICE_PROGID"])
        # Some COM servers may reuse a running application. Do not touch it.
        if cls.app.Workbooks.Count:
            cls.app = None
            raise RuntimeError("COM server has existing workbooks; refusing to change or close them")
        cls.app.Visible = False
        cls.app.DisplayAlerts = False
        print("Office COM:", os.environ["RUSTTOOLS_OFFICE_PROGID"],
              "version", cls.app.Version, flush=True)

    @classmethod
    def tearDownClass(cls):
        if cls.app is not None:
            cls.app.Quit()
            cls.app = None

    def setUp(self):
        self.assertEqual(self.app.Workbooks.Count, 0, "previous test left a workbook open")

    def test_xlsx_append_and_repeat_reopen_with_office(self):
        with tempfile.TemporaryDirectory() as folder:
            d = Path(folder)
            source = d / "source.xlsx"
            target = d / "target.xlsx"
            workbook = openpyxl.Workbook()
            workbook.active.title = "Data"
            workbook.active.append(["ID", "Amount"])
            workbook.active.append(["001", 10])
            workbook.active.append(["002", 20])
            workbook.save(source)
            workbook = openpyxl.Workbook()
            workbook.active.title = "Combined"
            workbook.active.append(["ID", "Amount"])
            workbook.active.append(["001", 99])
            workbook.save(target)
            case = boundaries.Boundaries()
            config = case.config(d, target)
            before = {p: digest(p) for p in (source, config)}
            for attempt in range(2):
                text = case.cli(config, "2\n3\n0\n", "office-xlsx-append-" + str(attempt))
                self.assertNotIn("执行失败", text)
                output_before = digest(target)
                book = self.app.Workbooks.Open(str(target), UpdateLinks=0, ReadOnly=True)
                try:
                    sheet = book.Worksheets("Combined")
                    self.assertEqual(sheet.Range("A1:B3").Value,
                                     (("ID", "Amount"), ("001", 99), ("002", 20)))
                    self.assertIsNone(sheet.Cells(4, 1).Value)
                finally:
                    book.Close(False)
                self.assertEqual(digest(target), output_before)
            for path, expected in before.items():
                self.assertEqual(digest(path), expected)

    def test_real_xls_refusal_preserves_workbook(self):
        with tempfile.TemporaryDirectory() as folder:
            d = Path(folder)
            target = d / "target.xls"
            source = d / "source.xlsx"
            book = self.app.Workbooks.Add()
            try:
                sheet = book.Worksheets(1)
                sheet.Name = "Combined"
                headers = ["ID", "Amount"] + ["Field" + str(i) for i in range(3, 28)]
                for column, value in enumerate(headers, 1):
                    sheet.Cells(1, column).Value = value
                sheet.Cells(2, 1).NumberFormat = "@"
                sheet.Cells(2, 1).Value = "001"
                sheet.Cells(2, 2).Value = 99
                sheet.Cells(12, 27).Formula = "=B2+1"
                book.SaveAs(str(target), FileFormat=56)
            finally:
                book.Close(False)
            self.assertEqual(target.read_bytes()[:8], bytes.fromhex("d0cf11e0a1b11ae1"))
            workbook = openpyxl.Workbook()
            workbook.active.title = "Data"
            workbook.active.append(headers)
            workbook.active.append(["002", 20])
            workbook.save(source)
            if os.environ.get("RUSTTOOLS_OFFICE_FIXTURE_DIR"):
                fixtures = Path(os.environ["RUSTTOOLS_OFFICE_FIXTURE_DIR"])
                fixtures.mkdir(parents=True, exist_ok=True)
                shutil.copyfile(source, fixtures / "source.xlsx")
                shutil.copyfile(target, fixtures / "target.xls")
            case = boundaries.Boundaries()
            config = case.config(d, target)
            before = {p: digest(p) for p in (source, target, config)}
            text = case.cli(config, "2\n3\n0\n", "office-xls-refusal")
            self.assertIn("既有 XLS 去重追加已安全停用", text)
            for path, expected in before.items():
                self.assertEqual(digest(path), expected)
            self.assertFalse(any("candidate" in p.name.lower() for p in d.rglob("*")))
            book = self.app.Workbooks.Open(str(target), UpdateLinks=0, ReadOnly=True)
            try:
                sheet = book.Worksheets("Combined")
                self.assertEqual(sheet.Cells(1, 1).Value, "ID")
                self.assertEqual(sheet.Cells(2, 1).Value, "001")
                self.assertEqual(sheet.Cells(2, 2).Value, 99)
                self.assertEqual(sheet.Cells(12, 27).Formula.upper(), "=B2+1")
            finally:
                book.Close(False)
            self.assertEqual(digest(target), before[target])

    def test_print_modes_reopen_with_office(self):
        for mode in (1, 2, 3):
            with self.subTest(mode=mode), tempfile.TemporaryDirectory() as folder:
                d = Path(folder)
                source = d / "source.xlsx"
                target = d / "printed.xlsx"
                workbook = openpyxl.Workbook()
                first = workbook.active
                first.title = "first"
                first["A1"] = "old"
                first.merge_cells("A1:B1")
                first.row_dimensions[1].height = 70
                first.column_dimensions["A"].width = 50
                second = workbook.create_sheet("second")
                second.append(["left", "right"])
                second.append([0, "0"])
                second.row_dimensions[1].height = 22
                second.column_dimensions["A"].width = 17
                workbook.save(source)
                before = digest(source)
                case = boundaries.Boundaries()
                config = case.config(d)
                settings = openpyxl.load_workbook(config)
                rules = settings.create_sheet("打印配置")
                rules.append(["是否启用", "打印模式", "源工作簿", "源工作表", "源工作表打印区域",
                              "目标工作簿", "目标工作表", "FitToPagesWide", "FitToPagesTall",
                              "打印方向", "水平居中", "垂直居中", "不输出批注", "零值不输出"])
                for name in ("first", "second"):
                    rules.append(["Y", mode, str(source), name, "A1:B2", str(target), "same",
                                  1, 2, "横向", "Y", "N", "Y", "Y"])
                settings.save(config)
                text = case.cli(config, "3\n1\n0\n0\n", "office-print-mode-" + str(mode))
                self.assertNotIn("执行失败", text)
                output_before = digest(target)
                book = self.app.Workbooks.Open(str(target), UpdateLinks=0, ReadOnly=True)
                try:
                    sheet = book.Worksheets("same")
                    self.assertEqual(sheet.Cells(1, 1).Value, "left")
                    self.assertEqual(sheet.Cells(1, 2).Value, "right")
                    self.assertFalse(sheet.Range("A1:B1").MergeCells)
                    self.assertIsNone(sheet.Cells(2, 1).Value)
                    self.assertFalse(sheet.Cells(2, 1).HasFormula)
                    self.assertEqual(sheet.Cells(2, 2).Value, "0")
                    self.assertEqual(
                        sheet.Rows(1).RowHeight, 22 if mode == 1 else sheet.StandardHeight
                    )
                    # Excel/WPS may round width units differently; reject stale width.
                    self.assertLess(sheet.Columns(1).ColumnWidth, 25)
                    setup = sheet.PageSetup
                    self.assertEqual(setup.Orientation, 2)
                    self.assertEqual((setup.FitToPagesWide, setup.FitToPagesTall), (1, 2))
                    self.assertIn("$A$1:$B$3", setup.PrintArea)
                    self.assertIn("共 1 页", setup.CenterFooter)
                finally:
                    book.Close(False)
                self.assertEqual(digest(source), before)
                self.assertEqual(digest(target), output_before)


if __name__ == "__main__":
    unittest.main(verbosity=2)
