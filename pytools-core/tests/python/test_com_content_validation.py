"""COM content-only validation: positive drift and negative corruption controls.

Pure snapshot cases need no Office. Opt-in real COM uses synthetic copies only.
"""
import copy
import hashlib
import importlib.util
import json
import os
from pathlib import Path
import subprocess
import tempfile
import unittest
from unittest.mock import Mock, patch

import openpyxl

ROOT = Path(__file__).resolve().parents[3]
spec = importlib.util.spec_from_file_location(
    "com_content_worker", ROOT / "pytools-core/python/pywin32_excel_conversion_worker.py")
worker = importlib.util.module_from_spec(spec)
spec.loader.exec_module(worker)


def snapshot():
    return [{"name": "Data", "formulas": {(12, 27): '=IF(A10="old",1,2)'},
             "formula_values": {(12, 27): 1.0}, "used_cells": 324,
             "values": {(1, 1): "header", (10, 1): "old", (11, 2): "001",
                        (12, 2): 0.0, (13, 2): False}}]


class SavedCalculation(unittest.TestCase):
    def test_host_can_skip_com_range_read_without_skipping_calculation_save(self):
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "validation-copy.xlsx"
            fixture = openpyxl.Workbook()
            fixture.save(path)
            fixture.close()
            original = path.read_bytes()
            session = worker.ConversionSession.__new__(worker.ConversionSession)
            session._application = Mock()
            session._application_name = "Excel.Application"
            session._process_id = None
            session._ensure_application = Mock()
            session._reset_application = Mock()
            workbook = Mock()
            session._open_workbook = Mock(return_value=workbook)
            regions = [{"template_sheet": "Sheet", "source_sheet": "Sheet", "address": "A1"}]

            with patch.object(worker, "_force_recalc", return_value=("CalculateFull", [])) as calc:
                response = session.recalculate_named_validation(str(path), regions, collect_cells=False)

            self.assertTrue(response["ok"], response)
            self.assertEqual(response["cells"], [])
            calc.assert_called_once_with(session._application)
            workbook.Save.assert_called_once_with()
            workbook.Close.assert_called_once_with(SaveChanges=False)
            workbook.Worksheets.assert_not_called()
            self.assertEqual(path.read_bytes(), original)


class ContentValidation(unittest.TestCase):
    def setUp(self):
        self.before = snapshot()
        self.after = copy.deepcopy(self.before)
        self.after[0]["values"][(10, 1)] = "new"
        self.updates = {("Data", (10, 1)): "new"}

    def verify(self):
        return worker._verify_institution_normalize_com_snapshots(
            self.before, self.after, self.updates)

    def test_cache_refresh_and_format_extent_are_not_content_changes(self):
        self.after[0]["formula_values"][(12, 27)] = 2.0
        self.after[0]["used_cells"] = 351
        self.assertEqual(self.verify(), {"worksheet_count": 1, "formula_cells": 1})

    def test_absent_or_error_cache_is_not_formula_loss(self):
        for cache in ({}, {(12, 27): "#VALUE!"}):
            with self.subTest(cache=cache):
                self.after[0]["formula_values"] = cache
                self.verify()

    def test_formula_text_change_rejected(self):
        self.after[0]["formulas"][(12, 27)] = "=2"
        with self.assertRaisesRegex(ValueError, "公式文本"):
            self.verify()

    def test_formula_address_change_or_removal_rejected(self):
        for formulas in ({}, {(13, 27): '=IF(A10="old",1,2)'}):
            with self.subTest(formulas=formulas):
                self.after[0]["formulas"] = formulas
                with self.assertRaisesRegex(ValueError, "公式文本"):
                    self.verify()

    def test_literal_change_removal_and_addition_rejected(self):
        for action in ("change", "remove", "add"):
            with self.subTest(action=action):
                self.after[0]["values"] = copy.deepcopy(self.before[0]["values"])
                self.after[0]["values"][(10, 1)] = "new"
                if action == "change": self.after[0]["values"][(1, 1)] = "wrong"
                if action == "remove": del self.after[0]["values"][(1, 1)]
                if action == "add": self.after[0]["values"][(99, 27)] = "unexpected"
                with self.assertRaisesRegex(ValueError, "非目标单元格"):
                    self.verify()

    def test_literal_type_change_rejected(self):
        self.after[0]["values"][(12, 2)] = "0"
        with self.assertRaisesRegex(ValueError, "值发生变化"):
            self.verify()

    def test_missing_requested_update_rejected(self):
        self.after[0]["values"][(10, 1)] = "old"
        with self.assertRaisesRegex(ValueError, "未保存预期文本"):
            self.verify()

    def test_sheet_rename_or_removal_rejected(self):
        self.after[0]["name"] = "Other"
        with self.assertRaisesRegex(ValueError, "名称或顺序"):
            self.verify()


@unittest.skipUnless(os.environ.get("RUSTTOOLS_COM_CONTENT_LIVE") == "1",
                     "requires explicit synthetic Windows COM run")
class RealXlsContent(unittest.TestCase):
    def setUp(self):
        import pythoncom
        # Each case owns its COM apartment; ConversionSession has its own
        # balanced initialization and must not define the next test's state.
        pythoncom.CoInitialize()
        self.addCleanup(pythoncom.CoUninitialize)

    def test_rust_xls_business_routes_match_xlsx_results(self):
        import openpyxl
        import win32com.client
        exe = os.environ["RUSTTOOLS_TEST_EXE"]
        spec = importlib.util.spec_from_file_location(
            "com_menu_fixtures", ROOT / "tools/AsposeWorker/test_menu_matrix.py")
        helper = importlib.util.module_from_spec(spec)
        spec.loader.exec_module(helper)
        root = Path(os.environ["RUSTTOOLS_COM_CONTENT_OUTPUT"]).with_name("com-rust-business")
        root.mkdir(parents=True, exist_ok=False)
        app = win32com.client.DispatchEx("Excel.Application")
        app.Visible = False
        app.DisplayAlerts = False

        def view(path):
            wb = app.Workbooks.Open(str(path), UpdateLinks=0, ReadOnly=True)
            try:
                # Values/formulas, without appearance or formula caches: compare
                # occupied cells rather than formatting-only UsedRange extents.
                return {s["name"]: (s["values"], s["formulas"]) for s in worker._snapshot(wb)}
            finally: wb.Close(False)

        records = []
        try:
            for case in ("5-1", "5-2", "5-7"):
                with self.subTest(case=case):
                    outputs = []
                    for legacy in (False, True):
                        directory = root / case / ("xls" if legacy else "xlsx")
                        command, stdin, expected, readonly = helper.prepare(directory, case, False)
                        if case == "5-2":
                            source = directory / "source.xlsx"
                            b = openpyxl.load_workbook(source)
                            b.active["AB12"] = '=IF(A2="ABC Bank",1,2)'
                            b.save(source)
                            b.close()
                        if legacy:
                            for path in directory.glob("*.xlsx"):
                                if path.name == "config.xlsx": continue
                                destination = path.with_suffix(".xls")
                                wb = app.Workbooks.Open(str(path), UpdateLinks=0)
                                try: wb.SaveAs(str(destination), FileFormat=56)
                                finally: wb.Close(False)
                                command = [arg.replace(str(path), str(destination)) for arg in command]
                                if stdin: stdin = stdin.replace(str(path), str(destination))
                                expected = [destination if p == path else p for p in expected]
                                readonly = [destination if p == path else p for p in readonly]
                        config = directory / "config.xlsx"
                        hashes = {p: hashlib.sha256(p.read_bytes()).hexdigest() for p in [config, *readonly]}
                        output_hashes = {p: hashlib.sha256(p.read_bytes()).hexdigest() for p in expected}
                        env = {**os.environ, "PYTOOLS_COM_PYTHON": os.sys.executable,
                               "PYTOOLS_COM_WORKER": str(ROOT / "pytools-core/python/pywin32_excel_conversion_worker.py"),
                               "PYTOOLS_OPENPYXL_PYTHON": os.sys.executable,
                               "PYTOOLS_OPENPYXL_WORKER": str(ROOT / "crates/excel-editor/python/openpyxl_worker.py"),
                               "LOCALAPPDATA": str(directory / "state"), "XDG_STATE_HOME": str(directory / "state"),
                               "PYTHONIOENCODING": "utf-8"}
                        result = subprocess.run([exe, *command], input=stdin, cwd=directory, env=env,
                                                capture_output=True, text=True, encoding="utf-8", errors="replace", timeout=180)
                        (directory / "stdout.txt").write_text(result.stdout, encoding="utf-8")
                        (directory / "stderr.txt").write_text(result.stderr, encoding="utf-8")
                        self.assertEqual(result.returncode, 0, result.stdout + result.stderr)
                        self.assertNotIn("执行失败", result.stdout + result.stderr)
                        for path, original in hashes.items():
                            self.assertEqual(hashlib.sha256(path.read_bytes()).hexdigest(), original, str(path))
                        for path, original in output_hashes.items():
                            self.assertNotEqual(hashlib.sha256(path.read_bytes()).hexdigest(), original, "business operation did not commit: " + result.stdout)
                        outputs.append([view(p) for p in expected])
                    self.assertEqual(outputs[0], outputs[1], case)
                    if case == "5-2":
                        values, formulas = outputs[1][0]["甲分机构本外币"]
                        self.assertEqual(values[(2, 1)], "标准机构")
                        self.assertEqual(formulas[(12, 28)], '=IF(A2="ABC Bank",1,2)')
                    if case == "5-7":
                        self.assertEqual(outputs[1][0]["惠州市涉农贷款分机构"][0][(3, 4)], 5.125)
                    records.append({"case": case, "status": "PASS", "comparison": "XLS COM vs XLSX business values/formulas"})
        finally:
            app.Quit()
            (root / "results.json").write_text(json.dumps(records, ensure_ascii=False, indent=2), encoding="utf-8")

    def test_normalize_rename_convert_and_append_refusal(self):
        import openpyxl
        import win32com.client
        root = Path(os.environ["RUSTTOOLS_COM_CONTENT_OUTPUT"]).resolve()
        root.mkdir(parents=True, exist_ok=False)
        xlsx, source, candidate = root / "fixture.xlsx", root / "source.xls", root / "candidate.xls"
        book = openpyxl.Workbook()
        sheet = book.active
        sheet.title = "Data"
        for address, value in {"A1": "header", "A10": "old", "B11": "001",
                               "B12": 0, "B13": False, "AA12": '=IF(A10="old",1,2)'}.items():
            sheet[address] = value
        book.save(xlsx)
        app = win32com.client.DispatchEx("Excel.Application")
        try:
            app.Visible = False
            app.DisplayAlerts = False
            wb = app.Workbooks.Open(str(xlsx), UpdateLinks=0)
            try: wb.SaveAs(str(source), FileFormat=56)
            finally: wb.Close(False)
        finally: app.Quit()
        original = source.read_bytes()
        candidate.write_bytes(original)
        session = worker.ConversionSession()
        outcomes = {}
        try:
            outcomes["5-2"] = session.institution_normalize_com(str(candidate), [
                {"sheet": "Data", "cell": "A10", "value": "new"}])
            self.assertTrue(outcomes["5-2"].get("ok"), outcomes)
            wb = session._open_workbook(candidate, read_only=True, label="synthetic verification")
            try:
                sh = wb.Worksheets("Data")
                self.assertEqual(sh.Range("A10").Value2, "new")
                self.assertEqual(sh.Range("B11").Value2, "001")
                self.assertEqual(sh.Range("AA12").Formula, '=IF(A10="old",1,2)')
                self.assertEqual(sh.Range("A1").Value2, "header")
            finally: wb.Close(False)
            before_refusal = hashlib.sha256(candidate.read_bytes()).hexdigest()
            with self.assertRaisesRegex(ValueError, "XLS append disabled"):
                session.append_rows_com(str(candidate), "Data", 14, [["unsafe"]])
            self.assertEqual(hashlib.sha256(candidate.read_bytes()).hexdigest(), before_refusal)
            outcomes["4-5"] = session.rename_sheets_com(str(candidate), [{"old_name": "Data", "new_name": "Renamed"}])
            self.assertTrue(outcomes["4-5"].get("ok"), outcomes)
            output = root / "converted.xlsx"
            outcomes["4-4"] = session.convert(str(candidate), str(output))
            self.assertTrue(outcomes["4-4"].get("ok"), outcomes)
            result = openpyxl.load_workbook(output, data_only=False)
            try:
                self.assertIn("Renamed", result.sheetnames)
                self.assertEqual(result["Renamed"]["A10"].value, "new")
                self.assertEqual(result["Renamed"]["AA12"].value, '=IF(A10="old",1,2)')
            finally: result.close()
            self.assertEqual(source.read_bytes(), original)
        finally:
            session.close()
            (root / "results.json").write_text(json.dumps(outcomes, ensure_ascii=False, indent=2), encoding="utf-8")


if __name__ == "__main__":
    unittest.main(verbosity=2)
