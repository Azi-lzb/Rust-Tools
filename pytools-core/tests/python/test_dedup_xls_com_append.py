"""Run isolated 2-3 acceptance with the selected, installed Windows COM host."""

import datetime
import hashlib
import json
import os
from pathlib import Path
import runpy
import shutil
import struct
import subprocess
import tempfile
import unittest

import openpyxl


def checksum(path):
    return hashlib.sha256(path.read_bytes()).hexdigest()


@unittest.skipUnless(os.name == "nt" and os.environ.get("RUSTTOOLS_XLS_APPEND_CLI"),
                     "requires an explicitly selected Windows RustTools binary")
class XlsAppend(unittest.TestCase):
    def setUp(self):
        self.cli = Path(os.environ["RUSTTOOLS_XLS_APPEND_CLI"]).resolve()
        worker_path = os.environ.get("PYTOOLS_COM_WORKER") or self.cli.parent / "python" / "pywin32_excel_conversion_worker.py"
        self.session = runpy.run_path(str(worker_path))["ConversionSession"]()
        self.addCleanup(self.session.close)
        self.session._ensure_application()
        self.assertEqual(self.session._application_name, os.environ["RUSTTOOLS_XLS_APPEND_APPLICATION"])
        folder = tempfile.TemporaryDirectory(prefix="xls-append-")
        self.addCleanup(folder.cleanup)
        self.root = Path(folder.name)
        self.target = self.root / "target.xls"
        self.source = self.root / "source.xlsx"
        self.config = self.root / "config.xlsx"
        self.headers = ["ID", "Amount", "Active", "Gap", "Formula", "Day", "Literal"] + [
            f"Field{column}" for column in range(8, 28)
        ]
        book = self.session._application.Workbooks.Add()
        try:
            data = book.Worksheets(1)
            data.Name = "Combined"
            data.Range("A1:AA1").Value2 = (tuple(self.headers),)
            data.Range("A2:B11").Value2 = tuple((f"old-{row}", row * 10) for row in range(2, 12))
            data.Range("AA12").Formula = "=SUM(B2:B11)"
            data.Range("J3:K3").Merge()
            data.Range("J3").Value2 = "Keep this label"
            data.Range("B2").AddComment("Synthetic original note")
            book.Names.Add("OriginalAmount", "=Combined!$B$2")
            book.SaveAs(str(self.target), 56)
        finally:
            book.Close(False)
        source = openpyxl.Workbook()
        sheet = source.active
        sheet.title = "Data"
        sheet.append(self.headers)
        sheet.append(["00123", 20.5, True, None, "=B2+5", datetime.datetime(2026, 11, 18), "=literal"])
        sheet["G2"].data_type = "s"
        sheet["AA2"] = "far column"
        sheet.append(["old-2", 333])
        source.save(self.source)
        config = openpyxl.Workbook()
        active = config.active
        active.title = "全局配置"
        active.append(["配置项", "值"])
        for key, value in (("输入目录", "."), ("输出目录", "output"), ("日志目录", "logs")):
            active.append([key, value])
        tasks = config.create_sheet("去重追加数据配置")
        tasks.append(["是否启用", "任务名", "源数据工作簿", "源数据工作表",
                      "标识列序号", "目标工作簿", "目标工作表", "执行模式"])
        tasks.append(["Y", "Synthetic", str(self.source), "Data", "1", str(self.target), "Combined", "追加"])
        config.save(self.config)

    def call(self, *arguments, menu_input=None, expected_code=0, env=None):
        result = subprocess.run([str(self.cli), *map(str, arguments)], cwd=self.root,
                                capture_output=True, text=True, encoding="utf-8",
                                input=menu_input, timeout=120, env=env)
        self.assertEqual(result.returncode, expected_code, result.stdout + result.stderr)
        return result.stdout, result.stderr

    def build_plan(self):
        output, _ = self.call("dedup-append-plan", "--config", self.config, "--output-dir", self.root / "plans")
        entry = json.loads(output)[0]
        self.assertEqual((entry["appended_rows"], entry["duplicate_rows_filtered"]), (1, 1))
        return Path(entry["plan_file"])

    def inspect_append(self, path):
        book = self.session._open_workbook(path, read_only=True, label="核验合成 XLS")
        try:
            self.assertEqual(book.FileFormat, 56)
            self.assertEqual(book.ActiveSheet.Name, "Combined")
            sheet = book.Worksheets("Combined")
            self.assertEqual(sheet.Range("A1:AA1").Value2[0], tuple(self.headers))
            self.assertEqual(sheet.Range("A2:B11").Value2,
                             tuple((f"old-{row}", row * 10) for row in range(2, 12)))
            self.assertEqual(sheet.Range("AA12").Formula, "=SUM(B2:B11)")
            self.assertEqual(sheet.Range("J3").MergeArea.Address, "$J$3:$K$3")
            self.assertEqual(sheet.Range("J3").Value2, "Keep this label")
            self.assertEqual(sheet.Range("B2").Comment.Text(), "Synthetic original note")
            self.assertEqual(book.Names("OriginalAmount").RefersTo, "=Combined!$B$2")
            self.assertEqual(sheet.Range("A13:D13").Value2[0], ("00123", 20.5, True, None))
            self.assertEqual(sheet.Range("E13").Formula, "=B2+5")
            self.assertEqual(sheet.Range("F13").Value2, 46344)
            self.assertEqual(sheet.Range("G13").Value2, "=literal")
            self.assertFalse(sheet.Range("G13").HasFormula)
            self.assertEqual(sheet.Range("AA13").Value2, "far column")
        finally:
            book.Close(False)

    def test_preview_menu_commit_and_repeat(self):
        before = {path: checksum(path) for path in (self.target, self.source, self.config)}
        if folder := os.environ.get("RUSTTOOLS_XLS_APPEND_FIXTURE_DIR"):
            destination = Path(folder)
            destination.mkdir()
            for file in (self.source, self.target):
                shutil.copyfile(file, destination / file.name)
        plan_path = self.build_plan()
        output, _ = self.call("dedup-append-execute", "--plan", plan_path)
        report = json.loads(output)
        self.assertEqual(report["worker"]["application"], self.session._application_name)
        self.assertEqual(checksum(self.target), before[self.target])
        self.inspect_append(Path(report["candidate"]))
        output, _ = self.call("menu", "--config", self.config, "--log-view", "developer", menu_input="2\n3\n0\n")
        self.assertNotIn("执行失败", output)
        self.inspect_append(self.target)
        updated = checksum(self.target)
        self.assertNotEqual(updated, before[self.target])
        self.call("menu", "--config", self.config, "--log-view", "developer", menu_input="2\n3\n0\n")
        self.assertEqual(checksum(self.target), updated)
        for path in (self.source, self.config):
            self.assertEqual(checksum(path), before[path])

    def test_refuses_unsafe_plan_without_touching_target(self):
        before = checksum(self.target)
        plan_path = self.build_plan()
        for index, (address, value, expected) in enumerate([
            ("C4", {"type": "text", "value": "unsafe"}, "原有数据区域"),
            ("@12", {"type": "text", "value": "unsafe"}, "解析追加单元格地址失败"),
            ("IW13", {"type": "text", "value": "unsafe"}, "无效的 BIFF8 单元格地址"),
            ("A65537", {"type": "text", "value": "unsafe"}, "无效的 BIFF8 单元格地址"),
            ("A13", {"type": "error", "value": "#VALUE!"}, "不支持直接写入 Excel 错误值"),
            ("A13", {"type": "text", "value": ""}, "不支持将空字符串与空白区分"),
        ]):
            with self.subTest(index=index):
                plan = json.loads(plan_path.read_text(encoding="utf-8"))
                plan["xls_operations"] = [{"operation": "set_cell_value", "sheet": "Combined",
                                          "cell": address, "value": value}]
                plan["candidate"] = str(self.root / f"refusal-{index}.xls")
                invalid = self.root / "unsafe.json"
                invalid.write_text(json.dumps(plan), encoding="utf-8")
                _, error = self.call("dedup-append-execute", "--plan", invalid, "--commit-target", expected_code=2)
                self.assertIn(expected, error)
                self.assertEqual(checksum(self.target), before)
                self.assertFalse(Path(plan["candidate"]).exists())

    def test_adds_missing_sheet_after_existing_and_restores_active_sheet(self):
        settings = openpyxl.load_workbook(self.config)
        settings["去重追加数据配置"]["G2"] = "NewData"
        settings.save(self.config)
        original = checksum(self.target)
        output, _ = self.call("dedup-append-plan", "--config", self.config, "--output-dir", self.root / "plans")
        plan = json.loads(output)[0]
        self.assertEqual(plan["appended_rows"], 2)
        output, _ = self.call("dedup-append-execute", "--plan", plan["plan_file"])
        report = json.loads(output)
        self.assertEqual(report["worker"]["application"], self.session._application_name)
        book = self.session._open_workbook(Path(report["candidate"]), read_only=True, label="核验新增 Sheet")
        try:
            sheet_names = [sheet.Name for sheet in book.Worksheets]
            self.assertEqual(sheet_names[0], "Combined")
            self.assertEqual(sheet_names[-1], "NewData")
            self.assertEqual(book.ActiveSheet.Name, "Combined")
            self.assertEqual(book.Worksheets("NewData").Range("A1:B3").Value2,
                             (("ID", "Amount"), ("00123", 20.5), ("old-2", 333)))
            self.assertEqual(book.Worksheets("Combined").Range("AA12").Formula, "=SUM(B2:B11)")
        finally:
            book.Close(False)
        self.assertEqual(checksum(self.target), original)

    def test_failed_readback_saves_unpublished_review_copy_beside_target(self):
        plan = self.build_plan()
        original = checksum(self.target)
        unwritten_worker = self.root / "unwritten_worker.py"
        unwritten_worker.write_text(
            "import json, sys\n"
            "print(json.dumps({'ready': True, 'protocol': 3}), flush=True)\n"
            "for line in sys.stdin.buffer:\n"
            "    op = json.loads(line)['op']\n"
            "    print(json.dumps({'ok': True, 'verified': True}), flush=True)\n"
            "    if op == 'shutdown': break\n",
            encoding="utf-8",
        )
        worker_python = os.environ.get("PYTOOLS_COM_PYTHON") or str(self.cli.parent / "python" / "python.exe")
        environment = dict(os.environ, PYTOOLS_COM_WORKER=str(unwritten_worker),
                           PYTOOLS_COM_PYTHON=worker_python)
        _, error = self.call("dedup-append-execute", "--plan", plan, "--commit-target",
                             expected_code=2, env=environment)
        self.assertIn("待核对副本", error)
        review, = self.root.glob("target_待核对_业务校验不符_*.xls")
        self.assertEqual(checksum(review), original)
        self.assertEqual(checksum(self.target), original)
        record = json.loads(review.with_name(review.name + ".run.json").read_text(encoding="utf-8"))
        self.assertIn("回读不一致", record["validation_issue"])
        self.assertEqual(Path(record["review_copy"]), review)
        self.assertEqual(Path(record["planned_target"]), self.target)
        self.assertFalse(record["original_target_replaced"])
        self.assertFalse(Path(json.loads(plan.read_text(encoding="utf-8"))["candidate"]).exists())

    def test_append_preserves_filldown_formulas_and_ignores_updated_caches(self):
        book = self.session._open_workbook(self.target, read_only=False, label="共享公式夹具")
        try:
            sheet = book.Worksheets("Combined")
            sheet.Range("L2").Formula = "=COUNTA(A:A)"
            sheet.Range("L2:L11").FillDown()
            sheet.Range("M2:M4").FormulaArray = "=LEN(A2:A4)"
            sheet.Calculate()
            book.Save()
        finally:
            book.Close(False)

        # Confirm whether this host saved actual BIFF shared-formula records.
        import pythoncom
        storage = pythoncom.StgOpenStorage(str(self.target), None, 0x20)
        stream = storage.OpenStream("Workbook", None, 0x10, 0)
        data = stream.Read(1024 * 1024)
        shared_count, array_count, offset = 0, 0, 0
        while offset + 4 <= len(data):
            kind, length = struct.unpack_from("<HH", data, offset)
            shared_count += kind == 0x04BC
            array_count += kind == 0x0221
            offset += 4 + length
        del stream, storage
        print(f"{self.session._application_name} BIFF records: SHRFMLA={shared_count}, ARRAY={array_count}", flush=True)
        self.assertGreater(shared_count, 0)
        self.assertGreater(array_count, 0)
        before = checksum(self.target)
        self.call("menu", "--config", self.config, "--log-view", "developer", menu_input="2\n3\n0\n")
        self.assertNotEqual(checksum(self.target), before)
        self.inspect_append(self.target)
        book = self.session._open_workbook(self.target, read_only=True, label="共享公式验收")
        try:
            sheet = book.Worksheets("Combined")
            self.assertEqual(sheet.Range("L2:L11").Formula,
                             tuple(("=COUNTA(A:A)",) for _ in range(10)))
            self.assertTrue(sheet.Range("M2").HasArray)
            self.assertEqual(sheet.Range("M2").CurrentArray.Address, "$M$2:$M$4")
        finally:
            book.Close(False)

    def write_deletion_fixture(self, sheet_name):
        book = self.session._application.Workbooks.Add()
        path = self.root / "delete.xls"
        try:
            sheet = book.Worksheets(1)
            sheet.Name = sheet_name
            sheet.Range("A1:B6").Value2 = (
                ("ID", "Code"), ("A", "p"), ("A", "q"), ("B", "r"), ("C", "p"), ("D", "s"))
            sheet.Range("D2").Formula = "=LEN(A2)"
            sheet.Range("D6").Formula = "=LEN(A6)"
            sheet.Range("A100").Font.Bold = True
            book.SaveAs(str(path), 56)
        finally:
            book.Close(False)
        return path

    def test_dedup_delete_current_row_coordinates_and_failed_write_statistics(self):
        path = self.write_deletion_fixture("Rows")
        settings = openpyxl.load_workbook(self.config)
        tasks = settings["去重追加数据配置"]
        tasks.delete_rows(2)
        for key in ("1", "2"):
            tasks.append(["Y", "Synthetic", str(path), "Rows", key])
        settings.save(self.config)
        bad_worker = self.root / "rejected_worker.py"
        bad_worker.write_text(
            "import json, sys\n"
            "print(json.dumps({'ready': True, 'protocol': 3}), flush=True)\n"
            "for line in sys.stdin.buffer:\n"
            "    if json.loads(line)['op'] == 'shutdown': break\n"
            "    print(json.dumps({'ok': True, 'verified': False, 'saved': True,"
            "'validation_issue': 'Synthetic readback rejection'}), flush=True)\n",
            encoding="utf-8",
        )
        original = checksum(path)
        env = dict(os.environ, PYTOOLS_COM_WORKER=str(bad_worker))
        output, error = self.call("menu", "--config", self.config, "--log-view", "developer",
                                  menu_input="2\n2\n0\n", env=env)
        self.assertNotIn("删除重复行数：2", output)
        self.assertIn("待核对副本", output + error)
        self.assertEqual(checksum(path), original)
        review, = self.root.glob("delete_待核对_业务校验不符_*.xls")
        self.assertEqual(checksum(review), original)

        output, _ = self.call("menu", "--config", self.config, "--log-view", "developer", menu_input="2\n2\n0\n")
        self.assertIn("删除重复行数：2", output)
        self.assertNotEqual(checksum(path), original)
        book = self.session._open_workbook(path, read_only=True, label="核验删行")
        try:
            sheet = book.Worksheets("Rows")
            self.assertEqual(sheet.Range("A1:A4").Value2, (("ID",), ("A",), ("B",), ("D",)))
            self.assertEqual(sheet.Range("D4").Formula, "=LEN(A4)")
        finally:
            book.Close(False)

    def test_foreign_bank_delete_keeps_input_and_reopens_copy(self):
        name = "(2)本外币存贷款分机构表"
        path = self.write_deletion_fixture(name)
        settings = openpyxl.load_workbook(self.config)
        mapping = settings.create_sheet("机构映射表")
        mapping.append(["占位", "机构名称", "外资行标志"])
        mapping.append([None, "A", "1"])
        settings.save(self.config)
        original = checksum(path)
        output, _ = self.call("remove-foreign-bank", "--config", self.config, "--create-copy",
                              "--output-dir", self.root / "output", "--log-view", "off", path)
        self.assertEqual(json.loads(output)["deleted_rows"], 2)
        copy = path.with_stem("delete(金融局)")
        self.assertEqual(checksum(path), original)
        book = self.session._open_workbook(copy, read_only=True, label="核验外资行副本")
        try:
            sheet = book.Worksheets(name)
            self.assertEqual(sheet.Range("A1:A4").Value2, (("ID",), ("B",), ("C",), ("D",)))
            self.assertEqual(sheet.Range("D4").Formula, "=LEN(A4)")
        finally:
            book.Close(False)

    def test_validation_region_reports_formula_errors_without_aborting(self):
        from openpyxl.workbook.defined_name import DefinedName

        source = self.root / "validation.xlsx"
        template = self.root / "template.xlsx"
        workbook = openpyxl.Workbook()
        sheet = workbook.active
        sheet.title = "Data"
        sheet.append(["项目", "金额"])
        sheet.append(["甲", 11])
        sheet.append(["乙", 20])
        sheet["G20"] = "=SYNTHETIC_UNKNOWN_NAME"
        workbook.save(source)
        source_hash = checksum(source)

        rules = openpyxl.Workbook()
        sheet = rules.active
        sheet.title = "Data"
        sheet.append(["项目", "金额"])
        sheet["D2"] = '=IF(SUM(B2:B3)=30,"通过","校验失败：金额不平")'
        sheet["D3"] = '=IF(B2=11,"","校验失败：不应产生")'
        rules.defined_names.add(DefinedName("表结构区域", attr_text="'Data'!$A$1:$B$1"))
        rules.defined_names.add(DefinedName("校验区域", attr_text="'Data'!$D$2:$D$5"))
        args = ["summary-named-validation", "--config", self.config, "--template", template,
                "--log-view", "developer", "--log-dir", self.root / "calclogs", source]
        for mode in ("com", "betteroffice"):
            with self.subTest(mode=mode):
                output_dir = self.root / mode
                rules.save(template)
                self.call(*args, "--calc-engine", mode, "--output-dir", output_dir)
                result = next(output_dir.glob("*.xlsx"))
                report = openpyxl.load_workbook(result, data_only=True)
                result_rows = list(report.active.values)
                self.assertEqual(len(result_rows), 2, result_rows)
                self.assertIn("D2", str(result_rows[1]))
                sheet["D2"] = "=SYNTHETIC_UNKNOWN_NAME"
                sheet["D4"] = "=INDIRECT(\"NoSuchSheet!A1\")"
                sheet["D5"] = "=1/0"
                rules.save(template)
                sheet["D2"] = '=IF(SUM(B2:B3)=30,"通过","校验失败：金额不平")'
                sheet["D4"] = None
                sheet["D5"] = None
                with_errors = self.root / f"{mode}-with-errors"
                self.call(*args, "--calc-engine", mode, "--output-dir", with_errors)
                report = openpyxl.load_workbook(next(with_errors.glob("*.xlsx")), data_only=True)
                rows = {row[4]: row[5:7] for row in list(report.active.values)[1:]}
                report.close()
                self.assertEqual(rows, {
                    "D2": ("公式计算错误", "#NAME?"),
                    "D4": ("公式计算错误", "#REF!"),
                    "D5": ("公式计算错误", "#DIV/0!"),
                })
        rules.save(template)
        no_recalculation = self.root / "no_recalculation.py"
        no_recalculation.write_text(
            "import json, sys\n"
            "print(json.dumps({'ready': True, 'protocol': 3}), flush=True)\n"
            "for line in sys.stdin.buffer:\n"
            "    if json.loads(line)['op'] == 'shutdown': break\n"
            "    print(json.dumps({'ok': True, 'application': 'Synthetic COM', 'cells': []}), flush=True)\n",
            encoding="utf-8",
        )
        missing = self.root / "missing-cache"
        self.call(*args, "--calc-engine", "com", "--output-dir", missing,
                  env=dict(os.environ, PYTOOLS_COM_WORKER=str(no_recalculation)))
        report = openpyxl.load_workbook(next(missing.glob("*.xlsx")), data_only=True)
        rows = {row[4]: row[5] for row in list(report.active.values)[1:]}
        report.close()
        self.assertEqual(rows, {"D2": "结果未保存", "D3": "结果未保存"})
        self.assertEqual(checksum(source), source_hash)


if __name__ == "__main__":
    unittest.main(verbosity=2)
