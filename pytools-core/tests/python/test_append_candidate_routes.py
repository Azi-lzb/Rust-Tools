"""Regression for 2-3's real menu candidate/validate/commit chain.

Only synthetic files in TemporaryDirectory are edited. An explicit executable
is required so this test cannot accidentally validate an older dist build.
"""
import hashlib
import json
import os
from pathlib import Path
import subprocess
import tempfile
import unittest

import openpyxl
from openpyxl.chart import BarChart, Reference
from openpyxl.workbook.defined_name import DefinedName


EXE = os.environ.get("RUSTTOOLS_TEST_EXE")


def digest(path):
    return hashlib.sha256(path.read_bytes()).hexdigest()


@unittest.skipUnless(EXE, "set RUSTTOOLS_TEST_EXE to the current test/debug CLI")
class AppendCandidateRoutes(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory(prefix="append-candidate-routes-")
        self.addCleanup(self.temp.cleanup)
        self.root = Path(self.temp.name)

    def fixture(self, route, named=False, chart=False):
        root = self.root / route
        root.mkdir()
        source, target, config = [root / name for name in ("source.xlsx", "target.xlsx", "config.xlsx")]
        headers = ["ID"] + ["Column" + str(i) for i in range(2, 28)]
        rows = [("R%03d" % i,) + (None,) * 24 + (0 if i % 2 == 0 else "000123", bool(i % 2)) for i in range(100)]
        book = openpyxl.Workbook()
        book.active.title = "Source"
        book.active.append(headers)
        for row in rows:
            book.active.append(row)
        book.save(source)
        book.close()
        book = openpyxl.Workbook()
        sheet = book.active
        sheet.title = "Combined"
        sheet.append(headers)
        if named:
            book.defined_names.add(DefinedName("Keep", attr_text="'Combined'!$A$1:$AA$1"))
            sheet.print_area = "A1:AA101"
            sheet.freeze_panes = "B2"
        if chart:
            graph = BarChart()
            graph.add_data(Reference(sheet, min_col=1, min_row=1, max_row=1))
            sheet.add_chart(graph, "AC1")
        book.save(target)
        book.close()
        book = openpyxl.Workbook()
        sheet = book.active
        sheet.title = "全局配置"
        sheet.append(["配置项", "值"])
        for key, value in (("输入目录", root), ("输出目录", root / "output"), ("日志目录", root / "logs")):
            sheet.append([key, str(value)])
        sheet = book.create_sheet("去重追加数据配置")
        sheet.append(["是否启用", "源数据工作簿", "源数据工作表", "标识列序号", "目标工作簿", "目标工作表", "执行模式"])
        sheet.append(["Y", str(source), "Source", "1", str(target), "Combined", "追加"])
        book.save(config)
        book.close()
        return root, source, target, config, [tuple(headers)] + rows

    def run_menu(self, root, config, route):
        env = os.environ.copy()
        env["PYTOOLS_EXCEL_EDIT_ROUTE"] = route
        result = subprocess.run([EXE, "menu", "--config", str(config), "--log-view", "developer"],
                                input="2\n3\n0\n", cwd=root, env=env, capture_output=True,
                                text=True, encoding="utf-8", errors="replace", timeout=120)
        self.assertEqual(result.returncode, 0, result.stderr)
        self.last_terminal = result.stdout + result.stderr
        reports = sorted((root / "logs").glob("pytools-dedup-config_append-*.json"), key=lambda p: p.stat().st_mtime_ns)
        self.assertTrue(reports, result.stdout + result.stderr)
        return json.loads(reports[-1].read_text(encoding="utf-8-sig"))

    def assert_rows(self, target, expected):
        # Both normal and streaming readers must see every committed row.
        for stream in (False, True):
            book = openpyxl.load_workbook(target, read_only=stream, data_only=False)
            try:
                self.assertEqual(list(book["Combined"].values), expected)
                self.assertEqual(book["Combined"].max_row, 101)
                self.assertEqual(book["Combined"].max_column, 27)
            finally:
                book.close()

    def test_empty_name_container_append_and_repeat_across_three_routes(self):
        for route in ("RustNativeOnly", "Auto", "UMYAOnly"):
            with self.subTest(route=route):
                root, source, target, config, expected = self.fixture(route)
                before = (digest(source), digest(config))
                first = self.run_menu(root, config, route)
                self.assertEqual(first["metrics"]["failed_targets"], 0, first)
                self.assertEqual(first["metrics"]["committed_targets"], 1, first)
                self.assertEqual(first["metrics"]["committed_appended_rows"], 100, first)
                self.assert_rows(target, expected)
                committed = digest(target)
                second = self.run_menu(root, config, route)
                self.assertEqual(second["metrics"]["committed_appended_rows"], 0, second)
                self.assertEqual(digest(target), committed, "repeat must be a byte-identical no-op")
                self.assert_rows(target, expected)
                self.assertEqual((digest(source), digest(config)), before)

    def test_umya_preserves_real_names_print_area_and_freeze_panes(self):
        root, source, target, config, expected = self.fixture("UMYAOnly", named=True)
        before = digest(source)
        result = self.run_menu(root, config, "UMYAOnly")
        self.assertEqual(result["metrics"]["committed_targets"], 1, self.last_terminal)
        self.assert_rows(target, expected)
        book = openpyxl.load_workbook(target)
        try:
            # Optional quotes around a simple Sheet name are a serialization
            # spelling; assert the independently resolved sheet/range instead.
            self.assertEqual(list(book.defined_names["Keep"].destinations), [("Combined", "$A$1:$AA$1")])
            self.assertEqual(set(book.defined_names), {"Keep"})
            self.assertEqual(str(book.active.print_area), "'Combined'!$A$1:$AA$101")
            self.assertEqual(book.active.freeze_panes, "B2")
        finally:
            book.close()
        self.assertEqual(digest(source), before)

    def test_umya_unsupported_chart_refusal_leaves_all_inputs_unchanged(self):
        root, source, target, config, _ = self.fixture("UMYAOnly", chart=True)
        before = {path: digest(path) for path in (source, target, config)}
        result = self.run_menu(root, config, "UMYAOnly")
        self.assertEqual(result["metrics"]["committed_targets"], 0, result)
        self.assertEqual(result["metrics"]["failed_targets"], 1, result)
        self.assertEqual({path: digest(path) for path in before}, before)
        self.assertFalse(list(root.rglob("*candidate*.xlsx")))


if __name__ == "__main__":
    unittest.main()
