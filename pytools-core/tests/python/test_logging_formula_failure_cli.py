"""Run with the bundled Python and RUSTTOOLS_LOGGING_CLI pointing at a built CLI."""
import hashlib
import json
import os
from pathlib import Path
import shutil
import subprocess
import sys
import tempfile
import unittest

import openpyxl
from openpyxl.workbook.defined_name import DefinedName


class FormulaFailureLoggingCliTests(unittest.TestCase):
    @unittest.skipUnless(os.name == "nt", "Windows COM Worker protocol scenario")
    def test_recalculation_failure_is_sanitized_and_never_counted_as_success(self):
        executable = os.environ.get("RUSTTOOLS_LOGGING_CLI")
        if not executable:
            self.skipTest("set RUSTTOOLS_LOGGING_CLI to an already built CLI")
        with tempfile.TemporaryDirectory(prefix="rusttools-log-formula-failure-") as directory:
            root = Path(directory)
            config = root / "config.xlsx"
            template = root / "template.xlsx"
            source = root / "formula-source.xlsx"
            worker = root / "failing-com-worker.py"
            marker = root / "worker-calls.jsonl"
            book = openpyxl.Workbook()
            sheet = book.active
            sheet.title = "全局配置"
            sheet.append(["键", "值", "备注"])
            sheet.append(["输出目录", str(root / "output"), ""])
            sheet.append(["日志目录", str(root / "logs"), ""])
            sheet.append(["校验计算引擎", "com", "existing platform route"])
            book.save(config)
            for path in (source, template):
                book = openpyxl.Workbook()
                sheet = book.active
                sheet.title = "Data"
                sheet.append(["指标", "金额"])
                sheet.append(["合成机构", 1])
                if path == template:
                    sheet["A3"] = '=IF(B2>0,"错|合成校验|差异","")'
                    for name, ref in [
                        ("表头区域", "'Data'!$A$1:$B$1"),
                        ("固定行汇总区域", "'Data'!$A$2:$B$2"),
                        ("校验区域", "'Data'!$A$3"),
                    ]:
                        book.defined_names.add(DefinedName(name, attr_text=ref))
                book.save(path)
            before = {path.name: hashlib.sha256(path.read_bytes()).hexdigest()
                      for path in (source, template)}
            worker.write_text(r"""import json, os, sys
from pathlib import Path
print(json.dumps({"ready": True, "protocol": 3}), flush=True)
for line in sys.stdin:
    request = json.loads(line)
    with Path(os.environ["RUSTTOOLS_FAULT_MARKER"]).open("a", encoding="utf-8") as log:
        log.write(json.dumps(request) + "\n")
    if request.get("op") == "shutdown":
        break
    print(json.dumps({"ok": False, "stage": "calculate",
                      "error": "HRESULT 0x800A03EC; Traceback synthetic recalculation failure"}),
          flush=True)
""", encoding="utf-8")
            environment = dict(os.environ, PYTOOLS_CONFIG_PATH=str(config),
                               PYTOOLS_COM_PYTHON=sys.executable, PYTOOLS_COM_WORKER=str(worker),
                               RUSTTOOLS_FAULT_MARKER=str(marker))
            output = subprocess.run(
                [executable, "menu", "--config", str(config), "--log-view", "user"],
                input=f"1\n5\n{template}\n1\n{source}\n0\n0\n",
                capture_output=True, text=True, encoding="utf-8", timeout=90,
                cwd=root, env=environment,
            )
            evidence_root = os.environ.get("RUSTTOOLS_LOGGING_EVIDENCE_DIR")
            if evidence_root:
                evidence = Path(evidence_root) / "formula-recalc-failure"
                evidence.mkdir(parents=True, exist_ok=True)
                (evidence / "stdout.txt").write_text(output.stdout, encoding="utf-8")
                (evidence / "stderr.txt").write_text(output.stderr, encoding="utf-8")
                (evidence / "exit.txt").write_text(str(output.returncode), encoding="utf-8")
                if (root / "logs").exists():
                    shutil.copytree(root / "logs", evidence / "logs", dirs_exist_ok=True)
                if marker.exists():
                    shutil.copyfile(marker, evidence / "worker-calls.jsonl")
            self.assertTrue(marker.exists(), output.stdout + output.stderr)
            self.assertIn("recalculate_named_validation", marker.read_text(encoding="utf-8"))
            user_text = output.stdout + output.stderr
            for internal in ("HRESULT", "Traceback", "pywin32"):
                self.assertNotIn(internal, user_text)
            self.assertIn("formula-source.xlsx", user_text)
            summaries = []
            for path in (root / "logs").glob("*.json"):
                value = json.loads(path.read_text(encoding="utf-8"))
                if "input_stats" in value:
                    summaries.append(value)
            self.assertTrue(summaries, user_text)
            summary = summaries[-1]
            self.assertEqual(summary["input_stats"]["succeeded"], 0)
            self.assertEqual(summary["input_stats"]["failed"], 1)
            self.assertEqual(summary["input_stats"]["pending"], 0)
            self.assertFalse(summary.get("output_stats", {}).get("saved", False))
            self.assertEqual(summary["output_stats"].get("workbooks_saved", 0), 0)
            self.assertEqual(summary["output_stats"].get("sheets_saved", 0), 0)
            for path in (source, template):
                self.assertEqual(hashlib.sha256(path.read_bytes()).hexdigest(), before[path.name])

    def test_missing_worksheet_is_a_business_skip(self):
        self._check_named_skip("missing-worksheet")

    def test_missing_required_header_is_a_business_skip(self):
        self._check_named_skip("missing-required-header")

    def test_all_missing_worksheets_is_business_noop(self):
        self._check_named_skip("all-missing-worksheet")

    def _check_named_skip(self, scenario):
        all_skipped = scenario == "all-missing-worksheet"
        executable = os.environ.get("RUSTTOOLS_LOGGING_CLI")
        if not executable:
            self.skipTest("set RUSTTOOLS_LOGGING_CLI to an already built CLI")
        with tempfile.TemporaryDirectory(prefix="rusttools-log-business-skip-") as directory:
            root = Path(directory)
            config = root / "config.xlsx"
            template = root / "template.xlsx"
            accepted = root / "accepted.xlsx"
            rejected = root / f"{scenario}.xlsx"
            book = openpyxl.Workbook()
            sheet = book.active
            sheet.title = "全局配置"
            sheet.append(["键", "值", "备注"])
            sheet.append(["输出目录", str(root / "output"), ""])
            sheet.append(["日志目录", str(root / "logs"), ""])
            book.save(config)
            for path in (template, accepted, rejected):
                book = openpyxl.Workbook()
                sheet = book.active
                sheet.title = "其他" if path == rejected and scenario in ("missing-worksheet", "all-missing-worksheet") else "模板"
                sheet.append(["指标", None if path == rejected and scenario == "missing-required-header" else "金额"])
                sheet.append(["accepted-row" if path == accepted else "excluded-row", 1])
                if path == template:
                    for name, ref in [("表头区域", "'模板'!$A$1:$B$1"),
                                      ("表结构区域", "'模板'!$A$1:$B$1"),
                                      ("固定行汇总区域", "'模板'!$A$2:$B$2")]:
                        book.defined_names.add(DefinedName(name, attr_text=ref))
                book.save(path)
            paths = (template, accepted, rejected)
            before = {path.name: hashlib.sha256(path.read_bytes()).hexdigest() for path in paths}
            selection = str(rejected) if all_skipped else f"{accepted};{rejected}"
            output = subprocess.run(
                [executable, "menu", "--config", str(config), "--log-view", "user"],
                input=f"1\n3\n{template}\n1\n{selection}\n0\n0\n",
                capture_output=True, text=True, encoding="utf-8", timeout=90,
                cwd=root, env=dict(os.environ, PYTOOLS_CONFIG_PATH=str(config)),
            )
            evidence_root = os.environ.get("RUSTTOOLS_LOGGING_EVIDENCE_DIR")
            evidence = Path(evidence_root) / scenario if evidence_root else None
            if evidence:
                evidence.mkdir(parents=True, exist_ok=True)
                (evidence / "stdout.txt").write_text(output.stdout, encoding="utf-8")
                (evidence / "stderr.txt").write_text(output.stderr, encoding="utf-8")
                (evidence / "exit.txt").write_text(str(output.returncode), encoding="utf-8")
                if (root / "logs").exists():
                    shutil.copytree(root / "logs", evidence / "logs", dirs_exist_ok=True)
            self.assertEqual(output.returncode, 0, output.stdout + output.stderr)
            summaries = [json.loads(path.read_text(encoding="utf-8"))
                         for path in (root / "logs").glob("*.json")]
            summary = next((value for value in summaries if "input_stats" in value), None)
            self.assertIsNotNone(summary, output.stdout + output.stderr)
            stats = summary["input_stats"]
            self.assertEqual((stats["succeeded"], stats["skipped"], stats["failed"]),
                             (0, 1, 0) if all_skipped else (1, 1, 0))
            self.assertEqual(summary["overall_status"], "NO_OP" if all_skipped else "SUCCESS")
            self.assertIn(rejected.name, output.stderr)
            self.assertNotIn(accepted.name, output.stderr)
            result_paths = list((root / "output").rglob("*.xlsx"))
            self.assertEqual(len(result_paths), 0 if all_skipped else 1)
            values = []
            if all_skipped:
                self.assertNotIn("执行失败", output.stdout + output.stderr)
                self.assertEqual(summary["status"], "skipped")
                diagnostic_path = root / "logs" / summary["diagnostic_log_path"]
                events = [json.loads(line) for line in diagnostic_path.read_text(encoding="utf-8").splitlines()]
                self.assertFalse(any(event["level"] == "ERROR" for event in events))
            else:
                result = openpyxl.load_workbook(result_paths[0], data_only=True)
                values = [value for sheet in result for row in sheet.iter_rows(values_only=True) for value in row]
                self.assertEqual(values.count("accepted-row"), 1)
                self.assertEqual(summary["record_stats"]["rows_written"], values.count("accepted-row"))
                self.assertNotIn("excluded-row", values)
            for path in paths:
                self.assertEqual(hashlib.sha256(path.read_bytes()).hexdigest(), before[path.name])
            if evidence:
                if result_paths:
                    shutil.copyfile(result_paths[0], evidence / "result.xlsx")
                (evidence / "verification.json").write_text(json.dumps({
                    "scenario": scenario, "run_id": summary["run_id"],
                    "input_stats": stats, "actual_data_rows": values.count("accepted-row"),
                    "record_stats": summary.get("record_stats"),
                    "source_hashes_unchanged": True,
                }, ensure_ascii=False, indent=2), encoding="utf-8")


if __name__ == "__main__":
    unittest.main()
