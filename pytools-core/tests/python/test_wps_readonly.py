"""Supplemental COM reads only; never quit/reconfigure an existing WPS app."""
import importlib.util
import json
import os
from pathlib import Path
import shutil
import sys
import tempfile
import unittest


spec = importlib.util.spec_from_file_location(
    "office_validation", Path(__file__).with_name("test_platform_office.py")
)
office = importlib.util.module_from_spec(spec)
spec.loader.exec_module(office)


@unittest.skipUnless(sys.platform == "win32" and os.environ.get("RUSTTOOLS_WPS_READONLY") == "1",
                     "requires explicit read-only WPS COM authorization")
class WpsReadOnly(unittest.TestCase):
    test_print_modes_reopen_with_office = office.OfficeValidation.test_print_modes_reopen_with_office
    test_xlsx_append_and_repeat_reopen_with_office = office.OfficeValidation.test_xlsx_append_and_repeat_reopen_with_office
    @classmethod
    def setUpClass(cls):
        import win32com.client
        # Do not create/restart another server or change application flags.
        cls.app = win32com.client.GetActiveObject("ket.Application")
        cls.initial_count = cls.app.Workbooks.Count
        print("Read-only WPS COM version", cls.app.Version, flush=True)

    @classmethod
    def tearDownClass(cls):
        # Release our proxy only. In particular, NEVER call Application.Quit.
        cls.app = None

    def setUp(self):
        self.assertEqual(self.app.Workbooks.Count, self.initial_count)

    def tearDown(self):
        self.assertEqual(self.app.Workbooks.Count, self.initial_count)

    def test_readonly_biff_does_not_change_bytes(self):
        # Read the synthetic BIFF created by the separate Excel validation.
        with tempfile.TemporaryDirectory() as folder:
            target = Path(folder) / "target.xls"
            shutil.copyfile(Path(os.environ["RUSTTOOLS_OFFICE_FIXTURE_DIR"]) / "target.xls", target)
            before = office.digest(target)
            original = target.read_bytes()
            book = self.app.Workbooks.Open(str(target), UpdateLinks=0, ReadOnly=True)
            after_open = office.digest(target)
            try:
                self.assertTrue(book.ReadOnly)
                sheet = book.Worksheets("Combined")
                self.assertEqual(sheet.Cells(1, 1).Value, "ID")
                self.assertEqual(sheet.Cells(2, 1).Value, "001")
                self.assertEqual(sheet.Cells(2, 2).Value, 99)
                self.assertEqual(sheet.Cells(12, 27).Formula.upper(), "=B2+1")
            finally:
                book.Close(False)
            after_close = office.digest(target)
            if after_close != before:
                evidence = (office.ROOT / "test-runs/platform-validation-20261113/wps-biff-diagnostic"
                            / (before[:12] + "-" + after_close[:12]))
                evidence.mkdir(parents=True, exist_ok=True)
                (evidence / "before.xls").write_bytes(original)
                (evidence / "after.xls").write_bytes(target.read_bytes())
                (evidence / "hashes.json").write_text(json.dumps({
                    "before": before, "after_open": after_open, "after_close": after_close,
                    "read_only": True, "save_changes": False,
                }, indent=2), encoding="utf-8")
            self.assertEqual(after_close, before)


if __name__ == "__main__":
    unittest.main(verbosity=2)
