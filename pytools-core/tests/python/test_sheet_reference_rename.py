"""Shared reference contract for both shipped Openpyxl rename workers."""
import runpy
import unittest
from pathlib import Path


ROOT = Path(__file__).resolve().parents[3]


class SheetReferenceRename(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        generic = runpy.run_path(str(ROOT / "crates/excel-editor/python/openpyxl_worker.py"))
        windows = runpy.run_path(str(ROOT / "pytools-core/python/pywin32_excel_conversion_worker.py"))
        cls.rewriters = [generic["_rewrite_formula_sheet_references"], windows["_rewrite_sheet_refs"]]

    def test_sheet_references_are_rewritten_once_without_changing_string_or_external_data(self):
        renames = {"A": "RenamedA", "B": "A", "O'Brien": "New O'Brien"}
        cases = [
            ("='B'!A1+B!A2", "=A!A1+A!A2"),
            ("='a'!A1+b!A2", "=RenamedA!A1+A!A2"),
            ('="B!A1"&\'B\'!A1', '="B!A1"&A!A1'),
            ("='[Book.xlsx]B'!A1+[1]B!A1", "='[Book.xlsx]B'!A1+[1]B!A1"),
            ("=NotB!A1+'Other B'!A1", "=NotB!A1+'Other B'!A1"),
            ("='O''Brien'!A1", "='New O''Brien'!A1"),
            ("='B'!$A$1", "=A!$A$1"),
        ]
        for rewrite in self.rewriters:
            for source, expected in cases:
                with self.subTest(worker=rewrite.__name__, formula=source):
                    self.assertEqual(rewrite(source, renames), expected)


if __name__ == "__main__":
    unittest.main()
