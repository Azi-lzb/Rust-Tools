"""Literal-empty preservation: semantics, coordinate mapping and corruption guards."""
import hashlib
import importlib.util
from pathlib import Path
import tempfile
import unittest
import zipfile
from xml.etree import ElementTree as ET

import openpyxl

ROOT = Path(__file__).resolve().parents[3]
spec = importlib.util.spec_from_file_location("empty_worker", ROOT / "crates/excel-editor/python/openpyxl_worker.py")
worker = importlib.util.module_from_spec(spec)
spec.loader.exec_module(worker)
NS = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"


class LiteralEmptyPreservation(unittest.TestCase):
    def setUp(self):
        self.temporary = tempfile.TemporaryDirectory()
        self.addCleanup(self.temporary.cleanup)
        self.root = Path(self.temporary.name)
        self.source, self.output = self.root / "source.xlsx", self.root / "output.xlsx"
        book = openpyxl.Workbook()
        book.active.title = "Data"
        book.active["A1"] = "copied"
        sheet = book.create_sheet("Keep")
        sheet["A1"] = 0; sheet["B1"] = False; sheet["C1"] = "0"; sheet["D1"] = " "
        sheet["AC126"] = "placeholder"
        sheet["AC127"] = "after"
        book.save(self.source)
        self.patch_source()
        self.source_hash = hashlib.sha256(self.source.read_bytes()).hexdigest()

    def patch_source(self, shared=False):
        with zipfile.ZipFile(self.source) as z:
            parts = {i.filename: (i, z.read(i.filename)) for i in z.infolist()}
        name = "xl/worksheets/sheet2.xml"
        info, data = parts[name]
        root = ET.fromstring(data)
        c = next(c for c in root.iter(NS + "c") if c.get("r") == "AC126")
        c[:] = []
        if shared:
            c.set("t", "str")
            ET.SubElement(c, NS + "v").text = ""
        else:
            c.set("t", "inlineStr")
            ET.SubElement(ET.SubElement(c, NS + "is"), NS + "t").text = ""
        parts[name] = (info, ET.tostring(root, encoding="utf-8"))
        with zipfile.ZipFile(self.source, "w") as z:
            for info, data in parts.values(): z.writestr(info, data)

    def save_candidate(self, change=None):
        book = openpyxl.load_workbook(self.source)
        if change: change(book)
        book.save(self.output); book.close()

    def restore(self, operations):
        worker._restore_unmodified_empty_strings(self.source, self.output, operations)
        self.assertEqual(self.source_hash, hashlib.sha256(self.source.read_bytes()).hexdigest())

    def test_roundtrip_preserves_empty_text_and_other_types(self):
        self.save_candidate()
        self.assertNotIn(("Keep", "AC126"), worker._explicit_empty_strings(self.output))
        self.restore([])
        self.assertIn(("Keep", "AC126"), worker._explicit_empty_strings(self.output))
        book = openpyxl.load_workbook(self.output)
        self.assertEqual([book["Keep"][p].value for p in ("A1", "B1", "C1", "D1")], [0, False, "0", " "])
        book.close()

    def test_missing_cell_is_materialized_for_literal_str(self):
        self.patch_source(shared=True)
        self.source_hash = hashlib.sha256(self.source.read_bytes()).hexdigest()
        self.save_candidate()
        self.restore([])
        self.assertIn(("Keep", "AC126"), worker._explicit_empty_strings(self.output))

    def test_planned_replacement_is_never_undone(self):
        for value in (0, False, "replacement", None, "=1+1"):
            with self.subTest(value=value):
                self.save_candidate(lambda w: setattr(w["Keep"]["AC126"], "value", value))
                kind = "set_formula_text" if value == "=1+1" else "set_cell_value"
                self.restore([dict(operation=kind, sheet="Keep", cell="AC126")])
                self.assertNotIn(("Keep", "AC126"), worker._explicit_empty_strings(self.output))

    def test_unexpected_nonempty_candidate_is_logged_and_left_intact(self):
        for value in (0, False, "lost original", "=1+1"):
            with self.subTest(value=value):
                self.save_candidate(lambda w: setattr(w["Keep"]["AC126"], "value", value))
                before = self.output.read_bytes()
                with self.assertLogs(worker.__name__, level="WARNING") as warnings:
                    self.restore([])
                self.assertTrue(any("Keep!AC126" in message for message in warnings.output))
                self.assertEqual(before, self.output.read_bytes())

    def test_rows_and_sheet_rename_map_coordinates(self):
        def move(w):
            w["Keep"].insert_rows(10, 3); w["Keep"].delete_rows(20, 1); w["Keep"].title = "Renamed"
        self.save_candidate(move)
        self.restore([dict(operation="insert_rows", sheet="Keep", first_row=10, count=3),
                      dict(operation="delete_rows", sheet="Keep", first_row=20, count=1),
                      dict(operation="rename_sheet", old_name="Keep", new_name="Renamed")])
        self.assertIn(("Renamed", "AC128"), worker._explicit_empty_strings(self.output))

    def test_deleted_row_and_reused_sheet_name_do_not_resurrect_data(self):
        self.save_candidate(lambda w: w["Keep"].delete_rows(126, 1))
        self.restore([dict(operation="delete_rows", sheet="Keep", first_row=126, count=1)])
        book = openpyxl.load_workbook(self.output)
        self.assertEqual(book["Keep"]["AC126"].value, "after"); book.close()
        def replace(w):
            del w["Keep"]; w.create_sheet("Keep")["AC126"] = "new sheet"
        self.save_candidate(replace)
        self.restore([dict(operation="delete_sheet", name="Keep"), dict(operation="create_sheet", name="Keep")])
        self.assertNotIn(("Keep", "AC126"), worker._explicit_empty_strings(self.output))

    def test_range_value_write_and_style_only_copy_are_distinct(self):
        operation = dict(operation="copy_range", source_sheet="Data", source_range="A1", target_sheet="Keep", target_start="AC126", copy_values=True, copy_formulas=False)
        self.save_candidate(lambda w: setattr(w["Keep"]["AC126"], "value", "copied"))
        self.restore([operation])
        self.assertNotIn(("Keep", "AC126"), worker._explicit_empty_strings(self.output))
        self.save_candidate()
        self.restore([dict(operation, copy_values=False, copy_styles=True)])
        self.assertIn(("Keep", "AC126"), worker._explicit_empty_strings(self.output))

    def test_batch_restore_of_many_cells_and_prefixed_xml(self):
        self.save_candidate()
        with zipfile.ZipFile(self.output) as z:
            parts = {i.filename: (i, z.read(i.filename)) for i in z.infolist()}
        name = "xl/worksheets/sheet2.xml"
        info, data = parts[name]
        root = ET.fromstring(data)
        sheet_data = root.find(NS + "sheetData")
        row = ET.SubElement(sheet_data, NS + "row", r="200")
        addresses = []
        for column in range(1, 1001):
            address = openpyxl.utils.get_column_letter(column) + "200"
            ET.SubElement(row, NS + "c", r=address)
            addresses.append(address)
        parts[name] = (info, ET.tostring(root, encoding="utf-8"))  # ns0:c, not literal <c>
        with zipfile.ZipFile(self.output, "w") as z:
            for info, payload in parts.values(): z.writestr(info, payload)
        worker._restore_planned_empty_strings(self.output, [dict(operation="set_cell_value", sheet="Keep", cell=address, value=dict(type="text", value="")) for address in addresses])
        self.assertTrue({("Keep", address) for address in addresses}.issubset(worker._explicit_empty_strings(self.output)))


if __name__ == "__main__": unittest.main()
