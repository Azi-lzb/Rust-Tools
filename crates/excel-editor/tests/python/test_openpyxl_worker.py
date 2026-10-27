import hashlib
import json
import os
from pathlib import Path
import subprocess
import sys
import tempfile
import unittest
import zipfile
from xml.etree import ElementTree as ET

import openpyxl
from openpyxl.styles import Font

WORKER = Path(__file__).resolve().parents[2] / "python" / "openpyxl_worker.py"


def digest(path):
    return hashlib.sha256(path.read_bytes()).hexdigest()


def run_worker(plan, output):
    request = output.with_suffix(".request.json")
    response = output.with_suffix(".response.json")
    request.write_text(json.dumps({
        "schema_version": 1,
        "request_id": plan["request_id"],
        "format": "xlsx",
        "input": plan["input"],
        "output": str(output),
        "plan": plan,
    }), encoding="utf-8")
    process = subprocess.run(
        [sys.executable, "-I", str(WORKER), "--request", str(request), "--response", str(response)],
        capture_output=True,
        text=True,
        timeout=30,
    )
    return process, json.loads(response.read_text(encoding="utf-8"))


def operation(kind, **values):
    return {"operation": kind, **values}


class OpenpyxlWorkerTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory(prefix="openpyxl-worker-test-")
        self.root = Path(self.temp.name)
        self.source = self.root / "source.xlsx"
        book = openpyxl.Workbook()
        sheet = book.active
        sheet.title = "Data"
        sheet["A1"] = "before"
        sheet["B1"] = 3
        sheet["C2"] = '=A1&"!"'
        sheet.merge_cells("D1:E1")
        sheet["D1"] = "merged"
        sheet["A1"].comment = openpyxl.comments.Comment("old note", "tester")
        sheet["B2"].comment = openpyxl.comments.Comment("retain note", "tester")
        book.create_sheet("Notes")["A1"] = "keep"
        book.create_sheet("DeleteMe")
        book.save(self.source)
        self.import_source = self.root / "import.xlsx"
        imported = openpyxl.Workbook()
        src = imported.active
        src.title = "Source"
        src["A1"] = "imported"
        src["B1"] = 9
        src["A1"].font = Font(bold=True)
        src["A1"].comment = openpyxl.comments.Comment("copied note", "source-author")
        src["C1"] = "=1+1"
        src.merge_cells("A2:B2")
        src["A2"] = "merged import"
        src.row_dimensions[1].height = 25
        src.column_dimensions["A"].width = 22
        imported.save(self.import_source)
        self._add_formula_cache(self.import_source, "xl/worksheets/sheet1.xml", "C1", "2")

    @staticmethod
    def _add_formula_cache(path, part_name, address, value):
        replacement = path.with_name(path.stem + ".cached.xlsx")
        with zipfile.ZipFile(path, "r") as original, zipfile.ZipFile(replacement, "w", zipfile.ZIP_DEFLATED) as patched:
            for item in original.infolist():
                data = original.read(item.filename)
                if item.filename == part_name:
                    root = ET.fromstring(data)
                    main_ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
                    cell = root.find(f".//{{{main_ns}}}c[@r='{address}']")
                    formula = cell.find(f"{{{main_ns}}}f")
                    if formula is None:
                        raise AssertionError("expected formula cell")
                    cached_value = cell.find(f"{{{main_ns}}}v")
                    if cached_value is None:
                        cached_value = ET.SubElement(cell, f"{{{main_ns}}}v")
                    cached_value.text = value
                    ET.register_namespace("", main_ns)
                    data = ET.tostring(root, encoding="utf-8", xml_declaration=True)
                patched.writestr(item, data)
        replacement.replace(path)

    def tearDown(self):
        self.temp.cleanup()

    def base_plan(self, output, operations):
        return {
            "request_id": "worker-test-1",
            "input": str(self.source),
            "output": str(output),
            "format": "xlsx",
            "workbook": {"format": "xlsx", "inventory_complete": True, "features_present": []},
            "operations": operations,
            "preserve": [],
        }

    def test_batch_modifies_one_workbook_and_preserves_supported_structures(self):
        output = self.root / "candidate.xlsx"
        source_hash = digest(self.source)
        import_hash = digest(self.import_source)
        operations = [
            operation("set_cell_value", sheet="Data", cell="A1", value={"type": "decimal", "value": "12.75"}),
            operation("set_formula_text", sheet="Data", cell="B3", formula="=A1*2"),
            operation("set_comment", sheet="Data", cell="A1", text="new note", author="worker"),
            operation("set_column_width", sheet="Data", column_index=2, width=18.5),
            operation("set_row_height", sheet="Data", row_index=3, height=27.0),
            operation("set_cell_fill_color", sheet="Data", cell="A1", argb="FF112233"),
            operation("set_page_settings", sheet="Data", settings={
                "orientation": "landscape", "fit_to_page": True, "fit_to_width": 1,
                "fit_to_height": 0, "clear_scale": True, "paper_size": 9,
                "print_area": "A1:E3", "horizontal_centered": True,
                "vertical_centered": False, "footer": "&P",
            }),
            operation("copy_range", source_sheet="Data", source_range="A1:B1",
                      target_sheet="Data", target_start="H5", copy_values=True,
                      copy_formulas=False, copy_styles=True, copy_comments=True),
            operation("import_range", source_workbook=str(self.import_source), source_sheet="Source",
                      source_range="A1:C2", target_sheet="Data", target_start="J1",
                      copy_values=True, copy_formulas=True, style_mode="full", copy_comments=True,
                      copy_merges=True, copy_row_heights=True, copy_column_widths=True,
                      copy_row_hidden=False, copy_column_hidden=False, copy_outline_levels=False,
                      copy_all_row_dimensions=False, copy_all_column_dimensions=False,
                      materialize_blank_cells=True, hide_zero=False),
            operation("set_workbook_calculation_properties", full_calc_on_load=True, force_full_calc=True),
            operation("rename_sheet", old_name="Notes", new_name="Information"),
            operation("create_sheet", name="Generated"),
            operation("create_sheet_at", name="History", index=0),
            operation("set_freeze_panes", sheet="History", cell="A2"),
            operation("set_auto_filter", sheet="History", range="A1:M2"),
            operation("delete_sheet", name="DeleteMe"),
        ]
        plan = self.base_plan(output, operations)
        process, response = run_worker(plan, output)
        self.assertEqual(process.returncode, 0, process.stderr)
        self.assertEqual(response["status"], "applied", response)
        self.assertEqual(response["operation_count"], len(operations))
        self.assertEqual(response["input_sha256_before"], response["input_sha256_after"])
        self.assertEqual(digest(self.source), source_hash)
        self.assertEqual(digest(self.import_source), import_hash)

        saved = openpyxl.load_workbook(output, data_only=False)
        self.assertEqual(saved.sheetnames, ["History", "Data", "Information", "Generated"])
        self.assertEqual(saved["History"].freeze_panes, "A2")
        self.assertEqual(saved["History"].auto_filter.ref, "A1:M2")
        sheet = saved["Data"]
        self.assertEqual(sheet["A1"].value, 12.75)
        self.assertEqual(sheet["B3"].value, "=A1*2")
        self.assertEqual(sheet["C2"].value, '=A1&"!"')
        self.assertEqual(sheet["A1"].comment.text, "new note")
        self.assertEqual(sheet["B2"].comment.text, "retain note")
        self.assertEqual(sheet["A1"].fill.fgColor.rgb, "FF112233")
        self.assertAlmostEqual(sheet.column_dimensions["B"].width, 18.5)
        self.assertAlmostEqual(sheet.row_dimensions[3].height, 27.0)
        self.assertEqual(sheet.print_area, "'Data'!$A$1:$E$3")
        self.assertEqual(sheet.page_setup.orientation, "landscape")
        self.assertTrue(sheet.sheet_properties.pageSetUpPr.fitToPage)
        self.assertEqual(sheet.oddFooter.center.text, "&P")
        self.assertEqual(sheet["H5"].value, 12.75)
        self.assertEqual(sheet["I5"].value, 3)
        self.assertEqual(sheet["H5"].comment.text, "new note")
        self.assertEqual(sheet["J1"].value, "imported")
        self.assertTrue(sheet["J1"].font.bold)
        self.assertEqual(sheet["J1"].comment.text, "copied note")
        self.assertEqual(sheet["L1"].value, "=1+1")
        self.assertTrue(saved.calculation.fullCalcOnLoad)
        self.assertTrue(saved.calculation.forceFullCalc)
        self.assertIn("J2:K2", {item.coord for item in sheet.merged_cells.ranges})
        self.assertAlmostEqual(sheet.row_dimensions[1].height, 25)
        self.assertAlmostEqual(sheet.column_dimensions["J"].width, 22)
        saved.close()

    def test_delete_and_recreate_same_sheet_validates_final_names_and_order(self):
        output = self.root / "recreated-sheet.xlsx"
        plan = self.base_plan(output, [
            operation("delete_sheet", name="DeleteMe"),
            operation("create_sheet_at", name="DeleteMe", index=1),
        ])

        process, response = run_worker(plan, output)

        self.assertEqual(process.returncode, 0, process.stderr)
        self.assertEqual(response["status"], "applied", response)
        self.assertEqual(response["operation_count"], 2)
        saved = openpyxl.load_workbook(output, data_only=False)
        self.assertEqual(saved.sheetnames, ["Data", "DeleteMe", "Notes"])
        saved.close()

    def test_unsupported_row_shift_is_rejected_before_output_write(self):
        output = self.root / "unsupported.xlsx"
        plan = self.base_plan(output, [operation("insert_rows", sheet="Data", first_row=2, count=1)])
        process, response = run_worker(plan, output)
        self.assertEqual(process.returncode, 0)
        self.assertEqual(response["status"], "unsupported_before_write")
        self.assertEqual(response["error"]["kind"], "unsupported_operation")
        self.assertFalse(output.exists())

    def test_formula_cache_is_rejected_before_write(self):
        cached = self.root / "cached.xlsx"
        with zipfile.ZipFile(self.source, "r") as original, zipfile.ZipFile(cached, "w", zipfile.ZIP_DEFLATED) as patched:
            for item in original.infolist():
                data = original.read(item.filename)
                if item.filename == "xl/worksheets/sheet1.xml":
                    root = ET.fromstring(data)
                    main_ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
                    cell = root.find(f".//{{{main_ns}}}c[@r='C2']")
                    self.assertIsNotNone(cell)
                    cached_value = cell.find(f"{{{main_ns}}}v")
                    self.assertIsNotNone(cached_value)
                    cached_value.text = "1"
                    ET.register_namespace("", main_ns)
                    data = ET.tostring(root, encoding="utf-8", xml_declaration=True)
                patched.writestr(item, data)
        self.assertTrue(zipfile.is_zipfile(cached))
        self.source = cached
        output = self.root / "cached-output.xlsx"
        plan = self.base_plan(output, [operation("set_cell_value", sheet="Data", cell="A1", value={"type": "text", "value": "changed"})])
        process, response = run_worker(plan, output)
        self.assertEqual(process.returncode, 0)
        self.assertEqual(response["status"], "unsupported_before_write")
        self.assertEqual(response["error"]["kind"], "formula_cache_not_preserved")
        self.assertFalse(output.exists())

    def test_formula_cache_inside_materialized_import_range_can_be_replaced(self):
        cached = self.root / "cached-target.xlsx"
        openpyxl.Workbook().save(cached)
        workbook = openpyxl.load_workbook(cached)
        workbook.active.title = "Data"
        workbook["Data"]["C2"] = "=1+1"
        workbook.save(cached)
        self._add_formula_cache(cached, "xl/worksheets/sheet1.xml", "C2", "2")

        output = self.root / "cached-target-output.xlsx"
        plan = self.base_plan(output, [operation(
            "import_range",
            source_workbook=str(self.import_source), source_sheet="Source",
            source_range="A1:C2", target_sheet="Data", target_start="A1",
            copy_values=True, copy_formulas=True, style_mode="full", copy_comments=True,
            copy_merges=True, copy_row_heights=False, copy_column_widths=False,
            copy_row_hidden=False, copy_column_hidden=False, copy_outline_levels=False,
            copy_all_row_dimensions=False, copy_all_column_dimensions=False,
            materialize_blank_cells=True, hide_zero=False,
        )])
        plan["input"] = str(cached)
        process, response = run_worker(plan, output)
        self.assertEqual(process.returncode, 0, process.stderr)
        self.assertEqual(response["status"], "applied", response)
        saved_values = openpyxl.load_workbook(output, data_only=True)
        self.assertIsNone(saved_values["Data"]["C2"].value)
        saved_values.close()
        saved_formulas = openpyxl.load_workbook(output, data_only=False)
        self.assertIsNone(saved_formulas["Data"]["C2"].value)
        saved_formulas.close()

    def test_materialized_import_range_does_not_cover_formula_cache_outside_target(self):
        cached = self.root / "cached-target-outside.xlsx"
        openpyxl.Workbook().save(cached)
        workbook = openpyxl.load_workbook(cached)
        workbook.active.title = "Data"
        workbook["Data"]["C2"] = "=1+1"
        workbook.save(cached)
        self._add_formula_cache(cached, "xl/worksheets/sheet1.xml", "C2", "2")

        output = self.root / "cached-target-outside-output.xlsx"
        plan = self.base_plan(output, [operation(
            "import_range",
            source_workbook=str(self.import_source), source_sheet="Source",
            source_range="A1:B2", target_sheet="Data", target_start="A1",
            copy_values=True, copy_formulas=True, style_mode="full", copy_comments=True,
            copy_merges=True, copy_row_heights=False, copy_column_widths=False,
            copy_row_hidden=False, copy_column_hidden=False, copy_outline_levels=False,
            copy_all_row_dimensions=False, copy_all_column_dimensions=False,
            materialize_blank_cells=True, hide_zero=False,
        )])
        plan["input"] = str(cached)
        process, response = run_worker(plan, output)
        self.assertEqual(process.returncode, 0, process.stderr)
        self.assertEqual(response["status"], "unsupported_before_write")
        self.assertEqual(response["error"]["kind"], "formula_cache_not_preserved")
        self.assertFalse(output.exists())

    def test_operation_error_is_structured_and_does_not_publish_output(self):
        output = self.root / "failed.xlsx"
        plan = self.base_plan(output, [operation("set_cell_value", sheet="Missing", cell="A1", value={"type": "text", "value": "x"})])
        process, response = run_worker(plan, output)
        self.assertNotEqual(process.returncode, 0)
        self.assertEqual(response["status"], "failed")
        self.assertEqual(response["error"]["kind"], "operation_failed")
        self.assertFalse(output.exists())


class StyleTableIntegrityTest(unittest.TestCase):
    """Direct coverage for the saved-output style-table integrity assertion."""

    def setUp(self):
        import importlib.util
        spec = importlib.util.spec_from_file_location("openpyxl_worker_module", WORKER)
        module = importlib.util.module_from_spec(spec)
        spec.loader.exec_module(module)
        self.integrity = module._verify_style_table_integrity
        self.root = Path(tempfile.mkdtemp(prefix="worker-style-integrity-"))
        self.addCleanup(self._cleanup)

    def _cleanup(self):
        import shutil
        shutil.rmtree(self.root, ignore_errors=True)

    def _workbook_with_styles(self, path, cellxfs_body, declared_count, sheet_xml):
        styles = (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
            '<fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>'
            '<fills count="2"><fill><patternFill patternType="none"/></fill>'
            '<fill><patternFill patternType="gray125"/></fill></fills>'
            '<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>'
            '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>'
            f'<cellXfs count="{declared_count}">{cellxfs_body}</cellXfs>'
            '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>'
            '</styleSheet>'
        )
        with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as archive:
            archive.writestr("[Content_Types].xml",
                '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
                '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
                '<Default Extension="xml" ContentType="application/xml"/>'
                '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
                '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
                '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/></Types>')
            archive.writestr("_rels/.rels",
                '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>')
            archive.writestr("xl/workbook.xml",
                '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
                '<sheets><sheet name="Data" sheetId="1" r:id="rId1"/></sheets></workbook>')
            archive.writestr("xl/_rels/workbook.xml.rels",
                '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>'
                '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>')
            archive.writestr("xl/styles.xml", styles)
            archive.writestr("xl/worksheets/sheet1.xml", sheet_xml)

    def test_out_of_range_style_reference_is_rejected(self):
        target = self.root / "oob.xlsx"
        self._workbook_with_styles(
            target,
            cellxfs_body='<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>',
            declared_count=1,
            sheet_xml='<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                      '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
                      '<sheetData><row r="1"><c r="A1" t="inlineStr" s="5"><is><t>x</t></is></c></row></sheetData></worksheet>',
        )
        with self.assertRaises(Exception) as caught:
            self.integrity(target)
        self.assertIn("beyond its", str(caught.exception))

    def test_out_of_range_column_dimension_style_is_rejected(self):
        target = self.root / "column-style-oob.xlsx"
        self._workbook_with_styles(
            target,
            cellxfs_body='<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>',
            declared_count=1,
            sheet_xml=(
                '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                '<x:worksheet xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
                '<x:cols><x:col min="1" max="1" style="1"/></x:cols>'
                '<x:sheetData><x:row r="1"><x:c r="A1" t="inlineStr"><x:is><x:t>x</x:t>'
                '</x:is></x:c></x:row></x:sheetData></x:worksheet>'
            ),
        )
        with self.assertRaises(Exception) as caught:
            self.integrity(target)
        self.assertIn("cell style 1 beyond its 1 cellXfs entries", str(caught.exception))

    def test_count_mismatch_is_rejected(self):
        target = self.root / "mismatch.xlsx"
        self._workbook_with_styles(
            target,
            cellxfs_body='<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>',
            declared_count=9,
            sheet_xml='<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                      '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
                      '<sheetData><row r="1"><c r="A1" t="inlineStr" s="0"><is><t>x</t></is></c></row></sheetData></worksheet>',
        )
        with self.assertRaises(Exception) as caught:
            self.integrity(target)
        self.assertIn("declares count=", str(caught.exception))

    def test_consistent_table_passes(self):
        target = self.root / "consistent.xlsx"
        self._workbook_with_styles(
            target,
            cellxfs_body='<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>'
                         '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyAlignment="1"><alignment vertical="center"/></xf>',
            declared_count=2,
            sheet_xml='<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                      '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
                      '<sheetData><row r="1"><c r="A1" t="inlineStr" s="1"><is><t>x</t></is></c></row></sheetData></worksheet>',
        )
        self.integrity(target)


if __name__ == "__main__":
    unittest.main()
