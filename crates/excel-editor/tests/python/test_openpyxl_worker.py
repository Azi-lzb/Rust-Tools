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

    @staticmethod
    def _add_cell_metadata(path, address):
        replacement = path.with_name(path.stem + ".metadata.xlsx")
        main_ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
        rel_ns = "http://schemas.openxmlformats.org/package/2006/relationships"
        content_ns = "http://schemas.openxmlformats.org/package/2006/content-types"
        rel_type = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sheetMetadata"
        content_type = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheetMetadata+xml"
        metadata_xml = (
            f'<metadata xmlns="{main_ns}" xmlns:xda="http://schemas.microsoft.com/office/spreadsheetml/2017/dynamicarray">'
            '<metadataTypes count="1"><metadataType name="XLDAPR" minSupportedVersion="120000" '
            'copy="1" pasteAll="1" pasteValues="1" merge="1" splitFirst="1" rowColShift="1" '
            'clearFormats="1" clearComments="1" assign="1" coerce="1" cellMeta="1"/></metadataTypes>'
            '<futureMetadata name="XLDAPR" count="1"><bk><extLst><ext uri="{bdbb8cdc-fa1e-496e-a857-3c3f30c029c3">'
            '<xda:dynamicArrayProperties fDynamic="1" fCollapsed="0"/></ext></extLst></bk></futureMetadata>'
            '<cellMetadata count="1"><bk><rc t="1" v="0"/></bk></cellMetadata></metadata>'
        ).encode("utf-8")
        with zipfile.ZipFile(path, "r") as original, zipfile.ZipFile(
            replacement, "w", zipfile.ZIP_DEFLATED
        ) as patched:
            for item in original.infolist():
                data = original.read(item.filename)
                if item.filename == "xl/worksheets/sheet1.xml":
                    root = ET.fromstring(data)
                    cell = root.find(f".//{{{main_ns}}}c[@r='{address}']")
                    if cell is None:
                        raise AssertionError("expected metadata cell")
                    cell.set("cm", "1")
                    ET.register_namespace("", main_ns)
                    data = ET.tostring(root, encoding="utf-8", xml_declaration=True)
                elif item.filename == "xl/_rels/workbook.xml.rels":
                    root = ET.fromstring(data)
                    ET.SubElement(root, f"{{{rel_ns}}}Relationship", {
                        "Id": "rIdMetadata", "Type": rel_type, "Target": "metadata.xml"
                    })
                    ET.register_namespace("", rel_ns)
                    data = ET.tostring(root, encoding="utf-8", xml_declaration=True)
                elif item.filename == "[Content_Types].xml":
                    root = ET.fromstring(data)
                    ET.SubElement(root, f"{{{content_ns}}}Override", {
                        "PartName": "/xl/metadata.xml", "ContentType": content_type
                    })
                    ET.register_namespace("", content_ns)
                    data = ET.tostring(root, encoding="utf-8", xml_declaration=True)
                elif item.filename == "xl/workbook.xml":
                    root = ET.fromstring(data)
                    extension_list = ET.SubElement(root, f"{{{main_ns}}}extLst")
                    extension = ET.SubElement(extension_list, f"{{{main_ns}}}ext", {"uri": "urn:worker-test:calc-features"})
                    ET.SubElement(extension, "{urn:worker-test:calc-features}feature", {"name": "retained"})
                    ET.register_namespace("", main_ns)
                    ET.register_namespace("xcalcf", "urn:worker-test:calc-features")
                    data = ET.tostring(root, encoding="utf-8", xml_declaration=True)
                patched.writestr(item, data)
            patched.writestr("xl/metadata.xml", metadata_xml)
        replacement.replace(path)

    @staticmethod
    def _add_unstyled_column_range(path):
        replacement = path.with_name(path.stem + ".column-range.xlsx")
        with zipfile.ZipFile(path, "r") as original, zipfile.ZipFile(
            replacement, "w", zipfile.ZIP_DEFLATED
        ) as patched:
            for item in original.infolist():
                data = original.read(item.filename)
                if item.filename == "xl/worksheets/sheet1.xml":
                    marker = b"<sheetData>"
                    fragment = b'<cols><col min="1" max="3" width="12.25" customWidth="1"/></cols>'
                    if data.count(marker) != 1:
                        raise AssertionError("expected one sheetData opening tag")
                    data = data.replace(marker, fragment + marker, 1)
                patched.writestr(item, data)
        replacement.replace(path)

    @staticmethod
    def _add_calculation_chain(path):
        replacement = path.with_name(path.stem + ".calc-chain.xlsx")
        relationship_ns = "http://schemas.openxmlformats.org/package/2006/relationships"
        content_types_ns = "http://schemas.openxmlformats.org/package/2006/content-types"
        calc_chain_type = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/calcChain"
        calc_chain_content_type = "application/vnd.openxmlformats-officedocument.spreadsheetml.calcChain+xml"
        calc_chain = (
            b'<calcChain xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
            b'<c r="C2" i="1"/></calcChain>'
        )
        with zipfile.ZipFile(path, "r") as original, zipfile.ZipFile(
            replacement, "w", zipfile.ZIP_DEFLATED
        ) as patched:
            for item in original.infolist():
                data = original.read(item.filename)
                if item.filename == "xl/_rels/workbook.xml.rels":
                    root = ET.fromstring(data)
                    ET.SubElement(root, f"{{{relationship_ns}}}Relationship", {
                        "Id": "rIdCalcChain", "Type": calc_chain_type, "Target": "calcChain.xml"
                    })
                    ET.register_namespace("", relationship_ns)
                    data = ET.tostring(root, encoding="utf-8", xml_declaration=True)
                elif item.filename == "[Content_Types].xml":
                    root = ET.fromstring(data)
                    ET.SubElement(root, f"{{{content_types_ns}}}Override", {
                        "PartName": "/xl/calcChain.xml", "ContentType": calc_chain_content_type
                    })
                    ET.register_namespace("", content_types_ns)
                    data = ET.tostring(root, encoding="utf-8", xml_declaration=True)
                patched.writestr(item, data)
            patched.writestr("xl/calcChain.xml", calc_chain)
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

    def test_insert_rows_is_applied_to_candidate_without_changing_source(self):
        output = self.root / "inserted-rows.xlsx"
        source_hash = digest(self.source)
        plan = self.base_plan(output, [operation("insert_rows", sheet="Data", first_row=2, count=1)])
        process, response = run_worker(plan, output)
        self.assertEqual(process.returncode, 0, process.stderr)
        self.assertEqual(response["status"], "applied", response)
        self.assertEqual(response["verified_operation_count"], 1)
        self.assertEqual(digest(self.source), source_hash)
        saved = openpyxl.load_workbook(output, data_only=False)
        self.assertEqual(saved["Data"]["A1"].value, "before")
        self.assertIsNone(saved["Data"]["A2"].value)
        self.assertEqual(saved["Data"]["C3"].value, '=A1&"!"')
        self.assertEqual(saved["Data"]["B3"].comment.text, "retain note")
        self.assertEqual(saved["Notes"]["A1"].value, "keep")
        saved.close()

    def test_formula_cache_dependent_on_write_is_left_uncached(self):
        """2026-11-10 用户口径：依赖被写入单元格的公式缓存不再拒绝整个计划。

        计划写入 Data!A1，而 C2 的公式 `=A1&"!"` 依赖 A1。旧行为是写前拒绝；
        新行为是正常写入，但 C2 的旧缓存值（过期）不恢复——输出里该单元格
        没有缓存值，Excel/WPS 打开时按 fullCalcOnLoad/forceFullCalc 重算。
        """
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
        self.assertEqual(process.returncode, 0, process.stderr)
        self.assertEqual(response["status"], "applied", response)
        self.assertTrue(output.exists())

        # 写入生效。
        saved_formulas = openpyxl.load_workbook(output, data_only=False)
        self.assertEqual(saved_formulas["Data"]["A1"].value, "changed")
        # 公式保留；依赖计划写入的旧缓存值不再恢复（data_only 读不到旧值）。
        self.assertEqual(saved_formulas["Data"]["C2"].value, '=A1&"!"')
        saved_formulas.close()
        saved_values = openpyxl.load_workbook(output, data_only=True)
        self.assertIsNone(saved_values["Data"]["C2"].value)
        saved_values.close()

        # 输出被标记为打开时全量重算。
        with zipfile.ZipFile(str(output), "r") as package:
            workbook_xml = package.read("xl/workbook.xml").decode("utf-8")
        self.assertIn('fullCalcOnLoad="1"', workbook_xml)
        self.assertIn('forceFullCalc="1"', workbook_xml)

    def test_pycode_openpyxl_parity_feature_keeps_formula_cache_and_requests_recalc(self):
        cached = self.root / "pycode-parity-cached.xlsx"
        workbook = openpyxl.load_workbook(self.source)
        workbook["Data"]["C2"] = "=1+1"
        workbook.save(cached)
        workbook.close()
        self._add_formula_cache(cached, "xl/worksheets/sheet1.xml", "C2", "2")

        output = self.root / "pycode-parity-output.xlsx"
        plan = self.base_plan(
            output,
            [operation("set_cell_value", sheet="Data", cell="A1", value={"type": "text", "value": "changed"})],
        )
        plan["request_id"] = "pytools-adjust-rural-loan"
        plan["input"] = str(cached)
        process, response = run_worker(plan, output)

        self.assertEqual(process.returncode, 0, process.stderr)
        self.assertEqual(response["status"], "applied", response)
        values = openpyxl.load_workbook(output, data_only=True)
        self.assertEqual(values["Data"]["C2"].value, 2)
        values.close()
        formulas = openpyxl.load_workbook(output, data_only=False)
        self.assertEqual(formulas["Data"]["C2"].value, "=1+1")
        self.assertTrue(formulas.calculation.fullCalcOnLoad)
        self.assertTrue(formulas.calculation.forceFullCalc)
        formulas.close()

    def test_delete_rows_preserves_unrelated_formula_cache(self):
        cached = self.root / "delete-unrelated-cache.xlsx"
        self.source.replace(cached)
        self._add_formula_cache(cached, "xl/worksheets/sheet1.xml", "C2", "1")

        output = self.root / "delete-unrelated-cache-output.xlsx"
        plan = self.base_plan(output, [operation(
            "delete_rows", sheet="DeleteMe", first_row=1, count=1
        )])
        plan["input"] = str(cached)
        process, response = run_worker(plan, output)

        self.assertEqual(process.returncode, 0, process.stderr)
        self.assertEqual(response["status"], "applied", response)
        saved = openpyxl.load_workbook(output, data_only=True)
        self.assertEqual(saved["Data"]["C2"].value, 1)
        self.assertIn("DeleteMe", saved.sheetnames)
        self.assertIsNone(saved["DeleteMe"]["A1"].value)
        saved.close()

    def test_delete_rows_is_rejected_when_formula_references_modified_sheet(self):
        source = self.root / "delete-referenced-sheet.xlsx"
        book = openpyxl.load_workbook(self.source)
        book["Notes"]["A2"] = "=DeleteMe!A1"
        book.save(source)
        self._add_formula_cache(source, "xl/worksheets/sheet2.xml", "A2", "1")

        output = self.root / "delete-referenced-sheet-output.xlsx"
        plan = self.base_plan(output, [operation(
            "delete_rows", sheet="DeleteMe", first_row=1, count=1
        )])
        plan["input"] = str(source)
        process, response = run_worker(plan, output)

        self.assertEqual(process.returncode, 0, process.stderr)
        self.assertEqual(response["status"], "unsupported_before_write", response)
        self.assertIn("depends on a worksheet being edited", response["error"]["message"])
        self.assertFalse(output.exists())

    def test_delete_rows_is_rejected_when_modified_sheet_contains_formula(self):
        source = self.root / "delete-formula-sheet.xlsx"
        book = openpyxl.load_workbook(self.source)
        book["DeleteMe"]["A2"] = "=Data!A1"
        book.save(source)

        output = self.root / "delete-formula-sheet-output.xlsx"
        plan = self.base_plan(output, [operation(
            "delete_rows", sheet="DeleteMe", first_row=1, count=1
        )])
        plan["input"] = str(source)
        process, response = run_worker(plan, output)

        self.assertEqual(process.returncode, 0, process.stderr)
        self.assertEqual(response["status"], "unsupported_before_write", response)
        self.assertIn("worksheet that contains formulas", response["error"]["message"])
        self.assertFalse(output.exists())

    def test_empty_text_write_keeps_an_explicit_inline_string(self):
        output = self.root / "empty-text-output.xlsx"
        plan = self.base_plan(output, [operation(
            "set_cell_value", sheet="Data", cell="B3", value={"type": "text", "value": ""}
        )])
        process, response = run_worker(plan, output)
        self.assertEqual(process.returncode, 0, process.stderr)
        self.assertEqual(response["status"], "applied", response)

        with zipfile.ZipFile(output, "r") as package:
            root = ET.fromstring(package.read("xl/worksheets/sheet1.xml"))
            cell = root.find(
                ".//{http://schemas.openxmlformats.org/spreadsheetml/2006/main}c[@r='B3']"
            )
            self.assertIsNotNone(cell)
            self.assertEqual(cell.attrib.get("t"), "inlineStr")
            inline = cell.find(
                "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}is"
            )
            self.assertIsNotNone(inline)
            self.assertIsNotNone(inline.find(
                "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}t"
            ))

        saved = openpyxl.load_workbook(output, data_only=False)
        self.assertEqual(saved["Data"]["B3"].value, "")
        saved.close()

    def test_formula_cache_is_preserved_when_write_is_not_a_dependency(self):
        cached = self.root / "cached-unrelated-write.xlsx"
        self.source.replace(cached)
        self._add_formula_cache(cached, "xl/worksheets/sheet1.xml", "C2", "1")

        output = self.root / "cached-unrelated-write-output.xlsx"
        plan = self.base_plan(output, [operation(
            "set_cell_value", sheet="Data", cell="B1", value={"type": "decimal", "value": "7"}
        )])
        plan["input"] = str(cached)
        process, response = run_worker(plan, output)
        self.assertEqual(process.returncode, 0, process.stderr)
        self.assertEqual(response["status"], "applied", response)
        saved = openpyxl.load_workbook(output, data_only=True)
        self.assertEqual(saved["Data"]["C2"].value, 1)
        saved.close()

    def test_formula_cache_and_dynamic_cell_metadata_survive_unrelated_append(self):
        cached = self.root / "cached-cell-metadata.xlsx"
        self.source.replace(cached)
        self._add_formula_cache(cached, "xl/worksheets/sheet1.xml", "C2", "1")
        self._add_cell_metadata(cached, "C2")

        output = self.root / "cached-cell-metadata-output.xlsx"
        plan = self.base_plan(output, [operation(
            "set_cell_value", sheet="Data", cell="B1", value={"type": "decimal", "value": "8"}
        )])
        plan["input"] = str(cached)
        plan["workbook"]["features_present"] = ["cell_metadata"]
        process, response = run_worker(plan, output)
        self.assertEqual(process.returncode, 0, process.stderr)
        self.assertEqual(response["status"], "applied", response)

        with zipfile.ZipFile(cached, "r") as source, zipfile.ZipFile(output, "r") as result:
            self.assertEqual(result.read("xl/metadata.xml"), source.read("xl/metadata.xml"))
            relationships = ET.fromstring(result.read("xl/_rels/workbook.xml.rels"))
            self.assertTrue(any(
                element.attrib.get("Type", "").endswith("/sheetMetadata")
                for element in relationships
            ))
            content_types = ET.fromstring(result.read("[Content_Types].xml"))
            self.assertTrue(any(
                element.attrib.get("PartName") == "/xl/metadata.xml"
                for element in content_types
            ))
            workbook_xml = result.read("xl/workbook.xml")
            self.assertIn(b"urn:worker-test:calc-features", workbook_xml)
            worksheet = ET.fromstring(result.read("xl/worksheets/sheet1.xml"))
            cell = worksheet.find(".//{http://schemas.openxmlformats.org/spreadsheetml/2006/main}c[@r='C2']")
            self.assertEqual(cell.attrib.get("cm"), "1")

        saved = openpyxl.load_workbook(output, data_only=True)
        self.assertEqual(saved["Data"]["C2"].value, 1)
        saved.close()

    def test_unmodified_column_dimensions_keep_the_original_ooxml_fragment(self):
        source = self.root / "unstyled-columns.xlsx"
        self.source.replace(source)
        self._add_unstyled_column_range(source)
        output = self.root / "unstyled-columns-output.xlsx"
        plan = self.base_plan(output, [operation(
            "set_cell_value", sheet="Data", cell="A1", value={"type": "text", "value": "changed"}
        )])
        plan["input"] = str(source)

        def columns(path):
            with zipfile.ZipFile(path, "r") as package:
                xml = package.read("xl/worksheets/sheet1.xml")
            start = xml.index(b"<cols>")
            end = xml.index(b"</cols>", start) + len(b"</cols>")
            return xml[start:end]

        source_columns = columns(source)
        process, response = run_worker(plan, output)
        self.assertEqual(process.returncode, 0, process.stderr)
        self.assertEqual(response["status"], "applied", response)
        self.assertEqual(columns(output), source_columns)
        self.assertNotIn(b"style=", columns(output))

    def test_header_footer_ooxml_is_preserved_with_only_planned_text_replaced(self):
        sheet_part = "xl/worksheets/sheet1.xml"
        original_fragment = (
            b'<headerFooter alignWithMargins="0" scaleWithDoc="0">'
            b'<oddHeader>&amp;L\n Unit=OldUnit\n</oddHeader>'
            b'<oddFooter xml:space="preserve">&amp;LFooter OldUnit\n\n</oddFooter>'
            b"</headerFooter>"
        )
        with zipfile.ZipFile(self.source, "r") as original:
            source_parts = {
                item.filename: original.read(item.filename)
                for item in original.infolist()
            }
        worksheet_xml = source_parts[sheet_part]
        self.assertIn(b"</worksheet>", worksheet_xml)
        source_parts[sheet_part] = worksheet_xml.replace(
            b"</worksheet>", original_fragment + b"</worksheet>"
        )
        with zipfile.ZipFile(self.source, "w", zipfile.ZIP_DEFLATED) as patched:
            for name, data in source_parts.items():
                patched.writestr(name, data)

        output = self.root / "header-footer.xlsx"
        plan = self.base_plan(
            output,
            [
                operation(
                    "replace_header_text",
                    sheet="Data",
                    search="OldUnit",
                    replacement="NewUnit",
                )
            ],
        )
        process, response = run_worker(plan, output)

        self.assertEqual(process.returncode, 0, process.stderr)
        self.assertEqual(response["status"], "applied", response)
        expected_fragment = original_fragment.replace(b"OldUnit", b"NewUnit", 1)
        with zipfile.ZipFile(output, "r") as saved:
            saved_xml = saved.read(sheet_part)
            self.assertIn(expected_fragment, saved_xml)
            notes_xml = saved.read("xl/worksheets/sheet2.xml")
            self.assertNotIn(b"<headerFooter", notes_xml)

    def test_header_text_replacement_matches_numeric_xml_character_references(self):
        sheet_part = "xl/worksheets/sheet1.xml"
        original_fragment = (
            b'<headerFooter alignWithMargins="0">'
            b'<oddHeader>&amp;L&#21333;&#20301;&#65306;&#19975;&#20803;</oddHeader>'
            b'<oddFooter>&amp;L&#19975;&#20803;</oddFooter>'
            b"</headerFooter>"
        )
        with zipfile.ZipFile(self.source, "r") as original:
            source_parts = {
                item.filename: original.read(item.filename)
                for item in original.infolist()
            }
        worksheet_xml = source_parts[sheet_part]
        source_parts[sheet_part] = worksheet_xml.replace(
            b"</worksheet>", original_fragment + b"</worksheet>"
        )
        with zipfile.ZipFile(self.source, "w", zipfile.ZIP_DEFLATED) as patched:
            for name, data in source_parts.items():
                patched.writestr(name, data)

        output = self.root / "numeric-header-footer.xlsx"
        plan = self.base_plan(
            output,
            [
                operation(
                    "replace_header_text",
                    sheet="Data",
                    search="万元",
                    replacement="万美元",
                )
            ],
        )
        process, response = run_worker(plan, output)

        self.assertEqual(process.returncode, 0, process.stderr)
        self.assertEqual(response["status"], "applied", response)
        with zipfile.ZipFile(output, "r") as saved:
            root = ET.fromstring(saved.read(sheet_part))
            ns = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
            header_footer = root.find(ns + "headerFooter")
            self.assertEqual(header_footer.attrib, {"alignWithMargins": "0"})
            self.assertEqual(header_footer.find(ns + "oddHeader").text, "&L单位：万美元")
            self.assertEqual(header_footer.find(ns + "oddFooter").text, "&L万元")

    def test_calculation_chain_is_invalidated_when_cell_value_changes(self):
        source = self.root / "calculation-chain.xlsx"
        self.source.replace(source)
        self._add_formula_cache(source, "xl/worksheets/sheet1.xml", "C2", "1")
        self._add_calculation_chain(source)
        output = self.root / "calculation-chain-output.xlsx"
        plan = self.base_plan(output, [operation(
            "set_cell_value", sheet="Data", cell="B1", value={"type": "decimal", "value": "7"}
        )])
        plan["input"] = str(source)

        process, response = run_worker(plan, output)
        self.assertEqual(process.returncode, 0, process.stderr)
        self.assertEqual(response["status"], "applied", response)
        with zipfile.ZipFile(output, "r") as result:
            self.assertNotIn("xl/calcChain.xml", result.namelist())
            relationships = ET.fromstring(result.read("xl/_rels/workbook.xml.rels"))
            self.assertFalse(any(
                element.attrib.get("Type", "").endswith("/calcChain")
                and element.attrib.get("Target") == "calcChain.xml"
                for element in relationships
            ))
            content_types = ET.fromstring(result.read("[Content_Types].xml"))
            self.assertFalse(any(
                element.attrib.get("PartName") == "/xl/calcChain.xml"
                and element.attrib.get("ContentType", "").endswith(".calcChain+xml")
                for element in content_types
            ))
            workbook = ET.fromstring(result.read("xl/workbook.xml"))
            calc_properties = workbook.find(
                "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}calcPr"
            )
            self.assertIsNotNone(calc_properties)
            self.assertEqual(calc_properties.attrib.get("fullCalcOnLoad"), "1")
            self.assertEqual(calc_properties.attrib.get("forceFullCalc"), "1")
        cached_values = openpyxl.load_workbook(output, data_only=True)
        self.assertEqual(cached_values["Data"]["C2"].value, 1)
        cached_values.close()

    def test_formula_cache_is_preserved_for_style_only_edits(self):
        cached = self.root / "cached-style.xlsx"
        self.source.replace(cached)
        self._add_formula_cache(cached, "xl/worksheets/sheet1.xml", "C2", "1")

        output = self.root / "cached-style-output.xlsx"
        plan = self.base_plan(output, [operation(
            "set_cell_fill_color", sheet="Data", cell="A1", argb="FFFF0000"
        )])
        plan["input"] = str(cached)
        process, response = run_worker(plan, output)
        self.assertEqual(process.returncode, 0, process.stderr)
        self.assertEqual(response["status"], "applied", response)

        saved_formulas = openpyxl.load_workbook(output, data_only=False)
        self.assertEqual(saved_formulas["Data"]["C2"].value, '=A1&"!"')
        saved_formulas.close()
        saved_values = openpyxl.load_workbook(output, data_only=True)
        self.assertEqual(saved_values["Data"]["C2"].value, 1)
        saved_values.close()

    def test_formula_caches_are_preserved_when_only_appending_a_new_sheet(self):
        cached = self.root / "cached-append.xlsx"
        book = openpyxl.Workbook()
        book.active.title = "Data"
        book["Data"]["A1"] = "=1+1"
        book.save(cached)
        self._add_formula_cache(cached, "xl/worksheets/sheet1.xml", "A1", "2")

        output = self.root / "cached-append-output.xlsx"
        plan = self.base_plan(output, [
            operation("create_sheet", name="Added"),
            operation(
                "import_range",
                source_workbook=str(self.import_source), source_sheet="Source",
                source_range="A1:B1", target_sheet="Added", target_start="A1",
                copy_values=True, copy_formulas=False, style_mode="full", copy_comments=False,
                copy_merges=False, copy_row_heights=False, copy_column_widths=False,
                copy_row_hidden=False, copy_column_hidden=False, copy_outline_levels=False,
                copy_all_row_dimensions=False, copy_all_column_dimensions=False,
                materialize_blank_cells=True, hide_zero=False,
            ),
        ])
        plan["input"] = str(cached)
        process, response = run_worker(plan, output)
        self.assertEqual(process.returncode, 0, process.stderr)
        self.assertEqual(response["status"], "applied", response)

        saved_formulas = openpyxl.load_workbook(output, data_only=False)
        self.assertEqual(saved_formulas["Data"]["A1"].value, "=1+1")
        saved_formulas.close()
        saved_values = openpyxl.load_workbook(output, data_only=True)
        self.assertEqual(saved_values["Data"]["A1"].value, 2)
        saved_values.close()

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

    def test_error_diagnostics_preserve_outer_and_deep_stack_frames(self):
        import importlib.util
        spec = importlib.util.spec_from_file_location("worker_full_trace_test", WORKER)
        module = importlib.util.module_from_spec(spec)
        spec.loader.exec_module(module)
        errors = [
            (RuntimeError("synthetic internal failure"), 1, "failed"),
            (module.WorkerError("operation_failed", "synthetic operation failure"), 1, "failed"),
            (module.WorkerError("unsupported_input", "synthetic rejection", unsupported=True), 0, "unsupported_before_write"),
        ]
        for index, (error, exit_code, status) in enumerate(errors):
            with self.subTest(status=status, kind=type(error).__name__):
                namespace = {"diagnostic_error": error}
                functions = ["def diagnostic_level_0(request): return diagnostic_level_1()"]
                functions += [
                    f"def diagnostic_level_{depth}(): return diagnostic_level_{depth + 1}()"
                    for depth in range(1, 20)
                ]
                functions.append("def diagnostic_level_20(): raise diagnostic_error")
                exec("\n".join(functions), namespace)
                module._apply = namespace["diagnostic_level_0"]
                request = self.root / f"trace-{index}.request.json"
                response = self.root / f"trace-{index}.response.json"
                request.write_text('{"request_id": "trace-check"}', encoding="utf-8")
                self.assertEqual(module._run(request, response), exit_code)
                result = json.loads(response.read_text(encoding="utf-8"))
                self.assertEqual(result["status"], status)
                self.assertEqual(result["request_id"], "trace-check")
                evidence_dir = os.environ.get("RUSTTOOLS_LOGGING_EVIDENCE_DIR")
                if evidence_dir:
                    evidence = Path(evidence_dir) / "worker-trace"
                    evidence.mkdir(parents=True, exist_ok=True)
                    (evidence / f"response-{index}.json").write_text(
                        json.dumps(result, ensure_ascii=False, indent=2), encoding="utf-8"
                    )
                trace = result["diagnostics"]["traceback"]
                self.assertIn("diagnostic_level_0", trace)
                self.assertIn("diagnostic_level_20", trace)



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
