"""Independent structural checks for real phase-two candidates."""
import json
from pathlib import Path
import zipfile
import xml.etree.ElementTree as ET

import openpyxl


ROOT = Path(__file__).resolve().parents[3]
EVIDENCE = ROOT / "artifacts" / "excel-edit-phase2"
MAIN_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
NS = {"m": MAIN_NS}


def json_rows(name):
    data = json.loads((EVIDENCE / name).read_text(encoding="utf-8"))
    return data.get("runs", []) if isinstance(data, dict) else data


def dim_tuple(dim):
    if dim is None:
        return None
    return tuple(getattr(dim, key, None) for key in (
        "height", "width", "hidden", "bestFit", "outlineLevel", "collapsed",
        "style", "min", "max", "thickTop", "thickBottom", "zeroHeight",
    ))


def style_signature(cell):
    return (str(cell.font), str(cell.fill), str(cell.border), str(cell.alignment),
            cell.number_format, str(cell.protection))


def verify_delete_row_styles():
    records = json_rows("dedup-results.json")
    original_path = Path(records[0]["input"])
    original_book = openpyxl.load_workbook(original_path, data_only=False)
    source_sheet = max(original_book.worksheets,
                       key=lambda sheet: sheet.max_row * sheet.max_column)
    source_rows = list(source_sheet.values)
    seen = set()
    duplicate_rows = set()
    for row_index, row in enumerate(source_rows[1:], start=2):
        value = row[7] if len(row) > 7 else None
        if value is None or str(value).strip() == "":
            continue
        if str(value) in seen:
            duplicate_rows.add(row_index)
        seen.add(str(value))
    row_map = {}
    target_row = 1
    for source_row in range(1, source_sheet.max_row + 1):
        if source_row in duplicate_rows:
            continue
        row_map[source_row] = target_row
        target_row += 1
    results = []
    for record in records:
        if record["menu"] != "2-2":
            continue
        output_book = openpyxl.load_workbook(record["output_copy"], data_only=False)
        output_sheet = output_book[source_sheet.title]
        style_differences = []
        for source_row, destination_row in row_map.items():
            for column in range(1, source_sheet.max_column + 1):
                before = style_signature(source_sheet.cell(source_row, column))
                after = style_signature(output_sheet.cell(destination_row, column))
                if before != after:
                    style_differences.append((source_row, destination_row, column))
                    if len(style_differences) >= 20:
                        break
            if len(style_differences) >= 20:
                break
        row_dimension_differences = []
        for source_row, destination_row in row_map.items():
            before = dim_tuple(source_sheet.row_dimensions.get(source_row))
            after = dim_tuple(output_sheet.row_dimensions.get(destination_row))
            if before != after:
                row_dimension_differences.append((source_row, destination_row, before, after))
                if len(row_dimension_differences) >= 20:
                    break
        column_dimension_equal = {
            key: dim_tuple(value) for key, value in source_sheet.column_dimensions.items()
        } == {
            key: dim_tuple(value) for key, value in output_sheet.column_dimensions.items()
        }
        results.append({
            "menu": "2-2",
            "mode": record["mode"],
            "deleted_rows": len(duplicate_rows),
            "survivor_rows_mapped": len(row_map),
            "cell_style_differences": style_differences,
            "row_dimension_differences": row_dimension_differences,
            "column_dimensions_equal": column_dimension_equal,
            "passed": not style_differences and not row_dimension_differences
                     and column_dimension_equal,
        })
        output_book.close()
    original_book.close()
    return results


def verify_5_1_inserted_row_styles():
    results = []
    for record in json_rows("special-results.json"):
        if record.get("menu") != "5-1" or not record.get("changed"):
            continue
        plan = json.loads(Path(record["plan_files"][0]).read_text(encoding="utf-8"))
        delete = next(op for op in plan["operations"]
                      if op["operation"] == "delete_rows")
        insert = next(op for op in plan["operations"]
                      if op["operation"] == "insert_rows")
        sheet_name = delete["sheet"]
        start = insert["first_row"]
        count = insert["count"]
        before_book = openpyxl.load_workbook(record["input"], data_only=False)
        after_book = openpyxl.load_workbook(record["output_copy"], data_only=False)
        before = before_book[sheet_name]
        after = after_book[sheet_name]
        template_row = delete["first_row"]
        template_styles = [(style_signature(before.cell(template_row, column)),
                            before.cell(template_row, column).style_id != 0)
                           for column in range(1, before.max_column + 1)]
        template_styled = sum(has_style for _, has_style in template_styles)
        inserted_styled = 0
        matching_template = 0
        holes = []
        for row in range(start, start + count):
            for column, (template_style, template_has_style) in enumerate(template_styles, start=1):
                target_cell = after.cell(row, column)
                actual_style = style_signature(target_cell)
                target_has_style = target_cell.style_id != 0
                if target_has_style:
                    inserted_styled += 1
                if target_has_style and actual_style == template_style:
                    matching_template += 1
                elif template_has_style:
                    holes.append(target_cell.coordinate)
        results.append({
            "menu": "5-1",
            "mode": record["mode"],
            "sheet": sheet_name,
            "template_row": template_row,
            "inserted_rows": list(range(start, start + count)),
            "template_styled_cells_per_row": template_styled,
            "inserted_styled_cells_total": inserted_styled,
            "inserted_cells_matching_template_style": matching_template,
            "style_holes": holes,
            "style_gap_observed": bool(holes),
        })
        before_book.close()
        after_book.close()
    return results


def header_values(root):
    element = root.find("m:headerFooter", NS)
    if element is None:
        return None
    return [(child.tag.rsplit("}", 1)[-1], dict(child.attrib), child.text)
            for child in element]


def verify_5_4_header_only():
    records = json_rows("special-results-final.json")
    results = []
    for record in records:
        if record["menu"] != "5-4" or not record.get("changed"):
            continue
        changed_parts = set(record["changed_parts"])
        expected_parts = {
            "xl/worksheets/sheet7.xml",
            "xl/worksheets/sheet8.xml",
            "xl/worksheets/sheet9.xml",
        }
        sheet_results = []
        with zipfile.ZipFile(record["input"]) as before_zip, \
                zipfile.ZipFile(record["output_copy"]) as after_zip:
            for name in sorted(expected_parts):
                before_root = ET.fromstring(before_zip.read(name))
                after_root = ET.fromstring(after_zip.read(name))
                before_header = header_values(before_root)
                after_header = header_values(after_root)
                for root in (before_root, after_root):
                    element = root.find("m:headerFooter", NS)
                    if element is not None:
                        root.remove(element)
                non_header_equal = ET.tostring(before_root) == ET.tostring(after_root)
                sheet_results.append({
                    "part": name,
                    "only_header_changed": non_header_equal,
                    "before_header": before_header,
                    "after_header": after_header,
                    "header_structure_equal": (
                        before_header is not None and after_header is not None
                        and [(tag, attrs) for tag, attrs, _ in before_header]
                            == [(tag, attrs) for tag, attrs, _ in after_header]
                    ),
                })
        results.append({
            "menu": "5-4",
            "mode": record["mode"],
            "changed_parts": sorted(changed_parts),
            "changed_parts_exact_expected": changed_parts == expected_parts,
            "sheet_header_checks": sheet_results,
            "passed": changed_parts == expected_parts
                     and all(row["only_header_changed"]
                             and row["header_structure_equal"] for row in sheet_results),
        })
    return results


def main():
    results = {
        "delete_row_style_mapping": verify_delete_row_styles(),
        "village_split_inserted_row_styles": verify_5_1_inserted_row_styles(),
        "fx_header_xml_scope": verify_5_4_header_only(),
    }
    (EVIDENCE / "structural-verify-results.json").write_text(
        json.dumps(results, ensure_ascii=False, indent=2), encoding="utf-8"
    )
    for row in results["delete_row_style_mapping"]:
        print("2-2", row["mode"], "styles", len(row["cell_style_differences"]),
              "row dimensions", len(row["row_dimension_differences"]),
              "columns", row["column_dimensions_equal"], "passed", row["passed"])
    for row in results["village_split_inserted_row_styles"]:
        print("5-1", row["mode"], "template cells", row["template_styled_cells_per_row"],
              "inserted styled", row["inserted_styled_cells_total"],
              "style gap", row["style_gap_observed"])
    for row in results["fx_header_xml_scope"]:
        print("5-4", row["mode"], "parts exact", row["changed_parts_exact_expected"],
              "header only", all(x["only_header_changed"] for x in row["sheet_header_checks"]),
              "passed", row["passed"])
    if any(not row.get("passed", True) for group in results.values() for row in group):
        raise SystemExit(1)


if __name__ == "__main__":
    main()
