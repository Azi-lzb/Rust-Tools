"""Independent semantic and package check for saved phase-2 artifacts."""
from __future__ import annotations

import hashlib
import json
import warnings
import zipfile
from pathlib import Path

import xlrd
from openpyxl import load_workbook

warnings.filterwarnings("ignore", category=UserWarning)
ROOT = Path(__file__).resolve().parents[1]


def sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for block in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(block)
    return digest.hexdigest()


def xlsx_semantics(path: Path) -> dict:
    wb = load_workbook(path, data_only=False)
    ws = wb["Data"] if "Data" in wb.sheetnames else wb.worksheets[0]
    formula = wb["Calc"]["A1"].value if "Calc" in wb.sheetnames else None
    return {
        "sheet_names": wb.sheetnames,
        "A3": ws["A3"].value,
        "B3": ws["B3"].value,
        "calc_A1_formula": formula,
        "merge_ranges": [str(r) for r in ws.merged_cells.ranges],
        "charts": len(ws._charts),
        "data_validations": len(ws.data_validations.dataValidation),
        "A2_comment": bool(ws["A2"].comment),
        "A3_hyperlink": ws["A3"].hyperlink.target if ws["A3"].hyperlink else None,
        "defined_names": sorted(wb.defined_names),
        "row1_height": ws.row_dimensions[1].height,
        "column_A_width": ws.column_dimensions["A"].width,
        "row4_hidden": ws.row_dimensions[4].hidden,
        "column_I_hidden": ws.column_dimensions["I"].hidden,
        "print_area": ws.print_area,
        "print_title_rows": ws.print_title_rows,
        "print_title_columns": ws.print_title_cols,
        "orientation": ws.page_setup.orientation,
        "paper_size": ws.page_setup.paperSize,
        "header_center": ws.oddHeader.center.text,
        "footer_right": ws.oddFooter.right.text,
    }


def xls_semantics(path: Path) -> dict:
    book = xlrd.open_workbook(path, formatting_info=True)
    sheet = book.sheet_by_index(0)
    return {
        "sheet_names": book.sheet_names(),
        "dimensions": [sheet.nrows, sheet.ncols],
        "A2": sheet.cell_value(1, 0),
        "B3": sheet.cell_value(2, 1),
        "B3_type": sheet.cell_type(2, 1),
        "merged_ranges": [list(r) for r in sheet.merged_cells],
    }


def parts(path: Path) -> list[str]:
    with zipfile.ZipFile(path) as package:
        return sorted(package.namelist())


def main() -> None:
    result: dict = {"fixtures": {}, "engines": {}}
    fixture_xlsx = ROOT / "fixtures/capabilities.xlsx"
    fixture_xls = ROOT / "fixtures/capabilities.xls"
    for path in (fixture_xlsx, fixture_xls):
        result["fixtures"][path.name] = {"bytes": path.stat().st_size, "sha256": sha256(path)}

    for engine in ("aspose", "gembox"):
        folder = ROOT / "reports/samples" / engine
        section = result["engines"].setdefault(engine, {})
        for suffix in ("noop.xlsx", "edited.xlsx", "capabilities.xlsx", "noop.xls", "edited.xls", "converted.xlsx"):
            path = folder / f"{engine}-{suffix}"
            if path.exists():
                artifact = {"bytes": path.stat().st_size, "sha256": sha256(path)}
                if path.suffix == ".xlsx":
                    artifact["semantics"] = xlsx_semantics(path)
                else:
                    artifact["semantics"] = xls_semantics(path)
                section[path.name] = artifact
        source_parts = set(parts(fixture_xlsx))
        saved = folder / f"{engine}-noop.xlsx"
        section["xlsx_noop_package_parts"] = {
            "input_count": len(source_parts),
            "output_count": len(parts(saved)),
            "removed": sorted(source_parts - set(parts(saved))),
            "added": sorted(set(parts(saved)) - source_parts),
        }

    libxl_dir = ROOT / "reports/samples/libxl"
    libxl_section = result["engines"].setdefault("libxl", {})
    for path in sorted(libxl_dir.glob("libxl-capabilities.*")):
        libxl_section[path.name] = {
            "bytes": path.stat().st_size,
            "sha256": sha256(path),
            "semantics": xls_semantics(path) if path.suffix == ".xls" else xlsx_semantics(path),
        }

    for engine in ("aspose", "gembox"):
        formula_name = "formula-check.xlsx" if engine == "aspose" else "gembox-formula-check.xlsx"
        formula_file = ROOT / f"reports/samples/{engine}/{formula_name}"
        if formula_file.exists():
            wb_formula = load_workbook(formula_file, data_only=False)
            wb_cached = load_workbook(formula_file, data_only=True)
            result["engines"][engine]["formula_check"] = {
                "formulas": [wb_formula["Calc"].cell(row, 1).value for row in range(1, 11)],
                "cached_values": [wb_cached["Calc"].cell(row, 1).value for row in range(1, 11)],
                "sha256": sha256(formula_file),
                "bytes": formula_file.stat().st_size,
            }

    out = ROOT / "reports/evidence/phase2-validation.json"
    out.parent.mkdir(parents=True, exist_ok=True)
    out.write_text(
        json.dumps(result, ensure_ascii=False, indent=2, default=lambda value: value.isoformat() if hasattr(value, "isoformat") else str(value)),
        encoding="utf-8",
    )
    print(out)


if __name__ == "__main__":
    main()
