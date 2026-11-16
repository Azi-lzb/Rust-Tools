"""Regression exposed by the Aspose 2-2 differential fixture, without Office."""
import argparse
import hashlib
import importlib.util
import json
from pathlib import Path
import openpyxl


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--worker", required=True)
    parser.add_argument("--output", required=True)
    args = parser.parse_args()
    root = Path(args.output).resolve(); root.mkdir(parents=True, exist_ok=False)
    spec = importlib.util.spec_from_file_location("production_openpyxl_worker", args.worker)
    worker = importlib.util.module_from_spec(spec); spec.loader.exec_module(worker)
    source = root / "source.xlsx"; output = root / "candidate.xlsx"
    wb = openpyxl.Workbook(); ws = wb.active; ws.title = "数据"
    for r in range(1, 13): ws.cell(r, 1, r)
    ws.row_dimensions[3].height = 40
    ws.row_dimensions[10].height = 28
    ws.row_dimensions[11].hidden = True; ws.row_dimensions[11].outlineLevel = 2
    ws.column_dimensions["AA"].width = 22
    wb.save(source)
    digest = hashlib.sha256(source.read_bytes()).hexdigest()
    operations = [dict(operation="delete_rows", sheet="数据", first_row=3, count=2)]
    request = dict(schema_version=worker.SCHEMA_VERSION, request_id="row-dimensions", input=str(source), output=str(output), format="xlsx",
                   plan=dict(request_id="row-dimensions", workbook=dict(inventory_complete=True), operations=operations))
    result = worker._apply(request)
    saved = openpyxl.load_workbook(output); s = saved["数据"]
    assert s["A3"].value == 5
    assert s.row_dimensions[8].height == 28
    assert s.row_dimensions[9].hidden and s.row_dimensions[9].outlineLevel == 2
    assert s.row_dimensions[3].height is None
    assert s.column_dimensions["AA"].width == 22
    assert hashlib.sha256(source.read_bytes()).hexdigest() == digest
    saved.close()
    (root / "summary.json").write_text(json.dumps(dict(status="PASS", assertions=6, source_sha256=digest, result=result), indent=2), encoding="utf8")
    print(json.dumps(dict(status="PASS", assertions=6)))


if __name__ == "__main__": main()
