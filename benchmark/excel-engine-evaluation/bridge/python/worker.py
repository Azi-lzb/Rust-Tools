"""Experiment-only, path-based openpyxl worker. Never imported by production code."""
import json
import os
import sys
import tempfile
import time
import traceback

def execute(request_path: str, response_path: str) -> int:
    request_id = "unknown"
    try:
        with open(request_path, "r", encoding="utf-8") as stream:
            request = json.load(stream)
        request_id = request["request_id"]
        input_path = os.path.abspath(request["input_path"])
        output_path = os.path.abspath(request["output_path"])
        if os.path.normcase(input_path) == os.path.normcase(output_path):
            raise ValueError("input_path and output_path must differ")
        start = time.perf_counter()
        import openpyxl
        book = openpyxl.load_workbook(input_path, read_only=False, data_only=False, keep_links=True)
        sheet = book[request.get("sheet_name") or book.sheetnames[0]]
        for values in request["append_rows"]:
            sheet.append(values)
        output_dir = os.path.dirname(output_path)
        fd, candidate = tempfile.mkstemp(prefix=".openpyxl-candidate-", suffix=".xlsx", dir=output_dir)
        os.close(fd)
        try:
            book.save(candidate)
            book.close()
            os.replace(candidate, output_path)
        finally:
            if os.path.exists(candidate):
                os.unlink(candidate)
        verify = openpyxl.load_workbook(output_path, read_only=True, data_only=False)
        check = verify[request.get("sheet_name") or verify.sheetnames[0]]
        expected = request["append_rows"][-1] if request["append_rows"] else None
        actual = tuple(check.iter_rows(min_row=check.max_row, max_row=check.max_row, values_only=True))[0] if expected else None
        verify.close()
        if expected and tuple(actual[:len(expected)]) != tuple(expected):
            raise RuntimeError("output readback mismatch")
        result = {"request_id": request_id, "ok": True, "appended_rows": len(request["append_rows"]),
                  "worker_total_ms": (time.perf_counter() - start) * 1000,
                  "openpyxl_version": openpyxl.__version__}
        code = 0
    except Exception as exc:
        result = {"request_id": request_id, "ok": False, "error": str(exc), "trace": traceback.format_exc(limit=4)}
        code = 1
    with open(response_path, "w", encoding="utf-8") as stream:
        json.dump(result, stream, ensure_ascii=False, separators=(",", ":"))
    return code

if len(sys.argv) == 3:
    raise SystemExit(execute(sys.argv[1], sys.argv[2]))
raise SystemExit("usage: worker.py REQUEST.json RESPONSE.json")
