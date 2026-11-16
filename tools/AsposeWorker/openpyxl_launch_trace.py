"""Benchmark-only launch counter; executes the unmodified production entry point."""
import json
import os
from pathlib import Path
import runpy
import sys

trace = os.environ.get("ASPOSE_BENCH_OPENPYXL_TRACE")
if trace:
    with open(trace, "a", encoding="utf8") as f:
        f.write(json.dumps(dict(pid=os.getpid(), arguments=sys.argv[1:])) + "\n")
runpy.run_path(str(Path(__file__).resolve().parents[2] / "crates/excel-editor/python/openpyxl_worker.py"), run_name="__main__")
