"""Real unlicensed Worker acceptance; evaluation output is explicitly accepted.

Only isolated synthetic files are used. No production menu/routing is invoked.
"""
import hashlib
import importlib.util
import json
import math
import os
import statistics
import shutil
import subprocess
import sys
import time
from pathlib import Path

import openpyxl
import xlwt
from test_memory_save import fixture

HERE = Path(__file__).resolve().parent
REPO = HERE.parents[1]
spec = importlib.util.spec_from_file_location("benchmark_worker_harness", REPO / "benchmark/excel-engine-evaluation/baselines/four_dim/run.py")
harness = importlib.util.module_from_spec(spec)
spec.loader.exec_module(harness)


def digest(path):
    return hashlib.sha256(path.read_bytes()).hexdigest()


def main():
    root = Path(sys.argv[1]).resolve()
    root.mkdir(exist_ok=False)
    dotnet, dll, host = sys.argv[2:5]
    fixture(root / "fixture")
    manifest = json.loads((root / "fixture/fixture.json").read_text())
    source = Path(manifest["input"])
    if len(sys.argv)>5:
        # Preserve the original xlwt fixture and failure evidence; independently
        # authored Excel XLS is a different, explicitly named acceptance cohort.
        source=root/"fixture/excel-authored-complex.xls"
        shutil.copyfile(sys.argv[5],source)
        manifest.update(input=str(source),sha256=digest(source),author="Excel-authored synthetic XLS; original xlwt fixture retained")
        (root/"acceptance-fixture.json").write_text(json.dumps(manifest,indent=2),"utf-8")
    before = digest(source)
    worker = harness.Worker([dotnet, dll, "--stdio"], root / "worker.stderr.jsonl")
    results = []
    counter = 0
    def call(method="Modify", input_path=source, output=None, operations=None, expected_ok=True, **extra):
        nonlocal counter
        counter += 1
        output = output or root / f"direct-{counter}.xlsx"
        q = dict(protocol_version=1,request_id=str(counter),run_id="resident-acceptance",method=method,
                 input=str(input_path),candidate=str(output))
        if operations is not None:q["operations"] = operations
        q.update(extra)
        result = worker.request(q)
        assert result["ok"] is expected_ok, result
        assert result["request_id"] == str(counter)
        if not expected_ok: assert not output.exists() or output.name == "existing.xlsx"
        results.append(dict(request=q,response=result,expected_ok=expected_ok))
        return result, output
    try:
        calculated, calc_path = call("Calculate")
        wb = openpyxl.load_workbook(calc_path,data_only=True)
        actual=[c.value for row in wb["Calc"] for c in row]
        assert len(actual)==2048 and all(isinstance(a,(int,float)) and math.isclose(a,b,abs_tol=1e-8) for a,b in zip(actual,manifest["expected"]))
        assert wb.sheetnames==["Data","Calc","Evaluation Warning"] and wb.active.title=="Evaluation Warning"
        wb.close()
        operations=[dict(op="set_value",sheet="Data",cell=cell,value=dict(type=kind,value=value)) for cell,kind,value in
                    [("E10","text","0012"),("E11","number",0),("E12","boolean",False),("E13","text","clear me"),("Z99","text","boundary"),("AA100","number",123)]]
        operations += [dict(op="set_value",sheet="Data",cell="E13",value=dict(type="blank")),
                       dict(op="set_formula",sheet="Data",cell="F10",formula="=SUM(Data!B2:B4)"),
                       dict(op="set_style",sheet="Data",cell="AA100",style=dict(bold=True,italic=True,font_name="Arial",font_size=14,fill_rgb="FFE699",number_format="0.00",wrap_text=True))]
        modified, modified_path=call(operations=operations)
        wb=openpyxl.load_workbook(modified_path,data_only=False)
        ws=wb["Data"]
        assert ws["E10"].value=="0012" and ws["E10"].data_type=="s"
        assert ws["E11"].value==0 and ws["E12"].value is False and ws["E13"].value is None
        assert ws["Z99"].value=="boundary" and ws["AA100"].value==123
        assert ws["F10"].value=="=SUM(Data!B2:B4)"
        assert ws["AA100"].font.bold and ws["AA100"].font.italic and ws["AA100"].font.sz==14
        assert ws["AA100"].fill.fgColor.rgb=="FFFFE699" and ws["AA100"].number_format=="0.00" and ws["AA100"].alignment.wrap_text
        wb.close()
        chained,chained_path=call("Calculate",input_path=modified_path)
        wb=openpyxl.load_workbook(chained_path,data_only=True)
        assert wb["Data"]["F10"].value==6
        wb.close()
        for address in ["A0","AA00","XFE1","A1048577"]:
            call(operations=[dict(op="set_value",sheet="Data",cell=address,value=dict(type="number",value=1))],expected_ok=False)
        call(operations=[dict(op="set_style",sheet="Data",cell="A1",style=dict(unknown_style=True))],expected_ok=False)
        call(operations=[dict(op="set_value",sheet="Missing",cell="A1",value=dict(type="text",value="x"))],expected_ok=False)
        call(operations=[dict(op="set_value",sheet="Data",cell="A1",value=dict(type="date",value="2026-11-15"))],expected_ok=False)
        call(operations=[dict(op="set_formula",sheet="Data",cell="A1",formula="not a formula")],expected_ok=False)
        existing=root/"existing.xlsx";existing.write_bytes(b"sentinel target")
        call("Calculate",output=existing,expected_ok=False)
        assert existing.read_bytes()==b"sentinel target"
        call("Calculate",output=root/"missing-parent/output.xlsx",expected_ok=False)
        corrupt=root/"corrupt.xls";corrupt.write_bytes(b"not BIFF")
        call("Calculate",input_path=corrupt,expected_ok=False)
        named=root/"business-warning-name.xls"
        wb=xlwt.Workbook();wb.add_sheet("Data");wb.add_sheet("Evaluation Warning").write(0,0,"user-owned original text");wb.save(str(named))
        named_hash=digest(named)
        call("Calculate",input_path=named,expected_ok=False)
        assert digest(named)==named_hash
        ping=worker.request(dict(protocol_version=1,request_id="ping",run_id="health",method="Ping"))
        assert ping["ok"] and ping["pid"]==worker.ready["pid"]
        assert digest(source)==before
    finally:worker.close()
    (root/"direct-results.json").write_text(json.dumps(results,indent=2),"utf-8")
    print(f"DIRECT_PASS: {len(results)} real tasks (including expected rejections), typed values/styles/formula and chain verified",flush=True)

    # Actual Rust -> .NET -> Rust acceptance and a real 99/100 boundary.
    tiny=root/"tiny.xlsx";wb=openpyxl.Workbook();wb.active.title="Data";wb["Data"]["A1"]=1;wb.create_sheet("Calc")["A1"]="=Data!A1+1";wb.active=1;wb.save(tiny);wb.close()
    tiny_before=digest(tiny)
    def host_run(name, input_path, count, fresh=False):
        out=root/name
        cmd=[host,dotnet,dll,str(input_path),str(out),str(count)]+(["fresh"] if fresh else [])
        start=time.perf_counter_ns()
        p=subprocess.run(cmd,capture_output=True,text=True,encoding="utf-8",timeout=180)
        wall=(time.perf_counter_ns()-start)/1e6
        (root/(name+".stdout.txt")).write_text(p.stdout,"utf-8")
        (root/(name+".stderr.txt")).write_text(p.stderr,"utf-8")
        assert p.returncode==0,(name,p.stdout,p.stderr)
        report=json.loads((out/"host-report.json").read_text())
        assert report["continued_in_rust"] and len(report["records"])==count
        for i,r in enumerate(report["records"]):
            wb=openpyxl.load_workbook(r["path"],data_only=i==0)
            assert wb.sheetnames==["Data","Calc","Evaluation Warning"] and wb.active.title=="Evaluation Warning"
            assert r["evaluation_output"] and r["source_active_index"]==1 and r["output_active_index"]==2
            if i>0:assert wb["Data"]["D12"].value==9999 and wb["Data"]["D12"].font.bold
            else:
                values=[c.value for row in wb["Calc"] for c in row]
                expected=manifest["expected"] if input_path==source else [2]
                assert values==expected
            wb.close()
        report["wall_including_rust_process_ms"]=wall
        return report
    boundary=host_run("boundary-100",tiny,100)
    assert boundary["generations"]==2
    records=boundary["records"]
    assert len({r["worker_pid"] for r in records[:99]})==1 and records[98]["draining"] and records[98]["worker_attempts"]==99
    assert records[99]["worker_attempts"]==1 and records[99]["worker_pid"]!=records[98]["worker_pid"]
    assert digest(tiny)==tiny_before
    print("BOUNDARY_PASS: 100 real Rust requests, 2 workers, response99 delivered before request100",flush=True)
    bench=[]
    for i in range(3):
        order=[False,True] if i%2==0 else [True,False]
        for fresh in order:
            name=f"bench-{i}-"+("fresh" if fresh else "resident")
            report=host_run(name,source,5,fresh)
            bench.append(dict(round=i,fresh=fresh,total_ms=report["wall_including_rust_process_ms"],host_ms=report["total_including_shutdown_ms"],generations=report["generations"],records=report["records"]))
            print(f"BENCH {name}: {report['wall_including_rust_process_ms']:.2f} ms",flush=True)
    summary=dict(status="PASS_EVALUATION_OUTPUT_ACCEPTED",direct_tasks=len(results),boundary_tasks=100,boundary_generations=2,
                 benchmark_batches=6,tasks_per_batch=5,benchmark=[dict(fresh=f,median_total_ms=statistics.median(x["total_ms"] for x in bench if x["fresh"]==f),min_total_ms=min(x["total_ms"] for x in bench if x["fresh"]==f),max_total_ms=max(x["total_ms"] for x in bench if x["fresh"]==f)) for f in (False,True)],
                 source_unchanged=digest(source)==before,watermark_removed=False,active_sheet_restored=False,
                 evaluation_behavior="Explicitly accepted by user; not counted as misalignment",production_route_changed=False)
    assert summary["source_unchanged"]
    (root/"benchmark-rounds.json").write_text(json.dumps(bench,indent=2),"utf-8")
    (root/"RESULT.json").write_text(json.dumps(summary,indent=2),"utf-8")
    print(json.dumps(summary),flush=True)


if __name__=="__main__":main()
