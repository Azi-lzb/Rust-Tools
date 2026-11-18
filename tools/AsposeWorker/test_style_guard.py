"""Keep the xlwt/Aspose style-loss regression: reject, never silently succeed."""
import hashlib
import importlib.util
import json
import sys
from pathlib import Path
import openpyxl

def main():
    root=Path(sys.argv[1]).resolve();root.mkdir(exist_ok=False)
    input_path=Path(sys.argv[2]).resolve();before=hashlib.sha256(input_path.read_bytes()).hexdigest()
    path=Path(__file__).resolve().parents[2]/"benchmark/excel-engine-evaluation/baselines/four_dim/run.py"
    spec=importlib.util.spec_from_file_location("benchmark_worker_harness",path)
    module=importlib.util.module_from_spec(spec);spec.loader.exec_module(module)
    worker=module.Worker([sys.argv[3],sys.argv[4],"--stdio"],root/"worker.stderr.jsonl")
    candidate=root/"rejected-candidate.xlsx"
    request=dict(protocol_version=1,request_id="style-regression",run_id="known-style-loss",method="Modify",verify_content=True,
                 input=str(input_path),candidate=str(candidate),operations=[
                     dict(op="set_value",sheet="Data",cell="AA100",value=dict(type="number",value=123)),
                     dict(op="set_style",sheet="Data",cell="AA100",style=dict(bold=True,font_size=14))])
    try: response=worker.request(request)
    finally:worker.close()
    assert response["ok"] is False and response["code"]=="OUTPUT_VALIDATION_FAILED",response
    assert "Data!AA100" in response["message"]
    assert hashlib.sha256(input_path.read_bytes()).hexdigest()==before
    wb=openpyxl.load_workbook(candidate)
    assert wb["Data"]["AA100"].font.bold is False  # Original failure remains visible.
    wb.close()
    report=dict(status="PASS_SAFE_REJECTION",underlying_style_loss_fixed=False,
                source_unchanged=True,response=response,candidate=str(candidate),candidate_published=False)
    (root/"RESULT.json").write_text(json.dumps(report,indent=2),"utf-8")
    print(json.dumps(report))

if __name__=="__main__":main()
