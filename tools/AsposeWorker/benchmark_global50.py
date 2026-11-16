"""Existing 1-8 policies exercised through real business entries, never no-op timing."""
import argparse
import json
from pathlib import Path
import re
import traceback

import benchmark_menu50 as bench


def traces(directory):
    result=[]
    for p in directory.glob("stderr-*.txt"):
        result += re.findall(r"ExcelEditPlan 路由：模式 ([^，]+)，引擎 ([^，]+)，",p.read_text(encoding="utf8"))
    return result


def confirmed(directory, engine, minimum):
    records=traces(directory)
    if len(records) < minimum: return False, records
    if any(mode != bench.GLOBAL_POLICIES[engine] for mode,_ in records):return False, records
    needle={"umya":"umya","betteroffice_modify":"betteroffice","rustnative":"rustnative"}.get(engine)
    if needle and any(needle not in re.sub(r"[-_]","",name.lower()) for _,name in records): return False,records
    return True,records


def main():
    p=argparse.ArgumentParser();p.add_argument("--output",required=True);p.add_argument("--references",required=True)
    p.add_argument("--exe",default="F:/RustBuild/pytools/release/RustTools.exe");a=p.parse_args()
    root=Path(a.output).resolve();root.mkdir(parents=True,exist_ok=False)
    refs=json.loads(Path(a.references).read_text(encoding="utf8"));results=[]
    for ref in refs:
        case=ref["case"];entry=dict(case=case,baseline=ref["baseline"],routes=[])
        baseline1=json.loads(Path(ref["preflight_snapshot"]).read_text(encoding="utf8"))
        baseline50=json.loads(Path(ref["batch_snapshot"]).read_text(encoding="utf8"))
        for engine in bench.GLOBAL_POLICIES:
            record=dict(engine=engine,status="NOT_TESTED")
            try:
                directory=root/case/"preflight"/engine
                info,view=bench.execute(a.exe,Path(ref["canonical1"]),directory,case,engine,1)
                ok,observed=confirmed(directory,engine,1)
                record["observed_routes"]=observed
                if not ok:
                    record.update(status="NOT_APPLICABLE",reason="Requested policy was not observed in the completed editing route")
                    continue
                diff,allowed=bench.compare(case,baseline1,view)
                bench.dump(directory/"DIFFERENCES.json",dict(differences=diff,approved=allowed))
                assert not diff,diff[:5]
                print(json.dumps(dict(case=case,engine=engine,stage="preflight",status="PASS")),flush=True)
                directory=root/case/"batch50"/engine
                info,view=bench.execute(a.exe,Path(ref["canonical50"]),directory,case,engine,50)
                ok,observed=confirmed(directory,engine,50)
                assert ok,("Actual completed policy/backend not confirmed for 50 outputs",observed)
                diff,allowed=bench.compare(case,baseline50,view)
                bench.dump(directory/"DIFFERENCES.json",dict(differences=diff,approved=allowed))
                assert not diff,diff[:5]
                record.update(info,status="PASS",evidence=str(directory),observed_routes=observed)
            except Exception:record.update(status="GATED",error=traceback.format_exc())
            finally:
                entry["routes"].append(record)
                bench.dump(root/"RESULTS.json",dict(results=results+[entry]))
                print(json.dumps(dict(case=case,engine=engine,status=record["status"],total_e2e_ms=record.get("total_e2e_ms"))),flush=True)
        results.append(entry)
    bench.dump(root/"RESULTS.json",dict(results=results))


if __name__ == "__main__":main()
