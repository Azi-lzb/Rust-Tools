"""Corrected-fixture modify-only cohort; prior measurements remain untouched."""
import json
import sys
from pathlib import Path
from run import Worker,summary,target_validation

root=Path(sys.argv[1]).resolve();evidence=root/'modify-corrected';evidence.mkdir(exist_ok=False);raw=evidence/'raw';raw.mkdir()
source=root/'fixtures/modify-complex-workbook-customxml.xlsx'
commands=dict(rust=['F:/RustBuild/pytools/release/four_dim.exe'],aspose=['F:/RustBuild/pytools/fourdim-dotnet/FourDim.exe'],openpyxl=[sys.executable,'-u',str(Path(__file__).parent/'python_worker.py')])
workers={};records=[];env={}
try:
    for name,cmd in commands.items():
        w=Worker(cmd,evidence/(name+'.stderr.txt'));workers[name]=w;env[name]=dict(ready=w.ready,launch_to_ready_ms=w.launch_ms,command=cmd)
    (evidence/'environment.json').write_text(json.dumps(env,indent=2),'utf-8')
    with (evidence/'rounds.jsonl').open('w',encoding='utf-8',buffering=1) as log:
        for op,condition in [('modify','modify-value-style'),('header','modify-header-narrow')]:
            for i in range(13):
                names=list(workers);names=names[i%3:]+names[:i%3]
                for name in names:
                    path=raw/f'{condition}-{name}-{i:02}.xlsx'
                    result=workers[name].request(dict(op=op,input=str(source),output=str(path)))
                    r=dict(condition=condition,op=op,engine=name,round=i,phase='warmup' if i<3 else 'measured',output=str(path),**result)
                    r['validation']=target_validation(path,op) if result.get('ok') else dict(ok=False,error=result.get('error'))
                    records.append(r);log.write(json.dumps(r)+'\n')
                    print(condition,i+1,name,result.get('processing_ms'),'PASS',r['validation']['ok'],flush=True)
                    if not r['validation']['ok']:raise RuntimeError(r)
finally:
    for w in workers.values():w.close()
    if records:summary(records,evidence,'summary')
print('CORRECTED_MODIFY_COMPLETE',flush=True)
