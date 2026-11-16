"""One serial four-dimensional run; resident API clocks are authoritative."""
import argparse
import hashlib
import json
import math
import os
import queue
import statistics
import subprocess
import sys
import threading
import time
from pathlib import Path

HERE=Path(__file__).resolve().parent
BASE=HERE.parent
sys.path.insert(0,str(BASE))


class Worker:
    def __init__(self,cmd,log,env=None):
        self.log=open(log,'w',encoding='utf-8');self.lines=queue.Queue()
        start=time.perf_counter_ns()
        self.p=subprocess.Popen(cmd,stdin=subprocess.PIPE,stdout=subprocess.PIPE,stderr=self.log,
                                text=True,encoding='utf-8',bufsize=1,env=env,creationflags=subprocess.CREATE_NO_WINDOW if os.name=='nt' else 0)
        def reader():
            for line in self.p.stdout:self.lines.put(line)
            self.lines.put(None)
        self.thread=threading.Thread(target=reader,daemon=True);self.thread.start()
        line=self.lines.get(timeout=60)
        self.launch_ms=(time.perf_counter_ns()-start)/1e6
        if not line:raise RuntimeError(f'worker failed startup: {cmd}')
        self.ready=json.loads(line)
        if not self.ready.get('ready'):raise RuntimeError(self.ready)

    def request(self,q):
        start=time.perf_counter_ns();self.p.stdin.write(json.dumps(q)+'\n');self.p.stdin.flush()
        line=self.lines.get(timeout=180)
        if not line:raise RuntimeError('worker exited')
        result=json.loads(line);result['request_e2e_ms']=(time.perf_counter_ns()-start)/1e6
        return result

    def close(self):
        if self.p.poll() is None:
            self.p.stdin.close()
            try:self.p.wait(timeout=20)
            except subprocess.TimeoutExpired:
                self.p.terminate();self.p.wait(timeout=10)
                pid=self.ready.get('owned_excel_pid')
                if pid:
                    # Only the dedicated process proven absent before DispatchEx.
                    try:os.kill(pid,15)
                    except OSError:pass
        self.thread.join(timeout=2);self.p.stdout.close();self.log.close()


def summary(records,root,name):
    groups={}
    for r in records:
        if r['phase']=='measured':groups.setdefault((r['condition'],r['engine']),[]).append(r)
    out=[]
    for (condition,engine),rows in groups.items():
        valid=[r for r in rows if r.get('ok') and r.get('validation',{}).get('ok')]
        times=[r['processing_ms'] for r in valid]
        out.append(dict(condition=condition,engine=engine,n=len(valid),total=len(rows),
                        median_ms=statistics.median(times) if times else None,min_ms=min(times) if times else None,max_ms=max(times) if times else None,
                        status='PASS' if len(valid)==len(rows)==10 else 'FAILED_OR_INCOMPLETE'))
    (root/(name+'.json')).write_text(json.dumps(out,indent=2),'utf-8')
    return out


def calc_validation(values,expected):
    bad=[]
    if len(values)!=len(expected):return dict(ok=False,error=f'count {len(values)} != {len(expected)}')
    for i,(a,b) in enumerate(zip(values,expected)):
        if not isinstance(a,(int,float)) or isinstance(a,bool) or not math.isfinite(a) or abs(a-b)>1e-8:bad.append(dict(index=i,actual=a,expected=b))
    return dict(ok=not bad,count=len(expected),mismatch_count=len(bad),examples=bad[:10])


def target_validation(path,op):
    import openpyxl
    start=time.perf_counter_ns();wb=openpyxl.load_workbook(path,read_only=True,data_only=False)
    try:
        if op=='modify':
            cell=wb['Data']['J12'];ok=cell.value==9999 and bool(cell.font.bold)
        else:
            import zipfile,xml.etree.ElementTree as ET
            with zipfile.ZipFile(path) as z:xml=ET.fromstring(z.read('xl/worksheets/sheet1.xml'))
            text=xml.findtext('{http://schemas.openxmlformats.org/spreadsheetml/2006/main}headerFooter/{http://schemas.openxmlformats.org/spreadsheetml/2006/main}oddHeader')
            ok=text is not None and 'edited' in text and 'benchmark' not in text
        return dict(ok=ok,scope='target change only; full structure analyzed separately',verify_ms=(time.perf_counter_ns()-start)/1e6)
    finally:wb.close()


def run(a):
    root=Path(a.root).resolve();fixture=Path(a.fixtures).resolve() if a.fixtures else root/'fixtures';manifest=json.loads((fixture/'manifest.json').read_text('utf-8'))
    evidence=root/'run';evidence.mkdir(exist_ok=False);raw=evidence/'raw';raw.mkdir()
    commands=dict(rust=[a.rust],aspose=[a.aspose],openpyxl=[sys.executable,'-u',str(HERE/'python_worker.py')],com=[sys.executable,'-u',str(HERE/'python_worker.py'),'--com'])
    clr_env=dict(os.environ)
    if a.no_tiered:clr_env.update(DOTNET_TieredCompilation='0',DOTNET_TieredPGO='0')
    env=dict(time=time.strftime('%Y-%m-%dT%H:%M:%S%z'),commands=commands,warmups=3,measured=10,
             fixtures_root=str(fixture),dotnet_profile='TieredCompilation=0; TieredPGO=0' if a.no_tiered else 'runtime defaults',
             modify_input=a.modify_input or manifest['modify'],
             timing='resident API clocks; Calc only recalculation; all file loading separately recorded; Write includes value generation/fill/save',
             clocks='Rust Instant; .NET Stopwatch; Python perf_counter_ns',disk_cache='OS managed; not flushed',aspose_budget=91)
    cold=[]
    for i in range(3):
        start=time.perf_counter_ns();p=subprocess.run([a.aspose,'--empty-run'],capture_output=True,text=True,encoding='utf-8',timeout=60,env=clr_env)
        ms=(time.perf_counter_ns()-start)/1e6
        if p.returncode:raise RuntimeError(p.stderr)
        cold.append(dict(sample=i,kind='DLL-load-and-exit-empty-run',elapsed_ms=ms,ready=json.loads(p.stdout)))
    (evidence/'empty-run.json').write_text(json.dumps(cold,indent=2),'utf-8')
    workers={};records=[]
    def add_worker(name):
        w=Worker(commands[name],evidence/(name+'.stderr.txt'),clr_env if name=='aspose' else None);workers[name]=w
        pings=[w.request(dict(op='ping')) for _ in range(5)]
        env[name]=dict(ready=w.ready,launch_to_ready_ms=w.launch_ms,ping=pings)
        (evidence/'environment.json').write_text(json.dumps(env,indent=2),'utf-8')
    conditions=[]
    for f in manifest['read']:conditions.append(dict(id='read-'+f['id'],op='read',engines=['rust','aspose'],input=f['input'],rows=f['rows'],cols=f['cols'],digest=f['digest']))
    conditions.append(dict(**manifest['write'],op='write',engines=['rust','aspose']))
    try:
        add_worker('rust');add_worker('aspose')
        with (evidence/'rounds.jsonl').open('w',encoding='utf-8',buffering=1) as log:
            def execute(condition):
                for i in range(13):
                    names=condition['engines'];names=names[i%len(names):]+names[:i%len(names)]
                    for name in names:
                        q={k:v for k,v in condition.items() if k in ('op','input','rows','cols')}
                        if q['op'] in ('write','modify','header'):q['output']=str(raw/f'{condition["id"]}-{name}-{i:02}.xlsx')
                        result=workers[name].request(q)
                        r=dict(condition=condition['id'],engine=name,op=q['op'],round=i,phase='warmup' if i<3 else 'measured',**result)
                        if q.get('output'):r['output']=q['output']
                        if result.get('ok'):
                            if q['op']=='read':r['validation']=dict(ok=result['digest']==condition['digest'],digest=result['digest'])
                            elif q['op']=='write':
                                v=workers['rust'].request(dict(op='verify',input=q['output'],rows=50000,cols=20))
                                r['validation']=dict(ok=v.get('ok') and v.get('digest')==condition['digest'],digest=v.get('digest'),verify_ms=v.get('processing_ms'))
                            elif q['op']=='calc':r['validation']=calc_validation(result.get('values',[]),manifest['calc_expected'])
                            else:r['validation']=target_validation(q['output'],q['op'])
                        else:r['validation']=dict(ok=False,error=result.get('error'))
                        records.append(r);log.write(json.dumps(r)+'\n')
                        print(f'{condition["id"]} {i+1}/13 {name}: {result.get("processing_ms")} ms PASS={r["validation"]["ok"]}',flush=True)
                        if q['op'] in ('read','write') and not r['validation']['ok']:raise RuntimeError('I/O validation failed; stop rather than publish misleading timings')
            for condition in conditions:execute(condition)
            first=summary(records,evidence,'read-write-summary');print('READ_WRITE_COMPLETE '+json.dumps(first),flush=True)
            # Finish implementing/inspecting structural checks without restarting workers.
            while not (root/'phase2.ready').exists():time.sleep(1)
            add_worker('openpyxl')
            com_available=True
            try:add_worker('com')
            except Exception as exc:
                com_available=False;(evidence/'com-unavailable.json').write_text(json.dumps(dict(status='NOT_TESTED',error=repr(exc))), 'utf-8')
            for ext in ('xls','xlsx'):
                engines=['aspose']+(['rust'] if ext=='xlsx' else [])+(['com'] if com_available else [])
                execute(dict(id='calc-'+ext,op='calc',engines=engines,input=str(fixture/('calc-stale.'+ext))))
            modify=a.modify_input or manifest['modify']
            execute(dict(id='modify-value-style',op='modify',engines=['rust','openpyxl','aspose'],input=modify))
            execute(dict(id='modify-header-narrow',op='header',engines=['rust','openpyxl','aspose'],input=modify))
            summary(records,evidence,'summary')
        checked={}
        for path,expected in json.loads((fixture/'sha256.json').read_text('utf-8')).items():checked[path]=hashlib.sha256(Path(path).read_bytes()).hexdigest()==expected
        for f in manifest['read']:checked[f['input']]=hashlib.sha256(Path(f['input']).read_bytes()).hexdigest()==f['sha256']
        (evidence/'input-integrity.json').write_text(json.dumps(checked,indent=2),'utf-8');assert all(checked.values())
    finally:
        for w in workers.values():w.close()
        if records:summary(records,evidence,'summary')
    print('ALL_TIMINGS_COMPLETE '+str(evidence),flush=True)


if __name__=='__main__':
    p=argparse.ArgumentParser()
    for n in ('root','rust','aspose'):p.add_argument('--'+n,required=True)
    p.add_argument('--fixtures');p.add_argument('--modify-input');p.add_argument('--no-tiered',action='store_true')
    run(p.parse_args())
