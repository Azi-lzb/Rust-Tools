import json
import sys
from pathlib import Path
from run import Worker,calc_validation,target_validation

root=Path(sys.argv[1]).resolve();out=root/'smoke';out.mkdir(exist_ok=False)
m=json.loads((root/'fixtures/manifest.json').read_text('utf-8'))
commands=dict(rust=['F:/RustBuild/pytools/release/four_dim.exe'],aspose=['F:/RustBuild/pytools/fourdim-dotnet/FourDim.exe'],
              openpyxl=[sys.executable,'-u',str(Path(__file__).parent/'python_worker.py')],com=[sys.executable,'-u',str(Path(__file__).parent/'python_worker.py'),'--com'])
workers={};results=[]
try:
    for name,cmd in commands.items():
        try:workers[name]=Worker(cmd,out/(name+'.stderr.txt'))
        except Exception as exc:
            results.append(dict(engine=name,status='UNAVAILABLE',error=repr(exc)));continue
    for name in ('rust','aspose'):
        f=m['read'][0];r=workers[name].request(dict(op='read',input=f['input'],rows=f['rows'],cols=f['cols']))
        assert r['ok'] and r['digest']==f['digest'],r
        path=out/(name+'-write.xlsx');r=workers[name].request(dict(op='write',output=str(path)))
        assert r['ok'],r
        check=workers['rust'].request(dict(op='verify',input=str(path),rows=50000,cols=20));assert check['ok'] and check['digest']==m['write']['digest'],check
        results.append(dict(engine=name,op='write',result=r,validation=check))
    for name in workers:
        if name=='openpyxl':continue
        r=workers[name].request(dict(op='calc',input=str(root/'fixtures/calc-stale.xlsx')))
        check=calc_validation(r.get('values',[]),m['calc_expected'])
        results.append(dict(engine=name,op='calc',result=r,validation=check))
        print(name,'calc',check,flush=True)
    for name in ('rust','openpyxl','aspose'):
        path=out/(name+'-modify.xlsx');r=workers[name].request(dict(op='modify',input=m['modify'],output=str(path)))
        assert r['ok'],r
        check=target_validation(path,'modify');assert check['ok'],check
        results.append(dict(engine=name,op='modify',result=r,validation=check))
finally:
    for w in workers.values():w.close()
    (out/'results.json').write_text(json.dumps(results,indent=2),'utf-8')
print('SMOKE_COMPLETE',flush=True)
