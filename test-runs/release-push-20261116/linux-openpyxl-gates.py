import json, os, re, subprocess, sys
from pathlib import Path
r=Path('/mnt/f/RustBuild/pytools/release-push-20261116' if sys.argv[1]=='wsl' else '/evidence')
platform=sys.argv[1]
s=(r/f'{platform}-workspace.log').read_text(errors='replace')
paths=re.findall(r'Running tests/openpyxl_worker\.rs \(([^)\r\n]+)\)',s)
assert len(paths)==1,paths
log=r/f'{platform}-openpyxl-gates.log'
with log.open('w') as f:
    p=subprocess.run([paths[0],'--ignored','--test-threads=1'],env=os.environ.copy(),stdout=f,stderr=subprocess.STDOUT,timeout=900)
s=log.read_text(errors='replace')
m=re.search(r'test result: (?:ok|FAILED)\. (\d+) passed; (\d+) failed; (\d+) ignored;',s)
result={'exit':p.returncode,'passed':int(m[1]) if m else None,'failed':int(m[2]) if m else None,'log':log.name,'command':[paths[0],'--ignored','--test-threads=1']}
(r/f'{platform}-openpyxl-gates.json').write_text(json.dumps(result,indent=2))
print(json.dumps(result),flush=True)
raise SystemExit(p.returncode)
