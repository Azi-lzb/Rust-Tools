"""Serial resident-process API benchmarks. All observations are durable JSONL."""
import argparse
import csv
import hashlib
import json
import os
import platform
import statistics
import subprocess
import sys
import time
from pathlib import Path
from compare.semantic import verify

BASE = Path(__file__).resolve().parent


class Worker:
    def __init__(self, cmd, log):
        self.stderr = open(log, 'w', encoding='utf-8')
        t = time.perf_counter_ns()
        self.p = subprocess.Popen(cmd, stdin=subprocess.PIPE, stdout=subprocess.PIPE, stderr=self.stderr,
                                  text=True, encoding='utf-8', bufsize=1)
        line = self.p.stdout.readline()
        self.ready_ms = (time.perf_counter_ns()-t)/1e6
        self.ready = json.loads(line)
        if not self.ready.get('ready'):
            raise RuntimeError(line)

    def request(self, q):
        t = time.perf_counter_ns()
        self.p.stdin.write(json.dumps(q)+'\n')
        self.p.stdin.flush()
        line = self.p.stdout.readline()
        response = json.loads(line)
        response['request_e2e_ms'] = (time.perf_counter_ns()-t)/1e6
        return response

    def close(self):
        if self.p.poll() is None:
            self.p.stdin.close()
            self.p.wait(timeout=30)
        self.p.stdout.close()
        self.stderr.close()


def summarize(records, output):
    rows = []
    for size in sorted({r['id'] for r in records}):
        for op in ('read', 'write', 'roundtrip'):
            for engine in ('rust', 'openpyxl', 'aspose'):
                group = [r for r in records if r['id']==size and r['op']==op and r['engine']==engine and r['phase']=='measured']
                good = [r for r in group if r.get('ok') and r.get('validation', {}).get('ok')]
                values = [r['processing_ms'] for r in good]
                if not group:
                    continue
                rows.append(dict(size=size, operation=op, engine=engine, n=len(values), total=len(group),
                                 median_ms=statistics.median(values) if values else None,
                                 min_ms=min(values) if values else None, max_ms=max(values) if values else None,
                                 status='PASS' if len(good)==len(group)==10 else 'INCOMPLETE_OR_FAILED'))
    (output/'summary.json').write_text(json.dumps(rows, indent=2), 'utf-8')
    with (output/'summary.csv').open('w', newline='', encoding='utf-8-sig') as f:
        w = csv.DictWriter(f, fieldnames=list(rows[0]))
        w.writeheader(); w.writerows(rows)
    return rows


def run(args):
    output = Path(args.output).resolve()
    output.mkdir(parents=True, exist_ok=False)
    raw = output/'raw'
    raw.mkdir()
    fixtures = json.loads(Path(args.manifest).read_text('utf-8'))
    commands = dict(rust=[args.rust], openpyxl=[sys.executable, '-u', str(BASE/'adapters/python/worker.py')], aspose=[args.aspose])
    env = dict(time=time.strftime('%Y-%m-%dT%H:%M:%S%z'), platform=platform.platform(),
               python=sys.version, executable=sys.executable, commands=commands,
               warmups=3, measured_rounds=10, clocks='perf_counter_ns / Instant / Stopwatch',
               disk_cache='OS managed; not flushed', timing='resident API spans; no IPC, init, digest or validation',
               fixtures=fixtures)
    (output/'environment.json').write_text(json.dumps(env,indent=2), 'utf-8')
    cold = []
    # Startup probes perform no workbook operation and do not consume evaluation workbook budget.
    for name, cmd in commands.items():
        for i in range(3):
            w = Worker(cmd, output/f'{name}-cold-{i}.stderr.txt')
            cold.append(dict(engine=name, sample=i, launch_to_ready_ms=w.ready_ms, ready=w.ready))
            w.close()
    (output/'cold-start.json').write_text(json.dumps(cold, indent=2), 'utf-8')
    workers = {}
    records = []
    try:
        for name, cmd in commands.items():
            w = Worker(cmd, output/f'{name}.stderr.txt')
            workers[name] = w
            response = w.request(dict(op='init', fixtures=fixtures))
            if not response['ok']:
                raise RuntimeError(response)
            env[name] = dict(ready=w.ready, launch_to_ready_ms=w.ready_ms, preload=response)
        (output/'environment.json').write_text(json.dumps(env,indent=2), 'utf-8')
        with (output/'rounds.jsonl').open('w', encoding='utf-8', buffering=1) as f:
            for fixture in fixtures:
                for op in ('read', 'write', 'roundtrip'):
                    for i in range(13):
                        # Rotate engine order deterministically, no concurrent load.
                        names = list(workers)
                        names = names[i%3:]+names[:i%3]
                        for name in names:
                            q = dict(op=op, id=fixture['id'], input=fixture['input'])
                            if op != 'read':
                                q['output'] = str(raw/f'{name}-{fixture["id"]}-{op}-{i:02}.xlsx')
                            result = workers[name].request(q)
                            r = dict(engine=name, id=fixture['id'], op=op, round=i,
                                     phase='warmup' if i<3 else 'measured', **result)
                            if result.get('ok'):
                                try:
                                    if result['digest'] != fixture['digest']:
                                        raise AssertionError('worker materialized digest mismatch')
                                    r['validation'] = verify(q['output'],fixture,name=='aspose') if op != 'read' else dict(ok=True,digest=result['digest'])
                                except Exception as exc:
                                    r['validation'] = dict(ok=False,error=repr(exc))
                            records.append(r)
                            f.write(json.dumps(r)+'\n')
                            print(f'{fixture["id"]} {op} {i+1}/13 {name}: {result.get("processing_ms", "ERROR")} ms validate={r.get("validation", {}).get("ok")}', flush=True)
                            if not result.get('ok') or not r.get('validation',{}).get('ok'):
                                raise RuntimeError(f'Benchmark stopped: {r}')
        checks = [{ 'id': x['id'], 'unchanged': hashlib.sha256(Path(x['input']).read_bytes()).hexdigest()==x['sha256']} for x in fixtures]
        (output/'input-integrity.json').write_text(json.dumps(checks,indent=2),'utf-8')
        assert all(x['unchanged'] for x in checks)
    finally:
        for w in workers.values():
            w.close()
        if records:
            summarize(records,output)
    return output


if __name__ == '__main__':
    parser = argparse.ArgumentParser()
    for key in ('manifest','rust','aspose','output'):
        parser.add_argument('--'+key, required=True)
    print(run(parser.parse_args()))
