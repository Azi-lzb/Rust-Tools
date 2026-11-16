"""Independent library baseline, never imported by production workers."""
import hashlib
import json
import os
import sys
import time
from pathlib import Path
import openpyxl


def digest(data):
    h = hashlib.sha256()
    for row in data:
        for v in row:
            if v is None:
                s = 'E\n'
            elif isinstance(v, bool):
                s = f'B{int(v)}\n'
            elif isinstance(v, (int, float)):
                if int(v) != v:
                    raise ValueError('non-integer fixture number')
                s = f'N{int(v)}\n'
            elif isinstance(v, str):
                s = f'S{v}\n'
            else:
                raise ValueError(f'unsupported type: {type(v)}')
            h.update(s.encode('utf-8'))
    return h.hexdigest()


def main():
    print(json.dumps(dict(ready=True, engine='openpyxl', version=openpyxl.__version__, pid=os.getpid())), flush=True)
    sets = {}
    for line in sys.stdin:
        try:
            q = json.loads(line)
            if q['op'] == 'init':
                sets = {f['id']: json.loads(Path(f['data']).read_text('utf-8')) for f in q['fixtures']}
                result = dict(ok=True)
            else:
                data = sets[q['id']]
                out = q.get('output')
                if out and Path(out).exists():
                    raise ValueError('refuse existing output')
                start = time.perf_counter_ns()
                if q['op'] == 'read':
                    wb = openpyxl.load_workbook(q['input'], read_only=True, data_only=False)
                    actual = [list(r) for r in wb['Data'].iter_rows(min_row=1, max_row=len(data), max_col=len(data[0]), values_only=True)]
                    wb.close()
                elif q['op'] == 'write':
                    wb = openpyxl.Workbook(write_only=True)
                    ws = wb.create_sheet('Data')
                    for row in data:
                        ws.append(row)
                    wb.save(out)
                    actual = data
                elif q['op'] == 'roundtrip':
                    wb = openpyxl.load_workbook(q['input'])
                    wb.save(out)
                    wb.close()
                    actual = data  # External verifier re-reads saved output, outside API timer.
                else:
                    raise ValueError('unknown op')
                ms = (time.perf_counter_ns() - start) / 1e6
                result = dict(ok=True, processing_ms=ms, digest=digest(actual), rows=len(data), cols=len(data[0]))
        except Exception as exc:
            result = dict(ok=False, error=repr(exc))
        print(json.dumps(result), flush=True)


if __name__ == '__main__':
    main()
