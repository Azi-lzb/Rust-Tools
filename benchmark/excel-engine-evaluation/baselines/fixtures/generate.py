"""Fixed-seed scalar matrices; target real compressed XLSX bytes, never ZIP padding."""
import hashlib
import json
import random
import string
import sys
from pathlib import Path
import openpyxl

BASE = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(BASE / 'adapters/python'))
from worker import digest


def matrix(rows, seed=20261114):
    rng = random.Random(seed)
    alphabet = string.ascii_letters + string.digits
    return [[r, ''.join(rng.choices(alphabet, k=80)), r % 2 == 0, None,
             ''.join(rng.choices(alphabet, k=80)), r * 17 - 12345,
             f'{r:012d}', ''.join(rng.choices(alphabet, k=80))] for r in range(rows)]


def save(path, data):
    wb = openpyxl.Workbook(write_only=True)
    ws = wb.create_sheet('Data')
    for row in data:
        ws.append(row)
    wb.save(path)


def generate(root, targets=(1, 10)):
    root = Path(root)
    root.mkdir(parents=True, exist_ok=True)
    result = []
    for size in targets:
        target = size * 1024 * 1024
        rows = 4500 * size
        path = root / f'{size}MiB.xlsx'
        for attempt in range(5):
            data = matrix(rows)
            save(path, data)
            actual = path.stat().st_size
            if abs(actual / target - 1) <= 0.05:
                break
            rows = max(1, round(rows * target / actual))
        else:
            raise AssertionError('size calibration failed')
        data_path = root / f'{size}MiB.matrix.json'
        data_path.write_text(json.dumps(data, separators=(',', ':')), 'utf-8')
        result.append(dict(id=f'{size}MiB', input=str(path.resolve()), data=str(data_path.resolve()),
                           bytes=actual, rows=rows, cols=8, positions=rows*8,
                           nonempty_cells=rows*7, seed=20261114, sheets=['Data'],
                           digest=digest(data), sha256=hashlib.sha256(path.read_bytes()).hexdigest()))
    (root / 'manifest.json').write_text(json.dumps(result, indent=2), 'utf-8')
    return result


if __name__ == '__main__':
    print(json.dumps(generate(sys.argv[1]), indent=2))
