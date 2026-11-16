"""External full scalar validator. Raw evaluation workbooks are never changed."""
import sys
import time
from pathlib import Path
import openpyxl

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / 'adapters/python'))
from worker import digest


def verify(path, fixture, aspose=False):
    start = time.perf_counter_ns()
    wb = openpyxl.load_workbook(path, read_only=True, data_only=False)
    try:
        missing = set(fixture['sheets']) - set(wb.sheetnames)
        if missing:
            raise AssertionError(f'missing original sheets: {missing}')
        excluded = []
        for name in set(wb.sheetnames) - set(fixture['sheets']):
            # Input identity plus exact vendor marker: never generic name matching.
            ws = wb[name]
            texts = [v for row in ws.iter_rows(max_row=10, max_col=5, values_only=True) for v in row if isinstance(v, str)]
            if aspose and name == 'Evaluation Warning' and any('Aspose' in t and 'Evaluation' in t for t in texts):
                excluded.append(dict(sheet=name, markers=texts))
            else:
                raise AssertionError(f'unattributed additional sheet {name!r}')
        ws = wb['Data']
        # write_only may omit <dimension>. Ignore declarations and scan the actual
        # entire sheet so incorrect declarations cannot hide appended cells.
        ws.reset_dimensions()
        data = [list(r) for r in ws.iter_rows(values_only=True)]
        if len(data) != fixture['rows'] or any(len(r) != fixture['cols'] for r in data):
            raise AssertionError('actual data dimensions != expected')
        h = digest(data)
        if h != fixture['digest']:
            raise AssertionError(f'typed/value digest mismatch: {h}')
        return dict(ok=True, digest=h, excluded_evaluation=excluded,
                    physical_bytes=Path(path).stat().st_size,
                    verify_ms=(time.perf_counter_ns()-start)/1e6)
    finally:
        wb.close()
