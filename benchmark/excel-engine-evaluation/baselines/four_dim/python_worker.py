import copy
import json
import os
import sys
import time
import traceback
from pathlib import Path
import openpyxl


def main():
    com='--com' in sys.argv;app=None;anchor=None
    try:
        if com:
            import win32com.client
            import win32process
            before=set(win32process.EnumProcesses())
            app=win32com.client.DispatchEx('Excel.Application')
            _,pid=win32process.GetWindowThreadProcessId(app.Hwnd)
            if pid in before:
                # Refuse to operate on or quit a pre-existing user's process.
                app=None;raise RuntimeError('Excel process was pre-existing; isolated ownership not proven')
            app.Visible=False;app.DisplayAlerts=False;app.EnableEvents=False
            app.AutomationSecurity=3;app.AskToUpdateLinks=False
            anchor=app.Workbooks.Add();app.Calculation=-4135
            ready=dict(ready=True,engine='excel-com',version=str(app.Version),owned_excel_pid=pid,pid=os.getpid())
        else: ready=dict(ready=True,engine='openpyxl',version=openpyxl.__version__,pid=os.getpid())
        print(json.dumps(ready),flush=True)
        for line in sys.stdin:
            try:
                q=json.loads(line);op=q['op']
                if op=='ping': result=dict(ok=True,processing_ms=0)
                elif com:
                    if op!='calc':raise RuntimeError('COM adapter is calculation-only')
                    start=time.perf_counter_ns()
                    wb=app.Workbooks.Open(str(Path(q['input']).resolve()),UpdateLinks=0,ReadOnly=True,AddToMru=False,IgnoreReadOnlyRecommended=True)
                    load_ms=(time.perf_counter_ns()-start)/1e6
                    try:
                        prior=[v for row in wb.Worksheets('Calc').Range('A1:H256').Value2 for v in row]
                        if any(v!=-9999 for v in prior):raise RuntimeError('COM input cache was not stale before timed calculation')
                        start=time.perf_counter_ns();app.CalculateFullRebuild();ms=(time.perf_counter_ns()-start)/1e6
                        values=[v for row in wb.Worksheets('Calc').Range('A1:H256').Value2 for v in row]
                    finally: wb.Close(SaveChanges=False)
                    result=dict(ok=True,processing_ms=ms,load_ms=load_ms,values=values)
                else:
                    output=Path(q['output'])
                    if output.exists():raise FileExistsError('refuse existing output')
                    start=time.perf_counter_ns();wb=openpyxl.load_workbook(q['input']);load_ms=(time.perf_counter_ns()-start)/1e6
                    try:
                        if op=='modify':
                            c=wb['Data']['J12'];c.value=9999;font=copy.copy(c.font);font.bold=True;c.font=font
                        elif op=='header':wb['Data'].oddHeader.left.text=wb['Data'].oddHeader.left.text.replace('benchmark','edited')
                        else:raise ValueError('unknown op')
                        wb.save(output)
                    finally:wb.close()
                    result=dict(ok=True,processing_ms=(time.perf_counter_ns()-start)/1e6,load_ms=load_ms)
            except Exception as exc:result=dict(ok=False,error=repr(exc),traceback=traceback.format_exc())
            print(json.dumps(result),flush=True)
    finally:
        if app is not None:
            if anchor is not None:anchor.Close(SaveChanges=False)
            app.Quit()


if __name__=='__main__':main()
