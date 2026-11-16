"""Protocol/failure injection only; never substitutes for real .NET acceptance."""
import json,os,sys,time,zipfile
from pathlib import Path

mode, record = sys.argv[1:]
def emit(v): print(json.dumps(v),flush=True)
if mode=='startup_timeout': time.sleep(60)
emit(dict(ready=True,protocol_version=1,engine='Aspose.Cells',pid=os.getpid()))
if mode=='write_timeout': time.sleep(60)
for line in sys.stdin:
 q=json.loads(line)
 with open(record,'a',newline='\n') as f:f.write(q['method']+'\n')
 if q['method']=='Ping':
  emit(dict(ok=True,request_id=q['request_id'],pid=os.getpid()));continue
 if mode=='request_timeout':time.sleep(60)
 if mode=='source_changed':Path(q['input']).write_bytes(b'changed externally by failure-injection stub')
 if mode=='crash':
  Path(q['candidate']).write_bytes(b'partial candidate');sys.exit(2)
 with zipfile.ZipFile(q['candidate'],'x') as z:
  z.writestr('[Content_Types].xml','<Types/>')
  z.writestr('_rels/.rels','<Relationships/>')
  z.writestr('xl/_rels/workbook.xml.rels','<Relationships><Relationship Id="rId1" Target="worksheets/sheet1.xml"/></Relationships>')
  z.writestr('xl/workbook.xml','<workbook xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><bookViews><workbookView activeTab="0"/></bookViews><sheets><sheet name="Data" r:id="rId1"/></sheets></workbook>')
  z.writestr('xl/worksheets/sheet1.xml','<worksheet/>' if mode!='bad_candidate' else 'invalid xml')
 emit(dict(ok=True,request_id='wrong' if mode=='wrong_id' else q['request_id'],candidate=q['candidate'],evaluation_output=True,
           source_sheets=['Data'],output_sheets=['Data'],source_active_index=0,output_active_index=0,
           processing_ms=1,load_ms=0,transform_ms=1,save_ms=0,pid=os.getpid(),attempts=99 if mode=='drain' else 1,draining=mode=='drain'))
 if mode=='drain':break
