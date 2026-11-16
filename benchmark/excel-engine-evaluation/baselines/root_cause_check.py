"""Run identical direct API reproductions before/after local dependency patches."""
import argparse
import hashlib
import json
import subprocess
import xml.etree.ElementTree as ET
import zipfile
from pathlib import Path
from root_cause_fixtures import biff


def norm(f):
    # Only remove redundant quotes around a known simple sheet identifier.
    return f.lstrip('=').replace("'Data'!",'Data!').replace("'Calc'!",'Calc!')


def validations(path):
    ns={'m':'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
    with zipfile.ZipFile(path) as z: xml=ET.fromstring(z.read('xl/worksheets/sheet1.xml'))
    return [dict(sqref=x.get('sqref'),formula1=x.findtext('m:formula1',None,ns),formula2=x.findtext('m:formula2',None,ns),
                 type=x.get('type'),operator=x.get('operator'),prompt=x.get('prompt'),promptTitle=x.get('promptTitle'))
            for x in xml.findall('m:dataValidations/m:dataValidation',ns)]


def main(a):
    root=Path(a.root).resolve(); spec=json.loads((root/'spec.json').read_text('utf-8'))
    expected=spec['formulas']; inputs=['xlwt-references.xls','aspose-references.xls','openpyxl-control.xlsx']
    independent=[json.loads(s) for s in (root/'aspose-independent.jsonl').read_text('utf-8-sig').splitlines()]
    sdk_diff=[]
    for r in independent:
        if 'formulas' in r:
            diff=[dict(row=i+1,expected=f,actual=r['formulas'][i]['formula']) for i,f in enumerate(expected) if norm(f)!=norm(r['formulas'][i]['formula'])]
            sdk_diff.append(dict(input=r['input'],differences=diff))
    assert not any(r['differences'] for r in sdk_diff),sdk_diff
    output=root/f'{a.label}-validation.xlsx'
    requests=[dict(op='formulas',input=str(root/name)) for name in inputs]
    requests.append(dict(op='validation',input=str(root/'validation-minimal.xlsx'),output=str(output)))
    req=root/f'{a.label}-requests.json';req.write_text(json.dumps(requests,indent=2),'utf-8')
    p=subprocess.run([a.rust,str(req)],text=True,capture_output=True,encoding='utf-8')
    (root/f'{a.label}.stdout.jsonl').write_text(p.stdout,'utf-8');(root/f'{a.label}.stderr.txt').write_text(p.stderr,'utf-8')
    assert p.returncode==0,p.stderr
    responses=[json.loads(s) for s in p.stdout.splitlines()]; assert len(responses)==4 and all(r['ok'] for r in responses)
    report=dict(label=a.label,independent_reader=sdk_diff,formulas=[],validation={})
    for name,r in zip(inputs,responses[:3]):
        assert len(r['formulas'])==len(expected)
        diff=[dict(row=i+1,expected=f,actual=r['formulas'][i]['formula']) for i,f in enumerate(expected) if norm(f)!=norm(r['formulas'][i]['formula'])]
        report['formulas'].append(dict(input=name,cases=len(expected),differences=diff))
    raw_before=validations(root/'validation-minimal.xlsx'); raw_after=validations(output)
    report['validation']=dict(in_memory_before_save=responses[-1]['in_memory_before_save'],after_reload=responses[-1]['after_reload'],
                              raw_xml_before=raw_before,raw_xml_after=raw_after,raw_xml_equal=raw_before==raw_after)
    for raw,mem in zip(raw_before,responses[-1]['in_memory_before_save']):
        assert (raw['formula1'] or '')==mem['formula1'] and (raw['formula2'] or '')==mem['formula2']
    report['input_hashes']={name:hashlib.sha256((root/name).read_bytes()).hexdigest() for name in inputs+['validation-minimal.xlsx']}
    (root/f'{a.label}-report.json').write_text(json.dumps(report,indent=2),'utf-8')
    (root/'aspose-biff-records.json').write_text(json.dumps(biff(root/'aspose-references.xls'),indent=2),'utf-8')
    print(json.dumps(dict(label=a.label,formula_mismatch=[(r['input'],len(r['differences'])) for r in report['formulas']],validation_preserved=raw_before==raw_after),indent=2))
    if a.expect_fixed:
        assert all(not r['differences'] for r in report['formulas'])
        assert raw_before==raw_after


if __name__=='__main__':
    p=argparse.ArgumentParser()
    for n in ('root','rust','label'):p.add_argument('--'+n,required=True)
    p.add_argument('--expect-fixed',action='store_true'); main(p.parse_args())
