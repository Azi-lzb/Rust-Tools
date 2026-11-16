"""Reuse the benchmark's independent complex-structure oracle for this worker."""
import hashlib
import json
import sys
from pathlib import Path

def main():
    root=Path(sys.argv[1]).resolve();root.mkdir(exist_ok=False)
    repo=Path(__file__).resolve().parents[2]
    sys.path.insert(0,str(repo/'benchmark/excel-engine-evaluation/baselines/four_dim'))
    from run import Worker
    from structure import snapshot,compare,parts,package_audit
    from chart_audit import projection,resolved_bindings
    source=repo/'benchmark/excel-engine-evaluation/baselines/results/four-dim-20261114/fixtures/modify-complex-workbook-customxml.xlsx'
    before=hashlib.sha256(source.read_bytes()).hexdigest()
    target=root/'candidate.xlsx'
    worker=Worker([sys.argv[2],sys.argv[3],'--stdio'],root/'worker.stderr.jsonl')
    try:
        response=worker.request(dict(protocol_version=1,request_id='complex',run_id='complex-xlsx-acceptance',method='Modify',input=str(source),candidate=str(target),operations=[
            dict(op='set_value',sheet='Data',cell='J12',value=dict(type='number',value=9999)),
            dict(op='set_style',sheet='Data',cell='J12',style=dict(bold=True))]))
        assert response['ok'] and response['evaluation_output'],response
    finally:worker.close()
    original=snapshot(source);actual=snapshot(target,normalize_target='modify')
    issues=compare(original,actual)
    a,b=parts(source),parts(target)
    chart_ok=resolved_bindings(projection(a))==resolved_bindings(projection(b))
    report=dict(status='PASS_SELECTED_SEMANTICS' if not issues and chart_ok else 'FAIL',issues=issues,
                formula_caches=len(original['caches']),chart_bindings_preserved=chart_ok,
                workbook_custom_xml_preserved=original['custom_xml']==actual['custom_xml'],
                source_unchanged=hashlib.sha256(source.read_bytes()).hexdigest()==before,
                evaluation_output_accepted=True,response=response,package_diff=package_audit(a,b,{}),
                precise_ooxml_patch=False,scope='Selected semantics of this synthetic fixture only')
    (root/'RESULT.json').write_text(json.dumps(report,indent=2),'utf-8')
    assert not issues and chart_ok and report['source_unchanged'],issues
    print(json.dumps({k:v for k,v in report.items() if k not in ('package_diff','response')}))

if __name__=='__main__':main()
