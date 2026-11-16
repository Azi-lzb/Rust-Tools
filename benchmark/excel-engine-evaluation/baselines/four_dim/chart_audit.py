"""Supplement chart counts with original chart type/title/reference checks."""
import json
import re
import sys
import xml.etree.ElementTree as ET
from pathlib import Path
from structure import parts,canonical


def projection(package):
    result=[]
    for name,blob in package.items():
        if not(name.startswith('xl/charts/chart') and name.endswith('.xml')):continue
        root=ET.fromstring(blob)
        local=lambda n:n.tag.rsplit('}',1)[-1]
        result.append(dict(plot_types=sorted(local(n) for n in root.iter() if local(n).endswith('Chart')),
                           series_count=sum(local(n)=='ser' for n in root.iter()),
                           texts=[n.text for n in root.iter() if local(n)=='t'],
                           formulas=sorted(canonical(n.text or '') for n in root.iter() if local(n)=='f')))
    return sorted(result,key=lambda x:json.dumps(x,sort_keys=True))


def resolved_bindings(projection):
    # Chart references have no formula-cell origin. Keep raw spellings in the
    # report; for this known Data-sheet fixture compare their resolved targets.
    result=[]
    for chart in projection:
        c=dict(chart);formulas=[]
        for f in c['formulas']:
            m=re.fullmatch(r'(Data)!(\$?[A-Z]+\$?\d+(?::\$?[A-Z]+\$?\d+)?)',f)
            formulas.append(m[1]+'!'+m[2].replace('$','') if m else f)
        c['formulas']=sorted(formulas);result.append(c)
    return result


if __name__=='__main__':
    root=Path(sys.argv[1]);explicit='--source' in sys.argv
    source_path=Path(sys.argv[sys.argv.index('--source')+1]) if explicit else root/'fixtures/modify-complex-workbook-customxml.xlsx'
    folder=root/('run' if explicit else 'modify-corrected');source=projection(parts(source_path))
    records=[json.loads(s) for s in (folder/'rounds.jsonl').read_text('utf-8').splitlines() if json.loads(s)['op'] in ('modify','header')]
    rows=[dict(output=r['output'],condition=r['condition'],engine=r['engine'],round=r['round'],
               match=resolved_bindings(p:=projection(parts(r['output'])))==resolved_bindings(source),
               raw_spelling_equal=p==source,actual=p) for r in records]
    (folder/'chart-audit.json').write_text(json.dumps(dict(expected=source,outputs=rows),indent=2),'utf-8')
    print('chart projection matches',sum(r['match'] for r in rows),'/',len(rows))
