"""Independent selected-semantics + all-uncompressed-parts preservation audit."""
import copy
import datetime
import hashlib
import json
import posixpath
import sys
import xml.etree.ElementTree as ET
from decimal import Decimal,InvalidOperation
from pathlib import Path
import openpyxl

BASE=Path(__file__).resolve().parents[1]
sys.path.insert(0,str(BASE))
from capabilities import eval_view,view_diff
from compare.package_diff import package_diff,parts
M='http://schemas.openxmlformats.org/spreadsheetml/2006/main'
R='http://schemas.openxmlformats.org/officeDocument/2006/relationships'


def canonical(v):
    if isinstance(v,(datetime.datetime,datetime.date)):return v.isoformat()
    if isinstance(v,float) and v.is_integer():return int(v)
    if isinstance(v,str):return v.replace("'Data'!",'Data!')
    return v


def style(cell):
    f=cell.font;fill=cell.fill
    color=fill.fgColor
    return dict(font_name=f.name,font_size=f.sz,bold=bool(f.bold),italic=bool(f.italic),number_format=cell.number_format,
                fill_type=fill.patternType,fill_color=(color.type,str(color.value)[-6:] if color.type=='rgb' else color.value))


def xml_semantics(blob):
    def walk(n):return [n.tag,sorted(n.attrib.items()),n.text or '',n.tail or '',[walk(c) for c in n]]
    return walk(ET.fromstring(blob))


def custom_xml_stores(package):
    result=[]
    for rel in ET.fromstring(package['xl/_rels/workbook.xml.rels']):
        if rel.get('Type')!=R+'/customXml':continue
        t=rel.get('Target');part=t.lstrip('/') if t.startswith('/') else posixpath.normpath('xl/'+t)
        if part not in package:result.append(dict(broken_target=part));continue
        payload=xml_semantics(package[part]);props=None
        relpart=posixpath.dirname(part)+'/_rels/'+posixpath.basename(part)+'.rels'
        if relpart in package:
            for r in ET.fromstring(package[relpart]):
                if r.get('Type')!=R+'/customXmlProps':continue
                t=r.get('Target');prop=t.lstrip('/') if t.startswith('/') else posixpath.normpath(posixpath.dirname(part)+'/'+t)
                if prop not in package:props=dict(broken_target=prop);continue
                element=ET.fromstring(package[prop]);ns='http://schemas.openxmlformats.org/officeDocument/2006/customXml'
                props=dict(item_id=element.get('{'+ns+'}itemID','').upper(),schemas=sorted(x.get('{'+ns+'}uri','') for x in element.iter('{'+ns+'}schemaRef')))
        result.append(dict(payload=payload,properties=props))
    return sorted(result,key=lambda x:json.dumps(x,sort_keys=True))


def snapshot(path,normalize_target=None,original_sheets=('Data','Calc')):
    wb=openpyxl.load_workbook(path,data_only=False)
    out=dict(sheets={},names={n:canonical(v.attr_text) for n,v in wb.defined_names.items()})
    for name in original_sheets:
        if name not in wb:out['sheets'][name]=dict(missing=True);continue
        ws=wb[name];h=hashlib.sha256();comments={};links={};styles={}
        for row in ws:
            for cell in row:
                value=cell.value;kind=cell.data_type
                if normalize_target=='modify' and name=='Data' and cell.coordinate=='J12':value=12;kind='n'
                if value is not None:
                    h.update((json.dumps([cell.coordinate,kind,canonical(value)],ensure_ascii=False,separators=(',',':'))+'\n').encode())
                if cell.comment:comments[cell.coordinate]=(cell.comment.text,cell.comment.author)
                if cell.hyperlink:links[cell.coordinate]=(cell.hyperlink.target,cell.hyperlink.location)
                if cell.has_style:
                    s=style(cell)
                    if normalize_target=='modify' and name=='Data' and cell.coordinate=='J12':s['bold']=False
                    styles[cell.coordinate]=s
        validations=[]
        for d in ws.data_validations.dataValidation:
            validations.append([str(d.sqref),d.type,d.operator or 'between',d.formula1,d.formula2])
        tables=[]
        for t in ws.tables.values():tables.append([t.name,t.ref,[c.name for c in t.tableColumns]])
        header=ws.oddHeader.left.text
        if normalize_target=='header' and name=='Data' and header is not None:header=header.replace('edited','benchmark')
        out['sheets'][name]=dict(business_digest=h.hexdigest(),merges=sorted(str(m) for m in ws.merged_cells.ranges),
              comments=comments,hyperlinks=links,styles=styles,validations=validations,tables=tables,
              chart_count=len(ws._charts),print_area=canonical(str(ws.print_area)),print_titles=ws.print_title_rows,
              header=header,footer=ws.oddFooter.right.text,orientation=ws.page_setup.orientation,paper=ws.page_setup.paperSize,
              freeze_panes=ws.freeze_panes,
              row_heights={str(k):v.height for k,v in ws.row_dimensions.items() if v.height is not None},
              col_widths={str(k):v.width for k,v in ws.column_dimensions.items()})
    package=parts(path);book=ET.fromstring(package['xl/workbook.xml']);rels={r.get('Id'):r.get('Target') for r in ET.fromstring(package['xl/_rels/workbook.xml.rels'])};caches={}
    for sheet in book.find(f'{{{M}}}sheets'):
        name=sheet.get('name')
        if name not in original_sheets:continue
        target=rels[sheet.get(f'{{{R}}}id')];part=target.lstrip('/') if target.startswith('/') else posixpath.normpath('xl/'+target)
        for cell in ET.fromstring(package[part]).iter(f'{{{M}}}c'):
            if cell.find(f'{{{M}}}f') is None:continue
            value=cell.find(f'{{{M}}}v');v=value.text if value is not None else None;typ=cell.get('t','n')
            if typ=='n' and v is not None:
                try:v=str(Decimal(v).normalize())
                except InvalidOperation:pass
            caches[name+'!'+cell.get('r')]=(typ,v)
    out['caches']=caches;out['custom_xml']=custom_xml_stores(package);wb.close();return out


def compare(source,actual):
    issues=[]
    if source['names']!=actual['names']:issues.append(dict(kind='defined_names',before=source['names'],after=actual['names']))
    if source['custom_xml']!=actual['custom_xml']:issues.append(dict(kind='workbook_custom_xml_changed_or_lost',before=source['custom_xml'],after=actual['custom_xml']))
    for name,s in source['sheets'].items():
        a=actual['sheets'].get(name,{})
        if a.get('missing'):issues.append(dict(sheet=name,kind='sheet_missing'));continue
        for key in ('business_digest','merges','comments','hyperlinks','validations','tables','chart_count','print_area','print_titles','header','footer','orientation','paper','freeze_panes'):
            if s[key]!=a.get(key):issues.append(dict(sheet=name,kind=key,before=s[key],after=a.get(key)))
        # Original explicit formatting must survive. Added default dimension
        # declarations are tracked as representations, not automatically loss.
        for key in ('styles','row_heights','col_widths'):
            mismatches=[k for k,v in s[key].items() if a.get(key,{}).get(k)!=v]
            if mismatches:issues.append(dict(sheet=name,kind=key,mismatch_count=len(mismatches),examples=mismatches[:10]))
    bad=[k for k,v in source['caches'].items() if actual['caches'].get(k)!=v]
    if bad:issues.append(dict(kind='formula_cache_changed_or_lost',count=len(bad),examples=[dict(cell=k,before=source['caches'][k],after=actual['caches'].get(k)) for k in bad[:8]]))
    return issues


def fingerprint(package):
    # Selected semantics do not inspect document timestamps. Every part is
    # still independently hashed/reported by package_audit below.
    return tuple((n,hashlib.sha256(b).hexdigest()) for n,b in sorted(package.items()) if n!='docProps/core.xml')


def package_audit(a,b,canonical_cache):
    def canonical_hash(blob):
        h=hashlib.sha256(blob).hexdigest()
        if h not in canonical_cache:
            try:canonical_cache[h]=hashlib.sha256(ET.canonicalize(blob).encode()).hexdigest()
            except (ET.ParseError,UnicodeError):canonical_cache[h]=None
        return canonical_cache[h]
    changed=[]
    for n in sorted(a.keys() & b.keys()):
        if a[n]==b[n]:continue
        item=dict(part=n,source_sha256=hashlib.sha256(a[n]).hexdigest(),target_sha256=hashlib.sha256(b[n]).hexdigest())
        if n.endswith(('.xml','.rels')):
            ah,bh=canonical_hash(a[n]),canonical_hash(b[n]);item['canonical_xml_equal']=ah is not None and ah==bh
        changed.append(item)
    return dict(added=sorted(b.keys()-a.keys()),removed=sorted(a.keys()-b.keys()),changed=changed,
                unchanged=sorted(n for n in a.keys() & b.keys() if a[n]==b[n]))


def audit(root,corrected=False,source_override=None):
    root=Path(root);evidence=root/('modify-corrected' if corrected else 'run')
    if source_override:source=Path(source_override)
    else:
        manifest=json.loads((root/'fixtures/manifest.json').read_text('utf-8'))
        source=(root/'fixtures/modify-complex-workbook-customxml.xlsx') if corrected else Path(manifest['modify'])
    expected=snapshot(source);source_parts=parts(source);has_workbook_custom_xml=bool(custom_xml_stores(source_parts))
    records=[json.loads(s) for s in (evidence/'rounds.jsonl').read_text('utf-8').splitlines()]
    all_results=[];semantic_cache={};canonical_cache={};previous={}
    if (evidence/'structure.jsonl').exists():
        previous={x['output']:x for x in [json.loads(s) for s in (evidence/'structure.jsonl').read_text('utf-8').splitlines()]}
    base_view=eval_view(source,source)
    for r in records:
        if r['op'] not in ('modify','header') or not r.get('ok'):continue
        path=Path(r['output']);p=parts(path);fp=(r['op'],fingerprint(p))
        if str(path) in previous:
            old=previous[str(path)];semantic_cache[fp]=old['issues'];all_results.append(old);continue
        reused=fp in semantic_cache
        issues=copy.deepcopy(semantic_cache[fp]) if reused else compare(expected,snapshot(path,r['op']))
        opaque=[n for n in source_parts if n.startswith('customXml/') and p.get(n)!=source_parts[n]]
        if opaque and not reused and not has_workbook_custom_xml:issues.append(dict(kind='opaque_parts_lost_or_changed',parts=opaque))
        semantic_cache[fp]=copy.deepcopy(issues)
        raw=package_audit(source_parts,p,canonical_cache)
        item=dict(condition=r['condition'],engine=r['engine'],round=r['round'],phase=r['phase'],output=str(path),
                  selected_semantics_preserved=not issues,issues=issues,package=raw)
        item['semantic_hash_cache_reused']=reused
        if r['engine']=='aspose':
            view=eval_view(source,path);item['evaluation_view']=view;item['filtered_package_diff']=view_diff(base_view,view)
        if r['op']=='header' and r['engine']=='rust':
            item['precise_patch']=not raw['added'] and not raw['removed'] and [x['part'] for x in raw['changed']]==['xl/worksheets/sheet1.xml']
        all_results.append(item)
        with (evidence/'structure.jsonl').open('a',encoding='utf-8') as f:f.write(json.dumps(item)+'\n')
        print('STRUCTURE',r['condition'],r['engine'],r['round'],'issues',[(x['kind'],x.get('count')) for x in issues],flush=True)
    (evidence/'structure-summary.json').write_text(json.dumps([dict(condition=condition,engine=engine,outputs=len(rows),
            preserved=sum(x['selected_semantics_preserved'] for x in rows),issue_kinds=sorted({i['kind'] for x in rows for i in x['issues']}),
            precise_patch=sum(x.get('precise_patch',False) for x in rows))
          for (condition,engine),rows in group(all_results).items()],indent=2),'utf-8')


def group(rows):
    groups={}
    for r in rows:groups.setdefault((r['condition'],r['engine']),[]).append(r)
    return groups


if __name__=='__main__':audit(sys.argv[1],'--corrected' in sys.argv,sys.argv[sys.argv.index('--source')+1] if '--source' in sys.argv else None)
