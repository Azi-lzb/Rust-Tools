"""Small synthetic formula and edit experiments, separate from speed measurements."""
import argparse
import copy
import hashlib
import json
import struct
import subprocess
import xml.etree.ElementTree as ET
import zipfile
from pathlib import Path
import openpyxl
import xlrd
import xlwt
from openpyxl.comments import Comment
from openpyxl.styles import Font, PatternFill
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.worksheet.datavalidation import DataValidation
from compare.package_diff import package_diff, parts

MAIN='http://schemas.openxmlformats.org/spreadsheetml/2006/main'
REL='http://schemas.openxmlformats.org/officeDocument/2006/relationships'
PKG='http://schemas.openxmlformats.org/package/2006/relationships'
CT='http://schemas.openxmlformats.org/package/2006/content-types'


def make_edit(path):
    wb=openpyxl.Workbook(); ws=wb.active; ws.title='Data'
    ws.append(['key','quantity','price','amount'])
    for r in range(2,17): ws.append([f'item{r}',r,r*3,f'=B{r}*C{r}'])
    ws['B12'].number_format='0.00'; ws['B12'].font=Font(name='Calibri',size=11,bold=False)
    ws['A2'].comment=Comment('synthetic annotation','probe')
    ws['A3'].hyperlink='https://example.invalid/synthetic'
    ws['A1'].fill=PatternFill('solid',fgColor='FFFF00')
    ws.merge_cells('F1:G1'); ws['F1']='merge sentinel'
    ws.row_dimensions[1].height=28; ws.column_dimensions['A'].width=21
    ws.print_area='A1:D16'; ws.print_title_rows='1:1'
    ws.page_setup.orientation='landscape'; ws.page_setup.paperSize='9'
    ws.oddHeader.left.text='synthetic &Bheader'; ws.oddFooter.right.text='&P / &N'
    dv=DataValidation(type='whole',operator='between',formula1=0,formula2=1000)
    ws.add_data_validation(dv); dv.add('B2:B16')
    calc=wb.create_sheet('Calc'); calc['A1']='=SUM(Data!B2:B16)'
    wb.defined_names.add(DefinedName('Quantity',attr_text="'Data'!$B$2:$B$16"))
    wb.save(path)
    # A package-level unknown OPC part tests byte retention. This is NOT a
    # WorkbookPart-associated custom XML store; do not infer loss of supported
    # Workbook custom XML from this older fixture. See four_dim's corrected case.
    package=parts(path)
    root=ET.fromstring(package['_rels/.rels'])
    ET.SubElement(root,f'{{{PKG}}}Relationship',Id='probeCustomXml',Type=REL+'/customXml',Target='customXml/item1.xml')
    package['_rels/.rels']=ET.tostring(root,encoding='utf-8',xml_declaration=True)
    package['customXml/item1.xml']=b'<probe xmlns="urn:rusttools:synthetic"><keep>opaque-sentinel</keep></probe>'
    with zipfile.ZipFile(path,'w',zipfile.ZIP_DEFLATED) as z:
        for n,b in package.items(): z.writestr(n,b)


def make_formula(path):
    wb=xlwt.Workbook(); ws=wb.add_sheet('Data')
    for r,row in enumerate([['key','quantity','price'],['alpha',2,3],['beta',4,5],['alpha',6,7]]):
        for c,v in enumerate(row): ws.write(r,c,v)
    calc=wb.add_sheet('Calc')
    formulas=['SUMIF(Data!A2:A4,"alpha",Data!B2:B4)',
              'SUMPRODUCT(Data!B2:B4,Data!C2:C4)',
              'IF(AND(Data!B2>0,Data!C2>0),ROUND(Data!B2*Data!C2,2),0)',
              'INDEX(Data!C2:C4,MATCH("beta",Data!A2:A4,0))']
    for r,f in enumerate(formulas): calc.write(r,0,xlwt.Formula(f))
    wb.save(str(path))
    # Inspect actual BIFF cached-result bytes, not just an empty value from a reader.
    stream=xlrd.compdoc.CompDoc(path.read_bytes()).get_named_stream('Workbook')
    offset=0; caches=[]
    while offset+4<=len(stream):
        code,length=struct.unpack_from('<HH',stream,offset); payload=stream[offset+4:offset+4+length]
        if code==0x0006:
            raw=payload[6:14]
            caches.append(dict(row=struct.unpack_from('<H',payload)[0]+1,column=struct.unpack_from('<H',payload,2)[0]+1,
                               cached_hex=raw.hex(),empty_cache=(raw==bytes.fromhex('030000000000ffff'))))
        offset+=4+length
    assert len(caches)==4 and all(c['empty_cache'] for c in caches),caches
    return dict(formulas=formulas,expected=[8,68,6,5],biff_formula_caches=caches)


def snapshot(path, originals=('Data','Calc')):
    wb=openpyxl.load_workbook(path,data_only=False)
    out={}
    for name in originals:
        if name not in wb: out[name]=dict(missing=True); continue
        ws=wb[name]; cells={}
        for row in ws:
            for c in row:
                if c.value is not None:
                    cells[c.coordinate]=dict(value=c.value,type=c.data_type,number_format=c.number_format,
                         bold=bool(c.font.bold),font_name=c.font.name,font_size=c.font.sz,
                         fill_type=c.fill.patternType,fill_color_type=c.fill.fgColor.type,
                         fill_color=c.fill.fgColor.value,comment=c.comment.text if c.comment else None,
                         hyperlink=c.hyperlink.target if c.hyperlink else None)
        out[name]=dict(cells=cells,merges=sorted(str(m) for m in ws.merged_cells.ranges),
                      row_heights={str(k):v.height for k,v in ws.row_dimensions.items() if v.height is not None},
                      col_widths={str(k):v.width for k,v in ws.column_dimensions.items()},
                      print_area=str(ws.print_area),print_title_rows=ws.print_title_rows,
                      orientation=ws.page_setup.orientation,paper=ws.page_setup.paperSize,
                      header=ws.oddHeader.left.text,footer=ws.oddFooter.right.text,
                      validations=[(str(d.sqref),d.type,d.operator,d.formula1,d.formula2) for d in ws.data_validations.dataValidation])
    out['defined_names']={n:v.attr_text for n,v in wb.defined_names.items()}
    package=parts(path)
    book=ET.fromstring(package['xl/workbook.xml'])
    rels={n.get('Id'):n.get('Target') for n in ET.fromstring(package['xl/_rels/workbook.xml.rels'])}
    caches={}
    for sheet in book.find(f'{{{MAIN}}}sheets'):
        name=sheet.get('name')
        if name not in originals: continue
        t=rels[sheet.get(f'{{{REL}}}id')]
        part=t.lstrip('/') if t.startswith('/') else 'xl/'+t
        for cell in ET.fromstring(package[part]).iter(f'{{{MAIN}}}c'):
            formula=cell.find(f'{{{MAIN}}}f'); value=cell.find(f'{{{MAIN}}}v')
            if formula is not None:
                caches[f'{name}!{cell.get("r")}']=dict(formula=formula.text,type=cell.get('t','n'),
                    cache_value=value.text if value is not None else None,cache_element=value is not None)
    out['formula_caches']=caches
    wb.close()
    return out


def differences(a,b,prefix=''):
    if isinstance(a,dict) and isinstance(b,dict):
        result=[]
        for k in sorted(a.keys()|b.keys()):
            key=f'{prefix}/{k}'
            if k not in a or k not in b: result.append(dict(path=key,before=a.get(k),after=b.get(k)))
            else: result.extend(differences(a[k],b[k],key))
        return result
    return [] if a==b else [dict(path=prefix,before=a,after=b)]


def classify(change):
    p,a,b=change['path'],change['before'],change['after']
    if '/defined_names/' in p and isinstance(a,str) and isinstance(b,str) and a.replace("'Data'!",'Data!')==b.replace("'Data'!",'Data!'):
        return 'equivalent_sheet_name_quoting'
    if '/row_heights/' in p and a is None and b==15:
        return 'explicit_default_row_height'
    if p.endswith('/fill_color') and isinstance(a,str) and isinstance(b,str) and a[-6:]==b[-6:]:
        return 'color_alpha_encoding_change'
    if p.startswith('/formula_caches/') and p.endswith('/type'):
        return 'empty_formula_cache_type_encoding_change'
    if p.endswith('/validations'):
        # OOXML operator defaults to between. Missing bounds are a different,
        # material issue and must never be normalized away.
        if all(v[3] is not None and v[4] is not None for v in a) and any(v[3] is None or v[4] is None for v in b):
            return 'DATA_VALIDATION_BOUNDS_LOST'
        if [list(v[:2])+[v[2] or 'between']+list(v[3:]) for v in a]==[list(v[:2])+[v[2] or 'between']+list(v[3:]) for v in b]:
            return 'explicit_default_validation_operator'
    if '/col_widths/' in p:
        return 'column_dimension_change_review'
    return 'UNEXPECTED_REQUIRES_REVIEW'


def view_diff(a,b):
    a=a['remaining_part_sha256']; b=b['remaining_part_sha256']
    return dict(added=sorted(b.keys()-a.keys()),removed=sorted(a.keys()-b.keys()),
                changed=sorted(n for n in a.keys() & b.keys() if a[n]!=b[n]))


def eval_view(source, target):
    """In-memory view only; remove exact added vendor sheet and its graph edges.

    Other parts (including shared styles and strings) remain visible. Their broad
    rewrite cannot be assumed to be only watermark noise.
    """
    source_wb=openpyxl.load_workbook(source,read_only=True)
    originals=source_wb.sheetnames; source_wb.close()
    wb=openpyxl.load_workbook(target,read_only=True)
    excluded=[]
    for name in set(wb.sheetnames)-set(originals):
        texts=[v for row in wb[name].iter_rows(max_row=10,max_col=5,values_only=True) for v in row if isinstance(v,str)]
        if name=='Evaluation Warning' and any('Aspose' in t and 'Evaluation' in t for t in texts): excluded.append(name)
        else: raise AssertionError(f'unattributed added sheet {name}')
    wb.close()
    p=parts(target); root=ET.fromstring(p['xl/workbook.xml']); rels=ET.fromstring(p['xl/_rels/workbook.xml.rels']); ct=ET.fromstring(p['[Content_Types].xml'])
    sheets=root.find(f'{{{MAIN}}}sheets'); removed=[]; eval_string_ids=set()
    for node in list(sheets):
        if node.get('name') not in excluded: continue
        rid=node.get(f'{{{REL}}}id'); sheets.remove(node)
        for rel in list(rels):
            if rel.get('Id')==rid:
                t=rel.get('Target'); part=t.lstrip('/') if t.startswith('/') else 'xl/'+t
                eval_xml=ET.fromstring(p[part])
                for cell in eval_xml.iter(f'{{{MAIN}}}c'):
                    v=cell.find(f'{{{MAIN}}}v')
                    if cell.get('t')=='s' and v is not None: eval_string_ids.add(int(v.text))
                removed.append(part); p.pop(part,None); rels.remove(rel)
                for override in list(ct):
                    if override.get('PartName')=='/'+part: ct.remove(override)
    for name,xml in [('xl/workbook.xml',root),('xl/_rels/workbook.xml.rels',rels),('[Content_Types].xml',ct)]:
        p[name]=ET.tostring(xml,encoding='utf-8')
    removed_strings=[]
    if eval_string_ids and 'xl/sharedStrings.xml' in p:
        # Remove only strings referenced exclusively by the confirmed added
        # evaluation sheet. Strings also used by user sheets always survive.
        worksheets={n:ET.fromstring(b) for n,b in p.items() if n.startswith('xl/worksheets/') and n.endswith('.xml')}
        used=set(); refs=[]
        for xml in worksheets.values():
            for cell in xml.iter(f'{{{MAIN}}}c'):
                v=cell.find(f'{{{MAIN}}}v')
                if cell.get('t')=='s' and v is not None: used.add(int(v.text)); refs.append(v)
        drop=eval_string_ids-used
        sst=ET.fromstring(p['xl/sharedStrings.xml']); entries=list(sst); mapping={}
        for old,node in enumerate(entries):
            if old in drop:
                removed_strings.append(dict(original_index=old,text=''.join(node.itertext())))
                sst.remove(node)
            else: mapping[old]=len(mapping)
        for v in refs: v.text=str(mapping[int(v.text)])
        sst.set('uniqueCount',str(len(mapping))); sst.set('count',str(len(refs)))
        p['xl/sharedStrings.xml']=ET.tostring(sst,encoding='utf-8')
        for n,xml in worksheets.items(): p[n]=ET.tostring(xml,encoding='utf-8')
    return dict(excluded_sheets=excluded,excluded_parts=removed,
                excluded_exclusive_evaluation_strings=removed_strings,
                remaining_part_sha256={n:hashlib.sha256(b).hexdigest() for n,b in sorted(p.items())},
                note='comparison view only; exclusive evaluation strings removed with reference remapping; remaining shared styles, user strings and active-sheet changes are not blanket-filtered')


def execute(cmd, requests, output, name):
    request_file=output/f'{name}-requests.json'; request_file.write_text(json.dumps(requests,indent=2),'utf-8')
    p=subprocess.run(cmd+[str(request_file)],text=True,encoding='utf-8',capture_output=True)
    (output/f'{name}.stdout.jsonl').write_text(p.stdout,'utf-8'); (output/f'{name}.stderr.txt').write_text(p.stderr,'utf-8')
    if p.returncode: raise RuntimeError(f'{name} failed: {p.stderr} {p.stdout}')
    results=[json.loads(line) for line in p.stdout.splitlines()]
    assert len(results)==len(requests) and all(x['ok'] for x in results),results
    return results


def run(args):
    output=Path(args.output).resolve(); output.mkdir(parents=True,exist_ok=False)
    source=output/'edit-input.xlsx'; make_edit(source)
    xls=output/'formula-missing-cache.xls'; formula=make_formula(xls)
    hashes={str(p):hashlib.sha256(p.read_bytes()).hexdigest() for p in (source,xls)}
    paths={engine:{op:output/f'{engine}-{op}.xlsx' for op in ('noop','value','style')} for engine in ('openpyxl','umya','aspose')}
    paths['openpyxl']['header']=output/'openpyxl-header.xlsx'
    paths['aspose']['header']=output/'aspose-header.xlsx'
    for op,target in paths['openpyxl'].items():
        wb=openpyxl.load_workbook(source); c=wb['Data']['B12']
        if op=='value': c.value=999
        elif op=='style': font=copy.copy(c.font); font.bold=True; c.font=font
        elif op=='header': wb['Data'].oddHeader.left.text=wb['Data'].oddHeader.left.text.replace('synthetic','edited')
        wb.save(target); wb.close()
    rust_requests=[dict(op=op,input=str(source),output=str(path)) for op,path in paths['umya'].items()]
    rust_formula=output/'rust-degraded.xlsx'; rust_requests.append(dict(op='formula',input=str(xls),output=str(rust_formula)))
    rust=execute([args.rust],rust_requests,output,'rust')
    assert rust[-1]['status']=='UNSUPPORTED' and not rust_formula.exists(),rust[-1]
    header=output/'ooxml-header.xlsx'
    execute([args.rust],[dict(op='header',input=str(source),output=str(header))],output,'ooxml')
    paths['ooxml-patch']={'header':header}
    aspose_requests=[dict(op=op,input=str(source),output=str(path)) for op,path in paths['aspose'].items()]
    calculated=output/'aspose-calculated.xlsx'; aspose_requests.append(dict(op='formula',input=str(xls),output=str(calculated)))
    aspose=execute([args.aspose],aspose_requests,output,'aspose')
    wb=openpyxl.load_workbook(calculated,data_only=True)
    actual=[wb['Calc'].cell(r,1).value for r in range(1,5)]; wb.close()
    assert actual==formula['expected'],actual
    fwb=openpyxl.load_workbook(calculated,data_only=False)
    assert all(fwb['Calc'].cell(r,1).data_type=='f' for r in range(1,5)); fwb.close()
    fwb=openpyxl.load_workbook(calculated,data_only=False)
    preserved=[fwb['Calc'].cell(r,1).value.lstrip('=').replace("'Data'!",'Data!') for r in range(1,5)]
    fwb.close()
    assert preserved==formula['formulas'],preserved
    formula.update(aspose_actual=actual,rust=rust[-1],classification='EXPECTED_CAPABILITY_GAP',
                   aspose_invocation='explicit CalculateFormula; opening/saving alone is not assumed to calculate')
    observed=[c['formula'] for c in rust[-1]['cells']]
    formula['formula_text_diagnostic']=dict(expected=formula['formulas'],calamine_observed=observed,
                status='UNEXPECTED_PARSER_DIFFERENCE' if observed!=formula['formulas'] else 'MATCH',
                action='XLS formula conversion remains intentionally unsupported; parsing diagnostics do not enable calculation or production routes')
    report=dict(formula=formula,edits=[])
    before=snapshot(source)
    for engine,outputs in paths.items():
        for op,path in outputs.items():
            expected=copy.deepcopy(before)
            if op=='value': expected['Data']['cells']['B12']['value']=999
            if op=='style': expected['Data']['cells']['B12']['bold']=True
            if op=='header': expected['Data']['header']=expected['Data']['header'].replace('synthetic','edited')
            diff=differences(expected,snapshot(path))
            for change in diff: change['classification']=classify(change)
            raw=package_diff(source,path)
            allowed={'xl/worksheets/sheet1.xml'} if op in ('value','header') else {'xl/worksheets/sheet1.xml','xl/styles.xml'} if op=='style' else set()
            unrelated=set(x['part'] for x in raw['changed'])-allowed
            item=dict(engine=engine,operation=op,semantic_unexpected=diff,raw_package=raw,
                      non_target_changed_parts=sorted(unrelated),
                      precise_part_patch=not(raw['added'] or raw['removed'] or unrelated))
            if engine=='aspose':
                item['evaluation_view']=eval_view(source,path)
                item['filtered_package_diff']=view_diff(eval_view(source,source),item['evaluation_view'])
                if op!='noop': item['versus_own_noop']=package_diff(outputs['noop'],path)
                if op!='noop': item['versus_own_noop_filtered']=view_diff(eval_view(source,outputs['noop']),item['evaluation_view'])
                # Raw graph differences remain in the evidence. A serializer is
                # not a precise patch merely because semantic checks pass.
            report['edits'].append(item)
    report['input_integrity']={p:hashlib.sha256(Path(p).read_bytes()).hexdigest()==h for p,h in hashes.items()}
    assert all(report['input_integrity'].values())
    (output/'report.json').write_text(json.dumps(report,indent=2,default=str),'utf-8')
    print(json.dumps(dict(output=str(output),formula=formula['classification'],edits=[dict(engine=x['engine'],op=x['operation'],semantic_diffs=len(x['semantic_unexpected']),precise=x['precise_part_patch']) for x in report['edits']]),indent=2))


if __name__=='__main__':
    p=argparse.ArgumentParser()
    for name in ('rust','aspose','output'): p.add_argument('--'+name,required=True)
    run(p.parse_args())
