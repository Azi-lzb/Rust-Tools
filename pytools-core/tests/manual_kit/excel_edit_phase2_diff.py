"""Independent semantic comparisons for the successful real-data candidates."""
import json
from pathlib import Path
import xml.etree.ElementTree as ET
import math
import openpyxl

repo=Path(__file__).resolve().parents[3]
evidence=repo/'artifacts/excel-edit-phase2'

def value(cell):
    v=cell.value
    if cell.data_type=='f' and not isinstance(v,str):
        v={'kind':type(v).__name__,'text':getattr(v,'text',None),'ref':getattr(v,'ref',None)}
    elif hasattr(v,'isoformat'):
        v=v.isoformat()
    return [cell.data_type,v]

def xml(obj):
    return ET.tostring(obj.to_tree(),encoding='unicode')

def snapshot(path):
    b=openpyxl.load_workbook(path)
    result={}
    for s in b:
        result[s.title]={
            'values':{c.coordinate:value(c) for c in s._cells.values() if c.value is not None},
            'styles':{c.coordinate:[str(c.font),str(c.fill),str(c.border),str(c.alignment),
                                    c.number_format,str(c.protection)]
                      for c in s._cells.values() if c.has_style},
            'comments':{c.coordinate:[c.comment.text,c.comment.author] for c in s._cells.values() if c.comment},
            'merges':sorted(map(str,s.merged_cells.ranges)),
            'row_dimensions':{str(k):[v.height,v.hidden,v.outlineLevel,v.collapsed] for k,v in s.row_dimensions.items()},
            'col_dimensions':{str(k):[v.min,v.max,v.width,v.hidden,v.outlineLevel,v.collapsed] for k,v in s.column_dimensions.items()},
            'printing':[xml(s.page_setup),xml(s.page_margins),xml(s.print_options),xml(s.HeaderFooter),str(s.print_area),str(s.print_title_rows),str(s.print_title_cols)],
        }
    b.close()
    return result

records=[]
special=json.loads((evidence/'special-results.json').read_text(encoding='utf-8'))
remaining=json.loads((evidence/'remaining-results.json').read_text(encoding='utf-8'))
for menu,source_records in [('5-2',remaining),('5-7',special)]:
    runs=[r for r in source_records if r['menu']==menu]
    baseline=next(r for r in runs if r['mode']=='OpenpyxlOnly')
    baseline_path=Path(baseline['case_directory'])/'source.xlsx' if menu=='5-2' else Path(baseline['output_copy'])
    expected=snapshot(baseline_path)
    original_path=Path(special[0]['input']).parent/'2026.09 外发用.xlsx' if menu=='5-2' else Path(baseline['input'])
    original=snapshot(original_path)
    for run in runs:
        path=Path(run['case_directory'])/'source.xlsx' if menu=='5-2' else Path(run['output_copy'])
        actual=snapshot(path)
        value_differences=[]
        exact_equal=True
        semantic_equal=True
        for sheet_name,expected_sheet in expected.items():
            actual_values=actual[sheet_name]['values']
            baseline_values=expected_sheet['values']
            for coordinate in set(actual_values)|set(baseline_values):
                left=actual_values.get(coordinate)
                right=baseline_values.get(coordinate)
                if left==right:
                    continue
                exact_equal=False
                close=False
                if (left is not None and right is not None and left[0]==right[0]=='n'
                        and isinstance(left[1],(int,float)) and isinstance(right[1],(int,float))):
                    close=math.isclose(left[1],right[1],rel_tol=0.0,abs_tol=1e-8)
                if not close:
                    semantic_equal=False
                value_differences.append({'sheet':sheet_name,'cell':coordinate,
                                          'baseline':right,'actual':left,
                                          'numeric_delta':(left[1]-right[1]
                                              if close else None)})
        record={'menu':menu,'mode':run['mode'],'output':str(path),
                'values_equal_openpyxl_baseline_exact':exact_equal,
                'values_equal_openpyxl_baseline_semantic':semantic_equal,
                'value_differences':value_differences,
                'styles_equal_original':all(actual[k]['styles']==v['styles'] for k,v in original.items()),
                'comments_equal_original':all(actual[k]['comments']==v['comments'] for k,v in original.items()),
                'printing_equal_original':all(actual[k]['printing']==v['printing'] for k,v in original.items()),
                'row_col_dimensions_equal_original':all(actual[k]['row_dimensions']==v['row_dimensions'] and actual[k]['col_dimensions']==v['col_dimensions'] for k,v in original.items()),
                'merges_equal_original':all(actual[k]['merges']==v['merges'] for k,v in original.items())}
        records.append(record)
        print(menu,run['mode'],record['values_equal_openpyxl_baseline_exact'],
              record['values_equal_openpyxl_baseline_semantic'],
              record['styles_equal_original'],flush=True)
imports=json.loads((evidence/'import-results.json').read_text(encoding='utf-8'))
for run in imports:
    if run['menu']!='5-6' or not run['changed']:continue
    source=snapshot(run['source'])['五篇大文章分机构']['values']
    target=snapshot(run['target'])['业务提取']['values']
    records.append({'menu':'5-6','mode':run['mode'],'imported_values_equal_source':source==target,
                    'existing_formulas_and_caches_equal':run['existing_formulas_and_caches_equal']})
(evidence/'semantic-differences.json').write_text(json.dumps(records,ensure_ascii=False,indent=2),encoding='utf-8')
