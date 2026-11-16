"""Reanalyze saved outputs only: never execute engines or modify workbook bytes."""
import json
import sys
from pathlib import Path
from capabilities import eval_view, view_diff

root=Path(sys.argv[1])
report_path=root/'report.json'
report=json.loads(report_path.read_text('utf-8'))
source=root/'edit-input.xlsx'
base=eval_view(source,source)
noop=eval_view(source,root/'aspose-noop.xlsx')
for item in report['edits']:
    if item['engine']!='aspose': continue
    item['evaluation_view']=eval_view(source,root/f'aspose-{item["operation"]}.xlsx')
    item['filtered_package_diff']=view_diff(base,item['evaluation_view'])
    if item['operation']!='noop': item['versus_own_noop_filtered']=view_diff(noop,item['evaluation_view'])
report_path.write_text(json.dumps(report,indent=2),'utf-8')
print(report_path)
