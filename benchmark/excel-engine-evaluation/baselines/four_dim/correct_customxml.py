"""Preserve the initial fixture; create a WorkbookPart-owned custom XML control."""
import hashlib
import json
import sys
import xml.etree.ElementTree as ET
import zipfile
from pathlib import Path

root=Path(sys.argv[1]).resolve();source=root/'fixtures/modify-complex.xlsx';target=root/'fixtures/modify-complex-workbook-customxml.xlsx'
if target.exists():raise FileExistsError(target)
with zipfile.ZipFile(source) as z:p={n:z.read(n) for n in z.namelist()}
P='http://schemas.openxmlformats.org/package/2006/relationships'
R='http://schemas.openxmlformats.org/officeDocument/2006/relationships'
C='http://schemas.openxmlformats.org/package/2006/content-types'
rels=ET.fromstring(p['_rels/.rels'])
for node in list(rels):
    if node.get('Id')=='benchmarkOpaque':rels.remove(node)
p['_rels/.rels']=ET.tostring(rels,encoding='utf-8')
rels=ET.fromstring(p['xl/_rels/workbook.xml.rels'])
ET.SubElement(rels,f'{{{P}}}Relationship',Id='benchmarkCustomXml',Type=R+'/customXml',Target='../customXml/item1.xml')
p['xl/_rels/workbook.xml.rels']=ET.tostring(rels,encoding='utf-8')
p['customXml/itemProps1.xml']=b'<ds:datastoreItem ds:itemID="{A88949E8-6510-4A69-9737-6AD4082D2AB5}" xmlns:ds="http://schemas.openxmlformats.org/officeDocument/2006/customXml"><ds:schemaRefs/></ds:datastoreItem>'
rels=ET.Element(f'{{{P}}}Relationships')
ET.SubElement(rels,f'{{{P}}}Relationship',Id='rId1',Type=R+'/customXmlProps',Target='itemProps1.xml')
p['customXml/_rels/item1.xml.rels']=ET.tostring(rels,encoding='utf-8')
types=ET.fromstring(p['[Content_Types].xml'])
ET.SubElement(types,f'{{{C}}}Override',PartName='/customXml/itemProps1.xml',ContentType='application/vnd.openxmlformats-officedocument.customXmlProperties+xml')
p['[Content_Types].xml']=ET.tostring(types,encoding='utf-8')
with zipfile.ZipFile(target,'w',zipfile.ZIP_DEFLATED) as z:
    for n,b in p.items():z.writestr(n,b)
(root/'customxml-fixture-correction.json').write_text(json.dumps(dict(original=str(source),corrected=str(target),sha256=hashlib.sha256(target.read_bytes()).hexdigest(),
    changed_parts=['_rels/.rels','xl/_rels/workbook.xml.rels','[Content_Types].xml'],added_parts=['customXml/itemProps1.xml','customXml/_rels/item1.xml.rels'],
    note='WorkbookPart customXml relationship and datastore properties; all worksheet/style/chart parts unchanged'),indent=2),'utf-8')
print(target)
