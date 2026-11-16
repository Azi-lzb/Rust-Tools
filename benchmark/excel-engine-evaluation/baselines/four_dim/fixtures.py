import hashlib
import json
import sys
import xml.etree.ElementTree as ET
import zipfile
from pathlib import Path
import openpyxl
from openpyxl.chart import BarChart,Reference
from openpyxl.comments import Comment
from openpyxl.styles import Font,PatternFill
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.table import Table,TableStyleInfo
from openpyxl.workbook.defined_name import DefinedName

BASE=Path(__file__).resolve().parents[1]
sys.path.insert(0,str(BASE/'adapters/python'))
from worker import digest
M='http://schemas.openxmlformats.org/spreadsheetml/2006/main'


def write_value(r,c):
    k=c%5
    return r*20+c if k==0 else f'row-{r:06}-col-{c:02}' if k==1 else (r+c)%2==0 if k==2 else None if k==3 else -(r+c)


def calc_spec():
    n=256; data=[['kind','quantity','price','key','code']]
    data += [['alpha' if i%2==0 else 'beta',i+1,i%7+1,f'key-{i}',i*3] for i in range(n)]
    total=sum(i+1 for i in range(n));alpha=sum(i+1 for i in range(n) if i%2==0)
    product=sum((i+1)*(i%7+1) for i in range(n)); formulas=[]; expected=[]
    for i in range(n):
        r=i+2
        formulas.append([f'SUM(Data!B$2:B$257)+Data!B{r}',
          'SUMIF(Data!A$2:A$257,"alpha",Data!B$2:B$257)',
          'COUNTIF(Data!B$2:B$257,">100")',
          f'IF(AND(Data!B{r}>0,Data!C{r}>0),ROUND(Data!B{r}*Data!C{r},2),0)',
          'VLOOKUP("key-123",Data!D$2:E$257,2,FALSE)',
          'INDEX(Data!C$2:C$257,MATCH("key-123",Data!D$2:D$257,0))',
          f'IF(OR(Data!B{r}=0,Data!C{r}=3),1,0)',
          'SUMPRODUCT(Data!B$2:B$257,Data!C$2:C$257)'])
        expected.extend([total+i+1,alpha,156,(i+1)*(i%7+1),369,5,int(i%7+1==3),product])
    return dict(data=data,formulas=formulas,expected=expected,wrong_cache=-9999,formula_count=2048)


def complex_fixture(path):
    wb=openpyxl.Workbook();ws=wb.active;ws.title='Data'
    ws.append(['key','quantity','price','amount','text','merged','other','extra','flag','marker'])
    for r in range(2,10002): ws.append([f'item-{r}',r,r+10,f'=B{r}*C{r}',f'text-{r}','',None,None,r%2==0,r])
    ws.merge_cells('F1:G1');ws['F1']='complex fixture';ws['A1'].fill=PatternFill('solid',fgColor='FFFF00')
    ws['A2'].comment=Comment('synthetic preservation sentinel','RustTools benchmark')
    ws['A3'].hyperlink='https://example.invalid/benchmark'
    ws['J12'].font=Font(name='Calibri',size=11,bold=False)
    ws.row_dimensions[1].height=28;ws.column_dimensions['A'].width=21
    ws.freeze_panes='B2';ws.auto_filter.ref='A1:J10001'
    dv=DataValidation(type='whole',operator='between',formula1='0',formula2='100000');ws.add_data_validation(dv);dv.add('B2:B10001')
    table=Table(displayName='DataTable',ref='A1:D10001');table.tableStyleInfo=TableStyleInfo(name='TableStyleMedium2',showRowStripes=True);ws.add_table(table)
    chart=BarChart();chart.title='benchmark quantities';chart.add_data(Reference(ws,min_col=2,min_row=1,max_row=11),titles_from_data=True);ws.add_chart(chart,'L2')
    ws.print_area='A1:J10001';ws.print_title_rows='1:1';ws.page_setup.orientation='landscape';ws.page_setup.paperSize='9'
    ws.oddHeader.left.text='benchmark &Bheader';ws.oddFooter.right.text='Page &P / &N'
    calc=wb.create_sheet('Calc');calc['A1']='=SUM(Data!B2:B10001)'
    wb.defined_names.add(DefinedName('Quantity',attr_text="'Data'!$B$2:$B$10001"))
    wb.save(path)
    with zipfile.ZipFile(path) as z: package={n:z.read(n) for n in z.namelist()}
    # Known correct caches, unaffected by target J12 or the header edits.
    for part in ['xl/worksheets/sheet1.xml','xl/worksheets/sheet2.xml']:
        root=ET.fromstring(package[part])
        for cell in root.iter(f'{{{M}}}c'):
            if cell.find(f'{{{M}}}f') is not None:
                v=cell.find(f'{{{M}}}v')
                if v is None:v=ET.SubElement(cell,f'{{{M}}}v')
                row=int(''.join(c for c in cell.get('r') if c.isdigit()))
                v.text=str(row*(row+10) if part.endswith('sheet1.xml') else 50015000)
        package[part]=ET.tostring(root,encoding='utf-8',xml_declaration=True)
    p='http://schemas.openxmlformats.org/package/2006/relationships'
    relation='http://schemas.openxmlformats.org/officeDocument/2006/relationships'
    rels=ET.fromstring(package['xl/_rels/workbook.xml.rels'])
    ET.SubElement(rels,f'{{{p}}}Relationship',Id='benchmarkCustomXml',Type=relation+'/customXml',Target='../customXml/item1.xml')
    package['xl/_rels/workbook.xml.rels']=ET.tostring(rels,encoding='utf-8')
    package['customXml/item1.xml']=b'<sentinel xmlns="urn:rusttools:benchmark">preserve-unmodified-part</sentinel>'
    package['customXml/itemProps1.xml']=b'<ds:datastoreItem ds:itemID="{A88949E8-6510-4A69-9737-6AD4082D2AB5}" xmlns:ds="http://schemas.openxmlformats.org/officeDocument/2006/customXml"><ds:schemaRefs/></ds:datastoreItem>'
    rels=ET.Element(f'{{{p}}}Relationships')
    ET.SubElement(rels,f'{{{p}}}Relationship',Id='rId1',Type=relation+'/customXmlProps',Target='itemProps1.xml')
    package['customXml/_rels/item1.xml.rels']=ET.tostring(rels,encoding='utf-8')
    content='http://schemas.openxmlformats.org/package/2006/content-types'
    types=ET.fromstring(package['[Content_Types].xml'])
    ET.SubElement(types,f'{{{content}}}Override',PartName='/customXml/itemProps1.xml',ContentType='application/vnd.openxmlformats-officedocument.customXmlProperties+xml')
    package['[Content_Types].xml']=ET.tostring(types,encoding='utf-8')
    with zipfile.ZipFile(path,'w',zipfile.ZIP_DEFLATED) as z:
        for n,b in package.items():z.writestr(n,b)


def generate(root):
    root=Path(root);root.mkdir(parents=True,exist_ok=False)
    (root/'calc-spec.json').write_text(json.dumps(calc_spec(),indent=2),'utf-8')
    complex_fixture(root/'modify-complex.xlsx')
    matrix=[[write_value(r,c) for c in range(20)] for r in range(50000)]
    write=dict(id='write-50kx20',rows=50000,cols=20,sheets=['Data'],positions=1000000,nonempty_cells=800000,digest=digest(matrix))
    previous=json.loads((BASE/'fixtures/generated/manifest.json').read_text('utf-8'))
    manifest=dict(read=previous,write=write,modify=str((root/'modify-complex.xlsx').resolve()),calc_expected=calc_spec()['expected'])
    (root/'manifest.json').write_text(json.dumps(manifest,indent=2),'utf-8')
    return manifest


if __name__=='__main__': print(json.dumps(generate(sys.argv[1]),indent=2)[:1800])
