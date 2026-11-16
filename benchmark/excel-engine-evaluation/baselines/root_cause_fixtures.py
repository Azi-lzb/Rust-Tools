"""Independent minimal fixtures for the two observed library discrepancies."""
import hashlib
import json
import struct
import sys
import xml.etree.ElementTree as ET
import zipfile
from pathlib import Path
import openpyxl
import xlrd
import xlwt
from openpyxl.worksheet.datavalidation import DataValidation


def specs():
    formulas=['SUMIF(Data!A2:A4,"alpha",Data!B2:B4)',
              'SUMPRODUCT(Data!B2:B4,Data!C2:C4)',
              'IF(AND(Data!B2>0,Data!C2>0),ROUND(Data!B2*Data!C2,2),0)',
              'INDEX(Data!C2:C4,MATCH("beta",Data!A2:A4,0))']
    formulas += ['B2','$B2','B$2','$B$2','SUM(Z99:AA100)','SUM($Z99:AA$100)']
    for sheet in ['Data',"'Data Set'","'O''Brien'"]:
        formulas += [sheet+'!'+r for r in ['B2','$B2','B$2','$B$2']]
        formulas += ['SUM('+sheet+'!'+r+')' for r in ['Z9:AA10','$Z9:AA$10','Z$99:$AA100','$Z$99:$AA$100']]
    formulas += ['Data!'+r for r in ['Z9','AA10','Z99','AA100']]
    formulas += ['Data!B2'+op+'0' for op in ['>','>=','<','<=','=','<>']]
    return dict(sheet_names=['Calc','Spare','Data','Data Set',"O'Brien"],formulas=formulas,
                values=[['A1','key'],['B1','quantity'],['C1','price'],['A2','alpha'],['B2',2],['C2',3],
                        ['A3','beta'],['B3',4],['C3',5],['A4','alpha'],['B4',6],['C4',7],
                        ['Z9',13],['AA10',17],['Z99',19],['AA100',23]])


def biff(path):
    stream=xlrd.compdoc.CompDoc(path.read_bytes()).get_named_stream('Workbook')
    if stream is None: stream=xlrd.compdoc.CompDoc(path.read_bytes()).get_named_stream('Book')
    offset=0; records=[]
    while offset+4<=len(stream):
        code,length=struct.unpack_from('<HH',stream,offset); data=stream[offset+4:offset+4+length]
        if code in (0x0006,0x0017):
            item=dict(offset=offset,record_type=hex(code),payload_hex=data.hex())
            if code==6: item.update(row=struct.unpack_from('<H',data)[0]+1,col=struct.unpack_from('<H',data,2)[0]+1,cache_hex=data[6:14].hex(),tokens_hex=data[22:22+struct.unpack_from('<H',data,20)[0]].hex())
            records.append(item)
        offset+=4+length
    return records


def main(root):
    root=Path(root); root.mkdir(exist_ok=True)
    spec=specs(); (root/'spec.json').write_text(json.dumps(spec,indent=2),'utf-8')
    if any(root.glob('*.xls')) or any(root.glob('*.xlsx')): raise FileExistsError('refuse overwriting fixtures')
    xl=xlwt.Workbook(); px=openpyxl.Workbook(); px.remove(px.active)
    xs={name:xl.add_sheet(name) for name in spec['sheet_names']}
    for name in spec['sheet_names']: px.create_sheet(name)
    for name in spec['sheet_names']:
        ws=xs[name]
        for address,value in spec['values']:
            row,col=openpyxl.utils.cell.coordinate_to_tuple(address)
            if name=='Calc' and col==1: continue
            ws.write(row-1,col-1,value); px[name][address]=value
        if name=='Calc':
            for r,f in enumerate(spec['formulas']): ws.write(r,0,xlwt.Formula(f))
    for r,f in enumerate(spec['formulas'],1): px['Calc'].cell(r,1,'='+f)
    xl.save(str(root/'xlwt-references.xls')); px.save(root/'openpyxl-control.xlsx')
    wb=openpyxl.Workbook(); ws=wb.active; ws.title='Data'; ws['A1']='minimal validation fixture'
    cases=[('whole','between','0','1000'),('whole','notBetween','-9','99'),
           ('decimal','between','-1.5','2.75'),('whole','greaterThan','0',None),
           ('list',None,'"A,B"',None),('custom',None,'AND(A1>0,A1<1000)',None),
           (None,None,None,None)]
    for i,(kind,op,a,b) in enumerate(cases,2):
        dv=DataValidation(type=kind,operator=op,formula1=a,formula2=b)
        dv.promptTitle='synthetic'; dv.prompt='keep bounds & message'
        wb.active.add_data_validation(dv); dv.add(f'{openpyxl.utils.get_column_letter(i)}2:{openpyxl.utils.get_column_letter(i)}16')
    wb.save(root/'validation-minimal.xlsx')
    with zipfile.ZipFile(root/'validation-minimal.xlsx') as z:
        xml=z.read('xl/worksheets/sheet1.xml')
        (root/'validation-input-worksheet.xml').write_bytes(xml)
    (root/'xlwt-biff-records.json').write_text(json.dumps(biff(root/'xlwt-references.xls'),indent=2),'utf-8')
    hashes={p.name:hashlib.sha256(p.read_bytes()).hexdigest() for p in root.glob('*') if p.suffix in ('.xls','.xlsx')}
    (root/'input-sha256.json').write_text(json.dumps(hashes,indent=2),'utf-8')
    print('formula cases',len(spec['formulas']),'validation cases',len(cases))


if __name__=='__main__': main(sys.argv[1])
