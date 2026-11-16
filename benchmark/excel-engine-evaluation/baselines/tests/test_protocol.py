import json
import sys
import tempfile
import unittest
from contextlib import nullcontext
from pathlib import Path
import openpyxl

BASE = Path(__file__).resolve().parents[1]
sys.path.insert(0,str(BASE))
from adapters.python.worker import digest
from compare.semantic import verify
from runner import Worker


class ValidationTests(unittest.TestCase):
    def fixture(self, root):
        path = root/'input.xlsx'
        data = [[0,'0000123',True,None,'text',-17,'long text',9], [10,'Z/AA',False,None,'=literal',0,'end',100]]
        # This fixture is scalar only; avoid interpreting leading '=' as a formula.
        data[1][4] = 'literal'
        wb = openpyxl.Workbook(); ws=wb.active; ws.title='Data'
        for row in data: ws.append(row)
        wb.save(path)
        d=root/'data.json'; d.write_text(json.dumps(data),'utf-8')
        return dict(id='tiny', input=str(path),data=str(d), rows=2,cols=8,sheets=['Data'],digest=digest(data))

    def test_detect_mismatch_extra_and_vendor_lookalike(self):
        with tempfile.TemporaryDirectory() as tmp:
            root=Path(tmp); f=self.fixture(root)
            self.assertTrue(verify(f['input'],f)['ok'])
            wb=openpyxl.load_workbook(f['input']); wb['Data']['A1']=1; wb.save(root/'bad.xlsx')
            with self.assertRaisesRegex(AssertionError,'digest'): verify(root/'bad.xlsx',f)
            wb=openpyxl.load_workbook(f['input']); wb.create_sheet('Evaluation Warning')['A1']='user data'; wb.save(root/'fake.xlsx')
            with self.assertRaisesRegex(AssertionError,'unattributed'): verify(root/'fake.xlsx',f,True)
            # Even a vendor-looking original sheet is part of input identity, not filtered away.
            f['sheets'].append('Evaluation Warning')
            self.assertEqual(verify(root/'fake.xlsx',f,True)['excluded_evaluation'],[])

    def test_all_adapters_small_and_existing_output_refused(self):
        commands = {
            'rust':['F:/RustBuild/pytools/release/excel-io-baseline.exe'],
            'openpyxl':[sys.executable,'-u',str(BASE/'adapters/python/worker.py')],
            'aspose':['F:/RustBuild/pytools/io-dotnet/IoBaseline.exe'],
        }
        # Preserve real SDK evaluation outputs even for protocol smoke tests.
        saved=tempfile.mkdtemp(prefix='protocol-',dir=BASE/'results')
        with nullcontext(saved) as tmp:
            root=Path(tmp); f=self.fixture(root)
            print(f'preserved protocol outputs: {root}',flush=True)
            for engine,cmd in commands.items():
                w=Worker(cmd,root/f'{engine}.stderr')
                try:
                    self.assertTrue(w.request(dict(op='init',fixtures=[f]))['ok'])
                    for op in ('read','write','roundtrip'):
                        q=dict(op=op,id='tiny',input=f['input'])
                        if op!='read': q['output']=str(root/f'{engine}-{op}.xlsx')
                        r=w.request(q)
                        self.assertTrue(r['ok'],r)
                        self.assertEqual(r['digest'],f['digest'])
                        self.assertGreater(r['processing_ms'],0)
                        if op!='read':
                            self.assertTrue(verify(q['output'],f,engine=='aspose')['ok'])
                            before=Path(q['output']).read_bytes()
                            self.assertFalse(w.request(q)['ok'])
                            self.assertEqual(Path(q['output']).read_bytes(),before)
                finally: w.close()


if __name__=='__main__': unittest.main()
