import unittest
from io import BytesIO
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment
from inspector import inspect
class InspectionTest(unittest.TestCase):
    def test_rules_and_metadata(self):
        w=Workbook();s=w.active;s.title='检查'
        s['A1']='@model@';s['A1'].number_format='@';s['A1'].font=Font(name='Target',size=10,bold=True,color='FF000000')
        s['A1'].alignment=Alignment(horizontal='left',vertical='distributed',wrap_text=True,shrink_to_fit=True)
        s['B2']='=1+1';s['B2'].font=Font(name='Other',size=10);s.merge_cells('C3:D3');s['C3']='合并';s.print_area='A1:C3';s['E5']='越界'
        s['A2']=1;s['A2'].number_format='0_'
        b=BytesIO();w.save(b);raw=b.getvalue();r=inspect(raw,'test.xlsx','Target',10)
        cells={c['address']:c for c in r['sheets'][0]['cells']}
        self.assertEqual(cells['A1']['colors'],['blue','red','blue','blue','blue','none','blue'])
        self.assertTrue(cells['A1']['pattern']);self.assertEqual(cells['B2']['colors'][6],'red');self.assertEqual(cells['C3']['colors'][5],'blue')
        self.assertEqual(cells['A2']['colors'][0],'green');self.assertEqual(r['sheets'][0]['outside'],'有り');self.assertTrue(r['new_type']);self.assertEqual(raw,b.getvalue())
    def test_missing_and_multi_print_area(self):
        w=Workbook();s=w.active;b=BytesIO();w.save(b)
        self.assertEqual(inspect(b.getvalue(),'blank.xlsx')['sheets'][0]['outside'],'不明')
        s.print_area=['A1:B2','D4:E5'];b=BytesIO();w.save(b)
        self.assertEqual(inspect(b.getvalue(),'multi.xlsx')['sheets'][0]['outside'],'不明')
if __name__=='__main__':unittest.main()
