"""API tests exercise the real handler without requiring local socket permissions."""
from io import BytesIO
import base64
import json
import unittest
from openpyxl import Workbook
from app import Handler, TOKEN


class ApiTest(unittest.TestCase):
    def request(self, path, body, token=TOKEN):
        data=json.dumps(body).encode()
        handler=object.__new__(Handler)
        handler.path=path;handler.headers={'Content-Length':str(len(data)),'X-Inspector-Token':token}
        handler.rfile=BytesIO(data);captured=[]
        handler.reply=lambda status,data,kind=None:captured.append((status,data,kind))
        handler.do_POST()
        status,payload,_=captured[0]
        return status,json.loads(payload)

    def test_compare_api_and_inspection_retained(self):
        book=Workbook();book.active['A1']='before';stream=BytesIO();book.save(stream)
        old={'name':'old.xlsx','data':base64.b64encode(stream.getvalue()).decode()}
        book.active['A1']='after';book.active['A1'].number_format='@';stream=BytesIO();book.save(stream)
        new={'name':'new.xlsx','data':base64.b64encode(stream.getvalue()).decode()}
        status,payload=self.request('/api/compare',{'old':old,'new':new,'font':'Target','size':10})
        self.assertEqual(status,200);self.assertEqual(len(payload['reports']),2)
        self.assertEqual(payload['comparison']['summary']['content'],1)
        self.assertEqual(payload['comparison']['summary']['format'],1)
        self.assertEqual(len(payload['reports'][1]['sheets'][0]['cells'][0]['colors']),7)
        status,_=self.request('/api/inspect',{**old,'font':'Target','size':10})
        self.assertEqual(status,200)

    def test_token_and_invalid_workbook(self):
        self.assertEqual(self.request('/api/compare',{},'wrong')[0],403)
        status,payload=self.request('/api/compare',{'font':'Target','old':{'name':'bad.xlsx','data':base64.b64encode(b'bad').decode()}})
        self.assertEqual(status,400);self.assertIn('无法读取工作簿',payload['error'])

    def test_static_routes(self):
        for path,status in [('/',200),('/js/comparison.js',200),('/css/styles.css',200),('/comparison.py',404),('/../app.py',404)]:
            handler=object.__new__(Handler);handler.path=path;captured=[]
            handler.reply=lambda code,data,kind:captured.append((code,data,kind))
            handler.do_GET();self.assertEqual(captured[0][0],status)
            if path=='/':self.assertIn(TOKEN.encode(),captured[0][1])

if __name__=='__main__':unittest.main()
