from http.server import HTTPServer, BaseHTTPRequestHandler
from pathlib import Path
import json, base64, secrets, webbrowser, argparse
from inspector import inspect
ROOT=Path(__file__).parent
TOKEN=secrets.token_urlsafe(32)
class Handler(BaseHTTPRequestHandler):
    def reply(self,status,data,kind='application/json; charset=utf-8'):
        self.send_response(status);self.send_header('Content-Type',kind);self.send_header('X-Content-Type-Options','nosniff');self.send_header('Cache-Control','no-store');self.end_headers();self.wfile.write(data)
    def do_GET(self):
        if self.path!='/': self.reply(404,b'Not found','text/plain');return
        self.reply(200,(ROOT/'web/index.html').read_text(encoding='utf-8').replace('__TOKEN__',TOKEN).encode(),'text/html; charset=utf-8')
    def do_POST(self):
        if self.path!='/api/inspect': self.reply(404,b'{}');return
        if self.headers.get('X-Inspector-Token')!=TOKEN: self.reply(403,b'{}');return
        try:
            length=int(self.headers.get('Content-Length',0))
            if not 0<length<=42*1024*1024: raise ValueError('单个上传请求上限 42 MB')
            body=json.loads(self.rfile.read(length)); name=str(body['name'])
            if not name.lower().endswith(('.xlsx','.xlsm')): raise ValueError('只支持 xlsx / xlsm')
            size=float(body.get('size',10));font=str(body.get('font','')).strip()
            if not font or not 0<size<=409: raise ValueError('请输入有效字体和字号')
            data=base64.b64decode(body['data'],validate=True)
            result=inspect(data,name,font,size)
            self.reply(200,json.dumps(result,ensure_ascii=False,allow_nan=False).encode())
        except Exception as e:self.reply(400,json.dumps({'error':str(e)},ensure_ascii=False).encode())
    def log_message(self,*args):pass
if __name__=='__main__':
    p=argparse.ArgumentParser();p.add_argument('--port',type=int,default=8765);p.add_argument('--no-browser',action='store_true');a=p.parse_args()
    server=HTTPServer(('127.0.0.1',a.port),Handler)
    url=f'http://127.0.0.1:{a.port}';print('Excel Inspector: '+url)
    if not a.no_browser:webbrowser.open(url)
    try:server.serve_forever()
    except KeyboardInterrupt:pass
    finally:server.server_close()
