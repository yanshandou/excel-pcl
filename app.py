from http.server import HTTPServer, BaseHTTPRequestHandler
from pathlib import Path
import json, base64, secrets, webbrowser, argparse
from zipfile import BadZipFile
from openpyxl.utils.exceptions import InvalidFileException
from inspector import inspect
from comparison import compare_reports
ROOT=Path(__file__).parent
TOKEN=secrets.token_urlsafe(32)
class Handler(BaseHTTPRequestHandler):
    def reply(self,status,data,kind='application/json; charset=utf-8'):
        self.send_response(status);self.send_header('Content-Type',kind);self.send_header('X-Content-Type-Options','nosniff');self.send_header('Cache-Control','no-store');self.end_headers();self.wfile.write(data)
    def do_GET(self):
        # Serve only explicitly public assets; never expose arbitrary project files.
        routes = {
            '/': ('index.html', 'text/html; charset=utf-8'),
            '/css/styles.css': ('css/styles.css', 'text/css; charset=utf-8'),
            '/js/app.js': ('js/app.js', 'text/javascript; charset=utf-8'),
            '/js/comparison.js': ('js/comparison.js', 'text/javascript; charset=utf-8'),
        }
        route = routes.get(self.path.split('?', 1)[0])
        if route is None:
            self.reply(404, b'Not found', 'text/plain; charset=utf-8')
            return
        filename, content_type = route
        content = (ROOT / 'web' / filename).read_text(encoding='utf-8')
        if filename == 'index.html':
            content = content.replace('__TOKEN__', TOKEN)
        self.reply(200, content.encode('utf-8'), content_type)
    def do_POST(self):
        if self.path not in ('/api/inspect','/api/compare'): self.reply(404,b'{}');return
        if self.headers.get('X-Inspector-Token')!=TOKEN: self.reply(403,b'{}');return
        try:
            length=int(self.headers.get('Content-Length',0))
            limit = 84 if self.path == '/api/compare' else 42
            if not 0<length<=limit*1024*1024: raise ValueError(f'上传请求上限 {limit} MB')
            body=json.loads(self.rfile.read(length))
            size=float(body.get('size',10));font=str(body.get('font','')).strip()
            if not font or not 0<size<=409: raise ValueError('请输入有效字体和字号')
            def read_file(item):
                name=str(item['name'])
                if not name.lower().endswith(('.xlsx','.xlsm')): raise ValueError('只支持 xlsx / xlsm')
                return inspect(base64.b64decode(item['data'],validate=True),name,font,size)
            if self.path == '/api/compare':
                before=read_file(body['old']);after=read_file(body['new'])
                result={'comparison':compare_reports(before,after,body.get('mapping')),
                        'reports':[before,after]}
            else:
                result=read_file(body)
            self.reply(200,json.dumps(result,ensure_ascii=False,allow_nan=False).encode())
        except (BadZipFile, InvalidFileException):
            self.reply(400,json.dumps({'error':'无法读取工作簿：文件可能已损坏、已加密，或不是有效的 xlsx / xlsm 文件'},ensure_ascii=False).encode())
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
