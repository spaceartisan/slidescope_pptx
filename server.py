"""Offline PowerPoint browser harness. Usage: python server.py [presentation.pptx]

Uses only Python 3.12 stdlib plus python-pptx/lxml; binds to loopback by default.
"""
from __future__ import annotations
import argparse
import json
from pathlib import Path
from http.server import ThreadingHTTPServer, BaseHTTPRequestHandler
from threading import Lock
from urllib.parse import urlsplit, unquote
import mimetypes
import re
from zipfile import ZipFile, BadZipFile
from io import BytesIO

from pptx_parser import parse_pptx, PREVIEW_LIMITS

ROOT=Path(__file__).resolve().parent
STATIC=(ROOT/'static').resolve()

class DeckState:
    def __init__(self, initial=None):
        self.lock=Lock(); self.model=None; self.assets={}; self.path=initial; self.modified=None; self.error=None; self.version=0; self.references={}; self.reference_name=None
    def update(self,data,filename):
        model,assets=parse_pptx(data,filename)
        with self.lock:
            self.model=model; self.assets=assets; self.error=None; self.version+=1
        return model
    def load_references(self, data, name):
        if len(data) > PREVIEW_LIMITS['bytes']:
            raise ValueError('Reference ZIP exceeds 64MB')
        images={}
        try:
            with ZipFile(BytesIO(data)) as archive:
                infos=archive.infolist()
                if len(infos)>500 or sum(info.file_size for info in infos)>150*1024*1024:
                    raise ValueError('Reference ZIP has too many or too-large files')
                for info in infos:
                    match=re.fullmatch(r'Slide(\d+)\.png',Path(info.filename).name,re.I)
                    if not match:continue
                    number=int(match.group(1))
                    if not 1<=number<=400 or info.file_size>20*1024*1024:continue
                    raw=archive.read(info)
                    if raw[:8]!=b'\x89PNG\r\n\x1a\n':continue
                    images[number]=raw
        except (BadZipFile, OSError) as exc:
            raise ValueError('Invalid reference ZIP') from exc
        if not images:raise ValueError('ZIP must contain Slide1.PNG, Slide2.PNG, etc.')
        with self.lock:
            self.references=images;self.reference_name=name
        return len(images)
    def refresh_file(self):
        if self.path is None: return
        try:
            stamp=(self.path.stat().st_mtime_ns,self.path.stat().st_size)
            if stamp!=self.modified:
                self.update(self.path.read_bytes(),self.path.name)
                self.modified=stamp
        except Exception as err:
            self.error=str(err)

class Handler(BaseHTTPRequestHandler):
    server_version='PPTXBrowser/0.1'
    def log_message(self,format,*args):
        print('%s - %s' % (self.address_string(), format%args))
    def reply(self,code,data,typ='application/json; charset=utf-8',cache='no-store'):
        self.send_response(code)
        self.send_header('Content-Type',typ)
        self.send_header('Content-Length',str(len(data)))
        self.send_header('Cache-Control',cache)
        self.send_header('X-Content-Type-Options','nosniff')
        self.send_header('X-Frame-Options','DENY')
        self.end_headers()
        self.wfile.write(data)
    def json(self,code,obj):
        self.reply(code,json.dumps(obj,ensure_ascii=False).encode('utf8'))
    def do_GET(self):
        url=urlsplit(self.path).path
        if url in ('/api/deck','/api/status'):
            state.refresh_file()
            with state.lock:
                if url=='/api/status':
                    self.json(200,{'loaded':state.model is not None,'filename':state.model['filename'] if state.model else None,'error':state.error,'version':state.version})
                else: self.json(200,state.model or {'count':0,'slides':[],'warnings':[]})
            return
        if url=='/api/references':
            with state.lock: self.json(200,{'count':len(state.references),'slides':sorted(state.references),'filename':state.reference_name})
            return
        if url.startswith('/api/reference/'):
            match=re.fullmatch(r'/api/reference/(\d+)',url)
            if not match:self.send_error(404);return
            with state.lock: raw=state.references.get(int(match.group(1)))
            if raw is None:self.send_error(404);return
            self.reply(200,raw,'image/png')
            return
        if url.startswith('/api/media/'):
            key=url[len('/api/media/'):]
            with state.lock:
                asset=state.assets.get(key)
            if not asset:
                self.send_error(404);return
            self.reply(200,asset['bytes'],asset['type'])
            return
        if url=='/': url='/index.html'
        path=(STATIC/unquote(url.lstrip('/'))).resolve()
        if not path.is_relative_to(STATIC) or not path.is_file():
            self.send_error(404);return
        ctype=mimetypes.guess_type(path.name)[0] or 'application/octet-stream'
        self.reply(200,path.read_bytes(),ctype+'; charset=utf-8' if ctype.startswith(('text/','application/javascript')) else ctype)
    def do_POST(self):
        path=urlsplit(self.path).path
        if path not in ('/api/upload','/api/reference'):
            self.send_error(404);return
        try:
            size=int(self.headers.get('Content-Length','0'))
            if size<=0 or size>PREVIEW_LIMITS['bytes']:
                self.json(413,{'error':'File must be between 1 byte and 64 MB'});return
            name=Path(unquote(self.headers.get('X-Filename','presentation.pptx').replace('\\','/'))).name[:200]
            if path=='/api/reference':
                if not name.lower().endswith('.zip'):
                    self.json(400,{'error':'Select the ZIP exported from PowerPoint'});return
                count=state.load_references(self.rfile.read(size),name)
                self.json(200,{'ok':True,'slides':count})
                return
            if not name.lower().endswith('.pptx'):
                self.json(400,{'error':'Select a .pptx file'});return
            data=self.rfile.read(size)
            state.update(data,name)
            # Browser uploads deliberately override CLI file watching.
            state.path=None;state.modified=None
            self.json(200,{'ok':True,'filename':name,'slides':state.model['count']})
        except (ValueError, KeyError, OSError) as exc:
            self.json(400,{'error':str(exc)})
        except Exception as exc:
            self.json(422,{'error':f'Cannot preview presentation: {exc}'})

if __name__=='__main__':
    p=argparse.ArgumentParser(description='Preview PPTX locally in a browser')
    p.add_argument('pptx',nargs='?',help='Optional initial PPTX; file changes are watched')
    p.add_argument('--port',type=int,default=8000)
    p.add_argument('--host',default='127.0.0.1',help='Default localhost only')
    args=p.parse_args()
    initial=Path(args.pptx).expanduser().resolve() if args.pptx else None
    state=DeckState(initial)
    if initial:
        state.refresh_file()
        if state.error: print('Could not load initial file:',state.error)
    httpd=ThreadingHTTPServer((args.host,args.port),Handler)
    print(f'PPTX browser ready at http://{args.host}:{args.port}')
    print('Open this address in your VS Code integrated browser, then upload a .pptx.')
    try: httpd.serve_forever()
    except KeyboardInterrupt: print('\nStopping')
    finally: httpd.server_close()
