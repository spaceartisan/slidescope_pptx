import json
import io
from zipfile import ZipFile
from pathlib import Path
import sys
import threading
import unittest
from urllib import request,error
from http.server import ThreadingHTTPServer

ROOT=Path(__file__).resolve().parents[1]
sys.path.insert(0,str(ROOT))
import server

class APITests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        server.state=server.DeckState()
        cls.httpd=ThreadingHTTPServer(('127.0.0.1',0),server.Handler)
        cls.thread=threading.Thread(target=cls.httpd.serve_forever, daemon=True)
        cls.thread.start()
        cls.base=f'http://127.0.0.1:{cls.httpd.server_port}'

    @classmethod
    def tearDownClass(cls):
        cls.httpd.shutdown();cls.httpd.server_close();cls.thread.join(timeout=2)

    def test_static_and_initial_empty_deck(self):
        with request.urlopen(self.base+'/') as response:
            self.assertEqual(response.status,200)
            self.assertIn(b'SLIDESCOPE',response.read())
        with request.urlopen(self.base+'/api/deck') as response:
            self.assertEqual(json.load(response)['count'],0)

    def test_upload_and_media(self):
        raw=(ROOT/'sample/demo.pptx').read_bytes()
        req=request.Request(self.base+'/api/upload',raw,headers={'X-Filename':'demo.pptx','Content-Type':'application/octet-stream'},method='POST')
        with request.urlopen(req) as response:
            self.assertEqual(json.load(response)['slides'],4)
        with request.urlopen(self.base+'/api/deck') as response:
            model=json.load(response)
            self.assertEqual(model['count'],4)
            pic=next(o for o in model['slides'][3]['shapes'] if o['kind']=='image')
        with request.urlopen(self.base+pic['image']['url']) as response:
            self.assertEqual(response.status,200)
            self.assertTrue(response.headers['Content-Type'].startswith('image/'))
        with request.urlopen(self.base+'/api/status') as response:
            self.assertGreaterEqual(json.load(response)['version'],1)

    def test_reference_roundtrip(self):
        buf=io.BytesIO()
        with ZipFile(buf,'w') as archive:
            archive.writestr('Export/Slide1.PNG',b'\x89PNG\r\n\x1a\n' + b'sample')
        req=request.Request(self.base+'/api/reference',buf.getvalue(),headers={'X-Filename':'exports.zip'},method='POST')
        with request.urlopen(req) as response:
            self.assertEqual(json.load(response)['slides'],1)
        with request.urlopen(self.base+'/api/references') as response:
            self.assertEqual(json.load(response)['slides'],[1])
        with request.urlopen(self.base+'/api/reference/1') as response:
            self.assertEqual(response.headers['Content-Type'],'image/png')
            self.assertTrue(response.read().startswith(b'\x89PNG'))
        with self.assertRaises(error.HTTPError) as cm:request.urlopen(self.base+'/api/reference/2')
        self.assertEqual(cm.exception.code,404)

    def test_upload_rejection(self):
        req=request.Request(self.base+'/api/upload',b'not a pptx',headers={'X-Filename':'not.pptx'},method='POST')
        with self.assertRaises(error.HTTPError) as cm:request.urlopen(req)
        self.assertEqual(cm.exception.code,400)

    def test_static_directory_escape_is_rejected(self):
        try:request.urlopen(self.base+'/../../server.py')
        except error.HTTPError as e:self.assertEqual(e.code,404)
        else:self.fail('Private source path should not be served')

if __name__=='__main__':unittest.main()
