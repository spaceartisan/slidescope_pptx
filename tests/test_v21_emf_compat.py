"""Portable EMF fallback regression tests with synthetic, local-only fixtures."""
import struct
import unittest
from lxml import etree
from metafile_compat import emf_to_svg_compat


def rec(kind, body=b''):
    size=(8+len(body)+3)//4*4
    return struct.pack('<II',kind,size)+body+b'\0'*(size-8-len(body))


def make_emf(*records):
    data=bytearray(108)
    struct.pack_into('<II',data,0,1,108)
    struct.pack_into('<4i',data,8,0,0,400,300)
    struct.pack_into('<4i',data,24,0,0,10000,7500)
    data[40:44]=b' EMF'
    body=bytes(data)+b''.join(records)+rec(14,b'\0'*12)
    data=bytearray(body)
    struct.pack_into('<I',data,48,len(data))
    return bytes(data)


class PortableEMFTests(unittest.TestCase):
    def test_rejects_invalid_and_oversized_headers(self):
        self.assertIsNone(emf_to_svg_compat(b'rubbish'))
        self.assertIsNone(emf_to_svg_compat(b'\0'*(16*1024*1024+1)))

    def test_fills_text_and_skips_noop_bitblt(self):
        brush=rec(39,struct.pack('<4I',5,0,0x0000FF00,0))
        select=rec(37,struct.pack('<I',5))
        rectangle=rec(43,struct.pack('<4i',20,20,380,200))
        txt='Portable!'.encode('utf-16le')
        # EMR_EXTTEXTOUTW offsets are relative to the record start.
        text_body=bytearray(68+len(txt))
        struct.pack_into('<ii',text_body,28,45,65)
        struct.pack_into('<I',text_body,36,len(txt)//2)
        struct.pack_into('<I',text_body,40,76)
        text_body[68:68+len(txt)]=txt
        text=rec(84,text_body)
        # A trailing DSTCOPY is a no-op. Previously it covered an entire table
        # in white and hid all earlier drawing commands.
        blt=bytearray(92)
        struct.pack_into('<5I',blt,16,0,0,400,300,0x00AA0029)
        svg=emf_to_svg_compat(make_emf(brush,select,rectangle,text,rec(76,blt)))
        self.assertIsNotNone(svg)
        root=etree.fromstring(svg)
        self.assertIn(b'Portable!',svg)
        self.assertIn(b'#00ff00',svg)
        self.assertEqual(len(root.xpath('//*[local-name()="rect"]')),1)

    def test_clip_region_and_shapes_are_svg_elements(self):
        # A copied rectangular GDI region must not be thrown away.
        hdr=struct.pack('<4I4i',32,1,1,16,10,10,200,200)
        rect=struct.pack('<4i',10,10,200,200)
        region=rec(75,struct.pack('<II',len(hdr)+len(rect),5)+hdr+rect)
        ellipse=rec(42,struct.pack('<4i',25,25,100,100))
        svg=emf_to_svg_compat(make_emf(region,ellipse))
        self.assertIsNotNone(svg)
        self.assertIn(b'clipPath',svg)
        self.assertIn(b'<ellipse',svg)


if __name__=='__main__':unittest.main()
