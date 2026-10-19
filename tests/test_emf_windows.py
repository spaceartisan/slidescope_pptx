"""EMF regression coverage without Windows, PowerPoint, or extra libraries.

Exercise the Windows GDI call signature with an emulated DLL, and ensure a
failed conversion is recoverable when a real PPTX object requests it.
"""
import ctypes
from io import BytesIO
from pathlib import Path
import struct
import sys
import unittest
from unittest.mock import patch
from zipfile import ZipFile, ZIP_DEFLATED

from PIL import Image  # import before temporarily simulating Windows platform

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))
import metafile
from pptx_parser import parse_pptx


def valid_header_only_emf():
    """Minimal, non-proprietary EMF document suitable for XML/media tests."""
    header = bytearray(88)
    fields = [
        (0, '<II', (1, 88)),
        (8, '<4i', (0, 0, 100, 100)),
        (24, '<4i', (0, 0, 2540, 2540)),
        (40, '<I', (0x464D4520,)),
        (44, '<IIII', (0x10000, 108, 2, 1)),
        (60, '<III', (0, 0, 0)),
        (72, '<2i', (100, 100)),
        (80, '<2i', (25, 25)),
    ]
    for offset, fmt, values in fields:
        struct.pack_into(fmt, header, offset, *values)
    return bytes(header) + struct.pack('<IIIII', 14, 20, 0, 0, 20)


def sample_with_emf():
    original = (ROOT / 'sample' / 'demo.pptx').read_bytes()
    buffer = BytesIO()
    with ZipFile(BytesIO(original)) as src, ZipFile(buffer, 'w', ZIP_DEFLATED) as dst:
        for member in src.infolist():
            name = member.filename
            content = src.read(name)
            if name == 'ppt/media/image1.png':
                name = 'ppt/media/image1.emf'
                content = valid_header_only_emf()
            elif name == '[Content_Types].xml':
                content = content.replace(
                    b'<Default Extension="png" ContentType="image/png"/>',
                    b'<Default Extension="emf" ContentType="image/x-emf"/>',
                )
            elif name == 'ppt/slides/_rels/slide4.xml.rels':
                content = content.replace(b'image1.png', b'image1.emf')
            dst.writestr(name, content)
    return buffer.getvalue()


class IntentionalStop(Exception):
    pass


class APIStub:
    def __init__(self, name):
        self.name = name
        self.argtypes = None
        self.restype = None

    def __call__(self, *args):
        # GDI is deliberately not invoked: signature inspection is cross-platform.
        if self.argtypes is None or len(args) != len(self.argtypes):
            raise AssertionError(f'{self.name} called with invalid GDI signature')
        raise IntentionalStop('All GDI signatures have now been bound')


class GDIStub:
    def __init__(self):
        self.functions = {}

    def __getattr__(self, name):
        return self.functions.setdefault(name, APIStub(name))


class WindowsEmfTests(unittest.TestCase):
    def test_windows_gdi_signatures_match_calls(self):
        gdi = GDIStub()
        with patch.object(metafile.os, 'name', 'nt'):
            with patch.object(ctypes, 'WinDLL', return_value=gdi, create=True):
                with self.assertRaises(IntentionalStop):
                    metafile.emf_to_png(valid_header_only_emf(), 120, 100)
        self.assertEqual(len(gdi.functions['CreateDIBSection'].argtypes), 6)
        self.assertEqual(len(gdi.functions['SetEnhMetaFileBits'].argtypes), 2)
        self.assertEqual(len(gdi.functions['PlayEnhMetaFile'].argtypes), 3)

    def test_failed_emf_conversion_does_not_abort_presentation(self):
        with patch('pptx_parser.emf_to_png', side_effect=RuntimeError('fake GDI error')) as converter:
            deck, _ = parse_pptx(sample_with_emf(), 'emf-demo.pptx')
        self.assertEqual(deck['count'], 4)
        self.assertEqual(converter.call_count, 1)
        self.assertTrue(any('EMF preview failed' in warning for warning in deck['warnings']))
        image = next(item for item in deck['slides'][3]['shapes'] if item['kind'] == 'image')
        self.assertIsNone(image['image']['url'])


if __name__ == '__main__':
    unittest.main()
