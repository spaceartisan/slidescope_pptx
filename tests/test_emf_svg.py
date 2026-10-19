"""Pure-stdlib EMF vector translation must not depend on Windows GDI."""
import struct
import unittest
import sys
from pathlib import Path
sys.path.insert(0,str(Path(__file__).resolve().parents[1]))
from metafile import emf_to_svg


def emf_fixture(unsupported=False):
    def rec(code, payload=b''):
        payload += b'\0' * (-len(payload) % 4)
        return struct.pack('<II', code, 8+len(payload))+payload
    records=[
        rec(9,struct.pack('<ii',100,40)),  # logical size
        rec(39,struct.pack('<IIII',1,0,0x0011B2EE,0)),  # solid orange BGR
        rec(37,struct.pack('<I',1)),
        rec(59),
        rec(27,struct.pack('<ii',10,10)),
        rec(54,struct.pack('<ii',90,10)),
        rec(54,struct.pack('<ii',90,30)),
        rec(54,struct.pack('<ii',10,30)),
        rec(61),rec(60),
        rec(62,struct.pack('<4i',0,0,100,40)),
    ]
    if unsupported: records.append(rec(81))  # STRETCHDIBITS: no silent omission
    records.append(rec(14, b'\0'*12))
    hdr=bytearray(108)
    struct.pack_into('<II',hdr,0,1,108)
    struct.pack_into('<4i',hdr,24,0,0,100,40)
    hdr[40:44]=b' EMF'
    struct.pack_into('<I',hdr,48,108+sum(map(len,records)))
    return bytes(hdr)+b''.join(records)


class CrossPlatformEMFTests(unittest.TestCase):
    def test_vector_fill_without_windows_dependencies(self):
        svg=emf_to_svg(emf_fixture())
        self.assertIsNotNone(svg)
        self.assertIn(b'<svg ',svg)
        self.assertIn(b'viewBox="0 0 100 40"',svg)
        self.assertIn(b'fill="#EEB211"',svg)
        self.assertIn(b'L90,30',svg)
        self.assertIn(b' Z',svg)

    def test_unsupported_drawing_command_fails_closed(self):
        self.assertIsNone(emf_to_svg(emf_fixture(unsupported=True)))

    def test_invalid_input_is_rejected(self):
        self.assertIsNone(emf_to_svg(b'not a metafile'))
        self.assertIsNone(emf_to_svg(emf_fixture()[:88]))

if __name__=='__main__': unittest.main()
