"""Offline archive and cached SmartArt regression tests (no test deck required)."""
import io
import sys
import unittest
from pathlib import Path
from zipfile import ZipFile, ZIP_DEFLATED
from types import SimpleNamespace
from lxml import etree

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0,str(ROOT))
import server
from pptx_parser import diagram_shapes

PNG_HEAD = b'\x89PNG\r\n\x1a\n'

def make_zip(files):
    dst=io.BytesIO()
    with ZipFile(dst,'w',ZIP_DEFLATED) as z:
        for name,content in files.items():z.writestr(name,content)
    return dst.getvalue()

class ReferenceTests(unittest.TestCase):
    def test_accepts_powerpoint_export_folder_and_slide_order(self):
        state=server.DeckState()
        content=make_zip({'Name/Slide2.PNG':PNG_HEAD+b'2','Name/Slide1.PNG':PNG_HEAD+b'1',
                          'Name/notes.txt':b'ignore'})
        self.assertEqual(state.load_references(content,'reference.zip'),2)
        self.assertEqual(sorted(state.references),[1,2])
        self.assertEqual(state.reference_name,'reference.zip')
    def test_rejects_bad_archive_without_erasing_previous_references(self):
        state=server.DeckState()
        state.load_references(make_zip({'Slide1.PNG':PNG_HEAD+b'1'}),'original.zip')
        for value in (b'notzip',make_zip({'Slide1.PNG':b'not png'}),make_zip({'other.png':PNG_HEAD})):
            with self.assertRaises(ValueError):state.load_references(value,'bad.zip')
        self.assertEqual(state.reference_name,'original.zip')
    def test_rejects_oversized_expanded_zip(self):
        state=server.DeckState()
        data=make_zip({'Slide1.PNG':PNG_HEAD+b'0'*(21*1024*1024)})
        with self.assertRaises(ValueError):state.load_references(data,'huge.zip')

class SmartArtTests(unittest.TestCase):
    def test_cached_shapes_are_inspectable_and_keep_geometry_and_text(self):
        ns='http://schemas.microsoft.com/office/drawing/2008/diagram'
        payload=f'''<dsp:drawing xmlns:dsp="{ns}" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><dsp:spTree><dsp:sp>
          <dsp:spPr><a:xfrm><a:off x="9525" y="19050"/><a:ext cx="952500" cy="476250"/></a:xfrm><a:prstGeom prst="rect"/><a:solidFill><a:srgbClr val="FFCC00"/></a:solidFill></dsp:spPr>
          <dsp:txBody><a:bodyPr anchor="ctr"/><a:p><a:r><a:rPr sz="2000" b="1"/><a:t>Safety review</a:t></a:r></a:p></dsp:txBody>
          </dsp:sp></dsp:spTree></dsp:drawing>'''
        dm_uri='http://schemas.openxmlformats.org/officeDocument/2006/relationships'
        fr=f'''<p:graphicFrame xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram" xmlns:r="{dm_uri}">
        <a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/diagram"><dgm:relIds r:dm="rId2"/></a:graphicData></a:graphic></p:graphicFrame>'''
        frame=SimpleNamespace(_element=etree.fromstring(fr.encode()),left=952500,top=952500,shape_id=7,name='Diagram 7')
        rels={'rId2':SimpleNamespace(target_part=SimpleNamespace(partname='/ppt/diagrams/data1.xml')),
              'rId6':SimpleNamespace(reltype='http://schemas.microsoft.com/office/2007/relationships/diagramDrawing',target_part=SimpleNamespace(partname='/ppt/diagrams/drawing1.xml',blob=payload.encode()))}
        # Non-diagram relationship still needs reltype for traversal.
        rels['rId2'].reltype='http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramData'
        slide=SimpleNamespace(part=SimpleNamespace(rels=rels))
        result=diagram_shapes(frame,slide,{'lt1':'#FFFFFF','accent1':'#FFC000'},{'minor':'Arial'},[])
        self.assertEqual(len(result),1)
        shape=result[0]
        self.assertEqual((shape['x'],shape['y'],shape['w'],shape['h']),(101.0,102.0,100.0,50.0))
        self.assertEqual(shape['fill'],'#FFCC00')
        self.assertEqual(shape['text']['paragraphs'][0]['runs'][0]['text'],'Safety review')
        self.assertTrue(shape['text']['paragraphs'][0]['runs'][0]['bold'])
    def test_unrelated_graphic_is_not_misclassified_as_smartart(self):
        frame=SimpleNamespace(_element=etree.fromstring(b'<a:root xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"/>'))
        self.assertIsNone(diagram_shapes(frame,None,{}, {}, []))

if __name__=='__main__':unittest.main()
