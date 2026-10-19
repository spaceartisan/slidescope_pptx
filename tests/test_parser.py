"""Python stdlib unittest suite; no browser automation dependency."""
import io
import sys
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))
from pptx_parser import parse_pptx, validate

SAMPLE = (ROOT / 'sample' / 'demo.pptx').read_bytes()

class ParserTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.deck, cls.assets = parse_pptx(SAMPLE, 'demo.pptx')

    def test_dimensions_and_slide_count(self):
        self.assertEqual(self.deck['count'],4)
        self.assertAlmostEqual(self.deck['width'], 1280, delta=2)
        self.assertAlmostEqual(self.deck['height'], 720, delta=1)

    def test_scene_graph_has_geometry(self):
        for slide in self.deck['slides']:
            self.assertGreater(len(slide['shapes']), 2)
            for shape in slide['shapes']:
                self.assertGreaterEqual(shape['w'], 0)
                self.assertGreaterEqual(shape['h'], 0)
                self.assertIn('id', shape)
                self.assertIn('kind', shape)
                self.assertIn('fill', shape)
                self.assertIn('source', shape)

    def test_text_content_and_color(self):
        shapes = self.deck['slides'][0]['shapes']
        blocks = [s['text'] for s in shapes if s['text'] is not None]
        words = ' '.join(p['text'] for block in blocks for p in block['paragraphs'])
        self.assertIn('Make every slide', words)
        self.assertIn('inspectable.', words)
        colors = [r['color'] for block in blocks for p in block['paragraphs'] for r in p['runs']]
        self.assertIn('#57CCBA', colors)

    def test_table_chart_image(self):
        s3 = self.deck['slides'][2]
        tables = [s for s in s3['shapes'] if s['kind']=='table']
        charts = [s for s in s3['shapes'] if s['kind']=='chart']
        self.assertEqual(len(tables), 1)
        self.assertEqual(len(charts), 1)
        self.assertEqual(tables[0]['table']['cells'][1][0]['text']['paragraphs'][0]['text'], 'Text boxes')
        self.assertEqual(charts[0]['chart']['categories'], ['Q1','Q2','Q3','Q4'])
        s4=self.deck['slides'][3]
        image=next(s for s in s4['shapes'] if s['kind']=='image')
        key=image['image']['url'].split('/')[-1]
        self.assertIn(key, self.assets)
        self.assertTrue(self.assets[key]['type'].startswith('image/'))

    def test_invalid_files_rejected(self):
        with self.assertRaisesRegex(ValueError, 'ZIP'):
            parse_pptx(b'wrong bytes', 'x.pptx')
        with self.assertRaisesRegex(ValueError, 'PowerPoint'):
            import zipfile
            bf=io.BytesIO()
            with zipfile.ZipFile(bf, 'w') as archive:
                archive.writestr('hello.txt','not a presentation')
            parse_pptx(bf.getvalue(),'x.pptx')

    def test_warnings_are_available(self):
        self.assertIsInstance(self.deck['warnings'],list)
        self.assertIsInstance(self.deck['slides'][0]['theme'], dict)

if __name__=='__main__':unittest.main()
