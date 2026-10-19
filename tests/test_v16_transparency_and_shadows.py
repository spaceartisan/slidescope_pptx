"""Regression checks for inherited effects and SmartArt transparency."""
import os
from pathlib import Path
import unittest
from lxml import etree
from pptx_parser import parse_pptx, smartart_soft_fill

class EffectsTest(unittest.TestCase):
    def test_cached_soft_material_preserves_authored_alpha(self):
        root=etree.fromstring('''<a:spPr xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
            <a:solidFill><a:schemeClr val="accent1"><a:alpha val="90000"/>
            <a:tint val="40000"/></a:schemeClr></a:solidFill></a:spPr>''')
        theme={'accent1':'#EEB211','accent3':'#AAAAAA','accent4':'#DADADA','lt2':'#EECD74'}
        fill=smartart_soft_fill(root,theme,'#FFCC00','#002C56')
        self.assertRegex(fill,r'^rgba\(\d+, \d+, \d+, 0\.9000\)$')
        # A transparent material must not be converted to an opaque swatch.
        self.assertNotEqual(fill,'#E1D3C0')

    @unittest.skipUnless(os.environ.get('SLIDESCOPE_REGRESSION_PPTX'),
                         'Provide SLIDESCOPE_REGRESSION_PPTX to test the supplied deck')
    def test_real_slide_master_shadow_and_soft_body_panels(self):
        model,_=parse_pptx(Path(os.environ['SLIDESCOPE_REGRESSION_PPTX']).read_bytes())
        title=next(obj for obj in model['slides'][1]['shapes']
                   if obj['name']=='Title 5' and obj.get('text'))
        shadow=title['text']['shadow']
        self.assertEqual(shadow['color'],'#000000')
        self.assertAlmostEqual(shadow['x'],1.333,places=2)
        self.assertAlmostEqual(shadow['y'],1.333,places=2)
        panels=[obj for obj in model['slides'][6]['shapes']
                if obj['source']=='smartart' and obj['h']>300]
        self.assertEqual(len(panels),3)
        self.assertTrue(all(p['fill'].endswith(', 0.9000)') for p in panels))
        headers=[obj for obj in model['slides'][6]['shapes']
                 if obj['source']=='smartart' and obj['h']<100]
        self.assertEqual(len(headers),3)
        self.assertTrue(all(h['fill'].startswith('#') for h in headers))
