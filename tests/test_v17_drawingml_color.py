"""PowerPoint color-transform and cached diagram material regression coverage."""
from pathlib import Path
import os
import unittest
from lxml import etree
from pptx_parser import color_node, parse_pptx

A = 'http://schemas.openxmlformats.org/drawingml/2006/main'

def color(xml, theme=None):
    return color_node(etree.fromstring(xml), theme or {})

class DrawingMLColorTests(unittest.TestCase):
    def test_hsl_modifiers_are_not_applied_as_rgb_multipliers(self):
        node=f'<a:srgbClr xmlns:a="{A}" val="FF0000"><a:hueOff val="7200000"/></a:srgbClr>'
        self.assertEqual(color(node), '#00FF00')
        node=f'<a:srgbClr xmlns:a="{A}" val="FF0000"><a:satMod val="0"/></a:srgbClr>'
        self.assertEqual(color(node), '#808080')

    def test_alpha_composition_and_tint(self):
        node=(f'<a:schemeClr xmlns:a="{A}" val="accent1">'
              '<a:alpha val="80000"/><a:alphaMod val="50000"/>'
              '<a:alphaOff val="10000"/></a:schemeClr>')
        self.assertEqual(color(node,{'accent1':'#EEB211'}),'rgba(238, 178, 17, 0.5000)')
        node=f'<a:srgbClr xmlns:a="{A}" val="000000"><a:tint val="100000"/></a:srgbClr>'
        self.assertEqual(color(node),'#000000')

    @unittest.skipUnless(os.environ.get('SLIDESCOPE_REGRESSION_PPTX'),
                         'Set SLIDESCOPE_REGRESSION_PPTX for a full-deck check')
    def test_cached_smartart_soft_outline_matches_panel_material(self):
        model,_=parse_pptx(Path(os.environ['SLIDESCOPE_REGRESSION_PPTX']).read_bytes())
        cards=[obj for obj in model['slides'][6]['shapes']
               if obj['source']=='smartart' and obj['h']>300]
        self.assertEqual(len(cards),3)
        for card in cards:
            self.assertAlmostEqual(card['line']['width'],2.667, places=2)
            self.assertTrue(card['line']['color'].startswith('rgba('))
            self.assertIn('0.9000',card['line']['color'])
            self.assertTrue(card['fill'].startswith('rgba('))

if __name__ == '__main__':
    unittest.main()
