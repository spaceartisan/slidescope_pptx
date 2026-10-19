"""Style cascade and connector decorations independent of reference decks."""
import unittest
from lxml import etree
from pptx_parser import shape_fill, line_style

A='http://schemas.openxmlformats.org/drawingml/2006/main'
P='http://schemas.openxmlformats.org/presentationml/2006/main'


class ShapeStyleTests(unittest.TestCase):
    def setUp(self):
        self.theme={'accent1':'#123456'}
        self.root=etree.fromstring(f'''<a:theme xmlns:a="{A}"><a:themeElements><a:fmtScheme>
          <a:fillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:fillStyleLst>
          <a:lnStyleLst><a:ln w="38100"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill>
          <a:prstDash val="dash"/><a:tailEnd type="triangle" w="lg" len="sm"/></a:ln></a:lnStyleLst>
        </a:fmtScheme></a:themeElements></a:theme>''')

    def shape(self,props='',index=1):
        return etree.fromstring(f'''<p:sp xmlns:p="{P}" xmlns:a="{A}"><p:spPr>{props}</p:spPr>
          <p:style><a:fillRef idx="{index}"><a:schemeClr val="accent1"/></a:fillRef>
          <a:lnRef idx="{index}"><a:srgbClr val="FEDCBA"/></a:lnRef></p:style></p:sp>''')

    def test_fill_and_line_use_independent_reference_colors(self):
        shape=self.shape()
        self.assertEqual(shape_fill(shape,self.theme,self.root),'#123456')
        line=line_style(shape,self.theme,self.root)
        self.assertEqual(line['color'],'#FEDCBA')
        self.assertEqual(line['width'],4)
        self.assertEqual(line['dash'],'dash')
        self.assertEqual(line['tailEnd'],{'type':'triangle','width':'lg','length':'sm'})

    def test_partial_direct_line_preserves_theme_color_and_dash(self):
        line=line_style(self.shape('<a:ln w="19050"><a:headEnd type="arrow"/><a:tailEnd type="none"/></a:ln>'),self.theme,self.root)
        self.assertEqual(line['width'],2)
        self.assertEqual(line['color'],'#FEDCBA')
        self.assertEqual(line['dash'],'dash')
        self.assertEqual(line['headEnd']['type'],'arrow')
        self.assertIsNone(line['tailEnd'])

    def test_direct_no_fill_and_no_line_override_style(self):
        shape=self.shape('<a:noFill/><a:ln><a:noFill/></a:ln>')
        self.assertEqual(shape_fill(shape,self.theme,self.root),'transparent')
        self.assertIsNone(line_style(shape,self.theme,self.root))

    def test_direct_colors_override_style(self):
        shape=self.shape('<a:solidFill><a:srgbClr val="112233"/></a:solidFill><a:ln><a:solidFill><a:srgbClr val="445566"/></a:solidFill></a:ln>')
        self.assertEqual(shape_fill(shape,self.theme,self.root),'#112233')
        self.assertEqual(line_style(shape,self.theme,self.root)['color'],'#445566')

    def test_zero_and_out_of_range_refs_do_not_invent_style(self):
        for index in (0,99):
            self.assertIsNone(shape_fill(self.shape(index=index),self.theme,self.root))
            self.assertIsNone(line_style(self.shape(index=index),self.theme,self.root))

    def test_picture_and_group_fills_do_not_use_style_solid(self):
        self.assertIsNone(shape_fill(self.shape('<a:blipFill/>'),self.theme,self.root))
        self.assertEqual(shape_fill(self.shape('<a:grpFill/>'),self.theme,self.root,'#ABCDEF'),'#ABCDEF')


if __name__=='__main__': unittest.main()
