"""Independent cases for legacy previews and DrawingML pixel substitution."""
from io import BytesIO
from types import SimpleNamespace
import struct
import unittest
from lxml import etree
from PIL import Image
from pptx_parser import image_color_changes, ole_preview
from wmf_text import wmf_text_to_svg

A='http://schemas.openxmlformats.org/drawingml/2006/main'
P='http://schemas.openxmlformats.org/presentationml/2006/main'


def record(op,data=b''):
    if len(data)%2:data+=b'\0'
    return struct.pack('<IH',3+len(data)//2,op)+data


def sample_wmf(extra=b'',symbol=False,advances=False):
    name=b'Symbol' if symbol else b'Times New Roman'
    font=struct.pack('<5h8B',-100,0,0,0,400,1,0,0,2 if symbol else 0,0,0,0,0)+name+b'\0'
    text=b'D=' if symbol else b'ab'
    dx=struct.pack('<hh',70,40) if advances else b''
    records=(record(0x20b,struct.pack('<hh',0,0))+record(0x20c,struct.pack('<hh',200,500))+
             record(0x102,struct.pack('<H',1))+record(0x12e,struct.pack('<H',24))+
             record(0x2fb,font)+record(0x12d,struct.pack('<H',0))+
             record(0xa32,struct.pack('<hhHH',100,20,len(text),0)+text+dx)+extra+record(0))
    return struct.pack('<HHHIHIH',1,9,0x300,(18+len(records))//2,2,100,0)+records


class WMFTextTests(unittest.TestCase):
    def test_text_font_baseline_and_window_are_preserved(self):
        svg=etree.fromstring(wmf_text_to_svg(sample_wmf()))
        self.assertEqual(svg.get('viewBox'),'0 0 500 200')
        text=svg[0]
        self.assertEqual(text.text,'ab');self.assertEqual(text.get('y'),'100')
        self.assertEqual(text.get('font-size'),'100')
        self.assertEqual(text.get('font-style'),'italic')

    def test_symbol_encoding_and_character_advances(self):
        svg=etree.fromstring(wmf_text_to_svg(sample_wmf(symbol=True,advances=True)))
        self.assertEqual(svg[0].text,'Δ=')
        self.assertEqual(svg[0].get('x'),'20 90')

    def test_unsupported_drawing_and_truncated_files_are_rejected(self):
        self.assertIsNone(wmf_text_to_svg(sample_wmf(extra=record(0x041b,b'\0'*8))))
        self.assertIsNone(wmf_text_to_svg(sample_wmf()[:-2]))
        self.assertIsNone(wmf_text_to_svg(b'not a metafile'))

    def test_escape_is_not_executed_or_silently_accepted(self):
        self.assertIsNone(wmf_text_to_svg(sample_wmf(extra=record(0x626,struct.pack('<HH',1,0)))))


class ColorChangeTests(unittest.TestCase):
    def converted(self,use_alpha=None):
        source=Image.new('RGBA',(3,1))
        source.putdata([(240,240,240,255),(239,240,240,255),(240,240,240,128)])
        data=BytesIO();source.save(data,format='PNG')
        attr='' if use_alpha is None else f' useA="{int(use_alpha)}"'
        blip=etree.fromstring(f'''<a:blip xmlns:a="{A}"><a:clrChange{attr}>
          <a:clrFrom><a:srgbClr val="F0F0F0"/></a:clrFrom>
          <a:clrTo><a:srgbClr val="123456"><a:alpha val="0"/></a:srgbClr></a:clrTo>
        </a:clrChange></a:blip>''')
        warnings=[]
        output,ctype=image_color_changes(data.getvalue(),'image/png',blip,{},warnings)
        self.assertEqual(ctype,'image/png');self.assertEqual(warnings,[])
        return Image.open(BytesIO(output))

    def test_exact_color_and_alpha_match(self):
        im=self.converted()
        self.assertEqual(im.getpixel((0,0)),(18,52,86,0))
        self.assertEqual(im.getpixel((1,0)),(239,240,240,255))
        self.assertEqual(im.getpixel((2,0)),(240,240,240,128))

    def test_use_alpha_false_preserves_source_alpha(self):
        im=self.converted(False)
        self.assertEqual(im.getpixel((0,0)),(18,52,86,255))
        self.assertEqual(im.getpixel((2,0)),(18,52,86,128))

    def test_no_effect_preserves_original_bytes_and_type(self):
        raw=b'original'
        self.assertEqual(image_color_changes(raw,'image/jpeg',None,{},[]),(raw,'image/jpeg'))


class EmbeddedPreviewTests(unittest.TestCase):
    def test_vml_shape_id_and_part_relationship_resolve_preview(self):
        image=SimpleNamespace(blob=sample_wmf(),content_type='image/x-wmf')
        vml=SimpleNamespace(blob=b'''<xml xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office">
          <v:shape id="unrelated"><v:imagedata o:relid="wrong"/></v:shape>
          <v:shape id="target"><v:imagedata o:relid="preview" gain="0" blacklevel="32768f"/></v:shape>
        </xml>''',related_part=lambda rid:image if rid=='preview' else None)
        rel=SimpleNamespace(reltype='test/vmlDrawing',is_external=False,target_part=vml)
        # There is intentionally no embedded-object relationship accessor;
        # resolving this preview must not read or run the embedded object.
        slide=SimpleNamespace(part=SimpleNamespace(rels={'vml':rel}))
        shape=etree.fromstring(f'''<p:graphicFrame xmlns:p="{P}" xmlns:a="{A}">
          <a:graphic><a:graphicData><p:oleObj spid="target" name="Equation" progId="Equation.3"/></a:graphicData></a:graphic>
        </p:graphicFrame>''')
        assets={};warnings=[]
        result=ole_preview(shape,slide,'slide',assets,warnings)
        self.assertEqual(warnings,[])
        self.assertEqual(result['vmlAdjustment'],{'gain':0,'blackLevel':.5})
        self.assertEqual(result['embeddedObject']['program'],'Equation.3')
        self.assertEqual(len(assets),1)
        asset=next(iter(assets.values()))
        self.assertEqual(asset['type'],'image/svg+xml')
        self.assertIn(b'>ab</text>',asset['bytes'])


if __name__=='__main__':unittest.main()
