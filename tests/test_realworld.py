"""Optional, reproducible regression checks for the real-world user test deck.

Set SLIDESCOPE_REGRESSION_PPTX to a local file path to activate these tests.
The user's private/source presentation is deliberately not included in the ZIP.
"""
import os
import unittest
from pathlib import Path

from pptx_parser import parse_pptx


class RealWorldDeckTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        candidate = os.environ.get('SLIDESCOPE_REGRESSION_PPTX')
        if not candidate:
            raise unittest.SkipTest('Set SLIDESCOPE_REGRESSION_PPTX to test with the real deck')
        cls.model, cls.assets = parse_pptx(Path(candidate).read_bytes(), Path(candidate).name)

    def test_title_slide_master_hidden_and_background_resolved(self):
        slide = self.model['slides'][0]
        self.assertEqual(self.model['count'], 12)
        self.assertFalse(any(o['source']=='master' for o in slide['shapes']))
        self.assertIn('/api/media/',slide['backgroundImage'])

    def test_title_placeholder_inherits_layout(self):
        slide=self.model['slides'][0]
        title=next(o for o in slide['shapes'] if o['name']=='Title 1' and o['source']=='slide')
        paragraph=title['text']['paragraphs'][0]
        self.assertEqual(paragraph['text'],'Causation v. Correlation')
        self.assertEqual(paragraph['align'],'center')
        self.assertEqual(paragraph['style']['color'],'#FFFFFF')
        self.assertAlmostEqual(paragraph['style']['size'],64.0)
        self.assertTrue(paragraph['style']['bold'])

    def test_content_slides_keep_master_gradient(self):
        for slide_number in (2,4):
            slide=self.model['slides'][slide_number-1]
            self.assertTrue(slide['background'].startswith('linear-gradient('))
            self.assertIn('#002C56',slide['background'])
            self.assertIsNone(slide['backgroundImage'])

    def test_master_bullets_are_inherited(self):
        slide=self.model['slides'][3]
        body=next(o for o in slide['shapes'] if o['source']=='slide'
                  and o['name']=='Content Placeholder 4')
        first, nested=body['text']['paragraphs'][:2]
        self.assertEqual(first['bullet'],'•')
        self.assertEqual(nested['bullet'],'–')
        self.assertEqual(first['bulletColor'],'#EEB211')
        self.assertEqual(nested['bulletColor'],'#EEB211')
        self.assertGreater(nested['indentLeft'],first['indentLeft'])
        self.assertLess(first['hanging'],0)

    def test_default_bold_only_applies_to_free_shapes(self):
        slide=self.model['slides'][1]
        body=next(o for o in slide['shapes'] if o['source']=='slide'
                  and o['name']=='Content Placeholder 4')
        callout=next(o for o in slide['shapes'] if o['source']=='slide'
                  and o['name']=='Rectangle 9')
        self.assertIsNone(body['text']['paragraphs'][0]['style']['bold'])
        self.assertTrue(callout['text']['paragraphs'][0]['style']['bold'])
        self.assertTrue(callout['text']['paragraphs'][0]['runs'][0]['bold'])

    def test_title_vertical_anchor_from_master(self):
        slide=self.model['slides'][1]
        title=next(o for o in slide['shapes'] if o['source']=='slide'
                   and o['name']=='Title 5')
        self.assertEqual(title['text']['vertical'],'middle')
        connector=next(o for o in slide['shapes'] if o['source']=='layout'
                       and o['kind']=='connector')
        self.assertEqual(connector['line']['color'],'#F0AF13')
        self.assertLess(connector['h'],1)

    def test_empty_paragraphs_are_preserved(self):
        slide=self.model['slides'][1]
        body=next(o for o in slide['shapes'] if o['source']=='slide'
                  and o['name']=='Content Placeholder 4')
        paragraphs=body['text']['paragraphs']
        self.assertEqual(len(paragraphs),5)
        self.assertEqual([i for i,p in enumerate(paragraphs) if not p['text']],[1,3])
        self.assertLess(body['text']['fontScale'],1.0)

    def test_free_textbox_resolves_embedded_paragraph_levels(self):
        slide=self.model['slides'][5]
        right=next(o for o in slide['shapes'] if o['source']=='slide'
                   and o['name']=='Content Placeholder 4' and not o['placeholderType'])
        first=right['text']['paragraphs'][0]
        nested=next(p for p in right['text']['paragraphs'] if p['level']==1)
        self.assertEqual(first['bullet'],'•')
        self.assertEqual(nested['bullet'],'–')
        self.assertEqual(first['bulletColor'],'#EEB211')
        self.assertAlmostEqual(first['runs'][0]['size'],26.4)
        self.assertGreater(nested['indentLeft'],first['indentLeft'])

    def test_cached_smartart_light_panels_and_paragraph_gaps(self):
        slide=self.model['slides'][6]
        panels=[o for o in slide['shapes'] if o['source']=='smartart'
                and o['h']>300]
        self.assertEqual(len(panels),3)
        # A soft panel must not become a saturated translucent yellow surface.
        self.assertTrue(all(p['fill'].startswith('rgba(') for p in panels))
        self.assertTrue(all(p['fill'].endswith(', 0.9000)') for p in panels))
        self.assertTrue(all(p['fill']!='#EEB211' for p in panels))
        self.assertTrue(all(any(para.get('spaceAfter')==26 for para in p['text']['paragraphs'])
                            for p in panels))

    def test_cached_smartart_color_close_to_neutral_material(self):
        panel = next(s for s in self.model['slides'][6]['shapes']
                     if s['source']=='smartart' and s['h']>300)
        fill=panel['fill']
        import re
        self.assertRegex(fill, r'^rgba\(\d+, \d+, \d+, 0\.9000\)$')
        values=[int(c) for c in re.findall(r'\d+',fill)[:3]]
        # The underlying source tint needs to be pale enough that its
        # translucency composites to neutral beige over a dark slide.
        self.assertTrue(all(190 <= channel <= 255 for channel in values))

    def test_master_text_insets_are_resolved(self):
        # The content and title placeholders have empty local a:bodyPr,
        # while the master stores 130622 EMU (13.714 px) side insets.
        # python-pptx otherwise returns its generic 9.6 px default.
        for n in (2, 3, 4, 5, 8, 10, 11, 12):
            slide=self.model['slides'][n-1]
            boxes=[s for s in slide['shapes'] if s['source']=='slide'
                   and s['kind']=='placeholder' and s.get('text')
                   and s['placeholderType'] in ('TITLE','OBJECT','BODY')]
            self.assertTrue(boxes, f'Missing placeholders on slide {n}')
            for obj in boxes:
                self.assertAlmostEqual(obj['text']['margin'][3],13.714,places=2)
                self.assertAlmostEqual(obj['text']['margin'][0],6.857,places=2)

    def test_numbered_text_preserves_hanging_indent_without_bullet(self):
        # Explicit a:buNone suppresses bullet glyph, not first-line indent.
        for n in (10, 11, 12):
            body=next(s for s in self.model['slides'][n-1]['shapes']
                      if s['source']=='slide' and s.get('text')
                      and any('.) ' in p['text'] for p in s['text']['paragraphs']))
            number=next(p for p in body['text']['paragraphs'] if '.) ' in p['text'])
            self.assertIsNone(number['bullet'])
            self.assertTrue(number['hanging'] is not None)
            if n in (10,11):self.assertLess(number['hanging'],0)

    def test_slide_number_uses_master_placeholder_alignment(self):
        # The master slide-number placeholder uses centered a:lvl1pPr,
        # overriding the presentation's generic left-aligned otherStyle.
        for n in (2, 4, 6, 7, 8, 10, 11, 12):
            num=next(s for s in self.model['slides'][n-1]['shapes']
                     if s['source']=='slide' and 'Slide Number' in s['name'])
            self.assertEqual(num['text']['paragraphs'][0]['align'], 'center')

    def test_autofit_factor_and_outline_survive_parsing(self):
        # Both dense text columns in slide 6 use normAutofit/fontScale.
        slide=self.model['slides'][5]
        frames=[obj for obj in slide['shapes'] if obj['source']=='slide'
                and obj.get('text') and len(obj['text']['paragraphs'])==6]
        self.assertEqual(len(frames),2)
        self.assertTrue(all(abs(obj['text']['fontScale']-.55)<.001 for obj in frames))
        callouts=[o for o in slide['shapes'] if o['source']=='slide'
                  and o.get('fill')=='#EEB211' and o.get('line')]
        self.assertGreaterEqual(len(callouts),2)
        self.assertTrue(all(o['line']['width']>=2 for o in callouts))

    def test_logo_is_not_rectangle_fallback(self):
        slide=self.model['slides'][0]
        paths=[o['paths'] for o in slide['shapes'] if o['source']=='layout' and o['customGeometry']]
        self.assertTrue(paths)
        self.assertTrue(all(p['supported'] for item in paths for p in item))


class MixedAutofitRunSizeTests(unittest.TestCase):
    """PowerPoint body placeholders can mix inherited and direct font sizes."""
    @classmethod
    def setUpClass(cls):
        source = os.environ.get('SLIDESCOPE_REGRESSION_PPTX')
        if not source:
            raise unittest.SkipTest('Set SLIDESCOPE_REGRESSION_PPTX for mixed font-size checks')
        cls.deck, _ = parse_pptx(Path(source).read_bytes(), Path(source).name)

    def test_unscaled_autofit_with_smaller_direct_runs(self):
        for slide_no in (4, 5):
            objects = self.deck['slides'][slide_no-1]['shapes']
            body = next(o for o in objects if o['source'] == 'slide' and
                        o.get('text') and len(o['text']['paragraphs']) >= 7)
            self.assertEqual(body['text']['autoSize'], 'TEXT_TO_FIT_SHAPE')
            self.assertAlmostEqual(body['text']['fontScale'], 1.0)
            p = body['text']['paragraphs'][0]
            self.assertLess(p['runs'][0]['size'], .85 * p['style']['size'])
            self.assertGreater(p['lineSpacingReduction'], 0)

    def test_other_body_text_keeps_existing_metrics(self):
        for slide_no in (2, 6):
            objects = self.deck['slides'][slide_no-1]['shapes']
            body = next(o for o in objects if o['source'] == 'slide' and
                        o.get('text') and len(o['text']['paragraphs']) >= 5)
            self.assertLess(body['text']['fontScale'], .999)


if __name__ == '__main__':
    unittest.main()
