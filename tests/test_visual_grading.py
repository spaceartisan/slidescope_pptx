"""Small synthetic grading tests, no private deck required."""
import io
import sys
import tempfile
import unittest
from pathlib import Path
from zipfile import ZipFile
from PIL import Image

sys.path.insert(0,str(Path(__file__).resolve().parents[1]/'tools'))
from compare_versions import compare, mae


class VisualGradingTests(unittest.TestCase):
    def test_rgb_mae_and_html_report(self):
        with tempfile.TemporaryDirectory() as temp:
            root=Path(temp)
            old=root/'old';new=root/'new';old.mkdir();new.mkdir()
            ref=Image.new('RGB',(40,20),(100,100,100))
            ref.save(old/'Slide1.png')
            Image.new('RGB',(40,20),(105,105,105)).save(new/'Slide1.png')
            # The previous score is perfect and the new score regresses.
            stream=io.BytesIO();ref.save(stream,format='PNG')
            with ZipFile(root/'slides.zip','w') as z:
                z.writestr('exports/Slide1.PNG',stream.getvalue())
            rows=compare(root/'slides.zip',old,new,root/'report')
            self.assertEqual(rows[0][0],1)
            self.assertAlmostEqual(rows[0][1],0)
            self.assertAlmostEqual(rows[0][2],5)
            self.assertTrue((root/'report/index.html').exists())
            self.assertTrue((root/'report/Slide1.jpg').exists())
            self.assertEqual(len((root/'report/regions.csv').read_text().splitlines()),10)
            (new/'Slide1.png').unlink()
            with self.assertRaisesRegex(ValueError, 'partial deck'):
                compare(root/'slides.zip',old,new,root/'incomplete-report')
    def test_mismatched_size_is_rejected(self):
        with self.assertRaises(ValueError):
            mae(Image.new('RGB',(2,2)),Image.new('RGB',(3,2)))
