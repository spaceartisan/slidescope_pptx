"""Compare browser screenshots with original PowerPoint PNG exports.

Usage: python tools/visual_diff.py references.zip screenshots_dir
Expected screenshots_dir/Slide1.png etc. Pillow is already a python-pptx dependency.
Outputs mean absolute error per RGB channel (0=exact match, 255=max mismatch).
"""
from __future__ import annotations
from pathlib import Path
from zipfile import ZipFile
from io import BytesIO
import argparse
import re
from PIL import Image, ImageChops, ImageStat


def scores(zip_path: Path, renders: Path, crop_bottom=0):
    results=[]
    with ZipFile(zip_path) as z:
        for filename in z.namelist():
            match = re.fullmatch(r'Slide(\d+)\.png',Path(filename).name,re.IGNORECASE)
            if not match:continue
            num=int(match.group(1))
            rendered=renders/f'Slide{num}.png'
            if not rendered.exists():continue
            with Image.open(BytesIO(z.read(filename))) as ref, Image.open(rendered) as out:
                ref=ref.convert('RGB');out=out.convert('RGB')
                if ref.size!=out.size:
                    raise ValueError(f'Slide {num}: reference {ref.size}, screenshot {out.size}; capture at original PPTX dimensions')
                if crop_bottom:
                    new_height=ref.height-crop_bottom
                    if new_height<=0:raise ValueError('Crop exceeds image height')
                    ref=ref.crop((0,0,ref.width,new_height));out=out.crop((0,0,out.width,new_height))
                delta=ImageChops.difference(ref,out)
                mae=sum(ImageStat.Stat(delta).mean)/3
                results.append((num,round(mae,3)))
    return sorted(results)


def main():
    p=argparse.ArgumentParser(description='Evaluate SlideScope PNG screenshots against PowerPoint exports')
    p.add_argument('reference_zip',type=Path)
    p.add_argument('screenshots',type=Path)
    p.add_argument('--crop-bottom',type=int,default=0,help='Ignore a footer of N pixels (for OS-dependent graphics, diagnostics only)')
    args=p.parse_args()
    out=scores(args.reference_zip,args.screenshots,args.crop_bottom)
    if not out:raise SystemExit('No SlideN.png screenshots matched the reference ZIP')
    for n,mae in out:print(f'Slide {n:>2}: RGB MAE {mae:>7.3f} / 255')
    print(f'Mean: {sum(x[1] for x in out)/len(out):.3f} / 255  ({len(out)} slides)')

if __name__=='__main__': main()
