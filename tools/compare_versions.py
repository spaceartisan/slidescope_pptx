"""Offline, image-based PowerPoint fidelity report, no extra dependencies.

Capture screenshots at the native reference resolution (e.g. 960x720 or 1536x864) with the existing
Playwright agent, then run:

  python tools/compare_versions.py slides.zip old_renders new_renders --output qa

The report uses the *actual* PowerPoint PNG exports, not scene graph estimates.
Requires Pillow, which is already an indirect python-pptx requirement.
"""
from __future__ import annotations
import argparse
import csv
from html import escape
from io import BytesIO
from pathlib import Path
from zipfile import ZipFile
import re
from PIL import Image, ImageChops, ImageEnhance, ImageStat, ImageDraw


def mae(left: Image.Image, right: Image.Image) -> float:
    if left.size != right.size:
        raise ValueError(f'Images have different dimensions: {left.size} vs {right.size}')
    return sum(ImageStat.Stat(ImageChops.difference(left, right)).mean) / 3


def compare(archive: Path, old_dir: Path, new_dir: Path, dest: Path, crop_bottom: int = 0):
    dest.mkdir(parents=True, exist_ok=True)
    rows=[]
    region_rows=[]
    with ZipFile(archive) as zip_file:
        entries=[]
        for entry in zip_file.namelist():
            match=re.fullmatch(r'Slide(\d+)\.png', Path(entry).name, re.I)
            if match:
                entries.append((int(match.group(1)),entry))
        for n,entry in sorted(entries):
            old_path=old_dir/f'Slide{n}.png'
            new_path=new_dir/f'Slide{n}.png'
            if not old_path.is_file() or not new_path.is_file():
                raise ValueError(f'Slide {n}: missing previous or current screenshot; cannot grade a partial deck')
            with Image.open(BytesIO(zip_file.read(entry))) as img:
                reference=img.convert('RGB')
            with Image.open(old_path) as img:
                old=img.convert('RGB')
            with Image.open(new_path) as img:
                new=img.convert('RGB')
            if not (reference.size == old.size == new.size):
                raise ValueError(f'Slide {n}: export and screenshots must have the same native dimensions')
            if crop_bottom:
                if crop_bottom >= reference.height:
                    raise ValueError('Crop exceeds slide height')
                region=(0,0,reference.width,reference.height-crop_bottom)
                a,b,c=[image.crop(region) for image in (reference,old,new)]
            else:
                a,b,c=reference,old,new
            former,current=mae(a,b),mae(a,c)
            rows.append((n,former,current,former-current))
            # A single full-slide average can mask large local failures. Keep
            # a normalized 3×3 grid so the QA agent can identify persistent
            # trouble spots without hand-picking slide-specific coordinates.
            for gy in range(3):
                for gx in range(3):
                    left,right=round(gx*a.width/3),round((gx+1)*a.width/3)
                    top,bottom=round(gy*a.height/3),round((gy+1)*a.height/3)
                    rectangle=(left,top,right,bottom)
                    old_reg=mae(a.crop(rectangle),b.crop(rectangle))
                    new_reg=mae(a.crop(rectangle),c.crop(rectangle))
                    region_rows.append((n,gx+1,gy+1,old_reg,new_reg,old_reg-new_reg))
            # Never crop the reference in the visual contact sheet: show the
            # unsupported footer, even when optional scoring excludes it.
            thumbnail_width=576
            tw=thumbnail_width
            th=round(reference.height*tw/reference.width)
            montage=Image.new('RGB',(tw*3,th+38),'#18212d')
            for idx,(image,title) in enumerate(((old,'PREVIOUS'),(new,'CURRENT'),(reference,'POWERPOINT REFERENCE'))):
                montage.paste(image.resize((tw,th),Image.Resampling.LANCZOS),(idx*tw,38))
                ImageDraw.Draw(montage).text((idx*tw+12,12),f'{title}  ·  Slide {n}',fill='#fff')
            montage.save(dest/f'Slide{n}.jpg',quality=88)
    if not rows:
        raise ValueError('No matching SlideN.png images in both screenshot directories')
    with (dest/'scores.csv').open('w',newline='',encoding='utf-8') as f:
        writer=csv.writer(f)
        writer.writerow(['slide','before_mae','after_mae','improvement'])
        for n,o,c,d in rows: writer.writerow([n,f'{o:.3f}',f'{c:.3f}',f'{d:+.3f}'])
    with (dest/'regions.csv').open('w',newline='',encoding='utf-8') as f:
        writer=csv.writer(f)
        writer.writerow(['slide','column','row','before_mae','after_mae','improvement'])
        for n,gx,gy,old,new,gain in region_rows:
            writer.writerow([n,gx,gy,f'{old:.3f}',f'{new:.3f}',f'{gain:+.3f}'])
    mean_before=sum(r[1] for r in rows)/len(rows)
    mean_after=sum(r[2] for r in rows)/len(rows)
    body='\n'.join(
        f'<article><header>Slide {n} &nbsp; {old:.2f} → {new:.2f} &nbsp; '
        f'<strong class="{"good" if delta >= 0 else "bad"}">{delta:+.2f}</strong></header>'
        f'<img alt="Slide {n}: previous, current, PowerPoint reference" src="Slide{n}.jpg"></article>'
        for n,old,new,delta in rows)
    page=f'''<!doctype html><html lang="en"><meta charset="utf-8"><meta name="viewport" content="width=device-width">
<title>SlideScope fidelity report</title><style>
:root{{color-scheme:dark;font-family:system-ui,Arial,sans-serif}}
body{{background:#131a22;color:#edf3fa;max-width:1800px;margin:30px auto;padding:0 24px}}
h1{{margin-bottom:0}}p{{color:#b0becd;line-height:1.6}}
article{{border:1px solid #344353;border-radius:8px;background:#1b2531;margin:18px 0;overflow:hidden}}
header{{padding:12px 18px;font-weight:600}}img{{width:100%;display:block}}
.good{{color:#73d9b3}}.bad{{color:#f0a196}}
</style><h1>SlideScope visual fidelity report</h1>
<p>PowerPoint reference: {escape(archive.name)}. Average RGB absolute error (0–255; lower is better).<br>
Before: <b>{mean_before:.3f}</b> &nbsp; After: <b>{mean_after:.3f}</b> &nbsp;
Change: <b>{mean_before-mean_after:+.3f}</b> across {len(rows)} slides.<br>
Scoring excluded bottom {crop_bottom}px; images show full slides. See <code>regions.csv</code> for per-slide 3×3 error maps, including untouched regions. Scores are sensitive to platform fonts,
anti-aliasing, and EMF support.</p>{body}</html>'''
    (dest/'index.html').write_text(page,encoding='utf-8')
    return rows


if __name__=='__main__':
    ap=argparse.ArgumentParser(description=__doc__)
    ap.add_argument('reference_zip',type=Path)
    ap.add_argument('previous_screenshots',type=Path)
    ap.add_argument('current_screenshots',type=Path)
    ap.add_argument('--output',type=Path,default=Path('qa-report'))
    ap.add_argument('--crop-bottom',type=int,default=0)
    args=ap.parse_args()
    results=compare(args.reference_zip,args.previous_screenshots,args.current_screenshots,args.output,args.crop_bottom)
    for n,old,new,delta in results:
        print(f'Slide {n:>2}: {old:>7.3f} -> {new:>7.3f} ({delta:+.3f})')
    print(f'Open {args.output / "index.html"}')
