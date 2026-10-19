"""PowerPoint -> browser scene graph. Python 3.12; python-pptx + lxml only.

Purposefully a visual *preview*, not an OOXML-compliant reference renderer.
"""
from __future__ import annotations

from io import BytesIO
from copy import deepcopy
import hashlib
import colorsys
import math
import re
from typing import Any
from zipfile import ZipFile, BadZipFile
from PIL import Image, ImageColor, ImageChops

from lxml import etree
from pptx import Presentation
from pptx.enum.dml import MSO_COLOR_TYPE
from pptx.enum.shapes import MSO_SHAPE_TYPE
from metafile import emf_to_png, emf_to_svg
from metafile_compat import emf_to_svg_compat
from wmf_text import wmf_text_to_svg

EMU_PER_PX = 914400 / 96
NS = {"a": "http://schemas.openxmlformats.org/drawingml/2006/main",
      "p": "http://schemas.openxmlformats.org/presentationml/2006/main"}
NS_A = NS["a"]
PREVIEW_LIMITS = {"bytes": 64 * 1024 * 1024, "uncompressed": 350 * 1024 * 1024,
                  "entries": 20000, "slides": 400}
DEFAULT_THEME = {"dk1":"#000000", "lt1":"#FFFFFF", "dk2":"#44546A", "lt2":"#E7E6E6",
                 "accent1":"#4472C4", "accent2":"#ED7D31", "accent3":"#A5A5A5", "accent4":"#FFC000",
                 "accent5":"#5B9BD5", "accent6":"#70AD47", "hlink":"#0563C1", "folHlink":"#954F72"}


def px(value: Any) -> float:
    return round((int(value or 0)) / EMU_PER_PX, 3)


def one(parent, path):
    return parent.find(path, namespaces=NS) if parent is not None else None


def many(parent, path):
    return parent.findall(path, namespaces=NS) if parent is not None else []


def number_attr(el, attr, default=0):
    try:
        return float(el.get(attr, default)) if el is not None else default
    except (ValueError, TypeError):
        return default


def theme_for(slide):
    theme = dict(DEFAULT_THEME)
    font = {"major": "Aptos Display", "minor": "Aptos"}
    try:
        master_part = slide.slide_layout.slide_master.part
        for rel in master_part.rels.values():
            if rel.reltype.endswith('/theme'):
                root = etree.fromstring(rel.target_part.blob)
                scheme = one(root, './/a:clrScheme')
                if scheme is not None:
                    for child in scheme:
                        if not isinstance(child.tag, str):
                            continue
                        name = etree.QName(child).localname
                        value = child[0] if len(child) else None
                        if value is not None:
                            if etree.QName(value).localname == 'sysClr':
                                v = value.get('lastClr') or value.get('val')
                            else:
                                v = value.get('val')
                            if v and re.fullmatch(r'[0-9A-Fa-f]{6}', v):
                                theme[name] = '#' + v.upper()
                for key, path in [('major', './/a:majorFont/a:latin'), ('minor', './/a:minorFont/a:latin')]:
                    elem = one(root, path)
                    if elem is not None and elem.get('typeface'):
                        font[key] = elem.get('typeface')
                break
    except Exception:
        pass
    # OOXML schemeClr aliases (tx1/bg1/...) are mapped by the slide master's
    # p:clrMap, not by fixed black/white assumptions.
    try:
        mapping = one(slide.slide_layout.slide_master._element, 'p:clrMap')
        override = one(slide.slide_layout._element, 'p:clrMapOvr/a:overrideClrMapping')
        slide_override = one(slide._element, 'p:clrMapOvr/a:overrideClrMapping')
        mapping = slide_override or override or mapping
        if mapping is not None:
            for alias, target in mapping.attrib.items():
                if target in theme: theme[alias] = theme[target]
    except (AttributeError, ValueError):
        pass
    return theme, font


def color_node(node, theme, default=None):
    if node is None:
        return default
    local = etree.QName(node).localname
    val = node.get('val', '')
    if local == 'srgbClr' and re.fullmatch(r'[0-9a-fA-F]{6}', val):
        result = '#' + val.upper()
    elif local == 'sysClr':
        val = node.get('lastClr') or val
        result = '#' + val.upper() if re.fullmatch(r'[0-9a-fA-F]{6}', val) else default
    elif local == 'schemeClr':
        result = theme.get(val, default)
    elif local == 'prstClr':
        result = {'black': '#000000', 'white': '#FFFFFF', 'red':'#FF0000',
                  'blue':'#0000FF', 'gray':'#808080', 'yellow':'#FFFF00'}.get(val, default)
    else:
        return default
    if not result:
        return default
    try:
        rgb = [int(result[n:n+2], 16) for n in (1,3,5)]
        alpha = 1.0
        for mod in node:
            key = etree.QName(mod).localname
            raw = number_attr(mod, 'val', 100000)
            v = raw / 100000
            if key in ('hueOff', 'hueMod', 'satOff', 'satMod', 'lumOff', 'lumMod'):
                # DrawingML adjusts hue, saturation and luminance in HSL
                # space, rather than scaling the encoded RGB channels.
                h, light, sat = colorsys.rgb_to_hls(*(c / 255 for c in rgb))
                if key == 'hueOff': h = (h + raw / (360 * 60000)) % 1
                elif key == 'hueMod': h = (h * v) % 1
                elif key == 'satOff': sat = max(0., min(1., sat + v))
                elif key == 'satMod': sat = max(0., min(1., sat * v))
                elif key == 'lumOff': light = max(0., min(1., light + v))
                elif key == 'lumMod': light = max(0., min(1., light * v))
                rgb = [round(c*255) for c in colorsys.hls_to_rgb(h,light,sat)]
            # DrawingML tint is the retained fraction of the input color.
            # 100% keeps the color; 0% produces white (unlike SpreadsheetML).
            elif key == 'tint':
                # Office blends the retained color with white in linear RGB,
                # then encodes it for display. Blending encoded bytes makes
                # pale table bands conspicuously too dark and saturated.
                def tinted(channel):
                    encoded = channel / 255
                    linear = encoded / 12.92 if encoded <= .04045 else ((encoded+.055)/1.055)**2.4
                    linear = max(0., min(1., linear*v + 1-v))
                    encoded = linear*12.92 if linear <= .0031308 else 1.055*linear**(1/2.4)-.055
                    return round(encoded*255)
                rgb = [tinted(c) for c in rgb]
            elif key == 'shade': rgb = [round(c*v) for c in rgb]
            elif key == 'alpha': alpha = max(0., min(1., v))
            elif key == 'alphaMod': alpha = max(0., min(1., alpha*v))
            elif key == 'alphaOff': alpha = max(0., min(1., alpha+v))
        rgb = [max(0,min(255,c)) for c in rgb]
        if alpha < 1:
            return f'rgba({rgb[0]}, {rgb[1]}, {rgb[2]}, {alpha:.4f})'
        return '#%02X%02X%02X' % tuple(rgb)
    except (ValueError, TypeError):
        return result


def fill_style(parent, theme, default=None):
    if parent is None:
        return default
    if one(parent, 'a:noFill') is not None:
        return 'transparent'
    solid = one(parent, 'a:solidFill')
    if solid is not None and len(solid):
        return color_node(solid[0], theme, default)
    grad = one(parent, 'a:gradFill')
    if grad is not None:
        stops = []
        for gs in many(grad, 'a:gsLst/a:gs'):
            c = color_node(gs[0], theme, '#CCCCCC') if len(gs) else '#CCCCCC'
            stops.append(f'{c} {number_attr(gs, "pos")/1000:.2f}%')
        if len(stops) >= 2:
            lin = one(grad, 'a:lin')
            angle = number_attr(lin,'ang',0) / 60000
            # OOXML angle is clockwise from horizontal; CSS from upwards clockwise.
            return f'linear-gradient({(90+angle)%360:.2f}deg, {", ".join(stops)})'
    return default


def image_luminance(blip):
    """Return native DrawingML image luminance instructions for the browser.

    a:lum applies to the bitmap itself, rather than a rectangular fill.
    Keep values separate so renderers can handle them independently.
    """
    lum = one(blip, 'a:lum')
    if lum is None:
        return None
    return {'brightness':max(-1.0,min(1.0,number_attr(lum,'bright',0)/100000)),
            'contrast':max(-1.0,min(1.0,number_attr(lum,'contrast',0)/100000))}


def image_color_changes(raw, content_type, blip, theme, warnings):
    """Apply exact DrawingML color substitutions before browser luminance."""
    changes=many(blip,'a:clrChange')
    if not changes:return raw,content_type
    try:
        with Image.open(BytesIO(raw)) as source:
            if getattr(source,'n_frames',1)>1:
                warnings.append('Color replacement on animated images is unsupported')
                return raw,content_type
            raster=source.convert('RGBA')
        for change in changes:
            fr=one(change,'a:clrFrom');to=one(change,'a:clrTo')
            if fr is None or to is None or not len(fr) or not len(to):continue
            def rgba(node):
                value=color_node(node,theme)
                if value and value.startswith('rgba('):
                    channels=[float(c.strip()) for c in value[5:-1].split(',')]
                    return tuple(round(c) for c in channels[:3])+(round(channels[3]*255),)
                return ImageColor.getcolor(value,'RGBA')
            before=rgba(fr[0]);after=rgba(to[0])
            use_alpha=change.get('useA','1') not in ('0','false')
            mask=Image.new('L',raster.size,255)
            for channel,value in zip(raster.split()[:4 if use_alpha else 3],before):
                mask=ImageChops.multiply(mask,channel.point([255 if i==value else 0 for i in range(256)]))
            replacement=Image.new('RGBA',raster.size,after)
            if not use_alpha:replacement.putalpha(raster.getchannel('A'))
            raster.paste(replacement,(0,0),mask)
        stream=BytesIO();raster.save(stream,format='PNG')
        return stream.getvalue(),'image/png'
    except (OSError,ValueError,TypeError) as exc:
        warnings.append(f'Image color replacement unsupported: {type(exc).__name__}')
        return raw,content_type


def ole_preview(shape_el, slide, source, assets, warnings):
    """Resolve the stored preview only, never the embedded OLE program/data."""
    ole=one(shape_el,'a:graphic/a:graphicData/p:oleObj')
    if ole is None or slide is None:return None
    owner=slide.part if source=='slide' else (slide.slide_layout.part if source=='layout' else slide.slide_layout.slide_master.part)
    rid_ns='{http://schemas.openxmlformats.org/officeDocument/2006/relationships}'
    metadata={'program':ole.get('progId'),'name':ole.get('name'),'shapeId':ole.get('spid')}
    try:
        part=None;adjustment=None
        blip=one(ole,'p:pic/p:blipFill/a:blip')
        if blip is not None and blip.get(rid_ns+'embed'):
            part=owner.related_part(blip.get(rid_ns+'embed'))
        if part is None:
            for rel in owner.rels.values():
                if not rel.reltype.endswith('/vmlDrawing') or rel.is_external:continue
                vml_part=rel.target_part;root=etree.fromstring(vml_part.blob)
                matches=root.findall('.//{urn:schemas-microsoft-com:vml}shape')
                match=next((s for s in matches if s.get('id')==ole.get('spid')),None)
                data=match.find('{urn:schemas-microsoft-com:vml}imagedata') if match is not None else None
                if data is None:continue
                rid=data.get('{urn:schemas-microsoft-com:office:office}relid') or data.get(rid_ns+'id')
                if not rid:continue
                part=vml_part.related_part(rid)
                def vml_value(value,default):
                    if value is None:return default
                    return float(value[:-1])/65536 if value.endswith('f') else float(value)
                adjustment={'gain':vml_value(data.get('gain'),1),'blackLevel':vml_value(data.get('blacklevel'),0)}
                break
        if part is None:
            warnings.append(f'Embedded object has no stored preview: {metadata["name"]}')
            return None
        raw=part.blob;ctype=part.content_type
        if ctype not in ('image/png','image/jpeg','image/gif','image/webp','image/svg+xml'):
            converted=emf_to_svg(raw) if raw[:4]==b'\x01\x00\x00\x00' else wmf_text_to_svg(raw)
            if not converted and raw[:4]==b'\x01\x00\x00\x00':converted=emf_to_svg_compat(raw)
            if not converted:
                warnings.append(f'Unsupported embedded-object preview: {ctype} ({metadata["name"]})')
                return None
            raw,ctype=converted,'image/svg+xml'
        digest=hashlib.sha1(raw).hexdigest()[:18];assets.setdefault(digest,{'bytes':raw,'type':ctype})
        return {'url':f'/api/media/{digest}','crop':[0,0,0,0],'vmlAdjustment':adjustment,'embeddedObject':metadata}
    except (KeyError,ValueError,AttributeError,etree.XMLSyntaxError) as exc:
        warnings.append(f'Cannot read embedded-object preview: {type(exc).__name__}')
        return None


def image_fill(shape_el, slide, source, assets, warnings, theme=None):
    """Resolve an image embedded as a fill of an ordinary p:sp shape.

    These are NOT p:pic shapes: many real PPTX decks use picture-filled
    rectangles for portraits and diagrams, which python-pptx exposes as
    AUTO_SHAPE. Their media relationships belong to the slide/layout/master.
    """
    fill = one(shape_el, 'p:spPr/a:blipFill')
    blip = one(fill, 'a:blip')
    if blip is None or slide is None:
        return None
    rid=blip.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed')
    if not rid:
        return None
    try:
        owner = (slide.part if source=='slide' else
                 slide.slide_layout.part if source=='layout' else
                 slide.slide_layout.slide_master.part)
        part=owner.related_part(rid)
        raw=part.blob
        ctype=part.content_type
        if ctype not in ('image/png','image/jpeg','image/gif','image/webp','image/svg+xml'):
            warnings.append(f'Unsupported bitmap fill format: {ctype}')
            return None
        raw,ctype=image_color_changes(raw,ctype,blip,theme or {},warnings)
        digest=hashlib.sha1(raw).hexdigest()[:18]
        if digest not in assets:
            assets[digest]={'bytes':raw,'type':ctype}
        rect=one(fill,'a:srcRect')
        crop=[max(0.0,min(.999,number_attr(rect,key,0)/100000)) for key in ('l','t','r','b')]
        return {'url':f'/api/media/{digest}','crop':crop,'luminance':image_luminance(blip)}
    except (KeyError,AttributeError,ValueError) as exc:
        warnings.append(f'Cannot read image shape fill: {type(exc).__name__}: {exc}')
        return None


def shape_fill(shape_el, theme, theme_root=None, group_fill=None):
    """Direct formatting overrides the theme's style-matrix fill."""
    sppr = one(shape_el, 'p:spPr')
    if one(sppr, 'a:grpFill') is not None:
        return group_fill
    # A bitmap/pattern fill must not fall through to a themed solid fill.
    if sppr is not None and any(etree.QName(c).localname.endswith('Fill') for c in sppr):
        return fill_style(sppr, theme)
    native, mapped = _table_reference(one(shape_el, 'p:style/a:fillRef'), theme_root, theme, 'fill')
    if native is None: return None
    wrapper = etree.Element('fill')
    wrapper.append(etree.fromstring(etree.tostring(native)))
    return fill_style(wrapper, mapped)


def line_style(shape_el, theme, theme_root=None):
    ln = one(shape_el, 'p:spPr/a:ln')
    if ln is None: ln = one(shape_el, 'p:grpSpPr/a:ln')
    base, mapped = _table_reference(one(shape_el, 'p:style/a:lnRef'), theme_root, theme, 'line')
    if ln is None and base is None:
        return None
    color = fill_style(ln, theme, fill_style(base, mapped, '#606979'))
    if color == 'transparent':
        return None
    width = number_attr(ln, 'w', number_attr(base, 'w', 12700)) / EMU_PER_PX
    def child(name):
        value = one(ln, 'a:'+name)
        return value if value is not None else one(base, 'a:'+name)
    def end(name):
        value = child(name)
        if value is None or value.get('type', 'none') == 'none': return None
        return {'type':value.get('type'), 'width':value.get('w', 'med'), 'length':value.get('len', 'med')}
    dash = child('prstDash')
    return {'color': color or '#606979', 'width': round(max(width, 0.5),2),
            'dash': dash.get('val') if dash is not None else None,
            'headEnd':end('headEnd'), 'tailEnd':end('tailEnd')}


def font_color(font, theme):
    try:
        c = font.color
        if c.type == MSO_COLOR_TYPE.RGB and c.rgb:
            return '#' + str(c.rgb)
        if c.type == MSO_COLOR_TYPE.SCHEME and c.theme_color:
            name = c.theme_color.name.lower()
            aliases = {'dark_1':'dk1', 'light_1':'lt1', 'dark_2':'dk2', 'light_2':'lt2',
                       'text_1':'dk1', 'background_1':'lt1', 'text_2':'dk2', 'background_2':'lt2',
                       'hyperlink':'hlink','followed_hyperlink':'folHlink'}
            return theme.get(aliases.get(name,name))
    except (ValueError, AttributeError):
        pass
    return None


def _inherited_text_defaults(shape, slide, theme, theme_fonts, level):
    """Resolve placeholder paragraph defaults from slide master and layout.

    Bullets, indent, and spacing generally live in the *paragraph* style,
    not the runs. Resolving font style alone loses entire presentation layouts.
    """
    result = {'font': theme_fonts['minor'], 'size': None, 'color': None,
              'bold': None, 'italic': None, 'align': None,
              'bullet': None, 'autoBullet': False, 'numbering': None, 'bulletColor': None, 'bulletFont': None,
              'marL': None, 'indent': None, 'spaceBeforePct': None,
              'spaceAfterPct': None, 'bulletSizePct': None}
    if not slide:
        return result
    # An ordinary text box can carry a complete style hierarchy in txBody/
    # lstStyle, independently of the master placeholder hierarchy. It is
    # especially common when a placeholder has been copied or converted into
    # a normal text box. Resolve this after presentation/master defaults.
    local_ppr = one(shape._element,
                    f'p:txBody/a:lstStyle/a:lvl{min(max(level+1,1),9)}pPr')
    def merge_local_style():
        if local_ppr is None:
            return
        ppr=local_ppr
        alignment=ppr.get('algn')
        if alignment:
            result['align']={'ctr':'center','l':'left','r':'right','just':'justify'}.get(alignment,'left')
        for key in ('marL','indent'):
            if ppr.get(key) is not None: result[key]=px(int(ppr.get(key)))
        for name,path in [('spaceBeforePct','a:spcBef/a:spcPct'),
                          ('spaceAfterPct','a:spcAft/a:spcPct'),
                          ('bulletSizePct','a:buSzPct')]:
            el=one(ppr,path)
            if el is not None:result[name]=number_attr(el,'val')/100000
        char=one(ppr,'a:buChar');num=one(ppr,'a:buAutoNum')
        if char is not None:
            result['bullet']=char.get('char','•');result['autoBullet']=False;result['numbering']=None
        elif num is not None:
            result['bullet']=None;result['autoBullet']=True
            result['numbering']={'type':num.get('type','arabicPeriod'),'startAt':int(num.get('startAt',1))}
        elif one(ppr,'a:buNone') is not None:
            result['bullet']=None;result['autoBullet']=False;result['numbering']=None
        bc=one(ppr,'a:buClr')
        if bc is not None and len(bc):result['bulletColor']=color_node(bc[0],theme,result['bulletColor'])
        bf=one(ppr,'a:buFont')
        if bf is not None:result['bulletFont']=bf.get('typeface')
        if one(ppr,'a:buFontTx') is not None:result['bulletFont']=None
        if one(ppr,'a:buClrTx') is not None:result['bulletColor']=None
        defr=one(ppr,'a:defRPr')
        if defr is None:return
        if defr.get('sz'):result['size']=round(number_attr(defr,'sz')/100*96/72,3)
        for attr,key in [('b','bold'),('i','italic')]:
            if defr.get(attr) is not None:result[key]=defr.get(attr) in ('1','true')
        latin=one(defr,'a:latin')
        if latin is not None and latin.get('typeface'):
            fn=latin.get('typeface')
            result['font']=theme_fonts['major' if fn.startswith('+mj-') else 'minor'] if fn.startswith(('+mj-','+mn-')) else fn
        fc=fill_style(defr,theme)
        if fc and fc!='transparent':result['color']=fc
    if not getattr(shape, 'is_placeholder', False):
        # Presentation defaults style free text on auto-shapes; placeholder
        # body styles are a separate inheritance chain and must NOT get the
        # generic defaultTextStyle bold flag from this presentation.
        defaults=one(slide.part.package.presentation_part._element,
                     f'p:defaultTextStyle/a:lvl{min(max(level+1,1),9)}pPr')
        if defaults is not None:
            defr=one(defaults,'a:defRPr')
            if defr is not None:
                if defr.get('sz'):result['size']=round(number_attr(defr,'sz')/100*96/72,3)
                for prop,key in [('b','bold'),('i','italic')]:
                    if defr.get(prop) is not None:result[key]=defr.get(prop) in ('1','true')
                lf=one(defr,'a:latin')
                if lf is not None and lf.get('typeface'):
                    fn=lf.get('typeface')
                    result['font']=theme_fonts['major' if fn.startswith('+mj-') else 'minor'] if fn.startswith(('+mj-','+mn-')) else fn
                fc=fill_style(defr,theme)
                if fc and fc!='transparent':result['color']=fc
        merge_local_style()
        return result
    from pptx.enum.shapes import PP_PLACEHOLDER
    ph_type = shape.placeholder_format.type
    group = 'titleStyle' if ph_type in (PP_PLACEHOLDER.TITLE, PP_PLACEHOLDER.CENTER_TITLE) else (
        'bodyStyle' if ph_type in (PP_PLACEHOLDER.BODY, PP_PLACEHOLDER.OBJECT) else 'otherStyle')
    master = slide.slide_layout.slide_master
    level_path = f'a:lvl{min(max(level+1,1),9)}pPr'
    master_ppr = one(master._element, f'p:txStyles/p:{group}/{level_path}')
    layout_shape = None
    try:
        layout_shape = next((p for p in slide.slide_layout.placeholders
                             if p.placeholder_format.idx == shape.placeholder_format.idx), None)
    except (KeyError, AttributeError):
        pass
    layout_ppr = one(layout_shape._element, f'p:txBody/a:lstStyle/{level_path}') if layout_shape is not None else None
    # Most inherited objects use bodyStyle even if their placeholder index
    # has no corresponding layout placeholder (e.g. legacy .ppt conversions).
    for ppr in (master_ppr, layout_ppr):
        if ppr is None: continue
        alignment = ppr.get('algn')
        if alignment:
            result['align'] = {'ctr':'center','l':'left','r':'right','just':'justify'}.get(alignment,'left')
        for key in ('marL','indent'):
            if ppr.get(key) is not None:
                result[key] = px(int(ppr.get(key)))
        for name,element in [('spaceBeforePct','a:spcBef/a:spcPct'),
                             ('spaceAfterPct','a:spcAft/a:spcPct'),
                             ('bulletSizePct','a:buSzPct')]:
            val=one(ppr, element)
            if val is not None:
                result[name] = number_attr(val,'val')/100000
        bullet=one(ppr,'a:buChar')
        numbered=one(ppr,'a:buAutoNum')
        if bullet is not None:
            result['bullet']=bullet.get('char','•');result['autoBullet']=False;result['numbering']=None
        elif numbered is not None:
            result['bullet']=None;result['autoBullet']=True
            result['numbering']={'type':numbered.get('type','arabicPeriod'),'startAt':int(numbered.get('startAt',1))}
        if one(ppr,'a:buNone') is not None:
            result['bullet']=None;result['autoBullet']=False;result['numbering']=None
        bcolor=one(ppr,'a:buClr')
        bf=one(ppr,'a:buFont')
        if bf is not None:result['bulletFont']=bf.get('typeface')
        if one(ppr,'a:buFontTx') is not None:result['bulletFont']=None
        if one(ppr,'a:buClrTx') is not None:result['bulletColor']=None
        if bcolor is not None:
            bc=color_node(bcolor[0],theme) if len(bcolor) else None
            if bc: result['bulletColor']=bc
        defr=one(ppr,'a:defRPr')
        if defr is None:continue
        if defr.get('sz'):
            result['size']=round(number_attr(defr,'sz')/100*96/72,3)
        for prop,key in [('b','bold'),('i','italic')]:
            if defr.get(prop) is not None:
                result[key] = defr.get(prop) in ('1','true')
        latin=one(defr,'a:latin')
        if latin is not None:
            fn=latin.get('typeface','')
            result['font']=theme_fonts['major' if fn.startswith('+mj-') else 'minor'] if fn.startswith(('+mj-','+mn-')) else (fn or result['font'])
        fill=fill_style(defr,theme)
        if fill and fill!='transparent':result['color']=fill
    # Slide-number and footer placeholders often have their own paragraph
    # alignment in the master's *placeholder* lstStyle. The generic
    # p:otherStyle is left aligned and must not override that more-specific
    # placeholder formatting. A missing slide-number alignment otherwise
    # shifts the field to the left throughout a presentation.
    if ph_type == PP_PLACEHOLDER.SLIDE_NUMBER:
        master_ph = next((ph for ph in master.placeholders
                          if ph.placeholder_format.type == ph_type), None)
        master_ph_style = one(master_ph._element,
                              f'p:txBody/a:lstStyle/{level_path}') if master_ph is not None else None
        if master_ph_style is not None and master_ph_style.get('algn'):
            result['align'] = {'ctr':'center','l':'left','r':'right',
                               'just':'justify'}.get(master_ph_style.get('algn'),result['align'])
    merge_local_style()
    return result


def custom_svg_paths(sppr):
    """Translate simple DrawingML custom paths into object-local SVG coordinates."""
    geom=one(sppr,'a:custGeom')
    if geom is None:return []
    result=[]
    for path in many(geom,'a:pathLst/a:path'):
        width,height=max(1,number_attr(path,'w',100)),max(1,number_attr(path,'h',100))
        def xy(point):
            return f'{number_attr(point,"x")/width*100:.4f} {number_attr(point,"y")/height*100:.4f}'
        chunks=[]
        supported=True
        for step in path:
            typ=etree.QName(step).localname
            coords=many(step,'a:pt')
            if typ=='moveTo' and coords:chunks.append('M'+xy(coords[0]))
            elif typ=='lnTo' and coords:chunks.append('L'+xy(coords[0]))
            elif typ=='cubicBezTo' and len(coords)==3:chunks.append('C'+' '.join(xy(p) for p in coords))
            elif typ=='quadBezTo' and len(coords)==2:chunks.append('Q'+' '.join(xy(p) for p in coords))
            elif typ=='close':chunks.append('Z')
            else:supported=False
        if chunks:result.append({'d':' '.join(chunks),'fill':path.get('fill','norm'),'supported':supported})
    return result


def background_for(slide, theme, assets=None, warnings=None):
    """Resolve actual slide/layout/master backgrounds, including image blips."""
    for obj in [slide, slide.slide_layout,slide.slide_layout.slide_master]:
        try:
            bg = one(obj._element, 'p:cSld/p:bg/p:bgPr')
            if bg is None:continue
            blip = one(bg, 'a:blipFill/a:blip')
            if blip is not None and assets is not None:
                rel=blip.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed')
                if rel:
                    part=obj.part.related_part(rel)
                    digest=hashlib.sha1(part.blob).hexdigest()[:18]
                    assets.setdefault(digest,{'bytes':part.blob,'type':part.content_type})
                    result={'color':'#FFFFFF','image':f'/api/media/{digest}',
                            'imageTile':None}
                    tile=one(bg,'a:blipFill/a:tile')
                    if tile is not None and part.content_type.startswith('image/'):
                        # DrawingML tiled background images are scaled from
                        # their *physical* dimensions, not stretched to fill
                        # the slide. A 72-DPI JPEG scaled to 55% is rendered
                        # at (pixels * 96 / 72 * .55) CSS pixels.
                        try:
                            with Image.open(BytesIO(part.blob)) as raster:
                                iw,ih=raster.size
                                dx,dy=raster.info.get('dpi',(96,96))
                            dx=dx if dx and dx>0 else 96
                            dy=dy if dy and dy>0 else 96
                            sx=number_attr(tile,'sx',100000)/100000
                            sy=number_attr(tile,'sy',100000)/100000
                            result['imageTile']={
                                'w':round(iw*96/dx*sx,3),
                                'h':round(ih*96/dy*sy,3),
                                'x':round(px(number_attr(tile,'tx',0)),3),
                                'y':round(px(number_attr(tile,'ty',0)),3)}
                        except (OSError,ValueError,TypeError,ZeroDivisionError):
                            pass
                    return result
            fill=fill_style(bg,theme)
            if fill:return {'color':fill,'image':None}
        except (AttributeError, KeyError, ValueError):
            pass
    return {'color':'#FFFFFF','image':None}


def inherited_text_margins(shape, slide, body_pr, defaults):
    """DrawingML txBody/bodyPr insets inherit from layout/master placeholders.

    python-pptx exposes its generic 0.1/0.05 inch defaults when the local
    a:bodyPr has no explicit lIns/tIns/rIns/bIns. That loses the master body's
    typically larger inset and shifts the text in many PPTX decks.
    """
    attrs=('tIns','rIns','bIns','lIns')
    margins=list(defaults)
    needs={attr for attr in attrs if body_pr is None or body_pr.get(attr) is None}
    if not needs or slide is None or not getattr(shape,'is_placeholder',False):
        return margins
    from pptx.enum.shapes import PP_PLACEHOLDER
    typ=shape.placeholder_format.type
    try:
        layout_ph=next((p for p in slide.slide_layout.placeholders
                        if p.placeholder_format.idx==shape.placeholder_format.idx),None)
        if layout_ph is None:
            layout_ph=next((p for p in slide.slide_layout.placeholders
                            if p.placeholder_format.type==typ),None)
    except (ValueError,AttributeError):
        layout_ph=None
    master_types=(PP_PLACEHOLDER.BODY,PP_PLACEHOLDER.OBJECT) if typ in (PP_PLACEHOLDER.BODY,PP_PLACEHOLDER.OBJECT) else (
        (PP_PLACEHOLDER.TITLE,PP_PLACEHOLDER.CENTER_TITLE) if typ in (PP_PLACEHOLDER.TITLE,PP_PLACEHOLDER.CENTER_TITLE) else (typ,))
    master_ph=next((p for p in slide.slide_layout.slide_master.placeholders
                    if p.placeholder_format.type in master_types),None)
    for ancestor in (layout_ph,master_ph):
        if ancestor is None:continue
        body=one(ancestor._element,'p:txBody/a:bodyPr')
        if body is None:continue
        for j,attr in enumerate(attrs):
            if attr in needs and body.get(attr) is not None:
                margins[j]=px(int(body.get(attr)))
                needs.remove(attr)
        if not needs:break
    return margins


def outer_text_shadow(shape, slide, theme):
    """Resolve an explicit DrawingML outer shadow through placeholder ancestry.

    A title often has no local p:spPr effects even though its master shape
    specifies a crisp, offset shadow. CSS text-shadow can reproduce those
    Office DrawingML vector effects without rasterizing text or adding fonts.
    """
    candidates = []
    if slide is not None and getattr(shape, 'is_placeholder', False):
        ph_type = shape.placeholder_format.type
        idx = shape.placeholder_format.idx
        master = slide.slide_layout.slide_master
        from pptx.enum.shapes import PP_PLACEHOLDER
        master_type = PP_PLACEHOLDER.TITLE if ph_type == PP_PLACEHOLDER.CENTER_TITLE else ph_type
        match = next((p for p in master.placeholders
                      if p.placeholder_format.type == master_type), None)
        if match is not None: candidates.append(match._element)
        layout = next((p for p in slide.slide_layout.placeholders
                       if p.placeholder_format.idx == idx), None)
        if layout is not None: candidates.append(layout._element)
    candidates.append(shape._element)
    result = None
    for candidate in candidates:
        effects = one(candidate, 'p:spPr/a:effectLst')
        if effects is None: continue
        outer = one(effects, 'a:outerShdw')
        if outer is None:
            result = None  # an explicit effect list can clear inherited effects
            continue
        color = color_node(outer[0], theme, '#000000') if len(outer) else '#000000'
        distance = number_attr(outer, 'dist', 0) / EMU_PER_PX
        direction = math.radians(number_attr(outer, 'dir', 0) / 60000)
        result = {'x': round(distance*math.cos(direction), 3),
                  'y': round(distance*math.sin(direction), 3),
                  'blur': round(number_attr(outer,'blurRad',0)/EMU_PER_PX,3),
                  'color': color}
    return result


def parse_text(shape, theme, theme_fonts, slide=None, text_defaults=None):
    tf = shape.text_frame
    body_pr=one(shape._element,'p:txBody/a:bodyPr') if hasattr(shape,'_element') else one(shape._tc,'a:txBody/a:bodyPr')
    anchor=body_pr.get('anchor') if body_pr is not None else None
    if anchor is None and slide is not None and getattr(shape,'is_placeholder',False):
        try:
            placeholders=[p for p in slide.slide_layout.placeholders
                          if p.placeholder_format.idx==shape.placeholder_format.idx]
            if not placeholders:
                placeholders=[p for p in slide.slide_layout.placeholders
                              if p.placeholder_format.type==shape.placeholder_format.type]
            masters=[p for p in slide.slide_layout.slide_master.placeholders
                     if p.placeholder_format.type==shape.placeholder_format.type]
            # For a center title, the master uses TITLE, not CENTER_TITLE.
            if not masters and shape.placeholder_format.type.name=='CENTER_TITLE':
                masters=[p for p in slide.slide_layout.slide_master.placeholders
                         if p.placeholder_format.type.name=='TITLE']
            for ph in placeholders+masters:
                bp=one(ph._element,'p:txBody/a:bodyPr')
                if bp is not None and bp.get('anchor') is not None:
                    anchor=bp.get('anchor');break
        except (AttributeError,ValueError):pass
    normal=one(body_pr,'a:normAutofit')
    font_scale=number_attr(normal,'fontScale',100000)/100000 if normal is not None else 1.0
    font_scale=max(.1,min(2,font_scale))
    spacing_scale=max(.3,min(2,number_attr(normal,'lnSpcReduction',0)/100000)) if normal is not None else 0
    result = {'margin': inherited_text_margins(shape,slide,body_pr,
                         [px(tf.margin_top),px(tf.margin_right),px(tf.margin_bottom),px(tf.margin_left)]),
              'vertical': ({'ctr':'middle','b':'bottom','t':'top'}.get(anchor,'top')),
              'direction':body_pr.get('vert','horz') if body_pr is not None else 'horz',
              'textRotation':number_attr(body_pr,'rot',0)/60000,
              'wrap': bool(tf.word_wrap) if tf.word_wrap is not None else True,
              'autoSize': tf.auto_size.name if tf.auto_size is not None else None,
              'fontScale':font_scale,
              'shadow':outer_text_shadow(shape,slide,theme) if hasattr(shape,'_element') else None,
              'paragraphs': []}
    def effective_size(value):
        return round(value*font_scale,3) if value is not None else None
    for para in tf.paragraphs:
        inherited = _inherited_text_defaults(shape,slide,theme,theme_fonts,para.level)
        if text_defaults:
            inherited.update(text_defaults)
        ppr = one(para._p,'a:pPr')
        bullet=one(ppr,'a:buChar')
        auto_element=one(ppr,'a:buAutoNum')
        auto_bullet=auto_element is not None
        bullet_char=bullet.get('char') if bullet is not None else inherited['bullet']
        numbering=inherited['numbering']
        if auto_element is not None:
            bullet_char=None
            numbering={'type':auto_element.get('type','arabicPeriod'),
                       'startAt':int(auto_element.get('startAt',1))}
        elif bullet is not None:
            auto_bullet=False
            numbering=None
        else:
            auto_bullet=inherited['autoBullet']
        if one(ppr,'a:buNone') is not None:
            bullet_char,auto_bullet,numbering=None,False,None
        indent_left = px(int(ppr.get('marL'))) if ppr is not None and ppr.get('marL') else inherited['marL']
        hanging = px(int(ppr.get('indent'))) if ppr is not None and ppr.get('indent') else inherited['indent']
        bullet_color=inherited['bulletColor']
        bc=one(ppr,'a:buClr')
        if bc is not None and len(bc): bullet_color=color_node(bc[0],theme,bullet_color)
        if one(ppr,'a:buClrTx') is not None: bullet_color=None
        bf=one(ppr,'a:buFont')
        bullet_font=bf.get('typeface') if bf is not None else inherited['bulletFont']
        if one(ppr,'a:buFontTx') is not None: bullet_font=None
        base=para.font
        style={'font': base.name or inherited['font'],
               'size': effective_size(round(base.size.pt*96/72,3) if base.size is not None else inherited['size']),
               'color':font_color(base,theme) or inherited['color'],
               'bold':base.bold if base.bold is not None else inherited['bold'],
               'italic':base.italic if base.italic is not None else inherited['italic']}
        # Even an empty DrawingML paragraph has endParaRPr. PowerPoint uses
        # its size to lay out the blank line; the master default can be much
        # larger and otherwise accumulates severe vertical drift.
        if not para.text:
            end_rpr=one(para._p,'a:endParaRPr')
            if end_rpr is not None and end_rpr.get('sz'):
                style['size']=effective_size(number_attr(end_rpr,'sz')/100*96/72)
        runs=[]
        all_runs=[]
        for run in para.runs:
            ft=run.font
            rpr=one(run._r,'a:rPr')
            direct_col=fill_style(rpr,theme) if rpr is not None else None
            if direct_col=='transparent':direct_col=None
            font_name=ft.name or style['font']
            if font_name and font_name.startswith(('+mj-', '+mn-')):
                font_name=theme_fonts['major' if font_name.startswith('+mj-') else 'minor']
            rstyle={'font':font_name,
                    'size':effective_size(round(ft.size.pt*96/72,3) if ft.size is not None else
                                          (round(base.size.pt*96/72,3) if base.size is not None else inherited['size'])),
                    'color':direct_col or font_color(ft,theme) or style['color'],
                    'bold':ft.bold if ft.bold is not None else style['bold'],
                    'italic':ft.italic if ft.italic is not None else style['italic'],
                    'underline':bool(ft.underline) if ft.underline is not None else False,
                    'baseline':number_attr(rpr,'baseline',number_attr(one(para._p,'a:pPr/a:defRPr'),'baseline',0))/100000,
                    'kerning': (rpr.get('kern') if rpr is not None else None)}
            item={'text':run.text,**rstyle}
            all_runs.append(item)
            if run.text:runs.append(item)
        # python-pptx excludes <a:br/> soft breaks from paragraph.runs.
        # Preserve their positions and text-run styling instead of flattening
        # the title/body into a single line. Paragraph.text represents them
        # as vertical tabs, which CSS does not reliably render as breaks.
        if '\v' in para.text:
            raw_parts=[]
            run_index=0
            previous_style=style
            for child in para._p:
                kind=etree.QName(child).localname
                if kind=='r':
                    if run_index < len(all_runs):
                        item=all_runs[run_index]
                        previous_style={k:v for k,v in item.items() if k!='text'}
                        if item['text']:
                            raw_parts.append(item)
                    run_index+=1
                elif kind=='br':
                    # A soft break has no own glyph, but is a distinct line.
                    raw_parts.append({'text':'\n',**previous_style})
                elif kind=='fld':
                    # Field text can be represented by python-pptx as a run.
                    t=one(child,'a:t')
                    if t is not None and (t.text or ''):
                        raw_parts.append({'text':t.text,**previous_style})
            runs=raw_parts or runs
        if not runs and para.text:runs=[{'text':para.text.replace('\v','\n'),**style}]
        align=para.alignment.name.lower() if para.alignment is not None else (inherited['align'] or 'left')
        sb=one(ppr,'a:spcBef/a:spcPct'); sa=one(ppr,'a:spcAft/a:spcPct')
        sb_pct=number_attr(sb,'val')/100000 if sb is not None else inherited['spaceBeforePct']
        sa_pct=number_attr(sa,'val')/100000 if sa is not None else inherited['spaceAfterPct']
        bsz=one(ppr,'a:buSzPct')
        bsz_pct=number_attr(bsz,'val')/100000 if bsz is not None else inherited['bulletSizePct']
        result['paragraphs'].append({'runs':runs,'text':para.text,'align':align,'level':para.level,
            'bullet':bullet_char,'autoBullet':auto_bullet,'numbering':numbering if auto_bullet else None,'bulletColor':bullet_color,
            'bulletSizePct':bsz_pct,'bulletFont':bullet_font,'indentLeft':indent_left,'hanging':hanging,
            'spaceBeforePct':sb_pct,'spaceAfterPct':sa_pct,
            'spaceBefore':round(para.space_before.pt*96/72,2) if para.space_before is not None else None,
            'spaceAfter':round(para.space_after.pt*96/72,2) if para.space_after is not None else None,
            'lineSpacing':para.line_spacing if not hasattr(para.line_spacing, 'pt') else None,
            'lineSpacingPx':round(para.line_spacing.pt * 96 / 72, 3)
                              if hasattr(para.line_spacing, 'pt') else None,
            'lineSpacingReduction':spacing_scale, 'style':style})
    return result


def _table_style(shape):
    """Find the table style in the presentation's relationship-owned part."""
    props = one(shape.table._tbl, 'a:tblPr')
    style_id = one(props, 'a:tableStyleId')
    for part in shape.part.package.iter_parts():
        if str(part.partname).endswith('/tableStyles.xml'):
            root = etree.fromstring(part.blob)
            wanted = style_id.text if style_id is not None else root.get('def')
            return next((s for s in root if s.get('styleId') == wanted), None)
    return None


def _table_regions(props, row, col, rows, cols):
    enabled = lambda key: props is not None and props.get(key) in ('1', 'true')
    names = ['wholeTbl']
    # Header and label rows/columns are excluded from alternating bands.
    if enabled('bandRow'):
        names.append('band1H' if (row-int(enabled('firstRow'))) % 2 == 0 else 'band2H')
    if enabled('bandCol'):
        names.append('band1V' if (col-int(enabled('firstCol'))) % 2 == 0 else 'band2V')
    first_r, last_r = row == 0 and enabled('firstRow'), row == rows-1 and enabled('lastRow')
    first_c, last_c = col == 0 and enabled('firstCol'), col == cols-1 and enabled('lastCol')
    for flag, name in [(last_c,'lastCol'),(first_c,'firstCol'),(last_r,'lastRow'),(first_r,'firstRow'),
                       (first_r and first_c,'nwCell'),(first_r and last_c,'neCell'),
                       (last_r and first_c,'swCell'),(last_r and last_c,'seCell')]:
        if flag: names.append(name)
    return names


def _table_theme_root(slide):
    if slide is not None:
        for rel in slide.slide_layout.slide_master.part.rels.values():
            if rel.reltype.endswith('/theme'):
                return etree.fromstring(rel.target_part.blob)
    return None


def _table_reference(ref, root, theme, kind):
    """Resolve DrawingML style-matrix references and their placeholder color."""
    if ref is None or root is None: return None, theme
    index = int(number_attr(ref, 'idx', 0))
    if index <= 0: return None, theme
    if kind == 'fill' and index >= 1001:
        path, index = './/a:bgFillStyleLst', index-1000
    else:
        path = './/a:fillStyleLst' if kind == 'fill' else './/a:lnStyleLst'
    styles = one(root, path)
    if styles is None or index > len(styles): return None, theme
    resolved = dict(theme)
    if len(ref): resolved['phClr'] = color_node(ref[0], theme, '#000000')
    return styles[index-1], resolved


def _table_border(line, theme):
    if line is None: return None
    color = fill_style(line, theme, '#000000')
    dash = one(line, 'a:prstDash')
    return {'color': color, 'width': px(number_attr(line, 'w', 12700)),
            'dash': dash.get('val') if dash is not None else 'solid'}


def parse_table(shape, theme, theme_fonts, slide=None):
    table = shape.table
    widths = [px(col.width) for col in table.columns]
    heights = [px(row.height) for row in table.rows]
    props = one(table._tbl, 'a:tblPr')
    style = _table_style(shape)
    theme_root = _table_theme_root(slide)
    cells=[]
    for ri, row in enumerate(table.rows):
        crow=[]
        for ci, c in enumerate(row.cells):
            bg, defaults, edges = '#FFFFFF', {}, {}
            for name in _table_regions(props, ri, ci, len(heights), len(widths)):
                region = one(style, 'a:'+name)
                tcstyle = one(region, 'a:tcStyle')
                fill = one(tcstyle, 'a:fill')
                bg = fill_style(fill, theme, bg)
                ref = one(tcstyle, 'a:fillRef')
                native, mapped = _table_reference(ref, theme_root, theme, 'fill')
                if native is not None:
                    wrapper = etree.Element('{'+NS_A+'}fill')
                    wrapper.append(deepcopy(native))
                    bg = fill_style(wrapper, mapped, bg)
                tx = one(region, 'a:tcTxStyle')
                if tx is not None:
                    for prop, key in [('b','bold'),('i','italic')]:
                        if tx.get(prop) in ('on','off','1','0','true','false'):
                            defaults[key] = tx.get(prop) in ('on','1','true')
                    fontref = one(tx,'a:fontRef')
                    if fontref is not None:
                        defaults['font'] = theme_fonts.get(fontref.get('idx'), theme_fonts['minor'])
                    for child in tx:
                        if etree.QName(child).localname.endswith('Clr'):
                            defaults['color'] = color_node(child,theme,defaults.get('color'))
                    latin = one(tx,'a:font/a:latin')
                    if latin is not None and latin.get('typeface'): defaults['font']=latin.get('typeface')
                borders = {}
                for edge in ('left','right','top','bottom','insideH','insideV'):
                    item = one(tcstyle,'a:tcBdr/a:'+edge)
                    ln = one(item,'a:ln')
                    mapped = theme
                    if ln is None:
                        ln,mapped = _table_reference(one(item,'a:lnRef'),theme_root,theme,'line')
                    if ln is not None: borders[edge]=_table_border(ln,mapped)
                # Resolve edges within each style layer so a header's bottom
                # rule overrides the whole-table interior rule.
                for edge, inner, interior in [('left','insideV',ci>0),('right','insideV',ci+c.span_width<len(widths)),
                                           ('top','insideH',ri>0),('bottom','insideH',ri+c.span_height<len(heights))]:
                    border = borders.get(inner if interior else edge, borders.get(edge))
                    if border is not None: edges[edge] = border
            tcpr = one(c._tc,'a:tcPr')
            bg = fill_style(tcpr,theme,bg)
            for edge, tag in [('left','lnL'),('right','lnR'),('top','lnT'),('bottom','lnB')]:
                line = one(tcpr,'a:'+tag)
                if line is not None: edges[edge]=_table_border(line,theme)
            ct = parse_text(c,theme,theme_fonts,text_defaults=defaults)
            ct['margin'] = [px(c.margin_top),px(c.margin_right),px(c.margin_bottom),px(c.margin_left)]
            if tcpr is not None and tcpr.get('anchor'):
                ct['vertical'] = {'ctr':'middle','b':'bottom','t':'top'}.get(tcpr.get('anchor'),'top')
            crow.append({'text':ct,'fill':bg,'margin':ct['margin'],'borders':edges,
                         'mergeOrigin': c.is_merge_origin, 'spanned': c.is_spanned,
                         'spanWidth': c.span_width, 'spanHeight': c.span_height})
        cells.append(crow)
    return {'columns':widths,'rows':heights,'cells':cells}


def parse_chart(shape, theme, slide=None):
    try:
        chart = shape.chart
        series = []
        for ser in chart.series:
            values = []
            for v in ser.values:
                try: values.append(float(v) if v is not None else None)
                except (TypeError, ValueError): values.append(None)
            series.append({'name':str(ser.name or ''),'values':values})
        categories = []
        if chart.plots and len(chart.plots):
            categories = [str(c.label) for c in chart.plots[0].categories]
        # Chart-space background defaults to transparent when the OOXML
        # does not specify a chart fill. The old renderer forced white even
        # on dark slide masters, obscuring imagery behind the chart.
        root=chart._chartSpace
        dark=False
        if slide is not None:
            for picture in slide.slide_layout.slide_master.shapes:
                for b in picture._element.xpath('.//a:blip'):
                    lum=image_luminance(b)
                    if lum and lum['brightness'] < -.25:
                        dark=True
                        break
                if dark:break
        return {'chartType':str(chart.chart_type.name).lower(), 'categories':categories,
                'background':'transparent' if not root.xpath('./c:spPr') else '#FFFFFF',
                'dark':dark,
                'series':series[:20], 'title': chart.chart_title.text_frame.text if chart.has_title else ''}
    except (AttributeError, ValueError, TypeError):
        return None


def shape_list(shapes, theme, theme_fonts, assets, warnings, source='slide', group=None, slide=None, group_fill=None):
    out=[]
    theme_root = _table_theme_root(slide)
    for shp in shapes:
        if shp.shape_type == MSO_SHAPE_TYPE.GROUP:
            group_el = one(shp._element,'p:grpSpPr/a:xfrm')
            coff = one(group_el,'a:chOff'); cext = one(group_el,'a:chExt')
            if group_el is None or coff is None or cext is None:
                warnings.append(f'Group {shp.name} has missing coordinates')
                continue
            off_x,off_y = number_attr(coff,'x'),number_attr(coff,'y')
            ch_w,ch_h = max(1,number_attr(cext,'cx')),max(1,number_attr(cext,'cy'))
            gr_w,gr_h = float(shp.width),float(shp.height)
            if group:
                bx,by,sx,sy,ox,oy = group
                x = bx+(float(shp.left)-ox)*sx
                y = by+(float(shp.top)-oy)*sy
                ww,hh = gr_w*sx,gr_h*sy
            else:
                x,y,ww,hh = float(shp.left),float(shp.top),gr_w,gr_h
            new_group = (x,y,ww/ch_w,hh/ch_h,off_x,off_y)
            group_color=fill_style(one(shp._element,'p:grpSpPr'),theme,group_fill)
            out += shape_list(shp.shapes,theme,theme_fonts,assets,warnings,source,new_group,slide,group_color)
            continue
        if group:
            gx,gy,sx,sy,ox,oy=group
            x,y = gx + (float(shp.left)-ox)*sx, gy+(float(shp.top)-oy)*sy
            w,h = float(shp.width)*sx, float(shp.height)*sy
        else:
            x,y,w,h = (float(shp.left),float(shp.top),float(shp.width),float(shp.height))
        el = shp._element
        sppr = one(el, 'p:spPr')
        geometry = one(sppr,'a:prstGeom')
        preset = geometry.get('prst') if geometry is not None else 'rect'
        custom_geom = one(sppr, 'a:custGeom') is not None
        xfrm = one(sppr,'a:xfrm')
        if xfrm is None: xfrm = one(el,'p:xfrm')
        kind='shape'
        if shp.shape_type == MSO_SHAPE_TYPE.PICTURE: kind='image'
        elif getattr(shp,'has_table',False): kind='table'
        elif getattr(shp,'has_chart',False): kind='chart'
        elif shp.shape_type == MSO_SHAPE_TYPE.LINE: kind='connector'
        elif shp.shape_type == MSO_SHAPE_TYPE.PLACEHOLDER: kind='placeholder'
        elif getattr(shp,'has_text_frame',False) and preset=='rect' and sppr is None: kind='text'
        obj={'id':f'{source}-{shp.shape_id}', 'source':source, 'name':shp.name,
             'kind':kind,'shapeType':str(shp.shape_type),
             'placeholderType':shp.placeholder_format.type.name if shp.is_placeholder else None, 'x':px(x),'y':px(y),
             'w':px(w),'h':px(h),'rotation':round(float(shp.rotation or 0),2),
             'flipH':bool(xfrm is not None and xfrm.get('flipH') == '1'),
             'flipV':bool(xfrm is not None and xfrm.get('flipV') == '1'),
             'preset':preset,'customGeometry':custom_geom, 'paths':custom_svg_paths(sppr),
             'fill':shape_fill(el,theme,theme_root,group_fill), 'line':line_style(el,theme,theme_root),
             'fillImage':image_fill(el,slide,source,assets,warnings,theme) if kind!='image' else None,
             'text':None,'image':None,'table':None,'chart':None,
             'frameAdj':float((one(geometry,'a:avLst/a:gd').get('fmla','val 5000').split()[-1]) or 5000)/100000 if preset=='frame' and one(geometry,'a:avLst/a:gd') is not None else 0.05,
             'fieldType':next((fld.get('type') for fld in el.xpath('.//a:fld') if fld.get('type') in ('slidenum','datetime')),None)}
        if custom_geom and (not obj['paths'] or any(not p['supported'] for p in obj['paths'])):
            warnings.append(f'Custom geometry for {shp.name} partially approximated')
        if kind=='image':
            try:
                im=shp.image
                digest=hashlib.sha1(im.blob).hexdigest()[:18]
                if digest not in assets: assets[digest]={'bytes':im.blob,'type':im.content_type}
                browser_types=('image/png','image/jpeg','image/gif','image/webp','image/svg+xml')
                image_ok=im.content_type in browser_types
                if not image_ok:
                    # On Windows, GDI handles legacy EMF, including EMF files
                    # mislabeled image/x-wmf in older PowerPoint packages.
                    converted=None
                    converted_type=None
                    try:
                        # Prefer SVG for inspectable, resolution-independent
                        # paths on *all* platforms. Reject unsupported EMF
                        # records instead of omitting part of a logo.
                        if im.blob[:4] == b'\x01\x00\x00\x00':
                            converted=emf_to_svg(im.blob)
                            converted_type='image/svg+xml' if converted else None
                        if not converted:
                            converted=emf_to_png(im.blob,round(px(w)*2),round(px(h)*2))
                            converted_type='image/png' if converted else None
                        if not converted and im.blob[:4] == b'\x01\x00\x00\x00':
                            converted=emf_to_svg_compat(im.blob)
                            converted_type='image/svg+xml' if converted else None
                    except Exception as exc:
                        converted=None
                        warnings.append(f'EMF preview failed for {shp.name}: {type(exc).__name__}: {exc}')
                    if converted:
                        digest=hashlib.sha1(converted).hexdigest()[:18]
                        assets[digest]={'bytes':converted,'type':converted_type}
                        image_ok=True
                    else:
                        warnings.append(f'Cannot preview metafile {im.content_type}: {shp.name} (unsupported EMF features)')
                if image_ok:
                    changed,ctype=image_color_changes(assets[digest]['bytes'],assets[digest]['type'],one(el,'p:blipFill/a:blip'),theme,warnings)
                    digest=hashlib.sha1(changed).hexdigest()[:18]
                    assets.setdefault(digest,{'bytes':changed,'type':ctype})
                    obj['image']={'url':f'/api/media/{digest}','crop':[float(shp.crop_left),float(shp.crop_top),float(shp.crop_right),float(shp.crop_bottom)],
                                  'luminance':image_luminance(one(el,'p:blipFill/a:blip'))}
                else:
                    obj['image']={'url':None,'crop':[0,0,0,0],'unsupported':im.content_type}
            except (AttributeError, ValueError):
                warnings.append(f'Cannot extract picture: {shp.name}')
        if one(el,'a:graphic/a:graphicData/p:oleObj') is not None:
            preview=ole_preview(el,slide,source,assets,warnings)
            if preview:
                obj['kind']='image';obj['image']=preview
        if getattr(shp,'has_text_frame',False):
            obj['text']=parse_text(shp,theme,theme_fonts,slide if source=='slide' else None)
        if kind=='table': obj['table']=parse_table(shp,theme,theme_fonts,slide)
        if kind=='chart':
            warnings.append('Charts use approximate browser rendering, not PowerPoint-native chart layout')
            obj['chart']=parse_chart(shp,theme,slide)
            if not obj['chart']: warnings.append(f'Chart could not be read: {shp.name}')
        if kind in ('shape','connector','placeholder') and preset not in ('rect','roundRect','ellipse','triangle','rtTriangle','diamond','chevron','rightArrow','leftArrow','downArrow','upArrow','hexagon','pentagon','parallelogram','trapezoid','line','arc','star5','heart','pie','donut','flowChartProcess','flowChartDecision','rect','frame','straightConnector1'):
            warnings.append(f'Approximate preset geometry: {preset} ({shp.name})')
        out.append(obj)
    return out



def smartart_soft_fill(sppr, theme, fallback, background=None):
    """Approximate the low-saturation material of cached SmartArt pastel cards.

    Some Office diagram caches specify a translucent tinted accent for the
    surface, while Office's simple SmartArt material produces a much more
    neutral panel. A muted neutral theme swatch is preferable to treating the
    raw tinted RGBA as an ordinary SVG rectangle over the dark slide.
    Only applies to softened accent fills in the *cached diagram drawing*;
    other DrawingML fills retain their exact declared colors.
    """
    solid=one(sppr,'a:solidFill')
    if solid is None or len(solid)!=1:return fallback
    color=solid[0]
    if etree.QName(color).localname!='schemeClr':return fallback
    tint=one(color,'a:tint');alpha=one(color,'a:alpha')
    if tint is None or alpha is None or number_attr(tint,'val') < 30000:return fallback
    accent=theme.get(color.get('val'))
    if not accent or not re.fullmatch(r'#[0-9A-Fa-f]{6}',accent):return fallback
    def triplet(value):return tuple(int(value[i:i+2],16) for i in (1,3,5))
    # An available muted theme swatch provides a deck-specific neutral anchor.
    # Avoid picking saturated or dark colors (or an unrelated palette accent).
    neutral_candidates=[]
    for key in ('accent3','accent4','accent5','accent6','lt2'):
        candidate=theme.get(key,'')
        if not re.fullmatch(r'#[0-9A-Fa-f]{6}',candidate):continue
        v=triplet(candidate)
        if min(v)<175 or max(v)-min(v)>35:continue
        neutral_candidates.append((max(v)-min(v)+abs(sum(v)/3-218)/6,v))
    if not neutral_candidates:return fallback
    neutral=min(neutral_candidates,key=lambda pair:pair[0])[1]
    source=triplet(accent)
    # Hue is retained subtly, rather than covering an unrelated dark surface
    # with nearly-opaque saturated yellow (which exaggerates this material).
    # Office's muted SmartArt material keeps the hue family but substantially
    # desaturates its underlying accent. A warm beige is not equivalent to a
    # translucent saturated fill over the blue slide background.
    hue, light, saturation = colorsys.rgb_to_hls(*(c / 255 for c in source))
    # Material's warmth follows the accent hue, rather than a deck-specific RGB.
    hue = (hue - 8/360) % 1
    r, g, b = colorsys.hls_to_rgb(hue, .816, .35)
    material_rgb = tuple(round(c*255) for c in (r,g,b))
    # A cached Office diagram's soft surface is explicitly semi-transparent.
    # Earlier versions approximated its *final* color as opaque beige, hiding
    # underlying watermark artwork. Preserve the OOXML alpha by calculating
    # a pale source material that yields the same neutral shade when composited
    # over this slide's theme background. The watermark can then shine through.
    opacity = number_attr(alpha, 'val', 100000) / 100000
    opacity = max(.01, min(1.0, opacity))
    if opacity >= .999:
        return '#%02X%02X%02X' % material_rgb
    background_color = background or ''
    # Resolve a representative theme background color, not a reference image.
    # For a gradient the upper stop is the material's base; the browser still
    # performs the actual per-pixel compositing over the whole slide.
    bg_match = re.search(r'#[0-9A-Fa-f]{6}', background_color)
    backdrop = triplet(bg_match.group(0)) if bg_match else (255,255,255)
    foreground = [max(0, min(255, round((v-(1-opacity)*back)/opacity)))
                  for v,back in zip(material_rgb,backdrop)]
    return 'rgba(%d, %d, %d, %.4f)' % (*foreground, opacity)


def diagram_shapes(frame, slide, theme, theme_fonts, warnings):
    """Expand cached Office SmartArt drawing shapes into inspectable objects.

    Office stores editable SmartArt data AND an already-laid-out DrawingML
    cache (ppt/diagrams/drawingN.xml). The latter preserves PowerPoint's
    layout without implementing the entire SmartArt layout algorithm. This
    produces native browser objects, not a baked image or slide-specific art.
    """
    root = frame._element
    graphic = root.find('.//a:graphicData', namespaces=NS)
    if graphic is None or 'drawingml/2006/diagram' not in graphic.get('uri',''):
        return None
    relids = next((child for child in graphic if etree.QName(child).localname=='relIds'), None)
    dm_rid = relids.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}dm') if relids is not None else None
    suffix = None
    if dm_rid in slide.part.rels:
        data_name = str(slide.part.rels[dm_rid].target_part.partname)
        m = re.search(r'/data(\d+)\.xml$',data_name)
        if m: suffix = m.group(1)
    possible = [rel for rel in slide.part.rels.values()
                if rel.reltype.endswith('/diagramDrawing') and
                (suffix is None or str(rel.target_part.partname).endswith('/drawing'+suffix+'.xml'))]
    if not possible:
        warnings.append(f'Cached SmartArt drawing unavailable: {frame.name}')
        return []
    try:
        doc = etree.fromstring(possible[0].target_part.blob)
    except etree.XMLSyntaxError:
        warnings.append(f'Invalid cached SmartArt drawing: {frame.name}')
        return []
    ns = {'a':NS['a'], 'dsp':'http://schemas.microsoft.com/office/drawing/2008/diagram'}
    elements = doc.findall('.//dsp:spTree/dsp:sp', namespaces=ns)
    # Content coordinates are local to the SmartArt graphic frame.
    # Only resolve the slide backdrop once; standalone cached drawing tests
    # may supply a minimal slide without a linked master.
    backdrop = background_for(slide,theme)['color'] if hasattr(slide,'slide_layout') else None
    result=[]
    for idx, element in enumerate(elements):
        sppr=element.find('dsp:spPr',namespaces=ns)
        xfrm=one(sppr,'a:xfrm')
        off=one(xfrm,'a:off'); ext=one(xfrm,'a:ext')
        if off is None or ext is None:continue
        left, top = px(frame.left+int(number_attr(off,'x'))),px(frame.top+int(number_attr(off,'y')))
        width,height=px(number_attr(ext,'cx')),px(number_attr(ext,'cy'))
        if width<=0 or height<=0:continue
        textbody=element.find('dsp:txBody',namespaces=ns)
        bodypr=one(textbody,'a:bodyPr')
        margin=[px(number_attr(bodypr,'tIns')),px(number_attr(bodypr,'rIns')),
                px(number_attr(bodypr,'bIns')),px(number_attr(bodypr,'lIns'))]
        styref=element.find('dsp:style/a:fontRef', namespaces=ns)
        default_color=color_node(styref[0],theme) if styref is not None and len(styref) else '#111111'
        paragraphs=[]
        for p in many(textbody,'a:p'):
            ppr=one(p,'a:pPr')
            align={'ctr':'center','r':'right','just':'justify'}.get(ppr.get('algn') if ppr is not None else '','left')
            bullet_el=one(ppr,'a:buChar')
            bullet=bullet_el.get('char','•') if bullet_el is not None else None
            spc_before=one(ppr,'a:spcBef/a:spcPct')
            spc_after=one(ppr,'a:spcAft/a:spcPct')
            line_pct=one(ppr,'a:lnSpc/a:spcPct')
            line=number_attr(line_pct,'val',112000)/100000 if line_pct is not None else 1.1
            runs=[]
            for r in p:
                if etree.QName(r).localname not in ('r','fld'):
                    continue
                t=one(r,'a:t')
                if t is None:continue
                rp=one(r,'a:rPr')
                font=one(rp,'a:latin')
                color=fill_style(rp,theme,default_color)
                sz=number_attr(rp,'sz',2200)/100*96/72
                runs.append({'text':t.text or '', 'font':font.get('typeface') if font is not None else theme_fonts['minor'],
                             'size':round(sz,3),'color':color,'bold':rp is not None and rp.get('b')=='1',
                             'italic':rp is not None and rp.get('i')=='1','underline':False})
            paragraphs.append({'runs':runs,'text':''.join(r['text'] for r in runs),'align':align,
                'level':int(number_attr(ppr,'lvl',0)), 'bullet':bullet,'autoBullet':False,
                'bulletColor':'#111111','bulletSizePct':1,
                'indentLeft':px(number_attr(ppr,'marL',0)) if bullet else None,
                'hanging':px(number_attr(ppr,'indent',0)) if bullet else None,
                'spaceBeforePct':number_attr(spc_before,'val')/100000 if spc_before is not None else 0,
                'spaceAfterPct':number_attr(spc_after,'val')/100000 if spc_after is not None else 0,
                # Empty paragraphs in Office's cached SmartArt typically
                # retain end-of-paragraph baseline spacing. Bare CSS line
                # boxes collapse these gaps and compress each card vertically.
                'spaceAfter':26 if not runs else None,
                'lineSpacing':line,'lineSpacingReduction':0,
                'style':{'font':theme_fonts['minor'],'size':round(22*96/72,3),
                         'color':default_color,'bold':False,'italic':False}})
        prst=one(sppr,'a:prstGeom')
        line=one(sppr,'a:ln')
        color=fill_style(sppr,theme,'transparent')
        color=smartart_soft_fill(sppr,theme,color,backdrop)
        # Simple SmartArt pastel material differs from the raw tint + alpha.
        lnfill=fill_style(line,theme) if line is not None else None
        if line is not None:
            lnfill=smartart_soft_fill(line,theme,lnfill,backdrop)
        line_w=number_attr(line,'w',0)/EMU_PER_PX if line is not None else 0
        title=' '.join(p['text'] for p in paragraphs if p['text'])[:70]
        result.append({'id':f'smartart-{frame.shape_id}-{idx}','source':'smartart','name':title or f'SmartArt {idx+1}',
          'kind':'shape','shapeType':'SMARTART_CACHED_SHAPE','placeholderType':None,
          'x':left,'y':top,'w':width,'h':height,'rotation':0,'flipH':False,'flipV':False,
          'preset':prst.get('prst','rect') if prst is not None else 'rect',
          'customGeometry':False,'paths':[],'fill':color,
          'line':{'color':lnfill,'width':line_w} if lnfill and line_w else None,
          'text':{'margin':margin,'vertical':{'ctr':'middle','b':'bottom'}.get(bodypr.get('anchor') if bodypr is not None else '','top'),
                  'wrap':True,'autoSize':None,'fontScale':1,'paragraphs':paragraphs},
          'image':None,'table':None,'chart':None})
    return result


def validate(data: bytes):
    if not data.startswith(b'PK'):
        raise ValueError('Not a valid .pptx ZIP container')
    if len(data) > PREVIEW_LIMITS['bytes']:
        raise ValueError('PPTX exceeds 64 MB upload limit')
    try:
        with ZipFile(BytesIO(data)) as z:
            infos=z.infolist()
            if len(infos)>PREVIEW_LIMITS['entries']:
                raise ValueError('PPTX contains too many entries')
            if sum(i.file_size for i in infos)>PREVIEW_LIMITS['uncompressed']:
                raise ValueError('Expanded PPTX exceeds safety limit')
            if 'ppt/presentation.xml' not in z.namelist():
                raise ValueError('ZIP does not contain a PowerPoint presentation')
    except BadZipFile as exc:
        raise ValueError('Invalid PPTX ZIP container') from exc


def parse_pptx(data: bytes, filename='presentation.pptx'):
    validate(data)
    prs=Presentation(BytesIO(data))
    if len(prs.slides)>PREVIEW_LIMITS['slides']:
        raise ValueError('Too many slides (>400)')
    assets={}
    warnings=[]
    slides=[]
    for i,slide in enumerate(prs.slides):
        theme,theme_fonts=theme_for(slide)
        shapes=[]
        if getattr(slide,'follow_master_graphics',True) and slide.slide_layout._element.get('showMasterSp','1') != '0':
            shapes += shape_list([s for s in slide.slide_layout.slide_master.shapes if not s.is_placeholder],theme,theme_fonts,assets,warnings,'master',slide=slide)
        shapes += shape_list([s for s in slide.slide_layout.shapes if not s.is_placeholder],theme,theme_fonts,assets,warnings,'layout',slide=slide)
        # Expand cached SmartArt in slide z-order rather than appending it on top.
        # Shape types vary between python-pptx releases; inspect OOXML instead.
        for shp in slide.shapes:
            parsed = (diagram_shapes(shp,slide,theme,theme_fonts,warnings)
                      if etree.QName(shp._element).localname=='graphicFrame' else None)
            if parsed is None:
                shapes += shape_list([shp],theme,theme_fonts,assets,warnings,'slide',slide=slide)
            else:
                shapes += parsed
        # Resolve dynamic slide-number fields through their master placeholder.
        # Their cached ‹#› text is not display text, and these layout fields
        # often have no local font/color despite a fully styled master level.
        from pptx.enum.shapes import PP_PLACEHOLDER
        master_num=next((p for p in slide.slide_layout.slide_master.placeholders
                         if p.placeholder_format.type==PP_PLACEHOLDER.SLIDE_NUMBER),None)
        master_level=one(master_num._element,'p:txBody/a:lstStyle/a:lvl1pPr') if master_num is not None else None
        master_rpr=one(master_level,'a:defRPr')
        field_color=fill_style(master_rpr,theme) if master_rpr is not None else None
        field_size=number_attr(master_rpr,'sz',0)/100*96/72 if master_rpr is not None else None
        field_align={'r':'right','ctr':'center','l':'left'}.get(master_level.get('algn')) if master_level is not None else None
        for obj in shapes:
            if obj.get('fieldType') != 'slidenum' or not obj.get('text'):
                continue
            for paragraph in obj['text']['paragraphs']:
                paragraph['text']=paragraph['text'].replace('‹#›',str(i+1))
                if field_align and paragraph['align']=='left':paragraph['align']=field_align
                if field_color and not paragraph['style']['color']:paragraph['style']['color']=field_color
                if field_size and not paragraph['style']['size']:paragraph['style']['size']=field_size
                for run in paragraph['runs']:
                    run['text']=run.get('text','').replace('‹#›',str(i+1))
                    if field_color and not run.get('color'):run['color']=field_color
                    if field_size and not run.get('size'):run['size']=field_size
        slide_title = next((s['text']['paragraphs'][0]['text'] for s in shapes if s['source']=='slide' and s.get('text') and s['text']['paragraphs'] and s['text']['paragraphs'][0]['text']), f'Slide {i+1}')
        bg=background_for(slide,theme,assets,warnings)
        slides.append({'number':i+1,'title':slide_title[:110], 'background':bg['color'], 'backgroundImage':bg['image'], 'backgroundImageTile':bg.get('imageTile'),
                       'shapes':shapes, 'theme':theme})
    model={'filename':filename,'width':px(prs.slide_width),'height':px(prs.slide_height),
           'count':len(slides),'slides':slides,'warnings':sorted(set(warnings)),
           'renderer':'browser scene graph / preview'}
    return model, assets
