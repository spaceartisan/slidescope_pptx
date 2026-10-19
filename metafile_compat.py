"""Portable, best-effort GDI EMF drawing fallback.

No external executables, runtime services, or new pip dependencies. Converts
frequent GDI fallback primitives (text, filled shapes, bitmaps, clipping,
world transforms) to browser SVG. On Windows, the native GDI raster path is
preferable for complex EMF and should be tried before this fallback.

This is not a full EMF/EMF+ interpreter: unsupported records are tracked and
conversion is rejected if too many unrecoverable drawing primitives remain.
"""
from __future__ import annotations
import base64
from io import BytesIO
import html
import math
import struct
from PIL import Image


def emf_to_svg_compat(blob: bytes) -> bytes | None:
    if len(blob) < 108 or len(blob)>16*1024*1024 or blob[:4]!=b'\x01\x00\x00\x00' or blob[40:44]!=b' EMF':
        return None
    def u32(p):return struct.unpack_from('<I',blob,p)[0]
    def i32(p):return struct.unpack_from('<i',blob,p)[0]
    def rgba(v):return f'#{v&255:02x}{(v>>8)&255:02x}{(v>>16)&255:02x}'
    def fmt(v):return format(v,'.6g')
    try:
        total=min(len(blob),u32(48))
        x0,y0,x1,y1=struct.unpack_from('<4i',blob,8)
        if not (x1>x0 and y1>y0 and x1-x0<10000000 and y1-y0<10000000):return None
        px_width=x1-x0;px_height=y1-y0
        nodes=[];objects={}; cur=(0,0);path=[];in_path=False;fillrule='nonzero'
        current_pen={'color':'#000','width':1};current_brush={'color':'#ffffff'}
        current_font={'family':'Arial','size':16,'weight':'400','italic':False}
        text_color='#000000';text_align=0
        matrix=(1.,0.,0.,1.,0.,0.)
        clip=None; state_stack=[]; ids=0
        unresolved=0; rendered=0
        win_org=(0.,0.);win_ext=(1.,1.);view_org=(0.,0.);view_ext=(1.,1.)
        # Applied in the SVG to keep native GDI logical coordinates.
        def transform():
            a,b,c,d,e,f=matrix
            sx=view_ext[0]/win_ext[0] if win_ext[0] else 1
            sy=view_ext[1]/win_ext[1] if win_ext[1] else 1
            # Source logical transform -> viewport transform.
            return (a*sx,b*sy,c*sx,d*sy,(e-win_org[0])*sx+view_org[0],(f-win_org[1])*sy+view_org[1])
        def t_attr():
            a,b,c,d,e,f=transform()
            if abs(a-1)<1e-10 and abs(b)+abs(c)+abs(d-1)+abs(e)+abs(f)<1e-10:return ''
            return ' transform="matrix('+ ' '.join(map(fmt,(a,b,c,d,e,f))) +')"'
        def clip_attr():return f' clip-path="url(#clip{clip})"' if clip else ''
        def emit(d, f=False,s=False):
            nonlocal rendered
            if not d:return
            fc=current_brush['color'] if f else 'none'
            sc=current_pen['color'] if s else 'none'
            sw=current_pen.get('width',1) if s else 0
            nodes.append(f'<path d="{html.escape(" ".join(d),quote=True)}" fill="{fc}" fill-rule="{fillrule}" stroke="{sc}" stroke-width="{fmt(sw)}"{t_attr()}{clip_attr()}/>')
            rendered+=1
        def xy(p):return fmt(p[0])+','+fmt(p[1])
        def rec_bitmap(pos,size,off_info,sz_info,off_pixels,sz_pixels):
            if not off_info or not off_pixels or off_info+sz_info>size or off_pixels+sz_pixels>size or sz_info<40:return None
            info=blob[pos+off_info:pos+off_info+sz_info]
            bits=blob[pos+off_pixels:pos+off_pixels+sz_pixels]
            try:
                # For BI_RGB / BI_BITFIELDS, the BMP file header lets Pillow
                # handle row pitch, palette and orientation correctly.
                header=struct.pack('<2sIHHI',b'BM',14+len(info)+len(bits),0,0,14+len(info))
                im=Image.open(BytesIO(header+info+bits));im.load()
                if im.width*im.height>20_000_000:return None
                output=BytesIO();im.convert('RGBA').save(output,format='PNG')
                return base64.b64encode(output.getvalue()).decode('ascii')
            except (OSError,ValueError,struct.error):return None
        def add_bitmap(data,x,y,w,h):
            nonlocal rendered
            if w==0 or h==0:return
            nodes.append(f'<image x="{fmt(x)}" y="{fmt(y)}" width="{fmt(w)}" height="{fmt(h)}" href="data:image/png;base64,{data}" preserveAspectRatio="none"{t_attr()}{clip_attr()}/>')
            rendered+=1
        def shape_rect(x,y,w,h,kind='rect'):
            nonlocal rendered
            if w<=0 or h<=0:return
            fc=current_brush['color'];sc=current_pen['color'];sw=current_pen.get('width',1)
            if kind=='rect':bounds=f'x="{x}" y="{y}" width="{w}" height="{h}"'
            else:bounds=f'cx="{fmt(x+w/2)}" cy="{fmt(y+h/2)}" rx="{fmt(w/2)}" ry="{fmt(h/2)}"'
            nodes.append(f'<{kind} {bounds} fill="{fc}" stroke="{sc}" stroke-width="{fmt(sw)}"{t_attr()}{clip_attr()}/>')
            rendered+=1
        def snapshot():
            return (cur[:],current_pen.copy(),current_brush.copy(),current_font.copy(),text_color,text_align,matrix,clip,win_org,win_ext,view_org,view_ext,fillrule)
        pos=0;count=0
        while pos+8<=total and count<100000:
            typ,size=struct.unpack_from('<II',blob,pos)
            if size<8 or size%4 or pos+size>total:return None
            count+=1
            if typ==14:break
            if typ==9 and size>=16:win_ext=tuple(float(v) for v in struct.unpack_from('<ii',blob,pos+8))
            elif typ==10 and size>=16:win_org=tuple(float(v) for v in struct.unpack_from('<ii',blob,pos+8))
            elif typ==11 and size>=16:view_ext=tuple(float(v) for v in struct.unpack_from('<ii',blob,pos+8))
            elif typ==12 and size>=16:view_org=tuple(float(v) for v in struct.unpack_from('<ii',blob,pos+8))
            elif typ==19 and size>=12:fillrule='evenodd' if u32(pos+8)==1 else 'nonzero'
            elif typ==24 and size>=12:text_color=rgba(u32(pos+8))
            elif typ==22 and size>=12:text_align=u32(pos+8)
            elif typ==35 and size>=32:matrix=struct.unpack_from('<6f',blob,pos+8)
            elif typ==36 and size>=36:
                m=struct.unpack_from('<6f',blob,pos+8)
                mode=u32(pos+32)
                if mode==1:matrix=(1.,0.,0.,1.,0.,0.)
                elif mode in (2,3):
                    a,b,c,d,e,f=matrix;A,B,C,D,E,F=m
                    if mode==2:matrix=(a*A+c*B,b*A+d*B,a*C+c*D,b*C+d*D,a*E+c*F+e,b*E+d*F+f)
                    else:matrix=(A*a+C*b,B*a+D*b,A*c+C*d,B*c+D*d,A*e+C*f+E,B*e+D*f+F)
            elif typ==33:state_stack.append(snapshot())
            elif typ==34 and size>=12 and state_stack:
                (cur,current_pen,current_brush,current_font,text_color,text_align,matrix,clip,win_org,win_ext,view_org,view_ext,fillrule)=state_stack.pop()
            elif typ in (38,95) and size>=28:
                ident=u32(pos+8)
                if typ==38:
                    style=u32(pos+12);width=max(1,abs(i32(pos+16)));color=rgba(u32(pos+24))
                else:
                    style=u32(pos+28);width=max(1,abs(i32(pos+32)));color=rgba(u32(pos+40))
                objects[ident]=('pen',{'color':color if style&0x0000000F!=5 else 'none','width':min(500,width)})
            elif typ==39 and size>=24:
                ident=u32(pos+8);style=u32(pos+12)
                objects[ident]=('brush',{'color':rgba(u32(pos+16)) if style==0 else 'none'})
            elif typ==82 and size>=104:
                ident=u32(pos+8);height=abs(i32(pos+12));weight=i32(pos+28)
                family=blob[pos+40:pos+104].decode('utf-16le',errors='ignore').split('\0')[0] or 'Arial'
                family=family.replace('@','')
                obj={'family':family,'size':min(1000,max(1,height)), 'weight':str(weight) if weight>0 else '400','italic':blob[pos+32]!=0}
                objects[ident]=('font',obj)
            elif typ==37 and size>=12:
                ident=u32(pos+8)
                if ident in objects:
                    role,value=objects[ident]
                    if role=='pen':current_pen=value.copy()
                    elif role=='brush':current_brush=value.copy()
                    elif role=='font':current_font=value.copy()
                elif ident&0x80000000:
                    stock=ident&0x7fffffff
                    if stock==5:current_brush={'color':'none'}
                    elif stock==8:current_pen={'color':'none','width':0}
                    elif stock==0:current_brush={'color':'#ffffff'}
                    elif stock==4:current_brush={'color':'#000000'}
                    elif stock==7:current_pen={'color':'#000000','width':1}
            elif typ==40 and size>=12:objects.pop(u32(pos+8),None)
            elif typ==27 and size>=16:
                cur=struct.unpack_from('<ii',blob,pos+8)
                if in_path:path.append('M'+xy(cur))
            elif typ==54 and size>=16:
                point=struct.unpack_from('<ii',blob,pos+8)
                if in_path:
                    if not path:path.append('M'+xy(cur))
                    path.append('L'+xy(point))
                else:emit(['M'+xy(cur),'L'+xy(point)],s=True)
                cur=point
            elif typ in (42,43,44) and size>=24:
                x,y,r,b=struct.unpack_from('<4i',blob,pos+8)
                shape_rect(x,y,r-x,b-y,'ellipse' if typ==42 else 'rect')
            elif typ in (3,4,5,6,86,87,88,89) and size>=28:
                n=u32(pos+24)
                stride=4 if typ>=85 else 8
                if n>200000 or pos+28+n*stride>pos+size:return None
                pts=[struct.unpack_from('<hh' if stride==4 else '<ii',blob,pos+28+stride*i) for i in range(n)]
                if pts:
                    if typ in (3,4,86,87):
                        d=['M'+xy(pts[0])]+['L'+xy(p) for p in pts[1:]]
                        polygon=typ in (3,86)
                        if polygon:d.append('Z')
                    else:
                        continuation=typ in (5,6,88,89)
                        d=['M'+xy(cur)]
                        if typ in (5,88) and n%3==0:
                            d+=['C'+' '.join(xy(p) for p in pts[i:i+3]) for i in range(0,n,3)]
                        else:d+=['L'+xy(p) for p in pts]
                        polygon=False
                    if in_path:path.extend(d)
                    else:emit(d,f=polygon,s=True)
                    cur=pts[-1]
            elif typ in (7,8,90,91) and size>=32:
                polygons=u32(pos+24);countpts=u32(pos+28)
                stride=4 if typ in (90,91) else 8
                if polygons>5000 or countpts>200000 or 32+polygons*4+stride*countpts>size:return None
                sizes=struct.unpack_from('<'+'I'*polygons,blob,pos+32) if polygons else []
                start=pos+32+4*polygons;pts=[struct.unpack_from('<hh' if stride==4 else '<ii',blob,start+stride*i) for i in range(countpts)]
                d=[];off=0
                for n in sizes:
                    q=pts[off:off+n];off+=n
                    if not q:continue
                    d.append('M'+xy(q[0]));d+=['L'+xy(p) for p in q[1:]]
                    if typ in (8,91):d.append('Z')
                if d:emit(d,f=typ in (8,91),s=True)
            elif typ==59:path=[];in_path=True
            elif typ==60:in_path=False
            elif typ==61:path.append('Z')
            elif typ in (62,63,64):emit(path,f=typ in (62,63),s=typ in (63,64));path=[];in_path=False
            elif typ==68:path=[];in_path=False
            elif typ==84 and size>=76:
                ptx,pty=struct.unpack_from('<ii',blob,pos+36)
                countchars=u32(pos+44);offstr=u32(pos+48);flags=u32(pos+52)
                if countchars>250000 or not offstr or offstr+countchars*2>size:
                    unresolved+=1
                else:
                    raw=blob[pos+offstr:pos+offstr+countchars*2]
                    txt=raw.decode('utf-16le',errors='replace').replace('\x00','')
                    if txt.strip():
                        fontsize=current_font['size']
                        v_align='0' if (text_align & 24)==24 else ('.82' if not (text_align&8) else '-.18')
                        cy=pty+float(v_align)*fontsize
                        anchor='middle' if (text_align&6)==6 else ('end' if text_align&2 else 'start')
                        family=html.escape(current_font['family'],quote=True)
                        nodes.append(f'<text x="{fmt(ptx)}" y="{fmt(cy)}" font-family="{family}" font-size="{fmt(fontsize)}" font-weight="{current_font["weight"]}" font-style="{"italic" if current_font["italic"] else "normal"}" fill="{text_color}" text-anchor="{anchor}"{t_attr()}{clip_attr()}>{html.escape(txt)}</text>')
                        rendered+=1
            elif typ==81 and size>=80:
                xdest,ydest,xsrc,ysrc,cxsrc,cysrc,offinfo,szinfo,offpixels,szpixels,usage,rop,cxdest,cydest=struct.unpack_from('<14i',blob,pos+24)
                bmp=rec_bitmap(pos,size,offinfo,szinfo,offpixels,szpixels)
                if bmp:add_bitmap(bmp,xdest,ydest,cxdest,cydest)
                else:unresolved+=1
            elif typ==76 and size>=100:
                xdest,ydest,cxdest,cydest,rop=struct.unpack_from('<5i',blob,pos+24)
                offinfo,szinfo,offpixels,szpixels=struct.unpack_from('<4I',blob,pos+84)
                bmp=rec_bitmap(pos,size,offinfo,szinfo,offpixels,szpixels)
                if bmp:add_bitmap(bmp,xdest,ydest,cxdest,cydest)
                elif not szpixels and rop!=0x00AA0029:shape_rect(xdest,ydest,cxdest,cydest)
                else:unresolved+=1
            elif typ==75 and size>=16:
                cb,mode=struct.unpack_from('<2I',blob,pos+8)
                if cb==0 and mode==5:
                    clip=None
                elif cb>=48 and 16+cb<=size:
                    header=pos+16
                    rgnhdr,nkind,nrects,rgnbytes=struct.unpack_from('<4I',blob,header)
                    if rgnhdr>=32 and nkind==1 and nrects<=10000 and cb>=rgnhdr+16*nrects:
                        rectangles=[struct.unpack_from('<4i',blob,header+rgnhdr+16*k) for k in range(nrects)]
                        ids+=1
                        shapes=''.join(f'<rect x="{x}" y="{y}" width="{max(0,r-x)}" height="{max(0,b-y)}"/>' for x,y,r,b in rectangles)
                        nodes.append(f'<defs><clipPath id="clip{ids}">{shapes}</clipPath></defs>')
                        clip=ids
                # RGN_COPY/AND are common; support the region for COPY and
                # conservative clipping for AND, without guessing obscure ops.
            elif typ==30 and size>=24:
                r=struct.unpack_from('<4i',blob,pos+8)
                ids+=1;clip=ids
                nodes.append(f'<defs><clipPath id="clip{ids}"><rect x="{r[0]}" y="{r[1]}" width="{max(0,r[2]-r[0])}" height="{max(0,r[3]-r[1])}"/></clipPath></defs>')
            elif typ in (67,28,29):
                # Complex clipping regions do not correspond to one SVG rect.
                # Keep producing useful geometry, while noting incompleteness.
                clip=None
            elif typ in (1,2,9,10,11,12,13,17,18,19,20,21,22,24,25,26,30,33,34,35,36,37,38,39,40,48,49,70,75,82,91,94,95,98):
                pass
            else:
                # Unsupported painting operation, not just metadata.
                unresolved+=1
            pos+=size
        if count>=100000 or not rendered:return None
        # Warn downstream by returning None on substantial missing content.
        if unresolved>rendered*.45+30:return None
        view=f'{x0} {y0} {px_width} {px_height}'
        svg='<svg xmlns="http://www.w3.org/2000/svg" xmlns:xlink="http://www.w3.org/1999/xlink" viewBox="'+view+'" preserveAspectRatio="none">'+''.join(nodes)+'</svg>'
        if len(svg)>16*1024*1024:return None
        return svg.encode('utf-8')
    except (ValueError,ZeroDivisionError,OverflowError,IndexError,struct.error,MemoryError):return None
