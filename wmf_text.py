"""Conservative, portable WMF text-preview reader; never executes OLE data.

Accepts font selection and positioned text with explicit window extents.
Rejects unsupported drawing/state records instead of silently dropping art.
"""
import html
import struct


def wmf_text_to_svg(blob):
    if len(blob)<18 or len(blob)>16*1024*1024:return None
    start=22 if blob[:4]==b'\xd7\xcd\xc6\x9a' else 0
    try:
        kind,header,version,words,capacity,_,_=struct.unpack_from('<HHHIHIH',blob,start)
        if kind not in (1,2) or header!=9 or version not in (0x100,0x300):return None
        end=start+words*2
        if end>len(blob) or capacity>4096:return None
        pos=start+18;fonts={};font=None;nodes=[];origin=(0,0);extent=None;align=0;color='#000000'
        greek=dict(zip('ABGDEZHQIKLMNXOPRSTUFCYW','ΑΒΓΔΕΖΗΘΙΚΛΜΝΞΟΠΡΣΤΥΦΧΨΩ'))
        greek.update(dict(zip('abgdezhqiklmnxoprstufcyw','αβγδεζηθικλμνξοπρστυφχψω')))
        greek.update({'V':'ς','J':'ϑ','j':'ϕ','v':'ϖ','-':'−'})
        eof=False
        while pos+6<=end:
            size,op=struct.unpack_from('<IH',blob,pos)
            if size<3 or pos+size*2>end:return None
            data=blob[pos+6:pos+size*2];pos+=size*2
            if op==0:eof=True;break
            if op==0x020b:
                y,x=struct.unpack_from('<hh',data);origin=(x,y)
            elif op==0x020c:
                h,w=struct.unpack_from('<hh',data)
                if w<=0 or h<=0:return None
                extent=(w,h)
            elif op==0x0209:
                r,g,b=struct.unpack_from('BBB',data);color=f'#{r:02x}{g:02x}{b:02x}'
            elif op==0x012e:
                align=struct.unpack_from('<H',data)[0]
                if align & ~30:return None
            elif op==0x0102:
                if struct.unpack_from('<H',data)[0]!=1:return None # opaque backgrounds unsupported
            elif op==0x0201:pass # background color has no effect in transparent mode
            elif op==0x0626:
                # MFCOMMENT is metadata (e.g. MathType); no escape is executed.
                escape,count=struct.unpack_from('<HH',data)
                if escape!=15 or count>len(data)-4:return None
            elif op==0x02fb:
                height,width,esc,orient,weight,italic,under,strike,charset,*_=struct.unpack_from('<5h8B',data)
                if esc or orient or under or strike or charset not in (0,1,2):return None
                family=data[18:].split(b'\0',1)[0].decode('cp1252')
                index=next((i for i in range(capacity) if i not in fonts),None)
                if index is None:return None
                fonts[index]=(abs(height),family,weight,italic,charset)
            elif op==0x012d:
                index=struct.unpack_from('<H',data)[0]
                if index not in fonts:return None
                font=fonts[index]
            elif op==0x01f0:
                fonts.pop(struct.unpack_from('<H',data)[0],None)
            elif op==0x0a32:
                if font is None or extent is None:return None
                y,x,count,options=struct.unpack_from('<hhHH',data)
                if options or count>len(data)-8:return None
                raw=data[8:8+count];size,family,weight,italic,charset=font
                if size<=0:return None
                if charset==2 or family.lower()=='symbol':
                    if any(c>=127 or (chr(c).isalpha() and chr(c) not in greek) for c in raw):return None
                    text=''.join(greek.get(chr(c),chr(c)) for c in raw);family='Times New Roman'
                else:text=raw.decode('cp1252')
                if any(ord(c)<32 for c in text):return None
                offset=8+count+(count%2);remaining=len(data)-offset
                positions=None
                if remaining:
                    if remaining!=count*2:return None
                    advances=struct.unpack_from('<'+'h'*count,data,offset)
                    positions=[];cur=x
                    for advance in advances:positions.append(str(cur));cur+=advance
                    if align&6:return None
                anchor='middle' if align&6==6 else ('end' if align&2 else 'start')
                baseline='' if align&24==24 else (' dominant-baseline="text-after-edge"' if align&8 else ' dominant-baseline="text-before-edge"')
                xs=' '.join(positions) if positions is not None else str(x)
                nodes.append(f'<text x="{xs}" y="{y}" font-family="{html.escape(family,quote=True)}, serif" font-size="{size}" font-weight="{weight}" font-style="{"italic" if italic else "normal"}" text-anchor="{anchor}" fill="{color}"{baseline}>{html.escape(text)}</text>')
            else:return None
        if not eof or not nodes or extent is None:return None
        return (f'<svg xmlns="http://www.w3.org/2000/svg" viewBox="{origin[0]} {origin[1]} {extent[0]} {extent[1]}" preserveAspectRatio="none">'+''.join(nodes)+'</svg>').encode()
    except (ValueError,struct.error,UnicodeError,IndexError):return None
