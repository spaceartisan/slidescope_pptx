"""Best-effort EMF rasterization through Windows' built-in GDI.

No binaries or third-party install required; Pillow is already a python-pptx
runtime dependency. Windows-only by design. Errors should be reported rather
than silently interpreted as working renderings.
"""
from __future__ import annotations
from io import BytesIO
import os


def emf_to_png(blob: bytes, width: int, height: int) -> bytes | None:
    if os.name != 'nt' or len(blob) < 44 or blob[:4] != b'\x01\x00\x00\x00' or blob[40:44] != b' EMF':
        return None
    import ctypes
    from ctypes import wintypes
    from PIL import Image

    width = min(3200, max(1, int(width)))
    height = min(2400, max(1, int(height)))
    class RECT(ctypes.Structure):
        _fields_ = [('left',ctypes.c_long),('top',ctypes.c_long),('right',ctypes.c_long),('bottom',ctypes.c_long)]
    class BITMAPINFOHEADER(ctypes.Structure):
        _fields_ = [('biSize',wintypes.DWORD),('biWidth',ctypes.c_long),('biHeight',ctypes.c_long),
                    ('biPlanes',wintypes.WORD),('biBitCount',wintypes.WORD),('biCompression',wintypes.DWORD),
                    ('biSizeImage',wintypes.DWORD),('biXPelsPerMeter',ctypes.c_long),
                    ('biYPelsPerMeter',ctypes.c_long),('biClrUsed',wintypes.DWORD),('biClrImportant',wintypes.DWORD)]
    class BITMAPINFO(ctypes.Structure):
        _fields_ = [('bmiHeader',BITMAPINFOHEADER),('bmiColors',wintypes.DWORD * 3)]
    gdi = ctypes.WinDLL('gdi32', use_last_error=True)
    def api(name, result, *argtypes):
        f = getattr(gdi,name)
        f.restype=result
        f.argtypes=list(argtypes)
        return f
    HF=ctypes.c_void_p
    make_emf=api('SetEnhMetaFileBits',HF,wintypes.UINT,ctypes.POINTER(ctypes.c_ubyte))
    close_emf=api('DeleteEnhMetaFile',wintypes.BOOL,HF)
    make_dc=api('CreateCompatibleDC',HF,HF)
    close_dc=api('DeleteDC',wintypes.BOOL,HF)
    make_bitmap=api('CreateDIBSection',HF,HF,ctypes.POINTER(BITMAPINFO),wintypes.UINT,
                    ctypes.POINTER(HF),HF,wintypes.DWORD)
    select=api('SelectObject',HF,HF,HF)
    delete_obj=api('DeleteObject',wintypes.BOOL,HF)
    play=api('PlayEnhMetaFile',wintypes.BOOL,HF,HF,ctypes.POINTER(RECT))
    data=(ctypes.c_ubyte * len(blob)).from_buffer_copy(blob)
    hemf=make_emf(len(blob),data)
    if not hemf:return None
    try:
        dc=make_dc(None)
        if not dc:return None
        try:
            info=BITMAPINFO()
            info.bmiHeader.biSize=ctypes.sizeof(BITMAPINFOHEADER)
            info.bmiHeader.biWidth=width
            info.bmiHeader.biHeight=-height  # top-down DIB, positive SVG/PPTX y-axis
            info.bmiHeader.biPlanes=1
            info.bmiHeader.biBitCount=32
            info.bmiHeader.biCompression=0
            pixels=HF()
            bitmap=make_bitmap(dc,ctypes.byref(info),0,ctypes.byref(pixels),None,0)
            if not bitmap or not pixels.value:return None
            old=select(dc,bitmap)
            try:
                buffers=[]
                for bg in (0,255):
                    ctypes.memset(pixels.value,bg,width*height*4)
                    if not play(dc,hemf,ctypes.byref(RECT(0,0,width,height))):return None
                    buffers.append(ctypes.string_at(pixels.value,width*height*4))
            finally:
                select(dc,old)
                delete_obj(bitmap)
        finally:
            close_dc(dc)
    finally:
        close_emf(hemf)
    black,white=buffers
    rgba=bytearray(width*height*4)
    # Estimate alpha from rendering over two different backgrounds, removing
    # the background contribution to retain transparency around logo artwork.
    for i in range(0,len(rgba),4):
        # Native DIB is BGRA; ignore GDI alpha byte, which is not reliable.
        opacity=255 - round(sum(max(0,white[i+j]-black[i+j]) for j in range(3))/3)
        opacity=max(0,min(255,opacity))
        for j in range(3):
            rgba[i+2-j]=min(255,round(black[i+j]*255/opacity)) if opacity else 0
        rgba[i+3]=opacity
    output=BytesIO()
    Image.frombytes('RGBA',(width,height),bytes(rgba)).save(output,format='PNG')
    return output.getvalue()


def emf_to_svg(blob: bytes) -> bytes | None:
    """Translate a conservative subset of vector EMF to browser-native SVG.

    Uses only the Python standard library. This handles common GDI paths
    (line segments, 16-bit cubic Béziers, polygons, pens, and solid brushes).
    Unsupported *drawing* commands cause an explicit fallback, rather than
    producing a misleadingly incomplete preview. The EMF bytes are untrusted.
    """
    import html
    import struct

    if (len(blob) < 108 or len(blob) > 16 * 1024 * 1024 or
            blob[:4] != b'\x01\x00\x00\x00' or blob[40:44] != b' EMF'):
        return None
    u32 = lambda off: struct.unpack_from('<I', blob, off)[0]
    i32 = lambda off: struct.unpack_from('<i', blob, off)[0]
    try:
        total = u32(48)
        frame = struct.unpack_from('<4i', blob, 24)
        # Frame coordinates are hundredths of millimeters, so the physical
        # aspect ratio remains independent of the original print DPI.
        fw, fh = frame[2]-frame[0], frame[3]-frame[1]
        if not 0 < fw < 10000000 or not 0 < fh < 10000000 or total > len(blob):
            return None
        wnd_x, wnd_y, wnd_w, wnd_h = 0, 0, fw, fh
        pos=0; count=0; nodes=[]; path=[]; current=(0,0)
        in_path=False
        objects={}
        pen={'color':'#000000','width':1}; brush={'color':'#000000'}
        fill_rule='nonzero'
        selected_pen=pen; selected_brush=brush

        def xy(x,y):
            # Preserve the native logical coordinate system; the SVG viewBox
            # maps it to the whole picture, just as the EMF viewport does.
            return f'{x:g},{y:g}'

        def emit(d, fill=False, stroke=False):
            if not d:return
            fc=selected_brush['color'] if fill and selected_brush else 'none'
            sc=selected_pen['color'] if stroke and selected_pen else 'none'
            sw=selected_pen.get('width',1) if stroke and selected_pen else 0
            nodes.append(f'<path d="{html.escape(" ".join(d), quote=True)}" '
                         f'fill="{fc}" fill-rule="{fill_rule}" stroke="{sc}" '
                         f'stroke-width="{sw:g}" stroke-linejoin="round"/>')

        while pos + 8 <= total and count < 100000:
            typ, size = struct.unpack_from('<II', blob, pos)
            if size < 8 or size % 4 or pos + size > total:
                return None
            count += 1
            if typ == 14:  # EMR_EOF
                break
            if typ == 9 and size >= 16:  # SETWINDOWEXTEX
                wnd_w, wnd_h = struct.unpack_from('<ii', blob, pos+8)
            elif typ == 10 and size >= 16:  # SETWINDOWORGEX
                wnd_x, wnd_y = struct.unpack_from('<ii', blob, pos+8)
            elif typ == 19 and size >= 12:
                fill_rule = 'evenodd' if u32(pos+8) == 1 else 'nonzero'
            elif typ == 95 and size >= 52:  # EXTCREATEPEN
                index=u32(pos+8)
                bgr=u32(pos+40)
                style=u32(pos+28)
                color=f'#{bgr&255:02X}{(bgr>>8)&255:02X}{(bgr>>16)&255:02X}'
                obj={'color':color,'width':max(1,min(10000,u32(pos+32))),
                     'style':style}
                objects[index]=('pen',obj)
            elif typ == 39 and size >= 24:  # CREATEBRUSHINDIRECT
                index=u32(pos+8);style=u32(pos+12);bgr=u32(pos+16)
                color=f'#{bgr&255:02X}{(bgr>>8)&255:02X}{(bgr>>16)&255:02X}'
                objects[index]=('brush',{'color':color if style==0 else 'none'})
            elif typ == 37 and size >= 12:  # SELECTOBJECT
                index=u32(pos+8)
                if index in objects:
                    role,obj=objects[index]
                    if role=='pen':selected_pen=obj
                    else:selected_brush=obj
                elif index & 0x80000000:  # GDI stock objects
                    stock=index&0x7FFFFFFF
                    if stock==5: selected_brush={'color':'none'}  # NULL_BRUSH
                    elif stock==8: selected_pen={'color':'none','width':0}  # NULL_PEN
                    elif stock==0: selected_brush={'color':'#FFFFFF'}
            elif typ == 40 and size >= 12:
                objects.pop(u32(pos+8),None)
            elif typ == 27 and size >= 16:  # MOVETOEX
                current=struct.unpack_from('<ii',blob,pos+8)
                if in_path:path.append('M'+xy(*current))
            elif typ == 54 and size >= 16:  # LINETO
                point=struct.unpack_from('<ii',blob,pos+8)
                if in_path:
                    if not path:path.append('M'+xy(*current))
                    path.append('L'+xy(*point))
                else:emit(['M'+xy(*current),'L'+xy(*point)],stroke=True)
                current=point
            elif typ in (86,87,88,89) and size >= 28:  # 16-bit polygons/lines/beziers
                n=u32(pos+24)
                if n > 200000 or 28 + 4*n > size:return None
                pts=[struct.unpack_from('<hh',blob,pos+28+4*i) for i in range(n)]
                if not pts:
                    pos += size;continue
                if typ in (86,87):
                    sub=['M'+xy(*pts[0])]+['L'+xy(*pt) for pt in pts[1:]]
                    if typ==86:sub.append('Z')
                    if in_path:path.extend(sub)
                    else:emit(sub,fill=typ==86,stroke=True)
                else:
                    sub=[]
                    if not in_path:sub=['M'+xy(*current)]
                    elif not path:sub=['M'+xy(*current)]
                    if typ==88:
                        if len(pts)%3:return None
                        for k in range(0,len(pts),3):
                            sub.append('C'+' '.join(xy(*p) for p in pts[k:k+3]))
                    else:  # POLYLINETO16
                        sub += ['L'+xy(*p) for p in pts]
                    if in_path:path.extend(sub)
                    else:emit(sub,stroke=True)
                current=pts[-1]
            elif typ == 59:  # BEGINPATH
                path=[];in_path=True
            elif typ == 60:  # ENDPATH
                in_path=False
            elif typ == 61:  # CLOSEFIGURE
                path.append('Z')
            elif typ in (62,63,64):  # FILLPATH, STROKEANDFILLPATH, STROKEPATH
                emit(path,fill=typ in (62,63),stroke=typ in (63,64))
                path=[];in_path=False
            elif typ == 58:  # ABORTPATH
                path=[];in_path=False
            elif typ in (1,2,3,4,5,6,7,8,11,12,13,17,18,20,21,22,24,25,
                         26,30,33,34,35,36,38,41,42,43,49,56,57,70,94):
                # Header, coordinate mode, state or non-drawing metadata.
                pass
            else:
                # No silently dropped drawing operations, e.g. bitmap blits
                # or text primitives. Let Windows GDI try those instead.
                return None
            pos += size
        if count >= 100000 or not nodes or not wnd_w or not wnd_h:return None
        view=f'{wnd_x:g} {wnd_y:g} {wnd_w:g} {wnd_h:g}'
        result=(f'<svg xmlns="http://www.w3.org/2000/svg" viewBox="{view}" '
                f'width="{fw}" height="{fh}" preserveAspectRatio="none">'
                + ''.join(nodes) + '</svg>')
        return result.encode('utf-8')
    except (ValueError, IndexError, struct.error, OverflowError, TypeError):
        return None
