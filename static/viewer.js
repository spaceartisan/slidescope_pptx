/* SlideScope: offline PowerPoint browser harness, no external JS dependencies. */
'use strict';
const $ = (id) => document.getElementById(id);
const SVG_NS = 'http://www.w3.org/2000/svg';
let imageFilterId=0;
const state = { deck: null, slide: 0, selected: null, scale: 1, fit: true, bounds: false, version: -1,
                references: [], compare: 'off' };
const els = { canvas: $('slideCanvas'), holder: $('stageHolder'), viewport: $('stageViewport'), thumbs: $('thumbs') };

function make(tag, cls, text) {
  const e = document.createElement(tag);
  if (cls) e.className = cls;
  if (text !== undefined && text !== null) e.textContent = String(text);
  return e;
}
function svgEl(tag, attrs = {}) {
  const e = document.createElementNS(SVG_NS, tag);
  for (const [name, val] of Object.entries(attrs)) e.setAttribute(name, String(val));
  return e;
}
function showToast(message) {
  $('toast').textContent = message; $('toast').classList.add('visible');
  clearTimeout(showToast.timer);
  showToast.timer = setTimeout(() => $('toast').classList.remove('visible'), 5000);
}
function presetPath(preset) {
  const maps = {
    triangle:'M50 0 L100 100 L0 100 Z', rtTriangle:'M0 0 L100 100 L0 100 Z',
    diamond:'M50 0 L100 50 L50 100 L0 50 Z', chevron:'M0 0 L72 0 L100 50 L72 100 L0 100 L28 50 Z',
    rightArrow:'M0 22 L63 22 L63 0 L100 50 L63 100 L63 78 L0 78 Z',
    leftArrow:'M100 22 L37 22 L37 0 L0 50 L37 100 L37 78 L100 78 Z',
    upArrow:'M22 100 L22 37 L0 37 L50 0 L100 37 L78 37 L78 100 Z',
    downArrow:'M22 0 L22 63 L0 63 L50 100 L100 63 L78 63 L78 0 Z',
    hexagon:'M25 0 L75 0 L100 50 L75 100 L25 100 L0 50 Z',
    pentagon:'M50 0 L100 37 L81 100 L19 100 L0 37 Z',
    parallelogram:'M25 0 L100 0 L75 100 L0 100 Z',
    trapezoid:'M25 0 L75 0 L100 100 L0 100 Z',
    star5:'M50 0 L62 34 L99 35 L69 57 L80 94 L50 72 L20 94 L31 57 L1 35 L38 34 Z',
    heart:'M50 92 C15 64 0 44 8 23 C15 1 40 -2 50 20 C60 -2 85 1 92 23 C100 44 85 64 50 92 Z',
    flowChartDecision:'M50 0 L100 50 L50 100 L0 50 Z',
    line:'M0 0 L100 100'
  };
  return maps[preset] || null;
}
function addShapeArt(node, obj) {
  const kind = obj.preset;
  const fill = obj.fill || 'transparent';
  if (fill.startsWith('linear-gradient')) {
    const bg = make('div','shape-gradient');
    bg.style.background = fill;
    bg.style.clipPath = kind === 'ellipse' ? 'ellipse(50% 50%)' : 'inset(0)';
    node.append(bg);
  }
  const border = obj.line;
  const stroke = border ? border.color || '#777' : 'none';
  // SVG vector-effect:non-scaling-stroke expresses the line width in display
  // pixels already; dividing by a shape's width made PowerPoint's 2px
  // borders nearly invisible on wide rectangles.
  const sw = border ? Math.max(0.25, border.width || 1) : 0;
  const attr = {fill:fill.startsWith('linear-gradient') ? 'none' : fill, stroke, 'stroke-width':sw, 'vector-effect':'non-scaling-stroke'};
  if (border && border.dash && border.dash !== 'solid') attr['stroke-dasharray']='4 3';
  let part;
  // OOXML 'frame' is a hollow picture border, not a solid rectangle.
  // Compute its normalized ring thickness from the adjustable preset guide.
  if (kind === 'frame') {
    const thickness=Math.max(0.3,Math.min(obj.w,obj.h)*(obj.frameAdj||.05));
    const tx=Math.min(49,100*thickness/Math.max(.1,obj.w));
    const ty=Math.min(49,100*thickness/Math.max(.1,obj.h));
    const outer=`M0 0 H100 V100 H0 Z`;
    const inner=`M${tx} ${ty} H${100-tx} V${100-ty} H${tx} Z`;
    const svg=svgEl('svg',{viewBox:'0 0 100 100',preserveAspectRatio:'none',class:'shape-svg','aria-hidden':'true'});
    svg.append(svgEl('path',{d:`${outer} ${inner}`,...attr,'fill-rule':'evenodd'}));
    node.append(svg);return;
  }
  if (obj.paths?.length) {
    const svg = svgEl('svg',{viewBox:'0 0 100 100',preserveAspectRatio:'none',class:'shape-svg','aria-hidden':'true'});
    for (const path of obj.paths) svg.append(svgEl('path',{d:path.d,fill:path.fill==='none'?'none':attr.fill,...(border?{'stroke':stroke,'stroke-width':sw,'vector-effect':'non-scaling-stroke'}:{})}));
    node.append(svg);
    return;
  }
  if (kind === 'ellipse' || kind === 'donut') {
    part = svgEl('ellipse', {cx:50,cy:50,rx:50,ry:50,...attr});
  } else if (kind === 'roundRect') {
    part = svgEl('rect',{x:0,y:0,width:100,height:100,rx:9,ry:9,...attr});
  } else if (kind === 'rect' || kind === 'flowChartProcess' || !presetPath(kind)) {
    part = svgEl('rect',{x:0,y:0,width:100,height:100,...attr});
  } else {
    part = svgEl('path',{d:presetPath(kind),...attr});
  }
  const svg = svgEl('svg',{viewBox:'0 0 100 100',preserveAspectRatio:'none',class:'shape-svg','aria-hidden':'true'});
  svg.append(part); node.append(svg);
}
function addConnectorArt(node, obj) {
  const line=obj.line;
  if(!line)return;
  // Use slide coordinates, including real zero extents. A normalized SVG
  // stretches arrowheads, and a zero-height viewport suppresses the line.
  const w=Math.max(1,obj.w),h=Math.max(1,obj.h),sw=Math.max(.5,line.width||1);
  const svg=svgEl('svg',{viewBox:`0 0 ${w} ${h}`,width:w,height:h,'aria-hidden':'true',class:'connector-svg'});
  Object.assign(svg.style,{position:'absolute',left:'0',top:'0',overflow:'visible'});
  const color=line.color||'#777';
  const path=svgEl('path',{d:`M0 0 L${obj.w} ${obj.h}`,fill:'none',stroke:color,'stroke-width':sw});
  const dash={dash:[4,3],sysDash:[3,1],dot:[1,3],sysDot:[1,1],lgDash:[8,3],dashDot:[4,3,1,3],lgDashDot:[8,3,1,3],lgDashDotDot:[8,3,1,3,1,3],sysDashDot:[3,1,1,1],sysDashDotDot:[3,1,1,1,1,1]}[line.dash];
  if(dash)path.setAttribute('stroke-dasharray',dash.map(v=>v*sw).join(' '));
  svg.append(path);
  const angle=Math.atan2(obj.h,obj.w)*180/Math.PI;
  function end(spec,start){
    if(!spec)return;
    const length=({sm:2,med:3,lg:5}[spec.length]||3)*sw;
    const width=({sm:2,med:3,lg:5}[spec.width]||3)*sw;
    const g=svgEl('g',{transform:`translate(${start?0:obj.w} ${start?0:obj.h}) rotate(${angle+(start?180:0)})`});
    let mark;
    if(spec.type==='oval')mark=svgEl('ellipse',{cx:0,cy:0,rx:length/2,ry:width/2,fill:color});
    else if(spec.type==='diamond')mark=svgEl('path',{d:`M${length/2} 0 L0 ${width/2} L${-length/2} 0 L0 ${-width/2} Z`,fill:color});
    else {
      const d=spec.type==='stealth'?`M0 0 L${-length} ${-width/2} L${-length*.65} 0 L${-length} ${width/2} Z`:
        `M${-length} ${-width/2} L0 0 L${-length} ${width/2}${spec.type==='arrow'?'':' Z'}`;
      mark=svgEl('path',{d,fill:spec.type==='arrow'?'none':color,stroke:spec.type==='arrow'?color:'none','stroke-width':sw,'stroke-linejoin':'miter'});
    }
    g.append(mark);svg.append(g);
  }
  end(line.headEnd,true);end(line.tailEnd,false);
  node.append(svg);
}
// Format the automatic numbered bullets encoded as a:buAutoNum.
function powerPointAutoNumber(n, format) {
  const alpha=(n, upper)=>{let out=''; for(let v=Math.max(1,n);v>0;v=Math.floor((v-1)/26))
    out=String.fromCharCode((upper?65:97)+(v-1)%26)+out;return out;};
  const roman=(n,upper)=>{if(n<1||n>3999)return String(n);let value=n,out='';
    for(const [num,label] of [[1000,'M'],[900,'CM'],[500,'D'],[400,'CD'],[100,'C'],[90,'XC'],[50,'L'],[40,'XL'],[10,'X'],[9,'IX'],[5,'V'],[4,'IV'],[1,'I']])
      while(value>=num){out+=label;value-=num;}return upper?out:out.toLowerCase();};
  const core=format.startsWith('alphaLc')?alpha(n,false):format.startsWith('alphaUc')?alpha(n,true):
    format.startsWith('romanLc')?roman(n,false):format.startsWith('romanUc')?roman(n,true):String(n);
  if(format.endsWith('ParenBoth'))return `(${core})`;
  if(format.endsWith('ParenR'))return `${core})`;
  return `${core}.`;
}
function renderText(textModel, root, opts = {}) {
  if (!textModel) return;
  const box = make('div','shape-text');
  const direction=textModel.direction||'horz';
  if(direction!=='horz') {
    // Vertical writing lays out against the available height, rather than
    // squeezing a horizontal paragraph into a narrow label's width.
    box.style.writingMode=direction==='mongolianVert'?'vertical-lr':'vertical-rl';
    box.style.textOrientation=direction==='eaVert'?'mixed':
      (direction==='wordArtVert'||direction==='wordArtVertRtl'?'upright':'sideways');
  }
  const textAngle=(textModel.textRotation||0)+(direction==='vert270'?180:0);
  if(textAngle)box.style.transform=`rotate(${textAngle}deg)`;
  if (opts.centerTitle) box.classList.add('center-title-text');
  // Respect DrawingML non-wrapping text frames.
  if (textModel.wrap === false) box.style.whiteSpace = 'pre';
  // PowerPoint's default text layout does not apply optional font-pair
  // kerning to ordinary slide text. Chromium's auto kerning changes glyph
  // advances and line wraps. Cached SmartArt has its own text metrics.
  if (!opts.smartArt) box.style.fontKerning = 'none';
  // Preserve crisp or blurred outer shadows inherited from slide masters.
  if (textModel.shadow) {
    const sh = textModel.shadow;
    box.style.textShadow = `${sh.x}px ${sh.y}px ${sh.blur}px ${sh.color}`;
  }
  // PowerPoint's normAutofit changes both glyph size and advance spacing.
  // Browsers using a scaled font size alone render tightly autofitted text
  // too narrow, changing wraps and sometimes clipping. Correct the advance
  // for tightly autofitted text frames (not normal body typography).
  if (textModel.fontScale > 0 && textModel.fontScale < 0.7)
    box.style.letterSpacing = `${((1-textModel.fontScale) * 0.30).toFixed(3)}px`;
  // OOXML can combine a normal-autofit text frame with locally sized runs
  // smaller than its inherited master paragraph style. In that case Office
  // preserves extra line leading. A straight CSS line-spacing reduction
  // collapses the nested paragraphs far too tightly.
  const hasSmallerDirectRuns = (textModel.paragraphs || []).some(p =>
    p.style?.size && p.runs?.some(r => r.text && r.size && r.size < p.style.size * 0.85));
  const preserveAutofitLeading = textModel.autoSize === 'TEXT_TO_FIT_SHAPE' &&
    textModel.fontScale >= 0.999 && hasSmallerDirectRuns;
  const [mt,mr,mb,ml] = textModel.margin || [0,0,0,0];
  box.style.padding = `${mt}px ${mr}px ${mb}px ${ml}px`;
  const vert = textModel.vertical;
  if (vert === 'middle') box.classList.add('vertical-middle');
  if (vert === 'bottom') box.classList.add('vertical-bottom');
  const numberingCounters=new Map();
  const scaledBlank = textModel.fontScale < 0.999 &&
    textModel.autoSize === "TEXT_TO_FIT_SHAPE" &&
    (textModel.paragraphs || []).some(p => !p.text && !(p.runs || []).some(r => r.text));
  (textModel.paragraphs || []).forEach((p, index) => {
    const para = make('p','text-paragraph');
    para.dataset.paragraphIndex=String(index);
    if (textModel.wrap === false) { para.style.whiteSpace='pre'; para.style.overflowWrap='normal'; }
    para.style.textAlign = {'center':'center','right':'right','justify':'justify'}[p.align] || 'left';
    const firstRun=p.runs?.find(r=>r.text);
    const fontSize=firstRun?.size || p.style?.size || opts.defaultFontSize || 24;
    para.style.fontSize=`${fontSize}px`;
    para.style.fontFamily=`"${String(firstRun?.font || p.style?.font || 'Arial').replaceAll('"','')}", Arial, sans-serif`;
    const before=p.spaceBefore ?? ((p.spaceBeforePct||0)*fontSize);
    const after=p.spaceAfter ?? ((p.spaceAfterPct||0)*fontSize);
    // Reposition the inherited leading of the first paragraph in standard,
    // non-autofit placeholders without changing subsequent paragraph flow.
    // This is a tested Office/browser metrics approximation.
    const transferFirstLeading = index===0 && opts.isPlaceholder &&
      textModel.autoSize == null && p.spaceBefore == null && before>0;
    if (before && !transferFirstLeading && (!scaledBlank || index > 0 || p.spaceBefore != null))
      para.style.marginTop=`${before}px`;
    if (after || transferFirstLeading) para.style.marginBottom=`${after+(transferFirstLeading?before:0)}px`;
    // PowerPoint subtitle placeholders use a roomier implicit line box than
    // ordinary HTML text (unless an explicit DrawingML line space exists).
    if (opts.subtitle && p.lineSpacing == null) para.style.lineHeight="1.44";
    if (typeof p.lineSpacingPx === 'number' && p.lineSpacingPx > 0)
      para.style.lineHeight=`${p.lineSpacingPx}px`;
    else if (typeof p.lineSpacing === 'number' && p.lineSpacing > 0 && p.lineSpacing < 4)
      para.style.lineHeight=String(p.lineSpacing);
    else if (p.lineSpacingReduction) {
      const ordinaryLeading=(textModel.fontScale < .8 ? 1.36 : 1.45)*(1-p.lineSpacingReduction);
      // Keep ~0.2em of the native line box's leading for explicit small runs
      // inside an unscaled normAutofit placeholder. This adjustment is
      // intentionally NOT applied to scaled autofit or ordinary text frames.
      para.style.lineHeight=String(Math.max(.85,ordinaryLeading + (preserveAutofitLeading ? .20 : 0)));
    }
    if(p.indentLeft !== null && p.indentLeft !== undefined) {
      para.style.paddingLeft=`${p.indentLeft}px`;
    } else if(p.level>0){
      // Fallback for texts without master paragraph styles.
      para.style.paddingLeft=`${p.level*23}px`;
    }
    const hasText=Boolean(p.text || p.runs?.some(r=>r.text));
    // A reduced-font normal-autofit frame with intentional empty paragraphs
    // needs the inherited paragraph leading between visible blocks, not
    // above its first visible paragraph. Keep explicit spacing directives.
    if (!hasText && scaledBlank && p.spaceBefore == null && before)
      para.style.marginBottom=`${after+before}px`;
    if(hasText && !p.bullet && !p.autoBullet && p.hanging != null) {
      para.style.textIndent=`${p.hanging}px`;
    }
    if(hasText && (p.bullet || p.autoBullet)){
      para.classList.add('has-bullet');
      let label=p.bullet || '•';
      if(p.autoBullet){
        const level=p.level || 0;
        if(!numberingCounters.has(level))numberingCounters.set(level,Math.max(1,p.numbering?.startAt||1)-1);
        const count=numberingCounters.get(level)+1;
        numberingCounters.set(level,count);
        for(const depth of [...numberingCounters.keys()])if(depth>level)numberingCounters.delete(depth);
        label=powerPointAutoNumber(count,p.numbering?.type||'arabicPeriod');
      }
      const bullet=make('span','bullet',label);
      // Legacy symbol fonts use character codes rather than Unicode glyphs.
      // Draw the common Wingdings square directly, independent of installed fonts.
      const symbolCode=label.length===1 ? (label.charCodeAt(0)&0xff) : -1;
      if (/^wingdings$/i.test(p.bulletFont||'') && [0xa7,0x6e,0xa0].includes(symbolCode)) {
        bullet.textContent='';
        const size=({167:.22,110:.54,160:.14})[symbolCode];
        const square=make('i');
        Object.assign(square.style,{display:'block',width:`${size}em`,height:`${size}em`,background:'currentColor',marginTop:`${.62-size/2}em`});
        bullet.append(square);
      } else if(p.bulletFont) bullet.style.fontFamily=`"${p.bulletFont.replaceAll('"','')}", Arial, sans-serif`;
      bullet.setAttribute('aria-hidden','true');
      const effectiveIndent=p.indentLeft??p.level*23;
      bullet.style.left=`${Math.max(0,effectiveIndent+(p.hanging??(-fontSize*.8)))}px`;
      bullet.style.color=p.bulletColor || firstRun?.color || p.style?.color || 'currentColor';
      bullet.style.fontSize=`${fontSize*(p.bulletSizePct||1)}px`;
      para.append(bullet);
    }
    for (const run of p.runs || []) {
      const span = make('span',null,run.text);
      // The run must inherit the paragraph's no-wrap setting. A fixed
      // pre-wrap rule on spans previously overrode the text-frame setting.
      span.style.whiteSpace='inherit';
      span.style.fontFamily = `"${String(run.font || p.style?.font || 'Arial').replaceAll('"','')}", Arial, sans-serif`;
      span.style.fontSize = `${run.size || p.style?.size || opts.defaultFontSize || 24}px`;
      if(run.baseline) {
        const size=run.size||p.style?.size||opts.defaultFontSize||24;
        // Office's script formatting retains the declared font size in
        // OOXML but draws a smaller glyph. Keep that original size in the
        // inspectable model and approximate script glyphs at 65% in CSS.
        span.style.fontSize=`${size*.65}px`;
        // Script glyphs must not inflate ordinary paragraph leading and
        // push every following bullet down the slide.
        span.style.lineHeight='0';
        span.style.verticalAlign=`${run.baseline*size}px`;
      }
      if (run.color || p.style?.color) span.style.color = run.color || p.style.color;
      if (run.bold) span.style.fontWeight = '700';
      if (run.italic) span.style.fontStyle = 'italic';
      if (run.kerning != null) {
        // OOXML kern is the minimum font size in 1/100 point at which
        // kerning takes effect; kern=0 explicitly disables it.
        const threshold = Number(run.kerning) * 96 / 7200;
        span.style.fontKerning = threshold > 0 && (run.size || fontSize) >= threshold ? 'normal' : 'none';
      }
      if (run.underline) span.style.textDecoration = 'underline';
      para.append(span);
    }
    // Empty PowerPoint paragraphs still have line-height and contribute to
    // layout. Dropping them collapsed the slide's intentional vertical gaps.
    if(!hasText) para.append(make('span','empty-paragraph','\u00a0'));
    box.append(para);
  });
  if (box.childNodes.length) root.append(box);
}

function renderImage(obj, node) {
  if (!obj.image) return;
  if (!obj.image.url) {
    const missing=make('div','image-unsupported');
    missing.title = `Unable to render ${obj.image.unsupported || 'image'}: ${obj.name}`;
    node.append(missing);
    return;
  }
  const wrap = make('div','shape-image');
  const img = make('img'); img.alt = obj.name || 'Slide image';
  img.src = obj.image.url;
  // a:lum brightness/contrast is applied to the *image pixels*, not to
  // surrounding text or shape geometry. CSS filters are a first approximation.
  const lum=obj.image.luminance;
  if(lum){
    // OOXML a:lum adjusts pixel intensity and contrast; ordinary CSS
    // brightness(.4) contrast(.37) washes dark slides gray. Approximate its
    // per-channel darkening as y=(1+contrast)*x + brightness*(40/255), then
    // express that transfer with native CSS filters (contrast first).
    const scale=Math.max(0,1+lum.contrast);
    const bias=lum.brightness*(40/255);
    const brightness=Math.max(0,scale+2*bias);
    const contrast=brightness>0?scale/brightness:1;
    img.style.filter=`contrast(${contrast}) brightness(${brightness})`;
  }
  const vml=obj.image.vmlAdjustment;
  if(vml && (vml.gain!==1 || vml.blackLevel!==0)) {
    const id=`image-adjust-${++imageFilterId}`;
    const defs=svgEl('svg',{width:0,height:0,'aria-hidden':'true'});
    defs.style.position='absolute';
    const filter=svgEl('filter',{id,'color-interpolation-filters':'sRGB'});
    const transfer=svgEl('feComponentTransfer');
    for(const channel of ['R','G','B'])transfer.append(svgEl(`feFunc${channel}`,{type:'linear',slope:vml.gain,intercept:(1-vml.gain)*.5+vml.blackLevel}));
    filter.append(transfer);defs.append(filter);node.append(defs);
    img.style.filter=`url(#${id})`;
  }

  const [l,t,r,b] = obj.image.crop || [0,0,0,0];
  const rx = Math.max(.05,1-l-r), ry = Math.max(.05,1-t-b);
  img.style.left = `${-l/rx*100}%`; img.style.top = `${-t/ry*100}%`;
  img.style.width = `${100/rx}%`; img.style.height = `${100/ry}%`;
  img.style.objectFit = 'fill';
  wrap.append(img); node.append(wrap);
}
function renderTable(obj, node) {
  const table=obj.table; if (!table) return;
  const grid=make('div','shape-table');
  const w=(table.columns || []).reduce((a,b)=>a+b,0)||obj.w;
  const h=(table.rows || []).reduce((a,b)=>a+b,0)||obj.h;
  const xEdges=[0],yEdges=[0],borderLines=new Map();
  table.columns.forEach(v=>xEdges.push(xEdges.at(-1)+v));
  table.rows.forEach(v=>yEdges.push(yEdges.at(-1)+v));
  grid.style.gridTemplateColumns = table.columns.map(x=>`${x/w*100}%`).join(' ');
  grid.style.gridTemplateRows = table.rows.map(x=>`${x/h*100}%`).join(' ');
  (table.cells || []).forEach((row,r)=>row.forEach((cell,c)=>{
    if (cell.spanned && !cell.mergeOrigin) return;
    const div=make('div','shape-table-cell');
    div.style.gridColumn = `${c+1} / span ${cell.mergeOrigin ? cell.spanWidth : 1}`;
    div.style.gridRow = `${r+1} / span ${cell.mergeOrigin ? cell.spanHeight : 1}`;
    div.style.background = cell.fill || '#FFFFFF';
    if(cell.borders) {
      div.style.border='none';
      const x1=xEdges[c],x2=xEdges[c+(cell.mergeOrigin?cell.spanWidth:1)];
      const y1=yEdges[r],y2=yEdges[r+(cell.mergeOrigin?cell.spanHeight:1)];
      const segments={left:[x1,y1,x1,y2],right:[x2,y1,x2,y2],top:[x1,y1,x2,y1],bottom:[x1,y2,x2,y2]};
      for(const [edge,line] of Object.entries(cell.borders)) {
        if(!line || line.color==='transparent' || !line.width) continue;
        const points=segments[edge],key=points.join(',');
        if(!borderLines.has(key)||borderLines.get(key).line.width<=line.width)
          borderLines.set(key,{points,line});
      }
    }
    renderText(cell.text,div,{defaultFontSize:16});
    grid.append(div);
  }));
  node.append(grid);
  // PowerPoint strokes are centered on cell boundaries and do not reduce
  // the text area. Draw each shared rule once rather than using CSS borders.
  const rules=svgEl('svg',{viewBox:`0 0 ${w} ${h}`,preserveAspectRatio:'none',class:'shape-svg table-rules'});
  for(const {points:[x1,y1,x2,y2],line} of borderLines.values()) {
    const attrs={x1,y1,x2,y2,stroke:line.color,'stroke-width':line.width};
    if(line.dash && line.dash!=='solid') attrs['stroke-dasharray']=line.dash.toLowerCase().includes('dot')?`${line.width} ${line.width*2}`:`${line.width*4} ${line.width*2}`;
    rules.append(svgEl('line',attrs));
  }
  node.append(rules);
}
function renderChart(obj, node) {
  const ch=obj.chart;
  if (!ch || !ch.series || !ch.series.length) { addShapeArt(node,{...obj,fill:'#eef1f3',preset:'rect'}); return; }
  const svg=svgEl('svg',{viewBox:'0 0 600 340',preserveAspectRatio:'none',class:'shape-chart','aria-label':ch.chartType+' chart'});
  svg.style.background=ch.background||'#FFFFFF';
  const textColor=ch.dark?'#ECF2F8':'#596e7d';
  const axisColor=ch.dark?'#8C9AA8':'#dbe1e7';
  const title=ch.title || '';
  if (title) { const titleElement=svgEl('text',{x:300,y:31,'text-anchor':'middle','font-size':19,'font-weight':700,fill:'#29364A'}); titleElement.textContent=title;svg.append(titleElement); }
  const categories=ch.categories || [];
  const series=ch.series.slice(0,8);
  const values=series.flatMap(s=>s.values).filter(v=>typeof v==='number' && Number.isFinite(v));
  const peak=Math.max(1,...values.map(v=>Math.abs(v)))*1.12;
  const colors=['#327C9A','#51A49A','#E09B50','#8E74B5','#D26A6A','#ABC37C','#607FCA','#CC7EAF'];
  const pie=ch.chartType.includes('pie') || ch.chartType.includes('doughnut');
  if (pie) {
    const arr=series[0].values.map(v=>Math.max(0,v||0));
    const total=arr.reduce((a,b)=>a+b,0)||1;
    let angle=-Math.PI/2;
    arr.forEach((val,i)=>{
      if (!val) return;
      const next=angle + val/total*Math.PI*2;
      const x1=300+Math.cos(angle)*108, y1=170+Math.sin(angle)*108;
      const x2=300+Math.cos(next)*108, y2=170+Math.sin(next)*108;
      const d=`M300 170 L${x1} ${y1} A108 108 0 ${next-angle>Math.PI?1:0} 1 ${x2} ${y2} Z`;
      svg.append(svgEl('path',{d,fill:colors[i%colors.length],stroke:'#fff','stroke-width':2}));
      angle=next;
    });
  } else {
    const isLine=ch.chartType.includes('line') || ch.chartType.includes('scatter') || ch.chartType.includes('area');
    const left=59,right=575,top=title?56:25,bottom=283;
    for(let k=0;k<=4;k++){
      const y=bottom-(bottom-top)*(k/4);
      svg.append(svgEl('line',{x1:left,y1:y,x2:right,y2:y,stroke:axisColor,'stroke-width':1}));
      const label=svgEl('text',{x:left-7,y:y+4,'text-anchor':'end','font-size':12,fill:textColor});label.textContent=(peak*k/4).toFixed(peak<10?1:0);svg.append(label);
    }
    const n=Math.max(1,categories.length,...series.map(s=>s.values.length));
    const step=(right-left)/n;
    series.forEach((s,j)=>{
      const width=Math.min(32,step*.7/series.length);
      const points=[];
      s.values.forEach((v,i)=>{
        if(typeof v!=='number') return;
        const x=left+step*(i+.5);
        const y=bottom-Math.max(0,v)/peak*(bottom-top);
        if(isLine) points.push(`${x},${y}`);
        else svg.append(svgEl('rect',{x:x+(j-(series.length-1)/2)*width-width/2,y,width,height:bottom-y,fill:colors[j%colors.length]}));
      });
      if(isLine && points.length>1) svg.append(svgEl('polyline',{points:points.join(' '),fill:'none',stroke:colors[j%colors.length],'stroke-width':3}));
    });
    categories.slice(0,14).forEach((cat,i)=>{
      const x=left+step*(i+.5);
      const label=svgEl('text',{x,y:bottom+21,'text-anchor':'middle','font-size':11,fill:textColor});label.textContent=String(cat).slice(0,13);svg.append(label);
    });
  }
  node.append(svg);
}
function renderObject(obj, inThumbnail = false) {
  const e=make('div','slide-object');
  e.style.left=`${obj.x}px`;e.style.top=`${obj.y}px`;
  e.style.width=`${Math.max(0.1,obj.w)}px`;e.style.height=`${Math.max(0.1,obj.h)}px`;
  if (obj.rotation || obj.flipH || obj.flipV) e.style.transform=`rotate(${obj.rotation||0}deg) scale(${obj.flipH?-1:1},${obj.flipV?-1:1})`;
  e.dataset.objectId=obj.id;e.dataset.objectName=obj.name;e.dataset.shapeType=obj.kind;
  e.dataset.source=obj.source;e.dataset.testid=inThumbnail?'thumbnail-object':'slide-object';
  e.classList.add(`source-${obj.source}`);
  if (obj.kind === 'image') renderImage(obj,e);
  else if (obj.preset === 'line' || obj.preset === 'straightConnector1') addConnectorArt(e,obj);
  else if(obj.kind === 'table') renderTable(obj,e);
  else if(obj.kind === 'chart') renderChart(obj,e);
  else if (obj.fillImage) {
    renderImage({...obj,image:obj.fillImage},e);
    // Overlay a transparent SVG outline, not a solid rectangle.
    if(obj.line)addShapeArt(e,{...obj,fill:'transparent'});
  }
  else addShapeArt(e,obj);
  if(obj.text && obj.kind!=='table') renderText(obj.text,e,{centerTitle:obj.placeholderType==='CENTER_TITLE',isPlaceholder:obj.kind==='placeholder',subtitle:obj.placeholderType==='SUBTITLE',smartArt:obj.source==='smartart'});
  if (!inThumbnail) {
    e.tabIndex=0;
    e.setAttribute('role','button');
    e.setAttribute('aria-label',`Inspect ${obj.name} (${obj.kind})`);
    e.addEventListener('click',event=>{event.stopPropagation();selectObject(obj.id);});
    e.addEventListener('keydown',event=>{if(event.key==='Enter'){event.stopPropagation();selectObject(obj.id);}});
  }
  return e;
}
function mountSlide(into, slide, thumbnail = false) {
  into.replaceChildren();
  if(!state.deck) return;
  into.style.width=`${state.deck.width}px`;into.style.height=`${state.deck.height}px`;
  // Setting backgroundImage='none' after the shorthand erases CSS gradient
  // backgrounds (the cause of dark grey slides in v0.4). Only override when
  // we actually have a background image from the PPTX theme/layout/slide.
  into.style.background=slide.background||'#fff';
  if (slide.backgroundImage) into.style.backgroundImage=`url("${slide.backgroundImage}")`;
  const tile=slide.backgroundImageTile;
  // OOXML a:tile preserves the source bitmap's physical size and scale;
  // stretching it across the entire slide distorts title-slide photography.
  into.style.backgroundSize=tile ? `${tile.w}px ${tile.h}px` : '100% 100%';
  into.style.backgroundPosition=tile ? `${tile.x}px ${tile.y}px` : 'center';
  into.style.backgroundRepeat=tile ? 'repeat' : 'no-repeat';
  const fragment=document.createDocumentFragment();
  for(const obj of slide.shapes||[]) fragment.append(renderObject(obj,thumbnail));
  into.append(fragment);
  if(!thumbnail) {
    const reference=make('img','reference-image');
    reference.id='referenceImage';reference.dataset.testid='reference-image';
    reference.alt='PowerPoint-rendered reference, not a SlideScope object';
    reference.draggable=false;
    const divider=make('div','reference-midline');divider.id='referenceDivider';
    into.append(reference,divider);
    refreshReference();
  }
}
function refreshReference(){
  const reference=$('referenceImage');if(!reference)return;
  const has=state.references.includes(state.slide+1);
  const active=has && state.compare!=='off';
  reference.style.display=active?'block':'none';
  reference.style.opacity=state.compare==='onion'?'0.5':'1';
  reference.style.clipPath=state.compare==='split'?'inset(0 0 0 50%)':'none';
  if(active)reference.src=`/api/reference/${state.slide+1}`;
  $('referenceDivider').style.display=active&&state.compare==='split'?'block':'none';
  $('compareMode').classList.toggle('is-on',active);
  $('compareMode').textContent=`COMPARE: ${state.compare.toUpperCase()}`;
}
function setCompare(mode){
  if(!['off','split','reference','onion'].includes(mode))return false;
  if(mode!=='off'&&!state.references.length){showToast('Load PowerPoint reference ZIP first.');return false;}
  state.compare=mode;refreshReference();return true;
}
function calcFit() {
  if (!state.deck) return 1;
  const rect=els.viewport.getBoundingClientRect();
  return Math.min(1.8,Math.max(.075,Math.min((rect.width-58)/state.deck.width,(rect.height-68)/state.deck.height)));
}
function applyScale() {
  if (!state.deck) return;
  if(state.fit) state.scale=calcFit();
  const scale=state.scale;
  els.holder.style.width=`${state.deck.width*scale}px`;
  els.holder.style.height=`${state.deck.height*scale}px`;
  els.canvas.style.transform=`scale(${scale})`;
  $('zoomLabel').textContent=state.fit?'FIT':`${Math.round(scale*100)}%`;
}
function updateThumbnails() {
  els.thumbs.replaceChildren();
  if(!state.deck) return;
  for(const [index,slide] of state.deck.slides.entries()) {
    const button=make('button','thumb-item');
    button.dataset.testid='slide-thumbnail';
    button.dataset.slideIndex=String(index+1);
    const wrapper=make('div','thumb-page-wrap');
    const canvas=make('div','thumb-canvas');
    mountSlide(canvas,slide,true);
    wrapper.append(canvas);
    const bottom=make('div','thumb-bottom');
    bottom.append(make('span','thumb-number',String(index+1).padStart(2,'0')),make('span','thumb-title',slide.title||`Slide ${index+1}`));
    button.append(wrapper,bottom);
    button.addEventListener('click',()=>goToSlide(index+1));
    els.thumbs.append(button);
  }
  scaleThumbnails();
}
function scaleThumbnails() {
  if(!state.deck) return;
  els.thumbs.querySelectorAll('.thumb-item').forEach(item=>{
    const wrap=item.querySelector('.thumb-page-wrap');
    wrap.style.aspectRatio=`${state.deck.width}/${state.deck.height}`;
    const canvas=item.querySelector('.thumb-canvas');
    canvas.style.transform=`scale(${wrap.clientWidth/state.deck.width})`;
  });
}
function goToSlide(n) {
  if(!state.deck || !state.deck.count) return false;
  state.slide=Math.min(Math.max(Math.round(n)-1,0),state.deck.count-1);
  state.selected=null;
  const slide=state.deck.slides[state.slide];
  mountSlide(els.canvas,slide);
  els.canvas.classList.toggle('canvas-show-bounds',state.bounds);
  $('slideNumber').textContent=`SLIDE ${String(state.slide+1).padStart(2,'0')} / ${String(state.deck.count).padStart(2,'0')}`;
  $('slideTitle').textContent=slide.title;
  $('slideDimensions').textContent=`${Math.round(state.deck.width)} × ${Math.round(state.deck.height)} PX  ·  ${slide.shapes.length} OBJECTS`;
  $('statusText').textContent='Browser preview ready';
  els.thumbs.querySelectorAll('.thumb-item').forEach(item=>{
    item.classList.toggle('is-active',Number(item.dataset.slideIndex)===state.slide+1);
    item.setAttribute('aria-current',Number(item.dataset.slideIndex)===state.slide+1?'true':'false');
  });
  els.thumbs.querySelector('.thumb-item.is-active')?.scrollIntoView({block:'nearest'});
  renderInspector(null);
  applyScale();
  return true;
}
function selectObject(id) {
  if(!state.deck) return false;
  const obj=state.deck.slides[state.slide].shapes.find(s=>s.id===id);
  if(!obj) return false;
  state.selected=obj.id;
  els.canvas.querySelectorAll('.slide-object').forEach(e=>e.classList.toggle('is-selected',e.dataset.objectId===id));
  renderInspector(obj);
  return true;
}
function renderInspector(obj) {
  const content=$('inspectorContent');
  content.replaceChildren();
  if(!obj){
    const empty=make('div','inspect-empty');
    empty.append(make('div','inspect-symbol','⌖'),make('h2',null,'Object details'),make('p',null,'Select a shape, image, or text box on a slide to inspect its coordinates and properties.'));
    content.append(empty);return;
  }
  const card=make('div','object-detail');
  card.append(make('div','detail-kicker',`${obj.kind.toUpperCase()} / ${obj.source.toUpperCase()}`),make('div','detail-title',obj.name));
  const section=make('div','detail-section');section.append(make('div','detail-section-title','GEOMETRY · PIXELS'));
  const grid=make('div','detail-grid');
  for(const [label,v] of [['X',obj.x],['Y',obj.y],['WIDTH',obj.w],['HEIGHT',obj.h],['ROTATION',`${obj.rotation}°`],['PRESET',obj.preset]]) {
    const m=make('div','detail-metric');m.append(make('span','metric-label',label),make('span','metric-value',v));grid.append(m);
  }
  section.append(grid);card.append(section);
  const sec2=make('div','detail-section');sec2.append(make('div','detail-section-title','PROPERTIES'));
  for(const [key,value] of [['Object ID',obj.id],['Fill',obj.fill||'Inherited / none'],['Stroke',obj.line?.color||'None'],['Shape type',obj.shapeType]]) {
    const row=make('div','detail-meta');row.append(make('span',null,key),make('span',null,String(value)));sec2.append(row);
  }
  card.append(sec2);
  const txt=obj.text?.paragraphs?.map(p=>p.text).filter(Boolean).join('\n') || obj.table?.cells?.map(row=>row.map(c=>c.text?.paragraphs?.map(p=>p.text).join(' ')).join(' | ')).join('\n');
  if(txt){const textSec=make('div','detail-section');textSec.append(make('div','detail-section-title','TEXT CONTENT'),make('pre','detail-text',txt));card.append(textSec);}
  content.append(card);
}
function updateDiagnostics() {
  const arr=state.deck?.warnings||[];
  $('warningCount').textContent=String(arr.length);
  $('issues').replaceChildren();
  if(!arr.length){$('issues').append(make('span','issue-ok',state.deck?'No known unsupported geometry detected':'No presentation loaded'));return;}
  for(const warning of arr.slice(0,30)) $('issues').append(make('div','issue',warning));
  if(arr.length>30) $('issues').append(make('div',null,`+${arr.length-30} more`));
}
async function reloadDeck() {
  const res=await fetch('/api/deck',{cache:'no-store'});
  if(!res.ok) throw new Error(`HTTP ${res.status}`);
  const deck=await res.json();
  state.deck=deck.count?deck:null;
  state.slide=0;state.selected=null;state.fit=true;
  $('emptyState').style.display=state.deck?'none':'flex';
  els.holder.style.display=state.deck?'block':'none';
  $('fileName').textContent=state.deck?.filename||'No presentation loaded';
  $('slideCount').textContent=String(state.deck?.count||0).padStart(2,'0');
  updateThumbnails();updateDiagnostics();
  if(state.deck)goToSlide(1);
  else {$('slideNumber').textContent='SLIDE — / —';$('statusText').textContent='Waiting for presentation';}
}
async function uploadFile(file) {
  if(!file)return;
  if(!file.name.toLowerCase().endsWith('.pptx')){showToast('Please select a .pptx presentation.');return;}
  $('statusText').textContent=`Loading ${file.name}…`;
  try{
    const data=await file.arrayBuffer();
    const response=await fetch('/api/upload',{method:'POST',headers:{'Content-Type':'application/octet-stream','X-Filename':encodeURIComponent(file.name)},body:data});
    const body=await response.json();
    if(!response.ok)throw new Error(body.error||'Unable to load PPTX');
    await reloadDeck();
    // A status poll sees this version on the next check. It will not reload unless changed.
    const status=await (await fetch('/api/status')).json();state.version=status.version;
  }catch(err){$('statusText').textContent='Import failed';showToast(err.message||String(err));}
  $('filePicker').value='';
}
async function refreshReferences(){
  const result=await (await fetch('/api/references',{cache:'no-store'})).json();
  state.references=result.slides||[];
  if(state.references.length&&state.compare==='off')state.compare='split';
  refreshReference();
  return result;
}
async function uploadReferences(file){
  if(!file)return;
  if(!file.name.toLowerCase().endsWith('.zip')){showToast('Select the PowerPoint-exported ZIP containing Slide1.PNG etc.');return;}
  try{
    const res=await fetch('/api/reference',{method:'POST',headers:{'X-Filename':encodeURIComponent(file.name),'Content-Type':'application/zip'},body:await file.arrayBuffer()});
    const data=await res.json();if(!res.ok)throw new Error(data.error||'Unable to import reference ZIP');
    await refreshReferences();setCompare('split');
    $('statusText').textContent=`PowerPoint references loaded: ${data.slides} slides`;
  }catch(e){showToast(e.message||String(e));}
  $('referencePicker').value='';
}
// Stable scripting interface for Playwright and VS Code agents.
window.PPTX_HARNESS={
  goToSlide,
  selectObject,
  getState:()=>({loaded:!!state.deck,slide:state.slide+1,slideCount:state.deck?.count||0,selectedId:state.selected,scale:state.scale,filename:state.deck?.filename||null}),
  getDeck:()=>state.deck,
  getObjects:(slideNumber=state.slide+1)=>state.deck?.slides?.[slideNumber-1]?.shapes||[],
  getWarnings:()=>state.deck?.warnings||[],
  inspect:(slideNumber,objectId)=>goToSlide(slideNumber)&&selectObject(objectId),
  setCompare, getReferenceState:()=>({count:state.references.length,slideNumbers:[...state.references],mode:state.compare})
};
$('filePicker').addEventListener('change',e=>uploadFile(e.target.files?.[0]));
$('referencePicker').addEventListener('change',e=>uploadReferences(e.target.files?.[0]));
$('compareMode').addEventListener('click',()=>{
 const modes=['off','split','reference','onion'];setCompare(modes[(modes.indexOf(state.compare)+1)%modes.length]);
});
$('nextSlide').addEventListener('click',()=>goToSlide(state.slide+2));
$('prevSlide').addEventListener('click',()=>goToSlide(state.slide));
$('toggleBounds').addEventListener('click',()=>{
  state.bounds=!state.bounds;
  els.canvas.classList.toggle('canvas-show-bounds',state.bounds);
  $('toggleBounds').classList.toggle('is-on',state.bounds);
});
$('zoomLabel').addEventListener('click',()=>{state.fit=true;applyScale()});
$('zoomIn').addEventListener('click',()=>{if(state.fit){state.scale=calcFit();state.fit=false;}state.scale=Math.min(3,state.scale*1.2);applyScale()});
$('zoomOut').addEventListener('click',()=>{if(state.fit){state.scale=calcFit();state.fit=false;}state.scale=Math.max(.1,state.scale/1.2);applyScale()});
els.canvas.addEventListener('click',(e)=>{if(e.target===els.canvas){state.selected=null;els.canvas.querySelectorAll('.is-selected').forEach(n=>n.classList.remove('is-selected'));renderInspector(null);}});
document.addEventListener('keydown',e=>{
  if(e.target.matches('input,textarea,[contenteditable]'))return;
  if(e.key==='ArrowRight'||e.key==='PageDown'){e.preventDefault();goToSlide(state.slide+2)}
  else if(e.key==='ArrowLeft'||e.key==='PageUp'){e.preventDefault();goToSlide(state.slide)}
  else if(e.key.toLowerCase()==='b'&&!e.ctrlKey){$('toggleBounds').click()}
  else if(e.key.toLowerCase()==='o'&&e.shiftKey){$('filePicker').click()}
  else if(e.key==='Escape'){state.selected=null;els.canvas.querySelectorAll('.is-selected').forEach(n=>n.classList.remove('is-selected'));renderInspector(null);}
});
let dragDepth=0;
document.addEventListener('dragenter',e=>{e.preventDefault();dragDepth++;$('dropOverlay').classList.add('visible')});
document.addEventListener('dragover',e=>e.preventDefault());
document.addEventListener('dragleave',e=>{e.preventDefault();dragDepth=Math.max(0,dragDepth-1);if(!dragDepth)$('dropOverlay').classList.remove('visible')});
document.addEventListener('drop',e=>{e.preventDefault();dragDepth=0;$('dropOverlay').classList.remove('visible');uploadFile(e.dataTransfer?.files?.[0]);});
window.addEventListener('resize',()=>{applyScale();scaleThumbnails()});
if('ResizeObserver' in window){const ro=new ResizeObserver(()=>{applyScale();scaleThumbnails()});ro.observe(els.viewport);ro.observe(els.thumbs);}
(async()=>{
  try{await reloadDeck();await refreshReferences();const r=await (await fetch('/api/status')).json();state.version=r.version??-1;}
  catch(e){showToast('Unable to connect to local Python server: '+e.message);}
  setInterval(async()=>{
    try{const stat=await (await fetch('/api/status',{cache:'no-store'})).json();if(stat.version!==undefined&&stat.version!==state.version){state.version=stat.version;await reloadDeck();}}
    catch(_e){/* page may remain open as server restarts */}
  },1800);
})();
