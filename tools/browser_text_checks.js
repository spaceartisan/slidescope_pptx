/* Optional development checks. Load after viewer.js in the running app,
   then call await runSlideScopeTextChecks(). No added runtime dependencies. */
async function runSlideScopeTextChecks() {
  const host=document.createElement('div');
  Object.assign(host.style,{position:'fixed',left:'0',top:'0',width:'500px',height:'500px',zIndex:'9999',background:'white'});
  document.body.append(host);
  const results=[];
  function check(condition,message) { if(!condition)throw Error(message);results.push(message); }
  function draw({wrap=true,direction='horz',runs,width=60,height=160}) {
    const root=document.createElement('div');
    Object.assign(root.style,{position:'absolute',width:width+'px',height:height+'px'});host.append(root);
    const defaults={font:'Arial',size:20,color:'#000'};
    renderText({margin:[0,0,0,0],wrap,direction,fontScale:1,vertical:'top',paragraphs:[{
      text:runs.map(r=>r.text).join(''),align:'left',level:0,style:defaults,
      runs:runs.map(r=>({...defaults,...r}))
    }]},root);
    return root;
  }
  function charBox(span,index=0) {const range=document.createRange();range.setStart(span.firstChild,index);range.setEnd(span.firstChild,index+1);return range.getBoundingClientRect();}
  try {
    const nowrap=draw({wrap:false,runs:[{text:'Alpha Beta Gamma'}]});
    const wrapped=draw({wrap:true,runs:[{text:'Alpha Beta Gamma'}]});
    const script=draw({width:200,runs:[{text:'x'},{text:'2',baseline:.3},{text:'3',baseline:-.25},{text:'2',size:13}]});
    const vertical=draw({direction:'vert',runs:[{text:'ABCDE'}]});
    const normal=draw({width:200,runs:[{text:'x23'}]});
    await document.fonts.ready;
    await new Promise(resolve=>requestAnimationFrame(()=>requestAnimationFrame(resolve)));
    const n=nowrap.querySelector('span'),w=wrapped.querySelector('span');
    check(Math.abs(charBox(n,0).top-charBox(n,11).top)<1,'No-wrap run stays on one line');
    check(charBox(w,11).top>charBox(w,0).top+5,'Wrapping remains enabled when requested');
    const [base,sup,sub,control]=script.querySelectorAll('span');
    // Compare equal-size glyphs to isolate the baseline shift from the
    // smaller font's ascent/descent metrics.
    check(charBox(sup).top<charBox(control).top-4,'Superscript baseline rises');
    check(charBox(sub).bottom>charBox(control).bottom+4,'Subscript baseline descends');
    check(charBox(sup).height<charBox(base).height,'Script glyphs render smaller');
    check(Math.abs(script.querySelector('p').getBoundingClientRect().height-normal.querySelector('p').getBoundingClientRect().height)<1,'Script offsets preserve paragraph leading');
    const v=vertical.querySelector('span');
    check(charBox(v,4).top>charBox(v,0).top+20,'Vertical text advances down the label');
    check(Math.abs(charBox(v,4).left-charBox(v,0).left)<1,'Vertical text remains in one column');
    return results;
  } finally {host.remove();}
}
