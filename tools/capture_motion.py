"""Render an explicitly labelled browser simulation, not a PowerPoint recording."""
from pathlib import Path
from functools import partial
from http.server import ThreadingHTTPServer,SimpleHTTPRequestHandler
from threading import Thread
import subprocess
from playwright.sync_api import sync_playwright

ROOT=Path(__file__).resolve().parents[1]
class Quiet(SimpleHTTPRequestHandler):
    def log_message(self,*args):pass
server=ThreadingHTTPServer(('127.0.0.1',0),partial(Quiet,directory=str(ROOT/'output')))
Thread(target=server.serve_forever,daemon=True).start()
with sync_playwright() as p:
    b=p.chromium.launch(executable_path='/usr/bin/chromium',headless=True,args=['--no-sandbox'])
    page=b.new_page(viewport={'width':1280,'height':720},device_scale_factor=1)
    page.goto(f'http://127.0.0.1:{server.server_port}/blue-motion-preview.html')
    page.add_style_tag(content='main{padding:0;margin:0}header,footer,.note{display:none}.frame{box-shadow:none}.item{transition:none!important}')
    page.evaluate('''() => { stage.innerHTML=''; window.frameNodes=new Map();
      window.drawFrame=(a,b,t)=>{const A=scenes[a],B=scenes[b];const ease=t*t*(3-2*t),seen=new Set();
      stage.style.background='#'+(t<.5?A.bg:B.bg);const ai=new Map(A.items.map(x=>[x.name,x])),bi=new Map(B.items.map(x=>[x.name,x]));
      const names=[...new Set([...ai.keys(),...bi.keys()])];
      for(const name of names){let x=ai.get(name),y=bi.get(name);let opacity=x&&y?1:x?1-ease:ease;x=x||y;y=y||x;
      let e=frameNodes.get(name);if(!e){e=document.createElement('div');e.className='item';stage.append(e);frameNodes.set(name,e)}seen.add(name);
      const m=(k)=>(x[k]||0)+((y[k]||0)-(x[k]||0))*ease,it=y;
      e.innerHTML=it.kind==='image'?`<img src="${it.dataUri}">`:it.kind==='text'?(it.html||it.text.replace(/&/g,'&amp;').replace(/</g,'&lt;')):'';
      Object.assign(e.style,{left:m('x')+'px',top:m('y')+'px',width:m('w')+'px',height:m('h')+'px',transform:`rotate(${m('rotation')}deg)`,opacity,fontSize:m('size')+'px',color:'#'+it.color,background:it.kind==='rect'?'#'+it.color:'transparent',fontWeight:it.bold?'700':'400',textAlign:it.align||'left',direction:it.rtl?'rtl':'ltr',zIndex:B.items.findIndex(k=>k.name===name)+2});}
      frameNodes.forEach((e,n)=>{if(!seen.has(n))e.style.opacity=0});};
    }''')
    # Automatic intro followed by the two accordion openings observed in the video.
    sc=page.evaluate('scenes');segments=[];clock=0
    for i in range(14):
        if i>0:
            segments.append((clock,clock+sc[i]['morph'],i-1,i));clock+=sc[i]['morph']
        hold=sc[i]['duration'] or 1.1
        segments.append((clock,clock+hold,i,i));clock+=hold
    fps=24;frames=round(clock*fps)
    cmd=['ffmpeg','-hide_banner','-loglevel','error','-y','-f','image2pipe','-vcodec','png','-r',str(fps),'-i','-','-an','-c:v','libx264','-preset','fast','-crf','20','-pix_fmt','yuv420p','-movflags','+faststart',str(ROOT/'output/motion-preview.mp4')]
    proc=subprocess.Popen(cmd,stdin=subprocess.PIPE)
    for k in range(frames):
        t=k/fps
        start,end,a,z=next(s for s in segments if s[0]<=t<s[1])
        frac=(t-start)/(end-start) if a!=z else 1
        page.evaluate('([a,b,t])=>window.drawFrame(a,b,t)',[a,z,frac])
        proc.stdin.write(page.locator('.frame').screenshot(type='png'))
        if k%120==0:print(f'Frame {k}/{frames}',flush=True)
    proc.stdin.close();assert proc.wait()==0
    b.close()
    print(f'Saved motion-preview.mp4: {frames} frames, {clock:.2f}s, 1280×720, browser simulation')
