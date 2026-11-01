#!/usr/bin/env python3
"""Build editable PowerPoint and a browser preview from one scene description.

Requires python-pptx, Pillow and lxml. No Office automation or macros.
"""
import argparse
import base64
import copy
import html
import json
import math
from pathlib import Path
import zipfile

from lxml import etree as ET
from PIL import Image
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.oxml.xmlchemy import OxmlElement
from pptx.util import Inches, Pt

ROOT = Path(__file__).resolve().parents[1]
W, H = 1600, 900
P = 'http://schemas.openxmlformats.org/presentationml/2006/main'
MC = 'http://schemas.openxmlformats.org/markup-compatibility/2006'
P14 = 'http://schemas.microsoft.com/office/powerpoint/2010/main'
P159 = 'http://schemas.microsoft.com/office/powerpoint/2015/09/main'


def scene(label, duration=0, morph=0.8, bg='10152E'):
    return dict(label=label, duration=duration, morph=morph, bg=bg, items=[])


def add(s, name, kind, x, y, w, h, **kw):
    item = dict(name='!!'+name, kind=kind, x=x, y=y, w=w, h=h, **kw)
    s['items'].append(item)
    return item


def make_scenes(c):
    C=c['colors']; out=[]
    def rect(s,n,x,y,w,h,color,rotation=0):
        return add(s,n,'rect',x,y,w,h,color=color,rotation=rotation)
    def text(s,n,t,x,y,w,h,size=32,color=None,bold=False,align=None,**kw):
        rtl=kw.pop('rtl',c['rtl'])
        return add(s,n,'text',x,y,w,h,text=t,size=size,color=color or C['text'],bold=bold,
                   align=align or ('right' if rtl else 'left'),rtl=rtl,**kw)
    def logo(s,x,y,sz=86,color=None):
        if c.get('logo_path'):
            add(s,'brand-logo','image',x,y,sz,sz,path=c['logo_path'])
            return
        rect(s,'logo-tile',x,y,sz,sz,color or '213771')
        # A replaceable two-piece geometric mark, rather than the source creator's mark.
        rect(s,'logo-left',x+sz*.23,y+sz*.23,sz*.23,sz*.54,'F6F7FF')
        rect(s,'logo-right',x+sz*.54,y+sz*.23,sz*.23,sz*.54,'F6F7FF')
        rect(s,'logo-cut',x+sz*.23,y+sz*.61,sz*.11,sz*.16,color or '213771')
    def hero(s,x,y,w,h):
        add(s,'hero','image',x,y,w,h,path=c.get('hero_path','assets/blue-folds.jpg'))
    def brand(s,x,y):
        logo(s,x,y,60)
        text(s,'brand',c['brand'],x+79,y+3,245,28,21,bold=True)
        text(s,'tagline',c['brand_tagline'],x+79,y+33,245,22,12,C['muted'])
    def cover(s):
        hero(s,64,64,704,772)
        rect(s,'title-panel',800,64,736,534,C['panel'])
        text(s,'title',c['title'],852,225,640,175,45,bold=True,line_colors=[C['text'],C['accent']])
        text(s,'subtitle',c['subtitle'],852,410,640,65,20,C['muted'])
        brand(s,825,735)
        rect(s,'presenter-divider',1210,730,2,72,C['rule'])
        text(s,'presenter',c['presenter'],1240,738,290,29,20,bold=True)
        text(s,'role',c['role'],1240,772,290,24,13,C['muted'])

    s=scene('01 · Logo on white',0.65,0.3,'F7F7FA')
    rect(s,'reveal',754,404,92,92,C['background'])
    logo(s,754,404,92); out.append(s)
    s=scene('02 · Diamond reveal',0.42,0.65,'F7F7FA')
    rect(s,'reveal',465,115,670,670,C['background'],45); logo(s,740,390,120); out.append(s)
    s=scene('03 · Mosaic resolves',0.45,0.65,C['background'])
    rect(s,'reveal',-1000,-1350,3600,3600,C['background'],90)
    for i,(x,y,w,h,col) in enumerate([(0,0,500,370,'172547'),(500,0,640,370,'1A254B'),(1140,0,460,530,'242744'),(0,370,500,530,'122045'),(500,370,640,530,'182A54'),(1140,530,460,370,'171D39')]):
        rect(s,'mosaic-'+str(i),x,y,w,h,col)
    logo(s,720,320,160); text(s,'brand',c['brand'],570,518,460,50,32,bold=True,align='center')
    text(s,'tagline',c['brand_tagline'],570,579,460,30,17,C['muted'],align='center'); out.append(s)
    s=scene('04 · Brand',0.65,0.8,C['background']); logo(s,720,320,160)
    text(s,'brand',c['brand'],570,518,460,50,32,bold=True,align='center')
    text(s,'tagline',c['brand_tagline'],570,579,460,30,17,C['muted'],align='center'); out.append(s)
    s=scene('05 · Cover assembles',0.45,0.9,C['background']); cover(s)
    for item in s['items']:
        if item['name'] in ['!!title','!!subtitle']: item['x']+=850
    out.append(s)
    s=scene('06 · Title cover',1.15,0.85,C['background']); cover(s); out.append(s)
    s=scene('07 · Outline transition',0.3,0.9,C['background'])
    hero(s,-750,0,750,900);rect(s,'title-panel',750,0,850,900,C['panel'])
    text(s,'outline',c['outline'],105,385,560,120,56,bold=True,align='center');out.append(s)
    for count in range(1,5):
        s=scene(f'{7+count:02} · Outline / {count}',0.15 if count<4 else 0.65,0.33,C['background'])
        rect(s,'title-panel',750,0,850,900,C['panel'])
        text(s,'outline',c['outline'],105,385,560,120,56,bold=True,align='center')
        for i,section in enumerate(c['sections'][:count]):
            y=225+i*118
            text(s,f'num-{i}',f'{i+1:02}',865,y,70,50,28,C['accent'],bold=True,rtl=False)
            text(s,f'row-{i}',section['title'],955,y,475,50,30,bold=True)
            text(s,f'plus-{i}','+',1460,y-5,50,60,35,align='center')
            rect(s,f'rule-{i}',865,y+72,645,1.5,C['rule'])
        out.append(s)

    def accordion(active=None):
        label='12 · Interactive outline' if active is None else f'{13+active:02} · Expanded / {active+1}'
        s=scene(label,0,0.85,C['background'])
        rect(s,'title-panel',0,0,1000,900,C['panel'])
        text(s,'outline',c['outline'],92,73,795,85,43,C['muted'],bold=True)
        hero(s,1032,92,520,716)
        if active==1:
            s['items'][-1]['path']=c.get('secondary_hero_path','assets/blue-bloom.jpg')
        logo(s,1420,672,132,C['background'])
        y=220
        for i,section in enumerate(c['sections']):
            target=12+i if active!=i else 11
            text(s,f'num-{i}',f'{i+1:02}',94,y,68,40,26,C['accent'],bold=True,rtl=False,link=target)
            text(s,f'row-{i}',section['title'],174,y,620,45,27,bold=True,link=target)
            text(s,f'plus-{i}','+',855,y-9,60,60,38,align='center',link=target)
            if active==i:
                text(s,f'desc-{i}',section['description'],96,y+58,785,115,20,C['muted'])
                y+=142
            rect(s,f'rule-{i}',94,y+63,820,1.5,C['rule'])
            y+=108
        if active is not None:
            text(s,'open-section',('ورود به بخش  ←' if c['rtl'] else 'Explore this section  →'),96,800,690,40,18,C['accent'],link=16+active)
        return s
    out.append(accordion())
    for i in range(4): out.append(accordion(i))
    # Four optional usable content layouts, beyond the video reference.
    for i,section in enumerate(c['sections']):
        s=scene(f'{17+i:02} · Content / {i+1}',0,0.85,C['background'])
        hero(s,1080,0,520,900)
        text(s,'section-label',f'{i+1:02}  /  '+section['title'],95,86,880,50,23,C['accent'],bold=True)
        text(s,'section-headline',section['headline'],95,232,885,180,47,bold=True)
        text(s,'section-body',section['body'],95,445,835,200,24,C['muted'])
        brand(s,95,748)
        text(s,'back',('بازگشت به فهرست  ←' if c['rtl'] else '←  Back to outline'),660,790,320,35,16,C['accent'],link=11)
        out.append(s)
    return out


def transition(slide, s):
    # Morph in PowerPoint 2016+; standard fade is the fallback for older consumers.
    root=slide._element
    ac=ET.Element('{'+MC+'}AlternateContent',nsmap={'mc':MC,'p159':P159,'p14':P14})
    ac.set('{'+MC+'}Ignorable','p14')
    choice=ET.SubElement(ac,'{'+MC+'}Choice',Requires='p159')
    fallback=ET.SubElement(ac,'{'+MC+'}Fallback')
    for parent,modern in [(choice,True),(fallback,False)]:
        attrs={'spd':'slow','advClick':'1','{'+P14+'}dur':str(round(s['morph']*1000))}
        if s['duration']: attrs['advTm']=str(round(s['duration']*1000))
        t=ET.SubElement(parent,'{'+P+'}transition',attrs)
        ET.SubElement(t,'{'+P159+'}morph' if modern else '{'+P+'}fade',{'option':'byObject'} if modern else {})
    # Schema order: cSld, clrMapOvr, transition / AlternateContent, timing, extLst.
    idx=len(root)
    for j,e in enumerate(root):
        if ET.QName(e).localname in ('timing','extLst'): idx=j;break
    root.insert(idx,ac)


def write_pptx(c,scenes,dest):
    prs=Presentation();prs.slide_width=Inches(13.333333);prs.slide_height=Inches(7.5)
    prs.core_properties.title='Blue Motion · Editable Presentation Template'
    prs.core_properties.subject='Reconstruction from visual reference, with native editable shapes and Morph transitions'
    prs.core_properties.author='Template Studio'
    slides=[prs.slides.add_slide(prs.slide_layouts[6]) for _ in scenes]
    def u(v): return Inches(v/120)
    for sl,s in zip(slides,scenes):
        sl.background.fill.solid();sl.background.fill.fore_color.rgb=RGBColor.from_string(s['bg'])
        for it in s['items']:
            x,y,w,h=[u(it[k]) for k in ['x','y','w','h']]
            if it['kind']=='image':
                sh=sl.shapes.add_picture(str(ROOT/it['path']),x,y,width=w,height=h)
                iw,ih=Image.open(ROOT/it['path']).size
                scale=max(it['w']/iw,it['h']/ih)
                cx=(1-it['w']/(iw*scale))/2;cy=(1-it['h']/(ih*scale))/2
                sh.crop_left=sh.crop_right=cx;sh.crop_top=sh.crop_bottom=cy
            elif it['kind']=='rect':
                sh=sl.shapes.add_shape(MSO_SHAPE.RECTANGLE,x,y,w,h)
                sh.fill.solid();sh.fill.fore_color.rgb=RGBColor.from_string(it['color']);sh.line.fill.background()
                sh.rotation=it.get('rotation',0)
            else:
                sh=sl.shapes.add_textbox(x,y,w,h);tf=sh.text_frame;tf.clear()
                tf.margin_left=tf.margin_right=0;tf.margin_top=tf.margin_bottom=0
                tf.word_wrap=True;tf.vertical_anchor=MSO_ANCHOR.TOP
                for j,line in enumerate(it['text'].split('\n')):
                    p=tf.paragraphs[0] if j==0 else tf.add_paragraph()
                    p.alignment={'left':PP_ALIGN.LEFT,'right':PP_ALIGN.RIGHT,'center':PP_ALIGN.CENTER}[it['align']]
                    p.space_before=Pt(0);p.space_after=Pt(2);p.line_spacing=1.18
                    pPr=p._p.get_or_add_pPr();pPr.set('rtl','1' if it.get('rtl') else '0')
                    r=p.add_run();r.text=line;r.font.name=c['font'];r.font.size=Pt(it['size']*.6)
                    line_colors=it.get('line_colors',[])
                    r.font.bold=it['bold'];r.font.color.rgb=RGBColor.from_string(line_colors[j] if j<len(line_colors) else it['color'])
                    # Explicit complex-script typeface for Persian runs.
                    rpr=r._r.get_or_add_rPr();cs=OxmlElement('a:cs');cs.set('typeface',c['font']);rpr.append(cs)
            sh.name=it['name']
            if 'link' in it: sh.click_action.target_slide=slides[it['link']]
        transition(sl,s)
        sl.notes_slide.notes_text_frame.text=(
            s['label']+'\n'+('Automatic build state. Keep this slide when retaining the motion sequence.' if s['duration'] else 'Click a section title or + to expand/collapse. Content links navigate inside this file.')+
            '\nMatching !! shape names preserve Morph object identity. Duplicate a slide before editing. Text, colors, geometry and image can be edited in PowerPoint.\n'+
            'Visual reconstruction from the supplied video; original transition settings and source assets are unknown.')
    # Presentation-wide theme colors and font are editable under Design > Variants.
    theme=prs.slide_masters[0].part.theme_part if hasattr(prs.slide_masters[0].part,'theme_part') else None
    prs.save(dest)
    # Add real template variant and align theme defaults with the supplied palette.
    with zipfile.ZipFile(dest) as z: parts={n:z.read(n) for n in z.namelist()}
    for name,data in list(parts.items()):
        if name.startswith('ppt/theme/theme') and name.endswith('.xml'):
            root=ET.fromstring(data);ns={'a':'http://schemas.openxmlformats.org/drawingml/2006/main'}
            for key,val in [('dk1',c['colors']['background']),('lt1',c['colors']['text']),('dk2',c['colors']['panel']),('accent1',c['colors']['accent'])]:
                node=root.find('.//a:clrScheme/a:'+key,ns)
                if node is not None:
                    for ch in list(node):node.remove(ch)
                    ET.SubElement(node,'{'+ns['a']+'}srgbClr',val=val)
            for node in root.findall('.//a:fontScheme//a:latin',ns):node.set('typeface',c['font'])
            parts[name]=ET.tostring(root,xml_declaration=True,encoding='UTF-8',standalone=True)
    with zipfile.ZipFile(dest,'w',zipfile.ZIP_DEFLATED) as z:
        for n,b in parts.items():z.writestr(n,b)
    parts['[Content_Types].xml']=parts['[Content_Types].xml'].replace(b'application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml',b'application/vnd.openxmlformats-officedocument.presentationml.template.main+xml')
    with zipfile.ZipFile(dest.with_suffix('.potx'),'w',zipfile.ZIP_DEFLATED) as z:
        for n,b in parts.items():z.writestr(n,b)


def write_preview(c,scenes,dest):

    preview_scenes=copy.deepcopy(scenes)
    for s in preview_scenes:
        for item in s['items']:
            if item['kind']=='text':
                colors=item.get('line_colors',[])
                item['html']='\n'.join('<span style="color:#'+(colors[j] if j<len(colors) else item['color'])+'">'+html.escape(line)+'</span>' for j,line in enumerate(item['text'].split('\n')))
            if item['kind']=='image':
                path=ROOT/item['path']
                mime='image/png' if path.suffix.lower()=='.png' else 'image/jpeg'
                item['dataUri']='data:'+mime+';base64,'+base64.b64encode(path.read_bytes()).decode()
    uri=''
    data=json.dumps(preview_scenes,ensure_ascii=False).replace('</','<\\/')
    lang='fa' if c['rtl'] else 'en'
    page='''<!DOCTYPE html><html lang="LANG"><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>Blue Motion · Preview</title><style>
*{box-sizing:border-box}body{margin:0;background:#080c1c;color:#e8efff;font-family:Arial,sans-serif}main{max-width:1400px;margin:auto;padding:30px 20px}header{display:flex;justify-content:space-between;align-items:center;gap:20px;margin-bottom:20px}h1{font-size:20px;font-weight:500;margin:0}small{color:#8293b7}button,select{font:14px Arial;border:1px solid #354566;color:#e8efff;background:#17223d;border-radius:6px;padding:9px 13px;cursor:pointer}button:hover{background:#26395a}.frame{width:100%;aspect-ratio:16/9;position:relative;overflow:hidden;box-shadow:0 30px 90px #0008;background:#10152e}.stage{width:1600px;height:900px;transform-origin:top left;position:absolute;overflow:hidden}.item{position:absolute;white-space:pre-wrap;line-height:1.18;transition:left var(--dur) cubic-bezier(.4,0,.2,1),top var(--dur) cubic-bezier(.4,0,.2,1),width var(--dur) cubic-bezier(.4,0,.2,1),height var(--dur) cubic-bezier(.4,0,.2,1),transform var(--dur) cubic-bezier(.4,0,.2,1),opacity var(--dur) ease,background var(--dur) ease;font-family:FONT,Arial,sans-serif}.item img{width:100%;height:100%;object-fit:cover}footer{display:flex;align-items:center;gap:10px;flex-wrap:wrap;margin-top:18px}.note{color:#8293b7;font-size:13px;line-height:1.7;margin-top:20px}.clickable{cursor:pointer}.clickable:hover{text-shadow:0 0 18px #79bfff88}
</style><main><header><h1>BLUE MOTION <small> / editable presentation</small></h1><small id="label"></small></header><div class="frame"><div class="stage"></div></div><footer><button id="play">▶ Play intro</button><button id="prev">←</button><button id="next">→</button><select id="select"></select><button id="fs">Full screen</button></footer><p class="note">Browser motion simulation · The downloadable PowerPoint uses native Morph transitions.<br>Click a section title or + in the interactive outline to open it. Use the arrow keys to move between slides.</p></main>
<script>const scenes=DATA,asset='ASSET',stage=document.querySelector('.stage'),frame=document.querySelector('.frame'),sel=document.querySelector('#select');let current=0,timer=null,playing=false;const nodes=new Map();scenes.forEach((s,i)=>sel.add(new Option(s.label,i)));
function scale(){stage.style.transform=`scale(${frame.clientWidth/1600})`}new ResizeObserver(scale).observe(frame);scale();
function show(idx,instant=false){clearTimeout(timer);current=Math.max(0,Math.min(scenes.length-1,idx));const s=scenes[current];stage.style.background='#'+s.bg;stage.style.setProperty('--dur',instant?'0s':s.morph+'s');const seen=new Set();s.items.forEach((it,z)=>{seen.add(it.name);let e=nodes.get(it.name);if(!e){e=document.createElement('div');e.className='item';e.style.opacity='0';stage.append(e);nodes.set(it.name,e)}e.innerHTML=it.kind==='image'?`<img src="${it.dataUri}">`:it.kind==='text'?(it.html||it.text.replace(/&/g,'&amp;').replace(/</g,'&lt;')):'';Object.assign(e.style,{left:it.x+'px',top:it.y+'px',width:it.w+'px',height:it.h+'px',background:it.kind==='rect'?'#'+it.color:'transparent',transform:`rotate(${it.rotation||0}deg)`,fontSize:(it.size||20)+'px',fontWeight:it.bold?'700':'400',color:'#'+it.color,textAlign:it.align||'left',direction:it.rtl?'rtl':'ltr',zIndex:z+1});e.onclick=it.link!==undefined?()=>{playing=false;show(it.link)}:null;e.classList.toggle('clickable',it.link!==undefined);requestAnimationFrame(()=>e.style.opacity='1')});nodes.forEach((e,n)=>{if(!seen.has(n)){e.style.opacity='0';e.style.pointerEvents='none'}else e.style.pointerEvents='auto'});sel.value=current;document.querySelector('#label').textContent=s.label;document.querySelector('#play').textContent=playing?'Ⅱ Pause':'▶ Play intro';if(playing&&s.duration)timer=setTimeout(()=>show(current+1),(s.duration+(instant?0:s.morph))*1000);else if(playing){playing=false;document.querySelector('#play').textContent='▶ Replay intro'}}
document.querySelector('#play').onclick=()=>{if(playing){playing=false;clearTimeout(timer);document.querySelector('#play').textContent='▶ Play intro'}else{playing=true;show(0,true)}};document.querySelector('#prev').onclick=()=>{playing=false;show(current-1)};document.querySelector('#next').onclick=()=>{playing=false;show(current+1)};sel.onchange=()=>{playing=false;show(+sel.value)};document.querySelector('#fs').onclick=()=>frame.requestFullscreen();document.addEventListener('keydown',e=>{if(e.key==='ArrowRight'){playing=false;show(current+1)}if(e.key==='ArrowLeft'){playing=false;show(current-1)}});window.showScene=(i)=>show(i,true);show(5,true);
</script></html>'''
    page=page.replace('LANG',lang).replace('FONT',c['font']).replace('DATA',data).replace('ASSET',uri)
    dest.write_text(page,encoding='utf-8')


def main():
    ap=argparse.ArgumentParser();ap.add_argument('--content',default='content.json');ap.add_argument('--name',default='blue-motion');args=ap.parse_args()
    c=json.loads((ROOT/args.content).read_text());assert len(c['sections'])==4,'The layout requires exactly four sections.'
    scenes=make_scenes(c);out=ROOT/'output';out.mkdir(exist_ok=True)
    write_pptx(c,scenes,out/(args.name+'.pptx'));write_preview(c,scenes,out/(args.name+'-preview.html'))
    (out/(args.name+'-scenes.json')).write_text(json.dumps(scenes,ensure_ascii=False,indent=2))
    print(f'Built {len(scenes)} slides: {out/args.name}.pptx / .potx / -preview.html')

if __name__=='__main__':main()
