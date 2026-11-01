"""Validate package structure, Morph states, editable content and internal links."""
from pathlib import Path
import posixpath
import zipfile
from lxml import etree as E
from pptx import Presentation

ROOT=Path(__file__).resolve().parents[1]
NS={'p':'http://schemas.openxmlformats.org/presentationml/2006/main',
    'a':'http://schemas.openxmlformats.org/drawingml/2006/main',
    'r':'http://schemas.openxmlformats.org/package/2006/relationships',
    'p159':'http://schemas.microsoft.com/office/powerpoint/2015/09/main'}
for f in sorted((ROOT/'output').glob('*.pptx')):
    prs=Presentation(f)
    assert len(prs.slides)==20
    assert len(prs.slides[5].shapes)>10,'Cover must have editable elements'
    assert any(s.has_text_frame and s.text for s in prs.slides[5].shapes)
    with zipfile.ZipFile(f) as z:
        assert z.testzip() is None
        names=set(z.namelist());links=0
        for n in names:
            if n.endswith(('.xml','.rels')):
                root=E.fromstring(z.read(n))
                if n.endswith('.rels'):
                    base=posixpath.dirname(posixpath.dirname(n))
                    for r in root:
                        if r.get('TargetMode')=='External':raise AssertionError('Unexpected external relationship')
                        target=r.get('Target')
                        path=posixpath.normpath(posixpath.join(base,target)) if not target.startswith('/') else target[1:]
                        assert path in names,(n,path)
                        if r.get('Type','').endswith('/slide'):links+=1
        for i in range(1,21):
            root=E.fromstring(z.read(f'ppt/slides/slide{i}.xml'))
            assert len(root.findall('.//p159:morph',NS))==1
            assert len(root.findall('.//p:fade',NS))==1
            shape_names=root.xpath('//p:cNvPr/@name',namespaces=NS)
            assert len(shape_names)==len(set(shape_names)),(i,'Duplicate Morph IDs')
        assert not any('vbaProject' in n for n in names)
    with zipfile.ZipFile(f.with_suffix('.potx')) as z:
        assert b'presentationml.template.main+xml' in z.read('[Content_Types].xml')
    print(f'{f.name}: 20 slides, 20 Morph/fade pairs, editable text/shapes, valid relationships, macro-free; template type correct')
