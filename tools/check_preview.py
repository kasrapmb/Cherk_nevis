from pathlib import Path
from playwright.sync_api import sync_playwright
from functools import partial
from http.server import ThreadingHTTPServer, SimpleHTTPRequestHandler
from threading import Thread

ROOT=Path(__file__).resolve().parents[1]
server=ThreadingHTTPServer(('127.0.0.1',0),partial(SimpleHTTPRequestHandler,directory=str(ROOT/'output')))
Thread(target=server.serve_forever,daemon=True).start()
with sync_playwright() as p:
    b=p.chromium.launch(executable_path='/usr/bin/chromium',headless=True,args=['--no-sandbox'])
    page=b.new_page(viewport={'width':1640,'height':1120},device_scale_factor=1)
    for locale,indices in [('',list(range(20))),('-fa',list(range(20)))]:
        errors=[]
        page.on('pageerror',lambda e: errors.append(str(e)))
        page.goto(f'http://127.0.0.1:{server.server_port}/blue-motion{locale}-preview.html')
        page.wait_for_timeout(300)
        for idx in indices:
            page.evaluate(f'window.showScene({idx})');page.wait_for_timeout(150)
            if idx in [5,11,13,16]:
                page.locator('.frame').screenshot(path=str(ROOT/f'output/preview{locale}-{idx+1:02}.png'))
            overflows=page.evaluate('''() => [...document.querySelectorAll('.item')].filter(e => getComputedStyle(e).opacity==='1' && e.textContent && (e.scrollHeight>e.clientHeight+3 || e.scrollWidth>e.clientWidth+3)).map(e=>e.textContent)''')
            assert not overflows,(locale,idx,overflows)
        page.evaluate('window.showScene(11)');page.wait_for_timeout(100)
        page.locator('.clickable').filter(has_text='+' if locale=='' else '+').first.click()
        page.wait_for_timeout(100)
        assert page.locator('#select').input_value()=='12'
        assert not errors,errors
        print(f'Preview {locale or "EN"}: 20 scenes, screenshots, text bounds and accordion navigation passed')
    b.close()
