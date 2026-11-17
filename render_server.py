#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Mac 端渲染微服务 —— 只干一件事：收字段 → 出 PPTX/PDF/PNG → 返回路径。

服务器（1.8GB 内存）跑 LibreOffice 会被 OOM 杀掉，这里把渲染挪到 16GB 的 Mac。
只监听 Tailscale 网卡（100.125.54.14:8021），不对公网暴露。

启动：
  python3 render_server.py
  # 或后台：nohup python3 render_server.py > logs/render.log 2>&1 &
"""
import os
import sys
import time
import json
import shutil
import hashlib
import subprocess
import threading
from typing import Optional

from fastapi import FastAPI
from fastapi.responses import JSONResponse, FileResponse
from pydantic import BaseModel

BASE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, BASE)

OUT_DIR = os.path.join(BASE, 'output')
LOG_DIR = os.path.join(BASE, 'logs')
os.makedirs(OUT_DIR, exist_ok=True)
os.makedirs(LOG_DIR, exist_ok=True)

# ⚠️ 必须用 7.6：macOS 版 LibreOffice 26.x 有 CJK 渲染 bug，中文字形全丢
# 7.6.7.2 实测正常（PDF 内嵌 SourceHanSansSC）
SOFFICE = '/Applications/LibreOffice76.app/Contents/MacOS/soffice'
if not os.path.exists(SOFFICE):
    SOFFICE = '/Applications/LibreOffice.app/Contents/MacOS/soffice'
PDFTOPPM = '/opt/homebrew/bin/pdftoppm'

_RENDER_LOCK = threading.Lock()   # soffice 不能并发跑

app = FastAPI(title='ParamSheet Render Service (Mac)')


class RenderIn(BaseModel):
    data: dict = {}
    lang: str = 'zh'
    pages: int = 4
    dpi: int = 110
    want_pdf: bool = True


def _cleanup(keep=200):
    try:
        items = []
        for fn in os.listdir(OUT_DIR):
            fp = os.path.join(OUT_DIR, fn)
            if os.path.isfile(fp):
                items.append((os.path.getmtime(fp), fp))
        if len(items) <= keep:
            return
        items.sort(reverse=True)
        for _, fp in items[keep:]:
            try:
                os.remove(fp)
            except Exception:
                pass
    except Exception:
        pass


@app.get('/health')
def health():
    return {
        'ok': True,
        'host': 'mac',
        'soffice': os.path.exists(SOFFICE),
        'pdftoppm': os.path.exists(PDFTOPPM),
        'tpl': os.path.exists(os.path.join(BASE, 'templates', 'template_zh.pptx')),
    }


@app.post('/render')
def render(body: RenderIn):
    t_start = time.time()
    try:
        import build_pptx
        f = dict(body.data or {})
        if not f:
            return JSONResponse({'ok': False, 'error': 'data 为空'}, 400)

        # 派生字段（与服务器端导出保持一致）
        if not f.get('freq_phase_pf'):
            fq = f.get('cur_freq') or ''
            f['freq_phase_pf'] = f.get('freq') or fq
        f['fc_lines'] = [f.get('fc_esp'), f.get('fc_prp'), f.get('fc_75'), f.get('fc_50')]

        tpl = os.path.join(BASE, 'templates',
                           'template_en.pptx' if body.lang == 'en' else 'template_zh.pptx')

        with _RENDER_LOCK:
            pptx = build_pptx.build(f, body.lang, tpl, OUT_DIR)
            pdf = pptx.rsplit('.', 1)[0] + '.pdf'
            if not os.path.exists(pdf):
                subprocess.run([SOFFICE, '--headless', '--convert-to', 'pdf',
                                '--outdir', OUT_DIR, pptx],
                               check=True, capture_output=True, timeout=180)

            stem = pdf.rsplit('.', 1)[0]
            pngs = []
            for i in range(1, max(1, body.pages) + 1):
                png = '%s_pg-%d.png' % (stem, i)
                if not os.path.exists(png):
                    subprocess.run([PDFTOPPM, '-png', '-r', str(body.dpi),
                                    '-f', str(i), '-l', str(i), pdf, '%s_pg' % stem],
                                   check=False, capture_output=True, timeout=120)
                if os.path.exists(png):
                    pngs.append(os.path.basename(png))

        _cleanup()
        return {
            'ok': True,
            'host': 'mac',
            'ms': int((time.time() - t_start) * 1000),
            'pdf': os.path.basename(pdf) if body.want_pdf else None,
            'pages': pngs,
        }
    except Exception as e:
        import traceback
        return JSONResponse({'ok': False, 'host': 'mac', 'error': str(e),
                             'trace': traceback.format_exc()[-1200:]}, 500)


@app.get('/file/{fn}')
def get_file(fn: str):
    """服务器从这里拉取渲染产物。"""
    # 防目录穿越
    if '/' in fn or '\\' in fn or fn.startswith('.'):
        return JSONResponse({'ok': False, 'error': 'bad name'}, 400)
    p = os.path.join(OUT_DIR, fn)
    if not os.path.exists(p):
        return JSONResponse({'ok': False, 'error': 'nf'}, 404)
    return FileResponse(p)


if __name__ == '__main__':
    import uvicorn
    host = os.environ.get('RENDER_HOST', '100.125.54.14')  # 只监听 Tailscale IP
    port = int(os.environ.get('RENDER_PORT', '8021'))
    print('Render service on http://%s:%d  (Tailscale only)' % (host, port))
    uvicorn.run(app, host=host, port=port, log_level='info')
