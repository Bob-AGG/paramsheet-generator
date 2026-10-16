#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
参数单生成器 V2 — FastAPI 服务
- GET  /               托管前端页面
- GET  /api/templates  列出可用语言模板
- POST /api/build      数据 JSON + lang -> 生成 pptx -> 返回文件
"""
import io
import json
import os
import sys
import tempfile

from fastapi import FastAPI, Request
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse, HTMLResponse, JSONResponse
from fastapi.staticfiles import StaticFiles

import build_pptx
import datasheet_parser
import datasheet_ai
import base64
import re as _re

HERE = os.path.dirname(os.path.abspath(__file__))
OUTPUT = os.path.join(HERE, "output")
IMAGES = os.path.join(HERE, "images")

app = FastAPI(title="ParamSheet Generator")
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_methods=["*"],
    allow_headers=["*"],
)


@app.get("/api/templates")
def templates():
    return {"templates": [
        {"lang": "en", "label": "English", "file": "template_map.en.json"},
        {"lang": "zh", "label": "中文", "file": "template_map.zh.json"},
    ]}


def pptx_to_pdf(pptx_path, outdir):
    """LibreOffice headless 转 PDF"""
    soffice = None
    for c in ("/usr/bin/soffice", "/Applications/LibreOffice.app/Contents/MacOS/soffice", "soffice"):
        if os.path.exists(c) or c == "soffice":
            soffice = c
            break
    if not soffice:
        return None
    cmd = [soffice, "--headless", "--norestore", "--convert-to", "pdf", "--outdir", outdir, pptx_path]
    try:
        import subprocess
        subprocess.run(cmd, timeout=180, check=True,
                       stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    except Exception as e:
        print("PDF 转换失败:", e, file=sys.stderr)
        return None
    base = os.path.splitext(os.path.basename(pptx_path))[0] + ".pdf"
    pdf = os.path.join(outdir, base)
    return pdf if os.path.exists(pdf) else None


@app.post("/api/build")
async def build(request: Request):
    try:
        body = await request.json()
    except Exception:
        return JSONResponse({"ok": False, "error": "JSON 解析失败"}, status_code=400)
    data = body.get("data") or {}
    lang = body.get("lang") or "en"
    title_sub = body.get("title_sub") or ""
    if title_sub:
        # slide1 顶部小标题（型号字，如 AS550E5）
        data["__title_sub__"] = title_sub
    if not data.get("model"):
        return JSONResponse({"ok": False, "error": "缺少 model（机组型号）"}, status_code=400)
    # 本地手选图片（dataURL）-> 写临时文件，供 build_pptx 使用
    tmp_files = []
    for key, field in (("__img_open_data", "__img_open__"), ("__img_silent_data", "__img_silent__")):
        v = data.pop(key, None)
        if v and isinstance(v, str) and v.startswith("data:image"):
            try:
                b64 = v.split(",", 1)[1]
                raw = base64.b64decode(b64)
                tf = tempfile.NamedTemporaryFile(delete=False, suffix=".png")
                tf.write(raw); tf.close()
                data[field] = tf.name
                tmp_files.append(tf.name)
            except Exception:
                pass
    fmt = (body.get("format") or "pptx").lower()
    try:
        os.makedirs(OUTPUT, exist_ok=True)
        out = build_pptx.build(data, lang, None, OUTPUT)
    except Exception as e:
        return JSONResponse({"ok": False, "error": str(e)}, status_code=500)
    if fmt == "pdf":
        pdf = pptx_to_pdf(out, OUTPUT)
        if pdf:
            return FileResponse(pdf, media_type="application/pdf",
                                filename=os.path.basename(pdf))
        return JSONResponse({"ok": False, "error": "PDF 转换失败（LibreOffice 不可用）"}, status_code=500)
    fname = os.path.basename(out)
    return FileResponse(
        out,
        media_type="application/vnd.openxmlformats-officedocument.presentationml.presentation",
        filename=fname,
    )


@app.post("/api/parse")
async def parse_datasheet_api(request: Request):
    """上传 Datasheet（PDF/Excel/TXT）-> 自动提取字段。

    表单：multipart/form-data，字段名 file；可选 model_hint。
    """
    try:
        form = await request.form()
    except Exception:
        return JSONResponse({"ok": False, "error": "需 multipart/form-data"}, status_code=400)
    up = form.get("file")
    if up is None:
        return JSONResponse({"ok": False, "error": "未收到文件"}, status_code=400)
    fname = getattr(up, "filename", "upload")
    content = await up.read()
    if not content:
        return JSONResponse({"ok": False, "error": "文件为空"}, status_code=400)
    if len(content) > 30 * 1024 * 1024:
        return JSONResponse({"ok": False, "error": "文件过大（>30MB）"}, status_code=400)
    model_hint = form.get("model_hint") or ""
    engine = (form.get("engine") or "ai").strip().lower()   # ai(默认) | rule
    try:
        if engine == "rule":
            text = datasheet_parser.extract_text(fname, content)
            fields = datasheet_parser.parse_text(text, model_hint)
            meta = {"chars": len(text), "ai": False}
        else:
            fields, meta = datasheet_ai.parse_datasheet_ai(fname, content, model_hint)
    except Exception as e:
        return JSONResponse({"ok": False, "error": "解析失败: %s" % e}, status_code=500)
    return {"ok": True, "filename": fname, "fields": fields,
            "found": len(fields), "meta": meta}


@app.post("/api/upload-image")
async def upload_image(request: Request):
    """上传机组照片（开架/静音）。

    表单：file + model + kind(open|silent)。按 型号_OPEN.png / 型号_SILENT.png 存 images/。
    """
    try:
        form = await request.form()
    except Exception:
        return JSONResponse({"ok": False, "error": "需 multipart/form-data"}, status_code=400)
    up = form.get("file")
    model = (form.get("model") or "").strip()
    kind = (form.get("kind") or "open").strip().lower()
    if up is None or not model:
        return JSONResponse({"ok": False, "error": "缺少 file 或 model"}, status_code=400)
    if kind not in ("open", "silent"):
        return JSONResponse({"ok": False, "error": "kind 必须是 open/silent"}, status_code=400)
    content = await up.read()
    if not content:
        return JSONResponse({"ok": False, "error": "文件为空"}, status_code=400)
    # 安全文件名（去掉路径分隔符）
    safe = "".join(c for c in model if c not in "/\\:*?\"<>|")
    os.makedirs(IMAGES, exist_ok=True)
    dst = os.path.join(IMAGES, "%s_%s.png" % (safe, kind.upper()))
    with open(dst, "wb") as f:
        f.write(content)
    return {"ok": True, "file": "%s_%s.png" % (safe, kind.upper()), "kind": kind}


@app.get("/api/images")
def list_images(model: str = ""):
    """列出某型号的可用图片（自动匹配结果）。"""
    import image_utils
    if not model:
        return {"ok": True, "images": {}}
    found = image_utils.find_images(IMAGES, model)
    return {"ok": True, "images": {k: (os.path.basename(v) if v else None)
                                    for k, v in found.items()}}


@app.get("/health")
def health():
    return {"ok": True, "langs": ["en", "zh"], "parse": True, "upload": True}


# 前端静态页
app.mount("/static", StaticFiles(directory=HERE), name="static")


@app.get("/", response_class=HTMLResponse)
def index():
    p = os.path.join(HERE, "index.html")
    with io.open(p, encoding="utf-8") as f:
        return HTMLResponse(f.read())


if __name__ == "__main__":
    import uvicorn

    uvicorn.run(app, host="127.0.0.1", port=8010)
