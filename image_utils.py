#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
图片处理：等比缩放 + 居中裁切 + （可选）白底抠透明
用于把厂家机组照片填进 pptx 图片占位框，保证不变形。
"""
import io
import os

from PIL import Image


def _remove_bg_flood(img, thresh=245):
    """只抠「与图像边缘连通的」近白背景（不碰图内部白色）。

    用从四条边缘开始的多源 flood fill，避免把机器上的白色区域也抠掉。
    """
    from collections import deque
    img = img.convert("RGBA")
    w, h = img.size
    px = img.load()

    def is_bg(x, y):
        r, g, b, a = px[x, y]
        return a > 0 and r >= thresh and g >= thresh and b >= thresh

    visited = bytearray(w * h)
    dq = deque()
    for x in range(w):
        for y in (0, h - 1):
            if not visited[y * w + x] and is_bg(x, y):
                visited[y * w + x] = 1
                dq.append((x, y))
    for y in range(h):
        for x in (0, w - 1):
            if not visited[y * w + x] and is_bg(x, y):
                visited[y * w + x] = 1
                dq.append((x, y))
    while dq:
        x, y = dq.popleft()
        r, g, b, a = px[x, y]
        px[x, y] = (r, g, b, 0)
        for nx, ny in ((x+1, y), (x-1, y), (x, y+1), (x, y-1)):
            if 0 <= nx < w and 0 <= ny < h and not visited[ny * w + nx] and is_bg(nx, ny):
                visited[ny * w + nx] = 1
                dq.append((nx, ny))
    return img

def _remove_white_bg(img, thresh=245):
    """把接近纯白/近白的背景像素变透明（适合白底厂家图）"""
    img = img.convert("RGBA")
    px = img.load()
    w, h = img.size
    for y in range(h):
        for x in range(w):
            r, g, b, a = px[x, y]
            if r >= thresh and g >= thresh and b >= thresh:
                px[x, y] = (r, g, b, 0)
    return img


def _crop_to_ratio(img, target_ratio, anchor="center"):
    """居中裁切到目标宽高比（不缩放，只裁边）"""
    w, h = img.size
    cur = w / h
    if abs(cur - target_ratio) < 1e-3:
        return img
    if cur > target_ratio:
        # 太宽 -> 裁左右
        new_w = int(round(h * target_ratio))
        if anchor == "center":
            left = (w - new_w) // 2
        else:
            left = 0
        return img.crop((left, 0, left + new_w, h))
    else:
        # 太高 -> 裁上下
        new_h = int(round(w / target_ratio))
        if anchor == "center":
            top = (h - new_h) // 2
        else:
            top = 0
        return img.crop((0, top, w, top + new_h))


def prepare_image(path, box_w_emu, box_h_emu, transparent_bg=True, white_thresh=245):
    """读图 -> 抠白底(可选) -> 按占位框宽高比居中裁切 -> 缩放。

    返回 PIL.Image（RGBA）。调用方按占位框位置/尺寸插入。
    """
    img = Image.open(path)
    if transparent_bg:
        img = _remove_bg_flood(img, white_thresh)
    else:
        img = img.convert("RGBA")
    # 裁成与占位框相同宽高比（不变形关键）
    ratio = box_w_emu / box_h_emu
    img = _crop_to_ratio(img, ratio)
    # 缩放到目标像素（按 96dpi 近似；EMU -> px: /9525）
    tw = max(1, int(round(box_w_emu / 9525)))
    th = max(1, int(round(box_h_emu / 9525)))
    # 限制最大边，避免巨图
    scale = min(1.0, 2000 / max(tw, th))
    tw, th = int(tw * scale), int(th * scale)
    img = img.resize((tw, th), Image.LANCZOS)
    buf = io.BytesIO()
    img.save(buf, format="PNG")
    buf.seek(0)
    return buf


def find_images(images_dir, model):
    """按型号找图：型号_OPEN.png/jpg（开架）、型号_SILENT.png/jpg（静音）。

    返回 {"open": path|None, "silent": path|None}
    """
    def pick(suffix):
        for ext in ("png", "jpg", "jpeg", "PNG", "JPG"):
            p = os.path.join(images_dir, "%s_%s.%s" % (model, suffix, ext))
            if os.path.exists(p):
                return p
        return None

    return {"open": pick("OPEN"), "silent": pick("SILENT")}
