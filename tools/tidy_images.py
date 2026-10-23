#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
图库整理：DEDUP/*.png -> images/<型号>_OPEN.png / <型号>_SILENT.png
- 同型号同名多张 -> 选主图（分辨率最高 + 文件最大），其余移到 _alt/
- 统一尺寸：最大边 1200px，保持比例
- 转 RGB 去透明，PNG 优化 / JPG 质量 85
- wmf/emf 用 Pillow 转 png
- 输出 inventory.json / report.txt
"""
import json
import os
import re
import shutil
import sys

from PIL import Image

SRC = "/tmp/dedup_check/DEDUP"
OUT = "/Users/bob/.openclaw/workspace/projects/参数单生成器/images"
ALT = "/Users/bob/.openclaw/workspace/projects/参数单生成器/images_alt"
CANVAS_W, CANVAS_H = 1600, 1200   # 统一画布（4:3），体积无所谓，清晰优先
MARGIN = 0.04                      # 四周留白比例
JPG_Q = 85

NAME_RE = re.compile(r"^(OPEN|SILENT)_(.+?)_(OPEN|SILENT)(?:_(\d+))?\.(png|jpg|jpeg|wmf|emf)$", re.I)


def parse(fn):
    m = NAME_RE.match(fn)
    if not m:
        return None
    kind = m.group(1).upper()
    model = m.group(2)
    ext = m.group(5).lower()
    idx = int(m.group(4) or 0)
    return kind, model, idx, ext


def load_img(path):
    im = Image.open(path)
    if im.mode in ("RGBA", "LA", "P"):
        im = im.convert("RGBA")
        bg = Image.new("RGB", im.size, (255, 255, 255))
        bg.paste(im, mask=im.split()[-1])
        im = bg
    else:
        im = im.convert("RGB")
    return im


def fit_canvas(im):
    """等比缩放进 CANVAS_W x CANVAS_H，白底居中，不裁切不拉伸"""
    avail_w = int(CANVAS_W * (1 - MARGIN * 2))
    avail_h = int(CANVAS_H * (1 - MARGIN * 2))
    w, h = im.size
    ratio = min(avail_w / float(w), avail_h / float(h))
    nw, nh = max(1, int(w * ratio)), max(1, int(h * ratio))
    im2 = im.resize((nw, nh), Image.LANCZOS)
    canvas = Image.new("RGB", (CANVAS_W, CANVAS_H), (255, 255, 255))
    canvas.paste(im2, ((CANVAS_W - nw) // 2, (CANVAS_H - nh) // 2))
    return canvas


def score(path, im):
    """选主图：分辨率优先，其次文件大小"""
    w, h = im.size
    return (w * h, os.path.getsize(path))


def main():
    if not os.path.isdir(SRC):
        print("源目录不存在:", SRC, file=sys.stderr)
        return 2
    os.makedirs(OUT, exist_ok=True)
    os.makedirs(ALT, exist_ok=True)

    groups = {}   # (kind, model) -> [files]
    skipped = []
    for fn in sorted(os.listdir(SRC)):
        p = os.path.join(SRC, fn)
        if not os.path.isfile(p):
            continue
        info = parse(fn)
        if not info:
            skipped.append(fn)
            continue
        kind, model, idx, ext = info
        groups.setdefault((kind, model), []).append(p)

    inventory, rows = {}, []
    n_main = n_alt = n_err = 0
    total_in = total_out = 0

    for (kind, model), files in sorted(groups.items()):
        loaded = []
        for p in files:
            try:
                im = load_img(p)
                loaded.append((p, im, score(p, im)))
            except Exception as e:
                n_err += 1
                rows.append({"model": model, "kind": kind, "file": os.path.basename(p),
                             "status": "error", "error": str(e)})
        if not loaded:
            continue
        loaded.sort(key=lambda x: -x[2][0] * 1e9 - x[2][1])
        loaded.sort(key=lambda x: (-x[2][0], -x[2][1]))
        main_p, main_im, _ = loaded[0]

        # 主图：等比缩放进统一画布，白底居中（不裁切，不拉伸）
        main_im = fit_canvas(main_im)
        # 统一存 JPG（白底），比 PNG 小 3~5 倍，PPT/网页都够用
        dst_name = "%s_%s.jpg" % (model, kind)
        dst = os.path.join(OUT, dst_name)
        main_im.save(dst, "JPEG", quality=JPG_Q, optimize=True, progressive=True)
        total_in += os.path.getsize(main_p)
        total_out += os.path.getsize(dst)
        n_main += 1
        inventory.setdefault(model, {})[kind] = {
            "file": dst_name, "w": main_im.size[0], "h": main_im.size[1],
            "kb": os.path.getsize(dst) // 1024, "alts": len(loaded) - 1,
        }
        rows.append({"model": model, "kind": kind, "file": dst_name, "status": "main",
                     "size": "%dx%d" % main_im.size, "alts": len(loaded) - 1})

        # 其余进 _alt（只缩放不覆盖）
        for p, im, sc in loaded[1:]:
            im = fit_canvas(im)
            alt_name = os.path.splitext(os.path.basename(p))[0] + ".jpg"
            im.save(os.path.join(ALT, alt_name), "JPEG", quality=JPG_Q, optimize=True)
            n_alt += 1

    with open(os.path.join(OUT, "..", "images_inventory.json"), "w", encoding="utf-8") as f:
        json.dump(inventory, f, ensure_ascii=False, indent=1)

    both = sum(1 for v in inventory.values() if "OPEN" in v and "SILENT" in v)
    only_o = sum(1 for v in inventory.values() if "OPEN" in v and "SILENT" not in v)
    only_s = sum(1 for v in inventory.values() if "SILENT" in v and "OPEN" not in v)

    with open("/Users/bob/.openclaw/workspace/projects/参数单生成器/images_report.txt", "w", encoding="utf-8") as f:
        f.write("图库整理报告\n" + "=" * 50 + "\n")
        f.write("型号总数      : %d\n" % len(inventory))
        f.write("  开架+静音都有: %d\n" % both)
        f.write("  只有开架     : %d\n" % only_o)
        f.write("  只有静音     : %d\n" % only_s)
        f.write("标准图数量    : %d\n" % n_main)
        f.write("备用图数量    : %d（images_alt）\n" % n_alt)
        f.write("转换失败      : %d\n" % n_err)
        f.write("原始总大小    : %.1f MB\n" % (total_in / 1048576.0))
        f.write("整理后总大小  : %.1f MB\n" % (total_out / 1048576.0))
        if skipped:
            f.write("\n无法识别文件名 %d 个:\n" % len(skipped))
            for s in skipped[:30]:
                f.write("  %s\n" % s)

    print("型号 %d（开架+静音 %d / 只开架 %d / 只静音 %d）"
          % (len(inventory), both, only_o, only_s))
    print("标准图 %d 张，备用图 %d 张，失败 %d" % (n_main, n_alt, n_err))
    print("体积 %.1f MB -> %.1f MB" % (total_in / 1048576.0, total_out / 1048576.0))
    print("输出:", OUT)
    return 0


if __name__ == "__main__":
    sys.exit(main())
