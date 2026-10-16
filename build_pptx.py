#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
参数单生成器 V2 — 后端最小闭环
读数据 JSON -> 套模板 (template_map.<lang>.json) -> 出 pptx

用法:
  python3 build_pptx.py --data data/D17.json --lang en --out output/
"""
import argparse
import copy
import io
import json
import os
import sys

from pptx import Presentation
from pptx.oxml.ns import qn

HERE = os.path.dirname(os.path.abspath(__file__))


def load_map(lang):
    p = os.path.join(HERE, "templates", "template_map.%s.json" % lang)
    with io.open(p, encoding="utf-8") as f:
        return json.load(f)


_GLOSSARY = None


def load_glossary():
    global _GLOSSARY
    if _GLOSSARY is None:
        p = os.path.join(HERE, "glossary.json")
        if os.path.exists(p):
            with io.open(p, encoding="utf-8") as f:
                raw = json.load(f)
            # rows 格式: [中文, 英文, 适用]，忽略以 __ 开头的分类标题行
            zh2en, en2zh, prefix = {}, {}, {}
            for row in raw.get("rows", []):
                if len(row) < 2:
                    continue
                zh, en = row[0], row[1]
                if not zh or zh.startswith("__"):
                    continue
                zh2en[zh] = en
                if en and en not in en2zh:
                    en2zh[en] = zh
                if zh.endswith(" "):
                    prefix[zh] = en
            _GLOSSARY = {"zh2en": zh2en, "en2zh": en2zh, "prefix": prefix}
        else:
            _GLOSSARY = {"zh2en": {}, "en2zh": {}, "prefix": {}}
    return _GLOSSARY


def translate(val, lang):
    """按语言翻译单元格值。

    - zh: 若值是英文术语（命中 en2zh）则转回中文，否则原样
    - en: 若值是中文术语（命中 zh2en）则转英文，否则原样
    - 整值精确命中优先；否则做前缀替换（如 '等级 H/H' -> 'Class H/H'）
    - 数值/含单位可变值不命中则不动
    """
    if val is None:
        return val
    g = load_glossary()
    s = str(val)
    if lang == "en":
        if s in g["zh2en"]:
            return g["zh2en"][s]
        for k, v in g["prefix"].items():
            if s.startswith(k):
                return v + s[len(k):]
    elif lang == "zh":
        if s in g["en2zh"]:
            return g["en2zh"][s]
    return val


def set_cell(cell, text):
    """写单元格，保留原格式（清空多余 run/段落）"""
    if text is None:
        return
    set_multiline(cell, [str(text)])


def set_multiline(cell, lines, autofit=True):
    """多行写入一个单元格：清空后按模板首个段落样式重建所有行。

    必须先在副本上改文本、再整体替换原始 <a:p>，否则按 _p 查找会重复命中。
    空行也必须保留为独立段落（否则会多出/丢掉行，如油耗 4 行夹着空值）。
    autofit=True 时，行数/单行字数偏多会按比例缩小字号，防止文字溢出。
    """
    if not lines:
        return
    lines = ["" if x is None else str(x) for x in lines]
    if autofit:
        try:
            _autofit_cell_font(cell, lines)
        except Exception:
            pass
    tf = cell.text_frame
    base_p = tf.paragraphs[0]
    proto = copy.deepcopy(base_p._p)
    txBody = tf._txBody

    def set_para_text(p_el, txt):
        ts = p_el.findall(qn("a:r"))
        if ts:
            first = ts[0]
            # 第一个 run 保留，写文本
            tnode = first.find(qn("a:t"))
            if tnode is None:
                tnode = first.makeelement(qn("a:t"), {})
                first.append(tnode)
            tnode.text = txt
            # 其余 run 删除
            for r in ts[1:]:
                p_el.remove(r)
        else:
            # 没有 run，就不要写文本（保持空段落）
            pass

    new_ps = []
    for ln in lines:
        pel = copy.deepcopy(proto)
        set_para_text(pel, ln)
        new_ps.append(pel)

    for old in txBody.findall(qn("a:p")):
        txBody.remove(old)
    for pel in new_ps:
        txBody.append(pel)


def _autofit_cell_font(cell, lines):
    """估算文本是否溢出单元格，若偏多则按比例缩小字号（下限 6pt）。"""
    from pptx.util import Pt
    tf = cell.text_frame
    cur = None
    for para in tf.paragraphs:
        for run in para.runs:
            if run.font.size:
                cur = run.font.size
                break
        if cur:
            break
    if not cur:
        return
    # 估算：中文 1.0 宽 / 英文 0.55 宽；单行容量约 46 个"标准字符"
    CAP = 46.0
    total_rows = 0
    max_w = 0.0
    for ln in lines:
        w = sum(1.0 if ord(ch) > 127 else 0.55 for ch in ln)
        max_w = max(max_w, w)
        total_rows += max(1, int(w / CAP) + (1 if w % CAP else 0))
    target = cur.pt
    if total_rows > 11 or max_w > CAP:
        ratio = min(1.0, 11.0 / total_rows, CAP / max(max_w, 1.0))
        target = max(6.0, round(cur.pt * ratio, 1))
    if target < cur.pt:
        for para in tf.paragraphs:
            for run in para.runs:
                run.font.size = Pt(target)


def set_shape_text(shape, text, lang=None):
    """写普通文本框（保留首个 run 样式）"""
    if not shape.has_text_frame or text is None:
        return
    tf = shape.text_frame
    p0 = tf.paragraphs[0]
    if p0.runs:
        p0.runs[0].text = text
        for r in p0.runs[1:]:
            r.text = ""
    else:
        p0.text = text
    for extra in tf.paragraphs[1:]:
        for r in extra.runs:
            r.text = ""


def get_path(obj, path):
    cur = obj
    for k in path:
        if isinstance(cur, dict):
            cur = cur.get(k)
        else:
            return None
    return cur


def replace_picture(slide, old_pic_shape, img_stream):
    """把已有图片形状换为新图，保持原位置/尺寸。"""
    try:
        blip = old_pic_shape._element.blipFill.blip
        part = slide.part
        image_part, new_rId = part.get_or_add_image_part(img_stream)
        blip.embed = new_rId
        return True
    except Exception as e:
        print("  ! 图片替换失败:", e, file=sys.stderr)
        return False


def fill_pictures(slide, img_map, pic_slots=None):
    """把机组照片填进 slide 的大图位。

    img_map: {"open": 流|None, "silent": 流|None}
    pic_slots: [左框shape, 右框shape]（不传则自动取面积最大的两个图位）
    规则：两张都有→左开架右静音；只有静音→居中（用两框合并居中位置）。
    """
    pics = [sh for sh in slide.shapes if sh.shape_type == 13]
    if not pics:
        return
    if pic_slots:
        left_pic, right_pic = pic_slots
    else:
        big = sorted(pics, key=lambda s: s.width * s.height, reverse=True)[:2]
        if len(big) < 2:
            return
        # 按 left 排序定左右
        big = sorted(big, key=lambda s: s.left)
        left_pic, right_pic = big[0], big[1]

    o, s = img_map.get("open"), img_map.get("silent")
    if o and s:
        replace_picture(slide, left_pic, o)
        replace_picture(slide, right_pic, s)
    elif s and not o:
        # 只有静音图 -> 居中（跨两框合并位置 + 居中）
        l = min(left_pic.left, right_pic.left)
        r = max(left_pic.left + left_pic.width, right_pic.left + right_pic.width)
        t = min(left_pic.top, right_pic.top)
        h = max(left_pic.height, right_pic.height)
        w = r - l
        # 居中放置：宽度用两框合并、高度不变、水平居中
        replace_picture(slide, left_pic, s)
        left_pic.left = l + (w - left_pic.width) // 2
        # 隐藏右框（设为极简空图占位？直接删除形状）
        right_pic._element.getparent().remove(right_pic._element)
    elif o and not s:
        replace_picture(slide, left_pic, o)
        right_pic._element.getparent().remove(right_pic._element)


def set_title_on_slide(slide, text, sub_text=None):
    """替换 slide1 顶部主标题 + 小标题（型号字）。

    模板结构（实测）：
      - 小标题（型号字，如 AS500E5/WP715D5）：top ≈ 0.12in，文本不含标题关键词
      - 主标题（大标题）：top ≈ 0.63~0.72in，文本含 'Industrial Range'/'Powered By'/'采用'
    定位规则：先按文本关键词找主标题；小标题取“顶部且非主标题”的那个框。
    """
    MAIN_KEYS = ("Industrial Range", "Powered By", "采用")

    def _is_main(t):
        return any(k in t for k in MAIN_KEYS)

    # 收集顶部（top < 2.0in）的文本框，按 top 升序
    tops = []
    for sh in slide.shapes:
        if not sh.has_text_frame or not sh.text_frame.text.strip():
            continue
        if sh.top is not None and sh.top < 2000000:
            tops.append(sh)
    tops.sort(key=lambda sh: sh.top or 0)

    main_sh = None
    sub_sh = None
    for sh in tops:
        t = sh.text_frame.text.strip()
        if _is_main(t):
            main_sh = sh          # 含关键词 -> 主标题
        elif sub_sh is None:
            sub_sh = sh           # 其余最上面的 -> 小标题（型号字）

    # 兜底：模板关键词变了，就按位置取（最上面=小标题，次上=主标题）
    if main_sh is None and len(tops) >= 2:
        sub_sh, main_sh = tops[0], tops[1]

    if text and main_sh is not None:
        set_shape_text(main_sh, text)
    if sub_text and sub_sh is not None:
        set_shape_text(sub_sh, sub_text)


def build(data, lang="en", template=None, outdir="output", images_dir=None, transparent_bg=True, pic_slots=None):
    tmap = load_map(lang)
    tpl = template or os.path.join(HERE, tmap["template"])
    prs = Presentation(tpl)
    images_dir = images_dir or os.path.join(HERE, "images")

    for slide_def in tmap["slides"]:
        s_idx = slide_def["slide"] - 1
        slide = prs.slides[s_idx]
        # 机组照片（仅 slide1）
        if s_idx == 0 and images_dir:
            import image_utils
            found = image_utils.find_images(images_dir, data.get("model", ""))
            # 手动指定图优先（data.__img_open__ / __img_silent__ 为绝对路径）
            at = data.get("__img_open__")
            st = data.get("__img_silent__")
            if at:
                found["open"] = at
            if st:
                found["silent"] = st
            if found.get("open") or found.get("silent"):
                img_map = {}
                # 先用占位框尺寸算裁切比（左框）
                pics = sorted([sh for sh in slide.shapes if sh.shape_type == 13],
                              key=lambda s: s.width * s.height, reverse=True)[:2]
                if len(pics) >= 2:
                    bw, bh = pics[0].width, pics[0].height
                else:
                    bw, bh = 3000000, 2000000
                for key in ("open", "silent"):
                    p = found.get(key)
                    if p and os.path.exists(p):
                        img_map[key] = image_utils.prepare_image(
                            p, bw, bh, transparent_bg=transparent_bg)
                if img_map:
                    fill_pictures(slide, img_map, pic_slots)
                    print("  已插入图片:", list(img_map.keys()))

        # 标题：模板串 + 品牌变量（取该语言的品牌名）
        title_tpl = tmap.get("title")
        sub_title = data.get("__title_sub__")
        title_custom = (data.get("__title_custom__") or "").strip()
        if title_tpl or sub_title or title_custom:
            brand_raw = data.get("e_brand") or ""
            brand = translate(brand_raw, lang) if brand_raw else ""
            if title_custom:
                # 用户手填大标题优先（支持 {engine_brand} 占位）
                title_text = title_custom.replace("{engine_brand}", str(brand))
            elif title_tpl:
                title_text = title_tpl.replace("{engine_brand}", str(brand))
            else:
                title_text = None
            set_title_on_slide(slide, title_text, sub_title)
        shape_tables = [sh for sh in slide.shapes if sh.has_table]
        for t_def in slide_def.get("tables", []):
            if t_def["idx"] >= len(shape_tables):
                print("  ! slide%d table idx=%d 不存在，跳过" % (s_idx + 1, t_def["idx"]), file=sys.stderr)
                continue
            table = shape_tables[t_def["idx"]].table
            # 普通字段
            for fname, coord in t_def.get("fields", {}).items():
                val = get_path(data, fname.split(".")) if "." in fname else data.get(fname)
                if val in (None, ""):
                    continue
                r, c = coord
                set_cell(table.cell(r, c), translate(val, lang))
            # 多行字段（油耗）
            if "fuel_consumption" in t_def:
                fc = t_def["fuel_consumption"]
                lines = data.get(fc["data_key"])
                if isinstance(lines, str):
                    lines = [x for x in lines.split("\n") if x.strip()]
                if lines and any(x not in (None, "") for x in lines):
                    r, c = fc["cell"]
                    set_multiline(table.cell(r, c), lines)
            # 多行文本字段（features）
            for fname, coord in t_def.get("multiline_fields", {}).items():
                raw = data.get(fname)
                if not raw:
                    continue
                lines = [x for x in str(raw).split("\n") if x.strip()]
                if lines:
                    r, c = coord
                    set_multiline(table.cell(r, c), [translate(x, lang) for x in lines])

    os.makedirs(os.path.join(HERE, outdir) if not os.path.isabs(outdir) else outdir, exist_ok=True)
    model = data.get("model", "model")
    outpath = os.path.join(os.path.join(HERE, outdir) if not os.path.isabs(outdir) else outdir,
                           "%s_paramsheet_%s.pptx" % (model, lang))
    prs.save(outpath)
    return outpath


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--data", required=True)
    ap.add_argument("--lang", default="en")
    ap.add_argument("--template")
    ap.add_argument("--out", default="output")
    a = ap.parse_args()
    with io.open(a.data, encoding="utf-8") as f:
        data = json.load(f)
    out = build(data, a.lang, a.template, a.out)
    print(out)


if __name__ == "__main__":
    main()
