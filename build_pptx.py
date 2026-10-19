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
import re
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


def _clean_value(v):
    """轻度清洗写入值：保留行首 `◆`/缩进等排版符号，只把制表符与不换行空格归一。

    注意：绝不能 strip 行首 —— features 行的 `◆ `、缩进都是模板设计要求，
    去掉会让文字挤成一团、断行错乱。只做：\xa0->空格、去行尾空白、去空行。
    """
    if v is None:
        return None
    s = str(v).replace("\u00a0", " ")
    lines = []
    for x in s.split("\n"):
        x = x.rstrip()          # 只去行尾空白，保留行首
        lines.append(x)
    # 去首尾空行，保留中间空行
    while lines and lines[0] == "":
        lines.pop(0)
    while lines and lines[-1] == "":
        lines.pop()
    return "\n".join(lines)


def set_cell(cell, text):
    """写单元格，保留原格式（清空多余 run/段落）"""
    if text is None:
        return
    text = _clean_value(text)
    set_multiline(cell, [str(text)])


def _para_prefix(p_el):
    """读出模板段落首个 run 的『符号前缀』（如 '◆ '）。

    模板 features 行首个 run 常是 '◆ '（符号单独成 run）；重建段落时必须
    原样保留，否则 ◆ 会被用户数据覆盖掉，导致文字挤成一团、断行错乱。
    判定：首个 run 文本去掉符号/空白后为空 → 视为纯前缀，返回 (前缀, 是否独立)。
    """
    ts = p_el.findall(qn("a:r"))
    if not ts:
        return None
    t = ts[0].find(qn("a:t"))
    txt = t.text if (t is not None and t.text) else ""
    core = txt.replace("◆", "").replace("●", "").replace("•", "").strip()
    if core == "":           # 纯符号/空白前缀
        return txt
    return None


def set_multiline(cell, lines, keep_prefix=False, keep_shape=False):
    """多行写入：清空后按模板首个段落样式重建所有行。

    cell 可为表格单元格，也可为 shape.text_frame（keep_shape=True 时）。
    必须先在副本上改文本、再整体替换原始 <a:p>，否则按 _p 查找会重复命中。
    空行也必须保留为独立段落（否则会多出/丢掉行，如油耗 4 行夹着空值）。
    keep_prefix=True 时，为每行自动补上模板的符号前缀（如 '◆ '），
    用于 features 这类「每行一个 ◆」的单元格。
    """
    if not lines:
        return
    lines = ["" if x is None else str(x) for x in lines]
    tf = cell if keep_shape else cell.text_frame
    base_p = tf.paragraphs[0]
    proto = copy.deepcopy(base_p._p)
    txBody = tf._txBody

    # 探测模板行前缀（如 '◆ '）：优先取"首个 run 为纯符号"的段落
    prefix = None
    if keep_prefix:
        for para in tf.paragraphs:
            pf = _para_prefix(para._p)
            if pf:
                prefix = pf
                break
        if prefix is None:
            # 回退：直接取第一段首个 run 里的符号字符
            pts = proto.findall(qn("a:r"))
            if pts:
                tn = pts[0].find(qn("a:t"))
                if tn is not None and tn.text:
                    m = re.match(r"^[◆●•►▶·]\s*", tn.text)
                    if m:
                        prefix = m.group(0)

    def set_para_text(p_el, txt):
        ts = p_el.findall(qn("a:r"))
        if ts:
            first = ts[0]
            tnode = first.find(qn("a:t"))
            if tnode is None:
                tnode = first.makeelement(qn("a:t"), {})
                first.append(tnode)
            tnode.text = txt
            for r in ts[1:]:
                p_el.remove(r)
        else:
            pass

    new_ps = []
    for ln in lines:
        pel = copy.deepcopy(proto)
        set_para_text(pel, ln)
        new_ps.append(pel)

    # 为每行补前缀：把前缀作为独立 run 插到段首（保持与模板同构）
    if prefix:
        for pel in new_ps:
            ts = pel.findall(qn("a:r"))
            if not ts:
                continue
            first = ts[0]
            tnode = first.find(qn("a:t"))
            cur = tnode.text if (tnode is not None and tnode.text is not None) else ""
            if cur.strip() == "":
                # 空行不补前缀
                continue
            # 若内容本身已带符号则不重复
            if re.match(r"^[◆●•►▶·]", cur.lstrip()):
                continue
            # 复制首个 run 的样式作为前缀 run
            pref_run = copy.deepcopy(first)
            pt = pref_run.find(qn("a:t"))
            if pt is None:
                pt = pref_run.makeelement(qn("a:t"), {})
                pref_run.append(pt)
            pt.text = prefix
            first.addprevious(pref_run)

    for old in txBody.findall(qn("a:p")):
        txBody.remove(old)
    for pel in new_ps:
        txBody.append(pel)


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


def _clear_non_data_pages(slide):
    """2026-10-19 用户规则：模板未标注「12345」的页（第3、4页）不得带任何模板示例文字，
    否则中英混排+换行与原文不一致，导出 PDF 会看起来多出空格/断行。
    做法：该页所有**表格**清成空单元格；**文本框**若含示例数据也清空（保留纯标题占位/符号框不动）。
    需在填完数据后调用；被显式填充过的单元格会在下面重新写回。"""
    A = '{http://schemas.openxmlformats.org/drawingml/2006/main}'
    for sh in slide.shapes:
        if sh.has_table:
            tbl = sh.table
            for row in tbl.rows:
                for cell in row.cells:
                    tf = cell.text_frame
                    p0 = tf.paragraphs[0]
                    # 每段保留一个空 run：字体/行高与模板一致，但显示为空
                    if p0.runs:
                        p0.runs[0].text = ""
                    for r in p0._p.findall(A + 'r')[1:]:
                        p0._p.remove(r)
                    for extra in list(tf.paragraphs)[1:]:
                        tf._txBody.remove(extra._p)



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

        # 第 3/4 页：模板示例文字先清空（用户 2026-10-19：模板里未标 12345 的内容不得带出）
        if s_idx >= 2:
            _clear_non_data_pages(slide)

        # 标题：仅第 1 页替换（第 2/3/4 页保持模板原样）
        title_tpl = tmap.get("title") if s_idx == 0 else None
        sub_title = data.get("__title_sub__") if s_idx == 0 else None
        title_custom = ((data.get("__title_custom__") or "").strip()) if s_idx == 0 else ""
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
        # 独立文本框字段（如 slide1 右下角频率/燃料小格、声压级符号框）
        for sdef in slide_def.get("shapes", []):
            si = sdef.get("shape")
            if si is None or si >= len(slide.shapes):
                print("  ! slide%d shape idx=%s 不存在，跳过" % (s_idx + 1, si), file=sys.stderr)
                continue
            sh = slide.shapes[si]
            if not sh.has_text_frame:
                continue
            fname = sdef.get("field")
            val = data.get(fname)
            if val in (None, ""):
                continue
            if sdef.get("mode") == "list":
                lines = val if isinstance(val, (list, tuple)) else [x for x in str(val).split("\n")]
                set_multiline(sh.text_frame, lines, keep_shape=True)
            else:
                set_shape_text(sh, translate(val, lang))

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
                    set_multiline(table.cell(r, c), [translate(x, lang) for x in lines],
                                  keep_prefix=True)

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
