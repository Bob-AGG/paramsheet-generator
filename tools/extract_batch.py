#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
批量提取历史参数单（公司电脑上跑，不联网、不上传、不调用 AI）

用途：
    把 NAS 上成百上千份历史参数单（PDF / PPT / PPTX）批量读出文本与表格，
    每份产出一个小的 JSON（~5KB），最后由你上传这批 JSON 给下游做结构化。
    原始文件始终留在公司，数据不出内网。

用法：
    python extract_batch.py --src "Z:\\参数单\\历史" --out "D:\\paramsheet_extract"

    # 只跑前 30 份做试跑
    python extract_batch.py --src "Z:\\参数单\\历史" --out "D:\\try30" --limit 30

    # 递归子目录（默认就递归，可用 --no-recursive 关掉）
    python extract_batch.py --src "Z:\\参数单" --out "D:\\out" --no-recursive

依赖（装一次）：
    pip install pdfplumber pypdf python-pptx

输出结构：
    <out>/
        index.json              ← 汇总清单（型号、来源、是否成功、字数）
        items/<key>.json        ← 每份参数单一份
        failed/                 ← 空文件/读不出的原文件路径清单
        report.txt              ← 人看的文字报告
"""

import argparse
import hashlib
import threading
import io
import json
import os
import re
import sys
import time
import traceback
from concurrent.futures import ThreadPoolExecutor, as_completed

EXTS = (".pdf", ".ppt", ".pptx")
PPTX_ONLY_EXTS = (".ppt", ".pptx")   # --pptx-only：跳过 PDF（PDF 表格会被压成文字流，只作兜底）


# ============================================================
# 文本抽取
# ============================================================

def extract_pdf(path):
    """PDF -> (text, tables, meta)。优先 pdfplumber，退 pypdf。"""
    text, tables, meta = "", [], {}
    try:
        import pdfplumber
        with pdfplumber.open(path) as pdf:
            meta["pages"] = len(pdf.pages)
            chunks = []
            for i, pg in enumerate(pdf.pages):
                chunks.append(pg.extract_text() or "")
                if i < 8:  # 表格只取前 8 页，够用且快
                    try:
                        for tb in (pg.extract_tables() or []):
                            if tb and len(tb) > 1:
                                tables.append({"page": i + 1,
                                               "rows": [[(c or "").strip() for c in row]
                                                        for row in tb]})
                    except Exception:
                        pass
            text = "\n".join(chunks)
        if text.strip():
            return text, tables, meta
    except Exception as e:
        meta["pdfplumber_error"] = str(e)

    try:
        from pypdf import PdfReader
        r = PdfReader(path)
        meta["pages"] = len(r.pages)
        text = "\n".join((p.extract_text() or "") for p in r.pages[:30])
    except Exception as e:
        meta["pypdf_error"] = str(e)
    return text, tables, meta


def extract_pptx(path):
    """PPTX -> (text, tables, meta)。python-pptx 读文本框与表格。"""
    text, tables, meta = "", [], {}
    try:
        from pptx import Presentation
        prs = Presentation(path)
        meta["slides"] = len(prs.slides)
        chunks = []
        for si, slide in enumerate(prs.slides):
            slide_lines = []
            for sh in slide.shapes:
                # 表格
                if getattr(sh, "has_table", False) and sh.has_table:
                    rows = []
                    for row in sh.table.rows:
                        rows.append([(c.text or "").strip() for c in row.cells])
                    if len(rows) > 1:
                        tables.append({"slide": si + 1, "rows": rows})
                    for row in rows:
                        slide_lines.append(" | ".join(x for x in row if x))
                    continue
                # 文本框
                if getattr(sh, "has_text_frame", False) and sh.has_text_frame:
                    t = sh.text_frame.text
                    if t and t.strip():
                        slide_lines.append(t.strip())
            if slide_lines:
                chunks.append("\n".join(slide_lines))
        text = "\n\n".join(chunks)
    except Exception as e:
        meta["pptx_error"] = str(e)
    return text, tables, meta


def extract_ppt_legacy(path):
    """老 .ppt 二进制格式：python-pptx 读不了，尝试用本机 LibreOffice 转换。

    若公司电脑装了 LibreOffice，会自动转成 pptx 再读；没装就记为需手工处理。
    """
    meta = {}
    cands = [
        r"C:\Program Files\LibreOffice\program\soffice.exe",
        r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
        "soffice",
    ]
    soffice = None
    for c in cands:
        if c == "soffice" or os.path.exists(c):
            soffice = c
            break
    if not soffice:
        meta["error"] = "老 .ppt 格式，且本机未装 LibreOffice，无法读取"
        return "", [], meta
    import subprocess
    import tempfile
    tmpdir = tempfile.mkdtemp(prefix="pptconv_")
    try:
        subprocess.run([soffice, "--headless", "--norestore",
                        "--convert-to", "pptx", "--outdir", tmpdir, path],
                       timeout=180, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    except Exception as e:
        meta["error"] = "LibreOffice 转换失败: %s" % e
        return "", [], meta
    base = os.path.splitext(os.path.basename(path))[0] + ".pptx"
    conv = os.path.join(tmpdir, base)
    if not os.path.exists(conv):
        meta["error"] = "LibreOffice 转换后未找到 pptx"
        return "", [], meta
    meta["converted_from"] = "ppt"
    t, tb, m2 = extract_pptx(conv)
    meta.update(m2)
    try:
        os.remove(conv)
        os.rmdir(tmpdir)
    except Exception:
        pass
    return t, tb, meta


def extract_any(path):
    low = path.lower()
    if low.endswith(".pdf"):
        return extract_pdf(path)
    if low.endswith(".pptx"):
        return extract_pptx(path)
    if low.endswith(".ppt"):
        return extract_ppt_legacy(path)
    return "", [], {"error": "不支持的类型"}


# ============================================================
# 型号猜测（仅用于给文件起名/索引，抽错不影响下游）
# ============================================================

# ⚠️ 型号判定顺序很重要：文件名往往比正文可靠（正文里会混进 "2000/14/CE" "EN60204" 之类），
# 所以先看文件名，再退回正文。旧版先扫正文，结果把 "345"(来自 12345 占位) 当型号。
# ⚠️ 不能用 \b：中文和 _ 在 Unicode 下都算 \w，"WP715D5_中文" 里 \b 不成立 -> 匹配失败。
# 改成 ASCII 前后否定。
_STRICT_MODEL = re.compile(r"(?<![A-Za-z0-9])([A-Z]{1,4}\d{2,5}[A-Z0-9\-]{0,10})(?![A-Za-z0-9])")
# 明显不是型号的：标准号、电压电流、纯数字、频率
_NOT_MODEL = re.compile(
    r"^(EN|ISO|CE|IEC|GB|DIN|IP|DC|AC|RPM|Hz|PF)\b|^\d{1,4}$|^\d+V$|^\d+A$", re.I)
# 数字开头的型号（6M33D633E200 / 4BTA3.9 / 6CT8.3），单独一条
# 数字开头的型号（12KTA58-G32 / 6M33D633E200 / 4BTA3.9 / 6CT8.3）
# 2026-10-20 修：原 pattern 要求 "2位数字+1~3字母"，把 12KTA58-G32 切成 G32。
# 改为 "1~2位数字 + 至少1位字母开头的主体"，主体内允许字母数字混合。
_STRICT_MODEL_DIGIT = re.compile(
    r"(?<![A-Za-z0-9])(\d{1,2}[A-Z][A-Z0-9\.]{1,14}(?:-[A-Z0-9\.]+)*)(?![A-Za-z0-9])")


def _clean_model(v):
    v = (v or "").strip(" .-_/")
    if not v or len(v) < 3 or len(v) > 30:
        return ""
    if _NOT_MODEL.match(v):
        return ""
    if not re.search(r"\d", v):     # 型号必须带数字
        return ""
    return v


def guess_model(text, filename):
    """先文件名、后正文（文件名更干净）。猜不到返回空，不影响下游。"""
    # 1) 文件名里的候选（去掉扩展名，去掉 12345 之类的模板占位）
    base = os.path.splitext(os.path.basename(filename))[0]
    base = re.sub(r"12345+", " ", base)
    # 数字开头的先试（12KTA58-G32），否则会被后面那条切成 G32
    for rx in (_STRICT_MODEL_DIGIT, _STRICT_MODEL):
        for m in rx.finditer(base):
            v = _clean_model(m.group(1))
            if v:
                return v

    # 2) 正文里带标签的（Model / 型号 / 机组型号）
    head = "\n".join(text.split("\n")[:40])
    m = re.search(r"(?:Model|型号|机组型号|Genset Model)[^\n:：]{0,8}[:：]?\s*([A-Za-z0-9][A-Za-z0-9\-\./]{2,29})",
                  head, re.I)
    if m:
        v = _clean_model(m.group(1))
        if v:
            return v

    # 3) 正文兜底（末位方案，容易误判，所以放最后）
    for rx in (_STRICT_MODEL_DIGIT, _STRICT_MODEL):
        for m in rx.finditer(head):
            v = _clean_model(m.group(1))
            if v:
                return v
    return ""


def safe_key(s, fallback):
    s = re.sub(r"[^A-Za-z0-9\-_\.]+", "_", (s or "").strip())
    return s[:60] or fallback


# ============================================================
# 主流程
# ============================================================

def collect_files(src, recursive=True, exts=EXTS):
    out = []
    if recursive:
        for root, _dirs, files in os.walk(src):
            for fn in files:
                if fn.lower().endswith(exts) and not fn.startswith("~$"):
                    out.append(os.path.join(root, fn))
    else:
        for fn in os.listdir(src):
            p = os.path.join(src, fn)
            if os.path.isfile(p) and fn.lower().endswith(exts) and not fn.startswith("~$"):
                out.append(p)
    out.sort()
    return out


_PRINT_LOCK = threading.Lock()
_tls = threading.local()


def main():
    ap = argparse.ArgumentParser(description="批量提取历史参数单文本（本地跑，不联网）")
    ap.add_argument("--src", required=True, help="源目录（NAS 映射盘或本地目录）")
    ap.add_argument("--out", required=True, help="输出目录")
    ap.add_argument("--limit", type=int, default=0, help="只处理前 N 份（试跑用）")
    ap.add_argument("--no-recursive", action="store_true", help="不递归子目录")
    ap.add_argument("--skip-existing", action="store_true", help="跳过已提取过的（断点续跑）")
    ap.add_argument("--pptx-only", action="store_true", help="只处理 PPT/PPTX，跳过 PDF（快一倍，推荐）")
    ap.add_argument("--jobs", type=int, default=8, help="并发线程数（默认 8，机械盘设 2~4，SSD 设 8~16）")
    a = ap.parse_args()

    if not os.path.isdir(a.src):
        print("源目录不存在: %s" % a.src)
        sys.exit(1)

    items_dir = os.path.join(a.out, "items")
    os.makedirs(items_dir, exist_ok=True)

    exts = PPTX_ONLY_EXTS if a.pptx_only else EXTS
    files = collect_files(a.src, recursive=not a.no_recursive, exts=exts)
    if a.limit > 0:
        files = files[:a.limit]
    total = len(files)
    print("发现 %d 个文件（%s），开始提取…（%d 线程）"
          % (total, "/".join(exts), max(1, a.jobs)))
    print("输出目录: %s\n" % a.out)

    index = []
    failed = []
    t_start = time.time()

    # 断点续跑：把已存在的 item key 收进集合（含 _N 后缀），命中即跳过
    existing = set()
    if a.skip_existing:
        for fn in os.listdir(items_dir):
            if fn.endswith(".json"):
                existing.add(fn[:-5])

    def handle_one(i, path):
        """处理单个文件 -> (row, item, ok, err)"""
        rel = os.path.relpath(path, a.src)
        try:
            size_mb = os.path.getsize(path) / 1048576.0
        except OSError:
            size_mb = 0.0
        try:
            text, tables, meta = extract_any(path)
        except Exception:
            text, tables, meta = "", [], {"error": traceback.format_exc(limit=2)}

        model = guess_model(text, path)
        chars = len(text.strip())
        tb_rows = sum(len(t.get("rows") or []) for t in tables)

        key = safe_key(model, "file%04d" % i)
        if a.skip_existing and (key in existing or any(
                k.startswith(key + "_") and k[len(key) + 1:].isdigit() for k in existing)):
            return i, {"key": key, "model_guess": model, "source_rel": rel,
                       "chars": chars, "tables": len(tables), "ok": True,
                       "error": "skipped"}, True, rel, chars, len(tables), "(已存在，跳过)"
        fname = key + ".json"
        with _PRINT_LOCK:
            n = 1
            while os.path.exists(os.path.join(items_dir, fname)):
                fname = "%s_%d.json" % (key, n)
                n += 1
            # 先占位，避免多线程撞名
            open(os.path.join(items_dir, fname), "a").close()

        ok = bool(chars >= 50)
        item = {
            "key": os.path.splitext(fname)[0],
            "model_guess": model,
            "source_file": os.path.basename(path),
            "source_rel": rel,
            "size_mb": round(size_mb, 2),
            "chars": chars,
            "table_count": len(tables),
            "table_rows": tb_rows,
            "ok": ok,
            "meta": meta,
            "text": text,
            "tables": tables,
        }
        with io.open(os.path.join(items_dir, fname), "w", encoding="utf-8") as f:
            json.dump(item, f, ensure_ascii=False, indent=1)

        row = {"key": item["key"], "model_guess": model, "source_rel": rel,
               "chars": chars, "tables": len(tables), "ok": ok,
               "error": (meta.get("error") or "")[:80]}
        return i, row, ok, rel, chars, len(tables), (meta.get("error") or "")

    done = 0
    results = [None] * total
    if a.jobs <= 1:
        for i, path in enumerate(files, 1):
            i, row, ok, rel, chars, ntb, err = handle_one(i, path)
            results[i - 1] = (row, ok)
            flag = "OK " if ok else "!! "
            print("[%4d/%d] %s%-45s %6d 字  %2d 表  %s"
                  % (i, total, flag, rel[:45], chars, ntb, err[:40]))
    else:
        with ThreadPoolExecutor(max_workers=a.jobs) as ex:
            futs = {ex.submit(handle_one, i, p): i for i, p in enumerate(files, 1)}
            for fut in as_completed(futs):
                i, row, ok, rel, chars, ntb, err = fut.result()
                results[i - 1] = (row, ok)
                done += 1
                flag = "OK " if ok else "!! "
                print("[%4d/%d] %s%-45s %6d 字  %2d 表  %s"
                      % (done, total, flag, rel[:45], chars, ntb, err[:40]))

    for r in results:
        if r is None:
            continue
        row, ok = r
        index.append(row)
        if not ok:
            failed.append(row)

    cost = time.time() - t_start
    with io.open(os.path.join(a.out, "index.json"), "w", encoding="utf-8") as f:
        json.dump({"src": a.src, "total": total, "ok": total - len(failed),
                   "failed": len(failed), "seconds": round(cost, 1),
                   "items": index}, f, ensure_ascii=False, indent=1)

    # 文字报告
    lines = ["批量提取报告", "=" * 50,
             "源目录: %s" % a.src,
             "文件总数: %d" % total,
             "成功: %d    失败/可疑: %d" % (total - len(failed), len(failed)),
             "耗时: %.1f 秒 (平均 %.2f 秒/份)" % (cost, cost / max(total, 1)),
             ""]
    if failed:
        lines.append("需要人工看的文件 (%d):" % len(failed))
        for r in failed:
            lines.append("  - %s    [%s]  %s" % (r["source_rel"], r["model_guess"] or "无型号", r["error"]))
        lines.append("")
    by_model = {}
    for r in index:
        by_model.setdefault(r["model_guess"] or "(未识别)", []).append(r["source_rel"])
    dup = {k: v for k, v in by_model.items() if len(v) > 1 and k != "(未识别)"}
    if dup:
        lines.append("⚠ 同型号多份文件 (%d 个型号)，注意是否重复录入:" % len(dup))
        for k, v in list(dup.items())[:40]:
            lines.append("  %s -> %d 份: %s" % (k, len(v), "; ".join(v[:3])))
        lines.append("")
    with io.open(os.path.join(a.out, "report.txt"), "w", encoding="utf-8") as f:
        f.write("\n".join(lines))

    print("\n" + "\n".join(lines[:8]))
    print("\n输出: %s" % a.out)
    print("下一步: 把整个输出目录打包压缩上传（体积约 %s）"
          % _dir_size_hint(a.out))


def _dir_size_hint(d):
    tot = 0
    for root, _dr, fs in os.walk(d):
        for fn in fs:
            try:
                tot += os.path.getsize(os.path.join(root, fn))
            except Exception:
                pass
    if tot < 1048576:
        return "%.0f KB" % (tot / 1024.0)
    return "%.1f MB" % (tot / 1048576.0)


if __name__ == "__main__":
    main()
