#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
convert_extract.py 的三重校验（用户要求：检验三次）

T1  单元测试：标签归一化 / 空值 / 单位拆分 / 别名 / 值映射（纯函数层）
T2  数据校验：字段合法性 / 数值类型 / 功率一致性 kW≈kVA×PF / 频率块非空 / 命名唯一
T3  端到端往返：写 SQLite -> 读回 -> 导出 -> 比对字段完全一致

用法：
    python tools/test_convert.py --data /tmp/models_30.json
    python tools/test_convert.py --data /tmp/models_30.json --extract /tmp/cu30/试跑30/items
"""
import argparse
import json
import os
import re
import sys
import tempfile

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
sys.path.insert(0, HERE)
sys.path.insert(0, ROOT)

import convert_extract as CE

PASS = []
FAIL = []


def ok(t, msg):
    PASS.append("%s: %s" % (t, msg))


def bad(t, msg):
    FAIL.append("%s: %s" % (t, msg))


# ============================================================
# T1 单元测试
# ============================================================
def t1_unit():
    T = "T1-单元"

    cases = [
        # (输入, 期望)
        ("Engine  Model", "enginemodel"),
        ("Bore x stroke", "borexstroke"),
        ("No.\xa0of\xa0bearings", "noofbearings"),
        ("Frequency / Phase / Power factor", "frequencyphasepowerfactor"),
        ("Total coolant capacity\n(only engine)", "totalcoolantcapacityonlyengine"),
    ]
    for src, exp in cases:
        got = CE.norm_label(src)
        (ok if got == exp else bad)(T, "norm_label(%r)=%r 期望 %r" % (src, got, exp))

    # clean() 空值识别
    for v in ("", "  ", "TBA", "N/A", "n/a", "-", "待定", "12345", None):
        got = CE.clean(v)
        (ok if got is None else bad)(T, "clean(%r) 应为 None，实际 %r" % (v, got))
    for v, exp in ((" 37.8 L ", "37.8 L"), ("480℃", "480℃"), ("8L", "8L")):
        got = CE.clean(v)
        (ok if got == exp else bad)(T, "clean(%r)=%r 期望 %r" % (v, got, exp))

    # split_unit()
    su = [("37.8 L", ("37.8", "L")), ("1500rpm", ("1500", "rpm")),
          ("650°C", ("650", "°C")), ("90 L", ("90", "L")),
          ("ABC", ("ABC", ""))]
    for src, exp in su:
        got = CE.split_unit(src)
        (ok if got == exp else bad)(T, "split_unit(%r)=%r 期望 %r" % (src, got, exp))

    # keep_with_unit()
    kw = [("111.1 m³/min", "111.1m³/min"), ("6.7 kPa", "6.7kPa"),
          ("10.0kPa", "10.0kPa")]
    for src, exp in kw:
        got = CE.keep_with_unit(src)
        (ok if got == exp else bad)(T, "keep_with_unit(%r)=%r 期望 %r" % (src, got, exp))

    # 标签别名精确匹配（不能误命中）
    if CE.LABELS.get("governor") == "e_gov":
        ok(T, "governor -> e_gov")
    else:
        bad(T, "governor 别名缺失")
    # "Engine model" 不能被 "enginebrand" 抢
    if CE.LABELS.get("enginemodel") == "e_model" and CE.LABELS.get("enginebrand") == "e_brand":
        ok(T, "enginebrand/enginemodel 不串位")
    else:
        bad(T, "enginebrand/enginemodel 串位")

    # 值映射
    vm = [("g_exc", "Self-excited", "Self-excited"),
          ("g_conn", "Star", "Star"),
          ("g_conn", "Serie", "Star"),
          ("g_phase", "Three phase", "Three phase"),
          ("g_brg", "Single bearing", "Single bearing")]
    for key, raw, exp in vm:
        table = CE.VALUE_MAP.get(key, {})
        got = table.get(CE.norm_label(raw))
        (ok if got == exp else bad)(T, "VALUE_MAP[%s][%r]=%r 期望 %r" % (key, raw, got, exp))

    # 频率识别
    gf = [("AS\\50HZ-1P\\x.pptx", None, "50Hz"),
          ("AS\\60hz\\x.pptx", None, "60Hz"),
          ("", "50Hz / 1P / PF 1.0", "50Hz"),
          ("", "60Hz / 3P / PF0.8", "60Hz"),
          ("weird\\path\\x.pptx", None, None)]
    for src, df, exp in gf:
        got = CE.guess_freq(src, df)
        (ok if got == exp else bad)(T, "guess_freq(%r,%r)=%r 期望 %r" % (src, df, got, exp))

    # 系列识别
    fam = [("AF11D5-1P", "", "AF"), ("AS1250E5A", "", "AS"),
           ("GC1000D5-KT", "", "GC"),
           ("XYZ", r"C:\01-CUMMINS\GC\a.pptx", "GC"),      # 路径里含 GC 目录
           ("XYZ", r"C:\01-CUMMINS\weird\a.pptx", "其他")]  # 路径和型号都无线索
    for m, s, exp in fam:
        got = CE.guess_family(m, s)
        (ok if got == exp else bad)(T, "guess_family(%r)=%r 期望 %r" % (m, got, exp))


# ============================================================
# T2 数据校验
# ============================================================
def t2_data(models):
    T = "T2-数据"
    if not models:
        bad(T, "机型列表为空")
        return
    names = set()
    for m in models:
        nm = m.get("name") or ""
        d = m.get("data") or {}
        if not nm:
            bad(T, "存在空 name")
        if nm in names:
            bad(T, "名称重复: %s" % nm)
        names.add(nm)
        if not d.get("model"):
            bad(T, "%s 缺 model" % nm)
        freqs = d.get("freqs") or {}
        if not freqs:
            bad(T, "%s 无 freqs" % nm)
        for fk, fd in freqs.items():
            if fk not in ("50Hz", "60Hz"):
                bad(T, "%s 非法频率 %r" % (nm, fk))
            if not fd:
                bad(T, "%s[%s] 频率块为空" % (nm, fk))
        if d.get("curFreq") not in freqs:
            bad(T, "%s curFreq=%r 不在 freqs 中" % (nm, d.get("curFreq")))
        # 未知字段
        allk = set()
        for fd in freqs.values():
            allk |= set(fd)
        allk |= set(d.get("common") or {})
        unknown = allk - CE.ALL_KEYS
        if unknown:
            bad(T, "%s 未知字段 %s" % (nm, sorted(unknown)))
        # 数值字段类型
        for fk, fd in freqs.items():
            for k in CE.NUMERIC_KEYS:
                v = fd.get(k)
                if v in (None, ""):
                    continue
                if not re.fullmatch(r"[-+]?\d+(?:\.\d+)?", str(v)):
                    bad(T, "%s[%s] %s 非数值 %r" % (nm, fk, k, v))
            # 功率一致性
            try:
                kv = float(fd.get("prp_kva") or 0)
                kw = float(fd.get("prp_kw") or 0)
                if kv > 0 and kw > 0:
                    pf = kw / kv
                    if not (0.7 <= pf <= 1.01):
                        bad(T, "%s[%s] PF=%.3f 异常" % (nm, fk, pf))
                    else:
                        ok(T, "%s[%s] PF=%.2f" % (nm, fk, pf))
            except (TypeError, ValueError):
                bad(T, "%s[%s] 功率无法转数字" % (nm, fk))
    ok(T, "机型 %d 个，字段校验完成" % len(models))


# ============================================================
# T3 数据库往返
# ============================================================
def t3_roundtrip(models):
    T = "T3-往返"
    tmp = tempfile.mkdtemp(prefix="psheet_test_")
    old_db = None
    import model_store
    old_db = model_store.DB_PATH
    model_store.DB_PATH = os.path.join(tmp, "test.db")
    try:
        model_store.init()
        items = [{"name": m["name"], "model": m["model"], "family": m.get("family", ""),
                  "src": m.get("src", ""), "data": m["data"]} for m in models]
        added, updated, skipped = model_store.save_many(items, overwrite=True)
        if added == len(items) and updated == 0:
            ok(T, "批量导入 %d 条" % added)
        else:
            bad(T, "导入计数不符 added=%d updated=%d skipped=%d" % (added, updated, skipped))
        if model_store.count() == len(items):
            ok(T, "count()=%d" % model_store.count())
        else:
            bad(T, "count()=%d 期望 %d" % (model_store.count(), len(items)))
        # 逐条读回比对
        mism = 0
        for m in models:
            r = model_store.get(m["name"])
            if not r:
                bad(T, "读不到 %s" % m["name"]); mism += 1; continue
            if r["data"] != json.loads(json.dumps(m["data"])):
                bad(T, "%s 数据往返不一致" % m["name"]); mism += 1
        if mism == 0:
            ok(T, "全部 %d 条读回完全一致" % len(models))
        # 导出
        ex = model_store.export_all()
        if len(ex["models"]) == len(models):
            ok(T, "导出 %d 条" % len(ex["models"]))
        else:
            bad(T, "导出 %d 条 期望 %d" % (len(ex["models"]), len(models)))
        # 覆盖语义
        one = items[0]
        okk = model_store.save(one["name"], one["model"], one["data"], "", "",
                               overwrite=False)
        (ok if okk is False else bad)(T, "overwrite=False 应拒绝重复写入")
        # 删除
        model_store.delete(one["name"])
        if model_store.count() == len(items) - 1:
            ok(T, "删除生效")
        else:
            bad(T, "删除后 count 不对")
        # 列表查询
        lst = model_store.list_models()
        if len(lst) == len(items) - 1:
            ok(T, "list_models 返回 %d" % len(lst))
        else:
            bad(T, "list_models 返回 %d" % len(lst))
        # 过滤
        fam_list = [m.get("family") for m in models if m.get("family")]
        if fam_list:
            f0 = fam_list[0]
            q = model_store.list_models(family=f0)
            exp = sum(1 for m in items if m.get("family") == f0) - (
                1 if one.get("family") == f0 else 0)
            (ok if len(q) == exp else bad)(
                T, "family 过滤 %s: %d 期望 %d" % (f0, len(q), exp))
    finally:
        model_store.DB_PATH = old_db
        try:
            import shutil
            shutil.rmtree(tmp, ignore_errors=True)
        except Exception:
            pass


# ============================================================
# T5 频率串归一（用户 2026-10-20 定：紧凑格式）
# ============================================================
def t5_freq_norm():
    T = "T5-频率归一"
    cases = [
        ("50Hz / 1 P / PF 1.0", "50Hz / 1P / PF1.0"),
        ("60Hz / 3P / PF0.8",   "60Hz / 3P / PF0.8"),
        ("50Hz / 3P / PF 0.8",  "50Hz / 3P / PF0.8"),
        ("60 Hz / 3 P / P.F. 0.8", "60Hz / 3P / PF0.8"),
        ("50Hz/1P/PF1.0",       "50Hz / 1P / PF1.0"),
        ("50Hz",                "50Hz"),
        ("  50Hz  /  3P  ",     "50Hz / 3P"),
    ]
    for src, exp in cases:
        got = CE.normalize_freq_str(src)
        (ok if got == exp else bad)(T, "normalize_freq_str(%r)=%r 期望 %r" % (src, got, exp))
    # 空值
    for v in ("", None, "   "):
        got = CE.normalize_freq_str(v)
        (ok if got is None else bad)(T, "normalize_freq_str(%r) 应为 None，得到 %r" % (v, got))
    # 三元组解析
    for src, exp in [("50Hz / 1P / PF1.0", ("50", "1", "1.0")),
                     ("60Hz / 3P / PF0.8", ("60", "3", "0.8"))]:
        got = CE.parse_freq_triplet(src)
        (ok if got == exp else bad)(T, "parse_freq_triplet(%r)=%r 期望 %r" % (src, got, exp))


# ============================================================
# T6 频率推断（field > filename > voltage > default）
# ============================================================
def t6_freq_infer():
    T = "T6-频率推断"
    cases = [
        ("a.pptx", "60Hz / 3P / PF0.8", {}, "60Hz", "field"),
        ("CU330D5_60Hz.pptx", "", {}, "60Hz", "filename"),
        ("x.pptx", "", {"std_v": "480/277V"}, "60Hz", "voltage"),
        ("x.pptx", "", {"std_v": "440/254V"}, "60Hz", "voltage"),
        ("x.pptx", "", {"std_v": "400/230V"}, "50Hz", "voltage"),
        ("x.pptx", "", {"std_v": "415/240V"}, "50Hz", "voltage"),
        ("x.pptx", "", {}, "50Hz", "default"),
    ]
    for src, raw, flds, ehz, esrc in cases:
        got, gsrc = CE.infer_freq(src, raw, flds)
        (ok if got == ehz else bad)(T, "infer_freq(%r,%r,%r) 频率=%r 期望 %r"
                                    % (src, raw, flds, got, ehz))
        (ok if gsrc == esrc else bad)(T, "infer_freq(%r,%r,%r) 来源=%r 期望 %r"
                                      % (src, raw, flds, gsrc, esrc))
    # 拼装兜底
    cases2 = [
        ({"std_v": "400/230V"}, "50Hz / 3P / PF0.8"),
        ({"std_v": "230/115V"}, "50Hz / 3P / PF0.8"),   # 230/115 = 三相线/相电压
        ({"std_v": "240/120V"}, "50Hz / 3P / PF0.8"),
        ({"model": "AF40D6-1P"}, "50Hz / 1P / PF1.0"),
        ({"model": "G32"}, "50Hz / 3P / PF0.8"),
    ]
    for flds, exp in cases2:
        got = CE.synth_freq_str("50Hz", flds)
        (ok if got == exp else bad)(T, "synth_freq_str(%r)=%r 期望 %r" % (flds, got, exp))


# ============================================================
# T7 冲突检测（同型号多文件不同值 -> needs_review）
# ============================================================
def t7_conflict():
    T = "T7-冲突检测"

    def mk(name, model, common=None, freq=None, src=None):
        cm = {"e_brand": "Cummins", "e_model": "X1"}
        cm.update(common or {})
        fd = {"prp_kva": "100", "prp_kw": "80"}
        fd.update(freq or {})
        return {"source_file": src or name, "model_guess": model,
                "text": "", "tables": [], "_pre_fields": (cm, fd)}

    # 直接构造 parse 后的分组：用假 item 走 merge_items 太重，这里直接测合并逻辑产物
    raw = []
    for i, (lwh, src) in enumerate([("5812*2290*2520", "a.pptx"),
                                    ("6058*2438*2591", "b.pptx")]):
        raw.append({"source_file": src, "model_guess": "AS1375E5",
                    "text": "", "tables": [], "chars": 500, "ok": True,
                    "_fake": {"s_lwh": lwh, "std_v": "400/230V",
                              "prp_kva": "1250", "prp_kw": "1000",
                              "e_model": "16XTAA38-G33"}})
    # monkeypatch parse_item 返回构造字段
    orig = CE.parse_item
    try:
        CE.parse_item = lambda it: (dict(it["_fake"]), [])
        models, _rn = CE.merge_items(raw)
    finally:
        CE.parse_item = orig

    if not models:
        bad(T, "构造数据未产出机型")
        return
    m = models[0]
    (ok if m.get("needs_review") else bad)(T, "冲突机型应标 needs_review")
    (ok if m.get("_conflicts") else bad)(T, "应记录 _conflicts")
    if m.get("_conflicts"):
        fields = [c[0] for c in m["_conflicts"]]
        (ok if "s_lwh" in fields else bad)(T, "应抓到 s_lwh 冲突，实际 %r" % fields)
        (ok if m["data"]["common"]["s_lwh"] == "5812*2290*2520"
         else bad)(T, "冲突时保留第一个值，实际 %r" % m["data"]["common"].get("s_lwh"))

    # 不冲突时应为 False
    raw2 = [{"source_file": "a.pptx", "model_guess": "X1", "text": "", "tables": [],
             "chars": 500, "ok": True,
             "_fake": {"s_lwh": "100", "prp_kva": "10", "prp_kw": "8"}}]
    try:
        CE.parse_item = lambda it: (dict(it["_fake"]), [])
        m2, _ = CE.merge_items(raw2)
    finally:
        CE.parse_item = orig
    (ok if m2 and not m2[0].get("needs_review") else bad)(T, "无冲突机型不应标 needs_review")


# ============================================================
# T8 交叉核对（备用<主用 / PF 不一致 / 电压频率组合）
# ============================================================
def t8_cross():
    T = "T8-交叉核对"

    def mk(name, fd, common=None):
        return {"name": name,
                "data": {"model": name, "curFreq": "50Hz",
                         "freqs": {"50Hz": fd},
                         "common": common or {}}}

    ms = [
        mk("A", {"prp_kva": "100", "eps_kva": "90", "prp_kw": "80", "eps_kw": "72"}),
        mk("B", {"prp_kva": "100", "eps_kva": "110", "prp_kw": "80", "eps_kw": "99"}),
        mk("C", {"prp_kva": "100", "eps_kva": "110", "prp_kw": "80", "eps_kw": "77"}),
        mk("D", {"prp_kva": "100", "prp_kw": "80"}, {"std_v": "480/277V"}),
        mk("E", {"prp_kva": "100", "prp_kw": "80"}, {"std_v": "400/230V"}),
    ]
    sus = CE.cross_check(ms)
    txt = "\n".join(sus)

    (ok if "备用 kVA" in txt and "A[" in txt else bad)(T, "应抓出 A 备用<主用")
    b_bad = any(line.startswith("B[") and "备用 kW" in line for line in sus)
    (ok if not b_bad else bad)(T, "B 备用 kW 正常不应告警")
    (ok if "PF" in txt and "C[" in txt else bad)(T, "应抓出 C 主备 PF 不一致")
    (ok if "D[" in txt else bad)(T, "应抓出 D 50Hz+480V 可疑")
    (ok if "E[" not in txt else bad)(T, "E 50Hz+400V 正常不应告警")


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--data", required=True, help="convert_extract.py 输出的 models json")
    ap.add_argument("--extract", default="", help="原始提取目录（可选，做重复转换一致性）")
    args = ap.parse_args()

    payload = json.load(open(args.data, encoding="utf-8"))
    models = payload.get("models") or []

    print("=" * 60)
    print("T1 单元测试")
    print("=" * 60)
    t1_unit()
    print("通过 %d / 失败 %d" % (len(PASS), len(FAIL)))
    for x in FAIL:
        print("  ✗", x)

    print()
    print("=" * 60)
    print("T2 数据校验 (%d 机型)" % len(models))
    print("=" * 60)
    n0, f0 = len(PASS), len(FAIL)
    t2_data(models)
    print("通过 %d / 失败 %d" % (len(PASS) - n0, len(FAIL) - f0))
    for x in FAIL[f0:]:
        print("  ✗", x)

    print()
    print("=" * 60)
    print("T3 数据库往返")
    print("=" * 60)
    n1, f1 = len(PASS), len(FAIL)
    t3_roundtrip(models)
    print("通过 %d / 失败 %d" % (len(PASS) - n1, len(FAIL) - f1))
    for x in FAIL[f1:]:
        print("  ✗", x)

    for label, fn in (("T5 频率串归一（紧凑格式）", t5_freq_norm),
                      ("T6 频率推断（四级兜底）", t6_freq_infer),
                      ("T7 冲突检测（待人工确认）", t7_conflict),
                      ("T8 交叉核对（备用量/PF/电压频率）", t8_cross)):
        print()
        print("=" * 60)
        print(label)
        print("=" * 60)
        n0, f0 = len(PASS), len(FAIL)
        fn()
        print("通过 %d / 失败 %d" % (len(PASS) - n0, len(FAIL) - f0))
        for x in FAIL[f0:]:
            print("  ✗", x)

    # 可选：重复转换一致性
    if args.extract:
        print()
        print("=" * 60)
        print("T4 重复转换一致性（同输入两次结果必须一致）")
        print("=" * 60)
        raw = CE.load_items(args.extract)
        m1, _ = CE.merge_items(raw)
        m2, _ = CE.merge_items(raw)
        s1 = json.dumps([x["data"] for x in m1], sort_keys=True, ensure_ascii=False)
        s2 = json.dumps([x["data"] for x in m2], sort_keys=True, ensure_ascii=False)
        if s1 == s2:
            print("✓ 幂等：两次转换结果完全相同 (%d 机型)" % len(m1))
        else:
            print("✗ 非幂等：两次结果不同")
            FAIL.append("T4: 非幂等")

    print()
    print("=" * 60)
    print("总计: 通过 %d / 失败 %d" % (len(PASS), len(FAIL)))
    print("=" * 60)
    return 1 if FAIL else 0


if __name__ == "__main__":
    sys.exit(main())
