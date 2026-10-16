#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
★ AI Datasheet 提取器（DeepSeek）

流程：PDF/Excel → 抽文本 → 丢给 DeepSeek 按固定字段表返回 JSON → 规则兜底/校验。

设计：
- AI 为主（跟对话里人工提取同效果），规则解析补 AI 漏读的字段
- 字段表固定（与前端表单一致），AI 只能从文本里取，不得编造
- 失败/无 key 时自动退回纯规则
"""
import json
import os
import re

import datasheet_parser

# 目标字段表（与前端 index.html GROUPS 一致）
FIELD_SPEC = """
发动机信息：e_brand 发动机品牌 / e_model 发动机型号 / e_fuel 燃料 / e_asp 进气方式 /
e_cool 冷却方式 / e_gov 调速器 / e_cyl 气缸排列 / e_app 应用等级(Prime & Standby)
油耗(L/H)：fc_esp 100%ESP / fc_prp 100%PRP / fc_75 75%PRP / fc_50 50%PRP
基础数据：displ 排量 / cyl_arr 缸数/排列 / bore 缸径×行程 / cr 压缩比 / rpm 转速 / etype 发动机类型
排气系统：ex_t 排气温度 / ex_f 排气流量 / ex_bp 排气背压
燃油系统：fu_res 输油泵阻力 / fu_max 最大燃油流量
冷却系统：co_air 散热器空气流量 / co_oil 机油容量 / co_cool 冷却液容量
电气系统：el_v 起动机电压 / el_bat 启动电池 / el_alt 充电发电机
进气系统：in_flt 空滤类型 / in_res 最大进气阻力 / in_flow 燃烧空气流量 / in_t 最高进气温度
电气输出(若表中有)：prp_kva / eps_kva / prp_kw / eps_kw / freq 频率·相数·PF / std_v 标准电压 / opt_v 可选电压
尺寸(如资料给出)：o_lwh/s_lwh 长宽高 · o_wt/s_wt 重量 · o_tank/s_tank 油箱
"""

# ★ 合法字段白名单（AI 返回值只准含这些 key，防乱填/幻觉）
VALID_FIELDS = {
    "e_brand", "e_model", "e_fuel", "e_asp", "e_cool", "e_gov", "e_cyl", "e_app",
    "fc_esp", "fc_prp", "fc_75", "fc_50",
    "displ", "cyl_arr", "bore", "cr", "rpm", "etype",
    "ex_t", "ex_f", "ex_bp", "fu_res", "fu_max",
    "co_air", "co_oil", "co_cool", "el_v", "el_bat", "el_alt",
    "in_flt", "in_res", "in_flow", "in_t",
    "prp_kva", "eps_kva", "prp_kw", "eps_kw", "freq", "std_v", "opt_v",
    "o_lwh", "s_lwh", "o_wt", "s_wt", "o_tank", "s_tank",
    "o_db1", "o_db7", "s_db1", "s_db7",
}
# 频率专属字段（不许留在 common）
FREQ_ONLY_FIELDS = {
    "freq", "prp_kva", "eps_kva", "prp_kw", "eps_kw",
    "rpm", "fc_esp", "fc_prp", "fc_75", "fc_50",
}

SYSTEM_PROMPT = """你是发电机组/发动机参数提取专家。用户会给你一份 Datasheet 的纯文本（从 PDF/Excel 抽出，可能有排版错乱）。

任务：把数据按给定字段表提取成 JSON。

★ 频率分流（最重要）：
同一份 Datasheet 常常同时给出 50Hz 和 60Hz 两套数据（如功率、转速、油耗不同）。
**必须把两个频率的数据分开输出**，结构如下：
{
  "50Hz": { 字段... },
  "60Hz": { 字段... },
  "common": { 两频率相同的字段... }
}
- 只在 Datasheet 中出现的频率才输出该键（比如只有 50Hz 数据，就只输出 "50Hz"）。
- 两频率都相同的字段（品牌/缸数/缸径行程/压缩比/排量/排气/燃油/冷却/电气等）放 "common"。
- 会随频率变的字段（prp_kva/eps_kva/prp_kw/eps_kw/rpm/fc_*/功率相关，以及任何随转速变的流量）放各自的频率键下。
- 如果某字段在 "common" 和频率键下都能放，优先放频率键（更精确）。

★ 单位归一（必须执行，直接输出归一后的单位，不要保留原文单位）：
- 压力（ex_bp 排气背压 / in_res 进气阻力 / fu_res 输油泵阻力）：一律转 **kPa**，保留 1 位小数
  · 1 mmHg = 0.1333 kPa　· 1 inHg = 3.386 kPa　· 1 mmH2O = 0.00981 kPa
  · 1 inH2O = 0.249 kPa　· 1 psi = 6.895 kPa　· 1 mbar = 0.1 kPa　· 1 bar = 100 kPa
  · 纯 kPa 原样保留
- 流量（ex_f 排气流量 / in_flow 燃烧空气流量 / co_air 散热器空气流量）：一律转 **m³/min**，保留 1 位小数
  · 1 L/S = 0.06 m³/min　· 1 cfm = 0.02832 m³/min　· 1 m³/min 原样
- 温度：一律 **°C**。°F -> °C：(F-32)*5/9，取整。例 "1094°F" -> "590°C"
- 容量（co_oil 机油 / co_cool 冷却液）：一律 **L**，格式 "13L"。
  · 1 US gal = 3.785 L　· 若给区间 "7.0-5.5"，取较大的总容量值或按原表主值
- 功率：kW/kVA 保留原值；只给 BHP 时 BHP x 0.746 = kW（保留整数）
- 排量 displ：如 "3.3L"；缸径行程 bore：如 "95x115mm"
- 压缩比 cr：如 "18:1"　· 转速 rpm：如 "1500rpm"
- 油耗 fc_*：只要数字（L/H），如 "13"
- 最大燃油流量 fu_max：务必是油量单位（L/H）；若原文是 US gph，1 gph = 3.785 L/H

★ 反幻觉（极重要，违反即失败）：
- 只能从文本中提取，绝对不许编造、猜测、推算未写明的内容。找不到的字段一律不输出。
- 特别注意以下陷阱，这些字段常被误填：
  · co_air（散热器空气流量）：不要把「Radiated Heat to Ambient / Heat Rejection」的热量值（kW/BTU）填进来。
  · in_t（最高进气温度）：不要把测试基准条件的「Air Temperature : 25°C」填进来，那是测试条件不是最高进气温度。找不到就空着。
  · ex_t（排气温度）：填实际排气温度，不要填测试条件温度。
  · 各字段不要跨行取相邻行的值。
- 若某字段原文有多个档位（如进气阻力分「脏滤芯/正常/重载」），取**正常/标准工况**那一档，并在值后可加括号注明，如 "3.0kPa(正常)"。

铁律：
1. 找不到的字段直接不输出（不要写 null / "-" / 空字符串）。
2. 术语用标准英文：Diesel / Water-cooled / Turbocharged & Intercooled /
   Electronic Governor / Prime & Standby / 4-stroke diesel 等。
3. 只返回一个 JSON 对象。不要 markdown 代码块，不要解释。

字段表：
%s
""" % FIELD_SPEC


def _get_key():
    key = os.environ.get("DEEPSEEK_API_KEY")
    if key:
        return key.strip()
    # 复用 genset-ai 的 .env
    for p in ("/root/genset-ai/.env", os.path.expanduser("~/.config/paramsheet/.env")):
        if os.path.exists(p):
            with open(p) as f:
                for line in f:
                    if line.startswith("DEEPSEEK_API_KEY"):
                        return line.split("=", 1)[1].strip().strip('"').strip("'")
    return None


def _call_deepseek(text, model="deepseek-chat", timeout=120):
    import urllib.request
    key = _get_key()
    if not key:
        return None
    # 文本过长时截断（DeepSeek 上下文足够，但控成本）
    body = {
        "model": model,
        "messages": [
            {"role": "system", "content": SYSTEM_PROMPT},
            {"role": "user", "content": text[:24000]},
        ],
        "temperature": 0,
        "max_tokens": 2000,
        "response_format": {"type": "json_object"},
    }
    req = urllib.request.Request(
        "https://api.deepseek.com/chat/completions",
        data=json.dumps(body).encode(),
        headers={"Content-Type": "application/json",
                 "Authorization": "Bearer " + key},
    )
    with urllib.request.urlopen(req, timeout=timeout) as r:
        resp = json.loads(r.read().decode())
    return resp["choices"][0]["message"]["content"]


def parse_datasheet_ai(filename, content, model_hint=""):
    """AI 为主 + 规则兜底。返回 (freqs, meta)。

    freqs 结构：{"50Hz": {..}, "60Hz": {..}, "common": {..}}
    """
    text = datasheet_parser.extract_text(filename, content)
    meta = {"chars": len(text), "ai": False, "rule_fields": 0, "ai_fields": 0}
    if not text.strip():
        return {}, meta

    ai_obj = {}
    try:
        raw = _call_deepseek(text)
        if raw:
            raw = raw.strip()
            raw = re.sub(r"^```(?:json)?|```$", "", raw).strip()
            parsed = json.loads(raw)
            if isinstance(parsed, dict):
                ai_obj = parsed
                meta["ai"] = True
    except Exception as e:
        meta["ai_error"] = str(e)[:200]

    # 结构识别：新式（含频率键）还是扁平（全部当 common）
    freq_keys = [k for k in ai_obj.keys() if re.match(r"^(50|60)\s*Hz$", str(k), re.I)]
    if freq_keys or "common" in ai_obj:
        out = {}
        for k, v in ai_obj.items():
            kk = re.sub(r"\s*Hz", "Hz", str(k), flags=re.I)
            if isinstance(v, dict):
                out[kk] = {fk: fv for fk, fv in v.items()
                           if fv not in (None, "", "-", "null", "N/A")}
        # common 去重：同名字段已在频率键下出现的，从 common 移除
        if "common" in out and isinstance(out["common"], dict):
            freq_field_names = set()
            for kk, vv in out.items():
                if kk != "common" and isinstance(vv, dict):
                    freq_field_names |= set(vv.keys())
            out["common"] = {fk: fv for fk, fv in out["common"].items()
                             if fk not in freq_field_names
                             and fk not in FREQ_ONLY_FIELDS}
        ai_fields = sum(len(v) for v in out.values() if isinstance(v, dict))
    else:
        flat = {k: v for k, v in ai_obj.items()
                if v not in (None, "", "-", "null", "N/A")}
        out = {"common": flat} if flat else {}
        ai_fields = len(flat)
    meta["ai_fields"] = ai_fields

    # 规则兜底：只补 common 里缺的（规则不区分频率，避免污染频率键）
    # ⚠️ 频率专属字段绝不从规则补进 common（会跟频率键打架）
    rule_fields = datasheet_parser.parse_text(text, model_hint)
    meta["rule_fields"] = len(rule_fields)
    if rule_fields:
        common = out.setdefault("common", {})
        for k, v in rule_fields.items():
            if k in FREQ_ONLY_FIELDS:
                continue
            common.setdefault(k, v)

    # 兜底：AI 没给出任何频率键时，用规则里的频率推断建一个
    if not any(re.match(r"^(50|60)Hz$", k) for k in out):
        freq = ""
        m = re.search(r"\b(50|60)\s*Hz\b", text, re.I)
        if m and "freq" in rule_fields:
            freq = m.group(1) + "Hz"
        if freq:
            # 把 common 里「随频率变」的字段挪到频率键
            moved = {}
            common = out.get("common", {})
            for k in list(common.keys()):
                if k in FREQ_ONLY_FIELDS:
                    moved[k] = common.pop(k)
            out[freq] = moved

    # ★ common 最终去重：清掉所有频率专属字段 + 与频率键重名的
    if isinstance(out.get("common"), dict):
        freq_names = set()
        for kk, vv in out.items():
            if kk != "common" and isinstance(vv, dict):
                freq_names |= set(vv.keys())
        out["common"] = {k: v for k, v in out["common"].items()
                         if k not in FREQ_ONLY_FIELDS and k not in freq_names}
        if not out["common"]:
            out.pop("common", None)

    # ★ 白名单过滤：清掉 AI 幻觉/乱填的非法 key
    out = _sanitize(out, text)
    return out, meta


def _sanitize(out, text=""):
    """白名单过滤 + 简单值校验（型号不得当品牌、单位归一）。"""
    clean = {}
    for fk, data in out.items():
        if not isinstance(data, dict):
            continue
        d = {}
        for k, v in data.items():
            if k not in VALID_FIELDS:
                continue
            sv = str(v).strip()
            # 占位符/无效值一律丢弃（TBD / N.A. / N/A / - / 空）
            if not sv or sv.lower() in ("null", "n/a", "na", "n.a.", "tbd", "-", "none", "—", "?"):
                continue
            # 纯占位符组合（如 "TBD/TBD"）
            if re.fullmatch(r"(tbd|n\.?a\.?|—|-|\?)(\s*[/、,]\s*(tbd|n\.?a\.?|—|-|\?))*", sv, re.I):
                continue
            # 型号不能当品牌：品牌值若含数字+连字符/与型号相同 -> 丢弃
            if k == "e_brand" and re.search(r"\d", sv) and len(sv) > 6:
                continue
            # fu_max 不允许压力单位（MPa/kPa/bar）
            if k == "fu_max" and re.search(r"(mpa|kpa|bar)\s*$", sv, re.I):
                continue
            # 容量字段归一为 "13L"
            if k in ("co_oil", "co_cool"):
                mm = re.match(r"^([\d.]+)\s*(l|liter|litre|liters|litres)\b", sv, re.I)
                if mm:
                    sv = mm.group(1) + "L"
            d[k] = sv
        if d:
            clean[fk] = d
    # 若 e_brand 与 e_model 相同，丢掉 e_brand
    for d in clean.values():
        if d.get("e_brand") and d.get("e_model") and d["e_brand"] == d["e_model"]:
            d.pop("e_brand", None)
    return clean
