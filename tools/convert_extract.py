#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
把 extract_batch.py 提取出来的 JSON 转换成参数单生成器的机型格式

输入：items/*.json（extract_batch 输出）
输出：{models:[...]} ，每项：
    {name, model, family, src, data:{model,curFreq,freqs:{'50Hz':{...},'60Hz':{...}},common:{...}}}

用法：
    python convert_extract.py --src <items目录> --out models_import.json [--dry-run] [--verify]

设计要点：
  · 只信 PPTX（表格结构化）；PDF 仅作缺失字段的兜底（可选 --use-pdf）
  · 标签按「归一化后精确匹配 + 别名表」，不做模糊包含，避免 e_gov 抢到别家
  · 单位与值分离：表单标了「连单位输入」的字段整串保留，其余拆数值
  · TBA / N/A / - / 待定 / 空 一律视为无数据（不写键）
"""
import argparse
import json
import os
import re
import sys

# ============================================================
# 常量
# ============================================================
NULL_TOKENS = {
    "", "-", "--", "—", "/", "n/a", "na", "tba", "tbd", "tbc",
    "待定", "无", "不适用", "nil", "none",
}

# 随频率变化的字段（与前端 FREQ_FIELDS 保持一致）
FREQ_FIELDS = ("freq", "prp_kva", "eps_kva", "prp_kw", "eps_kw",
               "rpm", "fc_esp", "fc_prp", "fc_75", "fc_50")

# ------------------------------------------------------------
# 标签别名表：归一化标签 -> 表单字段
# 归一化 = 小写 + 去所有非字母数字字符（含全角、\xa0、换行）
# ⚠️ 必须精确匹配，不做 in 判断
# ------------------------------------------------------------
LABELS = {
    # --- 第1页 功率 ---
    "power": "__power",                      # 特殊处理：kVA/kW 两行
    "frequencyphasepowerfactor": "freq",
    "frequency": "freq",
    "standardvoltage": "std_v",
    "availablevoltage": "opt_v",
    # --- 外观尺寸 ---
    "dimensionlwhmm": "__lwh",
    "dimensionlwh": "__lwh",
    "weightkg": "__weight",
    "weight": "__weight",
    "fueltankl": "__tank",
    "fueltank": "__tank",
    "soundpressurelevel75load1m dba": "o_db1",
    "soundpressurelevel75load7m dba": "o_db7",
    # --- 动力概览 ---
    "enginebrand": "e_brand",
    "enginemodel": "e_model",
    "fuel": "e_fuel",
    "aspiration": "e_asp",
    "cooling": "e_cool",
    "governor": "e_gov",
    "cylarrangement": "e_cyl",
    "application": "__ignore",
    "fuelconsumption": "__fuelcons",         # 多行：100% ESP / 100% PRP / 75% PRP / 50% PRP
    "fuelconsumptionlh": "__fuelcons",
    # --- 第2页 基础数据 ---
    "displacement": "displ",
    "noofcylarrangement": "cyl_arr",
    "borexstroke": "bore",
    "compressionratio": "__ignore",
    "enginespeed": "rpm",
    "enginetype": "__ignore",
    # --- 排气 ---
    "exhauststackgastemperaturemax": "ex_t",
    "exhaustgasflowratemax": "ex_f",
    "exhaustsystembackpressuremaximumallowable": "ex_bp",
    # --- 燃油 ---
    "maximumrestrictionatlifttump": "fu_res",
    "maximumrestrictionatliftpump": "fu_res",
    "maxfuelflow": "fu_max",
    # --- 冷却 ---
    "radiatorairflow": "co_air",
    "totaloilcapacity": "co_oil",
    "totalcoolantcapacity": "co_cool",
    "totalcoolantcapacityonlyengine": "co_cool",
    # --- 进气 ---
    "airfilter": "in_flt",
    "maxcombustionairintakerestriction": "in_res",
    "combustionairinletflowrate": "in_flow",
    "maxallowablecombustionairinlettemp": "in_t",
    # --- 电气 ---
    "startermotorratedvoltage": "el_v",
    "startbattery": "el_bat",
    "batterychargingalternator": "el_alt",
    # --- 电机 ---
    "insulationtemperaturerise": "g_insul",
    "excitationtype": "g_exc",
    "windingleads": "g_lead",
    "windingconnectionsstandard": "g_conn",
    "windingpitch": "g_pitch",
    "numberofphase": "g_phase",
    "type": "__ignore",                       # Brushless
    "noofbearings": "g_brg",
    "couplingsystem": "__ignore",
    "maximumoverspeed": "__ignore",
    "coatingtype": "__ignore",
    "iprating": "__ignore",

    # ===== 中文版数据表（康明斯国产机型，如 12KTA58-G32）=====
    # 中文表头归一化后是纯中文 key
    "转速": "rpm",
    "柴油机型号": "e_model",
    "机组型号": "e_model",
    "发动机型号": "e_model",
    "柴油机型式": "__ignore",
    "燃烧方式": "__ignore",
    "气缸型式": "__ignore",
    "气缸数": "cyl_arr",
    "缸径行程": "bore",
    "缸径": "__ignore",
    "行程": "__ignore",
    "压缩比": "__ignore",
    "发火次序": "__ignore",
    "喷油提前角": "__ignore",
    "调速器": "e_gov",
    "净重": "o_wt",
    "外形尺寸": "o_lwh",
    "冷却液容量": "co_cool",
    "进气流量": "in_flow",
    "排气流量": "ex_f",
    "排气温度": "ex_t",
    "燃料": "e_fuel",
    "机油容量": "co_oil",
    "机油总容量": "co_oil",
    "启动电压": "el_v",
    "蓄电池": "el_bat",
    "充电机": "el_alt",
    "进气方式": "e_asp",
    "冷却方式": "e_cool",
    "排量": "displ",
    "缸数": "cyl_arr",
}

# 值归一化：英文 -> 前端下拉的规范值
VALUE_MAP = {
    "g_exc": {
        "selfexcited": "Self-excited",
        "pmg": "PMG",
        "separatelyexcited": "Separately excited",
    },
    "g_conn": {
        "star": "Star",
        "delta": "Delta",
        "serie": "Star",          # 单相/串联写法，前端无 Serie 选项
        "series": "Star",
    },
    "g_pitch": {
        "twothirds": "Two thirds",
        "fivesixths": "Five sixths",
    },
    "g_phase": {
        "singlephase": "Single phase",
        "threephase": "Three phase",
    },
    "g_brg": {
        "singlebearing": "Single bearing",
        "doublebearing": "Double bearing",
    },
    "e_asp": {
        "turbocharged": "Turbocharged",
        "naturallyaspirated": "Naturally aspirated",
        "turbochargedaftercooled": "Turbocharged & Aftercooled",
        "turbochargedintercooled": "Turbocharged & Intercooled",
        "naturallyaspirated": "Naturally aspirated",
    },
    "e_cool": {
        "watercooled": "Water-cooled",
    },
    "e_fuel": {
        "diesel": "Diesel",
    },
}

# 中文版报表值（zh 模板 cell_value_map 用；此处只做英文，中文由模板映射）

# 单位后缀（长优先，避免 "m³/min" 被 "m" 抢先）
UNITS = ("m³/min", "m3/min", "m³/s", "m3/s", "kg/min", "L/S", "L/s",
         "kPa", "kW", "kVA", "kW/h", "L/H", "L/h", "rpm", "RPM",
         "°C", "mm", "kg", "Ah", "CCA", "V", "L", "A")


def norm_label(s):
    """标签归一化：小写 + 去所有非字母数字"""
    if s is None:
        return ""
    s = str(s).replace("\xa0", " ").replace("\u3000", " ")
    s = s.lower()
    return re.sub(r"[^0-9a-z]", "", s)


def norm_label_cn(s):
    """中文表头归一化：多行取首行 + 去圈号/项目符号/●○◆等装饰符 + 去"ㅇ"等韩文字符"""
    if s is None:
        return ""
    s = str(s).replace("\xa0", " ").replace("\u3000", " ")
    # 多行：中文表头形如 "转速\nEngine Speed"，取第一行
    first = s.split("\n")[0]
    if not first.strip() and len(s.split("\n")) > 1:
        first = s.split("\n")[1]
    t = re.sub(r"[\u25cb\u25cf\u25c6\u25c7\u25a0\u25a1\u2460-\u2473\u3007\u30fb\u3160-\u318f\u2022\u00b7]", "", first)
    t = re.sub(r"[^0-9A-Za-z\u4e00-\u9fff]", "", t)
    return t.lower()


def clean(v):
    """清洗值：去空白；空值返回 None"""
    if v is None:
        return None
    s = str(v).replace("\xa0", " ").strip()
    s = re.sub(r"[ \t]+", " ", s)
    s = re.sub(r"\n+", "\n", s)
    if not s:
        return None
    if s.lower() in NULL_TOKENS:
        return None
    # 纯占位符
    if s in ("12345", "00000"):
        return None
    return s


def split_unit(s):
    """'37.8 L' -> ('37.8', 'L')；无法拆则返回 (原串, '')"""
    if s is None:
        return None, ""
    t = s.strip()
    for u in UNITS:
        if t.lower().endswith(u.lower()):
            num = t[: len(t) - len(u)].strip()
            if num and re.fullmatch(r"[-+]?\d+(?:\.\d+)?", num):
                return num, u
    return t, ""


def keep_with_unit(s):
    """表单标了「连单位输入」：去掉空格但保留单位，如 '111.1 m³/min' -> '111.1m³/min'"""
    if s is None:
        return None
    s = s.strip()
    s = re.sub(r"\s*(m³/min|m3/min|m³/s|m3/s|kg/min|L/S|L/s|kPa|mm|kg|L|°C|A|V)\s*$",
               r"\1", s, flags=re.I)
    return s


def parse_power_series(text):
    """从 '50Hz / 3P / PF0.8' 提取频率/相数/PF，原样返回整串给 freq 字段"""
    return None  # freq 字段直接用整串，前端模板要 '50Hz / 3P / PF0.8'


def fmt_battery(raw):
    """启动电池格式化，对齐前端 collect()：`${n}X${ah}Ah(${cca}CCA)`

    输入形态：
        "2×180Ah"            -> "2X180Ah(TBA CCA)"
        "60(580CCA)"         -> "X60Ah(580CCA)"   (无个数)
        "2 / 60 / 720"       -> "2X60Ah(720CCA)"
        "2X180Ah(720CCA)"    -> 原样
    用户指定：缺 CCA 写 "（TBA CCA）" 语义 → 这里统一成 "X60Ah(TBA CCA)"。
    """
    if raw is None:
        return None
    s = str(raw).strip()
    if not s:
        return None
    # 已经是标准格式
    if re.fullmatch(r"\d*\s*(?:[xX×]\s*)?\d+(?:\.\d+)?\s*Ah\s*\(\s*(?:\d+(?:\.\d+)?\s*CCA|TBA\s*CCA)\s*\)",
                    s, re.I):
        return re.sub(r"\s+", "", s).replace("×", "X").replace("x", "X")

    n = ah = cca = None
    # 2×180Ah / 2x180Ah / 2 / 180
    m = re.search(r"(\d+)\s*[xX×]\s*(\d+(?:\.\d+)?)\s*Ah", s, re.I)
    if m:
        n, ah = m.group(1), m.group(2)
    else:
        m = re.search(r"(\d+)\s*/\s*(\d+(?:\.\d+)?)\s*/\s*(\d+(?:\.\d+)?)", s)
        if m:
            n, ah, cca = m.group(1), m.group(2), m.group(3)
        else:
            m = re.search(r"(\d+)\s*/\s*(\d+(?:\.\d+)?)", s)
            if m:
                n, ah = m.group(1), m.group(2)
            else:
                m = re.search(r"^(\d+(?:\.\d+)?)\s*Ah", s, re.I)
                if m:
                    ah = m.group(1)
    # CCA 单独出现（如 "60(580CCA)"）
    if cca is None:
        mc = re.search(r"(\d+(?:\.\d+)?)\s*CCA", s, re.I)
        if mc:
            cca = mc.group(1)
    # "60(580CCA)" -> ah=60, cca=580
    if ah is None:
        m2 = re.search(r"(\d+(?:\.\d+)?)\s*\(\s*\d", s)
        if m2:
            ah = m2.group(1)
    if ah is None:
        return s
    cca_s = ("%sCCA" % cca) if cca else "TBA CCA"
    if n:
        return "%sX%sAh(%s)" % (n, ah, cca_s)
    return "%sAh(%s)" % (ah, cca_s)


# ============================================================
# 频率串归一（用户 2026-10-20 定：紧凑格式）
#   "50Hz / 1 P / PF 1.0" -> "50Hz / 1P / PF1.0"
#   "60Hz / 3P / PF0.8"   -> "60Hz / 3P / PF0.8"
#   只规整分隔与小数，不改数值
# ============================================================
def normalize_freq_str(s):
    if s is None:
        return None
    t = str(s).replace("\xa0", " ").strip()
    if not t:
        return None
    # 分隔符统一
    t = re.sub(r"\s*[/／]\s*", " / ", t)
    # 相数：1 P / 3 P -> 1P / 3P
    t = re.sub(r"(\d)\s+([Pp])(?![A-Za-z])", r"\1\2", t)
    # PF：PF 0.8 / PF 1.0 / P.F. 0.8 -> PF0.8 / PF1.0
    t = re.sub(r"P\.?\s*F\.?\s*", "PF", t, flags=re.I)
    t = re.sub(r"PF\s+(\d)", r"PF\1", t, flags=re.I)
    # Hz：50 Hz -> 50Hz
    t = re.sub(r"(\d)\s+([Hh][Zz])", r"\1\2", t)
    # 压掉多余空格
    t = re.sub(r"\s{2,}", " ", t).strip()
    # 单相 1P 的 PF 固定 1.0（源里有时写 PF 1 或漏）
    return t


def parse_freq_triplet(s):
    """把归一后的频率串拆成 (hz, phase, pf)，取不到返回 None"""
    t = normalize_freq_str(s)
    if not t:
        return None
    mh = re.search(r"(50|60)\s*Hz", t, re.I)
    mp = re.search(r"(\d)\s*P(?![A-Za-z])", t)
    mf = re.search(r"PF\s*(\d+(?:\.\d+)?)", t, re.I)
    return (mh.group(1) if mh else None,
            mp.group(1) if mp else None,
            mf.group(1) if mf else None)


def synth_freq_str(freq_key, fields):
    """freq 字段缺失时，用「频率 + 相数 + PF」拼装（用户 2026-10-20 认的兜底）。
    相数从 std_v 猜：三相电压带斜杠（400/230）或已知三相电压 -> 3P；单相 -> 1P
    PF：相数 1 -> 1.0；三相默认 0.8（缺省值，报告里会标明是拼装的）
    """
    if not freq_key:
        return None
    hz = freq_key.replace("Hz", "")
    st = str(fields.get("std_v") or "")
    # 单相判定优先级：型号 -1P > 电压形态
    mdl = str(fields.get("model") or "") + " " + str(fields.get("e_model") or "")
    if re.search(r"-1P\b|\b1P-", mdl, re.I):
        phase = "1"
    else:
        m = re.search(r"(\d{2,4})\s*/\s*(\d{2,4})", st)
        if m:
            # 400/230、230/115、220/127、240/120 都是三相（线/相电压）
            phase = "3"
        else:
            mv = re.search(r"(\d{2,4})", st)
            volt = int(mv.group(1)) if mv else 0
            if volt:
                phase = "3" if volt >= 380 else "1"
            else:
                phase = "3"
    pf = "1.0" if phase == "1" else "0.8"
    return "%sHz / %sP / PF%s" % (hz, phase, pf)


def infer_freq(src, freq_raw, fields=None):
    """推断频率（用户 2026-10-20 定：多级兜底 + 记警告）
    1) freq 字段原串
    2) 文件名里的 50Hz / 60Hz
    3) 标准电压：440/480 -> 60Hz；380/400/415 -> 50Hz
    4) 都不中 -> 50Hz（调用方需记警告）
    返回 (freq, source)  source ∈ {"field","filename","voltage","default"}
    """
    m = re.search(r"(50|60)\s*Hz", str(freq_raw or ""), re.I)
    if m:
        return m.group(1) + "Hz", "field"
    m = re.search(r"(50|60)\s*Hz", str(src or ""), re.I)
    if m:
        return m.group(1) + "Hz", "filename"
    v = str((fields or {}).get("std_v") or "")
    mv = re.search(r"(\d{3,4})", v)
    if mv:
        volt = int(mv.group(1))
        if volt >= 440:
            return "60Hz", "voltage"
        if volt in (380, 400, 415, 220, 230, 240):
            return "50Hz", "voltage"
    return "50Hz", "default"


def guess_family(model, src):
    """系列：AF / AS / GC / 其他"""
    m = (model or "").upper()
    for f in ("GC", "AS", "AF"):
        if m.startswith(f):
            return f
    # 从路径猜
    up = (src or "").upper()
    for f in ("GC", "AS", "AF"):
        if re.search(r"(^|[\\/])" + f + r"[\\/-]", up):
            return f
    return "其他"


def guess_freq(src, data_freq):
    """频率：优先数据里的 freq 字段，其次路径"""
    if data_freq:
        m = re.search(r"(50|60)\s*Hz", data_freq, re.I)
        if m:
            return m.group(1) + "Hz"
    up = (src or "").upper()
    # 先判 1P / -1P 这类单相目录（"50HZ-1P" 同时含 50HZ，顺序无所谓）
    if "50HZ" in up or "50-1P" in up:
        return "50Hz"
    if "60HZ" in up or "60HZ-1P" in up or "60-1P" in up:
        return "60Hz"
    return None


def has_data(fields):
    """是否含有实质业务字段（不是只有电机绝缘那几项）"""
    core = ("prp_kva", "eps_kva", "prp_kw", "eps_kw", "e_model", "e_brand",
            "std_v", "displ", "rpm", "fc_prp")
    return any(fields.get(k) for k in core)


# ============================================================
# 主解析：把一个 extracted JSON 转成字段字典
# ============================================================
def parse_item(item):
    """
    返回 (fields, notes)
    fields: 扁平字段字典（含 freq 字段与 common 字段，尚未拆分）
    notes:  诊断信息列表
    """
    notes = []
    src = item.get("source_file") or ""
    tables = item.get("tables") or []

    f = {}
    _DIMS_ROWS = {}   # 每张尺寸表已读行数
    _full_text = item.get("text") or ""   # 中文表单行化文本，供 std_v 等补充解析

    def put(k, v, with_unit=False):
        if k in ("__ignore",) or k.startswith("__"):
            return
        if v is None:
            return
        if with_unit:
            v = keep_with_unit(v)
        if k in f and f[k]:
            return  # 先到先得（第1页优先于第2页重复字段）
        f[k] = v

    for t in tables:
        rows = t.get("rows") or []
        # ---- 中文功率矩阵：首行含 Engine Speed 且值里带 Prime/Standby ----
        flat = " ".join(" ".join(str(c) for c in r) for r in rows[:6])
        if re.search(r"engine\s*speed", flat, re.I) and re.search(r"standby\s*power", flat, re.I):
            for row in rows:
                cs = [clean(c) for c in row]
                txt = " ".join(x for x in cs if x)
                m_lab = re.search(r"(prime|standby)\s*power", txt, re.I)
                if not m_lab:
                    continue
                # 表结构：转速 | 运行模式 | 发动机功率kW | 机组功率kW | 机组功率kVA
                # 必须按列位取，不能靠 nums 顺序（含转速列，且单位行会干扰）
                idx = [i for i, x in enumerate(cs) if x]
                nums = [cs[i] for i in idx if re.fullmatch(r"\d+(?:\.\d+)?", cs[i] or "")]
                if len(nums) >= 3:
                    kw_g, kva_g = nums[-2], nums[-1]
                    if m_lab.group(1).lower() == "prime":
                        f["prp_kw"], f["prp_kva"] = kw_g, kva_g
                    else:
                        f["eps_kw"], f["eps_kva"] = kw_g, kva_g
            continue
        # 记录表头：OPEN / SILENT 列位置（尺寸/重量/油箱表用）
        col_hdr = {}
        for hc_i, hc in enumerate(rows[0] if rows else []):
            hn = norm_label(hc)
            if hn in ("open", "silent"):
                col_hdr[hc_i] = hn
        # 尺寸表首行常是 ['', 'OPEN', 'SILENT']，标签列是空的
        dims_mode = bool(col_hdr)
        for row in rows:
            cells = [c for c in row]
            if len(cells) < 2:
                continue
            # ---- 尺寸表：整表按行序取（第2行=尺寸，第3行=重量，第4行=油箱）----
            if dims_mode:
                vals_row = {i: clean(c) for i, c in enumerate(cells) if clean(c) is not None}
                if not vals_row:
                    continue
                rows_seen = _DIMS_ROWS.setdefault(id(t), 0)
                _DIMS_ROWS[id(t)] = rows_seen + 1
                # 跳过表头行（只有 OPEN/SILENT 字样）
                if all(norm_label(v) in ("open", "silent") for v in vals_row.values()):
                    continue
                pre = {0: "lwh", 1: "wt", 2: "tank"}.get(rows_seen - 1)
                if pre is None:
                    # 超过 3 行：按内容猜（含 * 或 x 的长串=尺寸）
                    for i, v in vals_row.items():
                        if re.fullmatch(r"\d{3,5}\s*[*x×]\s*\d{3,5}.*", v):
                            pre = "lwh"
                            break
                    if pre is None:
                        continue
                for i, v in vals_row.items():
                    side = "o_" if col_hdr.get(i) == "open" else (
                        "s_" if col_hdr.get(i) == "silent" else None)
                    if side and not f.get(side + pre):
                        f[side + pre] = v
                continue

            # ---- 无标签的数值行（尺寸/重量/油箱）→ 按表头列位置取 ----
            nonempty_idx = [i for i, c in enumerate(cells) if clean(c) is not None]
            if col_hdr and len(nonempty_idx) >= 1 and not any(
                    clean(cells[i]) in ("OPEN", "SILENT") for i in nonempty_idx):
                # 是否整行都无标签文本（全为数字/尺寸）
                vals_only = [clean(cells[i]) for i in nonempty_idx]
                if all(re.fullmatch(r"[\d\s.,*/x×\-]+", v or "") for v in vals_only):
                    pending = f.get("__dim_pending")
                    if pending:
                        pre = {"lwh": "lwh", "wt": "wt", "tank": "tank"}[pending]
                        for i, v in zip(nonempty_idx, vals_only):
                            side = "o_" if col_hdr.get(i) == "open" else (
                                "s_" if col_hdr.get(i) == "silent" else None)
                            if side and not f.get(side + pre):
                                f[side + pre] = v
                    continue
            # 找到第一个非空 cell 作为标签，其余非空作为值
            label = None
            vals = []
            for c in cells:
                c2 = clean(c)
                if c2 is None:
                    continue
                if norm_label(c2) in ("open", "silent"):
                    continue
                if label is None:
                    label = c2
                else:
                    vals.append(c2)
            if label is None:
                continue
            key = LABELS.get(norm_label(label))
            if key is None:
                # 中文表兜底：多行表头取首行、去圈号后匹配
                key = LABELS.get(norm_label_cn(label))

            # ---- 特殊：Power 行（kVA / kW），值里带单位需拆 ----
            if key == "__power":
                if not vals:
                    continue
                def _pw(s):
                    if s is None:
                        return None
                    num, _u = split_unit(s)
                    return num
                unit_v = vals[0]
                prp = _pw(vals[0])
                eps = _pw(vals[1]) if len(vals) > 1 else None
                if "kva" in unit_v.lower().replace(" ", ""):
                    put("prp_kva", prp)
                    put("eps_kva", eps)
                else:
                    put("prp_kw", prp)
                    put("eps_kw", eps)
                continue

            # ---- 特殊：油耗多行 ----
            if key == "__fuelcons":
                if len(vals) < 2:
                    continue
                labels_raw = [x.strip().lower() for x in vals[0].split("\n") if x.strip()]
                nums = [x.strip() for x in vals[1].split("\n") if x.strip()]
                for lab, num in zip(labels_raw, nums):
                    n = clean(num)
                    if n is None:
                        continue
                    if "esp" in lab:
                        put("fc_esp", n)
                    elif "75" in lab:
                        put("fc_75", n)
                    elif "50" in lab:
                        put("fc_50", n)
                    elif "prp" in lab:
                        put("fc_prp", n)
                continue

            # ---- 特殊：尺寸/重量/油箱（OPEN / SILENT 两列）----
            if key in ("__lwh", "__weight", "__tank"):
                f["__dim_pending"] = {"__lwh": "lwh", "__weight": "wt",
                                      "__tank": "tank"}[key]
                if not vals:
                    continue
                o = clean(vals[0]) if len(vals) > 0 else None
                s = clean(vals[1]) if len(vals) > 1 else None
                pre = {"__lwh": "lwh", "__weight": "wt", "__tank": "tank"}[key]
                if o:
                    put("o_" + pre, o)
                if s:
                    put("s_" + pre, s)
                f.pop("__dim_pending", None)
                continue

            # ---- 续行合并：值为 "/55A" 这类碎片，拼回上次同表字段 ----
            if key is None and vals and re.fullmatch(r"[\/]\s*\d+(?:\.\d+)?\s*[A-Za-z]*", vals[0] or ""):
                prev = f.get("__last_key")
                if prev and f.get(prev):
                    f[prev] = str(f[prev]) + vals[0].strip()
                    continue

            if not key:
                if label and len(label) > 2:
                    notes.append("未识别标签: %s" % label[:40])
                continue
            if key == "__ignore":
                continue
            if not vals:
                continue

            raw = vals[0]

            # 字段后处理
            if key in ("ex_f", "ex_bp", "fu_res", "fu_max", "co_air", "in_res", "in_flow"):
                put(key, raw, with_unit=True)
            elif key == "el_v":
                # 前端 collect() 格式：DC24V（用户 2026-10-20 定：单位不能丢）
                mm = re.search(r"(\d+(?:\.\d+)?)\s*V", raw, re.I)
                put(key, "DC%sV" % mm.group(1) if mm else raw)
            elif key == "el_alt":
                # 前端 collect() 格式：DC24V/55A；源里常见 "DC24V" 与 "/55A" 分两格
                mm = re.search(r"DC\s*(\d+(?:\.\d+)?)\s*V\s*/?\s*(\d+(?:\.\d+)?)\s*A", raw, re.I)
                if mm:
                    put(key, "DC%sV/%sA" % (mm.group(1), mm.group(2)))
                else:
                    put(key, raw)
            elif key == "el_bat":
                # 前端 collect() 格式：`${个数}X${容量}Ah(${CCA}CCA)`
                # 用户指定：缺 CCA 写 "2x60ah（TBA CCA）"
                put(key, fmt_battery(raw))
            elif key in ("displ", "co_oil", "co_cool", "in_t", "ex_t"):
                num, _u = split_unit(raw)
                put(key, num)
            elif key == "rpm":
                num, _u = split_unit(raw)
                # 中文表里 rpm 行常混入"运行模式/Type of Operation"这类文本
                put(key, num if re.fullmatch(r"\d+(?:\.\d+)?", num or "") else None)
            elif key == "bore":
                # "135 × 165 mm" -> "135x165"
                s = raw.replace("×", "x").replace("*", "x")
                s = re.sub(r"\s*mm\s*$", "", s, flags=re.I)
                s = re.sub(r"\s*", "", s)
                put(key, s)
            elif key == "g_insul":
                # "Class H/H" -> "H"
                m = re.search(r"\b([HF])\b", raw)
                put(key, m.group(1) if m else raw)
            elif key in VALUE_MAP:
                vn = norm_label(raw)
                put(key, VALUE_MAP[key].get(vn, raw))
            else:
                put(key, raw)
            f["__last_key"] = key

    f.pop("__dim_pending", None)
    f.pop("__last_key", None)

    # ---- 中文表补充解析（2026-10-20）----
    # 中文参数单里电压常写成 "400/230V"、"440/254V" 这种独立一行，
    # 但表格被拆散后抓不到标签 -> 从行化文本里补。
    if _full_text:
        if not f.get("std_v"):
            m = re.search(r"(\d{3,4}\s*/\s*\d{2,3})\s*V\b", _full_text)
            if m:
                f["std_v"] = re.sub(r"\s+", "", m.group(1)) + "V"
        if not f.get("g_insul"):
            m = re.search(r"绝缘\s*等级[^\n]{0,20}?\b([HF])\b", _full_text)
            if m:
                f["g_insul"] = m.group(1)
    return f, notes


def build_model(item, name=None, use_pdf_note=True):
    """把一个 extracted JSON 转成机型对象"""
    src = item.get("source_file") or ""
    model_guess = (item.get("model_guess") or "").strip() or "未识别"
    fields, notes = parse_item(item)

    freq_raw = fields.get("freq")
    freq_key, freq_src = infer_freq(src, freq_raw, fields)
    if freq_src == "default":
        notes.append("频率无法判定，默认 50Hz")

    freqs = {freq_key: {}}
    common = {}
    for k, v in fields.items():
        if k in FREQ_FIELDS:
            freqs[freq_key][k] = v
        else:
            common[k] = v
    # freq 字段统一紧凑格式（"50Hz / 3P / PF0.8"）
    freq_norm = normalize_freq_str(freq_raw)
    if not freq_norm:
        ff = dict(fields)
        ff.setdefault("model", model_guess)   # 型号不在 fields 里，补进去供单相/三相判断
        freq_norm = synth_freq_str(freq_key, ff)
        if freq_norm:
            notes.append("freq 由频率/相数/PF 拼装：%s" % freq_norm)
    if freq_norm:
        freqs[freq_key]["freq"] = freq_norm

    data = {
        "model": model_guess,
        "curFreq": freq_key,
        "freqs": freqs,
        "common": common,
    }
    return {
        "name": name or model_guess,
        "model": model_guess,
        "family": guess_family(model_guess, src),
        "src": src,
        "data": data,
        "_notes": notes,
    }


# ============================================================
# 区分标签：(400V开架)
# ============================================================
def box_kind(fields):
    """判断箱体：开架 open / 静音 silent / 都有"""
    has_o = any(fields.get(k) for k in ("o_lwh", "o_wt", "o_tank", "o_db1", "o_db7"))
    has_s = any(fields.get(k) for k in ("s_lwh", "s_wt", "s_tank", "s_db1", "s_db7"))
    if has_o and has_s:
        return "开架静音"
    if has_s:
        return "静音"
    if has_o:
        return "开架"
    return ""


def voltage_tag(fields):
    """从标准电压取主电压：'400/230V' -> '400V'"""
    v = fields.get("std_v") or ""
    m = re.search(r"(\d{2,5})", v)
    return (m.group(1) + "V") if m else ""


def make_label(fields):
    """生成区分标签，如 '400V开架'；都取不到返回空"""
    parts = [voltage_tag(fields), box_kind(fields)]
    tag = "".join(p for p in parts if p)
    return tag


# ============================================================
# 合并同型号多文件（PDF+PPTX 算一份；50/60Hz 合并成一个机型）
# ============================================================
def merge_items(raw_items):
    """
    分组键 = (纯型号, 区分标签)
    同组内：合并 freqs、合并 common（后者不覆盖前者已有值）
    """
    groups = {}
    diag = []
    for it in raw_items:
        src = it.get("source_file") or ""
        is_pdf = src.lower().endswith(".pdf")
        # 只信 PPTX；PDF 单独记，供兜底
        model_guess = (it.get("model_guess") or "").strip() or "未识别"
        fields, notes = parse_item(it)
        tag = make_label(fields)
        key = (model_guess, tag)
        g = groups.setdefault(key, {
            "model": model_guess, "tag": tag, "pptx": [], "pdf": [],
            "srcs": [], "notes": [],
        })
        g["srcs"].append(src)
        g["notes"].extend(notes)
        (g["pdf"] if is_pdf else g["pptx"]).append((it, fields))

    # ---- 同型号 tag 收敛（2026-10-20）----
    # PDF 版常缺箱体信息 -> tag 只有 "400V"，而 PPTX 版是 "400V开架静音"，
    # 会被当成两个机型。规则：同型号下，若 tag A 是 tag B 的前缀（A 信息更少，
    # 且 A 没有箱体词），把 A 的源并入 B。
    by_model = {}
    for (model, tag) in list(groups):
        by_model.setdefault(model, []).append(tag)
    merge_map = {}
    for model, tags in by_model.items():
        if len(tags) < 2:
            continue
        tags_sorted = sorted(tags, key=len, reverse=True)   # 信息最全的在前
        for i, rich in enumerate(tags_sorted):
            if rich in merge_map:
                continue
            for poor in tags_sorted[i + 1:]:
                if poor in merge_map:
                    continue
                # poor 必须是 rich 的前缀（如 "400V" 是 "400V开架静音" 的前缀）
                box_words = ("开架", "静音")
                if rich.startswith(poor) and not any(w in poor for w in box_words):
                    merge_map[(model, poor)] = (model, rich)

    for (model, poor), richkey in list(merge_map.items()):
        src_g, dst_g = groups.get((model, poor)), groups.get(richkey)
        if not src_g or not dst_g or src_g is dst_g:
            continue
        dst_g["pptx"] += src_g["pptx"]
        dst_g["pdf"] += src_g["pdf"]
        dst_g["srcs"] += src_g["srcs"]
        dst_g["notes"] += src_g["notes"]
        del groups[(model, poor)]

    models = []
    renames = []
    used_names = set()
    for (model, tag), g in groups.items():
        # 优先 PPTX；没有 PPTX 才用 PDF
        src_list = g["pptx"] if g["pptx"] else g["pdf"]
        if not src_list:
            continue
        # 【关键】PDF 单相页常只有电机段（无功率/发动机），会生成空壳机型
        # → 整组都没有实质数据就丢弃
        if not any(has_data(fields) for _it, fields in src_list):
            continue
        freqs = {}
        common = {}
        conflicts = []          # 同型号多文件、同字段不同值 -> 待人工确认
        for it, fields in src_list:
            src = it.get("source_file") or ""
            freq_raw = fields.get("freq")
            fk, fk_src = infer_freq(src, freq_raw, fields)
            if fk_src == "default":
                g["notes"].append("[%s] 频率无法判定，默认 50Hz" % src)
            fd = freqs.setdefault(fk, {})
            # freq 字段缺失时用「频率 + 相数 + PF」拼装兜底（用户 2026-10-20 认）
            if not normalize_freq_str(freq_raw):
                ff = dict(fields)
                ff.setdefault("model", model)
                syn = synth_freq_str(fk, ff)
                if syn:
                    fields = dict(fields)
                    fields["freq"] = syn
                    g["notes"].append("[%s] freq 由频率/相数/PF 拼装：%s" % (src, syn))
            for k, v in fields.items():
                if v in (None, ""):
                    continue
                if k in FREQ_FIELDS:
                    if k == "freq":
                        nv = normalize_freq_str(v)
                        if nv:
                            if fd.get("freq") and fd["freq"] != nv:
                                conflicts.append((k, fd["freq"], nv, src))
                            else:
                                fd["freq"] = nv
                    elif not fd.get(k):
                        fd[k] = v
                    elif str(fd[k]) != str(v):
                        # 同频率块内功率冲突：保留第一个，但记录下来
                        conflicts.append((k, fd[k], v, src))
                else:
                    if not common.get(k):
                        common[k] = v
                    elif str(common[k]) != str(v):
                        conflicts.append((k, common[k], v, src))
        if not freqs:
            continue
        base_name = model if not tag else "%s (%s)" % (model, tag)
        name = base_name
        n = 2
        while name in used_names:
            name = "%s #%d" % (base_name, n)
            n += 1
        if name != base_name:
            renames.append((base_name, name))
        used_names.add(name)

        cur_freq = "50Hz" if "50Hz" in freqs else list(freqs)[0]
        models.append({
            "name": name,
            "model": model,
            "family": guess_family(model, " ".join(g["srcs"])),
            "src": " ; ".join(g["srcs"][:3]),
            "data": {"model": model, "curFreq": cur_freq,
                     "freqs": freqs, "common": common},
            "_notes": g["notes"],
            "_srcs": g["srcs"],
            # 用户 2026-10-20 定：同型号多文件冲突 -> 标「待人工确认」，先不入库
            "_conflicts": conflicts,
            # 只有 PDF 没有 PPTX -> 数据残缺（PDF 常缺功率段/箱体），同样待确认
            "_pdf_only": (not g["pptx"]) and bool(g["pdf"]),
            "needs_review": bool(conflicts) or ((not g["pptx"]) and bool(g["pdf"])),
        })
    return models, renames


# ============================================================
# 校验：字段完整性 / 类型 / 交叉核对
# ============================================================
NUMERIC_KEYS = ("prp_kva", "eps_kva", "prp_kw", "eps_kw", "rpm",
                "fc_esp", "fc_prp", "fc_75", "fc_50",
                "displ", "ex_t", "co_oil", "co_cool", "in_t",
                "o_wt", "s_wt", "o_tank", "s_tank",
                "o_db1", "o_db7", "s_db1", "s_db7", "el_v")

ALL_KEYS = set()
for _g in ("prp_kva eps_kva prp_kw eps_kw freq std_v opt_v o_lwh o_wt o_tank s_lwh s_wt s_tank "
           "o_db1 o_db7 s_db1 s_db7 e_brand e_model e_fuel e_asp e_cool e_gov e_cyl "
           "fc_esp fc_prp fc_75 fc_50 displ cyl_arr bore rpm ex_t ex_f ex_bp "
           "fu_res fu_max co_air co_oil co_cool in_flt in_res in_flow in_t "
           "el_v el_bat el_alt g_insul g_exc g_lead g_conn g_pitch g_phase g_brg").split():
    ALL_KEYS.add(_g)


def verify(models, allowed_keys=None):
    """返回 (errors, warnings)"""
    errors, warnings = [], []
    allowed = allowed_keys or ALL_KEYS
    for m in models:
        nm = m["name"]
        d = m.get("data") or {}
        if not d.get("model"):
            errors.append("%s: 缺 model" % nm)
        freqs = d.get("freqs") or {}
        if not freqs:
            errors.append("%s: 无 freqs 数据" % nm)
        for fk, fd in freqs.items():
            if fk not in ("50Hz", "60Hz"):
                errors.append("%s: 非法频率键 %r" % (nm, fk))
            if not fd:
                warnings.append("%s[%s]: 频率块为空" % (nm, fk))
            for k, v in fd.items():
                if k not in allowed:
                    errors.append("%s[%s]: 未知字段 %s" % (nm, fk, k))
                if k in NUMERIC_KEYS and v not in (None, ""):
                    if not re.fullmatch(r"[-+]?\d+(?:\.\d+)?", str(v)):
                        warnings.append("%s[%s]: %s 非数值 %r" % (nm, fk, k, v))
            # 功率一致性：kW ≈ kVA × PF
            try:
                kv = float(fd.get("prp_kva") or 0)
                kw = float(fd.get("prp_kw") or 0)
                if kv > 0 and kw > 0:
                    pf = kw / kv
                    if not (0.7 <= pf <= 1.01):
                        warnings.append("%s[%s]: PF 异常 %.3f (kW/kVA)" % (nm, fk, pf))
            except Exception:
                pass
        for k, v in (d.get("common") or {}).items():
            if k not in allowed:
                errors.append("%s: 未知 common 字段 %s" % (nm, k))
        if not (d.get("common") or {}).get("e_model"):
            warnings.append("%s: 缺发动机型号" % nm)
        if m.get("needs_review"):
            for k, a, b, src in (m.get("_conflicts") or [])[:6]:
                warnings.append("%s: 字段冲突 %s = %r / %r (来自 %s)" % (nm, k, a, b, src))
    return errors, warnings


# ============================================================
# 交叉核对（用户 2026-10-20 加）
#   备用 ≥ 主用；主/备用 PF 一致；电压与频率组合合理
# ============================================================
def cross_check(models):
    """返回 suspicion 列表（都是「可疑」，不是硬错误）"""
    out = []
    for m in models:
        nm = m["name"]
        d = m.get("data") or {}
        for fk, fd in (d.get("freqs") or {}).items():
            def num(k):
                try:
                    return float(fd.get(k))
                except (TypeError, ValueError):
                    return None
            pk, ek = num("prp_kva"), num("eps_kva")
            pw, ew = num("prp_kw"), num("eps_kw")
            if pk and ek and ek < pk:
                out.append("%s[%s]: 备用 kVA(%.1f) < 主用(%.1f)，通常不合理" % (nm, fk, ek, pk))
            if pw and ew and ew < pw:
                out.append("%s[%s]: 备用 kW(%.1f) < 主用(%.1f)，通常不合理" % (nm, fk, ew, pw))
            pf_p = (pw / pk) if (pk and pw) else None
            pf_e = (ew / ek) if (ek and ew) else None
            if pf_p and pf_e and abs(pf_p - pf_e) > 0.02:
                out.append("%s[%s]: 主用 PF(%.3f) 与备用 PF(%.3f) 不一致" % (nm, fk, pf_p, pf_e))
            std_v = str((d.get("common") or {}).get("std_v") or "")
            mv = re.search(r"(\d{3,4})", std_v)
            hz = re.search(r"(50|60)", fk)
            if mv and hz:
                volt, h = int(mv.group(1)), int(hz.group(1))
                if h == 60 and volt in (380, 400, 415):
                    out.append("%s[%s]: 60Hz 却用 %.0fV 标准电压，可疑" % (nm, fk, volt))
                if h == 50 and volt >= 440:
                    out.append("%s[%s]: 50Hz 却用 %.0fV 标准电压，可疑" % (nm, fk, volt))
    return out


# ============================================================
# CLI
# ============================================================
def load_items(src_dir):
    files = []
    for root, _dirs, fs in os.walk(src_dir):
        for fn in fs:
            if fn.endswith(".json") and not fn.startswith("index"):
                files.append(os.path.join(root, fn))
    files.sort()
    out = []
    for p in files:
        try:
            out.append(json.load(open(p, encoding="utf-8")))
        except Exception as e:
            print("跳过坏文件 %s: %s" % (p, e), file=sys.stderr)
    return out


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--src", required=True)
    ap.add_argument("--out", default="models_import.json")
    ap.add_argument("--report", default="")
    ap.add_argument("--dry-run", action="store_true")
    ap.add_argument("--verify", action="store_true")
    args = ap.parse_args()

    raw = load_items(args.src)
    print("读入 %d 份提取 JSON" % len(raw))
    models, renames = merge_items(raw)
    n_rev = sum(1 for m in models if m.get("needs_review"))
    n_pdf = sum(1 for m in models if m.get("_pdf_only"))
    print("合并后 %d 个机型（待人工确认 %d 个，其中仅PDF无PPTX %d 个）"
          % (len(models), n_rev, n_pdf))

    errs, warns = verify(models)
    sus = cross_check(models)
    print("校验: %d 错误, %d 警告, %d 可疑" % (len(errs), len(warns), len(sus)))

    # 字段覆盖率统计
    cov = {}
    for m in models:
        keys = set()
        for fd in (m["data"].get("freqs") or {}).values():
            keys |= set(fd)
        keys |= set(m["data"].get("common") or {})
        for k in keys:
            cov[k] = cov.get(k, 0) + 1
    print("\n字段覆盖率（共 %d 机型）:" % len(models))
    for k in sorted(cov, key=lambda x: -cov[x]):
        print("  %-10s %4d  (%.0f%%)" % (k, cov[k], 100.0 * cov[k] / max(1, len(models))))
    missing = sorted(ALL_KEYS - set(cov))
    if missing:
        print("  从未出现:", ", ".join(missing))

    # 输出
    # 用户 2026-10-20 定：冲突机型「待人工确认」→ 不入库，单独列出
    ok_models = [m for m in models if not m.get("needs_review")]
    review_models = [m for m in models if m.get("needs_review")]
    payload = {"models": [{"name": m["name"], "model": m["model"],
                           "family": m["family"], "src": m["src"],
                           "data": m["data"]} for m in ok_models],
               "review": [{"name": m["name"], "model": m["model"],
                           "family": m["family"], "src": m["src"],
                           "data": m["data"],
                           "reason": ("仅PDF无PPTX" if m.get("_pdf_only")
                                      else "多文件字段冲突"),
                           "conflicts": [{"field": k, "kept": a, "other": b, "from": c}
                                         for k, a, b, c in (m.get("_conflicts") or [])]}
                          for m in review_models]}
    if not args.dry_run:
        with open(args.out, "w", encoding="utf-8") as f:
            json.dump(payload, f, ensure_ascii=False, indent=1)
        print("\n已写出 %s (入库 %d 机型, 待确认 %d 机型)"
              % (args.out, len(ok_models), len(review_models)))

    report = []
    report.append("合并后机型数: %d" % len(models))
    report.append("校验错误: %d" % len(errs))
    report.append("校验警告: %d" % len(warns))
    if renames:
        report.append("\n重名加序号 (%d):" % len(renames))
        for a, b in renames[:50]:
            report.append("  %s -> %s" % (a, b))
    if errs:
        report.append("\n[错误]")
        report += errs[:100]
    if warns:
        report.append("\n[警告]")
        report += warns[:100]
    if sus:
        report.append("\n[可疑·需核对] (%d)" % len(sus))
        report += sus[:100]
    if review_models:
        report.append("\n[待人工确认·不入库] (%d)" % len(review_models))
        for m in review_models[:50]:
            reason = "仅PDF无PPTX·数据残缺" if m.get("_pdf_only") else "多文件字段冲突"
            report.append("  %s  [%s]  <- %s" % (m["name"], reason, m["src"]))
            for k, a, b, c in (m.get("_conflicts") or [])[:6]:
                report.append("      %s: 保留 %r / 另有 %r (%s)" % (k, a, b, c))
    txt = "\n".join(report)
    if args.report:
        open(args.report, "w", encoding="utf-8").write(txt)
    print("\n" + txt[:3000])


if __name__ == "__main__":
    main()
