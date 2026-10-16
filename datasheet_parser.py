#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
★ Datasheet 解析器（动力/发动机参数自动提取）

设计原则：
- 规则驱动（正则 + 标签邻近匹配），不是「关键词 if in text」式硬编码
- 覆盖常见厂家样本（中/英，PDF/Excel）
- 能读的填，读不到的留空 —— 由前端表单补填（C 混合）
- 单位换算按 memory/oem-template.md 记录的口径

入口：parse_datasheet(filename, content_bytes) -> {field: value}
"""
import io
import re

# ============================================================
# 文本抽取层
# ============================================================

def extract_text(filename, content):
    """从 PDF / XLSX / XLS / TXT 抽取纯文本（保布局）。"""
    name = (filename or "").lower()
    if name.endswith(".pdf"):
        return _pdf_text(content)
    if name.endswith((".xlsx", ".xlsm")):
        return _xlsx_text(content)
    if name.endswith(".xls"):
        return _xls_text(content)
    # txt / 其它
    for enc in ("utf-8", "gbk", "latin-1"):
        try:
            return content.decode(enc)
        except Exception:
            continue
    return ""


def _pdf_text(content):
    # pdfplumber 优先（布局好），失败退回 pypdf
    try:
        import pdfplumber
        out = []
        with pdfplumber.open(io.BytesIO(content)) as pdf:
            for pg in pdf.pages[:12]:
                out.append(pg.extract_text() or "")
        return "\n".join(out)
    except Exception:
        pass
    try:
        from pypdf import PdfReader
        r = PdfReader(io.BytesIO(content))
        return "\n".join((p.extract_text() or "") for p in r.pages[:12])
    except Exception:
        return ""


def _xlsx_text(content):
    try:
        import openpyxl
        wb = openpyxl.load_workbook(io.BytesIO(content), read_only=True, data_only=True)
        out = []
        for ws in wb.worksheets[:3]:
            out.append("### SHEET: %s" % ws.title)
            for row in ws.iter_rows(values_only=True):
                cells = ["" if c is None else str(c) for c in row]
                if any(cells):
                    out.append(" | ".join(cells))
        return "\n".join(out)
    except Exception:
        return ""


def _xls_text(content):
    try:
        import xlrd
        wb = xlrd.open_workbook(file_contents=content)
        out = []
        for ws in wb.sheets()[:3]:
            out.append("### SHEET: %s" % ws.name)
            for r in range(min(ws.nrows, 400)):
                cells = [str(ws.cell_value(r, c)) for c in range(ws.ncols)]
                if any(cells):
                    out.append(" | ".join(cells))
        return "\n".join(out)
    except Exception:
        return ""


# ============================================================
# 字段提取规则
# ============================================================

def _num(s):
    """从字符串取第一个数字（含小数/负号）。逗号作千分位去掉。"""
    if s is None:
        return None
    s = str(s).replace(",", "")
    m = re.search(r"(-?\d+(?:\.\d+)?)", s)
    return m.group(1) if m else None


def _grab(text, patterns, group=1):
    """按顺序试多个正则，返回第一个命中的捕获组（去空白）。"""
    for p in patterns:
        m = re.search(p, text, re.I)
        if m:
            v = m.group(group)
            if v is not None and v.strip() != "":
                return v.strip()
    return ""


def _unit_line(text, label_pattern, unit_patterns, window=90):
    """标签 + 单位 邻近匹配：在标签附近找带单位的数值。"""
    for lp in label_pattern:
        for m in re.finditer(lp, text, re.I):
            seg = text[m.end():m.end() + window]
            for up in unit_patterns:
                um = re.search(up, seg, re.I)
                if um:
                    return um.group(0).strip()
    return ""


def parse_datasheet(filename, content):
    """主入口：返回待填字段 dict（键同前端表单字段名）。"""
    text = extract_text(filename, content)
    if not text:
        return {}
    return parse_text(text)


def parse_text(text, model_hint=""):
    """从纯文本提取字段。"""
    t = text
    t1 = re.sub(r"[ \t]+", " ", t)          # 压缩空格，保留换行
    flat = re.sub(r"\s+", " ", t)           # 全压平，给邻近匹配用
    f = {}

    # ---- 型号 / 品牌 ----
    f["e_model"] = _grab(t, [
        r"Engine Model[^\n:：]*[:：]\s*([A-Za-z0-9][A-Za-z0-9\-\.]{2,})",
        r"发动机型号[^\n:：]*[:：]\s*([A-Za-z0-9][A-Za-z0-9\-\.]{2,})",
        r"^([A-Z0-9]{2,}[A-Z0-9\-\.]{2,})\s*$",  # 首行大字型号
    ])
    if model_hint and not f["e_model"]:
        f["e_model"] = model_hint
    f["e_brand"] = _grab(t, [
        r"Brand[^\n:：]*[:：]\s*([^\n]{2,30})",
        r"品牌[^\n:：]*[:：]\s*([^\n]{2,30})",
    ])
    # 常见厂家名作为兜底（用词边界，避免 MAN 命中 manual）
    if not f["e_brand"]:
        for b in ("Cummins", "Perkins", "DEUTZ", "Volvo", "Yanmar", "KUBOTA",
                  "Weichai", "潍柴", "上柴", "玉柴", "MTU", "Baudouin", "Mitsubishi",
                  "Doosan", "Isuzu", "FPT", "Scania", "MAN", "Lister",
                  "Kirloskar", "Yuchai", "Shangchai"):
            if re.search(r"\b" + re.escape(b) + r"\b", t):
                f["e_brand"] = b
                break
            if re.search(re.escape(b), t) and not re.match(r"[A-Za-z]", b):
                f["e_brand"] = b
                break

    # ---- 排量 Displacement ----
    v = _unit_line(flat, [r"Displacement", r"排量"], [r"\d+(?:\.\d+)?\s*[Ll]\b"])
    if v:
        f["displ"] = v.replace(" ", "")
    else:
        m = re.search(r"(?:Displacement|排量)[^\d]{0,20}(\d+(?:\.\d+)?)\s*L", flat, re.I)
        if m:
            f["displ"] = m.group(1) + "L"

    # ---- 缸数 / 排列 ----
    m = re.search(r"(?:Number of cylinders|No\.?\s*of\s*cyl|缸数)[^\d]{0,20}(\d{1,2})\b", flat, re.I)
    cyl_n = ""
    if m:
        n = int(m.group(1))
        if 1 <= n <= 24:
            cyl_n = str(n)
    arr = ""
    if re.search(r"\bV[- ]?type\b|V型|Vee\b", flat, re.I):
        arr = "V"
    elif re.search(r"In-?line|直列", flat, re.I):
        arr = "L"
    if cyl_n or arr:
        f["cyl_arr"] = (cyl_n + arr) if (cyl_n and arr) else (cyl_n or arr)
    m2 = re.search(r"Cylinder (?:arrangement|type)[^\n:：]*[:：]\s*([^\n]{2,30})", t, re.I)
    if m2 and not arr and not cyl_n:
        f["cyl_arr"] = m2.group(1).strip()

    # ---- 缸径×行程 ----
    m = re.search(r"(?:Bore\s*[×x*]\s*stroke|缸径\s*[×x*]\s*行程)[^\d]{0,20}(\d+(?:\.\d+)?)\s*[×x*]\s*(\d+(?:\.\d+)?)", flat, re.I)
    if m:
        f["bore"] = "%sx%smm" % (m.group(1), m.group(2))

    # ---- 压缩比 ----
    m = re.search(r"(?:Compression\s*ratio|压缩比)[^\d]{0,20}(\d+(?:\.\d+)?)\s*[:：]\s*1", flat, re.I)
    if m:
        f["cr"] = m.group(1) + ":1"

    # ---- 转速 ----
    rpms = re.findall(r"\b(1500|1800|1200|3000|3600)\s*r/?min\b", flat, re.I)
    if not rpms:
        rpms = re.findall(r"\b(1500|1800|1200|3000|3600)\b", flat)
    if rpms:
        seen = []
        for r_ in rpms:
            if r_ not in seen:
                seen.append(r_)
        f["rpm"] = "/".join(seen[:2]) + "rpm"

    # ---- 发动机类型 ----
    if re.search(r"4-?stroke|四冲程|4\s*cycle", flat, re.I):
        f["etype"] = "4-stroke diesel"
    m = re.search(r"(?:Engine Type|发动机类型)[^\n:：]*[:：]\s*([^\n]{3,60})", t, re.I)
    if m:
        f["etype"] = m.group(1).strip()

    # ---- 进气方式 ----
    asp = []
    if re.search(r"turbo|增压|Turbocharg", flat, re.I):
        asp.append("Turbocharged")
    if re.search(r"aftercool|中冷|intercool|charge air cool", flat, re.I):
        asp.append("Intercooled")
    if not asp and re.search(r"naturally aspirated|自然吸气|natural", flat, re.I):
        asp.append("Naturally Aspirated")
    if asp:
        f["e_asp"] = " & ".join(asp)
        if len(asp) == 2:
            f["e_asp"] = "Turbocharged & Intercooled"

    # ---- 冷却方式 ----
    if re.search(r"water-?cooled|水冷", flat, re.I):
        f["e_cool"] = "Water-cooled"
    elif re.search(r"air-?cooled|风冷", flat, re.I):
        f["e_cool"] = "Air-cooled"

    # ---- 调速器 ----
    if re.search(r"electronic (?:governor|control)|电子调速|电控|ECU|electric control", flat, re.I):
        f["e_gov"] = "Electronic Governor"
    elif re.search(r"mechanical governor|机械调速|全程式机械", flat, re.I):
        f["e_gov"] = "Mechanical Governor"

    # ---- 燃油 ----
    if re.search(r"diesel|柴油", flat, re.I):
        f["e_fuel"] = "Diesel"

    # ---- 应用等级 ----
    apps = []
    if re.search(r"prime|常用", flat, re.I):
        apps.append("Prime")
    if re.search(r"standby|备用|emergency|应急", flat, re.I):
        apps.append("Standby")
    if apps:
        f["e_app"] = " & ".join(apps)

    # ---- 油耗 4 档（L/h）----
    fc = _fuel_consumption(flat, t)
    f.update(fc)

    # ---- 排气系统 ----
    m = re.search(r"(?:Exhaust\s*(?:gas\s*)?temp|排气温度)[^\d]{0,25}(\d{2,4})\s*°?\s*C", flat, re.I)
    if m:
        f["ex_t"] = m.group(1) + "°C"
    m = re.search(r"(?:Exhaust\s*(?:gas\s*)?flow|排气流量)[^\d]{0,25}(\d+(?:\.\d+)?)\s*(m3?/min|m³/min|L/[Ss]|m3?/h)", flat, re.I)
    if m and "fuel" not in flat[max(0, m.start()-60):m.start()].lower():
        f["ex_f"] = m.group(1) + m.group(2)
    m = re.search(r"(?:Exhaust\s*back\s*pressure|排气背压)[^\d]{0,25}(\d+(?:\.\d+)?)\s*(kPa|mbar|mmH2O)", flat, re.I)
    if m:
        val, unit = m.group(1), m.group(2)
        if unit.lower() == "mbar":
            val = str(round(float(val) / 10, 2))
            unit = "kPa"
        f["ex_bp"] = val + unit

    # ---- 燃油系统 ----
    m = re.search(r"(?:Fuel\s*(?:feed|transfer|pump).{0,20}?resist|输油泵阻力)[^\d]{0,25}(\d+(?:\.\d+)?)\s*(kPa|mbar)", flat, re.I)
    if m:
        val, unit = m.group(1), m.group(2)
        if unit.lower() == "mbar":
            val = str(round(float(val) / 10, 2)); unit = "kPa"
        f["fu_res"] = val + unit
    m = re.search(r"(?:Max(?:imum)?\s*fuel\s*(?:flow|consumption)|最大燃油流量)[^\d]{0,25}(\d+(?:\.\d+)?)\s*(L/[Hh]|L/h|kg/h)", flat, re.I)
    if m:
        f["fu_max"] = m.group(1) + "L/H"

    # ---- 冷却系统 ----
    m = re.search(r"(?:Radiator\s*(?:air\s*)?flow|散热器空气流量)[^\d]{0,25}(\d+(?:\.\d+)?)\s*(m3?/min|m³/min|L/[Ss]|m3?/h)", flat, re.I)
    if m:
        f["co_air"] = m.group(1) + m.group(2)
    m = re.search(r"(?:Lube\s*oil\s*(?:capacity|volume)|机油(?:总)?容量|Oil\s*pan)[^\d]{0,25}(\d+(?:\.\d+)?)\s*[Ll]", flat, re.I)
    if m:
        f["co_oil"] = m.group(1) + "L"
    m = re.search(r"(?:Coolant\s*(?:capacity|volume)|冷却液(?:总)?容量)[^\d]{0,25}(\d+(?:\.\d+)?)\s*[Ll]", flat, re.I)
    if m:
        f["co_cool"] = m.group(1) + "L"

    # ---- 电气系统 ----
    m = re.search(r"(?:Starter|起动机)[^\d]{0,25}(DC\s*24\s*V|24\s*V|DC\s*12\s*V|12\s*V)", flat, re.I)
    if m:
        f["el_v"] = "DC24V" if "24" in m.group(1) else "DC12V"
    m = re.search(r"(?:Batter|电池)[^\d]{0,25}(\d+\s*[x×]\s*\d+\s*Ah[^\n,\)]*)", flat, re.I)
    if m:
        f["el_bat"] = m.group(1).strip()
    m = re.search(r"(?:Alternator|充电发电机|Charging)[^\d]{0,25}(DC)?\s*(24|12)\s*V\s*/?\s*(\d+)\s*A", flat, re.I)
    if m:
        f["el_alt"] = "DC%sV/%sA" % (m.group(2), m.group(3))

    # ---- 进气系统 ----
    if re.search(r"dry\s*type|干式|纸芯|paper element", flat, re.I):
        f["in_flt"] = "Dry type"
    m = re.search(r"(?:Max(?:imum)?\s*(?:air\s*)?(?:intake|inlet)\s*restrict|最大进气阻力)[^\d]{0,25}(\d+(?:\.\d+)?)\s*(kPa|mbar)", flat, re.I)
    if m:
        val, unit = m.group(1), m.group(2)
        if unit.lower() == "mbar":
            val = str(round(float(val) / 10, 2)); unit = "kPa"
        f["in_res"] = val + unit
    m = re.search(r"(?:Combustion\s*air\s*flow|燃烧空气流量)[^\d]{0,25}(\d+(?:\.\d+)?)\s*(L/[Ss]|m3?/min|m³/min)", flat, re.I)
    if m:
        f["in_flow"] = m.group(1) + m.group(2)
    m = re.search(r"(?:Max(?:imum)?\s*(?:intake|inlet)\s*(?:air\s*)?temp|最高进气温度)[^\d]{0,25}(\d{2,3})\s*°?\s*C", flat, re.I)
    if m:
        f["in_t"] = m.group(1) + "°C"

    # ---- 频率 / 相数 / PF（若有）----
    freq = "50Hz" if re.search(r"\b50\s*Hz\b", flat, re.I) else ("60Hz" if re.search(r"\b60\s*Hz\b", flat, re.I) else "")
    ph = "3P" if re.search(r"three[- ]?phase|3[- ]?phase|三相", flat, re.I) else ""
    if freq or ph:
        parts = [x for x in (freq, ph, "PF0.8") if x]
        f["freq"] = " / ".join(parts)

    # 丢掉空值
    return {k: v for k, v in f.items() if v not in (None, "")}


def _fuel_consumption(flat, t):
    """油耗 4 档：100% ESP / 100% PRP / 75% PRP / 50% PRP（L/H）。"""
    out = {}
    # 常见：100%/75%/50% 标签后跟数值
    pats = [
        (r"100\s*%\s*(?:ESP|Standby|备用)[^\d]{0,20}(\d+(?:\.\d+)?)", "fc_esp"),
        (r"100\s*%\s*(?:PRP|Prime|主用|常用)[^\d]{0,20}(\d+(?:\.\d+)?)", "fc_prp"),
        (r"75\s*%[^\d]{0,20}(\d+(?:\.\d+)?)", "fc_75"),
        (r"50\s*%[^\d]{0,20}(\d+(?:\.\d+)?)", "fc_50"),
    ]
    for p, key in pats:
        m = re.search(p, flat, re.I)
        if m:
            out[key] = m.group(1)
    # 110%/100% 形式（如 4HTAA 表：25/50/75/100/110%）—— 取 100% 作 ESP
    if not out.get("fc_esp"):
        m = re.search(r"110\s*%[^\d]{0,20}(\d+(?:\.\d+)?)", flat, re.I)
        if m:
            out["fc_esp"] = m.group(1)
    return out
