# -*- coding: utf-8 -*-
"""发动机动力库加载（替代原 paramsheet.db 的机组机型列表）"""
import json, os

HERE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "data")

def _load(fn):
    p = os.path.join(HERE, fn)
    if not os.path.exists(p):
        return {}
    with open(p, encoding="utf-8") as f:
        return json.load(f)

def engines():
    """合并 50Hz + 60Hz 动力库，转成前端可直接用的机型结构"""
    out = []
    for fn in ("engine_50hz.json", "engine_60hz.json"):
        d = _load(fn)
        for e in d.get("engines", []):
            vals = dict(e.get("values") or {})
            vals["rpm"] = e.get("rpm", "")
            out.append({
                "name": e.get("model_name"),
                "model": e.get("engine"),
                "family": e.get("family") or "",
                "freq": e.get("freq"),
                "rpm": e.get("rpm"),
                "variant": e.get("variant") or "",
                "src": (e.get("src_models") or [""])[0],
                "needs_fill": bool(e.get("needs_fill")),
                "filled_fields": e.get("filled_fields", 0),
                "values": vals,
            })
    out.sort(key=lambda x: (x["family"], x["name"] or ""))
    return out

def families():
    from collections import Counter
    c = Counter((e["family"] or "其他") for e in engines())
    return dict(c.most_common())
