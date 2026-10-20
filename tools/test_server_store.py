#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""机型库服务器存储 —— 端到端验证（跑在临时库上，不碰生产 paramsheet.db）"""
import json
import os
import sys
import urllib.request
import urllib.error

BASE = os.environ.get("PS_BASE", "http://127.0.0.1:8099")
PASS, FAIL = [], []


def ok(m):
    PASS.append(m)
    print("  ✓", m)


def bad(m):
    FAIL.append(m)
    print("  ✗", m)


def req(method, path, body=None):
    url = BASE + path
    data = None
    headers = {}
    if body is not None:
        data = json.dumps(body, ensure_ascii=False).encode()
        headers["Content-Type"] = "application/json"
    r = urllib.request.Request(url, data=data, headers=headers, method=method)
    try:
        with urllib.request.urlopen(r, timeout=30) as resp:
            return resp.status, json.loads(resp.read().decode())
    except urllib.error.HTTPError as e:
        try:
            return e.code, json.loads(e.read().decode())
        except Exception:
            return e.code, {}


def mkitem(name, model, kva50, kva60=None, family="AS"):
    freqs = {"50Hz": {"prp_kva": kva50, "prp_kw": str(int(float(kva50) * 0.8)),
                      "eps_kva": str(int(float(kva50) * 1.1)),
                      "freq": "50Hz / 3P / PF0.8"}}
    if kva60:
        freqs["60Hz"] = {"prp_kva": kva60, "prp_kw": str(int(float(kva60) * 0.8)),
                         "eps_kva": str(int(float(kva60) * 1.1)),
                         "freq": "60Hz / 3P / PF0.8"}
    return {"name": name, "model": model, "family": family, "src": "test",
            "data": {"model": model, "curFreq": "50Hz", "freqs": freqs,
                     "common": {"model": model, "std_v": "400/230V",
                                "e_brand": "Cummins", "e_model": "QSB7-G5"}}}


def main():
    print("=" * 60)
    print("V1 基础连通 + 空库")
    print("=" * 60)
    st, j = req("GET", "/health")
    (ok if st == 200 and j.get("ok") else bad)("GET /health 200 ok")
    (ok if j.get("models") == 0 else bad)("空库 models=0，实际 %r" % j.get("models"))

    st, j = req("GET", "/api/models-all")
    (ok if st == 200 and j.get("total") == 0 else bad)("空库 /api/models-all total=0")
    (ok if isinstance(j.get("models"), list) else bad)("models-all 返回 list")

    st, j = req("GET", "/api/models")
    (ok if st == 200 and j.get("total") == 0 else bad)("空库 /api/models total=0")

    print()
    print("=" * 60)
    print("V2 单条保存 / 读取 / 双频")
    print("=" * 60)
    it = mkitem("CU330D5 (400V开架)", "CU330D5", "330", "375")
    st, j = req("POST", "/api/models/save", {
        "name": it["name"], "model": it["model"], "family": "AS",
        "src": "test", "overwrite": True, "data": it["data"]})
    (ok if st == 200 and j.get("ok") and j.get("total") == 1 else bad)(
        "保存 1 条 total=1，得到 %r %r" % (st, j))

    st, j = req("GET", "/api/models-all")
    ms = j.get("models") or []
    (ok if len(ms) == 1 else bad)("models-all 1 条，实际 %d" % len(ms))
    if ms:
        m = ms[0]
        (ok if m["name"] == it["name"] else bad)("name 一致 %r" % m["name"])
        (ok if m["model"] == "CU330D5" else bad)("model 一致 %r" % m["model"])
        (ok if sorted(m["freqs"].keys()) == ["50Hz", "60Hz"] else bad)(
            "双频都在 %r" % list(m["freqs"].keys()))
        (ok if m["freqs"]["50Hz"]["prp_kva"] == "330" else bad)(
            "50Hz prp_kva=330，实际 %r" % m["freqs"]["50Hz"].get("prp_kva"))
        (ok if m["freqs"]["60Hz"]["prp_kva"] == "375" else bad)(
            "60Hz prp_kva=375，实际 %r" % m["freqs"]["60Hz"].get("prp_kva"))
        (ok if m["common"]["e_brand"] == "Cummins" else bad)("common 完整")
        (ok if m["curFreq"] == "50Hz" else bad)("curFreq=50Hz，实际 %r" % m["curFreq"])

    st, j = req("GET", "/api/models/" + urllib.parse.quote(it["name"]))
    (ok if st == 200 and j.get("item", {}).get("name") == it["name"] else bad)(
        "单条读取 200")

    print()
    print("=" * 60)
    print("V3 路由不冲突（models-all 不被 {name} 抢）")
    print("=" * 60)
    st, j = req("GET", "/api/models-all")
    (ok if st == 200 and "models" in j else bad)(
        "/api/models-all 没被 /api/models/{name} 抢走（%r）" % st)
    st, j = req("GET", "/api/models/all")
    (ok if st == 404 else bad)("/api/models/all 应为 404，实际 %r" % st)

    print()
    print("=" * 60)
    print("V4 批量导入 / 覆盖 / 去重")
    print("=" * 60)
    items = [mkitem("AS1250E5A (400V开架静音)", "AS1250E5A", "1250", "1500"),
             mkitem("AS1375E5 (400V开架静音)", "AS1375E5", "1375", "1650"),
             mkitem("AF40D6-1P (240V开架静音)", "AF40D6-1P", "40", None, "AF")]
    st, j = req("POST", "/api/models/import", {"overwrite": True, "items": items})
    (ok if st == 200 and j.get("added") == 3 else bad)(
        "批量导入 added=3，实际 %r" % j)
    (ok if j.get("total") == 4 else bad)("total=4，实际 %r" % j.get("total"))

    # 再导一次，全部走 update
    st, j = req("POST", "/api/models/import", {"overwrite": True, "items": items})
    (ok if j.get("updated") == 3 and j.get("added") == 0 else bad)(
        "重复导入走 update，实际 %r" % j)
    (ok if j.get("total") == 4 else bad)("重复导入后 total 仍 4，实际 %r" % j.get("total"))

    # overwrite=False 应跳过
    st, j = req("POST", "/api/models/import", {"overwrite": False, "items": items})
    (ok if j.get("skipped") == 3 else bad)("overwrite=False 全部 skipped，实际 %r" % j)

    print()
    print("=" * 60)
    print("V5 单频机型 + 频率列表")
    print("=" * 60)
    st, j = req("GET", "/api/models-all")
    d = {m["name"]: m for m in j["models"]}
    one = d.get("AF40D6-1P (240V开架静音)")
    (ok if one and list(one["freqs"].keys()) == ["50Hz"] else bad)(
        "单频机型只有 50Hz，实际 %r" % (list(one["freqs"].keys()) if one else None))
    (ok if one and one["family"] == "AF" else bad)("family 保留 AF")
    two = d.get("AS1250E5A (400V开架静音)")
    (ok if two and two["freq_list"] == ["50Hz", "60Hz"] else bad)(
        "freq_list 双频，实际 %r" % (two["freq_list"] if two else None))

    st, j = req("GET", "/api/models")
    (ok if j.get("total") == 4 else bad)("列表 total=4，实际 %r" % j.get("total"))
    fams = [x["family"] if isinstance(x, dict) else x for x in (j.get("families") or [])]
    (ok if sorted(fams) == ["AF", "AS"] else bad)(
        "families=[AF,AS]，实际 %r" % fams)

    st, j = req("GET", "/api/models?family=AF")
    (ok if j.get("total") == 1 else bad)("family=AF 过滤 1 条，实际 %r" % j.get("total"))

    st, j = req("GET", "/api/models?q=1250")
    (ok if j.get("total") == 1 else bad)("q=1250 搜索 1 条，实际 %r" % j.get("total"))

    print()
    print("=" * 60)
    print("V6 更新后数据确实变了（不是只改计数）")
    print("=" * 60)
    it2 = mkitem("AS1250E5A (400V开架静音)", "AS1250E5A", "1300", "1560")
    req("POST", "/api/models/save", {"name": it2["name"], "model": "AS1250E5A",
                                     "family": "AS", "overwrite": True,
                                     "data": it2["data"]})
    st, j = req("GET", "/api/models-all")
    d = {m["name"]: m for m in j["models"]}
    (ok if d["AS1250E5A (400V开架静音)"]["freqs"]["50Hz"]["prp_kva"] == "1300" else bad)(
        "覆盖后 50Hz prp_kva=1300，实际 %r"
        % d["AS1250E5A (400V开架静音)"]["freqs"]["50Hz"].get("prp_kva"))

    print()
    print("=" * 60)
    print("V7 删除 / 不存在")
    print("=" * 60)
    st, j = req("DELETE", "/api/models/" + urllib.parse.quote("AS1375E5 (400V开架静音)"))
    (ok if st == 200 and j.get("ok") and j.get("total") == 3 else bad)(
        "删除后 total=3，实际 %r" % j)
    st, j = req("GET", "/api/models/" + urllib.parse.quote("不存在X"))
    (ok if st == 404 else bad)("不存在机型 404，实际 %r" % st)
    st, j = req("POST", "/api/models/save", {"name": "", "data": {}})
    (ok if st == 400 else bad)("缺 name 400，实际 %r" % st)
    st, j = req("POST", "/api/models/import", {"items": []})
    (ok if st == 400 else bad)("空 items 400，实际 %r" % st)

    print()
    print("=" * 60)
    print("V8 导出格式 + 与前端 DB 结构兼容")
    print("=" * 60)
    st, j = req("GET", "/api/models-export")
    (ok if st == 200 and len(j.get("models") or []) == 3 else bad)(
        "导出 3 条，实际 %r" % len(j.get("models") or []))
    sample = (j.get("models") or [{}])[0]
    need = {"model", "curFreq", "freqs", "common"}
    (ok if need.issubset(set(sample.keys())) else bad)(
        "导出项含前端所需键 %r，实际 %r" % (sorted(need), sorted(sample.keys())))

    st, j = req("GET", "/api/models-all")
    a = (j.get("models") or [{}])[0]
    need2 = {"name", "model", "family", "freqs", "common", "curFreq"}
    (ok if need2.issubset(set(a.keys())) else bad)(
        "models-all 项含前端所需键 %r" % sorted(need2))

    print()
    print("=" * 60)
    print("总计: 通过 %d / 失败 %d" % (len(PASS), len(FAIL)))
    print("=" * 60)
    return 1 if FAIL else 0


if __name__ == "__main__":
    import urllib.parse  # noqa
    sys.exit(main())
