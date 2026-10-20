#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""前端浏览器端到端验证：真实 Chrome 打开页面，检查 服务器存储 / 回退 / 增删改"""
import json
import sys
import time
import urllib.request
import urllib.error

from playwright.sync_api import sync_playwright

_BASE = os.environ.get("PS_URL", "http://127.0.0.1:8099/")
URL = _BASE.rstrip("/") + ("/paramsheet/" if os.environ.get("PS_PREFIX") else "/")
API = _BASE.rstrip("/")
PASS, FAIL = [], []


def ok(m):
    PASS.append(m)
    print("  ✓", m)


def bad(m):
    FAIL.append(m)
    print("  ✗", m)


def api(method, path, body=None):
    data = json.dumps(body, ensure_ascii=False).encode() if body is not None else None
    h = {"Content-Type": "application/json"} if body is not None else {}
    r = urllib.request.Request(API + path, data=data, headers=h, method=method)
    with urllib.request.urlopen(r, timeout=20) as resp:
        return json.loads(resp.read().decode())


def reset_db(items):
    """把临时库恢复成指定数据"""
    cur = [m["name"] for m in api("GET", "/api/models-all")["models"]]
    for n in cur:
        api("DELETE", "/api/models/" + urllib.parse.quote(n))
    if items:
        api("POST", "/api/models/import", {"overwrite": True, "items": items})


def mkitem(name, model, kva50, kva60=None, family="AS"):
    freqs = {"50Hz": {"prp_kva": kva50, "prp_kw": str(int(float(kva50) * 0.8)),
                      "eps_kva": str(int(float(kva50) * 1.1)),
                      "freq": "50Hz / 3P / PF0.8"}}
    if kva60:
        freqs["60Hz"] = {"prp_kva": kva60, "prp_kw": str(int(float(kva60) * 0.8)),
                         "eps_kva": str(int(float(kva60) * 1.1)),
                         "freq": "60Hz / 3P / PF0.8"}
    return {"name": name, "model": model, "family": family, "src": "browser-test",
            "data": {"model": model, "curFreq": "50Hz", "freqs": freqs,
                     "common": {"model": model, "std_v": "400/230V",
                                "e_brand": "Cummins", "e_model": "QSB7-G5",
                                "prp_kva": "", "eps_kva": ""}}}


def main():
    import urllib.parse  # noqa

    with sync_playwright() as p:
        exe = ("/Users/bob/Library/Caches/ms-playwright/chromium-1234/"
               "chrome-mac-arm64/Google Chrome for Testing.app/Contents/MacOS/"
               "Google Chrome for Testing")
        browser = p.chromium.launch(executable_path=exe, headless=True,
                                    args=["--no-sandbox"])
        ctx = browser.new_context()
        page = ctx.new_page()
        errors = []
        page.on("pageerror", lambda e: errors.append(str(e)))
        page.on("console", lambda m: errors.append("console:" + m.text)
                if m.type == "error" else None)

        print("=" * 60)
        print("B1 服务器模式：页面从服务器加载机型")
        print("=" * 60)
        reset_db([mkitem("CU330D5 (400V开架)", "CU330D5", "330", "375"),
                  mkitem("AF40D6-1P (240V开架静音)", "AF40D6-1P", "40", None, "AF")])
        # 预置一个「旧本地数据」，验证会被丢弃
        page.goto(URL)
        page.evaluate("""() => {
            localStorage.setItem('paramsheet_db_v2',
                JSON.stringify({models:[{model:'OLD_LOCAL_JUNK', curFreq:'50Hz',
                    freqs:{'50Hz':{}}, common:{model:'OLD_LOCAL_JUNK'}}]}));
        }""")
        page.reload()
        page.wait_for_timeout(2500)

        status = page.inner_text("#status")
        (ok if "数据存服务器" in status else bad)(
            "状态栏显示「数据存服务器」，实际 %r" % status)

        opts = page.eval_on_selector_all("#selModel option", "els=>els.map(e=>e.textContent)")
        (ok if set(opts) == {"CU330D5 (400V开架)", "AF40D6-1P (240V开架静音)"}
         else bad)("机型下拉来自服务器 %r" % opts)
        (ok if "OLD_LOCAL_JUNK" not in opts else bad)(
            "旧的 localStorage 机型已被丢弃 %r" % opts)

        ls = page.evaluate("() => localStorage.getItem('paramsheet_db_v2')")
        (ok if ls is None else bad)("服务器模式下 localStorage 旧键已清空，实际 %r" % ls)

        # 表单回填应跟「下拉当前选中的那台」一致（排序后才决定默认选谁）
        sel_v = page.eval_on_selector("#selModel", "e=>e.value")
        v = page.eval_on_selector('[data-k="prp_kva"]', "e=>e.value")
        exp = "330" if sel_v.startswith("CU330D5") else "40"
        (ok if v == exp else bad)(
            "表单回填与选中机型一致（%s -> %s），实际 %r" % (sel_v, exp, v))

        print()
        print("=" * 60)
        print("B2 切换机型 -> 表单跟着换")
        print("=" * 60)
        page.select_option("#selModel", "AF40D6-1P (240V开架静音)")
        page.wait_for_timeout(900)
        v2 = page.eval_on_selector('[data-k="prp_kva"]', "e=>e.value")
        (ok if v2 == "40" else bad)("切到 AF40D6 后 prp_kva=40，实际 %r" % v2)

        print()
        print("=" * 60)
        print("B3 改字段 -> 自动推送到服务器")
        print("=" * 60)
        page.select_option("#selModel", "CU330D5 (400V开架)")
        page.wait_for_timeout(900)
        before = page.eval_on_selector('[data-k="prp_kva"]', "e=>e.value")
        (ok if before == "330" else bad)("切到 CU330D5 后读到的原值 330，实际 %r" % before)
        page.fill('[data-k="prp_kva"]', "999")
        page.wait_for_timeout(2600)     # 防抖 800ms + 网络
        j = api("GET", "/api/models-all")
        d = {m["name"]: m for m in j["models"]}
        got = d["CU330D5 (400V开架)"]["freqs"]["50Hz"].get("prp_kva")
        (ok if got == "999" else bad)("服务器已存 prp_kva=999，实际 %r" % got)

        # 还原，避免污染后面断言
        page.fill('[data-k="prp_kva"]', "330")
        page.wait_for_timeout(2200)

        print()
        print("=" * 60)
        print("B4 新建机型 -> 服务器出现")
        print("=" * 60)
        page.evaluate("() => { window.prompt = () => 'TEST_NEW_1'; }")
        page.click("text=+ 新建")
        page.wait_for_timeout(2000)
        j = api("GET", "/api/models-all")
        names = [m["name"] for m in j["models"]]
        (ok if "TEST_NEW_1" in names else bad)("新建机型已写进服务器 %r" % names)

        print()
        print("=" * 60)
        print("B5 删除机型 -> 服务器同步删除")
        print("=" * 60)
        page.select_option("#selModel", "TEST_NEW_1")
        page.wait_for_timeout(500)
        page.evaluate("() => { window.confirm = () => true; }")
        page.click("text=删除")
        page.wait_for_timeout(2000)
        j = api("GET", "/api/models-all")
        names = [m["name"] for m in j["models"]]
        (ok if "TEST_NEW_1" not in names else bad)("删除后服务器已无该机型 %r" % names)
        (ok if len(names) == 2 else bad)("服务器仍剩 2 个机型，实际 %d" % len(names))

        print()
        print("=" * 60)
        print("B6 刷新页面 -> 数据还在（持久化）")
        print("=" * 60)
        page.reload()
        page.wait_for_timeout(2500)
        opts2 = page.eval_on_selector_all("#selModel option", "els=>els.map(e=>e.textContent)")
        (ok if set(opts2) == {"CU330D5 (400V开架)", "AF40D6-1P (240V开架静音)"}
         else bad)("刷新后机型仍从服务器来 %r" % opts2)

        print()
        print("=" * 60)
        print("B7 服务器挂掉 -> 回退本地并提示")
        print("=" * 60)
        # 用路由拦截模拟后端不可用
        page2 = ctx.new_page()
        page2.route("**/api/models-all", lambda route: route.abort())
        page2.route("**/health", lambda route: route.abort())
        page2.goto(URL)
        page2.wait_for_timeout(2500)
        st2 = page2.inner_text("#status")
        (ok if "未连接" in st2 else bad)("断线时状态栏提示「未连接」，实际 %r" % st2)
        (ok if "暂存本地" in st2 else bad)("断线时提示「暂存本地」，实际 %r" % st2)

        print()
        print("=" * 60)
        print("B8 页面无 JS 报错")
        print("=" * 60)
        real = [e for e in errors if "ERR_FAILED" not in e and "net::" not in e
                and "Failed to load resource" not in e]
        (ok if not real else bad)("无 JS 报错，实际: %r" % real[:4])

        browser.close()

    print()
    print("=" * 60)
    print("总计: 通过 %d / 失败 %d" % (len(PASS), len(FAIL)))
    print("=" * 60)
    return 1 if FAIL else 0


if __name__ == "__main__":
    sys.exit(main())
