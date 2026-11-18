# -*- coding: utf-8 -*-
"""参数单管理工具 — 后端接口（第2步，FastAPI 版）"""
import os, sys, json, sqlite3, re, subprocess, threading
import requests
from fastapi import FastAPI, Query, Request, Response
from fastapi.responses import JSONResponse, FileResponse, HTMLResponse
from fastapi.staticfiles import StaticFiles
from pydantic import BaseModel
from typing import Optional

BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, BASE)
STD_DB = os.path.join(BASE, 'paramsheet.db')
NON_DB = os.path.join(BASE, 'paramsheet_nonstd.db')
IMG_DIRS = [os.path.join(BASE, 'images'), os.path.join(BASE, 'images_alt')]
OUT_DIR = os.path.join(BASE, 'output')

# LibreOffice 可执行文件（macOS 本地 / Linux 服务器自动探测）
def _find_soffice():
    for p in ('/Applications/LibreOffice.app/Contents/MacOS/soffice',
              '/usr/bin/soffice', '/usr/bin/libreoffice',
              '/usr/lib64/libreoffice/program/soffice'):
        if os.path.exists(p):
            return p
    return 'soffice'


_SOFFICE = _find_soffice()

# ---------------- Mac 端渲染服务（Tailscale 内网） ----------------
# 服务器只有 1.8GB 内存，跑 LibreOffice 会被 OOM 杀掉（实测 rss 1.5GB）。
# 渲染挪到 16GB 的 Mac（100.125.54.14:8021），服务器只做转发 + 取回产物。
MAC_RENDER = os.environ.get('MAC_RENDER', 'http://100.125.54.14:8021')
MAC_ENABLED = os.environ.get('MAC_RENDER_DISABLE', '') != '1'


def _mac_health(timeout=10):
    """Mac 渲染服务是否可用。连不上就返回 False，走本地兜底。"""
    if not MAC_ENABLED:
        return False
    try:
        r = requests.get(MAC_RENDER + '/health', timeout=timeout)
        return r.ok and (r.json() or {}).get('ok') is True
    except Exception:
        return False


def _mac_render(f, lang='zh'):
    """调 Mac 渲染，返回 (pages_basenames, pdf_basename, errors)。"""
    r = requests.post(MAC_RENDER + '/render',
                      json={'data': f, 'lang': lang, 'pages': 4, 'dpi': 110},
                      timeout=180)
    j = r.json()
    if not (r.ok and j.get('ok')):
        raise RuntimeError(j.get('error') or ('mac render http %s' % r.status_code))
    return j.get('pages') or [], j.get('pdf'), None


def _mac_fetch(fn, dest_dir):
    """从 Mac 拉取渲染产物到服务器本地。"""
    r = requests.get(MAC_RENDER + '/file/' + fn, timeout=120, stream=True)
    if not r.ok:
        raise RuntimeError('fetch %s failed: %s' % (fn, r.status_code))
    dest = os.path.join(dest_dir, fn)
    with open(dest, 'wb') as fp:
        for chunk in r.iter_content(65536):
            if chunk:
                fp.write(chunk)
    return dest
ADMIN_USER, ADMIN_PASS = 'bob', '12345'

app = FastAPI(title='参数单管理工具')


@app.get('/api/download/{fn}')
def api_download(fn: str, as_: str = Query('', alias='as')):
    """直接下载文件（浏览器另存为，不跳打印界面）
    2026-11-17：文件名按「品牌+备用功率kVA+D/E+5/6」规则（as 参数），如 CU1000D5.pdf"""
    import urllib.parse as _up
    safe = os.path.basename(fn)
    p = os.path.join(OUT_DIR, safe)
    if not os.path.isfile(p) or not os.path.abspath(p).startswith(os.path.abspath(OUT_DIR)):
        return JSONResponse({'ok': False, 'error': '文件不存在'}, 404)
    ext = safe.rsplit('.', 1)[-1].lower()
    mt = {'pdf': 'application/pdf',
          'pptx': 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
          'png': 'image/png'}.get(ext, 'application/octet-stream')
    disp = ''
    if as_:
        # 只收字母数字下的点，防非法文件名
        a = re.sub(r'[^0-9A-Za-z._\-]', '', str(as_))[:80]
        if a:
            disp = a if a.lower().endswith('.' + ext) else '%s.%s' % (a, ext)
    if not disp:
        # 兜底：去掉指纹后缀，保留机型名
        disp = safe
        m = re.match(r'^(.*)_[0-9a-f]{6}(_paramsheet_[a-z]{2})?(\.[a-z]+)$', safe)
        if m:
            disp = m.group(1) + (m.group(2) or '') + m.group(3)
    quoted = _up.quote(disp)
    return FileResponse(p, media_type=mt, filename=disp,
                        headers={'Content-Disposition':
                                 "attachment; filename=\"%s\"; filename*=UTF-8''%s" % (safe, quoted)})


import logging
import traceback as _tb
import datetime as _dt

LOG_PATH = os.path.join(BASE, 'logs', 'api.log')
os.makedirs(os.path.dirname(LOG_PATH), exist_ok=True)
_log = logging.getLogger('paramsheet')
if not _log.handlers:
    _log.setLevel(logging.INFO)
    _fh = logging.FileHandler(LOG_PATH, encoding='utf-8')
    _fh.setFormatter(logging.Formatter('%(asctime)s %(levelname)s %(message)s'))
    _log.addHandler(_fh)
    _sh = logging.StreamHandler()
    _sh.setFormatter(logging.Formatter('%(asctime)s %(levelname)s %(message)s'))
    _log.addHandler(_sh)


@app.middleware('http')
async def _access_log(request: Request, call_next):
    """统一访问日志 + 异常兜底：任何接口崩了都返回 JSON 而不是 500 白屏"""
    import time as _t
    t0 = _t.time()
    try:
        resp = await call_next(request)
    except Exception as e:
        _log.error('EXC %s %s -> %s\n%s', request.method, request.url.path, e, _tb.format_exc()[-1500:])
        return JSONResponse({'ok': False, 'error': '服务器内部错误：' + str(e)}, 500)
    ms = int((_t.time() - t0) * 1000)
    if request.url.path.startswith('/api') or ms > 500:
        _log.info('%s %s %s %dms', request.method, request.url.path, resp.status_code, ms)
    resp.headers['X-Elapsed-Ms'] = str(ms)
    return resp


def _now():
    return _dt.datetime.now().strftime('%Y-%m-%d %H:%M:%S')


def conn(kind='std'):
    c = sqlite3.connect(STD_DB if kind == 'std' else NON_DB)
    c.row_factory = sqlite3.Row
    return c


def dicts(rows):
    return [dict(r) for r in rows]


# ---------------- 摊平层 ----------------
FREQ_KEYS = ['prp_kva', 'prp_kw', 'eps_kva', 'eps_kw', 'freq', 'rpm', 'fc_esp', 'fc_prp', 'fc_75', 'fc_50']
COMMON_KEYS = ['std_v', 'opt_v', 'e_brand', 'e_model', 'e_fuel', 'e_asp', 'e_cool', 'e_gov', 'e_cyl',
               'o_lwh', 'o_wt', 'o_db1', 'o_db7', 'o_tank', 's_lwh', 's_wt', 's_db1', 's_db7', 's_tank',
               'displ', 'cyl_arr', 'bore', 'ex_t', 'ex_f', 'ex_bp', 'fu_res', 'fu_max',
               'co_air', 'co_oil', 'co_cool', 'el_v', 'el_bat', 'el_alt',
               'in_flt', 'in_res', 'in_flow', 'in_t',
               'g_insul', 'g_exc', 'g_lead', 'g_conn', 'g_pitch', 'g_phase', 'g_brg']


def flatten(row):
    o = json.loads(row['data'])
    cm = o.get('common') or {}
    freqs = o.get('freqs') or {}
    out = {'name': row['name'], 'model': o.get('model'), 'src': row['src']}
    for k in ('shell_code', 'shell_name', 'family_code', 'family_name', 'zone', 'version'):
        try:
            out[k] = row[k]
        except Exception:
            out[k] = None
    cur_f = o.get('curFreq')
    out['cur_freq'] = cur_f
    for fk, label in (('50Hz', '50'), ('60Hz', '60')):
        fv = freqs.get(fk) or {}
        for k in FREQ_KEYS:
            out['%s_%s' % (label, k)] = fv.get(k)
    main = freqs.get(cur_f) or {}
    for k in FREQ_KEYS:
        out[k] = main.get(k)
    for k in COMMON_KEYS:
        out[k] = cm.get(k)
    out['fc_lines'] = [main.get('fc_esp'), main.get('fc_prp'), main.get('fc_75'), main.get('fc_50')]
    # 模板第1页「频率/相数/功率因数」格用 freq_phase_pf，必须由 freq 字段派生
    out['freq_phase_pf'] = main.get('freq') or cur_f
    return out


def parse_name(name):
    m = re.match(r'^([^\s\(]+)\s*\(([^)]*)\)\s*(?:\[([^\]]+)\])?', name or '')
    if not m:
        return {'model_code': name, 'voltage': None, 'variant': None, 'engine_model': None}
    vol = re.match(r'^(\d+(?:\.\d+)?V)(.*)$', m.group(2))
    return {'model_code': m.group(1),
            'voltage': vol.group(1) if vol else None,
            'variant': (vol.group(2) if vol else m.group(2)).replace('款', '').strip() or None,
            'engine_model': m.group(3)}


# --------- 图片 ---------
_img_inv = None


def img_inventory():
    global _img_inv
    if _img_inv is None:
        p = os.path.join(BASE, 'images_inventory.json')
        _img_inv = json.load(open(p, encoding='utf-8')) if os.path.exists(p) else {}
        _img_inv = {k.upper(): v for k, v in _img_inv.items()}
    return _img_inv


def images_for(name):
    code = parse_name(name)['model_code'].upper()
    e = img_inventory().get(code) or {}
    out = {}
    for kind in ('OPEN', 'SILENT'):
        f = (e.get(kind) or {}).get('file')
        out[kind] = '/images/' + f if f else None
    return out


@app.get('/images/{fn}')
def serve_img(fn: str):
    for d in IMG_DIRS:
        p = os.path.join(d, fn)
        if os.path.exists(p):
            return FileResponse(p)
    return JSONResponse({'error': 'not found'}, 404)


# ---------------- 字典 ----------------
@app.get('/api/dict')
def api_dict(kind: str = 'std'):
    c = conn(kind)
    shell = dicts(c.execute("select * from dict_shell order by sort_order").fetchall())
    family = dicts(c.execute("select * from dict_family order by sort_order").fetchall())
    volt = dicts(c.execute("select * from dict_voltage order by phase,freq,voltage").fetchall())
    fc = {r['family_code']: r['n'] for r in c.execute("select family_code, count(*) n from models group by family_code")}
    sc = {r['shell_code']: r['n'] for r in c.execute("select shell_code, count(*) n from models group by shell_code")}
    for f in family:
        f['count'] = fc.get(f['code'], 0)
    for s in shell:
        s['count'] = sc.get(s['code'], 0)
    total = c.execute("select count(*) from models").fetchone()[0]
    pend = c.execute("select count(*) from models where zone='pending'").fetchone()[0]
    zones = {r['zone'] or 'std': r['n'] for r in
             c.execute("select zone, count(*) n from models group by zone")}
    vols = [r['v'] for r in c.execute(
        "select json_extract(data,'$.common.std_v') v from models "
        "where json_extract(data,'$.common.std_v') is not null "
        "group by v order by cast(replace(v,'/','') as integer)")]
    c.close()
    return {'shell': shell, 'family': family, 'voltage': volt, 'volt_values': vols,
            'total': total, 'pending': pend, 'zones': zones, 'kind': kind}


# ---------------- 列表 ----------------
# 电压字段分段归一：把 "10000/5773V" 变成 "|10000|5773|"，
# 这样 "like '%|10000|%'" 只命中真正的 10000V 段，不会被 6600 之类子串误伤
_VOLT_SEG_FORM_LIKE = """('|' || replace(replace(replace(replace(
     lower(json_extract(m.data,'$.common.std_v')),
     'v',''),' ',''),'/','|'),'.','') || '|') like ?"""


def _build_search_where(q: str):
    """搜索词 -> SQL 条件，返回 (cond 列表, 参数列表)。

    单词内部是「或」，多个词之间是「且」。
    单词规则：
      ① 文字直匹配 名称/型号/发动机/发动机品牌/款式/电压原文/频率原文
      ② 电压段精确匹配 10000 / 10000V / 10000v -> std_v "10000/5773V" 的 10000 段
      ③ 带 k 写法 10kV / 6.6kV -> 换算伏特再走 ②
      ④ 功率数字 仅【纯数字】输入时启用（搜 1375）
    """
    words = [w for w in re.split(r'[\s,，]+', q.strip()) if w]
    if not words:
        return [], []
    groups, all_params = [], []
    for w in words:
        conds, params = _build_word_conds(w)
        groups.append('(' + ' or '.join(conds) + ')')
        all_params.extend(params)
    joined = (' and '.join(groups)) if len(groups) > 1 else groups[0]
    return ['(' + joined + ')'], all_params


def _build_word_conds(w: str):
    """单个词的匹配条件 + 参数"""
    like = '%' + w + '%'
    conds, params = [], []

    # ① 文字直匹配
    conds.append("""(m.name like ? or m.model like ? or m.family_name like ?
         or m.family_code like ? or m.shell_name like ? or m.shell_code like ?
         or json_extract(m.data,'$.common.e_model') like ?
         or json_extract(m.data,'$.common.e_brand') like ?
         or json_extract(m.data,'$.common.std_v') like ?
         or json_extract(m.data,'$.curFreq') like ?)""")
    params += [like] * 10

    # ② 电压段精确匹配（避免 6600/3810V 被搜 "60" 误命中）
    qn = re.sub(r'[^0-9.]', '', w).rstrip('.')
    if qn and re.fullmatch(r'\d+(?:\.\d+)?', qn):
        conds.append(_VOLT_SEG_FORM_LIKE)
        params.append('%|' + qn + '|%')

    # ③ 10kV / 6.6kV 写法
    mk = re.fullmatch(r'\s*(\d+(?:\.\d+)?)\s*k\s*v?\s*', w, re.I)
    if mk:
        volts = str(int(round(float(mk.group(1)) * 1000)))
        conds.append(_VOLT_SEG_FORM_LIKE)
        params.append('%|' + volts + '|%')

    # ④ 功率：仅纯数字输入
    if re.fullmatch(r'\d+', w.strip()):
        conds.append("""(cast(json_extract(m.data,'$.freqs."50Hz".prp_kva') as text) like ?
             or cast(json_extract(m.data,'$.freqs."50Hz".eps_kva') as text) like ?
             or cast(json_extract(m.data,'$.freqs."60Hz".prp_kva') as text) like ?
             or cast(json_extract(m.data,'$.freqs."60Hz".eps_kva') as text) like ?)""")
        params += ['%' + w.strip() + '%'] * 4

    return conds, params



@app.get('/api/list')
def api_list(kind: str = 'std', q: str = '', family: str = '', shell: str = '',
             zone: str = '', volt: str = '', page: int = 1, size: int = 20):
    page = max(1, page)
    size = min(100, max(5, size))
    c = conn(kind)
    where, params = [], []
    if family:
        where.append("m.family_code = ?"); params.append(family)
    if shell:
        where.append("m.shell_code = ?"); params.append(shell)
    if zone:
        where.append("m.zone = ?"); params.append(zone)
    if volt:
        where.append("json_extract(m.data,'$.common.std_v') = ?"); params.append(volt)
    if q.strip():
        # ---- 智能模糊搜索（2026-11-18 用户定）----
        # 一个搜索框搞定：型号 / 发动机 / 品牌 / 款式 / 电压 / 频率 / 功率
        # 「10000v」「10000V」「10kV」「10000/5773V」命中同一条；
        # 空格分隔多词 = 且（"50Hz 10000V"、"CU 11000V"）
        _conds, _params = _build_search_where(q)
        where.extend(_conds)
        params.extend(_params)
    wsql = (' where ' + ' and '.join(where)) if where else ''
    total = c.execute("select count(*) from models m" + wsql, params).fetchone()[0]
    rows = c.execute("""select m.name, m.shell_code, m.shell_name, m.family_code, m.family_name,
                               m.zone, m.version,
                               json_extract(m.data,'$.common.e_model') e_model,
                               json_extract(m.data,'$.common.e_brand') e_brand,
                               json_extract(m.data,'$.common.std_v') std_v,
                               json_extract(m.data,'$.curFreq') cur_freq,
                               json_extract(m.data,'$.freqs."50Hz".prp_kva') p50,
                               json_extract(m.data,'$.freqs."60Hz".prp_kva') p60
                        from models m""" + wsql + " order by m.name limit ? offset ?",
                        params + [size, (page - 1) * size]).fetchall()
    items = []
    for r in rows:
        d = dict(r)
        d['images'] = images_for(d['name'])
        d['parsed'] = parse_name(d['name'])
        items.append(d)
    c.close()
    return {'total': total, 'page': page, 'size': size,
            'pages': (total + size - 1) // size, 'items': items}


# ---------------- 详情 ----------------
@app.get('/api/detail')
def api_detail(kind: str = 'std', name: str = ''):
    c = conn(kind)
    r = c.execute("select * from models where name=?", (name,)).fetchone()
    c.close()
    if not r:
        return JSONResponse({'ok': False, 'error': 'not found'}, 404)
    f = flatten(r)
    f['parsed'] = parse_name(name)
    f['images'] = images_for(name)
    f['has_open'] = bool(f.get('o_lwh') or f.get('o_wt') or f.get('o_tank') or f['images'].get('OPEN'))
    f['has_silent'] = bool(f.get('s_lwh') or f.get('s_wt') or f.get('s_tank') or f['images'].get('SILENT'))
    diff = compare_with_std(f) if kind == 'nonstd' else None
    return {'ok': True, 'data': f, 'diff': diff}


# ---------------- 非标对比 ----------------
@app.get('/api/std_index')
def api_std_index(q: str = '', limit: int = 20):
    """标准单索引：非标派生时选基准（按型号/发动机/品牌搜）"""
    limit = min(50, max(5, limit))
    c = conn('nonstd')
    if q.strip():
        like = '%' + q.strip() + '%'
        rows = c.execute("""select * from std_index
            where name like ? or model_code like ? or e_model like ? or e_brand like ?
            order by model_code limit ?""", (like, like, like, like, limit)).fetchall()
    else:
        rows = c.execute("select * from std_index order by model_code limit ?", (limit,)).fetchall()
    # 补 prp_kva / eps_kva：前端选基准时要显示功率（std_index 表若无该列，从标准库回查）
    out = []
    sc = None
    for r in rows:
        d = dict(r)
        if d.get('prp_kva') in (None, ''):
            if sc is None:
                sc = conn('std')
            rr = sc.execute('select json_extract(data,\'$.freqs."50Hz".prp_kva\') a, '
                            'json_extract(data,\'$.freqs."50Hz".eps_kva\') b '
                            'from models where name=?', (d.get('name'),)).fetchone()
            if rr:
                d['prp_kva'] = rr['a']
                d['eps_kva'] = rr['b']
        out.append(d)
    if sc is not None:
        sc.close()
    rows = out
    total = c.execute("select count(*) from std_index").fetchone()[0]
    c.close()
    return {'ok': True, 'total': total, 'items': dicts(rows)}


class DeriveIn(BaseModel):
    base: str = ''            # 基准标准单主键
    name: str = ''            # 新非标单主键（留空则自动派生）
    data: dict = {}           # 覆盖字段（没给的沿用基准值）


@app.post('/api/derive')
def api_derive(body: DeriveIn, request: Request):
    """由标准单派生非标单：复制基准数据 → 覆盖差异字段 → 写进非标库"""
    if not _is_admin(request):
        return JSONResponse({'ok': False, 'error': '需要管理员登录'}, 403)
    base = (body.base or '').strip()
    if not base:
        return JSONResponse({'ok': False, 'error': '缺少基准标准单'}, 400)
    sc = conn('std')
    r = sc.execute("select * from models where name=?", (base,)).fetchone()
    if not r:
        sc.close()
        return JSONResponse({'ok': False, 'error': '基准标准单不存在：' + base}, 404)
    raw = json.loads(r['data'])
    base_shell, base_family, base_version = r['shell_code'], r['family_code'], r['version']
    base_name = r['name']
    sc.close()

    # 覆盖字段写进嵌套结构
    d = dict(body.data or {})
    cur_f = raw.get('curFreq') or '50Hz'
    fv = raw.setdefault('freqs', {}).setdefault(cur_f, {})
    cm = raw.setdefault('common', {})
    for k in FREQ_KEYS:
        if k in d and d[k] not in (None, ''):
            fv[k] = d[k]
    for k in COMMON_KEYS:
        if k in d and d[k] not in (None, ''):
            cm[k] = d[k]
    if d.get('model'):
        raw['model'] = d['model']

    name = (body.name or '').strip() or (base_name + ' · 非标')
    if name == base_name:
        return JSONResponse({'ok': False, 'error': '非标单不能和标准单同名'}, 400)

    nc = sqlite3.connect(NON_DB)
    nc.row_factory = sqlite3.Row
    if nc.execute("select 1 from models where name=?", (name,)).fetchone():
        nc.close()
        return JSONResponse({'ok': False, 'error': '非标单已存在：' + name}, 409)
    now = _now()
    nc.execute("""insert into models(name,model,family,freq_list,src,data,created_at,updated_at,
                 shell_code,shell_name,family_code,family_name,zone,version,display_name)
                 values(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)""",
               (name, raw.get('model'), d.get('family') or '', json.dumps(list((raw.get('freqs') or {}).keys())),
                '由标准单派生：' + base_name, json.dumps(raw, ensure_ascii=False), now, now,
                base_shell, base_shell, base_family, d.get('family_name'), 'nonstd',
                'V' + _dt.datetime.now().strftime('%Y%m%d') + '-01', None))
    try:
        nc.execute("""insert into models_fts(name,model,shell_name,family_name,e_model,e_brand,std_v,zone)
                      values(?,?,?,?,?,?,?,?)""",
                   (name, raw.get('model'), base_shell, d.get('family_name') or '',
                    cm.get('e_model'), cm.get('e_brand'), cm.get('std_v'), 'nonstd'))
    except Exception:
        pass
    nc.commit()
    nc.close()
    _log.info('derive %s <- %s', name, base_name)
    return {'ok': True, 'name': name, 'base': base_name}


def compare_with_std(nonstd):
    # 基准查找顺序：① 非标名自带「型号 (电压款式) [发动机]」→ 用型号代码
    #              ② 否派生的非标名不规范（如手工命名）→ 退回用模型型号 model 字段
    pn = parse_name(nonstd['name'])
    base = (pn['model_code'] or '').upper()
    if not re.match(r'^[A-Z0-9][A-Z0-9/\-\.]*$', base or ''):
        base = ''
    if (not base) or (not re.search(r'[\(（]', nonstd['name'] or '')):
        base = (nonstd.get('model') or base or '').upper()
    c = conn('std')
    cand = c.execute("select name from models where upper(model)=upper(?) or name like ?",
                     (base, base + ' (%')).fetchall()
    if not cand:
        c.close()
        return {'base': None, 'diffs': [], 'note': '找不到基准标准单 → 按规则不允许存非标'}
    pick = None
    for r in cand:
        sf = flatten(c.execute("select * from models where name=?", (r['name'],)).fetchone())
        if (sf.get('e_model') or '').strip() == (nonstd.get('e_model') or '').strip():
            pick = sf
            break
    if pick is None:
        pick = flatten(c.execute("select * from models where name=?", (cand[0]['name'],)).fetchone())
    c.close()
    diffs = []
    for k in COMMON_KEYS + ['prp_kva', 'eps_kva', 'prp_kw', 'eps_kw', 'rpm']:
        a, b = pick.get(k), nonstd.get(k)
        if (a or '') != (b or '') and (a or b):
            diffs.append({'field': k, 'std': a, 'nonstd': b})
    return {'base': pick['name'], 'diffs': diffs, 'note': ''}


@app.get('/api/compare')
def api_compare(name: str = ''):
    c = conn('nonstd')
    r = c.execute("select * from models where name=?", (name,)).fetchone()
    c.close()
    if not r:
        return JSONResponse({'ok': False, 'error': 'not found'}, 404)
    return {'ok': True, 'diff': compare_with_std(flatten(r))}


# ---------------- 登录 ----------------
class LoginIn(BaseModel):
    user: str = ''
    password: str = ''


@app.post('/api/login')
def api_login(body: LoginIn, response: Response):
    if body.user == ADMIN_USER and body.password == ADMIN_PASS:
        response.set_cookie('admin', '1', httponly=True, samesite='lax')
        return {'ok': True, 'admin': True}
    return JSONResponse({'ok': False, 'error': '账号或密码不对'}, 401)


@app.post('/api/logout')
def api_logout(response: Response):
    response.delete_cookie('admin')
    return {'ok': True}


@app.get('/api/whoami')
def api_whoami(request: Request):
    admin = request.cookies.get('admin') == '1'
    return {'admin': admin, 'user': ADMIN_USER if admin else None}


# ---------------- 导出 PDF ----------------
def _tpl_for(lang):
    """按语言取模板路径（zh / en）。未知语言回退中文。"""
    lg = 'en' if str(lang or '').lower().startswith('en') else 'zh'
    p = os.path.join(BASE, 'templates', 'template_%s.pptx' % lg)
    if not os.path.exists(p):
        lg = 'zh'
        p = os.path.join(BASE, 'templates', 'template_zh.pptx')
    return lg, p


# --------- 参数单文件名（用户 2026-11-17 定） ---------
# 格式：品牌缩写 + 备用功率kVA + D/E + 5/6
#   D = 机械/电子调速，E = ECU/ECM
#   5 = 50Hz，6 = 60Hz
# 例：CU1000D5 / CU1375E5 / AGG500E6
def _bcode(brand, model_code):
    """品牌缩写。e_brand 全称没有既有缩写时，用型号自身的字母前缀。"""
    b = str(brand or '').strip().upper().split()[0] if str(brand or '').strip() else ''
    MAP = {
        'CUMMINS': 'CU', 'PERKINS': 'PER', 'BAUDOUIN': 'BAU', 'DEUTZ': 'DE',
        'WEICHAI': 'WP', 'HYUNDAI': 'HY', 'MTU': 'MTU', 'YUCHAI': 'YC',
        'SCANIA': 'SC', 'SDEC': 'SD', 'VOLVO': 'VO', 'YANMAR': 'YM',
        'ISUZU': 'IS', 'KUBOTA': 'KB', 'YANGDONG': 'YD', 'FPT': 'FPT',
        'MITSUBISHI': 'MI', 'SME': 'SME', 'AGG': 'AGG',
    }
    if b in MAP:
        return MAP[b]
    mt = re.match(r'^([A-Za-z]+)', str(model_code or ''))
    return (mt.group(1).upper() if mt else b) or 'GEN'


def _gov_code(gov):
    """调速类型 → D（机械/电子调速）/ E（ECU/ECM）。"""
    g = str(gov or '').strip().lower()
    if not g:
        return 'D'
    if ('ecu' in g) or ('ecm' in g) or ('hpcr' in g) or ('management' in g):
        return 'E'
    if ('mechanical' in g) or ('mechanica' in g):
        return 'D'
    return 'D'


def _freq_code(freq, model_code):
    """频率 → 5（50Hz）/ 6（60Hz）。频率缺失时借型号末位。"""
    f = str(freq or '')
    if '60' in f:
        return '6'
    if '50' in f:
        return '5'
    mt = re.match(r'^[A-Za-z]+[\d.]+[DE]([56])$', str(model_code or ''))
    return mt.group(1) if mt else '5'


def doc_filename(f, lang='zh', ext='pdf'):
    """生成对外下载的参数单文件名，如 CU1000D5.pdf。

    命名标准（用户 2026-11-17 定）：
      品牌缩写 + 备用功率kVA + D/E + 5/6
      D=机械/电子调速，E=ECU/ECM，5=50Hz，6=60Hz
    型号本身已符合该标准时**直接沿用型号**（型号是给客户看的，
    库里 eps_kva/e_gov 的差异属历史笔误，不拿它们覆盖型号）。
    """
    model_code = str(parse_name(f.get('name') or '').get('model_code')
                     or f.get('model') or '').strip()
    # 型号本身就带「D/E + 5/6」或「G + 5/6」→ 直接沿用（型号是给客户看的）
    # 覆盖：CU1000D5 / CU1375E5 / SC2500D5C / P11D5/1P / W625G5 / CU27D5ATL
    #       AC440L-QSZ13-G3 / 4RT55-110D 等带引擎前缀的
    if re.search(r'(?:[DEG][56])(?![\d])', model_code):
        return '%s.%s' % (re.sub(r'[/\\:*?"<>|]', '-', model_code), ext)
    # AS/G 系列机身型号（如 AS25000-G31）本身就是成品型号，无 D/E/5-6 位也直接沿用
    if re.match(r'^AS\d+-G\d+[A-Z]?$', model_code):
        return '%s.%s' % (model_code, ext)
    # 否则按规则拼
    brand = f.get('e_brand') or (f.get('parsed') or {}).get('e_brand')
    try:
        pw = int(round(float(f.get('eps_kva'))))
    except Exception:
        mt = re.search(r'(\d+)', model_code)
        pw = int(mt.group(1)) if mt else 0
    base = '%s%d%s%s' % (_bcode(brand, model_code), pw,
                         _gov_code(f.get('e_gov')),
                         _freq_code(f.get('cur_freq') or f.get('curFreq'), model_code))
    base = re.sub(r'[\\/:*?"<>|]', '-', base)
    return '%s.%s' % (base, ext)


class ExportIn(BaseModel):
    kind: str = 'std'
    name: str = ''
    shell: str = ''
    lang: str = 'zh'


@app.post('/api/export')
def api_export(body: ExportIn):
    c = conn(body.kind)
    r = c.execute("select * from models where name=?", (body.name,)).fetchone()
    c.close()
    if not r:
        return JSONResponse({'ok': False, 'error': 'not found'}, 404)
    f = flatten(r)
    sh = (body.shell or '').upper()
    if sh == 'OPEN':
        for k in ('s_lwh', 's_wt', 's_tank'):
            f[k] = 'TBA'
    elif sh == 'SILENT':
        for k in ('o_lwh', 'o_wt', 'o_tank'):
            f[k] = 'TBA'
    try:
        import build_pptx
        os.makedirs(OUT_DIR, exist_ok=True)
        _lg, _tp = _tpl_for(body.lang)
        pptx = build_pptx.build(f, _lg, _tp, OUT_DIR)
        pdf = pptx.rsplit('.', 1)[0] + '.pdf'
        if not os.path.exists(pdf):
            subprocess.run([_SOFFICE, '--headless',
                            '--convert-to', 'pdf', '--outdir', OUT_DIR, pptx],
                           check=True, capture_output=True, timeout=180)
        return {'ok': True, 'pdf': '/output/' + os.path.basename(pdf),
                'pptx': '/output/' + os.path.basename(pptx),
                'dlname': doc_filename(f, _lg)}
    except Exception as e:
        import traceback
        return JSONResponse({'ok': False, 'error': str(e), 'trace': traceback.format_exc()[-800:]}, 500)


class PreviewIn(BaseModel):
    kind: str = 'std'
    name: str = ''
    shell: str = ''
    data: dict = {}
    page: int = 1
    lang: str = 'zh'


_PREV_LOCK = threading.Lock()


@app.post('/api/preview')
def api_preview(body: PreviewIn):
    """用当前编辑值走一遍真实模板渲染，返回前 N 页的 PNG（所见即所得预览）。"""
    try:
        import build_pptx
        os.makedirs(OUT_DIR, exist_ok=True)
        _lg, _tp = _tpl_for(body.lang)
        f = dict(body.data or {})
        if not f:
            c = conn(body.kind)
            r = c.execute("select * from models where name=?", (body.name,)).fetchone()
            c.close()
            if not r:
                return JSONResponse({'ok': False, 'error': 'not found'}, 404)
            f = flatten(r)
        sh = (body.shell or '').upper()
        if sh == 'OPEN':
            for k in ('s_lwh', 's_wt', 's_tank'):
                f[k] = 'TBA'
        elif sh == 'SILENT':
            for k in ('o_lwh', 'o_wt', 'o_tank'):
                f[k] = 'TBA'
        # 派生字段（与导出保持一致）
        if not f.get('freq_phase_pf'):
            fq = f.get('cur_freq') or ''
            f['freq_phase_pf'] = f.get('freq') or fq
        f['fc_lines'] = [f.get('fc_esp'), f.get('fc_prp'), f.get('fc_75'), f.get('fc_50')]

        # ---- 优先走 Mac 渲染（16GB 内存，约 4.5s）；失败则本地兜底 ----
        mac_err = None
        if _mac_health():
            try:
                names, pdf_name, _ = _mac_render(f, _lg)
                pngs = []
                for fn in names:
                    _mac_fetch(fn, OUT_DIR)
                    pngs.append('/output/' + fn)
                if pdf_name:
                    try:
                        _mac_fetch(pdf_name, OUT_DIR)
                    except Exception:
                        pass
                if pngs:
                    _cleanup_output()
                    return {'ok': True, 'pages': pngs, 'pdf': ('/output/' + pdf_name) if pdf_name else None,
                            'host': 'mac'}
            except Exception as e:
                mac_err = str(e)

        # ---- 本地兜底（服务器自己渲染，慢且可能 OOM）----
        with _PREV_LOCK:
            pptx = build_pptx.build(f, _lg, _tp, OUT_DIR)
            pdf = pptx.rsplit('.', 1)[0] + '.pdf'
            if not os.path.exists(pdf):
                subprocess.run([_SOFFICE, '--headless', '--convert-to', 'pdf',
                                '--outdir', OUT_DIR, pptx],
                               check=True, capture_output=True, timeout=180)
            stem = pdf.rsplit('.', 1)[0]
            pngs = []
            for i in range(1, 5):
                png = '%s_pg-%d.png' % (stem, i)
                if not os.path.exists(png):
                    subprocess.run(['pdftoppm', '-png', '-r', '110', '-f', str(i), '-l', str(i),
                                    pdf, '%s_pg' % stem],
                                   check=False, capture_output=True, timeout=120)
                if os.path.exists(png):
                    pngs.append('/output/' + os.path.basename(png))
        _cleanup_output()
        out = {'ok': True, 'pages': pngs, 'pdf': '/output/' + os.path.basename(pdf), 'host': 'server'}
        if mac_err:
            out['mac_error'] = mac_err[:200]
        return out
    except Exception as e:
        import traceback
        return JSONResponse({'ok': False, 'error': str(e), 'trace': traceback.format_exc()[-800:]}, 500)


def _cleanup_output(keep=120):
    """2026-11-17：每次渲染都会新增 pptx/pdf/4 张 png，从不清理会把磁盘塞满。
    保留最近 keep 个文件（按 mtime），其余删掉。"""
    try:
        items = []
        for fn in os.listdir(OUT_DIR):
            fp = os.path.join(OUT_DIR, fn)
            if os.path.isfile(fp):
                items.append((os.path.getmtime(fp), fp))
        if len(items) <= keep:
            return 0
        items.sort(reverse=True)
        n = 0
        for _, fp in items[keep:]:
            try:
                os.remove(fp)
                n += 1
            except Exception:
                pass
        return n
    except Exception:
        return 0


@app.get('/output/{fn}')
def serve_out(fn: str):
    p = os.path.join(OUT_DIR, fn)
    return FileResponse(p) if os.path.exists(p) else JSONResponse({'error': 'nf'}, 404)


# ---------------- 保存 / 新增 / 删除（仅管理员） ----------------
class SaveIn(BaseModel):
    kind: str = 'std'
    name: str = ''            # 原主键；新增时为空
    data: dict                # 前端编辑后的摊平字段（S.data）
    is_new: bool = False      # 新增标记：true 时走 insert，不做"名称不能为空"拦截


def _is_admin(request: Request):
    return request.cookies.get('admin') == '1'


def _today_version():
    import datetime
    return 'V' + datetime.datetime.now().strftime('%Y%m%d') + '-01'


def _derive(row_name, data, kind):
    """从摊平数据回推派生列（款式/系列/分区/版本）"""
    import re as _re
    sh = (data.get('shell_code') or '').strip().upper() or None
    if not sh:
        m = _re.search(r'开架', row_name)
        if m:
            sh = 'A'
        else:
            m = _re.search(r'([A-Z]{1,2})款', row_name)
            if m:
                sh = m.group(1)
    fam = (data.get('family_code') or '').strip().upper() or None
    return sh, fam, (data.get('zone') or ('std' if kind == 'std' else 'std'))


@app.post('/api/save')
def api_save(body: SaveIn, request: Request):
    """保存参数单：标准单直接覆盖原单（用户 2026-11-13 定口径，不留历史）"""
    if not _is_admin(request):
        return JSONResponse({'ok': False, 'error': '需要管理员登录'}, 403)
    kind = body.kind if body.kind in ('std', 'nonstd') else 'std'
    db = STD_DB if kind == 'std' else NON_DB
    d = dict(body.data or {})
    name = (body.name or '').strip()
    is_new = bool(body.is_new) or not name
    if is_new:
        # 新增：主键由前端拼好（型号 (电压款式) [发动机]）
        name = (d.get('name') or name).strip()
    if not name:
        return JSONResponse({'ok': False, 'error': '缺少参数单名称'}, 400)

    c = sqlite3.connect(db)
    c.row_factory = sqlite3.Row
    old = c.execute("select * from models where name=?", (name,)).fetchone()
    if old is not None and is_new:
        c.close()
        return JSONResponse({'ok': False, 'error': '同名参数单已存在：' + name}, 409)

    # 回写嵌套 data：保留未在界面出现的字段，只覆盖界面编辑过的那些
    if old is not None:
        raw = json.loads(old['data'])
    else:
        raw = {'model': d.get('model') or d.get('name'), 'curFreq': d.get('cur_freq') or '50Hz',
               'freqs': {}, 'common': {}}
    cur_f = raw.get('curFreq') or '50Hz'
    freqs = raw.setdefault('freqs', {})
    fv = freqs.setdefault(cur_f, {})
    cm = raw.setdefault('common', {})
    for k in FREQ_KEYS:
        if k in d:
            fv[k] = d.get(k)
    for k in COMMON_KEYS:
        if k in d:
            cm[k] = d.get(k)
    # 界面上的开架/静音尺寸（o_*/s_*）已在 COMMON_KEYS 内
    if d.get('model'):
        raw['model'] = d.get('model')
    if d.get('cur_freq'):
        raw['curFreq'] = d.get('cur_freq')
        cur_f = d.get('cur_freq')
    # 兜底：频率/相数/功率因素格
    if not fv.get('freq'):
        fv['freq'] = d.get('freq_phase_pf') or cur_f

    sh, fam, zone = _derive(name, d, kind)
    version = d.get('version') or (old['version'] if old is not None else _today_version())
    now = __import__('datetime').datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    created = old['created_at'] if old is not None else now
    payload = json.dumps(raw, ensure_ascii=False)

    if old is None:
        c.execute("""insert into models(name,model,family,freq_list,src,data,created_at,updated_at,
                     shell_code,shell_name,family_code,family_name,zone,version,display_name)
                     values(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)""",
                  (name, raw.get('model'), (fam or ''), json.dumps(list(freqs.keys())),
                   d.get('src') or 'web新增', payload, created, now, sh, sh, fam,
                   d.get('family_name'), zone, version, d.get('display_name')))
        action = 'created'
    else:
        c.execute("""update models set model=?, data=?, updated_at=?, shell_code=?, shell_name=?,
                     family_code=?, family_name=?, zone=?, version=? where name=?""",
                  (raw.get('model'), payload, now, sh, sh, fam,
                   d.get('family_name') or old['family_name'], zone, version, name))
        action = 'updated'
    # 同步 FTS
    try:
        c.execute("delete from models_fts where name=?", (name,))
        c.execute("""insert into models_fts(name,model,shell_name,family_name,e_model,e_brand,std_v,zone)
                     values(?,?,?,?,?,?,?,?)""",
                  (name, raw.get('model'), sh, d.get('family_name') or '',
                   cm.get('e_model'), cm.get('e_brand'), cm.get('std_v'), zone))
    except Exception:
        pass
    c.commit()
    c.close()
    return {'ok': True, 'action': action, 'name': name, 'version': version, 'updated_at': now}


class DeleteIn(BaseModel):
    kind: str = 'std'
    name: str = ''


@app.post('/api/delete')
def api_delete(body: DeleteIn, request: Request):
    """删除参数单：删标准单时其下非标跟着删"""
    if not _is_admin(request):
        return JSONResponse({'ok': False, 'error': '需要管理员登录'}, 403)
    kind = body.kind if body.kind in ('std', 'nonstd') else 'std'
    name = (body.name or '').strip()
    if not name:
        return JSONResponse({'ok': False, 'error': '缺少名称'}, 400)
    db = STD_DB if kind == 'std' else NON_DB
    c = sqlite3.connect(db)
    c.row_factory = sqlite3.Row
    n = c.execute("select count(*) from models where name=?", (name,)).fetchone()[0]
    if not n:
        c.close()
        return JSONResponse({'ok': False, 'error': 'not found'}, 404)
    deleted_nonstd = []
    if kind == 'std':
        # 删标准单 → 其下非标跟着删（按型号前缀配对）
        base = parse_name(name)['model_code'] or ''
        nc = sqlite3.connect(NON_DB)
        nc.row_factory = sqlite3.Row
        rows = nc.execute("select name from models where name like ?", (base + '%',)).fetchall()
        for r in rows:
            if (parse_name(r['name'])['model_code'] or '').upper() == base.upper():
                nc.execute("delete from models where name=?", (r['name'],))
                deleted_nonstd.append(r['name'])
        nc.commit()
        nc.close()
    c.execute("delete from models where name=?", (name,))
    try:
        c.execute("delete from models_fts where name=?", (name,))
    except Exception:
        pass
    c.commit()
    c.close()
    return {'ok': True, 'deleted': name, 'deleted_nonstd': deleted_nonstd}


# ---------------- 页面 ----------------
@app.get('/', response_class=HTMLResponse)
def index():
    p = os.path.join(BASE, 'web', 'index.html')
    return HTMLResponse(open(p, encoding='utf-8').read())


if __name__ == '__main__':
    import uvicorn
    uvicorn.run(app, host='0.0.0.0', port=8080, log_level='info')
