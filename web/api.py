# -*- coding: utf-8 -*-
"""参数单管理工具 — 后端接口（第2步，FastAPI 版）"""
import os, sys, json, sqlite3, re, subprocess
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
ADMIN_USER, ADMIN_PASS = 'bob', '12345'

app = FastAPI(title='参数单管理工具')


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
    c.close()
    return {'shell': shell, 'family': family, 'voltage': volt, 'total': total, 'pending': pend, 'kind': kind}


# ---------------- 列表 ----------------
@app.get('/api/list')
def api_list(kind: str = 'std', q: str = '', family: str = '', shell: str = '',
             zone: str = '', page: int = 1, size: int = 20):
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
    if q.strip():
        like = '%' + q.strip() + '%'
        where.append("""(m.name like ? or m.family_name like ? or m.shell_name like ?
             or json_extract(m.data,'$.common.e_model') like ?
             or json_extract(m.data,'$.common.e_brand') like ?
             or json_extract(m.data,'$.common.std_v') like ?)""")
        params += [like] * 6
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
def compare_with_std(nonstd):
    base = (parse_name(nonstd['name'])['model_code'] or '').upper()
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
class ExportIn(BaseModel):
    kind: str = 'std'
    name: str = ''
    shell: str = ''


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
        pptx = build_pptx.build(f, 'zh', os.path.join(BASE, 'templates', 'template_zh.pptx'), OUT_DIR)
        pdf = pptx.rsplit('.', 1)[0] + '.pdf'
        if not os.path.exists(pdf):
            subprocess.run(['/Applications/LibreOffice.app/Contents/MacOS/soffice', '--headless',
                            '--convert-to', 'pdf', '--outdir', OUT_DIR, pptx],
                           check=True, capture_output=True, timeout=180)
        return {'ok': True, 'pdf': '/output/' + os.path.basename(pdf),
                'pptx': '/output/' + os.path.basename(pptx)}
    except Exception as e:
        import traceback
        return JSONResponse({'ok': False, 'error': str(e), 'trace': traceback.format_exc()[-800:]}, 500)


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
