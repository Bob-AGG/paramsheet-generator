# -*- coding: utf-8 -*-
"""嵌套 data -> 48+ 字段摊平层（后端接口用）"""
import sqlite3, json, re

FREQ_KEYS=['prp_kva','prp_kw','eps_kva','eps_kw','freq','rpm','fc_esp','fc_prp','fc_75','fc_50']
COMMON_KEYS=['std_v','opt_v','e_brand','e_model','e_fuel','e_asp','e_cool','e_gov','e_cyl',
 'o_lwh','o_wt','o_db1','o_db7','o_tank','s_lwh','s_wt','s_db1','s_db7','s_tank',
 'displ','cyl_arr','bore','ex_t','ex_f','ex_bp','fu_res','fu_max','co_air','co_oil','co_cool',
 'el_v','el_bat','el_alt','in_flt','in_res','in_flow','in_t',
 'g_insul','g_exc','g_lead','g_conn','g_pitch','g_phase','g_brg']

def flatten(row):
    """models 一行 -> 摊平 dict（build_pptx 可直接吃）"""
    name=row['name']
    o=json.loads(row['data'])
    cm=o.get('common') or {}
    freqs=o.get('freqs') or {}
    out={}
    # 基础
    out['name']=name
    out['model']=o.get('model') or row['model'] if 'model' in row.keys() else o.get('model')
    out['family']=row['family'] if 'family' in row.keys() else None
    # 派生列
    for k in ('shell_code','shell_name','family_code','family_name','zone','version','src'):
        if k in row.keys(): out[k]=row[k]
    # 频率功率：主频率优先（curFreq），两套都给
    cur_f=o.get('curFreq')
    for fk,label in (('50Hz','50'),('60Hz','60')):
        fv=freqs.get(fk) or {}
        for k in FREQ_KEYS:
            out['%s_%s'%(label,k)]=fv.get(k)
    # 当前频率的那套，直接放顶级（build_pptx 用）
    main=freqs.get(cur_f) or {}
    for k in FREQ_KEYS:
        out[k]=main.get(k)
    out['cur_freq']=cur_f
    # 公共字段
    for k in COMMON_KEYS:
        out[k]=cm.get(k)
    return out

def build_name_fields(name):
    """把 raw_name 拆成 型号/变体/发动机"""
    m=re.match(r'^([^\s\(]+)\s*\(([^)]*)\)\s*(?:\[([^\]]+)\])?', name)
    if not m: return None
    vol=re.match(r'^(\d+V)(.*)$', m.group(2))
    return {
      'model_code': m.group(1),
      'voltage': vol.group(1) if vol else None,
      'variant': (vol.group(2) if vol else m.group(2)).replace('款','') or None,
      'engine_model': m.group(3),
    }

if __name__=='__main__':
    c=sqlite3.connect('paramsheet.db'); c.row_factory=sqlite3.Row
    rows=c.execute("select * from models limit 3").fetchall()
    for r in rows:
        f=flatten(r)
        print('---',f['name'])
        print('   型号:',f['model_code'] if 'model_code' in f else build_name_fields(f['name']))
        print('   款式:',f['shell_code'],f['shell_name'],'| 系列:',f['family_code'],f['family_name'],'| 分区:',f['zone'])
        print('   50Hz prp_kva:',f['50_prp_kva'],'| 60Hz:',f['60_prp_kva'],'| 电压:',f['std_v'],'| 发动机:',f['e_model'])
    c.close()
