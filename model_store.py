#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
机型库（服务器端存储）—— SQLite 单文件，替代浏览器 localStorage

表结构 models:
    name        TEXT PRIMARY KEY  机型名（含区分标签，如 "CU330D5 (400V开架)"）
    model       TEXT              纯型号（图片匹配用，如 "CU330D5"）
    family      TEXT              系列（AF / AS / GC / 其他）
    freq_list   TEXT              频率列表 JSON，如 ["50Hz","60Hz"]
    src         TEXT              来源说明
    data        TEXT              完整机型 JSON（{model,curFreq,freqs,common,...}）
    created_at  TEXT
    updated_at  TEXT
"""
import io
import json
import os
import sqlite3
import time

HERE = os.path.dirname(os.path.abspath(__file__))
DB_PATH = os.path.join(HERE, "paramsheet.db")

_SCHEMA = """
CREATE TABLE IF NOT EXISTS models (
    name       TEXT PRIMARY KEY,
    model      TEXT,
    family     TEXT,
    freq_list  TEXT,
    src        TEXT,
    data       TEXT NOT NULL,
    created_at TEXT,
    updated_at TEXT
);
CREATE INDEX IF NOT EXISTS idx_models_model  ON models(model);
CREATE INDEX IF NOT EXISTS idx_models_family ON models(family);
"""


def _conn():
    c = sqlite3.connect(DB_PATH, timeout=15)
    c.row_factory = sqlite3.Row
    return c


def init():
    with _conn() as c:
        c.executescript(_SCHEMA)


def _now():
    return time.strftime("%Y-%m-%d %H:%M:%S")


def list_models(family=None, q=None, limit=0, offset=0):
    """只返回轻量信息（不含完整 data），列表页/下拉框用"""
    sql = "SELECT name,model,family,freq_list,src,updated_at FROM models"
    where, args = [], []
    if family:
        where.append("family = ?"); args.append(family)
    if q:
        where.append("name LIKE ?"); args.append("%" + q + "%")
    if where:
        sql += " WHERE " + " AND ".join(where)
    sql += " ORDER BY family, name"
    if limit:
        sql += " LIMIT ? OFFSET ?"; args += [int(limit), int(offset)]
    with _conn() as c:
        rows = [dict(r) for r in c.execute(sql, args)]
    for r in rows:
        try:
            r["freqs"] = json.loads(r.pop("freq_list") or "[]")
        except Exception:
            r["freqs"] = []
    return rows


def count():
    with _conn() as c:
        return c.execute("SELECT COUNT(*) FROM models").fetchone()[0]


def get(name):
    with _conn() as c:
        r = c.execute("SELECT * FROM models WHERE name = ?", (name,)).fetchone()
    if not r:
        return None
    d = dict(r)
    try:
        d["data"] = json.loads(d["data"])
    except Exception:
        d["data"] = {}
    return d


def save(name, model, data, family="", src="", overwrite=True):
    """保存/更新一个机型。name 为唯一键。"""
    freqs = list((data.get("freqs") or {}).keys())
    now = _now()
    with _conn() as c:
        old = c.execute("SELECT created_at FROM models WHERE name = ?", (name,)).fetchone()
        if old and not overwrite:
            return False
        created = old["created_at"] if old else now
        c.execute(
            "INSERT OR REPLACE INTO models(name,model,family,freq_list,src,data,created_at,updated_at)"
            " VALUES(?,?,?,?,?,?,?,?)",
            (name, model, family, json.dumps(freqs, ensure_ascii=False), src,
             json.dumps(data, ensure_ascii=False), created, now),
        )
    return True


def save_many(items, overwrite=True):
    """批量保存。items: [{name, model, data, family, src}] -> (新增, 更新, 跳过)"""
    added = updated = skipped = 0
    now = _now()
    with _conn() as c:
        for it in items:
            name = it["name"]
            old = c.execute("SELECT created_at FROM models WHERE name = ?", (name,)).fetchone()
            if old and not overwrite:
                skipped += 1
                continue
            data = it.get("data") or {}
            freqs = list((data.get("freqs") or {}).keys())
            created = old["created_at"] if old else now
            c.execute(
                "INSERT OR REPLACE INTO models(name,model,family,freq_list,src,data,created_at,updated_at)"
                " VALUES(?,?,?,?,?,?,?,?)",
                (name, it.get("model") or "", it.get("family") or "",
                 json.dumps(freqs, ensure_ascii=False), it.get("src") or "",
                 json.dumps(data, ensure_ascii=False), created, now),
            )
            if old:
                updated += 1
            else:
                added += 1
    return added, updated, skipped


def delete(name):
    with _conn() as c:
        n = c.execute("DELETE FROM models WHERE name = ?", (name,)).rowcount
    return n > 0


def families():
    with _conn() as c:
        return [dict(r) for r in c.execute(
            "SELECT family, COUNT(*) AS n FROM models GROUP BY family ORDER BY family")]


def all_models():
    """返回全部机型（含完整数据），字段平铺给前端直接渲染。
    结构：[{name, model, family, freq_list, src, curFreq, freqs, common, updated_at}]
    """
    with _conn() as c:
        rows = c.execute(
            "SELECT name,model,family,freq_list,src,data,updated_at "
            "FROM models ORDER BY family, name").fetchall()
    out = []
    for r in rows:
        try:
            d = json.loads(r["data"])
        except Exception:
            d = {}
        freqs = d.get("freqs") or {}
        cur = d.get("curFreq") or (list(freqs)[0] if freqs else "50Hz")
        out.append({
            "name": r["name"],
            "model": r["model"] or d.get("model") or r["name"],
            "family": r["family"] or "",
            "freq_list": json.loads(r["freq_list"]) if r["freq_list"] else list(freqs),
            "src": r["src"] or "",
            "curFreq": cur,
            "freqs": freqs,
            "common": d.get("common") or {},
            "updated_at": r["updated_at"],
        })
    return out


def export_all():
    """导出成前端 DB 格式 {models:[...]}，用于备份"""
    with _conn() as c:
        rows = c.execute("SELECT data FROM models ORDER BY family, name").fetchall()
    models = []
    for r in rows:
        try:
            models.append(json.loads(r["data"]))
        except Exception:
            pass
    return {"models": models}


if __name__ == "__main__":
    init()
    print("db:", DB_PATH, "机型数:", count())
