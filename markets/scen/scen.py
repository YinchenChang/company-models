# 市場層產業情境（A／B／C）讀取（Siliconomics v0.16 起；決策 S41）。
# Siliconomics 只給基準（＝原情境 B，工作方式 v1.7「三層只給基準、情境在個股」）；產業情境 A、C 的參考值由本資料夾持有
# （industry_scen_v*.json，自 Siliconomics v0.15 In 頁逐字搬出）。B 一律讀 Siliconomics 快照，本檔不另存 B 值。
# 用法：
#   import sys, os; sys.path.insert(0, os.path.join(ROOT, "markets", "scen")); import scen as SCN
#   SCN.V["cpo_A_27"]（整列）；SCN.base("um28_C")；SCN.row8(...)／row9(...)／vals(...)：A／C 讀本檔、其他鍵（含 B）讀 Siliconomics
#   SCN.SCEN＝("A","B","C")；SCN.SCEN_DESC；SCN.merged_rows(sil_rows)：把 A／C 列插回 Siliconomics In 列序（記憶體公司 In 頁重播用，列序與 v0.15 相同）
import glob, json, os

_HERE = os.path.dirname(os.path.abspath(__file__))
_ROOT = os.path.dirname(os.path.dirname(_HERE))


def _vkey(f):
    v = os.path.basename(f).rsplit("_v", 1)[1].rsplit(".json", 1)[0]
    return tuple(int(x) for x in v.split("."))


FILE = sorted(glob.glob(os.path.join(_HERE, "industry_scen_v*.json")), key=_vkey)[-1]
DOC = json.load(open(FILE, encoding="utf-8"))
VERSION = DOC["version"]
SCEN = tuple(DOC["scen"])
SCEN_DESC = dict(DOC["scen_desc"])
SCEN_NAME = dict(DOC["scen_name"])
SRC = f"markets/scen {VERSION}（{os.path.basename(FILE)}）"

# Excel 不分整數與浮點數：比例、倍數、百萬顆、單價用浮點數（與 siliconomics/reader.py 相同）
_FLOAT_UNITS = {"%", "x", "百萬顆", "$/GB"}
OWN = []
for _r in DOC["rows"]:
    _r = dict(_r)
    for _f in ("lo", "base", "hi"):
        _x = _r[_f]
        if isinstance(_x, (int, float)) and not isinstance(_x, bool):
            _r[_f] = float(_x) if _r["unit"] in _FLOAT_UNITS else (int(_x) if float(_x).is_integer() else _x)
    OWN.append(_r)
_OWN = {r["id"]: r for r in OWN}

# Siliconomics 快照（B 與其他鍵）
_SILD = os.path.join(_ROOT, "siliconomics", "data")
_SNAP = json.load(open(os.path.join(_SILD, open(os.path.join(_SILD, "CURRENT_SNAPSHOT"), encoding="utf-8").read().strip()), encoding="utf-8"))
# 三層基準協調規則 v1.4：情境 B 基準顆數係數改讀 SIF_UnitsCoef_Base（siliconomics/coef_rule.py；與 reader、memory/market 同一函式）
import sys as _sys
_sys.path.insert(0, os.path.join(_ROOT, "siliconomics"))
import coef_rule as _CR
_CR.apply(_SNAP["rows"], _SNAP["names"], _SNAP["source"]["version"])
_SILV = {}
for _r in _SNAP["rows"]:
    if "id" in _r:
        _r = dict(_r)
        for _f in ("lo", "base", "hi"):
            _x = _r[_f]
            if isinstance(_x, (int, float)) and not isinstance(_x, bool):
                _r[_f] = float(_x) if _r["unit"] in _FLOAT_UNITS else (int(_x) if float(_x).is_integer() else _x)
        _SILV[_r["id"]] = _r
# v0.16 起 Siliconomics 快照不再有 A／C 列；讀到舊快照（v0.15 以前，升版建置的第一輪）時兩邊必須逐值相同
for _k in set(_OWN) & set(_SILV):
    assert all(_OWN[_k][f] == _SILV[_k][f] for f in ("lo", "base", "hi")), f"markets/scen 與 Siliconomics 快照不一致：{_k}"
    del _SILV[_k]


class _V(dict):
    def __missing__(self, k):
        return _SILV[k]


V = _V(_OWN)


def base(k):
    return V[k]["base"]


def row8(k):
    r = V[k]
    return (k, r["label"], r["unit"], r["lo"], r["base"], r["hi"], r["tag"], r["src"])


def row9(k):
    return row8(k) + (V[k].get("src2", ""),)


def vals(k):
    r = V[k]
    return (r["lo"], r["base"], r["hi"])


def merged_rows(sil_rows):
    """Siliconomics In 列（含節標題）＋本檔 A／C 列，依 after 錨點插回 v0.15 的列序。每列都必須找到錨點。"""
    pend = {}
    for r in OWN:
        pend.setdefault(r["after"], []).append(r)
    out, placed = [], 0
    for r in sil_rows:
        if "id" in r and r["id"] in _OWN:   # 舊快照（v0.15 以前）仍含 A／C 列：略過，改用本檔
            continue
        out.append(r)
        if "sec" in r:
            key = next((a for a in pend if a.startswith("sec:") and r["sec"].startswith(a[4:])), None)
        else:
            key = r["id"]
        while key in pend:
            add = pend.pop(key)
            out += [{k: v for k, v in x.items() if k != "after"} for x in add]
            placed += len(add)
            key = add[-1]["id"]   # 連續搬出的列（例如 cqA_8 … cqC_19）依原順序接續
    assert placed == len(OWN) and not pend, f"markets/scen：{len(OWN) - placed} 列找不到錨點 {list(pend)[:3]}"
    return out
