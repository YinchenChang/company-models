# Siliconomics 快照讀取（下游建檔程式共用）。用法：
#   import sys, os; sys.path.insert(0, os.path.join(ROOT, "siliconomics")); import reader as SIL
#   inp(*SIL.row9("pa_amd"))        # 9 欄 inp（鍵、項目、單位、低、基準、高、標記、來源1、來源2）
#   SIL.base("pa_amd")；SIL.V["pa_amd"]（整列）；SIL.NAMES["SIF_UnitsGrowth29"]（Interface 推導值）
# 只讀 data/CURRENT_SNAPSHOT 指向的快照；數值的權威在 Siliconomics Excel 的 In 頁。
import json, os

_D = os.path.join(os.path.dirname(os.path.abspath(__file__)), "data")
FILE = open(os.path.join(_D, "CURRENT_SNAPSHOT"), encoding="utf-8").read().strip()
SNAP = json.load(open(os.path.join(_D, FILE), encoding="utf-8"))
VERSION = SNAP["source"]["version"]
NAMES = SNAP["names"]
# Excel 不分整數與浮點數：比例、倍數、百萬顆、單價（$/GB）用浮點數；其他單位整數值用整數（與搬入前程式內的寫法相同）
_FLOAT_UNITS = {"%", "x", "百萬顆", "$/GB"}
ROWS = []
for _r in SNAP["rows"]:
    _r = dict(_r)
    if "id" in _r:
        for _f in ("lo", "base", "hi"):
            _x = _r[_f]
            if isinstance(_x, (int, float)) and not isinstance(_x, bool):
                _r[_f] = float(_x) if _r["unit"] in _FLOAT_UNITS else (int(_x) if float(_x).is_integer() else _x)
    ROWS.append(_r)
V = {r["id"]: r for r in ROWS if "id" in r}


def base(k):
    return V[k]["base"]


def row8(k):
    r = V[k]
    return (k, r["label"], r["unit"], r["lo"], r["base"], r["hi"], r["tag"], r["src"])


def row9(k):
    return row8(k) + (V[k].get("src2", ""),)


def vals(k):
    """(低, 基準, 高)：只取數值、保留呼叫端自己的標籤與說明時用"""
    r = V[k]
    return (r["lo"], r["base"], r["hi"])
