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
# 三層基準協調規則 v1.4：情境 B 基準顆數係數（um28_B、um29_B）改讀 SIF_UnitsCoef_Base（siliconomics/coef_rule.py）
import sys as _sys
_sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import coef_rule as _CR
RULE_COEF_KEYS = _CR.apply(ROWS, NAMES, VERSION)
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


# ---- v0.12：個股自訂情境（決策 S35、S36；工作方式 v1.7「三層只給基準、情境在個股」） ----
# 本庫基準＝原始顆數 × 基準係數（SIF_um28_Base、SIF_um29_Base）；個股情境係數（÷ 原始顆數）由個股自訂，
# 說明欄須寫明「相對本庫基準的倍數」與「產業 GW 相對 Wattonomics 基準 ±%」（本庫基準相對 Wattonomics 見 SIF_GWvsWT_Base）。
def _pct(x):
    return f"{x:+.0%}".replace("+-", "−").replace("-", "−")


def scen_note(s, yy, val):
    """個股情境係數列（um28_s／um29_s）的說明文字。s＝情境字母；yy＝28 或 29；val＝本模型該列的值（÷ 原始顆數）。"""
    bc = V[f"um{yy}_B"]["base"] if f"um{yy}_B" in RULE_COEF_KEYS else NAMES.get(f"SIF_um{yy}_Base")   # v0.19：依規則的基準係數
    gv = NAMES.get("SIF_GWvsWT_Base")
    if bc is None or gv is None:
        return "個股自訂情境（Siliconomics 快照未提供基準係數）"
    rel = val / bc
    yi = {28: 2, 29: 3}[yy]
    gw = (1 + gv[yi]) * rel - 1
    who = "＝Siliconomics 基準" if abs(rel - 1) < 1e-9 else f"＝Siliconomics 基準 ×{rel:.3f}"
    tail = ("；2026–27 三情境相同，產業 GW 相對 Wattonomics 基準 " + "、".join(f"20{26+i} {_pct(gv[i])}" for i in (0, 1))) if yy == 28 else ""
    return (f"個股自訂情境 {s}（{VERSION} 起；預設沿用 v0.11 行業情境值，判斷類可改）：{who}（基準係數 {bc}）；"
            f"20{yy} 產業 GW 相對 Wattonomics 基準 {_pct(gw)}（本庫基準 {_pct(gv[yi])}）{tail}")


# ---- v0.13：Wattonomics 錨定係數（三層校準原則，Andy 2026-11-04；決策 S37） ----
# 個股敏感度「Wattonomics 錨定」案例：以個股情境 B（＝Siliconomics 基準）為底，產業量（顆數、部署 GW、市場）自 CY2027 起 × WTA[年]；
# FY2026 維持錨點（實績＋指引），不乘。WTA＝R4 比值（Wattonomics 基準 ÷ 本庫推得上線 GW）。
WTA = dict(zip((2026, 2027, 2028, 2029), NAMES.get("SIF_WTAnchor_Base") or (1, 1, 1, 1)))
WTA_SRC = f"Siliconomics {VERSION} SIF_WTAnchor_Base（R4 比值 Wattonomics ÷ 本庫；Recon B 節）"


def wta_rows(inp, pre="wta"):
    """在 In 頁加三列 wta_27／wta_28／wta_29（Derived）。inp＝建檔程式的 9 欄 inp 函式。"""
    for y in (2027, 2028, 2029):
        inp(f"{pre}_{y % 100}", f"Wattonomics 錨定係數 CY{y}（敏感度案例用；產業量 × 本係數）", "x", None, WTA[y], None, "Derived", WTA_SRC,
            "三層校準原則（Andy 2026-11-04）：Wattonomics 為部署水位錨點；個股基準暫不改，2027-01～02 雲端資本支出指引後再決定" if y == 2027 else "同上")
