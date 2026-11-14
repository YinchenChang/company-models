# 三層基準協調規則 v1.4（docs/plan/三層基準協調規則.md C3、C4）：下游讀到的「情境 B＝基準顆數係數」改為依規則決定的值。
#
# 背景：Siliconomics v0.19 起，基準鏈上的顆數係數＝Coord 頁依規則決定的 crd_k（Interface 名稱 SIF_UnitsCoef_Base，
# 2026–2029 各一值）；In 頁 um28_B、um29_B 保留原 S33 值（0.8／0.85）只作參照。下游約 60 個建檔程式與市場層
# 都經 siliconomics/reader.py、markets/scen/scen.py、markets/memory/market.py 讀 um28_B／um29_B，
# 本函式在這三個讀取點把這兩列的基準值換成 SIF_UnitsCoef_Base，使規則基準一次到達全部下游（交接 5b）。
#
# - 只在快照有 SIF_UnitsCoef_Base 時作用（v0.18 以前的快照不變）。
# - 基準＝SIF_UnitsCoef_Base 的 2028、2029 值；低、高＝原 S33 區間依同比例縮放（規則係數 ÷ S33 基準）。
#   縮放是預設做法（判斷類，Andy 可改）：保留原區間的相對寬度，不另設新假設。
# - 原 S33 值寫在說明欄，數值不刪。
_YEARS = {"um28_B": 2028, "um29_B": 2029}


def _yidx(coef):
    """SIF_UnitsCoef_Base 為 [None, 2026, 2027, 2028, 2029] 或 [2026..2029]；回傳 年→值。"""
    vals = list(coef)
    return dict(zip(range(2029 - len(vals) + 1, 2030), vals))


def apply(rows, names, version):
    """就地修改 rows（快照 rows 清單）中 um28_B、um29_B 兩列；回傳改動的鍵清單。"""
    coef = names.get("SIF_UnitsCoef_Base")
    if not coef:
        return []
    by_year = _yidx(coef)
    done = []
    for r in rows:
        k = r.get("id")
        if k not in _YEARS:
            continue
        new = by_year[_YEARS[k]]
        old = r["base"]
        if new is None or old in (None, 0):
            continue
        f = new / old
        lo, hi = r.get("lo"), r.get("hi")
        r["base"] = float(new)
        if isinstance(lo, (int, float)):
            r["lo"] = float(lo) * f
        if isinstance(hi, (int, float)):
            r["hi"] = float(hi) * f
        r["tag"] = "Derived"
        same = abs(f - 1) < 1e-12
        r["src"] = (f"Siliconomics {version} SIF_UnitsCoef_Base（三層基準協調規則 v1.4 C3 依證據狀態決定的基準係數"
                    + ("；本年維持現行，與原 S33 值相同" if same else
                       f"；原 S33 值 {old:g}（區間 {lo:g}–{hi:g}）只作參照，區間依同比例縮放")
                    + "）")
        r["label"] = r["label"].replace("本庫基準係數", "本庫基準係數（依三層基準協調規則）") if "依三層基準協調規則" not in r["label"] else r["label"]
        done.append(k)
    return done
