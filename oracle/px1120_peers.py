# Oracle v0.5.4（原地修改，工作單 px1120 第 3 節）：同業報價由 2026-11-13 換到 2026-11-20 收盤。
# 方法同 px1113_peers.py：軟體同業 倍數 × 11-20 收盤 ÷ 11-13 收盤（近似）；MSFT 用市值法（市值以 10-07 值依 11-13、11-20 收盤推算）；
#   neocloud（CRWV、NBIS）EV 新＝EV 現行＋市值（11-13）×（11-20 收盤 ÷ 11-13 收盤 − 1），淨負債不變（近似 [Derived]）。
# 收盤價只取 markets/invest_map/data/closes_20261109_20261120.csv。
# 外部稽核優化 W3-7（2026-11-23；Andy 決定 J6 取預設 (a)）：軟體同業換價改為「分離淨負債」——有市值資料的才改，沒有的維持原式並標示。
#   原式（倍數 × 收盤比）等於假設淨負債跟著股價等比例變動：淨負債為正的公司變動被高估，淨現金的公司被低估。
#   CRM：EV 新＝EV（10-07）214,520＋市值（10-07）×（11-20 收盤 ÷ 10-07 收盤 − 1）；EV 取 oracle_facts_20261007.json val.peer.crm（StockAnalysis，2026-10-07）；
#        市值（10-07）＝流通股數 823.0 百萬股（StockAnalysis CRM 統計頁，頁面日期 2026-11-15；software_ai/salesforce/data/facts_salesforce.csv「流通股數」）× 10-07 收盤 224.56＝184,813；
#        反推淨負債 29,707（同檔「總債務（含租賃，第二來源）」42,380 扣現金後同方向、同量級）。股數日期晚於報價日（期間有買回），改用 10-Q 封面 819 百萬股時倍數差 0.01 倍以內。[Derived]
#   SAP、ADBE、NOW：repo 內沒有 2026-10-07 的市值、股數或淨負債 → 仍為整體縮放（近似）。原式結果一律保留在 newScaled、*Scaled 欄供對照。
CRM_EV_1007, CRM_SH_M = 214520.0, 823.0
# 用法：python3 oracle/px1120_peers.py → oracle/data/peers_sw_20261120.json（build_oracle.py、v052_post.py、scripts/sens_v04.py 讀）
import csv, json, os, statistics
HERE = os.path.dirname(os.path.abspath(__file__))
D = os.path.join(HERE, "data")
C = {}
for r in csv.DictReader(open(os.path.join(HERE, "..", "markets", "invest_map", "data", "closes_20261109_20261120.csv"), encoding="utf-8")):
    C[(r["symbol"], r["date"])] = float(r["close"])
NEW, OLD = "2026-11-20", "2026-11-13"
P13 = json.load(open(os.path.join(D, "peers_sw_20261113.json"), encoding="utf-8"))
out = {}
for t in ("MSFT", "SAP", "CRM", "ADBE", "NOW"):
    p = P13["peers"][t]; cur = p["new"]; a, b = C[(t, OLD)], C[(t, NEW)]; k = b / a
    assert abs(a - p["closeNew"]) < 1e-4, (t, a, p["closeNew"])      # 現行報價日收盤與 px1113 檔一致
    if t == "MSFT":
        k1 = p["priceRatio"]; EV13 = 3990000.0 + 3930000.0 * (k1 - 1); MC13 = 3930000.0 * k1
        f = (EV13 + MC13 * (k - 1)) / EV13
        how = "EV 新＝EV（11-13）＋市值（11-13）×（價格比−1）；EV、市值 11-13 值由 10-07 的 3,990,000、3,930,000 M 依價格比推得；EBITDA 不變"
    elif t == "CRM":      # W3-7：分離淨負債（由 10-07 原值直接換到 11-20）
        c07 = p["closeOld"]; mc07 = CRM_SH_M * c07; ev_new = CRM_EV_1007 + mc07 * (b / c07 - 1)
        f = p["old"] * ev_new / CRM_EV_1007 / cur
        how = f"分離淨負債（W3-7）：EV 新＝EV（10-07）{CRM_EV_1007:,.0f}＋市值（10-07）{mc07:,.0f}（823.0 百萬股 × {c07}）×（{b} ÷ {c07} − 1）＝{ev_new:,.0f}；淨負債 {CRM_EV_1007 - mc07:,.0f} 與 EBITDA 不變；倍數＝10-07 原值 {p['old']} × EV 新 ÷ EV 舊"
    else:
        f = k; how = "倍數 × 收盤比——仍為整體縮放（近似）：repo 內沒有市值或淨負債，等於假設淨負債隨股價等比例變動"
    out[t] = dict(old=cur, new=round(cur * f, 2), newScaled=(round(cur * k, 2) if t != "MSFT" else round(cur * f, 2)), refDate=OLD, closeOld=a, closeNew=b, priceRatio=round(k, 6), method=how,
                  separated=t in ("MSFT", "CRM"), tag="[Derived]" if t in ("MSFT", "CRM") else "[Derived]（近似）", old1007=p["old"])
vals = sorted(v["new"] for v in out.values()); old_vals = sorted(v["old"] for v in out.values())
neo = {}
for t, ev13, mc10_21, c1021 in (("CRWV", 75.33, 34.09, 61.82), ("NBIS", 33.65, 46.64, 171.57)):
    a, b = C[(t, OLD)], C[(t, NEW)]; mc13 = mc10_21 * a / c1021
    neo[t] = dict(evOld=ev13, evNew=round(ev13 + mc13 * (b / a - 1), 2), mcOld=round(mc13, 2), closeOld=a, closeNew=b)
res = dict(asOf=NEW, priceSource="markets/invest_map/data/closes_20261109_20261120.csv", peers=out, neocloud=neo,
           median=round(statistics.median(vals), 2), medianOld=round(statistics.median(old_vals), 2),
           lo=out["CRM"]["new"], hi=out["MSFT"]["new"], loOld=out["CRM"]["old"], hiOld=out["MSFT"]["old"],
           medianScaled=round(statistics.median(v["newScaled"] for v in out.values()), 2), loScaled=out["CRM"]["newScaled"],
           legacyMedianScaled=round(statistics.median([out["CRM"]["newScaled"], out["ADBE"]["newScaled"]]), 2),      # W3-7：原式結果對照
           legacyMedian=round(statistics.median([out["CRM"]["new"], out["ADBE"]["new"]]), 2), legacyMedianOld=round(statistics.median([out["CRM"]["old"], out["ADBE"]["old"]]), 2))
json.dump(res, open(os.path.join(D, "peers_sw_20261120.json"), "w", encoding="utf-8"), ensure_ascii=False, indent=1)
print({t: (v["old"], v["new"]) for t, v in out.items()}, neo, "median", res["medianOld"], "->", res["median"], "lo/hi", res["lo"], res["hi"], "legacy", res["legacyMedianOld"], res["legacyMedian"], "原式對照 median/lo/legacy", res["medianScaled"], res["loScaled"], res["legacyMedianScaled"])
