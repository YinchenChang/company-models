# Meta v0.2.2（原地修改，工作單 px1113 第 3 節）：數位廣告同業檔 EV／NTM EBITDA 由報價日（2026-10-07／10-08；Meta 自身 10-28）換到 2026-11-13 收盤。
# 方法：
#   GOOGL、META（自身）：檔內有 EV 與股數（市值）→ EV 新＝EV 舊＋市值舊 ×（11-13 收盤 ÷ 報價日收盤 − 1）；NTM EBITDA 不變，倍數＝EV 新 ÷ NTM EBITDA。
#     GOOGL 市值舊＝350.50 × 12,294M 股＝4,309.0（檔內 method 欄）；META 市值舊＝612.49 × 2,556.24M 股＝1,565.7（selfReference.ev.method）。
#   PINS、TTD、APP、RDDT、SNAP：檔內無市值 → 倍數 × 收盤比近似（EV 與倍數同比例縮放；淨負債不變的假設使變動被高估），標 [Derived] 近似。
# 收盤價只取 markets/invest_map/data/closes_20261020_20261113.csv。EBITDA、NTM 權重等基本面不變。
# 用法：python3 meta/px1113_peers.py → data/peers_ads_20261113.json（build_meta.py 讀；原檔 peers_ads_20261030.json 保留）
import csv, json, os, statistics
HERE = os.path.dirname(os.path.abspath(__file__))
D = os.path.join(HERE, "data")
C = {}
for r in csv.DictReader(open(os.path.join(HERE, "..", "markets", "invest_map", "data", "closes_20261020_20261113.csv"), encoding="utf-8")):
    C[(r["symbol"], r["date"])] = float(r["close"])
NEW = "2026-11-13"
MCAP = {"GOOGL": 350.50 * 12294e-3, "META": 612.49 * 2556.24e-3}   # US$ bn
j = json.load(open(os.path.join(D, "peers_ads_20261030.json"), encoding="utf-8"))
old_med = j["median"]["value"]
def conv(p, d):
    t = p["ticker"]; a, b = C[(t, d)], C[(t, NEW)]; k = b / a
    p["ntmEvEbitda1007"] = p["ntmEvEbitda"]; p["ev1007"] = p["ev"]["value"]; p["refDate"] = d
    p["closeOld"] = a; p["closeNew"] = b; p["priceRatio"] = round(k, 6)
    if t in MCAP:
        ev_new = p["ev"]["value"] + MCAP[t] * (k - 1)
        p["ntmEvEbitda"] = round(ev_new / p["ntmEbitda"]["value"], 2)
        how = f"EV 新＝EV 舊 {p['ev1007']}＋市值舊 {MCAP[t]:.1f} ×（{b} ÷ {a} − 1）＝{ev_new:.1f}；NTM EBITDA {p['ntmEbitda']['value']} 不變"
        p["ev"]["value"] = round(ev_new, 1)
    else:
        p["ntmEvEbitda"] = round(p["ntmEvEbitda"] * k, 2); p["ev"]["value"] = round(p["ev"]["value"] * k, 2)
        how = f"EV 舊 {p['ev1007']} × 收盤比 {b}/{a}；近似，無市值，淨負債不變"
    p["ev"]["asOf"] = NEW; p["ev"]["note"] = f"2026-11-13 換價（工作單 px1113；{d} → 11-13）：" + how
    p["tag"] = "[Derived]" if t in MCAP else "[Derived]（近似）"
for p in j["peers"]:
    conv(p, p["ev"]["asOf"])
conv(j["selfReference"], "2026-10-28")
j["median1030"] = j["median"]
j["median"] = {"value": round(statistics.median(p["ntmEvEbitda"] for p in j["peers"]), 2), "n": len(j["peers"]), "tag": "[Derived]（近似）", "note": f"6 家（GOOGL＋五家）中位數；2026-11-13 換價（舊 {old_med}）"}
j["asOf"] = "2026-11-13（EV 與倍數由報價日依 Yahoo Finance 收盤換算；GOOGL、META 用市值法，其餘倍數 × 收盤比近似）"; j["asOfOriginal"] = "2026-10-07（五家與 GOOGL）、2026-10-28（META 自身）"
j["priceSource"] = "markets/invest_map/data/closes_20261020_20261113.csv"
j["_readme"] = j.get("_readme", "") + "；v0.2.2 原地換價至 2026-11-13（meta/px1113_peers.py；工作單 px1113）"
json.dump(j, open(os.path.join(D, "peers_ads_20261113.json"), "w", encoding="utf-8"), ensure_ascii=False, indent=1)
print("median", old_med, "->", j["median"]["value"], [(p["ticker"], p["ntmEvEbitda1007"], p["ntmEvEbitda"]) for p in j["peers"]], "self", j["selfReference"]["ntmEvEbitda1007"], "->", j["selfReference"]["ntmEvEbitda"])
