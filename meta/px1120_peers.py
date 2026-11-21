# Meta v0.2.3（原地修改，工作單 px1120 第 3 節）：數位廣告同業檔 EV／NTM EBITDA 由 2026-11-13 收盤換到 2026-11-20 收盤。
# 方法同 px1113_peers.py：GOOGL、META（自身）有市值 → EV 新＝EV 舊＋市值舊 ×（11-20 收盤 ÷ 11-13 收盤 − 1），NTM EBITDA 不變；
#   市值舊＝11-13 收盤 × 股數（GOOGL 12,294M 股；META 2,556.24M 股）。PINS、TTD、APP、RDDT、SNAP 無市值 → 倍數與 EV × 收盤比近似，標 [Derived]（近似）。
# 收盤價只取 markets/invest_map/data/closes_20261109_20261120.csv（11-13 與 11-20 皆在此檔）。
# 用法：python3 meta/px1120_peers.py → data/peers_ads_20261120.json（build_meta.py 讀；peers_ads_20261113.json 保留供對照）
import csv, json, os, statistics
HERE = os.path.dirname(os.path.abspath(__file__))
D = os.path.join(HERE, "data")
C = {}
for r in csv.DictReader(open(os.path.join(HERE, "..", "markets", "invest_map", "data", "closes_20261109_20261120.csv"), encoding="utf-8")):
    C[(r["symbol"], r["date"])] = float(r["close"])
OLD, NEW = "2026-11-13", "2026-11-20"
SH = {"GOOGL": 12294e-3, "META": 2556.24e-3}   # 十億股
j = json.load(open(os.path.join(D, "peers_ads_20261113.json"), encoding="utf-8"))
old_med = j["median"]["value"]
def conv(p):
    t = p["ticker"]; a, b = C[(t, OLD)], C[(t, NEW)]; k = b / a
    assert abs(a - p["closeNew"]) < 1e-2, (t, a, p["closeNew"])
    p["ntmEvEbitda1113"] = p["ntmEvEbitda"]; p["ev1113"] = p["ev"]["value"]
    p["closeOld"] = a; p["closeNew"] = b; p["priceRatio"] = round(k, 6); p["refDate"] = OLD
    if t in SH:
        mc = a * SH[t]
        ev_new = p["ev"]["value"] + mc * (k - 1)
        p["ntmEvEbitda"] = round(ev_new / p["ntmEbitda"]["value"], 2)
        how = f"EV 新＝EV 舊 {p['ev1113']}＋市值舊 {mc:.1f}（{a}×{SH[t]*1000:,.2f}M 股）×（{b} ÷ {a} − 1）＝{ev_new:.1f}；NTM EBITDA {p['ntmEbitda']['value']} 不變"
        p["ev"]["value"] = round(ev_new, 1); p["tag"] = "[Derived]"
    else:
        p["ntmEvEbitda"] = round(p["ntmEvEbitda"] * k, 2); p["ev"]["value"] = round(p["ev"]["value"] * k, 2)
        how = f"EV 舊 {p['ev1113']} × 收盤比 {b}/{a}；近似，無市值，淨負債不變"; p["tag"] = "[Derived]（近似）"
    p["ev"]["asOf"] = NEW; p["ev"]["note"] = f"2026-11-20 換價（工作單 px1120；11-13 → 11-20）：" + how
for p in j["peers"]:
    conv(p)
conv(j["selfReference"])
j["median1113"] = j["median"]
j["median"] = {"value": round(statistics.median(p["ntmEvEbitda"] for p in j["peers"]), 2), "n": len(j["peers"]), "tag": "[Derived]（近似）", "note": f"6 家（GOOGL＋五家）中位數；2026-11-20 換價（11-13 為 {old_med}）"}
j["asOf"] = "2026-11-20（EV 與倍數由 2026-11-13 依 Yahoo Finance 收盤換算；GOOGL、META 用市值法，其餘倍數 × 收盤比近似）"
j["priceSource"] = "markets/invest_map/data/closes_20261109_20261120.csv"
j["_readme"] = j.get("_readme", "") + "；v0.2.3 原地換價至 2026-11-20（meta/px1120_peers.py；工作單 px1120）"
json.dump(j, open(os.path.join(D, "peers_ads_20261120.json"), "w", encoding="utf-8"), ensure_ascii=False, indent=1)
print("median", old_med, "->", j["median"]["value"], [(p["ticker"], p["ntmEvEbitda1113"], p["ntmEvEbitda"]) for p in j["peers"]], "self", j["selfReference"]["ntmEvEbitda1113"], "->", j["selfReference"]["ntmEvEbitda"])
