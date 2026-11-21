# Amazon v0.3.9（工作單 px1120）：同業檔（非 AI 雲端、數位廣告、零售）倍數由 2026-11-13 換到 2026-11-20 收盤。
# 方法同 px1113_peers.py：倍數新 ＝ 現行倍數 × 11-20 收盤 ÷ 現行報價日（11-13）收盤；無市值，EV 同比例縮放，標 [Derived] 近似。
# 收盤只取 markets/invest_map/data/closes_20261109_20261120.csv（11-13 與舊檔相同）；缺代號維持原值並列出。
# 用法：python3 amazon/px1120_peers.py → data/peers_{cloud,ads,retail}_20261120.json（build 讀）
import csv, json, os, statistics
HERE = os.path.dirname(os.path.abspath(__file__))
D = os.path.join(HERE, "data")
C = {}
for r in csv.DictReader(open(os.path.join(HERE, "..", "markets", "invest_map", "data", "closes_20261109_20261120.csv"), encoding="utf-8")):
    C[(r["symbol"], r["date"])] = float(r["close"])
OLD, NEW = "2026-11-13", "2026-11-20"
def med(v): return round(statistics.median(v), 2)
unchanged = []
for src, dst in (("peers_cloud_20261113.json", "peers_cloud_20261120.json"), ("peers_ads_20261113.json", "peers_ads_20261120.json"), ("peers_retail_20261113.json", "peers_retail_20261120.json")):
    j = json.load(open(os.path.join(D, src), encoding="utf-8"))
    old_med = j["median"]["value"]
    for p in j["peers"]:
        t = p["ticker"]
        if (t, OLD) not in C or (t, NEW) not in C:
            unchanged.append(t); continue
        a, b = C[(t, OLD)], C[(t, NEW)]; k = b / a
        assert abs(a - p["closeNew"]) < 1e-6 * a + 0.005, (t, a, p["closeNew"])   # 現行報價日收盤須與舊檔一致
        p["ntmEvEbitda1113"] = p["ntmEvEbitda"]; p["ev1113"] = p["ev"]["value"]
        p["closeOld"] = a; p["closeNew"] = b; p["priceRatio"] = round(k, 6); p["refDate"] = OLD
        p["ntmEvEbitda"] = round(p["ntmEvEbitda"] * k, 2); p["ev"]["value"] = round(p["ev"]["value"] * k, 2)
        p["ev"]["asOf"] = NEW; p["ev"]["note"] = f"2026-11-20 換價（工作單 px1120）：EV 舊 {p['ev1113']}（{OLD}）× 收盤比 {b}/{a}；近似，無市值，淨負債不變"
    j["median1113"] = j["median"]
    j["median"] = dict(value=med(p["ntmEvEbitda"] for p in j["peers"]), n=len(j["peers"]), tag="[Derived]（近似）", note=f"{len(j['peers'])} 家中位數；2026-11-20 換價（11-13 為 {old_med}）")
    j["asOf"] = "2026-11-20（倍數＝11-13 倍數 × Yahoo Finance 收盤比，近似）"
    j["priceSource"] = "markets/invest_map/data/closes_20261109_20261120.csv"
    j["_readme"] = j.get("_readme", "") + "；v0.3.9 再換價至 2026-11-20（px1120_peers.py；工作單 px1120）"
    json.dump(j, open(os.path.join(D, dst), "w", encoding="utf-8"), ensure_ascii=False, indent=1)
    print(dst, old_med, "->", j["median"]["value"], [(p["ticker"], p["ntmEvEbitda1113"] if "ntmEvEbitda1113" in p else p["ntmEvEbitda"], p["ntmEvEbitda"]) for p in j["peers"]])
print("未換價同業:", unchanged or "無")
