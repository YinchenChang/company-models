# Amazon v0.3.8（原地修改，工作單 px1113 第 3 節）：共用同業檔（非 AI 雲端、數位廣告）EV／NTM EBITDA 由報價日（2026-10-07／10-08）換到 2026-11-13 收盤。
# 方法：倍數新 ＝ 倍數舊 × 11-13 收盤 ÷ 報價日收盤（同業檔無市值，EV 與倍數同比例縮放；淨負債不變的假設使變動被高估），標 [Derived] 近似。EBITDA、NTM 權重不變。
# 收盤價只取 markets/invest_map/data/closes_20261020_20261113.csv（含補入的 09-28～10-19 同業日）。
# 用法：python3 amazon/px1113_peers.py → data/peers_cloud_20261113.json、peers_ads_20261113.json（build 讀）
import csv, json, os, statistics
HERE = os.path.dirname(os.path.abspath(__file__))
D = os.path.join(HERE, "data")
C = {}
for r in csv.DictReader(open(os.path.join(HERE, "..", "markets", "invest_map", "data", "closes_20261020_20261113.csv"), encoding="utf-8")):
    C[(r["symbol"], r["date"])] = float(r["close"])
NEW = "2026-11-13"
def med(v): return round(statistics.median(v), 2)
for src, dst in (("peers_cloud_20261008.json", "peers_cloud_20261113.json"), ("peers_ads_20261009.json", "peers_ads_20261113.json")):
    j = json.load(open(os.path.join(D, src), encoding="utf-8"))
    old_med = j["median"]["value"]
    for p in j["peers"]:
        d = p["ev"]["asOf"]; a, b = C[(p["ticker"], d)], C[(p["ticker"], NEW)]; k = b / a
        p["ntmEvEbitda1007"] = p["ntmEvEbitda"]; p["ev1007"] = p["ev"]["value"]; p["refDate"] = d
        p["closeOld"] = a; p["closeNew"] = b; p["priceRatio"] = round(k, 6)
        p["ntmEvEbitda"] = round(p["ntmEvEbitda"] * k, 2); p["ev"]["value"] = round(p["ev"]["value"] * k, 2)
        p["ev"]["asOf"] = NEW; p["ev"]["note"] = f"2026-11-13 換價（工作單 px1113）：EV 舊 {p['ev1007']}（{d}）× 收盤比 {b}/{a}；近似，無市值，淨負債不變"
    j["median1007"] = j["median"]
    j["median"] = {"value": med(p["ntmEvEbitda"] for p in j["peers"]), "n": len(j["peers"]), "tag": "[Derived]（近似）", "note": f"{len(j['peers'])} 家中位數；2026-11-13 換價（舊 {old_med}）"}
    j["asOf"] = "2026-11-13（倍數＝原報價日倍數 × Yahoo Finance 收盤比，近似）"; j["asOfOriginal"] = "2026-10-07／10-08"
    j["priceSource"] = "markets/invest_map/data/closes_20261020_20261113.csv"
    j["_readme"] = j.get("_readme", "") + "；v0.3.8 原地換價至 2026-11-13（px1113_peers.py；工作單 px1113）"
    json.dump(j, open(os.path.join(D, dst), "w", encoding="utf-8"), ensure_ascii=False, indent=1)
    print(dst, old_med, "->", j["median"]["value"], [(p["ticker"], p["ntmEvEbitda1007"], p["ntmEvEbitda"]) for p in j["peers"]])
