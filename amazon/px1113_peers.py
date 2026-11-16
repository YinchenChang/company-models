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

# 補做（N 批）：零售組 WMT、COST、TGT（價格檔 09-28～11-13 已補）。原值見事實總帳 val.peer.retail.*（報價日 2026-10-07，EV 與 NTM EBITDA 來自不同供應商）。
# 事實總帳只有 EV 與倍數，無市值 → 倍數新 ＝ 倍數舊 × 11-13 收盤 ÷ 10-07 收盤（近似，標 [Derived]）。
RET = [("WMT", "Walmart", 17.72, 921.68), ("COST", "Costco", 25.33, 406.42), ("TGT", "Target", 8.77, 82.35)]
peers = []
for t, nm, m0, ev0 in RET:
    a, b = C[(t, "2026-10-07")], C[(t, NEW)]; k = b / a
    peers.append(dict(ticker=t, name=nm, ntmEvEbitda1007=m0, ntmEvEbitda=round(m0 * k, 2), ev1007=ev0, ev=dict(value=round(ev0 * k, 2), asOf=NEW),
                      refDate="2026-10-07", closeOld=a, closeNew=b, priceRatio=round(k, 6)))
jr = dict(group="retail", peers=peers, median1007=dict(value=med(p["ntmEvEbitda1007"] for p in peers)),
          median=dict(value=med(p["ntmEvEbitda"] for p in peers), n=3, tag="[Derived]（近似）"), asOf="2026-11-13（倍數＝2026-10-07 倍數 × Yahoo Finance 收盤比，近似）",
          priceSource="markets/invest_map/data/closes_20261020_20261113.csv")
json.dump(jr, open(os.path.join(D, "peers_retail_20261113.json"), "w", encoding="utf-8"), ensure_ascii=False, indent=1)
print("retail", jr["median1007"]["value"], "->", jr["median"]["value"], [(p["ticker"], p["ntmEvEbitda1007"], p["ntmEvEbitda"]) for p in peers])
