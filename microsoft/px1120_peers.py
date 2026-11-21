# Microsoft v0.3.10（原地修改，工作單 px1120 第 3 節）：同業 EV／NTM EBITDA 由 2026-11-13 換到 2026-11-20 收盤。
# 方法：倍數新 ＝ 倍數現行（11-13）× 11-20 收盤 ÷ 11-13 收盤（同業檔無市值，EV 與倍數同比例縮放，近似，標 [Derived]）。
#       例外：MSFT 自身有市值 → 11-13 市值 ＝ 3,930,000 × (523.88 ÷ 529.76)、EV ＝ 3,990,000 ＋（市值 − 3,930,000）；EV 新＝EV＋市值×（價格比−1）；EBITDA 不變。
# 收盤價只取 markets/invest_map/data/closes_20261109_20261120.csv（含 11-13 與 11-20）。
# 用法：python3 microsoft/px1120_peers.py → data/peers_cloud_20261120.json、peers_ads_20261120.json、peers_px1120_20261120.json
import csv, json, os, statistics
HERE = os.path.dirname(os.path.abspath(__file__))
D = os.path.join(HERE, "data")
C = {}
for r in csv.DictReader(open(os.path.join(HERE, "..", "markets", "invest_map", "data", "closes_20261109_20261120.csv"), encoding="utf-8")):
    C[(r["symbol"], r["date"])] = float(r["close"])
OLD, NEW = "2026-11-13", "2026-11-20"
def med(v): return round(statistics.median(v), 2)
def ratio(t): return C[(t, NEW)] / C[(t, OLD)]

for src, dst in (("peers_cloud_20261113.json", "peers_cloud_20261120.json"), ("peers_ads_20261113.json", "peers_ads_20261120.json")):
    j = json.load(open(os.path.join(D, src), encoding="utf-8"))
    old_med = j["median"]["value"]
    for p in j["peers"]:
        t = p["ticker"]; a, b = C[(t, OLD)], C[(t, NEW)]; k = b / a
        p["ntmEvEbitda1113"] = p["ntmEvEbitda"]; p["ev1113"] = p["ev"]["value"]
        p["closeOld"] = a; p["closeNew"] = b; p["priceRatio"] = round(k, 6); p["refDate"] = OLD
        p["ntmEvEbitda"] = round(p["ntmEvEbitda"] * k, 2); p["ev"]["value"] = round(p["ev"]["value"] * k, 2)
        p["ev"]["asOf"] = NEW; p["ev"]["note"] = f"2026-11-20 換價（工作單 px1120）：EV 舊 {p['ev1113']}（{OLD}）× 收盤比 {b}/{a}；近似，無市值，淨負債不變"
    j["median1113"] = j["median"]
    j["median"] = {"value": med(p["ntmEvEbitda"] for p in j["peers"]), "n": len(j["peers"]), "tag": "[Derived]（近似）", "note": f"{len(j['peers'])} 家中位數；2026-11-20 換價（舊 {old_med}）"}
    j["asOf"] = "2026-11-20（倍數＝11-13 倍數 × Yahoo Finance 收盤比，近似）"
    j["priceSource"] = "markets/invest_map/data/closes_20261109_20261120.csv"
    j["_readme"] = j.get("_readme", "") + "；v0.3.10 原地換價至 2026-11-20（px1120_peers.py；工作單 px1120）"
    json.dump(j, open(os.path.join(D, dst), "w", encoding="utf-8"), ensure_ascii=False, indent=1)

P = json.load(open(os.path.join(D, "peers_px1113_20261113.json"), encoding="utf-8"))
ITEMS = {}
for id_, v in P["items"].items():
    t = v["ticker"]; a, b = C[(t, OLD)], C[(t, NEW)]; k = b / a
    if id_ == "val.msft.ownNtm":
        k13 = v["priceRatio"]  # 11-13 ÷ 10-07
        mc = 3930000.0 * k13; ev = 3990000.0 + (mc - 3930000.0); f = (ev + mc * (k - 1)) / ev
        method = f"EV 新＝EV（11-13：{ev:,.0f}）＋市值（11-13：{mc:,.0f}）×（價格比−1）；EBITDA 不變（事實總帳 val.ev、val.mktcap）"
    else:
        f = k; method = "倍數 × 收盤比（近似，無市值）"
    ITEMS[id_] = dict(ticker=t, old=v["new"], new=round(v["new"] * f, 2), refDate=OLD, closeOld=a, closeNew=b, priceRatio=round(k, 6), method=method)
MED = {"sw": med(ITEMS[f"val.peer.sw.{t}"]["new"] for t in ("orcl", "sap", "crm", "adbe", "now")),
       "game": med(ITEMS[f"val.peer.game.{t}"]["new"] for t in ("ttwo", "ntdoy", "ntes")),
       "pc": med(ITEMS[f"val.peer.pc.{t}"]["new"] for t in ("hpq", "dell", "0992.hk"))}
cl = json.load(open(os.path.join(D, "peers_cloud_20261120.json"), encoding="utf-8")); ad = json.load(open(os.path.join(D, "peers_ads_20261120.json"), encoding="utf-8"))
json.dump(dict(asOf=NEW, priceSource="markets/invest_map/data/closes_20261109_20261120.csv", tag="[Derived]（近似）", items=ITEMS, medians=MED, mediansOld=P["medians"],
               cloudMedian=cl["median"]["value"], adsMedian=ad["median"]["value"]),
          open(os.path.join(D, "peers_px1120_20261120.json"), "w", encoding="utf-8"), ensure_ascii=False, indent=1)
print("cloud", cl["median1113"]["value"], "->", cl["median"]["value"], "ads", ad["median1113"]["value"], "->", ad["median"]["value"])
print(P["medians"], "->", MED, {i: (v["old"], v["new"]) for i, v in ITEMS.items()})
