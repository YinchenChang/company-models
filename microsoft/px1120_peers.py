# Microsoft v0.3.10（原地修改，工作單 px1120 第 3 節）：同業 EV／NTM EBITDA 由 2026-11-13 換到 2026-11-20 收盤。
# 方法：倍數新 ＝ 倍數現行（11-13）× 11-20 收盤 ÷ 11-13 收盤（同業檔無市值，EV 與倍數同比例縮放，近似，標 [Derived]）。
#       例外：MSFT 自身有市值 → 11-13 市值 ＝ 3,930,000 × (523.88 ÷ 529.76)、EV ＝ 3,990,000 ＋（市值 − 3,930,000）；EV 新＝EV＋市值×（價格比−1）；EBITDA 不變。
# 收盤價只取 markets/invest_map/data/closes_20261109_20261120.csv（含 11-13 與 11-20）。
# 用法：python3 microsoft/px1120_peers.py → data/peers_cloud_20261120.json、peers_ads_20261120.json、peers_px1120_20261120.json
# 外部稽核優化 W3-7（2026-11-23；Andy 決定 J6 取預設 (a)）：企業軟體同業中有市值資料的兩家（ORCL、CRM）換價改為「分離淨負債」，由 10-07 原值直接換到 11-20：
#   EV 新＝EV（10-07）＋市值（10-07）×（11-20 收盤 ÷ 10-07 收盤 − 1）；淨負債與 EBITDA 不變。原式（倍數 × 收盤比）等於假設淨負債跟著股價等比例變動。
#   ORCL：EV 566,120、市值 434,050（oracle/data/oracle_facts_20261007.json val.ev.orcl 與其 note；StockAnalysis，2026-10-07，同一來源同一日）。
#   CRM：EV 214,520（同檔 val.peer.crm note）；市值＝流通股數 823.0 百萬股（StockAnalysis，頁面日期 2026-11-15；software_ai/salesforce/data/facts_salesforce.csv）× 10-07 收盤。
#   其餘（SAP、ADBE、NOW、遊戲 3 家、PC 3 家，以及共用的雲端 8 家、廣告 6 家同業檔）：repo 內沒有報價日的市值或淨負債 → 仍為整體縮放（近似）。
#   原式結果保留在各項的 newScaled 與 mediansScaled 供對照。
import csv, json, os, statistics
SEP = {"val.peer.sw.orcl": (566120.0, 434050.0, "EV 566,120、市值 434,050（StockAnalysis 2026-10-07；oracle_facts_20261007.json val.ev.orcl）"),
       "val.peer.sw.crm": (214520.0, None, "EV 214,520（oracle_facts_20261007.json val.peer.crm）、市值＝823.0 百萬股（StockAnalysis，頁面日期 2026-11-15）× 10-07 收盤")}
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
    elif id_ in SEP:      # W3-7：分離淨負債
        ev07, mc07, src = SEP[id_]; c07 = v["closeOld"]
        if mc07 is None: mc07 = 823.0 * c07
        ev_new = ev07 + mc07 * (b / c07 - 1); f = v["old"] * ev_new / ev07 / v["new"]
        method = f"分離淨負債（W3-7）：EV 新＝{ev07:,.0f}＋市值（10-07）{mc07:,.0f} ×（{b} ÷ {c07} − 1）＝{ev_new:,.0f}；淨負債 {ev07 - mc07:,.0f} 與 EBITDA 不變；倍數＝10-07 原值 {v['old']} × EV 新 ÷ EV 舊；{src}"
    else:
        f = k; method = "倍數 × 收盤比——仍為整體縮放（近似）：repo 內沒有市值或淨負債"
    ITEMS[id_] = dict(ticker=t, old=v["new"], new=round(v["new"] * f, 2), newScaled=round(v["new"] * (f if id_ == "val.msft.ownNtm" else k), 2), separated=(id_ in SEP or id_ == "val.msft.ownNtm"),
                      refDate=OLD, closeOld=a, closeNew=b, priceRatio=round(k, 6), method=method)
MED = {"sw": med(ITEMS[f"val.peer.sw.{t}"]["new"] for t in ("orcl", "sap", "crm", "adbe", "now")),
       "game": med(ITEMS[f"val.peer.game.{t}"]["new"] for t in ("ttwo", "ntdoy", "ntes")),
       "pc": med(ITEMS[f"val.peer.pc.{t}"]["new"] for t in ("hpq", "dell", "0992.hk"))}
MEDS = {"sw": med(ITEMS[f"val.peer.sw.{t}"]["newScaled"] for t in ("orcl", "sap", "crm", "adbe", "now")), "game": MED["game"], "pc": MED["pc"]}      # W3-7：原式結果對照
cl = json.load(open(os.path.join(D, "peers_cloud_20261120.json"), encoding="utf-8")); ad = json.load(open(os.path.join(D, "peers_ads_20261120.json"), encoding="utf-8"))
json.dump(dict(asOf=NEW, priceSource="markets/invest_map/data/closes_20261109_20261120.csv", tag="[Derived]（近似）", items=ITEMS, medians=MED, mediansScaled=MEDS, mediansOld=P["medians"],
               cloudMedian=cl["median"]["value"], adsMedian=ad["median"]["value"]),
          open(os.path.join(D, "peers_px1120_20261120.json"), "w", encoding="utf-8"), ensure_ascii=False, indent=1)
print("cloud", cl["median1113"]["value"], "->", cl["median"]["value"], "ads", ad["median1113"]["value"], "->", ad["median"]["value"])
print(P["medians"], "->", MED, {i: (v["old"], v["new"]) for i, v in ITEMS.items()})
