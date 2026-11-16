# Microsoft v0.3.9（原地修改，工作單 px1113 第 3 節）：同業 EV／NTM EBITDA（TTM）由 2026-10-07／10-08 報價日換到 2026-11-13 收盤。
# 方法：倍數新 ＝ 倍數舊 × 11-13 收盤 ÷ 報價日收盤（同業檔無市值，EV 與倍數同比例縮放；淨負債不變的假設使變動被高估），標 [Derived] 近似。
#       例外：MSFT 自身有市值（事實總帳 val.mktcap 3,930,000 M、val.ev 3,990,000 M）→ EV 新＝EV 舊＋市值舊×（價格比−1）；EBITDA 不變。
# 收盤價只取 markets/invest_map/data/closes_20261020_20261113.csv（含補入的 09-28～10-19 同業日）。EBITDA、NTM 權重等基本面不變。
# 用法：python3 microsoft/px1113_peers.py → data/peers_cloud_20261113.json、peers_ads_20261113.json、peers_px1113_20261113.json（build_microsoft.py 讀）
import csv, json, os, statistics
HERE = os.path.dirname(os.path.abspath(__file__))
D = os.path.join(HERE, "data")
C = {}
for r in csv.DictReader(open(os.path.join(HERE, "..", "markets", "invest_map", "data", "closes_20261020_20261113.csv"), encoding="utf-8")):
    C[(r["symbol"], r["date"])] = float(r["close"])
NEW = "2026-11-13"
def px(t, d): return C[(t, d)]
def conv(t, old, d, mcap_ev=None):
    a, b = px(t, d), px(t, NEW); k = b / a
    if mcap_ev:
        ev, mc = mcap_ev; f = (ev + mc * (k - 1)) / ev
    else:
        f = k
    return round(old * f, 2), a, b, k, f
def med(v): return round(statistics.median(v), 2)

out = {}
for src, dst in (("peers_cloud_20261008.json", "peers_cloud_20261113.json"), ("peers_ads_20261009.json", "peers_ads_20261113.json")):
    j = json.load(open(os.path.join(D, src), encoding="utf-8"))
    old_med = j["median"]["value"]
    for p in j["peers"]:
        d = p["ev"]["asOf"]
        new, a, b, k, f = conv(p["ticker"], p["ntmEvEbitda"], d)
        p["ntmEvEbitda1007"] = p["ntmEvEbitda"]; p["ev1007"] = p["ev"]["value"]; p["refDate"] = d
        p["closeOld"] = a; p["closeNew"] = b; p["priceRatio"] = round(k, 6)
        p["ntmEvEbitda"] = new; p["ev"]["value"] = round(p["ev"]["value"] * f, 2)
        p["ev"]["asOf"] = NEW; p["ev"]["note"] = f"2026-11-13 換價（工作單 px1113）：EV 舊 {p['ev1007']}（{d}）× 收盤比 {b}/{a}；近似，無市值，淨負債不變"
    j["median1007"] = j["median"]
    j["median"] = {"value": med(p["ntmEvEbitda"] for p in j["peers"]), "n": len(j["peers"]), "tag": "[Derived]（近似）", "note": f"{len(j['peers'])} 家中位數；2026-11-13 換價（舊 {old_med}）"}
    j["asOf"] = "2026-11-13（倍數＝原報價日倍數 × Yahoo Finance 收盤比，近似）"; j["asOfOriginal"] = "2026-10-07／10-08"
    j["priceSource"] = "markets/invest_map/data/closes_20261020_20261113.csv"
    j["_readme"] = j.get("_readme", "") + "；v0.3.9 原地換價至 2026-11-13（px1113_peers.py；工作單 px1113）"
    json.dump(j, open(os.path.join(D, dst), "w", encoding="utf-8"), ensure_ascii=False, indent=1)
    out[dst] = (old_med, j["median"]["value"], [(p["ticker"], p["ntmEvEbitda1007"], p["ntmEvEbitda"]) for p in j["peers"]])

F = json.load(open(os.path.join(D, "microsoft_facts_20261008.json"), encoding="utf-8"))
FV = {i["id"]: i for g in F["groups"].values() for i in g}
ITEMS = {}
for id_, t, d in (("val.peer.sw.orcl", "ORCL", "2026-10-07"), ("val.peer.sw.sap", "SAP", "2026-10-07"), ("val.peer.sw.crm", "CRM", "2026-10-07"),
                  ("val.peer.sw.adbe", "ADBE", "2026-10-07"), ("val.peer.sw.now", "NOW", "2026-10-07"),
                  ("val.peer.game.ttwo", "TTWO", "2026-10-08"), ("val.peer.game.ntdoy", "NTDOY", "2026-10-08"), ("val.peer.game.ntes", "NTES", "2026-10-08"),
                  ("val.peer.pc.hpq", "HPQ", "2026-10-08"), ("val.peer.pc.dell", "DELL", "2026-10-08"), ("val.peer.pc.0992.hk", "0992.HK", "2026-10-08")):
    new, a, b, k, f = conv(t, FV[id_]["value"], d)
    ITEMS[id_] = dict(ticker=t, old=FV[id_]["value"], new=new, refDate=d, closeOld=a, closeNew=b, priceRatio=round(k, 6), method="倍數 × 收盤比（近似，無市值）")
new, a, b, k, f = conv("MSFT", FV["val.msft.ownNtm"]["value"], "2026-10-07", (3990000.0, 3930000.0))
ITEMS["val.msft.ownNtm"] = dict(ticker="MSFT", old=FV["val.msft.ownNtm"]["value"], new=new, refDate="2026-10-07", closeOld=a, closeNew=b, priceRatio=round(k, 6),
                                method="EV 新＝EV 舊 3,990,000＋市值舊 3,930,000 ×（價格比−1）；EBITDA 不變（事實總帳 val.ev、val.mktcap）")
MED = {"sw": med(ITEMS[f"val.peer.sw.{t}"]["new"] for t in ("orcl", "sap", "crm", "adbe", "now")),
       "game": med(ITEMS[f"val.peer.game.{t}"]["new"] for t in ("ttwo", "ntdoy", "ntes")),
       "pc": med(ITEMS[f"val.peer.pc.{t}"]["new"] for t in ("hpq", "dell", "0992.hk"))}
OLDMED = {"sw": FV["val.peer.sw.median"]["value"], "game": FV["val.peer.game.median"]["value"], "pc": FV["val.peer.pc.median"]["value"]}
json.dump(dict(asOf=NEW, priceSource="markets/invest_map/data/closes_20261020_20261113.csv", tag="[Derived]（近似）", items=ITEMS, medians=MED, mediansOld=OLDMED,
               cloudMedian=out["peers_cloud_20261113.json"][1], adsMedian=out["peers_ads_20261113.json"][1]),
          open(os.path.join(D, "peers_px1113_20261113.json"), "w", encoding="utf-8"), ensure_ascii=False, indent=1)
for k_, v in out.items(): print(k_, v)
print("sw/game/pc", OLDMED, "->", MED, {i: (v["old"], v["new"]) for i, v in ITEMS.items()})
