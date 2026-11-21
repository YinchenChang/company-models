# Alphabet v0.3.11（工作單 px1120）：共用同業檔倍數由 2026-11-13 換到 2026-11-20 收盤。
# 方法：倍數新 ＝ 現行倍數（已是 11-13）× 11-20 收盤 ÷ 11-13 收盤（同業檔無市值，近似，標 [Derived]）。EBITDA、NTM 權重不變。
# 價格只取 markets/invest_map/data/closes_20261109_20261120.csv。
# 用法：python3 alphabet/px1120_peers.py → data/peers_{cloud,ads,subs,own}_20261120.json（build 讀）
import csv, json, os, statistics
HERE = os.path.dirname(os.path.abspath(__file__)); D = os.path.join(HERE, "data")
C = {}
for r in csv.DictReader(open(os.path.join(HERE, "..", "markets", "invest_map", "data", "closes_20261109_20261120.csv"), encoding="utf-8")):
    C[(r["symbol"], r["date"])] = float(r["close"])
OLD, NEW = "2026-11-13", "2026-11-20"
SRC = "markets/invest_map/data/closes_20261109_20261120.csv"
def med(v): return round(statistics.median(v), 2)
UNCH = []
def reprice(p):
    t = p["ticker"]
    if (t, OLD) not in C or (t, NEW) not in C:
        UNCH.append(t); return False
    a, b = C[(t, OLD)], C[(t, NEW)]; k = b / a
    p["ntmEvEbitda1113"] = p["ntmEvEbitda"]; p["close1113"] = a; p["close1120"] = b; p["priceRatio1120"] = round(k, 6)
    p["ntmEvEbitda"] = round(p["ntmEvEbitda"] * k, 2)
    return True
for grp in ("cloud", "ads", "subs"):
    j = json.load(open(os.path.join(D, f"peers_{grp}_20261113.json"), encoding="utf-8"))
    m1113 = j["median"]["value"]
    for p in j["peers"]:
        if reprice(p) and "ev" in p:
            p["ev1113"] = p["ev"]["value"]; p["ev"]["value"] = round(p["ev"]["value"] * p["priceRatio1120"], 2)
            p["ev"]["asOf"] = NEW; p["ev"]["note"] = f"2026-11-20 換價（工作單 px1120）：EV 11-13 值 {p['ev1113']} × 收盤比 {p['close1120']}/{p['close1113'] if 'close1113' in p else ''}；近似，無市值，淨負債不變"
    j["median1113"] = j["median"]
    j["median"] = {"value": med(p["ntmEvEbitda"] for p in j["peers"]), "n": len(j["peers"]), "tag": "[Derived]（近似）", "note": f"{len(j['peers'])} 家中位數；2026-11-20 換價（11-13 為 {m1113}）"}
    j["asOf"] = "2026-11-20（倍數＝11-13 倍數 × Yahoo Finance 收盤比，近似；原報價日 2026-10-07／10-08）"
    j["priceSource"] = SRC
    json.dump(j, open(os.path.join(D, f"peers_{grp}_20261120.json"), "w", encoding="utf-8"), ensure_ascii=False, indent=1)
    print(grp, m1113, "->", j["median"]["value"], [(p["ticker"], p["ntmEvEbitda1113"], p["ntmEvEbitda"]) for p in j["peers"]])
# Alphabet 自身：EV 新＝EV(11-13)＋市值(11-13) ×（11-20 ÷ 11-13 − 1）；EBITDA 不變。
o = json.load(open(os.path.join(D, "peers_own_20261113.json"), encoding="utf-8"))
a, b = C[("GOOGL", OLD)], C[("GOOGL", NEW)]
ev0, mc0 = o["evNew"], 12294e-3 * a
ev1 = ev0 + mc0 * (b / a - 1)
own = dict(old=o["old"], prev1113=o["new"], new=round(o["new"] * ev1 / ev0, 2), refDate=OLD, close1113=a, close1120=b, priceRatio=round(b / a, 6),
           ev1113=ev0, mcap1113=round(mc0, 1), evNew=round(ev1, 1),
           method=f"EV 新＝EV 11-13 {ev0}＋市值 11-13 {mc0:,.1f}（{a} × 12,294M 股）×（{b} ÷ {a} − 1）＝{ev1:,.1f}；EBITDA 不變；11-13 值 {o['new']}（10-07 為 {o['old']}）", tag="[Derived]")
json.dump(own, open(os.path.join(D, "peers_own_20261120.json"), "w", encoding="utf-8"), ensure_ascii=False, indent=1)
print("own", own)
print("未換價:", UNCH or "無")
