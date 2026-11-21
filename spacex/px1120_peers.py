# SpaceX v0.2.11（工作單 px1120，2026-11-21）：同業 EV／NTM EBITDA 由 2026-11-13 近似值換到 2026-11-20 收盤。
# 方法（工作單 px1113 第 3 節）：新倍數＝現行倍數（v0.2.10，11-13）× 11-20 收盤 ÷ 11-13 收盤；EV 同比例縮放。無市值，近似 [Derived]。
# 收盤價只取 markets/invest_map/data/closes_20261109_20261120.csv。基本面欄位與 NTM 權重不變。
# 用法：cd spacex && python3 px1120_peers.py → data/peers_connectivity_20261120.json、data/peers_si_20261120.json（讀 *_20261113.json）
import csv, json, os, statistics
HERE = os.path.dirname(os.path.abspath(__file__))
CSV = os.path.join(HERE, "..", "markets", "invest_map", "data", "closes_20261109_20261120.csv")
C = {(r["symbol"], r["date"]): float(r["close"]) for r in csv.DictReader(open(CSV, encoding="utf-8"))}
OLD, NEW = "2026-11-13", "2026-11-20"
def conv(src, dst):
    s = json.load(open(os.path.join(HERE, "data", src), encoding="utf-8"))
    d = json.load(open(os.path.join(HERE, "data", src), encoding="utf-8"))
    for p in d["peers"]:
        t = p["ticker"]; a, b = C[(t, OLD)], C[(t, NEW)]
        k = b / a
        p["ntmEvEbitda1113"] = p["ntmEvEbitda"]; p["evCurrent1113"] = p["evCurrent"]
        p["priceRatio1120"] = round(k, 6)
        p["closeOld"] = a; p["closeNew"] = b
        p["evCurrent"] = round(p["evCurrent"] * k, 0)
        p["ntmEvEbitda"] = round(p["ntmEvEbitda"] * k, 3)
        p["refPrice"] = b
        p["note"] = ((p.get("note") or "") + "；" + f"2026-11-20 換價（工作單 px1120）：倍數 {p['ntmEvEbitda1113']} × 收盤比 {b}/{a}（近似，無市值；淨負債不變）").strip("；")
    d["median"] = round(statistics.median(p["ntmEvEbitda"] for p in d["peers"]), 3)
    d["median1113"] = s["median"]
    d["asOf"] = "2026-11-20（收盤；MarketScreener 2026-11-06 倍數 × Yahoo Finance 收盤比，近似；11-13 值再乘 11-20÷11-13）"
    d["priceSource"] = "markets/invest_map/data/closes_20261109_20261120.csv（Yahoo Finance 日線收盤）"
    d["dateGap"] = "模型現價日與同業報價日同為 2026-11-20；同業倍數由 2026-11-06 原值依收盤價比例換算（近似 [Derived]）"
    d["_readme"] = d.get("_readme", "") + "；v0.2.11 換價至 2026-11-20（px1120_peers.py）"
    json.dump(d, open(os.path.join(HERE, "data", dst), "w", encoding="utf-8"), ensure_ascii=False, indent=1)
    print(dst, d["median1113"], "->", d["median"], [(p["ticker"], p["ntmEvEbitda1113"], p["ntmEvEbitda"]) for p in d["peers"]])
conv("peers_connectivity_20261113.json", "peers_connectivity_20261120.json")
conv("peers_si_20261113.json", "peers_si_20261120.json")
