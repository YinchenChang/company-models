# SpaceX v0.2.10（工作單 px1113，2026-11-16）：同業 EV／NTM EBITDA 由 2026-11-06 收盤換到 2026-11-13 收盤。
# 方法（工作單第 3 節）：本檔沒有同業市值，故以「倍數 × （11-13 收盤 ÷ 11-06 收盤）」近似（EV 與倍數同比例縮放；淨負債不變的假設會使變動被高估），標 [Derived] 近似。
# 收盤價只取 markets/invest_map/data/closes_20261020_20261113.csv。EBITDA、營收、淨負債等基本面欄位與 NTM 權重不變（仍為 2026-11-06 MarketScreener 原值）。
# 用法：cd spacex && python3 peers_v0210.py   → data/peers_connectivity_20261113.json、data/peers_si_20261113.json（讀 *_20261106.json）
import csv, json, os, statistics
HERE = os.path.dirname(os.path.abspath(__file__))
CSV = os.path.join(HERE, "..", "markets", "invest_map", "data", "closes_20261020_20261113.csv")
C = {}
for r in csv.DictReader(open(CSV, encoding="utf-8")):
    C[(r["symbol"], r["date"])] = float(r["close"])
OLD, NEW = "2026-11-06", "2026-11-13"
def conv(src, dst):
    d = json.load(open(os.path.join(HERE, "data", src), encoding="utf-8"))
    for p in d["peers"]:
        t = p["ticker"]; a, b = C[(t, OLD)], C[(t, NEW)]
        k = b / a
        p["ntmEvEbitda1106"] = p["ntmEvEbitda"]; p["evCurrent1106"] = p["evCurrent"]; p["refPrice1106"] = p["refPrice"]
        p["priceRatio"] = round(k, 6); p["closeOld"] = a; p["closeNew"] = b
        p["evCurrent"] = round(p["evCurrent"] * k, 0)
        p["ntmEvEbitda"] = round(p["ntmEvEbitda"] * k, 3)
        p["refPrice"] = b
        p["note"] = ((p.get("note") or "") + ("；" if p.get("note") else "") + f"2026-11-13 換價（工作單 px1113）：倍數 {p['ntmEvEbitda1106']} × 收盤比 {b}/{a}（近似，無市值；淨負債不變）").strip()
    d["median"] = round(statistics.median(p["ntmEvEbitda"] for p in d["peers"]), 3)
    d["median1106"] = json.load(open(os.path.join(HERE, "data", src), encoding="utf-8"))["median"]
    d["asOf"] = "2026-11-13（收盤；MarketScreener 2026-11-06 倍數 × Yahoo Finance 收盤比，近似）"
    d["asOfOriginal"] = "2026-11-06（收盤；頁面 Reference price）"
    d["priceSource"] = "markets/invest_map/data/closes_20261020_20261113.csv（Yahoo Finance 日線收盤）"
    d["dateGap"] = "模型現價日與同業報價日同為 2026-11-13；同業倍數由 2026-11-06 依收盤價比例換算（近似 [Derived]）"
    d["_readme"] = d.get("_readme", "") + "；v0.2.10 換價至 2026-11-13（peers_v0210.py）"
    json.dump(d, open(os.path.join(HERE, "data", dst), "w", encoding="utf-8"), ensure_ascii=False, indent=1)
    print(dst, d["median1106"], "->", d["median"], [(p["ticker"], p["ntmEvEbitda1106"], p["ntmEvEbitda"]) for p in d["peers"]])
conv("peers_connectivity_20261106.json", "peers_connectivity_20261113.json")
conv("peers_si_20261106.json", "peers_si_20261113.json")
