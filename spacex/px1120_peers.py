# SpaceX v0.2.11（工作單 px1120，2026-11-21）：同業 EV／NTM EBITDA 由 2026-11-13 近似值換到 2026-11-20 收盤。
# 方法（工作單 px1113 第 3 節）：新倍數＝現行倍數（v0.2.10，11-13）× 11-20 收盤 ÷ 11-13 收盤；EV 同比例縮放。無市值，近似 [Derived]。
# 收盤價只取 markets/invest_map/data/closes_20261109_20261120.csv。基本面欄位與 NTM 權重不變。
# 用法：cd spacex && python3 px1120_peers.py → data/peers_connectivity_20261120.json、data/peers_si_20261120.json（讀 *_20261113.json）
# 外部稽核優化 W3-7（2026-11-23；Andy 決定 J6 取預設 (a)）：換價改為「分離淨負債」。
#   原式把整個 EV／EBITDA 乘收盤比，等於假設淨負債跟著股價等比例變動；淨負債為正的公司（電信）變動被高估，淨現金的公司被低估。
#   分離式：EV 新＝EV（2026-11-06 原值）＋市值（11-06）×（11-20 收盤 ÷ 11-06 頁面參考價 − 1）；淨負債與 NTM EBITDA 不變。直接由 11-06 原值換到 11-20，不經 11-13 近似值。
#   市值（11-06）的來源（只用 repo 內已有資料，不估計）：
#     A. Connectivity 五家：同一來源恆等式。MarketScreener 的當年度 EV＝現價市值＋當年度預估淨負債（data/peers_connectivity_20261106.json method 欄），
#        故市值＝evCurrent − netDebt[2026]（同檔）。
#     B. NBIS、ORCL：MarketScreener 股數（千股；擷取 2026-11-13，coreweave/build_coreweave.py 的 pn_sh 註記 274,099、po_sh 3,117,708）× 11-06 參考價；
#        反推淨負債 13,659、115,597 與同來源預估淨負債（coreweave In pn_ndf 13.659、po_ndf 115.597）相符（差 < 1），等同恆等式。
#     C. CRWV：MarketScreener 頁面股數 repo 內沒有；用 Q3 基本加權 559 百萬股（coreweave In sh_out，Verified）× 11-06 參考價，
#        反推淨負債 44,879（申報口徑 43,400、共識 FY2026 預估 46,442，落在兩者之間）。若改用 46,442 反推市值，倍數差 0.015 倍以內。
#   原式結果保留在每家的 ntmEvEbitdaScaled、evCurrentScaled 與檔案層級 medianScaled，供 Recon 第 7 節對照欄使用。
import csv, json, os, statistics
HERE = os.path.dirname(os.path.abspath(__file__))
CSV = os.path.join(HERE, "..", "markets", "invest_map", "data", "closes_20261109_20261120.csv")
C = {(r["symbol"], r["date"]): float(r["close"]) for r in csv.DictReader(open(CSV, encoding="utf-8"))}
OLD, NEW = "2026-11-13", "2026-11-20"
# W3-7：SI 同業的股數（千股、出處、證據等級）；Connectivity 五家用同檔 netDebt，不需股數
MS_SHARES = {"NBIS": (274099, "MarketScreener 股數 274,099 千股", "A"), "ORCL": (3117708, "MarketScreener 股數 3,117,708 千股", "A"), "CRWV": (559000, "Q3 基本加權 559,000 千股（coreweave In sh_out）", "B")}
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
    # ── W3-7：分離淨負債（由 11-06 原值直接換到 11-20；原式結果另存對照）──
    for p in d["peers"]:
        t = p["ticker"]
        p["ntmEvEbitdaScaled"] = p["ntmEvEbitda"]; p["evCurrentScaled"] = p["evCurrent"]
        ev0, px0 = p["evCurrent1106"], p["refPrice1106"]
        nd = (p.get("netDebt") or {}).get("2026")
        if nd is not None:
            mc0 = ev0 - nd; how = f"市值（11-06）＝EV {ev0:,.0f} − MarketScreener 2026 年預估淨負債 {nd:,.0f}＝{mc0:,.0f}（同一來源恆等式）"; grade = "A"
        elif t in MS_SHARES:
            mc0 = MS_SHARES[t][0] * px0 / 1000; nd = ev0 - mc0; grade = MS_SHARES[t][2]
            how = f"市值（11-06）＝{MS_SHARES[t][1]} × 參考價 {px0}＝{mc0:,.0f}；反推淨負債 {nd:,.0f}"
        else:
            p["sepMethod"] = "仍為整體縮放（近似）：repo 內沒有市值或淨負債"; continue
        k0 = p["closeNew"] / px0
        ev1 = ev0 + mc0 * (k0 - 1)
        p["mcap1106"] = round(mc0, 0); p["netDebtHeld"] = round(nd, 0); p["priceRatio1106to1120"] = round(k0, 6)
        p["evCurrent"] = round(ev1, 0); p["ntmEvEbitda"] = round(ev1 / p["ntmEbitda"], 3); p["sepGrade"] = grade
        p["sepMethod"] = f"分離淨負債（W3-7）：EV 新＝{ev0:,.0f}＋{mc0:,.0f} ×（{p['closeNew']} ÷ {px0} − 1）＝{ev1:,.0f}；{how}；淨負債與 NTM EBITDA 不變"
        base = (p.get("note") or "").split("2026-11-13 換價")[0].rstrip("；")      # 只留 11-06 原註；兩次「近似，無市值」換價註記由下句取代
        p["note"] = (base + "；" if base else "") + f"2026-11-20 換價（工作單 px1120；外部稽核優化 W3-7 改為分離淨負債，2026-11-23）：11-06 原值 {p['ntmEvEbitda1106']} → {p['ntmEvEbitda']}（原式整個倍數 × 收盤比為 {p['ntmEvEbitdaScaled']}）；{how}"
    d["medianScaled"] = round(statistics.median(p["ntmEvEbitdaScaled"] for p in d["peers"]), 3)
    d["median"] = round(statistics.median(p["ntmEvEbitda"] for p in d["peers"]), 3)
    d["median1113"] = s["median"]
    d["asOf"] = "2026-11-20（收盤；MarketScreener 2026-11-06 EV 分離淨負債後，市值部分依 Yahoo Finance 11-20 收盤 ÷ 11-06 參考價換算；淨負債與 NTM EBITDA 不變）"
    d["priceSource"] = "markets/invest_map/data/closes_20261109_20261120.csv（Yahoo Finance 日線收盤）"
    d["dateGap"] = "模型現價日與同業報價日同為 2026-11-20；同業 EV 由 2026-11-06 原值只換市值部分（淨負債不隨股價變；[Derived]）"
    d["_readme"] = d.get("_readme", "") + "；v0.2.11 換價至 2026-11-20（px1120_peers.py）；外部稽核優化 W3-7（2026-11-23）改為分離淨負債，原式結果存 ntmEvEbitdaScaled、medianScaled"
    json.dump(d, open(os.path.join(HERE, "data", dst), "w", encoding="utf-8"), ensure_ascii=False, indent=1)
    print(dst, "原式中位數", d["medianScaled"], "→ 分離式", d["median"], [(p["ticker"], p["ntmEvEbitdaScaled"], p["ntmEvEbitda"], p.get("sepGrade")) for p in d["peers"]])
conv("peers_connectivity_20261113.json", "peers_connectivity_20261120.json")
conv("peers_si_20261113.json", "peers_si_20261120.json")
