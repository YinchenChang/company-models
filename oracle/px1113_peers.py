# Oracle v0.5.3（原地修改，工作單 px1113 第 3 節）：軟體同業（MSFT、SAP、CRM、ADBE、NOW）NTM EV/EBITDA 由報價日 2026-10-07 換到 2026-11-13 收盤。
# 方法：SAP、CRM、ADBE、NOW 無市值 → 倍數 × 收盤比近似（EV 與倍數同比例縮放；淨負債不變的假設使變動被高估），標 [Derived] 近似。
#       MSFT 有市值（microsoft/data/microsoft_facts_20261008.json val.mktcap 3,930,000 M、val.ev 3,990,000 M，2026-10-07）→ EV 新＝EV 舊＋市值舊 ×（價格比−1）；EBITDA 不變。
# 收盤價只取 markets/invest_map/data/closes_20261020_20261113.csv（含補入的 09-28～10-19）。EBITDA、NTM 權重等基本面不變。
# 用法：python3 oracle/px1113_peers.py → oracle/data/peers_sw_20261113.json（build_oracle.py、scripts/sens_v04.py 讀）
import csv, json, os, statistics
HERE = os.path.dirname(os.path.abspath(__file__))
D = os.path.join(HERE, "data")
C = {}
for r in csv.DictReader(open(os.path.join(HERE, "..", "markets", "invest_map", "data", "closes_20261020_20261113.csv"), encoding="utf-8")):
    C[(r["symbol"], r["date"])] = float(r["close"])
NEW, OLD = "2026-11-13", "2026-10-07"
F = json.load(open(os.path.join(D, "oracle_facts_20261007.json"), encoding="utf-8"))
FV = {i["id"]: i for g in F["groups"].values() for i in g}
out = {}
for t in ("MSFT", "SAP", "CRM", "ADBE", "NOW"):
    old = FV[f"val.peer.{t.lower()}"]["value"]; a, b = C[(t, OLD)], C[(t, NEW)]; k = b / a
    if t == "MSFT":
        EV, MC = 3990000.0, 3930000.0; f = (EV + MC * (k - 1)) / EV; how = "EV 新＝EV 舊 3,990,000＋市值舊 3,930,000 ×（價格比−1）；EBITDA 不變（MSFT 事實總帳 val.ev、val.mktcap）"
    else:
        f = k; how = "倍數 × 收盤比（近似，無市值；淨負債不變）"
    out[t] = dict(old=old, new=round(old * f, 2), refDate=OLD, closeOld=a, closeNew=b, priceRatio=round(k, 6), method=how, tag="[Derived]" if t == "MSFT" else "[Derived]（近似）")
vals = sorted(v["new"] for v in out.values()); old_vals = sorted(v["old"] for v in out.values())
res = dict(asOf=NEW, priceSource="markets/invest_map/data/closes_20261020_20261113.csv", peers=out,
           median=round(statistics.median(vals), 2), medianOld=round(statistics.median(old_vals), 2),
           lo=out["CRM"]["new"], hi=out["MSFT"]["new"], loOld=out["CRM"]["old"], hiOld=out["MSFT"]["old"],     # Sens 低／高＝第二低（CRM）／第二高（MSFT），與原 10.8／15.59 同位置
           legacyMedian=round(statistics.median([out["CRM"]["new"], out["ADBE"]["new"]]), 2), legacyMedianOld=round(statistics.median([out["CRM"]["old"], out["ADBE"]["old"]]), 2))
json.dump(res, open(os.path.join(D, "peers_sw_20261113.json"), "w", encoding="utf-8"), ensure_ascii=False, indent=1)
print({t: (v["old"], v["new"]) for t, v in out.items()}, "median", res["medianOld"], "->", res["median"], "lo/hi", res["lo"], res["hi"], "legacy", res["legacyMedianOld"], res["legacyMedian"])
