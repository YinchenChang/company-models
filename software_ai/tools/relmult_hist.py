#!/usr/bin/env python3
"""歷史相對倍數庫（PS 與 PSG）— 規格：software_ai/docs/plan/歷史相對倍數_PS與PSG.md（2026-11-17）

用法（在 software_ai/ 下）：
    python3 tools/relmult_hist.py            # raw -> CSV -> out/<公司>.json + summary.md
    python3 tools/relmult_hist.py --company snowflake

資料流：
  data/relmult_hist/raw/<代號>.txt   WebFetch 取得的 StockAnalysis 季度比率頁（季末日,市值,PS；# 開頭為檔頭）
  data/relmult_hist/<代號>.csv       每家公司一個 CSV（本程式由 raw 產生；市值、TTM 營收＝市值÷PS、TTM 營收年增、PS、PSG…）
  data/relmult_hist/out/<公司>.json  各模型讀取的結果
  data/relmult_hist/summary.md       一張彙總表

口徑（規格第 1、2 節）：
  PS  ＝ 市值 ÷ TTM 營收（資料商以季末市值與 TTM 營收計；TTM 營收由 市值÷PS 反推，PS 取兩位小數，誤差 <0.5%）
  PSG ＝ PS ÷ (TTM 營收年增 × 100)；年增 < 2% 或為負記 N/M
  相對值＝本公司 ÷ 同業中位數；同業該季有值者少於 2 家記 N/M
  PS 讀數＝相對 PS 歷史中位數
  PSG 讀數＝相對 PSG 歷史中位數 × 現行（本公司 ÷ 同業中位）TTM 營收成長比（NTM 未取得，退 TTM）
  基準＝兩讀數幾何平均；Sens 低高＝兩條序列各取最低／最高再幾何平均；另列只用 PS、只用 PSG
  差異旗標：|PSG 讀數 ÷ PS 讀數 − 1| > 30%
"""
import argparse, csv, datetime as dt, json, math, os, statistics as st, sys

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)                       # software_ai/
DATA = os.path.join(ROOT, "data", "relmult_hist")
RAW, OUT = os.path.join(DATA, "raw"), os.path.join(DATA, "out")
NM_G = 0.02          # TTM 年增低於此值記 N/M
MIN_PEERS = 2
WINDOW = 20          # 最近 5 年 × 4 季
JUMP = 0.25          # 隱含 TTM 營收單季跳動超過此值標記（重編、併購或資料錯誤）

# 公司 → (資料代號, 同業群組 {群組: [代號]})；同業取自各模型 build_*.py／facts／README 的 Comparable 名單
COMPANIES = {
    "accenture":  ("ACN",       {"all": ["CTSH", "INFY", "GIB"]}),
    "fujitsu":    ("FUJITSU",   {"all": ["HITACHI", "NEC", "NRI", "BIPROGY"]}),
    "nri":        ("NRI",       {"all": ["TIS", "BIPROGY", "FUJITSU", "NSSOL", "OBIC", "ACN", "CTSH", "INFY"]}),
    "salesforce": ("CRM",       {"all": ["NOW", "ADBE", "WDAY", "INTU", "HUBS", "SAP"]}),
    "hubspot":    ("HUBS",      {"all": ["CRM", "NOW", "WDAY", "ADBE", "INTU", "TEAM", "MNDY", "DOCU", "TWLO", "ZM", "DDOG"]}),
    "lycorp":     ("LYCORP",    {"all": ["CYBERAGENT", "RECRUIT", "NAVER", "META", "RAKUTEN", "ZOZO", "MERCARI", "MONOTARO", "PYPL", "XYZ"],
                                 "Media": ["CYBERAGENT", "RECRUIT", "NAVER", "META"],
                                 "Commerce": ["RAKUTEN", "ZOZO", "MERCARI", "MONOTARO"],
                                 "Strategic": ["RAKUTEN", "PYPL", "XYZ"]}),
    "naver":      ("NAVER",     {"all": ["GOOGL", "META", "PINS", "KAKAO", "LYCORP"]}),
    "applovin":   ("APP",       {"all": ["PINS", "TTD", "RDDT", "U", "META", "GOOGL"]}),
    "snowflake":  ("SNOW",      {"all": ["DDOG", "MDB", "NET", "CRWD", "NOW", "PLTR", "ESTC", "ZS"]}),
    "datadog":    ("DDOG",      {"all": ["SNOW", "NET", "MDB", "DT", "ESTC"]}),
    "palantir":   ("PLTR",      {"all": ["SNOW", "DDOG", "NET", "MDB", "NOW", "BAH", "LDOS", "CACI"],
                                 "商業": ["SNOW", "DDOG", "NET", "MDB", "NOW"],
                                 "政府": ["BAH", "LDOS", "CACI"]}),
    "crowdstrike": ("CRWD",     {"all": ["PANW", "FTNT", "ZS", "OKTA", "S", "CHKP"]}),
    "paloalto":   ("PANW",      {"all": ["FTNT", "ZS", "CHKP"]}),
    "trendmicro": ("TRENDMICRO", {"all": ["CHKP", "GEN", "QLYS", "FTNT"]}),
}
CUR = {t: "JPY" for t in ["FUJITSU", "NRI", "TIS", "TRENDMICRO", "LYCORP", "BIPROGY", "NSSOL", "OBIC", "NEC", "RECRUIT",
                          "CYBERAGENT", "RAKUTEN", "ZOZO", "MERCARI", "MONOTARO", "HITACHI"]}
CUR.update({"NAVER": "KRW", "KAKAO": "KRW"})


def q_end(d):
    """對齊到最近的日曆季末（財季末與日曆季末相差 ≤ 約 1 個月）。"""
    best = None
    for y in (d.year - 1, d.year, d.year + 1):
        for m, dd in ((3, 31), (6, 30), (9, 30), (12, 31)):
            c = dt.date(y, m, dd)
            k = abs((c - d).days)
            if best is None or k < best[0]:
                best = (k, c)
    return best[1]


def read_raw(tk):
    p = os.path.join(RAW, tk + ".txt")
    if not os.path.exists(p):
        return None
    head, rows = "", []
    for ln in open(p, encoding="utf-8"):
        ln = ln.strip()
        if not ln:
            continue
        if ln.startswith("#"):
            head = ln[1:].strip(); continue
        a = ln.split(",")
        d = dt.date.fromisoformat(a[0])
        mc = float(a[1]) if a[1] else None
        ps = float(a[2]) if len(a) > 2 and a[2] not in ("", "-") else None
        rows.append((d, mc, ps))
    rows.sort()
    src = head.split("source", 1)[1].split(";")[0].strip() if "source" in head else ""
    fetched = head.split("fetched")[1].strip() if "fetched" in head else ""
    return dict(tk=tk, head=head, rows=rows, src=src, fetched=fetched)


def build_points(tk):
    """由 raw 算每季點：市值、TTM 營收、年增、PS、PSG、標記。"""
    R = read_raw(tk)
    if R is None:
        return None
    rows = R["rows"][-WINDOW:]
    pts = []
    for i, (d, mc, ps) in enumerate(rows):
        rev = mc / ps if (mc and ps) else None
        pts.append(dict(qe=q_end(d), actual=d, mcap=mc, rev=rev, ps=ps, g=None, psg=None, flags=[]))
    for i, p in enumerate(pts):
        if p["ps"] is None:
            p["flags"].append("PS 缺（資料商未提供）")
        if i >= 4:
            q = pts[i - 4]
            ok_gap = abs((p["actual"] - q["actual"]).days - 365) <= 20
            if p["rev"] and q["rev"] and ok_gap:
                p["g"] = p["rev"] / q["rev"] - 1
            elif not ok_gap:
                p["flags"].append("年增：前 4 季日期不連續")
            else:
                p["flags"].append("年增：缺營收")
        else:
            p["flags"].append("年增：視窗起點前 4 季，無基期")
        if p["g"] is not None and p["ps"]:
            if p["g"] < NM_G:
                p["flags"].append("PSG N/M（TTM 年增 <2% 或為負）")
            else:
                p["psg"] = p["ps"] / (p["g"] * 100)
        if i >= 1 and p["rev"] and pts[i - 1]["rev"]:
            j = p["rev"] / pts[i - 1]["rev"] - 1
            if abs(j) > JUMP:
                p["flags"].append(f"隱含 TTM 營收單季變動 {j:+.0%}（重編／併購／資料錯誤，使用前核對）")
    cur = CUR.get(tk, "USD")
    R["pts"], R["cur"] = pts, cur
    return R


def write_csv(R):
    path = os.path.join(DATA, R["tk"] + ".csv")
    with open(path, "w", encoding="utf-8-sig", newline="") as f:
        w = csv.writer(f)
        w.writerow(["季末日（對齊日曆季末）", "財季實際期末日", "市值（百萬）", "TTM 營收（百萬；市值÷PS 反推）", "TTM 營收年增", "PS", "PSG",
                    "幣別", "來源", "擷取日", "標記"])
        for p in R["pts"]:
            f_ = lambda x, n: "" if x is None else round(x, n)
            w.writerow([p["qe"], p["actual"], f_(p["mcap"], 0), f_(p["rev"], 1), f_(p["g"], 4), f_(p["ps"], 2), f_(p["psg"], 4),
                        R["cur"], R["src"] + "（WebFetch 取自 StockAnalysis；Derived：TTM 營收＝市值÷PS）", R["fetched"], "；".join(p["flags"])])
    return path


def med(v):
    return st.median(v) if v else None


def gm(a, b):
    return math.sqrt(a * b) if a and b and a > 0 and b > 0 else None


def relative(co, peers, key):
    """key='ps' 或 'psg'；回傳序列 [{q, company, peer_median, n_peers, rel}]。"""
    pm = {tk: {p["qe"]: p[key] for p in P["pts"] if p[key] is not None} for tk, P in peers.items()}
    out = []
    for p in co["pts"]:
        row = dict(q=str(p["qe"]), actual=str(p["actual"]), company=p[key], peer_median=None, n_peers=0, rel=None)
        vals = [pm[tk][p["qe"]] for tk in pm if p["qe"] in pm[tk]]
        row["n_peers"] = len(vals)
        if p[key] is not None and len(vals) >= MIN_PEERS:
            row["peer_median"] = med(vals)
            row["rel"] = p[key] / row["peer_median"]
        out.append(row)
    return out


def stats(series):
    v = [r["rel"] for r in series if r["rel"] is not None]
    return dict(n=len(v), median=med(v), min=min(v) if v else None, max=max(v) if v else None)


def growth_ratio(co, peers):
    """現行（最近一個本公司與 ≥2 家同業皆有年增的日曆季）本公司 ÷ 同業中位 TTM 營收年增。"""
    pg = {tk: {p["qe"]: p["g"] for p in P["pts"] if p["g"] is not None and p["g"] >= NM_G} for tk, P in peers.items()}
    for p in reversed(co["pts"]):
        if p["g"] is None or p["g"] < NM_G:
            continue
        vals = [pg[tk][p["qe"]] for tk in pg if p["qe"] in pg[tk]]
        if len(vals) >= MIN_PEERS:
            m = med(vals)
            return dict(value=p["g"] / m, basis="TTM（NTM 未取得）", quarter=str(p["qe"]), company_g=p["g"], peer_median_g=m, n_peers=len(vals))
    return None


def analyse(name, tk, groups):
    co = build_points(tk)
    res = dict(company=name, ticker=tk, currency=CUR.get(tk, "USD"), generated=str(dt.date(2026, 11, 17)),
               method="PS＝市值÷TTM 營收；PSG＝PS÷TTM 營收年增(pt)；相對值＝本公司÷同業中位；讀數與基準見規格第 2 節",
               align="財季末對齊到最近的日曆季末（例：1/31→前一年 12/31、7/31→6/30、4/30→3/31、2/28→3/31、11/30→12/31）；各公司取最近 20 個財季",
               gaps=[])
    if co is None:
        res["gaps"].append(f"{tk}：本公司資料未取得")
        return res, None
    write_csv(co)
    for p in co["pts"]:
        for fl in p["flags"]:
            if "跳動" in fl or "變動" in fl or "PS 缺" in fl:
                res["gaps"].append(f"{tk} {p['qe']}：{fl}")
    res["ttm_growth_method"] = "TTM 營收由 市值÷PS 反推，年增＝本季 TTM ÷ 4 季前 TTM − 1；最早 4 季無基期，故 PSG 最多 16 點"
    res["groups"] = {}
    for gname, plist in groups.items():
        peers, miss = {}, []
        for t in plist:
            P = build_points(t)
            if P is None:
                miss.append(t); continue
            peers[t] = P
        for t, P in peers.items():
            for p in P["pts"]:
                for fl in p["flags"]:
                    if "變動" in fl or "PS 缺" in fl:
                        res["gaps"].append(f"同業 {t} {p['qe']}：{fl}") if (gname == "all" or gname == list(groups)[0]) else None
        if miss:
            res["gaps"].append(f"同業群組 {gname}：缺 {','.join(miss)}")
        ps_s, psg_s = relative(co, peers, "ps"), relative(co, peers, "psg")
        sp, sg = stats(ps_s), stats(psg_s)
        gr = growth_ratio(co, peers)
        G = dict(peers=list(peers), peers_missing=miss, ps_series=ps_s, ps_stats=sp, psg_series=psg_s, psg_stats=sg, growth_ratio=gr)
        ps_read = sp["median"]
        psg_read = sg["median"] * gr["value"] if (sg["median"] and gr) else None
        G["ps_reading"], G["psg_reading"] = ps_read, psg_read
        notes = []
        # 規格第 1 節：有效點 < 12 的序列視為取不到，只列參考、不進基準
        if psg_read is not None and (sg["n"] or 0) < 12:
            G["psg_reading_ref_only"] = psg_read
            notes.append(f"PSG 有效點 {sg['n']} < 12，只作參考，不進基準")
            psg_read = None
        if ps_read is not None and (sp["n"] or 0) < 12:
            G["ps_reading_ref_only"] = ps_read
            notes.append(f"PS 有效點 {sp['n']} < 12，只作參考，不進基準")
            ps_read = None
        if ps_read and psg_read:
            base = gm(ps_read, psg_read)
            low = gm(sp["min"], sg["min"] * gr["value"]); high = gm(sp["max"], sg["max"] * gr["value"])
            G.update(geo_base=base, base_source="PS 與 PSG 兩讀數幾何平均（Derived）", sens=dict(low=low, high=high, ps_only=ps_read, psg_only=psg_read,
                     ps_low=sp["min"], ps_high=sp["max"], psg_low=sg["min"] * gr["value"], psg_high=sg["max"] * gr["value"]))
            G["divergence"] = abs(psg_read / ps_read - 1)
            G["divergence_gt_30pct"] = G["divergence"] > 0.30
        elif ps_read:
            G.update(geo_base=ps_read, base_source="只有 PS（PSG 讀數無法計：" + ("成長比取不到" if not gr else "PSG 有效點不足 12") + "）",
                     sens=dict(low=sp["min"], high=sp["max"], ps_only=ps_read, psg_only=None))
            G["divergence"], G["divergence_gt_30pct"] = None, None
            notes.append("只用 PS（規格第 2 節第 6 條）")
        elif psg_read:
            G.update(geo_base=psg_read, base_source="只有 PSG", sens=dict(low=sg["min"] * gr["value"], high=sg["max"] * gr["value"], ps_only=None, psg_only=psg_read))
            G["divergence"], G["divergence_gt_30pct"] = None, None
        else:
            G.update(geo_base=1.0, base_source="兩條都取不到：1.00，列承重輸入", sens=dict(low=None, high=None, ps_only=None, psg_only=None))
            G["divergence"], G["divergence_gt_30pct"] = None, None
        G["valid_points"] = dict(ps=sp["n"], psg=sg["n"], window_quarters=len(co["pts"]), min_required=12,
                                 meets_12=sp["n"] >= 12 and sg["n"] >= 12)
        if sp["n"] < 12:
            G["gaps"] = G.get("gaps", []) + [f"PS 有效點 {sp['n']} < 12"]
        if sg["n"] < 12:
            G["gaps"] = G.get("gaps", []) + [f"PSG 有效點 {sg['n']} < 12"]
        G["notes"] = notes
        res["groups"][gname] = G
    top = res["groups"]["all"]
    res.update({k: top[k] for k in ("peers", "peers_missing", "ps_series", "ps_stats", "psg_series", "psg_stats", "growth_ratio", "ps_reading",
                                      "psg_reading", "geo_base", "base_source", "sens", "valid_points", "divergence", "divergence_gt_30pct")})
    res["window"] = dict(first=str(co["pts"][0]["qe"]), last=str(co["pts"][-1]["qe"]))
    res["gaps"] = sorted(set(res["gaps"] + sum([g.get("gaps", []) for g in res["groups"].values()], [])))
    return res, co


def fmt(x, n=2):
    return "—" if x is None else f"{x:.{n}f}"


def main():
    ap = argparse.ArgumentParser(); ap.add_argument("--company"); A = ap.parse_args()
    os.makedirs(OUT, exist_ok=True)
    rows = []
    for name, (tk, groups) in COMPANIES.items():
        if A.company and A.company != name:
            continue
        res, co = analyse(name, tk, groups)
        json.dump(res, open(os.path.join(OUT, name + ".json"), "w", encoding="utf-8"), ensure_ascii=False, indent=1, default=str)
        rows.append(res)
    # 同業只作同業、不是 14 家之一者也輸出 CSV
    needed = {t for _, (tk, g) in COMPANIES.items() for t in [tk] + sum(g.values(), [])}
    for t in sorted(needed):
        P = build_points(t)
        if P:
            write_csv(P)
    if A.company:
        return
    L = ["# 歷史相對倍數彙總（PS 與 PSG；2026-11-17）", "",
         "口徑：規格 `docs/plan/歷史相對倍數_PS與PSG.md`。讀數＝本公司÷同業中位的歷史中位；PSG 讀數已乘現行 TTM 成長比（NTM 未取得）；基準＝兩讀數幾何平均（Derived）。"
         "差異＝|PSG 讀數÷PS 讀數−1|。資料為 WebFetch 取自 StockAnalysis 季度比率頁（摘要器轉錄，未逐點二次核對）。", "",
         "| 公司 | PS 讀數 | PSG 讀數 | 基準 | Sens 低–高 | 有效點數（PS／PSG） | 成長比（本÷同業，TTM） | 兩讀數差異 | 差異 > 30% |",
         "|---|---|---|---|---|---|---|---|---|"]
    for r in rows:
        if "valid_points" not in r:
            L.append(f"| {r['company']} | 缺 | 缺 | 缺 | — | 0 | — | — | — |"); continue
        s = r["sens"]; v = r["valid_points"]; gr = r["growth_ratio"]
        dv = r["divergence"]
        L.append(f"| {r['company']} | {fmt(r['ps_reading'])} | {fmt(r['psg_reading'])} | {fmt(r['geo_base'])} | {fmt(s['low'])}–{fmt(s['high'])} | {v['ps']}／{v['psg']}"
                 f" | {fmt(gr['value']) if gr else '—'} | {('%.0f%%' % (dv*100)) if dv is not None else '—'} | {'是' if r['divergence_gt_30pct'] else ('否' if dv is not None else '—')} |")
    L += ["", "## 資料缺口與標記", ""]
    for r in rows:
        if r["gaps"]:
            L.append(f"- **{r['company']}**：" + "；".join(r["gaps"][:12]) + ("…" if len(r["gaps"]) > 12 else ""))
    open(os.path.join(DATA, "summary.md"), "w", encoding="utf-8").write("\n".join(L) + "\n")
    print("\n".join(L))


if __name__ == "__main__":
    main()
