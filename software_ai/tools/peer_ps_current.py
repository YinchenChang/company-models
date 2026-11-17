#!/usr/bin/env python3
"""同業現價 P/S 中位數（2026-11-13）。

讀 data/relmult_hist/peer_ps_20261113.csv，依 relmult_hist.py 的 COMPANIES 同業名單，
對 14 家各群組輸出同業中位 P/S 與家數，另列對照資料庫最近一季同業中位數的變動。
輸出：data/relmult_hist/out/peer_ps_20261113.json

用法（在 software_ai/ 下）：python3 tools/peer_ps_current.py
口徑：P/S＝市值÷TTM 營收（Derived）；同業少於 MIN_PEERS(2) 家有值記 N/M；
價格日不是 2026-11-13 者仍納入，但在 JSON 內列 n_stale 與清單；缺口（無價格）不納入。
"""
import csv, importlib.util, json, os, statistics as st

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
DATA = os.path.join(ROOT, "data", "relmult_hist")
OUT = os.path.join(DATA, "out")
TARGET = "2026-11-13"
MIN_PEERS = 2

spec = importlib.util.spec_from_file_location("relmult_hist", os.path.join(HERE, "relmult_hist.py"))
RH = importlib.util.module_from_spec(spec); spec.loader.exec_module(RH)


def load():
    d = {}
    for r in csv.DictReader(open(os.path.join(DATA, "peer_ps_20261113.csv"), encoding="utf-8-sig")):
        ps = r["P/S"].strip()
        d[r["代號"]] = dict(ps=float(ps) if ps else None, date=r["價格日"], cur=r["幣別"], flag=r["標記"])
    return d


def db_last_median(name):
    """資料庫最近一季（有同業中位數者）的同業中位 P/S，取 out/<公司>.json 各群組 ps_series。"""
    p = os.path.join(OUT, name + ".json")
    if not os.path.exists(p):
        return {}
    J = json.load(open(p, encoding="utf-8"))
    res = {}
    for g, G in J.get("groups", {}).items():
        last = [r for r in G.get("ps_series", []) if r.get("peer_median") is not None]
        if last:
            res[g] = dict(q=last[-1]["q"], median=last[-1]["peer_median"], n_peers=last[-1]["n_peers"])
    return res


def main():
    D = load()
    os.makedirs(OUT, exist_ok=True)
    out = dict(as_of=TARGET, source="data/relmult_hist/peer_ps_20261113.csv（StockAnalysis，WebFetch；Derived）",
               min_peers=MIN_PEERS, companies={})
    for name, (tk, groups) in RH.COMPANIES.items():
        db = db_last_median(name)
        co = dict(ticker=tk, own_ps=D.get(tk, {}).get("ps"), groups={})
        for g, plist in groups.items():
            used = [(t, D[t]) for t in plist if t in D and D[t]["ps"] is not None]
            miss = [t for t in plist if t not in D or D[t]["ps"] is None]
            vals = [x["ps"] for _, x in used]
            stale = [f"{t}({x['date']})" for t, x in used if x["date"] != TARGET]
            m = st.median(vals) if len(vals) >= MIN_PEERS else None
            dbm = db.get(g)
            chg = (m / dbm["median"] - 1) if (m and dbm) else None
            co["groups"][g] = dict(peer_median_ps=m, n_peers=len(vals), n_listed=len(plist), missing=miss,
                                   n_stale=len(stale), stale=stale,
                                   peers={t: x["ps"] for t, x in used},
                                   db_last_quarter=dbm, change_vs_db=chg,
                                   note="N/M（有值同業 < %d）" % MIN_PEERS if m is None else "")
        out["companies"][name] = co
    json.dump(out, open(os.path.join(OUT, "peer_ps_20261113.json"), "w", encoding="utf-8"), ensure_ascii=False, indent=1)
    print(f"{'公司':12s}{'群組':10s}{'同業中位P/S':>12s}{'家數':>6s}{'DB季':>12s}{'DB中位':>8s}{'變動':>8s}  缺／非11-13")
    for n, c in out["companies"].items():
        for g, G in c["groups"].items():
            f = lambda x, k=2: "N/M" if x is None else f"{x:.{k}f}"
            d = G["db_last_quarter"]
            print(f"{n:12s}{g:10s}{f(G['peer_median_ps']):>12s}{G['n_peers']:>3d}/{G['n_listed']:<2d}"
                  f"{(d['q'] if d else '—'):>12s}{f(d['median']) if d else '—':>8s}"
                  f"{('%+.1f%%' % (G['change_vs_db']*100)) if G['change_vs_db'] is not None else '—':>8s}  "
                  f"{','.join(G['missing']) or '-'}／{G['n_stale']}")


if __name__ == "__main__":
    main()
