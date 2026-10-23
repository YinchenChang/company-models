"""AI 半導體總覽：從各模型的重算 xlsx 抽取 HTML 所需數字，輸出 data.json。
只讀模型成品（dist/*.xlsx 的快取值），不重算、不改模型。估值規則（Andy 2026-10-13）：
合理倍數＝10 年（2016–2025 年底）倍數區間內依循環位置取分位：見頂年 25 分位、中段中位數、谷底 75 分位。
"""
import json, pathlib, numpy as np, openpyxl

ROOT = pathlib.Path(__file__).resolve().parents[2]
F = {
    "bridge": ROOT / "markets/chip_bridge/dist/20261022_晶片內容橋接表_v0.2.xlsx",
    "net": ROOT / "markets/network/dist/20261022_AI網通市場模型_v0.2.1.xlsx",
    "sk": ROOT / "skhynix/dist/20261023_SK海力士模型_v0.4.xlsx",
    "ss": ROOT / "samsung/dist/20261023_三星電子模型_v0.2.xlsx",
    "mu": ROOT / "micron/dist/20261023_美光模型_v0.3.xlsx",
    "tsmc": ROOT / "tsmc/dist/20261022_台積電模型_v0.2.xlsx",
    "eq": ROOT / "markets/equipment/dist/20261023_半導體設備市場模型_v0.1.2.xlsx",
    "adv": ROOT / "advantest/dist/20261023_Advantest模型_v0.2.xlsx",
    "eq_readme": ROOT / "markets/equipment/README.md",
}

def rows(path, sheet):
    """回傳 {A 欄標籤（去空白）: [B.. 欄值]}。"""
    ws = openpyxl.load_workbook(path, data_only=True)[sheet]
    out = {}
    for r in ws.iter_rows():
        if isinstance(r[0].value, str):
            out.setdefault(r[0].value.strip(), [c.value for c in r[1:]])
    return out

def name(path, key):
    wb = openpyxl.load_workbook(path, data_only=True)
    s, ref = list(wb.defined_names[key].destinations)[0]
    return wb[s][ref.replace("$", "")].value

# 10 年倍數（2016–2025 年底），兩個來源取各分位的平均；來源見 README
MULT = {
    "sk": {"kind": "P/B", "src": {"companiesmarketcap": [1.31, 1.59, .8832, 1.34, 1.56, 1.45, .81, 1.87, 1.62, 3.86],
                                   "marketcapof（2016–24）": [1.35, 1.60, .90, 1.34, 1.56, 1.45, .82, 1.83, 1.62]}},
    "ss": {"kind": "P/B", "src": {"companiesmarketcap": [1.27, 1.59, 1.04, 1.41, 1.97, 1.73, 1.04, 1.43, .8865, 1.96]},
           "check": "GuruFocus 10 年中位 1.44"},
    "mu": {"kind": "P/B", "src": {"macrotrends（會計年度 8 月底）": [1.29, 1.78, 1.79, 1.33, 1.27, 1.83, 1.22, 1.72, 2.35, 2.46]},
           "alt": {"companiesmarketcap（日曆年底）": [1.83, 2.03, 1.02, 1.64, 2.10, 2.27, 1.11, 2.20, 2.11, 6.04]}},
    "tsmc": {"kind": "P/E", "src": {"FullRatio": [14.5, 17.71, 16.13, 25.62, 31.11, 29.18, 11.94, 19.46, 28.72, 29.03],
                                     "companiesmarketcap": [11.6, 14.8, 14.2, 23.5, 28.3, 27.2, 11.2, 19.3, 28.8, 28.3]}},
}
def pct(m):
    ps = {}
    for q in (25, 50, 75):
        ps[f"p{q}"] = round(float(np.mean([np.percentile(v, q) for v in m["src"].values()])), 2)
    ps["min"] = round(min(min(v) for v in m["src"].values()), 2)
    ps["max"] = round(max(max(v) for v in m["src"].values()), 2)
    if "alt" in m:
        a = list(m["alt"].values())[0]
        ps["alt"] = {f"p{q}": round(float(np.percentile(a, q)), 2) for q in (25, 50, 75)}
    return ps

D = {}

# 橋接表
b = rows(F["bridge"], "Bridge")
cols = ["GB300", "VR200 2026", "VR200 2027 吸收"]
items = [("廠房（機房、電力、冷卻）", "├ 廠房（機房、電力、冷卻）"),
         ("NVIDIA GPU 毛利", "│  │  ├ GPU 晶片：NVIDIA 毛利（設計）"),
         ("台積電（晶圓＋CoWoS）", "│  │  ├ GPU 晶片＋CoWoS：台積電"),
         ("HBM", "│  │  ├ HBM"),
         ("其他封裝製造", "│  │  └ 其他製造（基板、測試、良率損失）"),
         ("CPU", "│  ├ CPU（Grace／Vera）"),
         ("LPDDR", "│  ├ 非 HBM DRAM（LPDDR5X）"),
         ("NAND／SSD", "│  ├ NAND／SSD（機架內）"),
         ("網卡／DPU", "│  ├ SuperNIC＋DPU（scale-out 端點，機架內）"),
         ("機架其他（電源、冷卻、整合、殘差）", None),
         ("網通（交換、光、銅）", "├ 網通"),
         ("儲存與管理", "└ 儲存與管理")]
def brow(lbl):
    for k, v in b.items():
        if k.endswith(lbl.split(" ", 1)[-1]) or k == lbl:
            return v
    raise KeyError(lbl)
tot = brow("每 GW capex 合計")
bridge = {c: [] for c in cols}
for i, c in enumerate(cols):
    acc = 0
    for lab, key in items:
        if key is None:
            v = brow("│  ├ 非 NVIDIA 機架內容（電源、冷卻、機構、整合）")[i] + brow("│  └ 殘差：匣級電路板、供電、NVIDIA／ODM 匣級加價")[i]
        else:
            v = brow(key)[i]
        acc += v
        bridge[c].append([lab, round(v, 3)])
    assert abs(acc - tot[i]) < 1e-6, (c, acc, tot[i])
D["bridge"] = {"cols": cols, "total": [round(x, 2) for x in tot[:3]], "items": bridge,
               "hbm_per_acc_k": [round(x, 2) for x in brow("每顆加速器 HBM 金額 $K")[:3]],
               "tsmc_of_gpu": [round(x, 3) for x in brow("台積電 ÷ GPU 售價")[:3]],
               "hbm_of_gpu": [round(x, 3) for x in brow("HBM ÷ GPU 售價")[:3]]}

# 記憶體市場層（取 SK 海力士工作簿的 Path／Mkt；三家共用）
p = openpyxl.load_workbook(F["sk"], data_only=True)["Path"]
q = [p.cell(3, j).value for j in range(3, 23)]
D["path"] = {"quarters": q,
             "dram": {s: [round(p.cell(r, j).value, 3) for j in range(3, 23)] for s, r in (("A", 4), ("B", 5), ("C", 6))},
             "nand": {s: [round(p.cell(r, j).value, 3) for j in range(3, 23)] for s, r in (("A", 8), ("B", 9), ("C", 10))}}
m = rows(F["sk"], "Mkt")
D["hbm"] = {"years": [2025, 2026, 2027, 2028, 2029],
            "mkt": {s: [round(x, 1) for x in m[f"HBM 市場（{s}）"][1:6]] for s in "ABC"},
            "price": {s: [round(x, 2) for x in m[f"HBM 單價（{s}）"][1:6]] for s in "ABC"},
            "supply_eb": [round(x, 2) for x in m["供給（三情境相同）"][1:6]],
            "gap_eb": [None] + [round(x, 2) for x in m["需求 − 供給（A；>0＝短缺）"][2:6]],
            "ship_eb": {s: [round(x, 2) for x in m[f"出貨＝MIN(需求, 供給)（{s}）"][1:6]] for s in "ABC"},
            "gap_eb_s": {s: [None] + [round(x, 2) for x in m[f"需求 − 供給（{s}；>0＝短缺）"][2:6]] for s in "ABC"},
            "units_m": [None] + [round(x, 1) for x in m["加速器顆數合計"][2:5]]}

# 記憶體三家
SC = ["2026", "2027 A", "2028 A", "2029 A", "2027 B", "2028 B", "2029 B", "2027 C", "2028 C", "2029 C"]
mem = {}
for k, ccy, fyl in (("sk", "KRW", "2027"), ("ss", "KRW", "2027"), ("mu", "USD", "FY2027")):
    y = rows(F[k], "Y")
    px = name(F[k], "px"); fx = name(F[k], "fx") if k != "mu" else 1
    def g(lbl): return {SC[i]: y[lbl][1 + i] for i in range(10)} | {"2025": y[lbl][0]}
    op = g("記憶體營業利益")  # 三星只取記憶體事業，三家口徑一致
    bv = g("每股淨值 BVPS"); eps = g("每股盈餘 EPS（期末股數）")
    mt = MULT[k]; ps = pct(mt)
    # 規則套用：2027＝見頂年（情境 A 命題）→ 25 分位 × 2027 年底 BVPS；谷底＝2028 C → 75 分位 × 2028 C BVPS
    peak = bv["2027 A"] * ps["p25"]; trough = bv["2028 C"] * ps["p75"]
    mem[k] = {"ccy": ccy, "fx": fx, "px": px,
              "op_usd_b": {s: (round(v * (1000 if ccy == "KRW" else 1) / fx, 1) if v is not None else None) for s, v in op.items()},
              "opm": {s: round(v, 3) for s, v in g("記憶體營業利益率").items() if v is not None},
              "hbm_share": {s: round(v, 3) for s, v in g("HBM ÷ 營收").items() if v is not None},
              "bvps": {s: round(v, 2) for s, v in bv.items() if v is not None},
              "eps": {s: round(v, 2) for s, v in eps.items() if v is not None},
              "roe": {s: round(v, 3) for s, v in g("ROE（期末權益）").items() if v is not None},
              "pb_now_on_2026": round(px / bv["2026"], 2), "pe_2027A": round(px / eps["2027 A"], 1),
              "mult": ps, "mult_src": list(mt["src"].keys()) + ([mt["check"]] if "check" in mt else []),
              "val": {"peak_px": round(peak, 1), "peak_vs": round(peak / px - 1, 3),
                      "mid_px": round(bv["2027 A"] * ps["p50"], 1), "mid_vs": round(bv["2027 A"] * ps["p50"] / px - 1, 3),
                      "trough_px": round(trough, 1), "trough_vs": round(trough / px - 1, 3),
                      "implied_pb_2027A": round(px / bv["2027 A"], 2)}}
    if "alt" in ps:
        mem[k]["val"]["peak_alt_vs"] = round(bv["2027 A"] * ps["alt"]["p25"] / px - 1, 3)
D["mem"] = mem

# 網通
n = rows(F["net"], "Mkt")
def nr(lbl, nd=1): return [round(x, nd) for x in n[lbl][1:5]]
D["net"] = {"years": [2026, 2027, 2028, 2029],
            "comp": [[lab, nr(lab)] for lab in ["scale-up 交換晶片", "scale-up 互連與交換匣其他（銅背板、線纜、電路板、機構）",
                                                 "scale-out 交換器", "scale-out 光模組", "scale-out 銅纜（DAC／AEC）", "網卡／DPU（SuperNIC）"]],
            "total": nr("AI 網通市場合計"), "su_share": nr("scale-up（交換＋互連）÷ 合計", 3),
            "cpo": nr("其中 CPO"), "plug": nr("其中可插拔"), "optics": nr("光學（可插拔＋CPO＋scale-up 光學）"),
            "copper": nr("銅（scale-up 背板與線纜＋scale-out DAC／AEC）"),
            "it_capex": nr("AI IT 資本支出（加速器 × 每顆 IT capex）"), "net_of_it": nr("網通 ÷ AI IT capex", 3),
            "nvda": nr("NVIDIA 網通（本模型歸屬）"), "avgo": nr("Broadcom 網通（本模型歸屬）"),
            "nvda_avgo_share": nr("NVIDIA＋Broadcom ÷ 市場", 3),
            "recon": {"nvda": [round(n["NVIDIA 網通（本模型歸屬）"][1], 1), n["對照：NVIDIA 實際（2026）"][1]],
                      "avgo": [round(n["Broadcom 網通（本模型歸屬）"][1], 1), n["對照：Broadcom 實際（2026）"][1]],
                      "optics": [round(n["光學：本模型 vs LightCounting（2026）"][1], 1), n["光學：本模型 vs LightCounting（2026）"][2]]}}

# 台積電
c = rows(F["tsmc"], "CoWoS"); pl = rows(F["tsmc"], "PL")
TS = ["2026", "2027 A", "2028 A", "2029 A", "2027 B", "2028 B", "2029 B", "2027 C", "2028 C"]
def tp(lbl, nd=2): return {TS[i]: round(pl[lbl][i], nd) for i in range(9)}
px = name(F["tsmc"], "px"); ps = pct(MULT["tsmc"]); eps = tp("EPS（新台幣）")
D["tsmc"] = {"px": px,
             "cowos": {"years": [2026, 2027, 2028, 2029], "demand": [round(x) for x in c["需求合計"][:4]],
                       "nvda": [round(x) for x in c["其中 NVIDIA"][:4]],
                       "cap": [round(x) for x in c["可用產能（含委外）"][:4]], "util": [round(x, 3) for x in c["需求 ÷ 可用產能"][:4]],
                       "hbm_gap_2027": c["對照：HBM 需求缺口 ÷ 供給（2027）"][1],
                       "ms": [c["對照：Morgan Stanley 全球需求"][0], c["對照：Morgan Stanley 全球需求"][1]]},
             "ai_rev": tp("AI 相關營收（含其他 AI）", 1), "rev": tp("營收", 1), "ai_share": tp("AI ÷ 營收", 3),
             "gm": tp("毛利率", 3), "eps": eps, "capex": tp("資本支出", 0),
             "mult": ps, "mult_src": list(MULT["tsmc"]["src"].keys()),
             # 規則套用：2027＝中段 → 中位數；區間 25–75 分位
             "val": {"mid_px": round(eps["2027 A"] * ps["p50"]), "mid_vs": round(eps["2027 A"] * ps["p50"] / px - 1, 3),
                     "lo_px": round(eps["2027 A"] * ps["p25"]), "lo_vs": round(eps["2027 A"] * ps["p25"] / px - 1, 3),
                     "hi_px": round(eps["2027 A"] * ps["p75"]), "hi_vs": round(eps["2027 A"] * ps["p75"] / px - 1, 3),
                     "pe_2027": round(px / eps["2027 A"], 1), "pe_2026": round(px / eps["2026"], 1),
                     "c28_px": round(eps["2028 C"] * ps["p75"]), "c28_vs": round(eps["2028 C"] * ps["p75"] / px - 1, 3)}}


# ---------------------------------------------------------------------------
# 美光替代估值（Andy 2026-10-22「都做」）：(1) P/E；(2) 買回加回的 P/B。
# 歷史值：SEC 10-K XBRL（Verified）——權益、庫藏股（成本）、歸屬淨利；P/B 為 macrotrends 會計年度底。
MU_HIST = {  # 會計年度：P/B, 權益 $B, 庫藏股 $B, 淨利 $B
    2016: (1.29, 12.080, 1.029, -0.276), 2017: (1.78, 18.621, 0.067, 5.089), 2018: (1.79, 32.294, 0.429, 14.135),
    2019: (1.33, 35.881, 3.221, 6.313), 2020: (1.27, 38.996, 3.495, 2.687), 2021: (1.83, 43.933, 4.695, 5.861),
    2022: (1.22, 49.907, 7.127, 8.687), 2023: (1.72, 44.120, 7.552, -5.833), 2024: (2.35, 45.131, 7.852, 0.778),
    2025: (2.46, 54.165, 7.852, 8.539)}
mu_pe = {y: pb * e / ni for y, (pb, e, t, ni) in MU_HIST.items() if ni > 0}          # 市值 ÷ 淨利＝P/B × 權益 ÷ 淨利
mu_pba = {y: pb * e / (e + t) for y, (pb, e, t, ni) in MU_HIST.items()}             # 市值 ÷（權益＋庫藏股）
def q3(v): return {f"p{q}": round(float(np.percentile(list(v), q)), 2) for q in (25, 50, 75)} | {"min": round(min(v), 2), "max": round(max(v), 2)}
mpe, mpba = q3(mu_pe.values()), q3(mu_pba.values())
yu = rows(F["mu"], "Y")
def gy(lbl, s): return yu[lbl][1 + SC.index(s)]
T_last = MU_HIST[2025][2]  # FY2026 年底庫藏股尚未公布，用 FY2025（偏低）
def adj_bvps(s, bbs):
    return (gy("期末權益", s) + T_last + sum(gy("買回", b) for b in bbs)) / gy("期末流通股數", s) * 1000
mu_px = D["mem"]["mu"]["px"]
adj27, adj28c = adj_bvps("2027 A", ["2027 A"]), adj_bvps("2028 C", ["2027 C", "2028 C"])
eps27, eps28c = D["mem"]["mu"]["eps"]["2027 A"], D["mem"]["mu"]["eps"]["2028 C"]
D["mu_alt"] = {"pe": mpe, "pba": mpba, "pe_hist": {y: round(v, 2) for y, v in mu_pe.items()}, "pba_hist": {y: round(v, 3) for y, v in mu_pba.items()},
               "adj_bvps": {"FY2027 A": round(adj27, 1), "FY2028 C": round(adj28c, 1)}}

# ---------------------------------------------------------------------------
# SK 海力士、三星 P/E 對照（Andy 2026-10-22：三家同樣呈現）。年底 P/E：companiesmarketcap（2016–2025）；
# stockanalysis 2021–2025 交叉核對（差距 ≤ 8%）；2016–2020 只有單一來源。虧損年（P/E < 0）不計。
KR_PE = {"sk": ("SK 海力士（P/E）", {2016: 9.57, 2017: 4.61, 2018: 2.54, 2019: 30.2, 2020: 16.2, 2021: 9.01, 2022: 22.6, 2023: -10.6, 2024: 5.99, 2025: 10.5}),
         "ss": ("三星電子（P/E）", {2016: 9.15, 2017: 6.95, 2018: 5.06, 2019: 15.3, 2020: 19.1, 2021: 12.5, 2022: 6.50, 2023: 36.1, 2024: 10.6, 2025: 18.2})}
KR_PE_CHECK = {"sk": {2021: 9.38, 2022: 23.13, 2024: 6.06, 2025: 10.47}, "ss": {2021: 13.40, 2022: 6.79, 2023: 35.92, 2024: 10.53, 2025: 17.39}}
kr_pe = {}
for k, (nm, hist) in KR_PE.items():
    v = {y: x for y, x in hist.items() if x > 0}
    kr_pe[k] = (nm, v, q3(v.values()))
    for y, x in KR_PE_CHECK[k].items():
        assert abs(x / hist[y] - 1) < 0.08, (k, y)

# 第二來源核對（只核對，不改 KR_PE）：DART 合併財報（Verified）的歸屬母公司權益與歸屬淨利（兆韓元），
# 換算 P/E = P/B × 權益 ÷ 淨利，與 companiesmarketcap 的年底 P/E 逐年比對。
# 取各年度自己的事業報告書（該年度有「기재정정」者取最新更正版）；來源 = DART 接收編號，
# 連結 https://dart.fss.or.kr/dsaf001/main.do?rcpNo=<接收編號>（III. 재무에 관한 사항 → 2. 연결재무제표）。
KR_HIST = {  # 年度：(歸屬母公司權益, 歸屬淨利, DART 接收編號)
    "sk": {2016: (24.017, 2.954, "20170331004537"), 2017: (33.815, 10.642, "20180615000111"), 2018: (46.846, 15.540, "20190404002329"),
           2019: (47.928, 2.013, "20200330004441"), 2020: (51.889, 4.755, "20210330000776"), 2021: (62.157, 9.602, "20220322000590"),
           2022: (63.266, 2.230, "20230321001209"), 2023: (53.504, -9.112, "20240319000684"), 2024: (73.903, 19.789, "20250319000665"),
           2025: (120.516, 42.919, "20260317000635")},
    "ss": {2016: (186.424, 22.416, "20170331004518"), 2017: (207.213, 41.345, "20180402005019"), 2018: (240.069, 43.891, "20190401004781"),
           2019: (254.915, 21.505, "20200330003851"), 2020: (267.670, 26.091, "20210309000744"), 2021: (296.238, 39.244, "20220308000798"),
           2022: (345.186, 54.730, "20230307000542"), 2023: (353.234, 14.473, "20240312000736"), 2024: (391.688, 33.621, "20250311001085"),
           2025: (424.313, 44.261, "20260310002820")}}
KR_PB = {"sk": {y: (a + b) / 2 if y < 2025 else a for y, (a, b) in enumerate(zip(MULT["sk"]["src"]["companiesmarketcap"], MULT["sk"]["src"]["marketcapof（2016–24）"] + [None]), 2016)},
         "ss": dict(enumerate(MULT["ss"]["src"]["companiesmarketcap"], 2016))}   # SK＝兩來源平均（2025 只有 cmc）；三星＝cmc
KR_PE_GAP = 0.15
kr_chk = []
for k, (nm, hist) in KR_PE.items():
    for y, (eq, ni, rc) in KR_HIST[k].items():
        pe_dart = KR_PB[k][y] * eq / ni
        gap = pe_dart / hist[y] - 1
        kr_chk.append({"company": nm, "year": y, "pb": round(KR_PB[k][y], 3), "equity_tn": eq, "net_income_tn": ni, "pe_dart": round(pe_dart, 2),
                       "pe_cmc": hist[y], "gap": round(gap, 3), "flag": abs(gap) > KR_PE_GAP, "dart_rcp_no": rc})
for r in kr_chk:
    if r["flag"]:
        print("KR P/E 差距 >15%：", r["company"], r["year"], "換算", r["pe_dart"], "cmc", r["pe_cmc"], f'{r["gap"]:+.1%}')

# ---------------------------------------------------------------------------
# 設備市場（markets/equipment v0.1.2）與 Advantest（advantest v0.1）
import re as _re
wbq = openpyxl.load_workbook(F["eq"], data_only=True)
mk = wbq["Mkt"]; sc = None; eqs = {}
LBL = {"設備總額": "total", "年增": "yoy", "其中晶圓前段": "wfe", "AI 相關設備": "ai", "AI 增量占總增量（相對 2025 累計）": "ai_inc_share", "見頂年度": "peak"}
for r in mk.iter_rows(min_row=4, max_row=30):
    a = r[0].value
    if isinstance(a, str) and a.startswith("情境 "):
        sc = a[-1]; continue
    if sc and isinstance(a, str) and a.strip() in LBL:
        eqs.setdefault(sc, {})[LBL[a.strip()]] = [c.value for c in r[2:7]]
EQY = [2025, 2026, 2027, 2028, 2029]
eq_rows = []
for s_, d in eqs.items():
    for i, y_ in enumerate(EQY):
        eq_rows.append({"key": f"{s_} {y_}", "scenario": s_, "year": y_, "total": round(d["total"][i], 1),
                        "wfe": round(d["wfe"][i], 1), "ai": round(d["ai"][i], 1),
                        "yoy": (round(d["yoy"][i], 3) if isinstance(d["yoy"][i], (int, float)) else None),
                        "ai_inc_share": (round(d["ai_inc_share"][i], 3) if isinstance(d["ai_inc_share"][i], (int, float)) else None),
                        "peak": d["peak"][0]})
def mrow(label):
    for r in mk.iter_rows():
        if isinstance(r[0].value, str) and r[0].value.strip().startswith(label): return [c.value for c in r]
    raise KeyError(label)
calib = {"sk_ss": mrow("校準值：SK 海力士、三星設備占比")[2], "sk_ss_lo": mrow("低：長鑫取高值")[2], "sk_ss_hi": mrow("高：業界下限")[2],
         "industry": mrow("對照：業界值")[2], "big3_2026": mrow("三家加權設備占比")[3], "other_mem_2025": mrow("其他記憶體廠設備（獨立估計）")[2],
         "semi_mem_2025": mrow("SEMI 2025 DRAM＋NAND 設備（實績）")[2]}
rc = wbq["Recon"]; eq_recon = []
for r in rc.iter_rows(min_row=4, max_row=21):
    v = [c.value for c in r[:9]]
    if v[0] and isinstance(v[3], (int, float)):
        src = _re.sub(r"\s*https?://\S+", "", str(v[2]))
        src = src + "）" * (src.count("（") - src.count("）"))
        eq_recon.append({"key": f"{v[0]} {v[1]} {src}", "item": v[0], "year": v[1], "source": src, "value": v[3],
                         "model_a": round(v[4], 1), "gap_a": round(v[5], 3), "model_c": round(v[8], 1)})
bt = []; txt = F["eq_readme"].read_text()
blk = txt[txt.index("## 瓶頸表"):txt.index("## 限制")]
for line in blk.splitlines():
    if line.startswith("| ") and not line.startswith("| 瓶頸") and not line.startswith("|---"):
        c = [x.strip() for x in line.strip("|").split("|")]
        clean = lambda t: _re.sub(r"\[([^\]]+)\]\([^)]+\)", r"\1", t).replace("**", "").replace("『", "「").replace("』", "」")
        bt.append({"bottleneck": c[0], "evidence": clean(c[1]), "source": clean(c[2]), "strength": clean(c[3])})
D["equip"] = {"scen": eq_rows, "calib": calib}

wa = openpyxl.load_workbook(F["adv"], data_only=True); pa = wa["PL"]; sc = None; adv = []
for r in pa.iter_rows(min_row=4, max_row=36):
    a, u = r[0].value, r[1].value
    if isinstance(a, str) and a.startswith("情境 "):
        sc = a[-1]; continue
    if sc and isinstance(a, str):
        k = {("營收", "¥B"): "rev", ("營收", "US$B"): "rev_usd", ("SoC 測試機", "¥B"): "soc", ("記憶體測試機", "¥B"): "mem",
             ("營業利益率", "%"): "opm", ("EPS", "¥"): "eps", ("營業利益", "¥B"): "op"}.get((a.strip(), u))
        if k:
            for i, fy in enumerate(["FY2025", "FY2026", "FY2027", "FY2028", "FY2029"]):
                row = next((x for x in adv if x["scenario"] == sc and x["fy"] == fy), None)
                if row is None:
                    row = {"key": f"{sc} {fy}", "scenario": sc, "fy": fy}; adv.append(row)
                row[k] = round(r[2 + i].value, 3)
va = wa["Val"]
def vrow(label):
    for r in va.iter_rows():
        if r[0].value == label: return [c.value for c in r]
adv_ps = {"p25": round(vrow("25 分位")[1], 2), "p50": round(vrow("中位數")[1], 2), "p75": round(vrow("75 分位")[1], 2)}
adv_px = vrow("現價（stockanalysis 2026-10-21 收盤）")[1]
adv_fpe = vrow("遠期 P/E：stockanalysis")[1]
adv_pe_hist = []
for r in va.iter_rows():
    if isinstance(r[0].value, str) and _re.match(r"^20\d\d-12$", r[0].value) and isinstance(r[4].value, (int, float)):
        adv_pe_hist.append((int(r[0].value[:4]), round(r[4].value, 2)))
adv_hist_vals = [x for _, x in adv_pe_hist]
adv_ps |= {"min": min(adv_hist_vals), "max": max(adv_hist_vals)}
eA27 = next(x for x in adv if x["key"] == "A FY2027")["eps"]; eC28 = next(x for x in adv if x["key"] == "C FY2028")["eps"]

out = pathlib.Path(__file__).with_name("data.json")
out.write_text(json.dumps(D, ensure_ascii=False, indent=1))
for k in ("sk", "ss", "mu"):
    print(k, D["mem"][k]["mult"], D["mem"][k]["val"], "P/B now", D["mem"][k]["pb_now_on_2026"], "PE27", D["mem"][k]["pe_2027A"])
print("tsmc", D["tsmc"]["mult"], D["tsmc"]["val"])

# ---------------------------------------------------------------------------
# 儀表板資料集（每個檔案一個 JSON 陣列，列物件）。欄位名稱即頁面 data-field。
DS = out.parent / "datasets"; DS.mkdir(exist_ok=True)
def dump(name, rows_):
    (DS / f"{name}.json").write_text(json.dumps(rows_, ensure_ascii=False, indent=0))

dump("bridge", [{"item": lab, "gb300": D["bridge"]["items"]["GB300"][i][1],
                 "vr200_2026": D["bridge"]["items"]["VR200 2026"][i][1],
                 "vr200_2027": D["bridge"]["items"]["VR200 2027 吸收"][i][1]} for i, (lab, _) in enumerate(items)])

CO = {"sk": "SK 海力士", "ss": "三星電子", "mu": "美光"}
my = []
for k, nm in CO.items():
    e = D["mem"][k]
    for s in ["2025", "2026"] + SC[1:]:
        yr, sc = (s, "實績" if s == "2025" else "基準") if " " not in s else s.split(" ")
        my.append({"company": nm, "year": int(yr), "scenario": sc, "key": f"{nm} {s}",
                   "op_usd_b": e["op_usd_b"].get(s), "opm": e["opm"].get(s), "hbm_share": e["hbm_share"].get(s),
                   "bvps": e["bvps"].get(s), "eps": e["eps"].get(s), "roe": e["roe"].get(s)})
dump("mem_y", my)

dump("mem_path", [{"quarter": qq, **{f"dram_{s}": D["path"]["dram"][s][i] for s in "ABC"},
                   **{f"nand_{s}": D["path"]["nand"][s][i] for s in "ABC"}} for i, qq in enumerate(D["path"]["quarters"])])

h = D["hbm"]
dump("hbm_mkt", [{"year": y_, **{f"mkt_{s}": h["mkt"][s][i] for s in "ABC"}, **{f"price_{s}": h["price"][s][i] for s in "ABC"},
                  "supply_eb": h["supply_eb"][i], **{f"ship_{s}": h["ship_eb"][s][i] for s in "ABC"}, **{f"gap_eb_{s}": h["gap_eb_s"][s][i] for s in "ABC"}, "gap_eb": h["gap_eb"][i],
                  "gap_pct": (round(h["gap_eb"][i] / h["supply_eb"][i], 3) if h["gap_eb"][i] is not None else None)}
                 for i, y_ in enumerate(h["years"])])

val = []
for k, nm in CO.items():
    e = D["mem"][k]; ps = e["mult"]; v = e["val"]
    base = {"company": nm, "metric": "P/B", "ccy": e["ccy"], "px": e["px"], "p25": ps["p25"], "p50": ps["p50"], "p75": ps["p75"],
            "hist_min": ps["min"], "hist_max": ps["max"]}
    yl = "FY2027" if k == "mu" else "2027"
    for case, pos, mult, basis_lbl, basis, ipx, vs in (
        ("見頂", "見頂年 → 25 分位", ps["p25"], f"{yl} 年底每股淨值（A）", e["bvps"]["2027 A"], v["peak_px"], v["peak_vs"]),
        ("中段", "中段 → 中位數", ps["p50"], f"{yl} 年底每股淨值（A）", e["bvps"]["2027 A"], v["mid_px"], v["mid_vs"]),
        ("谷底", "谷底 → 75 分位", ps["p75"], ("FY2028" if k == "mu" else "2028") + " 年底每股淨值（C）", e["bvps"]["2028 C"], v["trough_px"], v["trough_vs"])):
        val.append(base | {"key": f"{nm} {case}", "case": case, "position": pos, "multiple": mult, "basis_label": basis_lbl,
                           "basis": basis, "implied_px": ipx, "vs_px": vs, "px_multiple": round(e["px"] / basis, 2)})
t = D["tsmc"]; ps = t["mult"]; v = t["val"]
tb = {"company": "台積電", "metric": "P/E", "ccy": "TWD", "px": t["px"], "p25": ps["p25"], "p50": ps["p50"], "p75": ps["p75"],
      "hist_min": ps["min"], "hist_max": ps["max"]}
for case, pos, mult, ipx, vs in (("中段", "中段 → 中位數", ps["p50"], v["mid_px"], v["mid_vs"]),
                                 ("低", "25 分位（區間下緣）", ps["p25"], v["lo_px"], v["lo_vs"]),
                                 ("高", "75 分位（區間上緣）", ps["p75"], v["hi_px"], v["hi_vs"])):
    val.append(tb | {"key": f"台積電 {case}", "case": case, "position": pos, "multiple": mult, "basis_label": "2027 EPS（A）",
                     "basis": t["eps"]["2027 A"], "implied_px": ipx, "vs_px": vs, "px_multiple": round(t["px"] / t["eps"]["2027 A"], 2)})

for nm, metric, ps, cases in (
    ("美光（P/E）", "P/E", mpe, [("見頂", "見頂年 → 25 分位", "p25", "FY2027 EPS（A）", eps27), ("中段", "中段 → 中位數", "p50", "FY2027 EPS（A）", eps27),
                                 ("谷底", "谷底 → 75 分位", "p75", "FY2028 EPS（C）", eps28c)]),
    ("美光（買回加回 P/B）", "P/B", mpba, [("見頂", "見頂年 → 25 分位", "p25", "FY2027 年底調整每股淨值（A）", adj27), ("中段", "中段 → 中位數", "p50", "FY2027 年底調整每股淨值（A）", adj27),
                                     ("谷底", "谷底 → 75 分位", "p75", "FY2028 年底調整每股淨值（C）", adj28c)])):
    for case, pos, k, bl, basis in cases:
        ip = ps[k] * basis
        val.append({"company": nm, "metric": metric, "ccy": "USD", "px": mu_px, "p25": ps["p25"], "p50": ps["p50"], "p75": ps["p75"],
                    "hist_min": ps["min"], "hist_max": ps["max"], "key": f"{nm} {case}", "case": case, "position": pos, "multiple": ps[k],
                    "basis_label": bl, "basis": round(basis, 2), "implied_px": round(ip, 1), "vs_px": round(ip / mu_px - 1, 3),
                    "px_multiple": round(mu_px / basis, 2)})
for k, (nm, v_, ps) in kr_pe.items():
    e = D["mem"][k]
    for case, pos, q, bl, basis in (("見頂", "見頂年 → 25 分位", "p25", "2027 EPS（A）", e["eps"]["2027 A"]), ("中段", "中段 → 中位數", "p50", "2027 EPS（A）", e["eps"]["2027 A"]),
                                    ("谷底", "谷底 → 75 分位", "p75", "2028 EPS（C）", e["eps"]["2028 C"])):
        ip = ps[q] * basis
        val.append({"company": nm, "metric": "P/E", "ccy": "KRW", "px": e["px"], "p25": ps["p25"], "p50": ps["p50"], "p75": ps["p75"],
                    "hist_min": ps["min"], "hist_max": ps["max"], "key": f"{nm} {case}", "case": case, "position": pos, "multiple": ps[q],
                    "basis_label": bl, "basis": round(basis, 2), "implied_px": round(ip, 1), "vs_px": round(ip / e["px"] - 1, 3),
                    "px_multiple": round(e["px"] / basis, 2)})
for case, pos, q, bl, basis in (("見頂", "見頂年 → 25 分位", "p25", "FY2027 EPS（A）", eA27), ("中段", "中段 → 中位數", "p50", "FY2027 EPS（A）", eA27),
                                ("谷底", "谷底 → 75 分位", "p75", "FY2028 EPS（C）", eC28)):
    ip = adv_ps[q] * basis
    val.append({"company": "Advantest", "metric": "P/E", "ccy": "JPY", "px": adv_px, "p25": adv_ps["p25"], "p50": adv_ps["p50"], "p75": adv_ps["p75"],
                "hist_min": adv_ps["min"], "hist_max": adv_ps["max"], "key": f"Advantest {case}", "case": case, "position": pos, "multiple": adv_ps[q],
                "basis_label": bl, "basis": round(basis, 2), "implied_px": round(ip, 1), "vs_px": round(ip / adv_px - 1, 3), "px_multiple": round(adv_px / basis, 2),
                "fwd_pe": adv_fpe})
dump("valuation", val)

mh = []
for k, nm in list(CO.items()) + [("tsmc", "台積電")]:
    for src, vals in MULT[k]["src"].items():
        for i, x in enumerate(vals):
            mh.append({"company": nm, "metric": MULT[k]["kind"], "year": 2016 + i, "source": src, "multiple": x})
for y_, v_ in mu_pe.items(): mh.append({"company": "美光（P/E）", "metric": "P/E", "year": y_, "source": "SEC 10-K＋macrotrends", "multiple": round(v_, 2)})
for y_, v_ in mu_pba.items(): mh.append({"company": "美光（買回加回 P/B）", "metric": "P/B", "year": y_, "source": "SEC 10-K＋macrotrends", "multiple": round(v_, 3)})
for k, (nm, v_, ps) in kr_pe.items():
    for y_, x in v_.items(): mh.append({"company": nm, "metric": "P/E", "year": y_, "source": "companiesmarketcap", "multiple": x})
for y_, x in adv_pe_hist: mh.append({"company": "Advantest", "metric": "P/E", "year": y_, "source": "年底股價 ÷ 當期 FY 基本 EPS（自算）", "multiple": x})
dump("mult_hist", mh)
dump("equip_scen", eq_rows)
dump("equip_calib", [{k: (round(v, 4) if isinstance(v, float) else v) for k, v in calib.items()}])
dump("equip_recon", eq_recon)
dump("equip_bottle", bt)
dump("adv_pl", adv)
dump("kr_pe_check", kr_chk)
dump("mu_hist", [{"fy": y_, "pb": pb, "equity": e, "treasury": t, "net_income": ni, "roe": round(ni / e, 3),
                  "pe": (round(pb * e / ni, 2) if ni > 0 else None), "pb_adj": round(pb * e / (e + t), 3)} for y_, (pb, e, t, ni) in MU_HIST.items()])

nt = D["net"]
dump("net_comp", [{"component": lab, **{f"y{y_}": vals[i] for i, y_ in enumerate(nt["years"])}} for lab, vals in nt["comp"]])
nk = []
for i, y_ in enumerate(nt["years"]):
    r = {"year": y_, "total": nt["total"][i], "su_share": nt["su_share"][i], "so_share": round(1 - nt["su_share"][i], 3),
         "net_of_it": nt["net_of_it"][i], "it_capex": nt["it_capex"][i], "nvda": nt["nvda"][i], "avgo": nt["avgo"][i],
         "nvda_avgo_share": nt["nvda_avgo_share"][i], "optics": nt["optics"][i], "pluggable": nt["plug"][i], "cpo": nt["cpo"][i],
         "copper": nt["copper"][i]}
    if i == 0:
        r |= {"nvda_ref": nt["recon"]["nvda"][1], "avgo_ref": nt["recon"]["avgo"][1], "optics_ref": nt["recon"]["optics"][1]}
    nk.append(r)
dump("net_kpi", nk)

cw = t["cowos"]
dump("tsmc_cowos", [{"year": y_, "demand": cw["demand"][i], "nvidia": cw["nvda"][i], "capacity": cw["cap"][i], "util": cw["util"][i],
                     "ms_demand": (cw["ms"][i] if i < 2 else None), "hbm_gap": (cw["hbm_gap_2027"] if y_ == 2027 else None)}
                    for i, y_ in enumerate(cw["years"])])
dump("tsmc_pl", [{"key": s, "year": int(s[:4]), "scenario": (s[5:] or "基準"), "ai_rev": t["ai_rev"][s], "rev": t["rev"][s],
                  "non_ai_rev": round(t["rev"][s] - t["ai_rev"][s], 1), "ai_share": t["ai_share"][s], "gm": t["gm"][s],
                  "eps": t["eps"][s], "capex": t["capex"][s], "pe_now": round(t["px"] / t["eps"][s], 1)} for s in TS])
print("datasets:", sorted(p.name for p in DS.iterdir()))
