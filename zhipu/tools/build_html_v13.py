#!/usr/bin/env python3
"""智譜個股頁（個股 HTML 版面規格 v1.3；v0.1.5 起）：與一頁摘要（tools/build_html.py）共用取值與情境重算。

- 數字一律經 build_html.Values 取自 LibreOffice 重算後的 xlsx（基準與 html_scenarios.yaml 各情境），每個數字帶 data-ref／data-sc／data-fmt／data-v，
  tests/test_html.py 以 engine（pycel）逐一比對（與一頁摘要同一套檢查）。本檔不做模型計算，只做排版與 SVG 幾何。
- 本模型沒有目標價與評等（Reverse 只作反向對照），估值章標「估值：不適用（本模型未建目標價）」並帶 data-novaluation（規格 J 節 (b)：已上市但本模型未建估值；
  必列現價、市值、EV 與反向檢驗，至少一張反向檢驗圖）。
- 章節：命題 → 估值（不適用）→ 敏感度（對 2030 覆蓋率）→ 損益（PL 頁）→ 驅動因子（分部營收推導）→ 算力 → 資金 → 與 OpenAI 並排 → 對帳（Coord 頁 C0）→ 資料狀態。

用法：python3 tools/build_html_v13.py [--outdir dist]（會先跑 build_html.build() 產生一頁摘要與 dist xlsx，再輸出 <日期>_智譜個股頁_v<版本>.html）
"""
from __future__ import annotations

import argparse
import html
import re
from pathlib import Path

import build_html as BH
from build_html import YEARS, Values, chart_cash, chart_compute, chart_gw, chart_rev, chart_vs, nice_max, axis, ymap, year_table

REPO = BH.REPO
PL_YEARS = list(range(2023, 2031))
PL_COL = dict(zip(PL_YEARS, "CDEFGHIJ"))
CODE_RE = re.compile(r"（[^（）]*(?:INP_|SRC_ZP_)[^（）]*）")

CSS = re.sub(r"@media \(prefers-color-scheme: dark\)\{.*?\}\}", "", BH.CSS, flags=re.S) + r"""
main{max-width:1080px}
.hero{border:2px solid var(--line);border-radius:10px;padding:16px 18px;margin:14px 0;background:#f4f5f5;display:grid;gap:10px}
.hero-rec{display:flex;gap:14px;align-items:baseline;flex-wrap:wrap}.hero-k{color:var(--muted);font-weight:600}.hero-v{font-size:28px;font-weight:700}
.hero-tiles{display:grid;grid-template-columns:repeat(auto-fit,minmax(min(100%,14rem),1fr));gap:10px}
.hero-t{background:#fff;border:1px solid var(--line);border-radius:8px;padding:8px 12px}.hero-l{font-size:13px;color:var(--muted)}.hero-n{font-size:24px;font-weight:700}
.hero-note{margin:0;font-size:13px;color:var(--muted)}
nav.toc{display:flex;flex-wrap:wrap;gap:4px 12px;font-size:13.5px;margin:8px 0}nav.toc a{color:var(--accent);text-decoration:none}
details.more{margin:8px 0}details.more>summary{cursor:pointer;color:var(--accent)}
.trw{display:grid;grid-template-columns:minmax(120px,320px) 1fr 70px;gap:8px;align-items:center;margin:5px 0;font-size:13px}
.ttr{position:relative;height:14px;background:#ecefef;border-radius:3px}.ttr::before{content:"";position:absolute;left:50%;top:0;bottom:0;width:1px;background:#8a8c8c}
.ttr i{position:absolute;top:2px;bottom:2px;border-radius:2px}.ttr i.p{background:var(--ok)}.ttr i.n{background:var(--gap)}
.basis{border-left:3px solid var(--accent);padding:4px 10px;font-size:13.5px}
"""


def label(cfg, sc):
    return CODE_RE.sub("", cfg["scenarios"][sc]["label"]).strip()


def opt(V: Values, ref, year=None, f="f1"):
    try:
        v = V.get(ref, year)
    except ValueError:
        return "—"
    return V.n(ref, year, "text" if isinstance(v, str) else f)


def bars(V: Values, cells, unit, aria, f="f1"):
    """直條圖：cells＝[(類別, [(ref, 樣式)])]；每個 ref 是直接格位（頁!格），帶 data-ref 的數值標籤。"""
    W, H, L, R, T, B = 700, 290, 52, 10, 26, 34
    vals = [V.get(ref) for _, sers in cells for ref, _ in sers]
    hi, lo = nice_max(max(vals + [1e-9]) * 1.1), -nice_max(-min(vals + [-1e-9]) * 1.1)
    o = [f'<svg viewBox="0 0 {W} {H}" role="img" aria-label="{html.escape(aria)}">']
    o += axis(W, H, L, R, T, B, lo, hi, 4, unit)
    slot = (W - L - R) / len(cells)
    k = len(cells[0][1])
    bw = slot * 0.7 / k
    y0 = ymap(0, H, T, B, lo, hi)
    for i, (cat, sers) in enumerate(cells):
        x0 = L + slot * i + slot * 0.15
        for j, (ref, cls) in enumerate(sers):
            v = V.get(ref)
            yv = ymap(v, H, T, B, lo, hi)
            top, h_ = min(yv, y0), abs(yv - y0)
            o.append(f'<rect x="{x0 + j * bw:.1f}" y="{top:.1f}" width="{bw - 2:.1f}" height="{max(h_, 0.5):.1f}" style="fill:var(--{cls})"/>')
            o.append(V.t(ref, None, f, x=f"{x0 + j * bw + bw / 2 - 1:.1f}", y=f"{(yv - 4) if v >= 0 else (yv + 12):.1f}", text_anchor="middle", style="font-size:10.5px"))
        o.append(f'<text class="tick" x="{L + slot * i + slot / 2:.1f}" y="{H - 12}" text-anchor="middle">{html.escape(str(cat))}</text>')
    o.append("</svg>")
    return '<div class="cw">' + "\n".join(o) + "</div>"


def hero(V):
    tiles = (f'<div class="hero-t"><div class="hero-l">現價（2026-11-05 收盤，港元）</div><div class="hero-n">{V.n("Reverse!X38", None, "f0")}</div></div>'
             f'<div class="hero-t"><div class="hero-l">FY2030 覆蓋率（營收 ÷ 全成本）</div><div class="hero-n">{V.n("COST_Coverage", 2030, "pct0")}</div></div>'
             f'<div class="hero-t"><div class="hero-l">首次外部資金缺口年</div><div class="hero-n">{V.n("FND_FirstGapYear", None, "yrt")}</div></div>')
    note = (f'本模型未建目標價與評等（收支模型）。基準情境下 FY2030 營收只覆蓋全成本的 {V.n("COST_Coverage", 2030, "pct0")}，'
            f'已到位融資撐到 {V.n("FND_FirstGapYear", None, "yrt")} 年才出現外部資金缺口。')
    return (f'<div class="hero"><div class="hero-rec"><span class="hero-k">投資建議</span><span class="hero-v">本模型未建目標價與評等</span></div>'
            f'<div class="hero-tiles">{tiles}</div><p class="hero-note">{note}</p></div>')


def thesis(V):
    items = ["每 GW 收支：智譜每 VR 等值 GW 的年營收，能否覆蓋每 GW 年全成本（算力、本地化交付、人事與其他、股權報酬）。",
             "錢夠不夠：IPO、兩次配售與可換股債券的現金，能否撐過虧損期；不夠時何時、要多少外部資金。",
             "現價隱含什麼：以同業（MiniMax）倍數換算，市值要求的年營收是多少、模型哪一年達到（只作反向對照，不給目標價）。"]
    rep = (f'<ol><li>能否覆蓋？答：預測期內都不能，覆蓋率 FY2026 {V.n("COST_Coverage", 2026, "pct0")}、FY2030 {V.n("COST_Coverage", 2030, "pct0")}；'
           f'FY2030 每 VR 等值 GW 差額 {V.n("COST_PropGap_VR", 2030, "f0")} 億元／GW／年。</li>'
           f'<li>錢夠不夠？答：FY2026 年底現金 {V.n("FND_CashEnd", 2026, "f0")} 億元，{V.n("FND_FirstGapYear", None, "yrt")} 年首次需外部資金，至 FY2030 累計 {V.n("FND_ExtNeedCum", 2030, "f1")} 億元。</li>'
           f'<li>現價隱含什麼？答：企業價值 {V.n("RVS_EV", None, "f0")} 億元 ÷ 同業倍數 {V.n("RVS_PeerMult", None, "f1")} 倍＝年營收 {V.n("RVS_MktImpliedRev", None, "f1")} 億元，'
           f'模型在 {V.n("RVS_MktImpliedYear", None, "yrt")} 年達到。</li></ol>')
    obs = [f'<li><b>營收成長快於成本，但到 FY2030 仍未覆蓋：</b>營收由 FY2025 的 {V.n("REV_NetCapped", 2025, "f1")} 億元到 FY2030 的 {V.n("REV_NetCapped", 2030, "f0")}，'
           f'全成本由 {V.n("COST_FullCash", 2025, "f1")} 到 {V.n("COST_FullCash", 2030, "f0")}，差額 FY2030 仍 {V.n("COST_GapCash", 2030, "f1")}。</li>',
           f'<li><b>雲端（API 與 Coding Plan）是營收主體：</b>FY2030 雲端營收 {V.n("REV_Cloud", 2030, "f0")} 億元，本地化部署 {V.n("REV_OnPrem", 2030, "f1")}；'
           f'雲端覆蓋率 {V.n("COST_CoverageCloud", 2030, "pct0")}。</li>',
           f'<li><b>錢暫時夠：</b>2026 年已到位股權 {V.n("FND_EquityIn", 2026, "f0")} 億元，現金到 {V.n("FND_FirstGapYear", None, "yrt")} 年才不足。</li>']
    return ('<section id="thesis"><h2>命題與核心問題</h2>'
            '<p class="lead">命題：智譜每 VR 等值 GW 的年營收能否覆蓋每 GW 年全成本；若不能，缺口要多少外部資金、由誰以什麼條件提供。</p>'
            f'<ol>{"".join(f"<li>{html.escape(t)}</li>" for t in items)}</ol>'
            f'<details class="more"><summary>展開：命題回覆（依本模型計算的答案）</summary>{rep}</details>'
            f'<h3>每 VR 等值 GW：營收（藍）vs 全成本（橘），數字為差額</h3><div class="cw">{chart_gw(V)}</div>'
            f'<h3>重點觀察</h3><ol>{"".join(obs)}</ol></section>')


def chart_revcheck(V):
    """反向檢驗圖：模型營收（直條）vs 市值隱含年營收（虛線）與分析師共識（點）；數字全取自 xlsx。"""
    W, H, L, R, T, B = 700, 300, 52, 10, 26, 34
    imp = V.get("RVS_MktImpliedRev")
    rev = [V.get("REV_NetCapped", y) for y in YEARS]
    tgt = [V.get("RVS_Target", y) for y in YEARS]
    ymax = nice_max(max(rev + tgt + [imp]) * 1.1)
    o = [f'<svg viewBox="0 0 {W} {H}" role="img" aria-label="反向檢驗：模型營收對市值隱含年營收">']
    o += axis(W, H, L, R, T, B, 0, ymax, 4, "RMB 億")
    slot = (W - L - R) / len(YEARS)
    bw = slot * 0.42
    for i, y in enumerate(YEARS):
        xc = L + slot * i + slot / 2
        yv, y0 = ymap(rev[i], H, T, B, 0, ymax), ymap(0, H, T, B, 0, ymax)
        o.append(f'<rect x="{xc - bw / 2:.1f}" y="{yv:.1f}" width="{bw:.1f}" height="{max(y0 - yv, 0.5):.1f}" style="fill:var(--rev)"/>')
        o.append(V.t("REV_NetCapped", y, "f1", x=f"{xc:.1f}", y=f"{yv - 5:.1f}", text_anchor="middle", style="font-size:10.5px"))
        if y >= 2026:
            yt = ymap(tgt[i], H, T, B, 0, ymax)
            o.append(f'<circle cx="{xc + bw / 2 + 8:.1f}" cy="{yt:.1f}" r="4" style="fill:var(--tgt)"/>')
        o.append(f'<text class="tick" x="{xc:.1f}" y="{H - 12}" text-anchor="middle">{y}</text>')
    yi = ymap(imp, H, T, B, 0, ymax)
    o.append(f'<line x1="{L}" x2="{W - R}" y1="{yi:.1f}" y2="{yi:.1f}" style="stroke:var(--gap);stroke-width:2;stroke-dasharray:6 4"/>')
    o.append(V.t("RVS_MktImpliedRev", None, "f1", x=f"{L + 6}", y=f"{yi - 6:.1f}", style="font-size:11px;fill:var(--gap)"))
    o.append("</svg>")
    return '<div class="cw">' + "\n".join(o) + "</div>"


def valuation(V):
    rows = "".join(f'<tr><td>{lab}</td><td>{V.n(ref, None, f)}</td></tr>' for lab, ref, f in (
        ("現價：收盤價（港元，2026-11-05）", "Reverse!X38", "f0"), ("市值（人民幣億元）", "RVS_MktCap", "f0"), ("淨現金（估計，人民幣億元）", "RVS_NetCash", "f0"),
        ("企業價值 EV（人民幣億元）", "RVS_EV", "f0"), ("同業（MiniMax）企業價值 ÷ 年化營收（倍；MiniMax 市值取 2026-10-07 盤中，與智譜股價 2026-11-05 收盤不同日）", "RVS_PeerMult", "f1"),
        ("市值隱含年營收（人民幣億元）", "RVS_MktImpliedRev", "f1"), ("模型營收首次達到的年度", "RVS_MktImpliedYear", "yrt")))
    return ('<section id="valuation" data-novaluation="1"><h2>估值：不適用（本模型未建目標價）</h2>'
            '<p class="note">本模型未建目標價與評等。智譜已在港股上市、有市場價格，但本模型是收支模型：沒有 DCF、同業倍數加權與評等，因此不給目標價。'
            '版面規格 v1.3 的四個目標價圖、加權目標價敏感度與反向 DCF 章不適用（規格 J 節 (b)）；本章改列現價、市值、企業價值與反向檢驗（市值隱含營收）。'
            '若要給目標價需另建估值頁（判斷類，待 Andy 決定）。</p>'
            f'<h3>反向檢驗：模型營收（藍）vs 市值隱含年營收（紅虛線 {V.n("RVS_MktImpliedRev", None, "f1")} 億元），紫點＝分析師共識營收</h3>{chart_revcheck(V)}'
            f'<div class="tw"><table><thead><tr><th>現價、市值與反向檢驗</th><th>值</th></tr></thead><tbody>{rows}</tbody></table></div>'
            f'{year_table(V, [("正向營收 ÷ 市值隱含營收（倍）", "RVS_FwdOverImplied", "f2", ""), ("分析師共識營收（人民幣億元；FY2029–FY2030 為沿用值）", "RVS_Target", "f1", "")])}'
            f'<p class="note">分析師共識只到 FY2028；FY2029–FY2030 沿用 FY2028 共識平均，乘上反向模式年增率假設 {V.n("INP_134", None, "pct0")}（基準為 0，故三年同值），不是券商預測。</p>'
            '</section>')


def sens(V, cfg):
    scs = [s for s in cfg["scenarios"] if s != "base" and not s.startswith("flip_")]
    b = V.get("COST_Coverage", 2030)
    rows = sorted(((s, V.get("COST_Coverage", 2030, s)) for s in scs), key=lambda x: -abs(x[1] - b))
    mx = max([abs(v - b) for _, v in rows] + [1e-9])

    def line(s, v):
        d, w = v - b, abs(v - b) / mx * 50
        bar = (f'<i class="p" style="left:50%;width:{w:.2f}%"></i>' if d >= 0 else f'<i class="n" style="left:{50 - w:.2f}%;width:{w:.2f}%"></i>')
        return f'<div class="trw"><span>{html.escape(label(cfg, s))}</span><div class="ttr">{bar}</div><span>{V.n("COST_Coverage", 2030, "pct0", s)}</span></div>'
    big = [line(s, v) for s, v in rows if abs(v - b) / b >= 0.10]
    small = [line(s, v) for s, v in rows if abs(v - b) / b < 0.10]
    tab = "".join(f'<tr><td>{html.escape(label(cfg, s))}</td><td>{V.n("COST_Coverage", 2030, "pct0", s)}</td><td>{V.n("COST_PropGap_VR", 2030, "f0", s)}</td>'
                  f'<td>{V.n("FND_ExtNeedCum", 2030, "f1", s)}</td></tr>' for s in ["base"] + scs)
    return ('<section id="sens"><h2>敏感度</h2>'
            f'<h3>FY2030 覆蓋率（營收 ÷ 全成本；基準 {V.n("COST_Coverage", 2030, "pct0")}）：各情境（中線＝基準，右綠＝改善，左紅＝惡化）</h3>'
            + "".join(big) + (f'<details class="more"><summary>展開：變動未達 10% 的情境（{len(small)} 項）</summary>{"".join(small)}</details>' if small else "")
            + '<p class="note">本模型沒有目標價，因此敏感度改看收支關卡：每個情境改寫 Inputs 的低／高欄後整本重算（LibreOffice），看 FY2030 覆蓋率變動；只直接列出變動達 10% 的情境。</p>'
            f'<details class="more"><summary>展開：全部情境（覆蓋率、每 VR 等值 GW 差額、累計外部資金需求）</summary><div class="tw"><table><thead><tr><th>情境</th><th>FY2030 覆蓋率</th>'
            f'<th>FY2030 每 GW 差額（億元）</th><th>至 FY2030 累計外部資金（億元）</th></tr></thead><tbody>{tab}</tbody></table></div></details></section>')


def pl(V):
    wb = V.wb["base"]["PL"]
    names = []
    for r in range(6, 21):
        lab = wb.cell(r, 1).value
        if lab:
            names.append((r, lab))
    th = "".join(f"<th>FY{y}{'<br>預測' if y >= 2026 else ''}</th>" for y in PL_YEARS)
    trs = []
    for r, lab in names:
        f = "pct1" if ("率" in lab) else ("f2" if "EPS" in lab else "f2")
        cells = "".join(f"<td>{opt(V, f'PL!{PL_COL[y]}{r}', None, f)}</td>" for y in PL_YEARS)
        cls = "em" if lab.startswith(("營收", "毛利", "營業利益", "稅前", "稅後")) and "率" not in lab else ""
        trs.append(f'<tr class="{cls}"><td>{V.n(f"PL!A{r}", None, "text")}</td>{cells}</tr>')
    rr = next(r for r, l in names if l == "營收")
    ro = next(r for r, l in names if l.startswith("營業利益（"))
    ch = bars(V, [(f"FY{y}", [(f"PL!{PL_COL[y]}{rr}", "rev"), (f"PL!{PL_COL[y]}{ro}", "gap")]) for y in PL_YEARS], "RMB 億", "營收與營業利益", "f1")
    return ('<section id="pl"><h2>損益（標準損益表）</h2>'
            '<p class="basis"><strong>口徑：Non-GAAP（非 IFRS 經調整），人民幣億元。</strong>公司業績公告以經調整淨虧損為主要指標；IFRS 淨虧損含上市前投資人工具賬面值變動，不適合比較。'
            'FY2023–FY2025 為招股章程與年度業績公告實績；FY2026 起為本模型基準情境。歷史的稅前與所得稅公司未分開揭露，標「未揭露」；預測期本模型不計稅，標「未建模」。</p>'
            f'<h3>營收（藍）與營業利益（紅）</h3>{ch}'
            f'<div class="tw"><table><thead><tr><th>人民幣億元（Non-GAAP）</th>{th}</tr></thead><tbody>{"".join(trs)}</tbody></table></div></section>')


def drivers(V):
    t = year_table(V, [("API 按量計費 token（兆）", "DEM_Tok_API", "f0", ""), ("× API 有效單價（元／百萬 token）", "REV_ApiPrice", "f2", ""),
                       ("＝開放平台按量計費營收", "REV_API_PayGo", "f1", "em"), ("Coding Plan 訂閱者（萬人）", "DEM_Subs_CP", "f1", ""),
                       ("× 每訂閱者年營收（元）", "REV_CP_ARPU", "f0", ""), ("＝Coding Plan 訂閱營收", "REV_Sub", "f1", "em"),
                       ("＝雲端部署（開放平台及 API）", "REV_Cloud", "f1", "em"), ("＋本地化部署（智能體、通用大模型、技術服務）", "REV_OnPrem", "f1", ""),
                       ("＝營收淨額（截頂後）", "REV_NetCapped", "f1", "em sep")])
    return ('<section id="drivers"><h2>驅動因子 → 推導量</h2>'
            '<p class="note">雲端營收＝API 按量計費 token × 有效單價＋Coding Plan 訂閱者 × 每訂閱者營收；本地化部署依 2025 實績延伸。下圖為營收結構（API、訂閱、廣告堆疊）與分析師共識（紫菱形）。</p>'
            f'<h3>分部營收怎麼算出來</h3><div class="cw">{chart_rev(V)}</div>{t}</section>')


def mod_compute(V):
    t = year_table(V, [("推論 GW（有效）", "CMP_InfGW_Eff", "f3", ""), ("研發 GW", "CMP_RDGW", "f3", ""), ("供給 GW", "CMP_SupplyGW", "f3", "em"),
                       ("算力成本（現金口徑）", "COST_Compute", "f1", ""), ("覆蓋率", "COST_Coverage", "pct0", "em")])
    return f'<section id="compute"><h2>算力：供給 GW 與用途</h2><div class="cw">{chart_compute(V)}</div>{t}</section>'


def mod_funding(V):
    t = year_table(V, [("自由現金流", "FND_FCF", "f1", ""), ("期末現金", "FND_CashEnd", "f0", "em"), ("最低現金", "FND_MinCash", "f0", ""),
                       ("累計外部資金需求", "FND_ExtNeedCum", "f1", "em")])
    return f'<section id="funding"><h2>資金：現金與外部資金需求</h2><p class="note">年底現金：基準（藍）與可換股債券轉股情境（金）；虛線＝最低現金。</p><div class="cw">{chart_cash(V)}</div>{t}</section>'


def mod_peer(V):
    return (f'<section id="peer"><h2>與 OpenAI 並排（每 VR 等值 GW 差額，十億美元）</h2><p class="note">紅＝智譜，灰＝OpenAI 收支模型 v0.6.2（快照，只對照）。</p>'
            f'<div class="cw">{chart_vs(V)}</div></section>')


def recon(V):
    ws = V.wb["base"]["Coord"]
    r0 = next(r for r in range(1, ws.max_row + 1) if str(ws.cell(r, 1).value or "").startswith("C. C0"))
    trs = []
    r = r0 + 2
    while ws.cell(r, 1).value:
        c = lambda col, f="f2": opt(V, f"Coord!{col}{r}", None, f)  # noqa: E731
        trs.append(f'<tr><td>{c("A", "text")}</td><td>{c("B", "text")}</td><td>{c("C")}</td><td>{c("D")}</td><td>{c("E", "pct1")}</td>'
                   f'<td>{c("F")}</td><td>{c("G")}</td><td>{c("H", "pct1")}</td><td><b>{c("I", "text")}</b></td></tr>')
        r += 1
    rs = next(x for x in range(r, ws.max_row + 1) if ws.cell(x, 1).value == "可判定項目數（對帳中、時差邊界或一致）")
    sm = [V.n(f"Coord!C{rs + k}", None, "f0") for k in range(5)]
    rv = next(x for x in range(rs, ws.max_row + 1) if ws.cell(x, 1).value == "C0 結論")
    return ('<section id="recon"><h2>對帳</h2><h3>三層協調：C0（實績）公司實績 vs 層級隱含值</h3>'
            '<p class="note">個股側＝公司法定揭露（招股章程、年度與中期業績公告）；層級側＝Tokenomics 每 GW 營收 × Wattonomics 該玩家 GW、Wattonomics 該玩家資本支出。'
            'Wattonomics 的玩家名單沒有智譜（中國只有合計），層級側沒有值。兩讀數都超過門檻且同號才標「對帳中」。單位：人民幣億元、GW。</p>'
            '<div class="tw"><table><thead><tr><th>項目</th><th>單位</th><th>公司實績（主）</th><th>層級隱含（主）</th><th>差距</th><th>公司實績（第二）</th><th>層級隱含（第二）</th>'
            f'<th>第二讀數差距</th><th>判定</th></tr></thead><tbody>{"".join(trs)}</tbody></table></div>'
            f'<p class="note">判定摘要：可判定 {sm[0]} 項、對帳中 {sm[1]} 項、時差邊界 {sm[2]} 項、參考（同源校準） {sm[3]} 項、無可比實績 {sm[4]} 項。</p>'
            f'<p class="note">{V.n(f"Coord!C{rv}", None, "text")}</p>'
            '<h3>讀取層級（同源，不判定）</h3><p class="note">本模型只讀 Tokenomics 快照（算力物理、持有成本與價格參考）；不讀 Wattonomics，Wattonomics 也不讀本模型；Siliconomics 不適用（算力買方）。'
            '版本與檔名見 Coord 頁讀取層級節。</p></section>')


def status(V, ver, model):
    return ('<section id="status"><h2>資料狀態</h2><dl>'
            f'<dt>版本</dt><dd>v{html.escape(ver)}（{html.escape(model.name)}）</dd>'
            '<dt>公司資料</dt><dd>招股章程（2025-12-30）、2025 年度業績公告（2026-03-31）、2026 中期業績公告（2026-08-31）；SRC_ZP 頁逐筆列出處、標記與一手／二手。</dd>'
            '<dt>口徑</dt><dd>人民幣億元；曆年制；GW＝IT 關鍵電力。</dd></dl>'
            '<details class="more"><summary>附錄：版本要點</summary><p class="note">v0.1.5：版面 v1.3 個股頁、PL 標準損益表、Coord 三層協調（讀取層級、C0 實績）；既有數值不變。'
            'v0.1.4：現價改 2026-11-05 收盤。v0.1.2：官方分部對帳。</p></details></section>')


PARAM_RE = re.compile(r"[A-Za-z][A-Za-z0-9]*_[A-Za-z0-9_]+")


def _clean_titles(doc: str) -> str:
    """SVG 提示文字（<title>）去掉具名範圍代號（讀者可見文字不放參數代號；data-ref 屬性保留完整來源）。"""
    def f(m):
        parts = [x for x in m.group(1).split("｜") if not PARAM_RE.fullmatch(x.strip())]
        return "<title>" + "｜".join(parts) + "</title>"
    head, sep, body = doc.partition("<body>")
    return head + sep + re.sub(r"<title>(.*?)</title>", f, body)


def render(V, cfg, ver, stamp, model):
    navs = [("thesis", "命題"), ("valuation", "估值（不適用）"), ("sens", "敏感度"), ("pl", "損益"), ("drivers", "驅動因子"), ("compute", "算力"),
            ("funding", "資金"), ("peer", "與 OpenAI 並排"), ("recon", "對帳"), ("status", "資料狀態")]
    nav = "".join(f'<a href="#{a}">{t}</a>' for a, t in navs)
    body = (thesis(V) + valuation(V) + sens(V, cfg) + pl(V) + drivers(V) + mod_compute(V) + mod_funding(V) + mod_peer(V) + recon(V) + status(V, ver, model))
    return _clean_titles(f'<!DOCTYPE html><html lang="zh-Hant"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">'
            f'<title>智譜 個股頁 v{ver}</title><style>{CSS}</style></head><body><main><header><h1>智譜（智譜華章，02513.HK）個股頁</h1>'
            f'<div class="meta">版本 v{html.escape(ver)}｜{stamp}｜收支模型（未建目標價）｜版面 v1.3</div><nav class="toc">{nav}</nav></header>'
            f'{hero(V)}{body}<footer>智譜 個股頁 v{html.escape(ver)}｜tools/build_html_v13.py｜xlsx {html.escape(model.name)}</footer></main></body></html>')


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--outdir", type=Path, default=REPO / "dist")
    a = ap.parse_args()
    h, x, V = BH.build(a.outdir)
    model = BH.current_model_path()
    stamp = model.name.split("_")[0]
    ver = re.search(r"_v([\d.]+)\.xlsx$", model.name).group(1)
    out = a.outdir / f"{stamp}_智譜個股頁_v{ver.replace('.', '_')}.html"
    out.write_text(render(V, BH.load_scenarios(), ver, stamp, model), encoding="utf-8")
    print(h)
    print(out)


if __name__ == "__main__":
    main()
