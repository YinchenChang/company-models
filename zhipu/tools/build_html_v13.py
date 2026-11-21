#!/usr/bin/env python3
"""智譜個股頁（個股 HTML 版面規格 v1.3；v0.1.5 起）：與一頁摘要（tools/build_html.py）共用取值與情境重算。

- 數字一律經 build_html.Values 取自 LibreOffice 重算後的 xlsx（基準與 html_scenarios.yaml 各情境），每個數字帶 data-ref／data-sc／data-fmt／data-v，
  tests/test_html.py 以 engine（pycel）逐一比對（與一頁摘要同一套檢查）。本檔不做模型計算，只做排版與 SVG 幾何。
- v0.1.6 起有目標價與評等（Andy 2026-11-15）：估值章照 A–I 節（G：四個目標價與現價圖在最前；H：加權目標價相對基準低案、高案兩張主圖在敏感度章最前），
  新增反向 DCF 章；數字取自 Val、RevDCF 頁。原反向檢驗（市值隱含營收）移到估值章下方摺疊。
- 章節：命題 → 估值 → 敏感度（加權目標價；收支關卡敏感度摺疊）→ 反向 DCF → 損益（PL 頁）→ 驅動因子（分部營收推導）→ 算力 → 資金 → 與 OpenAI 並排 → 對帳（Coord 頁 C0）→ 資料狀態。

用法：python3 tools/build_html_v13.py [--outdir dist]（會先跑 build_html.build() 產生一頁摘要〔<日期>_智譜收支模型_v<版本>.html〕與 dist xlsx，再輸出 <日期>_智譜個股模型_v<版本>.html；v0.1.6 起命名對齊 MiniMax）
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
    tiles = (f'<div class="hero-t"><div class="hero-l">目標價（DCF 與同業倍數各半，港元）</div><div class="hero-n">{V.n("VAL_TP_W", None, "f2")}</div>'
             f'<div class="hero-l">上檔空間 {V.n("VAL_Up_W", None, "pct1")}</div></div>'
             f'<div class="hero-t"><div class="hero-l">現價（2026-11-20 收盤，港元）</div><div class="hero-n">{V.n("VAL_Px", None, "f2")}</div>'
             f'<div class="hero-l">市場共識目標價 {V.n("VAL_TP_Cons", None, "f0")}（2026-10-03）</div></div>'
             f'<div class="hero-t"><div class="hero-l">收支關卡：FY2026–FY2030 自由現金流合計（億元）</div><div class="hero-n">{V.n("VAL_Gate", None, "f1")}</div>'
             f'<div class="hero-l">估值評等 {V.n("VAL_Rating", None, "text")}</div></div>')
    floor = V.get("VAL_DCFFloorAll") == 1
    note = (f'同業倍數法（以同日的 MiniMax 倍數計）給 {V.n("VAL_TP_Comp", None, "f0")} 港元；DCF '
            + ("因基準情境到 FY2030 仍大幅虧損，只剩淨現金價值 " if floor else "給 ") + f'{V.n("VAL_TP_DCF", None, "f0")} 港元。'
            f'兩者各半得 {V.n("VAL_TP_W", None, "f0")} 港元，上檔空間 {V.n("VAL_Up_W", None, "pct0")}，估值評等{V.n("VAL_Rating", None, "text")}。'
            + ("收支關卡不通過（入不敷出），依規則不給買進。" if V.get("VAL_GateFail") == 1 else ""))
    return (f'<div class="hero"><div class="hero-rec"><span class="hero-k">投資建議</span><span class="hero-v">{V.n("VAL_Rec", None, "text")}</span></div>'
            f'<div class="hero-tiles">{tiles}</div><p class="hero-note">{note}</p></div>')


def thesis(V):
    items = ["每 GW 收支：智譜每 VR 等值 GW 的年營收，能否覆蓋每 GW 年全成本（算力、本地化交付、人事與其他、股權報酬）。",
             "錢夠不夠：IPO、兩次配售與可換股債券的現金，能否撐過虧損期；不夠時何時、要多少外部資金。",
             "值多少：DCF（現金流沿用本收支模型）與同業倍數（MiniMax，同日市值）各半的目標價與評等；收支關卡不通過時不給買進。"]
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


def vrow(V, pre, col="C"):
    """Val 頁 A 欄以 pre 開頭的第一列 → 格位（Val!欄列）。"""
    ws = V.wb["base"]["Val"]
    r = next(r for r in range(1, ws.max_row + 1) if str(ws.cell(r, 1).value or "").startswith(pre))
    return f"Val!{col}{r}"


def valuation(V):
    four = [("DCF 目標價（單純 DCF，三情境加權）", "VAL_TP_DCF", "VAL_Up_DCF", "無槓桿自由現金流折現；基準情境到 FY2030 仍虧損，三情境皆觸及淨現金下限"),
            ("Comparable 目標價（同業 EV ÷ 未來 12 個月營收，三情境加權）", "VAL_TP_Comp", "VAL_Up_Comp", "同業只有 MiniMax（同日 2026-11-20 市值）；未做相對倍數調整"),
            ("加權目標價（DCF 與 Comparable 各半；橫幅主值）", "VAL_TP_W", "VAL_Up_W", "評等依此判定"),
            ("市場共識目標價（Investing.com 24 位平均，2026-10-03）", "VAL_TP_Cons", "VAL_Up_Cons", "二手、早於現價日一個月，其間股價大跌，只作對照")]
    ch = bars(V, [(c, [(r, "rev")]) for c, r in (("DCF", "VAL_TP_DCF"), ("Comparable", "VAL_TP_Comp"), ("加權", "VAL_TP_W"), ("市場共識", "VAL_TP_Cons"), ("現價", "VAL_Px"))],
              "港元／股", "四個目標價與現價", "f1")
    t4 = "".join(f'<tr><td>{a}</td><td>{V.n(n, None, "f2")}</td><td>{V.n(u, None, "pct1")}</td><td class="note">{c}</td></tr>' for a, n, u, c in four)
    kv = lambda pre, f, col="C": V.n(vrow(V, pre, col), None, f)  # noqa: E731
    wacc = "".join(f'<tr><td>{lab}</td><td>{kv(pre, f)}</td></tr>' for lab, pre, f in (
        ("無風險利率", "無風險利率", "pct2"), ("β（Blume 調整後）", "β（", "f3"), ("股權風險溢酬", "股權風險溢酬", "pct1"), ("股東權益成本", "股東權益成本", "pct2"),
        ("WACC", "WACC＝", "pct2"), ("永續成長率", "永續成長率", "pct1"), ("評價股數（億股）", "評價股數", "f3"), ("淨現金（評價日 pro forma，億元）", "淨現金（評價日", "f1"),
        ("股權市值（億元）", "股權市值 E", "f0"), ("市場企業價值（億元）", "市場 EV", "f0")))
    dcf = "".join(f'<tr><td>{lab}</td>' + "".join(f"<td>{kv(pre, f, col)}</td>" for col in "CDE") + "</tr>" for lab, pre, f in (
        ("明確期現值合計（億元）", "明確期現值合計", "f1"), ("終值現值（億元）", "終值現值", "f1"), ("企業價值 EV（億元）", "企業價值 EV", "f1"),
        ("股權價值（億元）", "股權價值＝", "f1"), ("DCF 每股價值（港元）", "DCF 每股價值（目標時點；港元）", "f2"), ("Comparable 每股（港元）", "Comparable 每股（港元）", "f2"),
        ("情境目標價（港元）", "情境目標價", "f2")))
    peer = "".join(f'<tr><td>{lab}</td><td>{kv(pre, f, col)}</td></tr>' for lab, pre, f, col in (
        ("同業（MiniMax）EV ÷ 未來 12 個月共識營收（倍）", "同業 EV／NTM 營收中位數", "f2", "C"), ("智譜市場 EV ÷ 共識未來 12 個月營收（倍）", "對照：本公司市場 EV", "f2", "C"),
        ("智譜倍數 ÷ 同業（倍）", "對照：本公司倍數", "f2", "C"), ("本模型未來 12 個月營收（情境 B，億元）", "本模型 NTM 營收", "f1", "D")))
    return ('<section id="valuation"><h2>估值</h2>'
            f'<div class="figblk"><h3>四個目標價與現價（港元；現價 2026-11-20 收盤）</h3>{ch}</div>'
            '<p class="note">Andy 2026-11-15 決定智譜給目標價與評等。估值頁（xlsx Val 頁）＝DCF（無槓桿自由現金流，現金流沿用本收支模型）50%＋同業倍數 50%，'
            '三情境（2H26 API 任務數成長高／基準／低）25／50／25。折現率、同業選擇與權重為判斷類參數，依 Anthropic 估值頁樣板慣例預設，待確認。</p>'
            '<h3>四個目標價</h3><div class="tw"><table><thead><tr><th>目標價</th><th>港元</th><th>對現價上檔空間</th><th>說明</th></tr></thead><tbody>' + t4 + '</tbody></table></div>'
            f'<h3>評等規則與邊界揭露</h3><p class="note">{V.n("VAL_RatingNote", None, "text")}</p>'
            f'<p class="note">升一級所需加權目標價 {V.n("VAL_UpFlip", None, "f2")} 港元；收支關卡 FY2026–FY2030 自由現金流合計 {V.n("VAL_Gate", None, "f1")} 億元（負＝入不敷出：估值評等為買進以上時改為「收支不足」）。</p>'
            '<h3>折現率、股數與淨現金</h3><div class="tw"><table><thead><tr><th>項目</th><th>值</th></tr></thead><tbody>' + wacc + '</tbody></table></div>'
            '<p class="note"><b>稀釋未計：</b>受限制股份有揭露（招股章程附註 29(b)：2025-06-30 尚未行使 6,620 千份〔拆股前〕，1 拆 10 後約 6,620 萬股，行使價 0.10–2.00 元），'
            '未計入評價股數。若全部為增量，約占評價股數 13.6%，加權目標價約降至 138–140 港元，評等不變；這些股份是否已含在已發行股數內未查證。</p>'
            '<p class="note"><b>β 的來源：</b>智譜 2026-01 上市，自身 β 期間不足，借用 Anthropic 估值頁的美國軟體同業 5 年 β 中位數'
            '（Palantir、CrowdStrike、Cloudflare、Snowflake、Datadog；stockanalysis.com 2026-11-06），再經 Blume 調整；屬判斷類。</p>'
            '<p class="note"><b>DCF 淨現金下限：</b>三情境皆觸及下限，下限以評價日（2026-06-30 pro forma）淨現金按股東權益成本推到目標時點，'
            '未扣評價日之後的燒錢；若改用本模型目標時點（約 2027-11，年底期末值內插）的淨現金，加權目標價約 142 港元，評等不變。</p>'
            '<h3>DCF 與同業法（三情境）</h3><div class="tw"><table><thead><tr><th>項目</th><th>情境 A</th><th>情境 B</th><th>情境 C</th></tr></thead><tbody>' + dcf + '</tbody></table></div>'
            '<h3>同業倍數（2026-11-20 同日市值）</h3><div class="tw"><table><thead><tr><th>項目</th><th>值</th></tr></thead><tbody>' + peer + '</tbody></table></div>'
            '<p class="note">智譜與 MiniMax 互為唯一同業（MiniMax 模型的同業也只有智譜），兩家的 Comparable 結論本質上是同一組相對價差：一家偏貴即另一家偏便宜，不是兩個獨立的檢驗。</p>'
            f'<details class="more"><summary>展開：估值方法（模型原文）</summary><p class="note">{V.n("VAL_Method", None, "text")}</p></details>'
            '<details class="more"><summary>展開：原收支模型的反向檢驗（市值隱含營收；只作對照）</summary>' + _oldval_body(V) + '</details></section>')


def _oldval_body(V):
    rows = "".join(f'<tr><td>{lab}</td><td>{V.n(ref, None, f)}</td></tr>' for lab, ref, f in (
        ("現價：收盤價（港元，2026-11-20）", "Reverse!X38", "f0"), ("市值（人民幣億元）", "RVS_MktCap", "f0"), ("淨現金（估計，人民幣億元）", "RVS_NetCash", "f0"),
        ("企業價值 EV（人民幣億元）", "RVS_EV", "f0"), ("同業（MiniMax）企業價值 ÷ 年化營收（倍；MiniMax 市值取 2026-10-07 盤中，與智譜股價 2026-11-20 收盤不同日）", "RVS_PeerMult", "f1"),
        ("市值隱含年營收（人民幣億元）", "RVS_MktImpliedRev", "f1"), ("模型營收首次達到的年度", "RVS_MktImpliedYear", "yrt")))
    return (f'<h3>反向檢驗：模型營收（藍）vs 市值隱含年營收（紅虛線 {V.n("RVS_MktImpliedRev", None, "f1")} 億元），紫點＝分析師共識營收</h3>{chart_revcheck(V)}'
            f'<div class="tw"><table><thead><tr><th>現價、市值與反向檢驗</th><th>值</th></tr></thead><tbody>{rows}</tbody></table></div>'
            f'{year_table(V, [("正向營收 ÷ 市值隱含營收（倍）", "RVS_FwdOverImplied", "f2", ""), ("分析師共識營收（人民幣億元；FY2029–FY2030 為沿用值）", "RVS_Target", "f1", "")])}'
            f'<p class="note">口徑差異：本摺疊的淨現金 {V.n("RVS_NetCash", None, "f0")} 億元、企業價值 {V.n("RVS_EV", None, "f0")} 億元（Reverse 頁）是 2026-06-30 現金加 7、9 月配售與可換股債券，'
            f'另扣 7 月配售款至 8-31 已動用部分；估值章正文的淨現金 {V.n(vrow(V, "淨現金（評價日"), None, "f1")} 億元、市場企業價值 {V.n(vrow(V, "市場 EV"), None, "f0")} 億元是 2026-06-30 pro forma（未扣其後動用）。'
            f'同業倍數 {V.n("RVS_PeerMult", None, "f1")} 倍為 MiniMax 2026-10-07 盤中市值 ÷ 年化營收；正文 {V.n(vrow(V, "同業 EV／NTM 營收中位數"), None, "f2")} 倍為 2026-11-20 收盤市值 ÷ 未來 12 個月共識營收，兩者口徑不同。</p>'
            f'<p class="note">分析師共識只到 FY2028；FY2029–FY2030 沿用 FY2028 共識平均，乘上反向模式年增率假設 {V.n("INP_134", None, "pct0")}（基準為 0，故三年同值），不是券商預測。</p>'
            '')




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
    return ('<section id="sens"><h2>敏感度</h2>' + sens_val(V)
            + '<details class="more"><summary>展開：收支關卡的敏感度（FY2030 覆蓋率、累計外部資金需求）</summary>'
            f'<h3>FY2030 覆蓋率（營收 ÷ 全成本；基準 {V.n("COST_Coverage", 2030, "pct0")}）：各情境（中線＝基準，右綠＝改善，左紅＝惡化）</h3>'
            + "".join(big) + (f'<details class="more"><summary>展開：變動未達 10% 的情境（{len(small)} 項）</summary>{"".join(small)}</details>' if small else "")
            + '<p class="note">收支關卡的敏感度：每個情境改寫 Inputs 的低／高欄後整本重算（LibreOffice），看 FY2030 覆蓋率變動；只直接列出變動達 10% 的情境。</p>'
            f'<details class="more"><summary>展開：全部情境（覆蓋率、每 VR 等值 GW 差額、累計外部資金需求）</summary><div class="tw"><table><thead><tr><th>情境</th><th>FY2030 覆蓋率</th>'
            f'<th>FY2030 每 GW 差額（億元）</th><th>至 FY2030 累計外部資金（億元）</th></tr></thead><tbody>{tab}</tbody></table></div></details></details></section>')


def sens_val(V):
    """H 節：加權目標價相對基準（Val 第 7 節），低案、高案兩張主圖；|變動| ≥ 10% 直接列，其餘摺疊；變動為 0 的案例只列在全部案例表。"""
    ws = V.wb["base"]["Val"]
    r0 = next(r for r in range(1, ws.max_row + 1) if ws.cell(r, 1).value == "案例")
    rows, r = [], r0 + 1
    while isinstance(ws.cell(r, 2).value, (int, float)) and not str(ws.cell(r, 1).value).startswith("Sens 快照"):
        rows.append(r)
        r += 1
    mx = max([abs(ws[f"{c}{x}"].value) for x in rows for c in "FG"] + [0.1])

    def line(x, c, lc):
        v = ws[f"{c}{x}"].value
        w = abs(v) / mx * 50
        bar = f'<i class="p" style="left:50%;width:{w:.2f}%"></i>' if v >= 0 else f'<i class="n" style="left:{50 - w:.2f}%;width:{w:.2f}%"></i>'
        return (f'<div class="trw"><span>{V.n(f"Val!H{x}", None, "text")}（{V.n(f"Val!{lc}{x}", None, "text")}）</span><div class="ttr">{bar}</div>'
                f'<span>{V.n(f"Val!{c}{x}", None, "pct1")}</span></div>')

    def fig(c, lc, cap):
        nz = sorted([x for x in rows if abs(ws[f"{c}{x}"].value) > 1e-12], key=lambda x: -abs(ws[f"{c}{x}"].value))
        big = [line(x, c, lc) for x in nz if abs(ws[f"{c}{x}"].value) >= 0.10]
        small = [line(x, c, lc) for x in nz if abs(ws[f"{c}{x}"].value) < 0.10]
        body = "".join(big) if big else '<p class="note">沒有案例使加權目標價變動達 10%。</p>'
        more = f'<details class="more"><summary>展開：{cap}，變動未達 10% 的案例（{len(small)} 項）</summary>{"".join(small)}</details>' if small else ""
        return f'<div class="figblk"><h3>加權目標價相對基準：{cap}</h3>{body}{more}</div>'
    base = vrow(V, "Sens 快照的基準加權目標價", "C")
    allt = "".join(f'<tr><td>{V.n(f"Val!H{x}", None, "text")}</td><td>{V.n(f"Val!I{x}", None, "text")}</td><td>{V.n(f"Val!B{x}", None, "f2")}</td>'
                   f'<td>{V.n(f"Val!F{x}", None, "pct1")}</td><td>{V.n(f"Val!J{x}", None, "text")}</td><td>{V.n(f"Val!C{x}", None, "f2")}</td>'
                   f'<td>{V.n(f"Val!G{x}", None, "pct1")}</td></tr>' for x in rows)
    intro = (f'<p class="note">一次只動一個輸入（取低值或高值），整本模型重算，看加權目標價（基準 {V.n(base, None, "f2")} 港元）相對基準變動多少。'
             '長條依影響大小排序；只直接列出變動達 10% 的案例，其餘收在摺疊；變動為 0 的案例（DCF 觸及淨現金下限時，折現率、永續成長等不影響目標價）只列在全部案例表。</p>')
    return (fig("F", "I", "輸入取低值") + fig("G", "J", "輸入取高值") + intro
            + f'<details class="more"><summary>展開：全部案例表（{len(rows)} 項）</summary><div class="tw"><table><thead><tr><th>案例</th><th>低值設定</th><th>加權目標價</th>'
              f'<th>相對基準</th><th>高值設定</th><th>加權目標價</th><th>相對基準</th></tr></thead><tbody>{allt}</tbody></table></div></details>')


def revdcf(V):
    def tr(lab, imp, base, fi, fb, note):
        return f'<tr><td>{lab}</td><td><b>{V.n(imp, None, fi)}</b></td><td>{V.n(base, None, fb) if base else "—"}</td><td class="note">{note}</td></tr>'
    main = ('<div class="tw"><table><thead><tr><th>現價隱含什麼</th><th>現價隱含</th><th>模型基準</th><th>讀法</th></tr></thead><tbody>'
            + tr("支持現價所需的企業價值（億元）", "RVD_EVReq", "RVD_EVB", "f1", "f1", "DCF 情境 B 路徑；模型 EV 為負＝基準現金流不支持任何正價值")
            + tr("(b) 隱含終值期營業利益率（只改 2033 年）", "RVD_TermMargin", "RVD_TermMargin0", "pct1", "pct1", "其餘不變，永續期營業利益率要到這個水準")
            + tr("(c) 隱含 2031–2033 營業利益率加碼（相對 FY2030）", "RVD_ExtAdd", None, "pct0", "pct0", "對應 2033 年營業利益率 " + V.n("RVD_ExtMargin33", None, "pct0") + "（不計稅的上限讀法）")
            + tr("(d) 隱含 EV ÷ 未來 12 個月營收（本模型營收）", "RVD_CompMult", "RVD_PeerMult", "f2", "f2", "模型欄為同業（MiniMax）倍數")
            + tr("　隱含相對同業係數", "RVD_CompCoef", None, "f2", "f2", "1＝與同業同價")
            + tr("　支持現價所需的未來 12 個月營收（同業倍數；億元）", "RVD_NTMReq", "RVD_NTMB", "f1", "f1", "模型欄為本模型基準")
            + "</tbody></table></div>")
    read = ('<p class="note">讀法：以 DCF 看，基準情境到 FY2030 仍大幅虧損，要支持現價 ' + V.n("RVD_Px", None, "f2") + ' 港元，永續期營業利益率須達 '
            + V.n("RVD_TermMargin", None, "pct0") + '；以同業倍數看，現價隱含的營收倍數是同業的 ' + V.n("RVD_CompCoef", None, "f2") + ' 倍。'
            'WACC 與營收等比例放大的反解不做：基準自由現金流各年皆負，降低折現率或放大營收只會使價值更低，無解。</p>')
    chk = f'<p class="note">核對（應為 0）：(b) {V.n("RVD_TermChk", None, "f6")}；(d) {V.n("RVD_CompChk", None, "f6")}。計算在 xlsx 的 RevDCF 頁，該頁不被任何計算頁引用。</p>'
    return f'<section id="revdcf"><h2>反向 DCF：現價隱含什麼</h2>{main}{read}{chk}</section>'


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
            '<details class="more"><summary>附錄：版本要點</summary><p class="note">v0.1.8：現價與同業（MiniMax）報價改 2026-11-20 收盤（工作單 px1120）；其餘數值不變。v0.1.7：現價與同業（MiniMax）報價改 2026-11-20 收盤（工作單 px1113）；其餘數值不變。v0.1.6：估值頁 Val 與反向 DCF 頁 RevDCF（目標價、評等、收支關卡；Andy 2026-11-15 決定）；既有收支數值不變；成品檔名對齊 MiniMax（個股模型＝版面 v1.3 頁、收支模型＝一頁摘要）。v0.1.5：版面 v1.3 個股頁、PL 標準損益表、Coord 三層協調（讀取層級、C0 實績）；既有數值不變。'
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
    navs = [("thesis", "命題"), ("valuation", "估值"), ("sens", "敏感度"), ("revdcf", "反向 DCF"), ("pl", "損益"), ("drivers", "驅動因子"), ("compute", "算力"),
            ("funding", "資金"), ("peer", "與 OpenAI 並排"), ("recon", "對帳"), ("status", "資料狀態")]
    nav = "".join(f'<a href="#{a}">{t}</a>' for a, t in navs)
    body = (thesis(V) + valuation(V) + sens(V, cfg) + revdcf(V) + pl(V) + drivers(V) + mod_compute(V) + mod_funding(V) + mod_peer(V) + recon(V) + status(V, ver, model))
    return _clean_titles(f'<!DOCTYPE html><html lang="zh-Hant"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">'
            f'<title>智譜 個股頁 v{ver}</title><style>{CSS}</style></head><body><main><header><h1>智譜（智譜華章，02513.HK）個股頁</h1>'
            f'<div class="meta">版本 v{html.escape(ver)}｜{stamp}｜目標價 DCF＋同業倍數各半｜版面 v1.3</div><nav class="toc">{nav}</nav></header>'
            f'{hero(V)}{body}<footer>智譜 個股頁 v{html.escape(ver)}｜tools/build_html_v13.py｜xlsx {html.escape(model.name)}</footer></main></body></html>')


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--outdir", type=Path, default=REPO / "dist")
    a = ap.parse_args()
    h, x, V = BH.build(a.outdir)
    model = BH.current_model_path()
    stamp = model.name.split("_")[0]
    ver = re.search(r"_v([\d.]+)\.xlsx$", model.name).group(1)
    out = a.outdir / f"{stamp}_智譜個股模型_v{ver.replace('.', '_')}.html"
    out.write_text(render(V, BH.load_scenarios(), ver, stamp, model), encoding="utf-8")
    print(h)
    print(out)


if __name__ == "__main__":
    main()
