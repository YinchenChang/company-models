"""收支型個股 HTML：個股 HTML 版面規格 v1.3（docs/plan/個股HTML版面規格.md「v1.3 規格」；Andy 2026-11-14 08:11；Anthropic 試作）。

章節順序：投資建議橫幅 → 1 命題與核心問題 → 2 估值 → 3 敏感度 → 4 反向 DCF → 5 損益（標準損益表，xlsx PL 頁）→ 6 驅動因子
→ 7 個別模組（資金、每 GW 收支、算力）→ 8 對帳 → 9 資料狀態。

原則（與 fundmodel_v2 相同）：
- 數字一律取自 xlsx（新節讀 dist xlsx 的具名範圍或「頁!格」；重用的收支頁分區沿用原 HTML 的 data-ref），本檔不做模型計算；
  圖表只做幾何換算（座標、刻度）。
- 讀者可見文字不放參數代號（具名範圍只放在 data-ref／title 屬性）。
- 自含：內嵌 CSS 與 inline SVG，無外部網址、無 JavaScript。

由 fundmodel_v2.build() 在 profile 設 layout="v13" 時呼叫；OpenAI 等其他 profile 不受影響。
"""
from __future__ import annotations

import html
import math
import re
from pathlib import Path

from bs4 import BeautifulSoup

YRS = [2025, 2026, 2027, 2028, 2029, 2030]
COLS = "DEFGHI"                       # 計算頁 D–I＝2025–2030
PARAM_RE = re.compile(r"[A-Za-z][A-Za-z0-9]*_[A-Za-z0-9_]+")   # 與 check_layout_v13.py 的參數代號定義一致
C_TEAL, C_ACC, C_GOLD, C_PUR, C_NEG, C_GRAY ="#10a0a7", "#0b7a80", "#c29a2b", "#6f6bae", "#b4472f", "#9aa0a0"

CSS_V13 = """
.chart{margin:10px 0 4px}.chart h4{margin:0 0 4px;font-size:15px;font-weight:600}
.chart svg{width:100%;height:auto;max-width:860px;display:block}
.lg{display:flex;flex-wrap:wrap;gap:4px 14px;font-size:13px;color:var(--muted);margin:2px 0 4px}
.lg i{display:inline-block;width:11px;height:11px;border-radius:2px;margin-right:5px;vertical-align:-1px}
.charts2{display:grid;grid-template-columns:repeat(auto-fit,minmax(min(100%,22rem),1fr));gap:12px}
table.tp td:nth-child(2),table.tp td:nth-child(5),table.tp th:nth-child(2),table.tp th:nth-child(5){text-align:left}table.pl td,table.pl th{white-space:nowrap}table.pl td:first-child,table.pl th:first-child{position:sticky;left:0;background:#fff;z-index:1}table.pl tr.sec td{position:static}table.pl tr.b td{font-weight:600}table.pl tr.r td{color:var(--muted)}
table.pl td.fc,table.pl th.fc{background:#f4f9f9}
table.pl tr.sec td{background:var(--soft);font-weight:600;color:var(--accent)}
.basis{border-left:3px solid var(--teal);background:var(--soft);padding:6px 12px;margin:6px 0 10px;font-size:14px}
table.drv td:first-child{white-space:normal;min-width:15em}table.drv tr.eq td{font-weight:600;border-top:1px solid var(--muted)}
"""


# ── 讀值 ────────────────────────────────────────────────────────────────
def at(B, ref: str):
    """「頁!格」讀值（快取值）→ (ref, value)。"""
    sheet, cell = ref.split("!")
    return ref, B.wb[sheet.strip("'")][cell.replace("$", "")].value


def yrs_named(B, name: str, years=YRS):
    return [B.yr(name, y) for y in years]


def yrs_row(B, sheet: str, row: int, years=YRS):
    return [at(B, f"{sheet}!{COLS[y - 2025]}{row}") for y in years]


def find_row(B, sheet: str, code: str) -> int:
    """計算頁 A 欄的列 ID（如 R30、F02）→ 列號。"""
    ws = B.wb[sheet]
    for r in range(1, ws.max_row + 1):
        if ws.cell(r, 1).value == code:
            return r
    raise KeyError(f"{sheet} 找不到列 {code}")


def src_docdate(B, key: str):
    """SRC_ANT 某 key 的文件日期（N 欄）→ (ref, value)。"""
    ws = B.wb["SRC_ANT"]
    for r in range(5, ws.max_row + 1):
        if ws.cell(r, 2).value == key:
            return at(B, f"SRC_ANT!N{r}")
    raise KeyError(key)


# ── 圖表（SVG；只做幾何） ────────────────────────────────────────────────
def _nice(lo: float, hi: float, n: int = 4):
    span_ = hi - lo or 1.0
    raw = span_ / n
    mag = 10 ** math.floor(math.log10(raw))
    step = min((s * mag for s in (1, 2, 2.5, 5, 10) if s * mag >= raw), default=raw)
    a, b = math.floor(lo / step) * step, math.ceil(hi / step) * step
    ticks, t = [], a
    while t <= b + step * 1e-6:
        ticks.append(round(t, 10))
        t += step
    return a, b, ticks


def _tick(v: float) -> str:
    s = f"{abs(v):,.0f}" if abs(v) >= 10 or v == 0 else f"{abs(v):,.1f}".rstrip("0").rstrip(".")
    return ("−" if v < 0 else "") + s


def bar_chart(sp, cats, series, unit, title, stacked=False, totals=None, labels=True, fmt="n1", h=250, note=""):
    """長條圖。series＝[(名稱, 顏色, [(ref, v), …])]；stacked＝堆疊；totals＝堆疊頂端標示的 [(ref, v)]（取自 xlsx，不加總）。"""
    W, L, R, T, Bm = 760, 54, 10, 30, 30
    n, k = len(cats), len(series)
    vals = [[(v if isinstance(v, (int, float)) else 0.0) for _, v in s[2]] for s in series]
    if stacked:
        pos = [sum(max(0, vals[j][i]) for j in range(k)) for i in range(n)]
        neg = [sum(min(0, vals[j][i]) for j in range(k)) for i in range(n)]
        lo, hi = min(neg + [0]), max(pos + [0])
    else:
        flat = [x for row in vals for x in row]
        lo, hi = min(flat + [0]), max(flat + [0])
    a, b, ticks = _nice(lo, hi * 1.08 if hi > 0 else hi, 4)
    ph = h - T - Bm

    def y(v):
        return T + (b - v) / (b - a) * ph
    out = [f'<svg viewBox="0 0 {W} {h}" role="img" aria-label="{html.escape(title)}" xmlns="http://www.w3.org/2000/svg">']
    for t in ticks:
        out.append(f'<line x1="{L}" x2="{W - R}" y1="{y(t):.1f}" y2="{y(t):.1f}" stroke="{"#8a8c8c" if t == 0 else "#e3e8e8"}" stroke-width="1"/>'
                   f'<text x="{L - 6}" y="{y(t) + 4:.1f}" text-anchor="end" font-size="12" fill="#5f605f">{_tick(t)}</text>')
    out.append(f'<text x="4" y="12" font-size="12" fill="#5f605f">{html.escape(unit)}</text>')
    gw = (W - L - R) / n
    bw = gw * (0.62 if stacked else 0.78) / (1 if stacked else k)
    for i, c in enumerate(cats):
        x0 = L + gw * i + (gw - (bw if stacked else bw * k)) / 2
        out.append(f'<text x="{L + gw * (i + 0.5):.1f}" y="{h - 10}" text-anchor="middle" font-size="12.5" fill="#3b3838">{html.escape(str(c))}</text>')
        up, dn = 0.0, 0.0
        for j, (_nm, col, pts) in enumerate(series):
            ref, v = pts[i]
            if not isinstance(v, (int, float)):
                continue
            if stacked:
                base = up if v >= 0 else dn
                top, bot = (base + v, base) if v >= 0 else (base, base + v)
                if v >= 0:
                    up += v
                else:
                    dn += v
                xx = x0
            else:
                top, bot = max(v, 0), min(v, 0)
                xx = x0 + bw * j
            fill = col if (v >= 0 or stacked or col != "auto") else C_NEG
            if col == "auto":
                fill = C_TEAL if v >= 0 else C_NEG
            out.append(f'<rect x="{xx:.1f}" y="{y(top):.1f}" width="{bw:.1f}" height="{max(0.5, y(bot) - y(top)):.1f}" fill="{fill}" rx="1.5"/>')
            if labels and not stacked:
                ty = y(top) - 4 if v >= 0 else y(bot) + 13
                out.append(_svgnum(sp, ref, v, fmt, xx + bw / 2, ty, 11.5 if k > 1 else 12.5))
        if stacked and totals:
            ref, v = totals[i]
            out.append(_svgnum(sp, ref, v, fmt, x0 + bw / 2, y(up) - 5, 12.5, bold=True))
    out.append("</svg>")
    lg = "".join(f'<span><i style="background:{(C_TEAL if c == "auto" else c)}"></i>{html.escape(nm_)}</span>' for nm_, c, _ in series)
    nt = f'<p class="refline">{note}</p>' if note else ""
    return f'<div class="chart"><h4>{html.escape(title)}</h4><div class="lg">{lg}</div>{"".join(out)}{nt}</div>'


def _svgnum(sp, ref, v, fmt, x, y_, size, bold=False):
    from fundmodel_v2 import fmt as f_
    t = f_(v, fmt)
    fw = 'font-weight="700" ' if bold else ""
    return (f'<text x="{x:.1f}" y="{y_:.1f}" text-anchor="middle" font-size="{size}" {fw}fill="#3b3838" '
            f'data-ref="{html.escape(ref)}" data-fmt="{fmt}" data-v="{html.escape(repr(v))}">{html.escape(t)}</text>')


# ── 橫幅與各章 ──────────────────────────────────────────────────────────
def hero_v13(B, M):
    sp = M.span
    rref, rating = B.cell("VAL_Rating")
    cls = {"強力放空": "hero-neg", "放空": "hero-neg", "中立": "hero-neu", "買進": "hero-pos", "強力買進": "hero-pos"}.get(str(rating), "hero-neu")
    w, uw = B.cell("VAL_TP_W"), B.cell("VAL_Up_W")
    c, uc = B.cell("VAL_TP_CompMain"), B.cell("VAL_Up_CompMain")
    px, lo, hi = B.cell("VAL_Px"), B.cell("VAL_PxLo"), B.cell("VAL_PxHi")
    pxd = src_docdate(B, "pros_px_mid")
    tiles = (
        f'<div class="hero-t"><div class="hero-l">目標價（DCF 與同業倍數各半）</div><div class="hero-n">{sp(*w, "n2")}</div>'
        f'<div class="hero-s">上檔空間 {sp(*uw, "sp1")}</div></div>'
        f'<div class="hero-t"><div class="hero-l">市場主流同業倍數法（未調整）</div><div class="hero-n">{sp(*c, "n2")}</div>'
        f'<div class="hero-s">上檔空間 {sp(*uc, "sp1")}</div></div>'
        f'<div class="hero-t"><div class="hero-l">現價（IPO 價格區間中點；初步招股說明書 {sp(*pxd)}；尚未掛牌）</div><div class="hero-n">{sp(*px, "n2")}</div>'
        f'<div class="hero-s" style="color:var(--muted)">區間 {sp(*lo, "n2")}–{sp(*hi, "n2")}</div></div>')
    note = (f'以 IPO 區間中點 {sp(*px, "n0")} 美元計，目標價 {sp(*w, "n2")} 美元、上檔空間 {sp(*uw, "sp1")}，評等{sp(rref, rating)}。'
            f'DCF 只有 {sp(*B.cell("VAL_TP_DCF"), "n2")} 美元，市場主流同業倍數法卻到 {sp(*c, "n2")} 美元，兩者落差是本檔核心爭議。'
            f'IPO 後現金足以撐過 2027–2028 年算力付款高峰，主要風險在算力承諾的付款時點。')
    return (f'<div class="hero {cls}"><div class="hero-rec"><span class="hero-k">投資建議</span><span class="hero-v">{sp(rref, rating)}</span></div>'
            f'<div class="hero-tiles">{tiles}</div><p class="hero-note">{note}</p></div>')


def thesis_v13(soup_s1, B, M, P):
    sp = M.span
    base = M.thesis_section(soup_s1, B, "", P)
    gap = [B.yr("COST_PropGap_VR", y) for y in YRS[1:]]
    cash = [B.yr("FND_CashEnd", y) for y in YRS[1:]]
    mn = yrs_row(B, "Funding", find_row(B, "Funding", "F28"), YRS[1:])
    ch1 = bar_chart(sp, YRS[1:], [("營收減全成本（每 VR 等值 GW）", "auto", gap)], "十億美元／GW／年",
                    "問題一：每 GW 的營收能否覆蓋全成本（正＝覆蓋）", fmt="sn1",
                    note="全成本含算力、人事與管銷、股權報酬；2025 年分母極小、失真，不列。")
    ch2 = bar_chart(sp, YRS[1:], [("年底現金（基準）", C_ACC, cash), ("最低現金（次年營運費用六個月）", C_GOLD, mn)], "十億美元",
                    "問題二：錢夠不夠（年底現金對最低現金）", fmt="n0")
    charts = f'<h3>回答核心問題的圖</h3><div class="charts2">{ch1}{ch2}</div>'
    return base.replace('<h3>重點觀察</h3>', f'{charts}<h3>重點觀察</h3>', 1)


def valuation_v13(B, M):
    sp = M.span
    sec = M.valuation_section(B)
    # 原橫幅的評等規則與邊界揭露、方法說明移入估值章（v1.3 B 節）
    rule = ('<p class="note">投資建議依加權目標價判定：上檔空間 &gt;＋50% 強力買進、&gt;＋25% 買進、≥−25% 中立、≥−50% 放空、其餘強力放空。</p>')
    rn = f'<p class="note">{sp(*B.cell("VAL_RatingNote"))}</p>' if B.has("VAL_RatingNote") else ""
    meth = B.cell("VAL_Method") if B.has("VAL_Method") else None
    if meth:
        sec = sec.replace(f'<p class="note">{sp(*meth)}</p>', "", 1)
    extra = (f'<h3>評等規則與邊界揭露</h3>{rn or rule}'
             + (f'<details class="more"><summary>展開：估值方法（模型原文）</summary><p class="note">{sp(*meth)}</p></details>' if meth else ""))
    # v1.3 G 節：估值章開頭並排四個目標價（表＋長條圖〔原 bars_html，含現價線〕）
    px, pxd = B.cell("VAL_Px"), src_docdate(B, "pros_px_mid")

    def r4(no, lab, nm_, up, note):
        u = sp(*B.cell(up), "sp1") if up and B.has(up) else "—"
        return f"<tr><td>{no}</td><td>{lab}</td><td>{sp(*B.cell(nm_), 'n2')}</td><td>{u}</td><td class=\"note\">{note}</td></tr>"
    four = ('<h3>四個目標價（美元／股）</h3><div class="tw"><table class="tp"><thead><tr><th></th><th>目標價</th><th>值</th><th>對現價上檔空間</th><th>說明</th></tr></thead><tbody>'
            + r4("1", "DCF 目標價（單純 DCF，三情境加權）", "VAL_TP_DCF", "VAL_Up_DCF", "無槓桿自由現金流折現；價值多在 2030 年後的終值")
            + r4("2", "Comparable 目標價（三情境加權；調整後，進加權）", "VAL_TP_CompAdj", None, "同業營收倍數 × 依 2030 年自由現金流率調降的係數")
            + r4("2a", "　Comparable 未調整值（市場主流做法）", "VAL_TP_CompMain", "VAL_Up_CompMain", "同業營收倍數直接套用，不調整")
            + r4("3", "加權目標價（DCF 與 Comparable 各半；橫幅主值）", "VAL_TP_W", "VAL_Up_W", "評等依此判定")
            + '<tr><td>4</td><td>市場共識目標價</td><td>不適用</td><td>—</td><td class="note">尚未掛牌，無分析師共識目標價；掛牌後補進 Val 頁（附來源、家數、日期）</td></tr>'
            + f'</tbody></table></div><p class="refline">現價 {sp(*px, "n2")} 美元＝IPO 價格區間中點（初步招股說明書 {sp(*pxd)}）。下圖黑線為現價。</p>')
    sec = sec.replace("<h3>目標價與現價</h3>", f"{four}<h3>目標價與現價（圖）</h3>", 1)
    return sec.replace("<h3>結論主要項目</h3>", f"{extra}<h3>結論主要項目</h3>", 1)


def revdcf_v13(B, M):
    sp = M.span
    tgt = yrs_named(B, "RVS_Target")
    fwd = yrs_row(B, "Reverse", find_row(B, "Reverse", "X05"))
    gap_ = yrs_named(B, "RVS_Gap")
    mult = yrs_named(B, "RVS_MultProp")
    fcf = yrs_named(B, "RVS_FCF")
    ext = yrs_named(B, "RVS_ExtNeedCum")
    head = "".join(f"<th>{y}</th>" for y in YRS)

    def tr(lab, pts, f):
        return f"<tr><td>{lab}</td>" + "".join(f"<td>{sp(r, v, f) if v is not None else '—'}</td>" for r, v in pts) + "</tr>"
    tbl = (f'<div class="tw"><table><thead><tr><th>十億美元</th>{head}</tr></thead><tbody>'
           + tr("管理層營收目標（總額口徑）", tgt, "n1") + tr("本模型營收總額（基準）", fwd, "n1") + tr("差距（目標 − 本模型）", gap_, "sn1")
           + tr("要達成目標，兩條營收線需放大的倍數", mult, "n2") + tr("照目標營收、支出不變時的自由現金流", fcf, "sn1")
           + tr("照目標營收時的累計外部資金需求", ext, "n1") + "</tbody></table></div>")
    yr = B.cell("RVS_CFPosYear") if B.has("RVS_CFPosYear") else None
    return ('<section id="revdcf"><h2>反向 DCF：現價隱含什麼</h2>'
            '<p class="note">本模型尚未建「由現價反解成長率」的反向 DCF（列為待辦）。本章先以模型既有的反向頁回答相近的問題：'
            '若營收只走到管理層對外說過的目標，而算力與營運支出照已簽承諾不變，現金與外部資金會如何——用來檢查估值所依賴的營收路徑是否已比管理層自己的說法更樂觀。</p>'
            f'{tbl}<p class="note">讀法：倍數小於 1 表示本模型基準營收已高於管理層目標（2026–2029 年皆是）；若營收只走到管理層目標而支出照承諾，2029 年會出現外部資金需求。2025 年為實際值、2030 年無管理層目標，兩年均以本模型值代入。'
            '管理層目標多早於初步招股說明書，只作對照，不回饋基準。</p></section>')


def pl_v13(B, M):
    sp = M.span
    names = [("營收", "PL_Rev", "n2", "b"), ("營收年增率（YoY）", "PL_RevYoY", "p1", "r"), ("營業成本（COGS）", "PL_COGS", "n2", ""),
             ("毛利", "PL_GP", "n2", "b"), ("毛利率", "PL_GPM", "p1", "r"), ("營業費用", "PL_Opex", "n2", ""), ("營業利益", "PL_OI", "sn2", "b"),
             ("營業利益率", "PL_OIM", "p1", "r"), ("其他收入（淨額）", "PL_Oth", "sn2", ""), ("稅前利益", "PL_PBT", "sn2", "b"),
             ("所得稅", "PL_Tax", "n2", ""), ("稅後淨利", "PL_NI", "sn2", "b"), ("淨利率", "PL_NIM", "p1", "r"), ("EPS（完全稀釋，美元）", "PL_EPS", "sn2", "b")]
    ratios = [("營收年增率", "PL_RevYoY", "p1"), ("毛利率", "PL_GPM", "p1"), ("營業利益率", "PL_OIM", "p1"), ("淨利率", "PL_NIM", "p1"), ("淨利年增率（前期虧損時不適用）", "PL_NIYoY", "p1")]
    hdr = ["2023", "2024", "2025", "2026 前九月", "2026", "2027", "2028", "2029", "2030"]
    fc = [False] * 4 + [True] * 5
    th = "".join(f'<th class="{"fc" if f else ""}">{h.replace(" ", "<br>")}{"<br>預測" if f else ""}</th>' for h, f in zip(hdr, fc))

    def tr(lab, nm_, f, cls):
        cells = []
        for (r, v), isf in zip(B.grid(nm_)[0], fc):
            c = "fc" if isf else ""
            cells.append(f'<td class="{c}">{sp(r, v, f) if v is not None else "—"}</td>')
        return f'<tr class="{cls}"><td>{lab}</td>{"".join(cells)}</tr>'
    rows = "".join(tr(*x) for x in names)
    rrows = "".join(tr(lab, nm_, f, "r") for lab, nm_, f in ratios)
    tbl = (f'<div class="tw"><table class="pl"><thead><tr><th>十億美元（GAAP）</th>{th}</tr></thead><tbody>{rows}'
           f'<tr class="sec"><td colspan="10">比率</td></tr>{rrows}</tbody></table></div>')
    sh = B.grid("PL_Shares")[0]
    yrs_all = [2023, 2024, 2025, 2026, 2027, 2028, 2029, 2030]
    idx = [0, 1, 2, 4, 5, 6, 7, 8]
    g = lambda nm_: [B.grid(nm_)[0][i] for i in idx]  # noqa: E731
    ch1 = bar_chart(sp, yrs_all, [("營收", C_TEAL, g("PL_Rev")), ("營業利益", C_ACC, g("PL_OI")), ("稅後淨利", C_PUR, g("PL_NI"))],
                    "十億美元", "營收、營業利益與稅後淨利（2023–2025 實際；2026–2030 基準情境）", fmt="n0", h=290,
                    note="2025–2026 年稅後淨利的大額虧損主要來自可轉債等金融工具的公允價值變動（非現金）；IPO 時轉換後不再發生。")
    sc = lambda pre: [B.grid(f"{pre}_{s}")[0][4:] for s in "ABC"]  # noqa: E731
    ra, rb, rc = sc("PL_Rev")
    oa, ob, oc = sc("PL_OI")
    ch2 = bar_chart(sp, YRS[1:], [("情境 A（樂觀）", C_TEAL, ra), ("情境 B（基準）", C_ACC, rb), ("情境 C（保守）", C_GOLD, rc)],
                    "十億美元", "營收：三情境", fmt="n0", labels=False)
    ch3 = bar_chart(sp, YRS[1:], [("情境 A（樂觀）", C_TEAL, oa), ("情境 B（基準）", C_ACC, ob), ("情境 C（保守）", C_GOLD, oc)],
                    "十億美元", "營業利益：三情境", fmt="n0", labels=False)

    def sctr(lab, pts, f):
        return f"<tr><td>{lab}</td>" + "".join(f"<td>{sp(r, v, f)}</td>" for r, v in pts) + "</tr>"
    sctbl = (f'<div class="tw"><table><thead><tr><th>十億美元</th>{"".join(f"<th>{y}</th>" for y in YRS[1:])}</tr></thead><tbody>'
             + sctr("營收：情境 A", ra, "n1") + sctr("營收：情境 B", rb, "n1") + sctr("營收：情境 C", rc, "n1")
             + sctr("營業利益：情境 A", oa, "sn1") + sctr("營業利益：情境 B", ob, "sn1") + sctr("營業利益：情境 C", oc, "sn1") + "</tbody></table></div>")
    basis = ('<p class="basis"><strong>口徑：GAAP。</strong>Anthropic 招股書只提供 GAAP 財報（non-GAAP 僅揭露「調整後營業損益」一項），'
             '市場與媒體引用的營收、虧損也都是 GAAP 數字，因此本表採 GAAP。2023–2025 年與 2026 年前九月為招股書實際數，2026–2030 年為本模型基準情境預測。</p>')
    notes = ('<details class="more"><summary>展開：預測列的口徑說明</summary><ul class="note">'
             '<li>營收＝營收總額：招股書把經雲端市集（AWS、Google 等）的銷售以總額認列，平台抽成列在銷售、行銷與合作夥伴費用；本模型命題用的「營收淨額」已扣抽成，兩者不同。</li>'
             '<li>營業成本＝服務客戶的推論算力成本。2025 年招股書營收成本 2.77，對應本模型推論算力 3.23，模型口徑略偏保守。</li>'
             '<li>營業利益＝營收減算力、人事與管銷、平台抽成與股權報酬；2025 年模型 −8.06 與招股書 −8.06 一致。算力以合約付款計，與 GAAP 的租賃直線認列有時間差。</li>'
             '<li>其他收入＝現金利息（假設收益率 4%）；2026 年另含前九月實際的金融工具公允價值變動 −94.3。第四季轉換前的公允價值變動未揭露，暫以 0 計。</li>'
             '<li>所得稅：2026 年取前九月實際；2027 年起以 21% 稅率、先抵用累積虧損（2025 年底聯邦虧損扣抵 8.2 加之後營業虧損）。預測期內虧損扣抵尚未用完，所得稅為 0。</li>'
             f'<li>EPS：歷史為招股書每股淨損（分母為特別股轉換前的加權股數，約 {sp(*sh[2], "n2")} 十億股）；'
             f'預測為稅後淨利除以 IPO 後完全稀釋的評價股數 {sp(*sh[4], "n2")} 十億股。兩段分母口徑不同，不宜直接比較。</li></ul></details>')
    return (f'<section id="pl"><h2>損益（標準損益表）</h2>{basis}{tbl}{notes}{ch1}'
            f'<h3>情境 A／C 並列</h3><div class="charts2">{ch2}{ch3}</div>'
            f'<details class="more"><summary>展開：三情境數字（情境 A、C 為估值頁快照值）</summary>{sctbl}</details></section>')


def drivers_v13(B, M, rev_sec_html: str):
    sp = M.span
    rr = lambda code: yrs_row(B, "Revenue", find_row(B, "Revenue", code))  # noqa: E731
    fr = lambda code: yrs_row(B, "Funding", find_row(B, "Funding", code))  # noqa: E731
    users, fee_c = yrs_named(B, "DEM_Users_PaidConsumer"), rr("R30")
    seats, fee_s = yrs_named(B, "DEM_Seats_Total"), rr("R31")
    tok, price = yrs_named(B, "DEM_Tok_API"), rr("R23")
    sub_c, sub_s, api = yrs_named(B, "REV_SubConsumer"), yrs_named(B, "REV_SubSeats"), yrs_named(B, "REV_API")
    oth, gross = rr("R50"), yrs_named(B, "REV_Gross")
    cap, gross_c = rr("R64"), yrs_named(B, "REV_GrossCapped")
    ch_share, fee_r = rr("R53"), rr("R55")
    cut, net = yrs_named(B, "REV_PartnerShareCapped"), yrs_named(B, "REV_NetCapped")
    head = "".join(f"<th>{y}</th>" for y in YRS)

    def tr(lab, pts, f, cls=""):
        return f'<tr class="{cls}"><td>{lab}</td>' + "".join(f"<td>{sp(r, v, f) if v is not None else '—'}</td>" for r, v in pts) + "</tr>"
    seg = (f'<div class="tw"><table class="drv"><thead><tr><th>分部營收推導</th>{head}</tr></thead><tbody>'
           + tr("付費個人用戶（百萬人）", users, "n2") + tr("× 個人方案加權月費（美元／月）× 12", fee_c, "n1")
           + tr("＝個人訂閱營收（Pro、Max）", sub_c, "n2", "eq")
           + tr("企業付費席位（百萬席）", seats, "n2") + tr("× 席位加權月費（美元／月）× 12", fee_s, "n1")
           + tr("＝企業席位營收（Team、Enterprise）", sub_s, "n2", "eq")
           + tr("API 計費 token（兆）", tok, "n0") + tr("× API 組合有效單價（美元／百萬 token）", price, "n2")
           + tr("＝API 營收（含經雲端市集）", api, "n1", "eq")
           + tr("＋其他營收", oth, "sn2")
           + tr("＝營收總額", gross, "n1", "eq") + tr("× 算力容量上限係數（1＝不受限）", cap, "n2")
           + tr("＝營收總額（截頂後；即損益表營收）", gross_c, "n1", "eq")
           + tr("其中經雲端通路的 API 占比", ch_share, "p1") + tr("雲端平台抽成率", fee_r, "p1")
           + tr("− 雲端平台抽成", cut, "n1") + tr("＝營收淨額（命題與現金流用）", net, "n1", "eq")
           + "</tbody></table></div>")
    ch1 = bar_chart(sp, YRS, [("API", C_TEAL, api), ("個人訂閱", C_GOLD, sub_c), ("企業席位", C_PUR, sub_s)], "十億美元",
                    "分部營收（堆疊；頂端數字＝營收總額）", stacked=True, totals=gross, fmt="n1",
                    note="「其他營收」每年不到 0.01，圖中略去；頂端數字取自模型營收總額。")
    ch2 = bar_chart(sp, YRS, [("API 計費 token（兆）", C_TEAL, tok)], "兆 token", "量：API 計費 token", fmt="n0")
    ch3 = bar_chart(sp, YRS, [("API 組合有效單價", C_ACC, price)], "美元／百萬 token", "價：API 組合有效單價", fmt="n2")
    cc, nc = fr("F02"), fr("F04")
    inf, rd = yrs_named(B, "COST_InfCompute"), yrs_named(B, "COST_RDCompute")
    sbc, fcf = yrs_named(B, "COST_SBC"), yrs_named(B, "FND_FCF")
    idle = yrs_row(B, "Cost", find_row(B, "Cost", "K88"))
    cost = (f'<div class="tw"><table class="drv"><thead><tr><th>成本與現金面（基準；十億美元）</th>{head}</tr></thead><tbody>'
            + tr("算力成本（合約付款）", cc, "n1") + tr("　其中：推論（服務客戶）", inf, "n1") + tr("　其中：研發與訓練", rd, "n1") + tr("　其中：閒置保留", idle, "n1")
            + tr("人事與管銷（不含股權報酬）", nc, "n1") + tr("股權報酬（非現金）", sbc, "n1")
            + tr("營運自由現金流", fcf, "sn1", "eq") + "</tbody></table></div>")
    ch4 = bar_chart(sp, YRS, [("推論算力", C_TEAL, inf), ("研發與訓練算力", C_PUR, rd), ("閒置保留", C_GRAY, idle)], "十億美元", "算力成本：推論對研發（堆疊；頂端＝算力成本合計）",
                    stacked=True, totals=cc, fmt="n0", note="2027–2029 年研發算力隨已簽合約大增，是損益與現金流轉負的主因；推論算力隨營收成長。")
    intro = ('<p class="note">本章依命題挑出最重要的驅動因子：營收面看「量」（API 計費 token、付費用戶與席位）與「價」（API 有效單價、月費），'
             '成本面看已簽算力合約的付款與推論／研發分配。營收九成以上來自 API，因此 API token 量與單價是估值與命題最敏感的兩個變數。</p>')
    return (f'<section id="drivers"><h2>驅動因子 → 推導量</h2>{intro}<h3>分部營收怎麼算出來</h3>{ch1}{seg}'
            f'<div class="charts2">{ch2}{ch3}</div><p class="note">量增價跌：API 有效單價逐年下降（新模型降價與年降假設），營收成長全部來自 token 量增加。</p>'
            f'<h3>成本面</h3>{ch4}{cost}'
            f'<details class="more"><summary>展開：原收支頁「營收結構與年化營收里程碑」</summary>{rev_sec_html}</details></section>')


# ── 組裝 ────────────────────────────────────────────────────────────────
def build_v13(html_in: Path, xlsx: Path, out: Path, P: dict, M):
    src = html_in.read_text(encoding="utf-8")
    soup = BeautifulSoup(src, "lxml")
    B = M.Book(xlsx)
    m = re.search(r"v(\d+\.\d+\.\d+(?:\.\d+)?)", soup.header.h1.get_text())
    ver = m.group(1) if m else "?"
    meta_div = soup.header.find("div", class_="meta")
    meta_b = meta_div.find("b")
    ver_notes = str(meta_b) if meta_b else ""
    if meta_b:
        meta_b.decompose()
    meta_txt = meta_div.decode_contents().rstrip("｜")
    secs = {s["id"]: s for s in soup.select("main > section")}
    # 重用圖表的 SVG 提示文字（<title>）去掉具名範圍代號（v1.3 B 節：讀者可見文字不放參數代號；出處仍在 data-ref 與頁!格）
    for t in soup.select("main svg title"):
        parts_ = [x for x in t.get_text().split("｜") if not PARAM_RE.fullmatch(x.strip())]
        t.string = "｜".join(parts_)
    css = M.strip_dark(soup.style.string) + M.CSS_EXTRA + CSS_V13

    parts = [thesis_v13(secs["s1"], B, M, P), valuation_v13(B, M)]
    tv, th = M.tornado_html(B)
    sens = M.regroup(secs["s6"], soup, "sens", "敏感度與「要成立必須…」", summary="展開：收支命題的關鍵驅動表與全部案例")
    sens_html = str(sens)
    if tv:
        sens_html = sens_html.replace("</h2>", f"</h2><h3>目標價敏感度（龍捲風圖）</h3>{tv}{th}<h3>收支命題的敏感度</h3>", 1)
    parts.append(sens_html)
    parts.append(revdcf_v13(B, M))
    parts.append(pl_v13(B, M))
    rev_old = M.regroup(secs["s3"], soup, "x", "營收結構")
    rev_inner = "".join(str(c) for c in [c for c in rev_old.children if getattr(c, "name", None)][1:])
    parts.append(drivers_v13(B, M, rev_inner))
    mods = [("s5", "funding", "資金：現金路徑與融資", "資金"), ("s2", "unit", "每 GW 收支：每 VR 等值 GW 的營收對全成本", "每 GW 收支"),
            ("s4", "compute", "算力：需求對已簽合約與自建供給", "算力")]
    for sid, nid, title, _nav in mods:
        parts.append(str(M.regroup(secs[sid], soup, nid, title)))
    recon = soup.new_tag("section", id="recon")
    h = soup.new_tag("h2")
    h.string = P["recon_title"]
    recon.append(h)
    for sid, lab in P["recon"]:
        r = M.regroup(secs[sid], soup, "x", lab, keep_first_table=True)
        h3 = soup.new_tag("h3")
        h3.string = lab
        recon.append(h3)
        for c in [c for c in r.children if getattr(c, "name", None)][1:]:
            recon.append(c)
    parts.append(str(recon))
    hist_ids = [k for k in secs if re.fullmatch(r"s1[a-z]", k)]
    hist = "".join(f'<details class="hist"><summary>{html.escape(secs[k].h2.get_text(" ", strip=True))}</summary>{secs[k].decode_contents()}</details>'
                   for k in hist_ids)
    status = M.regroup(secs[P["status"]], soup, "status", "資料狀態", keep_first_table=True, summary="展開：資料與版本細節")
    status_html = str(status).replace("</section>", f'<h3>附錄</h3><details class="hist"><summary>版本要點（各版變更摘要）</summary><p class="note">{ver_notes}</p></details>{hist}</section>')
    parts.append(status_html)

    navs = [("thesis", "命題與核心問題"), ("valuation", "估值"), ("sens", "敏感度"), ("revdcf", "反向 DCF"), ("pl", "損益"), ("drivers", "驅動因子")] + \
        [(nid, nv) for _, nid, _t, nv in mods] + [("recon", "對帳"), ("status", "資料狀態")]
    nav = "".join(f'<a href="#{a}">{t}</a>' for a, t in navs)
    doc = (f'<!DOCTYPE html><html lang="zh-Hant"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">'
           f'<title>{P["name"]} 個股模型 v{ver}</title><style>{css}</style></head><body><div class="wrap">'
           f'<header class="top"><h1>{P["name"]} 個股模型</h1><div class="meta">{P["meta"]}｜版本 v{ver}｜{meta_txt}</div>'
           f'<nav class="toc">{nav}</nav></header>{hero_v13(B, M)}<main class="v2">{"".join(parts)}</main>'
           f'<footer class="v2">{P["name"]} 個股模型 v{ver}｜版面 v1.3（tools/company_html/fundmodel_v13.py；個股 HTML 版面規格 v1.3 試作）｜來源 HTML <code>{html.escape(html_in.name)}</code>｜'
           f'xlsx <code>{html.escape(xlsx.name)}</code></footer></div></body></html>')
    out.write_text(doc, encoding="utf-8")
    return out
