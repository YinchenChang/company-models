#!/usr/bin/env python3
"""個股 HTML 版面 v1.3「未上市、無市場價格的模型」生成器（規格 J 節；OpenAI 使用）。

與 fundmodel_v13.py（有估值頁的 v1.3）的差別：
  - 估值章＝「估值：不適用（未上市、無市場價格）」，帶 data-unlisted="1"（check_layout_v13.py 據此豁免 G、H 位置與反向 DCF 章）；不給目標價與評等。
  - 無反向 DCF 章；敏感度章改畫「對收支關卡（累計外部資金需求）」的案例長條圖（讀 s6 表，數字全取自 xlsx）。
  - 損益章讀 PL 頁具名範圍（PL_*）；歷史未揭露的格顯示「未揭露」，未建模的格顯示「未建模」。
  - 驅動因子章：分部營收推導（人數 × ARPU、token × 單價）＋成本與資金面。
數字只取自 xlsx（本檔不做模型計算，只做圖的幾何）。圖表函式沿用 fundmodel_v13.bar_chart。
"""
from __future__ import annotations

import html
import re
from pathlib import Path

from bs4 import BeautifulSoup

from fundmodel_v13 import (C_ACC, C_GOLD, C_GRAY, C_PUR, C_TEAL, CSS_V13, PARAM_RE, YRS, bar_chart)

YRS_F = YRS[1:]   # 2026–2030


def _yrs(B, name, years=YRS):
    return [B.yr(name, y) for y in years]


# ── 橫幅 ────────────────────────────────────────────────────────────────
def hero_v13u(B, M):
    sp = M.span
    g30 = B.yr("COST_PropGap_VR", 2030)
    cov30 = B.yr("COST_Coverage", 2030)
    ext = B.yr("FND_ExtNeedCum", 2030)
    tiles = (
        f'<div class="hero-t"><div class="hero-l">2030 年每 VR 等值 GW 營收減全成本（十億美元／年）</div><div class="hero-n">{sp(*g30, "sn1")}</div>'
        f'<div class="hero-s">覆蓋率 {sp(*cov30, "p0")}</div></div>'
        f'<div class="hero-t"><div class="hero-l">2026–2030 累計外部資金需求（十億美元）</div><div class="hero-n">{sp(*ext, "n1")}</div>'
        f'<div class="hero-s">2027 年起每年有缺口</div></div>'
        f'<div class="hero-t"><div class="hero-l">已到位融資（2026-03 輪，十億美元）</div><div class="hero-n">{sp(*B.yr("FND_Committed", 2026), "n0")}</div>'
        f'<div class="hero-s" style="color:var(--muted)">不含 Amazon 條件式部分</div></div>')
    note = (f'未上市、無市場價格，不給目標價與評等。基準情境下 2030 年營收仍不足以覆蓋全成本（每 VR 等值 GW 差額 {sp(*g30, "sn1")} 十億美元），'
            f'累計外部資金需求 {sp(*ext, "n1")} 十億美元。')
    return (f'<div class="hero hero-na"><div class="hero-rec"><span class="hero-k">投資建議</span><span class="hero-v">未上市，不給評等</span></div>'
            f'<div class="hero-tiles">{tiles}</div><p class="hero-note">{note}</p></div>')


# ── 命題 ────────────────────────────────────────────────────────────────
def thesis_v13u(soup_s1, B, M, P):
    sp = M.span
    base = M.thesis_section(soup_s1, B, "", P)
    gap = _yrs(B, "COST_PropGap_VR", YRS_F)
    no_ext = _yrs(B, "FND_CashNoExt", YRS_F)
    mn = _yrs(B, "FND_MinCash", YRS_F)
    ch1 = bar_chart(sp, YRS_F, [("營收減全成本（每 VR 等值 GW）", "auto", gap)], "十億美元／GW／年",
                    "問題一：每 GW 的營收能否覆蓋全成本（正＝覆蓋）", fmt="sn1",
                    note="全成本含算力、人事與管銷、股權報酬；2025 年分母極小、失真，不列。")
    ch2 = bar_chart(sp, YRS_F, [("只靠已到位融資的年底現金", C_ACC, no_ext), ("最低現金", C_GOLD, mn)], "十億美元",
                    "問題二：錢夠不夠（不另募資的年底現金對最低現金）", fmt="sn0",
                    note="低於最低現金的差額就是外部資金需求；2026 年因 2026-03 輪到位而有餘裕，2027 年起轉負。")
    charts = f'<h3>回答核心問題的圖</h3><div class="charts2">{ch1}{ch2}</div>'
    return base.replace('<h3>重點觀察</h3>', f'{charts}<h3>重點觀察</h3>', 1)


# ── 估值（不適用） ───────────────────────────────────────────────────────
def valuation_v13u(B, M):
    sp = M.span
    return ('<section id="valuation" data-unlisted="1"><h2>估值：不適用（未上市、無市場價格）</h2>'
            '<p class="note">OpenAI 未上市，沒有公開市場價格，也沒有分析師共識目標價，因此本模型不做 DCF 與同業倍數估值，不給目標價與投資評等。'
            '版面規格 J 節（未上市、無市場價格的模型）規定本章改為此說明；反向 DCF 章同樣不適用。</p>'
            f'<p class="note">可列的參考：2026-03 輪已到位的股權融資 {sp(*B.yr("FND_Committed", 2026), "n0")} 十億美元（Amazon、SoftBank、Nvidia 現金部分與其他；來源與組成見資金章）。'
            '融資輪隱含估值屬參考資訊，本模型未入模、不據以推導目標價。</p>'
            '<p class="note">本模型以收支與資金缺口取代估值：命題章回答營收能否覆蓋全成本，敏感度章看哪些假設使累計外部資金需求最大幅變動。'
            '掛牌或有公開價格後，再依 DCF 標準規格加入估值頁。</p></section>')


# ── 敏感度（對累計外部資金需求） ─────────────────────────────────────────
def sens_v13u(B, M, s6_html: str):
    sp = M.span
    soup = BeautifulSoup(s6_html, "html.parser")
    tbl = soup.find("table")
    rows = []
    for tr in tbl.find("tbody").find_all("tr"):
        tds = tr.find_all("td")
        if len(tds) < 5:
            continue
        s = tds[4].find("span")
        if s is None or s.get("data-v") is None:
            continue
        rows.append((tds[0].get_text(strip=True), s.get("data-ref"), float(s["data-v"]), s))
    base_row = next((x for x in rows if x[0] == "基準"), rows[0])
    bv = base_row[2]
    dev = [abs(x[2] - bv) for x in rows if x is not base_row]
    mx = max(dev + [1e-9])
    cases = sorted([x for x in rows if x is not base_row], key=lambda x: -abs(x[2] - bv))

    def line(x):
        d = x[2] - bv
        w = abs(d) / mx * 50
        bar = (f'<i class="n" style="left:50%;width:{w:.2f}%"></i>' if d >= 0 else f'<i class="p" style="left:{50 - w:.2f}%;width:{w:.2f}%"></i>')
        return f'<div class="trw"><span>{html.escape(x[0])}</span><div class="ttr">{bar}</div><span class="tvv">{str(x[3])}</span></div>'
    big = [line(x) for x in cases if abs(x[2] - bv) / bv >= 0.10]
    small = [line(x) for x in cases if abs(x[2] - bv) / bv < 0.10]
    head = (f'<h3>2026–2030 累計外部資金需求：各案例（基準 {str(base_row[3])} 十億美元；中線＝基準，右紅＝需求增加，左綠＝需求減少）</h3>')
    body = "".join(big) if big else '<p class="note">沒有案例使累計外部資金需求變動達 10%。</p>'
    more = (f'<details class="more"><summary>展開：變動未達 10% 的案例（{len(small)} 項）</summary>{"".join(small)}</details>' if small else "")
    intro = ('<p class="note">未上市、無目標價，因此敏感度改看「收支關卡」：一次只動一個輸入（或組合），整本模型重算，看 2026–2030 累計外部資金需求（資金缺口）變動多少。'
             '長條依影響大小排序；只直接列出變動達 10% 的案例，其餘收在摺疊；下方摺疊內有每 GW 差額、覆蓋率與「要成立必須…」。</p>')
    return ('<section id="sens"><h2>敏感度</h2><div class="figblk">' + head + body + more + '</div>' + intro
            + f'<details class="more"><summary>展開：收支命題的敏感度（每 GW 差額、外部資金需求與「要成立必須…」）</summary>{s6_html}</details></section>')


# ── 損益 ────────────────────────────────────────────────────────────────
def pl_v13u(B, M):
    sp = M.span
    names = [("營收", "PL_Rev", "n2", "b"), ("營收年增率（YoY）", "PL_RevYoY", "p1", "r"), ("營業成本（COGS）", "PL_COGS", "n2", ""),
             ("毛利", "PL_GP", "n2", "b"), ("毛利率", "PL_GPM", "p1", "r"), ("營業費用", "PL_Opex", "n2", ""), ("營業利益", "PL_OI", "sn2", "b"),
             ("營業利益率", "PL_OIM", "p1", "r"), ("其他收入（淨額）", "PL_Oth", "sn2", ""), ("稅前利益", "PL_PBT", "sn2", "b"),
             ("所得稅", "PL_Tax", "n2", ""), ("稅後淨利", "PL_NI", "sn2", "b"), ("淨利率", "PL_NIM", "p1", "r"), ("EPS（美元）", "PL_EPS", "sn2", "b"),
             ("淨利年增率", "PL_NIYoY", "p1", "r")]
    hdr = [f"FY{y}" for y in range(2023, 2031)]
    fc = [y >= 2026 for y in range(2023, 2031)]
    th = "".join(f'<th class="{"fc" if f else ""}">{h}{"<br>預測" if f else ""}</th>' for h, f in zip(hdr, fc))

    def tr(lab, nm_, f, cls):
        cells = []
        for (r, v), isf in zip(B.grid(nm_)[0], fc):
            cells.append(f'<td class="{"fc" if isf else ""}">{sp(r, v, f) if v is not None else "—"}</td>')
        return f'<tr class="{cls}"><td>{lab}</td>{"".join(cells)}</tr>'
    rows = "".join(tr(*x) for x in names)
    tbl = (f'<div class="tw"><table class="pl"><thead><tr><th>十億美元（GAAP）</th>{th}</tr></thead><tbody>{rows}</tbody></table></div>')
    basis = ('<p class="basis"><strong>口徑：GAAP（外流財報）。</strong>歷史只放公司（外流經查核財報，二手轉載）有揭露者：FY2024 只有營收；FY2025 有營收、營業成本、研發、銷售、管理費用與營業損失。'
             '未揭露的格標「未揭露」，不補編；FY2026–FY2030 為本模型基準情境，其他收入、稅前、所得稅、稅後淨利、EPS 本模型未建，標「未建模」（不補假設）。</p>')
    notes = ('<details class="more"><summary>展開：預測列的口徑說明</summary><ul class="note">'
             '<li>營收＝營收總額（截頂後，未扣 Microsoft 分成）；營業成本＝推論算力成本；營業費用＝研發算力＋閒置算力＋非算力成本（含股權報酬）＋Microsoft 分成。</li>'
             '<li>營業利益因此等於 Cost 頁「差額（含股權報酬；現金口徑）」；PL 頁的核對列（預測營業利益對該差額，應為 0；FY2025 營業損失對外流財報）列在 Checks。</li>'
             '<li>算力成本取合約付款（現金口徑），與 GAAP 租賃認列有時間差；Microsoft 分成在 GAAP 的列示位置未揭露，本表併入營業費用。</li></ul></details>')
    idx = list(range(1, 8))   # FY2024–FY2030（FY2023 未揭露，不畫）
    g = lambda nm_: [B.grid(nm_)[0][i] for i in idx]  # noqa: E731
    cats = [2024 + i for i in range(7)]
    ch1 = bar_chart(sp, cats, [("營收", C_TEAL, g("PL_Rev")), ("營業利益", C_ACC, g("PL_OI"))], "十億美元",
                    "營收與營業利益（FY2024–FY2025 為外流財報；FY2026–FY2030 為基準情境）", fmt="sn0", h=290,
                    note="FY2024 營業利益未揭露，不畫；營業損失在 2027 年後擴大，主因已簽算力合約付款大增（見驅動因子章）。")
    return (f'<section id="pl"><h2>損益（標準損益表）</h2>{basis}{ch1}{tbl}{notes}</section>')


# ── 驅動因子 ────────────────────────────────────────────────────────────
def drivers_v13u(B, M, rev_sec_html: str):
    sp = M.span
    head = "".join(f"<th>{y}</th>" for y in YRS)

    def tr(lab, nm_, f, cls=""):
        pts = _yrs(B, nm_)
        return f'<tr class="{cls}"><td>{lab}</td>' + "".join(f"<td>{sp(r, v, f) if v is not None else '—'}</td>" for r, v in pts) + "</tr>"
    seg = (f'<div class="tw"><table class="drv"><thead><tr><th>分部營收推導</th>{head}</tr></thead><tbody>'
           + tr("Go 用戶（百萬人）", "DEM_Users_Go", "n1") + tr("＝Go 訂閱營收（人數 × 實收 ARPU × 12）", "REV_Sub_Go", "n2", "eq")
           + tr("Plus 用戶（百萬人）", "DEM_Users_Plus", "n1") + tr("＝Plus 訂閱營收", "REV_Sub_Plus", "n2", "eq")
           + tr("Pro 用戶（百萬人）", "DEM_Users_Pro", "n2") + tr("＝Pro 訂閱營收", "REV_Sub_Pro", "n2", "eq")
           + tr("企業付費席位（百萬席）", "DEM_Users_Seats", "n0") + tr("＝企業席位營收", "REV_Sub_Seats", "n2", "eq")
           + tr("訂閱合計", "REV_Sub", "n2", "eq")
           + tr("API 計費 token（兆）", "DEM_Tok_API", "n0") + tr("× API 組合有效單價（美元／百萬 token）", "REV_ApiPrice", "n2")
           + tr("＝API 營收", "REV_API", "n2", "eq")
           + tr("＋廣告營收", "REV_Ads", "n2") + tr("＋其他營收", "REV_Other", "n2")
           + tr("＝營收總額（截頂後；即損益表營收）", "REV_GrossCapped", "n2", "eq")
           + tr("− Microsoft 分成", "REV_MSShareCapped", "n2") + tr("＝營收淨額（命題與現金流用）", "REV_NetCapped", "n2", "eq")
           + "</tbody></table></div>")
    ch1 = bar_chart(sp, YRS, [("API", C_TEAL, _yrs(B, "REV_API")), ("訂閱", C_GOLD, _yrs(B, "REV_Sub")), ("廣告", C_PUR, _yrs(B, "REV_Ads"))], "十億美元",
                    "分部營收（堆疊；頂端數字＝營收總額）", stacked=True, totals=_yrs(B, "REV_GrossCapped"), fmt="n0",
                    note="其他營收各年為 0，不畫；頂端數字取自模型營收總額。")
    ch2 = bar_chart(sp, YRS, [("API 計費 token（兆）", C_TEAL, _yrs(B, "DEM_Tok_API"))], "兆 token", "量：API 計費 token", fmt="n0")
    ch3 = bar_chart(sp, YRS, [("API 組合有效單價", C_ACC, _yrs(B, "REV_ApiPrice"))], "美元／百萬 token", "價：API 組合有效單價", fmt="n2")
    cost = (f'<div class="tw"><table class="drv"><thead><tr><th>成本與現金面（基準；十億美元）</th>{head}</tr></thead><tbody>'
            + tr("算力成本（現金口徑）", "COST_Compute", "n1", "eq") + tr("　其中：合約實付", "COST_ContractPay", "n1") + tr("　其中：自有資本支出", "COST_OwnedCapex", "n1")
            + tr("非算力成本（不含股權報酬）", "COST_NonCompExSBC", "n1") + tr("股權報酬（非現金）", "COST_SBC", "n1")
            + tr("營運自由現金流", "FND_FCF", "sn1", "eq") + tr("當年外部資金需求", "FND_ExtNeed", "n1") + "</tbody></table></div>")
    ch4 = bar_chart(sp, YRS, [("合約實付", C_TEAL, _yrs(B, "COST_ContractPay")), ("自有資本支出", C_PUR, _yrs(B, "COST_OwnedCapex"))], "十億美元",
                    "算力成本：合約實付與自有資本支出（堆疊；頂端＝算力成本合計）", stacked=True, totals=_yrs(B, "COST_Compute"), fmt="n0",
                    note="2027 年起合約付款大增，是自由現金流轉為大額負值、外部資金需求出現的主因。")
    intro = ('<p class="note">本章依命題挑出最重要的驅動因子：營收面看「量」（API 計費 token、付費用戶與席位）與「價」（API 有效單價、實收 ARPU），'
             '成本面看已簽算力合約的付款與自有資本支出。2026 年起 API 約占營收一半以上，因此 API token 量與單價是命題最敏感的兩個變數。</p>')
    return (f'<section id="drivers"><h2>驅動因子 → 推導量</h2>{intro}<h3>分部營收怎麼算出來</h3>{ch1}{seg}'
            f'<div class="charts2">{ch2}{ch3}</div><p class="note">量增價跌：API 組合有效單價 2026 年後逐年下降（牌價年降假設），營收成長來自 token 量。</p>'
            f'<h3>成本與資金面</h3>{ch4}{cost}'
            f'<details class="more"><summary>展開：原收支頁「營收結構與管理層目標」</summary>{rev_sec_html}</details></section>')


def payback_chart(B, M):
    """回本比值（實際 m×L ÷ 所需 m×L）三種口徑 × 低／基準／高（讀 PxQ 頁具名範圍 PXQ_Out_Ratio；比值＞1 表示實收足以覆蓋機隊持有成本）。"""
    g = B.grid("PXQ_Out_Ratio")
    cats = ["口徑一：Q2 年化", "口徑二：FY2026", "口徑三：9 月底"]
    ser = [(nm_, col, [row[i] for row in g]) for i, (nm_, col) in enumerate([("低案", C_GOLD), ("基準", C_ACC), ("高案", C_TEAL)])]
    return bar_chart(M.span, cats, ser, "倍", "回本比值（實際 m×L ÷ 所需 m×L；＞1＝覆蓋持有成本）", fmt="n2", h=260,
                     note="基準案三種口徑的比值都略高於 1 倍；重點是相對 2025 的趨勢（見下方明細表）。")


# ── 組裝 ────────────────────────────────────────────────────────────────
def build_v13u(html_in: Path, xlsx: Path, out: Path, P: dict, M):
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
    for t in soup.select("main svg title"):
        parts_ = [x for x in t.get_text().split("｜") if not PARAM_RE.fullmatch(x.strip())]
        t.string = "｜".join(parts_)
    css = M.strip_dark(soup.style.string) + M.CSS_EXTRA + CSS_V13

    parts = [thesis_v13u(secs["s1"], B, M, P), valuation_v13u(B, M)]
    s6 = M.regroup(secs["s6"], soup, "x", "收支命題的敏感度", summary="展開：收支命題的關鍵驅動表與全部案例")
    s6_inner = "".join(str(c) for c in [c for c in s6.children if getattr(c, "name", None)][1:])
    parts.append(sens_v13u(B, M, s6_inner))
    parts.append(pl_v13u(B, M))
    rev_old = M.regroup(secs["s3"], soup, "x", "營收結構")
    rev_inner = "".join(str(c) for c in [c for c in rev_old.children if getattr(c, "name", None)][1:])
    parts.append(drivers_v13u(B, M, rev_inner))
    mods = [("s5", "funding", "資金：現金與外部資金需求", "資金"), ("s2", "unit", "每 VR 等值 GW：營收 vs 全成本", "每 GW 收支"),
            ("s4", "compute", "算力：總需求 vs 合約＋自建", "算力"), ("s6b", "payback", "回本比值（實際 m×L ÷ 門檻）", "回本比值")]
    for sid, nid, title, _nav in mods:
        sec_html = str(M.regroup(secs[sid], soup, nid, title))
        if nid == "payback":
            sec_html = sec_html.replace("</h2>", "</h2>" + payback_chart(B, M), 1)
        parts.append(sec_html)
    recon = soup.new_tag("section", id="recon")
    h = soup.new_tag("h2")
    h.string = P["recon_title"]
    recon.append(h)
    _c = M.coord_c0_generic(B)
    if _c:
        for el in list(BeautifulSoup(_c, "html.parser").children):
            recon.append(el)
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

    navs = [("thesis", "命題與核心問題"), ("valuation", "估值（不適用）"), ("sens", "敏感度"), ("pl", "損益"), ("drivers", "驅動因子")] + \
        [(nid, nv) for _, nid, _t, nv in mods] + [("recon", "對帳"), ("status", "資料狀態")]
    nav = "".join(f'<a href="#{a}">{t}</a>' for a, t in navs)
    doc = (f'<!DOCTYPE html><html lang="zh-Hant"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">'
           f'<title>{P["name"]} 個股模型 v{ver}</title><style>{css}</style></head><body><div class="wrap">'
           f'<header class="top"><h1>{P["name"]} 個股模型</h1><div class="meta">{P["meta"]}｜版本 v{ver}｜{meta_txt}</div>'
           f'<nav class="toc">{nav}</nav></header>{hero_v13u(B, M)}<main class="v2">{"".join(parts)}</main>'
           f'<footer class="v2">{P["name"]} 個股模型 v{ver}｜版面 v1.3（未上市、無市場價格；tools/company_html/fundmodel_v13u.py；個股 HTML 版面規格 J 節）｜來源 HTML <code>{html.escape(html_in.name)}</code>｜'
           f'xlsx <code>{html.escape(xlsx.name)}</code></footer></div></body></html>')
    out.write_text(doc, encoding="utf-8")
    return out
