"""Anthropic 個股模型 HTML 新版面（v2，SpaceX 式；chat 端 2026-11-11 草稿）。

用途：把現行產生器（tools/build_html.py）輸出的 HTML 重排成「個股 HTML 版面規格 v1.0」的樣子：
  頁首＋目錄 → 投資建議橫幅（加權目標價、市場主流 Comparable 目標價、現價）→ 1. 命題與核心問題（分點、
  命題回覆摺疊、重點觀察）→ 2–5. 收支模型原有分區（圖表直接顯示、明細表與說明摺疊）→ 6. 估值（Val 頁）
  → 7. 敏感度 → 8. 對帳 → 9. 資料狀態（附錄：版本要點與各版更新紀錄，摺疊）。

原則（與 tools/company_html 相同）：
- 數字一律取自 xlsx（重用頁的數字沿用原 HTML 的 data-ref／data-v；新節由本檔讀 dist xlsx 的具名範圍），本檔不寫數字。
- 自含：內嵌 CSS 與 inline SVG，無外部網址、無 JavaScript；白底、青綠主色、不做深色模式（Andy 2026-10-23）。
- 不改原產生器；原 HTML 是本檔的輸入。

用法：python3 tools/build_html_v2.py --html dist/<v1>.html --xlsx dist/<同版>.xlsx --out <輸出>.html
"""
from __future__ import annotations

import argparse
import copy
import html
import re
from pathlib import Path

import openpyxl
from bs4 import BeautifulSoup, NavigableString

MINUS = "−"


# ── xlsx 具名範圍讀值 ────────────────────────────────────────────────────
class Book:
    def __init__(self, path: Path):
        self.path = path
        self.wb = openpyxl.load_workbook(path, data_only=True)
        self.names = {}
        for nm, dn in self.wb.defined_names.items():
            self.names[nm] = dn.attr_text

    def has(self, name: str) -> bool:
        return name in self.names

    def _cells(self, name: str):
        ref = self.names[name]
        sheet, rng = ref.split("!")
        sheet = sheet.strip("'")
        rng = rng.replace("$", "")
        ws = self.wb[sheet]
        if ":" in rng:
            return sheet, [[c for c in row] for row in ws[rng]]
        return sheet, [[ws[rng]]]

    def cell(self, name: str):
        sheet, rows = self._cells(name)
        c = rows[0][0]
        return f"{sheet}!{c.coordinate}", c.value

    def yr(self, name: str, year: int):
        """年度列（D–I＝2025–2030）取某年。"""
        row = self.grid(name)[0]
        return row[year - 2025]

    def grid(self, name: str):
        sheet, rows = self._cells(name)
        return [[(f"{sheet}!{c.coordinate}", c.value) for c in row] for row in rows]


def fmt(v, f: str) -> str:
    if v is None:
        return "—"
    if isinstance(v, str):
        return v
    if f == "t":
        return str(v)
    m = re.fullmatch(r"(s?)(n|p)(\d)", f)
    if not m:
        raise ValueError(f)
    sgn, kind, d = m.group(1), m.group(2), int(m.group(3))
    x = v * 100 if kind == "p" else v
    x = round(x, d)
    if x == 0:
        x = 0.0
    s = MINUS if x < 0 else ("＋" if (sgn and x > 0) else "")
    return f"{s}{abs(x):,.{d}f}{'%' if kind == 'p' else ''}"


def span(ref: str, v, f: str = "n2", cls: str = "n") -> str:
    """可追溯數字（與原 HTML 同一格式：data-ref、data-v、title）。"""
    if isinstance(v, str) or v is None:
        return f'<span class="lbl" data-ref="{html.escape(ref)}" title="{html.escape(ref)}">{html.escape(fmt(v, "t"))}</span>'
    return (f'<span class="{cls}" data-ref="{html.escape(ref)}" data-fmt="{f}" data-v="{html.escape(repr(v))}" '
            f'title="{html.escape(ref)}">{html.escape(fmt(v, f))}</span>')


# ── 版面 CSS（SpaceX v0.2.7 版面＋原收支頁圖表類別） ───────────────────────
CSS_EXTRA = """
:root{--bg:#fff;--fg:#3b3838;--muted:#5f605f;--line:#dce3e3;--card:#fff;--accent:#0b7a80;--teal:#10a0a7;--soft:#f7fbfb;--grid2:#ecefef;--negc:#b4472f}
body{background:var(--bg);color:var(--fg);font:15px/1.65 'Noto Sans TC','Microsoft JhengHei','微軟正黑體','PingFang TC',Arial,sans-serif;font-variant-numeric:tabular-nums}
.wrap{max-width:78rem;margin:0 auto;padding:32px 24px 40px}
header.top{display:grid;gap:8px}
header.top h1{font-size:28px;line-height:1.3;font-weight:400;margin:0}
nav.toc{display:flex;flex-wrap:wrap;gap:2px;border-bottom:1px solid var(--line);margin-top:10px;font-size:14px}
nav.toc a{color:var(--muted);text-decoration:none;padding:6px 12px;border-bottom:2px solid transparent;margin-bottom:-1px}
nav.toc a:hover{color:var(--fg);border-bottom-color:var(--teal)}
main.v2{counter-reset:sec;display:grid;gap:24px;margin-top:20px;max-width:none;padding:0}
main.v2>section{border-radius:6px;padding:20px;margin:0;min-width:0}
main.v2>section>h2{font-size:20px;font-weight:600;margin:0 0 8px;counter-increment:sec;display:block}
main.v2>section>h2::before{content:counter(sec) ". ";color:var(--teal)}
.thesis{font-size:18px;line-height:1.6;margin:4px 0 8px}
.q{border-left:3px solid var(--teal);padding:4px 12px;margin:10px 0;background:var(--soft)}
ol.items{margin:6px 0;padding-left:24px}
details.more{margin:10px 0}details.more>summary{cursor:pointer;color:var(--accent)}
details.more[open]>summary{margin-bottom:6px}
details.hist{border-top:1px solid var(--line);padding:6px 0}details.hist>summary{cursor:pointer;color:var(--accent);font-weight:600}
.note{font-size:13.5px}
main.v2 .cw svg{max-width:900px;margin:0 auto}
.lead{border-left-color:var(--teal)}
.hero{margin-top:18px;border-radius:8px;padding:18px 22px;border:2px solid var(--line);display:grid;gap:12px}
.hero-neg{background:#fbece8;border-color:#b4472f}.hero-pos{background:#e6f4f3;border-color:#0b7a80}.hero-neu{background:#fbf6e8;border-color:#c29a2b}.hero-na{background:#f4f5f5}
.hero-rec{display:flex;align-items:baseline;gap:14px;flex-wrap:wrap}
.hero-k{font-size:16px;color:var(--muted);font-weight:600}
.hero-v{font-size:34px;font-weight:700;line-height:1.2}
.hero-neg .hero-v,.hero-neg .hero-s{color:#b4472f}.hero-pos .hero-v,.hero-pos .hero-s{color:#0b7a80}.hero-neu .hero-v,.hero-neu .hero-s{color:#8a6d16}
.hero-tiles{display:grid;grid-template-columns:repeat(auto-fit,minmax(min(100%,15rem),1fr));gap:12px}
.hero-t{background:#fff;border:1px solid var(--line);border-radius:6px;padding:10px 14px}
.hero-l{font-size:14px;color:var(--muted)}
.hero-n{font-size:28px;font-weight:700;line-height:1.3}
.hero-s{font-size:17px;font-weight:600}
.hero-note{margin:0;font-size:13px;color:var(--muted)}
.br{display:grid;grid-template-columns:minmax(90px,220px) 1fr 72px;gap:8px;align-items:center;margin:6px 0}
.bt{position:relative;height:18px;background:var(--grid2);border-radius:3px}
.bt i{position:absolute;left:0;top:0;bottom:0;border-radius:3px;background:var(--teal)}
.bt i.w{background:var(--accent)}.bt i.c{background:#c29a2b}.bt i.d{background:#6f6bae}
.bt b{position:absolute;top:-4px;bottom:-4px;width:2px;background:var(--fg)}
.bt u{position:absolute;top:0;bottom:0;background:rgba(180,71,47,.15);text-decoration:none}
.bv{text-align:right}
.refline{font-size:13px;color:var(--muted);margin:4px 0 0}
.trw{display:grid;grid-template-columns:minmax(110px,260px) 1fr 70px;gap:8px;align-items:center;margin:5px 0;font-size:13px}
.ttr{position:relative;height:14px;background:var(--grid2);border-radius:3px}
.ttr::before{content:"";position:absolute;left:50%;top:0;bottom:0;width:1px;background:#8a8c8c}
.ttr i{position:absolute;top:2px;bottom:2px;border-radius:2px}.ttr i.p{background:var(--teal)}.ttr i.n{background:var(--negc)}
dl.kv{display:block;margin:8px 0;font-size:14px}dl.kv div{display:flex;gap:10px;border-bottom:1px solid var(--line);padding:4px 0;flex-wrap:wrap}
dl.kv dt{color:var(--muted);flex:1 1 220px}dl.kv dd{margin:0;flex:2 1 220px;word-break:normal}
footer.v2{margin:24px 0 0;padding-top:10px;border-top:1px solid var(--line);font-size:12px;color:var(--muted);text-align:left}
@media (max-width:640px){.wrap{padding:16px}header.top h1{font-size:23px}main.v2>section{padding:14px}.hero-v{font-size:28px}.hero-n{font-size:24px}
 .br{grid-template-columns:80px 1fr 58px}.trw{grid-template-columns:96px 1fr 58px;font-size:12px}}
"""


def strip_dark(css: str) -> str:
    """原 CSS 的深色模式區塊移除（白底版面，Andy 2026-10-23）。"""
    i = css.find("@media (prefers-color-scheme: dark)")
    if i < 0:
        return css
    depth, j = 0, css.find("{", i)
    while j < len(css):
        if css[j] == "{":
            depth += 1
        elif css[j] == "}":
            depth -= 1
            if depth == 0:
                return css[:i] + css[j + 1:]
        j += 1
    return css


# ── 原分區重排：圖表、主結論顯示；表格與說明摺疊 ────────────────────────
VISIBLE = {("div", "legend"), ("div", "cw"), ("p", "lead"), ("div", "kpis"), ("div", "must")}


def _key(tag):
    cls = tag.get("class") or [""]
    return (tag.name, cls[0])


def regroup(sec, soup, new_id: str, new_title: str, keep_first_table: bool = False, summary: str | None = None):
    kids = [c for c in sec.children if getattr(c, "name", None)]
    h2 = kids[0]
    body = kids[1:]
    vis, hid = [], []
    has_chart = any(_key(c) == ("div", "cw") for c in body)
    first_tbl_done = False
    for c in body:
        k = _key(c)
        if k in VISIBLE:
            vis.append(c)
        elif keep_first_table and not has_chart and not first_tbl_done and k in {("div", "tw"), ("ol", "")}:
            vis.append(c)
            first_tbl_done = True
        else:
            hid.append(c)
    out = soup.new_tag("section", id=new_id)
    h = soup.new_tag("h2")
    h.string = new_title
    out.append(h)
    # 原標題（含舊編號）保留為小字，方便對照舊版
    old = soup.new_tag("p", attrs={"class": "note"})
    old.string = "原收支頁：" + re.sub(r"\s+", " ", h2.get_text(" ", strip=True))
    for c in vis:
        out.append(copy.copy(c))
    if hid:
        d = soup.new_tag("details", attrs={"class": "more"})
        s = soup.new_tag("summary")
        ntab = sum(1 for c in hid if _key(c) == ("div", "tw") or c.find("table"))
        s.string = summary or f"展開：明細表與說明（{ntab} 張表）"
        d.append(s)
        for c in hid:
            d.append(copy.copy(c))
        d.append(old)
        out.append(d)
    else:
        out.append(old)
    return out


# ── 新節：橫幅、命題、估值、敏感度龍捲風 ──────────────────────────────────
def hero_html(B: Book) -> str:
    if not B.has("VAL_TP_W"):
        return ('<div class="hero hero-na"><div class="hero-rec"><span class="hero-k">投資建議</span>'
                '<span class="hero-v">估值頁建置中</span></div><p class="hero-note">本版 xlsx 尚無 Val 頁；'
                'v0.2.0 起由 Val 頁具名範圍帶入加權目標價、Comparable 目標價與現價。</p></div>')
    rref, rating = B.cell("VAL_Rating")
    cls = {"強力放空": "hero-neg", "放空": "hero-neg", "中立": "hero-neu", "買進": "hero-pos", "強力買進": "hero-pos"}.get(str(rating), "hero-neu")
    w = B.cell("VAL_TP_W"); uw = B.cell("VAL_Up_W")
    c = B.cell("VAL_TP_CompMain"); uc = B.cell("VAL_Up_CompMain")
    px = B.cell("VAL_Px"); lo = B.cell("VAL_PxLo"); hi = B.cell("VAL_PxHi")
    note = B.cell("VAL_RatingNote") if B.has("VAL_RatingNote") else None
    tiles = (
        f'<div class="hero-t"><div class="hero-l">目標價（DCF 與 Comparable 加權）</div><div class="hero-n">{span(*w, "n2")}</div>'
        f'<div class="hero-s">上檔空間 {span(*uw, "sp1")}</div></div>'
        f'<div class="hero-t"><div class="hero-l">Comparable 目標價（市場主流同業法，未調整）</div><div class="hero-n">{span(*c, "n2")}</div>'
        f'<div class="hero-s">上檔空間 {span(*uc, "sp1")}</div></div>'
        f'<div class="hero-t"><div class="hero-l">現價（IPO 價格區間中點；尚未掛牌）</div><div class="hero-n">{span(*px, "n2")}</div>'
        f'<div class="hero-s" style="color:var(--muted)">區間 {span(*lo, "n2")}–{span(*hi, "n2")}</div></div>')
    nt = f'<p class="hero-note">{span(*note)}</p>' if note else ""
    return (f'<div class="hero {cls}"><div class="hero-rec"><span class="hero-k">投資建議</span><span class="hero-v">{span(rref, rating)}</span></div>'
            f'<div class="hero-tiles">{tiles}</div>{nt}'
            '<p class="hero-note">投資建議依加權目標價判定，五級固定規則：上檔空間 &gt;＋50% 強力買進、＋25%～＋50% 買進、−25%～＋25% 中立、−50%～−25% 放空、低於 −50% 強力放空。</p></div>')


def bars_html(B: Book) -> str:
    rows = [("加權目標價", "VAL_TP_W", "w"), ("DCF（三情境加權）", "VAL_TP_DCF", "d"), ("　DCF 情境 A（樂觀）", "VAL_TP_DCF_A", "d"),
            ("　DCF 情境 B（基準）", "VAL_TP_DCF_B", "d"), ("　DCF 情境 C（保守）", "VAL_TP_DCF_C", "d"),
            ("Comparable（調整後，進加權）", "VAL_TP_CompAdj", "c"), ("Comparable（市場主流，未調整）", "VAL_TP_CompMain", "c")]
    rows = [r for r in rows if B.has(r[1])]
    vals = [B.cell(r[1]) for r in rows]
    px = B.cell("VAL_Px"); lo = B.cell("VAL_PxLo")[1]; hi = B.cell("VAL_PxHi")[1]
    mx = max([v for _, v in vals if isinstance(v, (int, float))] + [hi]) * 1.08
    out = []
    for (lab, _nm, c), (ref, v) in zip(rows, vals):
        w = max(0.0, (v or 0) / mx * 100)
        out.append(f'<div class="br"><span>{lab}</span><div class="bt"><u style="left:{lo / mx * 100:.2f}%;width:{(hi - lo) / mx * 100:.2f}%"></u>'
                   f'<i class="{c}" style="width:{w:.2f}%"></i><b style="left:{px[1] / mx * 100:.2f}%"></b></div><span class="bv">{span(ref, v, "n2")}</span></div>')
    return ("".join(out) + f'<p class="refline">黑線＝現價 {span(*px, "n2")}（IPO 價格區間中點）；淡紅底＝價格區間。單位：美元／股。</p>')


def kv(B: Book, items) -> str:
    out = []
    for lab, nm, f in items:
        if B.has(nm):
            out.append(f"<div><dt>{lab}</dt><dd>{span(*B.cell(nm), f)}</dd></div>")
    return f'<dl class="kv">{"".join(out)}</dl>'


def table_from_grid(B: Book, nm: str, head: list[str], fmts: list[str]) -> str:
    g = B.grid(nm)
    trs = []
    for row in g:
        if all(v is None for _, v in row):
            continue
        tds = "".join(f"<td>{span(r, v, f) if v is not None else ''}</td>" for (r, v), f in zip(row, fmts))
        trs.append(f"<tr>{tds}</tr>")
    th = "".join(f"<th>{h}</th>" for h in head)
    return f'<div class="tw"><table><thead><tr>{th}</tr></thead><tbody>{"".join(trs)}</tbody></table></div>'


def valuation_section(B: Book) -> str:
    if not B.has("VAL_TP_W"):
        return ('<section id="valuation"><h2>估值</h2><p class="note">本版 xlsx 尚無 Val 頁（v0.2.0 新增：DCF〔無槓桿自由現金流〕與 Comparable 各半、'
                '情境 A／B／C 25／50／25、評等依 IPO 價格）。</p></section>')
    meth = B.cell("VAL_Method") if B.has("VAL_Method") else None
    intro = f'<p class="note">{span(*meth)}</p>' if meth else ""
    main_kv = kv(B, [("加權目標價（DCF 50%＋調整後 Comparable 50%）", "VAL_TP_W", "n2"), ("DCF 目標價（三情境加權）", "VAL_TP_DCF", "n2"),
                     ("Comparable 目標價（調整後）", "VAL_TP_CompAdj", "n2"), ("Comparable 目標價（市場主流，未調整）", "VAL_TP_CompMain", "n2"),
                     ("現價（IPO 區間中點）", "VAL_Px", "n2"), ("上檔空間（加權）", "VAL_Up_W", "sp1"), ("上檔空間（市場主流 Comparable）", "VAL_Up_CompMain", "sp1"),
                     ("投資建議", "VAL_Rating", "t"), ("情境目標價 A／B／C（DCF）", "VAL_TP_DCF_A", "n2")])
    rates = kv(B, [("無風險利率 rf", "VAL_Rf", "p2"), ("β（同業中位數，Blume 調整）", "VAL_Beta", "n2"), ("股權風險溢酬 ERP", "VAL_ERP", "p1"),
                   ("股東權益成本 ke", "VAL_Ke", "p2"), ("稅後債務成本", "VAL_KdAT", "p2"), ("WACC", "VAL_WACC", "p2"),
                   ("永續期新投資報酬 RONIC", "VAL_RONIC", "p2"), ("永續成長率 g", "VAL_g", "p1")])
    dcf = kv(B, [("企業價值 EV（情境 B）", "VAL_EV_B", "n1"), ("　明確期現值", "VAL_PVExp_B", "n1"), ("　終值現值", "VAL_PVTV_B", "n1"),
                 ("終值占 EV", "VAL_TVShare_B", "p0"), ("＋淨現金（pro forma，含 IPO 與私募）", "VAL_NetCash", "n1"),
                 ("＝股權價值（情境 B）", "VAL_EqV_B", "n1"), ("評價股數（十億股，含 RSU、選擇權庫藏股法）", "VAL_Shares", "n3")])
    fc = ""
    if B.has("VAL_Yr"):
        yrs = [v for _, v in B.grid("VAL_Yr")[0]]
        rows = []
        for lab, nm, f in (("營收淨額", "VAL_Rev_B", "n1"), ("EBIT（現金口徑，含股權報酬）", "VAL_EBIT_B", "n1"), ("稅", "VAL_Tax_B", "n1"),
                           ("FCFF", "VAL_FCFF_B", "n1"), ("折現因子", "VAL_DF_B", "n3")):
            if B.has(nm):
                rows.append(f"<tr><td>{lab}</td>" + "".join(f"<td>{span(r, v, f)}</td>" for r, v in B.grid(nm)[0]) + "</tr>")
        fc = (f'<div class="tw"><table><thead><tr><th>$B（情境 B）</th>{"".join(f"<th>{html.escape(str(y))}</th>" for y in yrs)}</tr></thead>'
              f'<tbody>{"".join(rows)}</tbody></table></div>')
    peers = ""
    if B.has("VAL_PeerTbl"):
        peers = table_from_grid(B, "VAL_PeerTbl", ["代號", "名稱", "EV／遠期營收", "TTM FCF 利益率", "5 年 β", "資料日"], ["t", "t", "n2", "p1", "n2", "t"])
    comp = kv(B, [("同業 EV／NTM 營收中位數", "VAL_PeerMult", "n2"), ("同業 TTM FCF 利益率中位數", "VAL_PeerFCFm", "p1"),
                  ("調整係數（Anthropic FY2030 自由現金流率 ÷ 同業）", "VAL_AdjFactor", "n2"), ("Anthropic NTM 營收總額（情境 B）", "VAL_NTMRev", "n1")])
    return (f'<section id="valuation"><h2>估值</h2>{intro}<h3>目標價與現價</h3>{bars_html(B)}<h3>結論主要項目</h3>{main_kv}'
            f'<h3>DCF 折現率</h3>{rates}<h3>DCF 彙總（情境 B）</h3>{dcf}'
            f'<details class="more"><summary>展開：逐年自由現金流（FCFF）</summary>{fc}</details>'
            f'<details class="more"><summary>展開：同業倍數與 Comparable</summary>{comp}{peers}</details></section>')


def tornado_html(B: Book) -> tuple[str, str]:
    if not (B.has("VAL_SensTbl") and B.has("VAL_SensBase")):
        return "", ""
    bref, base = B.cell("VAL_SensBase")
    rows = []
    for row in B.grid("VAL_SensTbl"):
        (lr, lab), (r1, v1), (r2, v2) = row[:3]
        if lab is None or not isinstance(v1, (int, float)) or not isinstance(v2, (int, float)):
            continue
        d = max(abs(v1 / base - 1), abs(v2 / base - 1))
        rows.append((d, lab, lr, (r1, v1), (r2, v2)))
    rows.sort(key=lambda r: -r[0])
    mx = max([r[0] for r in rows] + [0.1])

    def line(r):
        d, lab, lr, (r1, v1), (r2, v2) = r
        bars = ""
        for v in (v1, v2):
            x = (v / base - 1) / mx * 50
            if x >= 0:
                bars += f'<i class="p" style="left:50%;width:{x:.2f}%"></i>'
            else:
                bars += f'<i class="n" style="left:{50 + x:.2f}%;width:{-x:.2f}%"></i>'
        return (f'<div class="trw"><span>{span(lr, lab)}</span><div class="ttr">{bars}</div>'
                f'<span class="tvv">{span(r1, v1, "n1")}／{span(r2, v2, "n1")}</span></div>')
    big = [line(r) for r in rows if r[0] >= 0.10]
    small = [line(r) for r in rows if r[0] < 0.10]
    head = f'<p class="note">加權目標價基準 {span(bref, base, "n2")}；長條＝低值／高值情況下的變動（中線＝基準）；只顯示變動達 10% 以上的案例。</p>'
    vis = head + "".join(big)
    hid = (f'<details class="more"><summary>展開：變動未達 10% 的案例（{len(small)} 項）</summary>{"".join(small)}</details>' if small else "")
    return vis, hid


def thesis_section(soup_s1, B: Book, ver: str) -> str:
    lead = soup_s1.find("p", class_="lead")
    kpis = soup_s1.find("div", class_="kpis")
    notes = soup_s1.find_all("p", class_="note")
    reply = "".join(str(x) for x in ([lead, kpis] + notes[2:]) if x is not None)
    hist_notes = "".join(str(x) for x in notes[:2])
    obs = []
    if B.has("VAL_TP_W"):
        obs.append(f'<li><b>IPO 價格偏貴或合理，看你相信哪一種估值：</b>加權目標價 {span(*B.cell("VAL_TP_W"), "n2")} 美元，對 IPO 區間中點 '
                   f'{span(*B.cell("VAL_Px"), "n2")} 的上檔空間 {span(*B.cell("VAL_Up_W"), "sp1")}；其中 DCF 只有 {span(*B.cell("VAL_TP_DCF"), "n2")}、'
                   f'市場主流的同業倍數法則是 {span(*B.cell("VAL_TP_CompMain"), "n2")}——兩者差距本身就是這檔股票的核心爭議。</li>')
    obs.append(f'<li><b>算力承諾要到 {span(*B.cell("COST_GapPosSustainYear"), "t")} 年才被營收穩定覆蓋：</b>在那之前每 VR 等值 GW 的營收多數年度低於全成本，'
               f'該年才持續轉正（差額 {span(*B.yr("COST_PropGap_VR", 2030), "sn1")} 十億美元／GW／年，'
               f'覆蓋率 {span(*B.yr("COST_Coverage", 2030), "p0")}；2027、2028 年的差額分別為 {span(*B.yr("COST_PropGap_VR", 2027), "sn1")}、{span(*B.yr("COST_PropGap_VR", 2028), "sn1")}，是承諾付款跳升、缺口最深的兩年）。</li>')
    obs.append(f'<li><b>錢目前夠用：</b>營運自由現金流在承諾付款高峰轉負（2027 年 {span(*B.yr("FND_FCF", 2027), "sn1")}、2028 年 {span(*B.yr("FND_FCF", 2028), "sn1")} 十億美元），'
               f'但只靠已到位融資，現金谷底仍有 {span(*B.cell("FND_CashNoExtTrough"), "n1")} 十億美元，2026–2030 累計外部資金需求 {span(*B.yr("FND_ExtNeedCum", 2030), "n1")}。'
               f'風險在承諾付款時點與 SpaceXAI 合約路徑，而不是眼前的現金。</li>')
    return (f'<section id="thesis"><h2>命題與核心問題</h2>'
            f'<p class="thesis">命題：Anthropic 的營收能否追上已簽下的算力承諾，以及 IPO 價格是否已反映這件事；三個問題如下。</p>'
            f'<p class="q"><strong>核心問題：</strong>每 VR 等值 GW 的年營收，能否覆蓋每 GW 年全成本？若不能，缺口要多少外部資金、由誰補？以 IPO 價格買進，每股值多少？</p>'
            f'<h3>命題</h3><ol class="items">'
            f'<li>算力承諾能否被營收覆蓋：已簽與招股書揭露的算力承諾，換成每 VR 等值 GW 的全成本，哪一年起被營收覆蓋。</li>'
            f'<li>錢夠不夠：營運自由現金流加上已到位融資與 IPO 募資，能否撐過承諾付款的高峰，需不需要再找外部資金。</li>'
            f'<li>每股值多少：以 DCF（無槓桿自由現金流）與同業倍數估算每股價值，對照 IPO 價格區間給出投資建議。</li></ol>'
            f'<details class="more"><summary>展開：命題回覆（依本模型計算的答案；想先看結論可展開，或讀完全文再看）</summary>{reply}</details>'
            f'<h3>重點觀察</h3><ol class="items">{"".join(obs)}</ol>'
            f'<details class="more"><summary>展開：本版與前版摘要</summary>{hist_notes}</details></section>')


def build(html_in: Path, xlsx: Path, out: Path):
    src = html_in.read_text(encoding="utf-8")
    soup = BeautifulSoup(src, "lxml")
    B = Book(xlsx)
    m = re.search(r"v(\d+\.\d+\.\d+(?:\.\d+)?)", soup.header.h1.get_text())
    ver = m.group(1) if m else "?"
    meta_div = soup.header.find("div", class_="meta")
    meta_b = meta_div.find("b")
    ver_notes = str(meta_b) if meta_b else ""
    if meta_b:
        meta_b.decompose()
    meta_txt = meta_div.decode_contents().rstrip("｜")
    secs = {s["id"]: s for s in soup.select("main > section")}
    css = strip_dark(soup.style.string) + CSS_EXTRA

    parts = []
    parts.append(thesis_section(secs["s1"], B, ver))
    plan = [("s5", "funding", "資金：現金路徑與融資"), ("s2", "unit", "每 VR 等值 GW：營收 vs 全成本"),
            ("s3", "revenue", "營收結構與年化營收"), ("s4", "compute", "算力：需求 vs 已簽合約與自建供給")]
    for sid, nid, title in plan:
        parts.append(str(regroup(secs[sid], soup, nid, title)))
    parts.append(valuation_section(B))
    tv, th = tornado_html(B)
    sens = regroup(secs["s6"], soup, "sens", "敏感度與「要成立必須…」", keep_first_table=False,
                   summary="展開：收支命題的關鍵驅動表與全部案例")
    sens_html = str(sens)
    if tv:
        sens_html = sens_html.replace("</h2>", f"</h2><h3>目標價敏感度（龍捲風圖）</h3>{tv}{th}<h3>收支命題的敏感度</h3>", 1)
    parts.append(sens_html)
    recon = soup.new_tag("section", id="recon")
    h = soup.new_tag("h2"); h.string = "對帳：OpenAI 並排與最該審的預設"; recon.append(h)
    for sid, lab in (("s7", "與 OpenAI 並排"), ("s8", "最該審的預設")):
        r = regroup(secs[sid], soup, "x", lab, keep_first_table=True)
        h3 = soup.new_tag("h3"); h3.string = lab; recon.append(h3)
        for c in [c for c in r.children if getattr(c, "name", None)][1:]:
            recon.append(c)
    parts.append(str(recon))
    hist_ids = [k for k in secs if re.fullmatch(r"s1[a-z]", k)]
    hist = "".join(f'<details class="hist"><summary>{html.escape(secs[k].h2.get_text(" ", strip=True))}</summary>{secs[k].decode_contents()}</details>'
                   for k in hist_ids)
    status = regroup(secs["s9"], soup, "status", "資料狀態", keep_first_table=True, summary="展開：資料與版本細節")
    status_html = str(status).replace("</section>", f'<h3>附錄</h3><details class="hist"><summary>版本要點（各版變更摘要）</summary><p class="note">{ver_notes}</p></details>{hist}</section>')
    parts.append(status_html)

    nav = "".join(f'<a href="#{a}">{t}</a>' for a, t in (("thesis", "命題與核心問題"), ("funding", "資金"), ("unit", "每 GW 收支"), ("revenue", "營收"),
                                                          ("compute", "算力"), ("valuation", "估值"), ("sens", "敏感度"), ("recon", "對帳"), ("status", "資料狀態")))
    doc = (f'<!DOCTYPE html><html lang="zh-Hant"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">'
           f'<title>Anthropic 個股模型 v{ver}</title><style>{css}</style></head><body><div class="wrap">'
           f'<header class="top"><h1>Anthropic 個股模型</h1><div class="meta">ANT（Nasdaq，待掛牌）｜版本 v{ver}｜{meta_txt}</div>'
           f'<nav class="toc">{nav}</nav></header>{hero_html(B)}<main class="v2">{"".join(parts)}</main>'
           f'<footer class="v2">Anthropic 個股模型 v{ver}｜版面 v2 草稿（tools/build_html_v2.py，SpaceX 式）｜來源 HTML <code>{html.escape(html_in.name)}</code>｜'
           f'xlsx <code>{html.escape(xlsx.name)}</code></footer></div></body></html>')
    out.write_text(doc, encoding="utf-8")
    return out


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--html", required=True, type=Path)
    ap.add_argument("--xlsx", required=True, type=Path)
    ap.add_argument("--out", required=True, type=Path)
    a = ap.parse_args()
    p = build(a.html, a.xlsx, a.out)
    print(p)


if __name__ == "__main__":
    main()
