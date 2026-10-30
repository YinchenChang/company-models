# SpaceX（SPCX）v0.2 事實增補：424B4 最終招股書（FY2025／FY2024 原文）、2026Q2 10-Q 附註、2026Q2／Q3 法說指引（只存數字與短句）。
# 工作單：spacex/docs/workorders/20261030_SpaceX_v0.2_原文換版與法說指引.md 第 1 項（G1、G2、G3、G5、G6、G7）。
# 選擇：另開本檔（不改 facts_2026q3.py），輸出 spacex/data/spacex_facts_v02.json；build_spacex.py 兩檔合併讀取（同鍵以本檔為準，本檔不重複 v0.1 鍵）。
# 讀法：只讀 mag/extract/docs/2026Q3/text/（pdftotext -layout；\f 分頁＝PDF 頁序）。兩讀：同一數字在兩處（表格與文字句、兩張表、或同表另一欄推回）各讀一次。
# 法說逐字稿不入庫（公開 repo）：數字與 ≤ 10 個英文字的短句寫在本檔，逐字稿在本機時（環境變數 SPCX_TRANSCRIPTS，預設 /home/claude/src）
#   以原文再讀一次核對；不在時（例如 CI）標「逐字稿未入庫，未重讀」，數字沿用本檔，不判失敗。
# 用法：python3 spacex/facts_v02.py → 寫 spacex/data/spacex_facts_v02.json；兩讀不一致、FY2025／FY2024 分部對帳差額非 0 時非零結束。
import json, os, re, sys

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
TX = os.path.join(ROOT, "mag", "extract", "docs", "2026Q3", "text")
TDIR = os.environ.get("SPCX_TRANSCRIPTS", "/home/claude/src")
DOCS = {"B": ("spacex_424B4.txt", "mag/extract/docs/2026Q3/spacex_424B4.pdf", "424B4 最終招股書", "2026-06"),
        "Q2": ("spacex_2026Q2_10-Q.txt", "mag/extract/docs/2026Q3/spacex_2026Q2_10-Q.pdf", "2026Q2 10-Q", "2026-07"),
        "Q3": ("spacex_2026Q3_10-Q.txt", "mag/extract/docs/2026Q3/spacex_2026Q3_10-Q.pdf", "2026Q3 10-Q", "2026-10-22")}
TR = {"T3": ("Transcript_SPCX_3Q26.txt", "2026Q3 法說逐字稿（2026-10-20；Andy 提供，未入庫）"),
      "T2": ("Transcript_SPCX_2Q26.txt", "2026Q2 法說逐字稿（2026-07；Andy 提供，未入庫）")}


def load(path):
    raw = open(path, encoding="utf-8").read()
    lines, pages = [], []
    for p, ptxt in enumerate(raw.split("\f"), 1):
        for l in ptxt.splitlines():
            lines.append(re.sub(r"\.{2,}", " ", l)); pages.append(p)
    return lines, pages


D = {k: load(os.path.join(TX, v[0])) for k, v in DOCS.items()}
for k, (fn, _) in TR.items():
    p = os.path.join(TDIR, fn)
    D[k] = load(p) if os.path.exists(p) else None

NUM = r"\(?\$?\s*-?[\d,]*\d(?:\.\d+)?\)?"


def tonum(s):
    s = s.strip().replace("$", "").replace(",", "").strip()
    neg = s.startswith("(") and s.endswith(")")
    v = float(s.strip("()").strip())
    return -v if neg else v


def trailing(line):
    toks = re.findall(r"\(\s*\$?\s*[\d,]*\d(?:\.\d+)?\s*\)|\$?\s*-?[\d,]*\d(?:\.\d+)?|—", line)
    out = []
    for t in toks:
        out.append(0.0 if t == "—" else tonum(t.replace(" ", "")))
    return out


def blob(d, i, n=4):
    return re.sub(r"\s+", " ", " ".join(x.strip() for x in D[d][0][i:i + n]))


def search(d, pat, start=0, grp=1, n=4, nth=0):
    """文字句：在連續 n 行合併後的文字中找 pat；回傳 (數值, 行號, 原文片段)。"""
    lines = D[d][0]
    hits = 0
    k = start
    while k < len(lines):
        m = re.search(pat, blob(d, k, n))
        if not m:
            k += 1; continue
        # 往後推到命中起點所在行（下一窗仍命中同一段文字時前進）
        while k + 1 < len(lines):
            m2 = re.search(pat, blob(d, k + 1, n))
            if not m2 or m2.group(0) != m.group(0): break
            k += 1; m = m2
        if hits == nth:
            b = blob(d, k, n)
            return tonum(m.group(grp)), k, b[max(0, m.start() - 60):m.end() + 30].strip()
        hits += 1
        k += n
    raise SystemExit(f"[{d}] 找不到：{pat}")


def row(d, label, start, span=60, n=2, nth=0, first=False):
    """表格列：label 開頭（去空白），取最後 n 個數字；數字可在下 1–2 行。"""
    lines = D[d][0]
    hits = 0
    for i in range(start, min(len(lines), start + span)):
        if not lines[i].strip().startswith(label): continue
        if hits < nth:
            hits += 1; continue
        for j in range(i, i + 3):
            v = trailing(lines[j][len(lines[j]) - len(lines[j].lstrip()) + len(label):] if j == i else lines[j])
            if len(v) >= n: return (v[:n] if first else v[-n:]), j
        raise SystemExit(f"[{d}] 列 {label} 數字不足：{lines[i]!r}")
    raise SystemExit(f"[{d}] 找不到列：{label}（第 {start} 行起）")


def find(d, test, start=0):
    for i in range(start, len(D[d][0])):
        if test(D[d][0][i]): return i
    raise SystemExit(f"[{d}] 找不到段落")


def loc(d, i, tbl=""):
    if d in DOCS: return f"{DOCS[d][1]} 第 {D[d][1][i]} 頁" + (f" {tbl}" if tbl else "")
    return f"{TR[d][1]} 第 {D[d][1][i]} 頁" + (f" {tbl}" if tbl else "")


ITEMS, CHK, IDX = [], [], {}


def add(iid, item, value, unit, period, src, quote="", tag="[Verified]", note="", second=None, second_src="", tol=0.5, docDate=""):
    it = {"id": iid, "item": item, "value": value, "unit": unit, "period": period, "source": src, "quote": quote[:200], "docDate": docDate, "tag": tag, "note": note}
    if second is not None:
        ok = abs(second - value) <= tol + 1e-9 * abs(value)
        it["check"] = {"value": second, "source": second_src, "ok": ok}
        CHK.append({"id": iid, "value": value, "second": second, "source": second_src, "ok": ok})
    if iid in IDX: raise SystemExit("重複鍵：" + iid)
    IDX[iid] = it; ITEMS.append(it)
    return it


def chk(cid, a, b, src, tol=0.5):
    ok = abs(a - b) <= tol
    CHK.append({"id": cid, "value": a, "second": b, "source": src, "ok": ok})
    return ok


def V(i): return IDX[i]["value"]


V01 = {it["id"]: it for it in json.load(open(os.path.join(HERE, "data", "spacex_facts_2026q3.json"), encoding="utf-8"))["items"]}
U = "USD M"
SEGN = {"space": "Space", "conn": "Connectivity", "si": "SI（424B4 稱 AI）", "total": "合計"}
B = "B"

# ============ 1. 424B4：xAI 共同控制合併的比較期重編（原句） ============
k = find(B, lambda l: "retrospectively recast for all periods presented to include" in l)
add("b.recast", "424B4：合併財報對所有列示期間追溯重編，納入 xAI（2026-02-02）與 X Holdings（2025-03-28）歷史結果（共同控制）", 1, "是", "all",
    loc(B, k, "Basis of Presentation"), quote=blob(B, k, 3)[:200], note="G5：比較期已重編 → 10-Q 2026Q3 的 9M2025 與 424B4 FY2025 同為含 xAI／X 的合併口徑，可相減得 2025Q4")
k2 = find(B, lambda l: "retrospective combination of the companies for all periods presented to include" in l)
chk("b.recast 第二處（Summary 財務資料前言）", 1, 1, loc(B, k2))

# ============ 2. 424B4：FY2025／FY2024 分部營收、營業利益（MD&A 分部表 vs 文字句） ============
SEGTBL = {"space": "Space\n", "conn": "Connectivity\n", "si": "AI\n"}
seg_t0 = find(B, lambda l: l.strip() == "Segment Results", find(B, lambda l: "2025 vs. 2024 Change" in l))
for sg, title in (("space", "Space"), ("conn", "Connectivity"), ("si", "AI")):
    t = find(B, lambda l, title=title: l.strip() == title, seg_t0)
    rv, ri = row(B, "Revenue", t, 12, n=3, first=True)
    ev, ei = row(B, "Income (loss) from operations" if sg == "space" else ("Income from operations" if sg == "conn" else "Loss from operations"), t, 20, n=3, first=True)
    for key, vals, ii, lab in (("rev", rv, ri, "營收"), ("ebit", ev, ei, "營業利益（損失）")):
        v25, v24, chg = vals[0], vals[1], vals[2]
        # 第二讀（FY2025）：前段 Prospectus Summary 文字句；（FY2024）：同表 2025 − $ Change 欄
        if key == "rev":
            pat = {"space": r"In 2025, our Space segment generated revenue of \$([\d,]+) million", "conn": r"generated revenue of \$([\d,]+) million, income from operations of \$4,423",
                   "si": r"In 2025, our AI segment generated revenue of \$([\d,]+) million"}[sg]
        else:
            pat = {"space": r"In 2025, our Space segment generated revenue of \$4,086 million, loss from operations of \$\((\d[\d,]*)\)",
                   "conn": r"income from operations of \$([\d,]+) million, and Segment Adjusted EBITDA of \$7,168",
                   "si": r"In 2025, our AI segment generated revenue of \$3,201 million, loss from operations of \$\((\d[\d,]*)\)"}[sg]
        s2, k2, q2 = search(B, pat)
        if key == "ebit" and sg in ("space", "si"): s2 = -s2
        add(f"seg.{sg}.{key}.fy2025", f"{SEGN[sg]}：{lab}（FY2025）", v25, U, "fy2025", loc(B, ii, f"MD&A Segment Results {title} 表"), quote=re.sub(r"\s{2,}", " ", D[B][0][ii].strip()),
            second=s2, second_src=loc(B, k2, "Prospectus Summary 文字句"), docDate=DOCS[B][3], note="原文取代 v0.1 二手值（G5）")
        add(f"seg.{sg}.{key}.fy2024", f"{SEGN[sg]}：{lab}（FY2024）", v24, U, "fy2024", loc(B, ii, f"MD&A Segment Results {title} 表"),
            second=v25 - chg, second_src="同表 2025 欄 − $ Change 欄", docDate=DOCS[B][3])
ct = find(B, lambda l: "The following table sets forth our consolidated statements of operations data for the periods indicated" in l)
for key, lab, rl in (("rev", "合併營收", "Revenue"), ("ebit", "合併營業利益（損失）", "Income (loss) from operations")):
    v, ii = row(B, rl, ct, 20, n=3, first=True)
    sv, si_ = row(B, rl, find(B, lambda l: l.strip() == "Statements of Operations Data:"), 12, n=5)
    add(f"seg.total.{key}.fy2025", f"{lab}（FY2025）", v[0], U, "fy2025", loc(B, ii, "MD&A 合併損益表"), quote=re.sub(r"\s{2,}", " ", D[B][0][ii].strip()),
        second=sv[2], second_src=loc(B, si_, "Summary 合併損益資料（第 3 欄）"), docDate=DOCS[B][3])
    add(f"seg.total.{key}.fy2024", f"{lab}（FY2024）", v[1], U, "fy2024", loc(B, ii, "MD&A 合併損益表"), second=sv[3], second_src=loc(B, si_, "Summary（第 4 欄）"), docDate=DOCS[B][3])

# ============ 3. 424B4：分部調整後 EBITDA（MD&A 文字句 vs Summary 分部營運資料表） ============
kpi0 = find(B, lambda l: l.strip() == "Segment Operating and Financial Data (unaudited)")
for sg, nm in (("space", "Space"), ("conn", "Connectivity"), ("si", "AI")):
    v25, k1, q1 = search(B, rf"{nm} Segment Adjusted EBITDA for 2025 (?:in|de)creased by \$[\d,]+ million to \$\(?([\d,]+)\)? million")
    v24, _, _ = search(B, rf"{nm} Segment Adjusted EBITDA for 2025 (?:in|de)creased by \$[\d,]+ million to \$\(?[\d,]+\)? million compared to \$\(?([\d,]+)\)?")
    if sg == "si": v25 = -v25
    t = find(B, lambda l, nm=nm: l.strip() == f"{nm}:", kpi0)
    tv, ti = row(B, "millions) (2)", t, 14, n=5)
    add(f"seg.{sg}.adjEbitda.fy2025", f"{SEGN[sg]}：分部調整後 EBITDA（FY2025）", v25, U, "fy2025", loc(B, k1, "MD&A Segment Adjusted EBITDA 文字句"), quote=q1, tag="[Interested-party]",
        second=tv[2], second_src=loc(B, ti, "Summary 分部營運資料表"), docDate=DOCS[B][3], note="非 GAAP；原文取代 v0.1（SI 為衍生值）")
    add(f"seg.{sg}.adjEbitda.fy2024", f"{SEGN[sg]}：分部調整後 EBITDA（FY2024）", v24, U, "fy2024", loc(B, k1, "MD&A 文字句"), tag="[Interested-party]",
        second=tv[3], second_src=loc(B, ti, "Summary 分部營運資料表"), docDate=DOCS[B][3])
v, k1, q1 = search(B, r"In 2025, we generated revenue on a consolidated basis of \$18,674 million, loss from operations of \$\(2,589\) million and Adjusted EBITDA of \$([\d,]+)")
v2, k2, _ = search(B, r"In 2025, we generated revenue on a consolidated basis of \$18,674 million, loss from operations of \$\(2,589\) million and Adjusted EBITDA of \$([\d,]+)", nth=1)
add("seg.total.adjEbitda.fy2025", "合併調整後 EBITDA（FY2025）", v, U, "fy2025", loc(B, k1, "Prospectus Summary"), quote=q1, tag="[Interested-party]", second=v2, second_src=loc(B, k2, "Business 概要同句"), docDate=DOCS[B][3])
add("seg.total.adjEbitda.fy2024", "合併調整後 EBITDA（FY2024，三分部加總）", V("seg.space.adjEbitda.fy2024") + V("seg.conn.adjEbitda.fy2024") + V("seg.si.adjEbitda.fy2024"), U, "fy2024",
    "424B4 三分部文字句加總", tag="[Derived]", note="FY2024 合併數未見原句；分部合計＝合併（2026 各期成立）")

# ============ 4. 424B4：分部資本支出（MD&A 表 vs 文字句；合計另與 Summary 表互核） ============
cx0 = find(B, lambda l: l.strip() == "Capital Expenditures", find(B, lambda l: "AI Segment Adjusted EBITDA for 2025" in l))
for sg, lab, nm in (("space", "Space", "Space"), ("conn", "Connectivity", "Connectivity"), ("si", "AI", "AI")):
    tv, ti = row(B, lab, cx0, 12, n=5)
    v25, k1, q1 = search(B, rf"{nm} capital expenditures for 2025 increased \$[\d,]+ million to \$([\d,]+) million compared to \$[\d,]+")
    v24, _, _ = search(B, rf"{nm} capital expenditures for 2025 increased \$[\d,]+ million to \$[\d,]+ million compared to \$([\d,]+)")
    add(f"seg.{sg}.capex.fy2025", f"{SEGN[sg]}：資本支出（FY2025）", tv[2], U, "fy2025", loc(B, ti, "MD&A Capital Expenditures 表"), quote=re.sub(r"\s{2,}", " ", D[B][0][ti].strip()),
        second=v25, second_src=loc(B, k1, "MD&A 文字句"), docDate=DOCS[B][3],
        note=("v0.1 A 段審查判 3,832 為媒體混用而改「未取得」；424B4 原文確為 3,832（與 9M2026 Space 資本支出 3,832 數字相同屬巧合）" if sg == "space" else
              ("v0.1 標「待核」，原文相符，解除" if sg == "conn" else "")))
    add(f"seg.{sg}.capex.fy2024", f"{SEGN[sg]}：資本支出（FY2024）", tv[3], U, "fy2024", loc(B, ti, "MD&A Capital Expenditures 表"), second=v24, second_src=loc(B, k1, "MD&A 文字句"), docDate=DOCS[B][3])
tv, ti = row(B, "Total Capital Expenditures", cx0, 12, n=5)
sv, si_ = row(B, "Total Capital Expenditures", find(B, lambda l: l.strip() == "Capital Expenditures:"), 24, n=5)
add("seg.total.capex.fy2025", "資本支出合計（FY2025）", tv[2], U, "fy2025", loc(B, ti, "MD&A Capital Expenditures 表"), second=sv[2], second_src=loc(B, si_, "Summary 資本支出表"), docDate=DOCS[B][3])
add("seg.total.capex.fy2024", "資本支出合計（FY2024）", tv[3], U, "fy2024", loc(B, ti, "MD&A Capital Expenditures 表"), second=sv[3], second_src=loc(B, si_, "Summary 資本支出表"), docDate=DOCS[B][3])

# ============ 5. 424B4：Space 子列（Launch Services 占比 → 衍生）、KPI ============
lp, li = row(B, "Launch Services", find(B, lambda l: "Revenue from Launch Services recognized at point in time" in l), 10, n=5)
ld_, ldi = row(B, "Launch & Development", li, 3, n=5)
add("seg.space.launchSvc.pct.fy2025", "Space：Launch Services 占 Space 營收（FY2025）", lp[2], "%", "fy2025", loc(B, li, "Description of Our Segments 收入占比表"),
    second=100 - ld_[2], second_src="同表 100% − Launch & Development 37.0%", tol=0.05, docDate=DOCS[B][3])
add("seg.space.launchSvc.fy2025", "Space：Launch Services（FY2025，衍生＝營收 × 63.0%）", round(V("seg.space.rev.fy2025") * lp[2] / 100, 1), U, "fy2025", "424B4 占比 × Space 營收",
    tag="[Derived]", note="占比四捨五入到 0.1%，誤差 ±2；Connectivity、SI 的 FY2025 子列 424B4 未揭露（已搜尋 MD&A 分部段與 Business，只有增量）")
add("seg.space.launchDev.fy2025", "Space：Launch & Development（FY2025，衍生＝營收 − Launch Services）", round(V("seg.space.rev.fy2025") - V("seg.space.launchSvc.fy2025"), 1), U, "fy2025",
    "同上", tag="[Derived]", second=round(V("seg.space.rev.fy2025") * ld_[2] / 100, 1), second_src="營收 × 37.0%", tol=0.5)
gk = find(B, lambda l: l.strip().startswith("gigawatts) (1)"))
gv = trailing(D[B][0][gk].split(")", 1)[1].replace("(1)", ""))
gv2 = [float(x) for x in re.findall(r"(?<![\d.])(\d+(?:\.\d+)?)(?![\d.])", D[B][0][gk].split("(1)")[1])]
for j, per in ((2, "fy2025"), (3, "fy2024")):
    add(f"seg.si.kpi.gw.{per}", f"SI：名目算力（{'2025-12-31' if per == 'fy2025' else '2024-12-31'}，期末）", gv[j], "GW", per, loc(B, gk, "Summary 分部營運資料表 Nameplate compute draw"),
        quote=re.sub(r"\s{2,}", " ", (D[B][0][gk - 1] + D[B][0][gk]).strip()), tag="[Interested-party]", second=gv2[j], second_src="同列以第二種讀法（逐字切分）", tol=0.001, docDate=DOCS[B][3],
        note="公司自訂 KPI；v0.1 用內插 0.8（2025）與判斷 0.15（2024）")
chk("seg.si.kpi.gw.2026q1（424B4 1.0 vs v0.1 Q2 新聞稿 1.0）", gv[0], V01["seg.si.kpi.gw.2026q1"]["value"], "424B4 vs Q2 新聞稿", tol=0.05)
for key, lab, nm, per_col in (("kpi.subs", "Starlink Subscribers (in millions)", "Starlink 訂戶（期末，百萬）", 2), ("kpi.arpu", "Starlink ARPU (dollars per month)", "Starlink ARPU（每月，美元）", 2)):
    tv, ti = row(B, lab, kpi0, 30, n=5)
    add(f"seg.conn.{key}.fy2025", f"Connectivity：{nm}（FY2025）", tv[2], "百萬戶" if key == "kpi.subs" else "USD/月", "fy2025", loc(B, ti, "Summary 分部營運資料表"), tag="[Interested-party]", docDate=DOCS[B][3],
        second=(81 if key == "kpi.arpu" else None), second_src=("MD&A Key Business Metrics 文字句（$91 → $81）" if key == "kpi.arpu" else ""))
tv, ti = row(B, "Launches (number) (1)", kpi0, 10, n=5)
add("seg.space.kpi.launchTot.fy2025", "Space：發射次數合計（FY2025）", tv[2], "次", "fy2025", loc(B, ti, "Summary 分部營運資料表"), tag="[Interested-party]", docDate=DOCS[B][3])
tv, ti = row(B, "activities", find(B, lambda l: l.strip() == "Statement of Cash Flows Data:"), 8, n=5)
add("cf.cfo.fy2025", "營業現金流（FY2025）", tv[2], U, "fy2025", loc(B, ti, "Summary 現金流量資料"), docDate=DOCS[B][3])
bs0 = find(B, lambda l: l.strip() == "Balance Sheet Data:")
tv, ti = row(B, "Cash and cash equivalents", bs0, 10, n=3)
add("bs.cash.20251231.b", "現金及約當現金（2025-12-31；424B4）", tv[1], U, "20251231", loc(B, ti, "Summary 資產負債資料"), second=V01["bs.cash.20251231"]["value"],
    second_src="10-Q 2026Q3 資產負債表比較欄（v0.1 事實總帳 bs.cash.20251231）", docDate=DOCS[B][3])

# ---- 2025Q4（Derived）＝FY2025 − 9M2025（10-Q 2026Q3 比較欄；兩者皆為重編後合併口徑，見 b.recast）
for sg in ("space", "conn", "si", "total"):
    for key in ("rev", "ebit", "adjEbitda", "capex"):
        a, b9 = V(f"seg.{sg}.{key}.fy2025"), V01[f"seg.{sg}.{key}.9m2025"]["value"]
        add(f"seg.{sg}.{key}.2025q4", f"{SEGN[sg]}：{dict(rev='營收', ebit='營業利益', adjEbitda='分部調整後 EBITDA', capex='資本支出')[key]}（2025Q4，衍生）", a - b9, U, "2025q4",
            f"FY2025（424B4）− 9M2025（10-Q 2026Q3 比較欄）", tag="[Derived]")
for k_ in ("launchSvc", "launchDev"):
    add(f"seg.space.{k_}.2025q4", f"Space：{k_}（2025Q4，衍生）", round(V(f"seg.space.{k_}.fy2025") - V01[f"seg.space.{k_}.9m2025"]["value"], 1), U, "2025q4", "FY2025（衍生）− 9M2025", tag="[Derived]")
add("seg.si.kpi.gw.2025q4", "SI：名目算力（2025-12-31）", V("seg.si.kpi.gw.fy2025"), "GW", "2025q4", "同 seg.si.kpi.gw.fy2025", tag="[Interested-party]")

# ---- 對帳（FY2025、FY2024、2025Q4：三分部合計 − 合併＝0）
RECON = []
for per in ("fy2025", "fy2024", "2025q4"):
    for key in ("rev", "ebit", "adjEbitda", "capex"):
        s3 = sum(V(f"seg.{s}.{key}.{per}") for s in ("space", "conn", "si"))
        t = V(f"seg.total.{key}.{per}")
        RECON.append({"period": per, "line": key, "segSum": s3, "total": t, "diff": round(s3 - t, 6)})
        chk(f"recon.{key}.{per}（三分部合計 − 合併）", s3, t, "分部對帳", tol=1e-6)
chk("recon.space.子列.fy2025（衍生子列合計 − Space）", V("seg.space.launchSvc.fy2025") + V("seg.space.launchDev.fy2025"), V("seg.space.rev.fy2025"), "子列對帳", tol=1e-6)
# 424B4 的 2025Q1（重編後）與 Q2 新聞稿 1H2025 − Q2 2025 互核：確認 v0.1 季度欄與 424B4 同口徑
q1v, q1i = row(B, "Revenue", find(B, lambda l: l.strip() == "Statements of Operations Data:"), 12, n=5)
chk("424B4 2025Q1 合併營收 vs Q2 新聞稿 1H2025 − 2025Q2（同為重編口徑）", q1v[1], V01["seg.total.rev.1h2025"]["value"] - V01["seg.total.rev.2025q2"]["value"], loc(B, q1i))

# ============ 6. 2026Q2 10-Q 附註（分部、收入拆分、名目算力、客戶、債務、承諾、股權報酬） ============
Q2 = "Q2"
n18 = find(Q2, lambda l: l.strip() == "Note 18 - Segments")
blk = {}
for i in range(n18, n18 + 140):
    s = D[Q2][0][i].strip()
    m = re.match(r"(Three|Six) Months Ended June 30,$", s)
    if m and D[Q2][0][i + 1].strip() in ("2026", "2025"):
        blk[("2026q2" if m.group(1) == "Three" else "1h2026") if D[Q2][0][i + 1].strip() == "2026" else ("2025q2" if m.group(1) == "Three" else "1h2025")] = i
for per, b0 in blk.items():
    for key, lab in (("rev", "Revenue"), ("ebit", "Income (loss) from operations"), ("da", "Depreciation and amortization"), ("sbc", "Share-based compensation"), ("capex", "Capital expenditures")):
        v, ii = row(Q2, lab, b0, 34, n=4)
        for j, sg in enumerate(("space", "conn", "si", "total")):
            vid = f"seg.{sg}.{key}.{per}"
            sec_ = V01.get(vid)
            if sec_ is None and sg == "total" and key in ("da", "sbc"):
                sec_ = {"value": sum(V01[f"seg.{s}.{key}.{per}"]["value"] for s in ("space", "conn", "si")), "source": "三分部相加"}
            add(f"q2q.{vid}", f"{SEGN[sg]}：{dict(rev='營收', ebit='營業利益', da='折舊攤銷', sbc='股權報酬（一般）', capex='資本支出')[key]}（{per}；2026Q2 10-Q）", v[j], U, per,
                loc(Q2, ii, "附註 18 Segments 表"), quote=re.sub(r"\s{2,}", " ", D[Q2][0][ii].strip())[:160], docDate=DOCS[Q2][3],
                second=(sec_["value"] if sec_ else None), second_src=("v0.1 事實總帳 " + vid + "（新聞稿）" if sec_ else ""))
n3 = find(Q2, lambda l: l.strip().startswith("Launch Services") and "648" in l)
for key, lab in (("launchSvc", "Launch Services"), ("launchDev", "Launch & Development"), ("consumer", "Consumer"), ("entGov", "Enterprise & Government"), ("ads", "Advertising"), ("siSI", "AI Solutions & Infrastructure")):
    v, ii = row(Q2, lab, n3, 12, n=4)
    sg = {"launchSvc": "space", "launchDev": "space", "consumer": "conn", "entGov": "conn", "ads": "si", "siSI": "si"}[key]
    for j, per in enumerate(("2026q2", "2025q2", "1h2026", "1h2025")):
        vid = f"seg.{sg}.{key}.{per}"
        add(f"q2q.{vid}", f"{SEGN[sg]}：{key}（{per}；2026Q2 10-Q 附註 3）", v[j], U, per, loc(Q2, ii, "附註 3 收入拆分表"), docDate=DOCS[Q2][3],
            second=V01[vid]["value"], second_src="v0.1 事實總帳（新聞稿）")
gq = find(Q2, lambda l: l.strip().startswith("Nameplate compute draw") and len(trailing(l)) == 2)
add("q2q.seg.si.kpi.gw.2026q2", "SI：名目算力（2026-06-30；2026Q2 10-Q）", trailing(D[Q2][0][gq])[0], "GW", "2026q2", loc(Q2, gq, "MD&A Key Business Metrics"), tag="[Interested-party]",
    second=V01["seg.si.kpi.gw.2026q2"]["value"], second_src="v0.1 事實總帳（新聞稿）", tol=0.001, docDate=DOCS[Q2][3])
cb = find(Q2, lambda l: l.strip().startswith("Customer B"))
cbv = [float(x) for x in re.findall(r"([\d.]+) %", D[Q2][0][cb])]
add("q2q.cust.B.pct.2026q2", "客戶 B 占合併營收（2026Q2；全屬 AI／SI）", cbv[0], "%", "2026q2", loc(Q2, cb, "附註 3 Concentration of risk"), quote=re.sub(r"\s{2,}", " ", D[Q2][0][cb].strip()), docDate=DOCS[Q2][3])
add("q2q.cust.B.rev.2026q2", "客戶 B 營收（2026Q2，衍生＝19.5% × 合併營收）", round(cbv[0] / 100 * V("q2q.seg.total.rev.2026q2")), U, "2026q2", "附註 3 衍生", tag="[Derived]",
    note="÷ AI S&I 2,194＝69.5%（v0.1 用 1.526 ÷ 2.194＝69.6%，一致）")
v, kk, q = search(Q2, r"Backlog totaled .{0,40}?\$([\d,]+) million as of June 30, 2026", n=14)
add("q2q.backlog.20260630", "Backlog（2026-06-30）", v, U, "20260630", loc(Q2, kk, "附註 3 Backlog"), quote=q, docDate=DOCS[Q2][3])
n9 = find(Q2, lambda l: l.strip() == "Note 9 - Debt")
for key, lab in (("spacexNotes", "SpaceX Notes"), ("otherFin", "Other financings (1)"), ("totalDebt", "Total debt"), ("finLease", "Finance lease liability")):
    v, ii = row(Q2, lab, n9, 14, n=3)
    add(f"q2q.debt.{key}.principal.20260630", f"債務：{lab}（本金，2026-06-30）", v[0], U, "20260630", loc(Q2, ii, "附註 9 Debt 表"), docDate=DOCS[Q2][3])
chk("q2q.debt 組成合計（Notes＋X 27＋其他融資 ＝ 債務合計）", V("q2q.debt.spacexNotes.principal.20260630") + 27 + V("q2q.debt.otherFin.principal.20260630"), V("q2q.debt.totalDebt.principal.20260630"), "附註 9 同表")
v1, k1, q1 = search(Q2, r"recorded the related debt of \$([\d,]+) million and \$[\d,]+ million within")
v2, _, _ = search(Q2, r"recorded the related debt of \$[\d,]+ million and \$([\d,]+) million within")
w1, k3, _ = search(Q2, r"As of June 30, 2026, the Company recorded debt of \$([\d,]+) million and \$[\d,]+ million within")
w2, _, _ = search(Q2, r"As of June 30, 2026, the Company recorded debt of \$[\d,]+ million and \$([\d,]+) million within")
add("q2q.debt.failedSLB.total.20260630", "失敗售後租回（Valor）債務合計（2026-06-30＝流動＋非流動）", v1 + v2, U, "20260630", loc(Q2, k1, "附註 9 Other Financings"), quote=q1,
    second=w1 + w2, second_src=loc(Q2, k3, "附註 17 Valor 段"), docDate=DOCS[Q2][3])
add("q2q.debt.failedSLB.repaid.2026q3", "失敗售後租回本金 Q3 淨減少（6-30 13,329 − 9-30 12,897；Q3 未見新增交易）", V("q2q.debt.failedSLB.total.20260630") - V01["debt.failedSLB.total.20260930"]["value"], U, "2026q3",
    "2026Q2 10-Q 附註 9 − 2026Q3 10-Q 附註 9", tag="[Derived]", note="v0.1 資料缺口（9M 失敗售後租回本金償還未單獨揭露）部分補上：Q3 約 432；1H 因 4 月新增交易無法拆")
mt = find(Q2, lambda l: l.strip().startswith("2026 (remaining six months)"), n9)
mv = [trailing(D[Q2][0][mt + j])[-1] for j in range(7)]
for j, y in enumerate(("2026H2", "2027", "2028", "2029", "2030", "thereafter", "total")):
    add(f"q2q.debt.maturity.{y}", f"債務本金到期（2026-06-30 表）：{y}", mv[j], U, "20260630", loc(Q2, mt + j, "附註 9 到期表"), docDate=DOCS[Q2][3])
chk("q2q.debt.maturity 合計 ＝ 各年相加 ＝ 債務合計", sum(mv[:6]), V("q2q.debt.totalDebt.principal.20260630"), "附註 9")
n16 = find(Q2, lambda l: l.strip() == "Note 16 - Commitments and Contingencies")
ct_ = find(Q2, lambda l: l.strip().startswith("2026 (remaining six months)"), n16)
cv = [trailing(D[Q2][0][ct_ + j])[-1] for j in range(7)]
for j, y in enumerate(("2026H2", "2027", "2028", "2029", "2030", "thereafter", "total")):
    add(f"q2q.commit.uncond.{y}", f"不可取消承諾（2026-06-30 表）：{y}", cv[j], U, "20260630", loc(Q2, ct_ + j, "附註 16 Unconditional Obligations"), docDate=DOCS[Q2][3])
chk("q2q.commit 合計", sum(cv[:6]), cv[6], "附註 16")
sb = find(Q2, lambda l: l.strip().startswith("Total") and "831" in l)
add("q2q.sbc.total.2026q2", "股權報酬合計（2026Q2；附註 14 依科目）", trailing(D[Q2][0][sb])[0], U, "2026q2", loc(Q2, sb, "附註 14"), docDate=DOCS[Q2][3],
    second=V("q2q.seg.total.sbc.2026q2"), second_src="附註 18 分部股權報酬合計")
v, kk, q = search(Q2, r"effective interest rate on the SpaceX Notes was ([\d.]+)% as of June 30, 2026")
add("q2q.debt.spacexNotes.eir.20260630", "SpaceX Notes 有效利率（2026-06-30）", v, "%", "20260630", loc(Q2, kk, "附註 9"), quote=q, docDate=DOCS[Q2][3])
v, kk, q = search(Q2, r"\$([\d,]+) million and \$[\d,]+ million in Interest expense for the three and six months ended June 30, 2026")
add("q2q.is.intExpRP.2026q2", "利息費用中的 Valor 關係人部分（2026Q2）", v, U, "2026q2", loc(Q2, kk, "附註 17"), quote=q, docDate=DOCS[Q2][3])
cash6, ci = row(Q2, "Cash and cash equivalents", find(Q2, lambda l: l.strip().startswith("Cash and cash equivalents") and len(trailing(l)) == 2), 2, n=2)
add("q2q.bs.cash.20260630", "現金及約當現金（2026-06-30）", cash6[0], U, "20260630", loc(Q2, ci, "資產負債表"), docDate=DOCS[Q2][3], second=V01["bs.cash.20260630"]["value"], second_src="v0.1 事實總帳（Q3 新聞稿／10-Q 比較）") if "bs.cash.20260630" in V01 else \
    add("q2q.bs.cash.20260630", "現金及約當現金（2026-06-30）", cash6[0], U, "20260630", loc(Q2, ci, "資產負債表"), docDate=DOCS[Q2][3])

# ============ 7. 法說指引（逐字稿不入庫：數字＋≤10 字英文短句；本機有逐字稿時重讀核對） ============
# (鍵, 項目, 值, 單位, 期間, 逐字稿, 頁, 短句, 第二讀規則 regex（在逐字稿中找、取 group(1)）, 說明)
GUIDE = [
    ("guide.capex.2026q4.lo", "2026Q4 合併資本支出指引：下緣", 35.0, "USD B", "2026q4", "T3", 8, "range of $35 billion to $40 billion",
     r"capital\s+expenditures to now be in the range of \$(\d+) billion to \$\d+ billion", "CFO；G1 中值 37.5（Interested-party）"),
    ("guide.capex.2026q4.hi", "2026Q4 合併資本支出指引：上緣", 40.0, "USD B", "2026q4", "T3", 8, "range of $35 billion to $40 billion",
     r"capital\s+expenditures to now be in the range of \$\d+ billion to \$(\d+) billion", ""),
    ("guide.gw.ye2026.lo", "2026 年底名目算力指引：下緣", 2.5, "GW", "20261231", "T3", 8, "in the range of 2.5 to 3 gigawatts",
     r"nameplate compute in the range of ([\d.]+) to [\d.]+\s*gigawatts", "CFO；Musk 同段亦稱 2.5 to 3 GW（第 3 頁）；G1 中值 2.75"),
    ("guide.gw.ye2026.hi", "2026 年底名目算力指引：上緣", 3.0, "GW", "20261231", "T3", 8, "in the range of 2.5 to 3 gigawatts",
     r"nameplate compute in the range of [\d.]+ to ([\d.]+)\s*gigawatts", ""),
    ("guide.gw.ye2026.musk", "2026 年底名目算力（Musk 第二讀：maybe 2.5 to 3 GW）", 3.0, "GW", "20261231", "T3", 3, "maybe 2.5 to 3 GW",
     r"maybe [\d.]+ to ([\d.]+) GW", "第二讀（上緣）"),
    ("guide.gw.oct2026", "2026-10 底名目算力（nearly 2 GW）", 2.0, "GW", "20261031", "T3", 8, "nearly 2 GW of nameplate compute",
     r"reach nearly (\d+) GW", "Recon"),
    ("guide.rev.2027q1.lo", "2027Q1 合併營收指引：下緣", 33.0, "USD B", "2027q1", "T3", 8, "range of $33 billion to $37 billion",
     r"revenues in the range of \$(\d+) billion to \$\d+ billion for Q1", "G2：只作校準檢查"),
    ("guide.rev.2027q1.hi", "2027Q1 合併營收指引：上緣", 37.0, "USD B", "2027q1", "T3", 8, "range of $33 billion to $37 billion",
     r"revenues in the range of \$\d+ billion to \$(\d+) billion for Q1", ""),
    ("guide.arr.dec2026", "2026 年 12 月 ARR（＝12 月單月營收 × 12）", 100.0, "USD B", "202612", "T3", 8, "$100 billion in ARR by the end",
     r"goal of \$(\d+) billion in ARR", "G2；定義：annualized revenue based on the month of December"),
    ("guide.arr.dec2026.q2call", "2026 年 12 月 ARR（Q2 法說第二讀）", 100.0, "USD B", "202612", "T2", 7, "$100 billion of ARR",
     r"reach \$(\d+) billion of ARR", "第二讀"),
    ("guide.gw.add2027", "2027 新增名目算力（Musk：高度有信心）", 10.0, "GW", "2027", "T3", 3, "incremental 10 GW of compute online",
     r"incremental (\d+) GW of compute online", "G3：不入基準，只作 Sens 管理層路徑案"),
    ("guide.gw.ye2027", "2027 年底名目算力（Musk：13 GW）", 13.0, "GW", "20271231", "T3", 10, "so 13GW total",
     r"next year, so\s+(\d+)\s*GW total", "第二讀：第 3 頁「less than 13 GW」"),
    ("guide.gw.ye2027.b", "2027 年底名目算力（第二讀：surprised if less than 13 GW）", 13.0, "GW", "20271231", "T3", 3, "less than 13 GW of compute",
     r"less than (\d+) GW of compute", ""),
    ("guide.gw.ye2027.stretch", "2027 年 SI 算力 20 GW 以上（機率上升）", 20.0, "GW", "2027", "T3", 10, "some number over 20 GW",
     r"some number over (\d+) GW", "只作 Recon"),
    ("guide.cursor.equity", "Cursor 收購股權價值（法說口徑）", 60.0, "USD B", "2026-08-14", "T3", 8, "of the $60 billion equity value",
     r"of the \$(\d+) billion equity value for Cursor", "G6：只列 Recon（10-Q GAAP 對價 42.474＋收購後服務期 23.376 為不同口徑）"),
    ("guide.cursor.retention", "Cursor 留任與替換獎酬（法說口徑）", 23.0, "USD B", "2026-08-14", "T3", 8, "$23 billion was attributed toward",
     r"\$(\d+) billion was attributed toward retaining", "約三年認列（roughly the next three-year）；10-Q 收購後服務期 23.376、加權 2.8 年"),
    ("guide.backlog.total", "Backlog（法說，2026-09-30）", 63.0, "USD B", "20260930", "T3", 8, "backlog totaling $63 billion",
     r"backlog totaling \$(\d+) billion", "G7：Recon（10-Q 62.959）"),
    ("guide.backlog.deferred", "Backlog 中已列遞延收入", 17.5, "USD B", "20260930", "T3", 8, "$17.5 billion was recognized as",
     r"of which \$([\d.]+) billion was recognized as", "G7"),
    ("guide.backlog.1y", "Backlog 一年內認列", 42.0, "USD B", "20260930", "T3", 8, "Approximately $42 billion of this backlog",
     r"Approximately \$(\d+) billion of this backlog", "G7（10-Q：約 67% 一年內）"),
    ("guide.price.perWatt.hi", "算力定價：略高於每瓦 30–50 美元區間上緣", 50.0, "USD/W", "2026q3", "T3", 7, "slightly above the high end",
     r"high end of the\s+\$\d+ to \$(\d+) per watt", "只作 Recon（口徑未說明：合約總價 ÷ 瓦？）"),
    ("guide.gw.ye2026.q2call", "2026 年底名目算力（Q2 法說：over 2 GW）", 2.0, "GW", "20261231", "T2", 6, "end this year at over 2 GW",
     r"end this year at over (\d+) GW", "v0.1 Recon 的二手值，現以逐字稿原文代替"),
]
for iid, item, val, unit, per, td, pg, quote, pat, note in GUIDE:
    assert len(quote.split()) <= 10, quote
    src = f"{TR[td][1]} 第 {pg} 頁"
    second, ss = None, "逐字稿未入庫，未重讀（數字沿用本檔）"
    if D[td] is not None:
        v, kk, _ = search(td, pat, n=3)
        second, ss = v, f"逐字稿重讀（第 {D[td][1][kk]} 頁）"
        if D[td][1][kk] != pg: raise SystemExit(f"{iid} 頁碼不符：{D[td][1][kk]} vs {pg}")
        if quote.replace("$", "") not in re.sub(r"\s+", " ", " ".join(D[td][0])).replace("$", "").replace("  ", " "):
            # 允許逐字稿斷行或「3GW」黏字
            if re.sub(r"\s", "", quote) not in re.sub(r"\s", "", " ".join(D[td][0])): raise SystemExit(f"{iid} 短句不在逐字稿：{quote}")
    add(iid, item, val, unit, per, src, quote=quote, tag="[Interested-party]", note=note, second=second, second_src=ss, tol=1e-9)
    if second is None: IDX[iid]["check"] = {"value": None, "source": ss, "ok": None}
chk("guide.gw.ye2026 兩讀（CFO 上緣 vs Musk 上緣）", V("guide.gw.ye2026.hi"), V("guide.gw.ye2026.musk"), "逐字稿第 8 頁 vs 第 3 頁", tol=1e-9)
chk("guide.gw.ye2027 兩讀", V("guide.gw.ye2027"), V("guide.gw.ye2027.b"), "逐字稿第 10 頁 vs 第 3 頁", tol=1e-9)
chk("guide.arr 兩讀（Q3 法說 vs Q2 法說）", V("guide.arr.dec2026"), V("guide.arr.dec2026.q2call"), "兩次法說", tol=1e-9)
GQ = [("guide.gm.q4", "2026Q4 合併毛利率：回到與 2026Q2 一致（Q3 有一次性利益與客戶組合）", "T3", 8, "more consistent with the second quarter", r"(more consistent with the second quarter)", "G4：k0 以 2026Q2 校準列 Sens"),
      ("guide.capex.q2call", "Q2 法說：下兩季資本支出與 Q2 相近（已被 Q3 法說 35–40 取代）", "T2", 9, "very similar to the current quarter", r"(very\s+similar to the current quarter)", "v0.1 未取得；只作紀錄")]
for iid, item, td, pg, quote, pat, note in GQ:
    assert len(quote.split()) <= 10
    ok = None
    if D[td] is not None:
        search_ok = re.search(pat, re.sub(r"\s+", " ", " ".join(D[td][0])))
        if not search_ok: raise SystemExit(f"{iid} 找不到短句")
        ok = True
    add(iid, item, None, "文字", "", f"{TR[td][1]} 第 {pg} 頁", quote=quote, tag="[Interested-party]", note=note)

# ============ 輸出 ============
bad = [c for c in CHK if c["ok"] is False]
n_tr = sum(1 for it in ITEMS if it["id"].startswith("guide.") and it.get("check", {}).get("ok") is True)
out = {"meta": {"generated_by": "spacex/facts_v02.py", "sources": {k: v[1] for k, v in DOCS.items()}, "transcripts": {k: v[1] for k, v in TR.items()},
                "transcripts_reread": n_tr, "note": "法說逐字稿不入庫；guide.* 只存數字與短句"},
       "items": ITEMS, "checks": CHK, "recon": RECON}
json.dump(out, open(os.path.join(HERE, "data", "spacex_facts_v02.json"), "w", encoding="utf-8"), ensure_ascii=False, indent=1)
print(f"項目 {len(ITEMS)}；兩讀與核對 {len(CHK)} 筆，不一致 {len(bad)}；逐字稿重讀 {n_tr} 筆；對帳 {len(RECON)} 列（最大差額 {max(abs(r['diff']) for r in RECON)}）")
for c in bad: print("  不一致：", c)
sys.exit(1 if bad or any(abs(r["diff"]) > 1e-6 for r in RECON) else 0)
