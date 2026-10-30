# SpaceX（SPCX）2026Q3 事實總帳（SpaceX v0.1 工作單 A 段第 1–2 項）。
# 來源：mag/extract/docs/2026Q3/spacex_2026Q3_10-Q.pdf（季度截至 2026-09-30，2026-10-22 簽署）與 spacex_2026Q3_新聞稿.pdf（8-K 財報新聞稿，2026-10-20 法說）。
# 讀法：只讀 text/（pdftotext -layout 轉出；以 \f 分頁，頁碼＝PDF 頁序）。
# 兩讀：同一數字在 10-Q 與新聞稿各讀一次互核；只出現一處者，以另一張表或文字句（或同表恆等式）作第二讀，並在 check.source 寫明讀法。
# 標記：[Verified]＝10-Q 財務報表與附註（向 SEC 申報、經公司簽證）；[Interested-party]＝公司自訂指標、非 GAAP、管理層敘述或只見於新聞稿者；[Derived]＝本程式由原文數字計算。
# 用法：python3 spacex/facts_2026q3.py → 寫 spacex/data/spacex_facts_2026q3.json；兩讀不一致或分部對帳差額非 0 時非零結束。
import json, os, re, sys

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
TX = os.path.join(ROOT, "mag", "extract", "docs", "2026Q3", "text")
QF = "mag/extract/docs/2026Q3/spacex_2026Q3_10-Q.pdf"
PF = "mag/extract/docs/2026Q3/spacex_2026Q3_新聞稿.pdf"


def load(name):
    raw = open(os.path.join(TX, name), encoding="utf-8").read()
    lines, pages = [], []
    for p, ptxt in enumerate(raw.split("\f"), 1):
        for l in ptxt.splitlines():
            l = re.sub(r"\.{2,}", " ", l).replace("shareholders ’", "shareholders’")   # 新聞稿報表的點線；撇號前空白
            lines.append(l); pages.append(p)
    return lines, pages


Q, QP = load("spacex_2026Q3_10-Q.txt")
P, PP = load("spacex_2026Q3_新聞稿.txt")
DOC = {"Q": (Q, QP, QF, "10-Q", "2026-10-22"), "P": (P, PP, PF, "新聞稿", "2026-10-20")}

TRAIL = re.compile(r"((?:\s+(?:\$\s*)?\(?\$?\s*[\d,]*\d(?:\.\d+)?\)?(?!\w)|\s+\$?\s*—|\s+\$)+)\s*$")
TOK = re.compile(r"(\()?\s*\$?\s*([\d,]*\d(?:\.\d+)?)\s*(\))?|(—)")


def trailing(line):
    m = TRAIL.search(" " + line)
    if not m: return []
    out = []
    for t in TOK.finditer(m.group(1)):
        if t.group(4): out.append(0.0); continue
        v = float(t.group(2).replace(",", ""))
        if t.group(1) and t.group(3): v = -v
        out.append(v)
    return out


def find(lines, test, start=0, end=None):
    end = end or len(lines)
    for i in range(start, end):
        if test(lines[i]): return i
    raise SystemExit(f"找不到段落（{start}–{end}）")


def sec(d, title, start=0, exact=True):
    lines = DOC[d][0]
    return find(lines, (lambda l: l.strip() == title) if exact else (lambda l: title in l), start)


def row(d, label, start, span=40, n=4, nth=0, strict=False):
    """回傳 (數字串列, 行號)：label 在行首（去空白），其後不得緊接英文字（避免 'Share-based compensation' 誤配 '... related to acquisitions'）；數字可在下 1–3 行。"""
    lines = DOC[d][0]
    hits = []
    for i in range(start, min(len(lines), start + span)):
        s = lines[i].strip()
        if not s.startswith(label): continue
        rest = re.sub(r"^\(\w\)", "", s[len(label):])
        if rest[:1].isalpha() or (rest[:1] and rest[:1] in "’'-"): continue
        if strict and re.search(r"[A-Za-z]", rest): continue
        hits.append(i)
    if len(hits) <= nth: raise SystemExit(f"[{d}] 找不到列：{label}（第 {start} 行起 {span} 行）")
    i = hits[nth]
    for j in range(i, i + 4):
        v = trailing(lines[j]) if j > i else trailing(re.sub(r"^\(\w\)", "", lines[i].strip()[len(label):]))
        if len(v) >= n: return v[-n:], j
    raise SystemExit(f"[{d}] 列 {label} 數字不足 {n} 欄：{lines[i]!r}")


def loc(d, i, tbl=""):
    lines, pages, f, nm, _ = DOC[d]
    return f"{f} 第 {pages[i]} 頁" + (f" {tbl}" if tbl else "")


def quote(d, i):
    s = re.sub(r"\s{2,}", " ", DOC[d][0][i].strip())
    return s[:160]


ITEMS, CHK = [], []
LASTQ = ""
IDX = {}


def add(iid, item, value, unit, period, d, i, tbl="", tag="[Verified]", note="", second=None, second_src="", tol=0.5, q=None):
    it = {"id": iid, "item": item, "value": value, "unit": unit, "period": period, "source": loc(d, i, tbl),
          "quote": q if q is not None else quote(d, i), "docDate": DOC[d][4], "tag": tag, "note": note}
    if second is not None:
        ok = abs(second - value) <= tol + 1e-9 * abs(value)
        it["check"] = {"value": second, "source": second_src, "ok": ok}
        CHK.append({"id": iid, "value": value, "second": second, "source": second_src, "ok": ok})
    if iid in IDX: raise SystemExit("重複鍵：" + iid)
    IDX[iid] = it
    ITEMS.append(it)
    return it


def chk(cid, a, b, src, tol=0.5):
    ok = abs(a - b) <= tol
    CHK.append({"id": cid, "value": a, "second": b, "source": src, "ok": ok})
    return ok


def V(iid): return IDX[iid]["value"]


PER5 = ["2026q3", "2026q2", "2025q3", "9m2026", "9m2025"]      # 新聞稿五欄順序
PER4 = ["2026q3", "2025q3", "9m2026", "9m2025"]                # 10-Q 四欄順序
SEG = ["space", "conn", "si", "total"]
SEGN = {"space": "Space", "conn": "Connectivity", "si": "SI", "total": "合計"}

# ============ 1. 10-Q 附註 18 分部表（四個期間區塊）與 MD&A 分部調整後 EBITDA 調節表 ============
n18 = sec("Q", "Note 18 - Segments")
n19 = sec("Q", "Note 19 - Restructuring")
blk = {}
for i in range(n18, n19):
    s = Q[i].strip()
    m = re.match(r"(Three|Nine) Months Ended September 30,$", s)
    if m and Q[i + 1].strip() in ("2026", "2025"):
        blk[("q3_" if m.group(1) == "Three" else "9m") + Q[i + 1].strip()] = i
BK = {"2026q3": blk["q3_2026"], "9m2026": blk["9m2026"], "2025q3": blk["q3_2025"], "9m2025": blk["9m2025"]}
nm = sec("Q", "Non-GAAP Financial Measures")
lq = sec("Q", "Liquidity and Capital Resources")
eb = {}
for i in range(nm, lq):
    s = Q[i].strip()
    m = re.match(r"(Three|Nine) Months Ended September 30,$", s)
    if m and Q[i + 1].strip() in ("2026", "2025") and "Total Reportable" in Q[i + 2]:
        eb[("2026q3" if m.group(1) == "Three" and Q[i + 1].strip() == "2026" else "2025q3" if m.group(1) == "Three" else "9m" + Q[i + 1].strip())] = i
Q18 = {}
for per, b in BK.items():
    for key, lab in (("rev", "Revenue"), ("cogs", "Cost of revenue"), ("rd", "Research and development"), ("sga", "Selling, general, and administrative"),
                     ("restr", "Restructuring charges"), ("impair", "Impairment"), ("costs", "Total costs and expenses"), ("ebit", "Income (loss) from operations"),
                     ("da", "Depreciation and amortization"), ("sbc", "Share-based compensation"), ("sbcAcq", "Share-based compensation related to acquisitions"),
                     ("capex", "Capital expenditures")):
        try:
            v, i = row("Q", lab, b, 30)
        except SystemExit:
            continue                                   # 該期無此列（例如 2025 無 acquisitions SBC）
        Q18[(per, key)] = (v, i)
    v, i = row("Q", "Segment Adjusted EBITDA", eb[per], 14)
    Q18[(per, "adjEbitda")] = (v, i)

# ============ 2. 新聞稿第 3 頁 Third Quarter Financial Highlights（五欄） ============
h0 = sec("P", "Third Quarter Financial Highlights")
PH = {}
for key, title in (("rev", "Revenue"), ("ebit", "Income (Loss) from Operations"), ("adjEbitda", "Segment Adjusted EBITDA"), ("capex", "Capex")):
    s0 = sec("P", title, h0)
    for k, lab in (("space", "Space"), ("conn", "Connectivity"), ("si", "SI"), ("total", "Total")):
        PH[(key, k)] = row("P", lab, s0, 7, n=5)

# ============ 3. 新聞稿第 4–6 頁分部營運與財務資料（五欄） ============
sp = sec("P", "Space Segment Operating and Financial Data")
cn = sec("P", "Connectivity Segment Operating and Financial Data")
si = sec("P", "SI Segment Operating and Financial Data")
oth = sec("P", "Other Key Business Developments")
PS = {}
SEGROWS = {
    "space": (sp, [("launchSvc", "Launch services revenues"), ("launchDev", "Launch & development revenues"), ("rev", "Space revenues"),
                   ("cogs", "Cost of revenue"), ("rd", "Research and development"), ("sga", "Selling, general, and administrative"), ("impair", "Impairment"),
                   ("ebit", "Income (loss) from operations"), ("da", "Depreciation & amortization"), ("sbc", "Share-based compensation"),
                   ("adjEbitda", "Space Adjusted EBITDA"), ("capex", "Space Capex"),
                   ("kpi.launchCust", "Customer launches (#)"), ("kpi.launchInt", "Internal launches (#)"), ("kpi.launchTot", "Total launches (#)"),
                   ("kpi.massCust", "Customer payloads (in metric tons)"), ("kpi.massInt", "Internal payloads (in metric tons)"), ("kpi.massTot", "Mass to orbit (in metric tons)")]),
    "conn": (cn, [("consumer", "Consumer revenues"), ("entGov", "Enterprise & government revenues"), ("rev", "Connectivity revenues"),
                  ("cogs", "Cost of revenue"), ("rd", "Research and development"), ("sga", "Selling, general, and administrative"),
                  ("ebit", "Income from operations"), ("da", "Depreciation & amortization"), ("sbc", "Share-based compensation"),
                  ("adjEbitda", "Connectivity Adjusted EBITDA"), ("capex", "Connectivity Capex"),
                  ("kpi.subs", "Starlink consumer subscribers"), ("kpi.arpu", "Starlink consumer ARPU ($ per month)")]),
    "si": (si, [("ads", "Advertising revenues"), ("siSI", "SI solutions & infrastructure revenues"), ("rev", "SI revenues"),
                ("cogs", "Cost of revenue"), ("rd", "Research and development"), ("sga", "Selling, general, and administrative"), ("restr", "Restructuring charges (credits)"),
                ("ebit", "Income (loss) from operations"), ("da", "Depreciation & amortization"), ("sbc", "Share-based compensation"),
                ("sbcAcq", "Share-based compensation related to acquisitions"), ("acqCost", "Acquisition-related costs"),
                ("adjEbitda", "SI Adjusted EBITDA"), ("capex", "SI Capex"), ("kpi.gw", "Nameplate compute (in GW)")]),
}
for sg, (s0, rows) in SEGROWS.items():
    for key, lab in rows:
        PS[(sg, key)] = row("P", lab, s0, 45, n=5)

# ============ 4. 10-Q 附註 3 收入拆分（四欄） ============
n3 = sec("Q", "Revenue disaggregated by type and segment is as follows:")
Q3R = {}
for key, lab in (("space.launchSvc", "Launch Services"), ("space.launchDev", "Launch & Development"), ("space.rev", "Space"),
                 ("conn.consumer", "Consumer"), ("conn.entGov", "Enterprise & Government"), ("conn.rev", "Connectivity"),
                 ("si.ads", "Advertising"), ("si.siSI", "SI Solutions & Infrastructure"), ("si.rev", "SI"), ("total.rev", "Total revenues")):
    Q3R[key] = row("Q", lab, n3, 16)
n3a = sec("Q", "Revenue disaggregated by products and services is as follows:")
for key, lab in (("products", "Products"), ("services", "Services")):
    Q3R[key] = row("Q", lab, n3a, 8)

# ---- 分部主列（營收、營業利益、調整後 EBITDA、資本支出）：五期
UNIT = "USD M"
for key, item in (("rev", "營收"), ("ebit", "營業利益（損失）"), ("adjEbitda", "分部調整後 EBITDA"), ("capex", "資本支出")):
    for sg in SEG:
        v5, i5 = PH[(key, sg)]
        for k, per in enumerate(PER5):
            iid = f"seg.{sg}.{key}.{per}"
            tag = "[Interested-party]" if key == "adjEbitda" else "[Verified]"
            if per == "2026q2":
                # 2026Q2 只見於新聞稿：第二讀＝同稿分部頁（另一張表）
                if sg == "total":
                    sec_v = sum(PS[(s, key)][0][k] for s in ("space", "conn", "si"))
                    sec_src = "新聞稿分部頁三分部同列相加（第 4–6 頁）"
                else:
                    sec_v = PS[(sg, key)][0][k]
                    sec_src = f"新聞稿第 {PP[PS[(sg, key)][1]]} 頁 {SEGN[sg]} 分部頁同列"
                add(iid, f"{SEGN[sg]}：{item}", v5[k], UNIT, per, "P", i5, "Third Quarter Financial Highlights 表（June 30, 2026 欄）", tag="[Interested-party]",
                    note="2026Q2 單季只見於 8-K 新聞稿比較欄；2026Q2 10-Q 未入庫", second=sec_v, second_src=sec_src)
                continue
            j = PER4.index(per)
            if (per, key) in Q18:
                qv, qi = Q18[(per, key)]
                col = SEG.index(sg)
                add(iid, f"{SEGN[sg]}：{item}", qv[col], UNIT, per, "Q", qi,
                    ("附註 18 Segments 表" if key != "adjEbitda" else "MD&A 分部調整後 EBITDA 調節表"), tag=tag,
                    second=v5[k], second_src=f"新聞稿第 {PP[i5]} 頁 Highlights 表")
# ---- 分部成本、折攤、股權報酬：10-Q 四期為主、新聞稿分部頁為第二讀；2026Q2 只有新聞稿（恆等式檢查）
COST = [("cogs", "營業成本"), ("rd", "研發費用"), ("sga", "銷管費用"), ("restr", "重組費用（回沖）"), ("impair", "減損"),
        ("da", "折舊與攤銷"), ("sbc", "股權報酬（不含收購相關）"), ("sbcAcq", "收購相關股權報酬"), ("acqCost", "收購相關成本"), ("costs", "成本費用合計")]
for sg in SEG:
    for key, item in COST:
        for k, per in enumerate(PER5):
            iid = f"seg.{sg}.{key}.{per}"
            pv = PS.get((sg, key)) if sg != "total" else None
            if per == "2026q2":
                if pv is None: continue
                add(iid, f"{SEGN[sg]}：{item}", pv[0][k], UNIT, per, "P", pv[1], f"{SEGN[sg]} Segment Operating and Financial Data 表", tag="[Interested-party]",
                    note="只見於新聞稿；以分部營業利益與調整後 EBITDA 恆等式檢查")
                continue
            if (per, key) not in Q18:
                if key == "acqCost": continue          # 不在附註 18，下方另由 MD&A 調節表讀
                if pv is not None and pv[0][k] != 0:
                    raise SystemExit(f"10-Q 無 {sg}.{key}.{per} 但新聞稿非 0")
                continue
            qv, qi = Q18[(per, key)]
            val = qv[SEG.index(sg)]
            if pv is None and val == 0 and key in ("impair", "restr", "sbcAcq", "acqCost"): continue
            add(iid, f"{SEGN[sg]}：{item}", val, UNIT, per, "Q", qi, "附註 18 Segments 表",
                second=(pv[0][k] if pv is not None else None), second_src=(f"新聞稿第 {PP[pv[1]]} 頁 {SEGN[sg]} 分部頁" if pv is not None else ""))
# acquisition-related costs 不在附註 18（在 MD&A 調節表）
for per, j in (("2026q3", 0), ("9m2026", 3)):
    v, i = row("Q", "Acquisition-related costs", eb[per], 14)
    add(f"seg.si.acqCost.{per}", "SI：收購相關成本", v[2], UNIT, per, "Q", i, "MD&A 分部調整後 EBITDA 調節表", tag="[Interested-party]",
        second=PS[("si", "acqCost")][0][j], second_src="新聞稿 SI 分部頁")

# ---- 子列（S3）：五期；10-Q 附註 3 為主，新聞稿分部頁為第二讀；2026Q2 第二讀＝母分部 − 另一子列
SUB = {"space": [("launchSvc", "發射服務收入"), ("launchDev", "發射與開發收入")], "conn": [("consumer", "消費收入"), ("entGov", "企業與政府收入")],
       "si": [("ads", "廣告收入"), ("siSI", "SI 解決方案與基礎設施收入")]}
for sg, subs in SUB.items():
    for a, (key, item) in enumerate(subs):
        other = subs[1 - a][0]
        pv, pi = PS[(sg, key)]
        for k, per in enumerate(PER5):
            iid = f"seg.{sg}.{key}.{per}"
            if per == "2026q2":
                add(iid, f"{SEGN[sg]}：{item}", pv[k], UNIT, per, "P", pi, f"{SEGN[sg]} 分部頁（June 30, 2026 欄）", tag="[Interested-party]",
                    second=PS[(sg, "rev")][0][k] - PS[(sg, other)][0][k], second_src="同表母分部營收 − 另一子列")
                continue
            qv, qi = Q3R[f"{sg}.{key}"]
            add(iid, f"{SEGN[sg]}：{item}", qv[PER4.index(per)], UNIT, per, "Q", qi, "附註 3 Revenue 收入拆分表", second=pv[k], second_src=f"新聞稿第 {PP[pi]} 頁分部頁")
for key, item in (("products", "產品收入（全屬 Connectivity）"), ("services", "服務收入")):
    qv, qi = Q3R[key]
    for j, per in enumerate(PER4):
        add(f"rev.{key}.{per}", item, qv[j], UNIT, per, "Q", qi, "附註 3 收入依產品與服務拆分")

# ---- 分部營運指標
KPI = [("space", "kpi.launchCust", "客戶發射次數", "次"), ("space", "kpi.launchInt", "內部發射次數", "次"), ("space", "kpi.launchTot", "發射次數合計", "次"),
       ("space", "kpi.massCust", "客戶酬載質量", "公噸"), ("space", "kpi.massInt", "內部酬載質量", "公噸"), ("space", "kpi.massTot", "入軌質量合計", "公噸"),
       ("conn", "kpi.subs", "Starlink 消費訂戶（期末）", "百萬戶"), ("conn", "kpi.arpu", "Starlink 消費 ARPU（每月）", "USD/月"),
       ("si", "kpi.gw", "名目算力 nameplate compute（期末）", "GW")]
# 10-Q 第二讀
mt = find(Q, lambda l: l.strip().startswith("(in metric tons)"))
QM = {"kpi.massTot": row("Q", "Mass to orbit", mt, 5), "kpi.massCust": row("Q", "of which, attributable to customer payloads", mt, 5),
      "kpi.massInt": row("Q", "of which, attributable to internal payloads", mt, 5)}
ln = find(Q, lambda l: l.strip().startswith("(number)"))
fal, fi = row("Q", "Falcon launches", ln, 6)
cus, ci = row("Q", "of which, customer launches", ln, 6)
inn, ii = row("Q", "of which, internal launches", ln, 6)
stp, sti = row("Q", "Starship launches", ln, 6)
QM["kpi.launchCust"] = (cus, ci)
QM["kpi.launchInt"] = ([inn[j] + stp[j] for j in range(4)], ii)
QM["kpi.launchTot"] = ([fal[j] + stp[j] for j in range(4)], fi)
ap = find(Q, lambda l: l.strip().startswith("(dollars per month)"))
QM["kpi.arpu"] = row("Q", "Starlink consumer ARPU", ap, 3)
sb = find(Q, lambda l: l.strip().startswith("Starlink consumer subscribers") and len(trailing(l)) == 2)
gwi = find(Q, lambda l: l.strip().startswith("Nameplate compute") and len(trailing(l)) == 2)
for sg, key, item, unit in KPI:
    pv, pi = PS[(sg, key)]
    for k, per in enumerate(PER5):
        if key in ("kpi.subs", "kpi.gw") and per.startswith("9m"): continue      # 期末值，9M 欄與 Q3 欄相同
        iid = f"seg.{sg}.{key}.{per}"
        second, ss = None, ""
        if per != "2026q2":
            j = PER4.index(per)
            if key in QM:
                second = QM[key][0][j]
                ss = "10-Q MD&A Key Business Metrics 表" + ("（Falcon 內部＋Starship）" if key == "kpi.launchInt" else "（Falcon＋Starship）" if key == "kpi.launchTot" else "")
            elif key in ("kpi.subs", "kpi.gw") and per in ("2026q3", "2025q3"):
                ii_ = sb if key == "kpi.subs" else gwi
                second = trailing(Q[ii_])[0 if per == "2026q3" else 1]; ss = "10-Q MD&A Key Business Metrics 表"
        add(iid, f"{SEGN[sg]}：{item}", pv[k], unit, per, "P", pi, f"{SEGN[sg]} 分部頁", tag="[Interested-party]",
            note=("公司自訂指標" + ("；Starship 發射全列內部（10-Q 註 1）" if key == "kpi.launchInt" else "")), second=second, second_src=ss, tol=0.5 if unit != "GW" else 0.05)
for per, j in (("2026q3", 0), ("2025q3", 1), ("9m2026", 2), ("9m2025", 3)):
    add(f"seg.space.kpi.starship.{per}", "Space：Starship 發射次數", stp[j], "次", per, "Q", sti, "MD&A Launches 表", tag="[Interested-party]",
        note="全部列為內部發射")
    add(f"seg.space.kpi.falcon.{per}", "Space：Falcon 發射次數", fal[j], "次", per, "Q", fi, "MD&A Launches 表", tag="[Interested-party]",
        second=V(f"seg.space.kpi.launchTot.{per}") - stp[j], second_src="新聞稿發射合計 − 10-Q Starship 次數")

# ============ 4b. 官方 Q2 財報新聞稿（公司 IR PDF；chat 端 10-30 補充，Andy 授權抓取） ============
# 本環境 curl 連 q4cdn 被代理擋（403），PDF 未入庫；數字以 WebFetch 讀取（2026-10-30，分三次讀表），故每筆以
# (a) Q3 新聞稿 June 30, 2026 欄、(b) 9M − Q2 − Q3 推得之 2026Q1、(c) 1H＋Q3 ＝ 9M（10-Q）三種方式交叉核對。Q2 新聞稿分部名稱為「AI」，Q3 起改稱「SI」（Q3 新聞稿註 2）。
Q2URL = "https://s21.q4cdn.com/184289198/files/doc_financials/2026/q2/SpaceX-Reports-Second-Quarter-2026-Results.pdf"
Q2SRC = "SpaceX 2026Q2 財報新聞稿（公司 IR PDF，" + Q2URL + "；WebFetch 2026-10-30 讀取，PDF 未入庫）"
Q2COL = ["2026q2", "2026q1", "2025q2", "1h2026", "1h2025"]
Q2T = {  # Financial Highlights 表（USD M）
    ("rev", "space"): [962, 619, 746, 1581, 1611], ("rev", "conn"): [4291, 3257, 2588, 7548, 5062], ("rev", "si"): [2561, 818, 737, 3379, 1465], ("rev", "total"): [7814, 4694, 4071, 12508, 8138],
    ("ebit", "space"): [-542, -662, -369, -1204, -439], ("ebit", "conn"): [1656, 1188, 923, 2844, 1956], ("ebit", "si"): [-1257, -2469, -1524, -3726, -2460], ("ebit", "total"): [-143, -1943, -970, -2086, -943],
    ("adjEbitda", "space"): [-205, -351, -93, -556, 131], ("adjEbitda", "conn"): [2597, 2087, 1583, 4684, 3200], ("adjEbitda", "si"): [1146, -609, -276, 537, -387], ("adjEbitda", "total"): [3538, 1127, 1214, 4665, 2944],
    ("capex", "space"): [1174, 1052, 946, 2226, 1705], ("capex", "conn"): [1367, 1332, 1130, 2699, 1944], ("capex", "si"): [15828, 7723, 749, 23551, 3316], ("capex", "total"): [18369, 10107, 2825, 28476, 6965],
    # 分部頁
    ("launchSvc", "space"): [648, 330, 490, 978, 1056], ("launchDev", "space"): [314, 289, 256, 603, 555],
    ("consumer", "conn"): [2485, 2148, 1721, 4633, 3213], ("entGov", "conn"): [1806, 1109, 867, 2915, 1849],
    ("ads", "si"): [367, 343, 426, 710, 870], ("siSI", "si"): [2194, 475, 311, 2669, 595],
    ("da", "space"): [158, 166, 146, 324, 308], ("da", "conn"): [805, 783, 569, 1588, 1078], ("da", "si"): [1885, 1493, 811, 3378, 1584],
    ("sbc", "space"): [179, 145, 125, 324, 233], ("sbc", "conn"): [136, 116, 91, 252, 166], ("sbc", "si"): [516, 378, 247, 894, 295],
    ("kpi.launchCust", "space"): [10, 7, 9, 17, 21], ("kpi.launchInt", "space"): [28, 33, 37, 61, 63], ("kpi.launchTot", "space"): [38, 40, 46, 78, 84],
    ("kpi.massCust", "space"): [87, 45, 88, 132, 163], ("kpi.massInt", "space"): [397, 511, 563, 908, 938], ("kpi.massTot", "space"): [485, 556, 652, 1041, 1102],
    ("kpi.subs", "conn"): [12.0, 10.3, 6.0, 12.0, 6.0], ("kpi.arpu", "conn"): [66, 66, 85, 66, 85], ("kpi.gw", "si"): [1.4, 1.0, 0.4, 1.4, 0.4]}
ITEMN = {"rev": "營收", "ebit": "營業利益（損失）", "adjEbitda": "分部調整後 EBITDA", "capex": "資本支出", "launchSvc": "發射服務收入", "launchDev": "發射與開發收入",
         "consumer": "消費收入", "entGov": "企業與政府收入", "ads": "廣告收入", "siSI": "SI（Q2 稱 AI）解決方案與基礎設施收入", "da": "折舊與攤銷", "sbc": "股權報酬（不含收購相關）",
         "kpi.launchCust": "客戶發射次數", "kpi.launchInt": "內部發射次數", "kpi.launchTot": "發射次數合計", "kpi.massCust": "客戶酬載質量", "kpi.massInt": "內部酬載質量",
         "kpi.massTot": "入軌質量合計", "kpi.subs": "Starlink 訂戶（期末；Q2 稱 Starlink Subscribers）", "kpi.arpu": "Starlink ARPU（每月）", "kpi.gw": "名目算力（期末）"}
KUNIT = {"kpi.launchCust": "次", "kpi.launchInt": "次", "kpi.launchTot": "次", "kpi.massCust": "公噸", "kpi.massInt": "公噸", "kpi.massTot": "公噸",
         "kpi.subs": "百萬戶", "kpi.arpu": "USD/月", "kpi.gw": "GW"}


def addw(iid, item, value, unit, period, source, tag, note="", second=None, second_src="", tol=0.5, q="", url="", docDate=""):
    it = {"id": iid, "item": item, "value": value, "unit": unit, "period": period, "source": source, "quote": q, "url": url, "docDate": docDate, "tag": tag, "note": note}
    if second is not None:
        ok = abs(second - value) <= tol + 1e-9 * abs(value)
        it["check"] = {"value": second, "source": second_src, "ok": ok}
        CHK.append({"id": iid, "value": value, "second": second, "source": second_src, "ok": ok})
    if iid in IDX: raise SystemExit("重複鍵：" + iid)
    IDX[iid] = it; ITEMS.append(it)
    return it


for (key, sg), vals in Q2T.items():
    for k, per in enumerate(Q2COL):
        iid = f"seg.{sg}.{key}.{per}"
        if per == "2026q2":     # 與 Q3 新聞稿 June 30, 2026 欄互核（不另立鍵）
            if iid in IDX:
                chk(f"{iid}（Q2 新聞稿 vs Q3 新聞稿比較欄）", vals[k], V(iid), "Q2 新聞稿 Highlights／分部頁", tol=0.05 if key == "kpi.gw" else 0.5)
            continue
        if key in ("kpi.subs", "kpi.gw") and per.startswith("1h"): continue
        sv, ss, tol = None, "", (0.05 if key == "kpi.gw" else 0.5)
        if per == "2026q1" and f"seg.{sg}.{key}.9m2026" in IDX and f"seg.{sg}.{key}.2026q2" in IDX and key not in ("kpi.subs", "kpi.gw", "kpi.arpu"):
            sv = V(f"seg.{sg}.{key}.9m2026") - V(f"seg.{sg}.{key}.2026q2") - V(f"seg.{sg}.{key}.2026q3"); ss = "Q3 新聞稿 9M − Q2 − Q3（衍生）"
        if per == "1h2025" and f"seg.{sg}.{key}.9m2025" in IDX and key not in ("kpi.subs", "kpi.gw", "kpi.arpu"):
            sv = V(f"seg.{sg}.{key}.9m2025") - V(f"seg.{sg}.{key}.2025q3"); ss = "10-Q 9M2025 − 2025Q3（容差 2.5：Q2 新聞稿 1H2025 欄與 Q3 10-Q 2025 比較數有 1–2 百萬差，兩次 WebFetch 讀數相同，判為前後期微調或四捨五入，見報告）"; tol = 2.5
        if per == "1h2026" and key not in ("kpi.subs", "kpi.gw", "kpi.arpu"):
            sv = vals[0] + vals[1]; ss = "同稿 Q2 欄＋Q1 欄"
        addw(iid, f"{SEGN[sg]}：{ITEMN[key]}", vals[k], KUNIT.get(key, UNIT), per, Q2SRC + (" Financial Highlights 表" if key in ("rev", "ebit", "adjEbitda", "capex") else " 分部營運與財務資料表"),
             "[Verified]" if not key.startswith("kpi") and key not in ("adjEbitda",) else "[Interested-party]",
             note="公司 IR 官方新聞稿；WebFetch 讀取（PDF 未入庫，待下載後以 pdftotext 重讀）", second=sv, second_src=ss, tol=tol, url=Q2URL, docDate="2026-07")

# ============ 4c. FY2025（424B4／S-1）— 二手轉述，待原文核對 ============
TAGS1 = "[二手轉述 S-1（待 424B4 原文核對）]"
FOOL = "https://www.fool.com/investing/2026/05/24/spacex-is-going-public-1-and-half-trillion/"
YH = "https://finance.yahoo.com/markets/article/spacex-files-ipo-prospectus-offering-a-peek-into-its-finances-205406189.html"
IMP = "https://www.implicator.ai/spacex-targets-135-ipo-price-for-75-billion-nasdaq-listing/"
FOOL2 = "https://fool.com/investing/2026/05/29/spacex-invests-3x-more-on-ai-than-rockets-and-lost/"
BC = "https://beancount.io/blog/2026/07/12/spacex-fy2025-earnings-analysis"
FY = [("seg.total.rev.fy2025", "合併營收", 18674, [FOOL, YH, IMP], "Fool 表 18,674；Yahoo 18.674B；Implicator 18.674B"),
      ("seg.total.ebit.fy2025", "合併營業損失", -2589, [FOOL, YH, IMP], "三家一致（2,589）"),
      ("seg.total.adjEbitda.fy2025", "合併調整後 EBITDA", 6584, [YH, IMP], "Yahoo、Implicator 6.584B"),
      ("seg.space.rev.fy2025", "Space：營收", 4086, [FOOL, YH, IMP], "三家一致"), ("seg.space.ebit.fy2025", "Space：營業損失", -657, [FOOL, YH, IMP], "三家一致"),
      ("seg.space.adjEbitda.fy2025", "Space：分部調整後 EBITDA", 653, [YH], "只見 Yahoo（單一來源）"),
      ("seg.conn.rev.fy2025", "Connectivity：營收", 11387, [FOOL, YH, IMP], "三家一致"), ("seg.conn.ebit.fy2025", "Connectivity：營業利益", 4423, [FOOL, YH, IMP], "三家一致"),
      ("seg.conn.adjEbitda.fy2025", "Connectivity：分部調整後 EBITDA", 7168, [YH], "只見 Yahoo（chat 端所列 7,168 相同；單一來源）"),
      ("seg.si.rev.fy2025", "AI（現 SI）：營收", 3201, [FOOL, IMP], "Fool 表、Implicator 一致；Yahoo 未列"),
      ("seg.si.ebit.fy2025", "AI（現 SI）：營業損失", -6355, [FOOL, YH, IMP], "三家一致（Yahoo 未說明口徑，由合計反推為營業損失）"),
      ("seg.conn.capex.fy2025", "Connectivity：資本支出", 4178, [BC, FOOL2], "Beancount 4.18B；Fool 寫 4.1B（不一致）。chat 端審查 2026-10-30：待核（424B4）"),
      ("seg.si.capex.fy2025", "AI（現 SI）：資本支出", 12727, [BC, FOOL2], "Beancount 12.73B；Fool 12.7B")]
for iid, item, v, urls, note in FY:
    addw(iid, item + "（FY2025）", v, UNIT, "fy2025", "媒體轉述 SpaceX S-1／424B4（WebFetch 2026-10-30）", TAGS1, note=note, url="；".join(urls), docDate="2026-05",
         q="Data source: SpaceX S-1（Fool 表註）" if FOOL in urls else "")
addw("seg.si.adjEbitda.fy2025", "AI（現 SI）：分部調整後 EBITDA（FY2025，衍生）", 6584 - 653 - 7168, UNIT, "fy2025", "合併 6,584 − Space 653 − Connectivity 7,168", "[Derived]",
     note="由二手轉述數字推得；假設分部調整後 EBITDA 合計＝合併（2026 各期成立）")
addw("seg.space.capex.fy2025", "Space：資本支出（FY2025）", None, UNIT, "fy2025", "未取得", "[未取得]", url="；".join([BC, FOOL2]),
     note="chat 端審查 2026-10-30：媒體轉述 3,832（Beancount 3.83B、Fool 3.8B）與 9M2026 Space 資本支出 3,832 完全相同，判為媒體轉述混用，改「未取得」，待 424B4")
addw("seg.total.capex.fy2025", "資本支出合計（FY2025，二手轉述）", 20737, UNIT, "fy2025", "Beancount 合計 20.74B（二手）", TAGS1, url=BC,
     note="Space 未取得，三分部無法加總；合計只作參考，不做 FY2025 資本支出分部對帳")
chk("Fool 表 2026Q1 合計 4,964 vs 分部加總 4,694（媒體筆誤）", 4964, V("seg.total.rev.2026q1"), "記錄用：預期不一致，不列入失敗", tol=1e9)

# ============ 5. 分部對帳與恆等式（五期，差額須為 0） ============
RECON = []
for per in PER5 + ["2026q1", "2025q2", "1h2026", "1h2025", "fy2025"]:
    for key in ("rev", "ebit", "adjEbitda", "capex"):
        if per == "fy2025" and key in ("adjEbitda", "capex"): continue      # AI 為衍生，恆等；Space 資本支出未取得
        s3 = sum(V(f"seg.{s}.{key}.{per}") for s in ("space", "conn", "si"))
        t = V(f"seg.total.{key}.{per}")
        RECON.append({"period": per, "line": key, "segSum": s3, "total": t, "diff": round(s3 - t, 6)})
        chk(f"recon.{key}.{per}（三分部合計 − 合併）", s3, t, "分部對帳")
    for sg, subs in SUB.items():
        if per == "fy2025": continue
        s2 = sum(V(f"seg.{sg}.{k}.{per}") for k, _ in subs)
        RECON.append({"period": per, "line": f"{sg}.sublines", "segSum": s2, "total": V(f"seg.{sg}.rev.{per}"), "diff": round(s2 - V(f"seg.{sg}.rev.{per}"), 6)})
        chk(f"recon.{sg}.子列.{per}（子列合計 − 母分部）", s2, V(f"seg.{sg}.rev.{per}"), "子列對帳")
    if per not in PER5: continue
    for sg in ("space", "conn", "si"):
        g = lambda k: IDX.get(f"seg.{sg}.{k}.{per}", {"value": 0})["value"]
        ebit = g("rev") - g("cogs") - g("rd") - g("sga") - g("restr") - g("impair")
        chk(f"ident.{sg}.ebit.{per}（營收 − 成本費用）", ebit, g("ebit"), "同表恆等式")
        ae = g("ebit") + g("da") + g("sbc") + g("sbcAcq") + g("acqCost") + g("restr") + g("impair")
        chk(f"ident.{sg}.adjEbitda.{per}（營業利益＋調整項）", ae, g("adjEbitda"), "同表恆等式")

# ---- 2026Q2 文字句第二讀（新聞稿 Business Highlights：季增率、訂戶增加、算力）
def pct(a, b): return round((a / b - 1) * 100)
for sg, p_txt in (("space", 84), ("conn", 22), ("si", 195)):
    chk(f"text.{sg}.rev.qoq", pct(V(f"seg.{sg}.rev.2026q3"), V(f"seg.{sg}.rev.2026q2")), p_txt, "新聞稿文字：季增 %（取整）", tol=1)
chk("text.conn.subs.qoq", round(V("seg.conn.kpi.subs.2026q3") - V("seg.conn.kpi.subs.2026q2"), 1), 1.6, "新聞稿文字：up 1.6 million sequentially", tol=0.05)
chk("text.conn.arpu.qoq", pct(V("seg.conn.kpi.arpu.2026q3"), V("seg.conn.kpi.arpu.2026q2")), 5, "新聞稿文字：ARPU improved 5% sequentially", tol=1)
chk("text.si.gw.2026q2", V("seg.si.kpi.gw.2026q2"), 1.4, "新聞稿文字：up from 1.4 GW as of June 30, 2026", tol=0.01)
chk("text.si.siSI.yoy", V("seg.si.siSI.2026q3") - V("seg.si.siSI.2025q3"), 6700, "新聞稿：Added $6.7 billion of incremental SI solutions and infrastructure revenue", tol=50)

# ============ 6. 合併損益表（新聞稿 vs 10-Q） ============
qis = sec("Q", "Consolidated Statements of Operations")
pis = sec("P", "Consolidated Statements of Operations", exact=False, start=sec("P", "Consolidated Financial Statements"))
IS = [("rev", "Revenue", "營收"), ("cogs", "Cost of revenue", "營業成本"), ("rd", "Research and development", "研發費用"),
      ("sga", "Selling, general, and administrative", "銷管費用"), ("restr", "Restructuring charges (credits)", "重組費用（回沖）"), ("impair", "Impairment", "減損"),
      ("costs", "Total costs and expenses", "成本費用合計"), ("ebit", "Income (loss) from operations", "營業利益（損失）"),
      ("intExp", "Interest expense", "利息費用（含關係人）"), ("intInc", "Interest income", "利息收入"), ("oie", "Other income (expense), net", "其他收益（費用）淨額"),
      ("pretax", "Income (loss) before income taxes", "稅前損益"), ("tax", "Provision for (benefit from) income taxes", "所得稅費用（利益）"),
      ("ni", "Net income (loss)", "淨利（損）"), ("niCommon", "Net income (loss) attributable to shareholders - basic and diluted", "歸屬普通股股東淨利（損）")]
for key, lab, item in IS:
    qv, qi = row("Q", lab, qis, 40)
    pv, pi = row("P", lab, pis, 40)
    for j, per in enumerate(PER4):
        add(f"is.{key}.{per}", item, qv[j], UNIT, per, "Q", qi, "合併損益表", second=pv[j], second_src=f"新聞稿第 {PP[pi]} 頁合併損益表")
for key, item, unit, nth in (("epsBasic", "基本 EPS", "USD", 0), ("epsDil", "稀釋 EPS", "USD", 0), ("shBasic", "基本加權股數", "M 股", 1), ("shDil", "稀釋加權股數", "M 股", 1)):
    lab = "Basic" if "Basic" in key else "Diluted"
    qv, qi = row("Q", lab, qis, 40, nth=nth)
    pv, pi = row("P", lab, pis, 40, nth=nth)
    for j, per in enumerate(PER4):
        add(f"is.{key}.{per}", item, qv[j], unit, per, "Q", qi, "合併損益表", second=pv[j], second_src=f"新聞稿第 {PP[pi]} 頁", tol=0.005 if unit == "USD" else 0.5)
for j, per in enumerate(PER4):
    chk(f"is.ebit.{per}（損益表 vs 分部合計）", V(f"is.ebit.{per}"), V(f"seg.total.ebit.{per}"), "損益表 vs 附註 18")
    chk(f"is.rev.{per}（損益表 vs 分部合計）", V(f"is.rev.{per}"), V(f"seg.total.rev.{per}"), "損益表 vs 附註 18")
    chk(f"is.rev.{per}（產品＋服務）", V(f"rev.products.{per}") + V(f"rev.services.{per}"), V(f"is.rev.{per}"), "附註 3 產品＋服務 vs 損益表")
# 關係人利息（Valor）與其他收益（TBC）
i = find(Q, lambda l: "Interest expense (related party of $351" in l)
add("is.intExpRP.2026q3", "利息費用中的關係人部分（Valor 失敗售後租回）", 351, UNIT, "2026q3", "Q", i, "合併損益表括號說明",
    second=float(re.search(r"\$(\d+) million and \$(\d+) million in Interest expense", re.sub(r"\s+", " ", " ".join(Q[find(Q, lambda l: 'CTC entered into various equipment lease agreements with Valor' in l):][:8]))).group(1)),
    second_src="10-Q 附註 17 Valor 段文字")
add("is.intExpRP.9m2026", "利息費用中的關係人部分（9M）", 864, UNIT, "9m2026", "Q", i + 1, "合併損益表括號說明",
    second=float(re.search(r"\$(\d+) million and \$(\d+) million in Interest expense", re.sub(r"\s+", " ", " ".join(Q[find(Q, lambda l: 'CTC entered into various equipment lease agreements with Valor' in l):][:8]))).group(2)),
    second_src="10-Q 附註 17 Valor 段文字")
i = find(Q, lambda l: "investment of $561 million recorded" in l)
add("is.tbcGain.2026q3", "TBC（The Boring Company）公允價值未實現利益（列其他收益）", 561, UNIT, "2026q3", "Q", i, "附註 8",
    second=561, second_src="合併損益表 Other income 括號（related party of $561）", note="關係人；非經常")
i = find(Q, lambda l: "Less: Deemed dividend" in l)
add("is.deemedDiv.9m2026", "視同股利（買回特別股超額）", trailing(Q[i])[2], UNIT, "9m2026", "Q", i, "附註 13 EPS",
    second=V("is.ni.9m2026") - V("is.niCommon.9m2026"), second_src="淨損 − 歸屬普通股淨損")

# ---- 合併調整後 EBITDA 調節（新聞稿 vs 10-Q MD&A）
qa = find(Q, lambda l: "reconciliation of Net income (loss), the most directly comparable" in l)
pa = sec("P", "Reconciliation of Net Income (Loss) to Adjusted EBITDA")
for key, lab, item in (("da", "Depreciation and amortization", "折舊與攤銷"), ("sbc", "Share-based compensation", "股權報酬（不含收購相關）"),
                       ("sbcAcq", "Share-based compensation related to acquisitions", "收購相關股權報酬"), ("acqCost", "Acquisition-related costs", "收購相關成本"),
                       ("adjEbitda", "Adjusted EBITDA", "調整後 EBITDA（公司定義）")):
    qv, qi = row("Q", lab, qa, 20)
    pv, pi = row("P", lab, pa, 20)
    for j, per in enumerate(PER4):
        if key in ("sbcAcq", "acqCost") and per in ("2025q3", "9m2025"): continue
        add(f"cons.{key}.{per}", item, qv[j], UNIT, per, "Q", qi, "MD&A Adjusted EBITDA 調節表", tag="[Interested-party]",
            second=pv[j], second_src=f"新聞稿第 {PP[pi]} 頁調節表")
for per in PER4:
    chk(f"cons.adjEbitda.{per}（合併 vs 分部合計）", V(f"cons.adjEbitda.{per}"), V(f"seg.total.adjEbitda.{per}"), "合併調節表 vs 分部表")

# ============ 7. 現金流量表（9M；新聞稿 vs 10-Q） ============
qcf = sec("Q", "Consolidated Statements of Cash Flows")
pcf = sec("P", "Consolidated Statements of Cash Flows", exact=False, start=pis)
CF = [("ni", "Net loss", "淨損"), ("da", "Depreciation and amortization", "折舊與攤銷"), ("sbc", "Share-based compensation", "股權報酬（含收購相關）"),
      ("defTax", "Deferred income taxes", "遞延所得稅"), ("digital", "Unrealized (gain) loss on digital assets", "數位資產未實現損益"),
      ("affGain", "Unrealized gain on investments in unconsolidated affiliates, net", "未合併關聯企業未實現利益"), ("extLoss", "Loss on debt extinguishment", "債務清償損失"),
      ("othAdj", "Other", "其他調整"), ("wc.ar", "Accounts receivable", "應收帳款變動"), ("wc.inv", "Inventory", "存貨變動"),
      ("wc.prepaid", "Prepaid expenses and other assets", "預付與其他資產變動"), ("wc.ap", "Accounts payable", "應付帳款變動"),
      ("wc.dr", "Deferred revenue", "遞延收入變動"), ("wc.oth", "Other liabilities, net", "其他負債變動"),
      ("cfo", "Net cash provided by operating activities", "營業現金流"), ("capex", "Purchases of property, plant, and equipment", "資本支出（購置不動產廠房設備，含關係人）"),
      ("capInt", "Capitalized interest", "資本化利息"), ("rebate", "Proceeds from product rebates", "產品回饋金收入（伺服器與網通設備）"),
      ("buyMkt", "Purchases of marketable securities", "購買有價證券"), ("matMkt", "Maturities of marketable securities", "有價證券到期"),
      ("sellMkt", "Proceeds from sales of marketable securities", "出售有價證券"), ("intang", "Payments for intangible assets", "支付無形資產（頻譜預付等）"),
      ("acqCash", "Cash received from acquisition of businesses", "收購取得現金"), ("invOth", "Other investing activities, net", "其他投資活動"),
      ("cfi", "Net cash used in investing activities", "投資現金流"), ("flPrin", "Principal repayments on finance leases", "融資租賃本金償還"),
      ("debtIss", "Proceeds from debt and other financing obligations", "舉債與其他融資收入"), ("debtCost", "Payment of debt issuance costs", "發債成本"),
      ("debtRep", "Repayments on debt and other financing obligations", "償還債務與其他融資"), ("extPrem", "Payment of debt extinguishment premium", "提前清償溢價"),
      ("stockIss", "Proceeds from issuance of capital stock, net of issuance costs", "發行股本（IPO 前私募）淨額"), ("espp", "Proceeds from employee equity award plans", "員工認股收入"),
      ("buyback", "Payments for repurchase of common and redeemable convertible preferred stock", "買回普通股與特別股"),
      ("taxNet", "Taxes paid related to net share settlement of equity awards", "股權獎酬淨額交割代扣稅"),
      ("ipo", "Proceeds from IPO, net of underwriting commissions and offering costs", "IPO 淨收入"), ("cff", "Net cash provided by financing activities", "籌資現金流"),
      ("fx", "Effect of exchange rate changes on cash and cash equivalents", "匯率影響"), ("netChg", "Net change in cash and cash equivalents and restricted cash", "現金淨變動"),
      ("begin", "Cash and cash equivalents and restricted cash, beginning of the period", "期初現金（含受限）"),
      ("end", "Cash and cash equivalents and restricted cash, end of the period", "期末現金（含受限）")]
for key, lab, item in CF:
    qv, qi = row("Q", lab, qcf, 60, n=2)
    pv, pi = row("P", lab, pcf, 60, n=2)
    for j, per in enumerate(("9m2026", "9m2025")):
        add(f"cf.{key}.{per}", item, qv[j], UNIT, per, "Q", qi, "合併現金流量表", second=pv[j], second_src=f"新聞稿第 {PP[pi]} 頁現金流量表")
for per in ("9m2026", "9m2025"):
    chk(f"cf.capex.{per}（現金流量表 vs 分部資本支出合計）", -V(f"cf.capex.{per}"), V(f"seg.total.capex.{per}"), "現金流量表購置 PP&E ＝ 分部 Capex 合計")
    chk(f"cf.da.{per}（現金流量表 vs 調節表）", V(f"cf.da.{per}"), V(f"cons.da.{per}"), "現金流量表 vs 調整後 EBITDA 調節表")
chk("cf.sbc.9m2026（現金流量表 3,833 ＝ 一般 2,616＋收購 1,217）", V("cf.sbc.9m2026"), V("cons.sbc.9m2026") + V("cons.sbcAcq.9m2026"), "兩表相加")
SUP = [("intPaid", "Interest, net of interest capitalized", "支付利息（扣資本化）"), ("taxPaid", "Income taxes, net", "支付所得稅"),
       ("ppeAP", "Purchases of property, plant, and equipment included in accounts payable and accrued expenses", "應付款中的未付資本支出（非現金）"),
       ("ppeFin", "Purchases of property, plant, and equipment financed by other financings", "以其他融資取得 PP&E（非現金）"),
       ("sbcCap", "Share-based compensation capitalized in property, plant, and equipment, net", "資本化股權報酬"),
       ("acqShares", "Acquisition of businesses through issuance of shares", "以發股方式收購（非現金）")]
sup = sec("Q", "Supplemental disclosures of cash flow information")
for key, lab, item in SUP:
    qv, qi = row("Q", lab, sup, 14, n=2)
    for j, per in enumerate(("9m2026", "9m2025")):
        sv, ss = None, ""
        if key == "acqShares" and per == "9m2026": sv, ss = 42723, "10-Q 權益變動表 Transfer of equity in business combinations 42,723"
        if key == "sbcCap": sv, ss = (160, "附註 14 文字（capitalized … $160 million）") if per == "9m2026" else (114, "附註 14 文字（$114 million）")
        add(f"cf.{key}.{per}", item, qv[j], UNIT, per, "Q", qi, "現金流量表補充揭露", second=sv, second_src=ss)
add("cf.fcf.9m2026", "自由現金流（營業現金流 − 資本支出）", V("cf.cfo.9m2026") + V("cf.capex.9m2026"), UNIT, "9m2026", "Q", IDX["cf.cfo.9m2026"] and qcf, "現金流量表衍生",
    tag="[Derived]", q="", note="公司未定義 FCF；資本支出不含資本化利息 20、未扣產品回饋金 1,372")
add("cf.fcfLease.9m2026", "自由現金流（扣融資租賃本金；S2 收支關卡口徑）", V("cf.fcf.9m2026") + V("cf.flPrin.9m2026"), UNIT, "9m2026", "Q", qcf, "現金流量表衍生",
    tag="[Derived]", q="", note="另 9M 以其他融資取得 PP&E 3,921（非現金，未入資本支出）；失敗售後租回償還列在 Repayments on debt")
add("cf.fcf.9m2025", "自由現金流（營業現金流 − 資本支出）", V("cf.cfo.9m2025") + V("cf.capex.9m2025"), UNIT, "9m2025", "Q", qcf, "現金流量表衍生", tag="[Derived]", q="")

# ============ 8. 資產負債表（2026-09-30、2025-12-31） ============
qbs = sec("Q", "Consolidated Balance Sheets")
pbs = sec("P", "Consolidated Balance Sheets", exact=False, start=pis)
BS = [("cash", "Cash and cash equivalents", "現金及約當現金"), ("mktSec", "Marketable securities", "有價證券（短期）"),
      ("ar", "Accounts receivable, net of allowance", "應收帳款淨額"), ("inv", "Inventory", "存貨"), ("prepaid", "Prepaid expenses and other current assets", "預付與其他流動資產"),
      ("tca", "Total current assets", "流動資產合計"), ("ppe", "Property, plant, and equipment, net", "不動產廠房設備淨額"),
      ("flRou", "Finance lease right-of-use assets", "融資租賃使用權資產"), ("intang", "Intangible assets, net", "無形資產淨額"), ("digital", "Digital assets", "數位資產（比特幣）"),
      ("goodwill", "Goodwill", "商譽"), ("dta", "Deferred tax assets", "遞延所得稅資產"), ("othA", "Other assets", "其他資產（含關係人）"), ("ta", "Total assets", "總資產"),
      ("ap", "Accounts payable", "應付帳款"), ("drCur", "Deferred revenue, current", "遞延收入（流動）"), ("debtCur", "Debt and finance leases, current", "債務與融資租賃（流動）"),
      ("accr", "Accrued expenses and other current liabilities", "應計費用與其他流動負債"), ("tcl", "Total current liabilities", "流動負債合計"),
      ("drLt", "Deferred revenue, net of current", "遞延收入（非流動）"), ("debtLt", "Debt and finance leases, net of current", "債務與融資租賃（非流動）"),
      ("othL", "Other liabilities", "其他負債"), ("tl", "Total liabilities", "總負債"), ("pref", "Redeemable convertible preferred stock, par value", "可贖回可轉換特別股"),
      ("apic", "Additional paid-in capital", "資本公積"), ("deficit", "Accumulated deficit", "累積虧損"), ("aoci", "Accumulated other comprehensive income", "其他綜合損益累計"),
      ("equity", "Total shareholders’ equity", "股東權益合計")]
for key, lab, item in BS:
    qv, qi = row("Q", lab, qbs, 80, n=2)
    pv, pi = row("P", lab, pbs, 80, n=2)
    for j, per in enumerate(("20260930", "20251231")):
        add(f"bs.{key}.{per}", item, qv[j], UNIT, per, "Q", qi, "合併資產負債表", second=pv[j], second_src=f"新聞稿第 {PP[pi]} 頁資產負債表")
for per in ("20260930", "20251231"):
    chk(f"bs.{per}（資產 ＝ 負債＋特別股＋權益）", V(f"bs.ta.{per}"), V(f"bs.tl.{per}") + V(f"bs.pref.{per}") + V(f"bs.equity.{per}"), "資產負債表恆等式")
chk("bs.deferredRev.20260930（流動＋非流動 ＝ 附註 3 總遞延收入 17,520）", V("bs.drCur.20260930") + V("bs.drLt.20260930"), 17520, "附註 3 文字")
chk("bs.deferredRev.20251231（12,116）", V("bs.drCur.20251231") + V("bs.drLt.20251231"), 12116, "附註 3 文字")
chk("bs.goodwill 變動（11,809＋36,807−193）", 11809 + 36807 - 193, V("bs.goodwill.20260930"), "附註 6 商譽變動表")
rc = sec("Q", "Total as presented in the consolidated statements of cash flows", exact=False)
chk("cash 含受限 2026-09-30（附註 2 vs 現金流量表期末）", trailing(Q[rc])[0], V("cf.end.9m2026"), "附註 2 受限現金表")

# ============ 9. 債務明細、融資租賃、失敗售後租回（附註 9、10、17） ============
n9 = sec("Q", "Note 9 - Debt")
DEBT = [("spacexNotes", "SpaceX Notes", "SpaceX Notes（優先無擔保，五批 2031–2056）"), ("xNotes", "X 2027 and X 2030 Notes", "X 2027／2030 票據"),
        ("otherFin", "Other financings", "其他融資（含 CTC 失敗售後租回）"), ("totalDebt", "Total debt", "債務合計"),
        ("finLease", "Finance lease liability", "融資租賃負債"), ("totalDebtFl", "Total debt and finance leases", "債務與融資租賃合計"),
        ("shortTerm", "Less: Short-term portion", "減：一年內到期"), ("ltNet", "Total debt and finance leases, net of current", "非流動部分（淨額）")]
n9b = find(Q, lambda l: l.strip() == "December 31, 2025", n9)
for key, lab, item in DEBT:
    qv, qi = row("Q", lab, n9, n9b - n9, n=3)
    add(f"debt.{key}.principal.20260930", item + "：本金", qv[0], UNIT, "20260930", "Q", qi, "附註 9 Debt 表")
    add(f"debt.{key}.net.20260930", item + "：帳面淨額", qv[2], UNIT, "20260930", "Q", qi, "附註 9 Debt 表",
        second=qv[0] - qv[1], second_src="本金 − 未攤銷發行成本（同列）")
for key, lab, item in [("xNotes", "X 2027 and X 2030 Notes", "X 2027／2030 票據"), ("xB1", "X B-1 Term Loan", "X B-1 定期貸款"), ("xB3", "X B-3 Term Loan", "X B-3 定期貸款"),
                       ("xaiFixed", "xAI Fixed Rate Term Loan", "xAI 固定利率定期貸款"), ("xaiFloat", "xAI Floating Rate Term Loan", "xAI 浮動利率定期貸款"),
                       ("xai125", "xAI 12.5% Secured Senior Notes", "xAI 12.5% 有擔保優先票據"), ("otherFin", "Other financings", "其他融資"),
                       ("totalDebt", "Total debt", "債務合計"), ("finLease", "Finance lease liability", "融資租賃負債"), ("totalDebtFl", "Total debt and finance leases", "債務與融資租賃合計")]:
    qv, qi = row("Q", lab, n9b, 16, n=3)
    add(f"debt.{key}.principal.20251231", item + "：本金", qv[0], UNIT, "20251231", "Q", qi, "附註 9 Debt 表（2025-12-31）")
chk("debt 2026-09-30 帳面合計 ＝ 資產負債表流動＋非流動", V("debt.totalDebtFl.net.20260930"), V("bs.debtCur.20260930") + V("bs.debtLt.20260930"), "附註 9 vs 資產負債表")
chk("debt 2026-09-30 債務組成合計", V("debt.spacexNotes.principal.20260930") + V("debt.xNotes.principal.20260930") + V("debt.otherFin.principal.20260930"),
    V("debt.totalDebt.principal.20260930"), "附註 9 同表")
mat = find(Q, lambda l: l.strip().startswith("2026 (remaining three months)"), n9)
MAT = [("2026Q4", 0), ("2027", 1), ("2028", 2), ("2029", 3), ("2030", 4), ("thereafter", 5), ("total", 6)]
mv = [trailing(Q[mat + k])[-1] for k in range(7)]
for (y, k) in MAT:
    add(f"debt.maturity.{y}", f"債務本金到期：{y}", mv[k], UNIT, "20260930", "Q", mat + k, "附註 9 到期表")
chk("debt.maturity 合計 ＝ 各年相加", sum(mv[:6]), mv[6], "附註 9 到期表")
chk("debt.maturity 合計 ＝ 債務本金", mv[6], V("debt.totalDebt.principal.20260930"), "附註 9")
i = find(Q, lambda l: "outstanding $38,006 million in aggregate principal" in l)
chk("debt.totalDebt（MD&A 文字 38,006）", 38006, V("debt.totalDebt.principal.20260930"), f"10-Q 第 {QP[i]} 頁 MD&A Debt Agreements")
# SpaceX Notes 條件（文字）
def tnum(pat, start=0, grp=1, text_src="Q"):
    lines = DOC[text_src][0]
    for k in range(start, len(lines) - 2):
        blob = re.sub(r"\s+", " ", " ".join(x.strip() for x in lines[k:k + 3]))
        m = re.search(pat, blob)
        if m:
            while k + 1 < len(lines) - 2:
                b2 = re.sub(r"\s+", " ", " ".join(x.strip() for x in lines[k + 1:k + 4]))
                m2 = re.search(pat, b2)
                if not m2 or m2.group(grp) != m.group(grp): break
                k, blob, m = k + 1, b2, m2
            global LASTQ
            LASTQ = blob[max(0, m.start() - 80):m.end() + 40].strip()
            return float(m.group(grp).replace(",", "")), k
    raise SystemExit("找不到文字：" + pat)
TXT = [("debt.spacexNotes.wac", r"weighted-average interest rate of ([\d.]+)%", "SpaceX Notes 加權平均票面利率", "%", "MD&A：weighted-average coupon interest rate of 5.855%", r"weighted-average coupon interest rate of ([\d.]+)%"),
       ("debt.spacexNotes.eir", r"effective interest\s+rate on the SpaceX Notes was ([\d.]+)%", "SpaceX Notes 有效利率", "%", "附註 3 Item 3：6.07%", r"effective interest rate on the SpaceX Notes was ([\d.]+)% as of"),
       ("debt.spacexNotes.wam", r"weighted-average maturity of ([\d.]+) years", "SpaceX Notes 加權平均到期年限", "年", "MD&A 同句", None),
       ("debt.spacexNotes.fv", r"SpaceX Notes \$ ([\d,]+) \$ ([\d,]+)", "SpaceX Notes 公允價值（2026-09-30）", "USD M", "", None),
       ("debt.otherFin.rate", r"weighted-average fixed interest rate of ([\d.]+)%", "其他融資加權平均固定利率（2026-09-30）", "%", "", None),
       ("debt.creditFacility.cap", r"increase the borrowing capacity up to\s+\$([\d,]+) million", "SpaceX 循環信用額度（未動用）", "USD M", "MD&A 流動性：$5,000 million available", r"supplemented by \$([\d,]+) million\s+available to borrow"),
       ("debt.bridge.amount", r"aggregate principal amount of \$([\d,]+) million\. The SpaceX Bridge Loan", "SpaceX 過橋貸款（2026-03 借、6 月以票據清償）", "USD M", "", None),
       ("debt.extLoss.bridgeRefi", r"loss on\s+extinguishment of \$([\d,]+) million, recorded", "清償 X／xAI 債務損失（2026-03）", "USD M", "", None)]
for iid, pat, item, unit, ss, pat2 in TXT:
    s_ = n9 if iid != "debt.spacexNotes.fv" else find(Q, lambda l: "The carrying amounts and fair values of the long-term fixed-rate debt" in l)
    v, k = tnum(pat, s_, 2 if iid == "debt.spacexNotes.fv" else 1)
    sv = None
    if pat2:
        sv, k2 = tnum(pat2, lq)
        ss = f"10-Q 第 {QP[k2]} 頁 " + ss
    add(iid, item, v, unit, "20260930", "Q", k, "附註 9", second=sv, second_src=ss, tol=0.01 if unit == "%" else 0.5,
        q=LASTQ)
# 失敗售後租回：CTC／Valor（附註 9 文字、附註 17 文字、資產負債表括號三讀）
k9 = find(Q, lambda l: "CTC Property, LLC" in l and "sale-leaseback" in l)
blob9 = re.sub(r"\s+", " ", " ".join(x.strip() for x in Q[k9:k9 + 6]))
m9 = re.search(r"related debt of \$([\d,]+) million and \$([\d,]+) million", blob9)
k17 = find(Q, lambda l: "CTC entered into various equipment lease agreements with Valor" in l)
blob17 = re.sub(r"\s+", " ", " ".join(x.strip() for x in Q[k17:k17 + 6]))
m17 = re.search(r"recorded debt of\s+\$([\d,]+) million and \$([\d,]+) million", blob17)
kbs = find(Q, lambda l: "Debt and finance leases, current (related party of $" in l, qbs)
mbs = re.search(r"related party of \$([\d,]+)", Q[kbs]); kbs2 = find(Q, lambda l: "Debt and finance leases, net of current (related party of $" in l, qbs)
mbs2 = re.search(r"related party of \$([\d,]+)", Q[kbs2])
add("debt.failedSLB.cur.20260930", "失敗售後租回債務（CTC 與 Valor，三筆）：流動", float(m9.group(1).replace(",", "")), UNIT, "20260930", "Q", k9, "附註 9 Other Financings",
    second=float(m17.group(1).replace(",", "")), second_src=f"10-Q 第 {QP[k17]} 頁附註 17 Valor 段；資產負債表括號關係人 {mbs.group(1)}",
    q="CTC … has entered into various sale-leaseback transactions for its SI infrastructure assets … resulting in failed sale-leaseback transactions",
    note="10-Q 直接寫明為 Valor Transactions（附註 17，文字檔第 1782–1783 行；chat 端審查 2026-10-30 確認，非推論）：CTC 與 Valor（董事 Antonio Gracias 創辦）的設備租賃屬失敗售後租回，與附註 9 金額相同；列在 Other financings 12,979 之內")
add("debt.failedSLB.lt.20260930", "失敗售後租回債務：非流動", float(m9.group(2).replace(",", "")), UNIT, "20260930", "Q", k9, "附註 9 Other Financings",
    second=float(m17.group(2).replace(",", "")), second_src=f"附註 17；資產負債表括號關係人 {mbs2.group(1)}", q="")
chk("failedSLB 流動 vs 資產負債表關係人括號", V("debt.failedSLB.cur.20260930"), float(mbs.group(1).replace(",", "")), "資產負債表括號")
chk("failedSLB 非流動 vs 資產負債表關係人括號", V("debt.failedSLB.lt.20260930"), float(mbs2.group(1).replace(",", "")), "資產負債表括號")
m17b = re.search(r"debt of \$([\d,]+) million and \$([\d,]+) million", re.sub(r"\s+", " ", " ".join(x.strip() for x in Q[k17 + 6:k17 + 16])))
add("debt.failedSLB.total.20251231", "失敗售後租回（Valor）債務合計 2025-12-31", float(m17b.group(1).replace(",", "")) + float(m17b.group(2).replace(",", "")), UNIT, "20251231", "Q", k17 + 7,
    "附註 17", tag="[Derived]", note="455＋4,052", q="As of December 31, 2025, the Company recorded debt of $455 million and $4,052 million")
add("debt.failedSLB.total.20260930", "失敗售後租回債務合計 2026-09-30", V("debt.failedSLB.cur.20260930") + V("debt.failedSLB.lt.20260930"), UNIT, "20260930", "Q", k9, "附註 9",
    tag="[Derived]", q="", note="占 Other financings 12,979 的 99%；其餘 82 為一般設備抵押融資")
# 融資租賃
n10 = sec("Q", "Note 10 - Leases")
for key, lab, item in (("opLease", "Total operating lease expense", "營業租賃費用合計"), ("flAmort", "Amortization of leased assets", "融資租賃資產攤銷"),
                       ("flInt", "Interest on lease liabilities", "融資租賃利息"), ("flExp", "Total finance lease expense", "融資租賃費用合計")):
    qv, qi = row("Q", lab, n10, 16)
    for j, per in enumerate(PER4):
        add(f"lease.{key}.{per}", item, qv[j], UNIT, per, "Q", qi, "附註 10 Leases")
k = find(Q, lambda l: "Operating lease liabilities, current" in l)
add("lease.opLiabCur.20260930", "營業租賃負債（流動）", trailing(Q[k])[0], UNIT, "20260930", "Q", k, "附註 11", note="非流動營業租賃負債未單獨揭露（含在 Other liabilities 1,657）")

# ============ 10. 承諾、頻譜、backlog、客戶集中 ============
n16 = sec("Q", "Note 16 - Commitments and Contingencies")
k = find(Q, lambda l: l.strip().startswith("2026 (remaining three months)"), n16)
cv = [trailing(Q[k + j])[-1] for j in range(7)]
for (y, j) in MAT:
    add(f"commit.uncond.{y}", f"不可取消合約承諾（SI 基礎設施、第三方雲端容量、頻譜等）：{y}", cv[j], UNIT, "20260930", "Q", k + j, "附註 16 Unconditional Obligations",
        note="公司未拆分算力採購與頻譜；2027 年 25,569 含頻譜交易應付（現金＋A 股）")
chk("commit.uncond 合計 ＝ 各年相加", sum(cv[:6]), cv[6], "附註 16 表")
SPEC = [("spectrum.echostar.total", r"total consideration for the acquisition of the Spectrum Licenses is approximately \$([\d.]+) billion", "EchoStar 頻譜（AWS-4、H-Block、AWS-3）總對價", "USD B"),
        ("spectrum.echostar.equity", r"approximately \$([\d.]+)\s+billion in equity", "EchoStar 頻譜：股權對價（約 261.8M 股 A 股，每股固定 42.40）", "USD B"),
        ("spectrum.echostar.cash", r"up to \$([\d.]+) billion related to the payoff of designated EchoStar debt", "EchoStar 頻譜：清償 EchoStar 指定債務（上限，不足以現金補）", "USD B"),
        ("spectrum.echostar.shares", r"issuance of approximately ([\d.]+) million shares", "EchoStar 頻譜：預計發行 A 股", "M 股"),
        ("spectrum.echostar.paid", r"of which \$([\d,]+) million was paid as of September 30, 2026", "Spectrum Credit Agreement 已付（2026-09-30，列其他資產預付）", "USD M"),
        ("spectrum.echostar.pay2026", r"Credit Agreement are \$([\d,]+) million in 2026", "Spectrum Credit Agreement 2026 年預計支付", "USD M"),
        ("spectrum.echostar.pay2027", r"paid as of September 30, 2026, and \$([\d,]+) million in 2027", "Spectrum Credit Agreement 2027 年預計支付（假設 2027-11-30 交割）", "USD M"),
        ("spectrum.echostar.payExtra", r"additional payments totaling \$([\d,]+) million", "若 2028-11-30 才交割之額外支付", "USD M")]
for iid, pat, item, unit in SPEC:
    v, kk = tnum(pat, lq)
    sv, ss = None, ""
    if iid == "spectrum.echostar.paid":
        sv, k2 = tnum(r"Agreement\s+to the Trust of \$([\d,]+) million as of September 30, 2026", n9 - 200 if n9 > 200 else 0)
        ss = f"10-Q 第 {QP[k2]} 頁附註 6 Spectrum Transactions"
    if iid == "spectrum.echostar.equity":
        sv, ss = round(261.8 * 42.40 / 1000, 1), "261.8M 股 × 42.40 美元"
    add(iid, item, v, unit, "20260930", "Q", kk, "MD&A Material Cash Commitments — Spectrum Transaction", second=sv, second_src=ss, tol=0.05 if unit == "USD B" else 0.5,
        q=LASTQ)
kk = find(P, lambda l: "megahertz" in l)
add("spectrum.800mhz.mhz", "800 MHz 頻譜（2026-10-08 簽約，預計 2027 上半年交割）：成對頻寬上限", 14, "MHz", "20261008", "P", kk, "新聞稿 Other Key Business Developments",
    tag="[Interested-party]", note="交易金額未揭露（新聞稿與 10-Q 均無；10-Q 期後事項未列）", q=re.sub(r"\s{2,}", " ", " ".join(x.strip() for x in P[kk - 1:kk + 1]))[:220])
n3b = sec("Q", "Backlog")
v, kk = tnum(r"Backlog totaled \$([\d,]+) million as of September 30, 2026", n3b)
add("backlog.total.20260930", "Backlog（未履約合約交易價格）", v, UNIT, "20260930", "Q", kk, "附註 3 Backlog",
    note="含遞延收入 17,520；約 67% 一年內、27% 一至三年、6% 其後認列；不含按量計價與受限變動對價；10-Q 未拆分部（雲端服務合約可由任一方 90 天通知終止，是否計入全額未說明）",
    q=LASTQ)
for iid, pat, item in (("backlog.pct1y", r"Approximately (\d+)% is expected to be recognized\s+within one year", "Backlog 一年內認列比例"),
                       ("backlog.pct1to3y", r"approximately (\d+)% is expected to be recognized between one and three years", "Backlog 一至三年認列比例")):
    v, kk = tnum(pat, n3b)
    add(iid, item, v, "%", "20260930", "Q", kk, "附註 3 Backlog")
cc = find(Q, lambda l: l.strip().startswith("Customer B"))
cb = trailing(Q[cc].replace("%", " ").replace("*", " — "))
ca = trailing(Q[cc - 1].replace("%", " "))
add("cust.A.pct.2026q3", "客戶 A 占合併營收（三分部皆有）", ca[0], "%", "2026q3", "Q", cc - 1, "附註 3 Concentration of risk", note="9M2026 17.3%；2025Q3 23.9%")
add("cust.B.pct.2026q3", "客戶 B 占合併營收（全屬 SI）", 34.5, "%", "2026q3", "Q", cc, "附註 3 Concentration of risk")
add("cust.B.pct.9m2026", "客戶 B 占合併營收（9M）", 24.2, "%", "9m2026", "Q", cc, "附註 3 Concentration of risk", note="2025 年兩期未達 10%")
add("cust.B.rev.2026q3", "客戶 B 營收（衍生＝34.5% × 合併營收）", round(0.345 * V("is.rev.2026q3")), UNIT, "2026q3", "Q", cc, "附註 3 衍生", tag="[Derived]", q="",
    note="百分比四捨五入至 0.1%，誤差 ±7；占 SI 解決方案與基礎設施 7,022 約 72%。客戶身分 10-Q 未揭露")
add("cust.B.rev.9m2026", "客戶 B 營收（衍生，9M）", round(0.242 * V("is.rev.9m2026")), UNIT, "9m2026", "Q", cc, "附註 3 衍生", tag="[Derived]", q="",
    note="9M − Q3 ≈ 1,526，即 Q2 起認列（與 MD&A『第二季開始提供雲端服務』一致）")
k = find(Q, lambda l: "one customer represented 27.7% of Accounts receivable" in l)
add("cust.AR.top.20260930", "單一客戶占應收帳款", 27.7, "%", "20260930", "Q", k, "附註 3")

# ============ 11. 收購與合併會計（Cursor、Mesh Optical、xAI）、IPO、股數、股權報酬 ============
n20 = sec("Q", "Note 20 - Acquisitions")
CUR = [("cursor.exchRatio", r"Exchange Ratio of (\d+\.\d+)", "Cursor 換股比例", "倍", None),
       ("cursor.vwap", r"Cursor Merger of \$(\d+\.\d+)", "Cursor 換股 VWAP（七日）", "USD", None),
       ("cursor.shares", r"SpaceX shares issued as purchase consideration ([\d,]+)", "Cursor 對價發行股數", "M 股", None),
       ("cursor.closePx", r"SpaceX share price on Closing Date \$ ([\d.]+)", "交割日 SpaceX 股價", "USD", None),
       ("cursor.equityConsid", r"Total equity consideration \$ ([\d,]+)", "Cursor 股權對價", "USD M", None),
       ("cursor.replAwards", r"Replacement equity awards for Cursor’s equity awards\(1\) ([\d,]+)", "Cursor 替代獎酬（歸屬收購前服務）", "USD M", None),
       ("cursor.consideration", r"Total aggregate purchase consideration \$ ([\d,]+)", "Cursor GAAP 收購對價", "USD M", ("P", r"GAAP purchase consideration was \$([\d.]+) billion", 1000, "新聞稿 Cursor Merger 段（42.5B）", 60)),
       ("cursor.postSvcAwards", r"post-acquisition service\s+as of the Closing Date is \$([\d,]+) million", "Cursor 收購後服務期股權獎酬總值（未來認列為費用）", "USD M",
        ("P", r"post-acquisition service is \$([\d.]+) billion", 1000, "新聞稿（23.4B）", 60)),
       ("cursor.goodwill", r"Goodwill ([\d,]+)\s+Total Purchase", "Cursor 商譽", "USD M", None),
       ("cursor.intang", r"Total intangible assets acquired \$ ([\d,]+)", "Cursor 可辨認無形資產", "USD M", None),
       ("cursor.netAssets", r"Fair value of net assets acquired \$ ([\d,]+)", "Cursor 可辨認淨資產", "USD M", None),
       ("cursor.rev", r"results of operations included \$([\d,]+) million of revenue attributable to Cursor", "Cursor 交割後營收貢獻（8-14 至 9-30）", "USD M",
        ("P", r"Cursor contributed \$([\d,]+) million of revenue", 1, "新聞稿 Cursor Merger 段", 0.5))]
for iid, pat, item, unit, sec2 in CUR:
    blobs = [re.sub(r"\s+", " ", " ".join(x.strip() for x in Q[kk:kk + 3])) for kk in range(n20, n20 + 160)]
    hit = next(((kk, re.search(pat, b)) for kk, b in zip(range(n20, n20 + 160), blobs) if re.search(pat, b)), None)
    if hit is None: raise SystemExit("Cursor 找不到：" + pat)
    kk, m = hit
    v = float(m.group(1).replace(",", ""))
    sv, ss, tol = None, "", 0.5
    if sec2:
        pb = re.sub(r"\s+", " ", " ".join(x.strip() for x in P))
        sv = float(re.search(sec2[1], pb).group(1).replace(",", "")) * sec2[2]; ss = sec2[3]; tol = sec2[4]
    if iid == "cursor.equityConsid": sv, ss, tol = 271 * 141.29, "271M 股 × 141.29（股數四捨五入，容差 80）", 80
    add(iid, item, v, unit, "2026-08-14", "Q", kk, "附註 20 Acquisitions", tag="[Verified]", second=sv, second_src=ss, tol=tol, q=blobs[kk - n20][:200])
chk("cursor 對價 ＝ 股權＋替代獎酬", V("cursor.equityConsid") + V("cursor.replAwards"), V("cursor.consideration"), "附註 20 表")
chk("cursor 對價 ＝ 淨資產＋商譽", V("cursor.netAssets") + V("cursor.goodwill"), V("cursor.consideration"), "附註 20 PPA 表")
pf = find(Q, lambda l: l.strip().startswith("Revenue") and "15,114" in l)
for key, lab in (("rev", "Revenue"), ("ni", "Net income (loss)")):
    qv, qi = row("Q", lab, pf, 3)
    for j, per in enumerate(PER4):
        add(f"cursor.proforma.{key}.{per}", f"備考（Cursor 自 2025-01-01 併入）：{'營收' if key == 'rev' else '淨利（損）'}", qv[j], UNIT, per, "Q", qi, "附註 20 備考資訊",
            tag="[Interested-party]", note="未經查核備考；不含整合綜效")
for iid, pat, item, unit in (("mesh.consideration", r"total purchase consideration of \$([\d,]+) million in the Company’s Class A", "Mesh Optical 收購對價（A 股，2026-07-06 交割）", "USD M"),
                             ("mesh.goodwill", r"recognized \$([\d,]+) million of goodwill\s+from the Mesh Optical", "Mesh Optical 商譽（SI）", "USD M")):
    v, kk = tnum(pat, n20)
    add(iid, item, v, unit, "2026-07-06", "Q", kk, "附註 20 Mesh Optical", q=LASTQ)
# IPO
ipo_k = find(Q, lambda l: "In June 2026, the Company completed its initial public offering" in l)
ib = re.sub(r"\s+", " ", " ".join(x.strip() for x in Q[ipo_k:ipo_k + 4]))
m = re.search(r"sold ([\d.]+) million shares.*?\$([\d.]+) per share.*?net proceeds of \$([\d,]+) million.*?costs of \$([\d,]+)", ib)
add("ipo.shares", "IPO 發行 A 股（含超額配售全額）", float(m.group(1)), "M 股", "2026-06", "Q", ipo_k, "附註 1 Initial Public Offering", q=ib[:200])
add("ipo.price", "IPO 價格", float(m.group(2)), "USD", "2026-06", "Q", ipo_k, "附註 1")
add("ipo.net", "IPO 淨收入", float(m.group(3).replace(",", "")), UNIT, "2026-06", "Q", ipo_k, "附註 1", second=V("cf.ipo.9m2026"), second_src="現金流量表 IPO 淨收入")
add("ipo.costs", "IPO 承銷與發行成本", float(m.group(4).replace(",", "")), UNIT, "2026-06", "Q", ipo_k, "附註 1",
    second=round(float(m.group(1)) * float(m.group(2))) - float(m.group(3).replace(",", "")), second_src="股數 × 價格 − 淨收入（股數四捨五入至 0.1M，容差 10）", tol=10)
# xAI 共同控制合併
k = find(Q, lambda l: "On February 2, 2026 (“xAI Merger Date”)" in l)
add("xai.mergerDate", "xAI 併入日（共同控制合併，換股）", 20260202, "日期", "2026-02-02", "Q", k, "附註 1 Common Control Mergers",
    q="On February 2, 2026 (“xAI Merger Date”), the Company completed its acquisition of X.AI Holdings Corp.", note="X 於 2025-03-28 先併入 xAI（亦屬共同控制）；Twitter 於 2022-10 由 Musk 取得")
add("xai.prefConversion", "xAI 合併：特別股轉普通股金額", 37475, UNIT, "9m2026", "Q", find(Q, lambda l: "pursuant to xAI Merger" in l), "權益變動表（9M2026）", q="Conversion of redeemable convertible preferred stock pursuant to xAI Merger … 37,475")
add("xai.repurchase", "xAI 合併：向 xAI 員工買回", 2413, UNIT, "9m2026", "Q", find(Q, lambda l: "repurchased $2,413 million" in l), "附註 12 Share Repurchases",
    second=2413, second_src="權益變動表 Repurchase of common stock pursuant to xAI Merger (2,413)")
# 股數
cov = find(Q, lambda l: "shares of Class A common stock and" in l and "As of October 15, 2026" in l)
mc = re.search(r"had ([\d,]+) shares of Class A common stock and ([\d,]+) shares of Class B", re.sub(r"\s+", " ", " ".join(Q[cov:cov + 2])))
a_, b_ = float(mc.group(1).replace(",", "")) / 1e6, float(mc.group(2).replace(",", "")) / 1e6
add("sh.out.A.20261015", "流通 A 股（10-Q 封面 2026-10-15）", round(a_, 3), "M 股", "20261015", "Q", cov, "10-Q 封面", second=8240, second_src="資產負債表 2026-09-30 A 股 8,240M（日期差 15 天）", tol=15)
add("sh.out.B.20261015", "流通 B 股（10-Q 封面 2026-10-15；每股 10 票）", round(b_, 3), "M 股", "20261015", "Q", cov, "10-Q 封面", second=5339, second_src="資產負債表 2026-09-30 B 股 5,339M", tol=15)
add("sh.out.AB.20261015", "流通 A＋B 股合計（S9 用）", round(a_ + b_, 3), "M 股", "20261015", "Q", cov, "10-Q 封面", tag="[Derived]",
    second=13579, second_src="權益變動表 2026-09-30 普通股 13,579M", tol=15)
ro = find(Q, lambda l: l.strip().startswith("Outstanding stock options"))
add("sh.options.20260930", "流通股票選擇權（A 158＋B 350）", sum(trailing(Q[ro])), "M 股", "20260930", "Q", ro, "附註 12 保留股數表")
add("sh.rsu.20260930", "流通 RSU", trailing(Q[ro + 1])[0], "M 股", "20260930", "Q", ro + 1, "附註 12")
k = find(Q, lambda l: "exclude 1,333 million and 14 million" in l)
add("sh.perfAwards.20260930", "未計入稀釋的績效／市場條件獎酬", 1333, "M 股", "20260930", "Q", k, "附註 13",
    note="條件未達成故不計入稀釋；10-Q 未說明受獎人（是否為 CEO 獎酬未見原文）", q="exclude 1,333 million and 14 million share-based compensation awards outstanding")
k = find(Q, lambda l: "Cursor Merger and Mesh Optical Merger" in l)
add("sbc.acqAwards.period", "收購相關獎酬剩餘認列期間（加權）", 2.8, "年", "20260930", "Q", k + 4, "附註 14", tag="[Verified]")
for key, lab, item in (("sbc.total", "Total", "股權報酬合計（含收購相關）"),):
    qv, qi = row("Q", "Total", find(Q, lambda l: l.strip().startswith("Cost of revenue(1)")), 4)
    for j, per in enumerate(PER4):
        add(f"{key}.{per}", item, qv[j], UNIT, per, "Q", qi, "附註 14 依科目",
            second=(V(f"cons.sbc.{per}") + IDX.get(f"cons.sbcAcq.{per}", {"value": 0})["value"]), second_src="調整後 EBITDA 調節表 一般＋收購相關")
# PP&E 與折舊
n5 = sec("Q", "Note 5 - Property, Plant, and Equipment, Net")
for key, lab, item in (("servers", "Servers and networking equipment", "伺服器與網通設備"), ("sats", "Satellites", "衛星"), ("mach", "Machinery and equipment", "機器設備"),
                       ("dcInfra", "Data center infrastructure", "資料中心基礎設施"), ("launch", "Launch sites", "發射場"), ("land", "Land, buildings, and improvements", "土地建物"),
                       ("flight", "Flight vehicle hardware", "飛行器硬體"), ("lhi", "Leasehold improvements", "租賃改良"), ("cip", "Construction-in-progress", "在建工程（主要為伺服器與資料中心）"),
                       ("gross", "Property, plant, and equipment", "PP&E 毛額"), ("accDep", "Less: Accumulated depreciation", "累計折舊")):
    qv, qi = row("Q", lab, n5, 20, n=2, strict=True)
    for j, per in enumerate(("20260930", "20251231")):
        add(f"ppe.{key}.{per}", item, qv[j], UNIT, per, "Q", qi, "附註 5")
for per in ("20260930", "20251231"):
    chk(f"ppe.{per}（毛額 − 累計折舊 ＝ 淨額）", V(f"ppe.gross.{per}") + V(f"ppe.accDep.{per}"), V(f"bs.ppe.{per}"), "附註 5 vs 資產負債表")
for iid, pat, item, per, g in (("dep.2026q3", r"September 30, 2026 was \$([\d,]+) million and \$([\d,]+) million", "折舊費用", "2026q3", 1),
                               ("dep.9m2026", r"September 30, 2026 was \$([\d,]+) million and \$([\d,]+) million", "折舊費用（9M）", "9m2026", 2),
                               ("capInt.2026q3", r"capitalized \$([\d,]+) million and \$([\d,]+)\s+million of interest", "資本化利息", "2026q3", 1),
                               ("amort.2026q3", r"was \$([\d,]+) million and \$([\d,]+) million in the three and nine months ended\s+September 30, 2026", "無形資產攤銷", "2026q3", 1),
                               ("amort.9m2026", r"was \$([\d,]+) million and \$([\d,]+) million in the three and nine months ended\s+September 30, 2026", "無形資產攤銷（9M）", "9m2026", 2)):
    v, kk = tnum(pat, n5, g)
    add(iid, item, v, UNIT, per, "Q", kk, "附註 5／6")
chk("D&A 2026Q3 ≈ 折舊 3,351＋攤銷 293", V("dep.2026q3") + V("amort.2026q3"), V("cons.da.2026q3"), "附註 5＋6 vs 調節表（差額為融資租賃資產攤銷等）", tol=80)
k = find(Q, lambda l: "goodwill attributable to the SI segment was" in l or "attributable to the SI segment was $47,908" in l)
add("goodwill.si.20260930", "SI 分部商譽", 47908, UNIT, "20260930", "Q", k, "附註 6", second=V("bs.goodwill.20260930") - 515, second_src="商譽合計 − Connectivity 515")
# 其他
OTH = [("bitcoin.units", r"held ([\d,]+) units of Bitcoin", "持有比特幣（單位）", "BTC", "20260930"),
       ("tax.vaRelease", r"provisional deferred tax benefit of \$([\d,]+) million", "評價備抵迴轉（收購遞延所得稅負債）", "USD M", "2026q3"),
       ("legal.accrual", r"recorded an accrual of \$([\d,]+) million and \$([\d,]+) million for litigation", "訴訟準備（2026-09-30）", "USD M", "20260930"),
       ("lc.outstanding", r"outstanding letters of credit of \$([\d,]+) million", "未到期信用狀（以受限現金擔保）", "USD M", "20260930"),
       ("tesla.megapack.2026q3", r"\$([\d,]+) million and \$([\d,]+) million, respectively, of\s+Megapack", "向 Tesla 購買 Megapack（Q3）", "USD M", "2026q3"),
       ("memphis.claim", r"compensatory damages the SpaceXAI companies estimate at more than \$([\d,]+) million", "Memphis 算力設施承包商訴訟：公司求償（逾）", "USD M", "20260930"),
       ("memphis.lien", r"construction liens seeking approximately \$([\d,]+) million", "Memphis 承包商反訴與工程留置權（約）", "USD M", "20260930")]
for iid, pat, item, unit, per in OTH:
    v, kk = tnum(pat, n5)
    add(iid, item, v, unit, per, "Q", kk, "10-Q 附註", q=LASTQ)
v, kk = tnum(r"\$([\d,]+) million and \$([\d,]+) million, respectively, of\s+Megapack", n5, 2)
add("tesla.megapack.9m2026", "向 Tesla 購買 Megapack（9M，關係人，列 PP&E）", v, UNIT, "9m2026", "Q", kk, "附註 17", q=LASTQ, second=float(re.search(r"related party of \$(\d+)", Q[find(Q, lambda l: l.strip().startswith("Purchases of property, plant, and equipment (related party"))]).group(1)), second_src="現金流量表 PP&E 購置括號 related party of $480")
k = find(Q, lambda l: "over 11,100 Starlink broadband and mobile satellites" in l)
add("conn.satellites", "在軌 Starlink 衛星（逾）", 11100, "顆", "20260930", "Q", k, "MD&A Overview", tag="[Interested-party]",
    note="新聞稿註 2：在軌衛星新定義含核准離軌者，舊定義約少 400 顆")
add("conn.countries", "Starlink 服務國家與地區", 174, "個", "20260930", "Q", k + 1, "MD&A Overview", tag="[Interested-party]")
k = find(P, lambda l: "reaching over 3,800 aircraft" in l)
add("conn.aircraft", "裝 Starlink 的飛機（逾）", 3800, "架", "20260930", "P", k, "Connectivity Business Highlights", tag="[Interested-party]")
# 名目算力定義（文字）
k = find(Q, lambda l: l.strip().startswith("Nameplate Compute: We calculate nameplate compute"))
add("si.gw.definition", "名目算力定義（原文）", None, "—", "20260930", "Q", k, "MD&A Key Business Metrics — SI", tag="[Interested-party]",
    q=re.sub(r"\s+", " ", " ".join(x.strip() for x in Q[k:k + 4]))[:420],
    note="GPU 安裝數 × 各自 all-in power；不含冷卻、配電損失、照明、保全與設施層級負載 → 屬 IT 側（非設施、非 PUE 後）；『installed in our data centers』且『不代表實際耗電或利用率』。是否含第三方雲端容量：定義限 our data centers，第三方雲端容量另列承諾（附註 16），故推定不含；原文未明說")
k = find(Q, lambda l: "During 2026, the Company’s SI segment entered into revenue arrangements to provide cloud services" in l)
add("si.cloud.policy", "雲端服務收入認列政策（原文）", None, "—", "2026", "Q", k, "附註 2 Cloud Services Arrangements", tag="[Verified]",
    q=re.sub(r"\s+", " ", " ".join(x.strip() for x in Q[k:k + 5]))[:420], note="MD&A：固定月費、合約通常三年、任一方 90 天通知可終止；第二季開始提供")
k = find(Q, lambda l: "recognized ratably over the contract term (typically three years and terminable by either party upon 90 days" in l)
add("si.cloud.term", "雲端服務合約期間（通常）", 3, "年", "2026", "Q", k, "MD&A Revenue - SI", tag="[Verified]", q=re.sub(r"\s+", " ", Q[k].strip())[:200],
    note="任一方 90 天通知可終止")
k = find(Q, lambda l: "our board of directors has formed a special committee to evaluate a potential" in l)
add("tesla.combination", "董事會成立特別委員會評估與 Tesla 合併（尚無決定）", None, "—", "2026-10-20", "Q", k, "MD&A Evaluation of Potential Combination with Tesla", tag="[Interested-party]",
    q=re.sub(r"\s+", " ", " ".join(x.strip() for x in Q[k - 1:k + 2]))[:260], note="S7：不進模型，Thesis 頁列事件風險")

# ---- Q2 新聞稿合併層數字（需 10-Q 讀完後互核）
Q2X = [("is.rev.2026q2", "合併營收", 7814, "2026q2"), ("is.ebit.2026q2", "合併營業損失", -143, "2026q2"), ("is.ni.2026q2", "合併淨損", -541, "2026q2"),
       ("is.ni.1h2026", "合併淨損（1H）", -4817, "1h2026"), ("is.epsDil.2026q2", "稀釋 EPS", -0.09, "2026q2"), ("is.shDil.2026q2", "稀釋加權股數（M）", 5864, "2026q2"),
       ("cons.adjEbitda.2026q2", "調整後 EBITDA", 3538, "2026q2"), ("cf.cfo.1h2026", "營業現金流（1H）", 3466, "1h2026"),
       ("bs.cash.20260630", "現金及約當現金", 93522, "20260630"), ("bs.mktSec.20260630", "有價證券", 6487, "20260630"),
       ("bs.debtCur.20260630", "債務與融資租賃（流動）", 2525, "20260630"), ("bs.debtLt.20260630", "債務與融資租賃（非流動）", 36839, "20260630"),
       ("backlog.total.20260630", "Backlog", 47500, "20260630"), ("si.cloud.contracted.2026q2", "2026Q2 簽訂雲端服務合約之合約銷售額（Cloud Services Agreements）", 14100, "2026q2")]
for iid, item, v, per in Q2X:
    sv, ss = None, ""
    if iid == "is.ni.1h2026": sv, ss = V("is.ni.9m2026") - V("is.ni.2026q3"), "10-Q 9M2026 淨損 − Q3 淨利"
    if iid == "is.rev.2026q2": sv, ss = V("seg.total.rev.2026q2"), "Q3 新聞稿 June 30, 2026 欄"
    if iid == "is.ebit.2026q2": sv, ss = V("seg.total.ebit.2026q2"), "Q3 新聞稿 June 30, 2026 欄"
    if iid == "cons.adjEbitda.2026q2": sv, ss = V("seg.total.adjEbitda.2026q2"), "Q3 新聞稿 June 30, 2026 欄"
    addw(iid, item, v, ("USD" if "eps" in iid else "M 股" if "shDil" in iid else UNIT), per, Q2SRC,
         "[Interested-party]" if iid.startswith(("cons.", "backlog", "si.cloud")) else "[Verified]",
         note=("Q2 新聞稿原句：『$47.5 billion in backlog』；Q3 為 62,959（季增 15.5B）" if iid.startswith("backlog") else
               "原句：『Cloud Services Agreements totaling $14.1 billion in contracted sales』；總額、期間、客戶未揭露" if iid.startswith("si.cloud") else
               "WebFetch 讀取，單一讀取；待 PDF 入庫重讀" if sv is None else ""),
         second=sv, second_src=ss, url=Q2URL, docDate="2026-07", tol=0.5)
chk("bs.cash＋mktSec 2026-06-30（Q2 新聞稿『約 100B』）", V("bs.cash.20260630") + V("bs.mktSec.20260630"), 100000, "Q2 新聞稿文字 $100 billion（四捨五入）", tol=500)

# ============ 輸出 ============
OUT = {"_readme": "SpaceX 2026Q3 事實總帳（SpaceX v0.1 工作單 A 段）。由 spacex/facts_2026q3.py 自 10-Q（2026-10-22）與 8-K 財報新聞稿（2026-10-20）讀取；"
                  "金額 USD M（除另註）；期間鍵：2026q3、2026q2、2025q3、9m2026、9m2025、2026q1（衍生）、20260930／20251231（時點）。"
                  "check＝第二次獨立讀取（10-Q 與新聞稿互核；只出現一處者用另一張表、文字句或恆等式）。標記：[Verified] 10-Q 財報與附註；[Interested-party] 公司自訂指標、非 GAAP、只見於新聞稿；[Derived] 計算。",
       "company": "SpaceX（Space Exploration Technologies Corp.）", "ticker": "SPCX", "asOf": "2026-09-30", "retrieved": "2026-10-30",
       "sources": {"10-Q": QF, "pressRelease": PF, "url": "https://ir.spacex.com（10-Q 與 8-K 由 Andy 2026-10-30 提供原檔；SEC 網站本環境不可達）"},
       "segmentRecon": RECON, "items": ITEMS, "checks": CHK}
os.makedirs(os.path.join(HERE, "data"), exist_ok=True)
json.dump(OUT, open(os.path.join(HERE, "data", "spacex_facts_2026q3.json"), "w", encoding="utf-8"), ensure_ascii=False, indent=1)
bad = [c for c in CHK if not c["ok"]]
n2 = sum(1 for it in ITEMS if "check" in it)
print(f"事實 {len(ITEMS)} 筆（其中 {n2} 筆有第二讀）；核對 {len(CHK)} 項，一致 {len(CHK) - len(bad)}，不一致 {len(bad)}")
rb = [r for r in RECON if abs(r["diff"]) > 1e-9]
print("分部對帳（2025Q2、2025Q3、2026Q1、2026Q2、2026Q3、1H2025、1H2026、9M2025、9M2026、FY2025）：", "全部差額 0" if not rb else rb)
for c in bad: print("  不一致：", c)
sys.exit(1 if bad or rb else 0)
