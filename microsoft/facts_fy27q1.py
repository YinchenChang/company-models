# Microsoft FY27Q1（季末 2026-09-30）事實總帳：兩次獨立讀取（財報新聞稿 docx 表格 vs 公司 IR 財務報表 xlsx；必要時加 10-Q 文字）後寫成 JSON。
# 用法：python3 microsoft/facts_fy27q1.py  → microsoft/data/microsoft_facts_FY27Q1_20261029.json；兩讀不一致時列在 mismatches 並以代碼 1 結束（已解釋者除外）。
# 原文：mag/extract/docs/2026Q3/（Andy 2026-10-29 提供）；法說逐字稿不在 repo，本檔只記必要數字與短語（「法說」標 Interested-party）。
import json
import os
import re
import sys

import docx
import openpyxl

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
D = os.path.join(ROOT, "mag/extract/docs/2026Q3")
PR = "mag/extract/docs/2026Q3/microsoft_FY27Q1_新聞稿.docx"
XL = "mag/extract/docs/2026Q3/microsoft_FY27Q1_財務報表.xlsx"
TQ = "mag/extract/docs/2026Q3/text/microsoft_FY27Q1_10-Q.txt"
PP = "mag/extract/docs/2026Q3/microsoft_FY27Q2_展望簡報.pptx"
CALL = "FY27Q1 法說逐字稿（2026-10-28；不在 repo，chat 端提供）"
OUT = os.path.join(HERE, "data/microsoft_facts_FY27Q1_20261029.json")


def num(s):
    s = str(s).replace("$", "").replace(",", "").replace("\xa0", "").strip()
    if s in ("", "-", "—"):
        return None
    neg = s.startswith("(") and s.endswith(")")
    s = s.strip("()")
    try:
        v = float(s)
    except ValueError:
        return None
    return -v if neg else v


# ---------- 讀取 A：財報新聞稿 docx（表格） ----------
doc = docx.Document(os.path.join(ROOT, PR))


def table_rows(ti):
    out = []
    for r in doc.tables[ti].rows:
        cells = []
        for c in r.cells:
            x = c.text.strip()
            if not cells or x != cells[-1]:
                cells.append(x)
        out.append(cells)
    return out


def pr_find(ti, label, nth=0, col=1, section=None):
    """docx 第 ti 個表格中 A 欄以 label 開頭的第 nth 列；section＝先找到該段標題列之後才開始比對"""
    seen = 0
    active = section is None
    for cells in table_rows(ti):
        a = re.sub(r"\s+", " ", cells[0])
        if section and a.startswith(section):
            active = True
            continue
        if active and a.startswith(label):
            if seen == nth:
                vals = [num(x) for x in cells[1:] if num(x) is not None]
                return vals[col - 1] if len(vals) >= col else None
            seen += 1
    raise KeyError((ti, label))


# ---------- 讀取 B：IR 財務報表 xlsx ----------
wb = openpyxl.load_workbook(os.path.join(ROOT, XL), data_only=True)


def xl_hist(sheet, label, colhdr="Q1-27", block=None):
    ws = wb[sheet]
    hdr_row, hdr = None, {}
    active = block is None
    for r in range(1, ws.max_row + 1):
        a = str(ws.cell(r, 1).value or "").strip()
        if block and a == block:
            active = True
            hdr = {str(ws.cell(r, c).value).strip(): c for c in range(2, ws.max_column + 1) if ws.cell(r, c).value is not None}
            continue
        if not block and colhdr in [str(ws.cell(r, c).value).strip() for c in range(2, ws.max_column + 1)]:
            hdr = {str(ws.cell(r, c).value).strip(): c for c in range(2, ws.max_column + 1) if ws.cell(r, c).value is not None}
            continue
        if active and hdr and a.startswith(label):
            return ws.cell(r, hdr[colhdr]).value
    raise KeyError((sheet, label, colhdr))


def xl_col(sheet, label, col=2):
    ws = wb[sheet]
    for r in range(1, ws.max_row + 1):
        a = re.sub(r"\s+", " ", str(ws.cell(r, 1).value or "")).strip()
        if a.startswith(label):
            return ws.cell(r, col).value
    raise KeyError((sheet, label))


def capex_q127():
    ws = wb["CapEx"]
    hdr = {str(ws.cell(5, c).value).strip(): c for c in range(1, ws.max_column + 1) if ws.cell(5, c).value is not None}
    return {k: ws.cell(6, c).value for k, c in hdr.items()}


def roc(label, q="Q1-27"):
    ws = wb["Return of Cash to Shareholders"]
    # 季度區塊：表頭列含 Q1-27，其下為 Dividend Per Share、Total Dividends、Share Buyback
    for r in range(1, ws.max_row + 1):
        row = [str(ws.cell(r, c).value).strip() for c in range(1, ws.max_column + 1)]
        if q in row:
            c = row.index(q) + 1
            for rr in range(r + 1, r + 6):
                if str(ws.cell(rr, 1).value or "").strip().startswith(label):
                    return ws.cell(rr, c).value
    raise KeyError(label)


TQT = open(os.path.join(ROOT, TQ), encoding="utf-8").read()


def tq_search(pat, grp=1):
    m = re.search(pat, TQT, re.S)
    if not m:
        raise KeyError(pat)
    return num(m.group(grp))


items, mism = [], []


def add(group, id_, item, value, unit, period, src, loc, tag, note="", src2=None, loc2=None, value2=None, tol=0.5, explained=None):
    it = dict(id=id_, item=item, value=value, unit=unit, period=period, source=src, loc=loc, tag=tag, note=note, docDate="2026-10-28", retrieved="2026-10-29")
    if src2 is not None:
        ok = value2 is not None and value is not None and abs(value - value2) <= tol
        it["check"] = dict(source=src2, loc=loc2, value=value2, match=bool(ok))
        if not ok:
            it["check"]["explained"] = explained
            mism.append(dict(id=id_, a=value, b=value2, explained=explained))
    items.append((group, it))


P = "FY27Q1（季末 2026-09-30）"
QP = "Quarterly Income Statements"
# 1. 分部與產品線（新聞稿 SEGMENT RESULTS／PRODUCT AND SERVICE OFFERINGS 表 vs xlsx Segment History／Products & Services History）
for sid, blk, lab in (("ai", "Agents and Infra", "Agents and Infra"), ("dc", "Devices and Consumer", "Devices and Consumer")):
    for k, lab2, xl in (("rev", "Revenue", "Revenue"), ("cogs", "Cost of revenue", "Cost of revenue"), ("opex", "Operating expenses", "Operating expenses"), ("oi", "Operating income", "Operating income")):
        a = pr_find(8, lab2, section=lab)
        b = xl_hist("Segment History", xl if k != "oi" else "Operating income", block=blk)
        add("1_分部與產品線", f"q.seg27.{sid}.{k}.fy27q1", f"季度分部 {lab}：{lab2}", a, "USD M", P, PR, "SEGMENT RESULTS 表（新聞稿第 7 頁附表）", "Interested-party",
            "新聞稿未經查核；10-Q 附註分部表同值", XL, "Segment History 工作表 Q1-27 欄", b)
for pid, lab in (("azure", "Azure"), ("m365cloud", "Microsoft 365 cloud"), ("licensing", "Productivity and server licensing"), ("searchads", "Search and advertising"),
                 ("xbox", "XBOX"), ("industry", "Industry solutions"), ("windows", "Windows OEM and devices"), ("frontier", "Frontier and support services")):
    a = pr_find(9, lab)
    b = xl_hist("Products & Services History", lab)
    add("1_分部與產品線", f"q.prod27.{pid}.fy27q1", f"季度產品線營收：{lab}", a, "USD M", P, PR, "PRODUCT AND SERVICE OFFERINGS REVENUE 表", "Interested-party", "", XL, "Products & Services History 工作表 Q1-27 欄", b)
# 前一年同季（FY26Q1）重編值：與 v0.3 事實總帳（8-K 2026-09-02）對照
for sid, lab in (("ai", "Agents and Infra"), ("dc", "Devices and Consumer")):
    add("1_分部與產品線", f"q.seg27.{sid}.rev.fy26q1.chk", f"FY26Q1 分部營收（本季新聞稿前期欄）：{lab}", pr_find(8, "Revenue", section=lab, col=2), "USD M", "FY26Q1", PR, "SEGMENT RESULTS 表 2025 欄", "Interested-party",
        "與 8-K 2026-09-02 重編值對照", XL, "Segment History Q1-26 欄", xl_hist("Segment History", "Revenue", "Q1-26", block=lab))

# 2. 合併損益（新聞稿 INCOME STATEMENTS vs xlsx Quarterly Income Statements）
for k, lab, xl in (("rev", "Total revenue", "Total revenue"), ("cogs", "Total cost of revenue", "Total cost of revenue"), ("gm", "Gross margin", "Gross margin"),
                   ("rd", "Research and development", "Research and development"), ("sm", "Sales and marketing", "Sales and marketing"), ("ga", "General and administrative", "General and administrative"),
                   ("oi", "Operating income", "Operating income"), ("oie", "Other income (expense), net", "Other income (expense), net"), ("pretax", "Income before income taxes", "Income before income taxes"),
                   ("tax", "Provision for income taxes", "Provision for"), ("ni", "Net income", "Net income")):
    add("2_合併損益", f"q.{k}.fy27q1", f"季度 {lab}", pr_find(5, lab), "USD M", P, PR, "INCOME STATEMENTS 表", "Interested-party", "", XL, f"{QP} 工作表 Q1-27 欄", xl_hist(QP, xl))
add("2_合併損益", "q.epsDil.fy27q1", "季度稀釋 EPS", pr_find(5, "Diluted"), "USD／股", P, PR, "INCOME STATEMENTS 表", "Interested-party", "", XL, f"{QP}（Diluted）", xl_hist(QP, "Diluted"), tol=0.005)
add("2_合併損益", "q.sharesDilW.fy27q1", "季度稀釋加權股數", pr_find(5, "Diluted", nth=1), "百萬股", P, PR, "INCOME STATEMENTS 表", "Interested-party", "", TQ, "10-Q 損益表（Diluted 加權股數）",
    tq_search(r"Diluted\s*\|[^\n]*?\|\s*([\d,]{5})\s*\|", 1))
add("2_合併損益", "q.sharesBasicW.fy27q1", "季度基本加權股數", pr_find(5, "Basic", nth=1), "百萬股", P, PR, "INCOME STATEMENTS 表", "Interested-party")
add("2_合併損益", "q.openaiGain.fy27q1", "OpenAI 投資淨利得（稅前，業外）", 272, "USD M", P, TQ, "10-Q 附註 3（Other income (expense), net 說明）", "Interested-party",
    "10-Q：「$272 million of net gains」；稅後對淨利 +207（新聞稿非 GAAP 調節）", PR, "新聞稿 Non-GAAP Definition：淨利 +207、EPS +0.03（稅後口徑，不同口徑不比）", None, explained="口徑不同（稅前 272 vs 稅後 207），不比值")
add("2_合併損益", "q.oie.exOAI.fy27q1", "業外收支（不含 OpenAI）", -234, "USD M", P, CALL, "CFO 致詞", "Interested-party", "「other income and expense was negative $234 million」", PR + "＋" + TQ, "業外 +38 − OpenAI 稅前利得 272＝−234", 38 - 272)

# 3. 現金流量（新聞稿 vs xlsx Cash Flows）
CF = (("ocf", "Net cash from operations"), ("da", "Depreciation, amortization, and other"), ("sbc", "Stock-based compensation expense"), ("capex", "Additions to property and equipment"),
      ("buyback", "Common stock repurchased"), ("div", "Common stock cash dividends paid"), ("debtRepay", "Repayments of debt"), ("deftax", "Deferred income taxes"), ("invGain", "Net recognized losses (gains)"),
      ("finOther", "Other, net"))
for k, lab in CF:
    add("3_現金流量", f"q.cf.{k}.fy27q1", f"季度現金流量：{lab}", pr_find(7, lab, section="CASH FLOWS"), "USD M", P, PR, "CASH FLOWS STATEMENTS 表", "Interested-party", "", XL, "Cash Flows 工作表 B 欄", xl_col("Cash Flows", lab))
WCL = ("Accounts receivable", "Inventories", "Other current assets", "Other long-term assets", "Accounts payable", "Unearned revenue", "Income taxes", "Other current liabilities", "Other long-term liabilities")
for col, pid in ((1, "fy27q1"), (2, "fy26q1")):
    a = sum(pr_find(7, l, section="CASH FLOWS", col=col) for l in WCL)
    b = sum(xl_col("Cash Flows", l, col=1 + col) for l in WCL)
    add("3_現金流量", f"q.cf.wc.{pid}", "季度營運資產與負債變動合計（九項）", a, "USD M", "FY27Q1" if col == 1 else "FY26Q1", PR, "CASH FLOWS 表 Changes in operating assets and liabilities 九項加總", "Derived",
        "C14：FY2027 剩餘三季的營運資金＝FY2026 全年 − FY26Q1（前一年度同期實績）", XL, "Cash Flows 工作表同九項加總", b)
add("3_現金流量", "q.cf.flPrincipal.fy27q1", "融資租賃本金（融資現金流）", 1000, "USD M", P, TQ, "10-Q 附註 12 租賃補充現金流（Financing cash flows from finance leases）", "Interested-party",
    "10-Q 未經查核", TQ, "同表（Operating cash flows from finance leases 769 為利息）", tq_search(r"Financing cash flows from finance leases\s*\|[^\d]*([\d,]+)"))

# 4. 資本支出（公司口徑，含融資租賃）
cx = capex_q127()
add("4_資本支出", "capexFL.fy27q1", "資本支出（含融資租賃取得資產；公司口徑）", cx["Q1-27"] * 1000, "USD M", P, XL, "CapEx 工作表 Q1-27", "Interested-party", "", CALL, "CFO：「Capital expenditures were $45.2 billion」", 45200)
add("4_資本支出", "capexFL.finLease.fy27q1", "其中融資租賃", 7400, "USD M", P, CALL, "CFO：「finance leases were $7.4 billion」", "Interested-party",
    "10-Q 附註 12 以融資租賃取得的使用權資產 5,192（不同口徑：公司資本支出口徑在起租時計全額本金）", TQ, "10-Q 附註 12 Right-of-use assets obtained in exchange for lease obligations：Finance leases", 5192,
    explained="口徑不同：法說 7.4（資本支出口徑）vs 10-Q 使用權資產 5.192；取法說值（資本支出口徑），列資料缺口")
add("4_資本支出", "capex.cashPPE.fy27q1", "現金 PP&E", 39367, "USD M", P, PR, "CASH FLOWS 表 Additions to property and equipment", "Interested-party", "", CALL, "CFO：「cash paid for P, P, and E was $39.4 billion」", 39400, tol=50)
add("4_資本支出", "capexFL.fy2026.chk", "FY2026 公司口徑資本支出（本季 xlsx 回查）", cx["Fiscal Year 2026"] * 1000, "USD M", "FY2026", XL, "CapEx 工作表 Fiscal Year 2026", "Interested-party", "", "microsoft/data/microsoft_facts_20261008.json", "capexFL.fy2026", 145300)
for q in ("Q2-25", "Q3-25"):
    add("4_資本支出", f"capexFL.fy25q{q[1]}", f"公司口徑資本支出 {q}（Q 頁八季窗口補值）", cx[q] * 1000, "USD M", q, XL, f"CapEx 工作表 {q}", "Interested-party", "Q 頁窗口 FY25Q2–FY27Q1；facts v0.1a 只收 FY25Q4 起")

# 5. 資產負債與資本結構（新聞稿 vs xlsx Balance Sheets；債務與租賃取 10-Q）
for k, lab in (("cashSti", "Total cash, cash equivalents, and short-term investments"), ("debtCur", "Current portion of long-term debt"), ("debtLT", "Long-term debt"), ("ppe", "Property and equipment"),
               ("equity", "Total stockholders' equity")):
    add("5_資產負債", f"bs.{k}.fy27q1", f"資產負債表：{lab}（2026-09-30）", pr_find(7, lab), "USD M", "2026-09-30", PR, "BALANCE SHEETS 表", "Interested-party", "", XL, "Balance Sheets 工作表 B 欄", xl_col("Balance Sheets", lab))
add("5_資產負債", "sh.out.fy27q1", "期末流通股數（2026-09-30）", 7427, "百萬股", "2026-09-30", PR, "BALANCE SHEETS 表（outstanding 7,427 and 7,427）", "Interested-party", "", TQ, "10-Q 資產負債表股數",
    tq_search(r"outstanding\s+([\d,]{5})\s+and\s+[\d,]{5}"))
add("5_資產負債", "debt.face.fy27q1", "長期債務面額合計（含一年內到期）", tq_search(r"maturities of our long-term debt, including the current portion, as of September 30, 2026:.*?Total\s+\|[^\d]*([\d,]+)"), "USD M", "2026-09-30", TQ, "10-Q 附註 9 到期表 Total", "Interested-party",
    "面額；帳面值 4,391＋31,090＝35,481", "microsoft/data/microsoft_facts_20261008.json＋" + PR, "FY2026 面額 46,136 − 本季償還 4,850＝41,286（差 17 推定匯率）", 46136 - 4850, tol=30)
DM = {2027: 4400, 2028: 0, 2029: 1988, 2030: 0, 2031: 500}
for y, v in DM.items():
    add("5_資產負債", f"debt.maturity.fy{y}.q1", f"債務到期 FY{y}" + ("（剩餘三季）" if y == 2027 else ""), v, "USD M", "2026-09-30", TQ, "10-Q 附註 9 到期表", "Interested-party")
add("5_資產負債", "debt.maturity.after.q1", "債務到期 FY2032 以後合計", 34381, "USD M", "2026-09-30", TQ, "10-Q 附註 9 到期表 Thereafter", "Interested-party")
add("5_資產負債", "lease.finLiab.fy27q1", "融資租賃負債合計（2026-09-30）", 70973, "USD M", "2026-09-30", TQ, "10-Q 附註 12 補充資產負債資訊（Total finance lease liabilities）", "Interested-party",
    "", TQ, "10-Q 附註 12 到期表：付款 95,246 − 設算利息 24,273", 95246 - 24273)
FLM = {2027: 6006, 2028: 8163, 2029: 7433, 2030: 7323, 2031: 7300}
for y, v in FLM.items():
    add("5_資產負債", f"lease.maturity.fin.fy{y}.q1", f"融資租賃付款 FY{y}（未折現）" + ("（剩餘三季）" if y == 2027 else ""), v, "USD M", "2026-09-30", TQ, "10-Q 附註 12 到期表", "Interested-party")
add("5_資產負債", "lease.notCommenced.fy27q1", "尚未起租的租賃（主要為資料中心）", 364600, "USD M", "2026-09-30", TQ, "10-Q 附註 12（not yet commenced of $364.6 billion）", "Interested-party", "FY2026 年底 329.1；FY2027–FY2034 起租")

# 6. 股東回饋（IR xlsx Return of Cash vs 新聞稿文字）
add("6_股東回饋", "roc.div.fy27q1", "季度股利（宣告口徑）", roc("Total Dividends") * 1000, "USD M", P, XL, "Return of Cash to Shareholders 工作表 Q1-27", "Interested-party", "現金流量表付現 6,757（付上季宣告的 0.91）")
add("6_股東回饋", "roc.bb.fy27q1", "季度回購（計畫內）", roc("Share Buyback") * 1000, "USD M", P, XL, "Return of Cash to Shareholders 工作表 Q1-27", "Interested-party", "另扣稅買回 1.5（10-Q 附註 14）", TQ, "10-Q：「6 million shares … for $2.9 billion」", 2900, tol=60)
add("6_股東回饋", "roc.total.fy27q1", "股利＋回購合計", (roc("Total Dividends") + roc("Share Buyback")) * 1000, "USD M", P, XL, "同上加總", "Interested-party", "", PR, "新聞稿：「returned $10.2 billion」", 10200, tol=60)
add("6_股東回饋", "roc.dps.fy27q1", "每股季度股利（FY27 第一次宣告）", roc("Dividend Per Share"), "USD／股", P, XL, "Return of Cash 工作表", "Interested-party", "", "microsoft/data/microsoft_facts_20261008.json", "sh.dpsQ.fy27q1", 0.98, tol=0.001)
add("6_股東回饋", "bb.authRemain.fy27q1", "回購授權剩餘", 37700, "USD M", "2026-09-30", TQ, "10-Q 附註 14", "Interested-party")

# 7. 指標、RPO、AI 揭露
mt = wb["Metrics"]
mhdr = {str(mt.cell(2, c).value).strip(): c for c in range(1, mt.max_column + 1) if mt.cell(2, c).value is not None}


def metric(label):
    for r in range(1, mt.max_row + 1):
        if str(mt.cell(r, 3).value or "").strip().startswith(label):
            return mt.cell(r, mhdr["Q1'27"]).value
    raise KeyError(label)


add("7_指標與AI", "rpo.commercial.fy27q1", "商用 RPO", metric("Commercial remaining performance obligation") * 1000, "USD M", "2026-09-30", XL, "Metrics 工作表 Q1'27", "Interested-party", "年增 73%；加權存續 2.3 年、約 30% 於 12 個月內認列（法說）",
    TQ, "10-Q 附註 11：commercial portion $677 billion", 677000)
add("7_指標與AI", "rpo.total.fy27q1", "全公司 RPO", 683000, "USD M", "2026-09-30", TQ, "10-Q 附註 11", "Interested-party")
add("7_指標與AI", "msCloud.rev.fy27q1", "Microsoft Cloud 營收", metric("Microsoft Cloud revenue") * 1000, "USD M", P, XL, "Metrics 工作表", "Interested-party", "年增 30%（cc 31%）", PR, "新聞稿：「Microsoft Cloud revenue was $63.9 billion」", 63900)
add("7_指標與AI", "msCloud.gm.fy27q1", "Microsoft Cloud 毛利率", metric("Microsoft Cloud gross margin percentage"), "比率", P, XL, "Metrics 工作表", "Interested-party", "", CALL, "CFO：「declined year-over-year to 66%」", 0.66, tol=0.005)
add("7_指標與AI", "q.azure27.growth.fy27q1", "Azure 年增（新口徑）", 0.48, "比率", P, PR, "新聞稿：「Azure revenue increased 48%」", "Interested-party", "cc 48%", XL, "Metrics 工作表（'48% / 48%'）", 0.48, tol=0.001)
add("7_指標與AI", "q.azure27.growth.calc.fy27q1", "Azure 年增（金額計算）", 33165 / 22384 - 1, "比率", P, PR, "產品線表 33,165 ÷ 22,384 − 1", "Derived")
add("7_指標與AI", "m365.seatGrowth.fy27q1", "M365 商用付費座席年增", metric("Microsoft 365 commercial seat growth"), "比率", P, XL, "Metrics 工作表", "Interested-party", "", CALL, "CFO：「Paid M365 commercial seats grew 5%」", 0.05, tol=0.001)
add("7_指標與AI", "oai.stake.fy27q1", "OpenAI 持股（as-converted）", 0.24, "比率（約）", "2026-09-30", TQ, "10-Q 附註 1（approximate 24% interest on an as-converted basis）", "Interested-party",
    "FY2026 10-K 約 25%；單一來源（10-Q）")
add("7_指標與AI", "oai.rev.fy27q1", "來自 OpenAI 的營收（含分潤）", 10000, "USD M", P, TQ, "10-Q 附註 1（關係人揭露）", "Interested-party", "應收 OpenAI 7.2（6-30 為 6.0）；只作驗證列，不回推 k（C3）")
add("7_指標與AI", "oai.funded.fy27q1", "OpenAI 投資承諾／已撥付", 12000, "USD M", "2026-09-30", TQ, "10-Q 附註 1（commitments $13.0 billion, $12.0 billion funded）", "Interested-party")
for id_, it, v, u, n in (("call.foundry.fy27q1", "Foundry 營收年增（倍）", 3, "倍（>）", "「revenue more than tripling year over year」；客戶 >120,000"),
                         ("call.copilotDau.fy27q1", "M365 Copilot 日活躍用戶季增", 0.5, "比率（約）", "「Daily users are up nearly 50 percent quarter over quarter」；付費座席本季未揭露"),
                         ("call.newDC.fy27q1", "本季新增資料中心", 24, "座", "「adding 24 new data centers across five continents」"),
                         ("call.shortLived.fy27q1", "短期資產（CPU／GPU）占資本支出", 0.75, "比率（約）", "「Roughly three quarters of our capex was for short-lived assets」"),
                         ("call.maia200.fy27q1", "Maia 200 每瓦 token（相對）", 1.7, "倍（至多）", "「up to 1.7 times more tokens per watt」（Teams、Copilot）；已安裝首批 Vera Rubin"),
                         ("call.fcf.fy27q1", "自由現金流（公司說法）", 18600, "USD M", "「free cash flow was $18.6 billion」＝營業現金流 58.0 − 現金 PP&E 39.4"),
                         ("call.capacity.fy27q1", "產能限制", None, "文字", "「Customer demand continues to exceed available supply」（Q2 指引段）")):
    add("7_指標與AI", id_, it, v, u, P, CALL, "法說", "Interested-party", n)

# 8. 指引（展望簡報 pptx 第 3、4 頁〔圖片，逐頁看圖〕vs 法說 CFO 指引段）
G = "FY27Q2（2026-10～12）"
for id_, it, v, u, n in (("guid.fy27q2.rev", "FY27Q2 營收", [96000, 97100], "USD M", "年增 18–19%"), ("guid.fy27q2.rev.ai", "FY27Q2 Agents and Infra 營收", [80100, 80700], "USD M", "年增 24–25%"),
                         ("guid.fy27q2.rev.dc", "FY27Q2 Devices and Consumer 營收", [15900, 16400], "USD M", ""), ("guid.fy27q2.cogs", "FY27Q2 營業成本", [32700, 32900], "USD M", "含自願離職方案估計"),
                         ("guid.fy27q2.opex", "FY27Q2 營業費用", [18500, 18600], "USD M", "年增 9%"), ("guid.fy27q2.oie", "FY27Q2 業外（不含 OpenAI）", -300, "USD M", "含資料中心融資租賃利息"),
                         ("guid.fy27q2.etr", "FY27Q2 有效稅率", 0.20, "比率", ""), ("guid.fy27q2.capex", "FY27Q2 資本支出（下限）", 50000, "USD M（下限）", "over $50 billion"),
                         ("guid.fy27q2.azure", "FY27Q2 Azure 年增（cc）", [0.49, 0.50], "比率", "部分第三方 GPU 服務合約延到下半年"),
                         ("guid.fy27q2.m365c", "FY27Q2 M365 商用雲年增（cc）", 0.19, "比率（約）", "匯率使報告成長少不到 1pt"),
                         ("guid.fy27q2.m365cons", "FY27Q2 M365 消費雲年增", "low-teens", "文字", ""), ("guid.fy27q2.lic", "FY27Q2 Productivity and server licensing", "decline mid- to high-single digits", "文字", ""),
                         ("guid.fy27q2.ind", "FY27Q2 Industry solutions cloud", "growth mid- to high-single digits", "文字", ""), ("guid.fy27q2.fro", "FY27Q2 Frontier and support services", "growth low- to mid-single digits", "文字", ""),
                         ("guid.fy27q2.srch", "FY27Q2 Search and advertising ex TAC", "growth low-double digits", "文字", "整體搜尋與廣告營收成長相近"), ("guid.fy27q2.xbox", "FY27Q2 XBOX content and services", "growth mid- to high-single digits", "文字", "硬體年減"),
                         ("guid.fy27q2.win", "FY27Q2 Windows OEM and devices", "decline mid-thirties", "文字", "法說：此趨勢延續本財年其餘各季"), ("guid.fy27q2.opm", "FY27Q2 營業利益率", "down slightly year-over-year", "文字", ""),
                         ("guid.fy27q2.msCloudGM", "FY27Q2 Microsoft Cloud 毛利率", 0.65, "比率（約）", "")):
    add("8_指引", id_, it, v, u, G, PP, "展望簡報第 3–4 頁（圖片）", "Interested-party", n + "；法說 CFO 指引段同值（兩讀一致）")
add("8_指引", "guid.cy26.capex", "曆年 2026 資本支出（公司口徑）", 170000, "USD M（約）", "CY2026", CALL, "CFO：「for the calendar year, spend should be approximately $170 billion due to the 3P capacity delays」", "Interested-party",
    "前次（FY26Q4 法說）約 175；減少原因為第三方產能交付延到 FY27 下半年", "microsoft/data/microsoft_facts_20261008.json＋本表", "FY26Q3 31.9＋FY26Q4 41.0＋FY27Q1 45.2＋FY27Q2 隱含 51.9＝170.0（Q2 指引 over 50 吻合）", 31900 + 41000 + 45200 + 51900)
add("8_指引", "guid.fy27.h2capex", "FY27 下半年資本支出", "accelerate", "文字", "FY27H2", CALL, "CFO：「In the second half of the fiscal year, we currently expect CapEx spend to accelerate」", "Interested-party", "主要為 CPU（含記憶體與儲存漲價）、GPU 及延後的第三方產能")
add("8_指引", "guid.fy27.fcf", "FY27 自由現金流", "positive", "文字", "FY2027", CALL, "CFO：「remain free cash flow positive in FY27」", "Interested-party", "另稱會在適當時機取得額外資金來源")
add("8_指引", "guid.fy27.revAccel.m365", "M365 商用雲全年", "accelerate through the year", "文字", "FY2027", CALL, "CFO", "Interested-party")
add("8_指引", "capex.fy27.derived.q1", "FY2027 資本支出推估（公司口徑）", 45200 + 51900 + (102100 + 5000), "USD M", "FY2027", "本表推導", "上半年＝Q1 45.2＋Q2 隱含 51.9＝97.1；下半年＝前次推估 102.1＋延後的第三方產能約 5.0（曆年 175 → 170）＝107.1", "Derived",
    "低＝下半年同上半年 97.1（合計 194.2）；高＝下半年較上半年加速 20%（116.5，合計 213.6）。前次推估（v0.3）204.2")

groups = {}
for g, it in items:
    groups.setdefault(g, []).append(it)
outj = {"_readme": "Microsoft FY27Q1 事實總帳（MAG v0.3.1 季度滾動，2026-10-29）。每筆附來源檔名＋表名／頁、標記；check＝第二次獨立讀取（不同文件或同文件不同表）的值與是否一致。金額 USD M。",
        "asOf": "2026-10-29", "company": "Microsoft", "ticker": "MSFT", "latestFiled": "FY27Q1（2026-09-30；10-Q 2026-10-28）",
        "sources": {"A": PR, "B": XL, "C": TQ, "D": PP, "E": CALL},
        "mismatches": mism, "groups": groups}
json.dump(outj, open(OUT, "w", encoding="utf-8"), ensure_ascii=False, indent=1)
n_chk = sum(1 for _, it in items if "check" in it)
bad = [m for m in mism if not m["explained"]]
print(f"寫出 {OUT}：{len(items)} 筆，雙讀核對 {n_chk} 筆，不一致 {len(mism)} 筆（未解釋 {len(bad)}）")
for m in mism:
    print("  不一致：", m)
sys.exit(1 if bad else 0)
