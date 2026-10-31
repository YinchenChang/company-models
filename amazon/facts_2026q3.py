# Amazon 2026Q3 事實總帳（MAG v0.3.1 工作單第 1 節第 1 項）：由 mag/extract/docs/2026Q3/ 的財報新聞稿逐列轉錄讀取，
# 每筆附檔名、表名、標記；同一數字由兩次獨立讀取核對一致，不一致列出並以非零結束。
# 兩次讀取：(1) 新聞稿 PDF 轉錄的不同表互對（分部表 vs Supplemental 季別表、損益表 vs Supplemental、營收線合計 vs 合併、
#           現金流量表 9M − 事實總帳 Q1＋Q2、TTM 對帳）；(2) PDF 轉錄 vs IR 網頁第二次讀取（amazon_2026Q3_核對_新聞稿PDF對IR網頁.md 22 項）；
#           (3) Supplemental 表中 2025Q3–2026Q2 各季 vs 既有事實總帳（data/amazon_facts_20261008.json；無重編則一致）。
# 用法：python3 amazon/facts_2026q3.py  → 寫 amazon/data/amazon_facts_2026q3.json；不一致時非零結束。
import json, os, re, sys

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
D = os.path.join(ROOT, "mag", "extract", "docs", "2026Q3")
PRF = "mag/extract/docs/2026Q3/amazon_2026Q3_新聞稿.md"
CKF = "mag/extract/docs/2026Q3/amazon_2026Q3_核對_新聞稿PDF對IR網頁.md"
URL = "https://s2.q4cdn.com/299287126/files/doc_earnings/2026/q3/earnings-result/AMZN-Q3-2026-Earnings-Release.pdf"
PR = open(os.path.join(D, "amazon_2026Q3_新聞稿.md"), encoding="utf-8").read().splitlines()
CK = open(os.path.join(D, "amazon_2026Q3_核對_新聞稿PDF對IR網頁.md"), encoding="utf-8").read().splitlines()
OLD = json.load(open(os.path.join(HERE, "data", "amazon_facts_20261008.json"), encoding="utf-8"))
F0 = {f["id"]: f for g in OLD["groups"].values() for f in g}


def num(s):
    s = s.strip().replace("$", "").replace(",", "").replace(" ", "")
    if s in ("—", "-", "", "N/A"): return 0.0 if s in ("—", "-") else None
    neg = s.startswith("(") and s.endswith(")")
    s = s.strip("()").rstrip("%")
    try:
        v = float(s)
    except ValueError:
        return None
    return -v if neg else v


def tables():
    """回傳 {段落標題: {列名: [數字...]}}；段落＝最近的 ## 或 ### 標題（### 之下再以「X Segment:」細分）"""
    out, cur = {}, None
    for ln in PR:
        if ln.startswith("## ") or ln.startswith("### "):
            cur = ln.lstrip("# ").strip(); out.setdefault(cur, {}); continue
        m = re.match(r"^([A-Za-z ]+?) Segment:\s*$", ln)
        if m and cur:
            cur = cur.split("｜")[0] + "｜" + m.group(1); out.setdefault(cur, {}); continue
        if cur and ln.startswith("|") and not ln.startswith("|---"):
            cells = [c.strip() for c in ln.strip().strip("|").split("|")]
            lab = cells[0]
            vals = [num(c) for c in cells[1:]]
            n = 2
            base = lab
            while lab in out[cur]:
                lab = f"{base}#{n}"; n += 1
            out[cur][lab] = vals
    return out


T = tables()


def sect(prefix):
    ks = [k for k in T if k.startswith(prefix)]
    if not ks: raise SystemExit("找不到段落：" + prefix)
    return T[ks[0]]


CF = sect("三、Consolidated Statements of Cash Flows")      # 欄：3M25 3M26 9M25 9M26 TTM25 TTM26
IS = sect("四、Consolidated Statements of Operations")      # 欄：3M25 3M26 9M25 9M26
SG = sect("六、Segment Information")
BS = sect("七、Consolidated Balance Sheets")                # 欄：2025-12-31 2026-09-30
S81 = sect("8.1 Cash Flows and Shares")                     # 欄：Q3 25 Q4 25 Q1 26 Q2 26 Q3 26 YoY
S82 = sect("8.2 Results of Operations")
S8NA, S8IN, S8AW = sect("8.3 Segments｜North"), sect("8.3 Segments｜International"), sect("8.3 Segments｜AWS")
S84 = sect("8.4 Net Sales")
S85 = sect("8.5 Stock-based")

ITEMS, CHK = [], []
Q = "2026q3"


def add(iid, item, value, unit="USD M", src="", tag="[Interested-party]", note="", period=None):
    ITEMS.append({"id": iid, "item": item, "value": value, "unit": unit, "period": period or ("2026Q3" if iid.endswith(Q) else iid.split(".")[-1]),
                  "source": src, "url": URL, "docDate": "2026-10-29", "retrieved": "2026-10-31", "tag": tag, "note": note})


def chk(label, a, b, src, tol=0.5):
    ok = a is not None and b is not None and abs(a - b) <= tol
    CHK.append({"id": label, "value": a, "second": b, "source": src, "ok": ok})


def seg_rows(name, tbl):
    """分部表：Net sales／Operating expenses／Operating income 依序出現於北美、國際、AWS、合併"""
    labs = list(tbl.keys())
    i0 = labs.index(name)
    return {k.split("#")[0]: tbl[k] for k in labs[i0 + 1:i0 + 4]}


SRC_SEG = f"{PRF} 第六節 Segment Information（三個月 2026 欄）"
SRC_SUP = f"{PRF} 第八節 Supplemental（Q3 2026 欄）"
for code, nm, sup in (("NA", "North America", S8NA), ("INTL", "International", S8IN), ("AWS", "AWS", S8AW)):
    rows = seg_rows(nm, SG)
    for key, lab in (("rev", "Net sales"), ("oi", "Operating income")):
        v = rows[lab][1]
        add(f"is.{code}.{key}.{Q}", f"{nm} 分部{'營收' if key == 'rev' else '營業利益'}", v, src=SRC_SEG)
        chk(f"is.{code}.{key}.{Q}（分部表 vs Supplemental）", v, sup[lab][4], SRC_SUP)
        add(f"is.{code}.{key}.2026m9", f"{nm} 分部{'營收' if key == 'rev' else '營業利益'}（9M26）", rows[lab][3], src=f"{PRF} 第六節（九個月 2026 欄）", period="9M26")
        add(f"is.{code}.{key}.2025m9", f"{nm} 分部{'營收' if key == 'rev' else '營業利益'}（9M25）", rows[lab][2], src=f"{PRF} 第六節（九個月 2025 欄）", period="9M25")
        # 9M26 − 事實總帳 Q1＋Q2＝Q3（第二讀）
        h1 = F0[f"is.{code}.{key}.2026q1"]["value"] + F0[f"is.{code}.{key}.2026q2"]["value"]
        chk(f"is.{code}.{key}.{Q}（9M26 − 事實總帳 Q1＋Q2）", v, rows[lab][3] - h1, "分部表九個月欄 − amazon_facts_20261008.json")
        # Supplemental 前四季 vs 既有事實總帳（無重編）
        for j, per in enumerate(("2025q3", "2025q4", "2026q1", "2026q2")):
            chk(f"is.{code}.{key}.{per}（Supplemental vs 既有總帳）", sup[lab][j], F0[f"is.{code}.{key}.{per}"]["value"], "amazon_facts_20261008.json")
cons = seg_rows("Consolidated", SG)
add(f"is.TOT.rev.{Q}", "合併營收", IS["Total net sales"][1], src=f"{PRF} 第四節損益表（三個月 2026 欄）")
add(f"is.TOT.oi.{Q}", "合併營業利益", IS["Operating income"][1], src=f"{PRF} 第四節損益表")
chk("is.TOT.rev.2026q3（損益表 vs Supplemental 8.2）", IS["Total net sales"][1], S82["Worldwide (WW) net sales"][4], SRC_SUP)
chk("is.TOT.rev.2026q3（損益表 vs 分部表合併）", IS["Total net sales"][1], cons["Net sales"][1], SRC_SEG)
chk("is.TOT.oi.2026q3（損益表 vs Supplemental 8.2）", IS["Operating income"][1], S82["Operating income"][4], SRC_SUP)
chk("is.TOT.rev.2026q3（三分部合計）", IS["Total net sales"][1], sum(seg_rows(n, SG)["Net sales"][1] for n in ("North America", "International", "AWS")), SRC_SEG)
chk("is.TOT.oi.2026q3（三分部合計）", IS["Operating income"][1], sum(seg_rows(n, SG)["Operating income"][1] for n in ("North America", "International", "AWS")), SRC_SEG)
for j, per in enumerate(("2025q3", "2025q4", "2026q1", "2026q2")):
    chk(f"is.TOT.rev.{per}（Supplemental vs 既有總帳）", S82["Worldwide (WW) net sales"][j], F0[f"is.TOT.rev.{per}"]["value"], "amazon_facts_20261008.json")
add("is.TOT.rev.2026m9", "合併營收（9M26）", IS["Total net sales"][3], src=f"{PRF} 第四節（九個月欄）", period="9M26")
add("is.TOT.oi.2026m9", "合併營業利益（9M26）", IS["Operating income"][3], src=f"{PRF} 第四節（九個月欄）", period="9M26")
# 營收線（Supplemental 8.4）
LN = (("online", "Online stores (1)"), ("physical", "Physical stores (2)"), ("seller3p", "Third-party seller services (3)"), ("ads", "Advertising services (4)"),
      ("subs", "Subscription services (5)"), ("other", "Other (6)"))
for k, lab in LN:
    add(f"is.cat.{k}.{Q}", f"營收線：{lab.split(' (')[0]}", S84[lab][4], src=f"{PRF} 第八節 8.4 Net Sales（Q3 2026 欄）")
    for j, per in enumerate(("2025q3", "2025q4", "2026q1", "2026q2")):
        chk(f"is.cat.{k}.{per}（Supplemental vs 既有總帳）", S84[lab][j], F0[f"is.cat.{k}.{per}"]["value"], "amazon_facts_20261008.json")
chk("營收線合計（含 AWS）＝合併營收（2026Q3）", sum(S84[lab][4] for _, lab in LN) + S84["AWS"][4], IS["Total net sales"][1], "8.4 各線加總 vs 損益表")
chk("營收線：AWS（8.4）＝AWS 分部（2026Q3）", S84["AWS"][4], seg_rows("AWS", SG)["Net sales"][1], SRC_SEG)
# 損益表其他列
for fid, lab, item in (("is.intExp", "Interest expense", "利息費用"), ("is.intInc", "Interest income", "利息收入"), ("is.otherInc", "Other income (expense), net", "其他業外收益"),
                       ("is.pretax", "Income before income taxes", "稅前淨利"), ("is.tax", "Provision for income taxes", "所得稅"), ("is.ni", "Net income", "淨利"),
                       ("is.eqMethod", "Equity-method investment activity, net of tax", "權益法投資損益（稅後）")):
    v3, v9 = IS[lab][1], IS[lab][3]
    sgn = -1 if fid in ("is.intExp", "is.tax") else 1
    add(f"{fid}.{Q}", item, sgn * v3, src=f"{PRF} 第四節損益表（三個月 2026 欄；費用以正值記）")
    add(f"{fid}.2026m9", item + "（9M26）", sgn * v9, src=f"{PRF} 第四節（九個月欄）", period="9M26")
    if fid in ("is.intExp", "is.pretax", "is.tax", "is.ni"):
        h1 = F0[f"{fid}.2026q1"]["value"] + F0[f"{fid}.2026q2"]["value"]
        chk(f"{fid}.{Q}（9M26 − 事實總帳 Q1＋Q2）", sgn * v3, sgn * v9 - h1, "損益表九個月欄 − amazon_facts_20261008.json")
chk("is.ni.2026q3（損益表 vs 現金流量表）", IS["Net income"][1], CF["Net income"][1], f"{PRF} 第三節現金流量表")
chk("is.ni.2026q3（損益表 vs Supplemental 8.2）", IS["Net income"][1], S82["Net income"][4], SRC_SUP)
chk("is.ni.2026q3（稅前 − 稅＋權益法）", IS["Net income"][1], IS["Income before income taxes"][1] + IS["Provision for income taxes"][1] + IS["Equity-method investment activity, net of tax"][1], "損益表三列")
add(f"sh.wDiluted.{Q}", "稀釋加權股數（三個月）", IS["Diluted"][1], unit="M 股", src=f"{PRF} 第四節損益表")
add("sh.wDiluted.2026m9", "稀釋加權股數（九個月）", IS["Diluted"][3], unit="M 股", src=f"{PRF} 第四節損益表", period="9M26")
add("sh.wDiluted.2025m9", "稀釋加權股數（2025 年九個月）", IS["Diluted"][2], unit="M 股", src=f"{PRF} 第四節損益表", period="9M25")
chk("is.eps.2026q3（淨利 ÷ 稀釋股數 vs 公司 1.89）", round(IS["Net income"][1] / IS["Diluted"][1], 2), IS["Diluted earnings per share"][1], "損益表 EPS 列", tol=0.005)
# 現金流量表
CFL = (("cf.cfo", "Net cash provided by (used in) operating activities", "營運現金流"), ("cf.capexGross", "Purchases of property and equipment", "購置不動產及設備（總額）"),
       ("cf.ppeProceeds", "Proceeds from property and equipment sales and incentives", "不動產及設備出售與獎勵收入"),
       ("cf.da", "Depreciation and amortization of property and equipment and capitalized content costs, operating lease assets, and other", "折舊攤銷（現金流量表口徑）"),
       ("cf.sbc", "Stock-based compensation", "股份基礎給付"), ("cf.finLeasePrin", "Principal repayments of finance leases", "融資租賃本金償還"),
       ("cf.finOblPrin", "Principal repayments of financing obligations", "融資義務本金償還"),
       ("cf.acqInvest", "Acquisitions, net of cash acquired, non-marketable investments, and other, net", "收購與非上市投資淨額"),
       ("cf.ltdProceeds", "Proceeds from long-term debt", "長期借款收入"), ("cf.ltdRepay", "Repayments of long-term debt", "長期借款償還"),
       ("cf.stdProceeds", "Proceeds from short-term debt, and other", "短期借款收入"), ("cf.stdRepay", "Repayments of short-term debt, and other", "短期借款償還"),
       ("cf.finLeasePPE", "Property and equipment acquired under finance leases, net of remeasurements and modifications", "融資租賃取得之不動產及設備"),
       ("cf.ppeUnpaid", "Increase (decrease) in property and equipment acquired but not yet paid", "已取得未付款之不動產及設備增加"),
       ("cf.nonop", "Non-operating expense (income), net", "業外（收益）費用淨額調整"), ("cf.defTax", "Deferred income taxes", "遞延所得稅"))
CFV = {}
for fid, lab, item in CFL:
    v = CF[lab]
    sgn = -1 if fid in ("cf.capexGross", "cf.finLeasePrin", "cf.finOblPrin", "cf.acqInvest", "cf.ltdRepay", "cf.stdRepay") else 1
    CFV[fid] = [sgn * x for x in v]
    add(f"{fid}.{Q}", item, sgn * v[1], src=f"{PRF} 第三節現金流量表（三個月 2026 欄；流出以正值記）")
    add(f"{fid}.2026m9", item + "（9M26）", sgn * v[3], src=f"{PRF} 第三節（九個月 2026 欄）", period="9M26")
    add(f"{fid}.2025m9", item + "（9M25）", sgn * v[2], src=f"{PRF} 第三節（九個月 2025 欄）", period="9M25")
    add(f"{fid}.ttm{Q}", item + "（TTM 至 2026-09-30）", sgn * v[5], src=f"{PRF} 第三節（十二個月 2026 欄）", period="TTM")
    if f"{fid}.2026q1" in F0 and f"{fid}.2026q2" in F0:
        h1 = F0[f"{fid}.2026q1"]["value"] + F0[f"{fid}.2026q2"]["value"]
        chk(f"{fid}.{Q}（9M26 − 事實總帳 Q1＋Q2）", sgn * v[1], sgn * v[3] - h1, "現金流量表九個月欄 − amazon_facts_20261008.json")
cxn = CFV["cf.capexGross"][1] - CFV["cf.ppeProceeds"][1]
add(f"cf.capexNet.{Q}", "資本支出淨額（購置 − 出售及獎勵；公司「cash CapEx」）", cxn, src="由同表兩列換算", tag="[Derived]")
add("cf.capexNet.2026m9", "資本支出淨額（9M26）", CFV["cf.capexGross"][3] - CFV["cf.ppeProceeds"][3], src="由同表兩列換算", tag="[Derived]", period="9M26")
add("cf.capexNet.2025m9", "資本支出淨額（9M25）", CFV["cf.capexGross"][2] - CFV["cf.ppeProceeds"][2], src="由同表兩列換算", tag="[Derived]", period="9M25")
chk("cf.capexNet.2026q3（換算 vs 法說「cash CapEx $75.9B」）", round(cxn / 100) * 100, 75900, "法說逐字稿（不入 repo；只引數字）", tol=1)
chk("cf.capexNet.ttm2026q3（換算 vs Supplemental 8.1）", CFV["cf.capexGross"][5] - CFV["cf.ppeProceeds"][5], S81["Purchases of property and equipment, net of proceeds from sales and incentives -- TTM"][4], f"{PRF} 8.1")
chk("cf.fcf.ttm2026q3（OCF − 資本支出淨額 vs Supplemental FCF TTM）", CFV["cf.cfo"][5] - (CFV["cf.capexGross"][5] - CFV["cf.ppeProceeds"][5]), S81["Free cash flow -- TTM (1)"][4], f"{PRF} 8.1")
chk("cf.cfo.ttm2026q3（現金流量表 vs Supplemental 8.1）", CFV["cf.cfo"][5], S81["Operating cash flow -- trailing twelve months (TTM)"][4], f"{PRF} 8.1")
chk("cf.sbc.2026q3（現金流量表 vs Supplemental 8.5）", CFV["cf.sbc"][1], S85["Total stock-based compensation expense"][4], f"{PRF} 8.5")
add(f"cf.fcf.{Q}", "自由現金流（公司口徑＝營運現金流 − 資本支出淨額）", CFV["cf.cfo"][1] - cxn, src="換算", tag="[Derived]")
# 資產負債表
for fid, lab, item in (("bs.cash", "Cash and cash equivalents", "現金及約當現金"), ("bs.mktSec", "Marketable securities", "有價證券"),
                       ("bs.ltd", "Long-term debt", "長期債務（帳面，不含流動部分）"), ("bs.ppe", "Property and equipment, net", "不動產及設備淨額"),
                       ("bs.ltLease", "Long-term lease liabilities", "長期租賃負債"), ("bs.equity", "Total stockholders' equity", "股東權益"),
                       ("bs.otherAssets", "Other assets", "其他資產（含 Anthropic 等投資）"), ("bs.totalAssets", "Total assets", "總資產")):
    add(f"{fid}.{Q}", item + "（2026-09-30）", BS[lab][1], src=f"{PRF} 第七節資產負債表")
    if f"{fid}.fy2025" in F0:
        chk(f"{fid}.fy2025（Q3 新聞稿比較欄 vs 既有總帳）", BS[lab][0], F0[f"{fid}.fy2025"]["value"], "amazon_facts_20261008.json")
add(f"bs.cashMs.{Q}", "現金、約當現金與有價證券（2026-09-30）", BS["Cash and cash equivalents"][1] + BS["Marketable securities"][1], src="資產負債表兩列相加", tag="[Derived]")
chk("bs.totalAssets（資產各列加總）", BS["Total assets"][1], sum(BS[k][1] for k in ("Total current assets", "Property and equipment, net", "Operating leases", "Goodwill", "Other assets")), "資產負債表")
# 股數（資產負債表股本說明 vs Supplemental）
stl = [k for k in BS if k.startswith("Common stock")][0]
m = re.search(r"11,246 and 11,326 shares issued; 10,731 and ([\d,]+) shares outstanding", stl)
sh_out = float(m.group(1).replace(",", ""))
add("sh.out.20260930", "流通股數（2026-09-30）", sh_out, unit="M 股", src=f"{PRF} 第七節資產負債表股本說明", note="10-Q 封面股數未取得（10-Q 缺）")
chk("sh.out.20260930（資產負債表 vs Supplemental 8.1）", sh_out, S81["Common shares outstanding"][4], f"{PRF} 8.1")
add(f"sh.awards.{Q}", "未歸屬股份獎勵（2026-09-30）", S81["Stock-based awards outstanding"][4], unit="M 股", src=f"{PRF} 8.1")
chk("sh.awards.2026q2（Supplemental vs 既有總帳）", S81["Stock-based awards outstanding"][3], F0["sh.awards.2026q2"]["value"], "amazon_facts_20261008.json")
# 推導：債務面額（10-Q 未取得）
face = F0["debt.face.2026q2"]["value"] + CFV["cf.ltdProceeds"][1] - CFV["cf.ltdRepay"][1]
add("debt.face.20260930", "債務面額（2026-09-30，推導）", face, src="2026-06-30 債務面額 132,995（10-Q 2026Q2 附註 5）＋Q3 長期借款收入 − 償還", tag="[Derived]",
    note="10-Q 附註未取得：收入為扣除折價後的淨額，外幣債未重估；帳面法（6/30 面額＋長期債帳面變動）為 177,226，差額主要是 2027 年到期票據轉入流動部分（10-Q 待補）")
# 第二讀：IR 網頁核對檔（22 項）
MAP = {"Q3 2026 Total net sales": f"is.TOT.rev.{Q}", "Q3 2026 Operating income": f"is.TOT.oi.{Q}", "Q3 2026 Net income": f"is.ni.{Q}",
       "North America net sales（Q3 2026）": f"is.NA.rev.{Q}", "North America operating income（Q3 2026）": f"is.NA.oi.{Q}",
       "International net sales（Q3 2026）": f"is.INTL.rev.{Q}", "International operating income（Q3 2026）": f"is.INTL.oi.{Q}",
       "AWS net sales（Q3 2026）": f"is.AWS.rev.{Q}", "AWS operating income（Q3 2026）": f"is.AWS.oi.{Q}",
       "Net cash provided by operating activities（Q3 2026，三個月）": f"cf.cfo.{Q}", "Purchases of property and equipment（Q3 2026）": f"cf.capexGross.{Q}",
       "Proceeds from property and equipment sales and incentives（Q3 2026）": f"cf.ppeProceeds.{Q}",
       "Property and equipment acquired under finance leases, net of remeasurements and modifications（Q3 2026）": f"cf.finLeasePPE.{Q}",
       "Cash and cash equivalents（2026-09-30）": f"bs.cash.{Q}", "Long-term debt（2026-09-30）": f"bs.ltd.{Q}",
       "Advertising services net sales（Q3 2026）": f"is.cat.ads.{Q}", "AWS net sales（Net sales by type，Q3 2026）": f"is.cat.aws.{Q}"}
V = {it["id"]: it["value"] for it in ITEMS}
V[f"is.cat.aws.{Q}"] = S84["AWS"][4]
n_ir = 0
for ln in CK:
    if not ln.startswith("|") or ln.startswith("|---") or ln.startswith("| 項目"): continue
    c = [x.strip() for x in ln.strip().strip("|").split("|")]
    if c[0] in MAP:
        ir = abs(num(c[2]))
        chk(f"{MAP[c[0]]}（PDF 轉錄 vs IR 網頁第二讀）", abs(V[MAP[c[0]]]), ir, CKF)
        n_ir += 1
    elif c[0] == "Free cash flow -- TTM（截至 2026-09-30）":
        chk("cf.fcf.ttm2026q3（PDF 轉錄 vs IR 網頁第二讀）", S81["Free cash flow -- TTM (1)"][4], num(c[2]), CKF); n_ir += 1
    elif c[0] == "Q3 2026 Diluted EPS":
        chk("is.eps.2026q3（PDF 轉錄 vs IR 網頁第二讀）", IS["Diluted earnings per share"][1], num(c[2]), CKF, tol=0.001); n_ir += 1
# 指引與法說（敘述段落：短引文；法說逐字稿不入 repo，只引數字與短語）
G = [("guide.2026q4.rev", "公司指引 2026Q4 營收（240.0–247.0；年增 12%–16%；匯率約 −60bp）", {"low": 240000, "high": 247000}, "新聞稿 Financial Guidance（PDF 與 IR 網頁一致）"),
     ("guide.2026q4.oi", "公司指引 2026Q4 營業利益（28.5–33.5；2025Q4 為 25.0）", {"low": 28500, "high": 33500}, "同上；Leo 成本 Q4 仍多數費用化（法說）"),
     ("guide.2026.capex", "2026 現金資本支出指引「approximately $220 billion」（維持）", 220000, "新聞稿 Financial Guidance；法說「still expect 2026 cash CapEx of approximately $220 billion」"),
     ("guide.2027.capex", "2027 資本支出方向：「spend meaningfully more in 2027」（無數字）", None, "法說逐字稿（Jassy）"),
     ("aws.backlog.2026q3", "AWS backlog（2026-09-30）：「$578 billion, growing triple digits」", 578000, "法說逐字稿（Jassy）；Wattonomics Q_review_2026q3 已登錄 578"),
     ("aws.runRate.2026q3", "AWS 年化營收 run-rate", 184000, "新聞稿副標題與法說（184B）"),
     ("aws.growthView.2027_2028", "公司說法：AWS 2027、2028 年各成長 40% 以上（只作驗證列）", 0.40, "新聞稿 CEO 引述；法說"),
     ("aws.aiGrowth.2026q3", "AWS AI 營收年增三位數（未給金額；推論過半為 inference）", None, "法說（CFO）"),
     ("trn.external.2027", "Trainium 對外銷售晶片：2027 年底開始", None, "法說（Jassy）"),
     ("ads.guide.2026q4", "廣告營收 2026Q4 年增「high teens」", None, "新聞稿")]
for iid, item, val, src in G:
    ITEMS.append({"id": iid, "item": item, "value": val, "unit": "USD M" if isinstance(val, (int, dict)) else "—", "period": "2026Q3", "source": src,
                  "url": URL, "docDate": "2026-10-29", "retrieved": "2026-10-31", "tag": "[Interested-party]", "note": ""})
OUT = {"_readme": "Amazon 2026Q3 事實總帳（MAG v0.3.1 工作單第 1 節第 1 項）。由 amazon/facts_2026q3.py 自財報新聞稿逐列轉錄讀取；10-Q 未取得（雲端擋 SEC）："
                  "債務明細、租賃、在建工程、分部折舊、RPO 附註沿用 2026Q2 或推導（見 build_amazon.py 與報告資料缺口）。金額 USD M（費用與流出以正值記）；checks＝第二次獨立讀取。",
       "asOf": "2026-09-30", "items": ITEMS, "checks": CHK}
json.dump(OUT, open(os.path.join(HERE, "data", "amazon_facts_2026q3.json"), "w", encoding="utf-8"), ensure_ascii=False, indent=1)
bad = [c for c in CHK if not c["ok"]]
print(f"事實 {len(ITEMS)} 筆；兩讀核對 {len(CHK)} 項（其中 IR 網頁第二讀 {n_ir} 項），一致 {len(CHK) - len(bad)}，不一致 {len(bad)}")
for c in bad: print("  不一致：", c)
sys.exit(1 if bad else 0)
