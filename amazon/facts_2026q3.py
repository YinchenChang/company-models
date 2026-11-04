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
add("sh.out.20260930", "流通股數（2026-09-30）", sh_out, unit="M 股", src=f"{PRF} 第七節資產負債表股本說明", note="模型 sh_out 用本值（與淨負債、未歸屬獎勵同日 9/30）；10-Q 封面 10/21 股數 sh.cover.20261021 只作驗證")
chk("sh.out.20260930（資產負債表 vs Supplemental 8.1）", sh_out, S81["Common shares outstanding"][4], f"{PRF} 8.1")
add(f"sh.awards.{Q}", "未歸屬股份獎勵（2026-09-30）", S81["Stock-based awards outstanding"][4], unit="M 股", src=f"{PRF} 8.1")
chk("sh.awards.2026q2（Supplemental vs 既有總帳）", S81["Stock-based awards outstanding"][3], F0["sh.awards.2026q2"]["value"], "amazon_facts_20261008.json")
# ---------------- 10-Q 2026Q3（v0.3.2；工作單 mag/docs/workorders/20261104_MAG_10Q補入_2026Q3.md 第 1 節第 1 項） ----------------
# 原文：Form 10-Q（期末 2026-09-30；2026-10-30 提交；accession 0001018724-26-000040）。全文不入 repo（公開 repo），只登錄必要數字與頁碼（PDF 頁＝10-Q 頁尾頁碼）。
# 兩讀：(1) 10-Q 內部互對（面額明細加總、帳面＋流動＋折價＝面額、租賃現值＝總額 − 利息、9M − 既有 Q1＋Q2、分部加總＝合併、承諾表列加總）與 10-Q 對新聞稿；
#       (2) 環境變數 AMZN_10Q_TXT 指向 pdftotext -layout 文字檔時，逐值檢查該數字出現在標註頁（雲端 /home/claude/src10q/amzn_q3.txt；不存在則略過並註明）。
TQ_URL = "https://www.sec.gov/Archives/edgar/data/1018724/000101872426000040/"
TQ = "Form 10-Q 2026Q3（期末 2026-09-30，核閱未查核；2026-10-30 提交，accession 0001018724-26-000040）"
PG = {}          # 事實 id → (頁碼, [應出現在該頁的數字字串])


def addq(iid, item, value, page, note_=None, unit="USD M", period=None, old=None, tag="[Verified]", find=None, note=""):
    nt = note
    if old is not None:
        nt = (nt + "；" if nt else "") + f"原（v0.3.1）：{old}"
    ITEMS.append({"id": iid, "item": item, "value": value, "unit": unit, "period": period or ("2026Q3" if iid.endswith(Q) else iid.split(".")[-1]),
                  "source": f"{TQ} 第 {page} 頁" + (f" {note_}" if note_ else ""), "url": TQ_URL, "docDate": "2026-10-30", "retrieved": "2026-11-04",
                  "tag": tag, "note": nt})
    if find is not None:
        PG[iid] = (page, find)


# 1) 債務（附註 5，第 18–19 頁）
NOTES = [("2014 票據（發行 6.0B）", 2750), ("2017 票據（17.0B）", 12000), ("2020 票據（10.0B）", 7750), ("2021 票據（18.5B）", 12250),
         ("2022-04 票據（12.8B）", 9750), ("2022-12 票據（8.3B）", 5750), ("2025 票據（15.0B）", 15000), ("2026-03 美元票據（37.0B；含浮動 2.8B）", 37000),
         ("2026-03 歐元票據（€14.5B；含浮動 €1.8B）", 16427), ("2026-05 瑞士法郎票據（CHF2.8B）", 3375), ("2026-06 加幣票據（C$14.0B）", 9834),
         ("2026-07 美元票據（25.0B；含浮動 0.75B）", 25000), ("2026-09 英鎊票據（£4.3B）", 5637), ("2026 定期貸款（Term Loan，2026-09 動用）", 17500), ("其他長期債", 864)]
face_old = F0["debt.face.2026q2"]["value"] + CFV["cf.ltdProceeds"][1] - CFV["cf.ltdRepay"][1]
addq("debt.face.20260930", "債務面額（2026-09-30）", 180887, 18, "附註 5 Total face value of long-term debt", period="2026-09-30",
     old=f"{face_old:,.0f}（推導：6/30 面額 132,995＋Q3 長期借款收入 48,102 − 償還 7）", find=["180,887"])
addq("debt.notes.2026q3", "長期債務明細（逐次發行，面額）", [[a, b] for a, b in NOTES], 18, "附註 5", period="2026-09-30",
     note="債券合計 162.5B（另 2026 Term Loan 17.5B、其他 0.864B）；票據加權剩餘年限 13.9 年；無財務承諾條款", find=["16,427", "3,375", "9,834", "5,637", "17,500", "864"])
chk("debt.face.20260930（10-Q 明細加總 vs 合計）", sum(b for _, b in NOTES), 180887, f"{TQ} 第 18 頁")
addq("debt.current.20260930", "長期債務流動部分（2026-09-30）", 6827, 18, "附註 5", period="2026-09-30", find=["6,827"])
addq("debt.discount.20260930", "未攤銷折價與發行成本（2026-09-30）", 935, 18, "附註 5", period="2026-09-30", find=["(935)"])
chk("debt.face.20260930（新聞稿長期債帳面＋10-Q 流動部分＋折價 vs 10-Q 面額）", BS["Long-term debt"][1] + 6827 + 935, 180887, f"{PRF} 第七節；{TQ} 第 18 頁")
addq("debt.notesOut.20260930", "優先票據流通在外（2026-09-30）", 162500, 18, "附註 5", period="2026-09-30", find=["$162.5 billion"])
addq("debt.fv.20260930", "票據公允價值（2026-09-30，Level 2）", 145400, 18, "附註 5", period="2026-09-30", note="面額 162.5B → 價格約 89.5%", find=["$145.4 billion"])
addq("debt.termLoan.20260930", "2026 Term Loan（無擔保、延後單次動用；2026-09 動用；2029-09 到期；SOFR＋0.625%–0.875%）", 17500, 19, "附註 5", period="2026-09-30", find=["$17.5 billion"])
gbp_old = F0["debt.postQ2.gbp202609"]["value"]
q3_gap = CFV["cf.ltdProceeds"][1] - F0["debt.postQ2.usd202607"]["value"] - gbp_old
addq("debt.q3gap.2026q3", "Q3 長期借款收入中未對上的約 18.9 的來源", 17500 + (5637 - gbp_old), 18, "附註 5（Term Loan 17,500＋英鎊票據實際 5,637 − 原估 4,250）",
     tag="[Derived]", old=f"未知 {q3_gap:,.0f}（48,102 − 7 月美元 25,000 − 9 月英鎊 4,250）", note="差額 −35 為發行折價與成本（收入為淨額）")
chk("debt.q3gap（Q3 長期借款收入 vs 7 月美元＋9 月英鎊＋Term Loan；容差 0.2B＝折價與成本）", CFV["cf.ltdProceeds"][1], 25000 + 5637 + 17500, f"{TQ} 第 18 頁；{PRF}", tol=200)
addq("debt.supplierLoan.2026q3", "供應商融資：技術基礎設施採購的擔保貸款額度（2027Q1–2028Q3 可動用，每筆 5 年到期，浮動利率；最高約 80.0B）", 80000, 19, "附註 5",
     note="期末未動用；動用後列長期債務、還本列長期借款償還；本版模型不計入（模型假設現金不足才舉債）", find=["$80.0 billion"])
addq("debt.revolver.20260930", "未動用循環信用額度（15.0B＋5.0B 364 天；2026-10 改簽：15.0B 至 2031-10、5.0B 至 2027-10）", 20000, 19, "附註 5", period="2026-09-30", find=["$20.0 billion"])
addq("debt.stBorrow.20260930", "其他短期借款（營運資金用；列「Accrued expenses and other」）", 380, 19, "附註 5", period="2026-09-30",
     note="2025-12-31 為 455；模型淨負債不計入（金額小、屬營運資金；計入時加權目標價約 −0.03），只作驗證", find=["$380 million"])
addq("debt.stbyLC.20260930", "備付信用狀（指定受益人；主要為採購承諾、支付服務、工傷與保險）", 21600, 19, "附註 5", period="2026-09-30",
     note="或有，不減少可用額度；不進模型", find=["$21.6 billion"])
addq("debt.cp.20260930", "商業本票餘額（上限 30.0B）", 0, 19, "附註 5", period="2026-09-30")
DS = {"2026Q4": 1413, "2027": 16566, "2028": 19567, "2029": 39185, "2030": 12489, "thereafter": 205034}
addq("debt.schedule.2026q3", "長期債務本金＋利息到期表（2026-09-30）", DS, 15, "附註 4 承諾表", period="2026-09-30", note="合計 294,254（含利息）", find=["294,254", "39,185"])
chk("debt.schedule（到期表列加總 vs 合計）", sum(DS.values()), 294254, f"{TQ} 第 15 頁")
# 2) 租賃與融資義務（附註 3、4，第 13–15 頁）
addq("lease.finPV.2026q3", "融資租賃負債現值（2026-09-30）", 14495, 14, "附註 3", period="2026-09-30", old="13,451（沿用 2026-06-30）", find=["14,495"])
addq("lease.finGross.2026q3", "融資租賃負債（未折現，含利息）", 17955, 14, "附註 3", period="2026-09-30", find=["17,955"])
addq("lease.opPV.2026q3", "營業租賃負債現值", 100288, 14, "附註 3", period="2026-09-30", find=["100,288"])
addq("lease.opGross.2026q3", "營業租賃負債（未折現）", 121492, 14, "附註 3", period="2026-09-30", find=["121,492"])
chk("lease.finPV（總額 − 利息 vs 現值）", 17955 - 3460, 14495, f"{TQ} 第 14 頁")
chk("lease.opPV（總額 − 利息 vs 現值）", 121492 - 21204, 100288, f"{TQ} 第 14 頁")
chk("bs.ltLease（10-Q 長期營業＋融資 vs 新聞稿資產負債表）", 86195 + 12629, BS["Long-term lease liabilities"][1], f"{TQ} 第 14 頁；{PRF} 第七節")
addq("lease.finObl.2026q3", "融資義務（物流與資料中心，含利息；2026-09-30；只作說明）", 11565, 15, "附註 4 承諾表", period="2026-09-30",
     note="未折現、含利息；模型改用本金 lease.finOblPrin.2026q3；加權剩餘 14.5 年、加權利率 3.4%", find=["11,565"])
addq("lease.finOblPrin.2026q3", "融資義務本金（不含利息；流動 426＋非流動 9,300；模型 fin_obl）", 9726, 15, "附註 4 註 (1)", period="2026-09-30", tag="[Verified]",
     old="11,070（沿用 2026-06-30，含利息口徑）；改用本金，與 debt_q3（面額）、fl_pv（現值）同為資產負債表口徑", find=["$426", "$9.3 billion"])
addq("lease.term.2026q3", "加權平均剩餘租期：營業／融資（年）", [9.9, 12.2], 13, "附註 3", unit="年", period="2026-09-30", note="折現率 4.0%／3.8%")
addq("lease.cost.2026q3", "租賃成本合計（營業 4,288＋融資 771＋變動 716）", 5775, 13, "附註 3", find=["5,775"])
UNC = {"2026Q4": 2590, "2027": 12448, "2028": 11477, "2029": 12193, "2030": 11385, "thereafter": 114829}
addq("lease.uncommenced.2026q3", "尚未起租的租賃承諾（逐年付款）", UNC, 15, "附註 4 承諾表", period="2026-09-30",
     note="合計 164,922（2026-06-30 為 137,214，季增 27,708）；起租時程與租期年限未揭露", find=["164,922", "114,829"])
addq("lease.uncommencedTotal.2026q3", "尚未起租的租賃承諾合計", 164922, 15, "附註 4", period="2026-09-30", old="137,214（2026-06-30）")
chk("lease.uncommenced（列加總 vs 合計）", sum(UNC.values()), 164922, f"{TQ} 第 15 頁")
addq("cf.ppeUnpaid.10q.2026q3", "已取得未付款之不動產及設備增加（10-Q 補充現金流資訊）", 853, 7, "附註 1 Supplemental Cash Flow", period="2026Q3",
     note="9M26 21,473；與新聞稿同值；設備延長付款期（最長一年）列應付帳款", find=["853", "21,473"])
chk("cf.ppeUnpaid.2026q3（10-Q vs 新聞稿）", 853, CFV["cf.ppeUnpaid"][1], f"{TQ} 第 7 頁；{PRF}")
chk("cf.ppeUnpaid.2026m9（10-Q vs 新聞稿）", 21473, CFV["cf.ppeUnpaid"][3], f"{TQ} 第 7 頁；{PRF}")
addq("lease.finLeaseGross.2026q3", "融資租賃取得資產總額（帳列 PP&E）", 56400, 13, "附註 3", period="2026-09-30", note="累計攤銷 39.5B", find=["$56.4 billion"])
# 3) 承諾（附註 4，第 15 頁）
CPPE = {"2026Q4": 13955, "2027": 63900, "2028": 25220, "2029": 25400, "2030": 28941, "thereafter": 2224}
addq("commit.ppe.2026q3", "不動產與設備長期採購承諾（未列資產負債表）", CPPE, 15, "附註 4 承諾表 註 (2)", period="2026-09-30", note="合計 159,640；2027 年 63.9B", find=["159,640", "63,900"])
chk("commit.ppe（列加總 vs 合計）", sum(CPPE.values()), 159640, f"{TQ} 第 15 頁")
addq("commit.purchase.2026q3", "無條件採購義務（能源、影音內容、軟體；不含設備）", 94759, 15, "附註 4 承諾表 註 (3)", period="2026-09-30",
     old="130,065（2026-06-30；當時口徑含設備，本季 10-Q 將設備另列一行，兩期不可直接比）", find=["94,759"])
addq("commit.other.2026q3", "其他承諾（資產除役、影音內容、建造中自建租賃）", 18994, 15, "附註 4", period="2026-09-30", find=["18,994"])
addq("commit.total.2026q3", "主要合約承諾合計", 883581, 15, "附註 4", period="2026-09-30", old="650,034（2026-06-30）", find=["883,581"])
chk("commit.total（各列合計 vs 總計）", 294254 + 121492 + 17955 + 11565 + 164922 + 159640 + 94759 + 18994, 883581, f"{TQ} 第 15 頁")
# 4) 分部折舊、PP&E 淨增加、分部 PP&E（附註 8，第 25–26 頁）
DAQ = {"NA": (4566, 13346, "4,566", "13,346"), "INTL": (1346, 3916, "1,346", "3,916"), "AWS": (9230, 24583, "9,230", "24,583")}
for code, (q3v, m9, s1, s2) in DAQ.items():
    oldv = F0[f"is.{code}.da.2026q2"]["value"] * CFV["cf.da"][1] / F0["cf.da.2026q2"]["value"]
    addq(f"is.{code}.da.{Q}", f"{code} 分部折舊與攤銷（不動產與設備，依使用量分攤）", q3v, 26, "附註 8", old=f"{oldv:,.1f}（推導：2026Q2 × 現金流量表 D&A 季增比）", find=[s1, s2])
    addq(f"is.{code}.da.2026m9", f"{code} 分部折舊與攤銷（9M26）", m9, 26, "附註 8")
    chk(f"is.{code}.da.2026m9（10-Q 9M vs 既有 Q1＋Q2＋10-Q Q3）", m9, F0[f"is.{code}.da.2026q1"]["value"] + F0[f"is.{code}.da.2026q2"]["value"] + q3v, f"{TQ} 第 26 頁；amazon_facts_20261008.json")
    chk(f"is.{code}.da.2025q3（10-Q 比較欄 vs 既有總帳）", {"NA": 4001, "INTL": 1276, "AWS": 5648}[code], F0[f"is.{code}.da.2025q3"]["value"], f"{TQ} 第 26 頁")
chk("分部折舊 Q3 合計（10-Q 合併列）", sum(v[0] for v in DAQ.values()), 15142, f"{TQ} 第 26 頁")
PAQ = {"NA": (14530, 37795), "INTL": (6068, 10335), "AWS": (57116, 147236), "CORP": (548, 1544), "TOT": (78262, 196910)}
for code, (q3v, m9) in PAQ.items():
    addq(f"capex.netAdd.{code}.{Q}", f"不動產與設備淨增加（{code}；含融資租賃與未付款，非現金口徑）", q3v, 26, "附註 8", find=[f"{q3v:,}", f"{m9:,}"])
    addq(f"capex.netAdd.{code}.2026m9", f"不動產與設備淨增加（{code}，9M26）", m9, 26, "附註 8")
    chk(f"capex.netAdd.{code}（9M − 1H26 vs Q3）", m9 - F0[f"capex.netAdd.{code}.2026h1"]["value"], q3v, f"{TQ} 第 26 頁；10-Q 2026Q2")
chk("capex.netAdd（分部加總 vs 合併，Q3）", sum(PAQ[c][0] for c in ("NA", "INTL", "AWS", "CORP")), PAQ["TOT"][0], f"{TQ} 第 26 頁")
PPE = {"NA": 145087, "INTL": 36263, "AWS": 313009, "CORP": 14503}
for code, v in PPE.items():
    addq(f"bs.ppe.{code}.{Q}", f"不動產與設備淨額（{code}，2026-09-30）", v, 26, "附註 8", period="2026-09-30", find=[f"{v:,}"])
chk("bs.ppe（10-Q 分部加總 vs 新聞稿資產負債表）", sum(PPE.values()), BS["Property and equipment, net"][1], f"{TQ} 第 26 頁；{PRF} 第七節")
addq("bs.segAssets.AWS.2026q3", "AWS 分部資產（2026-09-30）", 419594, 25, "附註 8", period="2026-09-30", note="2025-12-31 為 252,588", find=["419,594"])
addq("cip.20260930", "在建工程（construction in progress）", None, 26, "（10-Q 未揭露；只見於 10-K 附註）", unit="—", period="2026-09-30", tag="[不存在]",
     note="10-Q 不揭露 PP&E 組成；模型的在建以「實績期超出新增在役 MW 成本的 AI 資本支出」推導（Thesis cx_cip），不受影響")
# 5) RPO（附註 1，第 9–10 頁）
addq("rpo.2026q3", "RPO（原始期限 >1 年的未認列合約承諾，主要為 AWS）", 578000, 10, "附註 1 Unearned Revenue", period="2026-09-30", old="496,000（2026-06-30）；法說 backlog 578",
     note="加權平均剩餘年限 6.1 年（Q2 6.4 年）；12 或 24 個月內認列比例未揭露", find=["$578 billion", "6.1 years"])
addq("rpo.life.2026q3", "RPO 加權平均剩餘年限", 6.1, 10, "附註 1", unit="年", period="2026-09-30", old="6.4（2026-06-30）")
# 6) 股數、利息、持股（封面、第 4、12 頁）
addq("sh.cover.20261021", "封面流通股數（2026-10-21；只作驗證）", 10813.138306, 1, "封面", unit="M 股", period="2026-10-21",
     note="比 9/30 資產負債表 10,811 多約 2.1 百萬股（9/30–10/21 歸屬）；模型用 9/30 股數，避免與 9/30 淨負債、未歸屬獎勵重複計入", find=["10,813,138,306"])
chk("sh.cover − sh.out（10/21 − 9/30，應小於未歸屬獎勵 223）", 10813.138306 - sh_out, 2.1, f"{TQ} 第 1 頁；{PRF}", tol=0.1)
addq(f"is.intInc.{Q}.10q", "利息收入（10-Q 損益表）", 1187, 4, "合併損益表", period="2026Q3", note="MD&A：主要投資於投資級短中期債券與 AAA 貨幣基金；利息收入隨平均投資餘額與利率變動，10-Q 不提供更細的明細", find=["1,187", "3,617"])
chk("is.intInc.2026q3（10-Q vs 新聞稿）", 1187, IS["Interest income"][1], f"{TQ} 第 4 頁；{PRF}")
chk("is.intExp.2026q3（10-Q vs 新聞稿）", 1691, abs(IS["Interest expense"][1]), f"{TQ} 第 4 頁；{PRF}")
addq("inv.privEq.20260930", "非權益法非上市權益投資帳面（主要為 Anthropic 優先股與 OpenAI 優先股）", 143700, 12, "附註 2", period="2026-09-30", note="2025-12-31 為 16.2B", find=["$143.7 billion"])
addq("inv.antPref.20260930", "Anthropic 無表決權優先股帳面（164 百萬股）", 92500, 12, "附註 2", period="2026-09-30", note="9M26 依 Anthropic 募資的可觀察價格上調 62.8B", find=["$92.5 billion", "164 million"])
addq("inv.antNotes.20260930", "Anthropic 可轉換票據公允價值（不計持股上限可轉約 184 百萬股優先股）", 146200, 12, "附註 2", period="2026-09-30",
     note="未實現利益 140.3B 列其他綜合損益；每股優先股可轉 10 股普通股；票據轉換受 12% 持股上限限制，但上限「may be waived at our election」（第 11 頁）", find=["$146.2 billion", "184 million"])
addq("inv.oaiPref.20260930", "OpenAI Series C 優先股帳面（Q1 15.0＋Q2 13.7＋Q3 21.3）", 50000, 12, "附註 2", period="2026-09-30", find=["$21.3 billion"])
addq("inv.antFacility.20260930", "Anthropic 融資額度可動用上限（20.0B − Series H 5.0B，依算力交付里程碑開放）", 15000, 12, "附註 2", period="2026-09-30", find=["$15.0 billion"])
addq("acq.globalstar", "Globalstar 收購（現金＋股票，含債務約 10.9B；預計 2027 年完成）", 10900, 15, "附註 4", period="2026-04-13", note="本版模型不計入（2027 年的現金部分最多 40%）", find=["$10.9 billion"])
addq("capex.view.2027", "MD&A：預期 2027 年現金資本支出「meaningfully increase」，主要為技術基礎設施", None, 28, "MD&A Liquidity", unit="—", tag="[Interested-party]", find=["meaningfully increase in 2027"])
# 第三讀：原文文字檔（環境變數 AMZN_10Q_TXT；不入 repo）
TXT = os.environ.get("AMZN_10Q_TXT", "/home/claude/src10q/amzn_q3.txt")
n_txt = 0
if os.path.exists(TXT):
    pages = open(TXT, encoding="utf-8").read().split("\f")
    for iid, (pg, needles) in PG.items():
        for nd in needles:
            ok = nd in pages[pg - 1]
            CHK.append({"id": f"{iid}（10-Q 原文第 {pg} 頁含「{nd}」）", "value": nd, "second": "found" if ok else "missing", "source": "AMZN_10Q_TXT", "ok": ok}); n_txt += 1
    TXT_NOTE = f"10-Q 原文逐值比對 {n_txt} 項"
else:
    TXT_NOTE = "10-Q 原文文字檔不存在，略過逐值比對（10-Q 內部互對與對新聞稿仍執行）"
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
OUT = {"_readme": "Amazon 2026Q3 事實總帳（MAG v0.3.1 工作單第 1 節第 1 項；v0.3.2 補入 10-Q：mag/docs/workorders/20261104_MAG_10Q補入_2026Q3.md）。由 amazon/facts_2026q3.py 自財報新聞稿逐列轉錄讀取，"
                  "10-Q 附註數字（債務明細與面額、租賃、融資義務、分部折舊、分部 PP&E、RPO、承諾、封面股數、持股帳面）逐筆附頁碼登錄（全文不入 repo）；原推導或沿用值保留在 note「原（v0.3.1）：」。"
                  "金額 USD M（費用與流出以正值記）；checks＝第二次獨立讀取。",
       "asOf": "2026-09-30", "items": ITEMS, "checks": CHK}
json.dump(OUT, open(os.path.join(HERE, "data", "amazon_facts_2026q3.json"), "w", encoding="utf-8"), ensure_ascii=False, indent=1)
bad = [c for c in CHK if not c["ok"]]
print(f"事實 {len(ITEMS)} 筆；兩讀核對 {len(CHK)} 項（其中 IR 網頁第二讀 {n_ir} 項；{TXT_NOTE}），一致 {len(CHK) - len(bad)}，不一致 {len(bad)}")
for c in bad: print("  不一致：", c)
sys.exit(1 if bad else 0)
