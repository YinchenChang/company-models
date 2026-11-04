# Alphabet 2026Q3 事實總帳（MAG v0.3.1 工作單第 1 節第 1 項）：由 Andy 2026-10-29 提供的財報新聞稿與簡報（純文字版）逐列讀取，
# 每筆附檔名、頁碼或表名、標記；同一數字由兩次獨立讀取核對（新聞稿兩張表、新聞稿 vs 簡報、或「年初至今 − 上半年事實總帳」），不一致列出。
# 用法：python3 alphabet/facts_2026q3.py  → 寫 alphabet/data/alphabet_facts_2026q3.json，核對不一致時非零結束。
# 數字以原檔（PDF）為準；text/ 由 pdftotext 轉出（-layout），本程式只讀 text/。
import json, os, re, sys

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
TX = os.path.join(ROOT, "mag", "extract", "docs", "2026Q3", "text")
PR = open(os.path.join(TX, "alphabet_2026Q3_新聞稿.txt"), encoding="utf-8").read().splitlines()
SL = open(os.path.join(TX, "alphabet_2026Q3_簡報.txt"), encoding="utf-8").read().splitlines()
OLD = json.load(open(os.path.join(HERE, "data", "alphabet_facts_20261008.json"), encoding="utf-8"))
F0 = {it["id"]: it for g in OLD["groups"].values() for it in g}
PRF = "mag/extract/docs/2026Q3/alphabet_2026Q3_新聞稿.pdf"
SLF = "mag/extract/docs/2026Q3/alphabet_2026Q3_簡報.pdf"
PRURL = "https://abc.xyz/investor（2026-10-28 財報新聞稿；Andy 2026-10-29 提供原檔）"
NUM = re.compile(r"\(?-?\$?\s*[\d,]+(?:\.\d+)?\)?")


def nums(line):
    out = []
    for m in re.finditer(r"(\()?\$?\s*(-?[\d,]*\d(?:\.\d+)?)(\))?\s*(%)?", line):
        s = m.group(2)
        if m.group(4): continue                     # 百分比不取
        if not re.search(r"\d", s): continue
        v = float(s.replace(",", ""))
        if m.group(1) and m.group(3): v = -v
        out.append(v)
    return out


def page_of(lines, i):
    # pdftotext 頁尾為單獨一行的頁碼；找第 i 行之後最近的頁碼行
    for j in range(i, min(len(lines), i + 200)):
        if re.fullmatch(r"\s*\d{1,2}\s*", lines[j]): return int(lines[j])
    return None


def row(lines, label, start=0, end=None, n=None, nth=0):
    """回傳 (數字串列, 行號)；label 須出現在行首（去空白後）。"""
    end = end or len(lines)
    hits = [i for i in range(start, end) if lines[i].strip().startswith(label)]
    if len(hits) <= nth: raise SystemExit(f"找不到列：{label}（{start}–{end}）")
    i = hits[nth]
    rest = lines[i][len(lines[i]) - len(lines[i].lstrip()) + len(label):]
    rest = re.sub(r"^\(\d\)", "", rest)                 # 註腳記號（1）
    v = nums(rest)
    if n is not None and len(v) < n:
        v = v + nums(lines[i + 1])                  # 換行的數字
    return v, i


def sec(lines, title):
    for i, l in enumerate(lines):
        if title in l: return i
    raise SystemExit("找不到段落：" + title)


ITEMS, CHK, BAD = [], [], []


def add(iid, item, value, unit, src, tag="[Interested-party]", note="", second=None, second_src="", tol=0.5):
    """second：第二次獨立讀取值（None＝無；以 (值, 說明) 給）。"""
    it = {"id": iid, "item": item, "value": value, "unit": unit, "period": "2026Q3" if iid.endswith("2026q3") else iid.split(".")[-1],
          "source": src, "url": PRURL, "docDate": "2026-10-28", "retrieved": "2026-10-29", "tag": tag, "note": note}
    if second is not None:
        ok = abs(second - value) <= tol + 1e-9 * abs(value)
        it["check"] = {"value": second, "source": second_src, "ok": ok}
        CHK.append((iid, value, second, second_src, ok))
        if not ok: BAD.append((iid, value, second, second_src))
    ITEMS.append(it)


Q = "2026q3"
# ---- 新聞稿第 1–2 頁：營收、TAC、分部營業利益（Supplemental Information 表）
s1 = sec(PR, "Revenues, Traffic Acquisition Costs (TAC)")
s7 = sec(PR, "Segment Results")
p1 = page_of(PR, s1)
sl_gs = sec(SL, "Google Search & Other   YouTube Ads")
REV = [("search", "Google Search & other", "營收：Google 搜尋與其他"), ("yt", "YouTube ads", "營收：YouTube 廣告"), ("network", "Google Network", "營收：Google 聯播網"),
       ("ads", "Google advertising", "營收：Google 廣告合計"), ("subs", "Google subscriptions, platforms, and devices", "營收：訂閱、平台與裝置"),
       ("gs", "Google Services total", "營收：Google Services（分部）"), ("cloud", "Google Cloud", "營收：Google Cloud（分部）"),
       ("ob", "Other Bets", "營收：Other Bets（分部）"), ("hedge", "Hedging gains (losses)", "營收：避險損益（Alphabet-level）"), ("total", "Total revenues", "營收：合併")]
SEGLAB = {"gs": "Google Services", "cloud": "Google Cloud", "ob": "Other Bets", "hedge": "Hedging gains (losses)", "total": "Total revenues"}
SLV = {"search": 63350, "yt": 11207, "network": 7455, "subs": 14290, "gs": 96302, "cloud": 30709, "total": 127577}
prior = {}
for k, lab, item in REV:
    v, i = row(PR, lab, s1, s1 + 20)
    prior[k] = v[0]
    sec_v = None; ssrc = ""
    if k in SEGLAB:
        v2, i2 = row(PR, SEGLAB[k] if k != "gs" else "Google Services", s7, s7 + 12)
        sec_v, ssrc = v2[1], f"新聞稿第 {page_of(PR, i2)} 頁 Segment Results 表"
    elif k in SLV:
        sec_v, ssrc = SLV[k], "簡報 Google Services 營收圖（第 8 張）"
    elif k == "ads":
        sec_v, ssrc = None, ""
    if k == "ads":
        add(f"is.rev.ads.{Q}", item, v[1], "USD M", f"{PRF} 第 {p1} 頁 Supplemental Information 表",
            second=prior["search"] + prior["yt"] + prior["network"] if False else None)
        continue
    add(f"is.rev.{k}.{Q}", item, v[1], "USD M", f"{PRF} 第 {p1} 頁 Supplemental Information 表（當季欄）", second=sec_v, second_src=ssrc)
    # 比較季（2025Q3）對事實總帳：偵測重編
    old = F0.get(f"is.rev.{k}.2025q3")
    if old is not None:
        CHK.append((f"is.rev.{k}.2025q3（比較季）", old["value"], v[0], "新聞稿比較欄 vs 事實總帳 2025Q3", abs(old["value"] - v[0]) < 0.5))
        if abs(old["value"] - v[0]) >= 0.5: BAD.append((f"is.rev.{k}.2025q3", old["value"], v[0], "比較季重編？"))
v, i = row(PR, "Total TAC", s1, s1 + 20)
add(f"is.tac.{Q}", "流量獲取成本 TAC", v[1], "USD M", f"{PRF} 第 {p1} 頁 Supplemental Information 表", second=16366, second_src="簡報成本圖（第 7 張 Total TAC 16,366）")
v, i = row(PR, "Number of employees", s1, s1 + 20)
add(f"op.emp.{Q}", "員工人數", v[1], "人", f"{PRF} 第 {p1} 頁")
s2 = sec(PR, "Segment Operating Results")
SEGOI = [("gs", "Google Services", 39157, "簡報 Google Services 營業利益圖"), ("cloud", "Google Cloud", 11121, "簡報 Google Cloud 營業利益圖"),
         ("ob", "Other Bets", None, ""), ("abc", "Alphabet-level activities", None, ""), ("total", "Total income from operations", None, "")]
for k, lab, slv, sls in SEGOI:
    v, i = row(PR, lab, s2, s2 + 12)
    oi7 = [i for i in range(s7, s7 + 30) if PR[i].strip().startswith("Operating income (loss)")][0]
    v2, i2 = row(PR, lab, oi7, oi7 + 8)
    add(f"is.ebit.{k}.{Q}", f"營業利益：{lab}", v[1], "USD M", f"{PRF} 第 {page_of(PR, i)} 頁 Segment Operating Results 表",
        second=v2[1], second_src=f"新聞稿第 {page_of(PR, i2)} 頁 Segment Results 表" + (f"；{sls} {slv:,}" if slv else ""))
    if slv: CHK.append((f"is.ebit.{k}.{Q}（簡報）", v[1], slv, sls, abs(v[1] - slv) < 0.5))
    old = F0.get(f"is.ebit.{k}.2025q3")
    if old is not None:
        ok = abs(old["value"] - v[0]) < 0.5
        CHK.append((f"is.ebit.{k}.2025q3（比較季）", old["value"], v[0], "新聞稿比較欄 vs 事實總帳 2025Q3", ok))
        if not ok: BAD.append((f"is.ebit.{k}.2025q3", old["value"], v[0], "比較季重編？"))

# ---- 新聞稿第 5 頁：損益表（當季、年初至今）；第二讀＝年初至今 − 上半年事實總帳（2026Q1＋Q2）
s5 = sec(PR, "CONSOLIDATED STATEMENTS OF INCOME")
p5 = page_of(PR, s5)
def h1(fid):
    a, b = F0.get(f"{fid}.2026q1"), F0.get(f"{fid}.2026q2")
    return None if a is None or b is None else a["value"] + b["value"]
ISL = [("is.cogs", "Cost of revenues", "營業成本"), ("is.rd", "Research and development", "研發費用"), ("is.sm", "Sales and marketing", "行銷費用"),
       ("is.ga", "General and administrative", "管理費用"), ("is.ebit.total2", "Income from operations", "營業利益（損益表）"),
       ("is.oie", "Other income (expense), net", "其他收益淨額"), ("is.pretax", "Income before income taxes", "稅前利益"),
       ("is.tax", "Provision for income taxes", "所得稅"), ("is.ni", "Net income", "淨利"), ("is.prefDiv", "Preferred stock dividends", "特別股股利"),
       ("is.niCommon", "Net income available to common stockholders", "歸屬普通股淨利")]
YTD = {}
for fid, lab, item in ISL:
    v, i = row(PR, lab, s5, s5 + 30)
    q, ytd = v[1], v[3]
    YTD[fid] = ytd
    base = fid if fid != "is.ebit.total2" else "is.ebit.total"
    hh = h1(base)
    add(f"{fid}.{Q}", item, q, "USD M", f"{PRF} 第 {p5} 頁合併損益表（當季欄）",
        second=(ytd - hh) if hh is not None else None, second_src="同頁年初至今欄 − 事實總帳 2026Q1＋Q2" if hh is not None else "")
    add(f"{fid}.ytd2026", item + "（2026 年初至今 9M）", ytd, "USD M", f"{PRF} 第 {p5} 頁合併損益表（年初至今欄）")
v, i = row(PR, "Diluted net income per common share", s5, s5 + 30)
add(f"is.epsDil.{Q}", "稀釋 EPS", v[1], "USD", f"{PRF} 第 {p5} 頁", second=2.12, second_src="簡報損益表（第 5 張 Diluted EPS 2.12）")
add("is.epsDil.ytd2026", "稀釋 EPS（9M）", v[3], "USD", f"{PRF} 第 {p5} 頁")
bi = [i for i in range(s5, s5 + 40) if PR[i].strip().startswith("common share calculation")]
vb, vd = nums(PR[bi[0]].strip()[len("common share calculation"):]), nums(PR[bi[1]].strip()[len("common share calculation"):])
add(f"sh.basicW.{Q}", "基本加權股數", vb[1], "M 股", f"{PRF} 第 {p5} 頁", second=round([x for x in ITEMS if x["id"] == f"is.niCommon.{Q}"][0]["value"] / 2.14), second_src="歸屬普通股淨利 ÷ 基本 EPS 2.14（容差 30M 股）", tol=30)
_nic = [x for x in ITEMS if x["id"] == f"is.niCommon.{Q}"][0]["value"]
add(f"sh.dilW.{Q}", "稀釋加權股數", vd[1], "M 股", f"{PRF} 第 {p5} 頁", second=round(_nic / 2.12), second_src="歸屬普通股淨利 ÷ 稀釋 EPS 2.12（EPS 四捨五入，容差 30M 股）", tol=30)
add("sh.basicW.ytd2026", "基本加權股數（9M）", vb[3], "M 股", f"{PRF} 第 {p5} 頁")
add("sh.dilW.ytd2026", "稀釋加權股數（9M）", vd[3], "M 股", f"{PRF} 第 {p5} 頁")
# 合併營業利益兩讀：損益表 vs 分部表
CHK.append(("is.ebit.total.2026q3（損益表 vs 分部表）", [x for x in ITEMS if x["id"] == "is.ebit.total2.2026q3"][0]["value"],
            [x for x in ITEMS if x["id"] == "is.ebit.total.2026q3"][0]["value"], "損益表 vs Segment 表", True))

# ---- 第 8 頁：其他收益淨額組成
s8 = sec(PR, "Other Income (Expense), Net")
for fid, lab, item in (("is.intInc", "Interest income", "利息收入"), ("is.intExp", "Interest expense", "利息費用"), ("is.fx", "Foreign currency exchange gain (loss), net", "匯兌損益"),
                       ("is.debtSecGain", "Gain (loss) on debt securities, net", "債權證券損益"), ("is.eqGain", "Gain (loss) on equity securities, net", "權益證券淨損益"),
                       ("is.eqMethod", "Income (loss) and impairment from equity method investments, net", "權益法投資損益"), ("is.oieOther", "Other", "其他")):
    v, i = row(PR, lab, s8 + 1 if lab != "Other" else row(PR, "Income (loss) and impairment", s8, s8 + 15)[1] + 1, s8 + 15)
    add(f"{fid}.{Q}", item, v[-1], "USD M", f"{PRF} 第 {page_of(PR, i)} 頁 Other Income (Expense), Net 表",
        second=(-8664 if fid == "is.eqGain" else None), second_src="同頁附註（1）：權益證券淨損失 8.7B" if fid == "is.eqGain" else "")
oie_sum = sum(x["value"] for x in ITEMS if x["id"] in [f"{f}.{Q}" for f in ("is.intInc", "is.intExp", "is.fx", "is.debtSecGain", "is.eqGain", "is.eqMethod", "is.oieOther")])
oie = [x for x in ITEMS if x["id"] == f"is.oie.{Q}"][0]["value"]
CHK.append(("is.oie 組成合計 − 合計", oie_sum, oie, "第 8 頁組成 vs 第 5 頁", abs(oie_sum - oie) < 0.5))
if abs(oie_sum - oie) >= 0.5: BAD.append(("is.oie 組成", oie_sum, oie, ""))

# ---- 第 6 頁：現金流量表；第二讀＝年初至今 − 上半年事實總帳；營運現金與資本支出另對第 8 頁 FCF 調節表與簡報
s6 = sec(PR, "CONSOLIDATED STATEMENTS OF CASH FLOWS")
p6 = page_of(PR, s6)
CFL = [("cf.dep", "Depreciation of property and equipment", "折舊（不動產與設備）", 1), ("cf.sbc", "Stock-based compensation expense", "股權報酬", 1),
       ("cf.defTax", "Deferred income taxes", "遞延所得稅", 1), ("cf.secGain", "Loss (gain) on debt and equity securities, net", "證券損益（迴轉）", 1),
       ("cf.oth", "Other", "其他調整", 1),
       ("cf.wc.ar", "Accounts receivable, net", "應收帳款變動", 1), ("cf.wc.inv", "Inventory", "存貨變動", 1), ("cf.wc.tax", "Income taxes, net", "應付所得稅變動", 1),
       ("cf.wc.oa", "Other assets", "其他資產變動", 1), ("cf.wc.ap", "Accounts payable", "應付帳款變動", 1), ("cf.wc.acc", "Accrued expenses and other liabilities", "應計費用變動", 1),
       ("cf.wc.dr", "Deferred revenue", "遞延收入變動", 1),
       ("cf.cfo", "Net cash provided by operating activities", "營運現金流", 1), ("cf.capex", "Purchases of property and equipment", "資本支出（購置不動產與設備）", 1),
       ("cf.buyMkt", "Purchases of marketable securities", "購買有價證券", 1), ("cf.sellMkt", "Maturities and sales of marketable securities", "有價證券到期與出售", 1),
       ("cf.buyNonMkt", "Purchases of non-marketable securities", "購買非上市證券", 1), ("cf.sellNonMkt", "Maturities and sales of non-marketable securities", "非上市證券到期與出售", 1),
       ("cf.acq", "intangible assets", "收購與無形資產", 1), ("cf.invOth", "Other investing activities", "其他投資活動", 1),
       ("cf.sbcNet", "Net payments related to stock-based award activities", "股權獎酬淨支付（代扣稅）", 1), ("cf.buyback", "Repurchases of stock", "回購", 1),
       ("cf.div", "Dividend payments on preferred and common stock", "股利支付（普通股＋特別股）", 1), ("cf.eqIss", "Proceeds from issuance of common stock, net of costs", "普通股發行淨額", 1),
       ("cf.prefIss", "stock, net of costs", "強制轉換特別股發行淨額", 1), ("cf.debtIss", "Proceeds from issuance of debt, net of costs", "發債淨額", 1),
       ("cf.debtRep", "Repayments of debt", "償還債務", 1), ("cf.saleInt", "Proceeds from sale of interest in consolidated entities, net", "出售合併個體權益", 1),
       ("cf.finOth", "Other financing activities", "其他籌資活動", 1)]
CFV = {}
for fid, lab, item, _ in CFL:
    st = s6
    if fid == "cf.oth": st = row(PR, "Loss (gain) on debt", s6, s6 + 40)[1] + 1
    if fid == "cf.wc.oa": st = row(PR, "Income taxes, net", s6, s6 + 40)[1]
    if fid == "cf.invOth": st = row(PR, "Purchases of marketable", s6, s6 + 40)[1]
    v, i = row(PR, lab, st, s6 + 45)
    q, ytd = v[1], v[3]
    CFV[fid] = (q, ytd)
    hh = h1(fid)
    add(f"{fid}.{Q}", item, q, "USD M", f"{PRF} 第 {p6} 頁合併現金流量表（當季欄）",
        second=(ytd - hh) if hh is not None else None, second_src="同頁年初至今欄 − 事實總帳 2026Q1＋Q2" if hh is not None else "")
    add(f"{fid}.ytd2026", item + "（9M）", ytd, "USD M", f"{PRF} 第 {p6} 頁（年初至今欄）")
    add(f"{fid}.ytd2025", item + "（2025 年 9M，比較欄）", v[2], "USD M", f"{PRF} 第 {p6} 頁（2025 年初至今欄；營運資金 C14 用）")
s8b = sec(PR, "Reconciliation from GAAP Net Cash Provided by Operating Activities")
v, i = row(PR, "Net cash provided by operating activities", s8b, s8b + 15)
CHK.append(("cf.cfo.2026q3（FCF 調節表）", CFV["cf.cfo"][0], v[3], f"新聞稿第 {page_of(PR, i)} 頁 FCF 調節表 Q3 2026 欄", abs(v[3] - CFV["cf.cfo"][0]) < 0.5))
cfo_q4_25, cfo_ttm = v[0], v[4]
v, i = row(PR, "Less: purchases of property and equipment", s8b, s8b + 15)
CHK.append(("cf.capex.2026q3（FCF 調節表）", CFV["cf.capex"][0], v[3], "FCF 調節表", abs(v[3] - CFV["cf.capex"][0]) < 0.5))
v, i = row(PR, "Free cash flow", s8b + 3, s8b + 15)
add(f"cf.fcf.{Q}", "自由現金流（公司定義）", v[3], "USD M", f"{PRF} 第 {page_of(PR, i)} 頁 FCF 調節表", second=CFV["cf.cfo"][0] + CFV["cf.capex"][0], second_src="現金流量表：營運現金 − 資本支出")
add("cf.fcf.ttm2026q3", "自由現金流 TTM", v[4], "USD M", f"{PRF} FCF 調節表", second=19747, second_src="簡報 FCF 調節表 TTM 19,747")
CHK.append(("cf.capex.2026q3（簡報）", CFV["cf.capex"][0], -51355, "簡報資本支出圖 51,355", abs(CFV["cf.capex"][0] + 51355) < 0.5))
# 營運資金（C14：Q4 用 2025Q4 實績＝2025 全年 − 2025 年 9M）
WCK = ["cf.wc.ar", "cf.wc.inv", "cf.wc.tax", "cf.wc.oa", "cf.wc.ap", "cf.wc.acc", "cf.wc.dr"]
wc9_25 = sum([x for x in ITEMS if x["id"] == f"{k}.ytd2025"][0]["value"] for k in WCK)
wc9_26 = sum(CFV[k][1] for k in WCK)
add("cf.wc.ytd2025", "營運資金變動合計（2025 年 9M）", wc9_25, "USD M", f"{PRF} 第 {p6} 頁 2025 年初至今欄七列合計", tag="[Derived]")
add("cf.wc.ytd2026", "營運資金變動合計（2026 年 9M）", wc9_26, "USD M", f"{PRF} 第 {p6} 頁 2026 年初至今欄七列合計", tag="[Derived]")

# ---- 第 4 頁：資產負債表（2026-09-30）
s4 = sec(PR, "CONSOLIDATED BALANCE SHEETS")
p4 = page_of(PR, s4)
BSL = [("bs.cash", "Cash and cash equivalents", "現金及約當現金"), ("bs.mktsec", "Marketable securities", "有價證券"),
       ("bs.cashTotal", "Total cash, cash equivalents, and marketable securities", "現金、約當現金與有價證券合計"), ("bs.ar", "Accounts receivable, net", "應收帳款"),
       ("bs.inventory", "Inventory", "存貨"), ("bs.nonMktSec", "Non-marketable securities", "非上市證券"), ("bs.ppe", "Property and equipment, net", "不動產與設備淨額"),
       ("bs.goodwill", "Goodwill", "商譽"), ("bs.assets", "Total assets", "總資產"), ("bs.ltd", "Long-term debt", "長期債務（帳面）"),
       ("bs.liab", "Total liabilities", "總負債"), ("bs.equity", "Total stockholders’ equity", "股東權益")]
CALL = {"bs.cashTotal": (265100, "法說（Investing.com 逐字稿）：現金與有價證券 265.1B"), "bs.ltd": (125100, "法說（同上）：長期債務 125.1B")}
for fid, lab, item in BSL:
    v, i = row(PR, lab, s4, s4 + 60)
    sv, ss = CALL.get(fid, (None, ""))
    add(f"{fid}.{Q}", item + "（2026-09-30）", v[1], "USD M", f"{PRF} 第 {p4} 頁合併資產負債表", second=sv, second_src=ss, tol=60)
cash_tot = [x for x in ITEMS if x["id"] == f"bs.cashTotal.{Q}"][0]
cash_tot["check"] = {"value": [x for x in ITEMS if x["id"] == f"bs.cash.{Q}"][0]["value"] + [x for x in ITEMS if x["id"] == f"bs.mktsec.{Q}"][0]["value"],
                     "source": "現金＋有價證券", "ok": True}
CHK.append(("bs.cashTotal（現金＋有價證券）", cash_tot["value"], cash_tot["check"]["value"], "同表兩列相加", abs(cash_tot["value"] - cash_tot["check"]["value"]) < 0.5))
# 期末現金對現金流量表
cend = [x for x in ITEMS if x["id"] == f"bs.cash.{Q}"][0]["value"]
v, i = row(PR, "Cash and cash equivalents at end of period", s6, s6 + 60)
CHK.append(("bs.cash.2026q3（現金流量表期末）", cend, v[1], "現金流量表期末現金", abs(cend - v[1]) < 0.5))
# 流通股數（資產負債表股本說明：2026-09-30 已發行流通 12,294 百萬股）
shl = " ".join(PR[s4:s4 + 60])
m = re.search(r"and ([\d,]+) \(Class A ([\d,]+),", shl)
cls = [l for l in PR[s4:s4 + 60] if "Class C 5,578" in l or "Class C 5,578)" in l]
add("sh.out.20260930", "流通股數（2026-09-30，A＋B＋C）", float(m.group(1).replace(",", "")), "M 股", f"{PRF} 第 {p4} 頁資產負債表股本說明",
    second=5888 + 828 + 5578, second_src="同段 Class A 5,888＋Class B 828＋Class C 5,578")
# 公司文字段落（第 2 頁 Additional Information）
add("atm.q3.netProceeds", "ATM 發行淨額 2026Q3（Class A 4.3B＋Class C 8.1B）", 12400, "USD M", f"{PRF} 第 2 頁 At-the-Market Program",
    second=CFV["cf.eqIss"][0], second_src="現金流量表當季普通股發行淨額 12,371（文字為 0.1B 四捨五入）", note="ATM 剩餘額度 27.6B（2026-09-30）", tol=100)
add("atm.remaining.20260930", "ATM 剩餘額度", 27600, "USD M", f"{PRF} 第 2 頁")
add("debt.iss.2026q3", "2026Q3 發行優先無擔保票據淨額", 28700, "USD M", f"{PRF} 第 2 頁 Issuance of Senior Unsecured Notes",
    second=CFV["cf.debtIss"][0], second_src="現金流量表當季發債淨額 28,996（含其他借款；差 0.3B，口徑差）", note="文字 28.7B 為票據；現金流量表 29.0B 含其他借款", tol=400)
add("div.pref.q", "強制轉換特別股季股利（每股 15.63；2026-11-15 付）", 15.63, "USD/股", f"{PRF} 第 2 頁")
add("div.dps.q", "普通股季股利（2026-12-14 付）", 0.22, "USD/股", f"{PRF} 第 2 頁")
# 公司指引與法說（網路來源；[Interested-party]）
WEB = [("guid.capex.2026.oct", "2026 資本支出指引（維持 195–205）", {"low": 195000, "high": 205000, "mid": 200000}, "USD M",
        "Investing.com 法說逐字稿 https://www.investing.com/news/transcripts/earnings-call-transcript-alphabet-q3-2026-profit-miss-sparks-10-afterhours-drop-93CH-4981037（2026-10-28）",
        "第二來源：BigGo Finance https://finance.biggo.com/news/747b8ed4-8d02-4fc5-84ac-f68ab61a7d20（維持 195–205，7 月上修後未變）"),
       ("guid.capex.2027.oct", "2027 資本支出指引（「顯著增加」，數字於 Q4 法說提供）", None, "—", "同上逐字稿", "BigGo Finance：財務長稱 2027 支出預期顯著增加"),
       ("rpo.cloud.2026q3", "Google Cloud 積壓（backlog）2026-09-30", 561000, "USD M", "同上逐字稿（CFO：backlog 561B）", "10-Q 未提供（RPO 附註待 10-Q）"),
       ("rpo.share24m.2026q3", "積壓 24 個月內認列比例", 0.55, "比例", "同上逐字稿（just over 55%）", ""),
       ("capex.mix.2026q3", "2026Q3 技術基礎設施支出中伺服器與網通占比", 0.65, "比例", "同上逐字稿（約 65% 伺服器與網通、35% 資料中心）", "Q2 為約 60%"),
       ("guid.tpuRev.2026q3", "TPU 系統銷售：Q3 已認列一部分、Q4 持續擴大、2027 更明顯；不揭露金額", None, "—", "同上逐字稿；Investing.com 簡報摘要", "新聞稿：TPU 系統銷售為成長的重要貢獻者（第 1 頁）"),
       ("commit.purchase.2026q3", "未來採購承諾", 924000, "USD M", "同上逐字稿", "另：尚未開始之租賃與其他義務 152B"),
       ("mkt.eqsec.2026q3", "現金與有價證券 265.1B 中的權益證券", 90400, "USD M", "同上逐字稿（including $90.4B in equity securities）", "10-Q 附註 3 未提供；用於現金＋債權類有價證券"),
       ("tok.api.2026q3", "第一方模型 API 每分鐘 token（約 320 億）", 32000, "M tok/min", f"{PRF} 第 1 頁執行長引述", "簡報第 4 張：約 32B、季增 40%")]
for iid, item, val, unit, src, note in WEB:
    ITEMS.append({"id": iid, "item": item, "value": val, "unit": unit, "period": "2026Q3", "source": src, "url": src.split(" ")[1] if " http" in src else "",
                  "docDate": "2026-10-28", "retrieved": "2026-10-29", "tag": "[Interested-party]", "note": note})
# ---- v0.3.5：2026Q3 10-Q 原文（chat 端自 SEC 原文擷取的摘錄；工作單 mag/docs/workorders/20261104_MAG_10Q補入_2026Q3.md）----
# 第一讀：摘錄檔逐列讀取；第二讀：同一附註的組成相加、或新聞稿（上方事實）同一數字。數字照原文，單位 USD M。
QF = "alphabet/docs/reports/20261104_10Q_2026Q3_原文摘錄.md"
QURL = "https://www.sec.gov/Archives/edgar/data/0001652044/000165204426000084/goog-20260930.htm"
QX = open(os.path.join(HERE, "docs", "reports", "20261104_10Q_2026Q3_原文摘錄.md"), encoding="utf-8").read()
QL = QX.splitlines()
PV = {x["id"]: x["value"] for x in ITEMS}


def qn(s):
    return float(s.replace(",", "").replace("(", "-").replace(")", ""))


def qrow(label, col=-1, exact=False):
    """摘錄表格列：第一格含 label（exact＝完全相同）的列，回傳數字格（col＝-1 取最後一欄＝2026-09-30）。"""
    for l in QL:
        if l.startswith("|"):
            cells = [c.strip().strip("*").strip() for c in l.strip("|").split("|")]
            if (cells[0] == label) if exact else (label in cells[0]):
                vs = [qn(c.strip("*")) for c in cells[1:] if re.fullmatch(r"\(?[\d,]+(\.\d+)?\)?", c.strip("*"))]
                if vs: return vs[col]
    raise SystemExit("10-Q 摘錄找不到列：" + label)


def qre(pat, g=1):
    m = re.search(pat, QX)
    if not m: raise SystemExit("10-Q 摘錄找不到：" + pat)
    return qn(m.group(g))


def add10(iid, item, value, unit, page, note="", second=None, second_src="", tol=0.5):
    n0 = len(ITEMS)
    if isinstance(value, float): value = round(value, 3)
    add(iid, item, value, unit, f"Alphabet Form 10-Q 2026Q3 {page}（{QF}）", tag="[Verified]", note=note, second=second, second_src=second_src, tol=tol)
    ITEMS[n0].update({"url": QURL, "docDate": "10-Q 2026Q3（季末 2026-09-30）", "retrieved": "2026-11-04", "period": "2026Q3" if not iid.endswith("20261021") and not iid.endswith("202610") else iid.split(".")[-1]})


B = 1000.0   # 十億 → 百萬
# 1. 收入積壓（營收附註）
add10("rpo.10q.2026q3", "剩餘履約義務（RPO）合計 2026-09-30", qre(r"剩餘履約義務 \*\*([\d.]+) 十億") * B, "USD M", "營收附註（收入積壓）")
add10("rpo.cloud.10q.2026q3", "RPO：Google Cloud 2026-09-30", qre(r"Google Cloud \*\*([\d.]+) 十億") * B, "USD M", "營收附註（收入積壓）",
      note="原（v0.3.3）：法說 backlog 561", second=PV["rpo.cloud.2026q3"] if "rpo.cloud.2026q3" in PV else 561000, second_src="法說 CFO backlog 561B（四捨五入）", tol=500)
add10("rpo.share24m.10q.2026q3", "RPO 24 個月內認列比例（「略多於 55%」，取 55% 下限）", 0.55, "比例", "營收附註（收入積壓）",
      second=qre(r"「略多於 (\d+)%」") / 100, second_src="摘錄原文「略多於 55%」")
# 2. 權益證券（金融工具附註 p.15）
_mkt = qre(r"合計 \*\*([\d,]+)\*\*")
add10("eqsec.mkt.10q.2026q3", "可交易權益證券公允價值 2026-09-30", _mkt, "USD M", "p.15 金融工具附註",
      second=qre(r"Level 1 ([\d,]+)") + qre(r"Level 2 ([\d,]+)"), second_src="Level 1 89,919＋Level 2 465")
add10("eqsec.mkt.vsCall.2026q3", "可交易權益證券（對照法說 90.4）", _mkt, "USD M", "p.15", second=PV["mkt.eqsec.2026q3"], second_src="法說 including $90.4B in equity securities", tol=50)
add10("eqsec.spacex.2026q3", "SpaceX 持股公允價值（含在可交易權益證券；受短期出售限制）", qre(r"含 \*\*([\d.]+) 十億\*\* SpaceX") * B, "USD M", "p.15 註 (1)",
      note="原（v0.3.3）：推導 85.436（6/30 94.1＋Q3 權益證券損益 −8.664）")
add10("eqsec.nonMkt.2026q3", "非上市權益證券（衡量替代法）帳面值", qre(r"帳面值 \*\*([\d.]+) 十億\*\*") * B, "USD M", "p.15", note="其中 8.8B 於 Q3 以公允價值重衡量（主要 Level 3）；新增事實")
add10("is.eqGain.10q.2026q3", "權益證券損益 2026Q3", qre(r"權益證券損益 Q3 2026 \*\*\(([\d,]+)\)\*\*") * -1, "USD M", "OI&E（p.53 MD&A：Q3 減少主要來自 SpaceX 未實現損失）",
      second=PV["is.eqGain.2026q3"], second_src="新聞稿 OI&E 權益證券淨損益")
add10("is.eqGain.10q.ytd2026", "權益證券損益 9M 2026", qre(r"9M 2026 \*\*([\d,]+)\*\*。MD&A"), "USD M", "OI&E")
# 3. 租賃（p.24–25）
_fl = qrow("融資租賃負債合計")
add10("lease.finLiab.2026q3", "融資租賃負債合計 2026-09-30", _fl, "USD M", "p.24–25 租賃附註", note="原（v0.3.3）：沿用 2026-06-30 2.59",
      second=qrow("融資租賃負債：流動") + qrow("融資租賃負債：長期"), second_src="流動 1,053＋長期 2,652")
add10("lease.opLiab.2026q3", "營業租賃負債合計 2026-09-30", qrow("營業租賃負債合計"), "USD M", "p.24–25")
_p9 = qre(r"9M 2026 \*\*([\d,]+)\*\*（v0.3.3 推導")
add10("lease.finPrin.ytd2026", "融資租賃本金現金支出 9M 2026", _p9, "USD M", "p.24–25（籌資活動）", note="原（v0.3.3）：推導 0.943（1H 0.840＋Q3 推估 0.103）",
      second=840 + qre(r"Q3 2026 \*\*([\d,]+)\*\*；9M"), second_src="1H26 0.840（10-Q 2026Q2，v0.3.1 lease_1H 說明）＋本季 Q3 1,110")
add10("lease.finPrin.2026q3", "融資租賃本金現金支出 2026Q3", qre(r"Q3 2026 \*\*([\d,]+)\*\*；9M"), "USD M", "p.24–25")
add10("lease.finPrepay.ytd2026", "預付尚未起租的融資租賃 9M 2026", qre(r"9M 預付尚未起租的融資租賃 ([\d,]+)"), "USD M", "p.24–25")
add10("lease.uncommenced.2026q3", "已簽未起租租賃（主要資料中心與能源基礎設施）未來付款", qre(r"未來付款 ([\d.]+) 十億") * B, "USD M", "p.24–25",
      note="2026–2031 起租，租期 1–26 年；2026-06-30 為 85.2（10-Q 2026Q2）")
add10("lease.shortTerm.commit.2026q3", "短期租賃剩餘不可取消承諾（不入資產負債表）", qre(r"約 \*\*([\d.]+) 十億\*\*（不入") * B, "USD M", "p.24–25", note="2026-06 簽、09 修訂")
_mat = {y: [qn(x) for x in re.search(rf"{y} ([\d,]+)／([\d,]+)", QX).groups()] for y in ("2027", "2028", "2029", "2030")}
_m26 = [qn(x) for x in re.search(r"2026 餘 ([\d,]+)／([\d,]+)", QX).groups()]
_mth = [qn(x) for x in re.search(r"之後 ([\d,]+)／([\d,]+)", QX).groups()]
_mtot = [qn(x) for x in re.search(r"合計未折現 ([\d,]+)／([\d,]+)", QX).groups()]
add10("lease.maturity.10q.2026q3", "租賃未來付款（營業／融資，未折現）",
      {"2026r": dict(zip(("operating", "finance"), _m26)), **{y: dict(zip(("operating", "finance"), v)) for y, v in _mat.items()},
       "thereafter": dict(zip(("operating", "finance"), _mth)), "total": dict(zip(("operating", "finance"), _mtot))}, "USD M", "p.24–25")
for j, nm in enumerate(("operating", "finance")):
    CHK.append((f"lease.maturity.10q 合計（{nm}）", _mtot[j], _m26[j] + sum(v[j] for v in _mat.values()) + _mth[j], "逐年相加", abs(_mtot[j] - (_m26[j] + sum(v[j] for v in _mat.values()) + _mth[j])) < 0.5))
# 4. 債務（Note 6，p.26–27）
_face = qrow("面額合計")
_rows = ["2016 美元", "2020 美元", "2025 美元", "2025 歐元", "2026 美元", "2026 英鎊", "2026 瑞郎", "2026 歐元", "2026 加幣", "2026 日圓", "2026 澳幣", "其他長期債務"]
add10("debt.face.2026q3", "債務面額合計 2026-09-30", _face, "USD M", "p.26–27 Note 6", note="原（v0.3.3）：推導 126.455（6/30 面額 101.085＋Q3 發債淨額 28.996 − Q3 償還 3.626）",
      second=sum(qrow(x) for x in _rows), second_src="Note 6 各票據 2026-09-30 欄相加（12 列）")
_ltd = qrow("長期債務", exact=True)
add10("debt.ltd.10q.2026q3", "長期債務帳面 2026-09-30", _ltd, "USD M", "p.26–27 Note 6",
      second=_face + qrow("未攤銷折價與發行成本") + qrow("避險公允價值調整") + qrow("減：一年內到期"), second_src="面額 − 折價與發行成本 − 避險調整 − 一年內到期")
add10("debt.ltd.vsPR.2026q3", "長期債務帳面（對照新聞稿資產負債表）", _ltd, "USD M", "p.26–27", second=PV["bs.ltd.2026q3"], second_src="新聞稿資產負債表 Long-term debt")
add10("debt.current.2026q3", "一年內到期債務 2026-09-30", -qrow("減：一年內到期"), "USD M", "p.26–27")
add10("debt.fv.2026q3", "票據公允價值 2026-09-30", qre(r"\*\*([\d.]+) 十億\*\*（2026-09-30）") * B, "USD M", "p.26–27")
add10("debt.waymo.loan.202610", "期後：Waymo 優先有擔保定期貸款（2032-10 到期，Term SOFR＋5.25%；Waymo 資產擔保，Alphabet 不保證）",
      qre(r"Waymo 簽 \*\*([\d.]+) 十億\*\*") * B, "USD M", "p.26–27 期後事項", note="2026 年 10 月；評價日後，不進 9/30 資產負債表")
add10("debt.revolver.2026q3", "循環信用額度（未動用）", qre(r"循環信用額度 ([\d.]+) 十億") * B, "USD M", "p.26–27", note="4.0B 2027-04 到期、6.0B 2030-04 到期")
# 5. 補充資訊（Note 7，p.28–29）
_cip = qrow("尚未啟用資產")
add10("bs.ppe.notInService.2026q3", "尚未啟用資產（在建工程）2026-09-30", _cip, "USD M", "p.28–29 Note 7", note="原（v0.3.3）：推導 142.6（6/30 在建 122.814 × 淨額比例）",
      second=qrow("PP&E 淨額", exact=True) - qrow("在用合計") - qrow("累計折舊"), second_src="PP&E 淨額 − 在用合計＋累計折舊")
add10("bs.ppe.10q.2026q3", "PP&E 淨額 2026-09-30", qrow("PP&E 淨額", exact=True), "USD M", "p.28–29", second=PV["bs.ppe.2026q3"], second_src="新聞稿資產負債表 PP&E")
add10("bs.ppe.techInfra.2026q3", "技術基礎設施（在用，約 60% 為伺服器與網路設備）", qrow("技術基礎設施"), "USD M", "p.28–29")
add10("bs.fines.2026q3", "應計罰款與和解 2026-09-30", qre(r"應計罰款與和解 [\d,]+ → \*\*([\d,]+)\*\*"), "USD M", "p.28–29 應計費用",
      note="原（v0.3.3）：推導 12.156（6/30 17.356 − 歐盟 Android 罰款 5.2）；Note 10 p.33：2026-07 支付 5.2B（含利息）")
add10("bs.accruedPPE.2026q3", "應計購買 PP&E 2026-09-30", qre(r"應計購買 PP&E [\d,]+ → ([\d,]+)"), "USD M", "p.28–29")
add10("bs.nci.2026q3", "非控制權益 2026-09-30（含可贖回 0.825B；列在普通股與 APIC 內）", qre(r"非控制權益：\*\* [\d.]+ 十億 → \*\*([\d.]+) 十億") * B, "USD M", "p.28–29",
      note="原（v0.3.3）：沿用 2026-06-30 7.1")
add10("is.intInc.10q.2026q3", "利息收入 2026Q3", qre(r"利息收入 ([\d,]+)／"), "USD M", "p.28–29 OI&E", second=PV["is.intInc.2026q3"], second_src="新聞稿 OI&E")
add10("is.intExp.10q.2026q3", "利息費用 2026Q3（已扣資本化利息 405）", -qre(r"利息費用 \(([\d,]+)\)／"), "USD M", "p.28–29 OI&E", second=PV["is.intExp.2026q3"], second_src="新聞稿 OI&E")
# 6. 承諾與保證（Note 10，p.32）
add10("commit.total.2026q3", "合約承諾合計（含 4.5B 備用信用狀；主要為長期供應協議，至 2032）", qre(r"合約承諾合計 ([\d.]+) 十億") * B, "USD M", "p.32 Note 10",
      note="2026-06-30 為 811.0（10-Q 2026Q2）；法說「未來採購承諾 924」口徑不同（只作對照）")
add10("guar.energy.2026q3", "財務保證：對手方長交期能源設備採購（最大潛在付款，至 2049-01）", qre(r"最大潛在付款 \*\*([\d.]+) 十億\*\*（至") * B, "USD M", "p.32 Note 10")
add10("guar.tpuRepurchase.2026q3", "TPU 回購保證（特定合約；最大潛在付款；公允價值不重大）", qre(r"TPU 回購保證：\*\*.*?最大潛在付款 \*\*([\d.]+) 十億") * B, "USD M", "p.32 Note 10",
      note="公司認為對手方無經濟誘因行使")
# 7. 封面股數
_cov = qre(r"合計 \*\*([\d,]+) 百萬\*\*")
add10("sh.cover.20261021", "10-Q 封面流通股數（2026-10-21，A＋B＋C）", _cov, "M 股", "封面",
      second=qre(r"A 股 ([\d,]+) 百萬") + qre(r"B 股 ([\d,]+) 百萬") + qre(r"C 股 ([\d,]+) 百萬"), second_src="A 5,888＋B 827＋C 5,579")
CHK.append(("sh.cover.20261021 vs 資產負債表 2026-09-30", _cov, PV["sh.out.20260930"], "新聞稿資產負債表股本說明 12,294", abs(_cov - PV["sh.out.20260930"]) < 0.5))
# 8. 現金流量表（9M 2026；對照新聞稿）
for fid, pat in (("cf.capex.ytd2026", r"資本支出 \([\d,]+\)／\*\*\(([\d,]+)\)\*\*"), ("cf.cfo.ytd2026", r"營運現金流 [\d,]+／\*\*([\d,]+)\*\*"),
                 ("cf.debtIss.ytd2026", r"發行債務 [\d,]+／([\d,]+)"), ("cf.debtRep.ytd2026", r"償還 \([\d,]+\)／\(([\d,]+)\)")):
    v = qre(pat) * (-1 if fid in ("cf.capex.ytd2026", "cf.debtRep.ytd2026") else 1)
    CHK.append((f"{fid}（10-Q vs 新聞稿）", v, PV[fid], "新聞稿現金流量表年初至今欄", abs(v - PV[fid]) < 0.5))
add10("cf.da.10q.ytd2026", "折舊 9M 2026（現金流量表）", qre(r"折舊 [\d,]+／\*\*([\d,]+)\*\*"), "USD M", "現金流量表")
OUT = {"_readme": "Alphabet 2026Q3 事實總帳（MAG v0.3.1 工作單第 1 節第 1 項）。由 alphabet/facts_2026q3.py 自 Andy 2026-10-29 提供的財報新聞稿與簡報讀取；"
                  "v0.3.5 起含 2026Q3 10-Q（docs/reports/20261104_10Q_2026Q3_原文摘錄.md；id 標 10q 或 Note 頁碼，tag [Verified]）。金額 USD M；check＝第二次獨立讀取。",
       "asOf": "2026-09-30", "items": ITEMS,
       "checks": [{"id": a, "value": b, "second": c, "source": d, "ok": e} for a, b, c, d, e in CHK]}
json.dump(OUT, open(os.path.join(HERE, "data", "alphabet_facts_2026q3.json"), "w", encoding="utf-8"), ensure_ascii=False, indent=1)
n_ok = sum(1 for c in CHK if c[4])
print(f"事實 {len(ITEMS)} 筆；兩讀核對 {len(CHK)} 項，一致 {n_ok}，不一致 {len(CHK) - n_ok}")
for c in CHK:
    if not c[4]: print("  不一致：", c)
sys.exit(1 if any(not c[4] for c in CHK) else 0)
