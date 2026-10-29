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
OUT = {"_readme": "Alphabet 2026Q3 事實總帳（MAG v0.3.1 工作單第 1 節第 1 項）。由 alphabet/facts_2026q3.py 自 Andy 2026-10-29 提供的財報新聞稿與簡報讀取；"
                  "10-Q 尚未提供：10-Q 附註才有的數字（債務面額、租賃、在建工程、權益證券明細）見 build_alphabet.py 的推導與資料缺口。金額 USD M；check＝第二次獨立讀取。",
       "asOf": "2026-09-30", "items": ITEMS,
       "checks": [{"id": a, "value": b, "second": c, "source": d, "ok": e} for a, b, c, d, e in CHK]}
json.dump(OUT, open(os.path.join(HERE, "data", "alphabet_facts_2026q3.json"), "w", encoding="utf-8"), ensure_ascii=False, indent=1)
n_ok = sum(1 for c in CHK if c[4])
print(f"事實 {len(ITEMS)} 筆；兩讀核對 {len(CHK)} 項，一致 {n_ok}，不一致 {len(CHK) - n_ok}")
for c in CHK:
    if not c[4]: print("  不一致：", c)
sys.exit(1 if any(not c[4] for c in CHK) else 0)
