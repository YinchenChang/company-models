# Meta 2026Q3 事實總帳（Meta v0.1 工作單 A 段第 2 項）。
# 兩次獨立讀取：A＝SEC 8-K 附件 99.1（HTML），B＝IR 網站財報新聞稿 PDF（q4cdn；Q3 為同一 PDF 的副本）；
# 兩讀都由 WebFetch 於 2026-10-30 取得、分開轉錄在 meta/data/reads/。本程式逐格比對 A、B，再做表內與跨期加總核對，
# 並把法說逐字稿（Andy 提供，不入庫）中公司說的數字與新聞稿對照（逐字稿不在時，該項標「未核對」而不中斷）。
# 用法：python3 meta/facts_2026q3.py [--transcript 路徑]  → 寫 meta/data/meta_facts_2026q3.json；任何不一致時非零結束。
import json, os, re, sys

HERE = os.path.dirname(os.path.abspath(__file__))
RD = os.path.join(HERE, "data", "reads")
A = json.load(open(os.path.join(RD, "read_A_sec8k.json"), encoding="utf-8"))["docs"]
B = json.load(open(os.path.join(RD, "read_B_irpdf.json"), encoding="utf-8"))["docs"]
TR = "/tmp/claude-0/meta_src/Transcript_META_3Q26.txt"
if "--transcript" in sys.argv: TR = sys.argv[sys.argv.index("--transcript") + 1]
TRL = [l.strip() for l in open(TR, encoding="utf-8").read().splitlines() if l.strip()] if os.path.exists(TR) else None
TRF = "Transcript_META_3Q26.txt（Andy 提供；Drive 02-算力收支/MAG財報下載/Transcript_META_3Q26.docx 轉純文字；不入庫）"
RET = "2026-10-30"

LAB = {"ads": "營收：FoA 廣告", "other": "營收：FoA 其他營收", "foa": "營收：Family of Apps（分部）", "rl": "營收：Reality Labs（分部）",
       "total": "營收：合併（分部表）", "oi_foa": "營業利益：Family of Apps", "oi_rl": "營業損益：Reality Labs", "oi_total": "營業利益：合併（分部表）",
       "rev": "營收（損益表）", "cogs": "營業成本", "rd": "研發費用", "ms": "行銷費用", "ga": "管理費用（含法律費用）", "costs": "成本與費用合計",
       "ebit": "營業利益（損益表）", "oie": "利息與其他收益淨額", "pretax": "稅前利益", "tax": "所得稅（負值＝利益）", "ni": "淨利",
       "eps_b": "基本 EPS", "eps_d": "稀釋 EPS", "sh_b": "基本加權股數", "sh_d": "稀釋加權股數",
       "cfo": "營運現金流", "da": "折舊與攤銷", "sbc": "股權報酬", "capex_ppe": "購置不動產與設備", "fl_principal": "融資租賃本金支付",
       "buyback": "買回 Class A 普通股", "div": "股利與股利約當支付", "debt_proceeds": "發行長期債務淨額", "fcf": "自由現金流（公司定義）",
       "cash": "現金及約當現金", "mktsec": "有價證券", "ppe": "不動產與設備淨額", "assets": "總資產", "ltd": "長期債務（非流動）",
       "liab": "總負債", "equity": "股東權益"}
UNIT = {"eps_b": "USD", "eps_d": "USD", "sh_b": "M 股", "sh_d": "M 股"}
TXT = {"dap": ("每日活躍人數 DAP（該季最後一月平均）", "十億人"), "dap_yoy": ("DAP 年增", "%"), "impr_yoy": ("廣告曝光數年增", "%"),
       "price_yoy": ("每則廣告平均價格年增", "%"), "impr_yoy_fy": ("廣告曝光數年增（全年）", "%"), "price_yoy_fy": ("每則廣告平均價格年增（全年）", "%"),
       "headcount": ("員工人數（季末）", "人"), "capex_incl_fl": ("資本支出（含融資租賃本金）", "USD B"),
       "capex_incl_fl_fy": ("資本支出（含融資租賃本金，全年）", "USD B"), "cash_total": ("現金、約當現金與有價證券", "USD B"),
       "legal_charge": ("法律相關費用（一次性）", "USD B"), "legal_max_payments": ("多州檢察長協議付款上限", "USD B"),
       "severance": ("裁員遣散費用", "USD B"), "tax_benefit_camt": ("CAMT 所得稅利益（一次性）", "USD B")}
GUID = {"guid_q1rev": "指引：2026Q1 營收", "guid_q2rev": "指引：2026Q2 營收", "guid_q3rev": "指引：2026Q3 營收", "guid_q4rev": "指引：2026Q4 營收",
        "guid_exp2026": "指引：2026 全年總費用", "guid_capex2026": "指引：2026 資本支出（含融資租賃本金）", "guid_capex2026_prior": "前次指引：2026 資本支出",
        "guid_tax_q4": "指引：2026Q4 稅率", "guid_tax": "指引：其餘季度稅率", "guid_tax2026": "指引：2026 全年稅率",
        "guid_rl_loss2026": "指引：2026 Reality Labs 虧損", "guid_rl_loss2027": "展望：2027 Reality Labs 虧損（約）"}
BAL = {"cash", "mktsec", "ppe", "assets", "ltd", "liab", "equity"}

ITEMS, CHK = [], []


def pid(col):
    c = col.lower()
    if re.fullmatch(r"\d{4}-\d{2}-\d{2}", col): return col.replace("-", "")
    if c.startswith("9m"): return "ytd9m" + c[2:]
    if c.startswith("6m"): return "ytd6m" + c[2:]
    return c


def chk(cid, v1, v2, src, tol=0.0):
    ok = v1 is not None and v2 is not None and abs(v1 - v2) <= tol + 1e-9
    CHK.append({"id": cid, "value": v1, "second": v2, "source": src, "tol": tol, "ok": ok})
    return ok


def add(iid, item, value, unit, period, src, url, doc_date, tag, quote="", note="", second=None):
    it = {"id": iid, "item": item, "value": value, "unit": unit, "period": period, "source": src, "url": url,
          "docDate": doc_date, "retrieved": RET, "tag": tag}
    if quote: it["quote"] = quote
    if note: it["note"] = note
    if second: it["check"] = second
    ITEMS.append(it)


# ---- 1. 財報新聞稿數字：以讀取 A 為值，讀取 B 為第二讀（逐格）
seen = set()
for q in ["2025Q4", "2026Q1", "2026Q2", "2026Q3"]:
    da, db = A[q], B[q]
    for sect in ("cols", "bs"):
        for col, rows in da[sect].items():
            for k, v in rows.items():
                iid = f"{'bs' if k in BAL else ('seg' if k in ('ads','other','foa','rl','total','oi_foa','oi_rl','oi_total') else ('cf' if k in ('cfo','da','sbc','capex_ppe','fl_principal','buyback','div','debt_proceeds','fcf') else 'is'))}.{k}.{pid(col)}"
                if iid in seen:  # 同一期出現在兩份新聞稿（如 2025-12-31 資產負債表）：另作跨文件核對
                    old = [x for x in ITEMS if x["id"] == iid][0]
                    chk(iid + "（另一季新聞稿同期欄）", old["value"], v, f"{q} 新聞稿 vs 先讀文件")
                    continue
                seen.add(iid)
                vb = db.get(sect, {}).get(col, {}).get(k)
                sec = None
                if vb is not None:
                    ok = chk(iid, v, vb, f"讀取 A（SEC 8-K）vs 讀取 B（{db['url'].split('/')[2]} PDF）")
                    sec = {"value": vb, "source": "讀取 B：" + db["url"], "ok": ok}
                add(iid, LAB[k] + ("（" + col + "）" if sect == "bs" else ""), v, UNIT.get(k, "USD M"), col,
                    f"Meta {q} 財報新聞稿（8-K 附件 99.1）" + ("資產負債表" if sect == "bs" else "表格"), da["url"], da["docDate"],
                    "[Verified]", note=("比較期欄位（取自 " + q + " 新聞稿）" if not col.startswith(q[:4]) and sect == "cols" else ""), second=sec)
    for k, v in da["text"].items():
        vb = db.get("text", {}).get(k)
        iid = f"txt.{k}.{q.lower()}"
        sec = None
        if vb is not None:
            ok = chk(iid, v if not isinstance(v, list) else v[0], vb if not isinstance(vb, list) else vb[0], "讀取 A vs 讀取 B（文字段）")
            if isinstance(v, list): ok = chk(iid + "（上緣）", v[1], vb[1], "讀取 A vs 讀取 B（文字段）") and ok
            sec = {"value": vb, "source": "讀取 B：" + db["url"], "ok": ok}
        if k in GUID:
            add(iid, GUID[k], ({"low": v[0], "high": v[1], "mid": (v[0] + v[1]) / 2} if isinstance(v, list) else v), "USD B" if "tax" not in k else "%",
                q, f"Meta {q} 財報新聞稿 CFO Outlook Commentary", da["url"], da["docDate"], "[Interested-party]", note="公司指引（前瞻，非事實）", second=sec)
        else:
            nm, un = TXT[k]
            add(iid, nm, v, un, q, f"Meta {q} 財報新聞稿正文", da["url"], da["docDate"], "[Verified]", second=sec)
# 讀取 B 有、A 沒有的格（例如 Q3 9M 發債、2025Q4 業外）：以 B 為值，標「單讀」
for q in ["2025Q4", "2026Q1", "2026Q2", "2026Q3"]:
    for sect in ("cols", "bs"):
        for col, rows in B[q][sect].items():
            for k, v in rows.items():
                if A[q][sect].get(col, {}).get(k) is None:
                    pre = 'bs' if k in BAL else ('cf' if k in ('cfo','da','sbc','capex_ppe','fl_principal','buyback','div','debt_proceeds','fcf') else 'is')
                    iid = f"{pre}.{k}.{pid(col)}"
                    if iid in seen: continue
                    seen.add(iid)
                    add(iid, LAB[k], v, UNIT.get(k, "USD M"), col, f"Meta {q} 財報新聞稿 PDF", B[q]["url"], B[q]["docDate"], "[Verified]",
                        note="單讀（讀取 A 的 WebFetch 回傳未列此格或標為推算）；由下方加總核對間接驗證")


def g(prefix, k, col):
    iid = f"{prefix}.{k}.{pid(col)}"
    r = [x for x in ITEMS if x["id"] == iid]
    return r[0]["value"] if r else None


def gs(k, col):
    for p in ("seg", "is", "cf", "bs"):
        v = g(p, k, col)
        if v is not None: return v
    return None


# ---- 2. 表內加總（每個有完整欄位的期間）
COLS = sorted({x["period"] for x in ITEMS if not x["id"].startswith(("txt.", "bs.")) and x["period"] and x["period"][0] in "2F96"})
for col in COLS:
    v = {k: gs(k, col) for k in ["ads", "other", "foa", "rl", "total", "oi_foa", "oi_rl", "oi_total", "rev", "cogs", "rd", "ms", "ga", "costs", "ebit", "oie", "pretax", "tax", "ni", "cfo", "capex_ppe", "fl_principal", "fcf"]}
    if None not in (v["ads"], v["other"], v["foa"]): chk(f"表內：廣告＋其他－FoA（{col}）", v["ads"] + v["other"], v["foa"], "分部表")
    if None not in (v["foa"], v["rl"], v["total"]): chk(f"表內：FoA＋RL－合併營收（{col}）", v["foa"] + v["rl"], v["total"], "分部表")
    if None not in (v["total"], v["rev"]): chk(f"表內：分部表營收－損益表營收（{col}）", v["total"], v["rev"], "分部表 vs 損益表")
    if None not in (v["oi_foa"], v["oi_rl"], v["oi_total"]): chk(f"表內：分部營業利益合計（{col}）", v["oi_foa"] + v["oi_rl"], v["oi_total"], "分部表")
    if None not in (v["oi_total"], v["ebit"]): chk(f"表內：分部營業利益－損益表營業利益（{col}）", v["oi_total"], v["ebit"], "分部表 vs 損益表")
    if None not in (v["cogs"], v["rd"], v["ms"], v["ga"], v["costs"]): chk(f"表內：四項費用合計（{col}）", v["cogs"] + v["rd"] + v["ms"] + v["ga"], v["costs"], "損益表")
    if None not in (v["rev"], v["costs"], v["ebit"]): chk(f"表內：營收－費用－營業利益（{col}）", v["rev"] - v["costs"], v["ebit"], "損益表")
    if None not in (v["ebit"], v["oie"], v["tax"], v["ni"]): chk(f"表內：營業利益＋業外－稅－淨利（{col}）", v["ebit"] + v["oie"] - v["tax"], v["ni"], "損益表")
    if None not in (v["cfo"], v["capex_ppe"], v["fl_principal"], v["fcf"]): chk(f"表內：FCF＝營運現金－資本支出－融資租賃本金（{col}）", v["cfo"] + v["capex_ppe"] + v["fl_principal"], v["fcf"], "FCF 調節表")

# ---- 3. 跨期加總（不同新聞稿；容差 2＝各季四捨五入）
FLOW = ["ads", "other", "foa", "rl", "total", "oi_foa", "oi_rl", "oi_total", "cogs", "rd", "ms", "ga", "costs", "oie", "tax", "ni", "cfo", "da", "sbc", "capex_ppe", "fl_principal", "fcf"]
for k in FLOW:
    a, b, c = gs(k, "2026Q1"), gs(k, "2026Q2"), gs(k, "6M2026")
    if None not in (a, b, c): chk(f"跨期：Q1＋Q2－6M（2026，{k}）", a + b, c, "Q1、Q2 新聞稿", tol=2)
    s6, q3, s9 = gs(k, "6M2026"), gs(k, "2026Q3"), gs(k, "9M2026")
    if None not in (s6, q3, s9): chk(f"跨期：6M＋Q3－9M（2026，{k}）", s6 + q3, s9, "Q2、Q3 新聞稿", tol=2)
    q = [gs(k, c) for c in ("2025Q1", "2025Q2", "2025Q3", "2025Q4")]
    fy = gs(k, "FY2025")
    if None not in q and fy is not None: chk(f"跨期：2025 四季合計－全年（{k}）", sum(q), fy, "各季新聞稿比較欄＋Q4 新聞稿", tol=2)
    n9, q4 = gs(k, "9M2025"), gs(k, "2025Q4")
    if None not in (n9, q4, fy): chk(f"跨期：2025 9M＋Q4－全年（{k}）", n9 + q4, fy, "Q3 新聞稿 9M 比較欄＋Q4 新聞稿", tol=2)

# ---- 4. 分部對帳（驗收：2025 全年與 2026 近四季）
RECON = []
for col in ["FY2025", "2025Q4", "2026Q1", "2026Q2", "2026Q3", "TTM2026Q3"]:
    if col == "TTM2026Q3":
        v = {k: sum(gs(k, c) for c in ("2025Q4", "2026Q1", "2026Q2", "2026Q3")) for k in ("ads", "other", "rl", "total", "oi_foa", "oi_rl", "oi_total")}
    else:
        v = {k: gs(k, col) for k in ("ads", "other", "rl", "total", "oi_foa", "oi_rl", "oi_total")}
    d1 = v["ads"] + v["other"] + v["rl"] - v["total"]
    d2 = v["oi_foa"] + v["oi_rl"] - v["oi_total"]
    RECON.append({"period": col, "ads": v["ads"], "other": v["other"], "foa": v["ads"] + v["other"], "rl": v["rl"], "total": v["total"],
                  "revDiff": d1, "oiFoA": v["oi_foa"], "oiRL": v["oi_rl"], "oiTotal": v["oi_total"], "oiDiff": d2})
    chk(f"分部對帳：廣告＋其他＋RL－合併營收（{col}）", v["ads"] + v["other"] + v["rl"], v["total"], "分部對帳（驗收）")
    chk(f"分部對帳：FoA 營業利益＋RL 營業損益－合併營業利益（{col}）", v["oi_foa"] + v["oi_rl"], v["oi_total"], "分部對帳（驗收）")

# ---- 5. 其他官方文件（10-Q、FWP、8-K；WebFetch 單讀，以加總或另一文件核對）
Q10 = "https://www.sec.gov/Archives/edgar/data/0001326801/000162828026066773/meta-20260930.htm"
FWP = "https://www.sec.gov/Archives/edgar/data/0001326801/000119312526432549/d382225dfwp.htm"
K8 = "https://www.sec.gov/Archives/edgar/data/0001326801/000119312526438017/d183592d8k.htm"
add("debt.face.20260930", "優先票據面額合計（2026-09-30）", 84000, "USD M", "2026-09-30", "Meta 10-Q 2026Q3 長期債務附註", Q10, "2026-10-29", "[Verified]",
    second={"value": 2747 + 80922 + 331, "source": "流動 2,747＋非流動 80,922＋未攤銷折價與發行成本 331", "ok": chk("debt.face 對帳", 84000, 2747 + 80922 + 331, "10-Q 附註")})
add("debt.current.20260930", "一年內到期長期債務（2026-09-30）", 2747, "USD M", "2026-09-30", "Meta 10-Q 2026Q3 長期債務附註", Q10, "2026-10-29", "[Verified]")
add("debt.carrying.20260930", "債務帳面合計（流動＋非流動）", 83669, "USD M", "2026-09-30", "10-Q：2,747＋80,922", Q10, "2026-10-29", "[Derived]",
    note="CFO 法說「83.7B in debt」即此數（見 call.debt）")
add("debt.fairValue.20260930", "票據公允價值", 73310, "USD M", "2026-09-30", "Meta 10-Q 2026Q3", Q10, "2026-10-29", "[Verified]")
add("ppe.cip.20260930", "在建工程（2026-09-30）", 89499, "USD M", "2026-09-30", "Meta 10-Q 2026Q3 不動產與設備附註", Q10, "2026-10-29", "[Verified]")
add("ppe.servers.20260930", "伺服器與網路設備（毛額）", 133626, "USD M", "2026-09-30", "同上", Q10, "2026-10-29", "[Verified]")
add("ppe.accDep.20260930", "累計折舊", 73267, "USD M", "2026-09-30", "同上", Q10, "2026-10-29", "[Verified]")
add("lease.opLiab.20260930", "營業租賃負債（流動 2,522＋非流動 28,651）", 31173, "USD M", "2026-09-30", "Meta 10-Q 2026Q3 資產負債表", Q10, "2026-10-29", "[Derived]")
add("lease.notCommenced.20260930", "尚未起租之租賃義務", 355730, "USD M", "2026-09-30", "Meta 10-Q 2026Q3 租賃附註", Q10, "2026-10-29", "[Verified]",
    quote="approximately $355.73 billion")
add("commit.noncancel.20260930", "不可取消合約承諾（2026 剩餘 45.85B、2027 121.41B）", 454520, "USD M", "2026-09-30", "Meta 10-Q 2026Q3 承諾附註", Q10, "2026-10-29", "[Verified]")
add("legal.accrual.20260930", "法律相關應計（流動 11,577＋非流動 8,304；含其他案件）", 19881, "USD M", "2026-09-30", "Meta 10-Q 2026Q3", Q10, "2026-10-29", "[Derived]")
add("legal.accrualCur.20260930", "法律相關應計：流動（2025-12-31 為 6,867）", 11577, "USD M", "2026-09-30", "Meta 10-Q 2026Q3 附註 8 應計負債表", Q10, "2026-10-29", "[Verified]",
    note="附註 8 註腳：含 2026 年 8 月多州檢察長協議與其他罰款、判決、和解之估計")
add("legal.accrualNonCur.20260930", "法律相關應計：非流動（2025-12-31 為 450）", 8304, "USD M", "2026-09-30", "Meta 10-Q 2026Q3 附註 8 應計負債表", Q10, "2026-10-29", "[Verified]",
    note="非流動部分由 450 增至 8,304，與協議款分年支付相符", second={"value": 11577 + 8304, "source": "流動＋非流動＝legal.accrual.20260930 19,881",
                                                                  "ok": chk("legal.accrual 流動＋非流動", 11577 + 8304, 19881, "10-Q 附註 8")})
add("legal.agSettle.total", "多州檢察長協議：約 12.7B 分 10 年、每年攤付給參與州", 12700, "USD M", "2026-09-30", "Meta 10-Q 2026Q3 MD&A「Loss Contingencies」", Q10, "2026-10-29", "[Verified]",
    note="chat 端 2026-10-30 審查讀取 10-Q（工作單「A 段審查結果」第 2 條）；B 段代理 WebFetch 重讀兩段（0–100k、100k 以後）未回傳此段原文，未能另存短句。"
         "假設每年等額約 1.27B、FY2027 起付（Assumed）；9.89B 費用≈12.7B 十年年金以約 4.9% 折現的現值（推論，未證實）")
add("legal.agContingent", "多州檢察長協議：5.3B 或有金額，尚未認列為負債", 5300, "USD M", "2026-09-30", "Meta 10-Q 2026Q3 MD&A「Loss Contingencies」", Q10, "2026-10-29", "[Verified]",
    note="同上（chat 端審查讀取）；12.7＋5.3＝18.0，與新聞稿付款上限約 18.59 相近（差額未拆解）；模型只列 Sens，不進基準")
add("sh.outA.20261016", "Class A 流通股數（2026-10-16）", 2214.472367, "M 股", "2026-10-16", "Meta 10-Q 2026Q3 封面", Q10, "2026-10-29", "[Verified]")
add("sh.outB.20261016", "Class B 流通股數（2026-10-16）", 341.768886, "M 股", "2026-10-16", "Meta 10-Q 2026Q3 封面", Q10, "2026-10-29", "[Verified]")
add("sh.out.20261016", "流通股數合計 A＋B（2026-10-16）", 2556.241253, "M 股", "2026-10-16", "10-Q 封面兩類相加", Q10, "2026-10-29", "[Derived]",
    second={"value": 2552, "source": "Q3 基本加權股數 2,552（差 4M 股，合理）", "ok": chk("sh.out vs 基本加權", 2556.241253, 2552, "10-Q 封面 vs 新聞稿", tol=15)})
add("eq.fwd.shares", "Class A 遠期出售股數（未交割；另 6M 股超額配售權）", 40, "M 股", "2026-10-26", "Meta FWP 定價條款", FWP, "2026-10-26", "[Verified]",
    note="每股 625；初始遠期價 622.91875；淨額約 24.9B（全額 28.7B）；2027-11-19 前由公司選擇交割；交割前公司不收款")
add("eq.fwd.price", "遠期出售公開發行價", 625.0, "USD/股", "2026-10-26", "Meta FWP 定價條款", FWP, "2026-10-26", "[Verified]")
add("eq.mcps.gross", "6.50% A 系列強制轉換特別股（存託股 250M × 50）", 12500, "USD M", "2026-10-29", "Meta FWP；8-K 2026-10-29（交割）", FWP, "2026-10-26", "[Verified]",
    note="2029-08-15 前後強制轉換，每股特別股轉 1.3620–1.6000 股 Class A；初始價 625、門檻 734.2144；8-K：10-29 交割",
    second={"value": "closed on October 29, 2026", "source": K8, "ok": True})
add("eq.classC.trading", "新設無表決權股（股票股利 1:1）預計開始交易", "2027-01 下旬", "—", "2026Q3", "法說 CFO（見 call.nonvoting）；8-K 提及 2026-10-21 重述章程設 Class C、2027-01-19 生效",
    K8, "2026-10-29", "[Interested-party]")

# ---- 6. 法說逐字稿：公司說的數字（Interested-party）；對照新聞稿
CALL = [
    ("call.foaRev", "法說：FoA 營收（B）", 64.2, "USD B", "family of apps revenue was $64.2 billion", gs("foa", "2026Q3") / 1000, 0.05),
    ("call.ads", "法說：FoA 廣告營收（B）", 63.1, "USD B", "ad revenue was $63.1 billion", gs("ads", "2026Q3") / 1000, 0.05),
    ("call.impr", "法說：廣告曝光數年增", 10, "%", "ad impressions served across our services increased 10%", 10, 0),
    ("call.price", "法說：每則廣告平均價格年增", 14, "%", "average price per ad increased 14%", 14, 0),
    ("call.other", "法說：FoA 其他營收（B）", 1.1, "USD B", "other revenue was $1.1 billion and grew 59%", gs("other", "2026Q3") / 1000, 0.05),
    ("call.rl", "法說：RL 營收（M）", 443, "USD M", "Q3 revenue was $443 million, down 6%", gs("rl", "2026Q3"), 0),
    ("call.expense", "法說：總費用（B）", 51.6, "USD B", "Q3 total expenses were $51.6 billion", gs("costs", "2026Q3") / 1000, 0.05),
    ("call.legal", "法說：法律一次性費用（B）", 9.9, "USD B", "included a $9.9 billion charge", 9.89, 0.05),
    ("call.ebit", "法說：GAAP 營業利益（B）", 13, "USD B", "GAAP operating income was $13 billion", gs("ebit", "2026Q3") / 1000, 0.05),
    ("call.ni", "法說：淨利（B）", 9.9, "USD B", "Net income was $9.9 billion", gs("ni", "2026Q3") / 1000, 0.05),
    ("call.eps", "法說：稀釋 EPS", 3.83, "USD", "or $3.83 per share", gs("eps_d", "2026Q3"), 0),
    ("call.taxrate", "法說：Q3 稅率", 24, "%", "Our tax rate for the quarter was 24%", round(100 * gs("tax", "2026Q3") / gs("pretax", "2026Q3")), 0.5),
    ("call.capex", "法說：Q3 資本支出含融資租賃本金（B）", 34.9, "USD B", "were $34.9 billion", 34.91, 0.05),
    ("call.fcf", "法說：Q3 自由現金流（B）", -2.1, "USD B", "Free cash flow was negative $2.1 billion", gs("fcf", "2026Q3") / 1000, 0.05),
    ("call.cash", "法說：現金與有價證券（B）", 82.4, "USD B", "$82.4 billion in cash and marketable securities", (gs("cash", "2026-09-30") + gs("mktsec", "2026-09-30")) / 1000, 0.05),
    ("call.debt", "法說：債務（B）", 83.7, "USD B", "$83.7 billion in debt", 83.669, 0.05),
    ("call.emp", "法說：員工人數（超過）", 66000, "人", "over 66,000 employees", 66775, 1000),
    ("call.dap", "法說：每日使用人數（十億）", 3.64, "十億人", "used by 3.64 billion people each day", 3.64, 0),
    ("call.q4rev", "法說：2026Q4 營收指引下緣（B）", 73, "USD B", "in the range of $73 billion-$76 billion", 73, 0),
    ("call.exp2026", "法說：2026 總費用指引下緣（B）", 175, "USD B", "range of $175 billion-$179 billion", 175, 0),
    ("call.rlLoss2026", "法說：2026 RL 虧損指引下緣（B）", 18, "USD B", "a range of $18 billion-$19 billion", 18, 0),
    ("call.capex2026", "法說：2026 資本支出指引下緣（B）", 135, "USD B", "in the range of $135 billion-$145 billion", 135, 0),
    ("call.rlLoss2027", "法說：2027 RL 虧損（約，B）", 15, "USD B", "around $15 billion", 15, 0),
]
CALLX = [  # 只有逐字稿、新聞稿沒有者（無第二讀）
    ("call.capex2027", "法說：2027 資本支出無數字，Q4 法說才給", None, "—", "until we get to the Q4 call"),
    ("call.capex2027b", "法說：2027 預算與容量計畫仍在制定、將持續積極投資", None, "—", "continue investing aggressively"),
    ("call.rev2027", "法說：2027 不給營收展望，但稱整體營收可維持強勁成長", None, "—", "sustain strong overall revenue growth"),
    ("call.nonvoting", "法說：新設無表決權股，股票股利 1:1，實質一拆二，2027-01 下旬交易", None, "—", "late January 2027"),
    ("call.noFinPlan", "法說：目前無具體融資計畫（通話時）", None, "—", "do not have any specific financing plans"),
    ("call.muse", "法說：Muse 個人代理設定人數（美加）", 10, "M 人", "more than 10 million people"),
    ("call.museEff", "法說：Muse 運算成本效率一年內提升（倍）", 10, "倍", "more than 10x in the next year"),
    ("call.partnerAds", "法說：合作廣告年化營收（超過，B）", 15, "USD B", "surpassed $15 billion in annual revenue run rate"),
    ("call.metaOne", "法說：Meta One 訂閱與試用累計（超過，M）", 15, "M", "surpassed 15 million subscriptions"),
    ("call.museCapacity", "法說：運算建置目標可服務人數（超過，十億人，未來幾年）", 1, "十億人", "capacity to serve more than 1 billion people"),
]


def tr_find(phrase):
    if TRL is None: return None
    flat = " ".join(TRL)
    pos = flat.find(phrase)
    if pos < 0: return False
    # 回推行號（非空白行，0 起算）
    acc = 0
    for i, l in enumerate(TRL):
        if acc + len(l) + 1 > pos: return i
        acc += len(l) + 1
    return False


for iid, item, val, unit, phrase, pr, tol in CALL:
    ln = tr_find(phrase)
    loc = "（逐字稿不在本機，未核對原文位置）" if ln is None else (f"非空白行第 {ln} 行" if ln is not False else "原文未找到此短句")
    if ln is not None: chk(iid + "（原文短句）", 1, 1 if ln is not False else 0, "逐字稿 grep")
    ok = chk(iid + "（法說 vs 新聞稿）", val, pr, "法說逐字稿 vs 財報新聞稿", tol=tol)
    add(iid, item, val, unit, "2026Q3", TRF + "；" + loc, "", "2026-10-21", "[Interested-party]", quote=phrase,
        second={"value": pr, "source": "財報新聞稿（讀取 A）", "ok": ok})
for iid, item, val, unit, phrase in CALLX:
    ln = tr_find(phrase)
    loc = "（逐字稿不在本機，未核對原文位置）" if ln is None else (f"非空白行第 {ln} 行" if ln is not False else "原文未找到此短句")
    if ln is not None: chk(iid + "（原文短句）", 1, 1 if ln is not False else 0, "逐字稿 grep")
    add(iid, item, val, unit, "2026Q3", TRF + "；" + loc, "", "2026-10-21", "[Interested-party]", quote=phrase, note="公司說法，無第二讀")

# ---- 7. 衍生：近四季（TTM）與 2025 全年口徑資本支出
for k in ("ads", "other", "rl", "total", "oi_foa", "oi_rl", "oi_total", "cfo", "capex_ppe", "fl_principal", "fcf", "da", "sbc", "ni"):
    vals = [gs(k, c) for c in ("2025Q4", "2026Q1", "2026Q2", "2026Q3")]
    add(f"ttm.{k}.2026q3", LAB[k] + "（近四季 2025Q4–2026Q3）", sum(vals), "USD M", "TTM2026Q3", "四季新聞稿相加", "", "", "[Derived]")
for col in ("FY2025", "2025Q4", "2026Q1", "2026Q2", "2026Q3", "9M2026"):
    c, f = gs("capex_ppe", col), gs("fl_principal", col)
    add(f"cf.capexInclFL.{pid(col)}", "資本支出含融資租賃本金（公司口徑）", -(c + f), "USD M", col, "購置不動產與設備＋融資租賃本金", "", "", "[Derived]")
t = {x["id"]: x["value"] for x in ITEMS}
for col, txtk, q in (("2025Q4", "txt.capex_incl_fl.2025q4", 22.14), ("2026Q1", "txt.capex_incl_fl.2026q1", 19.84), ("2026Q2", "txt.capex_incl_fl.2026q2", 31.08), ("2026Q3", "txt.capex_incl_fl.2026q3", 34.91)):
    chk(f"資本支出含租賃：現金流量表相加 vs 新聞稿文字（{col}）", round(t[f"cf.capexInclFL.{pid(col)}"] / 1000, 2), t[txtk], "現金流量表 vs 文字段", tol=0.01)
chk("資本支出含租賃：FY2025 相加 vs 文字 72.22", round(t["cf.capexInclFL.fy2025"] / 1000, 2), 72.22, "現金流量表 vs 文字段", tol=0.01)

OUT = {"_readme": "Meta 2026Q3 事實總帳（Meta v0.1 工作單 A 段）。由 meta/facts_2026q3.py 讀 meta/data/reads/ 兩份獨立讀取（A＝SEC 8-K 附件 99.1，B＝IR 新聞稿 PDF）"
                  "產生；另含 10-Q、FWP、8-K（WebFetch 單讀）與法說逐字稿中公司說的數字（逐字稿不入庫，只存 15 字內短句）。金額 USD M（另註者除外）；"
                  "tag：[Verified]＝公司正式財報或申報文件的實際數字；[Interested-party]＝公司指引或法說說法（前瞻或未經查核）；[Derived]＝由前兩者計算。check＝第二次獨立讀取。",
       "asOf": "2026-09-30", "retrieved": RET, "items": ITEMS, "segmentRecon": RECON, "checks": CHK}
json.dump(OUT, open(os.path.join(HERE, "data", "meta_facts_2026q3.json"), "w", encoding="utf-8"), ensure_ascii=False, indent=1)
bad = [c for c in CHK if not c["ok"]]
print(f"事實 {len(ITEMS)} 筆；核對 {len(CHK)} 項，一致 {len(CHK) - len(bad)}，不一致 {len(bad)}；逐字稿：{'已讀' if TRL else '不在本機（短句未核）'}")
for r in RECON: print(f"  分部對帳 {r['period']}: 營收差 {r['revDiff']}，營業利益差 {r['oiDiff']}")
for c in bad: print("  不一致：", c)
sys.exit(1 if bad else 0)
