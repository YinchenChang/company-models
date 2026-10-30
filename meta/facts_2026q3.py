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
TR = os.environ.get("META_TR", "/tmp/claude-0/meta_src/Transcript_META_3Q26.txt")
if "--transcript" in sys.argv: TR = sys.argv[sys.argv.index("--transcript") + 1]
TR2 = os.environ.get("META_TR2", "/tmp/claude-0/meta_src/Transcript_follow up call_META_3Q26.txt")   # v0.1.1：follow-up call 逐字稿（不入庫）
def _lines(path): return [l.strip() for l in open(path, encoding="utf-8").read().splitlines() if l.strip()] if os.path.exists(path) else None
TRL, TRL2 = _lines(TR), _lines(TR2)
TRF = "Transcript_META_3Q26.txt（Andy 提供；Drive 02-算力收支/MAG財報下載/Transcript_META_3Q26.pdf 轉純文字；不入庫）"
TRF2 = "Transcript_follow up call_META_3Q26.txt（Andy 提供；Drive 02-算力收支/MAG財報下載/Transcript_follow up call_META_3Q26.pdf 轉純文字；不入庫）"
# v0.1.1：完整 10-Q（106 頁，SEC 公開，可入庫）：PDF 與 pdftotext -layout 純文字（換頁符 \f 分頁）
Q10TXT = os.path.join(HERE, "..", "mag", "extract", "docs", "2026Q3", "text", "meta_2026Q3_10-Q.txt")
Q10PDF = "mag/extract/docs/2026Q3/meta_2026Q3_10-Q.pdf"
Q10PAGES = open(Q10TXT, encoding="utf-8").read().split("\f")
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
    quote="payable to participating states in annual installments over a 10-year period",
    note="chat 端 2026-10-30 審查讀取 10-Q（工作單「A 段審查結果」第 2 條）；v0.1.1 以完整 10-Q 核對（PDF 第 26、51、58 頁），原文未載各年金額。"
         "模型假設每年等額約 1.27B、FY2027 起付（H3）；v0.1 推論「9.89B≈12.7B 以約 4.9% 折現」已撤回：協議費用前三季合計 11.89B（legal.agCharge.ytd9m2026）")
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
    ("call.foaRev", "法說：FoA 營收（B）", 64.2, "USD B", "Family of Apps revenue was $64.2 billion", gs("foa", "2026Q3") / 1000, 0.05),
    ("call.ads", "法說：FoA 廣告營收（B）", 63.1, "USD B", "ad revenue was $63.1 billion", gs("ads", "2026Q3") / 1000, 0.05),
    ("call.impr", "法說：廣告曝光數年增", 10, "%", "ad impressions served across our services increased 10%", 10, 0),
    ("call.price", "法說：每則廣告平均價格年增", 14, "%", "average price per ad increased 14%", 14, 0),
    ("call.other", "法說：FoA 其他營收（B）", 1.1, "USD B", "other revenue was $1.1 billion and grew 59%", gs("other", "2026Q3") / 1000, 0.05),
    ("call.rl", "法說：RL 營收（M）", 443, "USD M", "Q3 revenue was $443 million, down 6%", gs("rl", "2026Q3"), 0),
    ("call.expense", "法說：總費用（B）", 51.6, "USD B", "Q3 total expenses were $51.6 billion", gs("costs", "2026Q3") / 1000, 0.05),
    ("call.legal", "法說：法律一次性費用（B）", 9.9, "USD B", "included a $9.9 billion charge", 9.89, 0.05),
    ("call.ebit", "法說：GAAP 營業利益（B）", 13, "USD B", "GAAP operating income was $13.0 billion", gs("ebit", "2026Q3") / 1000, 0.05),
    ("call.ni", "法說：淨利（B）", 9.9, "USD B", "Net income was $9.9 billion", gs("ni", "2026Q3") / 1000, 0.05),
    ("call.eps", "法說：稀釋 EPS", 3.83, "USD", "or $3.83 per share", gs("eps_d", "2026Q3"), 0),
    ("call.taxrate", "法說：Q3 稅率", 24, "%", "Our tax rate for the quarter was 24%", round(100 * gs("tax", "2026Q3") / gs("pretax", "2026Q3")), 0.5),
    ("call.capex", "法說：Q3 資本支出含融資租賃本金（B）", 34.9, "USD B", "were $34.9 billion", 34.91, 0.05),
    ("call.fcf", "法說：Q3 自由現金流（B）", -2.1, "USD B", "Free cash flow was negative $2.1 billion", gs("fcf", "2026Q3") / 1000, 0.05),
    ("call.cash", "法說：現金與有價證券（B）", 82.4, "USD B", "$82.4 billion in cash and marketable securities", (gs("cash", "2026-09-30") + gs("mktsec", "2026-09-30")) / 1000, 0.05),
    ("call.debt", "法說：債務（B）", 83.7, "USD B", "$83.7 billion in debt", 83.669, 0.05),
    ("call.emp", "法說：員工人數（超過）", 66000, "人", "over 66,000 employees", 66775, 1000),
    ("call.dap", "法說：每日使用人數（十億）", 3.64, "十億人", "used by 3.64 billion people each day", 3.64, 0),
    ("call.q4rev", "法說：2026Q4 營收指引下緣（B）", 73, "USD B", "in the range of $73-76 billion", 73, 0),
    ("call.exp2026", "法說：2026 總費用指引下緣（B）", 175, "USD B", "range of $175-179 billion", 175, 0),
    ("call.rlLoss2026", "法說：2026 RL 虧損指引下緣（B）", 18, "USD B", "a range of $18-19 billion", 18, 0),
    ("call.capex2026", "法說：2026 資本支出指引下緣（B）", 135, "USD B", "the range of $135-145 billion", 135, 0),
    ("call.rlLoss2027", "法說：2027 RL 虧損（約，B）", 15, "USD B", "around $15 billion", 15, 0),
]
CALLX = [  # 只有逐字稿、新聞稿沒有者（無第二讀）
    ("call.capex2027", "法說：2027 資本支出無數字，Q4 法說才給", None, "—", "until we get to the Q4 call"),
    ("call.capex2027b", "法說：2027 預算與容量計畫仍在制定、將持續積極投資", None, "—", "continue investing aggressively"),
    ("call.rev2027", "法說：2027 不給營收展望，但稱整體營收可維持強勁成長", None, "—", "sustain strong overall revenue growth"),
    ("call.nonvoting", "法說：新設無表決權股，股票股利 1:1，實質一拆二，2027-01 下旬交易", None, "—", "late January 2027"),
    ("call.noFinPlan", "法說：目前無具體融資計畫（通話時）", None, "—", "don't have any specific financing plans"),
    ("call.muse", "法說：Muse 個人代理設定人數（美加）", 10, "M 人", "more than 10 million people"),
    ("call.museEff", "法說：Muse 運算成本效率一年內提升（倍）", 10, "倍", "more than 10x in the next year"),
    ("call.partnerAds", "法說：合作廣告年化營收（超過，B）", 15, "USD B", "surpassed $15 billion in annual revenue run-rate"),
    ("call.metaOne", "法說：Meta One 訂閱與試用累計（超過，M）", 15, "M", "surpassed 15 million subscriptions"),
    ("call.museCapacity", "法說：運算建置目標可服務人數（超過，十億人，未來幾年）", 1, "十億人", "capacity to serve more than a billion people"),
]


def _norm(t): return re.sub(r"\s+", " ", t.replace("’", "'").replace("‘", "'")).lower()


def tr_find(phrase, lines=None):
    """短句在逐字稿中的位置（非空白行，0 起算）；比對時不分大小寫、空白合併、彎引號視同直引號（v0.1.1：逐字稿改用 PDF 轉文字版）。"""
    lines = TRL if lines is None else lines
    if lines is None: return None
    ph = _norm(phrase)
    acc = 0
    for i, l in enumerate(lines):
        # 從第 i 行開始接上後續 3 行比對（短句可能跨行）
        if ph in _norm(" ".join(lines[i:i + 4])) and ph not in _norm(" ".join(lines[i + 1:i + 4])): return i
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

# ---- 8. v0.1.1：完整 10-Q 原文（repo 內純文字；第二讀＝程式從原文錨點後擷取數字，與轉錄值比對；頁碼＝PDF 頁）
def q10_find(anchor, nth=0):
    """回傳 (PDF 頁碼, 錨點後的文字)；錨點為正規表示式，比對時空白合併。"""
    hits = []
    for pi, pg in enumerate(Q10PAGES, 1):
        flat = re.sub(r"\s+", " ", pg)
        for m in re.finditer(anchor, flat): hits.append((pi, flat[m.end():m.end() + 160], m.group(0)))
    return hits[nth] if len(hits) > nth else (None, "", "")


def q10_num(after, k=0):
    ns = re.findall(r"\(?\$?\s*([0-9][0-9,]*(?:\.[0-9]+)?)\)?", after)
    return float(ns[k].replace(",", "")) if len(ns) > k else None


Q10F = [  # 鍵、項目、轉錄值、單位、期間、錨點（regex）、取錨點後第 k 個數字、標記、說明
    ("legal.accrualCur.20251231", "法律相關應計：流動（2025-12-31）", 6867, "USD M", "2025-12-31", r"Legal-related accruals, current \(?1?\)?", 1, "[Verified]", "2026-09-30 為 11,577（+4,710）"),
    ("legal.accrualNonCur.20251231", "法律相關應計：非流動（2025-12-31）", 450, "USD M", "2025-12-31", r"Legal-related accruals, non-current \(?1?\)?", 1, "[Verified]", "2026-09-30 為 8,304（+7,854）"),
    ("legal.agCharge.ytd9m2026", "多州檢察長協議費用：2026 年前三季合計（Q3 9.89）", 11.89, "USD B", "9M2026", r"attorneys general agreement of \$9\.89 billion in the three months ended September 30, 2026 and", 0, "[Verified]",
     "v0.1 以 9.89 當作協議全部應計；10-Q：前三季合計 11.89（即 Q3 以前已認列約 2.00）→ 其他法律應計＝19.881 − 11.89＝7.991"),
    ("legal.agCommittee", "多州檢察長協議：另付各州委員會調查與訴訟費用", 75, "USD M", "2026-09-30", r"An additional", 0, "[Verified]", "未入模型（金額小）"),
    ("legal.agPrivacy", "協議總付款含隱私案約 500M（49 州）", 500, "USD M", "2026-09-30", r"we agreed to payments of approximately", 0, "[Verified]",
     "12.7（70%）＋5.3（30%）＋0.075＋約 0.5 ≈ 18.59 上限；0.5 是否已含在 12.7 內原文未明"),
    ("legal.agPct", "12.7B 占青少年案總付款比例", 70, "%", "2026-09-30", r"Approximately \$12\.7 billion, representing approximately", 0, "[Verified]", "原文「annual installments over a 10-year period」，未載各年金額"),
    ("fl.rou.20260930", "融資租賃使用權資產（不動產與設備內）", 11704, "USD M", "2026-09-30", r"Finance lease right-of-use assets", 0, "[Verified]", "10-Q 未單獨揭露融資租賃負債（資產負債表無此科目；附註 8 其他流動／非流動負債未拆）"),
    ("fl.rou.20251231", "融資租賃使用權資產（2025-12-31）", 8187, "USD M", "2025-12-31", r"Finance lease right-of-use assets", 1, "[Verified]", ""),
    ("fl.principal.ytd9m2026.q10", "融資租賃本金支付（10-Q 現金流量表，前三季）", 2769, "USD M", "9M2026", r"Principal payments on finance leases", 0, "[Verified]", "與新聞稿 cf.fl_principal.ytd9m2026 相同"),
    ("bb.remaining.20260930", "買回授權剩餘額度（2025-12-31 與 2026-09-30 皆同）", 25.03, "USD B", "2026-09-30", r"2025 and September 30, 2026,", 0, "[Verified]", "2026 年前三季未買回"),
    ("ppe.dep.2026q3", "不動產與設備折舊費用（2026Q3）", 6.81, "USD B", "2026Q3", r"Depreciation expense on property and equipment was", 0, "[Verified]", "前三季 18.49"),
    ("ppe.depServers.2026q3", "其中伺服器與網路設備折舊（2026Q3）", 5.34, "USD B", "2026Q3", r"servers and network assets depreciation expense was", 0, "[Verified]", "前三季 14.34；2025Q3 3.60"),
    ("ppe.depServers.ytd9m2026", "伺服器與網路設備折舊（前三季）", 14.34, "USD B", "9M2026", r"servers and network assets depreciation expense was \$5\.34 billion and \$3\.60 billion for the three months ended September 30, 2026 and 2025, respectively, and", 0, "[Verified]", ""),
    ("ppe.servers.20251231", "伺服器與網路設備（毛額，2025-12-31）", 98040, "USD M", "2025-12-31", r"Servers and network assets", 1, "[Verified]", ""),
    ("cash.restricted.20260930", "受限約當現金（不含在現金＋有價證券 82.45 內）", 12826, "USD M", "2026-09-30", r"Restricted cash equivalents", 0, "[Verified]", "其中 10.80B 為基礎設施採購合約之保證金，預計 2028–2030 釋出"),
    ("cash.escrow.20260930", "基礎設施採購合約保證金（貨幣市場基金轉受限）", 10.80, "USD B", "2026-09-30", r"multi-year infrastructure purchase agreements,", 0, "[Verified]", "只列 Recon；不計入現金"),
    ("commit.cloudContingent", "或有雲端算力採購義務（五年，最多）", 14.72, "USD B", "2026-09-30", r"contingent obligations to purchase up to", 0, "[Verified]", "雲端供應商轉售他人則可減少"),
    ("commit.due2026", "不可取消合約承諾：2026 剩餘到期", 45.85, "USD B", "2026-09-30", r"respectively\. In addition|with approximately", 0, "[Verified]", "2027 到期 121.41"),
    ("debt.mat2027", "票據本金到期：2027", 2750, "USD M", "2026-09-30", r"Remainder of 2026 \$ — 2027", 0, "[Verified]", "2028 1,500；2029 1,000；2030 5,000；其後 73,750"),
    ("debt.interestFuture", "未來利息支付義務：一年內（之後 84.51B）", 4.40, "USD B", "2026-09-30", r"future interest payments obligations as of September 30, 2026 were", 0, "[Verified]", ""),
    ("hfs.20260930", "待出售資產淨額（預付及其他流動資產內）", 2.60, "USD B", "2026-09-30", r"total held-for-sale assets, net, were", 0, "[Verified]", "資料中心資產，預計一年內投入合資"),
    ("sh.classC.effective", "Class C 生效日（重述章程）", 2027, "年", "2026-10", r"with an effective date of January 19,", 0, "[Verified]", "2027-01-19；股票股利 1:1、會計上視同股票分割"),
]
Q10CHK = []
for iid, item, val, unit, period, anc, k, tag, note in Q10F:
    pg, after, hit = q10_find(anc)
    v2 = q10_num(after, k) if pg else None
    if iid == "debt.mat2027": pg, after, hit = q10_find(r"2027 2,750"); v2 = 2750.0 if pg else None
    if iid == "commit.due2026": pg, after, hit = q10_find(r"with approximately \$45\.85 billion and"); v2 = 45.85 if pg else None
    if iid == "sh.classC.effective": v2 = q10_num(after, 0)
    ok = chk(iid + "（完整 10-Q 原文）", val, v2, f"轉錄值 vs 程式自 {Q10PDF} 純文字擷取", tol=1e-6)
    add(iid, item, val, unit, period, f"Meta 10-Q 2026Q3（完整 PDF 第 {pg} 頁；{Q10PDF}）", Q10, "2026-10-29", tag,
        quote=" ".join(hit.split()[:12]), note=note, second={"value": v2, "source": f"完整 10-Q 純文字第 {pg} 頁（程式擷取）", "ok": ok})
    Q10CHK.append((iid, pg))
# v0.1 的 10-Q 事實（WebFetch 單讀）以完整 10-Q 再核對一次（工作單第 0 節第 6 項）
for iid, anc, k, scale in (("debt.face.20260930", r"Total face value of long-term debt", 0, 1), ("debt.current.20260930", r"Less: Long-term debt, current", 0, 1),
                           ("debt.fairValue.20260930", r"total estimated fair value of our outstanding Notes was", 0, 1000),
                           ("ppe.cip.20260930", r"Construction in progress \(1\)", 0, 1), ("ppe.servers.20260930", r"Servers and network assets", 0, 1),
                           ("ppe.accDep.20260930", r"Less: Accumulated depreciation", 0, 1), ("lease.notCommenced.20260930", r"These lease obligations were approximately", 0, 1000),
                           ("commit.noncancel.20260930", r"we had \$", 0, 1000), ("legal.accrualCur.20260930", r"Legal-related accruals, current \(?1?\)?", 0, 1),
                           ("legal.accrualNonCur.20260930", r"Legal-related accruals, non-current \(?1?\)?", 0, 1),
                           ("legal.agSettle.total", r"Approximately \$", 0, 1000), ("legal.agContingent", r"The remaining approximately \$", 0, 1000),
                           ("sh.outA.20261016", r"Class A Common Stock \$0\.000006 par value", 0, 1e-6), ("sh.outB.20261016", r"Class B Common Stock \$0\.000006 par value", 0, 1e-6)):
    it = [x for x in ITEMS if x["id"] == iid][0]
    pg, after, hit = q10_find(anc)
    v2 = q10_num(after, k)
    v2 = None if v2 is None else round(v2 * scale, 6)
    ok = chk(iid + "（v0.1 事實 vs 完整 10-Q）", it["value"], v2, "v0.1 WebFetch 讀值 vs 完整 10-Q 純文字", tol=1e-6)
    it["fullTenQ"] = {"page": pg, "value": v2, "ok": ok, "source": Q10PDF}

# ---- 8b. v0.1.1：424B5（強制轉換特別股）承銷折扣 → 淨額（WebFetch 單讀；算術核對）
B5 = "https://www.sec.gov/Archives/edgar/data/0001326801/000119312526436405/d133631d424b5.htm"
add("eq.mcps.uwDisc", "強制轉換特別股承銷折扣（每存託股 0.1665；250M 股）", 41.625, "USD M", "2026-10-27", "Meta 424B5（Series A 存託股）封面", B5, "2026-10-27", "[Verified]",
    quote="Underwriting discount $0.1665", note="2026-10-30 WebFetch 讀取（工具端）；公司稱淨額約 12.5B（超額配售全數行使約 14.3B，未知是否行使，不計）",
    second={"value": round(0.1665 * 250, 4), "source": "0.1665 × 250M 股", "ok": chk("eq.mcps.uwDisc 算術", 41.625, round(0.1665 * 250, 4), "424B5 封面")})
add("eq.mcps.net", "強制轉換特別股淨額（毛額 12,500 − 承銷折扣 41.625；費用未扣）", 12458.375, "USD M", "2026-10-29", "424B5 封面推算", B5, "2026-10-27", "[Derived]")
add("eq.fwd.uwDisc", "遠期出售承銷折扣（每股 2.08125；初始遠期價 622.91875）", 2.08125, "USD/股", "2026-10-26", "Meta 424B5（Class A，遠期出售）封面",
    "https://www.sec.gov/Archives/edgar/data/0001326801/000119312526436410/d94886d424b5.htm", "2026-10-27", "[Verified]",
    quote="subject to adjustment on a daily basis", note="遠期價每日依隔夜銀行融資利率減利差調整，並按預期股利調降（模型不計增值，保守）",
    second={"value": round(625 - 622.91875, 5), "source": "625 − 622.91875", "ok": chk("eq.fwd.uwDisc 算術", 2.08125, round(625 - 622.91875, 5), "424B5 封面")})

# ---- 8c. v0.1.1：follow-up call 逐字稿（Interested-party；只存 15 字內短句；不入庫）
CALLF = [
    ("fu.capex2027", "follow-up：2027 資本支出未量化（主法說：Q4 法說才給）", None, "—", "until we get to the Q4 call", TRL),
    ("fu.rl2027", "法說：RL 2027 虧損約 15B 並走向獲利（已在 v0.1 rl_loss27）", 15, "USD B", "trending towards profitability", TRL),
    ("fu.thirdCloudCost", "法說：基礎設施成本成長含第三方雲端支出", None, "—", "higher third party cloud spend", TRL),
    ("fu.tokenCost", "法說：費用成長含第三方 AI token 成本", None, "—", "third-party AI token costs", TRL),
    ("fu.ownInfra", "follow-up：以自建基礎設施為主", None, "—", "building our own infrastructure", TRL2),
    ("fu.cloudAug", "follow-up：近期以第三方雲端合約補足容量", None, "—", "augmented that with third-party cloud agreements", TRL2),
    ("fu.mixNotSpec", "follow-up：自建與租用比例未揭露", None, "—", "I don't think we've specified the mix", TRL2),
    ("fu.constrained", "follow-up：算力供給仍吃緊", None, "—", "remains pretty constrained", TRL2),
    ("fu.paceOrders", "follow-up：伺服器訂單可分期調節", None, "—", "pace our server orders over time", TRL2),
    ("fu.classCmax", "follow-up：資訊聲明中「數千億」是理論上限，非融資計畫", None, "—", "a theoretical maximum", TRL2),
    ("fu.noEquityPlan", "follow-up：非目前股權融資需求或計畫", None, "—", "not actually a description of our present equity financing needs", TRL2),
    ("fu.eff10x", "follow-up：10 倍效率仍需技術進展", None, "—", "10x is still going to require more technological advances", TRL2),
    ("fu.mtia", "法說與 follow-up：未提及 MTIA／自研晶片（搜尋 MTIA、silicon、chip、ASIC 皆無）", None, "—", "", None),
]
for iid, item, val, unit, phrase, lines in CALLF:
    if not phrase:
        add(iid, item, val, unit, "2026Q3", TRF + "；" + TRF2, "", "2026-10-21", "[Interested-party]", note="已搜尋，未找到"); continue
    ln = tr_find(phrase, lines)
    loc = "（逐字稿不在本機，未核對原文位置）" if ln is None else (f"非空白行第 {ln} 行" if ln is not False else "原文未找到此短句")
    if ln is not None: chk(iid + "（原文短句）", 1, 1 if ln is not False else 0, "逐字稿 grep")
    add(iid, item, val, unit, "2026Q3", (TRF if lines is TRL else TRF2) + "；" + loc, "", "2026-10-21", "[Interested-party]", quote=phrase, note="公司說法，無第二讀；方向性，未量化")

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
