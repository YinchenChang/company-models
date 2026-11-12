# NVIDIA v0.3.5／v0.3.6（2026-11-11）後處理：NVIDIA v0.3.4 驗收採用規劃書（Andy 2026-11-11 07:40 確認「請按預設繼續」：J1 a、J2 a）。
# 原則：docs/plan/個股模型原則_命題加營收分部.md 第 10 節第 1–51 條；不需驗收清單 docs/plan/個股模型驗收_不需驗收項目.md v1.5（X 列不做）。
# build_nvidia.py 在存檔前呼叫 apply(wb, here, ctx)。ctx["V36"]＝False 時只做展示與治理（數值不變，v0.3.5）；True 時加做數值項（v0.3.6）。
#   v0.3.5：原則 1、2、3、4、5、6、7、10、13、14、15、16、18、20、24、25、30、31、32、33、34、36、40（現金稅率並列）、41、42、43、44、45、47、50（展示）、51
#   v0.3.6：原則 11（評價股數）、26（目標時點）、35（折舊與資本支出比率最近四季）、49（月數加權；build_nvidia.py 引擎）、48（k_rec 衰減 Sens）、
#           50（Recon 毛利率差距分解與 abs_mem 反解）、23（供應預付與承諾另列）
import hashlib, json, os, re
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.workbook.defined_name import DefinedName

BOLD = Font(bold=True); BLUE = Font(color="0000FF"); HDR = PatternFill("solid", fgColor="DDEBF7")
AS_OF, CUTOFF, BUILT, VAL_D = "2026-11-05", "2026-11-05", "2026-11-11", "2026-12-31"
SEC = "https://www.sec.gov/Archives/edgar/data/1045810/"
Q10 = SEC + "000104581026000075/nvda-20260726.htm"
Q10N = "NVIDIA 10-Q（期間結束 2026-07-26；申報 2026-08-26；accession 0001045810-26-000075）"
K10 = SEC + "000104581026000021/nvda-20260125.htm"
K10N = "NVIDIA 10-K（FY2026，期間結束 2026-01-25；申報 2026-02-25；accession 0001045810-26-000021）"
Q3_10Q = "NVIDIA 10-Q（期間結束 2025-10-26；accession 0001045810-25-000230）"
PR = "https://nvidianews.nvidia.com/news/nvidia-announces-financial-results-for-"


def lit(t):   # Excel 公式字串常數上限 255 字元：切段以 & 相接
    t = t.replace('"', '""')
    return "&".join('"' + t[i:i + 200] + '"' for i in range(0, len(t), 200)) or '""'


def lab(ws, prefix, start=1, col=1):
    for r in range(start, ws.max_row + 1):
        v = ws.cell(r, col).value
        if isinstance(v, str) and v.strip().startswith(prefix): return r
    raise KeyError((ws.title, prefix))


def inrow(wb, key):
    IN = wb["In"]
    return next(r for r in range(1, IN.max_row + 1) if IN.cell(r, 1).value == key)


def dn(wb, name, ref):
    wb.defined_names[name] = DefinedName(name, attr_text=ref)


# ---------- 原則 14：Hist 頁（歷史財務移出公式；值不變）----------
HIST_Y = [  # (鍵, 項目, FY2024, FY2025, FY2026, 標記, 來源)
    ("rev", "營收", 60.922, 130.497, 215.938, "Verified", "NVIDIA 第四季新聞稿（FY2024、FY2025）；" + K10N + " " + K10),
    ("da", "折舊攤銷（現金流量表）", None, 1.864, 2.843, "Verified", K10N + "：Depreciation and amortization；SEC XBRL DepreciationDepletionAndAmortization"),
    ("cx", "資本支出（購置不動產、廠房設備與無形資產）", None, 3.236, 6.042, "Verified", K10N + "：Purchases related to property and equipment and intangible assets；XBRL PaymentsToAcquireProductiveAssets"),
    ("ar", "應收帳款（年底）", 9.999, 23.065, 38.466, "Verified", PR + "fourth-quarter-and-fiscal-2025（FY2024、FY2025）；Q2 FY27 8-K 比較欄（FY2026）"),
    ("inv", "存貨（年底）", 5.282, 10.08, 21.403, "Verified", "同上"),
    ("pre", "預付及其他流動資產（年底；含未單獨揭露的供應預付）", 3.08, 3.771, 3.18, "Verified", "同上"),
    ("ap", "應付帳款（年底）", 2.699, 6.31, 9.812, "Verified", "同上"),
    ("acc", "應計及其他流動負債（年底）", 6.682, 11.737, 21.352, "Verified", "同上"),
    ("tax", "所得稅費用（GAAP）", 4.058, 11.146, None, "Verified", PR + "fourth-quarter-and-fiscal-2025"),
    ("pti", "稅前利益（GAAP）", 33.818, 84.026, None, "Verified", "同上"),
]
HIST_P = [  # (鍵, 項目, 值, 標記, 來源)
    ("da_q3fy26", "Q3 FY26 折舊攤銷（前三季 2,031 − 上半年 1,280）", "=2.031-1.280", "Derived", Q3_10Q + "、" + Q10N + " 比較欄：前三季 2,031、上半年 1,280（US$M）"),
    ("da_q4fy26", "Q4 FY26 折舊攤銷（全年 2,843 − 前三季 2,031）", "=2.843-2.031", "Derived", K10N + "、" + Q3_10Q),
    ("da_h1fy27", "H1 FY27 折舊攤銷", 2.124, "Verified", Q10N + " 現金流量表 " + Q10),
    ("cx_q3fy26", "Q3 FY26 資本支出（前三季 4,758 − 上半年 3,122）", "=4.758-3.122", "Derived", Q3_10Q + "、" + Q10N + " 比較欄（US$M）"),
    ("cx_q4fy26", "Q4 FY26 資本支出（全年 6,042 − 前三季 4,758）", "=6.042-4.758", "Derived", K10N + "、" + Q3_10Q),
    ("cx_h1fy27", "H1 FY27 資本支出", 4.434, "Verified", Q10N + " 現金流量表 " + Q10),
    ("rev_h1fy27", "H1 FY27 營收（Q 頁 G4＋H4）", "=Q!G4+Q!H4", "Derived", "Q 頁（公司 8-K 與 10-Q 同值）"),
    ("rev_q4fy26", "Q4 FY26 營收（Q 頁 E4）", "=Q!E4", "Derived", "Q 頁（公司 8-K）"),
    ("rev_ltm", "最近四季營收（Q3 FY26～Q2 FY27；Q 頁 D4＋E4＋G4＋H4）", "=Q!D4+Q!E4+Q!G4+Q!H4", "Derived", "Q 頁"),
    ("shout", "流通股數（2026-07-26，十億股）", 24.147, "Verified", Q10N + " 股東權益變動表「Balances as of Jul 26, 2026」24,147 百萬股；封面 2026-08-21 為 24.1 billion（四捨五入） " + Q10),
    ("rsu", "未既得 RSU、PSU 與市場型 PSU（2026-07-26，十億股）", 0.168, "Verified", Q10N + " Note 2 Equity Award Activity：Balance as of Jul 26, 2026 168 百萬股（加權授予日公允價值 120.41 美元）"),
    ("opt", "選擇權（庫藏股法淨股數，十億股）", 0, "Verified", K10N + " 股權激勵附註：2007 Plan 只授予 RSU、PSU 與市場型 PSU（We grant RSUs, PSUs and market-based PSUs）；10-Q 股權變動表無選擇權列"),
    ("dil_q2", "Q2 FY27 稀釋加權股數（十億股）", "=Q!H21", "Verified", Q10N + " Note 3 Net Income Per Share：24,285 百萬股（Q 頁 H21）"),
    ("dps_q", "每季股息（美元／股）", 0.25, "Verified", Q10N + " 股東權益變動表：Cash dividends declared and paid ($0.25 per common share)"),
    ("taxpaid_26", "FY2026 已付所得稅（扣退稅）", 20.288, "Verified", K10N + "：Cash paid for income taxes, net 20,288（US$M）"),
    ("supply_c", "供應與產能承諾（2026-07-26）", 279, "Verified", Q10N + " Note 10 Commitments：Supply and capacity 279（US$B；前一季 119）"),
    ("supply_c_q1", "供應與產能承諾（2026-04-26）", 119, "Verified", Q10N + " Note 10：increasing supply commitments from $119 billion last quarter"),
    ("pre_q2", "預付及其他流動資產（2026-07-26）", 3.409, "Verified", Q10N + " 資產負債表"),
    ("oth_a_q2", "其他資產（非流動，2026-07-26）", 15.746, "Verified", Q10N + " 資產負債表（含長期預付、權益衍生工具等，未拆分）"),
    ("bs_cash", "現金及約當現金（2026-07-26）", 22.443, "Verified", Q10N + " 資產負債表 " + Q10),
    ("bs_mds", "有價債務證券", 34.143, "Verified", "同上"),
    ("bs_mes", "有價權益證券", 42.783, "Verified", "同上"),
    ("bs_nms", "非上市權益證券", 51.157, "Verified", "同上"),
    ("bs_std", "短期債務（帳面）", 1.0, "Verified", "同上"),
    ("bs_ltd", "長期債務（帳面）", 32.366, "Verified", "同上"),
    ("bs_disc", "票據未攤銷折價與發行成本（本金 33.5 − 帳面 33.366）", 0.134, "Derived", Q10N + " Note 9 Debt"),
    ("bs_lc", "營業租賃負債（流動）", 0.509, "Verified", Q10N + " 附註 14"),
    ("bs_lnc", "營業租賃負債（非流動）", 4.985, "Verified", Q10N + " 資產負債表"),
]
NOTES = [(3.5, 0.0439), (3.5, 0.0445), (4.0, 0.046), (3.5, 0.0482), (4.0, 0.0501), (3.0, 0.0561), (3.5, 0.0566)]


def hist_sheet(wb):
    H = wb.create_sheet("Hist", index=wb.sheetnames.index("Q") + 1)
    H["A1"] = ("Hist — 歷史財務與評價用資料（v0.3.5 起；原則 14、34）：In 頁的折舊、資本支出、營運資金、稅率、債務成本、股息、評價股數與 DCF 淨現金核對只引用本頁儲存格"
               "（定義名稱 h_*），不在公式內寫常數。資產負債與股數以 SEC 10-Q／10-K 為主來源（原則 34）。金額：十億美元；股數：十億股")
    for j, h in enumerate(["鍵", "項目", "FY2024", "FY2025", "FY2026", "", "標記", "來源"], 1):
        c = H.cell(3, j, h); c.font = BOLD; c.fill = HDR
    r = 4
    for k, labl, *vals, tag, src in HIST_Y:
        H.cell(r, 1, f"h_{k}"); H.cell(r, 2, labl)
        for j, (y, v) in enumerate(zip(("24", "25", "26"), vals)):
            if v is None: continue
            c = H.cell(r, 3 + j, v); c.font = BLUE; c.number_format = "#,##0.000"; dn(wb, f"h_{k}_{y}", f"Hist!${'CDE'[j]}${r}")
        H.cell(r, 7, tag); H.cell(r, 8, src); r += 1
    r += 1
    for j, h in enumerate(["鍵", "項目", "值", "", "", "", "標記", "來源"], 1):
        c = H.cell(r, j, h); c.font = BOLD; c.fill = HDR
    r += 1
    for k, labl, v, tag, src in HIST_P:
        H.cell(r, 1, f"h_{k}"); H.cell(r, 2, labl); c = H.cell(r, 3, v); c.number_format = "#,##0.000000"
        if not (isinstance(v, str) and v.startswith("=")): c.font = BLUE
        dn(wb, f"h_{k}", f"Hist!$C${r}"); H.cell(r, 7, tag); H.cell(r, 8, src); r += 1
    r += 1
    for j, h in enumerate(["鍵", "2026-06 發行票據", "本金（US$B）", "有效利率", "", "", "標記", "來源"], 1):
        c = H.cell(r, j, h); c.font = BOLD; c.fill = HDR
    r += 1; n0 = r
    for i, (a, k) in enumerate(NOTES, 1):
        H.cell(r, 1, f"h_note{i}"); H.cell(r, 2, f"第 {i} 批"); H.cell(r, 3, a).font = BLUE; H.cell(r, 4, k).font = BLUE; H.cell(r, 4).number_format = "0.00%"
        H.cell(r, 7, "Verified"); H.cell(r, 8, Q10N + " Note 9 Debt（2026-06 發行七批，有效利率 4.39%–5.66%）"); r += 1
    dn(wb, "h_note_amt", f"Hist!$C${n0}:$C${r - 1}"); dn(wb, "h_note_rate", f"Hist!$D${n0}:$D${r - 1}")
    r += 1
    for k, labl, f in (("asof", "資料基準日（現價日）", "=DATE(2026,11,5)"), ("cutoff", "資料截止日（最晚一筆資料：股價與同業報價）", "=DATE(2026,11,5)"),
                       ("built", "成品日（本版建置）", "=DATE(2026,11,11)"), ("valdate", "DCF 評價日", "=DATE(2026,12,31)"), ("bsdate", "資產負債表日（10-Q）", "=DATE(2026,7,26)")):
        H.cell(r, 1, f"h_{k}"); H.cell(r, 2, labl); H.cell(r, 3, f).number_format = "yyyy-mm-dd"; dn(wb, f"h_{k}", f"Hist!$C${r}"); H.cell(r, 7, "—"); r += 1
    H.column_dimensions["A"].width = 16; H.column_dimensions["B"].width = 60
    for c in "CDEF": H.column_dimensions[c].width = 13
    H.column_dimensions["G"].width = 12; H.column_dimensions["H"].width = 110
    return H


def hist_links(wb, ctx):
    """原則 14：In 頁公式內的歷史常數改引用 Hist（v0.3.5 數值不變）；v0.3.6 'ratio' 改最近四季（原則 35）、'sh' 評價股數（原則 11）"""
    IN = wb["In"]; P = ctx["V36P"]; OVR = ctx["OVR"]
    S = {
        "fc_wc": "=" + "+".join(f"(h_ar_{y}+h_inv_{y}+h_pre_{y}-h_ap_{y}-h_acc_{y})/h_rev_{y}" for y in ("24", "25", "26")).join(("(", ")/3")),
        "tax_w": "=(h_tax_24/h_pti_24+h_tax_25/h_pti_25+Q!F19)/3",
        "kd": "=SUMPRODUCT(h_note_amt,h_note_rate)/SUM(h_note_amt)",
        "dps_tgt": "=h_dps_q*4",
        "stub_div": "=h_dps_q*sh_v",
        "da_rev": "=(h_da_q3fy26+h_da_q4fy26+h_da_h1fy27)/h_rev_ltm" if "ratio" in P else "=h_da_26/h_rev_26",
        "cx_rev": "=(h_cx_q3fy26+h_cx_q4fy26+h_cx_h1fy27)/h_rev_ltm" if "ratio" in P else "=h_cx_26/h_rev_26",
        "sh_v": "=h_shout+h_rsu+h_opt" if "sh" in P else "=h_dil_q2",
    }
    for k, f in S.items():
        if k in OVR: continue
        c = IN.cell(inrow(wb, k), 5)
        assert isinstance(c.value, str) and c.value.startswith("="), (k, c.value)
        c.value = f
    # 原則 34、45：資產負債與股數以 10-Q 為主來源（URL、文件日期、段落）
    for k, s1, s2 in (
            ("cash_v", Q10N + " 資產負債表：現金及約當現金 22,443＋有價債務證券 34,143（US$M）" + Q10, "8-K 新聞稿資產負債表（同值，第二來源）"),
            ("inv_v", Q10N + " 資產負債表：有價權益證券 42,783＋非上市權益證券 51,157（US$M）" + Q10,
             "8-K 新聞稿（同值，第二來源）。權益證券損益不進 EPS（Q 頁），故以帳面計入評價；非上市部分為帳面，Sens 測不計。其他資產內受鎖定的長期上市權益證券 5.0 未計入（保守）"),
            ("debt_v", Q10N + " Note 9 Debt：票據本金合計 33,500（帳面 33,366＝短期 1,000＋長期 32,366）" + Q10, "8-K 新聞稿資產負債表（第二來源）"),
            ("lease_v", Q10N + " 附註 14：營業租賃負債 5,494（流動 509＋非流動 4,985）" + Q10, "8-K 新聞稿資產負債表：長期營業租賃負債 4,985（第二來源）"),
            ("bs_date", Q10N + " " + Q10, "評價日 2026-12-31；淨現金以 stub 期滾到評價日")):
        r = inrow(wb, k); IN.cell(r, 8, s1); IN.cell(r, 9, s2)
    r = inrow(wb, "sh_v")
    if "sh" in P:
        IN.cell(r, 2, "評價用股數（流通股＋RSU 全額＋選擇權庫藏股法）")
        IN.cell(r, 7, "Derived")
        IN.cell(r, 8, "原則 11（個股DCF標準規格第 3 節）：流通股 24,147＋未既得 RSU／PSU 168＋選擇權 0（百萬股；" + Q10N + " 股東權益變動表與 Note 2；Hist 頁）")
        IN.cell(r, 9, "沒收率未揭露，RSU 全額計入；PSU 達標機率未揭露，與 RSU 合併揭露無法拆分，全額計入（保守）。v0.3.5 以前＝Q2 FY27 稀釋加權股數 24.285（Sens 業務主因表有「只計流通股」案例）")
    else:
        IN.cell(r, 8, Q10N + " Note 3：Q2 FY27 稀釋加權股數 24,285 百萬股（Hist 頁 h_dil_q2＝Q 頁 H21；8-K 同值）")
    for k, ttl in (("da_rev", "折舊 ÷ 營收"), ("cx_rev", "資本支出 ÷ 營收")):
        r = inrow(wb, k); hk = k[:2]
        if "ratio" in P:
            IN.cell(r, 4, f"=h_{hk}_h1fy27/h_rev_h1fy27" if hk == "da" else "=h_cx_q4fy26/h_rev_q4fy26"); IN.cell(r, 6, f"=h_{hk}_26/h_rev_26")
            dn(wb, k + "_lo", f"In!$D${r}"); dn(wb, k + "_hi", f"In!$F${r}")
            for c in (4, 6): IN.cell(r, c).number_format = "0.00%"
            IN.cell(r, 7, "Derived")
            IN.cell(r, 8, f"原則 35（v0.3.6）：{ttl}，最近四季（Q3 FY26～Q2 FY27）實績、全期持平：(Q3 FY26＋Q4 FY26＋H1 FY27) ÷ 最近四季營收（Hist 頁；{Q10N}、{K10N}、{Q3_10Q}）")
            IN.cell(r, 9, "並列歷年：FY2025 " + ("1.43%" if hk == "da" else "2.48%") + "、FY2026（v0.3.5 以前的基準）見敏感度高欄；敏感度低＝"
                    + ("H1 FY27（低於基準）" if hk == "da" else "最近四季中最低的一季 Q4 FY26（1,284 ÷ 68,127；H1 FY27 的 2.49% 高於基準，不能當低）") + "。收斂路徑無證據，只作 Sens")
        else:
            IN.cell(r, 7, "Derived")


def v36_values(wb, ctx):
    IN = wb["In"]; P = ctx["V36P"]
    if "tt" in P:   # 原則 26：目標時點＝現價日＋12 個月（評價日 2026-12-31 不變；X20）
        r = inrow(wb, "t_tgt")
        IN.cell(r, 2, "目標時點距評價日（目標時點＝現價日＋12 個月）"); IN.cell(r, 5, "=(EDATE(h_asof,12)-h_valdate)/365")
        IN.cell(r, 8, "目標價為現價日起 12 個月的價格，與 Comparable 的 12 個月後前瞻 EPS 同一時點（原則 26）"); IN.cell(r, 9, "v0.3.5 以前＝1 年（評價日＋12 個月，即現價日＋約 14 個月）")
        r = inrow(wb, "dps_tgt")
        IN.cell(r, 5, "=h_dps_q*ROUND(t_tgt*4,0)"); IN.cell(r, 9, "評價日到目標時點之間的季度股息次數＝ROUND(t_tgt × 4)，每季金額不變（v0.3.5 以前固定 4 次）")


def overrides(wb, OVR):
    """敏感度快照用（nvidia/sens_nvidia.py）：In 頁基準欄覆寫；正式建置不設"""
    IN = wb["In"]
    for k, v in OVR.items():
        if k.startswith("_"): continue
        r = inrow(wb, k)
        if v in ("@lo", "@hi"):
            IN.cell(r, 5).value = IN.cell(r, 4 if v == "@lo" else 6).value
        else:
            IN.cell(r, 5).value = v if v.startswith("=") else float(v)
    for flag, pre in (("_hbmflat", "hbm"), ("_cdxflat", "cdx")):
        if flag in OVR:
            for s in "ABC":
                for y in (27, 28, 29):
                    IN.cell(inrow(wb, f"{pre}_{s}_{y}"), 5).value = f"={pre}_{s}_26"


# ---------- 原則 2：單一主要標記與來源誘因 ----------
def incentive(src):
    t = str(src or "")
    if any(k in t for k in ("共識", "賣方", "分析師", "stockanalysis", "247wallst", "Wallstreetzen", "Yahoo", "Webull", "Bernstein", "BofA", "HSBC", "Morgan Stanley", "Fubon")):
        return "賣方分析師或資料商：評等與交易業務誘因，預估與目標價傾向跟隨股價"
    if any(k in t for k in ("法說", "逐字稿", "二手", "implicator", "moomoo", "biggo", "CNBC")):
        return "公司管理層口頭陳述（二手轉述）：股價與市場預期管理誘因，未經查核、可能偏樂觀"
    if any(k in t for k in ("擁有者命題", "工作單")):
        return "本專案擁有者的命題設定：待模型檢驗的主張，不是證據"
    if any(k in t for k in ("TrendForce", "DIGITIMES", "研調", "UNITS", "market.py", "Siliconomics", "台積電模型")):
        return "研調機構或上游模型轉載的產業估計：研調機構有訂閱與客戶關係誘因，口徑由提供者定義"
    if any(k in t for k in ("指引", "公司", "新聞稿", "8-K")):
        return "公司指引或自揭數字：預期管理誘因，口徑由公司定義"
    return "利害關係方提供：未經獨立查核"


def norm_tags(wb):
    n = 0
    for sh, tcol, scol, icol in (("In", 7, 8, 10), ("Q", 12, 13, 15)):
        ws = wb[sh]
        hr = next(r for r in range(1, 15) if ws.cell(r, tcol).value == "標記")
        c = ws.cell(hr, icol, "來源誘因"); c.font = BOLD; c.fill = HDR; ws.column_dimensions[c.column_letter].width = 50
        for r in range(hr + 1, ws.max_row + 1):
            t = ws.cell(r, tcol).value
            if not isinstance(t, str) or t == "標記": continue
            if "／" in t:
                parts = [p.strip() for p in re.split(r"／|（", t)]
                prim = "Interested-party" if any(p.startswith("Interested-party") for p in parts) else parts[0]
                ws.cell(r, tcol).value = prim
                note = (f"欄別標記（v0.3.4 混標「{t}」）：主要標記取最近一期原文；實績季為財報（Verified）、Q3 FY27 為公司指引（Interested-party）、Q4 FY27 為假設（Assumed）" if sh == "Q"
                        else f"v0.3.4 混標「{t}」：主要標記取最近一期原文；低／高區間為假設（Assumed）")
                ws.cell(r, icol, (incentive(ws.cell(r, scol).value) + "；" + note) if prim == "Interested-party" else note); n += 1; t = prim
            if sh == "In" and t == "Verified":
                base = ws.cell(r, 5).value
                if isinstance(base, str) and base.startswith("=") and re.search(r"[+\-*/]", base[1:]) and not base.startswith("=DATE("):
                    ws.cell(r, tcol).value = t = "Derived"; n += 1
            if t not in ("Verified", "Interested-party", "Analogy", "Assumed", "Derived", "—"):
                ws.cell(r, tcol).value = t = "Verified" if t.startswith("Verified") else ("Assumed" if "Assumed" in t else ("Derived" if "Derived" in t else "Interested-party")); n += 1
            if t == "Interested-party" and not ws.cell(r, icol).value:
                ws.cell(r, icol, incentive(f"{ws.cell(r, scol).value or ''}｜{ws.cell(r, scol + 1).value or ''}｜{ws.cell(r, 2).value or ''}"))
    return n


# ---------- 原則 1：build_id、輸入指紋、模型狀態 ----------
def build_id(here, stage):
    h = hashlib.sha256(stage.encode())
    for fn in sorted(["build_nvidia.py", "v035_post.py", "px1105_post.py", "tp_v2.py", "accel_share.py", "sens.json", "segments.json", "hist_frozen.json"]):
        p = os.path.join(here, fn)
        if os.path.exists(p): h.update(fn.encode()); h.update(open(p, "rb").read())
    return f"nvidia-v{stage}-{h.hexdigest()[:10]}"


def governance(wb, SJ, bid, OVR):
    SE = wb["Sens"]; IN = wb["In"]
    hr = next(r for r in range(1, 15) if IN.cell(r, 7).value == "標記")
    last = next(r for r in range(hr + 1, IN.max_row + 1) if str(IN.cell(r, 1).value or "").startswith("情境驅動對照")) - 1
    rng = f"In!$D${hr + 1}:$F${last}"
    r = SE.max_row + 3; G = {}
    SE.cell(r, 1, "模型治理（v0.3.5 起；原則 1）：Sens 主表與目標價參數表為活頁簿公式；「業務主因（整本重算）」表與 HTML 是建置快照。只在 Excel 改 In 頁而未重跑 python3 nvidia/sens_nvidia.py 與建置時，下方模型狀態顯示 STALE").font = BOLD; r += 1
    SE.cell(r, 1, "輸入指紋（本檔即時重算；In 頁低／基準／高欄，情境驅動對照以上）"); SE.cell(r, 4, f"=SUMPRODUCT(N(+{rng}),ROW({rng})*3+COLUMN({rng}))").number_format = "0.000000"; G["fp"] = r; r += 1
    base = (SJ or {}).get("_base") if not OVR else None
    SE.cell(r, 1, "輸入指紋（建置快照的基準檔）")
    if base: SE.cell(r, 4, base["fp"]).number_format = "0.000000"
    fb = r; r += 1
    SE.cell(r, 1, "加權目標價（建置快照的基準檔）")
    if base: SE.cell(r, 4, base["tp"]).number_format = "#,##0.00"
    tb = r; r += 1
    tpc = f"Val!$B${lab(wb['Val'], '目標價（DCF 與 Comparable 加權）')}"
    SE.cell(r, 1, "模型狀態（CURRENT＝Sens 快照、HTML、README 與本檔輸入一致；STALE＝In 頁已改、須重跑 python3 nvidia/sens_nvidia.py 與建置）").font = BOLD
    SE.cell(r, 4, f'=IF(AND(ISNUMBER(D{fb}),ISNUMBER(D{tb}),ABS(D{G["fp"]}-D{fb})<=0.000000001*MAX(1,ABS(D{fb})),ABS({tpc}-D{tb})<0.005),"CURRENT","STALE")'); G["st"] = r; r += 1
    SE.cell(r, 1, "build_id（README、Sens、HTML 三處相同）"); SE.cell(r, 4, bid); G["bid"] = r
    return G


# ---------- 原則 32、47、48：目標價 Sens（DCF 單法美元欄；相對倍數係數＝1；業務主因整本重算快照）----------
BIZ = [("NVIDIA 顆數係數 ×0.9／×1.1（FY2028 起，三情境同乘）", "uk", "引擎顆數乘數換成 uk 低／高"),
       ("GPU 模組售價：Blackwell 低／高", "pm_bw", "In 頁敏感度低／高（k_rec、c_nh 整本重算）"),
       ("GPU 模組售價：Rubin 低／高", "pm_rb", "同上"),
       ("GPU 模組售價：Rubin Ultra 低／高", "pm_ru", "同上（售價與非 HBM 成本同比例，見讀法）"),
       ("GPU 模組售價：Hopper 低／高", "pm_hp", "同上"),
       ("匣級加價保留比例 k_tray 0／1", "k_tray", "k_rec、c_nh 隨之重算"),
       ("HBM／LPDDR 吸收比例 abs_mem 0／1（三情境）", "abs_mem", "全額轉嫁／全額吸收"),
       ("中國恢復 10／40（FY2028 起每年；基準 0）", "cn_amt", "事件型尾端（原則 19）：只進 Sens"),
       ("Edge Computing 毛利率 低／高", "gm_edge", "50%／70%"),
       ("Edge Computing 營收年增 低／高", "eg", "0%／15%"),
       ("Edge Computing 季增（Q3、Q4 FY27）低／高", "edge_q", "Data Center 錨點＝總營收 − Edge，Edge 誤差傳到錨點與 k_rec（原則 39）"),
       ("營業費用彈性 phi 低／高", "phi", "0.4／0.8"),
       ("營業費用年增下限 低／高", "g_opx_min", "0%／10%"),
       ("營業費用率 −2pt／＋2pt", "opx_adj", "引擎 opxa 換成 opx_adj 低／高"),
       ("有效稅率 tx_f 16%／18%", "tx_f", "公司 FY2027 指引區間"),
       ("其他收益淨額 低／高", "oth_f", "0.5／2.4（US$B）"),
       ("稀釋股數年變動 低／高", "sh_g", "−2.5%／0%"),
       ("Q3 FY27 營收指引 −2%／＋2%（錨點）", "q3_rev", "k_rec、c_nh 整本重算（原則 39）"),
       ("Q3 FY27 毛利率指引 ±50bp（錨點）", "q3_gm", "同上"),
       ("Q4 FY27 營收季增 5%／15%（錨點）", "q4_g", "同上"),
       ("Q4 FY27 毛利率 71%／74%（錨點）", "q4_gm", "同上"),
       ("Q4 FY27 營業費用季增 低／高（錨點）", "q4_opx_g", "同上"),
       ("Q3、Q4 FY27 其他收益 低／高（錨點）", "q_oth", "同上"),
       ("Q3、Q4 FY27 稅率 16%／18%（錨點）", "tx_27", "同上"),
       ("Q3、Q4 FY27 股數季變動 低／高（錨點）", "q_sh", "同上")]
BIZ36 = [("k_rec 自 FY2028 起線性衰減、FY2030 回到 1（原則 48；基準＝持續）", "krdecay", "只有高：k_rec＝校準值 → FY2028 +1/3、FY2029 +2/3、FY2030 回到 1 的差距（整本重算）"),
         ("折舊 ÷ 營收 低／高（原則 35）", "da_rev", "低＝H1 FY27、高＝FY2026（v0.3.5 基準）；基準＝最近四季"),
         ("資本支出 ÷ 營收 低／高（原則 35）", "cx_rev", "低＝Q4 FY26（最近四季最低季）、高＝FY2026（v0.3.5 基準）；基準＝最近四季"),
         ("評價股數只計流通股（原則 11 原口徑）", "sh_v", "只有低：不計 RSU（潛在稀釋全不計）")]
EXEMPT = ("pa_amd、pa_tpu、pa_trn、pa_oth 不列：只決定 Units 頁金額份額展示與「ASIC 金額份額 ±5pt」案例的份額基數，不進基準目標價，對加權目標價結構上恆為 0。"
          "r_disc 不列：只用於舊法規則值的折現參考欄，舊法不進結論，對加權目標價結構上恆為 0。"
          "rep_nv26 不列：只用於 Recon 網通與運算缺口對帳，不進基準目標價，對加權目標價結構上恆為 0。"
          "g28_gm 不列：公司 FY2028 毛利率指引只作 Recon 對帳與反解（不需驗收清單 X3），不進基準目標價。"
          "cons_eps28 不列：共識 FY2028 EPS 只用於 Recon 與 Val 的共識對照列，不進基準目標價。")


def sens_tp(wb, SJ, ctx):
    SE = wb["Sens"]; V = wb["Val"]
    r0 = lab(SE, "目標價參數敏感度"); hr = r0 + 1
    SE.cell(hr, 9, "低：DCF 單法變化（美元）").font = BOLD; SE.cell(hr, 10, "高：DCF 單法變化（美元）").font = BOLD
    r = hr + 1
    while SE.cell(r, 1).value and not str(SE.cell(r, 1).value).startswith("檢查"):
        is_w = SE.cell(r, 8).value == "tp_w_dcf"
        for cv, cd in (("C", "I"), ("E", "J")):
            SE[f"{cd}{r}"] = '="—（權重案例）"' if is_w else f'=IF(ISNUMBER({cv}{r}),({cv}{r}-G{r})/tp_w_dcf,"—")'
            SE[f"{cd}{r}"].number_format = "+#,##0.00;-#,##0.00"
        r += 1
    SE.cell(r0, 1).value = str(SE.cell(r0, 1).value) + "（v0.3.5 起 I、J 欄＝DCF 單法目標價變化，美元；Comparable 不變的案例＝加權變化 ÷ DCF 權重）"
    SE["A2"] = EXEMPT
    G = {"tp0": r0}
    tpc = f"Val!$B${lab(V, '目標價（DCF 與 Comparable 加權）')}"; dcf = f"Val!$B${lab(V, 'DCF 目標價（情境加權）')}"; cu = f"Val!$B${lab(V, 'Comparable 目標價（市場主流同業法）')}"
    r = SE.max_row + 3
    SE.cell(r, 1, "目標價敏感度：業務主因（整本重算快照；v0.3.5 起，原則 22、32、39、47、48）").font = BOLD; G["biz_h"] = r; r += 1
    SE.cell(r, 1, "每案把輸入換掉後整本重建、LibreOffice 重算（nvidia/sens_nvidia.py → nvidia/sens.json），讀加權目標價與 DCF 目標價；不是活頁簿公式，In 頁改過須重跑（模型狀態見本頁底部）。第一列「相對倍數係數＝1」為活頁簿公式（原則 47）"); r += 1
    for j, h in enumerate(["案例", "設定", "低：加權目標價", "低：相對基準", "高：加權目標價", "高：相對基準", "基準加權目標價", "鍵", "低：DCF 變化（美元）", "高：DCF 變化（美元）"], 1):
        c = SE.cell(r, j, h); c.font = BOLD; c.fill = HDR
    r += 1; G["biz0"] = r; G["bizrow"] = {}
    # 原則 47：相對倍數係數＝1（活頁簿公式）
    SE.cell(r, 1, "相對倍數係數＝1（Comparable 不調整；原則 47）"); SE.cell(r, 2, "Comparable 改用市場主流同業法（同業前瞻 P/E 中位數 × 12 個月後前瞻 EPS），DCF 不變；只有低")
    SE.cell(r, 3, f"=tp_w_dcf*{dcf}+(1-tp_w_dcf)*{cu}").number_format = "#,##0.00"; SE.cell(r, 4, f"=C{r}/G{r}-1").number_format = "+0.0%;-0.0%"
    SE.cell(r, 5, "—"); SE.cell(r, 7, f"={tpc}").number_format = "#,##0.00"; SE.cell(r, 8, "rm_factor"); SE.cell(r, 9, 0).number_format = "+#,##0.00;-#,##0.00"
    G["rm1"] = r; r += 1
    b = (SJ or {}).get("_base", {})
    for name, key, setting in BIZ + (BIZ36 if ctx["V36"] else []):
        d = (SJ or {}).get("biz", {}).get(key, {})
        SE.cell(r, 1, name); SE.cell(r, 2, setting); SE.cell(r, 8, key if key != "krdecay" else "k_rec"); G["bizrow"][key] = r
        for side, ct, cr, cdl in (("低", 3, 4, 9), ("高", 5, 6, 10)):
            v = d.get(side)
            if v and b:
                SE.cell(r, ct, v["tp"]).number_format = "#,##0.00"; SE.cell(r, cr, v["tp"] / b["tp"] - 1).number_format = "+0.0%;-0.0%"
                SE.cell(r, cdl, v["dcf"] - b["dcf"]).number_format = "+#,##0.00;-#,##0.00"
            else:
                SE.cell(r, ct, "—")
        if b: SE.cell(r, 7, b["tp"]).number_format = "#,##0.00"
        r += 1
    G["biz1"] = r - 1
    return G


# ---------- Recon：同口徑註記、反解（原則 4、33、36）；毛利率差距分解與 abs_mem 反解（原則 50，v0.3.6）----------
def recon(wb, SJ, ctx):
    RC = wb["Recon"]; G = {}; V = wb["Val"]; PB = ctx["PB"]
    # 驗收核對（2026-11-11）：B2 候選來源①「可查」欄的 10-Q 殘留說法
    r = lab(RC, "① Rubin CY2026 出貨時點")
    RC.cell(r, 10).value = str(RC.cell(r, 10).value).replace("（財務長書面說明、10-Q 未取得）", "（10-Q 已讀取，未揭露 Rubin 季別顆數；財務長書面說明未取得）")
    assert "10-Q 未取得" not in str(RC.cell(r, 10).value)
    # 情境 B／C 欄改引用各自的推導區塊（16–18、21 列；21 列未校準運算＝認列後 ÷ k_rec）
    for prefix, f in (("認列後 FY2027 Data Center 營收 vs 錨點", lambda s_: f"=PL!D{PB[s_]['comp']}+PL!D{PB[s_]['net']}"),
                      ("認列後 FY2027 毛利率 vs 錨點毛利率", lambda s_: f"=PL!D{PB[s_]['bugm']}"),
                      ("認列顆數（百萬顆）：k_rec × 市場層", lambda s_: f"=PL!D{PB[s_]['nvu']}"),
                      ("運算：未校準由下而上 vs 錨點", lambda s_: f"=PL!D{PB[s_]['comp']}/k_rec")):
        rr = lab(RC, prefix)
        for s_, col in (("B", 5), ("C", 7)): RC.cell(rr, col).value = f(s_)
        if ctx["MW"] and not prefix.startswith("認列顆數"):
            RC.cell(rr, 10).value = (str(RC.cell(rr, 10).value or "") + "；B、C 欄引用各自推導區塊。v0.3.6 起月數加權使 B、C 的 FY2027 含 1/12 的各自 CY2027 一般 DRAM 價格指數，"
                                     "而 k_rec、c_nh 只以情境 A 校準，故 B、C 與錨點有小差額（PL 頁情境 B／C 核對列）")
    r = lab(RC, "2027 年 NVIDIA 金額份額（加速器晶片")
    RC.cell(r, 10, str(RC.cell(r, 10).value or "") + "（原則 33、50：模型輸出——由 Assumed／Interested-party 的 ASIC 每顆金額與市場層顆數推得，非獨立證據；與命題同為模組層級金額份額）")
    r = RC.max_row + 2
    RC.cell(r, 1, "F. 反解：支持現價與評等翻轉所需的值（v0.3.5 起；原則 4、24、36）").font = BOLD; r += 1
    dcf = f"Val!$B${lab(V, 'DCF 目標價（情境加權）')}"
    feps = f"Val!$B${lab(V, '12 個月後的前瞻 EPS（情境加權）')}"; pmed = f"Val!$B${lab(V, '同業前瞻 P/E 中位數')}"
    up = f"Val!$B${lab(V, '上檔空間：加權目標價')}"
    RC.cell(r, 1, "評等級距邊界：升一級（上檔空間）"); RC.cell(r, 2, f"=IF({up}<-0.5,-0.5,IF({up}<-0.25,-0.25,IF({up}<=0.25,0.25,IF({up}<=0.5,0.5,NA()))))").number_format = "+0%;-0%"
    RC.cell(r, 10, "級距：大於 +50% 強力買進、+25%～+50% 買進、−25%～+25% 中立、−50%～−25% 放空、小於 −50% 強力放空（Andy 2026-10-25；Val 投資建議公式與 tools/target_summary/build_summary.py 同一套邊界，原則 24）"); G["e_up"] = r; r += 1
    RC.cell(r, 1, "評等級距邊界：降一級（上檔空間）"); RC.cell(r, 2, f"=IF({up}>0.5,0.5,IF({up}>0.25,0.25,IF({up}>=-0.25,-0.25,IF({up}>=-0.5,-0.5,NA()))))").number_format = "+0%;-0%"; G["e_dn"] = r; r += 1
    for j, h in enumerate(["項目", "需要的值", "模型現值", "需要的變動"], 1): RC.cell(r, j, h).font = BOLD
    r += 1
    w = "tp_w_dcf"; G["inv"] = {}
    tgt = {"px": "px", "up": f"px*(1+B{G['e_up']})", "dn": f"px*(1+B{G['e_dn']})"}
    nm = {"px": "支持現價", "up": "評等升一級", "dn": "評等降一級"}
    for t in ("up", "dn", "px"):
        for kind, f, fmt, cur in (("pe", f"=({tgt[t]}-{w}*{dcf})/((1-{w})*{feps})", "0.0", f"={pmed}*rm_factor"),
                                  ("rm", f"=({tgt[t]}-{w}*{dcf})/((1-{w})*{feps}*{pmed})", "0.00", "=rm_factor"),
                                  ("dcf", f"=({tgt[t]}-(1-{w})*tp_comp_adj)/{w}", "0.00", f"={dcf}")):
            labl = {"pe": f"{nm[t]}所需的 Comparable 採用倍數（同業中位數 × 相對倍數係數；DCF 不變）",
                    "rm": f"{nm[t]}所需的相對倍數調整係數（同業中位數與前瞻 EPS 不變；DCF 不變）",
                    "dcf": f"{nm[t]}所需的 DCF 目標價（Comparable 不變）"}[kind]
            RC.cell(r, 1, labl); RC.cell(r, 2, f).number_format = fmt; RC.cell(r, 3, cur).number_format = fmt
            RC.cell(r, 4, f"=B{r}/C{r}-1").number_format = "+0.0%;-0.0%"
            RC.cell(r, 10, "閉式：加權目標價＝w × DCF＋（1 − w）× 倍數 × EPS × 係數；另一法固定。反解值是目標價等於該點的值，不是評等門檻（評等門檻見升一級／降一級列）")
            G["inv"][f"{kind}_{t}"] = r; r += 1
    if ctx["V36"]:
        r += 1
        RC.cell(r, 1, "G. FY2028 毛利率與公司指引的差距分解（情境 B；v0.3.6 起，原則 50；整本重算快照 nvidia/sens.json，只展示、不以指引校準，X3）").font = BOLD; r += 1
        gm = (SJ or {}).get("gm", {})
        for j, h in enumerate(["項目", "百分點", "", "", "", "", "", "", "來源", "說明"], 1): RC.cell(r, j, h).font = BOLD
        r += 1; G["gm0"] = r
        rows = [("FY2028 毛利率（情境 B）− 公司 FY2028 毛利率指引中點（In 頁 g28_gm）", "total", "合計＝以下五項之和"),
                ("① HBM／LPDDR 吸收比例（abs_mem 0.5 vs 全額轉嫁 0）", "abs", "基準吸收一半漲價；全額轉嫁時毛利率較高"),
                ("② HBM 價格（全額轉嫁下，HBM 價格路徑 vs 持平在 2026）", "hbm", "轉嫁後售價與成本同額上升，毛利率仍被稀釋"),
                ("③ LPDDR 價格（全額轉嫁、HBM 持平下，一般 DRAM 價格路徑 vs 持平）", "lp", "同上"),
                ("④ 世代與顆數組合、網通與 Edge 占比（記憶體價格持平時 FY2028 − FY2027 錨點；餘項）", "mix", "Rubin／Rubin Ultra 占比與 c_nh 校準"),
                ("⑤ FY2027 錨點毛利率 − 指引中點（g28_gm）", "anchor", "錨點＝Q1、Q2 實績＋Q3 指引＋Q4 假設")]
        for labl, k, note in rows:
            RC.cell(r, 1, labl); v = gm.get(k)
            if v is not None: RC.cell(r, 2, v * 100).number_format = "+0.00;-0.00"
            else: RC.cell(r, 2, "—")
            RC.cell(r, 9, "nvidia/sens.json（整本重算）"); RC.cell(r, 10, note); G["gm_" + k] = r; r += 1
        RC.cell(r, 1, "HBM 單項（② HBM 價格＋① 吸收比例中按 ②／(②＋③) 分攤的 HBM 份額；命題 2 判定用）")
        RC.cell(r, 2, f'=IF(AND(ISNUMBER(B{G["gm_hbm"]}),ISNUMBER(B{G["gm_lp"]}),ISNUMBER(B{G["gm_abs"]})),B{G["gm_hbm"]}+B{G["gm_abs"]}*B{G["gm_hbm"]}/(B{G["gm_hbm"]}+B{G["gm_lp"]}),"—")').number_format = "+0.00;-0.00"
        RC.cell(r, 9, "本節①②③"); RC.cell(r, 10, "命題 2 只問 HBM 漲價：取 HBM 價格單項，加上吸收比例效果中屬於 HBM 的份額（以②、③ 比例分攤；LPDDR 份額不計）")
        G["gm_hbm1"] = r; r += 1
        RC.cell(r, 1, "檢查：① 至 ⑤ 合計 − 合計列（百分點，應為 0）"); RC.cell(r, 2, f"=SUM(B{G['gm_abs']}:B{G['gm_anchor']})-N(B{G['gm_total']})").number_format = "0.0000"; G["gm_chk"] = r; r += 1
        sv = (SJ or {}).get("abs72")
        RC.cell(r, 1, '="反解：FY2028 毛利率（情境 B）達指引下緣 "&TEXT(g28_gm_lo,"0%")&" 所需的 abs_mem（三情境同用；只展示，基準維持 "&TEXT(abs_mem,"0.0")&"）"')
        RC.cell(r, 2, sv if sv is not None else "—").number_format = "0.00"; RC.cell(r, 3, "=abs_mem").number_format = "0.00"
        RC.cell(r, 9, "nvidia/sens.json：abs_mem 0／0.5／1 整本重算，線性內插或外推")
        RC.cell(r, 10, '=IF(ISNUMBER(B' + str(r) + '),IF(B' + str(r) + '<0,"所需值小於 0：即使記憶體漲價全額轉嫁，毛利率仍低於指引下緣；要達到指引需售價漲幅超過記憶體成本增加，或非記憶體成本下降","所需值落在 0 以上"),"—")')
        G["abs72"] = r; r += 1
        RC.cell(r, 1, "k_rec 持續性（原則 48）：基準假設校準殘差 k_rec 自 FY2028 起持續；Sens 業務主因表有「FY2028 起線性衰減、FY2030 回到 1」案例。基準選持續的理由：k_rec 是顆數層與售價層的合併係數（Recon B2 各候選來源皆未驗證），沒有證據顯示缺口會收斂；收斂到 1 等於假設市場層顆數 × 橋接表售價在 FY2030 完全實現，對目標價是上偏")
        G["krec"] = r; r += 1
    return G


# ---------- Val：舊法標示、期間並列、權重占比、子群中位數、同業口徑（原則 5、7、18、25、30、42、43）----------
def val(wb, ctx):
    V = wb["Val"]; MW = ctx["MW"]
    V["A1"].value = ("Val — 估值（美元；EPS 口徑＝GAAP 扣除權益證券損益）。目標價＝DCF（標準版 FCFF＋WACC，第 0.1 節）與 Comparable（同業前瞻 P/E 中位數 × 相對倍數調整 × 12 個月後前瞻 EPS，第 0.2、0.5 節）加權（第 0.3 節）。"
                     "第 1 節「價格與市場隱含倍數」為對照；第 2–4 節與「舊法：規則值情境加權」「P/E：十年歷史倍數」是舊法（10 年落後 P/E），只作參考，不進結論（v0.3.5 起；原則 42）")
    V["A2"].value = "舊法不適用的理由：歷史倍數是落後（trailing）倍數，含高成長期溢價（FY2023 約 115 倍）；乘在 FY2028 EPS 上得到的是約 2028-01 的價格，遠高於現價。【舊法，不進結論】"
    for rr in range(1, V.max_row + 1):
        a = V.cell(rr, 1).value
        if isinstance(a, str) and a.startswith(("舊法：規則值情境加權", "P/E：十年歷史倍數", "2. 循環位置", "3. 規則值隱含股價", "4. What must be true")) and "不進結論" not in a:
            V.cell(rr, 1).value = a + "【舊法，不進結論】"
        e = V.cell(rr, 5).value
        if isinstance(e, str) and e.startswith("FY2028 EPS 三情境幾乎相同"):
            V.cell(rr, 5).value = ("分母：本模型各情境 FY2028 EPS（FY 口徑，2027-02～2028-01；B／C／D 欄＝情境 A／B／C）；現價日 2026-11-05。"
                                   + ("v0.3.6 起 FY2028 含 1/12 的 CY2028（月數加權），情境係數自 FY2028 起部分作用" if MW else "FY2028 三情境幾乎相同（情境係數從 FY2029 才開始）"))
        if isinstance(a, str) and a.startswith("讀法：1.00＝價格不變"):
            V.cell(rr, 1).value = "讀法（舊法，不進結論）：1.00＝價格不變、1.21＝兩年約 10% 年報酬；所需倍數與歷史年底倍數分位比較，見上表（數字由公式計算）。"
    for lbl, note in (("現價 ÷ FY2029 EPS（模型）", "分母：本模型各情境 FY2029 EPS（FY 口徑，2028-02～2029-01）；現價日 2026-11-05"),
                      ("現價 ÷ 共識 FY2028 EPS", "分母：共識 FY2028 EPS 15.91（247wallst 2026-10-19，新 non-GAAP 口徑含 SBC；FY 口徑，非 NTM）；資料商區間 15.70–16.13，口徑與日期未釐清前不據以判斷市場錯估（原則 43）"),
                      ("命題：現價約為", "擁有者命題原值（Interested-party，工作單）；現行值見上方模型與共識兩列"),
                      ("檢查用：現價 ÷ 各情境 FY2028 EPS", "分母：本模型各情境 FY2028 EPS（FY 口徑）；現價日 2026-11-05")):
        V.cell(lab(V, lbl), 5).value = note
    rr = lab(V, "同業前瞻 P/E 中位數")
    V.cell(rr, 5).value = "分母：各同業資料商未來 12 個月共識 EPS（stockanalysis 統計頁 2026-10-27～10-30）；價格 2026-11-05 收盤（原則 43）"
    rr = lab(V, "相對倍數調整係數")
    V.cell(rr, 3).value = f'=IF(B{lab(V, "原始比值")}>B{lab(V, "夾值上限")},"已達夾值上限：原始比值 "&TEXT(B{lab(V, "原始比值")},"0.00"),"未觸及夾值（原始比值 "&TEXT(B{lab(V, "原始比值")},"0.00")&"，夾值 0.5–2.0）")'
    G = {}
    r = V.max_row + 2
    V.cell(r, 1, "0.6 期間並列、方法占比與同業口徑（v0.3.5 起；原則 5、7、18、25、30）").font = BOLD; r += 1
    V.cell(r, 1, "折現期間：DCF 評價日 → 目標時點（年）"); V.cell(r, 2, "=t_tgt").number_format = "0.00"; V.cell(r, 5, '="評價日 "&TEXT(h_valdate,"yyyy-mm-dd")&"；DCF 每股價值以股權成本推 t_tgt 年"'); G["per1"] = r; r += 1
    V.cell(r, 1, "持有期間：現價日 → 目標時點（年）"); V.cell(r, 2, "=(h_valdate+t_tgt*365-h_asof)/365").number_format = "0.00"
    V.cell(r, 5, '="現價日 "&TEXT(h_asof,"yyyy-mm-dd")&"；目標時點 "&TEXT(h_valdate+t_tgt*365,"yyyy-mm-dd")&"；上檔空間是這段期間的價格變化，不是年化報酬"'); G["per2"] = r; r += 1
    tp = lab(V, "目標價（DCF 與 Comparable 加權）"); dcf = lab(V, "DCF 目標價（情境加權）")
    fe, pm = lab(V, "12 個月後的前瞻 EPS（情境加權）"), lab(V, "同業前瞻 P/E 中位數")
    V.cell(r, 1, "現價 ÷ 12 個月前瞻 EPS（情境加權；統一分母，原則 43）"); V.cell(r, 2, f"=px/B{fe}").number_format = "0.0"
    V.cell(r, 5, '="現價日 "&TEXT(h_asof,"yyyy-mm-dd")&"；前瞻 EPS＝3/12 × FY2028＋9/12 × FY2029（情境加權），與同業前瞻 P/E 同口徑"'); G["pe12"] = r; r += 1
    V.cell(r, 1, "現行相對倍數（現價 ÷ 12 個月前瞻 EPS ÷ 同業前瞻 P/E 中位數；原則 47）"); V.cell(r, 2, f"=px/B{fe}/B{pm}").number_format = "0.00"
    V.cell(r, 5, '="10 年相對倍數比（調整係數）"&TEXT(rm_factor,"0.00")&"；調整後 Comparable 需要的相對重估 "&TEXT(rm_factor/B' + str(r) + ',"0.0")&" 倍"'); G["rel"] = r; r += 1
    V.cell(r, 1, "DCF 在加權目標價中的占比（金額）"); V.cell(r, 2, f"=tp_w_dcf*B{dcf}/B{tp}").number_format = "0.0%"
    V.cell(r, 5, '="權重 "&TEXT(tp_w_dcf,"0%")&"；Comparable（相對倍數調整後）占 "&TEXT(1-tp_w_dcf*B' + str(dcf) + '/B' + str(tp) + ',"0.0%")&"（原則 25：進加權的 Comparable 是同業中位數 × 相對倍數調整的混合法）"'); G["share"] = r; r += 1
    pe = [lab(V, n) for n in ("AMD", "Broadcom", "Marvell", "台積電 ADR", "Qualcomm")]
    V.cell(r, 1, "子群中位數（純展示）：GPU／客製 ASIC 設計（AMD、Broadcom、Marvell）"); V.cell(r, 2, f"=MEDIAN(C{pe[0]},C{pe[1]},C{pe[2]})").number_format = "0.00"; G["sub1"] = r; r += 1
    V.cell(r, 1, "子群中位數（純展示）：代工與手機 SoC（台積電 ADR、Qualcomm）"); V.cell(r, 2, f"=MEDIAN(C{pe[3]},C{pe[4]})").number_format = "0.00"; G["sub2"] = r; r += 1
    V.cell(r, 1, "同業口徑（原則 7、30）"); V.cell(r, 2, "前瞻 P/E：價格日 2026-11-05、EPS 為資料商 NTM 共識（2026-10-27～10-30 頁面）。AMD、Broadcom、Marvell、Qualcomm 為 US GAAP（各家 non-GAAP 口徑不一），台積電 ADR 為 TIFRS；P/E 不涉 EV 與租賃處理。各同業成長率、EBITDA 率、淨負債 ÷ EBITDA、資本強度：未取得（同業少於 8 家不做迴歸，X8）")
    G["peer"] = r
    return G


# ---------- 原則 40（部分）：現金稅率並列（DCF 頁）----------
def dcf_tax(wb):
    D = wb["DCF"]; r = D.max_row + 3
    D.cell(r, 1, "EBIT 口徑與現金稅率（v0.3.5 起；原則 40）").font = BOLD; r += 1
    D.cell(r, 1, "EBIT 起點＝GAAP 營業利益（含股票薪酬與併購攤銷，皆為費用），不是 non-GAAP；權益證券損益不進 NOPAT。non-GAAP 調節橋不適用（本模型無加回項）"); r += 1
    D.cell(r, 1, "現金稅率＝FY2026 已付所得稅 ÷ GAAP 稅前利益（不含權益證券損益）"); D.cell(r, 2, "=h_taxpaid_26/(Q!F11+Q!F15)").number_format = "0.0%"
    dn(wb, "cash_tx", f"DCF!$B${r}")
    D.cell(r, 3, '="模型稅率 tx_f "&TEXT(tx_f,"0.0%")&"；差 "&TEXT((tx_f-B' + str(r) + ')*100,"0.0")&" 個百分點"')
    D.cell(r, 4, '=IF(ABS(tx_f-B' + str(r) + ')>0.03,"差距大於 3 個百分點：應加 Sens","差距在 3 個百分點內，不另建 Sens（Sens 業務主因表已有 tx_f 16%／18%）")'); r += 1
    D.cell(r, 1, "DCF 淨現金核對（本頁上方檢查列）的資產負債科目：Hist 頁 h_bs_*（" + Q10N + "）")
    return r


def dcf_check_links(wb):
    """原則 14：DCF 檢查列「淨現金對資產負債表」的常數改引用 Hist"""
    D = wb["DCF"]
    old = "22.443+34.143+42.783+51.157-(1+32.366+0.134)-(0.509+4.985)"
    new = "h_bs_cash+h_bs_mds+h_bs_mes+h_bs_nms-(h_bs_std+h_bs_ltd+h_bs_disc)-(h_bs_lc+h_bs_lnc)"
    n = 0
    for row in D.iter_rows():
        for c in row:
            if isinstance(c.value, str) and old in c.value:
                c.value = c.value.replace(old, new); n += 1
    assert n >= 1, n


# ---------- 原則 23：供應預付與承諾另列（v0.3.6；In 頁說明＋Recon）----------
def supply(wb):
    RC = wb["Recon"]; r = RC.max_row + 2
    RC.cell(r, 1, "H. 供應預付款與長期供應承諾（v0.3.6 起；原則 23：另列，不混入營運資金比率）").font = BOLD; r += 1
    RC.cell(r, 1, "供應與產能承諾（2026-07-26，US$B）"); RC.cell(r, 2, "=h_supply_c").number_format = "#,##0"; RC.cell(r, 3, "=h_supply_c_q1").number_format = "#,##0"
    RC.cell(r, 9, Q10N + " Note 10"); RC.cell(r, 10, "C 欄＝前一季（2026-04-26）。承諾主要為記憶體與製造產能，部分可取消或調整；不進 FCFF（未來採購隨營收認列進銷貨成本）"); r += 1
    RC.cell(r, 1, "供應預付款（單獨金額）"); RC.cell(r, 2, "未揭露"); RC.cell(r, 9, Q10N + "、" + K10N)
    RC.cell(r, 10, '="10-Q／10-K 未單獨揭露供應預付；含在預付及其他流動資產 "&TEXT(h_pre_q2,"0.0")&" 與其他資產 "&TEXT(h_oth_a_q2,"0.0")&"（US$B）內。營運資金比率 fc_wc 的預付及其他流動資產因此無法剔除供應預付（未揭露）"')
    return r


# ---------- README ----------
def readme(wb, ctx, bid, G, SJ):
    RD = wb["README"]; V = wb["Val"]; SE = wb["Sens"]; RC = wb["Recon"]
    VER = "v" + ctx["STAGE"]; V36 = ctx["V36"]; MW = ctx["MW"]; SUMR = ctx["SUMR"]
    tp = f"Val!B{lab(V, '目標價（DCF 與 Comparable 加權）')}"; up = f"Val!B{lab(V, '上檔空間：加權目標價')}"
    rec = f"Val!B{lab(V, '投資建議')}"; dcf = f"Val!B{lab(V, 'DCF 目標價（情境加權）')}"; cu = f"Val!B{lab(V, 'Comparable 目標價（市場主流同業法）')}"
    pmed = f"Val!B{lab(V, '同業前瞻 P/E 中位數')}"; band = lab(V, "價格帶（DCF")
    feps = f"Val!B{lab(V, '12 個月後的前瞻 EPS（情境加權）')}"
    eps = lambda s, c: f"PL!{c}{SUMR[(s, 'eps')]}"; gmr = lambda s, c: f"PL!{c}{SUMR[(s, 'gm')]}"
    SG = '"+0.0%;-0.0%"'
    rmrel = f"Val!B{lab(V, '現行相對倍數')}"           # 現行相對倍數（Val 0.6 節）
    noadj = f"(tp_w_dcf*{dcf}+(1-tp_w_dcf)*{cu})"
    FW = f'''"現價約為 12 個月前瞻 EPS（情境加權 "&TEXT({feps},"0.00")&"，"&TEXT(h_asof,"yyyy-mm-dd")&" 起 12 個月）的 "&TEXT(px/{feps},"0.0")&" 倍（分母統一用 12 個月前瞻情境加權 EPS，與同業前瞻 P/E 同口徑；以 FY2028 情境 B EPS 為分母為 "&TEXT(px/{eps("B", "E")},"0.0")&" 倍；擁有者原命題寫 "&TEXT(th_mult,"0.0")&" 倍，原文分母為「2027 年 EPS」）"'''
    REC5 = lambda x: f'IF({x}>0.5,"強力買進",IF({x}>0.25,"買進",IF({x}>=-0.25,"中立",IF({x}>=-0.5,"放空","強力放空"))))'
    # 原則 41：首列三個日期＋build_id（公式）
    a1 = str(RD["A1"].value)
    head, _, rest = a1.partition("　2026-10-25；")
    head = head.replace("收支與估值模型 v0.3.4", "個股模型 " + VER)
    RD["A1"].value = ("=" + lit(head) + '&"　資料截止日 "&TEXT(h_cutoff,"yyyy-mm-dd")&"；成品日 "&TEXT(h_built,"yyyy-mm-dd")&"；股價基準日 "&TEXT(h_asof,"yyyy-mm-dd")&"（v0.3.5 起三個日期分列，原則 41）；build_id '
                      + bid + '。版本：2026-10-25 首版；"&' + lit(rest + f"；v0.3.5：驗收採用規劃書展示與治理（數值不變）" + ("；v0.3.6：驗收採用規劃書數值項（評價股數、目標時點、折舊與資本支出比率、FY↔CY 月數加權）" if V36 else "")))
    for r in range(1, RD.max_row + 1):
        v = RD.cell(r, 1).value
        if not isinstance(v, str): continue
        w = v
        if v.startswith("模型命題：") and "現價約為 2027 年 EPS 的 13.5 倍" in v:   # 原則 13、43
            a, b = v.split("現價約為 2027 年 EPS 的 13.5 倍", 1)
            w = ("=" + lit(a) + f'&{FW}&' + lit(b))
        elif v.startswith("會計年度：FY(t)") and MW:   # 原則 49
            w = ("會計年度：FY(t)＝t−1 年 2 月～t 年 1 月（FY2027＝2026-02～2027-01）。上游市場層是日曆年，v0.3.6 起以月數加權：FY(t)＝11/12 × CY(t−1)＋1/12 × CY(t)（原則 49），"
                 "顆數、HBM 與一般 DRAM 價格指數、每顆 HBM 容量、網通金額各自加權（數量 × 單價的交叉項不分年，近似）；上游只到 CY2029，FY2030 的 1/12 以 CY2029 代替（持平近似，不引入新假設）。"
                 "FY2027 錨點的 k_rec、c_nh 以同一口徑重算（k_rec＝錨點 Data Center ÷ 未校準由下而上 FY2027）。季度出貨時點屬上游權威，不在個股估（X29）。預測期 FY2027–FY2030。")
        elif v.startswith("估值：P/E 規則"):   # 原則 42
            w = ("估值：目標價＝DCF（標準版：無槓桿自由現金流 FCFF＋WACC）與 Comparable（同業前瞻 P/E 中位數 × 相對倍數調整係數 × 12 個月後前瞻 EPS）加權，DCF 權重 tp_w_dcf（Andy 2026-10-25 定 50%）；"
                 "DCF 與 Comparable 各自以情境權重 A 0.3／B 0.5／C 0.2 加權。舊法（10 年落後 P/E 依循環位置取分位、規則值情境加權）保留在 Val 頁，標「舊法，不進結論」（v0.3.5 起；原則 42）。")
        elif v.startswith("標記：Verified"):   # 原則 2、34
            w = ("標記：每列單一主要標記（Verified／Interested-party／Analogy／Assumed／Derived），Interested-party 列另有「來源誘因」欄（v0.3.5 起；原則 2）。SemiAnalysis 未單獨使用。"
                 "資產負債與股數以 " + Q10N + " 為主來源（原則 34；Hist 頁），8-K 新聞稿為第二來源；CFO 書面說明與法說逐字稿未取得，相關數字只經搜尋摘要（標二手）。")
        elif v.startswith("已知限制：") and MW:
            w = v.replace("(5) FY↔CY 近似：Rubin 在 CY2026 的 1.7M 顆多數出貨在 FY2027 Q4 之後。", "(5) FY↔CY 以月數加權（v0.3.6），數量與單價分別加權為近似；季度出貨落點屬上游（X29）。")
        elif v.startswith("情境對應（v0.2 延伸期）") and MW:
            w = v.replace("三情境在 FY2027 相同、FY2028 幾乎相同（只差 LPDDR 成本），DCF 的情境差異主要來自 FY2029 之後。", "v0.3.6 起 FY2028 含 1/12 的 CY2028（月數加權），情境差異自 FY2028 起部分出現、主要來自 FY2029 之後。")
        elif v.startswith("重點觀察 3：最敏感的是 NVIDIA 顆數"):   # 原則 13
            b = v.split("下一個檢驗點是", 1)[1].split("：", 1)[0]
            s_u, s_a, s_1 = lab(SE, "NVIDIA 顆數係數 低"), lab(SE, "ASIC 金額份額 +5pt"), lab(SE, "HBM／LPDDR 吸收比例 1")
            w = (f'="重點觀察 3：EPS 敏感度（情境 A 為底）：顆數 ×0.9 時 FY2028 EPS "&TEXT(Sens!E{s_u},{SG})&"；ASIC 金額份額 +5pt 時 "&TEXT(Sens!E{s_a},{SG})&"；HBM／LPDDR 全額吸收時 "&TEXT(Sens!E{s_1},{SG})&"（Sens 頁）。目標價敏感度見 Sens 業務主因表。下一個檢驗點是"&' + lit(b) + '&"：營收對照 Q3 指引 "&TEXT(q3_rev*10,"#,##0")&" 億美元（區間 "&TEXT(q3_rev_lo*10,"#,##0")&"～"&TEXT(q3_rev_hi*10,"#,##0")&"），以及 Q4 指引。"')
        elif v.startswith("="):
            if v.startswith('="重點觀察 2：'):   # 原則 13：手寫 EPS 改公式
                w = (f'="重點觀察 2：FY2028 EPS 情境 A／B／C 為 "&TEXT({eps("A", "E")},"0.00")&"／"&TEXT({eps("B", "E")},"0.00")&"／"&TEXT({eps("C", "E")},"0.00")&"，FY2029 為 "&TEXT({eps("A", "F")},"0.00")&"／"&TEXT({eps("B", "F")},"0.00")&"／"&TEXT({eps("C", "F")},"0.00")'
                     f'&"、FY2030 為 "&TEXT({eps("A", "G")},"0.00")&"／"&TEXT({eps("B", "G")},"0.00")&"／"&TEXT({eps("C", "G")},"0.00")&"，DCF 價格帶 "&TEXT(Val!B{band},"0")&"～"&TEXT(Val!C{band},"0")&"（標準版 DCF）。核心問題（CY2028 顆數持平後 EPS 能否成長）在情境 B 的答案是 FY2029 "'
                     f'&TEXT({eps("B", "F")}/{eps("B", "E")}-1,{SG})&"、FY2030 "&TEXT({eps("B", "G")}/{eps("B", "F")}-1,{SG})&"。"')
            if v.startswith('="命題 4：'):   # 原則 43
                w = (f'="命題 4："&{FW}&"，這是低估，還是見頂前的折價。"')
            if v.startswith('="現價是低估，還是見頂前的折價？'):   # 原則 43：與同業同一分母
                w = re.sub(r'&" 約為本模型情境 B FY2028 EPS 的 "&TEXT\(Val!C\d+,"0\.0"\)&" 倍，低於同業前瞻本益比中位數 "&TEXT\((Val!B\d+),"0\.0"\)&" 倍（同業為目前的未來 12 個月口徑，與 FY2028 時點相近但不完全相同）。',
                           lambda m: (f'&" 約為 12 個月前瞻 EPS（情境加權 "&TEXT({feps},"0.00")&"）的 "&TEXT(px/{feps},"0.0")&" 倍，"&IF(px/{feps}<{m.group(1)},"低於","高於")&"同業前瞻本益比中位數 "&TEXT({m.group(1)},"0.0")&" 倍（同一 12 個月前瞻口徑；以 FY2028 情境 B EPS 為分母為 "&TEXT(px/{eps("B", "E")},"0.0")&" 倍）。'), v)
                assert w != v
            if v.startswith('="重點 2：三種情境在 FY2028 幾乎一樣'):   # 原則 31：方向詞由表格判定
                w = re.sub(r'^="重點 2：三種情境在 FY2028 幾乎一樣、分歧從 FY2029 開始：FY2028 EPS 約 "&TEXT\(PL!E\d+,"0\.00"\)&" 美元；',
                           lambda m: (f'="重點 2：FY2028 EPS 情境 A／B／C 為 "&TEXT({eps("A", "E")},"0.00")&"／"&TEXT({eps("B", "E")},"0.00")&"／"&TEXT({eps("C", "E")},"0.00")&" 美元（三情境差 "'
                                      f'&TEXT((MAX({eps("A", "E")},{eps("B", "E")},{eps("C", "E")})-MIN({eps("A", "E")},{eps("B", "E")},{eps("C", "E")}))/{eps("B", "E")},"0.0%")&"），分歧主要從 FY2029 開始：'), v)
            if v.startswith('="ASIC 會把 NVIDIA 的金額份額壓到多少？'):   # 原則 31、50
                rs = lab(RC, "2027 年 NVIDIA 金額份額（加速器晶片")
                w = v.replace('&"（模組層級），落在命題區間的下緣；', f'&"（模組層級），"&IF(Recon!C{rs}<th_sh_lo,"低於命題區間",IF(Recon!C{rs}<=th_sh_lo+0.01,"落在命題區間的下緣",IF(Recon!C{rs}<=th_sh_hi,"落在命題區間內","高於命題區間")))&"；')
                w = w.replace('份額是市場層顆數與各家每顆金額假設算出來的，不是公司揭露。', '這是模型輸出、非獨立證據：份額由 Assumed／Interested-party 的 ASIC 每顆金額與市場層顆數推得，不是公司揭露。可證偽指標：NVIDIA 資料中心營收 ÷（NVIDIA 資料中心＋AMD 資料中心 GPU＋Broadcom AI 半導體＋Marvell 客製 ASIC）的季度實績。')
                w = w.replace("本模型（三情境相同）", "本模型（三情境相同，CY2027）")
            if v.startswith('="HBM 漲價會把毛利率壓低多少？'):   # 原則 50：命題回覆依模型結果寫明支持或否定
                rg = lab(RC, "FY2028 毛利率 vs 指引"); rh = lab(RC, "HBM 成本占由下而上營收")
                dA, dB, dC = (f"({gmr(s, 'E')}-{gmr(s, 'D')})" for s in "ABC")
                tail = (f'"&"FY2028 毛利率情境 A "&TEXT({gmr("A", "E")},"0.00%")&"、B "&TEXT({gmr("B", "E")},"0.00%")&"、C "&TEXT({gmr("C", "E")},"0.00%")&"，比 FY2027 錨點 "&TEXT({gmr("B", "D")},"0.00%")&" 總降幅 "'
                        f'&TEXT(-{dA}*100,"0.0")&"／"&TEXT(-{dB}*100,"0.0")&"／"&TEXT(-{dC}*100,"0.0")&" 個百分點；情境 B "&IF(Recon!F{rg}<0,"低於","高於")&"公司 FY2028 指引中點 "&TEXT(g28_gm,"0.0%")&" "&TEXT(ABS(Recon!F{rg})*100,"0.0")&" 個百分點'
                        f'（指引區間 "&TEXT(g28_gm_lo,"0%")&"～"&TEXT(g28_gm_hi,"0%")&"）。同期 HBM 成本占營收（吸收前）上升 "&TEXT(Recon!E{rh}*100,"0.0")&" 個百分點。這是模型輸出、非獨立證據：毛利率由 Assumed 的吸收比例 abs_mem（"&TEXT(abs_mem,"0.0")&"，公司未揭露）與上游記憶體價格路徑推得；'
                        f'全額轉嫁時 FY2028 EPS "&TEXT(Sens!E{lab(SE, "HBM／LPDDR 吸收比例 0")},{SG})&"，全額吸收時 "&TEXT(Sens!E{lab(SE, "HBM／LPDDR 吸收比例 1")},{SG})&"。"')
                if V36 and G["recon"].get("gm_hbm1"):
                    h1 = f"Recon!B{G['recon']['gm_hbm1']}"   # 百分點（負值＝壓低）
                    w = (f'="HBM 漲價會把毛利率壓低多少？答："&IF(-{h1}>3,"命題被否定（HBM 單項降幅 "&TEXT(-{h1},"0.0")&" 個百分點，高於命題的 2～3 個百分點）",IF(-{h1}>=2,"命題被支持（HBM 單項降幅 "&TEXT(-{h1},"0.0")&" 個百分點，在 2～3 個百分點內）","命題被否定（HBM 單項降幅 "&TEXT(-{h1},"0.0")&" 個百分點，小於 2 個百分點）"))'
                         f'&IF(ABS(-{h1}-3)<1,"；結論臨界：距命題上緣 3 個百分點只有 "&TEXT(ABS(-{h1}-3),"0.0")&" 個百分點","")'
                         f'&"（HBM 單項＝HBM 價格效果加吸收比例效果中屬 HBM 的份額，情境 B，Recon G 節）。總降幅另含 LPDDR 與組合：' + tail + '&"差距分解與 abs_mem 反解見 Recon G 節。"')
                else:
                    w = (f'="HBM 漲價會把毛利率壓低多少？答："&IF(-{dB}>0.03,"命題被否定（降幅大於命題的 2～3 個百分點）",IF(-{dB}>=0.02,"命題被支持（降幅在 2～3 個百分點內）","命題被否定（降幅小於 2 個百分點）"))&"：' + tail)
            if v.startswith('="顆數持平以後 EPS 還能不能成長？'):   # 原則 31：方向詞由表格判定
                sf = lab(SE, "CY2029 顆數與 CY2028 持平"); e8, e9, e0 = eps("B", "E"), eps("B", "F"), eps("B", "G")
                w = (f'="顆數持平以後 EPS 還能不能成長？答："&IF(Sens!N{sf}<=0,"不能",IF(Sens!N{sf}<0.05,"能，但幅度小","能"))&"。敏感度案例「CY2029 顆數與 CY2028 持平」（以情境 A 為底、網通同步持平）中，FY2030 EPS 比 FY2029 "&TEXT(Sens!N{sf},{SG})&"（情境 A 原本為 "&TEXT(Sens!N4,{SG})&"）。'
                     f'情境 B（基準）的顆數在 CY2028 先降到原始的 "&TEXT(um28_B,"0.00")&" 倍，FY2029 EPS "&TEXT({e9},"0.00")&" 比 FY2028 的 "&TEXT({e8},"0.00")&" "&IF({e9}<{e8},"低","高")&" "&TEXT(ABS({e9}/{e8}-1),"0.0%")&"；CY2029 顆數回升後，FY2030 為 "&TEXT({e0},"0.00")&"（"&TEXT({e0}/{e9}-1,{SG})&"）。"')
        if w != v: RD.cell(r, 1).value = w
    rr = lab(RD, '="投資建議：')
    s_w3 = lab(SE, "DCF 權重 0.3")
    RD.cell(rr, 1).value = (f'="投資建議："&{rec}&"（依上檔空間的估值評等；加權目標價上檔空間 "&TEXT({up},{SG})&"；未評估借券、流通量與時點）。'
                            f'條件式（原則 3）：在 DCF 權重 "&TEXT(tp_w_dcf,"0%")&"、Comparable 採相對倍數調整（係數 "&TEXT(rm_factor,"0.00")&"）、k_rec 校準殘差自 FY2028 起持續、HBM 漲價吸收比例 "&TEXT(abs_mem,"0.0")&" 的基準假設下成立；'
                            f'移除相對倍數調整時評等為"&{REC5(noadj + "/px-1")}&"（"&TEXT({noadj}/px-1,{SG})&"），DCF 權重 0.3 時為"&{REC5(f"Sens!C{s_w3}/px-1")}&"。主要失效條件見本頁末。"')
    # 版本沿革列（插在 v0.3.4 列前）
    h = lab(RD, "v0.3.4：")
    RD.insert_rows(h)
    RD.cell(h, 1, f"v0.3.5（{BUILT}）：NVIDIA v0.3.4 驗收採用規劃書——展示與治理（數值不變；加權目標價、DCF、Comparable 不變）：build_id、輸入指紋與模型狀態；In、Q 頁單一主要標記與來源誘因；歷史財務移到 Hist 頁；"
                  "三個日期分列；10 年落後 P/E 標舊法；倍數標分母；README 手寫數字改公式；條件式投資建議與失效條件；上檔空間三段歸因與現行相對倍數（原則 47）；評等邊界距離；Recon 反解；"
                  "命題 2 依模型結果寫明被否定、命題 1 寫明模型輸出；資產負債與股數以 10-Q 為主來源；Sens 目標價表加 DCF 單法美元欄、相對倍數係數＝1 列與業務主因整本重算表；資料基準日與逾期；風險登錄與可證偽指標；展示層數字一致性檢查。")
    if V36:
        RD.insert_rows(h)
        RD.cell(h, 1, f"v0.3.6（{BUILT}）：驗收採用規劃書——數值項：評價股數改流通股＋RSU（原則 11）；目標時點改現價日＋12 個月（原則 26）；折舊與資本支出比率改最近四季（原則 35）；FY↔CY 改月數加權 11/12＋1/12、k_rec 與 c_nh 依新口徑重算（原則 49）；"
                      "Sens 新增 k_rec 線性衰減案例（原則 48）；Recon 新增 FY2028 毛利率與指引差距分解與 abs_mem 反解（原則 50）；供應預付與承諾另列（原則 23）。目標價變動拆解見下一列。")
        if SJ and SJ.get("attr"):
            RD.insert_rows(h + 1); RD.cell(h + 1, 1, SJ["attr"])
    # 新增各段（本頁末）
    r = RD.max_row + 2

    def put(t, bold=False):
        nonlocal r
        RD.cell(r, 1, t).font = Font(bold=bold); RD.cell(r, 1).alignment = Alignment(wrap_text=True, vertical="top"); r += 1
        return r - 1
    G["h"] = put(f"{VER} 驗收採用項（HTML 用；NVIDIA v0.3.4 驗收報告，採用規劃書 2026-11-11）", True)
    G["state"] = put(f'="build_id "&Sens!D{G["gov"]["bid"]}&"；模型狀態 "&Sens!D{G["gov"]["st"]}&"（Sens 業務主因表與 HTML 為建置快照：只在 Excel 改 In 頁而未重跑 python3 nvidia/sens_nvidia.py 與建置時顯示 STALE）"')
    u = up; e = f"IF({u}<-0.375,-0.5,IF({u}<0,-0.25,IF({u}<0.375,0.25,0.5)))"; d = f"ABS({u}-{e})"
    G["edge"] = put(f'=IF({d}<0.05,"評等貼近級距邊界（原則 36）：加權目標價上檔空間 "&TEXT({u},{SG})&" 距 "&TEXT({e},"+0%;-0%")&" 邊界只有 "&TEXT({d}*100,"0.0")&" 個百分點；加權目標價達 "&TEXT(px*(1+{e}),"0.00")&" 美元（目前 "&TEXT({tp},"0.00")&"）即翻轉評等為"&IF({u}<{e},IF({e}=0.5,"強力買進",IF({e}=0.25,"買進","中立")),IF({e}=-0.5,"強力放空",IF({e}=-0.25,"放空","中立")))&"。",'
                    f'"評等邊界（原則 36）：上檔空間距最近級距邊界 "&TEXT({d}*100,"0.0")&" 個百分點（5 個百分點以上，評等不貼近邊界）；加權目標價達 "&TEXT(px*(1+{e}),"0.00")&" 美元時翻轉。")')
    # 原則 47：上檔空間三段歸因與現行相對倍數
    G["attr47"] = put(f'="上檔空間三段歸因（原則 47）：加權上檔空間 "&TEXT({up},{SG})&"＝DCF 貢獻 "&TEXT(tp_w_dcf*({dcf}-px)/px,{SG})&"＋未調整 Comparable 貢獻 "&TEXT((1-tp_w_dcf)*({cu}-px)/px,{SG})&"＋相對倍數調整貢獻 "&TEXT((1-tp_w_dcf)*(tp_comp_adj-{cu})/px,{SG})&"。'
                      f'移除相對倍數調整時加權目標價 "&TEXT({noadj},"0.00")&"（"&TEXT({noadj}/px-1,{SG})&"），評等由"&{rec}&IF({REC5(noadj + "/px-1")}={rec},"維持不變","翻轉為"&{REC5(noadj + "/px-1")})&"。"')
    G["rel47"] = put(f'="現行相對倍數與 10 年比值（原則 47）：現價 ÷ 12 個月前瞻 EPS（情境加權 "&TEXT({feps},"0.00")&"）÷ 同業前瞻 P/E 中位數 "&TEXT({pmed},"0.00")&"＝"&TEXT({rmrel},"0.00")&"；10 年相對倍數比 "&TEXT(rm_factor,"0.00")&"（Bloomberg 前瞻倍數，Val 第 0.5 節）。'
                     f'調整後 Comparable 需要 NVIDIA 相對同業的倍數從 "&TEXT({rmrel},"0.00")&" 重估到 "&TEXT(rm_factor,"0.00")&"，即相對重估 "&TEXT(rm_factor/{rmrel},"0.0")&" 倍。'
                     f'兩法對競爭優勢的隱含假設相反：DCF 用 RONIC＝WACC（終值期無超額報酬），調整後 Comparable 用歷史溢價（假設超額報酬延續）；相對倍數調整仍在主估值（Andy J1 決定；X30）。Sens 業務主因表第一列為「相對倍數係數＝1」。"')
    # 原則 50：基準與公司指引的差距
    gB, g27 = gmr("B", "E"), f"PL!E{SUMR[('B', 'g')]}"
    G["guide50"] = put(f'="基準與公司指引（原則 50）：FY2028 毛利率（情境 B）"&TEXT({gB},"0.0%")&IF({gB}<g28_gm_lo,"，低於指引區間 "&TEXT(g28_gm_lo,"0%")&"～"&TEXT(g28_gm_hi,"0%")&" 下緣 "&TEXT((g28_gm_lo-{gB})*100,"0.0")&" 個百分點",IF({gB}>g28_gm_hi,"，高於指引上緣 "&TEXT(({gB}-g28_gm_hi)*100,"0.0")&" 個百分點","，落在指引區間內"))'
                       f'&"；FY2028 營收年增（情境 B）"&TEXT({g27},"0.0%")&IF({g27}<g28_gr,"，低於","，高於")&"公司說法約 "&TEXT(g28_gr,"0%")&" "&TEXT(ABS({g27}-g28_gr)*100,"0.0")&" 個百分點。指引只作對照、不校準（X3）"&IF({"TRUE" if V36 else "FALSE"},"；毛利率差距分解與 abs_mem 反解見 Recon G 節。","。")')
    G["asof_h"] = put(f"資料基準日（as_of）：{AS_OF}（現價日）。市場資料與基準日相差超過 2 個交易日標「逾期」；最近一次財報（Q2 FY27，2026-08-26）之後的共識不標財報前（原則 15）。", True)
    RD.cell(r, 1, "項目").font = BOLD; RD.cell(r, 2, "資料日期").font = BOLD; RD.cell(r, 3, "狀態").font = BOLD; G["asof_t"] = r; r += 1
    items = [("股價 px（Yahoo Finance 日線收盤）", "2026-11-05", 0), ("同業股價（五家；Comparable 前瞻 P/E 換算）", "2026-11-05", 0),
             ("同業前瞻 EPS（stockanalysis 統計頁，最早一頁）", "2026-10-27", 0), ("無風險利率 rf（美國 10 年期公債殖利率）", "2026-10-23", 0),
             ("相對倍數歷史資料（Bloomberg 前瞻倍數，10 年）", "2026-10-29", 0), ("市場共識目標價 cons_tp（stockanalysis）", "2026-10-25", 0),
             ("共識 FY2028 EPS（247wallst）", "2026-10-19", 0), ("資產負債與股數（10-Q）", "2026-07-26", 2)]
    G["asof_rows"] = []
    for labl, dd, kind in items:
        RD.cell(r, 1, labl); RD.cell(r, 2, dd)
        RD.cell(r, 3, '="最近一季資產負債表（以 stub 期滾到評價日）"' if kind == 2 else
                f'=IF(ABS(NETWORKDAYS(MIN(DATEVALUE(B{r}),DATEVALUE("{AS_OF}")),MAX(DATEVALUE(B{r}),DATEVALUE("{AS_OF}")))-1)>2,"逾期 "&TEXT(ABS(NETWORKDAYS(MIN(DATEVALUE(B{r}),DATEVALUE("{AS_OF}")),MAX(DATEVALUE(B{r}),DATEVALUE("{AS_OF}")))-1),"0")&" 個交易日","同期")')
        G["asof_rows"].append(r); r += 1
    G["asof_n"] = put("說明：rf、同業前瞻 EPS、相對倍數資料與共識逾期，但與其他 AI半導體 個股模型同一取值日，為維持跨模型一致暫不單獨更新（原則 15；不需驗收清單 X24）；同業 P/E 已以 2026-11-05 收盤換算。共識只作對帳，不進目標價。")
    inv = G["recon"]["inv"]
    G["fail_h"] = put("主要失效條件（可觀察；原則 3、20）", True)
    E_up, E_dn = f"Recon!B{G['recon']['e_up']}", f"Recon!B{G['recon']['e_dn']}"
    fails = [f'="(1) 相對倍數：NVIDIA 對同業的前瞻倍數比（現價 ÷ 12 個月前瞻 EPS ÷ 同業中位數；目前 "&TEXT({rmrel},"0.00")&"）是否回到 1.5 倍以上；調整後 Comparable 假設回到 10 年比值 "&TEXT(rm_factor,"0.00")&"。評等降一級（上檔 "&TEXT({E_dn},"+0%;-0%")&"）所需的相對倍數係數約 "&TEXT(Recon!B{inv["rm_dn"]},"0.00")&"，支持現價需約 "&TEXT(Recon!B{inv["rm_px"]},"0.00")&"。"',
             f'="(2) 毛利率：FY2028 實際毛利率是否 ≥ "&TEXT(g28_gm_lo,"0%")&"（公司指引下緣）；模型情境 B 為 "&TEXT({gB},"0.0%")&"。高於指引下緣表示吸收比例 abs_mem 或非記憶體成本假設偏保守。"',
             f'="(3) DCF：DCF 目標價（目前 "&TEXT(Recon!C{inv["dcf_up"]},"0.00")&"）升到約 "&TEXT(Recon!B{inv["dcf_up"]},"0.00")&" 評等升一級（上檔 "&TEXT({E_up},"+0%;-0%")&"；Comparable 不變）"&IF(Recon!B{inv["dcf_dn"]}>0,"，降到約 "&TEXT(Recon!B{inv["dcf_dn"]},"0.00")&" 降一級","；即使 DCF 為 0 也不會降一級（Comparable 不變時）")&"。可觀察：WACC 組成（rf、自算產業 β）與 FY2028–FY2030 營收、營業利益率。"',
             f'="(4) 倍數：Comparable 採用倍數（同業中位數 × 相對倍數係數；目前 "&TEXT({pmed}*rm_factor,"0.0")&" 倍）降到約 "&TEXT(Recon!B{inv["pe_dn"]},"0.0")&" 倍評等降一級；支持現價需約 "&TEXT(Recon!B{inv["pe_px"]},"0.0")&" 倍。可觀察：同業前瞻本益比中位數。"',
             f'="(5) 顆數與 k_rec：Q3 FY27 財報（約 2026-11 下旬）營收對照 Q3 指引 "&TEXT(q3_rev*10,"#,##0")&" 億美元（區間 "&TEXT(q3_rev_lo*10,"#,##0")&"～"&TEXT(q3_rev_hi*10,"#,##0")&"）；Rubin 出貨與認列節奏若使 k_rec 缺口收斂，目標價上升（Sens 業務主因表「k_rec 衰減」）。"']
    G["fail"] = [put(t) for t in fails]
    G["frame"] = put(f'="估值框架並列：100% DCF＝"&TEXT({dcf},"0.00")&" 美元、100% Comparable（相對倍數調整後）＝"&TEXT(tp_comp_adj,"0.00")&" 美元、只看 Comparable（未調整）"&TEXT({cu},"0.00")&" 美元、加權（DCF "&TEXT(tp_w_dcf,"0%")&"）＝"&TEXT({tp},"0.00")&" 美元；DCF 情境價格帶 "&TEXT(Val!B{band},"0.0")&"～"&TEXT(Val!C{band},"0.0")&" 美元。"')
    G["risk_h"] = put("風險登錄（原則 10、19；是否建模、理由、觀察指標）", True)
    G["risk"] = [put(t) for t in (
        '="中國出口管制與 H200 授權：建模為 Sens（中國恢復每年 "&TEXT(cn_amt_lo,"0")&"／"&TEXT(cn_amt,"0")&"／"&TEXT(cn_amt_hi,"0")&" 十億美元，基準 0＝公司指引假設；X31）。觀察：美國與中國的授權與海關政策、公司季度指引是否納入中國資料中心營收。"',
        "客戶集中（10-Q 揭露多家直接客戶各占營收一成以上）：未建模（個別客戶未具名、需求時點未揭露）。觀察：10-Q 客戶集中附註、主要雲端廠資本支出指引。",
        '="供應與產能承諾 "&TEXT(h_supply_c,"#,##0")&" 十億美元（Q2 FY27；前一季 "&TEXT(h_supply_c_q1,"#,##0")&"）：未進 FCFF（隨採購認列進成本）；需求下滑時可能產生超額採購損失（FY2026 H20 超額採購費用為先例）。觀察：Note 10 承諾金額、存貨與超額採購損失。"',
        "CoWoS／HBM 供給上限：讀 Siliconomics／markets/memory 共用顆數與價格路徑，不在個股估（X29）。觀察：台積電 CoWoS 產能分配、HBM4 合約價。",
        "客戶融資與延長付款條件（10-Q 流動性段）：未建模（金額未揭露）。觀察：應收帳款天數、10-Q 財務保證與信用支持揭露。",
        "流通量、借券與軋空：不進目標價（估值評等不含交易執行；X14）。觀察：借券費率、空單比例。")]
    G["ind_h"] = put("主要假設驅動與可證偽的公開指標（原則 20）", True)
    G["ind"] = [put(t) for t in (
        "k_rec（Derived，校準殘差、FY2028 起持續）：季度 Data Center 營收 ÷ 市場層 NVIDIA 顆數 × 橋接表售價；Rubin 出貨節奏（CFO 書面說明）。",
        '="HBM／LPDDR 吸收比例 abs_mem（Assumed，基準 "&TEXT(abs_mem,"0.0")&"）：公司季度毛利率與 FY2028 毛利率指引 "&TEXT(g28_gm_lo,"0%")&"～"&TEXT(g28_gm_hi,"0%")&" 的實現。"',
        "Rubin Ultra 模組售價 pm_ru（Assumed）：Rubin Ultra 機架報價（OEM、ODM 揭露）。",
        "匣級加價保留比例 k_tray（Assumed）：ODM（鴻海、廣達、緯創）AI 伺服器毛利率與認列方式。",
        "ASIC 每顆金額 pa_*（Siliconomics，Assumed／Interested-party）：Broadcom AI 半導體、Marvell 客製 ASIC、Google TPU 對外報價。",
        "Edge 毛利率與成長 gm_edge、eg（Assumed）：Edge Computing 平台季度營收；公司未揭露平台毛利率。",
        "相對倍數調整（10 年比值）：NVIDIA 對同業前瞻倍數比的月底值（Bloomberg BF）。")]
    return r


def apply(wb, here, ctx):
    V36, OVR = ctx["V36"], ctx["OVR"]
    sj_path = os.path.join(here, "sens.json")
    SJ = json.load(open(sj_path, encoding="utf-8")) if (os.path.exists(sj_path) and not OVR) else None
    if SJ and SJ.get("_stage") != ctx["STAGE"]: SJ = None
    hist_sheet(wb); hist_links(wb, ctx)
    if V36: v36_values(wb, ctx)
    overrides(wb, OVR)
    dcf_check_links(wb)
    n = norm_tags(wb)
    IN = wb["In"]
    IN["A2"].value = ("標記：每列單一主要標記 Verified／Interested-party／Analogy／Assumed／Derived（v0.3.5 起，原則 2；Interested-party 列見 J 欄「來源誘因」）；Analogy、Assumed 給低／高區間。SemiAnalysis 未單獨使用。"
                      "公司原文＝NVIDIA 8-K 新聞稿與 SEC 10-Q／10-K（資產負債與股數以 10-Q 為主來源，原則 34；歷史財務在 Hist 頁）；CFO 書面說明與法說逐字稿未取得")
    r = inrow(wb, "th_mult"); IN.cell(r, 9, "命題原文分母為「2027 年 EPS」（未標情境）。現行值統一以 12 個月前瞻情境加權 EPS 為分母（Val 0.6 節「現價 ÷ 12 個月前瞻 EPS」）；各情境 FY2028 口徑見 Val 第 1 節（原則 6、43：舊說明「約 13.4–13.5x」已過時，刪除）")
    RC = wb["Recon"]; rr = lab(RC, "現價 ÷ FY2028 EPS vs 命題")
    RC.cell(rr, 10).value = "分母：C／E／G 欄＝各情境 FY2028 EPS（FY 口徑，2026-11-05 現價）；統一口徑（12 個月前瞻情境加權 EPS）的倍數見 Val 0.6 節。EPS 越高，現價越便宜"
    r = inrow(wb, "rf"); IN.cell(r, 9, str(IN.cell(r, 9).value or "") + "。期限：10 年期；取值日 2026-10-23（股價日 2026-11-05，逾期；與其他個股模型同一取值日，見 README 資料基準日表，X24）")
    r = inrow(wb, "px"); IN.cell(r, 9, str(IN.cell(r, 9).value or "") + "。定位：Yahoo Finance NVDA 日線 2026-11-05 Close 181.60（未調整；原則 45）")
    for k, add in (("q3_rev", "定位：Q2 FY27 新聞稿 Outlook 段：Revenue is expected to be $108.0 billion, plus or minus 2%（原則 45）"),
                   ("q3_gm", "定位：Q2 FY27 新聞稿 Outlook 段：GAAP and non-GAAP gross margins 74.0%, plus or minus 50 basis points"),
                   ("beta_u", "依據：docs/plan/自算產業β_20261105.md（US 半導體 13 家；NVIDIA 本身在同業組內）")):
        c = IN.cell(inrow(wb, k), 9); c.value = str(c.value or "") + "。" + add
    SO = wb["Sources"]
    SO["A1"].value = "來源（公司原文＝NVIDIA 8-K 新聞稿與 SEC 10-Q／10-K，已讀取；資產負債與股數以 10-Q 為主來源（原則 34）；CFO 書面說明、法說逐字稿未取得，其內容只經搜尋摘要）"
    sr = SO.max_row + 1
    for a, b in ((Q10N + "（資產負債、股數、承諾、股權激勵）[Verified]", Q10), (K10N + "（現金流量、已付所得稅、股權激勵計畫）[Verified]", K10),
                 (Q3_10Q + "（前三季折舊與資本支出）[Verified]", SEC + "000104581025000230/"), ("SEC XBRL companyfacts（CIK 0001045810；第二來源）", "https://data.sec.gov/api/xbrl/companyfacts/CIK0001045810.json")):
        SO.cell(sr, 1, a); SO.cell(sr, 2, b); sr += 1
    PLs = wb["PL"]
    if ctx["MW"]:   # 驗收核對（2026-11-11）：情境 B／C 的 FY2027 校準恆等式差額
        r = PLs.max_row + 2; PB = ctx["PB"]
        PLs.cell(r, 1, "核對（情境 B／C；v0.3.6 起）：FY2027 由下而上毛利率 − 錨點毛利率；FY2027 由下而上營收 − 錨點營收。k_rec、c_nh 只以情境 A 校準（上方核對列為 0）；"
                       "月數加權後 FY2027 含 1/12 的 CY2027，B、C 的一般 DRAM 價格指數 cdx 與 A 不同，故有小差額（不影響報告值：FY2027 報告值三情境皆為錨點）").font = BOLD; r += 1
        for s_ in "BC":
            PLs.cell(r, 1, f"  情境 {s_}：毛利率差"); PLs[f"D{r}"] = f"=D{PB[s_]['bugm']}-D{PB[s_]['gm']}"; PLs[f"D{r}"].number_format = "0.0000"; r += 1
            PLs.cell(r, 1, f"  情境 {s_}：營收差（US$B）"); PLs[f"D{r}"] = f"=D{PB[s_]['bur']}-a27_rev"; PLs[f"D{r}"].number_format = "0.0000"; r += 1
    if ctx["MW"]:
        PLs["A1"].value = ("PL — NVIDIA 損益（US$B；FY2026A、FY2027 為實績與錨點，FY2028 起為模型，顆數用認列顆數＝市場層顆數 × k_rec）。v0.3.6 起 FY(t)＝11/12 × CY(t−1)＋1/12 × CY(t)（原則 49；FY2030 的 1/12 以 CY2029 代替）。"
                           "三情境各一個完整推導區塊（本頁下方），頂端為摘要。FY2028 含 1/12 的 CY2028，情境係數自 FY2028 起部分作用")
        UN = wb["Units"]
        UN["A1"].value = str(UN["A1"].value).replace("NVIDIA 會計年度 FY(t) 對照日曆年 t−1（FY2027＝CY2026 … FY2030＝CY2029）", "本頁為日曆年；PL 引擎以 FY(t)＝11/12 × CY(t−1)＋1/12 × CY(t) 換算（v0.3.6 起，原則 49）")
    bid = build_id(here, ctx["STAGE"])
    G = {}
    G["gov"] = governance(wb, SJ, bid, OVR)
    G["sens"] = sens_tp(wb, SJ, ctx)
    G["recon"] = recon(wb, SJ, ctx)
    if V36: supply(wb)
    G["val"] = val(wb, ctx)
    G["tax"] = dcf_tax(wb)
    readme(wb, ctx, bid, G, SJ)
    SE = wb["Sens"]
    SE["A1"].value = ("Sens — 一次動一個輸入（情境 A，k_rec 校準後的基準）：對 FY2028、FY2029、FY2030 的營收與 EPS 的影響；「舊法隱含價」欄＝EPS × 情境 A 的 10 年落後 P/E 規則倍數（舊法，只作參考，不進結論）。"
                      "每個案例是一個完整的推導區塊（本頁下方，與 PL 同一引擎）。目標價的敏感度見下方「目標價參數敏感度」與「業務主因（整本重算）」兩表；模型狀態與 build_id 見本頁底部")
    for c in (6, 7, 11, 12):
        v = SE.cell(3, c).value
        if isinstance(v, str) and "隱含價" in v:
            SE.cell(3, c).value = ("舊法 " + v.replace("（規則值，不適用，僅供對照）", "（參考）").replace("（規則值，不適用）", "（參考）")) if not v.startswith("隱含價相對") else "舊法隱含價相對基準（參考）"
    rn = lab(SE, "讀法：FY2028")
    SE.cell(rn, 1).value = (("讀法：v0.3.6 起 FY(t)＝11/12 × CY(t−1)＋1/12 × CY(t)，FY2028 含 1/12 的 CY2028，情境 B、C 的差異自 FY2028 起部分出現、主要從 FY2029 起；" if ctx["MW"] else
                             "讀法：FY2028（≈CY2027）三情境幾乎相同；情境 B、C 的差異從 FY2029 起；") +
                            "顆數係數、售價與吸收比例的影響從 FY2028 起。舊法隱含價欄只反映 EPS 變動（舊法，不進結論）。售價案例（Rubin Ultra 等）的非 HBM 成本隨基準售價等比例（c_nh × 售價），毛利率不因售價單獨改變（原則 44：已兩邊同動）；"
                            "HBM 吸收比例案例只動售價中的轉嫁部分，成本不變，屬單邊上界。")
    return f"v035_post（{ctx['STAGE']}）：標記正規化 {n} 格；build_id {bid}；快照 {'有' if SJ else '無'}"
