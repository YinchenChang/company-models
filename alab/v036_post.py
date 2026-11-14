# Astera Labs v0.3.6／v0.3.7（2026-11-09）後處理：Astera Labs v0.3.5 驗收報告的採用項。
# 採用規劃書：Drive 065-ALAB/驗收/20261109_ALAB驗收v0.3.5_採用規劃書.md（Andy 2026-11-09 23:22 確認，J1–J4 照建議預設）。
# build_alab.py 在存檔前呼叫 apply(wb, here, ctx)。ctx["V37"]＝False 時只做展示與治理（數值不變，v0.3.6）；True 時加做數值項（v0.3.7）。
#   v0.3.6：A1 Val 市場隱含 EPS 讀法改公式；A2 命題回覆條件式；A3 PL 標題、README 倍數、rf 說明改公式／更正；A5 目標價 Sens 加 DCF 單法美元欄；
#           A9 反解（閉式）與失效條件；A10 同業並列（未取得者寫明）；B3 Recon 每顆內容改列隱含採用率；B4 投資建議條件式＋風險登錄；
#           B5 β 依據；B6 資料基準日與逾期；B7（J3a）舊法標示；B8 估值框架並列；J4 評等邊界揭露；
#           第 10 節原則 1、2、3、5、10、13、14、15、20（build_id／狀態、單一標記與來源誘因、Hist 頁、as_of、可證偽指標）
#   v0.3.7：A4 目標時點＝現價日＋12 個月；A6 目標價 Sens 業務主因（整本重算快照 sens.json）；A8 資產負債改 10-Q；
#           B2（J2b）gaap_k 改最近四季實績；原則 11 評價日股數（流通股＋RSU＋選擇權庫藏股法）；J1(b) 情境 A／C 帶份額高／低路徑（build_alab.py）
import hashlib, json, os, re
from openpyxl.styles import Font, PatternFill
from openpyxl.workbook.defined_name import DefinedName

BOLD = Font(bold=True); BLUE = Font(color="0000FF"); HDR = PatternFill("solid", fgColor="DDEBF7")
AS_OF, VAL_D = "2026-11-05", "2026-12-31"
PR3 = "https://www.sec.gov/Archives/edgar/data/1736297/000173629726000044/q326exhibit991.htm"
PR4 = "https://www.sec.gov/Archives/edgar/data/1736297/000173629726000005/q425exhibit991.htm"
Q10 = "https://www.sec.gov/Archives/edgar/data/0001736297/000173629726000045/alab-20260930.htm"
SA = "stockanalysis（S&P Global）Astera Labs 年度財報 FY2023–FY2025（2026-11-05 讀取；v0.3.5 起沿用值）"
OPT = "opencapital.sh 整理之 XBRL（ShareBasedCompensationArrangement…OptionsOutstandingNumber／WeightedAverageExercisePrice，1Q26；3Q26 10-Q 本次未讀到選擇權表）"
PEERS = ("CRDO", "MRVL", "AVGO", "RMBS", "MTSI")


def lit(t):   # Excel 公式中單一字串常數上限 255 字元：切段以 & 相接
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


# ---------- 原則 14：Hist 頁（歷史財務、評價用資產負債、股數、日期）----------
HIST_Y = [  # (鍵, 項目, FY2023, FY2024, FY2025, 標記, 來源)
    ("rev", "營收", 0.11579, 0.39629, 0.85253, "Verified", SA + "：損益表"),
    ("da", "折舊攤銷", 0.00178, 0.00315, 0.00683, "Verified", SA + "：現金流量表"),
    ("cx", "資本支出", 0.00276, 0.03425, 0.03754, "Verified", SA + "：現金流量表"),
    ("ar", "應收帳款（年底）", 0.00834, 0.03881, 0.0832, "Verified", SA + "：資產負債表"),
    ("inv", "存貨（年底）", 0.0241, 0.04322, 0.05898, "Verified", SA + "：資產負債表"),
    ("ap", "應付帳款（年底）", 0.00634, 0.02692, 0.04236, "Verified", SA + "：資產負債表"),
]
HIST_P = [  # (鍵, 項目, 值, 標記, 來源)  non-GAAP 毛利與營業利益（營業費用＝毛利 − 營業利益）
    ("a24_gp", "FY2024 non-GAAP 毛利", 0.303544, "Verified", "公司 4Q24／FY2024 新聞稿（v0.3.5 沿用值）"),
    ("a24_oi", "FY2024 non-GAAP 營業利益", 0.119594, "Verified", "同上"),
    ("a25_gp", "FY2025 non-GAAP 毛利", 0.646384, "Verified", "公司 4Q25 新聞稿 " + PR4),
    ("a25_oi", "FY2025 non-GAAP 營業利益", 0.334406, "Verified", "同上"),
    ("q1_gp", "1Q26 non-GAAP 毛利", 0.23564, "Verified", "公司 1Q26 新聞稿（v0.3.5 沿用值）"),
    ("q1_oi", "1Q26 non-GAAP 營業利益", 0.111729, "Verified", "同上"),
    ("q2_gp", "2Q26 non-GAAP 毛利", 0.289259, "Verified", "公司 2Q26 新聞稿"),
    ("q2_oi", "2Q26 non-GAAP 營業利益", 0.153472, "Verified", "同上；與 3Q26 新聞稿比較欄一致"),
    ("q3_gp", "3Q26 non-GAAP 毛利", 0.423642, "Verified", "公司 3Q26 新聞稿 " + PR3),
    ("q3_oi", "3Q26 non-GAAP 營業利益", 0.261646, "Verified", "同上"),
    ("q2_ninc_a", "2Q26 利息與其他收益（報表值）", 0.013577, "Verified", "公司 2Q26 新聞稿"),
    ("q2_ninc_b", "2Q26 公允價值調整（non-GAAP 排除）", 0.0015, "Verified", "同上"),
    ("goi_25sa", "FY2025 GAAP 營業利益（資料商值）", 0.17342, "Verified", SA + "：損益表（v0.3.5 gaap_k 用值）"),
    ("goi_25", "FY2025 GAAP 營業利益", 0.173423, "Verified", "公司 4Q25 新聞稿 " + PR4),
    ("goi_4q25", "4Q25 GAAP 營業利益", 0.066958, "Verified", "同上"),
    ("noi_4q25", "4Q25 non-GAAP 營業利益", 0.108907, "Verified", "同上"),
    ("goi_9m26", "1Q26–3Q26 GAAP 營業利益", 0.330137, "Verified", "公司 3Q26 新聞稿 " + PR3),
    ("noi_9m26", "1Q26–3Q26 non-GAAP 營業利益", 0.526845, "Verified", "同上"),
    ("goi_3q26", "3Q26 GAAP 營業利益", 0.179056, "Verified", "同上"),
    ("sbc_9m26", "1Q26–3Q26 股權報酬", 0.195495, "Verified", "同上"),
    ("cash", "現金及約當現金（2026-09-30）", 0.209396, "Verified", "公司 10-Q 3Q26 資產負債表 " + Q10),
    ("ms", "有價證券（流動；2026-09-30）", 1.168847, "Verified", "同上（無非流動有價證券列）"),
    ("lsc", "營業租賃負債：流動（2026-09-30）", 0.006984, "Verified", "公司 10-Q 3Q26 租賃附註"),
    ("lsn", "營業租賃負債：非流動（2026-09-30）", 0.038009, "Verified", "同上"),
    ("shout", "流通股數（2026-10-31，十億股）", 0.174782075, "Verified", "公司 10-Q 3Q26 封面"),
    ("rsu", "未既得 RSU（2026-09-30，十億股）", 0.008463, "Verified", "公司 10-Q 3Q26 股權激勵附註（RSU 活動表）"),
    ("psu", "未既得 PSU（2026-09-30，十億股；績效條件，只作敏感度）", 0.000326, "Verified", "同上（PSU 活動表）"),
    ("opt", "流通在外選擇權（2026-03-31，十億股）", 0.00262, "Interested-party", OPT),
    ("optx", "選擇權加權平均履約價（美元）", 0.86, "Interested-party", OPT),
]
INCENT = {"Interested-party": "資料彙整商轉載 XBRL：未經本專案直接核對原文；選擇權數量為 1Q26，晚於基準日兩季"}


def hist_sheet(wb):
    H = wb.create_sheet("Hist", index=wb.sheetnames.index("Q") + 1)
    H["A1"] = ("Hist — 歷史財務與評價用資料（v0.3.6 起；驗收原則 14）：In 頁的折舊、資本支出、營運資金、GAAP 轉換、營業費用與評價用資產負債、股數只引用本頁儲存格"
               "（定義名稱 h_*），不在公式內寫常數。金額：十億美元；股數：十億股")
    for j, h in enumerate(["鍵", "項目", "FY2023", "FY2024", "FY2025", "標記", "來源", "來源誘因"], 1):
        c = H.cell(3, j, h); c.font = BOLD; c.fill = HDR
    r = 4
    for k, labl, a, b, c3, tag, src in HIST_Y:
        H.cell(r, 1, f"h_{k}"); H.cell(r, 2, labl)
        for j, (y, v) in enumerate(zip(("23", "24", "25"), (a, b, c3))):
            c = H.cell(r, 3 + j, v); c.font = BLUE; c.number_format = "#,##0.00000"; dn(wb, f"h_{k}_{y}", f"Hist!${'CDE'[j]}${r}")
        H.cell(r, 6, tag); H.cell(r, 7, src); r += 1
    r += 1
    for j, h in enumerate(["鍵", "項目", "值", "", "", "標記", "來源", "來源誘因"], 1):
        c = H.cell(r, j, h); c.font = BOLD; c.fill = HDR
    r += 1
    for k, labl, v, tag, src in HIST_P:
        H.cell(r, 1, f"h_{k}"); H.cell(r, 2, labl); c = H.cell(r, 3, v); c.font = BLUE; c.number_format = "#,##0.000000"
        dn(wb, f"h_{k}", f"Hist!$C${r}"); H.cell(r, 6, tag); H.cell(r, 7, src)
        if tag == "Interested-party": H.cell(r, 8, INCENT[tag])
        r += 1
    r += 1
    for k, labl, f in (("asof", "資料基準日（現價日）", "=DATE(2026,11,5)"), ("valdate", "DCF 評價日", "=DATE(2026,12,31)"), ("bsdate", "資產負債表日", "=DATE(2026,9,30)")):
        H.cell(r, 1, f"h_{k}"); H.cell(r, 2, labl); H.cell(r, 3, f).number_format = "yyyy-mm-dd"; dn(wb, f"h_{k}", f"Hist!$C${r}"); H.cell(r, 6, "—"); r += 1
    H.column_dimensions["A"].width = 14; H.column_dimensions["B"].width = 46
    for c in "CDE": H.column_dimensions[c].width = 13
    H.column_dimensions["F"].width = 16; H.column_dimensions["G"].width = 90; H.column_dimensions["H"].width = 50
    return H


def in_set(wb, key, col, val):
    wb["In"].cell(inrow(wb, key), col).value = val


def hist_links(wb, OVR=None):
    """原則 14：In 頁公式內的歷史常數改引用 Hist（數值不變）"""
    S = {
        ("fc_da", 4): "=h_da_25/h_rev_25", ("fc_da", 6): "=h_da_23/h_rev_23",
        ("fc_da", 5): "=(h_da_23/h_rev_23+h_da_24/h_rev_24+h_da_25/h_rev_25)/3",
        ("fc_cx", 4): "=h_cx_23/h_rev_23", ("fc_cx", 6): "=h_cx_24/h_rev_24",
        ("fc_cx", 5): "=(h_cx_23/h_rev_23+h_cx_24/h_rev_24+h_cx_25/h_rev_25)/3",
        ("fc_wc", 5): "=((h_ar_23+h_inv_23-h_ap_23)/h_rev_23+(h_ar_24+h_inv_24-h_ap_24)/h_rev_24+(h_ar_25+h_inv_25-h_ap_25)/h_rev_25)/3",
        ("q1_opx", 5): "=h_q1_gp-h_q1_oi", ("q2_opx", 5): "=h_q2_gp-h_q2_oi", ("q3_opx", 5): "=h_q3_gp-h_q3_oi",
        ("a24_opx_in", 5): "=h_a24_gp-h_a24_oi", ("a25_opx_in", 5): "=h_a25_gp-h_a25_oi",
        ("q2_ninc", 5): "=h_q2_ninc_a-h_q2_ninc_b",
        ("gaap_k", 5): "=h_goi_25sa/(a25_gp_in-a25_opx_in)",
        ("stub_frac", 5): "=(h_valdate-h_bsdate)/365",
    }
    for (k, c), f in S.items():
        cell = wb["In"].cell(inrow(wb, k), c)
        if OVR and k in OVR and c == 5: continue   # sens_alab.py 覆寫的基準值保留
        assert isinstance(cell.value, str) and cell.value.startswith("="), (k, c, cell.value)
        cell.value = f
    in_set(wb, "q2_ninc", 7, "Derived"); in_set(wb, "q1_opx", 7, "Derived")
    for k in ("q2_opx", "q3_opx", "a24_opx_in", "a25_opx_in"): in_set(wb, k, 7, "Derived")


# ---------- v0.3.7 數值項 ----------
def v37_values(wb, OVR, P):
    IN = wb["In"]
    if "gaap" in P:
        _gaap(wb, IN, OVR)
    if "bs" in P: _bs(wb, IN, OVR)
    if "sh" in P: _sh(wb, IN, OVR)
    if "tt" in P: _tt(wb, IN)


def _gaap(wb, IN, OVR):
    # J2(b)：GAAP ÷ non-GAAP 營業利益改用最近四季（4Q25–3Q26）實績；低＝FY2025、高＝3Q26 單季
    r = inrow(wb, "gaap_k")
    IN.cell(r, 2, "GAAP 營業利益 ÷ non-GAAP 營業利益（最近四季 4Q25–3Q26；EBIT 換成 GAAP 口徑，股權報酬視為費用；全期持平）")
    IN.cell(r, 4, "=h_goi_25/h_a25_oi"); IN.cell(r, 6, "=h_goi_3q26/h_q3_oi")
    IN.cell(r, 5, OVR["gaap_k"] if "gaap_k" in OVR else "=(h_goi_4q25+h_goi_9m26)/(h_noi_4q25+h_noi_9m26)")
    IN.cell(r, 7, "Derived"); IN.cell(r, 8, "公司 4Q25、3Q26 新聞稿（SEC 8-K）：GAAP 營業利益 66.958＋330.137 ÷ non-GAAP 營業利益 108.907＋526.845（US$ 百萬）；Hist 頁 h_goi_*、h_noi_*")
    IN.cell(r, 9, "驗收 J2(b)（Andy 2026-11-09）：基準用最近四季實績、全期持平；收斂路徑無證據，只作敏感度。低＝FY2025（51.9%）、高＝3Q26 單季")
    dn(wb, "gaap_k_lo", f"In!$D${r}"); dn(wb, "gaap_k_hi", f"In!$F${r}")


def _bs(wb, IN, OVR):
    # A8：資產負債改 SEC 10-Q
    for k, f, src in (("lease_v", "=h_lsc+h_lsn", "公司 10-Q 3Q26 租賃附註：營業租賃負債流動 6.984＋非流動 38.009（US$ 百萬；Hist 頁）"),
                      ("cash_v", "=h_cash+h_ms", "公司 10-Q 3Q26 資產負債表：現金 209.396＋有價證券 1,168.847（US$ 百萬；Hist 頁）")):
        rr = inrow(wb, k)
        if k not in OVR: IN.cell(rr, 5, f)
        IN.cell(rr, 7, "Derived"); IN.cell(rr, 8, src); IN.cell(rr, 9, "第二來源：stockanalysis（v0.3.5 用值：租賃 57.29、現金與證券 1,378）；v0.3.7 起以 10-Q 為主來源（原則 34）")


def _sh(wb, IN, OVR):
    # 原則 11：評價日股數＝流通股＋RSU（全額）＋選擇權庫藏股法；PSU 與原口徑只作敏感度
    rr = inrow(wb, "sh_v")
    IN.cell(rr, 5, "=q4_sh" if OVR.get("_shorig") else "=h_shout+h_rsu+h_opt*MAX(0,1-h_optx/px)")
    IN.cell(rr, 7, "Derived")
    IN.cell(rr, 8, "公司 10-Q 3Q26：流通股 174.78 百萬（2026-10-31）＋未既得 RSU 8.463 百萬＋選擇權 2.62 百萬 × (1 − 履約價 0.86 ÷ 現價)（庫藏股法；選擇權為 1Q26 值）")
    IN.cell(rr, 9, "個股DCF標準規格第 3 節（Andy 2026-11-09）；PSU 0.326 百萬（績效條件）不進基準；原口徑（公司 4Q26 稀釋股數指引 q4_sh）為 Sens 案例")


def _tt(wb, IN):
    # A4：目標時點＝現價日＋12 個月（評價日 2026-12-31 不變；不需驗收清單 X20）
    rr = inrow(wb, "t_tgt")
    IN.cell(rr, 2, "目標時點距評價日（目標時點＝現價日＋12 個月）"); IN.cell(rr, 5, "=(EDATE(h_asof,12)-h_valdate)/365")
    IN.cell(rr, 8, "目標價為現價日起 12 個月的價格，與 Comparable 的 12 個月後前瞻 EPS 同一時點（驗收 A4；原則 26）"); IN.cell(rr, 9, "v0.3.6 以前＝1 年（評價日＋12 個月，即現價日＋約 14 個月）")


# ---------- 原則 2：單一主要標記與來源誘因 ----------
def incentive(src):
    t = str(src or "")
    if any(k in t for k in ("共識", "賣方", "分析師", "Yahoo", "stockanalysis", "StockAnalysis", "MarketScreener", "marketscreener", "Tickflow", "S&P Global")):
        return "賣方分析師或資料商：評等與交易業務誘因，目標價與預估傾向跟隨股價"
    if any(k in t for k in ("法說", "逐字稿", "MarketBeat", "Futurum", "Motley", "TIKR")):
        return "公司管理層口頭陳述（二手轉述）：股價與市場預期管理誘因，未經查核、可能偏樂觀"
    if any(k in t for k in ("指引", "公司", "新聞稿", "8-K")):
        return "公司指引或自揭數字：預期管理誘因，口徑由公司定義"
    if "opencapital" in t: return INCENT["Interested-party"]
    return "利害關係方提供：未經獨立查核"


def norm_tags(wb):
    n = 0
    for sh, tcol, scol, icol in (("In", 7, 8, 10), ("Q", 9, 10, 12)):
        ws = wb[sh]
        hr = next(r for r in range(1, 15) if ws.cell(r, tcol).value == "標記")
        c = ws.cell(hr, icol, "來源誘因"); c.font = BOLD; c.fill = HDR; ws.column_dimensions[c.column_letter].width = 50
        for r in range(hr + 1, ws.max_row + 1):
            t = ws.cell(r, tcol).value
            if not isinstance(t, str) or t == "標記": continue
            if "／" in t:
                parts = [p.strip() for p in t.split("／")]
                prim = "Interested-party" if ("Interested-party" in parts or "指引" in parts) else parts[0]
                ws.cell(r, tcol).value = prim
                note = f"欄別標記（v0.3.5 混標「{t}」）：實績季為財報（Verified）、4Q26 為公司指引（Interested-party）；主要標記取最近一期原文"
                ws.cell(r, icol, (incentive("指引") + "；" + note) if prim == "Interested-party" else note); n += 1; t = prim
            if sh == "In" and t == "Verified":
                base = ws.cell(r, 5).value
                if isinstance(base, str) and base.startswith("=") and re.search(r"[+\-*/]", base[1:]):
                    ws.cell(r, tcol).value = t = "Derived"; n += 1
            if t not in ("Verified", "Interested-party", "Analogy", "Assumed", "Derived", "—"):
                ws.cell(r, tcol).value = t = "Assumed" if "Assumed" in t else ("Derived" if "Derived" in t else "Interested-party"); n += 1
            if t == "Interested-party" and not ws.cell(r, icol).value:
                ws.cell(r, icol, incentive(f"{ws.cell(r, scol).value or ''}｜{ws.cell(r, 2).value or ''}"))
    return n


# ---------- 原則 1：build_id、輸入指紋、模型狀態 ----------
def build_id(here, stage):
    h = hashlib.sha256(stage.encode())
    for fn in sorted(["build_alab.py", "v036_post.py", "v038_post.py", "px1105_post.py", "sens.json", "segments.json", "coord_alt.json"]):   # v0.3.8：加 v038_post.py、coord_alt.json
        p = os.path.join(here, fn)
        if os.path.exists(p): h.update(fn.encode()); h.update(open(p, "rb").read())
    return f"alab-v0.3.8-{h.hexdigest()[:10]}"   # v0.3.8 起標籤改版號（數值邏輯仍由 ALAB_STAGE＝0.3.7 控制，見 build_alab.py）


def governance(wb, SJ, bid, OVR):
    SE = wb["Sens"]; IN = wb["In"]
    hr = next(r for r in range(1, 15) if IN.cell(r, 7).value == "標記")
    last = next(r for r in range(hr + 1, IN.max_row + 1) if str(IN.cell(r, 1).value or "").startswith("情境驅動對照")) - 1
    rng = f"In!$D${hr + 1}:$F${last}"
    r = SE.max_row + 3; G = {}
    SE.cell(r, 1, "模型治理（v0.3.6 起；原則 1）：Sens 主表與 EPS 案例為活頁簿公式；「業務主因（整本重算）」表與 HTML 是建置快照。只在 Excel 改 In 頁而未重跑 python3 alab/sens_alab.py 與建置時，下方模型狀態顯示 STALE").font = BOLD; r += 1
    SE.cell(r, 1, "輸入指紋（本檔即時重算；In 頁低／基準／高欄，情境驅動對照以上）"); SE.cell(r, 4, f"=SUMPRODUCT({rng},ROW({rng})*3+COLUMN({rng}))").number_format = "0.000000"; G["fp"] = r; r += 1
    base = (SJ or {}).get("_base") if not OVR else None
    SE.cell(r, 1, "輸入指紋（建置快照的基準檔）")
    if base: SE.cell(r, 4, base["fp"]).number_format = "0.000000"
    fb = r; r += 1
    SE.cell(r, 1, "加權目標價（建置快照的基準檔）")
    if base: SE.cell(r, 4, base["tp"]).number_format = "#,##0.00"
    tb = r; r += 1
    tpc = f"Val!$B${lab(wb['Val'], '目標價（DCF 與 Comparable 加權）')}"
    SE.cell(r, 1, "模型狀態（CURRENT＝Sens 快照、HTML、README 與本檔輸入一致；STALE＝In 頁已改、須重跑 python3 alab/sens_alab.py 與建置）").font = BOLD
    SE.cell(r, 4, f'=IF(AND(ISNUMBER(D{fb}),ISNUMBER(D{tb}),ABS(D{G["fp"]}-D{fb})<=0.000000001*MAX(1,ABS(D{fb})),ABS({tpc}-D{tb})<0.005),"CURRENT","STALE")'); G["st"] = r; r += 1
    SE.cell(r, 1, "build_id（README、Sens、HTML 三處相同）"); SE.cell(r, 4, bid); G["bid"] = r
    return G


# ---------- A5：目標價 Sens 加 DCF 單法美元欄；A6：業務主因（整本重算快照）----------
def sens_tp(wb, SJ, V37):
    SE = wb["Sens"]
    r0 = lab(SE, "目標價參數敏感度"); hr = r0 + 1
    SE.cell(hr, 9, "低：DCF 單法變化（美元）").font = BOLD; SE.cell(hr, 10, "高：DCF 單法變化（美元）").font = BOLD
    r = hr + 1
    while SE.cell(r, 1).value and not str(SE.cell(r, 1).value).startswith("檢查"):
        is_w = SE.cell(r, 8).value == "tp_w_dcf"
        for cv, cd in (("C", "I"), ("E", "J")):
            SE[f"{cd}{r}"] = '="—（權重案例）"' if is_w else f'=IF(ISNUMBER({cv}{r}),({cv}{r}-G{r})/tp_w_dcf,"—")'
            SE[f"{cd}{r}"].number_format = "+#,##0.00;-#,##0.00"
        r += 1
    SE.cell(r0, 1).value = str(SE.cell(r0, 1).value) + "（v0.3.6 起 I、J 欄＝DCF 單法目標價變化，美元；Comparable 不變的案例＝加權變化 ÷ DCF 權重）"
    G = {"tp0": r0, "tp_end": r - 1}
    if not V37: return G
    # A6（v0.3.7）：業務主因——整本重算快照
    rows = [("Scorpio X 份額 低／高（三情境同用）", "relx", "三情境都用份額相對值低／高路徑（In 頁 rel_x_*_lo／_hi）"),
            ("PCIe 等份額 低／高（三情境同用）", "relp", "三情境都用 PCIe 等相對值低／高路徑（rel_p_*_lo／_hi）"),
            ("GAAP ÷ non-GAAP 營業利益 低／高", "gaap_k", "低＝FY2025 實績、高＝3Q26 單季（基準＝最近四季；J2）"),
            ("最大客戶轉單 15%／50%（尾端；基準 0）", "cl", "三情境 PCIe 等相對值再乘（1 − 最大終端客戶占比 × 轉單比例）；30% 見下一列"),
            ("最大客戶轉單 30%（尾端）", "cl30", "同上，轉單比例 30%（只有低）"),
            ("評價股數原口徑（公司 4Q26 稀釋股數指引）", "shorig", "原則 11 之前的口徑（只有一案，列在「低」）")]
    r = SE.max_row + 3
    SE.cell(r, 1, "目標價敏感度：業務主因（整本重算快照；v0.3.7 起，驗收 A6；原則 32）").font = BOLD; r += 1
    SE.cell(r, 1, "每案把 In 頁輸入換掉後整本重建、LibreOffice 重算（alab/sens_alab.py → alab/sens.json），讀加權目標價與 DCF 目標價；不是活頁簿公式，In 頁改過須重跑（模型狀態見本頁底部）"); r += 1
    for j, h in enumerate(["案例", "設定", "低：加權目標價", "低：相對基準", "高：加權目標價", "高：相對基準", "基準加權目標價", "鍵", "低：DCF 變化（美元）", "高：DCF 變化（美元）"], 1):
        c = SE.cell(r, j, h); c.font = BOLD; c.fill = HDR
    r += 1; G["biz0"] = r
    b = (SJ or {}).get("_base", {})
    for name, key, setting in rows:
        d = (SJ or {}).get("biz", {}).get(key, {})
        SE.cell(r, 1, name); SE.cell(r, 2, setting); SE.cell(r, 8, key)
        for side, ct, cr, cdl in (("低", 3, 4, 9), ("高", 5, 6, 10)):
            v = d.get(side)
            if v:
                SE.cell(r, ct, v["tp"]).number_format = "#,##0.00"; SE.cell(r, cr, v["tp"] / b["tp"] - 1).number_format = "+0.0%;-0.0%"
                SE.cell(r, cdl, v["dcf"] - b["dcf"]).number_format = "+#,##0.00;-#,##0.00"
            else:
                SE.cell(r, ct, "—")
        if b: SE.cell(r, 7, b["tp"]).number_format = "#,##0.00"
        r += 1
    G["biz1"] = r - 1
    return G


# ---------- Recon：B3 隱含採用率；A9 反解 ----------
def recon(wb, SJ, V37):
    RC = wb["Recon"]; G = {}
    r = lab(RC, "FY2028 Scorpio X 每顆非 NVIDIA 加速器營收")
    RC.cell(r, 1, "FY2028 隱含採用率上限＝Scorpio X 每顆營收（全部非 NVIDIA、非 TPU 加速器平均）÷ 公司採用平台每顆內容「逾 1,000 美元」")
    for cv, cd in ((3, 4), (5, 6), (7, 8)):
        L = chr(64 + cv)
        RC.cell(r, cd, f"={L}{r}/B{r}").number_format = "0.0%"
    RC.cell(r, 10, "驗收 B3（原則 33）：兩者口徑不同，比值是隱含採用率上限（每顆內容若真逾 1,000 美元，採用平台占非 NVIDIA 加速器不超過此比例），不是缺口；C、E、G 欄為每顆營收（美元）")
    G["adopt"] = r
    r = RC.max_row + 2
    RC.cell(r, 1, "E. 反解：支持現價與評等翻轉所需的值（驗收 A9；原則 4）").font = BOLD; r += 1
    V = wb["Val"]
    tp = f"Val!$B${lab(V, '目標價（DCF 與 Comparable 加權）')}"; dcf = f"Val!$B${lab(V, 'DCF 目標價（情境加權）')}"
    feps = f"Val!$B${lab(V, '12 個月後的前瞻 EPS')}"; pmed = f"Val!$B${lab(V, '同業前瞻 P/E 中位數')}"
    up = f"Val!$B${lab(V, '上檔空間：加權目標價')}"
    RC.cell(r, 1, "評等級距邊界：升一級（上檔空間）"); RC.cell(r, 2, f"=IF({up}<-0.5,-0.5,IF({up}<-0.25,-0.25,IF({up}<0.25,0.25,IF({up}<0.5,0.5,NA()))))").number_format = "+0%;-0%"
    RC.cell(r, 10, "級距：大於 +50% 強力買進、+25%～+50% 買進、−25%～+25% 中立、−50%～−25% 放空、小於 −50% 強力放空（Andy 2026-10-25）"); G["e_up"] = r; r += 1
    RC.cell(r, 1, "評等級距邊界：降一級（上檔空間）"); RC.cell(r, 2, f"=IF({up}>0.5,0.5,IF({up}>0.25,0.25,IF({up}>-0.25,-0.25,IF({up}>-0.5,-0.5,NA()))))").number_format = "+0%;-0%"; G["e_dn"] = r; r += 1
    for j, h in enumerate(["項目", "需要的值", "模型現值", "需要的變動"], 1): RC.cell(r, j, h).font = BOLD
    r += 1
    w = "tp_w_dcf"; G["inv"] = {}
    tgt = {"px": "px", "up": f"px*(1+B{G['e_up']})", "dn": f"px*(1+B{G['e_dn']})"}
    nm = {"px": "支持現價", "up": "評等升一級", "dn": "評等降一級"}
    for t in ("up", "dn", "px"):
        for kind, f, fmt, cur in (("pe", f"=({tgt[t]}-{w}*{dcf})/((1-{w})*{feps})", "0.0", f"={pmed}*rm_factor"),
                                  ("eps", f"=({tgt[t]}-{w}*{dcf})/((1-{w})*{pmed}*rm_factor)", "0.00", f"={feps}")):
            labl = (f"{nm[t]}所需的 Comparable 採用倍數（同業中位數 × 相對倍數係數；DCF 不變）" if kind == "pe"
                    else f"{nm[t]}所需的 12 個月後前瞻 EPS（同業中位數 × 係數；DCF 不變）")
            RC.cell(r, 1, labl); RC.cell(r, 2, f).number_format = fmt; RC.cell(r, 3, cur).number_format = fmt
            RC.cell(r, 4, f"=B{r}/C{r}-1").number_format = "+0.0%;-0.0%"
            RC.cell(r, 10, "閉式：加權目標價＝w × DCF＋（1 − w）× 倍數 × EPS；DCF 目標價固定。反解值是目標價等於該點的值")
            G["inv"][f"{kind}_{t}"] = r; r += 1
    if V37 and SJ and SJ.get("xk"):
        xk = SJ["xk"]
        for t in ("up", "dn", "px"):
            v = xk.get(t)
            RC.cell(r, 1, f"{nm[t]}所需的 Scorpio X 份額相對值乘數（三情境同乘；整本重算內插）"); RC.cell(r, 2, v if v is not None else "超出測試範圍").number_format = "0.00"; RC.cell(r, 3, 1).number_format = "0.00"
            if v is not None: RC.cell(r, 4, v - 1).number_format = "+0%;-0%"
            RC.cell(r, 10, "alab/sens_alab.py：乘數 " + "、".join(f"{p['k']:g}" for p in xk["pts"]) + " 各整本重算，線性內插；乘數作用於 FY2027–FY2029 份額相對值（FY2026 錨點不變）")
            G["inv"]["xk_" + t] = r; r += 1
    return G


# ---------- Val：A1 讀法公式、J3 舊法標示、原則 5 期間並列、J1 說明 ----------
def val(wb, ctx):
    V = wb["Val"]; SUMR, FC = ctx["SUMR"], ctx["FC"]
    eA, eB, eC = (f"PL!$F${SUMR[(s, 'eps')]}" for s in "ABC")
    r = lab(V, "讀法：以 25 倍")
    rows = [rr for rr in range(r - 4, r) if isinstance(V.cell(rr, 1).value, (int, float))]
    assert len(rows) == 4, rows
    def pos(rr):
        b = f"B{rr}"
        return (f'"以 "&TEXT(A{rr},"0")&" 倍計，隱含 "&TEXT({b},"0.00")&" 美元，"&IF({b}>={eA},"高於情境 A（"&TEXT({b}/{eA}-1,"+0%;-0%")&"）",'
                f'IF({b}>={eB},"介於情境 A 與 B（對 A "&TEXT({b}/{eA}-1,"+0%;-0%")&"、對 B "&TEXT({b}/{eB}-1,"+0%;-0%")&"）",'
                f'IF({b}>={eC},"介於情境 B 與 C（對 B "&TEXT({b}/{eB}-1,"+0%;-0%")&"、對 C "&TEXT({b}/{eC}-1,"+0%;-0%")&"）","低於情境 C（"&TEXT({b}/{eC}-1,"+0%;-0%")&"）")))')
    V.cell(r, 1).value = ('="讀法（公式組字，v0.3.6 起；驗收 A1）：模型 FY2028 EPS 情境 A／B／C＝"&TEXT(' + eA + ',"0.00")&"／"&TEXT(' + eB + ',"0.00")&"／"&TEXT(' + eC + ',"0.00")&" 美元。"&'
                         + '&"；"&'.join(pos(rr) for rr in rows) + '&"。現價是否「高於三個情境」取決於倍數：同業中位數約 "&TEXT(Val!B' + str(lab(V, "同業前瞻 P/E 中位數")) + ',"0.0")&" 倍附近時接近情境 A。"')
    G = {"read": r}
    # E 欄說明（FY2027 是否三情境相同）
    for rr in range(1, V.max_row + 1):
        v = V.cell(rr, 5).value
        if v == "FY2027 三情境相同（顆數係數自 CY2028 起）":
            V.cell(rr, 5).value = "FY2027 三情境不同（v0.3.7 起情境 A／C 帶 Scorpio X 與 PCIe 等份額高／低路徑；驗收 J1）" if "j1" in ctx["V37P"] else v
        if v == "B、C 的差異自 FY2028 起":
            V.cell(rr, 5).value = "顆數差異自 FY2028 起；份額路徑差異自 FY2027 起（J1）" if "j1" in ctx["V37P"] else v
    # J3(a)：舊法標示
    for rr in range(1, V.max_row + 1):
        a = V.cell(rr, 1).value
        if isinstance(a, str) and (a.startswith("舊法：規則值情境加權") or a.startswith("6. 十年倍數與規則值")):
            V.cell(rr, 1).value = a + "【只有 2025 年一個正觀測值，不可用於決策；驗收 J3】"
    rc = lab(V, "投資建議")
    V.cell(rc, 5).value = str(V.cell(rc, 5).value) + "。本列為估值評等，未評估借券、流通量與時點（驗收 B4；同 Arm J1 選項 a）"
    # 原則 5：期間並列
    r = V.max_row + 2
    V.cell(r, 1, "0.6 期間並列（v0.3.6 起；原則 5）").font = BOLD; r += 1
    V.cell(r, 1, "折現期間：DCF 評價日 → 目標時點（年）"); V.cell(r, 2, "=t_tgt").number_format = "0.00"; V.cell(r, 5, '="評價日 "&TEXT(h_valdate,"yyyy-mm-dd")&"；DCF 每股價值以股權成本推 t_tgt 年"'); G["per1"] = r; r += 1
    V.cell(r, 1, "持有期間：現價日 → 目標時點（年）"); V.cell(r, 2, "=(h_valdate+t_tgt*365-h_asof)/365").number_format = "0.00"
    V.cell(r, 5, '="現價日 "&TEXT(h_asof,"yyyy-mm-dd")&"；上檔空間是這段期間的價格變化，不是年化報酬；Comparable 的前瞻 EPS 對應現價日＋約 12 個月"'); G["per2"] = r
    return G


# ---------- README ----------
def readme(wb, ctx, bid, G):
    RD = wb["README"]; V = wb["Val"]; VER = "v" + ctx["STAGE"]
    tp = f"Val!B{lab(V, '目標價（DCF 與 Comparable 加權）')}"; up = f"Val!B{lab(V, '上檔空間：加權目標價')}"
    rec = f"Val!B{lab(V, '投資建議')}"; dcf = f"Val!B{lab(V, 'DCF 目標價（情境加權）')}"; cmp_ = f"Val!B{lab(V, 'Comparable 目標價（相對倍數調整後）')}"
    pmed = f"Val!B{lab(V, '同業前瞻 P/E 中位數')}"; band = f"Val!B{lab(V, '價格帶（DCF')}"
    RD["A1"].value = str(RD["A1"].value) + f"；{VER}（2026-11-09）：Astera Labs v0.3.5 驗收報告採用項；build_id {bid}"
    for r in range(1, RD.max_row + 1):
        v = RD.cell(r, 1).value
        if not isinstance(v, str): continue
        if v.startswith("模型命題：") and "現價約 64 倍共識 FY2026 EPS、40 倍 FY2027。" in v:   # A3：手寫倍數改公式
            RD.cell(r, 1).value = '=' + lit(v.replace("現價約 64 倍共識 FY2026 EPS、40 倍 FY2027。", "")) + '&"現價約 "&TEXT(px/cons_eps26,"0.0")&" 倍共識 FY2026 EPS、"&TEXT(px/cons_eps27,"0.0")&" 倍 FY2027（共識為 3Q26 財報前）。"'
        if v.startswith(("投資建議：", '="投資建議')):   # B4：條件式＋加註
            RD.cell(r, 1).value = (f'="投資建議："&{rec}&"（依上檔空間的估值評等；未評估借券、流通量與時點；加權目標價上檔空間 "&TEXT({up},"+0.0%;-0.0%")&"）。'
                                   f'承重假設：Comparable 以五家同業前瞻 P/E 中位數 "&TEXT({pmed},"0.0")&" 倍、相對倍數係數 "&TEXT(rm_factor,"0.00")&"（未調整：缺資料）；'
                                   f'DCF 的 EBIT 為 non-GAAP 營業利益 × "&TEXT(gaap_k,"0%")&"（GAAP 換算，股權報酬視為費用）；份額相對值為假設。主要失效條件見本頁末。"')
            G["recrow"] = r
        if v.startswith("=") and "現價隱含的成長高於本模型三個情境" in v:   # A2：條件式
            RD.cell(r, 1).value = v.replace("現價隱含的成長高於本模型三個情境",
                                            f'以同業中位數 "&TEXT({pmed},"0.0")&" 倍計，現價隱含 FY2028 EPS "&TEXT(px/{pmed},"0.00")&" 美元，相對情境 A "&TEXT(px/{pmed}/PL!$F${ctx["SUMR"][("A", "eps")]}-1,"+0%;-0%")&"（倍數越高，隱含 EPS 越低；Val 第 4 節列 20／25／35／45 倍）')
            G["a2"] = r
        if v.startswith("=") and "以情境 A 為基準" in v and "j1" in ctx["V37P"]:
            RD.cell(r, 1).value = v.replace("以情境 A 為基準", "以情境 B 為基準")
    # 版本沿革列
    h = lab(RD, "v0.3.5：")
    RD.insert_rows(h)
    RD.cell(h, 1, f"{VER}（2026-11-09）：Astera Labs v0.3.5 驗收報告採用項（Andy 2026-11-09 23:22 確認採用規劃書，J1–J4 照建議預設）。" +
            ("數值：情境 A／C 帶 Scorpio X 與 PCIe 等份額高／低路徑（J1b）；GAAP ÷ non-GAAP 營業利益改最近四季實績（J2b）；目標時點改現價日＋12 個月；評價股數改流通股＋RSU＋選擇權庫藏股法；資產負債改 10-Q；目標價 Sens 新增業務主因整本重算表。目標價變動拆解見下一列。"
             if ctx["V37"] else
             "展示與治理（數值不變）：build_id、輸入指紋與模型狀態；In、Q 頁單一主要標記與來源誘因；歷史財務移到 Hist 頁；Val 市場隱含 EPS 讀法、PL 標題、README 倍數改公式；命題回覆改條件式；Recon 隱含採用率與反解；資料基準日與逾期；風險登錄、失效條件、可證偽指標、評等邊界揭露；目標價 Sens 加 DCF 單法美元欄；舊法標示不可用於決策。"))
    if ctx["V37"] and ctx.get("ATTR"):
        RD.insert_rows(h + 1); RD.cell(h + 1, 1, ctx["ATTR"])
    # 新增各段（本頁末）
    r = RD.max_row + 2
    def put(t, bold=False):
        nonlocal r
        RD.cell(r, 1, t).font = Font(bold=bold); r += 1
        return r - 1
    G["h"] = put(f"{VER} 驗收採用項（HTML 用；Astera Labs v0.3.5 驗收報告，採用規劃書 2026-11-09）", True)
    G["state"] = put(f'="build_id "&Sens!D{G["gov"]["bid"]}&"；模型狀態 "&Sens!D{G["gov"]["st"]}&"（Sens 業務主因表與 HTML 為建置快照：只在 Excel 改 In 頁而未重跑 python3 alab/sens_alab.py 與建置時顯示 STALE）"')
    # J4：評等邊界
    u = up; e = f"IF({u}<-0.375,-0.5,IF({u}<0,-0.25,IF({u}<0.375,0.25,0.5)))"; d = f"ABS({u}-{e})"
    G["edge"] = put(f'=IF({d}<0.05,"評等貼近級距邊界（驗收 J4）：加權目標價上檔空間 "&TEXT({u},"+0.0%;-0.0%")&" 距 "&TEXT({e},"+0%;-0%")&" 邊界只有 "&TEXT({d}*100,"0.0")&" 個百分點；加權目標價 "&TEXT(px*(1+{e}),"0.00")&" 美元（目前 "&TEXT({tp},"0.00")&"）即翻轉評等。",'
                    f'"評等邊界（驗收 J4）：上檔空間距最近級距邊界 "&TEXT({d}*100,"0.0")&" 個百分點（5 個百分點以上，評等不貼近邊界）。")')
    # as_of
    G["asof_h"] = put(f"資料基準日（as_of）：{AS_OF}（現價日）。市場資料與基準日相差超過 2 個交易日標「逾期」；最近一次財報（3Q26，2026-11-03）前的共識標「財報前，只作歷史對照」（原則 15）。", True)
    RD.cell(r, 1, "項目").font = BOLD; RD.cell(r, 2, "資料日期").font = BOLD; RD.cell(r, 3, "狀態").font = BOLD; G["asof_t"] = r; r += 1
    items = [("股價 px", "2026-11-05", 0), ("同業股價（五家）", "2026-11-05", 0), ("同業前瞻 P/E 原值（stockanalysis 頁面）", "2026-10-23", 0),
             ("無風險利率 rf（美國 10 年期公債）", "2026-10-23", 0), ("市場共識目標價 tp_street", "2026-10-16", 1), ("共識 FY2026／FY2027 EPS（Yahoo）", "2026-10-23", 1),
             ("共識 3Q26、4Q26 營收（Recon 對照）", "2026-10-23", 1), ("資產負債（3Q26；v0.3.7 起以 10-Q 為主來源）" if ctx["V37"] else "資產負債（3Q26；stockanalysis，v0.3.7 改 10-Q）", "2026-09-30", 2)]
    G["asof_rows"] = []
    for labl, d, kind in items:
        RD.cell(r, 1, labl); RD.cell(r, 2, d)
        RD.cell(r, 3, '="財報前，只作歷史對照"' if kind == 1 else ('="最近一季資產負債表（以 stub 期滾到評價日）"' if kind == 2 else
                       f'=IF(ABS(NETWORKDAYS(MIN(DATEVALUE(B{r}),DATEVALUE("{AS_OF}")),MAX(DATEVALUE(B{r}),DATEVALUE("{AS_OF}")))-1)>2,"逾期 "&TEXT(ABS(NETWORKDAYS(MIN(DATEVALUE(B{r}),DATEVALUE("{AS_OF}")),MAX(DATEVALUE(B{r}),DATEVALUE("{AS_OF}")))-1),"0")&" 個交易日","同期")'))
        G["asof_rows"].append(r); r += 1
    G["asof_n"] = put("說明：rf 與同業前瞻 P/E 原值逾期，但與其他 AI 半導體個股模型同一取值日，為維持跨模型一致暫不單獨更新（原則 15；不需驗收清單 X24）；同業 P/E 已以 2026-11-05 收盤換算。共識只作對帳，不進目標價。")
    inv = G["recon"]["inv"]
    G["fail_h"] = put("主要失效條件（可觀察；原則 3）", True)
    E_up, E_dn = f"Recon!B{G['recon']['e_up']}", f"Recon!B{G['recon']['e_dn']}"
    fails = [f'="(1) 倍數：Comparable 採用倍數（同業中位數 × 相對倍數係數；目前 "&TEXT({pmed}*rm_factor,"0.0")&" 倍）升到約 "&TEXT(Recon!B{inv["pe_up"]},"0.0")&" 倍，上檔空間到 "&TEXT({E_up},"+0%;-0%")&"、評等升一級；降到約 "&TEXT(Recon!B{inv["pe_dn"]},"0.0")&" 倍則降一級；支持現價需約 "&TEXT(Recon!B{inv["pe_px"]},"0.0")&" 倍。"',
             f'="(2) 盈餘：12 個月後前瞻 EPS（情境加權；模型 "&TEXT(Recon!C{inv["eps_up"]},"0.00")&" 美元）在同業中位數下需到約 "&TEXT(Recon!B{inv["eps_up"]},"0.00")&" 美元評等才升一級；降到約 "&TEXT(Recon!B{inv["eps_dn"]},"0.00")&" 美元則降一級；支持現價需約 "&TEXT(Recon!B{inv["eps_px"]},"0.00")&" 美元。可觀察：4Q26、1Q27 營收與 non-GAAP EPS 對指引、FY2027 指引。"']
    if "xk_up" in inv:
        _x = lambda k: f'IF(ISNUMBER(Recon!B{inv[k]}),TEXT(Recon!B{inv[k]},"0.00")&" 倍","超出測試範圍")'
        fails.append(f'="(3) Scorpio X 份額：三情境份額相對值同乘約 "&{_x("xk_up")}&"時評等升一級、約 "&{_x("xk_dn")}&"時降一級；支持現價需約 "&{_x("xk_px")}&"。可觀察：Scorpio X 量產客戶數、公司揭露的 Scorpio 占營收比例。"')
        fails.append('="(4) GAAP 轉換：GAAP ÷ non-GAAP 營業利益若維持 3Q26 單季水準（"&TEXT(gaap_k_hi,"0%")&"），加權目標價 "&IF(ISNUMBER(Sens!E' + str(G["sens"]["biz0"] + 2) + '),TEXT(Sens!F' + str(G["sens"]["biz0"] + 2) + ',"+0.0%;-0.0%"),"（待重算）")&"；回到 FY2025（"&TEXT(gaap_k_lo,"0%")&"）則 "&IF(ISNUMBER(Sens!C' + str(G["sens"]["biz0"] + 2) + '),TEXT(Sens!D' + str(G["sens"]["biz0"] + 2) + ',"+0.0%;-0.0%"),"（待重算）")&"。可觀察：每季 10-Q 的 GAAP 與 non-GAAP 營業利益。"')
    G["fail"] = [put(t) for t in fails]
    G["frame"] = put(f'="估值框架並列（驗收 B8）：100% DCF＝"&TEXT({dcf},"0.00")&" 美元、100% Comparable（相對倍數調整後）＝"&TEXT({cmp_},"0.00")&" 美元、加權（DCF "&TEXT(tp_w_dcf,"0%")&"）＝"&TEXT({tp},"0.00")&" 美元；DCF 情境價格帶 "&{band}&" 美元。"')
    G["risk_h"] = put("風險登錄（原則 10、19；是否建模、理由、觀察指標）", True)
    G["risk"] = [put(t) for t in (
        "客戶集中：FY2025 單一終端客戶逾 70%、前三大約 86%（10-K）；1Q26–3Q26 直接客戶 A 40%、B 18%、C 14%、D 11%（10-Q）。建模：Sens 尾端「最大客戶轉單」15%／30%／50%（v0.3.7 起整本重算目標價），不進基準（無發生機率證據）。觀察：10-Q／10-K 客戶占比、主要客戶自研或第二供應商。",
        "NVLink Fusion、Broadcom Tomahawk Ultra 等搶非 NVIDIA scale-up 交換份額：建模於 Scorpio X 份額低路徑（v0.3.7 起為情境 C，並有三情境同用的目標價 Sens）。觀察：AMD、Trainium 機櫃的交換晶片選擇公告。",
        "非 NVIDIA 加速器顆數下修：建模於情境顆數係數（Siliconomics）與 Wattonomics 錨定 Sens。觀察：AMD、Trainium、其他 ASIC 出貨。",
        "股權報酬偏高（GAAP 營業利益約為 non-GAAP 的五至七成）：建模於 gaap_k（v0.3.7 起最近四季實績）與其低／高 Sens。觀察：每季 GAAP 與 non-GAAP 營業利益、未認列股權報酬（3Q26 為 8.03 億美元、2.1 年）。",
        "Amazon 認股權證（非現金營收減項，約 200bp 毛利率）：未建模（公司 non-GAAP 已排除）。觀察：認股權證既得進度與公司揭露。",
        "PSU（0.326 百萬股，績效條件）：不進評價股數（原則 11），稀釋影響約 0.2%。觀察：10-Q PSU 活動表。",
        "流通量、借券與軋空：不進目標價（估值評等不含交易執行；不需驗收清單 X14）。觀察：借券費率、空單比例、單日波動（11-05 一日 −19.6%）。")]
    G["ind_h"] = put("主要假設驅動與可證偽的公開指標（原則 20）", True)
    G["ind"] = [put(t) for t in (
        "Scorpio X FY2026 營收（x_26，Assumed）：公司對 Scorpio 占營收比例或 X／P 拆分的任何揭露；3Q26 已揭露 Scorpio 為最大產品家族。",
        "Scorpio X 份額相對值（rel_x_*，Assumed）：Scorpio X 量產客戶數（公司稱逾 10 家接洽）、AMD／Trainium 機櫃設計採用公告。",
        "PCIe 等份額相對值（rel_p_*，Assumed）：Aries／Taurus 營收年增評論、PCIe 6／7 重定時器競爭者設計案。",
        "毛利率路徑（gm_*，Assumed）：每季毛利率對公司長期目標 70%。",
        "GAAP ÷ non-GAAP 營業利益（gaap_k）：每季 10-Q 兩者比例。",
        "最大客戶轉單（cl_loss，Assumed）：10-Q 客戶占比變化。")]
    return r


def apply(wb, here, ctx):
    V37, OVR = ctx["V37"], ctx["OVR"]
    sj_path = os.path.join(here, "sens.json")
    SJ = json.load(open(sj_path, encoding="utf-8")) if (os.path.exists(sj_path) and not OVR) else None
    if SJ and SJ.get("_stage") != ctx["STAGE"]: SJ = None
    hist_sheet(wb); hist_links(wb, OVR)
    if V37: v37_values(wb, OVR, ctx["V37P"])
    n = norm_tags(wb)
    # A3：PL 標題、rf 說明、Units 註記、β 依據（B5）
    PLs = wb["PL"]
    _tail = ("情境 A／C 自 FY2027 起帶 Scorpio X 與 PCIe 等份額高／低路徑（J1），顆數係數差異自 CY2028 起" if "j1" in ctx["V37P"]
             else "FY2026–FY2027 三情境相同（顆數係數 CY2028 起），差異在 FY2028 起")
    PLs["A1"].value = ('="PL — Astera Labs 損益（US$B；non-GAAP）。FY2025A 實績、FY2026 錨點（"&Q!D3&"、"&Q!E3&"、"&Q!F3&"＋"&Q!G3&"），FY2027 起為模型。'
                       '會計年度＝日曆年。三情境各一個完整推導區塊（本頁下方），頂端為摘要。' + _tail + '"')
    in_set(wb, "rf", 9, "FRED DGS10：10-28～10-30 5.05%／4.98%／4.94%。股價日 2026-11-05；殖利率沿用 2026-10-23（逾期，與其他個股模型同一取值日；見 README 資料基準日表）")
    for k in ("beta", "beta_u"):
        r = inrow(wb, k); c = wb["In"].cell(r, 9)
        c.value = str(c.value or "") + "。依據：docs/plan/自算產業β_20261105.md（US 半導體 13 家；Astera Labs 本身在同業組內，上市不足 60 個月）"
    U = wb["Units"]; U["A9"].value = str(U["A9"].value) + "（全部非 NVIDIA、非 TPU 加速器平均；不是採用平台的每顆內容）"
    bid = build_id(here, ctx["STAGE"])
    G = {}
    G["gov"] = governance(wb, SJ, bid, OVR)
    G["sens"] = sens_tp(wb, SJ, V37)
    G["recon"] = recon(wb, SJ, V37)
    G["val"] = val(wb, ctx)
    if V37 and SJ and SJ.get("attr"):
        ctx["ATTR"] = SJ["attr"]
    readme(wb, ctx, bid, G)
    wb["Sens"]["A1"].value = str(wb["Sens"]["A1"].value) + "（模型狀態與 build_id 見本頁底部）"
    return f"v036_post（{ctx['STAGE']}）：標記正規化 {n} 格；build_id {bid}；快照 {'有' if SJ else '無'}"
