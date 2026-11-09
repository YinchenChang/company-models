# Arm v0.3.4（2026-11-09）後處理：Arm v0.3.3 驗收報告的採用項（採用規劃書 Drive 034-Arm/驗收/20261109_Arm驗收v0.3.3_採用規劃書.md；
# Andy 2026-11-09 18:46 確認，J1–J4 照建議預設）。建置腳本 build_arm.py 在存檔前呼叫 apply(wb, here, ctx)。
#   A1  README 核心問題的倍數改公式組字；A2 租賃說明（In 頁 kd 說明欄，build_arm.py 內改公式）
#   A3  build_id、In 頁輸入指紋、模型狀態 CURRENT／STALE（Sens 頁底部；README A1 與 HTML 同一 build_id）；每列單一主要標記與「來源誘因」欄
#   A4  資料基準日 as_of 與各市場資料日期、逾期標示（README）
#   A5  歷史財務改放 Hist 頁（定義名稱 h_*），In 頁公式只引用儲存格
#   A7  投資建議寫成條件式、加註「估值評等，未評估借券、流通量與時點」（J1 選項 a）；主要失效條件
#   A8  Recon：支持現價（與評等翻轉點）所需的同業倍數與 12 個月後前瞻 EPS（閉式）
#   B1  相對倍數係數＝1 標明「未調整：缺資料」並列為承重輸入（Val 子群中位數在 build_arm.py）
#   B2  風險登錄；B6 可證偽指標；B7 估值框架並列（100% DCF／100% Comparable）
import hashlib, os, re
from openpyxl.styles import Font, PatternFill
from openpyxl.workbook.defined_name import DefinedName

BOLD = Font(bold=True); BLUE = Font(color="0000FF"); HDR = PatternFill("solid", fgColor="DDEBF7")
AS_OF = "2026-11-05"
XB = "SEC XBRL companyfacts（CIK 0001973239；https://data.sec.gov/api/xbrl/companyfacts/CIK0001973239.json；20-F FY2024–FY2026；2026-11-05 讀取）"
Q2BS = "Arm 6-K 2026-11-04（季度財報，季度結束 2026-09-30；https://www.sec.gov/Archives/edgar/data/1973239/000197323926000142/arm-20260930.htm）資產負債表（US$M ÷ 1000）"
LET = "股東信（6-K Exhibit 99.2）non-GAAP EPS"

# ---------- A5：歷史財務（十億美元；EPS 美元）----------
HIST = [  # (鍵前綴, 項目, {欄: 值}, 標記, 來源)
    ("rev", "營收", {"24": 3.233, "25": 4.007, "26": 4.92}, "Verified", XB + "：Revenues"),
    ("da", "折舊攤銷", {"24": 0.162, "25": 0.183, "26": 0.249}, "Verified", XB + "：DepreciationDepletionAndAmortization"),
    ("ppe", "購置不動產廠房設備", {"24": 0.092, "25": 0.219, "26": 0.545}, "Verified", XB + "：PaymentsToAcquirePropertyPlantAndEquipment"),
    ("int", "購置無形資產", {"24": 0.051, "25": 0.02, "26": 0.03}, "Verified", XB + "：PaymentsToAcquireIntangibleAssets"),
    ("sbc", "股票薪酬", {"24": 1.037, "25": 0.82, "26": 1.052}, "Verified", XB + "：ShareBasedCompensation（FY2024 含上市相關一次性股票薪酬）"),
    ("ar", "營運資金＋：應收帳款", {"24": 0.781, "25": 1.107, "26": 1.3}, "Verified", XB + "（年底；以下九列依 v0.3.3 公式順序）"),
    ("ca1", "營運資金＋：合約資產（流動）", {"24": 0.336, "25": 0.642, "26": 0.977}, "Verified", XB),
    ("ca2", "營運資金＋：合約資產（非流動）", {"24": 0.24, "25": 0.346, "26": 0.32}, "Verified", XB),
    ("pre", "營運資金＋：預付與其他流動資產", {"24": 0.157, "25": 0.256, "26": 0.358}, "Verified", XB),
    ("ap", "營運資金−：應付帳款", {"24": 0.026, "25": 0.063, "26": 0.08}, "Verified", XB),
    ("emp", "營運資金−：應付員工", {"24": 0.298, "25": 0.14, "26": 0.154}, "Verified", XB),
    ("cl1", "營運資金−：合約負債（流動）", {"24": 0.198, "25": 0.209, "26": 0.294}, "Verified", XB),
    ("cl2", "營運資金−：合約負債（非流動）", {"24": 0.717, "25": 0.702, "26": 0.752}, "Verified", XB),
    ("ocl", "營運資金−：其他流動負債", {"24": 0.835, "25": 0.426, "26": 0.447}, "Verified", XB),
]
BS = [("cash", "現金及約當現金", 3.251), ("sti", "短期投資", 0.37), ("inv", "權益投資（含公允價值選擇權 0.157）", 0.599),
      ("lsc", "營業租賃負債（流動）", 0.045), ("lsn", "營業租賃負債（非流動）", 0.435)]
EPSQ = [("fy26", "FY2026 non-GAAP EPS", 1.77), ("1q26", "1Q26 non-GAAP EPS", 0.35), ("2q26", "2Q26 non-GAAP EPS", 0.39),
        ("1q27", "1Q27 non-GAAP EPS", 0.45), ("2q27", "2Q27 non-GAAP EPS", 0.52)]
NWC_P = ("ar", "ca1", "ca2", "pre"); NWC_N = ("ap", "emp", "cl1", "cl2", "ocl")


def _dn(wb, name, ref):
    wb.defined_names[name] = DefinedName(name, attr_text=ref)


def hist_sheet(wb):
    H = wb.create_sheet("Hist", index=wb.sheetnames.index("Q") + 1)
    H["A1"] = "Hist — 歷史財務（v0.3.4 起；驗收 A5）：In 頁 DCF 參數（折舊、資本支出、營運資金、股票薪酬占營收）、評價日資產負債與近 12 月 EPS 只引用本頁儲存格（定義名稱 h_*），不在公式內寫常數。金額：十億美元；EPS：美元"
    for j, h in enumerate(["鍵", "項目", "FY2024", "FY2025", "FY2026", "標記", "來源"], 1):
        c = H.cell(3, j, h); c.font = BOLD; c.fill = HDR
    r = 4
    for k, lab, vals, tag, src in HIST:
        H.cell(r, 1, f"h_{k}"); H.cell(r, 2, lab)
        for j, y in enumerate(("24", "25", "26")):
            c = H.cell(r, 3 + j, vals[y]); c.font = BLUE; c.number_format = "#,##0.000"
            _dn(wb, f"h_{k}_{y}", f"Hist!${'CDE'[j]}${r}")
        H.cell(r, 6, tag); H.cell(r, 7, src); r += 1
    H.cell(r, 1, "h_nwc"); H.cell(r, 2, "營運資金（＝四項資產 − 五項負債）")
    for j, y in enumerate(("24", "25", "26")):
        f = "=" + "+".join(f"h_{k}_{y}" for k in NWC_P) + "-" + "-".join(f"h_{k}_{y}" for k in NWC_N)
        H.cell(r, 3 + j, f).number_format = "#,##0.000"; _dn(wb, f"h_nwc_{y}", f"Hist!${'CDE'[j]}${r}")
    H.cell(r, 6, "Derived"); H.cell(r, 7, "上方九列相加減"); r += 2
    H.cell(r, 1, "評價用資產負債（2026-09-30，2Q27 季報）").font = BOLD; r += 1
    for j, h in enumerate(["鍵", "項目", "2026-09-30", "", "", "標記", "來源"], 1):
        c = H.cell(r, j, h); c.font = BOLD; c.fill = HDR
    r += 1
    for k, lab, v in BS:
        H.cell(r, 1, f"h_{k}_q2"); H.cell(r, 2, lab); c = H.cell(r, 3, v); c.font = BLUE; c.number_format = "#,##0.000"
        _dn(wb, f"h_{k}_q2", f"Hist!$C${r}"); H.cell(r, 6, "Verified"); H.cell(r, 7, Q2BS); r += 1
    r += 1
    H.cell(r, 1, "近 12 月 EPS 的季資料（non-GAAP，美元）").font = BOLD; r += 1
    for j, h in enumerate(["鍵", "項目", "EPS", "", "", "標記", "來源"], 1):
        c = H.cell(r, j, h); c.font = BOLD; c.fill = HDR
    r += 1
    for k, lab, v in EPSQ:
        H.cell(r, 1, f"h_eps_{k}"); H.cell(r, 2, lab); c = H.cell(r, 3, v); c.font = BLUE; c.number_format = "0.00"
        _dn(wb, f"h_eps_{k}", f"Hist!$C${r}"); H.cell(r, 6, "Verified"); H.cell(r, 7, LET); r += 1
    H.column_dimensions["A"].width = 12; H.column_dimensions["B"].width = 40
    for c in "CDE": H.column_dimensions[c].width = 12
    H.column_dimensions["F"].width = 12; H.column_dimensions["G"].width = 110
    return H


# ---------- A3：單一主要標記與來源誘因 ----------
def incentive(src):
    t = str(src or "")
    if any(k in t for k in ("共識", "賣方", "分析師", "stockanalysis", "marketscreener", "Hudson", "S&P Global")):
        return "賣方分析師：評等與交易業務誘因，目標價傾向跟隨股價"
    if any(k in t for k in ("法說", "逐字稿", "管理層", "CFO")):
        return "公司管理層（法說口頭陳述）：股價與市場預期管理誘因，未經查核、可能偏樂觀"
    if any(k in t for k in ("指引", "股東信", "公司")):
        return "公司指引或自揭數字：預期管理誘因，口徑由公司定義"
    if "SoftBank" in t or "whoistheownerof" in t:
        return "控股股東與關係人資訊（二手整理）：未經查核"
    return "利害關係方提供：未經獨立查核"


def norm_in(wb):
    ws = wb["In"]
    hr = next(r for r in range(1, 20) if ws.cell(r, 7).value == "標記")
    c = ws.cell(hr, 10, "來源誘因"); c.font = BOLD; c.fill = HDR
    ws.column_dimensions["J"].width = 50
    n = 0
    for r in range(hr + 1, ws.max_row + 1):
        t = ws.cell(r, 7).value
        if not isinstance(t, str) or t == "標記": continue
        src = f"{ws.cell(r, 8).value or ''}｜{ws.cell(r, 2).value or ''}"
        base = ws.cell(r, 5).value
        if t == "Verified" and (re.search(r"\d\s*＋|＋\s*\d", src) or (isinstance(base, str) and base.startswith("=") and re.search(r"[+\-*/]", base[1:]))):
            ws.cell(r, 7).value = t = "Derived"; n += 1
        if t == "Interested-party":
            ws.cell(r, 10, incentive(ws.cell(r, 8).value))
    return n


def norm_q(wb):
    ws = wb["Q"]
    c = ws.cell(3, 12, "來源誘因"); c.font = BOLD; c.fill = HDR
    ws.column_dimensions["L"].width = 70
    n = 0
    for r in range(4, ws.max_row + 1):
        t = ws.cell(r, 9).value
        if not isinstance(t, str) or "／" not in t: continue
        lab = str(ws.cell(r, 1).value or "")
        parts = t.split("／")
        prim = "Interested-party" if (lab.startswith("營業費用") or "公司揭露 non-GAAP EPS" in lab) else parts[0]
        prim = {"指引": "Interested-party"}.get(prim, prim)
        ws.cell(r, 9).value = prim
        note = f"欄別標記（v0.3.3 混標「{t}」）：FY2025A–2Q27 為財報實績；3Q27 為公司指引（Interested-party）或指引推得；4Q27 為假設（Assumed）。主要標記取最近一期原文"
        if prim == "Interested-party": note = incentive("指引") + "；" + note
        ws.cell(r, 12, note); n += 1
    return n


# ---------- A3：build_id、輸入指紋、模型狀態 ----------
def build_id(here, ver):
    h = hashlib.sha256()
    for fn in sorted(["build_arm.py", "v034_post.py", "px1105_post.py", "hist_frozen.json", "segments.json"]):
        p = os.path.join(here, fn)
        if os.path.exists(p): h.update(fn.encode()); h.update(open(p, "rb").read())
    return f"arm-{ver}-{h.hexdigest()[:10]}"


def sens_status(wb, ctx, bid):
    SE = wb["Sens"]; IN = wb["In"]
    hr = next(r for r in range(1, 20) if IN.cell(r, 7).value == "標記")
    last = next(r for r in range(hr + 1, IN.max_row + 1) if str(IN.cell(r, 1).value or "").startswith("情境驅動對照")) - 1
    rng = f"In!$D${hr + 1}:$F${last}"
    r = SE.max_row + 3
    SE.cell(r, 1, "模型治理（v0.3.4，驗收 A3）：Sens 主表與 HTML 是建置快照；只在 Excel 改 In 頁而未重跑 sens_arm.py 與建置時，下方模型狀態顯示 STALE").font = BOLD; r += 1
    SE.cell(r, 1, "輸入指紋（本檔即時重算；In 頁敏感度低／基準／高欄，情境驅動對照以上）")
    SE.cell(r, 4, f"=SUMPRODUCT({rng},ROW({rng})*3+COLUMN({rng}))").number_format = "0.000000"; fp_live = r; r += 1
    SE.cell(r, 1, "輸入指紋（敏感度重算時的基準檔）")
    base = (ctx["SJ"] or {}).get("_base") if not ctx["OVR"] else None
    if base: SE.cell(r, 4, base["fp"]).number_format = "0.000000"
    fp_base = r; r += 1
    SE.cell(r, 1, "加權目標價（敏感度重算時的基準檔）")
    if base: SE.cell(r, 4, base["tp"]).number_format = "#,##0.00"
    tp_base = r; r += 1
    SE.cell(r, 1, "模型狀態（CURRENT＝Sens、HTML、README 與本檔輸入一致；STALE＝In 頁已改、須重跑 python3 arm/sens_arm.py 與建置）").font = BOLD
    SE.cell(r, 4, f'=IF(AND(ISNUMBER(D{fp_base}),ISNUMBER(D{tp_base}),ABS(D{fp_live}-D{fp_base})<=0.000000001*MAX(1,ABS(D{fp_base})),ABS(tp_final-D{tp_base})<0.005),"CURRENT","STALE")')
    st = r; r += 1
    SE.cell(r, 1, "build_id（README、Sens、HTML 三處相同）"); SE.cell(r, 4, bid); bidr = r
    return dict(fp=fp_live, st=st, bid=bidr)


# ---------- README ----------
def readme(wb, ctx, bid, G):
    RD = wb["README"]; T = ctx["T"]; VER = ctx["VER"]; LEG = ctx["LEG"]
    a1 = str(RD["A1"].value)
    RD["A1"].value = a1.replace(f"個股模型 {LEG}", f"個股模型 {VER}", 1).replace("　2026-11-06", "　" + ctx["DATE"], 1) + f"；build_id {bid}；{VER}：Arm v0.3.3 驗收報告採用項（治理、揭露、尾端敏感度、首代晶片毛利率、共識更新）"
    RD["A2"].value = f"版本：{VER}（{ctx['DATE']}）：驗收報告採用項；" + str(RD["A2"].value)[len("版本："):]
    # A1：核心問題倍數改公式組字
    for r in range(1, 12):
        v = RD.cell(r, 1).value
        if isinstance(v, str) and v.startswith("核心問題："):
            RD.cell(r, 1).value = '="核心問題：現價約 "&TEXT(px/cons_feps,"0")&" 倍共識前瞻 EPS（2Q27 財報後），市場是否已把資料中心權利金倍增、CSS 費率上升與自研晶片（客戶需求逾 20 億美元）全部算進去；Qualcomm 判決與手機出貨下修會讓權利金少多少。"'
        if isinstance(v, str) and v.startswith("已知限制："):
            RD.cell(r, 1).value = v.replace("共識仍為 2Q27 前（10-21）。", "共識目標價與前瞻 EPS 已更新為 2Q27 財報後（11-05／11-07），FY2027 共識營收仍為財報前（10-21）。")
        if isinstance(v, str) and v.startswith("相對倍數調整："):
            RD.cell(r, 1).value = v + "v0.3.4 起標明「1.00（未調整：缺資料）」並列為承重輸入（Assumed）；Val 第 0.2 節另並列 IP／EDA 與晶片子群中位數（純展示）。"
    # A7（J1 選項 a）：投資建議列改條件式並加註
    rec = f"Val!B{T['rec']}"; upw = f"Val!B{T['upw']}"
    for r in range(1, RD.max_row + 1):
        v = RD.cell(r, 1).value
        if isinstance(v, str) and v.startswith('="投資建議："'):
            RD.cell(r, 1).value = (f'="投資建議："&{rec}&"（依上檔空間的估值評等；未評估借券、流通量與時點；加權目標價上檔空間 "&TEXT({upw},"+0.0%;−0.0%")&"）。'
                                   f'承重假設：Comparable 以六家同業前瞻 P/E 中位數 "&TEXT(Val!B{T["pmed"]},"0.0")&" 倍、相對倍數係數 "&TEXT(rm_factor,"0.00")&"（未調整：ARM 無前瞻倍數歷史）；'
                                   f'DCF 的 EBIT 扣股票薪酬（營收的 "&TEXT(sbc_rev,"0.0%")&"）。"')
            G["recrow"] = r
    # 版本沿革：在 v0.3.3 列前插入
    h = next(r for r in range(1, RD.max_row + 1) if str(RD.cell(r, 1).value or "").startswith(f"{LEG}："))
    RD.insert_rows(h)
    RD.cell(h, 1, f"{VER}（{ctx['DATE']}）：Arm v0.3.3 驗收報告採用項（Andy 2026-11-09 確認採用規劃書，J1–J4 照建議預設）。治理：build_id、輸入指紋、模型狀態（Sens 頁底部）；In、Q 頁單一主要標記與來源誘因欄；"
                    "歷史財務移到 Hist 頁（公式只引用儲存格）；README 核心問題與租賃說明改公式組字；資料基準日與逾期標示；投資建議改條件式並加註；Comparable 子群中位數並列；Recon 支持現價所需倍數與 EPS；"
                    "風險登錄、可證偽指標、失效條件。數值：首代晶片毛利率 FY2027–FY2028 改 30%／35%／40%（chip_gm1，公司揭露）；Sens 新增 Arm China ±20%、Qualcomm 停付一年／兩年、SoftBank 授權不續約、股票薪酬占營收低／高；共識目標價與前瞻 EPS 更新為財報後。目標價變動拆解見下一節。")
    return RD


def add_rows(wb, ctx, G):
    RD = wb["README"]; T = ctx["T"]
    r = RD.max_row + 2
    def put(t, bold=False):
        nonlocal r
        RD.cell(r, 1, t).font = Font(bold=bold); r += 1
        return r - 1
    G["v34h"] = put("v0.3.4 驗收採用項（HTML 用；Arm v0.3.3 驗收報告，採用規劃書 2026-11-09）", True)
    # 模型狀態
    G["state"] = put(f'="build_id "&Sens!D{G["sens"]["bid"]}&"；模型狀態 "&Sens!D{G["sens"]["st"]}&"（Sens 與 HTML 為建置快照：只在 Excel 改 In 頁而未重跑 python3 arm/sens_arm.py 與建置時顯示 STALE）"')
    # as_of
    G["asof_h"] = put(f"資料基準日（as_of）：{AS_OF}（現價日）。市場資料晚於或早於基準日超過 2 個交易日標「逾期」；最近一次財報（2Q27，2026-11-04）前的共識標「財報前，只作歷史對照」。", True)
    RD.cell(r, 1, "項目").font = BOLD; RD.cell(r, 2, "資料日期").font = BOLD; RD.cell(r, 3, "狀態").font = BOLD; G["asof_t"] = r; r += 1
    items = [("股價 px", "2026-11-05", 0), ("同業股價（六家）", "2026-11-05", 0), ("同業前瞻 EPS（stockanalysis 頁面，最早一家）", "2026-10-26", 0),
             ("無風險利率 rf（美國 10 年期公債）", "2026-10-23", 0), ("市場共識目標價 cons_tp", "2026-11-05", 0), ("共識前瞻 12 月 EPS cons_feps（價格日）", "2026-11-06", 0),
             ("共識 FY2027 EPS c_eps27", "2026-11-05", 0), ("共識 FY2027 營收 c_rev27", "2026-10-21", 1), ("共識 2Q27 營收 c_q2rev（只作 2Q27 實績對照）", "2026-10-21", 1)]
    G["asof_rows"] = []
    for lab, d, pre in items:
        RD.cell(r, 1, lab); RD.cell(r, 2, d)
        RD.cell(r, 3, ('="財報前，只作歷史對照"' if pre else
                       f'=IF(ABS(NETWORKDAYS(MIN(DATEVALUE(B{r}),DATEVALUE("{AS_OF}")),MAX(DATEVALUE(B{r}),DATEVALUE("{AS_OF}")))-1)>2,"逾期 "&TEXT(ABS(NETWORKDAYS(MIN(DATEVALUE(B{r}),DATEVALUE("{AS_OF}")),MAX(DATEVALUE(B{r}),DATEVALUE("{AS_OF}")))-1),"0")&" 個交易日","同期")'))
        G["asof_rows"].append(r); r += 1
    G["asof_n"] = put("說明：rf 與同業前瞻 EPS 逾期，但與其他 AI 半導體個股模型同一取值日，為維持跨模型一致暫不單獨更新（驗收 J4 選項 a：其他公司下次改版時一併處理）；共識只作對帳，不進目標價。")
    # 失效條件（A7）
    G["fail_h"] = put("主要失效條件（可觀察；評等由「強力放空」翻轉為「放空」的門檻＝加權目標價 ≥ 現價 × 0.5）", True)
    G["fail"] = [
        put(f'="(1) 倍數：同業前瞻 P/E 中位數（或 ARM 取得長期倍數資料後的相對倍數調整）使 Comparable 採用倍數升到約 "&TEXT(Recon!B{G["inv"]["pe_half"]},"0.0")&" 倍（目前 "&TEXT(Val!B{T["pmed"]},"0.0")&" 倍 × 係數 "&TEXT(rm_factor,"0.00")&"）；支持現價則需約 "&TEXT(Recon!B{G["inv"]["pe_px"]},"0")&" 倍。"'),
        put(f'="(2) 盈餘：12 個月後前瞻 EPS（情境加權）在同業中位數下需到約 "&TEXT(Recon!B{G["inv"]["eps_half"]},"0.00")&" 美元（模型 "&TEXT(Val!B{T["feps"]},"0.00")&"）；可觀察指標為 3Q27、4Q27 權利金年增是否回到雙位數、AGI CPU 首批出貨與毛利率。"'),
        put(f'="(3) 股票薪酬：股票薪酬占營收若降到 FY2025–FY2026 平均（約 "&TEXT(sbc_rev_lo,"0%")&"）並維持，加權目標價 "&IF(ISNUMBER(Sens!E{G["sbc_lo"]}),TEXT(Sens!E{G["sbc_lo"]},"+0.0%;-0.0%"),"（待重算）")&"；單一參數不足以翻轉評等。"'),
    ]
    # B7 估值框架並列
    G["frame"] = put(f'="估值框架並列（驗收 B7）：100% DCF＝"&TEXT(tp_dcf,"0.00")&" 美元、100% Comparable（相對倍數調整後）＝"&TEXT(tp_comp_adj,"0.00")&" 美元、加權（DCF "&TEXT(tp_w_dcf,"0%")&"）＝"&TEXT(tp_final,"0.00")&" 美元；IP／EDA 子群倍數的 Comparable＝"&TEXT(Val!B{T["ct_ip"]},"0.00")&"、晶片子群＝"&TEXT(Val!B{T["ct_chip"]},"0.00")&" 美元（純展示）。"')
    # B2 風險登錄
    G["risk_h"] = put("風險登錄（驗收 B2；是否建模、理由、觀察指標）", True)
    G["risk"] = [
        put("Arm China（最大單一客戶，FY2026 約占營收 16%）：Sens 尾端 cn_adj ±20%（權利金與授權同幅），不進基準；無發生機率證據。觀察：20-F 客戶占比、中國出口管制與 Arm China 授權續約。"),
        put("Qualcomm 停付權利金（Qualcomm 主張違約可停付五年，法官未裁定）：Sens 尾端 qc_stop 停付一年／兩年（兩年經收斂期延伸近似五年），停付額以 Qualcomm 占 FY2026 營收約 9%（含授權，上限）計。觀察：法官判決與上訴。"),
        put("SoftBank 關係人授權（約 2 億美元／季，期間與續約未找到）：Sens 尾端 sb_lic，FY2029 起不續約。觀察：20-F 關係人交易附註、協議期間。"),
        put("客戶集中（前五大合計約 57%）：未建模（無可量化輸入）。觀察：20-F 客戶占比變化。"),
        put("RPO／ACV 轉換與授權收入認列時點：未建模（授權成長 g_lic 為合併假設）。觀察：ACV 年增（2Q27 +12%）、剩餘履約義務與授權收入的季節性。"),
        put("流通股約一成（SoftBank 約九成）與軋空風險：不進目標價（估值評等不含交易執行）。觀察：可借券量、借券費率、SoftBank 減持或質押。"),
        put("AGI CPU 客戶衝突（自研晶片與 IP 客戶競爭）：未建模。觀察：既有 IP 客戶的授權續約與公開表態。"),
    ]
    # B6 可證偽指標
    G["ind_h"] = put("主要假設驅動與可證偽的公開指標（驗收 B6）", True)
    G["ind"] = [
        put("雲端自研 CPU 顆數年增（g_hc）：hyperscaler 自研 Arm CPU 部署公告與 IDC 伺服器出貨；公司揭露的資料中心權利金年增。"),
        put("資料中心每顆權利金率年增（dcr）：新世代核心數（Graviton、Vera、Cobalt）與 CSS 授權數；公司揭露的 CSS 占權利金比例。"),
        put("NVIDIA 占資料中心權利金比重（w_nv）：NVIDIA GB／VR 機櫃出貨（Siliconomics）與 Arm 資料中心權利金的同期變化。"),
        put("手機每支權利金年增（phr_A／B／C）：公司揭露的 v9 權利金占比、CSS 占權利金比例（目標兩年內逾 50%）與手機權利金年增。"),
        put("授權成長（g_lic）：ACV 與授權營收年增；SoftBank 協議期間。"),
        put("AGI CPU 營收與毛利（chip28／chip29、chip_gm1）：4Q27 首批商業出貨、客戶訂單、公司揭露的晶片毛利率。"),
    ]
    return r


def recon_inv(wb, ctx, G):
    RC = wb["Recon"]; T = ctx["T"]
    r = ctx["RECON_LAST"] + 2
    RC.cell(r, 1, "C. 反解：支持現價與評等翻轉所需的值（閉式；驗收 A8）").font = BOLD; r += 1
    w = "tp_w_dcf"; dcf = "tp_dcf"; feps = f"Val!B{T['feps']}"; pmed = f"Val!B{T['pmed']}"
    rows = [("pe_px", "支持現價所需的 Comparable 採用倍數（同業倍數 × 相對倍數係數；DCF 不變）", f"=(px-{w}*{dcf})/((1-{w})*{feps})", "0.0", f"={pmed}*rm_factor"),
            ("pe_half", "評等翻轉（加權目標價＝現價 × 0.5）所需的 Comparable 採用倍數", f"=(0.5*px-{w}*{dcf})/((1-{w})*{feps})", "0.0", f"={pmed}*rm_factor"),
            ("eps_px", "支持現價所需的 12 個月後前瞻 EPS（同業中位數 × 係數；DCF 不變）", f"=(px-{w}*{dcf})/((1-{w})*{pmed}*rm_factor)", "0.00", f"={feps}"),
            ("eps_half", "評等翻轉所需的 12 個月後前瞻 EPS", f"=(0.5*px-{w}*{dcf})/((1-{w})*{pmed}*rm_factor)", "0.00", f"={feps}")]
    G["inv"] = {}
    for k, lab, f, fmt, cur in rows:
        RC.cell(r, 1, lab); RC.cell(r, 2, f).number_format = fmt; RC.cell(r, 3, cur).number_format = fmt
        RC.cell(r, 4, f"=B{r}/C{r}-1").number_format = "0.0%"
        RC.cell(r, 9, "閉式：加權目標價＝w × DCF＋（1 − w）× 倍數 × EPS；DCF 目標價固定"); RC.cell(r, 10, "C 欄＝模型現值；D 欄＝需要的變動")
        G["inv"][k] = r; r += 1
    RC.cell(G["inv"]["pe_px"] - 1, 3, "模型現值"); RC.cell(G["inv"]["pe_px"] - 1, 4, "需要的變動")
    return r - 1


def apply(wb, here, ctx):
    G = {}
    bid = build_id(here, ctx["VER"])
    hist_sheet(wb)
    n_in = norm_in(wb); n_q = norm_q(wb)
    G["sens"] = sens_status(wb, ctx, bid)
    SE = wb["Sens"]
    G["sbc_lo"] = next((r for r in range(5, SE.max_row + 1) if SE.cell(r, 6).value == "sbc_rev" and SE.cell(r, 3).value == "低"), 4)
    SE["A1"].value = str(SE["A1"].value) + "（建置快照；模型狀態與 build_id 見本頁底部）"
    V = wb["Val"]
    V.cell(ctx["T"]["rec"], 5).value = str(V.cell(ctx["T"]["rec"], 5).value) + "。本列為估值評等，未評估借券、流通量與時點（Andy 2026-11-09 驗收 J1 選項 a）"
    for r in range(1, V.max_row + 1):
        if str(V.cell(r, 1).value or "").startswith("相對倍數調整係數"):
            V.cell(r, 5).value = (str(V.cell(r, 5).value or "") + "；" if V.cell(r, 5).value else "") + "資料不足時＝1.00（未調整：缺資料），列為承重輸入（Assumed；v0.3.4 驗收 B1）"
    rl = recon_inv(wb, ctx, G)
    readme(wb, ctx, bid, G)
    add_rows(wb, ctx, G)
    return dict(BUILD_ID=bid, RECON_LAST=rl, G=G, n_tag_in=n_in, n_tag_q=n_q)
