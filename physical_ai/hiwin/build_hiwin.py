# 上銀 HIWIN（2049.TW）個股模型 v0.1 建置腳本
# 規格：docs/plan/個股模型原則_命題加營收分部.md、docs/plan/個股DCF標準規格.md（實體 AI 模型不讀 Siliconomics／Tokenomics／Wattonomics 協調檔）
# 用法：python3 physical_ai/hiwin/build_hiwin.py
#   1. 建基準檔（公式）→ LibreOffice 重算
#   2. 敏感度與反解：複製基準檔、改 In 頁 E 欄輸入、整本重算（快照寫回 Sens／PRS敏感度／Recon）
#   3. 建成品 → 重算 → 錯誤值掃描 → 核對程式（in_layout、sens_coverage、governance、rating_gate、segment_check）→ 產生 README.md
# 金額單位：新台幣百萬元（NT$M）；EPS、股價：新台幣元；股數：百萬股
import copy, glob, hashlib, json, os, re, shutil, subprocess, sys
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.utils import get_column_letter as CL

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(os.path.dirname(HERE))
DIST = os.path.join(HERE, "dist")
TMP = "/tmp/hiwin_build"
VER = "0.1"
OUT = f"20261116_上銀個股模型_v{VER}.xlsx"

BOLD = Font(bold=True); BLUE = Font(color="0000FF"); RED = Font(color="C00000", bold=True)
HDR = PatternFill("solid", fgColor="DDEBF7"); SUB = PatternFill("solid", fgColor="EDEDED"); KEY = PatternFill("solid", fgColor="FFF2CC")
WRAP = Alignment(wrap_text=True, vertical="top")

# ---------------- 來源 ----------------
FF = "https://financialfilings.com/filings/hiwin-2049/"
S_Q3 = f"上銀 115 年第 3 季合併財務報告（202603_2049_AI1.pdf，2026-11-10 公告，勤業眾信核閱）；經 {FF}interim-quarterly-report/2026/76103351/ 擷取"
S_Q2 = f"上銀 115 年第 2 季合併財務報告（202602_2049_AI1.pdf，2026-08-12）；{FF}interim-quarterly-report/2026/55952776/"
S_Q1 = f"上銀 115 年第 1 季合併財務報告（202601_2049_AI1.pdf）；{FF}interim-quarterly-report/2026/48987334/"
S_FY = f"上銀 114 年度合併財務報告（202504_2049_AI1.pdf，2026-02-26）；{FF}annual-report/2026/33550514/"
S_PX = "臺灣證券交易所 個股日成交資訊 2049（https://www.twse.com.tw/rwd/zh/afterTrading/STOCK_DAY?date=20261101&stockNo=2049&response=json；2026-11-16 讀取）"
S_CT1111 = "中時新聞網 2026-11-11 上銀法說（https://wantrich.chinatimes.com/news/20261111900074-420101）"
S_UA1110 = "優分析 2026-11-10 上銀法說（https://uanalyze.com.tw/articles/3706058865）"
S_UA1110b = "優分析 2026-11-10 上銀法說 Q3（https://uanalyze.com.tw/articles/3895158844）"
S_CNA = "中央社 2026-08-19 台灣機器人展專訪（https://cna.com.tw/news/afe/202608190136.aspx）"
S_CTEE0602 = "工商時報 2026-06-02 COMPUTEX（https://www.ctee.com.tw/news/20260602701142-430503）"
S_UA2511 = "優分析 2025-11-11 上銀法說（https://uanalyze.com.tw/articles/6515838456）"
S_CTEE1111 = "工商時報 2026-11-11（https://www.ctee.com.tw/news/20261111700901-430201）"
S_OCT = "Win 投資 2026-11-06 上銀 10 月營收（https://winvest.tw/News/Detail/97257；原始為公開資訊觀測站月營收申報）"
S_UDN25 = "經濟日報 2026-02-26 上銀 2025 年 EPS 4.31 元（https://udn.com/news/story/7252/9347352）"
S_SA = "stockanalysis.com 統計頁（https://stockanalysis.com/quote/<市場>/<代號>/statistics/；2026-11-16 讀取）"
S_DAMO = "Damodaran「Country Default Spreads and Risk Premiums」2026-01-05 版（https://pages.stern.nyu.edu/~adamodar/New_Home_Page/datafile/ctryprem.html）；與台達電、雙鴻、光寶模型同值"
S_RF = "Trading Economics「Taiwan 10Y Bond Yield」2026-11-04 1.98%（https://tradingeconomics.com/taiwan/government-bond-yield）；與台達電、雙鴻、光寶模型同值（跨模型共用取值）"
S_RFUS = "美國財政部每日殖利率曲線 10 Yr：2026-10-23 5.00%（https://home.treasury.gov/）；與 NVIDIA、台達電模型同值"
S_FX = "Yahoo Finance USD/TWD 2026-11-05 33.034（台達電模型 build_delta.py 同值；跨模型共用取值）"
S_2B = "線 2b 人形機器人框架 v0.1（2026-11-16）第 3.2 節"
S_MS = "Morgan Stanley（經 xenospectrum、grabarobot 二手轉述；https://grabarobot.com/blog/humanoid-robot-bill-of-materials-cost-2026/）"
S_MITU = "覓途諮詢（經 鉅亨號 https://hao.cnyes.com/post/252957 轉述）"
S_BOFA = "BofA Institute（經 https://xenospectrum.com/en/humanoid-robot-bom-actuator-cost/ 轉述）"
S_TF = "TrendForce（經 大公資信報告 https://www.jtcopper.com/wp-content/uploads/2026/05/人形机器人：量产突破期的产业分化与国有资本投资逻辑.pdf 轉述）"
S_JARA = "線 2a 工業機器人框架 v0.1 第 5 節（JARA 2026 年 7–9 月受注 +46.6%、生產 +26.2%）"

# ---------------- 歷史（Hist 頁；新台幣百萬元） ----------------
HP = ["FY2024", "FY2025", "1Q25", "1H25", "9M25", "1Q26", "1H26", "9M26"]
N = None
HIST = [  # key, 項目, 值（依 HP）, 標記, 來源, 說明
    ("rev", "合併營收", [24391.684, 24262.638, 5838.780, 11764.271, 17759.602, 6379.288, 13887.696, 21997.456], "Verified", f"FY：{S_FY}；1Q：{S_Q1}；1H：{S_Q2}；9M：{S_Q3}", "各期報告比較欄"),
    ("LG", "線性滑軌 營收", [15496.851, 14951.829, 3624.350, 7266.195, 11038.062, 3820.856, 8363.599, 13318.836], "Verified", "同上，部門資訊附註", ""),
    ("BS", "滾珠螺桿 營收", [4842.602, 4886.269, 1114.116, 2300.179, 3553.556, 1259.785, 2737.231, 4365.338], "Verified", "同上，部門資訊附註", ""),
    ("RB", "機器人 營收", [N, N, 510.881, 1014.467, 1452.139, 747.419, 1613.272, 2584.587], "Verified", "同上；2026 年各季報告起單列（含 2025 比較期重編）", "FY2024、FY2025 年報部門只有 線性滑軌、滾珠螺桿、其他（機器人在其他內）：機器人單列未取得"),
    ("OT", "其他 營收", [4052.231, 4424.540, 589.433, 1183.430, 1715.845, 551.228, 1173.594, 1728.695], "Verified", "同上", "FY2024、FY2025 為舊口徑（含機器人）"),
    ("pLG", "線性滑軌 分部利益", [1663.735, N, 360.840, 715.585, 1097.053, 400.253, 1055.915, 1964.493], "Verified", "同上；FY2024 值由年報調節表依序對應（擷取標籤亂碼）", "FY2025 分部利益：未取得"),
    ("pBS", "滾珠螺桿 分部利益", [458.150, N, 88.074, 258.690, 345.068, 189.333, 471.252, 829.480], "Verified", "同上", ""),
    ("pRB", "機器人 分部利益", [N, N, 50.933, 107.940, 115.253, 120.417, 294.549, 532.017], "Verified", "同上", ""),
    ("pOT", "其他 分部利益", [-63.100, N, -81.096, -164.895, -229.508, -79.482, -157.294, -215.954], "Verified", "同上", "FY2024 舊口徑（含機器人）"),
    ("op", "營業利益（＝分部利益合計）", [2058.785, 1643.229, 418.751, 917.320, 1327.866, 630.521, 1664.422, 3110.036], "Verified", "同上（FY2024 為部門調節表分部合計）", "分部利益合計＝營業利益（公司未分攤項目為 0）"),
    ("gp", "營業毛利", [N, 6977.848, 1742.137, N, 5218.982, 2037.674, 4602.946, 7550.391], "Verified", "同上", ""),
    ("pt", "稅前淨利", [2249.631, 1764.093, 599.992, N, 1241.576, 719.225, N, 3378.161], "Verified", "同上", ""),
    ("tax", "所得稅費用", [N, 353.965, N, N, 259.415, N, N, 811.361], "Verified", "同上", ""),
    ("ni", "本期淨利（合併）", [1862.814, 1410.128, N, N, N, N, N, 2566.800], "Verified", "同上", ""),
    ("nip", "歸屬母公司淨利", [1971.948, 1525.866, 482.799, N, 1063.886, 579.338, 1426.930, 2657.131], "Verified", "同上", "非控制權益為虧損（歸屬母公司 ＞ 合併淨利）"),
    ("eps", "基本 EPS（元）", [5.57, 4.31, 1.36, N, 3.01, 1.64, 4.03, 7.51], "Verified", "同上", ""),
    ("da", "折舊＋攤銷", [2334.714, 2357.929, N, N, 1711.005, N, N, 1902.299], "Verified", "同上（現金流量表）", "9M25 只有折舊"),
    ("cx", "取得不動產、廠房及設備", [2881.118, 1691.702, N, N, 1117.068, N, N, 773.101], "Verified", "同上（現金流量表）", "另有預付設備款增加：9M25 829.372、9M26 767.018"),
    ("cxp", "預付設備款增加", [N, N, N, N, 829.372, N, N, 767.018], "Verified", "同上（現金流量表）", ""),
]
BS_ITEMS = [  # key, 項目, 2025-12-31, 2026-09-30, 來源
    ("cash", "現金及約當現金", 7624.371, 10133.460),
    ("fvtpl", "透過損益按公允價值衡量之金融資產－流動", N, 3.447),
    ("stb", "短期借款", 2033.445, 1972.471),
    ("ltc", "一年內到期長期借款", 977.336, 828.498),
    ("ltb", "長期借款", 6712.084, 6444.670),
    ("lease", "租賃負債（流動＋非流動）", 518.907, 483.364),
    ("nr", "應收票據（含關係人）", 530.986, 498.685),
    ("ar", "應收帳款（含關係人）", 3887.293, 4340.386),
    ("inv", "存貨", 7093.796, 7593.008),
    ("np", "應付票據", 4.576, 5.947),
    ("ap", "應付帳款（含關係人）", 3246.016, 4123.196),
    ("cl", "合約負債－流動", 54.528, 134.147),
    ("eqp", "歸屬母公司權益", 37345.766, 40608.388),
    ("nci", "非控制權益", 127.854, 37.480),
    ("hmc", "持有大銀微系統股票（8%，公允價值）", N, 2248.060),
    ("eqm", "採用權益法之投資", 368.670, N),
]


def defname(wb, name, ref):
    assert not re.fullmatch(r"[A-Za-z]{1,3}\d+", name), name
    wb.defined_names[name] = DefinedName(name, attr_text=ref)


def hdr(ws, r, labels, c0=1):
    for j, h in enumerate(labels, c0):
        if h in ("", None): continue
        c = ws.cell(r, j, h); c.font = BOLD; c.fill = HDR; c.alignment = WRAP


YRS = list(range(2025, 2033))
YC = {y: CL(3 + i) for i, y in enumerate(YRS)}  # C=2025 … J=2032
FY = YRS[2:]  # 2027–2032 明確期


def build(path, sens=None):
    """sens：None（基準檔）或 dict（Sens 快照、2D、反解、指紋）。回傳 cell 對照。"""
    wb = openpyxl.Workbook(); wb.remove(wb.active)
    cells = {}
    # ================= In =================
    IN = wb.create_sheet("In")
    IN["A1"] = ("In — 輸入（藍字＝輸入；黑字＝公式）。基準欄 E 進主模型；D／F「敏感度低／高」只用於 Sens 頁（一次動一個參數），不進目標價。"
                "情境 A（樂觀）／B（基準）／C（保守）的專屬輸入另列（鍵尾 _A／_B／_C），最後一節「情境驅動對照」並排展示。新台幣百萬元，股價與 EPS 為元，股數百萬股")
    IN["A2"] = "標記：Verified（附出處）、Interested-party（寫明誰與誘因）、Analogy（寫明類比對象）、Assumed、Derived；Analogy／Assumed 以 D／F 給區間。查不到寫「未取得」"
    hdr(IN, 3, ["鍵", "項目", "單位", "敏感度低", "基準", "敏感度高", "標記", "來源", "來源誘因", "說明"])
    rows = []
    def sec(t): rows.append(("sec", t))
    def inp(key, item, unit, lo, base, hi, tag, src, inc="", note=""):
        rows.append((key, item, unit, lo, base, hi, tag, src, inc, note))

    sec("M0. 市場資料與日期")
    inp("px", "現價（收盤）", "元", None, 289.0, None, "Verified", S_PX, "", "2026-11-16 收盤 289.00（11-13 為 314.00）")
    inp("px_date", "股價基準日（as_of）", "日期", None, "=DATE(2026,11,16)", None, "Verified", S_PX, "", "資料截止日＝股價基準日")
    inp("tp_date", "目標時點（＝EDATE(px_date,12)）", "日期", None, "=EDATE(px_date,12)", None, "Derived", "原則第 10 節第 26 條", "", "")
    inp("val_yrs", "折現期間：評價日 2026-12-31 → 目標時點（年）", "年", None, "=(tp_date-DATE(2026,12,31))/365", None, "Derived", "評價日 2026-12-31（同 repo 慣例）", "", "上檔空間不是年化報酬")
    inp("hold_yrs", "持有期間：股價基準日 → 目標時點（年）", "年", None, "=(tp_date-px_date)/365", None, "Derived", "原則第 10 節第 5 條", "", "")
    inp("fx", "匯率 USD/TWD（人形命題美元 → 新台幣；預測期固定）", "NT$/US$", None, 33.034, None, "Verified", S_FX, "", "基準假設匯率固定")
    inp("sh_v", "評價與 EPS 用稀釋股數", "百萬股", None, "=353.792+0.772", None, "Derived", S_Q3 + "：已發行 353,792 千股＋員工酬勞可能發股之稀釋 772 千股（9M26 稀釋 EPS 計算）", "",
        "無員工認股權或 RSU 計畫（9M26 報告未揭露有效計畫；資本公積有失效認股權 10,634 千元）；庫藏股：未取得（以 0 計）")

    sec("M1. FY2026 錨點（9M26 實績＋4Q26 模型）")
    inp("q4g", "4Q26 營收相對 3Q26 季增（各分部同比例）", "%", 0.0, 0.04, 0.08, "Assumed", S_OCT + "：10 月營收 2,850（月增 4.79%、年增 35.66%）× 3 ≈ 8,550，相對 3Q26 8,110 約 +5.4%",
        "", "公司指引 4Q 優於 3Q（Interested-party，見 Recon）；區間 0–8%")
    inp("capex_26", "FY2026 資本支出（含預付設備款）", "NT$M", None, 2700, None, "Interested-party", S_CNA + "：董事長稱今年資本支出約與 2025 年的 27 億元相當", "董事長（公司經營者，對外說法可能偏樂觀；資本支出屬計畫值）",
        "9M26 實付 773.1＋預付設備款 767.0＝1,540.1；4Q26 隱含 1,159.9")

    sec("M2. 分部核心成長（工業循環；基準路徑；命題人形子列另列）")
    for seg, nm, v27, v28, lo27, hi27, lo28, hi28 in (("LG", "線性滑軌", 0.12, 0.06, 0.05, 0.18, 0.0, 0.10), ("BS", "滾珠螺桿", 0.12, 0.06, 0.05, 0.18, 0.0, 0.10),
                                                    ("RB", "機器人", 0.18, 0.12, 0.08, 0.28, 0.04, 0.18), ("OT", "其他", 0.0, 0.0, -0.05, 0.05, -0.05, 0.05)):
        inp(f"g27_{seg}", f"FY2027 {nm} 核心營收年增（基準路徑，各情境共用）", "%", lo27, v27, hi27, "Assumed",
            S_CT1111 + "：公司稱 2027 年營運成長「可能維持雙位數」、機器人比重續升；" + S_UA1110 + "：滾珠螺桿與滑軌接單能見度到 2027 年中；" + S_JARA, "",
            "雙位數為公司說法（Interested-party），本列取區間中段；機器人 9M26 年增 78%（Hist），基期效應後放緩")
        inp(f"g28_{seg}", f"FY2028 {nm} 核心營收年增（基準路徑，各情境共用）", "%", lo28, v28, hi28, "Assumed",
            S_UA1110 + "：日本工具機工業會預期全球工具機投資 2027 年中起回升（Interested-party：產業公會）", "", "FY2029–FY2032 由 FY2028 線性收斂到長期成長率")
    inp("g_lt", "長期核心成長率（線性滑軌、滾珠螺桿；FY2032）", "%", 0.02, 0.03, 0.05, "Assumed", "工業自動化長期成長；FY2015–FY2025 上銀營收 CAGR：未取得", "", "")
    inp("g_lt_RB", "長期核心成長率（機器人；FY2032）", "%", 0.03, 0.06, 0.09, "Assumed", "IFR 預測 2026 年 65.5 萬台 → 2029 年 80.6 萬台（約 7%／年；線 2a 第 1 節）", "", "類比全球工業機器人裝機成長")
    inp("g_lt_OT", "長期核心成長率（其他；FY2032）", "%", None, 0.0, None, "Assumed", "其他分部（軸承、醫療、控制器等）FY2024→9M26 持平", "", "")

    sec("M3. 分部利益率（新分部口徑最近可得 9M26；比率參數以最近實績為起點，原則第 35 條）")
    for seg, nm in (("LG", "線性滑軌"), ("BS", "滾珠螺桿"), ("RB", "機器人"), ("OT", "其他")):
        inp(f"m_{seg}", f"{nm} 分部利益率（9M26 實績）", "%", None, f"=Hist!J{{p{seg}}}/Hist!J{{{seg}}}", None, "Derived", "9M26 分部利益 ÷ 分部營收（Hist 頁）", "",
            "最近四季口徑（4Q25 新分部）未揭露，以最近三季代替；全期持平，情境差另列")
    inp("mg_d", "分部利益率加減（全部分部；Sens 用，基準 0）", "pt", -0.02, 0.0, 0.02, "Assumed", "Sens：±2pt", "", "")
    inp("nonop_r", "業外淨額 ÷ 營收", "%", 0.005, "=(Hist!J{nonop9m26}+Hist!D{nonopq4}-Hist!C{fxq3})/(Hist!J{rev}+Hist!D{revq4})", 0.0167, "Derived",
        "最近四季（4Q25＋9M26）業外淨額扣除 3Q26 人民幣升值匯兌利益約 120 ÷ 營收", "", "高＝不扣匯兌（1.67%）；低 0.5%（Assumed）")
    inp("tax_r", "有效稅率（最近四季）", "%", 0.20, "=(Hist!J{tax}+Hist!D{taxq4})/(Hist!J{pt}+Hist!D{ptq4})", 0.24, "Derived", "4Q25＋9M26 所得稅 ÷ 稅前（Hist 頁）", "", "低＝FY2025 20.1%；高＝9M26 24.0%")
    inp("nci_r", "非控制權益損益 ÷ 營收（負＝子公司虧損由少數股東分擔）", "%", None, "=(Hist!J{ni}-Hist!J{nip})/Hist!J{rev}", None, "Derived", "9M26（合併淨利 − 歸屬母公司）÷ 營收", "", "")
    inp("payout", "現金股利發放率", "%", None, "=2/Hist!D{eps}", None, "Derived", S_UDN25 + "；FY2025 每股現金股利 2.0 元 ÷ EPS 4.31", "", "")

    sec("M4. 命題模組：人形機器人線性關節（行星滾柱絲杠、微型滑軌）— 讀線 2b")
    inp("hu_26", "全球人形機器人出貨 2026（台）", "台", None, 50000, None, "Interested-party", S_TF + "：2026 全球 ＞5 萬台", "研究機構預測（對產業有推廣立場）", "2026 年上銀人形營收以 0 計（ramp 0）")
    inp("bom_27", "每台硬體 BOM 2027（美元）", "US$", None, 35000, None, "Interested-party", S_BOFA + "：2025 年底業界 BOM 約 3.5 萬美元", "賣方研究（有承銷與交易利益）", "2026–2027 視為同一水準")
    inp("bom_32", "每台硬體 BOM 2032（美元）", "US$", 13000, 17000, 20000, "Assumed", S_BOFA + "：2030–35 年 1.3–1.7 萬美元（區間端點取自該情境）", "", "2027→2032 幾何下降")
    inp("prs_sh", "行星滾柱絲杠價值 ÷ 每台 BOM（本模型最大不確定）", "%", 0.19, 0.25, 0.52, "Assumed",
        S_2B + "：" + S_MITU + " 絲杠占 19%；" + S_MS + " 每台 12–14 支 × 1,350–2,700 美元＝BOM 的 35–69%（中點 52%）", "",
        "兩來源無法同時成立；基準取 25%（偏向較低端，原則第 10 節第 17 條：不利揭露可信度高）；完整區間見 PRS敏感度 頁")
    inp("n_scr", "每台行星滾柱絲杠支數（換算每支單價用）", "支", None, 14, None, "Interested-party", S_MS + "：12–14 支", "賣方研究（有承銷與交易利益）", "只用於 Recon 換算，不進模型")
    inp("ex_cn", "非中國供應鏈整機占出貨比例（上銀可及市場）", "%", 0.30, 0.40, 0.50, "Assumed", S_2B + "：中美供應鏈分流；2025 年中國廠占全球出貨多數（Omdia、Bloomberg）", "", "中國整機以陸系絲杠為主（五洲新春、貝斯特、北特等）")
    inp("hw_prs", "上銀在可及市場的行星滾柱絲杠份額（穩態，2030 起）", "%", 0.02, 0.05, 0.10, "Assumed",
        S_CNA + "：上銀行星滾柱絲杠「可進入量產階段」；" + S_UA2511 + "：西方廠已有產品、價格可達滾珠螺桿 10 倍以上；線 2b：舍弗勒、新劍傳動為兩大供應商", "", "公司未揭露客戶與份額：未取得")
    inp("gd_sh", "微型滑軌與小螺桿（手臂、靈巧手）價值 ÷ 每台 BOM", "%", 0.01, 0.02, 0.04, "Assumed", "未取得公開拆分；線 2b 第 3.2 節靈巧手合計約 9,500 美元（MS）", "", "")
    inp("hw_gd", "上銀在可及市場的微型滑軌與小螺桿份額（穩態）", "%", 0.05, 0.10, 0.20, "Assumed", S_CT1111 + "：獲美國人形機器人大廠小型滑軌訂單、批量五位數、2027 年初出貨", "", "訂單金額未揭露：未取得")
    inp("ramp_27", "上銀份額爬坡 2027（相對穩態）", "x", None, 0.25, None, "Assumed", "2027 年初首批出貨（同上）", "", "2026＝0")
    inp("ramp_28", "上銀份額爬坡 2028（相對穩態）", "x", None, 0.5, None, "Assumed", "—", "", "")
    inp("ramp_29", "上銀份額爬坡 2029（相對穩態）", "x", None, 0.75, None, "Assumed", "—", "", "2030 起＝1")

    sec("M5. 折現率與無槓桿自由現金流（標準版 DCF；評價日 2026-12-31；資產負債表日 2026-09-30）")
    inp("rf", "無風險利率（台灣 10 年期公債殖利率）", "%", None, 0.0198, None, "Verified", S_RF, "", "2026-11-04 值，距股價日逾 2 個交易日：逾期（跨模型共用，下次全面改版更新）")
    inp("rf_us", "美元口徑折現用：美國 10 年期公債殖利率（只用於 Sens）", "%", None, 0.05, None, "Verified", S_RFUS, "", "Sens「美元口徑折現」把 rf 換成本值")
    inp("blume_a", "Blume 調整係數", "x", None, 0.67, None, "Derived", "Blume（1971）；同 repo 其他模型", "", "")
    inp("beta_1", "原始 β：上銀 5 年（資料商）", "x", None, 1.08, None, "Verified", S_SA.replace("<市場>/<代號>", "tpe/2049") + "（資料日 2026-10-22）", "", "資料商計算，未調整")
    inp("beta_2", "原始 β：線性傳動同業中位數（Comps 頁 7 家）", "x", None, "=MEDIAN(Comps!J5:J11)", None, "Derived", "Comps 頁 β 欄中位數", "", "")
    inp("beta", "β（兩來源 Blume 調整後中位數）", "x", "=MEDIAN(blume_a*beta_1+1-blume_a,blume_a*beta_2+1-blume_a)-0.3", "=MEDIAN(blume_a*beta_1+1-blume_a,blume_a*beta_2+1-blume_a)", "=MEDIAN(blume_a*beta_1+1-blume_a,blume_a*beta_2+1-blume_a)+0.3", "Derived", "＝MEDIAN(各來源 blume_a × 原始 β ＋ 1 − blume_a)", "", "Sens：±0.3（Assumed）")
    inp("erp", "股權風險溢酬 ERP", "%", 0.04, 0.05, 0.06, "Assumed", "DCF 標準規格 v1.0：5%（敏感度 4–6%）", "", "")
    inp("crp", "國家風險溢酬 CRP（台灣）", "%", None, 0.0078, None, "Verified", S_DAMO + "：Taiwan（Moody's Aa3）0.78%", "", "")
    inp("kd", "稅前債務成本 kd", "%", None, "=Hist!J{intpaid}/0.75/((Hist!C{bdebt}+Hist!D{bdebt})/2)", None, "Derived", S_Q3 + "：9M26 支付利息 169,985 千元年化 ÷ 平均銀行借款", "", "")
    inp("tax_w", "WACC 用稅率", "%", None, "=tax_r", None, "Derived", "＝tax_r", "", "")
    inp("wacc_adj", "WACC 調整（Sens 用；基準 0）", "%", -0.01, 0.0, 0.01, "Assumed", "Sens：±1pt", "", "")
    inp("g_cap", "永續成長率上限", "%", None, 0.03, None, "Assumed", "DCF 標準規格：3%，且不超過計價幣別 10 年期公債殖利率", "", "")
    inp("g_tv", "永續成長率 g（＝MIN(g_cap, rf)）", "%", 0.0098, "=MIN(g_cap,rf)", None, "Derived", "DCF 標準規格", "", "Sens 低：−1pt；高＝上限 rf（與基準同，不另列）")
    inp("ronic_sp", "RONIC 相對 WACC 的差", "%", None, 0.0, 0.05, "Assumed", "DCF 標準規格：RONIC＝WACC", "", "Sens：WACC＋5pt")
    inp("t_tgt", "目標時點距評價日（年）", "年", None, 1.0, None, "Derived", "評價日 2026-12-31 → 2027-12-31", "", "")
    inp("cash_v", "現金及約當現金（2026-09-30）", "NT$M", None, "=Hist!D{bcash}+Hist!D{bfvtpl}", None, "Derived", S_Q3 + "：現金 10,133.460＋FVTPL 3.447", "", "")
    inp("debt_v", "銀行借款本金（2026-09-30）", "NT$M", None, "=Hist!D{bdebt}", None, "Derived", S_Q3 + "：短借＋一年內到期長借＋長借", "", "")
    inp("lease_v", "租賃負債（2026-09-30）", "NT$M", None, "=Hist!D{blease}", None, "Verified", S_Q3, "", "")
    inp("inv_v", "非營業投資（大銀微 8% 公允價值＋權益法投資）", "NT$M", None, "=Hist!D{bhmc}+Hist!C{beqm}", None, "Derived", S_Q3 + "：大銀微 2,248.060；" + S_FY + "：權益法投資 368.670（2025-12-31；2026-09-30 未取得）", "", "")
    inp("inv_k", "非營業投資計入比例", "x", 0.0, 1.0, None, "Assumed", "帳面（大銀微為公允價值）全額計入", "", "Sens 低：不計")
    inp("mi_v", "非控制權益（帳面）", "NT$M", None, "=Hist!D{bnci}", None, "Verified", S_Q3, "", "金額小，以帳面")
    inp("fc_da", "折舊攤銷 ÷ 營收（FY2027 起點）", "%", None, "=Hist!J{da}/Hist!J{rev}", None, "Derived", "9M26 折舊＋攤銷 1,902.3 ÷ 營收 21,997.5", "", "")
    inp("fc_da_ss", "折舊攤銷 ÷ 營收（FY2032 穩態）", "%", 0.06, 0.07, 0.08, "Assumed", "折舊金額大致持平、營收成長 → 比率下降", "", "FY2027→FY2032 線性收斂")
    inp("fc_cx", "資本支出 ÷ 營收（FY2027 起點）", "%", None, "=capex_26/Hist!D{rev26}", None, "Derived", "FY2026 資本支出（capex_26）÷ FY2026 營收錨點", "", "")
    inp("fc_cx_ss", "資本支出 ÷ 營收（FY2032 穩態＝穩態折舊 ×（1＋5g））", "%", None, "=fc_da_ss*(1+5*g_tv)", None, "Assumed", "設備年限約 10 年、以 g 成長時資本支出 ÷ 折舊約 1＋g × 年限 ÷ 2（同台達電模型）", "", "")
    inp("fc_cx_d", "資本支出 ÷ 營收額外加減（Sens；基準 0）", "pt", -0.01, 0.0, 0.01, "Assumed", "Sens：±1pt", "", "")
    inp("fc_wc", "營運資金 ÷ 營收（營運資金增加＝本值 × 營收增量）", "%", 0.25, "=Hist!D{wc}/(Hist!J{rev}+Hist!D{revq4})", 0.34, "Derived",
        "2026-09-30 營運資金（應收票據與帳款＋存貨 − 應付票據與帳款 − 合約負債）÷ 最近四季營收", "", "高＝2025-12-31 比率 33.8%；低 25%（Assumed）")
    inp("dps_tgt", "評價日到目標時點的每股股利（FY2026 盈餘，2027 年發放）", "元", None, "=payout*PL!D§EPSB§", None, "Derived", "發放率 × FY2026 EPS（情境 B）", "", "")

    sec("M6. Comparable（同業前瞻 P/E；分部加總）")
    inp("pe_lin", "線性傳動同業前瞻 P/E 中位數（線性滑軌、滾珠螺桿、其他）", "x", 15.36, "=MEDIAN(Comps!F5:F10)", 32.04, "Derived", "Comps 頁 6 家（五洲新春無前瞻 P/E）", "", "Sens 低／高＝NSK／恒立（約第 1、5 名）")
    inp("pe_rob", "機器人同業前瞻 P/E 中位數", "x", 21.49, "=MEDIAN(Comps!F12:F14)", 84.44, "Derived", "Comps 頁 Fanuc、安川、哈默納科", "", "Sens 低／高＝安川／哈默納科")
    inp("rel_k", "相對倍數調整係數", "x", None, 1.0, None, "Assumed", "1.00（未調整：缺上銀對同業 10 年前瞻 P/E 比值歷史）", "", "承重輸入；現行比值見 Val")
    inp("dcf_w", "DCF 權重（其餘為 Comparable）", "%", 0.3, 0.5, 0.7, "Assumed", "原則第 4 節預設 DCF 50%＋分部加總 50%", "", "")

    sec("M7. 情境專屬輸入（自 FY2027 起帶路徑；原則第 10 節第 46 條）")
    for s, v, note in (("A", 0.06, "情境 A（樂觀）：工業循環延續、產能擴充全數消化"), ("B", 0.0, "情境 B（基準）：無專屬調整"), ("C", -0.10, "情境 C（保守）：2027 年循環反轉（類比 2023–2024 下行）")):
        inp(f"d27_{s}", f"情境 {s}：FY2027 核心成長加減（各分部）", "pt", None, v, None, "Assumed", note, "", "")
    for s, v in (("A", 0.04), ("B", 0.0), ("C", -0.08)):
        inp(f"d28_{s}", f"情境 {s}：FY2028 核心成長加減（各分部）", "pt", None, v, None, "Assumed", "同上", "", "")
    inp("mg_A", "情境 A：分部利益率加減（＝3Q26 營業利益率 − 9M26）", "pt", None, "=Hist!J{opq3}/Hist!J{revq3}-Hist!J{op}/Hist!J{rev}", None, "Derived", "稼動率高檔延續（3Q26 單季水準）", "", "")
    inp("mg_B", "情境 B：分部利益率加減", "pt", None, 0.0, None, "Assumed", "9M26 水準持平", "", "")
    inp("mg_C", "情境 C：分部利益率加減（＝FY2025 營業利益率 − 9M26）", "pt", None, "=Hist!D{op}/Hist!D{rev}-Hist!J{op}/Hist!J{rev}", None, "Derived", "稼動率回到 FY2025 水準", "", "")
    for s, v in (("A", 1000000), ("B", 300000), ("C", 100000)):
        inp(f"hu30_{s}", f"情境 {s}：全球人形機器人出貨 2030（台）", "台", None, v, None, "Assumed", "線 2b 無出貨預測；Tesla Fremont 設計產能 100 萬台／年、宇樹 7.5 萬台、Hyundai 2028 年 3 萬台（線 2b，Interested-party）", "", "")
    for s, v in (("A", 3000000), ("B", 1000000), ("C", 300000)):
        inp(f"hu32_{s}", f"情境 {s}：全球人形機器人出貨 2032（台）", "台", None, v, None, "Assumed", "同上", "", "")
    for s, v in (("A", 0.25), ("B", 0.5), ("C", 0.25)):
        inp(f"w_{s}", f"情境權重 {s}", "%", None, v, None, "Assumed", "DCF 標準規格：預設 25／50／25", "", "")

    # 先建 Hist 以取得列號（In 的公式用 .format）
    H = wb.create_sheet("Hist")
    H["A1"] = "Hist — 歷史財務（新台幣百萬元；EPS 元）。藍字＝原文；黑字＝由原文相減（Derived）。部門口徑：2026 年起機器人單列（2025 比較期已重編），FY2024–FY2025 年報為舊口徑（機器人在其他）"
    hdr(H, 3, ["鍵", "項目"] + HP + ["標記", "來源", "來源誘因", "說明"])
    hr = {}
    r = 4
    for key, item, vals, tag, src, note in HIST:
        H.cell(r, 1, key); H.cell(r, 2, item)
        for j, v in enumerate(vals):
            if v is not None:
                c = H.cell(r, 3 + j, v); c.font = BLUE; c.number_format = "#,##0.00" if key == "eps" else "#,##0.0"
        H.cell(r, 11, tag); H.cell(r, 12, src); H.cell(r, 14, note); hr[key] = r; r += 1
    r += 1
    H.cell(r, 1, "檢查：分部合計 − 合併營收（各期）").font = BOLD
    for j in range(8):
        c = CL(3 + j)
        H[f"{c}{r}"] = f"={c}{hr['LG']}+{c}{hr['BS']}+N({c}{hr['RB']})+{c}{hr['OT']}-{c}{hr['rev']}"; H[f"{c}{r}"].number_format = "0.000"
    hr["chk"] = r; r += 1
    H.cell(r, 1, "檢查：分部利益合計 − 營業利益（有揭露的期別）").font = BOLD
    for j in (0, 2, 3, 4, 5, 6, 7):
        c = CL(3 + j)
        H[f"{c}{r}"] = f"={c}{hr['pLG']}+{c}{hr['pBS']}+N({c}{hr['pRB']})+{c}{hr['pOT']}-{c}{hr['op']}"; H[f"{c}{r}"].number_format = "0.000"
    r += 2
    # 衍生單季
    H.cell(r, 1, "衍生單季（Derived：期間相減）").font = BOLD; r += 1
    hdr(H, r, ["鍵", "項目", "4Q25", "2Q26", "3Q26", "", "", "", "", "", "標記", "來源", "來源誘因", "說明"]); r += 1
    q = {}
    for key, item in (("rev", "合併營收"), ("LG", "線性滑軌 營收"), ("BS", "滾珠螺桿 營收"), ("RB", "機器人 營收"), ("OT", "其他 營收"),
                      ("pLG", "線性滑軌 分部利益"), ("pBS", "滾珠螺桿 分部利益"), ("pRB", "機器人 分部利益"), ("pOT", "其他 分部利益"), ("op", "營業利益"),
                      ("pt", "稅前淨利"), ("tax", "所得稅費用")):
        H.cell(r, 1, key + "_q"); H.cell(r, 2, item)
        a = hr[key]
        if key in ("rev", "LG", "BS", "op", "pt", "tax"):
            H[f"C{r}"] = f"=D{a}-G{a}"  # 4Q25＝FY25 − 9M25
        if key == "OT":
            H[f"C{r}"] = f"=D{a}-G{a}-G{hr['RB']}"; H.cell(r, 14, "4Q25 舊口徑：其他＋機器人（FY2025 其他含機器人 − 9M25 其他 − 9M25 機器人）")
        H[f"D{r}"] = f"=I{a}-H{a}"; H[f"E{r}"] = f"=J{a}-I{a}"
        if key in ("pt", "tax"):
            H[f"D{r}"] = None; H[f"E{r}"] = None
        for c in "CDE": H[f"{c}{r}"].number_format = "#,##0.0"
        H.cell(r, 11, "Derived"); H.cell(r, 12, "上列期間值相減")
        q[key] = r; r += 1
    H.cell(r, 1, "nonop_q"); H.cell(r, 2, "業外淨額"); H[f"C{r}"] = f"=(D{hr['pt']}-D{hr['op']})-(G{hr['pt']}-G{hr['op']})"; H.cell(r, 11, "Derived"); H.cell(r, 12, "稅前 − 營業利益")
    H[f"J{r}"] = f"=J{hr['pt']}-J{hr['op']}"; H.cell(r, 14, "C＝4Q25；J＝9M26"); q["nonop"] = r; r += 1
    r += 1
    # 資產負債
    H.cell(r, 1, "資產負債表科目").font = BOLD; r += 1
    hdr(H, r, ["鍵", "項目", "2025-12-31", "2026-09-30", "", "", "", "", "", "", "標記", "來源", "來源誘因", "說明"]); r += 1
    b = {}
    for key, item, v1, v2 in BS_ITEMS:
        H.cell(r, 1, "b_" + key); H.cell(r, 2, item)
        for c, v in (("C", v1), ("D", v2)):
            if v is not None: H[f"{c}{r}"] = v; H[f"{c}{r}"].font = BLUE; H[f"{c}{r}"].number_format = "#,##0.0"
        H.cell(r, 11, "Verified"); H.cell(r, 12, S_Q3 + "（2025-12-31 為比較欄）"); b[key] = r; r += 1
    H.cell(r, 1, "b_debt"); H.cell(r, 2, "銀行借款合計"); H[f"C{r}"] = f"=C{b['stb']}+C{b['ltc']}+C{b['ltb']}"; H[f"D{r}"] = f"=D{b['stb']}+D{b['ltc']}+D{b['ltb']}"
    H.cell(r, 11, "Derived"); H.cell(r, 12, "加總"); b["debt"] = r; r += 1
    H.cell(r, 1, "b_wc"); H.cell(r, 2, "營運資金（應收票據與帳款＋存貨 − 應付票據與帳款 − 合約負債）")
    for c in "CD":
        H[f"{c}{r}"] = f"={c}{b['nr']}+{c}{b['ar']}+{c}{b['inv']}-{c}{b['np']}-{c}{b['ap']}-{c}{b['cl']}"; H[f"{c}{r}"].number_format = "#,##0.0"
    H.cell(r, 11, "Derived"); H.cell(r, 12, "加減"); b["wc"] = r; r += 1
    r += 1
    H.cell(r, 1, "其他揭露").font = BOLD; r += 1
    hdr(H, r, ["鍵", "項目", "值", "", "", "", "", "", "", "", "標記", "來源", "來源誘因", "說明"]); r += 1
    o = {}
    for key, item, v, tag, src, inc in (
            ("intpaid", "9M26 支付利息", 169.985, "Verified", S_Q3, ""),
            ("divpaid", "9M26 發放現金股利（FY2025 每股 2.0 元）", 707.584, "Verified", S_Q3, ""),
            ("oct", "2026 年 10 月營收", 2850.0, "Verified", S_OCT, ""),
            ("ytd10", "2026 年 1–10 月累計營收", 24848.0, "Verified", S_OCT, ""),
            ("fx_q3", "3Q26 人民幣升值匯兌利益（約）", 120.0, "Interested-party", S_UA1110b + "：業外 1.34 億元中約 1.2 億元為人民幣升值匯兌利益", "公司法說說法（經媒體轉述）"),
            ("rbsh", "機器人營收占比 2026 全年指引（＞）", 0.12, "Interested-party", S_CT1111 + "：全年可望超過 12%，2027 年續升", "公司法說指引")):
        H.cell(r, 1, key); H.cell(r, 2, item); H[f"C{r}"] = v; H[f"C{r}"].font = BLUE
        H.cell(r, 11, tag); H.cell(r, 12, src); H.cell(r, 13, inc); o[key] = r; r += 1
    r += 1
    # FY2026 錨點
    H.cell(r, 1, "FY2026 錨點（9M26 實績＋4Q26 模型：4Q 各分部＝3Q26 × (1＋q4g)，分部利益率＝3Q26）").font = BOLD; r += 1
    hdr(H, r, ["鍵", "項目", "9M26", "4Q26E", "FY2026E", "4Q26E 分部利益率", "", "", "", "", "標記", "來源", "來源誘因", "說明"]); r += 1
    a26 = {}
    for key, item in (("LG", "線性滑軌"), ("BS", "滾珠螺桿"), ("RB", "機器人"), ("OT", "其他")):
        H.cell(r, 1, "a26_" + key); H.cell(r, 2, item + " 營收")
        H[f"C{r}"] = f"=J{hr[key]}"; H[f"D{r}"] = f"=E{q[key]}*(1+q4g)"; H[f"E{r}"] = f"=C{r}+D{r}"
        H[f"F{r}"] = f"=E{q['p' + key]}/E{q[key]}"
        for c in "CDE": H[f"{c}{r}"].number_format = "#,##0.0"
        H[f"F{r}"].number_format = "0.0%"
        H.cell(r, 11, "Derived"); H.cell(r, 12, "9M26 實績＋4Q26 模型"); a26[key] = r; r += 1
        H.cell(r, 1, "a26_p" + key); H.cell(r, 2, item + " 分部利益")
        H[f"C{r}"] = f"=J{hr['p' + key]}"; H[f"D{r}"] = f"=D{r - 1}*F{r - 1}"; H[f"E{r}"] = f"=C{r}+D{r}"
        for c in "CDE": H[f"{c}{r}"].number_format = "#,##0.0"
        H.cell(r, 11, "Derived"); H.cell(r, 12, "9M26 實績＋4Q26 模型"); a26["p" + key] = r; r += 1
    H.cell(r, 1, "a26_rev"); H.cell(r, 2, "合併營收")
    H[f"C{r}"] = f"=J{hr['rev']}"; H[f"D{r}"] = "=" + "+".join(f"D{a26[k]}" for k in ("LG", "BS", "RB", "OT")); H[f"E{r}"] = f"=C{r}+D{r}"
    H.cell(r, 11, "Derived"); a26["rev"] = r; r += 1
    H.cell(r, 1, "a26_op"); H.cell(r, 2, "營業利益")
    H[f"C{r}"] = f"=J{hr['op']}"; H[f"D{r}"] = "=" + "+".join(f"D{a26['p' + k]}" for k in ("LG", "BS", "RB", "OT")); H[f"E{r}"] = f"=C{r}+D{r}"
    H.cell(r, 11, "Derived"); a26["op"] = r; r += 1
    H.cell(r, 1, "檢查：FY2026E 分部合計 − 合併營收"); H[f"E{r}"] = "=" + "+".join(f"E{a26[k]}" for k in ("LG", "BS", "RB", "OT")) + f"-E{a26['rev']}"; r += 1
    H.cell(r, 1, "檢查：1–10 月營收（9M26＋10 月）− 月營收累計"); H[f"E{r}"] = f"=J{hr['rev']}+C{o['oct']}-C{o['ytd10']}"
    H.cell(r, 14, "差額為季報合併數與月營收申報口徑差（月營收為自結）"); r += 1
    for col, w in zip("ABCDEFGHIJKLMN", [12, 40, 12, 12, 12, 12, 12, 12, 12, 12, 10, 70, 24, 50]): H.column_dimensions[col].width = w
    H.freeze_panes = "C4"

    fmtmap = dict(pLG=hr["pLG"], LG=hr["LG"], pBS=hr["pBS"], BS=hr["BS"], pRB=hr["pRB"], RB=hr["RB"], pOT=hr["pOT"], OT=hr["OT"],
                  nonop9m26=q["nonop"], nonopq4=q["nonop"], rev=hr["rev"], revq4=q["rev"], tax=hr["tax"], taxq4=q["tax"], pt=hr["pt"], ptq4=q["pt"],
                  ni=hr["ni"], nip=hr["nip"], eps=hr["eps"], intpaid=o["intpaid"], bdebt=b["debt"], bcash=b["cash"], bfvtpl=b["fvtpl"], blease=b["lease"],
                  bhmc=b["hmc"], beqm=b["eqm"], fxq3=o["fx_q3"], bnci=b["nci"], da=hr["da"], rev26=a26["rev"], wc=b["wc"], op=hr["op"], opq3=q["op"], revq3=q["rev"])
    # nonop9m26 在 J 欄、nonopq4 在 C 欄（同一列）：In 公式寫法見下
    # ---- 寫 In ----
    r = 4; inrow = {}
    for row in rows:
        if row[0] == "sec":
            IN.cell(r, 1, row[1]).font = BOLD; r += 1; continue
        key, item, unit, lo, base, hi, tag, src, inc, note = row
        if sens and sens.get("override", {}).get(key) is not None:
            base = sens["override"][key]
        IN.cell(r, 1, key); IN.cell(r, 2, item); IN.cell(r, 3, unit)
        for j, v in ((4, lo), (5, base), (6, hi)):
            if v is None: continue
            if isinstance(v, str) and "{" in v:
                v = v.format(**fmtmap)
                v = v.replace(f"Hist!D{q['nonop']}", f"Hist!C{q['nonop']}").replace(f"Hist!D{q['rev']}", f"Hist!C{q['rev']}").replace(f"Hist!D{q['tax']}", f"Hist!C{q['tax']}").replace(f"Hist!D{q['pt']}", f"Hist!C{q['pt']}")
                v = v.replace(f"Hist!J{q['op']}", f"Hist!E{q['op']}").replace(f"Hist!J{q['rev']}", f"Hist!E{q['rev']}")
                v = v.replace(f"Hist!D{a26['rev']}", f"Hist!E{a26['rev']}").replace(f"Hist!J{o['intpaid']}", f"Hist!C{o['intpaid']}")
            c = IN.cell(r, j, v)
            if not (isinstance(v, str) and v.startswith("=")): c.font = BLUE
            c.number_format = "0.0%" if unit in ("%", "pt") else ("yyyy-mm-dd" if unit == "日期" else "#,##0.00")
        IN.cell(r, 7, tag); IN.cell(r, 8, src); IN.cell(r, 9, inc if tag == "Interested-party" else (inc or None)); IN.cell(r, 10, note or None)
        defname(wb, key, f"In!$E${r}"); inrow[key] = r; r += 1
    last_in = r
    # 情境驅動對照
    r += 1
    IN.cell(r, 1, "情境驅動對照（純展示；修改請改上方藍字）").font = BOLD; sec_r = r; r += 1
    hdr(IN, r, ["鍵", "項目", "單位", "情境 A（樂觀）", "情境 B（基準）", "情境 C（保守）", "", "說明"]); r += 1
    for k, item, unit, note in (("d27", "FY2027 核心成長加減", "pt", "工業循環"), ("d28", "FY2028 核心成長加減", "pt", "工業循環"), ("mg", "分部利益率加減", "pt", "稼動率"),
                                ("hu30", "全球人形出貨 2030", "台", "命題"), ("hu32", "全球人形出貨 2032", "台", "命題")):
        IN.cell(r, 1, k); IN.cell(r, 2, item); IN.cell(r, 3, unit)
        for j, s in zip((4, 5, 6), "ABC"): IN.cell(r, j, f"=E{inrow[f'{k}_{s}']}")
        IN.cell(r, 8, f"引用 In 頁 {k}_A／{k}_B／{k}_C；{note}"); r += 1
    IN.cell(r, 1, "scn_w"); IN.cell(r, 2, "情境權重"); IN.cell(r, 3, "%")
    for j, s in zip((4, 5, 6), "ABC"): IN.cell(r, j, f"=E{inrow[f'w_{s}']}")
    IN.cell(r, 8, "引用 In 頁 w_A／w_B／w_C；合計＝1"); r += 1
    IN.cell(r + 1, 1, "註：其他輸入（稅率、業外、折舊、資本支出、營運資金、人形 BOM 與份額）情境不分化：無公司層共變機制證據（原則第 10 節第 46 條）；其不確定以 Sens 呈現")
    for col, w in zip("ABCDEFGHIJ", [12, 52, 9, 11, 12, 11, 15, 80, 28, 50]): IN.column_dimensions[col].width = w
    IN.freeze_panes = "D4"
    cells["in_last"] = last_in; cells["inrow"] = inrow

    # ================= Comps =================
    CP = wb.create_sheet("Comps")
    CP["A1"] = "Comps — 同業倍數（資料商 stockanalysis.com，2026-11-16 讀取；各檔資料日期不同，見 D 欄；與股價基準日相差逾 2 個交易日者標「逾期」）"
    CP["A2"] = "會計口徑：未查核（資料商以各公司申報口徑統一計算；日股 IFRS／日本基準、陸股中國會計準則、台股 IFRS 混合）。前瞻 P/E 分母＝資料商彙整之分析師下一年度 EPS 共識（NTM 口徑由資料商定義）"
    hdr(CP, 4, ["公司", "代號", "群組", "資料日期", "逾期", "前瞻 P/E", "EV/EBITDA", "EBITDA 率", "淨現金（＋）", "β（5Y）", "標記", "來源"])
    PEERS = [("THK", "6481.T", "線性傳動", "2026-10-19", 15.37, 13.17, 0.2609, "淨負債", 0.76),
             ("NSK", "6471.T", "線性傳動", "2026-09-18", 15.35, 3.17, 0.1074, "淨現金", 0.53),
             ("Nippon Thompson（IKO）", "6480.T", "線性傳動", "2026-09-04", 11.81, 12.96, 0.138, "淨現金", 1.04),
             ("直得", "1597.TW", "線性傳動", "2026-08-24", 58.97, 44.37, 0.1908, "淨現金", 0.99),
             ("貝斯特", "300580.SZ", "線性傳動", "2026-09-29", 26.78, 16.93, 0.3132, "淨現金", 0.60),
             ("恒立液壓", "601100.SS", "線性傳動", "2026-10-26", 32.04, 29.83, 0.2931, "淨現金", 0.87),
             ("五洲新春", "603667.SS", "線性傳動", "2026-10-13", None, 60.98, 0.0907, "淨現金", 0.86),
             ("Fanuc", "6954.T", "機器人", "2026-11-13", 23.22, 14.45, 0.2844, "淨現金", 0.98),
             ("安川電機", "6506.T", "機器人", "2026-09-16", 21.49, 17.57, 0.1214, "淨負債", 1.24),
             ("哈默納科", "6324.T", "機器人", "2026-10-02", 84.44, 53.83, 0.183, "淨現金", 0.96)]
    rr = 5
    for nm, tk, gp, d, pe, ev, em, nc, bt in PEERS:
        if rr == 12 and gp == "機器人": pass
        CP.cell(rr, 1, nm); CP.cell(rr, 2, tk); CP.cell(rr, 3, gp); CP.cell(rr, 4, d)
        CP.cell(rr, 5, f'=IF(px_date-DATEVALUE(D{rr})>4,"逾期","")')
        for j, v in ((6, pe), (7, ev), (8, em), (10, bt)):
            if v is not None: CP.cell(rr, j, v).font = BLUE
        if pe is None: CP.cell(rr, 6, "未取得")
        CP.cell(rr, 9, nc); CP.cell(rr, 11, "Verified"); CP.cell(rr, 12, S_SA)
        rr += 1
    # 列號：線性傳動 5–11（前瞻 P/E 5–10 有值，11 五洲新春無），機器人 12–14
    CP.cell(rr + 1, 1, "上銀（參照）"); CP.cell(rr + 1, 2, "2049.TW"); CP.cell(rr + 1, 4, "2026-10-22"); CP.cell(rr + 1, 6, 28.35).font = BLUE; CP.cell(rr + 1, 7, 21.77).font = BLUE
    CP.cell(rr + 1, 8, 0.1782).font = BLUE; CP.cell(rr + 1, 10, 1.08).font = BLUE; CP.cell(rr + 1, 11, "Verified"); CP.cell(rr + 1, 12, S_SA + "；分析師 17 家平均目標價 393.65（Interested-party，見 Recon）")
    CP.cell(rr + 1, 5, f'=IF(px_date-DATEVALUE(D{rr + 1})>4,"逾期","")')
    cells["hw_pe"] = f"Comps!F{rr + 1}"
    CP.cell(rr + 3, 1, "線性傳動中位數（前瞻 P/E）"); CP.cell(rr + 3, 6, "=MEDIAN(F5:F10)")
    CP.cell(rr + 4, 1, "機器人中位數（前瞻 P/E）"); CP.cell(rr + 4, 6, "=MEDIAN(F12:F14)")
    CP.cell(rr + 5, 1, "註：每組少於 8 家，不做迴歸（原則第 10 節第 7 條）；淨負債 ÷ EBITDA、資本強度：未取得（以淨現金正負代理）。哈默納科前瞻 P/E 84 倍反映人形機器人溢價，使機器人中位數對其敏感")
    for col, w in zip("ABCDEFGHIJKL", [22, 11, 10, 12, 6, 10, 11, 10, 10, 8, 10, 70]): CP.column_dimensions[col].width = w

    # ================= Thesis =================
    TH = wb.create_sheet("Thesis")
    TH["A1"] = "Thesis — 命題模組：人形機器人線性關節價值落到上銀多少（每台 BOM × 行星滾柱絲杠占比 × 非中國供應鏈 × 上銀份額）；結果進 PL 頁 線性滑軌、滾珠螺桿 的「人形」子列"
    TH["A2"] = "式：上銀 PRS 營收＝出貨 × BOM × prs_sh × ex_cn × hw_prs × 爬坡；微型滑軌＝出貨 × BOM × gd_sh × ex_cn × hw_gd × 爬坡。人形子列分部利益率＝母分部利益率（Assumed；PRS 實際利益率未取得）"
    PL = wb.create_sheet("PL")
    PL["A1"] = "PL — 分部營收 → 損益 → 現金流（新台幣百萬元）。FY2025＝實績（舊分部口徑：機器人在其他）；FY2026＝9M26 實績＋4Q26 模型；FY2027–FY2032＝模型。情境 B 在最上方（主列），其下為情境 A、C"
    TR, PR = {}, {}
    tr0 = 4; pr0 = 3
    for s in "BAC":
        # ---- Thesis 區塊 ----
        t = {}
        r = tr0
        TH.cell(r, 1, f"情境 {s}（{'樂觀' if s == 'A' else '基準' if s == 'B' else '保守'}）").font = BOLD
        for j in range(1, 11): TH.cell(r, j).fill = SUB
        r += 1
        hdr(TH, r, ["項目", "單位"] + [f"FY{y}" for y in YRS]); r += 1
        def trow(k, lab, unit, f, fm="#,##0"):
            nonlocal r
            TH.cell(r, 1, lab); TH.cell(r, 2, unit)
            for y in YRS[1:]:
                v = f(y)
                if v is not None: TH[f"{YC[y]}{r}"] = v; TH[f"{YC[y]}{r}"].number_format = fm
            t[k] = r; r += 1
        def units(y):
            if y == 2026: return "=hu_26"
            if y <= 2030: return f"=hu_26*(hu30_{s}/hu_26)^(({y}-2026)/4)"
            return f"=hu30_{s}*(hu32_{s}/hu30_{s})^(({y}-2030)/2)"
        trow("u", "全球人形機器人出貨", "台", units)
        trow("bom", "每台硬體 BOM", "US$", lambda y: "=bom_27" if y <= 2027 else f"=bom_27*(bom_32/bom_27)^(({y}-2027)/5)")
        trow("prs", "每台行星滾柱絲杠價值（＝BOM × prs_sh）", "US$", lambda y: f"={YC[y]}{t['bom']}*prs_sh")
        trow("ramp", "上銀份額爬坡（相對穩態）", "x", lambda y: 0 if y == 2026 else "=ramp_27" if y == 2027 else "=ramp_28" if y == 2028 else "=ramp_29" if y == 2029 else 1, "0.00")
        trow("prs_m", "上銀 PRS 營收（百萬美元）", "US$M", lambda y: f"={YC[y]}{t['u']}*{YC[y]}{t['prs']}*ex_cn*hw_prs*{YC[y]}{t['ramp']}/1e6", "#,##0.0")
        trow("gd_m", "上銀 微型滑軌與小螺桿 營收（百萬美元）", "US$M", lambda y: f"={YC[y]}{t['u']}*{YC[y]}{t['bom']}*gd_sh*ex_cn*hw_gd*{YC[y]}{t['ramp']}/1e6", "#,##0.0")
        trow("prs_t", "上銀 PRS 營收（新台幣百萬元）→ 滾珠螺桿 人形子列", "NT$M", lambda y: f"={YC[y]}{t['prs_m']}*fx", "#,##0.0")
        trow("gd_t", "上銀 微型滑軌與小螺桿 營收（新台幣百萬元）→ 線性滑軌 人形子列", "NT$M", lambda y: f"={YC[y]}{t['gd_m']}*fx", "#,##0.0")
        trow("mkt", "全球行星滾柱絲杠市場（人形；百萬美元）", "US$M", lambda y: f"={YC[y]}{t['u']}*{YC[y]}{t['prs']}/1e6", "#,##0")
        TR[s] = t; tr0 = r + 1
        # ---- PL 區塊 ----
        p = {}
        r = pr0
        PL.cell(r, 1, f"情境 {s}（{'樂觀' if s == 'A' else '基準' if s == 'B' else '保守'}）").font = BOLD
        for j in range(1, 11): PL.cell(r, j).fill = SUB
        r += 1
        hdr(PL, r, ["項目", "說明"] + [f"FY{y}{'A' if y == 2025 else 'E'}" for y in YRS]); r += 1
        def prow(k, lab, f, fm="#,##0.0", note=""):
            nonlocal r
            PL.cell(r, 1, lab); PL.cell(r, 2, note)
            for y in YRS:
                v = f(y)
                if v is not None: PL[f"{YC[y]}{r}"] = v; PL[f"{YC[y]}{r}"].number_format = fm
            p[k] = r; r += 1
        segs = (("LG", "線性滑軌", "g_lt"), ("BS", "滾珠螺桿", "g_lt"), ("RB", "機器人", "g_lt_RB"), ("OT", "其他", "g_lt_OT"))
        for sg, nm, glt in segs:
            def gf(y, sg=sg, glt=glt):
                if y <= 2026: return None
                if y == 2027: return f"=g27_{sg}+d27_{s}"
                if y == 2028: return f"=g28_{sg}+d28_{s}"
                return f"=$F{r}+({glt}-$F{r})*({y}-2028)/4"
            prow("g" + sg, f"{nm} 核心年增", gf, "0.0%", "FY2027–28 輸入＋情境；之後線性收斂")
        for sg, nm, _ in segs:
            def cf(y, sg=sg):
                if y == 2025: return None
                if y == 2026: return f"=Hist!E{a26[sg]}"
                return f"={CL(2 + YRS.index(y))}{r}*(1+{YC[y]}{p['g' + sg]})"
            lab = {"LG": "線性滑軌：核心（工業循環）", "BS": "滾珠螺桿：核心（工業循環）", "RB": "機器人 核心", "OT": "其他 核心"}[sg]
            prow("c" + sg, lab, cf)
        prow("hLG", "線性滑軌：人形（微型滑軌與小螺桿，命題）", lambda y: None if y == 2025 else f"=Thesis!{YC[y]}{t['gd_t']}")
        prow("hBS", "滾珠螺桿：人形（行星滾柱絲杠，命題）", lambda y: None if y == 2025 else f"=Thesis!{YC[y]}{t['prs_t']}")
        prow("LG", "線性滑軌 營收", lambda y: f"=Hist!D{hr['LG']}" if y == 2025 else f"={YC[y]}{p['cLG']}+{YC[y]}{p['hLG']}", note="官方分部")
        prow("BS", "滾珠螺桿 營收", lambda y: f"=Hist!D{hr['BS']}" if y == 2025 else f"={YC[y]}{p['cBS']}+{YC[y]}{p['hBS']}", note="官方分部")
        prow("RB", "機器人 營收", lambda y: None if y == 2025 else f"={YC[y]}{p['cRB']}", note="官方分部（2026 年起單列；FY2025 在其他）")
        prow("OT", "其他 營收", lambda y: f"=Hist!D{hr['OT']}" if y == 2025 else f"={YC[y]}{p['cOT']}", note="官方分部（FY2025 舊口徑含機器人）")
        prow("rev", "合併營收", lambda y: f"=Hist!D{hr['rev']}" if y == 2025 else f"=Hist!E{a26['rev']}" if y == 2026 else "=" + "+".join(f"{YC[y]}{p[k]}" for k in ("LG", "BS", "RB", "OT")),
             note="FY2025＝年報；FY2026＝錨點；其後＝分部加總")
        prow("chk", "檢查：分部合計 − 合併營收", lambda y: f"=SUM({YC[y]}{p['LG']}:{YC[y]}{p['OT']})-{YC[y]}{p['rev']}", "0.000")
        prow("chk2", "檢查：子列合計 − 分部（線性滑軌＋滾珠螺桿）", lambda y: None if y == 2025 else f"=({YC[y]}{p['cLG']}+{YC[y]}{p['hLG']}-{YC[y]}{p['LG']})+({YC[y]}{p['cBS']}+{YC[y]}{p['hBS']}-{YC[y]}{p['BS']})", "0.000")
        prow("yoy", "合併營收年增", lambda y: None if y == 2025 else f"={YC[y]}{p['rev']}/{CL(2 + YRS.index(y))}{p['rev']}-1", "0.0%")
        prow("hsh", "人形（命題）營收 ÷ 合併營收", lambda y: None if y == 2025 else f"=({YC[y]}{p['hLG']}+{YC[y]}{p['hBS']})/{YC[y]}{p['rev']}", "0.0%")
        for sg, nm, _ in segs:
            prow("m" + sg, f"{nm} 分部利益率", lambda y, sg=sg: None if y == 2025 else (f"=Hist!E{a26['p' + sg]}/Hist!E{a26[sg]}" if y == 2026 else f"=m_{sg}+mg_{s}+mg_d"), "0.0%",
                 "FY2026＝錨點；之後＝9M26 實績＋情境＋Sens 加減")
        for sg, nm, _ in segs:
            prow("p" + sg, f"{nm} 分部利益", lambda y, sg=sg: None if y == 2025 else f"={YC[y]}{p[sg]}*{YC[y]}{p['m' + sg]}")
        prow("ebit", "營業利益（分部利益合計）", lambda y: f"=Hist!D{hr['op']}" if y == 2025 else "=" + "+".join(f"{YC[y]}{p['p' + k]}" for k in ("LG", "BS", "RB", "OT")))
        prow("om", "營業利益率", lambda y: f"={YC[y]}{p['ebit']}/{YC[y]}{p['rev']}", "0.0%")
        prow("nonop", "業外淨額", lambda y: f"=Hist!D{hr['pt']}-Hist!D{hr['op']}" if y == 2025 else (f"=Hist!J{q['nonop']}+nonop_r*Hist!D{a26['rev']}" if y == 2026 else f"=nonop_r*{YC[y]}{p['rev']}"),
             note="FY2026＝9M26 實績＋4Q 模型")
        prow("pt", "稅前淨利", lambda y: f"={YC[y]}{p['ebit']}+{YC[y]}{p['nonop']}")
        prow("tax", "所得稅", lambda y: f"=Hist!D{hr['tax']}" if y == 2025 else (f"=Hist!J{hr['tax']}+tax_r*({YC[y]}{p['pt']}-Hist!J{hr['pt']})" if y == 2026 else f"=tax_r*{YC[y]}{p['pt']}"))
        prow("ni", "本期淨利（合併）", lambda y: f"={YC[y]}{p['pt']}-{YC[y]}{p['tax']}")
        prow("nci", "非控制權益損益", lambda y: f"=Hist!D{hr['ni']}-Hist!D{hr['nip']}" if y == 2025 else (f"=Hist!J{hr['ni']}-Hist!J{hr['nip']}+nci_r*Hist!D{a26['rev']}" if y == 2026 else f"=nci_r*{YC[y]}{p['rev']}"))
        prow("nip", "歸屬母公司淨利", lambda y: f"={YC[y]}{p['ni']}-{YC[y]}{p['nci']}")
        prow("eps", "EPS（元）", lambda y: f"=Hist!D{hr['eps']}" if y == 2025 else f"={YC[y]}{p['nip']}/sh_v", "0.00", "FY2025 為申報基本 EPS")
        prow("da", "折舊攤銷", lambda y: f"=Hist!D{hr['da']}" if y == 2025 else (f"=Hist!J{hr['da']}*4/3" if y == 2026 else f"={YC[y]}{p['rev']}*(fc_da+(fc_da_ss-fc_da)*({y}-2027)/5)"))
        prow("cx", "資本支出（含預付設備款）", lambda y: f"=Hist!D{hr['cx']}" if y == 2025 else ("=capex_26" if y == 2026 else f"={YC[y]}{p['rev']}*(fc_cx+(fc_cx_ss-fc_cx)*({y}-2027)/5+fc_cx_d)"),
             note="FY2025 只含取得不動產廠房設備（預付設備款增加未取得）")
        prow("dwc", "營運資金增加", lambda y: None if y == 2025 else (f"=(Hist!D{b['wc']}-Hist!C{b['wc']})+fc_wc*(Hist!D{a26['rev']}-Hist!C{q['rev']})" if y == 2026 else f"=fc_wc*({YC[y]}{p['rev']}-{CL(2 + YRS.index(y))}{p['rev']})"),
             note="FY2026＝9M 實際變動（2025-12-31 → 2026-09-30）＋4Q 模型")
        prow("nopat", "NOPAT（＝EBIT ×（1 − 稅率））", lambda y: None if y == 2025 else f"={YC[y]}{p['ebit']}*(1-tax_r)")
        prow("fcff", "FCFF（＝NOPAT＋折舊 − 資本支出 − 營運資金增加）", lambda y: None if y == 2025 else f"={YC[y]}{p['nopat']}+{YC[y]}{p['da']}-{YC[y]}{p['cx']}-{YC[y]}{p['dwc']}")
        prow("chk3", "檢查：FCFF − (EBIT − EBIT×稅率 ＋ 折舊 − 資本支出 − 營運資金增加)", lambda y: None if y == 2025 else f"={YC[y]}{p['fcff']}-({YC[y]}{p['ebit']}-{YC[y]}{p['ebit']}*tax_r+{YC[y]}{p['da']}-{YC[y]}{p['cx']}-{YC[y]}{p['dwc']})", "0.000")
        prow("div", "現金股利發放", lambda y: None if y == 2025 else (f"=Hist!C{o['divpaid']}" if y == 2026 else f"=payout*{CL(2 + YRS.index(y))}{p['eps']}*sh_v"), note="FY2026＝實付；之後＝發放率 × 上年 EPS × 股數")
        prow("fcfe", "股東自由現金流扣股利（＝歸屬母公司淨利＋折舊 − 資本支出 − 營運資金增加 − 股利）", lambda y: None if y == 2025 else f"={YC[y]}{p['nip']}+{YC[y]}{p['da']}-{YC[y]}{p['cx']}-{YC[y]}{p['dwc']}-{YC[y]}{p['div']}")
        PR[s] = p; pr0 = r + 1
    for col, w in zip("ABCDEFGHIJ", [52, 34, 11, 11, 11, 11, 11, 11, 11, 11]): PL.column_dimensions[col].width = w
    for col, w in zip("ABCDEFGHIJ", [52, 8, 11, 11, 11, 11, 11, 11, 11, 11]): TH.column_dimensions[col].width = w
    PL.freeze_panes = "C3"
    cells["eps_B"] = PR["B"]["eps"]
    # dps_tgt 公式用 PL!D{eps_B}
    _c = IN.cell(inrow["dps_tgt"], 5)
    if isinstance(_c.value, str): _c.value = _c.value.replace("§EPSB§", str(PR["B"]["eps"]))

    # ================= Val =================
    V = wb.create_sheet("Val")
    V["A1"] = "Val — 估值（DCF 50%＋分部加總 Comparable 50%；三情境 25／50／25）。網站不公開本檔目標價（台股）"
    vr = {}
    r = 3
    def vrow(k, lab, f, fm="#,##0.00", note="", cols=("C", "D", "E")):
        nonlocal r
        V.cell(r, 1, lab); V.cell(r, 2, note)
        if isinstance(f, str):
            V[f"C{r}"] = f; V[f"C{r}"].number_format = fm
        else:
            for c, s in zip(cols, "ABC"):
                v = f(s)
                if v is not None: V[f"{c}{r}"] = v; V[f"{c}{r}"].number_format = fm
        vr[k] = r; r += 1
    V.cell(r, 1, "0. 日期與期間").font = BOLD; r += 1
    vrow("px", "現價（元）", "=px", note="股價基準日見下")
    vrow("pxd", "股價基準日", "=px_date", "yyyy-mm-dd")
    vrow("tpd", "目標時點（＝EDATE(股價基準日,12)）", "=tp_date", "yyyy-mm-dd")
    vrow("valy", "模型折現期間：評價日 2026-12-31 → 目標時點（年）", "=val_yrs", "0.00")
    vrow("holdy", "持有期間：股價基準日 → 目標時點（年）", "=hold_yrs", "0.00", "上檔空間以此期間計，不是年化報酬")
    r += 1
    V.cell(r, 1, "1. 折現率").font = BOLD; r += 1
    vrow("ke", "股權成本 ke＝rf＋β ×（ERP＋CRP）", "=rf+beta*(erp+crp)", "0.00%")
    vrow("kdat", "稅後債務成本", "=kd*(1-tax_w)", "0.00%")
    vrow("we", "權益占比（市值）", "=px*sh_v/(px*sh_v+debt_v+lease_v)", "0.0%")
    vrow("wacc", "WACC", f"=C{r - 1}*C{r - 3}+(1-C{r - 1})*C{r - 2}+wacc_adj", "0.00%")
    vrow("g", "永續成長率 g", "=g_tv", "0.00%")
    vrow("ronic", "RONIC（＝WACC＋ronic_sp）", f"=C{vr['wacc']}+ronic_sp", "0.00%")
    for nm_, ref in (("ke", "ke"), ("wacc", "wacc")):
        defname(wb, nm_, f"Val!$C${vr[ref]}")
    r += 1
    V.cell(r, 1, "2. DCF（FCFF；明確期 FY2027–FY2032 期中折現；終值＝NOPAT(FY2033) ×（1 − g ÷ RONIC）÷（WACC − g））").font = BOLD; r += 1
    hdr(V, r, ["項目", "說明", "情境 A", "情境 B", "情境 C", "加權"]); r += 1
    W, G_, RO = f"$C${vr['wacc']}", f"$C${vr['g']}", f"$C${vr['ronic']}"
    def pv(s):
        p = PR[s]
        return "=" + "+".join(f"PL!{YC[y]}{p['fcff']}/(1+{W})^{i + 0.5}" for i, y in enumerate(FY))
    vrow("pv", "明確期 FCFF 現值", pv, "#,##0")
    vrow("tv", "終值（FY2032 年底）", lambda s: f"=PL!J{PR[s]['nopat']}*(1+{G_})*(1-{G_}/{RO})/({W}-{G_})", "#,##0")
    vrow("pvtv", "終值現值", lambda s: f"={'CDE'['ABC'.index(s)]}{r - 1}/(1+{W})^6", "#,##0")
    vrow("ev", "企業價值 EV", lambda s: f"={'CDE'['ABC'.index(s)]}{vr['pv']}+{'CDE'['ABC'.index(s)]}{vr['pvtv']}", "#,##0")
    vrow("nc", "＋淨現金（2026-09-30：現金 − 借款 − 租賃）", lambda s: "=cash_v-debt_v-lease_v", "#,##0")
    def stub(s):
        p = PR[s]; d = "D"
        return (f"=(Hist!D{a26['op']}*(1-tax_r))+PL!D{p['da']}/4-(capex_26-Hist!J{hr['cx']}-Hist!J{hr['cxp']})-fc_wc*(Hist!D{a26['rev']}-Hist!C{q['rev']})")
    vrow("stub", "＋Stub 期（4Q26）FCFF", stub, "#,##0", "4Q26 營業利益 ×（1 − 稅率）＋折舊 ÷ 4 − 4Q 資本支出（全年指引 − 9M 實付）− 營運資金增加")
    vrow("invv", "＋非營業投資 × 計入比例", lambda s: "=inv_v*inv_k", "#,##0")
    vrow("mi", "− 非控制權益", lambda s: "=-mi_v", "#,##0")
    vrow("eq", "股權價值", lambda s: "=" + "+".join(f"{'CDE'['ABC'.index(s)]}{vr[k]}" for k in ("ev", "nc", "stub", "invv", "mi")), "#,##0")
    vrow("ps", "每股價值（評價日）", lambda s: f"={'CDE'['ABC'.index(s)]}{vr['eq']}/sh_v")
    vrow("dcf12", "DCF 情境價值推到 12 個月後（×（1＋ke）− 期間股利）", lambda s: f"=MAX(0,{'CDE'['ABC'.index(s)]}{vr['ps']}*(1+ke)^t_tgt-dps_tgt)")
    V[f"F{vr['dcf12']}"] = f"=SUMPRODUCT(C{vr['dcf12']}:E{vr['dcf12']},C{vr['dcf12'] + 4}:E{vr['dcf12'] + 4})"
    vrow("tvsh", "終值現值占 EV", lambda s: f"=IF({'CDE'['ABC'.index(s)]}{vr['ev']}>0,{'CDE'['ABC'.index(s)]}{vr['pvtv']}/{'CDE'['ABC'.index(s)]}{vr['ev']},\"N/M\")", "0.0%", "超過 80% 時標示")
    vrow("tvm", "隱含終值 EV／EBITDA", lambda s: f"={'CDE'['ABC'.index(s)]}{vr['tv']}/((PL!J{PR[s]['ebit']}+PL!J{PR[s]['da']})*(1+{G_}))", "0.0")
    vrow("dchk", "檢查：EV ＞ 0 且 終值占比 ＜ 80%", lambda s: f"=IF(AND({'CDE'['ABC'.index(s)]}{vr['ev']}>0,{'CDE'['ABC'.index(s)]}{vr['tvsh']}<0.8),\"通過\",\"標示\")", "@")
    vrow("w", "情境權重", lambda s: f"=w_{s}", "0%")
    V[f"F{vr['w']}"] = f"=SUM(C{vr['w']}:E{vr['w']})"
    assert vr["w"] == vr["dcf12"] + 4
    vrow("dcf", "DCF 目標價（情境加權）", f"=F{vr['dcf12']}", "#,##0.00")
    r += 1
    V.cell(r, 1, "3. Comparable（分部加總；FY2028 分部淨利 × 同業前瞻 P/E；目標時點 2027-12-31 的 NTM＝FY2028）").font = BOLD; r += 1
    hdr(V, r, ["項目", "說明", "情境 A", "情境 B", "情境 C", "加權"]); r += 1
    for sg, nm, pe in (("LG", "線性滑軌", "pe_lin"), ("BS", "滾珠螺桿", "pe_lin"), ("RB", "機器人", "pe_rob"), ("OT", "其他", "pe_lin")):
        vrow("ni" + sg, f"{nm} FY2028 分部淨利（分部利益 ×（1 − 稅率）＋業外與非控制權益依營收分攤）",
             lambda s, sg=sg: f"=PL!F{PR[s]['p' + sg]}*(1-tax_r)+(PL!F{PR[s]['nonop']}*(1-tax_r)-PL!F{PR[s]['nci']})*PL!F{PR[s][sg]}/PL!F{PR[s]['rev']}", "#,##0")
    vrow("nichk", "檢查：分部淨利合計 − 歸屬母公司淨利（FY2028）", lambda s: "=" + "+".join(f"{'CDE'['ABC'.index(s)]}{vr['ni' + k]}" for k in ("LG", "BS", "RB", "OT")) + f"-PL!F{PR[s]['nip']}", "0.000")
    vrow("sotp", "分部加總股權價值（線性傳動 × pe_lin；機器人 × pe_rob）",
         lambda s: f"=({'CDE'['ABC'.index(s)]}{vr['niLG']}+{'CDE'['ABC'.index(s)]}{vr['niBS']}+{'CDE'['ABC'.index(s)]}{vr['niOT']})*pe_lin+{'CDE'['ABC'.index(s)]}{vr['niRB']}*pe_rob", "#,##0")
    vrow("cu", "Comparable（未調整，每股）", lambda s: f"={'CDE'['ABC'.index(s)]}{vr['sotp']}/sh_v")
    vrow("ca", "Comparable（相對倍數調整後，每股；× rel_k）", lambda s: f"={'CDE'['ABC'.index(s)]}{vr['cu']}*rel_k")
    V[f"F{vr['cu']}"] = f"=SUMPRODUCT(C{vr['cu']}:E{vr['cu']},C{vr['w']}:E{vr['w']})"
    V[f"F{vr['ca']}"] = f"=SUMPRODUCT(C{vr['ca']}:E{vr['ca']},C{vr['w']}:E{vr['w']})"
    vrow("comp", "Comparable 目標價（情境加權，進加權）", f"=F{vr['ca']}")
    vrow("relnow", "上銀現行前瞻 P/E ÷ 線性傳動同業中位數（顯示用）", f"={cells['hw_pe']}/pe_lin", "0.00", "10 年比值：未取得 → rel_k＝1.00（未調整）")
    r += 1
    V.cell(r, 1, "4. 加權目標價與評等").font = BOLD; r += 1
    hdr(V, r, ["項目", "說明", "情境 A", "情境 B", "情境 C", "加權"]); r += 1
    vrow("tps", "情境目標價（DCF × dcf_w ＋ Comparable ×（1 − dcf_w））", lambda s: f"=dcf_w*{'CDE'['ABC'.index(s)]}{vr['dcf12']}+(1-dcf_w)*{'CDE'['ABC'.index(s)]}{vr['ca']}")
    V[f"F{vr['tps']}"] = f"=SUMPRODUCT(C{vr['tps']}:E{vr['tps']},C{vr['w']}:E{vr['w']})"
    vrow("tp", "加權目標價（元）", f"=F{vr['tps']}", "#,##0.00", "網站不公開目標價（台股）")
    vrow("tpchk", "檢查：加權 − (dcf_w × DCF ＋ (1 − dcf_w) × Comparable)", f"=C{vr['tp']}-(dcf_w*C{vr['dcf']}+(1-dcf_w)*C{vr['comp']})", "0.000")
    vrow("rmin", "情境區間下緣（全部成員 MIN）", f"=MIN(C{vr['tps']}:E{vr['tps']},C{vr['tp']})")
    vrow("rmax", "情境區間上緣（全部成員 MAX）", f"=MAX(C{vr['tps']}:E{vr['tps']},C{vr['tp']})")
    vrow("rchk", "檢查：基準落在區間內", f"=IF(AND(D{vr['tps']}>=C{vr['rmin']},D{vr['tps']}<=C{vr['rmax']}),\"通過\",\"不通過\")", "@")
    vrow("up", "上檔空間（加權目標價 ÷ 現價 − 1；持有期間）", f"=C{vr['tp']}/px-1", "0.0%")
    vrow("rate", "估值評等（只看上檔空間）", f"=IF(C{vr['up']}>0.5,\"強力買進\",IF(C{vr['up']}>0.25,\"買進\",IF(C{vr['up']}>=-0.25,\"中立\",IF(C{vr['up']}>=-0.5,\"放空\",\"強力放空\"))))", "@")
    V[f"G{vr['rate']}"] = "級距：＞＋50% 強力買進、＞＋25% 買進、≥ −25% 中立、≥ −50% 放空、其餘強力放空（同 tools/target_summary）"
    vrow("gate", "收支關卡（FY2026–FY2030 股東自由現金流扣股利合計；A／B／C）", lambda s: f"=SUM(PL!D{PR[s]['fcfe']}:H{PR[s]['fcfe']})", "#,##0")
    vrow("gateB", "入不敷出判定用（情境 B 合計）", f"=D{vr['gate']}", "#,##0")
    vrow("gatej", "入不敷出判定", f"=IF(C{vr['gateB']}<0,\"不通過（入不敷出）\",\"通過\")", "@")
    vrow("rec", "投資建議", f"=IF(AND(C{vr['gateB']}<0,OR(C{vr['rate']}=\"買進\",C{vr['rate']}=\"強力買進\")),\"收支不足\",C{vr['rate']})", "@")
    V[f"G{vr['rec']}"] = "原則第 10 節第 63 條：收支關卡不通過且估值評等為買進以上時為「收支不足」"
    vrow("dist", "距最近級距邊界（百分點）", f"=MIN(ABS(C{vr['up']}-0.5),ABS(C{vr['up']}-0.25),ABS(C{vr['up']}+0.25),ABS(C{vr['up']}+0.5))*100", "0.0")
    vrow("flipup", "目標價達此值即升一級", f"=px*(1+IF(C{vr['up']}>0.5,99,IF(C{vr['up']}>0.25,0.5,IF(C{vr['up']}>=-0.25,0.25,IF(C{vr['up']}>=-0.5,-0.25,-0.5)))))", "#,##0.00")
    vrow("flipdn", "目標價低於此值即降一級", f"=px*(1+IF(C{vr['up']}>0.5,0.5,IF(C{vr['up']}>0.25,0.25,IF(C{vr['up']}>=-0.25,-0.25,IF(C{vr['up']}>=-0.5,-0.5,-99)))))", "#,##0.00")
    vrow("a_dcf", "上檔空間歸因：DCF 貢獻", f"=dcf_w*(C{vr['dcf']}/px-1)", "0.0%")
    vrow("a_cu", "上檔空間歸因：未調整 Comparable 貢獻", f"=(1-dcf_w)*(F{vr['cu']}/px-1)", "0.0%")
    vrow("a_rel", "上檔空間歸因：相對倍數調整貢獻", f"=(1-dcf_w)*(F{vr['ca']}-F{vr['cu']})/px", "0.0%")
    vrow("a_chk", "檢查：三段合計 − 上檔空間", f"=C{vr['a_dcf']}+C{vr['a_cu']}+C{vr['a_rel']}-C{vr['up']}", "0.0000")
    vrow("dcfsh", "DCF 在加權目標價的占比", f"=dcf_w*C{vr['dcf']}/C{vr['tp']}", "0.0%")
    vrow("humsh", "人形命題子列在 FY2028 分部淨利的占比（情境 B，顯示用）",
         f"=(PL!F{PR['B']['hLG']}*PL!F{PR['B']['mLG']}+PL!F{PR['B']['hBS']}*PL!F{PR['B']['mBS']})*(1-tax_r)/PL!F{PR['B']['nip']}", "0.0%")
    V.column_dimensions["A"].width = 62; V.column_dimensions["B"].width = 48
    for c in "CDEF": V.column_dimensions[c].width = 13
    cells["vr"] = vr

    # ================= Sens =================
    SE = wb.create_sheet("Sens")
    SE["A1"] = ("Sens — 敏感度（整本重算快照：複製本檔、改 In 頁 E 欄一個鍵或一組鍵、LibreOffice 重算後讀 Val；build_hiwin.py 產生）。"
                "一次動一個參數；情境與權重不變。估值評等欄只看上檔空間")
    SE["A2"] = "g_lt_OT、n_scr、ramp 不列：無區間或只用於換算（n_scr 只進 Recon）"
    SE["A3"] = "build_id"; SE["D3"] = sens["build_id"] if sens else "（基準檔）"
    SE["A4"] = "In 頁輸入指紋（敏感度重算時）"; SE["D4"] = sens["fp"] if sens else 0
    SE["A5"] = "In 頁輸入指紋（本檔即時）"; SE["D5"] = f"=SUM(In!E4:E{last_in - 1})"
    SE["A6"] = "基準目標價（敏感度重算時）"; SE["D6"] = sens["tp0"] if sens else 0
    SE["A7"] = "模型狀態"; SE["D7"] = f'=IF(AND(ABS(D5-D4)<=1E-6*MAX(1,ABS(D4)),ABS(Val!C{vr["tp"]}-D6)<1E-6),"CURRENT","STALE")'
    hdr(SE, 9, ["驅動", "設定", "鍵", "加權目標價", "相對基準", "DCF 目標價", "DCF 相對基準", "估值評等", "讀法"])
    if sens:
        rr = 10
        for case in sens["cases"]:
            SE.cell(rr, 1, case["name"]); SE.cell(rr, 2, case["set"]); SE.cell(rr, 3, case["keys"])
            SE.cell(rr, 4, case["tp"]).number_format = "#,##0.00"
            SE.cell(rr, 5, f"=D{rr}/Val!$C${vr['tp']}-1").number_format = "+0.0%;-0.0%"
            SE.cell(rr, 6, case["dcf"]).number_format = "#,##0.00"
            SE.cell(rr, 7, f"=F{rr}/Val!$C${vr['dcf']}-1").number_format = "+0.0%;-0.0%"
            SE.cell(rr, 8, f'=IF(D{rr}/px-1>0.5,"強力買進",IF(D{rr}/px-1>0.25,"買進",IF(D{rr}/px-1>=-0.25,"中立",IF(D{rr}/px-1>=-0.5,"放空","強力放空"))))')
            SE.cell(rr, 9, case.get("note", ""))
            rr += 1
        SE.cell(rr + 1, 1, "註：相對基準欄以本檔 Val 頁加權目標價（即時）為分母；模型狀態 STALE 表示 In 頁改過、須重跑 build_hiwin.py")
    for col, w in zip("ABCDEFGHI", [46, 40, 26, 12, 10, 12, 10, 10, 50]): SE.column_dimensions[col].width = w

    # ================= PRS 敏感度（2D） =================
    S2 = wb.create_sheet("PRS敏感度")
    S2["A1"] = "行星滾柱絲杠（PRS）價值占每台 BOM × 上銀可及市場份額 → 加權目標價（整本重算快照）。線 2b：覓途 19% vs MS 推算 35–69%（中點 52%）"
    S2["A2"] = "第二表：同一格點的 FY2030 人形（命題）營收 ÷ 合併營收（情境 B）"
    if sens:
        g = sens["grid"]
        hdr(S2, 4, ["prs_sh ＼ hw_prs"] + [f"{x:.0%}" for x in g["hw"]])
        for i, ps in enumerate(g["prs"]):
            S2.cell(5 + i, 1, f"{ps:.0%}")
            for j, _ in enumerate(g["hw"]):
                S2.cell(5 + i, 2 + j, g["tp"][i][j]).number_format = "#,##0"
        r2 = 6 + len(g["prs"])
        hdr(S2, r2, ["prs_sh ＼ hw_prs"] + [f"{x:.0%}" for x in g["hw"]])
        for i, ps in enumerate(g["prs"]):
            S2.cell(r2 + 1 + i, 1, f"{ps:.0%}")
            for j, _ in enumerate(g["hw"]):
                S2.cell(r2 + 1 + i, 2 + j, g["hs"][i][j]).number_format = "0.0%"
        S2.cell(r2 + 2 + len(g["prs"]), 1, f"現價 {sens['px']:.2f}；基準 prs_sh 25%、hw_prs 5%。評等邊界：升買進需加權目標價 ＞ 現價 × 1.25；降放空需 ＜ 現價 × 0.75")
    S2.column_dimensions["A"].width = 18

    # ================= Recon =================
    RC = wb.create_sheet("Recon")
    RC["A1"] = "Recon — 對照公司說法與市場（Interested-party 只作參照，不當參數）；反解為整本重算快照"
    hdr(RC, 3, ["項目", "模型值", "對照值", "差／隱含", "標記", "來源", "說明"])
    rc = [
        ("FY2027 合併營收年增（情境 B）", f"=PL!E{PR['B']['yoy']}", 0.10, "=B4-C4", "Interested-party", S_CT1111, "公司：2027 年可能維持雙位數成長（對照值取下緣 10%）"),
        ("4Q26 營收相對 3Q26", "=q4g", 0.0, "=B5-C5", "Interested-party", S_UA1110, "公司：4Q 營收可望優於 3Q（＞0）"),
        ("FY2026 機器人營收占比", f"=PL!D{PR['B']['RB']}/PL!D{PR['B']['rev']}", "=Hist!C" + str(o["rbsh"]), "=B6-C6", "Interested-party", S_CT1111, "公司：全年可望超過 12%"),
        ("分析師平均目標價（17 家）", f"=Val!C{vr['tp']}", 393.65, "=B7/C7-1", "Interested-party", S_SA.replace("<市場>/<代號>", "tpe/2049") + "（資料日 2026-10-22）", "賣方（有交易與承銷利益）；財報前，只作歷史對照"),
        ("美系外資目標價", f"=Val!C{vr['tp']}", 450, "=B8/C8-1", "Interested-party", S_CTEE1111, "外資「加碼」評等；財報後"),
        ("隱含每支 PRS 單價 2027（美元；＝每台 PRS 價值 ÷ n_scr）", f"=Thesis!E{TR['B']['prs']}/n_scr", 2025, "=B9/C9-1", "Interested-party", S_MS, "MS 1,350–2,700 美元（中點 2,025）；口徑：MS 為 Optimus Gen 2 類全尺寸"),
        ("FY2030 人形命題營收 ÷ 合併營收（情境 B）", f"=PL!H{PR['B']['hsh']}", None, None, "Derived", "模型輸出，非獨立證據", ""),
        ("FY2030 全球人形 PRS 市場（百萬美元，情境 B）", f"=Thesis!H{TR['B']['mkt']}", None, None, "Derived", "模型輸出", ""),
    ]
    for i, (a, b_, c_, d_, e_, f_, g_) in enumerate(rc):
        rr = 4 + i
        RC.cell(rr, 1, a); RC.cell(rr, 2, b_)
        if c_ is not None: RC.cell(rr, 3, c_)
        if d_ is not None: RC.cell(rr, 4, d_)
        RC.cell(rr, 5, e_); RC.cell(rr, 6, f_); RC.cell(rr, 7, g_)
        pct = any(x in a for x in ("年增", "占比", "季增", "相對", "÷"))
        for c in "BC": RC[f"{c}{rr}"].number_format = "0.0%" if pct else "#,##0.00"
        RC[f"D{rr}"].number_format = "+0.0%;-0.0%" if ("目標價" in a or "單價" in a) else ("0.0%" if pct else "#,##0.00")
    rr = 4 + len(rc) + 1
    RC.cell(rr, 1, "反解（整本重算；原則第 10 節第 4、36 條）").font = BOLD; rr += 1
    hdr(RC, rr, ["主導變數", "基準值", "支持現價所需值", "升一級（買進）所需值", "降一級（放空）所需值", "說明"]); rr += 1
    cells["rev_row"] = rr
    if sens:
        for item in sens["rev"]:
            RC.cell(rr, 1, item["name"]);
            for j, k in ((2, "base"), (3, "px"), (4, "up"), (5, "dn")):
                v = item.get(k)
                RC.cell(rr, j, v if v is not None else "超出可行範圍")
                if isinstance(v, (int, float)): RC.cell(rr, j).number_format = item.get("fmt", "0.0%")
            RC.cell(rr, 6, item.get("note", "")); rr += 1
    for col, w in zip("ABCDEFG", [56, 14, 14, 16, 16, 60, 50]): RC.column_dimensions[col].width = w

    # ================= 對帳 =================
    DZ = wb.create_sheet("對帳")
    DZ["A1"] = "對帳 — 本模型使用的線層級變數（與 physical_ai/lines/ 單位經濟模型對帳用；本代理讀不到其完成版，取值依線框架 v0.1）"
    hdr(DZ, 3, ["變數名", "所屬線", "本模型使用值（或區間）", "單位", "口徑", "出處", "標記", "影響本模型哪些儲存格"])
    dz = [
        ("全球人形機器人出貨 2026", "2b", "=TEXT(hu_26,\"#,##0\")", "台", "全球、各尺寸混合", S_TF, "Interested-party", "Thesis 各情境出貨列 → PL 人形子列"),
        ("全球人形機器人出貨 2030／2032（A／B／C）", "2b", "=TEXT(hu30_A,\"#,##0\")&\"／\"&TEXT(hu30_B,\"#,##0\")&\"／\"&TEXT(hu30_C,\"#,##0\")&\"；\"&TEXT(hu32_A,\"#,##0\")&\"／\"&TEXT(hu32_B,\"#,##0\")&\"／\"&TEXT(hu32_C,\"#,##0\")", "台", "全球、全尺寸＋小型混合（線 2b 無預測）", "本模型假設", "Assumed", "In hu30_*／hu32_* → Thesis → PL"),
        ("每台硬體 BOM 2027／2032", "2b", "=TEXT(bom_27,\"#,##0\")&\"／\"&TEXT(bom_32,\"#,##0\")&\"（2032 區間 13,000–20,000）\"", "US$", "全尺寸、不含軟體", S_BOFA, "Interested-party", "In bom_27／bom_32 → Thesis BOM 列"),
        ("行星滾柱絲杠價值 ÷ BOM", "2b", "=TEXT(prs_sh,\"0%\")&\"（區間 19%–52%；MS 推算上緣 69%）\"", "%", "每台全尺寸人形；含 12–14 支", S_2B, "Assumed", "In prs_sh → Thesis PRS 列 → PL 滾珠螺桿 人形子列；PRS敏感度 頁"),
        ("每台行星滾柱絲杠支數", "2b", "=TEXT(n_scr,\"0\")&\"（MS 12–14）\"", "支", "Optimus Gen 2 類", S_MS, "Interested-party", "Recon 隱含每支單價"),
        ("非中國供應鏈整機占出貨", "2b", "=TEXT(ex_cn,\"0%\")&\"（30%–50%）\"", "%", "全球出貨台數口徑", S_2B, "Assumed", "In ex_cn → Thesis"),
        ("微型滑軌與小螺桿價值 ÷ BOM", "2b", "=TEXT(gd_sh,\"0%\")&\"（1%–4%）\"", "%", "手臂、靈巧手線性元件", "未取得公開拆分", "Assumed", "In gd_sh → PL 線性滑軌 人形子列"),
        ("BOM 學習曲線（2027→2032 幾何下降）", "2b", "=TEXT(1-(bom_32/bom_27)^(1/5),\"0.0%\")&\" ／年\"", "%／年", "線 2b 以累積產量倍增定義學習率（10–25%）；本模型以年降幅表示，口徑不同", S_2B, "Derived", "Thesis BOM 列"),
        ("全球工業機器人裝機 2025／IFR 預測 2026、2029", "2a", "60.3 萬／65.5 萬／80.6 萬台", "台", "IFR 年度裝機", "IFR（線 2a 第 1、3 節）", "Verified", "g_lt_RB 區間依據（不直接進公式）"),
        ("JARA 受注 2026 年 7–9 月年增", "2a", "+46.6%", "%", "日本機器人工業會會員受注額", S_JARA, "Verified", "g27_* 依據（不直接進公式）"),
        ("哈默納科受注÷營收（2026 年 4–6 月）", "2a", "約 1.45", "x", "公司決算短信", "線 2a 第 5 節", "Derived", "工業循環位置參照（不直接進公式）"),
        ("美國工人全負擔時薪", "2b", "46.89", "US$／時", "BLS ECEC 2026-06 私部門", "線 2b 第 3.3 節", "Verified", "未使用（TCO 屬線 2b；本模型只取 BOM）"),
    ]
    for i, row in enumerate(dz):
        for j, v in enumerate(row, 1): DZ.cell(4 + i, j, v)
    DZ.cell(5 + len(dz), 1, "差異：線 2b 學習率以「累積產量倍增」定義，本模型以年降幅（BOM 2027→2032）表示，口徑不同，照寫待對帳；本模型未改口徑")
    for col, w in zip("ABCDEFGH", [40, 8, 44, 10, 36, 50, 14, 50]): DZ.column_dimensions[col].width = w

    # ================= Sources =================
    SO = wb.create_sheet("Sources")
    SO["A1"] = "Sources（公司原文經 financialfilings.com 擷取與網頁擷取工具讀取；表格以加總核對，見 Hist 檢查列）"
    for i, s_ in enumerate([S_Q3, S_Q2, S_Q1, S_FY, S_PX, S_CT1111, S_UA1110, S_UA1110b, S_CNA, S_CTEE0602, S_UA2511, S_CTEE1111, S_OCT, S_UDN25, S_SA, S_DAMO, S_RF, S_RFUS, S_FX, S_2B, S_MS, S_MITU, S_BOFA, S_TF, S_JARA], 3):
        SO.cell(i, 1, s_)
    SO.column_dimensions["A"].width = 200

    # ================= README（工作表） =================
    RD = wb.create_sheet("README", 0)
    bid = sens["build_id"] if sens else "（基準檔）"
    T = lambda ref, f="#,##0": f'TEXT({ref},"{f}")'
    VC = lambda k, c="C": f"Val!{c}{vr[k]}"
    RD["A1"] = (f'="上銀 HIWIN（2049.TW）個股模型 v{VER}｜build_id {bid}｜資料截止日 "&TEXT(px_date,"yyyy-mm-dd")&"｜成品日 2026-11-16｜股價基準日 "&TEXT(px_date,"yyyy-mm-dd")&"（收盤 "&TEXT(px,"0.00")&" 元）"')
    lines = [
        '="結論（條件式）：在 分部利益率維持 9M26 水準、FY2027 核心成長 "&TEXT(g27_LG,"0%")&"（線性滑軌）、行星滾柱絲杠占人形 BOM "&TEXT(prs_sh,"0%")&"、上銀可及份額 "&TEXT(hw_prs,"0%")&" 的基準假設下，加權目標價 "&' + T(VC("tp"), "#,##0") + '&" 元，上檔空間 "&' + T(VC("up"), "+0.0%;-0.0%") + '&"，估值評等「"&' + VC("rate") + '&"」，投資建議「"&' + VC("rec") + '&"」（網站不公開目標價）"',
        '="估值拆解：DCF "&' + T(VC("dcf"), "#,##0") + '&" 元（WACC "&TEXT(wacc,"0.0%")&"、g "&TEXT(g_tv,"0.00%")&"、情境 B 終值占 EV "&' + T(VC("tvsh", "D"), "0%") + '&"）；Comparable "&' + T(VC("comp"), "#,##0") + '&" 元（線性傳動 P/E "&TEXT(pe_lin,"0.0")&"、機器人 P/E "&TEXT(pe_rob,"0.0")&"，FY2028 分部淨利）；情境 A／B／C "&' + T(VC("tps", "C")) + '&"／"&' + T(VC("tps", "D")) + '&"／"&' + T(VC("tps", "E")) + '&" 元"',
        '="上檔空間歸因：DCF "&' + T(VC("a_dcf"), "+0.0%;-0.0%") + '&"＋未調整 Comparable "&' + T(VC("a_cu"), "+0.0%;-0.0%") + '&"＋相對倍數調整 "&' + T(VC("a_rel"), "+0.0%;-0.0%") + '&"；上銀現行前瞻 P/E 為線性傳動同業中位數的 "&' + T(VC("relnow"), "0.00") + '&" 倍（10 年比值未取得，rel_k＝1.00 未調整）"',
        '="評等距最近級距邊界 "&' + T(VC("dist"), "0.0") + '&" 個百分點；加權目標價達 "&' + T(VC("flipup"), "#,##0") + '&" 元升一級、低於 "&' + T(VC("flipdn"), "#,##0") + '&" 元降一級"',
        '="收支關卡：情境 B FY2026–FY2030 股東自由現金流扣股利合計 "&' + T(VC("gateB"), "#,##0") + '&" 百萬元，"&' + VC("gatej") + '&"；期間：折現 "&' + T(VC("valy"), "0.00") + '&" 年、持有 "&' + T(VC("holdy"), "0.00") + '&" 年"',
        '="命題：上銀被定價的核心問題＝工業循環（線性滑軌、滾珠螺桿占 9M26 營收 "&TEXT((Hist!J' + str(hr["LG"]) + '+Hist!J' + str(hr["BS"]) + ')/Hist!J' + str(hr["rev"]) + ',"0%")&"）上行能走多久，以及人形機器人線性關節（行星滾柱絲杠）能否成為第二成長曲線；情境 B 的人形子列 FY2030 只占營收 "&TEXT(PL!H' + str(PR["B"]["hsh"]) + ',"0.0%")&"（模型輸出，非獨立證據）"',
        "分部：線性滑軌、滾珠螺桿、機器人、其他（官方部門；2026 年起機器人單列）。人形（命題）子列掛在線性滑軌（微型滑軌與小螺桿）與滾珠螺桿（行星滾柱絲杠）之下，子列合計＝分部",
        '="最大不確定：行星滾柱絲杠占人形 BOM（覓途 19% vs MS 推算 35–69%）。Sens：prs_sh 19% → 加權目標價 "&TEXT(Sens!D' + "{s_prs_lo}" + ',"#,##0")&" 元、52% → "&TEXT(Sens!D' + "{s_prs_hi}" + ',"#,##0")&" 元；完整格點見 PRS敏感度 頁"',
        "失效條件（可觀察）：(1) 月營收年增連續兩個月轉負或接單能見度縮回 3 個月內（工業循環反轉，偏向情境 C）；(2) 單季營業利益率跌回 10% 以下（9M26 水準不可維持）；(3) 主要非中國人形整機（Tesla、Figure）公開以滾珠螺桿或旋轉關節取代腿部行星滾柱絲杠，或上銀 2027 年未見人形量產訂單",
        "承重 Assumed 的可證偽指標：g27_*＝公開資訊觀測站月營收年增；mg_*＝季報分部利益率；prs_sh＝整機拆解報告或供應商報價；hw_prs、hw_gd＝上銀法說人形營收或訂單揭露；hu30_*＝IDC／TrendForce 季度出貨",
        "實體 AI 模型不讀 Siliconomics／Tokenomics／Wattonomics 協調檔；自用 AI 讀法（原則第 7 節）不適用；收支模組、融資可行性檢查（第 9、62 條）不適用（淨現金、無新股）",
        "HTML 一頁摘要：本版未做（時間）；display_check.json（第 51 條）未建，repo README.md 數字由建置腳本從重算後儲存格產生",
        "風險登錄（第 10 條）：中國價格競爭（陸系滑軌、絲杠）— 未建模，觀察陸系同業毛利率；匯率（新台幣升值）— 未建模，基準匯率固定，觀察 USD/TWD；交期過長致客戶轉單（公司自述）— 未建模，觀察交期與訂單能見度；地緣政治與關稅 — 未建模，觀察美國關稅公告",
        "版本沿革：v0.1（2026-11-16）首版",
    ]
    for i, l in enumerate(lines, 2):
        RD.cell(i, 1, l)
    RD.column_dimensions["A"].width = 200
    cells["readme_lines"] = len(lines)

    # README 引用 Sens 列號（prs 案例）
    if sens:
        for i in range(2, 2 + len(lines)):
            v = RD.cell(i, 1).value
            if "{s_prs_lo}" in v:
                RD.cell(i, 1, v.replace("{s_prs_lo}", str(sens["row_prs_lo"])).replace("{s_prs_hi}", str(sens["row_prs_hi"])))
    else:
        for i in range(2, 2 + len(lines)):
            v = RD.cell(i, 1).value
            if "{s_prs_lo}" in v: RD.cell(i, 1, "（基準檔：Sens 未建）")
    for ws in wb.worksheets:
        ws.sheet_view.zoomScale = 90
    wb.save(path)
    cells["PR"] = PR; cells["TR"] = TR
    return cells


# ---------------- 重算與讀值 ----------------
def recalc(paths, outdir):
    os.makedirs(outdir, exist_ok=True)
    for i in range(0, len(paths), 20):
        subprocess.run(["soffice", "--headless", "--calc", "--convert-to", "xlsx", "--outdir", outdir] + paths[i:i + 20],
                       check=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, timeout=900)
    return [os.path.join(outdir, os.path.basename(p)) for p in paths]


def readv(path, cells):
    wb = openpyxl.load_workbook(path, data_only=True)
    V = wb["Val"]; vr = cells["vr"]
    PLs = wb["PL"]
    return dict(tp=V[f"C{vr['tp']}"].value, dcf=V[f"C{vr['dcf']}"].value, up=V[f"C{vr['up']}"].value,
                hs30=PLs[f"H{cells['PR']['B']['hsh']}"].value)


def errscan(path):
    wb = openpyxl.load_workbook(path, data_only=True)
    bad = []
    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for c in row:
                if isinstance(c.value, str) and (c.value.startswith("#") or c.value.startswith("Err:")) and c.value[:4] in ("#DIV", "#VAL", "#REF", "#NAM", "#N/A", "#NUM", "#NUL", "Err:"):
                    bad.append(f"{ws.title}!{c.coordinate}={c.value}")
    return bad


def variant(base, out, cells, over):
    wb = openpyxl.load_workbook(base)
    IN = wb["In"]
    for k, v in over.items():
        IN.cell(cells["inrow"][k], 5).value = v
    wb.save(out)


def main():
    shutil.rmtree(TMP, ignore_errors=True); os.makedirs(TMP)
    base = os.path.join(TMP, "base.xlsx")
    cells = build(base)
    rb = recalc([base], os.path.join(TMP, "rc"))[0]
    bad = errscan(rb)
    if bad: print("基準檔錯誤值：", bad[:20]); sys.exit(1)
    b0 = readv(rb, cells)
    print("基準：", b0)
    wbv = openpyxl.load_workbook(rb, data_only=True)
    INv = wbv["In"]
    val = lambda k, col=5: INv.cell(cells["inrow"][k], col).value
    import datetime as _dt
    def _num(v):
        if isinstance(v, _dt.datetime): return (v - _dt.datetime(1899, 12, 30)).days
        return v if isinstance(v, (int, float)) and not isinstance(v, bool) else 0
    fp = sum(_num(INv.cell(r, 5).value) for r in range(4, cells["in_last"]))
    # ---- 案例 ----
    C = []
    def case(name, setting, keys, over, note=""):
        C.append(dict(name=name, set=setting, keys=keys, over=over, note=note))
    lohi = lambda k: (val(k, 4), val(k, 6))
    for k, nm in (("q4g", "4Q26 季增"), ("bom_32", "BOM 2032"), ("prs_sh", "PRS 占 BOM"), ("ex_cn", "非中國供應鏈比例"), ("hw_prs", "上銀 PRS 份額"),
                  ("gd_sh", "微型滑軌占 BOM"), ("hw_gd", "上銀微型滑軌份額"), ("g_lt", "長期核心成長"), ("g_lt_RB", "機器人長期成長"), ("mg_d", "分部利益率 ±2pt"),
                  ("nonop_r", "業外 ÷ 營收"), ("tax_r", "有效稅率"), ("beta", "β ±0.3"), ("erp", "ERP 4%／6%"), ("wacc_adj", "WACC ±1pt"),
                  ("fc_da_ss", "穩態折舊 ÷ 營收"), ("fc_cx_d", "資本支出 ÷ 營收 ±1pt"), ("fc_wc", "營運資金 ÷ 營收"), ("pe_lin", "線性傳動 P/E"), ("pe_rob", "機器人 P/E"),
                  ("dcf_w", "DCF 權重 30%／70%")):
        lo, hi = lohi(k)
        for tag, v in (("低", lo), ("高", hi)):
            if v is None or v == val(k): continue
            case(f"{nm}（{tag}）", f"{k}＝{v:,.0f}" if abs(v) >= 100 else f"{k}＝{v:.4g}", k, {k: v}, "PRS 為本模型最大不確定（線 2b）" if k == "prs_sh" else "")
    for tag, col in (("低", 4), ("高", 6)):
        ks = ["g27_LG", "g27_BS", "g27_RB", "g27_OT"]
        case(f"FY2027 核心成長（全部{tag}）", "四分部同時取" + tag, " ".join(ks), {k: val(k, col) for k in ks})
        ks = ["g28_LG", "g28_BS", "g28_RB", "g28_OT"]
        case(f"FY2028 核心成長（全部{tag}）", "四分部同時取" + tag, " ".join(ks), {k: val(k, col) for k in ks})
    case("永續成長率 −1pt", f"g_tv＝{val('g_tv', 4):.4g}", "g_tv", {"g_tv": val("g_tv", 4)})
    case("RONIC＝WACC＋5pt", "ronic_sp＝0.05", "ronic_sp", {"ronic_sp": 0.05})
    case("非營業投資不計", "inv_k＝0", "inv_k", {"inv_k": 0})
    case("美元口徑折現", f"rf 改 rf_us＝{val('rf_us'):.2%}（g＝MIN(3%, rf)）", "rf", {"rf": val("rf_us")}, "匯率固定；只換折現利率")
    case("人形命題歸零", "hw_prs＝0、hw_gd＝0", "hw_prs hw_gd", {"hw_prs": 0, "hw_gd": 0}, "只剩工業循環的價值")
    case("情境 A 權重 100%", "w_A＝1、w_B＝0、w_C＝0", "w_A w_B w_C", {"w_A": 1, "w_B": 0, "w_C": 0})
    case("情境 C 權重 100%", "w_A＝0、w_B＝0、w_C＝1", "w_A w_B w_C", {"w_A": 0, "w_B": 0, "w_C": 1})
    case("相對倍數係數＝1", "rel_k＝1（同基準）", "rel_k", {"rel_k": 1.0}, "原則第 47 條：一律列此案")
    # 2D 格點
    PRS = [0.19, 0.25, 0.35, 0.52, 0.69]; HW = [0.0, 0.02, 0.05, 0.10, 0.15]
    jobs = []
    for i, c in enumerate(C):
        p = os.path.join(TMP, f"case{i:02d}.xlsx"); variant(base, p, cells, c["over"]); jobs.append(p)
    for i, ps in enumerate(PRS):
        for j, hw in enumerate(HW):
            p = os.path.join(TMP, f"grid{i}{j}.xlsx"); variant(base, p, cells, {"prs_sh": ps, "hw_prs": hw}); jobs.append(p)
    # 反解第一輪：mg_d、hw_prs、g27 全部加減（兩點線性）
    probes = {"mg_d": [0.0, 0.03], "hw_prs": [0.05, 0.20], "prs_sh": [0.25, 0.69]}
    for k, vs in probes.items():
        for v in vs:
            p = os.path.join(TMP, f"rv_{k}_{v}.xlsx"); variant(base, p, cells, {k: v}); jobs.append(p)
    print(f"重算 {len(jobs)} 檔…")
    outs = recalc(jobs, os.path.join(TMP, "rc"))
    res = {os.path.basename(o): readv(o, cells) for o in outs}
    for i, c in enumerate(C):
        rv = res[f"case{i:02d}.xlsx"]; c["tp"] = rv["tp"]; c["dcf"] = rv["dcf"]
    grid = dict(prs=PRS, hw=HW, tp=[[res[f"grid{i}{j}.xlsx"]["tp"] for j in range(len(HW))] for i in range(len(PRS))],
                hs=[[res[f"grid{i}{j}.xlsx"]["hs30"] for j in range(len(HW))] for i in range(len(PRS))])
    px = val("px")
    targets = dict(px=px, up=px * 1.25, dn=px * 0.75)
    rev = []
    # 線性外插求解，再整本重算驗證一次
    sol_jobs = {}
    for k, vs in probes.items():
        t0, t1 = res[f"rv_{k}_{vs[0]}.xlsx"]["tp"], res[f"rv_{k}_{vs[1]}.xlsx"]["tp"]
        slope = (t1 - t0) / (vs[1] - vs[0])
        for nm, tg in targets.items():
            x = vs[0] + (tg - t0) / slope if slope else None
            sol_jobs[(k, nm)] = x
    vjobs = []
    for (k, nm), x in sol_jobs.items():
        lim = {"mg_d": (-0.2, 0.3), "hw_prs": (0.0, 1.0), "prs_sh": (0.0, 1.0)}[k]
        if x is None or not (lim[0] <= x <= lim[1]): sol_jobs[(k, nm)] = None; continue
        p = os.path.join(TMP, f"sv_{k}_{nm}.xlsx"); variant(base, p, cells, {k: x}); vjobs.append(p)
    vouts = recalc(vjobs, os.path.join(TMP, "rc2")) if vjobs else []
    vres = {os.path.basename(o): readv(o, cells)["tp"] for o in vouts}
    for k, nm_, fmt, note in (("mg_d", "分部利益率加減（全部分部、全部情境）", "+0.0%;-0.0%", "基準 0；3Q26 單季營業利益率較 9M26 高約 3.7pt（情境 A 的設定）"),
                              ("hw_prs", "上銀可及市場 PRS 份額（穩態）", "0.0%", "基準 5%；可行範圍 0–100%"),
                              ("prs_sh", "PRS 價值占人形 BOM", "0.0%", "基準 25%；MS 推算上緣 69%")):
        item = dict(name=nm_, base=val(k), fmt=fmt, note=note)
        chk = []
        for nm in ("px", "up", "dn"):
            x = sol_jobs[(k, nm)]
            item[nm] = x
            if x is not None:
                chk.append(f"{nm}：重算 {vres[f'sv_{k}_{nm}.xlsx']:.1f} vs 目標 {targets[nm]:.1f}")
        item["note"] = note + "；線性解後整本重算驗證（" + "；".join(chk) + "）" if chk else note
        rev.append(item)
    h = hashlib.sha1(json.dumps([[INv.cell(r, c).value for c in range(1, 7)] for r in range(4, cells["in_last"])], default=str, ensure_ascii=False).encode()).hexdigest()[:8]
    build_id = f"hiwin-v{VER}-{h}"
    # 案例列號（README 引用 prs）
    row_lo = 10 + next(i for i, c in enumerate(C) if c["keys"] == "prs_sh" and "低" in c["name"])
    row_hi = 10 + next(i for i, c in enumerate(C) if c["keys"] == "prs_sh" and "高" in c["name"])
    C.insert(0, dict(name="基準（本檔）", set="—", keys="—", tp=b0["tp"], dcf=b0["dcf"]))
    sens = dict(build_id=build_id, fp=fp, tp0=b0["tp"], cases=C, grid=grid, rev=rev, px=px, row_prs_lo=row_lo + 1, row_prs_hi=row_hi + 1)
    os.makedirs(DIST, exist_ok=True)
    fin = os.path.join(TMP, OUT)
    build(fin, sens)
    outp = recalc([fin], DIST)[0]
    bad = errscan(outp)
    print("錯誤值：", len(bad), bad[:10])
    json.dump(dict(build_id=build_id, base=b0, cases=[{k: c.get(k) for k in ("name", "keys", "tp", "dcf")} for c in C], rev=rev, grid=grid),
              open(os.path.join(HERE, "sens.json"), "w"), ensure_ascii=False, indent=1, default=str)
    return outp, cells


CHECKS = [["tools/in_layout/check_in_layout.py", "{x}"], ["tools/sens_coverage/check_sens_coverage.py", "{x}", "--require-keycol"],
          ["tools/model_governance/check_governance.py", "{x}"], ["tools/model_governance/check_rating_gate.py", "{x}", "--strict"],
          ["tools/segment_check/check_segments.py", HERE]]


def run_checks(x):
    out, ok = [], True
    for c in CHECKS:
        cmd = [sys.executable, os.path.join(ROOT, c[0])] + [a.format(x=x) for a in c[1:]]
        rp = subprocess.run(cmd, capture_output=True, text=True, cwd=ROOT)
        line = (rp.stdout.strip().splitlines() or [""])[0]
        out.append(f"{os.path.basename(c[0])}：{line}")
        ok &= rp.returncode == 0
    return ok, out


def write_readme(x, chk):
    wb = openpyxl.load_workbook(x, data_only=True)
    V, IN, RD, SE, RC = wb["Val"], wb["In"], wb["README"], wb["Sens"], wb["Recon"]
    lab = {V.cell(r, 1).value: r for r in range(1, V.max_row + 1) if V.cell(r, 1).value}
    g = lambda l, c="C": V[f"{c}{lab[l]}"].value
    inr = {IN.cell(r, 1).value: r for r in range(4, IN.max_row + 1) if IN.cell(r, 1).value}
    iv = lambda k: IN.cell(inr[k], 5).value
    cases = [(SE.cell(r, 1).value, SE.cell(r, 4).value, SE.cell(r, 6).value) for r in range(11, SE.max_row + 1) if isinstance(SE.cell(r, 4).value, (int, float))]
    tp0 = g("加權目標價（元）")
    top = sorted(cases, key=lambda c: -abs(c[1] / tp0 - 1))[:8]
    rv = [[RC.cell(r, c).value for c in range(1, 7)] for r in range(1, RC.max_row + 1) if RC.cell(r, 1).value in ("分部利益率加減（全部分部、全部情境）", "上銀可及市場 PRS 份額（穩態）", "PRS 價值占人形 BOM")]
    f = lambda v, fm: (format(v, fm) if isinstance(v, (int, float)) else str(v))
    S2 = wb["PRS敏感度"]
    grid = [[S2.cell(r, c).value for c in range(1, 7)] for r in range(4, 10)]
    L = []
    L.append(f"# physical_ai/hiwin/ — 上銀 HIWIN（2049.TW）個股模型 v{VER}")
    L.append("")
    L.append(f"{RD['A1'].value}")
    L.append("")
    L.append("> 亞太掛牌（台股）。網站不公開本檔目標價。本 README 的數字由 `build_hiwin.py` 從重算後的成品儲存格產生，不手寫。")
    L.append("")
    L.append("## 結論（條件式）")
    for r in range(2, 7): L.append(f"- {RD.cell(r, 1).value}")
    L.append("")
    L.append("| 項目 | 情境 A（樂觀） | 情境 B（基準） | 情境 C（保守） | 加權 |")
    L.append("|---|---|---|---|---|")
    for l, fm in (("DCF 情境價值推到 12 個月後（×（1＋ke）− 期間股利）", ",.0f"), ("Comparable（相對倍數調整後，每股；× rel_k）", ",.0f"), ("情境目標價（DCF × dcf_w ＋ Comparable ×（1 − dcf_w））", ",.0f")):
        L.append(f"| {l.split('（')[0]} | {f(g(l, 'C'), fm)} | {f(g(l, 'D'), fm)} | {f(g(l, 'E'), fm)} | {f(g(l, 'F'), fm)} |")
    L.append(f"| 收支關卡（FY2026–FY2030 股東 FCF 扣股利，NT$M） | {f(g('收支關卡（FY2026–FY2030 股東自由現金流扣股利合計；A／B／C）', 'C'), ',.0f')} | {f(g('收支關卡（FY2026–FY2030 股東自由現金流扣股利合計；A／B／C）', 'D'), ',.0f')} | {f(g('收支關卡（FY2026–FY2030 股東自由現金流扣股利合計；A／B／C）', 'E'), ',.0f')} | — |")
    L.append("")
    L.append(f"現價 {f(iv('px'), '.2f')} 元（TWSE 2026-11-16 收盤）；加權目標價 {f(tp0, ',.0f')} 元；上檔空間 {f(g('上檔空間（加權目標價 ÷ 現價 − 1；持有期間）'), '+.1%')}；"
             f"估值評等 {g('估值評等（只看上檔空間）')}；投資建議 {g('投資建議')}。WACC {f(g('WACC'), '.2%')}、ke {f(g('股權成本 ke＝rf＋β ×（ERP＋CRP）'), '.2%')}、β {f(iv('beta'), '.2f')}、g {f(iv('g_tv'), '.2%')}。")
    L.append("")
    L.append("## 命題")
    L.append("- 上銀被市場定價的核心問題：(1) 工業循環（工具機、半導體設備、自動化）的上行能走多久、9M26 的利潤率（營業利益率 3Q26 單季 17.8%）能否維持；(2) 人形機器人線性關節（行星滾柱絲杠、微型滑軌）能否成為第二成長曲線。")
    L.append("- 命題模組（Thesis 頁）：上銀 PRS 營收＝全球人形出貨 × 每台 BOM × PRS 占 BOM × 非中國供應鏈比例 × 上銀份額 × 爬坡；微型滑軌同式。結果掛在官方分部之下的子列，子列合計＝分部。")
    L.append(f"- 模型回答：在基準假設下，人形子列 FY2030 只占營收 {f(wb['PL']['H' + str([r for r in range(1, 60) if wb['PL'].cell(r, 1).value == '人形（命題）營收 ÷ 合併營收'][0])].value, '.1%')}；估值主要由工業循環的利潤率決定，PRS 的不確定對目標價影響有限（見下）。這是模型輸出，非獨立證據。")
    L.append("")
    L.append("## 行星滾柱絲杠（PRS）敏感度（線 2b 第 3 節：覓途 19% vs MS 推算 35–69%）")
    L.append("加權目標價（元），列＝PRS 占每台 BOM，欄＝上銀在非中國供應鏈的 PRS 份額（穩態）：")
    L.append("")
    L.append("| " + " | ".join(str(x) for x in grid[0]) + " |")
    L.append("|" + "---|" * len(grid[0]))
    for row in grid[1:]: L.append("| " + " | ".join(f(x, ',.0f') for x in row) + " |")
    L.append("")
    L.append("## 反解（整本重算）")
    L.append("| 主導變數 | 基準 | 支持現價所需 | 升一級（買進）所需 | 降一級（放空）所需 |")
    L.append("|---|---|---|---|---|")
    for rw in rv: L.append(f"| {rw[0]} | {f(rw[1], '.1%')} | {f(rw[2], '.1%')} | {f(rw[3], '.1%')} | {f(rw[4], '.1%')} |")
    L.append("")
    L.append("## 敏感度前八大（加權目標價相對基準；整本重算）")
    L.append("| 案例 | 加權目標價 | 相對基準 | DCF 目標價 |")
    L.append("|---|---|---|---|")
    for c in top: L.append(f"| {c[0]} | {f(c[1], ',.0f')} | {f(c[1] / tp0 - 1, '+.1%')} | {f(c[2], ',.0f')} |")
    L.append("")
    L.append("## 分部（官方）與口徑")
    L.append("- 線性滑軌、滾珠螺桿、機器人、其他（合併財報部門資訊）。2026 年各季報告起機器人單列，2025 比較期已重編；FY2024、FY2025 年報為舊口徑（機器人在其他）。PL 頁 FY2025A 欄沿用年報口徑，機器人列空白。")
    L.append("- 分部利益合計＝營業利益（公司無未分攤項目）。命題子列：線性滑軌：人形（微型滑軌與小螺桿）、滾珠螺桿：人形（PRS）。PRS 歸滾珠螺桿分部為本模型假設（公司未說明歸屬）。")
    L.append("")
    L.append("## 主要假設（In 頁；藍字可改）")
    L.append("- FY2026 錨點：9M26 實績＋4Q26＝3Q26 ×（1＋4%）（10 月營收 2,850 百萬元年增 35.7%），4Q 分部利益率＝3Q26。")
    L.append("- 核心成長（基準路徑）：FY2027 線性滑軌、滾珠螺桿 +12%、機器人 +18%、其他 0；FY2028 +6%／+6%／+12%／0；FY2029–2032 線性收斂到長期 3%（機器人 6%）。公司稱 2027 年可能雙位數成長、接單能見度到 2027 年中（Interested-party）。")
    L.append("- 分部利益率：9M26 實績持平（線性滑軌 14.7%、滾珠螺桿 19.0%、機器人 20.6%、其他 −12.5%）。最近四季口徑（含 4Q25）因分部重編未揭露，以最近三季代替。")
    L.append("- 情境：A＝FY2027／28 核心成長 +6／+4pt、利益率回到 3Q26 單季水準（＋3.7pt）、人形 2030 年 100 萬台；B＝基準、人形 30 萬台；C＝FY2027／28 −10／−8pt、利益率回到 FY2025 水準（−7.4pt）、人形 10 萬台。權重 25／50／25。")
    L.append("- 人形：2026 全球出貨 5 萬台（TrendForce）、BOM 2027 年 3.5 萬美元 → 2032 年 1.7 萬美元（BofA）、PRS 占 BOM 25%（19–52%）、非中國供應鏈 40%、上銀 PRS 份額穩態 5%（2–10%）、微型滑軌占 BOM 2%、份額 10%；爬坡 2027 25% → 2030 100%。")
    L.append("- DCF：rf 1.98%（台灣 10 年期）、ERP 5%＋CRP 0.78%、β 為上銀 5 年 β 與線性傳動同業中位數各自 Blume 調整後中位數、kd＝9M26 付息年化 ÷ 平均借款、g＝MIN(3%, rf)、RONIC＝WACC。Comparable：FY2028 分部淨利 × 同業前瞻 P/E（線性傳動 6 家中位數、機器人 3 家中位數），rel_k＝1.00（未調整：缺 10 年相對倍數）。")
    L.append("")
    L.append("## 限制")
    L.append("- 同業倍數取自資料商統計頁，各檔資料日期不同（2026-08-24～11-13），多數與股價基準日相差逾 2 個交易日（Comps 頁標「逾期」）；會計口徑未查核。")
    L.append("- rf、CRP、匯率沿用其他台股模型 2026-11-04／05 取值（逾期，跨模型一致）。")
    L.append("- 人形出貨 2030／2032、上銀份額、PRS 占比皆為 Assumed；線 2b 無出貨預測，PRS 占比兩來源互相矛盾。")
    L.append("- 分部利益率基準用 9M26（三季）而非最近四季；3Q26 含一次性美國子公司關稅退款（約 +0.6pt 毛利率）與人民幣升值（約 +0.5pt）。")
    L.append("- HTML 一頁摘要與 display_check.json 本版未做。")
    L.append("")
    L.append("## 未取得資料")
    L.append("- FY2025 新口徑分部（機器人單列）與 FY2025 分部利益；4Q25 新口徑分部。")
    L.append("- 上銀 PRS 客戶、報價、產能、份額；美國人形大廠小型滑軌訂單金額；PRS 營收歸屬分部。")
    L.append("- 2026-09-30 權益法投資帳面（用 2025-12-31）；庫藏股；FY2025 預付設備款增加。")
    L.append("- 上銀對同業 10 年前瞻 P/E 比值；同業淨負債 ÷ EBITDA、資本強度。")
    L.append("- 人形 PRS 實際利益率（子列沿用母分部利益率）。")
    L.append("")
    L.append("## 對帳（線層級變數）")
    L.append("見 xlsx「對帳」頁：人形出貨（2b）、每台 BOM（2b）、PRS 占 BOM（2b）、每台 PRS 支數（2b）、非中國供應鏈比例（2b）、BOM 學習曲線（2b；本模型以年降幅表示，與線 2b 的累積產量倍增口徑不同，照寫待對帳）、IFR 工業機器人裝機與 JARA 受注（2a，只作依據）。")
    L.append("")
    L.append("## 不適用的規格項目")
    L.append("- 不讀 Siliconomics／Tokenomics／Wattonomics 協調檔（實體 AI 線）；自用 AI 讀法、MW 口徑、融資瀑布與股權吸收（原則第 7 節、第 10 節第 9、52–55、62 條）不適用。")
    L.append("")
    L.append("## 驗證")
    L.append(f"- LibreOffice 重算後錯誤值 0 格。核對程式：")
    for c in chk: L.append(f"  - {c}")
    L.append("- 尚未由未參與製作的代理獨立核對數字（Project 規則第 10 條）。")
    L.append("")
    L.append("## 檔案")
    L.append(f"- `build_hiwin.py`：建置、整本重算敏感度與反解、核對、產生本 README；`sens.json`：快照；`segments.json`：官方分部宣告。")
    L.append(f"- 成品：`dist/{OUT}`")
    open(os.path.join(HERE, "README.md"), "w", encoding="utf-8").write("\n".join(L) + "\n")


if __name__ == "__main__":
    outp, cells = main()
    ok, chk = run_checks(outp)
    for c in chk: print(c)
    write_readme(outp, chk)
    print("成品：", outp, "核對", "全部通過" if ok else "有失敗")
    sys.exit(0 if ok else 1)
