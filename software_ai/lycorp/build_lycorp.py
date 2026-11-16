#!/usr/bin/env python3
"""LY Corporation（4689）個股模型 v0.1.1：命題（原型 A 按廣告計價的單位經濟）＋官方營收分部 → 損益 → DCF／分部加總 → 目標價。

用法：
  python3 software_ai/lycorp/build_lycorp.py <raw.xlsx> [--sens sens.json] [--set key=value ...]
  （--set 只改 In 頁基準欄 E，供 sens_lycorp.py 逐案整本重算用）
本版未做先驗（工作單 G 節）。金額＝十億日圓，除非另註。FY＝4 月制（FY2026＝2026-04～2027-03）。
所有歷史數字取自 data/facts_lycorp.csv 的事實列（Q 頁複寫），來源與標記逐列寫在 Q 頁。
"""
import datetime, hashlib, json, os, re, sys
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter as L

HERE = os.path.dirname(os.path.abspath(__file__))
VER = "0.1.1"
PXDATE = datetime.date(2026, 11, 16)
VALDATE = datetime.date(2026, 11, 16)
DATADATE = datetime.date(2026, 11, 16)
BUILDDATE = datetime.date(2026, 11, 16)
TGT = datetime.date(2027, 11, 16)
FYE = datetime.date(2027, 3, 31)

args = sys.argv[1:]
OUT = args[0]
SENS = json.load(open(args[args.index("--sens") + 1], encoding="utf-8")) if "--sens" in args else None
OVR = {}
for i, a in enumerate(args):
    if a == "--set":
        k, v = args[i + 1].split("=", 1)
        OVR[k] = float(v)

BLUE = Font(color="0000FF")
BOLD = Font(bold=True)
YEL = PatternFill("solid", fgColor="FFF2CC")
HDR = PatternFill("solid", fgColor="D9EAD3")
SNAP = json.load(open(os.path.join(HERE, "data", "tokenomics_snapshot_v22.0.json"), encoding="utf-8"))
TKV = SNAP["source"]["version"]
tk_luna = SNAP["items"]["IF_PriceRef_Luna"]["values"]
tk_sol = SNAP["items"]["IF_PriceRef_Sol"]["values"]
tk_astra = SNAP["items"]["IF_PriceRef_Astra"]["values"]
tk_fc_luna = SNAP["items"]["IF_FullCost_Luna"]["values"]["VR200 NVL72"]["基準"]
tk_taskmult = SNAP["items"]["IF_TaskTokMult"]["values"]

# 同業倍數（stockanalysis.com 統計頁，EV／EBITDA TTM；資料日與擷取日 2026-11-16）
PEERS = [  # (分部, 公司, 代號, 會計準則, 租賃處理, EV 十億本幣, TTM EBITDA 十億本幣, 倍數, 資料日, 幣別, 來源路徑)
    ("Media", "CyberAgent", "4751", "未取得", "未取得", 645.62, 98.28, 6.59, datetime.date(2026, 10, 22), "JPY", "tyo/4751"),
    ("Media", "Recruit Holdings", "6098", "IFRS（公開常識，未於資料頁核對）", "未取得", 25870.0, 800.86, 31.19, datetime.date(2026, 10, 23), "JPY", "tyo/6098"),
    ("Media", "NAVER", "035420", "K-IFRS（公開常識，未於資料頁核對）", "未取得", 24150.0, 3020.0, 7.05, datetime.date(2026, 10, 23), "KRW", "krx/035420"),
    ("Media", "Meta Platforms", "META", "US GAAP（公開常識，未於資料頁核對）", "未取得", 1670.0, 116.19, 14.40, datetime.date(2026, 11, 13), "USD", "stocks/meta"),
    ("Commerce", "Rakuten Group", "4755", "IFRS（公開常識，未於資料頁核對）", "未取得", 1900.0, 451.70, 4.24, datetime.date(2026, 9, 11), "JPY", "tyo/4755"),
    ("Commerce", "ZOZO", "3092", "未取得", "未取得", 1010.0, 79.92, 12.62, datetime.date(2026, 11, 13), "JPY", "tyo/3092"),
    ("Commerce", "Mercari", "4385", "未取得", "未取得", 853.83, 46.78, 18.25, datetime.date(2026, 8, 28), "JPY", "tyo/4385"),
    ("Commerce", "MonotaRO", "3064", "未取得", "未取得", 1160.0, 57.92, 20.07, datetime.date(2026, 10, 16), "JPY", "tyo/3064"),
    ("Strategic", "Rakuten Group", "4755", "IFRS（公開常識，未於資料頁核對）", "未取得", 1900.0, 451.70, 4.24, datetime.date(2026, 9, 11), "JPY", "tyo/4755"),
    ("Strategic", "PayPal", "PYPL", "US GAAP（公開常識，未於資料頁核對）", "未取得", 53.21, 6.52, 8.16, datetime.date(2026, 11, 13), "USD", "stocks/pypl"),
    ("Strategic", "Block", "XYZ", "US GAAP（公開常識，未於資料頁核對）", "未取得", 45.50, 1.55, 29.39, datetime.date(2026, 10, 26), "USD", "stocks/xyz"),
]


def _quart(v, q):
    v = sorted(v); n = len(v); pos = (n - 1) * q; lo = int(pos); fr = pos - lo
    return v[lo] + (v[min(lo + 1, n - 1)] - v[lo]) * fr


MQ = {}
for g_ in ("Media", "Commerce", "Strategic"):
    v_ = [x[7] for x in PEERS if x[0] == g_]
    MQ[g_] = (round(_quart(v_, 0.25), 2), round(_quart(v_, 0.75), 2))

wb = openpyxl.Workbook()
wb.remove(wb.active)
S = {n: wb.create_sheet(n) for n in ["README", "In", "Q", "PL", "DCF", "Val", "Peers", "RevDCF", "Recon", "Sens", "Sources"]}

YRS = ["FY2023A", "FY2024A", "FY2025A", "FY2026E", "FY2027E", "FY2028E", "FY2029E", "FY2030E", "FY2031E", "FY2032E"]
C0 = 3  # 欄 C＝FY2023
CL = [L(C0 + i) for i in range(10)]  # C..L
ACT = CL[:3]
FC = CL[3:]


def put(ws, ref, v, font=None, fill=None, fmt=None, wrap=False):
    c = ws[ref]
    c.value = v
    if font: c.font = font
    if fill: c.fill = fill
    if fmt: c.number_format = fmt
    if wrap: c.alignment = Alignment(wrap_text=True, vertical="top")
    return c


# =====================================================================================
# In 頁
# =====================================================================================
IN = S["In"]
INR = {}   # key -> row
IN_ROWS = []  # (key,label,unit,lo,base,hi,tag,src,note)
SEC = []


def sec(title):
    IN_ROWS.append(("#", title))


def inp(key, label, unit, lo, base, hi, tag, src, note="", fmt=None):
    IN_ROWS.append((key, label, unit, lo, base, hi, tag, src, note, fmt))


F25 = "https://data.swcms.net/file/lycorp-ir/dam/jcr:977c1536-eca8-4f95-b8ed-5b6df7c8e884/140120260507518939.pdf"
T2Q = "https://www.lycorp.co.jp/en/ir/news/manual_20261105153006/main/0/link/pdfFile.pdf"
S2Q = "https://www.lycorp.co.jp/en/ir/library/presentations/main/017/teaserItems2/01/linkList/04/link/en2026q2_presentation.pdf"
A2Q = "https://www.lycorp.co.jp/en/ir/library/presentations/main/017/teaserItems2/01/linkList/05/link/en2026q2_presentation_app.pdf"
SNAPSRC = "data/tokenomics_snapshot_v22.0.json（Tokenomics v22.0，import_tokenomics.py --check 通過）"
UE = "v0.2 改讀 unit_economics（該模型 v0.1 完成前於此自建）"

sec("A. 日期、現價、股數與資產負債")
inp("data_date", "資料截止日（最新財報＝FY2026 Q2，2026-11-05 發布）", "日期", None, DATADATE, None, "Derived", "facts_lycorp.csv 擷取日")
inp("px_date", "股價基準日（現價日）", "日期", None, PXDATE, None, "Verified", "traders.co.jp 2026-11-16 15:30（15 分鐘延遲）")
inp("rf_date", "無風險利率基準日", "日期", None, PXDATE, None, "Verified", "tradingeconomics.com 頁面日期")
inp("cons_date", "共識目標價基準日", "日期", None, datetime.date(2026, 9, 8), None, "Verified", "kabuyoho.jp 頁面日期")
inp("val_date", "評價日", "日期", None, VALDATE, None, "Derived", "同股價基準日（資產負債表為 2026-09-30，見下）")
inp("tp_date", "目標時點＝EDATE(股價基準日, 12)", "日期", None, "=EDATE(E6,12)", None, "Derived", "原則第 10 節第 26 條")
inp("fye", "FY2026 會計年度結束日", "日期", None, FYE, None, "Verified", "公司 4 月制")
inp("px", "現價（2026-11-16 收盤）", "日圓", None, 543.3, None, "Verified", "traders.co.jp 個股頁 historical（stocks/61_4689；15 分鐘延遲）", "第二來源見下列 px2，兩源一致")
inp("px2", "現價第二來源（2026-11-16 收盤）", "日圓", None, 543.3, None, "Verified", "Yahoo!ファイナンス 4689.T chart 頁：11/16 543.3、前一交易日 11/13 534.7", "與 traders.co.jp 的 543.3、534.7 完全一致；Verified")
inp("sh_out", "發行股數（含庫藏，法定申報封面，2026-09-30）", "百萬股", None, 6888.742311, None, "Verified", T2Q + " 封面（第一來源）", "與 Yahoo!ファイナンス 2026-10-26 的 6,888,742,311 股一致")
inp("sh_tsy", "庫藏股（2026-09-30，封面；含下列三個股權交付信託持股）", "百萬股", None, 25.505766, None, "Verified", T2Q + " 封面", "")
inp("sh_trust", "其中：股票交付信託持股（J-ESOP、董事激勵計畫、Stock Delivery ESOP；視同待交付的 RSU 類，計入稀釋）", "百萬股", None, 20.089266, None, "Verified", T2Q + " 封面", "封面註記三信託合計 20,089,266 股")
inp("rsu_forf", "信託股預期沒收率", "%", None, 0.0, 0.2, "Assumed", "未取得：公司未揭露沒收率；基準 0（全數交付，偏保守）", "", "0%")
inp("opt_n", "選擇權、可轉債、ASR 稀釋股數（庫藏股法）", "百萬股", None, 0.0, None, "Assumed", "未取得：決算短信封面與附註未列選擇權、可轉債、ASR 數量，有價證券報告書未逐頁查；0 不代表確認為零", "")
inp("sh_dil", "評價日稀釋股數＝發行 − 庫藏 ＋ 信託股 ×（1 − 沒收率）＋ 選擇權庫藏股法", "百萬股", None, "=@@sh_out@@-@@sh_tsy@@+@@sh_trust@@*(1-@@rsu_forf@@)+@@opt_n@@", None, "Derived", "DCF 與 Comparable 規格第 3 節", "", "#,##0.000")
inp("debt0", "有息負債（不含金融事業，2026-09-30）", "十億日圓", None, 1384.7, None, "Verified", A2Q + " p.10", "口徑：債務本金＋租賃負債（IFRS 16，營業與融資租賃不分，全數計入），與 FCFF 的資本支出含使用權資產取得、折舊含使用權資產折舊一致；租賃負債金額未單獨取得；PayPay 銀行存款與放款不在內")
inp("nd0", "淨有息負債（不含金融事業，2026-09-30）", "十億日圓", None, 777.1, None, "Verified", A2Q + " p.11", "公司圖表；現金為 debt0 − nd0（Derived）")
inp("nci0", "非控制權益（帳面，2026-09-30）", "十億日圓", None, 715.636, None, "Derived", T2Q + " p.10：權益 3,763,781 − 母公司持分 3,048,145（百萬日圓）", "以帳面計；市值口徑（PayPay 少數股權）只作敏感度（nci_mult）")
inp("nci_mult", "非控制權益估值倍數（帳面＝1）", "倍", 1.0, 1.5, 2.5, "Assumed", "PayPay 持分公允價值 712.6bn 對權益比 31.0%（FY2025 Q4 附錄 p.8），少數股權市值高於帳面；無獨立估值", "預設 1.5 倍為判斷")

sec("B. 營收：Media 搜尋廣告（命題模組：查詢 × 每查詢收益）")
inp("q0", "Yahoo! JAPAN 搜尋查詢量基期（FY2025，十億次／年）", "十億次", 40, 70, 120, "Assumed", "公司未揭露查詢量（未找到）；以每查詢收益換算，只影響推論成本，不影響營收基準", UE)
inp("qg26", "搜尋基礎成長：FY2026（無 AI 影響的有機查詢×價格）", "%", -0.06, -0.02, 0.02, "Assumed", "FY2026 上半年搜尋廣告 YoY＝Q1 −6.8%、Q2 −1.7%（簡報；Derived 上半年約 −4.3%），含部分 AI 影響", "", "0.0%")
inp("qg32", "搜尋基礎成長：FY2032（線性收斂）", "%", -0.03, 0.0, 0.03, "Assumed", "無證據，收斂到零成長", "", "0.0%")
inp("div_A", "情境 A：FY2032 搜尋查詢分流到聊天介面的占比", "%", None, 0.10, None, "Assumed", "無證據；A＝分流小（LINE／Yahoo! 留住查詢）", "", "0%")
inp("div_B", "情境 B：FY2032 搜尋查詢分流占比", "%", None, 0.20, None, "Assumed", "無證據；基準", "", "0%")
inp("div_C", "情境 C：FY2032 搜尋查詢分流占比", "%", None, 0.35, None, "Assumed", "無證據；C＝聊天介面（含 ChatGPT Apps）大量接走查詢", "", "0%")
inp("lift_A", "情境 A：FY2032 AI 轉換率提升帶動的每次點擊價格提升", "%", None, 0.15, None, "Assumed", "公司自揭 AI 功能使用者購前轉換 1.7 倍（非隨機對照，Interested-party）只作方向", "", "0%")
inp("lift_B", "情境 B：FY2032 每次點擊價格提升", "%", None, 0.08, None, "Assumed", "同上", "", "0%")
inp("lift_C", "情境 C：FY2032 每次點擊價格提升", "%", None, 0.02, None, "Assumed", "同上", "", "0%")

sec("C. 營收：Media 其餘線、Commerce、Strategic")
inp("ag26", "帳號廣告基礎成長：FY2026", "%", 0.06, 0.12, 0.16, "Assumed", "FY2026 上半年帳號廣告 +12.5%（決算短信 FY2026 Q2 p.5）", "", "0.0%")
inp("ag32", "帳號廣告基礎成長：FY2032", "%", 0.0, 0.04, 0.08, "Assumed", "線性收斂；無證據", "", "0.0%")
inp("aim_A", "情境 A：FY2032 LINE 官方帳號 AI Mode／Agent i BIZ 帶動帳號廣告增量", "%", None, 0.12, None, "Assumed", "公司預告 AI Mode 2026 年內上線（未量化）", "", "0%")
inp("aim_B", "情境 B：同上", "%", None, 0.06, None, "Assumed", "同上", "", "0%")
inp("aim_C", "情境 C：同上", "%", None, 0.0, None, "Assumed", "同上", "", "0%")
inp("dg26", "展示廣告成長：FY2026", "%", -0.09, -0.05, 0.0, "Assumed", "FY2026 上半年展示約 −5.4%（Derived：Q1 −2.5%、Q2 −8.2%）", "", "0.0%")
inp("dg32", "展示廣告成長：FY2032", "%", -0.03, 0.0, 0.03, "Assumed", "線性收斂", "", "0.0%")
inp("og26", "Media 其他（含用戶付費）成長：FY2026", "%", 0.04, 0.10, 0.14, "Assumed", "用戶付費等 H1 +13.0%、LYP Premium 直接訂閱 6.37→10.90 百萬", "", "0.0%")
inp("og32", "Media 其他成長：FY2032", "%", 0.0, 0.03, 0.06, "Assumed", "線性收斂", "", "0.0%")
inp("cg26", "Commerce 成長：FY2026", "%", 0.08, 0.14, 0.20, "Assumed", "H1 +11.8%（ASKUL 網路攻擊使 FY2025 下半年基期偏低）；公司指引 +20.1% 只在 Recon 對照", "", "0.0%")
inp("cg32", "Commerce 成長：FY2032", "%", 0.0, 0.02, 0.05, "Assumed", "線性收斂", "", "0.0%")
inp("sg26", "Strategic 成長：FY2026", "%", 0.20, 0.28, 0.34, "Assumed", "H1 +31.4%；PayPay 營收 H1 YoY 約 +27%", "", "0.0%")
inp("sg32", "Strategic 成長：FY2032", "%", 0.0, 0.02, 0.06, "Assumed", "線性收斂", "", "0.0%")
inp("oth0", "Other 分部（FY2026 起每年，十億日圓）", "十億日圓", 5, 9.9, 12, "Derived", "FY2025 實績 9.861（決算短信 p.32）；H1 FY2026 3.8", "", "0.0")
inp("adj0", "Adjustments（分部間消除，每年）", "十億日圓", -14, -12.1, -10, "Derived", "FY2025 實績 −12.086", "", "0.0")

sec("D. AI 新營收：Agent i（使用者 × 付費占比 × 月費；使用者 × 廣告 ARPU）")
inp("ag_u0", "Agent i「日活」基期（2026-06，跨服務累計加總；百萬）", "百萬人", 8, 12.0, 14, "Interested-party", "https://www.lycorp.co.jp/en/ir/library/presentations/main/017/teaserItems2/00/linkList/04/link/en2026q1_presentation.pdf p.7", "公司定義寬鬆（累計加總）；來源誘因：管理層向市場展示 AI 採用率")
inp("ag_ug", "Agent i 使用者年增（FY2027 起）", "%", 0.0, 0.20, 0.40, "Assumed", "無證據", "", "0%")
inp("agp_A", "情境 A：FY2032 付費占 Agent i 使用者比例", "%", None, 0.06, None, "Assumed", "公司計畫高 token 用量者付費方案（價格未公布）", "", "0%")
inp("agp_B", "情境 B：同上", "%", None, 0.03, None, "Assumed", "同上", "", "0%")
inp("agp_C", "情境 C：同上", "%", None, 0.01, None, "Assumed", "同上", "", "0%")
inp("ag_price", "Agent i 付費方案月費（日圓／月）", "日圓", 500, 800, 1500, "Assumed", "價格未公布（LYP Premium 650、Lite 290 為對照）", "", "0")
inp("ag_adarpu", "Agent i 廣告 ARPU：FY2032 年化（日圓／使用者／年）", "日圓", 200, 500, 1000, "Assumed", "Agent i 展示廣告 2026-07 起部分計費，無金額揭露", "", "0")

sec("E. 推論成本（只讀 Tokenomics 快照；不自行估 token 價格）")
inp("tok_chat", "一般聊天每次互動 token 數", "token", 1000, 2000, 4000, "Assumed", "未取得：公司未揭露 token 量或代理任務量（Agent i 只揭露「日活」與購前轉換倍數），無法由揭露校準每任務 token；此為 Assumed，不是揭露值", "", "0")
inp("tok_mult", "任務倍數（IF_TaskTokMult 第 2 個任務類別）", "倍", 1.0, tk_taskmult[1], 9.2, "Assumed", f"Tokenomics {TKV} IF_TaskTokMult＝{[round(x,2) for x in tk_taskmult]}；對應哪個任務類別待 AI供應鏈 確認，低＝一般聊天、高＝第 3 類", "", "0.00")
inp("ptok", "向模型商購買的市場價（US$／百萬 token，參考請求混合）", "US$/M", tk_fc_luna, tk_luna, tk_sol, "Analogy", f"Tokenomics {TKV}：基準＝IF_PriceRef_Luna {tk_luna:.4f}；高＝IF_PriceRef_Sol {tk_sol:.3f}；低＝IF_FullCost_Luna（VR200 基準）{tk_fc_luna:.4f} 為模型商自己做的全成本下限，不是市場價；採用哪一級模型為 Assumed", "待 AI供應鏈 登錄軟體公司採購價名稱", "0.0000")
inp("route_cheap", "路由到便宜模型（IF_PriceRef_Luna 級）的 token 占比；只在 ptok 高於 Luna 時有效", "%", None, 0.0, 0.7, "Assumed", "模型組合假設：Agent i 以任務路由降低成本；無公司揭露", "", "0%")
inp("ptok_eff", "有效 token 價格＝ptok ×（1 − 路由占比）＋ Luna 價 × 路由占比", "US$/M", None, f"=@@ptok@@*(1-@@route_cheap@@)+{tk_luna}*@@route_cheap@@", None, "Derived", f"Tokenomics {TKV} IF_PriceRef_Luna", "", "0.0000")
inp("fx", "日圓兌美元", "日圓/US$", 140, 155, 170, "Assumed", "未查證現匯；僅換算推論成本", "", "0")
inp("ai_inter", "Agent i 每使用者每月互動次數", "次", 15, 30, 60, "Assumed", "無公司揭露", "", "0")
inp("ai_share", "FY2032 LINE／Yahoo! 自家搜尋由 AI 回答的比例", "%", 0.1, 0.35, 0.6, "Assumed", "公司未揭露；Agent i 併入 Yahoo! 搜尋框", "", "0%")

sec("F. 損益與現金流")
inp("ttm_oi", "最近四季營業利益（FY2025 Q3＋Q4＋FY2026 Q1＋Q2，IFRS）", "十億日圓", None, 325.658, None, "Derived", "FY2025 341.322 ＋ H1 FY2026 198.862 − H1 FY2025 214.526（T2Q26 p.11、T25 p.17）", "H1 FY2025 含 FY2025 一次性利益 61.4 十億日圓，TTM 已不含")
inp("ttm_rev", "最近四季營收", "十億日圓", None, 2158.667, None, "Derived", "FY2025 2,036.366 ＋ H1 FY2026 1,117.668 − H1 FY2025 995.367", "")
inp("ttm_cogs", "最近四季營業成本", "十億日圓", None, 564.099, None, "Derived", "FY2025 530.095 ＋ 305.372 − 271.368（T2Q26 p.11）", "")
inp("opm26", "營業利益率（無 AI 營收基礎）：FY2026（基準＝最近四季實績）", "%", 0.13, "=@@ttm_oi@@/@@ttm_rev@@", 0.19, "Derived", "最近四季營業利益 ÷ 營收；沒有改善證據，基準全期持平（G2 第 3 條）", "低／高只放 Sens（收斂情境）", "0.0%")
inp("opm32", "營業利益率（無 AI 營收基礎）：FY2032（基準＝持平於 FY2026）", "%", 0.13, "=@@opm26@@", 0.19, "Derived", "同上；低／高只放 Sens", "", "0.0%")
inp("m_inc", "AI 增量營收的流入營業利益比（邊際）", "%", 0.50, 0.75, 0.90, "Assumed", "廣告增量邊際利益高；無揭露", "", "0%")
inp("m_lost", "AI 侵蝕（損失營收）的流出營業利益比（邊際）", "%", 0.70, 0.85, 0.95, "Assumed", "損失的搜尋廣告固定成本不減；無揭露", "", "0%")
inp("cogs_pct", "營業成本占營收（不含推論成本；基準＝最近四季）", "%", 0.24, "=@@ttm_cogs@@/@@ttm_rev@@", 0.28, "Derived", "最近四季營業成本 ÷ 營收；基準持平（G2 第 3 條）", "", "0.0%")
inp("nonop", "營業外淨額（稅前利益 − 營業利益，每年）", "十億日圓", -50, -40, -25, "Derived", "FY2024 −40.2、FY2025 −47.1（含減損 19.6、出售益 10.8）", "", "0.0")
inp("tax", "稅率", "%", 0.25, 0.28, 0.31, "Assumed", "FY2023 23.4%、FY2024 26.4%、FY2025 3.8%（一次性）；法定約 30%", "", "0%")
inp("nci_sh", "非控制權益占稅後淨利比例", "%", 0.22, 0.30, 0.36, "Derived", "FY2023 18.6%、FY2024 24.2%、FY2025 31.6%（PayPay 少數股權）", "", "0%")
inp("da_pct", "折舊攤銷占營收", "%", 0.07, 0.087, 0.10, "Derived", "FY2025 176.4／2,036.4＝8.7%", "", "0.0%")
inp("cx_pct", "資本支出占營收（含使用權資產取得）", "%", 0.07, 0.087, 0.10, "Derived", "FY2025 177.6（P&E／無形 125.9＋使用權 51.7）／2,036.4＝8.7%（簡報 FY2025 Q4 附錄 p.10）", "", "0.0%")
inp("wc_k", "營運資金增加占營收增量", "%", 0.0, 0.05, 0.10, "Assumed", "含金融事業，無獨立揭露；PayPay 銀行資產負債不在 FCFF", "", "0%")
inp("dps0", "每股股利 FY2026（日圓）", "日圓", None, 12.0, None, "Interested-party", T2Q + " p.1（公司指引）", "來源誘因：管理層指引")
inp("dps_g", "每股股利年增（FY2027 起）", "%", 0.0, 0.05, 0.10, "Assumed", "公司股利政策改為與獲利連動（FY2025 Q4 簡報 slide 31-33）", "", "0%")

sec("G. 折現率與估值規則")
inp("rf", "無風險利率（日本 10 年期國債，2026-11-16）", "%", None, 0.0289, None, "Verified", "https://tradingeconomics.com/japan/government-bond-yield（2.888%）", "", "0.00%")
inp("beta", "β（未取兩來源 Blume 調整；Assumed）", "倍", 0.7, 0.9, 1.1, "Assumed", "未找到資料商 β", "", "0.00")
inp("erp", "股權風險溢酬", "%", 0.04, 0.05, 0.06, "Assumed", "DCF 標準規格預設 5%（敏感度 4–6%）", "", "0.0%")
inp("kd", "稅前負債成本", "%", 0.012, 0.018, 0.025, "Assumed", "未逐檔查證公司債與借款利率", "", "0.0%")
inp("g_tv", "永續成長率", "%", 0.01, 0.02, 0.03, "Assumed", "標準規格預設 3%；日本名目成長與 JGB 10 年 2.89% 下改用 2%（偏離標準，見 README）", "", "0.0%")
inp("m_media", "分部加總：Media EV／TTM EBITDA 倍數（同業中位數）", "倍", MQ["Media"][0], "=Peers!$H$17", MQ["Media"][1], "Analogy", "Peers 頁 4 家（CyberAgent、Recruit、NAVER、Meta）中位數；低／高＝第 1／第 3 四分位", "同業資料日不一，逾期列見 Peers 頁", "0.0")
inp("m_com", "分部加總：Commerce EV／TTM EBITDA 倍數（同業中位數）", "倍", MQ["Commerce"][0], "=Peers!$H$18", MQ["Commerce"][1], "Analogy", "Peers 頁 4 家（Rakuten、ZOZO、Mercari、MonotaRO）中位數；ASKUL 因 TTM EBITDA 為負排除", "同業為純電商或集團，成長與利益率不同於 LY Commerce（含 ASKUL 物流）", "0.0")
inp("m_str", "分部加總：Strategic（PayPay 等）EV／TTM EBITDA 倍數（同業中位數）", "倍", MQ["Strategic"][0], "=Peers!$H$19", MQ["Strategic"][1], "Analogy", "Peers 頁 3 家（Rakuten、PayPal、Block）中位數", "支付同業；PayPay 銀行與信用業務不同", "0.0")
inp("em_media", "分部調整後 EBITDA 率：Media（FY2025）", "%", None, 0.382, None, "Interested-party", "簡報 FY2025 Q4 slide 8：280.6／735.1", "公司 non-IFRS；來源誘因：管理層自訂口徑", "0.0%")
inp("em_com", "分部調整後 EBITDA 率：Commerce（FY2025）", "%", None, 0.152, None, "Interested-party", "129.9／857.6", "同上", "0.0%")
inp("em_str", "分部調整後 EBITDA 率：Strategic（FY2025）", "%", None, 0.211, None, "Interested-party", "93.9／445.7", "同上", "0.0%")
inp("w_dcf", "加權目標價：DCF 權重", "%", 0.3, 0.5, 0.7, "Assumed", "預設 DCF 50%＋分部加總 50%（個股模型原則第 4 節）", "", "0%")
inp("tpw_A", "情境權重 A（樂觀）", "%", None, 0.25, None, "Assumed", "DCF 標準規格第 3 節預設 25／50／25", "", "0%")
inp("tpw_B", "情境權重 B（基準）", "%", None, 0.50, None, "Assumed", "同上", "", "0%")
inp("tpw_C", "情境權重 C（保守）", "%", None, 0.25, None, "Assumed", "同上", "", "0%")
inp("cons_tp", "市場共識平均目標價（日圓）", "日圓", None, 583, None, "Interested-party", "https://kabuyoho.jp/sp/reportTarget?bcode=4689（10 家，2026-09-08，已過期：早於 Q2 上修）", "來源誘因：賣方預估", "0")

inp("ly_ebitda", "本公司 TTM EBITDA（資料商口徑＝營業利益＋折舊攤銷，與同業倍數同源）", "十億日圓", None, 485.95, None, "Analogy", "stockanalysis.com/quote/tyo/4689/statistics（TTM EBITDA 485.95，資料日見下列 ly_eb_date）", "非公司調整後 EBITDA；僅用於現行倍數對照")
inp("ly_eb_date", "本公司 TTM EBITDA 資料日", "日期", None, datetime.date(2026, 10, 13), None, "Analogy", "同上（頁面顯示價格日 2026-10-13）")

# 敏感度分組（HTML 驅動頁用）
NAMES_SCN = {"div_A", "div_B", "div_C", "lift_A", "lift_B", "lift_C", "aim_A", "aim_B", "aim_C", "agp_A", "agp_B", "agp_C", "tpw_A", "tpw_B", "tpw_C"}

put(IN, "A1", "In — 輸入（藍字；基準欄 E 進主模型；D／F＝敏感度低／高，只用於 Sens 頁，不進目標價）。標記＝Verified／Interested-party／Analogy／Assumed／Derived。本版未做先驗。", BOLD)
put(IN, "A2", "事實來源：data/facts_lycorp.csv、data/research_lycorp.md；推論成本只讀 data/tokenomics_snapshot_v22.0.json（Tokenomics v22.0）。金額十億日圓，EPS 與股價日圓。FY＝4 月制，FY2026＝2026-04～2027-03。", None)
for j, h in enumerate(["鍵", "項目", "單位", "敏感度低", "基準", "敏感度高", "標記", "來源", "來源誘因／說明"]):
    put(IN, f"{L(j+1)}3", h, BOLD, HDR)
r = 4
for row in IN_ROWS:
    if row[0] == "#":
        put(IN, f"A{r}", row[1], BOLD)
        r += 1
        continue
    key, label, unit, lo, base, hi, tag, src, note, *fm = row
    fmt = fm[0] if fm else None
    INR[key] = r
    put(IN, f"A{r}", key)
    put(IN, f"B{r}", label)
    put(IN, f"C{r}", unit)
    if key in OVR:
        base = OVR[key]
    for col, v in (("D", lo), ("E", base), ("F", hi)):
        if v is None: continue
        c = put(IN, f"{col}{r}", v, BLUE if col == "E" and not (isinstance(v, str) and v.startswith("=")) else None, YEL if tag == "Assumed" and col == "E" else None)
        if isinstance(v, (datetime.date,)): c.number_format = "yyyy-mm-dd"
        elif fmt: c.number_format = fmt
    if key == "tp_date":
        IN[f"E{r}"].value = f"=EDATE(E{INR['px_date']},12)"
        IN[f"E{r}"].number_format = "yyyy-mm-dd"
    put(IN, f"G{r}", tag)
    put(IN, f"H{r}", src)
    put(IN, f"I{r}", note)
    r += 1
IN_LAST = r - 1
for _rw in range(4, IN_LAST + 1):
    for _col in "DEF":
        _v = IN[f"{_col}{_rw}"].value
        if isinstance(_v, str) and "@@" in _v:
            IN[f"{_col}{_rw}"].value = re.sub(r"@@(\w+)@@", lambda m_: f"E{INR[m_.group(1)]}", _v)
r += 1
put(IN, f"A{r}", "J. 情境驅動對照（三情境並排，進加權目標價；本節只引用上方各列，修改請改上方藍字）", BOLD)
r += 1
for j, h in enumerate(["鍵", "項目", "單位", "情境 A（樂觀）", "情境 B（基準）", "情境 C（保守）", "標記", "說明"]):
    put(IN, f"{L(j+1)}{r}", h, BOLD, HDR)
r += 1
SCN_ROWS = {}
for k, lab, unit, kA, kB, kC in (("scn_div", "FY2032 搜尋查詢分流到聊天介面占比", "%", "div_A", "div_B", "div_C"),
                                 ("scn_lift", "FY2032 AI 轉換率提升帶動的每次點擊價格提升", "%", "lift_A", "lift_B", "lift_C"),
                                 ("scn_aim", "FY2032 官方帳號 AI Mode 帶動帳號廣告增量", "%", "aim_A", "aim_B", "aim_C"),
                                 ("scn_agp", "FY2032 Agent i 付費占使用者比例", "%", "agp_A", "agp_B", "agp_C"),
                                 ("scn_w", "情境權重", "%", "tpw_A", "tpw_B", "tpw_C")):
    SCN_ROWS[k] = r
    put(IN, f"A{r}", k)
    put(IN, f"B{r}", lab)
    put(IN, f"C{r}", unit)
    for col, kk in zip("DEF", (kA, kB, kC)):
        put(IN, f"{col}{r}", f"=E{INR[kk]}", None, None, "0%")
    put(IN, f"G{r}", "Derived")
    put(IN, f"H{r}", f"引用 In 頁 {kA}／{kB}／{kC}；三情境只在此四個 AI 驅動分化（見 README『情境分化依據』）")
    r += 1
IN.column_dimensions["A"].width = 14
IN.column_dimensions["B"].width = 58
IN.column_dimensions["H"].width = 60
IN.column_dimensions["I"].width = 50


def I(key, absolute=True):
    return f"In!$E${INR[key]}"


# =====================================================================================
# Q 頁：事實總帳（歷史數字；PL 實績欄引用）
# =====================================================================================
Q = S["Q"]
QR = {}
put(Q, "A1", "Q — 歷史事實（決算短信、簡報；逐列來源。十億日圓＝百萬日圓 ÷ 1000；FY2021–FY2022 另見 data/facts_lycorp.csv）", BOLD)
put(Q, "A2", "分部口徑：FY2023 取 FY2024 決算短信重編口徑、FY2024 取 FY2025 決算短信重編口徑，FY2023 與 FY2024 間 Media／Commerce 因 2025-04／10 重編不完全可比（見研究包）。", None)
for j, h in enumerate(["鍵", "指標", "單位", "FY2023A", "FY2024A", "FY2025A", "標記", "來源", "來源誘因"]):
    put(Q, f"{L(j+1)}3", h, BOLD, HDR)
FACTS = [
    ("seg_media", "Media 營收", [702.145, 732.180, 735.197], "Verified", "決算短信 FY2024 Note 3（FY2023）、FY2025 p.32（FY2024、FY2025）"),
    ("seg_com", "Commerce 營收", [827.085, 848.485, 857.622], "Verified", "同上"),
    ("seg_str", "Strategic 營收", [289.989, 341.201, 445.772], "Verified", "同上"),
    ("seg_oth", "Other 營收", [8.052, 7.828, 9.861], "Verified", "同上"),
    ("seg_adj", "Adjustments（分部間消除）", [-12.609, -12.216, -12.086], "Verified", "同上"),
    ("rev", "合併營收", [1814.663, 1917.478, 2036.366], "Verified", "決算短信 FY2024 p.17、FY2025 p.17"),
    ("cogs", "營業成本", [520.449, 529.522, 530.095], "Verified", "決算短信 FY2024 p.17、FY2025 p.17"),
    ("oi", "營業利益（IFRS）", [208.191, 315.033, 341.322], "Verified", "同上"),
    ("pbt", "稅前利益", [181.621, 274.882, 294.231], "Verified", "決算短信 FY2023 p.17、FY2025 p.17"),
    ("ptot", "當期利益（含非控制權益）", [139.073, 202.403, 283.090], "Verified", "同上"),
    ("nio", "歸屬母公司淨利", [113.199, 153.465, 193.692], "Verified", "同上"),
    ("eps", "基本 EPS（日圓）", [15.10, 21.00, 27.97], "Verified", "同上"),
    ("da", "折舊攤銷", [None, 159.874, 176.445], "Verified", "決算短信 FY2025 p.21（FY2023 未取）"),
    ("srch", "搜尋廣告（Media）", [None, None, 179.0], "Interested-party", "簡報 FY2025 Q4 slide 9-10；公司分類，十億日圓一位小數"),
    ("acct", "帳號廣告（Media）", [None, None, 145.7], "Interested-party", "同上"),
    ("disp", "展示廣告（Media）", [None, None, 254.2], "Interested-party", "同上"),
    ("moth", "其他（Media，含用戶付費）", [None, None, 156.2], "Interested-party", "同上"),
    ("adrev", "廣告營收合併合計（Media＋Commerce 廣告）", [606.4, 638.0, 657.6], "Interested-party", "簡報 FY2025 Q4 附錄 p.15"),
    ("emp", "員工數（合併，期末，人）", [28196, 27003, 29863], "Verified", "有價證券報告書（第二來源整理頁 the-shashi.com）；FY2023 欄為 2024-03 值"),
]
r = 4
for key, lab, vals, tag, src in FACTS:
    QR[key] = r
    put(Q, f"A{r}", key); put(Q, f"B{r}", lab); put(Q, f"C{r}", "人" if key == "emp" else ("日圓" if key == "eps" else "十億日圓"))
    for j, v in enumerate(vals):
        if v is not None:
            put(Q, f"{L(4+j)}{r}", v, BLUE)
    put(Q, f"G{r}", tag); put(Q, f"H{r}", src + "（全文逐列見 data/facts_lycorp.csv）")
    if tag == "Interested-party":
        put(Q, f"I{r}", "公司自訂分類與非 IFRS 揭露；來源誘因：管理層對市場說明口徑")
    r += 1
Q.column_dimensions["B"].width = 44; Q.column_dimensions["H"].width = 80


def QQ(key, i):  # i: 0,1,2 -> D,E,F
    return f"Q!${L(4+i)}${QR[key]}"


# =====================================================================================
# PL 頁：三情境分部營收與損益；標準損益表（情境 B）
# =====================================================================================
PL = S["PL"]
put(PL, "A1", "PL — 營收分部推導與損益（十億日圓；三情境 A／B／C；FY2023A–FY2025A 實績，FY2026E–FY2032E 預測）", BOLD)
put(PL, "A2", "命題模組（原型 A，按廣告計價）：搜尋廣告＝基礎 ×（1 − 分流）×（1 + 價格提升）；AI 淨效果＝營收增量 − 營收侵蝕 − 推論成本。變數名稱與 unit_economics 對齊，該模型 v0.1 完成前在 In 頁自建（v0.2 改讀）。", None)
HR = 4
put(PL, f"A{HR}", "項目", BOLD, HDR); put(PL, f"B{HR}", "單位", BOLD, HDR)
for j, y in enumerate(YRS):
    put(PL, f"{CL[j]}{HR}", y, BOLD, HDR)
TR = 5
put(PL, f"A{TR}", "年度序號 t（FY2025＝0）"); put(PL, f"B{TR}", "序號")
RR = 6
put(PL, f"A{RR}", "AI 效果爬升比例 t／7（FY2032＝1）"); put(PL, f"B{RR}", "比例")
SR = 7
put(PL, f"A{SR}", "線性收斂比例（FY2026＝0，FY2032＝1）"); put(PL, f"B{SR}", "比例")
for j in range(10):
    put(PL, f"{CL[j]}{TR}", j - 2)
    put(PL, f"{CL[j]}{RR}", f"=MAX(0,{CL[j]}{TR})/7", None, None, "0.000")
    put(PL, f"{CL[j]}{SR}", f"=MAX(0,{CL[j]}{TR}-1)/6", None, None, "0.000")


def path(a, b, col):  # 線性收斂值：FY2026=a，FY2032=b
    return f"({I(a)}+({I(b)}-{I(a)})*{col}${SR})"


PR = {}  # scenario -> dict of rows
BLOCK_ROWS = ["s_base", "s_ero", "s_inc", "srch", "a_base", "a_inc", "acct", "disp", "moth", "ag_sub", "ag_ad", "media", "com", "str", "oth", "adj", "rev",
              "chk_media", "chk_seg", "rev_noai", "ai_inc", "ai_ero", "q_bn", "ai_q", "inf_ag", "inf_s", "inf", "ai_net", "ai_oi", "opm0", "oi", "cogs", "gp", "opx", "nonop", "pbt", "tax", "ni", "nci", "nio", "eps", "om", "ag_users"]
LABEL = {
    "s_base": ("搜尋廣告基礎營收（無 AI 影響）", "十億日圓"), "s_ero": ("AI 侵蝕：查詢分流到聊天介面的搜尋廣告損失", "十億日圓"),
    "s_inc": ("AI 增量：轉換率提升（反映於每次點擊價格）", "十億日圓"), "srch": ("搜尋廣告", "十億日圓"),
    "a_base": ("帳號廣告基礎營收", "十億日圓"), "a_inc": ("AI 增量：官方帳號 AI Mode", "十億日圓"), "acct": ("帳號廣告（LINE 官方帳號等）", "十億日圓"),
    "disp": ("展示廣告", "十億日圓"), "moth": ("其他（含用戶付費）", "十億日圓"),
    "ag_sub": ("AI 增量：Agent i 付費訂閱營收", "十億日圓"), "ag_ad": ("AI 增量：Agent i 廣告營收", "十億日圓"),
    "media": ("Media", "十億日圓"), "com": ("Commerce", "十億日圓"), "str": ("Strategic", "十億日圓"), "oth": ("Other", "十億日圓"), "adj": ("Adjustments", "十億日圓"),
    "rev": ("合併營收", "十億日圓"),
    "chk_media": ("檢查：Media 子線合計 − Media", "十億日圓"), "chk_seg": ("檢查：分部合計 − 合併營收", "十億日圓"),
    "rev_noai": ("無 AI 影響的合併營收", "十億日圓"), "ai_inc": ("AI 營收增量合計", "十億日圓"), "ai_ero": ("AI 營收侵蝕合計（負值）", "十億日圓"),
    "q_bn": ("自家搜尋查詢量（分流後）", "十億次"), "ai_q": ("其中由 AI 回答的查詢量", "十億次"),
    "inf_ag": ("推論成本：Agent i", "十億日圓"), "inf_s": ("推論成本：搜尋 AI 回答", "十億日圓"), "inf": ("推論成本合計（向模型商購買）", "十億日圓"),
    "ai_net": ("AI 淨效果（營收：增量 − 侵蝕 − 推論成本）", "十億日圓"), "ai_oi": ("AI 淨效果（營業利益：增量×邊際 − 侵蝕×邊際 − 推論成本）", "十億日圓"),
    "opm0": ("營業利益率（無 AI 營收基礎）", "%"), "oi": ("營業利益", "十億日圓"), "cogs": ("營業成本（含推論成本）", "十億日圓"), "gp": ("毛利", "十億日圓"),
    "opx": ("營業費用（殘差＝毛利 − 營業利益）", "十億日圓"), "nonop": ("營業外淨額", "十億日圓"), "pbt": ("稅前利益", "十億日圓"), "tax": ("所得稅", "十億日圓"),
    "ni": ("稅後淨利（含非控制權益）", "十億日圓"), "nci": ("非控制權益", "十億日圓"), "nio": ("歸屬母公司淨利", "十億日圓"), "eps": ("EPS（完全稀釋，日圓）", "日圓"),
    "om": ("營業利益率", "%"), "ag_users": ("Agent i 使用者", "百萬人"),
}
PCT = {"opm0", "om"}
row = 9
for s in "ABC":
    put(PL, f"A{row}", f"情境 {s}（{'樂觀' if s=='A' else '基準' if s=='B' else '保守'}）", BOLD, HDR)
    row += 1
    rows = {}
    for k in BLOCK_ROWS:
        rows[k] = row
        put(PL, f"A{row}", LABEL[k][0]); put(PL, f"B{row}", LABEL[k][1])
        row += 1
    PR[s] = rows
    row += 1
STD0 = row

scn = lambda s, key: f"In!$E${INR[key + '_' + s]}"  # div_A 等
for s in "ABC":
    R = PR[s]
    for j in range(10):
        c = CL[j]
        p = CL[j - 1] if j > 0 else None
        act = j < 3
        f = {}
        if act:
            qi = j  # Q 頁欄 D/E/F
            f["media"] = f"={QQ('seg_media', qi)}"
            f["com"] = f"={QQ('seg_com', qi)}"
            f["str"] = f"={QQ('seg_str', qi)}"
            f["oth"] = f"={QQ('seg_oth', qi)}"
            f["adj"] = f"={QQ('seg_adj', qi)}"
            f["rev"] = f"={QQ('rev', qi)}"
            if j == 2:
                f["srch"] = f"={QQ('srch', 2)}"; f["acct"] = f"={QQ('acct', 2)}"; f["disp"] = f"={QQ('disp', 2)}"; f["moth"] = f"={QQ('moth', 2)}"
                f["s_base"] = f"={c}{R['srch']}"; f["a_base"] = f"={c}{R['acct']}"
                f["chk_media"] = f"=IF({c}{R['srch']}=\"\",0,{c}{R['srch']}+{c}{R['acct']}+{c}{R['disp']}+{c}{R['moth']}-{c}{R['media']})"
                f["ag_users"] = 0
                f["q_bn"] = f"={I('q0')}"
            else:
                f["chk_media"] = 0
            f["chk_seg"] = f"={c}{R['media']}+{c}{R['com']}+{c}{R['str']}+{c}{R['oth']}+{c}{R['adj']}-{c}{R['rev']}"
            f["cogs"] = f"={QQ('cogs', qi)}"
            f["oi"] = f"={QQ('oi', qi)}"
            f["pbt"] = f"={QQ('pbt', qi)}"
            f["ni"] = f"={QQ('ptot', qi)}"
            f["tax"] = f"={c}{R['pbt']}-{c}{R['ni']}"
            f["nio"] = f"={QQ('nio', qi)}"
            f["nci"] = f"={c}{R['ni']}-{c}{R['nio']}"
            f["eps"] = f"={QQ('eps', qi)}"
            f["gp"] = f"={c}{R['rev']}-{c}{R['cogs']}"
            f["opx"] = f"={c}{R['gp']}-{c}{R['oi']}"
            f["nonop"] = f"={c}{R['pbt']}-{c}{R['oi']}"
            f["om"] = f"={c}{R['oi']}/{c}{R['rev']}"
        else:
            t = f"{c}${TR}"
            rr_ = f"{c}${RR}"
            f["s_base"] = f"={p}{R['s_base']}*(1+{path('qg26','qg32',c)})" if j > 3 else f"={p}{R['s_base']}*(1+{I('qg26')})"
            f["s_ero"] = f"=-{c}{R['s_base']}*{scn(s,'div')}*{rr_}"
            f["s_inc"] = f"={c}{R['s_base']}*(1-{scn(s,'div')}*{rr_})*{scn(s,'lift')}*{rr_}"
            f["srch"] = f"={c}{R['s_base']}+{c}{R['s_ero']}+{c}{R['s_inc']}"
            f["a_base"] = f"={p}{R['a_base']}*(1+{path('ag26','ag32',c)})" if j > 3 else f"={p}{R['a_base']}*(1+{I('ag26')})"
            f["a_inc"] = f"={c}{R['a_base']}*{scn(s,'aim')}*{rr_}"
            f["acct"] = f"={c}{R['a_base']}+{c}{R['a_inc']}"
            f["disp"] = f"={p}{R['disp']}*(1+{path('dg26','dg32',c)})" if j > 3 else f"={p}{R['disp']}*(1+{I('dg26')})"
            f["moth"] = f"={p}{R['moth']}*(1+{path('og26','og32',c)})" if j > 3 else f"={p}{R['moth']}*(1+{I('og26')})"
            f["ag_users"] = f"={I('ag_u0')}*(1+{I('ag_ug')})^MAX(0,{t}-1)"
            f["ag_sub"] = f"={c}{R['ag_users']}*{scn(s,'agp')}*{rr_}*{I('ag_price')}*12/1000"
            f["ag_ad"] = f"={c}{R['ag_users']}*{I('ag_adarpu')}*{rr_}/1000"
            f["media"] = f"={c}{R['srch']}+{c}{R['acct']}+{c}{R['disp']}+{c}{R['moth']}+{c}{R['ag_sub']}+{c}{R['ag_ad']}"
            f["com"] = f"={p}{R['com']}*(1+{path('cg26','cg32',c)})" if j > 3 else f"={p}{R['com']}*(1+{I('cg26')})"
            f["str"] = f"={p}{R['str']}*(1+{path('sg26','sg32',c)})" if j > 3 else f"={p}{R['str']}*(1+{I('sg26')})"
            f["oth"] = f"={I('oth0')}"
            f["adj"] = f"={I('adj0')}"
            f["ai_inc"] = f"={c}{R['s_inc']}+{c}{R['a_inc']}+{c}{R['ag_sub']}+{c}{R['ag_ad']}"
            f["ai_ero"] = f"={c}{R['s_ero']}"
            f["rev_noai"] = f"={c}{R['s_base']}+{c}{R['a_base']}+{c}{R['disp']}+{c}{R['moth']}+{c}{R['com']}+{c}{R['str']}+{c}{R['oth']}+{c}{R['adj']}"
            f["rev"] = f"={c}{R['rev_noai']}+{c}{R['ai_inc']}+{c}{R['ai_ero']}"
            f["chk_media"] = f"={c}{R['srch']}+{c}{R['acct']}+{c}{R['disp']}+{c}{R['moth']}+{c}{R['ag_sub']}+{c}{R['ag_ad']}-{c}{R['media']}"
            f["chk_seg"] = f"={c}{R['media']}+{c}{R['com']}+{c}{R['str']}+{c}{R['oth']}+{c}{R['adj']}-{c}{R['rev']}"
            f["q_bn"] = f"={I('q0')}*({c}{R['s_base']}/${CL[2]}{R['s_base']})*(1-{scn(s,'div')}*{rr_})"
            f["ai_q"] = f"={c}{R['q_bn']}*{I('ai_share')}*{rr_}"
            tokq = f"{I('tok_chat')}*{I('tok_mult')}"
            f["inf_ag"] = f"={c}{R['ag_users']}*{I('ai_inter')}*12*{tokq}*{I('ptok_eff')}*{I('fx')}/1000000000"
            f["inf_s"] = f"={c}{R['ai_q']}*{tokq}*{I('ptok_eff')}*{I('fx')}/1000000"
            f["inf"] = f"={c}{R['inf_ag']}+{c}{R['inf_s']}"
            f["ai_net"] = f"={c}{R['ai_inc']}+{c}{R['ai_ero']}-{c}{R['inf']}"
            f["ai_oi"] = f"={c}{R['ai_inc']}*{I('m_inc')}+{c}{R['ai_ero']}*{I('m_lost')}-{c}{R['inf']}"
            f["opm0"] = f"={path('opm26','opm32',c)}"
            f["oi"] = f"={c}{R['rev_noai']}*{c}{R['opm0']}+{c}{R['ai_oi']}"
            f["cogs"] = f"={I('cogs_pct')}*{c}{R['rev']}+{c}{R['inf']}"
            f["gp"] = f"={c}{R['rev']}-{c}{R['cogs']}"
            f["opx"] = f"={c}{R['gp']}-{c}{R['oi']}"
            f["nonop"] = f"={I('nonop')}"
            f["pbt"] = f"={c}{R['oi']}+{c}{R['nonop']}"
            f["tax"] = f"={c}{R['pbt']}*{I('tax')}"
            f["ni"] = f"={c}{R['pbt']}-{c}{R['tax']}"
            f["nci"] = f"={c}{R['ni']}*{I('nci_sh')}"
            f["nio"] = f"={c}{R['ni']}-{c}{R['nci']}"
            f["eps"] = f"={c}{R['nio']}*1000/({I('sh_dil')})"
            f["om"] = f"={c}{R['oi']}/{c}{R['rev']}"
        for k, v in f.items():
            cell = put(PL, f"{c}{R[k]}", v)
            cell.number_format = "0.0%" if k in PCT else ("0.00" if k in ("eps",) else "#,##0.0")

# 標準損益表（情境 B；GAAP＝IFRS）
row = STD0
put(PL, f"A{row}", "標準損益表（情境 B；IFRS 合併，十億日圓，EPS 日圓）", BOLD, HDR)
row += 1
put(PL, f"A{row}", "項目", BOLD, HDR)
for j, y in enumerate(YRS):
    put(PL, f"{CL[j]}{row}", y, BOLD, HDR)
STDH = row
row += 1
RB = PR["B"]
IX = {}
std_def = [
    ("rev", "營收", "十億日圓", lambda c, j: f"={c}{RB['rev']}"),
    ("yoy", "營收年增率", "%", lambda c, j: (f"={c}{{rev}}/{CL[j-1]}{{rev}}-1" if j > 0 else None)),
    ("cogs", "營業成本（COGS）", "十億日圓", lambda c, j: f"={c}{RB['cogs']}"),
    ("gp", "毛利", "十億日圓", lambda c, j: f"={c}{RB['gp']}"),
    ("gm", "毛利率", "%", lambda c, j: f"={c}{{gp}}/{c}{{rev}}"),
    ("opx", "營業費用", "十億日圓", lambda c, j: f"={c}{RB['opx']}"),
    ("oi", "營業利益", "十億日圓", lambda c, j: f"={c}{RB['oi']}"),
    ("opm", "營業利益率", "%", lambda c, j: f"={c}{{oi}}/{c}{{rev}}"),
    ("oth", "其他收入（淨額）", "十億日圓", lambda c, j: f"={c}{RB['nonop']}"),
    ("pti", "稅前利益", "十億日圓", lambda c, j: f"={c}{RB['pbt']}"),
    ("tax", "所得稅", "十億日圓", lambda c, j: f"={c}{RB['tax']}"),
    ("ni", "稅後淨利（含非控制權益）", "十億日圓", lambda c, j: f"={c}{RB['ni']}"),
    ("nm", "淨利率", "%", lambda c, j: f"={c}{{ni}}/{c}{{rev}}"),
    ("eps", "EPS（完全稀釋；歸屬母公司，日圓）", "日圓", lambda c, j: f"={c}{RB['eps']}"),
]
for k, lab, unit, fn in std_def:
    IX[k] = row
    row += 1
for k, lab, unit, fn in std_def:
    rr_ = IX[k]
    put(PL, f"A{rr_}", lab); put(PL, f"B{rr_}", unit)
    for j in range(10):
        c = CL[j]
        v = fn(c, j)
        if v is None: continue
        v = v.format(**IX)
        cell = put(PL, f"{c}{rr_}", v)
        cell.number_format = "0.0%" if unit == "%" else ("0.00" if k == "eps" else "#,##0.0")
put(PL, f"A{row}", "比率（表尾）", BOLD)
row += 1
RAT = row
RID = {}
for k, lab in (("r_rev", "營收年增率"), ("r_gm", "毛利率"), ("r_opm", "營業利益率"), ("r_nm", "淨利率"), ("r_niy", "淨利年增率")):
    RID[k] = row
    put(PL, f"A{row}", lab); put(PL, f"B{row}", "%")
    for j in range(10):
        c = CL[j]
        v = {"r_rev": (f"={c}{IX['rev']}/{CL[j-1]}{IX['rev']}-1" if j else None), "r_gm": f"={c}{IX['gm']}", "r_opm": f"={c}{IX['opm']}", "r_nm": f"={c}{IX['nm']}",
             "r_niy": (f"={c}{IX['ni']}/{CL[j-1]}{IX['ni']}-1" if j else None)}[k]
        if v: put(PL, f"{c}{row}", v, None, None, "0.0%")
    row += 1
STD_NOTE = row
put(PL, f"A{row}", "標準損益表口徑：IFRS（公司法定口徑；公司另揭露非 IFRS 的調整後 EBITDA／EPS，見 Recon）。預測期營業成本＝營業成本率 × 營收＋推論成本；營業費用為殘差（營業利益由「無 AI 營收 × 營業利益率＋AI 淨效果」決定），營業利益以下取自上方情境 B 區塊。FY2023A 毛利與營業費用口徑取自 FY2024 決算短信（FY2023 SG&A 在兩份決算短信略有不同：1,091,094 與 1,095,832 百萬日圓，本表以營業成本與營業利益導出，不直接引用 SG&A）。")
row += 2
PLSUM = row
put(PL, f"A{row}", "分部對帳（三情境 × 全部年度；差額＝分部合計 − 合併營收，單位十億日圓，FY2021–FY2023 來源四捨五入 ±0.002）", BOLD)
row += 1
for s in "ABC":
    put(PL, f"A{row}", f"情境 {s}：分部合計 − 合併營收（絕對值最大者）")
    put(PL, f"C{row}", "=MAX(" + ",".join(f"ABS({c}{PR[s]['chk_seg']})" for c in CL) + ")", None, None, "0.000")
    row += 1
for s in "ABC":
    put(PL, f"A{row}", f"情境 {s}：Media 子線合計 − Media（絕對值最大者）")
    put(PL, f"C{row}", "=MAX(" + ",".join(f"ABS({c}{PR[s]['chk_media']})" for c in CL) + ")", None, None, "0.000")
    row += 1
PL.column_dimensions["A"].width = 62

# =====================================================================================
# DCF 頁：每情境 FCFF
# =====================================================================================
DC = S["DCF"]
put(DC, "A1", "DCF — 無槓桿自由現金流（FCFF＝NOPAT＋折舊攤銷 − 資本支出 − 營運資金增加；股票薪酬不加回；LY 不揭露 SBC 費用）", BOLD)
put(DC, "A2", "金額十億日圓。期中折現；FY2026 只計評價日之後的剩餘期間（stub）。WACC 與稅率見 Val 頁。PayPay 銀行之存款、放款不在 FCFF（營運資金僅以營收增量比率估計）。", None)
put(DC, "A4", "項目", BOLD, HDR); put(DC, "B4", "單位", BOLD, HDR)
for j, y in enumerate(YRS):
    put(DC, f"{CL[j]}4", y, BOLD, HDR)
put(DC, "A5", "期間占比（FY2026＝評價日至 3/31 的比例）")
put(DC, "A6", "折現期間（年，期中）")
put(DC, "A7", "折現因子（WACC）")
for j in range(3, 10):
    c = CL[j]
    if j == 3:
        put(DC, f"{c}5", f"=MAX(0,({I('fye')}-{I('val_date')}))/365", None, None, "0.000")
        put(DC, f"{c}6", f"={c}5/2", None, None, "0.000")
    else:
        put(DC, f"{c}5", 1, None, None, "0.000")
        put(DC, f"{c}6", f"=${CL[3]}5+{j-4}+0.5", None, None, "0.000")
    put(DC, f"{c}7", f"=1/(1+Val!$B$WACCROW)^{c}6", None, None, "0.0000")
DR = {}
row = 9
for s in "ABC":
    put(DC, f"A{row}", f"情境 {s}", BOLD, HDR)
    row += 1
    rows = {}
    for k, lab in (("oi", "營業利益（EBIT）"), ("nopat", "NOPAT＝EBIT ×（1 − 稅率）"), ("da", "＋折舊攤銷"), ("cx", "− 資本支出"), ("wc", "− 營運資金增加"), ("fcff", "FCFF"), ("pv", "FCFF 現值（含 stub 占比）"),
                   ("int", "稅後淨利息（有息負債 × 稅前負債成本 ×（1 − 稅率））"), ("div", "股利（每股股利 × 股數）"), ("gate", "現金收支：FCFF − 稅後利息 − 股利")):
        rows[k] = row
        put(DC, f"A{row}", lab)
        row += 1
    DR[s] = rows
    R = PR[s]
    for j in range(3, 10):
        c = CL[j]; p = CL[j - 1]
        put(DC, f"{c}{rows['oi']}", f"=PL!{c}{R['oi']}", None, None, "#,##0.0")
        put(DC, f"{c}{rows['nopat']}", f"={c}{rows['oi']}*(1-{I('tax')})", None, None, "#,##0.0")
        put(DC, f"{c}{rows['da']}", f"=PL!{c}{R['rev']}*{I('da_pct')}", None, None, "#,##0.0")
        put(DC, f"{c}{rows['cx']}", f"=PL!{c}{R['rev']}*{I('cx_pct')}", None, None, "#,##0.0")
        put(DC, f"{c}{rows['wc']}", f"=(PL!{c}{R['rev']}-PL!{p}{R['rev']})*{I('wc_k')}", None, None, "#,##0.0")
        put(DC, f"{c}{rows['fcff']}", f"={c}{rows['nopat']}+{c}{rows['da']}-{c}{rows['cx']}-{c}{rows['wc']}", None, None, "#,##0.0")
        put(DC, f"{c}{rows['pv']}", f"={c}{rows['fcff']}*{c}$5*{c}$7", None, None, "#,##0.0")
        put(DC, f"{c}{rows['int']}", f"={I('debt0')}*{I('kd')}*(1-{I('tax')})", None, None, "#,##0.0")
        put(DC, f"{c}{rows['div']}", f"={I('dps0')}*(1+{I('dps_g')})^{j-3}*{I('sh_dil')}/1000", None, None, "#,##0.0")
        put(DC, f"{c}{rows['gate']}", f"={c}{rows['fcff']}-{c}{rows['int']}-{c}{rows['div']}", None, None, "#,##0.0")
    row += 1
DC.column_dimensions["A"].width = 58

# =====================================================================================
# Val 頁
# =====================================================================================
V = S["Val"]
VR = {}
put(V, "A1", "Val — 目標價（股價基準日＋12 個月）＝DCF（情境加權）× 權重 ＋ 分部加總（情境加權）× 權重", BOLD)
put(V, "A2", "DCF 為標準 FCFF 版（個股DCF標準規格 v1.0）；分部加總為 FY2027E 分部營收 × 公司 FY2025 調整後 EBITDA 率 × 倍數（倍數為 Assumed，同業未蒐集）。五級評等：> +50% 強力買進、> +25% 買進、≥ −25% 中立、≥ −50% 放空、其餘強力放空；先看收支關卡。", None)
r = 4
put(V, f"A{r}", "0.3 結論", BOLD)
r += 1
for j, h in enumerate(["項目", "數值", "對照", "對照", "說明"]):
    put(V, f"{L(j+1)}{r}", h, BOLD, HDR)
r += 1
for k in ("tpdcf", "tpcomp", "tp", "px", "cons", "up", "upc", "tier", "valrate", "gate_chk", "gate_in", "rec"):
    VR[k] = r
    r += 1
r += 1
put(V, f"A{r}", "0.2 折現率", BOLD); r += 1
for k, lab in (("rf", "無風險利率（基準日與現價基準日差距見下列逾期核對）"), ("beta", "β"), ("erp", "股權風險溢酬"), ("ke", "股權成本 ke"), ("kd", "稅前負債成本"), ("kdat", "稅後負債成本"),
               ("emkt", "股權市值（十億日圓）"), ("dmkt", "有息負債（十億日圓）"), ("we", "股權權重"), ("wacc", "WACC"), ("ronic", "RONIC（＝WACC，新投資價值中性）"), ("g", "永續成長率")):
    VR[k] = r
    put(V, f"A{r}", lab)
    r += 1
r += 1
put(V, f"A{r}", "0.1 DCF 彙總（三情境）", BOLD); r += 1
VR["dh"] = r
for j, h in enumerate(["項目", "情境 A", "情境 B", "情境 C"]):
    put(V, f"{L(j+1)}{r}", h, BOLD, HDR)
r += 1
for k, lab in (("pvx", "明確期 FCFF 現值（含 stub）"), ("nop33", "FY2033 NOPAT＝FY2032 NOPAT ×（1 + g）"), ("tv", "永續價值＝NOPAT(FY2033)×(1 − g／RONIC)／(WACC − g)"),
               ("pvtv", "永續價值現值"), ("ev", "企業價值 EV"), ("tvsh", "永續價值占 EV"), ("nd", "− 淨有息負債（不含金融事業）"), ("nci", "− 非控制權益（帳面 × 倍數）"),
               ("eqv", "股權價值"), ("shv", "每股價值（日圓，評價日）"), ("dcf12", "DCF 情境價值推到 12 個月後（× (1+ke) − 股利）")):
    VR[k] = r
    put(V, f"A{r}", lab)
    r += 1
r += 1
put(V, f"A{r}", "0.0 分部加總（目標時點 TTM 分部 EBITDA＝FY2027E 與 FY2028E 各半 × 公司 FY2025 調整後 EBITDA 率 × 同業 TTM 倍數）", BOLD); r += 1
VR["sh"] = r
for j, h in enumerate(["分部", "情境 A", "情境 B", "情境 C", "倍數（分母＝TTM EBITDA）", "EBITDA 率"]):
    put(V, f"{L(j+1)}{r}", h, BOLD, HDR)
r += 1
for k, lab in (("s_media", "Media（EBITDA，十億日圓）"), ("s_com", "Commerce（EBITDA）"), ("s_str", "Strategic（EBITDA）"), ("s_ev", "分部加總 EV"), ("s_nd", "− 淨有息負債"), ("s_nci", "− 非控制權益"),
               ("s_eq", "股權價值"), ("s_ps", "每股價值（日圓，目標時點）")):
    VR[k] = r
    put(V, f"A{r}", lab)
    r += 1
r += 1
put(V, f"A{r}", "情境目標價", BOLD); r += 1
for k, lab in (("t_A", "情境 A 目標價"), ("t_B", "情境 B 目標價"), ("t_C", "情境 C 目標價"), ("wsum", "情境權重合計"), ("per", "加權目標價隱含 FY2027E 本益比（歸屬母公司 EPS；分母＝情境 B 模型）"),
               ("fcf_gate", "收支關卡（基準情境 B：FY2026–FY2030 FCFF − 稅後利息 − 股利 合計，十億日圓）"), ("gate_A", "情境 A 現金收支合計（FY2026–FY2030，十億日圓）"), ("gate_C", "情境 C 現金收支合計（FY2026–FY2030，十億日圓）")):
    VR[k] = r
    put(V, f"A{r}", lab)
    r += 1
VR["ai_B"] = r
put(V, f"A{r}", "AI 淨效果（營業利益）FY2032E：A／B／C（十億日圓）")
r += 1
VR["ai_tp"] = r
put(V, f"A{r}", "AI 淨效果對 DCF 價值的影響：情境價值 − 無 AI 對照（只在 FY2032E 營業利益層，十億日圓）")
r += 2
put(V, f"A{r}", "0.4 DCF 與 Comparable 落差分解（情境 B，每股日圓；兩法皆在目標時點）", BOLD); r += 1
for k, lab in (("g_tot", "落差＝Comparable 每股 − DCF 每股"), ("g_mult", "其中：倍數（分部加總 EV − DCF 隱含目標時點 EV，÷ 稀釋股數）"),
               ("g_lev", "其中：槓桿（兩法的淨債務與非控制權益扣除差）"), ("g_time", "其中：時點（兩法同在目標時點）"),
               ("g_sbc", "其中：股票薪酬（LY 不揭露；DCF 不加回，公司調整後 EBITDA 口徑未拆）"), ("g_chk", "檢查：分解合計 − 落差（應為零）"),
               ("g_evd", "DCF 隱含目標時點 EV（十億日圓）"), ("g_imp", "DCF 隱含 EV／目標時點 TTM 分部 EBITDA 合計（倍）"), ("g_sum", "分部加總使用的加權倍數（倍）")):
    VR[k] = r
    put(V, f"A{r}", lab)
    r += 1
r += 1
put(V, f"A{r}", "0.5 本公司現行倍數與同業（第 47 條）", BOLD); r += 1
for k, lab in (("c_ev", "本公司現行 EV（現價 × 稀釋股數 ＋ 淨債務 ＋ 非控制權益帳面 × 倍數，十億日圓）"), ("c_eb", "本公司 TTM EBITDA（資料商口徑，與同業同源；資料日見 In 頁）"),
               ("c_mult", "本公司現行 EV／TTM EBITDA（倍）"), ("c_peer", "同業中位數依分部 EBITDA 加權（倍）"), ("c_ratio", "現行倍數 ÷ 同業加權中位數"), ("c_rr", "隱含重估（同業中位數倍數 ÷ 現行倍數 − 1）")):
    VR[k] = r
    put(V, f"A{r}", lab)
    r += 1
r += 1
put(V, f"A{r}", "0.6 逾期核對（第 15 條；與現價基準日相差超過 2 個交易日標逾期）", BOLD); r += 1
for k, lab in (("o_rf", "無風險利率基準日"), ("o_peer", "同業倍數資料日"), ("o_cons", "共識目標價基準日"), ("o_ly", "本公司 TTM EBITDA 資料日")):
    VR[k] = r
    put(V, f"A{r}", lab)
    r += 1

b = lambda k: f"B{VR[k]}"
put(V, f"B{VR['rf']}", f"={I('rf')}", None, None, "0.00%")
put(V, f"B{VR['beta']}", f"={I('beta')}", None, None, "0.00")
put(V, f"B{VR['erp']}", f"={I('erp')}", None, None, "0.0%")
put(V, f"B{VR['ke']}", f"=B{VR['rf']}+B{VR['beta']}*B{VR['erp']}", None, None, "0.00%")
put(V, f"B{VR['kd']}", f"={I('kd')}", None, None, "0.00%")
put(V, f"B{VR['kdat']}", f"=B{VR['kd']}*(1-{I('tax')})", None, None, "0.00%")
put(V, f"B{VR['emkt']}", f"={I('px')}*{I('sh_dil')}/1000", None, None, "#,##0")
put(V, f"B{VR['dmkt']}", f"={I('debt0')}", None, None, "#,##0")
put(V, f"B{VR['we']}", f"=B{VR['emkt']}/(B{VR['emkt']}+B{VR['dmkt']})", None, None, "0.0%")
put(V, f"B{VR['wacc']}", f"=B{VR['we']}*B{VR['ke']}+(1-B{VR['we']})*B{VR['kdat']}", None, None, "0.00%")
put(V, f"B{VR['ronic']}", f"=B{VR['wacc']}", None, None, "0.00%")
put(V, f"B{VR['g']}", f"={I('g_tv')}", None, None, "0.0%")
for j in range(3, 10):
    DC[f"{CL[j]}7"].value = DC[f"{CL[j]}7"].value.replace("WACCROW", str(VR["wacc"]))

SC = {"A": "B", "B": "C", "C": "D"}
for s in "ABC":
    c = SC[s]
    D_ = DR[s]
    put(V, f"{c}{VR['pvx']}", f"=SUM(DCF!{CL[3]}{D_['pv']}:{CL[9]}{D_['pv']})", None, None, "#,##0.0")
    put(V, f"{c}{VR['nop33']}", f"=DCF!{CL[9]}{D_['nopat']}*(1+$B${VR['g']})", None, None, "#,##0.0")
    put(V, f"{c}{VR['tv']}", f"={c}{VR['nop33']}*(1-$B${VR['g']}/$B${VR['ronic']})/($B${VR['wacc']}-$B${VR['g']})", None, None, "#,##0.0")
    put(V, f"{c}{VR['pvtv']}", f"={c}{VR['tv']}*DCF!{CL[9]}7*(1+$B${VR['wacc']})^(-0.5)", None, None, "#,##0.0")  # 終值在 FY2032 年底：期中因子 × (1+WACC)^0.5
    put(V, f"{c}{VR['ev']}", f"={c}{VR['pvx']}+{c}{VR['pvtv']}", None, None, "#,##0.0")
    put(V, f"{c}{VR['tvsh']}", f"={c}{VR['pvtv']}/{c}{VR['ev']}", None, None, "0.0%")
    put(V, f"{c}{VR['nd']}", f"=-{I('nd0')}", None, None, "#,##0.0")
    put(V, f"{c}{VR['nci']}", f"=-{I('nci0')}*{I('nci_mult')}", None, None, "#,##0.0")
    put(V, f"{c}{VR['eqv']}", f"={c}{VR['ev']}+{c}{VR['nd']}+{c}{VR['nci']}", None, None, "#,##0.0")
    put(V, f"{c}{VR['shv']}", f"={c}{VR['eqv']}*1000/({I('sh_dil')})", None, None, "#,##0.00")
    put(V, f"{c}{VR['dcf12']}", f"={c}{VR['shv']}*(1+$B${VR['ke']})^(({I('tp_date')}-{I('val_date')})/365)-{I('dps0')}*(({I('tp_date')}-{I('val_date')})/365)", None, None, "#,##0.00")
    # 分部加總
    R = PR[s]
    f27 = CL[4]; f28 = CL[5]
    put(V, f"{c}{VR['s_media']}", f"=(PL!{f27}{R['media']}+PL!{f28}{R['media']})/2*{I('em_media')}", None, None, "#,##0.0")
    put(V, f"{c}{VR['s_com']}", f"=(PL!{f27}{R['com']}+PL!{f28}{R['com']})/2*{I('em_com')}", None, None, "#,##0.0")
    put(V, f"{c}{VR['s_str']}", f"=(PL!{f27}{R['str']}+PL!{f28}{R['str']})/2*{I('em_str')}", None, None, "#,##0.0")
    put(V, f"{c}{VR['s_ev']}", f"={c}{VR['s_media']}*$E${VR['s_media']}+{c}{VR['s_com']}*$E${VR['s_com']}+{c}{VR['s_str']}*$E${VR['s_str']}", None, None, "#,##0.0")
    put(V, f"{c}{VR['s_nd']}", f"=-{I('nd0')}", None, None, "#,##0.0")
    put(V, f"{c}{VR['s_nci']}", f"=-{I('nci0')}*{I('nci_mult')}", None, None, "#,##0.0")
    put(V, f"{c}{VR['s_eq']}", f"={c}{VR['s_ev']}+{c}{VR['s_nd']}+{c}{VR['s_nci']}", None, None, "#,##0.0")
    put(V, f"{c}{VR['s_ps']}", f"={c}{VR['s_eq']}*1000/({I('sh_dil')})", None, None, "#,##0.00")
    put(V, f"B{VR['t_' + s]}", f"={I('w_dcf')}*{c}{VR['dcf12']}+(1-{I('w_dcf')})*{c}{VR['s_ps']}", None, None, "#,##0.00")
    put(V, f"B{VR['gate_' + s] if s != 'B' else VR['fcf_gate']}", f"=SUM(DCF!{CL[3]}{D_['gate']}:{CL[7]}{D_['gate']})", None, None, "#,##0.0")
    put(V, f"{c}{VR['ai_B']}", f"=PL!{CL[9]}{R['ai_oi']}", None, None, "#,##0.0")
for k, col, ik in (("s_media", "E", "m_media"), ("s_com", "E", "m_com"), ("s_str", "E", "m_str")):
    put(V, f"E{VR[k]}", f"={I(ik)}", None, None, "0.0")
for k, ik in (("s_media", "em_media"), ("s_com", "em_com"), ("s_str", "em_str")):
    put(V, f"F{VR[k]}", f"={I(ik)}", None, None, "0.0%")
put(V, f"B{VR['wsum']}", f"={I('tpw_A')}+{I('tpw_B')}+{I('tpw_C')}", None, None, "0%")
put(V, f"B{VR['per']}", f"=B{VR['tp']}/PL!{CL[4]}{PR['B']['eps']}", None, None, "0.0")
put(V, f"B{VR['tpdcf']}", f"={I('tpw_A')}*B{VR['dcf12']}+{I('tpw_B')}*C{VR['dcf12']}+{I('tpw_C')}*D{VR['dcf12']}", None, None, "#,##0.00")
put(V, f"B{VR['tpcomp']}", f"={I('tpw_A')}*B{VR['s_ps']}+{I('tpw_B')}*C{VR['s_ps']}+{I('tpw_C')}*D{VR['s_ps']}", None, None, "#,##0.00")
put(V, f"B{VR['tp']}", f"={I('w_dcf')}*B{VR['tpdcf']}+(1-{I('w_dcf')})*B{VR['tpcomp']}", None, None, "#,##0.00")
put(V, f"B{VR['px']}", f"={I('px')}", None, None, "#,##0.00")
put(V, f"B{VR['cons']}", f"={I('cons_tp')}", None, None, "#,##0")
put(V, f"B{VR['up']}", f"=B{VR['tp']}/B{VR['px']}-1", None, None, "0.0%")
put(V, f"B{VR['upc']}", f"=B{VR['tpcomp']}/B{VR['px']}-1", None, None, "0.0%")
put(V, f"B{VR['tier']}", f'=IF(B{VR["up"]}>0.5,"> +50%",IF(B{VR["up"]}>0.25,"+25% 至 +50%",IF(B{VR["up"]}>=-0.25,"−25% 至 +25%",IF(B{VR["up"]}>=-0.5,"−50% 至 −25%","< −50%"))))')
put(V, f"B{VR['valrate']}", f'=IF(B{VR["up"]}>0.5,"強力買進",IF(B{VR["up"]}>0.25,"買進",IF(B{VR["up"]}>=-0.25,"中立",IF(B{VR["up"]}>=-0.5,"放空","強力放空"))))')
put(V, f"B{VR['gate_in']}", f"=B{VR['fcf_gate']}", None, None, "#,##0.0")
put(V, f"B{VR['gate_chk']}", f'=IF(B{VR["fcf_gate"]}>=0,"通過（FY2026–FY2030 自由現金流扣利息與股利合計為正）","不通過（資金缺口，不得買進）")')
put(V, f"B{VR['rec']}", f'=IF(AND(B{VR["gate_in"]}<0,OR(B{VR["valrate"]}="強力買進",B{VR["valrate"]}="買進")),"收支不足",B{VR["valrate"]})')
# 分解與現行倍數、逾期
sh = I('sh_dil')
put(V, f"B{VR['g_tot']}", f"=C{VR['s_ps']}-C{VR['dcf12']}", None, None, "#,##0.00")
put(V, f"B{VR['g_evd']}", f"=C{VR['dcf12']}*{sh}/1000-C{VR['nd']}-C{VR['nci']}", None, None, "#,##0")
put(V, f"B{VR['g_mult']}", f"=(C{VR['s_ev']}-B{VR['g_evd']})*1000/{sh}", None, None, "#,##0.00")
put(V, f"B{VR['g_lev']}", f"=((C{VR['s_nd']}+C{VR['s_nci']})-(C{VR['nd']}+C{VR['nci']}))*1000/{sh}", None, None, "#,##0.00")
put(V, f"B{VR['g_time']}", 0, None, None, "#,##0.00")
put(V, f"B{VR['g_sbc']}", "未取得")
put(V, f"B{VR['g_chk']}", f"=B{VR['g_mult']}+B{VR['g_lev']}+B{VR['g_time']}-B{VR['g_tot']}", None, None, "#,##0.000")
put(V, f"B{VR['g_imp']}", f"=B{VR['g_evd']}/(C{VR['s_media']}+C{VR['s_com']}+C{VR['s_str']})", None, None, "0.0")
put(V, f"B{VR['g_sum']}", f"=C{VR['s_ev']}/(C{VR['s_media']}+C{VR['s_com']}+C{VR['s_str']})", None, None, "0.0")
put(V, f"B{VR['c_ev']}", f"={I('px')}*{sh}/1000+{I('nd0')}+{I('nci0')}*{I('nci_mult')}", None, None, "#,##0")
put(V, f"B{VR['c_eb']}", f"={I('ly_ebitda')}", None, None, "#,##0.0")
put(V, f"B{VR['c_mult']}", f"=B{VR['c_ev']}/B{VR['c_eb']}", None, None, "0.0")
put(V, f"B{VR['c_peer']}", f"=(C{VR['s_media']}*E{VR['s_media']}+C{VR['s_com']}*E{VR['s_com']}+C{VR['s_str']}*E{VR['s_str']})/(C{VR['s_media']}+C{VR['s_com']}+C{VR['s_str']})", None, None, "0.0")
put(V, f"B{VR['c_ratio']}", f"=B{VR['c_mult']}/B{VR['c_peer']}", None, None, "0.00")
put(V, f"B{VR['c_rr']}", f"=B{VR['c_peer']}/B{VR['c_mult']}-1", None, None, "0.0%")
for k, dref in (("o_rf", I('rf_date')), ("o_cons", I('cons_date')), ("o_ly", I('ly_eb_date'))):
    put(V, f"B{VR[k]}", f"={dref}", None, None, "yyyy-mm-dd")
    put(V, f"C{VR[k]}", f'=IF(NETWORKDAYS(B{VR[k]},{I("px_date")})-1>2,"逾期","正常")')
put(V, f"B{VR['o_peer']}", "=MIN(Peers!$I$4:$I$14)", None, None, "yyyy-mm-dd")
put(V, f"C{VR['o_peer']}", '=COUNTIF(Peers!$J$4:$J$14,"逾期")&" 列逾期（共 11 列）；最舊資料日見 B"')
for k, lab, note in (("tpdcf", "DCF 目標價（情境加權）", "三情境 DCF 推到 12 個月後（扣期間股利）依權重 25／50／25 加權"),
                     ("tpcomp", "Comparable 目標價（分部加總，情境加權）", "同業 TTM 倍數（Analogy，Peers 頁）× 目標時點 TTM 分部 EBITDA"),
                     ("tp", "目標價（DCF 與 Comparable 加權）", "DCF 與分部加總依 In 頁權重加權"),
                     ("px", "現價（2026-11-16 收盤）", "traders.co.jp 與 Yahoo!ファイナンス兩源一致"),
                     ("cons", "市場共識目標價", "kabuyoho.jp 10 家，2026-09-08，已早於 Q2 財報（2026-11-05）上修；只作對照"),
                     ("up", "上檔空間：加權目標價", ""), ("upc", "上檔空間：Comparable 目標價", ""), ("tier", "上檔空間級距", "五級固定規則；> +50% 強力買進、> +25% 買進、≥ −25% 中立、≥ −50% 放空"), ("valrate", "估值評等（依加權目標價上檔空間，不含收支關卡）", "五級之一"), ("gate_in", "入不敷出判定用：基準情境 B 收支關卡合計（十億日圓）", "供 check_rating_gate 核對；負值時買進以上改判收支不足"),
                     ("gate_chk", "收支關卡判定", "先看收支再看估值（個股模型原則第 10 節第 63 條）"), ("rec", "投資建議", "估值評等；收支關卡不通過且評等為買進以上時改為收支不足")):
    put(V, f"A{VR[k]}", lab)
    put(V, f"E{VR[k]}", note)
# 三情境 AI 營業利益 FY2032（ai_B 列 B/C/D）
put(V, f"A{VR['ai_tp']}", "AI 淨效果 FY2032E 營業利益：情境 A − 情境 C（十億日圓）")
put(V, f"B{VR['ai_tp']}", f"=B{VR['ai_B']}-D{VR['ai_B']}", None, None, "#,##0.0")
V.column_dimensions["A"].width = 64; V.column_dimensions["E"].width = 70

# =====================================================================================
# Peers 頁：同業 EV／EBITDA（TTM）
# =====================================================================================
PE = S["Peers"]
put(PE, "A1", "Peers — 同業 EV／EBITDA 倍數（分母期間＝TTM EBITDA；資料商標準化口徑；標記 Analogy）", BOLD)
put(PE, "A2", "來源：stockanalysis.com 各公司 statistics 頁（2026-11-16 擷取，頁面顯示的價格日即資料日）。EV 與 EBITDA 單位為各公司本幣十億。會計準則與租賃處理多數頁面未載，寫「未取得」；資料商 EV 以總債務計算，是否含營業租賃未載。LY（IFRS 16，租賃負債已入淨債務、EBITDA 不扣租金）與 US GAAP／J-GAAP 同業的 EBITDA 口徑不完全相同，未調整。逾期＝資料日與現價基準日相差超過 2 個交易日（第 15 條）。", None)
for j, h in enumerate(["公司", "代號", "對應分部", "會計準則", "租賃處理", "EV（本幣十億）", "TTM EBITDA（本幣十億）", "EV／EBITDA（分母＝TTM EBITDA）", "資料日", "逾期", "來源", "標記"]):
    put(PE, f"{L(j+1)}3", h, BOLD, HDR)
for i_, (seg, nm, tk, std, ls, ev, eb, mu, dt, cur, path_) in enumerate(PEERS):
    rr_ = 4 + i_
    put(PE, f"A{rr_}", nm); put(PE, f"B{rr_}", tk); put(PE, f"C{rr_}", seg); put(PE, f"D{rr_}", std); put(PE, f"E{rr_}", ls)
    put(PE, f"F{rr_}", ev, BLUE, None, "#,##0.00"); put(PE, f"G{rr_}", eb, BLUE, None, "#,##0.00"); put(PE, f"H{rr_}", mu, BLUE, None, "0.00")
    put(PE, f"I{rr_}", dt, None, None, "yyyy-mm-dd")
    put(PE, f"J{rr_}", f'=IF(NETWORKDAYS(I{rr_},{I("px_date")})-1>2,"逾期","正常")')
    put(PE, f"K{rr_}", f"https://stockanalysis.com/quote/{path_}/statistics/（{cur}）" if "/" in path_ and not path_.startswith("stocks") else f"https://stockanalysis.com/{path_}/statistics/（{cur}）")
    put(PE, f"L{rr_}", "Analogy")
put(PE, "A15", "排除：ASKUL（2678）TTM EBITDA 為負（−12.02 十億日圓），倍數不適用。")
put(PE, "A16", "同業中位數", BOLD)
for rr_, nm, a_, b_ in ((17, "Media 中位數", 4, 7), (18, "Commerce 中位數", 8, 11), (19, "Strategic 中位數", 12, 14)):
    put(PE, f"A{rr_}", nm)
    put(PE, f"H{rr_}", f"=MEDIAN(H{a_}:H{b_})", None, None, "0.00")
    put(PE, f"I{rr_}", f"=MIN(H{a_}:H{b_})", None, None, "0.00"); put(PE, f"J{rr_}", f"=MAX(H{a_}:H{b_})", None, None, "0.00")
put(PE, "I16", "最小"); put(PE, "J16", "最大")
put(PE, "A21", "分部倍數：同業中位數為基準，第 1／第 3 四分位為 Sens 低／高；同業在成長與利益率上不同於 LY 各分部，未做調整（無證據可據以折溢價）。", None)
PE.column_dimensions["A"].width = 22; PE.column_dimensions["D"].width = 34; PE.column_dimensions["K"].width = 60

# =====================================================================================
# RevDCF：現價隱含（封閉解：FY2028–FY2032 統一營業利益率 m*；情境 B 路徑）
# =====================================================================================
RV = S["RevDCF"]
put(RV, "A1", "RevDCF — 現價隱含什麼（情境 B 營收路徑；以封閉解反解 FY2028–FY2032 統一營業利益率，其餘輸入不變）", BOLD)
put(RV, "A2", "FCFF 對 FY2028–FY2032 營業利益率為線性，故以 m＝0 與 m＝1 兩個端點求 EV，線性內插到現價所需 EV。永續價值亦隨 FY2032 營業利益率線性變動。", None)
RVR = {}
r = 4
for k, lab in (("px", "現價（日圓）"), ("eqt", "現價隱含股權價值（十億日圓，評價日；以 ke 倒推 12 個月前＝現價）"), ("evr", "現價隱含企業價值 EV（十億日圓）"),
               ("ev0", "EV(m＝0)"), ("ev1", "EV(m＝1)"), ("m1", "現價隱含 FY2028–FY2032 營業利益率"), ("m0", "模型情境 B FY2028–FY2032 平均營業利益率"),
               ("sps", "分部加總：現價隱含每股價值倍數 ÷ 模型（Comparable／現價）"), ("smul", "分部加總：現價隱含 EV／FY2027E 分部 EBITDA 合計（倍）"), ("smod", "分部加總：模型 EV／同一 EBITDA（倍）"),
               ("rv0", "模型情境 B FY2028 營收（十億日圓）"), ("w0", "模型 WACC")):
    RVR[k] = r
    put(RV, f"A{r}", lab)
    r += 1
RB = PR["B"]
DB = DR["B"]
yrs_c = [CL[5], CL[6], CL[7], CL[8], CL[9]]
put(RV, f"B{RVR['px']}", f"={I('px')}", None, None, "#,##0.00")
put(RV, f"B{RVR['eqt']}", f"=B{RVR['px']}*{I('sh_dil')}/1000", None, None, "#,##0")
put(RV, f"B{RVR['evr']}", f"=B{RVR['eqt']}+{I('nd0')}+{I('nci0')}*{I('nci_mult')}", None, None, "#,##0")
# EV(m)：FY2026–FY2027 以情境 B 模型現值；FY2028–FY2032 FCFF(m)=m*rev*(1-t)+(da-cx)*rev-wc；終值 NOPAT33=m*rev32*(1-t)*(1+g)


def ev_expr(m):
    parts = [f"SUM(DCF!{CL[3]}{DB['pv']}:{CL[4]}{DB['pv']})"]
    for c in yrs_c:
        parts.append(f"(({m})*PL!{c}{RB['rev']}*(1-{I('tax')})+PL!{c}{RB['rev']}*({I('da_pct')}-{I('cx_pct')})-DCF!{c}{DB['wc']})*DCF!{c}$7")
    tv = f"({m})*PL!{CL[9]}{RB['rev']}*(1-{I('tax')})*(1+Val!$B${VR['g']})*(1-Val!$B${VR['g']}/Val!$B${VR['ronic']})/(Val!$B${VR['wacc']}-Val!$B${VR['g']})*DCF!{CL[9]}$7*(1+Val!$B${VR['wacc']})^(-0.5)"
    parts.append(tv)
    return "+".join(parts)


put(RV, f"B{RVR['ev0']}", "=" + ev_expr("0"), None, None, "#,##0")
put(RV, f"B{RVR['ev1']}", "=" + ev_expr("1"), None, None, "#,##0")
put(RV, f"B{RVR['m1']}", f"=(B{RVR['evr']}-B{RVR['ev0']})/(B{RVR['ev1']}-B{RVR['ev0']})", None, None, "0.0%")
put(RV, f"B{RVR['m0']}", "=AVERAGE(" + ",".join(f"PL!{c}{RB['om']}" for c in yrs_c) + ")", None, None, "0.0%")
put(RV, f"B{RVR['sps']}", f"=B{RVR['px']}/Val!C{VR['s_ps']}", None, None, "0.00")
put(RV, f"B{RVR['smul']}", f"=B{RVR['evr']}/(Val!C{VR['s_media']}+Val!C{VR['s_com']}+Val!C{VR['s_str']})", None, None, "0.0")
put(RV, f"B{RVR['smod']}", f"=Val!C{VR['s_ev']}/(Val!C{VR['s_media']}+Val!C{VR['s_com']}+Val!C{VR['s_str']})", None, None, "0.0")
put(RV, f"B{RVR['rv0']}", f"=PL!{CL[5]}{RB['rev']}", None, None, "#,##0")
put(RV, f"B{RVR['w0']}", f"=Val!B{VR['wacc']}", None, None, "0.00%")
r += 1
RVR["sum"] = r
put(RV, f"A{r}", f'="現價隱含 FY2028–FY2032 營業利益率約 "&TEXT(B{RVR["m1"]},"0.0%")&"，模型情境 B 平均 "&TEXT(B{RVR["m0"]},"0.0%")&"；現價隱含 EV "&TEXT(B{RVR["evr"]},"#,##0")&" 十億日圓，是分部加總 FY2027E EBITDA 的 "&TEXT(B{RVR["smul"]},"0.0")&" 倍（模型 "&TEXT(B{RVR["smod"]},"0.0")&" 倍）。"')
RV.column_dimensions["A"].width = 80

# =====================================================================================
# Recon：公司指引、共識與模型
# =====================================================================================
RC = S["Recon"]
put(RC, "A1", "Recon — 公司指引、市場共識與模型的差距（公司指引與共識＝Interested-party，只作對照，不進基準）", BOLD)
put(RC, "A2", "指引來源：FY2026 Q2 簡報 slide 7（2026-11-05 上修後）。共識：FY2026／FY2027 營收與 EPS 共識未找到；平均目標價 583 日圓為 2026-09-08（Q2 上修前）。口徑：公司調整後 EBITDA／EPS 為非 IFRS，與本模型 IFRS 營業利益／EPS 不同口徑，不逐項相比。", None)
for j, h in enumerate(["項目", "年度", "模型情境 B", "公司指引", "差距（模型 ÷ 指引 − 1）", "說明"]):
    put(RC, f"{L(j+1)}4", h, BOLD, HDR)
RCR = {}
r = 5
GUIDE = [("rev", "合併營收", PR["B"]["rev"], 2330.0, "公司 2026-11-05 上修：2.33 兆（+14.8%）"),
         ("media", "Media 營收", PR["B"]["media"], 740.0, "上修後 740.0（+0.7%）"),
         ("com", "Commerce 營收", PR["B"]["com"], 1030.0, "上修後 1,030（+20.1%）；ASKUL、LINE MAN 表現優於預期"),
         ("str", "Strategic 營收", PR["B"]["str"], 570.0, "上修後 570（+27.9%）；PayPay 合併")]
for k, lab, prow, g, note in GUIDE:
    RCR[k] = r
    put(RC, f"A{r}", lab); put(RC, f"B{r}", "FY2026E")
    put(RC, f"C{r}", f"=PL!{CL[3]}{prow}", None, None, "#,##0.0")
    put(RC, f"D{r}", g, BLUE, None, "#,##0.0")
    put(RC, f"E{r}", f"=C{r}/D{r}-1", None, None, "0.0%")
    put(RC, f"F{r}", note)
    r += 1
RCR["ebitda"] = r
put(RC, f"A{r}", "調整後 EBITDA（公司非 IFRS）vs 模型 EBITDA＝營業利益＋折舊攤銷（差異含公司調整項，口徑不同）"); put(RC, f"B{r}", "FY2026E")
put(RC, f"C{r}", f"=PL!{CL[3]}{PR['B']['oi']}+PL!{CL[3]}{PR['B']['rev']}*{I('da_pct')}", None, None, "#,##0.0")
put(RC, f"D{r}", 640.0, BLUE, None, "#,##0.0")
put(RC, f"E{r}", f"=C{r}/D{r}-1", None, None, "0.0%")
put(RC, f"F{r}", "公司指引 640.0bn（+28.9%）；公司定義加回一次性與股票薪酬等，本欄只作量級對照")
r += 1
RCR["eps"] = r
put(RC, f"A{r}", "EPS：模型 IFRS 歸屬母公司 EPS vs 公司調整後 EPS 指引（口徑不同）"); put(RC, f"B{r}", "FY2026E")
put(RC, f"C{r}", f"=PL!{CL[3]}{PR['B']['eps']}", None, None, "#,##0.00")
put(RC, f"D{r}", 35.0, BLUE, None, "#,##0.00")
put(RC, f"E{r}", f"=C{r}/D{r}-1", None, None, "0.0%")
put(RC, f"F{r}", "公司調整後 EPS 指引 35.0 日圓；H1 FY2026 IFRS 基本 EPS 13.54、調整後 17.29，口徑差由收購攤提與一次性項目造成")
r += 2
RCR["h1"] = r
put(RC, f"A{r}", "上半年 FY2026 實績對模型（營收）", BOLD); r += 1
put(RC, f"A{r}", "H1 FY2026 營收（實績，十億日圓）"); put(RC, f"C{r}", 1117.668, BLUE, None, "#,##0.0"); put(RC, f"F{r}", "決算短信 FY2026 Q2 p.11（Verified）")
RCR["h1v"] = r; r += 1
put(RC, f"A{r}", "H1 占模型 FY2026E 營收比例"); put(RC, f"C{r}", f"=C{RCR['h1v']}/C{RCR['rev']}", None, None, "0.0%"); put(RC, f"F{r}", "上半年通常低於全年（Q3 年末檔期、ZOZO 等）；比例過高代表模型偏低")
r += 2
RCR["cons"] = r
put(RC, f"A{r}", "市場共識（未找到：FY2026、FY2027 營收與 EPS；平均目標價 583 日圓為 2026-09-08）", BOLD); r += 1
put(RC, f"A{r}", "共識目標價 vs 加權目標價"); put(RC, f"C{r}", f"=Val!B{VR['tp']}", None, None, "#,##0.00"); put(RC, f"D{r}", f"={I('cons_tp')}", None, None, "#,##0"); put(RC, f"E{r}", f"=C{r}/D{r}-1", None, None, "0.0%")
put(RC, f"F{r}", "共識過期（Q2 上修前）；來源誘因：賣方預估")
RCR["tpc"] = r; r += 2
put(RC, f"A{r}", "未找到：同業 EV／EBITDA 倍數、β、查詢量、ARPU、AI 營收與推論成本金額（見研究包第 9 節）。", None)
RC.column_dimensions["A"].width = 70; RC.column_dimensions["F"].width = 70

# =====================================================================================
# Sens
# =====================================================================================
SE = S["Sens"]
put(SE, "A1", "Sens — 一次改一個輸入（In 頁敏感度低／高值），整本活頁簿以 LibreOffice 重算；加權目標價與 DCF 目標價取自重算後的 Val 頁。", BOLD)
put(SE, "A2", "相對基準＝案例加權目標價 ÷ 基準加權目標價 − 1。各案例鍵與 In 頁鍵相同；本頁值為重算值（sens_lycorp.py），In 頁改過而未重跑時模型狀態顯示 STALE。", None)
for j, h in enumerate(["驅動", "設定", "方向", "加權目標價", "相對基準", "鍵", "說明", "收支關卡 B（五年，十億日圓）", "DCF 目標價", "投資建議"]):
    put(SE, f"{L(j+1)}3", h, BOLD, HDR)
SENS_KEYS = [k for k in INR if k not in NAMES_SCN and k not in ("data_date", "px_date", "val_date", "tp_date", "fye", "px", "sh_out", "dil", "debt0", "nd0", "em_media", "em_com", "em_str", "dps0", "cons_tp", "rf")]
# 有 D 或 F 的鍵
SENS_DEF = []
for k in SENS_KEYS:
    row_ = [x for x in IN_ROWS if x[0] == k][0]
    lo, base, hi = row_[3], row_[4], row_[5]
    if lo is not None and lo != base: SENS_DEF.append((k, "低", lo, row_[1]))
    if hi is not None and hi != base: SENS_DEF.append((k, "高", hi, row_[1]))
COMBO = {"@ptok_price": ("ptok、ag_price", "ptok=1.568;ag_price=1500", "價格與成本聯動：token 價格取 Sol 級（高案），同時 Agent i 月費取高案 1,500 日圓（成本轉嫁）"),
         "@ptok_route": ("ptok、route_cheap", "ptok=1.568;route_cheap=0.7", "模型組合：token 價格取 Sol 級，但 70% token 路由到 Luna 級便宜模型")}
for ck, (kt, vv_, lab_) in COMBO.items():
    SENS_DEF.append((ck, "聯動", vv_, lab_))
r = 4
SE_ROWS = []
sj = (SENS or {}).get("cases", {})
for k, side, v, lab in SENS_DEF:
    put(SE, f"A{r}", f"{lab}" if k.startswith("@") else (f"{lab}：{side}（上界：只動成本一邊，未連動價格與路由）" if (k == "ptok" and side == "高") else f"{lab}：{side}"))
    put(SE, f"B{r}", f"基準 → {v}")
    put(SE, f"C{r}", side)
    cs = sj.get(f"{k}|{side}")
    if cs:
        put(SE, f"D{r}", cs["tp"], None, None, "#,##0.00")
        put(SE, f"H{r}", cs["gate"], None, None, "#,##0.0")
        put(SE, f"I{r}", cs["dcf"], None, None, "#,##0.00")
        put(SE, f"J{r}", cs["rec"])
    put(SE, f"E{r}", f"=IF(D{r}=\"\",\"\",D{r}/$D$BASEROW-1)", None, None, "0.0%")
    put(SE, f"F{r}", COMBO[k][0] if k.startswith("@") else k)
    put(SE, f"G{r}", ("重算值（sens_lycorp.py；聯動案例）" if k.startswith("@") else "重算值（sens_lycorp.py）") if cs else "未重算")
    SE_ROWS.append(r)
    r += 1
SE_END = r - 1
r += 1
SE_BASE = r
put(SE, f"A{r}", "基準加權目標價（日圓）")
put(SE, f"D{r}", f"=Val!B{VR['tp']}", None, None, "#,##0.00")
for rr_ in SE_ROWS:
    SE[f"E{rr_}"].value = SE[f"E{rr_}"].value.replace("BASEROW", str(SE_BASE))
r += 1
SE_STALE = r
put(SE, f"A{r}", "Sens 重跑時的基準加權目標價（日圓）")
put(SE, f"D{r}", (SENS or {}).get("base_tp"), None, None, "#,##0.00")
r += 1
# 輸入指紋
num_inputs = [(k, [x for x in IN_ROWS if x[0] == k][0]) for k in INR]
SE_FP = r
put(SE, f"A{r}", "In 頁輸入指紋（現值：基準欄 E 數值型輸入的加權和）")
fp_terms = []
for i_, (k, row_) in enumerate(num_inputs):
    if isinstance(row_[4], (int, float)):
        fp_terms.append(f"{(i_ % 7) + 1}*In!E{INR[k]}")
fp_terms += [f"{(i_ % 5) + 1}*Peers!H{4 + i_}" for i_ in range(len(PEERS))]
put(SE, f"D{r}", "=" + "+".join(fp_terms), None, None, "0.000000")
r += 1
SE_FP0 = r
put(SE, f"A{r}", "In 頁輸入指紋（Sens 重跑時）")
put(SE, f"D{r}", (SENS or {}).get("fp"), None, None, "0.000000")
r += 1
SE_STATUS = r
put(SE, f"A{r}", "模型狀態（CURRENT＝Sens、HTML、README 與本檔輸入一致；STALE＝In 頁已改、須重跑 sens_lycorp.py）")
put(SE, f"D{r}", f'=IF(ABS(D{SE_FP}-D{SE_FP0})<0.000001,"CURRENT","STALE")')
r += 1
SE_BID = r
fp_val = sum(((i_ % 7) + 1) * row_[4] for i_, (k, row_) in enumerate(num_inputs) if isinstance(row_[4], (int, float)))
fp_val += sum(((i_ % 5) + 1) * PEERS[i_][7] for i_ in range(len(PEERS)))
BUILD_ID = "lycorp-v%s-%s" % (VER, hashlib.sha1(f"{fp_val:.6f}".encode()).hexdigest()[:10])
put(SE, f"A{r}", "build_id（README、Sens、HTML 三處相同）")
put(SE, f"D{r}", BUILD_ID)
r += 2
SE_TORN = r
put(SE, f"A{r}", "龍捲風圖說明：|相對基準| ≥ 10% 的案例列在前，未達者摺疊；全部案例見上表。", None)
SE.column_dimensions["A"].width = 60

# =====================================================================================
# README（含 hero 與重點觀察：公式組字）
# =====================================================================================
RD = S["README"]
Vb = lambda k, col="B": f"Val!{col}{VR[k]}"
put(RD, "A1", f'="LY Corporation（4689）個股模型 v{VER}｜build_id "&Sens!D{SE_BID}&"｜資料截止日 "&TEXT({I("data_date")},"yyyy-mm-dd")&"｜成品日 {BUILDDATE.isoformat()}｜股價基準日 "&TEXT({I("px_date")},"yyyy-mm-dd")&"｜評價日 "&TEXT({I("val_date")},"yyyy-mm-dd")&"｜目標時點 "&TEXT({I("tp_date")},"yyyy-mm-dd")&"（本版未做先驗；v0.1：Tokenomics 快照 {TKV}；單位經濟模組在 In 頁自建，v0.2 改讀 unit_economics）"')
put(RD, "A2", "v0.1 要點：首版。命題＝原型 A（按廣告計價）的單位經濟；官方分部 Media、Commerce、Strategic；DCF（標準 FCFF）與分部加總各 50%，三情境只在四個 AI 驅動（分流、價格提升、AI Mode、Agent i 付費）分化。未做項目：先驗、同業倍數、Tokenomics 自用 AI 讀法與 Wattonomics 對照（不適用）、三層協調（不適用：非 Siliconomics 覆蓋）。")
put(RD, "A3", "命題：AI 進入日本最大的 LINE／Yahoo! 廣告平台後，每單位廣告營收的變化＝AI 帶來的營收增量（轉換率與每次點擊價格提升、官方帳號 AI Mode、Agent i 訂閱與廣告）減去營收侵蝕（搜尋查詢分流到聊天介面）減去向模型商購買的推論成本；本模型以三情境檢驗淨效果的正負與大小。")
put(RD, "A4", f'="核心問題：FY2032E 的 AI 淨效果（營業利益）在情境 A／B／C 分別為 "&TEXT(Val!B{VR["ai_B"]},"#,##0.0")&"／"&TEXT(Val!C{VR["ai_B"]},"#,##0.0")&"／"&TEXT(Val!D{VR["ai_B"]},"#,##0.0")&" 十億日圓；基準下 AI 是否足以抵銷傳統搜尋與展示廣告的衰退。"')
put(RD, "A5", f'="投資建議："&Val!B{VR["rec"]}&"——在三情境權重 25／50／25、同業倍數為 Assumed 的基準假設下；加權目標價 "&TEXT(Val!B{VR["tp"]},"#,##0.0")&" 日圓，上檔空間 "&TEXT(Val!B{VR["up"]},"+0.0%;-0.0%")&"（現價 "&TEXT({I("px")},"#,##0.0")&"，2026-11-16 收盤）；收支關卡："&Val!B{VR["gate_chk"]}&"。"')
put(RD, "A6", f'="重點觀察 1：加權目標價 "&TEXT(Val!B{VR["tp"]},"#,##0.0")&"＝DCF "&TEXT(Val!B{VR["tpdcf"]},"#,##0.0")&" × "&TEXT({I("w_dcf")},"0%")&"＋分部加總 "&TEXT(Val!B{VR["tpcomp"]},"#,##0.0")&" × "&TEXT(1-{I("w_dcf")},"0%")&"；情境 A／B／C 為 "&TEXT(Val!B{VR["t_A"]},"#,##0.0")&"／"&TEXT(Val!B{VR["t_B"]},"#,##0.0")&"／"&TEXT(Val!B{VR["t_C"]},"#,##0.0")&"。分部加總的倍數與 β 皆為 Assumed。"')
put(RD, "A7", f'="重點觀察 2：現價隱含 FY2028–FY2032 營業利益率 "&TEXT(RevDCF!B{RVR["m1"]},"0.0%")&"，模型情境 B 為 "&TEXT(RevDCF!B{RVR["m0"]},"0.0%")&"；DCF 的永續價值占 EV "&TEXT(Val!C{VR["tvsh"]},"0%")&"。"')
put(RD, "A8", f'="重點觀察 3：FY2032E AI 淨效果（營業利益）A／B／C＝"&TEXT(Val!B{VR["ai_B"]},"#,##0.0")&"／"&TEXT(Val!C{VR["ai_B"]},"#,##0.0")&"／"&TEXT(Val!D{VR["ai_B"]},"#,##0.0")&" 十億日圓；情境 B 推論成本 FY2032E "&TEXT(PL!{CL[9]}{PR["B"]["inf"]},"0.0")&" 十億日圓（占營收 "&TEXT(PL!{CL[9]}{PR["B"]["inf"]}/PL!{CL[9]}{PR["B"]["rev"]},"0.00%")&"）。"')
put(RD, "A9", f'="重點觀察 4：FY2026E 模型營收 "&TEXT(Recon!C{RCR["rev"]},"#,##0")&" 十億日圓，對公司上修後指引 "&TEXT(Recon!D{RCR["rev"]},"#,##0")&" 為 "&TEXT(Recon!E{RCR["rev"]},"+0.0%;-0.0%")&"；Commerce 差距 "&TEXT(Recon!E{RCR["com"]},"+0.0%;-0.0%")&"（模型 Commerce 成長為 Assumed）。"')
put(RD, "A10", f'="重點觀察 5：收支關卡（FY2026–FY2030 FCFF 扣稅後利息與股利）A／B／C＝"&TEXT(Val!B{VR["gate_A"]},"#,##0")&"／"&TEXT(Val!B{VR["fcf_gate"]},"#,##0")&"／"&TEXT(Val!B{VR["gate_C"]},"#,##0")&" 十億日圓；WACC "&TEXT(Val!B{VR["wacc"]},"0.00%")&"，永續成長率 "&TEXT(Val!B{VR["g"]},"0.0%")&"（標準規格預設 3%，本模型改 2%，見下）。"')
put(RD, "A11", f'="橫幅："&Val!B{VR["rec"]}&"；加權目標價 "&TEXT(Val!B{VR["tp"]},"#,##0")&" 日圓（上檔空間 "&TEXT(Val!B{VR["up"]},"+0.0%;-0.0%")&"，現價 "&TEXT({I("px")},"#,##0.0")&"）。三情境 AI 淨效果分化（分流、價格提升、AI Mode、Agent i 付費）；同業倍數、β、查詢量為 Assumed。市場共識目標價 "&TEXT({I("cons_tp")},"0")&"（2026-09-08，已過期）。"')
put(RD, "A13", "情境分化依據：A／B／C 只在四個 AI 驅動分化——搜尋查詢分流占比（A 10%／B 20%／C 35%，Assumed）、AI 轉換率帶動的價格提升（15%／8%／2%）、官方帳號 AI Mode 增量（12%／6%／0%）、Agent i 付費占比（6%／3%／1%）。依據：公司只揭露 AI 功能使用者購前轉換 1.7 倍（非隨機對照）與 Agent i 計費起點，無量化 AI 營收；分化幅度為判斷，不是觀測值。其他分部成長三情境相同。")
put(RD, "A14", "與標準規格的偏離：(1) 永續成長率用 2%（標準預設 3%），因日本名目成長與 JGB 10 年期 2.89% 下 3% 偏高；(2) β 未取兩來源 Blume 調整中位數（Assumed）；(3) 評價日股價與資產負債表日不同：資產負債表取 2026-09-30，現價取 2026-11-16；(4) 非控制權益以帳面×1.5 倍（Assumed）；(5) 同業倍數未蒐集（分部加總倍數為 Assumed）；(6) 現金流含 PayPay 銀行，只以營收增量比率估計營運資金，未模擬存放款；(7) DCF 股數取流通股（含庫藏）×(1+稀釋係數)，未做庫藏股法。")
put(RD, "A15", "不適用：Tokenomics 自用 AI 讀法 2 與 Wattonomics 對照（LY 不自建算力，只向模型商購買推論，推論成本讀 Tokenomics 快照的市場價名稱）；三層協調（Coord 頁、C0、兩讀數）：非 Siliconomics 覆蓋；對帳中兩讀數目標價：不適用。")
put(RD, "A16", "已知限制：搜尋查詢量、ARPU、AI 營收、AI 推論成本金額、SBC、同業倍數、β、FY2026–FY2027 營收與 EPS 共識均未找到（見 data/research_lycorp.md 第 9 節）；FY2023–FY2024 的搜尋、帳號、展示拆分絕對金額未找到，歷史拆分只有 FY2025；分部口徑跨年有重編斷點。")
put(RD, "A18", "驅動 → 推導對照表（每頁輸入、推導量、因果方向）", BOLD)
for j, h in enumerate(["頁", "輸入", "推導量", "因果方向"]):
    put(RD, f"{L(j+1)}19", h, BOLD, HDR)
MAP = [("In→PL", "搜尋基礎成長、分流占比、價格提升", "搜尋廣告（基礎＋侵蝕＋增量）", "分流↑→搜尋營收↓；價格提升↑→營收↑"),
       ("In→PL", "帳號廣告成長、AI Mode 增量", "帳號廣告", "AI Mode↑→帳號廣告↑"),
       ("In→PL", "Agent i 使用者、付費占比、月費、廣告 ARPU", "Agent i 訂閱與廣告營收", "使用者×付費占比×月費"),
       ("In→PL", "Agent i 互動次數、任務倍數、token 價格（Tokenomics）、匯率", "推論成本（Agent i、搜尋 AI 回答）", "用量↑或單價↑→成本↑，計入營業成本"),
       ("PL→DCF", "營業利益、折舊、資本支出、營運資金", "FCFF 與現值", "FCFF＝NOPAT＋折舊−資本支出−營運資金增加"),
       ("DCF／PL→Val", "WACC、永續成長率、分部倍數", "DCF 目標價、分部加總目標價、加權目標價", "兩法依權重加權；評等依上檔空間與收支關卡"),
       ("Val→Sens", "每個有區間的輸入", "加權目標價相對基準", "一次改一個輸入，整本重算")]
for i_, rowv in enumerate(MAP):
    for j, v in enumerate(rowv):
        put(RD, f"{L(j+1)}{20+i_}", v)
put(RD, "A28", "市場資料日期、失效條件與風險登錄", BOLD)
put(RD, "A29", "市場資料日期：現價 2026-11-16 收盤（traders.co.jp）；JGB 10 年 2.89%（tradingeconomics.com，2026-11-16）；共識目標價 2026-09-08。")
put(RD, "A30", "主要失效條件", BOLD)
put(RD, "A31", "(1) 搜尋廣告 YoY 再度明顯低於 −5%（目前 Q2 −1.7%）而 AI 計費營收仍無揭露：情境 C 的分流假設需上調。(2) Agent i 付費方案價格與付費人數公布後與 Assumed 區間（500–1,500 日圓／月）顯著不同。(3) PayPay 與 ASKUL 利益率偏離指引。")
put(RD, "A32", "風險登錄", BOLD)
for j, h in enumerate(["風險", "是否建模", "理由", "觀察指標"]):
    put(RD, f"{L(j+1)}33", h, BOLD, HDR)
RISKS = [("AI 代理繞過 LY 平台（ChatGPT Apps、Gemini）", "是（情境分流）", "Yahoo!購物已對應 Apps in ChatGPT（2026-05-21）", "搜尋廣告 YoY、帳號廣告成長"),
         ("推論成本上升", "是", "讀 Tokenomics 市場價；高案 Sol 級", "Agent i 使用者與互動次數"),
         ("ASKUL 網路攻擊後遺與 Kakaku.com 收購", "否", "Kakaku.com TOB（約 690bn，與 Bain 各約 50%）未入模型；ASKUL 以 Commerce 成長帶過", "ASKUL 調整後 EBITDA、淨槓桿 1.94 倍"),
         ("PayPay 銀行與信用成本", "否", "FCFF 不含銀行存放款", "PayPay 銀行放款餘額、PayPay 調整後 EBITDA"),
         ("匯率", "只影響推論成本", "收入皆日圓", "USD／JPY")]
for i_, rowv in enumerate(RISKS):
    for j, v in enumerate(rowv):
        put(RD, f"{L(j+1)}{34+i_}", v)
put(RD, "A40", "工作表：README／In／Q／PL／DCF／Val／RevDCF／Recon／Sens／Sources。成品 dist/<日期>_LY個股模型_v0.1.0.xlsx 與 .html。")
for rr_ in (1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 13, 14, 15, 16, 29, 31):
    RD[f"A{rr_}"].alignment = Alignment(wrap_text=True, vertical="top")
RD.column_dimensions["A"].width = 120
RD.column_dimensions["B"].width = 36; RD.column_dimensions["C"].width = 36; RD.column_dimensions["D"].width = 36

# Sources
SO = S["Sources"]
put(SO, "A1", "Sources — 來源與版本", BOLD)
for i_, (a, b_) in enumerate([("事實總帳", "data/facts_lycorp.csv（267 列）、data/research_lycorp.md"), ("Tokenomics 快照", f"data/tokenomics_snapshot_v22.0.json（{TKV}；import_tokenomics.py --check 通過）"),
                              ("官方報告", "LY Corporation 決算短信（FY2021–FY2026 Q2）與簡報；URL 見 facts_lycorp.csv 出處欄"), ("現價與 JGB", "traders.co.jp、tradingeconomics.com（2026-11-16）"),
                              ("共識", "kabuyoho.jp（2026-09-08，過期）")]):
    put(SO, f"A{3+i_}", a); put(SO, f"B{3+i_}", b_)

# 輔助：匯出列號供 HTML 產生器使用
ROWMAP = {"IN": INR, "PR": PR, "IX": IX, "RID": RID, "VR": VR, "RVR": RVR, "RCR": RCR, "DR": DR, "SE_ROWS": SE_ROWS, "SE_BASE": SE_BASE, "SE_STALE": SE_STALE, "SE_STATUS": SE_STATUS,
          "SE_BID": SE_BID, "SE_END": SE_END, "SE_TORN": SE_TORN, "STDH": STDH, "STD_NOTE": STD_NOTE, "PLSUM": PLSUM, "HR": HR, "SCN_ROWS": SCN_ROWS, "SE_FP": SE_FP, "BUILD_ID": BUILD_ID,
          "SENS_DEF": [(k, s_, str(v)) for k, s_, v, _ in SENS_DEF], "IN_LAST": IN_LAST}
if os.environ.get("ROWMAP"):
    json.dump(ROWMAP, open(os.environ["ROWMAP"], "w", encoding="utf-8"), ensure_ascii=False, indent=1)
wb.save(OUT)
print("saved", OUT, BUILD_ID)
