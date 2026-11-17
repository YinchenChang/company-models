#!/usr/bin/env python3
"""Naver（035420）個股模型 v0.1：命題（原型 A，按廣告計價：查詢 x 廣告負載 x CTR x CPC）＋官方營收分部 → 損益 → DCF／分部加總 → 目標價。

用法：
  python3 software_ai/naver/build_naver.py <raw.xlsx> [--sens sens.json] [--set key=value ...]
  （--set 只改 In 頁基準欄 E，供 sens_naver.py 逐案整本重算用）
本版未做先驗（工作單 G 節）。金額＝十億韓元（K-IFRS，會計年度＝曆年），除非另註。歷史數字取自 facts_naver.py（同一份資料寫出 data/facts_naver.csv）。
"""
import datetime, hashlib, json, os, sys
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter as L

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
import facts_naver as FN  # noqa: E402

VER = "0.2.1"
PXDATE = datetime.date(2026, 11, 13)
VALDATE = datetime.date(2026, 9, 30)
DATADATE = datetime.date(2026, 11, 16)
BUILDDATE = datetime.date(2026, 11, 17)
FYE26 = datetime.date(2026, 12, 31)

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
TKI = SNAP["items"]
RM = json.load(open(os.path.join(HERE, "data", "relmult_naver.json"), encoding="utf-8"))
RMG = RM["groups"]["all"]
PSP = json.load(open(os.path.join(HERE, "data", "ps_peers_naver.json"), encoding="utf-8"))
tk_price = TKI["IF_PriceRef_Terra"]["values"]
tk_full = TKI["IF_FullCost_Terra"]["values"]["GB300 NVL72"]["基準"]
tk_def = TKI["IF_FullCostDefault_Terra"]["values"]["GB300 NVL72"]["基準"]
tk_tok = TKI["IF_TaskTokFresh"]["values"][0] + TKI["IF_TaskTokCached"]["values"][0] + TKI["IF_TaskTokDec"]["values"][0]
tk_mult2 = TKI["IF_TaskTokMult"]["values"][1]

wb = openpyxl.Workbook()
wb.remove(wb.active)
S = {n: wb.create_sheet(n) for n in ["README", "In", "Q", "Hist", "PL", "DCF", "CF", "Val", "RevDCF", "Recon", "Sens", "Sources"]}
YRS = ["FY2023A", "FY2024A", "FY2025A", "FY2026E", "FY2027E", "FY2028E", "FY2029E", "FY2030E", "FY2031E", "FY2032E"]
C0 = 3
CL = [L(C0 + i) for i in range(10)]  # C..L；j=0..9 ＝ FY2023..FY2032；j=3 ＝ FY2026E


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
INR = {}
IN_ROWS = []


def sec(title):
    IN_ROWS.append(("#", title))


def inp(key, label, unit, lo, base, hi, tag, src, note="", fmt=None):
    IN_ROWS.append((key, label, unit, lo, base, hi, tag, src, note, fmt))


def sg(key):  # 單點事實（facts_naver.SINGLE）
    return FN.SINGLE[key]


SNAPSRC = "data/tokenomics_snapshot_v22.0.json（Tokenomics v22.0，import_tokenomics.py --check 通過）"
D3 = FN.DECK["3Q26"]
sec("A. 日期、現價、股數與資產負債")
inp("data_date", "資料截止日（最新財報＝3Q26，2026-11-12 發布）", "日期", None, DATADATE, None, "Derived", "facts_naver.csv 擷取日")
inp("px_date", "股價基準日（現價日；最近一個交易日）", "日期", None, PXDATE, None, "Verified", "KRX 收盤（今日 2026-11-16 為週一，上一交易日為 11-13）")
inp("val_date", "評價日（資產負債表日；DCF 價值日）", "日期", None, VALDATE, None, "Derived", "3Q26 資產負債表日；NVIDIA 增資款（2026-11-06 繳款）依事件價另列入橋接")
inp("tp_date", "目標時點＝EDATE(股價基準日, 12)", "日期", None, "=EDATE(E5,12)", None, "Derived", "原則第 10 節第 26 條")
inp("fye26", "FY2026 會計年度結束日", "日期", None, FYE26, None, "Verified", "曆年制")
inp("px", "現價（2026-11-13 收盤）", "韓元", None, sg("px")[1], None, "Verified", sg("px")[5], "股價與當日跌幅取自單一券商頁並與教保報告隱含價相符；未三源核對", "#,##0")
inp("sh_out", "流通股數（扣庫藏，不含 NVIDIA 新股）", "百萬股", None, sg("sh")[1], None, "Derived", sg("sh")[5], "G2 第 4 條：法定申報封面股數未取得（未找到），暫用第二來源；2026-08-03 已註銷 4.9 百萬股（已反映於此數需另核）", "0.000")
inp("nv_sh", "NVIDIA 新發行股數（暫定）", "百萬股", None, sg("nv_sh")[1], None, "Verified", sg("nv_sh")[5], "2026-11-06 繳款、11-27 上市", "0.000000")
inp("nv_px", "NVIDIA 認購價（事件價）", "韓元", None, sg("nv_px")[1], None, "Verified", sg("nv_px")[5], "原則第 21 條：事件價不隨現價變動", "#,##0")
inp("dil", "稀釋係數（RSU 扣預期沒收、選擇權庫藏股法）", "%", 0.0, 0.005, 0.02, "Assumed", "RSU、選擇權明細未找到（G2 第 4 條要求庫藏股法，但資料缺）；ASR、可轉債：未發現", "以稀釋比例代替庫藏股法，列為偏離", "0.0%")
inp("sh_val", "評價日完全稀釋股數", "百萬股", None, "=(E7+E8)*(1+E10)", None, "Derived", "（流通股＋NVIDIA 新股）×（1＋稀釋係數）", "", "0.000")
inp("netcash0", "淨現金（公司定義，不含 NAVER Financial；2026-09-30）", "十億韓元", None, sg("netcash")[1], None, "Verified", D3, "G2 第 5 條：只計債務本金與融資租賃；租賃負債未找到、營業租賃不計", "#,##0.0")
inp("debt0", "借款（有息負債，2026-09-30）", "十億韓元", None, sg("debt")[1], None, "Verified", D3, "只用於資本結構權重與利息；淨現金已含其扣減", "#,##0.0")
inp("nci_v", "非控制權益（價值）", "十億韓元", 0, 0, 200, "Assumed", "非控制權益拆分未找到；暫為 0", "高案 200 為判斷", "#,##0")
inp("eqm_inc", "權益法投資年化收益（採計 NAVER 以外的投資組合）", "十億韓元", 50, 150, 300, "Assumed", "權益法損益 1Q25–3Q26 各季 −131.5 至 +276.2，無穩定基準（facts_naver.csv）；權益法投資帳面價值未找到", "同時作為預測期營業外收益（不入 EBIT）與非營運資產估值基礎", "#,##0")
inp("pe_inv", "權益法投資採用的本益比（類比：LY Corporation 前瞻 P/E）", "倍", 10, sg("pe_ly")[1], 22, "Analogy", sg("pe_ly")[5] + "（2026-10-13，逾期）", "類比，非同一資產；Naver 持有之被投資公司無獨立揭露", "0.0")

sec("B. 營收：Naver Platform 廣告（命題模組：查詢 x 廣告負載 x CTR x CPC）")
inp("q0", "搜尋查詢量 FY2026（十億次／年）", "十億次", 50, 120, 300, "Assumed", "公司未揭露查詢量（未找到）；只影響每查詢營收與推論成本的拆分，不影響校準年營收；推論成本與 q0 成正比", "", "0")
inp("gq", "查詢量年增（FY2027 起）", "%", 0.0, 0.025, 0.05, "Assumed", "無證據；韓國搜尋占有率 81.34%（Internet Trend）已高", "", "0.0%")
inp("al0", "每查詢廣告曝光數（FY2026，分解用）", "個", None, 1.0, None, "Assumed", "未揭露；僅作四因子分解，CPC 由校準反推（Derived）", "只作分解，校準年營收固定，不進敏感度", "0.00")
inp("ctr0", "廣告點擊率（FY2026，分解用）", "%", None, 0.025, None, "Assumed", "未揭露；同上", "", "0.0%")
inp("al_g", "每查詢廣告曝光數年增", "%", -0.01, 0.005, 0.02, "Assumed", "無證據", "", "0.0%")
inp("ctr_g", "點擊率年增", "%", -0.01, 0.005, 0.02, "Assumed", "無證據", "", "0.0%")
inp("cpc_g", "每次點擊價格（CPC）年增", "%", 0.0, 0.02, 0.05, "Assumed", "無證據；廣告成長 FY2026 約 7%（9M26 實績加 Q4 預測）", "", "0.0%")
inp("sigma", "AI 暴露廣告占 Naver Platform 廣告比例", "%", 0.6, 0.742, 0.85, "Derived", "（廣告 5,436.2 − 商務廣告 1,400.5）÷ 5,436.2（FY2025；商務廣告為舊分類，口徑不同故為近似）", "", "0.0%")
inp("g_q4ads", "廣告 Q4 26E 年增", "%", 0.03, 0.06, 0.09, "Assumed", "1Q26 +9.3%、2Q26 +7.5%、3Q26 +5.5%（facts_naver.csv，Derived）；公司無正式指引", "", "0.0%")
inp("s26", "FY2026 由 AI（AI 簡報、AI Tab）回答的查詢占比", "%", 0.10, 0.20, 0.35, "Assumed", "公司只揭露 AI 簡報月使用者逾 30 百萬、AI Tab 10 百萬（Interested-party）；占查詢比例未揭露", "", "0%")
inp("sai_A", "情境 A：FY2032 AI 回答占查詢比例", "%", None, 0.60, None, "Assumed", "無證據；A＝AI 簡報成為主要入口且變現順暢", "", "0%")
inp("sai_B", "情境 B：同上", "%", None, 0.50, None, "Assumed", "同上；基準", "", "0%")
inp("sai_C", "情境 C：同上", "%", None, 0.50, None, "Assumed", "同上；C＝占比相同但每則 AI 回答變現差", "", "0%")
inp("al_R", "AI 回答查詢的廣告負載相對值", "倍", 0.6, 0.8, 1.0, "Assumed", "AI 簡報版面擠壓廣告位，無揭露", "", "0.00")
inp("ctr_R", "AI 簡報廣告 CTR 相對值", "倍", 1.0, 1.30, 1.30, "Interested-party", FN.DECK["2Q26"] + "（公司稱 +30%，非隨機對照）", "來源誘因：管理層向市場說明 AI 成效；低案 1.0＝無提升", "0.00")
inp("cpc_R", "AI 回答查詢的每次點擊價格相對值", "倍", 0.9, 1.0, 1.1, "Assumed", "公司稱轉換 3 倍，但價格影響未揭露", "", "0.00")
inp("rai_A", "情境 A：AI 回答查詢的每查詢營收相對值 R", "倍", None, 1.15, None, "Assumed", "無證據；A＝AI 簡報的廣告變現高於傳統搜尋", "", "0.00")
inp("rai_B", "情境 B：R＝廣告負載 x CTR x CPC 相對值", "倍", None, "=E21*E22*E23", None, "Derived", "上列三項相乘（0.8 x 1.30 x 1.0）", "", "0.00")
inp("rai_C", "情境 C：同上", "倍", None, 0.85, None, "Assumed", "無證據；C＝AI 回答取代搜尋結果頁而廣告位被壓縮", "", "0.00")
inp("lam_A", "情境 A：FY2032 外部 AI（ChatGPT、Gemini）分流占查詢比例", "%", None, 0.0, None, "Assumed", "無證據；Naver 韓國搜尋占有率一年內由 59.26% 升至 81.34%（第三方統計）", "", "0%")
inp("lam_B", "情境 B：同上", "%", None, 0.03, None, "Assumed", "同上", "", "0%")
inp("lam_C", "情境 C：同上", "%", None, 0.08, None, "Assumed", "同上", "", "0%")
inp("oa27", "其他廣告（商務等非 AI 暴露）年增：FY2027", "%", 0.06, 0.12, 0.18, "Assumed", "商務廣告 2025 各季年增約 15% 至 19%（舊分類；facts_naver.csv）", "", "0.0%")
inp("oa32", "其他廣告年增：FY2032", "%", 0.02, 0.06, 0.09, "Assumed", "線性收斂；無證據", "", "0.0%")

sec("C. 營收：Services、Financial Platform、Global Opportunities")
inp("g_q4svc", "Services Q4 26E 年增", "%", 0.06, 0.10, 0.14, "Assumed", "3Q26 +13.1%；4Q25 基期偏高（434.9）", "", "0.0%")
inp("g_q4fin", "Financial Platform Q4 26E 年增", "%", 0.07, 0.11, 0.15, "Assumed", "3Q26 +14.1%；支付總額 3Q26 26.2 兆", "", "0.0%")
inp("g_q4c2c", "C2C Q4 26E 年增", "%", 0.15, 0.30, 0.45, "Assumed", "3Q26 +47.8%（FY2026 起 Poshmark 等併入後基期墊高），4Q25 基期 285.7 已高", "", "0.0%")
inp("g_q4cont", "Contents Q4 26E 年增", "%", -0.04, 0.0, 0.05, "Assumed", "3Q26 +5.9%、2Q26 +0.5%", "", "0.0%")
inp("g_q4ent", "Enterprise（不含 AI Factory）Q4 26E 年增", "%", 0.10, 0.20, 0.30, "Assumed", "3Q26 +33.0%（單季波動大）", "", "0.0%")
inp("svc27", "Services 年增：FY2027", "%", 0.08, 0.131, 0.16, "Derived", "以 3Q26 年增 13.1% 為起點（facts_naver.csv）", "", "0.0%")
inp("svc32", "Services 年增：FY2032", "%", 0.03, 0.06, 0.09, "Assumed", "線性收斂", "", "0.0%")
inp("fin27", "Financial Platform 年增：FY2027", "%", 0.08, 0.141, 0.18, "Derived", "以 3Q26 年增 14.1% 為起點", "", "0.0%")
inp("fin32", "Financial Platform 年增：FY2032", "%", 0.03, 0.07, 0.10, "Assumed", "線性收斂；NAVER Financial 與 Dunamu 合併未入模型", "", "0.0%")
inp("c2c27", "C2C 年增：FY2027", "%", 0.05, 0.15, 0.25, "Assumed", "FY2026 起 Poshmark 效應消化後的有機成長，無揭露", "", "0.0%")
inp("c2c32", "C2C 年增：FY2032", "%", 0.03, 0.08, 0.12, "Assumed", "線性收斂", "", "0.0%")
inp("cont27", "Contents 年增：FY2027", "%", -0.02, 0.03, 0.08, "Assumed", "近四季 Contents 營收持平（444.9／440.1／464.3／472.6）", "", "0.0%")
inp("cont32", "Contents 年增：FY2032", "%", -0.02, 0.03, 0.08, "Assumed", "持平於 FY2027 假設", "", "0.0%")
inp("ent27", "Enterprise（不含 AI Factory）年增：FY2027", "%", 0.05, 0.15, 0.25, "Assumed", "近四季波動大（150.5 至 193.3）", "", "0.0%")
inp("ent32", "Enterprise（不含 AI Factory）年增：FY2032", "%", 0.03, 0.08, 0.12, "Assumed", "線性收斂", "", "0.0%")

sec("D. AI Factory（算力出租，併入 Global Opportunities 的 Enterprise）")
inp("af_mw27", "年底上線容量 FY2027", "MW", None, sg("af_mw27")[1], None, "Interested-party", sg("af_mw27")[5], "來源誘因：管理層規劃", "0")
inp("af_mw28", "年底上線容量 FY2028", "MW", None, sg("af_mw28")[1], None, "Interested-party", sg("af_mw28")[5], "來源誘因：管理層規劃；整體目標 1GW 規模", "0")
inp("af_ext", "FY2029 起新增容量（MW）", "MW", 0, 0, 300, "Assumed", "無揭露；高案＝Brookfield 合作擴建", "", "0")
inp("af_rev", "每 MW 每年營收（十億韓元）", "十億韓元", 15, 22, 25, "Interested-party", sg("af_contract")[5] + "：約 2 兆韓元 ÷ 4.5 年 ÷ 20MW≈22", "來源誘因：管理層宣布合約，Naver 並保證使用費最長 8 年", "0.0")
inp("af_decl", "FY2029 起每 MW 營收年降幅", "%", 0.0, 0.02, 0.05, "Assumed", "算力租金隨世代更替下降；無揭露", "", "0.0%")
inp("af_m27", "AI Factory 營業利益率 FY2027", "%", 0.0, 0.03, 0.10, "Assumed", "公司稱 2027 年營業利益轉正（Interested-party）；利益率無揭露", "", "0.0%")
inp("af_m28", "AI Factory 營業利益率 FY2028 起", "%", 0.0, 0.08, 0.15, "Assumed", "無揭露；Naver 保證使用費，資產由 Asset Co 持有（預期合併，未入模型的資本支出與負債見 README）", "", "0.0%")

sec("E. 推論成本（只讀 Tokenomics 快照；不自行估 token 價格）")
inp("tok_base", "每次 AI 回答 token 數（一般聊天總 token＝新鮮輸入＋快取輸入＋decode）", "token", None, tk_tok, None, "Analogy", f"Tokenomics {TKV} IF_TaskTokFresh／IF_TaskTokCached／IF_TaskTokDec 第 1 項（{SNAPSRC}）", "G2 第 1 條不適用：Naver 未揭露 token 量與查詢量，無可校準的揭露；類比聊天基數，AI 簡報實際 token 未找到", "#,##0")
inp("tok_mult", "任務倍數（IF_TaskTokMult 第 1／2 項）", "倍", 1.0, 1.0, tk_mult2, "Analogy", f"Tokenomics {TKV} IF_TaskTokMult 第 2 項 {tk_mult2:.2f}（含思考的一般任務）", "對應哪個任務類別待 AI供應鏈 確認", "0.00")
inp("p1", "token 單價級距 1：全成本下限（IF_FullCost_Terra，GB300 基準）", "US$/M", None, tk_full, None, "Analogy", f"Tokenomics {TKV}；Terra 級對應 HyperCLOVA X 為類比，待 AI供應鏈 登錄", "模型商自營全成本，非市場價", "0.0000")
inp("p2", "token 單價級距 2：下游預設全成本（IF_FullCostDefault_Terra，GB300 基準）", "US$/M", None, tk_def, None, "Analogy", f"Tokenomics {TKV}；同上", "Naver 自建 HyperCLOVA X，採自行服務的全成本口徑（預設）", "0.0000")
inp("p3", "token 單價級距 3：市場價（IF_PriceRef_Terra）", "US$/M", None, tk_price, None, "Analogy", f"Tokenomics {TKV}；同上", "對外購買時的價格；Naver 自用不適用，僅作上界", "0.0000")
inp("ptok_idx", "採用的單價級距（1／2／3）", "級", 1, 2, 3, "Analogy", "選擇哪一層 token 價格屬類比（G2 第 1 條）", "高案（3）為只動成本一邊的上界，不含聯動", "0")
inp("r_link", "價格與成本聯動：單價級距每高於 2 一級，R 上調量", "倍", 0.0, 0.0, 0.08, "Assumed", "高價模型品質較佳換得較高變現（G2 第 2 條聯動案）；無證據", "只與單價級距 3 同時改才有效果，見 Sens 聯動案", "0.00")
inp("route", "模型組合：改走低價模型（級距 1）的查詢比例", "%", 0.0, 0.0, 0.5, "Assumed", "路由到便宜模型（G2 第 2 條模型組合案）", "", "0%")
inp("route_pen", "路由到低價模型的品質代價：R 下修量（全數改走時）", "倍", 0.05, 0.10, 0.20, "Assumed", "無證據", "只與路由比例同時改才有效果，見 Sens 組合案", "0.00")
inp("fx", "韓元兌美元", "韓元/US$", 1300, 1450, 1600, "Assumed", "未查證現匯；僅換算推論成本", "", "#,##0")

sec("F. 損益與現金流（比率參數基準＝持平於最近四季實績；收斂只放敏感度，G2 第 3 條）")
inp("m_np", "Naver Platform 營業利益率（近四季 4Q25–3Q26，營收加權）", "%", 0.24, "=Q!N{seg_oi_np}/Q!N{np}", 0.30, "Derived", "公司揭露各季利益率 33.3%／28.8%／28.5%／24.0% 乘營收（Q 頁）；低案＝3Q26 單季 24.0%", "基準持平於近四季（G3 第 1 條：歷史營業利益率無規模經濟趨勢，不延伸）；單季已降至 24.0% 的風險放在低案", "0.0%")
inp("m_fin", "Financial Platform 營業利益率（近四季；含未分攤調整的殘差）", "%", 0.02, "=Q!N{seg_oi_fin}/Q!N{fin}", 0.06, "Derived", "合併營業利益 − Naver Platform − Global Opportunities（Q 頁）；Financial 分部各季利益率只取得 3Q25 6.6%、3Q26 2.2%", "", "0.0%")
inp("m_glob", "Global Opportunities（不含 AI Factory）營業利益率（近四季）", "%", -0.05, "=Q!N{glob_oi}/Q!N{glob}", 0.02, "Derived", "分部損益 −31.2／−14.8／−35.6／−26.1 ÷ 近四季營收", "", "0.0%")
inp("cogs_pct", "營業成本率（Partner＋Infrastructure，近四季）", "%", 0.42, "=Q!N{cogs}/Q!N{rev}", 0.46, "Derived", "公司未揭露毛利；以 Partner 加 Infrastructure 為營業成本（模型分類，非公司口徑）；2026 起費用重分類", "", "0.0%")
inp("m_inc", "AI 增量營收的流入營業利益比（邊際）", "%", 0.50, 0.75, 0.90, "Assumed", "廣告增量邊際利益高；無揭露", "", "0%")
inp("m_lost", "AI 侵蝕（損失營收）的流出營業利益比（邊際）", "%", 0.70, 0.85, 0.95, "Assumed", "損失的搜尋廣告固定成本不減；無揭露", "", "0%")
inp("tax", "稅率（FY2025 有效稅率）", "%", 0.22, "=Q!L{tax}/Q!L{pbt}", 0.27, "Derived", "FY2025 所得稅 ÷ 稅前利益（Q 頁）", "", "0.0%")
inp("fin_inc", "預測期淨財務收益（淨現金加 NVIDIA 增資款 x 無風險利率）", "十億韓元", None, "=(E15+E8*E9/1000)*E{rf}", None, "Derived", "（淨現金＋新股 x 認購價）x 無風險利率；不入 EBIT", "", "#,##0.0")
inp("da25", "折舊攤銷 FY2025", "十億韓元", None, "=" + str(FN.SINGLE["ebitda25"][1]) + "-Q!L{oi}", None, "Derived", "第二來源 EBITDA 2,953（MarketScreener）減營業利益；公司未揭露 D&A（未找到）", "口徑不同，近似", "#,##0.0")
inp("capex25", "資本支出 FY2025（第二來源）", "十億韓元", None, sg("capex25")[1], None, "Derived", sg("capex25")[5], "公司簡報未揭露年度資本支出（未找到）", "#,##0")
inp("capex26", "資本支出 FY2026E", "十億韓元", 1455, 1800, 2400, "Assumed", "共識 1,455（財報前，Interested-party）；CFO 稱 GPU 單項 2026 起逾 2 兆韓元（2025-11）；3Q26 稱基礎設施投資 2026 年達高峰", "", "#,##0")
inp("capex_pct", "資本支出占營收（FY2027 起，持平）", "%", 0.07, 0.09, 0.13, "Assumed", "共識 FY2027E 資本支出 1,485 ÷ 營收 15,791＝9.4%（財報前，Interested-party）", "AI Factory 資產由 Asset Co 持有，其資本支出不在此", "0.0%")
inp("life", "折舊年限（GPU）", "年", None, sg("gpu_life")[1], None, "Verified", sg("gpu_life")[5], "D&A 隨前期資本支出以年限滾動", "0")
inp("wc_k", "營運資金增加占營收增量", "%", 0.0, 0.03, 0.08, "Assumed", "歷史比率未找到；金融事業營運資金不在 FCFF", "", "0%")
inp("dps0", "每股股利（FY2025 實績，FY2026 起以成長率推）", "韓元", None, sg("dps")[1], None, "Verified", sg("dps")[5], "", "#,##0")
inp("dps_g", "每股股利年增", "%", 0.0, 0.05, 0.10, "Assumed", "股東回饋 FY2025–27 為平均自由現金流 25% 至 35%（公司政策）", "", "0%")

sec("G. 折現率與估值規則")
inp("rf", "無風險利率（韓國 10 年期國債，2026-11-13 收盤）", "%", None, sg("rf")[1], None, "Verified", sg("rf")[5], "G4 第 3 條：同批模型統一取現價日前後同一日；頁面無 11-13 逐日值，由 11-16 讀數與單日變動回推", "0.00%")
inp("rf_date", "無風險利率基準日", "日期", None, datetime.date(2026, 11, 13), None, "Verified", sg("rf")[5], "寫入 In 頁（G4 第 3 條）")
inp("beta", "β（兩來源 1.19、1.20 平均後 Blume 調整）", "倍", 0.9, "=0.67*AVERAGE(1.19,1.2)+0.33", 1.4, "Derived", "StockAnalysis 1.19、Yahoo 1.20（facts_naver.csv）", "", "0.00")
inp("erp", "股權風險溢酬", "%", 0.04, 0.05, 0.06, "Assumed", "DCF 標準規格預設 5%", "", "0.0%")
inp("kd_sp", "負債成本相對無風險利率的利差", "%", 0.003, 0.006, 0.015, "Assumed", "未逐檔查證公司債利率", "", "0.0%")
inp("g_tv", "永續成長率", "%", 0.02, 0.03, 0.04, "Assumed", "DCF 標準規格預設 3%", "", "0.0%")
inp("rel_adj", "同業倍數相對調整 相對倍數（本公司 ÷ 同業；歷史 PS 與 PSG 兩讀數幾何平均）", "倍", RMG["sens"]["low"], RMG["geo_base"], RMG["sens"]["high"], "Derived", "software_ai/data/relmult_hist/out/naver.json（data/relmult_naver.json）：geo_base；" + RMG["base_source"] + "；規格 docs/plan/歷史相對倍數_PS與PSG.md", "G5 第 1 條：低／高＝sens.low／sens.high；Sens 另列只用 PS、只用 PSG 兩案。G6：Comparable 主腿為 P/S 口徑，與此相對值同口徑；前瞻本益比只作對照列", "0.000")
inp("sbc_pct", "股票薪酬占營收（只作對照，不進模型）", "%", 0.0, 0.0, 0.03, "Assumed", "未找到：股票薪酬金額未取得；0＝DCF（股票薪酬加回）讀數與標準 DCF 相同", "只作對照，不進模型；Recon 並列加回讀數（G4 第 2 條）", "0.0%")
inp("w_dcf", "加權目標價：DCF 權重", "%", 0.3, 0.5, 0.7, "Assumed", "預設 DCF 50%＋Comparable 50%（個股模型原則第 4 節）", "", "0%")
inp("tpw_A", "情境權重 A（樂觀）", "%", None, 0.25, None, "Assumed", "DCF 標準規格第 3 節預設 25／50／25", "", "0%")
inp("tpw_B", "情境權重 B（基準）", "%", None, 0.50, None, "Assumed", "同上", "", "0%")
inp("tpw_C", "情境權重 C（保守）", "%", None, 0.25, None, "Assumed", "同上", "", "0%")
inp("cons_tp", "市場共識平均目標價（財報前）", "韓元", None, sg("cons_tp")[1], None, "Interested-party", sg("cons_tp")[5], "來源誘因：賣方預估；只作對照，不進模型", "#,##0")

NAMES_SCN = {"sai_A", "sai_B", "sai_C", "rai_A", "rai_B", "rai_C", "lam_A", "lam_B", "lam_C", "tpw_A", "tpw_B", "tpw_C"}
SENS_SKIP = {"r_link", "route_pen"}
SENS_NOMODEL = {"sbc_pct"}   # 只在組合案中有效果，由組合案列涵蓋

put(IN, "A1", "In — 輸入（藍字；基準欄 E 進主模型；D／F＝敏感度低／高，只用於 Sens 頁，不進目標價）。標記＝Verified／Interested-party／Analogy／Assumed／Derived。本版未做先驗。", BOLD)
put(IN, "A2", "事實來源：facts_naver.py（寫出 data/facts_naver.csv）、data/research_naver.md；推論成本只讀 data/tokenomics_snapshot_v22.0.json（Tokenomics v22.0）。金額十億韓元，EPS 與股價韓元。曆年制，FY2026＝2026-01～2026-12。", None)
for j, h in enumerate(["鍵", "項目", "單位", "敏感度低", "基準", "敏感度高", "標記", "來源", "來源誘因／說明"]):
    put(IN, f"{L(j+1)}3", h, BOLD, HDR)
# 先定列號，再寫入（公式內有 Q 頁列號與 In 列號佔位）
r = 4
ROWNO = {}
for row in IN_ROWS:
    if row[0] == "#":
        r += 1
        continue
    ROWNO[row[0]] = r
    r += 1
INR = dict(ROWNO)
QROW = {}   # 於 Q 頁建立前先決定列號（下方 Q 頁用同一順序）
QKEYS_ORDER = ["rev", "np", "ads", "svc", "fin", "glob", "c2c", "cont", "ent", "devops", "partner", "infra", "mkt", "opex", "cogs", "oi", "seg_oi_np", "glob_oi", "seg_oi_fin", "nonop", "eqm", "pbt", "tax", "ni", "np_m", "tpv"]
for i, k in enumerate(QKEYS_ORDER):
    QROW[k] = 4 + i


def sub(v):
    if isinstance(v, str) and v.startswith("="):
        v = v.format(rf=INR["rf"], **QROW) if "{" in v else v
        for k in ("E5", "E7", "E8", "E9", "E10", "E15", "E21", "E22", "E23"):
            pass
    return v


# 固定列號的快捷公式（用實際列號重寫，避免硬碼錯誤）
def fix_formulas():
    IR = INR
    for row in IN_ROWS:
        pass


r = 4
IN_LAST = 0
for row in IN_ROWS:
    if row[0] == "#":
        put(IN, f"A{r}", row[1], BOLD)
        r += 1
        continue
    key, label, unit, lo, base, hi, tag, src, note, fmt = row
    put(IN, f"A{r}", key)
    put(IN, f"B{r}", label)
    put(IN, f"C{r}", unit)
    if isinstance(base, str) and base.startswith("="):
        base = base.format(rf=INR["rf"], **QROW) if "{" in base else base
        # 以鍵名替換位置佔位：把 E<固定列> 改為實際列號
        rep = {"E5": f"E{INR['px_date']}", "E7": f"E{INR['sh_out']}", "E8": f"E{INR['nv_sh']}", "E9": f"E{INR['nv_px']}", "E10": f"E{INR['dil']}", "E15": f"E{INR['netcash0']}",
               "E21": f"E{INR['al_R']}", "E22": f"E{INR['ctr_R']}", "E23": f"E{INR['cpc_R']}"}
        import re as _re
        base = _re.sub(r"E(5|7|8|9|10|15|21|22|23)(?![0-9])", lambda m: rep["E" + m.group(1)], base)
    if key in OVR:
        base = OVR[key]
    for col, v in (("D", lo), ("E", base), ("F", hi)):
        if v is None: continue
        isf = isinstance(v, str) and v.startswith("=")
        c = put(IN, f"{col}{r}", v, BLUE if col == "E" and not isf else None, YEL if tag == "Assumed" and col == "E" else None)
        if isinstance(v, datetime.date): c.number_format = "yyyy-mm-dd"
        elif fmt: c.number_format = fmt
    if key == "tp_date":
        IN[f"E{r}"].value = f"=EDATE(E{INR['px_date']},12)"
        IN[f"E{r}"].number_format = "yyyy-mm-dd"
    put(IN, f"G{r}", tag)
    put(IN, f"H{r}", src)
    put(IN, f"I{r}", note)
    IN_LAST = r
    r += 1
r += 1
put(IN, f"A{r}", "J. 情境驅動對照（三情境並排，進加權目標價；本節只引用上方各列，修改請改上方藍字）", BOLD)
r += 1
for j, h in enumerate(["鍵", "項目", "單位", "情境 A（樂觀）", "情境 B（基準）", "情境 C（保守）", "標記", "說明"]):
    put(IN, f"{L(j+1)}{r}", h, BOLD, HDR)
r += 1
SCN_ROWS = {}
for k, lab, unit, kA, kB, kC, fm in (("scn_sai", "FY2032 AI 回答占查詢比例", "%", "sai_A", "sai_B", "sai_C", "0%"),
                                      ("scn_rai", "AI 回答查詢的每查詢營收相對值 R", "倍", "rai_A", "rai_B", "rai_C", "0.00"),
                                      ("scn_lam", "FY2032 外部 AI 分流占查詢比例", "%", "lam_A", "lam_B", "lam_C", "0%"),
                                      ("scn_w", "情境權重", "%", "tpw_A", "tpw_B", "tpw_C", "0%")):
    SCN_ROWS[k] = r
    put(IN, f"A{r}", k)
    put(IN, f"B{r}", lab)
    put(IN, f"C{r}", unit)
    for col, kk in zip("DEF", (kA, kB, kC)):
        put(IN, f"{col}{r}", f"=E{INR[kk]}", None, None, fm)
    put(IN, f"G{r}", "Derived")
    put(IN, f"H{r}", f"引用 In 頁 {kA}／{kB}／{kC}；三情境只在 AI 驅動（AI 回答占比、每查詢營收相對值、外部分流）分化（見 README『情境分化依據』）")
    r += 1
IN.column_dimensions["A"].width = 14
IN.column_dimensions["B"].width = 60
IN.column_dimensions["H"].width = 60
IN.column_dimensions["I"].width = 50


def I(key):
    return f"In!$E${INR[key]}"


# =====================================================================================
# Q 頁
# =====================================================================================
Q = S["Q"]
put(Q, "A1", "Q — 歷史事實（Naver 財報簡報；逐列來源。十億韓元；各季口徑為 2026 年新分部與費用分類之重編；1Q25–4Q25 取自後續簡報比較欄）", BOLD)
put(Q, "A2", "L＝FY2025（1Q25–4Q25 加總）、M＝9M26 實績、N＝近四季（4Q25–3Q26）、O＝Q4 26E（In 頁成長率與近四季利益率）、P＝FY2026E（M＋O）。全文逐列見 data/facts_naver.csv。", None)
QH = ["鍵", "指標", "單位"] + FN.QS + ["FY2025", "9M26A", "近四季", "Q4 26E", "FY2026E", "標記", "來源", "來源誘因"]
for j, h in enumerate(QH):
    put(Q, f"{L(j+1)}3", h, BOLD, HDR)
QC = {q: L(4 + i) for i, q in enumerate(FN.QS)}   # D..K
cD, cE, cF, cG, cH, cI, cJ, cK = (L(4 + i) for i in range(8))
FLOW = set(QKEYS_ORDER) - {"np_m", "tpv"}
Q4G = {"ads": "g_q4ads", "svc": "g_q4svc", "fin": "g_q4fin", "c2c": "g_q4c2c", "cont": "g_q4cont", "ent": "g_q4ent"}
for k in QKEYS_ORDER:
    rr = QROW[k]
    put(Q, f"A{rr}", k)
    lab = FN.QLAB.get(k, {"cogs": "營業成本（模型分類：Partner＋Infrastructure）", "seg_oi_np": "Naver Platform 營業利益（＝營收 x 公司揭露利益率，Derived）", "seg_oi_fin": "Financial Platform 營業利益（殘差：合併 − Naver Platform − Global；含未分攤調整，Derived）"}.get(k, k))
    put(Q, f"B{rr}", lab)
    put(Q, f"C{rr}", FN.QUNIT.get(k, "十億韓元"))
    if k in FN.QDATA:
        for q, v in zip(FN.QS, FN.QDATA[k]):
            if v is not None:
                put(Q, f"{QC[q]}{rr}", v * 100 if k == "np_m" else v, BLUE, None, "0.0")
    tag = "Verified"
    if k in ("cogs", "seg_oi_np", "seg_oi_fin"):
        tag = "Derived"
        for q in FN.QS[1:]:
            c = QC[q]
            if k == "cogs": put(Q, f"{c}{rr}", f"={c}{QROW['partner']}+{c}{QROW['infra']}", None, None, "0.0")
            if k == "seg_oi_np": put(Q, f"{c}{rr}", f"={c}{QROW['np']}*{c}{QROW['np_m']}/100", None, None, "0.0")
            if k == "seg_oi_fin": put(Q, f"{c}{rr}", f"={c}{QROW['oi']}-{c}{QROW['seg_oi_np']}-{c}{QROW['glob_oi']}", None, None, "0.0")
    if k in FLOW:
        put(Q, f"L{rr}", f"=SUM({cE}{rr}:{cH}{rr})", None, None, "0.0")
        put(Q, f"M{rr}", f"=SUM({cI}{rr}:{cK}{rr})", None, None, "0.0")
        put(Q, f"N{rr}", f"=SUM({cH}{rr}:{cK}{rr})", None, None, "0.0")
        put(Q, f"P{rr}", f"=M{rr}+O{rr}", None, None, "0.0")
    if k in Q4G:
        put(Q, f"O{rr}", f"={cH}{rr}*(1+{I(Q4G[k])})", None, None, "0.0")
    put(Q, f"Q{rr}", tag if tag != "Verified" else ("Interested-party" if k == "np_m" else "Verified"))
    put(Q, f"R{rr}", FN.SRC_Q["3Q26"] + "（各季取該季簡報或後續簡報比較欄）" if tag == "Verified" else "由上列 Derived")
    if k == "np_m":
        put(Q, f"S{rr}", "公司自訂分部營業利益率；來源誘因：管理層自訂口徑")
# Q4E 合計列
R_ = QROW
put(Q, f"O{R_['ent']}", f"={cH}{R_['ent']}*(1+{I('g_q4ent')})", None, None, "0.0")
put(Q, f"O{R_['np']}", f"=O{R_['ads']}+O{R_['svc']}", None, None, "0.0")
put(Q, f"O{R_['glob']}", f"=O{R_['c2c']}+O{R_['cont']}+O{R_['ent']}", None, None, "0.0")
put(Q, f"O{R_['rev']}", f"=O{R_['np']}+O{R_['fin']}+O{R_['glob']}", None, None, "0.0")
put(Q, f"O{R_['seg_oi_np']}", f"={I('m_np')}*O{R_['np']}", None, None, "0.0")
put(Q, f"O{R_['seg_oi_fin']}", f"={I('m_fin')}*O{R_['fin']}", None, None, "0.0")
put(Q, f"O{R_['glob_oi']}", f"={I('m_glob')}*O{R_['glob']}", None, None, "0.0")
put(Q, f"O{R_['oi']}", f"=O{R_['seg_oi_np']}+O{R_['seg_oi_fin']}+O{R_['glob_oi']}", None, None, "0.0")
put(Q, f"O{R_['cogs']}", f"={I('cogs_pct')}*O{R_['rev']}", None, None, "0.0")
put(Q, f"O{R_['nonop']}", f"=({I('eqm_inc')}+{I('fin_inc')})/4", None, None, "0.0")
put(Q, f"O{R_['pbt']}", f"=O{R_['oi']}+O{R_['nonop']}", None, None, "0.0")
put(Q, f"O{R_['tax']}", f"=O{R_['pbt']}*{I('tax')}", None, None, "0.0")
put(Q, f"O{R_['ni']}", f"=O{R_['pbt']}-O{R_['tax']}", None, None, "0.0")
for k in ("devops", "partner", "infra", "mkt", "opex", "eqm"):
    for c in "OP": Q[f"{c}{R_[k]}"].value = None
# In 頁引用 Q 頁列號的公式佔位（{seg_oi_np} 等）已於上方 format 替換
QSR = QROW["tpv"] + 3
put(Q, f"A{QSR}", "單點事實（facts_naver.py 的 SINGLE；數值＝D 欄）", BOLD)
for j, h in enumerate(["鍵", "指標", "單位", "數值", "期間／日期", "口徑", "標記", "出處", "來源誘因"]):
    put(Q, f"{L(j+1)}{QSR+1}", h, BOLD, HDR)
QS1 = {}
rr = QSR + 2
for k, (m, v, u, p, b, s_, t, inc) in FN.SINGLE.items():
    QS1[k] = rr
    put(Q, f"A{rr}", k); put(Q, f"B{rr}", m); put(Q, f"C{rr}", u); put(Q, f"D{rr}", v, BLUE); put(Q, f"E{rr}", p); put(Q, f"F{rr}", b); put(Q, f"G{rr}", t); put(Q, f"H{rr}", s_); put(Q, f"I{rr}", inc)
    if k.startswith("pe_"):
        try:
            d = datetime.date.fromisoformat(p)
            put(Q, f"J{rr}", d, None, None, "yyyy-mm-dd")
        except ValueError:
            pass
    rr += 1
put(Q, f"A{rr+1}", "未找到清單（不估計）", BOLD)
for i_, nf in enumerate(FN.NOTFOUND):
    put(Q, f"A{rr+2+i_}", f"未找到：{nf}")
Q.column_dimensions["B"].width = 50

# =====================================================================================
# Hist 頁：舊五分部口徑年度（FY2021–FY2025）參考
# =====================================================================================
HS = S["Hist"]
put(HS, "A1", "Hist — 舊五分部口徑年度歷史（FY2021–FY2025；2026 年起分部改為三分部，下表為參考；FY2025 為 2026 年重編後之季度加總）", BOLD)
for j, h in enumerate(["項目", "單位"] + FN.YYR + ["標記", "來源"]):
    put(HS, f"{L(j+1)}3", h, BOLD, HDR)
HR_ = {}
rr = 4
for k, (lab, vals) in FN.YDATA.items():
    HR_[k] = rr
    put(HS, f"A{rr}", lab); put(HS, f"B{rr}", "十億韓元")
    for j, v in enumerate(vals):
        put(HS, f"{L(3+j)}{rr}", v, BLUE, None, "#,##0.0")
    put(HS, f"H{rr}", "Verified（FY2025 為 Derived：重編後季度加總）"); put(HS, f"I{rr}", "Naver 財報簡報 4Q22／4Q23／4Q24／4Q25；FY2021 淨利含 LINE-Z 合併利益")
    rr += 1
put(HS, f"A{rr}", "檢查：舊分部合計 − 合併營收（來源四捨五入）")
for j in range(5):
    c = L(3 + j)
    put(HS, f"{c}{rr}", "=" + "+".join(f"{c}{HR_[k]}" for k in ("seg_search", "seg_comm", "seg_fint", "seg_cont", "seg_cloud")) + f"-{c}{HR_['y_rev']}", None, None, "0.0")
rr += 1
for k, (lab, d) in FN.YOPX.items():
    HR_[k] = rr
    put(HS, f"A{rr}", lab); put(HS, f"B{rr}", "十億韓元")
    for y, v in d.items():
        put(HS, f"{L(3+FN.YYR.index(y))}{rr}", v, BLUE, None, "#,##0.0")
    put(HS, f"H{rr}", "Verified"); put(HS, f"I{rr}", "Naver 財報簡報 4Q23（FY2023）、4Q25（FY2024）")
    rr += 1
put(HS, f"A{rr}", "調整後 EBITDA（公司定義，不含股票薪酬；非 K-IFRS）"); put(HS, f"B{rr}", "十億韓元")
for y, v in FN.ADJ_EBITDA.items():
    put(HS, f"{L(3+FN.YYR.index(y))}{rr}", v, BLUE, None, "#,##0.0")
put(HS, f"H{rr}", "Interested-party"); put(HS, f"I{rr}", "管理層自訂口徑；FY2025 未取得")
HS.column_dimensions["A"].width = 52

# =====================================================================================
# PL 頁
# =====================================================================================
PL = S["PL"]
put(PL, "A1", "PL — 營收分部推導與損益（十億韓元；三情境 A／B／C；FY2023A–FY2025A 實績，FY2026E＝9M26 實績＋Q4 預測，FY2027E–FY2032E 預測）", BOLD)
put(PL, "A2", "命題模組（原型 A，按廣告計價）：Naver Platform 的 AI 暴露廣告＝查詢量 x（1 − 外部分流）x 每查詢營收 x（1 + AI 回答占比 x（R − 1））；每查詢營收＝廣告負載 x 點擊率 x 每次點擊價格。AI 淨效果＝增量（Q・s・(R−1)・u）− 侵蝕（−λ・Q・u）− 推論成本（Q・s・token・單價・匯率）。損益平衡 R*＝1 + 推論成本 ÷ 單位營收。Naver 自建 HyperCLOVA X，其自用推論成本在本模組，讀 Tokenomics 全成本（類比）。2026 年起新分部：FY2023–FY2024 無新口徑分部資料，分部列空白（舊口徑見 Hist 頁）。", None)
HRW = 4
put(PL, f"A{HRW}", "項目", BOLD, HDR); put(PL, f"B{HRW}", "單位", BOLD, HDR)
for j, y in enumerate(YRS):
    put(PL, f"{CL[j]}{HRW}", y, BOLD, HDR)
TR, SR_, CR = 5, 6, 7
put(PL, f"A{TR}", "年度序號 t（FY2026＝0）"); put(PL, f"B{TR}", "序號")
put(PL, f"A{SR_}", "AI 爬升比例（FY2026＝0，FY2032＝1）"); put(PL, f"B{SR_}", "比例")
put(PL, f"A{CR}", "成長率收斂比例（FY2027＝0，FY2032＝1）"); put(PL, f"B{CR}", "比例")
for j in range(10):
    put(PL, f"{CL[j]}{TR}", j - 3)
    put(PL, f"{CL[j]}{SR_}", f"=MAX(0,{CL[j]}{TR})/6", None, None, "0.000")
    put(PL, f"{CL[j]}{CR}", f"=MAX(0,{CL[j]}{TR}-1)/5", None, None, "0.000")


def path(a, b, col):
    return f"({I(a)}+({I(b)}-{I(a)})*{col}${CR})"


BLOCK = ["q", "lam", "sh", "reff", "al", "ctr", "cpc", "u", "base", "ero", "inc", "ads_exp", "chk_id", "oth_ads", "ads", "svc", "np", "fin", "c2c", "cont", "ent_core", "af_mw", "af", "ent", "glob", "rev",
         "chk_np", "chk_seg", "inf", "ai_net", "rstar", "dinc", "dinf", "ai_oi", "oi_np", "oi_fin", "oi_glob", "oi_af", "oi", "om", "cogs", "gp", "opx", "nonop", "pbt", "tax", "ni", "eps"]
LABEL = {"q": ("搜尋查詢量 Q（總，未扣外部分流）", "十億次"), "lam": ("外部 AI 分流占比 λ", "%"), "sh": ("AI 回答占查詢比例 s", "%"), "reff": ("AI 回答查詢的每查詢營收相對值 R（含聯動與路由）", "倍"),
         "al": ("每查詢廣告曝光數（廣告負載）", "個"), "ctr": ("廣告點擊率 CTR", "%"), "cpc": ("每次點擊價格 CPC（校準年由營收反推）", "韓元"), "u": ("每查詢營收 u＝廣告負載 x CTR x CPC", "韓元"),
         "base": ("基礎營收 Q・u（無 AI 差異、無侵蝕）", "十億韓元"), "ero": ("AI 侵蝕：外部 AI 分流損失 −λ・Q・u", "十億韓元"), "inc": ("AI 增量：(1−λ)・Q・u・s・(R−1)", "十億韓元"),
         "ads_exp": ("AI 暴露廣告（搜尋等）", "十億韓元"), "chk_id": ("檢查：AI 暴露廣告 − Q(1−λ)u(1+s(R−1))", "十億韓元"), "oth_ads": ("其他廣告（商務等非 AI 暴露）", "十億韓元"),
         "ads": ("廣告（Ads）", "十億韓元"), "svc": ("服務（Services）", "十億韓元"), "np": ("Naver Platform", "十億韓元"), "fin": ("Financial Platform", "十億韓元"),
         "c2c": ("C2C", "十億韓元"), "cont": ("Contents", "十億韓元"), "ent_core": ("Enterprise（不含 AI Factory）", "十億韓元"), "af_mw": ("AI Factory 年度平均上線容量", "MW"),
         "af": ("AI Factory 營收（模組，含於 Enterprise）", "十億韓元"), "ent": ("Enterprise", "十億韓元"), "glob": ("Global Opportunities", "十億韓元"), "rev": ("合併營收", "十億韓元"),
         "chk_np": ("檢查：子線合計 − 分部（來源四捨五入，|差| ≤ 0.1）", "十億韓元"), "chk_seg": ("檢查：分部合計 − 合併營收", "十億韓元"),
         "inf": ("推論成本（Naver 自用 HyperCLOVA X，讀 Tokenomics 全成本）", "十億韓元"), "ai_net": ("AI 淨效果（營收層：增量 + 侵蝕 − 推論成本）", "十億韓元"), "rstar": ("損益平衡 R*（增量剛好抵推論成本）", "倍"),
         "dinc": ("AI 增量相對 FY2026 校準年", "十億韓元"), "dinf": ("推論成本相對 FY2026 校準年", "十億韓元"),
         "ai_oi": ("AI 淨效果（營業利益層，相對 FY2026：增量 x 邊際 + 侵蝕 x 邊際 − 推論成本）", "十億韓元"),
         "oi_np": ("Naver Platform 營業利益", "十億韓元"), "oi_fin": ("Financial Platform 營業利益（含未分攤殘差）", "十億韓元"), "oi_glob": ("Global Opportunities 營業利益（不含 AI Factory）", "十億韓元"),
         "oi_af": ("AI Factory 營業利益", "十億韓元"), "oi": ("營業利益", "十億韓元"), "om": ("營業利益率", "%"), "cogs": ("營業成本（Partner＋Infrastructure，含推論成本增量）", "十億韓元"),
         "gp": ("毛利", "十億韓元"), "opx": ("營業費用（殘差＝毛利 − 營業利益）", "十億韓元"), "nonop": ("營業外淨額", "十億韓元"), "pbt": ("稅前利益", "十億韓元"), "tax": ("所得稅", "十億韓元"),
         "ni": ("稅後淨利（合計；歸屬母公司拆分未找到）", "十億韓元"), "eps": ("EPS（合計淨利 ÷ 完全稀釋股數，韓元；FY2026 含全年淨利）", "韓元")}
PCT = {"lam", "sh", "ctr", "om"}
SC = {"A": 0, "B": 1, "C": 2}
PR = {}
row = 9
for s in "ABC":
    put(PL, f"A{row}", f"情境 {s}（{'樂觀' if s=='A' else '基準' if s=='B' else '保守'}）", BOLD, HDR)
    row += 1
    rows = {}
    for k in BLOCK:
        rows[k] = row
        put(PL, f"A{row}", LABEL[k][0]); put(PL, f"B{row}", LABEL[k][1])
        row += 1
    PR[s] = rows
    row += 1
STD0 = row
scn = lambda s, key: f"In!$E${INR[key + '_' + s]}"
QP = lambda k: f"Q!$P${QROW[k]}"
QLc = lambda k: f"Q!$L${QROW[k]}"
tokx = f"({I('tok_base')}*{I('tok_mult')})"
pidx = f"MIN(3,MAX(1,{I('ptok_idx')}))"
pbase = f"CHOOSE({pidx},{I('p1')},{I('p2')},{I('p3')})"
peff = f"((1-{I('route')})*{pbase}+{I('route')}*{I('p1')})"
for s in "ABC":
    R = PR[s]
    for j in range(10):
        c = CL[j]; p = CL[j - 1] if j else None
        f = {}
        if j < 3:
            if j == 0:
                f["rev"] = f"=Hist!E{HR_['y_rev']}"; f["oi"] = f"=Hist!E{HR_['y_oi']}"
                f["cogs"] = f"=Hist!E{HR_['y_partner']}+Hist!E{HR_['y_infra']}"
                f["nonop"] = f"=Hist!E{HR_['y_nonop']}"; f["tax"] = f"=Hist!E{HR_['y_tax']}"
            elif j == 1:
                f["rev"] = f"=Hist!F{HR_['y_rev']}"; f["oi"] = f"=Hist!F{HR_['y_oi']}"
                f["cogs"] = f"=Hist!F{HR_['y_partner']}+Hist!F{HR_['y_infra']}"
                f["nonop"] = f"=Hist!F{HR_['y_nonop']}"; f["tax"] = f"=Hist!F{HR_['y_tax']}"
            else:
                f["rev"] = f"=Hist!G{HR_['y_rev']}"; f["oi"] = f"={QLc('oi')}"
                f["cogs"] = f"={QLc('cogs')}"; f["nonop"] = f"={QLc('nonop')}"; f["tax"] = f"={QLc('tax')}"
                f["ads"] = f"={QLc('ads')}"; f["svc"] = f"={QLc('svc')}"; f["np"] = f"={QLc('np')}"; f["fin"] = f"={QLc('fin')}"
                f["c2c"] = f"={QLc('c2c')}"; f["cont"] = f"={QLc('cont')}"; f["ent"] = f"={QLc('ent')}"; f["glob"] = f"={QLc('glob')}"
                f["chk_np"] = f"={c}{R['ads']}+{c}{R['svc']}-{c}{R['np']}"
                f["chk_seg"] = f"={c}{R['np']}+{c}{R['fin']}+{c}{R['glob']}-{c}{R['rev']}"
            f["gp"] = f"={c}{R['rev']}-{c}{R['cogs']}"; f["opx"] = f"={c}{R['gp']}-{c}{R['oi']}"
            f["pbt"] = f"={c}{R['oi']}+{c}{R['nonop']}"; f["ni"] = f"={c}{R['pbt']}-{c}{R['tax']}"; f["om"] = f"={c}{R['oi']}/{c}{R['rev']}"
        else:
            first = j == 3
            f["q"] = f"={I('q0')}" if first else f"={p}{R['q']}*(1+{I('gq')})"
            f["lam"] = 0 if first else f"={scn(s,'lam')}*{c}${SR_}"
            f["sh"] = f"={I('s26')}+({scn(s,'sai')}-{I('s26')})*{c}${SR_}"
            f["reff"] = f"={scn(s,'rai')}+{I('r_link')}*MAX(0,{pidx}-2)-{I('route_pen')}*{I('route')}"
            if first:
                f["ads_exp"] = f"={I('sigma')}*{QP('ads')}"
                f["u"] = f"={c}{R['ads_exp']}/({c}{R['q']}*(1+{c}{R['sh']}*({c}{R['reff']}-1)))"
                f["al"] = f"={I('al0')}"; f["ctr"] = f"={I('ctr0')}"; f["cpc"] = f"={c}{R['u']}/({c}{R['al']}*{c}{R['ctr']})"
                f["oth_ads"] = f"=(1-{I('sigma')})*{QP('ads')}"
                f["svc"] = f"={QP('svc')}"; f["fin"] = f"={QP('fin')}"; f["c2c"] = f"={QP('c2c')}"; f["cont"] = f"={QP('cont')}"; f["ent_core"] = f"={QP('ent')}"
                f["af_mw"] = 0; f["af"] = 0
                f["dinc"] = 0; f["dinf"] = 0
            else:
                f["al"] = f"={p}{R['al']}*(1+{I('al_g')})"; f["ctr"] = f"={p}{R['ctr']}*(1+{I('ctr_g')})"; f["cpc"] = f"={p}{R['cpc']}*(1+{I('cpc_g')})"
                f["u"] = f"={c}{R['al']}*{c}{R['ctr']}*{c}{R['cpc']}"
                f["ads_exp"] = f"={c}{R['base']}+{c}{R['ero']}+{c}{R['inc']}"
                f["oth_ads"] = f"={p}{R['oth_ads']}*(1+{path('oa27','oa32',c)})"
                f["svc"] = f"={p}{R['svc']}*(1+{path('svc27','svc32',c)})"; f["fin"] = f"={p}{R['fin']}*(1+{path('fin27','fin32',c)})"
                f["c2c"] = f"={p}{R['c2c']}*(1+{path('c2c27','c2c32',c)})"; f["cont"] = f"={p}{R['cont']}*(1+{path('cont27','cont32',c)})"
                f["ent_core"] = f"={p}{R['ent_core']}*(1+{path('ent27','ent32',c)})"
                mw_end = {4: I("af_mw27"), 5: I("af_mw28")}
                if j == 4: f["af_mw"] = f"={I('af_mw27')}/2"
                elif j == 5: f["af_mw"] = f"=({I('af_mw27')}+{I('af_mw28')})/2"
                else: f["af_mw"] = f"={I('af_mw28')}+{I('af_ext')}/2" if j == 6 else f"={I('af_mw28')}+{I('af_ext')}"
                f["af"] = f"={c}{R['af_mw']}*{I('af_rev')}" if j <= 5 else f"={c}{R['af_mw']}*{I('af_rev')}*(1-{I('af_decl')})^{j-5}"
                f["dinc"] = f"={c}{R['inc']}-$F{R['inc']}"
                f["dinf"] = f"={c}{R['inf']}-$F{R['inf']}"
            f["base"] = f"={c}{R['q']}*{c}{R['u']}"
            f["ero"] = f"=-{c}{R['lam']}*{c}{R['base']}"
            f["inc"] = f"=(1-{c}{R['lam']})*{c}{R['base']}*{c}{R['sh']}*({c}{R['reff']}-1)"
            f["chk_id"] = f"={c}{R['ads_exp']}-{c}{R['q']}*(1-{c}{R['lam']})*{c}{R['u']}*(1+{c}{R['sh']}*({c}{R['reff']}-1))"
            f["ads"] = f"={c}{R['ads_exp']}+{c}{R['oth_ads']}"
            f["np"] = f"={c}{R['ads']}+{c}{R['svc']}"
            f["ent"] = f"={c}{R['ent_core']}+{c}{R['af']}"
            f["glob"] = f"={c}{R['c2c']}+{c}{R['cont']}+{c}{R['ent']}"
            f["rev"] = f"={c}{R['np']}+{c}{R['fin']}+{c}{R['glob']}"
            f["chk_np"] = f"={c}{R['ads']}+{c}{R['svc']}-{c}{R['np']}"
            f["chk_seg"] = f"={c}{R['np']}+{c}{R['fin']}+{c}{R['glob']}-{c}{R['rev']}"
            f["inf"] = f"={c}{R['q']}*(1-{c}{R['lam']})*{c}{R['sh']}*{tokx}*{peff}*{I('fx')}/1000000"
            f["ai_net"] = f"={c}{R['inc']}+{c}{R['ero']}-{c}{R['inf']}"
            f["rstar"] = f"=IF({c}{R['sh']}*{c}{R['base']}*(1-{c}{R['lam']})=0,0,1+{c}{R['inf']}/({c}{R['sh']}*{c}{R['base']}*(1-{c}{R['lam']})))"
            f["ai_oi"] = f"={I('m_inc')}*{c}{R['dinc']}+{I('m_lost')}*{c}{R['ero']}-{c}{R['dinf']}"
            if first:
                f["oi_np"] = f"={QP('seg_oi_np')}"; f["oi_fin"] = f"={QP('seg_oi_fin')}"; f["oi_glob"] = f"={QP('glob_oi')}"; f["oi_af"] = 0
                f["cogs"] = f"={QP('cogs')}"; f["nonop"] = f"={QP('nonop')}"; f["tax"] = f"={QP('tax')}"
            else:
                f["oi_np"] = f"={I('m_np')}*({c}{R['np']}-{c}{R['dinc']}-{c}{R['ero']})+{I('m_inc')}*{c}{R['dinc']}+{I('m_lost')}*{c}{R['ero']}-{c}{R['dinf']}"
                f["oi_fin"] = f"={I('m_fin')}*{c}{R['fin']}"
                f["oi_glob"] = f"={I('m_glob')}*({c}{R['glob']}-{c}{R['af']})"
                f["oi_af"] = f"={c}{R['af']}*{I('af_m27') if j == 4 else I('af_m28')}"
                f["cogs"] = f"={I('cogs_pct')}*{c}{R['rev']}+{c}{R['dinf']}"
                f["nonop"] = f"={I('eqm_inc')}+{I('fin_inc')}"
                f["tax"] = f"={c}{R['pbt']}*{I('tax')}"
            f["oi"] = f"={c}{R['oi_np']}+{c}{R['oi_fin']}+{c}{R['oi_glob']}+{c}{R['oi_af']}"
            f["om"] = f"={c}{R['oi']}/{c}{R['rev']}"
            f["gp"] = f"={c}{R['rev']}-{c}{R['cogs']}"; f["opx"] = f"={c}{R['gp']}-{c}{R['oi']}"
            f["pbt"] = f"={c}{R['oi']}+{c}{R['nonop']}"; f["ni"] = f"={c}{R['pbt']}-{c}{R['tax']}"
            f["eps"] = f"={c}{R['ni']}*1000/{I('sh_val')}"
        for k, v in f.items():
            cell = put(PL, f"{c}{R[k]}", v)
            cell.number_format = "0.0%" if k in PCT else ("0.000" if k in ("reff", "rstar", "al") else ("#,##0" if k in ("eps", "u", "cpc") else "#,##0.0"))

# 標準損益表（情境 B）
row = STD0
put(PL, f"A{row}", "標準損益表（情境 B；K-IFRS 合併，十億韓元，EPS 韓元）", BOLD, HDR)
row += 1
put(PL, f"A{row}", "項目", BOLD, HDR)
for j, y in enumerate(YRS):
    put(PL, f"{CL[j]}{row}", y, BOLD, HDR)
STDH = row
row += 1
RB = PR["B"]
IX = {}
std_def = [("rev", "營收", "十億韓元", lambda c, j: f"={c}{RB['rev']}"), ("yoy", "營收年增率", "%", lambda c, j: (f"={c}{{rev}}/{CL[j-1]}{{rev}}-1" if j > 0 else None)),
           ("cogs", "營業成本（COGS）", "十億韓元", lambda c, j: f"={c}{RB['cogs']}"), ("gp", "毛利", "十億韓元", lambda c, j: f"={c}{RB['gp']}"), ("gm", "毛利率", "%", lambda c, j: f"={c}{{gp}}/{c}{{rev}}"),
           ("opx", "營業費用", "十億韓元", lambda c, j: f"={c}{RB['opx']}"), ("oi", "營業利益", "十億韓元", lambda c, j: f"={c}{RB['oi']}"), ("opm", "營業利益率", "%", lambda c, j: f"={c}{{oi}}/{c}{{rev}}"),
           ("oth", "其他收入（營業外淨額）", "十億韓元", lambda c, j: f"={c}{RB['nonop']}"), ("pti", "稅前利益", "十億韓元", lambda c, j: f"={c}{RB['pbt']}"), ("tax", "所得稅", "十億韓元", lambda c, j: f"={c}{RB['tax']}"),
           ("ni", "稅後淨利（合計）", "十億韓元", lambda c, j: f"={c}{RB['ni']}"), ("nm", "淨利率", "%", lambda c, j: f"={c}{{ni}}/{c}{{rev}}"),
           ("eps", "EPS（合計淨利 ÷ 完全稀釋股數；FY2023A–FY2025A 未找到）", "韓元", lambda c, j: (f"={c}{RB['eps']}" if j >= 3 else None))]
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
        cell = put(PL, f"{c}{rr_}", v.format(**IX))
        cell.number_format = "0.0%" if unit == "%" else ("#,##0" if k == "eps" else "#,##0.0")
put(PL, f"A{row}", "比率（表尾）", BOLD)
row += 1
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
put(PL, f"A{row}", "標準損益表口徑：K-IFRS（GAAP 口徑，非 Non-GAAP 調整後）。Naver 不揭露毛利；本表營業成本＝Partner＋Infrastructure（模型分類，非公司口徑），營業費用為殘差；FY2023A–FY2024A 為舊費用分類、FY2025A 起為 2026 年重編分類，兩者不連續。FY2026E＝9M26 實績＋Q4 預測。預測期營業外淨額＝權益法收益＋淨財務收益；歸屬母公司拆分未找到，EPS 以合計淨利計。")
row += 2
PLSUM = row
put(PL, f"A{row}", "分部對帳（三情境 × 全部年度；差額＝分部合計 − 合併營收，十億韓元）", BOLD)
row += 1
for s in "ABC":
    put(PL, f"A{row}", f"情境 {s}：分部合計 − 合併營收（絕對值最大者）")
    put(PL, f"C{row}", "=MAX(" + ",".join(f"ABS({c}{PR[s]['chk_seg']})" for c in CL[2:]) + ")", None, None, "0.000")
    row += 1
for s in "ABC":
    put(PL, f"A{row}", f"情境 {s}：AI 恆等式（AI 暴露廣告）差額（絕對值最大者）")
    put(PL, f"C{row}", "=MAX(" + ",".join(f"ABS({c}{PR[s]['chk_id']})" for c in CL[3:]) + ")", None, None, "0.000")
    row += 1
for s in "ABC":
    put(PL, f"A{row}", f"情境 {s}：子線合計 − 分部（絕對值最大者，來源四捨五入）")
    put(PL, f"C{row}", "=MAX(" + f"ABS({CL[2]}{PR[s]['chk_np']}),ABS({CL[3]}{PR[s]['chk_np']})" + ")", None, None, "0.000")
    row += 1
PL.column_dimensions["A"].width = 66

# =====================================================================================
# DCF 頁
# =====================================================================================
DC = S["DCF"]
put(DC, "A1", "DCF — 無槓桿自由現金流（FCFF＝NOPAT＋折舊攤銷 − 資本支出 − 營運資金增加；股票薪酬不加回，金額未找到）", BOLD)
put(DC, "A2", "金額十億韓元。期中折現；FY2026 只計評價日（2026-09-30）之後的 Q4 stub（占比＝評價日至年底天數 ÷ 365）。折舊攤銷隨前期資本支出滾動：D&A(t)＝D&A(t−1)＋(資本支出(t−1)−D&A(t−1))÷折舊年限。AI Factory 資產由 Asset Co 持有，其資本支出與負債不在本表（資金缺口風險見 README）。", None)
put(DC, "A4", "項目", BOLD, HDR); put(DC, "B4", "單位", BOLD, HDR)
for j, y in enumerate(YRS):
    put(DC, f"{CL[j]}4", y, BOLD, HDR)
put(DC, "A5", "期間占比（FY2026＝評價日至年底；其餘 1）")
put(DC, "A6", "折現期間（年，期中）")
put(DC, "A7", "折現因子（WACC）")
for j in range(3, 10):
    c = CL[j]
    if j == 3:
        put(DC, f"{c}5", f"=({I('fye26')}-{I('val_date')})/365", None, None, "0.000")
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
    for k, lab in (("oi", "營業利益（EBIT；不含權益法損益）"), ("nopat", "NOPAT＝EBIT ×（1 − 稅率）"), ("da", "＋折舊攤銷"), ("cx", "− 資本支出"), ("wc", "− 營運資金增加"), ("fcff", "FCFF"),
                   ("pv", "FCFF 現值（含 stub 占比）"), ("int", "稅後淨利息（借款 x 稅前負債成本 x (1 − 稅率)）"), ("div", "股利（每股股利 x 股數）"), ("gap", "收支差：FCFF − 稅後利息 − 股利")):
        rows[k] = row
        put(DC, f"A{row}", lab)
        row += 1
    DR[s] = rows
    R = PR[s]
    put(DC, f"{CL[2]}{rows['da']}", f"={I('da25')}", None, None, "#,##0.0")
    put(DC, f"{CL[2]}{rows['cx']}", f"={I('capex25')}", None, None, "#,##0.0")
    for j in range(3, 10):
        c = CL[j]; p = CL[j - 1]
        put(DC, f"{c}{rows['oi']}", f"=PL!{c}{R['oi']}", None, None, "#,##0.0")
        put(DC, f"{c}{rows['nopat']}", f"={c}{rows['oi']}*(1-{I('tax')})", None, None, "#,##0.0")
        put(DC, f"{c}{rows['da']}", f"={p}{rows['da']}+({p}{rows['cx']}-{p}{rows['da']})/{I('life')}", None, None, "#,##0.0")
        put(DC, f"{c}{rows['cx']}", f"={I('capex26')}" if j == 3 else f"=PL!{c}{R['rev']}*{I('capex_pct')}", None, None, "#,##0.0")
        put(DC, f"{c}{rows['wc']}", f"=(PL!{c}{R['rev']}-PL!{p}{R['rev']})*{I('wc_k')}", None, None, "#,##0.0")
        put(DC, f"{c}{rows['fcff']}", f"={c}{rows['nopat']}+{c}{rows['da']}-{c}{rows['cx']}-{c}{rows['wc']}", None, None, "#,##0.0")
        put(DC, f"{c}{rows['pv']}", f"={c}{rows['fcff']}*{c}$5*{c}$7", None, None, "#,##0.0")
        put(DC, f"{c}{rows['int']}", f"={I('debt0')}*(Val!$B$KDROW)*(1-{I('tax')})", None, None, "#,##0.0")
        put(DC, f"{c}{rows['div']}", f"={I('dps0')}*(1+{I('dps_g')})^{j-2}*{I('sh_out')}/1000", None, None, "#,##0.0")
        put(DC, f"{c}{rows['gap']}", f"={c}{rows['fcff']}-{c}{rows['int']}-{c}{rows['div']}", None, None, "#,##0.0")
    row += 1
DC.column_dimensions["A"].width = 60

# =====================================================================================
# CF 頁：收支關卡
# =====================================================================================
CF = S["CF"]
put(CF, "A1", "CF — 收支關卡（個股模型原則第 10 節第 63 條）：FY2026–FY2030 五年 FCFF − 稅後利息 − 股利 合計；三情境並排，B 為基準", BOLD)
put(CF, "A2", "FCFF 取 DCF 頁（全年，不扣 stub）；股利＝每股股利 x 流通股數；AI Factory 資產由 Asset Co 持有，其資本支出與負債不在本表。")
for j, h in enumerate(["項目", "情境 A", "情境 B", "情境 C"]):
    put(CF, f"{L(j+1)}4", h, BOLD, HDR)
CFR = {}
rr = 5
for k, lab in (("fcff", "FY2026–FY2030 FCFF 合計"), ("int", "FY2026–FY2030 稅後利息 合計"), ("div", "FY2026–FY2030 股利 合計")):
    CFR[k] = rr
    put(CF, f"A{rr}", lab)
    for s, c in zip("ABC", "BCD"):
        put(CF, f"{c}{rr}", f"=SUM(DCF!{CL[3]}{DR[s][k]}:{CL[7]}{DR[s][k]})", None, None, "#,##0.0")
    rr += 1
CFR["gate"] = rr
put(CF, f"A{rr}", "收支關卡 A／B／C（FY2026–FY2030 FCFF − 稅後利息 − 股利 合計，十億韓元）", BOLD)
for s, c in zip("ABC", "BCD"):
    put(CF, f"{c}{rr}", f"={c}{CFR['fcff']}-{c}{CFR['int']}-{c}{CFR['div']}", None, None, "#,##0.0")
CF.column_dimensions["A"].width = 70

# =====================================================================================
# Val 頁
# =====================================================================================
V = S["Val"]
VR = {}
put(V, "A1", "Val — 目標價（股價基準日＋12 個月）＝DCF（情境加權）x 權重 ＋ Comparable（同業中位 P/S x 相對倍數 x TTM 營收，情境加權）x 權重", BOLD)
put(V, "A2", "DCF 為標準 FCFF 版（個股DCF標準規格 v1.0），評價日 2026-09-30，推到目標時點；Comparable 主腿（G6）以同業中位 P/S 乘相對倍數乘目標時點 TTM 營收（FY2026E 與 FY2027E 依已過天數加權）；分部 NTM NOPAT 乘前瞻本益比只作對照列。五級評等：> +50% 強力買進、> +25% 買進、≥ −25% 中立、≥ −50% 放空、其餘強力放空；先看收支關卡。", None)
r = 4
put(V, f"A{r}", "0.3 結論", BOLD); r += 1
for j, h in enumerate(["項目", "數值", "對照", "對照", "說明"]):
    put(V, f"{L(j+1)}{r}", h, BOLD, HDR)
r += 1
for k in ("tpdcf", "tpcomp", "tp", "px", "cons", "up", "upc", "tier", "val_rating", "gate_val", "rec"):
    VR[k] = r; r += 1
r += 1
put(V, f"A{r}", "0.2 折現率", BOLD); r += 1
for k, lab in (("rf", "無風險利率"), ("rf_flag", "無風險利率基準日 vs 現價基準日（相差超過 2 個交易日標逾期）"), ("beta", "β"), ("erp", "股權風險溢酬"), ("ke", "股權成本 ke"), ("kd", "稅前負債成本（無風險利率＋利差）"), ("kdat", "稅後負債成本"),
               ("emkt", "股權市值（十億韓元，含 NVIDIA 新股）"), ("dmkt", "借款（十億韓元）"), ("we", "股權權重"), ("wacc", "WACC"), ("ronic", "RONIC（＝WACC，新投資價值中性）"), ("g", "永續成長率")):
    VR[k] = r; put(V, f"A{r}", lab); r += 1
r += 1
put(V, f"A{r}", "0.1 DCF 彙總（三情境）", BOLD); r += 1
VR["dh"] = r
for j, h in enumerate(["項目", "情境 A", "情境 B", "情境 C"]):
    put(V, f"{L(j+1)}{r}", h, BOLD, HDR)
r += 1
for k, lab in (("pvx", "明確期 FCFF 現值（含 stub）"), ("nop33", "FY2033 NOPAT＝FY2032 NOPAT x (1 + g)"), ("tv", "永續價值＝NOPAT(FY2033) x (1 − g／RONIC)／(WACC − g)"), ("pvtv", "永續價值現值"), ("ev", "企業價值 EV"),
               ("tvsh", "永續價值占 EV"), ("ncash", "＋淨現金（公司定義，不含 NAVER Financial；2026-09-30）"), ("nvp", "＋NVIDIA 增資款（事件價 x 新股數）"), ("eqm", "＋權益法投資價值（年化收益 x 本益比，類比）"),
               ("nci", "− 非控制權益"), ("eqv", "股權價值"), ("shv", "每股價值（韓元，評價日）"), ("nonop_ps", "其中非營運資產每股（淨現金、增資款、權益法投資 − 非控制權益）"), ("dcf12", "DCF 情境價值推到 12 個月後（x (1+ke) − 股利）")):
    VR[k] = r; put(V, f"A{r}", lab); r += 1
r += 1
put(V, f"A{r}", "0.0 Comparable 主腿：同業中位 P/S x 相對倍數 x 目標時點 TTM 營收 ÷ 股數（G6，與 PS／PSG 相對值同口徑）；其下前瞻本益比分部加總只作對照", BOLD); r += 1
for k, lab in (("ps_med", "同業中位 P/S（Val 同業表各家最近一季，季末日逐家標示；5 家，不做迴歸）"), ("ps_x", "主腿倍數＝同業中位 P/S x 相對倍數（歷史 PS 與 PSG 幾何平均）"), ("ttm_el", "目標日在 FY2027 內已過比例（TTM＝比例 x FY2027E ＋ (1 − 比例) x FY2026E）"),
               ("ntm_w", "對照：NTM 中 FY2027E 權重 w＝(FY2027 年底 − 目標日) ÷ 365（NTM＝w x FY2027E ＋ (1 − w) x FY2028E）"), ("pe_plat", "Platforms 群組前瞻本益比（GOOGL、META、PINS 中位數；3 家，不足 8 家，不做迴歸）"),
               ("pe_kj", "Korea／Japan 群組前瞻本益比（Kakao、LY 中位數；2 家，不做迴歸）"), ("pe_all", "全部同業前瞻本益比中位數"), ("rel", "相對倍數（歷史 PS 與 PSG 幾何平均，Derived；前瞻本益比對照列不套用）")):
    VR[k] = r; put(V, f"A{r}", lab); r += 1
VR["sh"] = r
for j, h in enumerate(["項目", "情境 A", "情境 B", "情境 C", "同業群組倍數"]):
    put(V, f"{L(j+1)}{r}", h, BOLD, HDR)
r += 1
for k, lab in (("ttm_t", "目標時點 TTM 營收（十億韓元）"), ("ps_eq", "主腿股權價值＝主腿倍數 x TTM 營收（十億韓元）"), ("ps_ps", "主腿每股價值（韓元，目標時點）＝Comparable 主腿"),
               ("n_np", "對照：Naver Platform NTM NOPAT（十億韓元）"), ("n_fin", "Financial Platform NTM NOPAT"), ("n_glob", "Global Opportunities（含 AI Factory）NTM NOPAT（負值不計價，下限 0）"),
               ("v_np", "Naver Platform 價值（Platforms 群組倍數）"), ("v_fin", "Financial Platform 價值（Korea／Japan 群組倍數）"), ("v_glob", "Global Opportunities 價值（Korea／Japan 群組倍數）"),
               ("v_op", "營運事業價值（分部加總）"), ("c_nonop", "＋非營運資產（淨現金、增資款、權益法投資 − 非控制權益）"), ("c_eq", "股權價值"), ("s_ps", "對照：前瞻本益比口徑每股價值（相對倍數 1.00；韓元，目標時點；不進目標價）"),
               ("ntm_tot", "NTM NOPAT 合計"), ("ntm_mult", "對照：前瞻本益比口徑隱含營運倍數（營運事業價值 ÷ NTM NOPAT 合計）")):
    VR[k] = r; put(V, f"A{r}", lab); r += 1
r += 1
put(V, f"A{r}", "同業逐家：P/S（主腿；TTM 營收口徑，季末日）與前瞻本益比（對照；NTM EPS 口徑，分母期間＝NTM）", BOLD); r += 1
VR["peh"] = r
for j, h in enumerate(["同業", "前瞻本益比", "基準日", "會計準則與租賃處理", "與現價日相差", "群組", "P/S（TTM）", "P/S 季末日", "P/S 逾期"]):
    put(V, f"{L(j+1)}{r}", h, BOLD, HDR)
r += 1
PEERS = [("pe_goog", "Alphabet（GOOGL）", "US GAAP（營業租賃費用化，租賃不計入 EBIT 以下利息）", "Platforms"), ("pe_meta", "Meta（META）", "US GAAP（同上）", "Platforms"), ("pe_pins", "Pinterest（PINS）", "US GAAP（同上）", "Platforms"),
         ("pe_kakao", "Kakao（035720）", "K-IFRS（租賃資本化：折舊與利息分列，盈餘口徑與 US GAAP 略有差異）", "Korea／Japan"), ("pe_ly", "LY Corporation（4689）", "IFRS（租賃資本化，同上）", "Korea／Japan")]
VR["peers"] = {}
for k, nm, acct, grp in PEERS:
    VR["peers"][k] = r
    put(V, f"A{r}", nm); put(V, f"B{r}", f"=Q!$D${QS1[k]}", None, None, "0.00"); put(V, f"C{r}", f"=Q!$J${QS1[k]}", None, None, "yyyy-mm-dd"); put(V, f"D{r}", acct)
    put(V, f"E{r}", f'=IF(ABS(NETWORKDAYS(MIN(C{r},{I("px_date")}),MAX(C{r},{I("px_date")}))-1)>2,"逾期","")'); put(V, f"F{r}", grp)
    pk = {"pe_goog": "GOOGL", "pe_meta": "META", "pe_pins": "PINS", "pe_kakao": "KAKAO", "pe_ly": "LYCORP"}[k]
    put(V, f"G{r}", PSP[pk]["ps"], BLUE, None, "0.00"); put(V, f"H{r}", datetime.date.fromisoformat(PSP[pk]["q"]), None, None, "yyyy-mm-dd")
    put(V, f"I{r}", f'=IF(ABS(NETWORKDAYS(MIN(H{r},{I("px_date")}),MAX(H{r},{I("px_date")}))-1)>2,"逾期（季末資料）","")')
    r += 1
VR["peer_end"] = r - 1
r += 1
put(V, f"A{r}", "本公司現行倍數 vs 同業（G2 第 6 條，第 47 條）", BOLD); r += 1
for k, lab in (("own_pe", "本公司現行前瞻本益比（現價 ÷ 情境 B 的 NTM EPS，分母期間＝NTM）"), ("own_ratio", "本公司現行倍數 ÷ 同業中位數"), ("own_rerate", "隱含重估（Comparable 目標價 ÷ 現價 − 1）"),
               ("rm_geo", "歷史相對倍數基準 相對倍數（PS 與 PSG 兩讀數幾何平均；Derived，歷史相對倍數資料庫）"), ("rm_ps", "　PS 讀數（歷史相對中位）"), ("rm_psg", "　PSG 讀數（歷史相對中位 x 成長比）"),
               ("rm_g", "　成長比（本公司 ÷ 同業中位的 TTM 營收年增）"), ("rm_cg", "　本公司 TTM 營收年增（2026-06-30）"), ("rm_pg", "　同業中位 TTM 營收年增"), ("rm_div", "　兩讀數差異（相對）"),
               ("comp_pe1", "對照：前瞻本益比口徑 Comparable 目標價（相對倍數 1.00；不進目標價）"), ("comp_alt", "對照：前瞻本益比口徑 Comparable 目標價（相對倍數取本公司現行 ÷ 同業中位；不進目標價）"), ("tp_alt", "診斷對照：加權目標價（同上；不進目標價）")):
    VR[k] = r; put(V, f"A{r}", lab); r += 1
r += 1
put(V, f"A{r}", "DCF 與 Comparable 落差分解（情境 B，每股，韓元；G2 第 8 條）", BOLD); r += 1
for k, lab in (("gp_gap", "落差＝Comparable − DCF（目標時點）"), ("gp_sbc", "其中股票薪酬（金額未找到，差 0）"), ("gp_lev", "其中槓桿（兩法皆為股權價值口徑，差 0）"),
               ("gp_time", "其中時點（DCF 由評價日推到目標時點並扣股利的差；P/S 主腿直接在目標時點）"), ("gp_mult", "其中倍數與口徑（P/S 主腿股權價值 − DCF 評價日股權價值）"), ("gp_chk", "檢查：分項合計 − 落差"),
               ("gp_dcfmult", "DCF 隱含 P/S（DCF 目標時點股權價值 ÷ TTM 營收）")):
    VR[k] = r; put(V, f"A{r}", lab); r += 1
r += 1
put(V, f"A{r}", "情境目標價", BOLD); r += 1
for k, lab in (("t_A", "情境 A 目標價"), ("t_B", "情境 B 目標價"), ("t_C", "情境 C 目標價"), ("wsum", "情境權重合計"), ("per", "加權目標價隱含 FY2027E 本益比（分母＝情境 B 模型）"),
               ("gate_A", "收支合計 A（十億韓元，CF 頁）"), ("gate_B", "收支合計 B"), ("gate_C", "收支合計 C")):
    VR[k] = r; put(V, f"A{r}", lab); r += 1
VR["ai_B"] = r; put(V, f"A{r}", "AI 淨效果（營業利益，相對 FY2026）FY2032E：A／B／C（十億韓元）"); r += 1
VR["ai_net"] = r; put(V, f"A{r}", "AI 淨效果（營收層，增量 + 侵蝕 − 推論成本）FY2032E：A／B／C（十億韓元）"); r += 1
VR["ai_tp"] = r; put(V, f"A{r}", "AI 淨效果 FY2032E 營業利益：情境 A − 情境 C（十億韓元）"); r += 1

put(V, f"B{VR['rf']}", f"={I('rf')}", None, None, "0.00%")
put(V, f"B{VR['rf_flag']}", f'=IF(ABS(NETWORKDAYS(MIN({I("rf_date")},{I("px_date")}),MAX({I("rf_date")},{I("px_date")}))-1)>2,"逾期","未逾期（取值日 2026-11-13）")')
put(V, f"B{VR['beta']}", f"={I('beta')}", None, None, "0.00")
put(V, f"B{VR['erp']}", f"={I('erp')}", None, None, "0.0%")
put(V, f"B{VR['ke']}", f"=B{VR['rf']}+B{VR['beta']}*B{VR['erp']}", None, None, "0.00%")
put(V, f"B{VR['kd']}", f"=B{VR['rf']}+{I('kd_sp')}", None, None, "0.00%")
put(V, f"B{VR['kdat']}", f"=B{VR['kd']}*(1-{I('tax')})", None, None, "0.00%")
put(V, f"B{VR['emkt']}", f"={I('px')}*{I('sh_val')}/1000", None, None, "#,##0")
put(V, f"B{VR['dmkt']}", f"={I('debt0')}", None, None, "#,##0")
put(V, f"B{VR['we']}", f"=B{VR['emkt']}/(B{VR['emkt']}+B{VR['dmkt']})", None, None, "0.0%")
put(V, f"B{VR['wacc']}", f"=B{VR['we']}*B{VR['ke']}+(1-B{VR['we']})*B{VR['kdat']}", None, None, "0.00%")
put(V, f"B{VR['ronic']}", f"=B{VR['wacc']}", None, None, "0.00%")
put(V, f"B{VR['g']}", f"={I('g_tv')}", None, None, "0.0%")
for j in range(3, 10):
    DC[f"{CL[j]}7"].value = DC[f"{CL[j]}7"].value.replace("WACCROW", str(VR["wacc"]))
    for s in "ABC":
        c_ = DC[f"{CL[j]}{DR[s]['int']}"]
        c_.value = c_.value.replace("KDROW", str(VR["kdat"] if False else VR["kd"]))
SCOL = {"A": "B", "B": "C", "C": "D"}
Tyrs = f"(({I('tp_date')}-{I('val_date')})/365)"
w_ntm = f"B{VR['ntm_w']}"
put(V, f"B{VR['ntm_w']}", f"=(DATE(2027,12,31)-{I('tp_date')})/365", None, None, "0.000")
pr_ = VR["peers"]
put(V, f"B{VR['pe_plat']}", f"=MEDIAN(B{pr_['pe_goog']}:B{pr_['pe_pins']})", None, None, "0.00")
put(V, f"B{VR['pe_kj']}", f"=MEDIAN(B{pr_['pe_kakao']}:B{pr_['pe_ly']})", None, None, "0.00")
put(V, f"B{VR['pe_all']}", f"=MEDIAN(B{VR['peers']['pe_goog']}:B{VR['peer_end']})", None, None, "0.00")
put(V, f"B{VR['rel']}", f"={I('rel_adj')}", None, None, "0.00")
put(V, f"E{VR['v_np']}", f"=B{VR['pe_plat']}", None, None, "0.00")
put(V, f"E{VR['v_fin']}", f"=B{VR['pe_kj']}", None, None, "0.00")
put(V, f"E{VR['v_glob']}", f"=B{VR['pe_kj']}", None, None, "0.00")
put(V, f"B{VR['ps_med']}", f"=MEDIAN(G{VR['peers']['pe_goog']}:G{VR['peer_end']})", None, None, "0.00")
put(V, f"B{VR['ps_x']}", f"=B{VR['ps_med']}*B{VR['rel']}", None, None, "0.000")
put(V, f"B{VR['ttm_el']}", f"=({I('tp_date')}-DATE(2026,12,31))/365", None, None, "0.000")
nonop_sum = f"({I('netcash0')}+{I('nv_sh')}*{I('nv_px')}/1000+{I('eqm_inc')}*{I('pe_inv')}-{I('nci_v')})"
for s in "ABC":
    c = SCOL[s]; D_ = DR[s]; R = PR[s]
    put(V, f"{c}{VR['pvx']}", f"=SUM(DCF!{CL[3]}{D_['pv']}:{CL[9]}{D_['pv']})", None, None, "#,##0.0")
    put(V, f"{c}{VR['nop33']}", f"=DCF!{CL[9]}{D_['nopat']}*(1+$B${VR['g']})", None, None, "#,##0.0")
    put(V, f"{c}{VR['tv']}", f"={c}{VR['nop33']}*(1-$B${VR['g']}/$B${VR['ronic']})/($B${VR['wacc']}-$B${VR['g']})", None, None, "#,##0.0")
    put(V, f"{c}{VR['pvtv']}", f"={c}{VR['tv']}*DCF!{CL[9]}7*(1+$B${VR['wacc']})^(-0.5)", None, None, "#,##0.0")
    put(V, f"{c}{VR['ev']}", f"={c}{VR['pvx']}+{c}{VR['pvtv']}", None, None, "#,##0.0")
    put(V, f"{c}{VR['tvsh']}", f"={c}{VR['pvtv']}/{c}{VR['ev']}", None, None, "0.0%")
    put(V, f"{c}{VR['ncash']}", f"={I('netcash0')}", None, None, "#,##0.0")
    put(V, f"{c}{VR['nvp']}", f"={I('nv_sh')}*{I('nv_px')}/1000", None, None, "#,##0.0")
    put(V, f"{c}{VR['eqm']}", f"={I('eqm_inc')}*{I('pe_inv')}", None, None, "#,##0.0")
    put(V, f"{c}{VR['nci']}", f"=-{I('nci_v')}", None, None, "#,##0.0")
    put(V, f"{c}{VR['eqv']}", f"={c}{VR['ev']}+{c}{VR['ncash']}+{c}{VR['nvp']}+{c}{VR['eqm']}+{c}{VR['nci']}", None, None, "#,##0.0")
    put(V, f"{c}{VR['shv']}", f"={c}{VR['eqv']}*1000/{I('sh_val')}", None, None, "#,##0")
    put(V, f"{c}{VR['nonop_ps']}", f"=({c}{VR['ncash']}+{c}{VR['nvp']}+{c}{VR['eqm']}+{c}{VR['nci']})*1000/{I('sh_val')}", None, None, "#,##0")
    put(V, f"{c}{VR['dcf12']}", f"={c}{VR['shv']}*(1+$B${VR['ke']})^{Tyrs}-{I('dps0')}*{Tyrs}", None, None, "#,##0")
    # Comparable
    def ntm(k): return f"({w_ntm}*PL!{CL[4]}{R[k]}+(1-{w_ntm})*PL!{CL[5]}{R[k]})"
    put(V, f"{c}{VR['n_np']}", f"={ntm('oi_np')}*(1-{I('tax')})", None, None, "#,##0.0")
    put(V, f"{c}{VR['n_fin']}", f"={ntm('oi_fin')}*(1-{I('tax')})", None, None, "#,##0.0")
    put(V, f"{c}{VR['n_glob']}", f"=({ntm('oi_glob')}+{ntm('oi_af')})*(1-{I('tax')})", None, None, "#,##0.0")
    put(V, f"{c}{VR['v_np']}", f"={c}{VR['n_np']}*$E${VR['v_np']}", None, None, "#,##0")
    put(V, f"{c}{VR['v_fin']}", f"={c}{VR['n_fin']}*$E${VR['v_fin']}", None, None, "#,##0")
    put(V, f"{c}{VR['v_glob']}", f"=MAX(0,{c}{VR['n_glob']})*$E${VR['v_glob']}", None, None, "#,##0")
    put(V, f"{c}{VR['v_op']}", f"={c}{VR['v_np']}+{c}{VR['v_fin']}+{c}{VR['v_glob']}", None, None, "#,##0")
    put(V, f"{c}{VR['c_nonop']}", f"={nonop_sum}", None, None, "#,##0")
    put(V, f"{c}{VR['c_eq']}", f"={c}{VR['v_op']}+{c}{VR['c_nonop']}", None, None, "#,##0")
    put(V, f"{c}{VR['s_ps']}", f"={c}{VR['c_eq']}*1000/{I('sh_val')}", None, None, "#,##0")
    put(V, f"{c}{VR['ttm_t']}", f"=$B${VR['ttm_el']}*PL!{CL[4]}{R['rev']}+(1-$B${VR['ttm_el']})*PL!{CL[3]}{R['rev']}", None, None, "#,##0.0")
    put(V, f"{c}{VR['ps_eq']}", f"=$B${VR['ps_x']}*{c}{VR['ttm_t']}", None, None, "#,##0")
    put(V, f"{c}{VR['ps_ps']}", f"={c}{VR['ps_eq']}*1000/{I('sh_val')}", None, None, "#,##0")
    put(V, f"{c}{VR['ntm_tot']}", f"={c}{VR['n_np']}+{c}{VR['n_fin']}+{c}{VR['n_glob']}", None, None, "#,##0.0")
    put(V, f"{c}{VR['ntm_mult']}", f"={c}{VR['v_op']}/{c}{VR['ntm_tot']}", None, None, "0.0")
    put(V, f"B{VR['t_' + s]}", f"={I('w_dcf')}*{c}{VR['dcf12']}+(1-{I('w_dcf')})*{c}{VR['ps_ps']}", None, None, "#,##0")
    put(V, f"B{VR['gate_' + s]}", f"=CF!{c}{CFR['gate']}", None, None, "#,##0.0")
    put(V, f"{c}{VR['ai_B']}", f"=PL!{CL[9]}{R['ai_oi']}", None, None, "#,##0.0")
    put(V, f"{c}{VR['ai_net']}", f"=PL!{CL[9]}{R['ai_net']}", None, None, "#,##0.0")
# 落差分解（情境 B＝欄 C）
put(V, f"B{VR['gp_gap']}", f"=C{VR['ps_ps']}-C{VR['dcf12']}", None, None, "#,##0")
put(V, f"B{VR['gp_sbc']}", 0, None, None, "#,##0")
put(V, f"B{VR['gp_lev']}", 0, None, None, "#,##0")
put(V, f"B{VR['gp_time']}", f"=-(C{VR['dcf12']}-C{VR['shv']})", None, None, "#,##0")
put(V, f"B{VR['gp_mult']}", f"=C{VR['ps_ps']}-C{VR['shv']}", None, None, "#,##0")
put(V, f"B{VR['gp_chk']}", f"=B{VR['gp_sbc']}+B{VR['gp_lev']}+B{VR['gp_time']}+B{VR['gp_mult']}-B{VR['gp_gap']}", None, None, "0.000")
put(V, f"B{VR['gp_dcfmult']}", f"=C{VR['dcf12']}*{I('sh_val')}/1000/C{VR['ttm_t']}", None, None, "0.00")
# 本公司現行倍數
put(V, f"B{VR['own_pe']}", f"={I('px')}/(({w_ntm}*PL!{CL[4]}{PR['B']['eps']})+(1-{w_ntm})*PL!{CL[5]}{PR['B']['eps']})", None, None, "0.0")
put(V, f"B{VR['own_ratio']}", f"=B{VR['own_pe']}/B{VR['pe_all']}", None, None, "0.00")
put(V, f"B{VR['own_rerate']}", f"=B{VR['tpcomp']}/B{VR['px']}-1", None, None, "0.0%")
for k_, v_, f_ in (("rm_geo", RMG["geo_base"], "0.000"), ("rm_ps", RMG["ps_reading"], "0.000"), ("rm_psg", RMG["psg_reading"], "0.000"), ("rm_g", RMG["growth_ratio"]["value"], "0.000"),
                   ("rm_cg", RMG["growth_ratio"]["company_g"], "0.0%"), ("rm_pg", RMG["growth_ratio"]["peer_median_g"], "0.0%"), ("rm_div", RMG["divergence"], "0.0%")):
    put(V, f"B{VR[k_]}", v_, BLUE, None, f_)
put(V, f"B{VR['comp_alt']}", "=" + "+".join(f"{I('tpw_'+s_)}*(({SCOL[s_]}{VR['v_op']}*B{VR['own_ratio']}+{SCOL[s_]}{VR['c_nonop']})*1000/{I('sh_val')})" for s_ in "ABC"), None, None, "#,##0")
put(V, f"B{VR['tp_alt']}", f"={I('w_dcf')}*B{VR['tpdcf']}+(1-{I('w_dcf')})*B{VR['comp_alt']}", None, None, "#,##0")
put(V, f"B{VR['wsum']}", f"={I('tpw_A')}+{I('tpw_B')}+{I('tpw_C')}", None, None, "0%")
put(V, f"B{VR['per']}", f"=B{VR['tp']}/PL!{CL[4]}{PR['B']['eps']}", None, None, "0.0")
put(V, f"B{VR['tpdcf']}", f"={I('tpw_A')}*B{VR['dcf12']}+{I('tpw_B')}*C{VR['dcf12']}+{I('tpw_C')}*D{VR['dcf12']}", None, None, "#,##0")
put(V, f"B{VR['tpcomp']}", f"={I('tpw_A')}*B{VR['ps_ps']}+{I('tpw_B')}*C{VR['ps_ps']}+{I('tpw_C')}*D{VR['ps_ps']}", None, None, "#,##0")
put(V, f"B{VR['comp_pe1']}", f"={I('tpw_A')}*B{VR['s_ps']}+{I('tpw_B')}*C{VR['s_ps']}+{I('tpw_C')}*D{VR['s_ps']}", None, None, "#,##0")
put(V, f"B{VR['tp']}", f"={I('w_dcf')}*B{VR['tpdcf']}+(1-{I('w_dcf')})*B{VR['tpcomp']}", None, None, "#,##0")
put(V, f"B{VR['px']}", f"={I('px')}", None, None, "#,##0")
put(V, f"B{VR['cons']}", f"={I('cons_tp')}", None, None, "#,##0")
put(V, f"B{VR['up']}", f"=B{VR['tp']}/B{VR['px']}-1", None, None, "0.0%")
put(V, f"B{VR['upc']}", f"=B{VR['tpcomp']}/B{VR['px']}-1", None, None, "0.0%")
put(V, f"B{VR['tier']}", f'=IF(B{VR["up"]}>0.5,"> +50%",IF(B{VR["up"]}>0.25,"+25% 至 +50%",IF(B{VR["up"]}>=-0.25,"−25% 至 +25%",IF(B{VR["up"]}>=-0.5,"−50% 至 −25%","< −50%"))))')
put(V, f"B{VR['val_rating']}", f'=IF(B{VR["up"]}>0.5,"強力買進",IF(B{VR["up"]}>0.25,"買進",IF(B{VR["up"]}>=-0.25,"中立",IF(B{VR["up"]}>=-0.5,"放空","強力放空"))))')
put(V, f"B{VR['gate_val']}", f"=B{VR['gate_B']}", None, None, "#,##0.0")
put(V, f"B{VR['rec']}", f'=IF(AND(B{VR["gate_val"]}<0,OR(B{VR["val_rating"]}="買進",B{VR["val_rating"]}="強力買進")),"收支不足",B{VR["val_rating"]})')
notes = {"tpdcf": "三情境 DCF 推到目標時點（扣期間股利），依權重 25／50／25 加權", "tpcomp": "同業中位 P/S x 相對倍數 x 目標時點 TTM 營收 ÷ 稀釋股數；情境加權（G6 主腿）",
         "tp": "DCF 與 Comparable 依 In 頁權重加權", "px": "2026-11-13 收盤（最近交易日）", "cons": "MarketScreener 26 位分析師，2026-11-11，財報前；只作對照，財報後各家目標價已下修（Recon）",
         "up": "", "upc": "", "tier": "五級固定規則", "val_rating": "僅依上檔空間級距", "gate_val": "同 CF 頁收支關卡 B（FY2026–FY2030 FCFF − 稅後利息 − 股利）", "rec": "收支關卡不通過且估值評等為買進以上時為收支不足"}
labs = {"tpdcf": "DCF 目標價（情境加權）", "tpcomp": "Comparable 目標價（P/S 主腿，情境加權）", "tp": "目標價（DCF 與 Comparable 加權）", "px": "現價（2026-11-13 收盤）", "cons": "市場共識目標價（財報前，對照）",
        "up": "上檔空間：加權目標價", "upc": "上檔空間：Comparable 目標價", "tier": "上檔空間級距", "val_rating": "估值評等（依加權目標價上檔空間）", "gate_val": "入不敷出判定用收支合計（十億韓元，FY2026–FY2030）", "rec": "投資建議"}
for k in labs:
    put(V, f"A{VR[k]}", labs[k]); put(V, f"E{VR[k]}", notes[k])
put(V, f"B{VR['ai_tp']}", f"=B{VR['ai_B']}-D{VR['ai_B']}", None, None, "#,##0.0")
V.column_dimensions["A"].width = 70; V.column_dimensions["E"].width = 70; V.column_dimensions["D"].width = 40

# =====================================================================================
# RevDCF
# =====================================================================================
RV = S["RevDCF"]
put(RV, "A1", "RevDCF — 現價隱含什麼（情境 B 營收路徑；以封閉解反解 FY2028–FY2032 統一營業利益率，其餘輸入不變）", BOLD)
put(RV, "A2", "FCFF 對 FY2028–FY2032 營業利益率為線性，故以 m＝0 與 m＝1 兩個端點求 EV，線性內插到現價所需 EV。永續價值亦隨 FY2032 營業利益率線性變動。另列 Comparable 的隱含倍數與現價的對照。", None)
RVR = {}
r = 4
for k, lab in (("px", "現價（韓元）"), ("eqt", "現價隱含股權價值（十億韓元，含 NVIDIA 新股；以 ke 倒推 12 個月前＝現價）"), ("evr", "現價隱含企業價值 EV（扣淨現金、增資款、權益法投資，加非控制權益）"), ("ev0", "EV(m＝0)"), ("ev1", "EV(m＝1)"),
               ("m1", "現價隱含 FY2028–FY2032 營業利益率"), ("m0", "模型情境 B FY2028–FY2032 平均營業利益率"), ("sps", "Comparable：現價 ÷ 模型每股價值"), ("smul", "Comparable 主腿：現價隱含 P/S（現價市值 ÷ 目標時點 TTM 營收）"),
               ("smod", "Comparable 主腿：模型 P/S（同業中位 x 相對倍數）"), ("rv0", "模型情境 B FY2028 營收（十億韓元）"), ("w0", "模型 WACC")):
    RVR[k] = r; put(RV, f"A{r}", lab); r += 1
RBd, DB = PR["B"], DR["B"]
yrs_c = [CL[5], CL[6], CL[7], CL[8], CL[9]]
put(RV, f"B{RVR['px']}", f"={I('px')}", None, None, "#,##0")
put(RV, f"B{RVR['eqt']}", f"=B{RVR['px']}*{I('sh_val')}/1000", None, None, "#,##0")
put(RV, f"B{RVR['evr']}", f"=B{RVR['eqt']}-{I('netcash0')}-{I('nv_sh')}*{I('nv_px')}/1000-{I('eqm_inc')}*{I('pe_inv')}+{I('nci_v')}", None, None, "#,##0")


def ev_expr(m):
    parts = [f"SUM(DCF!{CL[3]}{DB['pv']}:{CL[4]}{DB['pv']})"]
    for c in yrs_c:
        parts.append(f"(({m})*PL!{c}{RBd['rev']}*(1-{I('tax')})+DCF!{c}{DB['da']}-DCF!{c}{DB['cx']}-DCF!{c}{DB['wc']})*DCF!{c}$7")
    parts.append(f"({m})*PL!{CL[9]}{RBd['rev']}*(1-{I('tax')})*(1+Val!$B${VR['g']})*(1-Val!$B${VR['g']}/Val!$B${VR['ronic']})/(Val!$B${VR['wacc']}-Val!$B${VR['g']})*DCF!{CL[9]}$7*(1+Val!$B${VR['wacc']})^(-0.5)")
    return "+".join(parts)


put(RV, f"B{RVR['ev0']}", "=" + ev_expr("0"), None, None, "#,##0")
put(RV, f"B{RVR['ev1']}", "=" + ev_expr("1"), None, None, "#,##0")
put(RV, f"B{RVR['m1']}", f"=(B{RVR['evr']}-B{RVR['ev0']})/(B{RVR['ev1']}-B{RVR['ev0']})", None, None, "0.0%")
put(RV, f"B{RVR['m0']}", "=AVERAGE(" + ",".join(f"PL!{c}{RBd['om']}" for c in yrs_c) + ")", None, None, "0.0%")
put(RV, f"B{RVR['sps']}", f"=B{RVR['px']}/Val!C{VR['ps_ps']}", None, None, "0.00")
put(RV, f"B{RVR['smul']}", f"=B{RVR['eqt']}/Val!C{VR['ttm_t']}", None, None, "0.00")
put(RV, f"B{RVR['smod']}", f"=Val!B{VR['ps_x']}", None, None, "0.00")
put(RV, f"B{RVR['rv0']}", f"=PL!{CL[5]}{RBd['rev']}", None, None, "#,##0")
put(RV, f"B{RVR['w0']}", f"=Val!B{VR['wacc']}", None, None, "0.00%")
r += 1
RVR["sum"] = r
put(RV, f"A{r}", f'="現價隱含 FY2028–FY2032 營業利益率約 "&TEXT(B{RVR["m1"]},"0.0%")&"，模型情境 B 平均 "&TEXT(B{RVR["m0"]},"0.0%")&"；現價隱含企業價值 "&TEXT(B{RVR["evr"]},"#,##0")&" 十億韓元，P/S 約 "&TEXT(B{RVR["smul"]},"0.00")&" 倍（Comparable 主腿 "&TEXT(B{RVR["smod"]},"0.00")&" 倍）。"')
RV.column_dimensions["A"].width = 90

# =====================================================================================
# Recon
# =====================================================================================
RC = S["Recon"]
put(RC, "A1", "Recon — 公司指引、市場共識與模型的差距、反解（公司指引與共識＝Interested-party，只作對照，不進基準）", BOLD)
put(RC, "A2", "公司指引：未找到（Naver 無正式年度指引，只稱 Naver Platform 長期利益率穩定在 25%）。共識：MarketScreener 2026-11-11（財報前，26 位分析師；3Q26 營業利益 451.3 低於共識 493.0／497.6），財報後各券商已下修目標價。共識口徑為分析師自訂，與本模型 K-IFRS 合計口徑不逐項相比。", None)
for j, h in enumerate(["項目", "年度", "模型情境 B", "共識（財報前）", "差距（模型 ÷ 共識 − 1）", "說明"]):
    put(RC, f"{L(j+1)}4", h, BOLD, HDR)
RCR = {}
r = 5
for k, lab, mod, cons, note, fm in (("rev26", "合併營收", f"=PL!{CL[3]}{RBd['rev']}", sg("cons_rev26")[1], "共識為財報前", "#,##0.0"), ("oi26", "營業利益", f"=PL!{CL[3]}{RBd['oi']}", sg("cons_oi26")[1], "共識營業利益為分析師口徑", "#,##0.0"),
                                    ("eps26", "EPS（韓元）", f"=PL!{CL[3]}{RBd['eps']}", sg("cons_eps26")[1], "模型 EPS 以合計淨利計；共識 EPS 口徑未核", "#,##0"), ("rev27", "合併營收", f"=PL!{CL[4]}{RBd['rev']}", sg("cons_rev27")[1], "含 AI Factory 營收；共識是否含未找到", "#,##0.0"),
                                    ("oi27", "營業利益", f"=PL!{CL[4]}{RBd['oi']}", sg("cons_oi27")[1], "模型利益率基準持平於近四季（G2 第 3 條），共識含 3Q26 前的預期", "#,##0.0"), ("eps27", "EPS（韓元）", f"=PL!{CL[4]}{RBd['eps']}", sg("cons_eps27")[1], "同上", "#,##0"),
                                    ("cx26", "資本支出", f"={I('capex26')}", sg("cons_capex26")[1], "共識偏低（財報前）；CFO 稱 GPU 單項逾 2 兆", "#,##0"), ("cx27", "資本支出", f"=PL!{CL[4]}{RBd['rev']}*{I('capex_pct')}", sg("cons_capex27")[1], "", "#,##0")):
    RCR[k] = r
    put(RC, f"A{r}", lab); put(RC, f"B{r}", "FY2026E" if k.endswith("26") else "FY2027E")
    put(RC, f"C{r}", mod, None, None, fm); put(RC, f"D{r}", cons, BLUE, None, fm); put(RC, f"E{r}", f"=C{r}/D{r}-1", None, None, "0.0%"); put(RC, f"F{r}", note)
    r += 1
r += 1
RCR["gd"] = r
put(RC, f"A{r}", "公司指引：未找到（無正式年度指引；Naver Platform 長期利益率目標稱穩定在 25%，Interested-party，只作對照）", BOLD)
r += 2
RCR["h1"] = r
put(RC, f"A{r}", "9M26 實績占模型 FY2026E", BOLD); r += 1
RCR["h1v"] = r
put(RC, f"A{r}", "9M26 營收實績（十億韓元）"); put(RC, f"C{r}", f"=Q!M{QROW['rev']}", None, None, "#,##0.0"); put(RC, f"F{r}", "Q 頁（Verified）")
r += 1
put(RC, f"A{r}", "9M26 占模型 FY2026E 營收比例"); put(RC, f"C{r}", f"=C{RCR['h1v']}/C{RCR['rev26']}", None, None, "0.0%"); put(RC, f"F{r}", "9M 占比約 74% 至 76% 為合理；過高代表 Q4 預測偏低")
r += 2
RCR["tpc"] = r
put(RC, f"A{r}", "目標價對照（財報後各券商，Interested-party，賣方預估）", BOLD); r += 1
for k in ("tp_hana", "tp_kis", "tp_meritz", "tp_sk", "tp_kyobo"):
    put(RC, f"A{r}", sg(k)[0]); put(RC, f"D{r}", sg(k)[1], BLUE, None, "#,##0"); put(RC, f"C{r}", f"=Val!B{VR['tp']}", None, None, "#,##0"); put(RC, f"E{r}", f"=C{r}/D{r}-1", None, None, "0.0%"); put(RC, f"F{r}", sg(k)[4] + "；" + sg(k)[5])
    r += 1
r += 1
RCR["rv"] = r
put(RC, f"A{r}", "整本重算反解（G2 第 9 條）：主導變數需變到多少，加權目標價才等於現價", BOLD); r += 1
SOL = (SENS or {}).get("solve")
RCR["solve_key"] = r
put(RC, f"A{r}", "主導變數（敏感度最大的營運輸入）"); put(RC, f"C{r}", SOL["label"] if SOL else "待 sens_naver.py 反解"); put(RC, f"F{r}", "取 Sens 頁相對基準絕對值最大者（排除折現率與估值規則參數）")
r += 1
RCR["solve_val"] = r
put(RC, f"A{r}", "反解值（整本以 LibreOffice 重算，割線法）"); put(RC, f"C{r}", SOL["x"] if SOL else None, None, None, SOL.get("fmt", "0.0000") if SOL else None); put(RC, f"D{r}", SOL["base"] if SOL else None, None, None, SOL.get("fmt", "0.0000") if SOL else None); put(RC, f"E{r}", "=IF(D{0}=\"\",\"\",C{0}/D{0}-1)".format(r), None, None, "0.0%"); put(RC, f"F{r}", "C＝反解值、D＝基準、E＝相對基準")
r += 1
RCR["solve_chk"] = r
put(RC, f"A{r}", "反解後加權目標價（核對：應等於現價）"); put(RC, f"C{r}", SOL["tp"] if SOL else None, None, None, "#,##0"); put(RC, f"D{r}", f"=Val!B{VR['px']}", None, None, "#,##0"); put(RC, f"F{r}", "本列為重算靜態值，In 頁改動後須重跑敏感度與反解腳本")
r += 2
RCR["sbc"] = r
put(RC, f"A{r}", "DCF（股票薪酬加回）讀數（G4 第 2 條；只作對照，不進加權目標價；股票薪酬金額未找到，占營收比為 Assumed，0＝與標準 DCF 相同）", BOLD); r += 1
for j, h in enumerate(["項目", "情境 A", "情境 B", "情境 C", "說明"]):
    put(RC, f"{L(j+1)}{r}", h, BOLD, HDR)
r += 1
SB = {}
for k, lab in (("pv", "股票薪酬加回的明確期現值"), ("tv", "股票薪酬加回的永續價值現值"), ("eqv", "股權價值（加回後）"), ("shv", "每股價值（韓元，評價日）"), ("d12", "推到目標時點（韓元）")):
    SB[k] = r; put(RC, f"A{r}", lab); r += 1
for s_, c_ in zip("ABC", "BCD"):
    Rr, Dd = PR[s_], DR[s_]
    put(RC, f"{c_}{SB['pv']}", "=" + "+".join(f"{I('sbc_pct')}*PL!{CL[j]}{Rr['rev']}*DCF!{CL[j]}$5*DCF!{CL[j]}$7" for j in range(3, 10)), None, None, "#,##0.0")
    put(RC, f"{c_}{SB['tv']}", f"={I('sbc_pct')}*PL!{CL[9]}{Rr['rev']}*(1+Val!$B${VR['g']})*(1-Val!$B${VR['g']}/Val!$B${VR['ronic']})/(Val!$B${VR['wacc']}-Val!$B${VR['g']})*DCF!{CL[9]}$7*(1+Val!$B${VR['wacc']})^(-0.5)", None, None, "#,##0.0")
    put(RC, f"{c_}{SB['eqv']}", f"=Val!{c_}{VR['eqv']}+{c_}{SB['pv']}+{c_}{SB['tv']}", None, None, "#,##0.0")
    put(RC, f"{c_}{SB['shv']}", f"={c_}{SB['eqv']}*1000/{I('sh_val')}", None, None, "#,##0")
    put(RC, f"{c_}{SB['d12']}", f"={c_}{SB['shv']}*(1+Val!$B${VR['ke']})^{Tyrs}-{I('dps0')}*{Tyrs}", None, None, "#,##0")
put(RC, f"E{SB['pv']}", "股票薪酬加回＝占營收比 x 營收，加到 FCFF；永續價值同比例")
RCR["sbc_tp"] = r
put(RC, f"A{r}", "DCF 目標價（股票薪酬加回，情境加權）"); put(RC, f"B{r}", f"={I('tpw_A')}*B{SB['d12']}+{I('tpw_B')}*C{SB['d12']}+{I('tpw_C')}*D{SB['d12']}", None, None, "#,##0"); put(RC, f"E{r}", "對照：標準 DCF 目標價見 Val 頁；不進加權目標價"); r += 1
RCR["sbc_tp2"] = r
put(RC, f"A{r}", "加權目標價（以加回後 DCF 代入，診斷對照）"); put(RC, f"B{r}", f"={I('w_dcf')}*B{RCR['sbc_tp']}+(1-{I('w_dcf')})*Val!B{VR['tpcomp']}", None, None, "#,##0"); put(RC, f"E{r}", "只作對照"); r += 2
put(RC, f"A{r}", "未找到：公司年度指引、第二個共識來源、法定申報封面股數、RSU 與選擇權、非控制權益、租賃負債、折舊攤銷與股票薪酬、年度資本支出、查詢量與 AI token 量（見研究包第 9 節）。")
RC.column_dimensions["A"].width = 70; RC.column_dimensions["F"].width = 70

# =====================================================================================
# Sens
# =====================================================================================
SE = S["Sens"]
put(SE, "A1", "Sens — 一次改一個輸入（In 頁敏感度低／高值），整本活頁簿以 LibreOffice 重算；加權目標價與 DCF 目標價取自重算後的 Val 頁。", BOLD)
put(SE, "A2", "相對基準＝案例加權目標價 ÷ 基準加權目標價 − 1。聯動與組合案（G2 第 2 條）：單一成本側的高價案是上界，價格與成本聯動案、模型組合案另列。In 頁改過而未重跑時模型狀態顯示 STALE。", None)
for j, h in enumerate(["驅動", "設定", "方向", "加權目標價", "相對基準", "鍵", "說明", "收支關卡 B（五年，十億韓元）", "DCF 目標價", "投資建議"]):
    put(SE, f"{L(j+1)}3", h, BOLD, HDR)
SENS_DEF = []   # (鍵字串, 方向, {鍵:值}, 驅動, 說明)
for row_ in IN_ROWS:
    if row_[0] == "#": continue
    k, lab, unit, lo, base, hi = row_[:6]
    if k in NAMES_SCN or k in SENS_SKIP or k in SENS_NOMODEL: continue
    for side, v in (("低", lo), ("高", hi)):
        if v is None or (not isinstance(base, str) and base == v): continue
        if isinstance(v, str): continue
        ex = "（上界：只動成本一邊）" if k == "ptok_idx" and side == "高" else ""
        SENS_DEF.append((k, side, {k: v}, lab + ex, f"基準 → {v}"))
SENS_DEF.append(("rel_adj,ps_only", "高", {"rel_adj": RMG["sens"]["ps_only"]}, "相對倍數只用 PS 讀數（歷史相對中位）", f"rel_adj＝{RMG['sens']['ps_only']:.3f}"))
SENS_DEF.append(("rel_adj,psg_only", "低", {"rel_adj": RMG["sens"]["psg_only"]}, "相對倍數只用 PSG 讀數", f"rel_adj＝{RMG['sens']['psg_only']:.3f}"))
SENS_DEF.append(("ptok_idx,r_link", "高", {"ptok_idx": 3, "r_link": 0.08}, "價格與成本聯動：token 單價取市場價級距，同時 R 上調 0.08（高價模型換得較高變現）", "ptok_idx＝3；r_link＝0.08"))
SENS_DEF.append(("route", "高", {"route": 0.5}, "模型組合：50% 查詢路由到低價模型（品質代價 R 下修依 In 頁基準）", "route＝0.5"))
SENS_DEF.append(("route,route_pen", "高", {"route": 0.5, "route_pen": 0.20}, "模型組合：50% 路由到低價模型，品質代價加重（R 全數下修 0.20）", "route＝0.5；route_pen＝0.20"))
SENS_DEF.append(("route,route_pen", "低", {"route": 0.5, "route_pen": 0.05}, "模型組合：50% 路由到低價模型，品質代價輕微（R 全數下修 0.05）", "route＝0.5；route_pen＝0.05"))
r = 4
SE_ROWS = []
sj = (SENS or {}).get("cases", {})
for k, side, ov, lab, st in SENS_DEF:
    put(SE, f"A{r}", f"{lab}：{side}")
    put(SE, f"B{r}", st)
    put(SE, f"C{r}", side)
    cs = sj.get(f"{k}|{side}")
    if cs:
        put(SE, f"D{r}", cs["tp"], None, None, "#,##0"); put(SE, f"H{r}", cs["gate"], None, None, "#,##0.0"); put(SE, f"I{r}", cs["dcf"], None, None, "#,##0"); put(SE, f"J{r}", cs["rec"])
    put(SE, f"E{r}", f'=IF(D{r}="","",D{r}/$D$BASEROW-1)', None, None, "0.0%")
    put(SE, f"F{r}", k)
    put(SE, f"G{r}", "重算值（sens_naver.py）" if cs else "未重算")
    SE_ROWS.append(r)
    r += 1
SE_END = r - 1
r += 1
SE_BASE = r
put(SE, f"A{r}", "基準加權目標價（韓元）"); put(SE, f"D{r}", f"=Val!B{VR['tp']}", None, None, "#,##0")
for rr_ in SE_ROWS:
    SE[f"E{rr_}"].value = SE[f"E{rr_}"].value.replace("BASEROW", str(SE_BASE))
r += 1
SE_STALE = r
put(SE, f"A{r}", "Sens 重跑時的基準加權目標價（韓元）"); put(SE, f"D{r}", (SENS or {}).get("base_tp"), None, None, "#,##0")
r += 1
num_inputs = [(row_[0], row_) for row_ in IN_ROWS if row_[0] != "#"]
SE_FP = r
put(SE, f"A{r}", "In 頁輸入指紋（現值：基準欄 E 數值型輸入的加權和）")
terms = [f"{(i_ % 7) + 1}*In!E{INR[k]}" for i_, (k, row_) in enumerate(num_inputs) if isinstance(row_[4], (int, float)) and not isinstance(row_[4], bool) and k not in OVR]
put(SE, f"D{r}", "=" + "+".join(terms), None, None, "0.000000")
r += 1
SE_FP0 = r
put(SE, f"A{r}", "In 頁輸入指紋（Sens 重跑時）"); put(SE, f"D{r}", (SENS or {}).get("fp"), None, None, "0.000000")
r += 1
SE_STATUS = r
put(SE, f"A{r}", "模型狀態（CURRENT＝Sens、HTML、README 與本檔輸入一致；STALE＝In 頁已改、須重跑 sens_naver.py）")
put(SE, f"D{r}", f'=IF(ABS(D{SE_FP}-D{SE_FP0})<0.000001,"CURRENT","STALE")')
r += 1
SE_BID = r
fp_val = sum(((i_ % 7) + 1) * row_[4] for i_, (k, row_) in enumerate(num_inputs) if isinstance(row_[4], (int, float)) and not isinstance(row_[4], bool) and k not in OVR)
BUILD_ID = "naver-v%s-%s" % (VER, hashlib.sha1(f"{fp_val:.6f}".encode()).hexdigest()[:10])
put(SE, f"A{r}", "build_id（README、Sens、HTML 三處相同）"); put(SE, f"D{r}", BUILD_ID)
r += 2
SE_TORN = r
put(SE, f"A{r}", "龍捲風圖說明：|相對基準| ≥ 10% 的案例列在前，未達者摺疊；全部案例見上表。", None)
SE.column_dimensions["A"].width = 70

# =====================================================================================
# README
# =====================================================================================
RD = S["README"]
Vv = lambda k, col="B": f"Val!{col}{VR[k]}"
put(RD, "A1", f'="Naver（035420）個股模型 v{VER}｜build_id "&Sens!D{SE_BID}&"｜資料截止日 "&TEXT({I("data_date")},"yyyy-mm-dd")&"｜成品日 {BUILDDATE.isoformat()}｜股價基準日 "&TEXT({I("px_date")},"yyyy-mm-dd")&"｜評價日 "&TEXT({I("val_date")},"yyyy-mm-dd")&"｜目標時點 "&TEXT({I("tp_date")},"yyyy-mm-dd")&"（本版沒有做先驗；Tokenomics 快照 {TKV}）"')
put(RD, "A2", "v0.2.1 要點（G6）：Comparable 主腿改為 P/S 口徑，前瞻本益比 Comparable 改為對照列、不進加權。v0.2 要點：相對倍數改為歷史 PS 與 PSG 兩讀數幾何平均、無風險利率統一取 2026-11-13、並列 DCF（股票薪酬加回）對照，其餘同 v0.1。命題與結構：命題＝原型 A（按廣告計價）的單位經濟：查詢量、廣告負載、點擊率、每次點擊價格；官方分部 Naver Platform、Financial Platform、Global Opportunities（2026 年起新口徑）；DCF（標準 FCFF）與 Comparable（分部 NTM 淨利乘同業前瞻本益比）各 50%，三情境只在 AI 驅動分化。不適用項目：算力、Wattonomics 與 Tokenomics 自用讀法 2（Naver 的自用 AI 成本放在命題模組）、三層協調。")
put(RD, "A3", "命題：AI 進入 Naver 搜尋後，每單位廣告營收的變化＝AI 簡報帶來的營收增量（AI 回答查詢的廣告負載、點擊率與價格相對值，扣除版面擠壓）減去外部 AI 分流造成的侵蝕，再減去自建 HyperCLOVA X 的推論成本；三情境檢驗淨效果的正負與大小。")
put(RD, "A4", f'="核心問題：FY2032E 的 AI 淨效果（營收層）三情境為 "&TEXT({Vv("ai_net")},"#,##0")&"／"&TEXT({Vv("ai_net","C")},"#,##0")&"／"&TEXT({Vv("ai_net","D")},"#,##0")&" 十億韓元（A／B／C），損益平衡 R* 在基準為 "&TEXT(PL!{CL[9]}{PR["B"]["rstar"]},"0.00")&"。"')
put(RD, "A5", f'="投資建議："&{Vv("rec")}&"（估值評等 "&{Vv("val_rating")}&"）；加權目標價 "&TEXT({Vv("tp")},"#,##0")&" 韓元，上檔空間 "&TEXT({Vv("up")},"+0.0%;-0.0%")&"，現價 "&TEXT({I("px")},"#,##0")&"（"&TEXT({I("px_date")},"yyyy-mm-dd")&" 收盤）；收支合計 "&TEXT({Vv("gate_B")},"#,##0")&" 十億韓元。"')
put(RD, "A6", f'="觀察一：目標價 "&TEXT({Vv("tp")},"#,##0")&"＝DCF "&TEXT({Vv("tpdcf")},"#,##0")&" x "&TEXT({I("w_dcf")},"0%")&"＋Comparable（P/S 主腿）"&TEXT({Vv("tpcomp")},"#,##0")&" x "&TEXT(1-{I("w_dcf")},"0%")&"；三情境 "&TEXT({Vv("t_A")},"#,##0")&"／"&TEXT({Vv("t_B")},"#,##0")&"／"&TEXT({Vv("t_C")},"#,##0")&"。"')
put(RD, "A7", f'="觀察二：現價隱含 FY2028–FY2032 營業利益率 "&TEXT(RevDCF!B{RVR["m1"]},"0.0%")&"，模型為 "&TEXT(RevDCF!B{RVR["m0"]},"0.0%")&"；基準利益率持平於近四季，單季已降至 "&TEXT(Q!K{QROW["np_m"]}/100,"0.0%")&"（Naver Platform）。"')
put(RD, "A8", f'="觀察三：FY2032E 推論成本情境 B "&TEXT(PL!{CL[9]}{PR["B"]["inf"]},"0.0")&" 十億韓元，占廣告營收 "&TEXT(PL!{CL[9]}{PR["B"]["inf"]}/PL!{CL[9]}{PR["B"]["ads"]},"0.00%")&"；成本在三個級距間差距小，主導的是 R 與分流。"')
put(RD, "A9", f'="觀察四：DCF 與 Comparable 落差 "&TEXT({Vv("gp_gap")},"#,##0")&" 韓元，來自倍數 "&TEXT({Vv("gp_mult")},"#,##0")&"、時點 "&TEXT({Vv("gp_time")},"#,##0")&"；股票薪酬與槓桿無差（金額未找到）。"')
put(RD, "A10", f'="觀察五：FY2026E 營收 "&TEXT(Recon!C{RCR["rev26"]},"#,##0")&" 對財報前共識 "&TEXT(Recon!D{RCR["rev26"]},"#,##0")&" 為 "&TEXT(Recon!E{RCR["rev26"]},"+0.0%;-0.0%")&"；營業利益差 "&TEXT(Recon!E{RCR["oi26"]},"+0.0%;-0.0%")&"。"')
put(RD, "A11", f'="橫幅："&{Vv("rec")}&"；加權目標價 "&TEXT({Vv("tp")},"#,##0")&" 韓元（上檔空間 "&TEXT({Vv("up")},"+0.0%;-0.0%")&"，現價 "&TEXT({I("px")},"#,##0")&"）。AI 淨效果三情境分化；查詢量與利益率路徑為 Assumed。"')
put(RD, "A12", f'="落差歸因（情境 B，每股）：Comparable 比 DCF 高 "&TEXT({Vv("gp_gap")},"#,##0")&" 韓元＝倍數與口徑 "&TEXT({Vv("gp_mult")},"#,##0")&"＋時點 "&TEXT({Vv("gp_time")},"#,##0")&"＋股票薪酬 "&TEXT({Vv("gp_sbc")},"0")&"＋槓桿 "&TEXT({Vv("gp_lev")},"0")&"；DCF 隱含 P/S "&TEXT({Vv("gp_dcfmult")},"0.00")&" 倍，Comparable 主腿 P/S "&TEXT({Vv("ps_x")},"0.00")&" 倍（同業中位 "&TEXT({Vv("ps_med")},"0.00")&" x 相對倍數 "&TEXT({Vv("rel")},"0.00")&"）；對照列現行本益比 "&TEXT({Vv("own_pe")},"0.0")&" 倍，為同業中位數的 "&TEXT({Vv("own_ratio")},"0.00")&" 倍。"')
put(RD, "A13", "情境分化依據：A／B／C 只在三個 AI 驅動分化——FY2032 AI 回答占查詢比例（A 60%／B 50%／C 50%）、AI 回答查詢的每查詢營收相對值 R（1.15／1.04／0.85）、外部 AI 分流占比（0%／3%／8%）。依據：公司只揭露 AI 貢獻 1Q26 廣告成長逾 50%、2Q26 逾 60%，與 AI 簡報廣告點擊率 +30%（均為 Interested-party，非隨機對照），無查詢量與營收拆分；分化幅度為判斷，不是觀測值。其他分部成長三情境相同。權重 25／50／25（Assumed）。")
put(RD, "A14", "與標準規格的偏離：(1) 股數取第二來源，未取法定申報封面；RSU、選擇權明細未找到，以稀釋係數 Assumed 取代庫藏股法；(2) 非控制權益拆分未找到，暫為 0；(3) 比率參數（Naver Platform、Financial Platform、Global 利益率、營業成本率）基準持平於近四季，收斂與回升只放敏感度（G2 第 3 條；G3 第 1 條：FY2021–FY2025 營業利益率 19.4%、15.9%、15.4%、18.4%、18.3%，無規模經濟趨勢可延伸，故不加改善路徑，Sens 低案為單季 24.0%）；營運資金歷史比率未找到，非訂閱制，取 Assumed 3%（G3 第 2 條）；(4) 折舊攤銷與資本支出歷史由第二來源推得，公司未揭露；(5) 評價日取資產負債表日 2026-09-30，NVIDIA 增資款依事件價另加；(6) Comparable 的同業只有 5 家，不足 8 家不做迴歸，Global Opportunities 若為負盈餘不計價；相對倍數用 PS／PSG 歷史相對值近似套用於前瞻本益比（A41）；(7) AI Factory 資產由 Asset Co 持有，其資本支出與負債未入模型。")
put(RD, "A15", "不適用：算力與 Wattonomics、Tokenomics 自用 AI 讀法 2（Naver 自建 HyperCLOVA X，其自用推論成本在命題模組，只讀 Tokenomics 全成本，標類比）；三層協調（Coord 頁、C0、兩個讀數）：非 Siliconomics 覆蓋；對帳中兩讀數目標價不適用。")
put(RD, "A16", "已知限制：搜尋查詢量、AI 簡報 token 量、HyperCLOVA X 自用成本、年度資本支出、折舊攤銷、股票薪酬、租賃負債、RSU 與選擇權、非控制權益、公司年度指引均未找到（見 data/research_naver.md 第 9 節）；2026 年起分部與費用分類重編，FY2023A–FY2024A 無新口徑分部，歷史為舊口徑（Hist 頁）。NAVER Financial 與 Dunamu 合併未入模型。")
put(RD, "A17", f'="逾期標示（G2 第 10 條，與現價基準日 "&TEXT({I("px_date")},"yyyy-mm-dd")&" 相差超過 2 個交易日者）：無風險利率 "&{Vv("rf_flag")}&"；同業倍數 PINS "&IF({Vv("pe_pins","E").replace("B","E",1) if False else "Val!E"+str(VR["peers"]["pe_pins"])}="","未逾期","逾期")&"、Kakao "&IF(Val!E{VR["peers"]["pe_kakao"]}="","未逾期","逾期")&"、LY "&IF(Val!E{VR["peers"]["pe_ly"]}="","未逾期","逾期")&"、GOOGL "&IF(Val!E{VR["peers"]["pe_goog"]}="","未逾期","逾期")&"、META "&IF(Val!E{VR["peers"]["pe_meta"]}="","未逾期","逾期")&"。逾期倍數仍列入群組中位數，下次改版更新。"')
put(RD, "A18", "驅動 → 推導對照表（每頁輸入、推導量、因果方向）", BOLD)
for j, h in enumerate(["頁", "輸入", "推導量", "因果方向"]):
    put(RD, f"{L(j+1)}19", h, BOLD, HDR)
MAP = [("In→PL", "查詢量、廣告負載、點擊率、每次點擊價格、AI 回答占比、R、外部分流", "AI 暴露廣告（基礎＋侵蝕＋增量）", "分流↑→營收↓；R 高於 1 且占比↑→營收↑"),
       ("In→PL", "AI 回答占比、token 數、單價級距（Tokenomics）、匯率、路由比例", "推論成本", "用量或單價↑→成本↑，計入營業成本"),
       ("In→PL", "Services、Financial、C2C、Contents、Enterprise 成長、AI Factory 容量與每 MW 營收", "各分部營收與合併營收", "成長率由 FY2027 起點線性收斂到 FY2032"),
       ("In→PL", "近四季利益率（Derived）、AI 邊際流入流出", "分部營業利益與合併營業利益", "利益率持平，AI 影響只加減相對 FY2026 的增減"),
       ("PL→DCF", "營業利益、折舊（滾動）、資本支出、營運資金", "FCFF 與現值", "FCFF＝NOPAT＋折舊−資本支出−營運資金增加"),
       ("DCF／PL→Val", "WACC、永續成長率、同業前瞻本益比", "DCF 與 Comparable 目標價、加權目標價", "兩法依權重加權；評等依上檔空間與收支關卡"),
       ("Val→Sens", "每個有區間的輸入", "加權目標價相對基準", "一次改一個輸入，整本重算")]
for i_, rowv in enumerate(MAP):
    for j, v in enumerate(rowv):
        put(RD, f"{L(j+1)}{20+i_}", v)
put(RD, "A28", "市場資料日期、失效條件與風險登錄", BOLD)
put(RD, "A29", "市場資料日期：現價 2026-11-13 收盤（KRX，最近交易日）；韓國 10 年期國債 3.99%（2026-11-13，Trading Economics 回推）；同業本益比 GOOGL、META 2026-11-13、PINS 2026-11-06、LY 2026-10-13、Kakao 2026-09-09；共識 2026-11-11（財報前）。")
put(RD, "A30", f'="合理性關卡診斷（G3 第 3 條）：DCF 目標價 "&TEXT({Vv("tpdcf")},"#,##0")&" 為現價的 "&TEXT({Vv("tpdcf")}/{Vv("px")},"0.00")&" 倍（界線 0.25 至 4 倍："&IF(OR({Vv("tpdcf")}/{Vv("px")}<0.25,{Vv("tpdcf")}/{Vv("px")}>4),"觸發","未觸發")&"）；EV 為負："&IF({Vv("ev","C")}<0,"是","否")&"；評等 "&{Vv("val_rating")}&"（強力評等："&IF(OR({Vv("val_rating")}="強力買進",{Vv("val_rating")}="強力放空"),"是","否")&"）；FY2027 營收、營業利益對財報前共識 "&TEXT(Recon!E{RCR["rev27"]},"+0.0%;-0.0%")&"、"&TEXT(Recon!E{RCR["oi27"]},"+0.0%;-0.0%")&"（公司無年度指引，改對共識，界線 10%）。診斷：Comparable 主腿（P/S，"&TEXT({Vv("tpcomp")},"#,##0")&"）與 DCF 的落差來源是倍數：主腿 P/S "&TEXT({Vv("ps_x")},"0.00")&" 倍（同業中位 "&TEXT({Vv("ps_med")},"0.00")&" x 相對倍數 "&TEXT({Vv("rel")},"0.00")&"），DCF 隱含 P/S "&TEXT({Vv("gp_dcfmult")},"0.00")&" 倍。前瞻本益比 Comparable 改為對照列、不進加權：相對倍數 1.00 時 "&TEXT({Vv("comp_pe1")},"#,##0")&"，取本公司現行比值（"&TEXT({Vv("own_ratio")},"0.00")&"）時 "&TEXT({Vv("comp_alt")},"#,##0")&"（G6：PS／PSG 相對值已含利潤率差，不可再套前瞻本益比）。DCF 單獨為現價的 "&TEXT({Vv("tpdcf")}/{Vv("px")}-1,"+0%;-0%")&"。結論：不是建模錯誤（恆等式與分部對帳為零），是韓國折價是否維持歷史水準的經濟判斷，Comparable 主腿與評等依賴它。"')
put(RD, "A31", "(1) Naver Platform 利益率續降至單季 24% 以下而廣告成長放緩：基準的持平假設失效。(2) AI Factory 實際上線容量、每 MW 營收與 Naver 保證使用費的會計處理公布後與 Assumed 顯著不同。(3) NAVER Financial 與 Dunamu 合併條件（換股比例、會計）公布：Financial Platform 價值與權益法投資需重估。")
put(RD, "A32", "風險登錄", BOLD)
for j, h in enumerate(["風險", "是否建模", "理由", "觀察指標"]):
    put(RD, f"{L(j+1)}33", h, BOLD, HDR)
RISKS = [("外部 AI 取代搜尋入口", "是（情境外部分流）", "Naver 韓國搜尋占有率一年內升至 81.34%，但 ChatGPT、Gemini 持續滲透", "廣告營收年增、占有率"),
         ("AI 簡報壓縮廣告位", "是（R 與廣告負載相對值）", "公司稱 CTR +30% 為 Interested-party", "AI 簡報廣告營收揭露"),
         ("推論成本", "是", "讀 Tokenomics 全成本，三級距差距小；自用 GPU 折舊年限 6 年", "基礎設施費用、GPU 折舊"),
         ("AI Factory 資金與保證責任", "部分（營收與利益率，不含資本支出與負債）", "Asset Co 預期合併；Naver 保證使用費最長 8 年", "Asset Co 合併後負債"),
         ("利益率下行（行銷費用、基礎設施）", "否（基準持平，低案入 Sens）", "近四季利益率由 33.3% 降至 24.0%", "Naver Platform 單季利益率")]
for i_, rowv in enumerate(RISKS):
    for j, v in enumerate(rowv):
        put(RD, f"{L(j+1)}{34+i_}", v)
put(RD, "A41", f'="歷史相對倍數（G4，PS 與 PSG）：相對倍數 基準＝兩讀數幾何平均 "&TEXT({Vv("rm_geo")},"0.00")&"（PS 讀數 "&TEXT({Vv("rm_ps")},"0.00")&"、PSG 讀數 "&TEXT({Vv("rm_psg")},"0.00")&"），Comparable 主腿改為 P/S 口徑（G6），與 PS／PSG 相對值同口徑。兩讀數差 "&TEXT({Vv("rm_div")},"0%")&IF({Vv("rm_div")}>0.3,"，大於 30%：原因是本公司 TTM 營收年增 "&TEXT({Vv("rm_cg")},"0.0%")&" 低於同業中位 "&TEXT({Vv("rm_pg")},"0.0%")&"（成長比 "&TEXT({Vv("rm_g")},"0.00")&"），PSG 把歷史每單位成長的溢價按較低成長比折算，故低於 PS 讀數","，未超過 30%")&"。"')
put(RD, "A42", f'="DCF（股票薪酬加回）對照：每股價值 "&TEXT(Recon!C{SB["shv"]},"#,##0")&" 韓元（情境 B），目標價 "&TEXT(Recon!B{RCR["sbc_tp"]},"#,##0")&" 韓元（情境加權），標準 DCF "&TEXT({Vv("tpdcf")},"#,##0")&"；股票薪酬金額未找到，占營收比 Assumed "&TEXT({I("sbc_pct")},"0.0%")&"，故兩者相同。只作對照，不進加權目標價。"')
put(RD, "A44", "G6 口徑一致：歷史 PS 與 PSG 相對值已含利潤率差，不能再套前瞻本益比；Comparable 主腿因此改為同業中位 P/S（各家最近一季，季末日逐家標示，逾期者旗標）乘相對倍數乘目標日 TTM 營收除以稀釋股數。前瞻本益比 Comparable 降為對照列（相對倍數 1.00 與本公司現行比值兩個讀數），不進加權目標價。敏感度四案（只用 PS、只用 PSG、低案、高案）套用在 P/S 主腿。")
put(RD, "A40", "工作表：README／In／Q／Hist／PL／DCF／CF／Val／RevDCF／Recon／Sens／Sources。成品 dist/<日期>_Naver個股模型_v0.2.1.xlsx 與 .html。")
for rr_ in list(range(1, 18)) + [30, 41, 42]:
    RD[f"A{rr_}"].alignment = Alignment(wrap_text=True, vertical="top")
for rr_ in (29, 31):
    RD[f"A{rr_}"].alignment = Alignment(wrap_text=True, vertical="top")
RD.column_dimensions["A"].width = 120
RD.column_dimensions["B"].width = 36; RD.column_dimensions["C"].width = 36; RD.column_dimensions["D"].width = 36

SO = S["Sources"]
put(SO, "A1", "Sources — 來源與版本", BOLD)
for i_, (a, b_) in enumerate([("事實總帳", "facts_naver.py → data/facts_naver.csv、data/research_naver.md"), ("Tokenomics 快照", f"data/tokenomics_snapshot_v22.0.json（{TKV}；import_tokenomics.py --check 通過）"),
                              ("官方資料", "Naver 財報簡報（4Q22–3Q26）、新聞稿、整合報告 2025；URL 見 facts_naver.csv 出處欄"), ("現價與利率", "KIS 個股頁（2026-11-13）、tradingeconomics.com（2026-11-16）"),
                              ("共識與同業", "MarketScreener（2026-11-11）、StockAnalysis（同業前瞻本益比，日期逐家）")]):
    put(SO, f"A{3+i_}", a); put(SO, f"B{3+i_}", b_)

ROWMAP = {"IN": INR, "PR": PR, "IX": IX, "RID": RID, "VR": VR, "RVR": RVR, "RCR": RCR, "DR": DR, "CFR": CFR, "SE_ROWS": SE_ROWS, "SE_BASE": SE_BASE, "SE_STALE": SE_STALE, "SE_STATUS": SE_STATUS, "SE_BID": SE_BID,
          "SE_END": SE_END, "SE_TORN": SE_TORN, "STDH": STDH, "STD_NOTE": STD_NOTE, "PLSUM": PLSUM, "HR": HRW, "SCN_ROWS": SCN_ROWS, "SE_FP": SE_FP, "BUILD_ID": BUILD_ID, "QROW": QROW,
          "SENS_DEF": [[k, s_, ov, lab] for k, s_, ov, lab, st in SENS_DEF], "IN_LAST": IN_LAST, "IN_ROWS": [[x[0], x[1], x[3], (None if isinstance(x[4], str) else x[4]), x[5]] for x in IN_ROWS if x[0] != "#"]}
if os.environ.get("ROWMAP"):
    json.dump(ROWMAP, open(os.environ["ROWMAP"], "w", encoding="utf-8"), ensure_ascii=False, indent=1, default=str)
wb.save(OUT)
print("saved", OUT, BUILD_ID)
