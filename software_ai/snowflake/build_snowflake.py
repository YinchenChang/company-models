#!/usr/bin/env python3
"""Snowflake 個股模型 v0.1（原型 U：按用量計價；命題＋營收分部 → 非 GAAP 損益表 → FCFF DCF＋Comparable）。
用法：python3 build_snowflake.py <raw.xlsx> [--sens data/sens_v01.json] [--rowmap rowmap.json]
所有計算為活公式；輸入在 In 頁（藍字）。LibreOffice 重算後才有快取值。
"""
import datetime, hashlib, json, os, sys
import openpyxl
sys.path.insert(0, os.path.join(os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__)))), 'tools'))
import fcff_dcf as FD
from openpyxl.styles import Font, PatternFill
from openpyxl.workbook.defined_name import DefinedName

HERE = os.path.dirname(os.path.abspath(__file__))
OUT = sys.argv[1]
SENS = sys.argv[sys.argv.index("--sens") + 1] if "--sens" in sys.argv else None
ROWMAP = sys.argv[sys.argv.index("--rowmap") + 1] if "--rowmap" in sys.argv else None
VER, VDATE = "0.1", "20261116"
TK = json.load(open(os.path.join(HERE, "data/tokenomics_snapshot_v22.0.json"), encoding="utf-8"))
TKI = TK["items"]
BLUE = Font(color="0000FF"); BOLD = Font(bold=True); HDR = PatternFill("solid", fgColor="DDEBF7")
wb = openpyxl.Workbook()
SH = {n: wb.active if n == "README" else wb.create_sheet(n) for n in ["README", "In", "Q", "Seg", "PL", "DCF", "Val", "RevDCF", "Recon", "Sens", "Sources"]}
SH["README"].title = "README"

D0 = "Verified"; IP = "Interested-party"; AN = "Analogy"; AS = "Assumed"; DE = "Derived"
INP = []


def add(key, label, unit, lo, base, hi, tag, src, inc=""):
    if "http" in src: src = src.replace("未找到", "未取得")
    INP.append((key, label, unit, lo, base, hi, tag, src, inc))


U = "https://www.sec.gov/Archives/edgar/data/1640147/"
Q2PR = "Q2 FY27 財報新聞稿（8-K 附件 99.1，2026-09-02）" + U + "000164014726000033/fy2027q2earnings.htm"
Q2Q = "10-Q Q2 FY27（截至 2026-07-31）" + U + "000164014726000037/snow-20260731.htm"
SA = "StockAnalysis SNOW statistics／forecast 頁（2026-11-13 收盤；第二來源，非交易所）https://stockanalysis.com/stocks/snow/statistics/"
TKS = "Tokenomics v22.0（commit 79f6b30；data/tokenomics_snapshot_v22.0.json）"

# ---- 市場與資本結構
add("px_date", "股價基準日（收盤）", "日期", None, datetime.date(2026, 11, 13), None, D0, SA)
add("px", "現價（2026-11-13 收盤）", "US$/股", None, 291.00, None, D0, SA + "；收盤 291.00，盤後 290.00")
add("sh_out", "流通股數（10-Q 封面，2026-08-21）", "M股", None, 352.8, None, D0, Q2Q + " 封面：352.8 百萬股（2026-08-21）；法定申報為主來源，StockAnalysis 352.46M 為第二來源")
add("tax", "稅率（非 GAAP 稅與 NOPAT 稅）", "比例", 0.2, 0.25, 0.28, DE, "FY2024–FY2026 非 GAAP 淨利相對（非 GAAP 營業利益＋利息與其他收支）的隱含稅率約 26.0%、25.5%、24.0%（Derived；公司未揭露非 GAAP 稅前與稅額）")
add("kd_sp", "稅前負債信用利差（相對 rf）", "比例", 0.008, 0.012, 0.02, AS, "可轉債為 0 票息；公司債有效利率未取得（未找到）；以利差假設")
add("rsu", "在外 RSU／PSU（權益分類 23.529＋負債分類 1.435；沒收率未揭露，全額計入）", "M股", 23.529, 24.964, 25.089, D0, Q2Q + " Note 12：權益分類未歸屬 RSU（含領導層 PRSU）23,529 千股；負債分類 PRSU 1,435 千股；每股盈餘附註列 RSU 25,089 千股（高案）；沒收率未揭露（未找到）")
add("opt_n", "在外選擇權（2026-07-31）", "M股", None, 9.277, None, D0, Q2Q + " Note 12：9,277 千股")
add("opt_k", "選擇權加權平均履約價", "US$/股", None, 30.74, None, D0, Q2Q + " Note 12：US$30.74（庫藏股法，價格用現價）")
add("conv_prin", "可轉換優先票據本金（2027 與 2029 年票據各 1,150）", "US$M", None, 2300, None, D0, Q2Q + " Note 10：各 1.15B 本金")
add("conv_px", "可轉債轉換價", "US$/股", None, 157.50, None, D0, Q2Q + " Note 10：2027 與 2029 年票據轉換價皆為 US$157.50")
add("conv_sh", "可轉債轉換股數（兩檔合計）", "M股", None, 14.603, None, D0, Q2Q + " 每股盈餘附註：轉換選擇權基礎股數 14,603 千股")
add("cap_px", "Capped call 上限價", "US$/股", None, 225.00, None, D0, Q2Q + " MD&A：capped call 上限價 US$225.00（履約價 157.50）；現價高於轉換價時以 if-converted 計入股數、債務為 0，並扣 capped call 在 157.50–225 區間的回收股數")
add("rel_mult", "相對倍數調整係數（本公司 P/S ÷ 同業中位數的近五期中位數；夾在 0.5–2.0）", "倍", 1.0, "=Val!$B$@@RELM@@", 2.0, DE, "自算（Val 頁表）：StockAnalysis 各公司財務比率頁近五個會計年度末的 P/S，SNOW 除以八家同業中位數，取五期中位數，夾在 tools/relative_multiple.py 的 0.5–2.0（原則 G3 第 4 條）；低案 1.00＝未調整；高案＝夾值上限 2.0；期間為歷史落後 P/S，與前瞻倍數口徑不同，為承重輸入")
add("opx_a", "非 GAAP 營業費用：固定部分（US$M；營業費用＝固定＋變動×營收）", "US$M", 500, "=Q!$B$@@OPXA@@", 1000, DE, "Q 頁迴歸（FY2022–FY2027 實績）：最近三點（FY2025、FY2026、LTM 至 Q2 FY27）的截距，原則 G3 第 1 條；「全期持平」（固定 0、變動＝最近四季營業費用率）為 Sens 下緣案例")
add("opx_v", "非 GAAP 營業費用：變動部分（每增加 1 美元營收增加的營業費用）", "比例", 0.40, "=Q!$B$@@OPXV@@", 0.5705, DE, "Q 頁迴歸：最近三點斜率；低案 0.40 較強規模經濟（假設）；高案 0.5705＝六點全樣本斜率（較弱的規模經濟）")
add("sbc_a", "股票薪酬：固定部分（US$M；股票薪酬＝固定＋變動×營收）", "US$M", 800, "=Q!$B$@@SBCA@@", 1500, DE, "Q 頁迴歸：最近三點截距；「全期持平」（占營收比固定為最近四季）為 Sens 下緣案例")
add("w_dcf", "DCF 權重（其餘為 Comparable）", "比例", 0.3, 0.5, 0.7, AS, "原則第 4 節預設 50／50")
# ---- FY2027 錨點（H1 實績＋下半年假設）
add("q1p", "產品營收 Q1 FY27（實績）", "US$M", None, 1334.329, None, D0, "Q1 FY27 財報新聞稿 " + U + "000164014726000027/fy2027q1earnings.htm")
add("q2p", "產品營收 Q2 FY27（實績）", "US$M", None, 1491.861, None, D0, Q2PR)
add("q1s", "專業服務及其他營收 Q1 FY27（實績）", "US$M", None, 56.622, None, D0, "Q1 FY27 財報新聞稿")
add("q2s", "專業服務及其他營收 Q2 FY27（實績）", "US$M", None, 54.932, None, D0, Q2PR)
add("q3g", "產品營收 Q3 FY27 季增率", "比例", 0.04, 0.07, 0.10, AS, "公司 Q3 指引 1,588–1,593（季增約 6.5–6.8%，Interested-party）只作 Recon 對照，不進基準；基準 7% 為假設，未找到 Q3 實績（尚未公布）")
add("q4g", "產品營收 Q4 FY27 季增率", "比例", 0.03, 0.055, 0.08, AS, "Q4 FY26 季增 5.9%、Q2 FY27 季增 11.8%（Derived）；基準 5.5% 為假設")
add("ps_h2r", "專業服務及其他：下半年營收÷上半年營收", "倍", 0.95, 1.0, 1.05, AS, "FY2026 下半年 111.9 對上半年 99.7（Derived）；Q2 FY27 54.9 年減；基準持平為假設")
# ---- 命題模組（原型 U；名稱與 software_ai/unit_economics 一致）
add("g_px", "模型 API 市場單價年變化（同一層級、同一 token 類型）｜待 AI供應鏈 登錄", "比例/年", -0.5, -0.25, 0.0, AS, "Tokenomics 無未來價格路徑，待 AI供應鏈 登錄；與 software_ai/unit_economics 同值（-50%／-25%／0%）")
add("g_px_rs", "AI 功能轉售的價格路徑（售價隨成本的年變化；基準＝成本路徑，成本加成定價）", "比例/年", -0.5, -0.25, 0.0, AS, "基準與 g_px 同值＝成本加成、售價隨成本聯動；只動 g_px 而不動本列＝成本上升但售價不隨之變動，為「上界」案例（原則第 44 條）；Cortex 以 AI Credit 定價，價目調整頻率未揭露")
add("g_tok", "每任務 token 數年變化（代理迴圈加長、思考加深）", "比例/年", 0.0, 0.2, 0.5, AS, "與 software_ai/unit_economics 同值（0%／20%／50%）；無直接量測")
add("ovh", "實際工作負載 token ÷ Tokenomics 參考工作負載", "倍", 1.0, 3.0, 8.0, AN, "與 software_ai/unit_economics 同值（1／3／8）；Snowflake 實際 Cortex 呼叫的 token 數未揭露")
add("U_tier", "AI 功能使用的模型層級（1 Luna／2 Terra／3 Sol／4 Astra）", "序號", 1, 2, 3, AN, "Cortex 價目同時提供 Anthropic、OpenAI、Llama 多家模型（Service Consumption Table 6(a)）；預設中低層級；各層級使用比重未揭露")
add("U_gm0", "AI 前核心消耗的產品毛利率（非 GAAP；基期 FY2027）", "比例", 0.74, 0.77, 0.79, AS, "校準到最近四季實績：Q3 FY26 至 Q2 FY27 非 GAAP 產品毛利率 75.9／75.1／75.1／74.7%，平均約 75.2%（Derived）；本值使 FY2027 含 AI 後的產品毛利率約 75%；公司指引 74.0% 只在 Recon 對照")
add("gm_dr", "核心產品毛利率逐年變化（FY2028 起；基準持平，改善只放 Sens）", "百分點/年", 0.0, 0.0, 0.005, AS, "基準全期持平於最近四季實績（原則第 35 條）；FY2024→FY2026 非 GAAP 產品毛利率 77.8%→75.8% 是下滑（Derived），雲端費用占產品營業成本 67%→71%（Verified），故無改善證據；高案 +0.5 點/年僅為敏感度")
add("U_w26", "代理工作負載乘數（占 AI 前用量）— 基期 FY2027（對應 unit_economics 的 2026）", "比例", 0.01, 0.04, 0.08, AS, "公司稱 AI 產品貢獻成長加速約一半（Q2 FY27 CEO 口述，Interested-party）；AI 年化營收 100M 於 Q3 FY26（管理層口述，主要為產品組合）；與 unit_economics 同值")
add("U_e26", "工作負載被替代（模型直接讀資料、開放格式外移）— 基期 FY2027", "比例", 0.0, 0.01, 0.03, AS, "無直接量測；與 unit_economics 同值")
add("U_j26", "AI 功能呼叫數（每 1 美元 AI 前用量）— 基期 FY2027", "次/$", 0.1, 0.5, 1.0, AS, "任務固定為 Tokenomics 任務 2「推理聊天」；與 unit_economics 同值；Snowflake 未揭露呼叫數")
add("U_j32", "AI 功能呼叫數 — FY2032", "次/$", 1.0, 4.0, 10.0, AS, "與 unit_economics 同值；使用 Snowflake AI 功能的帳號 7,300（Q3 FY26）→9,100（Q4）→13,600（Q1 FY27，口徑放寬；Interested-party）只顯示採用方向")
add("g_b28", "AI 前基準消耗年增 FY2028（扣除 AI 工作負載與替代）", "比例", 0.17, 0.22, 0.26, AS, "上半年 FY2027 產品營收年增 34–37%（Verified），其中 AI 約一半的加速（Interested-party）；NRR 125–126%（Verified）；基準 22% 為假設")
add("g_b32", "AI 前基準消耗年增 FY2032", "比例", 0.06, 0.10, 0.14, AS, "線性收斂；無證據，判斷題")
add("ps_g", "專業服務及其他逐年成長（FY2028 起）", "比例/年", -0.05, 0.03, 0.08, AS, "FY2022–FY2026 年增 61%／10%／17%／29%（Derived）；占營收約 4.5%，影響小")
add("gm_ps", "專業服務及其他非 GAAP 毛利率", "比例", -0.05, 0.0, 0.05, DE, "FY2026 非 GAAP 合計毛利 3,388.431 減產品毛利 3,388.382＝0.049（約 0%）；FY2025 3.3%、FY2024 5.0%（Derived，" + U + "000162828026011631/fy2026q4earnings.htm）")
# ---- 費用與現金流
add("yld", "現金與投資的利息收益率（用於非 GAAP 其他收入）", "比例", 0.03, 0.04, 0.05, AS, "FY2026 利息收入 190.6 對現金與投資（Derived）；現金餘額不隨年度滾動，為簡化")
add("sh_ng", "FY2027 非 GAAP 稀釋加權股數（僅用於 EPS 顯示）", "M股", 376, 380, 384, IP, "公司 FY2027 非 GAAP 稀釋加權股數指引 380M（含可轉債 if-converted 約 10M 股；Q2 FY27 新聞稿）", "公司指引")
add("sh_g", "FY2028 起每年稀釋股數增加（僅 EPS 顯示）", "比例/年", 0.0, 0.01, 0.03, AS, "FY2024–FY2026 非 GAAP 稀釋股數 362.1／363.4／372.3（Verified）；年增約 0.4–2.5%")
add("gaap_adj", "GAAP 與非 GAAP 營業利益差異中 SBC 以外的調整占營收（僅參考列）", "比例", 0.04, 0.06, 0.07, DE, "FY2026：非 GAAP 營業利益 489.7 減 SBC 1,599.5 減 GAAP 營業損失（−1,435.2）＝325.3，占營收 6.9%（Derived）；僅用於參考列的 GAAP 營業利益近似，不進估值")
# ---- Tokenomics v22.0 快照（只讀，不自行估 token 價格）
TIERS = ["Luna", "Terra", "Sol", "Astra"]
for t in TIERS:
    v = lambda n: TKI[n]["values"]
    add(f"IF_PriceFresh_{t}", f"新鮮輸入 $/M — {t}（OpenAI 有效）", "US$/M token", None, v(f"IF_PriceFresh_{t}"), None, AN, TKS + " IF_PriceFresh_" + t)
    add(f"IF_PriceCached_{t}", f"快取輸入 $/M — {t}", "US$/M token", None, v(f"IF_PriceCached_{t}"), None, AN, TKS + " IF_PriceCached_" + t)
    add(f"IF_PriceOut_{t}", f"可見輸出與思考 $/M — {t}（思考依輸出計費）", "US$/M token", None, v(f"IF_PriceOut_{t}"), None, AN, TKS + " IF_PriceOut_" + t + "（IF_PriceThink 同值）")
    add(f"IF_TaskSucc_{t}", f"任務 2 成功率 — {t}", "比例", None, v(f"IF_TaskSucc_{t}")[1], None, AN, TKS + " IF_TaskSucc_" + t + " 第 2 項")
    add(f"IF_FullCost_{t}", f"全成本 $/M — {t}（VR200 基準；對照用，不進模型）", "US$/M token", None, v(f"IF_FullCost_{t}")["VR200 NVL72"]["基準"], None, AN, TKS + " IF_FullCost_" + t + "（自下而上攤提；參考請求混合口徑，只作量級對照）")
add("IF_TaskTokFresh", "任務 2「推理聊天」新鮮輸入 token", "token", None, TKI["IF_TaskTokFresh"]["values"][1], None, AN, TKS + " IF_TaskTokFresh 第 2 項（用途欄＝情境、相對比較、反轉門檻；信心 B）")
add("IF_TaskTokCached", "任務 2 快取輸入 token", "token", None, TKI["IF_TaskTokCached"]["values"][1], None, AN, TKS + " IF_TaskTokCached 第 2 項")
add("IF_TaskTokDec", "任務 2 decode token（思考＋可見輸出）", "token", None, TKI["IF_TaskTokDec"]["values"][1], None, AN, TKS + " IF_TaskTokDec 第 2 項")
SC = {
 "U_w32": ("代理工作負載乘數 — FY2032", "比例", 0.40, 0.25, 0.10, AS, "unit_economics 同名參數 2032 為 10%／40%／90%（低／基準／高）；Snowflake 以基準 25%、樂觀 40%、保守 10% 分化：AI 營收占比與採用速度為最大不確定性，無公司揭露的 AI 占比（未找到）", ""),
 "U_e32": ("工作負載被替代 — FY2032", "比例", 0.04, 0.08, 0.15, AS, "unit_economics 2032 為 0%／8%／20%；Snowflake 自身開放表格式（Iceberg）與模型直讀資料的外移速度為假設；無量測", ""),
 "U_mk": ("AI 功能轉售加成（售價÷推論成本−1）", "比例", 0.5, 0.3, 0.1, AS, "Cortex 價目：claude-sonnet-4-5 輸入 1.8／輸出 9.0 AI Credit 每百萬 token，AI Credit 全域 US$2＝US$3.6／18 每百萬 token（公司公開價目，Verified）；模型商牌價本版未取，故加成為假設；unit_economics 區間 0%／30%／100%", ""),
 "sbc_v": ("股票薪酬：變動部分（每增加 1 美元營收增加的股票薪酬）", "比例", 0.05, "=Q!$B$@@SBCV@@", 0.15, DE, "基準＝Q 頁最近三點迴歸斜率（FY2025、FY2026、LTM；Derived）；樂觀 5%、保守 15% 為假設（管理層稱目標 Q4 FY2028 達 GAAP 營業獲利，Interested-party，原文時點待核）；全樣本斜率 25.7% 為更弱的規模經濟，另列 Sens；GAAP 調整路徑的分化", ""),
 "tp_w": ("情境權重", "比例", 0.25, 0.5, 0.25, AS, "原則第 4 節預設 25／50／25", ""),
}

ws = SH["In"]
ws["A1"] = "Snowflake 個股模型 v0.1 輸入頁（藍字＝輸入；黑字＝公式）。金額 US$M，會計年度 2 月制（FY2027＝2026-02-01～2027-01-31）。"
ws["A2"] = "基準欄 E 進主模型；D／F（敏感度低／高）只用於 Sens 頁一次動一個參數，不進目標價。情境專屬列（鍵尾 A／B／C）由「情境驅動對照」節並列。本版未做先驗。"
for j, h in enumerate(["鍵", "項目", "單位", "敏感度低", "基準", "敏感度高", "標記", "來源", "來源誘因"], 1):
    c = ws.cell(3, j, h); c.font = BOLD; c.fill = HDR
IR = {}
r = 4
for key, label, unit, lo, base, hi, tag, src, inc in INP:
    IR[key] = r
    ws.cell(r, 1, key); ws.cell(r, 2, label); ws.cell(r, 3, unit)
    for col, v in ((4, lo), (5, base), (6, hi)):
        if v is not None:
            c = ws.cell(r, col, v); c.font = BLUE
            if isinstance(v, datetime.date): c.number_format = "yyyy-mm-dd"
    ws.cell(r, 7, tag); ws.cell(r, 8, src); ws.cell(r, 9, inc)
    r += 1
# ---- 標準版 DCF 參數（tools/fcff_dcf.py in_rows；定義名稱 key／key_lo／key_hi）
_I = lambda k: f"In!$E${IR[k]}"
PFD = dict(
    date="2026-11-13", ccy="US$", amt="US$M", shu="M股",
    rf=(0.0499, "沿用 repo meta/README.md 所載 2026-10-27 10 年美債 4.99%；距股價日 2026-11-13 逾兩個交易日，標逾期", "本次未重讀財政部頁面", "無風險利率（10 年美債；2026-10-27，逾期）"),
    betas=[("StockAnalysis（5 年）", 1.32, SA + "（beta 1.32）；第二來源未找到；docs/plan/自算產業β_20261105.md 的軟體產業 β 本版未套用")],
    beta_lohi=(1.0, 1.5, "單一來源，區間為 Assumed"),
    erp=(0.04, 0.05, 0.06, "DCF 標準規格預設 5%（4–6%）"),
    kd=("=rf+" + _I("kd_sp"), "Derived", "rf＋信用利差（kd_sp）", "可轉債 0 票息；若轉換則債務為 0"),
    tax_w=("=" + _I("tax"), "Derived", "同 In 頁 tax（非 GAAP 隱含稅率）", "—"),
    debt=("=IF(" + _I("px") + ">" + _I("conv_px") + ",0," + _I("conv_prin") + ")", "Derived", "可轉債本金；現價高於轉換價 157.50 時依 if-converted 視為股權（債務 0，股數加項見 sh_v）", Q2Q),
    lease=(0, "Assumed", "融資租賃負債未找到，以 0 計；營業租賃 479.8（流動 59.787＋非流動 420.043，10-Q 資產負債表）依規格不計入淨債務", "—"),
    cash=(4329.177, "Verified", Q2Q + " 資產負債表：現金及約當現金 1,707.187＋短期投資 637.508＋長期投資 1,984.482", "—"),
    inv=(0, "Assumed", "非營業權益投資未另列（已含於現金與投資）", "—"),
    mi=(0, "Assumed", "無少數股權（10-Q 損益表淨損失歸屬他方金額極小，未計）", "—"),
    sh=("=" + _I("sh_out") + "+" + _I("rsu") + "+" + _I("opt_n") + "*MAX(0,1-" + _I("opt_k") + "/" + _I("px") + ")+IF(" + _I("px") + ">" + _I("conv_px") + "," + _I("conv_sh") + "-" + _I("conv_sh") + "*(MIN(" + _I("px") + "," + _I("cap_px") + ")-" + _I("conv_px") + ")/" + _I("px") + ",0)",
        "Derived", "評價日股數＝流通股 352.8＋RSU（全額）＋選擇權庫藏股法＋可轉債 if-converted 股數減 capped call 回收（DCF 規格第 3 節）", "RSU 沒收率未揭露（未找到），全額計入；ASR 未建模"),
    fc_da=(0.035, 0.045, 0.055, "Assumed", "FY2026 220.443÷4,683.9＝4.7%、FY2025 5.0%（Derived）；含被併購無形資產攤銷，高估非 GAAP 內含折舊的部分", "—"),
    fc_cx=(0.015, 0.022, 0.03, "Derived", "FY2026 101.628÷4,683.9＝2.2%、FY2025 2.1%、FY2024 2.5%（營業現金流減公司自由現金流；10-K MD&A）", "—"),
    fc_da_ss=(0.03, "Assumed", "長期穩態：被併購無形資產攤銷退出後，折舊攤銷收斂到資本支出附近（FY2030–FY2032 線性收斂）", "—"),
    fc_cx_ss=(0.031, "Assumed", "長期穩態：資本支出占營收略高於折舊（終值要求資本支出 ≥ 折舊×(1+g)）", "—"),
    fc_wc=(-0.452, -0.363, -0.25, "Derived", "營運資金增加÷營收增量，負值＝流入（原則第 23 條）：10-K FY2026 MD&A 營運資產負債變動合計流入 383.8（FY2026；其中遞延收入 +755.2、遞延佣金 −305.1、應收 −380.0、應計 +393.3）÷營收增量 1,057.6＝36.3%；FY2025 443.6÷819.9＝54.1%；基準取最新年，低案取兩年平均，高案 25% 為假設", "FY2024 未找到"),
)
FDROWS = FD.in_rows(PFD)
for tup in FDROWS:
    if tup[0] == "sec":
        ws.cell(r, 1, tup[1]).font = BOLD; r += 1; continue
    key, label, unit, lo, base, hi, tag, s1, s2 = tup
    IR[key] = r
    ws.cell(r, 1, key); ws.cell(r, 2, label); ws.cell(r, 3, unit)
    for col, v in ((4, lo), (5, base), (6, hi)):
        if v is not None:
            c = ws.cell(r, col, v); c.font = BLUE
    ws.cell(r, 7, tag); ws.cell(r, 8, s1 + ("" if s2 in ("—", "") else "；" + s2)); ws.cell(r, 9, s2 if tag in ("Interested-party",) else "")
    for nm, colx in ((key, "E"), (key + "_lo", "D"), (key + "_hi", "F")):
        wb.defined_names[nm] = DefinedName(nm, attr_text=f"In!${colx}${r}")
    r += 1
SCK = {}
for _s in "ABC": pass
for pre, (label, unit, a, b, c_, tag, src, inc) in SC.items():
    for s, v in zip("ABC", (a, b, c_)):
        k = f"{pre}{s}" if pre == "tp_w" else f"{pre}_{s}"
        SCK.setdefault(pre, []).append(k)
        IR[k] = r
        if pre == "tp_w": wb.defined_names[k] = DefinedName(k, attr_text=f"In!$E${r}")
        ws.cell(r, 1, k); ws.cell(r, 2, f"情境 {s}：{label}"); ws.cell(r, 3, unit)
        ws.cell(r, 5, v).font = BLUE
        ws.cell(r, 7, tag); ws.cell(r, 8, src); ws.cell(r, 9, inc)
        r += 1
LAST_PARAM_ROW = max(v for k, v in IR.items() if not k.startswith(("U_w32_", "U_e32_", "U_mk_", "sbc_v_", "tp_w")))
r += 1
ws.cell(r, 1, "J. 情境驅動對照（純展示；A 樂觀／B 基準／C 保守）").font = BOLD
r += 1
for j, h in enumerate(["鍵", "項目", "單位", "情境 A", "情境 B", "情境 C", "標記", "說明"], 1):
    ws.cell(r, j, h).font = BOLD
SCN_H = r
r += 1
SCN_ROWS = {}
for pre, key in (("U_w32", "scn_w32"), ("U_e32", "scn_e32"), ("U_mk", "scn_mk"), ("sbc_v", "scn_sbc_v"), ("tp_w", "scn_w")):
    ks = SCK[pre]
    ws.cell(r, 1, key); ws.cell(r, 2, SC[pre][0]); ws.cell(r, 3, SC[pre][1])
    for j, k in zip((4, 5, 6), ks):
        ws.cell(r, j, f"=E{IR[k]}")
    ws.cell(r, 7, "Derived"); ws.cell(r, 8, "引用 In 頁 " + "／".join(ks) + "；三情境並排，只引用上方輸入列")
    SCN_ROWS[key] = r
    r += 1


def I(key):
    return f"In!$E${IR[key]}"


# ---------------------------------------------------------------- Q 頁
YEARS = ["FY2024", "FY2025", "FY2026", "FY2027", "FY2028", "FY2029", "FY2030", "FY2031", "FY2032"]
YC = {y: "CDEFGHIJK"[i] for i, y in enumerate(YEARS)}
HIST = ["FY2024", "FY2025", "FY2026"]
NF = "未找到"
K26 = "10-K FY2026（" + U + "000164014726000008/snow-20260131.htm）"
PR4 = "Q4 FY26 財報新聞稿（8-K 附件 99.1）" + U + "000162828026011631/fy2026q4earnings.htm"
PR24 = "Q4 FY24 財報新聞稿 " + U + "000164014724000031/fy2024q4earnings.htm"
Q = SH["Q"]
Q["A1"] = "Q 頁：公司法定申報與新聞稿的歷史數字（US$M；EPS 為美元）；FY2024–FY2026 為會計年度（2 月制）。每列單一標記；Interested-party 列有來源誘因。"
Q["A2"] = "非 GAAP 口徑為公司定義（排除股票薪酬與相關薪資稅、無形資產攤銷等）；公司未揭露非 GAAP 稅前與稅額，下方「其他收入」與「稅額」為推算（Derived）。FY2024 折舊攤銷與資本支出本次未讀到，標「未找到」。"
for j, h in enumerate(["指標", "FY2024", "FY2025", "FY2026", "單位", "來源", "標記", "來源誘因"], 1):
    Q.cell(3, j, h).font = BOLD
QR = {}
qdata = [
 ("prod", "Product revenue（產品營收）", 2666.849, 3462.422, 4472.317, "US$M", K26 + " Note 3", D0, ""),
 ("ps", "Professional services and other revenue", 139.640, 163.974, 211.629, "US$M", K26 + " Note 3", D0, ""),
 ("rev", "合併營收（公司揭露）", 2806.489, 3626.396, 4683.946, "US$M", K26 + " 損益表", D0, ""),
 ("cogs_g", "GAAP 營業成本", 898.558, 1214.673, 1537.805, "US$M", K26 + " 損益表", D0, ""),
 ("gp_g", "GAAP 毛利", 1907.931, 2411.723, 3146.141, "US$M", K26 + " 損益表", D0, ""),
 ("opx_g", "GAAP 營業費用", 3002.704, 3867.733, 4581.306, "US$M", K26 + " 損益表", D0, ""),
 ("oi_g", "GAAP 營業利益（損失）", -1094.773, -1456.010, -1435.165, "US$M", K26 + " 損益表", D0, ""),
 ("ni_g", "GAAP 淨利（損失）", -837.990, -1289.212, -1329.035, "US$M", K26 + " 損益表", D0, ""),
 ("eps_g", "GAAP 每股盈餘（US$）", -2.55, -3.86, -3.95, "US$", K26 + " 損益表", D0, ""),
 ("gp_n", "非 GAAP 毛利", 2082.899, 2648.711, 3388.431, "US$M", PR4 + " Reconciliations；FY2024 由 " + PR24 + " 推得（Derived）", IP, "公司定義之非 GAAP 數字"),
 ("opx_n", "非 GAAP 營業費用", 1853.189, 2416.988, 2898.713, "US$M", PR4 + " Reconciliations", IP, "公司定義之非 GAAP 數字"),
 ("oi_n", "非 GAAP 營業利益", 229.710, 231.723, 489.718, "US$M", PR4 + " Reconciliations", IP, "公司定義之非 GAAP 數字"),
 ("ni_n", "非 GAAP 淨利", 351.692, 299.991, 465.953, "US$M", PR4 + " Reconciliations", IP, "公司定義之非 GAAP 數字"),
 ("eps_n", "非 GAAP 稀釋每股盈餘（US$）", 0.98, 0.83, 1.25, "US$", PR4 + " Reconciliations", IP, "公司定義之非 GAAP 數字"),
 ("sh_n", "非 GAAP 稀釋加權股數（M）", 362.064, 363.374, 372.323, "M股", PR4 + " Reconciliations", IP, "公司定義之非 GAAP 數字"),
 ("sbc", "股票薪酬（現金流量表）", 1168.015, 1479.314, 1599.547, "US$M", K26 + " 現金流量表", D0, ""),
 ("da", "折舊與攤銷", NF, 182.508, 220.443, "US$M", K26 + " 現金流量表", D0, ""),
 ("capex", "資本支出（含資本化軟體；營業現金流減公司自由現金流）", 69.219, 75.712, 101.628, "US$M", K26 + " 現金流量表；FY2025 959.764−884.052、FY2026 1,221.942−1,120.314（Derived）", DE, ""),
 ("fcf", "公司口徑自由現金流", 778.903, 884.052, 1120.314, "US$M", PR4, IP, "公司定義之非 GAAP 數字"),
 ("inti", "利息收入", 200.663, 209.009, 190.556, "US$M", K26 + " 損益表", D0, ""),
 ("inte", "利息費用", 0, -2.759, -8.298, "US$M", K26 + " 損益表（FY2024 未單列，以 0 計）", D0, ""),
 ("othn", "其他收支淨額", 44.887, -35.339, -59.003, "US$M", K26 + " 損益表（FY2024 為「其他收入（費用）淨額」合併列）", D0, ""),
]
rr = 4
for key, lab, a, b, c_, unit, src, tag, inc in qdata:
    QR[key] = rr
    Q.cell(rr, 1, lab)
    for j, v in zip((2, 3, 4), (a, b, c_)):
        Q.cell(rr, j, v).font = BLUE
    Q.cell(rr, 5, unit); Q.cell(rr, 6, src); Q.cell(rr, 7, tag); Q.cell(rr, 8, inc)
    rr += 1
# 推算列
Q.cell(rr, 1, "其他收入（推算＝利息收入＋利息費用＋其他收支淨額；假設非 GAAP 不調整此三項）"); QR["oth"] = rr
for j, col in zip((2, 3, 4), "BCD"): Q.cell(rr, j, f"={col}{QR['inti']}+{col}{QR['inte']}+{col}{QR['othn']}")
Q.cell(rr, 5, "US$M"); Q.cell(rr, 6, "由上列相加"); Q.cell(rr, 7, DE); rr += 1
Q.cell(rr, 1, "非 GAAP 稅前（推算＝非 GAAP 營業利益＋其他收入）"); QR["ebt_n"] = rr
for j, col in zip((2, 3, 4), "BCD"): Q.cell(rr, j, f"={col}{QR['oi_n']}+{col}{QR['oth']}")
Q.cell(rr, 5, "US$M"); Q.cell(rr, 6, "由上列相加"); Q.cell(rr, 7, DE); rr += 1
Q.cell(rr, 1, "非 GAAP 稅額（推算＝非 GAAP 稅前−非 GAAP 淨利）"); QR["tax_n"] = rr
for j, col in zip((2, 3, 4), "BCD"): Q.cell(rr, j, f"={col}{QR['ebt_n']}-{col}{QR['ni_n']}")
Q.cell(rr, 5, "US$M"); Q.cell(rr, 6, "由上列相減；隱含稅率見 In 頁 tax"); Q.cell(rr, 7, DE); rr += 1
Q.cell(rr, 1, "核對：Product＋Professional services−合併營收（應為 0）"); QR["chk"] = rr
for j, col in zip((2, 3, 4), "BCD"): Q.cell(rr, j, f"={col}{QR['prod']}+{col}{QR['ps']}-{col}{QR['rev']}")
rr += 1
Q.cell(rr, 1, "核對：非 GAAP 毛利−營業費用−營業利益（應為 0）"); QR["chk2"] = rr
for j, col in zip((2, 3, 4), "BCD"): Q.cell(rr, j, f"={col}{QR['gp_n']}-{col}{QR['opx_n']}-{col}{QR['oi_n']}")
rr += 1
Q.cell(rr, 1, "核對：GAAP 營收−營業成本−毛利（應為 0）"); QR["chk3"] = rr
for j, col in zip((2, 3, 4), "BCD"): Q.cell(rr, j, f"={col}{QR['rev']}-{col}{QR['cogs_g']}-{col}{QR['gp_g']}")

# ---- 規模經濟迴歸與營運資金比率（原則 G3 第 1、2 條）
rr += 2
Q.cell(rr, 1, "規模經濟迴歸（FY2022–FY2027 實績；營業費用與股票薪酬＝固定＋變動×營收；US$M）").font = BOLD; rr += 1
RG = rr
for j, h in enumerate(["指標", "FY2022", "FY2023", "FY2024", "FY2025", "FY2026", "LTM（Q3 FY26–Q2 FY27）", "來源", "標記"], 1): Q.cell(rr, j, h).font = BOLD
rr += 1
QG = {}
reg = [("rev", "營收", [1219.327, 2065.659, "=B%d" % 0, None, None, 5434.653], "10-K FY2022／FY2024／FY2026 損益；LTM＝四季營收 1,212.909＋1,283.994＋1,390.951＋1,546.793（8-K 附件 99.1）"),
       ("pgp", "非 GAAP 產品毛利（FY2022、FY2023、LTM；專業服務非 GAAP 毛利約 0，以產品毛利代理合計）", [844.623, 1457.382, None, None, None, 3917.406], "FY2022、FY2023 非 GAAP 調節表（facts_snowflake.csv）；LTM＝879.165＋921.462＋1,002.695＋1,114.084"),
       ("oi", "非 GAAP 營業利益", [-31.175, 95.283, None, None, None, 673.241], "同上；LTM＝131.313＋139.175＋165.753＋237.0"),
       ("opx", "非 GAAP 營業費用（＝毛利−營業利益）", [None] * 6, "FY2022、FY2023、LTM＝產品毛利−營業利益（Derived）；FY2024–FY2026 取上方 Q 表"),
       ("sbc", "股票薪酬", [605.095, 861.533, None, None, None, 1641.922], "FY2022、FY2023 MD&A SBC 表；LTM＝412.278＋403.592＋402.47＋423.582（Derived）")]
for key, lab, vals, src in reg:
    QG[key] = rr
    Q.cell(rr, 1, lab)
    for j, v in enumerate(vals):
        if v is not None and not (isinstance(v, str)): Q.cell(rr, 2 + j, v).font = BLUE
    Q.cell(rr, 8, src); Q.cell(rr, 9, "Derived" if key in ("opx",) else "Interested-party" if key in ("pgp", "oi") else "Verified")
    rr += 1
for j, col in zip((4, 5, 6), "BCD"):   # FY2024–FY2026 取上方 Q 表
    Q.cell(QG["rev"], j, f"={col}{QR['rev']}"); Q.cell(QG["opx"], j, f"={col}{QR['opx_n']}"); Q.cell(QG["sbc"], j, f"={col}{QR['sbc']}")
for j, col in zip((2, 3, 7), "BCG"):
    Q.cell(QG["opx"], j, f"={col}{QG['pgp']}-{col}{QG['oi']}")
rr += 1
X3 = f"E{QG['rev']}:G{QG['rev']}"; X6 = f"B{QG['rev']}:G{QG['rev']}"
RS = {}
for key, lab, f in (("opxv", "營業費用變動部分（最近三點斜率）", f"=SLOPE(E{QG['opx']}:G{QG['opx']},{X3})"), ("opxa", "營業費用固定部分（最近三點截距；US$M）", f"=INTERCEPT(E{QG['opx']}:G{QG['opx']},{X3})"),
                    ("opxv6", "營業費用變動部分（六點全樣本斜率）", f"=SLOPE(B{QG['opx']}:G{QG['opx']},{X6})"), ("opxa6", "營業費用固定部分（六點全樣本截距）", f"=INTERCEPT(B{QG['opx']}:G{QG['opx']},{X6})"),
                    ("opxltm", "營業費用率（LTM；全期持平案用）", f"=G{QG['opx']}/G{QG['rev']}"),
                    ("sbcv", "股票薪酬變動部分（最近三點斜率）", f"=SLOPE(E{QG['sbc']}:G{QG['sbc']},{X3})"), ("sbca", "股票薪酬固定部分（最近三點截距；US$M）", f"=INTERCEPT(E{QG['sbc']}:G{QG['sbc']},{X3})"),
                    ("sbcv6", "股票薪酬變動部分（六點全樣本斜率）", f"=SLOPE(B{QG['sbc']}:G{QG['sbc']},{X6})"), ("sbcltm", "股票薪酬占營收（LTM；全期持平案用）", f"=G{QG['sbc']}/G{QG['rev']}")):
    RS[key] = rr; Q.cell(rr, 1, lab); Q.cell(rr, 2, f); Q.cell(rr, 7, DE); rr += 1
rr += 1
Q.cell(rr, 1, "營運資金（10-K FY2026 MD&A 營運資產負債變動；負值＝流入，原則第 23 條）").font = BOLD; rr += 1
WCR = rr
for j, h in enumerate(["指標", "FY2025", "FY2026", "來源", "標記"], 1): Q.cell(rr, j, h).font = BOLD
rr += 1
Q.cell(rr, 1, "營運資產負債變動合計（現金流入；US$M）"); Q.cell(rr, 2, 443.6).font = BLUE; Q.cell(rr, 3, 383.8).font = BLUE
Q.cell(rr, 4, K26 + " MD&A Liquidity：FY2026 合計 383.8（遞延收入 +755.2、遞延佣金 −305.1、應收 −380.0、應計 +393.3、營運租賃 −26.9、預付 −44.5）；FY2025 443.6"); Q.cell(rr, 5, D0); WC1 = rr; rr += 1
Q.cell(rr, 1, "營收增量（US$M）"); Q.cell(rr, 2, f"=C{QR['rev']}-B{QR['rev']}"); Q.cell(rr, 3, f"=D{QR['rev']}-C{QR['rev']}"); Q.cell(rr, 5, DE); rr += 1
Q.cell(rr, 1, "營運資金增加÷營收增量（負值＝流入）"); Q.cell(rr, 2, f"=-B{WC1}/B{WC1+1}"); Q.cell(rr, 3, f"=-C{WC1}/C{WC1+1}"); Q.cell(rr, 5, DE); WCRAT = rr; rr += 1
Q.cell(rr, 1, "其中遞延收入增加（FY2026；US$M）"); Q.cell(rr, 3, 755.2).font = BLUE; Q.cell(rr, 4, K26 + " MD&A"); Q.cell(rr, 5, D0); rr += 1
Q.cell(rr, 1, "核對：In 頁 fc_wc 基準 − FY2026 比率（應為 0；容許四捨五入 0.001）"); Q.cell(rr, 3, f"=ROUND(In!$E${IR['fc_wc']}-C{WCRAT},3)"); WCCK = rr; rr += 1


def QH(key, y):
    return f"Q!${'BCD'[HIST.index(y)]}${QR[key]}"


# ---------------------------------------------------------------- Seg 頁
S = SH["Seg"]
S["A1"] = "命題模組（原型 U：按用量計價）：產品營收＝核心平台消耗＋AI 工作負載。AI 工作負載＝基準消耗 B×代理工作負載乘數 w＋AI 功能轉售（呼叫數 j×每次推論成本×(1＋轉售加成 U_mk)）；侵蝕＝B×被替代比例 e；推論成本＝B×j×每次成功呼叫市場價（Tokenomics v22.0 任務 2「推理聊天」）。AI 淨效果＝AI 營收增量−核心侵蝕−推論成本。"
S["A2"] = "FY2027 為上半年實績加 Q3、Q4 假設，三情境共用產品營收錨點；情境分化自 FY2028 起（w、e、轉售加成、SBC 占比）。變數名稱與 software_ai/unit_economics 一致；v0.2 改讀 unit_economics。基準消耗 B 是「沒有 AI 工作負載與替代時」的消耗路徑，FY2027 由錨點反解。"
S["A4"] = "Tokenomics v22.0 任務 2「推理聊天」每次成功呼叫的成本（US$/次；2026 快照；市場價＝逐類 token×OpenAI 有效單價÷成功率；全成本＝總 token÷成功率×IF_FullCost（VR200 基準），只作量級對照）"
S["A5"] = "層級"; S["B5"] = "市場價（US$/次）"; S["C5"] = "全成本對照（US$/次）"
for i, t in enumerate(TIERS):
    rw = 6 + i
    S.cell(rw, 1, f"{i+1} {t}")
    S.cell(rw, 2, f"=({I('IF_TaskTokFresh')}*{I('IF_PriceFresh_'+t)}+{I('IF_TaskTokCached')}*{I('IF_PriceCached_'+t)}+{I('IF_TaskTokDec')}*{I('IF_PriceOut_'+t)})/1000000/{I('IF_TaskSucc_'+t)}")
    S.cell(rw, 3, f"=({I('IF_TaskTokFresh')}+{I('IF_TaskTokCached')}+{I('IF_TaskTokDec')})/1000000/{I('IF_TaskSucc_'+t)}*{I('IF_FullCost_'+t)}")
S["A10"] = "所選層級（U_tier）"; S["B10"] = f"=INDEX(B6:B9,{I('U_tier')})"; S["C10"] = f"=INDEX(C6:C9,{I('U_tier')})"
S["A12"] = "FY2027 產品營收錨點（Q1、Q2 實績；Q3、Q4 為假設；公司指引不進基準，只在 Recon 對照）"
for j, h in enumerate(["Q1A", "Q2A", "Q3E", "Q4E", "FY2027E"], 2): S.cell(12, j, h)
S["A13"] = "產品營收（US$M）"
S["B13"] = f"={I('q1p')}"; S["C13"] = f"={I('q2p')}"; S["D13"] = f"=C13*(1+{I('q3g')})"; S["E13"] = f"=D13*(1+{I('q4g')})"; S["F13"] = "=SUM(B13:E13)"
S["A14"] = "專業服務及其他營收（US$M；下半年＝上半年×比值）"
S["B14"] = f"={I('q1s')}"; S["C14"] = f"={I('q2s')}"; S["D14"] = f"=(B14+C14)*{I('ps_h2r')}/2"; S["E14"] = f"=(B14+C14)*{I('ps_h2r')}/2"; S["F14"] = "=SUM(B14:E14)"
SEGLAB = ["n", "gb", "B", "w", "e", "j", "pxidx", "cps", "infc", "resale", "wl", "air", "eros", "core", "prod", "ps", "rev", "net", "aish", "inf_full", "gm0n", "cogsx", "pgm", "chk1", "chk2"]
LABS = {"n": "年數（FY2027＝0）", "gb": "基準消耗年增（AI 前）", "B": "基準消耗 B（AI 前的消耗路徑；US$M）", "w": "代理工作負載乘數 w", "e": "工作負載被替代比例 e", "j": "AI 功能呼叫數 j（次/美元）",
        "pxidx": "路徑指數（ovh×((1+g_px)(1+g_tok))^n）", "cps": "每次成功呼叫推論成本（市場價；US$）", "infc": "推論成本（市場價；US$M）", "resale": "AI 功能轉售營收（推論成本×(1＋U_mk)）",
        "wl": "代理工作負載營收（B×w）", "air": "AI 營收增量（命題子列；工作負載＋轉售）", "eros": "核心侵蝕（B×e）", "core": "核心平台消耗（命題子列；B−侵蝕）", "prod": "Product revenue（產品營收）",
        "ps": "Professional services and other revenue", "rev": "合併營收", "net": "AI 淨效果（AI 營收增量−核心侵蝕−推論成本）", "aish": "AI 營收占產品營收",
        "inf_full": "對照：以全成本計的推論成本（不進模型；量級對照）", "gm0n": "核心產品毛利率（U_gm0＋逐年變化）", "cogsx": "產品營業成本：核心與工作負載（不含推論；US$M）", "pgm": "產品毛利率（非 GAAP）",
        "chk1": "檢查：核心平台消耗＋AI 營收增量−產品營收", "chk2": "檢查：產品營收＋專業服務−合併營收"}
SR = {}
BLK0 = 17
for si, s in enumerate("ABC"):
    r0 = BLK0 + si * (len(SEGLAB) + 4)
    S.cell(r0, 1, f"情境 {s}（{ {'A':'樂觀：AI 工作負載高、被替代低、轉售加成高','B':'基準','C':'保守：AI 工作負載低、被替代高、轉售加成低'}[s] }）").font = BOLD
    rows = {k: r0 + 2 + i for i, k in enumerate(SEGLAB)}
    SR[s] = rows | {"hdr": r0 + 1}
    for y, c in YC.items():
        S[f"{c}{r0+1}"] = y + ("A" if y in HIST else "E"); S[f"{c}{r0+1}"].font = BOLD
    S.cell(r0 + 1, 1, "US$M 除另註")
    for k, rw in rows.items(): S.cell(rw, 1, LABS[k])
    for y, c in YC.items():
        p = chr(ord(c) - 1)
        R = lambda k: f"{c}{rows[k]}"
        Pv = lambda k: f"{p}{rows[k]}"
        if y in HIST:
            S[R("prod")] = f"={QH('prod', y)}"; S[R("ps")] = f"={QH('ps', y)}"; S[R("rev")] = f"={QH('rev', y)}"
            S[R("chk2")] = f"={R('prod')}+{R('ps')}-{R('rev')}"
            for k in ("B", "w", "e", "j", "core", "air", "eros", "net"): S[R(k)] = "未找到（AI 營收未單列）" if k in ("air", "core") else None
            continue
        n = YEARS.index(y) - 3
        S[R("n")] = n
        S[R("w")] = f"={I('U_w26')}+({I('U_w32_'+s)}-{I('U_w26')})*{R('n')}/5"
        S[R("e")] = f"={I('U_e26')}+({I('U_e32_'+s)}-{I('U_e26')})*{R('n')}/5"
        S[R("j")] = f"={I('U_j26')}+({I('U_j32')}-{I('U_j26')})*{R('n')}/5"
        S[R("pxidx")] = f"={I('ovh')}*((1+{I('g_px')})*(1+{I('g_tok')}))^{R('n')}"
        S[R("cps")] = f"=$B$10*{R('pxidx')}"
        if y == "FY2027":
            S[R("B")] = f"=$F$13/(1+{R('w')}-{R('e')}+{R('j')}*{R('cps')}*(1+{I('U_mk_'+s)}))"
            S[R("ps")] = "=$F$14"
        else:
            S[R("gb")] = f"={I('g_b28')}+({I('g_b32')}-{I('g_b28')})*({R('n')}-1)/4"
            S[R("B")] = f"={Pv('B')}*(1+{R('gb')})"
            S[R("ps")] = f"={Pv('ps')}*(1+{I('ps_g')})"
        S[R("infc")] = f"={R('B')}*{R('j')}*{R('cps')}"
        S[R("resale")] = f"={R('infc')}*(1+{I('U_mk_'+s)})*((1+{I('g_px_rs')})/(1+{I('g_px')}))^{R('n')}"
        S[R("wl")] = f"={R('B')}*{R('w')}"
        S[R("air")] = f"={R('wl')}+{R('resale')}"
        S[R("eros")] = f"={R('B')}*{R('e')}"
        S[R("core")] = f"={R('B')}-{R('eros')}"
        S[R("prod")] = f"={R('core')}+{R('air')}"
        S[R("rev")] = f"={R('prod')}+{R('ps')}"
        S[R("net")] = f"={R('air')}-{R('eros')}-{R('infc')}"
        S[R("aish")] = f"={R('air')}/{R('prod')}"
        S[R("inf_full")] = f"={R('B')}*{R('j')}*$C$10*((1+{I('g_tok')})^{R('n')})*{I('ovh')}"
        S[R("gm0n")] = f"={I('U_gm0')}+{I('gm_dr')}*{R('n')}"
        S[R("cogsx")] = f"={R('B')}*(1+{R('w')}-{R('e')})*(1-{R('gm0n')})"
        S[R("pgm")] = f"=1-({R('cogsx')}+{R('infc')})/{R('prod')}"
        S[R("chk1")] = f"={R('core')}+{R('air')}-{R('prod')}"
        S[R("chk2")] = f"={R('prod')}+{R('ps')}-{R('rev')}"
SEG_END = BLK0 + 3 * (len(SEGLAB) + 4)
S.cell(SEG_END, 1, "註：推論成本只讀 Tokenomics v22.0 快照（市場價）；未來單價路徑 g_px 待 AI供應鏈 登錄（Assumed 區間）。全成本為「模型商自己做這件事的成本下限」對照，不進損益。").font = BOLD

# ---------------------------------------------------------------- PL 頁
P = SH["PL"]
P["A1"] = "損益表（非 GAAP 口徑，公司與賣方共識的口徑；單一官方營運分部：10-K 附註 one operating segment）：營收列取公司固定揭露的 Product revenue 與 Professional services and other revenue；命題子列為核心平台消耗與 AI 營收增量。US$M。"
P["A2"] = "非 GAAP 排除股票薪酬與相關薪資稅、無形資產攤銷等。DCF 的 EBIT＝非 GAAP 營業利益−股票薪酬（SBC 不加回；非 GAAP→FCFF 橋接見各情境區塊）。GAAP 營業利益與 EPS 為參考列（本版 GAAP 預測為近似，不進估值）。FY2024A–FY2026A 取自 Q 頁。"
PLB = ["rev", "cogs", "gp", "opex", "oi", "sbc", "ebit", "da", "capex", "nopat", "dwc", "fcff", "oth", "ebt", "taxx", "ni", "shs", "eps", "gaap_oi", "fcf_co"]
PLL = {"rev": "營收", "cogs": "營業成本（非 GAAP）", "gp": "毛利（非 GAAP）", "opex": "營業費用（非 GAAP）", "oi": "營業利益（非 GAAP）", "sbc": "股票薪酬",
       "ebit": "DCF 的 EBIT（非 GAAP 營業利益−股票薪酬）", "da": "折舊與攤銷", "capex": "資本支出", "nopat": "NOPAT（EBIT×(1−稅率)；取自 DCF 頁）",
       "dwc": "營運資金增加（負值＝流入；取自 DCF 頁）", "fcff": "無槓桿自由現金流 FCFF（取自 DCF 頁）", "oth": "其他收入（利息等；非 GAAP）", "ebt": "稅前（非 GAAP）", "taxx": "所得稅（非 GAAP）",
       "ni": "淨利（非 GAAP）", "shs": "稀釋股數（M；非 GAAP）", "eps": "EPS（非 GAAP 稀釋，美元）", "gaap_oi": "參考：GAAP 營業利益（近似＝非 GAAP 營業利益−SBC−其他調整）", "fcf_co": "參考：公司口徑自由現金流（FCFF＋SBC＋稅後其他收入）"}
PR = {}
BASE0 = 70
for si, s in enumerate("ABC"):
    r0 = BASE0 + si * (len(PLB) + 4)
    P.cell(r0, 1, f"情境 {s} 損益與現金流").font = BOLD
    for y, c in YC.items():
        P[f"{c}{r0+1}"] = y + ("A" if y in HIST else "E"); P[f"{c}{r0+1}"].font = BOLD
    rows = {k: r0 + 2 + i for i, k in enumerate(PLB)}
    PR[s] = rows | {"hdr": r0 + 1}
    for k, rw in rows.items(): P.cell(rw, 1, f"情境 {s}｜{PLL[k]}")
    sr = SR[s]
    for y, c in YC.items():
        p = chr(ord(c) - 1)
        R = lambda k: f"{c}{rows[k]}"
        Pv = lambda k: f"{p}{rows[k]}"
        if y in HIST:
            P[R("rev")] = f"={QH('rev', y)}"; P[R("gp")] = f"={QH('gp_n', y)}"; P[R("cogs")] = f"={R('rev')}-{R('gp')}"
            P[R("opex")] = f"={QH('opx_n', y)}"; P[R("oi")] = f"={QH('oi_n', y)}"; P[R("sbc")] = f"={QH('sbc', y)}"
            P[R("ebit")] = f"={R('oi')}-{R('sbc')}"
            P[R("da")] = f"={QH('da', y)}"; P[R("capex")] = f"={QH('capex', y)}"
            P[R("oth")] = f"={QH('oth', y)}"; P[R("ebt")] = f"={QH('ebt_n', y)}"; P[R("taxx")] = f"={QH('tax_n', y)}"; P[R("ni")] = f"={QH('ni_n', y)}"
            P[R("shs")] = f"={QH('sh_n', y)}"; P[R("eps")] = f"={QH('eps_n', y)}"; P[R("gaap_oi")] = f"={QH('oi_g', y)}"
            continue
        n = YEARS.index(y) - 3
        P[R("rev")] = f"=Seg!{c}{sr['rev']}"
        P[R("cogs")] = f"=Seg!{c}{sr['cogsx']}+Seg!{c}{sr['infc']}+Seg!{c}{sr['ps']}*(1-{I('gm_ps')})"
        P[R("gp")] = f"={R('rev')}-{R('cogs')}"
        P[R("opex")] = f"={I('opx_a')}+{I('opx_v')}*{R('rev')}"
        P[R("oi")] = f"={R('gp')}-{R('opex')}"
        P[R("sbc")] = f"=({I('sbc_a')}+{I('sbc_v_B')}*$F{rows['rev']})+{I('sbc_v_'+s)}*({R('rev')}-$F{rows['rev']})"
        P[R("ebit")] = f"={R('oi')}-{R('sbc')}"
        P[R("da")] = f"={R('rev')}*({I('fc_da')}+({I('fc_da_ss')}-{I('fc_da')})*MAX(0,{n}-2)/3)"
        P[R("capex")] = f"={R('rev')}*({I('fc_cx')}+({I('fc_cx_ss')}-{I('fc_cx')})*MAX(0,{n}-2)/3)"
        P[R("oth")] = f"={I('yld')}*{I('cash_v')}"
        P[R("ebt")] = f"={R('oi')}+{R('oth')}"
        P[R("taxx")] = f"={R('ebt')}*{I('tax')}"
        P[R("ni")] = f"={R('ebt')}-{R('taxx')}"
        P[R("shs")] = f"={I('sh_ng')}" if y == "FY2027" else f"={Pv('shs')}*(1+{I('sh_g')})"
        P[R("eps")] = f"={R('ni')}/{R('shs')}"
        P[R("gaap_oi")] = f"={R('oi')}-{R('sbc')}-{I('gaap_adj')}*{R('rev')}"
        P[R("fcf_co")] = f"={R('fcff')}+{R('sbc')}+{R('oth')}*(1-{I('tax')})"
PL_END = BASE0 + 3 * (len(PLB) + 4)
b = PR["B"]
STD = ["營收", "營收年增率", "營業成本", "毛利", "毛利率", "營業費用", "營業利益", "營業利益率", "其他收入", "稅前利益", "所得稅", "稅後淨利", "淨利率", "EPS（非 GAAP 稀釋，美元）", "淨利年增率"]
P.cell(4, 1, "標準損益表（非 GAAP；US$M，EPS 為美元；情境 B）：FY2024A–FY2026A 實績、FY2027E–FY2032E 預測").font = BOLD
for y, c in YC.items():
    P[f"{c}5"] = y + ("A" if y in HIST else "E"); P[f"{c}5"].font = BOLD
P.cell(5, 1, "項目")
STR = {nm: 6 + i for i, nm in enumerate(STD)}
for nm, rw in STR.items(): P.cell(rw, 1, nm)
for y, c in YC.items():
    p = chr(ord(c) - 1)
    ref = lambda k: f"{c}{b[k]}"
    P[f"{c}{STR['營收']}"] = f"={ref('rev')}"
    P[f"{c}{STR['營收年增率']}"] = f"={c}{STR['營收']}/{p}{STR['營收']}-1" if y != "FY2024" else f"=Q!B{QR['rev']}/2065.659-1"
    P[f"{c}{STR['營業成本']}"] = f"={ref('cogs')}"
    P[f"{c}{STR['毛利']}"] = f"={ref('gp')}"
    P[f"{c}{STR['毛利率']}"] = f"={c}{STR['毛利']}/{c}{STR['營收']}"
    P[f"{c}{STR['營業費用']}"] = f"={ref('opex')}"
    P[f"{c}{STR['營業利益']}"] = f"={ref('oi')}"
    P[f"{c}{STR['營業利益率']}"] = f"={c}{STR['營業利益']}/{c}{STR['營收']}"
    P[f"{c}{STR['其他收入']}"] = f"={ref('oth')}"
    P[f"{c}{STR['稅前利益']}"] = f"={ref('ebt')}"
    P[f"{c}{STR['所得稅']}"] = f"={ref('taxx')}"
    P[f"{c}{STR['稅後淨利']}"] = f"={ref('ni')}"
    P[f"{c}{STR['淨利率']}"] = f"={c}{STR['稅後淨利']}/{c}{STR['營收']}"
    P[f"{c}{STR['EPS（非 GAAP 稀釋，美元）']}"] = f"={ref('eps')}"
    P[f"{c}{STR['淨利年增率']}"] = f"={c}{STR['稅後淨利']}/{p}{STR['稅後淨利']}-1" if y != "FY2024" else "未找到"
P[f"A{STR['淨利年增率']+1}"] = "非 GAAP 稅前、稅額與其他收入的歷史值為推算（公司未揭露；見 Q 頁）。FY2024 營收年增率分母為 FY2023 營收 2,065.659（10-K FY2024 損益表）。GAAP 參考列見下。"
# GAAP 參考列
GR0 = STR["淨利年增率"] + 3
P.cell(GR0 - 1, 1, "參考：GAAP 與股票薪酬（GAAP 預測為近似，不進估值）").font = BOLD
P.cell(GR0, 1, "GAAP 營業利益（損失）"); P.cell(GR0 + 1, 1, "股票薪酬"); P.cell(GR0 + 2, 1, "GAAP 營業利益率"); P.cell(GR0 + 3, 1, "公司口徑自由現金流（FY2027E 起為模型值）"); P.cell(GR0 + 4, 1, "公司口徑自由現金流率")
for y, c in YC.items():
    P[f"{c}{GR0}"] = f"={c}{b['gaap_oi']}"; P[f"{c}{GR0+1}"] = f"={c}{b['sbc']}"; P[f"{c}{GR0+2}"] = f"={c}{GR0}/{c}{STR['營收']}"
    P[f"{c}{GR0+3}"] = f"={QH('fcf', y)}" if y in HIST else f"={c}{b['fcf_co']}"
    P[f"{c}{GR0+4}"] = f"={c}{GR0+3}/{c}{STR['營收']}"
P.cell(GR0 - 2, 1, None)
SEGROW0 = GR0 + 7
P.cell(SEGROW0 - 1, 1, "分部營收（情境 B；單一營運分部下的固定揭露營收線與命題子列）").font = BOLD
SEGP = {"營收合計（單一營運分部）": "rev", "Product revenue": "prod", "Professional services and other revenue": "ps", "核心平台消耗（命題子列）": "core", "AI 工作負載營收（命題子列）": "air"}
SPR = {}
for i, (lab, k) in enumerate(SEGP.items()):
    rw = SEGROW0 + i; SPR[lab] = rw
    P.cell(rw, 1, lab)
    for y, c in YC.items():
        if y in HIST and k in ("core", "air"): P[f"{c}{rw}"] = "未找到"
        elif y in HIST and k == "rev": P[f"{c}{rw}"] = f"={QH('rev', y)}"
        else: P[f"{c}{rw}"] = f"=Seg!{c}{SR['B'][k]}"
CK = SEGROW0 + len(SEGP) + 2
P.cell(CK - 1, 1, "對帳（應為 0）").font = BOLD
P.cell(CK, 1, "檢查：分部合計 − 合併營收")
P.cell(CK + 1, 1, "檢查：核心平台消耗＋AI 營收增量 − Product revenue（三情境）")
P.cell(CK + 2, 1, "檢查：分部合計 − 公司揭露合併營收（歷史年度；實績期）")
P.cell(CK + 3, 1, "檢查：營收 − 營業成本 − 毛利（FY2024A–FY2032E）")
for y, c in YC.items():
    P[f"{c}{CK}"] = "=MAX(" + ",".join(f"ABS(Seg!{c}{SR[s]['prod']}+Seg!{c}{SR[s]['ps']}-Seg!{c}{SR[s]['rev']})" for s in "ABC") + ")"
    if y in HIST:
        P[f"{c}{CK+1}"] = 0
        P[f"{c}{CK+2}"] = f"=ABS({c}{SPR['Product revenue']}+{c}{SPR['Professional services and other revenue']}-{QH('rev', y)})"
    else:
        P[f"{c}{CK+1}"] = "=MAX(" + ",".join(f"ABS(Seg!{c}{SR[s]['chk1']})" for s in "ABC") + ")"
        P[f"{c}{CK+2}"] = 0
    P[f"{c}{CK+3}"] = f"=ABS({c}{STR['營收']}-{c}{STR['營業成本']}-{c}{STR['毛利']})"

# ---------------------------------------------------------------- DCF 頁
Dp = SH["DCF"]
Dp["A1"] = "DCF（FCFF 標準規格；tools/fcff_dcf.py）：明確期 FY2027E 剩餘期間～FY2032E；永續價值＝NOPAT×(1+g)×(1−g÷RONIC)÷(WACC−g)；期中折現；評價日＝股價基準日。三情境各算一次。EBIT 已扣股票薪酬（SBC 不加回）。"
Dp["A2"] = "non-GAAP→FCFF 橋接：非 GAAP 營業利益−SBC＝EBIT；EBIT 稅後＋折舊攤銷−資本支出−營運資金增加＝FCFF（遞延收入增加為現金流入，fc_wc 為負值）；公司口徑自由現金流（PL 頁）＝FCFF＋SBC＋稅後其他收入。折舊與資本支出取 PL 頁（FY2027–FY2029 為近期比例，FY2030–FY2032 線性收斂到穩態比例 fc_da_ss／fc_cx_ss）。"
DYR = [2027, 2028, 2029, 2030, 2031, 2032]
DPATHS = []
for s in "ABC":
    pr = PR[s]
    DPATHS.append(dict(name=s, label=f"情境 {s}", rev={y: f"PL!{YC['FY'+str(y)]}{pr['rev']}" for y in DYR}, rev0=f"PL!E{pr['rev']}",
                       ebit={y: f"PL!{YC['FY'+str(y)]}{pr['ebit']}" for y in DYR}, tax={y: "tax_w" for y in DYR},
                       da={y: f"PL!{YC['FY'+str(y)]}{pr['da']}" for y in DYR}, cx={y: f"PL!{YC['FY'+str(y)]}{pr['capex']}" for y in DYR}))
FRAC0 = 79 / 365   # 2026-11-13 → 2027-01-31（FY2027 年底）
FDR = FD.write_detail(Dp, 4, DPATHS, DYR, FRAC0)
DCOL = FDR["col"]
for s in "ABC":
    pr = PR[s]; D = FDR["paths"][s]
    for y in ("FY2027", "FY2028", "FY2029", "FY2030", "FY2031", "FY2032"):
        c = YC[y]; dc = DCOL[int(y[2:])]
        P[f"{c}{pr['nopat']}"] = f"=DCF!{dc}{D['nopat']}"; P[f"{c}{pr['dwc']}"] = f"=DCF!{dc}{D['dwc']}"; P[f"{c}{pr['fcff']}"] = f"=DCF!{dc}{D['fcff']}"

# ---------------------------------------------------------------- Val 頁
V = SH["Val"]
VR = {}
vrows = ["intro", "hint", "blank0", "cap", "px", "px_date", "tp_date", "dcf_tp", "comp_tp", "tp", "up", "val_rec", "gate", "gate_use", "gate_res", "rec", "tp_ref", "sp1", "wh", "ke", "kd", "kdat", "wacc", "ronic", "g", "emkt", "dmkt", "ewt",
         "sp2", "ndh", "cashx", "ndebt", "netcash", "sh", "dps", "sp3", "sumh", "pvx", "tvpv", "ev", "eq", "vps0", "dcf12", "tvsh", "sp4", "ch", "ntm", "mult", "eqc", "comp12", "sp5", "scenh", "wa", "wb", "wc", "chk_w", "dec_h", "dec_gap", "dec_sbc", "dec_res", "rerate",
         "sp6", "peerh", "pe1", "pe2", "pe3", "pe4", "pe5", "pe6", "pe7", "pe8", "pe_med", "pe_n", "pe_own", "pe_pe", "rmx"]
for i, k in enumerate(vrows): VR[k] = 1 + i
V.cell(VR["intro"], 1, "估值：DCF（FCFF；EBIT 已扣股票薪酬）與 Comparable（同業前瞻 P/S 中位數×目標時點 NTM 營收）加權；目標價為 12 個月後（評價日＋12 個月）。三情境 25／50／25。上檔空間不是年化報酬。")
V.cell(VR["hint"], 1, "Comparable 的同業前瞻 P/S 取自 StockAnalysis statistics 頁（前瞻期間以市值÷前瞻 P/S 反推，與 NTM 不完全相同）；同業日期不一致（部分為 2026-11-06，逾期）。P/E 只作對照（SNOW 的非 GAAP EPS 偏低，P/E 不適合作主倍數）。相對倍數調整係數為 1.00（未調整：缺歷史前瞻倍數）。")
V.cell(VR["cap"], 1, "結論（情境加權）").font = BOLD
def vrow(k, label, b=None, c=None, d=None, e=None):
    r_ = VR[k]; V.cell(r_, 1, label)
    for col, v in zip("BCDE", (b, c, d, e)):
        if v is not None: V[f"{col}{r_}"] = v
vrow("px", "現價（US$）", f"={I('px')}")
vrow("px_date", "股價基準日", f"={I('px_date')}"); V[f"B{VR['px_date']}"].number_format = "yyyy-mm-dd"
vrow("tp_date", "目標時點（股價基準日＋12 個月）", f"=EDATE({I('px_date')},12)"); V[f"B{VR['tp_date']}"].number_format = "yyyy-mm-dd"
wr = lambda s: I("tp_w" + s)
vrow("dcf_tp", "DCF 目標價（情境加權；US$）", f"={wr('A')}*C{VR['dcf12']}+{wr('B')}*D{VR['dcf12']}+{wr('C')}*E{VR['dcf12']}")
vrow("comp_tp", "Comparable 目標價（情境加權；US$）", f"={wr('A')}*C{VR['comp12']}+{wr('B')}*D{VR['comp12']}+{wr('C')}*E{VR['comp12']}")
vrow("tp", "加權目標價（DCF 與 Comparable 加權；US$）", f"={I('w_dcf')}*B{VR['dcf_tp']}+(1-{I('w_dcf')})*B{VR['comp_tp']}")
vrow("up", "上檔空間（相對現價）", f"=B{VR['tp']}/B{VR['px']}-1")
vrow("val_rec", "估值評等（只看上檔空間）", f'=IF(B{VR["up"]}>0.5,"強力買進",IF(B{VR["up"]}>0.25,"買進",IF(B{VR["up"]}>=-0.25,"中立",IF(B{VR["up"]}>=-0.5,"放空","強力放空"))))')
gsum = lambda s_: f"Q!$D${QR['fcf']}+SUM(PL!F{PR[s_]['fcf_co']}:I{PR[s_]['fcf_co']})"
vrow("gate", "收支關卡（FY2026–FY2030 公司口徑自由現金流合計，未扣股利因無股利；情境 A／B／C；US$M）", None, f"={gsum('A')}", f"={gsum('B')}", f"={gsum('C')}")
vrow("gate_use", "入不敷出判定用（情境 B 期間合計；US$M）", f"=D{VR['gate']}")
vrow("gate_res", "入不敷出判定", f'=IF(B{VR["gate_use"]}<0,"不通過（入不敷出）","通過")')
vrow("rec", "投資建議", f'=IF(AND(B{VR["gate_use"]}<0,OR(B{VR["val_rec"]}="買進",B{VR["val_rec"]}="強力買進")),"收支不足",B{VR["val_rec"]})')
vrow("tp_ref", "只看 Comparable 的目標價（上檔空間見右）", f"=B{VR['comp_tp']}", f"=B{VR['comp_tp']}/B{VR['px']}-1")
V.cell(VR["wh"], 1, "折現率（DCF）").font = BOLD
vrow("ke", "股權成本 ke（rf＋β×(ERP＋CRP)；β＝Blume 調整；DCF 摘要區，見下）", "=ke")
vrow("kd", "稅前債務成本（rf＋利差）", "=kd")
vrow("kdat", "稅後債務成本", "=kd_at")
vrow("emkt", "股權市值（US$M；現價×評價股數）", f"={I('px')}*sh_v")
vrow("dmkt", "債務本金＋租賃（US$M）", "=debt_v+lease_v")
vrow("ewt", "股權權重", "=w_e")
vrow("wacc", "WACC", "=wacc")
vrow("ronic", "RONIC（＝WACC＋利差）", "=ronic")
vrow("g", "永續成長率", "=g_tv")
V.cell(VR["ndh"], 1, "淨現金（2026-07-31；US$M）").font = BOLD
vrow("cashx", "現金、約當現金與投資", "=cash_v")
vrow("ndebt", "債務（可轉債未轉換時）＋租賃負債", "=debt_v+lease_v")
vrow("netcash", "淨現金（負值＝淨債務）", "=dcf_nc")
vrow("sh", "評價股數（M；流通股＋RSU＋選擇權庫藏股法＋可轉債 if-converted 淨增）", "=sh_v")
vrow("dps", "期間股利（US$/股）", "=dps_tgt")
V.cell(VR["sumh"], 1, "DCF 彙總（三情境）").font = BOLD
for col, lab in zip("CDE", ("情境 A", "情境 B", "情境 C")): V[f"{col}{VR['sumh']}"] = lab
vrow("pvx", "明確期 FCFF 現值（US$M）"); vrow("tvpv", "永續價值現值（US$M）"); vrow("ev", "企業價值 EV（US$M）"); vrow("eq", "股權價值（EV＋淨現金；US$M）")
vrow("vps0", "每股價值（評價日；US$）"); vrow("dcf12", "DCF 情境目標價（12 個月後＝每股價值×(1+ke)−股利；US$）"); vrow("tvsh", "永續價值占 EV 比例")
V.cell(VR["ch"], 1, "Comparable（目標時點 NTM 營收 × 同業前瞻 P/S 中位數）").font = BOLD
vrow("ntm", "目標時點 NTM 營收（FY2028 與 FY2029 按剩餘月數加權；US$M）"); vrow("mult", "套用前瞻 P/S（同業中位數×調整係數）"); vrow("eqc", "Comparable 股權價值（US$M）"); vrow("comp12", "Comparable 情境目標價（US$；P/S 已是股權口徑，不加淨現金）")
V.cell(VR["scenh"], 1, "情境權重與加權").font = BOLD
vrow("wa", "情境權重 A／B／C", None, f"={wr('A')}", f"={wr('B')}", f"={wr('C')}")
vrow("wb", "情境加權目標價（各情境 DCF 與 Comparable 加權；US$）", None, *(f"={I('w_dcf')}*{c}{VR['dcf12']}+(1-{I('w_dcf')})*{c}{VR['comp12']}" for c in "CDE"))
vrow("wc", "各情境上檔空間", None, *(f"={c}{VR['wb']}/$B${VR['px']}-1" for c in "CDE"))
vrow("chk_w", "檢查：情境加權目標價與加權目標價之差（應為 0）", f"=ABS({I('tp_wA')}*C{VR['wb']}+{I('tp_wB')}*D{VR['wb']}+{I('tp_wC')}*E{VR['wb']}-B{VR['tp']})")
V.cell(VR["dec_h"], 1, "DCF 與 Comparable 落差分解（情境 B；每股 US$；目標時點）").font = BOLD
sb = PR["B"]
_DB = FDR["paths"]["B"]
_pv_sbc = "+".join(f"PL!{YC['FY'+str(y)]}{sb['sbc']}*DCF!{DCOL[y]}${FDR['frac']}*DCF!{DCOL[y]}{_DB['df']}" for y in DYR)
_tv_sbc = f"PL!K{sb['sbc']}*(1+B{VR['g']})*(1-B{VR['g']}/B{VR['ronic']})/(B{VR['wacc']}-B{VR['g']})/(1+B{VR['wacc']})^{FDR['tvt']}"
vrow("dec_gap", "Comparable 減 DCF（情境 B）", f"=D{VR['comp12']}-D{VR['dcf12']}")
vrow("dec_sbc", "其中：股票薪酬（DCF 扣除、P/S 未扣；稅後現值加永續價值，推到目標時點）", f"=(1-{I('tax')})*(({_pv_sbc})+{_tv_sbc})/B{VR['sh']}*(1+B{VR['ke']})")
vrow("dec_res", "其餘：槓桿（淨債務）、時點、前瞻期間與倍數水準差（殘差）", f"=B{VR['dec_gap']}-B{VR['dec_sbc']}")
vrow("rerate", "SNOW 前瞻 P/S 調整到同業中位數所隱含的價格（現價×同業中位數÷SNOW 自身；US$）", f"=B{VR['px']}*B{VR['pe_med']}/B{VR['pe_own']}")
V.cell(VR["peerh"], 1, "同業前瞻倍數（StockAnalysis statistics；八家；欄：前瞻 P/S、前瞻 P/E、市值 US$B、日期）").font = BOLD
PEERS = [("Datadog（DDOG）", 12.73, 61.83, 66.49, "2026-11-13"), ("MongoDB（MDB）", 7.67, 45.04, 25.20, "2026-11-06（逾期）"), ("Cloudflare（NET）", 29.98, 174.39, 109.60, "2026-11-13"),
         ("CrowdStrike（CRWD）", 22.82, 115.53, 122.72, "頁面未載日期"), ("ServiceNow（NOW）", 5.75, 22.46, 105.84, "2026-11-06（逾期）"), ("Palantir（PLTR）", 33.91, 73.10, 404.77, "2026-11-13"),
         ("Elastic（ESTC）", 4.36, 24.48, 9.04, "2026-11-06（逾期）"), ("Zscaler（ZS）", 7.98, 39.22, 31.38, "2026-11-13")]
for i, (nm, ps_, pe_, mc, dt) in enumerate(PEERS):
    rw = VR[f"pe{i+1}"]
    V.cell(rw, 1, f"{nm}｜{dt}")
    V.cell(rw, 2, ps_).font = BLUE; V.cell(rw, 3, pe_).font = BLUE; V.cell(rw, 4, mc).font = BLUE
    V.cell(rw, 5, "Verified（第二來源）"); V.cell(rw, 6, "US GAAP；P/S＝市值÷營收，租賃處理不影響分子分母；分母期間：前瞻（見 Recon）")
vrow("pe_med", "同業中位數前瞻 P/S", f"=MEDIAN(B{VR['pe1']}:B{VR['pe8']})", f"=MEDIAN(C{VR['pe1']}:C{VR['pe8']})")
V.cell(VR["pe_med"], 4, "（C 欄＝同業中位數前瞻 P/E，只作對照）")
vrow("pe_n", "家數", f"=COUNT(B{VR['pe1']}:B{VR['pe8']})")
vrow("pe_own", "Snowflake 自身前瞻 P/S（2026-11-13；僅對照）", 14.31, 115.37, 102.56)
V.cell(VR["pe_own"], 5, "C＝前瞻 P/E；D＝市值 US$B；隱含前瞻營收（US$M）見 Recon")
for col in "BCD": V[f"{col}{VR['pe_own']}"].font = BLUE
vrow("pe_pe", "Snowflake 前瞻 P/S 相對同業中位數", f"=B{VR['pe_own']}/B{VR['pe_med']}-1")
# rel_mult 表與 FD 摘要區在 DCF 檢查之後（見下）
# ---- rel_mult：最近 5 個期間 SNOW 倍數 ÷ 同業中位數（原則 G3 第 3 條；Derived）
RM0 = VR["rmx"] + 1
V.cell(RM0, 1, "相對倍數 rel_mult：SNOW 最近 5 個期間的 trailing P/S ÷ 同業（八家）trailing P/S 中位數；取 5 期比值中位數，夾在 0.5–2.0（Derived）").font = BOLD
RMH = ["期間（StockAnalysis 財務比率表，市值÷近四季營收）", "SNOW", "DDOG", "MDB", "NET", "CRWD", "NOW", "PLTR", "ESTC", "ZS", "同業中位數", "比值"]
for j, h in enumerate(RMH, 1): V.cell(RM0 + 1, j, h).font = BOLD
RMD = [("FY2026（2026-01-31）", 14.08, 13.92, 12.27, 31.85, 23.13, 11.97, 94.66, 2.76, 7.29),
       ("FY2025（2025-01-31）", 16.52, 18.08, 10.14, 22.13, 24.80, 19.91, 60.12, 6.07, 16.63),
       ("FY2024（2024-01-31）", 22.96, 18.74, 17.18, 21.56, 22.09, 16.11, 16.79, 8.13, 12.51),
       ("FY2023（2023-01-31）", 24.36, 13.93, 11.56, 15.23, 10.50, 10.81, 7.01, 5.17, 14.47),
       ("FY2022（2022-01-31）", 69.31, 54.02, 30.95, 64.45, 26.57, 21.91, 23.68, 8.25, 20.16)]
for i, row in enumerate(RMD):
    rw = RM0 + 2 + i
    V.cell(rw, 1, row[0])
    for j, v in enumerate(row[1:], 2): V.cell(rw, j, v).font = BLUE
    V.cell(rw, 11, f"=MEDIAN(C{rw}:J{rw})"); V.cell(rw, 12, f"=B{rw}/K{rw}")
RMR = RM0 + 2 + len(RMD)
V.cell(RMR, 1, "rel_mult（5 期比值中位數，夾在 0.5–2.0）"); V.cell(RMR, 2, f"=MIN(2,MAX(0.5,MEDIAN(L{RM0+2}:L{RMR-1})))"); V.cell(RMR, 3, "Derived；來源與缺漏見 Recon；trailing P/S 與前瞻 P/S 口徑不同，為主要載重輸入")
RELM_ROW = RMR
# ---- FD 摘要區（折現率、DCF 彙總）
FS0 = RMR + 3
FS = FD.write_summary(wb, V, FS0, FDR, ["A", "B", "C"], ["tp_wA", "tp_wB", "tp_wC"], px="px", beta_note="Blume 調整；單一來源")
fd_ = lambda k, col: f"=DCF!$C${FDR['paths'][col][k]}"
for si, s_ in enumerate("ABC"):
    col = "CDE"[si]; D_ = FDR["paths"][s_]
    V[f"{col}{VR['pvx']}"] = f"=DCF!$C${D_['pvx']}"; V[f"{col}{VR['tvpv']}"] = f"=DCF!$C${D_['pvtv']}"; V[f"{col}{VR['ev']}"] = f"=DCF!$C${D_['ev']}"
    V[f"{col}{VR['eq']}"] = f"={col}{VR['ev']}+$B${VR['netcash']}"; V[f"{col}{VR['vps0']}"] = f"=DCF!$C${D_['val']}"
    V[f"{col}{VR['dcf12']}"] = f"=DCF!$C${D_['v12']}"; V[f"{col}{VR['tvsh']}"] = f"=DCF!$C${D_['tvsh']}"
    V[f"{col}{VR['ntm']}"] = f"=((DATE(2028,1,31)-$B${VR['tp_date']})/365)*PL!G{PR[s_]['rev']}+(1-(DATE(2028,1,31)-$B${VR['tp_date']})/365)*PL!H{PR[s_]['rev']}"
    V[f"{col}{VR['mult']}"] = f"=$B${VR['pe_med']}*{I('rel_mult')}"
    V[f"{col}{VR['eqc']}"] = f"={col}{VR['ntm']}*{col}{VR['mult']}"
    V[f"{col}{VR['comp12']}"] = f"={col}{VR['eqc']}/$B${VR['sh']}"
wb.defined_names["px"] = DefinedName("px", attr_text=f"In!$E${IR['px']}")

for _row in SH['In'].iter_rows():
    for _c in _row:
        if isinstance(_c.value, str) and "@@" in _c.value:
            _c.value = (_c.value.replace("@@RELM@@", str(RELM_ROW)).replace("@@OPXA@@", str(RS["opxa"])).replace("@@OPXV@@", str(RS["opxv"]))
                        .replace("@@SBCA@@", str(RS["sbca"])).replace("@@SBCV@@", str(RS["sbcv"])))

# ---------------------------------------------------------------- RevDCF 頁
RD = SH["RevDCF"]
RV = {}
RD["A1"] = "反向 DCF：現價隱含什麼。以情境 B 路徑為底，兩個封閉解（DCF 對 EBIT 率與營收規模都是線性）；隱含 WACC 需整本重算，本版未做（Sens 頁有 WACC 相關案例）。"
RD["A2"] = "讀法：隱含值是「DCF 每股價值（評價日）＝現價」的點，不是評等門檻。EBIT 率是扣除股票薪酬後的口徑。"
rv_rows = [("lab", "項目"), ("v0", "情境 B 每股價值（評價日；US$）"), ("px", "現價（US$）"), ("gap", "現價與 DCF 每股價值之差（股權價值；US$M）"),
           ("k_m", "EBIT 率全期平移 1 個百分點對股權價值的影響（US$M）"), ("d_m", "現價隱含的 EBIT 率平移（百分點）"), ("m_fy32", "情境 B FY2032E EBIT 率（非 GAAP 營業利益率−SBC 占營收）"), ("m_imp", "現價隱含的 FY2032 EBIT 率"),
           ("ev0", "情境 B 企業價值（US$M）"), ("ev_imp", "現價隱含的企業價值（US$M）"), ("scale", "現價隱含的營收規模倍數（相對情境 B；FCFF 與營收同比例）"),
           ("mult_imp", "現價隱含的 EV／NTM 營收倍數（情境 B）"), ("mult_peer", "同業中位數前瞻 P/S"), ("ps", "分部加總：不適用（本模型為單一營運分部，Comparable 以前瞻 P/S）")]
for i, (k, lab) in enumerate(rv_rows):
    RV[k] = 4 + i; RD.cell(4 + i, 1, lab)
b_ = PR["B"]
RD[f"B{RV['v0']}"] = f"=Val!D{VR['vps0']}"; RD[f"B{RV['px']}"] = f"=Val!B{VR['px']}"
RD[f"B{RV['gap']}"] = f"=(B{RV['px']}-B{RV['v0']})*Val!B{VR['sh']}"
terms = "+".join(f"PL!{YC['FY'+str(y)]}{b_['rev']}*DCF!{DCOL[y]}${FDR['frac']}*DCF!{DCOL[y]}{FDR['paths']['B']['df']}" for y in DYR)
RD[f"B{RV['k_m']}"] = f"=0.01*(1-tax_w)*(({terms})+PL!K{b_['rev']}*(1+Val!B{VR['g']})*(1-Val!B{VR['g']}/Val!B{VR['ronic']})/(Val!B{VR['wacc']}-Val!B{VR['g']})/(1+Val!B{VR['wacc']})^{FDR['tvt']})"
RD[f"B{RV['d_m']}"] = f"=B{RV['gap']}/B{RV['k_m']}"
RD[f"B{RV['m_fy32']}"] = f"=PL!K{b_['ebit']}/PL!K{b_['rev']}"
RD[f"B{RV['m_imp']}"] = f"=B{RV['m_fy32']}+B{RV['d_m']}/100"
RD[f"B{RV['ev0']}"] = f"=Val!D{VR['ev']}"
RD[f"B{RV['ev_imp']}"] = f"=B{RV['px']}*Val!B{VR['sh']}-Val!B{VR['netcash']}"
RD[f"B{RV['scale']}"] = f"=B{RV['ev_imp']}/B{RV['ev0']}"
RD[f"B{RV['mult_imp']}"] = f"=B{RV['ev_imp']}/Val!D{VR['ntm']}"
RD[f"B{RV['mult_peer']}"] = f"=Val!B{VR['pe_med']}"
RD.cell(RV["ps"], 2, "—")

# ---------------------------------------------------------------- Recon 頁
Rc = SH["Recon"]
Rc["A1"] = "對帳：本模型 vs 公司指引與市場共識（共識與指引為 Interested-party；找不到者寫「未找到」）。口徑不同者不判定差異。"
RCR = {}
G27 = "公司 FY2027 指引（Q2 FY27 新聞稿，2026-09-02）"
rec_rows = [
 ("h1", "1. 合併層：FY2027（情境 B）", None),
 ("prod27", "產品營收（US$M）", (f"=PL!F{SPR['Product revenue']}", 6070, 6070, "未找到", G27 + "：產品營收 6,070（含 Observe 約 1 個百分點）；本模型 Q3、Q4 為假設（Assumed），高於指引屬預期（公司指引通常保守，Interested-party）；共識的產品營收未找到")),
 ("rev27", "合併營收（US$M）", (f"=PL!F{STR['營收']}", None, None, 6320, "賣方共識 6,320（StockAnalysis／S&P，2026-11-12，47 家；Interested-party）")),
 ("om27", "非 GAAP 營業利益率", (f"=PL!F{STR['營業利益率']}", 0.145, 0.145, "未找到", G27 + "：14.5%；本模型由營業費用率假設校準，非獨立預測")),
 ("gm27", "非 GAAP 產品毛利率", (f"=Seg!F{SR['B']['pgm']}", 0.74, 0.74, "未找到", G27 + "：74.0%；本模型為 U_gm0 與 AI 推論成本的結果（含 AI 稀釋）")),
 ("fcf27", "公司口徑自由現金流率（調整後）", (f"=PL!F{GR0+4}", 0.23, 0.23, "未找到", G27 + "：調整後自由現金流率 23%；模型口徑為 FCFF＋SBC＋稅後其他收入，不含薪資稅調整，非同一口徑")),
 ("sh27", "非 GAAP 稀釋加權股數（M）", (f"={I('sh_ng')}", 380, 380, "未找到", G27 + "：380M（含可轉債 if-converted 約 10M）；本模型直接採用指引（Interested-party）")),
 ("h2", "2. 共識（EPS 為非 GAAP；口徑相同）", None),
 ("eps27", "FY2027 非 GAAP EPS（US$）", (f"=PL!F{STR['EPS（非 GAAP 稀釋，美元）']}", None, None, 2.23, "賣方平均 2.23（區間 2.07–2.56；Interested-party）")),
 ("rev28", "FY2028 營收（US$M）", (f"=PL!G{STR['營收']}", None, None, 8160, "賣方平均 8,160（頁面摘要欄；Interested-party）")),
 ("eps28", "FY2028 非 GAAP EPS（US$）", (f"=PL!G{STR['EPS（非 GAAP 稀釋，美元）']}", None, None, 3.06, "賣方平均 3.06（Interested-party）")),
 ("tp", "賣方平均目標價與本模型加權目標價（US$）", (f"=Val!B{VR['tp']}", None, None, 434.09, "StockAnalysis 51 家平均 434.09、中位 447、區間 280–540（2026-11-12）；MarketBeat 41 家 414.24（2026-11-16 讀取）；本模型加權目標價見左（Interested-party 共識）")),
 ("fwdps", "SNOW 前瞻 P/S 隱含的前瞻營收（US$M；市值÷前瞻 P/S）", (f"=Val!D{VR['pe_own']}*1000/Val!B{VR['pe_own']}", None, None, "未找到", "約 7.2B，介於 FY2027 與 FY2028 共識營收之間，故同業前瞻 P/S 的期間約為 NTM 附近；各同業的前瞻期間未逐一反推")),
]
r = 3
for k, lab, spec in rec_rows:
    RCR[k] = r
    Rc.cell(r, 1, lab)
    if spec is None:
        Rc.cell(r, 1).font = BOLD
        for j, h in enumerate(["本模型", "指引低", "指引高", "共識", "說明"], 2): Rc.cell(r, j, h).font = BOLD
    else:
        for j, v in enumerate(spec, 2):
            if v is not None: Rc.cell(r, j, v)
    r += 1
KREQ = os.path.join(HERE, "data/kreq_v01.json")
if os.path.exists(KREQ):
    for n_, kq in enumerate(json.load(open(KREQ, encoding="utf-8"))):
        RCR["kreq" if n_ == 0 else "kreq2"] = r
        Rc.cell(r, 1, "反解（整本重算）：" + kq["label"] + " 為多少時，加權目標價＝現價")
        Rc.cell(r, 2, kq["base"]); Rc.cell(r, 5, kq["req"]); Rc.cell(r, 6, kq["note"])
        r += 1
Rc.cell(r + 1, 1, "3. 同業日期與口徑").font = BOLD
Rc.cell(r + 2, 1, "同業前瞻 P/S 取自 StockAnalysis；MongoDB、ServiceNow、Elastic 為 2026-11-06（距股價日逾兩個交易日，標逾期），CrowdStrike 頁面未載日期；八家家數達規格門檻但前瞻期間不一，不做迴歸；以調整係數作敏感度。")
Rc.cell(r + 3, 1, "4. Q3 FY27 實績（截至 2026-10-31）尚未公布，發布日未找到；本版以 Q2 FY27 為最新已公布季，Q3 指引（產品營收 1,588–1,593、非 GAAP 營業利益率 15.5%）只作對照。")
Rc.cell(r + 4, 1, "5. 未找到：AI 營收占比（公司只稱 meaningful step-up）、AI 專屬毛利率、Q2 FY27 租賃負債與流通股封面數、第二來源 Beta、FY2029 以後公司長期目標、共識的產品營收。")

# ---------------------------------------------------------------- Sens 頁
Se = SH["Sens"]
Se["A1"] = "敏感度：In 頁每個有「敏感度低／高」區間的輸入，一次動一個、整本 LibreOffice 重算；表中為加權目標價相對基準。情境專屬列與情境權重不列（三情境本身就是情境分析）。"
Se["A2"] = "白話標籤欄為「驅動」。群組：Tokenomics 每 MW 持有成本（IF_HoldEcon）本模型不適用（軟體公司向模型商購買推論，不持有算力）。"
fp = hashlib.sha256(json.dumps([(k, str(wb["In"].cell(IR[k], 5).value)) for k in sorted(IR)], ensure_ascii=False).encode()).hexdigest()[:12]
BUILD_ID = f"snowflake-v{VER}-{fp}"
sens = json.load(open(SENS, encoding="utf-8")) if SENS else None
Se.cell(3, 1, "build_id"); Se.cell(3, 4, BUILD_ID)
Se.cell(4, 1, "模型狀態"); Se.cell(4, 4, "CURRENT" if (sens and sens.get("fingerprint") == fp) else "STALE")
Se.cell(5, 1, "基準加權目標價（本檔）"); Se.cell(5, 4, f"=Val!B{VR['tp']}")
Se.cell(6, 1, "敏感度重算時的基準加權目標價"); Se.cell(6, 4, sens["base_tp"] if sens else None)
Se.cell(7, 1, "敏感度重算時的輸入指紋"); Se.cell(7, 4, sens["fingerprint"] if sens else None)
hdr_r = 9
for j, h in enumerate(["驅動", "設定", "方向", "加權目標價", "相對基準", "鍵", "DCF 目標價", "DCF 變動（美元）", "投資建議"], 1):
    Se.cell(hdr_r, j, h).font = BOLD
SE_ROWS = []
if sens:
    rr = hdr_r + 1
    for case in sens["cases"]:
        Se.cell(rr, 1, case["label"]); Se.cell(rr, 2, case["setting"]); Se.cell(rr, 3, case["side"])
        Se.cell(rr, 4, case["tp"]); Se.cell(rr, 5, f"=D{rr}/$D$5-1"); Se.cell(rr, 6, case["key"]); Se.cell(rr, 7, case["dcf"]); Se.cell(rr, 8, f"=G{rr}-Val!$B${VR['dcf_tp']}"); Se.cell(rr, 9, case["rec"])
        SE_ROWS.append(rr); rr += 1
    Se.cell(rr + 1, 1, "註：tp_w（情境權重）與 scn_*（情境驅動）不列：三情境並列已涵蓋；IF_HoldEcon（Tokenomics 持有成本）不適用。")
SE_END = hdr_r + len(SE_ROWS)

# ---------------------------------------------------------------- Sources 頁
So = SH["Sources"]
So["A1"] = "來源與版本"
src_rows = [("Tokenomics 快照", f"{TK['source']['version']}（{TK['source']['file']}；commit {TK['source']['commit'][:7]}）；名稱：IF_PriceFresh／Cached／Out／TaskSucc／FullCost（Luna、Terra、Sol、Astra）、IF_TaskTokFresh／Cached／Dec（data/tokenomics_snapshot_v22.0.json）"),
            ("研究包", "software_ai/snowflake/data/research_snowflake.md、facts_snowflake.csv"),
            ("公司申報", "Snowflake 8-K 附件 99.1（Q2 FY27 2026-09-02；Q1 FY27 2026-05-27；Q4 FY26 2026-02-25；Q3 FY26 2025-12）；10-K FY2026；10-Q Q2 FY27"),
            ("同業倍數與股價", "StockAnalysis statistics 頁（SNOW、DDOG、MDB、NET、CRWD、NOW、PLTR、ESTC、ZS；擷取日 2026-11-16）；forecast 頁（共識，2026-11-12 更新）"),
            ("本版先驗", "未做（Andy 2026-11-16 23:27）")]
for i, (a, b2) in enumerate(src_rows, 3):
    So.cell(i, 1, a); So.cell(i, 2, b2)
So.cell(2, 1, "項目"); So.cell(2, 2, "內容")

# ---------------------------------------------------------------- README 頁
Rm = SH["README"]
Rm["A1"] = f"Snowflake 個股模型 v{VER}｜資料截止 2026-11-16（最新已公布季 Q2 FY27，截至 2026-07-31）｜成品日 {VDATE}｜股價基準日 2026-11-13｜build_id {BUILD_ID}｜本版未做先驗"
Rm["A2"] = "命題、結論與重點觀察的數字由公式自本檔其他頁組字（不手寫）。"
VB = lambda k, col="B": f"Val!{col}{VR[k]}"
SB = SR["B"]
Rm["A3"] = ('="命題：AI 進入按用量計價的資料平台後，產品營收＝核心平台消耗＋AI 工作負載；AI 淨效果＝AI 營收增量−核心被替代的侵蝕−推論成本。情境 B 在 FY2032 的 AI 淨效果為 "&TEXT(Seg!K' + str(SB["net"]) + '/1000,"0.0")&" 十億美元（AI 營收 "&TEXT(Seg!K' + str(SB["air"]) + '/1000,"0.0")&" − 核心侵蝕 "&TEXT(Seg!K' + str(SB["eros"]) + '/1000,"0.0")&" − 推論成本 "&TEXT(Seg!K' + str(SB["infc"]) + '/1000,"0.0")&"），AI 營收占產品營收 "&TEXT(Seg!K' + str(SB["aish"]) + ',"0%")&"。"')
Rm["A4"] = "命題的可證偽指標：扣除價格因素後的用量成長（工作負載乘數）、AI 功能消耗的 credit 占比、產品毛利率是否因 AI 轉售而下降；任何一項偏離基準區間即須重估。"
Rm["A5"] = (f'="投資建議（在 AI 工作負載與被替代落在情境 A 到 C 區間、推論成本取 Tokenomics 市場價的基準假設下）："&{VB("rec")}&"；加權目標價 "&TEXT({VB("tp")},"0.00")&" 美元，上檔空間 "&TEXT({VB("up")},"0.0%")&"（現價 "&TEXT({VB("px")},"0.00")&" 美元，"&TEXT({VB("px_date")},"yyyy-mm-dd")&"）；估值評等 "&{VB("val_rec")}&"；收支關卡（FY2026–FY2030 公司口徑自由現金流合計 "&TEXT({VB("gate_use")},"#,##0")&" 百萬美元）"&{VB("gate_res")}')
Rm["A6"] = (f'="一、DCF 目標價 "&TEXT({VB("dcf_tp")},"0.00")&"、Comparable 目標價 "&TEXT({VB("comp_tp")},"0.00")&" 美元；情境 B 兩者相差 "&TEXT({VB("dec_gap")},"0.00")&" 美元，其中股票薪酬約 "&TEXT({VB("dec_sbc")},"0.00")&"、其餘（槓桿、時點與倍數）約 "&TEXT({VB("dec_res")},"0.00")&" 美元；同業中位數前瞻 P/S "&TEXT({VB("pe_med")},"0.0")&" 倍，SNOW 自身 "&TEXT({VB("pe_own")},"0.0")&" 倍（"&IF({VB("pe_pe")}>0,"高於","低於")&"中位數 "&TEXT(ABS({VB("pe_pe")}),"0%")&"），調整到中位數隱含價格 "&TEXT({VB("rerate")},"0.00")&" 美元。"')
Rm["A7"] = (f'="二、情境 B FY2027 產品營收 "&TEXT(PL!F{SPR["Product revenue"]}/1000,"0.00")&" 十億美元，"&IF(PL!F{SPR["Product revenue"]}>Recon!C{RCR["prod27"]},"高於","低於")&"公司指引 "&TEXT(Recon!C{RCR["prod27"]}/1000,"0.00")&" 十億美元；非 GAAP 營業利益率 "&TEXT(PL!F{STR["營業利益率"]},"0.0%")&"，"&IF(PL!F{STR["營業利益率"]}>Recon!C{RCR["om27"]},"高於","低於")&"公司指引 "&TEXT(Recon!C{RCR["om27"]},"0.0%")&"（模型費用率與毛利率依規格持平於最近四季實績，不含公司預期的槓桿；落差即此處假設的影響）。"')
Rm["A8"] = (f'="三、情境 B FY2032 推論成本（市場價）占產品營業成本 "&TEXT(Seg!K{SB["infc"]}/(Seg!K{SB["infc"]}+Seg!K{SB["cogsx"]}),"0%")&"，產品毛利率由 FY2027 的 "&TEXT(Seg!F{SB["pgm"]},"0.0%")&" 變為 "&TEXT(Seg!K{SB["pgm"]},"0.0%")&"；以全成本計的推論成本僅 "&TEXT(Seg!K{SB["inf_full"]}/1000,"0.0")&" 十億美元（對照用）。"')
Rm["A9"] = (f'="四、現價隱含 FY2032 EBIT 率（扣股票薪酬）"&TEXT(RevDCF!B{RV["m_imp"]},"0.0%")&"（情境 B 為 "&TEXT(RevDCF!B{RV["m_fy32"]},"0.0%")&"）；永續價值占 EV 約 "&TEXT(Val!D{VR["tvsh"]},"0%")&"。"')
Rm["A10"] = (f'="五、情境 A／B／C 加權目標價 "&TEXT(Val!C{VR["wb"]},"0.00")&"／"&TEXT(Val!D{VR["wb"]},"0.00")&"／"&TEXT(Val!E{VR["wb"]},"0.00")&" 美元；分化來自 AI 工作負載、被替代、轉售加成與股票薪酬占比。"')
Rm["A12"] = "v1.3 展示層"
Rm["A13"] = (f'="Snowflake："&{VB("rec")}&"，加權目標價 "&TEXT({VB("tp")},"0.00")&" 美元（上檔 "&TEXT({VB("up")},"0.0%")&"）；本版未做先驗。"')
Rm["A14"] = "本版要點"
Rm["A15"] = "單一營運分部（10-K）；營收列取公司固定揭露的 Product revenue 與 Professional services and other revenue；AI 營收不單列（未找到），核心與 AI 為本模型拆分（Assumed）。"
Rm["A16"] = "損益表為非 GAAP 口徑（公司與賣方共識口徑）；DCF 的 EBIT 扣除股票薪酬，GAAP 營業利益與 EPS 為參考列。"
Rm["A17"] = "推論成本只讀 Tokenomics v22.0 快照，不自行估 token 價格；未來價格路徑待 AI供應鏈 登錄，以 Assumed 區間作敏感度。"
Rm["A18"] = "未做事項：先驗（Andy 指示本版跳過）。不適用：三層基準協調（Coord、C0、兩讀數）：非 Siliconomics 覆蓋；Meta 的算力、Tokenomics 自用 AI 讀法 2、Wattonomics 對照：Snowflake 不自建算力，推論成本為向模型商購買。"
Rm["A20"] = "市場資料日期"
Rm["A21"] = (f'="現價 "&TEXT({I("px_date")},"yyyy-mm-dd")&" 收盤；無風險利率 4.99% 為 2026-10-27 取值（距現價日 "&TEXT({I("px_date")}-DATE(2026,10,27),"0")&" 日，逾期）；同業前瞻 P/S 2026-11-06 至 2026-11-13（2026-11-06 者逾期，CrowdStrike 未載日期）；共識 2026-11-12。"')
Rm["A22"] = "主要失效條件（可觀察）：Q3 FY27 產品營收低於指引 1,588 以下、產品毛利率跌破指引 74%、RPO 年增低於 25%、淨營收留存率跌破 120%、AI 功能 credit 占比停滯。"
Rm["A23"] = "風險登錄"
Rm["A24"] = "風險"; Rm["B24"] = "是否建模"; Rm["C24"] = "理由"; Rm["D24"] = "觀察指標"
risks = [("核心消耗被模型直讀資料或開放格式替代", "是", "U_e32 三情境", "淨營收留存率、儲存與運算 credit 成長"), ("AI 推論成本侵蝕產品毛利", "是", "Tokenomics 市場價×呼叫數", "產品毛利率、第三方雲端費用占產品營業成本"),
         ("股票薪酬與稀釋", "是", "DCF 扣 SBC；sbc32 三情境", "SBC 占營收、稀釋股數"), ("同業倍數與大盤評價", "部分", "Comparable 係數 cm_adj", "同業前瞻 P/S"), ("併購（Observe、Natoma）與整合", "否", "營收錨點已含 Observe", "產品營收指引含併購貢獻")]
for i, rk in enumerate(risks, 25):
    for j, v in enumerate(rk, 1): Rm.cell(i, j, v)
Rm.cell(31, 1, "工作表：README、In、Q、Seg、PL、DCF、Val、RevDCF、Recon、Sens、Sources")
Rm["A33"] = "頁"; Rm["B33"] = "輸入"; Rm["C33"] = "推導量"; Rm["D33"] = "因果方向"
mp = [("Seg", "q3g、q4g、U_w26、U_w32_A／B／C、U_e26、U_e32_A／B／C、U_j26、U_j32、U_mk_A／B／C、g_px、g_tok、ovh、U_tier、g_b28、g_b32", "基準消耗、AI 營收、核心侵蝕、推論成本、AI 淨效果", "基準消耗×(1＋w−e)＋轉售營收 → 產品營收；呼叫數×每次成本 → 推論成本"),
      ("PL", "U_gm0、gm_dr、opx27、opx_dr、sbc27、sbc32_A／B／C、da_pct、capex_pct、wc_k、tax、yld", "毛利、非 GAAP 營業利益、EBIT、FCFF、EPS", "營收 − 成本 − 費用 → 非 GAAP 營業利益 → 扣 SBC → NOPAT → FCFF"),
      ("Val", "rf、beta_raw、erp、kd_sp、g_tv、ronic_sp、w_dcf、cm_adj、sh_dil", "WACC、DCF、Comparable、加權目標價、評等", "FCFF 折現＋同業前瞻 P/S × NTM 營收 → 加權"),
      ("Sens", "In 頁敏感度低／高", "加權目標價、DCF 目標價、投資建議", "一次改一個輸入、整本重算")]
for i, rk in enumerate(mp, 34):
    for j, v in enumerate(rk, 1): Rm.cell(i, j, v)
Rm["A39"] = "版本沿革：v0.1（2026-11-16）首版；未做先驗；本版與其他版本無對照。"

for ws_ in wb.worksheets:
    ws_.column_dimensions["A"].width = 52
    for col in "BCDEFGHIJK": ws_.column_dimensions[col].width = 14
wb["In"].column_dimensions["B"].width = 60; wb["In"].column_dimensions["H"].width = 70
wb.defined_names["PL_hdr"] = DefinedName("PL_hdr", attr_text="PL!$C$5:$K$5")
wb.save(OUT)
rowmap = {"IR": IR, "SR": SR, "PR": PR, "STR": STR, "SPR": SPR, "VR": VR, "DR": FDR, "FS": FS, "RM": {"rel_mult": RELM_ROW, "first": RM0 + 2}, "RV": RV, "RCR": RCR, "CK": CK, "QR": QR, "SE_ROWS": SE_ROWS, "SE_END": SE_END, "SCN_H": SCN_H, "SCN_ROWS": SCN_ROWS,
          "BUILD_ID": BUILD_ID, "FP": fp, "HDR_R": hdr_r, "GR0": GR0, "LAST_PARAM_ROW": LAST_PARAM_ROW}
if ROWMAP: json.dump(rowmap, open(ROWMAP, "w"), ensure_ascii=False, indent=1)
print("built", OUT, BUILD_ID)
