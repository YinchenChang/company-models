# 京元電子（2449）收支與估值模型 v0.2.1（AI 半導體 Project：測試代工；AI GPU／ASIC 最終測試與 burn-in）
# 用法：python3 kyec/build_kyec.py raw.xlsx ；再以 LibreOffice headless 重算
# 會計年度＝日曆年，FY(t)＝CY(t)，不平移。金額：十億新台幣（NT$B）；EPS：新台幣；顆數：百萬顆。口徑：歸屬母公司淨利。
# 樣板：advantest/（測試需求＝加速器顆數 × 每顆測試時間；顆數與情境係數讀 tsmc/dist）；台股寫法（月營收、法說、新台幣）比照 delta/、emc/。
# 與 Advantest 不同：京元電賣的是測試時間（服務），不是測試機；營收＝加權顆數 × 每加權顆營收（FY2026 校準）× 份額相對值 × 價格指數，
# 毛利率由「固定現金成本＋變動現金成本＋折舊」推出，折舊由資本支出（測試機台、廠房）累積，資本支出隨當年營收增量擴產。
import sys, os, glob
import openpyxl
from openpyxl.styles import Font, PatternFill
from openpyxl.workbook.defined_name import DefinedName
ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, os.path.join(ROOT, "tools"))
sys.path.insert(0, os.path.join(ROOT, "siliconomics"))
import reader as SIL   # Siliconomics 快照（行業變數單一來源）
import cycle_position as CP
import valuation_rule as VR

BLUE = Font(color="0000FF"); BOLD = Font(bold=True); HDR = PatternFill("solid", fgColor="DDEBF7")
KEY = PatternFill("solid", fgColor="FFF2CC")
wb = openpyxl.Workbook()
def hdr(ws, row, labels):
    for j, h in enumerate(labels, 1):
        c = ws.cell(row, j, h); c.font = BOLD; c.fill = HDR
SC = ("A", "B", "C")

# ---------------- 讀入其他模型的成品（唯讀；不修改它們） ----------------
def only(pattern):
    fs = glob.glob(os.path.join(ROOT, pattern))
    assert len(fs) == 1, (pattern, fs)
    return fs[0]
TSMC_F = only("tsmc/dist/*.xlsx")
rel = lambda f: os.path.relpath(f, ROOT)
TS = {}; TSN = {}
for r in openpyxl.load_workbook(TSMC_F, data_only=True)["In"].iter_rows(min_row=4):
    k = r[0].value
    if k and (str(k).startswith("u_") or str(k).startswith("g29_") or str(k).startswith("um2")):
        TS[k] = (r[4].value, f"{rel(TSMC_F)}｜In!E{r[0].row}"); TSN[k] = r[1].value
PLAT = [("bw", "NVIDIA Blackwell"), ("rb", "NVIDIA Rubin"), ("ru", "NVIDIA Rubin Ultra"), ("hp", "NVIDIA Hopper"),
        ("amd", "AMD Instinct"), ("tpu", "Google TPU"), ("trn", "AWS Trainium"), ("oth", "其他 ASIC（MTIA、Maia 等）")]
NV = ("bw", "rb", "ru", "hp"); AX = ("amd", "tpu", "trn", "oth")
for k, _ in PLAT:
    for y in (26, 27, 28):
        assert f"u_{k}_{y}" in TS and TS[f"u_{k}_{y}"][0] is not None, (k, y)
    assert f"g29_{k}" in TS
assert all(f"um{y}_{s}" in TS for y in (28, 29) for s in SC), TS.keys()

# ---------------- In ----------------
IN = wb.active; IN.title = "In"
IN["A1"] = "In — 輸入（藍字；基準欄 E，低 D、高 F）。金額：十億新台幣（NT$B）；顆數：百萬顆；比例：%。每個輸入附兩個來源（找不到者寫「未找到」）；公司原文算一個"
IN["A2"] = "標記：Verified／Interested-party／Analogy／Assumed／Derived；Analogy、Assumed 給低／高區間。SemiAnalysis 未使用。公司原文＝京元電每季 Investor update（英文簡報）與月營收公告；AI 占比、客戶與券商估計為 Interested-party（媒體、券商轉述）"
hdr(IN, 3, ["鍵", "項目", "單位", "低", "基準", "高", "標記", "來源1", "來源2／說明"])
rows = []; sec = lambda t: rows.append(("sec", t)); inp = lambda *a: rows.append(a)
DECK2Q26 = "https://www.kyec.com.tw/Upfiles/EDUp/files/投資人關係/每季營運報告/KYEC(2449)_2Q26 Investor update_English.pdf"
DECK4Q25 = "https://www.kyec.com.tw/Upfiles/EDUp/files/投資人關係/每季營運報告/KYEC(2449)_4Q25 Investor update_English.pdf"
DECK2Q25 = "https://kyec.com.tw/Upfiles/EDUp/files/關於京元/KYEC(2449TW)_2Q25 Investor update_English.pdf"
CNA0306 = "https://www.cna.com.tw/news/afe/202603060261.aspx"
TN0511 = "https://finance.technews.tw/2026/05/11/king-yuan-electrics-first-quarter-financial-report-hit-a-new-quarterly-high/"
TN1008 = "https://finance.technews.tw/2026/10/08/kyec-2449-202609-financial-report/"
HSM = "https://histock.tw/stock/financial.aspx?no=2449"
MW0901 = "https://www.moneyweekly.com.tw/_Article?AID=246195"
MW0924 = "https://www.moneyweekly.com.tw/_Article?AID=251167"
UA0918 = "https://uanalyze.com.tw/articles/5209555591"
MS0623 = "https://finance.technews.tw/2026/06/23/morgan-stanley-expects-king-yuan-electronics-revenue-to-grow-by-40-in-2026/"
YAM0127 = "https://n.yam.com/Article/20260127965778"
CT0728 = "https://wantrich.chinatimes.com/news/20260728900199-420101"
PS = "https://poorstock.com/earningcall/2449"
FG1128 = "https://blog.fugle.tw/post/earnings-call-2449-2025-11-28"
FS0406 = "https://news.cnyes.com/news/id/6407870"
CNYQ = "https://www.cnyes.com/twstock/2449"
SD = "https://statementdog.com/analysis/2449/eps"

sec("A. 加速器顆數（讀 tsmc 成品 In 頁，與 HBM、台積電、Advantest 模型同一份；情境係數不另訂）")
for k, nm in PLAT:
    for y in (26, 27, 28):
        inp(f"u_{k}_{y}", f"{nm} 顆數 20{y}", "百萬顆", None, TS[f"u_{k}_{y}"][0], None, "Derived", TS[f"u_{k}_{y}"][1],
            "markets/memory UNITS（經台積電模型；顆數本身多為 Interested-party／Assumed，見上游）" if y == 26 else "同上")
    inp(f"g29_{k}", f"{nm} 2029 顆數成長", "%", None, TS[f"g29_{k}"][0], None, "Assumed", TS[f"g29_{k}"][1], "沿用台積電模型")
for s in SC:
    inp(f"um28_{s}", f"情境 {s}：2028 顆數 ÷ 基準", "x", None, TS[f"um28_{s}"][0], None, "Assumed", TS[f"um28_{s}"][1], "A＝基準；B＝AI 資本支出消化；C＝明顯下修（與台積電、Advantest、台達電模型相同）")
    inp(f"um29_{s}", f"情境 {s}：2029 顆數 ÷ 基準", "x", None, TS[f"um29_{s}"][0], None, "Assumed", TS[f"um29_{s}"][1], "同上")
inp("fx", "匯率（新台幣 ÷ 美元；只用於每顆美元換算的顯示）", "NT$/US$", *SIL.vals("fx_twd"), "Derived", "與台達電、台光電模型相同", "不影響營收與 EPS")

sec("B. 每顆測試時間（含 burn-in）係數與京元電曝險（Blackwell＝1.0；NVIDIA 曝險＝1.0）")
TT = (("bw", 1.0, 1.0, 1.0, "基準世代（定義為 1.0）"),
      ("rb", 1.5, 1.2, 1.8, "Advantest 模型 ATE 測試時間 1.3（Analogy）＋ burn-in 加長：Morgan Stanley 估 R 系列 burn-in 增量占 2026 營收 3–5%（2026-06-23）"),
      ("ru", 2.0, 1.5, 2.6, "Advantest 模型 1.7（4 顆光罩尺寸晶粒＋16 疊 HBM4E）＋ burn-in 功耗更高；Assumed"),
      ("hp", 0.6, 0.5, 0.8, "Advantest 模型 0.6；上一世代"),
      ("amd", 1.1, 0.9, 1.4, "Advantest 模型 1.1"),
      ("tpu", 0.9, 0.7, 1.2, "Advantest 模型 1.0；ASIC 的 burn-in 時間短於 GPU（Assumed）"),
      ("trn", 0.8, 0.6, 1.1, "Advantest 模型 0.8"),
      ("oth", 0.7, 0.5, 1.0, "Advantest 模型 0.7"))
for k, b, lo, hi, src in TT:
    if k == "bw":
        inp(f"tt_{k}", "每顆測試時間（含 burn-in）係數：NVIDIA Blackwell", "x", None, 1.0, None, "Assumed", "定義：Blackwell＝1.0", "校準在 FY2026 進行，係數只決定世代與平台間的相對權重")
    else:
        inp(f"tt_{k}", f"每顆測試時間（含 burn-in）係數：{dict(PLAT)[k]}", "x", lo, b, hi, "Analogy" if k in ("hp", "amd", "trn", "oth") else "Assumed", src,
            "公司：GPU 客戶要求更高可靠度，burn-in 後需再做最終測試，總測試時間拉長（理財周刊 2026-09-24）" if k == "rb" else "第二來源：未找到")
EX = (("amd", 0.2, 0.1, 0.4, "AMD 委外測試主要在日月光、Amkor（Assumed）；京元電份額：未找到"),
      ("tpu", 0.5, 0.3, 0.7, "Google 相關（TPU、CPU、最終測試、部分晶圓測試）約占 2026 營收 10%（Morgan Stanley 2026-06-23）；聯發科 TPU 後段測試 9 月起量產（優分析 2026-09-18）"),
      ("trn", 0.4, 0.2, 0.6, "AWS ASIC 專案出貨中（優分析 2026-09-18）；份額：未找到"),
      ("oth", 0.4, 0.2, 0.6, "Meta 等 ASIC（理財周刊 2026-09-24）；份額：未找到"))
for k, b, lo, hi, src in EX:
    inp(f"ex_{k}", f"京元電曝險（相對 NVIDIA＝1.0）：{dict(PLAT)[k]}", "x", lo, b, hi, "Assumed", src,
        "公司：GPU 測試是較大的營收來源，ASIC 2026–27 由低基期快速成長（富果 2025-11-28 法說摘要）" if k == "amd" else "—")

sec("C. 份額、價格與 AI 部署係數（FY2027 起）")
for y, v in ((27, (0.88, 0.95, 1.00)), (28, (0.80, 0.92, 1.00)), (29, (0.75, 0.90, 1.00))):
    inp(f"rel_{y}", f"京元電份額 20{y} ÷ 2026（相對值；對 NVIDIA 與 ASIC 委外測試）", "x", *v, "Assumed",
        "NVIDIA 分散測試供應鏈（日月光、矽品、Amkor 擴最終測試與 burn-in）；京元電份額：未找到" if y == 27 else "同上",
        "Morgan Stanley：AI GPU 與 ASIC 最終測試市占穩固（2026-06-23，Interested-party）" if y == 27 else "—")
for y, v in ((27, (0.95, 0.98, 1.02)), (28, (0.90, 0.96, 1.00)), (29, (0.88, 0.95, 1.00))):
    inp(f"pp_{y}", f"測試單價指數 20{y}（2026＝1；每單位測試時間的價格）", "x", *v, "Assumed",
        "公司未揭露測試單價：未找到；低毛利手機、網通訂單移出，產品組合改善（工商時報 2026-07-28）" if y == 27 else "同上",
        "新產能到位後價格競爭（Assumed）" if y == 27 else "—")
inp("mf", "AI 部署係數（Sens 用；FY2027 起乘在加權顆數上）", "x", 0.8, 1.0, 1.2, "Assumed", "無", "±20%")

sec("D. 營收結構錨點（AI 加速器測試＝NVIDIA GPU＋ASIC 的晶圓測試、最終測試、burn-in）")
inp("ai_25", "2025 年 AI 加速器測試營收 ÷ 營收", "%", 0.25, 0.275, 0.30, "Interested-party", "Morgan Stanley：2025 年 AI 占營收 25–30%（2026-06-23）", "中央社 2026-03-06：法人稱 AI 相關測試超過 25%；只用於 FY2025 欄拆分")
inp("ai_26", "2026 年 AI 加速器測試營收 ÷ 營收（校準用）", "%", 0.28, 0.35, 0.45, "Interested-party", "Morgan Stanley：2026 年超過 35%（2026-06-23）；中央社 2026-03-06：法人預估超過 30%",
    "媒體另稱「AI 相關」2026 年 70% 以上（理財周刊 2026-09-01、優分析 2026-09-18），口徑含 CPU、網通等，不採；公司 2Q26 資料處理應用占營收 36.3%（poorstock 摘要公司簡報），上限參考")
for y, v in ((27, (0.04, 0.12, 0.20)), (28, (0.0, 0.07, 0.12)), (29, (0.0, 0.05, 0.10))):
    inp(f"g_nai{y}", f"非 AI 營收年增 20{y}（伺服器 CPU、手機、消費、車用、工業）", "%", *v, "Assumed",
        "Morgan Stanley：最大客戶 CPU（V 系列）占 2026–27 營收 8–10%，加 Google CPU 後 2026 年 CPU 測試超過 15%（2026-06-23）" if y == 27 else "同上",
        "其他 IC 設計客戶 CPU 產能多在 2026 年底建置完成，支撐 2027 年（理財周刊 2026-09-01）；手機、網通訂單移出（工商時報 2026-07-28）" if y == 27 else "—")

sec("E. 資本支出與折舊（毛利率的關鍵）")
inp("capex_26", "2026 年資本支出（現金口徑）", "NT$B", 35.0, 42.0, 50.0, "Derived", "董事會 2026 年預算 500 億（2026-07 由 393.72 億上修；工商時報 2026-07-28、理財周刊 2026-09-24）；約五成用於廠房與無塵室（Morgan Stanley）",
    "1H26 現金資本支出 67.9＋71.0＝138.9 億（poorstock 摘要公司簡報）；2025 年預算約 370 億、現金 323.6 億 → 現金約為預算 87%；基準＝下半年加速（Assumed），高＝預算全額")
inp("m_cx", "維持性資本支出 ÷ 營收", "%", 0.06, 0.10, 0.15, "Assumed", "無：公司未拆分維持與擴產", "測試機汰換與升級（Assumed）")
inp("rev24", "2024 年營收（κ 校準用）", "NT$B", None, "=34.93354/1.3008", None, "Derived", "2025 年營收 349.34 億、年增 30.08%（雅虎新聞 2026-01-27）", "HiStock：2025 年累計年增 30.1%")
inp("dr", "新增資本支出的年折舊率（設備約 5–6 年、廠房約 20 年的混合）", "%/年", 0.09, 0.11, 0.14, "Assumed", "約一半用於廠房與無塵室（Morgan Stanley 2026-06-23）：0.5 ÷ 6＋0.5 ÷ 20 ≈ 0.11", "Recon 以同一公式回測 FY2026 折舊")
inp("ret", "既有折舊每年到期比例", "%/年", 0.0, 0.05, 0.10, "Assumed", "無", "—")
inp("v_cc", "變動現金營業成本 ÷ 營收（電力、耗材、直接人工）", "%", 0.10, 0.20, 0.30, "Assumed", "季資料（Q 頁）：營收 1Q25 73 億 → 2Q26 111 億，現金營業成本只從 32 億增至 40 億（斜率約 0.2；同季比較 2Q25→2Q26 約 0.14，Q 頁）",
    "公司：二季產能擴充後折舊、電費、人力成本上升（理財周刊 2026-09-24）；固定部分 F 由 FY2026 錨點推得")

sec("F. FY2026 錨點（1H26 實績＋3Q26 營收實績＋假設；3Q26 財報約 11 月上旬公布）")
inp("q3_rev", "3Q26 營收（月營收合計＝前 9 月 − 上半年）", "NT$B", None, "=33.475818-21.33356", None, "Verified", TN1008 + "：9 月 40.71 億、前 9 月 334.76 億（年增 34.1%）",
    "HiStock 月營收：7 月 39.91、8 月 40.80、9 月 40.71 億（" + HSM + "）")
inp("q4g", "4Q26 營收季增", "%", 0.08, 0.15, 0.22, "Assumed", "公司未給量化指引；法人估 4Q 季增約 20%（理財周刊 2026-09-01；優分析 2026-09-18 估 145.59 億）",
    "Rubin 首批測試延到 3Q 底、4Q 放量（優分析 2026-09-18）；9 月營收 40.71 億 × 3＝122 億（季增 0%）為下限參考")
inp("q3_gm", "3Q26 毛利率", "%", 0.37, 0.385, 0.40, "Assumed", "公司未給指引；2Q26 39.5%、1Q26 39.7%（公司簡報）", "二季產能擴充後折舊、電費、人力上升（理財周刊 2026-09-24）")
inp("q4_gm", "4Q26 毛利率", "%", 0.375, 0.395, 0.41, "Assumed", "4Q 營收放量、稼動率上升（Assumed）", "—")
inp("q3_da", "3Q26 折舊攤提", "NT$B", 2.80, 2.95, 3.10, "Assumed", "Q 頁：EBITDA − 營業利益，3Q25 1.91 → 4Q25 2.19 → 1Q26 2.49 → 2Q26 2.74（每季約 +0.25）", "公司未直接揭露折舊金額：未找到")
inp("q4_da", "4Q26 折舊攤提", "NT$B", 2.90, 3.15, 3.40, "Assumed", "同上", "—")
inp("q_opx", "3Q、4Q26 營業費用 ÷ 營收", "%", 0.115, 0.125, 0.14, "Assumed", "1Q26 12.5%、2Q26 14.1%（推銷管理＋研發；公司簡報）", "FY2025 10.1%")
inp("q_nir", "3Q、4Q26 歸屬母公司淨利 ÷ 營業利益", "%", 0.80, 0.85, 0.90, "Derived", "1H26：淨利 47.78 億 ÷ 營業利益 56.0 億＝0.853（公司簡報推得）", "2Q–4Q25：0.899（不含 1Q25 處分利益）")
inp("q_sh", "股數（十億股；含 2026 年盈餘配股每股 0.5 元＝5%）", "十億股", None, "=2.492/2.04*1.05", None, "Derived", "2Q26 淨利 24.92 億 ÷ EPS 2.04＝12.22 億股（理財周刊 2026-09-01）× 1.05（盈餘配股 0.5 元，2026-07-28 除權，工商時報）",
    "鉅亨網 10-22 顯示股本 12.23 億股（可能尚未反映配股）；配股依會計準則追溯調整 EPS，FY2026 起以 12.83 億股計")

sec("G. 營業費用、稅與股數（FY2027 起）")
inp("phi", "營業費用成長對營收成長的彈性", "x", 0.4, 0.6, 1.0, "Assumed", "FY2026 錨點實測（Q 頁）；FY2026 費用率上升含美國廠籌設與員工酬勞", "—")
inp("opx_adj", "營業費用額外占營收比（Sens 用；基準 0）", "pt", -0.01, 0.0, 0.01, "Assumed", "無", "—")
inp("nir_f", "歸屬母公司淨利 ÷ 營業利益", "%", 0.78, 0.84, 0.90, "Assumed", "1H26 0.853；資本支出擴大使利息費用上升（淨負債 ÷ 權益 2Q26 32.7%，poorstock 摘要公司簡報）", "2Q–4Q25 0.899")
inp("sh_g", "股數年變動", "%", 0.0, 0.0, 0.05, "Assumed", "2025 年度配股 5%；之後是否續配股：未找到", "高＝每年再配股 5%")

sec("H. 共識、券商與市場資料（只作 Recon 對帳，不作校準）")
inp("c_eps26", "FactSet 共識 FY2026 EPS 中位數（2026-04-06，15 位分析師）", "NT$", None, 9.9, None, "Interested-party", FS0406, "高 11.5、低 9.02；配股前股數口徑（推定）")
inp("c_eps27", "FactSet 共識 FY2027 EPS 中位數（2026-04-06）", "NT$", None, 14.4, None, "Interested-party", FS0406, "高 16.05、低 13.35")
inp("c_eps28", "FactSet 共識 FY2028 EPS 中位數（2026-04-06）", "NT$", None, 17.72, None, "Interested-party", FS0406, "高 18.77、低 16.67")
inp("c_eps26b", "法人 FY2026 EPS 中位數區間中點（理財周刊 2026-09-01：9.09–9.79）", "NT$", None, "=(9.09+9.79)/2", None, "Interested-party", MW0901, "樂觀 10.79")
inp("c_rev26", "法人 FY2026 營收（理財周刊 2026-09-01：近 500 億，485–500 億中點）", "NT$B", None, 49.25, None, "Interested-party", MW0901, "優分析 2026-09-18：年增約 36.7%（＝477.5 億）")
inp("c_q4rev", "法人 4Q26 營收估計", "NT$B", None, 14.559, None, "Interested-party", UA0918, "季增約 20%")
inp("ms_tp", "Morgan Stanley 目標價（2026-06-23）", "NT$", None, 388, None, "Interested-party", MS0623, "由 338 上調；法人目標價中位數 348–350（理財周刊 2026-09-01）")
inp("ms_goog", "Morgan Stanley：Google 相關占 2026 營收（TPU、CPU、最終測試、部分晶圓測試）", "%", None, 0.10, None, "Interested-party", MS0623, "口徑含 CPU，比本模型 TPU 測試寬")
inp("dp_26", "公司 2Q26 資料處理應用占營收", "%", None, 0.363, None, "Interested-party", PS + "（AI 摘要公司簡報，未核原文）", "3Q25 32.8%（富果 2025-11-28 法說摘要）")
inp("bi_25", "Burn-in 占營收（3Q25）", "%", None, 0.09, None, "Interested-party", FG1128 + "（AI 摘要公司法說）", "最終測試 58.2%、晶圓測試 30.4%")

sec("I. 價格與估值")
inp("px", "股價（2026-10-22 收盤）", "NT$", None, 262.5, None, "Interested-party", "鉅亨網 2449 個股頁（10-22 收盤 262.5，−5.58%；網址見 Sources）", "財報狗、HiStock：10-22 收盤 262.5；10-23 收盤：未找到")
inp("pe_ttm", "近 12 月 P/E（鉅亨網，10-22）", "x", None, 36.13, None, "Interested-party", CNYQ, "HiStock 10 月本益比 34.40；stockanalysis 2026-09-03 為 36.12")
sec("J. 目標價（DCF＋Comparable）與收斂法參數（黃底；Andy 可改；v0.2 由短期法／長期法改為 DCF＋Comparable）")
for s, w, nm in (("A", 0.3, "順利"), ("B", 0.5, "平平"), ("C", 0.2, "不如預期")):
    inp(f"tp_w{s}", f"目標價情境權重 {s}（{nm}）", "%", None, w, None, "Assumed", "Andy 可改（工作單 2026-10-25 預設 A 0.3／B 0.5／C 0.2）", "三者合計應為 100%（Val 頁第 0 節列示）")
inp("tp_r", "折現率／要求報酬（DCF）", "%", None, 0.10, None, "Assumed", "Andy 可改（工作單預設 10%；各市場同一值）", "DCF 與 Sens")
inp("tp_g", "長期成長率（DCF 永續；收斂法情境 B 的 FY2032 營收年增）", "%", None, 0.04, None, "Assumed", "Andy 可改（工作單預設 4%）", "—")
inp("tp_roe", "近三年平均 ROE（2023–2025；DCF 再投資率＝tp_g ÷ tp_roe）", "%", None, "=(0.1584+0.1584+0.2377)/3", None, "Derived", "玩股網年度 ROE：2023 15.84%、2024 15.84%、2025 23.77%（2025 含處分利益）；Andy 可改",
    "stockanalysis：2023 11.13%、2024 11.24%、2025 16.83%（口徑不同，平均 13.1%，不採）")
inp("tp_w_dcf", "目標價中 DCF 的權重；其餘為 Comparable", "%", None, 0.5, None, "Assumed", "Andy 可改（工作單 2026-10-25 11:53 定 DCF／Comparable 各半）", "—")
inp("fade_gA", "收斂法：情境 A 於 FY2032 的營收年增（＝tp_g＋2pt）", "%", None, "=tp_g+0.02", None, "Assumed", "Andy 可改（工作單 11:20）", "FY2030 起為收斂法，非驅動因子推導")
inp("fade_gC", "收斂法：情境 C 於 FY2032 的營收年增（＝tp_g−2pt）", "%", None, "=tp_g-0.02", None, "Assumed", "Andy 可改（工作單 11:20）；情境 B 收斂到 tp_g", "—")
inp("fade_mA", "收斂法：情境 A 的中段淨利率（＝FY2029 淨利率）", "%", None, None, None, "Derived", "PL 頁情境 A FY2029 淨利 ÷ 營收", "Andy 可改")
inp("fade_mB", "收斂法：情境 B 的中段淨利率（＝FY2029 與十年平均的中點）", "%", None, None, None, "Assumed", "十年平均淨利率（2016–2025）：未找到，以 FY2029 × 0.8 代替（工作單規則）；本列＝（FY2029＋FY2029 × 0.8）÷ 2", "Andy 可改")
inp("fade_mC", "收斂法：情境 C 的中段淨利率（＝十年平均）", "%", None, None, None, "Assumed", "十年平均淨利率：未找到，以情境 C FY2029 淨利率 × 0.8 代替", "Andy 可改")
inp("cons_tp", "市場共識目標價（賣方平均）", "NT$", None, 340.73, None, "Interested-party", "stockanalysis 統計頁：15 位分析師平均 340.73（頁面價格日期 2026-10-06）", "法人目標價中位數 348–350（理財周刊 2026-09-01）；Morgan Stanley 388（In 頁 ms_tp）")
inp("cons_feps", "共識 12 個月後前瞻 EPS（2/12 × FY2027＋10/12 × FY2028；換成配股後股數 ÷ 1.05）", "NT$", None, "=(2/12*c_eps27+10/12*c_eps28)/1.05", None, "Derived", "In 頁 FactSet 共識 FY2027、FY2028 中位數（2026-04-06；推定為配股前口徑）", "共識只放 Recon 與本列，不作校準；共識日期較舊")

r = 4
for srow in rows:
    if srow[0] == "sec":
        IN.cell(r, 1, srow[1]).font = BOLD; r += 1; continue
    key, item, unit, lo, base, hi, tag, s1, s2 = srow
    IN.cell(r, 1, key); IN.cell(r, 2, item); IN.cell(r, 3, unit)
    for j, v in ((4, lo), (5, base), (6, hi)):
        if v is not None:
            IN.cell(r, j, v)
            if not (isinstance(v, str) and v.startswith("=")): IN.cell(r, j).font = BLUE
    IN.cell(r, 7, tag); IN.cell(r, 8, s1); IN.cell(r, 9, s2)
    if key.startswith(("tp_", "fade_")):
        for j in range(1, 10): IN.cell(r, j).fill = KEY
    wb.defined_names[key] = DefinedName(key, attr_text=f"In!$E${r}")
    if lo is not None: wb.defined_names[key + "_lo"] = DefinedName(key + "_lo", attr_text=f"In!$D${r}")
    if hi is not None: wb.defined_names[key + "_hi"] = DefinedName(key + "_hi", attr_text=f"In!$F${r}")
    r += 1
for col, w in zip("ABCDEFGHI", [12, 58, 10, 9, 11, 9, 18, 70, 70]): IN.column_dimensions[col].width = w
IN.freeze_panes = "D4"
def defname(name, ref):
    wb.defined_names[name] = DefinedName(name, attr_text=ref)
# tt_bw 沒有低／高；Sens 的 tt 低／高案例用同名別名
for suf in ("_lo", "_hi"):
    defname("tt_bw" + suf, wb.defined_names["tt_bw"].attr_text)

# ---------------- Q：實績與 FY2026 錨點 ----------------
Q = wb.create_sheet("Q")
Q["A1"] = "Q — 實績與 FY2026 錨點（十億新台幣；EPS 為新台幣；股數為十億股）。藍字＝輸入；黑字＝公式。季資料取自京元電每季 Investor update（營收、毛利率、推銷管理費用、研發、EBITDA、EPS）；營業利益＝毛利 − 營業費用，折舊攤提＝EBITDA − 營業利益（Derived）"
Q["A2"] = "FY2026 錨點＝1H26 實績＋3Q26 營收實績（月營收）＋3Q、4Q 假設（毛利率、折舊、營業費用率、淨利 ÷ 營業利益；4Q 營收＝3Q × (1＋季增)）。3Q26 財報約 11 月上旬公布（確切日期：未找到），公布後更新。FY2025 淨利含 1Q25 出售蘇州京隆科技的處分利益（法人估約每股 2 元，未拆出金額：未找到）"
hdr(Q, 3, ["項目", "FY2025A", "1H26（實績）", "3Q26（營收實績）", "4Q26（模型）", "FY2026 錨點", "標記", "來源1", "來源2／說明"])
S1 = "季資料（本頁下方）：公司 Investor update 1Q25–2Q26（" + DECK2Q25 + "、" + DECK4Q25 + "、" + DECK2Q26 + "）"
S2 = "FY2025：中央社 2026-03-06（營收 349.33 億、毛利率 35.84%、營業利益 90.02 億、淨利 110.15 億、EPS 9.01）；3Q26：月營收"
QQ = 26   # 季資料第一列
def qrow(r, label, vals, fmt="#,##0.000", tag="Verified", s1=S1, s2=S2, ann=None):
    Q.cell(r, 1, label)
    for col, v in vals.items():
        c = Q[f"{col}{r}"]; c.value = v; c.number_format = fmt
        if not (isinstance(v, str) and v.startswith("=")): c.font = BLUE
    if ann != "":
        Q[f"F{r}"] = ann or f"=SUM(C{r}:E{r})"; Q[f"F{r}"].number_format = fmt
    Q.cell(r, 7, tag); Q.cell(r, 8, s1); Q.cell(r, 9, s2)
qs = lambda rr: {"B": f"=SUM(B{rr}:E{rr})", "C": f"=SUM(F{rr}:G{rr})"}
qrow(4, "營收", {**qs(QQ), "D": "=q3_rev", "E": "=D4*(1+q4g)"}, tag="Verified／Assumed（4Q 季增）", s2="FY2025 月營收合計 349.34 億（HiStock）、中央社 349.33 億；1H26 213.34 億（HiStock 月營收）；" + "3Q26：" + TN1008)
qrow(5, "毛利", {**qs(QQ + 5), "D": "=D4*q3_gm", "E": "=E4*q4_gm"}, tag="Derived（營收 × 毛利率）／Assumed")
qrow(6, "毛利率", {c: f"={c}5/{c}4" for c in "BCDEF"}, fmt="0.0%", ann="", tag="Derived", s1="—", s2="FY2025 35.84%（中央社 2026-03-06）")
qrow(7, "折舊攤提（EBITDA − 營業利益）", {**qs(QQ + 9), "D": "=q3_da", "E": "=q4_da"}, tag="Derived／Assumed", s2="公司未直接揭露折舊金額：未找到；以公司 EBITDA 減營業利益推得")
qrow(8, "現金營業成本（營收 − 毛利 − 折舊攤提）", {c: f"={c}4-{c}5-{c}7" for c in "BCDE"}, tag="Derived", s1="—", s2="假設折舊攤提全部在營業成本（部分在營業費用，Assumed）")
qrow(9, "營業費用（推銷管理＋研發）", {**qs(QQ + 6), "D": "=D4*q_opx", "E": "=E4*q_opx"}, tag="Verified／Assumed")
qrow(10, "營業利益", {c: f"={c}5-{c}9" for c in "BCDE"}, tag="Derived", s2="FY2025 公司 90.02 億（中央社）；1Q26 營益率 27.2%（科技新報 2026-05-11）")
qrow(11, "營業利益率", {c: f"={c}10/{c}4" for c in "BCDEF"}, fmt="0.0%", ann="", tag="Derived", s1="—", s2="—")
qrow(12, "歸屬母公司淨利", {"B": 11.015, "C": "=2.286+2.492", "D": "=D10*q_nir", "E": "=E10*q_nir"}, tag="Verified／Derived（3Q、4Q）",
     s1="FY2025 110.15 億（中央社 2026-03-06）；1Q26 22.86 億（科技新報 2026-05-11）、2Q26 24.92 億（理財周刊 2026-09-01）", s2="1H26 47.78 億（理財周刊）")
qrow(13, "淨利 ÷ 營業利益", {c: f"={c}12/{c}10" for c in "BCDEF"}, fmt="0.0%", ann="", tag="Derived", s1="—", s2="FY2025 含處分利益")
qrow(14, "股數（十億股）", {"B": "=11.015/9.01", "C": "=4.778/3.91", "D": "=q_sh", "E": "=q_sh"}, ann="=q_sh", tag="Derived", s1="淨利 ÷ EPS（FY2025、1H26 為配股前）", s2="FY2026 錨點用配股後 12.83 億股（In 頁 q_sh）")
qrow(15, "EPS", {c: f"={c}12/{c}14" for c in "BCDE"}, fmt="0.00", ann="=F12/F14", tag="Derived", s1="淨利 ÷ 股數", s2="FY2026 以配股後股數計（1H26 追溯後約 3.72）")
qrow(16, "  公司揭露 EPS（基本）", {"B": 9.01, "C": 3.91}, fmt="0.00", ann="", tag="Verified", s1="中央社 2026-03-06；公司簡報 1Q26 1.87、2Q26 2.04", s2="Yahoo 股市、HiStock 同值")
qrow(17, "資本支出（現金）", {"B": 32.3554, "C": "=6.7858+7.0978"}, ann="=capex_26", tag="Verified（單一來源）／Derived", s1=PS + "（AI 摘要公司簡報：2023 77.3、2024 148.6、2025 323.6 億；1Q26 67.9、2Q26 71.0 億）",
     s2="預算：2025 約 370 億（雅虎 2026-01-27）、2026 年 500 億（工商時報 2026-07-28）")
qrow(18, "AI 加速器測試營收（＝營收 × AI 占比）", {"B": "=B4*ai_25"}, ann="=F4*ai_26", tag="Derived／Interested-party", s1="In 頁 D 節", s2="公司未揭露 AI 營收金額：未找到")
qrow(19, "非 AI 營收（＝營收 − AI）", {"B": "=B4-B18"}, ann="=F4-F18", tag="Derived", s1="—", s2="伺服器 CPU、手機、消費、車用、工業、網通")
Q.cell(21, 1, "核對（參考）").font = BOLD
Q.cell(22, 1, "FY2025 營業利益（季資料推得）− 公司 90.02 億"); Q["B22"] = "=B10-9.002"; Q["B22"].number_format = "0.000"
Q.cell(23, 1, "FY2025 毛利率（季資料推得）− 公司 35.84%"); Q["B23"] = "=B6-0.3584"; Q["B23"].number_format = "0.000%"
Q.cell(24, 1, "FY2026 營收年增（錨點）"); Q["B24"] = "=F4/B4-1"; Q["B24"].number_format = "0.0%"
hdr(Q, 25, ["季資料（公司 Investor update）", "1Q25", "2Q25", "3Q25", "4Q25", "1Q26", "2Q26", "標記", "來源"])
QD = [("營收", [7.3155, 8.3619, 9.2905, 9.9657, 10.1919, 11.1416], "#,##0.000", "Verified", "公司簡報（2Q25、4Q25、2Q26 Investor update）；與 HiStock 月營收合計相符"),
      ("毛利率", [0.335, 0.355, 0.360, 0.377, 0.3974, 0.395], "0.0%", "Verified", "公司簡報（四捨五入至 0.1%）；1Q26 39.74%（科技新報 2026-05-11）"),
      ("推銷及管理費用", [0.6851, 0.5833, 0.6369, 0.6841, 0.9713, 1.2573], "#,##0.000", "Verified", "公司簡報"),
      ("研發費用", [0.2426, 0.2121, 0.2178, 0.2578, 0.3073, 0.3150], "#,##0.000", "Verified", "公司簡報"),
      ("EBITDA", [3.1729, 3.9445, 4.3960, 5.0021, 5.2584, 5.5696], "#,##0.000", "Verified", "公司簡報（定義未載明；本模型視為營業利益＋折舊攤提）")]
for i, (lab, vals, fmt, tag, src) in enumerate(QD):
    rr = QQ + i; Q.cell(rr, 1, lab)
    for j, v in enumerate(vals):
        c = Q.cell(rr, 2 + j, v); c.font = BLUE; c.number_format = fmt
    Q.cell(rr, 8, tag); Q.cell(rr, 9, src)
for i, (lab, f, fmt) in enumerate((("毛利（營收 × 毛利率）", "={c}{a}*{c}{b}", "#,##0.000"), ("營業費用（推銷管理＋研發）", "={c}{d}+{c}{e}", "#,##0.000"),
                                   ("營業利益（毛利 − 營業費用）", "={c}{g}-{c}{h}", "#,##0.000"), ("營業利益率", "={c}{i}/{c}{a}", "0.0%"),
                                   ("折舊攤提（EBITDA − 營業利益）", "={c}{f}-{c}{i}", "#,##0.000"), ("現金營業成本（營收 − 毛利 − 折舊攤提）", "={c}{a}-{c}{g}-{c}{j}", "#,##0.000"))):
    rr = QQ + 5 + i; Q.cell(rr, 1, lab)
    for c in "BCDEFG":
        Q[f"{c}{rr}"] = f.format(c=c, a=QQ, b=QQ + 1, d=QQ + 2, e=QQ + 3, f=QQ + 4, g=QQ + 5, h=QQ + 6, i=QQ + 7, j=QQ + 9); Q[f"{c}{rr}"].number_format = fmt
    Q.cell(rr, 8, "Derived")
rr = QQ + 11; Q.cell(rr, 1, "EPS（基本，公司揭露）")
for j, v in enumerate([3.51, 1.78, 1.88, 1.84, 1.87, 2.04]):
    c = Q.cell(rr, 2 + j, v); c.font = BLUE; c.number_format = "0.00"
Q.cell(rr, 8, "Verified"); Q.cell(rr, 9, "公司簡報；Yahoo 股市、HiStock 同值；1Q25 含處分利益")
rr += 1; Q.cell(rr, 1, "現金營業成本季增 ÷ 營收季增（2Q25→2Q26；In 頁 v_cc 的參考）"); Q[f"B{rr}"] = f"=(G{QQ + 10}-C{QQ + 10})/(G{QQ}-C{QQ})"; Q[f"B{rr}"].number_format = "0.00"
Q.cell(rr, 8, "Derived"); Q.cell(rr, 9, "同季比較排除季節性；斜率小代表現金成本多為固定（人力、電力基本量）")
QSLOPE = rr
for nm, rr_ in (("rev", 4), ("gp", 5), ("da", 7), ("cc", 8), ("opx", 9), ("oi", 10), ("ni", 12), ("sh", 14), ("eps", 15), ("capex", 17)):
    defname(f"a25_{nm}", f"Q!$B${rr_}"); defname(f"a26_{nm}", f"Q!$F${rr_}")
Q.column_dimensions["A"].width = 60
for c in "BCDEFG": Q.column_dimensions[c].width = 15
Q.column_dimensions["H"].width = 40; Q.column_dimensions["I"].width = 90
Q.freeze_panes = "B4"

# ---------------- 引擎 ----------------
FYS = (2025, 2026, 2027, 2028, 2029); FC = {2025: "C", 2026: "D", 2027: "E", 2028: "F", 2029: "G"}
BU = (2026, 2027, 2028, 2029)
def uref(k, y, s, mf):
    if y == 2026: return f"u_{k}_26"
    if y == 2027: return f"u_{k}_27*{mf}"
    if y == 2028: return f"u_{k}_28*um28_{s}*{mf}"
    return f"u_{k}_28*(1+g29_{k})*um29_{s}*{mf}"
DEFP = dict(scn="A", tt="", ex="", rel="rel", pp="pp", mf="1", ais="ai_26", gn="g_nai", v="v_cc", dr="dr", ret="ret", m="m_cx", phi="phi", nir="nir_f", opxa="0")
def engine(ws, r0, title, **over):
    P = dict(DEFP); P.update(over); s = P["scn"]
    ws.cell(r0, 1, title).font = BOLD
    R = {}; r = r0 + 1
    hdr(ws, r, ["項目", "單位", "FY2025A", "FY2026（錨點）", "FY2027", "FY2028", "FY2029", "說明"]); r += 1
    def row(key, label, unit, fs, fmt="#,##0.00", note=""):
        nonlocal r
        ws.cell(r, 1, label); ws.cell(r, 2, unit)
        for col, f in fs.items():
            ws[f"{col}{r}"] = f; ws[f"{col}{r}"].number_format = fmt
        ws.cell(r, 8, note); R[key] = r; r += 1
    def sec_(t):
        nonlocal r
        ws.cell(r, 1, t).font = BOLD; r += 1
    c_ = lambda y: FC[y]
    sec_("A. AI 加速器測試需求（加權顆數＝顆數 × 每顆測試時間（含 burn-in）× 京元電曝險；FY(t)＝CY(t)）")
    for k, nm in PLAT:
        exf = "" if k in NV else f"*ex_{k}{P['ex']}"
        row(f"w_{k}", f"  {nm}：加權顆數", "百萬顆", {c_(y): f"={uref(k, y, s, P['mf'])}*tt_{k}{P['tt']}{exf}" for y in BU}, fmt="0.00")
    row("u", "加速器顆數合計（未加權）", "百萬顆", {c_(y): "=" + "+".join(uref(k, y, s, P["mf"]) for k, _ in PLAT) for y in BU}, fmt="0.00")
    row("wnv", "NVIDIA 加權顆數", "百萬顆", {c_(y): "=" + "+".join(f"{c_(y)}{R['w_' + k]}" for k in NV) for y in BU}, fmt="0.00")
    row("wax", "AMD 與 ASIC 加權顆數（× 曝險）", "百萬顆", {c_(y): "=" + "+".join(f"{c_(y)}{R['w_' + k]}" for k in AX) for y in BU}, fmt="0.00")
    row("w", "加權顆數合計", "百萬顆", {c_(y): f"={c_(y)}{R['wnv']}+{c_(y)}{R['wax']}" for y in BU}, fmt="0.00")
    row("ai0", "FY2026 AI 加速器測試營收（校準）＝ 錨點營收 × AI 占比", "NT$B", {"D": f"=a26_rev*{P['ais']}"}, fmt="#,##0.00")
    row("rpu", "每加權顆營收（FY2026 校準）", "NT$B／百萬顆", {"D": f"=D{R['ai0']}/D{R['w']}"}, fmt="0.0000",
        note="Derived：＝ 千元新台幣／顆；乘積（測試時間 × 每小時價格 × 份額）只能整體識別")
    row("rel", "份額相對值（÷ FY2026）", "x", {"D": "=1", **{c_(y): f"={P['rel']}_{y % 100}" for y in BU[1:]}}, fmt="0.00")
    row("pp", "測試單價指數（FY2026＝1）", "x", {"D": "=1", **{c_(y): f"={P['pp']}_{y % 100}" for y in BU[1:]}}, fmt="0.00")
    row("ai", "AI 加速器測試營收 ＝ 每加權顆營收 × 加權顆數 × 份額 × 單價", "NT$B", {"C": "=a25_rev*ai_25", "D": f"=D{R['ai0']}",
        **{c_(y): f"=$D${R['rpu']}*{c_(y)}{R['w']}*{c_(y)}{R['rel']}*{c_(y)}{R['pp']}" for y in BU[1:]}}, fmt="#,##0.00")
    row("nai", "非 AI 營收", "NT$B", {"C": f"=a25_rev-C{R['ai']}", "D": f"=a26_rev-D{R['ai']}", **{c_(y): f"={c_(y - 1)}{r}*(1+{P['gn']}{y % 100})" for y in BU[1:]}}, fmt="#,##0.00")
    row("rev", "營收", "NT$B", {c_(y): f"={c_(y)}{R['ai']}+{c_(y)}{R['nai']}" for y in FYS}, fmt="#,##0.00")
    row("g", "營收年增", "%", {c_(y): f"={c_(y)}{R['rev']}/{c_(y - 1)}{R['rev']}-1" for y in FYS[1:]}, fmt="0.0%")
    row("aish", "AI 加速器測試 ÷ 營收", "%", {c_(y): f"={c_(y)}{R['ai']}/{c_(y)}{R['rev']}" for y in FYS}, fmt="0.0%")
    row("nvsh", "NVIDIA ÷ 加權顆數", "%", {c_(y): f"={c_(y)}{R['wnv']}/{c_(y)}{R['w']}" for y in BU}, fmt="0.0%")
    row("pue", "每顆 Blackwell 等值的京元電 AI 測試營收", "US$／顆", {c_(y): f"={c_(y)}{R['ai']}/{c_(y)}{R['w']}/fx*1000" for y in BU}, fmt="#,##0.0",
        note="含份額與單價指數；FY2027 起隨份額、單價變動")
    sec_("B. 資本支出與折舊（資本支出＝維持性 × 營收＋κ × 當年營收增量；折舊＝前一年 × (1 − 到期)＋折舊率 × 當年與前一年資本支出平均）")
    row("kap", "κ：每增加 1 元年營收所需擴產資本支出（FY2025 校準）", "x", {"C": f"=(a25_capex-{P['m']}*a25_rev)/(a25_rev-rev24)"}, fmt="0.00",
        note="Derived：（FY2025 現金資本支出 − 維持性）÷（FY2025 營收 − FY2024 營收）")
    row("capex", "資本支出", "NT$B", {"C": "=a25_capex", "D": "=capex_26",
        **{c_(y): f"={P['m']}*{c_(y)}{R['rev']}+$C${R['kap']}*MAX(0,{c_(y)}{R['rev']}-{c_(y - 1)}{R['rev']})" for y in BU[1:]}}, fmt="#,##0.00",
        note="FY2026＝In 頁 capex_26（公式值見 Recon）；營收下滑時只剩維持性資本支出")
    row("cxr", "資本支出 ÷ 營收", "%", {c_(y): f"={c_(y)}{R['capex']}/{c_(y)}{R['rev']}" for y in FYS}, fmt="0.0%")
    row("dep", "折舊攤提", "NT$B", {"C": "=a25_da", "D": "=a26_da",
        **{c_(y): f"={c_(y - 1)}{r}*(1-{P['ret']})+{P['dr']}*({c_(y - 1)}{R['capex']}+{c_(y)}{R['capex']})/2" for y in BU[1:]}}, fmt="#,##0.00")
    row("depr", "折舊攤提 ÷ 營收", "%", {c_(y): f"={c_(y)}{R['dep']}/{c_(y)}{R['rev']}" for y in FYS}, fmt="0.0%")
    row("depbt", "回測：以同一公式算 FY2026 折舊 − 錨點", "NT$B", {"D": f"=C{R['dep']}*(1-{P['ret']})+{P['dr']}*(C{R['capex']}+D{R['capex']})/2-D{R['dep']}"}, fmt="0.00",
        note="校準檢查（非獨立檢驗）：正值＝公式高估")
    sec_("C. 損益")
    row("ccF", "固定現金成本 F（FY2026 校準＝錨點現金成本 − v × 營收）", "NT$B", {"D": f"=a26_cc-{P['v']}*a26_rev"}, fmt="#,##0.00",
        note="Derived：人力、電力基本量、租金等；FY2026 恆等於錨點")
    row("cc", "現金營業成本 ＝ F ＋ v ×（AI 營收 ÷ 單價指數＋非 AI 營收）", "NT$B", {"C": "=a25_cc", "D": "=a26_cc",
        **{c_(y): f"=$D${R['ccF']}+{P['v']}*({c_(y)}{R['ai']}/{c_(y)}{R['pp']}+{c_(y)}{R['nai']})" for y in BU[1:]}}, fmt="#,##0.00",
        note="成本隨測試量（不含價格）變動；單價下降全部落在毛利")
    row("gp", "毛利 ＝ 營收 − 現金營業成本 − 折舊攤提", "NT$B", {c_(y): f"={c_(y)}{R['rev']}-{c_(y)}{R['cc']}-{c_(y)}{R['dep']}" for y in FYS}, fmt="#,##0.00")
    row("gm", "毛利率", "%", {c_(y): f"={c_(y)}{R['gp']}/{c_(y)}{R['rev']}" for y in FYS}, fmt="0.0%")
    row("opb", "營業費用（基準路徑）", "NT$B", {"C": "=a25_opx", "D": "=a26_opx", **{c_(y): f"={c_(y - 1)}{r}*(1+MAX(0,{P['phi']}*{c_(y)}{R['g']}))" for y in BU[1:]}}, fmt="#,##0.00")
    row("opx", "營業費用", "NT$B", {"C": f"=C{R['opb']}", "D": f"=D{R['opb']}", **{c_(y): f"={c_(y)}{R['opb']}+{P['opxa']}*{c_(y)}{R['rev']}" for y in BU[1:]}}, fmt="#,##0.00")
    row("oi", "營業利益", "NT$B", {c_(y): f"={c_(y)}{R['gp']}-{c_(y)}{R['opx']}" for y in FYS}, fmt="#,##0.00")
    row("oim", "營業利益率", "%", {c_(y): f"={c_(y)}{R['oi']}/{c_(y)}{R['rev']}" for y in FYS}, fmt="0.0%")
    row("ebitda", "EBITDA（營業利益＋折舊攤提）", "NT$B", {c_(y): f"={c_(y)}{R['oi']}+{c_(y)}{R['dep']}" for y in FYS}, fmt="#,##0.00")
    row("fcf", "EBITDA − 資本支出（近似，未計稅與營運資金）", "NT$B", {c_(y): f"={c_(y)}{R['ebitda']}-{c_(y)}{R['capex']}" for y in FYS}, fmt="#,##0.00")
    row("ni", "歸屬母公司淨利", "NT$B", {"C": "=a25_ni", "D": "=a26_ni", **{c_(y): f"={c_(y)}{R['oi']}*{P['nir']}" for y in BU[1:]}}, fmt="#,##0.00", note="FY2025 含處分利益")
    row("nioi", "淨利 ÷ 營業利益", "%", {c_(y): f"={c_(y)}{R['ni']}/{c_(y)}{R['oi']}" for y in FYS}, fmt="0.0%")
    row("sh_", "股數", "十億股", {"C": "=a25_sh", "D": "=a26_sh", **{c_(y): f"={c_(y - 1)}{r}*(1+sh_g)" for y in BU[1:]}}, fmt="0.0000")
    row("eps", "EPS", "NT$", {c_(y): f"={c_(y)}{R['ni']}/{c_(y)}{R['sh_']}" for y in FYS})
    row("epsg", "EPS 年增", "%", {c_(y): f"={c_(y)}{R['eps']}/{c_(y - 1)}{R['eps']}-1" for y in FYS[1:]}, fmt="0.0%", note="FY2026 對 FY2025：股數口徑不同（配股 5%）且 FY2025 含處分利益")
    R["end"] = r
    return R

# ---------------- PL ----------------
PL = wb.create_sheet("PL")
PL["A1"] = "PL — 京元電損益（十億新台幣；EPS 新台幣；歸屬母公司）。FY2025A 實績、FY2026 錨點（1H26 實績＋3Q26 營收實績＋假設），FY2027 起為模型。FY(t)＝CY(t)。三情境各一個完整推導區塊（本頁下方），頂端為摘要。FY2026–FY2027 三情境相同（顆數係數從 CY2028 開始）"
PB = {}; r0 = 60
SCN_T = {"A": "情境 A：基準（顆數係數 1.0）", "B": "情境 B：CY2028 顆數 ×0.8、CY2029 ×0.85", "C": "情境 C：CY2028–29 顆數 ×0.6"}
for s in SC:
    PB[s] = engine(PL, r0, SCN_T[s], scn=s)
    r0 = PB[s]["end"] + 2
# 收斂法（FY2030–FY2032）：上游顆數只到 CY2029，之後營收年增與淨利率線性收斂（Andy 2026-10-25 11:20）
FADE = PatternFill("solid", fgColor="E2EFDA")
FX = {2030: "H", 2031: "I", 2032: "J"}
FB = {}; fr = PB["C"]["end"] + 2
gT = {"A": "fade_gA", "B": "tp_g", "C": "fade_gC"}
for s_ in SC:
    PL.cell(fr, 1, f"收斂法（FY2030–FY2032；非驅動因子推導）：情境 {s_}").font = BOLD; fr += 1
    hdr(PL, fr, ["項目", "單位", "", "", "", "", "FY2029（模型）", "FY2030（收斂）", "FY2031（收斂）", "FY2032（收斂）", "說明"]); fr += 1
    E = PB[s_]; F = {}
    def frow(key, label, unit, g_, fs, fmt, note=""):
        global fr
        PL.cell(fr, 1, label); PL.cell(fr, 2, unit)
        PL[f"G{fr}"] = g_; PL[f"G{fr}"].number_format = fmt
        for col, f in fs.items():
            PL[f"{col}{fr}"] = f; PL[f"{col}{fr}"].number_format = fmt; PL[f"{col}{fr}"].fill = FADE
        PL.cell(fr, 11, note); F[key] = fr; fr += 1
    frow("g", "營收年增（FY2029 → 收斂目標，線性，FY2032 達到）", "%", f"=G{E['g']}", {FX[y]: f"=$G{fr}+({gT[s_]}-$G{fr})*{y - 2029}/3" for y in FX}, "0.0%", f"收斂目標＝In 頁 {gT[s_]}")
    frow("rev", "營收", "NT$B", f"=G{E['rev']}", {FX[y]: f"={chr(ord(FX[y]) - 1)}{fr}*(1+{FX[y]}{F['g']})" for y in FX}, "#,##0.00")
    frow("nm", "淨利率（FY2029 → 中段利潤率，線性）", "%", f"=G{E['ni']}/G{E['rev']}", {FX[y]: f"=$G{fr}+(fade_m{s_}-$G{fr})*{y - 2029}/3" for y in FX}, "0.0%", f"收斂目標＝In 頁 fade_m{s_}")
    frow("ni", "歸屬母公司淨利", "NT$B", f"=G{E['ni']}", {FX[y]: f"={FX[y]}{F['rev']}*{FX[y]}{F['nm']}" for y in FX}, "#,##0.00")
    frow("sh", "股數（沿用 sh_g）", "十億股", f"=G{E['sh_']}", {FX[y]: f"={chr(ord(FX[y]) - 1)}{fr}*(1+sh_g)" for y in FX}, "0.0000")
    frow("eps", "EPS", "NT$", f"=G{E['eps']}", {FX[y]: f"={FX[y]}{F['ni']}/{FX[y]}{F['sh']}" for y in FX}, "0.00")
    FB[s_] = F; fr += 1
FADE_NOTE = fr
PL.cell(fr, 1, "FY2030 起為收斂法，非驅動因子推導：營收年增由 FY2029 線性收斂到 In 頁的長期成長（A＝tp_g＋2pt、B＝tp_g、C＝tp_g−2pt），淨利率由 FY2029 線性收斂到中段利潤率（fade_mA／mB／mC）；綠底為收斂年度。")
hdr(PL, 3, ["情境／項目", "單位", "FY2025A", "FY2026（錨點）", "FY2027", "FY2028", "FY2029", "FY2030（收斂）", "FY2031（收斂）", "FY2032（收斂）", "說明"])
sr = 4; SUMR = {}
for s in SC:
    PL.cell(sr, 1, f"情境 {s}").font = BOLD; sr += 1
    for key, lab, unit, fmt in (("rev", "營收", "NT$B", "#,##0.0"), ("g", "  營收年增", "%", "0.0%"), ("ai", "  AI 加速器測試營收", "NT$B", "#,##0.0"),
                                ("nai", "  非 AI 營收", "NT$B", "#,##0.0"), ("capex", "  資本支出", "NT$B", "#,##0.0"), ("depr", "  折舊攤提 ÷ 營收", "%", "0.0%"),
                                ("gm", "毛利率", "%", "0.0%"), ("oim", "營業利益率", "%", "0.0%"), ("nm", "淨利率", "%", "0.0%"),
                                ("ni", "歸屬母公司淨利", "NT$B", "#,##0.0"), ("eps", "EPS", "NT$", "0.00"), ("epsg", "  EPS 年增", "%", "0.0%")):
        PL.cell(sr, 1, lab); PL.cell(sr, 2, unit)
        for y in FYS:
            if key in ("g", "epsg") and y == 2025: continue
            PL[f"{FC[y]}{sr}"] = f"={FC[y]}{PB[s]['ni']}/{FC[y]}{PB[s]['rev']}" if key == "nm" else f"={FC[y]}{PB[s][key]}"; PL[f"{FC[y]}{sr}"].number_format = fmt
        if key in ("rev", "g", "nm", "ni", "eps"):
            for y in FX:
                PL[f"{FX[y]}{sr}"] = f"={FX[y]}{FB[s][key]}"; PL[f"{FX[y]}{sr}"].number_format = fmt; PL[f"{FX[y]}{sr}"].fill = FADE
        if key == "epsg":
            for y in FX:
                PL[f"{FX[y]}{sr}"] = f"={FX[y]}{sr - 1}/{chr(ord(FX[y]) - 1)}{sr - 1}-1"; PL[f"{FX[y]}{sr}"].number_format = fmt; PL[f"{FX[y]}{sr}"].fill = FADE
        SUMR[(s, key)] = sr; sr += 1
    sr += 1
PL.cell(sr, 1, "核對（應為 0）：FY2026 模型營收、毛利、營業利益 − 錨點（情境 A）").font = BOLD; sr += 1
CHK = sr
PL.cell(sr, 1, "  情境 A：營收差（NT$B）"); PL[f"D{sr}"] = f"=D{PB['A']['rev']}-a26_rev"; PL[f"D{sr}"].number_format = "0.0000"; sr += 1
PL.cell(sr, 1, "  情境 A：毛利差（NT$B）"); PL[f"D{sr}"] = f"=D{PB['A']['gp']}-a26_gp"; PL[f"D{sr}"].number_format = "0.0000"; sr += 1
PL.cell(sr, 1, "  情境 A：營業利益差（NT$B）"); PL[f"D{sr}"] = f"=D{PB['A']['oi']}-a26_oi"; PL[f"D{sr}"].number_format = "0.0000"; sr += 1
PL.cell(sr, 1, "  情境 A：FY2025 毛利差（NT$B）"); PL[f"C{sr}"] = f"=C{PB['A']['gp']}-a25_gp"; PL[f"C{sr}"].number_format = "0.0000"
assert sr < 60, sr
PL.column_dimensions["A"].width = 62; PL.column_dimensions["B"].width = 12; PL.column_dimensions["K"].width = 50
for c in "CDEFGHIJ": PL.column_dimensions[c].width = 14

for s_ in SC:   # fade_mA／B／C：依各情境 FY2029 淨利率（十年平均淨利率未找到 → 以 × 0.8 代替）
    rr_f = next(r_.row for r_ in IN["A"] if r_.value == f"fade_m{s_}")
    last = f"PL!G{PB[s_]['ni']}/PL!G{PB[s_]['rev']}"
    IN.cell(rr_f, 5, {"A": f"={last}", "B": f"=({last}+{last}*0.8)/2", "C": f"={last}*0.8"}[s_])
# ---------------- Units：加權顆數、每顆營收與資本密度（權威：京元電的每顆測試營收） ----------------
UN = wb.create_sheet("Units")
UN["A1"] = "Units — 加速器顆數（tsmc 成品）× 每顆測試時間（含 burn-in）× 京元電曝險 → 加權顆數；FY2026 以 AI 占比校準每加權顆營收。情境 A；FY(t)＝CY(t)"
A_ = PB["A"]
cyc = lambda y: chr(ord("C") + y - 2026)
hdr(UN, 3, ["項目", "單位", "CY2026", "CY2027", "CY2028", "CY2029", "說明"])
ur = 4; UR = {}
ulist = [("u", "加速器顆數合計（未加權）", "百萬顆", lambda y: f"=PL!{FC[y]}{A_['u']}", "0.00", "讀台積電模型 In 頁（與 HBM、Advantest 同一份）")]
for k, nm in PLAT:
    ulist.append((f"w_{k}", f"  {nm}：加權顆數", "百萬顆", (lambda kk: lambda y: f"=PL!{FC[y]}{A_['w_' + kk]}")(k), "0.00", "顆數 × 測試時間係數" + ("" if k in NV else " × 曝險")))
ulist += [("w", "加權顆數合計（Blackwell 等值）", "百萬顆", lambda y: f"=PL!{FC[y]}{A_['w']}", "0.00", ""),
          ("nvsh", "NVIDIA ÷ 加權顆數", "%", lambda y: f"=PL!{FC[y]}{A_['nvsh']}", "0.0%", "公司：GPU 為較大的營收來源（富果 2025-11-28 法說摘要）"),
          ("ai", "京元電 AI 加速器測試營收", "US$B", lambda y: f"=PL!{FC[y]}{A_['ai']}/fx", "#,##0.00", ""),
          ("pue", "每顆 Blackwell 等值的京元電 AI 測試營收", "US$／顆", lambda y: f"=PL!{FC[y]}{A_['pue']}", "#,##0.0", "FY2026 校準；之後 × 份額 × 單價"),
          ("perunit", "每顆加速器（未加權）的京元電 AI 測試營收", "US$／顆", lambda y: f"=PL!{FC[y]}{A_['ai']}/fx/PL!{FC[y]}{A_['u']}*1000", "#,##0.0", "世代組合升級使每顆金額上升"),
          ("capex", "資本支出", "NT$B", lambda y: f"=PL!{FC[y]}{A_['capex']}", "#,##0.0", ""),
          ("capw", "資本支出 ÷ 加權顆數增量", "NT$／顆", lambda y: f"=IFERROR(PL!{FC[y]}{A_['capex']}/(PL!{FC[y]}{A_['w']}-PL!{FC[y - 1]}{A_['w']})*1000,\"—\")" if y > 2026 else "=\"—\"", "#,##0", "顆數下修（增量為負）時數值無意義"),
          ("depr", "折舊攤提 ÷ 營收", "%", lambda y: f"=PL!{FC[y]}{A_['depr']}", "0.0%", "毛利率的主要變數")]
for key, lab, unit, f, fmt, note in ulist:
    UN.cell(ur, 1, lab); UN.cell(ur, 2, unit)
    for y in BU:
        UN[f"{cyc(y)}{ur}"] = f(y); UN[f"{cyc(y)}{ur}"].number_format = fmt
    UN.cell(ur, 7, note); UR[key] = ur; ur += 1
ur += 1
UN.cell(ur, 1, "讀法：加權顆數的絕對水準與曝險只影響 FY2026 校準出的每加權顆營收，不影響 FY2026 營收；影響成長的是世代組合（Rubin、Rubin Ultra 的測試時間與 burn-in）、ASIC 曝險、份額相對值與單價。京元電的份額無法單獨識別（公司不揭露客戶別營收），以「每顆 Blackwell 等值營收」與份額相對值表示。").font = BOLD
UNOTE = ur
UN.column_dimensions["A"].width = 64; UN.column_dimensions["B"].width = 14
for c in "CDEF": UN.column_dimensions[c].width = 14
UN.column_dimensions["G"].width = 60

# ---------------- Val ----------------
V = wb.create_sheet("Val")
V["A1"] = "Val — P/E 估值（新台幣）。結論用「現價 ÷ FY2027、FY2028 EPS（模型）」與十年 P/E 規則值（年底落後 P/E，基準為各情境 FY2027 EPS）。本版未做先驗，無 Andy 合理 P/E 對照"
V["A2"] = "循環位置：京元電是測試代工，不是設備股，不套設備市場見頂年；以本模型營收最高的會計年度為見頂年（比照 NVIDIA、台達電、台光電；FY(t)＝CY(t)）。營收到 FY2029 仍在增加時，見頂年未知，判為中段"
# 第 1 節（價格與市場隱含倍數）；建置最後整體下移 23 列（第 1 節起於第 51 列），第 0 節放第 4–48 列
V.cell(28, 1, "1. 價格與市場隱含倍數").font = BOLD
hdr(V, 29, ["項目", "A", "B", "C", "說明"])
V.cell(30, 1, "現價（2026-10-22 收盤）"); V["B30"] = "=px"; V["B30"].number_format = "#,##0.0"
V.cell(31, 1, "現價 ÷ FY2026 EPS（錨點）"); V.cell(32, 1, "現價 ÷ FY2027 EPS（模型）"); V.cell(33, 1, "現價 ÷ FY2028 EPS（模型）"); V.cell(34, 1, "現價 ÷ FY2029 EPS（模型）")
for j, s in enumerate(SC):
    c = "BCD"[j]
    for rr, y in ((31, 2026), (32, 2027), (33, 2028), (34, 2029)):
        V[f"{c}{rr}"] = f"=px/PL!{FC[y]}{SUMR[(s, 'eps')]}"; V[f"{c}{rr}"].number_format = "0.0"
V.cell(35, 1, "未來 12 月 EPS（模型 A）≈ ¼ FY2026＋¾ FY2027"); V["B35"] = f"=0.25*PL!D{SUMR[('A', 'eps')]}+0.75*PL!E{SUMR[('A', 'eps')]}"; V["B35"].number_format = "0.00"
V.cell(36, 1, "  現價 ÷ 未來 12 月 EPS（模型 A）"); V["B36"] = "=px/B35"; V["B36"].number_format = "0.0"
V.cell(37, 1, "近 12 月 P/E（鉅亨網，10-22）"); V["B37"] = "=pe_ttm"; V["B37"].number_format = "0.0"
V["E32"] = "FY2027 三情境相同（顆數係數從 CY2028 開始）"
V["E31"] = "FY2026 起以配股後股數計"
V.cell(39, 1, "2. 循環位置（本模型營收最高年；FY(t)＝CY(t)）").font = BOLD
hdr(V, 40, ["情境", "營收最高年（FY2026–FY2029）", "FY2027 判定", "規則倍數（P/E）", "見頂年是否未知", "說明"])
for i, s in enumerate(SC, 41):
    rr = f"PL!D{PB[s]['rev']}:G{PB[s]['rev']}"
    V.cell(i, 1, s)
    V.cell(i, 2, f"=INDEX({{2026,2027,2028,2029}},MATCH(MAX({rr}),{rr},0))")
    V.cell(i, 3, '=IF(B{0}=2029,"中段",'.format(i) + CP.excel_rule("2027", f"B{i}") + ")")
    V.cell(i, 4, f'=IF(C{i}="見頂",pe_peak,IF(C{i}="谷底",pe_trough,pe_mid))').number_format = "0.0"
    V.cell(i, 5, f'=IF(B{i}=2029,"是：營收到觀測期末仍在增加","否")')
    V.cell(i, 6, "tools/cycle_position.py excel_rule；營收最高年取代設備市場見頂年" if s == "A" else "同上")
V.cell(45, 1, "3. 規則值隱含股價（新台幣；約 FY2027 年底，未折現）＝ FY2027 EPS × 規則倍數").font = BOLD
hdr(V, 46, ["情境", "FY2027 EPS", "（同左）", "規則倍數", "隱含股價", "相對現價", "現價 ÷ FY2027 EPS"])
VAL = {}
for j, s in enumerate(SC):
    i = 47 + j
    V.cell(i, 1, s)
    V.cell(i, 2, f"=PL!E{SUMR[(s, 'eps')]}").number_format = "0.00"
    V.cell(i, 3, f"=B{i}").number_format = "0.00"
    V.cell(i, 4, f"=D{41 + j}").number_format = "0.0"
    V.cell(i, 5, f"=C{i}*D{i}").number_format = "#,##0.0"
    V.cell(i, 6, f"=E{i}/px-1").number_format = "0%"
    V.cell(i, 7, f"=px/C{i}").number_format = "0.0"
    VAL[s] = i
V.cell(51, 1, "4. What must be true：FY2027 年底價 ＝ 現價 ×（1＋報酬）所需的倍數（FY2028 EPS；約 2027 年底的價格看下一年 EPS）").font = BOLD
hdr(V, 52, ["FY2027 年底價 ÷ 現價", "所需倍數（÷ FY2028 EPS，A）", "所需倍數（÷ FY2028 EPS，C）", "對照：十年中位數", "對照：十年 75 分位"])
vr = 53
for x in (0.5, 0.75, 1.0, 1.21, 1.5):
    V.cell(vr, 1, x).font = BLUE; V.cell(vr, 1).number_format = "0.00"
    V.cell(vr, 2, f"=px*A{vr}/PL!F{SUMR[('A', 'eps')]}").number_format = "0.0"
    V.cell(vr, 3, f"=px*A{vr}/PL!F{SUMR[('C', 'eps')]}").number_format = "0.0"
    V.cell(vr, 4, "=pe_mid").number_format = "0.0"; V.cell(vr, 5, "=pe_trough").number_format = "0.0"
    vr += 1
V.cell(vr, 1, "讀法：1.00＝價格不變、1.21＝一年約 21%。所需倍數遠高於歷史中位數時，現價依賴「AI 期倍數高於 2016–2022 年」與 FY2028 以後的成長延續；倍數是否合理，另看前瞻 P/E 歷史（Bloomberg）。"); WMT_NOTE = vr
assert vr < 64, vr
RB0 = 64
# 十年年底 P/E（2016–2025）：年底收盤（Yahoo Finance 12 月收盤 × 1.05 還原 2026 年配股調整）÷ 當年 EPS（玩股網、HiStock）
CLOSE_Y = [24.00, 28.81, 22.14, 35.76, 33.10, 42.67, 34.48, 80.86, 106.19, 235.71]
EPSY = [2.56, 1.88, 1.47, 2.49, 2.97, 4.23, 5.59, 4.78, 6.36, 9.01]
PE_D = [f"={c}*1.05/{e}" for c, e in zip(CLOSE_Y, EPSY)]
PE_SA = [None] * 5 + [10.59, 6.47, 17.77, 17.52, 27.47]     # stockanalysis 年度 P/E（2021–2025）
PE_HS = [None] * 5 + [12.95, 6.35, 17.99, 18.65, 28.03]     # HiStock 12 月本益比（2021–2025）
block_end = VR.add_rule_block(V, RB0, [{"id": "P/E", "name": "P/E", "sources": [
    {"label": "年底收盤 ÷ 當年 EPS（Yahoo Finance 12 月收盤 × 1.05、玩股網／HiStock EPS）", "vals": PE_D, "tag": "Derived（2016–2025 共 10 年，皆為正值且小於 200，無剔除；Yahoo 收盤已因 2026 年配股 5% 下調，乘 1.05 還原）"},
    {"label": "stockanalysis 年度 P/E", "vals": PE_SA, "tag": "Interested-party（資料商；只有 2021–2025，只作核對，不採用）", "use": False},
    {"label": "HiStock 12 月本益比", "vals": PE_HS, "tag": "Interested-party（資料商；近四季 EPS 口徑；只有 2021–2025，不採用）", "use": False}],
    "cases": [
        {"case": "見頂", "pos": "25 分位", "q": 25, "basis_label": "FY2027 EPS（A）", "basis": f"=C{VAL['A']}"},
        {"case": "中段", "pos": "中位數", "q": 50, "basis_label": "FY2027 EPS（A）", "basis": f"=C{VAL['A']}"},
        {"case": "谷底", "pos": "75 分位", "q": 75, "basis_label": "FY2027 EPS（A）", "basis": f"=C{VAL['A']}"}]}],
    "#,##0.00", "規則（Andy 2026-10-13）：合理倍數＝2016–2025 年底 P/E 依循環位置取分位。採用一個來源（年底收盤 ÷ 當年 EPS，Derived，10 年）；stockanalysis、HiStock 只有 2021–2025，與採用值同量級，只作核對。2016–2022 年京元電以手機、消費性晶片測試為主（約 6–16x），AI 前後倍數口徑不同；2025 年 EPS 含處分利益（每股約 2 元）使當年 P/E 偏低。兩者都使規則值可能低估 AI 期的合理倍數。")
rule_rows = [rw for rw in range(RB0, block_end) if str(V.cell(rw, 1).value or "").startswith("規則值：")]
assert len(rule_rows) == 3, rule_rows
for nm, rw in zip(("pe_peak", "pe_mid", "pe_trough"), rule_rows):
    defname(nm, f"Val!$C${rw}")
yr_rows = [rw for rw in range(RB0, block_end) if V.cell(rw, 1).value in [str(y) for y in range(2016, 2026)]]
assert len(yr_rows) == 10 and yr_rows[-1] - yr_rows[0] == 9, yr_rows
PE_RNG = f"$B${yr_rows[0]}:$B${yr_rows[-1]}"
defname("pe_hist", f"Val!{PE_RNG}")
V.cell(block_end, 1, "本版未做先驗：無 Andy 合理 P/E 對照。")
# 第 0 節（目標價＝DCF 與 Comparable 加權）在所有工作表建好後寫入（見檔尾 v0.2 節）；v0.1 的短期法／長期法已移除
V.column_dimensions["A"].width = 58
for c in "BCDEFGH": V.column_dimensions[c].width = 18

# ---------------- Sens ----------------
SE = wb.create_sheet("Sens")
SE["A1"] = "Sens — 一次動一個輸入（情境 A；每加權顆營收、固定現金成本、κ 在每個案例內依 FY2026 錨點重算，FY2026 恆等於錨點）：對 FY2027、FY2028 的營收、AI 測試營收、毛利率與 EPS 的影響。每個案例是一個完整推導區塊（本頁下方）"
hdr(SE, 3, ["案例", "設定", "FY2027 營收", "FY2027 AI 測試", "FY2027 毛利率", "FY2027 EPS", "FY2027 EPS 相對基準", "FY2028 營收", "FY2028 AI 測試", "FY2028 毛利率", "FY2028 EPS", "FY2028 EPS 相對基準"])
def alias(new, base, suf, years=(27, 28, 29), sep="_"):
    for y in years:
        wb.defined_names[f"{new}_{y}"] = DefinedName(f"{new}_{y}", attr_text=wb.defined_names[f"{base}{sep}{y}{suf}"].attr_text)
for nm, base, suf in (("RLO", "rel", "_lo"), ("RHI", "rel", "_hi"), ("PPLO", "pp", "_lo"), ("PPHI", "pp", "_hi")):
    alias(nm, base, suf)
for nm, suf in (("GNLO", "_lo"), ("GNHI", "_hi")):
    for y in (27, 28, 29):
        wb.defined_names[f"{nm}{y}"] = DefinedName(f"{nm}{y}", attr_text=wb.defined_names[f"g_nai{y}{suf}"].attr_text)
cases = [("測試時間係數 低", "全部平台取低值（Rubin 1.2、Rubin Ultra 1.5、TPU 0.7）", dict(tt="_lo")),
         ("測試時間係數 高", "全部平台取高值（Rubin 1.8、Rubin Ultra 2.6、TPU 1.2）", dict(tt="_hi")),
         ("ASIC 曝險 低", "AMD 0.1、TPU 0.3、Trainium 0.2、其他 0.2", dict(ex="_lo")),
         ("ASIC 曝險 高", "AMD 0.4、TPU 0.7、Trainium 0.6、其他 0.6", dict(ex="_hi")),
         ("份額流失（低）", "份額相對值 2027 0.88、2028 0.80、2029 0.75（日月光、Amkor 擴 GPU 測試）", dict(rel="RLO")),
         ("份額持平（高）", "份額相對值 2027–29 1.00", dict(rel="RHI")),
         ("測試單價下滑快", "單價指數 2027 0.95、2028 0.90、2029 0.88", dict(pp="PPLO")),
         ("測試單價持平", "單價指數 2027 1.02、2028–29 1.00", dict(pp="PPHI")),
         ("AI 部署 ×0.8", "FY2027 起加速器顆數 ×0.8", dict(mf="mf_lo")), ("AI 部署 ×1.2", "×1.2", dict(mf="mf_hi")),
         ("FY2026 AI 占比 28%", "校準用 AI 占比取低值（非 AI 基期較大）", dict(ais="ai_26_lo")), ("FY2026 AI 占比 45%", "取高值", dict(ais="ai_26_hi")),
         ("變動現金成本 10%", "v＝10%（固定成本占比較高）", dict(v="v_cc_lo")), ("變動現金成本 30%", "v＝30%", dict(v="v_cc_hi")),
         ("折舊率 9%", "新增資本支出年折舊率 9%", dict(dr="dr_lo")), ("折舊率 14%", "年折舊率 14%（設備占比較高）", dict(dr="dr_hi")),
         ("維持性資本支出 6%", "營收的 6%", dict(m="m_cx_lo")), ("維持性資本支出 15%", "營收的 15%", dict(m="m_cx_hi")),
         ("營業費用彈性 0.4", "", dict(phi="phi_lo")), ("營業費用彈性 1.0", "營業費用與營收同步成長", dict(phi="phi_hi")),
         ("淨利 ÷ 營業利益 0.78", "利息費用上升", dict(nir="nir_f_lo")), ("淨利 ÷ 營業利益 0.90", "回到 2025 年水準", dict(nir="nir_f_hi")),
         ("非 AI 成長 低", "2027 +4%、2028–29 0%", dict(gn="GNLO")), ("非 AI 成長 高", "2027 +20%、2028 +12%、2029 +10%", dict(gn="GNHI"))]
def put_case(i, name, setting, R, pre):
    SE.cell(i, 1, name); SE.cell(i, 2, setting)
    for c, key, y, fmt in (("C", "rev", 2027, "#,##0.0"), ("D", "ai", 2027, "#,##0.0"), ("E", "gm", 2027, "0.0%"), ("F", "eps", 2027, "0.00"),
                           ("H", "rev", 2028, "#,##0.0"), ("I", "ai", 2028, "#,##0.0"), ("J", "gm", 2028, "0.0%"), ("K", "eps", 2028, "0.00")):
        SE[f"{c}{i}"] = f"={pre}{FC[y]}{R[key]}"; SE[f"{c}{i}"].number_format = fmt
    SE[f"G{i}"] = f"=F{i}/F$4-1"; SE[f"L{i}"] = f"=K{i}/K$4-1"
    for c in "GL": SE[f"{c}{i}"].number_format = "0.0%"
put_case(4, "基準（情境 A）", "PL 區塊 A", PB["A"], "PL!")
base_row = 40; SB = {}
for nm, setting, ov in cases:
    SB[nm] = engine(SE, base_row, f"案例區塊：{nm}（情境 A）", scn="A", **ov)
    base_row = SB[nm]["end"] + 2
for n, (nm, setting, ov) in enumerate(cases):
    put_case(5 + n, nm, setting, SB[nm], "")
nn = 5 + len(cases)
put_case(nn, "顆數情境 B", "CY2028 ×0.8、CY2029 ×0.85", PB["B"], "PL!")
put_case(nn + 1, "顆數情境 C", "CY2028–29 ×0.6", PB["C"], "PL!")
assert nn + 3 < 40
SENS_LAST = nn + 1; SENS_NOTE = nn + 3
SE.cell(nn + 3, 1, "讀法：FY2027 三情境相同；顆數情境 B、C 的差異從 FY2028 起。資本支出隨當年營收增量擴產；FY2028 顆數下修、營收下滑時只剩維持性資本支出，但固定現金成本與既有折舊不降，毛利率跌幅大於營收。測試單價只改營收、不改成本，對 EPS 的影響大於對營收。")
SE.column_dimensions["A"].width = 30; SE.column_dimensions["B"].width = 56
for c in "CDEFGHIJKL": SE.column_dimensions[c].width = 14

# ---------------- Recon ----------------
RC = wb.create_sheet("Recon")
RC["A1"] = "Recon — 本模型 vs 共識、券商估計與公司揭露的結構數字。只作事後對照，不為對上數字而調參數。共識為 FactSet（2026-04-06，經鉅亨網）與媒體轉述的法人估計"
hdr(RC, 3, ["項目", "對照值", "模型 A", "差距 A", "模型 B", "差距 B", "模型 C", "差距 C", "對照來源", "說明"])
rr_ = 4
def rec(label, ref, getter, src, note="", fmt="#,##0.00", pt=False, gap=True):
    global rr_
    RC.cell(rr_, 1, label)
    if ref is not None: RC.cell(rr_, 2, ref).number_format = fmt
    for j, s in enumerate(SC):
        cv, cd = 3 + 2 * j, 4 + 2 * j
        RC.cell(rr_, cv, getter(s)).number_format = fmt
        if ref is not None and gap:
            L = chr(64 + cv)
            RC.cell(rr_, cd, f"={L}{rr_}-B{rr_}" if pt else f"={L}{rr_}/B{rr_}-1").number_format = "0.0%"
    RC.cell(rr_, 9, src); RC.cell(rr_, 10, note); rr_ += 1
pl = lambda key, y: (lambda s: f"=PL!{FC[y]}{PB[s][key]}")
pl05 = lambda key, y: (lambda s: f"=PL!{FC[y]}{PB[s][key]}*1.05")
RC.cell(rr_, 1, "A. 共識與法人估計（校準未使用）").font = BOLD; rr_ += 1
rec("FY2026 營收 vs 法人（近 500 億）", "=c_rev26", pl("rev", 2026), "理財周刊 2026-09-01", "模型 4Q 季增 15%；法人約 20%", fmt="#,##0.0")
rec("4Q26 營收 vs 法人估計", "=c_q4rev", lambda s: "=Q!E4", "優分析 2026-09-18", "", fmt="#,##0.00")
rec("FY2026 EPS vs FactSet 中位數（04-06）", "=c_eps26", pl("eps", 2026), "FactSet（鉅亨網）", "模型以配股後股數計；下一列換回配股前口徑")
rec("FY2026 EPS（× 1.05，配股前口徑）vs FactSet 中位數", "=c_eps26", pl05("eps", 2026), "同上", "共識的股數口徑未載明（推定為配股前）")
rec("FY2026 EPS（× 1.05）vs 法人中位數區間中點（09-01）", "=c_eps26b", pl05("eps", 2026), "理財周刊 2026-09-01", "3Q Rubin 延後後的估計")
rec("FY2027 EPS vs FactSet 中位數（04-06）", "=c_eps27", pl("eps", 2027), "FactSet（鉅亨網）", "")
rec("FY2027 EPS（× 1.05）vs FactSet 中位數", "=c_eps27", pl05("eps", 2027), "同上", "高 16.05、低 13.35")
rec("FY2028 EPS（× 1.05）vs FactSet 中位數", "=c_eps28", pl05("eps", 2028), "同上", "高 18.77、低 16.67")
rec("目標價（DCF 與 Comparable 加權，Val 第 0 節）vs Morgan Stanley 目標價", "=ms_tp", lambda s: "=tp_final", "Morgan Stanley 2026-06-23", "三情境同一個加權目標價；券商目標價多以 25–30 倍前瞻 P/E 計", fmt="#,##0.0")
RC.cell(rr_, 1, "B. 結構與機制檢查").font = BOLD; rr_ += 1
tpu_share = lambda s: f"=PL!D{PB[s]['w_tpu']}/PL!D{PB[s]['w']}*PL!D{PB[s]['ai']}/PL!D{PB[s]['rev']}"
rec("FY2026 Google TPU 測試 ÷ 營收 vs Morgan Stanley Google 相關 10%", "=ms_goog", tpu_share, "Morgan Stanley 2026-06-23", "對照值含 Google CPU、晶圓測試，口徑較寬；模型應低於對照值", fmt="0.0%", pt=True)
rec("FY2026 AI 加速器測試占比 vs 公司 2Q26 資料處理應用占比", "=dp_26", lambda s: f"=PL!D{PB[s]['aish']}", "poorstock 摘要公司簡報", "資料處理含 AI GPU、ASIC 與 CPU；AI 占比應不高於此值", fmt="0.0%", pt=True)
rec("FY2026 折舊公式回測（同一公式 − 錨點，NT$B）", None, lambda s: f"=PL!D{PB[s]['depbt']}", "Q 頁、In 頁 dr、ret", "接近 0 表示折舊率與到期比例和 FY2025–26 一致（校準檢查）", fmt="0.00")
rec("FY2026 資本支出：公式值（維持性＋κ × FY2026 營收增量）vs 輸入 capex_26", "=capex_26",
    lambda s: f"=m_cx*PL!D{PB[s]['rev']}+PL!C{PB[s]['kap']}*MAX(0,PL!D{PB[s]['rev']}-PL!C{PB[s]['rev']})", "In 頁 capex_26（公司預算 500 億的現金口徑）", "公式值接近預算 500 億、高於現金口徑輸入；差距反映預算與現金支出的時間差", fmt="#,##0.0")
rec("FY2027 每顆 Blackwell 等值 AI 測試營收（US$）", None, pl("pue", 2027), "—", "", fmt="#,##0.0")
rec("FY2027 資本支出 ÷ 營收", None, pl("cxr", 2027), "—", "FY2025 92.6%、FY2026 預算約 105%（現金基準約 88%）", fmt="0.0%")
rec("FY2028 折舊攤提 ÷ 營收", None, pl("depr", 2028), "—", "FY2025 21.5%", fmt="0.0%")
rec("FY2028 EBITDA − 資本支出（NT$B）", None, pl("fcf", 2028), "—", "負值＝擴產仍需外部資金", fmt="#,##0.0")
rec("Burn-in 占營收（3Q25，公司法說）", "=bi_25", lambda s: "=bi_25", "富果 2025-11-28（AI 摘要）", "本模型未拆 burn-in；Rubin 世代的 burn-in 加長含在測試時間係數內", fmt="0.0%", gap=False)
RC.cell(rr_, 1, "C. 預留").font = BOLD; rr_ += 1
RC.cell(rr_, 1, "3Q26 財報（約 11 月上旬）：更新 Q 頁 3Q 毛利率、折舊、營業費用與淨利，並檢查 4Q 季增假設"); RC.cell(rr_, 10, "預留列：本版無數值"); rr_ += 1
RC.cell(rr_, 1, "R4 產業部署路徑（markets/deploy_path/，尚未建立）：上線後以部署路徑取代顆數輸入"); RC.cell(rr_, 10, "預留列：本版無數值"); rr_ += 1
RECON_LAST = rr_ - 1
RC.column_dimensions["A"].width = 70
for c in "BCDEFGH": RC.column_dimensions[c].width = 12
RC.column_dimensions["I"].width = 36; RC.column_dimensions["J"].width = 60

# ---------------- Sources ----------------
SO = wb.create_sheet("Sources"); SO["A1"] = "來源（公司原文＝京元電每季 Investor update 與月營收；AI 占比、客戶與券商估計皆為 Interested-party）"; SO["A1"].font = BOLD
srcs = [("公司：2Q26 Investor update（1Q26、2Q26 營收、毛利率、費用、EBITDA、EPS；安裝測試機 5,900 台以上）", DECK2Q26),
        ("公司：4Q25 Investor update（3Q25、4Q25 營收、毛利率、費用、EBITDA、EPS）", DECK4Q25),
        ("公司：2Q25 Investor update（1Q25、2Q25；不含已出售的中國子公司）", DECK2Q25),
        ("中央社 2026-03-06（FY2025 營收、毛利率、營業利益、淨利、EPS；配股；處分利益）", CNA0306),
        ("科技新報 2026-05-11（1Q26 毛利率 39.74%、營益率 27.2%、淨利）", TN0511),
        ("月營收：科技新報 2026-10-08（9 月 40.71 億、前 9 月 334.76 億）", TN1008),
        ("月營收：HiStock（2025-01 至 2026-09）", HSM),
        ("媒體：理財周刊 2026-09-01（2Q26 淨利、EPS；法人 2026 年營收與 EPS；目標價中位數）", MW0901),
        ("媒體：理財周刊 2026-09-24（資本支出 500 億、美國廠 14 億美元、Rubin 延後、ASIC 客戶）", MW0924),
        ("媒體：優分析 2026-09-18（4Q26 營收估計、Rubin 影響、聯發科 TPU、博通占比）", UA0918),
        ("券商：Morgan Stanley（經科技新報 2026-06-23；AI 占比、Google 相關、CPU、burn-in、資本支出、目標價）", MS0623),
        ("媒體：雅虎新聞 2026-01-27（4Q25 營收、2025 年資本支出約 370 億、2026 年預算 393.72 億）", YAM0127),
        ("媒體：工商時報 2026-07-28（配股、美國投資 14 億美元、資本支出上修至 500 億、1H26 營收）", CT0728),
        ("資料：poorstock 法說會摘要（AI 摘要公司簡報：季資本支出、2023–2025 年資本支出、稼動率、資料處理占比）", PS),
        ("公司法說摘要：富果 2025-11-28（製程與應用別營收結構、burn-in 占比、GPU 與 ASIC）", FG1128),
        ("共識：FactSet（經鉅亨網 2026-04-06；2026–2028 EPS 中位數、目標價）", FS0406),
        ("資料商：鉅亨網、財報狗（10-22 收盤、近 12 月 P/E、股本）", CNYQ + "；" + SD),
        ("資料商：Yahoo Finance 月線收盤（2016–2025 年 12 月；已因 2026 年配股調整）", "https://finance.yahoo.com/quote/2449.TW/history/"),
        ("資料商：玩股網年度 EPS（2013–2025）；HiStock 季 EPS 與月本益比；stockanalysis 年度 P/E", "https://www.wantgoo.com/stock/2449/profitability/dividend-payout-ratio；https://histock.tw/stock/2449/%E6%9C%AC%E7%9B%8A%E6%AF%94；https://stockanalysis.com/quote/tpe/2449/financials/ratios/"),
        ("市場層：tsmc 成品（加速器顆數與情境係數，本 repo，唯讀）", ""),
        ("待補（人工）：3Q26 財報（約 11 月上旬）；10 月營收（11 月上旬）；折舊攤提金額（財報附註）；客戶別營收與 NVIDIA 委外份額；測試單價；處分利益金額；前瞻 P/E 十年歷史（Bloomberg）；FY2027–28 最新共識（付費）", "未找到"),
        ("SemiAnalysis", "未使用")]
for i, (a, b) in enumerate(srcs, 2): SO.cell(i, 1, a); SO.cell(i, 2, b)
SO.column_dimensions["A"].width = 90; SO.column_dimensions["B"].width = 140

# ---------------- README ----------------
RD = wb.create_sheet("README", 0)
lines = [("京元電子（2449）收支與估值模型 v0.2.1（AI 半導體 Project：測試代工；AI GPU／ASIC 最終測試與 burn-in）　2026-10-25", True),
 ("版本：v0.2.1（2026-10-25：補上檔空間級距、投資建議、重點觀察，數值不變）；v0.2（2026-10-25：目標價改為 DCF 與 Comparable 加權，取代 v0.1 的短期法／長期法）。驅動邏輯比照 Advantest 模型（測試需求＝加速器顆數 × 每顆測試時間；顆數與情境係數讀台積電模型成品）；台股寫法比照台達電、台光電模型。", False),
 ("模型命題：京元電的 AI 測試營收＝加速器顆數（NVIDIA、ASIC）× 每顆測試時間（含 burn-in；Rubin 世代加長）× 京元電曝險 × 每加權顆營收（FY2026 校準）× 份額相對值 × 測試單價；非 AI（伺服器 CPU、手機、消費、車用）另依成長率推。毛利率由固定現金成本、變動現金成本與折舊推出，折舊由隨當年營收增量擴產的資本支出累積。股價約 32 倍 FY2026、24 倍 FY2027 EPS（模型）。", False),
 ("核心問題：2026 年 500 億的擴產（約為營收的一倍）能否在 2027–28 年被 AI 測試量吸收；顆數放緩或份額流失時，累積的折舊會讓毛利率掉多少；現價隱含的路徑合不合理。", False), ("", False),
 ("會計年度＝日曆年，FY(t)＝CY(t)，不平移。預測期 FY2027–FY2029 由驅動因子推導；FY2030 起為收斂法，非驅動因子推導（營收年增與淨利率線性收斂，延伸到 FY2032；PL 頁綠底）。金額：十億新台幣。口徑：歸屬母公司淨利；FY2026 起以盈餘配股 5% 後的股數計 EPS。", False), ("", False),
 ("驅動因子 → 推導量（全部唯讀上游，不修改）", True),
 ("需求面：tsmc 成品的加速器顆數（八個平台）× 情境顆數係數 × 每顆測試時間係數（Assumed／Analogy，比照 Advantest 模型並加 burn-in）× 京元電曝險（ASIC、AMD 相對 NVIDIA，Assumed）＝加權顆數；× 每加權顆營收（FY2026 以 AI 占比 35% 校準，Derived）× 份額相對值 × 測試單價指數（Assumed）＝AI 加速器測試營收。", False),
 ("供給面：資本支出＝維持性比例 × 營收＋κ × 當年營收增量（κ 以 FY2025 現金資本支出與 FY2025 營收增量校準）；FY2026 取公司預算 500 億的現金口徑 420 億（Derived）。折舊＝前一年 ×（1 − 到期）＋折舊率 × 資本支出平均。", False),
 ("毛利率：營收 − 現金營業成本（固定 F，FY2026 校準；變動 v × 測試量）− 折舊。營業費用：彈性 phi。淨利 ÷ 營業利益 0.84（Assumed：利息費用上升）。", False),
 ("校準：FY2026 ＝ 1H26 實績＋3Q26 營收實績（月營收）＋假設（4Q 季增 15%、毛利率 38.5％／39.5%、折舊 29.5／31.5 億、營業費用率 12.5%）。3Q26 財報約 11 月上旬公布。不用共識數字（只在 Recon 對帳）。", False),
 ("估值：現價 ÷ FY2027、FY2028 EPS；十年 P/E 規則值（年底收盤 ÷ 當年 EPS，10 年）× 各情境 FY2027 EPS；循環位置以本模型營收最高年判定（京元電不套設備股見頂年）。目標價（Val 第 0 節）＝0.5 × DCF（股東現金流＝EPS ×（1 − g／ROE），FY2027–32 明確期＋永續，r 10%、g 4%，情境權重 0.3／0.5／0.2）＋0.5 × Comparable（台美封測同業前瞻 P/E 中位數 × 12 個月後前瞻 EPS）；十年 P/E 規則值保留參考。本版未做先驗，無 Andy 合理 P/E 對照。", False),
 ("標記：Verified、Interested-party、Analogy、Assumed、Derived。SemiAnalysis 未使用。公司原文＝每季 Investor update 與月營收；AI 占比、客戶、資本支出預算多經媒體與券商轉述。", False),
 ("", False),
 ("已知限制：(1) 公司不揭露 AI 營收、客戶別營收與 NVIDIA 委外份額，AI 占比取 Morgan Stanley 的 35%（媒體另有 70% 的寬口徑說法）；份額只能以相對值表示；(2) 每顆測試時間（含 burn-in）係數、ASIC 曝險、份額與單價路徑皆為 Assumed；(3) 折舊攤提金額公司未直接揭露，以 EBITDA − 營業利益推得，並假設全部在營業成本；(4) 資本支出以營收增量連結，2026 年預算含廠房與美國廠的提前投資，FY2027 起的資本支出可能低估、毛利率可能高估；(5) 十年 P/E 只有一個採用來源，2016–2022 年為 AI 前低倍數，2025 年 EPS 含處分利益；(6) 3Q26 財報（約 11 月上旬）、10 月營收後更新錨點；R4 產業部署路徑未建。", False)]
lines.append(("v0.2（2026-10-25）：Val 頁第 0 節由短期法／長期法改為「0. 目標價」（0.1 DCF、0.2 Comparable、0.3 結論），第 1 節起整體下移 23 列；In 頁 tp_mix 改為 tp_w_dcf，新增共識目標價與共識前瞻 EPS；Recon 的 Morgan Stanley 對照改用加權目標價；Sens 加目標價參數列。PL、FY2026–FY2032 EPS 不變。", False))
for i, (t, b) in enumerate(lines, 1): RD.cell(i, 1, t).font = Font(bold=b)
RD.column_dimensions["A"].width = 170


# ================= v0.2（2026-10-25）：FY2030–FY2032 收斂法；目標價＝DCF 與 Comparable 加權 =================
# 工作單：Andy 2026-10-25 11:18／11:20（三情境延伸到 FY2032、收斂法）、11:28（目標價＝DCF＋Comparable）、11:53（tp_w_dcf 0.5；A 股與其他市場同業分開）
# 本節在其他工作表都建好之後執行：(1) PL 說明欄由 H 移到 K，H–J 放 FY2030–32 收斂年度；(2) Val 既有內容整體下移，最前面放「0. 目標價」；(3) Sens 加目標價參數列
import re as _re
from copy import copy as _copy
FADE = (2030, 2031, 2032); FADEFILL = PatternFill("solid", fgColor="E7E6E6")
FCX = dict(FC); FCX.update({2030: "H", 2031: "I", 2032: "J"})
YRS6 = (2027, 2028, 2029, 2030, 2031, 2032)
INR = {IN.cell(i, 1).value: i for i in range(4, IN.max_row + 1) if IN.cell(i, 1).value}
for _k, _i in INR.items():
    if str(_k).startswith(("tp_", "fade_")) or _k in ("cons_tp", "cons_feps"):
        for _j in range(1, 10): IN.cell(_i, _j).fill = KEY

def _move_notes(ws):
    """說明欄（第 8 欄，H）移到第 11 欄（K），讓出 H–J；表頭列補上收斂年度。"""
    for r in range(1, ws.max_row + 1):
        c = ws.cell(r, 8)
        if c.value is None: continue
        assert not (isinstance(c.value, str) and c.value.startswith("=")), (ws.title, r, c.value)
        k = ws.cell(r, 11); k.value = c.value; k.font = _copy(c.font); k.fill = _copy(c.fill)
        c.value = None; c.font = Font(); c.fill = PatternFill()
        if str(ws.cell(r, 1).value or "") in ("項目", "情境／項目"):
            for j, y in zip((8, 9, 10), FADE):
                h = ws.cell(r, j, f"FY{y}（收斂）"); h.font = BOLD; h.fill = HDR
    ws.column_dimensions["K"].width = ws.column_dimensions["H"].width or 70
    for col in "HIJ": ws.column_dimensions[col].width = 14

def _fade(ws, R, s, eps_f):
    """FY2030–32：營收年增由 FY2029 線性收斂到 fade_g、淨利率收斂到 fade_m（FY2032 達到）；非驅動因子推導。"""
    for y in FADE:
        c, pc, k = FCX[y], FCX[y - 1], y - 2029
        fs = {"g": f"=$G{R['g']}+(fade_g{s}-$G{R['g']})*{k}/3", "rev": f"={pc}{R['rev']}*(1+{c}{R['g']})",
              "nim": f"=$G{R['nim']}+(fade_m{s}-$G{R['nim']})*{k}/3", "ni": f"={c}{R['rev']}*{c}{R['nim']}",
              "sh_": f"={pc}{R['sh_']}*(1+sh_g)", "eps": "=" + eps_f(c, R), "epsg": f"={c}{R['eps']}/{pc}{R['eps']}-1"}
        for key, f in fs.items():
            ws[f"{c}{R[key]}"] = f; ws[f"{c}{R[key]}"].number_format = ws[f"G{R[key]}"].number_format; ws[f"{c}{R[key]}"].fill = FADEFILL

def _shift_rows(wb, sheet, start, n):
    """把 sheet 第 start 列起的內容整體下移 n 列，並同步改寫所有公式與定義名稱中指向這些列的參照（等同插入列）。"""
    ws = wb[sheet]
    cell_ = r"\$?[A-Z]{1,3}\$?\d+"
    pat = _re.compile(r"(?<![A-Za-z0-9_.!$:])((?:'[^']+'|[A-Za-z_][A-Za-z0-9_]*)!)?(" + cell_ + r")(?::(" + cell_ + r"))?(?![A-Za-z0-9_(!])")
    one = _re.compile(r"(\$?[A-Z]{1,3}\$?)(\d+)")
    def fix(f, own):
        parts = f.split('"')
        for i in range(0, len(parts), 2):
            def rep(m):
                pre = m.group(1)
                tgt = (pre is None and own) or (pre is not None and pre.strip("'!") == sheet)
                def sh(x):
                    a, rr = one.fullmatch(x).groups()
                    rr = int(rr)
                    return f"{a}{rr + n if (tgt and rr >= start) else rr}"
                return (pre or "") + sh(m.group(2)) + (":" + sh(m.group(3)) if m.group(3) else "")
            parts[i] = pat.sub(rep, parts[i])
        return '"'.join(parts)
    for w in wb.worksheets:
        for row in w.iter_rows():
            for c in row:
                if isinstance(c.value, str) and c.value.startswith("="):
                    c.value = fix(c.value, w.title == sheet)
    for nm in list(wb.defined_names.keys()):
        d = wb.defined_names[nm]
        if sheet + "!" in d.attr_text: d.attr_text = fix(d.attr_text, False)
    cells = [(c.row, c.column) for row in ws.iter_rows(min_row=start) for c in row if c.value is not None or c.has_style]
    for r, col in sorted(cells, reverse=True):
        src = ws.cell(r, col); dst = ws.cell(r + n, col)
        dst.value = src.value; dst._style = _copy(src._style); dst.number_format = src.number_format
        src.value = None; src._style = _copy(ws.cell(1, 30)._style)

def _tp_section(V, eps_ref, fwd, fwd_note, peers, peer_note, cons_note, method, pe_rng, fmt, cons_feps_note):
    """Val 第 0 節（第 4–48 列）：0.1 DCF、0.2 Comparable、0.3 結論。列標籤依工作單逐字；B／C／D＝情境 A 順利／B 平平／C 不如預期。"""
    T = {}
    def tset(r, key, label, fs, nf, note=""):
        V.cell(r, 1, label); T[key] = r
        for c, f in zip("BCD", fs):
            if f is None: continue
            V[f"{c}{r}"] = f; V[f"{c}{r}"].number_format = nf
        if note: V[f"E{r}"] = note
    V.cell(4, 1, "0. 目標價").font = BOLD
    V.cell(5, 1, "0.1 DCF（股東權益現金流，FY2027–FY2032 明確期＋FY2032 後永續）").font = BOLD
    hdr(V, 6, ["項目", "A（順利）", "B（平平）", "C（不如預期）", "說明"])
    V.cell(7, 1, "情境 EPS（FY2027–FY2032）"); V["E7"] = "PL 頁各情境；FY2030 起為收斂法（PL 頁灰底）"
    for i, y in enumerate(YRS6):
        tset(8 + i, f"eps{y}", f"  FY{y}", [f"={eps_ref(s, y)}" for s in SC], "0.00")
    V.cell(14, 1, "股東現金流＝EPS ×（1 − tp_g ÷ tp_roe）")
    tset(15, "reinv", "  再投資率（tp_g ÷ tp_roe，上限 0.8）", ["=IF(tp_roe<=tp_g,0.8,MIN(tp_g/tp_roe,0.8))", None, None], "0.0%", "ROE＝In 頁 tp_roe；tp_roe ≤ tp_g 時取上限 0.8")
    for i, y in enumerate(YRS6):
        tset(16 + i, f"cf{y}", f"  FY{y}", [f"={c}{8 + i}*(1-$B$15)" for c in "BCD"], "0.00")
    tset(22, "dcfv", "DCF 情境每股價值（折現率 tp_r）", [f"=SUMPRODUCT({c}16:{c}21,1/(1+tp_r)^{{1;2;3;4;5;6}})+{c}21*(1+tp_g)/(tp_r-tp_g)/(1+tp_r)^6" for c in "BCD"], fmt,
         "Σ 明確期現金流 ÷ (1＋r)^t ＋ 永續價值 ÷ (1＋r)^6；永續價值＝FY2032 現金流 ×(1＋g)÷(r−g)；基準日 2026 年底，t＝1 對 FY2027")
    tset(23, "dcf12", "  DCF 情境價值推到 12 個月後（×(1＋tp_r)）", [f"={c}22*(1+tp_r)" for c in "BCD"], fmt)
    tset(24, "dcftp", "DCF 目標價（情境加權）", ["=tp_wA*B23+tp_wB*C23+tp_wC*D23", None, None], fmt, "Σ 權重 × 情境價值 ×(1＋r)；權重見 In 頁")
    V.cell(26, 1, "0.2 Comparable（市場主流同業前瞻 P/E）").font = BOLD
    hdr(V, 27, ["同業與前瞻 P/E", "代號", "前瞻 12 月 P/E", "日期", "來源", "標記"])
    assert 3 <= len(peers) <= 5, peers
    for i, (nm, tk, pe, dt, src) in enumerate(peers, 28):
        V.cell(i, 1, nm); V.cell(i, 2, tk); c = V.cell(i, 3, pe); c.font = BLUE; c.number_format = "0.0"
        V.cell(i, 4, dt); V.cell(i, 5, src); V.cell(i, 6, "Interested-party")
    tset(33, "pmed", "同業前瞻 P/E 中位數", [f"=MEDIAN(C28:C{27 + len(peers)})", None, None], "0.0", peer_note)
    tset(34, "feps", "12 個月後的前瞻 EPS（情境加權）",
         ["=" + "+".join(f"tp_w{s}*(" + "+".join(f"{w}*{c}{8 + YRS6.index(y)}" for y, w in fwd) + ")" for s, c in zip(SC, "BCD")), None, None], "0.00", fwd_note)
    tset(35, "comptp", "Comparable 目標價", ["=B33*B34", None, None], fmt, "同業中位數 × 前瞻 EPS")
    V.cell(37, 1, "0.3 結論").font = BOLD
    hdr(V, 38, ["項目", "數值", "", "", "說明"])
    tset(39, "tpw", "目標價（DCF 與 Comparable 加權）", ["=tp_w_dcf*B24+(1-tp_w_dcf)*B35", None, None], fmt, "tp_w_dcf × DCF＋（1 − tp_w_dcf）× Comparable；tp_w_dcf 見 In 頁")
    tset(40, "comp2", "Comparable 目標價（市場主流同業法）", ["=B35", None, None], fmt)
    tset(41, "cons", "市場共識目標價", ["=cons_tp", None, None], fmt, cons_note)
    tset(42, "upw", "上檔空間：加權目標價", ["=B39/px-1", None, None], "0.0%", "現價＝In 頁 px")
    tset(43, "upc", "上檔空間：Comparable 目標價", ["=B40/px-1", None, None], "0.0%")
    tset(44, "geps", "與市場共識目標價的差距：EPS 因素", ["=B34/cons_feps-1", None, None], "0.0%", cons_feps_note)
    tset(45, "gmul", "與市場共識目標價的差距：倍數因素", ["=(cons_tp/cons_feps)/B33-1", None, None], "0.0%")
    V["E45"] = '="共識目標價隱含 "&TEXT(cons_tp/cons_feps,"0.0")&"x，同業中位數 "&TEXT(B33,"0.0")&"x"'
    tset(46, "band", "價格帶（DCF 不如預期～順利）", [f'=TEXT(D23,"{fmt}")&"～"&TEXT(B23,"{fmt}")', None, None], "@", "C 情境價值～A 情境價值（推到 12 個月後）")
    tset(47, "pr", "歷史倍數檢查：現價隱含 P/E 的十年分位（A／B／C）",
         [f"=IF(px/{c}9>MAX({pe_rng}),1,IF(px/{c}9<MIN({pe_rng}),0,PERCENTRANK({pe_rng},px/{c}9)))" for c in "BCD"], "0%",
         "現價 ÷ 各情境 FY2028 EPS 在 2016–2025 年底 P/E（第 6 節採用來源）的分位；高於十年最大值＝100%；只作檢查，不進目標價")
    V.cell(48, 1, "目標價方法"); V["B48"] = method
    for r_ in (24, 35, 39):
        V[f"B{r_}"].fill = KEY
    return T

def _tp_sens(SE, T):
    """Sens：目標價參數（tp_r ±1pt、tp_g ±1pt、tp_w_dcf 0.3／0.7）；放在案例摘要與第 40 列推導區塊之間。"""
    last = max(r for r in range(1, 40) if any(SE.cell(r, j).value is not None for j in range(1, 13)))
    t0 = last + 2
    if t0 + 8 >= 40:          # 案例摘要與第 40 列推導區塊之間放不下：放在全部推導區塊之後
        t0 = SE.max_row + 3
    SE.cell(t0, 1, "目標價參數（只改 DCF 參數與權重；PL 收斂法的 fade_g 不隨 tp_g 變動）").font = BOLD
    hdr(SE, t0 + 1, ["案例", "設定", "DCF 目標價", "Comparable 目標價", "加權目標價", "相對基準"])
    def dcf_f(c, r, g, reinv):
        cf = lambda i: f"Val!{c}{8 + i}*(1-{reinv})"
        pv = "+".join(f"{cf(i)}/(1+{r})^{i + 1}" for i in range(6))
        return f"({pv}+{cf(5)}*(1+{g})/({r}-{g})/(1+{r})^6)"
    cases = [("基準", "r 10%、g 4%、DCF 權重 0.5", {}), ("折現率 +1pt", "r 11%", dict(r="(tp_r+0.01)")), ("折現率 −1pt", "r 9%", dict(r="(tp_r-0.01)")),
             ("長期成長 +1pt", "g 5%", dict(g="(tp_g+0.01)")), ("長期成長 −1pt", "g 3%", dict(g="(tp_g-0.01)")),
             ("DCF 權重 0.3", "Comparable 0.7", dict(w="0.3")), ("DCF 權重 0.7", "Comparable 0.3", dict(w="0.7"))]
    for k, (nm, st, kw) in enumerate(cases):
        i = t0 + 2 + k; r = kw.get("r", "tp_r"); g = kw.get("g", "tp_g"); w = kw.get("w", "tp_w_dcf")
        reinv = f"IF(tp_roe<={g},0.8,MIN({g}/tp_roe,0.8))"
        SE.cell(i, 1, nm); SE.cell(i, 2, st)
        SE[f"C{i}"] = "=(" + "+".join(f"tp_w{s}*{dcf_f(c, r, g, reinv)}" for s, c in zip(SC, "BCD")) + f")*(1+{r})"
        SE[f"D{i}"] = f"=Val!B{T['comptp']}"
        SE[f"E{i}"] = f"={w}*C{i}+(1-{w})*D{i}"
        SE[f"F{i}"] = f"=E{i}/E${t0 + 2}-1"
        for c in "CDE": SE[f"{c}{i}"].number_format = "0.00"
        SE[f"F{i}"].number_format = "0.0%"
    assert t0 + 1 + len(cases) < 40 or t0 > 40, t0
    return t0, t0 + 1 + len(cases)

def _fade_all(eps_f, nim10):
    """PL：三情境推導區塊與頂端摘要延伸到 FY2032；In 頁 fade_m 連到 PL 頁 FY2029 淨利率。nim10＝十年平均淨利率的名稱；None＝未找到，依工作單用 FY2029 × 0.8。"""
    _move_notes(PL)
    for s in SC:
        _fade(PL, PB[s], s, eps_f)
        for key in ("rev", "g", "nim", "ni", "eps", "epsg"):
            if (s, key) not in SUMR: continue
            for y in FADE:
                cc = PL[f"{FCX[y]}{SUMR[(s, key)]}"]; cc.value = f"={FCX[y]}{PB[s][key]}"
                cc.number_format = PL[f"G{SUMR[(s, key)]}"].number_format; cc.fill = FADEFILL
    nimr = {s: f"PL!$G${PB[s]['nim']}" for s in SC}
    IN.cell(INR["fade_mA"], 5, f"={nimr['A']}")
    IN.cell(INR["fade_mB"], 5, f"=({nimr['B']}+{nim10})/2" if nim10 else f"=({nimr['B']}+{nimr['B']}*0.8)/2")
    IN.cell(INR["fade_mC"], 5, f"={nim10}" if nim10 else f"={nimr['C']}*0.8")
    PL["A2"] = ("灰底＝FY2030–FY2032 收斂法，非驅動因子推導（上游市場層只到 2029）：營收年增由 FY2029 線性收斂到 fade_g（A＝tp_g＋2pt、B＝tp_g、C＝tp_g−2pt），"
                "淨利率由 FY2029 線性收斂到 fade_m（In 頁 K 節），FY2032 達到；股數依 sh_g")

def _rule_pe_range(V, ncols=1):
    """第 6 節十年倍數表（2016–2025 列）中採用來源的欄位範圍（PERCENTRANK 用；ncols＞1 時各來源合併成一個分布）。"""
    yr = [r for r in range(1, V.max_row + 1) if V.cell(r, 1).value in [str(y) for y in range(2016, 2026)]]
    yr = [r for r in yr if r > 50][:10]
    assert len(yr) == 10 and yr[-1] - yr[0] == 9, yr
    return f"$B${yr[0]}:${chr(65 + ncols)}${yr[-1]}"

_shift_rows(wb, "Val", 28, 23)
TPT = _tp_section(V, lambda s_, y: f"PL!{FCX[y]}{SUMR[(s_, 'eps')]}", [(2027, "2/12"), (2028, "10/12")],
    "約 2027-10 時的未來 12 月 EPS ≈ 2／12 × FY2027＋10／12 × FY2028（會計年度＝日曆年；配股後股數）",
    [("日月光投控（ASE Technology）", "ASX／3711.TW", 27.37, "2026-10-23", "stockanalysis forward PE（ASX；分析師共識）"),
     ("Amkor Technology", "AMKR", 14.02, "2026-10-23", "stockanalysis forward PE（分析師共識）"),
     ("力成（Powertech）", "6239.TW", 20.60, "2026-10-02", "stockanalysis forward PE（頁面價格日期較舊）"),
     ("矽格（Sigurd）", "6257.TW", 23.30, "2026-10-02", "stockanalysis forward PE（頁面價格日期較舊）"),
     ("南茂（ChipMOS）", "IMOS／8150.TW", 22.08, "2026-10-08", "stockanalysis forward PE（IMOS；頁面價格日期較舊）")],
    "台美封測同業（不放 A 股：長電、通富等 A 股同業有結構性溢價；工作單 Andy 11:53）",
    "stockanalysis：15 位分析師平均 340.73（頁面價格日期 2026-10-06）；法人中位數 348–350、Morgan Stanley 388",
    ("DCF：股東現金流＝EPS ×（1 − g／ROE），r＝10%、g＝4%、ROE 為近三年平均約 18%（玩股網；2025 含處分利益，偏高，再投資率約 22%），FY2027–32 明確期＋永續，FY2030 起為收斂法。"
     "京元電為資本密集的測試代工，模型有資本支出與折舊，但 FY2030 起只有收斂法、未延伸資本支出，未改用淨利＋折舊－資本支出；2026–27 年大幅擴產期的現金流偏高估。"
     "Comparable：台美封測同業（日月光、Amkor、力成、矽格、南茂）前瞻 P/E 中位數；京元電以 AI 測試為主、成長較高，同業中位數偏低估。"
     "情境對應：顆數情境 A＝順利、B＝平平、C＝不如預期；FY2027 三情境 EPS 相同（差異從 CY2028 起）。"),
    _rule_pe_range(V, 1), "#,##0.0",
    "本模型前瞻 EPS ÷ 共識前瞻 EPS − 1；共識前瞻 EPS＝In 頁 cons_feps（FactSet 2026-04-06，換成配股後股數）")
defname("tp_final", f"Val!$B${TPT['tpw']}")
TPS0, TPS_LAST = _tp_sens(SE, TPT)

# ---------------- v0.2.1（2026-10-25）：上檔空間級距、投資建議、重點觀察（Andy 14:00、14:17；既有數值與儲存格位置不變） ----------------
_UPW = [r_ for r_ in range(1, 60) if V.cell(r_, 1).value == "上檔空間：加權目標價"]
assert _UPW == [42], _UPW
GR_ROW, RC_ROW = 49, 50   # 0.3 結論後的空白預留列（不移動其他列）
for _r in (GR_ROW, RC_ROW):
    assert all(V.cell(_r, _c).value is None for _c in range(1, V.max_column + 1)), _r
V.cell(GR_ROW, 1, "上檔空間級距")
V.cell(GR_ROW, 2, '=IF(B42>0.5,"大於 50%",IF(B42>0.25,"25%–50%",IF(B42>=0,"0%–25%",IF(B42>=-0.25,"0%～−25%",IF(B42>=-0.5,"−25%～−50%","小於 −50%")))))').number_format = "@"
V.cell(GR_ROW, 5, "以「上檔空間：加權目標價」分六級（Andy 2026-10-25 14:00）")
V.cell(RC_ROW, 1, "投資建議")
V.cell(RC_ROW, 2, '=IF(B42>0.5,"強力買進",IF(B42>0.25,"買進",IF(B42>=-0.25,"中立",IF(B42>=-0.5,"放空","強力放空"))))').number_format = "@"
V.cell(RC_ROW, 5, "依上檔空間固定規則（Andy 2026-10-25 14:17）：大於 +50% 強力買進、+25%～+50% 買進、−25%～+25% 中立、−50%～−25% 放空、小於 −50% 強力放空")
OBS = [
       "目標價 NT$173（DCF 139、Comparable 208 各半），比現價 262.5（2026-10-22）低 34.0%，級距 −25%～−50%；Comparable 單獨為 −20.7%。主因是 EPS 不是倍數：本模型 12 個月後前瞻 EPS 9.43 比共識前瞻 EPS 16.35 低 42.3%（FactSet 2026-04-06，換成配股後股數），同業前瞻 P/E 中位數 22.1 倍與共識目標價 340.73 隱含的 20.8 倍相近；EPS 偏低是因 500 億擴產的折舊讓情境 A 毛利率由 FY2027 40.1% 降到 FY2028 37.6%，FY2028 EPS 11.10 不成長。",
       "最敏感的是折舊與固定成本：新增資本支出折舊率 14% 時 FY2028 EPS −14.4%、非 AI 成長低 −12.2%、變動現金成本 30% −10.9%；情境 C FY2028 EPS 6.17，比情境 A 低 44.4%。現價 ÷ FY2028 EPS 23.7～42.5 倍，在十年分位 96%～100%。",
       "追蹤 3Q26 財報與 10 月營收（皆約 11 月上旬）：折舊攤提金額公司未揭露（財報附註待補），AI 占比取 Morgan Stanley 的 35%（寬口徑 70%）；FY2027–28 最新共識未取得（付費），EPS 因素的差距可能隨共識更新改變。"
      ]   # 重點觀察（與 README.md 相同文字）
_n0 = RD.max_row
RD.cell(_n0 + 1, 1, "v0.2.1（2026-10-25）：補上檔空間級距、投資建議、重點觀察（Val 頁 0.3 結論加「上檔空間級距」「投資建議」兩列、README 頁加重點觀察三點；數值不變，既有儲存格位置不變）。")
RD.cell(_n0 + 2, 1, f'="投資建議："&Val!B{RC_ROW}&"（依上檔空間固定規則）"').font = BOLD
for _i, _t in enumerate(OBS, 1): RD.cell(_n0 + 2 + _i, 1, f"重點觀察 {_i}：" + _t)

# ---------------- v0.2.2（2026-10-28）：分部原則重建（Q 頁第 40 列起追加；不位移既有列，數值不變） ----------------
Q.cell(40, 1, "官方報告分部與測試類別（2025Q3；分部原則第 3 條）").font = BOLD
_S = "公司 3Q25 法說（富果研究整理，2025-11-28）"
for _r, _lab, _v, _tag, _s1 in (
        (41, "營運部門營收（測試服務約 97%）2025Q3", 9.2905, "Verified", "公司 4Q25 Investor update：3Q25 營收 9,290.5 百萬元；測試服務約占 97%"),
        (42, "  成品測試", "=B41*0.582", "Interested-party", _S + "：58.2%（比重；金額為推得）"),
        (43, "  晶圓測試", "=B41*0.304", "Interested-party", _S + "：30.4%"),
        (44, "  預燒", "=B41*0.09", "Interested-party", _S + "：9.0%"),
        (45, "  封裝及其他", "=B41*0.024", "Interested-party", _S + "：2.4%"),
        (46, "檢查：分部合計 − 合併營收", "=B42+B43+B44+B45-D26", "Derived", "測試類別合計 − 本頁季資料 3Q25 營收；應約為 0")):
    Q.cell(_r, 1, _lab); Q[f"B{_r}"] = _v; Q[f"B{_r}"].number_format = "#,##0.000"
    if not (isinstance(_v, str) and _v.startswith("=")): Q[f"B{_r}"].font = BLUE
    Q.cell(_r, 7, _tag); Q.cell(_r, 8, _s1)
_rr = RD.max_row + 1
RD.cell(_rr, 1, "v0.2.2（2026-10-28）：分部原則重建（數值不變）。Q 頁第 40 列起追加 2025Q3 營運部門營收、測試類別子列（成品測試 58.2%、晶圓測試 30.4%、預燒 9.0%、封裝及其他 2.4%，公司法說）與核對列「分部合計 − 合併營收」。公司年度測試類別金額未取得；推估年仍以 AI 加速器測試／非 AI 拆分。")
RD.cell(1, 1).value = RD.cell(1, 1).value.replace("v0.2.1", "v0.2.2")
order = ["README", "In", "Q", "Units", "PL", "Val", "Recon", "Sens", "Sources"]
wb._sheets = [wb[n] for n in order]
# 相對倍數調整（Andy 2026-10-27「照建議做」；tools/relative_multiple.py，2026-10-29 接入）
import os, sys  # noqa: E401
sys.path.insert(0, os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "tools"))
import relative_multiple as RMX  # noqa: E402
RMX.apply(wb, "kyec")
sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "tools", "in_layout")); import apply_in_layout  # noqa: E402
apply_in_layout.apply(wb, "v0.2.3", "2026-10-30")  # v0.2.3：In 頁版面（個股模型原則第 8 節；數值不變）
wb.save(sys.argv[1])
if os.environ.get("KYEC_ROWMAP"):   # 除錯用：輸出列號對照（供撰寫 html.json）；CI 不設此變數
    import json
    json.dump({"PB": PB, "SUMR": {f"{k[0]}|{k[1]}": v for k, v in SUMR.items()}, "UR": UR, "CHK": CHK, "VAL": VAL, "SB": list(SB), "SENS_LAST": SENS_LAST,
               "SENS_NOTE": SENS_NOTE, "RECON_LAST": RECON_LAST, "WMT_NOTE": WMT_NOTE, "UNOTE": UNOTE, "block_end": block_end, "QSLOPE": QSLOPE},
              open(os.environ["KYEC_ROWMAP"], "w"), ensure_ascii=False, indent=1)
