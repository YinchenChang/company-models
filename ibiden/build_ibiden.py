# Ibiden（4062）收支與估值模型 v0.1（AI 半導體 Project：載板與 PCB 次產業樣板）
# 用法：python3 ibiden/build_ibiden.py raw.xlsx ；再以 LibreOffice headless 重算
# 會計年度 4 月至隔年 3 月：FY(t)＝t/4–(t+1)/3（公司稱「(t+1) 年 3 月期」），對照 CY(t) 的加速器顆數（重疊 9 個月，不做平移）。
# 金額：十億日圓（¥B）；EPS：日圓（2026-10-01 一拆二之後的股數口徑）。美元換算只作對照（¥152／US$，公司 FY2026 前提）。
import sys, os, glob
import openpyxl
from openpyxl.styles import Font, PatternFill
from openpyxl.workbook.defined_name import DefinedName
ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, os.path.join(ROOT, "tools"))
import cycle_position as CP

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
NET_F = only("markets/network/dist/*.xlsx"); TSMC_F = only("tsmc/dist/*.xlsx")
rel = lambda f: os.path.relpath(f, ROOT)
UNITS = {}
for r in openpyxl.load_workbook(NET_F, data_only=True)["In"].iter_rows(min_row=4):
    k = r[0].value
    if k and (str(k).startswith("u_") or str(k).startswith("g29_")):
        UNITS[k] = (r[4].value, r[1].value, f"{rel(NET_F)}｜In!E{r[0].row}")
NETR = {}
for r in openpyxl.load_workbook(NET_F, data_only=True)["Mkt"].iter_rows(min_row=4):
    if r[0].value == "AI IT 資本支出（加速器 × 每顆 IT capex）":
        NETR["itc"] = ([r[2 + j].value for j in range(4)], f"{rel(NET_F)}｜Mkt!C{r[0].row}:F{r[0].row}")
assert "itc" in NETR
TS = {}
for r in openpyxl.load_workbook(TSMC_F, data_only=True)["In"].iter_rows(min_row=4):
    k = r[0].value
    if k and str(k).startswith("um"):
        TS[k] = (r[4].value, f"{rel(TSMC_F)}｜In!E{r[0].row}")
assert all(f"um{y}_{s}" in TS for y in (28, 29) for s in SC), TS.keys()
NV = ("bw", "rb", "ru", "hp"); OT = ("amd", "tpu", "trn", "oth")
assert all(f"u_{g}_{y}" in UNITS for g in NV + OT for y in (26, 27, 28)), sorted(UNITS)

# ---------------- In ----------------
IN = wb.active; IN.title = "In"
IN["A1"] = "In — 輸入（藍字；基準欄 E，低 D、高 F）。金額：十億日圓（¥B）；顆數：百萬顆；比例：%。每個輸入附兩個來源（找不到者寫「未找到」）；公司原文算一個"
IN["A2"] = "標記：Verified／Interested-party／Analogy／Assumed／Derived；Analogy、Assumed 給低／高區間。SemiAnalysis 未使用。公司原文＝Ibiden 決算短信與補足資料（TDnet）；公司預測為 Interested-party"
hdr(IN, 3, ["鍵", "項目", "單位", "低", "基準", "高", "標記", "來源1", "來源2／說明"])
rows = []; sec = lambda t: rows.append(("sec", t)); inp = lambda *a: rows.append(a)
TAN25 = "公司：2026 年 3 月期決算短信（Sources 頁）"
TAN25_URL = "https://www.ibiden.com/ir/items/ConsolidatedFinancialResultsFY2025.pdf"
TANQ1 = "公司：2027 年 3 月期 Q1 決算短信（2026-08-04；Sources 頁）"
TANQ1_URL = "https://finance.biggo.jp/news/jpx_tdnet_140120260731504718"
SUPQ1 = "公司：Q1 補足資料（分部門預測、設備投資、匯率前提；Sources 頁）"
SUPQ1_URL = "https://finance.biggo.jp/news/jpx_tdnet_140120260803507041"
STD = "SEMICON.TODAY 2026-08"
LIMO = "LIMO 2026-08-15"
TOMS = "Tom's Hardware：The state of ABF substrates in 2026"
DG = "Data Gravity 2026-06-14"

sec("A. 加速器顆數（讀 markets/network In 頁；CY2026–CY2028 為輸入、CY2029＝CY2028 ×（1＋成長率））")
for g in NV + OT:
    for y in (26, 27, 28):
        v, lab, src = UNITS[f"u_{g}_{y}"]
        inp(f"u_{g}_{y}", lab, "百萬顆", None, v, None, "Derived", src, "唯讀上游；不在本模型另訂")
    v, lab, src = UNITS[f"g29_{g}"]
    inp(f"g29_{g}", lab, "%", None, v, None, "Derived", src, "同上")
for j, y in enumerate((26, 27, 28, 29)):
    inp(f"itc_{y}", f"AI IT 資本支出 20{y}", "US$B", None, NETR["itc"][0][j], None, "Derived", NETR["itc"][1], "Recon 用：Ibiden AI 營收 ÷ AI IT 資本支出（每 GW 對照）" if y == 26 else "同上")
for s in SC:
    inp(f"um28_{s}", f"情境 {s}：2028 顆數 ÷ 基準", "x", None, TS[f"um28_{s}"][0], None, "Assumed", TS[f"um28_{s}"][1], "A＝基準；B＝AI 資本支出消化；C＝明顯下修（與台積電、NVIDIA、Broadcom、Arista 模型相同）")
    inp(f"um29_{s}", f"情境 {s}：2029 顆數 ÷ 基準", "x", None, TS[f"um29_{s}"][0], None, "Assumed", TS[f"um29_{s}"][1], "同上")

sec("B. Ibiden 份額與每顆載板金額（本專案權威：晶片商載板份額，規則第 13 條）")
inp("w_ot", "非 NVIDIA 加速器的權重（Ibiden 份額 × 每顆載板金額，相對 NVIDIA）", "x", 0.2, 0.4, 0.6, "Assumed",
    "公開資料只有 NVIDIA 份額：BofA（經 thebuildout 2026-03-26）Ibiden 占 GPU 載板 70–80%；高階 FC-BGA 整體約 35%（" + DG + "）", "ASIC 多由欣興、Shinko、AT&S 等分食（第三方研究，Interested-party）；每顆 ASIC 載板金額相對 GPU：未找到")
inp("ai_e", "FY2026 AI 伺服器載板 ÷ 電子事業營收", "%", 0.55, 0.65, 0.75, "Assumed",
    "公司未揭露：未找到（決算短信只說「生成 AI 伺服器向け・汎用サーバー向け高機能 IC パッケージ基板」帶動）", "semi-connect 2026：AI 高階載板營收多數為 NVIDIA 向（估計，Interested-party）")
inp("g_asp", "每顆載板金額年增（面積、層數）", "%", 0.10, 0.18, 0.22, "Analogy",
    TOMS + "：Ibiden 路線圖 2026 年 90×90 mm、2028 年 110×110 mm；每片 AI 伺服器載板的 SAP 負荷 2026 年為 2024 年 1.8 倍、2028 年 2.5 倍（→ 2026→2028 年約 +18%／年）",
    DG + "：ABF 載板平均單價 2024 $65 → 2025 $82（+26%，全品項）")
inp("rel_27", "Ibiden 份額 2027 ÷ 2026（相對值）", "x", 0.85, 0.95, 1.00, "Assumed", "第三方研究：欣興 2025 年拿到 NVIDIA AI 載板約 30% 顆數（Encyclopedia of Stocks 2026-03-04）", "三星電機獲 NVIDIA 載板擴產合作（Hardware Busters 2026）")
inp("rel_28", "Ibiden 份額 2028 ÷ 2026", "x", 0.75, 0.88, 1.00, "Assumed", "同上；競爭者擴產：欣興 2026 資本支出 NT$340 億、景碩 3 年 NT$235 億（" + TOMS + "）", "—")
inp("rel_29", "Ibiden 份額 2029 ÷ 2026", "x", 0.70, 0.85, 1.00, "Assumed", "同上", "—")

sec("C. 短缺溢價（價格指數；2026＝1）與 AI 載板利潤率")
inp("pp_27", "價格溢價指數 2027", "x", 0.95, 1.00, 1.05, "Assumed", TOMS + "（未具名供應鏈分析）：ABF 供需缺口 2026 下半年約 10%、2027 約 21%、2028 超過 40%", "BofA：短缺延續到 2028（經 thebuildout）")
inp("pp_28", "價格溢價指數 2028", "x", 0.85, 0.95, 1.05, "Assumed", "Ibiden 產能 2028 年為 2024 年 2.5–2.8 倍（公司路線圖，Interested-party；" + DG + "、" + TOMS + "）", "欣興、景碩、三星電機同時擴產（同上）")
inp("pp_29", "價格溢價指數 2029", "x", 0.75, 0.90, 1.00, "Assumed", "同上：新產能到位後溢價回落", "—")
inp("m_ne", "電子事業非 AI 部分營業利益率", "%", 0.10, 0.15, 0.20, "Assumed", "FY2024 電子事業營益率 13.6%（26.8 ÷ 197.2，AI 比重較低時；" + TAN25 + "）", "FY2025 18.6%（45.2 ÷ 243.3）")
inp("g_ne", "電子事業非 AI 營收年增（通用伺服器、PC 載板等）", "%", 0.00, 0.05, 0.10, "Assumed", STD + "：通用伺服器市場緩步成長", "PC：未找到")

sec("D. 陶瓷、其他事業與全公司")
inp("g_cer", "陶瓷事業營收年增（DPF、汽車排氣）", "%", -0.05, 0.00, 0.03, "Assumed", "FY2025 −1.8%（84.1 → 82.6；" + TAN25 + "）", "FY2026 預測 84.0（+1.8%；" + SUPQ1 + "）")
inp("m_cer", "陶瓷事業營益率", "%", 0.05, 0.077, 0.10, "Derived", "FY2026 預測 6.5 ÷ 84.0（" + SUPQ1 + "）", "FY2025 9.3%（7.6 ÷ 82.6）")
inp("g_oth", "其他事業營收年增（建材等）", "%", 0.00, 0.03, 0.05, "Assumed", "FY2025 +2.5%（88.1 → 90.3）", "FY2026 預測 91.0（+0.7%）")
inp("m_oth", "其他事業營益率", "%", 0.09, 0.115, 0.13, "Derived", "FY2026 預測 10.5 ÷ 91.0（" + SUPQ1 + "）", "FY2025 9.9%（9.0 ÷ 90.3）")
inp("ni_oi", "淨利 ÷ 營業利益（稅、非營業、少數股權）", "x", 0.62, 0.661, 0.70, "Derived", "FY2026 預測 84 ÷ 127（" + SUPQ1 + "）", "FY2025 1.03（含政策持股出售利益，不具代表性）")
inp("sh_g", "股數年變動（2031 年到期 CB 轉換）", "%", 0.00, 0.00, 0.03, "Assumed", "CB 轉換價 ¥2,243.1（分割後；" + STD + "），遠低於現價", "CB 發行額：未找到；FY2025 稀釋 EPS 比基本低 5.8%（214.91 ÷ 228.16）")
inp("fx", "匯率（美元換算用）", "¥／US$", None, 152, None, "Interested-party", SUPQ1 + "：FY2026 前提 ¥152", "—")

sec("E. FY2026 錨點（公司 2026-08-04 上修後預測；Q2 決算後更新）")
inp("f_rev", "FY2026 營收預測", "¥B", None, 550.0, None, "Interested-party", SUPQ1, LIMO + "：5,500 億（+32.1%）")
inp("f_e", "FY2026 電子事業營收預測", "¥B", None, 375.0, None, "Interested-party", SUPQ1, "Q1 實績 77.5（年增 37.7%；" + STD + "）")
inp("f_eoi", "FY2026 電子事業營業利益預測", "¥B", None, 110.0, None, "Interested-party", SUPQ1, "Q1 實績 21.4（營益率 27.6%）")
inp("f_cer", "FY2026 陶瓷事業營收預測", "¥B", None, 84.0, None, "Interested-party", SUPQ1, "—")
inp("f_ceroi", "FY2026 陶瓷事業營業利益預測", "¥B", None, 6.5, None, "Interested-party", SUPQ1, "—")
inp("f_oth", "FY2026 其他事業營收預測", "¥B", None, 91.0, None, "Interested-party", SUPQ1, "—")
inp("f_othoi", "FY2026 其他事業營業利益預測", "¥B", None, 10.5, None, "Interested-party", SUPQ1, "—")
inp("f_oi", "FY2026 營業利益預測", "¥B", None, 127.0, None, "Interested-party", SUPQ1, STD + "：900 → 1,270 億")
inp("f_ni", "FY2026 淨利預測（歸屬母公司）", "¥B", None, 84.0, None, "Interested-party", SUPQ1, STD + "：580 → 840 億")
inp("f_eps", "FY2026 EPS 預測（分割後）", "¥", None, 149.68, None, "Interested-party", STD + "：反映 2026-10-01 一拆二", "—")

sec("F. FY2025 實績（2026 年 3 月期）")
for k, lab, v in (("a_rev", "營收", 416.201), ("a_e", "電子事業營收", 243.316), ("a_eoi", "電子事業營業利益", 45.248), ("a_cer", "陶瓷事業營收", 82.554), ("a_ceroi", "陶瓷事業營業利益", 7.646),
                  ("a_oth", "其他事業營收", 90.330), ("a_othoi", "其他事業營業利益", 8.964), ("a_oi", "營業利益", 62.027), ("a_gp", "毛利", 131.668), ("a_ni", "淨利（歸屬母公司）", 63.713)):
    inp(k, f"FY2025 {lab}", "¥B", None, v, None, "Verified", TAN25, "SEMICON.TODAY、semi-connect 同值" if k in ("a_rev", "a_e", "a_oi") else "—")
inp("a_eps", "FY2025 基本 EPS（分割後換算＝228.16 ÷ 2）", "¥", None, "=228.16/2", None, "Derived", TAN25 + "：228.16（2026-01 一拆二後、2026-10 一拆二前）", "2026-10-01 再一拆二")

sec("G. 價格與市場資料（只作 Val、Recon 對照）")
inp("px", "股價（2026-10-16 收盤，分割後）", "¥", None, 9764, None, "Interested-party", "stockanalysis（Sources 頁）", "10-23 收盤：未找到")
inp("pe_fwd_sa", "前瞻 P/E（stockanalysis）", "x", None, 53.97, None, "Interested-party", "stockanalysis（口徑：未來 12 個月共識，頁面未明示）", "第二來源：未找到")
inp("pe_ttm", "近 12 月 P/E", "x", None, 83.85, None, "Interested-party", "stockanalysis", "—")
inp("pe_fair_andy", "Andy 先驗：合理 P/E", "x", None, 25, None, "Assumed", "先驗_Andy_Ibiden_20261024 I10b（Drive；不在本 repo）", "只作對照")
inp("tgt30_rev", "公司 FY2030 目標營收", "¥B", None, 1000, None, "Interested-party", "Quartr FY2025 摘要（AI 摘要；營收數字有誤處未採用）", "第二來源：未找到")
inp("tgt30_oi", "公司 FY2030 目標營業利益", "¥B", None, 300, None, "Interested-party", "同上", "第二來源：未找到")
sec("H. 目標價情境權重（v0.1.1 新增；Val 頁第 0 節用）")
for s, w in (("A", 0.3), ("B", 0.5), ("C", 0.2)):
    inp(f"tp_w{s}", f"目標價情境權重 {s}", "%", None, w, None, "Assumed", "Andy 可改（工作單 2026-10-25 預設 A 0.3／B 0.5／C 0.2）", "三者合計應為 100%（Val 頁第 0 節列示）")

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
    wb.defined_names[key] = DefinedName(key, attr_text=f"In!$E${r}")
    if lo is not None: wb.defined_names[key + "_lo"] = DefinedName(key + "_lo", attr_text=f"In!$D${r}")
    if hi is not None: wb.defined_names[key + "_hi"] = DefinedName(key + "_hi", attr_text=f"In!$F${r}")
    r += 1
for col, w in zip("ABCDEFGHI", [12, 58, 10, 9, 11, 9, 18, 80, 70]): IN.column_dimensions[col].width = w
IN.freeze_panes = "D4"
def defname(name, ref):
    wb.defined_names[name] = DefinedName(name, attr_text=ref)

# ---------------- Q：實績與錨點 ----------------
Q = wb.create_sheet("Q")
Q["A1"] = "Q — FY2025 實績與 FY2026 錨點（¥B）。FY2026 錨點＝公司 2026-08-04 上修後的全年預測（Q1 公布時；分部門預測取自 Q1 補足資料）。Q2 決算（預計 10 月底或 11 月初）後更新"
Q["A2"] = "公司未揭露 AI 伺服器載板營收；本模型以 In 頁 ai_e（Assumed）切分電子事業，並據此校準每顆有效金額與 AI 載板營業利益率（Derived）"
hdr(Q, 3, ["項目", "FY2025A", "FY2026 錨點", "年增", "標記", "說明"])
QR = {}
for i, (k, lab, a, f, fmt) in enumerate((("rev", "營收", "=a_rev", "=f_rev", "#,##0.0"), ("e", "  電子事業", "=a_e", "=f_e", "#,##0.0"), ("cer", "  陶瓷事業", "=a_cer", "=f_cer", "#,##0.0"),
                                          ("oth", "  其他事業", "=a_oth", "=f_oth", "#,##0.0"), ("oi", "營業利益", "=a_oi", "=f_oi", "#,##0.0"), ("eoi", "  電子事業", "=a_eoi", "=f_eoi", "#,##0.0"),
                                          ("ceroi", "  陶瓷事業", "=a_ceroi", "=f_ceroi", "#,##0.0"), ("othoi", "  其他事業", "=a_othoi", "=f_othoi", "#,##0.0"),
                                          ("adj", "  調整（全公司 − 三事業）", "=a_oi-a_eoi-a_ceroi-a_othoi", "=f_oi-f_eoi-f_ceroi-f_othoi", "#,##0.00"),
                                          ("ni", "淨利（歸屬母公司）", "=a_ni", "=f_ni", "#,##0.0"), ("eps", "EPS（分割後，日圓）", "=a_eps", "=f_eps", "#,##0.00")), 4):
    Q.cell(i, 1, lab); Q[f"B{i}"] = a; Q[f"C{i}"] = f; Q[f"D{i}"] = f"=C{i}/B{i}-1"
    for c in "BC": Q[f"{c}{i}"].number_format = fmt
    Q[f"D{i}"].number_format = "0.0%"
    Q.cell(i, 5, "Verified／Interested-party"); QR[k] = i
Q.cell(QR["ni"], 6, "FY2025 含政策持股出售利益（稅後約 345 億，Quartr 摘要），淨利高於營業利益")
r = QR["eps"] + 2
Q.cell(r, 1, "比率").font = BOLD; r += 1
for k, lab, fb, fc in (("em", "電子事業營益率", "=B{e}/B{s}".format(e=QR["eoi"], s=QR["e"]), "=C{e}/C{s}".format(e=QR["eoi"], s=QR["e"])),
                       ("om", "營業利益率", "=B{e}/B{s}".format(e=QR["oi"], s=QR["rev"]), "=C{e}/C{s}".format(e=QR["oi"], s=QR["rev"])),
                       ("esh", "電子事業 ÷ 營收", "=B{e}/B{s}".format(e=QR["e"], s=QR["rev"]), "=C{e}/C{s}".format(e=QR["e"], s=QR["rev"])),
                       ("gm", "毛利率（FY2026 預測未揭露）", "=a_gp/a_rev", None),
                       ("sh", "股數（十億股＝淨利 ÷ EPS）", f"=B{QR['ni']}/B{QR['eps']}", f"=C{QR['ni']}/C{QR['eps']}")):
    Q.cell(r, 1, lab); Q[f"B{r}"] = fb
    if fc: Q[f"C{r}"] = fc
    for c in "BC": Q[f"{c}{r}"].number_format = "0.0%" if k != "sh" else "0.0000"
    Q.cell(r, 5, "Derived"); QR[k] = r; r += 1
Q.cell(QR["sh"], 6, "FY2025 股數含政策持股出售前後變動，只作參考；預測期用 FY2026 錨點股數")
for nm, k in (("rev", "rev"), ("e", "e"), ("cer", "cer"), ("oth", "oth"), ("oi", "oi"), ("eoi", "eoi"), ("ni", "ni"), ("eps", "eps"), ("adj", "adj")):
    defname(f"q25_{nm}", f"Q!$B${QR[k]}"); defname(f"q26_{nm}", f"Q!$C${QR[k]}")
defname("q26_sh", f"Q!$C${QR['sh']}")
Q.column_dimensions["A"].width = 40
for c in "BCD": Q.column_dimensions[c].width = 14
Q.column_dimensions["E"].width = 26; Q.column_dimensions["F"].width = 80
Q.freeze_panes = "B4"

# ---------------- 引擎 ----------------
FYS = (2025, 2026, 2027, 2028, 2029); FC = {2025: "C", 2026: "D", 2027: "E", 2028: "F", 2029: "G"}
BU = (2026, 2027, 2028, 2029)
umx = lambda y, s: {2026: "1", 2027: "1", 2028: f"um28_{s}", 2029: f"um29_{s}"}[y]
def units_f(g, y):
    return f"u_{g}_{y % 100}" if y <= 2028 else f"u_{g}_28*(1+g29_{g})"
DEFP = dict(scn="A", w="w_ot", aie="ai_e", ga="g_asp", rel="rel", pp="pp", mne="m_ne", gne="g_ne", gcer="g_cer", goth="g_oth", nioi="ni_oi", shg="sh_g")
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
    sec_("A. AI 伺服器載板（D 欄＝FY2026 校準年；FY(t) 對照 CY(t) 顆數）")
    row("unv", "NVIDIA 加速器顆數 × 情境係數", "百萬顆", {c_(y): "=(" + "+".join(units_f(g, y) for g in NV) + f")*{umx(y, s)}" for y in BU}, note="markets/network In：Blackwell、Rubin、Rubin Ultra、Hopper")
    row("uot", "非 NVIDIA 加速器顆數 × 情境係數", "百萬顆", {c_(y): "=(" + "+".join(units_f(g, y) for g in OT) + f")*{umx(y, s)}" for y in BU}, note="AMD、TPU、Trainium、其他 ASIC")
    row("ueff", "有效顆數 ＝ NVIDIA ＋ 權重 × 非 NVIDIA", "百萬顆", {c_(y): f"={c_(y)}{R['unv']}+{P['w']}*{c_(y)}{R['uot']}" for y in BU})
    row("ai0", "FY2026 AI 載板營收（錨點電子事業 × AI 占比）", "¥B", {"D": f"=q26_e*{P['aie']}"}, note="Assumed 切分；公司未揭露")
    row("k", "每顆有效金額（Ibiden 份額 × 載板金額；FY2026 校準）", "千日圓／顆", {"D": f"=D{R['ai0']}/D{R['ueff']}"}, fmt="#,##0.0",
        note="Derived：¥B ÷ 百萬顆＝千日圓／顆。÷ 匯率即美元／顆（Units 頁）")
    row("asp", "每顆載板金額指數（面積、層數；FY2026＝1）", "x", {"D": "=1", **{c_(y): f"={c_(y - 1)}{r}*(1+{P['ga']})" for y in BU[1:]}}, fmt="0.000")
    row("rel", "Ibiden 份額相對值（÷ FY2026）", "x", {"D": "=1", **{c_(y): f"={P['rel']}_{y % 100}" for y in BU[1:]}}, fmt="0.00")
    row("pp", "短缺溢價指數（FY2026＝1）", "x", {"D": "=1", **{c_(y): f"={P['pp']}_{y % 100}" for y in BU[1:]}}, fmt="0.00")
    row("ai", "AI 伺服器載板營收", "¥B", {"D": f"=D{R['ai0']}", **{c_(y): f"={c_(y)}{R['ueff']}*$D${R['k']}*{c_(y)}{R['asp']}*{c_(y)}{R['rel']}*{c_(y)}{R['pp']}" for y in BU[1:]}})
    sec_("B. 電子事業")
    row("ne", "電子事業非 AI 營收（通用伺服器、PC 等）", "¥B", {"D": f"=q26_e-D{R['ai']}", **{c_(y): f"={c_(y - 1)}{r}*(1+{P['gne']})" for y in BU[1:]}})
    row("e", "電子事業營收", "¥B", {"C": "=q25_e", **{c_(y): f"={c_(y)}{R['ai']}+{c_(y)}{R['ne']}" for y in BU}})
    row("mai0", "AI 載板營業利益率（FY2026 校準）", "%", {"D": f"=(q26_eoi-{P['mne']}*D{R['ne']})/D{R['ai']}"}, fmt="0.0%",
        note="Derived：使 FY2026 電子事業營業利益＝錨點；非 AI 營益率為 In 頁 m_ne")
    row("mai", "AI 載板營業利益率 ＝ 1 −（1 − 校準值）÷ 溢價指數", "%", {"D": f"=D{R['mai0']}", **{c_(y): f"=1-(1-$D${R['mai0']})/{c_(y)}{R['pp']}" for y in BU[1:]}}, fmt="0.0%",
        note="Derived：每單位成本不變，價格隨溢價指數升降，差額全部落在利潤")
    row("eoi", "電子事業營業利益", "¥B", {"C": "=q25_eoi", **{c_(y): f"={c_(y)}{R['ai']}*{c_(y)}{R['mai']}+{c_(y)}{R['ne']}*{P['mne']}" for y in BU}})
    row("em", "電子事業營益率", "%", {c_(y): f"={c_(y)}{R['eoi']}/{c_(y)}{R['e']}" for y in FYS}, fmt="0.0%")
    sec_("C. 陶瓷、其他事業與全公司")
    row("cer", "陶瓷事業營收", "¥B", {"C": "=q25_cer", "D": "=q26_cer", **{c_(y): f"={c_(y - 1)}{r}*(1+{P['gcer']})" for y in BU[1:]}})
    row("oth", "其他事業營收", "¥B", {"C": "=q25_oth", "D": "=q26_oth", **{c_(y): f"={c_(y - 1)}{r}*(1+{P['goth']})" for y in BU[1:]}})
    row("rev", "營收", "¥B", {c_(y): f"={c_(y)}{R['e']}+{c_(y)}{R['cer']}+{c_(y)}{R['oth']}" for y in FYS})
    row("g", "營收年增", "%", {c_(y): f"={c_(y)}{R['rev']}/{c_(y - 1)}{R['rev']}-1" for y in FYS[1:]}, fmt="0.0%")
    row("aish", "AI 伺服器載板 ÷ 營收", "%", {c_(y): f"={c_(y)}{R['ai']}/{c_(y)}{R['rev']}" for y in BU}, fmt="0.0%")
    row("ceroi", "陶瓷＋其他事業營業利益＋調整", "¥B", {"C": "=q25_oi-q25_eoi", "D": "=q26_oi-q26_eoi",
        **{c_(y): f"={c_(y)}{R['cer']}*m_cer+{c_(y)}{R['oth']}*m_oth+q26_adj" for y in BU[1:]}})
    row("oi", "營業利益", "¥B", {c_(y): f"={c_(y)}{R['eoi']}+{c_(y)}{R['ceroi']}" for y in FYS})
    row("oim", "營業利益率", "%", {c_(y): f"={c_(y)}{R['oi']}/{c_(y)}{R['rev']}" for y in FYS}, fmt="0.0%")
    row("ni", "淨利（歸屬母公司）", "¥B", {"C": "=q25_ni", "D": "=q26_ni", **{c_(y): f"={c_(y)}{R['oi']}*{P['nioi']}" for y in BU[1:]}})
    row("sh", "股數（分割後）", "十億股", {"D": "=q26_sh", **{c_(y): f"={c_(y - 1)}{r}*(1+{P['shg']})" for y in BU[1:]}}, fmt="0.0000")
    row("eps", "EPS（日圓）", "¥", {"C": "=q25_eps", **{c_(y): f"={c_(y)}{R['ni']}/{c_(y)}{R['sh']}" for y in BU}})
    row("epsg", "EPS 年增", "%", {c_(y): f"={c_(y)}{R['eps']}/{c_(y - 1)}{R['eps']}-1" for y in FYS[1:]}, fmt="0.0%")
    R["end"] = r
    return R

# ---------------- PL ----------------
PL = wb.create_sheet("PL")
PL["A1"] = "PL — Ibiden 損益（¥B；EPS 日圓）。FY2025A 實績、FY2026 錨點（公司上修後預測），FY2027 起為模型。FY(t)＝t/4–(t+1)/3，對照 CY(t) 顆數。三情境各一個完整推導區塊（本頁下方），頂端為摘要。FY2026–FY2027 三情境相同（顆數係數從 CY2028 開始）"
PB = {}; r0 = 60
SCN_T = {"A": "情境 A：基準（顆數係數 1.0）", "B": "情境 B：CY2028 顆數 ×0.8、CY2029 ×0.85", "C": "情境 C：CY2028–29 顆數 ×0.6"}
for s in SC:
    PB[s] = engine(PL, r0, SCN_T[s], scn=s)
    r0 = PB[s]["end"] + 2
hdr(PL, 3, ["情境／項目", "單位", "FY2025A", "FY2026（錨點）", "FY2027", "FY2028", "FY2029", "說明"])
sr = 4; SUMR = {}
for s in SC:
    PL.cell(sr, 1, f"情境 {s}").font = BOLD; sr += 1
    for key, lab, unit, fmt in (("rev", "營收", "¥B", "#,##0.0"), ("g", "  營收年增", "%", "0.0%"), ("ai", "  AI 伺服器載板", "¥B", "#,##0.0"), ("e", "  電子事業", "¥B", "#,##0.0"),
                                ("em", "電子事業營益率", "%", "0.0%"), ("oim", "營業利益率", "%", "0.0%"), ("ni", "淨利", "¥B", "#,##0.0"),
                                ("eps", "EPS", "¥", "#,##0.0"), ("epsg", "  EPS 年增", "%", "0.0%")):
        PL.cell(sr, 1, lab); PL.cell(sr, 2, unit)
        for y in FYS:
            if key in ("g", "epsg") and y == 2025: continue
            if key == "ai" and y == 2025: continue
            PL[f"{FC[y]}{sr}"] = f"={FC[y]}{PB[s][key]}"; PL[f"{FC[y]}{sr}"].number_format = fmt
        SUMR[(s, key)] = sr; sr += 1
    sr += 1
PL.cell(sr, 1, "核對（應為 0）：FY2026 模型 − 錨點（校準恆等式）").font = BOLD; sr += 1
CHK = {}
for k, lab, ref in (("rev", "營收差（¥B）", "q26_rev"), ("oi", "營業利益差（¥B）", "q26_oi"), ("eps", "EPS 差（日圓）", "q26_eps")):
    PL.cell(sr, 1, f"  情境 A：{lab}"); PL[f"D{sr}"] = f"=D{PB['A'][k]}-{ref}"; PL[f"D{sr}"].number_format = "0.0000"; CHK[k] = sr; sr += 1
assert sr < 60, sr
PL.column_dimensions["A"].width = 62; PL.column_dimensions["B"].width = 12; PL.column_dimensions["H"].width = 70
for c in "CDEFG": PL.column_dimensions[c].width = 14

# ---------------- Units：份額與每顆金額（權威） ----------------
UN = wb.create_sheet("Units")
UN["A1"] = "Units — Ibiden 在 AI 加速器載板的有效金額與份額（本專案權威，規則第 13 條）。顆數取自 markets/network（情境 A）"
A_ = PB["A"]
cyc = lambda y: chr(ord("C") + y - 2026)
hdr(UN, 3, ["項目", "單位", "FY2026（CY2026）", "FY2027", "FY2028", "FY2029", "說明"])
ur = 4; UR = {}
for key, lab, unit, f, fmt, note in (
        ("unv", "NVIDIA 加速器顆數", "百萬顆", lambda y: f"=PL!{FC[y]}{A_['unv']}", "#,##0.00", "markets/network"),
        ("uot", "非 NVIDIA 加速器顆數", "百萬顆", lambda y: f"=PL!{FC[y]}{A_['uot']}", "#,##0.00", ""),
        ("ueff", "有效顆數", "百萬顆", lambda y: f"=PL!{FC[y]}{A_['ueff']}", "#,##0.00", "非 NVIDIA × 權重 w_ot"),
        ("kusd", "每顆 NVIDIA 加速器的 Ibiden 營收（美元）", "US$", lambda y: f"=PL!$D${A_['k']}*PL!{FC[y]}{A_['asp']}*PL!{FC[y]}{A_['rel']}*PL!{FC[y]}{A_['pp']}*1000/fx", "#,##0", "＝份額 × 每顆載板金額"),
        ("aspusd", "每顆 NVIDIA 載板金額（美元；假設 FY2026 份額 70%）", "US$", lambda y: f"=PL!$D${A_['k']}*PL!{FC[y]}{A_['asp']}*PL!{FC[y]}{A_['pp']}*1000/fx/0.7", "#,##0", "BofA：Ibiden 占 GPU 載板 70–80%（Interested-party）；先驗比對推導 $150–300"),
        ("ai", "AI 伺服器載板營收", "¥B", lambda y: f"=PL!{FC[y]}{A_['ai']}", "#,##0.0", ""),
        ("aiusd", "AI 伺服器載板營收（美元）", "US$B", lambda y: f"=PL!{FC[y]}{A_['ai']}/fx", "#,##0.00", "÷ 匯率 152"),
        ("itr", "AI 載板營收 ÷ AI IT 資本支出", "%", lambda y: f"=PL!{FC[y]}{A_['ai']}/fx/itc_{y % 100}", "0.000%", "× 每 GW IT capex 即每 GW 營收"),
        ("pgw", "每 GW Ibiden 營收（VR200 每 GW IT capex 37.59B）", "US$M", lambda y: f"=PL!{FC[y]}{A_['ai']}/fx/itc_{y % 100}*37.59*1000", "#,##0", "Tokenomics IF_CapexIT（VR200 基準，經橋接表）")):
    UN.cell(ur, 1, lab); UN.cell(ur, 2, unit)
    for y in BU:
        UN[f"{cyc(y)}{ur}"] = f(y); UN[f"{cyc(y)}{ur}"].number_format = fmt
    UN.cell(ur, 7, note); UR[key] = ur; ur += 1
ur += 1
UN.cell(ur, 1, "讀法：每顆 NVIDIA 加速器 Ibiden 只分到約 1–2 百美元，每 GW 約數千萬美元；瓶頸反映在利潤率（電子事業營益率）與倍數，不在營收規模。").font = BOLD
UNOTE = ur
UN.column_dimensions["A"].width = 60; UN.column_dimensions["B"].width = 10
for c in "CDEF": UN.column_dimensions[c].width = 16
UN.column_dimensions["G"].width = 70

# ---------------- Val ----------------
V = wb.create_sheet("Val")
V["A1"] = "Val — P/E 估值（日圓）。結論用「現價 ÷ FY2026–FY2028 EPS（模型）」與 what must be true。十年 P/E 規則值待 Bloomberg（10-27 抓取單加 4062）：公開資料商的 Ibiden 十年 P/E 本版未取得，規則值本版不計"
V["A2"] = "循環位置：Ibiden 不套設備股見頂年；以本模型營收最高的會計年度為見頂年（比照 NVIDIA、Arista）。營收到 FY2029 仍在增加時，見頂年未知，判為中段"
V.cell(4, 1, "1. 價格與市場隱含倍數").font = BOLD
hdr(V, 5, ["項目", "A", "B", "C", "說明"])
V.cell(6, 1, "現價（2026-10-16 收盤，分割後）"); V["B6"] = "=px"; V["B6"].number_format = "#,##0"
V.cell(7, 1, "現價 ÷ FY2026 EPS（錨點）"); V.cell(8, 1, "現價 ÷ FY2027 EPS（模型）"); V.cell(9, 1, "現價 ÷ FY2028 EPS（模型）"); V.cell(10, 1, "現價 ÷ FY2029 EPS（模型）")
for j, s in enumerate(SC):
    c = "BCD"[j]
    for rr, y in ((7, 2026), (8, 2027), (9, 2028), (10, 2029)):
        V[f"{c}{rr}"] = f"=px/PL!{FC[y]}{SUMR[(s, 'eps')]}"; V[f"{c}{rr}"].number_format = "0.0"
V.cell(11, 1, "前瞻 P/E（stockanalysis）"); V["B11"] = "=pe_fwd_sa"; V["B11"].number_format = "0.0"
V.cell(12, 1, "  隱含未來 12 月 EPS ＝ 現價 ÷ 前瞻 P/E"); V["B12"] = "=px/pe_fwd_sa"; V["B12"].number_format = "0.0"
V.cell(13, 1, "近 12 月 P/E（stockanalysis）"); V["B13"] = "=pe_ttm"; V["B13"].number_format = "0.0"
V["E8"] = "FY2027 三情境相同（顆數係數從 CY2028 開始）"; V["E11"] = "口徑：未來 12 個月"
V.cell(15, 1, "2. 循環位置（本模型營收最高年）").font = BOLD
hdr(V, 16, ["情境", "營收最高年（FY2026–FY2029）", "FY2027 判定", "見頂年是否未知", "說明"])
for i, s in enumerate(SC, 17):
    rr = f"PL!D{PB[s]['rev']}:G{PB[s]['rev']}"
    V.cell(i, 1, s)
    V.cell(i, 2, f"=INDEX({{2026,2027,2028,2029}},MATCH(MAX({rr}),{rr},0))")
    V.cell(i, 3, '=IF(B{0}=2029,"中段",'.format(i) + CP.excel_rule("2027", f"B{i}") + ")")
    V.cell(i, 4, f'=IF(B{i}=2029,"是：營收到觀測期末仍在增加","否")')
    V.cell(i, 5, "tools/cycle_position.py excel_rule；營收最高年取代設備市場見頂年；對應倍數待十年 P/E" if s == "A" else "同上")
V.cell(21, 1, "3. What must be true：FY2027 年底價 ＝ 現價 ×（1＋報酬）所需的倍數（看 FY2028 EPS）").font = BOLD
hdr(V, 22, ["FY2027 年底價 ÷ 現價", "所需倍數（÷ FY2028 EPS，A）", "所需倍數（÷ FY2028 EPS，B）", "所需倍數（÷ FY2028 EPS，C）", "對照：Andy 合理 P/E"])
vr = 23
for x in (0.5, 0.75, 1.0, 1.21, 1.5):
    V.cell(vr, 1, x).font = BLUE; V.cell(vr, 1).number_format = "0.00"
    for j, s in enumerate(SC):
        V.cell(vr, 2 + j, f"=px*A{vr}/PL!F{SUMR[(s, 'eps')]}").number_format = "0.0"
    V.cell(vr, 5, "=pe_fair_andy").number_format = "0.0"
    vr += 1
V.cell(vr, 1, "讀法：1.00＝價格不變、1.21＝一年約 21%。所需倍數愈高，現價愈依賴「短缺溢價延續到 2028 年以後」。"); WMT_NOTE = vr
vr += 2
V.cell(vr, 1, "4. 對照：以 Andy 合理 P/E 25x 計算的隱含價（只作對照，非規則值）").font = BOLD; vr += 1
hdr(V, vr, ["情境", "FY2027 EPS", "× 25x", "相對現價", "FY2028 EPS", "× 25x", "相對現價"]); IMPH = vr; vr += 1
IMP = {}
for s in SC:
    V.cell(vr, 1, s)
    V.cell(vr, 2, f"=PL!E{SUMR[(s, 'eps')]}").number_format = "0.0"; V.cell(vr, 3, f"=B{vr}*pe_fair_andy").number_format = "#,##0"; V.cell(vr, 4, f"=C{vr}/px-1").number_format = "0%"
    V.cell(vr, 5, f"=PL!F{SUMR[(s, 'eps')]}").number_format = "0.0"; V.cell(vr, 6, f"=E{vr}*pe_fair_andy").number_format = "#,##0"; V.cell(vr, 7, f"=F{vr}/px-1").number_format = "0%"
    for c in range(1, 8): V.cell(vr, c).fill = KEY
    IMP[s] = vr; vr += 1
V.column_dimensions["A"].width = 58
for c in "BCDEFG": V.column_dimensions[c].width = 18

# ---------------- Sens ----------------
SE = wb.create_sheet("Sens")
SE["A1"] = "Sens — 一次動一個輸入（情境 A；每顆有效金額與 AI 載板營益率在每個案例內依 FY2026 錨點重算，FY2026 恆等於錨點）：對 FY2027、FY2028 的營收、AI 載板營收、電子事業營益率與 EPS 的影響。每個案例是一個完整推導區塊（本頁下方）"
hdr(SE, 3, ["案例", "設定", "FY2027 營收", "FY2027 AI 載板", "FY2027 電子營益率", "FY2027 EPS", "FY2027 EPS 相對基準", "FY2028 營收", "FY2028 AI 載板", "FY2028 電子營益率", "FY2028 EPS", "FY2028 EPS 相對基準"])
cases = [("溢價快速回落", "價格溢價 2027 0.95、2028 0.85、2029 0.75（新產能提早到位）", dict(pp="PPLO")),
         ("溢價延續", "價格溢價 2027 1.05、2028 1.05、2029 1.00", dict(pp="PPHI")),
         ("份額流失（低）", "份額相對值 2027 0.85、2028 0.75、2029 0.70（欣興、三星電機加速）", dict(rel="RLO")),
         ("份額持平（高）", "份額相對值 1.00", dict(rel="RHI")),
         ("每顆金額年增 10%", "面積、層數成長較慢", dict(ga="g_asp_lo")), ("每顆金額年增 22%", "較快", dict(ga="g_asp_hi")),
         ("非 NVIDIA 權重 0.2", "ASIC 多由其他載板廠供應", dict(w="w_ot_lo")), ("非 NVIDIA 權重 0.6", "Ibiden 在 ASIC 份額較高", dict(w="w_ot_hi")),
         ("AI 占電子事業 55%", "FY2026 切分較低（校準跟著變）", dict(aie="ai_e_lo")), ("AI 占電子事業 75%", "較高", dict(aie="ai_e_hi")),
         ("非 AI 營益率 10%", "校準後 AI 載板營益率較高", dict(mne="m_ne_lo")), ("非 AI 營益率 20%", "較低", dict(mne="m_ne_hi")),
         ("淨利 ÷ 營業利益 0.62", "稅與非營業較不利", dict(nioi="ni_oi_lo")),
         ("CB 轉換：股數 +3%／年", "2031 年到期 CB 轉換稀釋", dict(shg="sh_g_hi"))]
for nm, base, suf in (("PPLO", "pp", "_lo"), ("PPHI", "pp", "_hi"), ("RLO", "rel", "_lo"), ("RHI", "rel", "_hi")):
    for y in (27, 28, 29):
        wb.defined_names[f"{nm}_{y}"] = DefinedName(f"{nm}_{y}", attr_text=wb.defined_names[f"{base}_{y}{suf}"].attr_text)
def put_case(i, name, setting, R, pre):
    SE.cell(i, 1, name); SE.cell(i, 2, setting)
    for c, key, y, fmt in (("C", "rev", 2027, "#,##0.0"), ("D", "ai", 2027, "#,##0.0"), ("E", "em", 2027, "0.0%"), ("F", "eps", 2027, "#,##0.0"),
                           ("H", "rev", 2028, "#,##0.0"), ("I", "ai", 2028, "#,##0.0"), ("J", "em", 2028, "0.0%"), ("K", "eps", 2028, "#,##0.0")):
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
SENS_LAST = nn + 1
SE.cell(nn + 3, 1, "讀法：FY2027 三個顆數情境相同；B、C 的差異從 FY2028 起。溢價指數同時改變營收與 AI 載板營益率（成本不變），所以對 EPS 的影響大於對營收的影響。"); SENS_NOTE = nn + 3
SE.column_dimensions["A"].width = 30; SE.column_dimensions["B"].width = 56
for c in "CDEFGHIJKL": SE.column_dimensions[c].width = 14

# ---------------- Recon ----------------
RC = wb.create_sheet("Recon")
RC["A1"] = "Recon — 本模型 vs 公司目標、研調份額、先驗。只作事後對照，不為對上數字而調參數。FY2027 以後的共識：未找到（付費）"
hdr(RC, 3, ["項目", "對照值", "模型 A", "差距 A", "模型 B", "差距 B", "模型 C", "差距 C", "對照來源", "說明"])
rr_ = 4
def rec(label, ref, getter, src, note="", fmt="#,##0.0", pt=False, gap=True):
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
RC.cell(rr_, 1, "A. 公司目標與市場（校準未使用）").font = BOLD; rr_ += 1
rec("FY2029 營收 vs 公司 FY2030 目標 ¥1 兆", "=tgt30_rev", pl("rev", 2029), "Quartr 摘要（Interested-party）", "模型到 FY2029；目標年晚一年")
rec("FY2029 營業利益 vs 公司 FY2030 目標 ¥3,000 億", "=tgt30_oi", pl("oi", 2029), "同上", "")
rec("FY2027 EPS vs 前瞻 P/E 隱含未來 12 月 EPS", "=px/pe_fwd_sa", pl("eps", 2027), "stockanalysis（推導）", "未來 12 月 ≈ ½ FY2026＋½ FY2027，口徑不同，只作量級比較")
RC.cell(rr_, 1, "B. 份額與每顆金額（先驗比對的機制檢查）").font = BOLD; rr_ += 1
rec("FY2026 每顆 NVIDIA 載板金額（美元）vs 先驗比對推導中點 $225", 225, lambda s: f"=Units!C{UR['aspusd']}", "先驗比對_Ibiden（Drive）", "推導區間 $150–300；封裝全部（中介層＋載板＋組裝）B200 約 $1,100（Silicon Analysts）", fmt="#,##0")
rec("FY2026 每 GW Ibiden 營收（US$M）vs 先驗比對推導中點 75", 75, lambda s: f"=Units!C{UR['pgw']}", "先驗比對_Ibiden（Drive）", "", fmt="#,##0")
rec("FY2027 AI 載板 ÷ 營收", None, pl("aish", 2027), "—", "", fmt="0.0%")
rec("FY2028 電子事業營益率 vs FY2026 錨點（百分點）", "=q26_eoi/q26_e", pl("em", 2028), "錨點", "短缺溢價回落使營益率下滑", fmt="0.0%", pt=True)
RC.cell(rr_, 1, "C. 預留").font = BOLD; rr_ += 1
RC.cell(rr_, 1, "產能上限（Ibiden 2028 年為 2024 年 2.5–2.8 倍）與需求面積的對照：2024 年基準面積未找到，本版不計"); RC.cell(rr_, 10, "預留列：本版無數值"); rr_ += 1
RC.cell(rr_, 1, "R4 產業部署路徑（markets/deploy_path/，尚未建立）：上線後以 AI 載板營收 ÷ 部署 GW 取代 Units 頁的每 GW 換算"); RC.cell(rr_, 10, "預留列：本版無數值"); rr_ += 1
RECON_LAST = rr_ - 1
RC.column_dimensions["A"].width = 70
for c in "BCDEFGH": RC.column_dimensions[c].width = 12
RC.column_dimensions["I"].width = 36; RC.column_dimensions["J"].width = 70

# ---------------- Sources ----------------
SO = wb.create_sheet("Sources"); SO["A1"] = "來源（公司原文＝決算短信與補足資料；分部門預測取自 Q1 補足資料的轉載）"; SO["A1"].font = BOLD
srcs = [("公司：2026 年 3 月期決算短信（FY2025 實績、分部門）", TAN25_URL),
        ("公司：2027 年 3 月期 Q1 決算短信（上修後全年預測）", TANQ1_URL),
        ("公司：Q1 補足資料（分部門全年預測、設備投資 2,100 億、匯率 ¥152）", SUPQ1_URL),
        ("二手：SEMICON.TODAY Q1 深掘（分部門 Q1、EPS 分割後 149.68）", "https://semicon.today/archives/12352"),
        ("二手：LIMO Q1（預收款、股價）", "https://finance.limo.media/articles/detail/683"),
        ("二手：Quartr FY2025 摘要（FY2030 目標；AI 摘要，營收數字有誤處未採用）", "https://quartr.com/events/ibiden-co-ltd-4062-q4-2026_F9Xvd7IW"),
        ("研調：Tom's Hardware ABF 2026（供需缺口、Ibiden 路線圖、競爭者擴產、ABF 膜漲價）", "https://www.tomshardware.com/tech-industry/semiconductors/the-state-of-abf-substrates-in-data-center-silicon-in-2026-solving-the-supply-crunch-and-material-wall-beneath-every-ai-accelerator"),
        ("研調：Data Gravity（高階 FC-BGA 份額、ABF 載板平均單價）", "https://www.datagravity.dev/p/the-abf-substrate-bottleneck"),
        ("研調：thebuildout（BofA：Ibiden 占 GPU 載板 70–80%、短缺到 2028）", "https://thebuildout.ai/intel/229c44f2-da57-4f46-b5fa-e5d03702236c"),
        ("研調：Encyclopedia of Stocks（欣興 2025 年 NVIDIA 約 30%）", "https://encyclopediaofstocks.substack.com/p/abf-substrate-market-ai-driven-super"),
        ("研調：Silicon Analysts（封裝成本，含載板）", "https://siliconanalysts.com/data/ai-chip-costs"),
        ("市場：stockanalysis（股價、前瞻與近 12 月 P/E）", "https://stockanalysis.com/quote/tyo/4062/"),
        ("市場層：markets/network、tsmc 成品（本 repo，唯讀）", ""),
        ("待補（人工）：十年 P/E 歷史（Bloomberg）；AI 伺服器載板營收（公司未揭露）；FY2027 以後共識（付費）；CB 發行額；ASIC 載板份額", "未找到"),
        ("SemiAnalysis", "未使用")]
for i, (a, b) in enumerate(srcs, 2): SO.cell(i, 1, a); SO.cell(i, 2, b)
SO.column_dimensions["A"].width = 90; SO.column_dimensions["B"].width = 140

# ---------------- README ----------------
RD = wb.create_sheet("README", 0)
lines = [("Ibiden（4062）收支與估值模型 v0.1.1（AI 半導體 Project：載板與 PCB 樣板）　2026-10-25", True),
 ("版本：v0.1（初版）。Ibiden 在 AI 加速器載板的有效金額與份額在本模型 Units 頁建立，為本專案權威資料（規則第 13 條）。", False),
 ("模型命題：Ibiden 的 AI 伺服器載板營收＝AI 加速器顆數（本專案網通模型）× Ibiden 份額 × 每顆載板金額（面積、層數逐年增加）× 短缺溢價；電子事業的非 AI 部分、陶瓷與其他事業另依成長率推。股價約 54–65 倍前瞻 EPS：2028 年產能擴到 2.5 倍、競爭者同時擴產之後，短缺溢價能否守住電子事業約 29% 的營益率？", False),
 ("核心問題：短缺溢價回落多少、何時回落，FY2028 EPS 會少多少；現價隱含的溢價路徑合不合理。", False), ("", False),
 ("會計年度 4 月至隔年 3 月：FY(t)＝t/4–(t+1)/3，對照 CY(t) 的加速器顆數（重疊 9 個月，不平移）。預測期 FY2027–FY2029（三年）。金額：十億日圓；EPS：日圓（2026-10-01 一拆二後）。", False), ("", False),
 ("驅動因子 → 推導量（全部唯讀上游，不修改）", True),
 ("AI 載板：markets/network 的加速器顆數（NVIDIA ＋ 權重 × 非 NVIDIA）× 每顆有效金額（FY2026 校準）× 金額指數（+18%／年，Analogy）× 份額相對值（Assumed）× 溢價指數（Assumed）。", False),
 ("利潤率：AI 載板營益率 FY2026 校準（非 AI 部分 15%，Assumed），之後＝ 1 −（1 − 校準值）÷ 溢價指數（成本不變）。陶瓷、其他事業營益率沿用 FY2026 預測。淨利 ÷ 營業利益 0.661（FY2026 預測）。", False),
 ("校準：FY2026 ＝ 公司 2026-08-04 上修後預測（營收 5,500 億、營業利益 1,270 億、淨利 840 億、EPS 149.68）；AI 伺服器載板占電子事業 65%（Assumed，公司未揭露）。", False),
 ("估值：現價 ÷ FY2026–FY2028 EPS 與 what must be true；十年 P/E 規則值待 Bloomberg（本版不計）；循環位置以本模型營收最高年判定。", False),
 ("標記：Verified、Interested-party、Analogy、Assumed、Derived。SemiAnalysis 未使用。", False),
 ("", False),
 ("已知限制：(1) AI 載板營收公司未揭露，以 Assumed 切分校準；(2) 溢價指數與份額相對值皆為 Assumed，是 FY2028 EPS 最敏感的輸入（Sens）；(3) 未設產能上限（2024 基準面積未找到）；(4) FY2026 錨點為公司預測，Q2 決算後更新；(5) 十年 P/E 規則值待 Bloomberg；(6) CB 轉換稀釋只在 Sens；(7) R4 產業部署路徑未建（Recon 預留列）。", False)]
lines.append(("v0.1.1（2026-10-25）：Val 頁新增「0. 目標價與上檔空間」與目標價用的十年 P/E（companiesmarketcap，Interested-party，單一來源）（放在 Val 既有內容之後，不位移既有儲存格），In 頁新增目標價情境權重 tp_wA／tp_wB／tp_wC（0.3／0.5／0.2，Assumed）。目標價＝十年 P/E 中位數 × 各情境 FY2027 EPS；其餘儲存格不變。", False))
for i, (t, b) in enumerate(lines, 1): RD.cell(i, 1, t).font = Font(bold=b)
RD.column_dimensions["A"].width = 170

# ---------------- Val 第 0 節：目標價與上檔空間（v0.1.1 新增；Andy 2026-10-25 要求，列標籤依工作單逐字） ----------------
# 為不位移既有儲存格（與舊版 dist 逐格相同），本節放在 Val 既有內容之後；HTML 估值章第一個區塊即為本節
def add_target_price(V, imp, note, fmt, method, px_note, t0=None):
    t0 = t0 or V.max_row + 2
    V.cell(t0, 1, "0. 目標價與上檔空間").font = BOLD
    hdr(V, t0 + 1, ["項目", "數值", "", "", "說明"])
    V.cell(t0 + 2, 1, "目標價（12 個月，情境加權）"); V[f"B{t0 + 2}"] = f"=tp_wA*B{t0 + 4}+tp_wB*B{t0 + 5}+tp_wC*B{t0 + 6}"; V[f"B{t0 + 2}"].number_format = fmt
    V[f"E{t0 + 2}"] = "權重 × 各情境隱含價"
    V.cell(t0 + 3, 1, "上檔空間（目標價 ÷ 現價 − 1）"); V[f"B{t0 + 3}"] = f"=B{t0 + 2}/px-1"; V[f"B{t0 + 3}"].number_format = "0.0%"; V[f"E{t0 + 3}"] = px_note
    for i, (s, lab) in enumerate((("A", "情境 A 隱含價（樂觀）"), ("B", "情境 B 隱含價（基準）"), ("C", "情境 C 隱含價（下檔）"))):
        rr = t0 + 4 + i
        V.cell(rr, 1, lab); V[f"B{rr}"] = imp[s]; V[f"B{rr}"].number_format = fmt; V[f"E{rr}"] = note[s]
    rr = t0 + 7
    V.cell(rr, 1, "情境權重 A／B／C"); V[f"B{rr}"] = "=tp_wA"; V[f"C{rr}"] = "=tp_wB"; V[f"D{rr}"] = "=tp_wC"
    for c in "BCD": V[f"{c}{rr}"].number_format = "0%"
    V[f"E{rr}"] = f'="合計 "&TEXT(B{rr}+C{rr}+D{rr},"0%")&"；In 頁 tp_wA／tp_wB／tp_wC（Assumed，Andy 可改）"'
    V.cell(t0 + 8, 1, "目標價方法"); V[f"B{t0 + 8}"] = method
    return t0
TP0 = V.max_row + 2
# 目標價用十年 P/E（本節下方）：companiesmarketcap 年底 P/E（2026-10-25 取得）；2017 為虧損（負值），不計入中位數
PE_CMC = [(2016, 44.1), (2017, -6.66), (2018, 35.3), (2019, 137.0), (2020, 55.2), (2021, 53.1), (2022, 40.7), (2023, 26.8), (2024, 57.2), (2025, 33.6)]
pr = TP0 + 10
V.cell(pr, 1, "0a. 十年 P/E（目標價用；companiesmarketcap 日曆年年底 P/E，Interested-party，資料商；第二來源：未找到）").font = BOLD
hdr(V, pr + 1, ["年底", "P/E", "計入中位數", "說明"])
for i, (y, v) in enumerate(PE_CMC):
    rw = pr + 2 + i
    V.cell(rw, 1, str(y)); V.cell(rw, 2, v).font = BLUE; V.cell(rw, 2).number_format = "0.0"
    V.cell(rw, 3, "否" if v <= 0 else "是")
    V.cell(rw, 4, "https://companiesmarketcap.com/ibiden/pe-ratio/（2026-10-25 讀取）" if i == 0 else ("虧損年，P/E 為負，不計" if v <= 0 else ""))
PE_MED_ROW = pr + 12
V.cell(PE_MED_ROW, 1, "中位數（計入者）"); V[f"B{PE_MED_ROW}"] = "=MEDIAN(" + ",".join(f"B{pr + 2 + i}" for i, (y, v) in enumerate(PE_CMC) if v > 0) + ")"
V[f"B{PE_MED_ROW}"].number_format = "0.0"
add_target_price(V, {s: f"=$B${PE_MED_ROW}*PL!E{SUMR[(s, 'eps')]}" for s in SC},
    {s: f"情境 {s} FY2027 EPS（PL 頁）× 十年 P/E 中位數（第 0a 節）" for s in SC}, "#,##0",
    "P/E：十年（2016–2025 日曆年年底）P/E 中位數（companiesmarketcap 單一來源，第 0a 節；2017 年虧損不計，9 年）× 各情境 FY2027 EPS（FY2027＝2027-04～2028-03）。"
    "本模型無規則值（十年 P/E 待 Bloomberg），依工作單改用此備案；倍數為日曆年年底口徑，與 3 月制會計年度不完全對齊。十年倍數多為 AI 載板需求前的時期（2019 年獲利低谷約 137x 為極值，中位數受影響小），偏誤方向不明確；"
    "FY2027 三情境 EPS 相同（情境差從 FY2028 起），故三情境隱含價相同；隱含價約為 FY2027 年底（2028-03）的價格、未折現，當作 12 個月目標價時偏高。",
    "現價＝In 頁 px（2026-10-16 收盤，分割後）", t0=TP0)

order =["README", "In", "Q", "Units", "PL", "Val", "Recon", "Sens", "Sources"]
wb._sheets = [wb[n] for n in order]
wb.save(sys.argv[1])
if os.environ.get("IBIDEN_ROWMAP"):   # 除錯用：輸出列號對照（供撰寫 html.json）；CI 不設此變數
    import json
    json.dump({"PB": PB, "SUMR": {f"{k[0]}|{k[1]}": v for k, v in SUMR.items()}, "UR": UR, "CHK": CHK, "IMP": IMP, "IMPH": IMPH, "SB": list(SB), "SENS_LAST": SENS_LAST,
               "SENS_NOTE": SENS_NOTE, "RECON_LAST": RECON_LAST, "WMT_NOTE": WMT_NOTE, "UNOTE": UNOTE, "QR": QR},
              open(os.environ["IBIDEN_ROWMAP"], "w"), ensure_ascii=False, indent=1)
