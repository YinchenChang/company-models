# Nabtesco（ナブテスコ，6268.T）個股模型 v0.1 — AI實體應用 Project（線 2a 工業機器人：減速機）
# 用法：python3 physical_ai/nabtesco/build_nabtesco.py            → 建置、整本重算 Sens 與反解、輸出 dist/ 成品與 README.md
#       python3 physical_ai/nabtesco/build_nabtesco.py --raw OUT  → 只寫一本（讀環境變數 NAB_OVERRIDE 的 JSON 覆寫 In 頁基準值），供重算
# 會計年度＝曆年（12 月決算，IFRS）。金額：十億日圓（¥B）；股數：百萬股；EPS、股價：日圓。
# 結構：README、In、Q（季度分部，含分部對帳）、Hist、PL（三情境：命題模組→分部→損益→現金流）、DCF、Val、Recon、Sens、對帳、Risk、Sources。
# 不讀 Siliconomics／Tokenomics／Wattonomics 協調檔（實體 AI 線；規格中 coord 部分略過）。DCF 為本檔自寫（未用 tools/fcff_dcf.py），公式同 docs/plan/個股DCF標準規格.md。
import sys, os, json, subprocess, shutil, hashlib, datetime, tempfile
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.workbook.defined_name import DefinedName

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(os.path.dirname(HERE))
VER = "v0.1"
DATE = "20261116"
NAME = f"{DATE}_Nabtesco個股模型_{VER}.xlsx"
BLUE = Font(color="0000FF"); BOLD = Font(bold=True)
HDR = PatternFill("solid", fgColor="DDEBF7"); KEY = PatternFill("solid", fgColor="FFF2CC")
SC = ("A", "B", "C"); SCN = {"A": "樂觀", "B": "基準", "C": "保守"}
YRS = list(range(2025, 2033)); COL = {y: "CDEFGHIJ"[i] for i, y in enumerate(YRS)}   # PL：C＝FY2025 … J＝FY2032

# ---------------- 出處 ----------------
TS3Q = "公司 2026年12月期 第3四半期決算短信〔IFRS〕(連結)（2026-11-13）https://www.nabtesco.com/cms/wp-content/uploads/Summary_of_Consolidated_Financial_Statements_for_FY2026_Q3.pdf"
TS2Q = "公司 2026年12月期 中間期決算短信（訂正版，2026-09-30 訂正後再訂正）https://www.nabtesco.com/cms/wp-content/uploads/Revision-of-Summary_of_Consolidated_Financial_Statements_for_FY2026_Q2_2-.pdf"
TS2QX = "JPX 適時開示 iXBRL 中間期サマリー（2026-08-14 訂正）https://www2.jpx.co.jp/disc/62680/081220260814521255_tse-scedifsy-62680-20260814362680-ixbrl.htm"
TS25 = "公司 2025年12月期 決算短信サマリー（2026-02-12；JPX iXBRL）https://www2.jpx.co.jp/disc/62680/081220260212556894_tse-acedifsm-62680-20260212362680-ixbrl.htm"
BR3Q = "公司 2026年12月期 第3四半期 決算説明会資料（2026-11-13）https://www.nabtesco.com/cms/wp-content/uploads/Results_Briefing_Material_for_FY2026_Q3.pdf"
BR2Q = "公司 2026年12月期 第2四半期 決算説明会資料（スクリプト付き，2026-08-05）https://www.nabtesco.com/cms/wp-content/uploads/Results_Briefing_Material_for_FY2026_Q2_script.pdf"
QA2Q = "公司 2026年12月期 第2四半期 決算説明会 Q&A https://www.nabtesco.com/cms/wp-content/uploads/Results_Briefing_QA_for_FY2026_Q2.pdf"
BR1Q = "公司 2026年12月期 第1四半期 決算説明会資料（2026-04-30）https://www.nabtesco.com/cms/wp-content/uploads/Results_Briefing_Material_for_FY2026_Q1.pdf"
REF = "公司 セグメント別参考データ（2026-07-31 修正）https://www.nabtesco.com/cms/wp-content/uploads/Referencial_data_for_sales_as_of_20260731.pdf"
ENV = "envalith.com 分部頁（轉載第 1 四半期決算短信分部表；百萬日圓）https://envalith.com/en/companies/6268/segments"
SA = "stockanalysis.com（資料商；讀取 2026-11-16）"
EDN = "EDINET 自己株券買付状況報告書（2026-10-15 提出；報告期間 2026-09）https://disclosure2dl.edinet-fsa.go.jp/searchdocument/pdf/S100Z80C.pdf"
TRD = "トレーダーズ・ウェブ 6268 時系列 https://www.traders.co.jp/stocks/61_6268/historical"
JGB10 = "日本財務省 國債金利情報 10 年：2026-11-02 2.817%（與 AI半導體 東京精密 v0.3.6 同一取值，跨模型一致）"
CO = "發行公司（立場：自家業績與前景）"
DV = "資料商彙整賣方共識（立場：中立彙整，但來源為賣方）"
L2A = "/mnt/user-data/outputs/lines/線2a_工業機器人_框架_v0.1.md"
L2B = "/mnt/user-data/outputs/lines/線2b_人形機器人_框架_v0.1.md"

# ---------------- In 列 ----------------
# (鍵, 項目, 單位, 低, 基準, 高, 標記, 來源1, 來源2／說明, 來源誘因)
ROWS = []
def sec(t): ROWS.append(("sec", t))
def inp(k, lab, u, lo, b, hi, tag, s1, s2="", inc=""): ROWS.append((k, lab, u, lo, b, hi, tag, s1, s2, inc))

sec("A. FY2025 實績（2025-01～12；持續營業口徑，油壓機器事業已列停業單位）")
inp("a25_rev", "合併營收 FY2025", "¥B", None, 307.912, None, "Verified", TS25 + "：売上収益 307,912 百萬日圓（+9.8%）", "FY2024 重編（持續營業）280,458")
inp("a25_op", "營業利益 FY2025", "¥B", None, 20.726, None, "Verified", TS25 + "：20,726（+60.3%）", "FY2024 12,933")
inp("a25_ni", "歸屬母公司淨利 FY2025（含停業單位）", "¥B", None, 15.695, None, "Verified", TS25 + "：15,695；基本 EPS 131.56", "2026-09-30 公司公告過年度訂正（持分法 TMT マシナリー），未取得訂正後 FY2025 淨利；本列為原公告值")
inp("a25_cmp", "コンポーネントソリューション（CMP）營收 FY2025", "¥B", None, 79.3, None, "Verified", BR1Q + "：10 年圖表 2025/12 通期実績 793 億、營業利益 54 億", "億日圓取整（決算短信分部注記原文未取得）")
inp("a25_trs", "トランスポートソリューション（TRS）營收 FY2025", "¥B", None, 100.5, None, "Verified", BR1Q + "：1,005 億、營業利益 136 億", "同上")
inp("a25_acb", "アクセシビリティソリューション（ACB）營收 FY2025", "¥B", None, 110.7, None, "Verified", BR1Q + "：1,107 億、營業利益 91 億", "同上")
inp("a25_mfr", "その他（マニュファクチャリングソリューション，MFR）營收 FY2025", "¥B", None, 17.4, None, "Verified", BR1Q + "：174 億、營業利益 22 億", "同上")

sec("B. FY2026 錨點：公司指引（2026-07-31 上修，2026-11-13 維持；Interested-party）")
inp("g26_rev", "公司指引：FY2026 營收", "¥B", None, 344.0, None, "Interested-party", TS3Q + "：344,000（+11.7%），第 3 季維持", "5 月原指引 327,000（" + TS25 + "）。9M 實績 250.813 → Q4 隱含 93.2（Derived）", CO)
inp("g26_op", "公司指引：FY2026 營業利益", "¥B", None, 32.6, None, "Interested-party", TS3Q + "：32,600（+55.6%）", "9M 實績 23.352 → Q4 隱含 9.2（Derived）", CO)
inp("g26_ni", "公司指引：FY2026 歸屬母公司淨利", "¥B", None, 24.1, None, "Interested-party", TS3Q + "：24,100；EPS 206.20；稅前 36,300", "含停業單位（9M 停業單位利益 1.359，含 70% 股權轉讓的換算差額）", CO)
inp("g26_eps", "公司指引：FY2026 基本 EPS", "¥", None, 206.20, None, "Interested-party", TS3Q + "：206.20", "", CO)
inp("g26_prg", "公司指引：FY2026 精密減速機營收", "¥B", None, 94.6, None, "Interested-party", REF + "：精密減速機 946 億（4/30 版 882 億）；上半實績 451 億", BR3Q + "：9M 精密減速機營收約 692 億（+122 億）→ Q4 隱含約 254 億（Derived）", CO)
inp("g26_cmpo", "公司指引：FY2026 CMP 其他（非精密減速機）營收", "¥B", None, 0.7, None, "Interested-party", REF + "：CMP その他 7 億（CMP 合計 953 億）", "", CO)
inp("g26_cmpop", "公司指引：FY2026 CMP 營業利益", "¥B", None, 9.7, None, "Interested-party", BR2Q + "：CMP 營業利益 97 億（10.2%）；4/30 版 70 億", BR3Q + "：9M 實績 74 億（進度 77%）", CO)
inp("g26_trs", "公司指引：FY2026 TRS 營收", "¥B", None, 112.7, None, "Interested-party", REF + "：1,127 億（鐵道 282、商用車 135、航空 304、船舶 312、其他 94）", "", CO)
inp("g26_trsop", "公司指引：FY2026 TRS 營業利益", "¥B", None, 19.4, None, "Interested-party", BR2Q + "：194 億（17.2%）", "", CO)
inp("g26_acb", "公司指引：FY2026 ACB 營收", "¥B", None, 117.1, None, "Interested-party", REF + "：自動門 1,171 億", "", CO)
inp("g26_acbop", "公司指引：FY2026 ACB 營業利益", "¥B", None, 10.7, None, "Interested-party", BR2Q + "：107 億（9.1%）", "", CO)
inp("g26_mfr", "公司指引：FY2026 MFR（その他）營收", "¥B", None, 18.9, None, "Interested-party", REF + "：189 億（包裝機 186）", "", CO)
inp("g26_mfrop", "公司指引：FY2026 MFR 營業利益", "¥B", None, 2.6, None, "Interested-party", BR2Q + "：26 億（13.8%）；9M 實績 11.65 億（Q4 須 14 億，包裝機 Q4 集中）", "", CO)
inp("corp_26", "FY2026 全社費用與消去（營業利益）", "¥B", None, "=g26_op-(g26_cmpop+g26_trsop+g26_acbop+g26_mfrop)", None, "Derived", "＝公司指引營業利益 − 四分部指引營業利益（約 −9.8；簡報約 −98 億）", "FY2027 起隨合併營收等比例")
inp("nonop_26", "FY2026 營業外（稅前 − 營業利益；持分法、金融損益）", "¥B", None, "=36.3-g26_op", None, "Derived", TS3Q + "：指引稅前 36,300 − 營業利益 32,600", "9M 實績 24.763 − 23.352 ＝ 1.411")
inp("sh_26", "FY2026 平均股數（指引 EPS 隱含）", "百萬股", None, "=g26_ni/g26_eps*1000", None, "Derived", "＝24,100 百萬 ÷ 206.20 ＝ 116.9 百萬股", "9M 期中平均 116.922（" + TS3Q + "）")
inp("cx_26", "公司計畫：FY2026 資本支出", "¥B", None, 13.0, None, "Interested-party", BR2Q + "：FY2026 設備投資計畫 130 億（7/31 版；上半實績 42 億）", "FY2027 起用 fc_cx", CO)
inp("g_anc", "FY2026 實績相對公司指引（營收、營業利益同比例）", "%", -0.02, 0.0, 0.02, "Assumed", "基準＝公司指引（9M 已實績，偏離只來自 Q4）", "低／高 ±2%：Q4 隱含營收 93.2 vs Q3 單季 83.4，季節性高（MFR 包裝機、ACB 年末）")

sec("C. 命題模組：精密減速機（RV）＝機器人用 × 全球工業機器人裝機 × 份額 ＋ 非機器人 ＋ RV mini／實體 AI 新品")
inp("r_rob", "精密減速機營收中「工業機器人用」比例（FY2026）", "%", 0.60, 0.70, 0.80, "Assumed", "未取得：公司不揭露用途別金額（簡報只有定性：產業機器人、工具機、半導體設備、一般產業）", "區間 60–80%：傳統 6 軸機器人底座與大臂用 RV（線 2a 第 3 節物理層註記）")
inp("pai_0", "RV mini 等新品營收 FY2026", "¥B", None, 0.2, None, "Interested-party", QA2Q + "：RV mini 下期出貨規模「2 億円弱」，先用於 6 軸機器人腕部，次為協作與工具機", "已含在 g26_prg 內（本模型自 g26_prg 扣出再分機器人／非機器人）", CO)
for s, v, lo, hi, note in (
        ("A", 0.15, None, None, "樂觀：JARA 受注 2026 年 7–9 月 +46.6%、生產 +26.2%（訂單÷生產約 1.24），傳導到 2027 年裝機；公司 9M 精密減速機受注 72.9 vs 59.0（+24%）"),
        ("B", 0.072, 0.03, 0.12, "基準：IFR 預測 2026 年 65.5 萬台 → 2029 年 80.6 萬台，年複合 7.2%（線 2a 第 1 節；IFR World Robotics 2026）"),
        ("C", -0.05, None, None, "保守：2026 年補庫存（9M 精密減速機營收 +21%，高於 IFR 2026 年 +8.6%）後，2027 年循環回落；2023–2024 年精密減速機下滑為類比")):
    inp(f"rg27_{s}", f"情境 {s}：全球工業機器人需求年增 FY2027（機器人用 RV 量）", "%", lo, v, hi, "Assumed" if s != "B" else "Analogy", note,
        "區間見情境 A／B／C 三列（A 15%、C −5%）" if s != "B" else "類比對象：IFR 全球裝機預測（台數；與 RV 營收的差異：價格、份額、在製品庫存）")
for s, v, lo, hi in (("A", 0.09, None, None), ("B", 0.072, 0.04, 0.10), ("C", 0.04, None, None)):
    inp(f"rgm_{s}", f"情境 {s}：全球工業機器人需求年增 FY2028–FY2029", "%/年", lo, v, hi, "Analogy" if s == "B" else "Assumed",
        "IFR 2026→2029 年複合 7.2%（Verified 預測值，Interested-party：IFR 為產業協會）" if s == "B" else "A／C 以 B ± 約 2–3pt 設定", "區間見情境 A–C")
for s, v, lo, hi in (("A", 0.07, None, None), ("B", 0.05, 0.03, 0.07), ("C", 0.03, None, None)):
    inp(f"rglt_{s}", f"情境 {s}：全球工業機器人需求年增 FY2030–FY2032", "%/年", lo, v, hi, "Assumed", "IFR 預測只到 2029；長期以 3–7% 區間", "")
for s, v, lo, hi in (("A", 0.0, None, None), ("B", -0.01, -0.03, 0.0), ("C", -0.03, None, None)):
    inp(f"shr_{s}", f"情境 {s}：機器人用 RV 份額＋價格年變化（中國 RV 廠替代）", "%/年", lo, v, hi, "Assumed",
        "公司自稱中大型機器人關節用 RV 全球份額約 6 成（公司產品頁，Interested-party）；中國國產比 2025 年 55%（IFR，線 2a）", "陸廠：雙環傳動 002472、中大力德 002896 等（線 2a 第 2 節）；份額數據未取得")
for s, v, lo, hi in (("A", 0.06, None, None), ("B", 0.04, 0.01, 0.06), ("C", 0.01, None, None)):
    inp(f"nr_{s}", f"情境 {s}：精密減速機非機器人用（工具機、半導體設備、一般產業）年增", "%/年", lo, v, hi, "Assumed",
        BR3Q + "：工具機、半導體設備需求堅實（定性）", QA2Q + "：一般產業「種まき」開始接單（定性）")
for s, v, lo, hi in (("A", 10.0, None, None), ("B", 3.0, 1.0, 6.0), ("C", 0.5, None, None)):
    inp(f"pai_{s}", f"情境 {s}：RV mini／實體 AI（含人形、協作）新品營收 FY2032", "¥B", lo, v, hi, "Assumed",
        "FY2026 約 0.2（公司）；線性增加到 FY2032。人形主流 BOM 用諧波＋行星，不用 RV（線 2b 第 3.2 節），故本列以協作、腕部、AMR 等小型關節為主", "A 含人形大關節（髖、膝）採用 RV mini 的選擇權；無公開訂單，純假設")
inp("mm_cmp", "CMP 邊際營業利益率（ΔCMP 營業利益 ÷ ΔCMP 營收）", "x", 0.22, 0.29, 0.35, "Derived",
    "9M：(7.425 − 3.895) ÷ (69.780 − 57.519) ＝ 28.8%（" + TS3Q + "）", "上半：(4.841 − 1.856) ÷ (45.521 − 36.821) ＝ 34.3%；低 22%（Assumed：含價格轉嫁 5 億的一次性）")
for s, v, lo, hi in (("A", 0.0, None, None), ("B", 0.0, -0.02, 0.0), ("C", -0.02, None, None)):
    inp(f"madj_{s}", f"情境 {s}：FY2028 起 CMP 營業利益率額外調整（價格戰、稼動）", "pt", lo, v, hi, "Assumed",
        "C：中國 RV 價格競爭（陸製諧波較哈默納科低 40–60%，線 2a；RV 價差未取得，類比）", "CMP 營業利益率史：2021 16.6%、2022 11.3%、2023 7.5%、2024 3.9%、2025 6.8%（含油壓，" + BR2Q + "）")

sec("D. 其他分部（官方分部；成長路徑）")
for s, v, lo, hi in (("A", 0.07, None, None), ("B", 0.05, 0.03, 0.07), ("C", 0.02, None, None)):
    inp(f"gtrs_{s}", f"情境 {s}：TRS 營收年增 FY2027 起（鐵道、航空、船舶、商用車）", "%/年", lo, v, hi, "Assumed",
        "Q1 TRS 受注 391 億 vs 275 億（+42%，" + BR1Q + "）；航空 777X 2027 年上市擴產（" + QA2Q + "）", "FY2025→FY2026 指引 +12.1%（Derived）")
for s, v, lo, hi in (("A", 0.05, None, None), ("B", 0.035, 0.02, 0.05), ("C", 0.015, None, None)):
    inp(f"gacb_{s}", f"情境 {s}：ACB（自動門、月台門）營收年增 FY2027 起", "%/年", lo, v, hi, "Assumed",
        "FY2025→FY2026 指引 +5.8%（Derived）；含匯率（歐洲比重高）", "")
for s, v, lo, hi in (("A", 0.05, None, None), ("B", 0.03, 0.0, 0.05), ("C", 0.0, None, None)):
    inp(f"gmfr_{s}", f"情境 {s}：MFR（包裝機等）營收年增 FY2027 起", "%/年", lo, v, hi, "Assumed", "FY2025→FY2026 指引 +8.6%（Derived）", "")
inp("m_trs", "TRS 營業利益率（FY2026 指引，全期持平）", "%", 0.15, "=g26_trsop/g26_trs", 0.19, "Derived", "＝194 ÷ 1,127 ＝ 17.2%", "9M 實績 18.0%；FY2025 13.5%（含鐵道關係會社整理損失）；低／高 Assumed")
inp("m_acb", "ACB 營業利益率（FY2026 指引，全期持平）", "%", 0.075, "=g26_acbop/g26_acb", 0.105, "Derived", "＝107 ÷ 1,171 ＝ 9.1%", "9M 實績 8.3%；FY2025 8.2%")
inp("m_mfr", "MFR 營業利益率（FY2026 指引，全期持平）", "%", 0.09, "=g26_mfrop/g26_mfr", 0.15, "Derived", "＝26 ÷ 189 ＝ 13.8%", "9M 實績 8.8%；FY2025 12.6%")

sec("E. 損益表下半段")
inp("nonop", "營業外（持分法 TMT マシナリー 33%、コムテスコ 30%＋金融）FY2027 起", "¥B", 1.5, 2.5, 3.5, "Assumed", "9M 實績 1.411、FY2026 指引隱含 3.7（含一次性）", "持分法投資帳面 33.971（2026-09-30）")
inp("tax_w", "有效稅率", "%", 0.28, 0.30, 0.32, "Analogy", "日本法定實效稅率約 30.6%（類比：法定稅率，未含海外子公司稅率差）", "FY2026 指引隱含（36.3 − 24.1 − 少數）÷ 36.3 約 31%（Derived，含停業單位）")
inp("nci_r", "少數股東損益 ÷ 稅後淨利", "%", None, 0.05, None, "Derived", TS3Q + "：9M 稅後 17.820、歸屬母公司 16.956 → 4.8%", "停業單位為主的 9M 外，取整 5%")
inp("pay_r", "FY2027 起配息率（股利 ÷ 歸屬母公司淨利）", "%", None, 0.40, None, "Interested-party", BR2Q + "：FY2026 配息率 39.8%（DOE 3.5% 目標）", "只用於收支關卡；不進 DCF", CO)
inp("dps_26", "FY2026 每股股利", "¥", None, 82, None, "Interested-party", TS3Q + "：年間 82（中間 41 實績＋期末 41 予想）", "", CO)

sec("F. 估值輸入（評價日 2026-12-31；資產負債表 2026-09-30）")
inp("px", "股價（2026-11-13 收盤）", "¥", None, 4149, None, "Verified", TRD + "：2026-11-13 終値 4,149", SA + "：4,149（2026-11-13 15:30 JST）。註：2026-11-16 盤中 3,651（−12.0%，第 3 季決算後），非收盤，未採用")
inp("px_date", "股價基準日", "日期", None, "=DATE(2026,11,13)", None, "Verified", "同上", "")
inp("tp_date", "目標價時點（股價日 ＋ 12 個月）", "日期", None, "=EDATE(px_date,12)", None, "Derived", "個股模型原則第 10 節第 26 條", "")
inp("sh_out", "流通股數（2026-09-30；發行 118,064,699 − 自己株式 2,711,630）", "百萬股", None, "=(118.064699-2.71163)", None, "Derived", TS3Q + "：期末發行済株式数、自己株式数（Verified）", "RSU／選擇權：公司以 BBT（株式給付信託）方式，股份已在自己株式內；無選擇權揭露 → 不另加")
inp("bb_rem", "評價日前尚待執行的庫藏股買回金額", "¥B", 0.0, "=15-8.5775835", None, "Derived", EDN + "：截至 2026-09-30 累計 1,861,100 股、8,577,583,500 円；上限 150 億円或 400 萬股、期間 2026-08-03～11-30", "Assumed：剩餘額度以現價全數執行（低案＝不再買回）。已買回股份 2026-12-15 註銷")
inp("sh_v", "評價日股數", "百萬股", None, "=sh_out-bb_rem/px*1000", None, "Derived", "＝流通股數 − 尚待買回金額 ÷ 股價", "DCF 標準規格第 3 節")
inp("cash_v", "現金及約當現金（2026-09-30）", "¥B", None, 85.115, None, "Verified", TS3Q + "：85,115", "")
inp("debt_v", "借款（2026-09-30；流動 43,297＋非流動 26）", "¥B", None, 43.323, None, "Verified", TS3Q, "")
inp("lease_v", "租賃負債（2026-09-30；流動 3,514＋非流動 9,904）", "¥B", None, 13.418, None, "Verified", TS3Q, "")
inp("inv_v", "持分法投資（2026-09-30）", "¥B", None, 33.971, None, "Verified", TS3Q + "：33,971（2025 年底 24,976；增加主因コムテスコ 30% 轉為持分法）", "")
inp("inv_k", "持分法投資計入比例", "x", 0.5, 1.0, None, "Assumed", "基準按帳面全額；低案 50%（TMT 訂正事項、コムテスコ 剩餘 30% 處分價值不確定）", "")
inp("mi_v", "非支配股東權益（2026-09-30）", "¥B", None, 12.049, None, "Verified", TS3Q, "")
inp("stub_frac", "Stub 期（2026-09-30 → 評價日 2026-12-31）", "年", None, "=(DATE(2026,12,31)-DATE(2026,9,30))/365", None, "Derived", "", "Stub FCFF＝FY2026 FCFF × stub_frac")
inp("rf", "無風險利率（日本 10 年期國債）", "%", None, 0.02817, None, "Verified", JGB10, "")
inp("erp", "股權風險溢酬", "%", 0.04, 0.05, 0.06, "Assumed", "DCF 標準規格第 2 節：預設 5%（敏感度 4–6%）", "")
inp("crp", "日本國家風險溢酬", "%", None, 0.0091, None, "Verified", "Damodaran 國家風險溢酬（2026-01-05）：日本 0.91%（與東京精密 v0.3.6 同值）", "")
inp("beta", "β（兩來源 Blume 調整後中位數）", "x", "=E{r}-0.3", "=MEDIAN(0.67*0.93+0.33,0.67*0.97+0.33)", "=E{r}+0.3", "Derived",
    SA + "：Nabtesco 5 年 β 0.93；同業 8 家 5 年 β 中位數 0.97（Fanuc 0.98、安川 1.24、THK 0.76、哈默納科 0.96、SMC 0.83、Nidec 1.07、Knorr-Bremse 1.11、ASSA ABLOY 0.82）", "Blume 0.67β＋0.33；敏感度 ±0.3（Assumed）")
inp("kd", "稅前負債成本", "%", None, "=rf+0.005", None, "Assumed", "rf ＋ 0.5pt（短期銀行借款為主）", "")
inp("g_cap", "永續成長率上限", "%", 0.02, 0.03, 0.04, "Assumed", "DCF 標準規格：3%（敏感度 2–4%），且不超過 rf → g＝MIN(g_cap, rf)", "")
inp("ronic_sp", "RONIC − WACC", "pt", None, 0.0, 0.05, "Assumed", "預設 RONIC＝WACC（新投資價值中性）；高案 +5pt", "")
inp("wacc_adj", "WACC 調整（只用於 Sens ±1pt）", "pt", -0.01, 0.0, 0.01, "Assumed", "", "")
inp("fc_da", "折舊攤銷 ÷ 營收", "%", 0.042, "=16.5/344", 0.055, "Interested-party", BR2Q + "：FY2026 計畫折舊 165 億 ÷ 營收 3,440 億", "FY2025 17.258 ÷ 307.912 ＝ 5.6%（含油壓，" + SA + "）", CO)
inp("fc_cx", "資本支出 ÷ 營收", "%", 0.04, 0.05, 0.065, "Assumed", "FY2026 計畫 130 億（3.8%，" + BR2Q + "）；FY2025 13.201、FY2024 24.898（" + SA + "，含油壓）", "精密減速機產能 115 萬台／年、津 85%、常州 95%（" + BR3Q + "）→ 量增超過約一成需擴產；基準 5% 高於折舊")
inp("fc_wc", "營運資金增加 ÷ 營收增量", "%", 0.20, 0.30, 0.35, "Derived", "（應收＋存貨 − 應付）÷ 營收：FY2025 (76.945＋52.838 − 45.161) ÷ 307.912 ＝ 27.5%；FY2024 (90.293＋53.387 − 51.178) ÷ 280.458 ＝ 33.0%（" + SA + "）", "取兩年平均約 30%")
inp("dps_tgt", "評價日到目標時點之每股股利", "¥", None, "=dps_26", None, "Assumed", "FY2026 期末 41（2027-03 支付）＋FY2027 中間（假設同 41）", "")
inp("t_tgt", "評價日到目標時點（年）", "年", None, 1.0, None, "Derived", "2026-12-31 → 2027-12-31", "")
inp("w_dcf", "DCF 權重（其餘為分部本益比加總）", "x", 0.3, 0.5, 0.7, "Assumed", "個股模型原則第 4 節預設各半", "")
inp("rel_k", "相對倍數調整係數", "x", None, 1.0, None, "Assumed", "1.00（未調整：缺 Nabtesco 對同業 10 年前瞻 P/E 比值歷史）", "第 10 節第 18、47 條：列為承重輸入")
for s, v in (("A", 0.25), ("B", 0.50), ("C", 0.25)):
    inp(f"w_{s}", f"情境權重 {s}", "x", None, v, None, "Assumed", "DCF 標準規格第 3 節預設 25／50／25", "")

sec("G. 同業前瞻 P/E（資料商頁面日期各異，未換算到同一日；Interested-party：賣方共識）")
PEERS = [("pe_fanuc", "Fanuc 6954（機器人、FA）", 23.22, "2026-11-13", "CMP"),
         ("pe_yask", "安川電機 6506（伺服、機器人）", 21.49, "2026-09-16", "CMP"),
         ("pe_hds", "哈默納科 6324（諧波減速機）", 84.44, "2026-10-02", "CMP"),
         ("pe_thk", "THK 6481（線性傳動）", 15.37, "2026-10-19", "CMP"),
         ("pe_smc", "SMC 6273（氣動元件）", 21.35, "2026-10-09", "CMP"),
         ("pe_nidec", "Nidec 6594（馬達）", 12.67, "2026-10-16", "CMP"),
         ("pe_kbx", "Knorr-Bremse KBX（鐵道與商用車煞車）", 20.01, "2026-09-30", "TRS"),
         ("pe_khi", "川崎重工 7012（鐵道、航空、船舶）", 15.44, "2026-10-09", "TRS"),
         ("pe_assa", "ASSA ABLOY ASSA-B（自動門、門禁）", 20.45, "2026-10-20", "ACB"),
         ("pe_doka", "dormakaba DOKA（自動門、門禁）", 16.75, "2026-08-28", "ACB")]
for k, lab, v, d, g in PEERS:
    inp(k, f"前瞻 P/E：{lab}", "x", None, v, None, "Interested-party", f"{SA}：{d} 頁面值", f"群組 {g}；會計口徑：日本公司 IFRS 或日本基準、歐洲 IFRS（租賃資本化），P/E 為股權口徑不受租賃處理影響", DV)
inp("pe_cmp", "CMP 群組前瞻 P/E 中位數（6 家）", "x", None, "=MEDIAN(pe_fanuc,pe_yask,pe_hds,pe_thk,pe_smc,pe_nidec)", None, "Derived", "少於 8 家，不做迴歸（第 10 節第 7 條）", "")
inp("pe_trs", "TRS 群組前瞻 P/E 中位數（2 家）", "x", None, "=MEDIAN(pe_kbx,pe_khi)", None, "Derived", "", "")
inp("pe_acb", "ACB 群組前瞻 P/E 中位數（2 家）", "x", None, "=MEDIAN(pe_assa,pe_doka)", None, "Derived", "", "")
inp("pe_all", "全體同業前瞻 P/E 中位數（10 家；MFR、全社費用、營業外用）", "x", None, "=MEDIAN(pe_fanuc,pe_yask,pe_hds,pe_thk,pe_smc,pe_nidec,pe_kbx,pe_khi,pe_assa,pe_doka)", None, "Derived", "", "")
inp("pe_self", "Nabtesco 現行前瞻 P/E（只作對照，不進模型）", "x", None, 18.40, None, "Interested-party", SA + "：2026-11-13 forward P/E 18.40", "隱含 NTM EPS ＝ 4,149 ÷ 18.40（Derived）", DV)
inp("cons_tp", "賣方平均目標價（只作對照，不進模型）", "¥", None, 5300.91, None, "Interested-party", SA + "：10 位分析師平均 5,300.91（2026-11-13 頁面）；評等 Hold", "", DV)
inp("cons_rev26", "共識 FY2026 營收（只作對照，不進模型）", "¥B", None, 347.0, None, "Interested-party", SA + " forecast 頁（2026-09-08）：平均 347.0（區間 344.7–350.0，8 位）", "財報前（第 3 季決算前），只作歷史對照", DV)


def build(path, ov=None, sens=None, rev=None):
    ov = ov or {}
    wb = openpyxl.Workbook()
    def dn(name, ref): wb.defined_names[name] = DefinedName(name, attr_text=ref)
    def hdr(ws, row, labels):
        for j, h in enumerate(labels, 1):
            c = ws.cell(row, j, h); c.font = BOLD; c.fill = HDR
    RD = wb.active; RD.title = "README"
    IN = wb.create_sheet("In")
    IN["A1"] = "In — 輸入（藍字）。基準欄 E 進主模型；D「敏感度低」、F「敏感度高」只用於 Sens 頁（一次動一個參數，整本重算），不進目標價。金額 ¥B、股數百萬股、EPS 與股價日圓"
    IN["A2"] = "標記：Verified（附出處）／Interested-party（寫明是誰、誘因欄）／Analogy（寫明類比對象）／Assumed（給區間）／Derived（寫明換算）；查不到寫「未取得」"
    hdr(IN, 3, ["鍵", "項目", "單位", "敏感度低", "基準", "敏感度高", "標記", "來源1", "來源2／說明", "來源誘因"])
    r = 4; KR = {}
    for row in ROWS:
        if row[0] == "sec":
            IN.cell(r, 1, row[1]).font = BOLD; r += 1; continue
        k, lab, u, lo, b, hi, tag, s1, s2, inc = row
        if k in ov: b = ov[k]
        IN.cell(r, 1, k); IN.cell(r, 2, lab); IN.cell(r, 3, u)
        for col, v in ((4, lo), (5, b), (6, hi)):
            if v is None: continue
            if isinstance(v, str): v = v.replace("{r}", str(r))
            c = IN.cell(r, col, v)
            if col == 5: c.font = BLUE
        IN.cell(r, 7, tag); IN.cell(r, 8, s1); IN.cell(r, 9, s2); IN.cell(r, 10, inc)
        dn(k, f"In!$E${r}"); KR[k] = r; r += 1
    # 情境驅動對照
    r += 1; IN.cell(r, 1, "情境驅動對照（純展示；修改請改上方藍字）").font = BOLD; r += 1
    hdr(IN, r, ["鍵", "驅動", "", "情境 A（樂觀）", "情境 B（基準）", "情境 C（保守）", "", "說明"]); r += 1
    FAM = [("rg27", "全球工業機器人需求年增 FY2027"), ("rgm", "同上 FY2028–FY2029"), ("rglt", "同上 FY2030–FY2032"), ("shr", "機器人用 RV 份額＋價格年變化"),
           ("nr", "非機器人用年增"), ("pai", "RV mini／實體 AI 新品營收 FY2032"), ("madj", "FY2028 起 CMP 營業利益率調整"),
           ("gtrs", "TRS 年增"), ("gacb", "ACB 年增"), ("gmfr", "MFR 年增")]
    for f, lab in FAM:
        IN.cell(r, 1, f + "_scn"); IN.cell(r, 2, lab)
        for j, s in enumerate(SC): IN.cell(r, 4 + j, f"=E{KR[f + '_' + s]}")
        IN.cell(r, 8, f"引用 In 頁 {f}_A／{f}_B／{f}_C；{lab}"); r += 1
    IN.cell(r, 1, "scn_w"); IN.cell(r, 2, "情境權重")
    for j, s in enumerate(SC): IN.cell(r, 4 + j, f"=E{KR['w_' + s]}")
    IN.cell(r, 8, "引用 In 頁 w_A／w_B／w_C；合計＝1"); r += 1
    IN.column_dimensions["B"].width = 52; IN.column_dimensions["H"].width = 60

    # ---------------- Hist ----------------
    H = wb.create_sheet("Hist")
    H["A1"] = "Hist — 歷史財務（年度；持續營業口徑；¥B）。公式只引用本頁儲存格（定義名稱 h_*）"
    hdr(H, 2, ["項目", "FY2024", "FY2025", "標記", "來源"])
    HR = [("h_rev", "營收", 280.458, 307.912, "Verified", TS25),
          ("h_op", "營業利益", 12.933, 20.726, "Verified", TS25),
          ("h_da", "折舊攤銷（含油壓）", 16.469, 17.258, "Interested-party", SA + "（資料商；法定原文未取得）"),
          ("h_cx", "資本支出（含油壓）", 24.898, 13.201, "Interested-party", SA),
          ("h_rec", "應收（含其他應收）", 90.293, 76.945, "Interested-party", SA),
          ("h_inv", "存貨", 53.387, 52.838, "Interested-party", SA),
          ("h_ap", "應付帳款", 51.178, 45.161, "Interested-party", SA),
          ("h_ocf", "營業現金流", 26.650, 32.824, "Verified", TS25),
          ("h_div", "股利支付總額", 9.651, 9.531, "Verified", TS25)]
    for i, (k, lab, a, b, tag, src) in enumerate(HR, 3):
        H.cell(i, 1, lab); H.cell(i, 2, a).font = BLUE; H.cell(i, 3, b).font = BLUE; H.cell(i, 4, tag); H.cell(i, 5, src)
        dn(k + "_24", f"Hist!$B${i}"); dn(k + "_25", f"Hist!$C${i}")
    i = 3 + len(HR)
    H.cell(i, 1, "營運資金 ÷ 營收（Derived）"); H.cell(i, 2, "=(h_rec_24+h_inv_24-h_ap_24)/h_rev_24"); H.cell(i, 3, "=(h_rec_25+h_inv_25-h_ap_25)/h_rev_25")
    H.cell(i + 1, 1, "FY2023 及以前：持續營業口徑重編值未取得（油壓機器事業 2025 Q3 起列停業單位；2026-09-30 公司另公告過年度訂正）")
    H.column_dimensions["A"].width = 40

    # ---------------- Q（季度分部） ----------------
    Q = wb.create_sheet("Q")
    Q["A1"] = "Q — 季度分部營收與營業利益（¥B）。累計值為 Verified 原文；單季為相減（Derived）。分部對帳：分部合計 − 合併營收"
    cum = {  # 期間: (CMP, TRS, ACB, MFR, 合併營收, CMP_OP, TRS_OP, ACB_OP, MFR_OP, 合併 OP, 來源)
        "Q1'25": (17.422, 21.901, 28.176, 3.540, None, 0.398, 3.258, 3.373, 0.285, None, ENV + "（第 1 四半期決算短信前年同期欄）"),
        "H1'25": (36.821, 46.467, 52.585, 7.399, 143.272, 1.856, 7.205, 4.247, 0.697, 9.305, TS2Q + "（前年同期欄，重編）"),
        "9M'25": (57.519, 71.480, 78.498, 11.690, 219.187, 3.895, 10.985, 6.236, 1.173, 15.478, TS3Q + "（前年同期欄）"),
        "FY'25": ("=a25_cmp", "=a25_trs", "=a25_acb", "=a25_mfr", "=a25_rev", 5.4, 13.6, 9.1, 2.2, "=a25_op", BR1Q + "（億日圓取整）；" + TS25),
        "Q1'26": (22.259, 24.891, 32.186, 3.696, None, 1.969, 4.074, 4.107, 0.168, None, ENV),
        "H1'26": (45.521, 54.422, 58.998, 8.459, 167.401, 4.841, 9.349, 5.675, 0.490, 15.759, TS2Q),
        "9M'26": (69.780, 81.384, 86.409, 13.240, 250.813, 7.425, 14.666, 7.151, 1.165, 23.352, TS3Q)}
    hdr(Q, 2, ["累計（輸入）"] + list(cum.keys()))
    labs = ["CMP 營收", "TRS 營收", "ACB 營收", "MFR 營收", "合併營收", "CMP 營業利益", "TRS 營業利益", "ACB 營業利益", "MFR 營業利益", "合併營業利益", "來源"]
    for i, lab in enumerate(labs):
        Q.cell(3 + i, 1, lab + "（累計）")
        for j, (p, v) in enumerate(cum.items()):
            x = v[i]
            if x is None: x = "未取得（合併單季 Q1 未單獨取得；以分部合計代）" if i == 4 else "未取得"
            c = Q.cell(3 + i, 2 + j, x)
            if isinstance(x, (int, float)): c.font = BLUE
    # 單季
    q0 = 16
    QS = ["Q1'25", "Q2'25", "Q3'25", "Q4'25", "Q1'26", "Q2'26", "Q3'26"]
    hdr(Q, q0, ["單季（Derived）"] + QS)
    src_col = {"Q1'25": "B", "H1'25": "C", "9M'25": "D", "FY'25": "E", "Q1'26": "F", "H1'26": "G", "9M'26": "H"}
    def qf(q, i):
        rr = 3 + i
        m = {"Q1'25": f"=B{rr}", "Q2'25": f"=C{rr}-B{rr}", "Q3'25": f"=D{rr}-C{rr}", "Q4'25": f"=E{rr}-D{rr}",
             "Q1'26": f"=F{rr}", "Q2'26": f"=G{rr}-F{rr}", "Q3'26": f"=H{rr}-G{rr}"}
        return m[q]
    names = ["CMP 營收", "TRS 營收", "ACB 營收", "MFR 營收"]
    for i, lab in enumerate(names):
        Q.cell(q0 + 1 + i, 1, lab)
        for j, q in enumerate(QS): Q.cell(q0 + 1 + i, 2 + j, qf(q, i))
    rr = q0 + 5; Q.cell(rr, 1, "合併營收")
    for j, q in enumerate(QS):
        L = "BCDEFGH"[j]
        if q == "Q1'25": Q.cell(rr, 2 + j, f"=SUM({L}{q0+1}:{L}{q0+4})")
        elif q == "Q2'25": Q.cell(rr, 2 + j, f"=C7-SUM(B{q0+1}:B{q0+4})")
        elif q == "Q1'26": Q.cell(rr, 2 + j, f"=SUM({L}{q0+1}:{L}{q0+4})")
        elif q == "Q2'26": Q.cell(rr, 2 + j, "=G7-SUM(F17:F20)")
        else: Q.cell(rr, 2 + j, qf(q, 4))
    Q.cell(rr + 1, 1, "檢查：分部合計 − 合併營收")
    for j in range(7):
        L = "BCDEFGH"[j]; Q.cell(rr + 1, 2 + j, f"=ROUND(SUM({L}{q0+1}:{L}{q0+4})-{L}{rr},2)")
    Q.cell(rr + 2, 1, "註：Q1'25、Q1'26 合併營收以分部合計代（合併值另見簡報 710、830 億，一致）；Q4'25 用 FY2025 億日圓取整分部值，差額 −0.01 為取整（容許 0.05）")
    Q.cell(rr + 3, 1, "年度分部對帳 FY2025：分部合計 − 合併營收"); Q.cell(rr + 3, 2, "=ROUND(a25_cmp+a25_trs+a25_acb+a25_mfr-a25_rev,2)")
    # 精密減速機子列與受注
    p0 = rr + 5
    hdr(Q, p0, ["精密減速機（公司簡報；Interested-party）", "9M'25", "H1'26", "9M'26", "FY2026 指引", "說明"])
    Q.cell(p0 + 1, 1, "精密減速機營收"); Q.cell(p0 + 1, 2, "=69.2-12.2"); Q.cell(p0 + 1, 3, 45.1); Q.cell(p0 + 1, 4, 69.2); Q.cell(p0 + 1, 5, "=g26_prg")
    Q.cell(p0 + 1, 6, BR3Q + "：9M 約 692 億、年增 122 億；" + REF + "：上半 451 億")
    Q.cell(p0 + 2, 1, "精密減速機受注"); Q.cell(p0 + 2, 2, 59.0); Q.cell(p0 + 2, 4, 72.9); Q.cell(p0 + 2, 6, BR3Q + "：9M 受注 729 億（前年 590 億）；Q3 單季 QoQ +0%、YoY +18%")
    Q.cell(p0 + 3, 1, "受注 ÷ 營收（9M，Derived）"); Q.cell(p0 + 3, 2, f"=B{p0+2}/B{p0+1}"); Q.cell(p0 + 3, 4, f"=D{p0+2}/D{p0+1}")
    Q.cell(p0 + 4, 1, "CMP − 精密減速機（9M'26，Derived）"); Q.cell(p0 + 4, 4, f"=H3-D{p0+1}")
    Q.cell(p0 + 5, 1, "稼動率（3 班、月 20 日）：津 25/3Q 70%→26/3Q 85%；常州 110%→95%；浜松 逐步增產；年產能約 115 萬台（" + BR3Q + "、" + BR1Q + "）")
    Q.cell(p0 + 6, 1, "精密減速機受注 YoY：25Q1 +15%、25Q2 +15%、25Q3 +11%、25Q4 +28%、26Q1 +21%、26Q2 +23%、26Q3 +18%（" + BR2Q + "、" + BR3Q + "）")
    Q.column_dimensions["A"].width = 42

    # ---------------- PL ----------------
    P = wb.create_sheet("PL")
    P["A1"] = "PL — 三情境：命題模組（精密減速機）→ 官方分部 → 損益 → 現金流（¥B；FY2025 實績、FY2026 公司指引 ×(1＋g_anc)、FY2027 起模型）"
    PR = {}
    rr = 3
    for s in SC:
        P.cell(rr, 1, f"情境 {s}（{SCN[s]}）").font = BOLD
        hdr(P, rr + 1, ["項目", "說明"] + [f"FY{y}" for y in YRS]); b = rr + 2
        lab = ["精密減速機－機器人用（子列）", "精密減速機－非機器人用（子列）", "精密減速機－RV mini／實體 AI 新品（子列）", "精密減速機（CMP 子列）", "CMP 其他（子列）",
               "コンポーネントソリューション（CMP）營收", "トランスポートソリューション（TRS）營收", "アクセシビリティソリューション（ACB）營收", "その他（MFR）營收", "合併營收",
               "檢查：CMP 子列合計 − CMP", "檢查：分部合計 − 合併營收", "合併營收年增",
               "CMP 營業利益", "TRS 營業利益", "ACB 營業利益", "MFR 營業利益", "全社費用與消去", "營業利益", "營業利益率", "CMP 營業利益率",
               "營業外", "稅前利益", "所得稅", "稅後淨利", "少數股東損益", "歸屬母公司淨利", "平均股數（百萬股）", "EPS（日圓）",
               "折舊攤銷", "資本支出", "營運資金增加", "FCFF（NOPAT＋折舊 − 資本支出 − 營運資金增加）", "股利", "自由現金流扣股利（收支關卡用）"]
        R = {k: b + i for i, k in enumerate(["rob", "nonrob", "pai", "prg", "cmpo", "cmp", "trs", "acb", "mfr", "rev", "chk1", "chk2", "gr",
                                             "cmpop", "trsop", "acbop", "mfrop", "corp", "op", "opm", "cmpm", "nonop", "pbt", "tax", "nit", "nci", "ni", "sh", "eps",
                                             "da", "cx", "dwc", "fcff", "div", "fcfad"])}
        for i, l in enumerate(lab): P.cell(b + i, 1, l)
        for y in YRS:
            c = COL[y]; p = COL.get(y - 1); cell = lambda k: f"{c}{R[k]}"; prev = lambda k: f"{p}{R[k]}"
            def w(k, f): P[cell(k)] = f
            if y == 2025:
                for k, nm in (("cmp", "a25_cmp"), ("trs", "a25_trs"), ("acb", "a25_acb"), ("mfr", "a25_mfr"), ("rev", "a25_rev"), ("op", "a25_op")):
                    w(k, f"={nm}")
                w("chk2", f"=ROUND(SUM({c}{R['cmp']}:{c}{R['mfr']})-{c}{R['rev']},2)")
                w("ni", "=a25_ni"); w("da", "=h_da_25"); w("cx", "=h_cx_25")
                for k, v in (("cmpop", 5.4), ("trsop", 13.6), ("acbop", 9.1), ("mfrop", 2.2)): w(k, v)
                w("corp", f"={c}{R['op']}-SUM({c}{R['cmpop']}:{c}{R['mfrop']})")
                w("opm", f"={c}{R['op']}/{c}{R['rev']}"); w("cmpm", f"={c}{R['cmpop']}/{c}{R['cmp']}")
                continue
            if y == 2026:
                w("prg", "=g26_prg*(1+g_anc)"); w("pai", "=pai_0")
                w("rob", f"=r_rob*({cell('prg')}-{cell('pai')})"); w("nonrob", f"=(1-r_rob)*({cell('prg')}-{cell('pai')})")
                w("cmpo", "=g26_cmpo*(1+g_anc)")
                w("trs", "=g26_trs*(1+g_anc)"); w("acb", "=g26_acb*(1+g_anc)"); w("mfr", "=g26_mfr*(1+g_anc)")
                w("cmpop", "=g26_cmpop*(1+g_anc)"); w("trsop", f"=m_trs*{cell('trs')}"); w("acbop", f"=m_acb*{cell('acb')}"); w("mfrop", f"=m_mfr*{cell('mfr')}")
                w("corp", "=corp_26*(1+g_anc)"); w("nonop", "=nonop_26"); w("sh", "=sh_26")
            else:
                rg = f"rg27_{s}" if y == 2027 else (f"rgm_{s}" if y <= 2029 else f"rglt_{s}")
                w("rob", f"={prev('rob')}*(1+{rg})*(1+shr_{s})"); w("nonrob", f"={prev('nonrob')}*(1+nr_{s})")
                w("pai", f"=pai_0+(pai_{s}-pai_0)*({y}-2026)/6")
                w("prg", f"={cell('rob')}+{cell('nonrob')}+{cell('pai')}"); w("cmpo", f"={prev('cmpo')}")
                w("trs", f"={prev('trs')}*(1+gtrs_{s})"); w("acb", f"={prev('acb')}*(1+gacb_{s})"); w("mfr", f"={prev('mfr')}*(1+gmfr_{s})")
                adj = f"+madj_{s}*{cell('cmp')}" if y >= 2028 else ""
                w("cmpop", f"=$D{R['cmpop']}+mm_cmp*({cell('cmp')}-$D{R['cmp']}){adj}")
                w("trsop", f"=m_trs*{cell('trs')}"); w("acbop", f"=m_acb*{cell('acb')}"); w("mfrop", f"=m_mfr*{cell('mfr')}")
                w("corp", f"=$D{R['corp']}*{cell('rev')}/$D{R['rev']}"); w("nonop", "=nonop"); w("sh", "=sh_v")
            w("cmp", f"={cell('prg')}+{cell('cmpo')}")
            w("rev", f"={cell('cmp')}+{cell('trs')}+{cell('acb')}+{cell('mfr')}")
            w("chk1", f"=ROUND({cell('rob')}+{cell('nonrob')}+{cell('pai')}+{cell('cmpo')}-{cell('cmp')},6)")
            w("chk2", f"=ROUND(SUM({c}{R['cmp']}:{c}{R['mfr']})-{c}{R['rev']},6)")
            w("gr", f"={cell('rev')}/{prev('rev')}-1")
            w("op", f"=SUM({c}{R['cmpop']}:{c}{R['corp']})"); w("opm", f"={cell('op')}/{cell('rev')}"); w("cmpm", f"={cell('cmpop')}/{cell('cmp')}")
            w("pbt", f"={cell('op')}+{cell('nonop')}"); w("tax", f"={cell('pbt')}*tax_w"); w("nit", f"={cell('pbt')}-{cell('tax')}")
            w("nci", f"={cell('nit')}*nci_r"); w("ni", f"={cell('nit')}-{cell('nci')}"); w("eps", f"={cell('ni')}/{cell('sh')}*1000")
            w("da", f"=fc_da*{cell('rev')}"); w("cx", "=cx_26" if y == 2026 else f"=fc_cx*{cell('rev')}"); w("dwc", f"=fc_wc*({cell('rev')}-{prev('rev')})")
            w("fcff", f"={cell('op')}*(1-tax_w)+{cell('da')}-{cell('cx')}-{cell('dwc')}")
            w("div", "=dps_26*sh_26/1000" if y == 2026 else f"=pay_r*{cell('ni')}")
            w("fcfad", f"={cell('fcff')}-{cell('div')}")
        P.cell(b, 2, "命題模組：前一年 ×(1＋機器人需求年增 rg)×(1＋份額變化 shr)")
        P.cell(b + 3, 2, "FY2026＝公司指引 946 億 ×(1＋g_anc)")
        P.cell(R["cmpop"], 2, "FY2027 起＝FY2026 ＋ mm_cmp × ΔCMP 營收 ＋ madj（FY2028 起）")
        P.cell(R["fcff"], 2, "NOPAT＝營業利益 ×(1 − tax_w)；股票報酬不加回")
        for k in ("gr", "opm", "cmpm"):
            for y in YRS: P[f"{COL[y]}{R[k]}"].number_format = "0.0%"
        PR[s] = R; rr = b + len(lab) + 2
    P.column_dimensions["A"].width = 44; P.column_dimensions["B"].width = 30
    for s in SC:
        for k in ("rev", "op", "ni", "eps", "fcff", "prg", "cmp", "fcfad"):
            dn(f"pl_{k}_{s}", f"PL!$C${PR[s][k]}:$J${PR[s][k]}")

    # ---------------- DCF ----------------
    D = wb.create_sheet("DCF")
    D["A1"] = "DCF — 標準版 FCFF（NOPAT＋折舊 − 資本支出 − 營運資金增加），WACC 折現，評價日 2026-12-31，明確期 FY2027–FY2032，期中慣例；終值＝NOPAT(FY2033) ×(1 − g ÷ RONIC) ÷ (WACC − g)"
    D["A3"] = "ke（股權成本）"; D["B3"] = "=rf+beta*(erp+crp)"; dn("ke", "DCF!$B$3")
    D["A4"] = "股權市值 E（¥B）"; D["B4"] = "=px*sh_v/1000"
    D["A5"] = "負債 D（借款＋租賃）"; D["B5"] = "=debt_v+lease_v"
    D["A6"] = "稅後 kd"; D["B6"] = "=kd*(1-tax_w)"
    D["A7"] = "WACC"; D["B7"] = "=(B4*ke+B5*B6)/(B4+B5)+wacc_adj"; dn("wacc", "DCF!$B$7")
    D["A8"] = "永續成長率 g＝MIN(g_cap, rf)"; D["B8"] = "=MIN(g_cap,rf)"; dn("g_tv", "DCF!$B$8")
    D["A9"] = "RONIC"; D["B9"] = "=wacc+ronic_sp"; dn("ronic", "DCF!$B$9")
    D["A10"] = "淨現金與投資（資產負債表 2026-09-30）＝現金 − 借款 − 租賃 − 非支配 ＋ 持分法投資 × inv_k"; D["B10"] = "=cash_v-debt_v-lease_v-mi_v+inv_k*inv_v"; dn("dcf_nc", "DCF!$B$10")
    D["A11"] = "檢查：WACC 由 In 頁輸入重算 − WACC"; D["B11"] = "=ROUND(((px*sh_v/1000)*(rf+beta*(erp+crp))+(debt_v+lease_v)*kd*(1-tax_w))/((px*sh_v/1000)+debt_v+lease_v)+wacc_adj-wacc,9)"
    for rr_ in range(3, 11): D[f"B{rr_}"].number_format = "0.00%" if rr_ in (3, 6, 7, 8, 9) else "#,##0.0"
    DR = {}; rr = 13
    for s in SC:
        R = PR[s]
        D.cell(rr, 1, f"情境 {s}（{SCN[s]}）").font = BOLD
        hdr(D, rr + 1, ["項目"] + [f"FY{y}" for y in range(2027, 2033)]); b = rr + 2
        rows = ["EBIT（營業利益）", "NOPAT", "＋折舊攤銷", "− 資本支出", "− 營運資金增加", "FCFF", "檢查：各列加總 − FCFF", "折現年數（期中）", "折現因子", "FCFF 現值"]
        for i, l in enumerate(rows): D.cell(b + i, 1, l)
        for j, y in enumerate(range(2027, 2033)):
            c = "BCDEFG"[j]; pc = COL[y]
            D[f"{c}{b}"] = f"=PL!{pc}{R['op']}"; D[f"{c}{b+1}"] = f"={c}{b}*(1-tax_w)"
            D[f"{c}{b+2}"] = f"=PL!{pc}{R['da']}"; D[f"{c}{b+3}"] = f"=-PL!{pc}{R['cx']}"; D[f"{c}{b+4}"] = f"=-PL!{pc}{R['dwc']}"
            D[f"{c}{b+5}"] = f"=PL!{pc}{R['fcff']}"; D[f"{c}{b+6}"] = f"=ROUND(SUM({c}{b+1}:{c}{b+4})-{c}{b+5},9)"
            D[f"{c}{b+7}"] = j + 0.5; D[f"{c}{b+8}"] = f"=1/(1+wacc)^{c}{b+7}"; D[f"{c}{b+9}"] = f"={c}{b+5}*{c}{b+8}"
        t = b + 11
        items = [("明確期現值合計", f"=SUM(B{b+9}:G{b+9})"),
                 ("NOPAT FY2033＝NOPAT FY2032 ×(1＋g)", f"=G{b+1}*(1+g_tv)"),
                 ("終值（FY2032 年底）", f"=B{t+1}*(1-g_tv/ronic)/(wacc-g_tv)"),
                 ("終值現值", f"=B{t+2}/(1+wacc)^6"),
                 ("企業價值 EV", f"=B{t}+B{t+3}"),
                 ("＋淨現金與投資（2026-09-30）", "=dcf_nc"),
                 ("＋Stub 期 FCFF（FY2026 FCFF × stub_frac）", f"=PL!D{R['fcff']}*stub_frac"),
                 ("− 尚待買回金額", "=-bb_rem"),
                 ("股權價值", f"=B{t+4}+B{t+5}+B{t+6}+B{t+7}"),
                 ("每股價值（評價日）", f"=MAX(0,B{t+8})/sh_v*1000"),
                 ("DCF 情境價值推到 12 個月後（×(1＋ke)^t − 期間股利）", f"=B{t+9}*(1+ke)^t_tgt-dps_tgt"),
                 ("終值占 EV", f"=IF(B{t+4}>0,B{t+3}/B{t+4},\"N/M\")"),
                 ("檢查：終值占 EV 是否超過 80%", f"=IF(AND(B{t+4}>0,B{t+11}<0.8),\"通過\",\"標示：超過 80% 或 EV≤0\")"),
                 ("隱含終值 EV/NOPAT 倍數", f"=B{t+2}/G{b+1}"),
                 ("終值替代法：FCFF(FY2032)×(1＋g)÷(WACC − g) 的每股價值（推到 12 個月後）", f"=((B{t}+G{b+5}*(1+g_tv)/(wacc-g_tv)/(1+wacc)^6)+B{t+5}+B{t+6}+B{t+7})/sh_v*1000*(1+ke)^t_tgt-dps_tgt")]
        for i, (l, f) in enumerate(items):
            D.cell(t + i, 1, l); D.cell(t + i, 2, f)
        DR[s] = dict(v12=t + 10, ps=t + 9, ev=t + 4, tvs=t + 11, alt=t + 14, chk=b + 6)
        dn(f"dcf_{s}", f"DCF!$B${t+10}")
        rr = t + len(items) + 2
    D.column_dimensions["A"].width = 60

    # ---------------- Val ----------------
    V = wb.create_sheet("Val")
    V["A1"] = "Val — 目標價：DCF（權重 w_dcf）＋ 分部本益比加總（Comparable，1 − w_dcf）；三情境 25／50／25；評等五級（> +50% 強力買進、> +25% 買進、≥ −25% 中立、≥ −50% 放空、其餘強力放空）"
    hdr(V, 3, ["項目", "情境 A", "情境 B", "情境 C", "加權", "說明"])
    vr = 4
    def vrow(lab, fs, wt=None, note=""):
        nonlocal vr
        V.cell(vr, 1, lab)
        for j, f in enumerate(fs): V.cell(vr, 2 + j, f)
        if wt: V.cell(vr, 5, wt)
        V.cell(vr, 6, note); vr += 1; return vr - 1
    r_dcf = vrow("DCF 情境價值推到 12 個月後", [f"=dcf_{s}" for s in SC], "=w_A*B{r}+w_B*C{r}+w_C*D{r}", "DCF 頁")
    V.cell(r_dcf, 5, f"=w_A*B{r_dcf}+w_B*C{r_dcf}+w_C*D{r_dcf}")
    cf = lambda s, k: f"PL!F{PR[s][k]}"
    comp_f = [f"=({cf(s,'cmpop')}*pe_cmp+{cf(s,'trsop')}*pe_trs+{cf(s,'acbop')}*pe_acb+({cf(s,'mfrop')}+{cf(s,'corp')}+{cf(s,'nonop')})*pe_all)*(1-tax_w)*(1-nci_r)/sh_v*1000" for s in SC]
    r_cu = vrow("Comparable（未調整）：分部本益比加總（FY2028 分部營業利益 × 稅後 × 群組前瞻 P/E）", comp_f, None,
                "CMP×pe_cmp、TRS×pe_trs、ACB×pe_acb、MFR＋全社費用＋營業外×pe_all；×(1 − tax_w)(1 − nci_r)；目標時點的 NTM 以 FY2028 代（2027-11 起算 NTM 約 1/12 FY2027＋11/12 FY2028）")
    V.cell(r_cu, 5, f"=w_A*B{r_cu}+w_B*C{r_cu}+w_C*D{r_cu}")
    r_ca = vrow("Comparable（相對倍數調整後，進加權）", [f"={L}{r_cu}*rel_k" for L in "BCD"], None, "rel_k＝1.00（未調整：缺歷史相對倍數）")
    V.cell(r_ca, 5, f"=w_A*B{r_ca}+w_B*C{r_ca}+w_C*D{r_ca}")
    r_tps = vrow("各情境目標價（w_dcf × DCF ＋ (1 − w_dcf) × Comparable）", [f"=w_dcf*{L}{r_dcf}+(1-w_dcf)*{L}{r_ca}" for L in "BCD"])
    V.cell(r_tps, 5, f"=w_A*B{r_tps}+w_B*C{r_tps}+w_C*D{r_tps}")
    vr += 1
    def one(lab, f, note=""):
        nonlocal vr
        V.cell(vr, 1, lab); V.cell(vr, 2, f); V.cell(vr, 6, note); vr += 1; return vr - 1
    r_tp = one("加權目標價（日圓）", f"=E{r_tps}"); dn("tp", f"Val!$B${r_tp}")
    r_dcft = one("DCF 目標價（情境加權）", f"=E{r_dcf}")
    r_cmpt = one("Comparable 目標價（情境加權，進加權）", f"=E{r_ca}")
    r_px = one("現價（In 頁 px）", "=px")
    r_up = one("上檔空間", f"=B{r_tp}/px-1"); dn("upside", f"Val!$B${r_up}")
    r_rt = one("估值評等（只看上檔空間）", f'=IF(B{r_up}>0.5,"強力買進",IF(B{r_up}>0.25,"買進",IF(B{r_up}>=-0.25,"中立",IF(B{r_up}>=-0.5,"放空","強力放空"))))')
    r_gate = vr; V.cell(vr, 1, "收支關卡：FY2026–FY2030 自由現金流扣股利合計（¥B；情境 A／B／C）")
    for j, s in enumerate(SC): V.cell(vr, 2 + j, f"=SUM(PL!D{PR[s]['fcfad']}:PL!H{PR[s]['fcfad']})")
    vr += 1
    r_vg = one("入不敷出判定用：情境 B 合計", f"=C{r_gate}")
    r_gj = one("入不敷出判定", f'=IF(B{r_vg}<0,"不通過","通過")')
    r_rec = one("投資建議", f'=IF(AND(B{r_vg}<0,OR(B{r_rt}="買進",B{r_rt}="強力買進")),"收支不足",B{r_rt})')
    r_dist = one("距最近級距邊界（百分點）", f"=MIN(ABS(B{r_up}-0.5),ABS(B{r_up}-0.25),ABS(B{r_up}+0.25),ABS(B{r_up}+0.5))*100")
    r_bnd = one("翻轉目標價：升一級／降一級", f'=TEXT(px*(1+IF(B{r_up}>0.5,9,IF(B{r_up}>0.25,0.5,IF(B{r_up}>=-0.25,0.25,IF(B{r_up}>=-0.5,-0.25,-0.5))))),"#,##0")&"／"&TEXT(px*(1+IF(B{r_up}>0.5,0.5,IF(B{r_up}>0.25,0.25,IF(B{r_up}>=-0.25,-0.25,IF(B{r_up}>=-0.5,-0.5,-9))))),"#,##0")', "升一級：上檔超過該級上緣；降一級：低於該級下緣")
    vr += 1
    r_a1 = one("上檔三段歸因：DCF 貢獻", f"=w_dcf*(B{r_dcft}/px-1)")
    r_a2 = one("上檔三段歸因：未調整 Comparable 貢獻", f"=(1-w_dcf)*(E{r_cu}/px-1)")
    r_a3 = one("上檔三段歸因：相對倍數調整貢獻", f"=(1-w_dcf)*(E{r_ca}-E{r_cu})/px")
    one("檢查：三段合計 − 上檔空間", f"=ROUND(B{r_a1}+B{r_a2}+B{r_a3}-B{r_up},9)")
    r_pd1 = one("期間：模型折現（評價日 2026-12-31 → 2027-12-31，年）", "=t_tgt")
    r_pd2 = one("期間：持有（股價日 → 目標價時點，年）", "=(tp_date-px_date)/365", "上檔空間不是年化報酬")
    one("WACC／ke／g／RONIC", '=TEXT(wacc,"0.00%")&"／"&TEXT(ke,"0.00%")&"／"&TEXT(g_tv,"0.00%")&"／"&TEXT(ronic,"0.00%")')
    r_pe = one("隱含前瞻 P/E：加權目標價 ÷ 情境 B FY2028 EPS", f"=B{r_tp}/PL!F{PR['B']['eps']}", "分母：模型情境 B FY2028 EPS")
    r_pe2 = one("現行前瞻 P/E（資料商 NTM，2026-11-13）÷ CMP 群組中位數", "=pe_self/pe_cmp", "只作對照")
    for rr_ in (r_up, r_a1, r_a2, r_a3): V[f"B{rr_}"].number_format = "0.0%"
    for rr_ in (r_dcf, r_cu, r_ca, r_tps):
        for L in "BCDE": V[f"{L}{rr_}"].number_format = "#,##0"
    for rr_ in (r_tp, r_dcft, r_cmpt): V[f"B{rr_}"].number_format = "#,##0"
    V.column_dimensions["A"].width = 70; V.column_dimensions["F"].width = 60
    VR = dict(tp=r_tp, dcf=r_dcft, cmp=r_cmpt, up=r_up, rt=r_rt, rec=r_rec, gate=r_gate, dist=r_dist, bnd=r_bnd, a1=r_a1, a2=r_a2, a3=r_a3,
              tps=r_tps, dcfr=r_dcf, cu=r_cu, pd2=r_pd2, pe=r_pe, gj=r_gj)

    # ---------------- Recon ----------------
    RC = wb.create_sheet("Recon")
    RC["A1"] = "Recon — 與公司指引、市場共識對照；反解（整本重算）"
    hdr(RC, 2, ["項目", "模型（情境 B）", "公司或共識", "差", "說明"])
    rc = [("FY2026 營收", f"=PL!D{PR['B']['rev']}", "=g26_rev", "公司指引（Interested-party）"),
          ("FY2026 營業利益", f"=PL!D{PR['B']['op']}", "=g26_op", "公司指引"),
          ("FY2026 歸屬母公司淨利", f"=PL!D{PR['B']['ni']}", "=g26_ni", "公司指引含停業單位利益 1.359（9M）；模型以營業外 3.7 代"),
          ("FY2026 精密減速機營收", f"=PL!D{PR['B']['prg']}", "=g26_prg", "公司參考資料"),
          ("FY2026 營收 vs 共識", f"=PL!D{PR['B']['rev']}", "=cons_rev26", "資料商共識（2026-09-08，財報前，只作歷史對照）"),
          ("FY2027 EPS（模型 B） vs 現價 ÷ 現行前瞻 P/E 隱含 NTM EPS", f"=PL!E{PR['B']['eps']}", "=px/pe_self", "口徑：資料商 NTM（2026-11-13 起 12 個月，約 1.5/12 FY2026＋10.5/12 FY2027）vs 模型 FY2027；口徑不同，只看方向"),
          ("加權目標價 vs 賣方平均目標價", "=tp", "=cons_tp", "賣方 10 位平均（Interested-party）")]
    for i, (l, a, b_, n) in enumerate(rc, 3):
        RC.cell(i, 1, l); RC.cell(i, 2, a); RC.cell(i, 3, b_); RC.cell(i, 4, f"=B{i}-C{i}"); RC.cell(i, 5, n)
    i = 3 + len(rc) + 1
    RC.cell(i, 1, "反解（整本重算，build_nabtesco.py 二分法；常數為建置時結果）").font = BOLD; i += 1
    hdr(RC, i, ["反解目標", "鍵", "反解值", "基準值", "說明"]); i += 1
    for item in (rev or []):
        RC.cell(i, 1, item["label"]); RC.cell(i, 2, item["key"]); RC.cell(i, 3, item["val"]); RC.cell(i, 4, item["base"]); RC.cell(i, 5, item["note"]); i += 1
    RC.column_dimensions["A"].width = 60; RC.column_dimensions["E"].width = 70

    # ---------------- Sens ----------------
    SE = wb.create_sheet("Sens")
    SE["A1"] = "Sens — 一次動一個 In 頁參數（敏感度低／高），整本重算（build_nabtesco.py 以 LibreOffice 重算每一案例）；結果為常數快照"
    base_tp = sens["base"]["tp"] if sens else None
    SE["A2"] = "快照基準值（加權，日圓）"; SE["B2"] = base_tp
    SE["C2"] = "本檔即時值（加權，日圓）"; SE["D2"] = "=tp"
    SE["E2"] = '=IF(ABS(D2-B2)>0.5,"STALE：In 頁已改，Sens 須重跑","OK：快照與本檔一致")'
    SE["F2"] = f"快照輸入指紋 {sens['fp'] if sens else ''}"
    hdr(SE, 4, ["驅動", "設定", "鍵", "加權目標價", "相對基準", "DCF 目標價", "DCF 相對基準", "估值評等"])
    sr = 5
    if sens:
        for cs in sens["cases"]:
            SE.cell(sr, 1, cs["drv"]); SE.cell(sr, 2, cs["set"]); SE.cell(sr, 3, cs["key"])
            SE.cell(sr, 4, cs["tp"]); SE.cell(sr, 5, cs["tp"] / base_tp - 1)
            SE.cell(sr, 6, cs["dcf"]); SE.cell(sr, 7, cs["dcf"] / sens["base"]["dcf"] - 1); SE.cell(sr, 8, cs["rt"])
            SE[f"E{sr}"].number_format = "0.0%"; SE[f"G{sr}"].number_format = "0.0%"; SE[f"D{sr}"].number_format = "#,##0"; SE[f"F{sr}"].number_format = "#,##0"
            sr += 1
    SE.cell(sr + 1, 1, "說明：相對倍數係數 rel_k 基準即為 1.00（未調整），故「相對倍數係數＝1」案例＝基準。nci_r、pay_r 等無區間的鍵不列案例。")
    SE.column_dimensions["A"].width = 46; SE.column_dimensions["B"].width = 34

    # ---------------- 對帳（線層級變數） ----------------
    RE = wb.create_sheet("對帳")
    RE["A1"] = "對帳 — 本模型使用的「線層級」變數（線框架 v0.1 取值；與 physical_ai/lines/ 完成版對帳時處理，不靜默改口徑）"
    hdr(RE, 2, ["變數", "所屬線", "本模型使用值（或區間）", "單位", "口徑", "出處", "標記", "影響本模型哪些儲存格"])
    rec = [("全球工業機器人裝機 2025", "2a", "60.3 萬", "台", "IFR 裝機（年度）", "IFR President's Report（" + L2A + " 第 3 節）", "Verified", "說明用；rg27／rgm 依其預測增速"),
           ("全球工業機器人裝機預測 2026、2029", "2a", "65.5 萬、80.6 萬（年複合 7.2%）", "台", "IFR 預測", L2A + " 第 1 節", "Interested-party（IFR 產業協會）", "In rg27_B、rgm_B → PL 各情境精密減速機機器人用列"),
           ("工業機器人需求年增 FY2027（情境 A／B／C）", "2a", "15%／7.2%／−5%", "%", "以台數增速代 RV 機器人用營收增速", "本模型 Assumed／Analogy", "Assumed", "In rg27_A／B／C"),
           ("工業機器人需求年增 FY2030–32", "2a", "7%／5%／3%", "%/年", "同上", "本模型", "Assumed", "In rglt_*"),
           ("JARA 受注（2026 年 7–9 月）", "2a", "3,254 億日圓，+46.6%；訂單÷生產 1.24", "¥", "日系整機訂單", L2A + " 第 5 節（JARA）", "Verified", "情境 A 依據"),
           ("中國工業機器人裝機占比、國產比 2025", "2a", "59%；國產 55%", "%", "IFR", L2A, "Verified", "shr_* 說明"),
           ("傳統 6 軸 BOM 減速機占比", "2a", "25–35%", "%", "本體成本", L2A + " 第 3 節", "Analogy", "說明用，未進公式"),
           ("RV 用於傳統 6 軸重負載關節、協作與人形以諧波為主", "2a／2b", "物理層註記", "—", "零組件結構", L2A + " 第 3 節；" + L2B + " 第 3.2 節", "Analogy", "pai_*（人形對 RV 的需求假設小）"),
           ("每台機器人 RV 用量", "2a", "未取得（傳統 6 軸約 3–4 顆 RV：底座、大臂、小臂；腕部諧波）", "顆／台", "物理", "未取得公開統計", "Assumed", "未進公式（以營收增速代）"),
           ("精密減速機產能", "2a", "約 115 萬台／年", "台", "公司產能", BR1Q, "Interested-party", "fc_cx 說明（量增超過約一成需擴產）"),
           ("精密減速機隱含平均單價", "2a", "約 8–10 萬日圓（946 億 ÷（115 萬 × 稼動約 85–95%））", "¥／台", "Derived，產能口徑與台數口徑可能不同", "Derived", "Derived", "說明用，未進公式"),
           ("精密減速機機器人用比例", "2a", "60–80%（基準 70%）", "%", "營收", "未取得", "Assumed", "In r_rob"),
           ("RV 全球份額（中大型機器人關節）", "2a", "約 6 成", "%", "公司自稱", "公司產品頁（經 jaredwatkins.com 轉述）", "Interested-party", "shr_* 說明"),
           ("人形機器人全球出貨 2026", "2b", ">5 萬台（TrendForce）；H1 追蹤 1.9–2.5 萬", "台", "出貨", L2B + " 第 7 節", "Interested-party", "pai_A（RV mini 進人形為選擇權）"),
           ("人形每台減速機", "2b", "諧波 14–16 顆，或行星替代；RV 未列入主流 BOM", "顆／台", "BOM", L2B + " 第 3.2 節", "Interested-party", "pai_*")]
    for i, row in enumerate(rec, 3):
        for j, v in enumerate(row): RE.cell(i, 1 + j, v)
    RE.column_dimensions["A"].width = 40; RE.column_dimensions["C"].width = 40; RE.column_dimensions["F"].width = 50

    # ---------------- Risk ----------------
    RK = wb.create_sheet("Risk")
    RK["A1"] = "Risk — 風險登錄（未建模的尾部風險：風險、是否建模、理由、觀察指標）"
    hdr(RK, 2, ["風險", "是否建模", "理由", "觀察指標"])
    risks = [("過年度財報訂正（持分法 TMT マシナリー）擴大或內控缺失", "否", "訂正內容已反映在 2026-09-30 後報表；金額未取得", "訂正報告書、監查意見、東證處分"),
             ("中國 RV 廠價格戰與份額流失", "是（shr_*、madj_C）", "情境 C 與 Sens", "精密減速機受注 YoY、常州稼動率、陸廠 RV 營收"),
             ("工業機器人循環轉折（2027 年回落）", "是（rg27_C）", "情境 C", "JARA 受注 YoY、安川／Fanuc Robot 受注"),
             ("匯率（人民幣、美元、歐元）", "否", "公司揭露敏感度：人民幣 1 円 → 營收 15.7 億、營業利益 2.5 億（" + BR2Q + "）；未逐年建模", "USDJPY、CNYJPY"),
             ("航空（777X）時程延誤", "否", "隱含在 gtrs_* 區間", "Boeing 777X 認證、交付"),
             ("產能瓶頸（量增超過約一成需擴產）", "部分（fc_cx）", "資本支出率 Sens", "津、常州稼動率；資本支出計畫"),
             ("人形機器人採用 RV mini 的選擇權", "部分（pai_A）", "無公開訂單", "公司法說、人形整機 BOM 拆解")]
    for i, row in enumerate(risks, 3):
        for j, v in enumerate(row): RK.cell(i, 1 + j, v)
    RK.column_dimensions["A"].width = 50; RK.column_dimensions["C"].width = 60

    # ---------------- Sources ----------------
    SO = wb.create_sheet("Sources")
    for i, s_ in enumerate([TS3Q, TS2Q, TS2QX, TS25, BR3Q, BR2Q, QA2Q, BR1Q, REF, ENV, EDN, TRD, SA + " statistics／forecast／財報頁", JGB10, L2A, L2B,
                            "IFR World Robotics 2026（經線 2a 框架）", "JARA 2026 年 7–9 月（經線 2a 框架）"], 1):
        SO.cell(i, 1, s_)

    # ---------------- README（工作表；文字由公式組字） ----------------
    RD["A1"] = f"Nabtesco（6268.T）個股模型 {VER}｜build_id {{BID}}"
    RD["A2"] = '="資料截止日 2026-11-16；成品日 2026-11-16；股價基準日 "&TEXT(px_date,"yyyy-mm-dd")&"（收盤 "&TEXT(px,"#,##0")&" 日圓）；目標價時點 "&TEXT(tp_date,"yyyy-mm-dd")'
    RD["A3"] = "命題：Nabtesco 的股價由精密減速機（RV，CMP 分部約 99%）的工業機器人循環決定——RV 營收＝機器人用（全球裝機 × 份額）＋非機器人用＋RV mini 新品；TRS、ACB、MFR 三個官方分部以成長路徑建模"
    RD["A4"] = (f'="在 FY2027 全球工業機器人需求情境 B 年增 "&TEXT(rg27_B,"0.0%")&"、RV 份額每年 "&TEXT(shr_B,"0.0%")&"、CMP 邊際利益率 "&TEXT(mm_cmp,"0%")&" 的基準假設下：加權目標價 "&TEXT(Val!B{r_tp},"#,##0")&" 日圓（DCF "&TEXT(Val!B{r_dcft},"#,##0")&"、分部本益比加總 "&TEXT(Val!B{r_cmpt},"#,##0")&"），上檔空間 "&TEXT(Val!B{r_up},"+0.0%;-0.0%")&"，估值評等 "&Val!B{r_rt}&"；收支關卡 "&Val!B{r_gj}&"（情境 B FY2026–30 自由現金流扣股利 "&TEXT(Val!C{r_gate},"#,##0.0")&"）；投資建議 "&Val!B{r_rec}')
    RD["A5"] = f'="距最近級距邊界 "&TEXT(Val!B{r_dist},"0.0")&" 個百分點；翻轉目標價（升一級／降一級）"&Val!B{r_bnd}&" 日圓。上檔三段歸因：DCF "&TEXT(Val!B{r_a1},"+0.0%;-0.0%")&"、未調整 Comparable "&TEXT(Val!B{r_a2},"+0.0%;-0.0%")&"、相對倍數調整 "&TEXT(Val!B{r_a3},"+0.0%;-0.0%")'
    RD["A6"] = "失效條件（可觀察）：精密減速機受注 YoY 連兩季轉負；常州稼動率跌破 80%；JARA 受注 YoY 轉負；CMP 單季營業利益率跌破 8%"
    RD["A7"] = "範圍：不讀 Siliconomics／Tokenomics／Wattonomics（算力線）；自用 AI 讀法、Wattonomics 上線時程等規格條文不適用。HTML 一頁摘要本版未做"
    wb.save(path)
    return VR, PR, DR


# ---------------- 重算與讀值 ----------------
LO = "file:///tmp/lo_nab_profile"
def recalc(src, outdir):
    os.makedirs(outdir, exist_ok=True)
    subprocess.run(["soffice", f"-env:UserInstallation={LO}", "--headless", "--convert-to", "xlsx", "--outdir", outdir, src],
                   check=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    return os.path.join(outdir, os.path.basename(src))

def read_vals(path, VR, DR=None):
    wb = openpyxl.load_workbook(path, data_only=True); V = wb["Val"]
    d = dict(tp=V[f"B{VR['tp']}"].value, dcf=V[f"B{VR['dcf']}"].value, cmp=V[f"B{VR['cmp']}"].value, up=V[f"B{VR['up']}"].value, rt=V[f"B{VR['rt']}"].value)
    if DR:
        I = wb["In"]; w = {}; wd = None
        for row in I.iter_rows(values_only=True):
            if row[0] in ("w_A", "w_B", "w_C"): w[row[0][-1]] = row[4]
            if row[0] == "w_dcf": wd = row[4]
        alt = sum(w[s] * wb["DCF"][f"B{DR[s]['alt']}"].value for s in SC)
        d["alt_dcf"] = alt; d["alt_tp"] = wd * alt + (1 - wd) * d["cmp"]
    return d

def run_case(ov, tag):
    d = tempfile.mkdtemp(prefix="nab_"); p = os.path.join(d, f"c_{tag}.xlsx")
    VR, _, DR = build(p, ov)
    out = recalc(p, os.path.join(d, "o"))
    v = read_vals(out, VR, DR if tag == "base" else None); shutil.rmtree(d, ignore_errors=True); return v

def in_val(k, col):
    for row in ROWS:
        if row[0] == k:
            return row[{"lo": 3, "base": 4, "hi": 5}[col]]
    raise KeyError(k)

def fingerprint():
    return hashlib.sha1(json.dumps([r for r in ROWS], ensure_ascii=False, default=str).encode()).hexdigest()[:10]

CASES = [  # (鍵, 驅動)
    ("g_anc", "FY2026 實績相對指引"), ("r_rob", "精密減速機機器人用比例"), ("rg27_B", "機器人需求年增 FY2027（情境 B）"), ("rgm_B", "機器人需求年增 FY2028–29（情境 B）"),
    ("rglt_B", "機器人需求年增 FY2030–32（情境 B）"), ("shr_B", "RV 份額＋價格年變化（情境 B）"), ("nr_B", "非機器人用年增（情境 B）"), ("pai_B", "RV mini／實體 AI 新品 FY2032（情境 B）"),
    ("mm_cmp", "CMP 邊際營業利益率"), ("madj_B", "CMP 利益率調整（情境 B）"), ("gtrs_B", "TRS 年增（情境 B）"), ("gacb_B", "ACB 年增（情境 B）"), ("gmfr_B", "MFR 年增（情境 B）"),
    ("m_trs", "TRS 營業利益率"), ("m_acb", "ACB 營業利益率"), ("m_mfr", "MFR 營業利益率"), ("nonop", "營業外"), ("tax_w", "有效稅率"),
    ("beta", "β"), ("erp", "ERP"), ("g_cap", "永續成長率上限"), ("ronic_sp", "RONIC − WACC"), ("wacc_adj", "WACC ±1pt"),
    ("fc_da", "折舊 ÷ 營收"), ("fc_cx", "資本支出 ÷ 營收"), ("fc_wc", "營運資金 ÷ 營收增量"), ("inv_k", "持分法投資計入比例"), ("bb_rem", "尚待買回金額"), ("w_dcf", "DCF 權重")]

def num(v, r=None):
    if isinstance(v, str) and v.startswith("="):
        return None
    return v

def write_readme(path, VR, PR, DR, bid, sens, revs):
    wb = openpyxl.load_workbook(path, data_only=True); V = wb["Val"]; P = wb["PL"]; D = wb["DCF"]; RD = wb["README"]
    g = lambda c: V[c].value
    f0 = lambda x: f"{x:,.0f}"; pc = lambda x: f"{x*100:+.1f}%"
    tp, dcf, cmp_, up = g(f"B{VR['tp']}"), g(f"B{VR['dcf']}"), g(f"B{VR['cmp']}"), g(f"B{VR['up']}")
    rows_s = "\n".join(f"| {c['drv']} | {c['set']} | `{c['key']}` | {f0(c['tp'])} | {pc(c['tp']/sens['base']['tp']-1)} | {c['rt']} |" for c in sens["cases"])
    big = sorted(sens["cases"], key=lambda c: -abs(c["tp"] / sens["base"]["tp"] - 1))[:6]
    rv = "\n".join(f"| {r['label']} | `{r['key']}` | {r['val'] if isinstance(r['val'], str) else f'{r[chr(118)+chr(97)+chr(108)]:.4f}'} | {r['base']} | {r['note']} |" for r in revs)
    yr = lambda s, k: " | ".join(f"{P[f'{COL[y]}{PR[s][k]}'].value:,.1f}" if isinstance(P[f'{COL[y]}{PR[s][k]}'].value, (int, float)) else "—" for y in YRS)
    eps = lambda s: " | ".join(f"{P[f'{COL[y]}{PR[s]['eps']}'].value:,.0f}" if isinstance(P[f'{COL[y]}{PR[s]['eps']}'].value, (int, float)) else "—" for y in YRS)
    scn = "\n".join(f"| 情境 {s}（{SCN[s]}） | {f0(g(chr(66+i)+str(VR['dcfr'])))} | {f0(g(chr(66+i)+str(VR['cu'])))} | {f0(g(chr(66+i)+str(VR['tps'])))} |" for i, s in enumerate(SC))
    txt = f"""# physical_ai/nabtesco — Nabtesco（ナブテスコ，6268.T）個股模型 {VER}

{RD['A2'].value}｜build_id `{bid}`｜成品 `dist/{NAME}`（LibreOffice 重算，錯誤值 0）

東證 Prime 上市。本 README 的數字由 `build_nabtesco.py` 從重算後的成品讀出產生，不手寫。

## 命題
{RD['A3'].value}

**結論（條件式）：** {RD['A4'].value}

{RD['A5'].value}

**一句話：** Nabtesco 是 RV 精密減速機的循環股（CMP 分部 FY2026 指引營收 953 億、其中精密減速機 946 億），但只占營收約 28%；另外 72% 是鐵道／航空／船舶（TRS）、自動門（ACB）與包裝機（MFR）。工業機器人上行循環（JARA 受注 +46.6%、公司精密減速機受注 9M +24%）已大致反映在 FY2026 指引與股價；人形機器人主流 BOM 用諧波＋行星，RV 對人形的直接受益小，只在情境 A 以 RV mini 選擇權呈現。

## 估值（評價日 2026-12-31；目標時點約 12 個月）
| 情境 | DCF（推到 12 個月後） | 分部本益比加總 | 情境目標價 |
|---|---|---|---|
{scn}

- 加權目標價 **{f0(tp)} 日圓**（DCF {f0(dcf)}、分部本益比加總 {f0(cmp_)}），現價 {f0(wb['In']['E' + str([r for r in range(1, 400) if wb['In']['A' + str(r)].value == 'px'][0])].value)}（2026-11-13 收盤），上檔 {pc(up)}，估值評等 **{g(f"B{VR['rt']}")}**；收支關卡 {g(f"B{VR['gj']}")} → 投資建議 **{g(f"B{VR['rec']}")}**。
- 方法：DCF 標準版 FCFF（WACC {V['B' + str(VR['pe'] - 1)].value}，為 WACC／ke／g／RONIC）50%＋分部本益比加總 50%（CMP 用機器人與 FA 同業 6 家中位數、TRS 用 Knorr-Bremse 與川崎重工、ACB 用 ASSA ABLOY 與 dormakaba、MFR 與全社費用用全體 10 家中位數；FY2028 分部營業利益 × 稅後）。相對倍數係數 1.00（未調整：缺 Nabtesco 對同業的歷史相對倍數）。
- 隱含前瞻 P/E（加權目標價 ÷ 情境 B FY2028 EPS）{g(f"B{VR['pe']}"):.1f} 倍；現行前瞻 P/E 18.40 倍（資料商 NTM）。
- **2026-11-16 盤中股價 3,651（−12.0%，第 3 季決算後），非收盤、未採用**；若以此價計，上檔約 {pc(tp/3651-1)}，評等仍落在中立區間（+25% 以下）。

## 營收分部（官方報告分部；PL 頁三情境，情境 B 摘要，¥B）
| 列 | {" | ".join(f"FY{y}" for y in YRS)} |
|---|{"---|" * len(YRS)}
| 精密減速機（CMP 子列；命題模組） | {yr('B','prg')} |
| CMP 營收 | {yr('B','cmp')} |
| TRS 營收 | {yr('B','trs')} |
| ACB 營收 | {yr('B','acb')} |
| MFR（その他）營收 | {yr('B','mfr')} |
| 合併營收 | {yr('B','rev')} |
| 營業利益 | {yr('B','op')} |
| 歸屬母公司淨利 | {yr('B','ni')} |
| EPS（日圓） | {eps('B')} |

- FY2025 實績（持續營業口徑；油壓機器事業 2025 Q3 起列停業單位、2026-01-01 轉讓 70% 給 Comer Industries 合資）；FY2026＝公司 2026-07-31 上修、11-13 維持的指引；FY2027 起為模型。
- 命題模組：精密減速機＝機器人用（前一年 ×(1＋全球工業機器人需求年增)×(1＋份額＋價格變化)）＋非機器人用（工具機、半導體設備、一般產業）＋RV mini／實體 AI 新品；三子列＋CMP 其他＝CMP，核對列為 0。
- 分部營業利益：CMP＝FY2026 指引 ＋ 邊際利益率 mm_cmp × ΔCMP 營收（9M 實績推得 28.8%）；TRS、ACB、MFR 用 FY2026 指引利益率持平；全社費用隨營收等比例。

## 主要假設（In 頁；Assumed／Analogy 皆給區間）
- 全球工業機器人需求年增（情境 A／B／C）：FY2027 15%／7.2%／−5%；FY2028–29 9%／7.2%／4%；FY2030–32 7%／5%／3%（B 以 IFR 2026→2029 年複合 7.2% 為類比）。
- RV 份額＋價格每年 0%／−1%／−3%（中國 RV 廠替代）；非機器人用年增 6%／4%／1%；RV mini／實體 AI 新品 FY2032 營收 10／3／0.5（¥B）；情境 C FY2028 起 CMP 利益率 −2pt。
- 精密減速機機器人用比例 70%（60–80%，未取得公司揭露）；TRS 年增 7%／5%／2%、ACB 5%／3.5%／1.5%、MFR 5%／3%／0%。
- 折現：rf 2.817%（日本 10 年期，與東京精密 v0.3.6 同值）、β {wb['In']['E' + str([r for r in range(1, 400) if wb['In']['A' + str(r)].value == 'beta'][0])].value:.3f}（兩來源 Blume 中位數）、ERP 5%、CRP 0.91%、g＝MIN(3%, rf)、RONIC＝WACC；資本支出 5%、折舊 4.8%、營運資金 30% × 營收增量。
- 評價日股數：2026-09-30 流通股 115.35 百萬股，扣除 11-30 前剩餘買回額度（約 64 億日圓，以現價全數執行；Assumed）。BBT 信託股已在自己株式內，無選擇權。

## 承重 Assumed 驅動與可證偽指標
| 驅動 | 公開可觀察指標 |
|---|---|
| 全球工業機器人需求（rg27、rgm） | JARA 季度受注與生產（訂單÷生產 1.24）、IFR 年報、中國國家統計局月產量 |
| RV 份額＋價格（shr） | 公司精密減速機受注 YoY（26Q3 +18%）、常州稼動率（95%）、雙環傳動／中大力德 RV 營收 |
| CMP 邊際利益率（mm_cmp） | CMP 單季營業利益率（26Q3 約 10.7%）、價格轉嫁金額 |
| TRS、ACB 利益率（m_trs、m_acb） | 分部季報；航空 777X 交付、鐵道訂單 |

失效條件：精密減速機受注 YoY 連兩季轉負；JARA 受注 YoY 轉負；常州稼動率跌破 80%；CMP 單季營業利益率跌破 8%。評等真正翻轉需要的條件見下方反解（單一驅動在合理區間內多半無解，評等穩定在中立）。

## 敏感度（整本重算；Sens 頁；影響最大的 6 項）
| 驅動 | 設定 | 鍵 | 加權目標價 | 相對基準 | 估值評等 |
|---|---|---|---|---|---|
""" + "\n".join(f"| {c['drv']} | {c['set']} | `{c['key']}` | {f0(c['tp'])} | {pc(c['tp']/sens['base']['tp']-1)} | {c['rt']} |" for c in big) + f"""

全部 {len(sens['cases'])} 案見 xlsx Sens 頁；所有案例估值評等皆為「{'、'.join(sorted(set(c['rt'] for c in sens['cases'])))}」。

## 反解（整本重算，二分法；Recon 頁）
| 反解目標 | 鍵 | 反解值 | 基準 | 說明 |
|---|---|---|---|---|
{rv}

## 對帳（線層級變數；xlsx「對帳」頁）
線 2a：IFR 全球裝機 2025 60.3 萬台、預測 2026 65.5 萬、2029 80.6 萬（年複合 7.2%）→ rg27_B、rgm_B；JARA 受注 +46.6% → 情境 A；中國國產比 55% → shr；傳統 6 軸減速機占 BOM 25–35%（Analogy，未進公式）；每台 RV 用量未取得（以營收增速代台數）。線 2b：人形 2026 出貨 >5 萬台（TrendForce）、人形主流 BOM 用諧波 14–16 顆或行星，RV 未列入 → pai_*。與 `physical_ai/lines/` 完成版對帳時，以上口徑不在本模型靜默修改。

## 限制與未取得資料
- **精密減速機用途別（機器人／非機器人）與地區別營收未取得**（公司不揭露），以 Assumed 70% 拆分；每台機器人 RV 用量、RV 平均單價（只有 Derived 粗估 8–10 萬日圓）未取得。
- FY2025 分部營收與營業利益只取得億日圓取整值（第 1 季簡報 10 年圖表）；FY2024 及以前持續營業口徑分部值未取得；FY2025 精密減速機營收未取得。
- 2026-09-30 公司公告過年度訂正（持分法 TMT マシナリー），訂正後 FY2025 淨利未取得；營收與營業利益不受影響（持分法在營業外）。
- 同業前瞻 P/E 取自資料商頁面，日期由 2026-08-28 到 2026-11-13 不一，未換算到同一估值日（違反第 10 節第 7 條的同日要求，列為限制）；各群組少於 8 家，不做迴歸；哈默納科 84 倍為離群值（中位數不受影響）。
- 共識 EPS 逐年值未取得（資料商付費牆），Recon 以現行前瞻 P/E 反推的 NTM EPS 對照。
- 折舊、資本支出、營運資金比率取自資料商（含油壓事業），法定原文未逐項核對。
- 未做：HTML 一頁摘要、tools/model_governance/check_governance.py 與 check_display.py、未參與製作的代理獨立核對（Project 規則第 10 條，待主代理安排）。
- 規格中 Siliconomics／Tokenomics／Wattonomics 協調（coord、自用 AI 讀法、上線時程）不適用於實體 AI 線，略過。DCF 為本模型自寫（未用 `tools/fcff_dcf.py`），公式與 `docs/plan/個股DCF標準規格.md` 相同。

## 檔案
- `build_nabtesco.py`：建置＋整本重算 Sens／反解＋產生本 README 與 `segments.json`、`sens.json`。
- `dist/{NAME}`：成品（工作表 README、In、Hist、Q、PL、DCF、Val、Recon、Sens、對帳、Risk、Sources）。
"""
    open(os.path.join(HERE, "README.md"), "w").write(txt)


def main():
    if "--raw" in sys.argv:
        ov = json.loads(os.environ.get("NAB_OVERRIDE", "{}"))
        build(sys.argv[sys.argv.index("--raw") + 1], ov); return
    base = run_case({}, "base"); print("基準", base); px0 = in_val("px", "base")
    cases = []
    # β 低／高為公式（基準 ± 0.3）：以數值覆寫
    import statistics
    beta_b = statistics.median([0.67 * 0.93 + 0.33, 0.67 * 0.97 + 0.33])
    for k, drv in CASES:
        for side in ("lo", "hi"):
            v = in_val(k, side)
            if k == "beta": v = beta_b + (-0.3 if side == "lo" else 0.3)
            if v is None: continue
            b_ = in_val(k, "base")
            if isinstance(b_, (int, float)) and abs(v - b_) < 1e-12: continue
            r = run_case({k: v}, f"{k}_{side}")
            cases.append(dict(drv=drv + ("（低）" if side == "lo" else "（高）"), set=f"{k}＝{v:.4g}", key=k, **r))
            print(k, side, round(r["tp"]), r["rt"])
    # 終值替代法、情境權重單一情境
    for s in SC:
        w = {f"w_{x}": (1.0 if x == s else 0.0) for x in SC}
        r = run_case(w, f"only_{s}")
        cases.append(dict(drv=f"只看情境 {s}", set=f"w_{s}＝1，其餘 0", key=f"w_{s}", **r))
    up_ = base["alt_tp"] / px0 - 1
    cases.append(dict(drv="終值替代法（DCF 終值＝FCFF(FY2032)×(1＋g)÷(WACC − g)）", set="DCF 頁替代法列", key="tv_alt", tp=base["alt_tp"], dcf=base["alt_dcf"], cmp=base["cmp"], up=up_,
                      rt="強力買進" if up_ > 0.5 else "買進" if up_ > 0.25 else "中立" if up_ >= -0.25 else "放空" if up_ >= -0.5 else "強力放空"))
    # 反解：rgm_B（FY2028–29 機器人需求年增，情境 B）使目標價＝現價、升一級、降一級
    px = 4149
    def solve(target_up, key="rgm_B", lo=-0.50, hi=0.60):
        f = lambda x: run_case({key: x}, "rev")["up"] - target_up
        flo, fhi = f(lo), f(hi)
        if flo * fhi > 0: return None
        for _ in range(16):
            mid = (lo + hi) / 2; fm = f(mid)
            if flo * fm <= 0: hi, fhi = mid, fm
            else: lo, flo = mid, fm
        return (lo + hi) / 2
    revs = []
    up0 = base["up"]
    tiers = [0.5, 0.25, -0.25, -0.5]
    upper = min([t for t in tiers if t > up0], default=None); lower = max([t for t in tiers if t <= up0], default=None)
    for lab, tgt in (("目標價＝現價（上檔 0%）", 0.0), (f"升一級邊界（上檔 {upper:+.0%}）" if upper is not None else None, upper), (f"降一級邊界（上檔 {lower:+.0%}）" if lower is not None else None, lower)):
        if lab is None: continue
        x = solve(tgt)
        revs.append(dict(label=lab, key="rgm_B", val=x if x is not None else "區間 −50%～+60% 內無解", base=0.072,
                         note="情境 B 的 FY2028–FY2029 全球工業機器人需求年增（其他不變，整本重算；情境 A／C 不動）"))
        print("反解", lab, x)
    for key, b0, lo_, hi_, note in (("shr_B", -0.01, -0.25, 0.25, "情境 B 機器人用 RV 份額＋價格年變化（FY2027–FY2032 每年；情境 A／C 不動）"),
                                    ("mm_cmp", 0.29, -0.5, 1.5, "CMP 邊際營業利益率（三情境同動）")):
        for lab, tgt in (("目標價＝現價（上檔 0%）", 0.0), (f"升一級邊界（上檔 {upper:+.0%}）", upper), (f"降一級邊界（上檔 {lower:+.0%}）", lower)):
            x = solve(tgt, key, lo_, hi_)
            revs.append(dict(label=lab, key=key, val=x if x is not None else f"區間 {lo_:+.0%}～{hi_:+.0%} 內無解", base=b0, note=note))
            print("反解", key, lab, x)
    sens = dict(base=base, cases=cases, fp=fingerprint())
    json.dump(sens, open(os.path.join(HERE, "sens.json"), "w"), ensure_ascii=False, indent=1)
    # 成品
    dist = os.path.join(HERE, "dist"); os.makedirs(dist, exist_ok=True)
    tmp = tempfile.mkdtemp(prefix="nabf_"); raw = os.path.join(tmp, NAME)
    VR, PR, DR = build(raw, {}, sens, revs)
    bid = f"{VER}-{fingerprint()}"
    wb = openpyxl.load_workbook(raw); wb["README"]["A1"] = wb["README"]["A1"].value.replace("{BID}", bid); wb["Sens"]["H2"] = f"build_id {bid}"; wb.save(raw)
    out = recalc(raw, os.path.join(tmp, "o"))
    shutil.copy(out, os.path.join(dist, NAME))
    # 錯誤值檢查
    wbv = openpyxl.load_workbook(os.path.join(dist, NAME), data_only=True)
    errs = [(ws.title, c.coordinate, c.value) for ws in wbv for row in ws.iter_rows() for c in row
            if isinstance(c.value, str) and (c.value.startswith("#") or c.value.startswith("Err:"))]
    print("錯誤值格數", len(errs), errs[:10])
    assert not errs, errs
    write_readme(os.path.join(dist, NAME), VR, PR, DR, bid, sens, revs)
    json.dump(dict(company="nabtesco", ticker="6268.T", model_sheet="Q",
                   official_segments=[dict(name=n, label=l, period="Q1'26（單季）", value=v, unit="¥B", source=TS2Q + "；" + ENV)
                                      for n, l, v in (("コンポーネントソリューション", "CMP 營收", 22.259), ("トランスポートソリューション", "TRS 營收", 24.891),
                                                      ("アクセシビリティソリューション", "ACB 營收", 32.186), ("その他（マニュファクチャリングソリューション）", "MFR 營收", 3.696))],
                   sublines=[], check_labels=["檢查：分部合計 − 合併營收"],
                   note="Q 頁單季分部列（累計原文相減）；PL 頁三情境另有「檢查：分部合計 − 合併營收」與「檢查：CMP 子列合計 − CMP」（精密減速機三子列＋CMP 其他＝CMP）。FY2025 分部為億日圓取整，差 −0.01"),
              open(os.path.join(HERE, "segments.json"), "w"), ensure_ascii=False, indent=1)

if __name__ == "__main__":
    main()
