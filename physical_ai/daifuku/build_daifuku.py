# 大福（6383.T）個股模型 v0.1 — AI實體應用 Project（線 5 倉儲物流自動化；並說明線 5 未涵蓋之半導體無塵室搬運）
# 用法：python3 physical_ai/daifuku/build_daifuku.py   → 建置、整本重算 Sens 與反解（LibreOffice 批次）、輸出 dist/ 成品、README.md、sens.json、segments.json
# 會計年度＝曆年（12 月決算；2025 年起由 3 月決算改 12 月決算）。日本基準（J-GAAP），2026 年底起預定改採 IFRS。金額：十億日圓（¥B）；股數：百萬股；EPS、股價：日圓。
# 結構：README、In、Hist、Q（官方報告分部＋客戶產業別）、PL（三情境：產業營收→標準損益表→現金流）、DCF、Val、Recon、Sens、對帳、Risk、Sources。
# 不讀 Siliconomics／Tokenomics／Wattonomics 協調檔。DCF 為本檔自寫，公式同 docs/plan/個股DCF標準規格.md。
import sys, os, json, subprocess, shutil, hashlib, tempfile, statistics
import openpyxl
from openpyxl.styles import Font, PatternFill
from openpyxl.workbook.defined_name import DefinedName

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(os.path.dirname(HERE))
VER = "v0.1"
DATE = "20261124"
NAME = f"{DATE}_Daifuku大福個股模型_{VER}.xlsx"
BLUE = Font(color="0000FF"); BOLD = Font(bold=True)
HDR = PatternFill("solid", fgColor="DDEBF7")
SC = ("A", "B", "C"); SCN = {"A": "樂觀", "B": "基準", "C": "保守"}
YRS = list(range(2023, 2033)); COL = {y: "CDEFGHIJKL"[i] for i, y in enumerate(YRS)}   # PL：C＝FY2023 … L＝FY2032

# ---------------- 出處 ----------------
TS3Q = "公司 FY2026 第 3 四半期決算短信〔J-GAAP〕（英文版，2026-11-11）https://www.daifuku.com/ir/assets/FY2026Q3e.pdf"
TS2Q = "公司 FY2026 中間期決算（2026-08-06）；資料商彙整 Quartr（https://quartr.com/events/daifuku-co-ltd-6383-q2-2027_F9NWHRfZ）、LNEWS（https://www.lnews.jp/2026/08/s0807501.html）"
LN25 = "LNEWS 報導公司 FY2025（2025 年 12 月期）決算（2026-02-12）https://www.lnews.jp/2026/02/s0212505.html"
SAF = "stockanalysis.com Daifuku 財務頁（資料商，2026-11-24 查閱；FY2023＝2024 年 3 月期）https://stockanalysis.com/quote/tyo/6383/financials/"
SAS = "stockanalysis.com Daifuku statistics（資料商，2026-11-24 頁面；價 4,372、前瞻 P/E 13.85、β 1.04、賣方 15 位平均目標價 7,328.67）https://stockanalysis.com/quote/tyo/6383/statistics/"
YCL = "Yahoo Finance 日線 2026-11-20 收盤（physical_ai/data/closes2_20261109_20261120.csv）"
JGB10 = "日本財務省 國債金利情報 10 年：2026-11-02 2.817%（與 physical_ai/nabtesco、thk 同一取值，跨模型一致；逾期）"
CO = "發行公司（立場：自家業績與前景）"
DV = "資料商彙整賣方共識（立場：中立彙整，但來源為賣方）"
L5 = "physical_ai/lines/warehouse/dist/20261116_倉儲物流單位經濟_v0.1.xlsx"
THX = "physical_ai/thk/dist/20261123_THK個股模型_v0.1.xlsx"
KGS = "stockanalysis.com KION statistics（2026-11-23 價 30.66 歐元、前瞻 P/E 8.17）https://stockanalysis.com/quote/etr/KGX/statistics/"
ATS = "stockanalysis.com AutoStore statistics（2026-11-23 價 15.55 挪威克朗、前瞻 P/E 24.04）https://stockanalysis.com/quote/osl/AUTO/statistics/"

# ---------------- In 列 ----------------
# (鍵, 項目, 單位, 低, 基準, 高, 標記, 來源1, 來源2／說明, 來源誘因)
ROWS = []
def sec(t): ROWS.append(("sec", t))
def inp(k, lab, u, lo, b, hi, tag, s1, s2="", inc=""): ROWS.append((k, lab, u, lo, b, hi, tag, s1, s2, inc))

IND = [("el", "電子（半導體、面板廠的無塵室搬運）"), ("co", "商業與零售"), ("au", "汽車與零組件"), ("ap", "機場"), ("tw", "運輸與倉儲"), ("ot", "其他產業（機械、化學醫藥、食品、鋼鐵、精密機器等）")]
sec("A. FY2025 實績：客戶產業別營收（公司揭露；LNEWS 引述，單位取整至 0.1 十億）")
for (k, n), v in zip(IND[:5], (251.0, 133.4, 85.7, 76.8, 29.3)):
    inp(f"a25_{k}", f"營收 FY2025：{n}", "¥B", None, v, None, "Verified", LN25 + "（產業別營收，億日圓取整）", "合計與合併營收 660.724 差由「其他產業」吸收")
inp("a25_ot", "營收 FY2025：其他產業（＝合併營收 − 上列五項）", "¥B", None, "=660.724-(251.0+133.4+85.7+76.8+29.3)", None, "Derived", LN25 + "；合併營收 660.724（" + SAF + "）", "含機械、化學醫藥、食品、鋼鐵、精密機器等，及合併調整")

sec("B. FY2026 錨點：9M 實績（客戶產業別）與公司指引（2026-11-11 第二次上修；Interested-party）")
for (k, n), v in zip(IND[:5], (226.680, 95.133, 70.168, 65.244, 21.369)):
    inp(f"a9_{k}", f"營收 9M'26：{n}", "¥B", None, v, None, "Verified", TS3Q + " 營收分解（產業別，外部客戶）", "2026-07 起含 Eisenmann（汽車塗裝系統）", "")
inp("a9_ot", "營收 9M'26：其他產業（＝合併營收 553.776 − 上列五項）", "¥B", None, "=553.776-(226.680+95.133+70.168+65.244+21.369)", None, "Derived", TS3Q, "含合併調整 0.726")
inp("a9_rev", "營收 9M'26：合併", "¥B", None, 553.776, None, "Verified", TS3Q + "：553,776（+13.9%）", "", "")
inp("g26_rev", "公司指引：FY2026 營收", "¥B", None, 770.0, None, "Interested-party", TS3Q + "：770,000（+16.5%）", "指引沿革：2026-02 700.0 → 08 735.0 → 11 770.0；Q4 隱含 216.2（Derived）；含 Eisenmann 與部分無塵室專案提前認列", CO)
inp("g26_op", "公司指引：FY2026 營業利益", "¥B", None, 119.0, None, "Interested-party", TS3Q + "：119,000（+18.0%）", "前次 113.0；9M 實績 83.487", CO)
inp("g26_pbt", "FY2026 稅前利益（9M 實績 87.992 ＋ 指引隱含 Q4 經常利益 36.072）", "¥B", None, "=87.992+(122.0-85.928)", None, "Derived", TS3Q + "：經常利益指引 122,000、9M 85,928、9M 稅前 87,992", "Q4 特別損益視為 0")
inp("g26_ni", "公司指引：FY2026 歸屬母公司淨利", "¥B", None, 89.5, None, "Interested-party", TS3Q + "：89,500；EPS 241.95", "9M 實績 63.284", CO)
inp("g26_eps", "公司指引：FY2026 基本 EPS", "¥", None, 241.95, None, "Interested-party", TS3Q + "：241.95", "9M 實績 171.15；稀釋後 164.07（可轉債）", CO)
inp("g26_gp", "FY2026 毛利（假設 Q4 毛利率＝9M'26 的 25.1%）", "¥B", None, "=g26_rev*139.070/553.776", None, "Derived", TS3Q + "：9M 毛利 139,070 ÷ 營收 553,776", "公司未揭露全年毛利指引；9M'25 毛利率 24.6%、FY2025 24.5%")
inp("g26_sga", "FY2026 銷管費用（＝毛利 − 營業利益指引）", "¥B", None, "=g26_gp-g26_op", None, "Derived", TS3Q + "：9M 銷管 55.583（+25.7%，含研發、Eisenmann）", "")
inp("g26_nci", "FY2026 非支配股東損益（9M 實績約 0.007）", "¥B", None, 0.009, None, "Derived", TS3Q + "：淨利 63,291 − 歸屬母公司 63,284", "")
inp("g26_ord", "公司指引：FY2026 訂單（中值；區間 970–1,000）", "¥B", None, 985.0, None, "Interested-party", TS3Q + "：970,000–1,000,000（前次 860–900）；9M 實績 781.230（+56.5%）", "訂單 ÷ 營收約 1.28；只作 Recon 對照，不進公式", CO)
inp("da_26", "FY2026 折舊（9M 實績 10.623 ×12÷9）", "¥B", None, "=10.623*12/9", None, "Derived", TS3Q + " 現金流量表：9M 折舊 10,623（9M'25 6,465；新租賃會計準則與 Eisenmann 使折舊上升）", "公司 2026-02 計畫 10.4（LNEWS）已被超過")
inp("cx_26", "FY2026 資本支出（9M 有形固定資產取得 19.371 ×12÷9）", "¥B", None, "=19.371*12/9", None, "Derived", TS3Q + " 現金流量表：9M 取得 19,371", "公司 2026-02 計畫 18.8（LNEWS）；美國產能倍增")
inp("dps_26", "FY2026 每股股利（中間 40＋期末 50 預定）", "¥", None, 90, None, "Interested-party", TS3Q + "：年間 90（前年 78）", "", CO)
inp("dps_q4", "FY2026 期末股利（2027 年支付）", "¥", None, 50, None, "Interested-party", TS3Q, "", CO)
inp("g_anc", "FY2026 實績相對公司指引（營收、毛利、費用同比例）", "%", -0.02, 0.0, 0.02, "Assumed", "基準＝公司指引（9M 已實績，偏離只來自 Q4）", "公司 2026 年已兩度上修；低／高 ±2%")
inp("mp27_rev", "公司中期計畫：2027 年營收目標", "¥B", None, 800.0, None, "Interested-party", LN25 + "：中計 2027 營收 8,000 億、營業利益 1,200 億（2026-02 上修）", "Recon 對照用，不進公式", CO)
inp("mp30_rev", "公司長期願景：2030 年營收目標", "¥B", None, 1000.0, None, "Interested-party", LN25 + "：2030 營收 1 兆、營業利益 1,500 億（利益率 15%）", "Recon 對照用，不進公式", CO)
inp("mp30_op", "公司長期願景：2030 年營業利益目標", "¥B", None, 150.0, None, "Interested-party", LN25, "Recon 對照用，不進公式", CO)

sec("C. 命題模組：產業營收驅動（電子＝半導體資本支出循環；倉儲物流＝線 5 AS/RS 市場成長 × 公司參與度；其餘＝一般成長）")
inp("cap_il", "倉儲物流（商業與零售、運輸與倉儲）營收成長 ÷ 線 5 AS/RS 市場成長（參與度）", "x", 0.2, 0.5, 1.0, "Assumed",
    "9M'26 商業與零售營收 95.1 vs 9M'25 99.6（−4.5%）、運輸與倉儲 21.4 vs 21.5（−0.7%），同期線 5 市場 2026 年成長約 +2%～+8%（" + L5 + " 下游取數!D8:D9）", "市場是 AS/RS 全市場（含製造、汽車）；大福主力為堆垛機式與系統整合，對照不完全；1.0＝與市場同速")
for yy, rr in ((2026, 8), (2028, 9), (2030, 10), (2032, 11)):
    for s, col in (("A", "E"), ("B", "D"), ("C", "C")):
        v = {("A", 2026): 10.9294503319294, ("B", 2026): 10.6943372354646, ("C", 2026): 10.4524071462717,
             ("A", 2028): 13.4289214189686, ("B", 2028): 12.5807865845056, ("C", 2028): 11.7461386987719,
             ("A", 2030): 16.5, ("B", 2030): 14.8, ("C", 2030): 13.2,
             ("A", 2032): 19.60365, ("B", 2032): 16.94452, ("C", 2032): 14.27712}[(s, yy)]
        inp(f"mk{str(yy)[2:]}_{s}", f"情境 {s}：AS/RS 市場 {yy}（線 5；{'高' if s == 'A' else '基準' if s == 'B' else '低'}）", "十億美元", None, v, None, "Derived",
            f"{L5} 下游取數!{col}{rr}（MarketsandMarkets 口徑）", "年間以幾何插值（Derived）；情境 A＝線模型高、B＝基準、C＝低")
GTXT = {
    "ge27": ("FY2027", {"A": (0.32, "樂觀：訂單 2026 指引 970–1,000 ÷ 營收 770 ≈ 1.28；9M 訂單 781 已超過去年全年；中間期受注殘超過 7,400 億（" + TS2Q + "）"),
                        "B": (0.20, "基準：受注殘推估 2026 年底約 8,700 億（＝中間期 7,400 ＋ H2 訂單 545 − H2 營收 415，Derived），約 FY2026 營收 1.1 倍；無塵室專案多為跨年認列"),
                        "C": (0.05, "保守：客戶延後認列或取消、專案進度順延")}),
    "ge28": ("FY2028", {"A": (0.12, "樂觀：半導體投資延續"), "B": (0.08, "基準：受注殘消化、新增訂單趨緩"), "C": (-0.12, "保守：半導體資本支出循環下行（以 2023 年記憶體循環下行為類比，幅度 Assumed）")}),
    "gem": ("FY2029–FY2030", {"A": (0.06, ""), "B": (0.04, "基準：與公司 2030 年營收 1 兆願景相容（情境 B 約 1.0 兆）"), "C": (0.00, "")}),
    "gelt": ("FY2031–FY2032", {"A": (0.04, ""), "B": (0.03, "長期：名目半導體廠自動化需求"), "C": (0.02, "")}),
    "gn27": ("FY2027", {"A": (0.06, ""), "B": (0.04, "基準：機場（美國投資）、汽車（客戶延後決策）、其他維持中低成長"), "C": (0.00, "")}),
    "gn28": ("FY2028", {"A": (0.06, ""), "B": (0.04, ""), "C": (-0.03, "保守：汽車、機場投資遞延")}),
    "gnm": ("FY2029–FY2030", {"A": (0.05, ""), "B": (0.04, ""), "C": (0.02, "")}),
    "gnlt": ("FY2031–FY2032", {"A": (0.04, ""), "B": (0.03, ""), "C": (0.02, "")}),
}
GRNG = {"ge27": (0.05, 0.30), "ge28": (-0.05, 0.12), "gem": (0.0, 0.07), "gelt": (0.01, 0.05), "gn27": (0.0, 0.08), "gn28": (0.0, 0.07), "gnm": (0.01, 0.06), "gnlt": (0.01, 0.04)}
for f, (yl, d) in GTXT.items():
    ft = "電子" if f.startswith("ge") else "汽車、機場、其他產業"
    for s in SC:
        v, note = d[s]
        lo, hi = GRNG[f] if s == "B" else (None, None)
        inp(f"{f}_{s}", f"情境 {s}：{ft}營收年增 {yl}", "%/年", lo, v, hi, "Assumed", note or "A／C 以 B ± 約 2pt 設定",
            "區間見情境 A／B／C 三列" if s == "B" else "")
inp("aud27", "汽車營收 FY2027 額外年增（Eisenmann 2026-07 納入，全年化）", "pt", 0.0, 0.08, 0.15, "Assumed", TS3Q + "：Eisenmann 等 9 家 2026-07 納入合併；其全年營收未揭露", "9M'26 其他分部汽車營收 28.0 vs 11.8（Derived，主要為 Eisenmann）；全年化效果以區間處理")

sec("D. 損益表結構（J-GAAP；營業成本、銷管費用公司有揭露）")
inp("mm_gp", "邊際毛利率（Δ毛利 ÷ Δ營收，相對 FY2026）", "x", 0.20, 0.25, 0.30, "Assumed",
    "9M'26 vs 9M'25：(139.070 − 119.429) ÷ (553.776 − 486.014) ＝ 29.0%（Derived，" + TS3Q + "；含無塵室獲利改善）", "FY2025 毛利率 24.5%、FY2023（2024 年 3 月期）19.1%（" + SAF + "）；基準取 25%（接近平均毛利率）以保守看待無塵室比重回落")
for s, v, lo, hi in (("A", 0.005, None, None), ("B", 0.0, -0.01, 0.005), ("C", -0.015, None, None)):
    inp(f"gmadj_{s}", f"情境 {s}：FY2027 起毛利率額外調整", "pt", lo, v, hi, "Assumed", "毛利率：FY2023 19.1%、FY2025 24.5%、9M'26 25.1%（Derived）", "C：材料成本上升與供應限制（公司 2026-11 說明風險）" if s == "C" else "")
inp("sga_el", "銷管費用對營收的彈性（費用年增 ÷ 營收年增）", "x", 0.4, 0.7, 1.0, "Assumed",
    "9M'26：SG&A +25.7% ÷ 營收 +13.9% ＝ 1.8（含 Eisenmann、研發）；FY2025 銷管 60.8 ÷ 營收 660.7（Derived）", "公司 2026 研發（實體 AI 與機器人）計畫 18.2（LNEWS，Interested-party）；基準 0.7")
inp("nonop", "業外與特別損益淨額（FY2027 起；經常外收支＋特別損益）", "¥B", 0.5, 2.5, 4.5, "Assumed", "FY2026 指引隱含 ＋5.1（稅前 124.1 − 營業利益 119.0，Derived）；9M'26 利息收入 3.0、匯損 1.8", "現金 2,566 億之利息約 4／年，扣匯損與其他；不假設證券出售益")
inp("tax_w", "有效稅率（FY2027 起）", "%", 0.25, 0.28, 0.31, "Assumed", "9M'26 28.1%（24.700 ÷ 87.992，Derived，" + TS3Q + "）", "日本法定實效約 30.6%")

sec("E. 現金流與股東回饋")
inp("fc_da", "折舊攤銷 ÷ 營收（FY2027 起）", "%", 0.015, "=da_26/g26_rev", 0.025, "Derived", "FY2026 推估折舊 ÷ 指引營收（見 da_26）", "")
inp("fc_cx", "資本支出 ÷ 營收（FY2027 起）", "%", 0.025, "=cx_26/g26_rev", 0.05, "Derived", "FY2026 推估資本支出 ÷ 指引營收（見 cx_26）", "")
inp("fc_wc", "營運資金增加 ÷ 營收增量", "%", 0.05, 0.15, 0.30, "Assumed", "9M'26：應收 +0.1、存貨 +12.8、應付 +1.7、合約負債 −4.2（現金流量表）→ 增加約 15.4 ÷ 9M 營收增量 67.8 ＝ 23%（Derived）", "合約負債（預收款）80.7 緩衝；基準 15%")
inp("pay", "配息率（股利 ÷ 淨利）", "%", None, 0.385, None, "Interested-party", "中間期決算：預定配息率 38.5%（" + TS2Q + "）；FY2026 年股利 90 ÷ EPS 241.95 ＝ 37.2%", "只用於收支關卡；不進 DCF（合理範圍約 30–50%，不影響目標價）", CO)
inp("eq_26", "FY2026 年底歸屬母公司權益（2026-09-30 權益 501.671 ＋ Q4 淨利）", "¥B", None, "=501.671+(g26_ni-63.284)", None, "Derived", TS3Q + "：2026-09-30 股東權益 501,671", "期末股利（2027-03 支付）不在 2026 扣除")

sec("F. 估值輸入（評價日 2026-12-31；資產負債表 2026-09-30）")
inp("px", "股價（2026-11-20 收盤）", "¥", None, 4129, None, "Verified", YCL + "：6383.T 2026-11-20 收盤 4,129", "未還原股利。決算日（2026-11-11）收盤 5,061（11-11）→ 4,129（11-20）−18.4%；下跌與全球 AI 與半導體股大跌同週，個別原因未查核")
inp("px_date", "股價基準日", "日期", None, "=DATE(2026,11,20)", None, "Verified", "同上", "")
inp("tp_date", "目標價時點（股價日 ＋ 12 個月）", "日期", None, "=EDATE(px_date,12)", None, "Derived", "個股模型原則第 10 節第 26 條", "")
inp("sh_out", "流通股數（2026-09-30；發行 379.830231 − 自己株式 7.604739）", "百萬股", None, "=379.830231-7.604739", None, "Derived", TS3Q + "：期末發行済株式数、自己株式数（Verified）", "含董事報酬信託 0.375 在自己株式內（未另調）；可轉債 449.58 億日圓視為負債，未計稀釋")
inp("sh_v", "評價日股數", "百萬股", None, "=sh_out", None, "Derived", "＝流通股數", "DCF 標準規格第 3 節")
inp("cash_v", "現金及銀行存款（2026-09-30）", "¥B", None, 256.621, None, "Verified", TS3Q + "：256,621", "")
inp("debt_v", "借款及附新股預約權可轉債（2026-09-30；短期借款 1.844＋可轉債 44.958）", "¥B", None, "=1.844+44.958", None, "Verified", TS3Q + "（原值相加）；與資料商總負債 46.8 一致（" + SAS + "）", "可轉債轉換則減債增股；稀釋 EPS 約 −4%（9M 164.07 vs 171.15）")
inp("lease_v", "租賃負債代理（其他非流動負債 15.514 − 2025 年底 8.052）", "¥B", 0.0, "=15.514-8.052", 15.514, "Derived", TS3Q + "：2026 年起適用新租賃會計準則，租賃負債併入「其他」；增加數為代理", "低案 0；高案＝其他非流動負債全額")
inp("mi_v", "非支配股東權益（2026-09-30）", "¥B", None, 0.09, None, "Verified", TS3Q + "：90 百萬日圓", "")
inp("stub_frac", "Stub 期（2026-09-30 → 評價日 2026-12-31）", "年", None, "=(DATE(2026,12,31)-DATE(2026,9,30))/365", None, "Derived", "", "Stub FCFF＝FY2026 FCFF × stub_frac")
inp("rf", "無風險利率（日本 10 年期國債）", "%", None, 0.02817, None, "Verified", JGB10, "")
inp("erp", "股權風險溢酬", "%", 0.04, 0.05, 0.06, "Assumed", "DCF 標準規格第 2 節：預設 5%（敏感度 4–6%）", "")
inp("crp", "日本國家風險溢酬", "%", None, 0.0091, None, "Verified", "Damodaran 國家風險溢酬（2026-01-05）：日本 0.91%（與 physical_ai/nabtesco、thk 同值）", "")
inp("beta", "β（5 年，Blume 調整 0.67β＋0.33）", "x", "=E{r}-0.3", "=0.67*1.04+0.33", "=E{r}+0.3", "Derived", SAS + "：5 年 β 1.04（單一來源）", "敏感度 ±0.3（Assumed）")
inp("kd", "稅前負債成本", "%", None, "=rf+0.005", None, "Assumed", "rf ＋ 0.5pt（可轉債與借款）", "")
inp("g_cap", "永續成長率上限", "%", 0.02, 0.03, 0.04, "Assumed", "DCF 標準規格：3%（敏感度 2–4%），且不超過 rf → g＝MIN(g_cap, rf)", "")
inp("ronic_sp", "RONIC − WACC", "pt", None, 0.0, 0.05, "Assumed", "預設 RONIC＝WACC（新投資價值中性）；高案 +5pt", "")
inp("wacc_adj", "WACC 調整（只用於 Sens ±1pt）", "pt", -0.01, 0.0, 0.01, "Assumed", "", "")
inp("dps_tgt", "評價日到目標時點之每股股利", "¥", None, "=dps_q4+45", None, "Assumed", "FY2026 期末 50（2027-03 支付）＋FY2027 中間（假設 45）", "")
inp("t_tgt", "評價日到目標時點（年）", "年", None, 1.0, None, "Derived", "2026-12-31 → 2027-12-31", "")
inp("w_dcf", "DCF 權重（其餘為 Comparable：同業前瞻 P/E）", "x", 0.3, 0.5, 0.7, "Assumed", "個股模型原則第 4 節預設各半", "")
inp("rel_k", "相對倍數調整係數", "x", None, 1.0, None, "Assumed", "1.00（未調整：缺大福對同業 10 年前瞻 P/E 比值歷史）", "第 10 節第 18、47 條：列為承重輸入")
for s, v in (("A", 0.25), ("B", 0.50), ("C", 0.25)):
    inp(f"w_{s}", f"情境權重 {s}", "x", None, v, None, "Assumed", "DCF 標準規格第 3 節預設 25／50／25", "")

sec("G. 同業前瞻 P/E（全部換算至 2026-11-20 收盤，與標的同一天；收盤檔沒有的同業不進中位數）")
PEERS = [("pe_fanuc", "Fanuc 6954（機器人、FA）", 21.1499, "Derived", "physical_ai/hiwin/dist/20261123_上銀個股模型_v0.1.1.xlsx Comps!F12（已換算至 2026-11-20；收盤檔 6954.T 5,047）"),
         ("pe_yask", "安川電機 6506（伺服、機器人）", 15.6739, "Derived", "同上 Comps!F13（已換算至 2026-11-20；收盤檔 6506.T 3,180）"),
         ("pe_nab", "Nabtesco 6268（精密減速機）", 13.7701, "Derived", "physical_ai/harmonic_drive/dist/20261123_哈默納科個股模型_v0.1.1.xlsx Comp!F5（已換算至 2026-11-20；收盤檔 6268.T 3,105）"),
         ("pe_kion", "KION KGX.DE（堆高機、供應鏈解決方案）", "=29.2/(30.66/8.17)", "Derived", KGS + "；換算：2026-11-20 收盤 29.20（收盤檔）÷ 隱含前瞻 EPS（30.66 ÷ 8.17 ＝ 3.753）"),
         ("pe_auto", "AutoStore AUTO.OL（立方格 AS/RS）", "=14.79/(15.55/24.04)", "Derived", ATS + "；換算：2026-11-20 收盤 14.79（收盤檔）÷ 隱含前瞻 EPS（15.55 ÷ 24.04 ＝ 0.647）")]
for k, lab, v, tag, src in PEERS:
    inp(k, f"前瞻 P/E：{lab}", "x", None, v, None, tag, src, "原值：賣方共識 EPS；P/E 為股權口徑；各家共識 EPS 年度基準（FY1／NTM）不同，只看量級", DV)
inp("pe_med", "同業前瞻 P/E 中位數（5 家；進 Comparable）", "x", 13.7701, "=MEDIAN(pe_fanuc,pe_yask,pe_nab,pe_kion,pe_auto)", 21.1499, "Derived",
    "5 家中位數（報價日皆為 2026-11-20）。SMC 6273（21.35）報價日 2026-10-09、收盤檔無該日價，列「報價日未取得」不納入；Symbotic（SYM）、Ocado（OCDO.L）虧損，前瞻 P/E 不適用；半導體搬運同業（Murata Machinery 等）無上市報價",
    "低＝Nabtesco 13.77、高＝Fanuc 21.15（Derived）；同業少於約 8 家，不做迴歸，區間寬")
inp("pe_self", "大福現行前瞻 P/E（換算至 2026-11-20）", "x", None, "=4129/(4372/13.85)", None, "Derived", SAS + "：2026-11-24 價 4,372、前瞻 P/E 13.85 → 隱含 EPS 315.7；以 2026-11-20 收盤 4,129 換算", "只作對照，不進模型", DV)
inp("cons_tp", "賣方平均目標價（15 位；資料商 2026-11-24 頁面，第 3 季決算後）", "¥", None, 7328.67, None, "Interested-party", SAS, "只作對照，不進模型", DV)



def build(path, ov=None, sens=None, rev=None, bid="{BID}"):
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
    r += 1; IN.cell(r, 1, "情境驅動對照（純展示；修改請改上方藍字）").font = BOLD; r += 1
    hdr(IN, r, ["鍵", "驅動", "", "情境 A（樂觀）", "情境 B（基準）", "情境 C（保守）", "", "說明"]); r += 1
    FAM = [("ge27", "電子營收年增 FY2027"), ("ge28", "電子營收年增 FY2028"), ("gem", "電子營收年增 FY2029–FY2030"), ("gelt", "電子營收年增 FY2031–FY2032"),
           ("gn27", "汽車、機場、其他年增 FY2027"), ("gn28", "汽車、機場、其他年增 FY2028"), ("gnm", "汽車、機場、其他年增 FY2029–FY2030"), ("gnlt", "汽車、機場、其他年增 FY2031–FY2032"),
           ("mk26", "AS/RS 市場 2026（線 5）"), ("mk28", "AS/RS 市場 2028（線 5）"), ("mk30", "AS/RS 市場 2030（線 5）"), ("mk32", "AS/RS 市場 2032（線 5）"), ("gmadj", "FY2027 起毛利率調整")]
    for f, lab in FAM:
        IN.cell(r, 1, f + "_scn"); IN.cell(r, 2, lab)
        for j, s in enumerate(SC): IN.cell(r, 4 + j, f"=E{KR[f + '_' + s]}")
        IN.cell(r, 8, f"引用 In 頁 {f}_A／{f}_B／{f}_C；{lab}"); r += 1
    IN.cell(r, 1, "scn_w"); IN.cell(r, 2, "情境權重")
    for j, s in enumerate(SC): IN.cell(r, 4 + j, f"=E{KR['w_' + s]}")
    IN.cell(r, 8, "引用 In 頁 w_A／w_B／w_C；合計＝1"); r += 1
    IN.column_dimensions["B"].width = 56; IN.column_dimensions["H"].width = 60

    # ---------------- Hist ----------------
    H = wb.create_sheet("Hist")
    H["A1"] = "Hist — 歷史財務（年度；¥B；J-GAAP）。FY2023＝2024 年 3 月期（舊決算期，與後兩年不可比）；FY2024＝2024 曆年（由公司 FY2025 年增率回推的參考值，Derived，細項未取得）；FY2025＝2025 年 12 月期。公式只引用本頁（定義名稱 h_*）"
    hdr(H, 2, ["項目", "FY2023（2024 年 3 月期）", "FY2024（曆年參考）", "FY2025", "標記", "來源"])
    NA = "未取得"
    HR = [("h_rev", "營收", 611.477, "=D3/1.026", 660.724, "Verified", SAF + "；FY2024＝FY2025 營收 ÷ (1＋2.6%)（" + LN25 + "，Derived）"),
          ("h_cogs", "營業成本", "=B3-B5", NA, "=D3-D5", "Derived", "營收 − 毛利"),
          ("h_gp", "毛利", 116.629, NA, 161.651, "Verified", SAF),
          ("h_sga", "銷管費用", "=B5-B7", NA, "=D5-D7", "Derived", "毛利 − 營業利益"),
          ("h_op", "營業利益", 62.080, "=D7/1.244", 100.816, "Verified", SAF + "；FY2024＝FY2025 營業利益 ÷ (1＋24.4%)（" + LN25 + "，Derived）"),
          ("h_pbt", "稅前利益", NA, NA, "=D3*0.1658", "Derived", SAF + "：稅前利益率 16.58% × 營收（頁面未列金額）"),
          ("h_tax", "所得稅", NA, NA, "=D8-D10", "Derived", "稅前 − 淨利（有效稅率約 28.7%）"),
          ("h_ni", "歸屬母公司淨利", 45.461, NA, 78.096, "Verified", SAF),
          ("h_eps", "基本 EPS（日圓）", 118.45, NA, 202.51, "Verified", SAF + "（已調整股票分割）"),
          ("h_da", "折舊", NA, NA, NA, "—", "全年值未揭露（9M'25 6.465、9M'26 10.623，見 " + TS3Q + "）"),
          ("h_cx", "資本支出", NA, NA, NA, "—", "全年值未揭露（9M'25 18.816、9M'26 19.371）")]
    HV = {}
    for i, (k, lab, a, b, c, tag, src) in enumerate(HR, 3):
        H.cell(i, 1, lab); HV[k] = (a, b, c)
        for j, v in enumerate((a, b, c)):
            cc = H.cell(i, 2 + j, v)
            if isinstance(v, (int, float)): cc.font = BLUE
        H.cell(i, 5, tag); H.cell(i, 6, src)
        dn(k + "_23", f"Hist!$B${i}"); dn(k + "_24", f"Hist!$C${i}"); dn(k + "_25", f"Hist!$D${i}")
    H.column_dimensions["A"].width = 44
    HRW = {k: 3 + i for i, (k, *_x) in enumerate(HR)}

    # ---------------- Q（官方報告分部＋客戶產業別） ----------------
    Q = wb.create_sheet("Q")
    Q["A1"] = "Q — 官方報告分部（公司別）9M 累計外部營收與訂單（¥B）。分部對帳：分部外部營收合計 ＋ 合併調整 − 合併營收（四捨五入差 < 0.005）"
    hdr(Q, 2, ["項目", "9M'25", "9M'26", "標記", "來源", "來源誘因"])
    segs = [("Daifuku（日本本體）營收", 193.412, 172.821), ("Contec 營收", 14.964, 17.913), ("Daifuku North America 營收", 120.595, 142.813),
            ("Clean Factomation（韓國）營收", 28.465, 55.038), ("Daifuku (Suzhou) Cleanroom Automation 營收", 29.543, 33.002), ("其他（未列報告分部的子公司）營收", 99.203, 131.460)]
    for i, (lab, a, b) in enumerate(segs, 3):
        Q.cell(i, 1, lab); Q.cell(i, 2, a).font = BLUE; Q.cell(i, 3, b).font = BLUE; Q.cell(i, 4, "Verified"); Q.cell(i, 5, TS3Q + " 營收分解（外部客戶）")
    Q.cell(9, 1, "報告分部＋其他 合計"); Q.cell(9, 2, "=SUM(B3:B8)"); Q.cell(9, 3, "=SUM(C3:C8)"); Q.cell(9, 4, "Derived"); Q.cell(9, 5, "加總")
    Q.cell(10, 1, "合併調整"); Q.cell(10, 2, -0.170).font = BLUE; Q.cell(10, 3, 0.726).font = BLUE; Q.cell(10, 4, "Verified"); Q.cell(10, 5, TS3Q)
    Q.cell(11, 1, "合併營收"); Q.cell(11, 2, 486.014).font = BLUE; Q.cell(11, 3, 553.776).font = BLUE; Q.cell(11, 4, "Verified"); Q.cell(11, 5, TS3Q)
    Q.cell(12, 1, "檢查：分部合計 ＋ 合併調整 − 合併營收（四捨五入差 < 0.005）"); Q.cell(12, 2, "=ROUND(B9+B10-B11,3)"); Q.cell(12, 3, "=ROUND(C9+C10-C11,3)"); Q.cell(12, 4, "Derived"); Q.cell(12, 5, "公司以百萬日圓捨去，分項加總有 1–3 百萬日圓差")
    Q.cell(13, 1, "訂單（合併）"); Q.cell(13, 2, "=C13/1.565").font = BLUE; Q.cell(13, 3, 781.230).font = BLUE; Q.cell(13, 4, "Verified"); Q.cell(13, 5, TS3Q + "：9M'26 781,230（+56.5%）；9M'25＝781.230 ÷ 1.565（Derived）")
    Q.cell(14, 1, "訂單 ÷ 營收（9M）"); Q.cell(14, 2, "=B13/B11"); Q.cell(14, 3, "=C13/C11"); Q.cell(14, 4, "Derived"); Q.cell(14, 5, "全年指引 985 ÷ 770 ＝ 1.28")
    Q.cell(15, 1, "註：報告分部為公司別（含日本與海外子公司），PL 頁營收列改用公司揭露的客戶產業別營收（同一份決算短信營收分解表），兩者口徑不同；Clean Factomation 9M 訂單 121.7（+199.2%）、Suzhou 65.1（+190.2%）")
    hdr(Q, 17, ["客戶產業別（外部客戶）", "9M'25", "9M'26", "標記", "來源", "來源誘因"])
    inds = [("電子", 186.346, 226.680), ("商業與零售", 99.613, 95.133), ("汽車與零組件", 62.811, 70.168), ("機場", 53.684, 65.244), ("運輸與倉儲", 21.523, 21.369),
            ("其他產業（機械、化學醫藥、食品、鋼鐵、精密、其他）", 62.037, 75.182)]
    for i, (lab, a, b) in enumerate(inds, 18):
        Q.cell(i, 1, lab); Q.cell(i, 2, a).font = BLUE; Q.cell(i, 3, b).font = BLUE; Q.cell(i, 4, "Verified" if i < 23 else "Derived"); Q.cell(i, 5, TS3Q + " 營收分解（產業別）")
    Q.cell(24, 1, "檢查：產業別合計 − 合併營收"); Q.cell(24, 2, "=ROUND(SUM(B18:B23)-B11,3)"); Q.cell(24, 3, "=ROUND(SUM(C18:C23)-C11,3)"); Q.cell(24, 4, "Derived"); Q.cell(24, 5, "其他產業含合併調整")
    Q.column_dimensions["A"].width = 46

    # ---------------- PL ----------------
    P = wb.create_sheet("PL")
    P["A1"] = "PL — 三情境：產業營收（電子＝半導體循環；倉儲物流＝線 5 市場 × 參與度；汽車、機場、其他＝一般成長）→ 標準損益表（J-GAAP）→ 現金流（¥B；FY2023 舊決算期、FY2024 參考值、FY2025 實績、FY2026 公司指引 ×(1＋g_anc)、FY2027 起模型）"
    PR = {}
    rr = 3
    KEYS = [f"c_{k}" for k, _ in IND] + ["mkt", "rev", "chk1", "gr", "cogs", "gp", "gm", "sga", "oth", "opex", "op", "opm", "fin", "pbt", "tax", "cont",
             "disc", "nci", "ni", "nim", "sh", "eps", "epsc", "nig", "da", "cx", "dwc", "fcff", "eq", "div", "fcfad"]
    LAB = ([f"營收：{n}（客戶產業別）" for _, n in IND] +
           ["參考：AS/RS 市場（十億美元，線 5）", "營收（合併）", "檢查：產業別合計 − 營收", "營收年增率", "營業成本", "毛利", "毛利率", "銷售及管理費用",
            "其他營業損益淨額（J-GAAP 無此列，設 0）", "營業費用（銷管費用 − 其他營業損益淨額）",
            "營業利益", "營業利益率", "業外與特別損益淨額（稅前 − 營業利益）", "稅前利益", "所得稅", "稅後淨利（含非支配）", "停業單位損益（無）",
            "非支配股東損益", "稅後淨利（歸屬母公司）", "淨利率", "平均股數（百萬股）", "EPS（基本，日圓）", "EPS（持續營業，日圓）", "淨利年增率",
            "折舊攤銷", "資本支出", "營運資金增加", "FCFF（NOPAT＋折舊 − 資本支出 − 營運資金增加）", "歸屬母公司權益（年底）", "股利", "自由現金流扣股利（收支關卡用）"])
    ICO = f"c_{IND[0][0]}"; ICL = f"c_{IND[-1][0]}"
    for s in SC:
        P.cell(rr, 1, f"情境 {s}（{SCN[s]}）").font = BOLD
        hdr(P, rr + 1, ["項目", "說明"] + [f"FY{y}" for y in YRS]); b = rr + 2
        R = {k: b + i for i, k in enumerate(KEYS)}
        for i, l in enumerate(LAB): P.cell(b + i, 1, l)
        for y in YRS:
            c = COL[y]; p = COL.get(y - 1); cell = lambda k: f"{c}{R[k]}"; prev = lambda k: f"{p}{R[k]}"
            def w(k, f): P[cell(k)] = f
            NAS = "未取得"
            if y <= 2025:
                i3 = {2023: 0, 2024: 1, 2025: 2}[y]; yy = str(y)[2:]
                ok = lambda k: isinstance(HV[k][i3], (int, float)) or (isinstance(HV[k][i3], str) and HV[k][i3].startswith("="))
                for k in ("rev", "cogs", "gp", "sga", "op", "pbt", "tax", "ni", "eps", "da", "cx"):
                    w(k, f"=h_{k}_{yy}" if ok("h_" + k) else NAS)
                w("gm", f"={cell('gp')}/{cell('rev')}" if ok("h_gp") else NAS)
                w("opex", f"={cell('sga')}" if ok("h_sga") else NAS); w("oth", 0)
                w("opm", f"={cell('op')}/{cell('rev')}" if ok("h_op") else NAS)
                w("fin", f"={cell('pbt')}-{cell('op')}" if ok("h_pbt") else NAS)
                w("cont", f"={cell('pbt')}-{cell('tax')}" if ok("h_tax") else NAS)
                w("disc", 0); w("nci", 0 if ok("h_ni") else NAS)
                w("nim", f"={cell('ni')}/{cell('rev')}" if ok("h_ni") else NAS)
                w("sh", f"={cell('ni')}/{cell('eps')}*1000" if ok("h_ni") and ok("h_eps") else NAS)
                w("epsc", f"={cell('eps')}" if ok("h_eps") else NAS)
                if y == 2025:
                    for k, _ in IND: w(f"c_{k}", f"=a25_{k}")
                    w("gr", f"={cell('rev')}/{prev('rev')}-1"); w("nig", f"={cell('ni')}/{prev('ni')}-1")
                    w("gr", '=0.026'); P[cell("gr")].number_format = "0.0%"
                    w("chk1", f"=ROUND(SUM({c}{R[ICO]}:{c}{R[ICL]})-{cell('rev')},3)")
                    w("nig", '="不可比（FY2024 細項未取得；公司揭露 FY2025 營收 +2.6%、營業利益 +24.4%）"')
                    w("gr", '=0.026')
                elif y == 2024:
                    w("gr", NAS); w("nig", NAS); w("chk1", NAS)
                    for k, _ in IND: w(f"c_{k}", NAS)
                else:
                    w("gr", NAS); w("nig", NAS); w("chk1", NAS)
                    for k, _ in IND: w(f"c_{k}", NAS)
                continue
            if y == 2026:
                for k, _ in IND: w(f"c_{k}", f"=a9_{k}*g26_rev/a9_rev*(1+g_anc)")
                w("gp", "=g26_gp*(1+g_anc)"); w("sga", "=g26_sga*(1+g_anc)"); w("oth", 0)
                w("fin", "=g26_pbt-g26_op"); w("disc", 0); w("nci", "=g26_nci")
                w("tax", f"={cell('pbt')}-(g26_ni+g26_nci)*(1+g_anc)")
                w("sh", "=g26_ni/g26_eps*1000")
                w("da", "=da_26"); w("cx", "=cx_26"); w("eq", "=eq_26"); w("div", f"=dps_26*{cell('sh')}/1000")
                w("mkt", f"=mk26_{s}")
            else:
                fe = "ge27" if y == 2027 else "ge28" if y == 2028 else "gem" if y <= 2030 else "gelt"
                fn = "gn27" if y == 2027 else "gn28" if y == 2028 else "gnm" if y <= 2030 else "gnlt"
                if y == 2027: mk = f"=mk26_{s}*(mk28_{s}/mk26_{s})^(({y}-2026)/2)"
                elif y == 2028: mk = f"=mk28_{s}"
                elif y == 2029: mk = f"=mk28_{s}*(mk30_{s}/mk28_{s})^0.5"
                elif y == 2030: mk = f"=mk30_{s}"
                elif y == 2031: mk = f"=mk30_{s}*(mk32_{s}/mk30_{s})^0.5"
                else: mk = f"=mk32_{s}"
                w("mkt", mk)
                w("c_el", f"={prev('c_el')}*(1+{fe}_{s})")
                for k in ("co", "tw"): w(f"c_{k}", f"={prev('c_' + k)}*(1+cap_il*({cell('mkt')}/{prev('mkt')}-1))")
                w("c_au", f"={prev('c_au')}*(1+{fn}_{s}" + ("+aud27)" if y == 2027 else ")"))
                for k in ("ap", "ot"): w(f"c_{k}", f"={prev('c_' + k)}*(1+{fn}_{s})")
                w("gp", f"=$F{R['gp']}+mm_gp*({cell('rev')}-$F{R['rev']})+gmadj_{s}*{cell('rev')}")
                w("sga", f"=$F{R['sga']}*(1+sga_el*({cell('rev')}/$F{R['rev']}-1))"); w("oth", 0)
                w("fin", "=nonop"); w("tax", f"={cell('pbt')}*tax_w"); w("disc", 0); w("nci", 0); w("sh", "=sh_v")
                w("da", f"=fc_da*{cell('rev')}"); w("cx", f"=fc_cx*{cell('rev')}")
                w("div", f"=pay*{cell('ni')}"); w("eq", f"={prev('eq')}+{cell('ni')}-{cell('div')}")
            w("rev", f"=SUM({c}{R[ICO]}:{c}{R[ICL]})")
            w("chk1", f"=ROUND(SUM({c}{R[ICO]}:{c}{R[ICL]})-{cell('rev')},6)")
            w("gr", f"={cell('rev')}/{prev('rev')}-1")
            w("cogs", f"={cell('rev')}-{cell('gp')}"); w("gm", f"={cell('gp')}/{cell('rev')}")
            w("opex", f"={cell('sga')}-{cell('oth')}"); w("op", f"={cell('gp')}-{cell('opex')}"); w("opm", f"={cell('op')}/{cell('rev')}")
            w("pbt", f"={cell('op')}+{cell('fin')}")
            w("cont", f"={cell('pbt')}-{cell('tax')}"); w("ni", f"={cell('cont')}+{cell('disc')}-{cell('nci')}")
            w("nim", f"={cell('ni')}/{cell('rev')}"); w("eps", f"={cell('ni')}/{cell('sh')}*1000"); w("epsc", f"=({cell('cont')}-{cell('nci')})/{cell('sh')}*1000")
            w("nig", f"={cell('ni')}/{prev('ni')}-1" if y > 2025 else "")
            w("dwc", f"=fc_wc*({cell('rev')}-{prev('rev')})")
            w("fcff", f"={cell('op')}*(1-tax_w)+{cell('da')}-{cell('cx')}-{cell('dwc')}")
            w("fcfad", f"={cell('fcff')}-{cell('div')}")
        P.cell(b, 2, "FY2027 起＝前一年 ×(1＋成長)")
        P.cell(R["c_co"], 2, "FY2027 起＝前一年 ×(1＋參與度 × 線 5 市場成長)")
        P.cell(R["c_el"], 2, "FY2027 起＝前一年 ×(1＋電子年增)；FY2026＝9M 產業結構 × 指引營收")
        P.cell(R["gp"], 2, "FY2027 起＝FY2026 毛利 ＋ 邊際毛利率 × Δ營收 ＋ 毛利率調整 × 營收")
        P.cell(R["sga"], 2, "FY2027 起＝FY2026 ×(1＋彈性 × 營收相對 FY2026 增幅)")
        P.cell(R["tax"], 2, "FY2026＝稅前 − 指引淨利（Derived）")
        P.cell(R["fcff"], 2, "NOPAT＝營業利益 ×(1 − tax_w)；股票報酬未取得（不加回）")
        P.cell(R["div"], 2, "FY2027 起＝配息率 × 淨利")
        P.cell(R["gr"], 2, "FY2025：公司揭露 +2.6%（對 2024 曆年）")
        for k in ("gr", "gm", "opm", "nim", "nig"):
            for y in YRS: P[f"{COL[y]}{R[k]}"].number_format = "0.0%"
        PR[s] = R; rr = b + len(LAB) + 2
    P.column_dimensions["A"].width = 52; P.column_dimensions["B"].width = 30
    for s in SC:
        for k in ("rev", "op", "ni", "eps", "fcff", "fcfad"):
            dn(f"pl_{k}_{s}", f"PL!$C${PR[s][k]}:$L${PR[s][k]}")


    # ---------------- DCF ----------------
    D = wb.create_sheet("DCF")
    D["A1"] = "DCF — 標準版 FCFF（NOPAT＋折舊 − 資本支出 − 營運資金增加），WACC 折現，評價日 2026-12-31，明確期 FY2027–FY2032，期中慣例；終值＝NOPAT(FY2033) ×(1 − g ÷ RONIC) ÷ (WACC − g)"
    D["A3"] = "ke（股權成本）"; D["B3"] = "=rf+beta*(erp+crp)"; dn("ke", "DCF!$B$3")
    D["A4"] = "股權市值 E（¥B）"; D["B4"] = "=px*sh_v/1000"
    D["A5"] = "負債 D（借款與可轉債＋租賃代理）"; D["B5"] = "=debt_v+lease_v"
    D["A6"] = "稅後 kd"; D["B6"] = "=kd*(1-tax_w)"
    D["A7"] = "WACC"; D["B7"] = "=(B4*ke+B5*B6)/(B4+B5)+wacc_adj"; dn("wacc", "DCF!$B$7")
    D["A8"] = "永續成長率 g＝MIN(g_cap, rf)"; D["B8"] = "=MIN(g_cap,rf)"; dn("g_tv", "DCF!$B$8")
    D["A9"] = "RONIC"; D["B9"] = "=wacc+ronic_sp"; dn("ronic", "DCF!$B$9")
    D["A10"] = "淨現金（2026-09-30）＝現金 − 借款與可轉債 − 租賃代理 − 非支配"; D["B10"] = "=cash_v-debt_v-lease_v-mi_v"; dn("dcf_nc", "DCF!$B$10")
    D["A11"] = "檢查：WACC 由 In 頁輸入重算 − WACC（應為 0）"; D["B11"] = "=ROUND(((px*sh_v/1000)*(rf+beta*(erp+crp))+(debt_v+lease_v)*kd*(1-tax_w))/((px*sh_v/1000)+debt_v+lease_v)+wacc_adj-wacc,9)"
    for rr_ in range(3, 11): D[f"B{rr_}"].number_format = "0.00%" if rr_ in (3, 6, 7, 8, 9) else "#,##0.0"
    DR = {}; rr = 13
    for s in SC:
        R = PR[s]
        D.cell(rr, 1, f"情境 {s}（{SCN[s]}）").font = BOLD
        hdr(D, rr + 1, ["項目"] + [f"FY{y}" for y in range(2027, 2033)]); b = rr + 2
        rows = ["EBIT（營業利益）", "NOPAT", "＋折舊攤銷", "− 資本支出", "− 營運資金增加", "FCFF", "檢查：各列加總 − FCFF（應為 0）", "折現年數（期中）", "折現因子", "FCFF 現值"]
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
                 ("＋淨現金（2026-09-30）", "=dcf_nc"),
                 ("＋Stub 期 FCFF（FY2026 FCFF × stub_frac）", f"=PL!F{R['fcff']}*stub_frac"),
                 ("− 尚待買回金額（無）", 0),
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
    V["A1"] = "Val — 目標價：DCF（權重 w_dcf）＋ Comparable（同業前瞻 P/E 中位數 × 情境 FY2028 EPS，1 − w_dcf）；三情境 25／50／25；評等五級（> +50% 強力買進、> +25% 買進、≥ −25% 中立、≥ −50% 放空、其餘強力放空）"
    hdr(V, 3, ["項目", "情境 A", "情境 B", "情境 C", "加權", "說明"])
    vr = 4
    def vrow(lab, fs, note=""):
        nonlocal vr
        V.cell(vr, 1, lab)
        for j, f in enumerate(fs): V.cell(vr, 2 + j, f)
        V.cell(vr, 5, f"=w_A*B{vr}+w_B*C{vr}+w_C*D{vr}")
        V.cell(vr, 6, note); vr += 1; return vr - 1
    r_dcf = vrow("DCF 情境價值推到 12 個月後", [f"=dcf_{s}" for s in SC], "DCF 頁")
    r_cu = vrow("Comparable（未調整）：同業前瞻 P/E 中位數 × 情境 FY2028 EPS（持續營業）", [f"=pe_med*PL!H{PR[s]['epsc']}" for s in SC],
                "目標時點的 NTM 以 FY2028 代（2027-11 起算 NTM 約 1/12 FY2027＋11/12 FY2028）；分母：模型各情境 FY2028 持續營業 EPS")
    r_ca = vrow("Comparable（相對倍數調整後，進加權）", [f"={L}{r_cu}*rel_k" for L in "BCD"], "rel_k＝1.00（未調整：缺歷史相對倍數）")
    r_tps = vrow("各情境目標價（w_dcf × DCF ＋ (1 − w_dcf) × Comparable）", [f"=w_dcf*{L}{r_dcf}+(1-w_dcf)*{L}{r_ca}" for L in "BCD"])
    vr += 1
    def one(lab, f, note=""):
        nonlocal vr
        V.cell(vr, 1, lab); V.cell(vr, 2, f); V.cell(vr, 6, note); vr += 1; return vr - 1
    r_tp = one("加權目標價（日圓）", f"=E{r_tps}"); dn("tp", f"Val!$B${r_tp}")
    r_dcft = one("DCF 目標價（情境加權）", f"=E{r_dcf}")
    r_cmpt = one("Comparable 目標價（情境加權，進加權）", f"=E{r_ca}")
    r_cons = one("市場共識目標價（賣方 15 位平均；資料商 2026-11-24 頁面，第 3 季決算後）", "=cons_tp", "只作對照（Interested-party）")
    r_px = one("現價（In 頁 px）", "=px")
    r_up = one("上檔空間", f"=B{r_tp}/px-1"); dn("upside", f"Val!$B${r_up}")
    r_rt = one("估值評等（只看上檔空間）", f'=IF(B{r_up}>0.5,"強力買進",IF(B{r_up}>0.25,"買進",IF(B{r_up}>=-0.25,"中立",IF(B{r_up}>=-0.5,"放空","強力放空"))))')
    r_gate = vr; V.cell(vr, 1, "收支關卡：FY2026–FY2030 自由現金流扣股利合計（¥B；情境 A／B／C）")
    for j, s in enumerate(SC): V.cell(vr, 2 + j, f"=SUM(PL!F{PR[s]['fcfad']}:PL!J{PR[s]['fcfad']})")
    vr += 1
    r_vg = one("入不敷出判定用：情境 B 合計", f"=C{r_gate}")
    r_gj = one("入不敷出判定", f'=IF(B{r_vg}<0,"不通過","通過")')
    r_rec = one("投資建議", f'=IF(AND(B{r_vg}<0,OR(B{r_rt}="買進",B{r_rt}="強力買進")),"收支不足",B{r_rt})')
    r_dist = one("距最近級距邊界（百分點）", f"=MIN(ABS(B{r_up}-0.5),ABS(B{r_up}-0.25),ABS(B{r_up}+0.25),ABS(B{r_up}+0.5))*100")
    r_bnd = one("翻轉目標價：升一級／降一級", f'=IF(B{r_up}>0.5,"已在最高級",TEXT(px*(1+IF(B{r_up}>0.25,0.5,IF(B{r_up}>=-0.25,0.25,IF(B{r_up}>=-0.5,-0.25,-0.5)))),"#,##0"))&"／"&TEXT(px*(1+IF(B{r_up}>0.5,0.5,IF(B{r_up}>0.25,0.25,IF(B{r_up}>=-0.25,-0.25,IF(B{r_up}>=-0.5,-0.5,-9))))),"#,##0")', "升一級：上檔超過該級上緣；降一級：低於該級下緣")
    vr += 1
    r_a1 = one("上檔三段歸因：DCF 貢獻", f"=w_dcf*(B{r_dcft}/px-1)")
    r_a2 = one("上檔三段歸因：未調整 Comparable 貢獻", f"=(1-w_dcf)*(E{r_cu}/px-1)")
    r_a3 = one("上檔三段歸因：相對倍數調整貢獻", f"=(1-w_dcf)*(E{r_ca}-E{r_cu})/px")
    one("檢查：三段合計 − 上檔空間（應為 0）", f"=ROUND(B{r_a1}+B{r_a2}+B{r_a3}-B{r_up},9)")
    r_pd1 = one("期間：模型折現（評價日 2026-12-31 → 2027-12-31，年）", "=t_tgt")
    r_pd2 = one("期間：持有（股價日 → 目標價時點，年）", "=(tp_date-px_date)/365", "上檔空間不是年化報酬")
    r_wk = one("WACC／ke／g／RONIC", '=TEXT(wacc,"0.00%")&"／"&TEXT(ke,"0.00%")&"／"&TEXT(g_tv,"0.00%")&"／"&TEXT(ronic,"0.00%")')
    r_pe = one("隱含前瞻 P/E：加權目標價 ÷ 情境 B FY2028 持續營業 EPS", f"=B{r_tp}/PL!H{PR['B']['epsc']}", "分母：模型情境 B FY2028 持續營業 EPS")
    r_pe1 = one("現價 ÷ 情境 B FY2027 持續營業 EPS", f"=px/PL!G{PR['B']['epsc']}", "分母：模型情境 B FY2027")
    r_pe2 = one("現行前瞻 P/E（資料商，換算至 2026-11-20）÷ 同業中位數", "=pe_self/pe_med", "只作對照；相對重估需要的倍數＝1 ÷ 本值")
    r_hum = one("電子（半導體廠）營收占 FY2030 合併營收（情境 B）", f"=PL!J{PR['B']['c_el']}/PL!J{PR['B']['rev']}")
    for rr_ in (r_up, r_a1, r_a2, r_a3, r_hum): V[f"B{rr_}"].number_format = "0.0%"
    for rr_ in (r_dcf, r_cu, r_ca, r_tps):
        for L in "BCDE": V[f"{L}{rr_}"].number_format = "#,##0"
    for rr_ in (r_tp, r_dcft, r_cmpt): V[f"B{rr_}"].number_format = "#,##0"
    V.column_dimensions["A"].width = 70; V.column_dimensions["F"].width = 60
    VR = dict(tp=r_tp, dcf=r_dcft, cmp=r_cmpt, up=r_up, rt=r_rt, rec=r_rec, gate=r_gate, dist=r_dist, bnd=r_bnd, a1=r_a1, a2=r_a2, a3=r_a3,
              tps=r_tps, dcfr=r_dcf, cu=r_cu, ca=r_ca, pd2=r_pd2, pe=r_pe, pe1=r_pe1, pe2=r_pe2, gj=r_gj, cons=r_cons, wk=r_wk, hum=r_hum)

    # ---------------- Recon ----------------
    RC = wb.create_sheet("Recon")
    RC["A1"] = "Recon — 與公司指引、公司中長期目標、市場共識對照；反解（整本重算）"
    hdr(RC, 2, ["項目", "模型（情境 B）", "公司或共識", "差", "說明"])
    rc = [("FY2026 營收", f"=PL!F{PR['B']['rev']}", "=g26_rev", "公司指引（Interested-party）"),
          ("FY2026 營業利益", f"=PL!F{PR['B']['op']}", "=g26_op", "公司指引"),
          ("FY2026 歸屬母公司淨利", f"=PL!F{PR['B']['ni']}", "=g26_ni", "公司指引"),
          ("FY2026 EPS", f"=PL!F{PR['B']['eps']}", "=g26_eps", "公司指引 241.95"),
          ("FY2027 營收 vs 公司中計 2027 目標", f"=PL!G{PR['B']['rev']}", "=mp27_rev", "公司中期計畫 2027 年營收 8,000 億（2026-02）；FY2026 指引已 7,700 億"),
          ("FY2030 營收 vs 公司 2030 願景", f"=PL!J{PR['B']['rev']}", "=mp30_rev", "公司長期願景 1 兆日圓"),
          ("FY2030 營業利益 vs 公司 2030 願景", f"=PL!J{PR['B']['op']}", "=mp30_op", "公司長期願景 1,500 億（利益率 15%）"),
          ("FY2027 EPS（模型 B） vs 現價 ÷ 現行前瞻 P/E 隱含 EPS", f"=PL!G{PR['B']['epsc']}", "=px/pe_self", "資料商前瞻 EPS（隱含 315.7，年度基準未明）；模型 FY2027 EPS 低於此，代表賣方共識較模型樂觀"),
          ("加權目標價 vs 賣方平均目標價", "=tp", "=cons_tp", "賣方 15 位平均（2026-11-24 頁面，第 3 季決算後；只作對照）"),
          ("FY2026 訂單指引中值 ÷ 營收指引", "=g26_ord/g26_rev", "=1", "訂單 ÷ 營收 > 1 表示受注殘增加；Derived 受注殘約 2026 年底 8,700 億")]
    for i, (l, a, b_, n) in enumerate(rc, 3):
        RC.cell(i, 1, l); RC.cell(i, 2, a); RC.cell(i, 3, b_); RC.cell(i, 4, f"=B{i}-C{i}"); RC.cell(i, 5, n)
    i = 3 + len(rc) + 1
    RC.cell(i, 1, "反解（整本重算，build_daifuku.py 批次格點＋線性內插；常數為建置時結果）").font = BOLD; i += 1
    hdr(RC, i, ["反解目標", "鍵", "反解值", "基準值", "說明"]); i += 1
    for item in (rev or []):
        RC.cell(i, 1, item["label"]); RC.cell(i, 2, item["key"]); RC.cell(i, 3, item["val"]); RC.cell(i, 4, item["base"]); RC.cell(i, 5, item["note"]); i += 1
    RC.column_dimensions["A"].width = 60; RC.column_dimensions["E"].width = 70

    # ---------------- Sens ----------------
    SE = wb.create_sheet("Sens")
    SE["A1"] = "Sens — 一次動一個 In 頁參數（敏感度低／高），整本重算（build_daifuku.py 以 LibreOffice 批次重算每一案例）；結果為常數快照"
    base_tp = sens["base"]["tp"] if sens else None
    SE["A2"] = "快照基準值（加權，日圓）"; SE["B2"] = base_tp
    SE["C2"] = "本檔即時值（加權，日圓）"; SE["D2"] = "=tp"
    SE["E2"] = f"快照輸入指紋 {sens['fp'] if sens else ''}"
    SE["A3"] = "build_id"; SE["D3"] = bid
    SE["A4"] = "模型狀態"; SE["D4"] = '=IF(ABS(D2-B2)>0.5,"STALE","CURRENT")'
    hdr(SE, 6, ["驅動", "設定", "鍵", "加權目標價", "相對基準", "DCF 目標價", "DCF 相對基準", "估值評等"])
    sr = 7
    if sens:
        for cs in sens["cases"]:
            SE.cell(sr, 1, cs["drv"]); SE.cell(sr, 2, cs["set"]); SE.cell(sr, 3, cs["key"])
            SE.cell(sr, 4, cs["tp"]); SE.cell(sr, 5, cs["tp"] / base_tp - 1)
            SE.cell(sr, 6, cs["dcf"]); SE.cell(sr, 7, cs["dcf"] / sens["base"]["dcf"] - 1); SE.cell(sr, 8, cs["rt"])
            SE[f"E{sr}"].number_format = "0.0%"; SE[f"G{sr}"].number_format = "0.0%"; SE[f"D{sr}"].number_format = "#,##0"; SE[f"F{sr}"].number_format = "#,##0"
            sr += 1
    SE.cell(sr + 1, 1, "說明：相對倍數係數 rel_k 基準即為 1.00（未調整），故「相對倍數係數＝1」案例＝基準。pay、eq_26 等不影響目標價或無區間的鍵不列案例；線 5 市場 mk26／mk28／mk30／mk32 的 B 列不另設區間（線 5 低／高已作情境 C／A）。")
    SE.column_dimensions["A"].width = 46; SE.column_dimensions["B"].width = 34

    # ---------------- 對帳（線層級變數） ----------------
    RE = wb.create_sheet("對帳")
    RE["A1"] = "對帳 — 本模型使用的線層級變數（讀線 5 模型成品「下游取數」頁；與線模型不一致的自設值標明理由，不靜默改口徑）"
    hdr(RE, 2, ["變數", "所屬線", "本模型使用值（或區間）", "單位", "口徑", "出處（檔案與儲存格）", "標記", "影響本模型哪些儲存格", "與線模型一致？"])
    rec = [("AS/RS 市場 2026（低／基準／高）", "5", "10.45／10.69／10.93", "十億美元", "MarketsandMarkets 口徑（含製造、汽車）", L5 + " 下游取數!C8:E8", "Derived", "In mk26_C／mk26_B／mk26_A → PL 倉儲物流營收成長", "一致（照取）"),
           ("AS/RS 市場 2028（低／基準／高）", "5", "11.75／12.58／13.43", "十億美元", "同上", L5 + " 下游取數!C9:E9", "Derived", "In mk28_*", "一致"),
           ("AS/RS 市場 2030（低／基準／高）", "5", "13.2／14.8／16.5", "十億美元", "同上", L5 + " 下游取數!C10:E10", "Derived", "In mk30_*", "一致"),
           ("AS/RS 市場 2032（低／基準／高）", "5", "14.28／16.94／19.60", "十億美元", "同上；2031 起為假設成長", L5 + " 下游取數!C11:E11", "Derived", "In mk32_*", "一致"),
           ("倉儲物流營收成長 ÷ 市場成長（參與度）", "5", "0.2／0.5／1.0（自設）", "x", "大福商業與零售、運輸與倉儲營收 vs AS/RS 市場", "本模型 In cap_il", "Assumed", "PL c_co、c_tw", "自設：線模型沒有大福份額（下游取數無此列）；9M'26 大福該兩產業營收 −4.5%／−0.7%，低於市場成長，故取 0.5"),
           ("電子（半導體廠無塵室搬運）需求", "—", "未納入線 5", "—", "線 5 為倉儲物流單位經濟，不含半導體 AMHS", "地圖 positions.csv：實體 AI 曝險判為「中」", "—", "PL c_el（占 FY2026 營收約 41%）", "不一致（自設）：電子營收以半導體資本支出循環建模（ge27／ge28 等 Assumed），線模型無對應變數；結論主要靠此項"),
           ("AMR 市場、AMR／AS/RS 回本期、AutoStore 份額 7.0%", "5", "未使用", "—", "線模型單位經濟", L5 + " 下游取數!D4:D7、D12、D14、D27", "Derived", "未進公式", "未使用（大福 AS/RS 為堆垛機式、系統整合，線模型回本期為立方格口徑，不適用）")]
    for i, row in enumerate(rec, 3):
        for j, v in enumerate(row): RE.cell(i, 1 + j, v)
    RE.column_dimensions["A"].width = 40; RE.column_dimensions["C"].width = 30; RE.column_dimensions["F"].width = 60; RE.column_dimensions["I"].width = 50

    # ---------------- Risk ----------------
    RK = wb.create_sheet("Risk")
    RK["A1"] = "Risk — 風險登錄（未建模的尾部風險：風險、是否建模、理由、觀察指標）"
    hdr(RK, 2, ["風險", "是否建模", "理由", "觀察指標"])
    risks = [("半導體資本支出循環反轉（無塵室專案延後、取消）", "是（ge27_C、ge28_C）", "情境 C；電子占營收約 41%", "訂單 ÷ 營收、受注殘、Clean Factomation 與 Suzhou 訂單、SEMI 設備出貨"),
             ("倉儲物流營收低於市場成長（參與度）", "是（cap_il）", "9M'26 該兩產業營收負成長", "商業與零售、運輸與倉儲營收 YoY"),
             ("Eisenmann 併購整合（全年營收、獲利未取得）", "部分（aud27 區間）", "營收貢獻未揭露，利潤率未建模", "汽車分部營收與毛利率"),
             ("美國關稅與汽車客戶投資延後", "部分（gn27_C、gn28_C）", "公司說明汽車客戶因美國貿易政策延後決策", "北美營收、汽車訂單"),
             ("材料成本與供應限制", "是（gmadj_C）", "公司 2026-11 說明已接單案件可能受影響", "季度毛利率"),
             ("可轉債轉換（稀釋）", "否", "可轉債視為負債；轉換後債減、股增，每股約中性至微稀釋（9M 稀釋 EPS −4%）", "稀釋後 EPS、轉換公告"),
             ("2026 年底改採 IFRS（利益口徑變動）", "否", "公司參考值：IFRS 營業利益 1,170 億（J-GAAP 1,190 億）", "首份 IFRS 決算"),
             ("匯率（美元、韓元、人民幣）", "否", "公司 FY2026 假設 158 日圓兌美元；未逐年建模", "USDJPY")]
    for i, row in enumerate(risks, 3):
        for j, v in enumerate(row): RK.cell(i, 1 + j, v)
    RK.column_dimensions["A"].width = 50; RK.column_dimensions["C"].width = 50

    # ---------------- Sources ----------------
    SO = wb.create_sheet("Sources")
    for i, s_ in enumerate([TS3Q, TS2Q, LN25, SAF, SAS, KGS, ATS, YCL, JGB10, L5 + "（下游取數）",
                            "physical_ai/hiwin/dist/20261123_上銀個股模型_v0.1.1.xlsx（Comps 頁 Fanuc、安川換算 P/E）", "physical_ai/harmonic_drive/dist/20261123_哈默納科個股模型_v0.1.1.xlsx（Comp 頁 Nabtesco 換算 P/E）",
                            "physical_ai/map/profiles/6383.md（公司背景；數字未採用）"], 1):
        SO.cell(i, 1, s_)

    # ---------------- README（工作表；文字由公式組字） ----------------
    RD["A1"] = f"Daifuku 大福（6383.T）個股模型 {VER}｜build_id {bid}"
    RD["A2"] = '="資料截止日 2026-11-20（第 3 季決算 2026-11-11）；成品日 2026-11-24；股價基準日 "&TEXT(px_date,"yyyy-mm-dd")&"（收盤 "&TEXT(px,"#,##0")&" 日圓）；目標價時點 "&TEXT(tp_date,"yyyy-mm-dd")'
    RD["A3"] = "命題：大福是全球物料搬運系統龍頭，但營收約四成來自半導體廠無塵室搬運（電子），倉儲物流自動化（線 5）只占小部分且近期營收未跟上市場；股價由半導體資本支出循環帶動的訂單與受注殘轉營收、獲利決定"
    RD["A4"] = (f'="在 FY2027 電子營收情境 B 年增 "&TEXT(ge27_B,"0%")&"、邊際毛利率 "&TEXT(mm_gp,"0%")&"、同業前瞻 P/E 中位數 "&TEXT(pe_med,"0.0")&" 倍的基準假設下：加權目標價 "&TEXT(Val!B{r_tp},"#,##0")&" 日圓（DCF "&TEXT(Val!B{r_dcft},"#,##0")&"、Comparable "&TEXT(Val!B{r_cmpt},"#,##0")&"），上檔空間 "&TEXT(Val!B{r_up},"+0.0%;-0.0%")&"，估值評等 "&Val!B{r_rt}&"；收支關卡 "&Val!B{r_gj}&"（情境 B FY2026–30 自由現金流扣股利 "&TEXT(Val!C{r_gate},"#,##0.0")&" 十億日圓）；投資建議 "&Val!B{r_rec}')
    RD["A5"] = f'="距最近級距邊界 "&TEXT(Val!B{r_dist},"0.0")&" 個百分點；翻轉目標價（升一級／降一級）"&Val!B{r_bnd}&" 日圓。上檔三段歸因：DCF "&TEXT(Val!B{r_a1},"+0.0%;-0.0%")&"、未調整 Comparable "&TEXT(Val!B{r_a2},"+0.0%;-0.0%")&"、相對倍數調整 "&TEXT(Val!B{r_a3},"+0.0%;-0.0%")'
    RD["A6"] = "失效條件（可觀察）：訂單 ÷ 營收連兩季低於 1；電子（半導體）訂單 YoY 轉負；季度毛利率跌破 22%；FY2027 公司指引營收低於 FY2026 實績"
    RD["A7"] = "範圍：J-GAAP 合併；不讀 Siliconomics／Tokenomics／Wattonomics（算力線）；自用 AI 讀法等規格條文不適用。HTML 展示層見 dist/ 同名 .html"
    RD["A8"] = f'="命題回覆：問：半導體循環與倉儲自動化哪個決定估值？答：情境 B FY2030 電子占合併營收 "&TEXT(Val!B{r_hum},"0.0%")&"，倉儲物流（商業與零售、運輸與倉儲）合計約兩成（FY2026 約 21%）；加權目標價 "&TEXT(Val!B{r_tp},"#,##0")&" 日圓，情境 B FY2028 EPS "&TEXT(PL!H{PR["B"]["epsc"]},"#,##0")&" 日圓，現價為其 "&TEXT(px/PL!H{PR["B"]["epsc"]},"0.0")&" 倍，同業中位數 "&TEXT(pe_med,"0.0")&" 倍。倉儲物流線 5 的市場變數只影響約兩成營收，結論主要由電子循環與 P/E 決定（模型輸出，非獨立證據）。"'
    wb.save(path)
    return VR, PR, DR


# ---------------- 重算與讀值 ----------------
def lo_profile():
    d = os.path.join(os.environ.get("HOME", tempfile.gettempdir()), "lo_daifuku")
    os.makedirs(d, exist_ok=True)
    return "file://" + d

def recalc_many(srcs, outdir):
    os.makedirs(outdir, exist_ok=True)
    for i in range(0, len(srcs), 40):
        chunk = srcs[i:i + 40]
        for attempt in range(2):
            try:
                subprocess.run(["soffice", f"-env:UserInstallation={lo_profile()}", "--headless", "--convert-to", "xlsx", "--outdir", outdir] + chunk,
                               check=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, timeout=600)
                break
            except subprocess.TimeoutExpired:
                subprocess.run(["pkill", "-f", "soffice"]); continue
    return [os.path.join(outdir, os.path.basename(s)) for s in srcs]

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

def run_cases(ovs):
    """ovs：[(tag, ov)] → {tag: vals}；一次 LibreOffice 批次重算。"""
    d = tempfile.mkdtemp(prefix="dk_"); paths = []; VR = DR = None
    for tag, ov in ovs:
        p = os.path.join(d, f"c_{tag}.xlsx"); VR, _, DR = build(p, ov); paths.append(p)
    outs = recalc_many(paths, os.path.join(d, "o"))
    res = {tag: read_vals(o, VR, DR if tag == "base" else None) for (tag, _), o in zip(ovs, outs)}
    shutil.rmtree(d, ignore_errors=True); return res

def in_val(k, col):
    for row in ROWS:
        if row[0] == k:
            return row[{"lo": 3, "base": 4, "hi": 5}[col]]
    raise KeyError(k)

def fingerprint():
    return hashlib.sha1(json.dumps([r for r in ROWS], ensure_ascii=False, default=str).encode()).hexdigest()[:10]

def rating(u):
    return "強力買進" if u > 0.5 else "買進" if u > 0.25 else "中立" if u >= -0.25 else "放空" if u >= -0.5 else "強力放空"

CASES = [  # (鍵, 驅動白話)
    ("g_anc", "FY2026 實績相對公司指引"), ("ge27_B", "電子營收年增 FY2027（情境 B）"), ("ge28_B", "電子營收年增 FY2028（情境 B）"), ("gem_B", "電子營收年增 FY2029–30（情境 B）"),
    ("gelt_B", "電子營收年增 FY2031–32（情境 B）"), ("gn27_B", "汽車、機場、其他年增 FY2027（情境 B）"), ("gn28_B", "汽車、機場、其他年增 FY2028（情境 B）"),
    ("gnm_B", "汽車、機場、其他年增 FY2029–30（情境 B）"), ("gnlt_B", "汽車、機場、其他年增 FY2031–32（情境 B）"), ("cap_il", "倉儲物流參與度（相對線 5 市場成長）"),
    ("aud27", "汽車 FY2027 額外年增（Eisenmann 全年化）"), ("mm_gp", "邊際毛利率"), ("gmadj_B", "毛利率調整（情境 B）"), ("sga_el", "銷管費用彈性"),
    ("nonop", "業外與特別損益"), ("tax_w", "有效稅率"), ("fc_da", "折舊 ÷ 營收"), ("fc_cx", "資本支出 ÷ 營收"), ("fc_wc", "營運資金 ÷ 營收增量"), ("lease_v", "租賃負債代理"),
    ("beta", "β"), ("erp", "股權風險溢酬"), ("g_cap", "永續成長率上限"), ("ronic_sp", "RONIC − WACC"), ("wacc_adj", "WACC ±1pt"), ("w_dcf", "DCF 權重"), ("pe_med", "同業前瞻 P/E 中位數")]
BETA_B = 0.67 * 1.04 + 0.33
FXD = {"fc_da": (10.623 * 12 / 9) / 770.0, "fc_cx": (19.371 * 12 / 9) / 770.0, "pe_med": statistics.median([21.1499, 15.6739, 13.7701, 29.2 / (30.66 / 8.17), 14.79 / (15.55 / 24.04)])}


def solve_batch(key, lo, hi, targets, px0):
    """格點批次重算（13 點）→ 對每個目標上檔找相鄰兩點線性內插；再以 7 點細格點精修一次。"""
    import numpy as np
    xs = list(np.linspace(lo, hi, 13))
    res = run_cases([(f"g{i}", {key: x}) for i, x in enumerate(xs)])
    ys = [res[f"g{i}"]["up"] for i in range(len(xs))]
    out = {}
    for tgt in targets:
        if tgt is None: continue
        br = next((i for i in range(len(xs) - 1) if (ys[i] - tgt) * (ys[i + 1] - tgt) <= 0), None)
        if br is None: out[tgt] = None; continue
        a, b_ = xs[br], xs[br + 1]
        xs2 = list(np.linspace(a, b_, 7)); r2 = run_cases([(f"f{i}", {key: x}) for i, x in enumerate(xs2)]); y2 = [r2[f"f{i}"]["up"] for i in range(7)]
        j = next(i for i in range(6) if (y2[i] - tgt) * (y2[i + 1] - tgt) <= 0)
        x0, x1, y0, y1 = xs2[j], xs2[j + 1], y2[j], y2[j + 1]
        out[tgt] = x0 if y1 == y0 else x0 + (tgt - y0) * (x1 - x0) / (y1 - y0)
    return out


def main():
    px0 = in_val("px", "base")
    jobs = [("base", {})]; meta = []
    for k, drv in CASES:
        for side in ("lo", "hi"):
            v = in_val(k, side)
            if k == "beta": v = BETA_B + (-0.3 if side == "lo" else 0.3)
            if v is None: continue
            b_ = in_val(k, "base")
            if not isinstance(b_, (int, float)): b_ = FXD.get(k, BETA_B if k == "beta" else None)
            if b_ is not None and abs(v - b_) < 1e-12: continue
            tag = f"{k}_{side}"; jobs.append((tag, {k: v})); meta.append((tag, drv + ("（低）" if side == "lo" else "（高）"), f"{k}＝{v:.4g}", k))
    for s in SC:
        tag = f"only_{s}"; jobs.append((tag, {f"w_{x}": (1.0 if x == s else 0.0) for x in SC})); meta.append((tag, f"只看情境 {s}", f"w_{s}＝1，其餘 0", f"w_{s}"))
    res = run_cases(jobs)
    base = res["base"]; print("基準", base)
    cases = [dict(drv=d, set=st, key=k, **res[t]) for t, d, st, k in meta]
    up_ = base["alt_tp"] / px0 - 1
    cases.append(dict(drv="終值替代法（DCF 終值＝FCFF(FY2032)×(1＋g)÷(WACC − g)）", set="DCF 頁替代法列", key="tv_alt", tp=base["alt_tp"], dcf=base["alt_dcf"], cmp=base["cmp"], up=up_, rt=rating(up_)))
    for c in cases: print(c["key"], c["drv"], round(c["tp"]), c["rt"])
    up0 = base["up"]; tiers = [0.5, 0.25, -0.25, -0.5]
    upper = min([t for t in tiers if t > up0], default=None); lower = max([t for t in tiers if t <= up0], default=None)
    revs = []
    for key, b0, lo_, hi_, note in (("ge27_B", 0.20, -0.40, 0.90, "情境 B 的 FY2027 電子營收年增（其他不變，整本重算；情境 A／C 不動；之後年度在此基礎上成長）"),
                                    ("mm_gp", 0.25, -0.50, 1.00, "邊際毛利率（三情境同動）"),
                                    ("pe_med", FXD["pe_med"], 0.0, 40.0, "同業前瞻 P/E 中位數（Comparable；DCF 不受影響）")):
        sol = solve_batch(key, lo_, hi_, [0.0, upper, lower], px0)
        for lab, tgt in (("目標價＝現價（上檔 0%）", 0.0), (f"升一級邊界（上檔 {upper:+.0%}）" if upper is not None else None, upper), (f"降一級邊界（上檔 {lower:+.0%}）" if lower is not None else None, lower)):
            if lab is None: continue
            x = sol.get(tgt)
            revs.append(dict(label=lab, key=key, val=x if x is not None else f"區間 {lo_:+.2f}～{hi_:+.2f} 內無解", base=b0, note=note))
            print("反解", key, lab, x)
    sens = dict(base=base, cases=cases, fp=fingerprint())
    json.dump(sens, open(os.path.join(HERE, "sens.json"), "w"), ensure_ascii=False, indent=1)
    dist = os.path.join(HERE, "dist"); os.makedirs(dist, exist_ok=True)
    tmp = tempfile.mkdtemp(prefix="dkf_"); raw = os.path.join(tmp, NAME)
    bid = f"{VER}-{fingerprint()}"
    VR, PR, DR = build(raw, {}, sens, revs, bid)
    out = recalc_many([raw], os.path.join(tmp, "o"))[0]
    shutil.copy(out, os.path.join(dist, NAME))
    wbv = openpyxl.load_workbook(os.path.join(dist, NAME), data_only=True)
    errs = [(ws.title, c.coordinate, c.value) for ws in wbv for row in ws.iter_rows() for c in row
            if isinstance(c.value, str) and (c.value.startswith("#") or c.value.startswith("Err:"))]
    print("錯誤值格數", len(errs), errs[:10]); assert not errs, errs
    print("Sens 狀態", wbv["Sens"]["D4"].value, "tp", wbv["Val"][f"B{VR['tp']}"].value)
    write_readme(os.path.join(dist, NAME), VR, PR, DR, bid, sens, revs)
    json.dump(dict(company="daifuku", ticker="6383.T", model_sheet="Q",
                   official_segments=[dict(name=n, label=l, period="9M'26（累計，外部顧客）", value=v, unit="¥B", source=TS3Q + " 營收分解")
                                      for n, l, v in (("Daifuku", "Daifuku（日本本體）營收", 172.821), ("Contec", "Contec 營收", 17.913), ("Daifuku North America", "Daifuku North America 營收", 142.813),
                                                      ("Clean Factomation", "Clean Factomation（韓國）營收", 55.038), ("Daifuku Suzhou", "Daifuku (Suzhou) Cleanroom Automation 營收", 33.002),
                                                      ("其他", "其他（未列報告分部的子公司）營收", 131.460))],
                   sublines=[], check_labels=["檢查：分部合計 ＋ 合併調整 − 合併營收（四捨五入差 < 0.005）"],
                   note="官方報告分部為公司別（日本本體、Contec、北美、韓國 Clean Factomation、蘇州無塵室、其他），Q 頁以「分部合計 ＋ 合併調整 − 合併營收」對帳。PL 頁營收列改用公司揭露的客戶產業別營收（同一份決算短信），另有「檢查：產業別合計 − 營收」。"),
              open(os.path.join(HERE, "segments.json"), "w"), ensure_ascii=False, indent=1)


def write_readme(path, VR, PR, DR, bid, sens, revs):
    wb = openpyxl.load_workbook(path, data_only=True); V = wb["Val"]; P = wb["PL"]; RD = wb["README"]; I = wb["In"]
    g = lambda c: V[c].value
    f0 = lambda x: f"{x:,.0f}"; pc = lambda x: f"{x*100:+.1f}%"
    inv = lambda k: I['E' + str([r for r in range(1, I.max_row + 1) if I['A' + str(r)].value == k][0])].value
    tp, dcf, cmp_, up = g(f"B{VR['tp']}"), g(f"B{VR['dcf']}"), g(f"B{VR['cmp']}"), g(f"B{VR['up']}")
    big = sorted(sens["cases"], key=lambda c: -abs(c["tp"] / sens["base"]["tp"] - 1))[:8]
    rv = "\n".join(f"| {r['label']} | `{r['key']}` | {r['val'] if isinstance(r['val'], str) else format(r['val'], '.4f')} | {r['base']:.4g} | {r['note']} |" for r in revs)
    def yr(s, k, fmt="{:,.1f}", ys=range(2024, 2033)):
        out = []
        for y in ys:
            v = P[f"{COL[y]}{PR[s][k]}"].value
            out.append(fmt.format(v) if isinstance(v, (int, float)) else "—")
        return " | ".join(out)
    scn = "\n".join(f"| 情境 {s}（{SCN[s]}） | {f0(g(chr(66+i)+str(VR['dcfr'])))} | {f0(g(chr(66+i)+str(VR['cu'])))} | {f0(g(chr(66+i)+str(VR['tps'])))} | {g(chr(66+i)+str(VR['gate'])):,.1f} |" for i, s in enumerate(SC))
    rts = sorted(set(c['rt'] for c in sens['cases']))
    YS = range(2024, 2033)
    rows = "\n".join(f"| {n} | {yr('B', 'c_' + k)} |" for k, n in IND)
    txt = f"""# physical_ai/daifuku — Daifuku 大福（6383.T）個股模型 {VER}

{RD['A2'].value}｜build_id `{bid}`｜成品 `dist/{NAME}`（LibreOffice 重算，錯誤值 0）

東證 Prime 上市。本 README 的數字由 `build_daifuku.py` 從重算後的成品讀出產生，不手寫。

## 完成狀態與未完成項目
- 已完成：個股模型 v0.1（In、Hist、Q、PL、DCF、Val、Recon、Sens、對帳、Risk、Sources）與 HTML 展示層。
- 未完成或未取得：(1) FY2024 細項（只有由公司年增率回推的營收、營業利益參考值）、FY2025 折舊與資本支出全年值；(2) Eisenmann 併購（2026-07 納入）的全年營收與利潤率；(3) 公司不揭露訂單產業別，無塵室占營收只能看電子產業別；(4) 同業：SMC 報價日 2026-10-09 在收盤檔之外，列「報價日未取得」未納入；半導體搬運同業（Murata Machinery 等）無上市報價；(5) 公司第 3 季簡報（受注殘、各分部訂單明細）只讀決算短信，未讀簡報。

## 命題
{RD['A3'].value}

**結論（條件式）：** {RD['A4'].value}

{RD['A5'].value}

{RD['A8'].value}

**讀法：** 現價只有模型 FY2027 EPS 的 {g(f"B{VR['pe1']}"):.1f} 倍；賣方平均目標價 {f0(inv('cons_tp'))}（15 位，第 3 季決算後）。Comparable 用 5 家同業前瞻 P/E 中位數 {inv('pe_med'):.1f} 倍（皆為 2026-11-20 報價）。結論主要建立在 Assumed 的電子營收成長（FY2027、FY2028）、邊際毛利率與同業 P/E 上；線 5 倉儲物流市場變數（Derived）只影響約兩成營收，不是承重項。

## 估值（評價日 2026-12-31；目標時點約 12 個月）
| 情境 | DCF（推到 12 個月後） | Comparable（P/E × FY2028 EPS） | 情境目標價 | 收支關卡（FY2026–30 FCF 扣股利，¥B） |
|---|---|---|---|---|
{scn}

- 加權目標價 **{f0(tp)} 日圓**（DCF {f0(dcf)}、Comparable {f0(cmp_)}），現價 {f0(inv('px'))}（2026-11-20 收盤），上檔 {pc(up)}，估值評等 **{g(f"B{VR['rt']}")}**；收支關卡 {g(f"B{VR['gj']}")} → 投資建議 **{g(f"B{VR['rec']}")}**。賣方平均目標價 {f0(inv('cons_tp'))}，只作對照。
- 方法：DCF 標準版 FCFF（{g(f"B{VR['wk']}")}，為 WACC／ke／g／RONIC）50%＋Comparable 50%（同業 5 家前瞻 P/E 中位數 {inv('pe_med'):.2f} 倍 × 各情境 FY2028 EPS）。相對倍數係數 1.00（未調整）。
- 隱含前瞻 P/E（加權目標價 ÷ 情境 B FY2028 EPS）{g(f"B{VR['pe']}"):.1f} 倍；大福現行前瞻 P/E {inv('pe_self'):.2f} 倍 ＝ 同業中位數的 {g(f"B{VR['pe2']}"):.2f} 倍。

## 營收與標準損益（情境 B，¥B；FY2025 實績、FY2026 公司指引、FY2027 起模型）
| 列 | {" | ".join(f"FY{y}" for y in YS)} |
|---|{"---|" * len(YS)}
{rows}
| 營收（合併） | {yr('B','rev')} |
| 毛利 | {yr('B','gp')} |
| 營業利益 | {yr('B','op')} |
| 歸屬母公司淨利 | {yr('B','ni')} |
| EPS（日圓） | {yr('B','epsc','{:,.0f}')} |

- 決算期：2025 年起由 3 月決算改 12 月決算；FY2023 欄＝2024 年 3 月期（舊決算期，不可比）；FY2024 欄只有由公司揭露的 FY2025 年增率（營收 +2.6%、營業利益 +24.4%）回推的參考值。資料商頁面另有一列標為「Dec 2024」營收 737.3、營業利益 96.5，與公司年增率回推值不一致，期間定義未查清，**未採用**。
- 營收依客戶產業別（公司決算短信的營收分解表）；FY2025 產業別取自 LNEWS 報導（億日圓取整），「其他產業」為餘項。FY2026 產業結構＝9M'26 實績結構 × 指引營收 770（Q4 結構同 9M 為假設）。

## 主要假設（In 頁；Assumed／Analogy 皆給區間）
- 電子營收年增（A／B／C）：FY2027 32%／20%／5%；FY2028 12%／8%／−12%；FY2029–30 6%／4%／0%；FY2031–32 4%／3%／2%。依據：公司 FY2026 訂單指引 970–1,000 億（營收 770），9M 訂單 781（+56.5%），中間期受注殘超過 7,400 億。
- 倉儲物流（商業與零售、運輸與倉儲）：線 5 AS/RS 市場成長（Derived，下游取數）× 參與度 0.5（Assumed，區間 0.2–1.0；9M'26 這兩個產業營收為負成長）。
- 汽車、機場、其他：年增 4%（A 6%、C 0% → −3%）；汽車 FY2027 另加 Eisenmann 全年化 8pt（區間 0–15pt）。
- 毛利：FY2026 毛利（9M 毛利率 25.1% × 指引營收）＋ 邊際毛利率 25%（區間 20–30%）× Δ營收 ＋ 情境調整；銷管費用彈性 0.7；稅率 28%。
- 現金流：折舊、資本支出為 9M 年化比率（1.8%、3.4%）；營運資金 15% × 營收增量；股利＝配息率 38.5% × 淨利。
- 折現：rf 2.817%（逾期）、β {inv('beta'):.3f}（單一來源）、ERP 5%、CRP 0.91%、g＝MIN(3%, rf)、RONIC＝WACC。評價日股數 372.2 百萬股；淨現金約 {wb['DCF']['B10'].value:,.1f}（現金 256.6 − 借款與可轉債 46.8 − 租賃代理）。可轉債視為負債，未計稀釋。

## 承重 Assumed 驅動與可證偽指標
| 驅動 | 公開可觀察指標 |
|---|---|
| FY2027–28 電子營收成長（ge27、ge28） | 訂單 ÷ 營收（9M 1.41）、受注殘、Clean Factomation／Suzhou 訂單；FY2027 公司指引（2027-02） |
| 邊際毛利率（mm_gp） | 季度毛利率（9M'26 25.1%）、Δ毛利 ÷ Δ營收 |
| 同業 P/E（pe_med） | 同業前瞻 P/E 換算到同一日期後的中位數（目前 5 家） |

失效條件：{RD['A6'].value.split('：', 1)[1]}。

## 敏感度（整本重算；Sens 頁；影響最大的 8 項）
| 驅動 | 設定 | 鍵 | 加權目標價 | 相對基準 | 估值評等 |
|---|---|---|---|---|---|
""" + "\n".join(f"| {c['drv']} | {c['set']} | `{c['key']}` | {f0(c['tp'])} | {pc(c['tp']/sens['base']['tp']-1)} | {c['rt']} |" for c in big) + f"""

全部 {len(sens['cases'])} 案見 xlsx Sens 頁；各案例估值評等出現「{'、'.join(rts)}」。

## 反解（整本重算；Recon 頁）
| 反解目標 | 鍵 | 反解值 | 基準 | 說明 |
|---|---|---|---|---|
{rv}

## 對帳（線層級變數；xlsx「對帳」頁）
- 線 5：AS/RS 市場 2026／2028／2030／2032 低／基準／高（下游取數 C8:E11）→ mk26～mk32 的情境 C／B／A，照取。
- 自設（與線模型不一致）：倉儲物流參與度 0.5（線模型無大福份額）；電子（半導體無塵室）占營收約 41%，線 5 未涵蓋，以半導體資本支出循環自設成長（Assumed）；線模型的 AMR、立方格回本期與 AutoStore 份額不適用大福，未使用。

## 限制與未取得資料
- 公司訂單產業別、半導體無塵室營收未單獨揭露；FY2024 細項、FY2025 折舊與資本支出全年值未取得；Eisenmann 全年營收與利潤率未取得。
- 同業前瞻 P/E 僅 5 家（Fanuc、安川、Nabtesco 取自同專案其他模型的 2026-11-20 換算值；KION、AutoStore 以資料商 2026-11-23 前瞻 P/E 推隱含 EPS 後以 2026-11-20 收盤換算），區間寬；SMC 報價日未取得。賣方共識 EPS 年度基準不一致。
- rf、CRP 沿用 2026-11-02／01-05 取值（逾期，跨模型一致）；β 單一來源。DCF 為本模型自寫，公式與 `docs/plan/個股DCF標準規格.md` 相同。
- 決算日（2026-11-11）至 2026-11-20 股價 −18.4%，與全球 AI 與半導體股同週大跌；個別原因未查核。

## 檔案
- `build_daifuku.py`：建置＋整本重算 Sens／反解（LibreOffice 批次）＋產生本 README、`segments.json`、`sens.json`。
- `dist/{NAME}`：成品（工作表 README、In、Hist、Q、PL、DCF、Val、Recon、Sens、對帳、Risk、Sources）。
- `html.json`＋`gen_html_json.py`：HTML 展示層設定（tools/company_html，layout v1.3）；`dist/{NAME[:-5]}.html`。

## 版本紀錄
- v0.1（2026-11-24）：初版；現價 {f0(inv('px'))}（2026-11-20 收盤）；加權目標價 {f0(tp)}、上檔 {pc(up)}、估值評等 {g(f"B{VR['rt']}")}、投資建議 {g(f"B{VR['rec']}")}。
"""
    open(os.path.join(HERE, "README.md"), "w").write(txt)


if __name__ == "__main__":
    main()
