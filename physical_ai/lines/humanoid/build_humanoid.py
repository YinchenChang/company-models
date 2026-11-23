# -*- coding: utf-8 -*-
"""線 2b 人形機器人單位經濟 v0.1 產生器。

python3 build_humanoid.py  ->  dist/20261116_人形機器人單位經濟_v0.1.xlsx
所有計算都寫成 Excel 公式（可追溯）；本程式只放輸入值、出處與版面。
"""
import os
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.utils import get_column_letter as CL

VERSION = "v0.1"
DATE = "20261116"
OUT = os.path.join(os.path.dirname(os.path.abspath(__file__)), "dist",
                   f"{DATE}_人形機器人單位經濟_{VERSION}.xlsx")

BLUE = Font(color="0000FF")
BLUE_B = Font(color="0000FF", bold=True)
BOLD = Font(bold=True)
TITLE = Font(bold=True, size=14)
HDR = Font(bold=True, color="FFFFFF")
HDR_FILL = PatternFill("solid", fgColor="1F6F78")
YEL = PatternFill("solid", fgColor="FFF2CC")
GRY = PatternFill("solid", fgColor="EDEDED")
GRN = PatternFill("solid", fgColor="E2F0D9")
WRAP = Alignment(wrap_text=True, vertical="top")
THIN = Side(style="thin", color="BFBFBF")

# ---------------------------------------------------------------- 出處
S = {
    "PICEA": "humanoid.guide 關節 BOM（引 PICEA Motion、X.PIN、MS via GrabaRobot），2026：https://humanoid.guide/?p=17836",
    "MS": "Morgan Stanley Optimus Gen 2 BOM（二手轉述 GrabaRobot，2026-07）：https://grabarobot.com/blog/humanoid-robot-bill-of-materials-cost-2026/",
    "XENO": "xenospectrum 彙整 MS／BofA Institute（2025-04-29）／Goldman：https://xenospectrum.com/en/humanoid-robot-bom-actuator-cost/",
    "SCREW": "humanoid.guide 行星滾柱絲杠（引 China Daily 2026-01、36Kr）：https://humanoid.guide/?p=20588",
    "CD": "China Daily 寧波 雙林 ¥499 絲杠量產（2026-01-27）：https://ningbo.chinadaily.com.cn/2026-01/27/c_1157428.htm",
    "MITU": "覓途諮詢零件占比（經鉅亨號）：https://hao.cnyes.com/post/252957",
    "BLS": "BLS ECEC 私部門全負擔時薪（2026-06，2026-09 發布）：https://www.bls.gov/regions/southwest/news-release/2026/employercostsforemployeecompensation_regions_20260909.htm",
    "UNI": "宇樹招股書經 KrASIA：https://kr-asia.com/what-is-the-secret-behind-unitrees-high-gross-margins",
    "UBT": "優必選 2026H1 經 KrASIA：https://kr-asia.com/ubtechs-full-size-humanoid-robot-revenue-jumps-1445-in-h1-2026",
    "TSLA": "Tesla 2026Q3 法說回顧：https://notateslaapp.com/news/4867/tesla-q3-2026-earnings-call-recap-everything-you-need-to-know",
    "MUSK": "Musk 成本與租賃說法：https://tech-insider.org/tesla-optimus-rent-not-sell-musk-30000-build-cost-2026/",
    "MSQ": "Morgan Stanley 中國出貨預測（TNW，2026-06-24）：https://thenextweb.com/news/china-humanoid-robot-forecast-morgan-stanley-50000",
    "GS": "Goldman 出貨預測上修（humanoid.guide，2026-09-15）：https://humanoid.guide/?p=18115",
    "TF": "大公資信報告（TrendForce 2026 全球 >5 萬台等）：https://www.jtcopper.com/wp-content/uploads/2026/05/人形机器人：量产突破期的产业分化与国有资本投资逻辑.pdf",
    "OMDIA": "Omdia 2025 出貨（TechNode，2026-01-09）：https://technode.com/2026/01/09/chinas-agibot-leads-global-humanoid-robot-shipments-in-2025-omdia-says/",
    "IDC": "IDC 2025 出貨約 1.8 萬台（界面）：https://en.jiemian.com/article/13925066.html",
    "NEXT": "各廠宣布 vs 實際出貨（2026H1 追蹤）：https://www.nextmsc.com/industry-deep-dive/humanoid-robots-announced-vs-shipped-units",
    "FW": "本線框架 線2b_人形機器人_框架_v0.1.md（2026-11-16）",
    "UA": "優分析 諧波減速機（2025-03-10）：https://uanalyze.com.tw/articles/4453411375",
}

wb = Workbook()
ws0 = wb.active
ws0.title = "說明"
wsI = wb.create_sheet("輸入")
wsS = wb.create_sheet("敏感度")
wsC = wb.create_sheet("計算")
wsN = wb.create_sheet("情境")
wsD = wb.create_sheet("下游取數")
wsT = wb.create_sheet("追蹤指標")
wsK = wb.create_sheet("核對")


def hdr(ws, row, labels, col=1):
    for i, t in enumerate(labels):
        c = ws.cell(row, col + i, t)
        c.font = HDR
        c.fill = HDR_FILL
        c.alignment = WRAP


def widths(ws, ws_w):
    for i, w in enumerate(ws_w, 1):
        ws.column_dimensions[CL(i)].width = w


# =============================================================== 輸入
IN = {}  # key -> row
widths(wsI, [12, 30, 6, 12, 11, 11, 11, 11, 12, 18, 60, 11, 30])
wsI["A1"] = "輸入｜全尺寸人形機器人（Optimus Gen 2 類，約 30 關節＋2 靈巧手）2026 年基準"
wsI["A1"].font = TITLE
wsI["A2"] = ("藍字＝輸入；黃底＝判斷值（Assumed／Analogy，或可改的選擇）。使用值＝覆寫欄有值時用覆寫，否則用基準。"
             "注意：低／高是數值低／高，不是看多／看空。金額一律美元（人民幣輸入以匯率換算）。")
wsI["A2"].alignment = WRAP
wsI.merge_cells("A2:M2")
wsI.row_dimensions[2].height = 32
hdr(wsI, 4, ["代號", "變數", "類別", "單位", "低", "基準", "高", "覆寫", "使用值", "標記",
             "出處／說明", "日期", "在式中的位置"])

r = 5


def sec(title):
    global r
    c = wsI.cell(r, 1, title)
    c.font = BOLD
    c.fill = GRY
    for j in range(2, 14):
        wsI.cell(r, j).fill = GRY
    r += 1


def inp(key, name, cat, unit, lo, base, hi, tag, src, date, where, fmt="#,##0", use=None):
    """use: 自訂使用值公式（覆寫欄仍優先）。"""
    global r
    IN[key] = r
    vals = [key, name, cat, unit, lo, base, hi, None, None, tag, src, date, where]
    for j, v in enumerate(vals, 1):
        c = wsI.cell(r, j, v)
        c.alignment = WRAP if j in (2, 11, 13) else Alignment(vertical="top")
    judg = tag.startswith("[Assumed]") or tag.startswith("[Analogy]")
    for j in (5, 6, 7):
        c = wsI.cell(r, j)
        if isinstance(c.value, str) and c.value.startswith("="):
            c.font = Font(color="000000")
        else:
            c.font = BLUE
        c.number_format = fmt
        if judg:
            c.fill = YEL
    ov = wsI.cell(r, 8)
    ov.fill = YEL
    ov.font = BLUE
    ov.number_format = fmt
    basef = use if use else f"F{r}"
    u = wsI.cell(r, 9, f'=IF(H{r}="",{basef},H{r})')
    u.number_format = fmt
    u.font = BOLD
    r += 1


def U(k):
    return f"'輸入'!$I${IN[k]}"


def L(k):
    return f"'輸入'!$E${IN[k]}"


def H(k):
    return f"'輸入'!$G${IN[k]}"


def B(k):
    return f"'輸入'!$F${IN[k]}"


sec("A. 匯率與結構數量")
inp("fx", "人民幣兌美元匯率", "[A]", "CNY/USD", 6.9, 7.1, 7.3, "[Assumed]", "換算用；未逐日查證，給區間", "2026-11", "人民幣輸入 ÷ 匯率", "0.00")
inp("n_rot", "旋轉關節數（每台）", "[P]", "個", 14, 16, 17, "[Interested-party]",
    "BofA Institute：30 致動器＝16 旋轉＋14 線性（" + S["XENO"] + "）；humanoid.guide 表列 12–17", "2025-04／2026", "BOM：旋轉關節數 × 旋轉關節單價")
inp("n_lin", "線性關節數（每台）", "[P]", "個", 12, 14, 14, "[Interested-party]",
    "Optimus 約 14 支倒置行星滾柱絲杠（36Kr，經 " + S["SCREW"] + "）；MS 約 12", "2026", "BOM：線性關節數 × 線性關節單價")
inp("n_hand", "靈巧手數（每台）", "[P]", "隻", 2, 2, 2, "[Verified]", "全尺寸人形雙手", "—", "BOM：2 × 靈巧手")
inp("n_ft", "六維力感測器（每台）", "[P]", "顆", 2, 2, 4, "[Interested-party]", "通常在手腕與腳踝 2–4 顆（X.PIN，經 " + S["PICEA"] + "）", "2026", "BOM：感測")
inp("n_enc_rot", "編碼器／旋轉關節", "[P]", "顆", 1, 2, 2, "[Interested-party]", "馬達端＋輸出端雙編碼器（產業網站圖示，" + S["PICEA"] + "）", "2026", "旋轉關節單價")
inp("n_enc_lin", "編碼器／線性關節", "[P]", "顆", 1, 1, 2, "[Assumed]", "未取得；框架設 1 顆", "—", "線性關節單價")

sec("B. 關節零組件單價（美元／顆）")
inp("p_harm", "諧波減速機", "[P]", "USD", 400, 800, 1500, "[Interested-party]",
    "高階 800–1,500（PICEA Motion，" + S["PICEA"] + "）；低＝中國廠便宜約 50%（同出處稱便宜 30–50%）", "2026", "旋轉關節單價")
inp("p_motor", "無框力矩馬達", "[P]", "USD", 150, 250, 450, "[Assumed]",
    "單價未取得；框架以「占零件價值約 16%」× BOM 5–6 萬推得 290–340（" + S["MITU"] + "）", "—", "旋轉＋線性關節單價")
inp("p_enc", "編碼器", "[P]", "USD", 15, 40, 100, "[Assumed]", "單價未取得", "—", "旋轉＋線性關節單價")
inp("p_drv", "驅動板（每關節）", "[P]", "USD", 60, 100, 200, "[Assumed]", "單價未取得；X.PIN 稱控制器約占零件成本 12%（核對頁比較）", "—", "旋轉＋線性關節單價")
inp("p_brg", "軸承（每關節合計）", "[P]", "USD", 30, 60, 120, "[Assumed]", "單價未取得；覓途稱軸承約占 5.5%（核對頁比較）", "—", "旋轉＋線性關節單價")
inp("p_tsens", "關節扭矩感測器（每旋轉關節）", "[P]", "USD", 0, 50, 300, "[Assumed]", "是否每關節配置未取得；低＝不配", "—", "旋轉關節單價")
inp("p_hous", "關節外殼（每關節）", "[P]", "USD", 30, 50, 100, "[Assumed]", "未取得", "—", "旋轉＋線性關節單價")
r_screw_row = r
inp("p_screw", "行星滾柱絲杠（使用值＝下方情形選擇）", "[P]", "USD", None, None, None, "[Interested-party]",
    "五個出處互相矛盾，做成情形選擇與敏感度；低＝情形 2、基準＝情形 3、高＝情形 5", "2025–2026", "線性關節單價", use="0")
inp("p_hand", "靈巧手（每隻，含空心杯馬達、觸覺）", "[P]", "USD", 2100, 4750, 6500, "[Interested-party]",
    "基準＝MS 兩隻合計 9,500（" + S["MS"] + "）；低＝中國量產目標單隻 ¥15,000（[Analogy]，" + S["MITU"] + "）；高＝[Assumed] 22 自由度手", "2026", "BOM：2 × 靈巧手")
inp("p_ft", "六維力感測器（每顆）", "[P]", "USD", 700, 1000, 4200, "[Interested-party]", "¥5,000–30,000（約 700–4,200 美元；X.PIN，" + S["PICEA"] + "）；基準為判斷", "2026", "BOM：感測")
inp("p_head", "頭部（相機等）", "[P]", "USD", 1500, 2100, 3000, "[Interested-party]", "基準 MS 2,100（" + S["MS"] + "）；低／高 [Assumed]", "2026", "BOM")
inp("p_comp", "主控運算", "[P]", "USD", 3000, 4000, 6000, "[Interested-party]", "Goldman 3,000–6,000+（" + S["XENO"] + "）；晶片細節只讀 AI供應鏈", "2026", "BOM")
inp("bat_kwh", "電池容量", "[P]", "kWh", 2.0, 2.3, 3.0, "[Verified]", "Optimus 2.3 kWh（Tesla 2022 AI Day）；低／高 [Assumed]", "2022", "BOM：電池＝kWh × 每 kWh", "0.0")
inp("p_kwh", "電池包每 kWh", "[P]", "USD/kWh", 100, 130, 200, "[Derived]", "MS 電池約 300 美元 ÷ 2.3 kWh ≈ 130（" + S["MS"] + "）；低／高 [Assumed]", "2026", "BOM：電池")
inp("p_struct", "結構件", "[P]", "USD", 2000, 3000, 5000, "[Assumed]", "未取得絕對值；X.PIN 稱約占零件成本 10%（核對頁比較）", "—", "BOM")
inp("p_harn", "線束與其他電子", "[P]", "USD", 400, 800, 1500, "[Assumed]", "未取得", "—", "BOM")

sec("C. 製造成本")
inp("yld", "整機良率", "[P]", "%", 0.70, 0.85, 0.95, "[Assumed]", "所有廠商都未揭露", "—", "製造成本＝BOM ÷ 良率 ＋ …", "0%")
inp("labor", "組裝測試人工（每台）", "[A]", "USD", 1000, 2500, 5000, "[Assumed]", "未取得", "—", "製造成本", )
inp("fixcap", "年固定成本／每單位年產能", "[A]", "USD/台·年", 1000, 2500, 5000, "[Assumed]",
    "折舊＋廠務＋研發攤提；產能錨點 [Interested-party]：Tesla Fremont 設計 100 萬台/年（" + S["TSLA"] + "）、Figure BotQ 1.2 萬、宇樹 7.5 萬、優必選 2026 年 5,000", "2026", "每台固定成本＝本值 ÷（工廠稼動 × 良率）")
inp("futil", "工廠產能利用率", "[P]", "%", 0.30, 0.60, 0.85, "[Assumed]", "量產初期產能遠高於出貨", "—", "固定成本分攤", "0%")
inp("gm", "整機毛利率（定價用）", "[A]", "%", 0.15, 0.25, 0.45, "[Analogy]",
    "低＝車廠 15–20%；高＝優必選 44.7%（" + S["UBT"] + "）；宇樹 62.9% 為小型機科研市場，不作上限", "2025–2026", "售價＝製造成本 ÷（1 − 毛利率）", "0%")

sec("D. 使用端 TCO")
inp("deploy", "部署整合費（占售價）", "[A]", "%", 0.10, 0.30, 0.50, "[Analogy]", "類比工業機器人系統整合費比例", "—", "初始投入＝售價 ×（1 ＋ 本值）", "0%")
inp("maint", "年維護費（占售價）", "[A]", "%/年", 0.05, 0.10, 0.15, "[Assumed]", "無公開實測", "—", "年營運成本", "0%")
inp("life", "使用壽命", "[P]", "年", 3, 5, 7, "[Assumed]", "Tesla 稱壽命與可靠度約 2027 年底解決（[Interested-party]，" + S["TSLA"] + "）；無公開 MTBF", "—", "CRF", "0")
inp("rate", "資金成本", "[A]", "%", 0.05, 0.08, 0.12, "[Assumed]", "—", "—", "CRF＝r ÷ [1 −（1＋r）^−n]", "0%")
inp("util", "運轉時數占全年（1／2／3 班）", "[P]", "%", 0.23, 0.46, 0.68, "[Assumed]", "約 2,000／4,000／6,000 小時 ÷ 8,760；Tesla 目標全天候（[Interested-party]）", "—", "有效工時", "0%")
inp("down", "停機率", "[P]", "%", 0.05, 0.15, 0.30, "[Assumed]", "無公開實測；含充電／換電", "—", "有效工時", "0%")
inp("prod", "相對人工生產力（2026）", "[P]", "倍", 0.15, 0.30, 0.50, "[Assumed]", "無公開實測，未取得；承重變數", "—", "有效工時 ＝ 8,760 × 運轉 ×（1 − 停機）× 本值", "0.00")
inp("pkw", "平均耗電", "[P]", "kW", 0.3, 0.5, 1.0, "[Assumed]", "2.3 kWh 電池可用數小時推估", "—", "年電費", "0.0")
inp("elec", "電價", "[A]", "USD/kWh", 0.06, 0.09, 0.15, "[Assumed]", "未逐一查證", "—", "年電費", "0.00")
inp("soft", "年軟體授權", "[A]", "USD/年", 0, 3000, 12000, "[Assumed]", "未取得；機器人智能暫緩，只放成本", "—", "年營運成本")
inp("wage_us", "美國工人全負擔時薪", "[A]", "USD/小時", 46.89, 46.89, 46.89, "[Verified]", "私部門全體 2026-06（" + S["BLS"] + "）", "2026-06", "替代人工價值", "0.00")
inp("cn_wage", "中國工廠工人全負擔年成本", "[A]", "CNY/年", 80000, 100000, 120000, "[Assumed]", "待補國家統計局", "—", "替代人工價值（中國）")
inp("whours", "工人年工時", "[A]", "小時", 1900, 2000, 2400, "[Assumed]", "—", "—", "中國時薪換算；等效 FTE")
inp("wsel", "工人所在地（1＝美國，2＝中國）", "—", "選擇", 1, 1, 2, "[Assumed]", "選擇值，不是不確定性；敏感度頁另列", "—", "時薪選擇", "0")
inp("thr", "商業化門檻：回本期上限", "[A]", "年", 2, 3, 4, "[Assumed]", "框架第 4 節門檻 2", "—", "跨門檻判斷", "0.0")

sec("E. 學習曲線基期")
inp("q0", "2025 年底全球累積出貨", "[P]", "台", 15000, 20000, 30000, "[Assumed]",
    "2025 全年約 1.3 萬（Omdia，" + S["OMDIA"] + "）至 1.8 萬（IDC，" + S["IDC"] + "）＋更早年份", "2026-01", "學習曲線 Q₀")

sec("F. 對照與交叉檢核用（不進主計算）")
inp("ms_bom", "MS Optimus Gen 2 BOM（不含軟體）", "[A]", "USD", 50000, 55000, 60000, "[Interested-party]", "Morgan Stanley（賣方）二手轉述：" + S["MS"], "2026-07", "核對")
inp("ms_act", "MS 致動器占 BOM", "[A]", "%", 0.56, 0.56, 0.56, "[Interested-party]", S["XENO"], "2026", "核對", "0%")
inp("mitu_screw", "覓途：絲杠占零件價值", "[A]", "%", 0.19, 0.19, 0.19, "[Interested-party]", S["MITU"], "—", "核對（與 MS 單價矛盾）", "0.0%")
inp("mitu_motor", "覓途：無框馬達占零件價值", "[A]", "%", 0.16, 0.16, 0.16, "[Interested-party]", S["MITU"], "—", "核對", "0.0%")
inp("mitu_brg", "覓途：軸承占零件價值", "[A]", "%", 0.055, 0.055, 0.055, "[Interested-party]", S["MITU"], "—", "核對", "0.0%")
inp("xpin_ctrl", "X.PIN：控制器占零件成本", "[A]", "%", 0.12, 0.12, 0.12, "[Interested-party]", S["PICEA"], "2026", "核對", "0.0%")
inp("xpin_str", "X.PIN：結構件占零件成本", "[A]", "%", 0.10, 0.10, 0.10, "[Interested-party]", S["PICEA"], "2026", "核對", "0.0%")
inp("bofa_25", "BofA：2025 年底業界 BOM", "[A]", "USD", 35000, 35000, 35000, "[Interested-party]", S["XENO"], "2025-04-29", "核對")
inp("bofa_30", "BofA：2030–35 年 BOM 目標", "[A]", "USD", 13000, 15000, 17000, "[Interested-party]", S["XENO"], "2025-04-29", "核對（情境頁 BOM 路徑）")
inp("musk", "Musk：Optimus 成本說法", "[I]", "USD", 20000, 25000, 30000, "[Interested-party]", "初期約 3 萬往 2 萬（Musk，股東利害相關）：" + S["MUSK"], "2026", "核對")
inp("uni_asp", "宇樹人形 ASP（2025 前三季）", "[A]", "CNY", 167600, 167600, 167600, "[Verified]", "招股書：" + S["UNI"], "2025-09", "核對：中國小型機隱含單台成本")
inp("uni_gm", "宇樹人形毛利率（2025 前三季）", "[A]", "%", 0.629, 0.629, 0.629, "[Verified]", "招股書：" + S["UNI"], "2025-09", "核對", "0.0%")
inp("ubt_rev", "優必選全尺寸人形營收 2026H1", "[A]", "CNY", 590300000, 590300000, 590300000, "[Verified]", "中報：" + S["UBT"], "2026-06", "核對")
inp("ubt_units", "優必選全尺寸人形交付 2026H1", "[A]", "台", 921, 921, 921, "[Verified]", "中報：" + S["UBT"], "2026-06", "核對")

# 絲杠情形表
r += 1
wsI.cell(r, 1, "G. 行星滾柱絲杠單價情形（五個出處互相矛盾；選擇值決定使用值）").font = BOLD
r += 1
hdr(wsI, r, ["情形", "說明", "", "原幣", "原幣單價", "美元單價", "", "", "", "標記", "出處", "日期", ""])
r += 1
SCREW_CASES = [
    (1, "中國低價量產（雙林 ¥499）", "CNY", 499, "[Interested-party]", S["CD"], "2026-01"),
    (2, "中國主流均價（五洲新春 2025 均價約 ¥1,360，年降 39%，約 8.7 萬支）", "CNY", 1360, "[Interested-party]", S["SCREW"], "2025"),
    (3, "國際品（>¥5,000，China Daily 引業界估計）", "CNY", 5000, "[Interested-party]", S["CD"], "2026-01"),
    (4, "MS Optimus 級（低端 1,350 美元）", "USD", 1350, "[Interested-party]", S["MS"], "2026-07"),
    (5, "MS Optimus 級（高端 2,700 美元）", "USD", 2700, "[Interested-party]", S["MS"], "2026-07"),
]
case_first = r
for n, desc, ccy, px, tag, src, d in SCREW_CASES:
    wsI.cell(r, 1, n)
    wsI.cell(r, 2, desc).alignment = WRAP
    wsI.cell(r, 4, ccy)
    c = wsI.cell(r, 5, px)
    c.font = BLUE
    c.number_format = "#,##0"
    f = f"=E{r}/{U('fx')}" if ccy == "CNY" else f"=E{r}"
    c = wsI.cell(r, 6, f)
    c.number_format = "#,##0"
    wsI.cell(r, 10, tag if ccy == "USD" else "[Derived]（原幣 " + tag + "）")
    wsI.cell(r, 11, src).alignment = WRAP
    wsI.cell(r, 12, d)
    r += 1
case_last = r - 1
r += 1
wsI.cell(r, 2, "絲杠情形選擇（1–5）").font = BOLD
c = wsI.cell(r, 6, 3)
c.font = BLUE_B
c.fill = YEL
SCREW_SEL = f"'輸入'!$F${r}"
wsI.cell(r, 11, "基準選 3（國際品）：讓自下而上 BOM 與 MS 標題 5–6 萬美元大致對得上；選 4／5 時 BOM 超過 MS 標題，見核對頁。這是判斷。").alignment = WRAP
r += 1
# 回填絲杠列
rs = IN["p_screw"]
wsI.cell(rs, 5, f"=F{case_first + 1}").number_format = "#,##0"
wsI.cell(rs, 6, f"=INDEX(F{case_first}:F{case_last},{SCREW_SEL})").number_format = "#,##0"
wsI.cell(rs, 7, f"=F{case_last}").number_format = "#,##0"
for j in (5, 6, 7):
    wsI.cell(rs, j).font = Font(color="000000")
wsI.cell(rs, 9, f'=IF(H{rs}="",F{rs},H{rs})')
SCREW_CASE_RANGE = (case_first, case_last)
wsI.freeze_panes = "C5"

# =============================================================== 計算
widths(wsC, [44, 14, 14, 14, 12, 64])
wsC["A1"] = "計算｜全尺寸人形機器人每台單位經濟（2026 年基準）"
wsC["A1"].font = TITLE
wsC["A2"] = "所有值皆為公式。基準＝輸入頁使用值；低＝全部取低值；高＝全部取高值（低高同向，只作 BOM 區間）。"
CC = {}
cr = 4


def csec(t):
    global cr
    cr += 1
    c = wsC.cell(cr, 1, t)
    c.font = BOLD
    c.fill = GRY
    for j in range(2, 7):
        wsC.cell(cr, j).fill = GRY
    cr += 1
    hdr(wsC, cr, ["項目", "基準", "低", "高", "單位", "等式"])
    cr += 1


def crow(key, label, fb, fl=None, fh=None, unit="USD", eq="", fmt="#,##0"):
    global cr
    CC[key] = cr
    wsC.cell(cr, 1, label)
    for j, f in ((2, fb), (3, fl), (4, fh)):
        if f is not None:
            c = wsC.cell(cr, j, f)
            c.number_format = fmt
    wsC.cell(cr, 5, unit)
    wsC.cell(cr, 6, eq).alignment = WRAP
    cr += 1


def C(k, col="B"):
    return f"'計算'!${col}${CC[k]}"


def tri(expr):
    """expr 使用 {X} 代表輸入引用函式；回傳 (基準, 低, 高)。"""
    return ("=" + expr.format(**{k: U(k) for k in IN}),
            "=" + expr.format(**{k: L(k) for k in IN}),
            "=" + expr.format(**{k: H(k) for k in IN}))


def lc(key, col):
    return f"$" + col + f"${CC[key]}"


csec("1. 旋轉關節單價（每關節）")
for k, lab, ex in [
    ("r_harm", "諧波減速機", "{p_harm}"),
    ("r_motor", "無框力矩馬達", "{p_motor}"),
    ("r_enc", "編碼器 × 每關節顆數", "{n_enc_rot}*{p_enc}"),
    ("r_drv", "驅動板", "{p_drv}"),
    ("r_brg", "軸承", "{p_brg}"),
    ("r_ts", "關節扭矩感測器", "{p_tsens}"),
    ("r_hous", "外殼", "{p_hous}"),
]:
    b, lo, hi = tri(ex)
    crow(k, "  " + lab, b, lo, hi)
first = CC["r_harm"]
crow("rot_unit", "旋轉關節單價", f"=SUM(B{first}:B{cr-1})", f"=SUM(C{first}:C{cr-1})", f"=SUM(D{first}:D{cr-1})",
     eq="減速機＋馬達＋編碼器×2＋驅動板＋軸承＋扭矩感測＋外殼")
wsC.cell(CC["rot_unit"], 1).font = BOLD

csec("2. 線性關節單價（每關節）")
for k, lab, ex in [
    ("l_screw", "行星滾柱絲杠（情形選擇）", "{p_screw}"),
    ("l_motor", "無框力矩馬達", "{p_motor}"),
    ("l_enc", "編碼器 × 每關節顆數", "{n_enc_lin}*{p_enc}"),
    ("l_drv", "驅動板", "{p_drv}"),
    ("l_brg", "軸承", "{p_brg}"),
    ("l_hous", "外殼", "{p_hous}"),
]:
    b, lo, hi = tri(ex)
    crow(k, "  " + lab, b, lo, hi)
first = CC["l_screw"]
crow("lin_unit", "線性關節單價", f"=SUM(B{first}:B{cr-1})", f"=SUM(C{first}:C{cr-1})", f"=SUM(D{first}:D{cr-1})",
     eq="絲杠＋馬達＋編碼器＋驅動板＋軸承＋外殼")
wsC.cell(CC["lin_unit"], 1).font = BOLD

csec("3. 每台 BOM")
bom_lines = [
    ("b_rot", "旋轉關節（數量 × 單價）", "{n_rot}*" , "rot_unit"),
    ("b_lin", "線性關節（數量 × 單價）", "{n_lin}*", "lin_unit"),
]
for k, lab, ex, uk in bom_lines:
    b = "=" + ex.format(**{x: U(x) for x in IN}) + lc(uk, "B")
    lo = "=" + ex.format(**{x: L(x) for x in IN}) + lc(uk, "C")
    hi = "=" + ex.format(**{x: H(x) for x in IN}) + lc(uk, "D")
    crow(k, lab, b, lo, hi)
for k, lab, ex in [
    ("b_hand", "靈巧手 × 2", "{n_hand}*{p_hand}"),
    ("b_ft", "六維力感測器", "{n_ft}*{p_ft}"),
    ("b_head", "頭部（相機等）", "{p_head}"),
    ("b_comp", "主控運算", "{p_comp}"),
    ("b_bat", "電池包（kWh × 每 kWh）", "{bat_kwh}*{p_kwh}"),
    ("b_str", "結構件", "{p_struct}"),
    ("b_harn", "線束與其他電子", "{p_harn}"),
]:
    b, lo, hi = tri(ex)
    crow(k, lab, b, lo, hi)
first = CC["b_rot"]
crow("bom", "每台 BOM 合計", f"=SUM(B{first}:B{cr-1})", f"=SUM(C{first}:C{cr-1})", f"=SUM(D{first}:D{cr-1})",
     eq="Σ 旋轉關節＋Σ 線性關節＋2×靈巧手＋感測＋運算＋電池＋結構與線束")
wsC.cell(CC["bom"], 1).font = BOLD
for j in (2, 3, 4):
    wsC.cell(CC["bom"], j).font = BOLD
    wsC.cell(CC["bom"], j).fill = GRN

csec("4. 每台零組件價值（下游取數）")
for k, lab, ex in [
    ("v_harm", "諧波減速機價值／台", "{n_rot}*{p_harm}"),
    ("v_screw", "行星滾柱絲杠價值／台", "{n_lin}*{p_screw}"),
    ("v_motor", "無框馬達價值／台", "({n_rot}+{n_lin})*{p_motor}"),
    ("v_enc", "編碼器價值／台", "({n_rot}*{n_enc_rot}+{n_lin}*{n_enc_lin})*{p_enc}"),
    ("v_drv", "驅動板價值／台", "({n_rot}+{n_lin})*{p_drv}"),
    ("v_brg", "軸承價值／台", "({n_rot}+{n_lin})*{p_brg}"),
]:
    b, lo, hi = tri(ex)
    crow(k, lab, b, lo, hi)
crow("act", "致動器合計（旋轉＋線性）", f"=B{CC['b_rot']}+B{CC['b_lin']}", f"=C{CC['b_rot']}+C{CC['b_lin']}", f"=D{CC['b_rot']}+D{CC['b_lin']}")
crow("act_sh", "致動器占 BOM", f"=B{CC['act']}/B{CC['bom']}", f"=C{CC['act']}/C{CC['bom']}", f"=D{CC['act']}/D{CC['bom']}", unit="%", fmt="0.0%")
crow("screw_sh", "絲杠占 BOM", f"=B{CC['v_screw']}/B{CC['bom']}", f"=C{CC['v_screw']}/C{CC['bom']}", f"=D{CC['v_screw']}/D{CC['bom']}", unit="%", fmt="0.0%")
crow("harm_sh", "諧波減速機占 BOM", f"=B{CC['v_harm']}/B{CC['bom']}", f"=C{CC['v_harm']}/C{CC['bom']}", f"=D{CC['v_harm']}/D{CC['bom']}", unit="%", fmt="0.0%")
crow("n_harm_pr", "諧波減速機用量／台（全尺寸）", f"={U('n_rot')}", f"={L('n_rot')}", f"={H('n_rot')}", unit="顆", fmt="0")
crow("n_screw_pr", "行星滾柱絲杠用量／台（全尺寸）", f"={U('n_lin')}", f"={L('n_lin')}", f"={H('n_lin')}", unit="支", fmt="0")

csec("5. 每台製造成本與售價")
crow("m_bomy", "BOM ÷ 良率", f"=B{CC['bom']}/{U('yld')}", eq="良率損失攤入")
crow("m_lab", "組裝測試人工", f"={U('labor')}")
crow("m_fix", "固定成本分攤", f"={U('fixcap')}/({U('futil')}*{U('yld')})", eq="年固定成本／單位產能 ÷（工廠稼動 × 良率）")
crow("mfg", "每台製造成本", f"=B{CC['m_bomy']}+B{CC['m_lab']}+B{CC['m_fix']}", eq="BOM ÷ 良率 ＋ 人工 ＋ 年固定成本 ÷（產能 × 稼動 × 良率）")
wsC.cell(CC["mfg"], 1).font = BOLD
crow("gm", "毛利率", f"={U('gm')}", unit="%", fmt="0%")
crow("price", "售價（出廠）", f"=B{CC['mfg']}/(1-B{CC['gm']})", eq="製造成本 ÷（1 − 毛利率）[A]")
wsC.cell(CC["price"], 1).font = BOLD
crow("gp", "每台毛利", f"=B{CC['price']}-B{CC['mfg']}")
crow("conv", "非 BOM 轉換成本（人工＋固定）", f"=B{CC['m_lab']}+B{CC['m_fix']}", eq="情境頁學習曲線一併下降")

csec("6. 使用端 TCO 與回本期")
crow("wage", "替代工人全負擔時薪（依所在地選擇）",
     f"=IF({U('wsel')}=1,{U('wage_us')},{U('cn_wage')}/{U('whours')}/{U('fx')})", unit="USD/小時", fmt="0.00")
crow("capex", "初始投入（售價 ×（1＋部署費））", f"=B{CC['price']}*(1+{U('deploy')})")
crow("crf", "資本回收係數 CRF", f"={U('rate')}/(1-(1+{U('rate')})^(-{U('life')}))", unit="—", fmt="0.0000", eq="r ÷ [1 −（1＋r）^−n]")
crow("hop", "年運轉時數", f"=8760*{U('util')}", unit="小時")
crow("heff", "年有效工時（人工等效）", f"=B{CC['hop']}*(1-{U('down')})*{U('prod')}", unit="小時", eq="8,760 × 運轉 ×（1 − 停機）× 相對生產力")
crow("o_maint", "年維護", f"={U('maint')}*B{CC['price']}")
crow("o_elec", "年電費", f"={U('pkw')}*B{CC['hop']}*{U('elec')}")
crow("o_soft", "年軟體", f"={U('soft')}")
crow("opex", "年營運成本", f"=B{CC['o_maint']}+B{CC['o_elec']}+B{CC['o_soft']}")
crow("ann", "年化機器人成本", f"=B{CC['capex']}*B{CC['crf']}+B{CC['opex']}", eq="初始投入 × CRF ＋ 年營運")
crow("cph", "每有效工時成本", f"=B{CC['ann']}/B{CC['heff']}", unit="USD/小時", fmt="0.00")
crow("val", "替代人工年價值", f"=B{CC['heff']}*B{CC['wage']}", eq="有效工時 × 工人時薪")
crow("net", "年淨節省", f"=B{CC['val']}-B{CC['opex']}")
crow("payback", "回本期（99＝不回本）", f"=IF(B{CC['net']}<=0,99,B{CC['capex']}/B{CC['net']})", unit="年", fmt="0.00",
     eq="（售價＋部署）÷（有效工時 × 時薪 − 年營運）")
wsC.cell(CC["payback"], 1).font = BOLD
wsC.cell(CC["payback"], 2).fill = GRN
crow("fte", "等效工人數", f"=B{CC['heff']}/{U('whours')}", unit="FTE", fmt="0.00")
crow("lease", "損益兩平月租上限（使用者不另付維護）", f"=(B{CC['val']}-B{CC['o_elec']}-B{CC['o_soft']})/12", eq="（替代人工年價值 − 電費 − 軟體）÷ 12")
crow("pass", "是否跨過回本門檻", f'=IF(B{CC["payback"]}<={U("thr")},"已跨過","未跨過")', unit="—", fmt="General")
crow("p_req", "回本＝門檻所需售價", f"={U('thr')}*(B{CC['val']}-B{CC['o_elec']}-B{CC['o_soft']})/(1+{U('deploy')}+{U('thr')}*{U('maint')})",
     eq="P＝T×（V − 電 − 軟）÷（1＋部署＋T×維護）")
crow("bom_req", "回本＝門檻所需 BOM", f"=(B{CC['p_req']}*(1-B{CC['gm']})-B{CC['m_lab']}-B{CC['m_fix']})*{U('yld')}",
     eq="（P×（1−毛利率）− 人工 − 固定分攤）× 良率")
crow("bom_gap", "現況 BOM ÷ 所需 BOM", f"=B{CC['bom']}/B{CC['bom_req']}", unit="倍", fmt="0.00")
crow("prod_req", "現價下回本＝門檻所需相對生產力",
     f"=(B{CC['capex']}/{U('thr')}+B{CC['opex']})/(B{CC['hop']}*(1-{U('down')})*B{CC['wage']})", unit="倍", fmt="0.00")

csec("7. 交叉檢核（資訊比較，非通過／失敗）")
crow("x_ms", "自下而上 BOM ÷ MS 標題 BOM", f"=B{CC['bom']}/{U('ms_bom')}", unit="倍", fmt="0.00", eq="MS 5–6 萬美元")
crow("x_act", "模型致動器占比 − MS 56%", f"=B{CC['act_sh']}-{U('ms_act')}", unit="pp", fmt="0.0%")
crow("x_scr_imp", "覓途 19% 隱含絲杠單價", f"={U('mitu_screw')}*B{CC['bom']}/{U('n_lin')}", eq="19% × BOM ÷ 線性關節數")
crow("x_scr_ms_lo", "MS 絲杠低端 × 用量 ÷ MS BOM", f"='輸入'!$F${case_first+3}*{U('n_lin')}/{U('ms_bom')}", unit="%", fmt="0%")
crow("x_scr_ms_hi", "MS 絲杠高端 × 用量 ÷ MS BOM", f"='輸入'!$F${case_last}*{U('n_lin')}/{U('ms_bom')}", unit="%", fmt="0%", eq="與覓途 19% 無法同時成立")
crow("x_motor", "模型馬達占比（覓途 16%）", f"=B{CC['v_motor']}/B{CC['bom']}", unit="%", fmt="0.0%")
crow("x_brg", "模型軸承占比（覓途 5.5%）", f"=B{CC['v_brg']}/B{CC['bom']}", unit="%", fmt="0.0%")
crow("x_drv", "模型驅動板占比（X.PIN 控制器 12%）", f"=B{CC['v_drv']}/B{CC['bom']}", unit="%", fmt="0.0%")
crow("x_str", "模型結構件占比（X.PIN 10%）", f"=B{CC['b_str']}/B{CC['bom']}", unit="%", fmt="0.0%")
crow("x_uni", "宇樹人形隱含單台成本（美元）", f"={U('uni_asp')}*(1-{U('uni_gm')})/{U('fx')}", eq="ASP ×（1 − 毛利率）÷ 匯率；中小型機混合，不可直比全尺寸")
crow("x_ubt", "優必選全尺寸每台營收（美元，含服務）", f"={U('ubt_rev')}/{U('ubt_units')}/{U('fx')}")
crow("x_musk", "模型製造成本 ÷ Musk 3 萬美元說法", f"=B{CC['mfg']}/{H('musk')}", unit="倍", fmt="0.00")
crow("x_bofa", "模型 BOM ÷ BofA 2025 年底 3.5 萬", f"=B{CC['bom']}/{U('bofa_25')}", unit="倍", fmt="0.00")
wsC.freeze_panes = "B4"

# =============================================================== 情境
widths(wsN, [40] + [12] * 12 + [40])
wsN["A1"] = "情境｜2026–2035 全球出貨 × 學習曲線 → BOM、售價、回本期、關鍵零組件需求"
wsN["A1"].font = TITLE
wsN["A2"] = ("學習曲線：BOM(t)＝BOM₂₀₂₆ ×（累積中點(t) ÷ 累積中點(2026)）^log₂(1 − 學習率)；轉換成本與零組件單價同比例下降。"
             "出貨錨點年 2026／2030／2035 為輸入，中間年幾何內插。全球平均每台用量＝全尺寸用量 × 採用比例（小型機多用行星減速、不用絲杠）。")
wsN["A2"].alignment = WRAP
wsN.merge_cells("A2:N2")
wsN.row_dimensions[2].height = 45
YEARS = list(range(2026, 2036))
YC = {y: CL(2 + i) for i, y in enumerate(YEARS)}  # B..K

nr = 4
hdr(wsN, nr, ["情境輸入", "低", "基準", "高", "標記", "出處"], col=1)
nr += 1
SIN = {}
SC_IN = [
    ("s26", "全球出貨 2026（台）", 40000, 55000, 75000, "低 [Assumed]；基準 [Interested-party]；高 [Interested-party]",
     "基準：TrendForce 全球 >5 萬（" + S["TF"] + "）、MS 中國 5 萬（" + S["MSQ"] + "）；高：Goldman 7.5 萬（" + S["GS"] + "）；2026H1 追蹤 1.9–2.5 萬（" + S["NEXT"] + "）", "#,##0"),
    ("s30", "全球出貨 2030（台）", 200000, 500000, 890000, "低 [Assumed]；基準 [Assumed]；高 [Interested-party]",
     "基準以 MS 中國 2030 年 44.6 萬台、中國占八成以上推估；高：Goldman 89 萬", "#,##0"),
    ("s35", "全球出貨 2035（台）", 1000000, 3000000, 6500000, "低 [Assumed]；基準 [Assumed]；高 [Interested-party]",
     "高：Goldman 648 萬（年度或累積口徑未明）；Goldman 前版 138 萬", "#,##0"),
    ("lr", "學習率（每倍增累積產量降幅）", 0.10, 0.15, 0.20, "[Analogy]", "工業機器人與車用電池成本曲線", "0%"),
    ("p30", "相對人工生產力 2030（2026→2030 線性，之後持平）", 0.35, 0.60, 0.90, "[Assumed]", "無公開實測", "0.00"),
    ("hsh", "全球旋轉關節採用諧波減速機比例", 0.30, 0.50, 0.70, "[Assumed]", "宇樹等小型機多用行星減速；Tesla、Figure 用諧波", "0%"),
    ("ssh", "全球出貨採用滾柱絲杠線性關節比例", 0.15, 0.30, 0.50, "[Assumed]", "只有混合架構（Optimus 類）用絲杠；中國小型機多為全旋轉", "0%"),
]
for k, lab, lo, b, hi, tag, src, fmt in SC_IN:
    SIN[k] = nr
    wsN.cell(nr, 1, lab)
    for j, v in ((2, lo), (3, b), (4, hi)):
        c = wsN.cell(nr, j, v)
        c.font = BLUE
        c.number_format = fmt
        if "[Assumed]" in tag or "[Analogy]" in tag:
            c.fill = YEL
    wsN.cell(nr, 5, tag).alignment = WRAP
    wsN.cell(nr, 6, src).alignment = WRAP
    wsN.merge_cells(start_row=nr, start_column=6, end_row=nr, end_column=14)
    wsN.row_dimensions[nr].height = 30
    nr += 1


def SI(k, s):  # s=0,1,2
    return f"'情境'!${'BCD'[s]}${SIN[k]}"


nr += 1
SCN = {}  # (s, key) -> row
SNAME = ["低", "基準", "高"]
for s in range(3):
    wsN.cell(nr, 1, f"情境：{SNAME[s]}").font = BOLD
    for j in range(1, 13):
        wsN.cell(nr, j).fill = GRY
    nr += 1
    hdr(wsN, nr, ["項目"] + [str(y) for y in YEARS] + ["單位"])
    yr_row = nr
    for y in YEARS:
        wsN.cell(nr, 2 + YEARS.index(y), y).font = HDR
    nr += 1

    def srow(key, label, fn, unit, fmt="#,##0"):
        global nr
        SCN[(s, key)] = nr
        wsN.cell(nr, 1, label)
        for y in YEARS:
            c = wsN.cell(nr, 2 + YEARS.index(y), fn(y))
            c.number_format = fmt
        wsN.cell(nr, 12, unit)
        nr += 1

    def R(key, y):
        return f"{YC[y]}{SCN[(s, key)]}"

    def YH(y):
        return f"{YC[y]}${yr_row}"

    def ship(y):
        if y == 2026:
            return f"={SI('s26', s)}"
        if y <= 2030:
            return f"={SI('s26', s)}*({SI('s30', s)}/{SI('s26', s)})^(({YH(y)}-2026)/4)"
        return f"={SI('s30', s)}*({SI('s35', s)}/{SI('s30', s)})^(({YH(y)}-2030)/5)"

    srow("ship", "全球人形出貨", ship, "台")
    srow("cum", "年底累積出貨", lambda y: f"={U('q0')}+{R('ship', y)}" if y == 2026 else f"={R('cum', y-1)}+{R('ship', y)}", "台")
    srow("mid", "累積中點", lambda y: f"={R('cum', y)}-{R('ship', y)}/2", "台")
    srow("lf", "學習係數", lambda y: f"=({R('mid', y)}/${YC[2026]}${SCN[(s, 'mid')]})^LOG(1-{SI('lr', s)},2)", "倍", "0.000")
    srow("bom", "全尺寸每台 BOM", lambda y: f"={C('bom')}*{R('lf', y)}", "USD")
    srow("mfg", "每台製造成本", lambda y: f"={R('bom', y)}/{U('yld')}+{C('conv')}*{R('lf', y)}", "USD")
    srow("price", "售價", lambda y: f"={R('mfg', y)}/(1-{U('gm')})", "USD")
    srow("prod", "相對人工生產力",
         lambda y: f"={U('prod')}+({SI('p30', s)}-{U('prod')})*MIN(1,({YH(y)}-2026)/4)", "倍", "0.00")
    srow("capex", "初始投入", lambda y: f"={R('price', y)}*(1+{U('deploy')})", "USD")
    srow("heff", "年有效工時", lambda y: f"={C('hop')}*(1-{U('down')})*{R('prod', y)}", "小時")
    srow("opex", "年營運成本", lambda y: f"={U('maint')}*{R('price', y)}+{C('o_elec')}+{U('soft')}", "USD")
    srow("payback", "回本期（99＝不回本）",
         lambda y: f"=IF({R('heff', y)}*{C('wage')}-{R('opex', y)}<=0,99,{R('capex', y)}/({R('heff', y)}*{C('wage')}-{R('opex', y)}))", "年", "0.00")
    srow("cph", "每有效工時成本", lambda y: f"=({R('capex', y)}*{C('crf')}+{R('opex', y)})/{R('heff', y)}", "USD/小時", "0.00")
    srow("flag", "跨門檻年（輔助：未跨＝9999）", lambda y: f"=IF({R('payback', y)}<={U('thr')},{YH(y)},9999)", "年", "0")
    srow("harm_pr", "全球平均每台諧波減速機用量", lambda y: f"={U('n_rot')}*{SI('hsh', s)}", "顆", "0.0")
    srow("scr_pr", "全球平均每台滾柱絲杠用量", lambda y: f"={U('n_lin')}*{SI('ssh', s)}", "支", "0.0")
    srow("harm_u", "人形用諧波減速機需求", lambda y: f"={R('ship', y)}*{R('harm_pr', y)}", "顆")
    srow("scr_u", "人形用滾柱絲杠需求", lambda y: f"={R('ship', y)}*{R('scr_pr', y)}", "支")
    srow("harm_p", "諧波減速機單價（學習曲線）", lambda y: f"={U('p_harm')}*{R('lf', y)}", "USD")
    srow("scr_p", "滾柱絲杠單價（學習曲線）", lambda y: f"={U('p_screw')}*{R('lf', y)}", "USD")
    srow("harm_m", "人形用諧波減速機市場", lambda y: f"={R('harm_u', y)}*{R('harm_p', y)}/1e6", "百萬 USD", "#,##0.0")
    srow("scr_m", "人形用滾柱絲杠市場", lambda y: f"={R('scr_u', y)}*{R('scr_p', y)}/1e6", "百萬 USD", "#,##0.0")
    cy = nr
    SCN[(s, "cross")] = nr
    wsN.cell(nr, 1, "跨過回本門檻年份").font = BOLD
    fr = SCN[(s, "flag")]
    c = wsN.cell(nr, 2, f'=IF(MIN(B{fr}:K{fr})=9999,"2035 後",MIN(B{fr}:K{fr}))')
    c.font = BOLD
    c.fill = GRN
    nr += 2
wsN.freeze_panes = "B4"

# =============================================================== 敏感度
widths(wsS, [32, 10, 12, 12, 12, 12, 12, 12] + [11] * 20)
wsS["A1"] = "敏感度｜承重變數排序（先看這頁，再看計算頁基準結果）"
wsS["A1"].font = TITLE
wsS["A2"] = ("每列只改一個輸入（低值／高值），其餘維持使用值，整條式子（BOM → 製造成本 → 售價 → 回本期）完整重算。"
             "依回本期擺幅排序；回本期 99＝不回本。壽命與資金成本只影響每有效工時成本，不影響簡單回本期。工人所在地是選擇，另列。")
wsS["A2"].alignment = WRAP
wsS.merge_cells("A2:L2")
wsS.row_dimensions[2].height = 45

# 敏感變數：(key, label, kind, mult_expr)
QTY = lambda e: e
SENS = [
    ("prod", "相對人工生產力", "drv", None),
    ("util", "運轉時數（班次）", "drv", None),
    ("down", "停機率", "drv", None),
    ("gm", "整機毛利率", "drv", None),
    ("deploy", "部署整合費", "drv", None),
    ("maint", "年維護費", "drv", None),
    ("yld", "整機良率", "drv", None),
    ("labor", "組裝測試人工", "drv", None),
    ("fixcap", "固定成本／單位產能", "drv", None),
    ("futil", "工廠產能利用率", "drv", None),
    ("soft", "年軟體授權", "drv", None),
    ("pkw", "平均耗電", "drv", None),
    ("elec", "電價", "drv", None),
    ("p_screw", "行星滾柱絲杠單價", "bom", "{n_lin}"),
    ("p_harm", "諧波減速機單價", "bom", "{n_rot}"),
    ("p_motor", "無框馬達單價", "bom", "({n_rot}+{n_lin})"),
    ("p_hand", "靈巧手單價", "bom", "{n_hand}"),
    ("p_ft", "六維力感測器單價", "bom", "{n_ft}"),
    ("p_drv", "驅動板單價", "bom", "({n_rot}+{n_lin})"),
    ("p_comp", "主控運算", "bom", "1"),
    ("p_struct", "結構件", "bom", "1"),
    ("p_tsens", "關節扭矩感測器", "bom", "{n_rot}"),
    ("p_enc", "編碼器單價", "bom", "({n_rot}*{n_enc_rot}+{n_lin}*{n_enc_lin})"),
    ("p_brg", "軸承", "bom", "({n_rot}+{n_lin})"),
    ("p_hous", "關節外殼", "bom", "({n_rot}+{n_lin})"),
    ("p_head", "頭部", "bom", "1"),
    ("p_harn", "線束", "bom", "1"),
    ("p_kwh", "電池每 kWh", "bom", "{bat_kwh}"),
    ("n_rot", "旋轉關節數", "bomq", "rot_unit"),
    ("n_lin", "線性關節數", "bomq", "lin_unit"),
    ("n_ft", "力感測器顆數", "bom", "{p_ft}"),
    ("life", "使用壽命（只影響每工時成本）", "drv", None),
    ("rate", "資金成本（只影響每工時成本）", "drv", None),
]
DRV = ["yld", "labor", "fixcap", "futil", "gm", "deploy", "maint", "life", "rate", "util", "down", "prod", "pkw", "elec", "soft"]
# detail table columns
DCOLS = ["變數", "情形", "測試值", "BOM"] + [k for k in DRV] + ["時薪", "製造成本", "售價", "初始投入", "有效工時", "年營運", "回本期", "每工時成本"]
NS = len(SENS)
rank_top = 5
summ_top = rank_top + NS + 4
base_row_r = None

# --- 排序表（頂端）
wsS.cell(rank_top - 1, 1, "A. 承重變數排序（依回本期擺幅）").font = BOLD
hdr(wsS, rank_top, ["排名", "變數", "低值", "高值", "回本期@低", "回本期@高", "回本期擺幅", "每工時成本擺幅"])
# --- 未排序摘要
s_hdr = summ_top
detail_top = summ_top + NS + 22
wsS.cell(s_hdr - 1, 1, "B. 未排序摘要（排序表讀這裡）").font = BOLD
hdr(wsS, s_hdr, ["變數", "低值", "高值", "回本期@低", "回本期@高", "擺幅", "排序鍵", "每工時成本擺幅", "BOM@低", "BOM@高"])

# --- 明細
wsS.cell(detail_top - 1, 1, "D. 明細（每列為一次完整重算）").font = BOLD
hdr(wsS, detail_top, DCOLS)
col = {n: i + 1 for i, n in enumerate(DCOLS)}


def dref(row, name):
    return f"{CL(col[name])}{row}"


dr = detail_top + 1


def detail_row(label, case, test_f, key, kind, mult):
    global dr
    wsS.cell(dr, 1, label)
    wsS.cell(dr, 2, case)
    if test_f is not None:
        wsS.cell(dr, 3, test_f).number_format = "#,##0.00"
    # BOM
    if kind == "bom" and key:
        m = mult.format(**{k: U(k) for k in IN})
        bomf = f"={C('bom')}+{m}*({dref(dr, '測試值')}-{U(key)})"
    elif kind == "bomq":
        bomf = f"={C('bom')}+{C(mult)}*({dref(dr, '測試值')}-{U(key)})"
    else:
        bomf = f"={C('bom')}"
    wsS.cell(dr, col["BOM"], bomf).number_format = "#,##0"
    for k in DRV:
        f = f"={dref(dr, '測試值')}" if (kind == "drv" and key == k) else f"={U(k)}"
        wsS.cell(dr, col[k], f).number_format = "0.00"
    wsS.cell(dr, col["時薪"], f"={C('wage')}" if case != "地" else test_f).number_format = "0.00"
    g = lambda n: dref(dr, n)
    wsS.cell(dr, col["製造成本"], f"={g('BOM')}/{g('yld')}+{g('labor')}+{g('fixcap')}/({g('futil')}*{g('yld')})").number_format = "#,##0"
    wsS.cell(dr, col["售價"], f"={g('製造成本')}/(1-{g('gm')})").number_format = "#,##0"
    wsS.cell(dr, col["初始投入"], f"={g('售價')}*(1+{g('deploy')})").number_format = "#,##0"
    wsS.cell(dr, col["有效工時"], f"=8760*{g('util')}*(1-{g('down')})*{g('prod')}").number_format = "#,##0"
    wsS.cell(dr, col["年營運"], f"={g('maint')}*{g('售價')}+{g('pkw')}*8760*{g('util')}*{g('elec')}+{g('soft')}").number_format = "#,##0"
    wsS.cell(dr, col["回本期"],
             f"=IF({g('有效工時')}*{g('時薪')}-{g('年營運')}<=0,99,{g('初始投入')}/({g('有效工時')}*{g('時薪')}-{g('年營運')}))").number_format = "0.00"
    wsS.cell(dr, col["每工時成本"],
             f"=({g('初始投入')}*{g('rate')}/(1-(1+{g('rate')})^(-{g('life')}))+{g('年營運')})/{g('有效工時')}").number_format = "0.00"
    dr += 1
    return dr - 1


base_d = detail_row("基準（全部使用值）", "基準", None, None, "drv", None)
pairs = {}
for key, label, kind, mult in SENS:
    rl = detail_row(label, "低", f"={L(key)}", key, kind, mult)
    rh = detail_row(label, "高", f"={H(key)}", key, kind, mult)
    pairs[key] = (rl, rh)
# 絲杠五情形
screw_rows = []
for i in range(5):
    rr = detail_row(f"絲杠情形 {i+1}", "情形", f"='輸入'!$F${case_first+i}", "p_screw", "bom", "{n_lin}")
    screw_rows.append(rr)
# 工人所在地
wage_us_r = detail_row("工人在美國", "地", f"={U('wage_us')}", None, "drv", None)
wage_cn_r = detail_row("工人在中國", "地", f"={U('cn_wage')}/{U('whours')}/{U('fx')}", None, "drv", None)

# 摘要
sr = s_hdr + 1
summ_rows = []
for key, label, kind, mult in SENS:
    rl, rh = pairs[key]
    wsS.cell(sr, 1, label)
    wsS.cell(sr, 2, f"=C{rl}").number_format = "#,##0.00"
    wsS.cell(sr, 3, f"=C{rh}").number_format = "#,##0.00"
    wsS.cell(sr, 4, f"={dref(rl, '回本期')}").number_format = "0.00"
    wsS.cell(sr, 5, f"={dref(rh, '回本期')}").number_format = "0.00"
    wsS.cell(sr, 6, f"=ABS(E{sr}-D{sr})").number_format = "0.00"
    wsS.cell(sr, 7, f"=F{sr}+ABS({dref(rh, '每工時成本')}-{dref(rl, '每工時成本')})/1E6+ROW()/1E9").number_format = "0.000000"
    wsS.cell(sr, 8, f"=ABS({dref(rh, '每工時成本')}-{dref(rl, '每工時成本')})").number_format = "0.00"
    wsS.cell(sr, 9, f"={dref(rl, 'BOM')}").number_format = "#,##0"
    wsS.cell(sr, 10, f"={dref(rh, 'BOM')}").number_format = "#,##0"
    summ_rows.append(sr)
    sr += 1
s1, s2 = summ_rows[0], summ_rows[-1]
for k in range(1, NS + 1):
    rr = rank_top + k
    wsS.cell(rr, 1, k)
    m = f"MATCH(LARGE($G${s1}:$G${s2},{k}),$G${s1}:$G${s2},0)"
    for j, cc in ((2, "A"), (3, "B"), (4, "C"), (5, "D"), (6, "E"), (7, "F"), (8, "H")):
        c = wsS.cell(rr, j, f"=INDEX(${cc}${s1}:${cc}${s2},{m})")
        c.number_format = "0.00" if j >= 5 else "#,##0.00"
    if k <= 3:
        for j in range(1, 9):
            wsS.cell(rr, j).fill = YEL

# C. 基準結果、絲杠情形、所在地
cr2 = s2 + 2
wsS.cell(cr2, 1, "C. 基準結果與兩個特別表").font = BOLD
cr2 += 1
hdr(wsS, cr2, ["項目", "BOM", "製造成本", "售價", "回本期", "每工時成本", "絲杠占 BOM", "對 MS 5.5 萬"])
cr2 += 1
SENS_BASE_ROW = cr2
for lab, rr in [("基準（＝計算頁）", base_d)] + [(f"絲杠情形 {i+1}：" + SCREW_CASES[i][1], screw_rows[i]) for i in range(5)] + \
        [("工人在美國（BLS）", wage_us_r), ("工人在中國（¥8–12 萬/年）", wage_cn_r)]:
    wsS.cell(cr2, 1, lab).alignment = WRAP
    wsS.cell(cr2, 2, f"={dref(rr, 'BOM')}").number_format = "#,##0"
    wsS.cell(cr2, 3, f"={dref(rr, '製造成本')}").number_format = "#,##0"
    wsS.cell(cr2, 4, f"={dref(rr, '售價')}").number_format = "#,##0"
    wsS.cell(cr2, 5, f"={dref(rr, '回本期')}").number_format = "0.00"
    wsS.cell(cr2, 6, f"={dref(rr, '每工時成本')}").number_format = "0.00"
    tv = dref(rr, "測試值") if lab.startswith("絲杠") else U("p_screw")
    wsS.cell(cr2, 7, f"={tv}*{U('n_lin')}/{dref(rr, 'BOM')}").number_format = "0.0%"
    wsS.cell(cr2, 8, f"={dref(rr, 'BOM')}/{U('ms_bom')}").number_format = "0.00"
    cr2 += 1
SCREW_TABLE_FIRST = SENS_BASE_ROW + 1
wsS.freeze_panes = "B4"

# =============================================================== 下游取數
widths(wsD, [26, 40, 14, 14, 14, 12, 50, 26])
wsD["A1"] = "下游取數｜給個股模型（Tesla、哈默納科、Nabtesco、上銀、Fanuc、安川等）讀的輸出"
wsD["A1"].font = TITLE
wsD["A2"] = "每列有 Excel 具名範圍 DS_<代號>（基準值）。口徑：全尺寸＝Optimus Gen 2 類；全球平均＝全部人形出貨（含小型機）。金額美元。"
hdr(wsD, 4, ["代號", "名稱", "低", "基準", "高", "單位", "口徑", "來源"])
dr2 = 5
DS = {}


def ds(code, name, lo, b, hi, unit, scope, src, fmt="#,##0"):
    global dr2
    DS[code] = dr2
    wsD.cell(dr2, 1, code)
    wsD.cell(dr2, 2, name).alignment = WRAP
    for j, f in ((3, lo), (4, b), (5, hi)):
        if f is not None:
            wsD.cell(dr2, j, f).number_format = fmt
    wsD.cell(dr2, 6, unit)
    wsD.cell(dr2, 7, scope).alignment = WRAP
    wsD.cell(dr2, 8, src).alignment = WRAP
    wb.defined_names[f"DS_{code}"] = DefinedName(f"DS_{code}", attr_text=f"'下游取數'!$D${dr2}")
    dr2 += 1


def dsec(t):
    global dr2
    wsD.cell(dr2, 1, t).font = BOLD
    for j in range(1, 9):
        wsD.cell(dr2, j).fill = GRY
    dr2 += 1


dsec("1. 每台 BOM 與成本（2026，全尺寸）")
ds("HUM_BOM_2026", "每台 BOM", f"={C('bom','C')}", f"={C('bom')}", f"={C('bom','D')}", "USD/台", "全尺寸、硬體、不含軟體；低＝全部取低，高＝全部取高", "計算!B/C/D BOM 合計")
ds("HUM_BOM_SCREW_CASES", "每台 BOM（絲杠情形 1／3／5）", f"='敏感度'!B{SCREW_TABLE_FIRST}", f"='敏感度'!B{SCREW_TABLE_FIRST+2}", f"='敏感度'!B{SCREW_TABLE_FIRST+4}", "USD/台", "其他輸入不變，只換絲杠單價", "敏感度 C 表")
ds("HUM_ACT_SHARE", "致動器占 BOM", f"={C('act_sh','C')}", f"={C('act_sh')}", f"={C('act_sh','D')}", "%", "全尺寸", "計算", "0.0%")
ds("HUM_MFG_2026", "每台製造成本", None, f"={C('mfg')}", None, "USD/台", "BOM÷良率＋人工＋固定分攤", "計算")
ds("HUM_PRICE_2026", "每台售價（出廠）", None, f"={C('price')}", None, "USD/台", "製造成本÷（1−毛利率）", "計算")
ds("HUM_PAYBACK_2026", "對美國工人回本期", None, f"={C('payback')}", None, "年", "99＝不回本", "計算", "0.00")
ds("HUM_BOM_REQ", "回本＝門檻所需 BOM", None, f"={C('bom_req')}", None, "USD/台", "其他輸入不變", "計算", "#,##0")
ds("HUM_PROD_REQ", "現價下回本＝門檻所需相對生產力", None, f"={C('prod_req')}", None, "倍", "相對一名工人", "計算", "0.00")

dsec("2. 零組件單價與每台用量（2026，全尺寸）")
for code, name, key, unit, fmt in [
    ("P_HARMONIC", "諧波減速機單價", "p_harm", "USD/顆", "#,##0"),
    ("P_ROLLER_SCREW", "行星滾柱絲杠單價（低＝情形2，高＝情形5）", "p_screw", "USD/支", "#,##0"),
    ("P_FRAMELESS_MOTOR", "無框力矩馬達單價", "p_motor", "USD/顆", "#,##0"),
    ("P_ENCODER", "編碼器單價", "p_enc", "USD/顆", "#,##0"),
    ("P_DRIVER", "驅動板單價", "p_drv", "USD/關節", "#,##0"),
    ("P_BEARING", "軸承（每關節）", "p_brg", "USD/關節", "#,##0"),
    ("P_HAND", "靈巧手單價", "p_hand", "USD/隻", "#,##0"),
    ("P_FT_SENSOR", "六維力感測器單價", "p_ft", "USD/顆", "#,##0"),
    ("P_COMPUTE", "主控運算", "p_comp", "USD/台", "#,##0"),
    ("P_BATTERY_KWH", "電池每 kWh", "p_kwh", "USD/kWh", "#,##0"),
    ("Q_ROTARY", "諧波減速機（旋轉關節）用量／台", "n_rot", "顆/台", "0"),
    ("Q_ROLLER_SCREW", "行星滾柱絲杠（線性關節）用量／台", "n_lin", "支/台", "0"),
    ("Q_MOTOR", "無框馬達用量／台", None, "顆/台", "0"),
    ("Q_FT", "六維力感測器用量／台", "n_ft", "顆/台", "0"),
    ("Q_BATTERY_KWH", "電池容量／台", "bat_kwh", "kWh/台", "0.0"),
]:
    if key:
        ds(code, name, f"={L(key)}", f"={U(key)}", f"={H(key)}", unit, "全尺寸", "輸入頁", fmt)
    else:
        ds(code, name, f"={L('n_rot')}+{L('n_lin')}", f"={U('n_rot')}+{U('n_lin')}", f"={H('n_rot')}+{H('n_lin')}", unit, "全尺寸，每關節 1 顆", "輸入頁", fmt)
for code, name, k in [("V_HARMONIC", "諧波減速機價值／台", "v_harm"), ("V_ROLLER_SCREW", "滾柱絲杠價值／台", "v_screw"),
                      ("V_MOTOR", "無框馬達價值／台", "v_motor")]:
    ds(code, name, f"={C(k,'C')}", f"={C(k)}", f"={C(k,'D')}", "USD/台", "全尺寸", "計算")

dsec("3. 全球人形出貨量情境（台）")
for y in YEARS:
    ds(f"SHIP_{y}", f"全球人形出貨 {y}", *[f"='情境'!{YC[y]}{SCN[(s, 'ship')]}" for s in range(3)], "台/年", "全球、含小型機", "情境頁")
dsec("4. 人形用諧波減速機與滾柱絲杠需求（全球）")
ds("Q_HARMONIC_AVG", "全球平均每台諧波減速機用量", *[f"='情境'!B{SCN[(s, 'harm_pr')]}" for s in range(3)], "顆/台", "全尺寸用量 × 採用比例", "情境頁", "0.0")
ds("Q_SCREW_AVG", "全球平均每台滾柱絲杠用量", *[f"='情境'!B{SCN[(s, 'scr_pr')]}" for s in range(3)], "支/台", "全尺寸用量 × 採用比例", "情境頁", "0.0")
for y in YEARS:
    ds(f"HARM_UNITS_{y}", f"人形用諧波減速機需求 {y}", *[f"='情境'!{YC[y]}{SCN[(s, 'harm_u')]}" for s in range(3)], "顆/年", "全球人形", "情境頁")
for y in YEARS:
    ds(f"SCREW_UNITS_{y}", f"人形用滾柱絲杠需求 {y}", *[f"='情境'!{YC[y]}{SCN[(s, 'scr_u')]}" for s in range(3)], "支/年", "全球人形", "情境頁")
for y in (2026, 2028, 2030, 2035):
    ds(f"HARM_MKT_{y}", f"人形用諧波減速機市場 {y}", *[f"='情境'!{YC[y]}{SCN[(s, 'harm_m')]}" for s in range(3)], "百萬 USD", "單價隨學習曲線", "情境頁", "#,##0.0")
    ds(f"SCREW_MKT_{y}", f"人形用滾柱絲杠市場 {y}", *[f"='情境'!{YC[y]}{SCN[(s, 'scr_m')]}" for s in range(3)], "百萬 USD", "單價隨學習曲線；絲杠情形 3", "情境頁", "#,##0.0")
dsec("5. 成本路徑與門檻")
for y in (2028, 2030, 2035):
    ds(f"HUM_BOM_{y}", f"全尺寸每台 BOM {y}", *[f"='情境'!{YC[y]}{SCN[(s, 'bom')]}" for s in range(3)], "USD/台", "學習曲線；低＝低出貨情境（BOM 較高）", "情境頁")
    ds(f"HUM_PRICE_{y}", f"全尺寸售價 {y}", *[f"='情境'!{YC[y]}{SCN[(s, 'price')]}" for s in range(3)], "USD/台", "同上", "情境頁")
ds("HUM_CROSS_YEAR", "跨過回本門檻年份", *[f"='情境'!B{SCN[(s, 'cross')]}" for s in range(3)], "年", "回本 ≤ 門檻（基準 3 年），美國工人", "情境頁", "0")
wsD.freeze_panes = "C5"

# =============================================================== 追蹤指標
widths(wsT, [30, 8, 38, 34, 12, 60, 40])
wsT["A1"] = "追蹤指標（框架第 5 節，含最新值與日期）"
wsT["A1"].font = TITLE
hdr(wsT, 3, ["指標", "頻率", "資料來源", "最新值", "日期", "出處", "為何重要／對應輸入"])
TRK = [
    ("全球人形出貨（年）", "季", "IDC、Omdia、Counterpoint、TrendForce；各廠財報", "2025：約 1.3 萬（Omdia）／約 1.8 萬（IDC）", "2026-01", S["OMDIA"] + "；" + S["IDC"], "學習曲線 Q₀；情境頁出貨"),
    ("2026 上半年追蹤出貨", "季", "產業追蹤", "1.9–2.5 萬台（已超過 2025 全年）", "2026-H1", S["NEXT"], "情境 2026 錨點"),
    ("機構出貨預測", "事件", "MS、Goldman、TrendForce", "MS 中國 2026 5 萬、2030 44.6 萬；Goldman 全球 2026 7.5 萬、2030 89 萬、2035 約 648 萬", "2026-06／2026-09", S["MSQ"] + "；" + S["GS"], "情境錨點"),
    ("宇樹人形 ASP 與毛利率", "季", "宇樹財報（科創板）", "ASP ¥16.76 萬（年減 35.7%）；毛利率 62.9%", "2025 前三季", S["UNI"], "唯一有售價＋毛利的人形廠；核對"),
    ("優必選全尺寸台數與毛利", "半年", "優必選中報、年報", "921 台、營收 ¥5.903 億；毛利率 44.7%（分部口徑待核）", "2026H1", S["UBT"], "工業場景真實訂單；毛利率高值"),
    ("Optimus 量產節奏與成本說法", "季", "Tesla 法說、股東信", "Fremont 設計年產 100 萬台；達每週 2 萬台至少需一年", "2026Q3", S["TSLA"], "產能、固定成本分攤"),
    ("滾柱絲杠報價", "季", "五洲新春、貝斯特、雙林法說；上銀月營收", "五洲新春 2025 均價約 ¥1,360（年降 39%）；雙林 ¥499；國際品 >¥5,000", "2025／2026-01", S["SCREW"] + "；" + S["CD"], "解決絲杠單價矛盾（承重）"),
    ("諧波減速機報價", "季", "哈默納科、綠的諧波法說", "高階 800–1,500 美元；中國便宜 30–50%", "2026", S["PICEA"], "旋轉關節成本"),
    ("工人全負擔成本", "季", "BLS ECEC；國家統計局", "美國私部門 46.89 美元/小時", "2026-06", S["BLS"], "TCO 分母"),
    ("中國政策與標準", "事件", "發改委、工信部人形機器人標準化技術委員會", "標準體系 2026 版；發改委泡沫警示（150 家以上）", "2026", "見框架 Sources", "補貼與產能過剩"),
    ("首批工廠部署的生產力／MTBF", "事件", "Figure、Agility、優必選客戶揭露", "未取得", "—", "—", "相對生產力、停機率、壽命（回本期最大承重）"),
]
for i, row in enumerate(TRK):
    for j, v in enumerate(row, 1):
        wsT.cell(4 + i, j, v).alignment = WRAP

# =============================================================== 核對
widths(wsK, [6, 60, 16, 16, 10])
wsK["A1"] = "核對（內部一致性；全部 PASS 才交付）"
wsK["A1"].font = TITLE
hdr(wsK, 3, ["#", "檢查", "值 A", "值 B", "結果"])
CHK = [
    ("BOM 合計＝各列加總", f"={C('bom')}", f"=SUM('計算'!B{CC['b_rot']}:B{CC['b_harn']})"),
    ("致動器＝旋轉＋線性", f"={C('act')}", f"={C('b_rot')}+{C('b_lin')}"),
    ("敏感度基準列回本期＝計算頁", f"='敏感度'!E{SENS_BASE_ROW}", f"={C('payback')}"),
    ("敏感度基準列 BOM＝計算頁", f"='敏感度'!B{SENS_BASE_ROW}", f"={C('bom')}"),
    ("敏感度基準列每工時成本＝計算頁", f"='敏感度'!F{SENS_BASE_ROW}", f"={C('cph')}"),
    ("情境 2026 BOM（基準）＝計算頁", f"='情境'!B{SCN[(1, 'bom')]}", f"={C('bom')}"),
    ("情境 2026 回本期（基準）＝計算頁", f"='情境'!B{SCN[(1, 'payback')]}", f"={C('payback')}"),
    ("絲杠情形 3 列 BOM＝計算頁（選 3 時）", f"=IF({SCREW_SEL}=3,'敏感度'!B{SCREW_TABLE_FIRST+2},{C('bom')})", f"={C('bom')}"),
    ("下游 BOM＝計算頁", f"='下游取數'!D{DS['HUM_BOM_2026']}", f"={C('bom')}"),
    ("下游 2030 出貨＝情境頁", f"='下游取數'!D{DS['SHIP_2030']}", f"='情境'!F{SCN[(1, 'ship')]}"),
    ("情境 2030 出貨＝錨點輸入", f"='情境'!F{SCN[(1, 'ship')]}", f"={SI('s30', 1)}"),
    ("回本＝門檻所需售價代回後回本期＝門檻", f"={C('p_req')}*(1+{U('deploy')})/({C('val')}-{C('o_elec')}-{C('o_soft')}-{U('maint')}*{C('p_req')})", f"={U('thr')}"),
    ("BOM 低 ≤ 基準 ≤ 高", f"=IF(AND({C('bom','C')}<={C('bom')},{C('bom')}<={C('bom','D')}),1,0)", "=1"),
    ("絲杠情形選擇在 1–5", f"=IF(AND({SCREW_SEL}>=1,{SCREW_SEL}<=5),1,0)", "=1"),
    ("學習係數 2026＝1（三情境）", f"='情境'!B{SCN[(0, 'lf')]}+'情境'!B{SCN[(1, 'lf')]}+'情境'!B{SCN[(2, 'lf')]}", "=3"),
    ("排序表第 1 名擺幅＝摘要最大擺幅", f"='敏感度'!G{rank_top+1}", f"=MAX('敏感度'!F{s1}:F{s2})"),
]
for i, (lab, a, b) in enumerate(CHK):
    rr = 4 + i
    wsK.cell(rr, 1, i + 1)
    wsK.cell(rr, 2, lab)
    wsK.cell(rr, 3, a).number_format = "#,##0.0000"
    wsK.cell(rr, 4, b).number_format = "#,##0.0000"
    wsK.cell(rr, 5, f'=IF(ABS(C{rr}-D{rr})<=1E-6*MAX(1,ABS(D{rr})),"PASS","FAIL")')
kl = 4 + len(CHK)
wsK.cell(kl + 1, 2, "全部通過").font = BOLD
wsK.cell(kl + 1, 5, f'=IF(COUNTIF(E4:E{kl-1},"PASS")={len(CHK)},"PASS","FAIL")').font = BOLD
CHECK_ALL = f"'核對'!E{kl+1}"

# =============================================================== 說明
widths(ws0, [22, 110])
ws0["A1"] = f"人形機器人單位經濟 {VERSION}（線 2b，只做硬體）"
ws0["A1"].font = TITLE
info = [
    ("產生", f"2026-11-16｜build_humanoid.py｜設計稿：{S['FW']} 第 3 節"),
    ("核心問題", "全尺寸人形的硬體成本何時降到「買一台換一個（或半個）工人」3 年內回本；價值落在整機還是關節零組件。"),
    ("分析單位", "每台全尺寸人形（Optimus Gen 2 類：16 旋轉＋14 線性關節＋2 靈巧手），2026 年基準，美元。小型機（宇樹 G1、R1）只作核對。"),
    ("頁面", "輸入 → 敏感度（承重排序，先看）→ 計算（BOM、製造成本、售價、TCO、回本）→ 情境（2026–2035 低／基準／高）→ 下游取數 → 追蹤指標 → 核對"),
    ("顏色", "藍字＝輸入；黃底＝判斷值（Assumed／Analogy）或可改的選擇；綠底＝主要結果；黑字＝公式"),
    ("標記", "[Verified] 可公開驗證；[Interested-party] 有利害關係的來源（賣方、公司、媒體轉述）；[Analogy] 類比；[Assumed] 假設；[Derived] 由其他值推得。Analogy、Assumed 一律給低／基準／高。"),
    ("類別", "[P] 物理；[A] 會計恆等式；[I] 誘因資訊"),
    ("絲杠矛盾", "行星滾柱絲杠單價有五個互相矛盾的出處：雙林 ¥499、五洲新春 2025 均價約 ¥1,360、國際品 >¥5,000、MS 1,350–2,700 美元。MS 單價 × 14 支已占 MS 自己 BOM 的 34–69%，與覓途「絲杠占 19%」不能同時成立。做法：輸入頁 G 表選情形（基準選 3），敏感度頁 C 表列出五個情形的 BOM 與回本期。"),
    ("口徑提醒", "(1) 全球出貨大多是中國小型機，不用絲杠、多用行星減速，所以全球平均每台用量＝全尺寸用量 × 採用比例；(2) 學習曲線用全產業累積出貨，零組件單價同比例下降；(3) 售價＝製造成本 ÷（1 − 毛利率），不是任何廠商的列表價。"),
    ("未取得", "無框馬達、編碼器、驅動板、軸承、線束的單價；良率；MTBF；相對人工生產力實測；出貨按用途拆分。這些都標 [Assumed] 並給區間。"),
    ("不做", "機器人智能（模型、資料、訓練算力）；晶片規格與產量（AI供應鏈只讀）；四足機器人。"),
    ("核對", "核對頁全部 PASS；check_humanoid.py 以 LibreOffice 重算後檢查 0 錯誤值。"),
]
for i, (a, b) in enumerate(info):
    ws0.cell(3 + i, 1, a).font = BOLD
    c = ws0.cell(3 + i, 2, b)
    c.alignment = WRAP
rr = 3 + len(info) + 1
ws0.cell(rr, 1, "核對結果").font = BOLD
ws0.cell(rr, 2, f"={CHECK_ALL}")
rr += 2
ws0.cell(rr, 1, "Sources").font = BOLD
for k, v in S.items():
    rr += 1
    ws0.cell(rr, 1, k)
    ws0.cell(rr, 2, v).alignment = WRAP

os.makedirs(os.path.dirname(OUT), exist_ok=True)
wb.save(OUT)


def lo_recalc(path):
    """以 LibreOffice headless 重算並寫回，讓快取值可被 data_only 讀取（下游取數用）。"""
    import shutil
    import subprocess
    import tempfile
    tmp = tempfile.mkdtemp()
    src = os.path.join(tmp, "in.xlsx")
    shutil.copy(path, src)
    subprocess.run(["soffice", "--headless", "--calc", "--convert-to", "xlsx", "--outdir", os.path.join(tmp, "out"), src],
                   check=True, capture_output=True, timeout=300)
    shutil.copy(os.path.join(tmp, "out", "in.xlsx"), path)


if "--no-recalc" not in os.sys.argv:
    lo_recalc(OUT)
print(OUT)
