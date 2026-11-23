# -*- coding: utf-8 -*-
"""Tesla（TSLA）個股模型 v0.3.1 建置腳本（AI實體應用；physical_ai/tesla/）
v0.3.1：現價換 2026-11-20 收盤（Yahoo Finance 日線，physical_ai/data/closes_20261109_20261120.csv）；其他假設不變。
v0.3：Robotaxi 讀分部 v0.3（對帳 2）；公開前清除原檔個人字樣（修改紀錄_3Q26!A53 等）；「原檔假設」標記。

以既有檔 20261028_TeslaTSLA.US_post_3Q26.xlsx（post-3Q26 版）為底改版，不從頭重建：
  * 保留原檔全部工作表（Non-GAAP 季度引擎、財報預估差異、長期評價分析、2025PA、Robotaxi 試算等），只做
    (1) 年度欄公式範圍錯誤修正、(2) 滾入 3Q26 後新資訊（Cybertruck 2027 年中停產）；
  * 新增：README、In、Hist、Optimus、PL（官方分部→損益，三情境）、DCF（標準 FCFF）、SOTP（沿用原檔 2035 年分部倍數法）、
    Val、Sens、Recon、對帳、風險登錄、改版紀錄_v0.2。
v0.2：Robotaxi 分部價值（SOTP）與 FY2027–FY2035 Robotaxi 營收（PL／DCF 腿）改讀 physical_ai/tesla_robotaxi 分部 v0.3 成品
（具名儲存格 RT_VAL_A/B/C、RT_REV_<情境>_<年>；建置時從分部成品讀值，不手打），取代暫定值與「原檔第 131 列 × rt_k」。
用法：
  python3 physical_ai/tesla/build_tesla.py [--base <原檔路徑>] [--nosens]
  產出：physical_ai/tesla/dist/20261116_Tesla個股模型_v0.2.xlsx（LibreOffice 重算後的成品，含公式與快取值）
"""
import glob, hashlib, os, re, shutil, subprocess, sys, tempfile, json
from concurrent.futures import ThreadPoolExecutor
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.comments import Comment

HERE = os.path.dirname(os.path.abspath(__file__))
VER = "v0.3.1"
DATE_TAG = "20261123"
OUT_NAME = f"{DATE_TAG}_Tesla個股模型_{VER}.xlsx"
BASE = "/mnt/user-data/uploads/03-AI實體應用/20261028_TeslaTSLA.US_post_3Q26.xlsx"
args = sys.argv[1:]
if "--base" in args:
    BASE = args[args.index("--base") + 1]
DO_SENS = "--nosens" not in args

# ------------------------------------------------------------------ 讀 Robotaxi 分部 v0.3 成品（權威：physical_ai/tesla_robotaxi/dist）
SEG_GLOB = os.path.join(HERE, "..", "tesla_robotaxi", "dist", "*TeslaRobotaxi個股模型_v0.3.xlsx")


def load_seg():
    path = sorted(glob.glob(SEG_GLOB))[-1]
    fn = os.path.basename(path)
    wb = openpyxl.load_workbook(path, data_only=True)
    out = {"file": fn, "path": path, "val": {}, "rev": {}, "cell": {}}

    def get(nm):
        sh, ref = list(wb.defined_names[nm].destinations)[0]
        v = wb[sh][ref.replace("$", "")].value
        if not isinstance(v, (int, float)):
            raise SystemExit(f"{fn} 的 {nm} 沒有快取值（{v}）：請先對分部成品跑 recalc.py")
        return v, f"{sh}!{ref.replace('$', '')}"
    for sc in ("A", "B", "C"):
        out["val"][sc], out["cell"][f"RT_VAL_{sc}"] = get(f"RT_VAL_{sc}")
        for y in range(2027, 2036):
            out["rev"][(sc, y)], out["cell"][f"RT_REV_{sc}_{y}"] = get(f"RT_REV_{sc}_{y}")
    return out


SEG = load_seg()

BLUE = Font(color="0000FF")
BOLD = Font(bold=True)
HDR = PatternFill("solid", fgColor="DDEBF7")
SEC = PatternFill("solid", fgColor="FFF2CC")
WARN = PatternFill("solid", fgColor="F8CBAD")
NG = "'Non-GAAP'!"

YEARS = list(range(2025, 2036))                      # PL 欄：C..M
PLC = {y: openpyxl.utils.get_column_letter(3 + i) for i, y in enumerate(YEARS)}
ENG = {y: openpyxl.utils.get_column_letter(87 + (y - 2025)) for y in YEARS}   # 原檔年度欄 CI=FY25 … CS=FY35
DY = list(range(2027, 2036))                          # DCF 明確期 FY2027–FY2035；C 欄＝4Q26 stub
DCC = {y: openpyxl.utils.get_column_letter(4 + i) for i, y in enumerate(DY)}
SCN = ["A", "B", "C"]
SCN_L = SCN

# ------------------------------------------------------------------ 來源字串
S_XBRL = "SEC XBRL companyfacts（CIK 1318605）：10-Q 2026-09-30（accn 0001628280-26-066775，2026-10-22 申報）https://data.sec.gov/api/xbrl/companyfacts/CIK0001318605.json"
S_10K = "SEC XBRL companyfacts：10-K FY2025（accn 0001628280-26-003952，2026-01-29 申報）"
S_Q3 = "Tesla Q3 2026 Update（8-K Ex.99.1，2026-10-21），經 Electrek https://electrek.co/2026/10/21/tesla-tsla-q3-2026-financial-results/ 與 Drive Tesla Canada 轉述"
S_CALL = "Tesla Q3 2026 法說（2026-10-21），經 Not a Tesla App 整理 https://notateslaapp.com/news/4867/tesla-q3-2026-earnings-call-recap-everything-you-need-to-know"
S_ORIG = "原檔 20261028_TeslaTSLA.US_post_3Q26.xlsx（post-3Q26 版）"
S_2B = "線框架 2b 人形機器人 v0.1（2026-11-16）第 3 節"
S_1A = "線框架 1a Robotaxi v0.1（2026-11-16）第 3 節"
S_1B = "線框架 1b 智慧車 v0.1（2026-11-16）第 3 節"
INC_CO = "公司自揭（管理層有維持估值與股價的誘因；CEO 獎酬與市值、營運里程碑掛鉤）"
INC_SELL = "賣方或媒體轉述賣方估計：交易與投行業務誘因"
INC_MUSK = "CEO（大股東，2025 CEO 獎酬與市值、100 萬台機器人等里程碑掛鉤）"
INC_ORIG = "原檔假設（非外部來源）"

# ------------------------------------------------------------------ In 頁列
# (鍵, 項目, 單位, 低, 基準, 高, 標記, 來源1, 來源2／說明, 來源誘因)
IN = []
def sec(t): IN.append(("sec", t))
def row(*a): IN.append(a)

sec("A. 市場資料與日期（資料基準日＝股價日 2026-11-20）")
row("px", "股價（2026-11-20 收盤）", "US$", None, 345.82, None, "Verified",
    "Yahoo Finance 日線 2026-11-20 收盤（physical_ai/data/closes_20261109_20261120.csv）：345.82",
    "前版 v0.3 股價：2026-11-13 收盤 390.22（stockanalysis.com，2026-11-14 擷取），只作對照", None)
row("px_date", "股價日（資料基準日 as_of）", "日期", None, "=DATE(2026,11,20)", None, "Verified", "同上", "原則 15、26、41", None)
row("tp_date", "目標價時點（＝EDATE(px_date,12)）", "日期", None, "=EDATE(px_date,12)", None, "Derived", "原則 26", None, None)
row("cons_tp", "市場共識目標價（44 位分析師）", "US$", None, 412.48, None, "Interested-party",
    "stockanalysis.com 2026-11-14 擷取：共識 Buy、目標價 412.48", "只作對照，不進模型；對照值未更新（維持 2026-11-14）", INC_SELL)

sec("B. 折現率（DCF 規格第 2 節；rf 與其他個股模型同一取值日）")
row("rf", "無風險利率（美國 10 年期公債殖利率，2026-10-23）", "%", None, 0.05, None, "Verified",
    "美國財政部每日殖利率曲線 10 Yr 2026-10-23 5.00%（與 nvidia/ 等個股模型同一取值；跨模型一致，原則 15）",
    "取值日 2026-10-23 早於股價日 2026-11-20，逾期；下次全面改版一併更新", None)
row("beta_y", "β 來源 1：Yahoo Finance Beta (5Y Monthly)", "x", None, 1.92, None, "Verified",
    "https://finance.yahoo.com/quote/TSLA/key-statistics/（2026-11-16 擷取）", "原始值，未調整", None)
row("beta_s", "β 來源 2：stockanalysis Beta (5Y)", "x", None, 1.74, None, "Verified",
    "https://stockanalysis.com/stocks/tsla/statistics/（2026-11-14 更新）", "原始值，未調整", None)
row("beta", "β（兩來源 Blume 調整 0.67β＋0.33 後中位數）", "x", "=MEDIAN(0.67*beta_y+0.33,0.67*beta_s+0.33)-0.3",
    "=MEDIAN(0.67*beta_y+0.33,0.67*beta_s+0.33)", "=MEDIAN(0.67*beta_y+0.33,0.67*beta_s+0.33)+0.3", "Derived",
    "DCF 規格第 2 節：β 採兩來源 Blume 調整後中位數", "Sens：±0.3（Assumed 區間）；未做自算產業 β（汽車／機器人同業組未建）", None)
row("erp", "股權風險溢酬 ERP", "%", 0.04, 0.05, 0.06, "Assumed", "DCF 規格第 2 節預設 5%（敏感度 4–6%）", "Sens：低／高", None)
row("kd", "稅前債務成本 kd", "%", None, 0.06, None, "Assumed",
    "Tesla 債務多為無追索 ABS 與中國營運資金額度（sec-api.io 10-Q 分析：中國額度 60 億、汽車 ABS 27 億、能源 ABS 7 億）；以 rf＋1pt 近似",
    "負債占資本不到 1%，對 WACC 影響可忽略；未查核各批利率", None)
row("tax_w", "WACC 用稅率（債務稅盾）", "%", None, 0.25, None, "Assumed", "沿用原檔預估期稅率 25%", "只用於債務稅盾", INC_ORIG)
row("wacc_adj", "WACC 調整（Sens 用；基準 0）", "%", -0.01, 0, 0.01, "Assumed", "DCF 規格第 4 節：WACC ±1pt", "—", None)
row("g_cap", "永續成長率上限", "%", None, 0.03, None, "Assumed", "DCF 規格第 3 節：g 預設 3%", "g＝MIN(上限, rf)", None)
row("g_tv", "永續成長率 g（FY2035 後）", "%", "=MIN(g_cap,rf)-0.01", "=MIN(g_cap,rf)", "=MIN(g_cap,rf)+0.01", "Derived",
    "DCF 規格第 3 節：3%，敏感度 2–4%", "Sens：−1pt／＋1pt", None)
row("ronic_sp", "RONIC 相對 WACC 的差（RONIC＝WACC＋本值）", "%", None, 0, 0.05, "Assumed", "DCF 規格第 3 節：RONIC＝WACC（價值中性）", "Sens：WACC＋5pt", None)
row("tv_alt", "終值算法開關（0＝NOPAT×(1−g/RONIC)；1＝FCFF×(1＋g) 替代法）", "0/1", None, 0, 1, "Assumed",
    "DCF 規格第 4 節：敏感度須含終值替代法", "Sens：高＝替代法", None)

sec("C. 營運：沿用原檔 Non-GAAP 季度引擎（汽車量價、FSD、Robotaxi、能源、費用），本頁只放 FY2027 起的調整與情境係數")
row("tax_t", "預估期稅率（FY2027 起）", "%", 0.20, 0.25, 0.25, "Assumed",
    "原檔預估期 25%（原檔待決定第 4 項：維持 25% 或改近四季平均約 20%）", "XBRL：9M26 有效稅率 20%、FY2025 27%；低＝20%", INC_ORIG)
row("auto_gm_adj", "非 FSD 汽車毛利率調整（pt，加在原檔路徑上，FY2027 起）", "pt", -0.03, 0, 0.03, "Assumed",
    "原檔路徑：3Q26 經常性 16.5% 起每季 +0.1pt（原檔待決定第 1 項，預設依 3Q26 實際校準）",
    "區間：原檔前版起點 20.4% 與 3Q26 實際之差約 3–4pt", INC_ORIG)
row("fsd_k", "FSD 營收倍數（相對原檔路徑，FY2027 起）", "x", 0.6, 1.0, 1.4, "Assumed",
    "原檔路徑：活躍數＝累積交車×滲透率（FY27 20% → FY35 60%）；買斷存量 87 萬不付月費",
    "低：EU 未過、滲透停滯；高：EU 通過且無監督 FSD 2027Q2 上路（法說）", INC_ORIG)
row("fsd_gm", "FSD 毛利率（調整額用）", "%", None, 0.80, None, "Assumed", "原檔 Non-GAAP 第 169 列 FSD 毛利率約 80%", "只用於 fsd_k、rt_k 與分部 v0.3 營收相對原檔路徑之差的毛利調整額", INC_ORIG)
row("rt_gm", "Robotaxi 毛利率（調整額用）", "%", None, 0.55, None, "Assumed", "原檔 Non-GAAP 第 170 列 Robotaxi 毛利率約 55%", "同上", INC_ORIG)
row("rt_k", "Robotaxi 營收倍數（DCF 腿；只在 Sens 動，基準 1.0；乘在分部 v0.3 的 RT_REV 上）", "x", 0.5, 1.0, 1.5, "Assumed",
    "v0.2：Robotaxi 營收路徑讀分部 v0.3（見 G2 節），本倍數只供敏感度", "區間；SOTP 的 Robotaxi 價值不乘此倍數", INC_ORIG)
row("sbc_r", "股票薪酬 ÷ 營收（只用於收支關卡的自由現金流）", "%", 0.03, 0.0425, 0.06, "Derived",
    "XBRL：9M26 股票薪酬 3,346 ÷ 9M26 營收 78,669＝4.25%", "DCF 以 NOPAT 扣股票薪酬，不加回（規格第 1 節）", None)

sec("D. Optimus（讀線框架 2b；本模型為唯一出處，Tesla 未揭露任何 Optimus 數字）")
OPT_U = {2027: 2000, 2028: 20000, 2029: 80000, 2030: 200000, 2031: 350000, 2032: 550000, 2033: 800000, 2034: 1050000, 2035: 1300000}
for y, u in OPT_U.items():
    row(f"opt_u_{y}", f"Optimus 對外交付量（基準路徑）FY{y}", "台", round(u * 0.25), u, u * 2, "Assumed",
        "法說 2026-10-21：對外交付約 2027 年底開始、先工業客戶；Fremont 產線設計 100 萬台/年、達每週 2 萬台至少一年（" + S_2B + "）",
        "低＝×0.25、高＝×2（整組一案，Sens「Optimus 量路徑」）", INC_MUSK)
row("opt_k_A", "情境 A：Optimus 量倍數（乘基準路徑）", "x", None, 1.6, None, "Assumed", "情境 A：量產節奏接近公司說法", "—", None)
row("opt_k_B", "情境 B：Optimus 量倍數", "x", None, 1.0, None, "Assumed", "基準路徑", "—", None)
row("opt_k_C", "情境 C：Optimus 量倍數", "x", None, 0.25, None, "Assumed", "情境 C：壽命與可靠度未在 2027 年底解決（2b 第 4 節門檻 3）", "—", None)
row("opt_asp27", "Optimus 平均售價 FY2027", "千美元/台", 30, 40, 60, "Assumed",
    "Musk：初期成本約 3 萬美元、往 2 萬下降（Interested-party，2b 第 3.2 節）；1X NEO 2 萬美元（家用對照）", "售價高於成本說法；低高為區間", INC_MUSK)
row("opt_asp35", "Optimus 平均售價 FY2035（FY2027 起線性）", "千美元/台", 18, 25, 35, "Assumed",
    "BofA：2030–35 年 BOM 1.3–1.7 萬美元（2b）；售價取成本加毛利", "—", INC_SELL)
row("opt_bom0", "Optimus 每台 BOM（累積產量 opt_q0 時）", "千美元/台", 35, 55, 65, "Interested-party",
    "Morgan Stanley：Optimus Gen 2 類 BOM 5–6 萬美元（二手轉述，2b 第 3.2 節）", "低：BofA 2025 年底業界 BOM 3.5 萬美元；高：MS 上緣", INC_SELL)
row("opt_q0", "學習曲線起點累積產量 Q0", "台", None, 1000, None, "Assumed", "BOM₀ 對應的累積產量（試產規模）", "—", None)
row("opt_lr", "學習率（累積產量每倍增成本降幅）", "%", 0.10, 0.15, 0.25, "Analogy",
    "類比工業機器人與車用電池成本曲線 10–25%（2b 第 3.3 節）", "—", None)
row("opt_oh", "製造附加（組裝測試、良率損失、折舊）÷ BOM", "%", 0.10, 0.20, 0.40, "Assumed", "良率未揭露（2b 第 7 節）", "—", None)
row("opt_sga", "Optimus 銷售與服務費用 ÷ 營收", "%", 0.05, 0.08, 0.12, "Analogy", "類比汽車業 SG&A 占營收（Tesla FY2025 6.2%）", "研發已在原檔引擎研發費用內，不重複", None)
row("opt_cx", "Optimus 每台新增年產能資本支出", "千美元/台", 3, 5, 10, "Assumed",
    "Fremont 100 萬台/年產線（法說）；投資額未揭露", "資本支出＝年交付增量 × 本值", None)
row("opt_wage", "美國工人全負擔時薪（回本期展示用）", "US$/小時", None, 46.89, None, "Verified", "BLS ECEC 2026-06（2b 第 3.3 節）", "只用於 Optimus 頁回本期，不進目標價", None)
row("opt_prod", "相對人工生產力（回本期展示用）", "x", 0.3, 0.5, 1.0, "Assumed", "無公開實測（2b 第 4 節）", "只用於回本期展示，不進目標價", None)

sec("E. 現金流（DCF 與收支關卡）")
row("cx_26", "FY2026 資本支出", "US$M", 19000, 20000, 22000, "Interested-party",
    "法說 2026-10-21：2026 資本支出修正為約 200 億美元（Q2 時為 >250 億）", "XBRL：9M26 資本支出 13,014；10-Q 稱 in excess of $20bn（sec-api.io 轉述）", INC_CO)
row("cx_r27", "FY2027 核心資本支出 ÷ 核心營收（不含 Optimus）", "%", 0.18, 0.25, 0.32, "Assumed",
    "FY2026 約 200 億 ÷ 原檔 FY26 營收 1,053 億＝19%；Terafab 設備數百億美元至 2028 年初（法說），2027 指引未取得", "FY2027 起線性收斂到 cx_ss（FY2035）", None)
row("cx_ss", "FY2035 核心資本支出 ÷ 核心營收（穩態）", "%", 0.06, 0.08, 0.10, "Analogy", "類比 Tesla FY2019–FY2023 資本支出占營收 5–8%（10-K）", "—", None)
row("life", "PP&E 平均折舊年限（折舊＝期初淨額 ÷ 年限）", "年", 6, 8, 10, "Analogy",
    "XBRL：FY2025 折舊 5,030 ÷ 期初 PP&E 淨額 35,836＝14%（約 7 年）", "會計口徑；折舊滾動起點＝2026-09-30 PP&E 淨額", None)
row("wc_r", "營運資金增加 ÷ 營收增量", "%", -0.05, 0, 0.10, "Assumed",
    "XBRL：FY2024、FY2025 營運資產負債變動皆為淨流入（應付帳款增加大於存貨與應收增加）；比率無穩定意義，基準 0", "原則 23：不省略營運資金，以 Sens 呈現", None)
row("tail_tfab", "尾端：Terafab 額外資本支出（FY2027、FY2028 各年；基準 0）", "US$M", None, 0, 15000, "Assumed",
    "法說：Terafab 設備數百億美元至 2028 年初、長期可能逾千億；與 SpaceX 分攤未定", "原則 19：事件型尾端風險進 Sens、不進基準", None)

sec("F. 股數與或有股份（DCF 規格第 3 節；原則 37）")
row("ceo_ovr", "2025 CEO 獎酬歸屬比例覆寫（−1＝依市值閘門；0／1＝Sens）", "x", 0, -1, 1, "Assumed",
    "12 批、各約 3,531 萬股；市值閘門 2 兆起每 0.5 兆一批至 6.5 兆，再 7.5、8.5 兆；另需營運里程碑（Meridian 2025-11 整理股東會資料）",
    "基準只依市值閘門（以本情境股權價值判定，營運里程碑視為同步達成）；Sens 0%／100%；每股價值低於抵付價 334.09 時淨增股數為 0", None)

sec("G. 分部加總（SOTP；沿用原檔「長期評價分析」2035 年倍數法）")
row("sotp_r", "SOTP 折現率（原檔 15%）", "%", 0.12, 0.15, 0.18, "Assumed", "原檔長期評價分析 Z9＝1.15^10、Z11＝×1.15", "—", INC_ORIG)
row("sotp_n", "SOTP 折現年數（原檔 10 年；原檔待決定第 3 項）", "年", 9, 10, None, "Assumed", "原檔預設 10 年；以 2026 為基期可改 9 年", "—", INC_ORIG)
row("pe_auto", "車輛業務（非 FSD／Robotaxi）2035 年本益比", "x", 8, 12, 16, "Assumed", "原檔 Y2＝12", "未逐家取得同業倍數（原則 7 未做）", INC_ORIG)
row("npm_auto", "車輛業務 2035 年淨利率", "%", 0.08, 0.12, 0.14, "Assumed", "原檔 W2＝營收×12%", "—", INC_ORIG)
row("pe_non", "非汽車業務（能源＋服務）2035 年本益比", "x", 10, 15, 20, "Assumed", "原檔 Y5＝15", "—", INC_ORIG)
row("npm_non", "非汽車業務 2035 年淨利率", "%", 0.10, 0.15, 0.18, "Assumed", "原檔 W5＝營收×15%", "—", INC_ORIG)
row("peg_k", "成長分部市銷率係數（P/S＝FY2035 營收年增率×100×本值）", "x", 0.5, 0.8, 1.0, "Assumed",
    "原檔 X3、X4＝年增率×100×0.8（FSD、Robotaxi）；本版同法用於 Optimus", "—", INC_ORIG)
for _s in SCN_L:
    row(f"rt_val_{_s}", f"情境 {_s}：Robotaxi 分部價值（現值，讀分部 v0.3 RT_VAL_{_s}）", "US$B", None, SEG["val"][_s], None, "Derived",
        f"{SEG['file']}｜{SEG['cell']['RT_VAL_' + _s]}（具名儲存格 RT_VAL_{_s}；建置時讀值）", "分部 EV（不含 Tesla 淨現金）；線 1a v0.2 路徑，Assumed 為主", None)
sec("G2. Robotaxi 營收路徑（讀分部 v0.3；Derived；取代原檔第 131 列 × rt_k；FY2026 以前維持原檔）")
for _y in range(2027, 2036):
    for _s in SCN_L:
        row(f"rt_rev_{_y}_{_s}", f"情境 {_s}：Tesla Robotaxi 營收 FY{_y}（分部 v0.3）", "US$M", None, SEG["rev"][(_s, _y)], None, "Derived",
            f"{SEG['file']}｜{SEG['cell'][f'RT_REV_{_s}_{_y}']}（具名儲存格 RT_REV_{_s}_{_y}，百萬美元；建置時讀值）",
            "自有車隊全額＋車主車輛抽成；成本與毛利不在此，沿用本模型 rt_gm（待對帳）", None)

sec("H. 權重")
row("tp_wA", "情境權重 A", "%", None, 0.25, None, "Assumed", "DCF 規格第 3 節預設 25／50／25", "—", None)
row("tp_wB", "情境權重 B", "%", None, 0.50, None, "Assumed", "同上", "—", None)
row("tp_wC", "情境權重 C", "%", None, 0.25, None, "Assumed", "同上", "—", None)
row("w_dcf", "DCF 權重（其餘為分部加總）", "%", 0.3, 0.5, 0.7, "Assumed", "個股模型原則第 4 節：DCF 50%＋分部加總 50%", "Sens：0.3／0.7", None)

SCN_DRV = ([("opt_k", "Optimus 量倍數")] + [(f"rt_rev_{_y}", f"Robotaxi 營收 FY{_y}（分部 v0.3，US$M）") for _y in range(2027, 2036)]
           + [("rt_val", "Robotaxi 分部價值（SOTP，讀分部 v0.3，US$B）")])

# ------------------------------------------------------------------ 工具
def dn(wb, name, ref):
    wb.defined_names[name] = DefinedName(name, attr_text=ref)

def hdr(ws, r, labels, c0=1):
    for j, h in enumerate(labels):
        c = ws.cell(r, c0 + j, h); c.font = BOLD; c.fill = HDR

def setw(ws, widths):
    for k, v in widths.items():
        ws.column_dimensions[k].width = v


def _personal_terms(wb):
    """個人字樣清單由原檔文件屬性（作者）推得，不在公開程式碼寫字面姓名（v0.3）"""
    import re as _re
    terms = set()
    for v in (wb.properties.creator, wb.properties.lastModifiedBy):
        if v:
            terms.add(v.strip())
            terms.update(t for t in _re.split(r"[()（）\s,，/]+", v) if len(t) >= 2)
    return sorted(terms, key=len, reverse=True)


def scrub(wb):
    """原檔中的個人字樣（儲存格、註解、文件屬性）改為中性用語；回傳改動清單（寫入改版紀錄）"""
    terms = _personal_terms(wb)
    hits = []

    def clean(v):
        # 試算表匯出的註解標頭「ID#… ⏎ 姓名 (日期時間)」：姓名可能與文件屬性的寫法不同（連字號、分段），一律改中性用語
        v = re.sub(r"(ID#\S+\s*\n)[^\n(]*\(", r"\1原檔作者 (", v)
        for t in terms:
            v = v.replace(f"待 {t} 決定", "原檔待決定事項").replace(f"待{t}決定", "原檔待決定事項")
        for t in terms:
            v = v.replace(t, "原檔作者")
        # 姓名其餘部分（緊貼在已替換字樣前的拉丁字名，例如連字號寫法）一併移除
        v = re.sub(r"[A-Za-z][A-Za-z\-]{1,30}[ \u00a0]+原檔作者", "原檔作者", v)
        return v
    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for c in row:
                if isinstance(c.value, str) and not c.value.startswith("="):
                    v = clean(c.value)
                    if v != c.value:
                        hits.append(f"{ws.title}!{c.coordinate}"); c.value = v
                if c.comment is not None:
                    v = clean(c.comment.text)
                    if v != c.comment.text:
                        hits.append(f"{ws.title}!{c.coordinate}（註解）")
                        c.comment = Comment(v, "原檔作者" if c.comment.author and any(t in c.comment.author for t in terms) else (c.comment.author or ""))
                    elif c.comment.author and any(t in c.comment.author for t in terms):
                        c.comment = Comment(c.comment.text, "原檔作者")
    for attr in ("creator", "lastModifiedBy"):
        if getattr(wb.properties, attr):
            setattr(wb.properties, attr, "company-models"); hits.append(f"文件屬性 {attr}")
    return hits


def build(path_out):
    wb = openpyxl.load_workbook(BASE)
    SCRUBBED = scrub(wb)
    print("個人字樣改中性：", SCRUBBED)
    ng = wb["Non-GAAP"]
    # ---------------- (1) 原檔年度欄公式範圍修正（只影響年度顯示欄，季度數值不變）
    fixes = []
    for cell, f in [("CI34", "=AVERAGE(AD34:AG34)"), ("CI51", "=SUM(AD51:AG51)"), ("CI52", "=SUM(AD52:AG52)"),
                    ("CI147", "=AVERAGE(AD147:AG147)"), ("CI149", "=AVERAGE(AD149:AG149)"),
                    ("CJ78", "=SUM(AH78:AK78)"), ("CK78", "=SUM(AL78:AO78)"), ("CL78", "=SUM(AP78:AS78)"), ("CM78", "=SUM(AT78:AW78)")]:
        fixes.append((cell, ng[cell].value, f)); ng[cell] = f
    # ---------------- (2) 滾入 3Q26 後新資訊：Cybertruck 2027 年中停產（Musk X 2026-11-06）
    for cell in ("AN86", "AO86"):
        fixes.append((cell, ng[cell].value, -1)); ng[cell] = -1
        ng[cell].comment = Comment("v0.1（2026-11-16）：Musk 2026-11-06 於 X 表示 Giga Texas 約 2027 年中停止 Cybertruck 生產、改產 Cybercab（heygotrade 2026-11-09 轉述；Tesla 未發新聞稿）[Interested-party]。原值 0（沿用 3Q26 估計）。", "v0.1")
    # Cybertruck 停產後同期為 0，年增率格除以零：包 IFERROR（只影響顯示列 86、124）
    for rr in (86, 124):
        for cc in range(2, ng.max_column + 1):
            v = ng.cell(rr, cc).value
            if isinstance(v, str) and v.startswith("=") and "/" in v and "IFERROR" not in v and cc >= 44:
                ng.cell(rr, cc).value = f'=IFERROR({v[1:]},"-")'
    lt = wb["長期評價分析"]
    lt["AC1"] = "舊法（原檔 2035 年分部倍數、折現 15%）：參考，不進結論。新版 SOTP 頁沿用其結構，並修正 V5 誤用 FY2030（N 欄）營收的問題。"
    lt["AC1"].font = Font(color="C00000", bold=True)

    # ---------------- In
    I = wb.create_sheet("In", 0)
    I["A1"] = ("In — 輸入（藍字；基準欄 E 進主模型，D／F＝敏感度低／高，只用於 Sens 頁（一次動一個參數）、不進目標價；情境 A／B／C 的驅動見本頁最後一節「情境驅動對照」）。"
               "金額：百萬美元（US$M，與原檔引擎同）；分部價值另註 US$B。")
    I["A2"] = "標記：每列單一主要標記 Verified／Interested-party／Analogy／Assumed／Derived；Analogy、Assumed 給低／高區間；Interested-party 見 J 欄來源誘因。實體 AI 模型不讀 Siliconomics／Tokenomics／Wattonomics 協調檔（算力線），規格中 coord 部分略過。"
    hdr(I, 3, ["鍵", "項目", "單位", "敏感度低", "基準", "敏感度高", "標記", "來源1", "來源2／說明", "來源誘因"])
    keyrow = {}
    r = 4
    for it in IN:
        if it[0] == "sec":
            I.cell(r, 1, it[1]).font = BOLD; I.cell(r, 1).fill = SEC; r += 1; continue
        k, lab, unit, lo, base, hi, tag, s1, s2, inc = it
        I.cell(r, 1, k); I.cell(r, 2, lab); I.cell(r, 3, unit)
        for c, v in ((4, lo), (5, base), (6, hi)):
            if v is not None:
                cc = I.cell(r, c, v)
                if not (isinstance(v, str) and v.startswith("=")): cc.font = BLUE
                if unit == "%": cc.number_format = "0.0%"
                elif unit == "日期": cc.number_format = "yyyy-mm-dd"
                elif isinstance(v, (int, float)) and abs(v) >= 1000: cc.number_format = "#,##0"
        I.cell(r, 7, tag); I.cell(r, 8, s1); I.cell(r, 9, s2); I.cell(r, 10, inc)
        keyrow[k] = r
        dn(wb, k, f"In!$E${r}")
        r += 1
    IN_LAST = r - 1
    # 情境驅動對照
    r += 1
    I.cell(r, 1, "Z. 情境驅動對照（純展示：三情境並排，每格引用上方輸入列；修改請改上方藍字。本節不得被任何公式引用）").font = BOLD
    I.cell(r, 1).fill = SEC; r += 1
    hdr(I, r, ["鍵", "情境驅動", "單位", "情境 A（樂觀）", "情境 B（基準）", "情境 C（保守）", "", "說明"]); r += 1
    for k, lab in SCN_DRV:
        I.cell(r, 1, k); I.cell(r, 2, lab); I.cell(r, 3, I.cell(keyrow[k + "_B"], 3).value)
        for j, s in enumerate(SCN):
            I.cell(r, 4 + j, f"=E{keyrow[k + '_' + s]}")
        I.cell(r, 8, f"引用 In 頁 {k}_A／{k}_B／{k}_C；三情境的唯一差異來源之一")
        r += 1
    I.cell(r, 1, "scn_w"); I.cell(r, 2, "情境權重"); I.cell(r, 3, "%")
    for j, s in enumerate(SCN):
        I.cell(r, 4 + j, f"=E{keyrow['tp_w' + s]}").number_format = "0%"
    I.cell(r, 8, "引用 In 頁 tp_wA／tp_wB／tp_wC；合計須＝1"); r += 2
    I.cell(r, 2, "說明：其他輸入（稅率、毛利率調整、FSD 倍數、Optimus 單價與成本、資本支出、折現率）三情境相同——無經濟機制或證據支持分化（原則 46）；情境差異只來自 Robotaxi 與 Optimus 的規模。")
    setw(I, {"A": 14, "B": 52, "C": 10, "D": 11, "E": 11, "F": 11, "G": 15, "H": 70, "I": 50, "J": 30})

    # ---------------- Hist
    H = wb.create_sheet("Hist", 1)
    H["A1"] = "Hist — 歷史財務與評價用資產負債（原則 14、34：法定申報為主來源；公式只引用本頁定義名稱 h_*）。單位：US$M、百萬股。"
    hdr(H, 3, ["鍵", "項目", "值", "單位", "期間", "標記", "來源"])
    HI = [
        ("h_rev25", "營收 FY2025", 94827, "US$M", "FY2025", S_10K + " Revenues"),
        ("h_ocf25", "營業現金流 FY2025", 14747, "US$M", "FY2025", S_10K),
        ("h_cx25", "資本支出 FY2025", 8527, "US$M", "FY2025", S_10K + " PaymentsToAcquirePropertyPlantAndEquipment"),
        ("h_da25", "折舊 FY2025", 5030, "US$M", "FY2025", S_10K + " us-gaap:Depreciation（不含營業租賃車輛折舊口徑待核）"),
        ("h_sbc25", "股票薪酬 FY2025", 2825, "US$M", "FY2025", S_10K + " ShareBasedCompensation"),
        ("h_rev9m", "營收 9M26", 78669, "US$M", "2026-01-01～09-30", S_XBRL),
        ("h_ocf9m", "營業現金流 9M26", 14897, "US$M", "2026-01-01～09-30", S_XBRL),
        ("h_cx9m", "資本支出 9M26", 13014, "US$M", "2026-01-01～09-30", S_XBRL),
        ("h_sbc9m", "股票薪酬 9M26", 3346, "US$M", "2026-01-01～09-30", S_XBRL),
        ("h_da9m", "折舊 9M26", 4290, "US$M", "2026-01-01～09-30", S_XBRL + " us-gaap:Depreciation"),
        ("h_da_q3", "折舊 3Q26", 1580, "US$M", "2026-07-01～09-30", S_XBRL),
        ("h_cash", "現金及約當現金", 17853, "US$M", "2026-09-30", S_XBRL),
        ("h_sti", "短期投資", 27594, "US$M", "2026-09-30", S_XBRL),
        ("h_debt", "債務（DebtInstrumentCarryingAmount）", 9284, "US$M", "2026-09-30", S_XBRL + "；sec-api.io 轉述本金 9.3bn"),
        ("h_fl", "融資租賃負債（流動 76＋非流動 180）", 256, "US$M", "2026-09-30", S_XBRL),
        ("h_mi", "非控制權益（含可贖回 54）", 600, "US$M", "2026-09-30", S_XBRL + " MinorityInterest 546＋RedeemableNoncontrollingInterest 54"),
        ("h_spx", "SpaceX（含 xAI）股權投資公允價值", 2865, "US$M", "2026-09-30", S_XBRL + " EquityMethodInvestmentsFairValueDisclosure；投入 2,002（9M26）"),
        ("h_btc", "比特幣公允價值", 962, "US$M", "2026-09-30", S_XBRL + " CryptoAssetFairValue"),
        ("h_ppe", "PP&E 淨額（含融資租賃使用權）", 50591, "US$M", "2026-09-30", S_XBRL),
        ("h_sh_dil", "稀釋加權股數 3Q26", 3541, "百萬股", "2026-07-01～09-30", S_XBRL + " WeightedAverageNumberOfDilutedSharesOutstanding"),
        ("h_sh_out", "流通股數（封面，含 2025 CEO 獎酬未賺得限制股）", 3952.758, "百萬股", "2026-10-15", S_XBRL + " dei:EntityCommonStockSharesOutstanding"),
        ("h_ceo_sh", "2025 CEO 獎酬限制股（未賺得）", 423.744, "百萬股", "2025-11 核准", "Meridian Compensation Partners 2025-11 整理股東會資料：423,743,904 股、12 批"),
        ("h_ceo_off", "CEO 獎酬每股抵付金額（歸屬時支付，可以股抵付）", 334.09, "US$/股", "—", "同上：offset amount 334.09 美元/股"),
    ]
    hrow = {}
    r = 4
    for k, lab, v, u, per, src in HI:
        H.cell(r, 1, k); H.cell(r, 2, lab); c = H.cell(r, 3, v); c.font = BLUE; c.number_format = "#,##0.0" if v < 1000 else "#,##0"
        H.cell(r, 4, u); H.cell(r, 5, per); H.cell(r, 6, "Verified"); H.cell(r, 7, src)
        dn(wb, k, f"Hist!$C${r}"); hrow[k] = r; r += 1
    r += 1
    H.cell(r, 1, "2025 CEO 獎酬市值閘門（兆美元；12 批）").font = BOLD; r += 1
    g0 = r
    for i, g in enumerate([2, 2.5, 3, 3.5, 4, 4.5, 5, 5.5, 6, 6.5, 7.5, 8.5]):
        H.cell(r, 1, f"第 {i + 1} 批"); c = H.cell(r, 3, g); c.font = BLUE; H.cell(r, 4, "US$T")
        H.cell(r, 6, "Verified" if i in (0, 6, 11) else "Derived")
        H.cell(r, 7, "Meridian 2025-11：首批 2 兆、第 7 批 5 兆、最高 8.5 兆" + ("" if i in (0, 6, 11) else "；中間批次依每批 +0.5 兆（至 6.5 兆）再 7.5、8.5 兆推得"))
        r += 1
    dn(wb, "h_gate", f"Hist!$C${g0}:$C${r - 1}")
    r += 1
    H.cell(r, 1, "分部營收歷史（官方分部：汽車＝汽車營收＋服務及其他；能源＝能源發電與儲存）").font = BOLD; r += 1
    hdr(H, r, ["期間", "汽車營收（含碳權、租賃）", "服務及其他", "汽車分部", "能源分部", "合併營收", "檢查：分部合計 − 合併營收", "標記", "來源", "對原檔引擎差"]); r += 1
    QH = [("4Q25", "AG", 17693, 3371, 3837.2, 24901.2), ("1Q26", "AH", 16234, 3745, 2408, 22387),
          ("2Q26", "AI", 20516, 4581, 3139, 28236), ("3Q26", "AJ", 20560, 4605, 2881, 28046),
          ("FY2024", "CH", 77070, 10534, 10086, 97690), ("FY2025", "CI", 69526, 12530, 12770.69, 94826.69)]
    HSEG = r
    for per, col, au, sv, en, tot in QH:
        H.cell(r, 1, per)
        for j, v in enumerate((au, sv)):
            c = H.cell(r, 2 + j, v); c.font = BLUE; c.number_format = "#,##0"
        H.cell(r, 4, f"=B{r}+C{r}").number_format = "#,##0"
        c = H.cell(r, 5, en); c.font = BLUE; c.number_format = "#,##0"
        c = H.cell(r, 6, tot); c.font = BLUE; c.number_format = "#,##0"
        H.cell(r, 7, f"=D{r}+E{r}-F{r}").number_format = "0.0"
        H.cell(r, 8, "Verified")
        H.cell(r, 9, "Tesla 各季 Update（8-K Ex.99.1）與 10-K；原檔已逐項核對（原檔修改紀錄第一節）；9M26 合計 78,669 與 XBRL 一致")
        H.cell(r, 10, f"=F{r}-{NG}{col}46").number_format = "0.0"
        r += 1
    setw(H, {"A": 12, "B": 44, "C": 12, "D": 12, "E": 18, "F": 12, "G": 70, "H": 10, "I": 60, "J": 12})

    # ---------------- Optimus
    O = wb.create_sheet("Optimus", 2)
    O["A1"] = "Optimus — 人形機器人分部（讀線框架 2b 單位經濟：BOM 學習曲線、售價、TCO）。US$M；量為台；單價千美元。Tesla 未揭露任何 Optimus 財務數字，本頁全部為假設。"
    hdr(O, 3, ["項目", "單位"] + [f"FY{y}" for y in YEARS])
    oc = PLC  # 同 PL 欄位
    orow = {}
    r = 4
    O.cell(r, 1, "基準量路徑"); O.cell(r, 2, "台")
    for y in YEARS:
        O[f"{oc[y]}{r}"] = f"=opt_u_{y}" if y >= 2027 else 0
        O[f"{oc[y]}{r}"].number_format = "#,##0"
    orow["base"] = r; r += 1
    O.cell(r, 1, "平均售價"); O.cell(r, 2, "千美元")
    for y in YEARS:
        O[f"{oc[y]}{r}"] = f"=opt_asp27+(opt_asp35-opt_asp27)*({y}-2027)/8" if y >= 2027 else 0
        O[f"{oc[y]}{r}"].number_format = "0.0"
    orow["asp"] = r; r += 2
    for s in SCN:
        O.cell(r, 1, f"情境 {s}").font = BOLD; O.cell(r, 1).fill = SEC; r += 1
        rows = {}
        for key, lab, unit in [("u", "對外交付量", "台"), ("cum", "累積產量", "台"), ("bom", "每台 BOM（學習曲線）", "千美元"),
                               ("cost", "每台製造成本", "千美元"), ("rev", "Optimus 營收", "US$M"), ("gp", "Optimus 毛利", "US$M"),
                               ("gm", "毛利率", "%"), ("sga", "Optimus 銷售與服務費用", "US$M"), ("op", "Optimus 營業利益（不含研發）", "US$M"),
                               ("cx", "Optimus 產能資本支出", "US$M")]:
            O.cell(r, 1, lab); O.cell(r, 2, unit); rows[key] = r; r += 1
        for i, y in enumerate(YEARS):
            c = oc[y]; p = oc[y - 1] if y > 2025 else None
            O[f"{c}{rows['u']}"] = f"={c}{orow['base']}*opt_k_{s}"
            O[f"{c}{rows['cum']}"] = f"={p}{rows['cum']}+{c}{rows['u']}" if p else f"={c}{rows['u']}"
            O[f"{c}{rows['bom']}"] = f"=IF({c}{rows['u']}>0,opt_bom0*(MAX({c}{rows['cum']},opt_q0)/opt_q0)^LOG(1-opt_lr,2),0)"
            O[f"{c}{rows['cost']}"] = f"={c}{rows['bom']}*(1+opt_oh)"
            O[f"{c}{rows['rev']}"] = f"={c}{rows['u']}*{c}{orow['asp']}/1000"
            O[f"{c}{rows['gp']}"] = f"={c}{rows['u']}*({c}{orow['asp']}-{c}{rows['cost']})/1000"
            O[f"{c}{rows['gm']}"] = f"=IF({c}{rows['rev']}>0,{c}{rows['gp']}/{c}{rows['rev']},0)"
            O[f"{c}{rows['sga']}"] = f"={c}{rows['rev']}*opt_sga"
            O[f"{c}{rows['op']}"] = f"={c}{rows['gp']}-{c}{rows['sga']}"
            O[f"{c}{rows['cx']}"] = f"=MAX(0,{c}{rows['u']}-{p}{rows['u']})*opt_cx/1000" if p else "=0"
            for k in rows:
                O[f"{c}{rows[k]}"].number_format = "0.0%" if k == "gm" else ("0.0" if k in ("bom", "cost") else "#,##0")
        orow[s] = rows; r += 1
    O.cell(r, 1, "TCO 與回本期（2b 第 3.1 節；展示，不進目標價；情境 B）").font = BOLD; r += 1
    tco = {}
    for key, lab, unit in [("lab", "替代人工年成本（兩班 4,000 小時 × 時薪 × 生產力）", "千美元/年"), ("opx", "機器人年營運成本（售價 15%）", "千美元/年"),
                           ("pb", "回本期（售價×1.3 部署費 ÷ 年淨節省）", "年")]:
        O.cell(r, 1, lab); O.cell(r, 2, unit); tco[key] = r; r += 1
    for y in YEARS[2:]:
        c = oc[y]
        O[f"{c}{tco['lab']}"] = "=4000*opt_wage*opt_prod/1000"
        O[f"{c}{tco['opx']}"] = f"={c}{orow['asp']}*0.15"
        O[f"{c}{tco['pb']}"] = f"=IF({c}{tco['lab']}-{c}{tco['opx']}>0,{c}{orow['asp']}*1.3/({c}{tco['lab']}-{c}{tco['opx']}),\"N/M\")"
        for k in tco: O[f"{c}{tco[k]}"].number_format = "0.00"
    r += 1
    O.cell(r, 1, "讀法：門檻不在價格而在相對生產力與壽命（2b 第 4 節）；本頁回本期只隨 opt_prod 變動，壽命、MTBF 無公開資料。研發費用已含在原檔引擎研發列（FY27 起年增＝營收年增＋40pt），不另計。")
    setw(O, {"A": 46, "B": 10})

    # ---------------- PL
    P = wb.create_sheet("PL", 3)
    P["A1"] = "PL — 官方分部（汽車、能源）＋ Optimus（尚未列報）→ 損益 → 自由現金流。US$M。FY2025 實績、FY2026 含 9M 實績（原檔引擎）；FY2027 起依 In 頁調整與情境係數。三情境各一區塊。"
    hdr(P, 3, ["項目", "說明"] + [f"FY{y}" for y in YEARS])
    prow = {}
    r = 4
    for s in SCN:
        P.cell(r, 1, f"情境 {s}").font = BOLD; P.cell(r, 1).fill = SEC; r += 1
        L = [("seg_auto", "汽車分部營收（官方分部）", "＝以下六列"), ("veh", "  車輛銷售（非 FSD／Robotaxi）", "原檔第 119–127 列"),
             ("fsd", "  FSD", "原檔第 129 列 × fsd_k"), ("rt", "  Robotaxi", f"FY2027 起讀分部 v0.3 RT_REV_{s}_<年> × rt_k；FY2026 以前原檔第 131 列"), ("crd", "  碳權", "原檔第 133 列"),
             ("lease", "  汽車租賃", "原檔第 135 列"), ("svc", "  服務及其他", "原檔第 44 列"), ("seg_en", "能源分部營收（官方分部）", "原檔第 42 列"),
             ("seg_opt", "Optimus（尚未列報；模型新增）", "Optimus 頁"), ("rev", "合併營收", "原檔第 46 列＋調整＋Optimus"),
             ("chk", "檢查：分部合計 − 合併營收", "應為 0"), ("gp", "毛利", "原檔第 62 列＋調整＋Optimus 毛利"), ("gm", "毛利率", ""),
             ("opex", "營業費用", "原檔第 69 列＋Optimus 銷售費用"), ("ebit", "營業利益（EBIT）", ""), ("nonop", "業外", "原檔第 7 列"),
             ("pti", "稅前淨利", ""), ("tax", "所得稅", "FY2027 起 × tax_t"), ("mi", "少數股權", "原檔 3Q26 比率"),
             ("ni", "淨利（GAAP）", ""), ("sh", "稀釋股數（百萬）", "原檔第 34 列"), ("eps", "EPS（稀釋）", ""),
             ("da", "折舊", "DCF 頁 PP&E 滾動"), ("sbc", "股票薪酬（非現金）", "營收 × sbc_r"), ("cx", "資本支出", "DCF 頁"),
             ("wc", "營運資金增加", "DCF 頁"), ("fcf", "自由現金流（扣股利；Tesla 不配息）", "淨利＋折舊＋股票薪酬 − 資本支出 − 營運資金增加")]
        rows = {}
        for k, lab, note in L:
            P.cell(r, 1, lab); P.cell(r, 2, note); rows[k] = r; r += 1
        prow[s] = rows; r += 1
    for s in SCN:
        R = prow[s]
        for y in YEARS:
            c = PLC[y]; e = ENG[y]; adj = y >= 2027
            P[f"{c}{R['veh']}"] = f"={NG}{e}119+{NG}{e}121+{NG}{e}123+{NG}{e}125+{NG}{e}127"
            P[f"{c}{R['fsd']}"] = f"={NG}{e}129*fsd_k" if adj else f"={NG}{e}129"
            P[f"{c}{R['rt']}"] = (f"=rt_rev_{y}_{s}*rt_k" if adj else f"=N({NG}{e}131)")
            P[f"{c}{R['crd']}"] = f"={NG}{e}133"
            P[f"{c}{R['lease']}"] = f"={NG}{e}135"
            P[f"{c}{R['svc']}"] = f"={NG}{e}44"
            P[f"{c}{R['seg_auto']}"] = f"=SUM({c}{R['veh']}:{c}{R['svc']})"
            P[f"{c}{R['seg_en']}"] = f"={NG}{e}42"
            P[f"{c}{R['seg_opt']}"] = f"=Optimus!{c}{orow[s]['rev']}"
            P[f"{c}{R['rev']}"] = (f"={NG}{e}46+({c}{R['fsd']}-{NG}{e}129)+({c}{R['rt']}-N({NG}{e}131))+{c}{R['seg_opt']}")
            P[f"{c}{R['chk']}"] = f"={c}{R['seg_auto']}+{c}{R['seg_en']}+{c}{R['seg_opt']}-{c}{R['rev']}"
            P[f"{c}{R['gp']}"] = (f"={NG}{e}62+" + (f"auto_gm_adj*{c}{R['veh']}+" if adj else "") +
                                  f"({c}{R['fsd']}-{NG}{e}129)*fsd_gm+({c}{R['rt']}-N({NG}{e}131))*rt_gm+Optimus!{c}{orow[s]['gp']}")
            P[f"{c}{R['gm']}"] = f"={c}{R['gp']}/{c}{R['rev']}"
            P[f"{c}{R['opex']}"] = f"={NG}{e}69+Optimus!{c}{orow[s]['sga']}"
            P[f"{c}{R['ebit']}"] = f"={c}{R['gp']}-{c}{R['opex']}"
            P[f"{c}{R['nonop']}"] = f"={NG}{e}7"
            P[f"{c}{R['pti']}"] = f"={c}{R['ebit']}+{c}{R['nonop']}"
            P[f"{c}{R['tax']}"] = f"={c}{R['pti']}*tax_t" if adj else f"={NG}{e}9"
            P[f"{c}{R['mi']}"] = f"=({c}{R['pti']}-{c}{R['tax']})*{NG}$AJ$36" if adj else f"={NG}{e}11"
            P[f"{c}{R['ni']}"] = f"={c}{R['pti']}-{c}{R['tax']}-{c}{R['mi']}"
            P[f"{c}{R['sh']}"] = f"={NG}{e}34"
            P[f"{c}{R['eps']}"] = f"={c}{R['ni']}/{c}{R['sh']}"
            if y == 2025:
                P[f"{c}{R['da']}"] = "=h_da25"; P[f"{c}{R['sbc']}"] = "=h_sbc25"; P[f"{c}{R['cx']}"] = "=h_cx25"; P[f"{c}{R['wc']}"] = "—"
                P[f"{c}{R['fcf']}"] = "=h_ocf25-h_cx25"
            elif y == 2026:
                P[f"{c}{R['da']}"] = "=h_da9m+h_da_q3"; P[f"{c}{R['sbc']}"] = f"=h_sbc9m+{NG}AK2*sbc_r"; P[f"{c}{R['cx']}"] = "=cx_26"; P[f"{c}{R['wc']}"] = 0
                P[f"{c}{R['fcf']}"] = f"=h_ocf9m-h_cx9m+({NG}AK12+h_da_q3+{NG}AK2*sbc_r-(cx_26-h_cx9m))"
            else:
                P[f"{c}{R['da']}"] = f"=DCF!{DCC[y]}{{da_{s}}}"
                P[f"{c}{R['sbc']}"] = f"={c}{R['rev']}*sbc_r"
                P[f"{c}{R['cx']}"] = f"=DCF!{DCC[y]}{{cx_{s}}}"
                P[f"{c}{R['wc']}"] = f"=DCF!{DCC[y]}{{wc_{s}}}"
                P[f"{c}{R['fcf']}"] = f"={c}{R['ni']}+{c}{R['da']}+{c}{R['sbc']}-{c}{R['cx']}-{c}{R['wc']}"
            for k in R:
                cell = P[f"{c}{R[k]}"]
                cell.number_format = "0.0%" if k == "gm" else ("0.00" if k == "eps" else "#,##0")
    setw(P, {"A": 40, "B": 34})

    # ---------------- DCF
    D = wb.create_sheet("DCF", 4)
    D["A1"] = "DCF — 標準 FCFF（docs/plan/個股DCF標準規格.md）：NOPAT＋折舊 − 資本支出 − 營運資金增加；明確期 FY2027–FY2035（原檔引擎期間），C 欄＝4Q26 stub；期中折現；評價日＝股價日。US$M。"
    r = 3
    D.cell(r, 1, "折現率").font = BOLD; r += 1
    drow = {}
    for k, lab, f, fmt in [("ke", "ke＝rf＋β×ERP", "=rf+beta*erp", "0.00%"), ("kd_at", "稅後 kd", "=kd*(1-tax_w)", "0.00%"),
                           ("E_v", "股權市值（股價 × 稀釋股數；US$M）", "=px*h_sh_dil", "#,##0"), ("D_v", "負債（債務＋融資租賃；US$M）", "=h_debt+h_fl", "#,##0"),
                           ("w_e", "權益占比", f"=E_v/(E_v+D_v)", "0.0%"), ("wacc", "WACC", "=w_e*ke+(1-w_e)*kd_at+wacc_adj", "0.00%"),
                           ("g_dcf", "永續成長率 g", "=g_tv", "0.00%"), ("ronic", "RONIC", "=wacc+ronic_sp", "0.00%"),
                           ("nc", "淨現金與投資（現金＋短期投資 − 債務 − 融資租賃 − 非控制權益＋SpaceX 持股＋比特幣）", "=h_cash+h_sti-h_debt-h_fl-h_mi+h_spx+h_btc", "#,##0")]:
        D.cell(r, 1, lab); D.cell(r, 3, f).number_format = fmt; dn(wb, k, f"DCF!$C${r}"); drow[k] = r; r += 1
    r += 1
    hdr(D, r, ["項目", "說明", "4Q26 stub"] + [f"FY{y}" for y in DY]); r += 1
    drow_s = {}
    for s in SCN:
        D.cell(r, 1, f"情境 {s}").font = BOLD; D.cell(r, 1).fill = SEC; r += 1
        rows = {}
        for k, lab in [("rev", "合併營收"), ("core", "核心營收（不含 Optimus）"), ("ebit", "EBIT"), ("nopat", "NOPAT＝EBIT×(1−tax_t)"),
                       ("cxr", "核心資本支出 ÷ 核心營收"), ("cx", "資本支出（核心＋Optimus＋尾端）"), ("ppe0", "期初 PP&E 淨額"),
                       ("da", "折舊＝期初淨額 ÷ life"), ("ppe1", "期末 PP&E 淨額"), ("wc", "營運資金增加＝wc_r × 營收增量"),
                       ("fcff", "FCFF＝NOPAT＋折舊 − 資本支出 − 營運資金增加"), ("t", "折現年數（期中）"), ("df", "折現因子"), ("pv", "現值")]:
            D.cell(r, 1, lab); rows[k] = r; r += 1
        R = prow[s]
        # stub
        D[f"C{rows['rev']}"] = f"={NG}AK2"; D[f"C{rows['core']}"] = f"={NG}AK2"; D[f"C{rows['ebit']}"] = f"={NG}AK72"
        D[f"C{rows['nopat']}"] = f"=C{rows['ebit']}*(1-tax_w)"; D[f"C{rows['cx']}"] = "=cx_26-h_cx9m"
        D[f"C{rows['ppe0']}"] = "=h_ppe"; D[f"C{rows['da']}"] = "=h_da_q3"; D[f"C{rows['ppe1']}"] = f"=C{rows['ppe0']}+C{rows['cx']}-C{rows['da']}"
        D[f"C{rows['wc']}"] = 0
        D[f"C{rows['fcff']}"] = f"=C{rows['nopat']}+C{rows['da']}-C{rows['cx']}-C{rows['wc']}"
        D[f"C{rows['t']}"] = "=(DATE(2026,12,31)-px_date)/365.25/2"
        D[f"C{rows['df']}"] = f"=1/(1+wacc)^C{rows['t']}"; D[f"C{rows['pv']}"] = f"=C{rows['fcff']}*C{rows['df']}"
        for y in DY:
            c = DCC[y]; p = DCC[y - 1] if y > 2027 else "C"; pc = PLC[y]; ppc = PLC[y - 1]
            D[f"{c}{rows['rev']}"] = f"=PL!{pc}{R['rev']}"
            D[f"{c}{rows['core']}"] = f"=PL!{pc}{R['rev']}-PL!{pc}{R['seg_opt']}"
            D[f"{c}{rows['ebit']}"] = f"=PL!{pc}{R['ebit']}"
            D[f"{c}{rows['nopat']}"] = f"={c}{rows['ebit']}*(1-tax_t)"
            D[f"{c}{rows['cxr']}"] = f"=cx_r27+(cx_ss-cx_r27)*({y}-2027)/8"
            D[f"{c}{rows['cx']}"] = (f"={c}{rows['core']}*{c}{rows['cxr']}+Optimus!{pc}{orow[s]['cx']}" + ("+tail_tfab" if y in (2027, 2028) else ""))
            D[f"{c}{rows['ppe0']}"] = f"={p}{rows['ppe1']}"
            D[f"{c}{rows['da']}"] = f"={c}{rows['ppe0']}/life"
            D[f"{c}{rows['ppe1']}"] = f"={c}{rows['ppe0']}+{c}{rows['cx']}-{c}{rows['da']}"
            D[f"{c}{rows['wc']}"] = f"=wc_r*(PL!{pc}{R['rev']}-PL!{ppc}{R['rev']})"
            D[f"{c}{rows['fcff']}"] = f"={c}{rows['nopat']}+{c}{rows['da']}-{c}{rows['cx']}-{c}{rows['wc']}"
            D[f"{c}{rows['t']}"] = f"=(DATE({y},7,1)-px_date)/365.25"
            D[f"{c}{rows['df']}"] = f"=1/(1+wacc)^{c}{rows['t']}"
            D[f"{c}{rows['pv']}"] = f"={c}{rows['fcff']}*{c}{rows['df']}"
        for k in rows:
            for c in ["C"] + [DCC[y] for y in DY]:
                D[f"{c}{rows[k]}"].number_format = "0.0%" if k == "cxr" else ("0.000" if k in ("t", "df") else "#,##0")
        # summary
        L9 = DCC[2035]
        sm = {}
        for k, lab, f, fmt in [
            ("sumpv", "明確期 FCFF 現值合計（含 stub）", f"=SUM(C{rows['pv']}:{L9}{rows['pv']})", "#,##0"),
            ("tv", "終值（FY2035 年底）", f"=IF(tv_alt=1,{L9}{rows['fcff']}*(1+g_dcf)/(wacc-g_dcf),{L9}{rows['nopat']}*(1+g_dcf)*(1-g_dcf/ronic)/(wacc-g_dcf))", "#,##0"),
            ("tvpv", "終值現值", f"={{tv}}/(1+wacc)^((DATE(2035,12,31)-px_date)/365.25)", "#,##0"),
            ("ev", "企業價值 EV", "={sumpv}+{tvpv}", "#,##0"),
            ("tvsh", "終值占 EV（>80% 標示；EV≤0 顯示 N/M）", '=IF({ev}>0,{tvpv}/{ev},"N/M")', "0.0%"),
            ("eq", "股權價值（EV＋淨現金與投資）", "={ev}+nc", "#,##0"),
            ("eqT", "股權價值（兆美元，CEO 獎酬閘門判定用）", "={eq}/1000000", "0.000"),
            ("vfr", "CEO 獎酬歸屬比例", '=IF(ceo_ovr<0,COUNTIF(h_gate,"<="&{eqT})/12,ceo_ovr)', "0.0%"),
            ("vsh", "CEO 獎酬淨增股數（扣每股抵付 334.09 以股支付）", "=h_ceo_sh*{vfr}*MAX(0,1-h_ceo_off/MAX(0.01,{eq}/h_sh_dil))", "#,##0.0"),
            ("shv", "評價股數（稀釋＋CEO 獎酬淨增）", "=h_sh_dil+{vsh}", "#,##0.0"),
            ("vps", "每股價值（評價日）", "=MAX(0,{eq}/{shv})", "0.00"),
            ("v12", "DCF 情境價值推到 12 個月後（×(1+ke)，Tesla 不配息）", "={vps}*(1+ke)", "0.00"),
            ("chk", "檢查：FCFF 各列加總 − FCFF（全期）", f"=SUMPRODUCT(C{rows['nopat']}:{L9}{rows['nopat']}+C{rows['da']}:{L9}{rows['da']}-C{rows['cx']}:{L9}{rows['cx']}-C{rows['wc']}:{L9}{rows['wc']})-SUM(C{rows['fcff']}:{L9}{rows['fcff']})", "0.000"),
        ]:
            D.cell(r, 1, lab); sm[k] = r; r += 1
            D.cell(sm[k], 3, f.format(**{kk: f"C{vv}" for kk, vv in sm.items()})).number_format = fmt
        rows.update({"sum_" + k: v for k, v in sm.items()})
        drow_s[s] = rows; r += 1
    # backfill PL links
    for s in SCN:
        R = prow[s]
        for y in YEARS:
            for k in ("da", "cx", "wc"):
                v = P[f"{PLC[y]}{R[k]}"].value
                if isinstance(v, str) and "{" in v:
                    P[f"{PLC[y]}{R[k]}"] = v.replace("{da_%s}" % s, str(drow_s[s]["da"])).replace("{cx_%s}" % s, str(drow_s[s]["cx"])).replace("{wc_%s}" % s, str(drow_s[s]["wc"]))
    D.cell(r, 1, "檢查：WACC 由 In 頁輸入重算 − WACC"); D.cell(r, 3, "=(px*h_sh_dil*(rf+beta*erp)+(h_debt+h_fl)*kd*(1-tax_w))/(px*h_sh_dil+h_debt+h_fl)+wacc_adj-wacc").number_format = "0.000000"; DCHK_W = r; r += 1
    D.cell(r, 1, "檢查：淨現金與 Hist 資產負債核對"); D.cell(r, 3, "=nc-(h_cash+h_sti-h_debt-h_fl-h_mi+h_spx+h_btc)").number_format = "0.0"; r += 1
    setw(D, {"A": 52, "B": 4, "C": 12})

    # ---------------- SOTP
    S = wb.create_sheet("SOTP", 5)
    S["A1"] = ("SOTP — 分部加總（沿用原檔「長期評價分析」2035 年倍數法：2035 年各分部價值 ÷ (1+sotp_r)^sotp_n 折回現值，×(1+sotp_r) 為 12 個月目標）；"
               "Robotaxi 分部價值讀 tesla_robotaxi 分部 v0.3（RT_VAL_A/B/C，In rt_val_*）。US$M。")
    hdr(S, 3, ["項目", "說明"] + [f"情境 {s}" for s in SCN])
    srow = {}
    r = 4
    M = PLC[2035]; Lc = PLC[2034]
    L = [("a_rev", "車輛業務（非 FSD／Robotaxi）2035 營收", "PL 車輛銷售＋碳權＋租賃", lambda s: f"=PL!{M}{prow[s]['veh']}+PL!{M}{prow[s]['crd']}+PL!{M}{prow[s]['lease']}"),
         ("a_val", "  價值＝營收 × npm_auto × pe_auto", "原檔 W2、Y2", lambda s: "={a_rev}*npm_auto*pe_auto"),
         ("f_rev", "FSD 2035 營收", "PL", lambda s: f"=PL!{M}{prow[s]['fsd']}"),
         ("f_ps", "  P/S＝FY2035 年增率×100×peg_k", "原檔 X3", lambda s: f"=(PL!{M}{prow[s]['fsd']}/PL!{Lc}{prow[s]['fsd']}-1)*100*peg_k"),
         ("f_val", "  價值", "", lambda s: "={f_rev}*{f_ps}"),
         ("n_rev", "非汽車業務（能源＋服務）2035 營收", "原檔 V5 誤用 FY2030，本版改 FY2035", lambda s: f"=PL!{M}{prow[s]['seg_en']}+PL!{M}{prow[s]['svc']}"),
         ("n_val", "  價值＝營收 × npm_non × pe_non", "原檔 W5、Y5", lambda s: "={n_rev}*npm_non*pe_non"),
         ("o_rev", "Optimus 2035 營收", "Optimus 頁", lambda s: f"=PL!{M}{prow[s]['seg_opt']}"),
         ("o_ps", "  P/S＝FY2035 年增率×100×peg_k", "比照原檔成長分部", lambda s: f"=IF(PL!{Lc}{prow[s]['seg_opt']}>0,(PL!{M}{prow[s]['seg_opt']}/PL!{Lc}{prow[s]['seg_opt']}-1)*100*peg_k,0)"),
         ("o_val", "  價值", "", lambda s: "={o_rev}*{o_ps}"),
         ("v35", "2035 年價值合計（不含 Robotaxi）", "", lambda s: "={a_val}+{f_val}+{n_val}+{o_val}"),
         ("pv", "折回現值", "÷(1+sotp_r)^sotp_n", lambda s: "={v35}/(1+sotp_r)^sotp_n"),
         ("rt", "Robotaxi 分部價值（現值，讀分部 v0.3）", "In rt_val_*（US$B×1000）", lambda s: f"=rt_val_{s}*1000"),
         ("eq", "分部加總股權價值（現值）", "原檔不加淨現金，本版沿用", lambda s: "={pv}+{rt}"),
         ("vfr", "CEO 獎酬歸屬比例（同 DCF 邏輯）", "", lambda s: '=IF(ceo_ovr<0,COUNTIF(h_gate,"<="&{eq}/1000000)/12,ceo_ovr)'),
         ("vsh", "CEO 獎酬淨增股數", "", lambda s: "=h_ceo_sh*{vfr}*MAX(0,1-h_ceo_off/MAX(0.01,{eq}/h_sh_dil))"),
         ("shv", "評價股數", "", lambda s: "=h_sh_dil+{vsh}"),
         ("vps", "每股價值（現值）", "", lambda s: "={eq}/{shv}"),
         ("v12", "分部加總推到 12 個月後（×(1+sotp_r)，原檔 Z11）", "", lambda s: "={vps}*(1+sotp_r)"),
         ("rtps", "  其中 Robotaxi 分部價值貢獻（12 個月後，每股；讀分部 v0.3）", "", lambda s: "={rt}/{shv}*(1+sotp_r)")]
    for k, lab, note, fn in L:
        S.cell(r, 1, lab); S.cell(r, 2, note); srow[k] = r; r += 1
    for j, s in enumerate(SCN):
        col = openpyxl.utils.get_column_letter(3 + j)
        for k, lab, note, fn in L:
            f = fn(s).format(**{kk: f"{col}{vv}" for kk, vv in srow.items()})
            c = S[f"{col}{srow[k]}"]; c.value = f
            c.number_format = "0.0%" if k == "vfr" else ("0.0" if k in ("f_ps", "o_ps") else ("0.00" if k in ("vps", "v12", "rtps") else "#,##0"))
    setw(S, {"A": 58, "B": 36, "C": 14, "D": 14, "E": 14})

    # ---------------- Val
    V = wb.create_sheet("Val", 6)
    V["A1"] = "Val — 目標價＝DCF（w_dcf）＋分部加總（1 − w_dcf）；三情境加權（25／50／25）。評等五級：> +50% 強力買進、> +25% 買進、≥ −25% 中立、≥ −50% 放空、其餘強力放空（原則 24）。"
    hdr(V, 3, ["項目", "說明", "情境 A", "情境 B", "情境 C", "加權"])
    vrow = {}
    r = 4
    def vline(k, lab, note, cells, wf=None, fmt="0.00"):
        nonlocal r
        V.cell(r, 1, lab); V.cell(r, 2, note)
        for j, f in enumerate(cells):
            if f is not None: V.cell(r, 3 + j, f).number_format = fmt
        if wf: V.cell(r, 6, wf).number_format = fmt
        vrow[k] = r; r += 1
    vline("dcf", "DCF 情境價值推到 12 個月後", "DCF 頁", [f"=DCF!C{drow_s[s]['sum_v12']}" for s in SCN],
          "=SUMPRODUCT(C{r}:E{r},C{w}:E{w})")
    vline("sotp", "分部加總推到 12 個月後", "SOTP 頁", [f"=SOTP!{openpyxl.utils.get_column_letter(3 + j)}{srow['v12']}" for j, s in enumerate(SCN)],
          "=SUMPRODUCT(C{r}:E{r},C{w}:E{w})")
    vline("w", "情境權重", "In tp_w*", ["=tp_wA", "=tp_wB", "=tp_wC"], "=SUM(C{r}:E{r})", fmt="0%")
    for k in ("dcf", "sotp"):
        V.cell(vrow[k], 6, f"=SUMPRODUCT(C{vrow[k]}:E{vrow[k]},C{vrow['w']}:E{vrow['w']})")
    V.cell(vrow["w"], 6, f"=SUM(C{vrow['w']}:E{vrow['w']})")
    vline("tps", "情境目標價（DCF 與分部加總加權）", "", [f"=w_dcf*{c}{vrow['dcf']}+(1-w_dcf)*{c}{vrow['sotp']}" for c in "CDE"])
    r += 1
    V.cell(r, 1, "DCF 目標價（情境加權）"); V.cell(r, 3, f"=F{vrow['dcf']}").number_format = "0.00"; vrow["dcf_w"] = r; r += 1
    V.cell(r, 1, "分部加總目標價（情境加權；原檔 2035 年倍數法＋Robotaxi 讀分部 v0.3）"); V.cell(r, 3, f"=F{vrow['sotp']}").number_format = "0.00"; vrow["sotp_w"] = r; r += 1
    V.cell(r, 1, "目標價（DCF 與分部加總加權）").font = BOLD; V.cell(r, 3, f"=w_dcf*C{vrow['dcf_w']}+(1-w_dcf)*C{vrow['sotp_w']}").number_format = "0.00"; vrow["tp"] = r; dn(wb, "tp_final", f"Val!$C${r}"); r += 1
    V.cell(r, 1, "股價（股價日）"); V.cell(r, 3, "=px").number_format = "0.00"; r += 1
    V.cell(r, 1, "上檔空間（不是年化報酬）"); V.cell(r, 3, f"=tp_final/px-1").number_format = "0.0%"; vrow["up"] = r; r += 1
    RAT = f'IF({{x}}>0.5,"強力買進",IF({{x}}>0.25,"買進",IF({{x}}>=-0.25,"中立",IF({{x}}>=-0.5,"放空","強力放空"))))'
    V.cell(r, 1, "估值評等（只看上檔空間）"); V.cell(r, 3, "=" + RAT.format(x=f"C{vrow['up']}")); vrow["rate"] = r; r += 1
    gate_f = [f"=SUM(PL!{PLC[2026]}{prow[s]['fcf']}:{PLC[2030]}{prow[s]['fcf']})" for s in SCN]
    V.cell(r, 1, "收支關卡（FY2026–FY2030 自由現金流合計，扣股利；情境 A／B／C）"); V.cell(r, 2, "原則 63：基準情境 B < 0＝入不敷出")
    for j, f in enumerate(gate_f): V.cell(r, 3 + j, f).number_format = "#,##0"
    vrow["gate"] = r; r += 1
    V.cell(r, 1, "入不敷出判定用（情境 B 合計）"); V.cell(r, 3, f"=D{vrow['gate']}").number_format = "#,##0"; vrow["gate_b"] = r
    V.cell(r, 4, f'=IF(C{r}<0,"不通過","通過")'); r += 1
    V.cell(r, 1, "投資建議").font = BOLD
    V.cell(r, 3, f'=IF(AND(C{vrow["gate_b"]}<0,OR(C{vrow["rate"]}="買進",C{vrow["rate"]}="強力買進")),"收支不足",C{vrow["rate"]})'); vrow["rec"] = r; r += 1
    V.cell(r, 1, "條件式投資建議句")
    V.cell(r, 3, (f'="在 Optimus 與 Robotaxi 依情境路徑放量、Robotaxi 分部價值讀分部 v0.3、DCF 腿 Robotaxi 營收讀分部 v0.3、SOTP 沿用原檔 2035 年倍數法的假設下，加權目標價 "&TEXT(tp_final,"0.0")&" 美元，'
                  f'相對 "&TEXT(px_date,"yyyy-mm-dd")&" 收盤 "&TEXT(px,"0.00")&" 上檔空間 "&TEXT(C{vrow["up"]},"0.0%")&"，估值評等"&C{vrow["rate"]}&"；收支關卡情境 B 合計 "&TEXT(C{vrow["gate_b"]},"#,##0")&" 百萬美元（"&D{vrow["gate_b"]}&"），投資建議"&C{vrow["rec"]}&"。"'))
    vrow["sent"] = r; r += 1
    # 距邊界
    V.cell(r, 1, "距最近級距邊界（pt）"); V.cell(r, 3, f"=MIN(ABS(C{vrow['up']}-0.5),ABS(C{vrow['up']}-0.25),ABS(C{vrow['up']}+0.25),ABS(C{vrow['up']}+0.5))").number_format = "0.0%"; vrow["dist"] = r; r += 1
    V.cell(r, 1, "升一級所需目標價"); V.cell(r, 3, f"=px*(1+IF(C{vrow['up']}<-0.5,-0.5,IF(C{vrow['up']}<-0.25,-0.25,IF(C{vrow['up']}<=0.25,0.25000001,IF(C{vrow['up']}<=0.5,0.50000001,99)))))").number_format = "0.00"; vrow["up1"] = r; r += 1
    V.cell(r, 1, "降一級的目標價（低於此值即降級）"); V.cell(r, 3, f'=IF(C{vrow["up"]}<-0.5,"N/M（已是最低級）",px*(1+IF(C{vrow["up"]}>0.5,0.5,IF(C{vrow["up"]}>0.25,0.25,IF(C{vrow["up"]}>=-0.25,-0.25,-0.5)))))').number_format = "0.00"; vrow["dn1"] = r; r += 1
    # 三段歸因
    V.cell(r, 1, "上檔空間歸因：DCF 貢獻＝w_dcf ×（DCF ÷ 股價 − 1）"); V.cell(r, 3, f"=w_dcf*(C{vrow['dcf_w']}/px-1)").number_format = "0.0%"; r += 1
    V.cell(r, 1, "上檔空間歸因：分部加總貢獻（未做相對倍數調整）"); V.cell(r, 3, f"=(1-w_dcf)*(C{vrow['sotp_w']}/px-1)").number_format = "0.0%"; r += 1
    V.cell(r, 1, "上檔空間歸因：相對倍數調整貢獻"); V.cell(r, 3, 0).number_format = "0.0%"; V.cell(r, 4, "1.00（未調整：缺同業倍數與 10 年相對比值）"); r += 1
    V.cell(r, 1, "Robotaxi 分部價值（讀分部 v0.3）在加權目標價中的占比"); V.cell(r, 3, f"=(1-w_dcf)*SUMPRODUCT(SOTP!C{srow['rtps']}:E{srow['rtps']},C{vrow['w']}:E{vrow['w']})/tp_final").number_format = "0.0%"; vrow["rtsh"] = r; r += 1
    V.cell(r, 1, "情境與方法區間：最低"); V.cell(r, 3, f"=MIN(C{vrow['dcf']}:E{vrow['sotp']},C{vrow['tps']}:E{vrow['tps']},tp_final)").number_format = "0.00"; vrow["mn"] = r; r += 1
    V.cell(r, 1, "情境與方法區間：最高"); V.cell(r, 3, f"=MAX(C{vrow['dcf']}:E{vrow['sotp']},C{vrow['tps']}:E{vrow['tps']},tp_final)").number_format = "0.00"; vrow["mx"] = r; r += 1
    V.cell(r, 1, "檢查：基準落在區間內"); V.cell(r, 3, f'=IF(AND(tp_final>=C{vrow["mn"]},tp_final<=C{vrow["mx"]}),"通過","不通過")'); r += 1
    V.cell(r, 1, "期間：模型折現期間（股價日 → 目標價時點）"); V.cell(r, 3, '=TEXT(px_date,"yyyy-mm-dd")&" → "&TEXT(tp_date,"yyyy-mm-dd")'); r += 1
    V.cell(r, 1, "期間：持有期間（現價日起 12 個月）；上檔空間不是年化報酬"); V.cell(r, 3, "=(tp_date-px_date)/365.25").number_format = "0.00"; r += 1
    V.cell(r, 1, "DCF 終值占 EV（情境 B）"); V.cell(r, 3, f"=DCF!C{drow_s['B']['sum_tvsh']}").number_format = "0.0%"; vrow["tvsh"] = r; r += 1
    V.cell(r, 1, "收支關卡旁並列：情境 B 期間最低單年自由現金流"); V.cell(r, 3, f"=MIN(PL!{PLC[2026]}{prow['B']['fcf']}:{PLC[2030]}{prow['B']['fcf']})").number_format = "#,##0"; r += 1
    V.cell(r, 1, "收支關卡旁並列：評價日淨現金與投資"); V.cell(r, 3, "=nc").number_format = "#,##0"; r += 1
    setw(V, {"A": 62, "B": 30, "C": 14, "D": 14, "E": 14, "F": 12})

    # ---------------- Recon
    RC = wb.create_sheet("Recon", 7)
    RC["A1"] = "Recon — 對照與反解（原則 4、33、50）。"
    hdr(RC, 3, ["項目", "本模型", "對照值", "差", "說明"])
    r = 4
    rcon = {}
    for k, lab, a, b, note in [
        ("cons", "目標價 vs 市場共識（44 位分析師）", "=tp_final", "=cons_tp", "共識：stockanalysis 2026-11-14 擷取（對照值未更新）；Interested-party，只作對照"),
        ("old_fair", "原檔合理每股現價（長期評價分析 Z10）", "=tp_final", "='長期評價分析'!Z10", "原檔舊法（2035 年分部倍數 ÷ 1.15^10；不含 Optimus；Robotaxi 占 71%）"),
        ("old_tp", "原檔未來一年目標價（Z11）vs 本版加權目標價", "=tp_final", "='長期評價分析'!Z11", "差異來源：DCF 腿（規格新增）、Optimus 分部（新增）、Robotaxi 改讀分部 v0.3（價值與 FY2027 起營收）、非汽車業務改用 FY2035 營收"),
        ("old_sotp", "原檔 Z11 vs 本版分部加總（情境 B）", f"=SOTP!D{srow['v12']}", "='長期評價分析'!Z11", "本版情境 B 分部加總與原檔同法；差＝Optimus＋非汽車業務年度修正"),
        ("rt_old", "Robotaxi 分部價值（情境 B，US$B，讀分部 v0.3）vs 原檔舊法現值", "=rt_val_B", f"='長期評價分析'!Z4/'長期評價分析'!Z9/1000", "原檔舊法（FY35 營收 × P/S ÷ 1.15^10）；v0.1 的暫定值 627 即此數；v0.2 已由分部成品取代"),
        ("eps26", "FY2026 EPS（情境 B）vs 原檔", f"=PL!{PLC[2026]}{prow['B']['eps']}", f"={NG}CJ13", "FY2026 無調整；年度 EPS＝淨利÷年度股數，原檔為季 EPS 加總，差為口徑"),
        ("eps27", "FY2027 EPS（情境 B）vs 原檔", f"=PL!{PLC[2027]}{prow['B']['eps']}", f"={NG}CK13", "差＝FY2027 起 Optimus 與調整項；另含年度股數 vs 季 EPS 加總口徑"),
        ("rev35", "FY2035 合併營收（情境 B）vs 原檔", f"=PL!{PLC[2035]}{prow['B']['rev']}", f"={NG}CS46", "差＝Optimus 營收"),
    ]:
        RC.cell(r, 1, lab); RC.cell(r, 2, a).number_format = "#,##0.00"; RC.cell(r, 3, b).number_format = "#,##0.00"
        RC.cell(r, 4, f"=B{r}-C{r}").number_format = "#,##0.00"; RC.cell(r, 5, note); rcon[k] = r; r += 1
    r += 1
    RC.cell(r, 1, "反解（原則 4、36）：Robotaxi 分部價值三情境同加多少（US$B），加權目標價才會等於下列價格").font = BOLD; r += 1
    hdr(RC, r, ["目標", "目標價", "所需 Robotaxi 價值變動（US$B，三情境同加）", "所需情境 B Robotaxi 價值（US$B）", "說明"]); r += 1
    slope = f"((1-w_dcf)*SUMPRODUCT(1/SOTP!C{srow['shv']}:E{srow['shv']},Val!C{vrow['w']}:E{vrow['w']})*(1+sotp_r)*1000)"
    for lab, tgt in [("支持現價（目標價＝股價）", "=px"), ("升一級（估值評等邊界）", f"=Val!C{vrow['up1']}"), ("降一級（估值評等邊界）", f"=Val!C{vrow['dn1']}")]:
        RC.cell(r, 1, lab); RC.cell(r, 2, tgt).number_format = "0.00"
        RC.cell(r, 3, f'=IF(ISNUMBER(B{r}),(B{r}-tp_final)/{slope},"N/M")').number_format = "#,##0"
        RC.cell(r, 4, f'=IF(ISNUMBER(C{r}),rt_val_B+C{r},"N/M")').number_format = "#,##0"
        RC.cell(r, 5, "線性近似：評價股數固定在本版各情境值（CEO 獎酬閘門不隨之重判）")
        rcon.setdefault("req_px", r)
        r += 1
    r += 1
    RC.cell(r, 1, "Optimus 分部價值（情境 B，分部加總口徑，12 個月後每股）"); RC.cell(r, 2, f"=SOTP!D{srow['o_val']}/(1+sotp_r)^sotp_n/SOTP!D{srow['shv']}*(1+sotp_r)").number_format = "0.00"; r += 1
    RC.cell(r, 1, "Optimus 2035 營收（情境 B）÷ 合併營收"); RC.cell(r, 2, f"=PL!{PLC[2035]}{prow['B']['seg_opt']}/PL!{PLC[2035]}{prow['B']['rev']}").number_format = "0.0%"; r += 1
    RC.cell(r, 1, "Optimus FY2027 毛利率（情境 B）"); RC.cell(r, 2, f"=Optimus!{PLC[2027]}{orow['B']['gm']}").number_format = "0.0%"; r += 1
    RC.cell(r, 1, "Optimus FY2035 每台 BOM（情境 B，千美元）vs 2b BofA 2030–35 情境 1.3–1.7 萬"); RC.cell(r, 2, f"=Optimus!{PLC[2035]}{orow['B']['bom']}").number_format = "0.0"; RC.cell(r, 3, 15).number_format = "0.0"; RC.cell(r, 5, "BofA 中點 15（Interested-party，2b）"); r += 1
    RC.cell(r, 1, "FY2026 資本支出（模型）vs 公司說法約 200 億"); RC.cell(r, 2, "=cx_26"); RC.cell(r, 3, 20000); RC.cell(r, 5, "法說 2026-10-21（Interested-party）"); r += 1
    RC.cell(r, 1, "市場共識 FY2026／FY2027 EPS"); RC.cell(r, 2, "未取得"); RC.cell(r, 5, "免費來源未提供前瞻 EPS（stockanalysis 2026-11-14 只列 forward P/E 182.53，以當日股價計，未更新）"); r += 1
    setw(RC, {"A": 60, "B": 16, "C": 22, "D": 22, "E": 70})

    # ---------------- 對帳
    RB = wb.create_sheet("對帳", 8)
    RB["A1"] = "對帳 — 本模型使用的線層級變數（Robotaxi 各列 v0.2 起讀 tesla_robotaxi 分部 v0.3，其成本與價格讀線 1a v0.2；其餘線尚無完成版，先依線框架 v0.1 取值，不同處照寫，對帳時處理，不靜默改口徑）"
    hdr(RB, 3, ["變數名", "所屬線", "本模型使用值（或區間）", "單位", "口徑", "出處", "標記", "影響本模型哪些儲存格"])
    RBL = [
        ("FSD 活躍數（含早期買斷）", "1b", "174 萬（3Q26）", "戶", "全球、含買斷", S_1B + "；原檔 AJ93", "Verified", "Non-GAAP AJ93 → 季度 FSD 營收 → PL FSD 列"),
        ("FSD 月費", "1b", "原檔季 ASP 297 美元（＝99 美元／月）", "美元", "美國", S_1B + "；原檔第 110 列", "Verified", "Non-GAAP 第 110 列"),
        ("FSD 買斷存量（不付月費）", "1b", "87 萬（＝50% × 活躍數）", "戶", "公司稱過半為訂閱", "原檔第 173 列；法說", "Interested-party", "Non-GAAP 第 129、173 列"),
        ("FSD 採用率路徑（累積交車滲透率）", "1b", "FY26 18% → FY35 60%", "%", "累積交車分母", "原檔第 94 列（原檔假設）", "Assumed", "PL FSD 列、SOTP FSD 價值"),
        ("FSD 北美新車搭配率", "1b", ">75%（3Q26）", "%", "北美交車", S_1B, "Interested-party", "未直接使用（只作採用率路徑佐證）"),
        ("Robotaxi 車隊（運行台數）", "1a", "4Q26 720 台 → FY35 163 萬台（原檔）；德州車籍 >1,200 台、其中 Cybercab 約 570（Fox Business 轉述 DMV，2026-11）", "台", "原檔＝全球運行台數；德州＝登記數", "原檔第 95 列；heygotrade 2026-11-09；" + S_1A, "Assumed", "PL Robotaxi 營收改讀分部 v0.3 RT_REV（FY2027 起）；原檔車隊路徑只剩 FY2026 以前與原檔舊法對照；SOTP 不用（讀 rt_val）"),
        ("Robotaxi 每英里票價", "1a", "原檔 1.5 美元／英里（3Q26）→ FY30 1.0", "美元/英里", "乘客付費英里", "原檔第 115 列；1a：奧斯汀起跳 3.25＋每英里 1.00", "Assumed", "Non-GAAP 第 112 列每車每季營收（FY2026 以前）；FY2027 起由分部 v0.3 的價格路徑（線 1a v0.2 Tesla 車隊加權：B 2030 1.11、2032 0.90）取代，不再進 PL"),
        ("Robotaxi 每車每日單量、每單里程", "1a", "18 單／日、6.06 英里／單", "—", "付費趟", "原檔第 113–114 列", "Assumed", "同上"),
        ("Robotaxi 每英里營運成本", "1a", "DCF 腿未用（以毛利率 rt_gm 55% 代替，調整額＝分部營收與原檔路徑之差）；分部 v0.3 讀線：B 2027 5.69、2030 1.07、2032 0.71（不含車輛折舊），Tesla 目標 0.30 美元／英里", "美元/英里", "定義未明", S_1A + "；" + S_CALL + "；" + SEG["file"], "Interested-party", "PL 毛利（rt_gm）。DCF 腿的毛利與分部 v0.3 每英里經濟（B 情境 EBIT 近 0）不同，待對帳"),
        ("Robotaxi 分部價值", "1a", f"情境 A／B／C＝{SEG['val']['A']:.1f}／{SEG['val']['B']:.1f}／{SEG['val']['C']:.1f}", "US$B", "現值（分部 EV）", f"{SEG['file']}｜給Tesla模型 RT_VAL_A/B/C", "Derived", "In rt_val_* → SOTP Robotaxi 列 → 分部加總目標價（v0.1 暫定值 1,200／627／150 已取代）"),
        ("Robotaxi 營收路徑 FY2027–FY2035", "1a", f"情境 B FY2030／FY2035＝{SEG['rev'][('B', 2030)]:,.0f}／{SEG['rev'][('B', 2035)]:,.0f}（原檔路徑 B：約 13,960／106,000）", "US$M", "自有車隊全額＋車主車輛抽成", f"{SEG['file']}｜給Tesla模型 RT_REV_<情境>_<年>", "Derived", "In rt_rev_<年>_<情境> → PL Robotaxi 列 → 營收、毛利調整、DCF 腿"),
        ("人形每台 BOM（Gen 2 類）", "2b", "55（35–65）", "千美元", "全尺寸、不含軟體", S_2B + "（MS 5–6 萬、BofA 3.5 萬）", "Interested-party", "Optimus 頁 BOM 列"),
        ("人形 BOM 學習率", "2b", "15%（10–25%）", "%", "每倍增累積產量", S_2B, "Analogy", "Optimus 頁 BOM 列"),
        ("Optimus 對外交付時點與產能", "2b", "2027 年底起對外；Fremont 設計 100 萬台／年", "台", "公司說法", S_2B + "；" + S_CALL, "Interested-party", "In opt_u_*（基準路徑）"),
        ("Optimus 售價", "2b", "FY27 40 → FY35 25（千美元）", "千美元", "出廠價", S_2B + "（Musk 成本 3 萬→2 萬）", "Assumed", "Optimus 頁營收"),
        ("美國工人全負擔時薪", "2b", "46.89", "美元/小時", "私部門全體", S_2B + "（BLS ECEC 2026-06）", "Verified", "Optimus 頁回本期（不進目標價）"),
        ("相對人工生產力", "2b", "0.5（0.3–1.0）", "x", "—", S_2B, "Assumed", "Optimus 頁回本期（不進目標價）"),
        ("Tesla 自建訓練算力（AI5、Terafab、400 MW）", "算力收支／AI供應鏈（只讀）", "未取得（本模型不建模；資本支出以占營收比率與 Terafab 尾端 Sens 涵蓋）", "MW、US$M", "—", "法說：2026 年底訓練算力目標 400 MW（Drive Tesla Canada 轉述）", "—", "DCF 資本支出（cx_r27、tail_tfab）"),
    ]
    r = 4
    for t in RBL:
        for j, v in enumerate(t): RB.cell(r, 1 + j, v)
        r += 1
    setw(RB, {"A": 30, "B": 10, "C": 46, "D": 10, "E": 20, "F": 60, "G": 15, "H": 44})

    # ---------------- 風險登錄
    RK = wb.create_sheet("風險登錄", 9)
    RK["A1"] = "風險登錄（原則 10、19）：未建模或只在 Sens 呈現的風險。"
    hdr(RK, 3, ["風險", "是否建模", "理由", "觀察指標"])
    for i, t in enumerate([
        ("Terafab 與 AI 算力資本支出超支（與 SpaceX 分攤未定）", "Sens（tail_tfab）", "公司未給 2027 資本支出指引", "2026Q4 法說 2027 資本支出指引；10-K 購買承諾"),
        ("EU FSD 表決（2026-12）未過", "部分（fsd_k 低）", "影響 FSD 採用率路徑", "EU TCMV 表決結果、RDW 公告"),
        ("中國 FSD 與資料不出境限制", "否", "原檔路徑未區分地區", "中國工信部核准、Tesla 中國 FSD 推送"),
        ("Robotaxi 事故與責任（NHTSA 調查、召回）", "部分（rt_rev_<年>_C、rt_val_C）", "情境 C 以規模縮小代表", "NHTSA SGO 事故資料、德州／加州許可"),
        ("關鍵人物（CEO 注意力分散至 SpaceX、xAI）與關係人交易", "否", "不可量化", "董事會關係人交易揭露、CEO 獎酬歸屬"),
        ("2025 CEO 獎酬稀釋", "是（市值閘門；Sens 0%／100%）", "營運里程碑未逐項建模", "市值、FSD 1,000 萬訂閱、100 萬台 Robotaxi／機器人、EBITDA 4,000 億"),
        ("碳權收入消失（美國政策）", "部分（原檔每年遞減 20%）", "2026-11 國會改選結果影響未評估", "碳權季營收（3Q26 113，年減 73%）"),
        ("Cybertruck 停產與車款老化", "是（2027 年中停產）", "Musk X 說法、Tesla 未正式公告", "Giga Texas 產線公告、季度 Other models 交車"),
        ("Optimus 壽命與可靠度（MTBF）", "否（以情境 C 量倍數代表）", "無公開實測", "首批客戶部署揭露、Tesla 法說"),
        ("SpaceX 股權公允價值波動（2026-12 解禁）", "是（計入淨現金與投資 2,865）", "以 9/30 公允價值計", "SpaceX 估值、Tesla 10-K 公允價值"),
    ]):
        for j, v in enumerate(t): RK.cell(4 + i, 1 + j, v)
    setw(RK, {"A": 50, "B": 28, "C": 40, "D": 60})

    # ---------------- 改版紀錄
    CH = wb.create_sheet("改版紀錄_v0.3", 10)
    CH["A1"] = "改版紀錄（v0.3 → v0.3.1：2026-11-23；v0.2 → v0.3：2026-11-17；v0.1 → v0.2：2026-11-16）；v0.1 列為相對原檔 20261028_TeslaTSLA.US_post_3Q26.xlsx 的變更；v0.3 個人字樣改動儲存格：" + ("、".join(SCRUBBED) or "無")
    hdr(CH, 3, ["#", "變更", "位置", "原值", "新值／說明"])
    r = 4
    items = [("v0.3.1 改", "現價換 2026-11-20 收盤（Yahoo Finance 日線，physical_ai/data/closes_20261109_20261120.csv）；股價日改 2026-11-20；其他假設不變；共識目標價、forward P/E 等對照值未更新", "In px、px_date；README", "390.22（2026-11-13）", "345.82（2026-11-20）"),
             ("v0.3 改", "Robotaxi 分部價值 rt_val_A/B/C 與 Robotaxi 營收 rt_rev_* 改讀分部 v0.3（對帳 2：市場、滲透、份額讀線 1a，與 Waymo 共用；每趟英里收斂；誘發相對 2026）", "In rt_val_*、rt_rev_*、SOTP、PL", "分部 v0.2：133.2／−1.0／−12.0", f"{SEG['val']['A']:.1f}／{SEG['val']['B']:.1f}／{SEG['val']['C']:.1f}（US$B）"),
             ("v0.3 改", "公開前清除個人字樣：原檔「修改紀錄_3Q26」頁 A53 等含人名的儲存格改為中性用語；原檔判斷的標記改為「原檔假設（非外部來源）」", "修改紀錄_3Q26、In、對帳", "含人名的舊標記", "原檔待決定事項／原檔假設"),
             ("v0.3 註記", "口徑差異：DCF 腿情境股權價值為負時歸零（原檔 MAX(0,·) 限責處理）；SOTP 腿 Robotaxi 分部價值為負時不歸零（分部虧損是補貼訂價的一部分，對帳 1 判斷題 3）", "README、Val", "—", "兩腿不同口徑，刻意保留"),
             ("v0.2 改", "Robotaxi 分部價值 rt_val_A/B/C：由暫定 1,200／627／150 改讀分部 v0.2 RT_VAL_A/B/C", "In rt_val_*、SOTP", "1,200／627／150（Assumed，B＝原檔舊法現值）", f"{SEG['val']['A']:.1f}／{SEG['val']['B']:.1f}／{SEG['val']['C']:.1f}（Derived；{SEG['file']}）"),
             ("v0.2 改", "PL Robotaxi 營收 FY2027–FY2035：由「原檔第 131 列 × rt_k_A/B/C（1.5／1.0／0.3）」改讀分部 v0.2 RT_REV_<情境>_<年>", "In 第 G2 節 rt_rev_*、PL", "原檔路徑（B FY2035 約 1,060 億）", f"分部 v0.2（B FY2035 {SEG['rev'][('B', 2035)]:,.0f} 百萬）；rt_k 改為單一 Sens 乘數（基準 1.0）"),
             ("v0.2 改", "Val、SOTP、README、Recon、對帳頁文字：「Robotaxi 暫定值」改「讀分部 v0.2」", "Val、SOTP、README、Recon、對帳", "暫定值", "讀分部 v0.2"),
             ("v0.1 新增", "In、Hist、Optimus、PL、DCF、SOTP、Val、Sens、Recon、對帳、風險登錄、README 頁", "新工作表", "—", "依 docs/plan 個股模型原則與 DCF 標準規格；原檔各頁保留")]
    for cell, old, new in fixes:
        items.append(("v0.1 " + ("修正" if cell[:2] in ("CI", "CJ", "CK", "CL", "CM") else "滾入新資訊"), f"Non-GAAP!{cell}", "Non-GAAP", "原式 " + str(old), "新式 " + str(new)))
    items.append(("v0.1 滾入新資訊", "Non-GAAP 第 86、124 列（Cybertruck 年增率）預估期包 IFERROR", "Non-GAAP", "除以零", "停產後顯示 -"))
    items += [("v0.1 修正", "非汽車業務 2035 年價值改用 FY2035 營收（原檔 V5＝N30＋N31 為 FY2030）", "SOTP 頁", "FY2030", "FY2035；原檔頁保留並加註「舊法」"),
              ("v0.1 新增", "Optimus 分部（原檔無）", "Optimus、PL、SOTP", "—", "讀線框架 2b"),
              ("v0.1 改", "Robotaxi 分部價值改為輸入（讀 tesla_robotaxi 模型，暫定）", "In rt_val_*、SOTP", "原檔 2035 營收×P/S 23.9", "暫定值＝原檔舊法現值 627（B）"),
              ("v0.1 新增", "DCF 腿（標準 FCFF，FY2027–FY2035）", "DCF", "原檔無", "目標價＝DCF 50%＋分部加總 50%"),
              ("v0.1 新增", "資產負債、股數、SpaceX 持股、CEO 獎酬", "Hist", "原檔無", "SEC XBRL 10-Q 2026-09-30")]
    for i, t in enumerate(items):
        CH.cell(r, 1, i + 1)
        for j, v in enumerate(t): CH.cell(r, 2 + j, v)
        r += 1
    setw(CH, {"A": 4, "B": 14, "C": 60, "D": 24, "E": 60})

    # ---------------- Sens（骨架；數值於重算後填入）
    SE = wb.create_sheet("Sens", 7)
    SE["A1"] = "Sens — 一次改一個輸入（In 頁敏感度低／高值），整本活頁簿以 LibreOffice 重算後的加權目標價與 DCF 單法目標價（建置快照，數值）；模型狀態與 build_id 見本頁底部。"
    SE["A2"] = ("豁免：opt_prod 不列：只用於 Optimus 頁回本期展示，不進目標價。"
                "sbc_r 不列：只用於收支關卡的自由現金流，DCF 以 NOPAT 扣股票薪酬，對目標價結構上恆為 0。"
                "tax_t 的高值與基準相同，只列低值。")
    hdr(SE, 4, ["驅動", "設定", "鍵", "加權目標價（低）", "加權目標價（高）", "DCF 目標價（低）", "DCF 目標價（高）", "估值評等（低）", "估值評等（高）", "說明"])
    SE["L2"] = "=SUMPRODUCT(In!$D$4:$F$%d,ROW(In!$D$4:$F$%d)*3+COLUMN(In!$D$4:$F$%d))" % (IN_LAST, IN_LAST, IN_LAST)
    SE["K2"] = "指紋暫存"
    wb._tesla = dict(keyrow=keyrow, IN_LAST=IN_LAST, vrow=vrow, prow=prow, srow=srow, drow_s=drow_s, rcon=rcon)
    # README（內容於 finalize 填）
    RD = wb.create_sheet("README", 0)
    wb.save(path_out)
    return wb._tesla


# ------------------------------------------------------------------ LibreOffice 重算
def recalc(src, outdir, tag="x"):
    prof = tempfile.mkdtemp(prefix=f"lo_{tag}_")
    cmd = ["soffice", f"-env:UserInstallation=file://{prof}", "--headless", "--calc", "--convert-to", "xlsx", "--outdir", outdir, src]
    subprocess.run(cmd, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, timeout=600)
    shutil.rmtree(prof, ignore_errors=True)
    out = os.path.join(outdir, os.path.basename(src))
    return out


def read_vals(path, meta):
    wb = openpyxl.load_workbook(path, data_only=True)
    V = wb["Val"]; vr = meta["vrow"]
    return dict(tp=V.cell(vr["tp"], 3).value, dcf=V.cell(vr["dcf_w"], 3).value, rate=V.cell(vr["rate"], 3).value)


def fingerprint_formula(meta):
    rng = f"In!$D$4:$F${meta['IN_LAST']}"
    return f"=SUMPRODUCT({rng},ROW({rng})*3+COLUMN({rng}))"


SENS_CASES = []   # (驅動, 設定, 鍵欄, {鍵: (低值, 高值)}, 說明)


def sens_cases(meta, basewb_vals):
    """依 In 頁 D／F 值（重算後）建立案例；Optimus 量路徑整組一案。"""
    kv = basewb_vals
    cases = []
    opt_keys = [f"opt_u_{y}" for y in OPT_U]
    done = set()
    for it in IN:
        if it[0] == "sec": continue
        k = it[0]
        if k in done or k in ("opt_prod", "sbc_r"): continue
        lo, base, hi = kv[k]
        if k in opt_keys:
            cases.append(("Optimus 量路徑（整組 ×0.25／×2）", "opt_u_2027…opt_u_2035 同時改", "、".join(opt_keys),
                          {kk: (kv[kk][0], kv[kk][2]) for kk in opt_keys}, "情境 B 量路徑；情境 A／C 再乘 opt_k"))
            done.update(opt_keys); continue
        if not isinstance(lo, (int, float)) and not isinstance(hi, (int, float)): continue
        if (lo is None or lo == base) and (hi is None or hi == base): continue
        lab = it[1]
        cases.append((lab, f"低 {lo if lo is not None else '—'}／高 {hi if hi is not None else '—'}", k,
                      {k: (lo if lo != base else None, hi if hi != base else None)}, it[8] if len(it) > 8 else ""))
    return cases


def run_sens(base_xlsx, meta, kv, workdir):
    cases = sens_cases(meta, kv)
    jobs = []
    for i, (drv, st, key, m, note) in enumerate(cases):
        for side in (0, 1):
            vals = {k: v[side] for k, v in m.items()}
            if all(v is None for v in vals.values()): continue
            jobs.append((i, side, vals))

    def one(job):
        i, side, vals = job
        d = tempfile.mkdtemp(prefix="sens_", dir=workdir)
        src = os.path.join(d, f"case_{i}_{side}.xlsx")
        wb = openpyxl.load_workbook(base_xlsx)
        I = wb["In"]
        for k, v in vals.items():
            if v is not None: I.cell(meta["keyrow"][k], 5).value = v
        wb.save(src)
        od = os.path.join(d, "out"); os.makedirs(od)
        out = recalc(src, od, tag=f"{i}_{side}")
        return job, read_vals(out, meta)
    res = {}
    with ThreadPoolExecutor(max_workers=4) as ex:
        for job, v in ex.map(one, jobs):
            res[(job[0], job[1])] = v
    return cases, res


def read_in_values(path, meta):
    wb = openpyxl.load_workbook(path, data_only=True)
    I = wb["In"]
    return {k: tuple(I.cell(r, c).value for c in (4, 5, 6)) for k, r in meta["keyrow"].items()}


def finalize(stage1, meta, base_vals, fp_base, cases, res, build_id, out):
    wb = openpyxl.load_workbook(stage1)
    SE = wb["Sens"]
    r = 5
    for i, (drv, st, key, m, note) in enumerate(cases):
        SE.cell(r, 1, drv); SE.cell(r, 2, st); SE.cell(r, 3, key)
        for side, (ct, cd, cr) in enumerate(((4, 6, 8), (5, 7, 9))):
            v = res.get((i, side))
            if v:
                SE.cell(r, ct, v["tp"]).number_format = "0.00"; SE.cell(r, cd, v["dcf"]).number_format = "0.00"; SE.cell(r, cr, v["rate"])
        SE.cell(r, 10, note)
        r += 1
    r += 1
    SE.cell(r, 1, "基準加權目標價（建置時）"); SE.cell(r, 4, base_vals["tp"]).number_format = "0.00"; SE_BASE = r; r += 1
    SE.cell(r, 1, "基準加權目標價（本檔即時）"); SE.cell(r, 4, "=tp_final").number_format = "0.00"; SE_LIVE = r; r += 1
    SE.cell(r, 1, "基準 DCF 目標價（建置時）"); SE.cell(r, 4, base_vals["dcf"]).number_format = "0.00"; r += 1
    SE.cell(r, 1, "輸入指紋（本檔即時重算）"); SE.cell(r, 4, fingerprint_formula(meta)).number_format = "0.000000"; FP_LIVE = r; r += 1
    SE.cell(r, 1, "輸入指紋（敏感度重算時的基準檔）"); SE.cell(r, 4, fp_base).number_format = "0.000000"; FP_BASE = r; r += 1
    SE.cell(r, 1, "模型狀態（CURRENT＝Sens 與本檔輸入一致；STALE＝In 頁已改、須重跑 build_tesla.py）").font = BOLD
    SE.cell(r, 4, f'=IF(AND(ISNUMBER(D{FP_BASE}),ABS(D{SE_LIVE}-D{SE_BASE})<0.005,ABS(D{FP_LIVE}-D{FP_BASE})<=0.000000001*MAX(1,ABS(D{FP_BASE}))),"CURRENT","STALE")'); SE_STAT = r; r += 1
    SE.cell(r, 1, "build_id"); SE.cell(r, 4, build_id); SE_BID = r; r += 1
    for c, w in zip("ABCDEFGHIJ", (46, 30, 30, 14, 14, 14, 14, 12, 12, 50)):
        SE.column_dimensions[c].width = w

    # README 頁
    RD = wb["README"]
    vr = meta["vrow"]; pr = meta["prow"]
    RD["A1"] = f'="Tesla（TSLA）個股模型 {VER}｜build_id "&Sens!D{SE_BID}&"｜資料截止日 2026-11-16（股價 2026-11-20）｜成品日 2026-11-23｜股價基準日 "&TEXT(px_date,"yyyy-mm-dd")'
    lines = [
        ("命題", "Tesla 的價值有多少來自「AI 進入實體世界」（FSD、Robotaxi、Optimus），而不是汽車與能源本業？本業在 3Q26 營業利益率僅 2.6%，市值主要押在三個尚未有分部揭露的業務。"),
        ("命題回覆", f'="模型輸出（非獨立證據）：Robotaxi 分部價值（SOTP 腿，讀分部 v0.3）占加權目標價 "&TEXT(Val!C{vr["rtsh"]},"0%")&"；DCF 目標價 "&TEXT(Val!C{vr["dcf_w"]},"0.0")&"、分部加總 "&TEXT(Val!C{vr["sotp_w"]},"0.0")&"，支持現價所需的 Robotaxi 分部價值（情境 B）約 "&TEXT(Recon!D{meta["rcon"]["req_px"]},"#,##0")&" 十億美元（分部 v0.3 值 "&TEXT(rt_val_B,"#,##0")&"；見 Recon 反解）。"'),
        ("投資建議", f"=Val!C{vr['sent']}"),
        ("估值評等／收支", f'="估值評等 "&Val!C{vr["rate"]}&"；距最近級距邊界 "&TEXT(Val!C{vr["dist"]},"0.0%")&"（目標價達 "&TEXT(Val!C{vr["up1"]},"0.00")&" 即升一級）；收支關卡（FY2026–FY2030 自由現金流，情境 B）"&TEXT(Val!C{vr["gate_b"]},"#,##0")&" 百萬美元 "&Val!D{vr["gate_b"]}&"；投資建議 "&Val!C{vr["rec"]}'),
        ("分部", "官方報告分部：汽車（含 FSD、Robotaxi、碳權、租賃、服務及其他）、能源發電與儲存；Optimus 尚未列報，模型另列一列（歷史為 0）。命題子列 FSD、Robotaxi 在汽車分部之下，合計＝汽車分部。"),
        ("方法", "以原檔 Non-GAAP 季度引擎為底（量價、毛利、費用、FSD、Robotaxi 路徑全部保留）；FY2027 起在 PL 頁加調整與情境係數；DCF＝標準 FCFF 至 FY2035；FY2027 起 Robotaxi 營收讀 tesla_robotaxi 分部 v0.3（RT_REV_<情境>_<年>），其餘維持原檔；分部加總＝原檔 2035 年倍數法（折現 15%、10 年），Robotaxi 價值讀分部 v0.3（RT_VAL_A/B/C）、新增 Optimus。"),
        ("主要失效條件", "(1) 2027 年 Tesla 揭露每車每日付費英里 > 100 且無安全員城市 ≥ 10（線的成本錨點應下修）；(2) Optimus 2027 年底未開始對外交付；(3) 非 FSD 汽車毛利率連兩季低於 15%；(4) 2027 資本支出指引高於 350 億美元。"),
        ("可證偽指標（承重 Assumed）", "rt_val／rt_rev：德州與加州車籍、Robotaxi 每週付費趟數、每車每日付費英里、價格減全成本（tesla_robotaxi 分部 v0.3、線 1a）；opt_u：Tesla 法說對外交付台數；opt_bom0／opt_lr：Tesla 對 Optimus 成本說法與供應商報價；cx_r27：2026Q4 法說 2027 資本支出指引；fsd_k：EU 12 月表決、FSD 活躍數季增。"),
        ("判斷題", "DCF 明確期到 FY2035（原檔引擎期間），FY2035 時 Robotaxi 營收年增仍約 30%，終值以 g 3%、RONIC＝WACC 計，等於假設 2035 年後新投資不創造價值；分部加總則以 2035 年成長率給 P/S（原檔 PEG 法）。兩法對 2035 年後成長的隱含假設相反，加權目標價接近強力放空／放空邊界，權重 w_dcf 是最大的方法判斷（Sens：0.3／0.7）。"),
        ("限制", "未做同業倍數（原則 7）、HTML 一頁摘要、自算產業 β；Robotaxi 的 DCF 腿毛利（rt_gm 55%，套在分部營收與原檔路徑之差）與分部 v0.3 的每英里經濟不同，屬待對帳項；Optimus 全部為假設（公司未揭露）；SOTP 沿用原檔「不加淨現金」口徑；訓練算力資本支出未取得（AI供應鏈只讀）。"),
        ("模型狀態", f'="build_id "&Sens!D{SE_BID}&"；模型狀態 "&Sens!D{SE_STAT}'),
    ]
    r = 3
    for a, b in lines:
        RD.cell(r, 1, a).font = BOLD; RD.cell(r, 2, b); RD.cell(r, 2).alignment = Alignment(wrap_text=True, vertical="top"); r += 1
    r += 1
    RD.cell(r, 1, "資料基準日").font = BOLD; r += 1
    for a, b in [("股價", "2026-11-20 收盤（Yahoo Finance 日線，physical_ai/data/closes_20261109_20261120.csv）；前版 2026-11-13 收盤 390.22"), ("財報", "3Q26 10-Q（2026-09-30；2026-10-22 申報）"), ("rf", "2026-10-23（逾期：早於股價日 >2 個交易日；跨模型一致）"),
                 ("β", "Yahoo 2026-11-16、stockanalysis 2026-11-14 擷取"), ("共識目標價", "2026-11-14（財報後；對照值，v0.3.1 未更新）")]:
        RD.cell(r, 1, a); RD.cell(r, 2, b); r += 1
    RD.column_dimensions["A"].width = 24; RD.column_dimensions["B"].width = 140
    wb.save(out)


def main():
    work = tempfile.mkdtemp(prefix="tsla_")
    stage1 = os.path.join(work, OUT_NAME)
    meta = build(stage1)
    # build_id＝版本＋建置輸入雜湊（原檔＋本腳本）
    h = hashlib.sha1(open(BASE, "rb").read() + open(__file__, "rb").read()).hexdigest()[:10]
    build_id = f"{VER}-{h}"
    od = os.path.join(work, "base"); os.makedirs(od)
    base_calc = recalc(stage1, od, "base")
    base_vals = read_vals(base_calc, meta)
    kv = read_in_values(base_calc, meta)
    print("基準：", base_vals)
    # 指紋（以重算後 In D:F 數值）
    fp = openpyxl.load_workbook(base_calc, data_only=True)["Sens"]["L2"].value
    print("指紋", fp)
    cases, res = ([], {})
    if DO_SENS:
        cases, res = run_sens(stage1, meta, kv, work)
    stage2 = os.path.join(work, "final_" + OUT_NAME)
    finalize(stage1, meta, base_vals, fp, cases, res, build_id, stage2)
    od2 = os.path.join(work, "final"); os.makedirs(od2)
    fin = recalc(stage2, od2, "final")
    dst = os.path.join(HERE, "dist", OUT_NAME)
    shutil.copy(fin, dst)
    print("成品：", dst, "build_id", build_id)
    json.dump({"build_id": build_id, "base": base_vals, "n_cases": len(cases)}, open(os.path.join(work, "summary.json"), "w"), ensure_ascii=False)
    print("work", work)


if __name__ == "__main__":
    main()
