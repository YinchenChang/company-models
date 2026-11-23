# -*- coding: utf-8 -*-
"""線 3 倉儲物流自動化 單位經濟模型 v0.1 產生器。

用法：python3 build_warehouse.py
輸出：dist/20261116_倉儲物流單位經濟_v0.1.xlsx（經 LibreOffice 重算，含快取值）
檢查：重算後 0 錯誤值；Excel 重算值與 Python 同式求值一致（相對誤差 < 1e-6）。
"""
import math, os, re, shutil, subprocess, sys, tempfile
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

HERE = os.path.dirname(os.path.abspath(__file__))
DIST = os.path.join(HERE, "dist")
OUT_NAME = "20261116_倉儲物流單位經濟_v0.1.xlsx"
VERSION = "v0.1"
YEARS = list(range(2026, 2033))
TOPK = 3  # 情境悲觀／樂觀：同時移動的承重變數個數
SCEN_FIXED = {"H"}  # 年營運時數是場址設定（班次），情境固定在基準 2 班；其影響見敏感度頁

BLUE = Font(color="0000FF")
BOLD = Font(bold=True)
TITLE = Font(bold=True, size=13)
HDR = PatternFill("solid", fgColor="DDEBF7")
YEL = PatternFill("solid", fgColor="FFF2CC")
GRY = PatternFill("solid", fgColor="F2F2F2")
WRAP = Alignment(wrap_text=True, vertical="top")

# ---------------------------------------------------------------- 輸入
# (代號, 變數, 單位, 低, 基準, 高, 標記, 類別, 出處, 日期, 備註)
INPUTS = [
 # --- 人工基準
 ("W0", "美國倉儲業平均時薪（生產與非主管員工）", "USD/時", 26.02, 26.02, 26.02, "[Verified]", "[A]",
  "BLS CES，NAICS 493 Warehousing and Storage（https://www.bls.gov/iag/tgs/iag493.htm）", "2026-09（初值）",
  "全體員工 26.87；本模型用生產與非主管口徑（揀貨員較接近）"),
 ("LF", "全負擔係數（工資＋福利）", "倍", 1.30, 1.35, 1.40, "[Interested-party]", "[A]",
  "Conveyco（整合商，有推銷自動化誘因）：全負擔比時薪高 30–40%", "2026",
  "中國與亞太時薪未取得，本版只做美國"),
 ("M", "人工揀貨生產力", "訂單行/人時", 80, 100, 120, "[Interested-party]", "[P]",
  "Pallite（包材商部落格）：人工電商倉基準 80–120 行/人時", "2026", "低於 60 或高於 150 都有；品類差異大"),
 # --- AMR 貨到人
 ("G", "AMR 貨到人工作站生產力", "訂單行/人時", 180, 250, 400, "[Analogy]", "[P]",
  "類比 Dematic（供應商說法）：貨架到人約為人工 2 倍；料箱到人最高 500 行/時", "未註明", "供應商行銷口徑，未註明條件"),
 ("R", "每工作站配置 AMR 數", "台/站", 3, 5, 8, "[Assumed]", "[P]", "未取得公開配置數據", "—", "決定每台分攤的節省工時"),
 ("P", "AMR 單價（含核心軟體與試車）", "USD/台", 20000, 30000, 45000, "[Interested-party]", "[A]",
  "Cleverence（倉儲軟體商部落格，非報價）：每台 2–4.5 萬美元", "2026-04-08", "獨立報價未取得；中美價差未取得"),
 ("s", "軟體與整合占專案總成本", "比例", 0.20, 0.275, 0.35, "[Interested-party]", "[A]",
  "Cleverence：中型與企業專案 20–35%", "2026-04-08", "含車隊調度軟體（一次性 5–25 萬美元）"),
 ("Cst", "揀貨工作站單價", "USD/站", 15000, 25000, 40000, "[Interested-party]", "[A]",
  "Cleverence：每站 1.5–4 萬美元", "2026-04-08", ""),
 ("Sa", "每台 AMR 年支援費", "USD/台/年", 1000, 2000, 3000, "[Interested-party]", "[A]",
  "Cleverence：每台每年 1,000–3,000 美元", "2026-04-08", ""),
 ("H", "年營運時數", "時/年", 2000, 4000, 6000, "[Assumed]", "[P]", "1–3 班 × 250 天 × 8 小時", "—", "直接放大節省工時"),
 ("U", "AMR 工作站穩態稼動率", "比例", 0.70, 0.85, 0.95, "[Assumed]", "[P]", "未取得", "—", ""),
 ("U1", "第一年爬坡（占穩態產能）", "比例", 0.40, 0.50, 0.60, "[Interested-party]", "[P]",
  "Conveyco：新系統第一年只跑 40–60% 產能", "2026", "AMR 與 AS/RS 共用"),
 ("r", "資金成本（年化）", "比例", 0.06, 0.08, 0.12, "[Assumed]", "[A]", "未取得客戶 WACC", "—", "用於每單位成本的資本回收係數"),
 ("La", "AMR／手臂經濟壽命", "年", 5, 7, 10, "[Assumed]", "[P]", "未取得", "—", ""),
 ("d", "AMR 年降價幅度", "比例/年", 0.03, 0.05, 0.08, "[Assumed]", "[A]", "未取得；中國廠商價格競爭方向已知、幅度未知", "—", "情境頁 2026–2032 使用"),
 ("g", "倉儲時薪年增率", "比例/年", 0.02, 0.035, 0.05, "[Assumed]", "[A]", "未取得 BLS 年增率（本輪未查）", "—", "情境頁使用"),
 # --- AS/RS 立方格
 ("Csys", "立方格 AS/RS 系統總價（典型）", "USD/系統", 3000000, 4500000, 6000000, "[Interested-party]", "[A]",
  "Kardex（AutoStore 整合商）：多數系統 300–600 萬美元；小型約 100 萬、大型逾 5,000 萬", "未註明", "視為全含（結構、料箱、機器人、埠、整合）"),
 ("Np", "每系統揀選埠數", "埠", 6, 10, 15, "[Assumed]", "[P]", "未取得", "—", "Csys ÷ Np ＝ 每埠全含資本支出"),
 ("Up", "AS/RS 埠稼動率", "比例", 0.40, 0.60, 0.80, "[Assumed]", "[P]", "未取得", "—", "埠有人且有料箱可揀的時間比例"),
 ("Mx", "AS/RS 埠人員生產力（相對人工倍數）", "倍", 3, 4, 5, "[Interested-party]", "[P]",
  "Dematic（供應商說法）：AutoStore 生產力提升 300%–500%", "未註明", "若讀為『增加』則為 4–6 倍；本版取 3–5 倍偏保守"),
 ("Oa", "AS/RS 年維運費占初始投資", "比例/年", 0.05, 0.10, 0.20, "[Interested-party]", "[A]",
  "Conveyco 5–8%；Cleverence 12–20%（傳統自動化）", "2026", "兩個利害關係人來源衝突，區間取聯集，基準取中間"),
 ("Ls", "AS/RS 經濟壽命", "年", 10, 12, 15, "[Assumed]", "[P]", "框架 v0.1 結構限制：系統壽命 10–15 年（未取得原始出處）", "—", ""),
 ("d2", "AS/RS 年降價幅度", "比例/年", 0.00, 0.02, 0.04, "[Assumed]", "[A]", "未取得", "—", "情境頁使用"),
 # --- 單件揀選手臂
 ("Ka", "單件揀選手臂工作站單價", "USD/站", 150000, 250000, 400000, "[Assumed]", "[A]",
  "未取得公開報價（本輪查無可信來源）", "—", "區間為假設，只用於情境比較，不是事實"),
 ("Ga", "手臂揀選速度", "件/時", 300, 450, 700, "[Assumed]", "[P]", "未取得獨立數據（門檻 2 的缺口）", "—", "手臂取代 AS/RS 埠的人員"),
 ("fa", "手臂站監看人力", "人/站", 0.10, 0.20, 0.33, "[Assumed]", "[P]", "未取得", "—", "例外處理與補揀"),
 ("Oarm", "手臂年維運費占單價", "比例/年", 0.08, 0.10, 0.15, "[Assumed]", "[A]", "未取得", "—", ""),
 # --- 人形
 ("Hraas", "人形機器人 RaaS 年費", "USD/台/年", 60000, 100000, 140000, "[Interested-party]", "[A]",
  "Agility 投資人簡報示意值（約 8,500 美元/月，轉引 beri.net）；低高為假設", "2026-10", "非公開牌價；GXO 費率未揭露"),
 ("Hh", "人形年工作時數", "時/年", 3000, 5000, 6000, "[Assumed]", "[P]", "未取得", "—", ""),
 ("th", "人形相對人工生產力", "比例", 0.30, 0.50, 0.80, "[Assumed]", "[P]",
  "未取得；GXO 執行長 2026-08-05：45 個人形試點尚未達 ROI", "—", "承重且最不確定"),
 ("dh", "人形 RaaS 年降價幅度", "比例/年", 0.05, 0.10, 0.20, "[Assumed]", "[A]", "未取得", "—", "情境頁使用"),
 # --- 市場
 ("A24", "行動機器人營收 2024", "十億 USD", 5.0, 5.0, 5.0, "[Verified]", "[P]",
  "Interact Analysis（轉引 Robotics 24/7、CFO Dive）：略低於 50 億美元", "2026-01-15",
  "口徑：製造與物流物料搬運，含 AGV、自主堆高機、訂單履行機器人"),
 ("A25g", "行動機器人營收成長 2025", "比例", 0.24, 0.24, 0.24, "[Verified]", "[P]",
  "Interact Analysis（轉引 SDCExec）：2025 營收 +24%、出貨 +18%", "2026-01-23", ""),
 ("A30", "行動機器人營收 2030", "十億 USD", 11.2, 14.0, 15.6, "[Verified]", "[P]",
  "基準：Interact 2026-01 預測 140 億；高：Interact 2025-05 前版 156 億；低：[Assumed] 基準 −20%（Interact 已下修 12%）",
  "2026-01-20", "基準與高為 Verified 預測值；低為假設"),
 ("Apost", "行動機器人 2031–32 年增", "比例/年", 0.08, 0.12, 0.15, "[Assumed]", "[P]", "未取得 2030 後預測", "—", ""),
 ("S25", "AS/RS 市場 2025", "十億 USD", 9.86, 9.86, 9.86, "[Verified]", "[P]",
  "MarketsandMarkets 新聞稿（研究機構，口徑不透明）：2025 年 98.6 億美元", "2025-01-31", "2025 為該報告預估值；含製造業與汽車"),
 ("S30", "AS/RS 市場 2030", "十億 USD", 13.2, 14.8, 16.5, "[Verified]", "[P]",
  "基準：MarketsandMarkets 148 億（CAGR 8.5%）；低：[Analogy] Interact 整體倉儲自動化 2025–30 年增 6%；高：[Assumed]",
  "2025-01-31", ""),
 ("Spost", "AS/RS 2031–32 年增", "比例/年", 0.04, 0.07, 0.09, "[Assumed]", "[P]", "未取得", "—", ""),
 # --- 對照
 ("AS26", "AutoStore 2026 營收指引", "十億 USD", 0.75, 0.75, 0.75, "[Interested-party]", "[I]",
  "AutoStore Q3 2026 報告（公司指引，由 7 億上調至約 7.5 億美元）", "2026-11-05", "毛利 71.3%、調整後 EBITDA 率 42.3%"),
 ("GK25", "極智嘉 2025 營收", "十億 RMB", 3.171, 3.171, 3.171, "[Verified]", "[I]",
  "Geekplus 2025 年度業績新聞稿（港交所 2590）", "2026-03-31", "毛利率 35.5%；海外占 75.3%、海外毛利 46.6%"),
 ("FX", "美元兌人民幣", "RMB/USD", 7.0, 7.15, 7.3, "[Assumed]", "[A]", "本輪未查證匯率", "—", "只用於份額對照"),
 ("IFR25", "IFR 運輸物流專業服務機器人 2025 銷量", "台", 117500, 117500, 117500, "[Verified]", "[P]",
  "IFR World Robotics 2026（轉引 Malay Mail）：+21%，樣本 238 家", "2026-09-30", "IFR 為樣本統計，與 Interact 口徑不同"),
 ("AMZ", "Amazon 自動化每件節省", "USD/件", 0.30, 0.30, 0.30, "[Interested-party]", "[A]",
  "Amazon 內部文件經 NYT 揭露（轉引 Computerworld）；公司稱不代表完整策略", "2025-10", "基準口徑未知"),
 ("BP", "Brightpick 自主揀貨 AMR RaaS 月費", "USD/台/月", 2190, 2190, 2190, "[Interested-party]", "[A]",
  "Brightpick 官網定價：起價，3 年約、最少 15 台", "2026-11-16 查", "用於 RaaS 對照"),
]
IN = {x[0]: x for x in INPUTS}

# ---------------------------------------------------------------- 計算（框架第 3 節等式）
# (代號, 項目, 單位, 運算式, 說明)
CALC = [
 ("SEC", "A. 人工基準"),
 ("W", "全負擔時薪", "USD/時", "{W0}*{LF}", "＝時薪 × 全負擔係數"),
 ("SEC", "B. 每台 AMR 回本（貨到人）"),
 ("amr_cap", "每台全含資本支出", "USD/台", "{P}/(1-{s})+{Cst}/{R}", "＝單價 ÷（1 − 整合占比）＋ 工作站 ÷ 每站台數"),
 ("amr_hrs", "每台年取代工時（穩態）", "時/台/年", "({G}/{M}-1)*{H}*{U}/{R}", "＝（貨到人生產力 ÷ 人工生產力 − 1）× 年時數 × 稼動率 ÷ 每站台數"),
 ("amr_sav", "每台年淨節省（穩態）", "USD/台/年", "{amr_hrs}*{W}-{Sa}", "＝取代工時 × 全負擔時薪 − 年支援費"),
 ("amr_pb", "每台 AMR 回本期（穩態）", "年", "{amr_cap}/{amr_sav}", "＝全含資本支出 ÷ 年淨節省"),
 ("amr_sav1", "第一年淨節省（爬坡）", "USD/台/年", "{amr_hrs}*{U1}*{W}-{Sa}", "＝取代工時 × 爬坡比例 × 時薪 − 支援費"),
 ("amr_pb1", "每台 AMR 回本期（含第一年爬坡）", "年", "IF({amr_cap}<={amr_sav1},{amr_cap}/{amr_sav1},1+({amr_cap}-{amr_sav1})/{amr_sav})", "第一年用爬坡節省，之後用穩態"),
 ("amr_raas_be", "供應商 RaaS 對照：Brightpick 月費 ÷ 本模型每台每月攤提", "倍", "({BP}*12)/({amr_cap}*{crf_a}+{Sa})", "＞1 表示 RaaS 價格高於自購年化成本（含風險轉移與服務）"),
 ("SEC", "C. 每座倉／每埠回本（立方格 AS/RS）"),
 ("Kp", "每埠全含資本支出", "USD/埠", "{Csys}/{Np}", "＝系統總價 ÷ 揀選埠數"),
 ("asrs_hrs", "每埠年取代工時（穩態）", "時/埠/年", "{H}*{Up}*({Mx}-1)", "＝年時數 × 埠稼動率 ×（生產力倍數 − 1）"),
 ("asrs_sav", "每埠年淨節省（穩態）", "USD/埠/年", "{asrs_hrs}*{W}-{Oa}*{Kp}", "＝取代工時 × 時薪 − 年維運費"),
 ("asrs_pb", "AS/RS 回本期（穩態；每埠＝每座倉，線性）", "年", "{Kp}/{asrs_sav}", "＝每埠資本支出 ÷ 年淨節省"),
 ("asrs_sav1", "第一年淨節省（爬坡）", "USD/埠/年", "{asrs_hrs}*{U1}*{W}-{Oa}*{Kp}", ""),
 ("asrs_pb1", "AS/RS 回本期（含第一年爬坡）", "年", "IF({Kp}<={asrs_sav1},{Kp}/{asrs_sav1},1+({Kp}-{asrs_sav1})/{asrs_sav})", ""),
 ("wh_cap", "每座倉 AS/RS 投資（典型）", "USD", "{Kp}*{Np}", "＝系統總價（對照用）"),
 ("wh_fte", "每座倉取代全職人力（穩態）", "人", "{asrs_hrs}*{Np}/2000", "以每人每年 2,000 工時換算"),
 ("SEC", "D. 每單位（訂單行）處理成本"),
 ("crf_a", "資本回收係數（AMR／手臂）", "比例/年", "{r}/(1-POWER(1+{r},-{La}))", "＝r ÷（1 −（1+r）^−壽命）"),
 ("crf_s", "資本回收係數（AS/RS）", "比例/年", "{r}/(1-POWER(1+{r},-{Ls}))", ""),
 ("c_man", "人工揀貨每行成本", "USD/行", "{W}/{M}", "＝全負擔時薪 ÷ 人工生產力"),
 ("c_amr", "AMR 貨到人每行成本", "USD/行", "({W}+({R}*{P}/(1-{s})+{Cst})*{crf_a}/({H}*{U})+{R}*{Sa}/({H}*{U}))/{G}", "＝（人員時薪 ＋ 每站時資本成本 ＋ 每站時支援費）÷ 每站時行數"),
 ("c_asrs", "AS/RS（立方格）每行成本", "USD/行", "({W}+{Kp}*({crf_s}+{Oa})/({H}*{Up}))/({M}*{Mx})", "＝（人員時薪 ＋ 每埠時資本與維運成本）÷ 每埠時行數"),
 ("c_arm", "AS/RS＋單件揀選手臂每行成本", "USD/行", "({fa}*{W}+({Kp}*({crf_s}+{Oa})+{Ka}*({crf_a}+{Oarm}))/({H}*{Up}))/{Ga}", "手臂取代埠上人員；監看人力按比例計"),
 ("sv_amr", "AMR 每行節省（相對人工）", "USD/行", "{c_man}-{c_amr}", "對照 Amazon 約 0.30 美元/件（口徑不同）"),
 ("sv_asrs", "AS/RS 每行節省（相對人工）", "USD/行", "{c_man}-{c_asrs}", ""),
 ("sv_arm", "手臂相對 AS/RS 人工埠的每行節省", "USD/行", "{c_asrs}-{c_arm}", "＞0 代表手臂已比人員埠便宜（門檻 2 的成本面）"),
 ("SEC", "E. 人形／移動操作（門檻 4）"),
 ("hum_hr", "人形每人工當量小時成本", "USD/時", "{Hraas}/({Hh}*{th})", "＝RaaS 年費 ÷（年工時 × 相對生產力）"),
 ("hum_gap", "人形成本 − 全負擔時薪", "USD/時", "{hum_hr}-{W}", "＜0 代表跨過門檻 4"),
 ("hum_be_th", "打平所需相對生產力", "比例", "{Hraas}/({Hh}*{W})", "＝RaaS 年費 ÷（年工時 × 時薪）"),
 ("SEC", "F. 市場與份額對照"),
 ("A25", "行動機器人營收 2025", "十億 USD", "{A24}*(1+{A25g})", "＝2024 × （1＋2025 成長）[Derived]"),
 ("cagr_amr", "行動機器人 2025–30 CAGR（基準）", "比例/年", "POWER({A30}/{A25},1/5)-1", ""),
 ("cagr_asrs", "AS/RS 2025–30 CAGR（基準）", "比例/年", "POWER({S30}/{S25},1/5)-1", ""),
 ("gk_share", "極智嘉占行動機器人營收（2025）", "比例", "{GK25}/{FX}/{A25}", "口徑不同（極智嘉含軟體與服務、Interact 含製造業），僅量級對照"),
 ("as_share", "AutoStore 占 AS/RS 市場（2026）", "比例", "{AS26}/({S25}*POWER({S30}/{S25},1/5))", "AutoStore 營收為出廠價（不含整合商加價），低估其系統占比"),
 ("asp_impl", "隱含每台行動機器人營收（Interact 營收 ÷ IFR 台數）", "USD/台", "{A25}*1000000000/{IFR25}", "口徑不一致（IFR 只含服務機器人樣本、Interact 含堆高機），只看量級"),
]
CALC_ROWS = [c for c in CALC if c[0] != "SEC"]
CD = {c[0]: c for c in CALC_ROWS}

FUNCS = {"POWER": lambda a, b: a ** b, "IF": lambda c, a, b: a if c else b,
         "MAX": max, "MIN": min, "ABS": abs}
TOK = re.compile(r"\{(\w+)\}")
COL = {"low": "D", "base": "E", "high": "F"}
IDX = {"low": 3, "base": 4, "high": 5}

in_row, calc_row = {}, {}


def expand(expr, choose, over=None, mode="xl"):
    """把運算式展開到只剩輸入。choose(id)->'low'/'base'/'high'；over：id→已展開的替代運算式（Excel 與 Python 共用語法）。"""
    over = over or {}

    def rep(m):
        k = m.group(1)
        if k in over:
            return "(" + over[k][mode] + ")"
        if k in IN:
            c = choose(k)
            if mode == "xl":
                return f"'輸入'!${COL[c]}${in_row[k]}"
            return repr(float(IN[k][IDX[c]]))
        return "(" + expand(CD[k][3], choose, over, mode) + ")"
    return TOK.sub(rep, expr)


def pyeval(expr):
    return eval(expr, {"__builtins__": {}}, FUNCS)


def deps(expr, acc=None):
    acc = set() if acc is None else acc
    for k in TOK.findall(expr):
        if k in IN:
            acc.add(k)
        else:
            deps(CD[k][3], acc)
    return acc


def val(cid, choose=lambda k: "base", over=None):
    return pyeval(expand(CD[cid][3], choose, over, "py"))


def year_over(t, choose):
    """隨年份變動的覆寫：AMR 價格、時薪、AS/RS 系統價、人形 RaaS。"""
    n = t - 2026
    o = {}
    for key, base, drv, sign in (("P", "{P}", "d", -1), ("W0", "{W0}", "g", 1),
                                 ("Csys", "{Csys}", "d2", -1), ("Hraas", "{Hraas}", "dh", -1)):
        e = f"{base}*POWER(1{'+' if sign > 0 else '-'}{{{drv}}},{n})"
        o[key] = {"xl": expand(e, choose), "py": expand(e, choose, mode="py")}
    return o


def worst_choice(cid, worse_is_higher=True):
    """悲觀＝敏感度前 TOPK 名承重變數同時取不利端（其餘基準）＋年度驅動取不利端；樂觀相反。"""
    pess, opt = {}, {}
    sw = []
    for k in deps(CD[cid][3]) - SCEN_FIXED:
        lo = val(cid, lambda x: "low" if x == k else "base")
        hi = val(cid, lambda x: "high" if x == k else "base")
        if abs(hi - lo) > 1e-12:
            sw.append((abs(hi - lo), k, hi > lo))
    sw.sort(reverse=True)
    for _, k, hi_gt in sw[:TOPK]:  # 只動敏感度前 TOPK 名承重變數，其餘維持基準
        hi_worse = hi_gt == worse_is_higher
        pess[k], opt[k] = ("high", "low") if hi_worse else ("low", "high")
    # 時間驅動變數：方向直接定義
    for k, p in (("d", "low"), ("d2", "low"), ("dh", "low")):
        pess[k], opt[k] = p, ("high" if p == "low" else "low")
    # 時薪成長：對自動化而言，時薪漲得慢＝悲觀
    pess["g"], opt["g"] = "low", "high"
    return (lambda k: pess.get(k, "base")), (lambda k: opt.get(k, "base"))


def setw(ws, widths):
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w


def hdr(ws, r, vals):
    for j, v in enumerate(vals, 1):
        c = ws.cell(r, j, v)
        c.font = BOLD
        c.fill = HDR
        c.alignment = WRAP


def build(path):
    wb = openpyxl.Workbook()
    ws0 = wb.active
    ws0.title = "說明"
    wi = wb.create_sheet("輸入")
    wc = wb.create_sheet("計算")
    wsn = wb.create_sheet("情境")
    wse = wb.create_sheet("敏感度")
    wd = wb.create_sheet("下游取數")
    wt = wb.create_sheet("追蹤指標")

    # ---------------- 輸入
    wi["A1"] = "輸入｜藍字＝輸入；黃底＝判斷值（可改）；計算頁一律讀『基準』欄，低／高欄供情境與敏感度"
    wi["A1"].font = TITLE
    hdr(wi, 3, ["代號", "變數", "單位", "低", "基準", "高", "標記", "類別", "出處", "日期", "備註"])
    r = 4
    for x in INPUTS:
        in_row[x[0]] = r
        for j, v in enumerate(x, 1):
            c = wi.cell(r, j, v)
            c.alignment = WRAP
        judg = x[6] != "[Verified]" or not (x[3] == x[4] == x[5])
        for j in (4, 5, 6):
            c = wi.cell(r, j)
            c.font = BLUE
            if judg:
                c.fill = YEL
        r += 1
    setw(wi, [8, 34, 13, 11, 11, 11, 17, 7, 60, 13, 40])
    wi.freeze_panes = "D4"

    # ---------------- 計算
    wc["A1"] = "計算｜框架第 3 節等式（全部讀『輸入』基準欄；黑字＝公式）"
    wc["A1"].font = TITLE
    hdr(wc, 3, ["代號", "項目", "值", "單位", "等式", "標記"])
    r = 4
    for c in CALC:
        if c[0] == "SEC":
            wc.cell(r, 1, c[1]).font = BOLD
            wc.cell(r, 1).fill = GRY
            r += 1
            continue
        calc_row[c[0]] = r
        r += 1
    for c in CALC_ROWS:
        rr = calc_row[c[0]]

        def rep(m):
            k = m.group(1)
            return f"'輸入'!$E${in_row[k]}" if k in IN else f"$C${calc_row[k]}"
        wc.cell(rr, 1, c[0])
        wc.cell(rr, 2, c[1])
        wc.cell(rr, 3, "=" + TOK.sub(rep, c[3]))
        wc.cell(rr, 4, c[2])
        wc.cell(rr, 5, c[4]).alignment = WRAP
        wc.cell(rr, 6, "[Derived]")
        fmt = "0.000" if c[2].startswith("USD/行") or c[2] in ("比例", "比例/年", "倍") else ("0.00" if c[2] in ("年", "USD/時") else "#,##0")
        if c[2] in ("比例", "比例/年"):
            fmt = "0.0%"
        if c[2] == "十億 USD":
            fmt = "0.00"
        wc.cell(rr, 3).number_format = fmt
    setw(wc, [11, 46, 14, 12, 70, 10])
    wc.freeze_panes = "C4"

    # ---------------- 情境
    wsn["A1"] = "情境｜2026–2032，低／基準／高（市場）與悲觀／基準／樂觀（單位經濟；悲觀／樂觀＝該輸出敏感度前 3 名承重變數同時取對自動化不利／有利端，加上年度驅動；其餘輸入取基準）"
    wsn["A1"].font = TITLE
    r = 3
    scen_ref = {}
    # 市場
    for key, lab, s25, s30, post in (("amr", "行動機器人（AMR／AGV／自主堆高機）營收，十億 USD（Interact 口徑）", None, "A30", "Apost"),
                                      ("asrs", "AS/RS 市場，十億 USD（MarketsandMarkets 口徑）", "S25", "S30", "Spost")):
        wsn.cell(r, 1, lab).font = BOLD
        r += 1
        hdr(wsn, r, ["情境", "2025"] + [str(y) for y in YEARS] + ["2025–30 CAGR"])
        r += 1
        for sc, nm in (("low", "低"), ("base", "基準"), ("high", "高")):
            wsn.cell(r, 1, nm)
            if key == "amr":
                b = f"'計算'!$C${calc_row['A25']}"
            else:
                b = f"'輸入'!$E${in_row['S25']}"
            wsn.cell(r, 2, "=" + b)
            end = f"'輸入'!${COL[sc]}${in_row[s30]}"
            for j, y in enumerate(YEARS):
                col = 3 + j
                if y <= 2030:
                    f = f"=$B{r}*POWER({end}/$B{r},{y - 2025}/5)"
                else:
                    f = f"={get_column_letter(col - 1)}{r}*(1+'輸入'!${COL[sc]}${in_row[post]})"
                wsn.cell(r, col, f).number_format = "0.00"
                scen_ref[(key, sc, y)] = f"'情境'!${get_column_letter(col)}${r}"
            wsn.cell(r, 2).number_format = "0.00"
            c = wsn.cell(r, 3 + len(YEARS), f"=POWER({end}/$B{r},1/5)-1")
            c.number_format = "0.0%"
            r += 1
        r += 1
    # 單位經濟路徑
    paths = [("amr_pb", "每台 AMR 回本期（穩態，年）", True, "0.00"),
             ("amr_pb1", "每台 AMR 回本期（含第一年爬坡，年）", True, "0.00"),
             ("asrs_pb", "AS/RS 回本期（穩態，年）", True, "0.00"),
             ("c_man", "人工揀貨每行成本（USD/行；悲觀＝人工較便宜）", False, "0.000"),
             ("c_amr", "AMR 貨到人每行成本（USD/行）", True, "0.000"),
             ("c_asrs", "AS/RS 每行成本（USD/行）", True, "0.000"),
             ("c_arm", "AS/RS＋手臂每行成本（USD/行）", True, "0.000"),
             ("hum_hr", "人形每人工當量小時成本（USD/時）", True, "0.00"),
             ("W", "全負擔時薪（USD/時；悲觀＝較低）", False, "0.00")]
    checks = []
    wsn.cell(r, 1, "單位經濟路徑（年度驅動：AMR 降價 d、時薪成長 g、AS/RS 降價 d2、人形 RaaS 降價 dh；其餘輸入不隨年變）").font = BOLD
    r += 1
    for cid, lab, wih, fmt in paths:
        wsn.cell(r, 1, lab).font = BOLD
        r += 1
        hdr(wsn, r, ["情境", ""] + [str(y) for y in YEARS])
        r += 1
        pess, opt = worst_choice(cid, wih)
        moved = [k for k in sorted(IN) if pess(k) != "base" and k not in ("d", "g", "d2", "dh")]
        wsn.cell(r - 2, 2, "悲觀／樂觀移動：" + "、".join(f"{k}（{IN[k][1]}）" for k in moved) + "＋年度驅動")
        sets = [("悲觀", pess), ("基準", lambda k: "base"), ("樂觀", opt)]
        for nm, ch in sets:
            wsn.cell(r, 1, nm)
            for j, y in enumerate(YEARS):
                ov = year_over(y, ch)
                xl = expand(CD[cid][3], ch, ov, "xl")
                py = pyeval(expand(CD[cid][3], ch, ov, "py"))
                c = wsn.cell(r, 3 + j, "=" + xl)
                c.number_format = fmt
                scen_ref[(cid, nm, y)] = f"'情境'!${get_column_letter(3 + j)}${r}"
                checks.append(("情境", r, 3 + j, py))
            r += 1
        r += 1
    setw(wsn, [36, 9] + [10] * len(YEARS) + [12])

    # ---------------- 敏感度
    wse["A1"] = "敏感度｜每個輸入單獨取低／高、其餘取基準；依擺幅排序（承重變數在前，基準結果列在排序之後）"
    wse["A1"].font = TITLE
    r = 3
    sens_top = {}
    for cid, lab, fmt in (("amr_pb", "每台 AMR 回本期（穩態，年）", "0.00"),
                          ("asrs_pb", "AS/RS 回本期（穩態，年）", "0.00"),
                          ("c_amr", "AMR 貨到人每行成本（USD/行）", "0.000"),
                          ("c_asrs", "AS/RS 每行成本（USD/行）", "0.000"),
                          ("hum_gap", "人形成本 − 全負擔時薪（USD/時）", "0.00")):
        wse.cell(r, 1, lab).font = BOLD
        r += 1
        hdr(wse, r, ["排序", "代號", "變數", "輸入低", "輸入高", "結果（輸入取低）", "結果（輸入取高）", "擺幅", "標記"])
        r += 1
        rows = []
        for k in deps(CD[cid][3]):
            lo = val(cid, lambda x, k=k: "low" if x == k else "base")
            hi = val(cid, lambda x, k=k: "high" if x == k else "base")
            if abs(hi - lo) > 1e-12:
                rows.append((abs(hi - lo), k, lo, hi))
        rows.sort(reverse=True)
        sens_top[cid] = rows[:3]
        first = r
        for i, (sw, k, lo, hi) in enumerate(rows, 1):
            wse.cell(r, 1, i)
            wse.cell(r, 2, k)
            wse.cell(r, 3, IN[k][1])
            wse.cell(r, 4, f"='輸入'!$D${in_row[k]}")
            wse.cell(r, 5, f"='輸入'!$F${in_row[k]}")
            fl = "=" + expand(CD[cid][3], lambda x, k=k: "low" if x == k else "base")
            fh = "=" + expand(CD[cid][3], lambda x, k=k: "high" if x == k else "base")
            wse.cell(r, 6, fl).number_format = fmt
            wse.cell(r, 7, fh).number_format = fmt
            wse.cell(r, 8, f"=ABS(G{r}-F{r})").number_format = fmt
            wse.cell(r, 9, IN[k][6])
            checks.append(("敏感度", r, 6, lo))
            checks.append(("敏感度", r, 7, hi))
            r += 1
        wse.cell(r, 3, "基準結果").font = BOLD
        c = wse.cell(r, 6, f"='計算'!$C${calc_row[cid]}")
        c.number_format = fmt
        c.font = BOLD
        wse.cell(r, 8, f"=IF(MAX(H{first}:H{r - 1})=H{first},\"排序一致\",\"排序需重建\")")
        r += 2
    setw(wse, [6, 8, 36, 11, 11, 15, 15, 11, 18])

    # ---------------- 下游取數
    wd["A1"] = "下游取數｜給個股模型讀（AutoStore、Symbotic、大福、極智嘉、KION、Amazon 等）；值皆為公式連結，勿手改"
    wd["A1"].font = TITLE
    hdr(wd, 3, ["代號", "名稱", "低／悲觀", "基準", "高／樂觀", "單位", "口徑", "來源儲存格（基準）", "標記"])
    r = 4
    DS = []
    for y in (2026, 2028, 2030, 2032):
        DS.append((f"WH_AMR_MKT_{y}", f"行動機器人營收 {y}", scen_ref[("amr", "low", y)], scen_ref[("amr", "base", y)],
                   scen_ref[("amr", "high", y)], "十億 USD", "Interact 口徑：製造＋物流物料搬運，含 AGV、自主堆高機、訂單履行機器人；2031 起為假設成長", "[Derived]"))
    for y in (2026, 2028, 2030, 2032):
        DS.append((f"WH_ASRS_MKT_{y}", f"AS/RS 市場 {y}", scen_ref[("asrs", "low", y)], scen_ref[("asrs", "base", y)],
                   scen_ref[("asrs", "high", y)], "十億 USD", "MarketsandMarkets 口徑（含製造、汽車）；低端類比 Interact 6%", "[Derived]"))

    def cref(cid):
        return f"'計算'!$C${calc_row[cid]}"
    for cid, code, nm, unit, kj in (
        ("amr_pb", "WH_AMR_PAYBACK", "每台 AMR 回本期（穩態）", "年", "美國、貨到人、2026；悲觀／樂觀＝前 3 名承重變數同時取不利／有利端（見情境頁）"),
        ("amr_pb1", "WH_AMR_PAYBACK_RAMP", "每台 AMR 回本期（含第一年爬坡）", "年", "同上，第一年產能 40–60%"),
        ("asrs_pb", "WH_ASRS_PAYBACK", "AS/RS（立方格）回本期（穩態）", "年", "每埠＝每座倉（線性）；美國 2026"),
        ("c_man", "WH_COST_MANUAL", "人工揀貨每行成本", "USD/訂單行", "美國全負擔時薪 ÷ 人工生產力"),
        ("c_amr", "WH_COST_AMR", "AMR 貨到人每行成本", "USD/訂單行", "含資本回收（r、壽命）與支援費"),
        ("c_asrs", "WH_COST_ASRS", "AS/RS 每行成本", "USD/訂單行", "含資本回收與維運費"),
        ("c_arm", "WH_COST_ARM", "AS/RS＋單件揀選手臂每行成本", "USD/訂單行", "手臂價格與速度皆為假設（未取得），低信心"),
        ("hum_hr", "WH_HUMANOID_COST_HR", "人形每人工當量小時成本", "USD/時", "Agility 示意 RaaS ÷（工時 × 相對生產力）；生產力為假設"),
    ):
        DS.append((code, nm, scen_ref[(cid, "悲觀", 2026)], cref(cid), scen_ref[(cid, "樂觀", 2026)], unit, kj, "[Derived]"))
    for cid, code, nm, unit, kj in (
        ("amr_pb", "WH_AMR_PAYBACK_2030", "每台 AMR 回本期（穩態）2030", "年", "AMR 降價 d、時薪成長 g 推至 2030"),
        ("c_amr", "WH_COST_AMR_2030", "AMR 貨到人每行成本 2030", "USD/訂單行", "同上"),
        ("hum_hr", "WH_HUMANOID_COST_HR_2030", "人形每人工當量小時成本 2030", "USD/時", "RaaS 年降價 dh"),
    ):
        DS.append((code, nm, scen_ref[(cid, "悲觀", 2030)], scen_ref[(cid, "基準", 2030)], scen_ref[(cid, "樂觀", 2030)], unit, kj, "[Derived]"))
    DS.append(("WH_W_LOADED_US", "美國倉儲全負擔時薪 2026", f"'輸入'!$E${in_row['W0']}*'輸入'!$D${in_row['LF']}", cref("W"),
               f"'輸入'!$E${in_row['W0']}*'輸入'!$F${in_row['LF']}", "USD/時", "BLS NAICS 493 生產與非主管 × 全負擔係數", "[Derived]"))
    DS.append(("WH_AMR_CAPEX_UNIT", "每台 AMR 全含資本支出", None, cref("amr_cap"), None, "USD/台", "單價＋整合分攤＋工作站分攤", "[Derived]"))
    DS.append(("WH_ASRS_CAPEX_PORT", "AS/RS 每埠全含資本支出", None, cref("Kp"), None, "USD/埠", "系統總價 ÷ 埠數", "[Derived]"))
    DS.append(("WH_GEEKPLUS_SHARE", "極智嘉占行動機器人營收 2025", None, cref("gk_share"), None, "比例", "口徑不同，量級對照", "[Derived]"))
    DS.append(("WH_AUTOSTORE_SHARE", "AutoStore 占 AS/RS 市場 2026", None, cref("as_share"), None, "比例", "出廠價口徑，低估系統占比", "[Derived]"))
    DS.append(("WH_HUM_BREAKEVEN_TH", "人形打平所需相對人工生產力", None, cref("hum_be_th"), None, "比例", "基準 RaaS 與工時下", "[Derived]"))
    for code, nm, lo, b, hi, unit, kj, tag in DS:
        wd.cell(r, 1, code)
        wd.cell(r, 2, nm)
        for j, v in ((3, lo), (4, b), (5, hi)):
            if v:
                c = wd.cell(r, j, "=" + v)
                c.number_format = "0.000" if "USD/訂單行" in unit else ("0.0%" if unit == "比例" else ("#,##0" if unit.startswith("USD/台") or unit.startswith("USD/埠") else "0.00"))
        wd.cell(r, 6, unit)
        wd.cell(r, 7, kj).alignment = WRAP
        wd.cell(r, 8, b)
        wd.cell(r, 9, tag)
        r += 1
    setw(wd, [26, 36, 12, 12, 12, 13, 60, 22, 10])
    wd.freeze_panes = "C4"

    # ---------------- 追蹤指標
    wt["A1"] = "追蹤指標｜框架第 5 節；最新值與日期（本輪未更新者照框架 v0.1 並註明）"
    wt["A1"].font = TITLE
    hdr(wt, 3, ["指標", "頻率", "資料來源", "最新值", "日期", "標記", "為何重要／連到哪個輸入"])
    TRK = [
     ("運輸物流機器人台數與成長", "年（9 月）", "IFR World Robotics", "2025 年 11.75 萬台，+21%，占專業服務機器人 47%", "2026-09-30", "[Verified]", "產業總量；IFR25"),
     ("行動機器人營收與成長", "年", "Interact Analysis", "2024 略低於 50 億美元；2025 營收 +24%、出貨 +18%；2030 預測 140 億（前版 156 億）", "2026-01", "[Verified]", "A24、A25g、A30"),
     ("倉儲自動化接單", "年", "Interact Analysis", "2025 接單 +7%（美國 +12%，其中 6 點為漲價）；2025–30 全球年增 6%；亞太 2025 營收 −8%", "2026-02-24", "[Verified]", "S30 低端類比"),
     ("美國倉儲業時薪", "月", "BLS CES NAICS 493", "生產與非主管 26.02 美元；全體 26.87 美元；就業 183.95 萬人", "2026-09（初值）", "[Verified]", "W0"),
     ("Symbotic backlog、12 月交付比例、運作中系統數", "季", "Symbotic 10-Q／法說", "Backlog 225 億美元；12 月交付約 12%；Q3 FY26 運作 56 套、部署中 77 套（框架 v0.1）", "2026-08", "[Verified]", "客戶集中、backlog 轉換（未進本模型計算）"),
     ("AutoStore 營收、訂單、毛利", "季", "AutoStore 季報", "Q3 營收 1.921 億美元（+38.2%）、接單 1.776 億（+16.5%）、毛利 71.3%、調整後 EBITDA 42.3%；全年指引約 7.5 億；約 2,050 套系統", "2026-11-05", "[Verified]", "AS26；標準化 IP 毛利"),
     ("KION 智慧自動化訂單與 EBIT 率", "季", "KION 季報", "9M 2026 調整後 EBIT 率 6.8%；訂單 −12.4%（框架 v0.1，本輪未更新）", "2026-10", "[Verified]", "整合商利潤（價值歸屬）"),
     ("極智嘉營收、訂單、訂閱、海外毛利", "半年", "港交所公告（2590）", "2025 營收 31.71 億人民幣（+31.6%）、毛利 35.5%、海外毛利 46.6%、調整後淨利 4,380 萬；接單 41.37 億（+31.7%）；累計逾 7.2 萬台；H1 2026 訂閱訂單 1.56 億（+75%）", "2026-03-31／2026-08", "[Verified]", "GK25；亞太掛牌"),
     ("大福產業別營收", "季", "大福決算短信", "1H 2026 電子（含半導體 AMHS）42.6%、運輸倉儲 4.5%、商業零售 16.9%（框架 v0.1）", "2026-08", "[Verified]", "倉儲曝險權重"),
     ("Amazon 機器人數、每件 cost to serve", "季", "Amazon 法說與部落格", "逾 100 萬台、300+ 設施；自動化每件節省約 0.30 美元（內部文件）", "2025-10／2026", "[Interested-party]", "AMZ 對照"),
     ("Ocado CFC 上線與關閉", "半年", "Ocado 半年報", "Kroger 關 3 座 CFC、Sobeys 關 1 座（框架 v0.1）", "2026", "[Verified]", "大型雜貨倉經濟性（門檻 3）"),
     ("人形倉儲部署站點與時數", "不定期", "Agility、GXO 公告與法說", "GXO 45 個人形試點、尚未達 ROI（2026-08-05）；Digit 累計逾 10 萬箱、全部客戶逾 6.5 萬小時；2027 初擴至 2 座倉", "2026-10-26", "[Interested-party]", "th、Hraas（門檻 4）"),
     ("AMR RaaS 牌價", "不定期", "供應商官網", "Brightpick 每台每月 2,190 美元起（3 年約、最少 15 台）", "2026-11-16 查", "[Interested-party]", "BP"),
     ("MHI 年度調查（自動化投資意向）", "年", "MHI Annual Industry Report", "未取得（本輪未查）", "—", "—", "需求端領先指標"),
     ("AMR／AS/RS 獨立報價", "不定期", "第三方採購資料", "未取得", "—", "—", "P、Csys 是承重變數，目前只有利害關係人來源"),
    ]
    r = 4
    for t in TRK:
        for j, v in enumerate(t, 1):
            wt.cell(r, j, v).alignment = WRAP
        r += 1
    setw(wt, [34, 10, 26, 70, 16, 17, 30])

    # ---------------- 說明
    L = [
     ("線 3 倉儲物流自動化｜單位經濟模型 " + VERSION, TITLE),
     ("產生：2026-11-16｜框架：線3_倉儲物流_框架_v0.1.md｜產生器：physical_ai/lines/warehouse/build_warehouse.py", None),
     ("", None),
     ("核心問題：哪一層倉儲自動化（AMR 貨到人、立方格 AS/RS、單件揀選手臂、人形）的回本期與每單位處理成本已低於客戶門檻？", BOLD),
     ("分析單位：(a) 每台 AMR 回本期；(b) 每座倉／每埠 AS/RS 回本期；(c) 每訂單行處理成本（人工、AMR、AS/RS、AS/RS＋手臂）；(d) 人形每人工當量小時成本。基準：美國人工揀貨倉。", None),
     ("", None),
     ("頁面：輸入（藍字、黃底＝判斷）→ 計算（框架第 3 節等式）→ 情境（市場 2025–2032 低／基準／高；單位經濟 2026–2032 悲觀／基準／樂觀）→ 敏感度（承重變數排序）→ 下游取數 → 追蹤指標。", None),
     ("標記：[Verified] 可公開驗證；[Interested-party] 利害關係人（供應商、整合商）說法；[Analogy] 類比；[Assumed] 假設；[Derived] 由前者計算。Analogy、Assumed 一律給低／基準／高。", None),
     ("情境定義：市場的低／高＝輸入的低／高欄；單位經濟的悲觀／樂觀＝該輸出敏感度前 3 名承重變數同時取對自動化不利／有利的一端，加上年度驅動（降價、時薪成長），其餘取基準；年營運時數 H 屬場址設定（班次），情境固定在 2 班（4,000 時），其影響只在敏感度頁呈現。這是包絡，不是機率區間。", None),
     ("", None),
     ("重要限制（v0.1）", BOLD),
     ("1. AMR 單價、AS/RS 系統價、維運費率全部來自利害關係人（Cleverence、Kardex、Conveyco、Dematic），獨立報價未取得。維運費兩來源衝突（5–8% vs 12–20%），區間取聯集。", None),
     ("2. 單件揀選手臂的價格與速度、人形的相對生產力皆為假設（未取得），相關輸出低信心。", None),
     ("3. 只做美國人工基準；中國與亞太時薪未取得。", None),
     ("4. 市場口徑：行動機器人（Interact）含製造業與自主堆高機；AS/RS（MarketsandMarkets）含製造與汽車，2025 為 2025-01 的預估。兩者不可相加成『倉儲自動化市場』。", None),
     ("5. 未建模：空間節省（AutoStore 宣稱省 25% 面積）、電力、保險、勞動法規；Symbotic／Walmart 客戶集中（只在追蹤頁）。", None),
     ("6. 本模型不含個人情境；不放部位。", None),
    ]
    for i, (t, f) in enumerate(L, 1):
        c = ws0.cell(i, 1, t)
        c.alignment = Alignment(wrap_text=True, vertical="top")
        if f:
            c.font = f
    ws0.column_dimensions["A"].width = 140

    wb.save(path)
    return checks, sens_top


def recalc(src, outdir):
    prof = tempfile.mkdtemp(prefix="lo_wh_")
    cmd = ["soffice", f"-env:UserInstallation=file://{prof}", "--headless", "--calc",
           "--convert-to", "xlsx:Calc MS Excel 2007 XML", "--outdir", outdir, src]
    subprocess.run(cmd, check=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, timeout=180)
    shutil.rmtree(prof, ignore_errors=True)


def verify(path, checks):
    wb = openpyxl.load_workbook(path, data_only=True)
    errs = []
    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for c in row:
                if isinstance(c.value, str) and c.value.startswith("#") and c.value[1:4].isupper():
                    errs.append(f"{ws.title}!{c.coordinate}={c.value}")
    # 計算頁：Excel vs Python
    wc = wb["計算"]
    bad = []
    for cid in CD:
        xv = wc.cell(calc_row[cid], 3).value
        pv = val(cid)
        if xv is None or abs(xv - pv) > 1e-6 * max(1, abs(pv)):
            bad.append((cid, xv, pv))
    for sh, r, c, pv in checks:
        xv = wb[sh].cell(r, c).value
        if xv is None or abs(xv - pv) > 1e-6 * max(1, abs(pv)):
            bad.append((f"{sh}!R{r}C{c}", xv, pv))
    for t in wb["敏感度"].iter_rows(values_only=True):
        if "排序需重建" in t:
            bad.append(("敏感度排序", t, None))
    return errs, bad


def main():
    os.makedirs(DIST, exist_ok=True)
    tmp = tempfile.mkdtemp(prefix="wh_build_")
    raw = os.path.join(tmp, OUT_NAME)
    checks, sens_top = build(raw)
    rc = os.path.join(tmp, "rc")
    os.makedirs(rc)
    recalc(raw, rc)
    out = os.path.join(DIST, OUT_NAME)
    shutil.copy(os.path.join(rc, OUT_NAME), out)
    errs, bad = verify(out, checks)
    print("輸出：", out)
    print("錯誤值：", len(errs), errs[:10])
    print("Python 對照不一致：", len(bad), bad[:10])
    b = {k: val(k) for k in ("W", "amr_pb", "amr_pb1", "asrs_pb", "c_man", "c_amr", "c_asrs", "c_arm", "hum_hr", "A25", "cagr_amr", "gk_share", "as_share", "asp_impl", "amr_raas_be")}
    for k, v in b.items():
        print(f"  {k} = {v:.4f}")
    for k, rows in sens_top.items():
        print("  敏感度前三", k, [(x[1], round(x[2], 3), round(x[3], 3)) for x in rows])
    shutil.rmtree(tmp, ignore_errors=True)
    if errs or bad:
        sys.exit(1)


if __name__ == "__main__":
    main()
