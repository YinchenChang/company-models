# -*- coding: utf-8 -*-
"""Fanuc（ファナック，6954.T）個股模型 v0.1.1 建置腳本。

用法：python3 build_fanuc.py            # 建置成品＋敏感度快照＋反解＋核對
      python3 build_fanuc.py --quick    # 只建置成品（不重跑敏感度，Sens 頁以上次 sens.json 填值）

結構（docs/plan/個股模型原則_命題加營收分部.md）：命題模組（訂單→營收傳導）→ 營收分部（官方單一分部之下
四個公司固定揭露的部門：FA、Robot、Robomachine、Service）→ 損益表 → 標準 FCFF DCF（docs/plan/個股DCF標準規格.md）
＋同業本益比 Comparable，各 50%，三情境加權。實體 AI 線，不讀 Siliconomics／Tokenomics／Wattonomics 協調檔。
"""
import hashlib
import json
import os
import shutil
import subprocess
import sys
import tempfile

import openpyxl
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.workbook.defined_name import DefinedName

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(os.path.dirname(HERE))
VER = "v0.1.1"
DATE = "20261123"
NAME = "發那科"
OUT = os.path.join(HERE, "dist", f"{DATE}_{NAME}個股模型_{VER}.xlsx")

BLUE = Font(color="0000FF")
BOLD = Font(bold=True)
HFILL = PatternFill("solid", fgColor="DDEBF7")
SFILL = PatternFill("solid", fgColor="F2F2F2")
WRAP = Alignment(wrap_text=True, vertical="top")

# ---------------------------------------------------------------- 來源
S_REF09 = "Fanuc 2027年3月期 上半期決算補足資料（Supplementary data, reference202609_e.pdf，2026-10-26）https://www.fanuc.co.jp/en/ir/announce/pdf/2026/reference202609_e.pdf"
S_REF03 = "Fanuc 2026年3月期 決算補足資料（reference202603_e.pdf，2026-04）https://www.fanuc.co.jp/en/ir/announce/pdf/2026/reference202603_e.pdf"
S_TS09 = "Fanuc 2027年3月期 第2四半期（中間期）決算短信〔日本基準〕（連結）（2026-10-26）https://www.fanuc.co.jp/ja/ir/announce/pdf/2026/financialresult202609.pdf"
S_TS03 = "Fanuc 2026年3月期 決算短信〔日本基準〕（連結）（2026-04）https://www.fanuc.co.jp/ja/ir/announce/pdf/2026/financialresult202603.pdf"
S_PX = "traders.co.jp 個股頁（2026-11-16 15:30 收盤，15 分延遲；PER 為該站以各公司業績予想計）"
S_YF = "Yahoo Finance 日線 2026-11-20 收盤（physical_ai/data/closes_20261109_20261120.csv）"
S_JGB = "日本財務省 國債金利情報（jgbcme.csv）10 年：2026-11-02 2.817%（https://www.mof.go.jp/english/policy/jgbs/reference/interest_rate/jgbcme.csv；沿用 accretech v0.3.6 同一取值，跨模型一致）"
S_CRP = "Damodaran 國家風險溢酬（2026-01-05）：日本 A1、CRP 0.91%（https://pages.stern.nyu.edu/~adamodar/New_Home_Page/datafile/ctryprem.html；沿用 accretech v0.3.6）"
S_NEWS = "日刊工業新聞／ニュースイッチ（2026-10-27）https://newswitch.jp/p/50631"
S_L2A = "線 2a 工業機器人框架 v0.1（/mnt/user-data/outputs/lines/線2a_工業機器人_框架_v0.1.md，2026-11-16）"

# ---------------------------------------------------------------- 歷史（Hist 頁）
# (名稱, 項目, 單位, FY2024, FY2025, FY2026 上半, 標記, 來源)
HIST = [
    ("sec", "A. 部門營收（¥B；公司固定揭露的部門別，官方報告分部＝單一分部）"),
    ("h_rfa", "FA 部門營收", "¥B", 194.8, 208.5, "=Q!H6+Q!I6", "Verified", S_REF03 + "；上半＝Q 頁 FY26Q1＋Q2"),
    ("h_rrb", "Robot 部門營收", "¥B", 329.6, 378.6, "=Q!H7+Q!I7", "Verified", S_REF03 + "；上半＝Q 頁"),
    ("h_rrm", "Robomachine 部門營收", "¥B", 137.6, 129.6, "=Q!H8+Q!I8", "Verified", S_REF03 + "；上半＝Q 頁"),
    ("h_rsv", "Service 部門營收", "¥B", 135.2, 141.1, "=Q!H9+Q!I9", "Verified", S_REF03 + "；上半＝Q 頁"),
    ("h_rev", "合併營收", "¥B", 797.129, 857.831, 473.708, "Verified", S_TS03 + "；" + S_TS09),
    ("sec", "B. 部門訂單（¥B；FY2025＝四季加總，FY2024 全年未取得）"),
    ("h_ofa", "FA 部門訂單", "¥B", None, "=SUM(Q!D13:G13)", "=Q!H13+Q!I13", "Derived", S_REF03 + "、" + S_REF09 + "（季資料加總）"),
    ("h_orb", "Robot 部門訂單", "¥B", None, "=SUM(Q!D14:G14)", "=Q!H14+Q!I14", "Derived", "同上"),
    ("h_orm", "Robomachine 部門訂單", "¥B", None, "=SUM(Q!D15:G15)", "=Q!H15+Q!I15", "Derived", "同上"),
    ("h_osv", "Service 部門訂單", "¥B", None, "=SUM(Q!D16:G16)", "=Q!H16+Q!I16", "Derived", "同上"),
    ("sec", "C. 損益（¥B；日本會計準則）"),
    ("h_cogs", "營業成本", "¥B", 502.204, 529.365, 280.0, "Verified", S_TS03 + "；上半 280.0 取自 " + S_REF09 + "（取整到億日圓）"),
    ("h_sga", "銷管費", "¥B", 136.079, 144.703, "=h_rev_h1-h_cogs_h1-h_oi_h1", "Verified", S_TS03 + "；上半欄＝營收 − 營業成本 − 營業利益（推得）"),
    ("h_oi", "營業利益", "¥B", 158.846, 183.763, 116.662, "Verified", S_TS03 + "；" + S_TS09),
    ("h_eq", "持分法投資利益", "¥B", 28.040, 30.531, None, "Verified", S_TS03 + "（上半未逐行取得）"),
    ("h_ord", "經常利益", "¥B", 196.738, 227.485, 149.045, "Verified", S_TS03 + "；" + S_TS09),
    ("h_tax", "法人稅等合計", "¥B", 45.391, 58.088, None, "Verified", S_TS03),
    ("h_nci", "非支配股東利益", "¥B", 3.790, 2.854, None, "Verified", S_TS03),
    ("h_ni", "歸屬母公司淨利", "¥B", 147.557, 166.543, 112.576, "Verified", S_TS03 + "；" + S_TS09),
    ("h_eps", "基本 EPS", "¥", 157.31, 178.47, 120.90, "Verified", S_TS03 + "；" + S_TS09),
    ("h_dps", "每股股利", "¥", 94.39, 107.09, 72.54, "Verified", S_TS03 + "（配當性向 60%）；上半＝中間配當實績 72.54"),
    ("sec", "D. 現金流與資產負債（¥B）"),
    ("h_da", "減價償却費", "¥B", 46.448, 47.782, 23.147, "Verified", S_TS03 + "；" + S_TS09),
    ("h_cx", "有形固定資產取得支出", "¥B", 40.805, 21.181, 16.603, "Verified", S_TS03 + "；" + S_TS09 + "（無形固定資產取得未單列）"),
    ("h_dwc", "營運資金增加（＋＝占用現金）", "¥B", "=4.613-52.489+4.480", "=0.658-27.869-5.468", "=-11.172+11.059-10.067", "Derived",
     "CF 計算書：−(売上債権の増減額) −(棚卸資産の増減額) −(仕入債務の増減額)；FY2024 △4,613／52,489／△4,480；FY2025 △658／27,869／5,468；上半 11,172／△11,059／10,067（百萬日圓）"),
    ("h_wc", "營運資金水準（應收＋存貨 − 應付）", "¥B", "=156.437+306.439-37.454", "=170.775+292.186-46.262", None, "Derived", S_TS03 + "：受取手形及び売掛金、棚卸資産、支払手形及び買掛金"),
    ("h_taxpaid", "法人稅等支付額", "¥B", 32.363, 46.095, 35.039, "Verified", S_TS03 + "；" + S_TS09),
    ("h_cash", "現金及び預金（期末）", "¥B", 590.504, 718.071, 746.285, "Verified", S_TS03 + "；" + S_TS09 + "（上半欄＝2026-09-30）"),
    ("h_sec", "有価証券（流動）", "¥B", 16.000, 35.800, 40.800, "Verified", "同上"),
    ("h_invsec", "投資有価証券（含持分法投資）", "¥B", 192.214, 223.287, 242.860, "Verified", "同上；不另加入股權價值（持分法價值以 eq_pe 另估，避免重複）"),
    ("h_ncib", "非支配株主持分（帳面）", "¥B", None, 18.076, 18.676, "Verified", "同上"),
    ("h_shout", "期末流通股數（發行 − 自己株式）", "十億股", None, "=(982.383493-49.224684)/1000", "=(982.272430-57.358897)/1000", "Derived", S_TS09 + "：期末発行済株式数 982,272,430、自己株式 57,358,897（2026-09-30）"),
]

# ---------------------------------------------------------------- 季資料（Q 頁）：FY24Q4, FY25Q1..Q4, FY26Q1, Q2
QCOLS = ["FY24 Q4", "FY25 Q1", "FY25 Q2", "FY25 Q3", "FY25 Q4", "FY26 Q1", "FY26 Q2"]
QDATA = [
    ("sec", "部門營收（¥B）"),
    ("FA 營收", [47.4, 49.7, 54.2, 49.8, 54.8, 57.4, 62.5]),
    ("Robot 營收", [87.2, 81.0, 91.6, 96.6, 109.4, 96.1, 98.2]),
    ("Robomachine 營收", [44.8, 33.9, 29.6, 33.0, 33.1, 41.7, 41.6]),
    ("Service 營收", [32.8, 31.8, 35.8, 36.3, 37.2, 35.9, 40.3]),
    ("合併營收（公司揭露）", [212.1, 196.4, 211.2, 215.7, 234.5, 231.0, 242.7]),
    ("檢查：部門合計 − 合併營收（季）", None),
    ("sec", "部門訂單（¥B）"),
    ("FA 訂單", [50.7, 55.0, 48.8, 54.7, 71.1, 90.2, 112.2]),
    ("Robot 訂單", [90.9, 90.6, 91.8, 98.8, 100.6, 97.9, 103.4]),
    ("Robomachine 訂單", [37.5, 29.0, 29.7, 30.8, 43.8, 57.6, 51.2]),
    ("Service 訂單", [32.3, 31.2, 35.1, 35.9, 36.5, 36.3, 39.3]),
    ("訂單合計（公司揭露）", [211.4, 205.9, 205.3, 220.2, 252.0, 281.9, 306.0]),
    ("sec", "訂單 ÷ 營收（Derived）"),
    ("FA 訂單÷營收", "b2b:13:6"),
    ("Robot 訂單÷營收", "b2b:14:7"),
    ("Robomachine 訂單÷營收", "b2b:15:8"),
    ("合計 訂單÷營收", "b2b:17:10"),
    ("sec", "地區別訂單（¥B）"),
    ("日本 訂單", [None, None, 27.9, 29.2, 34.3, 39.3, 43.0]),
    ("美洲 訂單", [None, None, 56.1, 58.9, 61.2, 62.8, 64.5]),
    ("歐洲 訂單", [None, None, 37.2, 40.1, 37.5, 38.9, 39.2]),
    ("中國 訂單", [None, None, 50.7, 57.7, 79.0, 88.7, 91.2]),
    ("中國以外亞洲 訂單", [None, None, 30.8, 32.4, 36.8, 49.1, 65.3]),
    ("sec", "營業利益率（%；公司揭露）"),
    ("營業利益率", [0.228, 0.216, 0.206, 0.193, 0.239, 0.232, 0.260]),
]

# ---------------------------------------------------------------- In 頁輸入
# (鍵, 項目, 單位, 低, 基準, 高, 標記, 來源, 說明, 來源誘因)
SCN_DIV = {  # 訂單年增 FY2027／FY2028／FY2029
    "fa": {"A": (0.05, -0.05, 0.05), "B": (-0.10, -0.10, 0.04), "C": (-0.35, -0.10, 0.03)},
    "rb": {"A": (0.15, 0.12, 0.10), "B": (0.06, 0.05, 0.05), "C": (-0.08, 0.00, 0.04)},
    "rm": {"A": (0.05, -0.05, 0.05), "B": (-0.10, -0.08, 0.04), "C": (-0.30, -0.10, 0.03)},
}
DIVN = {"fa": "FA", "rb": "Robot", "rm": "Robomachine"}
SCN_NAME = {"A": "樂觀", "B": "基準", "C": "保守"}
GSV = {"A": 0.08, "B": 0.06, "C": 0.03}
GLT = {"A": 0.05, "B": 0.035, "C": 0.02}
SCW = {"A": 0.25, "B": 0.50, "C": 0.25}

PEERS = [  # (名稱, 代號, 2026-11-20 收盤, 2026-11-16 收盤（traders，分割前口徑）, 2026-11-16 PER（traders）, 分割比率：舊收盤口徑股數 → 新收盤口徑)
    ("SMC", "6273.T", 55300, 60270, 15.81, 1),
    ("Keyence", "6861.T", 14025, 73450, 31.25, 5),   # 2026-11-19 1 拆 5；csv 11-16 調整後 14,690 × 5＝73,450，與舊收盤一致
    ("Nabtesco", "6268.T", 3105, 3574, 16.89, 1),
    ("Okuma", "6103.T", 3630, 3935, 11.45, 1),
    ("DMG Mori", "6141.T", 2361.5, 2572.5, 25.91, 1),
]
PX_OLD = 5343      # Fanuc 2026-11-16 收盤（traders.co.jp；v0.1 現價）
PE_OLD = 21.65     # Fanuc 2026-11-16 PER（traders.co.jp）

INROWS = []
def sec(t): INROWS.append(("sec", t))
def inp(*a): INROWS.append(a)

sec("A. 價格、股數與日期")
inp("px", "現價（2026-11-20 收盤）", "¥", None, 5047, None, "Verified", S_YF + "：5,047", "as_of＝現價日；v0.1 對照：2026-11-16 收盤 5,343（traders.co.jp）", "")
inp("px_date", "現價日", "日期", None, "=DATE(2026,11,20)", None, "Verified", S_YF, "資料基準日（as_of）", "")
inp("tp_date", "目標價時點（px_date＋12 個月）", "日期", None, "=EDATE(px_date,12)", None, "Derived", "＝EDATE(px_date,12)（原則第 10 節第 26 條）", "", "")
inp("val_date", "評價日（資產負債表日）", "日期", None, "=DATE(2026,9,30)", None, "Verified", S_TS09 + "：中間期末 2026-09-30", "FY2026 只計下半（2026-10～2027-03）現金流", "")
inp("sh_v", "評價日股數（期末流通股數）", "十億股", None, "=h_shout_h1", None, "Derived", "＝Hist 頁期末流通股數 924.91 百萬股（2026-09-30）", "RSU／股票選擇權：公司未揭露重大稀釋（未取得），不另加；2026-10 新宣布上限 500 億日圓自社株買い未扣（按公允價買回對每股價值中性）", "")
inp("payout", "配當性向（公司方針）", "%", None, 0.60, None, "Interested-party", S_TS03 + "：配當性向 60%（FY2024、FY2025 實績皆 60.0%）", "用於期間股利與收支關卡", "發行公司股東回饋方針，可變更")
inp("div_int", "FY2026 中間配當（已決議，2026-09-30 基準、評價日後支付）", "¥/股", None, 72.54, None, "Verified", S_TS09 + "：2027年3月期 第2四半期末 72.54 円（期末予想未定）", "在評價日資產負債表現金內，自股權價值扣除", "")

sec("B. FY2026 下半（2026-10～2027-03）：部門營收與訂單（共用，不分情境）")
inp("h2_fa", "FA 下半營收 ÷ 上半 − 1", "%", 0.10, 0.25, 0.40, "Assumed", "上半 FA 訂單 202.4 對營收 119.9（訂單÷營收 1.69，Q 頁）；7–9 月 FA 訂單 1,122 億、2.3 倍，AI、資料中心、半導體、人形零件加工需求（" + S_NEWS + "）", "記憶體等部分零件吃緊（同上）→ 低案；公司下半全體營收指引 540.0（Recon）", "")
inp("h2_rb", "Robot 下半營收 ÷ 上半 − 1", "%", 0.0, 0.08, 0.15, "Assumed", "上半 Robot 訂單÷營收 1.04（Q 頁）；FY2025 下半 ÷ 上半＝(96.6＋109.4)÷(81.0＋91.6)−1＝19%（季節性）", "", "")
inp("h2_rm", "Robomachine 下半營收 ÷ 上半 − 1", "%", 0.0, 0.15, 0.25, "Assumed", "上半 Robomachine 訂單÷營收 1.31（Q 頁）", "", "")
inp("h2_sv", "Service 下半營收 ÷ 上半 − 1", "%", 0.0, 0.05, 0.10, "Assumed", "Service 訂單≈營收（Q 頁）；FY2025 下半 ÷ 上半＝8%", "", "")
inp("o2h_fa", "FA 下半訂單 ÷（7–9 月訂單 × 2）− 1", "%", -0.30, -0.10, 0.05, "Assumed", "7–9 月 FA 訂單 112.2 為歷史新高（" + S_REF09 + "）；基準假設下半回落 10%", "", "")
inp("o2h_rb", "Robot 下半訂單 ÷（7–9 月訂單 × 2）− 1", "%", -0.10, 0.0, 0.08, "Assumed", "7–9 月 Robot 訂單 103.4，YoY +12.6%、QoQ +5.6%", "", "")
inp("o2h_rm", "Robomachine 下半訂單 ÷（7–9 月訂單 × 2）− 1", "%", -0.30, -0.10, 0.05, "Assumed", "7–9 月 Robomachine 訂單 51.2，QoQ −11.1%", "", "")
inp("opm_h2", "FY2026 下半營業利益率", "%", 0.22, "=h_oi_h1/h_rev_h1", 0.26, "Derived", "＝上半營業利益 116.662 ÷ 營收 473.708＝24.6%（" + S_TS09 + "）", "低／高 Assumed：公司下半指引隱含 23.0%（下半匯率假設 150 円／美元，上半實績 159.37）", "")
inp("nonop_h2", "FY2026 下半營業外收支 ÷ 上半 − 1", "%", -0.30, 0.0, 0.10, "Assumed", "上半營業外收支（經常利益 − 營業利益）＝149.045 − 116.662＝32.383；公司下半指引隱含 35.7（Recon）", "上半含日圓走貶的匯兌益（金額未逐行取得）", "")

sec("C. 命題模組：訂單 → 營收傳導")
inp("lag_fa", "FA：當年營收中來自上一年訂單的比重（L）", "比率", 0.30, 0.50, 0.70, "Assumed", "營收＝(1−L)×當年訂單＋L×上年訂單；CNC 交期與上半訂單÷營收 1.69 顯示積壓", "公司不揭露受注殘（未取得）", "")
inp("lag_rb", "Robot：L", "比率", 0.10, 0.30, 0.50, "Assumed", "Robot 訂單÷營收接近 1（Q 頁），傳導快", "", "")
inp("lag_rm", "Robomachine：L", "比率", 0.20, 0.40, 0.60, "Assumed", "Robomachine 訂單÷營收 1.31（上半）", "", "")
inp("o_shift", "FY2027–FY2029 訂單年增平移（FA、Robot、Robomachine；三情境同加）", "pt", -0.05, 0.0, 0.05, "Assumed", "反解用（Recon）；基準 0", "", "")
for d in ("fa", "rb", "rm"):
    for s in "ABC":
        for i, y in enumerate((27, 28, 29)):
            inp(f"o_{d}{y}_{s}", f"情境 {s}（{SCN_NAME[s]}）：{DIVN[d]} 訂單年增 FY20{y}", "%", None, SCN_DIV[d][s][i], None, "Assumed",
                {"fa": "FA：工具機（含 AI、資料中心、半導體、人形零件加工）循環；FY2026 訂單約為 FY2025 的 1.8 倍，基準假設 FY2027–28 回落",
                 "rb": "Robot：線 2a IFR 全球裝機 2025 年 60.3 萬台、預測 2029 年 80.6 萬台（年增約 7.5%）；JARA 7–9 月受注 +46.6%",
                 "rm": "Robomachine（ROBODRILL 等，連動 IT／智慧手機與汽車零件加工）：循環性高"}[d],
                "三情境即區間（A／C 為上下緣）", "")
for s in "ABC":
    inp(f"gsv_{s}", f"情境 {s}（{SCN_NAME[s]}）：Service 營收年增 FY2027–FY2029", "%", None, GSV[s], None, "Assumed", "裝機基數成長；FY2025 +4.4%、上半 FY2026 +13.5%（Hist 頁）", "", "")
for s in "ABC":
    inp(f"glt_{s}", f"情境 {s}（{SCN_NAME[s]}）：FY2030–FY2032 各部門訂單與營收年增", "%", None, GLT[s], None, "Assumed", "收斂到長期成長", "", "")

sec("D. 損益參數")
inp("inc_m", "營業利益的增量利潤率（Δ營業利益 ÷ Δ營收，FY2027 起相對 FY2026）", "%", 0.30, "=((h_oi_25-h_oi_24)/(h_rev_25-h_rev_24)+(h_oi_h1-85.964)/(h_rev_h1-407.568))/2", 0.50, "Derived",
    "＝FY2025 對 FY2024：24.9÷60.7＝41.0%；上半 FY2026 對上半 FY2025：30.7÷66.1＝46.4%（上半 FY2025 營收 407.568、營業利益 85.964，" + S_TS09 + "）；兩者平均", "同時作用於營收下滑（減量利潤率同值）；公司不揭露部門利益", "")
inp("eq_link", "持分法投資利益隨 Robot 營收等比例", "開關", None, 1, None, "Assumed", "持分法主要為中國機器人合資（上海發那科，本輪未以原文核對）；以 Robot 營收比例外推", "", "")
inp("nonop_o", "FY2027 起營業外收支（不含持分法）", "¥B", 8.0, "=h_ord_25-h_oi_25-h_eq_25", 18.0, "Derived", "＝FY2025 經常利益 − 營業利益 − 持分法＝13.2（受取利息 7.5、受取配當 3.4 等）", "低／高 Assumed", "")
inp("tax_w", "稅率（法人稅 ÷（經常利益 − 持分法））", "%", 0.25, "=(h_tax_24/(h_ord_24-h_eq_24)+h_tax_25/(h_ord_25-h_eq_25))/2", 0.31, "Derived",
    "FY2024 45.391÷168.698＝26.9%；FY2025 58.088÷196.954＝29.5%；平均", "NOPAT 也用本稅率；現金稅率（支付額÷稅前）FY2025 約 23.4%（Hist）", "")
inp("nci_r", "非支配股東利益 ÷ 稅後利益", "%", None, "=h_nci_25/(h_ni_25+h_nci_25)", None, "Derived", "FY2025 2.854 ÷ 169.397", "", "")
inp("da_life", "折舊調整年數（FY2027 起：折舊＋(上年資本支出 − 上年折舊) ÷ 本值）", "年", 6, 10, 14, "Assumed", "FY2026 折舊＝上半 23.147 × 2（Derived）；FY2021–FY2025 折舊 46–50 十億日圓、有形資本支出 21–53（補足資料、決算短信），資本支出低於折舊時折舊緩降", "取代「折舊 ÷ 營收」比例法：Fanuc 折舊金額穩定，不隨營收等比例", "")
inp("cx_r", "資本支出 ÷ 營收", "%", 0.025, "=(h_cx_24/h_rev_24+h_cx_25/h_rev_25+h_cx_h1/h_rev_h1)/3", 0.055, "Derived", "FY2024 5.1%、FY2025 2.5%、上半 FY2026 3.5% 平均", "", "")
inp("wc_r", "營運資金增加 ÷ 營收增量", "%", 0.30, "=h_wc_25/h_rev_25", 0.55, "Derived", "＝FY2025 末營運資金水準 416.7 ÷ 營收 857.8＝48.6%（增量比率 FY2024、FY2025 皆為負：存貨去化，不適用）", "", "")

sec("E. 折現率與估值")
inp("rf", "無風險利率（日本 10 年期國債）", "%", None, 0.02817, None, "Verified", S_JGB, "與現價日相差逾 2 個交易日：逾期（跨模型一致，下次全面改版更新）", "")
inp("crp", "日本國家風險溢酬", "%", None, 0.0091, None, "Verified", S_CRP, "", "")
inp("erp", "股權風險溢酬", "%", 0.04, 0.05, 0.06, "Assumed", "DCF 標準規格預設 5%（敏感度 4–6%）", "", "")
inp("beta", "β（槓桿；公司無有息負債＝無槓桿）", "x", 0.85, 1.10, 1.35, "Assumed", "未取得：本輪無法讀取股價序列自算 β（規格要求兩來源 Blume 調整中位數）；以日本 FA／機械大型股常見區間假設", "承重：β 每 ±0.25 影響 ke 約 ±1.5pt", "")
inp("wacc_adj", "WACC 調整", "pt", -0.01, 0.0, 0.01, "Assumed", "Sens 用 ±1pt；基準 0", "", "")
inp("g_tv", "永續成長率 g", "%", 0.02, "=MIN(0.03,rf)", 0.03, "Derived", "＝MIN(3%, rf)，同 accretech 慣例", "", "")
inp("tv_alt", "永續價值改用 FCFF×(1＋g) 法（0＝NOPAT 法）", "開關", None, 0, 1, "Assumed", "DCF 標準規格第 4 節替代法", "", "")
inp("eq_pe", "持分法投資價值倍數（× FY2027 持分法利益）", "x", 8, 12, 16, "Assumed", "持分法利益為稅後；倍數低於同業中位數以反映少數股權與中國風險", "投資有価証券 242.9 不另加入（假設主要為持分法投資，避免重複；其他持股未取得）", "")
inp("peer_pe", "同業前瞻本益比中位數", "x", None, "=MEDIAN(Val!D40:D44)", None, "Derived", "Val 頁同業表（SMC、Keyence、Nabtesco、Okuma、DMG Mori；收盤 " + S_YF + "；PER 由 " + S_PX + " 按收盤比例換算）；安川 6506、哈默納科 6324 頁面讀取失敗，未取得", "少於 8 家，不做迴歸（原則第 10 節第 7 條）", "")
inp("rel_m", "相對倍數係數", "x", None, 1.0, "=Val!D46/peer_pe", "Assumed", "1.00（未調整：缺 Fanuc 對同業 10 年相對倍數歷史）", "高案＝現行 Fanuc PER ÷ 同業中位數（維持現行溢價）", "")
inp("w_dcf", "DCF 權重（Comparable＝1−本值）", "%", 0.30, 0.50, 0.70, "Assumed", "個股模型原則第 4 節預設各半", "", "")
for s in "ABC":
    inp(f"scw_{s}", f"情境權重 {s}（{SCN_NAME[s]}）", "%", None, SCW[s], None, "Assumed", "DCF 標準規格第 3 節預設 25／50／25", "", "")
IN_LAST_KEY = "scw_C"

# ---------------------------------------------------------------- 對帳頁（線層級變數）
RECON_LINES = [
    ("全球工業機器人裝機 2025", "2a", "60.3 萬台（+11%）", "台", "IFR 年度裝機", "IFR President's Report 2026（經 " + S_L2A + "）", "Verified", "Robot 訂單年增情境（o_rb27_*）的背景；本模型不直接用台數"),
    ("全球工業機器人裝機 IFR 預測 2026／2029", "2a", "65.5 萬／80.6 萬台（2025–2029 年增約 7.5%）", "台", "IFR 預測", "同上", "Interested-party", "情境 B Robot 訂單年增 5–6%（金額口徑；含價格下滑）低於台數成長；情境 A 10–15%"),
    ("中國裝機與國產比 2025", "2a", "35.4 萬台；國產 55%", "台／%", "IFR", "同上", "Verified", "未直接使用；Fanuc 中國營收占比（上半 28.6%）與持分法（中國合資）風險"),
    ("JARA 受注額 2026 年 7–9 月", "2a", "3,254 億日圓，+46.6%", "¥", "日本機器人工業會", "JARA（經線 2a 框架）", "Verified", "Fanuc Robot 訂單同期 +12.6%，遠低於 JARA 全體：本模型未假設 Fanuc 追上產業增速"),
    ("中國工業機器人產量 2026 年 1–9 月", "2a", "79.5 萬台，+29.1%", "台", "國家統計局（含出口、口徑≠IFR）", "同上", "Verified", "Robot 情境 A 的背景"),
    ("整機廠營業利益率（安川機器人 FY2026 預估）", "2a", "約 7.1%", "%", "安川財測推得", "technews（經線 2a 框架）", "Derived", "Fanuc 不揭露部門利益；合併營業利益率 24.6%（上半）遠高於此值：Fanuc 利潤主要來自 FA（CNC）與服務，不是 Robot（未取得部門利益，推論）"),
    ("傳統 6 軸 BOM：減速機占比", "2a", "25–35%", "%", "業界常引", "線 2a 框架", "Analogy", "未使用（Fanuc 減速機多自製／外購比例未取得）"),
    ("人形機器人零件加工需求", "2b", "定性：Fanuc 稱 FA 需求「一段引き上がっている」（含人形零件加工）", "—", "公司說法", S_NEWS, "Interested-party", "FA 下半營收與訂單假設（h2_fa、o2h_fa）的背景；人形 BOM 單價本模型未用"),
]

RISKS = [
    ("FA 訂單為一次性（AI／資料中心設備投資、中國拉貨）而非新循環", "是（情境 C：FY2027 FA 訂單 −35%）", "最大承重變數", "Fanuc 每季 FA 訂單、中國地區訂單、日本工作機械工業會受注"),
    ("零件（記憶體等）短缺限制出貨", "部分（h2_fa 低案）", "公司說法：部品調達為最優先課題", "下半營收對指引、存貨變動"),
    ("中國價格戰與國產替代（Robot、持分法合資）", "部分（eq_pe、情境 C Robot）", "IFR 國產比 55%", "中國 Robot 訂單、持分法投資利益"),
    ("日圓升值（公司下半假設 150 円／美元，上半 159.37）", "否", "匯率不建模，營業利益率以上半實績外推", "季平均匯率與營業利益率"),
    ("美國關稅與貿易限制", "否", "不可量化", "美洲訂單"),
    ("physical AI 帶動新機器人應用（NVIDIA 合作、VTLA）", "部分（情境 A Robot）", "目前無營收揭露", "Robot 訂單中一般工業占比（未揭露）"),
]

SOURCES = [S_REF09, S_REF03, S_TS09, S_TS03, S_PX + "：https://www.traders.co.jp/stocks/6954/（另 6273、6861、6268、6103、6141）", S_YF + "（6954、6273、6861、6268、6103、6141；6861 於 2026-11-19 1 拆 5，見 csv.splits.txt）", S_JGB, S_CRP, S_NEWS,
           "ビジネス+IT（2026-10-31）Fanuc 上半與上限 500 億日圓自社株買い：https://www.sbbit.jp/article/st/187873",
           "Business Standard（2025-12）Fanuc 與 NVIDIA physical AI 合作：https://www.business-standard.com/world-news/japan-fanuc-nvidia-physical-ai-industrial-robots-voice-command-tech-125120200779_1.html",
           "Fanuc America 新聞稿（2026-03-16）FANUC Accelerates Physical AI in Industrial Robotics, Leveraging NVIDIA Technologies：https://www.fanucamerica.com/news-resources/fanuc-america-press-releases/2026/03/16/fanuc-accelerates-physical-ai-in-industrial-robotics-leveraging-nvidia-technologies",
           S_L2A]

YCOLS = "CDEFGHIJK"           # FY2024..FY2032
YEARS = list(range(2024, 2033))


def col_of(y):
    return YCOLS[y - 2024]


def build(over=None, sens=None, rev=None, path=OUT):
    over = over or {}
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    names = {}

    def defname(n, ref):
        wb.defined_names[n] = DefinedName(n, attr_text=ref)
        names[n] = ref

    # ---------------- Q
    Q = wb.create_sheet("Q")
    Q["A1"] = "Q — 季資料（十億日圓；FY＝4 月起；FY26＝2026-04～2027-03）"; Q["A1"].font = BOLD
    Q["A2"] = "來源：FY24Q4–FY25Q4 " + S_REF03 + "；FY25Q2–FY26Q2 " + S_REF09
    hdr = ["項目", "單位"] + QCOLS + ["標記", "來源", "來源誘因"]
    for j, h in enumerate(hdr):
        Q.cell(4, j + 1, h).font = BOLD; Q.cell(4, j + 1).fill = HFILL
    r = 5
    for it in QDATA:
        if it[0] == "sec":
            Q.cell(r, 1, it[1]).font = BOLD; r += 1; continue
        lab, vals = it
        Q.cell(r, 1, lab); Q.cell(r, 2, "%" if "率" in lab else ("x" if "÷" in lab else "¥B"))
        if isinstance(vals, list):
            for j, v in enumerate(vals):
                if v is not None:
                    c = Q.cell(r, 3 + j, v); c.font = BLUE
            Q.cell(r, 10, "Verified"); Q.cell(r, 11, S_REF09 if "地區" in lab or "日本" in lab or "美洲" in lab or "歐洲" in lab or "中國" in lab or "亞洲" in lab else (S_REF03 + "；" + S_REF09))
        elif vals is None:   # 檢查列
            for j in range(7):
                cl = "CDEFGHI"[j]
                Q.cell(r, 3 + j, f"=ROUND(SUM({cl}6:{cl}9)-{cl}10,1)")
            Q.cell(r, 10, "Derived"); Q.cell(r, 11, "部門取整到億日圓，容許 ±0.1")
        else:
            _, a, b = vals.split(":")
            for j in range(7):
                cl = "CDEFGHI"[j]
                Q.cell(r, 3 + j, f"={cl}{a}/{cl}{b}").number_format = "0.00"
            Q.cell(r, 10, "Derived"); Q.cell(r, 11, "訂單 ÷ 營收")
        if "率" in lab:
            for j in range(7): Q.cell(r, 3 + j).number_format = "0.0%"
        r += 1
    assert Q["A6"].value == "FA 營收" and Q["A13"].value == "FA 訂單" and Q["A17"].value.startswith("訂單合計"), (Q["A6"].value, Q["A13"].value, Q["A17"].value)
    Q.column_dimensions["A"].width = 30; Q.column_dimensions["K"].width = 60

    # ---------------- Hist
    H = wb.create_sheet("Hist")
    H["A1"] = "Hist — 歷史財務（十億日圓；FY2024＝2024-04～2025-03，FY2025＝2025-04～2026-03，FY2026 上半＝2026-04～09）"; H["A1"].font = BOLD
    for j, h in enumerate(["名稱", "項目", "FY2024", "FY2025", "FY2026 上半", "單位", "標記", "來源"]):
        H.cell(3, j + 1, h).font = BOLD; H.cell(3, j + 1).fill = HFILL
    # 版面：C＝FY2024、D＝FY2025、E＝FY2026 上半；公式中以 Hist!C9 等引用
    r = 4
    hrow = {}
    for it in HIST:
        if it[0] == "sec":
            H.cell(r, 1, it[1]).font = BOLD; r += 1; continue
        k, lab, unit, v24, v25, vh1, tag, src = it
        H.cell(r, 1, k); H.cell(r, 2, lab)
        for j, v in enumerate((v24, v25, vh1)):
            if v is not None:
                c = H.cell(r, 3 + j, v)
                if not (isinstance(v, str) and v.startswith("=") and "Q!" in v or (isinstance(v, str) and "Hist" in v)):
                    c.font = BLUE
                c.number_format = "#,##0.000" if unit != "¥" else "#,##0.00"
        H.cell(r, 6, unit); H.cell(r, 7, tag); H.cell(r, 8, src)
        hrow[k] = r
        for j, yy in enumerate(("24", "25", "h1")):
            defname(f"{k}_{yy}", f"Hist!${'CDE'[j]}${r}")
        r += 1
    H.column_dimensions["B"].width = 34; H.column_dimensions["H"].width = 80

    # ---------------- In
    I = wb.create_sheet("In", 0)
    I["A1"] = ("In — 輸入（藍字；基準欄 E 進主模型；D／F＝敏感度低／高，只用於 Sens 頁（一次動一個參數，整本重算）、不進目標價；"
               "情境 A／B／C 的驅動見本頁最後一節「情境驅動對照」）。金額：十億日圓（¥B）；每股：日圓。")
    I["A2"] = "標記：Verified（公司文件、交易所資料）／Interested-party（寫明是誰、立場）／Analogy／Assumed（給低／高區間或以情境 A／C 為區間）／Derived（寫明換算）。查不到者寫「未取得」。"
    heads = ["鍵", "項目", "單位", "敏感度低", "基準", "敏感度高", "標記", "來源", "說明", "來源誘因"]
    for j, h in enumerate(heads):
        I.cell(3, j + 1, h).font = BOLD; I.cell(3, j + 1).fill = HFILL
    r = 4
    inrow = {}
    for it in INROWS:
        if it[0] == "sec":
            I.cell(r, 1, it[1]).font = BOLD; I.cell(r, 1).fill = SFILL; r += 1; continue
        k, lab, unit, lo, base, hi, tag, src, note, inc = it
        if k in over:
            base = over[k]
        I.cell(r, 1, k); I.cell(r, 2, lab); I.cell(r, 3, unit)
        for j, v in ((4, lo), (5, base), (6, hi)):
            if v is not None:
                c = I.cell(r, j, v); c.font = BLUE
                if unit in ("%", "pt"): c.number_format = "0.0%"
                elif unit == "日期": c.number_format = "yyyy-mm-dd"
                elif unit == "比率": c.number_format = "0.00"
        I.cell(r, 7, tag); I.cell(r, 8, src); I.cell(r, 9, note); I.cell(r, 10, inc or None)
        inrow[k] = r
        defname(k, f"In!$E${r}")
        r += 1
    in_last = r - 1
    # 情境驅動對照（純展示）
    r += 1
    I.cell(r, 1, "情境驅動對照（純展示：三情境並排；修改請改上方藍字；本節不被任何公式引用）").font = BOLD
    r += 1
    for j, h in enumerate(["鍵", "項目", "單位", "情境 A（樂觀）", "情境 B（基準）", "情境 C（保守）", "標記", "說明"]):
        I.cell(r, j + 1, h).font = BOLD; I.cell(r, j + 1).fill = HFILL
    r += 1
    pairs = [(f"scn_o_{d}{y}", f"{DIVN[d]} 訂單年增 FY20{y}", "%", [f"o_{d}{y}_{s}" for s in "ABC"]) for d in ("fa", "rb", "rm") for y in (27, 28, 29)]
    pairs += [("scn_gsv", "Service 營收年增 FY2027–29", "%", [f"gsv_{s}" for s in "ABC"]),
              ("scn_glt", "FY2030–32 年增", "%", [f"glt_{s}" for s in "ABC"]),
              ("scn_w", "情境權重", "%", [f"scw_{s}" for s in "ABC"])]
    for k, lab, unit, ks in pairs:
        I.cell(r, 1, k); I.cell(r, 2, lab); I.cell(r, 3, unit)
        for j, kk in enumerate(ks):
            c = I.cell(r, 4 + j, f"=E{inrow[kk]}"); c.number_format = "0.0%"
        I.cell(r, 7, "Derived"); I.cell(r, 8, "引用 In 頁 " + "／".join(ks) + "；修改請改上方藍字")
        r += 1
    I.column_dimensions["A"].width = 14; I.column_dimensions["B"].width = 46
    for cl in "HIJ": I.column_dimensions[cl].width = 60

    # ---------------- PL（三情境）
    P = wb.create_sheet("PL")
    P["A1"] = "PL — 命題模組（訂單→營收）＋部門營收＋損益表＋現金流（十億日圓；FY＝4 月起，FY2026＝2026-04～2027-03）"; P["A1"].font = BOLD
    P["A2"] = ("官方報告分部＝單一分部（決算短信：FA 綜合供應商單一業種，セグメント情報省略）；FA、Robot、Robomachine、Service 為公司在補足資料固定揭露的部門別，"
               "列為合併營收之下的子列（子列合計＝合併營收，核對列）。FY2024、FY2025 為實績；FY2026＝上半實績＋下半假設；FY2027 起由訂單推得。")
    blocks = {}
    r0 = 4
    for s in "ABC":
        R = {}
        P.cell(r0, 1, f"情境 {s}（{SCN_NAME[s]}）").font = Font(bold=True, size=12)
        for j, y in enumerate(YEARS):
            c = P.cell(r0, 3 + j, f"FY{y}"); c.font = BOLD; c.fill = HFILL
        P.cell(r0, 2, "單位").font = BOLD
        rr = r0 + 1
        def row(key, label, unit="¥B", fmt="#,##0.0"):
            nonlocal rr
            P.cell(rr, 1, label); P.cell(rr, 2, unit)
            R[key] = rr; rr += 1
            return R[key]
        def put(key, y, f, fmt=None):
            c = P[f"{col_of(y)}{R[key]}"]; c.value = f
            c.number_format = fmt or ("0.0%" if P.cell(R[key], 2).value == "%" else ("0.00" if P.cell(R[key], 2).value == "x" else "#,##0.0"))
        P.cell(rr, 1, "命題模組：訂單 → 營收").font = BOLD; rr += 1
        for d in ("fa", "rb", "rm"):
            row(f"o_{d}", f"{DIVN[d]} 訂單")
        row("o_sum", "訂單合計（FA＋Robot＋Robomachine）")
        for d in ("fa", "rb", "rm"):
            row(f"bb_{d}", f"{DIVN[d]} 訂單÷營收", "x")
        P.cell(rr, 1, "營收分部（官方單一分部之下的部門子列）").font = BOLD; rr += 1
        for d, lab in (("fa", "FA 營收"), ("rb", "Robot 營收"), ("rm", "Robomachine 營收"), ("sv", "Service 營收")):
            row(f"r_{d}", lab)
        row("rev", "合併營收")
        row("chk", "檢查：分部合計 − 合併營收")
        row("g", "合併營收年增", "%")
        P.cell(rr, 1, "損益表").font = BOLD; rr += 1
        row("oi", "營業利益"); row("opm", "營業利益率", "%")
        row("eq", "持分法投資利益"); row("nop", "其他營業外收支（淨額）"); row("ord", "經常利益")
        row("tax", "法人稅等"); row("nci", "非支配股東利益"); row("ni", "歸屬母公司淨利")
        row("eps", "EPS", "¥", "#,##0.00"); row("dps", "每股股利（配當性向）", "¥", "#,##0.00")
        P.cell(rr, 1, "現金流").font = BOLD; rr += 1
        row("da", "減價償却"); row("cx", "資本支出"); row("dwc", "營運資金增加")
        row("fcf", "自由現金流（淨利 − 持分法 ＋ 折舊 − 資本支出 − 營運資金增加）")
        row("divp", "股利支付（每股股利 × 評價日股數）")
        row("fcfd", "自由現金流（扣股利）")
        for d in ("fa", "rb", "rm"):
            k = f"o_{d}"
            put(k, 2025, f"=h_o{d}_25")
            put(k, 2026, f"=h_o{d}_h1+2*Q!I{ {'fa': 13, 'rb': 14, 'rm': 15}[d] }*(1+o2h_{d})")
            for y in range(2027, 2033):
                p = col_of(y - 1)
                gr = f"(o_{d}{y - 2000}_{s}+o_shift)" if y <= 2029 else f"glt_{s}"
                put(k, y, f"={p}{R[k]}*(1+{gr})")
        for y in range(2025, 2033):
            c = col_of(y)
            put("o_sum", y, f"=SUM({c}{R['o_fa']}:{c}{R['o_rm']})")
            for d in ("fa", "rb", "rm"):
                put(f"bb_{d}", y, f"={c}{R['o_' + d]}/{c}{R['r_' + d]}")
        for d in ("fa", "rb", "rm", "sv"):
            k = f"r_{d}"
            put(k, 2024, f"=h_r{d}_24"); put(k, 2025, f"=h_r{d}_25")
            put(k, 2026, f"=h_r{d}_h1*(2+h2_{d})")
            for y in range(2027, 2033):
                c, p = col_of(y), col_of(y - 1)
                if d == "sv":
                    gr = f"gsv_{s}" if y <= 2029 else f"glt_{s}"
                    put(k, y, f"={p}{R[k]}*(1+{gr})")
                else:
                    put(k, y, f"=(1-lag_{d})*{c}{R['o_' + d]}+lag_{d}*{p}{R['o_' + d]}")
        for y in YEARS:
            c = col_of(y)
            ssum = f"SUM({c}{R['r_fa']}:{c}{R['r_sv']})"
            if y <= 2025:
                put("rev", y, f"=h_rev_{y - 2000}")
                put("chk", y, f"=IF(ABS({ssum}-{c}{R['rev']})<=0.1,0,{ssum}-{c}{R['rev']})")
            else:
                put("rev", y, f"={ssum}")
                put("chk", y, f"={ssum}-{c}{R['rev']}")
            if y >= 2025:
                put("g", y, f"={c}{R['rev']}/{col_of(y - 1)}{R['rev']}-1")
        # 損益
        for y in YEARS:
            c, p = col_of(y), col_of(y - 1) if y > 2024 else None
            yy = y - 2000
            if y <= 2025:
                put("oi", y, f"=h_oi_{yy}"); put("eq", y, f"=h_eq_{yy}")
                put("nop", y, f"=h_ord_{yy}-h_oi_{yy}-h_eq_{yy}"); put("tax", y, f"=h_tax_{yy}")
                put("nci", y, f"=h_nci_{yy}"); put("ni", y, f"=h_ni_{yy}"); put("eps", y, f"=h_eps_{yy}"); put("dps", y, f"=h_dps_{yy}")
                put("da", y, f"=h_da_{yy}"); put("cx", y, f"=h_cx_{yy}"); put("dwc", y, f"=h_dwc_{yy}")
            else:
                if y == 2026:
                    put("oi", y, f"=h_oi_h1+opm_h2*({c}{R['rev']}-h_rev_h1)")
                    put("eq", y, f"=h_eq_25*IF(eq_link=1,{c}{R['r_rb']}/h_rrb_25,1)")
                    put("nop", y, f"=(h_ord_h1-h_oi_h1)*(2+nonop_h2)-{c}{R['eq']}")
                else:
                    put("oi", y, f"=$E{R['oi']}+inc_m*({c}{R['rev']}-$E{R['rev']})")
                    put("eq", y, f"=$E{R['eq']}*IF(eq_link=1,{c}{R['r_rb']}/$E{R['r_rb']},1)")
                    put("nop", y, "=nonop_o")
                put("tax", y, f"=({c}{R['ord']}-{c}{R['eq']})*tax_w")
                put("nci", y, f"=({c}{R['ord']}-{c}{R['tax']})*nci_r")
                put("ni", y, f"={c}{R['ord']}-{c}{R['tax']}-{c}{R['nci']}")
                put("eps", y, f"={c}{R['ni']}/sh_v")
                put("dps", y, f"={c}{R['eps']}*payout")
                if y == 2026:
                    put("da", y, "=h_da_h1*2")
                else:
                    put("da", y, f"={p}{R['da']}+({p}{R['cx']}-{p}{R['da']})/da_life")
                put("cx", y, f"={c}{R['rev']}*cx_r")
                put("dwc", y, f"=wc_r*({c}{R['rev']}-{p}{R['rev']})")
            put("opm", y, f"={c}{R['oi']}/{c}{R['rev']}")
            put("ord", y, f"={c}{R['oi']}+{c}{R['eq']}+{c}{R['nop']}") if y > 2025 else put("ord", y, f"=h_ord_{yy}")
            put("fcf", y, f"={c}{R['ni']}-{c}{R['eq']}+{c}{R['da']}-{c}{R['cx']}-{c}{R['dwc']}")
            put("divp", y, f"={c}{R['dps']}*sh_v")
            put("fcfd", y, f"={c}{R['fcf']}-{c}{R['divp']}")
        R["_end"] = rr
        blocks[s] = R
        r0 = rr + 2
    P.column_dimensions["A"].width = 52
    for cl in YCOLS: P.column_dimensions[cl].width = 11

    # ---------------- DCF
    D = wb.create_sheet("DCF")
    D["A1"] = "DCF — 標準版 FCFF（NOPAT＋折舊 − 資本支出 − 營運資金增加），WACC 折現到評價日 2026-09-30；期中慣例；FY2026 只計下半；明確期到 FY2032"; D["A1"].font = BOLD
    D["A2"] = "股權報酬：公司未揭露重大股票報酬（未取得），NOPAT 不另調整。持分法投資利益不在營業利益內，另以 eq_pe 倍數估值加回；現金＝現金及び預金＋有価証券；無有息負債。"
    dcf = {}
    r0 = 4
    for s in "ABC":
        R = blocks[s]; X = {}
        D.cell(r0, 1, f"情境 {s}（{SCN_NAME[s]}）").font = Font(bold=True, size=12)
        for j, lab in enumerate(["FY2026 下半"] + [f"FY{y}" for y in range(2027, 2033)]):
            c = D.cell(r0, 5 + j, lab); c.font = BOLD; c.fill = HFILL
        rr = r0 + 1
        labels = [("ebit", "EBIT（營業利益）"), ("nopat", "NOPAT"), ("da", "＋折舊"), ("cx", "− 資本支出"), ("dwc", "− 營運資金增加"),
                  ("fcff", "＝FCFF"), ("chk", "檢查：各列加總 − FCFF"), ("t", "折現期間（年，期中）"), ("df", "折現因子"), ("pv", "現值")]
        for k, lab in labels:
            D.cell(rr, 1, lab); X[k] = rr; rr += 1
        cols = "EFGHIJK"
        for j, cl in enumerate(cols):
            pc = "E" if j == 0 else col_of(2026 + j)  # PL 欄
            if j == 0:
                D[f"{cl}{X['ebit']}"] = f"=PL!E{R['oi']}-h_oi_h1"
                D[f"{cl}{X['da']}"] = f"=PL!E{R['da']}-h_da_h1"
                D[f"{cl}{X['cx']}"] = f"=PL!E{R['cx']}-h_cx_h1"
                D[f"{cl}{X['dwc']}"] = f"=PL!E{R['dwc']}-h_dwc_h1"
                D[f"{cl}{X['t']}"] = 0.25
            else:
                D[f"{cl}{X['ebit']}"] = f"=PL!{pc}{R['oi']}"
                D[f"{cl}{X['da']}"] = f"=PL!{pc}{R['da']}"
                D[f"{cl}{X['cx']}"] = f"=PL!{pc}{R['cx']}"
                D[f"{cl}{X['dwc']}"] = f"=PL!{pc}{R['dwc']}"
                D[f"{cl}{X['t']}"] = f"={cols[j - 1]}{X['t']}+{0.75 if j == 1 else 1}"
            D[f"{cl}{X['nopat']}"] = f"={cl}{X['ebit']}*(1-tax_w)"
            D[f"{cl}{X['fcff']}"] = f"={cl}{X['nopat']}+{cl}{X['da']}-{cl}{X['cx']}-{cl}{X['dwc']}"
            D[f"{cl}{X['chk']}"] = f"={cl}{X['nopat']}+{cl}{X['da']}-{cl}{X['cx']}-{cl}{X['dwc']}-{cl}{X['fcff']}"
            D[f"{cl}{X['df']}"] = f"=1/(1+wacc)^{cl}{X['t']}"
            D[f"{cl}{X['pv']}"] = f"={cl}{X['fcff']}*{cl}{X['df']}"
            for k in ("ebit", "nopat", "da", "cx", "dwc", "fcff", "pv"):
                D[f"{cl}{X[k]}"].number_format = "#,##0.0"
            D[f"{cl}{X['df']}"].number_format = "0.000"
        rr += 1
        items = [
            ("sumpv", "明確期 FCFF 現值合計", f"=SUM(E{X['pv']}:K{X['pv']})"),
            ("nopat33", "NOPAT FY2033＝NOPAT FY2032 ×（1＋g）", f"=K{X['nopat']}*(1+g_tv)"),
            ("tv", "永續價值（FY2032 末）", f"=IF(tv_alt=1,K{X['fcff']}*(1+g_tv)/(wacc-g_tv),nopat33_{s}*(1-g_tv/ronic)/(wacc-g_tv))"),
            ("pvtv", "永續價值現值", f"=tv_{s}/(1+wacc)^6.5"),
            ("ev", "企業價值 EV", f"=sumpv_{s}+pvtv_{s}"),
            ("tvsh", "永續價值占 EV", f"=IF(ev_{s}>0,pvtv_{s}/ev_{s},\"N/M\")"),
            ("cash", "＋現金及び預金＋有価証券（2026-09-30）", "=h_cash_h1+h_sec_h1"),
            ("eqv", "＋持分法投資價值（FY2027 持分法利益 × eq_pe）", f"=PL!F{R['eq']}*eq_pe"),
            ("nci", "− 非支配株主持分（帳面）", "=h_ncib_h1"),
            ("divx", "− 已決議中間配當（評價日後支付）", "=div_int*sh_v"),
            ("eqty", "股權價值（評價日）", f"=ev_{s}+cash_{s}+eqv_{s}-nci_{s}-divx_{s}"),
            ("ps", "每股價值（評價日）", f"=eqty_{s}/sh_v"),
            ("pdiv", "− 期間股利（FY2026 期末＋FY2027 中間，每股）", f"=payout*(PL!E{R['eps']}-h_eps_h1)+payout*PL!F{R['eps']}/2"),
            ("tp", "DCF 情境目標價（以 ke 推到目標時點，扣期間股利）", f"=ps_{s}*(1+ke)^t_tp-pdiv_{s}"),
        ]
        for k, lab, f in items:
            D.cell(rr, 1, lab); c = D.cell(rr, 5, f)
            c.number_format = "0.0%" if k == "tvsh" else ("#,##0" if k in ("ps", "tp", "pdiv") else "#,##0.0")
            defname(f"{k}_{s}", f"DCF!$E${rr}")
            X[k] = rr; rr += 1
        D.cell(X["tvsh"], 6, f"=IF(AND(ev_{s}>0,tvsh_{s}>=0,tvsh_{s}<0.8),\"通過（<80%）\",\"標示：≥80% 或 N/M\")")
        dcf[s] = X
        r0 = rr + 2
    D.column_dimensions["A"].width = 52
    for cl in "EFGHIJK": D.column_dimensions[cl].width = 12

    # ---------------- Val
    V = wb.create_sheet("Val")
    V["A1"] = "Val — 估值（DCF 50%＋同業本益比 Comparable 50%；三情境 25／50／25；約 12 個月目標價）"; V["A1"].font = BOLD
    vr = {}
    def vrow(r_, key, lab, f, fmt="#,##0", note=None):
        V.cell(r_, 1, lab); c = V.cell(r_, 2, f); c.number_format = fmt
        if note: V.cell(r_, 5, note)
        if key:
            defname(key, f"Val!$B${r_}"); vr[key] = r_
    V.cell(3, 1, "0.1 折現率").font = BOLD
    vrow(4, "ke", "股權成本 ke＝rf＋β ×（ERP＋CRP）", "=rf+beta*(erp+crp)", "0.00%")
    vrow(5, "wacc", "WACC（無有息負債：＝ke＋wacc_adj）", "=ke+wacc_adj", "0.00%", "權重：股權 100%（2026-09-30 無借入金、リース債務未單列）")
    vrow(6, "ronic", "RONIC（＝WACC，新投資價值中性）", "=wacc", "0.00%")
    vrow(7, None, "永續成長率 g", "=g_tv", "0.00%")
    vrow(8, "t_tp", "折現期間：評價日 → 目標時點（年）", "=(tp_date-val_date)/365", "0.00")
    vrow(9, None, "持有期間：現價日 → 目標時點（年）", "=(tp_date-px_date)/365", "0.00", "上檔空間不是年化報酬")
    V.cell(11, 1, "0.2 DCF 三情境").font = BOLD
    for j, h in enumerate(["情境 A", "情境 B", "情境 C", "加權"]):
        V.cell(11, 2 + j, h).font = BOLD
    rows = [("EV", "ev"), ("永續價值占 EV", "tvsh"), ("股權價值（評價日）", "eqty"), ("每股價值（評價日）", "ps"), ("DCF 情境目標價", "tp")]
    for i, (lab, k) in enumerate(rows):
        rr = 12 + i
        V.cell(rr, 1, lab)
        for j, s in enumerate("ABC"):
            V.cell(rr, 2 + j, f"={k}_{s}").number_format = "0.0%" if k == "tvsh" else "#,##0"
    V["E16"] = "=B16*scw_A+C16*scw_B+D16*scw_C"; V["E16"].number_format = "#,##0"
    defname("tp_dcf", "Val!$E$16")
    V.cell(18, 1, "0.3 Comparable（同業前瞻本益比 × 目標時點 NTM EPS）").font = BOLD
    for j, h in enumerate(["情境 A", "情境 B", "情境 C", "加權"]):
        V.cell(18, 2 + j, h).font = BOLD
    V["A19"] = "NTM EPS（目標時點 2027-11：FY2027 × 4/12＋FY2028 × 8/12）"
    V["A20"] = "Comparable（未調整）＝peer_pe × NTM EPS"
    V["A21"] = "Comparable（相對倍數調整後）＝× rel_m（進加權）"
    for j, s in enumerate("ABC"):
        R = blocks[s]; cl = "BCD"[j]
        V[f"{cl}19"] = f"=PL!F{R['eps']}*4/12+PL!G{R['eps']}*8/12"
        V[f"{cl}20"] = f"=peer_pe*{cl}19"
        V[f"{cl}21"] = f"={cl}20*rel_m"
        for rr_ in (19, 20, 21): V[f"{cl}{rr_}"].number_format = "#,##0"
    for rr_ in (19, 20, 21):
        V[f"E{rr_}"] = f"=B{rr_}*scw_A+C{rr_}*scw_B+D{rr_}*scw_C"; V[f"E{rr_}"].number_format = "#,##0"
    defname("tp_comp", "Val!$E$21"); defname("tp_comp_u", "Val!$E$20")
    V["F19"] = "分母：本模型情境 EPS（非共識）；倍數分子：同業 2026-11-16 PER（traders.co.jp，收盤 ÷ 各公司本期業績予想 EPS）× 2026-11-20 收盤 ÷ 2026-11-16 收盤（EPS 不變）。同業會計年度不一（SMC、Keyence、Okuma 3 月；Nabtesco、DMG Mori 12 月）"
    V.cell(23, 1, "0.4 加權目標價與評等").font = BOLD
    vrow(24, "tp_w", "加權目標價＝w_dcf × DCF＋(1−w_dcf) × Comparable", "=w_dcf*tp_dcf+(1-w_dcf)*tp_comp")
    vrow(25, None, "現價（2026-11-20 收盤）", "=px")
    vrow(26, "upside", "上檔空間", "=tp_w/px-1", "0.0%")
    vrow(27, None, "估值評等（只看上檔空間）", '=IF(upside>0.5,"強力買進",IF(upside>0.25,"買進",IF(upside>=-0.25,"中立",IF(upside>=-0.5,"放空","強力放空"))))', "General",
         "級距：> +50% 強力買進、> +25% 買進、≥ −25% 中立、≥ −50% 放空、其餘強力放空（target_summary 同一套）")
    V.cell(27, 1, "估值評等（只看上檔空間）")
    V["A28"] = "收支關卡：FY2026–FY2030 自由現金流（扣股利）合計（情境 A／B／C）"
    for j, s in enumerate("ABC"):
        R = blocks[s]
        V.cell(28, 2 + j, f"=SUM(PL!E{R['fcfd']}:I{R['fcfd']})").number_format = "#,##0.0"
    vrow(29, "gate", "入不敷出判定用（情境 B 合計）", "=C28", "#,##0.0")
    V["C29"] = '=IF(gate<0,"不通過（入不敷出）","通過")'
    V["A30"] = "投資建議"; V["B30"] = '=IF(AND(gate<0,OR(B27="強力買進",B27="買進")),"收支不足",B27)'
    vrow(31, None, "上檔空間距最近級距邊界（pt）", "=MIN(ABS(upside-0.5),ABS(upside-0.25),ABS(upside+0.25),ABS(upside+0.5))", "0.0%")
    vrow(32, "upside_c", "Comparable 目標價上檔空間（只作對照，評等不依此）", "=tp_comp/px-1", "0.0%")
    V["A33"] = "0.5 上檔空間三段歸因"; V["A33"].font = BOLD
    vrow(34, None, "DCF 貢獻＝w_dcf ×（DCF ÷ 現價 − 1）", "=w_dcf*(tp_dcf/px-1)", "0.0%")
    vrow(35, None, "未調整 Comparable 貢獻", "=(1-w_dcf)*(tp_comp_u/px-1)", "0.0%")
    vrow(36, None, "相對倍數調整貢獻", "=(1-w_dcf)*(tp_comp-tp_comp_u)/px", "0.0%")
    V.cell(38, 1, "0.6 同業（2026-11-20 收盤；PER＝2026-11-16 PER × 新收盤 ÷ 舊收盤，分母 EPS 不變）").font = BOLD
    for j, h in enumerate(["公司", "代號", "收盤", "PER", "標記", "來源"]):
        V.cell(39, 1 + j, h).font = BOLD
    for i, (n_, t_, p_, p0_, pe_, sp_) in enumerate(PEERS):
        rr = 40 + i
        V.cell(rr, 1, n_); V.cell(rr, 2, t_); c = V.cell(rr, 3, p_); c.font = BLUE
        V.cell(rr, 4, f"={pe_}*C{rr}*{sp_}/{p0_}" if sp_ != 1 else f"={pe_}*C{rr}/{p0_}").number_format = "0.00"
        V.cell(rr, 5, "Derived")
        V.cell(rr, 6, f"收盤：{S_YF}（Verified）；PER＝2026-11-16 PER {pe_}（https://www.traders.co.jp/stocks/{t_[:4]}/）× 2026-11-20 收盤"
               + (f" × 分割比率 {sp_}（2026-11-19 1 拆 {sp_}；舊收盤 {p0_:,} 為分割前口徑）" if sp_ != 1 else "")
               + f" ÷ 2026-11-16 收盤 {p0_:,}（分母 EPS 不變）")
    V["A45"] = "中位數"; V["D45"] = "=MEDIAN(D40:D44)"
    V["A46"] = "Fanuc（現行）"; V["B46"] = "6954.T"; V["C46"] = "=px"; V["D46"] = f"={PE_OLD}*px/{PX_OLD}"; V["D46"].number_format = "0.00"
    V["E46"] = "Derived"; V["F46"] = f"＝2026-11-16 PER {PE_OLD}（traders.co.jp，≈ {PX_OLD:,} ÷ 公司 FY2026 EPS 予想 246.01）× 現價 ÷ {PX_OLD:,}（分母 EPS 不變）"
    V["A47"] = "Fanuc 現行 PER ÷ 同業中位數"; V["D47"] = "=D46/D45"; V["D47"].number_format = "0.00"
    V["A48"] = "未取得：安川電機 6506、哈默納科 6324（頁面讀取失敗）；Fanuc 對同業 10 年相對倍數歷史。同業僅 5 家，不做迴歸；會計口徑皆為日本會計準則或 IFRS（Nabtesco、DMG Mori 為 IFRS），PER 為股權口徑、不受租賃處理影響。"
    V.column_dimensions["A"].width = 60
    for cl in "BCDE": V.column_dimensions[cl].width = 14

    # ---------------- Recon
    RC = wb.create_sheet("Recon")
    RC["A1"] = "Recon — 對照公司指引與反解（只展示，不進目標價）"; RC["A1"].font = BOLD
    for j, h in enumerate(["項目", "模型（情境 B）", "公司指引", "差（模型 − 指引）", "差 %", "說明"]):
        RC.cell(3, 1 + j, h).font = BOLD
    RB = blocks["B"]
    guid = [("FY2026 營收", f"=PL!E{RB['rev']}", 1013.7), ("FY2026 營業利益", f"=PL!E{RB['oi']}", 240.7),
            ("FY2026 經常利益", f"=PL!E{RB['ord']}", 308.8), ("FY2026 淨利", f"=PL!E{RB['ni']}", 228.3),
            ("FY2026 EPS（¥）", f"=PL!E{RB['eps']}", 246.01),
            ("FY2026 下半營收", f"=PL!E{RB['rev']}-h_rev_h1", "=1013.7-473.708"),
            ("FY2026 下半營業利益率", f"=(PL!E{RB['oi']}-h_oi_h1)/(PL!E{RB['rev']}-h_rev_h1)", "=(240.7-116.662)/(1013.7-473.708)"),
            ("FY2026 下半營業外收支", f"=PL!E{RB['ord']}-PL!E{RB['oi']}-(h_ord_h1-h_oi_h1)", "=(308.8-149.045)-(240.7-116.662)")]
    for i, (lab, m, g) in enumerate(guid):
        rr = 4 + i
        RC.cell(rr, 1, lab); RC.cell(rr, 2, m); c = RC.cell(rr, 3, g); c.font = BLUE
        RC.cell(rr, 4, f"=B{rr}-C{rr}"); RC.cell(rr, 5, f"=B{rr}/C{rr}-1")
        RC.cell(rr, 6, "公司 2026-10-26 修正予想（Interested-party：發行公司；下半匯率假設 150 円／美元）")
        for cc in (2, 3, 4): RC.cell(rr, cc).number_format = "0.0%" if "率" in lab else "#,##0.0"
        RC.cell(rr, 5).number_format = "0.0%"
    RC["A13"] = "方向（公式組字）："
    RC["B13"] = '="FY2026 營收模型"&IF(B4>C4,"高於","低於")&"指引 "&TEXT(ABS(E4),"0.0%")&"；營業利益"&IF(B5>C5,"高於","低於")&"指引 "&TEXT(ABS(E5),"0.0%")&"；EPS "&IF(B8>C8,"高於","低於")&"指引 "&TEXT(ABS(E8),"0.0%")'
    RC["A14"] = "市場共識（FY2026／FY2027 EPS）"; RC["B14"] = "未找到（本輪未取得可引用的共識來源）"
    RC["A16"] = "反解（整本重算，o_shift＝FY2027–FY2029 三部門訂單年增三情境同加的平移）"; RC["A16"].font = BOLD
    for j, h in enumerate(["條件", "所需 o_shift", "說明"]):
        RC.cell(17, 1 + j, h).font = BOLD
    rv = rev or {}
    for i, (lab, k, note) in enumerate([("加權目標價＝現價（上檔空間 0%）", "zero", "反解值是目標價＝現價的點，不是評等門檻"),
                                        ("升一級：上檔空間＝+25%（買進）", "up", "評等翻轉條件"),
                                        ("降一級：上檔空間＝−25%（放空）", "dn", "評等翻轉條件")]):
        rr = 18 + i
        RC.cell(rr, 1, lab); v = rv.get(k)
        c = RC.cell(rr, 2, v if v is not None else "未解（超出 ±60pt 搜尋區間）"); c.number_format = "+0.0%;-0.0%"
        RC.cell(rr, 3, note)
    RC["A22"] = "對照：情境 B FY2027 FA 訂單年增（基準，未含平移）"; RC["B22"] = "=o_fa27_B"; RC["B22"].number_format = "0.0%"
    RC.column_dimensions["A"].width = 48; RC.column_dimensions["F"].width = 60

    # ---------------- 對帳
    DZ = wb.create_sheet("對帳")
    DZ["A1"] = "對帳 — 本模型用到的線層級變數（與五條線單位經濟模型互相對帳；本模型不改線的口徑）"; DZ["A1"].font = BOLD
    for j, h in enumerate(["變數名", "所屬線", "本模型使用值（或區間）", "單位", "口徑", "出處", "標記", "影響本模型哪些儲存格"]):
        DZ.cell(3, 1 + j, h).font = BOLD; DZ.cell(3, 1 + j).fill = HFILL
    for i, row_ in enumerate(RECON_LINES):
        for j, v in enumerate(row_):
            DZ.cell(4 + i, 1 + j, v).alignment = WRAP
    DZ.column_dimensions["A"].width = 30; DZ.column_dimensions["C"].width = 30; DZ.column_dimensions["F"].width = 40; DZ.column_dimensions["H"].width = 60

    # ---------------- Risk
    RK = wb.create_sheet("Risk")
    RK["A1"] = "風險登錄（原則第 10 節第 10、19 條）"; RK["A1"].font = BOLD
    for j, h in enumerate(["風險", "是否建模", "理由", "觀察指標"]):
        RK.cell(3, 1 + j, h).font = BOLD
    for i, row_ in enumerate(RISKS):
        for j, v in enumerate(row_):
            RK.cell(4 + i, 1 + j, v).alignment = WRAP
    for cl, w in zip("ABCD", (50, 30, 40, 50)): RK.column_dimensions[cl].width = w

    # ---------------- Sens
    SE = wb.create_sheet("Sens")
    SE["A1"] = "Sens — 一次動一個輸入（In 頁 D／F 欄），整本重算（LibreOffice）的快照；Comparable 也隨 EPS 重算"; SE["A1"].font = BOLD
    sd = sens or {}
    SE["A2"] = "build_id"; SE["D2"] = sd.get("build_id", "未建置")
    SE["A3"] = "模型狀態"
    SE["B3"] = sd.get("fp", 0); SE["C3"] = f"=ROUND(SUM(In!E4:E{in_last}),6)"
    SE["D3"] = '=IF(ROUND(B3,6)=C3,"CURRENT","STALE")'
    SE["E3"] = "B＝敏感度重算時的 In 頁基準欄指紋；C＝本檔即時指紋"
    hdr = ["案例", "設定", "鍵", "DCF 目標價", "加權目標價", "相對基準", "上檔空間", "估值評等", "說明"]
    for j, h in enumerate(hdr):
        SE.cell(5, 1 + j, h).font = BOLD; SE.cell(5, 1 + j).fill = HFILL
    rr = 6
    for case in sd.get("cases", []):
        SE.cell(rr, 1, case["name"]); SE.cell(rr, 2, case["set"]); SE.cell(rr, 3, case["key"])
        SE.cell(rr, 4, case["dcf"]).number_format = "#,##0"; SE.cell(rr, 5, case["tp"]).number_format = "#,##0"
        SE.cell(rr, 6, case["rel"]).number_format = "+0.0%;-0.0%"; SE.cell(rr, 7, case["up"]).number_format = "+0.0%;-0.0%"
        SE.cell(rr, 8, case["tier"]); SE.cell(rr, 9, case.get("note", ""))
        rr += 1
    SE.cell(rr + 1, 1, "說明：nci_r、payout、rf、crp、div_int 等無低／高區間的輸入不列。o_*_A／C、gsv、glt 為情境專屬輸入（三情境即區間），不另做單參數案例。")
    SE.column_dimensions["A"].width = 40; SE.column_dimensions["B"].width = 34; SE.column_dimensions["I"].width = 50

    # ---------------- Sources
    SO = wb.create_sheet("Sources")
    SO["A1"] = "來源"; SO["A1"].font = BOLD
    for i, s_ in enumerate(SOURCES):
        SO.cell(3 + i, 1, s_)
    SO.column_dimensions["A"].width = 160

    # ---------------- README（第一頁）
    RD = wb.create_sheet("README", 0)
    bid = sd.get("build_id", "未建置")
    RD["A1"] = f"Fanuc（ファナック，6954.T）個股模型 {VER}｜build_id {bid}｜資料截止日 2026-11-16（股價 2026-11-20）｜成品日 2026-11-23｜股價基準日 2026-11-20（收盤）"
    RD["A1"].font = Font(bold=True, size=13)
    RB, XB = blocks["B"], dcf["B"]
    lines = [
        ("命題", "Fanuc 的訂單在 FY2026 上半出現 FA（CNC）主導的跳升（7–9 月 FA 訂單 2.3 倍，AI／資料中心／半導體與人形零件加工需求），Robot 訂單溫和成長（+12.6%）。本模型問：這波訂單是否會在 FY2027 轉成營收與利潤，以及循環回落時現價隱含什麼。實體 AI 角度：Fanuc 同時是工業機器人整機（線 2a）與人形／AI 硬體零件加工設備（CNC）的供應商。"),
        ("結論（公式組字）", '="在情境 B 訂單傳導（FA 營收中上一年訂單比重 "&TEXT(lag_fa,"0%")&"）與 β "&TEXT(beta,"0.00")&" 的假設下：加權目標價 "&TEXT(tp_w,"#,##0")&" 円（DCF "&TEXT(tp_dcf,"#,##0")&"、Comparable "&TEXT(tp_comp,"#,##0")&"），現價 "&TEXT(px,"#,##0")&" 円，上檔空間 "&TEXT(upside,"+0.0%;-0.0%")&"，估值評等 "&Val!B27&"；收支關卡 "&Val!C29&"（情境 B FY2026–30 自由現金流扣股利合計 "&TEXT(gate,"#,##0")&" 十億日圓）；投資建議 "&Val!B30&"。"'),
        ("距級距邊界", '="上檔空間距最近級距邊界 "&TEXT(Val!B31,"0.0%")&"（"&IF(Val!B31<0.05,"不到 5 個百分點，評等貼近邊界","5 個百分點以上")&"）。"'),
        ("三段歸因", '="上檔空間＝DCF 貢獻 "&TEXT(Val!B34,"+0.0%;-0.0%")&"＋未調整 Comparable 貢獻 "&TEXT(Val!B35,"+0.0%;-0.0%")&"＋相對倍數調整 "&TEXT(Val!B36,"+0.0%;-0.0%")&"；Fanuc 現行 PER 為同業中位數的 "&TEXT(Val!D47,"0.00")&" 倍，本模型未給溢價（rel_m＝1）。"'),
        ("FY2026 對指引", "=Recon!B13"),
        ("DCF 情境", '="DCF 情境 A／B／C："&TEXT(tp_A,"#,##0")&"／"&TEXT(tp_B,"#,##0")&"／"&TEXT(tp_C,"#,##0")&" 円；WACC "&TEXT(wacc,"0.00%")&"、g "&TEXT(g_tv,"0.00%")&"、情境 B 永續價值占 EV "&TEXT(tvsh_B,"0%")&"。"'),
        ("失效條件", "（1）FA 訂單連兩季 QoQ 回落 20% 以上且中國訂單同步下滑（情境 C 路徑）；（2）下半營收明顯低於指引 540 十億日圓（零件短缺）；（3）營業利益率跌回 20% 以下；（4）Robot 訂單年增轉負（中國價格戰）。"),
        ("可證偽指標", "lag_fa → 每季 FA 營收 ÷ 前兩季 FA 訂單；o_fa27_* → FY2027 每季 FA 訂單（5、8、11、2 月公布）；inc_m → 每季營業利益率；beta → 自算 β（下一版補）。"),
        ("協調檔", "實體 AI 線模型不讀 Siliconomics／Tokenomics／Wattonomics 協調檔，DCF 規格中 coord、兩讀數、Wattonomics 相關條款不適用；自用 AI 讀法（第 7 節）不適用。"),
        ("限制", "公司為單一報告分部，不揭露部門利益與受注殘；營業利益以合併增量利潤率推得。β 未自算（Assumed）。持分法投資（中國合資）以倍數估值，投資有価証券帳面不另加。HTML 展示層見 dist 同名 .html（v0.1.1 起）。"),
    ]
    # 展示層（HTML 讀取；v0.1.1 新增，白話、公式帶數字；上方原列保留）
    SENS_REL = 'INDEX(Sens!E:E,MATCH("rel_m",Sens!C:C,0))'
    lines += [
        ("展示層：命題分項 1", "FY2026 上半訂單 ÷ 營收明顯高於 1，主因 FA（CNC）：公司稱需求來自 AI、資料中心、半導體與人形機器人零件加工，以中國為主。"),
        ("展示層：命題分項 2", "同期 Robot 訂單成長溫和，明顯低於日本機器人工業會全體受注增速。"),
        ("展示層：命題分項 3", "要回答：這波訂單會不會在 FY2027 轉成營收與利潤？循環回落時，現價隱含了什麼？"),
        ("展示層：橫幅說明", '="加權目標價 "&TEXT(tp_w,"#,##0")&" 円"&IF(upside<0,"低於","高於")&"現價，上檔空間 "&TEXT(upside,"+0.0%;-0.0%")&"，投資建議"&Val!B30&"。主因：基準情境假設 FA 訂單 FY2027、FY2028 各年增 "&TEXT(o_fa27_B,"+0%;-0%")&"、"&TEXT(o_fa28_B,"+0%;-0%")&"，且不給 Fanuc 對同業的本益比溢價；若維持現行溢價，加權目標價 "&TEXT(' + SENS_REL + ',"#,##0")&" 円。"'),
        ("展示層：重點 1", '="估值：DCF（情境加權）"&TEXT(tp_dcf,"#,##0")&" 円、同業本益比法 "&TEXT(tp_comp,"#,##0")&" 円，各占一半，加權目標價 "&TEXT(tp_w,"#,##0")&" 円，對現價 "&TEXT(px,"#,##0")&" 円的上檔空間 "&TEXT(upside,"+0.0%;-0.0%")&"，落在中立區間（距最近級距邊界 "&TEXT(Val!B31*100,"0.0")&" 個百分點）。"'),
        ("展示層：重點 2", '="訂單路徑是最承重的變數：基準情境 FY2027 營收 "&TEXT(PL!F' + str(RB["rev"]) + ',"#,##0")&" 十億日圓、營業利益 "&TEXT(PL!F' + str(RB["oi"]) + ',"#,##0")&"（FY2026 為 "&TEXT(PL!E' + str(RB["rev"]) + ',"#,##0")&"、"&TEXT(PL!E' + str(RB["oi"]) + ',"#,##0")&"）；保守情境 FY2027 FA 訂單年增 "&TEXT(o_fa27_C,"+0%;-0%")&"，DCF 目標價 "&TEXT(tp_C,"#,##0")&" 円。"'),
        ("展示層：重點 3", '="同業倍數：Fanuc 現行本益比 "&TEXT(Val!D46,"0.0")&" 倍，是同業中位數 "&TEXT(peer_pe,"0.0")&" 倍的 "&TEXT(Val!D47,"0.00")&" 倍；本模型不給溢價，若維持現行溢價，加權目標價 "&TEXT(' + SENS_REL + ',"#,##0")&" 円（上檔空間 "&TEXT(' + SENS_REL.replace("E:E", "G:G") + ',"+0.0%;-0.0%")&"）。"'),
        ("展示層：重點 4", '="反解：三部門 FY2027–FY2029 訂單年增在三情境同時上調 "&TEXT(Recon!B18*100,"0.0")&" 個百分點，加權目標價才等於現價；升到買進需上調 "&TEXT(Recon!B19*100,"0.0")&" 個百分點，降到放空需下調 "&TEXT(-Recon!B20*100,"0.0")&" 個百分點。"'),
        ("展示層：重點 5", '="收支關卡"&Val!C29&"：基準情境 FY2026–FY2030 自由現金流扣股利合計 "&TEXT(gate,"#,##0.0")&" 十億日圓；其中 FY2026 "&TEXT(PL!E' + str(RB["fcfd"]) + ',"+#,##0.0;-#,##0.0")&"、FY2027 "&TEXT(PL!F' + str(RB["fcfd"]) + ',"+#,##0.0;-#,##0.0")&"（營運資金隨訂單增加）。"'),
        ("展示層：命題回覆", '="這波訂單會不會在 FY2027 轉成營收與利潤？答：基準情境 FY2027 營收年增 "&TEXT(PL!F' + str(RB["g"]) + ',"+0.0%;-0.0%")&"、營業利益率 "&TEXT(PL!F' + str(RB["opm"]) + ',"0.0%")&"，FY2028 營收年增 "&TEXT(PL!G' + str(RB["g"]) + ',"+0.0%;-0.0%")&"（訂單回落）。循環回落時現價隱含什麼？答：現價 "&TEXT(px,"#,##0")&" 円需要三部門訂單年增比基準高 "&TEXT(Recon!B18*100,"0.0")&" 個百分點，或維持對同業的本益比溢價。"'),
    ]
    for i, (k, v) in enumerate(lines):
        RD.cell(3 + i, 1, k).font = BOLD; c = RD.cell(3 + i, 2, v); c.alignment = WRAP
    RD.column_dimensions["A"].width = 18; RD.column_dimensions["B"].width = 140

    order = ["README", "In", "Q", "Hist", "PL", "DCF", "Val", "Sens", "Recon", "對帳", "Risk", "Sources"]
    wb._sheets = [wb[n] for n in order]
    os.makedirs(os.path.dirname(path), exist_ok=True)
    wb.save(path)
    return dict(inrow=inrow, blocks=blocks, dcf=dcf, in_last=in_last)


def recalc(paths, outdir):
    subprocess.run(["soffice", "--headless", "--calc", "--convert-to", "xlsx", "--outdir", outdir] + list(paths),
                   check=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, timeout=600)
    return [os.path.join(outdir, os.path.basename(p)) for p in paths]


def readv(path):
    wb = openpyxl.load_workbook(path, data_only=True)
    V = wb["Val"]
    return dict(tp=V["B24"].value, dcf=V["E16"].value, up=V["B26"].value, tier=V["B27"].value, comp=V["E21"].value)


def base_inputs():
    return {it[0]: it for it in INROWS if it[0] != "sec"}


def fingerprint(path):
    wb = openpyxl.load_workbook(path, data_only=True)
    return wb["Sens"]["C3"].value


def run_cases(cases, tmp):
    """cases: list of (name, over)；回傳每案 readv 結果"""
    paths = []
    for i, (name, ov) in enumerate(cases):
        p = os.path.join(tmp, f"c{i:03d}.xlsx")
        build(over=ov, path=p)
        paths.append(p)
    outd = os.path.join(tmp, "out"); os.makedirs(outd, exist_ok=True)
    res = []
    for k in range(0, len(paths), 20):
        res += recalc(paths[k:k + 20], outd)
    return [readv(p) for p in res]


def main():
    quick = "--quick" in sys.argv
    tmp = tempfile.mkdtemp(prefix="fanuc_")
    sens_f = os.path.join(HERE, "sens.json")
    bi = base_inputs()
    # build_id：版本＋建置輸入雜湊
    h = hashlib.sha1(json.dumps([INROWS, HIST, QDATA, PEERS], ensure_ascii=False, default=str).encode()).hexdigest()[:10]
    build_id = f"{VER}-{h}"
    if quick and os.path.exists(sens_f):
        sens = json.load(open(sens_f, encoding="utf-8"))
    else:
        # 基準
        base = run_cases([("base", {})], tmp)[0]
        cases = [("基準", "—", "—", {})]
        for k, it in bi.items():
            lo, b, hi = it[3], it[4], it[5]
            for lab, v in (("低", lo), ("高", hi)):
                if v is None or v == b:
                    continue
                cases.append((f"{it[1]}：敏感度{lab}", f"{k}＝" + ("現行 PER ÷ 同業中位數" if k == "rel_m" and isinstance(v, str) else (v if isinstance(v, str) else str(round(v, 4)))), k, {k: v}))
        out = run_cases([(c[0], c[3]) for c in cases], tmp)
        rows = []
        for c, o in zip(cases, out):
            rows.append(dict(name=c[0], set=c[1], key=c[2], dcf=o["dcf"], tp=o["tp"], rel=o["tp"] / base["tp"] - 1, up=o["up"], tier=o["tier"],
                             note="整本重算" + ("；β 低／高 ±0.25" if c[2] == "beta" else "")))
        # 加 DCF 單法顯示由 dcf 欄
        sens = dict(build_id=build_id, cases=rows, base=base)
        # 反解 o_shift
        rev = {}
        px = bi["px"][4]
        for tag, target in (("zero", 0.0), ("up", 0.25), ("dn", -0.25)):
            lo_, hi_ = -0.6, 0.6
            f_lo = run_cases([("x", {"o_shift": lo_})], tmp)[0]["up"] - target
            f_hi = run_cases([("x", {"o_shift": hi_})], tmp)[0]["up"] - target
            if f_lo * f_hi > 0:
                rev[tag] = None; continue
            for _ in range(12):
                mid = (lo_ + hi_) / 2
                fm = run_cases([("x", {"o_shift": mid})], tmp)[0]["up"] - target
                if fm * f_lo <= 0:
                    hi_ = mid
                else:
                    lo_, f_lo = mid, fm
            rev[tag] = round((lo_ + hi_) / 2, 4)
        sens["rev"] = rev
        json.dump(sens, open(sens_f, "w", encoding="utf-8"), ensure_ascii=False, indent=1, default=str)
    sens["build_id"] = build_id
    # 指紋：先建一次讀即時指紋
    tmpb = os.path.join(tmp, "fp.xlsx")
    build(sens=dict(sens, fp=0), rev=sens.get("rev"), path=tmpb)
    fp = fingerprint(recalc([tmpb], os.path.join(tmp, "fpout"))[0])
    sens["fp"] = fp
    build(sens=sens, rev=sens.get("rev"), path=OUT)
    # 重算後覆蓋成品（快取值），檢查錯誤值
    outd = os.path.join(tmp, "final"); os.makedirs(outd, exist_ok=True)
    rc = recalc([OUT], outd)[0]
    wb = openpyxl.load_workbook(rc, data_only=True)
    errs = []
    for ws in wb:
        for row in ws.iter_rows():
            for c in row:
                if isinstance(c.value, str) and c.value.startswith("#") and c.value[1:4].isupper():
                    errs.append(f"{ws.title}!{c.coordinate}={c.value}")
                if isinstance(c.value, str) and c.value in ("Err:502", "Err:504", "Err:508", "Err:522"):
                    errs.append(f"{ws.title}!{c.coordinate}={c.value}")
    print("重算錯誤格：", len(errs), errs[:10])
    shutil.copy(rc, OUT)   # 成品帶快取值（LibreOffice 重算）
    v = readv(OUT)
    print("加權目標價", round(v["tp"]), "DCF", round(v["dcf"]), "Comparable", round(v["comp"]), "上檔", f"{v['up']:+.1%}", v["tier"])
    print("反解", sens.get("rev"))
    shutil.rmtree(tmp, ignore_errors=True)
    return 1 if errs else 0


if __name__ == "__main__":
    sys.exit(main())
