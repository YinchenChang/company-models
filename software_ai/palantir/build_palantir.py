#!/usr/bin/env python3
"""Palantir（PLTR）個股模型 v0.1：命題（原型 U 按用量計價的單位經濟）＋官方分部（Government／Commercial）→ 損益表。
用法：python3 build_palantir.py <輸出 raw.xlsx> [--sens sens.json] [--set 鍵=值 ...]
環境變數 ROWMAP＝輸出列號對照 json（供 gen_html_json.py 與 sens_pltr.py 使用）。
全部活公式（openpyxl 寫入、LibreOffice 重算）；藍字＝輸入、黑字＝公式、黃底＝判斷題。金額 US$M；股數百萬股；EPS 美元。
本版未做先驗（工作單 G 節）。"""
import argparse, datetime, hashlib, json, os, sys
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter as CL

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.join(HERE, "data"))
from facts_palantir import FACTS, V, PEERS, FACT_BY_ID  # noqa: E402

ap = argparse.ArgumentParser()
ap.add_argument("out"); ap.add_argument("--sens"); ap.add_argument("--set", action="append", default=[])
A = ap.parse_args()
OVR = {k: float(v) for k, v in (s.split("=", 1) for s in A.set)}
SENS = json.load(open(A.sens, encoding="utf-8")) if A.sens and os.path.exists(A.sens) else {}

VER = "v0.1"; VDATE = "2026-11-16"
def _build_id():
    h = hashlib.sha256()
    for fn in ["build_palantir.py", "data/facts_palantir.py", "data/tokenomics_snapshot_v22.0.json"]:
        h.update(fn.encode()); h.update(open(os.path.join(HERE, fn), "rb").read())
    return f"palantir-{VER}-{h.hexdigest()[:10]}"
BUILD_ID = _build_id()
TK = json.load(open(os.path.join(HERE, "data", "tokenomics_snapshot_v22.0.json"), encoding="utf-8"))
TKI = TK["items"]; TKV = TK["source"]["version"]

BLUE = Font(color="0000FF"); BOLD = Font(bold=True); HDR = Font(bold=True, color="FFFFFF")
YEL = PatternFill("solid", fgColor="FFF2CC"); HFILL = PatternFill("solid", fgColor="0B7A80"); GREY = PatternFill("solid", fgColor="EDEDED")
wb = openpyxl.Workbook(); wb.remove(wb.active)
SH = {n: wb.create_sheet(n) for n in ["README", "In", "Q", "Seg", "PL", "CF", "DCF", "Val", "RevDCF", "Recon", "Sens", "Sources"]}
SC = ["B", "A", "C"]                       # 情境排列：B（基準）先
SCN_NAME = {"A": "情境 A（樂觀）", "B": "情境 B（基準）", "C": "情境 C（保守）"}
YRS = list(range(2023, 2033)); FC = list(range(2026, 2033))
def yc(y): return CL(3 + y - 2023)          # 2023→C … 2032→L
MCOL = "M"                                   # 9M26A


def put(ws, r, c, v, fmt=None, bold=False, blue=False, yel=False, wrap=False):
    cell = ws.cell(r, c, v)
    if fmt: cell.number_format = fmt
    if bold: cell.font = BOLD
    if blue: cell.font = BLUE
    if yel: cell.fill = YEL
    if wrap: cell.alignment = Alignment(wrap_text=True, vertical="top")
    return cell


# =====================================================================================================
# Q 頁：事實總帳（每列單一標記；Interested-party 列有來源誘因）
Q = SH["Q"]
Q["A1"] = "Q — 事實總帳（US$M；股數百萬股；EPS 美元）。每列一個數字：指標、期間、口徑、出處、標記；來源與擷取日見 data/facts_palantir.csv（本頁與 csv 同源）。公司自揭 KPI、指引、賣方預估＝Interested-party；由原文相加相除＝Derived"
Q["A2"] = "核對列（實績必須為 0）：Government＋Commercial－合併營收（見本頁底部與 PL 頁）。財報科目以 SEC 10-K／10-Q／8-K 為主來源，資料商與媒體只作第二來源；找不到寫「未找到」，見 README 與 Recon。"
for j, h in enumerate(["指標", "期間", "數值", "單位", "口徑", "來源", "標記", "來源誘因（Interested-party 列；混合來源列寫其他欄位）", "擷取日", "事實 id"], 1):
    c = Q.cell(3, j, h); c.font = HDR; c.fill = HFILL
QR = {}
for i, x in enumerate(FACTS):
    r = 4 + i; QR[x["id"]] = r
    put(Q, r, 1, x["指標"]); put(Q, r, 2, x["期間"]); put(Q, r, 3, x["數值"], blue=True); put(Q, r, 4, x["單位"]); put(Q, r, 5, x["口徑"]); put(Q, r, 6, x["出處"])
    put(Q, r, 7, x["標記"]); put(Q, r, 8, x["來源誘因"] or None); put(Q, r, 9, x["擷取日"]); put(Q, r, 10, x["id"])
def q(i): return f"Q!$C${QR[i]}"
QLAST = 4 + len(FACTS) - 1
r = QLAST + 2
put(Q, r, 1, "檢查：Government＋Commercial－合併營收（FY2021–FY2025；應為 0）", bold=True); QCHK = r
for k, y in enumerate(range(2021, 2026)):
    put(Q, r, 3 + k, f"={q(f'seg.gov.{y}')}+{q(f'seg.com.{y}')}-{q(f'rev.{y}')}", "0.000")
r += 1; put(Q, r, 1, "檢查：Government＋Commercial－合併營收（最近四季；應為 0）", bold=True); QCHK2 = r
for k, qq in enumerate(["2025Q4", "2026Q1", "2026Q2", "2026Q3"]):
    put(Q, r, 3 + k, f"={q(f'seg.gov.{qq}')}+{q(f'seg.com.{qq}')}-{q(f'rev.{qq}')}", "0.000")
for col, w in zip("ABCDEFGHIJ", [46, 18, 14, 8, 36, 60, 16, 36, 12, 18]): Q.column_dimensions[col].width = w

# =====================================================================================================
# In 頁
IN = SH["In"]
INS = []      # (kind, ...)
def sec(t): INS.append(("sec", t))
def inp(key, label, unit, lo, base, hi, tag, src, note="", inc="", fmt=None):
    INS.append(("row", dict(key=key, label=label, unit=unit, lo=lo, base=base, hi=hi, tag=tag, src=src, note=note, inc=inc, fmt=fmt)))
def qv(i): return V(i)
pct = "0.0%"
sec("A. 日期、現價、股數與資產負債（評價日 2026-09-30；現價日 2026-11-13）")
inp("data_date", "資料截止日（本版最晚一筆輸入：StockAnalysis 同業頁更新日）", "日期", None, datetime.datetime(2026, 11, 15), None, "Derived", "facts_palantir.py 同業列最晚頁面更新日 2026-11-15（LDOS）；財報原文截至 2026-11-03 的 10-Q", "原則第 41 條：README 首列分列資料截止日、成品日、股價基準日", fmt="yyyy-mm-dd")
inp("px_date", "股價基準日（現價日）", "日期", None, datetime.datetime(2026, 11, 13), None, "Verified", "同 px 列來源（2026-11-13 收盤）", "原則第 26 條：目標時點＝本值＋12 個月；比成品日早 3 天（11-14、11-15 為週末），未逾 2 個交易日門檻", fmt="yyyy-mm-dd")
inp("val_date", "評價日（最近一季資產負債表日）", "日期", None, datetime.datetime(2026, 9, 30), None, "Verified", "Palantir 2026Q3 10-Q 資產負債表日", "評價日早於現價日本身不改；DCF 折現期間自評價日起、持有期間自現價日起（Val 頁並列）", fmt="yyyy-mm-dd")
INS.append(("tpdate",))
inp("px", "現價（2026-11-13 收盤）", "US$", None, V("mkt.px.1113"), None, "Verified", "StockAnalysis PLTR 歷史股價表 Nov 13, 2026 收盤 167.99（https://stockanalysis.com/stocks/pltr/history/）", "11-13 為單一來源；11-09 收盤 169.20 與 heygotrade 報導相符。只影響上檔空間、評等與選擇權庫藏股股數", fmt="0.00")
inp("sh_a", "Class A 流通股（2026-10-26 封面）", "百萬股", None, V("sh.a"), None, "Verified", "10-Q 2026Q3 封面", fmt="#,##0.000")
inp("sh_b", "Class B 流通股", "百萬股", None, V("sh.b"), None, "Verified", "同上", fmt="#,##0.000")
inp("sh_f", "Class F 流通股", "百萬股", None, V("sh.f"), None, "Verified", "同上", fmt="#,##0.000")
inp("rsu_n", "未歸屬 RSU（全額計入；沒收率未揭露）", "百萬股", None, V("sh.rsu"), None, "Verified", "10-Q Note 9", "DCF 標準規格第 3 節：RSU 全額進基準；P-RSU 0.14 百萬股數量小未計入（列為已知差異）", fmt="#,##0.000")
inp("opt_n", "選擇權流通數", "百萬股", None, V("sh.opt"), None, "Verified", "10-Q Note 9", "庫藏股法，價格用現價", fmt="#,##0.000")
inp("opt_k", "選擇權加權履約價", "US$", None, V("sh.optk"), None, "Verified", "10-Q Note 9", fmt="0.00")
inp("sar_n", "SAR 流通數", "百萬股", None, V("sh.sar"), None, "Verified", "10-Q Note 9", "股票增值權以選擇權庫藏股法處理（履約價 134.77 低於現價才有稀釋）", fmt="#,##0.000")
inp("sar_k", "SAR 加權履約價", "US$", None, V("sh.sark"), None, "Verified", "10-Q Note 9", fmt="0.00")
inp("cash_0", "現金及約當現金（2026-09-30）", "US$M", None, V("bs.cash"), None, "Verified", "10-Q 資產負債表 p.3", fmt="#,##0.0")
inp("sec_0", "有價證券（短期美國公債；2026-09-30）", "US$M", None, V("bs.sec"), None, "Verified", "10-Q 資產負債表 p.3", "新聞稿標題寫現金與公債 10.4B，資產負債表兩項相加 10.56B，公司未說明差異；本模型採資產負債表", fmt="#,##0.0")
inp("debt_0", "借款餘額（循環額度未動用）", "US$M", None, V("bs.debt"), None, "Verified", "10-Q Note 6", "營業租賃負債 225.3（非流動）不計為債務：營業利益已含租金費用（避免重複）", fmt="#,##0.0")
sec("B. 折現率與永續（DCF 標準規格第 2、3 節）")
inp("rf", "無風險利率（10 年美債，2026-10-27）", "%", None, V("mkt.rf"), None, "Verified", "沿用 Meta v0.2.1 In 頁 rf（美國財政部每日殖利率曲線 2026-10-27 4.99%）", "跨模型共用取值，逾期（距現價日 12 個交易日）時維持跨模型一致，下次全面改版更新（原則第 15 條）", fmt="0.00%")
inp("beta_y", "原始 β：Yahoo Finance「Beta (5Y Monthly)」", "x", None, V("mkt.beta.y"), None, "Interested-party", "https://finance.yahoo.com/quote/PLTR/key-statistics（頁面日期未載）", "資料商計算值", "資料商計算值，無交易誘因；基準指數與日期未載", fmt="0.00")
inp("beta_s", "原始 β：StockAnalysis「Beta (5Y)」", "x", None, V("mkt.beta.s"), None, "Interested-party", "https://stockanalysis.com/stocks/pltr/statistics/（更新 2026-11-14）", "資料商計算值", "資料商計算值，無交易誘因", fmt="0.00")
inp("blume", "Blume 調整係數", "x", None, 0.67, None, "Derived", "Blume（1971）；同 Meta、Alphabet 模型", fmt="0.00")
INS.append(("beta",))
inp("erp", "股權風險溢酬", "%", 0.04, 0.05, 0.06, "Assumed", "同 Meta、Alphabet 模型（ERP 5%，敏感度 4–6%）", fmt=pct)
inp("kd_sp", "債務信用利差（公司無評等、無借款）", "%", 0.01, 0.02, 0.03, "Assumed", "判斷：Palantir 無公開債務與評等；權重為 0，只影響顯示", "債務權重＝0（無借款），不影響 WACC", fmt=pct)
inp("wacc_adj", "WACC 調整（直接加在 WACC；敏感度用）", "%", -0.01, 0, 0.01, "Assumed", "DCF 標準規格第 4 節：WACC ±1pt", fmt="0.0%")
inp("g_tv", "永續成長率（FY2032 後）", "%", 0.02, 0.03, 0.04, "Assumed", "DCF 標準規格第 3 節預設 3%（敏感度 2–4%）", fmt=pct)
inp("ronic_sp", "永續期 RONIC 相對 WACC 的差", "%", None, 0, 0.05, "Assumed", "DCF 標準規格：RONIC＝WACC（新投資價值中性）；Sens：WACC＋5pt", fmt=pct)
inp("tv_alt", "終值方法（0＝NOPAT×(1−g÷RONIC)；1＝FCFF×(1＋g)）", "旗標", None, 0, 1, "Assumed", "DCF 標準規格第 4 節要求的替代法 Sens", fmt="0")
inp("tax_L", "長期有效稅率（FY2032 與終值；全額備抵釋放後）", "%", 0.15, 0.21, 0.25, "Assumed", "判斷：美國聯邦 21%；目前美英遞延所得稅資產仍全額備抵、有效稅率約 1.3%（10-Q Note 10），備抵釋放時點未揭露", "對 DCF 終值影響大；Sens 15%／25%", fmt=pct)
inp("tax_0", "FY2026 起點有效稅率（前三季實績）", "%", None, V("tax.rate.9M26"), None, "Derived", "10-Q：稅費 41.142 ÷ 稅前 3,117.947；FY2027–FY2032 線性升到 tax_L", fmt="0.00%")
inp("yld", "現金與公債收益率（利息收入 ÷ 期初現金）", "%", 0.025, 0.035, 0.045, "Assumed", "判斷：2026Q3 利息收入 88.9 × 4 ÷ 約 10.5B 現金 ≈ 3.4%", "只影響業外收入與現金累積", fmt=pct)
sec("C. FY2026 Q4 橋接（9M26 實績＋Q4E；三情境共用）")
inp("q4_usc", "2026Q4 美國商業營收季增", "%", 0.12, 0.20, 0.28, "Assumed", "判斷：近三季季增 18%／28%／29%（公司揭露）逐季放緩；公司 Q4 指引（Interested-party）隱含更高，Recon 並列", "不以公司指引校準（原則第 59 條）", fmt=pct)
inp("q4_ic", "2026Q4 國際商業營收季增", "%", -0.02, 0.03, 0.08, "Assumed", "判斷：Q3 國際商業營收 184 百萬、季增 1%（Futurum 轉述公司揭露）", fmt=pct)
inp("q4_gov", "2026Q4 政府營收季增", "%", 0.05, 0.10, 0.15, "Assumed", "判斷：近三季政府季增 18%／15%／16%；放緩假設", fmt=pct)
inp("q4_cust", "2026Q4 美國商業客戶數季增", "%", 0.02, 0.05, 0.07, "Assumed", "判斷：Q3 季增 5%、Q2 6%（Futurum 轉述）", fmt=pct)
sec("D. 命題模組（原型 U 按用量計價）：美國商業營收＝期末客戶數 × 每客戶年營收；每客戶年營收成長＝代理工作負載乘數成長 ＋ 單價侵蝕")
sec_start_fam = len(INS)
def fam(key, label, unit, vals, tag, src, note):
    for s in ("A", "B", "C"):
        inp(f"{key}_{s}", f"{label}（情境 {s} 值）", unit, None, vals[s], None, tag, src, note, fmt=pct)
fam("cg27", "FY2027 美國商業客戶數年增", "%", {"A": 0.24, "B": 0.20, "C": 0.13}, "Assumed", "判斷：Q3 客戶數年增 29%、季增 5%（Futurum 轉述）逐季放緩；情境間差異＝客戶擴張速度（A 維持較高、C 較快飽和）", "情境 A／C 值即區間（In 頁 D／F 欄留空）；FY2027 起向 cg_L 線性收斂到 FY2032")
fam("wl27", "FY2027 每客戶代理工作負載乘數成長", "%", {"A": 0.42, "B": 0.32, "C": 0.20}, "Assumed", "判斷：公司揭露 Q3 每位美國商業客戶近十二個月營收 +88%（Futurum）、NDR 165%（Interested-party）；本欄為其中『用量』部分，FY2027 假設大幅放緩；AIP 任務數沒有揭露", "AI 增量營收來源；工作負載同時推高推論成本（tasks 隨本欄成長）")
fam("pe27", "FY2027 每單位工作負載單價變動（侵蝕為負）", "%", {"A": -0.05, "B": -0.08, "C": -0.13}, "Assumed", "判斷：Q3 說明 Sovereign AI 以開源權重模型自建（L3Harris 稱成本降 95%，Interested-party）、數筆八位數合約年限較短；單價侵蝕公司未量化", "情境 C 單價侵蝕較大")
fam("gg27", "FY2027 政府營收年增", "%", {"A": 0.42, "B": 0.32, "C": 0.20}, "Assumed", "判斷：政府營收 Q3 年增 81%、Q3 TCV 隱含年減約 6%（Futurum 推算）；Maven 用戶超過 10 萬（公司揭露）", "政府分部無用量量價拆分，用有依據的成長路徑；FY2027 起向 gg_L 收斂")
inp("ic27", "FY2027 國際商業營收年增", "%", 0.10, 0.18, 0.25, "Assumed", "判斷：Q3 國際商業 +21% 年增（Futurum 轉述）", "三情境共用", fmt=pct)
inp("cg_L", "FY2032 美國商業客戶數年增（收斂終點）", "%", 0.03, 0.05, 0.08, "Assumed", "判斷：規模化後客戶擴張趨緩", fmt=pct)
inp("wl_L", "FY2032 每客戶工作負載乘數成長（收斂終點）", "%", 0.04, 0.08, 0.12, "Assumed", "判斷", fmt=pct)
inp("pe_L", "FY2032 單價變動（收斂終點）", "%", -0.04, -0.02, 0.0, "Assumed", "判斷：價格侵蝕隨產品組合穩定而趨緩", fmt=pct)
inp("gg_L", "FY2032 政府營收年增（收斂終點）", "%", 0.04, 0.08, 0.12, "Assumed", "判斷：長期接近美國國防與民用 IT 預算成長加滲透", fmt=pct)
inp("ic_L", "FY2032 國際商業營收年增（收斂終點）", "%", 0.04, 0.08, 0.12, "Assumed", "判斷", fmt=pct)
inp("gov_shock", "尾端事件：FY2028 政府營收一次性減少比例（基準 0）", "%", -0.30, 0, None, "Assumed", "原則第 10 節第 19 條：單一客戶群占營收一成以上的事件型尾端風險進 Sens；美國政府約占營收 41%（2026Q3：953÷2,313）", "預算或政策重大變動情境；不進基準", fmt=pct)
sec("E. 推論成本（只讀 Tokenomics 快照 v22.0；不自行估 token 價格）")
inp("tk_p", "軟體公司向模型商購買的市場價（每百萬 token，混合）", "US$/M", TKI["IF_PriceRef_Luna"]["values"], TKI["IF_PriceRef_Terra"]["values"], TKI["IF_PriceRef_Sol"]["values"], "Derived", f"Tokenomics {TKV} IF_PriceRef_Terra（Interface!C409）；低＝IF_PriceRef_Luna（C90）、高＝IF_PriceRef_Sol（C108）；快照 data/tokenomics_snapshot_v22.0.json", "中層模型市場價；Palantir 使用的模型組合未揭露（AIP 連接第三方 LLM）。全成本 IF_FullCost_Terra（GB200 基準 0.0809）只作模型商成本下限對照，見 Recon", fmt="0.0000")
inp("tk_chat", "一般聊天每次任務總 token（新鮮輸入＋快取輸入＋decode）", "token", None, float(TKI["IF_TaskTokFresh"]["values"][0] + TKI["IF_TaskTokCached"]["values"][0] + TKI["IF_TaskTokDec"]["values"][0]), None, "Derived", f"Tokenomics {TKV} IF_TaskTokFresh＋IF_TaskTokCached＋IF_TaskTokDec 第一欄（一般聊天）：1,500＋1,000＋500", fmt="#,##0")
inp("tk_mult", "代理任務 token ÷ 一般聊天 token（單代理工具迴圈）", "x", TKI["IF_TaskTokMult"]["values"][1], TKI["IF_TaskTokMult"]["values"][2], TKI["IF_TaskTokMult"]["values"][3], "Derived", f"Tokenomics {TKV} IF_TaskTokMult（Interface!C331:G331）：基準＝單代理（工具迴圈）9.2；低＝推理聊天 2.07；高＝多代理研究 36.8", "Palantir 任務型態未揭露，選單代理為中位", fmt="0.00")
inp("tk_n", "FY2026 每位美國商業客戶每年 AIP 任務數（百萬件）", "百萬件", 1, 5, 20, "Assumed", "未找到：公司未揭露 AIP 任務數、token 用量或推論成本；判斷值使推論成本約占 Commercial 營收 2%，Sens 1–20 百萬件", "承重 Assumed 之一；可證偽指標：Palantir 揭露 AIP 用量或 Commercial 分部毛利率變化", fmt="0.0")
inp("tk_pg", "每百萬 token 市場價年變動（FY2027 起）", "%", -0.30, 0, 0.10, "Assumed", "判斷：Tokenomics 快照只有現行有效價、無預測路徑，基準持平（不自行重估）；低＝每年降 30%", "待 AI供應鏈 登錄價格路徑名稱後改讀", fmt=pct)
sec("F. 分部利益、成本與現金流比率（原則第 35 條：以最近實績為起點，全期持平）")
def ratio(i, j): return round(V(i) / V(j), 4)
inp("cm_g", "Government contribution 率（公司定義分部利益 ÷ 分部營收）", "%", 0.66, ratio("ctb.gov.9M26", "seg.gov.9M26"), 0.75, "Derived", "10-Q Note 12：前三季 2,127.775 ÷ 2,993.982（FY2025：65.6%；2026Q3：69.6%）", "分部利益含銷售與直接成本、不含研發、一般管理與股票薪酬等未分攤項", fmt=pct)
inp("cm_c", "Commercial contribution 率", "%", 0.70, ratio("ctb.com.9M26", "seg.com.9M26"), 0.80, "Derived", "10-Q Note 12：前三季 2,195.425 ÷ 2,887.476（FY2025：65.9%；2026Q3：76.8%）", fmt=pct)
inp("ua_r", "未分攤成本率（總 contribution − GAAP 營業利益，占營收）", "%", 0.24, round((4323.2 - V("is.oi.9M26")) / V("rev.9M26"), 4), 0.30, "Derived", "10-Q：前三季 (4,323.200 − 2,750.847) ÷ 5,881.458；含股票薪酬（前三季占營收 13.4%）", "GAAP 營業利益＝總 contribution − 未分攤成本；股票薪酬不加回（DCF 標準規格預設）", fmt=pct)
inp("cogs_r", "營業成本率（GAAP，不含增量推論成本；只影響損益表列拆分，不改營業利益）", "%", None, round(V("is.cogs.9M26") / V("rev.9M26"), 4), None, "Derived", "10-Q：前三季營業成本 895.954 ÷ 營收 5,881.458", "預測期營業成本＝本率×營收＋增量推論成本；營業費用為殘差（HTML 規格 D 節）", fmt=pct)
inp("da_r", "折舊攤銷占營收", "%", 0.003, round(V("is.da.9M26") / V("rev.9M26"), 4), 0.006, "Derived", "10-Q 現金流量表：前三季 22.031 ÷ 5,881.458", fmt="0.00%")
inp("cx_r", "資本支出占營收", "%", 0.004, round(V("is.capex.9M26") / V("rev.9M26"), 4), 0.008, "Derived", "10-Q 現金流量表：前三季 29.969 ÷ 5,881.458（輕資產：無自建算力）", fmt="0.00%")
inp("wc_k", "營運資金與其他調整 ÷ 營收年增量", "%", 0.13, 0.20, 0.25, "Derived", "（營業現金流 − 淨利 − 股票薪酬 − 折舊攤銷）÷ 營收年增量：前三季 564.6 ÷ 2,812.8 ≈ 20.1%；FY2025 13.1%；FY2024 5.8%（10-K 現金流量表）", "DCF 標準規格第 5 節、原則第 23 條：歷史比率法；10-K／10-Q 逐年勾稽見 Recon", fmt=pct)
inp("sbc_r", "股票薪酬占營收（只用於現金流加回；損益表已含）", "%", 0.10, round(V("sbc.rev.9M26"), 4), 0.16, "Derived", "10-Q 現金流量表：前三季 785.690 ÷ 5,881.458（FY2025：15.3%）", "現金流加回非現金股票薪酬；DCF 不加回（NOPAT 已扣）", fmt=pct)
sec("G. 估值規則與同業（Comparable：同業 EV／營收經成長調整）")
inp("w_dcf", "目標價中 DCF 的權重（其餘＝Comparable）", "%", 0.3, 0.5, 0.7, "Assumed", "預設 DCF 50%＋Comparable 50%（個股模型原則第 4 節）", fmt=pct)
for s, w in (("A", 0.25), ("B", 0.5), ("C", 0.25)):
    inp(f"tp_w{s}", f"情境權重 {s}", "%", None, w, None, "Assumed", "預設 A 25%／B 50%／C 25%（DCF 標準規格第 3 節）", fmt=pct)
inp("gth", "成長調整指數 θ（倍數 × (分部成長 ÷ 同業成長)^θ；0＝不調整）", "x", 0, 0.5, 1, "Assumed", "判斷：θ＝0 為未調整同業倍數（缺資料時不調整的邊界，原則第 18、47 條）；θ＝1 為倍數與成長等比例", "承重 Assumed；Sens θ＝0 即「相對倍數係數＝1」案", fmt="0.00")
inp("gcap", "成長比值上限（分部成長 ÷ 同業成長 超過者夾到此值）", "x", 2, 3, 99, "Assumed", "判斷：避免與低成長同業比時倍數失控；原則第 38 條：觸及夾值時 Val 揭露原始值與夾值，Sens 高值 99＝不夾值", fmt="0.0")
cm_peers = [p for p in PEERS if p[2] == "商業"]; gv_peers = [p for p in PEERS if p[2] == "政府"]
for p in cm_peers + gv_peers:
    inp(f"pm_{p[0].lower()}", f"{p[1]}（{p[0]}）EV／營收（TTM；{p[2]}同業）", "x", None, p[3], None, "Verified", f"{p[6]}（頁面更新 {p[5]}）；US GAAP，EV 不含營業租賃", "原則第 7、30 條：各家頁面日期不同（2026-10-26 至 2026-11-15），未同一估值日；口徑為 TTM、非 NTM", fmt="0.00")
for p in cm_peers + gv_peers:
    inp(f"pg_{p[0].lower()}", f"{p[1]}（{p[0]}）賣方 3 年營收成長預測", "%", None, round(p[4] / 100, 4), None, "Interested-party", f"{p[6]}（頁面更新 {p[5]}）", "賣方預測", "賣方預測；投行業務關係", fmt=pct)
sec("H. 市場共識與同業對照（只作對照；不進模型）")
inp("cons_tp", "賣方平均目標價（32 家；2026-08-31，Q3 財報前）", "US$", None, V("cons.tp"), None, "Interested-party", "https://ca.marketscreener.com/quote/stock/PALANTIR-TECHNOLOGIES-INC-113108869/consensus/", "財報前，只作歷史對照", "賣方目標價；投行業務關係", fmt="0.00")
inp("cons_rev26", "共識營收 FY2026（29 家；2026-10-23，Q3 財報前）", "US$M", None, V("cons.rev26"), None, "Interested-party", "https://finance.yahoo.com/quote/PLTR/analysis/", "財報前，只作歷史對照", "賣方預估", fmt="#,##0")
inp("cons_rev27", "共識營收 FY2027（28 家；2026-10-23，Q3 財報前）", "US$M", None, V("cons.rev27"), None, "Interested-party", "同上", "財報前，只作歷史對照", "賣方預估", fmt="#,##0")
inp("cons_eps26", "共識 EPS FY2026（29 家；多為 non-GAAP；財報前）", "US$", None, V("cons.eps26"), None, "Interested-party", "同上", "口徑與本模型 GAAP EPS 不同", "賣方預估", fmt="0.00")
inp("cons_eps27", "共識 EPS FY2027（29 家；財報前）", "US$", None, V("cons.eps27"), None, "Interested-party", "同上", "口徑與本模型 GAAP EPS 不同", "賣方預估", fmt="0.00")
inp("g_q4lo", "公司 2026Q4 營收指引下限", "US$M", None, V("gd.q4.rev.lo"), None, "Interested-party", "Q3 財報新聞稿 Outlook", "只作對照", "公司指引", fmt="#,##0")
inp("g_q4hi", "公司 2026Q4 營收指引上限", "US$M", None, V("gd.q4.rev.hi"), None, "Interested-party", "同上", "只作對照", "公司指引", fmt="#,##0")
inp("g_fylo", "公司 FY2026 營收指引下限", "US$M", None, V("gd.fy.rev.lo"), None, "Interested-party", "同上", "只作對照", "公司指引", fmt="#,##0")
inp("g_fyhi", "公司 FY2026 營收指引上限", "US$M", None, V("gd.fy.rev.hi"), None, "Interested-party", "同上", "只作對照", "公司指引", fmt="#,##0")
inp("g_usc", "公司 FY2026 美國商業營收指引（超過）", "US$M", None, V("gd.fy.usc"), None, "Interested-party", "同上", "只作對照", "公司指引", fmt="#,##0")

# 寫入 In
IN["A1"] = ("In — 輸入（藍字；基準欄 E 進主模型；D／F＝敏感度低／高，只用於 Sens 頁、不進目標價；情境 A／B／C 的驅動見本頁最後一節「情境驅動對照」）。金額 US$M；股數百萬股；每股 US$。"
            "標記：Verified／Interested-party／Analogy／Assumed／Derived（Analogy、Assumed 一律給低／高區間；情境專屬輸入以情境 A／C 值為區間，D／F 欄留空）")
IN["A2"] = (f"事實總帳：data/facts_palantir.csv（本頁與 Q 頁同源）；Tokenomics {TKV} 快照 data/tokenomics_snapshot_v22.0.json。評價日 2026-09-30，股價基準日 2026-11-13，目標時點＝股價基準日＋12 個月。黃底＝判斷題。"
            "本版未做先驗（工作單 G 節：Andy 2026-11-16 23:27 指示先建模型）。")
for j, h in enumerate(["鍵", "項目", "單位", "敏感度低", "基準", "敏感度高", "標記", "來源", "說明", "來源誘因（Interested-party 列；混合來源列寫其他欄位）"], 1):
    c = IN.cell(3, j, h); c.font = HDR; c.fill = HFILL
IR = {}; r = 4; BETA_ROWS = {}
for it in INS:
    if it[0] == "sec":
        put(IN, r, 1, it[1], bold=True); r += 1; continue
    if it[0] == "tpdate":
        IR["tp_date"] = r
        put(IN, r, 1, "tp_date"); put(IN, r, 2, "目標時點＝EDATE(股價基準日, 12)"); put(IN, r, 3, "日期")
        put(IN, r, 5, f"=EDATE(E{IR['px_date']},12)", "yyyy-mm-dd"); put(IN, r, 7, "Derived"); put(IN, r, 8, "原則第 26 條"); put(IN, r, 9, "持有期間＝現價日起 12 個月")
        r += 1; continue
    if it[0] == "beta":
        IR["beta"] = r
        b1 = f"E{IR['blume']}*E{IR['beta_y']}+(1-E{IR['blume']})"; b2 = f"E{IR['blume']}*E{IR['beta_s']}+(1-E{IR['blume']})"
        put(IN, r, 1, "beta"); put(IN, r, 2, "β（Blume 調整後兩個來源的中位數）"); put(IN, r, 3, "x")
        put(IN, r, 4, f"=MIN({b1},{b2})", "0.000"); put(IN, r, 5, f"=MEDIAN({b1},{b2})", "0.000"); put(IN, r, 6, f"=MAX({b1},{b2})", "0.000")
        put(IN, r, 7, "Derived"); put(IN, r, 8, "兩個資料商原始 β 各自 Blume 調整（0.67×β＋0.33）後取中位數"); put(IN, r, 9, "低／高＝兩個來源各自的調整值")
        BETA_ROWS["r"] = r; r += 1; continue
    d = it[1]; IR[d["key"]] = r
    put(IN, r, 1, d["key"]); put(IN, r, 2, d["label"]); put(IN, r, 3, d["unit"])
    base = OVR.get(d["key"], d["base"])
    if d["lo"] is not None: put(IN, r, 4, d["lo"], d["fmt"], blue=True)
    c = put(IN, r, 5, base, d["fmt"], blue=True)
    if d["tag"] == "Assumed": c.fill = YEL
    if d["hi"] is not None: put(IN, r, 6, d["hi"], d["fmt"], blue=True)
    put(IN, r, 7, d["tag"]); put(IN, r, 8, d["src"]); put(IN, r, 9, d["note"] or None); put(IN, r, 10, d["inc"] or None)
    r += 1
IN_LAST = r - 1
def I(k): return f"In!$E${IR[k]}"
def Ilo(k): return f"In!$D${IR[k]}"
def Ihi(k): return f"In!$F${IR[k]}"
# 情境驅動對照
r += 1
put(IN, r, 1, "J. 情境驅動對照（三情境並排，進主模型的加權目標價；本節只引用上方各列，修改請改上方藍字。上方「敏感度低／高」不進目標價）", bold=True); SCN_SEC = r; r += 1
for j, h in enumerate(["鍵", "項目", "單位", "情境 A（樂觀）", "情境 B（基準）", "情境 C（保守）", "標記", "說明"], 1):
    c = IN.cell(r, j, h); c.font = HDR; c.fill = HFILL
SCN_H = r; r += 1
SCN_ROWS = {}
for k, lab, unit in (("scn_cg", "FY2027 美國商業客戶數年增", "%"), ("scn_wl", "FY2027 每客戶工作負載乘數成長", "%"), ("scn_pe", "FY2027 每單位工作負載單價變動", "%"), ("scn_gg", "FY2027 政府營收年增", "%")):
    base = k[4:]
    SCN_ROWS[k] = r
    put(IN, r, 1, k); put(IN, r, 2, lab); put(IN, r, 3, unit)
    for j, s in zip((4, 5, 6), "ABC"): put(IN, r, j, f"=E{IR[f'{base}27_{s}']}", pct)
    put(IN, r, 7, "Derived"); put(IN, r, 8, f"引用 In 頁 {base}27_A／{base}27_B／{base}27_C；情境 B 為基準；分化依據見 README「情境分化依據」")
    r += 1
SCN_ROWS["scn_w"] = r
put(IN, r, 1, "scn_w"); put(IN, r, 2, "情境權重"); put(IN, r, 3, "%")
for j, s in zip((4, 5, 6), "ABC"): put(IN, r, j, f"=E{IR[f'tp_w{s}']}", pct)
put(IN, r, 7, "Derived"); put(IN, r, 8, "引用 In 頁 tp_wA／tp_wB／tp_wC；[判斷] 黃底，Andy 可改"); r += 1
for col, w in zip("ABCDEFGHIJ", [12, 58, 9, 11, 11, 11, 16, 70, 60, 40]): IN.column_dimensions[col].width = w
