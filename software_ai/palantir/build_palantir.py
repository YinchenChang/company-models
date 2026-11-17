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
from facts_palantir import FACTS, V, PEERS, FACT_BY_ID, PSL, PS13  # noqa: E402

ap = argparse.ArgumentParser()
ap.add_argument("out"); ap.add_argument("--sens"); ap.add_argument("--set", action="append", default=[])
A = ap.parse_args()
OVR = {k: float(v) for k, v in (s.split("=", 1) for s in A.set)}
SENS = json.load(open(A.sens, encoding="utf-8")) if A.sens and os.path.exists(A.sens) else {}

VER = "v0.2.2"; VDATE = "2026-11-17"
def _build_id():
    h = hashlib.sha256()
    for fn in ["build_palantir.py", "data/facts_palantir.py", "../data/relmult_hist/out/peer_ps_20261113.json", "../data/relmult_hist/peer_ps_20261113.csv", "data/tokenomics_snapshot_v22.0.json", "../data/relmult_hist/out/palantir.json"] + [f"../data/relmult_hist/{_t}.csv" for _t in PSL]:
        h.update(fn.encode()); h.update(open(os.path.join(HERE, fn), "rb").read())
    return f"palantir-{VER}-{h.hexdigest()[:10]}"
BUILD_ID = _build_id()
TK = json.load(open(os.path.join(HERE, "data", "tokenomics_snapshot_v22.0.json"), encoding="utf-8"))
RELM = json.load(open(os.path.join(HERE, "..", "data", "relmult_hist", "out", "palantir.json"), encoding="utf-8"))["groups"]
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
    put(Q, r, 11 + k, f"={q(f'seg.gov.{y}')}+{q(f'seg.com.{y}')}-{q(f'rev.{y}')}", "0.000")
r += 1; put(Q, r, 1, "檢查：Government＋Commercial－合併營收（最近四季；應為 0）", bold=True); QCHK2 = r
for k, qq in enumerate(["2025Q4", "2026Q1", "2026Q2", "2026Q3"]):
    put(Q, r, 11 + k, f"={q(f'seg.gov.{qq}')}+{q(f'seg.com.{qq}')}-{q(f'rev.{qq}')}", "0.000")
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
inp("data_date", "資料截止日（本版最晚一筆輸入：共用庫相對倍數 JSON 產生日）", "日期", None, datetime.datetime(2026, 11, 17), None, "Derived", "software_ai/data/relmult_hist/out/palantir.json generated 2026-11-17；同業倍數最晚頁面日 2026-11-15（LDOS）；財報原文截至 2026-11-03 的 10-Q", "原則第 41 條：README 首列分列資料截止日、成品日、股價基準日", fmt="yyyy-mm-dd")
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
inp("rf", "無風險利率（10 年美債，2026-11-12）", "%", None, V("mkt.rf"), None, "Verified", "FRED DGS10 2026-11-12 收盤 5.10%（https://fred.stlouisfed.org/series/DGS10；工作單 G5 第 2 條統一美國 10 年期）", "現價日 2026-11-13 前 1 個交易日，未逾期（原則第 15 條）", fmt="0.00%")
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
inp("tk_p", "軟體公司向模型商購買的市場價（每百萬 token，混合）", "US$/M", TKI["IF_PriceRef_Luna"]["values"], TKI["IF_PriceRef_Terra"]["values"], TKI["IF_PriceRef_Sol"]["values"], "Analogy", f"Tokenomics {TKV} IF_PriceRef_Terra（Interface!C409）；低＝IF_PriceRef_Luna（C90）、高＝IF_PriceRef_Sol（C108）；快照 data/tokenomics_snapshot_v22.0.json", "中層模型市場價；Palantir 使用的模型組合未揭露（AIP 連接第三方 LLM）。全成本 IF_FullCost_Terra（GB200 基準 0.0809）只作模型商成本下限對照，見 Recon", fmt="0.0000")
inp("tk_chat", "一般聊天每次任務總 token（新鮮輸入＋快取輸入＋decode）", "token", None, float(TKI["IF_TaskTokFresh"]["values"][0] + TKI["IF_TaskTokCached"]["values"][0] + TKI["IF_TaskTokDec"]["values"][0]), None, "Derived", f"Tokenomics {TKV} IF_TaskTokFresh＋IF_TaskTokCached＋IF_TaskTokDec 第一欄（一般聊天）：1,500＋1,000＋500", fmt="#,##0")
inp("tk_mult", "代理任務 token ÷ 一般聊天 token（單代理工具迴圈）", "x", TKI["IF_TaskTokMult"]["values"][1], TKI["IF_TaskTokMult"]["values"][2], TKI["IF_TaskTokMult"]["values"][3], "Analogy", f"Tokenomics {TKV} IF_TaskTokMult（Interface!C331:G331）：基準＝單代理（工具迴圈）9.2；低＝推理聊天 2.07；高＝多代理研究 36.8", "公司未揭露 token 量或代理任務量（未找到，G2-1 無從校準），改以快照單代理為類比（Analogy）；一旦揭露即以揭露推得的每任務 token 取代", fmt="0.00")
inp("tk_n", "FY2026 每位美國商業客戶每年 AIP 任務數（百萬件）", "百萬件", 1, 5, 20, "Assumed", "未找到：公司未揭露 AIP 任務數、token 用量或推論成本；判斷值使推論成本約占 Commercial 營收 2%，Sens 1–20 百萬件", "承重 Assumed 之一；可證偽指標：Palantir 揭露 AIP 用量或 Commercial 分部毛利率變化", fmt="0.0")
inp("tk_pg", "每百萬 token 市場價年變動（FY2027 起）", "%", -0.30, 0, 0.10, "Assumed", "判斷：Tokenomics 快照只有現行有效價、無預測路徑，基準持平（不自行重估）；低＝每年降 30%", "待 AI供應鏈 登錄價格路徑名稱後改讀", fmt=pct)
sec("F. 分部利益、成本與現金流比率（原則第 35 條：以最近實績為起點，全期持平）")
def ratio(i, j): return round(V(i) / V(j), 4)
inp("cm_g", "Government contribution 率（公司定義分部利益 ÷ 分部營收）", "%", 0.66, ratio("ctb.gov.9M26", "seg.gov.9M26"), 0.75, "Derived", "10-Q Note 12：前三季 2,127.775 ÷ 2,993.982（FY2025：65.6%；2026Q3：69.6%）", "分部利益含銷售與直接成本、不含研發、一般管理與股票薪酬等未分攤項", fmt=pct)
inp("cm_c", "Commercial contribution 率", "%", 0.70, ratio("ctb.com.9M26", "seg.com.9M26"), 0.80, "Derived", "10-Q Note 12：前三季 2,195.425 ÷ 2,887.476（FY2025：65.9%；2026Q3：76.8%）", fmt=pct)
inp("ua_r", "未分攤成本率（總 contribution − GAAP 營業利益，占營收）", "%", 0.24, round((4323.2 - V("is.oi.9M26")) / V("rev.9M26"), 4), 0.30, "Derived", "10-Q：前三季 (4,323.200 − 2,750.847) ÷ 5,881.458；含股票薪酬（前三季占營收 13.4%）", "GAAP 營業利益＝總 contribution − 未分攤成本；股票薪酬不加回（DCF 標準規格預設）", fmt=pct)
inp("ua_d", "FY2032 未分攤成本率相對前三季水準的變動（規模經濟延伸；基準＝0＝全期持平）", "%", -0.07, 0, None, "Assumed", "G3-1：未分攤成本率歷史有規模經濟（FY2023 約 50.6%、FY2024 49.2%、FY2025 34.2%、前三季 26.7%，Q 頁推得），但前三季水準已是歷史最低、其中股票薪酬占 13.4% 屬半固定；沒有固定／變動成本資料可回歸，基準不延伸（持平於前三季實績），延伸案放 Sens 低值（FY2032 降約 7 個百分點，線性）", "全期持平為基準；低值＝規模經濟延伸；承重 Assumed；可證偽指標：未分攤成本率逐季變化", fmt=pct)
inp("cogs_r", "營業成本率（GAAP，不含增量推論成本；只影響損益表列拆分，不改營業利益）", "%", None, round(V("is.cogs.9M26") / V("rev.9M26"), 4), None, "Derived", "10-Q：前三季營業成本 895.954 ÷ 營收 5,881.458", "預測期營業成本＝本率×營收＋增量推論成本；營業費用為殘差（HTML 規格 D 節）", fmt=pct)
inp("da_r", "折舊攤銷占營收", "%", 0.003, round(V("is.da.9M26") / V("rev.9M26"), 4), 0.006, "Derived", "10-Q 現金流量表：前三季 22.031 ÷ 5,881.458", fmt="0.00%")
inp("cx_r", "資本支出占營收", "%", 0.004, round(V("is.capex.9M26") / V("rev.9M26"), 4), 0.008, "Derived", "10-Q 現金流量表：前三季 29.969 ÷ 5,881.458（輕資產：無自建算力）", fmt="0.00%")
inp("wc_k", "營運資金與其他調整 ÷ 營收年增量（負值＝現金流出）", "%", -0.25, -0.20, -0.13, "Derived", "（營業現金流 − 淨利 − 股票薪酬 − 折舊攤銷）÷ 營收年增量：前三季 −564.6 ÷ 2,812.8 ≈ −20.1%；FY2025 −13.1%；FY2024 −5.8%（10-K 現金流量表）", "DCF 標準規格第 5 節、原則第 23 條：歷史比率法；10-K／10-Q 逐年勾稽見 Recon", fmt=pct)
inp("sbc_r", "股票薪酬占營收（只用於現金流加回；損益表已含）", "%", 0.10, round(V("sbc.rev.9M26"), 4), 0.16, "Derived", "10-Q 現金流量表：前三季 785.690 ÷ 5,881.458（FY2025：15.3%）", "現金流加回非現金股票薪酬；DCF 不加回（NOPAT 已扣）", fmt=pct)
sec("G. 估值規則與同業（Comparable 主腿：同業 P/S × 共用庫相對倍數；G6 與 PS／PSG 同口徑）")
inp("w_dcf", "目標價中 DCF 的權重（其餘＝Comparable）", "%", 0.3, 0.5, 0.7, "Assumed", "預設 DCF 50%＋Comparable 50%（個股模型原則第 4 節）", fmt=pct)
for s, w in (("A", 0.25), ("B", 0.5), ("C", 0.25)):
    inp(f"tp_w{s}", f"情境權重 {s}", "%", None, w, None, "Assumed", "預設 A 25%／B 50%／C 25%（DCF 標準規格第 3 節）", fmt=pct)
inp("gth", "成長調整指數 θ（倍數 × (分部成長 ÷ 同業成長)^θ；v0.2 基準 0＝改由共用庫相對倍數取代，θ 只留作 Sens）", "x", None, 0, 0.5, "Assumed", "判斷：θ＝0 為未調整同業倍數（缺資料時不調整的邊界，原則第 18、47 條）；θ＝1 為倍數與成長等比例", "v0.2：基準 0（相對倍數改讀共用庫）；Sens 高值 0.5＝v0.1 基準的成長調整", fmt="0.00")
_rg = RELM["商業"]; _rv = RELM["政府"]; _ra = RELM["all"]
for _k, _nm in (("rel_c", "Commercial"), ("rel_g", "Government")):
    inp(_k, f"{_nm}：整體相對倍數（本公司合併 P/S ÷ 8 家同業合併中位 P/S；共用庫 all.geo_base，兩分部同值）", "x", round(_ra["sens"]["low"], 4), round(_ra["geo_base"], 4), round(_ra["sens"]["high"], 4), "Derived",
        f"software_ai/data/relmult_hist/out/palantir.json groups.all.geo_base；base_source＝{_ra['base_source']}；低／高＝sens.low／sens.high；v0.2.1：整體單一口徑，分組讀數（商業、政府）只列對照",
        "Comparable 倍數＝8 家同業合併中位 P/S × 本係數；兩分部同值，Sens 同步改動；PS 與 PSG 兩讀數差異大（見 README）", fmt="0.00")
for _k, _g, _nm in (("c", _rg, "商業"), ("g", _rv, "政府"), ("a", _ra, "整體")):
    for _kk, _lab, _v, _f in ((f"ps_{_k}", "PS 讀數（20 季歷史相對值中位）", _g["ps_reading"], "0.00"), (f"psg_{_k}", "PSG 讀數（相對值）", _g["psg_reading"], "0.00"),
                              (f"gr_{_k}", "成長比（本公司 TTM 營收年增 ÷ 同業中位）", _g["growth_ratio"]["value"], "0.00"), (f"cgr_{_k}", "本公司 TTM 營收年增", _g["growth_ratio"]["company_g"], "0.0%"),
                              (f"pgr_{_k}", "同業 TTM 營收年增中位", _g["growth_ratio"]["peer_median_g"], "0.0%"), (f"dv_{_k}", "PS 與 PSG 讀數差異倍數（divergence）", _g["divergence"], "0.00")):
        inp(_kk, f"{_nm}：{_lab}", "x", None, round(_v, 4), None, "Derived", f"relmult_hist/out/palantir.json groups.{_nm}", "只作對照，不進模型（v0.2.1 起分組讀數不進加權）", fmt=_f)
inp("rel_c_ref", "對照：商業分組 geo_base（不進模型）", "x", None, round(_rg["geo_base"], 4), None, "Derived", "relmult_hist/out/palantir.json groups.商業", "只作對照，不進加權", fmt="0.00")
inp("rel_g_ref", "對照：政府分組 geo_base（不進模型）", "x", None, round(_rv["geo_base"], 4), None, "Derived", "relmult_hist/out/palantir.json groups.政府", "只作對照，不進加權", fmt="0.00")
inp("gcap", "成長比值上限（分部成長 ÷ 同業成長 超過者夾到此值）", "x", 2, 3, 99, "Assumed", "判斷：避免與低成長同業比時倍數失控；原則第 38 條：觸及夾值時 Val 揭露原始值與夾值，Sens 高值 99＝不夾值", fmt="0.0")
cm_peers = [p for p in PEERS if p[2] == "商業"]; gv_peers = [p for p in PEERS if p[2] == "政府"]
for p in cm_peers + gv_peers:
    inp(f"pm_{p[0].lower()}", f"{p[1]}（{p[0]}）P/S（市值÷TTM 營收；{p[2]}同業；peer_ps_20261113.json，價格日 {PS13[p[0]][0]}）", "x", None, V(f"peer.{p[0]}.ps13"), None, "Derived", f"software_ai/data/relmult_hist/out/peer_ps_20261113.json groups.all.peers.{p[0]}（價格日 {PS13[p[0]][0]}，基準日 2026-11-13）", "G7：同業 P/S 取現價日（2026-11-13）口徑；價格日非 11-13 者於 Recon 第 5 節標逾期；資料庫最近一季值列 Sens 案", fmt="0.00")
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
        put(IN, r, 4, f"=MIN({b1},{b2})", "0.000"); put(IN, r, 5, OVR["beta"] if "beta" in OVR else f"=MEDIAN({b1},{b2})", "0.000"); put(IN, r, 6, f"=MAX({b1},{b2})", "0.000")
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

# =====================================================================================================
# 版面：先配置各頁列號（公式互相引用）
def alloc(start, keys):
    return {k: start + i for i, k in enumerate(keys)}
YL = {y: f"FY{y}{'A' if y <= 2025 else 'E'}" for y in YRS}
# ---- Seg
SB = dict(hdr=6, usc=7, ic=8, com=9, gov=10, rev=11, cust=12, arpc=13)
SCM = dict(hdr=15, price=16, tokens=17, icg=18)
SEGK = ["hd", "cg", "wl", "pe", "gg", "cust", "arpc", "usc", "ic", "com", "gov", "rev", "yoy", "inc", "ero", "tasks", "cost", "share", "incinf", "net", "be", "chk"]
SEGR = {}; r0 = 21
for s in SC:
    SEGR[s] = alloc(r0, SEGK); r0 += len(SEGK) + 1
# ---- PL
PLQ = alloc(6, ["hd", "gov", "com", "rev", "cg", "cc", "ctb", "ua", "oi", "cogs", "oth", "pbt", "tax", "ni"])
PLK = ["hd", "gov", "com", "rev", "ctb_g", "ctb_c", "ctb", "ua", "oi", "cogs", "gp", "opx", "oth", "pbt", "taxr", "tax", "ni", "sh", "eps", "chk_seg", "chk_pbt", "chk_ni"]
PLR = {}; r0 = 22
for s in SC:
    PLR[s] = alloc(r0, PLK); r0 += len(PLK) + 1
PL_STD0 = r0
STDK = ["hd", "rev", "yoy", "cogs", "gp", "gm", "opx", "oi", "opm", "oth", "pbt", "tax", "ni", "nm", "eps", "sub", "r_yoy", "r_gm", "r_opm", "r_nm", "r_niy", "chk1", "chk2"]
PLS = alloc(PL_STD0, STDK)
# ---- CF
CFQ = alloc(6, ["hd", "sbc", "da", "wc", "cfo", "cx", "fcf"])
CFK = ["hd", "ni", "sbc", "da", "wc", "cfo", "cx", "fcf", "div", "fcfd", "cash", "oth", "chk"]
CFR = {}; r0 = 16
for s in SC:
    CFR[s] = alloc(r0, CFK); r0 += len(CFK) + 1
CF_GATE = r0; CF_GATE_ROWS = alloc(r0, ["gate", "gate_ok", "liq_max", "liq_cash"])
# ---- DCF
DK = ["hd", "oi", "taxr", "nopat", "da", "cx", "wc", "fcff", "t", "df", "pv", "sumpv", "tvn", "tv", "tvt", "pvtv", "ev", "nc", "eq", "ps", "pstp", "tvsh", "sbc", "pvsb", "pvsbtv", "evsb", "pssb", "pstpsb"]
DR = {}; r0 = 5
for s in SC:
    DR[s] = alloc(r0, DK); r0 += len(DK) + 1
# ---- Val
VL = ["cap1", "px", "pxd", "tpd", "tpdcf", "tpcomp", "tp", "cons", "up", "updcf", "upcomp", "tpdcf_sbc", "updcf_sbc", "tier", "est", "gate", "rec", "bdist", "att_dcf", "att_cmp", "att_adj",
      "cap2", "sh_basic", "sh_rsu", "sh_opt", "sh_sar", "sh_dil", "nc0",
      "cap3", "rf", "beta", "erp", "ke", "kd", "kdat", "mcap", "debt", "we", "wacc", "ronic", "g", "tau", "tau_h",
      "cap4", "d_pv", "d_pvtv", "d_ev", "d_nc", "d_eq", "d_ps", "d_pstp", "d_tvsh", "d_pstp_sbc",
      "cap5", "m_c", "m_g", "rl_c", "rl_g", "pg_c", "pg_g", "sg_c", "sg_g", "rr_c", "rc_c", "rr_g", "rc_g", "ma_c", "ma_g", "ftp", "ttm_c", "ttm_g", "ev_c", "ev_g", "ev_t", "nc_t", "eq_t", "ps_c", "ps_u", "imp_m", "clamp", "meth"]
VR = alloc(4, VL)
SCOL = {"A": "B", "B": "C", "C": "D"}
def val(k, col="B"): return f"Val!${col}${VR[k]}"

def fade(k27, kL, c):
    return f"={I(kL)}+({I(k27)}-{I(kL)})*(2032-{c}$4)/5"

# =====================================================================================================
# Seg
S = SH["Seg"]
S["A1"] = "Seg — 命題模組（原型 U 按用量計價）與官方分部營收：美國商業營收＝期末美國商業客戶數 × 每客戶年營收；每客戶年營收成長＝代理工作負載乘數成長＋單價侵蝕（AI 增量與侵蝕）；推論成本只讀 Tokenomics 快照 v22.0。金額 US$M"
S["A2"] = ("官方分部＝Government 與 Commercial（10-K 分部）。Commercial 之下子列＝美國商業與國際商業（新聞稿固定揭露，合計＝Commercial）。FY2026E＝前三季實績＋Q4E（S0 橋接，三情境共用）；FY2027 起三情境分化（客戶數、工作負載、單價、政府成長）。"
           "AI 淨效果＝AI 增量（客戶數 × 前一年每客戶營收 × 工作負載成長）＋單價侵蝕（同底數 × 單價變動，為負）－增量推論成本（相對 FY2026 隱含占比）。變數名稱與 software_ai/unit_economics 原型 U 對應；該模型完成前於本模型 In 頁自建（v0.2 改讀 unit_economics）。")
def year_header(ws, with_m=None):
    put(ws, 3, 1, "項目", bold=True); put(ws, 3, 2, "單位", bold=True)
    for y in YRS: put(ws, 3, 3 + y - 2023, YL[y], bold=True)
    put(ws, 4, 1, "年（數值，公式用）")
    for y in YRS: put(ws, 4, 3 + y - 2023, y)
    if with_m: put(ws, 3, 13, with_m, bold=True)
year_header(S)
for j, h in enumerate(["S0. FY2026E 橋接（US$M）", "2026Q3A", "Q4 季增（In）", "2026Q4E", "2026 前三季 A", "FY2026E", "2025Q4A", "Q4 年增"], 1):
    c = S.cell(SB["hdr"], j, h); c.font = HDR; c.fill = HFILL
R_ = SB
put(S, R_["usc"], 1, "美國商業營收（公司自揭 KPI，四捨五入到百萬）")
put(S, R_["usc"], 2, f"={q('us.com.2026Q3')}", "#,##0.0"); put(S, R_["usc"], 3, f"={I('q4_usc')}", pct); put(S, R_["usc"], 4, f"=B{R_['usc']}*(1+C{R_['usc']})", "#,##0.0")
put(S, R_["usc"], 5, f"={q('us.com.2026Q1')}+{q('us.com.2026Q2')}+{q('us.com.2026Q3')}", "#,##0.0"); put(S, R_["usc"], 6, f"=D{R_['usc']}+E{R_['usc']}", "#,##0.0")
put(S, R_["usc"], 7, f"={q('us.com.2025Q4')}", "#,##0.0"); put(S, R_["usc"], 8, f"=D{R_['usc']}/G{R_['usc']}-1", pct)
put(S, R_["ic"], 1, "國際商業營收（＝Commercial − 美國商業；Derived）")
put(S, R_["ic"], 2, f"={q('seg.com.2026Q3')}-B{R_['usc']}", "#,##0.0"); put(S, R_["ic"], 3, f"={I('q4_ic')}", pct); put(S, R_["ic"], 4, f"=B{R_['ic']}*(1+C{R_['ic']})", "#,##0.0")
put(S, R_["ic"], 5, f"={q('seg.com.9M26')}-E{R_['usc']}", "#,##0.0"); put(S, R_["ic"], 6, f"=D{R_['ic']}+E{R_['ic']}", "#,##0.0")
put(S, R_["ic"], 7, f"={q('seg.com.2025Q4')}-G{R_['usc']}", "#,##0.0"); put(S, R_["ic"], 8, f"=D{R_['ic']}/G{R_['ic']}-1", pct)
put(S, R_["com"], 1, "Commercial（官方分部）")
for j, cc in ((2, "B"), (4, "D"), (6, "F")): put(S, R_["com"], j, f"={cc}{R_['usc']}+{cc}{R_['ic']}", "#,##0.0")
put(S, R_["com"], 5, f"={q('seg.com.9M26')}", "#,##0.0"); put(S, R_["com"], 7, f"={q('seg.com.2025Q4')}", "#,##0.0"); put(S, R_["com"], 8, f"=D{R_['com']}/G{R_['com']}-1", pct)
put(S, R_["gov"], 1, "Government（官方分部）")
put(S, R_["gov"], 2, f"={q('seg.gov.2026Q3')}", "#,##0.0"); put(S, R_["gov"], 3, f"={I('q4_gov')}", pct); put(S, R_["gov"], 4, f"=B{R_['gov']}*(1+C{R_['gov']})", "#,##0.0")
put(S, R_["gov"], 5, f"={q('seg.gov.9M26')}", "#,##0.0"); put(S, R_["gov"], 6, f"=D{R_['gov']}+E{R_['gov']}", "#,##0.0"); put(S, R_["gov"], 7, f"={q('seg.gov.2025Q4')}", "#,##0.0"); put(S, R_["gov"], 8, f"=D{R_['gov']}/G{R_['gov']}-1", pct)
put(S, R_["rev"], 1, "合併營收", bold=True)
for j, cc in ((2, "B"), (4, "D"), (5, "E"), (6, "F"), (7, "G")): put(S, R_["rev"], j, f"={cc}{R_['com']}+{cc}{R_['gov']}", "#,##0.0")
put(S, R_["rev"], 8, f"=D{R_['rev']}/G{R_['rev']}-1", pct)
put(S, R_["cust"], 1, "美國商業客戶數（期末；公司定義）"); put(S, R_["cust"], 2, f"={q('kpi.usc.q3')}", "#,##0"); put(S, R_["cust"], 3, f"={I('q4_cust')}", pct)
put(S, R_["cust"], 4, f"=B{R_['cust']}*(1+C{R_['cust']})", "#,##0"); put(S, R_["cust"], 6, f"=D{R_['cust']}", "#,##0")
put(S, R_["arpc"], 1, "每客戶年營收（US$M；＝FY2026E 美國商業營收 ÷ 期末客戶數；Derived）"); put(S, R_["arpc"], 6, f"=F{R_['usc']}/F{R_['cust']}", "0.000")
c = S.cell(SCM["hdr"], 1, "S1. 命題模組共用（Tokenomics 快照 v22.0；不自行估 token 價格）"); c.font = HDR; c.fill = HFILL
put(S, SCM["price"], 1, "每百萬 token 市場價（US$/M；IF_PriceRef_Terra 路徑）"); put(S, SCM["price"], 2, "US$/M")
put(S, SCM["tokens"], 1, "每任務 token 數（一般聊天總 token × IF_TaskTokMult）"); put(S, SCM["tokens"], 2, "token")
put(S, SCM["icg"], 1, "國際商業營收年增路徑（三情境共用）"); put(S, SCM["icg"], 2, "%")
for y in FC:
    c = yc(y); p = yc(y - 1)
    put(S, SCM["price"], 3 + y - 2023, f"={I('tk_p')}" if y == 2026 else f"={p}{SCM['price']}*(1+{I('tk_pg')})", "0.0000")
    put(S, SCM["tokens"], 3 + y - 2023, f"={I('tk_chat')}*{I('tk_mult')}", "#,##0")
    if y >= 2027: put(S, SCM["icg"], 3 + y - 2023, fade("ic27", "ic_L", c), pct)
LABSEG = {"cg": ("美國商業客戶數年增", "%"), "wl": ("每客戶代理工作負載乘數成長", "%"), "pe": ("每單位工作負載單價變動（侵蝕為負）", "%"), "gg": ("政府營收年增", "%"),
       "cust": ("美國商業客戶數（期末）", "家"), "arpc": ("每客戶年營收", "US$M"), "usc": ("美國商業營收＝客戶數 × 每客戶年營收", "US$M"), "ic": ("國際商業營收", "US$M"),
       "com": ("Commercial（美國商業＋國際商業）", "US$M"), "gov": ("Government", "US$M"), "rev": ("合併營收", "US$M"), "yoy": ("合併營收年增率", "%"),
       "inc": ("AI 增量營收（客戶數 × 前一年每客戶營收 × 工作負載成長）", "US$M"), "ero": ("單價侵蝕（同底數 × 單價變動）", "US$M"),
       "tasks": ("每客戶每年 AIP 任務數（隨工作負載成長）", "百萬件"), "cost": ("推論成本（客戶數 × 任務數 × 每任務 token × 市場價）", "US$M"), "share": ("推論成本占 Commercial 營收", "%"),
       "incinf": ("增量推論成本（相對 FY2026 隱含占比；進損益）", "US$M"), "net": ("AI 淨效果＝增量＋侵蝕－增量推論成本", "US$M"),
       "be": ("損益兩平：淨效果＝0 所需的工作負載成長（反解）", "%"), "chk": ("檢查：美國商業營收－客戶數×每客戶年營收（應為 0）", "US$M")}
for s in SC:
    R = SEGR[s]
    c = S.cell(R["hd"], 1, f"{SCN_NAME[s]}"); c.font = HDR; c.fill = HFILL
    for k, (t, u) in LABSEG.items():
        put(S, R[k], 1, t); put(S, R[k], 2, u)
    for y in YRS:
        col = 3 + y - 2023; c = yc(y); p = yc(y - 1)
        if y <= 2025:
            put(S, R["gov"], col, f"={q(f'seg.gov.{y}')}", "#,##0.0"); put(S, R["com"], col, f"={q(f'seg.com.{y}')}", "#,##0.0"); put(S, R["rev"], col, f"={q(f'rev.{y}')}", "#,##0.0")
            if y >= 2024: put(S, R["yoy"], col, f"={c}{R['rev']}/{p}{R['rev']}-1", pct)
            continue
        if y == 2026:
            put(S, R["cust"], col, f"=$F${SB['cust']}", "#,##0"); put(S, R["arpc"], col, f"=$F${SB['arpc']}", "0.000"); put(S, R["usc"], col, f"=$F${SB['usc']}", "#,##0.0")
            put(S, R["ic"], col, f"=$F${SB['ic']}", "#,##0.0"); put(S, R["gov"], col, f"=$F${SB['gov']}", "#,##0.0"); put(S, R["tasks"], col, f"={I('tk_n')}", "0.00")
            put(S, R["incinf"], col, 0, "#,##0.0")
        else:
            put(S, R["cg"], col, fade(f"cg27_{s}", "cg_L", c), pct); put(S, R["wl"], col, fade(f"wl27_{s}", "wl_L", c), pct)
            put(S, R["pe"], col, fade(f"pe27_{s}", "pe_L", c), pct); put(S, R["gg"], col, fade(f"gg27_{s}", "gg_L", c), pct)
            put(S, R["cust"], col, f"={p}{R['cust']}*(1+{c}{R['cg']})", "#,##0"); put(S, R["arpc"], col, f"={p}{R['arpc']}*(1+{c}{R['wl']}+{c}{R['pe']})", "0.000")
            put(S, R["usc"], col, f"={c}{R['cust']}*{c}{R['arpc']}", "#,##0.0"); put(S, R["ic"], col, f"={p}{R['ic']}*(1+{c}${SCM['icg']})", "#,##0.0")
            put(S, R["gov"], col, f"={p}{R['gov']}*(1+{c}{R['gg']})*(1+IF({c}$4=2028,{I('gov_shock')},0))", "#,##0.0")
            put(S, R["inc"], col, f"={c}{R['cust']}*{p}{R['arpc']}*{c}{R['wl']}", "#,##0.0"); put(S, R["ero"], col, f"={c}{R['cust']}*{p}{R['arpc']}*{c}{R['pe']}", "#,##0.0")
            put(S, R["tasks"], col, f"={p}{R['tasks']}*(1+{c}{R['wl']})", "0.00")
            put(S, R["incinf"], col, f"={c}{R['cost']}-$F{R['cost']}/$F{R['com']}*{c}{R['com']}", "#,##0.0")
            put(S, R["net"], col, f"={c}{R['inc']}+{c}{R['ero']}-{c}{R['incinf']}", "#,##0.0")
            put(S, R["be"], col, f"=-{c}{R['pe']}+{c}{R['incinf']}/({c}{R['cust']}*{p}{R['arpc']})", pct)
        put(S, R["yoy"], col, f"={c}{R['rev']}/{p}{R['rev']}-1", pct)
        put(S, R["com"], col, f"={c}{R['usc']}+{c}{R['ic']}", "#,##0.0"); put(S, R["rev"], col, f"={c}{R['com']}+{c}{R['gov']}", "#,##0.0")
        put(S, R["cost"], col, f"={c}{R['cust']}*{c}{R['tasks']}*{c}${SCM['tokens']}*{c}${SCM['price']}/1000000", "#,##0.0")
        put(S, R["share"], col, f"={c}{R['cost']}/{c}{R['com']}", "0.00%")
        put(S, R["chk"], col, f"={c}{R['usc']}-{c}{R['cust']}*{c}{R['arpc']}", "0.000")
S.column_dimensions["A"].width = 70
for col in "BCDEFGHIJKLM": S.column_dimensions[col].width = 12

# =====================================================================================================
# PL
P = SH["PL"]
P["A1"] = "PL — 損益（GAAP；US$M；EPS 美元）：官方分部營收 → 分部 contribution（公司定義分部利益）→ 扣未分攤成本 → GAAP 營業利益 → 業外、稅、淨利、EPS。FY2026E＝前三季實績（M 欄）＋Q4E"
P["A2"] = ("營業利益＝總 contribution − 未分攤成本（含股票薪酬）；contribution 率與未分攤成本率以 2026 前三季實績為起點、全期持平（原則第 35 條）；Commercial contribution 另扣增量推論成本。"
           "預測期營業成本＝營業成本率×營收＋增量推論成本，營業費用為殘差（營業利益不受此拆分影響）。稅率自前三季實績線性升到長期稅率；股數固定為評價日稀釋股數（未計未來回購與新股，DCF 以 NOPAT 扣股票薪酬處理）。")
year_header(P, "9M26A（前三季實績）")
c = P.cell(PLQ["hd"], 1, "Q4 2026E 單季（三情境共用；US$M）"); c.font = HDR; c.fill = HFILL
for k, (t, f_) in {"gov": ("Government 營收", f"=Seg!D{SB['gov']}"), "com": ("Commercial 營收", f"=Seg!D{SB['com']}"), "rev": ("合併營收", f"=B{PLQ['gov']}+B{PLQ['com']}"),
                    "cg": ("Government contribution", f"=B{PLQ['gov']}*{I('cm_g')}"), "cc": ("Commercial contribution", f"=B{PLQ['com']}*{I('cm_c')}"),
                    "ctb": ("總 contribution", f"=B{PLQ['cg']}+B{PLQ['cc']}"), "ua": ("未分攤成本", f"={I('ua_r')}*B{PLQ['rev']}"), "oi": ("營業利益（GAAP）", f"=B{PLQ['ctb']}-B{PLQ['ua']}"),
                    "cogs": ("營業成本", f"={I('cogs_r')}*B{PLQ['rev']}"), "oth": ("其他收入（淨額；＝利息收入，現金×收益率÷4）", f"={I('yld')}*({I('cash_0')}+{I('sec_0')})/4"),
                    "pbt": ("稅前利益", f"=B{PLQ['oi']}+B{PLQ['oth']}"), "tax": ("所得稅", f"=B{PLQ['pbt']}*{I('tax_0')}"), "ni": ("淨利", f"=B{PLQ['pbt']}-B{PLQ['tax']}")}.items():
    put(P, PLQ[k], 1, t); put(P, PLQ[k], 2, f_, "#,##0.0")
LAB_PL = {"gov": "Government", "com": "Commercial", "rev": "合併營收", "ctb_g": "Government contribution", "ctb_c": "Commercial contribution", "ctb": "總 contribution", "ua": "未分攤成本（含股票薪酬）",
          "oi": "營業利益（GAAP）", "cogs": "營業成本", "gp": "毛利", "opx": "營業費用（殘差）", "oth": "其他收入（淨額）", "pbt": "稅前利益", "taxr": "有效稅率", "tax": "所得稅", "ni": "淨利",
          "sh": "稀釋股數（百萬股）", "eps": "EPS（GAAP 稀釋，美元）"}
for s in SC:
    R = PLR[s]; sfx = "" if s == "B" else f"（情境 {s}）"
    c = P.cell(R["hd"], 1, SCN_NAME[s]); c.font = HDR; c.fill = HFILL
    for k, t in LAB_PL.items():
        put(P, R[k], 1, t); put(P, R[k], 2, "%" if k == "taxr" else "百萬股" if k == "sh" else "US$" if k == "eps" else "US$M")
    put(P, R["chk_seg"], 1, "檢查：分部合計－合併營收" + sfx); put(P, R["chk_pbt"], 1, "檢查：營業利益＋其他收入－稅前利益（歷史對 10-K；應為 0）" + sfx); put(P, R["chk_ni"], 1, "檢查：稅前利益－所得稅－淨利（歷史；應為 0）" + sfx)
    SR = SEGR[s]
    for y in YRS + ["M"]:
        if y == "M":
            col = 13
            put(P, R["gov"], col, f"={q('seg.gov.9M26')}", "#,##0.0"); put(P, R["com"], col, f"={q('seg.com.9M26')}", "#,##0.0"); put(P, R["rev"], col, f"={q('rev.9M26')}", "#,##0.0")
            put(P, R["ctb_g"], col, f"={q('ctb.gov.9M26')}", "#,##0.0"); put(P, R["ctb_c"], col, f"={q('ctb.com.9M26')}", "#,##0.0"); put(P, R["ctb"], col, f"=M{R['ctb_g']}+M{R['ctb_c']}", "#,##0.0")
            put(P, R["oi"], col, f"={q('is.oi.9M26')}", "#,##0.0"); put(P, R["ua"], col, f"=M{R['ctb']}-M{R['oi']}", "#,##0.0"); put(P, R["cogs"], col, f"={q('is.cogs.9M26')}", "#,##0.0")
            put(P, R["gp"], col, f"=M{R['rev']}-M{R['cogs']}", "#,##0.0"); put(P, R["opx"], col, f"=M{R['gp']}-M{R['oi']}", "#,##0.0")
            put(P, R["pbt"], col, f"={q('is.pbt.9M26')}", "#,##0.0"); put(P, R["oth"], col, f"=M{R['pbt']}-M{R['oi']}", "#,##0.0"); put(P, R["tax"], col, f"={q('is.tax.9M26')}", "#,##0.0")
            put(P, R["taxr"], col, f"=M{R['tax']}/M{R['pbt']}", "0.0%"); put(P, R["ni"], col, f"=M{R['pbt']}-M{R['tax']}", "#,##0.0")
            put(P, R["sh"], col, f"={q('sh.dil.q3')}", "#,##0.0"); put(P, R["eps"], col, f"={q('is.eps.9M26')}", "0.00")
            put(P, R["chk_seg"], col, f"=M{R['gov']}+M{R['com']}-M{R['rev']}", "0.000"); put(P, R["chk_ni"], col, f"=M{R['ni']}-{q('is.ni.9M26')}", "0.000")
            continue
        col = 3 + y - 2023; c = yc(y); p = yc(y - 1)
        if y <= 2025:
            put(P, R["gov"], col, f"={q(f'seg.gov.{y}')}", "#,##0.0"); put(P, R["com"], col, f"={q(f'seg.com.{y}')}", "#,##0.0"); put(P, R["rev"], col, f"={q(f'rev.{y}')}", "#,##0.0")
            put(P, R["ctb_g"], col, f"={q(f'ctb.gov.{y}')}", "#,##0.0"); put(P, R["ctb_c"], col, f"={q(f'ctb.com.{y}')}", "#,##0.0")
            put(P, R["oi"], col, f"={q(f'is.oi.{y}')}", "#,##0.0"); put(P, R["ua"], col, f"={c}{R['ctb']}-{c}{R['oi']}", "#,##0.0")
            put(P, R["cogs"], col, f"={q(f'is.cogs.{y}')}", "#,##0.0"); put(P, R["oth"], col, f"={q(f'is.int.{y}')}+{q(f'is.oth.{y}')}", "#,##0.0")
            put(P, R["tax"], col, f"={q(f'is.tax.{y}')}", "#,##0.0"); put(P, R["sh"], col, f"={q(f'is.dsh.{y}')}", "#,##0.0"); put(P, R["eps"], col, f"={q(f'is.eps.{y}')}", "0.00")
            put(P, R["chk_pbt"], col, f"={c}{R['oi']}+{c}{R['oth']}-{q(f'is.pbt.{y}')}", "0.000"); put(P, R["chk_ni"], col, f"={c}{R['pbt']}-{c}{R['tax']}-{q(f'is.ni.{y}')}", "0.000")
        elif y == 2026:
            put(P, R["gov"], col, f"=Seg!F{SR['gov']}", "#,##0.0"); put(P, R["com"], col, f"=Seg!F{SR['com']}", "#,##0.0"); put(P, R["rev"], col, f"=Seg!F{SR['rev']}", "#,##0.0")
            put(P, R["ctb_g"], col, f"=M{R['ctb_g']}+$B${PLQ['cg']}", "#,##0.0"); put(P, R["ctb_c"], col, f"=M{R['ctb_c']}+$B${PLQ['cc']}", "#,##0.0")
            put(P, R["ua"], col, f"=M{R['ua']}+$B${PLQ['ua']}", "#,##0.0"); put(P, R["oi"], col, f"={c}{R['ctb']}-{c}{R['ua']}", "#,##0.0")
            put(P, R["cogs"], col, f"=M{R['cogs']}+$B${PLQ['cogs']}", "#,##0.0"); put(P, R["oth"], col, f"=M{R['oth']}+$B${PLQ['oth']}", "#,##0.0")
            put(P, R["tax"], col, f"=M{R['tax']}+$B${PLQ['tax']}", "#,##0.0"); put(P, R["sh"], col, f"={val('sh_dil')}", "#,##0.0"); put(P, R["eps"], col, f"={c}{R['ni']}/{c}{R['sh']}", "0.00")
        else:
            put(P, R["gov"], col, f"=Seg!{c}{SR['gov']}", "#,##0.0"); put(P, R["com"], col, f"=Seg!{c}{SR['com']}", "#,##0.0"); put(P, R["rev"], col, f"=Seg!{c}{SR['rev']}", "#,##0.0")
            put(P, R["ctb_g"], col, f"={c}{R['gov']}*{I('cm_g')}", "#,##0.0"); put(P, R["ctb_c"], col, f"={c}{R['com']}*{I('cm_c')}-Seg!{c}{SR['incinf']}", "#,##0.0")
            put(P, R["ua"], col, f"=({I('ua_r')}+{I('ua_d')}*({c}$4-2026)/6)*{c}{R['rev']}", "#,##0.0"); put(P, R["oi"], col, f"={c}{R['ctb']}-{c}{R['ua']}", "#,##0.0")
            put(P, R["cogs"], col, f"={I('cogs_r')}*{c}{R['rev']}+Seg!{c}{SR['incinf']}", "#,##0.0")
            put(P, R["oth"], col, f"={I('yld')}*CF!{p}{CFR[s]['cash']}", "#,##0.0")
            put(P, R["taxr"], col, f"={I('tax_0')}+({I('tax_L')}-{I('tax_0')})*({c}$4-2026)/6", "0.0%")
            put(P, R["tax"], col, f"={c}{R['pbt']}*{c}{R['taxr']}", "#,##0.0"); put(P, R["sh"], col, f"={val('sh_dil')}", "#,##0.0"); put(P, R["eps"], col, f"={c}{R['ni']}/{c}{R['sh']}", "0.00")
        put(P, R["ctb"], col, f"={c}{R['ctb_g']}+{c}{R['ctb_c']}", "#,##0.0"); put(P, R["gp"], col, f"={c}{R['rev']}-{c}{R['cogs']}", "#,##0.0")
        put(P, R["opx"], col, f"={c}{R['gp']}-{c}{R['oi']}", "#,##0.0"); put(P, R["pbt"], col, f"={c}{R['oi']}+{c}{R['oth']}", "#,##0.0")
        if y <= 2026: put(P, R["taxr"], col, f"={c}{R['tax']}/{c}{R['pbt']}", "0.0%")
        put(P, R["ni"], col, f"={c}{R['pbt']}-{c}{R['tax']}", "#,##0.0"); put(P, R["chk_seg"], col, f"={c}{R['gov']}+{c}{R['com']}-{c}{R['rev']}", "0.000")
RB = PLR["B"]
c = P.cell(PLS["hd"], 1, "標準損益表（GAAP；情境 B 基準；US$M，EPS 美元）"); c.font = HDR; c.fill = HFILL
for y in YRS: put(P, PLS["hd"], 3 + y - 2023, YL[y], bold=True)
put(P, PLS["hd"], 13, "9M26A（YTD）", bold=True)
std = {"rev": ("營收（百萬美元）", "rev"), "cogs": ("營業成本", "cogs"), "gp": ("毛利", "gp"), "opx": ("營業費用", "opx"), "oi": ("營業利益", "oi"), "oth": ("其他收入（淨額）", "oth"),
       "pbt": ("稅前利益", "pbt"), "tax": ("所得稅", "tax"), "ni": ("稅後淨利", "ni"), "eps": ("EPS（完全稀釋，美元）", "eps")}
for k, (t, src) in std.items():
    put(P, PLS[k], 1, t)
    for col in list(range(3, 14)):
        cl = CL(col); put(P, PLS[k], col, f"={cl}{RB[src]}", "0.00" if k == "eps" else "#,##0.0")
for k, t, num in (("gm", "毛利率", "gp"), ("opm", "營業利益率", "oi"), ("nm", "淨利率", "ni")):
    put(P, PLS[k], 1, t)
    for col in range(3, 14): cl = CL(col); put(P, PLS[k], col, f"={cl}{PLS[num]}/{cl}{PLS['rev']}", pct)
put(P, PLS["yoy"], 1, "營收年增率")
put(P, PLS["yoy"], 3, f"=C{PLS['rev']}/{q('rev.2022')}-1", pct)
for col in range(4, 13): put(P, PLS["yoy"], col, f"={CL(col)}{PLS['rev']}/{CL(col-1)}{PLS['rev']}-1", pct)
put(P, PLS["yoy"], 13, f"=M{PLS['rev']}/{q('rev.9M25')}-1", pct)
put(P, PLS["sub"], 1, "比率區", bold=True)
for k, t, src in (("r_yoy", "營收年增率", "yoy"), ("r_gm", "毛利率", "gm"), ("r_opm", "營業利益率", "opm"), ("r_nm", "淨利率", "nm")):
    put(P, PLS[k], 1, t)
    for col in range(3, 14): cl = CL(col); put(P, PLS[k], col, f"={cl}{PLS[src]}", pct)
put(P, PLS["r_niy"], 1, "淨利年增率")
for col in range(4, 13): put(P, PLS["r_niy"], col, f"={CL(col)}{PLS['ni']}/{CL(col-1)}{PLS['ni']}-1", pct)
put(P, PLS["r_niy"], 13, f"=M{PLS['ni']}/{q('is.ni.9M25')}-1", pct)
put(P, PLS["chk1"], 1, "核對：營業費用（殘差）－歷史 S&M＋R&D＋G&A（應為 0）")
for y in (2023, 2024, 2025): put(P, PLS["chk1"], 3 + y - 2023, f"={yc(y)}{PLS['opx']}-({q(f'is.sm.{y}')}+{q(f'is.rd.{y}')}+{q(f'is.ga.{y}')})", "0.000")
put(P, PLS["chk2"], 1, "核對：毛利－營業費用－營業利益（恆等式，應為 0）")
for col in range(3, 14): cl = CL(col); put(P, PLS["chk2"], col, f"={cl}{PLS['gp']}-{cl}{PLS['opx']}-{cl}{PLS['oi']}", "0.000")
P.column_dimensions["A"].width = 64
for col in "BCDEFGHIJKLM": P.column_dimensions[col].width = 12

# =====================================================================================================
# CF
C_ = SH["CF"]
C_["A1"] = "CF — 現金流（US$M）：營業現金流＝淨利＋股票薪酬＋折舊攤銷＋營運資金與其他調整；自由現金流＝營業現金流－資本支出（輕資產：無自建算力）。FY2026E＝前三季實績＋Q4E。現金＝現金＋有價證券"
C_["A2"] = "收支關卡＝FY2026–FY2030 自由現金流扣股利合計（含 FY2026 前三季實績；公司無股利）；負值＝入不敷出（原則第 63 條）。營運資金調整＝比率（負值，歷史為現金流出）× 營收年增量（來源見 In wc_k）。"
year_header(C_, "9M26A（前三季實績）")
c = C_.cell(CFQ["hd"], 1, "Q4 2026E 單季（三情境共用；US$M）"); c.font = HDR; c.fill = HFILL
for k, (t, f_) in {"sbc": ("股票薪酬", f"={I('sbc_r')}*PL!B{PLQ['rev']}"), "da": ("折舊攤銷", f"={I('da_r')}*PL!B{PLQ['rev']}"),
                   "wc": ("營運資金與其他調整（營收季增量 × wc_k）", f"={I('wc_k')}*(PL!B{PLQ['rev']}-Seg!B{SB['rev']})"),
                   "cfo": ("營業現金流", f"=PL!B{PLQ['ni']}+B{CFQ['sbc']}+B{CFQ['da']}+B{CFQ['wc']}"), "cx": ("資本支出", f"={I('cx_r')}*PL!B{PLQ['rev']}"),
                   "fcf": ("自由現金流", f"=B{CFQ['cfo']}-B{CFQ['cx']}")}.items():
    put(C_, CFQ[k], 1, t); put(C_, CFQ[k], 2, f_, "#,##0.0")
LAB_CF = {"ni": "淨利", "sbc": "＋股票薪酬", "da": "＋折舊攤銷", "wc": "＋營運資金與其他調整", "cfo": "營業現金流", "cx": "－資本支出", "fcf": "自由現金流", "div": "－股利（無）",
          "fcfd": "自由現金流扣股利", "cash": "期末現金＋有價證券", "oth": "（保留）", "chk": "檢查：歷史營業現金流－財報（應為 0）"}
for s in SC:
    R = CFR[s]; PR = PLR[s]
    c = C_.cell(R["hd"], 1, SCN_NAME[s]); c.font = HDR; c.fill = HFILL
    for k, t in LAB_CF.items():
        if k == "oth": continue
        put(C_, R[k], 1, t); put(C_, R[k], 2, "US$M")
    # 9M26 欄
    put(C_, R["ni"], 13, f"=PL!M{PR['ni']}", "#,##0.0"); put(C_, R["sbc"], 13, f"={q('is.sbc.9M26')}", "#,##0.0"); put(C_, R["da"], 13, f"={q('is.da.9M26')}", "#,##0.0")
    put(C_, R["cfo"], 13, f"={q('is.cfo.9M26')}", "#,##0.0"); put(C_, R["wc"], 13, f"=M{R['cfo']}-M{R['ni']}-M{R['sbc']}-M{R['da']}", "#,##0.0")
    put(C_, R["cx"], 13, f"={q('is.capex.9M26')}", "#,##0.0"); put(C_, R["fcf"], 13, f"=M{R['cfo']}-M{R['cx']}", "#,##0.0")
    for y in YRS:
        col = 3 + y - 2023; c = yc(y); p = yc(y - 1)
        put(C_, R["ni"], col, f"=PL!{c}{PR['ni']}", "#,##0.0")
        if y <= 2025:
            put(C_, R["sbc"], col, f"={q(f'is.sbc.{y}')}", "#,##0.0"); put(C_, R["da"], col, f"={q(f'is.da.{y}')}", "#,##0.0"); put(C_, R["cfo"], col, f"={q(f'is.cfo.{y}')}", "#,##0.0")
            put(C_, R["wc"], col, f"={c}{R['cfo']}-{c}{R['ni']}-{c}{R['sbc']}-{c}{R['da']}", "#,##0.0"); put(C_, R["cx"], col, f"={q(f'is.capex.{y}')}", "#,##0.0")
            put(C_, R["chk"], col, f"={c}{R['ni']}+{c}{R['sbc']}+{c}{R['da']}+{c}{R['wc']}-{c}{R['cfo']}", "0.000")
            if y == 2025: put(C_, R["cash"], col, f"={q('bs.cash25')}+{q('bs.sec25')}", "#,##0.0")
        elif y == 2026:
            put(C_, R["sbc"], col, f"=M{R['sbc']}+$B${CFQ['sbc']}", "#,##0.0"); put(C_, R["da"], col, f"=M{R['da']}+$B${CFQ['da']}", "#,##0.0")
            put(C_, R["wc"], col, f"=M{R['wc']}+$B${CFQ['wc']}", "#,##0.0"); put(C_, R["cx"], col, f"=M{R['cx']}+$B${CFQ['cx']}", "#,##0.0")
            put(C_, R["cfo"], col, f"={c}{R['ni']}+{c}{R['sbc']}+{c}{R['da']}+{c}{R['wc']}", "#,##0.0")
            put(C_, R["cash"], col, f"={I('cash_0')}+{I('sec_0')}+$B${CFQ['fcf']}", "#,##0.0")
        else:
            put(C_, R["sbc"], col, f"={I('sbc_r')}*PL!{c}{PR['rev']}", "#,##0.0"); put(C_, R["da"], col, f"={I('da_r')}*PL!{c}{PR['rev']}", "#,##0.0")
            put(C_, R["wc"], col, f"={I('wc_k')}*(PL!{c}{PR['rev']}-PL!{p}{PR['rev']})", "#,##0.0"); put(C_, R["cx"], col, f"={I('cx_r')}*PL!{c}{PR['rev']}", "#,##0.0")
            put(C_, R["cfo"], col, f"={c}{R['ni']}+{c}{R['sbc']}+{c}{R['da']}+{c}{R['wc']}", "#,##0.0")
            put(C_, R["cash"], col, f"={p}{R['cash']}+{c}{R['fcfd']}", "#,##0.0")
        put(C_, R["fcf"], col, f"={c}{R['cfo']}-{c}{R['cx']}", "#,##0.0"); put(C_, R["div"], col, 0, "#,##0.0"); put(C_, R["fcfd"], col, f"={c}{R['fcf']}-{c}{R['div']}", "#,##0.0")
c = C_.cell(CF_GATE - 1, 1, "收支關卡與期間流動性（原則第 55、63 條）"); c.font = HDR; c.fill = HFILL
for j, h in enumerate(["項目", "情境 A", "情境 B", "情境 C"], 1): put(C_, CF_GATE, j, h, bold=True) if False else None
GR = CF_GATE_ROWS
put(C_, GR["gate"], 1, "收支關卡：FY2026–FY2030 自由現金流扣股利合計（含 FY2026 前三季實績）")
put(C_, GR["gate_ok"], 1, "收支關卡判定（合計 ≥ 0 為通過）"); put(C_, GR["liq_max"], 1, "期間最低年末現金＋有價證券（FY2026–FY2030）"); put(C_, GR["liq_cash"], 1, "FY2026–FY2030 是否任一年自由現金流為負（年數）")
for j, s in enumerate("ABC"):
    R = CFR[s]; cl = "BCD"[j]
    put(C_, GR["gate"], 2 + j, f"=SUM(F{R['fcfd']}:J{R['fcfd']})", "#,##0.0"); put(C_, GR["gate_ok"], 2 + j, f'=IF({cl}{GR["gate"]}>=0,"通過","不通過")')
    put(C_, GR["liq_max"], 2 + j, f"=MIN(F{R['cash']}:J{R['cash']})", "#,##0.0"); put(C_, GR["liq_cash"], 2 + j, f'=COUNTIF(F{R["fcfd"]}:J{R["fcfd"]},"<0")', "0")
C_.column_dimensions["A"].width = 64
for col in "BCDEFGHIJKLM": C_.column_dimensions[col].width = 12

# =====================================================================================================
# DCF
D_ = SH["DCF"]
D_["A1"] = "DCF — 無槓桿自由現金流（FCFF）折現：NOPAT＝GAAP 營業利益×(1−稅率)，不加回股票薪酬（以費用計）；期中折現；評價日 2026-09-30，FY2026 只計 Q4；終值＝NOPAT(FY2033)×(1−g÷RONIC)÷(WACC−g)，RONIC＝WACC"
D_["A2"] = "每股價值＝(EV＋淨現金)÷評價日稀釋股數，再以 ke 推到目標時點（股價基準日＋12 個月）；淨現金＝現金＋有價證券−借款（營業租賃不計：營業利益已含租金）。三情境各一塊，加權見 Val。"
put(D_, 3, 1, "項目", bold=True); put(D_, 4, 1, "年（數值）")
for y in YRS[3:]: put(D_, 3, 3 + y - 2023, YL[y] if y > 2026 else "Q4 2026E", bold=True); put(D_, 4, 3 + y - 2023, y)
LAB_D = {"hd": "", "oi": "營業利益（GAAP）", "taxr": "稅率", "nopat": "NOPAT", "da": "＋折舊攤銷", "cx": "－資本支出", "wc": "＋營運資金與其他調整", "fcff": "FCFF",
         "t": "折現年數（評價日起，期中）", "df": "折現因子", "pv": "現值", "sumpv": "明確期現值合計", "tvn": "FY2033 NOPAT", "tv": "終值（FY2032 年底）", "tvt": "終值折現年數",
         "pvtv": "終值現值", "ev": "企業價值 EV", "nc": "＋淨現金（現金＋有價證券−借款）", "eq": "股權價值", "ps": "每股價值（評價日）", "pstp": "每股價值（推到目標時點）", "tvsh": "終值占 EV", "sbc": "股票薪酬（加回對照；FY2026 只計 Q4）", "pvsb": "股票薪酬稅後現值", "pvsbtv": "股票薪酬終值現值", "evsb": "股票薪酬加回的 EV 增量", "pssb": "DCF（股票薪酬加回）每股價值（評價日）", "pstpsb": "DCF（股票薪酬加回）每股價值（目標時點）"}
for s in SC:
    R = DR[s]; PR = PLR[s]; CR_ = CFR[s]
    c = D_.cell(R["hd"], 1, SCN_NAME[s]); c.font = HDR; c.fill = HFILL
    for k, t in LAB_D.items():
        if k != "hd": put(D_, R[k], 1, t)
    for y in range(2026, 2033):
        col = 3 + y - 2023; c = yc(y)
        if y == 2026:
            put(D_, R["oi"], col, f"=PL!B{PLQ['oi']}", "#,##0.0"); put(D_, R["taxr"], col, f"={I('tax_0')}", "0.0%")
            put(D_, R["da"], col, f"=CF!B{CFQ['da']}", "#,##0.0"); put(D_, R["cx"], col, f"=CF!B{CFQ['cx']}", "#,##0.0"); put(D_, R["wc"], col, f"=CF!B{CFQ['wc']}", "#,##0.0")
            put(D_, R["t"], col, 0.125, "0.000")
        else:
            put(D_, R["oi"], col, f"=PL!{c}{PR['oi']}", "#,##0.0"); put(D_, R["taxr"], col, f"=PL!{c}{PR['taxr']}", "0.0%")
            put(D_, R["da"], col, f"=CF!{c}{CR_['da']}", "#,##0.0"); put(D_, R["cx"], col, f"=CF!{c}{CR_['cx']}", "#,##0.0"); put(D_, R["wc"], col, f"=CF!{c}{CR_['wc']}", "#,##0.0")
            put(D_, R["t"], col, f"={c}$4-2026-0.25", "0.000")
        put(D_, R["nopat"], col, f"={c}{R['oi']}*(1-{c}{R['taxr']})", "#,##0.0"); put(D_, R["fcff"], col, f"={c}{R['nopat']}+{c}{R['da']}-{c}{R['cx']}+{c}{R['wc']}", "#,##0.0")
        put(D_, R["df"], col, f"=1/(1+{val('wacc')})^{c}{R['t']}", "0.0000"); put(D_, R["pv"], col, f"={c}{R['fcff']}*{c}{R['df']}", "#,##0.0")
    for y in range(2026, 2033):
        col = 3 + y - 2023; c = yc(y)
        put(D_, R["sbc"], col, f"=CF!B{CFQ['sbc']}" if y == 2026 else f"=CF!{c}{CR_['sbc']}", "#,##0.0")
        put(D_, R["pvsb"], col, f"={c}{R['sbc']}*(1-{c}{R['taxr']})*{c}{R['df']}", "#,##0.0")
    put(D_, R["pvsbtv"], 3, f"=IF({I('tv_alt')}=1,L{R['sbc']}*(1-L{R['taxr']})*(1+{val('g')}),L{R['sbc']}*(1-L{R['taxr']})*(1+{val('g')})*(1-{val('g')}/{val('ronic')}))/({val('wacc')}-{val('g')})/(1+{val('wacc')})^6.25", "#,##0.0")
    put(D_, R["evsb"], 3, f"=SUM(F{R['pvsb']}:L{R['pvsb']})+C{R['pvsbtv']}", "#,##0.0")
    put(D_, R["pssb"], 3, f"=(C{R['ev']}+C{R['evsb']}+C{R['nc']})/{val('sh_dil')}", "#,##0.00"); put(D_, R["pstpsb"], 3, f"=C{R['pssb']}*(1+{val('ke')})^{val('tau')}", "#,##0.00")
    put(D_, R["sumpv"], 3, f"=SUM(F{R['pv']}:L{R['pv']})", "#,##0.0")
    put(D_, R["tvn"], 3, f"=L{R['nopat']}*(1+{val('g')})", "#,##0.0")
    put(D_, R["tv"], 3, f"=IF({I('tv_alt')}=1,L{R['fcff']}*(1+{val('g')}),C{R['tvn']}*(1-{val('g')}/{val('ronic')}))/({val('wacc')}-{val('g')})", "#,##0.0")
    put(D_, R["tvt"], 3, "=6.25", "0.00"); put(D_, R["pvtv"], 3, f"=C{R['tv']}/(1+{val('wacc')})^C{R['tvt']}", "#,##0.0")
    put(D_, R["ev"], 3, f"=C{R['sumpv']}+C{R['pvtv']}", "#,##0.0"); put(D_, R["nc"], 3, f"={val('nc0')}", "#,##0.0")
    put(D_, R["eq"], 3, f"=C{R['ev']}+C{R['nc']}", "#,##0.0"); put(D_, R["ps"], 3, f"=C{R['eq']}/{val('sh_dil')}", "#,##0.00")
    put(D_, R["pstp"], 3, f"=C{R['ps']}*(1+{val('ke')})^{val('tau')}", "#,##0.00"); put(D_, R["tvsh"], 3, f"=C{R['pvtv']}/C{R['ev']}", "0.0%")
D_.column_dimensions["A"].width = 52
for col in "BCDEFGHIJKL": D_.column_dimensions[col].width = 12

# =====================================================================================================
# Val
Vs = SH["Val"]
Vs["A1"] = "Val — 目標價：DCF 與 Comparable（同業 P/S × 共用庫相對倍數）以 In 頁 DCF 權重加權；情境 A／B／C 以 In 頁權重加權。目標時點＝股價基準日＋12 個月；上檔空間是持有期間報酬，不是年化"
Vs["A2"] = "評等五級固定規則（原則第 24 條）：上檔空間 > +50% 強力買進、> +25% 買進、≥ −25% 中立、≥ −50% 放空、其餘強力放空；收支關卡不通過且評等為買進以上時，投資建議為「收支不足」（第 63 條）。B、C、D 欄＝情境 A、B、C"
def vp(k, label, vals, fmt=None, bold=False, note=None):
    r = VR[k]; put(Vs, r, 1, label, bold=bold)
    for j, v in enumerate(vals):
        if v is None: continue
        put(Vs, r, 2 + j, v, fmt[j] if isinstance(fmt, list) else fmt, bold=bold)
    if note: put(Vs, r, 6, note)
def cap(k, t):
    c = Vs.cell(VR[k], 1, t); c.font = HDR; c.fill = HFILL
    for j in range(2, 7): Vs.cell(VR[k], j).fill = HFILL
W3 = lambda row: f"{I('tp_wA')}*B{row}+{I('tp_wB')}*C{row}+{I('tp_wC')}*D{row}"
S3 = lambda fn: [fn(s) for s in "ABC"]
cap("cap1", "1. 目標價與評等")
vp("px", "現價（US$）", [f"={I('px')}"], "#,##0.00", note="stockanalysis.com 2026-11-13 收盤；2026-11-09 另源 169.20 交叉核對")
vp("pxd", "股價基準日", [f"={I('px_date')}"], "yyyy-mm-dd")
vp("tpd", "目標時點（＝股價基準日＋12 個月）", [f"={I('tp_date')}"], "yyyy-mm-dd")
vp("tpdcf", "DCF 目標價（情境加權）", [f"={W3(VR['d_pstp'])}"], "#,##0.00", bold=True)
vp("tpcomp", "Comparable 目標價（情境加權）", [f"={W3(VR['ps_c'])}"], "#,##0.00", bold=True)
vp("tp", "目標價（DCF 與 Comparable 加權）", [f"={I('w_dcf')}*B{VR['tpdcf']}+(1-{I('w_dcf')})*B{VR['tpcomp']}"], "#,##0.00", bold=True)
vp("cons", "賣方平均目標價（只作對照；2026-08-31，Q3 財報前，標「逾期」）", [f"={I('cons_tp')}"], "#,##0.00")
U = f"B{VR['up']}"
vp("up", "上檔空間：加權目標價", [f"=B{VR['tp']}/B{VR['px']}-1"], "0.0%", bold=True)
vp("updcf", "上檔空間：DCF 目標價", [f"=B{VR['tpdcf']}/B{VR['px']}-1"], "0.0%")
vp("upcomp", "上檔空間：Comparable 目標價", [f"=B{VR['tpcomp']}/B{VR['px']}-1"], "0.0%")
def recf(u): return f'IF({u}>0.5,"強力買進",IF({u}>0.25,"買進",IF({u}>=-0.25,"中立",IF({u}>=-0.5,"放空","強力放空"))))'
def tierf(u): return f'IF({u}>0.5,"大於 50%",IF({u}>0.25,"25%–50%",IF({u}>=0,"0%–25%",IF({u}>=-0.25,"0%～−25%",IF({u}>=-0.5,"−25%～−50%","小於 −50%")))))'
vp("tpdcf_sbc", "DCF（股票薪酬加回）目標價（情境加權；只作對照，不進加權目標價）", [f"={W3(VR['d_pstp_sbc'])}"], "#,##0.00")
vp("updcf_sbc", "上檔空間：DCF（股票薪酬加回）目標價（只作對照）", [f"=B{VR['tpdcf_sbc']}/B{VR['px']}-1"], "0.0%")
vp("tier", "上檔空間級距", ["=" + tierf(U)], "@")
vp("est", "估值評等（依上檔空間五級）", ["=" + recf(U)], "@", bold=True)
vp("gate", "入不敷出判定用：FY2026–30 自由現金流扣股利合計（B）", [f"=CF!C{GR['gate']}"], "#,##0.0")
vp("rec", "投資建議", [f'=IF(AND(B{VR["gate"]}<0,OR(B{VR["est"]}="強力買進",B{VR["est"]}="買進")),"收支不足",B{VR["est"]})'], "@", bold=True)
vp("bdist", "距最近評等邊界（百分點）／升一級所需目標價／降一級的目標價", [
    f"=MIN(ABS({U}-0.5),ABS({U}-0.25),ABS({U}+0.25),ABS({U}+0.5))*100",
    f"=B{VR['px']}*(1+IF({U}>0.5,9.99,IF({U}>0.25,0.5,IF({U}>=-0.25,0.25,IF({U}>=-0.5,-0.25,-0.5)))))",
    f"=B{VR['px']}*(1+IF({U}>0.5,0.5,IF({U}>0.25,0.25,IF({U}>=-0.25,-0.25,IF({U}>=-0.5,-0.5,-9.99)))))"], ["0.0", "#,##0.00", "#,##0.00"])
vp("att_dcf", "上檔空間歸因：DCF 貢獻（權重×DCF 上檔）", [f"={I('w_dcf')}*B{VR['updcf']}"], "+0.0%;-0.0%")
vp("att_cmp", "上檔空間歸因：Comparable 貢獻（權重×未調整 Comparable 上檔）", [f"=(1-{I('w_dcf')})*(({W3(VR['ps_u'])})/B{VR['px']}-1)"], "+0.0%;-0.0%")
vp("att_adj", "上檔空間歸因：相對倍數調整（共用庫相對倍數＋成長調整）", [f"=(1-{I('w_dcf')})*(B{VR['upcomp']}-(({W3(VR['ps_u'])})/B{VR['px']}-1))"], "+0.0%;-0.0%")
cap("cap2", "2. 評價日稀釋股數（百萬股；原則第 11、34 條）")
vp("sh_basic", "流通股（Class A＋B＋F；2026-10-26 10-Q 封面）", [f"={I('sh_a')}+{I('sh_b')}+{I('sh_f')}"], "#,##0.0")
vp("sh_rsu", "＋未歸屬 RSU（全額；預期沒收率公司未揭露，未找到）", [f"={I('rsu_n')}"], "#,##0.0")
vp("sh_opt", "＋選擇權（庫藏股法，以現價）", [f"=MAX(0,{I('opt_n')}*(1-{I('opt_k')}/{I('px')}))"], "#,##0.0")
vp("sh_sar", "＋SAR（庫藏股法，以現價）", [f"=MAX(0,{I('sar_n')}*(1-{I('sar_k')}/{I('px')}))"], "#,##0.0")
vp("sh_dil", "評價日稀釋股數", [f"=B{VR['sh_basic']}+B{VR['sh_rsu']}+B{VR['sh_opt']}+B{VR['sh_sar']}"], "#,##0.0", bold=True)
vp("nc0", "淨現金（現金＋有價證券−借款；營業租賃不計）", [f"={I('cash_0')}+{I('sec_0')}-{I('debt_0')}"], "#,##0.0")
cap("cap3", "3. 折現率")
vp("rf", "無風險利率（2026-11-12；FRED DGS10；現價日前 1 個交易日）", [f"={I('rf')}"], "0.00%")
vp("beta", "β（Blume 調整，兩源中位數）", [f"={I('beta')}"], "0.000")
vp("erp", "股權風險溢酬", [f"={I('erp')}"], "0.00%")
vp("ke", "股權成本 ke", [f"=B{VR['rf']}+B{VR['beta']}*B{VR['erp']}"], "0.00%")
vp("kd", "債務成本（rf＋利差；無借款，權重為 0）", [f"=B{VR['rf']}+{I('kd_sp')}"], "0.00%")
vp("kdat", "稅後債務成本", [f"=B{VR['kd']}*(1-{I('tax_L')})"], "0.00%")
vp("mcap", "市值（現價×稀釋股數）", [f"=B{VR['px']}*B{VR['sh_dil']}"], "#,##0.0")
vp("debt", "債務", [f"={I('debt_0')}"], "#,##0.0")
vp("we", "股權權重", [f"=B{VR['mcap']}/(B{VR['mcap']}+B{VR['debt']})"], "0.0%")
vp("wacc", "WACC（含調整）", [f"=B{VR['we']}*B{VR['ke']}+(1-B{VR['we']})*B{VR['kdat']}+{I('wacc_adj')}"], "0.00%", bold=True)
vp("ronic", "永續期 RONIC", [f"=B{VR['wacc']}+{I('ronic_sp')}"], "0.00%")
vp("g", "永續成長率", [f"={I('g_tv')}"], "0.00%")
vp("tau", "評價日至目標時點（年）", [f"=({I('tp_date')}-{I('val_date')})/365"], "0.000")
vp("tau_h", "股價基準日至目標時點（年）", [f"=({I('tp_date')}-{I('px_date')})/365"], "0.000")
cap("cap4", "4. DCF 彙總（B、C、D＝情境 A、B、C）")
for k, lab, rk, fm in (("d_pv", "明確期現值合計", "sumpv", "#,##0.0"), ("d_pvtv", "終值現值", "pvtv", "#,##0.0"), ("d_ev", "企業價值 EV", "ev", "#,##0.0"), ("d_nc", "淨現金", "nc", "#,##0.0"),
                       ("d_eq", "股權價值", "eq", "#,##0.0"), ("d_ps", "每股價值（評價日）", "ps", "#,##0.00"), ("d_pstp", "每股價值（目標時點）", "pstp", "#,##0.00"), ("d_tvsh", "終值占 EV", "tvsh", "0.0%")):
    vp(k, lab, S3(lambda s, rk=rk: f"=DCF!C{DR[s][rk]}"), fm, bold=(k == "d_pstp"))
vp("d_pstp_sbc", "DCF（股票薪酬加回）每股價值（目標時點；只作對照）", S3(lambda s: f"=DCF!C{DR[s]['pstpsb']}"), "#,##0.00")
cap("cap5", "5. Comparable 主腿：同業 P/S（現價日 2026-11-13 口徑）×共用庫相對倍數×本公司目標時點 TTM 營收 ÷ 稀釋股數（P/S 為市值口徑，不另加淨現金；G6）")
cmk = [p[0].lower() for p in PEERS if p[2] == "商業"]; gvk = [p[0].lower() for p in PEERS if p[2] == "政府"]
vp("m_c", "8 家同業合併中位 P/S（peer_ps_20261113.json all 群組，現價日口徑；整體單一口徑）", ["=MEDIAN(" + ",".join(I(f"pm_{k}") for k in cmk + gvk) + ")"], "0.00")
vp("m_g", "同上（Government 分部沿用合併中位；不再單獨取政府同業）", [f"=$B${VR['m_c']}"], "0.00")
vp("pg_c", "商業同業賣方 3 年營收成長中位數", ["=MEDIAN(" + ",".join(I(f"pg_{k}") for k in cmk) + ")"], "0.0%")
vp("pg_g", "政府服務同業賣方 3 年營收成長中位數", ["=MEDIAN(" + ",".join(I(f"pg_{k}") for k in gvk) + ")"], "0.0%")
vp("sg_c", "Commercial 分部 FY2026–FY2029 營收年複合成長", S3(lambda s: f"=(Seg!J{SEGR[s]['com']}/Seg!F{SEGR[s]['com']})^(1/3)-1"), "0.0%")
vp("sg_g", "Government 分部 FY2026–FY2029 營收年複合成長", S3(lambda s: f"=(Seg!J{SEGR[s]['gov']}/Seg!F{SEGR[s]['gov']})^(1/3)-1"), "0.0%")
vp("rr_c", "Commercial 成長比值（原始）", S3(lambda s: f"={SCOL[s]}{VR['sg_c']}/$B${VR['pg_c']}"), "0.00")
vp("rc_c", "Commercial 成長比值（夾 gcap 後）", S3(lambda s: f"=MIN({I('gcap')},{SCOL[s]}{VR['rr_c']})"), "0.00")
vp("rr_g", "Government 成長比值（原始）", S3(lambda s: f"={SCOL[s]}{VR['sg_g']}/$B${VR['pg_g']}"), "0.00")
vp("rc_g", "Government 成長比值（夾 gcap 後）", S3(lambda s: f"=MIN({I('gcap')},{SCOL[s]}{VR['rr_g']})"), "0.00")
vp("rl_c", "Commercial 相對倍數（整體 geo_base；In 頁）", [f"={I('rel_c')}"], "0.00")
vp("rl_g", "Government 相對倍數（整體 geo_base；In 頁，與 Commercial 同值）", [f"={I('rel_g')}"], "0.00")
vp("ma_c", "Commercial 採用倍數＝同業中位數×相對倍數×(成長比值)^θ（θ 基準 0）", S3(lambda s: f"=$B${VR['m_c']}*$B${VR['rl_c']}*{SCOL[s]}{VR['rc_c']}^{I('gth')}"), "0.00")
vp("ma_g", "Government 採用倍數", S3(lambda s: f"=$B${VR['m_g']}*$B${VR['rl_g']}*{SCOL[s]}{VR['rc_g']}^{I('gth')}"), "0.00")
vp("ftp", "目標日 TTM 中 FY2027 占比（TTM 至 2027Q3＝Q4'26＋前三季'27，近似 25%／75%）", ["=0.75"], "0.00")
vp("ttm_c", "Commercial 目標日 TTM 營收", S3(lambda s: f"=$B${VR['ftp']}*Seg!G{SEGR[s]['com']}+(1-$B${VR['ftp']})*Seg!F{SEGR[s]['com']}"), "#,##0.0")
vp("ttm_g", "Government 目標日 TTM 營收", S3(lambda s: f"=$B${VR['ftp']}*Seg!G{SEGR[s]['gov']}+(1-$B${VR['ftp']})*Seg!F{SEGR[s]['gov']}"), "#,##0.0")
vp("ev_c", "Commercial 隱含市值（合併倍數 × 分部營收，拆分僅供檢視）", S3(lambda s: f"={SCOL[s]}{VR['ma_c']}*{SCOL[s]}{VR['ttm_c']}"), "#,##0.0")
vp("ev_g", "Government 隱含市值（合併倍數 × 分部營收，拆分僅供檢視）", S3(lambda s: f"={SCOL[s]}{VR['ma_g']}*{SCOL[s]}{VR['ttm_g']}"), "#,##0.0")
vp("ev_t", "合計隱含市值（P/S 口徑）", S3(lambda s: f"={SCOL[s]}{VR['ev_c']}+{SCOL[s]}{VR['ev_g']}"), "#,##0.0")
vp("nc_t", "淨現金（目標日；僅供對照，P/S 口徑不加入；CF 現金 FY2026／FY2027 以 25%／75% 內插，減借款）", S3(lambda s: f"=(1-$B${VR['ftp']})*CF!F{CFR[s]['cash']}+$B${VR['ftp']}*CF!G{CFR[s]['cash']}-{I('debt_0')}"), "#,##0.0")
vp("eq_t", "股權價值（＝P/S 口徑市值，不另加淨現金）", S3(lambda s: f"={SCOL[s]}{VR['ev_t']}"), "#,##0.0")
vp("ps_c", "Comparable 每股價值（目標時點；成長調整後）", S3(lambda s: f"={SCOL[s]}{VR['eq_t']}/$B${VR['sh_dil']}"), "#,##0.00", bold=True)
vp("ps_u", "Comparable 對照：相對倍數＝1.00（同業中位 P/S）每股價值", S3(lambda s: f"=($B${VR['m_c']}*{SCOL[s]}{VR['ttm_c']}+$B${VR['m_g']}*{SCOL[s]}{VR['ttm_g']})/$B${VR['sh_dil']}"), "#,##0.00")
vp("imp_m", "隱含合併 P/S（倍）", S3(lambda s: f"={SCOL[s]}{VR['ev_t']}/({SCOL[s]}{VR['ttm_c']}+{SCOL[s]}{VR['ttm_g']})"), "0.00")
vp("clamp", "是否觸及成長比值上限（原則第 38 條：揭露原始值與夾值）", S3(lambda s: f'=IF(OR({SCOL[s]}{VR["rr_c"]}>{I("gcap")},{SCOL[s]}{VR["rr_g"]}>{I("gcap")}),"觸及夾值：商業原始 "&TEXT({SCOL[s]}{VR["rr_c"]},"0.00")&"、政府原始 "&TEXT({SCOL[s]}{VR["rr_g"]},"0.00")&" 夾到 "&TEXT({I("gcap")},"0.0"),"未觸及")'), "@")
vp("meth", "目標價方法", [f'="DCF（FCFF、WACC、RONIC＝WACC）"&TEXT({I("w_dcf")},"0%")&"＋Comparable（同業 P/S × 共用庫相對倍數；θ＝"&TEXT({I("gth")},"0.0")&"）"&TEXT(1-{I("w_dcf")},"0%")&"；情境權重 A／B／C＝"&TEXT({I("tp_wA")},"0%")&"／"&TEXT({I("tp_wB")},"0%")&"／"&TEXT({I("tp_wC")},"0%")'], "@")
Vs.column_dimensions["A"].width = 70
for col in "BCDE": Vs.column_dimensions[col].width = 14
Vs.column_dimensions["F"].width = 60

# =====================================================================================================
# RevDCF（反向 DCF：現價隱含什麼；情境 B 路徑）
RV = SH["RevDCF"] if "RevDCF" in SH else None
RV = SH["RevDCF"]
RVK = ["hd1", "sum_dcf", "sum_cmp", "px", "mkt_ev", "mod_ev", "scale", "rv28_0", "rv28_1", "cagr0", "cagr1", "w0", "w1", "hd2", "tgrid", "ctt", "cur_mult", "peer_mult", "cur_rel", "imp_up", "hd3"]
RVR = alloc(4, RVK)
GRID = [0.05, 0.06, 0.07, 0.08, 0.09, 0.10, 0.11, 0.12, 0.13, 0.14, 0.15, 0.16, 0.18, 0.20]
RV["A1"] = "RevDCF — 反向 DCF：現價隱含什麼（情境 B 路徑）。現價隱含企業價值＝現價×稀釋股數−淨現金；以兩種單變數反解對照模型基準：（一）整條自由現金流路徑同乘一個倍數；（二）WACC。"
RV["A2"] = "兩種反解都不改營業利益率等其他假設；倍數反解為線性近似（FCFF 全期同乘；終值同乘）；WACC 反解以格點內插（終值 RONIC 維持 WACC＋In 頁差值）。主導變數的整本重算反解見 Recon 第 4 節。"
def rvp(k, lab, v, fmt=None, bold=False, extra=None):
    put(RV, RVR[k], 1, lab, bold=bold)
    if v is not None: put(RV, RVR[k], 2, v, fmt, bold=bold)
    for j, e in enumerate(extra or [], 3): put(RV, RVR[k], j, e[0], e[1])
DB = DR["B"]
c = RV.cell(RVR["hd1"], 1, "1. 現金流折現法：現價隱含 vs 模型"); c.font = HDR; c.fill = HFILL
rvp("px", "現價（US$）", f"={val('px')}", "#,##0.00")
rvp("mkt_ev", "現價隱含企業價值 EV（現價×稀釋股數−淨現金；US$M）", f"={val('px')}*{val('sh_dil')}-{val('nc0')}", "#,##0")
rvp("mod_ev", "模型企業價值 EV（情境 B，評價日；US$M）", f"=DCF!C{DB['ev']}", "#,##0")
rvp("scale", "現價隱含的現金流倍數（現價隱含 EV ÷ 模型 EV；＝各年自由現金流同乘此倍數）", f"=B{RVR['mkt_ev']}/B{RVR['mod_ev']}", "0.00")
rvp("rv28_0", "FY2028 營收：模型（US$M）", f"=PL!H{PLR['B']['rev']}", "#,##0")
rvp("rv28_1", "FY2028 營收：現金流同乘倍數且利益率不變時（US$M）", f"=B{RVR['rv28_0']}*B{RVR['scale']}", "#,##0")
rvp("cagr0", "FY2026–FY2028 營收年複合成長：模型", f"=(B{RVR['rv28_0']}/PL!F{PLR['B']['rev']})^(1/2)-1", "0.0%")
rvp("cagr1", "FY2026–FY2028 營收年複合成長：現價隱含", f"=(B{RVR['rv28_1']}/PL!F{PLR['B']['rev']})^(1/2)-1", "0.0%")
rvp("w0", "WACC：模型", f"={val('wacc')}", "0.00%")
rvp("w1", "WACC：現價隱含（格點內插；超出範圍則顯示說明）", None)
RV.cell(RVR["w1"], 2).value = None
c = RV.cell(RVR["hd2"], 1, "2. WACC 格點（情境 B 自由現金流；EV 隨 WACC 遞減）"); c.font = HDR; c.fill = HFILL
put(RV, RVR["tgrid"], 1, "WACC 格點／對應 EV（US$M）", bold=True)
GR0 = RVR["tgrid"] + 1
# 格點放 D 欄起：直向 14 列
for i, w in enumerate(GRID):
    rr = RVR["hd3"] + 2 + i
    put(RV, rr, 1, "格點 WACC／EV"); put(RV, rr, 2, w, "0.0%", blue=True)
    put(RV, rr, 3, f"=SUMPRODUCT(DCF!$F${DB['fcff']}:$L${DB['fcff']}/(1+B{rr})^DCF!$F${DB['t']}:$L${DB['t']})+IF({I('tv_alt')}=1,DCF!$L${DB['fcff']}*(1+{val('g')}),DCF!$L${DB['nopat']}*(1+{val('g')})*(1-{val('g')}/(B{rr}+{I('ronic_sp')})))/(B{rr}-{val('g')})/(1+B{rr})^6.25", "#,##0")
GR1 = RVR["hd3"] + 2 + len(GRID) - 1; GR0 = RVR["hd3"] + 2
c = RV.cell(RVR["hd3"], 1, "WACC 格點明細（B＝WACC，C＝情境 B 企業價值）"); c.font = HDR; c.fill = HFILL
put(RV, RVR["w1"], 2, f'=IFERROR(INDEX(B{GR0}:B{GR1},MATCH(B{RVR["mkt_ev"]},C{GR0}:C{GR1},-1))+(INDEX(C{GR0}:C{GR1},MATCH(B{RVR["mkt_ev"]},C{GR0}:C{GR1},-1))-B{RVR["mkt_ev"]})/(INDEX(C{GR0}:C{GR1},MATCH(B{RVR["mkt_ev"]},C{GR0}:C{GR1},-1))-INDEX(C{GR0}:C{GR1},MATCH(B{RVR["mkt_ev"]},C{GR0}:C{GR1},-1)+1))*(INDEX(B{GR0}:B{GR1},MATCH(B{RVR["mkt_ev"]},C{GR0}:C{GR1},-1)+1)-INDEX(B{GR0}:B{GR1},MATCH(B{RVR["mkt_ev"]},C{GR0}:C{GR1},-1))),"格點範圍內無解")', "0.00%")
c = RV.cell(RVR["ctt"] - 0, 1); 
put(RV, RVR["ctt"], 1, "3. Comparable：現價隱含的 P/S vs 同業 P/S（市值口徑；TTM 營收）", bold=True)
RV.cell(RVR["ctt"], 1).font = HDR; RV.cell(RVR["ctt"], 1).fill = HFILL
rvp("cur_mult", "本公司現行 P/S（現價×稀釋股數 ÷ 最近四季營收 2025Q4–2026Q3）", f"={val('px')}*{val('sh_dil')}/({q('rev.2025Q4')}+{q('rev.2026Q1')}+{q('rev.2026Q2')}+{q('rev.2026Q3')})", "0.0")
rvp("peer_mult", "同業 P/S 中位數（B 情境，按目標日 TTM 營收加權）", f"=({val('m_c')}*{val('ttm_c','C')}+{val('m_g')}*{val('ttm_g','C')})/({val('ttm_c','C')}+{val('ttm_g','C')})", "0.0")
rvp("cur_rel", "本公司現行倍數 ÷ 同業中位數倍數（倍）", f"=B{RVR['cur_mult']}/B{RVR['peer_mult']}", "0.0")
rvp("imp_up", "Comparable 隱含重估：目標價 ÷ 現價−1（情境 B）", f"={val('ps_c','C')}/{val('px')}-1", "+0.0%;-0.0%")
put(RV, RVR["sum_dcf"], 1, f'="現金流折現法：現價隱含的企業價值約為模型情境 B 的 "&TEXT(B{RVR["scale"]},"0.0")&" 倍；若只靠調整折現率，需要 WACC "&IF(ISNUMBER(B{RVR["w1"]}),"約 "&TEXT(B{RVR["w1"]},"0.0%"),"低於格點下限（逼近永續成長率，不可行）")&"（模型 "&TEXT(B{RVR["w0"]},"0.0%")&"）。"')
put(RV, RVR["sum_cmp"], 1, f'="Comparable：本公司現行 P/S 約為同業中位數 P/S 的 "&TEXT(B{RVR["cur_rel"]},"0.0")&" 倍（"&TEXT(B{RVR["cur_mult"]},"0.0")&" 對 "&TEXT(B{RVR["peer_mult"]},"0.0")&"）；成長調整後情境 B 隱含重估 "&TEXT(B{RVR["imp_up"]},"+0%;-0%")&"。"')
RV.column_dimensions["A"].width = 80
for col in "BCD": RV.column_dimensions[col].width = 16

# =====================================================================================================
# Recon
RCs = SH["Recon"]
RCs["A1"] = ("Recon — 口徑對照：（1）合併層對賣方共識（財報前，標逾期）；（2）對公司指引；（3）DCF 與 Comparable 落差的公式分解；（4）主導變數整本重算反解；（5）市場資料日期與逾期標示；"
             "（6）Tokenomics 成本下限對照。共識與指引只作對照，不進模型。")
RCs["A2"] = "未找到：公司未揭露 AIP 任務數、token 用量或推論成本；賣方共識分部拆分未取得；RSU 預期沒收率未揭露；各同業倍數非同一估值日（見第 5 節）。"
RCK = {"s1": 4, "h1": 5, "rev26": 6, "rev27": 7, "eps26": 8, "eps27": 9,
       "s2": 11, "h2": 12, "q4": 13, "fy": 14, "usc": 15,
       "s3": 17, "h3": 18, "d_ps": 19, "d_time": 20, "d_dcf": 21, "d_sbc": 22, "d_lev": 23, "d_mult": 24, "d_comp": 25, "d_chk": 26, "d_sbcw": 27,
       "s4": 29, "h4": 30, "k_wacc": 31, "k_wl": 32, "k_chk": 33,
       "s5": 35, "h5": 36, "dt0": 37,
       "s6": 48, "h6": 49, "fc": 50, "mkt": 51, "ratio": 52, "joint": 53}
def rc_hdr(r, t, heads):
    c = RCs.cell(r, 1, t); c.font = HDR; c.fill = HFILL
    for j, h in enumerate(heads, 1): put(RCs, r + 1, j, h, bold=True)
rc_hdr(RCK["s1"], "1. 合併層：本模型 vs 賣方共識（財報前；FY2026、FY2027）", ["項目", "年度", "情境 A", "情境 B", "情境 C", "共識", "B 對共識", "說明"])
for k, lab, y, col, ck, note in (("rev26", "營收（US$M）", "FY2026E", "F", "cons_rev26", "共識日 2026-10-23（Q3 財報前，逾期）；29 家"), ("rev27", "營收（US$M）", "FY2027E", "G", "cons_rev27", "同上；28 家")):
    r = RCK[k]; put(RCs, r, 1, lab); put(RCs, r, 2, y)
    for j, s in enumerate("ABC"): put(RCs, r, 3 + j, f"=PL!{col}{PLR[s]['rev']}", "#,##0")
    put(RCs, r, 6, f"={I(ck)}", "#,##0"); put(RCs, r, 7, f"=D{r}/F{r}-1", "+0.0%;-0.0%"); put(RCs, r, 8, note)
for k, y, col, ck in (("eps26", "FY2026E", "F", "cons_eps26"), ("eps27", "FY2027E", "G", "cons_eps27")):
    r = RCK[k]; put(RCs, r, 1, "EPS（美元；本模型 GAAP 稀釋，共識多為 non-GAAP，口徑不同，只作方向對照）"); put(RCs, r, 2, y)
    for j, s in enumerate("ABC"): put(RCs, r, 3 + j, f"=PL!{col}{PLR[s]['eps']}", "0.00")
    put(RCs, r, 6, f"={I(ck)}", "0.00"); put(RCs, r, 7, f"=D{r}/F{r}-1", "+0.0%;-0.0%"); put(RCs, r, 8, "口徑不同：GAAP 含股票薪酬，non-GAAP 不含；差距不能解讀為預測差")
rc_hdr(RCK["s2"], "2. 對公司指引（Q3 財報新聞稿 Outlook）", ["項目", "本模型（B）", "指引低", "指引高", "位置", "說明"])
r = RCK["q4"]; put(RCs, r, 1, "2026Q4 營收（US$M）"); put(RCs, r, 2, f"=Seg!D{SB['rev']}", "#,##0"); put(RCs, r, 3, f"={I('g_q4lo')}", "#,##0"); put(RCs, r, 4, f"={I('g_q4hi')}", "#,##0")
put(RCs, r, 5, f'=IF(B{r}<C{r},"低於指引區間",IF(B{r}>D{r},"高於指引區間","區間內"))'); put(RCs, r, 6, "Q4 季增 In 頁判斷值；指引為公司自揭")
r = RCK["fy"]; put(RCs, r, 1, "FY2026 營收（US$M）"); put(RCs, r, 2, f"=Seg!F{SB['rev']}", "#,##0"); put(RCs, r, 3, f"={I('g_fylo')}", "#,##0"); put(RCs, r, 4, f"={I('g_fyhi')}", "#,##0")
put(RCs, r, 5, f'=IF(B{r}<C{r},"低於指引區間",IF(B{r}>D{r},"高於指引區間","區間內"))'); put(RCs, r, 6, "前三季實績＋Q4E")
r = RCK["usc"]; put(RCs, r, 1, "FY2026 美國商業營收（US$M；指引為「超過」下限）"); put(RCs, r, 2, f"=Seg!F{SB['usc']}", "#,##0"); put(RCs, r, 3, f"={I('g_usc')}", "#,##0"); put(RCs, r, 4, None)
put(RCs, r, 5, f'=IF(B{r}>=C{r},"達指引下限","低於指引下限")'); put(RCs, r, 6, "指引為下限")
# 3. DCF 與 Comparable 落差分解（G2-8）
rc_hdr(RCK["s3"], "3. DCF 與 Comparable 落差分解（每股，目標時點；公式分解：時點、股票薪酬、槓桿、倍數與成長）", ["項目", "情境 A", "情境 B", "情境 C", "說明"])
rows3 = {
 "d_ps": ("DCF 每股價值（評價日）", lambda s, c: f"=DCF!C{DR[s]['ps']}", "0.00", "DCF 頁"),
 "d_time": ("＋時點：推到目標時點（×(1＋ke)^年數−1）", lambda s, c: f"=DCF!C{DR[s]['pstp']}-DCF!C{DR[s]['ps']}", "+0.00;-0.00", "Comparable 以目標日 TTM 營收與目標日現金評價，DCF 要推到同一時點才可比"),
 "d_dcf": ("DCF 每股價值（目標時點）", lambda s, c: f"={c}{RCK['d_ps']}+{c}{RCK['d_time']}", "0.00", "＝Val 頁 DCF 每股（目標時點）"),
 "d_sbc": ("＋股票薪酬：DCF 以費用扣除（NOPAT 已扣），P/S 倍數不扣；稅後 SBC 現值（含終值）÷ 股數，推到目標時點", lambda s, c: f"={val('d_pstp_sbc', SCOL[s])}-DCF!C{DR[s]['pstp']}", "+0.00;-0.00", "同業倍數隱含不計 SBC 成本（口徑差）；此列＝DCF（股票薪酬加回）目標時點每股價值減 DCF 每股價值，只作對照"),
 "d_lev": ("槓桿：本公司無借款、同業 EV 為無槓桿口徑，差額為 0", lambda s, c: "=0", "+0.00;-0.00", "營業租賃不計入雙方 EV；本公司借款為 0"),
 "d_mult": ("倍數與成長假設差（殘差＝Comparable − DCF（目標時點）− 股票薪酬 − 槓桿）", lambda s, c: f"={c}{RCK['d_comp']}-{c}{RCK['d_dcf']}-{c}{RCK['d_sbc']}-{c}{RCK['d_lev']}", "+0.00;-0.00", "同業倍數與成長調整 θ 隱含的未來成長與利益率，對比 DCF 明確期路徑"),
 "d_comp": ("Comparable 每股價值（目標時點）", lambda s, c: f"={val('ps_c', SCOL[s])}", "0.00", "Val 頁第 5 節"),
}
for k, (lab, fn, fm, note) in rows3.items():
    r = RCK[k]; put(RCs, r, 1, lab, bold=k in ("d_dcf", "d_comp"))
    for j, s in enumerate("ABC"): put(RCs, r, 2 + j, fn(s, "BCD"[j]), fm)
    put(RCs, r, 5, note)
r = RCK["d_chk"]; put(RCs, r, 1, "檢查：DCF（目標時點）＋SBC＋槓桿＋倍數與成長差 − Comparable（應為 0）")
for j in range(3): cc = "BCD"[j]; put(RCs, r, 2 + j, f"={cc}{RCK['d_dcf']}+{cc}{RCK['d_sbc']}+{cc}{RCK['d_lev']}+{cc}{RCK['d_mult']}-{cc}{RCK['d_comp']}", "0.000")
r = RCK["d_sbcw"]; put(RCs, r, 1, "DCF（股票薪酬加回）每股價值（目標時點；只作對照，不進加權目標價）", bold=True)
for j, s in enumerate("ABC"): put(RCs, r, 2 + j, f"={val('d_pstp_sbc', SCOL[s])}", "0.00")
put(RCs, r, 5, f'="情境加權 DCF（股票薪酬加回）目標價 "&TEXT({val("tpdcf_sbc")},"0.00")&"（上檔空間 "&TEXT({val("updcf_sbc")},"+0.0%;-0.0%")&"）；對照 DCF 目標價 "&TEXT({val("tpdcf")},"0.00")')
# 4. 反解
REV = {}
_rp = os.path.join(HERE, "data", "reverse_solve.json")
if os.path.exists(_rp): REV = json.load(open(_rp, encoding="utf-8"))
rc_hdr(RCK["s4"], "4. 主導變數整本重算反解：使加權目標價等於現價所需的值（逐步整本重算；反解腳本；數值為建置快照）", ["項目", "基準值", "反解值", "反解後目標價", "說明"])
for k, lab, key, fm, note in (("k_wacc", "WACC 調整（加在 WACC；負值＝降低折現率）", "wacc_adj", "0.00%", "使加權目標價＝現價所需的 WACC 調整"), ("k_wl", "FY2027 每客戶工作負載成長（A／B／C 同加 Δ 後的情境 B 值）", "wl27", "0.0%", "三情境同加 Δ；FY2032 收斂終點不動")):
    r = RCK[k]; put(RCs, r, 1, lab)
    put(RCs, r, 2, f"={I('wacc_adj')}" if key == "wacc_adj" else f"={I('wl27_B')}", fm)
    if key in REV and REV[key].get("value") is not None:
        c = put(RCs, r, 3, REV[key]["value"], fm, blue=True); put(RCs, r, 4, REV[key]["tp"], "#,##0.00", blue=True)
    else: put(RCs, r, 3, "未重算"); put(RCs, r, 4, "未重算")
    put(RCs, r, 5, note if key in REV and REV[key].get("value") is not None else (REV.get(key, {}).get("note") or note))
r = RCK["k_chk"]; put(RCs, r, 1, "檢查：反解後目標價 − 現價（應接近 0；單位美元）")
for j, k in enumerate(("k_wacc", "k_wl")): put(RCs, r, 3 + j, f'=IF(ISNUMBER(D{RCK[k]}),D{RCK[k]}-{val("px")},"未重算")', "0.00")
# 5. 市場資料日期與逾期
rc_hdr(RCK["s5"], "5. 市場資料日期與逾期標示（與現價日相差超過 2 個交易日標逾期；原則第 15 條）", ["項目", "資料日期", "與現價日相差（交易日）", "標示", "說明"])
dts = [("無風險利率 rf（10 年美債）", datetime.datetime(2026, 11, 12), "FRED DGS10"), ("賣方平均目標價", datetime.datetime(2026, 8, 31), "MarketScreener；Q3 財報前"),
       ("賣方共識營收與 EPS", datetime.datetime(2026, 10, 23), "Yahoo Finance analysis；Q3 財報前")] + [(f"同業 P/S：{p[1]}（{p[0]}）", datetime.datetime.strptime(PS13[p[0]][0], "%Y-%m-%d"), "peer_ps_20261113.json 價格日") for p in PEERS]
for i, (lab, d, note) in enumerate(dts):
    r = RCK["dt0"] + i; put(RCs, r, 1, lab); put(RCs, r, 2, d, "yyyy-mm-dd", blue=True)
    put(RCs, r, 3, f"=NETWORKDAYS(B{r},{I('px_date')})-1", "0"); put(RCs, r, 4, f'=IF(ABS(C{r})>2,"逾期","")'); put(RCs, r, 5, note)
assert RCK["dt0"] + len(dts) <= RCK["s6"], len(dts)
rc_hdr(RCK["s6"], "6. Tokenomics 成本下限對照與聯動（只讀快照 v22.0；不自行估推論成本）", ["項目", "值", "單位", "說明"])
put(RCs, RCK["fc"], 1, "模型商全成本下限（Tokenomics 全成本，Terra 層，GB200 基準）"); put(RCs, RCK["fc"], 2, TKI["IF_FullCost_Terra"]["values"]["GB200 NVL72"]["基準"], "0.0000", blue=True); put(RCs, RCK["fc"], 3, "US$/M token")
put(RCs, RCK["fc"], 4, "模型商的成本下限；軟體公司付的是市場價（Analogy：選 Terra 層），兩者價差是模型商毛利，不是 Palantir 成本")
put(RCs, RCK["mkt"], 1, "本模型採用市場價（Tokenomics 市場價，Terra 層，基準）"); put(RCs, RCK["mkt"], 2, f"={I('tk_p')}", "0.0000"); put(RCs, RCK["mkt"], 3, "US$/M token"); put(RCs, RCK["mkt"], 4, "選哪一層價格屬 Analogy，不標 Verified")
put(RCs, RCK["ratio"], 1, "市場價 ÷ 全成本下限"); put(RCs, RCK["ratio"], 2, f"=B{RCK['mkt']}/B{RCK['fc']}", "0.0"); put(RCs, RCK["ratio"], 3, "倍"); put(RCs, RCK["ratio"], 4, "Sens 中「只動價格」的高案為成本上界（價格升而營收不變）；另列價格與成本聯動案")
put(RCs, RCK["joint"], 1, "推論成本占 Commercial 營收（FY2028，情境 B）"); put(RCs, RCK["joint"], 2, f"=Seg!H{SEGR['B']['share']}", "0.00%"); put(RCs, RCK["joint"], 3, "%"); put(RCs, RCK["joint"], 4, "公司未揭露任何推論成本或 token 用量（未找到）；此為 In 頁判斷值（每客戶任務數）推得")
RCs.column_dimensions["A"].width = 80
for col in "BCDEFG": RCs.column_dimensions[col].width = 16
RCs.column_dimensions["H"].width = 60

# =====================================================================================================
# Sens
SE = SH["Sens"]
SE["A1"] = "Sens — 一次改一個輸入（In 頁敏感度低／高值），整本活頁簿以 LibreOffice 重算後的加權目標價（Val 頁）與 DCF 單法目標價；由建置腳本（敏感度模式）產生，建置時寫入（數值，屬建置快照）；模型狀態與 build_id 見本頁底部"
SE["A2"] = ("相對基準＝案例目標價 ÷ 基準加權目標價 − 1。只動成本一邊的案例（token 價格高、每任務 token 倍數高、每客戶任務數高）標「上界」：價格升而營收不變，非一致的情境；"
            "另列價格與成本聯動案（token 價格下降同時客戶要求分享節省、單價侵蝕加深）與模型組合案（路由到便宜模型）。")
def hdrrow(ws, r, labs):
    for j, h in enumerate(labs, 1):
        c = ws.cell(r, j, h); c.font = HDR; c.fill = HFILL
hdrrow(SE, 3, ["驅動", "設定", "方向", "加權目標價", "相對基準", "鍵", "說明", "收支關卡 B（五年）", "FY2027 美國商業營收（B）", "DCF 目標價", "DCF 相對基準（美元）", "投資建議"])
COST_ONLY = {"tk_p": "上界：只動成本邊（營收不變）", "tk_mult": "上界：只動成本邊（營收不變）", "tk_n": "上界：只動成本邊（營收不變）"}
_rc = _rg2 = RELM["all"]["sens"]
JOINT = {
    "rel_c+rel_g:ps_only": ("相對倍數只用 PS 讀數（整體 json ps_only）", "低", {"rel_c": round(_rc["ps_only"], 4), "rel_g": round(_rg2["ps_only"], 4)}, "PS 讀數不計成長"),
    "rel_c+rel_g:psg_only": ("相對倍數只用 PSG 讀數（整體 json psg_only）", "高", {"rel_c": round(_rc["psg_only"], 4), "rel_g": round(_rg2["psg_only"], 4)}, "PSG 讀數把成長放進分母，高成長公司被放大"),
    "rel_c+rel_g:sens.low": ("相對倍數取共用庫敏感度低值（sens.low）", "低", {"rel_c": round(_rc["low"], 4), "rel_g": round(_rg2["low"], 4)}, "共用庫 sens.low"),
    "rel_c+rel_g:sens.high": ("相對倍數取共用庫敏感度高值（sens.high）", "高", {"rel_c": round(_rc["high"], 4), "rel_g": round(_rg2["high"], 4)}, "共用庫 sens.high"),
    "pm_all:db_last_q": ("同業 P/S＝資料庫最近一季中位（各同業 <同業>.csv 最近一季值；取代現價日口徑）", "高", {f"pm_{k.lower()}": PSL[k][1] for k in PSL}, "資料庫最近一季，季末較舊（2026-06-30 至 09-30）；較現價日口徑偏高"),
    "tk_pg+pe27": ("價格與成本聯動：token 價格每年降 30% 且客戶要求分享節省（FY2027 單價變動三情境各再 −5pt）", "高", {"tk_pg": -0.30, "pe27_A": -0.10, "pe27_B": -0.13, "pe27_C": -0.18}, "價格降使推論成本降，但單價侵蝕加深（聯動）"),
    "tk_p+tk_n": ("模型組合路由到便宜模型：token 價格取 Luna 層、同時每客戶任務數加倍（便宜模型使任務變多）", "低", {"tk_p": float(TKI["IF_PriceRef_Luna"]["values"]), "tk_n": 10}, "價格降、用量升（聯動）"),
}
SENS_CASES = []   # (key, label, side, note)
for it in INS:
    if it[0] == "beta":
        SENS_CASES += [("beta", "β（Blume 調整，兩源）", "低", "取兩個來源調整值的較低者"), ("beta", "β（Blume 調整，兩源）", "高", "取兩個來源調整值的較高者")]
    elif it[0] == "row":
        d = it[1]
        if d["lo"] is None and d["hi"] is None: continue
        if d["key"] in ("px_date", "rel_c", "rel_g"): continue
        short = d["label"].split("（")[0][:30]
        for side, v in (("低", d["lo"]), ("高", d["hi"])):
            if v is None: continue
            SENS_CASES.append((d["key"], short, side, COST_ONLY.get(d["key"], "In 頁敏感度" + side + "值")))
for jk, (lab, side, sets, note) in JOINT.items(): SENS_CASES.append((jk, lab, side, note))
r = 4; SE_ROWS = []
for key, lab, side, note in SENS_CASES:
    put(SE, r, 1, f"{lab}：{side}"); put(SE, r, 2, note); put(SE, r, 3, side); put(SE, r, 6, key)
    res = SENS.get(f"{key}|{side}")
    if res is not None:
        put(SE, r, 4, res["tp"], "#,##0.00", blue=True); put(SE, r, 5, f"=D{r}/Val!$B${VR['tp']}-1", "+0.0%;-0.0%;0.0%")
        put(SE, r, 8, res["gateB"], "#,##0.0", blue=True); put(SE, r, 9, res["usc27"], "#,##0.0", blue=True)
        put(SE, r, 10, res["tpdcf"], "#,##0.00", blue=True); put(SE, r, 11, f"=J{r}-Val!$B${VR['tpdcf']}", "+#,##0.00;-#,##0.00")
        put(SE, r, 12, "=" + recf(f"D{r}/{I('px')}-1"))
    put(SE, r, 7, "重算值（make_v01.sh）")
    SE_ROWS.append((key, side, lab, r)); r += 1
SE_END = r - 1
r += 1
put(SE, r, 1, "基準加權目標價（Val 頁）"); put(SE, r, 4, f"=Val!B{VR['tp']}", "#,##0.00"); SE_BASE = r
put(SE, r, 8, f"=CF!C{GR['gate']}", "#,##0.0"); put(SE, r, 9, f"=Seg!G{SEGR['B']['usc']}", "#,##0.0"); r += 1
put(SE, r, 1, "敏感度重算時的基準值（應與上一列相同；不同表示敏感度已過期）")
if "base|base" in SENS:
    put(SE, r, 4, SENS["base|base"]["tp"], "#,##0.00", blue=True); put(SE, r, 5, f"=D{r}-D{r - 1}", "0.000")
SE_STALE = r; r += 1
FP_RNG = f"In!$D$4:$F${IN_LAST}"
put(SE, r, 1, "輸入指紋（本檔即時重算）"); put(SE, r, 4, f"=SUMPRODUCT({FP_RNG},ROW({FP_RNG})*3+COLUMN({FP_RNG}))", "0.000000"); SE_FP = r; r += 1
put(SE, r, 1, "輸入指紋（敏感度重算時的基準檔）")
if "base|base" in SENS and SENS["base|base"].get("fp") is not None: put(SE, r, 4, SENS["base|base"]["fp"], "0.000000", blue=True)
r += 1
put(SE, r, 1, "模型狀態（CURRENT＝Sens、HTML、README 與本檔輸入一致；STALE＝In 頁已改、須重跑敏感度）", bold=True)
put(SE, r, 4, f'=IF(AND(ISNUMBER(D{r - 1}),ISNUMBER(D{SE_STALE}),ABS(D{SE_STALE}-D{SE_BASE})<0.005,ABS(D{r - 2}-D{r - 1})<=0.000000001*MAX(1,ABS(D{r - 1}))),"CURRENT","STALE")')
SE_STATUS = r; r += 1
put(SE, r, 1, "build_id（README、Sens、HTML 三處相同）"); put(SE, r, 4, BUILD_ID); SE_BID = r; r += 1
r += 1
put(SE, r, 1, "加權目標價相對基準（HTML 圖用：低案、高案；每個驅動一列）", bold=True); SE_REL0 = r; r += 1
for j, h in enumerate(["驅動", "低案相對基準", "高案相對基準"], 1): put(SE, r, j, h, bold=True)
r += 1
SE_REL1 = r
drv_order = []
for key, side, lab, rr in SE_ROWS:
    if (key, lab) not in [(a, b) for a, b, _ in drv_order]: drv_order.append((key, lab, {}))
    for a, b, dct in drv_order:
        if a == key and b == lab: dct[side] = rr
for key, lab, dct in drv_order:
    put(SE, r, 1, lab)
    if "低" in dct: put(SE, r, 2, f"=E{dct['低']}", "+0.0%;-0.0%;0.0%")
    if "高" in dct: put(SE, r, 3, f"=E{dct['高']}", "+0.0%;-0.0%;0.0%")
    r += 1
SE_REL2 = r - 1
SE.column_dimensions["A"].width = 60; SE.column_dimensions["B"].width = 40; SE.column_dimensions["G"].width = 22

# =====================================================================================================
# Sources
SO = SH["Sources"]
SO["A1"] = "Sources — 來源（編號、文件、日期、網址、標記、用途）；逐筆事實與頁碼見 data/facts_palantir.csv 與 Q 頁"
hdrrow(SO, 3, ["編號", "文件", "日期", "網址", "標記", "用途"])
SRCS = [
    ("Palantir 2026Q3 財報新聞稿（8-K 附件 99.1）", "2026-11-03", "https://investors.palantir.com/", "Verified", "2026Q3 分部營收、美國商業營收、客戶數、指引、Outlook"),
    ("Palantir 2026Q3 10-Q", "2026-11（封面股數 2026-10-26）", "https://www.sec.gov/cgi-bin/browse-edgar?action=getcompany&CIK=0001321655&type=10-Q", "Verified", "損益表、現金流量表、資產負債表、分部資訊（Note）、股數、選擇權／RSU／SAR"),
    ("Palantir FY2025、FY2024、FY2023 10-K", "2026-02／2025-02／2024-02", "https://www.sec.gov/cgi-bin/browse-edgar?action=getcompany&CIK=0001321655&type=10-K", "Verified", "歷史分部營收與分部利益、損益、現金流量"),
    ("StockAnalysis PLTR 統計頁與歷史價格", "2026-11-14", "https://stockanalysis.com/stocks/pltr/statistics/", "Verified", "現價 167.99（2026-11-13 收盤）、β、市值、EV／營收"),
    ("Yahoo Finance PLTR 統計與 analysis 頁", "2026-10-23", "https://finance.yahoo.com/quote/PLTR/analysis/", "Interested-party", "β（5Y Monthly）；賣方 FY2026／FY2027 營收與 EPS 共識（Q3 財報前）"),
    ("MarketScreener PLTR 共識頁", "2026-08-31", "https://ca.marketscreener.com/quote/stock/PALANTIR-TECHNOLOGIES-INC-113108869/consensus/", "Interested-party", "賣方平均目標價（32 家；財報前）"),
    ("共用庫同業 P/S（software_ai/data/relmult_hist/out/peer_ps_20261113.json all 群組，現價日 2026-11-13 口徑；資料庫最近一季值見 <同業>.csv，列為 Sens 案；SNOW、DDOG、NET、MDB、NOW、BAH、LDOS、CACI）與 StockAnalysis 同業頁（賣方 3 年營收成長預測）", "價格日 2026-10-26 至 2026-11-13", "repo 內／https://stockanalysis.com/", "Derived", "同業 P/S（市值÷TTM 營收）；賣方 3 年營收成長預測（Interested-party）"),
    ("FRED DGS10（10 年期美國公債殖利率）", "2026-11-12", "https://fred.stlouisfed.org/series/DGS10", "Verified", "rf 5.10%（工作單 G5 統一取值；現價日前 1 個交易日）"),
    ("共用庫相對倍數 software_ai/data/relmult_hist/out/palantir.json", "2026-11-17", "repo 內（software_ai/data/relmult_hist）", "Derived", "Comparable 相對倍數 geo_base（商業、政府分組）與 PS／PSG 讀數、sens 四案"),
    ("Futurum 與其他法說轉述（Q3 美國商業客戶數、NDR、每客戶營收等）", "2026-11", "見 data/facts_palantir.csv 各筆 url", "Interested-party", "美國商業客戶數、淨美元留存、客戶層級 KPI（公司自揭或轉述）"),
    (f"Tokenomics 快照 data/tokenomics_snapshot_{TKV}.json（{TK['source']['file']}，commit {TK['source']['commit'][:7]}）", TK["source"]["extractedAt"], "repo 內（YinchenChang/Tokenomics 唯讀 clone）", "Analogy", "IF_PriceRef_Luna／Terra／Sol、IF_TaskTok*、IF_FullCost_Terra（token 價格層級選擇為類比）"),
    ("事實總帳 data/facts_palantir.csv（facts_palantir.py）", "2026-11-16", "repo 內", "各筆標記", "Q 頁全部實績與多數輸入"),
]
for i, (doc, dt, url, tag, use) in enumerate(SRCS, 1):
    for j, v in enumerate((f"S{i:02d}", doc, dt, url, tag, use), 1): put(SO, 3 + i, j, v)
for col, w in zip("ABCDEF", [6, 80, 22, 70, 18, 60]): SO.column_dimensions[col].width = w

# =====================================================================================================
# README（文字中的數字一律公式組字；原則第 13、16、31、41 條）
RD = SH["README"]
VB = lambda k, c="B": f"Val!{c}{VR[k]}"
SB_ = SEGR["B"]; SA_ = SEGR["A"]; SC_ = SEGR["C"]
TXT = lambda ref, f: f'TEXT({ref},"{f}")'
def bn(sh, col, row): return f'TEXT({sh}!{col}{row}/1000,"0.0")'
RDR = {}
rr = 1
def rput(key, formula, bold=False):
    global rr
    RDR[key] = rr; c = RD.cell(rr, 1, formula); c.alignment = Alignment(wrap_text=True, vertical="top")
    if bold: c.font = BOLD
    rr += 1
rput("hdr", (f'="Palantir（PLTR）個股模型 {VER}｜build_id {BUILD_ID}｜資料截止日 "&TEXT({I("data_date")},"yyyy-mm-dd")&"｜成品日 {VDATE}｜股價基準日 "&TEXT({I("px_date")},"yyyy-mm-dd")'
             f'&"｜評價日 "&TEXT({I("val_date")},"yyyy-mm-dd")&"｜目標時點 "&TEXT({I("tp_date")},"yyyy-mm-dd")&"（命題＝原型 U 按用量計價的單位經濟＋官方分部 → 損益表；本版未做先驗）"'))
rput("sheets", "工作表：README → In → Q → Seg → PL → CF → DCF → Val → RevDCF → Recon → Sens → Sources。藍字＝輸入、黑字＝公式；黃底＝判斷題。命題模組變數（客戶數、工作負載、單價、推論成本）暫建於 In 頁，標「v0.2 改讀 unit_economics」。")
rput("thesis", (f'="命題：AI 對 Palantir 的價值重分配＝美國商業營收（客戶數 × 每客戶年營收）中 AI 增量減單價侵蝕再減推論成本的淨效果；情境 B FY2028 AI 增量 US$"&{bn("Seg","H",SB_["inc"])}&"B、單價侵蝕 US$"&{bn("Seg","H",SB_["ero"])}&"B、增量推論成本 US$"&TEXT(Seg!H{SB_["incinf"]}/1000,"0.00")&"B，淨效果 US$"&{bn("Seg","H",SB_["net"])}&"B。三情境的差別全在這三個分量與政府成長。"'))
rr += 1
rput("rec", (f'="投資建議："&{VB("rec")}&"（估值評等 "&{VB("est")}&"；收支關卡 "&CF!C{GR["gate_ok"]}&"）；加權目標價 US$"&TEXT({VB("tp")},"0.00")&"（DCF "&TEXT({VB("tpdcf")},"0.00")&"、Comparable "&TEXT({VB("tpcomp")},"0.00")&"），現價 US$"&TEXT({VB("px")},"0.00")&"（"&TEXT({I("px_date")},"yyyy-mm-dd")&"），上檔空間 "&TEXT({VB("up")},"+0.0%;-0.0%")&"。"'), True)
O = {}
O["o1"] = (f'="重點觀察 1：目標價 "&TEXT({VB("tp")},"0.00")&"（DCF "&TEXT({VB("tpdcf")},"0.00")&"、Comparable "&TEXT({VB("tpcomp")},"0.00")&"），現價 "&TEXT({VB("px")},"0.00")&"；兩法都遠低於現價，落差由規則算出，不是調參結果。"')
O["o2"] = (f'="重點觀察 2：情境 B FY2028 AI 淨效果 US$"&{bn("Seg","H",SB_["net"])}&"B；情境 A "&{bn("Seg","H",SA_["net"])}&"B、情境 C "&{bn("Seg","H",SC_["net"])}&"B。"')
O["o3"] = (f'="重點觀察 3：推論成本只占 Commercial 營收 "&TEXT(Seg!H{SB_["share"]},"0.0%")&"（情境 B FY2028；token 用量公司未揭露，為判斷值）；損益兩平的工作負載成長僅 "&TEXT(Seg!H{SB_["be"]},"0.0%")&"。"')
O["o4"] = (f'="重點觀察 4：現價隱含的企業價值約為模型情境 B 的 "&TEXT(RevDCF!B{RVR["scale"]},"0.0")&" 倍；只調折現率則需 WACC "&IF(ISNUMBER(RevDCF!B{RVR["w1"]}),"約 "&TEXT(RevDCF!B{RVR["w1"]},"0.0%"),"低於格點下限（逼近永續成長率，不可行）")&"。"')
O["o5"] = (f'="重點觀察 5：本公司現行 P/S 約為同業中位數的 "&TEXT(RevDCF!B{RVR["cur_rel"]},"0.0")&" 倍；政府營收約占 FY2026 營收 "&TEXT(Seg!F{SB_["gov"]}/Seg!F{SB_["rev"]},"0%")&"。"')
rput("o1_legacy", "=" + O["o1"][1:]); rput("o2_legacy", "=" + O["o2"][1:]); rput("o3_legacy", "=" + O["o3"][1:])
rr += 1
rput("v13", "v1.3 展示層（HTML 讀取；本區文字數字全由公式組字）", True)
rput("hero", (f'="加權目標價 US$"&TEXT({VB("tp")},"0.00")&"，現價 US$"&TEXT({VB("px")},"0.00")&"（上檔空間 "&TEXT({VB("up")},"+0%;-0%")&"）。情境 B AI 淨效果 FY2028 US$"&{bn("Seg","H",SB_["net"])}&"B；"&IF(AND({VB("tpdcf")}<{VB("px")},{VB("tpcomp")}<{VB("px")}),"DCF 與 Comparable 兩法都低於現價。",IF(AND({VB("tpdcf")}>={VB("px")},{VB("tpcomp")}>={VB("px")}),"DCF 與 Comparable 兩法都高於現價。","兩法一高一低。"))'))
rput("rec2", "=A" + str(RDR["rec"]))
for k in ("o1", "o2", "o3", "o4", "o5"): rput(k, "=" + O[k][1:])
rput("drv", (f'="分部營收推導：美國商業營收＝客戶數 × 每客戶年營收，每客戶年營收成長＝代理工作負載成長＋單價侵蝕；國際商業與政府分部另依成長路徑；兩個官方分部加總＝合併營收（FY2026E 為前三季實績加 Q4E）。"'))
rput("ver", f'="{VER}（2026-11-17）：同業中位 P/S 改讀 peer_ps_20261113.json（G7，現價日口徑），Sens 加「同業 P/S＝資料庫最近一季中位」案；v0.2.1：Comparable 改整體單一口徑（8 家合併中位 P/S × 整體 geo_base，分組讀數只列對照）；Comparable 主腿改 P/S（G6，同業 P/S 取共用庫最近一季）；Comparable 相對倍數改讀共用庫（θ 成長調整改為 Sens）；rf 統一 5.10%（FRED DGS10，2026-11-12）；新增 DCF（股票薪酬加回）對照讀數；Sens 加 ps_only、psg_only、sens.low、sens.high 四案；合理性關卡重跑。"')
rput("ver1", "v0.1（2026-11-16）：首版：研究包與個股模型；Government／Commercial 分部、原型 U 單位經濟（v0.2 改讀 unit_economics）、DCF 至 FY2032、Comparable＝同業 EV／營收成長調整、五級評等與收支關卡。")
rr += 1
rput("rel0", "相對倍數（共用庫）與 PS／PSG 差異原因", True)
rput("rel1", (f'="Comparable 主腿（v0.2.2）：8 家同業合併中位 P/S（peer_ps_20261113.json all 群組，現價日 2026-11-13 口徑；價格日非 11-13 者標逾期，Recon 第 5 節；資料庫最近一季中位另列 Sens 案）"&TEXT({val("m_c")},"0.00")&" × 整體相對倍數 "&TEXT({I("rel_c")},"0.00")&"（palantir.json groups.all.geo_base，Derived；PS "&TEXT({I("ps_a")},"0.00")&"、PSG "&TEXT({I("psg_a")},"0.00")&" 的幾何平均）× 合併目標時點 TTM 營收 ÷ 稀釋股數。改整體單一口徑的理由：分組讀數是以本公司合併 P/S 除以分組同業（政府 IT 服務：BAH、LDOS、CACI），政府分部等於按 Palantir 自己的歷史 P/S 計價，且政府同業與 Palantir 業務不同類，比值被放大成雜訊（政府 geo_base "&TEXT({I("rel_g_ref")},"0.00")&"、商業 "&TEXT({I("rel_c_ref")},"0.00")&"），故商業、政府分組讀數只列對照、不進加權目標價。主腿為 P/S（G6），與 PS／PSG 同口徑；EV／營收與盈餘型倍數只作對照，不進加權。"'))
rput("rel2", (f'="PS 與 PSG 兩讀數差異大（整體：PS "&TEXT({I("ps_a")},"0.00")&"、PSG "&TEXT({I("psg_a")},"0.00")&"，差 "&TEXT({I("dv_a")},"0.00")&" 倍，超過 30%）。原因是成長比：本公司 TTM 營收年增 "&TEXT({I("cgr_a")},"0%")&"、同業中位 "&TEXT({I("pgr_a")},"0%")&"，成長比 "&TEXT({I("gr_a")},"0.00")&"。PS 不計成長，PSG 把成長放進分母，高成長公司的 PSG 相對值因此被放大；兩者取幾何平均為基準，單獨取 PS 或 PSG 的結果列在 Sens。"'))
rr += 1
rput("qa", "命題回覆（依模型計算逐一回答）", True)
rput("qa1", (f'="問：AI 讓 Palantir 的單位經濟變好還是變差？答：情境 B FY2028 AI 增量 US$"&{bn("Seg","H",SB_["inc"])}&"B 大於單價侵蝕 US$"&{bn("Seg","H",SB_["ero"])}&"B 與增量推論成本 US$"&TEXT(Seg!H{SB_["incinf"]}/1000,"0.00")&"B 之和，淨效果 US$"&{bn("Seg","H",SB_["net"])}&"B；情境 C（工作負載成長 "&TEXT(In!E{IR["wl27_C"]},"0%")&"、單價 "&TEXT(In!E{IR["pe27_C"]},"0%")&"）淨效果降到 US$"&{bn("Seg","H",SC_["net"])}&"B。"'))
rput("qa2", (f'="問：推論成本會不會吃掉增量？答：不會，在本模型假設下：推論成本占 Commercial 營收約 "&TEXT(Seg!H{SB_["share"]},"0.0%")&"，且每客戶任務數為判斷值（公司未揭露）；token 價格上界案（Tokenomics 高層價格）見 Sens。真正的風險在單價侵蝕，不在推論成本。"'))
rput("qa3", (f'="問：現價隱含什麼？答：以情境 B 現金流，現價隱含企業價值約為模型的 "&TEXT(RevDCF!B{RVR["scale"]},"0.0")&" 倍；DCF 目標價 US$"&TEXT({VB("tpdcf")},"0.00")&" 與 Comparable US$"&TEXT({VB("tpcomp")},"0.00")&" 的落差，股票薪酬解釋 US$"&TEXT(Recon!C{RCK["d_sbc"]},"0.00")&"，其餘為倍數與成長假設。"'))
rr += 1
rput("gate0", "合理性關卡與診斷結論（工作單 G3；發布前）", True)
rput("gate1", f'="DCF 目標價 "&TEXT({VB("tpdcf")},"0.00")&" ÷ 現價 "&TEXT({VB("px")},"0.00")&" ＝ "&TEXT({VB("tpdcf")}/{VB("px")},"0.00")&" 倍："&IF({VB("tpdcf")}/{VB("px")}<0.25,"低於 0.25 倍，觸發",IF({VB("tpdcf")}/{VB("px")}>4,"高於 4 倍，觸發","落在 0.25–4 倍內"))&"。"')
rput("gate2", f'="企業價值 EV（情境 A／B／C，US$B）"&TEXT({VB("d_ev","B")}/1000,"0.0")&"／"&TEXT({VB("d_ev","C")}/1000,"0.0")&"／"&TEXT({VB("d_ev","D")}/1000,"0.0")&"："&IF(MIN({VB("d_ev","B")},{VB("d_ev","C")},{VB("d_ev","D")})<0,"有負值，觸發","皆為正，未觸發")&"。評等 "&{VB("est")}&"："&IF(OR({VB("est")}="強力放空",{VB("est")}="強力買進"),"強力評等，觸發","未觸發")&"。"')
rput("gate3", f'="FY2026E 營收 US$"&TEXT(Seg!F{SB_["rev"]},"#,##0")&"M 對公司指引中點 US$"&TEXT(({I("g_fylo")}+{I("g_fyhi")})/2,"#,##0")&"M 差 "&TEXT(Seg!F{SB_["rev"]}/(({I("g_fylo")}+{I("g_fyhi")})/2)-1,"+0.0%;-0.0%")&"；FY2027E 營收對賣方共識（財報前）差 "&TEXT(PL!G{PLR["B"]["rev"]}/{I("cons_rev27")}-1,"+0.0%;-0.0%")&"；公司未給 FY2027 指引（未找到），營業利益率無指引可對。差距皆在 10% 內。"')
rput("gate4", f'="診斷：現價隱含 EV 約 US$"&TEXT(RevDCF!B{RVR["mkt_ev"]}/1000,"0")&"B，為 FY2026E 自由現金流 US$"&TEXT(CF!F{CFR["B"]["fcf"]}/1000,"0.0")&"B 的 "&TEXT(RevDCF!B{RVR["mkt_ev"]}/CF!F{CFR["B"]["fcf"]},"0")&" 倍，為情境 B 的 DCF EV 的 "&TEXT(RevDCF!B{RVR["scale"]},"0.0")&" 倍。造成落差的輸入：（一）FY2026 起營業利益率與自由現金流已取前三季實績（營業利益率 "&TEXT(PL!M{PLR["B"]["oi"]}/PL!M{PLR["B"]["rev"]},"0%")&"），未凍結在虧損水準；未分攤成本率歷史下降但前三季已是歷史最低，延伸案（In 頁未分攤成本率變動低值）放 Sens；（二）折現率 "&TEXT({VB("wacc")},"0.0%")&"（β "&TEXT({VB("beta")},"0.00")&"）與終值占 EV "&TEXT({VB("d_tvsh","C")},"0%")&"；（三）營收成長自 FY2027 起逐年收斂，情境 A 營收 FY2028 US$"&TEXT(PL!H{PLR["A"]["rev"]}/1000,"0.0")&"B 仍不足以支撐現價。結論：不是建模錯誤（歷史與前三季各核對列為零、營收對指引與共識差距在 10% 內），而是經濟判斷：現價需要 DCF 路徑之外的更高成長或更低折現率，反解見 Recon 第 4 節；評等為規則算出（"&{VB("rec")}&"）；DCF 與 Comparable 兩腿差距大（Comparable 以整體單一口徑：8 家同業合併中位 P/S × 整體相對倍數 "&TEXT({I("rel_c")},"0.00")&"，見 README），且上檔空間距評等邊界僅 "&TEXT({VB("bdist")},"0.0")&" 個百分點，評等對輸入敏感，保留並於 Sens 四案與反解揭露。"')
rr += 1
rput("dates", "市場資料日期", True)
rput("dt1", f'="現價 "&TEXT({I("px")},"0.00")&" 為 "&TEXT({I("px_date")},"yyyy-mm-dd")&" 收盤；評價日 "&TEXT({I("val_date")},"yyyy-mm-dd")&"；目標時點 "&TEXT({I("tp_date")},"yyyy-mm-dd")&"（股價基準日起 12 個月）。"')
rput("dt2", f'="rf 5.10%（2026-11-12，FRED DGS10，未逾期）、賣方共識（2026-08-31 與 2026-10-23，Q3 財報前）與同業倍數（2026-10-26 至 2026-11-15）與現價日相差超過 2 個交易日者標「逾期」（Recon 第 5 節逐筆列示）；同業倍數非同一估值日，限制已揭露。"')
rr += 1
rput("fail", "主要失效條件", True)
rput("f1", f'="工作負載成長低於損益兩平 "&TEXT(Seg!H{SB_["be"]},"0.0%")&"（情境 B FY2028；單價侵蝕下 AI 淨效果轉負）。可證偽：公司揭露美國商業每客戶營收年增與留存率。"')
rput("f2", f'="美國政府營收一次性減少（In 頁尾端事件，FY2028 基準 0；Sens 取 −30%）：政府約占 FY2026 營收 "&TEXT(Seg!F{SB_["gov"]}/Seg!F{SB_["rev"]},"0%")&"。可證偽：政府分部 TCV 與合約續約。"')
rput("f3", f'="單價侵蝕深於情境 C（FY2027 單價變動 "&TEXT(In!E{IR["pe27_C"]},"0%")&"）：Sovereign AI 以開源權重模型自建使每單位工作負載收費下降。可證偽：新簽合約的單位價格或 Commercial 毛利變化。"')
rr += 1
rput("risk", "風險登錄", True)
rput("risk_h", "風險")
RD.cell(RDR["risk_h"], 2, "是否建模"); RD.cell(RDR["risk_h"], 3, "理由"); RD.cell(RDR["risk_h"], 4, "觀察指標")
RISKS = [("美國政府集中度", "部分建模", "尾端事件列在 In 頁基準 0、只進 Sens；政府成長三情境分化", "政府分部營收與 TCV、續約"),
         ("單價侵蝕（開源權重、Sovereign AI 自建）", "建模", "單價變動為情境驅動之一（三情境 −5%／−8%／−13%）", "新簽合約單價、Commercial 分部毛利"),
         ("推論成本與 token 價格", "建模", "只讀 Tokenomics 快照；用量公司未揭露，為判斷值", "公司揭露 AIP 用量或推論成本"),
         ("股票薪酬與稀釋", "建模", "GAAP 營業利益已含 SBC；股數固定為評價日稀釋股數，未計未來回購與新股", "稀釋股數、SBC 占營收"),
         ("估值倍數壓縮", "建模", "Comparable 以同業 P/S 中位數 × 共用庫相對倍數；四案與 θ 列入 Sens", "同業 P/S")]
for i, (a, b, c_, d) in enumerate(RISKS):
    rrr = rr; RDR[f"risk{i}"] = rrr
    for j, v in enumerate((a, b, c_, d), 1): RD.cell(rrr, j, v)
    rr += 1
rr += 1
rput("na", "不適用項目（本模型不做）", True)
rput("na1", "不適用：Coord、三層基準協調、Wattonomics、自用算力（輕資產：資本支出占營收不到 1%，無自建算力）。先驗未做（工作單 G 節：先建模型，先驗後補）。")
rr += 1
rput("ld", "承重 Assumed 與可證偽指標", True)
rput("ld1", f'="每客戶工作負載成長（情境 A／B／C FY2027 "&TEXT(In!E{IR["wl27_A"]},"0%")&"／"&TEXT(In!E{IR["wl27_B"]},"0%")&"／"&TEXT(In!E{IR["wl27_C"]},"0%")&"）：可證偽指標＝公司揭露美國商業每客戶營收年增與留存率（客戶層級資料）。"')
rput("ld2", f'="美國商業客戶數年增（FY2027 "&TEXT(In!E{IR["cg27_A"]},"0%")&"／"&TEXT(In!E{IR["cg27_B"]},"0%")&"／"&TEXT(In!E{IR["cg27_C"]},"0%")&"）：可證偽指標＝季度客戶數揭露。"')
rput("ld3", f'="單價變動（FY2027 "&TEXT(In!E{IR["pe27_A"]},"0%")&"／"&TEXT(In!E{IR["pe27_B"]},"0%")&"／"&TEXT(In!E{IR["pe27_C"]},"0%")&"）：可證偽指標＝新簽合約單位價格與 Commercial 毛利變化。"')
rr += 1
rput("sc", "情境分化依據", True)
rput("sc1", f'="情境差異只放在 AI 淨效果的三個分量與政府成長：A＝客戶擴張與工作負載成長維持、單價侵蝕小；B＝逐季放緩；C＝客戶趨飽和、工作負載成長減半、Sovereign AI 自建使單價侵蝕深。FY2028 AI 淨效果 A／B／C＝US$"&{bn("Seg","H",SA_["net"])}&"B／"&{bn("Seg","H",SB_["net"])}&"B／"&{bn("Seg","H",SC_["net"])}&"B。基礎：公司 Q3 揭露的客戶數年增、每客戶營收年增與淨留存率；分化幅度為判斷（Assumed）。"')
RD.column_dimensions["A"].width = 150
for col in "BCD": RD.column_dimensions[col].width = 30

# =====================================================================================================
# 存檔與列號對照
for ws in wb.worksheets: ws.sheet_view.zoomScale = 90
wb.save(A.out)
if os.environ.get("ROWMAP"):
    json.dump({"VR": VR, "SEGR": SEGR, "SB": SB, "SCM": SCM, "PLR": PLR, "PLQ": PLQ, "PLS": PLS, "CFR": CFR, "GR": GR, "DR": DR, "RVR": RVR, "RCK": RCK, "IR": IR, "IN_ROW": IR, "IN_LAST": IN_LAST,
               "SE_ROWS": SE_ROWS, "SE_BASE": SE_BASE, "SE_STALE": SE_STALE, "SE_STATUS": SE_STATUS, "SE_BID": SE_BID, "SE_FP": SE_FP, "SE_REL0": SE_REL0, "SE_REL1": SE_REL1, "SE_REL2": SE_REL2,
               "RDR": RDR, "JOINT": {k: v[2] for k, v in JOINT.items()}, "SCN_H": SCN_H, "SCN_ROWS": SCN_ROWS, "QCHK": QCHK, "QCHK2": QCHK2, "BUILD_ID": BUILD_ID}, open(os.environ["ROWMAP"], "w", encoding="utf-8"), ensure_ascii=False, indent=1)
print("built", A.out, BUILD_ID)
