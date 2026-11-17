# 趨勢科技（4704）個股模型 v0.2（原型 P 按防護範圍計價；官方報告分部＝所在地別四分部；命題子列＝Vision One）
# 用法：python3 build_trendmicro.py raw.xlsx [--sens sens.json] [--set 鍵=值 ...]；再以 LibreOffice 重算（make_v02.sh）。
# 金額：十億日圓（¥B）；ARR：百萬美元；股數：百萬股；每股：日圓。FY＝日曆年（12 月制）。評價日 2026-06-30（最近可得資產負債表）。
# 本版未做先驗（工作單 G 節）。Meta 的算力、Tokenomics 自用 AI 讀法 2、Wattonomics、三層協調：不適用（見 README）。
import sys, os, json, argparse, hashlib, datetime
import openpyxl
from openpyxl.styles import Font, PatternFill
from openpyxl.workbook.defined_name import DefinedName
HERE = os.path.dirname(os.path.abspath(__file__))
ap = argparse.ArgumentParser(); ap.add_argument("out"); ap.add_argument("--sens"); ap.add_argument("--set", action="append", default=[])
A = ap.parse_args()
OVR = {k: float(v) for k, v in (s.split("=", 1) for s in A.set)}
BLUE = Font(color="0000FF"); BOLD = Font(bold=True); HDR = PatternFill("solid", fgColor="DDEBF7")
VER = "v0.2.1"; VDATE = "2026-11-17"
def _bid():
    h = hashlib.sha256()
    for fn in ["build_trendmicro.py", "data/facts_trendmicro.json", "data/tk_snapshot.json", "../data/relmult_hist/out/trendmicro.json", "../data/relmult_hist/out/peer_ps_20261113.json"]:
        h.update(fn.encode()); h.update(open(os.path.join(HERE, fn), "rb").read())
    return f"trendmicro-{VER}-{h.hexdigest()[:10]}"
BUILD_ID = _bid()
FJ = {x["id"]: x for x in json.load(open(os.path.join(HERE, "data/facts_trendmicro.json"), encoding="utf-8"))}
TK = json.load(open(os.path.join(HERE, "data/tk_snapshot.json"), encoding="utf-8"))
RM = json.load(open(os.path.join(HERE, "../data/relmult_hist/out/trendmicro.json"), encoding="utf-8"))["groups"]["all"]; RMS = RM["sens"]
PP = json.load(open(os.path.join(HERE, "../data/relmult_hist/out/peer_ps_20261113.json"), encoding="utf-8"))["companies"]["trendmicro"]; PPG = PP["groups"]["all"]; PPP = PPG["peers"]
TKV = TK["source"]["version"]; TKI = TK["items"]
def fv(i): return FJ[i]["value"]
def fsrc(i): return FJ[i]["source"].split("；")[0][:160] + f"（事實總帳 {i}）"
wb = openpyxl.Workbook()
RD = wb.active; RD.title = "README"
IN = wb.create_sheet("In"); HI = wb.create_sheet("Hist"); PL = wb.create_sheet("PL"); CF = wb.create_sheet("CF")
VA = wb.create_sheet("Val"); RC = wb.create_sheet("Recon"); RV = wb.create_sheet("RevDCF"); SE = wb.create_sheet("Sens"); SO = wb.create_sheet("Sources")
def dn(name, ref): wb.defined_names[name] = DefinedName(name, attr_text=ref)
def hdr(ws, row, labels, c0=1):
    for j, h in enumerate(labels, c0):
        c = ws.cell(row, j, h); c.font = BOLD; c.fill = HDR
IV = "公司自揭 KPI／指引／管理層說法：有呈現成長與可達成的誘因"
# ================= In =================
IN["A1"] = "In — 輸入（藍字；基準欄 E 進主模型；D／F＝敏感度低／高，只用於 Sens 頁、不進目標價；情境 A／B／C 的驅動見最後一節「情境驅動對照」）。金額 ¥B；ARR 百萬美元；股數百萬股；每股日圓。標記：Verified／Interested-party／Analogy／Assumed／Derived（Analogy、Assumed 給低／高區間；分配類參數對合併營收無影響者不給區間，見 README）"
IN["A2"] = f"事實總帳 data/facts_trendmicro.csv；Tokenomics {TKV} 快照 data/tk_snapshot.json（只讀 IF_ 名稱）。評價日 2026-06-30、股價基準日 2026-11-13。"
hdr(IN, 3, ["鍵", "項目", "單位", "敏感度低", "基準", "敏感度高", "標記", "來源", "說明", "來源誘因（Interested-party 列）"])
ROWS = []; INR = {}
def sec(t): ROWS.append(("sec", t))
def inp(key, item, unit, lo, base, hi, tag, src, note="", inc=""):
    if tag == "Interested-party" and not inc: inc = IV
    ROWS.append((key, item, unit, lo, base, hi, tag, src, note, inc))
D = datetime.date
sec("A. 日期、現價、股數與資產負債")
inp("data_date", "資料截止日", "日期", None, D(2026, 11, 16), None, "Derived", "本代理讀取日")
inp("px_date", "股價基準日（現價日）", "日期", None, D(2026, 11, 13), None, "Verified", fsrc("px"))
inp("val_date", "評價日（最近可得資產負債表日；2026Q3 資產負債表未找到）", "日期", None, D(2026, 6, 30), None, "Verified", fsrc("cashsec.q2"))
inp("tp_date", "目標時點＝EDATE(股價基準日,12)", "日期", None, "=EDATE(px_date,12)", None, "Derived", "原則第 10 節第 26 條")
inp("px", "現價（2026-11-13 收盤）", "日圓", None, 7214, None, "Verified", fsrc("px"), "Q3 財報後 +8.32%")
inp("sh_now", "流通股數（2026-11-13）", "百萬股", None, 129.93, None, "Verified", fsrc("sh.out.now"), "RSU／選擇權未見重大揭露，股票薪酬未找到（見 sbc_r）")
inp("sh_fy25", "流通股數（2025-12-31）", "百萬股", None, 130.680236, None, "Verified", fsrc("sh.out.fy25"))
inp("cash_0", "現金及存款＋短期投資（2026-06-30）", "¥B", None, 231.363, None, "Verified", fsrc("cashsec.q2"))
inp("fin_lease", "有利息負債與租賃負債（類債；皆未找到）", "¥B", 0.0, 0.0, 8.0, "Assumed", "資料商資產負債表無借款與租賃列（未找到）", "基準 0；高＝假設 8 十億日圓租賃負債")
sec("B. 命題模組：Vision One（原型 P：受保護範圍 × 每範圍模組數 × 每模組單價；單位＝每 US$1M ARR）")
inp("v1_arr_q3", "Vision One SaaS ARR（2026Q3）", "US$M", None, fv("arr.v1.q3"), None, "Interested-party", fsrc("arr.v1.q3"), "法說稱近 800（口徑不同）；本模型用新聞稿口徑")
inp("v1_g_q3", "Vision One ARR 年增率（2026Q3）", "比率", None, fv("arr.v1.g.q3"), None, "Interested-party", fsrc("arr.v1.g.q3"), "用於回推 2025 年底 ARR")
inp("fx", "日圓兌美元（ARR 換算營收；全期持平）", "日圓/US$", 140.0, 156.0, 165.0, "Interested-party", "公司 FY2026 指引假設 156（" + fsrc("guide.rev26") + "）", "FY2026H2 起匯率持平；Q3 季均約 159")
inp("s_b26", "受保護範圍成長（非 AI；FY2026；之後逐年乘 fade_b）", "比率", 0.14, 0.19, 0.24, "Assumed", "Vision One 年增約 46%＝範圍×模組×單價三項乘積的拆解，公司未揭露（未找到）", "錨：企業客戶 ARR 年增 14%、採用客戶逾 2,900")
inp("m_b26", "每範圍模組數成長（非 AI；FY2026）", "比率", 0.08, 0.12, 0.16, "Assumed", "同上；擴張客戶平均 4.2 個方案（" + fsrc("arr.v1.q3") + "）")
inp("p_b26", "每模組單價成長（非 AI；FY2026）", "比率", 0.01, 0.04, 0.06, "Assumed", "同上；公司稱曾漲價（" + fsrc("price.raise") + "）")
inp("ag26", "AI 增量：代理／非人類身分擴大受保護範圍（FY2026，加到範圍成長）", "pp", 0.01, 0.03, 0.05, "Assumed", "公司未揭露 AI 營收（未找到）；AI Security Package 見 research §5")
inp("am26", "AI 增量：AI 模組加購（FY2026，加到模組成長）", "pp", 0.005, 0.02, 0.035, "Assumed", "同上")
inp("fade_b", "非 AI 成長逐年衰減係數", "倍", 0.65, 0.75, 0.85, "Assumed", "無公開證據，Sens 涵蓋")
inp("fade_ai", "AI 增量逐年衰減係數", "倍", 0.70, 0.80, 0.90, "Assumed", "無公開證據，Sens 涵蓋")
for s, ka, kb, kc in (("kag", 1.5, 1.0, 0.5), ("kmod", 1.5, 1.0, 0.5)):
    for L, v in zip("ABC", (ka, kb, kc)):
        inp(f"{s}_{L}", ("AI 範圍增量" if s == "kag" else "AI 模組增量") + f"相對 FY2026 水準的倍數（FY2027 起；情境 {L}）", "倍", None, v, None, "Assumed", "情境設定：A 樂觀（代理防護需求快速放大）、C 保守", "分化依據：AI 增量與侵蝕的強度，三情境共用同一非 AI 路徑")
for L, v in zip("ABC", (0.005, 0.015, 0.035)):
    inp(f"aier_{L}", f"AI／競爭造成的每模組單價侵蝕（FY2027 起每年 pp；情境 {L}）", "pp", None, v, None, "Assumed", "Q2 公告提「競爭環境」、免費授權遷移方案（" + fsrc("price.raise") + "）；幅度未找到")
for L, v in zip("ABC", (0.20, 0.30, 0.45)):
    inp(f"tpmg_{L}", f"每 ARR 的 AI 任務量年增（情境 {L}）", "比率", None, v, None, "Assumed", "法說：單位 token 成本下降但總用量上升（" + fsrc("price.raise") + "），幅度未找到")
for L, v in zip("ABC", (0.01, 0.0, -0.02)):
    inp(f"gosh_{L}", f"其他營收（非 Vision One）年增調整（FY2027 起；情境 {L}）", "pp", None, v, None, "Assumed", "情境設定：舊產品萎縮速度")
sec("C. 營收分部（所在地別）")
for k, z, g, o, v in (("jp", "Japan", 0.01, -0.03, 0.22), ("am", "Americas", 0.06, -0.05, 0.38), ("eu", "Europe", 0.14, -0.01, 0.20)):
    inp(f"g26_{k}", f"{z} FY2026 營收年增（錨：2026H1 年增 0.7%／7.5%／16.6%／22.2% 依序為日本／美洲／歐洲／亞太，含弱日圓）", "比率", None, g, None, "Assumed", fsrc(f"seg.{z}.h1fy2026"), "H2 假設成長放緩（匯率順風消退）")
    inp(f"go_{k}", f"{z} 其他營收（非 Vision One）FY2027 起年增", "比率", None, o, None, "Assumed", "FY2025 分部年增（日本 +2.4%、美洲 −6.2%、歐洲 +4.9%、亞太 +2.9%）與舊產品萎縮說法（" + fsrc("arr.v1.q3.call") + "）")
    inp(f"vsh_{k}", f"Vision One 營收在{z}的占比", "比率", None, v, None, "Assumed", "法說：美國為最大、日本第二、德國、巴西第四（" + fsrc("arr.v1.q3.call") + "）；占比未揭露", "分配類，不改合併營收")
inp("g26_ap", "Asia Pacific FY2026 營收年增", "比率", None, 0.18, None, "Assumed", fsrc("seg.Asia Pacific.h1fy2026"))
inp("go_ap", "Asia Pacific 其他營收 FY2027 起年增", "比率", None, 0.0, None, "Assumed", "同上")
inp("vsh_ap", "Vision One 營收在 Asia Pacific 的占比（＝1 減其他三者）", "比率", None, "=1-vsh_jp-vsh_am-vsh_eu", None, "Derived", "同上")
inp("g26_sh", "FY2026 各分部年增同步調整（Sens 用）", "pp", -0.025, 0.0, 0.025, "Assumed", "H2 放緩幅度不確定")
sec("D. 成本：推論成本（只讀 Tokenomics 快照；公司未揭露 token 量與代理任務量，故每任務 token 為 Analogy，不標 Verified）")
tf, tc, td, tm = (TKI[n]["values"] for n in ("IF_TaskTokFresh", "IF_TaskTokCached", "IF_TaskTokDec", "IF_TaskTokMult"))
inp("chat_tok", "一般聊天每次請求總 token（新鮮＋快取＋decode）", "tok", None, tf[0] + tc[0] + td[0], None, "Derived", f"Tokenomics {TKV} IF_TaskTokFresh／Cached／Dec 第 1 欄加總（快照 data/tk_snapshot.json）", "")
inp("tok_mult", "每個安全分析任務的 token 倍數（相對一般聊天）", "倍", tm[1], tm[2], tm[3], "Analogy", f"Tokenomics {TKV} IF_TaskTokMult：0.1 小時任務 {tm[1]:.2f}、0.5 小時任務 {tm[2]:.1f}、2 小時任務 {tm[3]:.1f}", "公司未揭露每任務 token（未找到）；以 0.5 小時人類任務類比，低／高＝0.1／2 小時任務")
inp("ptok", "向模型商購買的 token 市價（混合請求）", "US$/百萬 token", TKI["IF_PriceRef_Luna"]["values"], TKI["IF_PriceRef_Terra"]["values"], TKI["IF_PriceRef_Sol"]["values"], "Analogy", f"Tokenomics {TKV} IF_PriceRef_Luna／Terra／Sol（OpenAI 有效價）", "選哪一層價格屬類比；公司稱採 Claude 與 Nemotron 開源模型，實際單價未找到。全成本 IF_FullCost 只是模型商成本下限，不用")
inp("tpm", "每 US$1M ARR 每年的 AI 任務數（FY2026）", "百萬任務/US$M", 0.8, 2.5, 6.0, "Assumed", "公司未揭露（未找到）", "承重最大 Assumed 之一；FY2026 推論成本占 Vision One 營收約 5%，上界檢查：公司下修營業利益指引 12 十億日圓（含非 token 項目）")
inp("defl", "token 單價年降幅", "比率", 0.05, 0.15, 0.30, "Assumed", "法說：單位成本下降（" + fsrc("price.raise") + "）；幅度未找到")
inp("cogs_r", "營業成本率（不含推論成本；基準全期持平於 FY2025 實績）", "比率", 0.215, round(fv("cogs.fy2025") / fv("rev.fy2025"), 4), 0.25, "Derived", fsrc("cogs.fy2025") + " ÷ " + fsrc("rev.fy2025"), "改善路徑無證據，只放 Sens；FY2025 推論成本視為 0（未找到）")
inp("oi_9m", "2026 前三季營業利益（Q1＋Q2＋Q3）", "¥B", None, round((fv("oi.2026Q1") + fv("oi.2026Q2") + fv("oi.2026Q3")) / 1000, 3), None, "Derived", "三季事實相加（Q3 為新聞稿四捨五入值）")
inp("rev_9m", "2026 前三季營收", "¥B", None, round((fv("rev.2026Q1") + fv("rev.2026Q2") + fv("rev.2026Q3")) / 1000, 3), None, "Derived", "三季事實相加")
inp("q4_opm", "2026Q4 營業利益率（含約 4 十億日圓離職費用）", "比率", 0.08, 0.12, 0.17, "Assumed", "Q3 約 17.5%、Q2 8%；Q4 離職費用見 " + fsrc("sev.q4"))
inp("sev26", "FY2026 一次性離職費用（Q3 1＋Q4 4）", "¥B", None, 5.0, None, "Interested-party", fsrc("sev.q4"), "FY2027 營業費用基期扣除")
inp("sav27", "FY2027 裁員節省（300–400 人）", "¥B", 1.5, 4.0, 7.0, "Assumed", fsrc("emp.q3"), "有公司揭露的減員數、單位成本未揭露；原則 35：改善路徑只在此一項且有揭露減員為證")
for L, v in zip("ABC", (0.02, 0.03, 0.045)):
    inp(f"opxg_{L}", f"營業費用年增（FY2027 起；情境 {L}）", "比率", None, v, None, "Assumed", "FY2025 營業費用 −3.2%、FY2026H1 營業成本加費用 +20.3%（" + fsrc("oi.h1fy2026") + "）", "情境設定：A 生產力提升、C 競爭與投資壓力；G3：FY2021–2025 營業費用率 55.0%／61.7%／61.3%／58.5%／56.0% 無單調規模經濟，故以年增率而非費用率建模（費用率隨營收成長而降，是輸出）；全期費用率持平案見 Sens 的 opxg 高案")
inp("nonop", "營業外收支淨額（年）", "¥B", -3.0, 2.0, 5.0, "Assumed", "FY2024 +7.5、FY2025 −5.4（稅前減營業利益）；2026H1 經常利益減營業利益 +3.0")
inp("te", "稅與少數股權有效稅率（稅前到歸屬淨利）", "比率", 0.30, 0.335, 0.37, "Derived", "FY2025 (52,329−34,523)÷52,329＝34.0%；FY2026 指引 1−30.7÷46.0＝33.3%")
sec("E. 股東回饋、現金流")
inp("payout", "配息率（占歸屬淨利）", "比率", 0.5, 0.70, 0.9, "Assumed", "FY2025 185 日圓×130.68M÷34,523＝70.0%（" + fsrc("div.fy25") + "）")
inp("bb", "庫藏股買回（每年）", "¥B", 0.0, 10.0, 25.0, "Assumed", "FY2024 40、FY2025 約 5（2026H1 為 5.0；" + fsrc("bb.h1fy2026") + "）")
inp("dan", "折舊攤銷淨額（扣資本化合約成本）占營收", "比率", 0.01, 0.025, 0.05, "Assumed", "FY2024 營業現金流 46.8 減淨利 34.4＝12.4（營收 4.6%），含遞延收入變動；FY2021–2024 同差 8.9%／12.1%／18.7%／4.6%（高波動）")
inp("cx_r", "資本支出占營收", "比率", 0.004, 0.006, 0.01, "Derived", "FY2021–2024 資本支出占營收 1.1%、0.6%、0.9%、0.4%（" + fsrc("capex.fy2024") + "）")
inp("nwc_k", "營運資金增加 ÷ 營收增量（負＝遞延收入融資）", "倍", -0.5, -0.25, 0.0, "Assumed", "FY2025 合約負債 236.1、應收 76.1，淨營運資金約負 160（營收 58%）；保守取 −0.25")
inp("sbc_r", "未計入營業利益的股票薪酬占營收（Sens 用）", "比率", 0.0, 0.0, 0.01, "Assumed", "合併股票薪酬未單獨揭露（未找到），基準 0")
sec("F. 折現率與目標價規則")
inp("rf", "無風險利率（日本 10 年期公債，2026-11-11）", "比率", None, fv("rf"), None, "Verified", fsrc("rf"), "與現價日相差 2 個交易日，不逾期")
inp("beta_own", "Beta（5 年，資料商；第二來源未找到）", "—", None, fv("beta.sa"), None, "Verified", fsrc("beta.sa"))
inp("beta_peer", "同業 Beta 中位數（七家）", "—", None, 0.84, None, "Derived", "七家 5 年 Beta：0.42／1.26／0.60／1.04／0.84／0.72／0.97 中位數（事實總帳 peer.*）")
inp("beta", "Beta（兩個來源各做 Blume 調整後取中位數＝平均）", "—", 0.62, "=AVERAGE(0.67*beta_own+0.33,0.67*beta_peer+0.33)", 1.0, "Derived", "DCF 規格第 2 節；低＝僅用自身 Beta，高＝1.0")
inp("wacc_adj", "WACC 加減（Sens 用，基準 0）", "pp", -0.01, 0.0, 0.01, "Assumed", "DCF 規格第 4 節：WACC ±1pt")
inp("erp", "股權風險溢酬", "比率", 0.04, 0.05, 0.06, "Assumed", "DCF 規格預設 5%")
inp("g_tv", "永續成長率（RONIC＝WACC 下終值＝NOPAT÷WACC，對 g 近乎中性；預設 3% 改 2%：日圓名目）", "比率", 0.01, 0.02, 0.03, "Assumed", "DCF 規格預設 3%、敏感度 2–4%；本模型取 2% 並寫明")
inp("w_dcf", "DCF 權重（Comparable 為 1 減本值）", "比率", 0.3, 0.5, 0.7, "Assumed", "原則第 4 節預設各半")
inp("pm_chkp", "同業 EV/EBIT（TTM）：Check Point（2026-11-13；US GAAP）", "倍", None, 16.51, None, "Verified", "stockanalysis.com/stocks/chkp/statistics/", "低成長子群（3 年營收成長預估 ≤10%）")
inp("pm_gen", "同業 EV/EBIT（TTM）：Gen Digital（2026-10-26；US GAAP；逾期）", "倍", None, 10.20, None, "Verified", "stockanalysis.com/stocks/gen/statistics/", "價格日 2026-10-26 與現價日相差逾 2 個交易日：逾期；Gen 有槓桿（淨負債），EV 口徑對 EV")
inp("pm_qlys", "同業 EV/EBIT（TTM）：Qualys（2026-10-26；US GAAP；逾期）", "倍", None, 29.48, None, "Verified", "stockanalysis.com/stocks/qlys/statistics/", "逾期，同上")
RMSRC = "software_ai/data/relmult_hist/out/trendmicro.json（groups.all；2026-11-17 產出；窗口 2021-09-30～2026-06-30，20 季）"
inp("cm_adj", "相對倍數 rel_mult（本公司 ÷ 同業中位；PS 與 PSG 兩讀數幾何平均＝geo_base；基準＝geo_base，低／高＝sens.low／sens.high）", "倍", RMS["low"], RM["geo_base"], RMS["high"], "Derived", RMSRC,
    "base_source：" + RM["base_source"] + "。同業 CHKP、GEN、QLYS、FTNT；Comparable 主腿為 P/S 口徑（G6），與 PS／PSG 相對值同口徑（同業 P/S 中位 × rel_mult）。v0.1 為 1.00（EV/EBIT）")
inp("rm_ps", "PS 相對值讀數（20 季中位數；Sens 案 ps_only）", "倍", None, RM["ps_reading"], None, "Derived", RMSRC, "PS＝市值÷TTM 營收；範圍 %.2f–%.2f" % (RM["ps_stats"]["min"], RM["ps_stats"]["max"]))
inp("rm_psg", "PSG 相對值讀數（14 季中位數；Sens 案 psg_only）", "倍", None, RM["psg_reading"], None, "Derived", RMSRC, "PSG＝PS÷TTM 營收年增(pt)；最早 4 季無基期")
inp("rm_psgm", "PSG 相對值中位數（14 季；未乘成長比）", "倍", None, RM["psg_stats"]["median"], None, "Derived", RMSRC, "PSG 讀數＝本值 × 成長比")
inp("rm_gco", "本公司 TTM 營收年增（2026-06-30）", "比率", None, RM["growth_ratio"]["company_g"], None, "Derived", RMSRC, "TTM（NTM 未取得）")
inp("rm_gpe", "同業中位 TTM 營收年增（2026-06-30）", "比率", None, RM["growth_ratio"]["peer_median_g"], None, "Derived", RMSRC, "同業 %d 家" % RM["growth_ratio"]["n_peers"])
inp("rm_gr", "成長比（本公司年增 ÷ 同業中位年增）", "倍", None, RM["growth_ratio"]["value"], None, "Derived", RMSRC, "解釋 PS 與 PSG 兩讀數分歧")
inp("rm_div", "兩讀數分歧度（|PS 讀數 − PSG 讀數| ÷ PS 讀數；divergence_gt_30pct＝" + str(RM["divergence_gt_30pct"]) + "）", "比率", None, RM["divergence"], None, "Derived", RMSRC, "超過 30% 要在 README 與 HTML 寫原因")
inp("sbc_dir", "董事股票選擇權報酬（FY2025；僅董事，合併股票薪酬未找到）", "¥B", None, fv("sbc.dir.fy25") / 1000, None, "Verified", fsrc("sbc.dir.fy25"), "只作 SBC 加回對照的下限；不進加權目標價")
inp("sbc_ab", "SBC 加回對照：股票薪酬占營收（FY2025 董事選擇權 ÷ 合併營收）", "比率", None, fv("sbc.dir.fy25") / fv("rev.fy2025"), None, "Derived", fsrc("sbc.dir.fy25") + "；" + fsrc("rev.fy2025"), "合併 SBC 未找到，此為下限（不含員工）；只用於 Val 的對照列")
PPSRC = "software_ai/data/relmult_hist/out/peer_ps_20261113.json（companies.trendmicro；StockAnalysis；Derived：市值÷最新 TTM 營收；價格日目標 2026-11-13）"
inp("ps_chkp", "同業 P/S（TTM）：Check Point（2026-11-13；US GAAP）", "倍", None, PPP["CHKP"], None, "Verified", PPSRC, "價格日 2026-11-13，未逾期；P/S 不受會計準則差異影響")
inp("ps_gen", "同業 P/S（TTM）：Gen Digital（2026-10-26；US GAAP；逾期）", "倍", None, PPP["GEN"], None, "Verified", PPSRC, "逾期：價格日 2026-10-26，與現價日相差逾 2 個交易日（stale 列：" + "、".join(PPG["stale"]) + "）")
inp("ps_qlys", "同業 P/S（TTM）：Qualys（2026-10-26；US GAAP；逾期）", "倍", None, PPP["QLYS"], None, "Verified", PPSRC, "逾期，同上")
inp("ps_ftnt", "同業 P/S（TTM）：Fortinet（2026-11-13；US GAAP）", "倍", None, PPP["FTNT"], None, "Verified", PPSRC, "價格日 2026-11-13，未逾期")
inp("ps_med", "同業 P/S 中位數（四家；Comparable 主腿；Sens 案改為資料庫最近一季中位）", "倍", None, "=MEDIAN(ps_chkp,ps_gen,ps_qlys,ps_ftnt)", None, "Derived", PPSRC, "v0.2 為資料庫 2026-06-30 季末中位 %.2f（對照案 ps_db）；G7" % PPG["db_last_quarter"]["median"])
inp("ps_db", "對照：同業 P/S 資料庫最近一季（2026-06-30）中位數（Sens 案用）", "倍", None, PPG["db_last_quarter"]["median"], None, "Derived", PPSRC, "即 v0.2 的值；季末列，相對現價日逾期")
inp("ps_tm", "本公司 P/S（TTM；2026-11-13）", "倍", None, PP["own_ps"], None, "Verified", PPSRC, "與同業同價格日，用於現行相對值")
inp("tp_wA", "情境權重 A", "比率", None, 0.25, None, "Assumed", "DCF 規格第 3 節預設 25／50／25")
inp("tp_wB", "情境權重 B", "比率", None, 0.50, None, "Assumed", "同上")
inp("tp_wC", "情境權重 C", "比率", None, 0.25, None, "Assumed", "同上")
inp("cons_tp", "賣方平均目標價（9 家；財報前，只作對照，不進模型）", "日圓", 5000, fv("cons.tp"), 9000, "Interested-party", fsrc("cons.tp"), "只作對照，不進模型", "賣方有維持客戶關係與評等立場的誘因；只作對照，不進模型")
SCN = [("scn_kag", "AI 範圍增量倍數", "kag"), ("scn_kmod", "AI 模組增量倍數", "kmod"), ("scn_aier", "每模組單價侵蝕", "aier"), ("scn_tpmg", "AI 任務量年增", "tpmg"),
       ("scn_gosh", "其他營收年增調整", "gosh"), ("scn_opxg", "營業費用年增", "opxg"), ("scn_w", "情境權重", "tp_w")]
FP = hashlib.sha256()
r = 4
for row in ROWS:
    if row[0] == "sec":
        IN.cell(r, 1, row[1]).font = BOLD; r += 1; continue
    key, item, unit, lo, base, hi, tag, src, note, inc = row
    if key in OVR: base = OVR[key]
    for c, v in enumerate((key, item, unit, lo, base, hi, tag, src, note, inc), 1):
        if v is not None: IN.cell(r, c, v)
    for c in (4, 5, 6):
        if isinstance(IN.cell(r, c).value, (int, float)): IN.cell(r, c).font = BLUE
        if isinstance(IN.cell(r, c).value, (datetime.date,)): IN.cell(r, c).number_format = "yyyy-mm-dd"
    INR[key] = r; dn(key, f"In!$E${r}")
    FP.update(f"{key}={base}".encode()); r += 1
IN_END = r
FPR = FP.hexdigest()[:12]
r += 1
IN.cell(r, 1, "情境驅動對照（純展示：D／E／F＝情境 A／B／C，只引用上方輸入列；修改請改上方藍字）").font = BOLD; SCN_SEC = r; r += 1
hdr(IN, r, ["鍵", "項目", "單位", "情境 A（樂觀）", "情境 B（基準）", "情境 C（保守）", "", "說明"]); SCN_H = r; r += 1
for k, nm, pre in SCN:
    ks = [f"{pre}_{L}" for L in "ABC"] if pre != "tp_w" else ["tp_wA", "tp_wB", "tp_wC"]
    IN.cell(r, 1, k); IN.cell(r, 2, nm)
    for c, kk in zip((4, 5, 6), ks): IN.cell(r, c, f"=E{INR[kk]}")
    IN.cell(r, 8, "引用 In 頁 " + "／".join(ks) + "；分化依據：A＝AI 增量大、侵蝕小、推論成本成長慢；C 相反")
    r += 1
for c, w in zip("ABCDEFGHIJ", (14, 60, 14, 12, 12, 12, 16, 70, 60, 40)): IN.column_dimensions[c].width = w
# ================= Hist =================
HI["A1"] = "Hist — 歷史財務（¥B；公式只引用本頁；來源與標記見各列）。FY2021–FY2025＝C–G 欄"
hdr(HI, 3, ["項目", "單位", "FY2021", "FY2022", "FY2023", "FY2024", "FY2025", "標記", "來源"])
HR = {}
def hrow(key, label, unit, vals, tag, src, scale=1000.0):
    rr = 4 + len(HR); HR[key] = rr
    HI.cell(rr, 1, label); HI.cell(rr, 2, unit)
    for j, v in enumerate(vals):
        if v is not None: HI.cell(rr, 3 + j, round(v / scale, 6)).font = BLUE
    HI.cell(rr, 8, tag); HI.cell(rr, 9, src)
def fl(pat): return [FJ.get(pat.format(y), {}).get("value") for y in range(2021, 2026)]
hrow("rev", "營收", "¥B", fl("rev.fy{}"), "Verified", "Annual Report FY2023–2025 p.4；FY2021 資料商二源")
hrow("cogs", "營業成本", "¥B", fl("cogs.fy{}"), "Verified", "AR24 p.26、AR25 p.20；其餘資料商")
hrow("oi", "營業利益", "¥B", fl("oi.fy{}"), "Verified", "AR23 p.2、AR24 p.26、AR25 p.20")
hrow("pti", "稅前利益", "¥B", fl("pti.fy{}"), "Verified", "stockanalysis（資料商；法定申報原文未讀）")
hrow("ni", "歸屬母公司淨利", "¥B", fl("ni.fy{}"), "Verified", "AR25 p.4")
hrow("eps", "EPS（基本，日圓）", "日圓", [None] + [FJ[f"eps.fy{y}"]["value"] for y in (2022, 2023, 2024, 2025)], "Verified", "AR25 p.4", scale=1.0)
for rn in ("Japan", "Americas", "Europe", "Asia Pacific"):
    hrow("s_" + rn, rn, "¥B", [None] + [FJ.get(f"seg.{rn}.fy{y}", {}).get("value") for y in (2022, 2023, 2024, 2025)], "Verified", "AR25 p.38／AR24 p.40／AR23 p.1；FY2022 為回推（Derived）")
hrow("emp", "員工數", "人", [None, None, None, fv("emp.fy24"), fv("emp.fy25")], "Verified", "AR24／AR25 p.5", scale=1.0)
hrow("cfo", "營業現金流", "¥B", [fv(f"cfo.fy{y}") for y in (2021, 2022, 2023, 2024)] + [None], "Verified", "stockanalysis（資料商）")
hrow("capex", "資本支出", "¥B", [fv(f"capex.fy{y}") for y in (2021, 2022, 2023, 2024)] + [None], "Verified", "stockanalysis；FY2023／2024 與 AR 一致")
HI.column_dimensions["A"].width = 24; HI.column_dimensions["I"].width = 60
def HC(key, y): return f"Hist!{'CDEFG'[y - 2021]}{HR[key]}"
# ================= PL（三情境）=================
YRS = list(range(2023, 2033)); FC = list(range(2026, 2033))
def cl(y): return "CDEFGHIJKL"[y - 2023]
PL["A1"] = "PL — 三情境損益（¥B；GAAP＝日本會計準則財報口徑）。命題＝Vision One 的受保護範圍×模組數×單價（成長形式）；分部＝所在地別四分部；FY2026E＝前三季實績＋Q4 估。"
PL["A2"] = "AI 淨效果（¥B）＝AI 營收增量 − 單價侵蝕 − 推論成本；推論成本＝平均 ARR × 每 ARR 任務數 × 每任務 token × token 市價（Tokenomics 快照）。"
HD = 4
hdr(PL, HD, ["項目", "單位"] + [f"FY{y}{'A' if y <= 2025 else 'E'}" for y in YRS])
R = {}; cur = [HD + 1]
def nr(): 
    x = cur[0]; cur[0] += 1; return x
def row(s, key, label, unit, fn, hist=None, fmt="#,##0.0", bold=False):
    rr = nr(); R[s][key] = rr
    PL.cell(rr, 1, label); PL.cell(rr, 2, unit)
    if bold: PL.cell(rr, 1).font = BOLD
    for y in YRS:
        v = hist(y) if (y <= 2025 and hist) else (fn(y) if y >= 2026 and fn else None)
        if v is not None:
            c = PL[f"{cl(y)}{rr}"]; c.value = v; c.number_format = fmt
    return rr
def ref(s, key, y): return f"{cl(y)}{R[s][key]}"
REG = [("jp", "Japan"), ("am", "Americas"), ("eu", "Europe"), ("ap", "Asia Pacific")]
def block(s):
    R[s] = {}
    PL.cell(nr(), 1, f"情境 {s}").font = BOLD
    P = lambda k, y: ref(s, k, y)
    row(s, "tt", "距 FY2026 年數", "年", lambda y: y - 2026, fmt="0")
    row(s, "sb", "範圍成長（非 AI）", "比率", lambda y: f"=s_b26*fade_b^{P('tt', y)}", fmt="0.0%")
    row(s, "mb", "模組數成長（非 AI）", "比率", lambda y: f"=m_b26*fade_b^{P('tt', y)}", fmt="0.0%")
    row(s, "pb", "單價成長（非 AI）", "比率", lambda y: f"=p_b26*fade_b^{P('tt', y)}", fmt="0.0%")
    row(s, "ag", "AI 範圍增量", "pp", lambda y: f"=IF({P('tt', y)}=0,ag26,ag26*kag_{s}*fade_ai^{P('tt', y)})", fmt="0.0%")
    row(s, "am", "AI 模組增量", "pp", lambda y: f"=IF({P('tt', y)}=0,am26,am26*kmod_{s}*fade_ai^{P('tt', y)})", fmt="0.0%")
    row(s, "er", "單價侵蝕", "pp", lambda y: f"=IF({P('tt', y)}=0,0,aier_{s})", fmt="0.0%")
    for k, nm in (("a0", "ARR 路徑 0：無 AI 增量、無侵蝕"), ("a1", "ARR 路徑 1：加 AI 增量"), ("a2", "ARR 路徑 2：加 AI 增量與侵蝕（進營收）")):
        rr = nr(); R[s][k] = rr; PL.cell(rr, 1, f"Vision One {nm}（年底）"); PL.cell(rr, 2, "US$M")
        PL[f"E{rr}"] = "=v1_arr_q3/(1+v1_g_q3)^0.75"
        for y in FC:
            sc = f"(1+{P('sb', y)}{'+' + P('ag', y) if k != 'a0' else ''})"; mo = f"(1+{P('mb', y)}{'+' + P('am', y) if k != 'a0' else ''})"
            pr = f"(1+{P('pb', y)}{'-' + P('er', y) if k == 'a2' else ''})"
            PL[f"{cl(y)}{rr}"] = f"={cl(y - 1)}{rr}*{sc}*{mo}*{pr}"
        for y in [2025] + FC: PL[f"{cl(y)}{rr}"].number_format = "#,##0"
    av = lambda k, y: f"(({cl(y - 1)}{R[s][k]}+{cl(y)}{R[s][k]})/2)"
    row(s, "v1", "Vision One 營收", "¥B", lambda y: f"={av('a2', y)}*fx/1000")
    row(s, "inc", "AI 營收增量（路徑 1 − 路徑 0）", "¥B", lambda y: f"=({av('a1', y)}-{av('a0', y)})*fx/1000")
    row(s, "ero", "單價侵蝕（路徑 1 − 路徑 2）", "¥B", lambda y: f"=({av('a1', y)}-{av('a2', y)})*fx/1000")
    row(s, "inf", "AI 推論成本", "¥B", lambda y: f"={av('a2', y)}*tpm*chat_tok*tok_mult*ptok*(1-defl)^{P('tt', y)}*(1+tpmg_{s})^{P('tt', y)}/1000000*fx/1000")
    row(s, "ain", "AI 淨效果（增量 − 侵蝕 − 推論成本）", "¥B", lambda y: f"={P('inc', y)}-{P('ero', y)}-{P('inf', y)}", bold=True)
    row(s, "inf_r", "推論成本占營業成本", "比率", lambda y: None, fmt="0.0%")
    for k, z in REG:
        r1_ = row(s, "r_" + k, z, "¥B", None, hist=lambda y, z=z: f"={HC('s_' + z, y)}", bold=True)
        r2_ = row(s, "v_" + k, "Vision One（命題子列）｜" + z, "¥B", None)
        r3_ = row(s, "o_" + k, "其他｜" + z, "¥B", None)
        for y in FC:
            PL[f"{cl(y)}{r1_}"] = (f"={HC('s_' + z, 2025)}*(1+g26_{k}+g26_sh)" if y == 2026 else f"={cl(y)}{r2_}+{cl(y)}{r3_}")
            PL[f"{cl(y)}{r2_}"] = f"={P('v1', y)}*vsh_{k}"
            PL[f"{cl(y)}{r3_}"] = (f"={cl(y)}{r1_}-{cl(y)}{r2_}" if y == 2026 else f"={cl(y - 1)}{r3_}*(1+go_{k}+gosh_{s})")
            for r_ in (r1_, r2_, r3_): PL[f"{cl(y)}{r_}"].number_format = "#,##0.0"
    row(s, "rev", "合併營收", "¥B", lambda y: "=" + "+".join(P("r_" + k, y) for k, _ in REG), hist=lambda y: "=" + "+".join(P("r_" + k, y) for k, _ in REG), bold=True)
    row(s, "chk1", "核對：分部合計 − 公司合併營收（歷史）", "¥B", None, hist=lambda y: f"={P('rev', y)}-{HC('rev', y)}", fmt="0.000")
    row(s, "chk2", "核對：命題子列合計 − 分部", "¥B", lambda y: "=" + "+".join(f"ABS({P('v_' + k, y)}+{P('o_' + k, y)}-{P('r_' + k, y)})" for k, _ in REG), hist=lambda y: 0, fmt="0.000")
    row(s, "cogs", "營業成本", "¥B", lambda y: f"=cogs_r*{P('rev', y)}+{P('inf', y)}", hist=lambda y: f"={HC('cogs', y)}")
    row(s, "gp", "毛利", "¥B", lambda y: f"={P('rev', y)}-{P('cogs', y)}", hist=lambda y: f"={P('rev', y)}-{P('cogs', y)}")
    ro = row(s, "opx", "營業費用（銷管費含研發）", "¥B", None, hist=lambda y: f"={P('gp', y)}-{HC('oi', y)}")
    rq = row(s, "oi", "營業利益", "¥B", None, hist=lambda y: f"={HC('oi', y)}", bold=True)
    for y in FC:
        PL[f"{cl(y)}{rq}"] = (f"=oi_9m+({P('rev', y)}-rev_9m)*q4_opm" if y == 2026 else f"={P('gp', y)}-{P('opx', y)}")
        PL[f"{cl(y)}{ro}"] = (f"={P('gp', y)}-{P('oi', y)}" if y == 2026 else f"=({cl(y - 1)}{ro}-IF({P('tt', y)}=1,sev26,0))*(1+opxg_{s})-IF({P('tt', y)}=1,sav27,0)")
        PL[f"{cl(y)}{rq}"].number_format = PL[f"{cl(y)}{ro}"].number_format = "#,##0.0"
    row(s, "non", "營業外收支淨額", "¥B", lambda y: "=nonop", hist=lambda y: f"={HC('pti', y)}-{HC('oi', y)}")
    row(s, "pti", "稅前利益", "¥B", lambda y: f"={P('oi', y)}+{P('non', y)}", hist=lambda y: f"={HC('pti', y)}")
    row(s, "tax", "所得稅與少數股權", "¥B", lambda y: f"={P('pti', y)}*te", hist=lambda y: f"={HC('pti', y)}-{HC('ni', y)}")
    row(s, "ni", "歸屬母公司淨利", "¥B", lambda y: f"={P('pti', y)}-{P('tax', y)}", hist=lambda y: f"={HC('ni', y)}", bold=True)
    row(s, "sh", "流通股數（年底）", "百萬股", lambda y: ("=sh_now-bb*0.3/px" if y == 2026 else f"={cl(y - 1)}{R[s]['sh']}-bb/(px*1.05^{P('tt', y)})"), fmt="#,##0.00")
    PL[f"E{R[s]['sh']}"] = "=sh_fy25"
    row(s, "eps", "EPS（日圓）", "日圓", lambda y: f"={P('ni', y)}*1000/(({cl(y - 1)}{R[s]['sh']}+{cl(y)}{R[s]['sh']})/2)",
        hist=lambda y: (f"={HC('eps', y)}" if y >= 2023 else None), fmt="#,##0.0")
    row(s, "dps", "每股股利（日圓）", "日圓", lambda y: f"=payout*{P('eps', y)}", fmt="#,##0.0")
    row(s, "div", "股利支出", "¥B", lambda y: f"={P('dps', y)}*(({cl(y - 1)}{R[s]['sh']}+{cl(y)}{R[s]['sh']})/2)/1000")
    row(s, "fcf", "自由現金流（淨利＋折舊攤銷淨額 − 資本支出 − 營運資金增加）", "¥B",
        lambda y: f"={P('ni', y)}+(dan-cx_r)*{P('rev', y)}-nwc_k*({P('rev', y)}-{cl(y - 1)}{R[s]['rev']})")
    row(s, "fcfd", "自由現金流扣股利", "¥B", lambda y: f"={P('fcf', y)}-{P('div', y)}", bold=True)
    row(s, "nopat", "NOPAT（營業利益扣未計股票薪酬後 ×(1−稅率)）", "¥B", lambda y: f"=({P('oi', y)}-sbc_r*{P('rev', y)})*(1-te)")
    row(s, "fcff", "FCFF（NOPAT＋折舊攤銷淨額 − 資本支出 − 營運資金增加）", "¥B", lambda y: f"={P('nopat', y)}+(dan-cx_r)*{P('rev', y)}-nwc_k*({P('rev', y)}-{cl(y - 1)}{R[s]['rev']})", bold=True)
    for y in FC: PL[f"{cl(y)}{R[s]['inf_r']}"] = f"={P('inf', y)}/{P('cogs', y)}"
    cur[0] += 1
for s in "ABC": block(s)
# ---- 標準損益表（情境 B；GAAP）與核對列 ----
PL.cell(cur[0], 1, "標準損益表（情境 B 基準；日本會計準則，財報口徑；歷史 EPS 為基本 EPS）").font = BOLD; cur[0] += 1
IX = {}; B = "B"
def isr(key, label, unit, f, fmt="#,##0.0"):
    rr = cur[0]; cur[0] += 1; IX[key] = rr; PL.cell(rr, 1, label); PL.cell(rr, 2, unit)
    for y in YRS:
        v = f(y)
        if v is not None: PL[f"{cl(y)}{rr}"] = v; PL[f"{cl(y)}{rr}"].number_format = fmt
def pb(k, y): return ref("B", k, y)
isr("rev", "營收", "¥B", lambda y: f"={pb('rev', y)}")
isr("yoy", "營收年增率（YoY）", "比率", lambda y: (f"={pb('rev', y)}/{pb('rev', y - 1)}-1" if y > 2023 else f"={pb('rev', 2023)}/{HC('rev', 2022)}-1"), "0.0%")
isr("cogs", "營業成本（COGS）", "¥B", lambda y: f"={pb('cogs', y)}")
isr("gp", "毛利", "¥B", lambda y: f"={pb('gp', y)}")
isr("gm", "毛利率", "比率", lambda y: f"={pb('gp', y)}/{pb('rev', y)}", "0.0%")
isr("opx", "營業費用", "¥B", lambda y: f"={pb('opx', y)}")
isr("oi", "營業利益", "¥B", lambda y: f"={pb('oi', y)}")
isr("opm", "營業利益率", "比率", lambda y: f"={pb('oi', y)}/{pb('rev', y)}", "0.0%")
isr("non", "其他收入（淨額；營業外）", "¥B", lambda y: f"={pb('non', y)}")
isr("pti", "稅前利益", "¥B", lambda y: f"={pb('pti', y)}")
isr("tax", "所得稅（含少數股權）", "¥B", lambda y: f"={pb('tax', y)}")
isr("ni", "稅後淨利（歸屬母公司）", "¥B", lambda y: f"={pb('ni', y)}")
isr("nm", "淨利率", "比率", lambda y: f"={pb('ni', y)}/{pb('rev', y)}", "0.0%")
isr("eps", "EPS（日圓；預測為完全稀釋口徑，無稀釋證券）", "日圓", lambda y: f"={pb('eps', y)}")
PL.cell(cur[0], 1, "比率區").font = BOLD; IX["ratio"] = cur[0]; cur[0] += 1
isr("niy", "淨利年增率", "比率", lambda y: (f"={pb('ni', y)}/{pb('ni', y - 1)}-1" if y > 2023 else f"={pb('ni', 2023)}/{HC('ni', 2022)}-1"), "0.0%")
cur[0] += 1
PL.cell(cur[0], 1, "分部加總與核對（三情境取絕對值最大者，應為 0）").font = BOLD; cur[0] += 1
CK1 = cur[0]; CK2 = cur[0] + 1
for rr, lab, k in ((CK1, "檢查：分部合計 − 合併營收", "chk1"), (CK2, "檢查：命題子列合計 − 分部", "chk2")):
    PL.cell(rr, 1, lab); PL.cell(rr, 2, "¥B")
    for y in YRS:
        if k == "chk1" and y >= 2026: PL[f"{cl(y)}{rr}"] = 0
        else: PL[f"{cl(y)}{rr}"] = "=MAX(" + ",".join(f"ABS({ref(s, k, y)})" for s in "ABC") + ")"
        PL[f"{cl(y)}{rr}"].number_format = "0.000"
cur[0] += 3
PL.column_dimensions["A"].width = 62; PL.column_dimensions["B"].width = 10
for y in YRS: PL.column_dimensions[cl(y)].width = 11
# ---- CF：收支關卡 ----
CF["A1"] = "CF — 收支關卡（自由現金流扣股利，FY2026–FY2030 合計；原則第 10 節第 55、63 條）。口徑：淨利＋折舊攤銷淨額 − 資本支出 − 營運資金增加，扣股利；庫藏股買回另列不扣。"
hdr(CF, 3, ["項目", "情境 A", "情境 B", "情境 C"])
GATE = 4
CF.cell(GATE, 1, "收支關卡（FY2026–FY2030 自由現金流扣股利合計，¥B）")
for c, s in zip((2, 3, 4), "ABC"):
    CF.cell(GATE, c, f"=SUM(PL!{cl(2026)}{R[s]['fcfd']}:{cl(2030)}{R[s]['fcfd']})").number_format = "#,##0.0"
CF.cell(GATE + 1, 1, "期間自由現金流合計（¥B）")
CF.cell(GATE + 2, 1, "期間股利合計（¥B）")
CF.cell(GATE + 3, 1, "期間庫藏股買回合計（¥B；不進關卡）")
for c, s in zip((2, 3, 4), "ABC"):
    CF.cell(GATE + 1, c, f"=SUM(PL!{cl(2026)}{R[s]['fcf']}:{cl(2030)}{R[s]['fcf']})").number_format = "#,##0.0"
    CF.cell(GATE + 2, c, f"=SUM(PL!{cl(2026)}{R[s]['div']}:{cl(2030)}{R[s]['div']})").number_format = "#,##0.0"
    CF.cell(GATE + 3, c, "=5*bb").number_format = "#,##0.0"
CF.cell(GATE + 4, 1, "關卡判定"); 
for c in (2, 3, 4): CF.cell(GATE + 4, c, f'=IF({"BCD"[c - 2]}{GATE}<0,"不通過","通過")')
CF.cell(GATE + 6, 1, "最低年末現金：不適用（無有利息負債；基準不模擬現金餘額，現金 cash_0 為評價日值）；舉債：無")
CF.column_dimensions["A"].width = 70
# ================= Val =================
VR = {}; vcur = [4]
VA["A1"] = "Val — 估值：DCF（FCFF，明確期到 FY2032，評價日 2026-06-30，期中折現）與 Comparable（同業 P/S 中位 × rel_mult × 目標日 TTM 營收，G6 主腿）各半；EV/EBIT Comparable 只作對照列，不進加權目標價；三情境依 In 頁權重加權。金額 ¥B，每股日圓。"
VA["A2"] = "目標時點＝股價基準日＋12 個月；DCF 值由評價日以股權成本推到目標時點並扣期間股利；上檔空間為持有期間（現價日起 12 個月）價差，不是年化報酬。"
VA["A3"] = "三情境差異只在 Vision One 的 AI 增量／侵蝕／推論成本、其他營收年增與營業費用年增；DCF 與 Comparable 落差分解見下方。"
hdr(VA, 4, ["項目", "情境 A／數值", "情境 B", "情境 C", "說明"]); vcur[0] = 5
def vr(key, label, b=None, c=None, d=None, fmt="#,##0.00", note=""):
    rr = vcur[0]; vcur[0] += 1; VR[key] = rr; VA.cell(rr, 1, label)
    for col, v in zip((2, 3, 4), (b, c, d)):
        if v is not None: VA.cell(rr, col, v).number_format = fmt
    if note: VA.cell(rr, 5, note)
    return rr
def v3(f, fmt="#,##0.00"): return [f(s) for s in "ABC"]
def PLs(s, k, y): return f"PL!{ref(s, k, y)}"
vr("rf", "無風險利率 rf", "=rf", fmt="0.00%", note="日本 10 年期公債 2026-11-11（相差 2 個交易日，不逾期）")
vr("ke", "股權成本 ke＝rf＋β×ERP", "=rf+beta*erp", fmt="0.00%")
vr("wacc", "WACC（無有利息負債，權重全為股權；含 wacc_adj）", "=B%d+wacc_adj" % VR["ke"], fmt="0.00%", note="kd 不適用：基準無有利息負債（資料商無借款列，未找到）")
vr("ronic", "RONIC（＝WACC，新投資價值中性）", "=B%d" % VR["wacc"], fmt="0.00%")
vr("stub", "評價日到 FY2026 年底的年數", "=(DATE(2026,12,31)-val_date)/365", fmt="0.000")
vr("ytp", "評價日到目標時點的年數", "=(tp_date-val_date)/365", fmt="0.000")
W_, KE, ST, YT, G = f"$B${VR['wacc']}", f"$B${VR['ke']}", f"$B${VR['stub']}", f"$B${VR['ytp']}", "g_tv"
def pvx(s):
    t = [f"PL!{ref(s, 'fcff', 2026)}*{ST}/(1+{W_})^({ST}/2)"]
    for y in range(2027, 2033): t.append(f"PL!{ref(s, 'fcff', y)}/(1+{W_})^({ST}+{y - 2027}+0.5)")
    return "=" + "+".join(t)
vr("pvx", "明確期 FCFF 現值", *v3(pvx, ), fmt="#,##0.0")
vr("tv", "終值 FY2033 NOPAT×(1−g÷RONIC)÷(WACC−g)", *v3(lambda s: f"=PL!{ref(s, 'nopat', 2032)}*(1+{G})*(1-{G}/$B${VR['ronic']})/({W_}-{G})"), fmt="#,##0.0")
vr("pvtv", "終值現值", *v3(lambda s: f"={'BCD'[ 'ABC'.index(s)]}{VR['tv']}/(1+{W_})^({ST}+6)"), fmt="#,##0.0")
vr("ev", "企業價值 EV", *v3(lambda s: f"={'BCD'['ABC'.index(s)]}{VR['pvx']}+{'BCD'['ABC'.index(s)]}{VR['pvtv']}"), fmt="#,##0.0")
vr("tvsh", "終值占 EV", *v3(lambda s: f'=IF({"BCD"["ABC".index(s)]}{VR["ev"]}>0,{"BCD"["ABC".index(s)]}{VR["pvtv"]}/{"BCD"["ABC".index(s)]}{VR["ev"]},"N/M")'), fmt="0.0%", note="分母≤0 顯示 N/M")
vr("eq", "股權價值＝EV＋現金（2026-06-30）− 有利息負債與租賃", *v3(lambda s: f"={'BCD'['ABC'.index(s)]}{VR['ev']}+cash_0-fin_lease"), fmt="#,##0.0")
vr("ps0", "DCF 每股價值（評價日；÷現流通股）", *v3(lambda s: f"={'BCD'['ABC'.index(s)]}{VR['eq']}*1000/sh_now"), fmt="#,##0")
vr("dcf", "DCF 目標價（推到目標時點、扣 FY2026 股利）", *v3(lambda s: f"={'BCD'['ABC'.index(s)]}{VR['ps0']}*(1+{KE})^{YT}-PL!{ref(s, 'dps', 2026)}"), fmt="#,##0")
vr("pm", "Comparable 倍數：同業 P/S 中位數（CHKP、GEN、QLYS、FTNT；TTM；GEN、QLYS 價格日 2026-10-26 逾期）×相對倍數 rel_mult（cm_adj）", "=ps_med*cm_adj", fmt="0.00", note="P/S 為股權口徑（不加淨現金、不扣債務；無非控制權益，未找到）；與 PS／PSG 相對值同口徑（G6）；CHKP、FTNT 價格日 2026-11-13，其餘逾期")
vr("ttm", "目標時點 TTM 營收（FY2027E 與 FY2026E 依日數加權；目標時點＝tp_date）", *v3(lambda s: f"=PL!{ref(s, 'rev', 2027)}*(tp_date-DATE(2026,12,31))/365+PL!{ref(s, 'rev', 2026)}*(DATE(2026,12,31)-EDATE(tp_date,-12))/365"), fmt="#,##0.0", note="¥B；相鄰會計年度依日數加權（G7 第 4 條）")
vr("evc", "Comparable 股權價值＝P/S 倍數×目標時點 TTM 營收", *v3(lambda s: f"=$B${VR['pm']}*{'BCD'['ABC'.index(s)]}{VR['ttm']}"), fmt="#,##0.0")
vr("comp", "Comparable 目標價（股權價值 ÷ 評價日稀釋股數 sh_now；無 RSU 與選擇權揭露，等於流通股）", *v3(lambda s: f"={'BCD'['ABC'.index(s)]}{VR['evc']}*1000/sh_now"), fmt="#,##0")
vr("tps", "各情境加權目標價（w_dcf×DCF＋(1−w_dcf)×Comparable）", *v3(lambda s: f"=w_dcf*{'BCD'['ABC'.index(s)]}{VR['dcf']}+(1-w_dcf)*{'BCD'['ABC'.index(s)]}{VR['comp']}"), fmt="#,##0")
vr("wts", "情境權重 A／B／C", "=tp_wA", "=tp_wB", "=tp_wC", fmt="0%")
vr("tpdcf", "DCF 目標價（情境加權）", "=SUMPRODUCT(B%d:D%d,B%d:D%d)" % (VR["dcf"], VR["dcf"], VR["wts"], VR["wts"]), fmt="#,##0")
vr("tpcomp", "Comparable 目標價（情境加權）", "=SUMPRODUCT(B%d:D%d,B%d:D%d)" % (VR["comp"], VR["comp"], VR["wts"], VR["wts"]), fmt="#,##0")
vr("tp", "加權目標價（DCF 與 Comparable 加權；目標時點 12 個月）", "=SUMPRODUCT(B%d:D%d,B%d:D%d)" % (VR["tps"], VR["tps"], VR["wts"], VR["wts"]), fmt="#,##0")
vr("cons", "市場共識目標價（9 家平均；S&P 經 stockanalysis，價格日 2026-09-30，財報前，只作歷史對照）", "=cons_tp", fmt="#,##0")
vr("px", "現價（2026-11-13 收盤）", "=px", fmt="#,##0")
vr("up", "加權目標價上檔空間（持有期間 12 個月，非年化）", "=B%d/B%d-1" % (VR["tp"], VR["px"]), fmt="0.0%")
vr("upd", "DCF 單法上檔空間", "=B%d/B%d-1" % (VR["tpdcf"], VR["px"]), fmt="0.0%")
vr("upc", "Comparable 單法上檔空間", "=B%d/B%d-1" % (VR["tpcomp"], VR["px"]), fmt="0.0%")
def recf(u): return f'IF({u}>0.5,"強力買進",IF({u}>0.25,"買進",IF({u}>=-0.25,"中立",IF({u}>=-0.5,"放空","強力放空"))))'
vr("vrate", "估值評等（只看上檔空間；> +50%、> +25%、≥ −25%、≥ −50%）", "=" + recf("B%d" % VR["up"]), fmt="@")
vr("gate", "入不敷出判定用（基準情境 B；FY2026–FY2030 自由現金流扣股利合計，¥B）", "=CF!C%d" % GATE, fmt="#,##0.0")
vr("gatej", "入不敷出判定", '=IF(B%d<0,"不通過","通過")' % VR["gate"], fmt="@")
vr("rec", "投資建議（收支不足＝入不敷出且估值評等為買進以上）", '=IF(AND(B%d<0,OR(B%d="買進",B%d="強力買進")),"收支不足",B%d)' % (VR["gate"], VR["vrate"], VR["vrate"], VR["vrate"]), fmt="@")
vr("dist", "距最近評等邊界（百分點；<5 要揭露）", "=MIN(ABS(B%d-0.5),ABS(B%d-0.25),ABS(B%d+0.25),ABS(B%d+0.5))*100" % ((VR["up"],) * 4), fmt="0.0")
# SBC 加回對照（只作對照，不進加權目標價）
def pvab(sc):
    t = [f"sbc_ab*PL!{ref(sc, 'rev', 2026)}*{ST}/(1+{W_})^({ST}/2)"]
    for y in range(2027, 2033): t.append(f"sbc_ab*PL!{ref(sc, 'rev', y)}/(1+{W_})^({ST}+{y - 2027}+0.5)")
    t.append(f"sbc_ab*PL!{ref(sc, 'rev', 2032)}*(1+{G})/({W_}-{G})/(1+{W_})^({ST}+6)")
    return "=" + "+".join(t)
vr("pvab", "SBC 加回對照：股票薪酬加回的現值（FY2026 下半年到終值；sbc_ab×營收；終值為永續年金，不乘稅盾）", *v3(pvab), fmt="#,##0.00", note="只作對照，不進加權目標價；合併 SBC 未找到，sbc_ab 僅董事選擇權（下限）")
vr("dcfs", "DCF 每股目標價（股票薪酬加回；對照）", *v3(lambda s: f"=(({'BCD'['ABC'.index(s)]}{VR['eq']}+{'BCD'['ABC'.index(s)]}{VR['pvab']})*1000/sh_now)*(1+{KE})^{YT}-PL!{ref(s, 'dps', 2026)}"), fmt="#,##0", note="同 DCF 目標價的推算方式，EV 加上加回現值")
vr("ps0s", "DCF 每股價值（股票薪酬加回；評價日；對照）", *v3(lambda s: f"=({'BCD'['ABC'.index(s)]}{VR['eq']}+{'BCD'['ABC'.index(s)]}{VR['pvab']})*1000/sh_now"), fmt="#,##0")
vr("tpdcfs", "DCF 目標價（股票薪酬加回；情境加權；對照，不進加權目標價）", "=SUMPRODUCT(B%d:D%d,B%d:D%d)" % (VR["dcfs"], VR["dcfs"], VR["wts"], VR["wts"]), fmt="#,##0")
vr("sbcd", "  加回後比 DCF 目標價多（日圓）", "=B%d-B%d" % (VR["tpdcfs"], VR["tpdcf"]), fmt="#,##0.0", note="差距極小，因合併 SBC 未找到、僅董事選擇權為下限")
# DCF 與 Comparable 落差分解（情境 B）
vr("gap", "落差分解（情境 B）：DCF 目標價 − Comparable 目標價", "=C%d-C%d" % (VR["dcf"], VR["comp"]), fmt="#,##0")
vr("gap_t", "  時點與股利：DCF 由評價日推到目標時點並扣股利的增減", "=C%d-C%d" % (VR["dcf"], VR["ps0"]), fmt="#,##0")
vr("mdcf", "  DCF 隱含 P/S（DCF 股權價值 ÷ 目標時點 TTM 營收）", "=C%d/C%d" % (VR["eq"], VR["ttm"]), fmt="0.00")
vr("gap_m", "  倍數：(DCF 隱含 P/S − Comparable P/S)×目標時點 TTM 營收 ÷ 股數", "=(B%d-B%d)*C%d*1000/sh_now" % (VR["mdcf"], VR["pm"], VR["ttm"]), fmt="#,##0")
vr("gap_s", "  股票薪酬與槓桿：基準 0（sbc_r＝0、無負債）", 0, fmt="#,##0")
vr("gap_c", "  分解核對（時點＋倍數＋其他−落差，應為 0）", "=B%d+B%d+B%d-B%d" % (VR["gap_t"], VR["gap_m"], VR["gap_s"], VR["gap"]), fmt="0.000")
vr("cur", "本公司現行 P/S（2026-11-13）", "=ps_tm", fmt="0.00", note="TTM P/S；與同業 P/S 同價格日對照")
vr("rel", "現行倍數 ÷ 同業 P/S 中位數（本公司現行相對值；對照 rel_mult）", "=ps_tm/ps_med", fmt="0.00", note="與 rel_mult（歷史 20 季 PS／PSG 讀數）並列；現行值高於 rel_mult 表示市場已不把成長差全數計入")
vr("pmx", "對照列：同業 EV/EBIT 中位數（TTM；CHKP、GEN、QLYS；GEN、QLYS 價格日 2026-10-26 逾期）", "=MEDIAN(pm_chkp,pm_gen,pm_qlys)", fmt="0.00", note="盈餘型倍數已內含利潤率差，與 PS／PSG 相對值不同口徑，只作對照，不進加權目標價（G6）")
vr("curx", "對照列：本公司現行 EV/EBIT（現價 EV ÷ FY2026E 營業利益，情境 B）", "=(px*sh_now/1000-cash_0+fin_lease)/PL!%s" % ref("B", "oi", 2026), fmt="0.00", note="分子分母同為 FY2026E，與同業 TTM 倍數並列")
vr("relx", "對照列：現行 EV/EBIT ÷ 同業中位數（本公司現行相對值）", "=B%d/B%d" % (VR["curx"], VR["pmx"]), fmt="0.00")
vr("cx1", "對照列：EV/EBIT Comparable 目標價，rel_mult＝1.00（同業中位 × FY2027E 營業利益＋現金−負債）", *v3(lambda s: f"=($B${VR['pmx']}*PL!{ref(s, 'oi', 2027)}+cash_0-fin_lease)*1000/sh_now"), fmt="#,##0")
vr("cx2", "對照列：EV/EBIT Comparable 目標價，rel＝本公司現行相對值", *v3(lambda s: f"=($B${VR['pmx']}*$B${VR['relx']}*PL!{ref(s, 'oi', 2027)}+cash_0-fin_lease)*1000/sh_now"), fmt="#,##0")
vr("cx1w", "對照列：EV/EBIT Comparable（rel＝1.00）情境加權；不進加權目標價", "=SUMPRODUCT(B%d:D%d,B%d:D%d)" % (VR["cx1"], VR["cx1"], VR["wts"], VR["wts"]), fmt="#,##0")
vr("cx2w", "對照列：EV/EBIT Comparable（rel＝現行）情境加權；不進加權目標價", "=SUMPRODUCT(B%d:D%d,B%d:D%d)" % (VR["cx2"], VR["cx2"], VR["wts"], VR["wts"]), fmt="#,##0")
vr("stale", "同業倍數日期（逾期標示）", '="P/S：CHKP、FTNT 2026-11-13（未逾期）；GEN、QLYS 2026-10-26（逾期，相差超過 2 個交易日）；對照案資料庫季末 2026-06-30（逾期）。對照列 EV/EBIT：CHKP 2026-11-13、GEN 與 QLYS 2026-10-26（逾期）"', fmt="@")
VA.column_dimensions["A"].width = 80; VA.column_dimensions["E"].width = 80
for c in "BCD": VA.column_dimensions[c].width = 14
# ================= RevDCF =================
RV["A1"] = "RevDCF — 反向 DCF：現價隱含什麼（情境 B 路徑；FY2028–FY2032 營業利益率統一一個值，其餘不變；封閉解，FCFF 對營業利益線性）"
RV["A2"] = "把現價折回評價日得到市場隱含 EV，反解 FY2028–FY2032 統一營業利益率；Comparable 側則列現價隱含倍數。反解值是「目標價＝現價」的點，不是評等門檻。"
RVR = {}
def rv(key, label, f, fmt="#,##0.00", note=""):
    rr = 4 + len(RVR); RVR[key] = rr; RV.cell(rr, 1, label); RV.cell(rr, 2, f).number_format = fmt
    if note: RV.cell(rr, 3, note)
W = f"Val!$B${VR['wacc']}"; S_ = f"Val!$B${VR['stub']}"
def dfa(y): return f"(1+{W})^({S_}+{y - 2027}+0.5)"
rv("evreq", "市場隱含 EV（現價市值折回評價日 − 現金＋負債）", f"=(px*sh_now/1000)/(1+Val!$B${VR['ke']})^((px_date-val_date)/365)-cash_0+fin_lease", "#,##0.0", "¥B；未扣期間股利（略高估隱含 EV 約一次股利）")
rv("c1", "每 1.00 營業利益率的現值（FY2028–FY2032＋終值）", "=" + "+".join(f"PL!{ref('B', 'rev', y)}*(1-te)/{dfa(y)}" for y in range(2028, 2033)) + f"+PL!{ref('B', 'rev', 2032)}*(1+g_tv)*(1-te)*(1-g_tv/Val!$B${VR['ronic']})/({W}-g_tv)/(1+{W})^({S_}+6)", "#,##0.0")
rv("pvoi", "模型 FY2028–FY2032 營業利益相關現值（含終值）", "=" + "+".join(f"(PL!{ref('B', 'oi', y)}-sbc_r*PL!{ref('B', 'rev', y)})*(1-te)/{dfa(y)}" for y in range(2028, 2033)) + f"+PL!{ref('B', 'nopat', 2032)}*(1+g_tv)*(1-g_tv/Val!$B${VR['ronic']})/({W}-g_tv)/(1+{W})^({S_}+6)", "#,##0.0")
rv("rest", "其餘現值（FY2026–2027 與折舊、資本支出、營運資金）", f"=Val!$C${VR['ev']}-B{4 + 2}", "#,##0.0")
rv("m1", "現價隱含 FY2028–FY2032 統一營業利益率", f"=(B{4}-B{4 + 3})/B{4 + 1}+sbc_r", "0.0%")
rv("m0", "模型情境 B FY2028–FY2032 營收加權營業利益率", "=(" + "+".join(f"PL!{ref('B', 'oi', y)}" for y in range(2028, 2033)) + ")/(" + "+".join(f"PL!{ref('B', 'rev', y)}" for y in range(2028, 2033)) + ")", "0.0%")
rv("gapm", "現價隱含 − 模型（百分點）", f"=(B{4 + 4}-B{4 + 5})*100", "0.0")
rv("mdc", "現價隱含 EV/EBIT（現價 EV ÷ FY2027E 營業利益，情境 B）", f"=(px*sh_now/1000-cash_0+fin_lease)/PL!{ref('B', 'oi', 2027)}", "0.00", "對照：EV/EBIT 對照列見 Val 頁")
rv("sum", "摘要（公式組字）", f'="現價 "&TEXT(px,"#,##0")&" 日圓隱含 FY2028–FY2032 營業利益率約 "&TEXT(B{4 + 4},"0.0%")&"，模型情境 B 為 "&TEXT(B{4 + 5},"0.0%")&"（"&IF(B{4 + 4}>B{4 + 5},"現價要求高於模型","現價要求低於模型")&" "&TEXT(ABS(B{4 + 6}),"0.0")&" 個百分點）；管理層 FY2028 目標 28–31%（Interested-party，不進模型）。"', "@")
# ================= Recon =================
RC["A1"] = "Recon — 對帳：公司指引、前三季實績、共識（財報前）、分部與單位經濟。比較字樣由公式依數值組字。"
RCR = {}; rcur = [3]
def rch(title, cols):
    rcur[0] += 1; RC.cell(rcur[0], 1, title).font = BOLD; rcur[0] += 1; hdr(RC, rcur[0], cols); rcur[0] += 1
def rc(key, *vals, fmts=None):
    rr = rcur[0]; rcur[0] += 1; RCR[key] = rr
    for j, v in enumerate(vals, 1):
        if v is not None:
            c = RC.cell(rr, j, v)
            if fmts and j - 1 < len(fmts) and fmts[j - 1]: c.number_format = fmts[j - 1]
    return rr
rch("1. 合併層：模型 FY2026E vs 公司指引（Interested-party）與前三季實績", ["項目", "模型情境 B", "公司指引", "差", "說明"])
rc("r26", "FY2026E 營收（¥B）", f"=PL!{ref('B', 'rev', 2026)}", 301.5, "=B{r}-C{r}", "指引 2026-02-18 發布、11-12 維持；前三季實績 223.9，隱含 Q4＝模型減前三季", fmts=[None, "#,##0.0", "#,##0.0", "#,##0.0"])
rc("o26", "FY2026E 營業利益（¥B）", f"=PL!{ref('B', 'oi', 2026)}", 44.4, "=B{r}-C{r}", "指引 2026-08-13 由 56.4 下修；模型＝前三季實績 35.0＋Q4×Q4 利益率（Assumed，未反推指引）", fmts=[None, "#,##0.0", "#,##0.0", "#,##0.0"])
rc("n26", "FY2026E 歸屬淨利（¥B）", f"=PL!{ref('B', 'ni', 2026)}", 30.7, "=B{r}-C{r}", "", fmts=[None, "#,##0.0", "#,##0.0", "#,##0.0"])
rc("m28", "FY2028E 營業利益率（情境 B）vs 管理層目標 28–31%", f"=PL!{ref('B', 'oi', 2028)}/PL!{ref('B', 'rev', 2028)}", 0.28, "=(B{r}-C{r})*100", "目標低緣；差為百分點；管理層口頭目標（Interested-party），不進模型、不反推", fmts=[None, "0.0%", "0.0%", "0.0"])
rc("gm", "FY2026E 毛利率 vs 2026Q3 近四季 77.4%（資料商）", f"=PL!{ref('B', 'gp', 2026)}/PL!{ref('B', 'rev', 2026)}", 0.774, "=(B{r}-C{r})*100", "模型含推論成本、營業成本率持平於 FY2025；差為百分點", fmts=[None, "0.0%", "0.0%", "0.0"])
rc("c26", "FY2026E 營收 vs 賣方共識 298.96（財報前，只作歷史對照）", f"=PL!{ref('B', 'rev', 2026)}", 298.96, "=B{r}/C{r}-1", "S&P 經 stockanalysis，更新 2026-09-11", fmts=[None, "#,##0.0", "#,##0.0", "0.0%"])
rc("c27", "FY2027E 營收 vs 賣方共識 316.69（財報前）", f"=PL!{ref('B', 'rev', 2027)}", 316.69, "=B{r}/C{r}-1", "EPS 共識：未找到", fmts=[None, "#,##0.0", "#,##0.0", "0.0%"])
for k in ("r26", "o26", "n26", "m28", "gm", "c26", "c27"):
    rr = RCR[k]
    for c in range(1, 5):
        v = RC.cell(rr, c).value
        if isinstance(v, str): RC.cell(rr, c).value = v.replace("{r}", str(rr))
rch("2. 分部（所在地別）：2026H1 實績 vs 模型 FY2026E", ["分部", "H1 實績（¥B）", "H1 年增", "模型 FY2026E", "隱含 H2（¥B）", "隱含 H2 年增", "FY2025 實績"])
for k, z in REG:
    rr = rc("h1" + k, z, round(fv(f"seg.{z}.h1fy2026") / 1000, 3), {"Japan": 0.007, "Americas": 0.075, "Europe": 0.166, "Asia Pacific": 0.222}[z],
            f"=PL!{ref('B', 'r_' + k, 2026)}", None, None, f"={HC('s_' + z, 2025)}", fmts=[None, "#,##0.0", "0.0%", "#,##0.0", "#,##0.0", "0.0%", "#,##0.0"])
    RC.cell(rr, 5, f"=D{rr}-B{rr}").number_format = "#,##0.0"
    RC.cell(rr, 6, f"=E{rr}/(G{rr}-B{rr}/(1+C{rr}))-1").number_format = "0.0%"
rc("h1t", "合計", f"=SUM(B{RCR['h1jp']}:B{RCR['h1ap']})", None, f"=SUM(D{RCR['h1jp']}:D{RCR['h1ap']})", f"=SUM(E{RCR['h1jp']}:E{RCR['h1ap']})", None, f"=SUM(G{RCR['h1jp']}:G{RCR['h1ap']})", fmts=[None, "#,##0.0", None, "#,##0.0", "#,##0.0", None, "#,##0.0"])
rcur[0] += 0
rch("3. 原型 P 單位經濟（每 US$1M Vision One ARR 的年度數，US$K；情境 A／B／C，FY2027E／FY2030E）", ["項目", "情境 A", "情境 B", "情境 C", "說明"])
for y in (2027, 2030):
    for key, lab in (("inc", "AI 營收增量"), ("ero", "單價侵蝕"), ("inf", "AI 推論成本"), ("ain", "AI 淨效果")):
        rc(f"u{key}{y}", f"FY{y}E 每 US$1M 平均 ARR 的{lab}（US$K）",
           *[f"=PL!{ref(s, key, y)}/PL!{ref(s, 'v1', y)}*1000" for s in "ABC"],
           "¥B 換 US$：÷fx；除以平均 ARR（US$M）再 ×1,000", fmts=[None, "#,##0", "#,##0", "#,##0"])
rc("uinfr", "FY2030E 推論成本占 Vision One 營收", *[f"=PL!{ref(s, 'inf', 2030)}/PL!{ref(s, 'v1', 2030)}" for s in "ABC"], "推論成本占 Vision One 營收", fmts=[None, "0.0%", "0.0%", "0.0%"])
rc("uain", "AI 淨效果（¥B）FY2030E", *[f"=PL!{ref(s, 'ain', 2030)}" for s in "ABC"], "三情境分化：A＝增量大、侵蝕小；C 相反", fmts=[None, "#,##0.0", "#,##0.0", "#,##0.0"])
rcur[0] += 1
RC.cell(rcur[0], 1, "4. 反解（整本重算）：見 Sens 頁下方「反解」區（建置時由敏感度重算寫入）"); RCR["kreq_hint"] = rcur[0]
rcur[0] += 2
rch("5. DCF（股票薪酬加回）對照：只作對照，不進加權目標價", ["項目", "情境 A", "情境 B", "情境 C", "情境加權", "說明"])
rc("sb0", "DCF 每股價值（評價日；現行）", *[f"=Val!{c}{VR['ps0']}" for c in "BCD"], None, "現行 DCF，sbc_r＝0", fmts=[None, "#,##0", "#,##0", "#,##0"])
rc("sb1", "DCF 每股價值（評價日；股票薪酬加回）", *[f"=Val!{c}{VR['ps0s']}" for c in "BCD"], None, "加回現值見 Val 頁；sbc_ab 為董事選擇權下限", fmts=[None, "#,##0", "#,##0", "#,##0"])
rc("sb2", "DCF 目標價（現行）", *[f"=Val!{c}{VR['dcf']}" for c in "BCD"], f"=Val!B{VR['tpdcf']}", "", fmts=[None, "#,##0", "#,##0", "#,##0", "#,##0"])
rc("sb3", "DCF 目標價（股票薪酬加回）", *[f"=Val!{c}{VR['dcfs']}" for c in "BCD"], f"=Val!B{VR['tpdcfs']}", "對照；不進加權目標價", fmts=[None, "#,##0", "#,##0", "#,##0", "#,##0"])
rc("sbs", "摘要（公式組字）", f'="DCF 目標價（股票薪酬加回）"&TEXT(Val!B{VR["tpdcfs"]},"#,##0")&" 日圓，較現行 DCF "&TEXT(Val!B{VR["tpdcf"]},"#,##0")&" 日圓多 "&TEXT(Val!B{VR["sbcd"]},"#,##0.0")&" 日圓；合併股票薪酬未找到，加回只含董事選擇權（占營收 "&TEXT(sbc_ab,"0.000%")&"，下限），故差距極小；只作對照，不進加權目標價。"')
RC.column_dimensions["A"].width = 70
for c in "BCDEFG": RC.column_dimensions[c].width = 16
# ================= Sens =================
SENS = json.load(open(A.sens, encoding="utf-8")) if A.sens and os.path.exists(A.sens) else {}
cases = []
for row_ in ROWS:
    if row_[0] == "sec": continue
    key, item, unit, lo, base, hi = row_[:6]
    if key in ("cons_tp", "tp_wA", "tp_wB", "tp_wC") or key.endswith(("_A", "_B", "_C")): continue
    for side, v in (("低", lo), ("高", hi)):
        if isinstance(v, (int, float)) and isinstance(row_[4], (int, float, str)) and v != (base if isinstance(base, (int, float)) else None):
            cases.append((key, side, [(key, v)], item))
EXTRA = [("combo_route", "低", [("ptok", TKI["IF_PriceRef_Luna"]["values"]), ("tok_mult", tm[3])], "聯動：改走便宜模型（token 市價取低、每任務 token 取高）"),
         ("combo_pxcost", "高", [("ptok", TKI["IF_PriceRef_Sol"]["values"]), ("defl", 0.05)], "聯動：token 市價高且降幅小（價格與降幅同向）"),
         ("peer_ps_db", "低", [("ps_med", PPG["db_last_quarter"]["median"])], "同業 P/S＝資料庫最近一季中位（對照，G7 第 2 條）"),
         ("rel_ps_only", "高", [("cm_adj", RM["ps_reading"])], "相對倍數只用 PS 讀數（ps_only）"),
         ("rel_psg_only", "低", [("cm_adj", RM["psg_reading"])], "相對倍數只用 PSG 讀數（psg_only）"),
         ("combo_ai_off", "低", [("ag26", 0.0), ("am26", 0.0), ("tpm", 0.0)], "AI 全關：無 AI 增量也無推論成本（對照）")]
cases += EXTRA
SE["A1"] = "Sens — 敏感度：一次改一個 In 頁輸入（敏感度低／高，整本以 LibreOffice 重算），讀加權目標價。token 市價高案只動成本一邊，標「上界」；聯動案另列。"
SE["A2"] = "豁免：cons_tp 不列：只作對照，不進模型，不影響任何目標價；kag_A 等情境專屬列不列：已由情境 A／B／C 加權涵蓋；vsh_jp 不列：分配類參數，不改合併營收與目標價。"
SE["A3"] = "讀法：相對基準＝該案加權目標價 ÷ 基準 − 1；基準檔 In 指紋與模型狀態見下方。"
rr = 5
SE.cell(rr, 1, "基準加權目標價（建置時敏感度重算）"); SE.cell(rr, 4, SENS.get("base|base", {}).get("tp")); SE_BASE = rr; rr += 1
SE.cell(rr, 1, "基準檔 In 輸入指紋（敏感度重算時）"); SE.cell(rr, 4, SENS.get("base|base", {}).get("fp", "")); SE_FPR = rr; rr += 1
SE.cell(rr, 1, "本檔 In 輸入指紋（即時）"); SE.cell(rr, 4, FPR); SE_FPC = rr; rr += 1
SE.cell(rr, 1, "模型狀態"); SE_ST = rr; rr += 1
SE.cell(rr, 1, "build_id"); SE.cell(rr, 4, BUILD_ID); SE_BID = rr; rr += 2
hdr(SE, rr, ["驅動", "設定", "方向", "加權目標價", "相對基準", "DCF 目標價", "DCF 變動（日圓）", "投資建議", "收支關卡 B", "鍵"]); SE_HDR = rr; rr += 1
SE_ROWS = []
for key, side, sets, item in cases:
    res = SENS.get(f"{key}|{side}")
    SE.cell(rr, 1, item + ("（上界：只動成本一邊）" if key == "ptok" and side == "高" else "")); SE.cell(rr, 2, "、".join(f"{k}＝{v:g}" for k, v in sets)); SE.cell(rr, 3, "輸入取" + side + "值")
    if res:
        SE.cell(rr, 4, res["tp"]).number_format = "#,##0"; SE.cell(rr, 5, f"=D{rr}/$D${SE_BASE}-1").number_format = "0.0%"
        SE.cell(rr, 6, res["tpdcf"]).number_format = "#,##0"; SE.cell(rr, 7, f"=F{rr}-{SENS['base|base']['tpdcf']}").number_format = "#,##0"
        SE.cell(rr, 8, res["rec"]); SE.cell(rr, 9, res["gate"]).number_format = "#,##0.0"
    SE.cell(rr, 10, "、".join(k for k, _ in sets)); SE_ROWS.append((key, side, rr, item)); rr += 1
SE_END = rr - 1
rr += 1
SE.cell(rr, 1, "加權目標價相對基準（HTML 龍捲風圖用；低案與高案）").font = BOLD; SE_T0 = rr; rr += 1
hdr(SE, rr, ["驅動（白話）", "輸入取低值", "輸入取高值"]); rr += 1
SE_T1 = rr
seen = []
for key, side, srow, item in SE_ROWS:
    if key in seen: continue
    seen.append(key)
    SE.cell(rr, 1, item[:60] + ("（上界）" if key == "ptok" else ""))
    for col, sd in ((2, "低"), (3, "高")):
        m = next((x for x in SE_ROWS if x[0] == key and x[1] == sd), None)
        if m: SE.cell(rr, col, f"=E{m[2]}").number_format = "0.0%"
    rr += 1
SE_T2 = rr - 1
rr += 1
SE.cell(rr, 1, "反解（整本重算；支持現價所需值與評等翻轉值；建置時由 sens_trendmicro.py 寫入）").font = BOLD; SE_K0 = rr; rr += 1
hdr(SE, rr, ["變數", "支持現價（目標價＝現價）所需值", "基準值", "評等升一級邊界（上檔空間 +50%，目標價＝1.5×現價）所需值", "評等降一級邊界（上檔空間 +25%，目標價＝1.25×現價）所需值"]); rr += 1
KR = SENS.get("kreq", {})
SE_KR = {}
for kk, lab in (("s_b26", "非 AI 受保護範圍成長（FY2026）"), ("tpm", "每 US$1M ARR 每年 AI 任務數"), ("fx", "日圓兌美元")):
    SE.cell(rr, 1, lab); SE_KR[kk] = rr
    v = KR.get(kk)
    if v:
        for c, x in zip((2, 4, 5), (v.get("px"), v.get("up"), v.get("down"))):
            if x is not None: SE.cell(rr, c, x).number_format = "0.000"
    SE.cell(rr, 3, f"={kk}"); rr += 1
SE.column_dimensions["A"].width = 70
# 模型狀態（公式）
SE.cell(SE_ST, 4, f'=IF(AND(D{SE_FPR}=D{SE_FPC},ABS(Val!B{VR["tp"]}-D{SE_BASE})<0.5),"CURRENT","STALE")')
# ================= README =================
def T(c, f="#,##0"): return f'TEXT({c},"{f}")'
VB = lambda k: f"Val!B{VR[k]}"
PBR = lambda k, y: f"PL!{ref('B', k, y)}"
RDR = {}; rrd = [1]
def rd(key, text, formula=True):
    r_ = rrd[0]; rrd[0] += 1; RDR[key] = r_
    RD.cell(r_, 1, ("=" + text) if formula else text); return r_
rd("a1", f'"趨勢科技（4704）個股模型 {VER}｜build_id {BUILD_ID}｜資料截止日 "&TEXT(data_date,"yyyy-mm-dd")&"｜成品日 {VDATE}｜股價基準日 "&TEXT(px_date,"yyyy-mm-dd")&"｜目標時點 "&TEXT(tp_date,"yyyy-mm-dd")&"｜本版未做先驗"')
rd("sp", '""')
rd("thesis", '"命題：AI 進入資安軟體後，Trend Micro 的價值取決於 Vision One 的受保護範圍（含代理與非人類身分）、每範圍模組數與每模組單價三項乘積，能否蓋過單價侵蝕與推論成本；市場同時評價舊產品萎縮與管理層 FY2028 利益率目標。"')
rd("rec", f'"投資建議："&{VB("rec")}&"（估值評等 "&{VB("vrate")}&"；收支合計 "&{T(VB("gate"), "#,##0.0")}&" 十億日圓，"&{VB("gatej")}&"）。條件式：在 Vision One 範圍、模組、單價成長依基準衰減、推論成本依 Tokenomics 市價與任務量假設的前提下成立；AI 增量被侵蝕吃掉或推論成本超出假設即失效。"')
for i, t in enumerate(("Vision One 的營收＝受保護範圍 × 每範圍模組數 × 每模組單價（成長形式；公司未揭露單位數，以 ARR 代理）。",
                       "AI 淨效果＝AI 營收增量 − 單價侵蝕 − 推論成本，三情境分化。",
                       "官方報告分部是所在地別四區（日本、美洲、歐洲、亞太）；Vision One 是各分部之下的命題子列，加上其他營收等於分部。",
                       "價值由 DCF 與 Comparable 各半，並看收支關卡。"), 1):
    rd(f"p{i}", f'"{t}"')
rd("ans1", f'"Q：Vision One 能否撐起成長？答：模型情境 B 下 Vision One 營收 FY2026E 為 "&{T(PBR("v1", 2026), "#,##0")}&" 十億日圓、FY2030E 為 "&{T(PBR("v1", 2030), "#,##0")}&"，合併營收 FY2030E 為 "&{T(PBR("rev", 2030), "#,##0")}&"；模型輸出，非獨立證據。"')
rd("ans2", f'"Q：AI 推論成本吃掉多少？答：FY2030E 情境 A／B／C 的 AI 淨效果為 "&{T("PL!"+ref("A","ain",2030), "#,##0.0")}&"／"&{T("PL!"+ref("B","ain",2030), "#,##0.0")}&"／"&{T("PL!"+ref("C","ain",2030), "#,##0.0")}&" 十億日圓，推論成本占 Vision One 營收（情境 B）"&{T("PL!"+ref("B","inf",2030)+"/PL!"+ref("B","v1",2030), "0.0%")}&"。"')
rd("ans3", f'"Q：現價要求什麼？答："&RevDCF!B{RVR["sum"]}')
rd("o1", f'"1. 營收：FY2026E "&{T(PBR("rev", 2026), "#,##0.0")}&" 十億日圓（指引 301.5），FY2028E "&{T(PBR("rev", 2028), "#,##0.0")}&"，FY2030E "&{T(PBR("rev", 2030), "#,##0.0")}&"；Vision One 貢獻成長，舊產品萎縮（以上為情境 B 模型輸出）。"')
rd("o2", f'"2. 利益率：FY2026E 營業利益 "&{T(PBR("oi", 2026), "#,##0.0")}&" 十億日圓（指引 44.4，已因 AI token 與雲端成本下修），FY2028E 營業利益率 "&{T(PBR("oi", 2028)+"/"+PBR("rev", 2028), "0.0%")}&"，"&IF({PBR("oi", 2028)}/{PBR("rev", 2028)}<0.28,"低於","不低於")&"管理層目標 28%（口頭目標，不進模型）。"')
rd("o3", f'"3. 估值：DCF "&{T(VB("tpdcf"))}&"、Comparable "&{T(VB("tpcomp"))}&"、加權 "&{T(VB("tp"))}&" 日圓；現價 "&{T(VB("px"))}&"（"&TEXT(px_date,"yyyy-mm-dd")&"）；DCF 與 Comparable 落差 "&{T(VB("gap"))}&" 日圓，時點與倍數可完全分解。"')
rd("o4", f'"4. 收支：FY2026–FY2030 自由現金流扣股利合計 "&{T(VB("gate"), "#,##0.0")}&" 十億日圓，關卡"&{VB("gatej")}&"；無有利息負債，現金 "&{T("cash_0", "#,##0")}&" 十億日圓。"')
rd("o5", f'"5. 傳聞：2025 年私募興趣報導未見後續成交，公司仍上市（2026-11-12 發布 Q3 財報）；模型不含任何收購溢價。"')
rd("hero", f'"加權目標價 "&{T(VB("tp"))}&" 日圓，上檔空間 "&{T(VB("up"), "+0.0%;-0.0%")}&"（現價 "&{T(VB("px"))}&"，"&TEXT(px_date,"yyyy-mm-dd")&" 收盤）；DCF "&{T(VB("tpdcf"))}&"、Comparable "&{T(VB("tpcomp"))}&"。Vision One 高成長，但推論成本已使 FY2026 營業利益指引下修；結論依賴 AI 增量不被單價侵蝕蓋過。"')
rd("diag", f'"發布前合理性診斷（G3）：DCF 目標價 ÷ 現價 "&{T(VB("tpdcf")+"/"+VB("px"), "0.00")}&" 倍（"&IF(AND({VB("tpdcf")}/{VB("px")}>=0.25,{VB("tpdcf")}/{VB("px")}<=4),"在 0.25–4 倍內","超出 0.25–4 倍，須診斷")&"）；EV "&IF(Val!C{VR["ev"]}>0,"為正","為負")&"；評等 "&{VB("vrate")}&IF(OR({VB("vrate")}="強力買進",{VB("vrate")}="強力放空"),"，須診斷","，非強力評等")&"；FY2027E 營收較賣方共識（財報前）"&TEXT({PBR("rev", 2027)}/316.69-1,"+0.0%;-0.0%")&"，FY2026E 營收較公司指引"&TEXT({PBR("rev", 2026)}/301.5-1,"+0.0%;-0.0%")&"。診斷結論：Comparable（P/S 主腿）目標價 "&{T(VB("tpcomp"))}&" 日圓，DCF "&{T(VB("tpdcf"))}&" 日圓，落差 "&{T(VB("gap"))}&" 日圓，由倍數主導（Comparable P/S "&{T(VB("pm"), "0.00")}&" 倍，DCF 隱含 P/S "&{T(VB("mdcf"), "0.00")}&" 倍）；倍數低是因 rel_mult "&TEXT(cm_adj,"0.00")&"（PS 與 PSG 讀數分歧，成長比 "&TEXT(rm_gr,"0.00")&"），而 DCF 的 WACC 約 "&TEXT({VB("wacc")},"0.0%")&" 配 Vision One 成長與費用年增低於營收年增（利益率擴張）。屬估值方法間的經濟判斷與讀數分歧，非建模錯誤；FY2027 公司無指引。"')
rd("dist", f'IF({VB("dist")}<5,"評等距最近邊界僅 "&{T(VB("dist"), "0.0")}&" 個百分點，目標價小幅變動即翻轉評等。","評等距最近邊界 "&{T(VB("dist"), "0.0")}&" 個百分點。")')
rd("drv", f'"驅動：Vision One 營收＝平均 ARR × 匯率；ARR 年增＝(1＋範圍成長)×(1＋模組成長)×(1＋單價成長)，AI 增量加在範圍與模組、侵蝕扣在單價；各分部營收＝Vision One 占比分配＋其他營收，合計＝合併營收（核對列為 0）。"')
rd("ver", f'"{VER}（{VDATE}）：依 G7 同業 P/S 改讀 2026-11-13 現價（逾期者標示）、Comparable 分母改目標時點 TTM 營收、Sens 加同業 P/S 資料庫季末案；v0.2 起：相對倍數改讀 relmult_hist 的 geo_base（"&TEXT(cm_adj,"0.00")&"，v0.1 為 1.00，Derived）；依 G6 Comparable 主腿改 P/S 口徑、EV/EBIT 改對照列；新增 Sens 四案 ps_only、psg_only、sens.low、sens.high；DCF（股票薪酬加回）對照列（不進加權目標價）；rf 沿用日本 10 年期 "&TEXT(rf,"0.00%")&"（2026-11-11）；其餘同 v0.1。本版未做先驗；未做 Meta 的算力、自用 AI 讀法 2、Wattonomics、三層協調（不適用：非算力公司、非 Siliconomics 覆蓋）。"')
rd("rm1", f'"相對倍數與 Comparable 口徑（G6）：rel_mult＝"&TEXT(cm_adj,"0.00")&"（PS 與 PSG 兩讀數幾何平均，來源 relmult_hist/out/trendmicro.json，Derived）。Comparable 主腿改 P/S 口徑，與 PS／PSG 相對值同口徑：同業 P/S 中位 "&TEXT(ps_med,"0.00")&"（2026-11-13 價；v0.2 為資料庫季末中位 "&TEXT(ps_db,"0.00")&"）× rel_mult "&TEXT(cm_adj,"0.00")&" ＝ "&TEXT({VB("pm")},"0.00")&" 倍，乘目標時點 TTM 營收（FY2027E 與 FY2026E 依日數加權）÷ 評價日稀釋股數；EV/EBIT Comparable 只作對照列（rel＝1.00 為 "&TEXT({VB("cx1w")},"#,##0")&"、rel＝現行 "&TEXT({VB("relx")},"0.00")&" 為 "&TEXT({VB("cx2w")},"#,##0")&" 日圓），不進加權目標價。Sens 另列同業 P/S＝資料庫最近一季中位案、ps_only "&TEXT(rm_ps,"0.00")&"、psg_only "&TEXT(rm_psg,"0.00")&"、sens.low、sens.high（In 頁 cm_adj 低／高欄）。"')
rd("rm2", f'IF(rm_div>0.3,"PS 與 PSG 兩讀數差異大（分歧度 "&TEXT(rm_div,"0%")&"，超過 30%）：PS 讀數 "&TEXT(rm_ps,"0.00")&"、PSG 讀數 "&TEXT(rm_psg,"0.00")&"。原因是成長比：本公司 TTM 營收年增 "&TEXT(rm_gco,"0.0%")&"，同業中位 "&TEXT(rm_gpe,"0.0%")&"，成長比 "&TEXT(rm_gr,"0.00")&"；PSG 讀數＝PSG 相對值中位數 "&TEXT(rm_psgm,"0.00")&" × 成長比 "&TEXT(rm_gr,"0.00")&" ≈ "&TEXT(rm_psgm*rm_gr,"0.00")&"。PS 讀數不扣成長差；PSG 把成長差全數計入，所以讀數約減半。geo_base 取兩者折衷。","PS 與 PSG 兩讀數分歧度 "&TEXT(rm_div,"0%")&"，未超過 30%。")')
rd("sp2", '""')
rd("h_md", '"市場資料日期"', formula=True)
rd("md1", f'"現價 "&TEXT(px_date,"yyyy-mm-dd")&"；rf 2026-11-11（"&IF(px_date-DATE(2026,11,11)>2,"逾期","未逾期")&"）；同業 P/S CHKP、FTNT 2026-11-13，GEN、QLYS 2026-10-26（逾期）；對照列 EV/EBIT：CHKP 2026-11-13、GEN 與 QLYS 2026-10-26（逾期）；共識目標價價格日 2026-09-30、資料更新 2026-09-11（財報前，只作歷史對照）。"')
rd("h_fail", '"主要失效條件"', formula=True)
rd("f1", '"Vision One ARR 年增掉到非 AI 基準以下，或單價侵蝕大於 AI 增量（觀察：公司每季 Vision One ARR 與淨新增 ARR）。"')
rd("f2", '"AI 推論成本占 Vision One 營收超過假設區間（觀察：營業利益指引是否再因雲端與 token 成本下修）。"')
rd("f3", '"營收分部中美洲持續萎縮、日本 ARR 持平（觀察：分部營收季度揭露）。"')
rd("h_risk", '"風險登錄"', formula=True)
RISK_HDR = rrd[0]
for t in (("風險", "是否建模", "理由", "觀察指標"),
          ("私募收購或私有化", "否", "傳聞未成交，建模加溢價屬猜測", "TSE 公告、EDINET 公開收購申報"),
          ("客戶資安事件（GMO 事件）", "否", "金額與責任未揭露", "公司說明與訴訟申報"),
          ("匯率（日圓兌美元）", "是（Sens：fx）", "ARR 以美元計價、營收以日圓報", "季度匯率假設"),
          ("同業競爭與價格戰", "是（單價侵蝕）", "Q2 公告提競爭環境", "每模組單價與 NRR"),
          ("Magna AI 合資前期成本", "否", "金額未揭露", "季度營業利益說明")):
    r_ = rrd[0]; rrd[0] += 1
    for c, v in enumerate(t, 1): RD.cell(r_, c, v)
rd("sp3", '""')
rd("h_load", '"承重最大的 Assumed 與可證偽指標"', formula=True)
rd("l1", '"每 US$1M ARR 的 AI 任務數（tpm）：公司未揭露；可證偽：Trend 披露雲端／AI token 成本占營收比或指引再下修。"')
rd("l2", '"Vision One 範圍、模組、單價成長拆解（s_b26／m_b26／p_b26）：公司未揭露；可證偽：公司揭露客戶數、每客戶 ARR、方案數。"')
rd("l3", '"單價侵蝕（aier_A／B／C）：公司只說競爭與免費授權；可證偽：Vision One ARR 增速與每客戶 ARR 的背離。"')
rd("l4", '"相對倍數 rel_mult（cm_adj）：兩讀數分歧大；主腿為 P/S 口徑；可證偽：本公司現行 P/S 相對值（Val 頁 rel）持續高於或低於 rel_mult；Sens 有 ps_only、psg_only、low、high 四案。"')
rd("sp4", '""')
rd("h_na", '"不適用說明"', formula=True)
rd("na1", '"不適用：Meta 的算力、Tokenomics 自用 AI 讀法 2、Wattonomics 對照：Trend Micro 不自建 AI 算力，AI 成本以向模型商購買的 token 計。"')
rd("na2", '"不適用：三層協調（Coord 頁、對帳中旗標、兩讀數）：非 Siliconomics 覆蓋。"')
rd("na3", '"股數：RSU 與選擇權庫藏股法：合併股票薪酬與 RSU 未揭露（未找到），以流通股數為準；ASR、可轉債：無。淨債務：只計有利息負債與融資租賃（資料未見，基準 0），營業租賃不計。"')
rd("na4", '"本版未做先驗。"')
rd("nf", '"未找到：FY2021–2022 分部營收與單季分部（Q2–Q4）；受保護單位數、模組數、單價；AI 營收、訂單與付費客戶；漲價日期與幅度；股票薪酬；有利息負債與租賃負債；EPS 共識；第二個 beta 來源；2026Q3 資產負債表；EDINET 法定申報原文；同業 NTM 倍數與本公司 5 年歷史相對倍數；私募談判 2025-02 後進展。"')
rd("sp5", '""')
rd("h_ws", '"工作表：README、In、Hist、PL、CF、Val、Recon、RevDCF、Sens、Sources"', formula=True)
RD.column_dimensions["A"].width = 160; RD.column_dimensions["B"].width = 16; RD.column_dimensions["C"].width = 50; RD.column_dimensions["D"].width = 40
# ================= Sources =================
SO["A1"] = "Sources — 來源與版本"
SO["A2"] = "Tokenomics 快照"; SO["B2"] = f"{TKV}（{TK['source']['file']}；commit {TK['source']['commit'][:8]}；擷取 {TK['source']['extractedAt']}；names：data/tk_names.txt；重現檢查 import_tokenomics.py --check）"
SO["A3"] = "事實總帳"; SO["B3"] = "data/facts_trendmicro.csv／json（177 列）；研究包 data/research_trendmicro.md"
SO["A4"] = "公開資料局限"; SO["B4"] = "法定申報（EDINET 有價證券報告書）原文未直接讀取；財報科目取自公司事業報告 PDF（Annual Report FY2023–2025）與資料商轉載，資料商只作第二來源。"
dn("tk_version_note", "Sources!$B$2")
# ================= 輸出 =================
for ws in wb.worksheets:
    ws.sheet_view.zoomScale = 90
wb.save(A.out)
rm = {"R": R, "IX": IX, "VR": VR, "RVR": RVR, "RCR": RCR, "RDR": RDR, "GATE": GATE, "CK1": CK1, "CK2": CK2, "INR": INR, "SCN_H": SCN_H, "HD": HD,
      "SE": {"BASE": SE_BASE, "ST": SE_ST, "BID": SE_BID, "HDR": SE_HDR, "END": SE_END, "T0": SE_T0, "T1": SE_T1, "T2": SE_T2, "K0": SE_K0, "KR": SE_KR, "FPR": SE_FPR},
      "SE_ROWS": SE_ROWS, "CASES": [(k, sd, st) for k, sd, st, _ in cases], "RISK_HDR": RISK_HDR, "HR": HR, "BUILD_ID": BUILD_ID, "FPR": FPR, "IN_END": IN_END, "SCN_SEC": SCN_SEC}
json.dump(rm, open(os.environ.get("ROWMAP", os.path.join(HERE, "data", "rowmap.json")), "w", encoding="utf-8"), ensure_ascii=False)
print("saved", A.out, BUILD_ID)
