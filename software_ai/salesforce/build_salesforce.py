#!/usr/bin/env python3
"""Salesforce 個股模型 v0.2.3（原型 S：按席位計價；命題＋營收分部 → 損益表 → FCFF DCF＋Comparable）。
用法：python3 build_salesforce.py <raw.xlsx> [--sens data/sens_v01.json] [--rowmap rowmap.json]
所有計算為活公式；輸入在 In 頁（藍字）。LibreOffice 重算後才有快取值。
"""
import datetime, hashlib, json, sys
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.workbook.defined_name import DefinedName

OUT = sys.argv[1]
SENS = sys.argv[sys.argv.index("--sens") + 1] if "--sens" in sys.argv else None
KREQ = sys.argv[sys.argv.index("--kreq") + 1] if "--kreq" in sys.argv else None
ROWMAP = sys.argv[sys.argv.index("--rowmap") + 1] if "--rowmap" in sys.argv else None
VER, VDATE = "0.2.3", "20261117"
PP = json.load(open("/home/claude/company-models/software_ai/data/relmult_hist/out/peer_ps_20261113.json", encoding="utf-8"))["companies"]["salesforce"]["groups"]["all"]
PPD = json.load(open("/home/claude/company-models/software_ai/data/relmult_hist/out/peer_ps_20261113.json", encoding="utf-8"))["as_of"]
RM = json.load(open("/home/claude/company-models/software_ai/data/relmult_hist/out/salesforce.json", encoding="utf-8"))["groups"]["all"]
TK = json.load(open("/home/claude/company-models/software_ai/salesforce/data/tokenomics_snapshot_v22.0.json", encoding="utf-8"))
TKI = TK["items"]
BLUE = Font(color="0000FF"); BOLD = Font(bold=True); YEL = PatternFill("solid", fgColor="FFF2CC"); HDR = PatternFill("solid", fgColor="DDEBF7")
wb = openpyxl.Workbook()
SH = {n: wb.active if n == "README" else wb.create_sheet(n) for n in ["README", "In", "Q", "Seg", "PL", "DCF", "Val", "RevDCF", "Recon", "Sens", "Sources"]}
SH["README"].title = "README"

# ---------------------------------------------------------------- In 頁
D0 = "Verified"; IP = "Interested-party"; AN = "Analogy"; AS = "Assumed"; DE = "Derived"
Q2PR = "Q2 FY27 財報新聞稿（8-K 2026-08-26 附件 99.1）https://s205.q4cdn.com/626266368/files/doc_financials/2027/q2/CRM-Q2-FY27-Earnings-Press-Release.pdf"
INP = []  # (key,label,unit,lo,base,hi,tag,source,incentive)


def add(key, label, unit, lo, base, hi, tag, src, inc=""):
    INP.append((key, label, unit, lo, base, hi, tag, src, inc))


add("px_date", "股價基準日（收盤）", "日期", None, datetime.date(2026, 11, 13), None, D0, "StockAnalysis CRM statistics 2026-11-13 收盤（第二來源）；research_salesforce.md")
add("px", "現價（2026-11-13 收盤）", "US$/股", None, 217.84, None, D0, "StockAnalysis CRM statistics（第二來源）")
add("sh_out", "流通股數（10-Q 封面）", "M股", 800, 819.0, 845, D0, "Q1 FY27 10-Q 封面：約 819M 股（2026-05-21；SEC 0001108524-26-000127，經 boardroomalpha 鏡像讀取）；Q2 FY27 10-Q 封面股數本次未取得（未找到）；StockAnalysis 2026-11-15 顯示 823.0M 僅作第二來源對照")
add("rsu_vest", "限制性股票：預計歸屬（扣預期沒收；Q2 FY27 流通 39M）", "M股", 33, 33, 39, D0, "Q2 FY27 10-Q 股權報酬附註（獨立核對者引述：未歸屬 39M、預計歸屬 33M；本代理未直接讀到 Q2 原檔，Q1 10-Q：流通 40M、預計歸屬 34M）；高端＝全額計入")
add("opt_n", "選擇權流通數", "M股", None, 7, None, D0, "Q2 FY27 10-Q（獨立核對者引述 7M；Q1 10-Q 為 7M，加權履約價 203.99）")
add("opt_k", "選擇權加權平均履約價", "US$/股", None, 205.71, None, D0, "Q2 FY27 10-Q（獨立核對者引述）；庫藏股法：淨股數＝數量×max(0,1−履約價÷現價)")
add("asr_sh", "ASR 最終交割再交付股數（負值＝再減少；封面股數未含）", "M股", -40, 0, 0, AS, "$25B ASR 預計 2026-10 結算、至少回購 14% 流通股、迄今均價約 176 美元（財務長法說）；Q1 已交付約 103M 股（Guggenheim 轉述）；最終再交付股數未揭露，低端 -40M 為假設")
add("cash", "現金及約當現金（2026-07-31）", "US$M", None, 8310, None, D0, Q2PR + " 資產負債表")
add("msec", "有價證券（2026-07-31）", "US$M", None, 3093, None, D0, Q2PR + " 資產負債表")
add("strat", "策略投資（2026-07-31）", "US$M", None, 11324, None, D0, Q2PR + " 資產負債表")
add("strat_hc", "策略投資計入淨現金的比例", "比例", 0.5, 1.0, 1.0, AS, "預設全額計入（帳面）；低端打五折反映流動性與評價風險；判斷題")
add("debt", "債務（2026-07-31，非流動；流動為 0）", "US$M", None, 39288, None, D0, Q2PR + " 資產負債表")
add("lease", "融資租賃負債（營業租賃不計）", "US$M", None, 633, None, D0, "Q2 FY27 10-Q 租賃附註（獨立核對者引述 633；Q1 10-Q 為 664）；營業租賃 2,455（Q2 10-Q，含流動 531）不計入淨債務（口徑）；本代理未直接讀到融資租賃餘額")
add("dps", "每股股利（年）", "US$/股", None, 1.77, None, DE, "Q2 FY27 股利支出 364M × 4 ÷ 稀釋股數 821.9M（淨利 3,526 ÷ 稀釋 EPS 4.29）")
add("rf", "無風險利率（10 年美債）", "比例", 0.045, 0.051, 0.055, D0, "美國 10 年期公債殖利率 5.10%（FRED DGS10，2026-11-12，現價日前 1 個交易日；本代理已讀 FRED 原檔，2026-11-13 為 5.15%）；工作單 G5 第 2 條統一取值")
add("rf_date", "無風險利率資料日期", "日期", None, datetime.date(2026, 11, 12), None, D0, "FRED DGS10（https://fred.stlouisfed.org/series/DGS10）")
add("data_date", "資料截止日", "日期", None, datetime.date(2026, 11, 16), None, D0, "本研究包擷取日")
add("build_date", "成品日", "日期", None, datetime.date(2026, 11, 17), None, D0, "本版建置日（2026-11-17 工作單 G2 修正）")
add("crpo_g", "Q3 FY27 cRPO 年增指引（固定匯率；不含 Contentful 與 Fin）", "比例", None, 0.14, None, IP, "Q2 FY27 新聞稿 Q3 指引", "公司指引")
add("beta_raw", "Beta（5 年，單一來源）", "—", 1.0, 1.23, 1.4, D0, "StockAnalysis CRM statistics 2026-11-15；規格要兩來源取 Blume 中位數，本次只取得一個（未找到第二來源）")
add("erp", "股權風險溢酬", "比例", 0.04, 0.05, 0.06, AS, "DCF 標準規格預設 5%（4–6%）")
add("kd_sp", "稅前負債信用利差（相對 rf）", "比例", 0.0, 0.012, 0.02, AS, "未找到公司債有效利率；對照：獨立核對者稱有效利率約 4.8%（低於 rf＋利差＝6.2%，Recon 對照）；負債權重約 19%，影響小")
add("g_tv", "永續成長率", "比例", 0.02, 0.03, 0.04, AS, "DCF 標準規格預設 3%（2–4%）")
add("ronic_sp", "RONIC 減 WACC 的利差", "比例", -0.02, 0.0, 0.02, AS, "DCF 標準規格預設 RONIC＝WACC（新投資價值中性）")
add("tax", "稅率（用於 NOPAT）", "比例", 0.18, 0.21, 0.25, IP, "公司 FY2027 non-GAAP 稅率 20.5%（Q2 FY27 新聞稿）；GAAP 有效稅率 FY2026＝2,063÷9,520＝21.7%（Derived）", "公司自訂 non-GAAP 稅率")
add("w_dcf", "DCF 權重（其餘為 Comparable）", "比例", 0.3, 0.5, 0.7, AS, "原則第 4 節預設 50／50")
add("cm_adj", "相對倍數 rel_mult（本公司 ÷ 同業中位數；乘在同業中位 P／S 上）", "倍", RM["sens"]["low"], RM["geo_base"], RM["sens"]["high"], DE, "讀 software_ai/data/relmult_hist/out/salesforce.json（2026-11-17）：geo_base＝" + RM["base_source"] + "；PS 讀數 %.3f、PSG 讀數 %.3f；低／高取 sens.low／sens.high；窗口 %s 至 %s（%d 季）；Comparable 主腿為 P／S，與 PS／PSG 相對值同口徑（G6）" % (RM["ps_reading"], RM["psg_reading"], "2021-09-30", "2026-06-30", RM["valid_points"]["window_quarters"]))
PSS = RM["ps_series"][-1]
add("ps_peer", "同業中位 P／S（現價日；Comparable 主腿）", "倍", None, PP["peer_median_ps"], None, DE, "software_ai/data/relmult_hist/out/peer_ps_20261113.json（StockAnalysis；市值÷最新 TTM 營收；價格日目標 %s；同業 %d 家中位數；逾期 %d 家：%s）" % (PPD, PP["n_peers"], PP["n_stale"], "、".join(PP["stale"])))
add("ps_peer_date", "同業 P／S 價格日（目標日；個別同業逾期見說明）", "日期", None, datetime.date.fromisoformat(PPD), None, DE, "peer_ps_20261113.json as_of；逾期同業 %s（照原則第 15 條標示）" % "、".join(PP["stale"]))
add("ps_peer_db", "同業中位 P／S：資料庫最近一季中位（對照；Sens 用，即 v0.2.1 的值）", "倍", None, PP["db_last_quarter"]["median"], None, DE, "peer_ps_20261113.json db_last_quarter（季末 %s，%d 家）；相對現價日逾期" % (PP["db_last_quarter"]["q"], PP["db_last_quarter"]["n_peers"]))
add("rel_ps", "相對倍數：僅 PS 讀數（ps_only；Sens 用）", "倍", None, RM["sens"]["ps_only"], None, DE, "relmult_hist JSON sens.ps_only")
add("rel_psg", "相對倍數：僅 PSG 讀數（psg_only；Sens 用）", "倍", None, RM["sens"]["psg_only"], None, DE, "relmult_hist JSON sens.psg_only")
add("rel_div", "相對倍數：PS 與 PSG 兩讀數分歧度（divergence）", "比例", None, RM["divergence"], None, DE, "relmult_hist JSON divergence；divergence_gt_30pct＝" + str(RM["divergence_gt_30pct"]))
add("rel_gr", "相對倍數：成長比（本公司 TTM 營收年增 ÷ 同業中位；TTM，NTM 未取得）", "比例", None, RM["growth_ratio"]["value"], None, DE, "relmult_hist JSON growth_ratio（本公司 %.1f%%、同業中位 %.1f%%，%s 季末）" % (RM["growth_ratio"]["company_g"]*100, RM["growth_ratio"]["peer_median_g"]*100, RM["growth_ratio"]["quarter"]))
# 營收基期（Verified：公司揭露）
add("apps_h1", "Agentforce Apps 營收：FY2027 上半年", "US$M", None, 14102, None, D0, Q2PR + " Revenue by Service Offering 六個月欄")
add("apps_q2", "Agentforce Apps 營收：Q2 FY27", "US$M", None, 7193, None, D0, Q2PR)
add("data_h1", "Data 360, Headless Platform, and Other 營收：FY2027 上半年", "US$M", None, 7299, None, D0, Q2PR)
add("data_q2", "Data 360, Headless Platform, and Other 營收：Q2 FY27", "US$M", None, 3618, None, D0, Q2PR)
add("hedge27", "Hedging gains（上半年；下半年假設 0）", "US$M", None, 12, None, D0, Q2PR)
add("ps_h1", "專業服務與其他營收：FY2027 上半年（540＋525）", "US$M", None, 1065, None, DE, "Q1 FY27 新聞稿 540 ＋ Q2 FY27 新聞稿 525，相加")
add("ps_q2", "專業服務與其他營收：Q2 FY27", "US$M", None, 525, None, D0, Q2PR)
add("cogs_h1", "營業成本：FY2027 上半年", "US$M", None, 5219, None, D0, Q2PR + " 六個月欄")
add("rev_h1", "合併營收：FY2027 上半年", "US$M", None, 22478, None, D0, Q2PR + " 六個月欄")
add("ebit_h1", "營業利益：FY2027 上半年（2,347＋2,331）", "US$M", None, 4678, None, DE, "Q1 FY27 新聞稿 2,347 ＋ Q2 FY27 新聞稿 2,331，相加")
# 下半年與 FY2027 假設
add("h_apps", "Agentforce Apps 下半年平均季度相對 Q2 的增幅", "比例", 0.0, 0.035, 0.06, AS, "Q1→Q2 季增 4.1%（6,910→7,193，Derived）；下半年兩季平均高於 Q2 約 3.5% 為假設")
add("h_data", "Data 360 等 下半年平均季度相對 Q2 的增幅", "比例", 0.0, 0.03, 0.06, AS, "Q1→Q2 季減 1.8%（3,683→3,618，Derived）；下半年假設平均高於 Q2 3%，含 Informatica 與資料需求；判斷題")
add("acq27", "Contentful 與 Fin 併入 FY2027 的營收（歸 Data 360 等）", "US$M", 0, 200, 200, IP, "財務長法說：指引含約 200M 來自預期交割（Q2 FY27 逐字稿）；歸類到 Data 360 等 為假設", "公司指引，交割未完成")
add("ps_h2g", "專業服務下半年平均季度相對 Q2 的增幅", "比例", -0.05, -0.02, 0.0, AS, "Q2 年減 3.8%（Derived）；假設延續微幅下滑")
# 命題模組（原型 S）
add("agent27", "Agentforce 代理營收：FY2027（歸 Agentforce Apps）", "US$M", 1200, 1500, 1800, IP, "Futurum Q2 FY27 摘要：on track for roughly $1.5B revenue this year；Q2 Agentforce ARR 逾 1.5B（公司自揭，且自 Q2 起納入 Slackbot 與 Headless 360）；歸入 Apps 為假設", "公司自揭 KPI；ARR 口徑 Q2 擴大")
add("p0", "每任務實現單價（FY2027）", "US$/任務", 0.05, 0.10, 0.15, AS, "list price 為每動作 $0.10（20 Flex Credits；2025-05-15 新聞稿）；實際混合 Flex Credits、每席位無限量與 AELA，實現單價未揭露")
add("gs", "席位毛成長（客戶人員成長與新客，未扣替代）", "比例/年", 0.01, 0.03, 0.05, AS, "公司未揭露席位數；Q2 FY27 法說稱銷售、服務、Slack 席位年增（口述）；數值為假設")
add("pu", "每席位淨價格／版本組合提升", "比例/年", 0.01, 0.03, 0.05, AS, "2025-08-01 list price 平均 +6%（官方）；高階版溢價 60–80% 但採用約 5%（口述）；折為每年 3% 為假設")
add("gt_dec", "任務量成長逐年衰減係數", "倍", 0.5, 0.75, 0.9, AS, "無證據；判斷題")
add("tok_task", "每個代理任務（AWU）的總 token 數", "token/任務", 7000, 7526, 8333, DE, "公司累計 token ÷ 累計 AWU：Q1 FY27 28.6 兆 ÷ 3.8B＝7,526；Q4 FY26 約 20 兆 ÷ 2.4B＝8,333（Q4 FY26／Q1 FY27 新聞稿與 Futurum 轉述；累計值對累計值，AWU 2.4→3.8→7.0B 與 Q2 單季 3.2B 一致）。公司自揭分子分母，另標 Interested-party 性質")
add("ptok", "向模型商購買的每百萬 token 價格（IF_PriceRef_Sol 層級；類比）", "US$/M token", 0.0804, 1.568, 8.0356, AN, "Tokenomics v22.0 IF_PriceRef_Sol 1.568（OpenAI 有效價、參考請求混合）；低端 Luna 0.0804（模型組合路由到便宜模型）；高端 Astra 8.0356（上界：只動成本一邊，不含售價同步）。Salesforce 實際向誰以何價購買未揭露，故標類比")
add("px_dec", "每百萬 token 市場價逐年變動（FY2028 起）", "比例/年", -0.25, 0.0, 0.0, AS, "價格路徑待 AI供應鏈 登錄；基準 0（不降價，保守）；低端每年降 25%")
add("gd0", "Data 360 等 FY2028 年增", "比例", 0.08, 0.12, 0.16, AS, "Q2 FY27 年增 +20.3%（含 Informatica 併入，Derived）；有機成長低於此；假設 FY2028 為 12%")
add("gd_term", "Data 360 等 FY2032 年增", "比例", 0.04, 0.06, 0.08, AS, "線性收斂；判斷題")
add("ps_g", "專業服務與其他逐年成長（FY2028 起）", "比例/年", -0.06, -0.03, 0.0, AS, "FY2023–FY2026 專業服務由 2,331 降到 2,137（Derived）")
# 成本與現金流
add("cogs_dr", "營業成本率（不含代理推論）逐年下降（基準持平）", "百分點/年", -0.005, 0.0, 0.005, AS, "基準＝上半年營業成本率扣 FY2027 推論成本後全期持平（原則第 35 條）；FY2024–FY2026 毛利率 75.5%→77.7%（Derived）")
add("opx_dr", "營業費用率逐年下降（基準持平於最近實績，改善只放敏感度）", "百分點/年", -0.005, 0.0, 0.013, AS, "基準＝上半年營業費用率 56.0% 全期持平（原則第 35 條）；FY2026 57.6%、上半年 56.0%（Derived）；高端 1.3 點/年為假設收斂，低端 −0.5 點為惡化")
add("da_pct", "折舊與攤銷占營收", "比例", 0.075, 0.085, 0.095, AS, "FY2026 3,631÷41,525＝8.7%、FY2025 9.2%（Derived）；Informatica 無形資產攤銷使其上升或持平")
add("capex_pct", "資本支出占營收", "比例", 0.012, 0.015, 0.02, IP, "公司 FY2027 指引約 1.5%（Q2 FY27 新聞稿）；FY2026 594÷41,525＝1.4%", "公司指引")
add("wc_k", "營運資金流入（負營運資金）占營收增量比例", "比例", 0.05, 0.10, 0.15, AS, "獨立核對者由 10-K 現金流量表推得歷史比率：FY2025 約 8%、FY2026 約 11%（本代理未重算）；取約 10%")
add("sbc_pct", "股票薪酬占營收（只用於 P／FCF 對照列與落差分解、DCF 加回對照；不進目標價）", "比例", 0.07, 0.08, 0.09, DE, "Q2 FY27 904÷11,345＝8.0%、FY2026 8.4%（Derived）")
add("oth27", "FY2027 業外淨額（含 Q2 策略投資利得 2,613 減利息）", "US$M", 0, 700, 2000, AS, "Q2 利得 2,613（Verified）、利息費用 Q2 473（Verified）；全年合計為假設")
add("oth_run", "FY2028 起每年業外淨額（利息費用淨額）", "US$M", -1800, -1200, -600, AS, "Q2 FY27 利息費用 473／季（Verified）；利息收入與策略投資損益未預測；假設淨約 −1,200")
add("sh27", "FY2027 加權稀釋股數（僅用於 EPS 顯示）", "M股", 850, 880, 900, AS, "Q2 FY27 稀釋股數約 822（Derived）；上半年較高；僅影響 EPS 顯示")
add("bb", "FY2028 起每年股數淨減少（僅 EPS 顯示）", "比例/年", 0.0, 0.02, 0.04, AS, "$50B 授權（Q4 FY26 新聞稿）；僅影響 EPS 顯示")
SC = {  # 情境專屬列（A 樂觀／B 基準／C 保守）：key 前綴 → (標籤, 單位, A, B, C, 標記, 來源, 誘因)
 "rho": ("客戶人力替代造成的席位年縮減率（FY2028 起）", "比例/年", 0.01, 0.025, 0.05, AS, "無揭露；Q2 FY27 法說稱席位仍年增（口述，Interested-party）；三情境以替代率分化", ""),
 "gt0": ("每任務動作量 FY2028 年增", "比例", 1.0, 0.7, 0.4, AS, "公司稱 Agentforce ARR 年增逾 240%（Q2 FY27，Interested-party）；FY2027 為推得值；逐年以 gt_dec 衰減；三情境分化為假設", ""),
 "dp": ("每任務實現單價逐年變動（FY2028 起）", "比例/年", -0.05, -0.10, -0.18, AS, "list price $0.10 對照前代每對話 $2（2024-10；2025-05-15 新聞稿）；單價下降速度為假設", ""),
 "tp_w": ("情境權重", "比例", 0.25, 0.5, 0.25, AS, "原則第 4 節預設 25／50／25", ""),
}

ws = SH["In"]
ws["A1"] = "Salesforce 個股模型 v0.1 輸入頁（藍字＝輸入；黑字＝公式）。金額 US$M，會計年度 2 月制（FY2027＝2026-02-01～2027-01-31）。"
ws["A2"] = "基準欄 E 進主模型；D／F（敏感度低／高）只用於 Sens 頁一次動一個參數，不進目標價。情境專屬列（鍵尾 A／B／C）由「情境驅動對照」節並列。本版未做事前驗證。"
for j, h in enumerate(["鍵", "項目", "單位", "敏感度低", "基準", "敏感度高", "標記", "來源", "來源誘因"], 1):
    c = ws.cell(3, j, h); c.font = BOLD; c.fill = HDR
IR = {}  # key -> row
r = 4
for key, label, unit, lo, base, hi, tag, src, inc in INP:
    IR[key] = r
    ws.cell(r, 1, key); ws.cell(r, 2, label); ws.cell(r, 3, unit)
    for col, v in ((4, lo), (5, base), (6, hi)):
        if v is not None:
            c = ws.cell(r, col, v); c.font = BLUE
            if isinstance(v, datetime.date): c.number_format = "yyyy-mm-dd"
    ws.cell(r, 7, tag); ws.cell(r, 8, src); ws.cell(r, 9, inc)
    r += 1
for pre, (label, unit, a, b, c_, tag, src, inc) in SC.items():
    for s, v in zip("ABC", (a, b, c_)):
        k = f"{pre}{s}" if pre == "tp_w" else f"{pre}_{s}"
        IR[k] = r
        ws.cell(r, 1, k); ws.cell(r, 2, f"情境 {s}：{label}"); ws.cell(r, 3, unit)
        ws.cell(r, 5, v).font = BLUE
        ws.cell(r, 7, tag); ws.cell(r, 8, src); ws.cell(r, 9, inc)
        r += 1
r += 1
ws.cell(r, 1, "J. 情境驅動對照（純展示；A 樂觀／B 基準／C 保守）").font = BOLD
r += 1
for j, h in enumerate(["鍵", "項目", "單位", "情境 A", "情境 B", "情境 C", "標記", "說明"], 1):
    ws.cell(r, j, h).font = BOLD
SCN_H = r
r += 1
SCN_ROWS = {}
for pre, key in (("rho", "scn_rho"), ("gt0", "scn_gt0"), ("dp", "scn_dp"), ("tp_w", "scn_w")):
    ks = [f"{pre}{s}" if pre == "tp_w" else f"{pre}_{s}" for s in "ABC"]
    ws.cell(r, 1, key); ws.cell(r, 2, SC[pre][0]); ws.cell(r, 3, SC[pre][1])
    for j, k in zip((4, 5, 6), ks):
        ws.cell(r, j, f"=E{IR[k]}")
    ws.cell(r, 7, "Derived"); ws.cell(r, 8, "引用 In 頁 " + "／".join(ks) + "；三情境並排，只引用上方輸入列")
    SCN_ROWS[key] = r
    r += 1
IN_LAST = r
for k in ("px_date", "px"):
    pass


def I(key):
    return f"In!$E${IR[key]}"


def _fp():
    return hashlib.sha256(json.dumps([(k, (ws.cell(IR[k], 5).value if not isinstance(ws.cell(IR[k], 5).value, datetime.date) else str(ws.cell(IR[k], 5).value))) for k in sorted(IR)], ensure_ascii=False).encode()).hexdigest()[:12]
fp0 = _fp()

# ---------------------------------------------------------------- Q 頁（法定申報歷史；每格一列出處）
YEARS = ["FY2024", "FY2025", "FY2026", "FY2027", "FY2028", "FY2029", "FY2030", "FY2031", "FY2032"]
YC = {y: "CDEFGHIJK"[i] for i, y in enumerate(YEARS)}
HIST = ["FY2024", "FY2025", "FY2026"]
K26 = "10-K FY2026（SEC 0001108524-26-000060；經第三方鏡像讀取，待回 SEC 原檔核對頁碼）MD&A"
K25 = "10-K FY2025 Item 7；https://www.sec.gov/Archives/edgar/data/1108524/000110852425000006/crm-20250131.htm"
PR4 = "Q4 FY26 新聞稿 https://s205.q4cdn.com/626266368/files/doc_financials/2026/q4/CRM-Q4-FY26-Earnings-Press-Release.pdf"
Q = SH["Q"]
Q["A1"] = "Q 頁：公司法定申報與新聞稿的歷史數字（US$M）；FY2024–FY2026 為會計年度（2 月制）。每列單一標記；Interested-party 列有來源誘因。"
Q["A2"] = "FY2024 稅前利益以下科目本次未讀到，標「未找到」。"
for j, h in enumerate(["指標", "FY2024", "FY2025", "FY2026", "單位", "來源", "標記", "來源誘因"], 1):
    Q.cell(3, j, h).font = BOLD
QR = {}
qdata = [
 ("sub", "訂閱與支援營收", 32537, 35679, 39388, K26, D0),
 ("ps", "專業服務與其他營收", 2320, 2216, 2137, K26, D0),
 ("rev", "合併營收（公司揭露）", 34857, 37895, 41525, K26, D0),
 ("cogs", "營業成本", 8541, 8643, 9270, K26 + "；FY2026 另見 " + PR4, D0),
 ("gp", "毛利", 26316, 29252, 32255, K26, D0),
 ("ebit", "營業利益", 5011, 7205, 8331, K26, D0),
 ("ebt", "稅前利益", "未找到", 7438, 9520, PR4, D0),
 ("tax", "所得稅", "未找到", 1241, 2063, PR4, D0),
 ("ni", "淨利", "未找到", 6197, 7457, PR4, D0),
 ("eps", "稀釋 EPS（US$）", "未找到", 6.36, 7.80, PR4, D0),
 ("sh", "加權平均稀釋股數（M）", "未找到", 974, 956, PR4, D0),
 ("da", "折舊與攤銷", "未找到", 3477, 3631, PR4, D0),
 ("capex", "資本支出", "未找到", 658, 594, PR4, D0),
 ("sbc", "股票薪酬（含重組相關）", 2787, 3183, 3509, "10-K FY2025 Item 7（FY2024、FY2025）；FY2026＝Q4 FY26 新聞稿 3,480＋重組相關 29（Derived）", DE),
]
rr = 4
for key, lab, a, b, c_, src, tag in qdata:
    QR[key] = rr
    Q.cell(rr, 1, lab)
    for j, v in zip((2, 3, 4), (a, b, c_)):
        Q.cell(rr, j, v).font = BLUE
    Q.cell(rr, 5, "US$M" if key not in ("eps", "sh") else ("US$" if key == "eps" else "M股"))
    Q.cell(rr, 6, src); Q.cell(rr, 7, tag); Q.cell(rr, 8, "")
    rr += 1
Q.cell(rr, 1, "核對：訂閱與支援＋專業服務－合併營收（應為 0）"); QR["chk"] = rr
for j, col in zip((2, 3, 4), "BCD"):
    Q.cell(rr, j, f"={col}{QR['sub']}+{col}{QR['ps']}-{col}{QR['rev']}")
rr += 1
Q.cell(rr, 1, "核對：營收－營業成本－毛利（應為 0）"); QR["chk2"] = rr
for j, col in zip((2, 3, 4), "BCD"):
    Q.cell(rr, j, f"={col}{QR['rev']}-{col}{QR['cogs']}-{col}{QR['gp']}")


def QH(key, y):  # 歷史值引用
    return f"Q!${'BCD'[HIST.index(y)]}${QR[key]}"


# ---------------------------------------------------------------- Seg 頁（命題模組＋營收分部）
S = SH["Seg"]
S["A1"] = "命題模組（原型 S：按席位計價）：Agentforce Apps ＝ 席位型營收（席位 × 每席位年費，扣客戶人力替代）＋ 代理營收（任務數 × 每任務單價）；推論成本＝任務數 × 每任務 token × 向模型商購買的市場價（Tokenomics v22.0 快照，不自行重估）。AI 淨效果＝代理營收增量 − 席位侵蝕 − 推論成本。"
S["A2"] = "FY2027 為上半年實績加下半年假設，三情境相同；情境分化自 FY2028 起（席位替代率、任務量成長、每任務單價）。變數名稱與 software_ai/unit_economics 一致；該模型 v0.1 完成前在本頁自建（v0.2 改讀 unit_economics）。"
SEGLAB = ["tasks", "price", "agent", "seat", "seat_cf", "apps", "data", "hedge", "sub", "ps", "rev", "eros", "incr", "inf", "net", "inf_full", "chk1", "chk2"]
LABS = {"tasks": "代理任務數（十億次）", "price": "每任務實現單價（US$）", "agent": "Agentforce 代理營收（命題子列；歸 Agentforce Apps）", "seat": "席位型營收（命題子列；歸 Agentforce Apps）",
        "seat_cf": "席位型營收：無 AI 替代的對照路徑（rho＝0）", "apps": "Agentforce Apps", "data": "Data 360, Headless Platform, and Other", "hedge": "Hedging gains",
        "sub": "訂閱與支援", "ps": "專業服務與其他", "rev": "合併營收", "eros": "AI 淨效果之一：席位侵蝕（對照路徑減實際）", "incr": "AI 淨效果之二：代理營收增量",
        "inf": "AI 淨效果之三：推論成本（Tokenomics 市場價）", "net": "AI 淨效果（增量 − 侵蝕 − 推論成本）", "inf_full": "對照：以模型商全成本計的推論成本（IF_FullCostDefault_Sol，只作成本下限對照，不進模型）",
        "chk1": "檢查：Apps＋Data＋Hedging－訂閱與支援", "chk2": "檢查：訂閱與支援＋專業服務－合併營收"}
SR = {}
for si, s in enumerate("ABC"):
    r0 = 4 + si * (len(SEGLAB) + 4)
    S.cell(r0, 1, f"情境 {s}（{ {'A':'樂觀：席位替代低、任務量與單價佳','B':'基準','C':'保守：席位替代高、任務量與單價差'}[s] }）").font = BOLD
    for y, c in YC.items():
        S[f"{c}{r0+1}"] = y + ("A" if y in HIST else "E"); S[f"{c}{r0+1}"].font = BOLD
    S.cell(r0 + 1, 1, "US$M 除另註");
    rows = {k: r0 + 2 + i for i, k in enumerate(SEGLAB)}
    SR[s] = rows | {"hdr": r0 + 1}
    for k, rw in rows.items():
        S.cell(rw, 1, LABS[k])
    T = "tasks"
    for y, c in YC.items():
        p = chr(ord(c) - 1)
        R = lambda k: f"{c}{rows[k]}"
        P = lambda k: f"{p}{rows[k]}"
        if y in HIST:
            S[R("sub")] = f"={QH('sub', y)}"; S[R("ps")] = f"={QH('ps', y)}"; S[R("rev")] = f"={QH('rev', y)}"
            S[R("chk2")] = f"={R('sub')}+{R('ps')}-{R('rev')}"
            for k in ("apps", "data"): S[R(k)] = "未找到（FY2026 以前未回溯至新二分類）"
            continue
        if y == "FY2027":
            S[R("price")] = f"={I('p0')}"
            S[R("agent")] = f"={I('agent27')}"
            S[R("tasks")] = f"={R('agent')}/({R('price')}*1000)"
            S[R("apps")] = f"={I('apps_h1')}+2*{I('apps_q2')}*(1+{I('h_apps')})"
            S[R("seat")] = f"={R('apps')}-{R('agent')}"
            S[R("seat_cf")] = f"={R('seat')}"
            S[R("data")] = f"={I('data_h1')}+2*{I('data_q2')}*(1+{I('h_data')})+{I('acq27')}"
            S[R("hedge")] = f"={I('hedge27')}"
            S[R("ps")] = f"={I('ps_h1')}+2*{I('ps_q2')}*(1+{I('ps_h2g')})"
        else:
            n = YEARS.index(y) - YEARS.index("FY2028")
            S[R("price")] = f"={P('price')}*(1+{I('dp_'+s)})"
            S[R("tasks")] = f"={P('tasks')}*(1+{I('gt0_'+s)}*{I('gt_dec')}^{n})"
            S[R("agent")] = f"={R('tasks')}*{R('price')}*1000"
            S[R("seat")] = f"={P('seat')}*(1+{I('gs')}-{I('rho_'+s)})*(1+{I('pu')})"
            S[R("seat_cf")] = f"={P('seat_cf')}*(1+{I('gs')})*(1+{I('pu')})"
            S[R("apps")] = f"={R('seat')}+{R('agent')}"
            S[R("data")] = f"={P('data')}*(1+{I('gd0')}+({I('gd_term')}-{I('gd0')})*{n}/4)"
            S[R("hedge")] = 0
            S[R("ps")] = f"={P('ps')}*(1+{I('ps_g')})"
        if y == "FY2027":
            S[R("apps")] = f"={R('seat')}+{R('agent')}"
            S[R("seat")] = f"={I('apps_h1')}+2*{I('apps_q2')}*(1+{I('h_apps')})-{R('agent')}"
        S[R("sub")] = f"={R('apps')}+{R('data')}+{R('hedge')}"
        S[R("rev")] = f"={R('sub')}+{R('ps')}"
        yrs = YEARS.index(y) - YEARS.index("FY2027")
        S[R("inf")] = f"={R('tasks')}*{I('tok_task')}*{I('ptok')}*(1+{I('px_dec')})^{max(yrs-0,0) if yrs>0 else 0}/1000"
        S[R("inf_full")] = f"={R('tasks')}*{I('tok_task')}*{TKI['IF_FullCostDefault_Sol']['values']['VR200 NVL72']['基準']}/1000"
        S[R("incr")] = f"={R('agent')}"
        S[R("eros")] = f"={R('seat_cf')}-{R('seat')}"
        S[R("net")] = f"={R('incr')}-{R('eros')}-{R('inf')}"
        S[R("chk1")] = f"={R('apps')}+{R('data')}+{R('hedge')}-{R('sub')}"
        S[R("chk2")] = f"={R('sub')}+{R('ps')}-{R('rev')}"
SEG_END = 4 + 3 * (len(SEGLAB) + 4)
S.cell(SEG_END, 1, "註：Tokenomics 價格兩欄：(1) 軟體公司向模型商購買的市場價（In 頁 ptok，IF_PriceRef_Sol；低端 Luna、高端 Astra）進模型；(2) 全成本（IF_FullCostDefault_Sol，VR200 基準 0.128 US$/M）只列為「模型商自己做這件事的成本下限」對照。").font = BOLD

# ---------------------------------------------------------------- PL 頁
P = SH["PL"]
P["A1"] = "損益表：單一官方營運分部（10-K：one operating segment），營收列＝公司揭露的營收型態與兩項固定產品線（Agentforce Apps；Data 360, Headless Platform, and Other）；命題子列為席位型與代理營收。GAAP 口徑，US$M。"
P["A2"] = "FY2024A–FY2026A 取自 Q 頁；FY2027E 為上半年實績加下半年假設；FY2028E–FY2032E 為情境預測。Non-GAAP 營業利益率為公司定義，本模型不預測，只在 Recon 對照。"
PLB = ["rev", "cogs_x", "inf", "cogs", "gp", "opex", "ebit", "da", "capex", "nopat", "dwc", "fcff", "sbc", "fcf_u", "oth", "fcf_cf", "fcfl", "ebt", "taxx", "ni", "shs", "eps"]
PLL = {"rev": "營收", "cogs_x": "營業成本（不含代理推論）", "inf": "代理推論成本", "cogs": "營業成本", "gp": "毛利", "opex": "營業費用", "ebit": "營業利益", "da": "折舊與攤銷",
       "capex": "資本支出", "nopat": "NOPAT（營業利益×(1−稅率)；股票薪酬不加回）", "dwc": "營運資金增加（負值＝流入）", "fcff": "無槓桿自由現金流 FCFF", "sbc": "股票薪酬（Comparable 現金流口徑用）",
       "fcf_u": "無槓桿、股票薪酬加回的自由現金流（FCFF＋股票薪酬）", "fcf_cf": "有槓桿公司口徑自由現金流（FCFF＋股票薪酬＋稅後業外淨額；對應同業 OCF 減資本支出；僅供 Comparable）", "fcfl": "收支關卡用自由現金流（FCFF＋稅後業外淨額−股利）", "oth": "其他收入（業外淨額）", "ebt": "稅前利益", "taxx": "所得稅", "ni": "稅後淨利", "shs": "稀釋股數（M）", "eps": "EPS（GAAP 稀釋，美元）"}
PR = {}
BASE0 = 70
for si, s in enumerate("ABC"):
    r0 = BASE0 + si * (len(PLB) + 4)
    P.cell(r0, 1, f"情境 {s} 損益與現金流").font = BOLD
    for y, c in YC.items():
        P[f"{c}{r0+1}"] = y + ("A" if y in HIST else "E"); P[f"{c}{r0+1}"].font = BOLD
    rows = {k: r0 + 2 + i for i, k in enumerate(PLB)}
    PR[s] = rows | {"hdr": r0 + 1}
    for k, rw in rows.items():
        P.cell(rw, 1, f"情境 {s}｜{PLL[k]}")
    sr = SR[s]
    for y, c in YC.items():
        p = chr(ord(c) - 1)
        R = lambda k: f"{c}{rows[k]}"
        Pv = lambda k: f"{p}{rows[k]}"
        if y in HIST:
            P[R("rev")] = f"={QH('rev', y)}"; P[R("cogs")] = f"={QH('cogs', y)}"; P[R("gp")] = f"={QH('gp', y)}"; P[R("ebit")] = f"={QH('ebit', y)}"
            P[R("opex")] = f"={R('gp')}-{R('ebit')}"
            P[R("sbc")] = f"={QH('sbc', y)}"
            if y != "FY2024":
                P[R("da")] = f"={QH('da', y)}"; P[R("capex")] = f"={QH('capex', y)}"
                P[R("ebt")] = f"={QH('ebt', y)}"; P[R("taxx")] = f"={QH('tax', y)}"; P[R("ni")] = f"={QH('ni', y)}"; P[R("eps")] = f"={QH('eps', y)}"; P[R("shs")] = f"={QH('sh', y)}"
                P[R("oth")] = f"={R('ebt')}-{R('ebit')}"
            else:
                for k in ("da", "capex", "ebt", "taxx", "ni", "eps", "shs", "oth"): P[R(k)] = "未找到"
            continue
        yi = YEARS.index(y)
        P[R("rev")] = f"=Seg!{c}{sr['rev']}"
        P[R("inf")] = f"=Seg!{c}{sr['inf']}"
        if y == "FY2027":
            # 上半年營業成本率（含推論）扣掉 FY2027 推論占比＝不含推論成本率
            P[R("cogs_x")] = f"={R('rev')}*({I('cogs_h1')}/{I('rev_h1')})-{R('inf')}"
            P[R("opex")] = f"={R('rev')}*(({I('rev_h1')}-{I('cogs_h1')}-{I('ebit_h1')})/{I('rev_h1')})"
        else:
            P[R("cogs_x")] = f"={R('rev')}*({Pv('cogs_x')}/{Pv('rev')}-{I('cogs_dr')})"
            P[R("opex")] = f"={R('rev')}*({Pv('opex')}/{Pv('rev')}-{I('opx_dr')})"
        P[R("cogs")] = f"={R('cogs_x')}+{R('inf')}"
        P[R("gp")] = f"={R('rev')}-{R('cogs')}"
        P[R("ebit")] = f"={R('gp')}-{R('opex')}"
        P[R("da")] = f"={R('rev')}*{I('da_pct')}"
        P[R("capex")] = f"={R('rev')}*{I('capex_pct')}"
        P[R("nopat")] = f"={R('ebit')}*(1-{I('tax')})"
        P[R("dwc")] = f"=-{I('wc_k')}*({R('rev')}-{Pv('rev')})"
        P[R("fcff")] = f"={R('nopat')}+{R('da')}-{R('capex')}-{R('dwc')}"
        P[R("sbc")] = f"={R('rev')}*{I('sbc_pct')}"
        P[R("fcf_u")] = f"={R('fcff')}+{R('sbc')}"
        P[R("oth")] = f"={I('oth27')}" if y == "FY2027" else f"={I('oth_run')}"
        P[R("fcf_cf")] = f"={R('fcf_u')}+{R('oth')}*(1-{I('tax')})"
        P[R("fcfl")] = f"={R('fcff')}+{R('oth')}*(1-{I('tax')})-{I('dps')}*{I('sh_out')}"
        P[R("ebt")] = f"={R('ebit')}+{R('oth')}"
        P[R("taxx")] = f"={R('ebt')}*{I('tax')}"
        P[R("ni")] = f"={R('ebt')}-{R('taxx')}"
        P[R("shs")] = f"={I('sh27')}" if y == "FY2027" else f"={Pv('shs')}*(1-{I('bb')})"
        P[R("eps")] = f"={R('ni')}/{R('shs')}"
PL_END = BASE0 + 3 * (len(PLB) + 4)
# 標準損益表（情境 B；v1.3 必備列）
b = PR["B"]
STD = ["營收", "營收年增率", "營業成本", "毛利", "毛利率", "營業費用", "營業利益", "營業利益率", "其他收入", "稅前利益", "所得稅", "稅後淨利", "淨利率", "EPS（GAAP 稀釋，美元）", "淨利年增率"]
P.cell(4, 1, "標準損益表（GAAP；US$M，EPS 為美元；情境 B）：FY2024A–FY2026A 實績、FY2027E–FY2032E 預測").font = BOLD
for y, c in YC.items():
    P[f"{c}5"] = y + ("A" if y in HIST else "E"); P[f"{c}5"].font = BOLD
P.cell(5, 1, "項目")
STR = {n: 6 + i for i, n in enumerate(STD)}
for n, rw in STR.items(): P.cell(rw, 1, n)
for y, c in YC.items():
    p = chr(ord(c) - 1)
    ref = lambda k: f"{c}{b[k]}"
    P[f"{c}{STR['營收']}"] = f"={ref('rev')}"
    if y != "FY2024": P[f"{c}{STR['營收年增率']}"] = f"={c}{STR['營收']}/{p}{STR['營收']}-1"
    else: P[f"{c}{STR['營收年增率']}"] = "未找到"
    P[f"{c}{STR['營業成本']}"] = f"={ref('cogs')}"
    P[f"{c}{STR['毛利']}"] = f"={ref('gp')}"
    P[f"{c}{STR['毛利率']}"] = f"={c}{STR['毛利']}/{c}{STR['營收']}"
    P[f"{c}{STR['營業費用']}"] = f"={ref('opex')}"
    P[f"{c}{STR['營業利益']}"] = f"={ref('ebit')}"
    P[f"{c}{STR['營業利益率']}"] = f"={c}{STR['營業利益']}/{c}{STR['營收']}"
    P[f"{c}{STR['其他收入']}"] = f"={ref('oth')}"
    P[f"{c}{STR['稅前利益']}"] = f"={ref('ebt')}"
    P[f"{c}{STR['所得稅']}"] = f"={ref('taxx')}"
    P[f"{c}{STR['稅後淨利']}"] = f"={ref('ni')}"
    P[f"{c}{STR['淨利率']}"] = f'=IF(ISNUMBER({c}{STR["稅後淨利"]}),{c}{STR["稅後淨利"]}/{c}{STR["營收"]},"未找到")'
    P[f"{c}{STR['EPS（GAAP 稀釋，美元）']}"] = f"={ref('eps')}"
    P[f"{c}{STR['淨利年增率']}"] = f'=IF(AND(ISNUMBER({c}{STR["稅後淨利"]}),ISNUMBER({p}{STR["稅後淨利"]})),{c}{STR["稅後淨利"]}/{p}{STR["稅後淨利"]}-1,"未找到")' if y != "FY2024" else "未找到"
P.cell(STR["淨利年增率"] + 1, 1, "GAAP 稅後淨利 FY2027E 含 Q2 策略投資利得；FY2028E 起不預測策略投資損益。Non-GAAP 營業利益率（公司定義）：FY2026 34.1%、FY2027 指引 34.3%，見 Recon。")
# 分部營收段（情境 B；官方營收型態與固定產品線）
P.cell(23, 1, "分部營收（情境 B；單一營運分部下的營收型態與固定產品線）").font = BOLD
SEGP = {"營收合計（單一營運分部）": "rev", "訂閱與支援": "sub", "專業服務與其他": "ps", "Agentforce Apps": "apps", "Data 360, Headless Platform, and Other": "data", "Hedging gains": "hedge",
        "Agentforce 代理營收（命題子列）": "agent", "席位型營收（命題子列）": "seat"}
SPR = {}
for i, (lab, k) in enumerate(SEGP.items()):
    rw = 24 + i; SPR[lab] = rw
    P.cell(rw, 1, lab)
    for y, c in YC.items():
        if y in HIST and k in ("apps", "data", "hedge", "agent", "seat"):
            P[f"{c}{rw}"] = "未找到"
        elif y in HIST and k == "rev": P[f"{c}{rw}"] = f"={QH('rev', y)}"
        else: P[f"{c}{rw}"] = f"=Seg!{c}{SR['B'][k]}"
CK = 24 + len(SEGP) + 1
P.cell(CK - 1, 1, "對帳（應為 0）").font = BOLD
P.cell(CK, 1, "檢查：分部合計 − 合併營收")
P.cell(CK + 1, 1, "檢查：Apps＋Data＋Hedging − 訂閱與支援（三情境）")
P.cell(CK + 2, 1, "檢查：席位型＋代理 − Agentforce Apps（三情境）")
P.cell(CK + 3, 1, "檢查：分部合計 − 公司揭露合併營收（歷史年度；實績期）")
P.cell(CK + 4, 1, "檢查：營收 − 營業成本 − 毛利（FY2024A–FY2026A）")
for y, c in YC.items():
    refs = []
    for s in "ABC":
        refs.append(f"ABS(Seg!{c}{SR[s]['sub']}+Seg!{c}{SR[s]['ps']}-Seg!{c}{SR[s]['rev']})")
    P[f"{c}{CK}"] = "=MAX(" + ",".join(refs) + ")"
    if y in HIST:
        P[f"{c}{CK+1}"] = 0; P[f"{c}{CK+2}"] = 0
        P[f"{c}{CK+3}"] = f"=ABS({c}{SPR['訂閱與支援']}+{c}{SPR['專業服務與其他']}-{c}{SPR['營收合計（單一營運分部）']})"
        P[f"{c}{CK+4}"] = f"=ABS({c}{STR['營收']}-{c}{STR['營業成本']}-{c}{STR['毛利']})"
    else:
        P[f"{c}{CK+1}"] = "=MAX(" + ",".join(f"ABS(Seg!{c}{SR[s]['chk1']})" for s in "ABC") + ")"
        P[f"{c}{CK+2}"] = "=MAX(" + ",".join(f"ABS(Seg!{c}{SR[s]['seat']}+Seg!{c}{SR[s]['agent']}-Seg!{c}{SR[s]['apps']})" for s in "ABC") + ")"
        P[f"{c}{CK+3}"] = 0 if y != "FY2027" else f"=ABS({c}{SPR['訂閱與支援']}+{c}{SPR['專業服務與其他']}-{c}{SPR['營收合計（單一營運分部）']})"
        P[f"{c}{CK+4}"] = 0
CHK_ROWS = (CK, CK + 4)

# ---------------------------------------------------------------- DCF 頁
Dp = SH["DCF"]
Dp["A1"] = "DCF（FCFF 標準規格）：明確期 FY2027E 剩餘期間～FY2032E；永續價值 ＝ NOPAT(FY2033)×(1−g÷RONIC)÷(WACC−g)；期中折現。評價日＝股價基準日。三情境各算一次。"
Dp["A2"] = "WACC 由 Val 頁推出（讀 In 頁基準欄，不讀敏感度欄）。"
DR = {}
for si, s in enumerate("ABC"):
    r0 = 4 + si * 9
    DR[s] = {k: r0 + 1 + i for i, k in enumerate(["fcff", "frac", "t", "df", "pv"])}
    Dp.cell(r0, 1, f"情境 {s}").font = BOLD
    for y in YEARS[3:]:
        c = YC[y]; Dp[f"{c}{r0}"] = y + "E"
    for k, lab in zip(["fcff", "frac", "t", "df", "pv"], ["FCFF", "計入比例（FY2027 只計評價日之後）", "折現時點（年）", "折現因子", "現值"]):
        Dp.cell(DR[s][k], 1, f"情境 {s}｜{lab}")
    for y in YEARS[3:]:
        c = YC[y]; i = YEARS.index(y) - 3
        Dp[f"{c}{DR[s]['fcff']}"] = f"=PL!{c}{PR[s]['fcff']}"
        Dp[f"{c}{DR[s]['frac']}"] = "=Val!$B$%d" % 0  # placeholder，稍後覆寫
DCF_ROWS = DR

# ---------------------------------------------------------------- Val 頁
V = SH["Val"]
VR = {}
vrows = ["intro", "hint", "blank0", "cap", "px", "px_date", "tp_date", "dcf_tp", "comp_tp", "tp", "up", "val_rec", "gate", "gate_b", "rec", "tp_ref", "sp1", "wh", "ke", "kd", "kdat", "wacc", "ronic", "g", "emkt", "dmkt", "ewt",
         "sp2", "ndh", "cashx", "ndebt", "netcash", "sh", "dps", "sp3", "sumh", "pvx", "tvpv", "ev", "eq", "vps0", "dcf12", "tvsh", "sp4", "ch", "rv", "psm", "ps_eq", "psc12", "ntm", "mult", "ev_c", "comp12", "ref_cur", "sp5", "scenh", "wa", "wb", "wc", "chk_w",
         "sp6", "peerh", "pe1", "pe2", "pe3", "pe4", "pe5", "pe6", "pe_med", "pe_n", "pe_own", "pe_ratio", "pe_reval", "ev_med", "ev_own", "ev_ratio", "sp7", "dh", "d_sbc", "d_lev", "d_time", "d_mult", "d_tot", "uh", "u_dcf", "u_comp", "u_adj", "u_tot", "sp8", "sbh", "sb_pvx", "sb_tvpv", "sb_ev", "sb_eq", "sb_vps0", "sb_dcf12", "sb_tp", "sb_up", "sb_tpw"]
for i, k in enumerate(vrows): VR[k] = 1 + i
V["A1"] = None
for k in ("intro", "hint", "cap"): pass
V.cell(VR["intro"], 1, "估值：DCF（FCFF）與 Comparable（P／S 主腿：同業中位 P／S × rel_mult × 目標日 TTM 營收 ÷ 評價股數）加權；P／FCF Comparable 只作對照列，不進加權。目標價為 12 個月後（評價日＋12 個月）。三情境 25／50／25。上檔空間不是年化報酬。")
V.cell(VR["hint"], 1, "口徑一致：PS／PSG 歷史相對值已內含本公司與同業的利潤率差，故主腿用 P／S；若把它套在 P／FCF 等盈餘型倍數會重複計入利潤率差（工作單 G6）。同業 P／S 取 peer_ps_20261113.json（現價日 2026-11-13；WDAY、HUBS 為 2026-10-26，逾期）；P／S 一律股權口徑；P／FCF 對照列給 rel_mult＝1.00 與本公司現行相對值兩讀數。同業家數 6，不做迴歸。")
V.cell(VR["cap"], 1, "結論（情境加權）").font = BOLD
hdrs = {"A": "項目", "B": "數值", "C": "情境 A", "D": "情境 B", "E": "情境 C"}
def vrow(k, label, b=None, c=None, d=None, e=None):
    r_ = VR[k]; V.cell(r_, 1, label)
    for col, v in zip("BCDE", (b, c, d, e)):
        if v is not None: V[f"{col}{r_}"] = v
vrow("px", "現價（US$）", f"={I('px')}")
vrow("px_date", "股價基準日", f"={I('px_date')}"); V[f"B{VR['px_date']}"].number_format = "yyyy-mm-dd"
vrow("tp_date", "目標時點（股價基準日＋12 個月）", f"=EDATE({I('px_date')},12)"); V[f"B{VR['tp_date']}"].number_format = "yyyy-mm-dd"
wr = lambda s: I("tp_w" + s)
vrow("dcf_tp", "DCF 目標價（情境加權；US$）", f"={wr('A')}*C{VR['dcf12']}+{wr('B')}*D{VR['dcf12']}+{wr('C')}*E{VR['dcf12']}")
vrow("comp_tp", "Comparable 目標價（P／S 主腿；情境加權；US$）", f"={wr('A')}*C{VR['psc12']}+{wr('B')}*D{VR['psc12']}+{wr('C')}*E{VR['psc12']}")
vrow("tp", "加權目標價（DCF 與 Comparable 加權；US$）", f"={I('w_dcf')}*B{VR['dcf_tp']}+(1-{I('w_dcf')})*B{VR['comp_tp']}")
vrow("up", "上檔空間（相對現價）", f"=B{VR['tp']}/B{VR['px']}-1")
vrow("val_rec", "估值評等（依上檔空間五級；未套收支關卡）", f'=IF(B{VR["up"]}>0.5,"強力買進",IF(B{VR["up"]}>0.25,"買進",IF(B{VR["up"]}>=-0.25,"中立",IF(B{VR["up"]}>=-0.5,"放空","強力放空"))))')
vrow("gate", "收支關卡（FY2027–FY2031 自由現金流扣股利合計；US$M；情境 A／B／C）", None, *(f"=SUM(PL!F{PR[s_]['fcfl']}:J{PR[s_]['fcfl']})" for s_ in "ABC"))
vrow("gate_b", "入不敷出判定用（情境 B 收支合計；US$M）", f"=D{VR['gate']}")
vrow("rec", "投資建議", f'=IF(AND(B{VR["gate_b"]}<0,OR(B{VR["val_rec"]}="買進",B{VR["val_rec"]}="強力買進")),"收支不足",B{VR["val_rec"]})')
vrow("tp_ref", "只看 Comparable 的目標價（上檔空間見右）", f"=B{VR['comp_tp']}", f"=B{VR['comp_tp']}/B{VR['px']}-1")
V.cell(VR["wh"], 1, "折現率（DCF）").font = BOLD
vrow("ke", "股權成本 ke（rf＋β×ERP；β＝Blume 調整＝0.67×原始＋0.33）", f"={I('rf')}+(0.67*{I('beta_raw')}+0.33)*{I('erp')}")
vrow("kd", "稅前債務成本（rf＋利差）", f"={I('rf')}+{I('kd_sp')}")
vrow("kdat", "稅後債務成本", f"=B{VR['kd']}*(1-{I('tax')})")
vrow("emkt", "股權市值（US$M；現價×評價股數）", f"={I('px')}*B{VR['sh']}")
vrow("dmkt", "債務本金＋融資租賃（US$M）", f"={I('debt')}+{I('lease')}")
vrow("ewt", "股權權重", f"=B{VR['emkt']}/(B{VR['emkt']}+B{VR['dmkt']})")
vrow("wacc", "WACC", f"=B{VR['ewt']}*B{VR['ke']}+(1-B{VR['ewt']})*B{VR['kdat']}")
vrow("ronic", "RONIC（＝WACC＋利差）", f"=B{VR['wacc']}+{I('ronic_sp')}")
vrow("g", "永續成長率", f"={I('g_tv')}")
V.cell(VR["ndh"], 1, "淨現金（2026-07-31 資產負債表；US$M）").font = BOLD
vrow("cashx", "現金＋有價證券＋策略投資×計入比例", f"={I('cash')}+{I('msec')}+{I('strat')}*{I('strat_hc')}")
vrow("ndebt", "債務本金＋融資租賃（營業租賃不計）", f"={I('debt')}+{I('lease')}")
vrow("netcash", "淨現金（負值＝淨債務）", f"=B{VR['cashx']}-B{VR['ndebt']}")
vrow("sh", "評價股數（M）＝封面流通股＋ASR 再交付＋預計歸屬 RSU＋選擇權庫藏股法淨股數", f"={I('sh_out')}+{I('asr_sh')}+{I('rsu_vest')}+MAX(0,{I('opt_n')}*(1-{I('opt_k')}/{I('px')}))")
vrow("dps", "期間股利（US$/股；持有期間）", f"={I('dps')}")
V.cell(VR["sumh"], 1, "DCF 彙總（三情境）").font = BOLD
for col, lab in zip("CDE", ("情境 A", "情境 B", "情境 C")): V[f"{col}{VR['sumh']}"] = lab
for si, s in enumerate("ABC"):
    col = "CDE"[si]
    pvr = DR[s]["pv"]
    V[f"{col}{VR['pvx']}"] = f"=SUM(DCF!F{pvr}:K{pvr})"
    V[f"{col}{VR['tvpv']}"] = f"=PL!K{PR[s]['nopat']}*(1+$B${VR['g']})*(1-$B${VR['g']}/$B${VR['ronic']})/($B${VR['wacc']}-$B${VR['g']})/(1+$B${VR['wacc']})^DCF!K{DR[s]['t']+0}"
    V[f"{col}{VR['ev']}"] = f"={col}{VR['pvx']}+{col}{VR['tvpv']}"
    V[f"{col}{VR['eq']}"] = f"={col}{VR['ev']}+$B${VR['netcash']}"
    V[f"{col}{VR['vps0']}"] = f"={col}{VR['eq']}/$B${VR['sh']}"
    V[f"{col}{VR['dcf12']}"] = f"={col}{VR['vps0']}*(1+$B${VR['ke']})-$B${VR['dps']}"
    V[f"{col}{VR['tvsh']}"] = f"={col}{VR['tvpv']}/{col}{VR['ev']}"
    # Comparable
    V[f"{col}{VR['rv']}"] = f"=((DATE(2027,1,31)-{I('px_date')})/365)*PL!F{PR[s]['rev']}+(1-(DATE(2027,1,31)-{I('px_date')})/365)*PL!G{PR[s]['rev']}"
    V[f"{col}{VR['psm']}"] = f"={I('ps_peer')}*{I('cm_adj')}"
    V[f"{col}{VR['ps_eq']}"] = f"={col}{VR['rv']}*{col}{VR['psm']}"
    V[f"{col}{VR['psc12']}"] = f"={col}{VR['ps_eq']}/$B${VR['sh']}"
    V[f"{col}{VR['mult']}"] = f"=$B${VR['pe_med']}"
    V[f"{col}{VR['ref_cur']}"] = f"={col}{VR['ntm']}*$B${VR['pe_own']}/$B${VR['sh']}"
    V[f"{col}{VR['ev_c']}"] = f"={col}{VR['ntm']}*{col}{VR['mult']}"
    V[f"{col}{VR['comp12']}"] = f"={col}{VR['ev_c']}/$B${VR['sh']}"
    V[f"{col}{VR['wa']}"] = None
vrow("pvx", "明確期 FCFF 現值（US$M）"); vrow("tvpv", "永續價值現值（US$M）"); vrow("ev", "企業價值 EV（US$M）"); vrow("eq", "股權價值（EV＋淨現金；US$M）")
vrow("vps0", "每股價值（評價日；US$）"); vrow("dcf12", "DCF 情境目標價（12 個月後＝每股價值×(1+ke)−股利；US$）"); vrow("tvsh", "永續價值占 EV 比例")
V.cell(VR["ch"], 1, "Comparable（目標時點 NTM 自由現金流 × 同業 EV／FCF）").font = BOLD
vrow("rv", "目標日 TTM 營收（FY2027E 與 FY2028E 按日數加權；US$M）"); vrow("psm", "套用倍數（同業中位 P／S，現價日 2026-11-13，WDAY 與 HUBS 逾期；分母期間 TTM；×rel_mult）"); vrow("ps_eq", "P／S 股權價值（目標日；US$M）"); vrow("psc12", "Comparable 情境目標價（P／S 主腿；US$；股權價值÷評價股數）")
vrow("ntm", "對照列（不進加權）：目標日 TTM 有槓桿自由現金流（FY2027E 與 FY2028E 按日數加權；US$M）"); vrow("mult", "對照列：套用倍數（同業中位 P／FCF，TTM；rel_mult＝1.00）"); vrow("ev_c", "對照列：P／FCF 股權價值（rel_mult＝1.00；US$M）"); vrow("comp12", "對照列：P／FCF 目標價（rel_mult＝1.00；US$；不進加權）")
vrow("ref_cur", "對照列：P／FCF 目標價（本公司現行相對值＝本公司現行 P／FCF；US$；不進加權）")
V.cell(VR["scenh"], 1, "情境權重與加權").font = BOLD
vrow("wa", "情境權重 A／B／C", f"={wr('A')}", f"={wr('B')}", f"={wr('C')}")
for col, nm in zip("CDE", ("A", "B", "C")): pass
V[f"B{VR['wa']}"] = None
V[f"C{VR['wa']}"] = f"={wr('A')}"; V[f"D{VR['wa']}"] = f"={wr('B')}"; V[f"E{VR['wa']}"] = f"={wr('C')}"
vrow("wb", "情境加權目標價（各情境 DCF 與 Comparable 加權；US$）", None, *(f"={I('w_dcf')}*{c}{VR['dcf12']}+(1-{I('w_dcf')})*{c}{VR['psc12']}" for c in "CDE"))
vrow("wc", "各情境上檔空間", None, *(f"={c}{VR['wb']}/$B${VR['px']}-1" for c in "CDE"))
vrow("chk_w", "檢查：情境加權目標價與加權目標價之差（應為 0）", f"=ABS({I('tp_wA')}*C{VR['wb']}+{I('tp_wB')}*D{VR['wb']}+{I('tp_wC')}*E{VR['wb']}-B{VR['tp']})")
V.cell(VR["peerh"], 1, "同業倍數（StockAnalysis；P／FCF＝市值÷（EV÷(EV／FCF)）；家數 6，少於 8 不做迴歸）").font = BOLD
for j, h in enumerate(["P／FCF（TTM 分母）", "市值（US$B）", "EV（US$B）", "EV／FCF（TTM）", "資料日", "會計準則", "租賃處理", "逾期（相對現價日超過 2 個交易日）"], 2):
    V.cell(VR["peerh"], j, h).font = BOLD
PEERS = [("ServiceNow（NOW）", 105.84, 107.29, 22.59, datetime.date(2026, 11, 6), "US GAAP", "營業租賃不計入債務（US GAAP；StockAnalysis EV 未載明）"),
         ("Adobe（ADBE）", 93.14, 94.28, 8.90, datetime.date(2026, 11, 6), "US GAAP", "營業租賃不計入債務（US GAAP；未載明）"),
         ("Workday（WDAY）", 50.25, 50.62, 17.70, datetime.date(2026, 10, 26), "US GAAP", "營業租賃不計入債務（US GAAP；未載明）"),
         ("Intuit（INTU）", 91.31, 92.54, 10.68, datetime.date(2026, 11, 13), "US GAAP", "營業租賃不計入債務（US GAAP；未載明）"),
         ("HubSpot（HUBS）", 13.16, 12.06, 15.13, datetime.date(2026, 10, 26), "US GAAP", "營業租賃不計入債務（US GAAP；未載明）"),
         ("SAP（SAP）", 273.96, 274.08, 26.73, datetime.date(2026, 10, 26), "IFRS", "IFRS 16 租賃負債通常計入債務（推定；未載明）")]
for i, (nm, mc, ev, evf, dt, acc, ls) in enumerate(PEERS):
    r_ = VR[f"pe{i+1}"]
    V.cell(r_, 1, nm)
    for j, v in zip((3, 4, 5), (mc, ev, evf)): V.cell(r_, j, v).font = BLUE
    V.cell(r_, 2, f"=C{r_}/(D{r_}/E{r_})")
    V.cell(r_, 6, dt).number_format = "yyyy-mm-dd"; V.cell(r_, 6).font = BLUE
    V.cell(r_, 7, acc); V.cell(r_, 8, ls)
    V.cell(r_, 9, f'=IF(NETWORKDAYS(F{r_},{I("px_date")})-1>2,"逾期","")')
vrow("pe_med", "同業中位數 P／FCF（TTM）", f"=MEDIAN(B{VR['pe1']}:B{VR['pe6']})")
vrow("pe_n", "家數", f"=COUNT(B{VR['pe1']}:B{VR['pe6']})")
r_ = VR["pe_own"]; V.cell(r_, 1, "Salesforce 自身（2026-11-13）P／FCF（TTM；StockAnalysis 市值 179.28、EV 210.26、EV／FCF 13.87）")
for j, v in zip((3, 4, 5), (179.28, 210.26, 13.87)): V.cell(r_, j, v).font = BLUE
V.cell(r_, 2, f"=C{r_}/(D{r_}/E{r_})"); V.cell(r_, 6, datetime.date(2026, 11, 13)).number_format = "yyyy-mm-dd"; V.cell(r_, 7, "US GAAP"); V.cell(r_, 8, "營業租賃不計入債務（推定）")
vrow("pe_ratio", "本公司現行倍數 ÷ 同業中位數（P／FCF）", f"=B{VR['pe_own']}/B{VR['pe_med']}")
vrow("pe_reval", "調整後 Comparable 需要的相對重估（同業中位數 ÷ 本公司現行倍數 − 1；僅供對照，基準改用歷史相對倍數）", f"=1/B{VR['pe_ratio']}-1")
vrow("ev_med", "對照：同業中位數 EV／FCF（TTM；舊口徑）", f"=MEDIAN(E{VR['pe1']}:E{VR['pe6']})")
vrow("ev_own", "對照：本公司 EV／FCF（TTM）", f"=E{VR['pe_own']}")
vrow("ev_ratio", "對照：EV／FCF 的本公司 ÷ 同業中位數", f"=B{VR['ev_own']}/B{VR['ev_med']}")
V.cell(VR["dh"], 1, "DCF 與 Comparable 的落差分解（情境 B；US$／股；公式）").font = BOLD
tt = f"PL!F{PR['B']['fcf_u']}"
vrow("d_sbc", "股票薪酬：P／FCF 對照口徑加回股票薪酬（DCF 不加回；以 rel_mult＝1.00 計）", f"=D{VR['mult']}*((DATE(2027,1,31)-{I('px_date')})/365*PL!F{PR['B']['sbc']}+(1-(DATE(2027,1,31)-{I('px_date')})/365)*PL!G{PR['B']['sbc']})/B{VR['sh']}")
vrow("d_lev", "槓桿：有槓桿自由現金流扣除稅後利息（Comparable 對股權；為負）", f"=D{VR['mult']}*((DATE(2027,1,31)-{I('px_date')})/365*PL!F{PR['B']['oth']}+(1-(DATE(2027,1,31)-{I('px_date')})/365)*PL!G{PR['B']['oth']})*(1-{I('tax')})/B{VR['sh']}")
vrow("d_time", "時點：Comparable 取目標日 TTM 現金流，相對以 FY2027E 現金流計的成長", f"=D{VR['mult']}*(((DATE(2027,1,31)-{I('px_date')})/365)*PL!F{PR['B']['fcff']}+(1-(DATE(2027,1,31)-{I('px_date')})/365)*PL!G{PR['B']['fcff']}-PL!F{PR['B']['fcff']})/B{VR['sh']}")
vrow("d_mult", "倍數與口徑殘差（P／S 主腿相對 DCF；含利潤率差、rel_mult、淨現金差、折現）", f"=D{VR['psc12']}-D{VR['dcf12']}-B{VR['d_sbc']}-B{VR['d_lev']}-B{VR['d_time']}")
vrow("d_tot", "落差合計（Comparable 目標價 − DCF 目標價；情境 B）", f"=D{VR['psc12']}-D{VR['dcf12']}")
V.cell(VR["sbh"], 1, "DCF（股票薪酬加回）讀數：只作對照，不進加權目標價（工作單 G5 第 3 條）").font = BOLD
for col, lab in zip("CDE", ("情境 A", "情境 B", "情境 C")): V[f"{col}{VR['sbh']}"] = lab
for si, s_ in enumerate("ABC"):
    col = "CDE"[si]
    V[f"{col}{VR['sb_pvx']}"] = f"={col}{VR['pvx']}+SUMPRODUCT(PL!F{PR[s_]['sbc']}:K{PR[s_]['sbc']},DCF!F{DR[s_]['frac']}:K{DR[s_]['frac']},DCF!F{DR[s_]['df']}:K{DR[s_]['df']})"
    V[f"{col}{VR['sb_tvpv']}"] = f"=(PL!K{PR[s_]['nopat']}+PL!K{PR[s_]['sbc']})*(1+$B${VR['g']})*(1-$B${VR['g']}/$B${VR['ronic']})/($B${VR['wacc']}-$B${VR['g']})/(1+$B${VR['wacc']})^DCF!K{DR[s_]['pv']+1}"
    V[f"{col}{VR['sb_ev']}"] = f"={col}{VR['sb_pvx']}+{col}{VR['sb_tvpv']}"
    V[f"{col}{VR['sb_eq']}"] = f"={col}{VR['sb_ev']}+$B${VR['netcash']}"
    V[f"{col}{VR['sb_vps0']}"] = f"={col}{VR['sb_eq']}/$B${VR['sh']}"
    V[f"{col}{VR['sb_dcf12']}"] = f"={col}{VR['sb_vps0']}*(1+$B${VR['ke']})-$B${VR['dps']}"
V.cell(VR["sb_pvx"], 1, "加回後明確期現值（FCFF＋股票薪酬；US$M）"); V.cell(VR["sb_tvpv"], 1, "加回後永續價值現值（(NOPAT＋股票薪酬)×(1−g÷RONIC)；US$M）")
V.cell(VR["sb_ev"], 1, "加回後企業價值（US$M）"); V.cell(VR["sb_eq"], 1, "加回後股權價值（US$M）"); V.cell(VR["sb_vps0"], 1, "加回後每股價值（評價日；US$；股數未計股票薪酬稀釋）")
V.cell(VR["sb_dcf12"], 1, "加回後 DCF 情境目標價（12 個月後；US$）")
vrow("sb_tp", "DCF（股票薪酬加回）目標價（情境加權；US$；不進加權目標價）", f"={wr('A')}*C{VR['sb_dcf12']}+{wr('B')}*D{VR['sb_dcf12']}+{wr('C')}*E{VR['sb_dcf12']}")
vrow("sb_up", "DCF（股票薪酬加回）目標價相對現價", f"=B{VR['sb_tp']}/B{VR['px']}-1")
vrow("sb_tpw", "對照：以加回版 DCF 取代 DCF 的加權目標價（US$；僅對照）", f"={I('w_dcf')}*B{VR['sb_tp']}+(1-{I('w_dcf')})*B{VR['comp_tp']}")
V.cell(VR["uh"], 1, "上檔空間三段歸因（原則第 47 條；加權上檔空間＝三項合計）").font = BOLD
vrow("u_dcf", "DCF 貢獻（DCF 目標價權重 ×（DCF 目標價 ÷ 現價 − 1））", f"={I('w_dcf')}*(B{VR['dcf_tp']}/B{VR['px']}-1)")
vrow("u_comp", "未調整 Comparable 貢獻（rel_mult＝1 的 Comparable）", f"=(1-{I('w_dcf')})*(B{VR['comp_tp']}/{I('cm_adj')}/B{VR['px']}-1)")
vrow("u_adj", "相對倍數調整貢獻（rel_mult − 1 的部分；基準為歷史相對倍數）", f"=(1-{I('w_dcf')})*(B{VR['comp_tp']}-B{VR['comp_tp']}/{I('cm_adj')})/B{VR['px']}")
vrow("u_tot", "加權上檔空間（檢查：等於上列三項合計）", f"=B{VR['u_dcf']}+B{VR['u_comp']}+B{VR['u_adj']}")
# DCF 頁公式補完（需 Val 位置）
for s in "ABC":
    d = DR[s]
    for y in YEARS[3:]:
        c = YC[y]; i = YEARS.index(y) - 3
        if y == "FY2027":
            Dp[f"{c}{d['frac']}"] = f"=(DATE(2027,1,31)-{I('px_date')})/365"
            Dp[f"{c}{d['t']}"] = f"={c}{d['frac']}/2"
        else:
            Dp[f"{c}{d['frac']}"] = 1
            Dp[f"{c}{d['t']}"] = f"=$F{d['frac']}+{i-1}+0.5"
        Dp[f"{c}{d['df']}"] = f"=1/(1+Val!$B${VR['wacc']})^{c}{d['t']}"
        Dp[f"{c}{d['pv']}"] = f"={c}{d['fcff']}*{c}{d['frac']}*{c}{d['df']}"
    # 永續價值折現時點＝FY2032 年底 ＝ 剩餘期間 ＋5 年
    Dp.cell(d["pv"] + 1, 1, f"情境 {s}｜永續價值折現時點（年）"); Dp[f"K{d['pv']+1}"] = f"=$F{d['frac']}+5"
    Dp.cell(d["pv"] + 2, 1, f"情境 {s}｜檢查：FCFF 各列加總（NOPAT＋折舊−資本支出−營運資金增加）−FCFF（應為 0）")
    for y in YEARS[3:]:
        c = YC[y]
        Dp[f"{c}{d['pv']+2}"] = f"={c}{d['fcff']}-(PL!{c}{PR[s]['nopat']}+PL!{c}{PR[s]['da']}-PL!{c}{PR[s]['capex']}-PL!{c}{PR[s]['dwc']})"
for si, s in enumerate("ABC"):
    col = "CDE"[si]
    V[f"{col}{VR['tvpv']}"] = f"=PL!K{PR[s]['nopat']}*(1+$B${VR['g']})*(1-$B${VR['g']}/$B${VR['ronic']})/($B${VR['wacc']}-$B${VR['g']})/(1+$B${VR['wacc']})^DCF!K{DR[s]['pv']+1}"
    V[f"{col}{VR['ntm']}"] = f"=((DATE(2027,1,31)-{I('px_date')})/365)*PL!F{PR[s]['fcf_cf']}+(1-(DATE(2027,1,31)-{I('px_date')})/365)*PL!G{PR[s]['fcf_cf']}"

# ---------------------------------------------------------------- RevDCF 頁
RD = SH["RevDCF"]
RV = {}
RD["A1"] = "反向 DCF：現價隱含什麼。以情境 B 路徑為底，兩個封閉解（DCF 對營業利益率與營收規模都是線性）；隱含 WACC 需整本重算，本版未做（Sens 頁有 WACC ±1 點案例）。"
RD["A2"] = "讀法：隱含值是「DCF 每股價值（評價日）＝現價」的點，不是評等門檻。"
rv_rows = [("lab", "項目"), ("v0", "情境 B 每股價值（評價日；US$）"), ("px", "現價（US$）"), ("gap", "現價與 DCF 每股價值之差（股權價值；US$M）"),
           ("k_m", "營業利益率全期平移 1 個百分點對股權價值的影響（US$M）"), ("d_m", "現價隱含的營業利益率平移（百分點）"), ("m_fy32", "情境 B FY2032E 營業利益率"), ("m_imp", "現價隱含的 FY2032 營業利益率"),
           ("ev0", "情境 B 企業價值（US$M）"), ("ev_imp", "現價隱含的企業價值（US$M）"), ("scale", "現價隱含的營收規模倍數（相對情境 B；FCFF 與營收同比例）"),
           ("mult_imp", "現價隱含的 P／FCF 倍數（現價市值÷目標日 TTM 有槓桿自由現金流，情境 B）"), ("mult_peer", "同業中位數 P／FCF（TTM）"), ("ps", "分部加總：不適用（本模型為單一營運分部，Comparable 以現金流倍數）")]
for i, (k, lab) in enumerate(rv_rows):
    RV[k] = 4 + i; RD.cell(4 + i, 1, lab)
b_ = PR["B"]
RD[f"B{RV['v0']}"] = f"=Val!D{VR['vps0']}"; RD[f"B{RV['px']}"] = f"=Val!B{VR['px']}"
RD[f"B{RV['gap']}"] = f"=(B{RV['px']}-B{RV['v0']})*Val!B{VR['sh']}"
# k_m：每年 rev*(1-t)*0.01*DF*frac 之和 ＋ 永續項
terms = "+".join(f"PL!{YC[y]}{b_['rev']}*DCF!{YC[y]}{DR['B']['frac']}*DCF!{YC[y]}{DR['B']['df']}" for y in YEARS[3:])
RD[f"B{RV['k_m']}"] = f"=0.01*(1-{I('tax')})*(({terms})+PL!K{b_['rev']}*(1+Val!B{VR['g']})*(1-Val!B{VR['g']}/Val!B{VR['ronic']})/(Val!B{VR['wacc']}-Val!B{VR['g']})/(1+Val!B{VR['wacc']})^DCF!K{DR['B']['pv']+1})"
RD[f"B{RV['d_m']}"] = f"=B{RV['gap']}/B{RV['k_m']}"
RD[f"B{RV['m_fy32']}"] = f"=PL!K{STR['營業利益率']}"
RD[f"B{RV['m_imp']}"] = f"=B{RV['m_fy32']}+B{RV['d_m']}/100"
RD[f"B{RV['ev0']}"] = f"=Val!D{VR['ev']}"
RD[f"B{RV['ev_imp']}"] = f"=B{RV['px']}*Val!B{VR['sh']}-Val!B{VR['netcash']}"
RD[f"B{RV['scale']}"] = f"=B{RV['ev_imp']}/B{RV['ev0']}"
RD[f"B{RV['mult_imp']}"] = f"=Val!B{VR['emkt']}/Val!D{VR['ntm']}"
RD[f"B{RV['mult_peer']}"] = f"=Val!B{VR['pe_med']}"
RD.cell(RV["ps"], 2, "—")

# ---------------------------------------------------------------- Recon 頁
Rc = SH["Recon"]
Rc["A1"] = "對帳：本模型 vs 公司指引與市場共識（共識與指引為 Interested-party；找不到者寫「未找到」）。口徑不同者不判定差異。"
RCR = {}
rec_rows = [
 ("h1", "1. 合併層：FY2027（情境 B）", None),
 ("rev27", "營收（US$M）", ("=PL!F%d" % STR["營收"], 46100, 46400, 46110, "公司指引 46.1–46.4B（Q2 FY27 新聞稿）；共識 46.11B 為財報前（Futurum 轉述），Interested-party")),
 ("om27", "GAAP 營業利益率", ("=PL!F%d" % STR["營業利益率"], 0.201, 0.201, "未找到", "公司指引 GAAP 20.1%（Q2 FY27 新聞稿）；本模型 FY2027 以上半年費用率外推，高於指引的差額來自下半年重組與費用假設")),
 ("eps27", "GAAP 稀釋 EPS（US$）", ("=PL!F%d" % STR["EPS（GAAP 稀釋，美元）"], 10.21, 10.25, "未找到", "公司 GAAP EPS 指引 10.21–10.25；本模型業外淨額與股數為假設（Assumed），差異主要來自股數與策略投資利得")),
 ("capex27", "資本支出占營收", ("=In!E%d" % IR["capex_pct"], 0.015, 0.015, "未找到", "公司指引約 1.5%，本模型直接採用（Interested-party）")),
 ("h2", "2. 共識（EPS 口徑不同，不判定）", None),
 ("c1", "FY2027 EPS 共識 16.7553；FY2028 EPS 共識 16.0067（247wallst 2026-09-29；口徑未載明，與 non-GAAP 指引 16.67–16.71 一致，含策略投資利得）", ("—", None, None, None, "本模型為 GAAP，不與 non-GAAP 共識直接比較；non-GAAP 預測本版未建")),
 ("c2", "FY2028 營收共識", ("=PL!G%d" % STR["營收"], None, None, "未找到", "FY2028 營收共識未找到（stockanalysis 頁面為 2025-04 舊資料，不採用）")),
 ("c3", "賣方平均目標價與家數", ("=Val!B%d" % VR["tp"], None, None, "未找到", "同日期可引用的平均目標價與家數未找到；247wallst 2026-09-29：56 則評等，40 買進、14 持有、2 強力賣出；TIKR 引 Guggenheim 2026-07-01 稱 Street 平均約 246 美元（單點）")),
]
r = 3
for k, lab, spec in rec_rows:
    RCR[k] = r
    Rc.cell(r, 1, lab)
    if spec is None:
        Rc.cell(r, 1).font = BOLD
        for j, h in enumerate(["本模型", "指引低", "指引高", "共識", "說明"], 2): Rc.cell(r, j, h).font = BOLD
    else:
        for j, v in enumerate(spec, 2):
            if v is not None: Rc.cell(r, j, v)
    r += 1
Rc.cell(r + 1, 1, "3. 同業日期與口徑").font = BOLD
Rc.cell(r + 2, 1, "同業 EV／FCF 取自 StockAnalysis 的日期為 2026-10-26 至 2026-11-13，與現價日不一致；家數 6，少於 8 不做迴歸；EV／FCF 的 FCF 口徑加回股票薪酬，本模型以 FCFF 加股票薪酬對應。未取得 NTM 口徑倍數（原則第 10 節第 7 條），以 TTM 倍數代理。")
Rc.cell(r + 3, 1, "4. Q3 FY27 實績（截至 2026-10-31）預計 2026-12 初公布；本版以 Q2 FY27 為最新已公布季。")
Rc.cell(r + 4, 1, "5. 新二分類 Agentforce Apps／Data 360 等 於 FY2026 全年與 Q3、Q4 FY26 未回溯：未找到，FY2027 以後才有可比基期。")
Rc.cell(r + 6, 1, "7. 其他對照與註記").font = BOLD
Rc.cell(r + 7, 1, "負債成本對照：本模型稅前 kd＝rf＋利差；獨立核對者稱公司有效利率約 4.8%（本代理未讀到原檔，未找到一手來源）；負債權重約 19%，兩者差距對 WACC 影響小。")
Rc.cell(r + 8, 1, "Hedging gains：Q2 FY27 新聞稿列 9；Q1 FY27 新聞稿我讀到的新二分類合計等於訂閱與支援 10,593、未單列 Hedging；獨立核對者稱 Q1 有 3 且 Apps／Data 與公布值有 1–2 的差異，本代理未能核對（Q1 10-Q 原檔頁碼待查），模型只用 Q2 與 FY2027 上半年欄。")
Rc.cell(r + 9, 1, "FY2028 起每年的 FCFF 逐年不連續：FY2027E 為上半年實績外推、FY2028 起為情境路徑，營運資金流入與業外淨額採新基準（Assumed）。")
RCR["dom_h"] = r + 11
Rc.cell(r + 11, 1, "8. 主導變數反解：使加權目標價＝現價所需的值（整本重算割線法；原則第 4 條；不是評等門檻）").font = BOLD
for j, h in enumerate(["基準值", "反解值", "單位", "說明"], 2): Rc.cell(r + 12, j, h).font = BOLD
kq = json.load(open(KREQ, encoding="utf-8")) if KREQ else None
for i_, row_ in enumerate(kq["rows"] if kq else []):
    for j, v in enumerate([row_["base"], row_["solved"], row_["unit"], row_["note"]], 2): Rc.cell(r + 13 + i_, j, v)
    Rc.cell(r + 13 + i_, 1, row_["label"])
if kq: Rc.cell(r + 13 + len(kq["rows"]), 1, "反解結果對應輸入指紋 " + str(kq["fingerprint"]) + "；" + ("與本檔一致" if kq["fingerprint"] == fp0 else "與本檔不一致（過期，須重跑）"))
Rc.cell(r + 13 + len(kq["rows"] if kq else []) + 2, 1, "9. DCF 股票薪酬加回對照（不進加權目標價）").font = BOLD
Rc.cell(r + 13 + len(kq["rows"] if kq else []) + 3, 1, f'="DCF（股票薪酬加回）每股價值（評價日）：情境 A／B／C＝"&TEXT(Val!C{VR["sb_vps0"]},"0.00")&"／"&TEXT(Val!D{VR["sb_vps0"]},"0.00")&"／"&TEXT(Val!E{VR["sb_vps0"]},"0.00")&" 美元；不加回＝"&TEXT(Val!C{VR["vps0"]},"0.00")&"／"&TEXT(Val!D{VR["vps0"]},"0.00")&"／"&TEXT(Val!E{VR["vps0"]},"0.00")&" 美元"')
Rc.cell(r + 13 + len(kq["rows"] if kq else []) + 4, 1, f'="DCF（股票薪酬加回）目標價（情境加權）"&TEXT(Val!B{VR["sb_tp"]},"0.00")&" 美元，相對現價 "&TEXT(Val!B{VR["sb_up"]},"0.0%")&"；不加回 DCF 目標價 "&TEXT(Val!B{VR["dcf_tp"]},"0.00")&" 美元；差額主要來自股票薪酬占營收約 8%（加回後 DCF 與 Comparable 的現金流口徑一致）。"')
Rc.cell(r + 5, 1, "6. 未找到清單：見 research_salesforce.md 第 9 節；另 Blume β 只有單一來源、公司債有效利率未找到、流通股數淨額於 ASR 結算後的股數未找到。")

# ---------------------------------------------------------------- Sens 頁
Se = SH["Sens"]
Se["A1"] = "敏感度：In 頁每個有「敏感度低／高」區間的輸入，一次動一個、整本 LibreOffice 重算；表中為加權目標價相對基準。情境專屬列與情境權重不列（三情境本身就是情境分析）。"
Se["A2"] = "白話標籤欄為「驅動」。群組：Tokenomics 每 MW 持有成本（IF_HoldEcon）本模型不適用（軟體公司不持有算力）。"
base_vals = {k: IR[k] for k in IR}
fp = fp0
BUILD_ID = f"salesforce-v{VER}-{fp}"
sens = json.load(open(SENS, encoding="utf-8")) if SENS else None
Se.cell(3, 1, "build_id"); Se.cell(3, 4, BUILD_ID)
Se.cell(4, 1, "模型狀態"); Se.cell(4, 4, "CURRENT" if (sens and sens.get("fingerprint") == fp) else "STALE")
Se.cell(5, 1, "基準加權目標價（本檔）"); Se.cell(5, 4, f"=Val!B{VR['tp']}")
Se.cell(6, 1, "敏感度重算時的基準加權目標價"); Se.cell(6, 4, sens["base_tp"] if sens else None)
Se.cell(7, 1, "敏感度重算時的輸入指紋"); Se.cell(7, 4, sens["fingerprint"] if sens else None)
hdr_r = 9
for j, h in enumerate(["驅動", "設定", "方向", "加權目標價", "相對基準", "鍵", "DCF 目標價", "DCF 變動（美元）", "投資建議"], 1):
    Se.cell(hdr_r, j, h).font = BOLD
SE_ROWS = []
if sens:
    rr = hdr_r + 1
    for case in sens["cases"]:
        Se.cell(rr, 1, case["label"]); Se.cell(rr, 2, case["setting"]); Se.cell(rr, 3, case["side"])
        Se.cell(rr, 4, case["tp"]); Se.cell(rr, 5, f"=D{rr}/$D$5-1"); Se.cell(rr, 6, case["key"]); Se.cell(rr, 7, case["dcf"]); Se.cell(rr, 8, f"=G{rr}-Val!$B${VR['dcf_tp']}"); Se.cell(rr, 9, case["rec"])
        SE_ROWS.append(rr); rr += 1
    Se.cell(rr + 1, 1, "註：tp_w（情境權重）與 scn_*（情境驅動）不列：三情境並列已涵蓋；he_*（Tokenomics 持有成本）不適用。")
SE_END = hdr_r + len(SE_ROWS)

# ---------------------------------------------------------------- Sources 頁
So = SH["Sources"]
So["A1"] = "來源與版本"
src_rows = [("Tokenomics 快照", f"{TK['source']['version']}（{TK['source']['file']}；commit {TK['source']['commit'][:7]}）；名稱：IF_TaskTokMult、IF_PriceRef_Luna／Sol／Astra、IF_FullCostDefault_Sol／Luna、IF_FullCost_Sol（data/tokenomics_snapshot_v22.0.json）"),
            ("研究包", "software_ai/salesforce/data/research_salesforce.md、facts_salesforce.csv"),
            ("公司申報", "Salesforce 8-K 附件 99.1（2026-08-26 Q2 FY27；2026-05 Q1 FY27；2026-02-25 Q4 FY26；2025-12-03 Q3 FY26）；10-K FY2026（2026-03-02）；10-K FY2025"),
            ("同業倍數", "StockAnalysis statistics 頁（NOW、ADBE、WDAY、INTU、HUBS、SAP、CRM）"),
            ("本版事前驗證", "未做（專案負責人 2026-11-16 23:27）")]
for i, (a, b2) in enumerate(src_rows, 3):
    So.cell(i, 1, a); So.cell(i, 2, b2)
So.cell(2, 1, "項目"); So.cell(2, 2, "內容")

# ---------------------------------------------------------------- README 頁
Rm = SH["README"]
Rm["A1"] = (f'="Salesforce 個股模型 v{VER}｜資料截止 "&TEXT({I("data_date")},"yyyy-mm-dd")&"（最新已公布季 Q2 FY27，截至 2026-07-31）｜成品日 "&TEXT({I("build_date")},"yyyy-mm-dd")&"｜股價基準日 "&TEXT({I("px_date")},"yyyy-mm-dd")&"｜build_id {BUILD_ID}｜本版未做事前驗證"')
Rm["A2"] = "命題、結論與重點觀察的數字由公式自本檔其他頁組字（不手寫）。"
VB = lambda k, col="B": f"Val!{col}{VR[k]}"
Rm["A3"] = ('="命題：AI 進入席位型軟體後，每單位營收＝AI 帶來的代理營收增量 − 席位被替代的侵蝕 − 推論成本；本模型以 Salesforce 的席位型營收（席位 × 每席位年費，扣替代）加上代理任務數 × 每任務單價推導 Agentforce Apps，其餘沿用公司揭露口徑。情境 B 在 FY"'
            f'&"2032 的 AI 淨效果為 "&TEXT(Seg!K{SR["B"]["net"]}/1000,"0.0")&" 十億美元（代理營收 "&TEXT(Seg!K{SR["B"]["incr"]}/1000,"0.0")&" − 席位侵蝕 "&TEXT(Seg!K{SR["B"]["eros"]}/1000,"0.0")&" − 推論成本 "&TEXT(Seg!K{SR["B"]["inf"]}/1000,"0.0")&"）。"')
Rm["A4"] = "命題的可證偽指標：每任務推論成本占實現單價的比例、席位年增、Agentforce 營收占比；任何一項偏離基準區間即須重估。"
Rm["A5"] = (f'="投資建議（在席位替代率落在情境 A 到 C 區間、推論成本取每任務 token 公司揭露推得值與 Tokenomics 市場價且不降價的基準假設下）："&{VB("rec")}&"；估值評等 "&{VB("val_rec")}&"；加權目標價 "&TEXT({VB("tp")},"0.00")&" 美元，上檔空間 "&TEXT({VB("up")},"0.0%")&"（現價 "&TEXT({VB("px")},"0.00")&" 美元，"&TEXT({VB("px_date")},"yyyy-mm-dd")&"）；收支關卡（FY2027–FY2031 自由現金流扣股利合計，情境 B）"&TEXT({VB("gate_b")}/1000,"0.0")&" 十億美元"')
Rm["A6"] = (f'="一、DCF 目標價 "&TEXT({VB("dcf_tp")},"0.00")&"、Comparable 目標價 "&TEXT({VB("comp_tp")},"0.00")&" 美元；情境 B 落差 "&TEXT({VB("d_tot","D") if False else VB("d_tot")},"0.0")&" 美元＝股票薪酬 "&TEXT({VB("d_sbc")},"0.0")&"、槓桿 "&TEXT({VB("d_lev")},"0.0")&"、時點 "&TEXT({VB("d_time")},"0.0")&"、倍數殘差 "&TEXT({VB("d_mult")},"0.0")&"。"')
Rm["A7"] = (f'="二、情境 B FY2027 營收 "&TEXT(PL!F{STR["營收"]}/1000,"0.0")&" 十億美元，公司指引 "&TEXT(Recon!C{RCR["rev27"]}/1000,"0.0")&" 至 "&TEXT(Recon!D{RCR["rev27"]}/1000,"0.0")&" 十億美元（Recon）；其中 Agentforce 代理營收約 "&TEXT(Seg!F{SR["B"]["agent"]}/1000,"0.0")&" 十億美元（公司自揭 ARR，Interested-party）。"')
Rm["A8"] = (f'="三、每任務推論成本（公司揭露的 token 數 × Tokenomics 市場價）占實現單價 "&TEXT(Seg!G{SR["B"]["inf"]}/Seg!G{SR["B"]["agent"]},"0%")&"；以模型商全成本計僅約 "&TEXT(Seg!G{SR["B"]["inf_full"]}/Seg!G{SR["B"]["agent"]},"0%")&"（只作下限對照）。"')
Rm["A9"] = (f'="四、現價隱含 FY2032 營業利益率 "&TEXT(RevDCF!B{RV["m_imp"]},"0.0%")&"（情境 B 為 "&TEXT(RevDCF!B{RV["m_fy32"]},"0.0%")&"）；永續價值占 EV 約 "&TEXT(Val!D{VR["tvsh"]},"0%")&"。"')
Rm["A10"] = (f'="五、上檔空間 "&TEXT({VB("u_tot")},"0.0%")&"＝DCF 貢獻 "&TEXT({VB("u_dcf")},"0.0%")&"＋未調整 Comparable 貢獻 "&TEXT({VB("u_comp")},"0.0%")&"＋相對倍數調整貢獻 "&TEXT({VB("u_adj")},"0.0%")&"；相對倍數 rel_mult＝"&TEXT({I("cm_adj")},"0.00")&"（PS 與 PSG 歷史相對值幾何平均，Derived；PS 單用 "&TEXT({I("rel_ps")},"0.00")&"、PSG 單用 "&TEXT({I("rel_psg")},"0.00")&"），套在同業中位 P／S 上（口徑一致，G6）；兩讀數分歧 "&TEXT({I("rel_div")},"0%")&IF({I("rel_div")}>0.3,"，超過 30%：原因為成長比 "&TEXT({I("rel_gr")},"0.00")&"（本公司營收年增低於同業）","，未超過 30%")&"；本公司現行 P／FCF 為同業中位數的 "&TEXT({VB("pe_ratio")},"0.00")&" 倍。"')
Rm["A12"] = "v1.3 展示層"
Rm["A13"] = (f'="Salesforce："&{VB("rec")}&"，加權目標價 "&TEXT({VB("tp")},"0.00")&" 美元（上檔 "&TEXT({VB("up")},"0.0%")&"）；本版未做事前驗證。"')
Rm["A14"] = "本版要點"
Rm["A15"] = "單一營運分部（10-K）；營收分部取公司自 FY2027 Q1 起揭露的新二分類（Agentforce Apps、Data 360 等）與營收型態；舊五分類只有歷史。"
Rm["A16"] = "命題子列（席位型、代理營收）合計等於 Agentforce Apps；子列為本模型拆分（Assumed），不是公司揭露。"
Rm["A17"] = "推論成本只讀 Tokenomics v22.0 快照，不自行估 token 價格；每任務 token 取公司揭露的累計 token÷累計 AWU（Derived／Interested-party），ptok 為 Analogy，待 AI供應鏈 登錄。"
Rm["A19"] = (f'="DCF（股票薪酬加回）每股價值（情境 B，評價日）"&TEXT({VB("sb_vps0","D")},"0.00")&" 美元、目標價（情境加權）"&TEXT({VB("sb_tp")},"0.00")&" 美元（相對現價 "&TEXT({VB("sb_up")},"0.0%")&"），不加回的 DCF 目標價 "&TEXT({VB("dcf_tp")},"0.00")&" 美元；僅作對照，不進加權目標價。"')
Rm["A18"] = "未做事項：事前驗證（專案負責人 指示本版跳過）；三層基準協調（Coord、C0、兩讀數）不適用：非 Siliconomics 覆蓋；Meta 的算力、Tokenomics 自用 AI 讀法 2、Wattonomics 對照不適用：Salesforce 不自建算力；非 GAAP 預測、新二分類 FY2026 回溯值：未找到。"
Rm["A20"] = "市場資料日期"
Rm["A21"] = (f'="現價 "&TEXT({I("px_date")},"yyyy-mm-dd")&" 收盤；無風險利率日期 "&TEXT({I("rf_date")},"yyyy-mm-dd")&IF(NETWORKDAYS({I("rf_date")},{I("px_date")})-1>2,"（逾期）","")&"；同業倍數日期 "&TEXT(MIN(Val!F{VR["pe1"]}:F{VR["pe6"]}),"yyyy-mm-dd")&" 至 "&TEXT(MAX(Val!F{VR["pe1"]}:F{VR["pe6"]}),"yyyy-mm-dd")&"，逾期 "&COUNTIF(Val!I{VR["pe1"]}:I{VR["pe6"]},"逾期")&" 家；同業 P／S 價格日 "&TEXT({I("ps_peer_date")},"yyyy-mm-dd")&"（逾期 2 家：WDAY、HUBS，2026-10-26）；股數取 10-Q 封面（Q1 FY27，2026-05-21）。"')
Rm["A22"] = (f'="主要失效條件（可觀察）：Q3 FY27 cRPO 年增（固定匯率）低於指引 "&TEXT({I("crpo_g")},"0%")&"；Agentforce ARR 單季增量轉負（Q2 FY27 逾 1.5 十億美元）；席位年增轉為負；每任務推論成本占單價超過本頁基準 "&TEXT(Seg!G{SR["B"]["inf"]}/Seg!G{SR["B"]["agent"]},"0%")&"。"')
Rm["A23"] = "風險登錄"
Rm["A24"] = "風險"; Rm["B24"] = "是否建模"; Rm["C24"] = "理由"; Rm["D24"] = "觀察指標"
risks = [("席位壓縮", "是", "rho 三情境", "席位年增、cRPO"), ("推論成本侵蝕毛利", "是", "Tokenomics 市場價與 token 數", "毛利率、每任務成本"), ("定價模式轉變（AELA、Flex Credits）", "部分", "以每任務單價與其年變動代理", "高階版採用率、Flex Credits 續購占比"),
         ("Informatica 與 Contentful、Fin 整合", "部分", "併入 Data 360 等 成長假設", "Data 360 等 年增（固定匯率）"), ("槓桿與回購", "否", "只影響淨債務與股數", "非流動債務、ASR 結算股數")]
for i, rk in enumerate(risks, 25):
    for j, v in enumerate(rk, 1): Rm.cell(i, j, v)
Rm.cell(31, 1, "工作表：README、In、Q、Seg、PL、DCF、Val、RevDCF、Recon、Sens、Sources")
Rm["A33"] = "頁"; Rm["B33"] = "輸入"; Rm["C33"] = "推導量"; Rm["D33"] = "因果方向"
mp = [("Seg", "gs、pu、rho_A／B／C、gt0_A／B／C、dp_A／B／C、agent27、p0、tok_task、ptok、px_dec、gd0、gd_term、ps_g", "席位型營收、代理營收、Data 360 等、推論成本、AI 淨效果", "席位 × 每席位年費＋任務數 × 單價 → Apps；任務數 × token × 市場價 → 推論成本"),
      ("PL", "cogs_dr、opx_dr、da_pct、capex_pct、wc_k、tax、oth27、oth_run", "毛利、營業利益、FCFF、淨利、EPS", "營收 − 成本 − 費用 → 營業利益 → NOPAT → FCFF"),
      ("Val", "rf、beta_raw、erp、kd_sp、g_tv、ronic_sp、w_dcf、cm_adj、策略投資比例", "WACC、DCF、Comparable、加權目標價、評等", "FCFF 折現＋同業倍數 × NTM 現金流 → 加權"),
      ("Sens", "In 頁敏感度低／高", "加權目標價、DCF 目標價、投資建議", "一次改一個輸入、整本重算")]
for i, rk in enumerate(mp, 34):
    for j, v in enumerate(rk, 1): Rm.cell(i, j, v)
Rm["A39"] = "版本沿革：v0.1（2026-11-16）首版；未做事前驗證；v0.2.3（2026-11-17）依工作單 G8 只改文字（個人情境用語中性化），數字不變；本版與其他版本無對照。"

# ---------------------------------------------------------------- 樣式與輸出
for ws_ in wb.worksheets:
    ws_.column_dimensions["A"].width = 52
    for col in "BCDEFGHIJK": ws_.column_dimensions[col].width = 14
wb["In"].column_dimensions["B"].width = 60; wb["In"].column_dimensions["H"].width = 70
for name, rng in (("PL_hdr", "PL!$C$5:$K$5"),):
    wb.defined_names[name] = DefinedName(name, attr_text=rng)
wb.save(OUT)
rowmap = {"IR": IR, "SR": SR, "PR": PR, "STR": STR, "SPR": SPR, "VR": VR, "DR": DR, "RV": RV, "RCR": RCR, "CK": CK, "QR": QR, "SE_ROWS": SE_ROWS, "SE_END": SE_END, "SCN_H": SCN_H, "SCN_ROWS": SCN_ROWS,
          "BUILD_ID": BUILD_ID, "FP": fp, "HDR_R": hdr_r}
if ROWMAP: json.dump(rowmap, open(ROWMAP, "w"), ensure_ascii=False, indent=1)
print("built", OUT, BUILD_ID)
