# Oracle（ORCL）個股模型 v0.4（命題＋營收分部架構；docs/plan/個股模型原則_命題加營收分部.md）
# 用法：python3 oracle/build_oracle.py raw.xlsx ；再以 LibreOffice headless 重算：
#   soffice --headless --calc --convert-to xlsx --outdir oracle/dist raw.xlsx
# 會計年度 FY(t)＝t−1 年 6 月～t 年 5 月（FY2027＝2026-06～2027-05）。金額：十億美元（US$B），MW 為 IT 口徑（MW-IT）。
# 工作表（順序固定）：README → In → Q → Thesis → Seg → PL → CF → Val → Recon → Sens → Sources
# 環境變數 SENS_OVR＝JSON {In 鍵: 值}：只替換基準值（敏感度建置用，scripts/sens_v04.py）；成品不設。
import sys, os, json
import openpyxl
from openpyxl.styles import Font, PatternFill
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.utils import get_column_letter

HERE = os.path.dirname(os.path.abspath(__file__))
OVR = json.loads(os.environ.get("SENS_OVR", "{}"))
TK = json.load(open(os.path.join(HERE, "data", "tokenomics_snapshot_v21.4.json"), encoding="utf-8"))["items"]
SENS_F = os.path.join(HERE, "data", "sens_v04.json")
SENS = json.load(open(SENS_F, encoding="utf-8")) if (os.path.exists(SENS_F) and not OVR) else None

BLUE = Font(color="0000FF"); BOLD = Font(bold=True); HDR = PatternFill("solid", fgColor="DDEBF7")
KEY = PatternFill("solid", fgColor="FFF2CC"); GREY = Font(color="808080")
wb = openpyxl.Workbook()
SC = ("A", "B", "C")
SCN = {"A": "情境 A（樂觀：依公司時程交付、不延誤）", "B": "情境 B（基準：依 FY27Q1 交付速度、延誤 3 個月）", "C": "情境 C（保守：OpenAI 相關容量實現 60%、延誤 6 個月）"}
YRS = list(range(2026, 2033))                       # Seg／PL／CF 年度欄：C＝FY2026A … I＝FY2032
yc = lambda y: get_column_letter(3 + y - 2026)
FY = [y for y in YRS if y >= 2027]
# Thesis 欄：C＝FY27Q1（估計）、D＝FY27 Q2–Q4（模型）、E＝FY2027 全年、F＝FY2028 … J＝FY2032
tc = lambda y: get_column_letter(5 + y - 2027)
GEN = ("h100", "gb200", "gb300", "vr200")
GNAME = {"h100": "Hopper H100", "gb200": "GB200 NVL72", "gb300": "GB300 NVL72", "vr200": "VR200 NVL72"}


def hdr(ws, row, labels, c0=1):
    for j, h in enumerate(labels, c0):
        c = ws.cell(row, j, h); c.font = BOLD; c.fill = HDR


def dn(name, ref):
    wb.defined_names[name] = DefinedName(name, attr_text=ref)


def tkv(name, g):
    return TK[name]["values"][GNAME[g]]["基準"]


# =====================================================================================
# In
# =====================================================================================
IN = wb.active; IN.title = "In"
IN["A1"] = "In — 輸入（藍字＝輸入；黑字＝公式；基準欄 E，低 D、高 F）。金額：十億美元（US$B）；MW＝IT 口徑（MW-IT）；每 MW 金額：百萬美元／MW·年（＝US$B／GW·年）"
IN["A2"] = "標記：Verified／Interested-party／Analogy／Assumed／Derived；Analogy、Assumed 給低／高區間。SemiAnalysis 未使用。「沿用 v0.3.1」＝沿用 oracle/legacy_v0.3 舊引擎 company.json 已核准的參數（原出處見來源欄）"
hdr(IN, 3, ["鍵", "項目", "單位", "低", "基準", "高", "標記", "來源1", "來源2／說明"])
rows = []; sec = lambda t: rows.append(("sec", t)); inp = lambda *a: rows.append(a)
PR_Q1 = "Oracle Q1 FY27 財報新聞稿（8-K 2026-09-10 附件 99.1）"
K10 = "Oracle FY2026 10-K（2026-06-22；附註 13 分部）"
Q10 = "Oracle FY27 Q1 10-Q（2026-09-11）"
FACTS = "oracle/data/oracle_facts_20261007.json"
CO = "oracle/legacy_v0.3/company.json"

sec("A. 股價、股數")
inp("px", "現價（2026-10-26 收盤）", "US$", None, 126.06, None, "Verified", "stockanalysis.com 歷史股價頁（https://stockanalysis.com/stocks/orcl/history/，2026-10-27 擷取）：2026-10-26 收盤 126.06", "markets/invest_map/data/prices.csv：2026-10-23 收盤 126.16（stockanalysis；stockscan 兩來源），與同頁 10-23 值相同；v0.3.1 為 10-21 收盤 129.13")
inp("sh_out", "普通股流通在外（2026-09-07）", "十億股", None, 3.023736, None, "Interested-party", Q10 + " 封面（" + FACTS + " sh.out.20260907）", "Q1 FY27 已含 ATM 發行 1.41 億股")
inp("opt_dil", "稀釋股數加項（選擇權等）", "十億股", None, 0.034, None, "Derived", "Q1 FY27 稀釋加權 3.000 − 基本加權 2.966（" + FACTS + " sh.dilW／sh.basicW.fy27q1）", "各年沿用")
inp("mcps_sh", "強制轉換特別股轉換股數（2029-01-15，FY2029 計入）", "十億股", None, 0.031238, None, "Derived", "Series D 50,000 股 × 轉換比率上限 624.7657（" + FACTS + " mcps.convShares）", "沿用 v0.3.1：一律轉股、取稀釋最大")

sec("B. 命題：已連網 MW 路徑與延誤（沿用 v0.3.1 scenarios.mwPath／delayMonths）")
inp("mw_fy26", "已連網 AI MW-IT（FY2026 年底）", "MW", None, 1835, None, "Derived", "FY26 交付 1.2 GW ÷ PUE 1.2 等（" + CO + " defaults.mwYearEnd）", "Oracle 未揭露累計在役 MW（不存在）；口徑不明者 ÷1.2")
inp("mw_q1", "已連網 AI MW-IT（FY27Q1 末，2026-08-31）", "MW", None, 2543, None, "Derived", "1,835 ＋ Q1 交付 850 MW ÷ 1.2＝708（" + CO + " fleet.openMw）", "Q1 交付 850 MW：Q1 FY27 法說（二手逐字稿，Interested-party）")
for s, p, c, d in (("A", 3333, 8778, 0), ("B", 2833, 8778, 3), ("C", 1000, 6284, 6)):
    inp(f"pace_{s}", f"情境 {s}：併網速度", "MW-IT／年", None, p, None, "Interested-party／Derived" if s != "C" else "Derived",
        {"A": "公司 Q1 預告每季約 1 GW × 4 ÷ 1.2", "B": "FY27Q1 實際 850 MW × 4 ÷ 1.2", "C": "FY26 歷史 1.2 GW ÷ 1.2"}[s], "沿用 v0.3.1（" + CO + " scenarios.mwPath.pace）")
    inp(f"cap_{s}", f"情境 {s}：合約容量上限（已簽站點）", "MW-IT", None, c, None, "Derived／Assumed" if s == "C" else "Derived",
        "2,543 ＋ Oracle 承租五個 Stargate 站點 8.1 GW（Epoch 估計，設施）÷1.2 − 已交付 Abilene 515" + ("；× OpenAI 相關實現 60%（Project Jupiter 不可抗力，2026-09-24；區間 50–70%）" if s == "C" else ""), "沿用 v0.3.1（" + CO + " scenarios.mwPath.contracted）；上限以後不假設新簽站點")
    inp(f"delay_{s}", f"情境 {s}：建設延誤（計費路徑後移月數）", "月", None, d, None, "Assumed",
        "沿用 v0.3.1（" + CO + " scenarios.delayMonths）：保守 6（Jupiter 天然氣管線延後近 6 個月）、基準 3、積極 0", "敏感度 0／3／6／12 月")
inp("u_bill", "計費比例（計費 MW ÷ 延誤後在役 MW）", "x", 0.90, 0.95, 1.00, "Assumed", "公司稱 AI 基礎設施利用率 97.9%（Q1 FY27 法說，Interested-party）；新交付產能約一季後才全額計費", "v0.3.1 以收斂比例 80–90% 處理（不同口徑）")

sec("C. 命題：世代組合（沿用 v0.3.1 fleet；Assumed）")
for g, v, lo, hi in (("h100", 0.20, 0.10, 0.30), ("gb200", 0.55, 0.45, 0.70), ("gb300", 0.25, 0.10, 0.35)):
    inp(f"om_{g}", f"Q1 末在役世代占比：{GNAME[g]}", "%", lo, v, hi, "Assumed", "Oracle 未揭露機隊世代組合（找不到）；Abilene（已連網約 20%）為 GB200（二手）", "沿用 v0.3.1（" + CO + " fleet.openMix）")
inp("nm27_gb300", "FY2027 新增產能中 GB300 占比（其餘 VR200）", "%", None, 1.0, None, "Assumed", "沿用 v0.3.1 fleet.newMix：FY27 GB300 100%", "VR200 2027 下半年起交付")
inp("nm28_gb300", "FY2028 新增產能中 GB300 占比（其餘 VR200）", "%", 0.0, 0.5, 1.0, "Assumed", "沿用 v0.3.1 fleet.newMix：FY28 GB300 50%／VR200 50%", "區間：VR200 提前或延後一年")
inp("nm29_gb300", "FY2029 起新增產能中 GB300 占比（其餘 VR200）", "%", 0.0, 0.0, 0.5, "Assumed", "沿用 v0.3.1 fleet.newMix：FY29–FY31 VR200 100%", "—")
inp("rf_h100_30", "FY2030 汰換：H100 在役量的比例（換成 VR200）", "%", None, 0.5, None, "Assumed", "GPU 經濟壽命 6 年（Tokenomics IF_DeprLifeIT；Oracle 伺服器耐用年限 6 年，10-K）：H100 約 FY24–FY25 上線", "v0.3.1 為觸頂後穩態汰換；本版改依世代年份（判斷，見報告第 5 節）")
inp("rf_h100_31", "FY2031 汰換：H100 在役量的比例", "%", None, 0.5, None, "Assumed", "同上", "—")
inp("rf_gb200_32", "FY2032 汰換：GB200 在役量的比例", "%", None, 1.0, None, "Assumed", "GB200 約 FY26 上線", "GB300 以後的汰換落在 FY2033 以後（模型期外；永續價值改用穩態汰換）")

sec("D. 命題：每 MW 經濟性（Tokenomics v21.4 快照，基準成本情境；修正一自 v19.0 升級，與 MAG 相同）")
for g in GEN:
    inp(f"he_{g}", f"每 MW 年持有成本（經濟口徑）IF_HoldEcon｜{GNAME[g]}", "US$M／MW·年", round(TK["IF_HoldEcon"]["values"][GNAME[g]]["低成本"], 4), round(tkv("IF_HoldEcon", g), 4), round(TK["IF_HoldEcon"]["values"][GNAME[g]]["高成本"], 4), "Derived", "oracle/data/tokenomics_snapshot_v21.4.json（Tokenomics v21.4，3ac70ac；複製自 microsoft/data）Interface!13 列", "低／高＝Tokenomics 低成本／高成本情境（只作敏感度）")
for g in GEN:
    inp(f"cx_{g}", f"每 MW IT 資本支出 IF_CapexIT｜{GNAME[g]}", "US$M／MW", None, round(tkv("IF_CapexIT", g), 4), None, "Derived", "同上 Interface!9 列（只含 IT；機房以租賃取得）", "沿用 v0.3.1 methodology.perMw.capexName")
for g in GEN:
    inp(f"ox_{g}", f"每 MW 年營運費用（不含折舊；含電費）IF_OpexGW｜{GNAME[g]}", "US$M／MW·年", None, round(tkv("IF_OpexGW", g), 4), None, "Derived", "同上 IF_OpexGW（電費、IT 與廠房維護、人員、稅與保險）", "本版新增：AI 營運成本由 Tokenomics 正向推導（v0.3.1 用 EBITDAR 率）")
inp("life_ai", "AI 設備折舊年限", "年", 5, 6, 6, "Derived", "Tokenomics IF_DeprLifeIT｜GB300 基準 6（高成本 4）", "Oracle 伺服器與網通設備耐用年限 6 年（10-K，" + FACTS + " ppe.serverLife）")
inp("p_long", "長約價格證據：IREN–Microsoft GB300 五年約", "US$M／MW·年", 8.9073, 9.70, 12.7247, "Derived", "IREN 新聞稿 2025-11-03：$9.7bn ÷ 5 年 ÷ 200 MW（含 20% 預付）", "沿用 v0.3.1 pricing.anchorMultiple（低／高為 v0.3 的 k 0.70／1.00 × v5.27 錨換成固定價格）")
inp("p_spot", "現貨價格證據：H100 Silicon Data 指數", "US$／GPU-hr", None, 2.82, None, "Verified", "Silicon Data H100 Index（2026-10-06）", "沿用 v0.3.1")
inp("gh_h100", "每 GPU 小時持有成本 IF_GPUhrEcon｜Hopper H100", "US$／GPU-hr", None, round(TK["IF_GPUhrEcon"]["values"]["Hopper H100"]["基準"], 5), None, "Derived", "Tokenomics v21.4 快照", "—")
inp("od_share", "隨需（現貨）占計費產能比例", "%", 0.0, 0.10, 0.20, "Assumed", "公司未揭露長約／隨需比例（找不到）；Oracle 以 take-or-pay 大型容量合約為主", "沿用 v0.3.1 pricing.anchorMultiple.onDemandShare")
inp("k_long", "長約定價倍數 k_長約（＝長約價 ÷ GB300 持有成本）", "x", "=p_long_lo/he_gb300", "=p_long/he_gb300", "=p_long_hi/he_gb300", "Derived", "契約第 4 條預設解讀：長約以略低於打平租金換取預付與五年確定性", "Oracle–OpenAI 合約隱含 k 約 1.22（4.5 GW 視為 IT）／1.47（視為設施），金額未經證實，只作敏感度")
inp("k_spot", "現貨定價倍數 k_現貨", "x", None, "=p_spot/gh_h100", None, "Derived", "H100 現貨指數 ÷ H100 每 GPU 小時持有成本", "—")
inp("k_mix", "定價倍數 k（隨需占比加權）", "x", None, "=od_share*k_spot+(1-od_share)*k_long", None, "Derived", "k＝隨需占比 × k_現貨 ＋（1 − 隨需占比）× k_長約", "每 MW 年收入＝在役世代加權 IF_HoldEcon × k；不以 Oracle 營收、RPO、指引反推")
inp("k_adj", "k 調整係數（敏感度用；基準 1）", "x", 0.9, 1.0, 1.2, "Assumed", "敏感度：k ×0.9／×1.2（1.2 約當 Oracle 自身合約隱含 k 1.22）", "—")
inp("lam", "成長型資本支出提前比例 λ（次期新增產能的設備在本期付款）", "%", None, 0.35, None, "Assumed", "沿用 v0.3.1 defaults.lambda", "—")

sec("E. 命題：租賃（未起租 288B；沿用 v0.3.1 leases.uncommenced）")
inp("lease_unc", "未起租租賃承諾（未折現）", "US$B", None, 288.0, None, "Interested-party", Q10 + " 附註 6（" + FACTS + " lease.uncommenced.fy27q1）", "FY27Q2 至 FY2029 起租、期限 15–19 年")
inp("lease_q", "未起租租約起租季數（FY27Q2 起平均）", "季", None, 11, None, "Interested-party", Q10 + " 附註 6：FY27Q2 至 FY2029 起租", "沿用 v0.3.1")
inp("lease_term", "未起租租約期限（直線付租）", "年", 15, 17, 19, "Assumed", "沿用 v0.3.1：區間中點 17 年", "11 季平均起租（FY27Q2–FY29Q4）；三情境相同（合約性，不隨 MW 縮放）")
inp("oplease_q1", "營業租賃負債（FY27Q1 末）", "US$B", None, 34.621, None, "Interested-party", Q10 + "（" + FACTS + " bs.opLease.fy27q1）", "—")
inp("finlease_q1", "融資租賃負債（FY27Q1 末；列入淨負債與 WACC 的負債）", "US$B", None, 9.185, None, "Interested-party", Q10 + "（" + FACTS + " bs.finLease.fy27q1）", "修正一：列入淨負債（F6）")
inp("lease_on", "在帳租賃負債（營業＋融資，FY27Q1 末）", "US$B", None, "=oplease_q1+finlease_q1", None, "Interested-party", Q10 + "（" + FACTS + " bs.opLease／bs.finLease.fy27q1）", "不含未起租")
inp("lease_r", "租賃折現率", "%", None, 0.057, None, "Verified", K10 + " 營業／融資租賃加權平均折現率 5.7%", "沿用 v0.3.1")
inp("oplcost_q1", "營業租賃費用（Q1 FY27）", "US$B", None, 1.101, None, "Interested-party", Q10 + "（" + FACTS + " lease.opCost.fy27q1）", "—")
inp("rent_on", "在帳租約年租金（營業租賃費用年化；算槓桿用）", "US$B", None, "=oplcost_q1*4", None, "Derived", Q10 + " 營業租賃費用 1.101 × 4", "已含在各分部成本內，只用於調整後槓桿")

sec("F. 命題：OCI 其他（非 AI）營收路徑")
inp("gna0", "OCI 其他（非 AI）年增：FY2027", "%", 0.10, 0.20, 0.30, "Assumed", "公司未揭露 OCI 的 AI／非 AI 拆分（不存在）；FY24 OCI（AI 規模化前）年增約 50%（含早期 AI）", "Q1 FY27 的非 AI 部分為殘差（Thesis 頁 C 欄），v0.3.1 非 AI 占比區間 12–30%")
inp("gnaLT", "OCI 其他（非 AI）年增：FY2032（線性收斂）", "%", 0.05, 0.08, 0.12, "Assumed", "一般 IaaS 長期成長", "—")

sec("G. 其他營收線：長期年增（FY2032；FY2027 起點＝近四季年增率，Q 頁 Derived，線性收斂）")
for k, nm, v, lo, hi, note in (("saas", "Cloud Applications（SaaS）", 0.08, 0.05, 0.10, "沿用 v0.3.1 legacyBiz SaaS gLT 8%"),
                                ("lic", "Software license", -0.02, -0.05, 0.0, "v0.3.1 為授權＋支援合併 −1%；本版拆開，授權 −2%（判斷）"),
                                ("sup", "Software support", -0.01, -0.03, 0.01, "v0.3.1 為授權＋支援合併 −1%"),
                                ("hw", "Hardware", 0.0, -0.03, 0.03, "沿用 v0.3.1 legacyBiz 硬體 0%"),
                                ("svc", "Services", 0.03, 0.0, 0.05, "沿用 v0.3.1 legacyBiz 服務 3%")):
    inp(f"gLT_{k}", f"{nm} 長期年增（FY2032）", "%", lo, v, hi, "Assumed", note, "FY2028–FY2031 由近四季年增率線性收斂")

sec("H. 分部利潤率、未分攤費用（Q 頁實績推得）")
inp("m_hw", "Hardware 分部利潤率", "%", None, "=Q!L17/Q!L8", None, "Derived", K10 + "：FY2026 硬體分部利潤 2,017 ÷ 營收 3,084", "Q1 FY27 為 58.4%")
inp("m_svc", "Services 分部利潤率", "%", None, "=Q!L18/Q!L9", None, "Derived", K10 + "：FY2026 服務分部利潤 1,533 ÷ 營收 5,743", "Q1 FY27 為 31.5%")
inp("m_sw", "Cloud and software 分部中「OCI AI 以外」的利潤率（以 Q1 FY27 校準）", "%", None, "=Thesis!C47", None, "Derived", "（Q1 FY27 分部利潤 9.358 − 命題估計的 Q1 OCI AI 貢獻）÷（Q1 分部營收 17.157 − 命題估計的 Q1 OCI AI 營收）；見 Thesis 頁 C 欄", "對照：FY2024（AI 規模化前）雲端與軟體分部利潤率 64.1%（Q 頁）")
inp("m_adj", "「OCI AI 以外」利潤率調整（敏感度用；基準 0）", "pt", -0.03, 0.0, 0.03, "Assumed", "SaaS 占比上升、授權與支援下滑的組合效果未分開建模", "—")
inp("un_rr", "未分攤費用 FY27 Q2–Q4 每季 ÷ Q1 實績", "x", 1.0, "=Q!L20/4/Q!I20", 1.35, "Assumed", "修正一預設：回到 FY2026 季均 4.353（Q1 FY27 3.527 偏低，可能含一次性因素）", "低＝Q1 實績水準（v0.4 原預設）；Sens 頁列示")
inp("un_g", "未分攤費用（研發、管理、攤銷、SBC 等）年增（FY2028 起）", "%", 0.02, 0.04, 0.06, "Assumed", "無：公司不揭露未分攤費用預算", "SBC 同比例")
inp("sbc_q1", "股份基礎給付 SBC（Q1 FY27）", "US$B", None, 1.127, None, "Interested-party", Q10 + "（" + FACTS + " is.sbc.fy27q1）", "FY2026 4.811")
inp("da_q1", "折舊與攤銷（Q1 FY27）", "US$B", None, "=Q!I24", None, "Interested-party", "Q 頁", "—")
inp("amort", "無形資產攤銷（年）", "US$B", None, "=(da_q1-dep_q1)*4", None, "Derived", "Q1 FY27 D&A 3.358 − 折舊 3.200（" + FACTS + " is.da／dep.fy27q1）× 4", "FY2026 1.671，逐年下降；各年沿用 Q1 年化（偏高）")
inp("dep_q1", "PP&E 折舊（Q1 FY27）", "US$B", None, 3.2, None, "Interested-party", Q10 + "（" + FACTS + " dep.fy27q1）", "—")
inp("dep_fy24", "PP&E 折舊（FY2024，AI 規模化前）", "US$B", None, 3.129, None, "Verified", K10 + "（" + FACTS + " is.dep.fy2024）", "—")
inp("nonai_dep_q", "非 AI 折舊（每季）", "US$B", None, "=dep_fy24/4", None, "Derived", K10 + "：FY2024（AI 規模化前）折舊 3.129 ÷ 4", "Q1 其餘折舊視為 AI 在役設備；FY2028 起隨非 AI 營收等比例")
inp("cip_q1", "在建工程（FY27Q1 末；視為 AI 設備，FY27 Q2–Q4 起投入使用）", "US$B", None, 48.546, None, "Interested-party", Q10 + "（" + FACTS + " bs.cip.fy27q1）", "判斷：Q1 末之前的支出不是已在 Q1 折舊內，就是在建工程")
inp("nonai_cx", "非 AI 資本支出 ÷ 非 AI 營收", "%", 0.05, 0.08, 0.12, "Assumed", "FY2024 資本支出 6.866 ÷ 營收 52.961＝13%（含早期 OCI 擴建）", "非 AI 折舊另以 FY2024 水準隨非 AI 營收等比例（不隨本參數變動）")

sec("I. 現金流、融資（沿用 v0.3.1 defaults；FY27 以 Q1 實績＋指引為錨）")
inp("capex_g27", "FY2027 資本支出（公司指引中點）", "US$B", 90, 92.5, 95, "Interested-party", "Q1 FY27 法說：FY27 資本支出 90–95（二手逐字稿；" + FACTS + " guide.fy27.capex）", "C10：近期以實績＋公司指引為錨；FY2028 起由命題推導")
inp("prepay_cov", "客戶出資覆蓋比（預付 ÷ AI 資本支出）", "%", 0.222, 0.243, 0.37, "Derived", "FY27 指引：資本支出 90–95、淨現金資本支出 ≤70 → 22.2–26.3% 中點（沿用 v0.3.1 defaults.prepay.capexCover）", "高＝Q1 實際 37.0%")
inp("prepay_open", "客戶預付（重大財務組成）累計現金（FY27Q1 末）", "US$B", None, 15.955, None, "Derived", "FY26 4.592 ＋ Q1 FY27 11.363（" + FACTS + " prepay.cash.cum）", "5 年認列為非現金營收（自營運現金扣除）")
inp("prepay_yrs", "預付認列年數", "年", 3, 5, 5, "Assumed", "沿用 v0.3.1：OpenAI 合約 2027 起 5 年", "—")
inp("etr", "有效稅率", "%", 0.126, 0.151, 0.17, "Interested-party", PR_Q1 + "：Q1 FY27 GAAP 有效稅率 15.1%", "FY2026 12.62%（10-K）")
inp("debt_q1", "借款本金（FY27Q1 末）", "US$B", None, 125.0, None, "Interested-party", Q10 + " 附註 2（" + FACTS + " bs.debtPar.fy27q1）", "到期再融資假設以相同利率展延")
inp("int_q1", "利息費用（Q1 FY27）", "US$B", None, "=Q!I26", None, "Interested-party", "Q 頁", "—")
inp("ii_q1", "利息收入（Q1 FY27）", "US$B", None, 0.306, None, "Interested-party", PR_Q1 + "（" + FACTS + " is.interestIncome.fy27q1）", "Q 頁未列（新聞稿只有 Q1）")
inp("r_old", "既有借款平均利率（Q1 利息費用年化 ÷ 本金）", "%", None, "=int_q1*4/debt_q1", None, "Derived", PR_Q1 + "：Q1 利息費用 1.428", "—")
inp("r_new", "新增投資級債利率", "%", 0.05, 0.058, 0.0811, "Assumed", "沿用 v0.3.1 新債票面 5.8%", "高＝6.55% 2046 票據殖利率 8.11%（2026-10-01，Benzinga）")
inp("r_junk", "高息債利率", "%", None, 0.12, None, "Assumed", "沿用 v0.3.1 defaults.junkRate", "—")
inp("r_cash", "現金利息收益率", "%", None, "=ii_q1*4/cashonly_q1", None, "Derived", PR_Q1 + "：Q1 利息收入 0.306 ÷ 期末現金 36.369 × 4", "—")
inp("cashonly_q1", "現金與約當現金（FY27Q1 末）", "US$B", None, 36.369, None, "Interested-party", Q10 + "（" + FACTS + " bs.cash.fy27q1）", "—")
inp("mktsec_q1", "有價證券（FY27Q1 末）", "US$B", None, 0.708, None, "Interested-party", Q10 + "（" + FACTS + " bs.mktSec.fy27q1）", "—")
inp("cash_q1", "現金與有價證券（FY27Q1 末）", "US$B", None, "=cashonly_q1+mktsec_q1", None, "Interested-party", Q10 + "（" + FACTS + " bs.cash／bs.mktSec.fy27q1）", "—")
inp("min_cash", "最低現金", "US$B", None, 10, None, "Assumed", "沿用 v0.3.1 defaults.minCash", "循環額度 10 不列入")
inp("lev_max", "投資級上限：（借款＋租賃負債）÷（EBITDA＋租金）", "x", 4.0, 4.5, 5.0, "Assumed", "沿用 v0.3.1 defaults.debtEbitdaMax（租賃調整後）", "只限制新增投資級債；超過走股權再走高息債")
inp("eq_cap", "每年新股上限（占市值）", "%", None, 0.05, None, "Assumed", "沿用 v0.3.1 defaults.eqCapPct", "Q1 FY27 已用 ATM 發行 19.9（不計入模型期）")
inp("eq_disc", "新股發行折價", "%", None, 0.10, None, "Assumed", "沿用 v0.3.1 defaults.eqDisc", "—")
inp("dps_q", "普通股每季股利", "US$／股", None, 0.50, None, "Interested-party", PR_Q1 + "（Q1 FY27 宣告）", "不回購（沿用 v0.3.1；授權餘額 6.3）")
inp("pref_q1", "特別股股利（Q1 FY27）", "US$B", None, 0.081, None, "Interested-party", Q10 + "（" + FACTS + " is.prefDiv.fy27q1）", "—")
inp("shd_q1", "稀釋加權股數（Q1 FY27）", "十億股", None, 3.000, None, "Interested-party", Q10 + "（" + FACTS + " sh.dilW.fy27q1）", "—")
inp("pref27", "特別股股利：FY27 Q2–Q4", "US$B", None, 0.244, None, "Derived", "6.50% Series D $5B × 6.5% × 3 季（沿用 v0.3.1 defaults.dividend.preferred）", "FY28 0.325、FY29 0.244（2029-01-15 轉換）")
inp("pref28", "特別股股利：FY2028", "US$B", None, 0.325, None, "Derived", "同上", "—")
inp("pref29", "特別股股利：FY2029", "US$B", None, 0.244, None, "Derived", "同上", "—")
inp("cfo_q1", "營運現金流（Q1 FY27）", "US$B", None, 23.103, None, "Interested-party", Q10 + "（含客戶預付 11.363）", "FY27 全年＝Q1 實績＋Q2–Q4 模型")
inp("capex_q1", "資本支出（Q1 FY27）", "US$B", None, 28.499, None, "Interested-party", Q10, "—")
inp("div_q1", "股利支付（普通股＋特別股，Q1 FY27）", "US$B", None, 1.565, None, "Interested-party", Q10, "—")

sec("J. 評價（沿用 v0.3.1 CAPM 輸入；權重為判斷預設）")
inp("rf", "無風險利率（美國 10 年期公債，2026-10-26）", "%", None, 0.05, None, "Verified", "美國財政部 10 年期公債殖利率 2026-10-26：5.00%（與 Amazon v0.3 同一輸入，跨模型一致；修正一追加）", "v0.4 初版為 2026-10-20 的 5.12%（oracle/data/oracle_facts_20261022.json val.ust10.treasury.20261020）")
inp("beta_r1", "原始 Beta 來源 1（5 年；Seeking Alpha／Yahoo）", "x", None, 1.77, None, "Interested-party", FACTS + " val.beta.sa／val.beta.yahoo（2026-10-07）", "兩站同值，視為一個來源")
inp("beta_r2", "原始 Beta 來源 2（5 年月資料）", "x", None, 1.712, None, "Interested-party", "oracle/data/oracle_facts_20261022.json val.beta.cand1（2026-10 初）", "—")
inp("beta", "Beta（Blume 調整：0.67 × 原始＋0.33，兩來源中位數）", "x", 1.2, "=MEDIAN(0.67*beta_r1+0.33,0.67*beta_r2+0.33)", 2.0, "Derived", "修正一（F4）：v0.4 原為原始 1.77", "Sens：1.5／2.0")
inp("erp", "股權風險溢酬", "%", None, 0.05, None, "Assumed", "沿用 v0.3.1", "—")
inp("kd", "稅前債務成本（約 10 年期）", "%", 0.05, "=r_new", 0.0811, "Assumed", "修正一（F4）：取約 10 年期口徑＝新增投資級債利率 5.8%（2026-02 發行 5.70% 2036 票據票面相近）；ORCL 10 年期票據現行殖利率：未找到", "高＝6.55% 2046 長債殖利率 8.11%（v0.4 原值）")
inp("wacc_adj", "WACC 調整（敏感度用；基準 0）", "pt", -0.01, 0.0, 0.01, "Assumed", "Sens：±1 個百分點、WACC × g 格", "—")
inp("ronic_adj", "永續期新投資報酬率 RONIC − WACC（基準 0）", "pt", None, 0.0, 0.05, "Assumed", "修正一（F5）：RONIC＝WACC（永續成長不創造價值）", "Sens：WACC＋5pt")
inp("val_date", "評價日（FY27Q1 季末）", "日期", None, "=DATE(2026,8,31)", None, "Verified", "FY27Q1 10-Q 資產負債表日", "—")
inp("tp_date", "目標價時點（約 12 個月後）", "日期", None, "=DATE(2027,10,31)", None, "Assumed", "今天 2026-10-27 ＋ 約 12 個月", "—")
inp("fy28_end", "FY2028 年底（分部加總錨定日）", "日期", None, "=DATE(2028,5,31)", None, "Verified", "會計年度 5 月底", "—")
inp("tp_g", "永續成長率", "%", 0.02, 0.03, 0.04, "Assumed", "沿用 v0.3.1 valuation.g", "—")
inp("w_dcf", "目標價中 DCF 的權重（其餘為分部加總）", "%", 0.3, 0.5, 0.7, "Assumed", "工作單預設 50%（與 AI半導體 v2「DCF＋Comparable 各半」一致；Comparable 在分部層做）", "Andy 可改")
for s, w in (("A", 0.25), ("B", 0.5), ("C", 0.25)):
    inp(f"w{s}", f"目標價情境權重 {s}", "%", None, w, None, "Assumed", "工作單預設 A 25%／B 50%／C 25%", "Andy 可改")
inp("ev_crwv", "CoreWeave 企業價值（2026-10-21）", "US$B", None, 80.12, None, "Interested-party", "oracle/data/oracle_facts_20261022.json val.peer.crwv", "—")
inp("eb_crwv", "CoreWeave 2027 年共識 EBITDA", "US$B", None, 16.656, None, "Interested-party", "coreweave/data/consensus_crwv_20260925.json FY27（MarketScreener）", "—")
inp("ev_nbis", "Nebius 企業價值（2026-10-21）", "US$B", None, 48.79, None, "Interested-party", "oracle/data/oracle_facts_20261022.json val.peer.nbis", "—")
inp("eb_nbis", "Nebius 2027 年共識 EBITDA", "US$B", None, 6.194, None, "Interested-party", "nebius/data/consensus_nbis_20261007.json FY27（MarketScreener）", "—")
inp("mult_oci", "OCI AI 前瞻 EV/EBITDA（neocloud 同業中位數；修正一起只作對照列）", "x", 4.81, "=MEDIAN(Val!D61:D62)", 7.88, "Derived", "CoreWeave、Nebius：EV（2026-10-21）÷ 2027 年共識 EBITDA（Val 頁同業表）", "第三家同業（IREN、WhiteFiber）EBITDA 共識：未找到；v0.3.1 為 6x [Assumed]")
inp("mult_sw", "其餘業務（軟體、硬體、服務）前瞻 EV/EBITDA（軟體同業中位數）", "x", 10.8, "=MEDIAN(Val!D65:D69)", 15.59, "Derived", "Microsoft、SAP、Salesforce、Adobe、ServiceNow NTM EV/EBITDA（2026-10-07；Val 頁同業表）", "硬體與服務以軟體倍數計（偏高，影響小，見報告第 5 節）")
inp("mult_leg", "傳統業務（授權、支援、硬體、服務）前瞻 EV/EBITDA", "x", 6.6, "=MEDIAN(Val!D68,Val!D69)", 10.8, "Derived", "軟體同業中成熟、低成長的兩家（Salesforce 10.8、Adobe 6.6）中位數；修正一（F2／核對意見 3）", "高成長的 SaaS 與 OCI 非 AI 仍用軟體中位數")
inp("mult_adj", "分部倍數調整係數（敏感度用；基準 1）", "x", 0.8, 1.0, 1.2, "Assumed", "敏感度 ±20%", "—")

sec("K. 共識與公司指引（只作 Recon 對帳，不作校準）")
inp("cons_tp", "市場共識目標價（S&P，43 家平均）", "US$", None, 237.36, None, "Interested-party", "S&P Global Market Intelligence 經 StockAnalysis（頁面表列 2026-10-20；oracle/data/consensus_orcl_20261022.json）", "MarketBeat 249.52（47 家）")
for y, r, e, eps in ((2027, 90.423, 51.942, 6.643), (2028, 131.711, 74.642, 9.144), (2029, 183.655, 107.543, 13.43)):
    inp(f"cr{y}", f"共識營收 FY{y}", "US$B", None, r, None, "Interested-party", "MarketScreener（S&P Capital IQ 系，2026-10-22；consensus_orcl_20261022.json）", "LSEG FY27 90.42、FY28 131.77" if y < 2029 else "第二來源：未找到")
    inp(f"ce{y}", f"共識 EBITDA FY{y}（非 GAAP 營業利益＋折舊）", "US$B", None, e, None, "Interested-party", "同上", "口徑：不含 SBC 費用")
    inp(f"cp{y}", f"共識 GAAP EPS FY{y}", "US$", None, eps, None, "Interested-party", "MarketScreener（2026-10-07 值沿用）", "調整後 EPS：FY27 8.14、FY28 11.00（LSEG）" if y < 2029 else "FY29 第二來源：未找到")
for y, cx, fc in ((2027, 93.961, -37.32), (2028, 100.962, -32.56), (2029, 84.848, 6.154)):
    inp(f"cx{y}", f"共識資本支出 FY{y}", "US$B", None, cx, None, "Interested-party", "MarketScreener（2026-10-07 值沿用；consensus_orcl_20261022.json）", "第二來源：未找到")
    inp(f"cf{y}", f"共識自由現金流 FY{y}", "US$B", None, fc, None, "Interested-party", "同上", "第二來源：未找到")
for y, v in ((2027, 32), (2028, 73), (2029, 114), (2030, 144)):
    inp(f"go{y}", f"公司 OCI 營收路徑 FY{y}", "US$B", None, v, None, "Interested-party", "Oracle Q4 FY26 新聞稿（2026-06-10；" + FACTS + " guide.ociPath.fy26pr）", "2025-10 分析師日：FY30 OCI 166、總營收 225" if y == 2030 else "—")
inp("g_rev27", "公司 FY2027 營收指引（下限）", "US$B", None, 90, None, "Interested-party", PR_Q1, "—")
inp("rpo", "剩餘履約義務 RPO（FY27Q1 末）", "US$B", None, 664, None, "Interested-party", Q10 + "：12 個月內 13%、13–36 個月 37%、37–60 個月 34%、其後 16%", "OpenAI 約占一半（二手，Moody's 判斷）")
inp("oai_annual", "Oracle–OpenAI 合約年額（約 $300bn ÷ 5 年）", "US$B／年", None, 60, None, "Interested-party", "WSJ 轉述（The Register 2025-09-11），未經當事方證實", "只作對照")
inp("tp_v031", "v0.3.1 基準目標價（加權）", "US$", 77.4, 84.2, 91.9, "Derived", "oracle/legacy_v0.3 交接檔 v0.3.1：保守 77.4／基準 84.2／積極 91.9（賣出）", "Recon v0.4 vs v0.3.1")

base_vals = {}
r = 4
for srow in rows:
    if srow[0] == "sec":
        IN.cell(r, 1, srow[1]).font = BOLD; r += 1; continue
    key, item, unit, lo, base, hi, tag, s1, s2 = srow
    if key in OVR:
        base = OVR[key]
    base_vals[key] = (lo, base, hi)
    IN.cell(r, 1, key); IN.cell(r, 2, item); IN.cell(r, 3, unit)
    for j, v in ((4, lo), (5, base), (6, hi)):
        if v is not None:
            c = IN.cell(r, j, v)
            if not (isinstance(v, str) and v.startswith("=")): c.font = BLUE
            if unit in ("%", "pt"): c.number_format = "0.0%"
            elif unit == "x": c.number_format = "0.000"
    IN.cell(r, 7, tag); IN.cell(r, 8, s1); IN.cell(r, 9, s2)
    dn(key, f"In!$E${r}")
    if lo is not None: dn(key + "_lo", f"In!$D${r}")
    if hi is not None: dn(key + "_hi", f"In!$F${r}")
    if key in ("w_dcf", "wA", "wB", "wC"):
        for c in range(1, 10): IN.cell(r, c).fill = KEY
    r += 1
for col, w in zip("ABCDEFGHI", [14, 60, 12, 10, 11, 10, 18, 80, 70]): IN.column_dimensions[col].width = w
IN.freeze_panes = "D4"


# =====================================================================================
# Q：實績（8 季＋FY2024–FY2026）
# =====================================================================================
Q = wb.create_sheet("Q")
Q["A1"] = "Q — 實績（十億美元）。會計年度 FY(t)＝t−1 年 6 月～t 年 5 月。藍字＝輸入；黑字＝公式。季度＝公司新聞稿「Supplemental analysis of GAAP revenues」；年度＝10-K（經查核）"
Q["A2"] = "資料取自 oracle/data/oracle_facts_20261007.json（v0.1a 起 SEC 一手資料：10-K [Verified]、10-Q 與新聞稿 [Interested-party]）。官方報告分部（10-K 附註 13）：雲端與軟體（Cloud and software）、硬體、服務；營收線（新聞稿）：Cloud Infrastructure、Cloud Applications、Software license、Software support、Hardware、Services"
QC = ["FY25 Q2（2024-09～11）", "FY25 Q3（～2025-02）", "FY25 Q4（～05）", "FY26 Q1（～08）", "FY26 Q2（～11）", "FY26 Q3（～2026-02）", "FY26 Q4（～05）", "FY27 Q1（2026-06～08）", "FY2024", "FY2025", "FY2026"]
hdr(Q, 3, ["項目"] + QC + ["標記", "來源", "說明"])
QK = ["fy25q2", "fy25q3", "fy25q4", "fy26q1", "fy26q2", "fy26q3", "fy26q4", "fy27q1", "fy2024", "fy2025", "fy2026"]
F_ = {it["id"]: it["value"] for g in json.load(open(os.path.join(HERE, "data", "oracle_facts_20261007.json"), encoding="utf-8"))["groups"].values() for it in g if isinstance(it, dict) and "id" in it}
SQ = "公司財報新聞稿（8-K 附件 99.1，各季）；FY2024–FY2026＝10-K"
def fv(fid, k):
    v = F_.get(f"{fid}.{k}")
    return None if v is None else v / 1000.0
def qline(r, label, fid, tag="Verified（年度）／Interested-party（季度）", src=SQ, note="", fmt="#,##0.000", vals=None):
    Q.cell(r, 1, label)
    for j, k in enumerate(QK):
        v = vals[j] if vals is not None else fv(fid, k)
        if v is None: continue
        c = Q.cell(r, 2 + j, v); c.number_format = fmt
        if not (isinstance(v, str) and v.startswith("=")): c.font = BLUE
    Q.cell(r, 13, tag); Q.cell(r, 14, src); Q.cell(r, 15, note)
LINES = [("Cloud Infrastructure（OCI）", "is.rev.iaas"), ("Cloud Applications（SaaS）", "is.rev.saas"), ("Software license", "is.rev.swLic"),
         ("Software support", "is.rev.swSup"), ("Hardware", "is.rev.hw"), ("Services", "is.rev.svc")]
for i, (lab, fid) in enumerate(LINES):
    qline(4 + i, lab, fid)
qline(10, "合併營收（公司揭露）", "is.rev.total")
cols = [get_column_letter(2 + j) for j in range(11)]
qline(11, "分部：Cloud and software（＝OCI＋SaaS＋授權＋支援）", None, tag="Derived", src=K10 + "：雲端與軟體分部營收＝雲端（IaaS＋SaaS）＋軟體（授權＋支援）", note="FY2026 58.530、Q1 FY27 17.157（10-K／10-Q 分部表）", vals=[f"=SUM({c}4:{c}7)" for c in cols])
qline(12, "分部：Hardware", None, tag="Derived", src=K10, vals=[f"={c}8" for c in cols])
qline(13, "分部：Services", None, tag="Derived", src=K10, vals=[f"={c}9" for c in cols])
qline(14, "檢查：營收線合計 − 合併營收（應為 0）", None, tag="—", src="—", vals=[f"=SUM({c}4:{c}9)-{c}10" for c in cols])
qline(15, "檢查：分部合計 − 合併營收（應為 0）", None, tag="—", src="—", vals=[f"=SUM({c}11:{c}13)-{c}10" for c in cols])
SEGM = "10-K 附註 13（年度）；10-Q 附註 9（Q1 FY27 與去年同季）"
qline(16, "分部利潤：Cloud and software", None, tag="Verified（年度）／Interested-party（季度）", src=SEGM, note="分部利潤只含直接成本，不含研發、管理、無形資產攤銷、重組、SBC、利息。FY25 Q2–Q4、FY26 Q2–Q4 季度分部利潤：未找到（10-Q 在本環境無法開啟）",
      vals=[None, None, None, 7.691, None, None, None, 9.358, 28.514, 30.930, 34.468])
qline(17, "分部利潤：Hardware", None, src=SEGM, note="季度（除 Q1 FY27）：未找到", vals=[None] * 7 + [0.452, 1.915, 1.918, 2.017])
qline(18, "分部利潤：Services", None, src=SEGM, note="季度（除 Q1 FY27）：未找到", vals=[None] * 7 + [0.445, 0.916, 0.993, 1.533])
qline(19, "分部利潤合計", None, tag="Derived", src="—", vals=[None] * 7 + [f"=SUM({c}16:{c}18)" for c in "IJKL"])
qline(20, "未分攤費用（＝分部利潤合計 − 營業利益）", None, tag="Derived", src="—", note="研發、管理、無形資產攤銷、重組、SBC、收購相關等", vals=[None] * 7 + [f"={c}19-{c}21" for c in "IJKL"])
qline(21, "營業利益（GAAP）", "is.ebit")
qline(22, "檢查：分部利潤三項合計 − 公司揭露分部利潤合計（FY2026 38.018，應為 0）", None, tag="—", src=K10, vals=[None] * 10 + ["=L19-L51"])
qline(23, "營業利益（非 GAAP）", "is.ebitNG", tag="Interested-party")
qline(24, "折舊與攤銷（D&A）", "is.da")
qline(25, "股份基礎給付（SBC）", "is.sbc")
qline(26, "利息費用", "is.intExp")
qline(27, "淨利（GAAP）", "is.ni")
qline(28, "營運現金流", "cf.cfo")
qline(29, "資本支出（現金）", "cf.capex")
qline(30, "自由現金流（營運現金流 − 資本支出）", None, tag="Derived", src="—", vals=[f"={c}28-{c}29" for c in cols])
qline(31, "股利支付（普通股＋特別股）", "cf.div", note="季度（除 Q1 FY27）：未找到（只取得年度與 Q1）")
qline(32, "買回普通股", "cf.buyback", note="季度（除 Q1 FY27）：未找到")
qline(33, "檢查：自由現金流 − 公司揭露值（應為 0）", None, tag="—", src="—",
      vals=[f"={get_column_letter(2+j)}30-{F_.get('cf.fcf.'+k)/1000}" if F_.get('cf.fcf.' + k) is not None else None for j, k in enumerate(QK)])
Q.cell(35, 1, "近四季年增率（FY26Q2–FY27Q1 合計 ÷ FY25Q2–FY26Q1 合計 − 1；Seg 頁 FY2027 起點；沿用 v0.3.1 g0 方法）").font = BOLD
for i, (lab, fid) in enumerate(LINES):
    rr = 36 + i
    Q.cell(rr, 1, f"{lab}：近四季年增")
    Q.cell(rr, 2, f"=SUM(F{4+i}:I{4+i})/SUM(B{4+i}:E{4+i})-1").number_format = "0.0%"
    Q.cell(rr, 13, "Derived")
Q.cell(43, 1, "利潤率對照").font = BOLD
Q.cell(44, 1, "Cloud and software 分部利潤率"); [Q.cell(44, 2 + j, f"={c}16/{c}11") for j, c in enumerate(cols) if c in "EIJKL"]
Q.cell(45, 1, "FY2024 Cloud and software 分部利潤率（AI 規模化前；In 頁 m_sw 的對照）"); Q["B45"] = "=J16/J11"
for rr in (44, 45):
    for c in range(2, 13): Q.cell(rr, c).number_format = "0.0%"
Q.cell(47, 1, "其他實績（PL 頁引用）").font = BOLD
qline(48, "特別股股利", None, vals=[None] * 7 + [0.081, None, None, 0.103], note="FY2026 特別股股利 103（損益表）；Q1 FY27 81")
qline(49, "稀釋加權股數（十億股）", None, vals=[None] * 7 + [3.000, 2.823, 2.866, 2.914])
qline(50, "稀釋 EPS（GAAP，美元）", None, vals=[None] * 7 + [1.56, 3.71, 4.34, 5.83], fmt="0.00")
qline(51, "公司揭露分部利潤合計", None, vals=[None] * 10 + [38.018], note="10-K 附註 13（" + FACTS + " seg.total.margin.fy2026）")
Q.column_dimensions["A"].width = 62
for j in range(11): Q.column_dimensions[get_column_letter(2 + j)].width = 12
Q.column_dimensions["M"].width = 22; Q.column_dimensions["N"].width = 60; Q.column_dimensions["O"].width = 60
Q.freeze_panes = "B4"
for k, rr in (("oci", 4), ("saas", 5), ("lic", 6), ("sup", 7), ("hw", 8), ("svc", 9), ("rev", 10), ("cs", 11)):
    dn(f"q1_{k}", f"Q!$I${rr}"); dn(f"a26_{k}", f"Q!$L${rr}"); dn(f"q0_{k}", f"Q!$E${rr}")
for i, k in enumerate(("oci", "saas", "lic", "sup", "hw", "svc")):
    dn(f"g0_{k}", f"Q!$B${36+i}")


# =====================================================================================
# Thesis：命題模組（OCI AI 算力）
# =====================================================================================
TH = wb.create_sheet("Thesis")
TH["A1"] = "Thesis — 命題模組：OCI 的已簽約 backlog（RPO 664B，OpenAI／Stargate 集中）能否以高於資金成本的報酬轉成營收與現金。只驅動 Cloud Infrastructure 營收線的「OCI AI 算力」子列；AI 資本支出接到 AI 折舊與分部利潤"
TH["A2"] = "已連網 MW（期末）＝MIN(合約上限, 前期＋併網速度 × 期間)；計費 MW＝（平均在役 − 延誤月數 ÷ 12 × 期間新增速度）× 計費比例；每 MW 年收入＝在役世代加權 IF_HoldEcon × k；營運費用＝在役 MW × IF_OpexGW；租金＝未起租 288B 依 11 季起租、17 年直線；折舊＝Q1 在役基數＋在建工程＋新資本支出 ÷ 6 年（當年半年）"
TCOLS = ["項目", "單位", "FY27 Q1（估計）", "FY27 Q2–Q4（模型）", "FY2027", "FY2028", "FY2029", "FY2030", "FY2031", "FY2032", "說明"]
hdr(TH, 3, TCOLS)
TL = {"C": 0.25, "D": 0.75, "E": 1, "F": 1, "G": 1, "H": 1, "I": 1, "J": 1}
TR = {}
r = 4
ALLC = "CDEFGHIJ"; FC = "FGHIJ"
def put(ws, row, lab, unit, vals, fmt="#,##0.000", note=None, bold=False):
    ws.cell(row, 1, lab).font = Font(bold=bold); ws.cell(row, 2, unit)
    for c, v in vals.items():
        if v is None: continue
        x = ws[f"{c}{row}"]; x.value = v; x.number_format = fmt
    if note: ws.cell(row, 11 if ws.title == "Thesis" else 10, note)
for s in SC:
    TH.cell(r, 1, SCN[s]).font = BOLD; r += 1
    R = {}
    def nr(k):
        global r
        R[k] = r; r += 1; return R[k]
    cap, pace, dly = f"cap_{s}", f"pace_{s}", f"delay_{s}"
    # 期間長度
    put(TH, nr("L"), "期間長度", "年", {c: TL[c] for c in ALLC}, "0.00")
    # 已連網 MW（期末）
    v = {"C": "=mw_q1", "D": f"=MIN({cap},mw_q1+{pace}*D{R['L']})", "E": f"=D{R['L']+1}"}
    prevc = {"F": "E", "G": "F", "H": "G", "I": "H", "J": "I"}
    for c in "FG": v[c] = f"=MIN({cap},{prevc[c]}{R['L']+1}+{pace}*{c}{R['L']})"
    for k_, c in enumerate("HIJ", 1):
        p_ = f"{prevc[c]}{R['L']+1}"; inc = f"MAX(0,MIN({cap},{p_}+{pace})-{p_})"
        v[c] = f"={p_}+{inc}+(tp_g*{p_}-{inc})*{k_}/3"
    put(TH, nr("conn"), "已連網 AI MW-IT（期末）", "MW", v, "#,##0", "FY2028–29：MIN(上限 " + cap + ", 前期＋速度)；FY2030–32：新增量線性收斂到「前期 × 永續成長 g」（修正一 F5；超過已簽上限的部分視為新簽約）")
    v = {"C": f"=C{R['conn']}-mw_fy26", "D": f"=D{R['conn']}-C{R['conn']}", "E": f"=E{R['conn']}-mw_fy26"}
    for c in FC: v[c] = f"={c}{R['conn']}-{prevc[c]}{R['conn']}"
    put(TH, nr("add"), "期間新增已連網 MW", "MW", v, "#,##0")
    v = {"C": f"=(mw_fy26+C{R['conn']})/2", "D": f"=(C{R['conn']}+D{R['conn']})/2", "E": f"=(C{{a}}*C{R['L']}+D{{a}}*D{R['L']})/E{R['L']}"}
    for c in FC: v[c] = f"=({prevc[c]}{R['conn']}+{c}{R['conn']})/2"
    ra = nr("avg"); v["E"] = v["E"].replace("{a}", str(ra))
    put(TH, ra, "平均已連網 MW", "MW", v, "#,##0")
    v = {"C": f"=C{ra}", "D": f"=MAX(0,D{ra}-{dly}/12*D{R['add']}/D{R['L']})", "E": f"=(C{{x}}*C{R['L']}+D{{x}}*D{R['L']})/E{R['L']}"}
    for c in FC: v[c] = f"=MAX(0,{c}{ra}-{dly}/12*{c}{R['add']}/{c}{R['L']})"
    rl = nr("lag"); v["E"] = v["E"].replace("{x}", str(rl))
    put(TH, rl, "延誤後平均在役 MW（Q1 為實績期，不延誤）", "MW", v, "#,##0", "延誤月數：In 頁 " + dly)
    put(TH, nr("bill"), "計費 MW（× 計費比例）", "MW", {c: f"={c}{rl}*u_bill" for c in ALLC}, "#,##0")
    # 世代存量（期末）
    TH.cell(r, 1, "世代存量（期末，MW）：Q1 末依 In 頁占比；新增依 In 頁新增組合；汰換換成 VR200").font = GREY; r += 1
    for g in GEN: R[f"st_{g}"] = r + GEN.index(g)
    cR = R["conn"]; aR = R["add"]
    rf = {"H": f"C{R['st_h100']}*rf_h100_30", "I": f"C{R['st_h100']}*rf_h100_31", "J": f"C{R['st_gb200']}*rf_gb200_32"}
    for g in GEN:
        rr = R[f"st_{g}"]
        v = {"C": f"=C{cR}*om_{g}" if g != "vr200" else "=0"}
        if g in ("h100", "gb200"):
            v["D"] = f"=C{rr}"; v["E"] = f"=D{rr}"
            for c in FC: v[c] = f"={prevc[c]}{rr}" + (f"-{rf[c]}" if (g == "h100" and c in "HI") or (g == "gb200" and c == "J") else "")
        elif g == "gb300":
            v["D"] = f"=C{rr}+D{aR}*nm27_gb300"; v["E"] = f"=D{rr}"
            v["F"] = f"=E{rr}+F{aR}*nm28_gb300"
            for c in "GHIJ": v[c] = f"={prevc[c]}{rr}+{c}{aR}*nm29_gb300"
        else:
            v["D"] = f"=C{rr}+D{aR}*(1-nm27_gb300)"; v["E"] = f"=D{rr}"
            v["F"] = f"=E{rr}+F{aR}*(1-nm28_gb300)"
            for c in "GHIJ": v[c] = f"={prevc[c]}{rr}+{c}{aR}*(1-nm29_gb300)" + (f"+{rf[c]}" if c in rf else "")
        put(TH, rr, f"  {GNAME[g]}", "MW", v, "#,##0")
    r += 4
    put(TH, nr("stchk"), "檢查：世代存量合計 − 已連網 MW（應為 0）", "MW", {c: f"=SUM({c}{R['st_h100']}:{c}{R['st_vr200']})-{c}{cR}" for c in ALLC}, "0.000")
    # 平均世代占比 → 每 MW 錨、營運費用
    def mixsum(c, pref):
        if c == "C":
            return "+".join(f"om_{g}*{pref}_{g}" for g in GEN if g != "vr200")
        p = {"D": "C", "F": "E", "G": "F", "H": "G", "I": "H", "J": "I"}[c]
        return "(" + "+".join(f"({p}{R['st_'+g]}+{c}{R['st_'+g]})*{pref}_{g}" for g in GEN) + f")/({p}{cR}+{c}{cR})"
    v = {c: "=" + mixsum(c, "he") for c in "CDFGHIJ"}
    rh = nr("anc"); v["E"] = f"=(C{rh}*C{R['bill']}*C{R['L']}+D{rh}*D{R['bill']}*D{R['L']})/(C{R['bill']}*C{R['L']}+D{R['bill']}*D{R['L']})"
    put(TH, rh, "每 MW 錨：在役世代加權 IF_HoldEcon", "US$M／MW·年", v, "0.00", "期初＋期末存量平均加權（Tokenomics v21.4 基準成本）")
    put(TH, nr("pm"), "每 MW 年收入（錨 × k × k 調整）", "US$M／MW·年", {c: f"={c}{rh}*k_mix*k_adj" for c in ALLC}, "0.00", "k：In 頁 k_mix")
    v = {c: f"={c}{R['bill']}*{c}{R['pm']}*{c}{R['L']}/1000" for c in "CDFGHIJ"}; v["E"] = f"=C{r}+D{r}"
    put(TH, nr("rev"), "OCI AI 算力營收（命題子列）", "US$B", v, "#,##0.00", "Q1 為估計（Q1 實際 OCI 7.388 的 AI 部分；其餘為非 AI 殘差）")
    v = {c: "=" + mixsum(c, "ox") for c in "CDFGHIJ"}
    rox = nr("oxm"); v["E"] = f"=(C{rox}*C{rl}*C{R['L']}+D{rox}*D{rl}*D{R['L']})/(C{rl}*C{R['L']}+D{rl}*D{R['L']})"
    put(TH, rox, "每 MW 年營運費用：在役世代加權 IF_OpexGW", "US$M／MW·年", v, "0.000")
    v = {c: f"={c}{rl}*{c}{rox}*{c}{R['L']}/1000" for c in "CDFGHIJ"}; v["E"] = f"=C{r}+D{r}"
    put(TH, nr("opex"), "AI 營運費用（電費、維護、人員、稅與保險；不含折舊、租金）", "US$B", v, "#,##0.00", "延誤後在役 MW × 每 MW 營運費用")
    full = "lease_unc/lease_term"
    v = {"C": "=0", "D": f"={full}*MIN(1+2+3,lease_q*(lease_q+1)/2)/(lease_q*4)", "E": "=D{}", "F": f"={full}*MIN(4+5+6+7,lease_q*4)/(lease_q*4)", "G": f"={full}*MIN(8+9+10+11,lease_q*4)/(lease_q*4)"}
    for c in "HIJ": v[c] = f"={full}"
    for c in "FGHIJ": v[c] += f"+MAX(0,{c}{ra}-cap_B)*{full}/(cap_B-mw_q1)"
    rr = nr("rent"); v["E"] = f"=D{rr}"
    put(TH, rr, "AI 機房租金（未起租 288B 依序起租＋超過已簽範圍的新租約）", "US$B", v, "#,##0.00", "11 季平均起租、17 年直線；平均在役超過 cap_B 的部分以同一每 MW 租金加計（修正一）")
    # 資本支出
    v = {"C": f"=capex_q1-nonai_cx*(q1_rev-C{R['rev']})", "E": f"=capex_g27-nonai_cx*(Seg!D{{tot}}-E{R['rev']})"}
    rcx = nr("cxg")
    v["D"] = f"=E{rcx}-C{rcx}"
    nxt = {"F": "G", "G": "H", "H": "I", "I": "J", "J": None}
    cxn = {"F": "(nm28_gb300*cx_gb300+(1-nm28_gb300)*cx_vr200)", "G": "(nm29_gb300*cx_gb300+(1-nm29_gb300)*cx_vr200)"}
    for c in "HIJ": cxn[c] = cxn["G"]
    for c in FC:
        v[c] = f"=({c}{aR}*(1-lam)*{cxn[c]}" + (f"+{nxt[c]}{aR}*lam*{cxn[nxt[c]]}" if nxt[c] else f"+{c}{aR}*(1+tp_g)*lam*{cxn[c]}") + ")/1000"
    put(TH, rcx, "AI 成長型資本支出（FY27＝指引中點 − 非 AI；其後＝新增 MW × IF_CapexIT，λ 提前付款）", "US$B", v, "#,##0.00", "FY2027：In 頁 capex_g27（三情境相同）")
    v = {c: "=0" for c in "CDEFG"}; v.update({"H": f"={rf['H']}*cx_vr200/1000", "I": f"={rf['I']}*cx_vr200/1000", "J": f"={rf['J']}*cx_vr200/1000"})
    put(TH, nr("cxr"), "AI 汰換資本支出（依世代年份，換成 VR200）", "US$B", v, "#,##0.00")
    put(TH, nr("cx"), "AI 資本支出合計", "US$B", {c: f"={c}{rcx}+{c}{R['cxr']}" for c in ALLC}, "#,##0.00")
    # 折舊
    put(TH, nr("dbase"), "Q1 在役 AI 設備折舊（年率；扣汰換世代）", "US$B／年",
        {**{c: "=(dep_q1-nonai_dep_q)*4" for c in "CDEFG"},
         "H": f"=MAX(0,(dep_q1-nonai_dep_q)*4-rf_h100_30*C{R['st_h100']}*cx_h100/life_ai/1000)",
         "I": f"=MAX(0,(dep_q1-nonai_dep_q)*4-(rf_h100_30+rf_h100_31)*C{R['st_h100']}*cx_h100/life_ai/1000)",
         "J": f"=MAX(0,(dep_q1-nonai_dep_q)*4-(rf_h100_30+rf_h100_31)*C{R['st_h100']}*cx_h100/life_ai/1000-rf_gb200_32*C{R['st_gb200']}*cx_gb200/life_ai/1000)"}, "#,##0.00",
        "Q1 折舊 3.2 − 非 AI 折舊（FY2024 水準）；汰換世代的舊折舊扣除")
    v = {"C": f"=C{R['dbase']}*C{R['L']}", "D": f"=D{R['dbase']}*D{R['L']}+(cip_q1+D{R['cx']})/life_ai*D{R['L']}/2"}
    for c in FC:
        prior = "+".join(f"{x}{R['cx']}" for x in "FGHIJ"[:"FGHIJ".index(c)])
        v[c] = f"={c}{R['dbase']}+(cip_q1+D{R['cx']}" + (f"+{prior}" if prior else "") + f")/life_ai+{c}{R['cx']}/life_ai/2"
    rd = nr("dep"); v["E"] = f"=C{rd}+D{rd}"
    put(TH, rd, "AI 折舊（進分部利潤與營業利益）", "US$B", v, "#,##0.00", "在建工程 48.5 與 FY27 Q2–Q4 起的資本支出 ÷ 6 年，當年半年")
    put(TH, nr("ebitda"), "OCI AI EBITDA（租金後）", "US$B", {c: f"={c}{R['rev']}-{c}{R['opex']}-{c}{R['rent']}" for c in ALLC}, "#,##0.00")
    put(TH, nr("contrib"), "OCI AI 分部利潤貢獻（營業利益口徑）", "US$B", {c: f"={c}{R['ebitda']}-{c}{rd}" for c in ALLC}, "#,##0.00", "進 PL 頁 Cloud and software 分部利潤")
    put(TH, nr("om"), "OCI AI 營業利益率", "%", {c: f"=IF({c}{R['rev']}=0,0,{c}{R['contrib']}/{c}{R['rev']})" for c in ALLC}, "0.0%")
    # 非 AI
    v = {"C": f"=q1_oci-C{R['rev']}", "D": f"=C{{o}}*((1+gna0)^0.25+(1+gna0)^0.5+(1+gna0)^0.75)", "E": "=C{o}+D{o}"}
    ro = nr("oth")
    v = {k: x.replace("{o}", str(ro)) for k, x in v.items()}
    gr = ro + 1
    for i, c in enumerate(FC, 1):
        v[c] = f"={prevc[c]}{ro}*(1+{c}{gr})"
    put(TH, ro, "OCI 其他（非 AI）營收（子列）", "US$B", v, "#,##0.00", "Q1＝實際 OCI − 命題 AI 估計（殘差）；之後依成長路徑")
    put(TH, nr("othg"), "  OCI 其他年增", "%", {c: f"=gna0+(gnaLT-gna0)*{i}/5" for i, c in enumerate(FC, 1)}, "0.0%")
    put(TH, nr("oci"), "Cloud Infrastructure（OCI）營收", "US$B", {c: f"={c}{R['rev']}+{c}{ro}" for c in ALLC}, "#,##0.00")
    put(TH, nr("ocim"), "OCI（含非 AI）營業利益率（非 AI 以 m_sw 計）", "%", {c: f"=({c}{R['contrib']}+(m_sw+m_adj)*{c}{ro})/{c}{R['oci']}" for c in ALLC}, "0.0%")
    # 報酬
    put(TH, nr("nopat"), "OCI AI 稅後營業利益", "US$B", {c: f"={c}{R['contrib']}*(1-etr)" for c in ALLC}, "#,##0.00")
    v = {"C": "=(dep_q1-nonai_dep_q)*4*life_ai", "D": f"=C{{k}}+cip_q1", "E": f"=D{{k}}+D{R['cx']}"}
    rk = nr("cap")
    v = {k_: x.replace("{k}", str(rk)) for k_, x in v.items()}
    for c in FC: v[c] = f"={prevc[c]}{rk}+{prevc[c] if prevc[c] != 'E' else 'E'}{R['cx']}*0" if False else f"={prevc[c]}{rk}+{c}{R['cx']}"
    v["F"] = f"=E{rk}+F{R['cx']}"
    put(TH, rk, "AI 累計投入資本（毛額，期末）", "US$B", v, "#,##0.0", "Q1 在役基數（年折舊 × 6）＋在建工程＋其後資本支出；機房以租賃取得不計")
    put(TH, nr("roic"), "AI 增量 ROIC（稅後營業利益 ÷ 期初投入資本毛額）", "%", {c: f"={c}{R['nopat']}/{prevc[c]}{rk}" for c in FC}, "0.0%", "對照 Val 頁 WACC")
    TR[s] = R
    r += 1
# Q1 校準：m_sw（只用情境 B 的 Q1 估計；三情境 Q1 相同）
rr = TH.max_row + 1
RB = TR["B"]
TH.cell(rr + 1, 1, "Q1 FY27 校準（三情境 Q1 相同）：Cloud and software 分部中「OCI AI 以外」的利潤率 m_sw").font = BOLD
CAL = rr + 2
TH.cell(CAL, 1, "Q1 分部利潤 − 命題 OCI AI 貢獻（估計）"); TH[f"C{CAL}"] = f"=Q!I16-C{RB['contrib']}"
TH.cell(CAL + 1, 1, "Q1 分部營收 − 命題 OCI AI 營收（估計）"); TH[f"C{CAL+1}"] = f"=Q!I11-C{RB['rev']}"
TH.cell(CAL + 2, 1, "m_sw（In 頁引用）"); TH[f"C{CAL+2}"] = f"=C{CAL}/C{CAL+1}"; TH[f"C{CAL+2}"].number_format = "0.0%"
TH.cell(CAL + 3, 1, "對照：FY2024 雲端與軟體分部利潤率（AI 規模化前）"); TH[f"C{CAL+3}"] = "=Q!B45"; TH[f"C{CAL+3}"].number_format = "0.0%"
TH.cell(CAL + 4, 1, "對照：Q1 OCI 非 AI 殘差占 OCI 比例（v0.3.1 區間 12–30%）"); TH[f"C{CAL+4}"] = f"=C{RB['oth']}/q1_oci"; TH[f"C{CAL+4}"].number_format = "0.0%"
for x in range(CAL, CAL + 2): TH[f"C{x}"].number_format = "#,##0.000"
M_SW_CELL = f"Thesis!C{CAL+2}"
for _r in range(4, IN.max_row + 1):
    if IN.cell(_r, 1).value == "m_sw" and "m_sw" not in OVR: IN.cell(_r, 5).value = "=" + M_SW_CELL
# RPO 轉換檢查
RP = CAL + 6
TH.cell(RP, 1, "RPO 轉換檢查（命題的對照列；不作校準）").font = BOLD
hdr(TH, RP + 1, ["項目", "單位", "A", "B", "C", "說明"])
TH.cell(RP + 2, 1, "RPO 13–36 個月部分的年化認列（約 FY2028Q2–FY2030Q1）"); TH.cell(RP + 2, 2, "US$B／年")
for j, s in enumerate(SC): TH.cell(RP + 2, 3 + j, "=rpo*0.37/2").number_format = "#,##0.0"
TH.cell(RP + 2, 6, "RPO 含所有合約（OCI、SaaS、支援）；OpenAI 約一半（二手）")
TH.cell(RP + 3, 1, "模型 FY2029 OCI AI 營收 ÷ RPO 年化認列"); TH.cell(RP + 3, 2, "x")
for j, s in enumerate(SC): TH.cell(RP + 3, 3 + j, f"=G{TR[s]['rev']}/(rpo*0.37/2)").number_format = "0.00"
TH.cell(RP + 4, 1, "模型 FY2029 OCI AI 營收 ÷ OpenAI 合約年額 60"); TH.cell(RP + 4, 2, "x")
for j, s in enumerate(SC): TH.cell(RP + 4, 3 + j, f"=G{TR[s]['rev']}/oai_annual").number_format = "0.00"
TH.cell(RP + 5, 1, "模型 FY2029 OCI（含非 AI）÷ 公司 OCI 路徑 FY2029（114）"); TH.cell(RP + 5, 2, "x")
for j, s in enumerate(SC): TH.cell(RP + 5, 3 + j, f"=G{TR[s]['oci']}/go2029").number_format = "0.00"
RPO_ROWS = (RP + 2, RP + 5)
# 命題答案摘要（核心問題：FY2029 OCI 營收與營業利益率 vs 資本支出與租約、資金缺口）
SM = RP + 7
TH.cell(SM, 1, "命題答案摘要（FY2029；核心問題）").font = BOLD
hdr(TH, SM + 1, ["項目", "單位", "情境 A", "情境 B", "情境 C", "說明"])
SUMQ = [("OCI 營收 FY2029（含非 AI）", "US$B", lambda s: f"=G{TR[s]['oci']}", "#,##0.0", "公司路徑 114"),
        ("其中 OCI AI 算力營收 FY2029", "US$B", lambda s: f"=G{TR[s]['rev']}", "#,##0.0", ""),
        ("OCI AI 營業利益率 FY2029（含 AI 折舊、租金）", "%", lambda s: f"=G{TR[s]['om']}", "0.0%", ""),
        ("OCI（含非 AI）營業利益率 FY2029", "%", lambda s: f"=G{TR[s]['ocim']}", "0.0%", ""),
        ("AI 資本支出累計 FY2027–FY2029", "US$B", lambda s: f"=E{TR[s]['cx']}+F{TR[s]['cx']}+G{TR[s]['cx']}", "#,##0.0", ""),
        ("AI 機房租金 FY2029", "US$B", lambda s: f"=G{TR[s]['rent']}", "#,##0.0", "未起租 288B 起租"),
        ("AI 增量 ROIC FY2029", "%", lambda s: f"=G{TR[s]['roic']}", "0.0%", "對照 WACC（Val 頁）"),
        ("收支關卡：FY2027–FY2031 自由現金流扣股利", "US$B", lambda s: f"=CF!{'CDE'[SC.index(s)]}4", "#,##0.0", "CF 頁")]
for i, (lab, unit, f, fmt, note) in enumerate(SUMQ):
    rr_ = SM + 2 + i
    TH.cell(rr_, 1, lab); TH.cell(rr_, 2, unit); TH.cell(rr_, 6, note)
    for j, s in enumerate(SC):
        TH.cell(rr_, 3 + j, f(s)).number_format = fmt
SUM_ROWS = (SM + 1, SM + 1 + len(SUMQ))
TH.column_dimensions["A"].width = 66; TH.column_dimensions["B"].width = 14
for c in "CDEFGHIJ": TH.column_dimensions[c].width = 12
TH.column_dimensions["K"].width = 70
TH.freeze_panes = "C4"


# =====================================================================================
# Seg：營收分部層
# =====================================================================================
SG = wb.create_sheet("Seg")
SG["A1"] = "Seg — 營收分部層：營收線（公司新聞稿）→ 官方報告分部（10-K 附註 13：Cloud and software、Hardware、Services）→ 合併營收。FY2026 為實績（Q 頁）；FY2027＝Q1 實績＋Q2–Q4 模型；延伸到 FY2032"
SG["A2"] = "驅動：OCI AI 算力＝Thesis 頁命題模組；OCI 其他＝成長路徑；SaaS、授權、支援、硬體、服務＝FY2027 為 Q1 實績＋去年 Q2–Q4 ×（1＋近四季年增率），其後線性收斂到 In 頁長期年增（三情境相同）。核對列：子列合計＝母列、分部合計＝合併營收"
YH = ["項目", "單位"] + [f"FY{y}" + ("A" if y == 2026 else "") for y in YRS] + ["說明"]
hdr(SG, 3, YH)
r = 4
SG.cell(r, 1, "共同（三情境相同）：OCI 以外的營收線").font = BOLD; r += 1
SGC = {}
OTH = (("saas", "Cloud Applications（SaaS）"), ("lic", "Software license"), ("sup", "Software support"), ("hw", "Hardware"), ("svc", "Services"))
for k, lab in OTH:
    rv, rg = r, r + 1
    SG.cell(rv, 1, lab); SG.cell(rv, 2, "US$B"); SG.cell(rg, 1, f"  {lab} 年增"); SG.cell(rg, 2, "%")
    SG[f"C{rv}"] = f"=a26_{k}"; SG[f"D{rv}"] = f"=q1_{k}+(a26_{k}-q0_{k})*(1+g0_{k})"
    SG[f"D{rg}"] = f"=D{rv}/C{rv}-1"
    for i, y in enumerate(range(2028, 2033), 1):
        c = yc(y); p = yc(y - 1)
        SG[f"{c}{rg}"] = f"=g0_{k}+(gLT_{k}-g0_{k})*{i}/5"
        SG[f"{c}{rv}"] = f"={p}{rv}*(1+{c}{rg})"
    for c in "CDEFGHI":
        SG[f"{c}{rv}"].number_format = "#,##0.000"; SG[f"{c}{rg}"].number_format = "0.0%"
    SG.cell(rv, 10, "FY2027＝Q1 實績＋（FY2026 − FY26Q1）×（1＋近四季年增率 Q 頁）；FY2028 起年增率線性收斂到 In 頁 gLT_" + k)
    SGC[k] = rv; r += 2
r += 1
SR = {}
for s in SC:
    SG.cell(r, 1, SCN[s]).font = BOLD; r += 1
    R = {}; T = TR[s]
    def srow(k, lab, vals, fmt="#,##0.000", note=None):
        global r
        R[k] = r; SG.cell(r, 1, lab); SG.cell(r, 2, "US$B" if fmt != "0.0%" else "%")
        for c, v in vals.items():
            SG[f"{c}{r}"] = v; SG[f"{c}{r}"].number_format = fmt
        if note: SG.cell(r, 10, note)
        r += 1
    srow("oci", "Cloud Infrastructure（OCI）", {"C": "=a26_oci", **{yc(y): f"={yc(y)}{r+1}+{yc(y)}{r+2}" for y in FY}}, note="FY2026 實績（不拆 AI／非 AI：公司未揭露）")
    srow("ai", "　OCI AI 算力（命題子列）", {yc(y): f"=Thesis!{tc(y)}{T['rev']}" for y in FY}, note="Thesis 頁命題模組")
    srow("oth", "　OCI 其他（非 AI）（子列）", {yc(y): f"=Thesis!{tc(y)}{T['oth']}" for y in FY})
    for k, lab in OTH:
        srow(k, lab, {c: f"={c}{SGC[k]}" for c in "CDEFGHI"})
    srow("cs", "分部：Cloud and software", {c: f"={c}{R['oci']}+{c}{R['saas']}+{c}{R['lic']}+{c}{R['sup']}" for c in "CDEFGHI"}, note="10-K 官方分部（雲端與軟體）")
    srow("shw", "分部：Hardware", {c: f"={c}{R['hw']}" for c in "CDEFGHI"})
    srow("ssvc", "分部：Services", {c: f"={c}{R['svc']}" for c in "CDEFGHI"})
    srow("tot", "合併營收", {c: f"={c}{R['cs']}+{c}{R['shw']}+{c}{R['ssvc']}" for c in "CDEFGHI"})
    srow("g", "  合併營收年增", {c: f"={c}{R['tot']}/{p}{R['tot']}-1" for c, p in zip("DEFGHI", "CDEFGH")}, "0.0%")
    srow("chk1", "檢查：OCI 子列合計 − OCI（應為 0）", {yc(y): f"={yc(y)}{R['ai']}+{yc(y)}{R['oth']}-{yc(y)}{R['oci']}" for y in FY}, "0.000")
    srow("chk2", "檢查：營收線合計 − 合併營收（應為 0）", {c: f"={c}{R['oci']}+{c}{R['saas']}+{c}{R['lic']}+{c}{R['sup']}+{c}{R['hw']}+{c}{R['svc']}-{c}{R['tot']}" for c in "CDEFGHI"}, "0.000")
    srow("chk3", "檢查：FY2026 合併營收 − 公司揭露（應為 0）", {"C": f"=C{R['tot']}-a26_rev"}, "0.000")
    SR[s] = R; r += 1
SG.column_dimensions["A"].width = 46; SG.column_dimensions["B"].width = 8
for c in "CDEFGHI": SG.column_dimensions[c].width = 12
SG.column_dimensions["J"].width = 80
SG.freeze_panes = "C4"
# Thesis 的 FY2027 非 AI 資本支出需要合併營收：用情境 B 的 D 欄（三情境 FY2027 營收僅 OCI AI 不同 → 改用各自情境）
for s in SC:
    T = TR[s]
    TH[f"E{T['cxg']}"] = f"=capex_g27-nonai_cx*(Seg!D{SR[s]['tot']}-E{T['rev']})"


# =====================================================================================
# PL：分部營收 → 分部利潤 → 合併營業利益 → 淨利 → EPS
# =====================================================================================
PL = wb.create_sheet("PL")
PL["A1"] = "PL — 三情境損益（十億美元；EPS 美元）。分部營收（Seg 頁）→ 分部利潤（Cloud and software＝OCI AI 貢獻〔Thesis 頁，含 AI 折舊與租金〕＋其餘營收 × m_sw；硬體、服務＝營收 × FY2026 利潤率）→ 減未分攤費用＝GAAP 營業利益 → 利息、稅 → 淨利 → EPS（GAAP）"
PL["A2"] = "FY2026 為實績；FY2027＝Q1 實績＋Q2–Q4 模型。持股評價損益：預測期設 0。EPS 為 GAAP 稀釋口徑（共識 GAAP EPS 對照見 Recon）"
hdr(PL, 3, YH)
r = 4
PL.cell(r, 1, "分部核對（三情境合計絕對差；每年度應為 0）").font = BOLD; r += 1
CHK_SEG = r; PL.cell(r, 1, "檢查：分部合計 − 合併營收"); r += 1
CHK_SUB = r; PL.cell(r, 1, "檢查：子列合計 − 母列（OCI 兩子列；Cloud and software 四營收線）"); r += 1
CHK_SUM = r; PL.cell(r, 1, "檢查：營收線合計 − 合併營收"); r += 2
PR_ = {}
for s in SC:
    PL.cell(r, 1, SCN[s]).font = BOLD; r += 1
    R = {}; T = TR[s]; G = SR[s]
    def prow(k, lab, vals, fmt="#,##0.000", note=None, unit="US$B"):
        global r
        R[k] = r; PL.cell(r, 1, lab); PL.cell(r, 2, unit)
        for c, v in vals.items():
            if v is None: continue
            PL[f"{c}{r}"] = v; PL[f"{c}{r}"].number_format = fmt
        if note: PL.cell(r, 10, note)
        r += 1
    for k, lab in (("oci", "Cloud Infrastructure（OCI）"), ("ai", "　OCI AI 算力（命題子列）"), ("oth", "　OCI 其他（非 AI）（子列）"),
                   ("saas", "Cloud Applications（SaaS）"), ("lic", "Software license"), ("sup", "Software support"), ("hw", "Hardware"), ("svc", "Services")):
        prow(k, lab, {c: (f"=Seg!{c}{G[k]}" if not (k in ("ai", "oth") and c == "C") else None) for c in "CDEFGHI"})
    prow("cs", "分部：Cloud and software", {c: f"={c}{R['oci']}+{c}{R['saas']}+{c}{R['lic']}+{c}{R['sup']}" for c in "CDEFGHI"})
    prow("shw", "分部：Hardware", {c: f"={c}{R['hw']}" for c in "CDEFGHI"})
    prow("ssvc", "分部：Services", {c: f"={c}{R['svc']}" for c in "CDEFGHI"})
    prow("rev", "合併營收", {c: f"=Seg!{c}{G['tot']}" for c in "CDEFGHI"})
    prow("revg", "  合併營收年增", {c: f"={c}{R['rev']}/{p}{R['rev']}-1" for c, p in zip("DEFGHI", "CDEFGH")}, "0.0%", unit="%")
    msw = "(m_sw+m_adj)"
    v = {"C": "=Q!L16", "D": f"=Q!I16+Thesis!D{T['contrib']}+{msw}*((D{R['cs']}-Q!I11)-Thesis!D{T['rev']})"}
    for y in range(2028, 2033): v[yc(y)] = f"=Thesis!{tc(y)}{T['contrib']}+{msw}*({yc(y)}{R['cs']}-Thesis!{tc(y)}{T['rev']})"
    prow("mcs", "分部利潤：Cloud and software", v, note="OCI AI 貢獻（Thesis，含 AI 折舊與租金）＋（分部營收 − OCI AI 營收）× m_sw")
    prow("mcs_ai", "  其中：OCI AI 貢獻（命題）", {yc(y): f"=Thesis!{tc(y)}{T['contrib']}" for y in FY})
    prow("mhw", "分部利潤：Hardware", {"C": "=Q!L17", "D": f"=Q!I17+m_hw*(D{R['shw']}-Q!I12)", **{yc(y): f"=m_hw*{yc(y)}{R['shw']}" for y in range(2028, 2033)}})
    prow("msvc", "分部利潤：Services", {"C": "=Q!L18", "D": f"=Q!I18+m_svc*(D{R['ssvc']}-Q!I13)", **{yc(y): f"=m_svc*{yc(y)}{R['ssvc']}" for y in range(2028, 2033)}})
    prow("mtot", "分部利潤合計", {c: f"={c}{R['mcs']}+{c}{R['mhw']}+{c}{R['msvc']}" for c in "CDEFGHI"})
    prow("mcsr", "  Cloud and software 分部利潤率", {c: f"={c}{R['mcs']}/{c}{R['cs']}" for c in "CDEFGHI"}, "0.0%", unit="%")
    prow("un", "未分攤費用（研發、管理、攤銷、SBC 等）", {"C": "=Q!L20", "D": "=Q!I20*(1+3*un_rr)", **{yc(y): f"={yc(y-1)}{r}*(1+un_g)" for y in range(2028, 2033)}}, note="FY2027＝Q1 實績 ×（1＋3 × un_rr）；其後年增 un_g")
    prow("oi", "營業利益（GAAP）", {c: f"={c}{R['mtot']}-{c}{R['un']}" for c in "CDEFGHI"})
    prow("oim", "  營業利益率", {c: f"={c}{R['oi']}/{c}{R['rev']}" for c in "CDEFGHI"}, "0.0%", unit="%")
    prow("ociom", "  OCI AI 營業利益率（命題）", {yc(y): f"=Thesis!{tc(y)}{T['om']}" for y in FY}, "0.0%", unit="%")
    prow("int", "利息費用", {"C": "=Q!L26", **{yc(y): f"=CF!{yc(y)}{{int}}" for y in FY}}, note="CF 頁：既有借款 × r_old＋期初新債 × r_new＋高息債 × 12%")
    prow("ii", "利息收入", {"C": None, **{yc(y): f"=CF!{yc(y)}{{ii}}" for y in FY}}, note="FY2026 含在其他收益（未拆）")
    prow("pti", "稅前利益", {"C": f"=Q!L27/(1-Q!L26*0+0.1262)" if False else None, **{yc(y): f"={yc(y)}{R['oi']}-{yc(y)}{R['int']}+{yc(y)}{R['ii']}" for y in FY}}, note="其他收益與持股評價損益：預測期設 0")
    prow("tax", "所得稅", {yc(y): f"={yc(y)}{R['pti']}*etr" for y in FY})
    prow("ni", "淨利", {"C": "=Q!L27", **{yc(y): f"={yc(y)}{R['pti']}-{yc(y)}{R['tax']}" for y in FY}})
    prow("pref", "特別股股利", {"C": "=Q!L48", "D": "=pref_q1+pref27", "E": "=pref28", "F": "=pref29", "G": "=0", "H": "=0", "I": "=0"}, note="FY27 Q1 實績 0.081＋Q2–Q4；2029-01-15 強制轉換")
    prow("nic", "歸屬普通股淨利", {c: f"={c}{R['ni']}-{c}{R['pref']}" for c in "CDEFGHI"})
    prow("sh", "稀釋股數", {"C": "=Q!L49", **{yc(y): f"=CF!{yc(y)}{{shd}}" for y in FY}}, unit="十億股")
    prow("eps", "EPS（GAAP 稀釋）", {"C": "=Q!L50", **{yc(y): f"={yc(y)}{R['nic']}/{yc(y)}{R['sh']}" for y in FY}}, "0.00", unit="US$")
    prow("epsg", "  EPS 年增", {c: f"={c}{R['eps']}/{p}{R['eps']}-1" for c, p in zip("DEFGHI", "CDEFGH")}, "0.0%", unit="%")
    v = {"C": "=Q!L24", "D": f"=Q!I24+Thesis!D{T['dep']}+nonai_dep_q*3+amort*0.75"}
    for y in range(2028, 2033): v[yc(y)] = f"=Thesis!{tc(y)}{T['dep']}+nonai_dep_q*({yc(y)}{R['rev']}-Thesis!{tc(y)}{T['rev']})/(q1_rev-Thesis!C{T['rev']})+amort"
    prow("da", "折舊與攤銷（D&A）", v, note="AI 折舊（Thesis）＋非 AI 折舊（FY2024 水準每季 0.782，隨非 AI 營收相對 Q1 等比例）＋無形資產攤銷")
    prow("sbc", "股份基礎給付（SBC）", {"C": "=Q!L25", "D": "=sbc_q1*(1+3*un_rr)", **{yc(y): f"={yc(y-1)}{r}*(1+un_g)" for y in range(2028, 2033)}})
    prow("ebitda", "EBITDA（營業利益＋D&A）", {c: f"={c}{R['oi']}+{c}{R['da']}" for c in "CDEFGHI"})
    prow("aebitda", "調整後 EBITDA（＋SBC；共識與同業口徑）", {c: f"={c}{R['ebitda']}+{c}{R['sbc']}" for c in "CDEFGHI"})
    prow("chk", "檢查：FY2026 營業利益 − 公司揭露 20.606（應為 0）", {"C": f"=C{R['oi']}-Q!L21"}, "0.000")
    PR_[s] = R; r += 1
for c in "CDEFGHI":
    PL[f"{c}{CHK_SEG}"] = "=" + "+".join(f"ABS({c}{PR_[s]['cs']}+{c}{PR_[s]['shw']}+{c}{PR_[s]['ssvc']}-{c}{PR_[s]['rev']})" for s in SC)
    PL[f"{c}{CHK_SUM}"] = "=" + "+".join(f"ABS({c}{PR_[s]['oci']}+{c}{PR_[s]['saas']}+{c}{PR_[s]['lic']}+{c}{PR_[s]['sup']}+{c}{PR_[s]['hw']}+{c}{PR_[s]['svc']}-{c}{PR_[s]['rev']})" for s in SC)
    sub = "+".join(f"ABS({c}{PR_[s]['oci']}+{c}{PR_[s]['saas']}+{c}{PR_[s]['lic']}+{c}{PR_[s]['sup']}-{c}{PR_[s]['cs']})" for s in SC)
    if c != "C":
        sub += "+" + "+".join(f"ABS({c}{PR_[s]['ai']}+{c}{PR_[s]['oth']}-{c}{PR_[s]['oci']})" for s in SC)
    PL[f"{c}{CHK_SUB}"] = "=" + sub
    for x in (CHK_SEG, CHK_SUB, CHK_SUM): PL[f"{c}{x}"].number_format = "0.000"
PL.column_dimensions["A"].width = 48; PL.column_dimensions["B"].width = 8
for c in "CDEFGHI": PL.column_dimensions[c].width = 12
PL.column_dimensions["J"].width = 80
PL.freeze_panes = "C4"


# =====================================================================================
# CF：現金流與融資瀑布（收支模組）
# =====================================================================================
CF = wb.create_sheet("CF")
CF["A1"] = "CF — 現金流與融資瀑布（收支模組）。營運現金＝淨利＋D&A＋SBC＋客戶預付流入 − 預付認列（非現金營收）；資本支出＝AI（Thesis）＋非 AI；自由現金流扣股利後的缺口依序由現金（保留最低 10）→ 投資級債（租賃調整後槓桿 ≤ 4.5×）→ 新股（每年 ≤ 市值 5%）→ 高息債 補足"
CF["A2"] = "FY2027＝Q1 實績＋Q2–Q4 模型（瀑布只計 Q2–Q4，起點為 Q1 末現金）。利息以期初餘額計（不循環）。租賃負債＝在帳 43.8＋已起租的未起租租約現值（17 年、5.7%）"
hdr(CF, 3, ["項目", "單位", "情境 A", "情境 B", "情境 C", "說明"])
GATE = 4
CF.cell(GATE, 1, "收支關卡：自由現金流扣股利（五年合計）"); CF.cell(GATE, 2, "US$B")
CF.cell(GATE + 1, 1, "融資需求（五年合計）：新增投資級債"); CF.cell(GATE + 1, 2, "US$B")
CF.cell(GATE + 2, 1, "融資需求（五年合計）：新股"); CF.cell(GATE + 2, 2, "US$B")
CF.cell(GATE + 3, 1, "融資需求（五年合計）：高息債"); CF.cell(GATE + 3, 2, "US$B")
CF.cell(GATE + 4, 1, "調整後槓桿峰值（FY2027–FY2031）"); CF.cell(GATE + 4, 2, "x")
CF.cell(GATE + 5, 1, "是否需要發新股或高息債（五年）"); CF.cell(GATE + 5, 2, "—")
CF.cell(GATE + 6, 1, "累計自由現金流（五年合計）"); CF.cell(GATE + 6, 2, "US$B")
CF.cell(GATE + 7, 1, "檢查：收支關卡＝首年起五年自由現金流扣股利（逐欄加總，應為 0）"); CF.cell(GATE + 7, 2, "US$B")
CF.cell(GATE, 6, "FY2027–FY2031 自由現金流 − 普通股與特別股股利 的合計（負值＝需外部融資；地圖收支關卡讀這一列）")
r = 13
hdr(CF, r, YH); r += 1
CR_ = {}
for s in SC:
    CF.cell(r, 1, SCN[s]).font = BOLD; r += 1
    R = {}; T = TR[s]; P = PR_[s]
    def crow(k, lab, vals, fmt="#,##0.000", note=None, unit="US$B"):
        global r
        R[k] = r; CF.cell(r, 1, lab); CF.cell(r, 2, unit)
        for c, v in vals.items():
            if v is None: continue
            CF[f"{c}{r}"] = v; CF[f"{c}{r}"].number_format = fmt
        if note: CF.cell(r, 10, note)
        r += 1
    rows_spec = ["ni", "da", "sbc", "ppin", "pprec", "cfo", "cxai", "cxna", "cx", "fcf", "div", "fcfd", "L", "cash0", "flow", "need", "ebr", "lease", "cap", "ig", "eq", "hy", "cash1", "igc", "hyc", "debt", "lev", "sh0", "neweq", "sh1", "shd", "int", "ii"]
    base = r
    RR = {k: base + i for i, k in enumerate(rows_spec)}
    R.update(RR)
    def F(k): return RR[k]
    vals = {}
    L = {"D": 0.75, **{yc(y): 1 for y in range(2028, 2033)}}
    # 期間長度
    CF.cell(F("L"), 1, "瀑布期間長度（FY2027 只計 Q2–Q4）"); CF.cell(F("L"), 2, "年")
    for c, v in L.items(): CF[f"{c}{F('L')}"] = v
    crow_ = lambda k, lab, vals, fmt="#,##0.000", note=None, unit="US$B": (CF.cell(F(k), 1, lab), CF.cell(F(k), 2, unit), [setattr(CF[f"{c}{F(k)}"], "value", v) or setattr(CF[f"{c}{F(k)}"], "number_format", fmt) for c, v in vals.items() if v is not None], CF.cell(F(k), 10, note) if note else None)
    crow_("ni", "淨利", {c: f"=PL!{c}{P['ni']}" for c in "CDEFGHI"})
    crow_("da", "折舊與攤銷", {c: f"=PL!{c}{P['da']}" for c in "CDEFGHI"})
    crow_("sbc", "SBC（非現金）", {c: f"=PL!{c}{P['sbc']}" for c in "CDEFGHI"})
    crow_("ppin", "客戶預付流入（AI 資本支出 × 覆蓋比）", {"D": f"=prepay_cov*Thesis!D{T['cx']}", **{yc(y): f"=prepay_cov*Thesis!{tc(y)}{T['cx']}" for y in range(2028, 2033)}}, note="FY2027 只計 Q2–Q4（Q1 實績 11.363 已在營運現金流）")
    v = {"D": "=prepay_open/prepay_yrs*0.75"}
    for i, y in enumerate(range(2028, 2033)):
        c = yc(y)
        prev = "+".join(f"{x}{F('ppin')}" for x in "DEFGHI"[: "DEFGHI".index(c)])
        v[c] = f"=prepay_open/prepay_yrs*{1 if y < 2032 else 0.25}+({prev})/prepay_yrs"
    crow_("pprec", "預付認列（非現金營收，自營運現金扣除）", v, note="期初餘額與新流入皆 5 年直線（次年起）")
    crow_("cfo", "營運現金流", {"C": "=Q!L28", "D": f"=cfo_q1+(D{F('ni')}-Q!I27)+(D{F('da')}-Q!I24)+(D{F('sbc')}-sbc_q1)+D{F('ppin')}-D{F('pprec')}",
                               **{yc(y): f"={yc(y)}{F('ni')}+{yc(y)}{F('da')}+{yc(y)}{F('sbc')}+{yc(y)}{F('ppin')}-{yc(y)}{F('pprec')}" for y in range(2028, 2033)}},
         note="FY2027＝Q1 實績 23.103＋Q2–Q4；營運資金變動除預付外不建模")
    crow_("cxai", "資本支出：AI（Thesis）", {yc(y): f"=Thesis!{tc(y)}{T['cx']}" for y in FY})
    crow_("cxna", "資本支出：非 AI", {"D": f"=capex_g27-D{F('cxai')}", **{yc(y): f"=nonai_cx*(PL!{yc(y)}{P['rev']}-Thesis!{tc(y)}{T['rev']})" for y in range(2028, 2033)}})
    crow_("cx", "資本支出合計", {"C": "=Q!L29", **{yc(y): f"={yc(y)}{F('cxai')}+{yc(y)}{F('cxna')}" for y in FY}}, note="FY2027＝公司指引中點 92.5（含 Q1 實績 28.499）")
    crow_("fcf", "自由現金流", {c: f"={c}{F('cfo')}-{c}{F('cx')}" for c in "CDEFGHI"})
    crow_("div", "股利（普通股＋特別股）", {"C": "=Q!L31", "D": f"=div_q1+dps_q*3*D{F('sh0')}+pref27", **{yc(y): f"=dps_q*4*{yc(y)}{F('sh0')}+PL!{yc(y)}{P['pref']}" for y in range(2028, 2033)}})
    crow_("fcfd", "自由現金流扣股利", {c: f"={c}{F('fcf')}-{c}{F('div')}" for c in "CDEFGHI"})
    # 瀑布
    crow_("cash0", "期初現金（FY2027＝Q1 末）", {"D": "=cash_q1", **{yc(y): f"={yc(y-1)}{F('cash1')}" for y in range(2028, 2033)}})
    crow_("flow", "融資前現金流（FY2027 只計 Q2–Q4）", {"D": f"=D{F('fcfd')}-(cfo_q1-capex_q1-div_q1)", **{yc(y): f"={yc(y)}{F('fcfd')}" for y in range(2028, 2033)}})
    crow_("need", "外部融資需求（扣除可動用現金後）", {c: f"=MAX(0,-({c}{F('flow')}+{c}{F('cash0')}-min_cash))" for c in "DEFGHI"})
    crow_("ebr", "EBITDA＋租金（槓桿分母）", {c: f"=PL!{c}{P['ebitda']}+Thesis!{tc(int('20'+str(26+'CDEFGHI'.index(c))))}{T['rent']}*0+rent_on" for c in "DEFGHI"}, note="EBITDA 已扣 AI 租金；加回 AI 租金與在帳租金")
    for c in "DEFGHI":
        y = 2026 + "CDEFGHI".index(c)
        CF[f"{c}{F('ebr')}"] = f"=PL!{c}{P['ebitda']}+Thesis!{tc(y)}{T['rent']}+rent_on"
    full = "lease_unc/lease_term"
    frac = {"D": "6/11*0.5", "E": "7/11", "F": "1", "G": "1", "H": "1", "I": "1"}
    for c in "DEFGHI":
        y = 2026 + "CDEFGHI".index(c)
        CF[f"{c}{F('lease')}"] = f"=lease_on+Thesis!{tc(y)}{T['rent']}/(lease_unc/lease_term)*{full}*(1-(1+lease_r)^-lease_term)/lease_r" if False else f"=lease_on+MIN(1,{ {'D':'3/lease_q','E':'7/lease_q','F':'11/lease_q','G':'11/lease_q','H':'11/lease_q','I':'11/lease_q'}[c] })*{full}*(1-(1+lease_r)^-lease_term)/lease_r"
        CF[f"{c}{F('lease')}"].number_format = "#,##0.0"
    CF.cell(F("lease"), 1, "租賃負債（在帳＋已起租的未起租租約現值，期末）"); CF.cell(F("lease"), 2, "US$B")
    CF.cell(F("lease"), 10, "期末已起租比例：FY27 3/11、FY28 7/11、FY29 起 11/11；在帳部分持平（判斷）")
    crow_("cap", "投資級債可新增額度", {c: f"=MAX(0,lev_max*{c}{F('ebr')}/{c}{F('L')}-{c}{F('lease')}-({'debt_q1' if c == 'D' else chr(ord(c)-1)+str(F('igc'))}+{'0' if c == 'D' else ''}))" for c in "DEFGHI"},
         note="lev_max ×（EBITDA＋租金，年化）− 租賃負債 − 既有借款")
    for c in "DEFGHI":
        prev_debt = "debt_q1" if c == "D" else f"(debt_q1+{chr(ord(c)-1)}{F('igc')})"
        CF[f"{c}{F('cap')}"] = f"=MAX(0,lev_max*{c}{F('ebr')}-{c}{F('lease')}-{prev_debt})"
    crow_("ig", "新增投資級債", {c: f"=MIN({c}{F('need')},{c}{F('cap')})" for c in "DEFGHI"})
    crow_("eq", "新股", {c: f"=MIN({c}{F('need')}-{c}{F('ig')},eq_cap*px*{c}{F('sh0')}*{c}{F('L')})" for c in "DEFGHI"})
    crow_("hy", "高息債", {c: f"={c}{F('need')}-{c}{F('ig')}-{c}{F('eq')}" for c in "DEFGHI"})
    crow_("cash1", "期末現金", {c: f"={c}{F('cash0')}+{c}{F('flow')}+{c}{F('ig')}+{c}{F('eq')}+{c}{F('hy')}" for c in "DEFGHI"})
    crow_("igc", "累計新增投資級債（期末）", {"D": f"=D{F('ig')}", **{yc(y): f"={yc(y-1)}{F('igc')}+{yc(y)}{F('ig')}" for y in range(2028, 2033)}})
    crow_("hyc", "累計高息債（期末）", {"D": f"=D{F('hy')}", **{yc(y): f"={yc(y-1)}{F('hyc')}+{yc(y)}{F('hy')}" for y in range(2028, 2033)}})
    crow_("debt", "借款合計（期末）", {c: f"=debt_q1+{c}{F('igc')}+{c}{F('hyc')}" for c in "DEFGHI"})
    crow_("lev", "調整後槓桿（借款＋租賃負債）÷（EBITDA＋租金）", {c: f"=({c}{F('debt')}+{c}{F('lease')})/{c}{F('ebr')}" for c in "DEFGHI"}, "0.00", unit="x")
    crow_("sh0", "期初普通股（十億股）", {"D": "=sh_out", **{yc(y): f"={yc(y-1)}{F('sh1')}" for y in range(2028, 2033)}}, unit="十億股")
    crow_("neweq", "新發行股數", {c: f"={c}{F('eq')}/(px*(1-eq_disc))" for c in "DEFGHI"}, unit="十億股")
    crow_("sh1", "期末普通股", {c: f"={c}{F('sh0')}+{c}{F('neweq')}" + ("+mcps_sh" if c == "F" else "") for c in "DEFGHI"}, unit="十億股", note="FY2029 加強制轉換特別股 0.031")
    crow_("shd", "稀釋股數（平均＋選擇權）", {"D": f"=(0.75*(D{F('sh0')}+D{F('sh1')})/2+0.25*shd_q1)+opt_dil*0.75", **{yc(y): f"=({yc(y)}{F('sh0')}+{yc(y)}{F('sh1')})/2+opt_dil" for y in range(2028, 2033)}}, unit="十億股", note="FY2027 混合 Q1 實績稀釋加權 3.000")
    crow_("int", "利息費用", {"D": f"=int_q1+r_old*debt_q1*0.75", **{yc(y): f"=r_old*debt_q1+r_new*{yc(y-1)}{F('igc')}+r_junk*{yc(y-1)}{F('hyc')}" for y in range(2028, 2033)}}, note="FY2027＝Q1 實績 1.428＋Q2–Q4；新債以期初餘額計息")
    crow_("ii", "利息收入", {"D": f"=ii_q1+r_cash*D{F('cash0')}*0.75", **{yc(y): f"=r_cash*{yc(y)}{F('cash0')}" for y in range(2028, 2033)}})
    r = base + len(rows_spec)
    CR_[s] = R; r += 1
# 回填 PL 的 CF 參照
for s in SC:
    P = PR_[s]; C_ = CR_[s]
    for k, rk in (("int", "int"), ("ii", "ii"), ("sh", "shd")):
        for y in FY:
            c = yc(y)
            PL[f"{c}{P[k]}"] = f"=CF!{c}{C_[rk]}"
# 收支關卡摘要
for j, s in enumerate(SC):
    c = "CDE"[j]; C_ = CR_[s]
    CF[f"{c}{GATE}"] = f"=SUM(D{C_['fcfd']}:H{C_['fcfd']})"
    CF[f"{c}{GATE+1}"] = f"=SUM(D{C_['ig']}:H{C_['ig']})"
    CF[f"{c}{GATE+2}"] = f"=SUM(D{C_['eq']}:H{C_['eq']})"
    CF[f"{c}{GATE+3}"] = f"=SUM(D{C_['hy']}:H{C_['hy']})"
    CF[f"{c}{GATE+4}"] = f"=MAX(D{C_['lev']}:H{C_['lev']})"
    CF[f"{c}{GATE+5}"] = f'=IF({c}{GATE+2}+{c}{GATE+3}>0.05,"需要","不需要")'
    CF[f"{c}{GATE+6}"] = f"=SUM(D{C_['fcf']}:H{C_['fcf']})"
    CF[f"{c}{GATE+7}"] = f"={c}{GATE}-(" + "+".join(f"{x}{C_['fcfd']}" for x in "DEFGH") + ")"
    for x in range(GATE, GATE + 8):
        CF[f"{c}{x}"].number_format = "0.00" if x == GATE + 4 else "#,##0.0"
CF.column_dimensions["A"].width = 50; CF.column_dimensions["B"].width = 8
for c in "CDEFGHI": CF.column_dimensions[c].width = 12
CF.column_dimensions["J"].width = 70; CF.column_dimensions["F"].width = 14
CF.freeze_panes = "C4"


# =====================================================================================
# Val：DCF 50%＋分部加總（SOTP）50%
# =====================================================================================
V = wb.create_sheet("Val")
V["A1"] = "Val — 目標價（約 12 個月後，2027-10）＝DCF（UFCF、WACC）50%＋分部加總 50%（OCI AI 以 AI 模組自身 DCF 計，同業倍數版列對照；SaaS 與 OCI 非 AI × 軟體同業倍數；授權、支援、硬體、服務 × 成熟軟體倍數）；情境權重 A 25%／B 50%／C 25%（判斷預設，Andy 可改）"
V["A2"] = "評價日 2026-08-31（FY27Q1 資產負債表）。SBC 口徑（F2）：DCF 以 GAAP 營業利益計（SBC 視為現金成本、股數不另計稀釋）；分部加總的 EBITDA 加回 SBC（與同業倍數同口徑）。淨借款含融資租賃 9.185；營業租賃不列（UFCF 已扣租金）。永續期：RONIC＝WACC（F5），AI 折舊＝穩態汰換"
hdr(V, 3, ["項目", "單位", "數值", "", "", "說明"])
def vrow(row, lab, unit, val, fmt="#,##0.00", note=None):
    V.cell(row, 1, lab); V.cell(row, 2, unit)
    if val is not None:
        V.cell(row, 3, val).number_format = fmt
    if note: V.cell(row, 6, note)
vrow(4, "現價（2026-10-26 收盤）", "US$", "=px", note="In 頁 px（stockanalysis.com）")
vrow(5, "股權成本 ke（CAPM：rf＋β × ERP；β 為 Blume 調整）", "%", "=rf+beta*erp+wacc_adj", "0.00%", "wacc_adj 為敏感度調整（基準 0）")
vrow(6, "稅後債務成本（約 10 年期）", "%", "=kd*(1-etr)", "0.00%")
vrow(7, "股權市值（現價 × 流通股）", "US$B", "=px*sh_out", "#,##0.0")
vrow(8, "負債（借款本金＋融資租賃，FY27Q1 末）", "US$B", "=debt_q1+finlease_q1", "#,##0.0")
vrow(9, "WACC", "%", "=(C7*C5+C8*C6)/(C7+C8)", "0.00%", "修正一（F4）：Blume β、約 10 年期 kd、負債含融資租賃（強制轉換特別股視為股權）")
vrow(10, "淨借款（借款＋融資租賃 − 現金，FY27Q1 末）", "US$B", "=debt_q1+finlease_q1-cash_q1", "#,##0.0", "修正一：含融資租賃（F6）")
vrow(11, "完全稀釋股數（流通＋選擇權＋強制轉換特別股）", "十億股", "=sh_out+opt_dil+mcps_sh", "0.000")
vrow(12, "目標價時點距評價日", "年", "=(tp_date-val_date)/365.25", "0.000", "In 頁 val_date → tp_date")
vrow(13, "FY2028 年底距目標價時點", "年", "=(fy28_end-tp_date)/365.25", "0.000", "分部加總錨定日 → 目標價時點")
vrow(14, "永續期新投資報酬率 RONIC", "%", "=C9+ronic_adj", "0.00%", "RONIC＝WACC：永續成長不創造價值（F5）；Sens：WACC＋5pt")
for nm, ref in (("ke", "Val!$C$5"), ("wacc", "Val!$C$9"), ("netdebt0", "Val!$C$10"), ("shfd", "Val!$C$11"), ("t_tp", "Val!$C$12"), ("t_28", "Val!$C$13"), ("ronic", "Val!$C$14")):
    dn(nm, ref)
r = 16
V.cell(r, 1, "0.1 DCF（無槓桿自由現金流，WACC；SBC 為現金成本）").font = BOLD; r += 1
hdr(V, r, ["項目", "折現年數 t"] + [f"情境 {s}" for s in SC] + ["說明"]); r += 1
TT = {2027: 0.375, 2028: 1.25, 2029: 2.25, 2030: 3.25, 2031: 4.25, 2032: 5.25}
DC = {}
for y in FY:
    V.cell(r, 1, f"  UFCF FY{y}" + ("（Q2–Q4）" if y == 2027 else "")); V.cell(r, 2, TT[y]).number_format = "0.000"
    for j, s in enumerate(SC):
        P = PR_[s]; T = TR[s]; C_ = CR_[s]; c = yc(y)
        if y == 2027:
            f = f"=(PL!D{P['oi']}-Q!I21)*(1-etr)+(PL!D{P['da']}-Q!I24)-(CF!D{C_['cx']}-capex_q1)+CF!D{C_['ppin']}-CF!D{C_['pprec']}"
        else:
            f = f"=PL!{c}{P['oi']}*(1-etr)+PL!{c}{P['da']}-CF!{c}{C_['cx']}+CF!{c}{C_['ppin']}-CF!{c}{C_['pprec']}"
        V.cell(r, 3 + j, f).number_format = "#,##0.00"
    DC[y] = r; r += 1
V.cell(r, 2, None)
V.cell(r, 1, "  永續期 NOPAT（FY2032 營業利益，AI 折舊改為穩態汰換；稅後）")
for j, s in enumerate(SC):
    P = PR_[s]; T = TR[s]
    V.cell(r, 3 + j, f"=(PL!I{P['oi']}+Thesis!J{T['dep']}-Thesis!J{T['conn']}*cx_vr200/life_ai/1000)*(1-etr)").number_format = "#,##0.00"
V.cell(r, 6, "穩態 AI 折舊＝穩態汰換資本支出＝已連網 MW × VR200 每 MW 資本支出 ÷ 年限（修正一：取代 FY2032 帳面 AI 折舊）"); DC["tvb"] = r; r += 1
V.cell(r, 1, "  永續價值（FY2032 年底，t＝5.75；NOPAT ×(1＋g)×(1 − g ÷ RONIC) ÷ (WACC − g)）"); V.cell(r, 2, 5.75)
for j in range(3):
    cc = get_column_letter(3 + j)
    V.cell(r, 3 + j, f"={cc}{DC['tvb']}*(1+tp_g)*(1-tp_g/ronic)/(wacc-tp_g)").number_format = "#,##0.0"
V.cell(r, 6, "再投資率＝g ÷ RONIC（成長所需資本支出與 g 一致，F5）")
DC["tv"] = r; r += 1
V.cell(r, 1, "企業價值（評價日）")
for j in range(3):
    cc = get_column_letter(3 + j)
    V.cell(r, 3 + j, "=" + "+".join(f"{cc}{DC[y]}/(1+wacc)^$B${DC[y]}" for y in FY) + f"+{cc}{DC['tv']}/(1+wacc)^$B${DC['tv']}").number_format = "#,##0.0"
DC["ev"] = r; r += 1
V.cell(r, 1, "  其中：永續價值現值占比")
for j in range(3):
    cc = get_column_letter(3 + j)
    V.cell(r, 3 + j, f"={cc}{DC['tv']}/(1+wacc)^$B${DC['tv']}/{cc}{DC['ev']}").number_format = "0.0%"
DC["tvsh"] = r; r += 1
V.cell(r, 1, "股權價值（企業價值 − 淨借款〔含融資租賃〕）")
for j in range(3):
    cc = get_column_letter(3 + j)
    V.cell(r, 3 + j, f"={cc}{DC['ev']}-netdebt0").number_format = "#,##0.0"
DC["eqv"] = r; r += 1
V.cell(r, 1, "DCF 每股價值（評價日；小於 0 取 0）")
for j in range(3):
    cc = get_column_letter(3 + j)
    V.cell(r, 3 + j, f"=MAX(0,{cc}{DC['eqv']}/shfd)").number_format = "#,##0.00"
DC["ps"] = r; r += 1
V.cell(r, 1, "DCF 情境價值推到 12 個月後（×(1＋ke)^t − 期間股利）")
for j in range(3):
    cc = get_column_letter(3 + j)
    V.cell(r, 3 + j, f"=MAX(0,{cc}{DC['ps']}*(1+ke)^t_tp-dps_q*4*t_tp)").number_format = "#,##0.00"
    V.cell(r, 3 + j).fill = KEY
DC["v12"] = r; r += 2
# SOTP
V.cell(r, 1, "0.2 分部加總（FY2029 EBITDA × 前瞻倍數；OCI AI 以 AI 模組 DCF）").font = BOLD; r += 1
hdr(V, r, ["同業", "代號", "EV（US$B）", "前瞻 EV/EBITDA", "日期", "來源與標記"]); r += 1
V.cell(r, 1, "OCI AI 同業（neocloud）：EV ÷ 2027 年共識 EBITDA（修正一起只作對照列）").font = GREY; r += 1
for nm, tk, ev, eb, dt, src in (("CoreWeave", "CRWV", "ev_crwv", "eb_crwv", "2026-10-21", "In 頁 ev_crwv、eb_crwv（MarketScreener，Interested-party）"),
                                ("Nebius", "NBIS", "ev_nbis", "eb_nbis", "2026-10-21", "In 頁 ev_nbis、eb_nbis（MarketScreener，Interested-party）")):
    V.cell(r, 1, nm); V.cell(r, 2, tk); V.cell(r, 3, f"={ev}").number_format = "#,##0.0"
    V.cell(r, 4, f"={ev}/{eb}").number_format = "0.00"; V.cell(r, 5, dt); V.cell(r, 6, src); r += 1
OCI_PEERS = (r - 2, r - 1)
V.cell(r, 1, "  第三家（IREN、WhiteFiber）：EBITDA 共識未找到——不符合「至少 3 家」規格；本版 OCI AI 主值改用 DCF，倍數只作對照").font = GREY; r += 1
V.cell(r, 1, "其餘業務同業（軟體）：NTM EV/EBITDA（SBC 加回口徑）").font = GREY; r += 1
SW0 = r
for nm, tk, m in (("Microsoft", "MSFT", 15.59), ("SAP", "SAP", 15.02), ("Salesforce", "CRM", 10.8), ("Adobe", "ADBE", 6.6), ("ServiceNow", "NOW", 20.65)):
    V.cell(r, 1, nm); V.cell(r, 2, tk); c = V.cell(r, 4, m); c.font = BLUE; c.number_format = "0.00"
    V.cell(r, 5, "2026-10-07"); V.cell(r, 6, f"oracle/data/oracle_facts_20261007.json val.peer.{tk.lower()}（EV ÷ 共識 EBITDA 時間加權，Derived；2026-10-22 找不到新值）"); r += 1
SW_PEERS = (SW0, r - 1)
for _r in range(4, IN.max_row + 1):
    k_ = IN.cell(_r, 1).value
    if k_ == "mult_oci" and "mult_oci" not in OVR: IN.cell(_r, 5).value = f"=MEDIAN(Val!D{OCI_PEERS[0]}:D{OCI_PEERS[1]})"
    if k_ == "mult_sw" and "mult_sw" not in OVR: IN.cell(_r, 5).value = f"=MEDIAN(Val!D{SW_PEERS[0]}:D{SW_PEERS[1]})"
    if k_ == "mult_leg" and "mult_leg" not in OVR: IN.cell(_r, 5).value = f"=MEDIAN(Val!D{SW0+2},Val!D{SW0+3})"
r += 1
hdr(V, r, ["項目", "單位"] + [f"情境 {s}" for s in SC] + ["說明"]); r += 1
SO_ = {}
col = lambda s: get_column_letter(3 + SC.index(s))
def vr(k, lab, unit, f, fmt="#,##0.00", note=None):
    global r
    V.cell(r, 1, lab); V.cell(r, 2, unit)
    for j, s in enumerate(SC):
        V.cell(r, 3 + j, f(s)).number_format = fmt
    if note: V.cell(r, 6, note)
    SO_[k] = r; r += 1
vr("e_oci", "FY2029 OCI AI EBITDA（租金後）", "US$B", lambda s: f"=Thesis!G{TR[s]['ebitda']}", note="Thesis 頁")
for k, y in enumerate(range(2029, 2033), 1):
    vr(f"aif{y}", f"  OCI AI 自由現金流 FY{y}（EBITDA − 稅 − AI 資本支出＋預付淨額）", "US$B",
       lambda s, y=y: f"=Thesis!{tc(y)}{TR[s]['ebitda']}-etr*Thesis!{tc(y)}{TR[s]['contrib']}-Thesis!{tc(y)}{TR[s]['cx']}+CF!{yc(y)}{CR_[s]['ppin']}-CF!{yc(y)}{CR_[s]['pprec']}")
vr("ai_tv", "  OCI AI 永續價值（FY2032 年底；穩態汰換、RONIC＝WACC）", "US$B",
   lambda s: f"=(Thesis!J{TR[s]['ebitda']}-Thesis!J{TR[s]['conn']}*cx_vr200/life_ai/1000)*(1-etr)*(1+tp_g)*(1-tp_g/ronic)/(wacc-tp_g)", "#,##0.0")
vr("ev_oci", "OCI AI 企業價值（AI 模組 DCF，FY2028 年底；主值）", "US$B",
   lambda s: "=" + "+".join(f"{col(s)}{SO_[f'aif{y}']}/(1+wacc)^{k - 0.5}" for k, y in enumerate(range(2029, 2033), 1)) + f"+{col(s)}{SO_['ai_tv']}/(1+wacc)^4", "#,##0.0",
   note="修正一（F3）：同一 WACC；FY2029–FY2032 期中折現到 FY2028 年底")
vr("ev_ocim", "  對照：OCI AI 企業價值（× neocloud 倍數 mult_oci）", "US$B", lambda s: f"={col(s)}{SO_['e_oci']}*mult_oci*mult_adj", "#,##0.0", note="v0.4 原主值；只作對照")
vr("e_rest", "FY2029 其餘業務調整後 EBITDA（合併調整後 EBITDA − OCI AI；SBC 加回）", "US$B", lambda s: f"=PL!F{PR_[s]['aebitda']}-{col(s)}{SO_['e_oci']}", note="SBC 加回（與同業倍數同口徑，F2）")
vr("g_sh", "  其中成長型（SaaS＋OCI 非 AI）占比（依分部利潤貢獻分攤）", "%",
   lambda s: f"=(m_sw+m_adj)*(PL!F{PR_[s]['saas']}+PL!F{PR_[s]['oth']})/((m_sw+m_adj)*(PL!F{PR_[s]['saas']}+PL!F{PR_[s]['oth']}+PL!F{PR_[s]['lic']}+PL!F{PR_[s]['sup']})+m_hw*PL!F{PR_[s]['hw']}+m_svc*PL!F{PR_[s]['svc']})", "0.0%")
vr("ev_g", "成長型企業價值（× 軟體同業中位數 mult_sw）", "US$B", lambda s: f"={col(s)}{SO_['e_rest']}*{col(s)}{SO_['g_sh']}*mult_sw*mult_adj", "#,##0.0")
vr("ev_l", "傳統業務企業價值（授權、支援、硬體、服務 × mult_leg）", "US$B", lambda s: f"={col(s)}{SO_['e_rest']}*(1-{col(s)}{SO_['g_sh']})*mult_leg*mult_adj", "#,##0.0", note="修正一：成熟軟體倍數（Salesforce、Adobe 中位數）")
vr("ev", "企業價值（FY2028 年底）", "US$B", lambda s: f"={col(s)}{SO_['ev_oci']}+{col(s)}{SO_['ev_g']}+{col(s)}{SO_['ev_l']}", "#,##0.0")
vr("nd", "淨借款（FY2028 年底：借款＋融資租賃 − 現金）", "US$B", lambda s: f"=CF!E{CR_[s]['debt']}+finlease_q1-CF!E{CR_[s]['cash1']}", "#,##0.0")
vr("sh", "完全稀釋股數（FY2028 年底）", "十億股", lambda s: f"=CF!E{CR_[s]['sh1']}+opt_dil+mcps_sh", "0.000")
vr("ps", "分部加總每股價值（FY2028 年底）", "US$", lambda s: f"=MAX(0,({col(s)}{SO_['ev']}-{col(s)}{SO_['nd']})/{col(s)}{SO_['sh']})")
vr("v12", "分部加總每股價值（折回目標價時點，加回期間股利）", "US$", lambda s: f"=({col(s)}{SO_['ps']}+dps_q*4*t_28)/(1+ke)^t_28", note="FY2028 年底 → 目標價時點（Val 頁 t_28）")
for j in range(3): V.cell(SO_["v12"], 3 + j).fill = KEY
vr("v12m", "  對照：分部加總每股（OCI AI 以同業倍數）", "US$", lambda s: f"=(MAX(0,({col(s)}{SO_['ev']}-{col(s)}{SO_['ev_oci']}+{col(s)}{SO_['ev_ocim']}-{col(s)}{SO_['nd']})/{col(s)}{SO_['sh']})+dps_q*4*t_28)/(1+ke)^t_28")
vr("oci_sh", "OCI AI 占分部加總企業價值", "%", lambda s: f"={col(s)}{SO_['ev_oci']}/{col(s)}{SO_['ev']}", "0.0%")
r += 1
V.cell(r, 1, "0.3 結論").font = BOLD; r += 1
hdr(V, r, ["項目", "單位"] + [f"情境 {s}" for s in SC] + ["說明"]); r += 1
vr("tps", "情境目標價（DCF × w_dcf＋分部加總 ×（1 − w_dcf））", "US$", lambda s: f"=w_dcf*{col(s)}{DC['v12']}+(1-w_dcf)*{col(s)}{SO_['v12']}")
hdr(V, r, ["項目", "單位", "數值", "對照", "對照", "說明"]); r += 1
CON = {}
def cr(k, lab, unit, f, fmt="#,##0.00", note=None, d=None, e=None, fill=False):
    global r
    V.cell(r, 1, lab); V.cell(r, 2, unit); V.cell(r, 3, f).number_format = fmt
    if d is not None: V.cell(r, 4, d[0]).number_format = d[1]
    if e is not None: V.cell(r, 5, e[0]).number_format = e[1]
    if note: V.cell(r, 6, note)
    if fill:
        for c in range(1, 7): V.cell(r, c).fill = KEY
    CON[k] = r; r += 1
wsum = lambda row: f"=wA*C{row}+wB*D{row}+wC*E{row}"
cr("dcf", "DCF 目標價（情境加權）", "US$", wsum(DC["v12"]))
cr("comp", "Comparable 目標價（市場主流同業法）", "US$", wsum(SO_["v12"]), note="分部加總（OCI AI 以 AI 模組 DCF；SaaS 與 OCI 非 AI × 軟體倍數；傳統業務 × 成熟軟體倍數）情境加權")
cr("tp", "目標價（DCF 與 Comparable 加權）", "US$", f"=w_dcf*C{r-2}+(1-w_dcf)*C{r-1}", fill=True, note='="DCF 權重 "&TEXT(w_dcf,"0%")&"、分部加總權重 "&TEXT(1-w_dcf,"0%")&"；情境權重 A／B／C "&TEXT(wA,"0%")&"／"&TEXT(wB,"0%")&"／"&TEXT(wC,"0%")')
cr("px", "現價（2026-10-26）", "US$", "=px")
cr("cons", "市場共識目標價", "US$", "=cons_tp", note="S&P（43 家平均，2026-10-20）；Interested-party")
cr("up", "上檔空間：加權目標價", "%", f"=C{CON['tp']}/px-1", "0.0%", fill=True)
cr("upc", "上檔空間：Comparable 目標價", "%", f"=C{CON['comp']}/px-1", "0.0%")
cr("tier", "上檔空間級距", "—", f'=IF(C{CON["up"]}>0.5,"大於 50%",IF(C{CON["up"]}>0.25,"25%–50%",IF(C{CON["up"]}>=0,"0%–25%",IF(C{CON["up"]}>=-0.25,"0%～−25%",IF(C{CON["up"]}>=-0.5,"−25%～−50%","小於 −50%")))))', "@")
cr("rec", "投資建議", "—", f'=IF(C{CON["up"]}>=0.5,"強力買進",IF(C{CON["up"]}>=0.25,"買進",IF(C{CON["up"]}>-0.25,"中立",IF(C{CON["up"]}>-0.5,"放空","強力放空"))))', "@", fill=True,
   note="五級固定規則：≥ +50% 強力買進、+25%～+50% 買進、−25%～+25% 中立、−50%～−25% 放空、≤ −50% 強力放空")
cr("band", "價格帶（DCF 不如預期～順利）", "US$", f"=E{DC['v12']}", d=(f"=C{DC['v12']}", "#,##0.00"), note="C＝情境 C 的 DCF 價值（悲觀情境價）、D＝情境 A 的 DCF 價值（皆推到 12 個月後）")
cr("bear", "悲觀情境價（情境 C DCF）", "US$", f"=E{DC['v12']}")
cr("dcfvs", "DCF 情境價值推到 12 個月後（A／B／C）", "US$", f"=C{DC['v12']}", d=(f"=D{DC['v12']}", "#,##0.00"), e=(f"=E{DC['v12']}", "#,##0.00"))
cr("gcons", "與市場共識目標價的差距", "%", f"=C{CON['tp']}/cons_tp-1", "0.0%")
cr("tvsh", "DCF 永續價值占比（情境 B）", "%", f"=D{DC['tvsh']}", "0.0%")
cr("compm", "對照：加權目標價（分部加總的 OCI AI 改用同業倍數）", "US$", f"=w_dcf*C{CON['dcf']}+(1-w_dcf)*({wsum(SO_['v12m'])[1:]})", note="v0.4 原口徑；差額＝AI 資本密集度（F3）")
cr("wchk", "檢查：情境權重合計（應為 1）", "x", "=ROUND(wA+wB+wC,6)", "0.000")
cr("meth", "目標價方法", "—", "DCF（UFCF、WACC、永續 g、RONIC＝WACC）與分部加總（FY2029 EBITDA × 同業前瞻倍數；OCI AI 以 AI 模組 DCF）各半；情境 A／B／C 權重 25／50／25。偏誤方向：OCI 同業只有兩家（不符合至少三家，只作對照）；傳統業務倍數取成熟軟體兩家；永續價值占 DCF 比例高，對 WACC 與 g 敏感（Sens 頁）", "@")
V.column_dimensions["A"].width = 62; V.column_dimensions["B"].width = 12
for c in "CDE": V.column_dimensions[c].width = 14
V.column_dimensions["F"].width = 90
VAL = {"DC": DC, "SO": SO_, "CON": CON, "PEERS": [OCI_PEERS, SW_PEERS]}


# =====================================================================================
# Recon
# =====================================================================================
RC = wb.create_sheet("Recon")
RC["A1"] = "Recon — 逐分部、逐營收線對照市場共識與公司指引（校準時未使用）；另列 v0.4 vs v0.3.1。共識不分部（只有合併營收、EBITDA、EPS）：分部層的共識「未找到」，OCI 以公司營收路徑對照"
hdr(RC, 3, ["項目", "對照值", "模型 A", "差距 A", "模型 B", "差距 B", "模型 C", "差距 C", "對照來源", "說明"])
r = 4
def rec(lab, ref, model, src, note="", fmt="#,##0.00", pct=True):
    global r
    RC.cell(r, 1, lab)
    if ref is not None: RC.cell(r, 2, ref).number_format = fmt
    for j, s in enumerate(SC):
        c = 3 + 2 * j
        RC.cell(r, c, model(s)).number_format = fmt
        if ref is not None:
            RC.cell(r, c + 1, f"={get_column_letter(c)}{r}/B{r}-1" if pct else f"={get_column_letter(c)}{r}-B{r}").number_format = "0.0%" if pct else fmt
    RC.cell(r, 9, src); RC.cell(r, 10, note); r += 1
RC.cell(r, 1, "A. 合併（共識）").font = BOLD; r += 1
for y in (2027, 2028, 2029):
    rec(f"合併營收 FY{y} vs 共識", f"=cr{y}", lambda s, y=y: f"=PL!{yc(y)}{PR_[s]['rev']}", "MarketScreener（S&P 系）")
for y in (2027, 2028, 2029):
    rec(f"調整後 EBITDA FY{y} vs 共識（非 GAAP 營業利益＋折舊）", f"=ce{y}", lambda s, y=y: f"=PL!{yc(y)}{PR_[s]['aebitda']}", "同上", "模型＝GAAP 營業利益＋D&A＋SBC；共識口徑不含攤銷，略有差異")
for y in (2027, 2028, 2029):
    rec(f"GAAP EPS FY{y} vs 共識", f"=cp{y}", lambda s, y=y: f"=PL!{yc(y)}{PR_[s]['eps']}", "MarketScreener（2026-10-07 值）")
for y in (2027, 2028, 2029):
    rec(f"資本支出 FY{y} vs 共識", f"=cx{y}", lambda s, y=y: f"=CF!{yc(y)}{CR_[s]['cx']}", "MarketScreener（2026-10-07 值）", "模型只建已簽站點（上限 8,778 MW-IT），FY2029 起資本支出低於共識：觀點（見報告第 5 節）")
for y in (2027, 2028, 2029):
    rec(f"自由現金流 FY{y} vs 共識（差距以金額計）", f"=cf{y}", lambda s, y=y: f"=CF!{yc(y)}{CR_[s]['fcf']}", "MarketScreener（2026-10-07 值）", "", pct=False)
rec("FY2027 營收 vs 公司指引下限 90", "=g_rev27", lambda s: f"=PL!D{PR_[s]['rev']}", "Q1 FY27 新聞稿", "指引為下限")
rec("FY2027 資本支出 vs 公司指引 90–95", "=capex_g27", lambda s: f"=CF!D{CR_[s]['cx']}", "Q1 FY27 法說（二手）", "本版以指引中點為錨（恆等）")
RC.cell(r, 1, "B. 逐分部與營收線（共識未分部：未找到；OCI 對照公司路徑）").font = BOLD; r += 1
for y in (2027, 2028, 2029, 2030):
    rec(f"Cloud Infrastructure（OCI）FY{y} vs 公司 OCI 營收路徑", f"=go{y}", lambda s, y=y: f"=PL!{yc(y)}{PR_[s]['oci']}", "Oracle Q4 FY26 新聞稿（Interested-party）", "公司路徑不作校準")
for lab, k in (("Cloud Applications（SaaS）", "saas"), ("Software license", "lic"), ("Software support", "sup"), ("Hardware", "hw"), ("Services", "svc")):
    rec(f"{lab} FY2027：共識未找到（列模型值）", None, lambda s, k=k: f"=PL!D{PR_[s][k]}", "未找到", "共識不分營收線")
for lab, k in (("分部：Cloud and software", "cs"), ("分部：Hardware", "shw"), ("分部：Services", "ssvc")):
    rec(f"{lab} FY2027：共識未找到（列模型值）", None, lambda s, k=k: f"=PL!D{PR_[s][k]}", "未找到", "共識不分部")
RC.cell(r, 1, "C. 命題檢查（不為對上而調參數）").font = BOLD; r += 1
rec("Q1 FY27 OCI 非 AI 殘差占比 vs v0.3.1 區間上緣 30%", 0.30, lambda s: f"=Thesis!C{TR[s]['oth']}/q1_oci", "v0.3.1 E2", "殘差包含 Oracle 合約價高於 k 的部分", "0.0%", pct=False)
rec("m_sw（Q1 校準）vs FY2024 雲端與軟體分部利潤率", "=Q!B45", lambda s: "=m_sw", "Q 頁", "兩者接近＝Q1 的 OCI AI 貢獻估計合理", "0.0%", pct=False)
rec("FY2029 OCI AI 營收 vs OpenAI 合約年額 60", "=oai_annual", lambda s: f"=Thesis!G{TR[s]['rev']}", "WSJ 轉述（未證實）", "")
rec("FY2029 OCI AI 營業利益率 vs 公司說法：AI 合約毛利率 30–40% 中點", 0.35, lambda s: f"=Thesis!G{TR[s]['om']}", "2025-10 分析師日（Interested-party）", "口徑不同：模型為營業利益率（含折舊、租金）；外部報導毛利率 13.9%（The Information）", "0.0%", pct=False)
rec("FY2029 AI 增量 ROIC vs WACC", "=wacc", lambda s: f"=Thesis!G{TR[s]['roic']}", "Val 頁", "低於 WACC＝AI 擴張摧毀價值", "0.0%", pct=False)
RC.cell(r, 1, "D. v0.4 vs v0.3.1（目標價與主要差異來源）").font = BOLD; r += 1
rec("情境目標價 vs v0.3.1（A 對積極、B 對基準、C 對保守）", None, lambda s: f"=Val!{get_column_letter(3+SC.index(s))}{VAL['SO']['tps']}", "v0.3.1 交接檔", "")
for j, s in enumerate(SC):
    RC.cell(r - 1, 2 + 0, None)
RC.cell(r, 1, "v0.3.1 情境目標價（積極／基準／保守）"); RC.cell(r, 3, "=tp_v031_hi"); RC.cell(r, 5, "=tp_v031"); RC.cell(r, 7, "=tp_v031_lo")
for c in (3, 5, 7): RC.cell(r, c).number_format = "#,##0.00"
RC.cell(r, 10, "v0.3.1：DCF 45%＋分部 EV/EBITDA 55%（OCI 6x、傳統事業 15.02x，錨定 FY29）"); r += 1
RC.cell(r, 1, "差異來源（定性；數字拆解見報告第 6 節）"); RC.cell(r, 10, "(1) 營收結構：SaaS、授權、支援、硬體、服務各自路徑，OCI 拆 AI／非 AI；(2) AI 折舊與租金進分部利潤與 EPS；(3) 評價法：DCF 50%＋分部加總 50%、OCI 倍數改為同業中位數、情境權重 25／50／25；(4) 市場資料：現價 129.13 → 126.06"); r += 1
RC.column_dimensions["A"].width = 62
for c in "BCDEFGH": RC.column_dimensions[c].width = 12
RC.column_dimensions["I"].width = 36; RC.column_dimensions["J"].width = 80


# =====================================================================================
# Sens：建置時以完整模型重算（scripts/sens_v04.py）
# =====================================================================================
SE = wb.create_sheet("Sens")
SE["A1"] = "Sens — 一次動一個輸入（低／高），以完整模型重算加權目標價與 FY2029 指標（scripts/sens_v04.py 建置時以 LibreOffice 重算整本活頁簿；數值寫入此頁）"
SE["A2"] = "未重算（data/sens_v04.json 不存在）" if SENS is None else f"重算日期 {SENS['date']}；基準加權目標價 {SENS['base']['tp']:.4f}（應等於 Val 頁加權目標價，見下方檢查列）"
hdr(SE, 3, ["輸入", "低／高設定", "低：加權目標價", "低：相對基準", "高：加權目標價", "高：相對基準", "最大絕對變動", "低：FY2029 OCI AI 營業利益率（B）", "高：FY2029 OCI AI 營業利益率（B）", "低：收支關卡（B）", "高：收支關卡（B）"])
SENS_ROWS = (4, 4); SENS_GRID = None
if SENS:
    rr = 4
    for case in SENS["cases"]:
        SE.cell(rr, 1, case["label"]); SE.cell(rr, 2, case["setting"])
        SE.cell(rr, 3, case["lo"]["tp"]).number_format = "#,##0.00"; SE.cell(rr, 4, f"=C{rr}/Val!$C${CON['tp']}-1").number_format = "0.0%"
        SE.cell(rr, 5, case["hi"]["tp"]).number_format = "#,##0.00"; SE.cell(rr, 6, f"=E{rr}/Val!$C${CON['tp']}-1").number_format = "0.0%"
        SE.cell(rr, 7, f"=MAX(ABS(D{rr}),ABS(F{rr}))").number_format = "0.0%"
        SE.cell(rr, 8, case["lo"]["om29"]).number_format = "0.0%"; SE.cell(rr, 9, case["hi"]["om29"]).number_format = "0.0%"
        SE.cell(rr, 10, case["lo"]["gate"]).number_format = "#,##0.0"; SE.cell(rr, 11, case["hi"]["gate"]).number_format = "#,##0.0"
        for c in (3, 5, 8, 9, 10, 11): SE.cell(rr, c).font = BLUE
        rr += 1
    SENS_ROWS = (4, rr - 1)
    SE.cell(rr + 1, 1, "檢查：重算基準加權目標價 − Val 頁加權目標價（應為 0）"); SE.cell(rr + 1, 3, f"={SENS['base']['tp']}-Val!C{CON['tp']}").number_format = "0.0000"
    SENS_CHK = rr + 1
    if SENS.get("grid"):
        g0 = rr + 5
        SE.cell(g0, 1, "WACC × 永續成長率 g：加權目標價（完整重算；列＝WACC 調整、欄＝g）").font = BOLD
        hdr(SE, g0 + 1, ["WACC 調整", "", "g＝2%", "g＝3%", "g＝4%"])
        for i, w in enumerate((-0.01, 0.0, 0.01)):
            SE.cell(g0 + 2 + i, 1, f"WACC {'−1pt' if w < 0 else ('基準' if w == 0 else '+1pt')}")
            for j, g in enumerate((0.02, 0.03, 0.04)):
                v = [x["tp"] for x in SENS["grid"] if abs(x["wacc_adj"] - w) < 1e-9 and abs(x["tp_g"] - g) < 1e-9][0]
                c = SE.cell(g0 + 2 + i, 3 + j, v); c.number_format = "#,##0.00"; c.font = BLUE
        SENS_GRID = (g0, g0 + 4)
    SE.cell(rr + 3, 1, "說明：tornado 以「低：相對基準」「高：相對基準」兩欄；情境結構（A／B／C）不變，只動一個輸入。容量情境本身的差異見 Val 頁情境目標價")
SE.column_dimensions["A"].width = 46; SE.column_dimensions["B"].width = 40
for c in "CDEFGHIJK": SE.column_dimensions[c].width = 14


# =====================================================================================
# Sources
# =====================================================================================
SO = wb.create_sheet("Sources")
SO["A1"] = "來源（每列：編號、文件、日期、網址、標記）。SEC 網站在建置環境被擋（403）：SEC 文件數字取自 v0.1a 事實總帳（oracle/data/oracle_facts_20261007.json，當時直接自 EDGAR 取得）"; SO["A1"].font = BOLD
hdr(SO, 2, ["編號", "文件", "日期", "網址", "標記"])
srcs = [("Oracle FY2026 10-K（附註 13 分部、經查核財報）", "2026-06-22", "https://www.sec.gov/Archives/edgar/data/1341439/000119312526277521/orcl-20260531.htm", "Verified"),
        ("Oracle FY27 Q1 10-Q（附註 6 租賃、附註 9 分部）", "2026-09-11", "https://www.sec.gov/Archives/edgar/data/1341439/000119312526389274/orcl-20260831.htm", "Interested-party"),
        ("Oracle Q1 FY27 財報新聞稿（8-K 附件 99.1）", "2026-09-10", "https://www.sec.gov/Archives/edgar/data/1341439/000119312526387905/orcl-ex99_1.htm", "Interested-party"),
        ("Oracle Q4 FY26 財報新聞稿（OCI 營收路徑）", "2026-06-10", "https://s23.q4cdn.com/440135859/files/doc_earnings/2026/q4/earnings-result/4q26-pressrelease-final.pdf", "Interested-party"),
        ("Oracle 事實總帳 v0.1a（434 筆，SEC 一手）", "2026-10-07", "oracle/data/oracle_facts_20261007.json", "Verified／Interested-party"),
        ("Oracle 事實總帳 v0.3.1 增補（現價、共識、美債、同業 EV）", "2026-10-22", "oracle/data/oracle_facts_20261022.json", "Verified／Interested-party"),
        ("市場共識（MarketScreener／S&P、LSEG）", "2026-10-22", "oracle/data/consensus_orcl_20261022.json", "Interested-party"),
        ("現價：stockanalysis.com 歷史股價（2026-10-26 收盤 126.06）", "2026-10-27", "https://stockanalysis.com/stocks/orcl/history/", "Verified"),
        ("現價對照：markets/invest_map/data/prices.csv（2026-10-23 收盤 126.16）", "2026-10-23", "markets/invest_map/data/prices.csv", "Verified"),
        ("Tokenomics v21.4 快照（IF_HoldEcon、IF_CapexIT、IF_OpexGW、IF_DeprLifeIT）", "2026-10-25", "oracle/data/tokenomics_snapshot_v21.4.json（YinchenChang/Tokenomics 3ac70ac；與 microsoft/data 同檔）", "Derived"),
        ("IREN–Microsoft GB300 五年約 $9.7bn", "2025-11-03", "https://iren.gcs-web.com/news-releases/news-release-details/iren-secures-97bn-ai-cloud-contract-microsoft", "Derived"),
        ("Silicon Data H100 Index", "2026-10-06", "https://www.silicondata.com/products/silicon-index/h100", "Verified"),
        ("Oracle–OpenAI 約 $300bn／4.5 GW（WSJ 轉述）", "2025-09-11", "https://www.theregister.com/2025/09/11/openai_reportedly_on_the_hook/", "Interested-party"),
        ("CoreWeave 共識（MarketScreener）", "2026-09-25", "coreweave/data/consensus_crwv_20260925.json", "Interested-party"),
        ("Nebius 共識（MarketScreener）", "2026-10-07", "nebius/data/consensus_nbis_20261007.json", "Interested-party"),
        ("v0.3.1 舊引擎參數（company.json）與交接檔", "2026-10-22", "oracle/legacy_v0.3/company.json；oracle/docs/handoff/20261022_Oracle收支模型_交接檔_v0_3_1.md", "—"),
        ("研究筆記：官方營收線、主要客戶（新聞稿、10-K 風險因子）", "2026-10-27", "研究檔（本專案研究代理）", "Interested-party"),
        ("SemiAnalysis", "—", "未使用", "—")]
for i, (a, d, u, t) in enumerate(srcs, 3):
    SO.cell(i, 1, i - 2); SO.cell(i, 2, a); SO.cell(i, 3, d); SO.cell(i, 4, u); SO.cell(i, 5, t)
SO.column_dimensions["B"].width = 70; SO.column_dimensions["C"].width = 12; SO.column_dimensions["D"].width = 110; SO.column_dimensions["E"].width = 22


# =====================================================================================
# README（公式引用數字）
# =====================================================================================
RD = wb.create_sheet("README", 0)
B_ = PR_["B"]; TB = TR["B"]; CB = CR_["B"]
up = f"Val!C{CON['up']}"; tp = f"Val!C{CON['tp']}"
lines = [
    ("Oracle（ORCL）個股模型 v0.4（命題＋營收分部架構）　2026-10-27", True),
    ("命題：OCI 已簽約的 backlog（RPO 約 6,640 億美元，約一半來自 OpenAI／Stargate）能否以高於資金成本的報酬轉成營收與現金，同時不把資產負債表撐破。命題模組（已連網 MW × 每 MW 年收入〔Tokenomics 持有成本 × 定價倍數 k〕、延誤、租金、AI 折舊）只驅動 Cloud Infrastructure 營收線的「OCI AI 算力」子列；SaaS、授權、支援、硬體、服務各有自己的成長路徑。", False),
    ("核心問題：FY2029 的 OCI 營收與營業利益率，相對於資本支出與租約的擴張，以及 Oracle 的資金缺口（自由現金流為負）要如何補。", False),
    (f'="投資建議："&Val!C{CON["rec"]}&"（加權目標價 US$"&TEXT({tp},"0.0")&"，現價 US$"&TEXT(px,"0.00")&"〔2026-10-26〕，上檔空間 "&TEXT({up},"+0.0%;-0.0%")&"）"', True),
    (f'="重點觀察 1：情境 B 的 FY2029 OCI 營收 "&TEXT(PL!F{B_["oci"]},"0.0")&" 十億美元（其中 AI 算力 "&TEXT(PL!F{B_["ai"]},"0.0")&"），只有公司路徑 114 的 "&TEXT(Thesis!D{RPO_ROWS[1]},"0%")&"；OCI AI 營業利益率 "&TEXT(Thesis!G{TB["om"]},"0.0%")&"，AI 增量 ROIC "&TEXT(Thesis!G{TB["roic"]},"0.0%")&"，對照 WACC "&TEXT(wacc,"0.0%")&"。"', False),
    (f'="重點觀察 2：收支關卡（FY2027–FY2031 自由現金流扣股利合計）情境 A／B／C 為 "&TEXT(CF!C{GATE},"0")&"／"&TEXT(CF!D{GATE},"0")&"／"&TEXT(CF!E{GATE},"0")&" 十億美元；情境 B 需新增投資級債 "&TEXT(CF!D{GATE+1},"0")&"、新股 "&TEXT(CF!D{GATE+2},"0")&"、高息債 "&TEXT(CF!D{GATE+3},"0")&"，調整後槓桿峰值 "&TEXT(CF!D{GATE+4},"0.00")&" 倍。"', False),
    (f'="重點觀察 3：DCF 情境加權 "&TEXT(Val!C{CON["dcf"]},"0.0")&"、分部加總 "&TEXT(Val!C{CON["comp"]},"0.0")&"（OCI AI 以 AI 模組 DCF 計，只占情境 B 分部企業價值的 "&TEXT(Val!D{SO_["oci_sh"]},"0%")&"）。若 OCI AI 改用 neocloud 同業倍數，加權目標價為 "&TEXT(Val!C{CON["compm"]},"0.0")&"：差額就是 AI 的資本密集度（汰換資本支出幾乎吃掉 EBITDA）；市場共識目標價 "&TEXT(cons_tp,"0")&"，高出本模型 "&TEXT(cons_tp/{tp}-1,"0%")&"。"', False),
    ("", False),
    ("會計年度：FY(t)＝t−1 年 6 月～t 年 5 月。FY2027＝Q1 實績（2026-06～08）＋Q2–Q4 模型；延伸到 FY2032。金額十億美元，MW 為 IT 口徑。", False),
    ("營收分部層：官方報告分部（10-K 附註 13）＝Cloud and software（雲端與軟體）、Hardware、Services；其下營收線（新聞稿）＝Cloud Infrastructure（OCI）、Cloud Applications（SaaS）、Software license、Software support（以上屬 Cloud and software）、Hardware、Services。命題子列：OCI＝OCI AI 算力（命題模組）＋OCI 其他（非 AI，成長路徑）；子列合計＝母列（PL 頁核對列）。", False),
    ("驅動 → 推導對照（每頁輸入、推導量、因果方向）", True),
    ("Thesis：In 頁 MW 速度、上限、延誤、計費比例、世代組合、Tokenomics 每 MW 成本、k → 已連網／計費 MW → OCI AI 營收；營運費用＝在役 MW × IF_OpexGW；租金＝未起租 288B 起租；AI 資本支出（FY27＝指引中點、其後新增 MW × IF_CapexIT、汰換依世代年份）→ AI 折舊 → OCI AI 分部利潤貢獻 → ROIC。", False),
    ("Seg：近四季年增率（Q 頁）→ 線性收斂到長期年增（In 頁 G 節）→ SaaS、授權、支援、硬體、服務；OCI＝Thesis 兩子列 → 三分部 → 合併營收。", False),
    ("PL：分部利潤＝OCI AI 貢獻＋其餘雲端與軟體營收 × m_sw（Q1 FY27 校準，對照 FY2024 64.1%）；硬體、服務 × FY2026 利潤率；− 未分攤費用 → GAAP 營業利益 → 利息（CF 頁）、稅 → 淨利 → EPS。", False),
    ("CF：淨利＋D&A＋SBC＋客戶預付 − 預付認列 → 營運現金；− 資本支出 → 自由現金流；− 股利 → 收支關卡；缺口 → 現金 → 投資級債 → 新股 → 高息債（融資瀑布沿用 v0.3.1）→ 股數、淨借款、槓桿。", False),
    ("Val：DCF（UFCF、WACC、g 3%）50%＋分部加總（OCI AI × neocloud 倍數、其餘 × 軟體倍數，FY2029 EBITDA）50%；情境 25／50／25；五級投資建議。Recon：共識與公司路徑對照；Sens：一次動一個輸入的完整重算。", False),
    ("", False),
    ("與 v0.3.1 的差異：v0.3.1 是 CRWV 樣板（OCI MW × 每 MW 收入＋傳統事業四線壓成成長率），本版改為官方分部與六條營收線；AI 營運費用由 Tokenomics 正向推導（取代 EBITDAR 率校準）、AI 折舊與租金進分部利潤與 EPS；評價改為 DCF 50%＋分部加總 50%。舊引擎保留在 oracle/legacy_v0.3/。", False),
    ("標記：Verified、Interested-party、Analogy、Assumed、Derived；SemiAnalysis 未使用。SEC 網站在建置環境被擋，SEC 數字取自 v0.1a 事實總帳。", False),
    ("已知限制：(1) OCI 的 AI／非 AI 拆分不存在公開資料，Q1 非 AI 為殘差；(2) 季度分部利潤只有 Q1 FY27 與 FY26Q1；(3) OCI 同業倍數只有兩家；(4) 預付與營運資金外的營運資金變動不建模；(5) 在帳租賃負債持平、到期借款以原利率展延；(6) 續約價格衰退不建模（公司稱續約 +20%）。", False),
]
for i, (t, b) in enumerate(lines, 1):
    RD.cell(i, 1, t).font = Font(bold=b)
RD.column_dimensions["A"].width = 170

order = ["README", "In", "Q", "Thesis", "Seg", "PL", "CF", "Val", "Recon", "Sens", "Sources"]
wb._sheets = [wb[n] for n in order]
wb.save(sys.argv[1])
if os.environ.get("ORCL_ROWMAP"):
    json.dump({"TR": TR, "SR": SR, "PR": PR_, "CR": CR_, "VAL": VAL, "GATE": GATE, "CHK": [CHK_SEG, CHK_SUB, CHK_SUM], "RPO": RPO_ROWS, "SUM": SUM_ROWS, "SENS": SENS_ROWS, "GRID": SENS_GRID, "CAL": CAL, "SGC": SGC},
              open(os.environ["ORCL_ROWMAP"], "w"), ensure_ascii=False, indent=1)
