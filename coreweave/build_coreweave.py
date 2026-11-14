# CoreWeave（CRWV）個股模型 v5.0（Oracle 收支型樣板重建＋2026Q3 滾入＋v4.13 驗收採用項；工作單 coreweave/docs/workorders/20261112_CoreWeave_v5.0_Oracle樣板重建_Q3.md）
# 用法：python3 coreweave/build_coreweave.py raw.xlsx ；再以 LibreOffice headless 重算（coreweave/make_v5.sh 一次跑完）。
# 會計年度＝日曆年。金額：十億美元（US$B）；MW＝IT 關鍵電力口徑（公司 active／contracted power 視為 IT，Assumed）。
# 工作表（順序固定）：README → In → Q → Thesis → Seg → PL → CF → Fund → RPO → Sites → Debt → Lease → Val → RevDCF → Recon → Sens → Sources
# 環境變數：SENS_OVR＝JSON {In 鍵: 值}（只替換基準值；敏感度、反解、拆解用）；CRWV_TKV＝Tokenomics 快照版本（拆解用，預設 v22.0）；
#           CRWV_ROWMAP＝列號對照輸出路徑（HTML 設定檔產生器讀）。成品三者都不設（ROWMAP 除外）。
import sys, os, json, hashlib, datetime
import openpyxl
from openpyxl.styles import Font, PatternFill
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.utils import get_column_letter as CL

HERE = os.path.dirname(os.path.abspath(__file__))
VER = "v5.0.3"
BUILT = "2026-11-14"
OVR = json.loads(os.environ.get("SENS_OVR", "{}"))
TKV = os.environ.get("CRWV_TKV", "v22.0")
TKS = json.load(open(os.path.join(HERE, "data", f"tokenomics_snapshot_{TKV}.json"), encoding="utf-8"))
TK = TKS["items"]
FACTS_F = "coreweave/data/coreweave_facts_20261112.json"
FX = {it["id"]: it for it in json.load(open(os.path.join(HERE, "data", "coreweave_facts_20261112.json"), encoding="utf-8"))["items"]}
SENS_F = os.path.join(HERE, "data", "sens_v50.json")
SENS = json.load(open(SENS_F, encoding="utf-8")) if (os.path.exists(SENS_F) and not OVR) else None
REV_F = os.path.join(HERE, "data", "reverse_v50.json")
REV = json.load(open(REV_F, encoding="utf-8")) if (os.path.exists(REV_F) and not OVR) else None
DEC_F = os.path.join(HERE, "data", "decomp_v50.json")
DEC = json.load(open(DEC_F, encoding="utf-8")) if (os.path.exists(DEC_F) and not OVR) else None


def _bid():
    h = hashlib.sha256()
    for p in (__file__, os.path.join(HERE, "data", "coreweave_facts_20261112.json"), os.path.join(HERE, "data", f"tokenomics_snapshot_{TKV}.json"), os.path.join(HERE, "data", "coreweave_is_hist_20261114.json")):
        h.update(open(p, "rb").read())
    return f"{VER}-{h.hexdigest()[:10]}"


BUILD_ID = _bid()
BLUE = Font(color="0000FF"); BOLD = Font(bold=True); GREY = Font(color="808080")
HDR = PatternFill("solid", fgColor="DDEBF7"); KEY = PatternFill("solid", fgColor="FFF2CC")
wb = openpyxl.Workbook()
SC = ("A", "B", "C")
SCN = {"A": "情境 A（樂觀：2030 年底在役 8 GW）", "B": "情境 B（基準：2030 年底在役 5.6 GW）", "C": "情境 C（保守：在役到已簽約 4.6 GW 為止）"}
GEN = ("h100", "gb200", "gb300", "vr200")
GNAME = {"h100": "Hopper H100", "gb200": "GB200 NVL72", "gb300": "GB300 NVL72", "vr200": "VR200 NVL72"}
YRS = list(range(2025, 2033))
yc = lambda y: CL(3 + y - 2025)              # Seg／PL／CF：C＝FY2025A、D＝FY2026、E＝FY2027 … J＝FY2032
TCOL = {"9m": "C", "q4": "D", 2026: "E", **{y: CL(6 + y - 2027) for y in range(2027, 2033)}}   # Thesis：C＝9M26、D＝Q4、E＝FY2026、F＝FY2027 … K＝FY2032
FYM = list(range(2027, 2033))
ROWMAP = {}


def hdr(ws, row, labels, c0=1):
    for j, h in enumerate(labels, c0):
        c = ws.cell(row, j, h); c.font = BOLD; c.fill = HDR


def dn(name, ref):
    wb.defined_names[name] = DefinedName(name, attr_text=ref)


def tkv(name, g, case="基準"):
    return round(TK[name]["values"][GNAME[g]][case], 4)


def fv(fid):
    return FX[fid]["v"]


SRCF = {"PR_Q3": "CoreWeave Q3 2026 新聞稿（2026-11-10；8-K 附件 99.1，investors.coreweave.com）", "Q2_10Q": "CoreWeave Q2 2026 10-Q（2026-08-12）",
        "CALL_Q3": "CoreWeave Q3 2026 法說（2026-11-10；chat 端提供，新聞稿未見）", "PR_0917": "CoreWeave 新聞稿 2026-09-17（finviz 轉載）",
        "PX": "debelegger.nl CRWV 頁（2026-11-11）；Motley Fool、heygotrade", "UST": "CAPP Notes Closing Look 2026-11-10：10 年期收 4.98%；heygotrade 2026-11-11 引 IBD 4.98%；CNBC 11-10「races back to 5%」", "K10_25": "CoreWeave FY2025 10-K"}
INC_CO = "公司自揭 KPI／非 GAAP：管理層有呈現成長與獲利的誘因"
INC_GUIDE = "管理層指引：有設定可達成目標、維持市場預期的誘因"
INC_SELL = "賣方共識：投行業務、承銷關係，偏多傾向"
INC_PEER = "同業公司自揭或賣方共識：有利害關係"
INC_TK = "—"

# =====================================================================================
# In
# =====================================================================================
IN = wb.active; IN.title = "In"
IN["A1"] = "In — 輸入（藍字＝輸入；黑字＝公式；基準欄 E 進主模型；D／F＝敏感度低／高，只用於 Sens 頁、不進目標價；情境 A／B／C 的驅動見本頁最後一節「情境驅動對照」）。金額：十億美元（US$B）；MW＝IT 口徑；每 MW 金額：百萬美元／MW·年（＝US$B／GW·年）"
IN["A2"] = "標記：每列單一主要標記 Verified／Interested-party／Analogy／Assumed／Derived（原則第 10 節第 2 條；Interested-party 列見 J 欄「來源誘因」）；Analogy、Assumed 給低／高區間。SemiAnalysis 未使用。「沿用 v4.13」＝沿用 coreweave/legacy_v4/company.json 已核准參數（原出處見來源欄）"
hdr(IN, 3, ["鍵", "項目", "單位", "敏感度低", "基準", "敏感度高", "標記", "來源", "說明", "來源誘因"])
rows = []
sec = lambda t: rows.append(("sec", t))
def inp(key, item, unit, lo, base, hi, tag, src, note="—", inc=None):
    if inc is None:
        inc = INC_CO if tag == "Interested-party" else ""
    rows.append((key, item, unit, lo, base, hi, tag, src, note, inc))

PRQ3 = SRCF["PR_Q3"]; Q2Q = SRCF["Q2_10Q"]; CALL = SRCF["CALL_Q3"]
sec("A. 股價、股數、日期")
inp("px", "現價（2026-11-11 收盤）", "US$", None, fv("px.1111"), None, "Verified", SRCF["PX"] + "：最新收盤 60.77", "v4.13 為 2026-11-05 收盤 43.66；第二來源：Motley Fool、heygotrade（獨立核對代理 2026-11-12 確認 11-11 收盤 60.77）；11-10 收盤 48.73（heygotrade），財報後單日 +24.7%")
inp("px_date", "現價日（px 的收盤日）", "日期", None, "=DATE(2026,11,11)", None, "Verified", "同 px", "README、Val 與 tp_date 以本列組字（原則第 10 節第 26、41 條）")
inp("cutoff", "資料截止日（Q3 新聞稿）", "日期", None, "=DATE(2026,11,10)", None, "Verified", PRQ3, "原則第 10 節第 41 條：三個日期分列")
inp("built", "成品日", "日期", None, "=DATE(2026,11,14)", None, "Verified", "建置日", "—")
inp("sh_out", "普通股流通在外（Q3 基本加權股數代理）", "十億股", None, fv("q3.waso"), None, "Verified", PRQ3 + "：Q3 基本加權 559 百萬股（" + FACTS_F + " q3.waso）", "期末流通股數待 Q3 10-Q 封面（Q2 10-Q 期末 0.5515）；ATM（上限 35 百萬股、約 6.63 十億美元）已發行數未揭露、未計入（資料缺口，待 Q3 10-Q）")
inp("rsu_net", "RSU＋選擇權庫藏股法淨股數", "十億股", None, 0, 0.03, "Assumed", "找不到：Q3 10-Q 在建置環境無法開啟（SEC 被擋），RSU 與選擇權餘額未取得", "DCF 標準規格第 3 節要求計入；暫以 0 並設 Sens 高 0.03（約 5%，Assumed）；資料缺口")
inp("val_date", "評價日（Q3 資產負債表日）", "日期", None, "=DATE(2026,9,30)", None, "Verified", PRQ3, "—")
inp("tp_date", "目標價時點（現價日＋12 個月）", "日期", None, "=EDATE(px_date,12)", None, "Derived", "公式（原則第 10 節第 26 條）", "評價日價值以 ke 推到本日（不需驗收清單 X20）")
inp("fy27_end", "Comparable 錨定日（FY2027 年底）", "日期", None, "=DATE(2027,12,31)", None, "Derived", "會計年度＝日曆年", "Comparable 以 FY2028 調整後 EBITDA 為前瞻值")

sec("B. 命題：在役 MW 路徑（情境 A／B／C；J2）")
inp("mw_ye25", "在役 MW（2025 年底）", "MW", None, 850, None, "Interested-party", SRCF["K10_25"] + "（legacy_v4/company.json defaults.mwYearEnd）", "公司 active power；口徑視為 IT（Assumed，同 v4.13 mwBasis）")
inp("mw_q2", "在役 MW（2026-06-30）", "MW", None, 1500, None, "Interested-party", Q2Q + "（法說 active power 1.5 GW）", "世代占比的基數")
inp("mw_q3", "在役 MW（2026-09-30）", "MW", None, fv("q3.active_mw"), None, "Interested-party", PRQ3 + "：over 1.7 GW（取下限）", "季增 200+ MW")
inp("mw_ye26", "在役 MW（2026 年底；三情境相同）", "MW", None, fv("ye26.active_mw"), None, "Interested-party", PRQ3 + "：more than 1.85 GW at year-end（公司預期）", "近期以公司預期為錨（只差一季）；Q4 新增 150 MW")
inp("ctr_mw", "已簽約電力（contracted power，2026-09-30）", "MW", None, fv("q3.ctr_mw"), None, "Interested-party", PRQ3 + "：over 4.6 GW（季增 900+ MW；國際 >1.3 GW）", "情境 C 的上限（J2：v4.13 的 4.2 GW 低於已簽約 4.6 GW → C 改為已簽約量）")
MWP = {"A": (3100, 4600, 6300, 8000, 1000), "B": (3100, 4000, 4800, 5600, 500), "C": (3100, 3800, "=ctr_mw", "=ctr_mw", 0)}
for s in SC:
    a = MWP[s]
    for j, y in enumerate((2027, 2028, 2029, 2030)):
        inp(f"mw{s}_{y % 100}", f"情境 {s}：{y} 年底在役 MW", "MW", None, a[j], None, "Assumed",
            "沿用 v4.13 scenarios.accepted（" + {"A": "積極 8 GW", "B": "基準 5.6 GW", "C": "保守 4.2 GW → v5.0 改為已簽約 4.6 GW（J2）"}[s] + "）", "公司 2030 年目標最多 8 GW（Interested-party，只作情境 A 的上緣）" if s == "A" and y == 2030 else "—")
    inp(f"add{s}_post", f"情境 {s}：2031–2032 每年新增在役 MW", "MW／年", None, a[4], None, "Assumed", "沿用 v4.13 defaults.mw31（低 0／基準 500／高 1,000）", "之後的永續成長由 DCF 的 g 處理")
inp("u_bill", "計費比例（計費 MW ÷ 平均在役 MW）", "x", 0.88, 0.92, 0.96, "Assumed", "沿用 v4.13 scenarios.billableRatio（計費 ÷ 驗收約 0.92–0.94）", "公司未揭露計費 MW（找不到）；可證偽指標：每季營收 ÷ 平均在役 MW（Recon 公司實況驗證）")

sec("C. 命題：世代組合與汰換（沿用 v4.13 fleet；Assumed）")
for g, v, lo, hi in (("h100", 0.2727, 0.24, 0.31), ("gb200", 0.2937, 0.20, 0.42)):
    inp(f"om_{g}", f"2026-06-30 在役世代占比：{GNAME[g]}", "%", lo, v, hi, "Assumed", "沿用 v4.13 fleet.openMix（W1 第 6b 步：依年份分層）", "公司未揭露機隊世代組合（找不到）")
inp("om_gb300", "2026-06-30 在役世代占比：GB300 NVL72（＝1 − 其餘）", "%", None, "=1-om_h100-om_gb200", None, "Derived", "合計 100%", "—")
inp("nm26_gb300", "2026 下半年新增產能中 GB300 占比（其餘 GB200）", "%", None, 0.8, None, "Assumed", "沿用 v4.13 fleet.newMix（2H26 GB300 80%／GB200 20%）", "—")
inp("nm27_gb300", "2027 年新增產能中 GB300 占比（其餘 VR200）", "%", 0.3, 0.5, 0.8, "Assumed", "沿用 v4.13 fleet.newMix（FY27 GB300 50%／VR200 50%）", "區間：VR200 提前或延後")
inp("nm28_gb300", "2028 年起新增產能中 GB300 占比（其餘 VR200）", "%", None, 0.0, 0.3, "Assumed", "沿用 v4.13 fleet.newMix（FY28–FY30 VR200 100%）", "—")
for k, y, v in (("ret29", 2029, 70), ("ret30", 2030, 290), ("ret31", 2031, 490)):
    inp(k, f"{y} 年汰換 MW（{y - 6} 年新增批次，換成 VR200）", "MW", None, v, None, "Interested-party", "legacy_v4/company.json fleet.vintages（2023 年 70、2024 年 290、2025 年 490；10-K 年底在役 MW 相減）", "經濟壽命 6 年（Tokenomics IF_DeprLifeIT 基準）")
inp("ret32", "2032 年汰換 MW（2026 年新增批次）", "MW", None, "=mw_ye26-mw_ye25", None, "Derived", "2026 年新增＝2026 年底 − 2025 年底在役", "—")
inp("v25_h100", "2025 年新增批次中 H100（H200）占比", "%", None, 0.1, None, "Assumed", "沿用 v4.13 fleet.vintages（GB200 70%／GB300 20%／H200 10%）", "—")
inp("v25_gb200", "2025 年新增批次中 GB200 占比", "%", None, 0.7, None, "Assumed", "同上", "—")
inp("v26_gb300", "2026 年新增批次中 GB300 占比（其餘 GB200）", "%", None, 0.835, None, "Assumed", "1H26 650 MW × 85%＋2H26 350 MW × 80%＝83.3%（v4.13 fleet.vintages、newMix）", "取 83.5%：使期初占比推得的 GB200 存量在 2032 年汰換後不為負（兩組 Assumed 占比的捨入差）")

sec(f"D. 命題：每 MW 經濟性（Tokenomics {TKV} 快照，基準成本情境）與定價")
TKSRC = f"coreweave/data/tokenomics_snapshot_{TKV}.json（Tokenomics {TKV}，{TKS['source']['commit'][:7]}）"
for g in GEN:
    inp(f"he_{g}", f"每 MW 年持有成本（經濟口徑）IF_HoldEcon｜{GNAME[g]}", "US$M／MW·年", tkv("IF_HoldEcon", g, "低成本"), tkv("IF_HoldEcon", g), tkv("IF_HoldEcon", g, "高成本"), "Derived", TKSRC + " Interface!13 列", "低／高＝Tokenomics 低成本／高成本情境（只作敏感度，鍵 he_all）")
for g in GEN:
    inp(f"cx_{g}", f"每 MW IT 資本支出 IF_CapexIT｜{GNAME[g]}", "US$M／MW", None, tkv("IF_CapexIT", g), None, "Derived", TKSRC + " Interface!9 列", "只含 IT；機房以租賃取得為主")
for g in GEN:
    inp(f"ox_{g}", f"每 MW 年營運費用（不含折舊、租金）IF_OpexGW｜{GNAME[g]}", "US$M／MW·年", None, tkv("IF_OpexGW", g), None, "Derived", TKSRC + " IF_OpexGW（電費、IT 與廠房維護、人員、稅與保險）", "CLAUDE.md 第 13 條：營運成本以 Tokenomics 滿載穩態為主值")
inp("fac_cx", "每 MW 自建機房資本支出 L1_FacCapexMW", "US$M／MW", None, round(TK["L1_FacCapexMW"]["values"] if not isinstance(TK["L1_FacCapexMW"]["values"], dict) else 12.67, 4), None, "Derived", TKSRC + " L1!D10", "只用於自建比例 own_share")
inp("life_ai", "AI 設備折舊年限（會計口徑）", "年", 5, 6, 7, "Derived", f"Tokenomics IF_DeprLifeIT｜GB300 基準 {TK['IF_DeprLifeIT']['values']['GB300 NVL72']['基準']:g}（高成本 {TK['IF_DeprLifeIT']['values']['GB300 NVL72']['高成本']:g}）；CoreWeave 伺服器耐用年限 6 年（10-K）", "會計年限≠GPU 經濟壽命；Sens 另有經濟汰換壓力 4 年案（Assumed，待 Tokenomics 登錄；原則第 10 節第 27 條）")
inp("fac_life", "自建機房折舊年限", "年", None, 20, None, "Assumed", "一般廠房 15–25 年（判斷）", "影響小（自建比例 15%）")
inp("p_long", "長約價格證據：IREN–Microsoft GB300 五年約（Analogy）", "US$M／MW·年", 8.9073, 9.70, 12.7247, "Analogy", "IREN 新聞稿 2025-11-03：$9.7bn ÷ 5 年 ÷ 200 MW（critical IT；全套 GPU 雲端服務；含 20% 預付）", "原則第 10 節第 61 條：他家合約換算一律標 Analogy；與 CoreWeave 差異：客戶（Microsoft 對 OpenAI／Meta 等）、世代同為 GB300、IREN 自有機房（含設施成本）、預付 20%；第二類比 IREN–NVIDIA 氣冷 Blackwell 11.33（同一供應商）。低／高沿用 v4.13 pricing.anchorMultiple")
inp("p_spot", "現貨價格證據：H100 Silicon Data 指數", "US$／GPU-hr", None, 2.82, None, "Verified", "Silicon Data H100 Index（2026-10-06；沿用 v4.13 pricing.marketRefs）", "—")
inp("gh_h100", "每 GPU 小時持有成本 IF_GPUhrEcon｜Hopper H100", "US$／GPU-hr", None, round(TK["IF_GPUhrEcon"]["values"]["Hopper H100"]["基準"], 5), None, "Derived", TKSRC, "—")
inp("od_share", "隨需（現貨）占計費產能比例", "%", 0.0, 0.10, 0.20, "Assumed", "公司未揭露長約／隨需比例（找不到；10-K：已承諾合約加權年期約 5 年）", "v4.13 基準 0%（10%、20% 只作敏感度）；v5.0 起基準改 10%，此變更未列入 v4.13 驗收採用規劃書（2026-11-13 稽核指出，基準值待 Andy 決定；回到 0% 見 Sens 低值）")
inp("k_long", "長約定價倍數 k_長約（＝長約價 ÷ GB300 持有成本）", "x", None, "=p_long/he_gb300", None, "Derived", "CLAUDE.md 第 14 條：每 MW 收入＝IF_HoldEcon × 公司因素 k；不以營收、ARR、RPO 反推 k", "—")
inp("k_spot", "現貨定價倍數 k_現貨", "x", None, "=p_spot/gh_h100", None, "Derived", "H100 現貨指數 ÷ H100 每 GPU 小時持有成本", "—")
inp("k_mix", "定價倍數 k（隨需占比加權；單一 k，CLAUDE.md 第 16(3) 條）", "x", None, "=od_share*k_spot+(1-od_share)*k_long", None, "Derived", "k＝隨需占比 × k_現貨＋（1 − 隨需占比）× k_長約", "J4：公司宣稱的新約價格不改基準 k")
inp("k_adj", "k 調整係數（敏感度與反解用；基準 1）", "x", 0.9, 1.0, 1.2, "Assumed", "敏感度 ×0.9／×1.2", "RevDCF 以本鍵反解支持現價所需的每 MW 收入")
inp("p_new_ovr", "新增產能每 MW 年收入改用公司宣稱值（0＝不用；J4 Sens）", "US$M／MW·年", None, 0, 22.5, "Assumed", "公司宣稱新約 4–5 年約每 MW 年收入 US$20–25m（Interested-party，K 節 p_claim_*）；高＝中點 22.5", "J4：只進 Sens、不改基準 k；作用於 2026-09-30 之後新增的在役 MW")
inp("lam", "成長型資本支出提前比例 λ（次年新增產能的設備在本年付款）", "%", 0.20, 0.35, 0.50, "Assumed", "沿用 v4.13 defaults.lambda", "可證偽指標：10-Q 在建工程 ÷ 次年新增產能資本支出")
inp("own_share", "自建機房比例（其餘租賃）", "%", None, 0.15, None, "Derived", "1 − 承租比例 85%（legacy_v4/company.json leases.facts.share）", "—")
inp("rent_mw", "租賃機房每 MW 年租金", "US$M／MW·年", 0.80, 1.00, 1.57, "Derived", "v4.13 站點檢查：FY2030 每 MW 年租金 1.00（Helios、Polaris 等租約總額 ÷ MW ÷ 年）", "高＝Q2 營業租賃費用 0.5 × 4 ÷（1,500 MW × 85%）＝1.57（含未通電站點的租金，A11 租金高案）；低＝0.80")
inp("sga_r", "管銷費用率（扣 SBC）", "%", 0.045, 0.0551, 0.065, "Derived", Q2Q + "：（銷售行銷 0.060 − SBC 0.012＋一般管理 0.178 − SBC 0.084）÷ 營收 2.575", "Q3 明細待 Q3 10-Q；全期持平（原則第 10 節第 35 條）")
inp("sbc_r", "股票薪酬 ÷ 營收", "%", 0.045, "=sbc9m/rev9m", 0.065, "Derived", PRQ3 + "：9M SBC 0.478 ÷ 9M 營收 8.314", "全期持平；DCF 以 SBC 為現金成本（不加回）")
inp("q4cx_ovr", "Q4 資本支出改用公司指引隱含值（0＝不用；J5 Sens）", "US$B", None, 0, 7.35, "Assumed", "高＝指引中點 38 − 9M 公司口徑 30.65（Interested-party 推得）", "J5：指引只進 Recon 與 Sens；基準為模型公式（原則第 10 節第 59 條）")
inp("pp_q4", "Q4 額外客戶預付（0＝不計；公司稱未來三個月約 9）", "US$B", None, 0, 9, "Assumed", PRQ3 + "：CEO 稱未來三個月約 90 億美元客戶預付與押金（Interested-party）", "只作 Sens（基準不計：預付對應的資本支出與認列時點未揭露）")

sec("E. 實績錨點（Q3 新聞稿；FY2026＝9M 實績＋Q4 模型）")
for k, fid, item in (("rev9m", "9m.rev", "營收（9M26）"), ("oi9m", "9m.oi", "營業利益（GAAP，9M26）"), ("aeb9m", "9m.aebitda", "調整後 EBITDA（9M26）"),
                     ("aoi9m", "9m.aoi", "調整後營業利益（9M26）"), ("ni9m", "9m.ni", "淨損（9M26）"), ("int9m", "9m.int", "利息費用淨額（9M26）"),
                     ("da9m", "9m.da", "折舊攤銷（9M26）"), ("sbc9m", "9m.sbc", "股票薪酬（9M26）"), ("cfo9m", "9m.cfo", "營運現金流（9M26）"),
                     ("cx9m_cash", "9m.cashcapex", "現金資本支出（購置 PP&E，9M26）"), ("rev_q3", "q3.rev", "營收（Q3 2026）"), ("aeb_q3", "q3.aebitda", "調整後 EBITDA（Q3 2026）"),
                     ("da_q3", "q3.da", "折舊攤銷（Q3 2026）"), ("int_q3", "q3.int", "利息費用淨額（Q3 2026）")):
    t = FX[fid]["tag"]
    inp(k, item, "US$B", None, fv(fid), None, t, PRQ3 + f"（{FACTS_F} {fid}）", FX[fid].get("note", "—"), INC_CO if t == "Interested-party" else "")
inp("cx_q3_def", "公司口徑資本支出（Q3 2026；含應付設備款、OEM 融資）", "US$B", None, fv("q3.capex_def"), None, "Interested-party", CALL + f"（{FACTS_F} q3.capex_def）", "新聞稿只有現金購置 9.441；待 Q3 10-Q 核對")
inp("cx_1h_def", "公司口徑資本支出（1H26）", "US$B", None, fv("1h.capex_def"), None, "Verified", Q2Q + "：PP&E 增加＋融資租賃 − CIP 變動（legacy_v4/company.json ytdActual.capex）", "與全年指引同口徑")
inp("cx9m_def", "公司口徑資本支出（9M26）＝1H 實績＋Q3", "US$B", None, "=cx_1h_def+cx_q3_def", None, "Derived", "16.139＋14.511（J5）", "FY2026 公司口徑＝本列＋Q4 模型")
inp("ppe_q3", "PP&E 淨額（2026-09-30）", "US$B", None, fv("q3.ppe"), None, "Verified", PRQ3, "Thesis 投入資本起點")
inp("cip_q2", "在建工程（2026-06-30；Q4 起分 6 年折舊）", "US$B", None, 11.9, None, "Verified", Q2Q + "（legacy_v4/company.json latestQuarter.cip）", "Q3 值待 Q3 10-Q")
inp("def_open", "遞延收入（客戶預付）餘額（2026-06-30）", "US$B", None, 9.692, None, "Verified", Q2Q + "（latestQuarter.deferredTotal）", "自 Q4 起依 prepay_yrs 年直線認列（非現金營收）；Q3 值待 Q3 10-Q")

sec("F. 現金流、融資")
inp("cash_q3", "現金與約當現金（2026-09-30）", "US$B", None, fv("q3.cash"), None, "Verified", PRQ3, "受限現金 2.282 另列，不計入可動用現金")
inp("rcash_q3", "受限現金（2026-09-30）", "US$B", None, fv("q3.rcash"), None, "Verified", PRQ3, "不計入淨負債的現金（保守）")
inp("debt_q3", "借款帳面值（追索＋非追索，2026-09-30）", "US$B", None, fv("q3.debt"), None, "Derived", PRQ3 + "：追索 9.183＋32.417、非追索 1.389＋3.884", "本金與到期表待 Q3 10-Q（Debt 頁沿用 Q2 10-Q 逐檔）")
inp("finl_q3", "融資租賃負債（2026-09-30）", "US$B", None, fv("q3.finlease"), None, "Verified", PRQ3, "列入淨負債與 WACC 負債")
inp("opl_q3", "營業租賃負債（2026-09-30）", "US$B", None, fv("q3.oplease"), None, "Verified", PRQ3, "租金已在 EBITDA 扣除，不列入淨負債")
inp("r_old", "既有借款平均利率（Q3 利息淨額 × 4 ÷ 借款）", "%", None, "=int_q3*4/debt_q3", None, "Derived", PRQ3 + "：0.846 × 4 ÷ 46.873", "公司稱加權債務成本 7.9%（Interested-party）與 Q4 利息指引隱含值並列於 Recon")
inp("kd", "稅前債務成本（2032 年優先票據發行票面利率，2026-06 發行）", "%", 0.08, 0.09625, 0.11, "Derived", "legacy_v4/company.json debt.instruments「2032 9.625% Senior Notes（2026-06 發行）」發行票面利率（原則第 10 節第 21 條）；現行市場殖利率：未取得（SEC、TRACE 報價在建置環境無法取得）", "與公司稱加權債務成本 7.9%（含可轉債低票息，Interested-party，In kd_co）並列於 Val；DCF 三情境皆失效取 0，kd 目前不影響目標價")
inp("r_new", "瀑布新增投資級債利率", "%", None, "=kd", None, "Derived", "同 kd", "—")
inp("r_junk", "高息債利率", "%", 0.15, 0.12, 0.18, "Assumed", "沿用 v4.13 defaults.junkRate 12%", "A12：Sens 15%／18%")
inp("min_cash", "最低現金", "US$B", None, 2, None, "Assumed", "沿用 v4.13 defaults.minCash", "—")
inp("lev_max", "投資級債上限：借款 ÷ 調整後 EBITDA（年化）", "x", 4.0, 4.5, 5.0, "Assumed", "原則第 10 節第 9 條預設 4.5", "只限制新增投資級債；超過走新股、再走高息債")
inp("icov_min", "利息保障倍數下限（調整後 EBITDA ÷ 利息；只 WARN）", "x", None, 3, None, "Assumed", "原則第 10 節第 9 條預設 3", "—")
inp("eq_cap", "每年新股上限（占市值）", "%", 0.10, 0.20, 0.30, "Assumed", "沿用 v4.13 defaults.eqCapPct 20%", "A12、原則第 62 條：Sens 10%／30%")
inp("eq_disc", "新股發行折價", "%", None, 0.10, None, "Assumed", "沿用 v4.13 defaults.eqDisc", "—")
inp("prepay_cov", "客戶預付覆蓋比（預付 ÷ 資本支出）", "%", 0.0, 0.08, 0.15, "Assumed", "沿用 v4.13 scenarios.capexTemplate.customerFund 8%", "—")
inp("prepay_yrs", "預付認列年數", "年", None, 5, None, "Assumed", "10-K：已承諾合約加權年期約 5 年", "—")
inp("nwc_r", "營運資金增加 ÷ 營收增量（預付除外）", "%", 0.0, 0.02, 0.04, "Assumed", "沿用 v4.13 valuation.wcPctOfRevGrowth 2%", "A12：Sens 0%／4%；DCF 標準規格：FCFF 扣營運資金增加")
inp("tax_lt", "長期現金稅率", "%", None, 0.21, 0.25, "Assumed", "美國聯邦 21%（B3：Assumed）", "Sens 25%（含州稅）")
inp("nol0", "淨營業虧損（NOL）餘額（2025 年底）", "US$B", None, 4.0, None, "Interested-party", "沿用 v4.13 valuation.nol（10-K 稅務附註，聯邦 NOL 約 4）", "Section 382 使用限制列風險登錄（Unknown）")
inp("nol_lim", "NOL 每年可抵稅前所得上限", "%", None, 0.8, None, "Verified", "美國稅法（TCJA 後 NOL 抵減上限 80%）", "B4")

sec("G. 評價")
inp("rf", "無風險利率（美國 10 年期公債，2026-11-10 收盤）", "%", None, fv("ust10.1110"), None, "Verified", SRCF["UST"], "11-11 退伍軍人節債市休市，取前一交易日收盤（與現價日相差 1 個交易日）；v5.0 初版誤用 Trading Economics 盤中 4.88%")
inp("beta_u", "無槓桿 β（自算產業 β：US 伺服器／EMS 代理）", "x", 1.022, 1.322, 1.622, "Analogy", "docs/plan/自算產業β_20261105.md：US 伺服器／EMS 同業 8 家中位數（Yahoo 月報酬 2021-11～2026-10，對 S&P 500 迴歸，Hamada 去槓桿、現金修正），Blume 調整後 1.322", "新雲沒有自算組（CoreWeave 上市不足 60 個月）；以 AI 伺服器硬體組代理（判斷題）；低／高＝±0.3（規則）")
inp("beta", "β（依市值 D/E 重新加槓桿：β_u ×（1＋(1 − 稅率) × D/E））", "x", None, "=beta_u*(1+(1-tax_lt)*(debt_q3+finl_q3)/(px*(sh_out+rsu_net)))", None, "Derived", "Hamada（自算產業 β 檔寫法）", "—")
inp("erp", "股權風險溢酬", "%", 0.04, 0.05, 0.06, "Assumed", "DCF 標準規格預設 5%（敏感度 4–6%）", "—")
inp("wacc_adj", "WACC 調整（敏感度用；基準 0）", "pt", -0.01, 0.0, 0.01, "Assumed", "Sens：±1 個百分點", "—")
inp("ronic_adj", "永續期 RONIC − WACC（基準 0）", "pt", None, 0.0, 0.05, "Assumed", "DCF 標準規格：RONIC＝WACC", "Sens：WACC＋5pt")
inp("tp_g", "永續成長率 g（＝MIN(3%, 10 年期)）", "%", 0.02, "=MIN(0.03,rf)", 0.04, "Derived", "A5：g＝MIN(3%, rf)", "—")
inp("w_dcf", "目標價中 DCF 的權重（其餘 Comparable）", "%", 0.3, 0.5, 0.7, "Assumed", "A1：DCF 50%＋Comparable 50%（取代 v4.13 的 45%）", "Andy 可改")
for s, w in (("A", 0.25), ("B", 0.5), ("C", 0.25)):
    inp(f"w{s}", f"目標價情境權重 {s}", "%", None, w, None, "Assumed", "A14：25／50／25", "Andy 可改")
PEER_D = "2026-11-11"
inp("pn_px", "Nebius 股價（2026-11-11 收盤）", "US$", None, 154.69, None, "Verified", "MarketScreener 共識，擷取 2026-11-13（URL 見 Sources MS；chat 端內建瀏覽器擷取，頁面「Market Closed 2026-11-12」）：Nasdaq 5 日表 11-11 154.69（11-10 134.00、11-12 128.64）", "同業估值日＝現價日 2026-11-11（原則 7、15）")
inp("pn_sh", "Nebius 普通股流通在外（2026-09-30；A 股 259.597＋B 股 33.305 百萬股）", "十億股", None, 0.292902405, None, "Verified", "Nebius Q3 2026 6-K 附件 99.1（2026-11-12；URL 見 Sources NBIS_6K）：292,902,405 股（不含庫藏 29.1 百萬股）", "口徑：基本股數配含可轉債（列為負債）的淨負債；MarketScreener 274,099 千股與 Nebius v0.5.3 評價股數 0.3193（含可轉債轉股）均未採用")
inp("pn_nd", "Nebius 淨負債：最新申報（2026-09-30；借款含可轉債 − 現金；受限現金不扣）", "US$B", None, "=14.1908-7.4256", None, "Derived", "Nebius Q3 2026 6-K 附件 99.1（2026-11-12；URL 見 Sources NBIS_6K）：借款 流動 0.1875＋非流動 14.0033＝14.1908；現金與約當現金 7.4256；受限現金 1.2155 不扣（與 CoreWeave 自身口徑一致）", "融資租賃負債：找不到（6-K 未列）；計入只會提高 Nebius 倍數，Nebius 已高於中位數，不影響中位數。10 月 Citi 額度為期後")
inp("pn_ndf", "Nebius 淨負債：MarketScreener 2026 年預估", "US$B", None, 13.659, None, "Interested-party", "MarketScreener 共識，擷取 2026-11-13（URL 見 Sources MS；chat 端內建瀏覽器擷取，頁面「Market Closed 2026-11-12」）", "替代口徑（Val 並列、Sens peer_nd_mode）", INC_SELL)
inp("pn_eb", "Nebius NTM 共識 EBITDA（2026 × 1.6/12＋2027 × 10.4/12；12 月制）", "US$B", None, "=1.348*1.6/12+6.556*10.4/12", None, "Interested-party", "MarketScreener 共識，擷取 2026-11-13（URL 見 Sources MS；chat 端內建瀏覽器擷取，頁面「Market Closed 2026-11-12」）：2026 1,348、2027 6,556、2028 13,700（US$M）", "原則 49：自 2026-11-11 起月份加權", INC_SELL)
inp("ir_px", "IREN 股價（2026-11-11 收盤）", "US$", None, 20.84, None, "Verified", "MarketScreener 共識，擷取 2026-11-13（URL 見 Sources MS；chat 端內建瀏覽器擷取，頁面「Market Closed 2026-11-12」）：Nasdaq 5 日表 11-11 20.84（11-10 19.73、11-12 19.48）", "—")
inp("ir_sh", "IREN 股數（Nbr of stocks）", "十億股", None, 0.424357, None, "Derived", "MarketScreener 共識，擷取 2026-11-13（URL 見 Sources MS；chat 端內建瀏覽器擷取，頁面「Market Closed 2026-11-12」）：424,357 千股", "口徑：基本股數配含可轉債的淨負債（10-Q 股數：找不到，用 MarketScreener）")
inp("ir_nd", "IREN 淨負債：最新申報（2026-09-30；借款＋融資租賃 − 現金；受限現金不扣）", "US$B", None, "=10.208-3.522", None, "Derived", "IREN Q1 FY27 新聞稿（2026-11-05；URL 見 Sources IREN_Q1）：GAAP 借款含融資租賃 10.208（公司自列）；現金與約當現金 3.522；受限現金 2.256 不扣（與 CoreWeave 自身口徑一致；公司自算淨負債 4.430 有扣受限現金）", "替代口徑見 ir_ndf")
inp("ir_ndf", "IREN 淨負債：MarketScreener FY2027 預估", "US$B", None, 18.889, None, "Interested-party", "MarketScreener 共識，擷取 2026-11-13（URL 見 Sources MS；chat 端內建瀏覽器擷取，頁面「Market Closed 2026-11-12」）", "替代口徑", INC_SELL)
inp("ir_eb", "IREN NTM 共識 EBITDA（FY27 × 7.6/12＋FY28 × 4.4/12；6 月制）", "US$B", None, "=1.395*7.6/12+5.123*4.4/12", None, "Interested-party", "MarketScreener 共識，擷取 2026-11-13（URL 見 Sources MS；chat 端內建瀏覽器擷取，頁面「Market Closed 2026-11-12」）：FY2027 1,395、FY2028 5,123、FY2029 9,826（US$M）", "原則 49 月份加權；v5.0 初版曾用 chartmill（FY27 2.21），兩來源差異大，改用與另兩家同源的 MarketScreener", INC_SELL)
inp("ir_use", "IREN 進同業中位數（1＝是；拆解用）", "—", None, 1, None, "Assumed", "2026-11-13 起取得同日股價、股數與共識", "—")
inp("po_px", "Oracle 股價（2026-11-11 收盤）", "US$", None, 128.38, None, "Verified", "MarketScreener 共識，擷取 2026-11-13（URL 見 Sources MS；chat 端內建瀏覽器擷取，頁面「Market Closed 2026-11-12」）：NYSE 5 日表 11-11 128.38（11-10 117.98、11-12 127.11）", "—")
inp("po_sh", "Oracle 股數（Nbr of stocks）", "十億股", None, 3.117708, None, "Derived", "MarketScreener 共識，擷取 2026-11-13（URL 見 Sources MS；chat 端內建瀏覽器擷取，頁面「Market Closed 2026-11-12」）：3,117,708 千股（含 2026-11-04 新股）", "—")
inp("po_nd", "Oracle 淨負債：最新實際（含融資租賃；2026-11 增資 pro forma）", "US$B", None, 87.208, None, "Derived", "Oracle v0.5.1 Val!C10", "主口徑；MarketScreener FY2026 實際 97.647（增資前）")
inp("po_ndf", "Oracle 淨負債：MarketScreener FY2027 預估", "US$B", None, 115.597, None, "Interested-party", "MarketScreener 共識，擷取 2026-11-13（URL 見 Sources MS；chat 端內建瀏覽器擷取，頁面「Market Closed 2026-11-12」）", "替代口徑", INC_SELL)
inp("po_eb", "Oracle NTM 共識 EBITDA（FY27 × 6.6/12＋FY28 × 5.4/12；5 月制）", "US$B", None, "=52.09*6.6/12+74.148*5.4/12", None, "Interested-party", "MarketScreener 共識，擷取 2026-11-13（URL 見 Sources MS；chat 端內建瀏覽器擷取，頁面「Market Closed 2026-11-12」）：FY2027 52,090、FY2028 74,148（US$M）", "原則 49 月份加權", INC_SELL)
inp("peer_nd_mode", "同業淨負債口徑（0＝最新實際；1＝MarketScreener 當年度預估）", "—", None, 0, 1, "Assumed", "主口徑同 Oracle 樣板（EV 以估值日的資產負債計）", "Sens：改用預估口徑（IREN 差異最大：1.752 → 18.889）")
inp("pn_rev", "Nebius NTM 共識營收（FY26 × 2/12＋FY27 × 10/12）", "US$B", None, "=3.379*2/12+11.43*10/12", None, "Interested-party", "nebius/data/consensus_nbis_20261007.json（FY27 營收由 EBITDA 6.194 ÷ EBITDA 率 54.2% 推得）", "同業表成長與利潤率欄", INC_SELL)
inp("po_rev", "Oracle NTM 共識營收（FY27 × 7/12＋FY28 × 5/12）", "US$B", None, "=90.423*7/12+131.711*5/12", None, "Interested-party", "oracle/data/consensus_orcl_20261022.json", "同業表成長與利潤率欄", INC_SELL)
inp("po_cx", "Oracle NTM 共識資本支出（FY27 × 7/12＋FY28 × 5/12）", "US$B", None, "=93.961*7/12+100.962*5/12", None, "Interested-party", "oracle/In cx2027、cx2028（MarketScreener）", "同業表資本支出率欄", INC_SELL)
inp("rel_mult", "相對倍數調整係數", "x", None, 1.0, None, "Assumed", "1.00（未調整：缺 CoreWeave 10 年歷史倍數，上市不足 2 年；tools/relative_multiple.py 需要 10 年比值）", "原則第 10 節第 18、47 條：列入承重輸入")
inp("mult_adj", "Comparable 倍數調整（敏感度用；基準 1）", "x", 0.8, 1.0, 1.2, "Assumed", "敏感度 ±20%", "—")

sec("K. 共識、公司指引與對照（只作 Recon 與 Sens，不進模型）")
GR = lambda fid, i: fv(fid)[i]
for k, fid, i, item in (("g_rev_lo", "g26.rev", 0, "公司 FY2026 營收指引：下緣"), ("g_rev_hi", "g26.rev", 1, "公司 FY2026 營收指引：上緣"),
                        ("g_aoi_lo", "g26.aoi", 0, "公司 FY2026 調整後營業利益指引：下緣"), ("g_aoi_hi", "g26.aoi", 1, "公司 FY2026 調整後營業利益指引：上緣"),
                        ("g_cx_lo", "g26.capex", 0, "公司 FY2026 資本支出指引：下緣"), ("g_cx_hi", "g26.capex", 1, "公司 FY2026 資本支出指引：上緣"),
                        ("g_arr_lo", "g26.arr", 0, "公司 2026 年底 ARR 指引：下緣"), ("g_arr_hi", "g26.arr", 1, "公司 2026 年底 ARR 指引：上緣"),
                        ("g_int_lo", "g26.q4int", 0, "公司 Q4 2026 利息費用指引：下緣"), ("g_int_hi", "g26.q4int", 1, "公司 Q4 2026 利息費用指引：上緣")):
    inp(k, item, "US$B", None, GR(fid, i), None, "Interested-party", PRQ3 + "（Outlook 段）", "只作對照，不進模型", INC_GUIDE)
inp("kd_co", "公司稱加權債務成本", "%", None, fv("q3.kd_co"), None, "Interested-party", CALL, "只作對照，不進模型", INC_CO)
inp("backlog", "Revenue backlog（2026-09-30）", "US$B", None, fv("q3.backlog"), None, "Interested-party", PRQ3, "只作對照，不進模型", INC_CO)
inp("rpo", "RPO（2026-09-30）", "US$B", None, fv("q3.rpo"), None, "Interested-party", CALL, "只作對照，不進模型", INC_CO)
for k, v, lab in (("rpo_24", 0.36, "24 個月內"), ("rpo_48", 0.38, "25–48 個月"), ("rpo_l", 0.26, "48 個月後")):
    inp(k, f"RPO 認列時點：{lab}", "%", None, v, None, "Interested-party", CALL, "只作對照，不進模型", INC_CO)
inp("p_claim_lo", "公司宣稱：新簽 4–5 年約每 MW 年收入（下緣）", "US$M／MW·年", None, GR("claim.permw", 0), None, "Interested-party", CALL + "；分母為 IT 或設施 MW 未說明（待查）", "只作對照，不進模型（J4）", INC_CO)
inp("p_claim_hi", "公司宣稱：新簽 4–5 年約每 MW 年收入（上緣）", "US$M／MW·年", None, GR("claim.permw", 1), None, "Interested-party", CALL, "只作對照，不進模型（J4）", INC_CO)
inp("p_short", "公司宣稱：短天期（3–6 個月）合約每 MW 年收入", "US$M／MW·年", None, fv("claim.short"), None, "Interested-party", SRCF["PR_0917"] + "：年化收入 ÷ 服務該叢集所需電力", "只作對照，不進模型（J4）", INC_CO)
inp("cons_eb27", "共識 FY2027 EBITDA（2026-09-25，財報前）", "US$B", None, 16.656, None, "Interested-party", "coreweave/data/consensus_crwv_20260925.json（MarketScreener，S&P 系）", "只作對照，不進模型（財報前，只作歷史對照）", INC_SELL)
inp("cons_eb26", "共識 FY2026 EBITDA（2026-09-25，財報前）", "US$B", None, 7.574, None, "Interested-party", "同上", "只作對照，不進模型（財報前，只作歷史對照）", INC_SELL)
inp("cons_tp", "市場共識目標價（2026-09-25，40 家平均，財報前）", "US$", None, 140.81, None, "Interested-party", "同上 priceTarget（S&P 經 StockAnalysis）", "只作對照，不進模型（財報前，只作歷史對照）", INC_SELL)
inp("tp_v413", "v4.13 加權目標價（舊法、現價 43.66）", "US$", None, 25.51, None, "Derived", "coreweave/dist/_待刪除/20261106_CoreWeave收支模型_v4_13.xlsx（基準 5.6 GW）", "只作對照，不進模型")
inp("dcf_v413", "v4.13 融資調整 DCF（舊法，不進結論）", "US$", None, 25.51, None, "Derived", "v4.13 基準情境（DCF 45%＋EV/EBITDA 55% 的加權結果，融資後股數）", "只作對照，不進模型（B1：舊法，不進結論）")
inp("wat26", "Wattonomics：CoreWeave 2026 年底上線 MW", "MW", None, 1850, None, "Derived", "wattonomics/data/wattonomics_snapshot_v0.3.json 列 mw_crwv（2026）", "只作對照，不進模型（三層規則：個股不改 Wattonomics）")
inp("wat30", "Wattonomics：CoreWeave 2030 年底上線 MW", "MW", None, 8000, None, "Derived", "同上 mw_crwv（2030；公司目標最多 8 GW）", "只作對照，不進模型")

base_vals, IR = {}, {}
r = 4
for srow in rows:
    if srow[0] == "sec":
        IN.cell(r, 1, srow[1]).font = BOLD; r += 1; continue
    key, item, unit, lo, base, hi, tag, s1, s2, inc = srow
    if key in OVR:
        base = OVR[key]
    base_vals[key] = (lo, base, hi)
    IN.cell(r, 1, key); IN.cell(r, 2, item); IN.cell(r, 3, unit)
    for j, v in ((4, lo), (5, base), (6, hi)):
        if v is not None:
            c = IN.cell(r, j, v)
            if not (isinstance(v, str) and v.startswith("=")): c.font = BLUE
            if unit in ("%", "pt"): c.number_format = "0.00%"
            elif unit == "x": c.number_format = "0.000"
            elif unit == "日期": c.number_format = "yyyy-mm-dd"
    IN.cell(r, 7, tag); IN.cell(r, 8, s1); IN.cell(r, 9, s2); IN.cell(r, 10, inc)
    dn(key, f"In!$E${r}")
    if lo is not None: dn(key + "_lo", f"In!$D${r}")
    if hi is not None: dn(key + "_hi", f"In!$F${r}")
    if key in ("w_dcf", "wA", "wB", "wC"):
        for c in range(1, 11): IN.cell(r, c).fill = KEY
    IR[key] = r
    r += 1
IN_LAST = r - 1
r += 1
IN.cell(r, 1, "L. 情境驅動對照（三情境並排，進主模型的加權目標價；本節只引用上方各列，修改請改上方藍字。上方「敏感度低／高」欄只用於 Sens 頁，不進目標價）").font = BOLD; r += 1
hdr(IN, r, ["鍵", "項目", "單位", "情境 A（樂觀）", "情境 B（基準）", "情境 C（保守）", "標記", "說明"]); r += 1
for key, item, unit, ks, fmt, note in (
        ("scn_mw27", "2027 年底在役 MW", "MW", ("mwA_27", "mwB_27", "mwC_27"), "#,##0", "沿用 v4.13 容量路徑"),
        ("scn_mw28", "2028 年底在役 MW", "MW", ("mwA_28", "mwB_28", "mwC_28"), "#,##0", "—"),
        ("scn_mw29", "2029 年底在役 MW", "MW", ("mwA_29", "mwB_29", "mwC_29"), "#,##0", "C＝已簽約 4.6 GW（J2）"),
        ("scn_mw30", "2030 年底在役 MW", "MW", ("mwA_30", "mwB_30", "mwC_30"), "#,##0", "A 8 GW＝公司 2030 目標上緣"),
        ("scn_post", "2031–2032 每年新增 MW", "MW／年", ("addA_post", "addB_post", "addC_post"), "#,##0", "—"),
        ("scn_w", "情境權重", "%", ("wA", "wB", "wC"), "0%", "[判斷] 黃底，Andy 可改")):
    IN.cell(r, 1, key); IN.cell(r, 2, item); IN.cell(r, 3, unit)
    for j, k_ in zip((4, 5, 6), ks):
        IN.cell(r, j, f"=E{IR[k_]}").number_format = fmt
    IN.cell(r, 7, "Derived"); IN.cell(r, 8, "引用 In 頁 " + "／".join(ks) + "；" + note)
    r += 1
for col, w in zip("ABCDEFGHIJ", [13, 58, 12, 10, 11, 10, 16, 70, 60, 34]): IN.column_dimensions[col].width = w
IN.freeze_panes = "D4"

# =====================================================================================
# Q：實績
# =====================================================================================
Q = wb.create_sheet("Q")
Q["A1"] = "Q — 實績（十億美元）。會計年度＝日曆年。藍字＝輸入；黑字＝公式。Q3 2026＝公司新聞稿（8-K 附件 99.1，2026-11-10）；Q1、Q2＝10-Q（Q1 由 1H 減 Q2 推得者標 Derived）；FY2024–FY2025＝10-K"
Q["A2"] = f"資料取自 {FACTS_F}（Q3）與 coreweave/legacy_v4/company.json（Q2 10-Q、10-K）。官方報告分部：CoreWeave 只有一個應報導部門（AI 雲端平台），無其他營收線揭露"
QH = ["項目", "Q1 2026", "Q2 2026", "Q3 2026", "9M 2026", "FY2024", "FY2025", "標記", "來源", "說明", "來源誘因"]
hdr(Q, 3, QH)
QD = [("營收", "rev", (2.078, 2.575, 3.661, 8.314, 1.915, 5.131), "Verified", ""),
      ("營業利益（GAAP）", "oi", ("=E5-C5-D5", -0.049, 0.207, 0.014, 0.324, -0.046), "Verified", "Q1＝9M − Q2 − Q3（Derived 公式）"),
      ("調整後 EBITDA（非 GAAP）", "aeb", (1.157, 1.510, 2.256, 4.923, None, None), "Interested-party", "Q1、Q2 引自財報報導轉述（v4.13 ytdActual.adjEbitdaMeta）"),
      ("調整後營業利益（非 GAAP）", "aoi", ("=E7-C7-D7", 0.128, 0.379, 0.528, None, None), "Interested-party", "Q1＝9M − Q2 − Q3"),
      ("淨損（GAAP）", "ni", ("=E8-C8-D8", -0.626, -0.709, -2.075, -0.863, -1.167), "Verified", ""),
      ("利息費用淨額", "int", ("=E9-C9-D9", 0.640, 0.846, 2.022, None, None), "Verified", ""),
      ("折舊攤銷", "da", ("=E10-C10-D10", 1.393, 1.889, 4.429, None, None), "Verified", ""),
      ("股票薪酬", "sbc", ("=E11-C11-D11", 0.165, 0.160, 0.478, None, None), "Verified", ""),
      ("營運現金流", "cfo", (None, None, 1.975, 5.638, None, None), "Verified", "1H 3.663（Q1、Q2 個別：未取得）"),
      ("現金資本支出（購置 PP&E）", "cx", (None, 9.352, 9.441, 23.558, None, None), "Verified", "1H 14.117；Q2 欄為公司口徑 9.352（見說明）"),
      ("公司口徑資本支出", "cxd", (None, None, 14.511, "=16.139+D14", None, None), "Interested-party", "Q3 14.511 為法說數（新聞稿未見，待 Q3 10-Q）；1H 16.139 為 10-Q")]
QR = {}
for i, (lab, k, vals, tag, note) in enumerate(QD):
    rr = 4 + i; QR[k] = rr
    Q.cell(rr, 1, lab)
    for j, v in enumerate(vals):
        if v is None: continue
        c = Q.cell(rr, 2 + j, v); c.number_format = "#,##0.000"
        if not (isinstance(v, str) and v.startswith("=")): c.font = BLUE
    Q.cell(rr, 8, tag); Q.cell(rr, 9, PRQ3 + "；" + Q2Q + "；" + SRCF["K10_25"]); Q.cell(rr, 10, note or "—")
    Q.cell(rr, 11, INC_CO if tag == "Interested-party" else "")
rr = 4 + len(QD)
Q.cell(rr, 1, "檢查：Q1＋Q2＋Q3 營收 − 9M 營收（應為 0）"); Q.cell(rr, 5, "=B4+C4+D4-E4").number_format = "0.000"; Q.cell(rr, 8, "—"); QR["chk"] = rr
Q.cell(rr + 1, 1, "對照：9M 調整後 EBITDA − 營業利益 − 股票薪酬 − 折舊攤銷（其他調整）"); Q.cell(rr + 1, 5, "=E6-E5-E11-E10").number_format = "0.000"; Q.cell(rr + 1, 8, "—")
Q.cell(rr + 2, 1, "季增在役 MW ÷ 季營收年化（每 MW 年營收，Q3：營收 × 4 ÷ 平均在役 MW）"); Q.cell(rr + 2, 4, "=D4*4/((mw_q2+mw_q3)/2)*1000").number_format = "0.00"; Q.cell(rr + 2, 8, "Derived"); QR["permw"] = rr + 2
Q.cell(rr + 2, 10, "公司實況驗證（CLAUDE.md 第 15 條）：與 Thesis 每 MW 年收入 × 計費比例對照，見 Recon")
Q.column_dimensions["A"].width = 58
for c in "BCDEFG": Q.column_dimensions[c].width = 11
Q.column_dimensions["H"].width = 16; Q.column_dimensions["I"].width = 60; Q.column_dimensions["J"].width = 50; Q.column_dimensions["K"].width = 30
Q.freeze_panes = "B4"

# =====================================================================================
# Thesis：命題模組（AI 雲端：在役 MW × 每 MW 年收入）
# =====================================================================================
TH = wb.create_sheet("Thesis")
TH["A1"] = "Thesis — 命題模組：CoreWeave 已簽約的 backlog 能否以高於資金成本的報酬轉成營收與現金，同時不讓資金缺口失控。驅動唯一的報告分部（AI 雲端平台）"
TH["A2"] = "在役 MW（期末）依情境路徑；計費 MW＝平均在役 × 計費比例；每 MW 年收入＝在役世代加權 IF_HoldEcon × k；營運費用＝平均在役 × IF_OpexGW；租金＝平均在役 × 租賃比例 × 每 MW 年租金；資本支出＝新增 MW × IF_CapexIT（λ 提前付款）＋自建機房＋汰換；折舊＝Q3 年化基數 − 汰換批次＋在建工程與新資本支出 ÷ 年限（當年半年）"
hdr(TH, 3, ["項目", "單位", "9M 2026（實績）", "Q4 2026（模型）", "FY2026", "FY2027", "FY2028", "FY2029", "FY2030", "FY2031", "FY2032", "說明"])
TL = {"C": 0.75, "D": 0.25, "E": 1, **{TCOL[y]: 1 for y in FYM}}
FC = [TCOL[y] for y in FYM]                    # F..K
PREV = {"D": "C", "F": "E", "G": "F", "H": "G", "I": "H", "J": "I", "K": "J"}
TR = {}
r = 4
def put(ws, row, lab, unit, vals, fmt="#,##0.000", note=None, ncol=12):
    ws.cell(row, 1, lab); ws.cell(row, 2, unit)
    for c, v in vals.items():
        if v is None: continue
        x = ws[f"{c}{row}"]; x.value = v; x.number_format = fmt
    if note: ws.cell(row, ncol, note)
CXN = {"D": "(nm26_gb300*cx_gb300+(1-nm26_gb300)*cx_gb200)", "F": "(nm27_gb300*cx_gb300+(1-nm27_gb300)*cx_vr200)"}
for c in "GHIJK": CXN[c] = "(nm28_gb300*cx_gb300+(1-nm28_gb300)*cx_vr200)"
NXT = {"D": "F", "F": "G", "G": "H", "H": "I", "I": "J", "J": "K", "K": "K"}
RETC = {"H": "ret29", "I": "ret30", "J": "ret31", "K": "ret32"}
for s in SC:
    TH.cell(r, 1, SCN[s]).font = BOLD; r += 1
    R = {}
    def nr(k):
        global r
        R[k] = r; r += 1; return R[k]
    put(TH, nr("L"), "期間長度", "年", TL, "0.00")
    rc = nr("conn")
    put(TH, rc, "在役 MW（期末）", "MW", {"C": "=mw_q3", "D": "=mw_ye26", "E": f"=D{rc}", "F": f"=mw{s}_27", "G": f"=mw{s}_28", "H": f"=mw{s}_29", "I": f"=mw{s}_30",
                                          "J": f"=I{rc}+add{s}_post", "K": f"=J{rc}+add{s}_post"}, "#,##0", "情境路徑（In 頁 B 節）；口徑：公司 active power，視為 IT（Assumed）")
    ra = nr("add")
    put(TH, ra, "期間新增在役 MW", "MW", {"C": f"=C{rc}-mw_ye25", "D": f"=D{rc}-C{rc}", "E": f"=E{rc}-mw_ye25", **{c: f"={c}{rc}-{PREV[c]}{rc}" for c in FC}}, "#,##0")
    rt = nr("ret")
    put(TH, rt, "汰換 MW（期間；換成 VR200，在役量不變）", "MW", {"C": "=0", "D": "=0", "E": "=0", "F": "=0", "G": "=0", **{c: f"={RETC[c]}" for c in "HIJK"}}, "#,##0", "經濟壽命 6 年：各年新增批次在第 6 年汰換")
    rv = nr("avg")
    put(TH, rv, "平均在役 MW", "MW", {"C": f"=(mw_ye25+C{rc})/2", "D": f"=(C{rc}+D{rc})/2", "E": f"=(C{rv}*0.75+D{rv}*0.25)", **{c: f"=({PREV[c]}{rc}+{c}{rc})/2" for c in FC}}, "#,##0")
    rb = nr("bill")
    put(TH, rb, "計費 MW（× 計費比例）", "MW", {c: f"={c}{rv}*u_bill" for c in "CDEFGHIJK"}, "#,##0")
    TH.cell(r, 1, "世代存量（期末，MW）：2026-06-30 依 In 頁占比；新增依新增組合；汰換換成 VR200").font = GREY; r += 1
    for g in GEN: R[f"st_{g}"] = r + GEN.index(g)
    stv = {}
    for g in GEN:
        rr_ = R[f"st_{g}"]
        if g == "h100":
            v = {"C": "=mw_q2*om_h100", "D": f"=C{rr_}", "E": f"=D{rr_}", "F": f"=E{rr_}", "G": f"=F{rr_}", "H": f"=G{rr_}-ret29", "I": f"=H{rr_}-ret30", "J": f"=I{rr_}-ret31*v25_h100", "K": f"=J{rr_}"}
        elif g == "gb200":
            v = {"C": "=mw_q2*om_gb200+(mw_q3-mw_q2)*(1-nm26_gb300)", "D": f"=C{rr_}+D{ra}*(1-nm26_gb300)", "E": f"=D{rr_}", "F": f"=E{rr_}",
                 "G": f"=F{rr_}", "H": f"=G{rr_}", "I": f"=H{rr_}", "J": f"=I{rr_}-ret31*v25_gb200", "K": f"=J{rr_}-ret32*(1-v26_gb300)"}
        elif g == "gb300":
            v = {"C": "=mw_q2*om_gb300+(mw_q3-mw_q2)*nm26_gb300", "D": f"=C{rr_}+D{ra}*nm26_gb300", "E": f"=D{rr_}", "F": f"=E{rr_}+F{ra}*nm27_gb300",
                 **{c: f"={PREV[c]}{rr_}+{c}{ra}*nm28_gb300" for c in "GHI"}, "J": f"=I{rr_}+J{ra}*nm28_gb300-ret31*(1-v25_h100-v25_gb200)", "K": f"=J{rr_}+K{ra}*nm28_gb300-ret32*v26_gb300"}
        else:
            v = {"C": "=0", "D": "=0", "E": "=0", "F": f"=E{rr_}+F{ra}*(1-nm27_gb300)", **{c: f"={PREV[c]}{rr_}+{c}{ra}*(1-nm28_gb300)+{c}{rt}" for c in "GHIJK"}}
        put(TH, rr_, f"  {GNAME[g]}", "MW", v, "#,##0")
    r += 4
    put(TH, nr("stchk"), "檢查：世代存量合計 − 在役 MW（應為 0）", "MW", {c: f"=SUM({c}{R['st_h100']}:{c}{R['st_vr200']})-{c}{rc}" for c in "CDEFGHIJK"}, "0.000")
    def mix(c, pref):
        if c == "C":
            return "(" + "+".join(f"C{R['st_' + g]}*{pref}_{g}" for g in GEN) + f")/C{rc}"
        p = PREV[c]
        return "(" + "+".join(f"({p}{R['st_' + g]}+{c}{R['st_' + g]})*{pref}_{g}" for g in GEN) + f")/({p}{rc}+{c}{rc})"
    rh = nr("anc")
    put(TH, rh, "每 MW 錨：在役世代加權 IF_HoldEcon", "US$M／MW·年", {**{c: "=" + mix(c, "he") for c in "CD"}, "E": f"=(C{rh}*C{rb}*0.75+D{rh}*D{rb}*0.25)/(C{rb}*0.75+D{rb}*0.25)", **{c: "=" + mix(c, "he") for c in FC}}, "0.00", f"期初＋期末存量平均加權（Tokenomics {TKV} 基準成本）")
    rpm = nr("pm")
    put(TH, rpm, "每 MW 年收入（錨 × k × k 調整）", "US$M／MW·年", {c: f"={c}{rh}*k_mix*k_adj" for c in "CDEFGHIJK"}, "0.00", "k：In 頁 k_mix（IREN 長約 Analogy＋H100 現貨）")
    rpe = nr("pme")
    put(TH, rpe, "每 MW 年收入（實際採用；p_new_ovr＞0 時新增 MW 改用公司宣稱值）", "US$M／MW·年",
        {c: f"=IF(p_new_ovr>0,(MIN({c}{rv},mw_q3)*{c}{rpm}+MAX(0,{c}{rv}-mw_q3)*p_new_ovr)/{c}{rv},{c}{rpm})" for c in "DEFGHIJK"}, "0.00", "基準 p_new_ovr＝0（J4：公司宣稱只進 Sens）")
    rr_ = nr("rev")
    put(TH, rr_, "AI 雲端營收（命題）", "US$B", {"C": "=rev9m", "D": f"=D{rb}*D{rpe}*D{R['L']}/1000", "E": f"=C{rr_}+D{rr_}", **{c: f"={c}{rb}*{c}{rpe}/1000" for c in FC}}, "#,##0.000", "9M＝實績（新聞稿）；其後＝計費 MW × 每 MW 年收入")
    put(TH, nr("revchk"), "  對照：9M 模型估計營收（計費 MW × 每 MW 年收入 × 0.75；只作公司實況驗證）", "US$B", {"C": f"=C{rb}*C{rpm}*0.75/1000"}, "#,##0.000", "與實績差距見 Recon C 節（CLAUDE.md 第 15 條）")
    rox = nr("oxm")
    put(TH, rox, "每 MW 年營運費用：在役世代加權 IF_OpexGW", "US$M／MW·年", {**{c: "=" + mix(c, "ox") for c in "D"}, **{c: "=" + mix(c, "ox") for c in FC}}, "0.000")
    rop = nr("opex")
    put(TH, rop, "營運費用（電費、維護、人員、稅與保險；不含折舊、租金）", "US$B", {"D": f"=D{rv}*D{rox}*0.25/1000", **{c: f"={c}{rv}*{c}{rox}/1000" for c in FC}}, "#,##0.000", "平均在役 MW × 每 MW 營運費用")
    rrt = nr("rent")
    put(TH, rrt, "機房租金（租賃比例 × 每 MW 年租金）", "US$B", {"D": f"=D{rv}*(1-own_share)*rent_mw*0.25/1000", **{c: f"={c}{rv}*(1-own_share)*rent_mw/1000" for c in FC}}, "#,##0.000")
    rsg = nr("sga")
    put(TH, rsg, "管銷費用（扣 SBC）", "US$B", {"D": f"=D{rr_}*sga_r", **{c: f"={c}{rr_}*sga_r" for c in FC}}, "#,##0.000")
    re_ = nr("ebitda")
    put(TH, re_, "調整後 EBITDA（扣 SBC 前；公司口徑）", "US$B", {"C": "=aeb9m", "D": f"=D{rr_}-D{rop}-D{rrt}-D{rsg}", "E": f"=C{re_}+D{re_}", **{c: f"={c}{rr_}-{c}{rop}-{c}{rrt}-{c}{rsg}" for c in FC}}, "#,##0.000", "9M＝實績（非 GAAP，Interested-party）")
    put(TH, nr("ebm"), "  調整後 EBITDA 率", "%", {c: f"=IF({c}{rr_}=0,0,{c}{re_}/{c}{rr_})" for c in "CDEFGHIJK"}, "0.0%")
    rcg = nr("cxg")
    put(TH, rcg, "成長型 IT 資本支出（新增 MW × IF_CapexIT，λ 提前付款）", "US$B",
        {"D": f"=(D{ra}*(1-lam)*{CXN['D']}+F{ra}*lam*{CXN['F']}*0.25)/1000", **{c: f"=({c}{ra}*(1-lam)*{CXN[c]}+{NXT[c]}{ra}*lam*{CXN[NXT[c]]})/1000" for c in FC}}, "#,##0.000",
        "Q4：Q4 新增 ×（1 − λ）＋ 2027 年新增 × λ × 0.25；FY2032 的次年新增假設與當年相同")
    rcf = nr("cxf")
    put(TH, rcf, "自建機房資本支出（新增 MW × 自建比例 × L1_FacCapexMW）", "US$B", {c: f"={c}{ra}*own_share*fac_cx/1000" for c in "D" + "".join(FC)}, "#,##0.000")
    rcr = nr("cxr")
    put(TH, rcr, "汰換資本支出（汰換 MW × VR200 IF_CapexIT）", "US$B", {"D": "=0", "F": "=0", "G": "=0", **{c: f"={c}{rt}*cx_vr200/1000" for c in "HIJK"}}, "#,##0.000")
    rcx = nr("cx")
    put(TH, rcx, "資本支出合計（公司口徑；Q4 起模型）", "US$B", {"C": "=cx9m_def", "D": f"=IF(q4cx_ovr>0,q4cx_ovr,D{rcg}+D{rcf}+D{rcr})", "E": f"=C{rcx}+D{rcx}", **{c: f"={c}{rcg}+{c}{rcf}+{c}{rcr}" for c in FC}}, "#,##0.000",
        "9M＝公司口徑實績（1H 10-Q 16.139＋Q3 法說 14.511）；公司指引 37–39 只進 Recon 與 Sens（J5、原則 59：不以指引作界值）")
    # 折舊
    retdep = {"F": "0", "G": "0", "H": "ret29*cx_h100", "I": "(ret29+ret30)*cx_h100",
              "J": "((ret29+ret30+ret31*v25_h100)*cx_h100+ret31*v25_gb200*cx_gb200+ret31*(1-v25_h100-v25_gb200)*cx_gb300)",
              "K": "((ret29+ret30+ret31*v25_h100)*cx_h100+ret31*v25_gb200*cx_gb200+ret31*(1-v25_h100-v25_gb200)*cx_gb300+ret32*(v26_gb300*cx_gb300+(1-v26_gb300)*cx_gb200))"}
    rdp = nr("dep")
    dv = {"C": "=da9m", "D": f"=da_q3+cip_q2/life_ai*0.25+(D{rcx}-D{rcf})/life_ai*0.25/2+D{rcf}/fac_life*0.25/2", "E": f"=C{rdp}+D{rdp}"}
    for i, c in enumerate(FC):
        prior = ["D"] + FC[:i]
        dv[c] = (f"=MAX(0,da_q3*4-{retdep[c]}/life_ai/1000)+cip_q2/life_ai+(" + "+".join(f"{x}{rcx}-{x}{rcf}" for x in prior) + ")/life_ai+("
                 + "+".join(f"{x}{rcf}" for x in prior) + f")/fac_life+({c}{rcx}-{c}{rcf})/life_ai/2+{c}{rcf}/fac_life/2")
    put(TH, rdp, "折舊攤銷", "US$B", dv, "#,##0.000", "Q3 年化 7.556 − 已汰換批次＋在建工程 11.9 ÷ 年限＋新資本支出 ÷ 年限（當年半年）")
    rsb = nr("sbc")
    put(TH, rsb, "股票薪酬", "US$B", {"C": "=sbc9m", "D": f"=D{rr_}*sbc_r", "E": f"=C{rsb}+D{rsb}", **{c: f"={c}{rr_}*sbc_r" for c in FC}}, "#,##0.000")
    rei = nr("ebit")
    put(TH, rei, "營業利益（GAAP）", "US$B", {"C": "=oi9m", "D": f"=D{re_}-D{rsb}-D{rdp}", "E": f"=C{rei}+D{rei}", **{c: f"={c}{re_}-{c}{rsb}-{c}{rdp}" for c in FC}}, "#,##0.000")
    put(TH, nr("om"), "  營業利益率", "%", {c: f"=IF({c}{rr_}=0,0,{c}{rei}/{c}{rr_})" for c in "CDEFGHIJK"}, "0.0%")
    rcap = nr("capb")
    put(TH, rcap, "累計投入資本（PP&E 淨額＋其後資本支出 − 折舊，期末）", "US$B", {"C": "=ppe_q3", "D": f"=C{rcap}+D{rcx}-D{rdp}", "E": f"=D{rcap}", **{c: f"={PREV[c]}{rcap}+{c}{rcx}-{c}{rdp}" for c in FC}}, "#,##0.0", "2026-09-30 PP&E 淨額 56.778（新聞稿）")
    put(TH, nr("roic"), "  稅後 ROIC（營業利益 ×（1 − 稅率）÷ 期初投入資本）", "%", {c: f"=IF({PREV[c]}{rcap}<=0,0,{c}{rei}*(1-tax_lt)/{PREV[c]}{rcap})" for c in FC}, "0.0%", "對照 Val 頁 WACC")
    TR[s] = R
    r += 1
RPO_ = TH.max_row + 2
TH.cell(RPO_, 1, "命題答案摘要（FY2029；核心問題）").font = BOLD
hdr(TH, RPO_ + 1, ["項目", "單位", "情境 A", "情境 B", "情境 C", "說明"])
SUMQ = [("在役 MW（2029 年底）", "MW", lambda s: f"=H{TR[s]['conn']}", "#,##0", ""),
        ("AI 雲端營收 FY2029", "US$B", lambda s: f"=H{TR[s]['rev']}", "#,##0.0", ""),
        ("調整後 EBITDA 率 FY2029", "%", lambda s: f"=H{TR[s]['ebm']}", "0.0%", ""),
        ("營業利益率 FY2029（GAAP）", "%", lambda s: f"=H{TR[s]['om']}", "0.0%", ""),
        ("稅後 ROIC FY2029", "%", lambda s: f"=H{TR[s]['roic']}", "0.0%", "對照 WACC（Val 頁）"),
        ("資本支出累計 FY2027–FY2029", "US$B", lambda s: f"=F{TR[s]['cx']}+G{TR[s]['cx']}+H{TR[s]['cx']}", "#,##0.0", ""),
        ("收支關卡：FY2026–FY2030 自由現金流扣股利", "US$B", lambda s: f"=Fund!{'CDE'[SC.index(s)]}5", "#,##0.0", "Fund 頁")]
for i, (lab, unit, f, fmt, note) in enumerate(SUMQ):
    rr_ = RPO_ + 2 + i
    TH.cell(rr_, 1, lab); TH.cell(rr_, 2, unit); TH.cell(rr_, 6, note)
    for j, s in enumerate(SC): TH.cell(rr_, 3 + j, f(s)).number_format = fmt
SUM_ROWS = (RPO_ + 1, RPO_ + 1 + len(SUMQ))
TH.column_dimensions["A"].width = 62; TH.column_dimensions["B"].width = 13
for c in "CDEFGHIJK": TH.column_dimensions[c].width = 11
TH.column_dimensions["L"].width = 70
TH.freeze_panes = "C4"

# =====================================================================================
# Seg：唯一報告分部
# =====================================================================================
SG = wb.create_sheet("Seg")
SG["A1"] = "Seg — 營收分部層：CoreWeave 只有一個應報導部門（AI 雲端平台，10-K 分部附註），沒有其他營收線揭露；命題模組（在役 MW × 每 MW 年收入）即驅動整個分部。FY2025 實績、FY2026＝9M 實績＋Q4 模型、延伸到 FY2032"
SG["A2"] = "核對列：命題子列＝分部、分部＝合併營收；「其他／服務」：公司未揭露（不存在），不另列"
YH = ["項目", "單位"] + [f"FY{y}" + ("A" if y == 2025 else "") for y in YRS] + ["說明"]
hdr(SG, 3, YH)
r = 4
SR = {}
for s in SC:
    SG.cell(r, 1, SCN[s]).font = BOLD; r += 1
    R = {}
    for k, lab, f in (("ai", "AI 雲端：命題模組（在役 MW × 每 MW 年收入）", lambda y: f"=Thesis!{TCOL[y]}{TR[s]['rev']}"),
                      ("seg", "分部：AI 雲端平台", None), ("tot", "合併營收", None)):
        R[k] = r; SG.cell(r, 1, lab); SG.cell(r, 2, "US$B"); r += 1
    SG[f"C{R['seg']}"] = "=Q!G4"; SG[f"C{R['tot']}"] = f"=C{R['seg']}"
    for y in range(2026, 2033):
        c = yc(y)
        SG[f"{c}{R['ai']}"] = f"=Thesis!{TCOL[y]}{TR[s]['rev']}"
        SG[f"{c}{R['seg']}"] = f"={c}{R['ai']}"
        SG[f"{c}{R['tot']}"] = f"={c}{R['seg']}"
    R["g"] = r; SG.cell(r, 1, "  合併營收年增"); SG.cell(r, 2, "%")
    for y in range(2026, 2033): SG[f"{yc(y)}{r}"] = f"={yc(y)}{R['tot']}/{yc(y - 1)}{R['tot']}-1"; SG[f"{yc(y)}{r}"].number_format = "0.0%"
    r += 1
    R["chk1"] = r; SG.cell(r, 1, "檢查：命題子列 − 分部（應為 0）")
    for y in range(2026, 2033): SG[f"{yc(y)}{r}"] = f"={yc(y)}{R['ai']}-{yc(y)}{R['seg']}"
    r += 1
    R["chk2"] = r; SG.cell(r, 1, "檢查：分部合計 − 合併營收（應為 0）")
    for y in YRS: SG[f"{yc(y)}{r}"] = f"={yc(y)}{R['seg']}-{yc(y)}{R['tot']}"
    r += 1
    for k in ("ai", "seg", "tot", "chk1", "chk2"):
        for y in YRS:
            if SG[f"{yc(y)}{R[k]}"].value is not None: SG[f"{yc(y)}{R[k]}"].number_format = "#,##0.000"
    SR[s] = R; r += 1
SG.column_dimensions["A"].width = 48
for c in "CDEFGHIJ": SG.column_dimensions[c].width = 11
SG.column_dimensions["K"].width = 60

# =====================================================================================
# PL
# =====================================================================================
PL = wb.create_sheet("PL")
PL["A1"] = "PL — 三情境損益（十億美元；EPS 美元）。營收（Seg）→ 營運費用、租金、管銷（Thesis）→ 調整後 EBITDA → 減股票薪酬、折舊＝GAAP 營業利益 → 利息（CF）→ 稅（NOL 80% 上限）→ 淨利 → EPS"
PL["A2"] = "FY2025 為實績（10-K）；FY2026＝9M 實績（新聞稿）＋Q4 模型，上方費用明細只列 FY2027 起（9M 營業成本、其他收入、所得稅已由 Q3 10-Q XBRL 拆分，見本頁末標準損益表）。所得稅：長期現金稅率 × MAX(0, 稅前 − NOL 使用)；上方 FY2026 的稅前與所得稅列只含模型項，淨利＝9M 實際淨損（已含 9M 實際所得稅與其他收入）＋Q4 模型"
hdr(PL, 3, YH)
r = 4
PL.cell(r, 1, "分部核對（三情境合計絕對差；每年度應為 0）").font = BOLD; r += 1
CHK_SEG = r; PL.cell(r, 1, "檢查：分部合計 − 合併營收"); r += 1
CHK_SUB = r; PL.cell(r, 1, "檢查：命題子列 − 分部"); r += 2
PR_ = {}
def prow(R, k, lab, vals, fmt="#,##0.000", note=None, unit="US$B"):
    global r
    R[k] = r; PL.cell(r, 1, lab); PL.cell(r, 2, unit)
    for c, v in vals.items():
        if v is None: continue
        PL[f"{c}{r}"] = v; PL[f"{c}{r}"].number_format = fmt
    if note: PL.cell(r, 11, note)
    r += 1
for s in SC:
    PL.cell(r, 1, SCN[s]).font = BOLD; r += 1
    R = {}; T = TR[s]; G = SR[s]
    prow(R, "ai", "AI 雲端：命題模組（在役 MW × 每 MW 年收入）", {yc(y): f"=Seg!{yc(y)}{G['ai']}" for y in range(2026, 2033)})
    prow(R, "seg", "分部：AI 雲端平台", {yc(y): f"=Seg!{yc(y)}{G['seg']}" for y in YRS})
    prow(R, "rev", "合併營收", {yc(y): f"=Seg!{yc(y)}{G['tot']}" for y in YRS})
    prow(R, "revg", "  合併營收年增", {yc(y): f"={yc(y)}{R['rev']}/{yc(y - 1)}{R['rev']}-1" for y in range(2026, 2033)}, "0.0%", unit="%")
    prow(R, "opex", "營運費用（Tokenomics；不含折舊、租金）", {yc(y): f"=Thesis!{TCOL[y]}{T['opex']}" for y in FYM})
    prow(R, "rent", "機房租金", {yc(y): f"=Thesis!{TCOL[y]}{T['rent']}" for y in FYM})
    prow(R, "sga", "管銷費用（扣 SBC）", {yc(y): f"=Thesis!{TCOL[y]}{T['sga']}" for y in FYM})
    prow(R, "aeb", "調整後 EBITDA", {yc(y): f"=Thesis!{TCOL[y]}{T['ebitda']}" for y in range(2026, 2033)})
    prow(R, "aebm", "  調整後 EBITDA 率", {yc(y): f"={yc(y)}{R['aeb']}/{yc(y)}{R['rev']}" for y in range(2026, 2033)}, "0.0%", unit="%")
    prow(R, "sbc", "股票薪酬", {yc(y): f"=Thesis!{TCOL[y]}{T['sbc']}" for y in range(2026, 2033)})
    prow(R, "da", "折舊攤銷", {yc(y): f"=Thesis!{TCOL[y]}{T['dep']}" for y in range(2026, 2033)})
    prow(R, "oi", "營業利益（GAAP）", {"C": "=Q!G5", **{yc(y): f"=Thesis!{TCOL[y]}{T['ebit']}" for y in range(2026, 2033)}})
    prow(R, "oim", "  營業利益率", {yc(y): f"={yc(y)}{R['oi']}/{yc(y)}{R['rev']}" for y in range(2026, 2033)}, "0.0%", unit="%")
    prow(R, "aoi", "調整後營業利益（＝營業利益＋股票薪酬；公司口徑近似）", {"D": f"=aoi9m+Thesis!D{T['ebit']}+Thesis!D{T['sbc']}", **{yc(y): f"={yc(y)}{R['oi']}+{yc(y)}{R['sbc']}" for y in FYM}}, note="FY2026＝9M 實績 0.528＋Q4 模型")
    prow(R, "int", "利息費用淨額", {}, note="CF 頁：既有借款 × r_old＋瀑布新債 × kd＋高息債 × r_junk（期初餘額）")
    prow(R, "pti", "稅前利益", {}, note="其他收益：預測期設 0")
    prow(R, "nolb", "NOL 期初餘額", {}, note="FY2026 期初＝2025 年底 nol0")
    prow(R, "nolu", "NOL 使用", {})
    prow(R, "tax", "所得稅（長期現金稅率 × 抵減後應稅所得）", {})
    prow(R, "nole", "NOL 期末餘額", {})
    prow(R, "ni", "淨利", {"C": "=Q!G8"})
    prow(R, "sh", "稀釋股數（平均）", {}, unit="十億股")
    prow(R, "eps", "EPS（GAAP 稀釋）", {"C": -2.81}, "0.00", unit="US$")
    PR_[s] = R; r += 1
for c in "CDEFGHIJ":
    PL[f"{c}{CHK_SEG}"] = "=" + "+".join(f"ABS(Seg!{c}{SR[s]['chk2']})" for s in SC)
    PL[f"{c}{CHK_SUB}"] = "=" + ("+".join(f"ABS(Seg!{c}{SR[s]['chk1']})" for s in SC) if c != "C" else "0")
    for x in (CHK_SEG, CHK_SUB): PL[f"{c}{x}"].number_format = "0.000"
PL.column_dimensions["A"].width = 50
for c in "CDEFGHIJ": PL.column_dimensions[c].width = 11
PL.column_dimensions["K"].width = 70
PL.freeze_panes = "C4"

# =====================================================================================
# CF：現金流與融資瀑布（收支模組）
# =====================================================================================
CF = wb.create_sheet("CF")
CF["A1"] = "CF — 現金流與融資瀑布（收支模組）。營運現金＝淨利＋折舊＋股票薪酬＋客戶預付流入 − 預付認列（非現金營收）− 營運資金增加；自由現金流的缺口依序由現金（保留最低現金）→ 投資級債（借款 ÷ 調整後 EBITDA ≤ lev_max）→ 新股（每年 ≤ 市值 eq_cap）→ 高息債 補足"
CF["A2"] = "FY2026＝9M 實績＋Q4 模型；瀑布只計 Q4（起點為 2026-09-30 現金 4.122，受限現金不計）。利息以期初餘額計（不循環）。到期借款假設以原利率展延（Debt 頁列再融資備忘）。無股利、無回購"
CFH = 4
hdr(CF, CFH, YH)
r = CFH + 1
CR_ = {}
SPEC = ["ni", "da", "sbc", "ppin", "pprec", "nwc", "cfo", "cx", "fcf", "L", "cash0", "flow", "need", "ebd", "cap", "ig", "eq", "hy", "cash1", "igc", "hyc", "debt",
        "nde", "ndw", "icov", "icw", "sh0", "neweq", "sh1", "shd", "eqm", "dil", "int"]
WF = "DEFGHIJ"
for s in SC:
    CF.cell(r, 1, SCN[s]).font = BOLD; r += 1
    R = {k: r + i for i, k in enumerate(SPEC)}; T = TR[s]; P = PR_[s]
    def cw(k, lab, vals, fmt="#,##0.000", note=None, unit="US$B"):
        CF.cell(R[k], 1, lab); CF.cell(R[k], 2, unit)
        for c, v in vals.items():
            if v is None: continue
            CF[f"{c}{R[k]}"] = v
            if fmt: CF[f"{c}{R[k]}"].number_format = fmt
        if note: CF.cell(R[k], 11, note)
    F_ = lambda k: R[k]
    cw("ni", "淨利", {c: f"=PL!{c}{P['ni']}" for c in WF})
    cw("da", "折舊攤銷", {c: f"=PL!{c}{P['da']}" for c in WF})
    cw("sbc", "股票薪酬（非現金）", {c: f"=PL!{c}{P['sbc']}" for c in WF})
    cw("ppin", "客戶預付流入（資本支出 × 覆蓋比；Q4 另加 pp_q4）", {"D": f"=prepay_cov*Thesis!D{T['cx']}+pp_q4", **{yc(y): f"=prepay_cov*Thesis!{TCOL[y]}{T['cx']}" for y in FYM}}, note="FY2026 只計 Q4（9M 預付已在營運現金流實績內）")
    pv = {}
    for i, c in enumerate(WF):
        f_ = f"=def_open/prepay_yrs*{'0.25' if c == 'D' else '1'}"
        for j, x in enumerate(WF[:i]):
            f_ += f"+IF({i - j}<=prepay_yrs,{x}{F_('ppin')}/prepay_yrs,0)"
        pv[c] = f_
    cw("pprec", "預付認列（非現金營收，自營運現金扣除）", pv, note="2026-06-30 遞延收入 9.692 自 Q4 起、各年新流入自次年起，依 prepay_yrs 年直線")
    cw("nwc", "營運資金增加（預付除外）", {"D": f"=nwc_r*(Thesis!D{T['rev']}-rev_q3)", **{yc(y): f"=nwc_r*(PL!{yc(y)}{P['rev']}-PL!{yc(y - 1)}{P['rev']})" for y in FYM}}, note="FY2026 只計 Q4 對 Q3 的增量")
    cw("cfo", "營運現金流", {"D": f"=cfo9m+(Thesis!D{T['ebit']}-(PL!D{P['int']}-int9m))+Thesis!D{T['dep']}+Thesis!D{T['sbc']}+D{F_('ppin')}-D{F_('pprec')}-D{F_('nwc')}",
                             **{yc(y): f"={yc(y)}{F_('ni')}+{yc(y)}{F_('da')}+{yc(y)}{F_('sbc')}+{yc(y)}{F_('ppin')}-{yc(y)}{F_('pprec')}-{yc(y)}{F_('nwc')}" for y in FYM}},
       note="FY2026＝9M 實績 5.638＋Q4（Q4 無稅）")
    cw("cx", "現金資本支出", {"D": f"=cx9m_cash+Thesis!D{T['cx']}", **{yc(y): f"=Thesis!{TCOL[y]}{T['cx']}" for y in FYM}}, note="FY2026＝9M 現金購置 23.558＋Q4 模型；9M 公司口徑 30.650 與現金 23.558 的差額為應付設備款與 OEM 融資（已含在 Q3 借款餘額）")
    cw("fcf", "自由現金流（無股利，即收支關卡口徑）", {c: f"={c}{F_('cfo')}-{c}{F_('cx')}" for c in WF})
    cw("L", "瀑布期間長度（FY2026 只計 Q4）", {"D": 0.25, **{yc(y): 1 for y in FYM}}, "0.00", unit="年")
    cw("cash0", "期初現金（FY2026＝2026-09-30）", {"D": "=cash_q3", **{yc(y): f"={yc(y - 1)}{F_('cash1')}" for y in FYM}})
    cw("flow", "融資前現金流（FY2026 只計 Q4）", {"D": f"=D{F_('fcf')}-(cfo9m-cx9m_cash)", **{yc(y): f"={yc(y)}{F_('fcf')}" for y in FYM}})
    cw("need", "外部融資需求（扣除可動用現金後）", {c: f"=MAX(0,-({c}{F_('flow')}+{c}{F_('cash0')}-min_cash))" for c in WF})
    cw("ebd", "調整後 EBITDA（年化；槓桿分母）", {"D": f"=Thesis!D{T['ebitda']}*4", **{yc(y): f"=PL!{yc(y)}{P['aeb']}" for y in FYM}}, note="FY2026 以 Q4 × 4")
    cw("cap", "投資級債可新增額度", {c: f"=MAX(0,lev_max*{c}{F_('ebd')}-" + ("debt_q3" if c == "D" else f"{chr(ord(c) - 1)}{F_('debt')}") + ")" for c in WF}, note="lev_max × 年化調整後 EBITDA − 期初借款")
    cw("ig", "新增投資級債", {c: f"=MIN({c}{F_('need')},{c}{F_('cap')})" for c in WF})
    cw("eq", "新股", {c: f"=MIN({c}{F_('need')}-{c}{F_('ig')},eq_cap*px*{c}{F_('sh0')}*{c}{F_('L')})" for c in WF})
    cw("hy", "高息債", {c: f"={c}{F_('need')}-{c}{F_('ig')}-{c}{F_('eq')}" for c in WF})
    cw("cash1", "期末現金", {c: f"={c}{F_('cash0')}+{c}{F_('flow')}+{c}{F_('ig')}+{c}{F_('eq')}+{c}{F_('hy')}" for c in WF})
    cw("igc", "累計新增投資級債（期末）", {"D": f"=D{F_('ig')}", **{yc(y): f"={yc(y - 1)}{F_('igc')}+{yc(y)}{F_('ig')}" for y in FYM}})
    cw("hyc", "累計高息債（期末）", {"D": f"=D{F_('hy')}", **{yc(y): f"={yc(y - 1)}{F_('hyc')}+{yc(y)}{F_('hy')}" for y in FYM}})
    cw("debt", "借款合計（期末）", {c: f"=debt_q3+{c}{F_('igc')}+{c}{F_('hyc')}" for c in WF})
    cw("nde", "淨負債 ÷ 調整後 EBITDA（分母 ≤0 為 N/M）", {c: f'=IF({c}{F_("ebd")}<=0,"N/M",({c}{F_("debt")}+finl_q3-{c}{F_("cash1")})/{c}{F_("ebd")})' for c in WF}, "0.00", unit="x", note="原則第 10 節第 9、56 條")
    cw("ndw", "  檢查：淨負債 ÷ EBITDA ≤ lev_max", {c: f'=IF(AND(ISNUMBER({c}{F_("nde")}),{c}{F_("nde")}>=0,{c}{F_("nde")}<=lev_max),"OK",IF(ISNUMBER({c}{F_("nde")}),IF({c}{F_("nde")}<0,"OK","WARN"),"WARN（N/M）"))' for c in WF}, None, unit="—")
    cw("icov", "利息保障倍數（調整後 EBITDA ÷ 利息；分母 ≤0 為 N/M）", {c: f'=IF({c}{F_("int")}<=0,"N/M",{c}{F_("ebd")}*{c}{F_("L")}/{c}{F_("int")})' for c in WF}, "0.00", unit="x")
    cw("icw", "  檢查：利息保障倍數 ≥ icov_min", {c: f'=IF(AND(ISNUMBER({c}{F_("icov")}),{c}{F_("icov")}>=icov_min),"OK","WARN")' for c in WF}, None, unit="—")
    cw("sh0", "期初普通股", {"D": "=sh_out", **{yc(y): f"={yc(y - 1)}{F_('sh1')}" for y in FYM}}, "0.0000", unit="十億股")
    cw("neweq", "新發行股數", {c: f"={c}{F_('eq')}/(px*(1-eq_disc))" for c in WF}, "0.0000", unit="十億股")
    cw("sh1", "期末普通股", {c: f"={c}{F_('sh0')}+{c}{F_('neweq')}" for c in WF}, "0.0000", unit="十億股")
    cw("shd", "稀釋股數（平均＋RSU／選擇權淨股數）", {"D": f"=(sh_out*0.75+(D{F_('sh0')}+D{F_('sh1')})/2*0.25)+rsu_net", **{yc(y): f"=({yc(y)}{F_('sh0')}+{yc(y)}{F_('sh1')})/2+rsu_net" for y in FYM}}, "0.0000", unit="十億股")
    cw("eqm", "新股金額 ÷ 當年期初市值（原則第 62 條）", {c: f"={c}{F_('eq')}/(px*{c}{F_('sh0')})" for c in WF}, "0.0%", unit="%")
    cw("dil", "累計稀釋（期末股數 ÷ 2026-09-30 股數 − 1）", {c: f"={c}{F_('sh1')}/sh_out-1" for c in WF}, "0.0%", unit="%")
    cw("int", "利息費用淨額", {"D": "=r_old*debt_q3*0.25", **{yc(y): f"=r_old*debt_q3+r_new*{yc(y - 1)}{F_('igc')}+r_junk*{yc(y - 1)}{F_('hyc')}" for y in FYM}}, note="FY2026 欄只計 Q4（PL 加 9M 實績）")
    CR_[s] = R
    r = R["int"] + 2
# 回填 PL（利息、稅、NOL、淨利、股數、EPS）
for s in SC:
    P = PR_[s]; C_ = CR_[s]; T = TR[s]
    PL[f"D{P['int']}"] = f"=int9m+CF!D{C_['int']}"
    for y in FYM: PL[f"{yc(y)}{P['int']}"] = f"=CF!{yc(y)}{C_['int']}"
    for y in range(2026, 2033):
        c = yc(y)
        PL[f"{c}{P['pti']}"] = f"={c}{P['oi']}-{c}{P['int']}"
        PL[f"{c}{P['nolb']}"] = "=nol0" if y == 2026 else f"={yc(y - 1)}{P['nole']}"
        PL[f"{c}{P['nolu']}"] = f"=MIN({c}{P['nolb']},nol_lim*MAX(0,{c}{P['pti']}))"
        PL[f"{c}{P['tax']}"] = f"=tax_lt*MAX(0,{c}{P['pti']}-{c}{P['nolu']})"
        PL[f"{c}{P['nole']}"] = f"={c}{P['nolb']}-{c}{P['nolu']}+MAX(0,-{c}{P['pti']})"
        PL[f"{c}{P['ni']}"] = f"={c}{P['pti']}-{c}{P['tax']}"
        PL[f"{c}{P['sh']}"] = f"=CF!{c}{C_['shd']}"
        PL[f"{c}{P['eps']}"] = f"={c}{P['ni']}/{c}{P['sh']}"
        for k, fm in (("int", "#,##0.000"), ("pti", "#,##0.000"), ("nolb", "#,##0.000"), ("nolu", "#,##0.000"), ("tax", "#,##0.000"), ("nole", "#,##0.000"), ("ni", "#,##0.000"), ("sh", "0.0000"), ("eps", "0.00")):
            PL[f"{c}{P[k]}"].number_format = fm
    PL[f"D{P['ni']}"] = f"=ni9m+(Thesis!D{T['ebit']}-CF!D{C_['int']})-D{P['tax']}"
    PL.cell(P["ni"], 11, "FY2026＝9M 實績 −2.075＋Q4 模型；其後＝稅前 − 所得稅")
CF.column_dimensions["A"].width = 52
for c in "CDEFGHIJ": CF.column_dimensions[c].width = 11
CF.column_dimensions["K"].width = 70
CF.freeze_panes = "C5"

# =====================================================================================
# Fund：資金缺口與融資（收支模組摘要；全部引用 CF）
# =====================================================================================
FU = wb.create_sheet("Fund")
FU["A1"] = "Fund — 資金缺口與融資（收支模組摘要；全部引用 CF 頁，模型只有一組現金數字）"
FU["A2"] = "收支關卡＝FY2026–FY2030 自由現金流合計（無股利）；旁列期間流動性：最大單年新增舉債、最低年末現金、淨負債 ÷ EBITDA 與利息保障倍數的最差年（原則第 10 節第 55 條）；新股占市值與累計稀釋（第 62 條，門檻是否 WARN 待 Andy 決定，只呈現）"
hdr(FU, 3, ["項目", "單位", "情境 A", "情境 B", "情境 C", "說明"])
FS = {}
FSPEC = [("gate", "收支關卡：FY2026–FY2030 自由現金流合計", "US$B", lambda C_: f"=SUM(CF!D{C_['fcf']}:H{C_['fcf']})", "#,##0.0", "負值＝需外部融資；地圖收支關卡讀這一列"),
         ("cum", "Q4 2026–FY2030 外部融資合計（投資級債＋新股＋高息債）", "US$B", lambda C_: f"=SUM(CF!D{C_['need']}:H{C_['need']})", "#,##0.0", ""),
         ("ig", "  其中：新增投資級債", "US$B", lambda C_: f"=SUM(CF!D{C_['ig']}:H{C_['ig']})", "#,##0.0", ""),
         ("eq", "  其中：新股", "US$B", lambda C_: f"=SUM(CF!D{C_['eq']}:H{C_['eq']})", "#,##0.0", ""),
         ("hy", "  其中：高息債", "US$B", lambda C_: f"=SUM(CF!D{C_['hy']}:H{C_['hy']})", "#,##0.0", ""),
         ("maxb", "期間最大單年新增舉債（投資級債＋高息債）", "US$B", lambda C_: "=MAX(" + ",".join(f"CF!{c}{C_['ig']}+CF!{c}{C_['hy']}" for c in "DEFGH") + ")", "#,##0.0", "原則 55"),
         ("minc", "期間最低年末現金", "US$B", lambda C_: f"=MIN(CF!D{C_['cash1']}:H{C_['cash1']})", "#,##0.0", ""),
         ("ndmax", "淨負債 ÷ 調整後 EBITDA 最差年（FY2026–FY2030；N/M 年不計）", "x", lambda C_: f"=MAX(CF!D{C_['nde']}:H{C_['nde']})", "0.00", "門檻 4.5（WARN，不截斷）"),
         ("icmin", "利息保障倍數最差年（FY2026–FY2030）", "x", lambda C_: f"=MIN(CF!D{C_['icov']}:H{C_['icov']})", "0.00", "門檻 3（WARN）"),
         ("warn", "原則 9 兩指標是否 WARN（FY2026–FY2030）", "—", lambda C_: f'=IF(COUNTIF(CF!D{C_["ndw"]}:H{C_["ndw"]},"WARN*")+COUNTIF(CF!D{C_["icw"]}:H{C_["icw"]},"WARN")>0,"WARN","OK")', None, ""),
         ("eqmax", "新股 ÷ 期初市值：最高年", "%", lambda C_: f"=MAX(CF!D{C_['eqm']}:H{C_['eqm']})", "0.0%", "原則 62（只呈現）"),
         ("dil", "累計稀釋（2030 年底股數 ÷ 2026-09-30 − 1）", "%", lambda C_: f"=CF!H{C_['dil']}", "0.0%", "未來新股不進 DCF 每股（X35）"),
         ("chk", "檢查：Q4 2026–FY2030 外部融資 − 三項合計（應為 0）", "US$B", lambda C_: None, "0.000", "")]
r = 4
for k, lab, unit, f, fmt, note in FSPEC:
    FS[k] = r + 1; r += 1
    FU.cell(FS[k], 1, lab); FU.cell(FS[k], 2, unit); FU.cell(FS[k], 6, note)
    for j, s in enumerate(SC):
        v = f(CR_[s]) if k != "chk" else f"={'CDE'[j]}{FS['cum']}-{'CDE'[j]}{FS['ig']}-{'CDE'[j]}{FS['eq']}-{'CDE'[j]}{FS['hy']}"
        c = FU.cell(FS[k], 3 + j, v)
        if fmt: c.number_format = fmt
FS["trough"] = FS["gate"]
r = FS["chk"] + 2
FU.cell(r, 1, "情境 B：逐年來源與用途（引用 CF 頁）").font = BOLD; r += 1
FUH = r
hdr(FU, r, ["項目", "單位", ""] + [f"FY{y}" + ("（Q4）" if y == 2026 else "") for y in range(2026, 2033)]); r += 1
FR = {}
CB = CR_["B"]
for k, lab in (("flow", "融資前現金流"), ("cash0", "期初現金"), ("need", "外部融資需求"), ("ig", "新增投資級債"), ("eq", "新股"), ("hy", "高息債"), ("cash1", "期末現金"), ("debt", "借款合計（期末）"), ("nde", "淨負債 ÷ 調整後 EBITDA"), ("icov", "利息保障倍數"), ("eqm", "新股 ÷ 期初市值"), ("dil", "累計稀釋")):
    FR[k] = r; FU.cell(r, 1, lab); FU.cell(r, 2, "x" if k in ("nde", "icov") else ("%" if k in ("eqm", "dil") else "US$B"))
    for j, c in enumerate(WF):
        x = FU.cell(r, 4 + j, f"=CF!{c}{CB[k]}"); x.number_format = "0.00" if k in ("nde", "icov") else ("0.0%" if k in ("eqm", "dil") else "#,##0.00")
    r += 1
FU.column_dimensions["A"].width = 60; FU.column_dimensions["F"].width = 40
for c in "CDEGHIJ": FU.column_dimensions[c].width = 11

# =====================================================================================
# RPO：backlog 轉營收路徑
# =====================================================================================
RP = wb.create_sheet("RPO")
RP["A1"] = "RPO — 合約排程對模型營收（同口徑檢查，不作校準）。RPO 132.3（2026-09-30；24 個月內 36%、25–48 個月 38%、48 個月後 26%，Interested-party）依桶平均攤到年度：24 個月內＝每年 1/2、25–48 個月＝每年 1/2、48 個月後＝2031–2034 每年 1/4"
RP["A2"] = "覆蓋率＝RPO 排程 ÷ 模型營收（情境 B）；>1＝合約已排程金額高於模型營收（模型偏保守），<1＝模型營收需靠新簽或隨需。Revenue backlog 132.5 另含公司估計的非 RPO 金額"
hdr(RP, 4, ["項目", "單位", "", "FY2026（Q4）", "FY2027", "FY2028", "FY2029", "FY2030", "FY2031", "FY2032"])
RPB = {}
r = 5
RP.cell(r, 1, "RPO 排程（年化分攤）"); RP.cell(r, 2, "US$B")
sched = {"D": "=rpo*rpo_24/2*0.25", "E": "=rpo*rpo_24/2", "F": "=rpo*rpo_24/2*0.75+rpo*rpo_48/2*0.25", "G": "=rpo*rpo_48/2", "H": "=rpo*rpo_48/2*0.75+rpo*rpo_l/4*0.25", "I": "=rpo*rpo_l/4", "J": "=rpo*rpo_l/4"}
for c, v in sched.items(): RP[f"{c}{r}"] = v; RP[f"{c}{r}"].number_format = "#,##0.0"
RPB["sched"] = r; r += 1
for s in SC:
    RP.cell(r, 1, f"模型營收：{SCN[s]}"); RP.cell(r, 2, "US$B")
    RP[f"D{r}"] = f"=Thesis!D{TR[s]['rev']}"
    for j, y in enumerate(FYM): RP[f"{CL(5 + j)}{r}"] = f"=Thesis!{TCOL[y]}{TR[s]['rev']}"
    for c in "DEFGHIJ": RP[f"{c}{r}"].number_format = "#,##0.0"
    RPB["rev" + s] = r; r += 1
RP.cell(r, 1, "覆蓋率：RPO 排程 ÷ 模型營收（情境 B）"); RP.cell(r, 2, "x")
for c in "DEFGHIJ": RP[f"{c}{r}"] = f"={c}{RPB['sched']}/{c}{RPB['revB']}"; RP[f"{c}{r}"].number_format = "0.00"
RPB["cov"] = r; r += 1
RP.cell(r, 1, "模型營收中未被 RPO 覆蓋的部分（情境 B；需新簽或隨需）"); RP.cell(r, 2, "US$B")
for c in "DEFGHIJ": RP[f"{c}{r}"] = f"=MAX(0,{c}{RPB['revB']}-{c}{RPB['sched']})"; RP[f"{c}{r}"].number_format = "#,##0.0"
RPB["gap"] = r; r += 1
RP.cell(r, 1, "對照：2026Q4–2032 RPO 排程合計 − RPO（負值＝2033 年後才認列的部分）"); RP.cell(r, 2, "US$B")
RP[f"D{r}"] = f"=SUM(D{RPB['sched']}:J{RPB['sched']})-rpo"; RP[f"D{r}"].number_format = "#,##0.0"; RPB["chk"] = r
RP.column_dimensions["A"].width = 62
for c in "DEFGHIJ": RP.column_dimensions[c].width = 11

# =====================================================================================
# Sites：具名站點與容量上限
# =====================================================================================
SI = wb.create_sheet("Sites")
SI["A1"] = "Sites — 具名站點（沿用 v4.13 legacy_v4/company.json defaults.sites，Interested-party／二手；Q3 未逐站更新：待 Q3 10-Q 與房東公告）與容量上限對照"
SI["A2"] = "MW 為 critical IT 口徑（房東揭露）；具名站點合計遠小於已簽約 4.6 GW——其餘未具名（公司未逐站揭露，找不到）"
hdr(SI, 4, ["站點", "房東／營運", "規劃 MW", "已通電 MW", "時程", "把握"])
SITES = [("Helios, Dickens County, Texas", "Galaxy Digital（15 年租約）", 526, 133, "2026 下半年起分期", "中"), ("Polaris Forge 1, Ellendale, ND", "Applied Digital（15 年租約）", 400, 150, "2026–2027", "中"),
         ("Core Scientific 多站（Denton／Muskogee／Dalton 等）", "Core Scientific", 590, 420, "2026–2027", "中"), ("Kenilworth (NEST), New Jersey", "CoreWeave 自建", 100, 0, "2026 年底", "低"),
         ("Lancaster, Pennsylvania", "CoreWeave 自建／Blue Owl JV", 300, 0, "2027 起", "低"), ("Indonesia（三站）", "合作開發", 360, 0, "2028", "低")]
ST = {"s0": 5}
for i, row in enumerate(SITES):
    for j, v in enumerate(row): SI.cell(5 + i, 1 + j, v)
r = 5 + len(SITES)
SI.cell(r, 1, "具名站點合計"); SI.cell(r, 3, f"=SUM(C5:C{r - 1})"); SI.cell(r, 4, f"=SUM(D5:D{r - 1})"); ST["sum"] = r; r += 1
SI.cell(r, 1, "已簽約電力（2026-09-30）"); SI.cell(r, 3, "=ctr_mw"); ST["ctr"] = r; r += 1
SI.cell(r, 1, "未具名（已簽約 − 具名站點）"); SI.cell(r, 3, f"=C{ST['ctr']}-C{ST['sum']}"); ST["res"] = r; r += 1
SI.cell(r, 1, "情境 A／B／C：2030 年底在役 MW ÷ 已簽約");
for j, s in enumerate(SC): SI.cell(r, 3 + j, f"=mw{s}_30/ctr_mw").number_format = "0.00"
ST["capc"] = r
SI.column_dimensions["A"].width = 50; SI.column_dimensions["B"].width = 34

# =====================================================================================
# Debt：逐檔與到期表（Q2 10-Q；待 Q3 10-Q）
# =====================================================================================
DB = wb.create_sheet("Debt")
DB["A1"] = "Debt — 債務逐檔（2026-06-30，Q2 10-Q；沿用 legacy_v4/company.json debt.instruments）＋Q3 新增（2033 可轉債 4.2、DDTL 5.5 額度 2.6）；Q3 逐檔與到期表待 Q3 10-Q（資料缺口）"
DB["A2"] = "基準假設：到期借款以原利率展延（再融資備忘列於下方到期表）；瀑布新債另計（CF 頁）。帳面合計與 Q3 資產負債表 46.873 的差額為 Q3 新增借款與攤還（逐檔未取得）"
hdr(DB, 4, ["債務", "類別", "到期", "利率", "本金（US$B）", "說明"])
INST = [("DDTL 1.0", "追索", "2028-03", 0.15, 1.3, "資產層、隨客戶付款攤還"), ("DDTL 2.0", "追索", "2030-08", 0.11, 3.19, ""), ("DDTL 2.1", "追索", "2031-03", 0.09, 3, ""),
        ("DDTL 3.0", "追索", "2030-08", 0.09, 2.215, ""), ("DDTL 5.0", "追索", "2031-11", 0.09, 1.101, "SOFR+4.50%；額度 3.1"), ("2030 Senior Notes", "追索", "2030-06", 0.10, 2, ""),
        ("2031 9.00% Senior Notes", "追索", "2031-02", 0.09, 1.75, ""), ("2031 9.75% Senior Notes", "追索", "2031-10", 0.0975, 2.75, "2026-04 發行"),
        ("2032 9.625% Senior Notes", "追索", "2032-07", 0.09625, 1.25, "2026-06 發行（kd 依據）"), ("2032 EUR Senior Notes（€2.0bn）", "追索", "2032-07", 0.085, 2.279, "已換匯換利避險"),
        ("2031 可轉債", "追索", "2031-12", 0.0175, 2.588, ""), ("2032 可轉債", "追索", "2032-10", 0.0175, 4, ""), ("OEM 與軟體授權融資", "追索", "2026-12～2030-07", 0.11, 4.22, ""),
        ("Magnetar Loan", "追索", "2029-01", 0.12, 0.189, ""), ("DDTL 4.0", "非追索", "2032-03", 0.07, 2.837, "SOFR+2.25%；額度 8.5"), ("OEM 融資（非追索）", "非追索", "2026-08～2028-08", 0.09, 0.882, ""),
        ("2033 可轉債（2026-09 發行，含超額配售）", "追索", "2033", 0.02875, 4.2, "Q3 新聞稿：upsized convertible senior notes 約 4.2"), ("DDTL 5.5（額度 2.6；動用額未揭露）", "追索", "—", None, 0, "Q3 新聞稿：$2.6B DDTL 5.5 已完成")]
for i, row in enumerate(INST):
    for j, v in enumerate(row):
        c = DB.cell(5 + i, 1 + j, v)
        if j == 3 and v is not None: c.number_format = "0.000%"
D_ = {"r0": 5}
r = 5 + len(INST)
DB.cell(r, 1, "逐檔合計（Q2 逐檔＋Q3 新增可轉債）"); DB.cell(r, 5, f"=SUM(E5:E{r - 1})").number_format = "#,##0.000"; D_["tot"] = r; r += 1
DB.cell(r, 1, "Q3 資產負債表借款帳面值"); DB.cell(r, 5, "=debt_q3").number_format = "#,##0.000"; r += 1
DB.cell(r, 1, "差額（Q3 新增與攤還，逐檔待 Q3 10-Q）"); DB.cell(r, 5, f"=E{r - 1}-E{r - 2}").number_format = "#,##0.000"; D_["diff"] = r; r += 2
DB.cell(r, 1, "到期與攤還表（Q2 10-Q；US$B）").font = BOLD; r += 1
D_["hdr"] = r
hdr(DB, r, ["項目", "", "", "2026 下半年", "2027", "2028", "2029", "2030", "2031 以後", "合計"]); r += 1
DB.cell(r, 1, "既有借款到期／攤還")
for j, v in enumerate((4.413, 6.184, 4.416, 2.421, 3.221, 14.896)): DB.cell(r, 4 + j, v).font = BLUE
DB.cell(r, 10, f"=SUM(D{r}:I{r})"); D_["old"] = r; r += 1
DB.cell(r, 1, "累計到期（基準：以原利率展延，不另耗現金）")
for j, c in enumerate("DEFGHI"): DB.cell(r, 4 + j, f"=SUM($D{D_['old']}:{c}{D_['old']})")
D_["cum"] = r; r += 1
DB.cell(r, 1, "再融資備忘：情境 B 瀑布新增投資級債（同期）")
for j, c in enumerate("DEFGH"): DB.cell(r, 4 + j, f"=CF!{c}{CR_['B']['ig']}")
D_["refi"] = r
for rr_ in range(D_["old"], r + 1):
    for c in range(4, 11): DB.cell(rr_, c).number_format = "#,##0.000"
DB.column_dimensions["A"].width = 46; DB.column_dimensions["F"].width = 40

# =====================================================================================
# Lease：租賃與租金路徑
# =====================================================================================
LS = wb.create_sheet("Lease")
LS["A1"] = "Lease — 租賃。在帳營業租賃負債 18.697、融資租賃 0.649（2026-09-30）；未起租租約 35.5（Q2 10-Q；Q3 待 10-Q）；單一站點 393 MW 未交付、租金上限 14.7／16 年（Q2 10-Q）"
LS["A2"] = "租金路徑＝平均在役 MW × 租賃比例 × 每 MW 年租金（Thesis）。A11 站點檢查：v4.13 推得五期租金可能低估 3.29；Q2 實際營業租賃費用年化 ÷ 租賃 MW＝1.57（含未通電站點），列為 Sens 高"
hdr(LS, 4, ["項目", "單位", "", "FY2026（Q4）", "FY2027", "FY2028", "FY2029", "FY2030", "FY2031", "FY2032"])
LSB = {}
r = 5
for s in SC:
    LS.cell(r, 1, f"機房租金：{SCN[s]}"); LS.cell(r, 2, "US$B")
    LS[f"D{r}"] = f"=Thesis!D{TR[s]['rent']}"
    for j, y in enumerate(FYM): LS[f"{CL(5 + j)}{r}"] = f"=Thesis!{TCOL[y]}{TR[s]['rent']}"
    LSB["rent" + s] = r; r += 1
LS.cell(r, 1, "情境 B：每 MW 年租金（租金 ÷ 平均租賃 MW）"); LS.cell(r, 2, "US$M／MW·年")
for j, y in enumerate(FYM): LS[f"{CL(5 + j)}{r}"] = f"=Thesis!{TCOL[y]}{TR['B']['rent']}/(Thesis!{TCOL[y]}{TR['B']['avg']}*(1-own_share))*1000"
LSB["permw"] = r; r += 1
LS.cell(r, 1, "情境 B：FY2027–FY2031 租金合計 ÷（在帳營業租賃負債＋未起租 35.5）"); LS.cell(r, 2, "x")
LS[f"D{r}"] = f"=SUM(E{LSB['rentB']}:I{LSB['rentB']})/(opl_q3+35.5)"; LSB["ratio"] = r
for rr_ in range(5, r + 1):
    for c in "DEFGHIJ": LS[f"{c}{rr_}"].number_format = "#,##0.00"
LS.column_dimensions["A"].width = 62

# =====================================================================================
# Val：DCF 50%＋Comparable 50%
# =====================================================================================
V = wb.create_sheet("Val")
V["A1"] = "Val — 目標價（現價日＋12 個月）＝DCF（FCFF、WACC；DCF 標準規格）50%＋Comparable（新雲同業前瞻 EV/EBITDA）50%；情境權重 A 25%／B 50%／C 25%（判斷預設，Andy 可改）"
V["A2"] = "評價日 2026-09-30（Q3 資產負債表）。DCF：NOPAT＝營業利益 − 長期稅率 × MAX(0, 營業利益)（股票薪酬視為現金成本）＋折舊 − 資本支出 − 營運資金增加＋預付淨額；永續 NOPAT ×(1 − g ÷ RONIC) ÷ (WACC − g)，RONIC＝WACC；股數＝評價日股數＋RSU／選擇權淨股數，未來瀑布新股不進 DCF 每股（X35）。Comparable 錨定 FY2027 年底（含當時的淨負債與股數）"
hdr(V, 3, ["項目", "單位", "數值", "", "", "說明"])
VR = {}
def vrow(k, lab, unit, val, fmt="#,##0.00", note=None):
    global r
    VR[k] = r; V.cell(r, 1, lab); V.cell(r, 2, unit)
    if val is not None: V.cell(r, 3, val).number_format = fmt
    if note: V.cell(r, 6, note)
    r += 1
r = 4
vrow("px", "現價", "US$", "=px", note='="In 頁 px（"&TEXT(px_date,"yyyy-mm-dd")&" 收盤）；v4.13 為 2026-11-05 收盤 43.66"')
vrow("ke", "股權成本 ke（CAPM：rf＋β × ERP）", "%", "=rf+beta*erp", "0.00%", '="rf "&TEXT(rf,"0.00%")&"（美債 10 年期 2026-11-10）、β "&TEXT(beta,"0.00")&"（自算產業 β "&TEXT(beta_u,"0.000")&" 依市值 D/E 重新加槓桿）、ERP "&TEXT(erp,"0%")')
vrow("kdat", "稅後債務成本", "%", "=kd*(1-tax_lt)", "0.00%", '="kd "&TEXT(kd,"0.000%")&"（市場：2032 年 9.625% 票據）；公司稱加權 "&TEXT(kd_co,"0.0%")&"（Interested-party，只作對照）"')
vrow("E", "股權市值（現價 × 評價日股數）", "US$B", "=px*(sh_out+rsu_net)", "#,##0.0")
vrow("D", "負債（借款帳面值＋融資租賃）", "US$B", "=debt_q3+finl_q3", "#,##0.0")
vrow("wacc", "WACC（＋wacc_adj，敏感度用）", "%", f"=(C{r - 2}*C{r - 4}+C{r - 1}*C{r - 3})/(C{r - 2}+C{r - 1})+wacc_adj", "0.00%", "市值權重（A5）")
vrow("ronic", "永續期新投資報酬率 RONIC", "%", f"=C{r - 1}+ronic_adj", "0.00%", "RONIC＝WACC（DCF 標準規格）")
vrow("nd", "淨負債（借款＋融資租賃 − 現金；受限現金不扣）", "US$B", "=debt_q3+finl_q3-cash_q3", "#,##0.0")
vrow("sh", "評價日股數（流通＋RSU／選擇權淨股數）", "十億股", "=sh_out+rsu_net", "0.000")
vrow("t_tp", "折現期間：目標價時點距評價日", "年", "=(tp_date-val_date)/365.25", "0.000")
vrow("t_hold", "持有期間：目標價時點距現價日", "年", "=(tp_date-px_date)/365.25", "0.000", '="現價日 "&TEXT(px_date,"yyyy-mm-dd")&" → 目標價時點 "&TEXT(tp_date,"yyyy-mm-dd")&"；上檔空間為此期間報酬，非年化（原則第 10 節第 5 條）"')
vrow("t27", "Comparable 錨定日距目標價時點", "年", "=(fy27_end-tp_date)/365.25", "0.000")
for nm, k in (("ke", "ke"), ("wacc", "wacc"), ("ronic", "ronic"), ("netdebt0", "nd"), ("shv", "sh"), ("t_tp", "t_tp"), ("t27", "t27")):
    dn(nm, f"Val!$C${VR[k]}")
r += 1
V.cell(r, 1, "0.1 DCF（FCFF；評價日 2026-09-30；期中折現）").font = BOLD; r += 1
hdr(V, r, ["項目", "折現年數 t", "情境 A", "情境 B", "情境 C", "說明"]); r += 1
DC = {}
TT = {2026: 0.125, 2027: 0.75, 2028: 1.75, 2029: 2.75, 2030: 3.75, 2031: 4.75, 2032: 5.75}
for y in range(2026, 2033):
    V.cell(r, 1, f"  FCFF FY{y}" + ("（Q4）" if y == 2026 else "")); V.cell(r, 2, TT[y]).number_format = "0.000"
    for j, s in enumerate(SC):
        T = TR[s]; C_ = CR_[s]; tc_ = "D" if y == 2026 else TCOL[y]; cc = yc(y)
        f = (f"=Thesis!{tc_}{T['ebit']}-tax_lt*MAX(0,Thesis!{tc_}{T['ebit']})+Thesis!{tc_}{T['dep']}-Thesis!{tc_}{T['cx']}"
             f"+CF!{cc}{C_['ppin']}-CF!{cc}{C_['pprec']}-CF!{cc}{C_['nwc']}")
        V.cell(r, 3 + j, f).number_format = "#,##0.00"
    DC[y] = r; r += 1
V.cell(r, 1, "  永續期 NOPAT 基礎（FY2032 調整後 EBITDA − SBC − 穩態汰換與機房折舊；稅後）")
for j, s in enumerate(SC):
    T = TR[s]
    V.cell(r, 3 + j, f"=(K{T['ebitda']}-K{T['sbc']}-K{T['conn']}*(cx_vr200/life_ai+own_share*fac_cx/fac_life)/1000)*(1-tax_lt)".replace("K", "Thesis!K")).number_format = "#,##0.00"
V.cell(r, 6, "穩態折舊＝在役 MW × VR200 每 MW 資本支出 ÷ 年限（＝穩態汰換資本支出）"); DC["tvb"] = r; r += 1
V.cell(r, 1, "  永續價值（FY2032 年底；NOPAT ×(1＋g)×(1 − g ÷ RONIC) ÷ (WACC − g)）"); V.cell(r, 2, 6.25)
for j in range(3):
    cc = CL(3 + j); V.cell(r, 3 + j, f"={cc}{DC['tvb']}*(1+tp_g)*(1-tp_g/ronic)/(wacc-tp_g)").number_format = "#,##0.0"
DC["tv"] = r; r += 1
V.cell(r, 1, "企業價值（評價日）")
for j in range(3):
    cc = CL(3 + j)
    V.cell(r, 3 + j, "=" + "+".join(f"{cc}{DC[y]}/(1+wacc)^$B${DC[y]}" for y in range(2026, 2033)) + f"+{cc}{DC['tv']}/(1+wacc)^$B${DC['tv']}").number_format = "#,##0.0"
DC["ev"] = r; r += 1
V.cell(r, 1, "  其中：永續價值現值占企業價值（企業價值 ≤0 為 N/M；原則 56）")
for j in range(3):
    cc = CL(3 + j); V.cell(r, 3 + j, f'=IF({cc}{DC["ev"]}<=0,"N/M",{cc}{DC["tv"]}/(1+wacc)^$B${DC["tv"]}/{cc}{DC["ev"]})').number_format = "0.0%"
DC["tvsh"] = r; r += 1
V.cell(r, 1, "  檢查：DCF 有效（企業價值 >0 且 0 ≤ 永續占比 < 100%）")
for j in range(3):
    cc = CL(3 + j); V.cell(r, 3 + j, f'=IF(AND({cc}{DC["ev"]}>0,ISNUMBER({cc}{DC["tvsh"]}),{cc}{DC["tvsh"]}>=0,{cc}{DC["tvsh"]}<1),"通過","不通過（DCF 失效）")')
V.cell(r, 6, "原則第 10 節第 56 條：分母 ≤0 一律不通過；永續占比 >80% 時報告標示"); DC["valid"] = r; r += 1
V.cell(r, 1, "NOL 稅盾價值（NOL × 稅率，以 WACC 折現 3 年；Assumed）")
for j in range(3): V.cell(r, 3 + j, "=nol0*tax_lt/(1+wacc)^3").number_format = "#,##0.00"
DC["nol"] = r; r += 1
V.cell(r, 1, "股權價值（企業價值 − 淨負債＋NOL 稅盾）")
for j in range(3):
    cc = CL(3 + j); V.cell(r, 3 + j, f"={cc}{DC['ev']}-netdebt0+{cc}{DC['nol']}").number_format = "#,##0.0"
DC["eqv"] = r; r += 1
V.cell(r, 1, "DCF 每股價值（評價日；股權價值 <0 取 0＝有限責任）")
for j in range(3):
    cc = CL(3 + j); V.cell(r, 3 + j, f"=MAX(0,{cc}{DC['eqv']}/shv)").number_format = "#,##0.00"
DC["ps"] = r; r += 1
V.cell(r, 1, "DCF 情境價值推到目標價時點（×(1＋ke)^t；無股利）")
for j in range(3):
    cc = CL(3 + j); c_ = V.cell(r, 3 + j, f"={cc}{DC['ps']}*(1+ke)^t_tp"); c_.number_format = "#,##0.00"; c_.fill = KEY
DC["v12"] = r; r += 2
# Comparable
V.cell(r, 1, "0.2 Comparable（新雲同業 NTM EV/EBITDA；同業估值日 2026-11-11）").font = BOLD; r += 1
PEERH = r
hdr(V, r, ["同業", "代號", "EV（US$B）", "NTM EV/EBITDA", "估值日", "口徑與來源", "NTM EBITDA 率", "淨負債 ÷ NTM EBITDA", "資本支出 ÷ 營收", "營收成長（下一年）"]); r += 1
P0 = r
for nm, tk, ev, eb, src, mg, nd, cx_, gr in (
        ("Nebius", "NBIS", "pn_px*pn_sh+IF(peer_nd_mode=1,pn_ndf,pn_nd)", "pn_eb", "US GAAP；MarketScreener 共識（擷取 2026-11-13）；Derived＝EV ÷ NTM EBITDA", "=pn_eb/pn_rev", "=IF(peer_nd_mode=1,pn_ndf,pn_nd)/pn_eb", "未取得", "=6.556/1.348-1"),
        ("Oracle（混合業務）", "ORCL", "po_px*po_sh+IF(peer_nd_mode=1,po_ndf,po_nd)", "po_eb", "US GAAP；MarketScreener 共識（擷取 2026-11-13）；SaaS 與授權占多數", "=po_eb/po_rev", "=IF(peer_nd_mode=1,po_ndf,po_nd)/po_eb", "=po_cx/po_rev", "=74.148/52.09-1"),
        ("IREN", "IREN", "ir_px*ir_sh+IF(peer_nd_mode=1,ir_ndf,ir_nd)", "ir_eb", "US GAAP（10-K）；MarketScreener 共識（擷取 2026-11-13）；淨負債為 2026-09-30 申報口徑：借款含融資租賃 10.208 − 現金 3.522＝6.686（受限現金 2.256 不扣，與 CoreWeave 一致；公司自算 4.430 有扣受限現金）", None, "=IF(peer_nd_mode=1,ir_ndf,ir_nd)/ir_eb", "未取得", "=5.123/1.395-1")):
    V.cell(r, 1, nm); V.cell(r, 2, tk)
    if ev:
        V.cell(r, 3, "=" + ev).number_format = "#,##0.0"; V.cell(r, 4, f'=IF(ir_use=0,"不納入",({ev})/{eb})' if tk == "IREN" else f"=({ev})/{eb}").number_format = "0.00"
    else:
        V.cell(r, 3, "未取得"); V.cell(r, 4, "未取得")
    V.cell(r, 5, PEER_D); V.cell(r, 6, src)
    for j, v in enumerate((mg, nd, cx_, gr)):
        if v is not None:
            c_ = V.cell(r, 7 + j, v)
            if isinstance(v, str) and v.startswith("="): c_.number_format = "0.0%" if j != 1 else "0.00"
    r += 1
P1 = r - 1
V.cell(r, 1, "  估值日 2026-11-11（＝現價日）；NTM 自 11-11 起依月份加權（原則 49）；淨負債主口徑＝最新申報＋已發生事件（同 Oracle 樣板；Nebius、IREN 為 2026-09-30，受限現金不扣，與 CoreWeave 自身一致），替代口徑見下兩列。WhiteFiber：共識無 EBITDA（未納入）。同業 3 家，少於 8 家不做迴歸（原則 7）").font = GREY; r += 1
V.cell(r, 1, "  替代口徑：淨負債改用 MarketScreener 當年度年底預估時的 NTM EV/EBITDA（NBIS／ORCL／IREN；時點不一致：11-11 市值配年底預估淨負債，只作對照）")
for j, (pxs, nds, ebs) in enumerate((("pn_px*pn_sh", "pn_ndf", "pn_eb"), ("po_px*po_sh", "po_ndf", "po_eb"), ("ir_px*ir_sh", "ir_ndf", "ir_eb"))):
    V.cell(r, 3 + j, f"=({pxs}+{nds})/{ebs}").number_format = "0.00"
VR["altm"] = r; r += 1
V.cell(r, 1, "  替代口徑中位數（對照；時點不一致；Sens peer_nd_mode）"); V.cell(r, 4, f"=MEDIAN(C{r - 1}:E{r - 1})").number_format = "0.00"; VR["altmed"] = r; r += 1
V.cell(r, 1, "  對照（不進中位數）：CoreWeave 自身＝現價市值＋淨負債 ÷ 2027 年共識 EBITDA（財報前共識）；後四欄為模型情境 B FY2027").font = GREY
V.cell(r, 3, "=px*shv+netdebt0").number_format = "#,##0.0"; V.cell(r, 4, f"=C{r}/cons_eb27").number_format = "0.00"; V.cell(r, 5, "2026-11-11")
VR["selfm"] = r; SELF_R = r
for j, f_ in enumerate((f"=PL!E{PR_['B']['aebm']}", f"=netdebt0/PL!E{PR_['B']['aeb']}", f"=Thesis!F{TR['B']['cx']}/PL!E{PR_['B']['rev']}", f"=PL!E{PR_['B']['revg']}")):
    V.cell(r, 7 + j, f_).number_format = "0.00" if j == 1 else "0.0%"
r += 1
V.cell(r, 1, "同業中位數（前瞻 EV/EBITDA）"); V.cell(r, 4, f"=MEDIAN(D{P0}:D{P1})").number_format = "0.00"; VR["mult"] = r; dn("mult", f"Val!$D${r}"); r += 2
hdr(V, r, ["項目", "單位", "情境 A", "情境 B", "情境 C", "說明"]); r += 1
CO = {}
def vr3(k, lab, unit, f, fmt="#,##0.00", note=None):
    global r
    CO[k] = r; V.cell(r, 1, lab); V.cell(r, 2, unit)
    for j, s in enumerate(SC): V.cell(r, 3 + j, f(s, CL(3 + j))).number_format = fmt
    if note: V.cell(r, 6, note)
    r += 1
vr3("eb", "FY2028 調整後 EBITDA（FY2027 年底的前瞻一年）", "US$B", lambda s, c: f"=PL!F{PR_[s]['aeb']}")
vr3("ev0", "企業價值（未調整：× 同業中位數）", "US$B", lambda s, c: f"={c}{CO['eb']}*mult*mult_adj", "#,##0.0")
vr3("ev", "企業價值（× 相對倍數係數 rel_mult）", "US$B", lambda s, c: f"={c}{CO['ev0']}*rel_mult", "#,##0.0", '="rel_mult "&TEXT(rel_mult,"0.00")&"（未調整：缺 CoreWeave 10 年歷史倍數）"')
vr3("nd27", "淨負債（FY2027 年底：借款＋融資租賃 − 現金）", "US$B", lambda s, c: f"=CF!E{CR_[s]['debt']}+finl_q3-CF!E{CR_[s]['cash1']}", "#,##0.0")
vr3("sh27", "股數（FY2027 年底，含瀑布新股＋RSU／選擇權）", "十億股", lambda s, c: f"=CF!E{CR_[s]['sh1']}+rsu_net", "0.000")
vr3("ps0", "Comparable 每股（未調整；FY2027 年底）", "US$", lambda s, c: f"=MAX(0,({c}{CO['ev0']}-{c}{CO['nd27']})/{c}{CO['sh27']})")
vr3("ps", "Comparable 每股（FY2027 年底）", "US$", lambda s, c: f"=MAX(0,({c}{CO['ev']}-{c}{CO['nd27']})/{c}{CO['sh27']})")
vr3("v120", "Comparable 情境價值（未調整；折回目標價時點）", "US$", lambda s, c: f"={c}{CO['ps0']}/(1+ke)^t27")
vr3("v12", "Comparable 情境價值（折回目標價時點）", "US$", lambda s, c: f"={c}{CO['ps']}/(1+ke)^t27")
for j in range(3): V.cell(CO["v12"], 3 + j).fill = KEY
r += 1
V.cell(r, 1, "0.3 結論").font = BOLD; r += 1
hdr(V, r, ["項目", "單位", "數值", "對照", "", "說明"]); r += 1
CON = {}
def crow(k, lab, unit, val, fmt="#,##0.00", note=None, d=None):
    global r
    CON[k] = r; V.cell(r, 1, lab); V.cell(r, 2, unit)
    if val is not None:
        x_ = V.cell(r, 3, val)
        if fmt: x_.number_format = fmt
    if d is not None: V.cell(r, 4, d).number_format = fmt
    if note: V.cell(r, 6, note)
    r += 1
sw = lambda row: f"wA*C{row}+wB*D{row}+wC*E{row}"
crow("tps", "情境目標價（加權 DCF 與 Comparable；A／B／C）", "US$", None)
for j in range(3):
    cc = CL(3 + j); V.cell(CON["tps"], 3 + j, f"=w_dcf*{cc}{DC['v12']}+(1-w_dcf)*{cc}{CO['v12']}").number_format = "#,##0.00"
crow("dcf", "DCF 目標價（情境加權）", "US$", "=" + sw(DC["v12"]))
crow("comp", "Comparable 目標價（情境加權）", "US$", "=" + sw(CO["v12"]))
crow("comp0", "Comparable 目標價（未調整，情境加權）", "US$", "=" + sw(CO["v120"]))
crow("tp", "目標價（DCF 與 Comparable 加權）", "US$", f"=w_dcf*C{CON['dcf']}+(1-w_dcf)*C{CON['comp']}", note='="DCF "&TEXT(w_dcf,"0%")&"＋Comparable "&TEXT(1-w_dcf,"0%")&"；情境 "&TEXT(wA,"0%")&"／"&TEXT(wB,"0%")&"／"&TEXT(wC,"0%")')
V.cell(CON["tp"], 3).fill = KEY
crow("px", "現價", "US$", "=px")
crow("up", "上檔空間：加權目標價", "%", f"=C{CON['tp']}/C{CON['px']}-1", "0.0%")
crow("upc", "上檔空間：Comparable 目標價", "%", f"=C{CON['comp']}/C{CON['px']}-1", "0.0%")
RATE = lambda x: f'IF({x}>0.5,"強力買進",IF({x}>0.25,"買進",IF({x}>=-0.25,"中立",IF({x}>=-0.5,"放空","強力放空"))))'
crow("valrec", "估值評等（只看上檔空間；未套收支關卡）", "—", "=" + RATE(f"C{CON['up']}"), None, "五級規則：上檔空間 > +50% 強力買進、> +25% 買進、≥ −25% 中立、≥ −50% 放空、其餘強力放空（原則 24、57）")
crow("gatev", "入不敷出判定用：情境 B 的 FY2026–FY2030 自由現金流合計（＝Fund 頁收支關卡）", "US$B", f"=Fund!D{FS['gate']}", "#,##0.0", "原則 63：合計 < 0＝入不敷出＝收支關卡不通過（判準與 AI 投資地圖相同；本模型無容忍帶）")
crow("gatep", "入不敷出判定結果（原則 63）", "—", f'=IF(C{CON["gatev"]}<0,"不通過","通過")', None)
crow("rec", "投資建議", "—", f'=IF(AND(C{CON["gatep"]}="不通過",OR(C{CON["valrec"]}="強力買進",C{CON["valrec"]}="買進")),"收支不足",C{CON["valrec"]})', None)
V.cell(CON["rec"], 6, "先看收支、再看估值（原則 63；Andy 2026-10-25、10-26、11-13）：收支關卡不通過時，估值評等為買進以上者一律改列「收支不足」，中立、放空、強力放空維持；估值評等五級：上檔空間 > +50% 強力買進、> +25% 買進、≥ −25% 中立、≥ −50% 放空、其餘強力放空")
crow("dist", "距最近級距邊界（百分點；原則 36）", "pt", f"=MIN(ABS(C{CON['up']}-0.5),ABS(C{CON['up']}-0.25),ABS(C{CON['up']}+0.25),ABS(C{CON['up']}+0.5))", "0.0%")
crow("dcfup", "上檔空間歸因：DCF 貢獻（w_dcf ×(DCF ÷ 現價 − 1)）", "%", f"=w_dcf*(C{CON['dcf']}/C{CON['px']}-1)", "0.0%", "原則 47：三段相加＝加權上檔空間")
crow("cmpup", "上檔空間歸因：未調整 Comparable 貢獻", "%", f"=(1-w_dcf)*(C{CON['comp0']}/C{CON['px']}-1)", "0.0%")
crow("adjup", "上檔空間歸因：相對倍數調整貢獻", "%", f"=(1-w_dcf)*(C{CON['comp']}-C{CON['comp0']})/C{CON['px']}", "0.0%")
crow("upchk", "檢查：三段相加 − 上檔空間（應為 0）", "%", f"=C{CON['dcfup']}+C{CON['cmpup']}+C{CON['adjup']}-C{CON['up']}", "0.000%")
crow("lo", "情境區間下緣（三情境 MIN；原則 58）", "US$", f"=MIN(C{CON['tps']}:E{CON['tps']})")
crow("hi", "情境區間上緣（三情境 MAX）", "US$", f"=MAX(C{CON['tps']}:E{CON['tps']})")
crow("inr", "檢查：情境 B 落在區間內", "—", f'=IF(AND(D{CON["tps"]}>=C{CON["lo"]},D{CON["tps"]}<=C{CON["hi"]}),"通過","不通過")', None)
crow("dvalid", "DCF 有效性（三情境）", "—", f'=IF(COUNTIF(C{DC["valid"]}:E{DC["valid"]},"通過")=3,"三情境皆通過","有情境 DCF 失效："&IF(C{DC["valid"]}<>"通過","A ","")&IF(D{DC["valid"]}<>"通過","B ","")&IF(E{DC["valid"]}<>"通過","C",""))', None)
crow("npvAB", "每增加 1 GW 的增量 DCF 企業價值：A 對 B（原則 60）", "US$B／GW", f"=(C{DC['ev']}-D{DC['ev']})/((mwA_30-mwB_30)/1000)", "#,##0.00", "2030 年底在役 GW 差；負值＝擴張毀損價值")
crow("npvBC", "每增加 1 GW 的增量 DCF 企業價值：B 對 C", "US$B／GW", f"=(D{DC['ev']}-E{DC['ev']})/((mwB_30-mwC_30)/1000)", "#,##0.00")
crow("npvps", "每增加 1 GW 的增量情境目標價：A 對 B（US$／股）", "US$", f"=(C{CON['tps']}-D{CON['tps']})/((mwA_30-mwB_30)/1000)", "#,##0.00")
crow("cons", "市場共識目標價（2026-09-25，財報前，只作歷史對照）", "US$", "=cons_tp")
crow("v413", "v4.13 加權目標價（舊法；現價 43.66）", "US$", "=tp_v413")
crow("chg", "目標價變動：v5.0 − v4.13", "US$", f"=C{CON['tp']}-C{CON['v413']}")
crow("tvsh", "情境 B：永續價值占企業價值", "%", f"=D{DC['tvsh']}", "0.0%")
# 橫幅與白話字串（公式組字；原則 13、31、57、60）
crow("bn_scn", "橫幅：擴張與目標價", "—", f'=IF(AND(C{CON["tps"]}<D{CON["tps"]},D{CON["tps"]}<E{CON["tps"]}),"A 情境擴張最多但目標價最低："&TEXT(C{CON["tps"]},"0.00")&"／"&TEXT(D{CON["tps"]},"0.00")&"／"&TEXT(E{CON["tps"]},"0.00")&"（A／B／C）；每增加 1 GW（A 對 B）DCF 企業價值變動 "&TEXT(C{CON["npvAB"]},"+0.0;-0.0")&" 十億美元",IF(AND(C{CON["tps"]}>D{CON["tps"]},D{CON["tps"]}>E{CON["tps"]}),"擴張越多目標價越高："&TEXT(C{CON["tps"]},"0.00")&"／"&TEXT(D{CON["tps"]},"0.00")&"／"&TEXT(E{CON["tps"]},"0.00")&"（A／B／C）","情境目標價 "&TEXT(C{CON["tps"]},"0.00")&"／"&TEXT(D{CON["tps"]},"0.00")&"／"&TEXT(E{CON["tps"]},"0.00")&"（A／B／C；非單調）")&IF(C{CON["npvAB"]}<0,"；但 DCF 口徑每增加 1 GW（A 對 B）企業價值 "&TEXT(C{CON["npvAB"]},"+0.0;-0.0")&" 十億美元、（B 對 C）"&TEXT(C{CON["npvBC"]},"+0.0;-0.0")&"：擴張價值為負，情境目標價的差異來自 Comparable 的 EBITDA 倍數",""))', None)
UB = f'IF(C{CON["up"]}>0.5,NA(),IF(C{CON["up"]}>0.25,0.5,IF(C{CON["up"]}>=-0.25,0.25,IF(C{CON["up"]}>=-0.5,-0.25,-0.5))))'
LB = f'IF(C{CON["up"]}>0.5,0.5,IF(C{CON["up"]}>0.25,0.25,IF(C{CON["up"]}>=-0.25,-0.25,IF(C{CON["up"]}>=-0.5,-0.5,NA()))))'
MREQ = lambda b, cond: (f'=IF({cond},"—",((px*(1+{b})-w_dcf*C{CON["dcf"]})/(1-w_dcf)*(1+ke)^t27+wA*C{CO["nd27"]}/C{CO["sh27"]}+wB*D{CO["nd27"]}/D{CO["sh27"]}+wC*E{CO["nd27"]}/E{CO["sh27"]})'
                  f'/((wA*C{CO["eb"]}/C{CO["sh27"]}+wB*D{CO["eb"]}/D{CO["sh27"]}+wC*E{CO["eb"]}/E{CO["sh27"]})*mult_adj*rel_mult))')
crow("mup", "升一級所需同業倍數（其他不變；Comparable 每股未觸及 0 下限時）", "x", MREQ("0.5" if False else UB.replace("NA()","0.5"), f'C{CON["up"]}>0.5'), "0.00", "原則 36：反解加權目標價＝升一級邊界")
crow("mdn", "降一級所需同業倍數", "x", MREQ(LB.replace("NA()","-0.5"), f'C{CON["up"]}<-0.5'), "0.00", "反解加權目標價＝降一級邊界（現行邊界）")
crow("bn_edge", "橫幅：距級距邊界", "—", f'=IF(C{CON["dist"]}<0.05,"接近級距邊界：","")&"距最近級距邊界 "&TEXT(C{CON["dist"]}*100,"0.0")&" 個百分點；同業倍數 "&TEXT(mult,"0.00")&" 倍，升一級需 "&IF(ISNUMBER(C{CON["mup"]}),TEXT(C{CON["mup"]},"0.00")&" 倍","—（已最高級）")&"、降一級需 "&IF(ISNUMBER(C{CON["mdn"]}),TEXT(C{CON["mdn"]},"0.00")&" 倍","—")', None)
crow("bn_dcf", "橫幅：DCF 有效性", "—", f'=C{CON["dvalid"]}', None)
crow("val_lab", "估值評等標示（距邊界 <5pt 時加「邊界」）", "—", f'=C{CON["valrec"]}&IF(C{CON["dist"]}<0.05,"（邊界，±"&TEXT(C{CON["dist"]}*100,"0.0")&"pt）","")', None)
crow("rec_lab", "評等標示（收支不足時並列估值評等與收支合計）", "—", f'=IF(C{CON["rec"]}="收支不足","收支不足〔估值評等："&C{CON["val_lab"]}&"；情境 B 的 FY2026–FY2030 自由現金流合計 "&TEXT(C{CON["gatev"]},"0.0")&" 十億美元，入不敷出〕",C{CON["val_lab"]})', None)
_flip = []
if SENS:
    _brec = SENS["cases"][0]["lo"]["rec"]
    _LOW = ("中立", "放空", "強力放空")
    _flip = [4 + i for i, c in enumerate(SENS["cases"]) if any(c.get(sd) and c[sd]["rec"] in _LOW for sd in ("lo", "hi"))]
_SR = [4, 3 + len(SENS['cases'])] if SENS else [4, 4]
_med = f'IF(ABS(mult-D{P0})<1E-9,A{P0},IF(ABS(mult-D{P0 + 1})<1E-9,A{P0 + 1},A{P0 + 2}))'
crow("bn_frag", "橫幅：評等脆弱點（原則 36）", "—", '=IF(C{d}<0.05,"評等脆弱點：(a) 同業中位數剛好是 "&{m}&"（"&TEXT(mult,"0.00")&" 倍）；(b) Nebius 淨負債用 11-12 公布的 9-30 申報（晚估值日一天，屬已發生事件的最新申報口徑）；(c) Sens 單一案例中有 "&TEXT({n},"0")&" 案的估值評等會由買進降為中立或更低"{lst},"")'.format(
    d=CON["dist"], m=_med, n=f'SUMPRODUCT(--((ISNUMBER(MATCH(Sens!M{_SR[0]}:M{_SR[1]},{{"中立","放空","強力放空"}},0))+ISNUMBER(MATCH(Sens!N{_SR[0]}:N{_SR[1]},{{"中立","放空","強力放空"}},0)))>0))' if SENS else "0", lst=('&"（如 "&' + '&"、"&'.join(f"Sens!L{x}" for x in _flip[:6]) + ('&" 等）"' if len(_flip) > 6 else '&"）"')) if _flip else ''), None)
V.column_dimensions["A"].width = 64; V.column_dimensions["B"].width = 11
for c in "CDE": V.column_dimensions[c].width = 12
V.column_dimensions["F"].width = 80

# =====================================================================================
# RevDCF：反向 DCF（整本重算反解；scripts/reverse_v50.py）
# =====================================================================================
RV = wb.create_sheet("RevDCF")
RV["A1"] = "RevDCF — 反向 DCF：一次動一個驅動、整本重算，反解「加權目標價＝目標值」所需的值（scripts/reverse_v50.py → data/reverse_v50.json；建置時寫入）"
RV["A2"] = "主導變數＝每 MW 年收入（k 調整係數 k_adj）：其他輸入不變，反解加權目標價＝現價、＝升一級邊界、＝降一級邊界所需的 k_adj，換算成 FY2030 情境 B 每 MW 年收入，與模型基準、Tokenomics 錨與公司宣稱新約價格並列（J4）。反解值是目標價等於該值的點，不是評等門檻（原則 4）"
RV["A3"] = "讀法：支持現價所需的每 MW 年收入若落在公司宣稱區間內，表示現價隱含公司說法成立；本模型基準以 Tokenomics 持有成本 × k 為錨，不採用公司宣稱值（J4、CLAUDE.md 第 14 條）"
hdr(RV, 4, ["反解目標", "驅動", "目標加權目標價", "反解 k_adj", "", "", "FY2030 每 MW 年收入（情境 B）", "模型基準", "相對基準", "公司宣稱中點"])
RVN = {}
r = 5
if REV:
    for it in REV["solves"]:
        RV.cell(r, 1, it["label"]); RV.cell(r, 2, it["key"]); RV.cell(r, 3, it["target"]).number_format = "#,##0.00"
        RV.cell(r, 4, it["value"]).number_format = "0.000"
        RV.cell(r, 7, it["pm30"]).number_format = "0.00"; RV.cell(r, 8, f"=Thesis!I{TR['B']['pm']}").number_format = "0.00"
        RV.cell(r, 9, f"=G{r}/H{r}-1").number_format = "+0.0%;-0.0%"; RV.cell(r, 10, "=(p_claim_lo+p_claim_hi)/2").number_format = "0.00"
        RVN[it["id"]] = r; r += 1
    RV.cell(r, 1, f"快照：{REV['date']}；基準加權目標價 {REV['base_tp']:.4f}（建置時重算的快照值）"); RVN["snap"] = r; r += 1
    RV.cell(r, 1, "快照檢查：反解快照的基準加權目標價 − Val 頁現值（應為 0；不為 0＝重跑 scripts/reverse_v50.py）"); RV.cell(r, 3, f"={REV['base_tp']}-Val!C{CON['tp']}").number_format = "0.0000"; RVN["chk"] = r; r += 1
else:
    RV.cell(r, 1, "反解快照未產生（先跑 scripts/reverse_v50.py）"); RVN["none"] = r; r += 1
RV.column_dimensions["A"].width = 50
for c in "BCDGHIJ": RV.column_dimensions[c].width = 14

# =====================================================================================
# Recon：公司指引、公司實況驗證、對照
# =====================================================================================
RC = wb.create_sheet("Recon")
RC["A1"] = "Recon — 對帳：公司指引（只作對照，不校準）、公司實況驗證（CLAUDE.md 第 15 條）、J4 公司宣稱價格、Wattonomics、類比、舊法。差距 = 模型 ÷ 對照 − 1"
hdr(RC, 3, ["項目", "對照值", "模型 A", "A 差距", "模型 B", "B 差距", "模型 C", "C 差距", "說明"])
RCR = {}
r = 4
def rsec(t):
    global r
    RC.cell(r, 1, t).font = BOLD; r += 1
def rline(k, lab, ref, f, note="", fmt="#,##0.00", gap=True):
    global r
    RCR[k] = r; RC.cell(r, 1, lab)
    if ref is not None:
        x_ = RC.cell(r, 2, ref)
        if fmt: x_.number_format = fmt
    for j, s in enumerate(SC):
        v = f(s) if callable(f) else f
        if v is None: continue
        x_ = RC.cell(r, 3 + 2 * j, v)
        if fmt: x_.number_format = fmt
        if gap and ref is not None: RC.cell(r, 4 + 2 * j, f'=IF(B{r}=0,"N/M",{CL(3 + 2 * j)}{r}/B{r}-1)').number_format = "+0.0%;-0.0%"
    RC.cell(r, 9, note); r += 1
rsec("A. 公司 FY2026 指引 vs 模型（原則 50：基準落在區間外時標示並分解；指引不作界值，原則 59）")
rline("rev26", "FY2026 營收 vs 指引中點", "=(g_rev_lo+g_rev_hi)/2", lambda s: f"=PL!D{PR_[s]['rev']}", "指引 13.0–13.3；三情境 FY2026 相同（Q4 共用路徑）")
rline("rev26pos", "  FY2026 營收：相對指引區間", None, lambda s: f'=IF(PL!D{PR_[s]["rev"]}>g_rev_hi,"高於指引上緣 "&TEXT((PL!D{PR_[s]["rev"]}/g_rev_hi-1)*100,"0.0")&" pt",IF(PL!D{PR_[s]["rev"]}<g_rev_lo,"低於指引下緣 "&TEXT((1-PL!D{PR_[s]["rev"]}/g_rev_lo)*100,"0.0")&" pt","落在指引區間內"))', "方向字樣由公式依數值決定（原則 57）", None, False)
rline("aoi26", "FY2026 調整後營業利益 vs 指引中點", "=(g_aoi_lo+g_aoi_hi)/2", lambda s: f"=PL!D{PR_[s]['aoi']}", "模型＝9M 實績 0.528＋Q4（調整後 EBITDA − 折舊）")
rline("aoi26pos", "  FY2026 調整後營業利益：相對指引區間", None, lambda s: f'=IF(PL!D{PR_[s]["aoi"]}>g_aoi_hi,"高於指引上緣 "&TEXT((PL!D{PR_[s]["aoi"]}/g_aoi_hi-1)*100,"0.0")&" pt",IF(PL!D{PR_[s]["aoi"]}<g_aoi_lo,"低於指引下緣 "&TEXT((1-PL!D{PR_[s]["aoi"]}/g_aoi_lo)*100,"0.0")&" pt","落在指引區間內"))', "", None, False)
rline("cx26", "FY2026 資本支出（公司口徑）vs 指引中點（J5）", "=(g_cx_lo+g_cx_hi)/2", lambda s: f"=Thesis!E{TR[s]['cx']}", "模型＝9M 公司口徑 30.650（Verified＋Interested-party）＋Q4 模型公式；不設 MAX(公式, 指引下緣)")
rline("cxq4", "  Q4 2026 資本支出：模型 vs 指引隱含 Q4 中點（指引 − 9M；隱含區間 6.35–8.35）", "=(g_cx_lo+g_cx_hi)/2-cx9m_def", lambda s: f"=Thesis!D{TR[s]['cx']}", "公式值、採用值、差額並列（原則 59）：採用值＝公式值（無夾值）")
rline("cxq4pos", "  Q4 資本支出：相對指引隱含區間", None, lambda s: f'=IF(Thesis!D{TR[s]["cx"]}>g_cx_hi-cx9m_def,"高於指引隱含上緣 "&TEXT(Thesis!D{TR[s]["cx"]}-(g_cx_hi-cx9m_def),"0.00"),IF(Thesis!D{TR[s]["cx"]}<g_cx_lo-cx9m_def,"低於指引隱含下緣 "&TEXT((g_cx_lo-cx9m_def)-Thesis!D{TR[s]["cx"]},"0.00"),"落在指引隱含區間內"))', "單位：十億美元", None, False)
rline("arr", "2026 年底 ARR vs 指引中點（模型＝Q4 營收 × 4）", "=(g_arr_lo+g_arr_hi)/2", lambda s: f"=Thesis!D{TR[s]['rev']}*4", "口徑不同：公司 ARR 為年底月營收 × 12（含 12 月新上線），模型為 Q4 平均 × 4 → 只列隱含值（原則 33）")
rline("int4", "Q4 2026 利息費用 vs 指引中點", "=(g_int_lo+g_int_hi)/2", lambda s: f"=CF!D{CR_[s]['int']}", "模型＝Q3 平均利率 × Q3 借款 × 0.25；指引高出代表 Q4 舉債增加或利率上升（反解：Q4 平均利率需 "  + "見下列）")
rline("int4r", "  達到指引中點所需的年化利率（÷ Q3 借款）", "=kd_co", lambda s: "=(g_int_lo+g_int_hi)/2*4/debt_q3", "對照：公司稱加權債務成本 7.9%（Interested-party）；只展示、不校準（X3）", "0.00%")
rline("int4pos", "  Q4 利息：相對指引區間（原則 50）", None, lambda s: f'=IF(CF!D{CR_[s]["int"]}<g_int_lo,"低於指引下緣 "&TEXT(g_int_lo-CF!D{CR_[s]["int"]},"0.00"),IF(CF!D{CR_[s]["int"]}>g_int_hi,"高於指引上緣 "&TEXT(CF!D{CR_[s]["int"]}-g_int_hi,"0.00"),"落在指引區間內"))', "機制：利息以期初餘額計（Oracle 樣板慣例，不改公式），當期新增借款不計息", None, False)
rline("intnew", "  新增借款當年未計的半年利息（FY2027；只標示，不進模型）", None, lambda s: f"=0.5*(r_new*CF!E{CR_[s]['ig']}+r_junk*CF!E{CR_[s]['hy']})", "利息以平均餘額計時約多出此數；風險登錄第 6 條")
rline("aoiq4", "  Q4 調整後營業利益：模型 vs Q3 實績 0.379", "=0.379", lambda s: f"=Thesis!D{TR[s]['ebitda']}-Thesis!D{TR[s]['dep']}", "機制：Q4 折舊攤銷較 Q3 季增（下一列），在建工程 11.9 與 Q4 資本支出開始折舊")
rline("depq4", "  Q4 折舊攤銷季增率（模型 Q4 ÷ Q3 實績 1.889 − 1）", None, lambda s: f"=Thesis!D{TR[s]['dep']}/da_q3-1", "", "0.0%", False)
rline("rev4req", "  達到 FY2026 營收指引中點所需的 Q4 每 MW 年收入（÷ Q4 計費 MW；反解，只展示、不校準）", f"=Thesis!D{TR['B']['pm']}", lambda s: f"=((g_rev_lo+g_rev_hi)/2-rev9m)/(Thesis!D{TR[s]['bill']}*0.25)*1000", "原則 50：模型低於指引主要來自每 MW 年收入（模型＝Tokenomics 錨 × k；公司稱 7 月起漲價、新約價格較高，Interested-party）")
rline("rev4mw", "  或：維持模型每 MW 收入時所需的 Q4 平均計費 MW", f"=Thesis!D{TR['B']['bill']}", lambda s: f"=((g_rev_lo+g_rev_hi)/2-rev9m)/(Thesis!D{TR[s]['pme']}*0.25)*1000", "對照：公司年底在役 >1.85 GW", "#,##0")
rsec("B. RPO 與 backlog（同口徑：RPO 排程 vs 模型營收）")
rline("rpo27", "FY2027：RPO 排程 vs 模型營收", f"=RPO!E{RPB['sched']}", lambda s: f"=RPO!E{RPB['rev' + s]}", "RPO 排程為平均攤分（Interested-party 分桶）")
rline("rpo29", "FY2029：RPO 排程 vs 模型營收", f"=RPO!G{RPB['sched']}", lambda s: f"=RPO!G{RPB['rev' + s]}", "")
rsec("C. 公司實況驗證（CLAUDE.md 第 15 條：差距 >10% 找有證據的機制，找不到則 Tokenomics 為基準、公司實際為 Sens）")
rline("q9rev", "9M 2026 營收：實績 vs 模型估計（計費 MW × 每 MW 年收入 × 0.75）", "=rev9m", lambda s: f"=Thesis!C{TR[s]['revchk']}", "差距 ≤10%：Tokenomics 錨 × k 與實績一致")
rline("q3pm", "Q3 每 MW 年營收（營收 × 4 ÷ 平均在役）vs 模型（每 MW 年收入 × 計費比例）", f"=Q!D{QR['permw']}", lambda s: f"=Thesis!C{TR[s]['pm']}*u_bill", "")
rline("q3eb", "Q3 調整後 EBITDA 率：實績 vs 模型 FY2027", "=aeb_q3/rev_q3", lambda s: f"=Thesis!F{TR[s]['ebm']}", "模型 FY2027 以 Tokenomics 營運成本＋租金＋管銷；差距 >10% 時 rent_mw、sga_r 有 Sens", "0.0%")
rsec("D. J4：公司宣稱的新約價格（Interested-party；只對照、不改基準 k）")
rline("j4m", "每 MW 年收入：公司宣稱新約中點 vs 模型 FY2027（情境 B 欄＝基準）", "=(p_claim_lo+p_claim_hi)/2", lambda s: f"=Thesis!F{TR[s]['pm']}", "公司宣稱分母（IT 或設施 MW）未說明；若為設施 MW，換成 IT 口徑需 ×1.2（IF_FacilityGW 基準）")
rline("j4tk", "  Tokenomics 錨（GB300 IF_HoldEcon）vs 公司宣稱中點", "=(p_claim_lo+p_claim_hi)/2", lambda s: "=he_gb300", "")
rline("j4s", "  短天期合約（2026-09-17 新聞稿）vs 模型 FY2027", "=p_short", lambda s: f"=Thesis!F{TR[s]['pm']}", "隨需／短天期；模型以 od_share 10% × k_現貨 涵蓋")
if REV:
    rline("j4req", "  支持現價所需每 MW 年收入（FY2030 情境 B；RevDCF 反解）vs 公司宣稱中點", "=(p_claim_lo+p_claim_hi)/2", lambda s: f"=RevDCF!G{RVN['px']}" if s == "B" else None, "反解值不是評等門檻（原則 4）")
rsec("E. 類比與容量對照")
rline("iren", "IREN–Microsoft 長約每 MW 年收入（Analogy，原則 61）vs 模型 FY2027", "=p_long", lambda s: f"=Thesis!F{TR[s]['pm']}", "差異：IREN 自有機房（價格含設施）、200 MW、GB300、預付 20%；CoreWeave 多為租賃機房、多客戶")
rline("wat26", "Wattonomics 2026 年底上線 MW vs 模型（原則 54）", "=wat26", lambda s: f"=Thesis!E{TR[s]['conn']}", "wattonomics_snapshot_v0.3 mw_crwv", "#,##0")
rline("wat30", "Wattonomics 2030 年底上線 MW vs 模型", "=wat30", lambda s: f"=Thesis!I{TR[s]['conn']}", "差距 >15%：寫對帳單草稿（coreweave/docs/reports/20261112_對帳單草稿_CoreWeave_J4與Wattonomics.md）", "#,##0")
rsec("F. 舊法與舊版（只作對照）")
rline("old", "v4.13 融資調整 DCF（舊法，不進結論）vs v5.0 加權目標價", "=dcf_v413", lambda s: f"=Val!C{CON['tps']}".replace("C", "CDE"[SC.index(s)], 1) if False else f"=Val!{'CDE'[SC.index(s)]}{CON['tps']}", "B1：舊法保留為對照")
rline("cons", "市場共識 FY2027 EBITDA（財報前）vs 模型", "=cons_eb27", lambda s: f"=PL!E{PR_[s]['aeb']}", "財報前，只作歷史對照（原則 15）")
r += 1
RC.cell(r, 1, "G. 圖表資料：每 MW 年收入（情境 B；US$M／MW·年；HTML 主題圖一）").font = BOLD; r += 1
RCR["chart_h"] = r
for j, y in enumerate(FYM): RC.cell(r, 3 + j, f"FY{y}").font = BOLD
r += 1
for k, lab, f in (("ch_m", "模型：Tokenomics 錨 × k（基準）", lambda y: f"=Thesis!{TCOL[y]}{TR['B']['pm']}"),
                  ("ch_tk", "Tokenomics 錨：在役世代加權 IF_HoldEcon", lambda y: f"=Thesis!{TCOL[y]}{TR['B']['anc']}"),
                  ("ch_cl", "公司宣稱新約中點（Interested-party）", lambda y: "=(p_claim_lo+p_claim_hi)/2")) + ((("ch_req", "支持現價所需（RevDCF 反解 k_adj × 模型）", lambda y: f"=Thesis!{TCOL[y]}{TR['B']['pm']}*RevDCF!D{RVN['px']}"),) if REV else ()):
    RCR[k] = r; RC.cell(r, 1, lab)
    for j, y in enumerate(FYM): RC.cell(r, 3 + j, f(y)).number_format = "0.00"
    r += 1
RC.column_dimensions["A"].width = 70; RC.column_dimensions["I"].width = 70
for c in "BCDEFGH": RC.column_dimensions[c].width = 11

# =====================================================================================
# Sens
# =====================================================================================
SE = wb.create_sheet("Sens")
SE["A1"] = "Sens — 一次動一個輸入（低／高），整本重建並以 LibreOffice 重算，讀加權目標價、DCF 單法目標價、收支關卡（scripts/sens_v50.py → data/sens_v50.json；建置時寫入此頁）"
SE["A2"] = ("重算日期 " + SENS["date"] + f"；基準加權目標價 {SENS['base']['tp']:.4f}" if SENS else "敏感度快照未產生（先跑 scripts/sens_v50.py）") + "。In 頁每個有低／高區間的鍵都有案例（L 欄「鍵」）；情境路徑 mw*_xx 由 Val 頁情境目標價呈現。"
hdr(SE, 3, ["輸入", "低／高設定", "低：加權目標價", "低：相對基準", "高：加權目標價", "高：相對基準", "最大絕對變動", "低：DCF 目標價", "高：DCF 目標價", "低：收支關卡（B）", "高：收支關卡（B）", "鍵", "低：估值評等", "高：估值評等"])
SENS_R = [4, 4]
r = 4
if SENS:
    bt = SENS["base"]["tp"]
    for c in SENS["cases"]:
        SE.cell(r, 1, c["label"]); SE.cell(r, 2, c["setting"]); SE.cell(r, 12, c["key"])
        for col, side in ((3, "lo"), (5, "hi")):
            v = c.get(side)
            if v is None: continue
            SE.cell(r, col, v["tp"]).number_format = "#,##0.00"; SE.cell(r, col + 1, v["tp"] / bt - 1).number_format = "+0.0%;-0.0%"
            SE.cell(r, 8 if side == "lo" else 9, v["dcf"]).number_format = "#,##0.00"
            SE.cell(r, 10 if side == "lo" else 11, v["gate"]).number_format = "#,##0.0"
            SE.cell(r, 13 if side == "lo" else 14, v["rec"])
        SE.cell(r, 7, max(abs((c[s_]["tp"] / bt - 1)) for s_ in ("lo", "hi") if c.get(s_))).number_format = "0.0%"
        r += 1
    SENS_R = [4, r - 1]
    r += 1
    SE.cell(r, 1, "快照檢查：data/sens_v50.json 的基準加權目標價 − Val 頁現值（應為 0；不為 0＝Sens 已過期，重跑 scripts/sens_v50.py）"); SE.cell(r, 3, f"={bt}-Val!C{CON['tp']}").number_format = "0.000000"; SNAPR = r; r += 2
    SE.cell(r, 1, "WACC × 永續成長率 g：加權目標價（整本重算；列＝WACC 調整、欄＝g）").font = BOLD; GRID0 = r; r += 1
    hdr(SE, r, ["WACC 調整", "", "g＝2%", "g＝3%", "g＝4%"]); r += 1
    for lab, row in zip(("WACC −1pt", "WACC 基準", "WACC +1pt"), SENS["grid"]):
        SE.cell(r, 1, lab)
        for j, v in enumerate(row): SE.cell(r, 3 + j, v).number_format = "#,##0.00"
        r += 1
    GRID = (GRID0, r - 1); r += 1
else:
    SNAPR = r; SE.cell(r, 1, "快照檢查：未產生"); r += 2; GRID = (r, r)
SE.cell(r, 1, "說明：tornado 以「低：相對基準」「高：相對基準」兩欄；情境結構（A／B／C）不變，只動一個輸入。容量情境本身的差異見 Val 頁情境目標價。IF_HoldEcon 四世代同時一案（鍵 he_all）。只作對照的列（K 節公司指引、宣稱價格、共識、舊版目標價、Wattonomics）不列。"); r += 2
# 治理（原則 1）：build_id、輸入指紋、模型狀態
SE.cell(r, 1, "模型治理（原則第 10 節第 1 條）：Sens 主表與 RevDCF 是建置快照；只在 Excel 改 In 頁而未重跑 scripts/sens_v50.py 與建置時，下方模型狀態顯示 STALE").font = BOLD; r += 1
rng = f"In!$D$4:$F${IN_LAST}"
SE.cell(r, 1, "輸入指紋（本檔即時重算；In 頁低／基準／高欄，情境驅動對照以上）"); SE.cell(r, 4, f"=SUMPRODUCT(N(+{rng}),ROW({rng})*3+COLUMN({rng}))").number_format = "0.000000"; FPR = r; r += 1
SE.cell(r, 1, "輸入指紋（建置快照的基準檔）"); SE.cell(r, 4, SENS["base"]["fp"] if SENS else None); FBR = r; r += 1
SE.cell(r, 1, "基準加權目標價（建置快照）"); SE.cell(r, 4, SENS["base"]["tp"] if SENS else None); TBR = r; r += 1
SE.cell(r, 1, "模型狀態（CURRENT＝Sens 快照、HTML、README 與本檔輸入一致；STALE＝In 頁已改、須重跑 scripts/sens_v50.py 與建置）").font = BOLD
SE.cell(r, 4, f'=IF(AND(ISNUMBER(D{FBR}),ISNUMBER(D{TBR}),ABS(D{FPR}-D{FBR})<=0.000000001*MAX(1,ABS(D{FBR})),ABS(Val!C{CON["tp"]}-D{TBR})<0.005),"CURRENT","STALE")'); STR = r; r += 1
SE.cell(r, 1, "build_id（README、Sens、HTML 三處相同）"); SE.cell(r, 4, BUILD_ID); BIDR = r
SE.column_dimensions["A"].width = 62; SE.column_dimensions["B"].width = 36
for c in "CDEFGHIJKMN": SE.column_dimensions[c].width = 12
SE.column_dimensions["L"].width = 22

# =====================================================================================
# Sources
# =====================================================================================
SO = wb.create_sheet("Sources")
SO["A1"] = "Sources — 來源清單（事實總帳 coreweave/data/coreweave_facts_20261112.json；URL 與擷取日）"
facts = json.load(open(os.path.join(HERE, "data", "coreweave_facts_20261112.json"), encoding="utf-8"))
hdr(SO, 3, ["代號", "來源"])
for i, (k, v) in enumerate(facts["sources"].items()):
    SO.cell(4 + i, 1, k); SO.cell(4 + i, 2, v)
SO.cell(5 + len(facts["sources"]), 1, "Tokenomics"); SO.cell(5 + len(facts["sources"]), 2, TKSRC + "；名稱清單 coreweave/data/tokenomics_names.txt")
SO.column_dimensions["A"].width = 14; SO.column_dimensions["B"].width = 160

# =====================================================================================
# v5.0.2：HTML 版面規格 v1.3 需要的新增列（只附加在各頁最後，不移動既有儲存格；以「新增列以外逐格比對 0 差異」驗證）
#   Seg：分部營收推導（長約／隨需）；PL：標準損益表（GAAP，歷史 3 年＋預測 5 年＋9M 比較欄）；Sens：O 欄白話標籤；Sources：SEC XBRL 來源
# =====================================================================================
HIST = json.load(open(os.path.join(HERE, "data", "coreweave_is_hist_20261114.json"), encoding="utf-8"))
HY = HIST["years"]
TB_ = TR["B"]; SB_ = SR["B"]; PB_ = PR_["B"]
# ---- Seg：分部營收推導（情境 B）
r = SG.max_row + 2
SG.cell(r, 1, "分部營收推導（情境 B；v5.0.2 新增，HTML 驅動因子章）：AI 雲端平台營收＝計費 MW × 每 MW 年收入，依合約型態拆成長約與隨需（公司未揭露拆分；拆分比例為 In 頁隨需占比假設）").font = BOLD; SEGD = {"title": r}; r += 1
hdr(SG, r, YH); SEGD["hdr"] = r; r += 1
SGC = {2026: "D", **{y: yc(y) for y in FYM}}              # Seg 欄
THC = {2026: "D", **{y: TCOL[y] for y in FYM}}            # Thesis 欄（FY2026 用 Q4 模型欄 D）
LEN = {2026: "*0.25", **{y: "" for y in FYM}}
def sgrow(k, lab, unit, vals, fmt="#,##0.000", note=None):
    global r
    SEGD[k] = r; SG.cell(r, 1, lab); SG.cell(r, 2, unit)
    for c, v in vals.items():
        SG[f"{c}{r}"] = v; SG[f"{c}{r}"].number_format = fmt
    if note: SG.cell(r, 11, note)
    r += 1
sgrow("bill", "計費 MW（平均在役 MW × 計費比例；FY2026 欄＝第四季）", "MW", {SGC[y]: f"=Thesis!{THC[y]}{TB_['bill']}" for y in SGC}, "#,##0")
sgrow("anc", "每 MW 年持有成本（Tokenomics，在役世代加權）", "US$M／MW·年", {SGC[y]: f"=Thesis!{THC[y]}{TB_['anc']}" for y in SGC}, "0.00")
sgrow("pl", "長約每 MW 年收入（持有成本 × 長約定價倍數 × k 調整）", "US$M／MW·年", {SGC[y]: f"={SGC[y]}{SEGD['anc']}*k_long*k_adj" for y in SGC}, "0.00", "長約定價倍數＝IREN–Microsoft 長約價 ÷ GB300 持有成本（Analogy）")
sgrow("ps", "隨需每 MW 年收入（持有成本 × 隨需定價倍數 × k 調整）", "US$M／MW·年", {SGC[y]: f"={SGC[y]}{SEGD['anc']}*k_spot*k_adj" for y in SGC}, "0.00", "隨需定價倍數＝H100 現貨指數 ÷ H100 每 GPU 小時持有成本")
sgrow("od", "隨需占計費產能比例", "%", {SGC[y]: "=od_share" for y in SGC}, "0.0%", "公司未揭露長約／隨需比例（Assumed）")
sgrow("act", "實績（公司未揭露長約／隨需拆分）", "US$B", {"C": f"=C{SB_['seg']}", "D": "=rev9m"}, "#,##0.000", "FY2025＝10-K；FY2026＝9M 實績（Q3 新聞稿）")
sgrow("long", "長約營收（計費 MW ×（1 − 隨需占比）× 長約每 MW 年收入）", "US$B", {SGC[y]: f"={SGC[y]}{SEGD['bill']}*(1-{SGC[y]}{SEGD['od']})*{SGC[y]}{SEGD['pl']}{LEN[y]}/1000" for y in SGC})
sgrow("spot", "隨需營收（計費 MW × 隨需占比 × 隨需每 MW 年收入）", "US$B", {SGC[y]: f"={SGC[y]}{SEGD['bill']}*{SGC[y]}{SEGD['od']}*{SGC[y]}{SEGD['ps']}{LEN[y]}/1000" for y in SGC})
sgrow("ovr", "新增產能改用公司宣稱價格的差額（基準為 0；只在敏感度案例不為 0）", "US$B",
      {SGC[y]: f"=Thesis!{THC[y]}{TB_['rev']}-Thesis!{THC[y]}{TB_['bill']}*Thesis!{THC[y]}{TB_['pm']}{LEN[y]}/1000" for y in SGC})
sgrow("tot", "合計（＝AI 雲端平台分部營收＝合併營收）", "US$B", {"C": f"=C{SEGD['act']}", **{SGC[y]: f"=SUM({SGC[y]}{SEGD['act']}:{SGC[y]}{SEGD['ovr']})" for y in SGC}})
sgrow("lsh", "長約占營收比例（不含未拆分實績）", "%", {SGC[y]: f"={SGC[y]}{SEGD['long']}/({SGC[y]}{SEGD['long']}+{SGC[y]}{SEGD['spot']})" for y in SGC}, "0.0%")
sgrow("chk", "檢查：分部營收推導合計 − 合併營收（應為 0）", None, {c: f"={c}{SEGD['tot']}-{c}{SB_['tot']}" for c in "CDEFGHIJ"}, "0.000")
# ---- PL：標準損益表（GAAP；情境 B）
r = PL.max_row + 2
PLS = {"title": r}
PL.cell(r, 1, "標準損益表（GAAP；情境 B；v5.0.2 新增，HTML 損益章讀取）：FY2023–FY2025＝10-K 實績；FY2026＝9M 實績（Q3 10-Q）＋第四季模型；FY2027–FY2030＝模型；K 欄＝9M 2026 實績（比較欄）").font = BOLD; r += 1
PL.cell(r, 1, "口徑：GAAP（公司財報損益與 EPS 以 GAAP 為主；公司另揭露的調整後指標是調整後 EBITDA 與調整後營業利益，見上方列與 Q 頁）。營業成本＝公司損益表 Cost of revenue，不含折舊（折舊列在營業費用的 Technology and infrastructure），故毛利率偏高；預測期營業成本＝營運費用＋機房租金，營業費用＝管銷（扣 SBC）＋股票薪酬＋折舊攤銷。歷史與 9M 實績來源：" + HIST["source"]); PLS["note"] = r; r += 1
hdr(PL, r, ["項目", "單位", "FY2023A", "FY2024A", "FY2025A", "FY2026E", "FY2027E", "FY2028E", "FY2029E", "FY2030E", "9M 2026A", "說明／來源"]); PLS["hdr"] = r; r += 1
HC = {"FY2023": "C", "FY2024": "D", "FY2025": "E", "9M2026": "K"}
MC = {2026: "F", 2027: "G", 2028: "H", 2029: "I", 2030: "J"}     # 標準表欄 → 年度
ALLC = "CDEFGHIJK"
def plrow(k, lab, unit, vals, fmt="#,##0.000", note=None, inputs=()):
    global r
    PLS[k] = r; PL.cell(r, 1, lab); PL.cell(r, 2, unit)
    for c, v in vals.items():
        if v is None: continue
        PL[f"{c}{r}"] = v; PL[f"{c}{r}"].number_format = fmt
        if c in inputs: PL[f"{c}{r}"].font = BLUE
    if note: PL.cell(r, 12, note)
    r += 1
hin = lambda key, ys: {HC[y]: HY[y][key] for y in ys}
SRC_K = "10-K（FY2023–FY2025）／10-Q（9M 2026）經 SEC XBRL，見 Sources SEC_XBRL"
plrow("rev", "營收", "US$B", {"C": HY["FY2023"]["rev"], "D": "=Q!F4", "E": "=Q!G4", **{MC[y]: f"={yc(y)}{PB_['rev']}" for y in MC}, "K": "=rev9m"}, inputs="C", note="FY2023：" + SRC_K + "；FY2024–FY2025：Q 頁；預測：上方情境 B 合併營收")
plrow("revg", "營收年增率（YoY）", "%", {c: f"={c}{PLS['rev']}/{p}{PLS['rev']}-1" for p, c in zip("CDEFGHI", "DEFGHIJ")}, "0.0%")
plrow("cogs", "營業成本（COGS，不含折舊）", "US$B", {**hin("cogs", ("FY2023", "FY2024", "FY2025", "9M2026")), "F": f"=K{r}+Thesis!D{TB_['opex']}+Thesis!D{TB_['rent']}",
      **{MC[y]: f"={yc(y)}{PB_['opex']}+{yc(y)}{PB_['rent']}" for y in range(2027, 2031)}}, inputs="CDEK", note="歷史：" + SRC_K + "；預測＝營運費用（Tokenomics）＋機房租金；FY2026＝9M 實績（公司口徑）＋第四季模型口徑，毛利率的變動部分來自口徑差異")
plrow("gp", "毛利", "US$B", {c: f"={c}{PLS['rev']}-{c}{PLS['cogs']}" for c in ALLC})
plrow("gpm", "毛利率", "%", {c: f"={c}{PLS['gp']}/{c}{PLS['rev']}" for c in ALLC}, "0.0%")
plrow("opx", "營業費用（技術與基礎設施含折舊、研發、銷售行銷、一般管理）", "US$B", {**hin("opex", ("FY2023", "FY2024", "FY2025", "9M2026")), "F": f"=K{r}+Thesis!D{TB_['sga']}+Thesis!D{TB_['sbc']}+Thesis!D{TB_['dep']}",
      **{MC[y]: f"={yc(y)}{PB_['sga']}+{yc(y)}{PB_['sbc']}+{yc(y)}{PB_['da']}" for y in range(2027, 2031)}}, inputs="CDEK", note="歷史＝Total operating expenses − Cost of revenue（" + SRC_K + "）；預測＝管銷（扣 SBC）＋股票薪酬＋折舊攤銷")
plrow("oi", "營業利益", "US$B", {"C": HY["FY2023"]["oi"], "D": "=Q!F5", "E": "=Q!G5", **{MC[y]: f"={yc(y)}{PB_['oi']}" for y in MC}, "K": "=oi9m"}, inputs="C", note="預測：上方情境 B 營業利益（GAAP）")
plrow("oim", "營業利益率", "%", {c: f"={c}{PLS['oi']}/{c}{PLS['rev']}" for c in ALLC}, "0.0%")
plrow("int", "利息費用淨額（負值＝費用）", "US$B", {**hin("int", ("FY2023", "FY2024", "FY2025")), **{MC[y]: f"=-{yc(y)}{PB_['int']}" for y in MC}, "K": "=-int9m"}, inputs="CDE", note="預測：上方情境 B 利息費用淨額（CF 頁）")
plrow("oth", "其他收入（淨額，含公允價值調整損益）", "US$B", {**hin("oth", ("FY2023", "FY2024", "FY2025", "9M2026")), "F": f"=K{r}", **{MC[y]: 0 for y in range(2027, 2031)}}, inputs="CDEGHIJK",
      note="歷史：Other income (expense), net＋公允價值調整損益（" + SRC_K + "）；預測期設 0（與上方稅前列一致）；FY2026＝9M 實績＋第四季 0")
plrow("pti", "稅前利益", "US$B", {c: f"={c}{PLS['oi']}+{c}{PLS['int']}+{c}{PLS['oth']}" for c in ALLC}, note="FY2026 含 9M 實際其他收入，故與上方「稅前利益」列（只含模型項）不同；稅後淨利兩者一致")
plrow("tax", "所得稅", "US$B", {**hin("tax", ("FY2023", "FY2024", "FY2025", "9M2026")), "F": f"=K{r}+{yc(2026)}{PB_['tax']}", **{MC[y]: f"={yc(y)}{PB_['tax']}" for y in range(2027, 2031)}}, inputs="CDEK",
      note="預測：上方情境 B 所得稅（長期現金稅率 × 抵減 NOL 後應稅所得）；FY2026＝9M 實績＋第四季模型")
plrow("ni", "稅後淨利", "US$B", {c: f"={c}{PLS['pti']}-{c}{PLS['tax']}" for c in ALLC})
plrow("nim", "淨利率", "%", {c: f"={c}{PLS['ni']}/{c}{PLS['rev']}" for c in ALLC}, "0.0%")
plrow("eps", "EPS（GAAP，完全稀釋）", "US$", {**hin("eps", ("FY2023", "FY2024", "9M2026")), "E": f"=C{PB_['eps']}", **{MC[y]: f"={yc(y)}{PB_['eps']}" for y in MC}}, "0.00", inputs="CDK",
      note="歷史：歸屬普通股東淨損 ÷ 稀釋加權股數（" + SRC_K + "）；預測：上方情境 B EPS（淨利 ÷ 稀釋股數）")
plrow("sh", "完全稀釋加權股數", "十億股", {**hin("sh", ("FY2023", "FY2024", "FY2025", "9M2026")), **{MC[y]: f"={yc(y)}{PB_['sh']}" for y in MC}}, "0.000", inputs="CDEK")
PL.cell(r, 1, "比率").font = BOLD; PLS["ratio"] = r; r += 1
plrow("nig", "淨利年增率（÷ 前一年淨利絕對值；負值＝虧損擴大）", "%", {c: f"=({c}{PLS['ni']}-{p}{PLS['ni']})/ABS({p}{PLS['ni']})" for p, c in zip("CDEFGHI", "DEFGHIJ")}, "0.0%")
plrow("chk1", "檢查：毛利 − 營業費用 − 營業利益（應為 0）", None, {c: f"={c}{PLS['gp']}-{c}{PLS['opx']}-{c}{PLS['oi']}" for c in ALLC}, "0.000")
plrow("chk2", "檢查：稅後淨利 − 10-K／10-Q 或上方模型淨利（應為 0）", None, {"C": f"=C{PLS['ni']}-({HY['FY2023']['ni']})", "D": f"=D{PLS['ni']}-Q!F8", "E": f"=E{PLS['ni']}-Q!G8",
      **{MC[y]: f"={MC[y]}{PLS['ni']}-{yc(y)}{PB_['ni']}" for y in MC}, "K": f"=K{PLS['ni']}-ni9m"}, "0.000")
plrow("chk3", "檢查：稅前利益 − 10-K／10-Q 報導值（應為 0）", None, {HC[y]: f"={HC[y]}{PLS['pti']}-({HY[y]['pti']})" for y in HC}, "0.000")
PL.column_dimensions["L"].width = 70
# ---- CF：累計自由現金流（HTML 命題章收支關卡圖的折線；FY2030 欄＝Fund 頁收支關卡）
r = CF.max_row + 2
CFC = {"title": r}
CF.cell(r, 1, "累計自由現金流（FY2026 起；v5.0.2 新增，HTML 命題章收支關卡圖；FY2030 欄＝收支關卡）").font = BOLD; r += 1
for s in SC:
    CFC[s] = r; CF.cell(r, 1, f"累計自由現金流：{SCN[s]}"); CF.cell(r, 2, "US$B")
    for y in range(2026, 2031):
        CF[f"{yc(y)}{r}"] = f"=SUM($D${CR_[s]['fcf']}:{yc(y)}{CR_[s]['fcf']})"; CF[f"{yc(y)}{r}"].number_format = "#,##0.0"
    r += 1
CFC["chk"] = r; CF.cell(r, 1, "檢查：FY2030 累計自由現金流 − Fund 頁收支關卡（三情境絕對差合計，應為 0）")
CF[f"H{r}"] = "=" + "+".join(f"ABS(H{CFC[s]}-Fund!{c}{FS['gate']})" for s, c in zip(SC, "CDE")); CF[f"H{r}"].number_format = "0.000"
# ---- Sens：O 欄白話標籤（HTML 龍捲風圖用；不改 A 欄）
SE.cell(3, 15, "白話標籤（HTML 用）").font = BOLD; SE.cell(3, 15).fill = HDR
PLAIN = (("IF_HoldEcon", "每 MW 年持有成本（Tokenomics）"), ("IF_CapexIT", "每 MW 設備資本支出"), ("IF_OpexGW", "每 MW 營運費用"))
for rr in range(SENS_R[0], SENS_R[1] + 1):
    a_ = SE.cell(rr, 1).value
    if isinstance(a_, str):
        for x_, y_ in PLAIN: a_ = a_.replace(x_, y_)
        SE.cell(rr, 15, a_)
SE.column_dimensions["O"].width = 50
# ---- Val：估值章四個目標價（版面規格 v1.3 G 節，Andy 2026-11-14 09:03；附加在 Val 頁最後，不移動既有列）
r = V.max_row + 2
V.cell(r, 1, "估值章四個目標價（v5.0.2 新增；HTML 估值章表與圖；E 欄＝圖上簡稱）").font = BOLD; CON["t4_title"] = r; r += 1
hdr(V, r, ["項目", "單位", "目標價", "現價", "簡稱", "說明", "相對現價"]); CON["t4_hdr"] = r; r += 1
def t4(k, lab, val, short, note, inp=False):
    global r
    CON[k] = r; V.cell(r, 1, lab); V.cell(r, 2, "US$")
    c_ = V.cell(r, 3, val); c_.number_format = "#,##0.00"
    if inp: c_.font = BLUE
    V.cell(r, 4, "=px").number_format = "#,##0.00"; V.cell(r, 5, short); V.cell(r, 6, note)
    V.cell(r, 7, f"=C{r}/D{r}-1").number_format = "+0.0%;-0.0%"
    r += 1
t4("t4_dcf", "DCF 目標價（單純 DCF，情境加權）", f"=C{CON['dcf']}", "DCF", f'=IF(C{CON["dcf"]}=0,"三情境 DCF 皆失效（企業價值 ≤ 0），每股取 0（有限責任）","DCF 有效性："&C{CON["dvalid"]})')
t4("t4_comp", "Comparable 目標價（情境加權）", f"=C{CON['comp']}", "Comparable", "新雲同業前瞻 EV/EBITDA 中位數 × FY2028 調整後 EBITDA，情境 25／50／25 加權")
t4("t4_tp", "加權目標價（DCF 與 Comparable 加權；橫幅主值）", f"=C{CON['tp']}", "加權", f'="DCF "&TEXT(w_dcf,"0%")&"＋Comparable "&TEXT(1-w_dcf,"0%")&"；投資建議依此判定（再套收支關卡）"')
t4("t4_cons", "市場共識目標價（TIKR 平均目標價，2026-11-11，Q3 財報後）", 132, "市場共識",
   "TIKR 文章「Stifel Just Cut Its CoreWeave Target From $120 to $60, Even After a Raised Outlook」（2026-11-11）：The Street's average target … to $132；TIKR 追蹤 42 位分析師（文中未明示平均值的家數）；賣方共識（Interested-party），只作對照，不進目標價；出處見 Sources 頁（TIKR 共識列）", True)
t4("t4_comp0", "Comparable 目標價（未調整，情境加權）", f"=C{CON['comp0']}", "Comparable（未調整）", f'=IF(ABS(C{CON["comp0"]}-C{CON["comp"]})<1E-9,"相對倍數係數＝1，與調整後相同","相對倍數調整前")')
# ---- Sources：SEC XBRL
r = SO.max_row + 1
SO.cell(r, 1, HIST["source_key"]); SO.cell(r, 2, HIST["source"] + "；URL：https://data.sec.gov/api/xbrl/companyfacts/CIK0001769628.json；資料檔 coreweave/data/coreweave_is_hist_20261114.json")
SO.cell(r + 1, 1, "TIKR_CONS"); SO.cell(r + 1, 2, "TIKR 部落格〈Stifel Just Cut Its CoreWeave Target From $120 to $60, Even After a Raised Outlook〉（2026-11-11 發布與更新；https://www.tikr.com/blog/stifel-just-cut-its-coreweave-target-from-120-to-60-even-after-a-raised-outlook）：分析師平均目標價 US$132（一年前 US$144）；TIKR 追蹤 42 位分析師；擷取 2026-11-14。對照：MarketScreener 共識頁快取為 2026-10-09（平均 141.47、42 家）、Investing.com 快取為 2026-10-23（135.82），皆早於 Q3 財報，未採用")

# =====================================================================================
# v5.0.3（2026-11-14）：Coord 頁——三層協調（讀取層級、C0 公司實績 vs 層級隱含值）。只新增頁；既有儲存格不動（Recon 兩列說明文字改標除外）。
#   工作單 docs/workorders/20261114_第二批13家_v1.3與三層協調.md 第 2 節；規則 docs/plan/三層基準協調規則.md v1.4.2 C0 第 2 點；tools/coord/README.md
#   算力買賣方（新雲樣板）：不做兩讀數目標價；Nebius 等新雲複製本節（README v5.0.3 說明）。
# =====================================================================================
WATF = os.path.join(HERE, "..", "wattonomics", "data", "wattonomics_snapshot_v0.3.2.json")
WAT = json.load(open(WATF, encoding="utf-8"))
WN = WAT["names"]
WATV = WAT["source"]["version"]
WATFILE = WAT["source"]["file"]
CD = wb.create_sheet("Coord")
CD.column_dimensions["A"].width = 52
for c_, w_ in zip("BCDEFGHIJK", (30, 14, 16, 18, 12, 22, 70, 18, 70, 22)):
    CD.column_dimensions[c_].width = w_
GAPF = "+0.0%;-0.0%"
CD["A1"] = ("Coord — 三層協調（三層基準協調規則 v1.4.2；工作單 20261114 第二批；算力買賣方樣板：新雲 CoreWeave）。個股不裁決、不改上游值；"
            "只列：（G1）本模型讀取哪些層級名稱與版本、（G2）公司已公布實績 vs 層級隱含值（C0）。算力買賣方不做兩讀數目標價。")
CD["A1"].font = BOLD
rc = 3
CD.cell(rc, 1, "G1. 讀取層級（同源，不判定）：本模型讀的 Tokenomics、Wattonomics 名稱與版本").font = BOLD; G1h = rc; rc += 1
CD.cell(rc, 1, ("本模型只用 Tokenomics 計價與成本（每 MW 年持有成本 IF_HoldEcon、設備資本支出 IF_CapexIT、營運費用 IF_OpexGW、機房資本支出 L1_FacCapexMW、每 GPU 小時持有成本 IF_GPUhrEcon、折舊年限等，共 "
                f"{sum(1 for l_ in open(os.path.join(HERE, 'data', 'tokenomics_names.txt'), encoding='utf-8') if l_.strip() and not l_.startswith('#'))} 個名稱，清單見 coreweave/data/tokenomics_names.txt）；快照 Tokenomics {TKV}（{TKS['source']['file']}，commit {TKS['source']['commit'][:7]}）。"
                f"Wattonomics 不進估值：In 頁 wat26、wat30 是手填對照值（來源快照 v0.3 列 mw_crwv），已核對 v0.3.2 同值（2026 年底 {WAT['rows'][105]['values'][3]:,} MW、2030 年底 {WAT['rows'][105]['values'][7]:,} MW），本輪不換快照；"
                f"C0 表（G2）另讀 Wattonomics {WATV}（{WATFILE}）的玩家 GW 路徑與玩家資本支出作層級側，只用於 Coord 頁。以下各列個股自層級讀入、或層級讀公司目標值，差距必為零或無檢驗力，標「讀取層級（同源，不判定）」，不計入對帳中項目數。"))
G1note = rc; rc += 1
hdr(CD, rc, ["名稱", "層級與版本", "單位", "本模型採用值（或公司目標）", "層級值", "差距 %", "狀態", "說明"]); rc += 1
G1a = rc
SAMEL = "讀取層級（同源，不判定）"
def g1(name, ver, unit, mv, lv, fmt, note, gap=True):
    global rc
    CD.cell(rc, 1, name); CD.cell(rc, 2, ver); CD.cell(rc, 3, unit)
    CD.cell(rc, 4, mv).number_format = fmt; CD.cell(rc, 5, lv).number_format = fmt
    if gap: CD.cell(rc, 6, f'=IF(E{rc}=0,"N/M",D{rc}/E{rc}-1)').number_format = GAPF
    CD.cell(rc, 7, SAMEL); CD.cell(rc, 8, note); rc += 1
for g_ in ("gb300", "vr200"):
    g1(f"每 MW 年持有成本 IF_HoldEcon｜{GNAME[g_]}", f"Tokenomics {TKV}", "US$M／MW·年", f"=he_{g_}", tkv("IF_HoldEcon", g_), "#,##0.000", "In 頁 he_ 列直接讀入（基準成本情境）；差距恆為 0")
g1("每 MW IT 資本支出 IF_CapexIT｜GB300 NVL72", f"Tokenomics {TKV}", "US$M／MW", "=cx_gb300", tkv("IF_CapexIT", "gb300"), "#,##0.000", "In 頁 cx_ 列直接讀入；差距恆為 0")
g1("每 MW 年營運費用 IF_OpexGW｜GB300 NVL72", f"Tokenomics {TKV}", "US$M／MW·年", "=ox_gb300", tkv("IF_OpexGW", "gb300"), "#,##0.000", "In 頁 ox_ 列直接讀入；差距恆為 0")
g1("Wattonomics 2026 年底 CoreWeave 上線 MW（mw_crwv）", f"Wattonomics {WATV}", "MW", fv("ye26.active_mw"), WAT["rows"][105]["values"][3], "#,##0", "層級值＝公司 2026 年底目標（原文：more than 1.85 GW；Interested-party），與公司目標同源；Recon 頁 wat26 列沿用，差距與判定改標同源參考")
g1("Wattonomics 2030 年底 CoreWeave 上線 MW（mw_crwv）", f"Wattonomics {WATV}", "MW", f"=Thesis!I{TR['A']['conn']}", WAT["rows"][105]["values"][7], "#,##0", "層級值＝公司 2030 年目標上緣（最多 8 GW）；對照本模型情境 A（沿用該目標）。情境 B、C 與層級的差距（B −30%、C −42.5%）是模型情境與公司目標的差，不是實績對帳；對帳單見 Drive 02-AI供應鏈/對帳/ 20261113_對帳單_CoreWeave_J4新約價格與Wattonomics上線GW.md（只列參考）")
G1b = rc - 1
rc += 1
# ---------------- G2：C0（實績）----------------
CD.cell(rc, 1, "G2. C0（實績）：公司已公布實績 vs 層級隱含值（規則 v1.4.2 C0 第 2 點；工作單 2026-11-14）").font = BOLD; G2h = rc; rc += 1
CD.cell(rc, 1, ("方法：個股側只用已公布期間的公司實績（2026 年前三季＝2026-01～2026-09，Q3 新聞稿 2026-11-10 與 10-Q 2026-11-12；指引、RPO 與簽約電力的未來展望不算實績）。"
                "層級側的 GW 量（Wattonomics 玩家 GW 路徑）部分或全部讀公司實績，所以這張表實際檢驗的是單位經濟：營收＝層級在役 GW × Tokenomics 每 MW 持有成本（IF_HoldEcon 在役世代加權；層級側稱「持有成本隱含」，不乘本模型的定價倍數 k 與計費比例），檢驗的是每 MW 收入對持有成本；"
                "資本支出＝每 GW 資本支出（公司口徑總 PP&E 增量 ÷ 2026 年前三季淨增在役 GW，對 Wattonomics WIF_Gap_Player_crwv_2026_CapexAIGW ÷ WIF_Gap_Player_crwv_2026_AddGW）；在役電力的層級值同源（見下），不判定。"
                "時差：層級值與本公司會計年度都是日曆年。營收主讀數＝2026 年已公布期間累計實績 vs 層級 CY2026 × 已公布月數 ÷ 12（線性近似，Derived）；第二讀數＝最近一季（Q3，2026-07～09）× 4 vs 層級 CY2026（Q3 中點在 8 月中，晚於年中，成長年的年化速率會被線性近似壓低或推高，方向依該季相對全年平均的位置而定，不得寫成夾住真值）。"
                "資本支出主讀數＝公司口徑每 GW（Verified，10-Q 推得）；現金口徑每 GW 是口徑下界，只列讀數、不進判定。判定：營收兩讀數差距都超過 ±門檻且同號＝對帳中，只有一個超過＝時差邊界（參考），都在內＝一致；資本支出以主讀數超過門檻＝對帳中，否則一致。"))
G2note = rc; rc += 1
CD.cell(rc, 1, "對帳門檻（±）"); CD.cell(rc, 4, 0.15).number_format = "0%"; CD.cell(rc, 4).font = BLUE; CD.cell(rc, 8, "三層基準協調規則 C0／工作單：差距超過 ±15% 標「對帳中」；門檻為輸入（藍字）"); TOLC = f"$D${rc}"; rc += 1
hdr(CD, rc, ["項目", "期間／口徑", "單位", "公司實績", "層級隱含（同期間換算）", "差距 %", "狀態／判定", "公司實績口徑與出處", "公司實績等級", "層級口徑與出處", "層級等級"]); rc += 1
G2a = rc
PRU = SRCF["PR_Q3"]
LAYSRC = f"Wattonomics {WATV}（{WATFILE}）"
def c0row(item, per, unit, val, fmt, src="", tag="", lsrc="", ltag="", blue=False, mv=None):
    global rc
    CD.cell(rc, 1, item); CD.cell(rc, 2, per); CD.cell(rc, 3, unit)
    if val is not None:
        CD.cell(rc, 4, val).number_format = fmt
        if blue: CD.cell(rc, 4).font = BLUE
    CD.cell(rc, 8, src); CD.cell(rc, 9, tag); CD.cell(rc, 10, lsrc); CD.cell(rc, 11, ltag); rc += 1; return rc - 1
def reading(item, per, unit, comp, lay, fmt, src, tag, lsrc, ltag, status=None):
    global rc
    CD.cell(rc, 1, item); CD.cell(rc, 2, per); CD.cell(rc, 3, unit)
    CD.cell(rc, 4, comp).number_format = fmt; CD.cell(rc, 5, lay).number_format = fmt
    CD.cell(rc, 6, f"=D{rc}/E{rc}-1").number_format = GAPF; CD.cell(rc, 7, status if status else f'=IF(ABS(F{rc})>{TOLC},"超過門檻","門檻內")')
    CD.cell(rc, 8, src); CD.cell(rc, 9, tag); CD.cell(rc, 10, lsrc); CD.cell(rc, 11, ltag); rc += 1; return rc - 1
def judge(label, a, b):
    global rc
    CD.cell(rc, 1, label).font = BOLD; CD.cell(rc, 2, "兩讀數都超過門檻且同號＝對帳中")
    CD.cell(rc, 7, f'=IF(AND(ABS(F{a})>{TOLC},ABS(F{b})>{TOLC},SIGN(F{a})=SIGN(F{b})),"對帳中",IF(OR(ABS(F{a})>{TOLC},ABS(F{b})>{TOLC}),"時差邊界（參考）","一致"))')
    rc += 1; return rc - 1
# 輸入列
SAMEC = "參考（同源校準）"
r_mo = c0row("已公布月數", "2026-01～2026-09", "月", 9, "0", "Q3 2026 新聞稿與 10-Q（2026 年前三季）", "Verified", blue=True)
r_rev9 = c0row("營收 2026 年前三季累計（公司實績）", "2026-01～2026-09", "US$B", fv("9m.rev"), "#,##0.000", f"{PRU}；損益表 Revenue（facts 9m.rev；亦見 10-Q 0001769628-26-000465，2026-11-12）", "Verified")
r_revq = c0row("營收 Q3（公司實績）", "2026-07～2026-09", "US$B", fv("q3.rev"), "#,##0.000", f"{PRU}；損益表 Revenue（facts q3.rev）", "Verified")
r_pp0 = c0row("總 PP&E（gross）2025-12-31", "2025-12-31", "US$B", 33.941, "#,##0.000", "CoreWeave Q3 2026 10-Q（0001769628-26-000465）附註 Property and equipment, gross；US$33,941M（獨立核對代理自 10-Q 讀出，建置環境無法直接開啟 SEC）", "Verified", blue=True)
r_pp1 = c0row("總 PP&E（gross）2026-09-30", "2026-09-30", "US$B", 64.554, "#,##0.000", "同上；US$64,554M", "Verified", blue=True)
r_ppd = c0row("資本支出 2026 年前三季（公司口徑＝總 PP&E 增量，含應付設備款與融資取得）", "2026-01～2026-09", "US$B", f"=D{r_pp1}-D{r_pp0}", "#,##0.000", "總 PP&E 增量（10-Q 推得；與 Thesis 9M 公司口徑 30.65＝上半年 10-Q 16.139＋Q3 法說 14.511 相差約 0.04）", "Verified（10-Q 推得）")
r_cx9 = c0row("資本支出 2026 年前三季（現金購置 PP&E；口徑下界）", "2026-01～2026-09", "US$B", fv("9m.cashcapex"), "#,##0.000", f"{PRU}；現金流量表 Purchases of property and equipment, including capitalized internal-use software（facts 9m.cashcapex）；現金口徑，不含應付設備款與 OEM 融資", "Verified")
r_mw = c0row("在役電力（active power）2026-09-30（公司揭露）", "2026-09-30（時點）", "MW", fv("q3.active_mw"), "#,##0", f"{PRU}；原文 over 1.7 GW（季增 200+ MW；facts q3.active_mw）；營運指標，非財報審計數字", "Interested-party")
r_wq = c0row("Q3 淨增在役電力（公司揭露）", "2026-07～2026-09", "MW", 200, "#,##0", f"{PRU}；原文 季增 200+ MW（下限）", "Interested-party", blue=True)
r_l25 = c0row("層級：玩家 GW 年底 2025（WIF_PlayerGW_YE_crwv_2025）", "2025 年底", "GW", WN["WIF_PlayerGW_YE_crwv_2025"], "0.000", "", "", f"{LAYSRC} 名稱 WIF_PlayerGW_YE_crwv_2025（主路徑；公司 2025 年底 850 MW 屬實）", "Verified（2025）", blue=True)
r_l26 = c0row("層級：玩家 GW 年底 2026（WIF_PlayerGW_YE_crwv_2026）", "2026 年底", "GW", WN["WIF_PlayerGW_YE_crwv_2026"], "0.000", "", "", f"{LAYSRC} 名稱 WIF_PlayerGW_YE_crwv_2026＝CapexGW 頁 G190；其公式 G189＝MAX(個股資本支出換算 GW G187〔CapexAIModel 32.978 → G173 → G186 → 上線延遲調整；個股模型回寫值，同源〕, w1_crwv_q3 ÷ 1000 − F190)，w1_crwv_q3＝1,700（公司 Q3 新聞稿，Wattonomics In 第 467 列），延遲 GW 為 0；MAX 兩個分支都同源（個股資本支出換算 GW、公司 Q3 實績），1.70 GW 等於公司 2026-09-30 實績形成的下限（同源）", "同源（公司實績下限）", blue=True)
r_lhe = c0row("層級：在役世代加權每 MW 年持有成本 FY2026（Tokenomics IF_HoldEcon 加權）", "CY2026", "US$M／MW·年", f"=Thesis!E{TR['B']['anc']}", "#,##0.000", "", "", f"Tokenomics {TKV}（{TKS['source']['commit'][:7]}）IF_HoldEcon 依本模型在役世代占比加權（Thesis 頁錨列；未乘 k）", "Derived（世代占比 Assumed）")
r_lrev = c0row("持有成本隱含營收 CY2026（玩家 GW 年平均 × 加權每 MW 年持有成本）", "CY2026", "US$B", f"=(D{r_l25}+D{r_l26})/2*D{r_lhe}", "#,##0.000", "", "", "＝（YE2025＋YE2026）÷ 2 × 每 MW 年持有成本；GW × US$M／MW 單位相乘即 US$B；GW 量部分讀公司實績（消去），層級側不含計費比例與定價倍數 k", "Derived（GW 量部分同源）")
r_lcx = c0row("層級：玩家資本支出 CY2026（WIF_Gap_Player_crwv_2026_CapexAIGW）", "CY2026", "US$B", WN["WIF_Gap_Player_crwv_2026_CapexAIGW"], "#,##0.000", "", "", f"{LAYSRC} 名稱 WIF_Gap_Player_crwv_2026_CapexAIGW（新增 GW × Tokenomics 每 GW 資本支出；非 CapexAIModel）", "Derived（GW 量部分同源）", blue=True)
r_lad = c0row("層級：玩家新增 GW 2026（WIF_Gap_Player_crwv_2026_AddGW）", "CY2026", "GW", WN["WIF_Gap_Player_crwv_2026_AddGW"], "0.000", "", "", f"{LAYSRC} 名稱 WIF_Gap_Player_crwv_2026_AddGW（＝YE2026 − YE2025）", "Derived（同源）", blue=True)
r_lpg = c0row("層級：每 GW 資本支出（CapexAIGW ÷ AddGW）", "CY2026", "US$B／GW", f"=D{r_lcx}/D{r_lad}", "#,##0.00", "", "", "Tokenomics 每 GW 資本支出經 Wattonomics 換算；GW 量在分子分母消去，剩下單位成本", "Derived")
r_apg = c0row("公司：2026 年前三季淨增在役 GW（2026-09-30 實績 − 層級 YE2025）", "2026-01～2026-09", "GW", f"=(D{r_mw}-D{r_l25}*1000)/1000", "0.000", "1,700 MW（新聞稿，Interested-party）減 2025 年底 850 MW（Verified）", "Interested-party／Verified")
CD.cell(rc, 1, "Wattonomics 快照核對（上列藍字常數取自此檔）"); CD.cell(rc, 2, f"{WATFILE}（快照 coreweave 讀 wattonomics/data/wattonomics_snapshot_{WATV}.json）"); CD.cell(rc, 8, f"xlsx sha256 {WAT['source']['sha256']}（快照 source 欄）；建置時逐名核對 WIF_PlayerGW_YE_crwv_2025／2026、WIF_Gap_Player_crwv_2026_CapexAIGW／AddGW 取自該 JSON，不手寫"); r_sha = rc; rc += 1
# 讀數
r_m1 = reading("營收｜主讀數：已公布期間累計 vs 持有成本隱含 CY2026 × 已公布月數 ÷ 12", "2026-01～2026-09 累計", "US$B", f"=D{r_rev9}", f"=D{r_lrev}*D{r_mo}/12", "#,##0.000", "前三季累計營收（10-Q／新聞稿）", "Verified", "持有成本隱含營收 × 9 ÷ 12（線性時差近似）", "Derived（部分同源）")
r_m2 = reading("營收｜第二讀數：最近一季（Q3）× 4 vs 持有成本隱含 CY2026", "2026-07～2026-09 年化", "US$B", f"=D{r_revq}*4", f"=D{r_lrev}", "#,##0.000", "Q3 營收 × 4（Q3 中點 8 月中）", "Verified／Derived", "持有成本隱含營收（全年）", "Derived（部分同源）")
r_j1 = judge("營收（部分同源：GW 量讀公司實績；實際檢驗每 MW 收入對持有成本）：C0（實績）判定", r_m1, r_m2)
r_c1 = reading("資本支出｜主讀數：公司口徑每 GW（總 PP&E 增量 ÷ 淨增在役 GW）vs 層級每 GW 資本支出", "2026-01～2026-09", "US$B／GW", f"=D{r_ppd}/D{r_apg}", f"=D{r_lpg}", "#,##0.00", "總 PP&E 增量（10-Q 推得，Verified）÷ 淨增在役 GW（Interested-party／Verified）", "Verified（10-Q 推得）", "層級每 GW 資本支出（Tokenomics 單位成本）", "Derived")
r_c2 = reading("資本支出｜口徑下界（不進判定）：現金口徑每 GW vs 層級每 GW 資本支出", "2026-01～2026-09（現金購置）", "US$B／GW", f"=D{r_cx9}/D{r_apg}", f"=D{r_lpg}", "#,##0.00", "前三季現金購置 PP&E（不含應付設備款與 OEM 融資）÷ 淨增在役 GW", "Verified", "同上", "Derived", status="參考（口徑下界，不判定）")
CD.cell(rc, 1, "資本支出（部分同源：GW 量消去；實際檢驗每 GW 資本支出）：C0（實績）判定").font = BOLD; CD.cell(rc, 2, "主讀數差距超過門檻＝對帳中；現金口徑列不進判定")
CD.cell(rc, 7, f'=IF(ABS(F{r_c1})>{TOLC},"對帳中","一致")'); r_j2 = rc; rc += 1
r_w1 = reading("在役電力｜讀數一：2026-09-30 實績 vs 層級 YE2025→YE2026 線性內插", "2026-09-30（時點）", "MW", f"=D{r_mw}", f"=(D{r_l25}+(D{r_l26}-D{r_l25})*D{r_mo}/12)*1000", "#,##0", "公司揭露在役電力（over 1.7 GW）", "Interested-party", "層級年底 GW 路徑內插到 9 月底（× 1000 換 MW）；YE2026 為公司實績下限", "同源", status=SAMEC)
r_w2 = reading("在役電力｜讀數二：2026-09-30 實績 vs 層級 YE2026", "2026-09-30 vs 2026 年底", "MW", f"=D{r_mw}", f"=D{r_l26}*1000", "#,##0", "同上", "Interested-party", "層級 2026 年底 GW × 1000（＝公司 2026-09-30 實績下限，差距必為 0）", "同源", status=SAMEC)
CD.cell(rc, 1, "在役電力：C0（實績）判定").font = BOLD; CD.cell(rc, 2, "層級 YE2026 ＝ 公司 Q3 實績下限（CapexGW!G189 的 MAX）")
CD.cell(rc, 7, "參考（同源校準）：無可比實績（層級值同源）"); r_j3 = rc; rc += 1
r_wf = rc
CD.cell(rc, 1, "參考（流量版）：Q3 淨增在役 MW × 4 vs 層級 CY2026 淨增（YE2026 − YE2025）"); CD.cell(rc, 2, "2026-07～2026-09 年化 vs CY2026"); CD.cell(rc, 3, "MW")
CD.cell(rc, 4, f"=D{r_wq}*4").number_format = "#,##0"; CD.cell(rc, 5, f"=(D{r_l26}-D{r_l25})*1000").number_format = "#,##0"; CD.cell(rc, 6, f"=D{rc}/E{rc}-1").number_format = GAPF
CD.cell(rc, 7, "參考（同源，不判定）"); CD.cell(rc, 8, "Q3 公司揭露季增 200+ MW（下限）× 4"); CD.cell(rc, 9, "Interested-party"); CD.cell(rc, 10, "層級淨增＝YE2026 下限 − YE2025；YE2026 下限來自公司 9-30 實績，故此列只看速率、不是獨立檢驗"); CD.cell(rc, 11, "同源"); rc += 1
r_ctr = rc
CD.cell(rc, 1, "簽約電力（contracted power）2026-09-30 vs 層級"); CD.cell(rc, 2, "2026-09-30（時點）"); CD.cell(rc, 3, "MW"); CD.cell(rc, 4, fv("q3.ctr_mw")).number_format = "#,##0"
CD.cell(rc, 7, "無可比實績"); CD.cell(rc, 8, f"{PRU}；原文 over 4.6 GW（季增 900+ MW；facts q3.ctr_mw）；營運指標"); CD.cell(rc, 9, "Interested-party")
CD.cell(rc, 10, "Wattonomics 玩家 GW 路徑是在役量，沒有簽約電力序列；缺：層級側的簽約（或承諾）GW。2030 年底在役路徑（WIF_PlayerGW_YE_crwv_2030＝" + f"{WN['WIF_PlayerGW_YE_crwv_2030']:.2f} GW）只作參考，不同時點、不同概念，不判定"); rc += 1
r_cxd = rc
CD.cell(rc, 1, "參考：公司口徑資本支出前三季（上半年 10-Q 16.139＋Q3 法說 14.511）每 GW vs 層級每 GW"); CD.cell(rc, 2, "2026-01～2026-09 累計"); CD.cell(rc, 3, "US$B／GW")
CD.cell(rc, 4, f"=({fv('1h.capex_def')}+{fv('q3.capex_def')})/D{r_apg}").number_format = "#,##0.00"; CD.cell(rc, 5, f"=D{r_lpg}").number_format = "#,##0.00"; CD.cell(rc, 6, f"=D{rc}/E{rc}-1").number_format = GAPF
CD.cell(rc, 7, "參考（不判定）"); CD.cell(rc, 8, f"上半年 {fv('1h.capex_def')}（Q2 10-Q，Verified）＋Q3 {fv('q3.capex_def')}（法說簡報第 16 頁，原文未核；Interested-party）；等級混合；與主讀數的 10-Q 總 PP&E 增量相差約 0.04"); CD.cell(rc, 9, "Verified／Interested-party"); CD.cell(rc, 10, "同上層級每 GW"); CD.cell(rc, 11, "Derived"); rc += 1
r_xd = rc
CD.cell(rc, 1, "參考：每 MW 年收入——實績隱含 vs Tokenomics 加權持有成本"); CD.cell(rc, 2, "2026-01～2026-09 年化 ÷ 平均在役 MW（YE2025 與 2026-09-30 的平均）"); CD.cell(rc, 3, "US$M／MW·年")
CD.cell(rc, 4, f"=D{r_rev9}/D{r_mo}*12/((D{r_l25}*1000+D{r_mw})/2)*1000").number_format = "#,##0.00"; CD.cell(rc, 5, f"=D{r_lhe}").number_format = "#,##0.00"; CD.cell(rc, 6, f"=D{rc}/E{rc}-1").number_format = GAPF
CD.cell(rc, 7, "參考（不判定）"); CD.cell(rc, 8, "前三季營收年化 ÷（2025 年底 850 MW 與 2026-09-30 1,700 MW 的平均）；與營收主讀數同差（GW 量消去），這才是實際被檢驗的量；分母是平均在役 MW，含未計費與爬坡中產能；公司宣稱新約每 MW 年收入 20–25（Interested-party，分母口徑未說明）屬 J4 對帳單，只列參考"); CD.cell(rc, 9, "Derived"); CD.cell(rc, 10, "Tokenomics IF_HoldEcon 在役世代加權（持有成本，不含 k）"); CD.cell(rc, 11, "Derived"); rc += 1
G2b = rc - 1
rc += 1
# 核對與計數
CD.cell(rc, 1, "核對：營收 9 個月實績 − Thesis 9M 實績（應為 0）"); CD.cell(rc, 4, f"=D{r_rev9}-Thesis!C{TR['B']['rev']}").number_format = "0.000000"; r_k1 = rc; rc += 1
CD.cell(rc, 1, "對照：10-Q 總 PP&E 增量 − Thesis 9M 公司口徑資本支出（Thesis＝上半年 10-Q＋Q3 法說；差距小，不要求為 0）"); CD.cell(rc, 4, f"=D{r_ppd}-Thesis!C{TR['B']['cx']}").number_format = "0.000"; r_k2 = rc; rc += 1
CD.cell(rc, 1, "對帳中項目數（營收、資本支出；在役電力同源不判定，G1 與參考列不計）"); CD.cell(rc, 4, f'=COUNTIF(G{r_j1},"對帳中")+COUNTIF(G{r_j2},"對帳中")'); r_n = rc; rc += 1
CD.cell(rc, 1, "有判定的項目數（營收、資本支出）"); CD.cell(rc, 4, f'=COUNTA(G{r_j1},G{r_j2})'); r_nj = rc; rc += 1
CD.cell(rc, 1, "C0 對帳摘要（HTML 對帳章讀取）")
CD.cell(rc, 4, f'="實績對帳（公司已公布實績對 Wattonomics、Tokenomics 層級隱含值；層級的 GW 量讀公司實績，故實際只檢驗單位經濟）：營收前三季 US$"&TEXT(D{r_m1},"0.0")&"B（持有成本隱含同期間 US$"&TEXT(E{r_m1},"0.0")&"B，主讀數 "&TEXT(F{r_m1},"+0%;-0%")&"；Q3 年化 US$"&TEXT(D{r_m2},"0.0")&"B，第二讀數 "&TEXT(F{r_m2},"+0%;-0%")&"）判定："&G{r_j1}&"；資本支出公司口徑每 GW US$"&TEXT(D{r_c1},"0.0")&"B 對層級 US$"&TEXT(E{r_c1},"0.0")&"B（"&TEXT(F{r_c1},"+0%;-0%")&"；現金口徑下界 "&TEXT(F{r_c2},"+0%;-0%")&"，不判定）判定："&G{r_j2}&"；在役電力 "&TEXT(D{r_w1},"#,##0")&" MW："&G{r_j3}&"（層級 YE2026 是公司 Q3 實績下限）。對帳中項目 "&D{r_n}&" 個（門檻 ±"&TEXT({TOLC},"0%")&"）。簽約電力無可比實績。個股不裁決，不改上游值；證據回饋走對帳單。"'); r_sum = rc; rc += 1
CD.cell(rc, 1, "同源檢查（哪些部分同源、什麼會消去、這張表實際檢驗什麼）")
CD.cell(rc, 2, ("（一）本模型沒有把校準值寫入 Wattonomics 或 Tokenomics：定價倍數 k、計費比例、世代占比只留在本模型；層級側不含 k 與計費比例。"
                "（二）Wattonomics 玩家 GW 路徑 YE2026（CapexGW 頁 G190）的公式 G189＝MAX(個股資本支出換算 GW G187〔CapexAIModel 32.978 → G173 → G186 → 上線延遲調整；個股模型回寫值，同源〕, w1_crwv_q3 ÷ 1000 − F190)，w1_crwv_q3＝1,700 取自公司 Q3 新聞稿（Wattonomics In 第 467 列），延遲 GW 為 0，MAX 的兩個分支都同源（一支是個股模型回寫的資本支出，一支是公司 Q3 實績），所以層級 2026 年底 1.70 GW 不是獨立證據，等於公司 2026-09-30 實績形成的下限；在役電力列因此標「參考（同源校準）」，結論「無可比實績（層級值同源）」，不判定、不計數。"
                "（三）營收與資本支出的 GW 量也讀公司實績，在比值中消去：營收列實際檢驗每 MW 收入對 Tokenomics 持有成本（IF_HoldEcon 加權；層級側稱「持有成本隱含」，它是成本不是價格）；資本支出列改以每 GW 資本支出比較，實際檢驗 Tokenomics 單位資本支出對公司口徑實績。兩列標「部分同源（GW 量讀公司實績）」。"
                "（四）資本支出讀 CapexAIGW（GW 路徑 × Tokenomics 每 GW 資本支出），不讀 CapexAIModel（個股模型回寫值，同源）；不用 _SiteOnly 路徑（公司站點目標）。世代占比取自本模型 In 頁（Assumed），只影響每 MW 持有成本的加權。")); r_org = rc; rc += 1
CD.cell(rc, 1, "口徑換算與限制")
CD.cell(rc, 2, ("營收：層級側是持有成本（經濟口徑成本，不是價格）；公司實績含 GPU 雲端以外的軟體與服務，且計費 MW 小於在役 MW，爬坡期收入落後在役量。資本支出：公司口徑（總 PP&E 增量）含應付設備款與 OEM 融資，並含為後續上線預付的設備，所以每淨增 GW 的資本支出偏高；現金口徑不含應付設備款，是下界，兩者夾住的是口徑而非時差。層級單位資本支出由 Tokenomics 設備與機房成本組成，與公司口徑的差距只在單位成本。"
                "在役電力：公司以 active power 揭露，Wattonomics 以 IT 口徑記錄（mwBasis 為 Assumed）。另註（交主控評估是否開對帳單）：CoreWeave 2026-09-30 的在役 1,700 MW 已等於層級 YE2026 值，公司指引年底 over 1.85 GW，層級 YE2026 可能偏低（只有下限、沒有 Q4 新增）；這是同源下限造成的，不是實績對帳結論。")); r_k = rc; rc += 1
CD.freeze_panes = "B3"
COORD_ROWS = {"G1h": G1h, "G1a": G1a, "G1b": G1b, "G2h": G2h, "G2a": G2a, "G2b": G2b, "tol": int(TOLC.split("$")[-1]), "wf": r_wf, "sha": r_sha, "m1": r_m1, "m2": r_m2, "j1": r_j1, "c1": r_c1, "c2": r_c2, "j2": r_j2, "w1": r_w1, "w2": r_w2, "j3": r_j3,
              "ctr": r_ctr, "cxd": r_cxd, "xd": r_xd, "n": r_n, "nj": r_nj, "sum": r_sum, "org": r_org, "k": r_k, "g1note": G1note, "g2note": G2note, "mo": r_mo}
# Recon 兩列說明改標（只改說明文字；數值不動）
RC.cell(RCR["wat26"], 9, "讀取層級（同源，不判定）：Wattonomics 值＝公司 2026 年底目標；wattonomics_snapshot_v0.3 mw_crwv（v0.3.2 同值；Coord 頁 G1）")
RC.cell(RCR["wat30"], 9, "讀取層級（同源，不判定）：Wattonomics 值＝公司 2030 年目標上緣；模型情境 B、C 與其差距是情境差，不是實績對帳（Coord 頁 G1；對帳單草稿 coreweave/docs/reports/20261112_對帳單草稿_CoreWeave_J4與Wattonomics.md，只列參考）")

# =====================================================================================
# README（最後建，最前排；文字中的數字一律公式組字）
# =====================================================================================
RD = wb.create_sheet("README", 0)
VC = lambda k: f"Val!$C${CON[k]}"
T_ = lambda ref, f: f'TEXT({ref},"{f}")'
RD["A1"] = ('="CoreWeave（CRWV）個股模型 ' + VER + '（Oracle 收支型樣板重建＋2026Q3 滾入＋v4.13 驗收採用項）　資料截止日 "&TEXT(cutoff,"yyyy-mm-dd")&"；成品日 "&TEXT(built,"yyyy-mm-dd")&"；股價基準日 "&TEXT(px_date,"yyyy-mm-dd")&"；build_id '
            + BUILD_ID + '"')
RL = {}
_alt = [i for i, c in enumerate(SENS["cases"]) if c["key"] == "peer_nd_mode"] if SENS else []
ALT_SENT = (f'&"；同業淨負債改用 MarketScreener 年底預估（替代口徑，時點不一致）時加權目標價 "&TEXT(Sens!E{4 + _alt[0]},"0.00")&"（"&Sens!N{4 + _alt[0]}&"）"') if _alt else ""
lines = [
    ("thesis", '="命題：CoreWeave 已簽約的 backlog（"&TEXT(backlog,"0.0")&" 十億美元；RPO "&TEXT(rpo,"0.0")&"）能否以高於資金成本的報酬轉成營收與現金，同時不讓資金缺口失控。命題模組（在役 MW × 每 MW 年收入〔Tokenomics 持有成本 × 定價倍數 k〕、營運費用、租金、資本支出與折舊）驅動唯一的報告分部 AI 雲端平台"'),
    ("q", "核心問題：每增加 1 GW 的產能，是否創造高於資金成本的價值；資金缺口（自由現金流為負）要靠多少新債與新股補，股東被稀釋多少。"),
    ("rec", f'="投資建議："&{VC("rec_lab")}&"（加權目標價 US$"&TEXT({VC("tp")},"0.00")&"，現價 US$"&TEXT(px,"0.00")&"〔"&TEXT(px_date,"yyyy-mm-dd")&"〕，上檔空間 "&TEXT({VC("up")},"+0.0%;-0.0%")&"）——條件式：在每 MW 年收入＝Tokenomics 持有成本 × k "&TEXT(k_mix,"0.000")&"（IREN 長約 Analogy 9.70 與 H100 現貨）、計費比例 "&TEXT(u_bill,"0%")&"、AI 設備折舊 "&TEXT(life_ai,"0")&" 年（Tokenomics）的基準假設下；失效條件見本頁末"'),
    ("o1", f'="重點觀察 1：DCF 情境加權 "&TEXT({VC("dcf")},"0.00")&"、Comparable "&TEXT({VC("comp")},"0.00")&"（同業中位數 "&TEXT(mult,"0.00")&" 倍 × FY2028 調整後 EBITDA）；"&{VC("bn_dcf")}&"。上檔空間 "&TEXT({VC("up")},"+0.0%;-0.0%")&"＝DCF 貢獻 "&TEXT({VC("dcfup")},"+0.0%;-0.0%")&"＋Comparable 貢獻 "&TEXT({VC("cmpup")},"+0.0%;-0.0%")&"＋相對倍數調整 "&TEXT({VC("adjup")},"+0.0%;-0.0%"){ALT_SENT}'),
    ("o2", f'="重點觀察 2："&{VC("bn_scn")}&"。情境 B 的 FY2029 稅後 ROIC "&TEXT(Thesis!D{SUM_ROWS[0] + 5},"0.0%")&"，對照 WACC "&TEXT(wacc,"0.0%")&"。"'),
    ("o3", f'="重點觀察 3：收支關卡（FY2026–FY2030 自由現金流合計）情境 A／B／C 為 "&TEXT(Fund!C{FS["gate"]},"0.0")&"／"&TEXT(Fund!D{FS["gate"]},"0.0")&"／"&TEXT(Fund!E{FS["gate"]},"0.0")&" 十億美元；情境 B 期間最大單年新增舉債 "&TEXT(Fund!D{FS["maxb"]},"0.0")&"、最低年末現金 "&TEXT(Fund!D{FS["minc"]},"0.0")&"、淨負債 ÷ EBITDA 最差 "&TEXT(Fund!D{FS["ndmax"]},"0.00")&" 倍、利息保障最差 "&TEXT(Fund!D{FS["icmin"]},"0.00")&" 倍（"&IF(Fund!D{FS["ndmax"]}>lev_max,"淨負債 ÷ EBITDA 觸發 WARN","淨負債 ÷ EBITDA 未觸發")&"、"&IF(Fund!D{FS["icmin"]}<icov_min,"利息保障觸發 WARN","利息保障未觸發")&"）；2030 年底累計稀釋 "&TEXT(Fund!D{FS["dil"]},"0.0%")'),
    ("o4", '="J4：公司稱新簽 4–5 年約每 MW 年收入 US$"&TEXT(p_claim_lo,"0")&"–"&TEXT(p_claim_hi,"0")&"m（Interested-party），是本模型 FY2027 情境 B 每 MW 年收入 "&TEXT(Thesis!F' + str(TR["B"]["pm"]) + ',"0.00")&" 的 "&TEXT((p_claim_lo+p_claim_hi)/2/Thesis!F' + str(TR["B"]["pm"]) + ',"0.0")&" 倍；新增產能改用公司宣稱中點的 Sens 見 Sens 頁 p_new_ovr 列；不改基準 k"'),
    ("o5", f'="橫幅："&{VC("bn_edge")}&IF({VC("bn_frag")}="","","。"&{VC("bn_frag")})'),
    ("o6", f'="指引對照（原則 50）：模型 FY2026 營收 "&TEXT(PL!D{PR_["B"]["rev"]},"0.00")&"，"&Recon!E{RCR["rev26pos"]}&"（指引 "&TEXT(g_rev_lo,"0.0")&"–"&TEXT(g_rev_hi,"0.0")&"）；調整後營業利益 "&TEXT(PL!D{PR_["B"]["aoi"]},"0.000")&"，"&Recon!E{RCR["aoi26pos"]}&"；差距主要來自每 MW 年收入：達到營收指引中點需 Q4 每 MW 年收入 "&TEXT(Recon!E{RCR["rev4req"]},"0.00")&"（模型 "&TEXT(Recon!B{RCR["rev4req"]},"0.00")&"，模型輸出，非獨立證據）"'),
]
for i, (k, f) in enumerate(lines):
    RD.cell(2 + i, 1, f); RL[k] = 2 + i
r = 2 + len(lines) + 1
RD.cell(r, 1, "會計年度＝日曆年；FY2026＝9M 實績（Q3 新聞稿 2026-11-10）＋Q4 模型；延伸到 FY2032。金額十億美元，MW 為 IT 口徑（公司 active power 視為 IT，Assumed）。"); RL["cal"] = r; r += 1
RD.cell(r, 1, "營收分部層：CoreWeave 只有一個應報導部門（AI 雲端平台）；命題模組即驅動整個分部，核對列見 Seg 與 PL 頁。"); RL["seg"] = r; r += 2
RD.cell(r, 1, "失效條件（可觀察證據；出現任一，結論須重估）").font = BOLD; RL["fail"] = r; r += 1
for t in ('="1. 每 MW 年收入：若 10-Q 或公司揭露的新增產能營收顯示每 MW 年收入持續高於 "&TEXT(Thesis!F' + str(TR["B"]["pm"]) + '*1.5,"0.0")&"（本模型 FY2027 × 1.5），定價倍數 k 低估；可證偽指標：每季營收 ÷ 平均在役 MW"',
          '="2. 資金：若淨負債 ÷ 調整後 EBITDA 超過 "&TEXT(lev_max,"0.0")&" 倍仍能以投資級利率借款，或客戶預付顯著高於覆蓋比 "&TEXT(prepay_cov,"0%")&"，融資成本與稀釋低估"',
          '="3. 經濟壽命：若 GPU 實際汰換早於 "&TEXT(life_ai,"0")&" 年（Tokenomics 會計口徑），折舊與汰換資本支出低估；可證偽指標：二手 H100 租金與 Silicon Data 指數"'):
    RD.cell(r, 1, t); r += 1
r += 1
RD.cell(r, 1, "風險登錄（未進主模型的尾部風險；原則 10、19）").font = BOLD; RL["risk"] = r; r += 1
for t in ("1. 客戶集中（Q2 10-Q：最大客戶 36%、次大 26%）：是否建模＝否（事件型，進 Sens 尾端案待補）；理由：合約條款未揭露；觀察指標：10-Q 客戶占比與 RPO 客戶結構",
          "2. Section 382 NOL 使用限制：是否建模＝否（Unknown）；理由：股權變動測試資料未揭露；觀察指標：10-K 稅務附註",
          "3. 可轉債轉股與 ATM（35 百萬股上限）稀釋：是否建模＝部分（ATM 未計入評價股數）；觀察指標：10-Q 股數",
          "4. 站點交付延誤（房東 Core Scientific、Galaxy、Applied Digital）：是否建模＝情境 C；觀察指標：房東季報交付 MW",
          "5. 利率與信用利差：是否建模＝Sens（kd、高息債利率）；觀察指標：CRWV 優先票據殖利率與 CDS",
          "6. 利息以期初餘額計（Oracle 樣板慣例）：新增借款當年不計息，FY2027 情境 B 約少算 1.1 十億美元（Recon「新增借款當年未計的半年利息」列）；是否建模＝否（避免循環參照）；觀察指標：季報利息費用 vs 指引"):
    RD.cell(r, 1, t); r += 1
r += 1
RD.cell(r, 1, "已知限制").font = BOLD; r += 1
RD.cell(r, 1, "Q3 10-Q（2026-11-12 申報）：損益數字已經 SEC XBRL 取得並用於標準損益表（9M 營業成本、營業費用、其他收入、所得稅）；10-Q 原文未逐行對讀，債務逐檔、到期表、租賃、在建工程、期末股數、RSU 與選擇權仍沿用 Q2 10-Q 或標待補；Q3 公司口徑資本支出 14.511、RPO 分桶、加權債務成本 7.9% 為法說數（新聞稿未見）。Comparable 同業 3 家（Nebius、Oracle、IREN；MarketScreener 共識，估值日 2026-11-11），少於 8 家不做迴歸；中位數落在 Oracle（混合業務），淨負債口徑影響大（Val 並列兩口徑）。β 以 US 伺服器／EMS 組代理（新雲無自算組）。kd 為 2032 票據發行票面，市場殖利率未取得。"); RL["lim"] = r; r += 2
RD.cell(r, 1, "驅動 → 推導對照（每頁輸入、推導量、因果方向）").font = BOLD; RL["drv"] = r; r += 1
for t in ("In：一列一鍵（值、單位、標記、來源、說明、來源誘因）→ 所有公式以鍵名引用",
          "Thesis：在役 MW 路徑 × 每 MW 年收入（Tokenomics × k）→ 營收；Tokenomics 營運費用、租金、管銷 → 調整後 EBITDA；新增 MW × IF_CapexIT → 資本支出 → 折舊 → 營業利益",
          "Seg、PL：唯一分部 → 合併營收 → 營業利益 → 利息（CF）、稅（NOL）→ 淨利、EPS",
          "CF、Fund、Debt、Lease、RPO、Sites：自由現金流 → 現金 → 投資級債 → 新股 → 高息債；收支關卡與期間流動性；RPO 排程與站點對照（不校準）",
          "Val、RevDCF、Recon、Sens：DCF（FCFF）＋Comparable → 加權目標價、評等；反解支持現價所需每 MW 收入；指引與公司說法只對照"):
    RD.cell(r, 1, t); r += 1
r += 1
RD.cell(r, 1, "展示層（HTML 版面 v1.2；以下各列供 HTML 引用）").font = BOLD; RL["show"] = r; r += 1
for t in ("1. 每 MW 的經濟性：CoreWeave 每增加 1 MW，收入（Tokenomics 持有成本 × k）扣掉營運費用、租金與折舊後，是否高於資金成本？",
          "2. 資金：建置所需的資本支出遠大於營運現金流，缺口靠多少新債與新股補、股東被稀釋多少？",
          "3. 公司說法：公司稱新約每 MW 年收入遠高於 Tokenomics 錨，現價是否已經隱含公司說法成立？"):
    RD.cell(r, 1, t); r += 1
r += 1
RD.cell(r, 1, "命題回覆（HTML 摺疊）").font = BOLD; RL["ans"] = r; r += 1
RD.cell(r, 1, f'="1. 每 MW 經濟性？答：情境 B 的 FY2029 稅後 ROIC "&TEXT(Thesis!D{SUM_ROWS[0] + 5},"0.0%")&"，對照 WACC "&TEXT(wacc,"0.0%")&"；"&IF(Thesis!D{SUM_ROWS[0] + 5}<wacc,"低於資金成本，擴張不創造價值（模型輸出，非獨立證據）","高於資金成本（模型輸出，非獨立證據）")'); r += 1
RD.cell(r, 1, f'="2. 資金？答：情境 B 在 Q4 2026–FY2030 需外部融資 "&TEXT(Fund!D{FS["cum"]},"0.0")&" 十億美元（投資級債 "&TEXT(Fund!D{FS["ig"]},"0.0")&"、新股 "&TEXT(Fund!D{FS["eq"]},"0.0")&"、高息債 "&TEXT(Fund!D{FS["hy"]},"0.0")&"），2030 年底累計稀釋 "&TEXT(Fund!D{FS["dil"]},"0.0%")'); r += 1
if REV:
    RD.cell(r, 1, f'="3. 公司說法？答：支持現價所需的 FY2030 每 MW 年收入（情境 B）為 "&TEXT(RevDCF!G{RVN["px"]},"0.00")&"，"&IF(RevDCF!G{RVN["px"]}<p_claim_lo,"低於",IF(RevDCF!G{RVN["px"]}>p_claim_hi,"高於","落在"))&"公司宣稱新約區間 "&TEXT(p_claim_lo,"0")&"–"&TEXT(p_claim_hi,"0")&"；本模型基準 "&TEXT(Thesis!I{TR["B"]["pm"]},"0.00")&"（Tokenomics 錨 × k）"'); r += 1
else:
    RD.cell(r, 1, "3. 公司說法？答：反解快照未產生"); r += 1
r += 1
RD.cell(r, 1, "v5.0.1（2026-11-13）：Andy 2026-11-13「个股模型还是要遵守这个收之不过就不能买进的这个规则……个股模型这边还是要用相同的原则来评断它的評等」→ 投資建議加上收支關卡（原則 63）：情境 B 的 FY2026–FY2030 自由現金流合計為負、而估值評等為買進以上時，改列「收支不足」。加權目標價與所有數值不變（76.20）；另更正已知限制中過時的同業家數與 In 頁隨需占比說明。"); RL["v501"] = r; r += 1
RD.cell(r, 1, "v5.0（2026-11-12）：Andy 2026-11-12「Coreweave 已经公布财报了，所以要更新 Coreweave 的财报。再将 html 的版面套用 Oracle 最新的版本。」＋確認 v4.13 驗收採用規劃書 J1–J5。舊 JS 雙引擎（v4.13）移 coreweave/legacy_v4/；成品 v4.13 移 dist/_待刪除/。"); RL["v50"] = r; r += 1
if DEC:
    RD.cell(r, 1, "目標價變動拆解（scripts/decomp_v50.py）：" + "；".join(f"{d['label']} {d['delta']:+.2f}" for d in DEC["steps"]) + f"；合計 {DEC['total']:+.2f}"); RL["dec"] = r; r += 1
r += 1
RD.cell(r, 1, "展示層 v1.3（HTML 版面規格 v1.3，2026-11-14；以下各列供 HTML 引用；白話、不含參數代號）").font = BOLD; RL["s13"] = r; r += 1
_V = VC
_GATE3 = f'TEXT(Fund!C{FS["gate"]},"0.0")&"、"&TEXT(Fund!E{FS["gate"]},"0.0")'
_dcf0 = f'IF({_V("dcf")}=0,"現金流折現法（DCF）三個情境的每股價值都是 0，加權目標價全部來自同業倍數。","")'
L13 = [
    ("hero13", f'=IF({_V("rec")}="收支不足","投資建議「收支不足」：估值面上檔空間 "&TEXT({_V("up")},"+0.0%;-0.0%")&"，估值評等為"&{_V("val_lab")}&"；但基準情境（B）FY2026–FY2030 自由現金流合計 "&TEXT({_V("gatev")},"0.0")&" 十億美元，入不敷出，依收支關卡不得列為買進。","投資建議「"&{_V("rec")}&"」：上檔空間 "&TEXT({_V("up")},"+0.0%;-0.0%")&"；基準情境 FY2026–FY2030 自由現金流合計 "&TEXT({_V("gatev")},"0.0")&" 十億美元。")&{_dcf0}'),
    ("ob1", f'="入不敷出是結論的關鍵：基準情境 FY2026–FY2030 自由現金流合計 "&TEXT({_V("gatev")},"0.0")&" 十億美元，樂觀與保守情境為 "&{_GATE3}&"，三個情境都要靠外部資金；估值面上檔空間 "&TEXT({_V("up")},"+0.0%;-0.0%")&"，距最近的評等級距只有 "&TEXT({_V("dist")}*100,"0.0")&" 個百分點。"'),
    ("ob2", f'="目標價全部來自同業倍數：DCF 情境加權每股 "&TEXT({_V("dcf")},"0.00")&" 美元，同業倍數法 "&TEXT({_V("comp")},"0.00")&" 美元（同業中位數 "&TEXT(mult,"0.00")&" 倍 × FY2028 調整後 EBITDA），各占 "&TEXT(w_dcf,"0%")&" 得 "&TEXT({_V("tp")},"0.00")&" 美元；同業倍數降到 "&TEXT({_V("mdn")},"0.00")&" 倍，估值評等就降一級。"'),
    ("ob3", f'=IF({_V("npvAB")}<0,"擴張越多、價值越低：樂觀相對基準情境每多 1 GW，DCF 企業價值減少 "&TEXT(-{_V("npvAB")},"0.0"),"擴張創造價值：樂觀相對基準情境每多 1 GW，DCF 企業價值增加 "&TEXT({_V("npvAB")},"0.0"))&" 十億美元；基準情境 FY2029 稅後投入資本報酬率 "&TEXT(Thesis!D{SUM_ROWS[0] + 5},"0.0%")&"，"&IF(Thesis!D{SUM_ROWS[0] + 5}<wacc,"遠低於","高於")&"資金成本 "&TEXT(wacc,"0.0%")&"。"'),
    ("ob4", f'="缺口靠舉債與增資補：基準情境 2026 年第四季至 FY2030 需外部融資 "&TEXT(Fund!D{FS["cum"]},"0.0")&" 十億美元（投資級債 "&TEXT(Fund!D{FS["ig"]},"0.0")&"、新股 "&TEXT(Fund!D{FS["eq"]},"0.0")&"、高息債 "&TEXT(Fund!D{FS["hy"]},"0.0")&"），2030 年底股東累計稀釋 "&TEXT(Fund!D{FS["dil"]},"0.0%")&"；淨負債對 EBITDA 最差 "&TEXT(Fund!D{FS["ndmax"]},"0.00")&" 倍"&IF(Fund!D{FS["ndmax"]}>lev_max,"，超過 "&TEXT(lev_max,"0.0")&" 倍警戒","")&"。"'),
    ("ob5", '="公司說的新約價格是模型的 "&TEXT((p_claim_lo+p_claim_hi)/2/Thesis!I' + str(TR["B"]["pm"]) + ',"0.0")&" 倍，但現價用不到那麼高：公司稱新約每 MW 年收入 "&TEXT(p_claim_lo,"0")&"–"&TEXT(p_claim_hi,"0")&" 百萬美元，模型 FY2030 為 "&TEXT(Thesis!I' + str(TR["B"]["pm"]) + ',"0.00")&"' + (f'；支撐現價只需 "&TEXT(RevDCF!G{RVN["px"]},"0.00")&"（同為 FY2030）' if REV else '') + f'。另 FY2026 營收 "&TEXT(PL!D{PR_["B"]["rev"]},"0.00")&" 十億美元，"&Recon!E{RCR["rev26pos"]}&"。"'),
    ("rev13", (f'="反向 DCF 問的是：現價 "&TEXT(px,"0.00")&" 美元隱含每 MW 年收入要多少。其他條件不變，FY2030 基準情境每 MW 年收入 "&TEXT(RevDCF!G{RVN["px"]},"0.00")&" 百萬美元即可支撐現價，"&IF(RevDCF!I{RVN["px"]}<0,"比模型基準 "&TEXT(RevDCF!H{RVN["px"]},"0.00")&" 低 "&TEXT(-RevDCF!I{RVN["px"]},"0.0%"),"比模型基準 "&TEXT(RevDCF!H{RVN["px"]},"0.00")&" 高 "&TEXT(RevDCF!I{RVN["px"]},"0.0%"))&"；估值評等升一級需 "&TEXT(RevDCF!G{RVN["up"]},"0.00")&"、降一級的邊界在 "&TEXT(RevDCF!G{RVN["dn"]},"0.00")&"。公司宣稱新約中點 "&TEXT(RevDCF!J{RVN["px"]},"0.00")&"，"&IF(RevDCF!J{RVN["px"]}>RevDCF!G{RVN["up"]},"遠高於支撐現價所需：現價並未反映公司說法。","並未明顯高於支撐現價所需。")' if REV else "反解快照未產生")),
    ("pl13", '="口徑：GAAP。CoreWeave 財報的損益表與每股盈餘以 GAAP 為主，公司另揭露的調整後指標是調整後 EBITDA 與調整後營業利益（不是調整後 EPS），因此本表用 GAAP、EPS 為完全稀釋。公司的營業成本不含折舊（折舊列在營業費用的技術與基礎設施項下），所以毛利率偏高，營業利益率才是可比的獲利指標。FY2023–FY2025 為 10-K 實績，FY2026 為前三季實績加第四季模型，FY2027 起為基準情境模型；預測期其他收入設為零。FY2026 的營業成本是前三季公司口徑加第四季模型口徑（營運費用＋機房租金），毛利率的變動部分來自口徑差異。歷史 EPS 以歸屬普通股東淨損計算（含優先股股利等調整），不等於淨利 ÷ 股數：例如 FY2024 淨利 "&TEXT(PL!D' + str(PLS["ni"]) + ',"0.000")&" ÷ 稀釋股數 "&TEXT(PL!D' + str(PLS["sh"]) + ',"0.000")&"＝"&TEXT(PL!D' + str(PLS["ni"]) + '/PL!D' + str(PLS["sh"]) + ',"0.00")&"，報導 EPS 為 "&TEXT(PL!D' + str(PLS["eps"]) + ',"0.00")&"。"'),
    ("plb13", '="FY2026 欄的淨利含前三季實際所得稅 "&TEXT(PL!K' + str(PLS["tax"]) + ',"0.000")&" 與其他收入 "&TEXT(PL!K' + str(PLS["oth"]) + ',"0.000")&"；本表的稅前利益與所得稅列只含模型項，所以 FY2026 稅前 − 所得稅 ≠ 淨利。完整對照見本章標準損益表。"'),
    ("drv13", "驅動主線：在役 MW（情境路徑）× 計費比例＝計費 MW；計費 MW × 每 MW 年收入（Tokenomics 每 MW 持有成本 × 定價倍數）＝營收；營運費用與機房租金依在役 MW 計，資本支出＝新增 MW × 每 MW 設備成本，折舊依會計年限攤提。下方先看分部營收推導，再看每 MW 年收入與 backlog 兩張圖，參數明細收在摺疊。"),
    ("seg13", '="分部營收推導：CoreWeave 只有一個報導分部（AI 雲端平台），本模型依合約型態拆成長約與隨需。長約每 MW 年收入＝持有成本 × "&TEXT(k_long,"0.000")&"（IREN–Microsoft 長約類比），隨需＝持有成本 × "&TEXT(k_spot,"0.000")&"（H100 現貨指數），隨需占計費產能 "&TEXT(od_share,"0%")&"（假設，公司未揭露）；兩者各乘上計費 MW 後加總＝分部營收＝合併營收。FY2026 只拆第四季，前三季為未拆分實績。"'),
    ("px13", '="現價（"&TEXT(px_date,"yyyy-mm-dd")&" 收盤）"'),
    ("fund13", '="收支模組：基準情境 FY2026–FY2030 自由現金流合計 "&TEXT(Fund!D' + str(FS["gate"]) + ',"0.0")&" 十億美元；缺口依序以現金、投資級債（借款不超過 EBITDA 的 "&TEXT(lev_max,"0.0")&" 倍）、新股（每年不超過市值 "&TEXT(eq_cap,"0%")&"）、高息債補足。"'),
]
for k, t in L13:
    RD.cell(r, 1, t); RL[k] = r; r += 1
r += 1
RD.cell(r, 1, "v5.0.3（2026-11-14）：三層協調（三層基準協調規則 v1.4.2；工作單 20261114 第二批）。版面 v1.3 補缺項：估值章第一個區塊＝四個目標價與現價圖（G）、敏感度章前兩個區塊＝低案、高案主圖（H），章節說明移到圖下。新增 Coord 頁：G1 讀取層級（Tokenomics v22.0 名稱、Wattonomics 對照值；同源，不判定）、G2 C0（實績）——營收（對 Tokenomics 持有成本隱含值）與每 GW 資本支出（公司口徑，對 Wattonomics／Tokenomics 單位資本支出）；在役電力因層級 YE2026 即公司 Q3 實績下限，標參考（同源校準）、不判定；簽約電力無可比實績；HTML 對帳章同步新增。只新增頁與改標 Recon wat26、wat30 兩列說明文字；目標價 76.20、投資建議收支不足不變。新雲樣板：Nebius 照本節複製（README 的 v5.0.3 節）。"); RL["v503"] = r; r += 1
RD.cell(r, 1, "v5.0.2（2026-11-14）：Andy 2026-11-14 08:11 個股 HTML 版面規格 v1.3 試作（章節改為命題→估值→敏感度→反向 DCF→損益→驅動→資金缺口與融資→對帳→資料狀態；橫幅說明 ≤200 字、重點觀察 ≤5 點 ≤750 字、每章至少一張圖、標準損益表）。xlsx 只在各頁最後附加列：PL 標準損益表（GAAP，FY2023–FY2030＋9M 2026，歷史取 SEC XBRL）、Seg 分部營收推導（長約／隨需）、Sens O 欄白話標籤、Sources SEC_XBRL；既有儲存格逐格不變（目標價 76.20、投資建議收支不足）。"); RL["v502"] = r; r += 1
RD.column_dimensions["A"].width = 180

# 列號對照
ROWMAP.update({"TR": TR, "SR": SR, "PR": PR_, "CR": CR_, "FS": FS, "FR": FR, "FUH": FUH, "RPB": RPB, "ST": ST, "D": D_, "LSB": LSB, "VR": VR, "DC": DC, "CO": CO, "CON": CON,
               "PEERS": [P0, P1], "PEERH": PEERH, "RVN": RVN, "RCR": RCR, "SENS": SENS_R, "GRID": list(GRID), "SNAP": SNAPR, "RL": RL, "SUM": list(SUM_ROWS), "CHK": [CHK_SEG, CHK_SUB],
               "QR": QR, "IR": IR, "SEGD": SEGD, "PLS": PLS, "CFC": CFC, "IN_LAST": IN_LAST, "GOV": {"fp": FPR, "fb": FBR, "tb": TBR, "st": STR, "bid": BIDR}, "BUILD_ID": BUILD_ID, "COORD": COORD_ROWS})
if os.environ.get("CRWV_ROWMAP"):
    json.dump(ROWMAP, open(os.environ["CRWV_ROWMAP"], "w", encoding="utf-8"), ensure_ascii=False, indent=0)
wb.save(sys.argv[1])
print("saved", sys.argv[1], BUILD_ID)
