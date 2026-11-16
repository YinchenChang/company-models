# WhiteFiber（WYFI）個股模型 v0.4.0（CoreWeave v5.0.3 新雲樣板重建〔經 Nebius v0.6.0 分部版〕＋2026Q3 滾入＋三層協調 Coord；工作單 docs/workorders/20261114_第二批13家_v1.3與三層協調.md）
# 用法：python3 whitefiber/build_whitefiber.py raw.xlsx ；再以 LibreOffice headless 重算（whitefiber/make_v1.sh 一次跑完）。
# 會計年度＝日曆年。金額：百萬美元（US$M）；股數：百萬股；MW＝IT 關鍵電力口徑（雲端在役 MW；託管另列）。
# 工作表：README → In → Q → Thesis → Seg → PL → CF → Fund → RPO → Sites → Debt → Lease → Val → RevDCF → Recon → Sens → Sources → Coord
# 環境變數：SENS_OVR＝JSON {In 鍵: 值}；WYFI_TKV＝Tokenomics 快照版本（預設 v22.0）；WYFI_ROWMAP＝列號對照輸出路徑。
import sys, os, json, hashlib, datetime
import openpyxl
from openpyxl.styles import Font, PatternFill
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.utils import get_column_letter as CL

HERE = os.path.dirname(os.path.abspath(__file__))
VER = "v0.4.3"
BUILT = "2026-11-15"
OVR = json.loads(os.environ.get("SENS_OVR", "{}"))
TKV = os.environ.get("WYFI_TKV", "v22.0")
TKS = json.load(open(os.path.join(HERE, "data", f"tokenomics_snapshot_{TKV}.json"), encoding="utf-8"))
TK = TKS["items"]
FACTS_F = "whitefiber/data/whitefiber_facts_20261115.json"
FX = {it["id"]: it for it in json.load(open(os.path.join(HERE, "data", "whitefiber_facts_20261115.json"), encoding="utf-8"))["items"]}
SENS_F = os.path.join(HERE, "data", "sens_v50.json")
SENS = json.load(open(SENS_F, encoding="utf-8")) if (os.path.exists(SENS_F) and not OVR) else None
REV_F = os.path.join(HERE, "data", "reverse_v50.json")
REV = json.load(open(REV_F, encoding="utf-8")) if (os.path.exists(REV_F) and not OVR) else None
DEC_F = os.path.join(HERE, "data", "decomp_v50.json")
DEC = json.load(open(DEC_F, encoding="utf-8")) if (os.path.exists(DEC_F) and not OVR) else None


def _bid():
    h = hashlib.sha256()
    for p in (__file__, os.path.join(HERE, "data", "whitefiber_facts_20261115.json"), os.path.join(HERE, "data", f"tokenomics_snapshot_{TKV}.json"), os.path.join(HERE, "data", "whitefiber_is_hist_20261115.json")):
        h.update(open(p, "rb").read())
    return f"{VER}-{h.hexdigest()[:10]}"


BUILD_ID = _bid()
BLUE = Font(color="0000FF"); BOLD = Font(bold=True); GREY = Font(color="808080")
HDR = PatternFill("solid", fgColor="DDEBF7"); KEY = PatternFill("solid", fgColor="FFF2CC")
wb = openpyxl.Workbook()
SC = ("A", "B", "C")
SCN = {"A": "情境 A（樂觀：2030 年底雲端在役 85 MW）", "B": "情境 B（基準：2030 年底雲端在役 59 MW）", "C": "情境 C（保守：只計已具名合約 25 MW）"}
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


_FACTS = json.load(open(os.path.join(HERE, "data", "whitefiber_facts_20261115.json"), encoding="utf-8"))
import re as _re
SRCF = {k_: _re.sub(r"https?://\S+", "（網址見 Sources 頁 " + k_ + "）", v_).strip() for k_, v_ in _FACTS["sources"].items()}   # 可見文字不含網址（HTML 自含檢查）；網址列在 Sources 頁
WAT0 = json.load(open(os.path.join(HERE, "..", "wattonomics", "data", "wattonomics_snapshot_v0.3.5.json"), encoding="utf-8"))
INC_CO = "公司自揭 KPI／非 GAAP：管理層有呈現成長與獲利的誘因"
INC_GUIDE = "管理層指引：有設定可達成目標、維持市場預期的誘因"
INC_SELL = "賣方共識：投行業務、承銷關係，偏多傾向"
INC_PEER = "同業公司自揭或賣方共識：有利害關係"
INC_TK = "—"

# =====================================================================================
# In
# =====================================================================================
IN = wb.active; IN.title = "In"
IN["A1"] = "In — 輸入（藍字＝輸入；黑字＝公式；基準欄 E 進主模型；D／F＝敏感度低／高，只用於 Sens 頁、不進目標價；情境 A／B／C 的驅動見本頁最後一節「情境驅動對照」）。金額：百萬美元（US$B）；MW＝IT 口徑；每 MW 金額：百萬美元／MW·年（＝US$B／GW·年）"
IN["A2"] = "標記：每列單一主要標記 Verified／Interested-party／Analogy／Assumed／Derived（原則第 10 節第 2 條；Interested-party 列見 J 欄「來源誘因」）；Analogy、Assumed 給低／高區間。SemiAnalysis 未使用。「沿用 v4.13」＝沿用 coreweave/legacy_v4/company.json 已核准參數（原出處見來源欄）"
hdr(IN, 3, ["鍵", "項目", "單位", "敏感度低", "基準", "敏感度高", "標記", "來源", "說明", "來源誘因"])
rows = []
sec = lambda t: rows.append(("sec", t))
def inp(key, item, unit, lo, base, hi, tag, src, note="—", inc=None):
    if inc is None:
        inc = INC_CO if tag == "Interested-party" else ""
    rows.append((key, item, unit, lo, base, hi, tag, src, note, inc))

PRQ3 = SRCF["Q3_10Q"]; CALL = SRCF["CALL_Q3"]; K10 = SRCF["K10_25"]
sec("A. 股價、股數、日期")
inp("px", "現價（2026-11-13 收盤）", "US$", None, 9.43, None, "Verified", "工作單 px1113；Yahoo Finance 日線收盤 2026-11-13（markets/invest_map/data/closes_20261020_20261113.csv）：9.43", "v0.4.2 為 2026-11-10 收盤 " + str(fv("mkt.price.20261110")) + "；v0.2.1 現價 7.80（2026-10-06 前後）")
inp("px_date", "現價日（px 的收盤日）", "日期", None, "=DATE(2026,11,13)", None, "Verified", "同 px", "README、Val 與 tp_date 以本列組字")
inp("cutoff", "資料截止日（Q3 10-Q 與法說）", "日期", None, "=DATE(2026,11,10)", None, "Verified", PRQ3, "原則第 10 節第 41 條：三個日期分列")
inp("built", "成品日", "日期", None, "=DATE(2026,11,16)", None, "Verified", "建置日", "—")
inp("sh_base", "普通股流通在外（2026-09-30）", "百萬股", None, fv("bs.sharesOut.2026Q3"), None, "Verified", PRQ3 + "（bs.sharesOut.2026Q3；封面 2026-11-06 為 45.288）", "零履約價買權 5.906 百萬股（2031 交割）未扣除，見風險登錄")
inp("pfw", "其他視同已發行股數（無）", "百萬股", None, 0, None, "Verified", "10-Q 未見預付認股權證", "—")
inp("sh_out", "評價股數基礎", "百萬股", None, "=sh_base+pfw", None, "Derived", "sh_base＋pfw", "可轉債列為負債（淨負債），不計轉股")
inp("rsu_net", "RSU＋選擇權庫藏股法淨股數", "百萬股", None, 0, 2.0, "Assumed", "找不到：本輪未登錄 10-Q 的 RSU 與選擇權期末餘額", "暫以 0 並設 Sens 高 2.0（約 4%）")
inp("val_date", "評價日（Q3 資產負債表日）", "日期", None, "=DATE(2026,9,30)", None, "Verified", PRQ3, "—")
inp("tp_date", "目標價時點（現價日＋12 個月）", "日期", None, "=EDATE(px_date,12)", None, "Derived", "公式", "—")
inp("fy27_end", "Comparable 錨定日（FY2027 年底）", "日期", None, "=DATE(2027,12,31)", None, "Derived", "會計年度＝日曆年", "Comparable 以 FY2028 調整後 EBITDA 為前瞻值")

sec("B. 命題：雲端在役 MW 路徑（情境 A／B／C；IT 口徑）")
inp("mw_ye25", "雲端在役 MW（2025 年底）", "MW", None, 5.5, None, "Interested-party", "legacy_v02/company.json defaults.mwYearEnd（冰島 Blönduós 已簽 5.5 MW-IT，10-Q）；Wattonomics mw_wyfi 2025 同值", "亞特蘭大、加拿大承租 MW 未揭露（下限）")
inp("mw_q2", "雲端在役 MW（2026-06-30；世代占比的基數）", "MW", None, 5.5, None, "Interested-party", "legacy_v02 fleet.openMw（10-Q）", "—")
inp("mw_q3", "雲端在役 MW（2026-09-30）", "MW", 5.5, 6.5, 8.0, "Assumed", "找不到：公司未揭露季末雲端在役 MW；Q3 經常性雲端營收 14.706（扣終止費 3.4）× 4 ÷（Tokenomics 錨 × k 約 9.4）÷ 計費比例 0.92 ≈ 6.8", "公司實況驗證只檢驗 MW 假設與 k 的合併結果")
inp("mw_ye26", "雲端在役 MW（2026 年底；三情境相同）", "MW", None, 14.5, None, "Assumed", "已簽 Q4 起算合約依 Tokenomics 每 GPU 功耗換算：Baseten 1,392 顆 B300 2.858、968＋1,024 顆 B300 約 4.09、128 顆 B300 0.263、冰島 576 顆 B300 1.183（Interested-party）＋ Q3 6.5", "Wattonomics mw_wyfi 2026 年 14.574（舊版個股模型，同源，Coord G1）；公司預期 Q4 雲端營收約與 Q3 持平 18.1（起算集中在年底）")
inp("ctr_mw", "已簽約雲端 GPU 換算 MW（約 20,000 顆扣託管服務 2,048 顆）", "MW", None, 37, None, "Interested-party", CALL + "：已簽專用運算 GPU 合計約 20,000 顆（含 2026 年新簽約 13,000）；以 B300 每 MW 約 487 顆換算（Derived）", "只作 Sites 頁容量對照；公司 2027 年中雲端年化營收目標 365")
MWP = {"A": (40, 55, 70, 85, 15), "B": (35, 43, 51, 59, 8), "C": (25, 25, 25, 25, 0)}
for s in SC:
    a = MWP[s]
    for j, y in enumerate((2027, 2028, 2029, 2030)):
        inp(f"mw{s}_{y % 100}", f"情境 {s}：{y} 年底雲端在役 MW", "MW", None, a[j], None, "Assumed",
            {"A": "已簽約約 37 MW 全數上線並每年新增 15 MW（舊版積極每年 +6 的 2.5 倍，對應 2026 年新簽約速度）", "B": "已簽約約 35 MW 上線（2027 年中年化 365 目標約對應 36 MW）、每年新增 8 MW", "C": "只計已具名合約（雪梨 2.46、巴黎 2.7、Prime VR200 1.97 等）約 25 MW，不再新增"}[s], "公司目標 2027 年中雲端年化營收 365（Interested-party）" if s == "B" and y == 2027 else "—")
    inp(f"add{s}_post", f"情境 {s}：2031–2032 每年新增雲端 MW", "MW／年", None, a[4], None, "Assumed", "延續 2028–2030 速度", "之後的永續成長由 DCF 的 g 處理")
inp("u_bill", "計費比例（計費 MW ÷ 平均在役 MW）", "x", 0.85, 0.92, 0.96, "Assumed", "沿用 CoreWeave v5.0 樣板", "公司未揭露計費 MW（找不到）")

sec("C. 命題：世代組合與汰換（Assumed）")
for g, v, lo, hi in (("h100", 0.6, 0.55, 0.65), ("gb200", 0.35, 0.30, 0.40)):
    inp(f"om_{g}", f"2026-06-30 雲端在役世代占比：{GNAME[g]}", "%", lo, v, hi, "Assumed", "legacy_v02/company.json fleet.openMix（Hopper 60%／GB200 類 35%／GB300 5%）", "公司未揭露機隊世代組合（找不到）；H200 以 H100 欄代表")
inp("om_gb300", "2026-06-30 雲端在役世代占比：GB300 NVL72（＝1 − 其餘）", "%", None, "=1-om_h100-om_gb200", None, "Derived", "合計 100%", "—")
inp("nm26_gb300", "2026 下半年新增產能中 GB300 占比（其餘 GB200）", "%", None, 0.95, None, "Derived", "legacy_v02 fleet.newMix FY26（GB300 95.3%／GB200 4.7%）；B300 以 GB300 欄近似", "—")
inp("nm27_gb300", "2027 年新增產能中 GB300 占比（其餘 VR200）", "%", 0.7, 0.85, 0.95, "Assumed", "2027 起算合約以 B300 為主（雪梨 1,200、10 月全球客戶），Prime Intellect 576 顆 VR200（Interested-party）", "—")
inp("nm28_gb300", "2028 年起新增產能中 GB300 占比（其餘 VR200）", "%", None, 0.0, 0.3, "Assumed", "legacy_v02 fleet.newMix（FY29 起 VR200 100%）", "—")
for k, y, v in (("ret29", 2029, 0), ("ret30", 2030, 2.5), ("ret31", 2031, 3.0)):
    inp(k, f"{y} 年汰換 MW（{y - 6} 年新增批次，換成 VR200）", "MW", None, v, None, "Assumed", "2025 年底 5.5 MW 依年份拆（2024 年 2.5、2025 年 3.0；公司未揭露）", "經濟壽命 6 年（Tokenomics IF_DeprLifeIT 基準）")
inp("ret32", "2032 年汰換 MW（2026 年新增批次）", "MW", None, "=mw_ye26-mw_ye25", None, "Derived", "2026 年新增", "—")
inp("v25_h100", "2025 年新增批次中 H100（H200）占比", "%", None, 0.25, None, "Assumed", "2025 年新增以 B200 為主（Initial Customer 464 顆 B200 等）", "使世代存量不為負")
inp("v25_gb200", "2025 年新增批次中 GB200 占比", "%", None, 0.75, None, "Assumed", "同上", "—")
inp("v26_gb300", "2026 年新增批次中 GB300 占比（其餘 GB200）", "%", None, 0.99, None, "Assumed", "2026 新增幾乎全為 B300", "使 2032 年汰換後 GB200 存量不為負")

sec(f"D. 命題：每 MW 經濟性（Tokenomics {TKV} 快照，基準成本情境）與定價")
TKSRC = f"whitefiber/data/tokenomics_snapshot_{TKV}.json（Tokenomics {TKV}，{TKS['source']['commit'][:7]}）"
for g in GEN:
    inp(f"he_{g}", f"每 MW 年持有成本（經濟口徑）IF_HoldEcon｜{GNAME[g]}", "US$M／MW·年", tkv("IF_HoldEcon", g, "低成本"), tkv("IF_HoldEcon", g), tkv("IF_HoldEcon", g, "高成本"), "Derived", TKSRC + " Interface!13 列", "低／高＝Tokenomics 低成本／高成本情境（只作敏感度，鍵 he_all）")
for g in GEN:
    inp(f"cx_{g}", f"每 MW IT 資本支出 IF_CapexIT｜{GNAME[g]}", "US$M／MW", None, tkv("IF_CapexIT", g), None, "Derived", TKSRC + " Interface!9 列", "只含 IT（雲端機房向第三方承租）")
for g in GEN:
    inp(f"ox_{g}", f"每 MW 年營運費用（不含折舊、租金）IF_OpexGW｜{GNAME[g]}", "US$M／MW·年", None, tkv("IF_OpexGW", g), None, "Derived", TKSRC + " IF_OpexGW", "營運成本以 Tokenomics 滿載穩態為主值")
inp("fac_cx", "每 MW 自建機房資本支出 L1_FacCapexMW", "US$M／MW", None, round(TK["L1_FacCapexMW"]["values"] if not isinstance(TK["L1_FacCapexMW"]["values"], dict) else 12.67, 4), None, "Derived", TKSRC + " L1!D10", "雲端自建比例 0，只作對照")
inp("life_ai", "AI 設備折舊年限（會計口徑）", "年", 4, 5, 6, "Interested-party", "legacy_v02/company.json defaults.gpuLife 5 年（10-K 伺服器耐用年限）", "高＝Tokenomics IF_DeprLifeIT 基準 6 年")
inp("fac_life", "機房與託管建物折舊年限", "年", None, 20, None, "Assumed", "一般廠房 15–25 年（判斷）", "託管資本支出依此年限折舊")
inp("p_long", "長約價格證據：IREN–Microsoft GB300 五年約（Analogy）", "US$M／MW·年", 8.9073, 9.70, 12.7247, "Analogy", "IREN 新聞稿 2025-11-03：$9.7bn ÷ 5 年 ÷ 200 MW（critical IT）", "原則第 10 節第 61 條；WhiteFiber 自身合約換算每 MW 需假設 GPU 功耗（K 節 p_claim，只作對照）")
inp("p_spot", "現貨價格證據：H100 Silicon Data 指數", "US$／GPU-hr", None, 2.82, None, "Verified", "Silicon Data H100 Index（2026-10-06；同 CoreWeave v5.0）", "—")
inp("gh_h100", "每 GPU 小時持有成本 IF_GPUhrEcon｜Hopper H100", "US$／GPU-hr", None, round(TK["IF_GPUhrEcon"]["values"]["Hopper H100"]["基準"], 5), None, "Derived", TKSRC, "—")
inp("od_share", "隨需（現貨）占計費產能比例", "%", None, 0.0, 0.20, "Assumed", "新簽約皆為專用運算多年約（平均 4.2 年，Interested-party）；隨需比例未揭露", "高 20% 只作敏感度")
inp("k_long", "長約定價倍數 k_長約（＝長約價 ÷ GB300 持有成本）", "x", None, "=p_long/he_gb300", None, "Derived", "每 MW 收入＝IF_HoldEcon × 公司因素 k；不以營收、RPO、合約價反推 k", "—")
inp("k_spot", "現貨定價倍數 k_現貨", "x", None, "=p_spot/gh_h100", None, "Derived", "H100 現貨指數 ÷ H100 每 GPU 小時持有成本", "—")
inp("k_mix", "定價倍數 k（隨需占比加權；單一 k）", "x", None, "=od_share*k_spot+(1-od_share)*k_long", None, "Derived", "k＝隨需占比 × k_現貨＋（1 − 隨需占比）× k_長約", "—")
inp("k_adj", "k 調整係數（敏感度與反解用；基準 1）", "x", 0.9, 1.0, 1.2, "Assumed", "敏感度 ×0.9／×1.2", "RevDCF 以本鍵反解")
inp("p_new_ovr", "新增產能每 MW 年收入改用公司合約換算值（0＝不用；Sens）", "US$M／MW·年", None, 0, 17.5, "Assumed", "K 節 p_claim（已簽合約 TCV ÷ 年數 ÷ Tokenomics 功耗換算 MW，Interested-party）", "只進 Sens")
inp("lam", "成長型資本支出提前比例 λ", "%", 0.20, 0.35, 0.50, "Assumed", "沿用 legacy_v02 defaults.lambda", "可證偽指標：PP&E 訂金 ÷ 次年新增產能資本支出")
inp("own_share", "雲端自建機房比例（其餘向第三方承租）", "%", None, 0.0, None, "Interested-party", "雲端站點（冰島、安大略、巴黎、雪梨、亞特蘭大）皆為第三方機房（10-Q）", "託管自有站點另在 H 節")
inp("rent_mw", "雲端承租機房每 MW 年租金", "US$M／MW·年", 1.5, 1.9, 2.4, "Analogy", "WhiteFiber 自身託管租金：NC-1 第 1 年 1.886、MTL-3 2.35（legacy_v02 事實總帳 ct.nscale.perMW.y1、ct.cerebras.perMW）", "第三方機房租金未逐站揭露（找不到）")
inp("sga_r", "營業費用率（一般及管理，扣 SBC；含託管與總部）", "%", 0.15, 0.20, 0.40, "Assumed", PRQ3 + "：Q3 一般及管理 17.998 − 股票薪酬約 3.9＝14.1 ÷ 營收 24.0＝59%", "規模效應：公司目標 2027 年中雲端年化 365；基準 20%（Assumed）；高案 40%")
inp("sbc_r", "股票薪酬 ÷ 營收", "%", 0.08, "=sbc_q3/rev_q3", 0.25, "Derived", PRQ3 + "：Q3 股票薪酬（9M 14.94 − Q1 7.346 − Q2 3.671）÷ Q3 營收", "全期持平；DCF 以 SBC 為現金成本")
inp("q4cx_ovr", "Q4 資本支出改用其他值（0＝不用；Sens）", "US$M", None, 0, 400, "Assumed", "公司未給 2026 資本支出指引；高＝400（約 9M 年化）", "只作敏感度")
inp("pp_q4", "Q4 額外客戶預付（0＝不計）", "US$M", None, 0, 100, "Assumed", "Q4 預付已由覆蓋比 × Q4 資本支出計入；10 月全球客戶 9 個月服務費存入托管帳戶（金額未揭露）", "只作敏感度")

sec("E. 實績錨點（Q3 10-Q；FY2026＝9M 實績＋Q4 模型）")
for k, fid, item in (("rev9m", "is.rev.9M26", "合併營收（9M26）"), ("oi9m", "is.opInc.9M26", "營業利益（GAAP，9M26）"), ("aeb9m", "ng.adjEbitda.9M26", "調整後 EBITDA（9M26；公司定義）"),
                     ("ni9m", "is.ni.9M26", "淨損（9M26；含誘導轉換費用 80.8）"), ("da9m", "is.da.9M26", "折舊攤銷（9M26）"), ("sbc9m", "cf.sbc.9M26", "股票薪酬（9M26）"), ("cfo9m", "cf.cfo.9M26", "營運現金流（9M26；含客戶預付 226.5）"),
                     ("cx9m_cash", "cf.capex.9M26", "現金資本支出（購置 PP&E 與設備訂金，9M26）"), ("rev_q3", "is.rev.2026Q3", "合併營收（Q3 2026）"), ("aeb_q3", "ng.adjEbitda.2026Q3", "調整後 EBITDA（Q3 2026）"),
                     ("da_q3", "is.da.2026Q3", "折舊攤銷（Q3 2026）")):
    t = FX[fid]["tag"]
    inp(k, item, "US$M", None, fv(fid), None, t, SRCF[FX[fid]["src"]] + f"（{fid}）", FX[fid].get("note", "—"), INC_CO if t == "Interested-party" else "")
inp("sbc_q3", "股票薪酬（Q3 2026）", "US$M", None, "=sbc9m-7.346-3.671", None, "Derived", PRQ3 + "：9M 14.94 − Q1 7.346 − Q2 3.671（調整後 EBITDA 調節表口徑）", "—")
inp("intx_q3", "利息費用（Q3 2026；第三方 6.426＋關係人 3.987）", "US$M", None, "=6.426+3.987", None, "Verified", PRQ3, "含 DDTL 1.1× MOIC 攤提")
inp("int9m", "利息費用（9M26）", "US$M", None, fv("is.int.9M26"), None, "Verified", PRQ3, "利息收入未登錄（不外推）")
inp("int_q3", "利息費用（Q3 2026）", "US$M", None, "=intx_q3", None, "Derived", "同上", "—")
inp("aoi9m", "調整後營業利益（9M26；營業利益＋股票薪酬，Derived）", "US$M", None, "=oi9m+sbc9m", None, "Derived", "公司未揭露調整後營業利益", "—")
for k, fid, item in (("ai1", "is.cloud.2026Q1", "雲端服務營收（Q1 2026）"), ("ai2", "is.cloud.2026Q2", "雲端服務營收（Q2 2026；含 Initial Customer 終止費 12.3）"), ("ai3", "is.cloud.2026Q3", "雲端服務營收（Q3 2026；含終止費追加 3.4）")):
    inp(k, item, "US$M", None, fv(fid), None, "Verified", SRCF[FX[fid]["src"]] + f"（{fid}）", "10-Q 分部營收", "")
inp("rev9m_ai", "雲端服務營收（9M26；分部）", "US$M", None, "=ai1+ai2+ai3", None, "Derived", "Q1＋Q2＋Q3", "命題模組 9M 實績；合併營收與其差額＝託管與其他")
inp("rev_q3_ai", "雲端服務營收（Q3 2026）", "US$M", None, "=ai3", None, "Derived", "同上", "—")
inp("nonrec_q3", "Q3 雲端非經常性（Initial Customer 終止費追加）", "US$M", None, fv("rev.termFee.2026Q3"), None, "Verified", PRQ3, "Q2 另有終止費 12.3")
inp("cx_q3_def", "公司口徑資本支出（Q3 2026；現金購置與訂金，9M − 1H 344.7）", "US$M", None, "=cx9m_cash-344.7", None, "Derived", PRQ3 + "；1H 344.7 取 legacy_v02/company.json defaults.prepay.note（Q2 10-Q）", "公司未另定應計口徑")
inp("cx_1h_def", "公司口徑資本支出（1H26）", "US$M", None, 344.7, None, "Verified", SRCF["Q2_10Q"] + "（legacy_v02 defaults.prepay.note）", "—")
inp("cx9m_def", "公司口徑資本支出（9M26）＝1H＋Q3", "US$M", None, "=cx_1h_def+cx_q3_def", None, "Derived", "＝9M 現金購置", "FY2026 公司口徑＝本列＋Q4 模型")
inp("ppe_q3", "PP&E 淨額（2026-09-30）", "US$M", None, fv("bs.ppe.2026Q3"), None, "Verified", PRQ3, "Thesis 投入資本起點（含託管建物）")
inp("cip_q2", "在建工程（2026-09-30；Q4 起依 cip_life 年折舊）", "US$M", None, fv("bs.cip.2026Q3"), None, "Verified", PRQ3 + "：在建工程 312.326（另有 PP&E 訂金 158.437，屬 GPU 預付，對應的上線產能已在模型資本支出內，不重複折舊）", "多為 NC-1 第二期與 MTL-2 建物（判斷）")
inp("cip_life", "在建工程折舊年限", "年", 5, 20, None, "Assumed", "NC-1、MTL-2 建物為主（機房年限）；低＝全部以 AI 設備 5 年", "—")
inp("def_open", "遞延收入（客戶預付）餘額（2026-09-30）", "US$M", None, "=54.636+251.398", None, "Verified", PRQ3 + "：流動 54.636＋非流動 251.398", "自 Q4 起依 prepay_yrs 年直線認列（非現金營收）")

sec("F. 現金流、融資")
inp("cash_q3", "現金與約當現金（2026-09-30）", "US$M", None, fv("bs.cash.2026Q3"), None, "Verified", PRQ3, "期後：DDTL 追加動用 70（10-27）、NC-2／NC-3 收購付現 60（10-28）、NC-1 信用協議 300 簽約未動用（11-09）")
inp("rcash_q3", "受限現金（2026-09-30）", "US$M", None, fv("bs.restricted.2026Q3"), None, "Verified", PRQ3, "—")
inp("debt_q3", "借款帳面值（含可轉債，2026-09-30）", "US$M", None, "=330.088+11.873+29.989+30.683", None, "Verified", PRQ3 + "：可轉債 330.088、短期第三方 11.873、關係人 DDTL 29.989、長期第三方 30.683", "期後借款不計入評價日淨負債")
inp("finl_q3", "融資租賃負債（2026-09-30）", "US$M", None, 0, None, "Assumed", "找不到：本輪未登錄期末融資租賃餘額（9M 償還 12.598）", "—")
inp("opl_q3", "營業租賃負債（2026-09-30）", "US$M", None, fv("bs.opLeaseLiab.2026Q3"), None, "Verified", PRQ3, "—")
inp("lease_nc", "未起租租約（加拿大託管 2.7 MW、60 個月）", "US$M", None, fv("lease.notCommenced.2026Q3"), None, "Verified", PRQ3, "—")
inp("hold", "非上市權益投資（無）", "US$M", None, 0, None, "Verified", "10-Q 未見重大權益投資", "—")
inp("r_old", "既有借款平均利率（Q3 利息費用 × 4 ÷ 借款）", "%", None, "=intx_q3*4/debt_q3", None, "Derived", PRQ3, "含 DDTL MOIC 攤提")
inp("kd", "稅前債務成本（邊際純債）", "%", 0.079, 0.095, 0.106, "Analogy", "legacy_v02 事實總帳 val.kd.default：Bit Digital DDTL 票面 9.5%；區間＝冰島貸款票面 7.92%–有效 10.64%；NC-1 信用協議 SOFR＋6.00% 約 9.89%（期後）", "不用可轉債票面（含轉換選擇權價值）")
inp("r_new", "瀑布新增投資級債利率", "%", None, "=kd", None, "Derived", "同 kd", "—")
inp("r_junk", "高息債利率", "%", 0.15, 0.12, 0.18, "Assumed", "沿用 CoreWeave v5.0 樣板 12%", "—")
inp("min_cash", "最低現金", "US$M", None, 25, None, "Assumed", "規模比例（判斷）", "—")
inp("lev_max", "投資級債上限：借款 ÷ 調整後 EBITDA（年化）", "x", 4.0, 4.5, 5.0, "Assumed", "原則第 10 節第 9 條預設 4.5", "—")
inp("icov_min", "利息保障倍數下限（只 WARN）", "x", None, 3, None, "Assumed", "原則第 10 節第 9 條預設 3", "—")
inp("eq_cap", "每年新股上限（占市值）", "%", 0.10, 0.20, 0.30, "Assumed", "沿用 legacy_v02 defaults.eqCapPct 20%", "—")
inp("eq_disc", "新股發行折價", "%", None, 0.10, None, "Assumed", "沿用 legacy_v02 defaults.eqDisc", "—")
inp("prepay_cov", "客戶預付覆蓋比（預付 ÷ 資本支出）", "%", 0.20, 0.40, 0.50, "Interested-party", CALL + "：已簽未上線專用運算合約預付約占估計資本支出 40%；9M 實績 226.5 ÷ 549.8＝41.2%（Derived）", "Q4 起套用", INC_CO)
inp("prepay_yrs", "預付認列年數", "年", None, 5, None, "Derived", "遞延收入 306.0 ÷ 流動部分 54.6 ≈ 5.6（取 5）", "—")
inp("nwc_r", "營運資金增加 ÷ 營收增量（預付除外）", "%", 0.0, 0.02, 0.04, "Assumed", "沿用 legacy_v02 valuation.wcPctOfRevGrowth 2%", "—")
inp("tax_lt", "長期現金稅率", "%", None, 0.21, 0.265, "Assumed", "美國聯邦 21%（legacy_v02 valuation.tax）；加拿大約 26.5%", "Sens 26.5%")
inp("nol0", "淨營業虧損（NOL）可用餘額", "US$M", None, 0, 100, "Assumed", "legacy_v02 valuation.nol 0（加拿大、日本遞延所得稅資產全額備抵）", "基準 0；高 100 只作敏感度")
inp("nol_lim", "NOL 每年可抵稅前所得上限", "%", None, 0.8, None, "Verified", "美國稅法（TCJA 後 80%）", "—")

sec("G. 評價")
inp("rf", "無風險利率（美國 10 年期公債，2026-11-10 收盤）", "%", None, 0.0498, None, "Verified", SRCF["UST"], "與現價日相差 3 個交易日（工作單 px1113 不改 rf）")
inp("beta_u", "無槓桿 β（自算產業 β：US 伺服器／EMS 代理）", "x", 1.022, 1.322, 1.622, "Analogy", "docs/plan/自算產業β_20261105.md（同 CoreWeave、Nebius 新雲樣板）", "舊版 v0.2.1 用 β 2.5（個股 β 0.84–4.77 不可靠）；判斷題")
inp("beta", "β（依市值 D/E 重新加槓桿）", "x", None, "=beta_u*(1+(1-tax_lt)*(debt_q3+finl_q3)/(px*(sh_out+rsu_net)))", None, "Derived", "Hamada", "—")
inp("erp", "股權風險溢酬", "%", 0.04, 0.05, 0.06, "Assumed", "DCF 標準規格預設 5%", "—")
inp("wacc_adj", "WACC 調整（敏感度用；基準 0）", "pt", -0.01, 0.0, 0.01, "Assumed", "Sens：±1 個百分點", "—")
inp("ronic_adj", "永續期 RONIC − WACC（基準 0）", "pt", None, 0.0, 0.05, "Assumed", "DCF 標準規格：RONIC＝WACC", "—")
inp("tp_g", "永續成長率 g（＝MIN(3%, 10 年期)）", "%", 0.02, "=MIN(0.03,rf)", 0.04, "Derived", "g＝MIN(3%, rf)", "—")
inp("w_dcf", "目標價中 DCF 的權重（其餘 Comparable）", "%", 0.3, 0.5, 0.7, "Assumed", "CoreWeave v5.0 樣板：DCF 50%＋Comparable 50%", "Andy 可改")
for s, w in (("A", 0.25), ("B", 0.5), ("C", 0.25)):
    inp(f"w{s}", f"目標價情境權重 {s}", "%", None, w, None, "Assumed", "CoreWeave v5.0 樣板：25／50／25", "Andy 可改")
PEER_D = "2026-11-13"
MSS = SRCF["MS"]
inp("pc_px", "CoreWeave 股價（2026-11-13 收盤）", "US$", None, 53.13, None, "Verified", "工作單 px1113；Yahoo Finance 日線收盤 2026-11-13（closes_20261020_20261113.csv）；同 coreweave/build_coreweave.py In px（v5.0.5）", "同業估值日 2026-11-13＝本模型現價日；v0.4.2 為 2026-11-11 收盤 60.77")
inp("pc_sh", "CoreWeave 股數（Q3 基本加權）", "百萬股", None, 559, None, "Verified", SRCF["CRWV_PR"], "—")
inp("pc_nd", "CoreWeave 淨負債：最新申報（2026-09-30）", "US$M", None, "=46873+649-4122", None, "Derived", SRCF["CRWV_PR"] + "：借款 46,873、融資租賃 649、現金 4,122", "—")
inp("pc_ndf", "CoreWeave 淨負債：共識 FY2026 預估（財報前）", "US$M", None, 46442, None, "Interested-party", SRCF["CONS_CRWV"], "替代口徑", INC_SELL)
inp("pc_eb", "CoreWeave NTM 共識 EBITDA", "US$M", None, "=7574*1.6/12+16656*10.4/12", None, "Interested-party", SRCF["CONS_CRWV"] + "：FY26 7,574、FY27 16,656、FY28 28,857", "原則 49 月份加權", INC_SELL)
inp("pc_rev", "CoreWeave NTM 共識營收", "US$M", None, "=12882*1.6/12+26376*10.4/12", None, "Interested-party", SRCF["CONS_CRWV"], "—", INC_SELL)
inp("pc_cx", "CoreWeave NTM 共識資本支出", "US$M", None, "=35953*1.6/12+46476*10.4/12", None, "Interested-party", SRCF["CONS_CRWV"], "—", INC_SELL)
inp("pn_px", "Nebius 股價（2026-11-13 收盤）", "US$", None, 115.88, None, "Verified", "工作單 px1113；Yahoo Finance 日線收盤 2026-11-13（closes_20261020_20261113.csv）：115.88", "v0.4.2 為 2026-11-11 收盤 154.69（MarketScreener Nasdaq 5 日表）")
inp("pn_sh", "Nebius 普通股流通在外（2026-09-30）", "百萬股", None, 292.902405, None, "Verified", SRCF["NBIS_6K"], "—")
inp("pn_nd", "Nebius 淨負債：最新申報（2026-09-30）", "US$M", None, "=14190.8-7425.6", None, "Derived", SRCF["NBIS_6K"] + "：借款 14,190.8 − 現金 7,425.6", "—")
inp("pn_ndf", "Nebius 淨負債：MarketScreener 2026 年預估", "US$M", None, 13659, None, "Interested-party", MSS, "替代口徑", INC_SELL)
inp("pn_eb", "Nebius NTM 共識 EBITDA", "US$M", None, "=1348*1.6/12+6556*10.4/12", None, "Interested-party", MSS + "：2026 1,348、2027 6,556、2028 13,700", "原則 49 月份加權", INC_SELL)
inp("pn_rev", "Nebius NTM 共識營收", "US$M", None, "=3379*1.6/12+11418*10.4/12", None, "Interested-party", "nebius/data/consensus_nbis_20261007.json（MarketScreener）", "—", INC_SELL)
inp("ir_px", "IREN 股價（2026-11-13 收盤）", "US$", None, 18.17, None, "Verified", "Yahoo Finance 日線收盤 2026-11-13（`markets/invest_map/data/closes_20261020_20261113.csv`，IREN 2026-11-13＝18.17）；v0.4.3 初稿為 2026-11-11 收盤 20.84（MarketScreener Nasdaq 5 日表 11-11）；股數與淨負債來源："+MSS, "工作單 px1113：IREN 價格檔補齊後換 11-13 收盤 18.17；EV＝股價 × 股數＋淨負債，依公式重算（等同 EV 新＝EV 舊＋市值舊 ×（18.17 ÷ 20.84 − 1））；股數、淨負債、共識 EBITDA 不變")
inp("ir_sh", "IREN 股數", "百萬股", None, 424.357, None, "Derived", MSS, "—")
inp("ir_nd", "IREN 淨負債：最新申報（2026-09-30）", "US$M", None, "=10208-3522", None, "Derived", SRCF["IREN_Q1"], "—")
inp("ir_ndf", "IREN 淨負債：MarketScreener FY2027 預估", "US$M", None, 18889, None, "Interested-party", MSS, "替代口徑", INC_SELL)
inp("ir_eb", "IREN NTM 共識 EBITDA", "US$M", None, "=1395*7.6/12+5123*4.4/12", None, "Interested-party", MSS, "原則 49 月份加權", INC_SELL)
inp("ir_use", "IREN 進同業中位數（1＝是）", "—", None, 1, None, "Assumed", "同 CoreWeave v5.0", "—")
inp("po_px", "Oracle 股價（2026-11-13 收盤）", "US$", None, 125.35, None, "Verified", "工作單 px1113；Yahoo Finance 日線收盤 2026-11-13（closes_20261020_20261113.csv）：125.35", "v0.4.2 為 2026-11-11 收盤 128.38（MarketScreener）")
inp("po_sh", "Oracle 股數", "百萬股", None, 3117.708, None, "Derived", MSS, "—")
inp("po_nd", "Oracle 淨負債：最新實際（含融資租賃；2026-11 增資 pro forma）", "US$M", None, 87208, None, "Derived", "Oracle v0.5.1 Val!C10（同 CoreWeave v5.0 In po_nd）", "主口徑")
inp("po_ndf", "Oracle 淨負債：MarketScreener FY2027 預估", "US$M", None, 115597, None, "Interested-party", MSS, "替代口徑", INC_SELL)
inp("po_eb", "Oracle NTM 共識 EBITDA", "US$M", None, "=52090*6.6/12+74148*5.4/12", None, "Interested-party", MSS, "原則 49 月份加權", INC_SELL)
inp("peer_nd_mode", "同業淨負債口徑（0＝最新實際；1＝共識預估）", "—", None, 0, 1, "Assumed", "主口徑同 Oracle 樣板", "Sens：改用預估口徑")
inp("po_rev", "Oracle NTM 共識營收", "US$M", None, "=90423*7/12+131711*5/12", None, "Interested-party", "oracle/data/consensus_orcl_20261022.json", "—", INC_SELL)
inp("po_cx", "Oracle NTM 共識資本支出", "US$M", None, "=93961*7/12+100962*5/12", None, "Interested-party", "oracle/In cx2027、cx2028", "—", INC_SELL)
inp("rel_mult", "相對倍數調整係數", "x", None, 1.0, None, "Assumed", "1.00（未調整：缺 WhiteFiber 10 年歷史倍數，上市不足 2 年）", "原則第 10 節第 18、47 條")
inp("mult_adj", "Comparable 倍數調整（敏感度用；基準 1；雲端與託管兩分部同時）", "x", 0.8, 1.0, 1.2, "Assumed", "敏感度 ±20%", "—")
# 託管分部同業（Andy 2026-11-15：同業法改為分部加總；託管用資料中心／託管同業倍數，雲端用新雲同業倍數）
inp("sotp_on", "Comparable 分部加總（1＝託管 × 資料中心同業、雲端 × 新雲同業；0＝v0.4.0 全公司 × 新雲同業）", "—", None, 1, None, "Assumed", "Andy 2026-11-15 決定改為分部加總", "0 只供拆解（scripts/decomp_v50.py）")
inp("dlr_ev", "Digital Realty（DLR）企業價值（2026-10-07 EV 86,198 × 收盤價比 11-13 ÷ 10-07；Derived 近似）", "US$M", None, "=86198*179.99/180.47", None, "Derived", SRCF["MS_DLR"] + "：2026-10-07 EV 86,198（股價 180.47）；Yahoo Finance 日線收盤 DLR 10-07＝180.47、11-13＝179.99（`closes_20261020_20261113.csv`，9-28～10-19 段已補）", "工作單 px1113：模型無 DLR 市值，倍數 × 價格比近似（淨負債不隨股價變，略高估變動；共識 EBITDA 不變）；v0.4.2 為 EV 86,198（倍數 20.42）")
inp("dlr_eb", "Digital Realty NTM 共識 EBITDA（2026 × 0.25＋2027 × 0.75）", "US$M", None, "=3841*0.25+4349*0.75", None, "Interested-party", SRCF["MS_DLR"] + "：2026 3,841、2027 4,349", "原則 49 月份加權（估值日 2026-10 起 12 個月）", INC_SELL)
inp("eqix_ev", "Equinix（EQIX）企業價值（約 2026-10-02 EV 121,482 × 收盤價比 11-13 ÷ 10-02；Derived 近似）", "US$M", None, "=121482*1053.79/1025.72", None, "Derived", SRCF["MS_EQIX"] + "：約 2026-10-02 EV 121,482（頁面股價 1,012.33）；Yahoo Finance 日線收盤 EQIX 10-02＝1,025.72、11-13＝1,053.79（`closes_20261020_20261113.csv`）；依工作單以報價日收盤計價格比（頁面股價與 10-02 收盤差 1.3%，未另調）", "工作單 px1113：模型無 EQIX 市值，倍數 × 價格比近似（淨負債不隨股價變，略高估變動；共識 EBITDA 不變）；v0.4.2 為 EV 121,482（倍數 21.06）")
inp("eqix_eb", "Equinix NTM 共識 EBITDA（2026 × 0.25＋2027 × 0.75）", "US$M", None, "=5286*0.25+5930*0.75", None, "Interested-party", SRCF["MS_EQIX"] + "：2026 5,286、2027 5,930", "原則 49 月份加權", INC_SELL)
inp("colo_disc", "託管倍數相對資料中心同業中位數的折價（小型、單一主要承租人、非 REIT）", "%", None, 0.0, 0.28, "Assumed", "基準不折價（Andy 2026-11-15 指定用資料中心／託管同業倍數）；高＝28%（約 15 倍，data/whitefiber_facts_20261008.json val.peer.coloDefault 下緣）", "只進 Sens")
inp("tp_v040", "v0.4.0 加權目標價（全公司 × 新雲同業中位數）", "US$", None, 5.39, None, "Derived", "whitefiber/dist/_待刪除/20261115_WhiteFiber個股模型_v0.4.0.xlsx（Val 加權目標價 5.3855；投資建議放空）", "只作對照")

sec("H. 託管分部（NC-1、MTL；Seg 其他事業列、Thesis 託管列）")
inp("colo_q4", "託管＋其他營收（Q4 2026）", "US$M", None, 21.8, None, "Derived", "NC-1 Nscale 10 年約 865 ÷ 10 ÷ 40 MW＝每 MW 年 2.1625（直線）：第一期 20 MW 全季 10.81＋第二期 20 MW 自 11-09 起 6.16＋MTL-1／MTL-3 以 Q3 水準 4.52＋其他 0.3（Interested-party 合約、Derived）", "—")
inp("colo_core", "託管核心營收（FY2027 起每年；NC-1 40 MW 直線＋MTL）", "US$M／年", 90, 105.8, 115, "Derived", "NC-1 86.5（865 ÷ 10）＋MTL-1／MTL-3 18.1（Q3 扣 NC-1 1.1 後 × 4）＋其他 1.2；到期以同價續約", "NC-1 第二期擴充、NC-2／NC-3（初始 60 毛 MW、2027 年底 RFS、僅意向書）未計入")
inp("mgd_27", "代管服務營收（FY2027；客戶自有 2,048 顆 B300）", "US$M", 0, 30, 45, "Interested-party", CALL + "：代管服務年化經常性營收目標 50（2027 年第一季末）；FY2027 取約 0.6 年", "公司目標，基準打折", INC_GUIDE)
inp("mgd_28", "代管服務營收（FY2028 起每年）", "US$M／年", 0, 50, 60, "Interested-party", CALL + "：年化 50（目標）", "—", INC_GUIDE)
inp("colo_m", "託管與代管 EBITDA 率（扣一般及管理前）", "%", 0.50, 0.65, 0.75, "Assumed", PRQ3 + "：Q3 託管分部毛利 3.208 ÷ 營收 5.617＝57%；NC-1 為修正毛租（電費轉嫁）", "營業費用率另計")
inp("cxo_q4", "託管資本支出（Q4 2026；NC-2／NC-3 收購 60＋MTL-2）", "US$M", None, 70, None, "Interested-party", SRCF["Q3_10Q"] + "：期後 NC-2／NC-3 收購付現 60；MTL-2 預算約 25（法說）", "—")
inp("cxo_27", "託管資本支出（FY2027；MTL-2 完工）", "US$M", None, 15, 60, "Assumed", "MTL-2 預算約 25 的剩餘部分；高＝加 NC-2／NC-3 初期工程", "NC-2／NC-3 開發（約 9／毛 MW）未簽約，不計入基準")
inp("cxo_28", "託管維護資本支出（FY2028 起每年）", "US$M／年", None, 5, None, "Assumed", "判斷", "—")

sec("K. 共識、公司指引與對照（只作 Recon 與 Sens，不進模型）")
inp("g_q4cloud", "公司預期 Q4 2026 雲端營收（約與 Q3 持平）", "US$M", None, fv("guid.q4Cloud"), None, "Interested-party", CALL, "只作對照", INC_GUIDE)
inp("g_arr27", "公司目標：簽約雲端組合全部上線後年化營收（2027 年中）", "US$M", None, fv("guid.arr2027"), None, "Interested-party", CALL, "只作對照", INC_GUIDE)
inp("g_rev_lo", "FY2026 合併營收：9M 實績＋公司 Q4 雲端預期＋模型 Q4 託管（對照用）", "US$M", None, "=rev9m+g_q4cloud+colo_q4", None, "Derived", "公司未給全年營收指引；以 Q4 雲端預期組合", "只作對照", INC_GUIDE)
inp("g_rev_hi", "同上（上緣取同值）", "US$M", None, "=g_rev_lo", None, "Derived", "同上", "只作對照", INC_GUIDE)
inp("rpo", "雲端 RPO（2026-09-30）", "US$M", None, fv("rpo.cloud.2026Q3"), None, "Verified", PRQ3 + "（RPO 合計 728.982；託管 RPO 12.697；另有出租方租賃收款 421.49）", "只作對照", "")
inp("arr", "新簽雲端合約價值（2026 年五家客戶）", "US$M", None, fv("kpi.newTcv"), None, "Interested-party", CALL, "只作對照", INC_CO)
for k, v, lab in (("rpo_24", 0.0, "未用（RPO 年度分布已揭露）"), ("rpo_48", 0.0, "未用"), ("rpo_l", 0.0, "未用")):
    inp(k, f"RPO 分桶：{lab}", "%", None, v, None, "Verified", PRQ3 + "：雲端 RPO 年度分布（RPO 頁直接列）", "只作對照")
inp("p_claim_lo", "公司合約換算每 MW 年收入（下緣：冰島 576 顆 B300）", "US$M／MW·年", None, 14.8, None, "Interested-party", "legacy_v02 事實總帳 ct.icelandB300（TCV ÷ 年數 ÷ Tokenomics 功耗換算 MW，Derived）", "只作對照", INC_CO)
inp("p_claim_hi", "公司合約換算每 MW 年收入（上緣：128 顆 B300）", "US$M／MW·年", None, 20.29, None, "Interested-party", "legacy_v02 事實總帳 ct.b300_128（另 Baseten 19.27、Prime VR200 18.27）", "只作對照", INC_CO)
inp("p_short", "短期 GPU 合約每 MW 年收入（Prime GB200 12 個月）", "US$M／MW·年", None, 15.8, None, "Interested-party", "legacy_v02 事實總帳 ct.primeGB200", "只作對照", INC_CO)
inp("cons_eb27", "共識 FY2027 EBITDA（未取得，以營收共識 × 40% 代理）", "US$M", None, "=296*0.4", None, "Interested-party", "legacy_v02 事實總帳 cons.est.sws（FY2027 營收 296；EBITDA 共識找不到）", "只作對照（Comparable 自身對照列）", INC_SELL)
inp("cons_eb26", "共識 FY2026 營收（S&P 經 Simply Wall St）", "US$M", None, 137, None, "Interested-party", "legacy_v02 事實總帳 cons.est.sws", "只作對照", INC_SELL)
inp("cons_tp", "市場共識目標價（11 家平均，2026-11-10）", "US$", None, fv("mkt.consensusTp.20261110"), None, "Interested-party", SRCF["MS_WYFI"], "只作對照", INC_SELL)
inp("tp_v053", "v0.2.1 加權目標價（舊 JS 引擎、現價 7.80）", "US$", None, 15.87, None, "Derived", "whitefiber/dist/_待刪除/20261106_WhiteFiber收支模型_v0_2_1.xlsx（摘要頁；評等中立）", "只作對照")
inp("wat26", "Wattonomics：WhiteFiber（雲端）2026 年底上線 MW", "MW", None, round(WAT0["names"]["WIF_PlayerGW_YE_wyfi_2026"] * 1000, 3), None, "Derived", "wattonomics/data/wattonomics_snapshot_v0.3.5.json WIF_PlayerGW_YE_wyfi_2026（× 1000）", "只作對照（同源：讀 legacy_v02/company.json scenarios.mwPath，Coord G1）")
inp("wat30", "Wattonomics：WhiteFiber（雲端）2030 年底上線 MW", "MW", None, round(WAT0["names"]["WIF_PlayerGW_YE_wyfi_2030"] * 1000, 3), None, "Derived", "同上 2030", "只作對照")

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
        ("scn_mw27", "2027 年底在役 MW", "MW", ("mwA_27", "mwB_27", "mwC_27"), "#,##0", "A＝公司每年超過 1 GW 節奏"),
        ("scn_mw28", "2028 年底在役 MW", "MW", ("mwA_28", "mwB_28", "mwC_28"), "#,##0", "—"),
        ("scn_mw29", "2029 年底在役 MW", "MW", ("mwA_29", "mwB_29", "mwC_29"), "#,##0", "—"),
        ("scn_mw30", "2030 年底在役 MW", "MW", ("mwA_30", "mwB_30", "mwC_30"), "#,##0", "已簽約超過 4 GW（口徑未說明）"),
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
Q["A1"] = "Q — 實績（百萬美元）。會計年度＝日曆年。藍字＝輸入；黑字＝公式。Q1–Q3 2026＝10-Q；FY2024–FY2025＝10-K（數字經 PR #451 事實總帳登錄，原文出處見 Sources）"
Q["A2"] = f"資料取自 {FACTS_F}。官方分部：雲端服務、託管服務（另有其他營收：設備銷售型租賃利息）。雲端 Q2 含 Initial Customer 終止費 12.3、Q3 含追加 3.4"
QH = ["項目", "Q1 2026", "Q2 2026", "Q3 2026", "9M 2026", "FY2024", "FY2025", "標記", "來源", "說明", "來源誘因"]
hdr(Q, 3, QH)
QD = [("合併營收", "rev", (21.923, 28.839, 24.017, 74.779, 47.639, 79.164), "Verified", ""),
      ("營業利益（GAAP）", "oi", ("=E5-C5-D5", -9.255, -12.403, -32.678, 0.628, -26.821), "Verified", "Q2＝營收 − 營業成本 − 折舊 − 減損 5.006 − 一般及管理（Derived）；Q1＝9M − Q2 − Q3"),
      ("調整後 EBITDA（非 GAAP，公司定義）", "aeb", (3.001, 5.535, -0.024, 8.512, 21.926, 17.295), "Interested-party", ""),
      ("雲端服務營收（分部）", "ai", (16.767, 23.806, 18.106, "=B7+C7+D7", 45.728, 68.754), "Verified", "Q2 含終止費 12.3、Q3 含追加 3.4"),
      ("淨損益（GAAP）", "ni", (-12.042, -14.976, -101.435, -128.453, 1.37, -24.683), "Verified", "Q3 含誘導轉換費用 80.767"),
      ("利息費用（第三方＋關係人）", "int", (1.995, "=4.578+1.438", "=6.426+3.987", 18.424, None, None), "Verified", ""),
      ("折舊攤銷", "da", (6.441, 6.567, 8.077, 21.085, 16.511, 23.441), "Verified", ""),
      ("股票薪酬", "sbc", (7.346, 3.671, "=E11-B11-C11", 14.94, 3.171, 19.246), "Verified", "調整後 EBITDA 調節表口徑"),
      ("營運現金流", "cfo", (None, None, None, 211.312, None, 45.655), "Verified", "9M 含客戶預付（遞延營收增加）226.524"),
      ("現金資本支出（購置 PP&E 與設備訂金）", "cx", (None, None, "=E13-344.7", 549.804, None, 268.403), "Verified", "1H 344.7（Q2 10-Q）"),
      ("公司口徑資本支出（未另定應計口徑）", "cxd", (None, None, None, None, None, None), "—", "公司未揭露")]
QR = {}
for i, (lab, k, vals, tag, note) in enumerate(QD):
    rr = 4 + i; QR[k] = rr
    Q.cell(rr, 1, lab)
    for j, v in enumerate(vals):
        if v is None: continue
        c = Q.cell(rr, 2 + j, v); c.number_format = "#,##0.000"
        if not (isinstance(v, str) and v.startswith("=")): c.font = BLUE
    Q.cell(rr, 8, tag); Q.cell(rr, 9, PRQ3 + "；" + K10); Q.cell(rr, 10, note or "—")
    Q.cell(rr, 11, INC_CO if tag == "Interested-party" else "")
rr = 4 + len(QD)
Q.cell(rr, 1, "檢查：Q1＋Q2＋Q3 營收 − 9M 營收（應為 0）"); Q.cell(rr, 5, "=B4+C4+D4-E4").number_format = "0.000"; Q.cell(rr, 8, "—"); QR["chk"] = rr
Q.cell(rr + 1, 1, "對照：9M 調整後 EBITDA − 營業利益 − 股票薪酬 − 折舊攤銷（其他調整）"); Q.cell(rr + 1, 5, "=E6-E5-E11-E10").number_format = "0.000"; Q.cell(rr + 1, 8, "—")
Q.cell(rr + 2, 1, "每 MW 年營收（Q3：雲端營收扣終止費 × 4 ÷ 平均在役 MW；在役 MW 為假設）"); Q.cell(rr + 2, 4, "=(D7-nonrec_q3)*4/((mw_q2+mw_q3)/2)").number_format = "0.00"; Q.cell(rr + 2, 8, "Assumed"); QR["permw"] = rr + 2
Q.cell(rr + 2, 10, "公司實況驗證（CoreWeave 樣板 CLAUDE.md 第 15 條）：分母在役 MW 為假設（公司未揭露），只檢驗 MW 假設與 k 的合併結果，見 Recon")
Q.column_dimensions["A"].width = 58
for c in "BCDEFG": Q.column_dimensions[c].width = 11
Q.column_dimensions["H"].width = 16; Q.column_dimensions["I"].width = 60; Q.column_dimensions["J"].width = 50; Q.column_dimensions["K"].width = 30
Q.freeze_panes = "B4"

# =====================================================================================
# Thesis：命題模組（AI 雲端：在役 MW × 每 MW 年收入）
# =====================================================================================
TH = wb.create_sheet("Thesis")
TH["A1"] = "Thesis — 命題模組：WhiteFiber 以多年期專用運算合約與客戶預付支撐的雲端擴張，加上 NC-1 託管長約，能否以高於資金成本的報酬轉成營收與現金，同時不讓資金缺口失控。驅動雲端服務分部；託管與其他（NC-1、MTL、代管服務）列在本頁託管列（三情境相同）"
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
                                          "J": f"=I{rc}+add{s}_post", "K": f"=J{rc}+add{s}_post"}, "#,##0", "情境路徑（In 頁 B 節）；2026-09-30 與 2026 年底為假設（公司未揭露季末雲端在役 MW）；口徑：IT")
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
        {c: f"=IF(p_new_ovr>0,(MIN({c}{rv},mw_q3)*{c}{rpm}+MAX(0,{c}{rv}-mw_q3)*p_new_ovr)/{c}{rv},{c}{rpm})" for c in "DEFGHIJK"}, "0.00", "基準 p_new_ovr＝0（公司宣稱只進 Sens）")
    rr_ = nr("rev")
    put(TH, rr_, "雲端營收（命題）", "US$M", {"C": "=rev9m_ai", "D": f"=D{rb}*D{rpe}*D{R['L']}", "E": f"=C{rr_}+D{rr_}", **{c: f"={c}{rb}*{c}{rpe}" for c in FC}}, "#,##0.000", "9M＝雲端服務分部實績（10-Q；含終止費 15.7）；其後＝計費 MW × 每 MW 年收入")
    put(TH, nr("revchk"), "  對照：9M 模型估計雲端營收（計費 MW × 每 MW 年收入 × 0.75；只作公司實況驗證）", "US$M", {"C": f"=C{rb}*C{rpm}*0.75"}, "#,##0.000", "與實績差距見 Recon C 節（CLAUDE.md 第 15 條）")
    rco = nr("colo")
    put(TH, rco, "託管與其他營收（NC-1、MTL、代管、其他；三情境相同）", "US$M", {"C": "=rev9m-rev9m_ai", "D": "=colo_q4", "E": f"=C{rco}+D{rco}", "F": "=colo_core+mgd_27", **{c: "=colo_core+mgd_28" for c in "GHIJK"}}, "#,##0.000", "In 頁 H 節；9M＝合併營收 − 雲端營收")
    rcoc = nr("cxo")
    put(TH, rcoc, "託管資本支出（NC-2／NC-3 收購、MTL-2、維護；三情境相同）", "US$M", {"D": "=cxo_q4", "F": "=cxo_27", **{c: "=cxo_28" for c in "GHIJK"}}, "#,##0.000", "列入自建機房資本支出列，依機房年限折舊")
    rox = nr("oxm")
    put(TH, rox, "每 MW 年營運費用：在役世代加權 IF_OpexGW", "US$M／MW·年", {**{c: "=" + mix(c, "ox") for c in "D"}, **{c: "=" + mix(c, "ox") for c in FC}}, "0.000")
    rop = nr("opex")
    put(TH, rop, "營運費用（電費、維護、人員、稅與保險；不含折舊、租金）", "US$M", {"D": f"=D{rv}*D{rox}*0.25", **{c: f"={c}{rv}*{c}{rox}" for c in FC}}, "#,##0.000", "平均在役 MW × 每 MW 營運費用")
    rrt = nr("rent")
    put(TH, rrt, "機房租金（租賃比例 × 每 MW 年租金）", "US$M", {"D": f"=D{rv}*(1-own_share)*rent_mw*0.25", **{c: f"={c}{rv}*(1-own_share)*rent_mw" for c in FC}}, "#,##0.000")
    rsg = nr("sga")
    put(TH, rsg, "營業費用（一般及管理，扣 SBC；雲端＋託管營收 × 費用率）", "US$M", {"D": f"=(D{rr_}+D{rco})*sga_r", **{c: f"=({c}{rr_}+{c}{rco})*sga_r" for c in FC}}, "#,##0.000")
    re_ = nr("ebitda")
    put(TH, re_, "調整後 EBITDA（扣 SBC 前；含託管 EBITDA＝託管營收 × 託管 EBITDA 率）", "US$M", {"C": "=aeb9m", "D": f"=D{rr_}-D{rop}-D{rrt}-D{rsg}+D{rco}*colo_m", "E": f"=C{re_}+D{re_}", **{c: f"={c}{rr_}-{c}{rop}-{c}{rrt}-{c}{rsg}+{c}{rco}*colo_m" for c in FC}}, "#,##0.000", "9M＝全公司調整後 EBITDA 實績（非 GAAP，Interested-party；含託管）")
    put(TH, nr("ebm"), "  調整後 EBITDA 率（÷ 雲端＋託管營收）", "%", {c: f"=IF({c}{rr_}+{c}{rco}=0,0,{c}{re_}/({c}{rr_}+{c}{rco}))" for c in "CDEFGHIJK"}, "0.0%")
    rcg = nr("cxg")
    put(TH, rcg, "成長型 IT 資本支出（新增 MW × IF_CapexIT，λ 提前付款）", "US$M",
        {"D": f"=(D{ra}*(1-lam)*{CXN['D']}+F{ra}*lam*{CXN['F']}*0.25)", **{c: f"=({c}{ra}*(1-lam)*{CXN[c]}+{NXT[c]}{ra}*lam*{CXN[NXT[c]]})" for c in FC}}, "#,##0.000",
        "Q4：Q4 新增 ×（1 − λ）＋ 2027 年新增 × λ × 0.25；FY2032 的次年新增假設與當年相同")
    rcf = nr("cxf")
    put(TH, rcf, "機房資本支出（新增 MW × 自建比例 × L1_FacCapexMW＋託管資本支出）", "US$M", {c: f"={c}{ra}*own_share*fac_cx+{c}{rcoc}" for c in "D" + "".join(FC)}, "#,##0.000")
    rcr = nr("cxr")
    put(TH, rcr, "汰換資本支出（汰換 MW × VR200 IF_CapexIT）", "US$M", {"D": "=0", "F": "=0", "G": "=0", **{c: f"={c}{rt}*cx_vr200" for c in "HIJK"}}, "#,##0.000")
    rcx = nr("cx")
    put(TH, rcx, "資本支出合計（公司口徑；Q4 起模型）", "US$M", {"C": "=cx9m_def", "D": f"=IF(q4cx_ovr>0,q4cx_ovr,D{rcg}+D{rcf}+D{rcr})", "E": f"=C{rcx}+D{rcx}", **{c: f"={c}{rcg}+{c}{rcf}+{c}{rcr}" for c in FC}}, "#,##0.000",
        "9M＝現金購置與設備訂金實績（10-Q）；公司未給 2026 資本支出指引")
    # 折舊
    retdep = {"F": "0", "G": "0", "H": "ret29*cx_h100", "I": "(ret29+ret30)*cx_h100",
              "J": "((ret29+ret30+ret31*v25_h100)*cx_h100+ret31*v25_gb200*cx_gb200+ret31*(1-v25_h100-v25_gb200)*cx_gb300)",
              "K": "((ret29+ret30+ret31*v25_h100)*cx_h100+ret31*v25_gb200*cx_gb200+ret31*(1-v25_h100-v25_gb200)*cx_gb300+ret32*(v26_gb300*cx_gb300+(1-v26_gb300)*cx_gb200))"}
    rdp = nr("dep")
    dv = {"C": "=da9m", "D": f"=da_q3+cip_q2/cip_life*0.25+(D{rcx}-D{rcf})/life_ai*0.25/2+D{rcf}/fac_life*0.25/2", "E": f"=C{rdp}+D{rdp}"}
    for i, c in enumerate(FC):
        prior = ["D"] + FC[:i]
        dv[c] = (f"=MAX(0,da_q3*4-{retdep[c]}/life_ai)+cip_q2/cip_life+(" + "+".join(f"{x}{rcx}-{x}{rcf}" for x in prior) + ")/life_ai+("
                 + "+".join(f"{x}{rcf}" for x in prior) + f")/fac_life+({c}{rcx}-{c}{rcf})/life_ai/2+{c}{rcf}/fac_life/2")
    put(TH, rdp, "折舊攤銷", "US$M", dv, "#,##0.000", "Q3 折舊攤銷年化 − 已汰換批次＋在建工程 ÷ 機房年限＋新資本支出 ÷ 年限（當年半年）")
    rsb = nr("sbc")
    put(TH, rsb, "股票薪酬", "US$M", {"C": "=sbc9m", "D": f"=D{rr_}*sbc_r", "E": f"=C{rsb}+D{rsb}", **{c: f"={c}{rr_}*sbc_r" for c in FC}}, "#,##0.000")
    rei = nr("ebit")
    put(TH, rei, "營業利益（GAAP）", "US$M", {"C": "=oi9m", "D": f"=D{re_}-D{rsb}-D{rdp}", "E": f"=C{rei}+D{rei}", **{c: f"={c}{re_}-{c}{rsb}-{c}{rdp}" for c in FC}}, "#,##0.000")
    put(TH, nr("om"), "  營業利益率（÷ 雲端＋託管營收）", "%", {c: f"=IF({c}{rr_}+{c}{rco}=0,0,{c}{rei}/({c}{rr_}+{c}{rco}))" for c in "CDEFGHIJK"}, "0.0%")
    rcap = nr("capb")
    put(TH, rcap, "累計投入資本（PP&E 淨額＋其後資本支出 − 折舊，期末）", "US$M", {"C": "=ppe_q3", "D": f"=C{rcap}+D{rcx}-D{rdp}", "E": f"=D{rcap}", **{c: f"={PREV[c]}{rcap}+{c}{rcx}-{c}{rdp}" for c in FC}}, "#,##0.0", "起點＝2026-09-30 PP&E 淨額（6-K）")
    put(TH, nr("roic"), "  稅後 ROIC（營業利益 ×（1 − 稅率）÷ 期初投入資本）", "%", {c: f"=IF({PREV[c]}{rcap}<=0,0,{c}{rei}*(1-tax_lt)/{PREV[c]}{rcap})" for c in FC}, "0.0%", "對照 Val 頁 WACC")
    TR[s] = R
    r += 1
RPO_ = TH.max_row + 2
TH.cell(RPO_, 1, "命題答案摘要（FY2029；核心問題）").font = BOLD
hdr(TH, RPO_ + 1, ["項目", "單位", "情境 A", "情境 B", "情境 C", "說明"])
SUMQ = [("在役 MW（2029 年底）", "MW", lambda s: f"=H{TR[s]['conn']}", "#,##0", ""),
        ("AI 雲端營收 FY2029", "US$M", lambda s: f"=H{TR[s]['rev']}", "#,##0.0", ""),
        ("調整後 EBITDA 率 FY2029", "%", lambda s: f"=H{TR[s]['ebm']}", "0.0%", ""),
        ("營業利益率 FY2029（GAAP）", "%", lambda s: f"=H{TR[s]['om']}", "0.0%", ""),
        ("稅後 ROIC FY2029", "%", lambda s: f"=H{TR[s]['roic']}", "0.0%", "對照 WACC（Val 頁）"),
        ("資本支出累計 FY2027–FY2029", "US$M", lambda s: f"=F{TR[s]['cx']}+G{TR[s]['cx']}+H{TR[s]['cx']}", "#,##0.0", ""),
        ("收支關卡：FY2026–FY2030 自由現金流扣股利", "US$M", lambda s: f"=Fund!{'CDE'[SC.index(s)]}5", "#,##0.0", "Fund 頁")]
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
SG["A1"] = "Seg — 營收分部層：WhiteFiber 官方分部為雲端服務與託管服務（另有其他營收：設備銷售型租賃利息；10-Q 分部附註）。命題模組（雲端在役 MW × 每 MW 年收入）驅動雲端服務；託管與其他（NC-1 Nscale 長約、MTL-1／MTL-3、代管服務、其他）取 Thesis 託管列（In 頁 H 節，三情境相同）。FY2025 實績、FY2026＝9M 實績＋Q4 模型、延伸到 FY2032"
SG["A2"] = "核對列：命題子列＝雲端分部、分部合計＝合併營收"
YH = ["項目", "單位"] + [f"FY{y}" + ("A" if y == 2025 else "") for y in YRS] + ["說明"]
hdr(SG, 3, YH)
r = 4
SR = {}
for s in SC:
    SG.cell(r, 1, SCN[s]).font = BOLD; r += 1
    R = {}
    for k, lab, f in (("ai", "雲端：命題模組（在役 MW × 每 MW 年收入）", lambda y: f"=Thesis!{TCOL[y]}{TR[s]['rev']}"),
                      ("seg", "分部：雲端服務", None), ("oth", "分部：託管與其他（NC-1、MTL、代管、其他）", None), ("tot", "合併營收", None)):
        R[k] = r; SG.cell(r, 1, lab); SG.cell(r, 2, "US$M"); r += 1
    SG[f"C{R['seg']}"] = "=Q!G7"; SG[f"C{R['oth']}"] = f"=Q!G4-C{R['seg']}"; SG[f"C{R['tot']}"] = f"=C{R['seg']}+C{R['oth']}"
    for y in range(2026, 2033):
        c = yc(y)
        SG[f"{c}{R['ai']}"] = f"=Thesis!{TCOL[y]}{TR[s]['rev']}"
        SG[f"{c}{R['seg']}"] = f"={c}{R['ai']}"
        SG[f"{c}{R['oth']}"] = f"=Thesis!{TCOL[y]}{TR[s]['colo']}"
        SG[f"{c}{R['tot']}"] = f"={c}{R['seg']}+{c}{R['oth']}"
    R["g"] = r; SG.cell(r, 1, "  合併營收年增"); SG.cell(r, 2, "%")
    for y in range(2026, 2033): SG[f"{yc(y)}{r}"] = f"={yc(y)}{R['tot']}/{yc(y - 1)}{R['tot']}-1"; SG[f"{yc(y)}{r}"].number_format = "0.0%"
    r += 1
    R["chk1"] = r; SG.cell(r, 1, "檢查：命題子列 − 分部（應為 0）")
    for y in range(2026, 2033): SG[f"{yc(y)}{r}"] = f"={yc(y)}{R['ai']}-{yc(y)}{R['seg']}"
    r += 1
    R["chk2"] = r; SG.cell(r, 1, "檢查：分部合計 − 合併營收（應為 0）")
    for y in YRS: SG[f"{yc(y)}{r}"] = f"={yc(y)}{R['seg']}+{yc(y)}{R['oth']}-{yc(y)}{R['tot']}"
    r += 1
    for k in ("ai", "seg", "oth", "tot", "chk1", "chk2"):
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
PL["A1"] = "PL — 三情境損益（百萬美元；EPS 美元）。營收（Seg：雲端＋託管與其他）→ 營運費用、租金、營業費用、託管 EBITDA（Thesis）→ 調整後 EBITDA → 減股票薪酬、折舊＝GAAP 營業利益 → 利息（CF）→ 稅（NOL 80% 上限）→ 淨利 → EPS"
PL["A2"] = "FY2025 為實績（10-K）；FY2026＝9M 實績（Q3 10-Q）＋Q4 模型，上方費用明細只列 FY2027 起（9M 營業成本、其他收入、所得稅見本頁末標準損益表）。所得稅：長期現金稅率 × MAX(0, 稅前 − NOL 使用)；上方 FY2026 的稅前與所得稅列只含模型項，淨利＝9M 實際淨損（含誘導轉換費用 80.8）＋Q4 模型"
hdr(PL, 3, YH)
r = 4
PL.cell(r, 1, "分部核對（三情境合計絕對差；每年度應為 0）").font = BOLD; r += 1
CHK_SEG = r; PL.cell(r, 1, "檢查：分部合計 − 合併營收"); r += 1
CHK_SUB = r; PL.cell(r, 1, "檢查：命題子列 − 分部"); r += 2
PR_ = {}
def prow(R, k, lab, vals, fmt="#,##0.000", note=None, unit="US$M"):
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
    prow(R, "ai", "雲端：命題模組（在役 MW × 每 MW 年收入）", {yc(y): f"=Seg!{yc(y)}{G['ai']}" for y in range(2026, 2033)})
    prow(R, "seg", "分部：雲端服務", {yc(y): f"=Seg!{yc(y)}{G['seg']}" for y in YRS})
    prow(R, "oth", "分部：託管與其他（NC-1、MTL、代管、其他）", {yc(y): f"=Seg!{yc(y)}{G['oth']}" for y in YRS})
    prow(R, "rev", "合併營收", {yc(y): f"=Seg!{yc(y)}{G['tot']}" for y in YRS})
    prow(R, "revg", "  合併營收年增", {yc(y): f"={yc(y)}{R['rev']}/{yc(y - 1)}{R['rev']}-1" for y in range(2026, 2033)}, "0.0%", unit="%")
    prow(R, "opex", "營運費用（Tokenomics；不含折舊、租金）", {yc(y): f"=Thesis!{TCOL[y]}{T['opex']}" for y in FYM})
    prow(R, "rent", "機房租金", {yc(y): f"=Thesis!{TCOL[y]}{T['rent']}" for y in FYM})
    prow(R, "sga", "營業費用（一般及管理，扣 SBC）", {yc(y): f"=Thesis!{TCOL[y]}{T['sga']}" for y in FYM})
    prow(R, "aeb", "調整後 EBITDA", {yc(y): f"=Thesis!{TCOL[y]}{T['ebitda']}" for y in range(2026, 2033)})
    prow(R, "aebm", "  調整後 EBITDA 率", {yc(y): f"={yc(y)}{R['aeb']}/{yc(y)}{R['rev']}" for y in range(2026, 2033)}, "0.0%", unit="%")
    prow(R, "sbc", "股票薪酬", {yc(y): f"=Thesis!{TCOL[y]}{T['sbc']}" for y in range(2026, 2033)})
    prow(R, "da", "折舊攤銷", {yc(y): f"=Thesis!{TCOL[y]}{T['dep']}" for y in range(2026, 2033)})
    prow(R, "oi", "營業利益（GAAP）", {"C": "=Q!G5", **{yc(y): f"=Thesis!{TCOL[y]}{T['ebit']}" for y in range(2026, 2033)}})
    prow(R, "oim", "  營業利益率", {yc(y): f"={yc(y)}{R['oi']}/{yc(y)}{R['rev']}" for y in range(2026, 2033)}, "0.0%", unit="%")
    prow(R, "aoi", "調整後營業利益（＝營業利益＋股票薪酬；公司口徑近似）", {"D": f"=aoi9m+Thesis!D{T['ebit']}+Thesis!D{T['sbc']}", **{yc(y): f"={yc(y)}{R['oi']}+{yc(y)}{R['sbc']}" for y in FYM}}, note="FY2026＝9M（營業利益＋股票薪酬）＋Q4 模型")
    prow(R, "int", "利息費用淨額", {}, note="CF 頁：既有借款 × r_old＋瀑布新債 × kd＋高息債 × r_junk（期初餘額）")
    prow(R, "pti", "稅前利益", {}, note="其他收益：預測期設 0")
    prow(R, "nolb", "NOL 期初餘額", {}, note="FY2026 期初＝2025 年底 nol0")
    prow(R, "nolu", "NOL 使用", {})
    prow(R, "tax", "所得稅（長期現金稅率 × 抵減後應稅所得）", {})
    prow(R, "nole", "NOL 期末餘額", {})
    prow(R, "ni", "淨利", {"C": "=Q!G8"})
    prow(R, "sh", "稀釋股數（平均）", {}, unit="百萬股")
    prow(R, "eps", "EPS（GAAP 稀釋）", {}, "0.00", unit="US$")
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
CF["A1"] = "CF — 現金流與融資瀑布（收支模組）。營運現金＝淨利＋折舊＋股票薪酬＋客戶預付流入（資本支出 × 覆蓋比）− 預付認列（非現金營收）− 營運資金增加；自由現金流的缺口依序由現金（保留最低現金）→ 投資級債（借款 ÷ 調整後 EBITDA ≤ lev_max）→ 新股（每年 ≤ 市值 eq_cap）→ 高息債 補足"
CF["A2"] = "FY2026＝9M 實績＋Q4 模型；瀑布只計 Q4（起點為 2026-09-30 現金與約當現金，受限現金不計；期後 DDTL 動用 70、NC-2／NC-3 付現 60、NC-1 信用協議 300 未動用，由瀑布自行決定融資）。利息以期初餘額計（不循環）。可轉債到期假設以原利率展延。無股利、無回購"
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
    def cw(k, lab, vals, fmt="#,##0.000", note=None, unit="US$M"):
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
    cw("pprec", "預付認列（非現金營收，自營運現金扣除）", pv, note="2026-09-30 遞延收入自 Q4 起、各年新流入自次年起，依 prepay_yrs 年直線")
    cw("nwc", "營運資金增加（預付除外）", {"D": f"=nwc_r*(Thesis!D{T['rev']}-rev_q3_ai)", **{yc(y): f"=nwc_r*(PL!{yc(y)}{P['rev']}-PL!{yc(y - 1)}{P['rev']})" for y in FYM}}, note="FY2026 只計 Q4 對 Q3 的增量")
    cw("cfo", "營運現金流", {"D": f"=cfo9m+(Thesis!D{T['ebit']}-(PL!D{P['int']}-int9m))+Thesis!D{T['dep']}+Thesis!D{T['sbc']}+D{F_('ppin')}-D{F_('pprec')}-D{F_('nwc')}",
                             **{yc(y): f"={yc(y)}{F_('ni')}+{yc(y)}{F_('da')}+{yc(y)}{F_('sbc')}+{yc(y)}{F_('ppin')}-{yc(y)}{F_('pprec')}-{yc(y)}{F_('nwc')}" for y in FYM}},
       note="FY2026＝9M 實績＋Q4（Q4 無稅）")
    cw("cx", "現金資本支出", {"D": f"=cx9m_cash+Thesis!D{T['cx']}", **{yc(y): f"=Thesis!{TCOL[y]}{T['cx']}" for y in FYM}}, note="FY2026＝9M 現金購置＋Q4 模型（公司未另定應計口徑）")
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
    cw("sh0", "期初普通股", {"D": "=sh_out", **{yc(y): f"={yc(y - 1)}{F_('sh1')}" for y in FYM}}, "0.0000", unit="百萬股")
    cw("neweq", "新發行股數", {c: f"={c}{F_('eq')}/(px*(1-eq_disc))" for c in WF}, "0.0000", unit="百萬股")
    cw("sh1", "期末普通股", {c: f"={c}{F_('sh0')}+{c}{F_('neweq')}" for c in WF}, "0.0000", unit="百萬股")
    cw("shd", "稀釋股數（平均＋RSU／選擇權淨股數）", {"D": f"=(sh_out*0.75+(D{F_('sh0')}+D{F_('sh1')})/2*0.25)+rsu_net", **{yc(y): f"=({yc(y)}{F_('sh0')}+{yc(y)}{F_('sh1')})/2+rsu_net" for y in FYM}}, "0.0000", unit="百萬股")
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
    PL.cell(P["ni"], 11, "FY2026＝9M 實績淨損＋Q4 模型；其後＝稅前 − 所得稅")
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
FSPEC = [("gate", "收支關卡：FY2026–FY2030 自由現金流合計", "US$M", lambda C_: f"=SUM(CF!D{C_['fcf']}:H{C_['fcf']})", "#,##0.0", "負值＝需外部融資；地圖收支關卡讀這一列"),
         ("cum", "Q4 2026–FY2030 外部融資合計（投資級債＋新股＋高息債）", "US$M", lambda C_: f"=SUM(CF!D{C_['need']}:H{C_['need']})", "#,##0.0", ""),
         ("ig", "  其中：新增投資級債", "US$M", lambda C_: f"=SUM(CF!D{C_['ig']}:H{C_['ig']})", "#,##0.0", ""),
         ("eq", "  其中：新股", "US$M", lambda C_: f"=SUM(CF!D{C_['eq']}:H{C_['eq']})", "#,##0.0", ""),
         ("hy", "  其中：高息債", "US$M", lambda C_: f"=SUM(CF!D{C_['hy']}:H{C_['hy']})", "#,##0.0", ""),
         ("maxb", "期間最大單年新增舉債（投資級債＋高息債）", "US$M", lambda C_: "=MAX(" + ",".join(f"CF!{c}{C_['ig']}+CF!{c}{C_['hy']}" for c in "DEFGH") + ")", "#,##0.0", "原則 55"),
         ("minc", "期間最低年末現金", "US$M", lambda C_: f"=MIN(CF!D{C_['cash1']}:H{C_['cash1']})", "#,##0.0", ""),
         ("ndmax", "淨負債 ÷ 調整後 EBITDA 最差年（FY2026–FY2030；N/M 年不計）", "x", lambda C_: f"=MAX(CF!D{C_['nde']}:H{C_['nde']})", "0.00", "門檻 4.5（WARN，不截斷）"),
         ("icmin", "利息保障倍數最差年（FY2026–FY2030）", "x", lambda C_: f"=MIN(CF!D{C_['icov']}:H{C_['icov']})", "0.00", "門檻 3（WARN）"),
         ("warn", "原則 9 兩指標是否 WARN（FY2026–FY2030）", "—", lambda C_: f'=IF(COUNTIF(CF!D{C_["ndw"]}:H{C_["ndw"]},"WARN*")+COUNTIF(CF!D{C_["icw"]}:H{C_["icw"]},"WARN")>0,"WARN","OK")', None, ""),
         ("eqmax", "新股 ÷ 期初市值：最高年", "%", lambda C_: f"=MAX(CF!D{C_['eqm']}:H{C_['eqm']})", "0.0%", "原則 62（只呈現）"),
         ("dil", "累計稀釋（2030 年底股數 ÷ 2026-09-30 − 1）", "%", lambda C_: f"=CF!H{C_['dil']}", "0.0%", "未來新股不進 DCF 每股（X35）"),
         ("chk", "檢查：Q4 2026–FY2030 外部融資 − 三項合計（應為 0）", "US$M", lambda C_: None, "0.000", "")]
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
    FR[k] = r; FU.cell(r, 1, lab); FU.cell(r, 2, "x" if k in ("nde", "icov") else ("%" if k in ("eqm", "dil") else "US$M"))
    for j, c in enumerate(WF):
        x = FU.cell(r, 4 + j, f"=CF!{c}{CB[k]}"); x.number_format = "0.00" if k in ("nde", "icov") else ("0.0%" if k in ("eqm", "dil") else "#,##0.00")
    r += 1
FU.column_dimensions["A"].width = 60; FU.column_dimensions["F"].width = 40
for c in "CDEGHIJ": FU.column_dimensions[c].width = 11

# =====================================================================================
# RPO：backlog 轉營收路徑
# =====================================================================================
RP = wb.create_sheet("RPO")
RP["A1"] = "RPO — 合約排程對模型營收（同口徑檢查，不作校準）。雲端 RPO 716.3（2026-09-30，10-Q 年度分布：2026 年餘 12.1、2027 230.5、2028 202.9、2029 190.3、2030 45.5、其後 35.0；Verified）；託管 NC-1 等以租賃認列（出租方租賃收款 421.5，不在 RPO 內）"
RP["A2"] = "覆蓋率＝雲端 RPO 排程 ÷ 模型雲端營收（情境 B）；>1＝合約已排程金額高於模型營收，<1＝模型營收需靠新簽或續約。2031–2032 欄取「其後」平均分攤"
hdr(RP, 4, ["項目", "單位", "", "FY2026（Q4）", "FY2027", "FY2028", "FY2029", "FY2030", "FY2031", "FY2032"])
RPB = {}
r = 5
RP.cell(r, 1, "雲端 RPO 排程（10-Q 年度分布）"); RP.cell(r, 2, "US$M")
sched = {"D": 12.107, "E": 230.545, "F": 202.867, "G": 190.280, "H": 45.469, "I": "=35.017/2", "J": "=35.017/2"}
for c, v in sched.items():
    RP[f"{c}{r}"] = v; RP[f"{c}{r}"].number_format = "#,##0.0"
    if not isinstance(v, str): RP[f"{c}{r}"].font = BLUE
RPB["sched"] = r; r += 1
for s in SC:
    RP.cell(r, 1, f"模型雲端營收：{SCN[s]}"); RP.cell(r, 2, "US$M")
    RP[f"D{r}"] = f"=Thesis!D{TR[s]['rev']}"
    for j, y in enumerate(FYM): RP[f"{CL(5 + j)}{r}"] = f"=Thesis!{TCOL[y]}{TR[s]['rev']}"
    for c in "DEFGHIJ": RP[f"{c}{r}"].number_format = "#,##0.0"
    RPB["rev" + s] = r; r += 1
RP.cell(r, 1, "覆蓋率：RPO 排程 ÷ 模型營收（情境 B）"); RP.cell(r, 2, "x")
for c in "DEFGHIJ": RP[f"{c}{r}"] = f"={c}{RPB['sched']}/{c}{RPB['revB']}"; RP[f"{c}{r}"].number_format = "0.00"
RPB["cov"] = r; r += 1
RP.cell(r, 1, "模型營收中未被 RPO 覆蓋的部分（情境 B；需新簽或隨需）"); RP.cell(r, 2, "US$M")
for c in "DEFGHIJ": RP[f"{c}{r}"] = f"=MAX(0,{c}{RPB['revB']}-{c}{RPB['sched']})"; RP[f"{c}{r}"].number_format = "#,##0.0"
RPB["gap"] = r; r += 1
RP.cell(r, 1, "對照：2026Q4–2032 RPO 排程合計 − 雲端 RPO（應為 0）"); RP.cell(r, 2, "US$M")
RP[f"D{r}"] = f"=SUM(D{RPB['sched']}:J{RPB['sched']})-rpo"; RP[f"D{r}"].number_format = "#,##0.0"; RPB["chk"] = r
RP.column_dimensions["A"].width = 62
for c in "DEFGHIJ": RP.column_dimensions[c].width = 11

# =====================================================================================
# Sites：具名站點與容量上限
# =====================================================================================
SI = wb.create_sheet("Sites")
SI["A1"] = "Sites — 具名站點（10-Q、法說、簡報；Interested-party）與容量對照；雲端站點為第三方機房（IT MW），託管站點為自有（IT MW；毛 MW 另註）"
SI["A2"] = "已簽專用運算 GPU 約 20,000 顆（含代管 2,048 顆）以 Tokenomics 每 GPU 功耗換算約 37 MW-IT（In ctr_mw）；部分雲端站點 MW 未揭露（找不到）"
hdr(SI, 4, ["站點", "營運／產權", "規劃 MW（IT）", "已通電 MW（IT）", "時程", "把握"])
SITES = [("冰島 Blönduós（雲端）", "第三方機房", 5.5, 5.5, "營運中；576 顆 B300 2026-12", "高"), ("安大略（Baseten 1,392 顆 B300）", "第三方機房", 2.858, 0, "2026-11 起", "中"),
         ("雪梨（1,200 顆 B300）", "第三方機房", 2.5, 0, "2027 第一季", "中"), ("巴黎（Tier-3+）", "第三方機房", 2.7, 0, "2027-01／02", "中"),
         ("MTL-2（Prime Intellect 576 顆 VR200）", "自有（雲端自用）", 3, 0, "2027 年初完工、第二季服務", "中"),
         ("NC-1 Madison, North Carolina（Nscale 10 年）", "自有（託管）", 40, 40, "第一期 20 MW 9-23、第二期 20 MW 11-09 計費", "高"),
         ("MTL-1（14 家客戶）", "承租（託管）", 3, 3, "營運中", "高"), ("MTL-3（Cerebras 5 年）", "自有（託管）", 5, 5, "營運中", "高"),
         ("NC-2／NC-3 Yadkinville（收購完成）", "自有（託管，未簽約）", 44, 0, "初始 60 毛 MW、2027 年底 RFS（意向書）", "低")]
ST = {"s0": 5}
for i, row in enumerate(SITES):
    for j, v in enumerate(row): SI.cell(5 + i, 1 + j, v)
r = 5 + len(SITES)
SI.cell(r, 1, "具名站點合計"); SI.cell(r, 3, f"=SUM(C5:C{r - 1})"); SI.cell(r, 4, f"=SUM(D5:D{r - 1})"); ST["sum"] = r; r += 1
SI.cell(r, 1, "已簽約雲端 GPU 換算 MW"); SI.cell(r, 3, "=ctr_mw"); ST["ctr"] = r; r += 1
SI.cell(r, 1, "對照：已簽約雲端 MW − 具名雲端站點（前五列）"); SI.cell(r, 3, f"=C{ST['ctr']}-SUM(C5:C9)"); ST["res"] = r; r += 1
SI.cell(r, 1, "情境 A／B／C：2030 年底雲端在役 MW ÷ 已簽約雲端 MW");
for j, s in enumerate(SC): SI.cell(r, 3 + j, f"=mw{s}_30/ctr_mw").number_format = "0.00"
ST["capc"] = r
SI.column_dimensions["A"].width = 50; SI.column_dimensions["B"].width = 34

# =====================================================================================
# Debt：逐檔與到期表（Q2 10-Q；待 Q3 10-Q）
# =====================================================================================
DB = wb.create_sheet("Debt")
DB["A1"] = "Debt — 債務逐檔（2026-09-30，10-Q；本金，非帳面值）；期後：DDTL 追加 70（2026-10-27）、NC-1 信用協議 300（2026-11-09 簽約、未動用）不計入評價日淨負債，由瀑布另計新債"
DB["A2"] = "基準假設：到期借款以原利率展延；可轉債視為負債（2031 轉換價約 25.91、2032 約 33.84，皆遠高於現價）。逐檔本金合計與 Q3 帳面值 402.6 的差額為發行成本與折價"
hdr(DB, 4, ["債務", "類別", "到期", "利率", "本金（US$M）", "說明"])
INST = [("4.50% 可轉債 2031", "可轉債", "2031-02", 0.045, 31.85, "2026-08 誘導轉換後剩餘；2029-02 起可賣回"), ("5.00% 可轉債 2032", "可轉債", "2032-09", 0.05, 310.0, "2026-08 發行；2030-09 起可賣回"),
        ("Bit Digital DDTL（關係人）", "擔保", "2027-01", 0.095, 30.0, "額度 100（可增至 150）；1.1× MOIC；期後再動用 70"), ("Landsbankinn 冰島擔保貸款", "擔保", "2028", 0.0815, 17.8, "SOFR＋4.25%；季攤還"),
        ("RBC 聯貸（CAD 36.8）", "擔保", "2029", 0.0525, 26.2, "CORRA＋2.45%；15 年攤還")]
D_ = {"r0": 5}
r = 5 + len(INST)
DB.cell(r, 1, "逐檔合計（本金）"); DB.cell(r, 5, f"=SUM(E5:E{r - 1})").number_format = "#,##0.000"; D_["tot"] = r; r += 1
DB.cell(r, 1, "Q3 資產負債表借款帳面值"); DB.cell(r, 5, "=debt_q3").number_format = "#,##0.000"; r += 1
DB.cell(r, 1, "差額（發行成本與折價）"); DB.cell(r, 5, f"=E{r - 1}-E{r - 2}").number_format = "#,##0.000"; D_["diff"] = r; r += 2
DB.cell(r, 1, "到期與攤還表（本金；US$B）").font = BOLD; r += 1
D_["hdr"] = r
hdr(DB, r, ["項目", "", "", "2026 第四季", "2027", "2028", "2029", "2030", "2031 以後", "合計"]); r += 1
DB.cell(r, 1, "既有借款到期／攤還（DDTL 2027-01、冰島與 RBC 以到期年粗列、可轉債 2031 以後）")
for j, v in enumerate((0, 30.0, 17.8, 26.2, 0, 341.85)): DB.cell(r, 4 + j, v).font = BLUE
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
LS["A1"] = "Lease — 租賃。營業租賃負債 69.3（2026-09-30，流動 12.955＋非流動 56.297）；未起租加拿大託管租約 34.0（2.7 MW、60 個月）；GPU 伺服器租賃為變動租金（9M 12.2，含在營業成本）"
LS["A2"] = "租金路徑＝雲端平均在役 MW × 承租比例（雲端 100% 第三方機房）× 每 MW 年租金（Thesis）"
hdr(LS, 4, ["項目", "單位", "", "FY2026（Q4）", "FY2027", "FY2028", "FY2029", "FY2030", "FY2031", "FY2032"])
LSB = {}
r = 5
for s in SC:
    LS.cell(r, 1, f"機房租金：{SCN[s]}"); LS.cell(r, 2, "US$M")
    LS[f"D{r}"] = f"=Thesis!D{TR[s]['rent']}"
    for j, y in enumerate(FYM): LS[f"{CL(5 + j)}{r}"] = f"=Thesis!{TCOL[y]}{TR[s]['rent']}"
    LSB["rent" + s] = r; r += 1
LS.cell(r, 1, "情境 B：每 MW 年租金（租金 ÷ 平均租賃 MW）"); LS.cell(r, 2, "US$M／MW·年")
for j, y in enumerate(FYM): LS[f"{CL(5 + j)}{r}"] = f"=Thesis!{TCOL[y]}{TR['B']['rent']}/(Thesis!{TCOL[y]}{TR['B']['avg']}*(1-own_share))"
LSB["permw"] = r; r += 1
LS.cell(r, 1, "情境 B：FY2027–FY2031 租金合計 ÷（在帳營業租賃負債＋已簽約未起租）"); LS.cell(r, 2, "x")
LS[f"D{r}"] = f"=SUM(E{LSB['rentB']}:I{LSB['rentB']})/(opl_q3+lease_nc)"; LSB["ratio"] = r
for rr_ in range(5, r + 1):
    for c in "DEFGHIJ": LS[f"{c}{rr_}"].number_format = "#,##0.00"
LS.column_dimensions["A"].width = 62

# =====================================================================================
# Val：DCF 50%＋Comparable 50%
# =====================================================================================
V = wb.create_sheet("Val")
V["A1"] = "Val — 目標價（現價日＋12 個月）＝DCF（FCFF、WACC；DCF 標準規格）50%＋Comparable（分部加總：雲端分部 × 新雲同業前瞻 EV/EBITDA、託管分部 × 上市資料中心同業前瞻 EV/EBITDA；Andy 2026-11-15）50%；情境權重 A 25%／B 50%／C 25%（判斷預設，Andy 可改）"
V["A2"] = "評價日 2026-09-30（Q3 資產負債表）。金額百萬美元、股數百萬股。DCF：NOPAT＝營業利益 − 長期稅率 × MAX(0, 營業利益)（股票薪酬視為現金成本）＋折舊 − 資本支出 − 營運資金增加＋預付淨額；永續 NOPAT ×(1 − g ÷ RONIC) ÷ (WACC − g)，RONIC＝WACC；股數＝評價日流通股數＋RSU／選擇權淨股數，未來瀑布新股不進 DCF 每股（X35）；可轉債列負債。Comparable 錨定 FY2027 年底（含當時的淨負債與股數）"
hdr(V, 3, ["項目", "單位", "數值", "", "", "說明"])
VR = {}
def vrow(k, lab, unit, val, fmt="#,##0.00", note=None):
    global r
    VR[k] = r; V.cell(r, 1, lab); V.cell(r, 2, unit)
    if val is not None: V.cell(r, 3, val).number_format = fmt
    if note: V.cell(r, 6, note)
    r += 1
r = 4
vrow("px", "現價", "US$", "=px", note='="In 頁 px（"&TEXT(px_date,"yyyy-mm-dd")&" 收盤）；v0.4.2 為 2026-11-10 收盤 9.21；v0.2.1 為 7.80"')
vrow("ke", "股權成本 ke（CAPM：rf＋β × ERP）", "%", "=rf+beta*erp", "0.00%", '="rf "&TEXT(rf,"0.00%")&"（美債 10 年期 2026-11-10）、β "&TEXT(beta,"0.00")&"（自算產業 β "&TEXT(beta_u,"0.000")&" 依市值 D/E 重新加槓桿）、ERP "&TEXT(erp,"0%")')
vrow("kdat", "稅後債務成本", "%", "=kd*(1-tax_lt)", "0.00%", '="kd "&TEXT(kd,"0.00%")&"（邊際純債：Bit Digital DDTL 票面）；既有借款平均利率 "&TEXT(r_old,"0.0%")&"（可轉債低票息，只作對照）"')
vrow("E", "股權市值（現價 × 評價日股數）", "US$M", "=px*(sh_out+rsu_net)", "#,##0.0")
vrow("D", "負債（借款帳面值＋融資租賃）", "US$M", "=debt_q3+finl_q3", "#,##0.0")
vrow("wacc", "WACC（＋wacc_adj，敏感度用）", "%", f"=(C{r - 2}*C{r - 4}+C{r - 1}*C{r - 3})/(C{r - 2}+C{r - 1})+wacc_adj", "0.00%", "市值權重（A5）")
vrow("ronic", "永續期新投資報酬率 RONIC", "%", f"=C{r - 1}+ronic_adj", "0.00%", "RONIC＝WACC（DCF 標準規格）")
vrow("nd", "淨負債（借款含可轉債＋融資租賃 − 現金；受限現金不扣）", "US$M", "=debt_q3+finl_q3-cash_q3", "#,##0.0")
vrow("sh", "評價日股數（流通＋RSU／選擇權淨股數）", "百萬股", "=sh_out+rsu_net", "0.000")
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
    V.cell(r, 3 + j, f"=(K{T['ebitda']}-K{T['sbc']}-K{T['conn']}*(cx_vr200/life_ai+own_share*fac_cx/fac_life)-K{T['cxo']})*(1-tax_lt)".replace("K", "Thesis!K")).number_format = "#,##0.00"
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
V.cell(r, 1, "NOL 稅盾價值（NOL × 稅率，以 WACC 折現 3 年；Assumed；基準 NOL 0）")
for j in range(3): V.cell(r, 3 + j, "=nol0*tax_lt/(1+wacc)^3").number_format = "#,##0.00"
DC["nol"] = r; r += 1
V.cell(r, 1, "股權價值（企業價值 − 淨負債＋NOL 稅盾＋非上市權益證券投資）")
for j in range(3):
    cc = CL(3 + j); V.cell(r, 3 + j, f"={cc}{DC['ev']}-netdebt0+{cc}{DC['nol']}+hold").number_format = "#,##0.0"
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
V.cell(r, 1, "0.2 Comparable：雲端分部同業（新雲 NTM EV/EBITDA；同業估值日 2026-11-13＝現價日；IREN 11-11 收盤 20.84 → 11-13 收盤 18.17）").font = BOLD; r += 1
PEERH = r
hdr(V, r, ["同業", "代號", "EV（US$B）", "NTM EV/EBITDA", "估值日", "口徑與來源", "NTM EBITDA 率", "淨負債 ÷ NTM EBITDA", "資本支出 ÷ 營收", "營收成長（下一年）"]); r += 1
P0 = r
for nm, tk, ev, eb, src, mg, nd, cx_, gr in (
        ("CoreWeave", "CRWV", "pc_px*pc_sh+IF(peer_nd_mode=1,pc_ndf,pc_nd)", "pc_eb", "US GAAP；共識為 2026-09-25 MarketScreener（CoreWeave 財報前；財報後共識未取得）；淨負債 2026-09-30 申報（借款＋融資租賃 − 現金）", "=pc_eb/pc_rev", "=IF(peer_nd_mode=1,pc_ndf,pc_nd)/pc_eb", "=pc_cx/pc_rev", "=28.857/16.656-1"),
        ("Nebius", "NBIS", "pn_px*pn_sh+IF(peer_nd_mode=1,pn_ndf,pn_nd)", "pn_eb", "US GAAP；MarketScreener 共識（擷取 2026-11-13）；淨負債 2026-09-30 申報（借款含可轉債 − 現金，受限現金不扣）", "=pn_eb/pn_rev", "=IF(peer_nd_mode=1,pn_ndf,pn_nd)/pn_eb", "未取得", "=13700/6556-1"),
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
V.cell(r, 1, "  估值日 2026-11-13（＝本模型現價日；IREN 亦為 11-13 收盤 18.17，原 11-11 收盤 20.84）；NTM 自 11-11 起依月份加權（原則 49）；淨負債主口徑＝最新申報（CoreWeave、Nebius、IREN 為 2026-09-30，受限現金不扣；Oracle 含 2026-11 增資 pro forma），替代口徑見下兩列；CoreWeave 共識為其財報前。同業 4 家，少於 8 家不做迴歸（原則 7）。託管同業（DLR、EQIX 約 20 倍）不納入：本模型以單一新雲倍數評整家公司（判斷題，見 README）").font = GREY; r += 1
V.cell(r, 1, "  替代口徑：淨負債改用共識年底預估時的 NTM EV/EBITDA（CRWV／NBIS／ORCL／IREN；時點不一致：11-13 市值配年底預估淨負債，只作對照）")
for j, (pxs, nds, ebs) in enumerate((("pc_px*pc_sh", "pc_ndf", "pc_eb"), ("pn_px*pn_sh", "pn_ndf", "pn_eb"), ("po_px*po_sh", "po_ndf", "po_eb"), ("ir_px*ir_sh", "ir_ndf", "ir_eb"))):
    V.cell(r, 3 + j, f"=({pxs}+{nds})/{ebs}").number_format = "0.00"
VR["altm"] = r; r += 1
V.cell(r, 1, "  替代口徑中位數（對照；時點不一致；Sens peer_nd_mode）"); V.cell(r, 4, f"=MEDIAN(C{r - 1}:F{r - 1})").number_format = "0.00"; VR["altmed"] = r; r += 1
V.cell(r, 1, "  對照（不進中位數）：WhiteFiber 自身＝現價市值＋淨負債 ÷ 2027 年 EBITDA 代理（營收共識 × 40%；EBITDA 共識找不到）；後四欄為模型情境 B FY2027").font = GREY
V.cell(r, 3, "=px*shv+netdebt0").number_format = "#,##0.0"; V.cell(r, 4, f"=C{r}/cons_eb27").number_format = "0.00"; V.cell(r, 5, "2026-11-13")
VR["selfm"] = r; SELF_R = r
for j, f_ in enumerate((f"=PL!E{PR_['B']['aebm']}", f"=netdebt0/PL!E{PR_['B']['aeb']}", f"=Thesis!F{TR['B']['cx']}/PL!E{PR_['B']['rev']}", f"=PL!E{PR_['B']['revg']}")):
    V.cell(r, 7 + j, f_).number_format = "0.00" if j == 1 else "0.0%"
r += 1
V.cell(r, 1, "同業中位數（前瞻 EV/EBITDA；雲端分部倍數）"); V.cell(r, 4, f"=MEDIAN(D{P0}:D{P1})").number_format = "0.00"; VR["mult"] = r; dn("mult", f"Val!$D${r}"); r += 2
V.cell(r, 1, "0.2b 託管分部同業（上市資料中心營運商 NTM EV/EBITDA；Andy 2026-11-15 分部加總）").font = BOLD; r += 1
hdr(V, r, ["同業", "代號", "EV（US$M）", "NTM EV/EBITDA", "估值日", "口徑與來源"]); VR["hph"] = r; r += 1
H0 = r; VR["hp0"] = r
for nm, tk, ev, eb, dt, src in (("Digital Realty", "DLR", "dlr_ev", "dlr_eb", "2026-11-13（EV 自 10-07 依收盤價比換算）", "MarketScreener 估值頁（EV Derived；EBITDA 為賣方共識 Interested-party）；REIT、批發型資料中心為主，最接近 NC-1 的批發長約"),
                                ("Equinix", "EQIX", "eqix_ev", "eqix_eb", "2026-11-13（EV 自 10-02 依收盤價比換算）", "MarketScreener 估值頁（EV Derived；EBITDA 為賣方共識 Interested-party）；REIT、零售型共置與互連為主，最接近 MTL 多客戶共置")):
    V.cell(r, 1, nm); V.cell(r, 2, tk); V.cell(r, 3, "=" + ev).number_format = "#,##0.0"; V.cell(r, 4, f"={ev}/{eb}").number_format = "0.00"
    V.cell(r, 5, dt); V.cell(r, 6, src); r += 1
H1 = r - 1
V.cell(r, 1, "託管同業中位數（前瞻 EV/EBITDA）"); V.cell(r, 4, f"=MEDIAN(D{H0}:D{H1})").number_format = "0.00"; VR["multh0"] = r; r += 1
V.cell(r, 1, "託管分部倍數（同業中位數 ×（1 − 折價））"); V.cell(r, 4, f"=D{VR['multh0']}*(1-colo_disc)").number_format = "0.00"; VR["multh"] = r; dn("mult_colo", f"Val!$D${r}"); r += 1
V.cell(r, 1, "  同業選取：上市資料中心營運商中有 NTM EBITDA 共識者 2 家（Digital Realty、Equinix）；Applied Digital、Core Scientific、Cipher Mining（HPC 託管）無可用 NTM EBITDA 共識或為負、Galaxy 未取得（data/whitefiber_facts_20261008.json 資料缺口），未納入；同業少於 8 家不做迴歸（原則 7）。EV 原為 2026-10 初報價（DLR 10-07、EQIX 10-02），已依報價日收盤至 11-13 收盤的價格比換算（倍數 × 價格比近似，Derived；EBITDA 共識仍為 2026-10-08 擷取）；WhiteFiber 規模小、NC-1 單一主要承租人、非 REIT，折價 28%（約 15 倍）列 Sens").font = GREY; r += 2
hdr(V, r, ["項目", "單位", "情境 A", "情境 B", "情境 C", "說明"]); r += 1
CO = {}
def vr3(k, lab, unit, f, fmt="#,##0.00", note=None):
    global r
    CO[k] = r; V.cell(r, 1, lab); V.cell(r, 2, unit)
    for j, s in enumerate(SC): V.cell(r, 3 + j, f(s, CL(3 + j))).number_format = fmt
    if note: V.cell(r, 6, note)
    r += 1
vr3("eb", "FY2028 調整後 EBITDA（FY2027 年底的前瞻一年）", "US$M", lambda s, c: f"=PL!F{PR_[s]['aeb']}")
vr3("rvh", "  FY2028 託管分部營收（NC-1、MTL、代管服務、其他；Seg 第二分部）", "US$M", lambda s, c: f"=PL!F{PR_[s]['oth']}")
vr3("ebh", "  FY2028 託管分部 EBITDA（營收 ×（託管 EBITDA 率 − 營業費用率）；營業費用依營收比例分攤）", "US$M", lambda s, c: f"={c}{CO['rvh']}*(colo_m-sga_r)")
vr3("ebc", "  FY2028 雲端分部 EBITDA（調整後 EBITDA − 託管分部）", "US$M", lambda s, c: f"={c}{CO['eb']}-{c}{CO['ebh']}")
vr3("evc0", "  雲端分部企業價值（× 新雲同業中位數）", "US$M", lambda s, c: f"=IF(sotp_on=1,{c}{CO['ebc']},{c}{CO['eb']})*mult*mult_adj", "#,##0.0")
vr3("evh0", "  託管分部企業價值（× 資料中心同業倍數）", "US$M", lambda s, c: f"=IF(sotp_on=1,{c}{CO['ebh']}*mult_colo*mult_adj,0)", "#,##0.0")
vr3("ev0", "企業價值（未調整：分部加總）", "US$M", lambda s, c: f"={c}{CO['evc0']}+{c}{CO['evh0']}", "#,##0.0")
vr3("sotpchk", "檢查：雲端＋託管分部 EBITDA − 調整後 EBITDA（應為 0）", "US$M", lambda s, c: f"={c}{CO['ebc']}+{c}{CO['ebh']}-{c}{CO['eb']}", "0.000000")
vr3("ev", "企業價值（× 相對倍數係數 rel_mult）", "US$M", lambda s, c: f"={c}{CO['ev0']}*rel_mult", "#,##0.0", '="rel_mult "&TEXT(rel_mult,"0.00")&"（未調整：缺 WhiteFiber 10 年歷史倍數）"')
vr3("nd27", "淨負債（FY2027 年底：借款＋融資租賃 − 現金）", "US$M", lambda s, c: f"=CF!E{CR_[s]['debt']}+finl_q3-CF!E{CR_[s]['cash1']}", "#,##0.0")
vr3("sh27", "股數（FY2027 年底，含瀑布新股＋RSU／選擇權）", "百萬股", lambda s, c: f"=CF!E{CR_[s]['sh1']}+rsu_net", "0.000")
vr3("ps0", "Comparable 每股（未調整；FY2027 年底；含非上市權益證券投資）", "US$", lambda s, c: f"=MAX(0,({c}{CO['ev0']}-{c}{CO['nd27']}+hold)/{c}{CO['sh27']})")
vr3("ps", "Comparable 每股（FY2027 年底；含非上市權益證券投資）", "US$", lambda s, c: f"=MAX(0,({c}{CO['ev']}-{c}{CO['nd27']}+hold)/{c}{CO['sh27']})")
vr3("v120", "Comparable 情境價值（未調整；折回目標價時點）", "US$", lambda s, c: f"={c}{CO['ps0']}/(1+ke)^t27")
vr3("v12", "Comparable 情境價值（折回目標價時點）", "US$", lambda s, c: f"={c}{CO['ps']}/(1+ke)^t27")
for j in range(3): V.cell(CO["v12"], 3 + j).fill = KEY
vr3("pcc", "  分部貢獻：雲端分部企業價值每股（折回目標價時點；未扣淨負債）", "US$", lambda s, c: f"={c}{CO['evc0']}*rel_mult/{c}{CO['sh27']}/(1+ke)^t27")
vr3("pch", "  分部貢獻：託管分部企業價值每股（同上）", "US$", lambda s, c: f"={c}{CO['evh0']}*rel_mult/{c}{CO['sh27']}/(1+ke)^t27")
vr3("pcn", "  分部貢獻：淨負債、非上市投資與 0 下限（＝情境價值 − 兩分部）", "US$", lambda s, c: f"={c}{CO['v12']}-{c}{CO['pcc']}-{c}{CO['pch']}")
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
crow("gatev", "入不敷出判定用：情境 B 的 FY2026–FY2030 自由現金流合計（＝Fund 頁收支關卡）", "US$M", f"=Fund!D{FS['gate']}", "#,##0.0", "原則 63：合計 < 0＝入不敷出＝收支關卡不通過（判準與 AI 投資地圖相同；本模型無容忍帶）")
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
crow("npvAB", "每增加 1 MW 的增量 DCF 企業價值：A 對 B（原則 60）", "US$B／GW", f"=(C{DC['ev']}-D{DC['ev']})/(mwA_30-mwB_30)", "#,##0.00", "2030 年底雲端在役 MW 差；負值＝擴張毀損價值")
crow("npvBC", "每增加 1 MW 的增量 DCF 企業價值：B 對 C", "US$B／GW", f"=(D{DC['ev']}-E{DC['ev']})/(mwB_30-mwC_30)", "#,##0.00")
crow("npvps", "每增加 1 MW 的增量情境目標價：A 對 B（US$／股）", "US$", f"=(C{CON['tps']}-D{CON['tps']})/(mwA_30-mwB_30)", "#,##0.00")
crow("cons", "市場共識目標價（2026-11-10，11 家平均，只作對照）", "US$", "=cons_tp")
crow("v413", "v0.2.1 加權目標價（舊 JS 引擎；現價 7.80）", "US$", "=tp_v053")
crow("chg", f"目標價變動：{VER} − v0.2.1", "US$", f"=C{CON['tp']}-C{CON['v413']}")
crow("v040", "v0.4.0 加權目標價（全公司 × 新雲同業中位數；現價 9.21）", "US$", "=tp_v040")
crow("chg040", f"目標價變動：{VER} − v0.4.0（同業法改分部加總）", "US$", f"=C{CON['tp']}-C{CON['v040']}")
crow("tvsh", "情境 B：永續價值占企業價值", "%", f"=D{DC['tvsh']}", "0.0%")
# 橫幅與白話字串（公式組字；原則 13、31、57、60）
crow("bn_scn", "橫幅：擴張與目標價", "—", f'=IF(AND(C{CON["tps"]}<D{CON["tps"]},D{CON["tps"]}<E{CON["tps"]}),"A 情境擴張最多但目標價最低："&TEXT(C{CON["tps"]},"0.00")&"／"&TEXT(D{CON["tps"]},"0.00")&"／"&TEXT(E{CON["tps"]},"0.00")&"（A／B／C）；每增加 1 MW（A 對 B）DCF 企業價值變動 "&TEXT(C{CON["npvAB"]},"+0.0;-0.0")&" 百萬美元",IF(AND(C{CON["tps"]}>D{CON["tps"]},D{CON["tps"]}>E{CON["tps"]}),"擴張越多目標價越高："&TEXT(C{CON["tps"]},"0.00")&"／"&TEXT(D{CON["tps"]},"0.00")&"／"&TEXT(E{CON["tps"]},"0.00")&"（A／B／C）","情境目標價 "&TEXT(C{CON["tps"]},"0.00")&"／"&TEXT(D{CON["tps"]},"0.00")&"／"&TEXT(E{CON["tps"]},"0.00")&"（A／B／C；非單調）")&IF(C{CON["npvAB"]}<0,"；但 DCF 口徑每增加 1 MW（A 對 B）企業價值 "&TEXT(C{CON["npvAB"]},"+0.0;-0.0")&" 百萬美元、（B 對 C）"&TEXT(C{CON["npvBC"]},"+0.0;-0.0")&"：擴張價值為負，情境目標價的差異來自 Comparable 的 EBITDA 倍數",""))', None)
UB = f'IF(C{CON["up"]}>0.5,NA(),IF(C{CON["up"]}>0.25,0.5,IF(C{CON["up"]}>=-0.25,0.25,IF(C{CON["up"]}>=-0.5,-0.25,-0.5))))'
LB = f'IF(C{CON["up"]}>0.5,0.5,IF(C{CON["up"]}>0.25,0.25,IF(C{CON["up"]}>=-0.25,-0.25,IF(C{CON["up"]}>=-0.5,-0.5,NA()))))'
MREQ = lambda b, cond: (f'=IF({cond},"—",((px*(1+{b})-w_dcf*C{CON["dcf"]})/(1-w_dcf)*(1+ke)^t27+wA*C{CO["nd27"]}/C{CO["sh27"]}+wB*D{CO["nd27"]}/D{CO["sh27"]}+wC*E{CO["nd27"]}/E{CO["sh27"]}-hold*(wA/C{CO["sh27"]}+wB/D{CO["sh27"]}+wC/E{CO["sh27"]})-rel_mult*(wA*C{CO["evh0"]}/C{CO["sh27"]}+wB*D{CO["evh0"]}/D{CO["sh27"]}+wC*E{CO["evh0"]}/E{CO["sh27"]}))'
                  f'/((wA*IF(sotp_on=1,C{CO["ebc"]},C{CO["eb"]})/C{CO["sh27"]}+wB*IF(sotp_on=1,D{CO["ebc"]},D{CO["eb"]})/D{CO["sh27"]}+wC*IF(sotp_on=1,E{CO["ebc"]},E{CO["eb"]})/E{CO["sh27"]})*mult_adj*rel_mult))')
crow("mup", "升一級所需雲端同業倍數（託管倍數與其他不變；Comparable 每股未觸及 0 下限時）", "x", MREQ("0.5" if False else UB.replace("NA()","0.5"), f'C{CON["up"]}>0.5'), "0.00", "原則 36：反解加權目標價＝升一級邊界")
crow("mdn", "降一級所需雲端同業倍數（託管倍數不變）", "x", MREQ(LB.replace("NA()","-0.5"), f'C{CON["up"]}<-0.5'), "0.00", "反解加權目標價＝降一級邊界（現行邊界）")
crow("bn_edge", "橫幅：距級距邊界", "—", f'=IF(C{CON["dist"]}<0.05,"接近級距邊界：","")&"距最近級距邊界 "&TEXT(C{CON["dist"]}*100,"0.0")&" 個百分點；雲端同業倍數 "&TEXT(mult,"0.00")&" 倍（託管 "&TEXT(mult_colo,"0.00")&" 倍不變），升一級需 "&IF(ISNUMBER(C{CON["mup"]}),TEXT(C{CON["mup"]},"0.00")&" 倍","—（已最高級）")&"、降一級需 "&IF(ISNUMBER(C{CON["mdn"]}),TEXT(C{CON["mdn"]},"0.00")&" 倍","—")', None)
crow("bn_dcf", "橫幅：DCF 有效性", "—", f'=C{CON["dvalid"]}', None)
crow("val_lab", "估值評等標示（距邊界 <5pt 時加「邊界」）", "—", f'=C{CON["valrec"]}&IF(C{CON["dist"]}<0.05,"（邊界，±"&TEXT(C{CON["dist"]}*100,"0.0")&"pt）","")', None)
crow("rec_lab", "評等標示（收支不足時並列估值評等與收支合計）", "—", f'=IF(C{CON["rec"]}="收支不足","收支不足〔估值評等："&C{CON["val_lab"]}&"；情境 B 的 FY2026–FY2030 自由現金流合計 "&TEXT(C{CON["gatev"]},"0.0")&" 百萬美元，入不敷出〕",C{CON["val_lab"]})', None)
_flip = []
if SENS:
    _brec = SENS["cases"][0]["lo"]["rec"]
    _LOW = ("中立", "放空", "強力放空")
    _flip = [4 + i for i, c in enumerate(SENS["cases"]) if any(c.get(sd) and c[sd]["rec"] in _LOW for sd in ("lo", "hi"))]
_SR = [4, 3 + len(SENS['cases'])] if SENS else [4, 4]
_med = '"四家中間兩家的平均"'
crow("bn_frag", "橫幅：評等脆弱點（原則 36）", "—", '=IF(C{d}<0.05,"評等脆弱點：(a) 同業中位數剛好是 "&{m}&"（"&TEXT(mult,"0.00")&" 倍）；(b) 同業估值日 11-13 與現價日同日（IREN 亦為 11-13 收盤；DLR、EQIX 維持 10 月初）、CoreWeave 共識為其財報前；(c) Sens 單一案例中有 "&TEXT({n},"0")&" 案的估值評等會由買進降為中立或更低"{lst},"")'.format(
    d=CON["dist"], m=_med, n=f'SUMPRODUCT(--((ISNUMBER(MATCH(Sens!M{_SR[0]}:M{_SR[1]},{{"中立","放空","強力放空"}},0))+ISNUMBER(MATCH(Sens!N{_SR[0]}:N{_SR[1]},{{"中立","放空","強力放空"}},0)))>0))' if SENS else "0", lst=('&"（如 "&' + '&"、"&'.join(f"Sens!L{x}" for x in _flip[:6]) + ('&" 等）"' if len(_flip) > 6 else '&"）"')) if _flip else ''), None)
V.column_dimensions["A"].width = 64; V.column_dimensions["B"].width = 11
for c in "CDE": V.column_dimensions[c].width = 12
V.column_dimensions["F"].width = 80

# =====================================================================================
# RevDCF：反向 DCF（整本重算反解；scripts/reverse_v50.py）
# =====================================================================================
RV = wb.create_sheet("RevDCF")
RV["A1"] = "RevDCF — 反向 DCF：一次動一個驅動、整本重算，反解「加權目標價＝目標值」所需的值（scripts/reverse_v50.py → data/reverse_v50.json；建置時寫入）"
RV["A2"] = "主導變數＝每 MW 年收入（k 調整係數 k_adj）：其他輸入不變，反解加權目標價＝現價、＝升一級邊界、＝降一級邊界所需的 k_adj，換算成 FY2030 情境 B 每 MW 年收入，與模型基準與公司宣稱新約 ACV 並列。反解值是目標價等於該值的點，不是評等門檻（原則 4）"
RV["A3"] = "讀法：支持現價所需的每 MW 年收入若落在公司宣稱區間內，表示現價隱含公司說法成立；本模型基準以 Tokenomics 持有成本 × k 為錨，不採用公司宣稱值（CoreWeave 樣板 CLAUDE.md 第 14 條）"
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
RC["A1"] = "Recon — 對帳：公司預期與目標（只作對照，不校準）、公司實況驗證、公司合約換算每 MW 收入、Wattonomics、類比、舊版。差距 = 模型 ÷ 對照 − 1"
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
rsec("A. 公司預期與目標 vs 模型（公司未給全年指引；原則 50、59：只對照、不校準）")
rline("rev26", "FY2026 合併營收 vs 9M 實績＋公司 Q4 雲端預期＋模型 Q4 託管", "=g_rev_lo", lambda s: f"=PL!D{PR_[s]['rev']}", "公司稱 Q4 雲端營收約與 Q3 18.1 持平（Interested-party）；三情境 FY2026 相同")
rline("rev26pos", "  FY2026 營收：相對對照值", None, lambda s: f'=IF(PL!D{PR_[s]["rev"]}>g_rev_hi*1.05,"高於對照值 "&TEXT((PL!D{PR_[s]["rev"]}/g_rev_hi-1)*100,"0.0")&" pt",IF(PL!D{PR_[s]["rev"]}<g_rev_lo*0.95,"低於對照值 "&TEXT((1-PL!D{PR_[s]["rev"]}/g_rev_lo)*100,"0.0")&" pt","與對照值相差 5% 以內"))', "方向字樣由公式依數值決定（原則 57）", None, False)
rline("q4c", "Q4 2026 雲端營收：模型 vs 公司預期（約與 Q3 持平）", "=g_q4cloud", lambda s: f"=Thesis!D{TR[s]['rev']}", "模型以期初期末平均在役 MW 計；新合約多在 11–12 月起算，模型 Q4 偏高")
rline("arr27", "2027 年底雲端年化營收（年底計費 MW × 每 MW 年收入）vs 公司 2027 年中目標 365", "=g_arr27", lambda s: f"=Thesis!F{TR[s]['conn']}*u_bill*Thesis!F{TR[s]['pm']}", "時點不同（年底 vs 年中）；目標為 Interested-party")
rline("intnew", "  新增借款當年未計的半年利息（FY2027；只標示，不進模型）", None, lambda s: f"=0.5*(r_new*CF!E{CR_[s]['ig']}+r_junk*CF!E{CR_[s]['hy']})", "利息以平均餘額計時約多出此數；風險登錄第 6 條")
rline("ebq4", "  Q4 調整後 EBITDA 率：模型 vs Q3 實績", "=aeb_q3/rev_q3", lambda s: f"=Thesis!D{TR[s]['ebm']}", "Q3 實績含終止費 3.4、一般及管理 18.0；模型營業費用率 20% 為規模效應假設", "0.0%")
rline("depq4", "  Q4 折舊攤銷季增率（模型 Q4 ÷ Q3 實績 − 1）", None, lambda s: f"=Thesis!D{TR[s]['dep']}/da_q3-1", "在建工程 312.3 自 Q4 起依機房年限折舊", "0.0%", False)
rline("rev4req", "  達到對照值所需的 Q4 每 MW 年收入（÷ Q4 計費 MW；反解，只展示、不校準）", f"=Thesis!D{TR['B']['pm']}", lambda s: f"=(g_rev_lo-rev9m-colo_q4)/(Thesis!D{TR[s]['bill']}*0.25)", "在役 MW 為假設，此列同時反映 MW 假設與每 MW 收入")
rline("rev4mw", "  或：維持模型每 MW 收入時所需的 Q4 平均計費 MW", f"=Thesis!D{TR['B']['bill']}", lambda s: f"=(g_rev_lo-rev9m-colo_q4)/(Thesis!D{TR[s]['pme']}*0.25)", "對照：公司未揭露在役 MW", "#,##0.0")
rsec("B. RPO 與 backlog（同口徑：RPO 排程 vs 模型營收）")
rline("rpo27", "FY2027：雲端 RPO 排程 vs 模型雲端營收", f"=RPO!E{RPB['sched']}", lambda s: f"=RPO!E{RPB['rev' + s]}", "RPO 年度分布為 10-Q 揭露（Verified）")
rline("rpo29", "FY2029：雲端 RPO 排程 vs 模型雲端營收", f"=RPO!G{RPB['sched']}", lambda s: f"=RPO!G{RPB['rev' + s]}", "")
rsec("C. 公司實況驗證（CLAUDE.md 第 15 條：差距 >10% 找有證據的機制，找不到則 Tokenomics 為基準、公司實際為 Sens）")
rline("q9rev", "9M 2026 雲端營收（扣終止費 15.7）：實績 vs 模型估計（計費 MW × 每 MW 年收入 × 0.75）", "=rev9m_ai-15.7", lambda s: f"=Thesis!C{TR[s]['revchk']}", "在役 MW 為假設（公司未揭露）：差距同時反映 MW 假設與 Tokenomics 錨 × k")
rline("q3pm", "Q3 每 MW 年營收（雲端扣終止費 × 4 ÷ 平均在役假設）vs 模型（每 MW 年收入 × 計費比例）", f"=Q!D{QR['permw']}", lambda s: f"=Thesis!C{TR[s]['pm']}*u_bill", "分母在役 MW 為假設")
rline("q3eb", "Q3 調整後 EBITDA 率：實績 vs 模型 FY2027", "=aeb_q3/rev_q3", lambda s: f"=Thesis!F{TR[s]['ebm']}", "模型 FY2027 以 Tokenomics 營運成本＋租金＋營業費用率＋託管 EBITDA；Q3 實績含終止費；差距 >10% 時 sga_r 有 Sens", "0.0%")
rsec("D. 公司合約換算每 MW 年收入（Interested-party；只對照、不改基準 k）")
rline("j4m", "每 MW 年收入：已簽 B300 合約換算中點 vs 模型 FY2027（情境 B 欄＝基準）", "=(p_claim_lo+p_claim_hi)/2", lambda s: f"=Thesis!F{TR[s]['pm']}", "合約 TCV ÷ 年數 ÷ Tokenomics 每 GPU 功耗換算 MW（B300 HGX 實際功耗較低，換算偏高估每 MW 收入）")
rline("j4tk", "  Tokenomics 錨（GB300 IF_HoldEcon）vs 公司宣稱中點", "=(p_claim_lo+p_claim_hi)/2", lambda s: "=he_gb300", "")
rline("j4s", "  短期合約（Prime GB200 12 個月）vs 模型 FY2027", "=p_short", lambda s: f"=Thesis!F{TR[s]['pm']}", "隨需／短天期；模型隨需占比基準 0%（Sens 20%）")
if REV:
    rline("j4req", "  支持現價所需每 MW 年收入（FY2030 情境 B；RevDCF 反解）vs 公司宣稱中點", "=(p_claim_lo+p_claim_hi)/2", lambda s: f"=RevDCF!G{RVN['px']}" if s == "B" else None, "反解值不是評等門檻（原則 4）")
rsec("E. 類比與容量對照")
rline("iren", "IREN–Microsoft 長約每 MW 年收入（Analogy，原則 61）vs 模型 FY2027", "=p_long", lambda s: f"=Thesis!F{TR[s]['pm']}", "差異：IREN 自有機房（價格含設施）、200 MW、GB300；WhiteFiber 雲端在第三方機房、合約 1–5 年、預付約 40%")
rline("wat26", "Wattonomics 2026 年底上線 MW vs 模型（原則 54）", "=wat26", lambda s: f"=Thesis!E{TR[s]['conn']}", "wattonomics_snapshot_v0.3.5 WIF_PlayerGW_YE_wyfi_2026", "#,##0.0")
rline("wat30", "Wattonomics 2030 年底上線 MW vs 模型", "=wat30", lambda s: f"=Thesis!I{TR[s]['conn']}", "同上 2030", "#,##0.0")
rsec("F. 舊法與舊版（只作對照）")
rline("old", "v0.2.1 加權目標價（舊 JS 引擎，現價 7.80）vs v0.4.0 情境目標價", "=tp_v053", lambda s: f"=Val!{'CDE'[SC.index(s)]}{CON['tps']}", "舊版保留為對照；差異來源見 README 版本表與 v0.6.0 節")
rline("cons", "共識 FY2027 營收 × 40%（EBITDA 代理）vs 模型調整後 EBITDA", "=cons_eb27", lambda s: f"=PL!E{PR_[s]['aeb']}", "只作對照（原則 15）")
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
SO["A1"] = "Sources — 來源清單（事實總帳 whitefiber/data/whitefiber_facts_20261115.json；URL 與擷取日）"
facts = json.load(open(os.path.join(HERE, "data", "whitefiber_facts_20261115.json"), encoding="utf-8"))
hdr(SO, 3, ["代號", "來源"])
for i, (k, v) in enumerate(facts["sources"].items()):
    SO.cell(4 + i, 1, k); SO.cell(4 + i, 2, v)
SO.cell(5 + len(facts["sources"]), 1, "Tokenomics"); SO.cell(5 + len(facts["sources"]), 2, TKSRC + "；名稱清單 whitefiber/data/tokenomics_names.txt")
SO.column_dimensions["A"].width = 14; SO.column_dimensions["B"].width = 160

# =====================================================================================
# HTML 版面規格 v1.3 需要的附加列（附加在各頁最後）
#   Seg：分部營收推導（長約／隨需）；PL：標準損益表（GAAP，歷史 3 年＋預測 5 年＋9M 比較欄）；Sens：O 欄白話標籤；Sources：SEC XBRL 來源
# =====================================================================================
HIST = json.load(open(os.path.join(HERE, "data", "whitefiber_is_hist_20261115.json"), encoding="utf-8"))
HY = HIST["years"]
TB_ = TR["B"]; SB_ = SR["B"]; PB_ = PR_["B"]
# ---- Seg：分部營收推導（情境 B）
r = SG.max_row + 2
SG.cell(r, 1, "分部營收推導（情境 B；HTML 驅動因子章）：雲端服務分部營收＝計費 MW × 每 MW 年收入，依合約型態拆成長約與隨需（公司未揭露拆分；拆分比例為 In 頁隨需占比假設）").font = BOLD; SEGD = {"title": r}; r += 1
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
sgrow("act", "實績（公司未揭露長約／隨需拆分）", "US$M", {"C": f"=C{SB_['seg']}", "D": "=rev9m_ai"}, "#,##0.000", "FY2025＝10-K；FY2026＝9M 分部實績（10-Q；含終止費 15.7）")
sgrow("long", "長約營收（計費 MW ×（1 − 隨需占比）× 長約每 MW 年收入）", "US$M", {SGC[y]: f"={SGC[y]}{SEGD['bill']}*(1-{SGC[y]}{SEGD['od']})*{SGC[y]}{SEGD['pl']}{LEN[y]}" for y in SGC})
sgrow("spot", "隨需營收（計費 MW × 隨需占比 × 隨需每 MW 年收入）", "US$M", {SGC[y]: f"={SGC[y]}{SEGD['bill']}*{SGC[y]}{SEGD['od']}*{SGC[y]}{SEGD['ps']}{LEN[y]}" for y in SGC})
sgrow("ovr", "新增產能改用公司宣稱價格的差額（基準為 0；只在敏感度案例不為 0）", "US$M",
      {SGC[y]: f"=Thesis!{THC[y]}{TB_['rev']}-Thesis!{THC[y]}{TB_['bill']}*Thesis!{THC[y]}{TB_['pm']}{LEN[y]}" for y in SGC})
sgrow("tot", "合計（＝雲端服務分部營收）", "US$M", {"C": f"=C{SEGD['act']}", **{SGC[y]: f"=SUM({SGC[y]}{SEGD['act']}:{SGC[y]}{SEGD['ovr']})" for y in SGC}})
sgrow("lsh", "長約占營收比例（不含未拆分實績）", "%", {SGC[y]: f"={SGC[y]}{SEGD['long']}/({SGC[y]}{SEGD['long']}+{SGC[y]}{SEGD['spot']})" for y in SGC}, "0.0%")
sgrow("chk", "檢查：分部營收推導合計 − 雲端分部營收（應為 0）", None, {c: f"={c}{SEGD['tot']}-{c}{SB_['seg']}" for c in "CDEFGHIJ"}, "0.000")
# ---- PL：標準損益表（GAAP；情境 B）
r = PL.max_row + 2
PLS = {"title": r}
PL.cell(r, 1, "標準損益表（GAAP；情境 B；HTML 損益章讀取）：FY2023 找不到（2025 年分拆上市，10-K 只含 2024–2025）；FY2024–FY2025＝10-K 實績；FY2026＝9M 實績（Q3 10-Q）＋第四季模型；FY2027–FY2030＝模型；K 欄＝9M 2026 實績（比較欄）").font = BOLD; r += 1
PL.cell(r, 1, "口徑：GAAP（公司另揭露調整後 EBITDA）。營業成本＝雲端＋託管營業成本，不含折舊；預測期營業成本＝營運費用（Tokenomics）＋機房租金＋託管與其他營業成本（託管營收 ×（1 − 託管 EBITDA 率）），營業費用＝一般及管理（扣 SBC）＋股票薪酬＋折舊攤銷。FY2024–FY2025 利息未逐項登錄，利息與其他合計列在其他收入。歷史來源：" + HIST["source"]); PLS["note"] = r; r += 1
hdr(PL, r, ["項目", "單位", "FY2023A", "FY2024A", "FY2025A", "FY2026E", "FY2027E", "FY2028E", "FY2029E", "FY2030E", "9M 2026A", "說明／來源"]); PLS["hdr"] = r; r += 1
HC = {"FY2024": "D", "FY2025": "E", "9M2026": "K"}
MC = {2026: "F", 2027: "G", 2028: "H", 2029: "I", 2030: "J"}
ALLC = "DEFGHIJK"
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
SRC_K = "10-K（FY2024–FY2025）／Q3 10-Q（9M 2026），見 Sources K10_10Q"
YH3 = ("FY2024", "FY2025", "9M2026")
plrow("rev", "營收", "US$M", {"D": "=Q!F4", "E": "=Q!G4", **{MC[y]: f"={yc(y)}{PB_['rev']}" for y in MC}, "K": "=rev9m"}, note="FY2024–FY2025：Q 頁（10-K）；預測：上方情境 B 合併營收；FY2023 找不到")
plrow("revg", "營收年增率（YoY）", "%", {c: f"={c}{PLS['rev']}/{p}{PLS['rev']}-1" for p, c in zip("DEFGHI", "EFGHIJ")}, "0.0%")
plrow("cogs", "營業成本（COGS，不含折舊）", "US$M", {**hin("cogs", YH3), "F": f"=K{r}+Thesis!D{TB_['opex']}+Thesis!D{TB_['rent']}+Thesis!D{TB_['colo']}*(1-colo_m)",
      **{MC[y]: f"={yc(y)}{PB_['opex']}+{yc(y)}{PB_['rent']}+{yc(y)}{PB_['oth']}*(1-colo_m)" for y in range(2027, 2031)}}, inputs="DEK", note="歷史：" + SRC_K + "；預測＝營運費用（Tokenomics）＋機房租金＋託管與其他營業成本")
plrow("gp", "毛利", "US$M", {c: f"={c}{PLS['rev']}-{c}{PLS['cogs']}" for c in ALLC})
plrow("gpm", "毛利率", "%", {c: f"={c}{PLS['gp']}/{c}{PLS['rev']}" for c in ALLC}, "0.0%")
plrow("opx", "營業費用（一般及管理含 SBC、折舊攤銷、減損）", "US$M", {**hin("opex", YH3), "F": f"=K{r}+Thesis!D{TB_['sga']}+Thesis!D{TB_['sbc']}+Thesis!D{TB_['dep']}",
      **{MC[y]: f"={yc(y)}{PB_['sga']}+{yc(y)}{PB_['sbc']}+{yc(y)}{PB_['da']}" for y in range(2027, 2031)}}, inputs="DEK", note="歷史＝折舊攤銷＋一般及管理＋減損（" + SRC_K + "）；預測＝一般及管理（扣 SBC）＋股票薪酬＋折舊攤銷")
plrow("oi", "營業利益", "US$M", {"D": "=Q!F5", "E": "=Q!G5", **{MC[y]: f"={yc(y)}{PB_['oi']}" for y in MC}, "K": "=oi9m"}, note="預測：上方情境 B 營業利益（GAAP）")
plrow("oim", "營業利益率", "%", {c: f"={c}{PLS['oi']}/{c}{PLS['rev']}" for c in ALLC}, "0.0%")
plrow("int", "利息淨額（負值＝費用）", "US$M", {**hin("int", ("FY2024", "FY2025")), **{MC[y]: f"=-{yc(y)}{PB_['int']}" for y in MC}, "K": "=-int9m"}, inputs="DE", note="FY2024–FY2025 未逐項登錄（記 0，併入其他收入）；預測：上方情境 B 利息費用（CF 頁）")
plrow("oth", "其他收入（淨額；9M 含誘導轉換費用 −80.8）", "US$M", {**hin("oth", YH3), "F": f"=K{r}", **{MC[y]: 0 for y in range(2027, 2031)}}, inputs="DEGHIJK",
      note="歷史：稅前 − 營業利益 − 利息（Derived）；預測期設 0；FY2026＝9M 實績＋第四季 0")
plrow("pti", "稅前利益", "US$M", {c: f"={c}{PLS['oi']}+{c}{PLS['int']}+{c}{PLS['oth']}" for c in ALLC}, note="FY2026 含 9M 實際其他收入，故與上方「稅前利益」列（只含模型項）不同；稅後淨利兩者一致")
plrow("tax", "所得稅（正＝費用）", "US$M", {**hin("tax", YH3), "F": f"=K{r}+{yc(2026)}{PB_['tax']}", **{MC[y]: f"={yc(y)}{PB_['tax']}" for y in range(2027, 2031)}}, inputs="DEK",
      note="預測：上方情境 B 所得稅；FY2026＝9M 實績＋第四季模型")
plrow("ni", "稅後淨利", "US$M", {c: f"={c}{PLS['pti']}-{c}{PLS['tax']}" for c in ALLC})
plrow("nim", "淨利率", "%", {c: f"={c}{PLS['ni']}/{c}{PLS['rev']}" for c in ALLC}, "0.0%")
plrow("eps", "EPS（GAAP，基本＝稀釋）", "US$", {"K": -3.25, **{MC[y]: f"={yc(y)}{PB_['eps']}" for y in MC}}, "0.00", inputs="K",
      note="歷史：9M 10-Q；FY2024–FY2025 未登錄（找不到）；預測：上方情境 B EPS（淨利 ÷ 稀釋股數）")
plrow("sh", "完全稀釋加權股數", "百萬股", {"K": 39.581712, **{MC[y]: f"={yc(y)}{PB_['sh']}" for y in MC}}, "0.000", inputs="K")
PL.cell(r, 1, "比率").font = BOLD; PLS["ratio"] = r; r += 1
plrow("nig", "淨利年增率（÷ 前一年淨利絕對值；負值＝虧損擴大）", "%", {c: f"=({c}{PLS['ni']}-{p}{PLS['ni']})/ABS({p}{PLS['ni']})" for p, c in zip("DEFGHI", "EFGHIJ")}, "0.0%")
plrow("chk1", "檢查：毛利 − 營業費用 − 營業利益（應為 0）", None, {c: f"={c}{PLS['gp']}-{c}{PLS['opx']}-{c}{PLS['oi']}" for c in ALLC}, "0.000")
plrow("chk2", "檢查：稅後淨利 − 10-K／10-Q 或上方模型淨利（應為 0）", None, {"D": f"=D{PLS['ni']}-Q!F8", "E": f"=E{PLS['ni']}-Q!G8",
      **{MC[y]: f"={MC[y]}{PLS['ni']}-{yc(y)}{PB_['ni']}" for y in MC}, "K": f"=K{PLS['ni']}-ni9m"}, "0.000")
plrow("chk3", "檢查：稅前利益 − 10-K／10-Q 推得值（應為 0）", None, {HC[y]: f"={HC[y]}{PLS['pti']}-({HY[y]['pti']})" for y in HC}, "0.000")
PL.column_dimensions["L"].width = 70
# ---- CF：累計自由現金流（HTML 命題章收支關卡圖的折線；FY2030 欄＝Fund 頁收支關卡）
r = CF.max_row + 2
CFC = {"title": r}
CF.cell(r, 1, "累計自由現金流（FY2026 起；HTML 命題章收支關卡圖；FY2030 欄＝收支關卡）").font = BOLD; r += 1
for s in SC:
    CFC[s] = r; CF.cell(r, 1, f"累計自由現金流：{SCN[s]}"); CF.cell(r, 2, "US$M")
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
V.cell(r, 1, "估值章四個目標價（HTML 估值章表與圖；E 欄＝圖上簡稱）").font = BOLD; CON["t4_title"] = r; r += 1
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
t4("t4_comp", "Comparable 目標價（情境加權）", f"=C{CON['comp']}", "Comparable", "分部加總：雲端分部 FY2028 EBITDA × 新雲同業中位數＋託管分部 × 上市資料中心同業中位數（Andy 2026-11-15），情境 25／50／25 加權")
t4("t4_tp", "加權目標價（DCF 與 Comparable 加權；橫幅主值）", f"=C{CON['tp']}", "加權", f'="DCF "&TEXT(w_dcf,"0%")&"＋Comparable "&TEXT(1-w_dcf,"0%")&"；投資建議依此判定（再套收支關卡）"')
t4("t4_cons", "市場共識目標價（11 家平均，2026-11-10，Q3 財報後）", "=cons_tp", "市場共識",
   "MarketScreener（S&P Capital IQ）新聞頁 2026-11-10：分析師平均目標價（出處見 Sources 頁）；賣方共識（Interested-party），只作對照，不進目標價")
t4("t4_comp0", "Comparable 目標價（未調整，情境加權）", f"=C{CON['comp0']}", "Comparable（未調整）", f'=IF(ABS(C{CON["comp0"]}-C{CON["comp"]})<1E-9,"相對倍數係數＝1，與調整後相同","相對倍數調整前")')
# ---- Sources：SEC XBRL
r = SO.max_row + 1
SO.cell(r, 1, HIST["source_key"]); SO.cell(r, 2, HIST["source"] + "；資料檔 whitefiber/data/whitefiber_is_hist_20261115.json")

# =====================================================================================
# Coord 頁——三層協調（讀取層級、C0 公司實績 vs 層級隱含值）；結構同 CoreWeave v5.0.3（新雲樣板），換成 WhiteFiber 的玩家 wyfi 與實績。
#   工作單 docs/workorders/20261114_第二批13家_v1.3與三層協調.md 第 2 節；規則 docs/plan/三層基準協調規則.md v1.4.2 C0 第 2 點；tools/coord/README.md
#   追公式：tools/coord/trace_wattonomics.py --player wyfi → whitefiber/data/wattonomics_trace_v0.3.5.json
# =====================================================================================
WATF = os.path.join(HERE, "..", "wattonomics", "data", "wattonomics_snapshot_v0.3.5.json")
WAT = json.load(open(WATF, encoding="utf-8"))
WN = WAT["names"]
WATV = WAT["source"]["version"]
WATFILE = WAT["source"]["file"]
WROW = {x.get("id"): x for x in WAT["rows"] if x.get("id")}
TRF = "whitefiber/data/wattonomics_trace_v0.3.5.json"
TRJ = json.load(open(os.path.join(HERE, "data", "wattonomics_trace_v0.3.5.json"), encoding="utf-8"))
assert TRJ["source"]["sha256"] == WAT["source"]["sha256"] if isinstance(TRJ.get("source"), dict) and "sha256" in TRJ["source"] else True
CD = wb.create_sheet("Coord")
CD.column_dimensions["A"].width = 52
for c_, w_ in zip("BCDEFGHIJK", (30, 14, 16, 18, 12, 22, 70, 18, 70, 22)):
    CD.column_dimensions[c_].width = w_
GAPF = "+0.0%;-0.0%"
CD["A1"] = ("Coord — 三層協調（三層基準協調規則 v1.4.2；工作單 20261114 第二批；新雲樣板 CoreWeave v5.0.3，經 Nebius v0.6.0）。個股不裁決、不改上游值；"
            "只列：（G1）本模型讀取哪些層級名稱與版本、（G2）公司已公布實績 vs 層級隱含值（C0）。算力買賣方不做兩讀數目標價。")
CD["A1"].font = BOLD
rc = 3
CD.cell(rc, 1, "G1. 讀取層級（同源，不判定）：本模型讀的 Tokenomics、Wattonomics 名稱與版本").font = BOLD; G1h = rc; rc += 1
CD.cell(rc, 1, ("本模型只用 Tokenomics 計價與成本（每 MW 年持有成本 IF_HoldEcon、設備資本支出 IF_CapexIT、營運費用 IF_OpexGW、每 GPU 小時持有成本 IF_GPUhrEcon 等，"
                f"清單見 whitefiber/data/tokenomics_names.txt）；快照 Tokenomics {TKV}（{TKS['source']['file']}，commit {TKS['source']['commit'][:7]}；舊版 v0.2.1 讀 v19.0，本版隨樣板改讀 {TKV}）。"
                f"Wattonomics 不進估值：In 頁 wat26、wat30 取自 Wattonomics {WATV} 快照（{WATFILE}），只作對照。"
                "追公式（tools/coord/trace_wattonomics.py --player wyfi，輸出 " + TRF + "）顯示：WhiteFiber（雲端）的年底上線 GW 直接讀 Wattonomics In 頁 mw_wyfi（2025 年 5.5、2026 年 14.574、2027 年 16.548 MW），"
                "其來源註記為 whitefiber/company.json scenarios.mwPath（舊版 v0.2 個股模型以合約 GPU 數換算），玩家資本支出 CapexAI 由這條 GW 路徑 × Tokenomics 每 GW 資本支出推得（WhiteFiber 沒有付款年資本支出列，Wattonomics v0.3.5 的付款年主讀法不適用）；資本支出公司實績與指引（cap_wyfi）空白；個股模型回寫值 gcap_wyfi 已由 Wattonomics v0.3.5 重匯為本公司 v0.4.1 的值（只進 CapexAIModel）。"
                "所以層級的 WhiteFiber 路徑與本公司模型同源，以下各列標「讀取層級（同源，不判定）」，不計入對帳中項目數。"))
G1note = rc; rc += 1
hdr(CD, rc, ["名稱", "層級與版本", "單位", "本模型採用值（或公司揭露）", "層級值", "差距 %", "狀態", "說明"]); rc += 1
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
g1("Wattonomics 2026 年底 WhiteFiber（雲端）上線 MW（WIF_PlayerGW_YE_wyfi_2026）", f"Wattonomics {WATV}", "MW", "=mw_ye26", round(WN["WIF_PlayerGW_YE_wyfi_2026"] * 1000, 3), "#,##0.0",
   "層級值＝In mw_wyfi＝舊版個股模型 scenarios.mwPath（合約 GPU 數換算，同源）；本模型 2026 年底為假設（同樣以合約 GPU 換算）")
g1("Wattonomics 2030 年底 WhiteFiber（雲端）上線 MW（WIF_PlayerGW_YE_wyfi_2030）", f"Wattonomics {WATV}", "MW", f"=Thesis!I{TR['B']['conn']}", round(WN["WIF_PlayerGW_YE_wyfi_2030"] * 1000, 3), "#,##0.0",
   "層級 2028 年起持平 16.548（舊版個股模型只到 2027 年已簽約合約）；對照本模型情境 B")
g1("Wattonomics 2026 年 WhiteFiber 資本支出（個股模型回寫 WIF_Gap_Player_wyfi_2026_CapexAIModel）", f"Wattonomics {WATV}", "US$M", f"=Thesis!E{TR['B']['cx']}", round(WN["WIF_Gap_Player_wyfi_2026_CapexAIModel"] * 1000, 3), "#,##0.0",
   "層級值＝WhiteFiber v0.4.1 個股模型 FY2026 資本支出合計回寫（gcapai_wyfi＝gcap_wyfi，Thesis!E70；Wattonomics v0.3.5 重匯，v0.3.2 為舊版 v0.2 的 473.2），同源；本版 v0.4.2 基準輸入與 v0.4.1 相同，差距應為 0。WhiteFiber 沒有付款年資本支出列，Wattonomics Gap 主讀法對它仍為 GW 推算，CapexAI 不讀此值")
G1b = rc - 1
rc += 1
# ---------------- G2：C0（實績）----------------
CD.cell(rc, 1, "G2. C0（實績）：公司已公布實績 vs 層級隱含值（規則 v1.4.2 C0 第 2 點；工作單 2026-11-14）").font = BOLD; G2h = rc; rc += 1
CD.cell(rc, 1, ("方法：個股側只用已公布期間的公司實績（2026 年前三季＝2026-01～2026-09，Q3 10-Q；指引、目標、RPO、合約換算不算實績）。"
                "層級側：雲端營收＝Wattonomics 玩家 MW 年平均 × Tokenomics 每 MW 持有成本（IF_HoldEcon 在役世代加權；層級側稱「持有成本隱含」，不乘本模型的定價倍數 k 與計費比例）；"
                "資本支出＝Wattonomics 玩家 2026 年資本支出（WIF_Gap_Player_wyfi_2026_CapexAI，只含雲端）。"
                "時差：主讀數＝2026 年已公布期間累計實績 vs 層級 CY2026 × 已公布月數 ÷ 12（線性近似，Derived）；第二讀數＝最近一季（Q3）× 4 vs 層級 CY2026（Q3 中點在 8 月中；方向依該季相對全年平均的位置而定，不得寫成夾住真值）。"
                "判定：兩讀數差距都超過 ±門檻且同號＝對帳中，只有一個超過＝時差邊界（參考），都在內＝一致；層級值同源者標「參考（同源校準）」、不判定。"
                "同源檢查（追公式）：層級 MW 路徑讀舊版個股模型（mw_wyfi），資本支出由該路徑推得，所以營收、資本支出兩列都標參考（同源校準），結論「無可比實績（層級值同源）」；在役電力公司未揭露，無可比實績。"))
G2note = rc; rc += 1
CD.cell(rc, 1, "對帳門檻（±）"); CD.cell(rc, 4, 0.15).number_format = "0%"; CD.cell(rc, 4).font = BLUE; CD.cell(rc, 8, "三層基準協調規則 C0／工作單：差距超過 ±15% 標「對帳中」；門檻為輸入（藍字）"); TOLC = f"$D${rc}"; rc += 1
hdr(CD, rc, ["項目", "期間／口徑", "單位", "公司實績", "層級隱含（同期間換算）", "差距 %", "狀態／判定", "公司實績口徑與出處", "公司實績等級", "層級口徑與出處", "層級等級"]); rc += 1
G2a = rc
PRU = SRCF["Q3_10Q"]
LAYSRC = f"Wattonomics {WATV}（{WATFILE}）"
SAMEC = "參考（同源校準）"
def c0row(item, per, unit, val, fmt, src="", tag="", lsrc="", ltag="", blue=False):
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
r_mo = c0row("已公布月數", "2026-01～2026-09", "月", 9, "0", "Q3 2026 10-Q（2026 年前三季）", "Verified", blue=True)
r_rev9 = c0row("雲端服務營收 2026 年前三季累計（公司實績）", "2026-01～2026-09", "US$M", "=rev9m_ai", "#,##0.000", "Q1–Q3 10-Q 分部營收（In ai1–ai3；含 Initial Customer 終止費 15.7）", "Verified")
r_revq = c0row("雲端服務營收 Q3（公司實績）", "2026-07～2026-09", "US$M", "=rev_q3_ai", "#,##0.000", PRU + "：雲端服務營收 18.106（含終止費追加 3.4）", "Verified")
r_cx9 = c0row("資本支出 2026 年前三季（購置 PP&E 與設備訂金；全公司）", "2026-01～2026-09", "US$M", "=cx9m_cash", "#,##0.000", PRU + "：現金流量表 549.804（含 NC-1 建置與託管）", "Verified")
r_cxq = c0row("資本支出 Q3（9M − 上半年 344.7）", "2026-07～2026-09", "US$M", "=cx_q3_def", "#,##0.000", "Derived（10-Q 9M − Q2 10-Q 上半年）", "Verified（推得）")
r_mw = c0row("雲端在役電力 2026-09-30（公司揭露）", "2026-09-30（時點）", "MW", None, "#,##0.0", "公司未揭露：10-Q 只列冰島已簽 5.5 MW-IT、雪梨 2.5 MW（2026 第四季起），其餘站點 MW 未揭露", "—")
r_l25 = c0row("層級：玩家 GW 年底 2025（WIF_PlayerGW_YE_wyfi_2025，× 1000）", "2025 年底", "MW", round(WN["WIF_PlayerGW_YE_wyfi_2025"] * 1000, 3), "0.000", "", "", f"{LAYSRC} 名稱 WIF_PlayerGW_YE_wyfi_2025（＝In mw_wyfi 2025 年 5.5）", "同源（舊版個股模型）", blue=True)
r_l26 = c0row("層級：玩家 GW 年底 2026（WIF_PlayerGW_YE_wyfi_2026，× 1000）", "2026 年底", "MW", round(WN["WIF_PlayerGW_YE_wyfi_2026"] * 1000, 3), "0.000", "", "", f"{LAYSRC} 名稱 WIF_PlayerGW_YE_wyfi_2026＝In mw_wyfi 2026 年 14.574（whitefiber/company.json scenarios.mwPath，舊版個股模型；追公式 {TRF}）", "同源（舊版個股模型）", blue=True)
r_lhe = c0row("層級：在役世代加權每 MW 年持有成本 FY2026（Tokenomics IF_HoldEcon 加權）", "CY2026", "US$M／MW·年", f"=Thesis!E{TR['B']['anc']}", "#,##0.000", "", "", f"Tokenomics {TKV}（{TKS['source']['commit'][:7]}）IF_HoldEcon 依本模型在役世代占比加權（未乘 k）", "Derived（世代占比 Assumed）")
r_lrev = c0row("持有成本隱含雲端營收 CY2026（玩家 MW 年平均 × 加權每 MW 年持有成本）", "CY2026", "US$M", f"=(D{r_l25}+D{r_l26})/2*D{r_lhe}", "#,##0.000", "", "", "＝（YE2025＋YE2026）÷ 2 × 每 MW 年持有成本；MW 路徑同源", "Derived（MW 量同源）")
r_lcx = c0row("層級：玩家 AI 資本支出 CY2026（WIF_Gap_Player_wyfi_2026_CapexAI，× 1000）", "CY2026", "US$M", round(WN["WIF_Gap_Player_wyfi_2026_CapexAI"] * 1000, 3), "#,##0.000", "", "", f"{LAYSRC} 名稱 WIF_Gap_Player_wyfi_2026_CapexAI（新增 MW × Tokenomics 每 GW 資本支出；只含雲端，不含託管建置；新增 MW 同源）", "同源（舊版個股模型）", blue=True)
r_lcm = c0row("層級：個股模型回寫資本支出 CY2026（WIF_Gap_Player_wyfi_2026_CapexAIModel，× 1000）", "CY2026", "US$M", round(WN["WIF_Gap_Player_wyfi_2026_CapexAIModel"] * 1000, 3), "#,##0.000", "", "", f"{LAYSRC} 名稱 WIF_Gap_Player_wyfi_2026_CapexAIModel＝In gcapai_wyfi（WhiteFiber v0.4.1 個股模型 Thesis!E70；原值 US$M ÷ 1000）", "同源（個股模型回寫）", blue=True)
CD.cell(rc, 1, "Wattonomics 快照核對（上列藍字常數取自此檔）"); CD.cell(rc, 2, f"{WATFILE}（快照 whitefiber 讀 wattonomics/data/wattonomics_snapshot_{WATV}.json）"); CD.cell(rc, 8, f"xlsx sha256 {WAT['source']['sha256']}（快照 source 欄）；建置時逐名讀 WIF_PlayerGW_YE_wyfi_2025／2026、WIF_Gap_Player_wyfi_2026_CapexAI／CapexAIModel（GW、十億美元 × 1000 換成 MW、百萬美元），不手寫"); r_sha = rc; rc += 1
r_m1 = reading("雲端營收｜主讀數：已公布期間累計 vs 持有成本隱含 CY2026 × 已公布月數 ÷ 12", "2026-01～2026-09 累計", "US$M", f"=D{r_rev9}", f"=D{r_lrev}*D{r_mo}/12", "#,##0.000", "前三季雲端營收（含終止費）", "Verified", "持有成本隱含營收 × 9 ÷ 12", "Derived（同源）", status=SAMEC)
r_m2 = reading("雲端營收｜第二讀數：最近一季（Q3）× 4 vs 持有成本隱含 CY2026", "2026-07～2026-09 年化", "US$M", f"=D{r_revq}*4", f"=D{r_lrev}", "#,##0.000", "Q3 雲端營收 × 4（含終止費 3.4）", "Verified／Derived", "持有成本隱含營收（全年）", "Derived（同源）", status=SAMEC)
CD.cell(rc, 1, "雲端營收：C0（實績）判定").font = BOLD; CD.cell(rc, 2, "層級 MW 路徑讀舊版個股模型")
CD.cell(rc, 7, "參考（同源校準）：無可比實績（層級值同源）"); r_j1 = rc; rc += 1
r_c1 = reading("資本支出｜主讀數：已公布期間累計 vs 層級 CapexAI CY2026 × 已公布月數 ÷ 12", "2026-01～2026-09 累計", "US$M", f"=D{r_cx9}", f"=D{r_lcx}*D{r_mo}/12", "#,##0.000", "前三季現金資本支出（全公司，含託管建置）", "Verified", "層級 CapexAI × 9 ÷ 12（只含雲端）", "同源", status=SAMEC)
r_c2 = reading("資本支出｜第二讀數：Q3 × 4 vs 層級 CapexAI CY2026", "2026-07～2026-09 年化", "US$M", f"=D{r_cxq}*4", f"=D{r_lcx}", "#,##0.000", "Q3 現金資本支出 × 4", "Verified（推得）", "層級 CapexAI（全年）", "同源", status=SAMEC)
CD.cell(rc, 1, "資本支出：C0（實績）判定").font = BOLD; CD.cell(rc, 2, "層級資本支出由舊版個股模型 MW 路徑推得；口徑亦不同（層級只含雲端）")
CD.cell(rc, 7, "參考（同源校準）：無可比實績（層級值同源）"); r_j2 = rc; rc += 1
r_w1 = rc
CD.cell(rc, 1, "雲端在役電力：2026-09-30 實績 vs 層級 YE2025→YE2026 內插"); CD.cell(rc, 2, "2026-09-30（時點）"); CD.cell(rc, 3, "MW")
CD.cell(rc, 5, f"=D{r_l25}+(D{r_l26}-D{r_l25})*D{r_mo}/12").number_format = "#,##0.0"
CD.cell(rc, 7, "無可比實績"); CD.cell(rc, 8, "公司未揭露季末雲端在役 MW"); CD.cell(rc, 9, "—"); CD.cell(rc, 10, "層級內插值只列參考（同源）"); CD.cell(rc, 11, "同源"); rc += 1
CD.cell(rc, 1, "雲端在役電力：C0（實績）判定").font = BOLD; CD.cell(rc, 2, "缺公司揭露")
CD.cell(rc, 7, "無可比實績（公司未揭露季末在役 MW）"); r_j3 = rc; rc += 1
r_ctr = rc
CD.cell(rc, 1, "已簽約雲端容量（GPU 換算 MW）vs 層級"); CD.cell(rc, 2, "2026-11（法說）"); CD.cell(rc, 3, "MW"); CD.cell(rc, 4, "=ctr_mw").number_format = "#,##0.0"
CD.cell(rc, 7, "無可比實績"); CD.cell(rc, 8, f"{CALL}：已簽專用運算 GPU 約 20,000 顆（Interested-party），以 Tokenomics 功耗換算（Derived）"); CD.cell(rc, 9, "Interested-party")
CD.cell(rc, 10, "Wattonomics 沒有 WhiteFiber 簽約容量序列（缺：層級側簽約 GW）"); CD.cell(rc, 11, "—"); rc += 1
r_cxm1 = reading("參考：資本支出主讀數改對個股模型回寫值（CapexAIModel × 9 ÷ 12）", "2026-01～2026-09 累計", "US$M", f"=D{r_cx9}", f"=D{r_lcm}*D{r_mo}/12", "#,##0.000", "前三季現金資本支出", "Verified", "層級 CapexAIModel（v0.4.1 回寫）× 9 ÷ 12", "同源", status="參考（同源，不判定）")
r_cxm2 = reading("參考：資本支出第二讀數改對個股模型回寫值", "2026-07～2026-09 年化", "US$M", f"=D{r_cxq}*4", f"=D{r_lcm}", "#,##0.000", "Q3 現金資本支出 × 4", "Verified（推得）", "層級 CapexAIModel（v0.4.1 回寫）", "同源", status="參考（同源，不判定）")
G2b = rc - 1
rc += 1
CD.cell(rc, 1, "核對：雲端營收 9 個月實績 − Thesis 9M 實績（應為 0）"); CD.cell(rc, 4, f"=D{r_rev9}-Thesis!C{TR['B']['rev']}").number_format = "0.000000"; r_k1 = rc; rc += 1
CD.cell(rc, 1, "核對：資本支出 9 個月實績 − Thesis 9M 公司口徑資本支出（應為 0）"); CD.cell(rc, 4, f"=D{r_cx9}-Thesis!C{TR['B']['cx']}").number_format = "0.000000"; r_k2 = rc; rc += 1
CD.cell(rc, 1, "對帳中項目數（營收、資本支出、在役電力；G1 與參考列不計）"); CD.cell(rc, 4, f'=COUNTIF(G{r_j1},"對帳中")+COUNTIF(G{r_j2},"對帳中")+COUNTIF(G{r_j3},"對帳中")'); r_n = rc; rc += 1
CD.cell(rc, 1, "有判定的項目數（同源與無可比實績不計）"); CD.cell(rc, 4, f'=COUNTIF(G{r_j1},"對帳中")+COUNTIF(G{r_j1},"一致")+COUNTIF(G{r_j1},"時差邊界*")+COUNTIF(G{r_j2},"對帳中")+COUNTIF(G{r_j2},"一致")+COUNTIF(G{r_j2},"時差邊界*")'); r_nj = rc; rc += 1
CD.cell(rc, 1, "C0 對帳摘要（HTML 對帳章讀取）")
CD.cell(rc, 4, f'="實績對帳（公司已公布實績對 Wattonomics、Tokenomics 層級隱含值）：層級的 WhiteFiber（雲端）MW 路徑讀舊版個股模型、資本支出由其推得（同源），所以雲端營收（前三季 "&TEXT(D{r_m1},"0.0")&" 對持有成本隱含同期間 "&TEXT(E{r_m1},"0.0")&"，"&TEXT(F{r_m1},"+0%;-0%")&"）與資本支出（前三季 "&TEXT(D{r_c1},"0.0")&" 對層級同期間 "&TEXT(E{r_c1},"0.0")&"，"&TEXT(F{r_c1},"+0%;-0%")&"；百萬美元）只列參考："&G{r_j1}&"；在役電力："&G{r_j3}&"；簽約容量無可比實績。對帳中項目 "&D{r_n}&" 個、有判定 "&D{r_nj}&" 個（門檻 ±"&TEXT({TOLC},"0%")&"）。個股不裁決，不改上游值；證據回饋走對帳單。"'); r_sum = rc; rc += 1
CD.cell(rc, 1, "同源檢查（哪些部分同源、什麼會消去、這張表實際檢驗什麼）")
CD.cell(rc, 2, ("（一）本模型沒有把校準值寫入 Wattonomics 或 Tokenomics：定價倍數 k、計費比例、世代占比、在役 MW 假設只留在本模型。"
                "（二）Wattonomics 玩家 wyfi 的年底上線 GW 直接讀 In 頁 mw_wyfi（2025 年 5.5、2026 年 14.574、2027 年 16.548 MW），註記來源為 whitefiber/company.json scenarios.mwPath（舊版 v0.2 個股模型以合約 GPU 數換算）；玩家資本支出 CapexAI＝新增 GW × Tokenomics 每 GW 資本支出；公司資本支出實績與指引 cap_wyfi 空白。"
                "（三）因此營收列的層級 MW 量、資本支出列的層級金額都等於舊版本公司模型的推估，比式裡沒有獨立於本公司模型的成分，這張表實際只檢驗「實績 vs 舊版本公司模型」，依規則標「參考（同源校準）」，結論「無可比實績（層級值同源）」。"
                "（四）資本支出口徑另有差異：公司實績含 NC-1 託管建置與設備訂金，層級只含雲端；兩者不可直接比較。Wattonomics v0.3.5 已重匯 gcap_wyfi（本公司 v0.4.1：2026 年 844.6 百萬美元，原 v0.2 的 473.2），只進 CapexAIModel（參考列）；mw_wyfi 仍為舊版 v0.2 的 MW 路徑。交主控：重匯 mw_wyfi 時一併評估，並考慮由公司揭露（已簽 GPU 數）而非個股模型推得層級 MW。")); r_org = rc; rc += 1
CD.cell(rc, 1, "口徑換算與限制")
CD.cell(rc, 2, ("營收：層級側是持有成本（經濟口徑成本，不是價格）；公司雲端實績含 Initial Customer 終止費 15.7（Q2 12.3、Q3 3.4）。資本支出：公司只揭露現金購置（含設備訂金、NC-1 與 MTL 建置），無應計口徑。"
                "在役電力：公司以合約 GPU 數揭露，未揭露季末雲端在役 MW。缺哪一項獨立揭露：季末雲端在役 MW，或 Wattonomics 以公司揭露（非個股模型）推得的 WhiteFiber 路徑。")); r_k = rc; rc += 1
CD.freeze_panes = "B3"
COORD_ROWS = {"G1h": G1h, "G1a": G1a, "G1b": G1b, "G2h": G2h, "G2a": G2a, "G2b": G2b, "tol": int(TOLC.split("$")[-1]), "sha": r_sha, "m1": r_m1, "m2": r_m2, "j1": r_j1, "c1": r_c1, "c2": r_c2, "j2": r_j2, "w1": r_w1, "j3": r_j3,
              "ctr": r_ctr, "cxm1": r_cxm1, "cxm2": r_cxm2, "n": r_n, "nj": r_nj, "sum": r_sum, "org": r_org, "k": r_k, "g1note": G1note, "g2note": G2note, "mo": r_mo, "k1": r_k1, "k2": r_k2}
RC.cell(RCR["wat26"], 9, "讀取層級（同源，不判定）：Wattonomics wyfi 路徑讀舊版個股模型 mw_wyfi（Coord 頁 G1）")
RC.cell(RCR["wat30"], 9, "讀取層級（同源，不判定）：同上；差距是兩個模型版本與情境的差，不是實績對帳（Coord 頁 G1）")

# =====================================================================================
# README（最後建，最前排；文字中的數字一律公式組字）
# =====================================================================================
RD = wb.create_sheet("README", 0)
VC = lambda k: f"Val!$C${CON[k]}"
T_ = lambda ref, f: f'TEXT({ref},"{f}")'
RD["A1"] = ('="WhiteFiber（WYFI）個股模型 ' + VER + '（CoreWeave 新雲樣板重建＋2026Q3 滾入＋三層協調）　資料截止日 "&TEXT(cutoff,"yyyy-mm-dd")&"；成品日 "&TEXT(built,"yyyy-mm-dd")&"；股價基準日 "&TEXT(px_date,"yyyy-mm-dd")&"；build_id '
            + BUILD_ID + '"')
RL = {}
_alt = [i for i, c in enumerate(SENS["cases"]) if c["key"] == "peer_nd_mode"] if SENS else []
ALT_SENT = (f'&"；同業淨負債改用共識年底預估（替代口徑，時點不一致）時加權目標價 "&TEXT(Sens!E{4 + _alt[0]},"0.00")&"（"&Sens!N{4 + _alt[0]}&"）"') if _alt else ""
lines = [
    ("thesis", '="命題：WhiteFiber 以多年期專用運算合約（雲端 RPO "&TEXT(rpo,"0.0")&" 百萬美元）、客戶預付（遞延收入 "&TEXT(def_open,"0.0")&"）與 NC-1 託管長約支撐的擴張，能否以高於資金成本的報酬轉成營收與現金，同時不讓資金缺口失控。命題模組（雲端在役 MW × 每 MW 年收入〔Tokenomics 持有成本 × 定價倍數 k〕、營運費用、租金、資本支出與折舊）驅動雲端服務分部；託管與其他另列"'),
    ("q", "核心問題：小型新雲靠合約與預付擴張，每增加 1 MW 雲端產能是否創造高於資金成本的價值；NC-1 託管長約能否撐住現金流；資金缺口要靠多少新債、可轉債與新股補，股東被稀釋多少。"),
    ("rec", f'="投資建議："&{VC("rec_lab")}&"（加權目標價 US$"&TEXT({VC("tp")},"0.00")&"，現價 US$"&TEXT(px,"0.00")&"〔"&TEXT(px_date,"yyyy-mm-dd")&"〕，上檔空間 "&TEXT({VC("up")},"+0.0%;-0.0%")&"）——條件式：在每 MW 年收入＝Tokenomics 持有成本 × k "&TEXT(k_mix,"0.000")&"（IREN 長約 Analogy）、計費比例 "&TEXT(u_bill,"0%")&"、AI 設備折舊 "&TEXT(life_ai,"0")&" 年、雲端在役 MW 假設（公司未揭露）、託管營收依 NC-1 合約的基準下；失效條件見本頁末"'),
    ("o1", f'="重點觀察 1：DCF 情境加權 "&TEXT({VC("dcf")},"0.00")&"、Comparable "&TEXT({VC("comp")},"0.00")&"（分部加總：雲端分部 FY2028 EBITDA × 新雲同業中位數 "&TEXT(mult,"0.00")&" 倍＋託管分部 × 資料中心同業 "&TEXT(mult_colo,"0.00")&" 倍）；"&{VC("bn_dcf")}&"。上檔空間 "&TEXT({VC("up")},"+0.0%;-0.0%")&"＝DCF 貢獻 "&TEXT({VC("dcfup")},"+0.0%;-0.0%")&"＋Comparable 貢獻 "&TEXT({VC("cmpup")},"+0.0%;-0.0%")&"＋相對倍數調整 "&TEXT({VC("adjup")},"+0.0%;-0.0%"){ALT_SENT}'),
    ("o2", f'="重點觀察 2："&{VC("bn_scn")}&"。情境 B 的 FY2029 稅後 ROIC "&TEXT(Thesis!D{SUM_ROWS[0] + 5},"0.0%")&"，對照 WACC "&TEXT(wacc,"0.0%")&"。"'),
    ("o3", f'="重點觀察 3：收支關卡（FY2026–FY2030 自由現金流合計）情境 A／B／C 為 "&TEXT(Fund!C{FS["gate"]},"0.0")&"／"&TEXT(Fund!D{FS["gate"]},"0.0")&"／"&TEXT(Fund!E{FS["gate"]},"0.0")&" 百萬美元；情境 B 期間最大單年新增舉債 "&TEXT(Fund!D{FS["maxb"]},"0.0")&"、最低年末現金 "&TEXT(Fund!D{FS["minc"]},"0.0")&"、淨負債 ÷ EBITDA 最差 "&TEXT(Fund!D{FS["ndmax"]},"0.00")&" 倍、利息保障最差 "&TEXT(Fund!D{FS["icmin"]},"0.00")&" 倍（"&IF(Fund!D{FS["ndmax"]}>lev_max,"淨負債 ÷ EBITDA 觸發 WARN","淨負債 ÷ EBITDA 未觸發")&"、"&IF(Fund!D{FS["icmin"]}<icov_min,"利息保障觸發 WARN","利息保障未觸發")&"）；2030 年底累計稀釋 "&TEXT(Fund!D{FS["dil"]},"0.0%")'),
    ("o4", '="公司合約換算：已簽 B300 合約每 MW 年收入約 "&TEXT(p_claim_lo,"0.0")&"–"&TEXT(p_claim_hi,"0.0")&"（Interested-party；TCV ÷ Tokenomics 功耗換算 MW），是本模型 FY2027 情境 B 每 MW 年收入 "&TEXT(Thesis!F' + str(TR["B"]["pm"]) + ',"0.00")&" 的 "&TEXT((p_claim_lo+p_claim_hi)/2/Thesis!F' + str(TR["B"]["pm"]) + ',"0.0")&" 倍；新增產能改用合約換算值的 Sens 見 Sens 頁 p_new_ovr 列；不改基準 k"'),
    ("o5", f'="橫幅："&{VC("bn_edge")}&IF({VC("bn_frag")}="","","。"&{VC("bn_frag")})'),
    ("o6", f'="公司預期對照：模型 FY2026 營收 "&TEXT(PL!D{PR_["B"]["rev"]},"0.0")&"，"&Recon!E{RCR["rev26pos"]}&"（對照值 "&TEXT(g_rev_lo,"0.0")&"＝9M 實績＋公司 Q4 雲端預期＋模型 Q4 託管）；2027 年底雲端年化 "&TEXT(Recon!E{RCR["arr27"]},"0.0")&" 對公司 2027 年中目標 "&TEXT(g_arr27,"0")&"（百萬美元）"'),
]
for i, (k, f) in enumerate(lines):
    RD.cell(2 + i, 1, f); RL[k] = 2 + i
r = 2 + len(lines) + 1
RD.cell(r, 1, "會計年度＝日曆年；FY2026＝9M 實績（Q3 10-Q）＋Q4 模型；延伸到 FY2032。金額百萬美元、股數百萬股，MW 為 IT 口徑（雲端在役 MW 公司未揭露，2026 年為假設；託管以合約 MW 計）。"); RL["cal"] = r; r += 1
RD.cell(r, 1, "營收分部層：官方分部為雲端服務與託管服務（另有其他營收）；命題模組驅動雲端服務，託管與其他取 NC-1 合約、MTL 與代管服務路徑（三情境相同）；核對列見 Seg 與 PL 頁。"); RL["seg"] = r; r += 2
RD.cell(r, 1, "失效條件（可觀察證據；出現任一，結論須重估）").font = BOLD; RL["fail"] = r; r += 1
for t in ('="1. 每 MW 年收入：若公司揭露季末在役 MW，而營收 ÷ 平均在役 MW 持續高於 "&TEXT(Thesis!F' + str(TR["B"]["pm"]) + '*1.5,"0.0")&"（本模型 FY2027 × 1.5），定價倍數 k 低估；可證偽指標：季報 AI cloud 營收（扣非經常性）÷ 平均在役 MW"',
          '="2. 資金：若客戶預付持續覆蓋資本支出超過 "&TEXT(prepay_cov*1.5,"0%")&"（本模型 "&TEXT(prepay_cov,"0%")&" × 1.5），或擔保融資利率遠低於 "&TEXT(kd,"0.0%")&"，融資成本與稀釋高估"',
          '="3. 經濟壽命：若 GPU 實際汰換早於 "&TEXT(life_ai,"0")&" 年（公司會計年限），折舊與汰換資本支出低估；可證偽指標：二手 H100 租金與 Silicon Data 指數"'):
    RD.cell(r, 1, t); r += 1
r += 1
RD.cell(r, 1, "風險登錄（未進主模型的尾部風險；原則 10、19）").font = BOLD; RL["risk"] = r; r += 1
for t in ("1. 客戶集中（9M 最大客戶占雲端營收 50%；NC-1 單一客戶 Nscale）：是否建模＝否（事件型）；理由：合約條款未揭露；觀察指標：10-Q 客戶占比、Nscale 融資",
          "2. 控股股東 Bit Digital 持股 59.7%、DDTL 為關係人借款（1.1× MOIC）：是否建模＝部分（DDTL 利率入 kd 區間）；觀察指標：關係人交易附註",
          "3. 可轉債與零履約價買權：2031／2032 可轉債轉換價約 25.91／33.84（遠高於現價），列負債；公司 2026-01 支付 120 的零履約價買權（2031 交割 5.906 百萬股）未自評價股數扣除；是否建模＝否；觀察指標：10-Q 股數與權益附註",
          "4. 站點交付與 NC-2／NC-3 開發（初始 60 毛 MW、僅意向書）：是否建模＝否（基準只含已簽 NC-1 與 MTL）；觀察指標：租約簽署、NC-1 信用協議動用",
          "5. 利率與信用利差：是否建模＝Sens（kd、高息債利率）；觀察指標：NC-1 信用協議 SOFR＋6.00%",
          "6. 利息以期初餘額計（Oracle 樣板慣例）：新增借款當年不計息（Recon「新增借款當年未計的半年利息」列）；是否建模＝否；觀察指標：季報利息費用"):
    RD.cell(r, 1, t); r += 1
r += 1
RD.cell(r, 1, "已知限制").font = BOLD; r += 1
RD.cell(r, 1, "公司未揭露季末雲端在役 MW，2026-09-30 與 2026 年底為假設（以合約 GPU 數經 Tokenomics 功耗換算），公司實況驗證只檢驗 MW 假設與 k 的合併結果；雲端營收含 Initial Customer 終止費 15.7；GPU 伺服器租賃（變動租金）視為自購處理（資本支出偏高、租金偏低）；託管營收不分情境，NC-2／NC-3 與 NC-1 擴充未計入。Comparable 以 4 家新雲同業（CoreWeave、Nebius、IREN、Oracle）單一倍數評整家公司，未分部加總（託管同業約 20 倍，判斷題）；CoreWeave 共識為其財報前；WhiteFiber EBITDA 共識找不到。β 以 US 伺服器／EMS 組代理。FY2023 損益找不到（2025 年分拆上市）。事實數字轉錄自 open PR #451 的事實總帳。"); RL["lim"] = r; r += 2
RD.cell(r, 1, "驅動 → 推導對照（每頁輸入、推導量、因果方向）").font = BOLD; RL["drv"] = r; r += 1
for t in ("In：一列一鍵（值、單位、標記、來源、說明、來源誘因）→ 所有公式以鍵名引用",
          "Thesis：雲端在役 MW 路徑 × 每 MW 年收入（Tokenomics × k）→ 雲端營收；Tokenomics 營運費用、租金、營業費用率、託管 EBITDA → 調整後 EBITDA；新增 MW × IF_CapexIT＋託管資本支出 → 資本支出 → 折舊 → 營業利益",
          "Seg、PL：雲端＋託管與其他 → 合併營收 → 營業利益 → 利息（CF）、稅（NOL）→ 淨利、EPS",
          "CF、Fund、Debt、Lease、RPO、Sites：自由現金流（含客戶預付流入與預付認列）→ 現金 → 投資級債 → 新股 → 高息債；收支關卡與期間流動性；RPO 排程與站點對照（不校準）",
          "Val、RevDCF、Recon、Sens、Coord：DCF（FCFF）＋Comparable → 加權目標價、評等；反解支持現價所需每 MW 收入；指引與公司說法只對照；三層協調"):
    RD.cell(r, 1, t); r += 1
r += 1
RD.cell(r, 1, "展示層（以下各列供 HTML 引用）").font = BOLD; RL["show"] = r; r += 1
for t in ("1. 每 MW 的經濟性：WhiteFiber 每增加 1 MW 雲端產能，收入（Tokenomics 持有成本 × k）扣掉營運費用、租金與折舊後，是否高於資金成本？",
          "2. 資金：客戶預付與 NC-1 租金能覆蓋多少資本支出，其餘缺口靠多少新債與新股補、股東被稀釋多少？",
          "3. 公司合約：已簽合約換算的每 MW 年收入高於 Tokenomics 錨，現價是否已經隱含合約價格持續？"):
    RD.cell(r, 1, t); r += 1
r += 1
RD.cell(r, 1, "命題回覆（HTML 摺疊）").font = BOLD; RL["ans"] = r; r += 1
RD.cell(r, 1, f'="1. 每 MW 經濟性？答：情境 B 的 FY2029 稅後 ROIC "&TEXT(Thesis!D{SUM_ROWS[0] + 5},"0.0%")&"，對照 WACC "&TEXT(wacc,"0.0%")&"；"&IF(Thesis!D{SUM_ROWS[0] + 5}<wacc,"低於資金成本，擴張不創造價值（模型輸出，非獨立證據）","高於資金成本（模型輸出，非獨立證據）")'); r += 1
RD.cell(r, 1, f'="2. 資金？答：情境 B 在 Q4 2026–FY2030 客戶預付流入約 "&TEXT(SUM(CF!D{CR_["B"]["ppin"]}:H{CR_["B"]["ppin"]}),"0.0")&" 百萬美元後，仍需外部融資 "&TEXT(Fund!D{FS["cum"]},"0.0")&" 百萬美元（投資級債 "&TEXT(Fund!D{FS["ig"]},"0.0")&"、新股 "&TEXT(Fund!D{FS["eq"]},"0.0")&"、高息債 "&TEXT(Fund!D{FS["hy"]},"0.0")&"），2030 年底累計稀釋 "&TEXT(Fund!D{FS["dil"]},"0.0%")'); r += 1
if REV:
    RD.cell(r, 1, f'="3. 公司合約？答：支持現價所需的 FY2030 每 MW 年收入（情境 B）為 "&TEXT(RevDCF!G{RVN["px"]},"0.00")&"，"&IF(RevDCF!G{RVN["px"]}<p_claim_lo,"低於","不低於")&"公司合約換算下緣 "&TEXT(p_claim_lo,"0.0")&"；本模型基準 "&TEXT(Thesis!I{TR["B"]["pm"]},"0.00")&"（Tokenomics 錨 × k）"'); r += 1
else:
    RD.cell(r, 1, "3. 公司說法？答：反解快照未產生"); r += 1
r += 1
if DEC:
    RD.cell(r, 1, "目標價變動拆解（scripts/decomp_v50.py）：" + "；".join(f"{d['label']} {d['delta']:+.2f}" for d in DEC["steps"]) + f"；合計 {DEC['total']:+.2f}"); RL["dec"] = r; r += 1
r += 1
RD.cell(r, 1, "展示層 v1.3（HTML 版面規格 v1.3；以下各列供 HTML 引用；白話、不含參數代號）").font = BOLD; RL["s13"] = r; r += 1
_V = VC
_GATE3 = f'TEXT(Fund!C{FS["gate"]},"0.0")&"、"&TEXT(Fund!E{FS["gate"]},"0.0")'
_dcf0 = f'IF({_V("dcf")}=0,"現金流折現法（DCF）三個情境的每股價值都是 0，加權目標價全部來自同業倍數。","")'
L13 = [
    ("hero13", f'=IF({_V("rec")}="收支不足","投資建議「收支不足」：估值面上檔空間 "&TEXT({_V("up")},"+0.0%;-0.0%")&"，估值評等為"&{_V("val_lab")}&"；但基準情境（B）FY2026–FY2030 自由現金流合計 "&TEXT({_V("gatev")},"0.0")&" 百萬美元，入不敷出，依收支關卡不得列為買進。","投資建議「"&{_V("rec")}&"」：上檔空間 "&TEXT({_V("up")},"+0.0%;-0.0%")&"；基準情境 FY2026–FY2030 自由現金流合計 "&TEXT({_V("gatev")},"0.0")&" 百萬美元。")&{_dcf0}'),
    ("ob1", f'="收支是結論的前提：基準情境 FY2026–FY2030 自由現金流合計 "&TEXT({_V("gatev")},"0.0")&" 百萬美元，樂觀與保守情境為 "&{_GATE3}&"；客戶預付與 NC-1 租金已計入營運現金流，仍"&IF({_V("gatev")}<0,"入不敷出","有餘")&"；估值面上檔空間 "&TEXT({_V("up")},"+0.0%;-0.0%")&"，距最近的評等級距 "&TEXT({_V("dist")}*100,"0.0")&" 個百分點。"'),
    ("ob2", f'="目標價的組成：DCF 情境加權每股 "&TEXT({_V("dcf")},"0.00")&" 美元，同業倍數法 "&TEXT({_V("comp")},"0.00")&" 美元（分部加總：雲端分部 × 新雲同業 "&TEXT(mult,"0.00")&" 倍、託管分部 × 資料中心同業 "&TEXT(mult_colo,"0.00")&" 倍；基準情境每股雲端 "&TEXT(Val!D{CO["pcc"]},"0.00")&"、託管 "&TEXT(Val!D{CO["pch"]},"0.00")&"、扣淨負債 "&TEXT(Val!D{CO["pcn"]},"0.00")&"），各占 "&TEXT(w_dcf,"0%")&" 得 "&TEXT({_V("tp")},"0.00")&" 美元；雲端同業倍數"&IF(ISNUMBER({_V("mdn")}),"降到 "&TEXT({_V("mdn")},"0.00")&" 倍，估值評等就降一級。","估值評等已是最低級。")'),
    ("ob3", f'=IF({_V("npvAB")}<0,"擴張越多、價值越低：樂觀相對基準情境每多 1 MW 雲端產能，DCF 企業價值減少 "&TEXT(-{_V("npvAB")},"0.0"),"擴張創造價值：樂觀相對基準情境每多 1 MW 雲端產能，DCF 企業價值增加 "&TEXT({_V("npvAB")},"0.0"))&" 百萬美元；基準情境 FY2029 稅後投入資本報酬率 "&TEXT(Thesis!D{SUM_ROWS[0] + 5},"0.0%")&"，"&IF(Thesis!D{SUM_ROWS[0] + 5}<wacc,"低於","高於")&"資金成本 "&TEXT(wacc,"0.0%")&"。"'),
    ("ob4", f'="缺口靠舉債與增資補：基準情境 2026 年第四季至 FY2030 需外部融資 "&TEXT(Fund!D{FS["cum"]},"0.0")&" 百萬美元（投資級債 "&TEXT(Fund!D{FS["ig"]},"0.0")&"、新股 "&TEXT(Fund!D{FS["eq"]},"0.0")&"、高息債 "&TEXT(Fund!D{FS["hy"]},"0.0")&"），2030 年底股東累計稀釋 "&TEXT(Fund!D{FS["dil"]},"0.0%")&"；淨負債對 EBITDA 最差 "&TEXT(Fund!D{FS["ndmax"]},"0.00")&" 倍"&IF(Fund!D{FS["ndmax"]}>lev_max,"，超過 "&TEXT(lev_max,"0.0")&" 倍警戒","")&"。"'),
    ("ob5", '="公司合約換算的每 MW 年收入是模型的 "&TEXT((p_claim_lo+p_claim_hi)/2/Thesis!I' + str(TR["B"]["pm"]) + ',"0.0")&" 倍：已簽 B300 合約換算約 "&TEXT(p_claim_lo,"0.0")&"–"&TEXT(p_claim_hi,"0.0")&" 百萬美元／MW·年，模型 FY2030 為 "&TEXT(Thesis!I' + str(TR["B"]["pm"]) + ',"0.00")&"' + (f'；支撐現價需 "&TEXT(RevDCF!G{RVN["px"]},"0.00")&"（同為 FY2030）' if REV and REV["solves"][0]["value"] is not None else '') + f'。託管營收 FY2027 起每年約 "&TEXT(Thesis!F{TR["B"]["colo"]},"0")&" 百萬美元（NC-1 長約為主）。"'),
    ("rev13", (f'="反向 DCF 問的是：現價 "&TEXT(px,"0.00")&" 美元隱含每 MW 年收入要多少。其他條件不變，FY2030 基準情境每 MW 年收入 "&TEXT(RevDCF!G{RVN["px"]},"0.00")&" 百萬美元即可支撐現價，"&IF(RevDCF!I{RVN["px"]}<0,"比模型基準 "&TEXT(RevDCF!H{RVN["px"]},"0.00")&" 低 "&TEXT(-RevDCF!I{RVN["px"]},"0.0%"),"比模型基準 "&TEXT(RevDCF!H{RVN["px"]},"0.00")&" 高 "&TEXT(RevDCF!I{RVN["px"]},"0.0%"))&"。公司合約換算中點 "&TEXT(RevDCF!J{RVN["px"]},"0.00")&"，"&IF(RevDCF!J{RVN["px"]}>RevDCF!G{RVN["px"]},"高於支撐現價所需。","不高於支撐現價所需：現價已隱含合約價格持續。")' if REV and "px" in RVN and REV["solves"][0]["value"] is not None else '="反向 DCF：在 k 調整係數 0.5–3.0 內無解（見下表）。"')),
    ("pl13", '="口徑：GAAP。WhiteFiber 財報的損益表與每股盈餘以 GAAP 為主，公司另揭露的調整後指標是調整後 EBITDA，因此本表用 GAAP。營業成本不含折舊（折舊攤銷列在營業費用），所以毛利率偏高，營業利益率才是可比的獲利指標。FY2023 找不到（2025 年分拆上市，10-K 只含 2024–2025），FY2024–FY2025 為 10-K 實績，FY2026 為前三季實績加第四季模型，FY2027 起為基準情境模型；預測期其他收入設為零。前三季淨損含誘導轉換費用 "&TEXT(80.767,"0.0")&" 百萬美元（2031 可轉債提前換回）。"'),
    ("plb13", '="FY2026 欄的淨利含前三季實際所得稅 "&TEXT(PL!K' + str(PLS["tax"]) + ',"0.000")&" 與其他收入 "&TEXT(PL!K' + str(PLS["oth"]) + ',"0.000")&"；本表的稅前利益與所得稅列只含模型項，所以 FY2026 稅前 − 所得稅 ≠ 淨利。完整對照見本章標準損益表。"'),
    ("drv13", "驅動主線：雲端在役 MW（情境路徑；2026 年為假設）× 計費比例＝計費 MW；計費 MW × 每 MW 年收入（Tokenomics 每 MW 持有成本 × 定價倍數）＝雲端營收；營運費用與第三方機房租金依在役 MW 計，資本支出＝新增 MW × 每 MW 設備成本＋託管資本支出，折舊依會計年限攤提；託管與其他營收取 NC-1 合約與 MTL 路徑。下方先看分部營收推導，再看每 MW 年收入與 backlog 兩張圖，參數明細收在摺疊。"),
    ("seg13", '="分部營收推導：雲端服務分部依合約型態拆成長約與隨需。長約每 MW 年收入＝持有成本 × "&TEXT(k_long,"0.000")&"（IREN–Microsoft 長約類比），隨需＝持有成本 × "&TEXT(k_spot,"0.000")&"（H100 現貨指數），隨需占計費產能 "&TEXT(od_share,"0%")&"（假設）；兩者各乘上計費 MW 後加總＝雲端分部營收；託管與其他另列於分部表。FY2026 只拆第四季，前三季為未拆分實績。"'),
    ("px13", '="現價（"&TEXT(px_date,"yyyy-mm-dd")&" 收盤）"'),
    ("fund13", '="收支模組：基準情境 FY2026–FY2030 自由現金流合計 "&TEXT(Fund!D' + str(FS["gate"]) + ',"0.0")&" 百萬美元（已含客戶預付流入，預付覆蓋資本支出 "&TEXT(prepay_cov,"0%")&"）；缺口依序以現金、投資級債（借款不超過 EBITDA 的 "&TEXT(lev_max,"0.0")&" 倍）、新股（每年不超過市值 "&TEXT(eq_cap,"0%")&"）、高息債補足。"'),
]
for k, t in L13:
    RD.cell(r, 1, t); RL[k] = r; r += 1
r += 1
RD.cell(r, 1, "v0.4.0（2026-11-15）：依 CoreWeave v5.0.3 新雲樣板（經 Nebius v0.6.0 分部版）重建（舊 JS 雙引擎 v0.2.1 移 whitefiber/legacy_v02/）＋滾入 2026Q3 實績（10-Q）＋版面 v1.3＋三層協調 Coord 頁。金額改為百萬美元；估值改為 DCF 50%＋Comparable 50%（新雲同業 EV/EBITDA，不再分部加總）、評等五級並套收支關卡；Tokenomics 快照由 v19.0 改 v22.0。Coord：Wattonomics 的 WhiteFiber 路徑讀舊版個股模型（同源），營收、資本支出標參考（同源校準），在役電力無可比實績。v0.3 版號由 open PR #451 占用。新舊目標價與評等差異見 whitefiber/README.md v0.4.0 節。"); RL["v060"] = r; r += 1
RD.cell(r, 1, "v0.4.1（2026-11-15）：Andy 2026-11-15 決定 Comparable 改為分部加總：託管分部（NC-1 Nscale 10 年長約、MTL、代管服務、其他）FY2028 EBITDA 乘上市資料中心同業（Digital Realty、Equinix）前瞻 EV/EBITDA 中位數，雲端分部乘現行新雲同業中位數；營業費用依營收比例分攤到兩分部；DCF 不變。新增 Sens：託管倍數折價（小型、單一主要承租人、非 REIT）。新舊目標價與評等見 whitefiber/README.md v0.4.1 節。"); RL["v041"] = r; r += 1
RD.cell(r, 1, "v0.4.2（2026-11-15）：Wattonomics 快照由 v0.3.2 改讀 v0.3.5（追公式改讀 whitefiber/data/wattonomics_trace_v0.3.5.json）。層級個股模型回寫資本支出 CapexAIModel 2026 由舊版 v0.2 的 0.473 重匯為本公司 v0.4.1 的 0.845（十億美元，同源，只列參考）；WhiteFiber 沒有付款年資本支出列，Gap 主讀法仍為 GW 推算，玩家 MW 路徑（mw_wyfi，舊版 v0.2）與 CapexAI 不變。Wattonomics 不進估值，基準輸入、目標價與評等不變。"); RL["v042"] = r; r += 1
RD.cell(r, 1, "v0.4.3（2026-11-16）：現價與同業報價改 2026-11-13 收盤（工作單 px1113；Yahoo Finance 日線，markets/invest_map/data/closes_20261020_20261113.csv）。現價 9.21（2026-11-10）→ 9.43；雲端同業報價：CoreWeave 60.77 → 53.13、Nebius 154.69 → 115.88、Oracle 128.38 → 125.35（均 2026-11-11 → 11-13；EV＝市值＋淨負債，依公式重算）。IREN 20.84（2026-11-11）→ 18.17（2026-11-13 收盤；EV＝股價 × 股數＋淨負債，公式重算，EV 15,530 → 14,397）；託管同業依報價日收盤換 11-13（價格檔已補 9-28～10-19）：DLR 10-07 收盤 180.47 → 179.99（EV 86,198 → 85,969，倍數 20.42 → 20.36）、EQIX 10-02 收盤 1,025.72 → 1,053.79（EV 121,482 → 124,806，倍數 21.06 → 21.63）；模型無 DLR／EQIX 市值，採倍數 × 價格比近似（Derived），淨負債不隨股價變。價格檔補齊後原地修改，不升版號；無未換價同業。雲端同業前瞻 EV/EBITDA 中位數 6.74 → 6.08（IREN 5.62 → 5.21、CoreWeave 5.01 → 4.73、Nebius 8.88 → 6.94、Oracle 7.86 → 7.71）；託管同業中位數 20.74 → 21.00；Comparable 目標價 27.59 → 26.62；加權目標價 13.79 → 13.31；上檔空間 +49.8% → +41.1%；估值評等 買進 → 買進，投資建議 收支不足 → 收支不足。rf、共識目標價（2026-11-10 對照值）、共識 EBITDA、淨負債、股數、營運假設與情境權重不變。Sens、RevDCF 快照已重跑；拆解沿用。")
RD.column_dimensions["A"].width = 180

# 列號對照
ROWMAP.update({"TR": TR, "SR": SR, "PR": PR_, "CR": CR_, "FS": FS, "FR": FR, "FUH": FUH, "RPB": RPB, "ST": ST, "D": D_, "LSB": LSB, "VR": VR, "DC": DC, "CO": CO, "CON": CON,
               "PEERS": [P0, P1], "PEERH": PEERH, "RVN": RVN, "RCR": RCR, "SENS": SENS_R, "GRID": list(GRID), "SNAP": SNAPR, "RL": RL, "SUM": list(SUM_ROWS), "CHK": [CHK_SEG, CHK_SUB],
               "QR": QR, "IR": IR, "SEGD": SEGD, "PLS": PLS, "CFC": CFC, "IN_LAST": IN_LAST, "GOV": {"fp": FPR, "fb": FBR, "tb": TBR, "st": STR, "bid": BIDR}, "BUILD_ID": BUILD_ID, "COORD": COORD_ROWS})
if os.environ.get("WYFI_ROWMAP"):
    json.dump(ROWMAP, open(os.environ["WYFI_ROWMAP"], "w", encoding="utf-8"), ensure_ascii=False, indent=0)
wb.save(sys.argv[1])
print("saved", sys.argv[1], BUILD_ID)
