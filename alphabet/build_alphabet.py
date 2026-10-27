# Alphabet（GOOGL）個股模型 v0.3：命題模組＋官方營收分部 → 損益表（docs/plan/個股模型原則_命題加營收分部.md）
# 工作單：mag/docs/workorders/20261027_MAG_v0.3_個股模型重建_分部架構.md（第 1、2、3「Alphabet」、4、5、6 節）
# 用法：python3 alphabet/build_alphabet.py raw.xlsx [--sens sens.json]
#       再以 LibreOffice headless 重算（alphabet/make_v03.sh 一次跑完：建置 → 敏感度重算 → 成品 → 檢查）。
# 金額：十億美元（US$B）；每股：美元；股數：十億股；MW＝IT 關鍵電力 MW。日曆年＝會計年度。
# 欄位（模型頁共用）：C FY2024A、D FY2025A、E 1H26A、F 2H26E、G FY2026E（＝E＋F）、H–M FY2027–FY2032。
import sys, os, json, argparse
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.workbook.defined_name import DefinedName

HERE = os.path.dirname(os.path.abspath(__file__))
ap = argparse.ArgumentParser()
ap.add_argument("out")
ap.add_argument("--sens", help="敏感度重算結果（make_v03.sh 產生）；省略時 Sens 頁結果欄留空")
ap.add_argument("--set", action="append", default=[], help="覆寫 In 頁輸入：鍵=值（敏感度重算用）")
A = ap.parse_args()
OVR = {}
for s in A.set:
    k, v = s.split("=", 1); OVR[k] = float(v)

BLUE = Font(color="0000FF"); BOLD = Font(bold=True); RED = Font(color="C00000", bold=True)
HDR = PatternFill("solid", fgColor="DDEBF7"); KEY = PatternFill("solid", fgColor="FFF2CC"); SUB = PatternFill("solid", fgColor="E2EFDA")
wb = openpyxl.Workbook()

FACTS_F = os.path.join(HERE, "data", "alphabet_facts_20261008.json")
FD = json.load(open(FACTS_F, encoding="utf-8"))
F = {it["id"]: it for g in FD["groups"].values() for it in g}
TK = json.load(open(os.path.join(HERE, "data", "tokenomics_snapshot_v21.4.json"), encoding="utf-8"))
GEN = {x["code"]: x["name"] for x in TK["source"]["generations"]}
def tk(name, code=None):
    v = TK["items"][name]["values"]
    return v if code is None else v[GEN[code]]["基準"]

def hdr(ws, row, labels, c0=1):
    for j, h in enumerate(labels, c0):
        c = ws.cell(row, j, h); c.font = BOLD; c.fill = HDR

def dn(name, ref):
    wb.defined_names[name] = DefinedName(name, attr_text=ref)

YK = ["24", "25", "1H", "2H", "26", 27, 28, 29, 30, 31, 32]
COL = dict(zip(YK, "CDEFGHIJKLM"))
YH = {"24": "FY2024A", "25": "FY2025A", "1H": "1H26A", "2H": "2H26E", "26": "FY2026E", 27: "FY2027E", 28: "FY2028E", 29: "FY2029E",
      30: "FY2030E", 31: "FY2031E", 32: "FY2032E"}
FC = ["2H", 27, 28, 29, 30, 31, 32]           # 預測期間
ANN = [27, 28, 29, 30, 31, 32]
LEN = {p: (0.5 if p in ("1H", "2H") else 1) for p in YK}
PY = {"2H": 2026, 27: 2027, 28: 2028, 29: 2029, 30: 2030, 31: 2031, 32: 2032}
def prevp(p):                                   # 前一期（流量年增用 FY；存量用前一期期末）
    return {"2H": "1H", 27: "2H", 28: 27, 29: 28, 30: 29, 31: 30, 32: 31}[p]
def prevfy(p):                                  # 年度流量的前一年度欄（FY2027 → FY2026）
    return {27: "26", 28: 27, 29: 28, 30: 29, 31: 30, 32: 31}[p]
MODEL_HDR = ["項目", "單位"] + [YH[k] for k in YK] + ["說明"]
NOTE_C = 14   # N 欄

def put(ws, r, label, unit, vals, fmt="#,##0.00", note=None, bold=False, fill=None, font=None):
    ws.cell(r, 1, label); ws.cell(r, 2, unit)
    if bold: ws.cell(r, 1).font = BOLD
    for p, v in vals.items():
        c = ws[f"{COL[p]}{r}"] if p in COL else ws[f"{p}{r}"]
        c.value = v; c.number_format = fmt
        if not (isinstance(v, str) and v.startswith("=")) and isinstance(v, (int, float)):
            c.font = BLUE
        if font: c.font = font
    if note: ws.cell(r, NOTE_C, note)
    if fill:
        for cc in range(1, NOTE_C + 1): ws.cell(r, cc).fill = fill

def widths(ws, a=58, b=10, mid=11, note=90, ncols=11):
    ws.column_dimensions["A"].width = a; ws.column_dimensions["B"].width = b
    for i in range(ncols): ws.column_dimensions[chr(ord("C") + i)].width = mid
    ws.column_dimensions["N"].width = note

# =====================================================================================================
# README（內容在最後補；先建頁以固定順序）
RD = wb.active; RD.title = "README"
IN = wb.create_sheet("In"); Q = wb.create_sheet("Q"); TH = wb.create_sheet("Thesis"); SG = wb.create_sheet("Seg")
PL = wb.create_sheet("PL"); CF = wb.create_sheet("CF"); V = wb.create_sheet("Val"); RC = wb.create_sheet("Recon")
SE = wb.create_sheet("Sens"); SO = wb.create_sheet("Sources")

# =====================================================================================================
# In：輸入（一列一鍵；E＝基準，D＝低，F＝高）
IN["A1"] = "In — 輸入（藍字；基準欄 E，低 D、高 F）。金額 US$B；每股 US$；股數十億股；MW＝IT MW。標記：Verified／Interested-party／Analogy／Assumed／Derived（Analogy、Assumed 一律給低／高區間）"
IN["A2"] = "「沿用 v0.2.2」＝MAG v0.2.2 已核准的參數（原出處：alphabet/legacy_v0.2/company.json、kspec.py、data/kevidence_alphabet_20261008.json、data/tokenomics_snapshot_v21.4.json；對照表 r1 C1–C23）。Tokenomics 維持 v21.4。黃底＝判斷題（Andy 可改）"
hdr(IN, 3, ["鍵", "項目", "單位", "低", "基準", "高", "標記", "來源", "說明"])
ROWS = []
def sec(t): ROWS.append(("sec", t))
def inp(key, item, unit, lo, base, hi, tag, src, note="", judg=False):
    ROWS.append((key, item, unit, lo, base, hi, tag, src, note, judg))
def _fid(i):
    if i in F: return i
    m = sorted(k for k in F if k.startswith(i + "."))
    assert m, i
    return m[-1]
def fv(i): return F[_fid(i)]["value"]
def fs(i): i = _fid(i); it = F[i]; return f"{it['source']}（{it['docDate']}；事實總帳 {i}）"

V22 = "沿用 v0.2.2（legacy_v0.2/company.json）"
sec("A. 現價、股數與資產負債（評價日 2026-06-30，2026Q2 10-Q）")
inp("px", "現價（2026-10-26 收盤）", "US$", None, 316.80, None, "Verified",
    "StockAnalysis 歷史股價頁 https://stockanalysis.com/stocks/googl/history/（2026-10-27 讀取：10-26 收盤 316.80、10-23 收盤 320.21）",
    "第二來源：markets/invest_map/data/prices.csv 2026-10-23 收盤 320.21（stockscan、bullfincher）與同頁 10-23 值相同。v0.2.2 現價 350.50（2026-10-07）")
inp("sh_out", "流通股數（2026-07-15，10-Q 封面）", "十億股", None, fv("sh.out.20260715") / 1000, None, "Verified", fs("sh.out.20260715"))
inp("sh_dil_q2", "2026Q2 稀釋加權股數", "十億股", None, fv("sh.dilW.2026q2") / 1000, None, "Interested-party", fs("sh.dilW.2026q2"))
inp("sh_basic_q2", "2026Q2 基本加權股數", "十億股", None, 12.151, None, "Interested-party", "Alphabet 2026Q2 10-Q（2026-07-23）；legacy_v0.2/company.json latestQuarter.basicWaso",
    "2026-06 股權發行（6/4）只在 Q2 部分期間加權；期末稀釋股數＝Q2 稀釋加權＋(流通股 − Q2 基本加權)")
inp("conv_sh", "強制轉換特別股 2029-05-15 轉換股數（上限）", "十億股", 0.0436, 0.0546, 0.0546, "Interested-party",
    "8-K 2026-06-05：A 系列每股轉 2.2520–2.8160 股、B 系列 2.2740–2.8420 股 → 43.6–54.6M 股；" + V22 + " debt.convertibles", "取稀釋最大（保守）；評價時直接加入股數，特別股股利照付至轉換")
inp("cash_0", "現金＋債權類有價證券（2026-06-30）", "US$B", None, 155.411, None, "Derived",
    "10-Q 2026Q2：現金 55.911＋有價證券 186.563 − 有價權益證券 87.063（事實總帳 bs.cash、bs.mktsec、bs.mktEq）", V22 + " defaults.cash（權益證券改列持股）")
inp("cash_ye25", "現金＋債權類有價證券（2025-12-31）", "US$B", None, 120.53, None, "Derived", "10-K 2025：30.708＋96.135 − 6.313（" + V22 + " ytdActual.cash1231）")
inp("debt_0", "債務面額（2026-06-30）", "US$B", None, fv("debt.face.2026q2") / 1000 if "debt.face.2026q2" in F else 101.085, None, "Interested-party", "10-Q 2026Q2 附註 6（事實總帳 debt.face）")
inp("fines", "應計罰款與和解（類債）", "US$B", None, 17.356, None, "Interested-party", "10-Q 2026Q2（事實總帳 bs.accruedFines）", "2026-07 已付 Android 案 5.2B（仍以 6/30 餘額扣除，與 v0.2.2 同）")
inp("nci", "非控制權益（類債）", "US$B", None, 7.1, None, "Interested-party", "10-Q 2026Q2（事實總帳 bs.nci；Waymo 等少數股權）")

sec("B. 持股（分部加總與 DCF 皆加入；非上市 × (1 − 流動性折價)；上市不折價）")
inp("ant_val", "Anthropic 最近一輪投後估值（Series H）", "US$B", None, 965.0, None, "Interested-party", fs("rp.anthropic.val.seriesH"))
inp("ant_stake", "Anthropic 持股比例", "%", 0.11, 0.125, 0.14, "Assumed", "一手不存在：2025 法院文件約 14%；非上市證券帳面 124.3B ÷ 965B 隱含上限 12.9%（" + V22 + " valuation.holdings）")
inp("illiq", "非上市持股流動性折價", "%", 0.0, 0.2, 0.4, "Assumed", "工作單第 2 節 Val 0.2（1 − 流動性折價 20%）；對照表 r1 第 5 節第 9 條", judg=True)
inp("spacex", "SpaceX 持股公允價值（上市、限售）", "US$B", None, 94.1, None, "Interested-party", "10-Q 2026Q2 附註 3：短期 80.0＋長期限售 14.1（事實總帳 rp.spacex.holding）", "評價日公允價值；之後股價變動未更新（資料缺口）")
inp("oth_eq", "其他有價權益證券", "US$B", None, 7.089, None, "Derived", "有價權益證券 87.063＋14.126 − SpaceX 94.1（" + V22 + "）")

sec("C. 折現率與目標價規則")
inp("rf", "無風險利率（10 年美債，2026-10-07）", "%", None, 0.0528, None, "Verified", fs("val.rf.10y"), V22)
inp("beta", "β（獨立來源平均）", "x", 1.21, 1.28845, 1.3669, "Derived", "StockAnalysis 5 年月 1.21（val.beta.sa）、GuruFocus 3 年 1.3669（val.beta.gurufocus）平均（C13）", V22)
inp("erp", "股權風險溢酬", "%", 0.04, 0.05, 0.06, "Assumed", V22 + "（val.erp）")
inp("kd", "稅前債務成本（2026-08 新發 2066 票據殖利率）", "%", None, 0.06527, None, "Verified", fs("val.kd.2066"), V22)
inp("tax_w", "WACC 用稅率（2025 有效稅率）", "%", None, 0.168, None, "Derived", "10-K 2025：所得稅 26.656 ÷ 稅前 158.826", V22)
inp("g_tv", "永續成長率（FY2032 後）", "%", 0.02, 0.03, 0.04, "Assumed", "工作單第 2 節 Val 0.1（g 3%）", judg=True)
inp("ronic", "永續期新投資報酬率（終值價值驅動公式）", "%", None, 0.12, 0.20, "Assumed",
    "判斷：終值＝NOPAT(FY2033) ×(1 − g ÷ RONIC) ÷ (WACC − g)；基準 12%（略高於 WACC，永續期新投資只賺略高於資金成本）", "替代：以 FY2032 無槓桿自由現金流 ×(1＋g) 計終值（資本支出仍在擴張期，終值偏低，Sens 列）", judg=True)
inp("tp_wA", "情境權重 A", "%", None, 0.25, None, "Assumed", "工作單第 2 節 Val 0.3（A 25%／B 50%／C 25%）", judg=True)
inp("tp_wB", "情境權重 B", "%", None, 0.50, None, "Assumed", "同上", judg=True)
inp("tp_wC", "情境權重 C", "%", None, 0.25, None, "Assumed", "同上", judg=True)
inp("w_dcf", "目標價中 DCF 的權重（其餘＝分部加總 SOTP）", "%", 0.3, 0.5, 0.7, "Assumed", "工作單第 2 節 Val 0.3（DCF 50%＋分部加總 50%；與 AI半導體「DCF＋Comparable 各半」一致）", judg=True)
inp("t_tgt", "目標時點距評價日（2026-06-30 → 2027-10-27）", "年", None, 1.3233, None, "Derived", "(2027-10-27 − 2026-06-30) ÷ 365＝483 天", "目標價約 12 個月後（今天 2026-10-27 起算）")
inp("w_ntm27", "目標時點 NTM 中 FY2027 的權重", "%", None, 0.1781, None, "Derived", "(2027-12-31 − 2027-10-27) ÷ 365＝65 天；FY2028 權重＝1 − 本值", "分部加總用目標時點的未來 12 個月 EBITDA")
inp("cons_tp", "市場共識目標價（平均）", "US$", None, 429.12, None, "Interested-party",
    "StockAnalysis 預測頁 https://stockanalysis.com/stocks/googl/forecast/（2026-10-26 更新；61 位分析師；中位數 430、低 340、高 500）", "v0.2.2 時 429.47（2026-10-08）；賣方有承銷關係，只作對照")

inp("mb_ttm", "分部利益率起點改用近四季（1＝是；0＝1H26）", "x", None, 0, 1, "Assumed", "判斷：預設 1H26（利潤率擴張中、與評價日機隊同時點）；1＝v0.2.2 的近四季口徑", "", judg=True)
inp("rd2", "讀法 2：自用 AI 增量投資價值中性（1＝評價時排除自用 AI 增量成本與資本支出；0＝讀法 1：照計成本）", "x", None, 0, 1, "Assumed",
    "對照表 r1 C16：讀法 1（自用 AI 只計成本，v0.2.2 主值）／讀法 2（自用 AI 價值中性）；主值待 Andy 決定", "只影響評價（DCF、SOTP）；損益表與現金流照實計入自用 AI 成本", judg=True)
inp("tv_alt", "終值改用 FY2032 無槓桿自由現金流 ×(1＋g)（1＝是；Sens 用）", "x", None, 0, 1, "Assumed", "判斷：基準用價值驅動公式（ronic）；1＝替代口徑", "", judg=True)

sec("C2. 分部加總（SOTP）倍數：同業 NTM EV/EBITDA（沿用 v0.2.2 共用同業檔，基準日 2026-10-07；中位數在 Val 頁計算）")
PEERS = {"ads": ("數位廣告（Search、YouTube、Network、Alphabet 層級費用）", "data/peers_ads_20261009.json（C18；MAG 三家共用）"),
         "subs": ("訂閱影音（Subscriptions, platforms, and devices）", "legacy_v0.2/company.json valuation.segmentMultiples.subs（NFLX 為 NTM；SPOT、DIS 為 TTM，amazon/data 事實總帳）"),
         "cloud": ("非 AI 雲端／軟體（Cloud 其他）", "data/peers_cloud_20261008.json（C12；MAG 三家共用）")}
_pa = json.load(open(os.path.join(HERE, "data", "peers_ads_20261009.json"), encoding="utf-8"))["peers"]
_pc = json.load(open(os.path.join(HERE, "data", "peers_cloud_20261008.json"), encoding="utf-8"))["peers"]
PEERL = {"ads": [(p["ticker"], p["name"], p["ntmEvEbitda"]) for p in _pa],
         "subs": [("NFLX", "Netflix", 15.41), ("SPOT", "Spotify", 31.97), ("DIS", "Disney", 10.58)],
         "cloud": [(p["ticker"], p["name"], p["ntmEvEbitda"]) for p in _pc]}
for grp, lst in PEERL.items():
    for tkr, nm, x in lst:
        inp(f"pm_{grp}_{tkr.lower()}", f"{PEERS[grp][0]}：{nm}（{tkr}）NTM EV/EBITDA", "x", None, x, None, "Derived", PEERS[grp][1], "EV（StockAnalysis）÷ NTM EBITDA（MarketScreener 共識 FY1／FY2 時間加權）")
inp("mult_ai", "Cloud AI、TPU 系統銷售、Other Bets 的 EV/EBITDA", "x", 4.0, 6.0, 10.0, "Assumed", V22 + " valuation.evEbitda（AI 基礎設施 6×；CRWV TTM 25×、ORCL NTM 9.5× 為對照，事實總帳 val.peer.aiInfra）",
    "Other Bets EBITDA 為負 → 虧損 × 倍數（保守；Waymo 估值未找到）", judg=True)
inp("mult_ads_own", "Alphabet 自身 NTM EV/EBITDA（Sens 用：廣告線改用自身倍數）", "x", None, 14.68, None, "Derived", "事實總帳 val.googl.ownNtm（2026-10-07）")
inp("ads_own", "廣告線倍數改用自身倍數（1＝是；Sens 用）", "x", None, 0, 1, "Assumed", "Sens 用開關")

sec("D. 命題模組：AI 容量（沿用 v0.2.2；總 AI MW＝對外 MW ÷ 對外比例）")
inp("ext_sh", "對外比例（Google Cloud 客戶 MW ÷ AI 總 MW）", "%", 0.25, 0.30, 0.40, "Assumed", fs("mw.extShare.2026q2") if "mw.extShare.2026q2" in F else "事實總帳 mw.extShare（v0.1a 區間 25%–40%）",
    V22 + " capexModel.extShare；C23：公司未揭露，是最大不確定（Sens 另列 ±20pt）")
inp("mw_tot0", "AI 在役 IT MW（2026-06-30，自用＋對外）", "MW", 5012, 6337, 7891, "Derived", fs("mw.ai.it.2026q2"),
    "Epoch AI 晶片功耗 3,841 MW × IT/TDP 1.65；對外 1,901＝6,337 × 30%（" + V22 + " defaults.billableOpen）")
inp("mw_ye25", "AI 在役 IT MW（2025-12-31）", "MW", None, 4873, None, "Derived", fs("mw.ai.it.epoch.2025q4"), "對外 1,462（" + V22 + " defaults.mwYearEnd 2025）÷ 30%")
inp("mw_ye24", "AI 在役 IT MW（2024-12-31；汰換用）", "MW", None, 2143.3, None, "Derived", V22 + " defaults.mwYearEnd 2024 對外 643 ÷ 30%")
inp("mw_ye23", "AI 在役 IT MW（2023-12-31；汰換用）", "MW", None, 636.7, None, "Derived", V22 + " defaults.mwYearEnd 2023 對外 191 ÷ 30%")
inp("pace", "AI 總 MW 每年新增（三情境相同）", "MW/年", 2833.3, 4159.67, 6239.67, "Derived",
    V22 + "：對外速度 1,247.9 MW／年（C10：2026 資本支出指引中點 200B 解出，scripts/calib_pace.js）÷ 30%",
    "低＝v0.2.2 保守 850 ÷ 30%（已公開在建）、高＝積極 1,871.9 ÷ 30%（1.5 × 基準）；本版不作情境軸（預設 1，見報告第 5 節），只列 Sens")
inp("mix0_h", "期初機隊：Hopper 錨（TPU v4／v5、H100／H200）", "%", None, 0.25, None, "Assumed", V22 + " pricing.chips（TPU v4／v5 15%＋Hopper 10%）", "C2：TPU 以同期 NVIDIA 世代代理")
inp("mix0_b2", "期初機隊：GB200 錨（Trillium v6e、GB200／B200）", "%", None, 0.45, None, "Assumed", V22 + "（Trillium 35%＋GB200 10%）", "C2")
inp("mix0_b3", "期初機隊：GB300 錨（Ironwood）", "%", None, 0.30, None, "Assumed", V22 + "（Ironwood 30%）", "C2")
inp("am_27", "2027 新增中 GB300 錨占比（Ironwood＋GB300；其餘 VR200 錨）", "%", None, 0.5, None, "Assumed", V22 + "（2027 Ironwood 40%＋GB300 10%）", "2026 新增全為 GB300 錨（Ironwood 75%＋GB300 25%）；2028 起全為 VR200 錨（下一代 TPU＋VR200）")
inp("tpu_sh", "TPU 占 AI MW 比例", "%", None, 0.75, None, "Derived", "Epoch 2025 年新增功耗 TPU 約 77%（" + V22 + " pricing.mixNote）", "只在 TPU 係數不等於 1 時有作用")
inp("tpu_f", "TPU 每 MW 收入錨係數（相對同期 NVIDIA 世代）", "x", 0.7, 1.0, 1.0, "Assumed", "C2：基準 × 1.0、敏感度 × 0.7（" + V22 + " pricing.customFactor）", "Tokenomics v21.4 NonNV（IF_NonNV_Hold_TPU 0.8）只登錄、不切換（工作單：Tokenomics 本版維持 v21.4）")
inp("tpu_cf", "TPU 每 MW IT 資本支出係數", "x", 0.7, 1.0, 1.0, "Assumed", V22 + " pricing.customCapexFactor", "雙邊敏感度時與 tpu_f 同乘 0.7")

sec("E. 命題模組：每 MW 價格（k＝證據價格 ÷ Tokenomics 同世代成本；沿用 v0.2.2 kspec.py）")
inp("kl_px", "長約證據價格（判斷倍數 0.75 × v5.27 VR200 12.762）", "US$m/MW-年", 7.6572, 9.5715, None, "Assumed",
    "legacy_v0.2/company.json pricing.kEvidence long.base／long.low（data/kevidence_alphabet_20261008.json kspecInputs）", "不是觀察到的市場價格；分母＝VR200 IF_HoldEcon")
inp("ks_px", "現貨證據價格（Silicon Data H100 指數）", "US$/GPU-hr", None, 2.82, 4.858, "Verified",
    "Silicon Data H100 Index（2026-10-06）https://www.silicondata.com/products/silicon-index/h100；高＝GCP a3-highgpu-8g 3 年預留 4.858（Interested-party）", "分母＝IF_GPUhrEcon｜H100；C1 三家一致")
inp("long_sh", "長約占比", "%", 0.4, 0.6, 0.8, "Assumed", V22 + " pricing.longShare（Cloud 積壓以多年期大客戶為主）")
for code, nm in (("H100", "h"), ("GB200", "b2"), ("GB300", "b3"), ("VR200", "vr")):
    inp(f"he_{nm}", f"IF_HoldEcon｜{code}（每 MW 年持有成本，經濟口徑）", "US$m/MW-年", tk("IF_HoldEcon")[GEN[code]]["低成本"], tk("IF_HoldEcon", code), tk("IF_HoldEcon")[GEN[code]]["高成本"], "Derived",
        f"data/tokenomics_snapshot_v21.4.json（{TK['source']['file']}，{TK['items']['IF_HoldEcon']['ref']}）", "Tokenomics WACC 10% 下的打平持有成本；收入端不乘 IF_Util（MW-year 計價）")
inp("gpuhr_h", "IF_GPUhrEcon｜H100", "US$/GPU-hr", None, tk("IF_GPUhrEcon", "H100"), None, "Derived", "data/tokenomics_snapshot_v21.4.json")
for code, nm in (("H100", "h"), ("GB200", "b2"), ("GB300", "b3"), ("VR200", "vr")):
    inp(f"opx_{nm}", f"IF_OpexGW｜{code}（每 MW 年營運費用，不含折舊）", "US$m/MW-年", None, tk("IF_OpexGW", code), None, "Derived", "data/tokenomics_snapshot_v21.4.json")
for code, nm in (("H100", "h"), ("GB200", "b2"), ("GB300", "b3"), ("VR200", "vr")):
    inp(f"cit_{nm}", f"IF_CapexIT｜{code}（每 MW IT 資本支出）", "US$m/MW", None, tk("IF_CapexIT", code), None, "Derived", "data/tokenomics_snapshot_v21.4.json")
inp("cfac", "IF_CapexFacility（每 MW 機房資本支出）", "US$m/MW", None, tk("IF_CapexFacility", "VR200"), None, "Derived", "data/tokenomics_snapshot_v21.4.json")
for code, nm in (("H100", "h"), ("GB200", "b2"), ("GB300", "b3"), ("VR200", "vr")):
    inp(f"cap_{nm}", f"IFW_RevGWFleet_Util｜{code}（上限檢查用，不得作預測）", "US$m/MW-年", None, tk("IFW_RevGWFleet_Util", code), None, "Derived", "data/tokenomics_snapshot_v21.4.json", "每 MW 收入 ÷ 本值 > 50% 示警")
inp("selfb", "自建機房比例（租用部分不計機房資本支出）", "%", 0.75, 0.85, 0.95, "Assumed", V22 + " capexModel.selfBuild")
inp("life_it", "AI IT 折舊年限", "年", 5, 6, 7, "Verified", "Tokenomics IF_DeprLifeIT 6 年；10-K 2025：伺服器與網通設備 6 年（事實總帳 ppe.serverLife）")
inp("life_fac", "機房折舊年限", "年", 15, 20, 25, "Derived", "Tokenomics IF_CapexFacility 12.67 ÷ IF_DeprFac 0.634＝20 年")
inp("life_na", "非 AI 資本支出折舊年限", "年", 8, 10, 15, "Assumed", V22 + " capexModel.nonAiLife（C17）")
inp("phi_0", "資本支出當期折舊比例（在役時點與半年慣例）", "x", 0.0, 0.25, 0.5, "Assumed", "判斷：支出平均於期中、約半年後上線 → 當期 25%；最後一年補足 75%", "AI 折舊進損益的時點（工作單第 2 節 Thesis）")
inp("da_base", "評價日在役資產年折舊（2026Q2 × 4）", "US$B/年", None, 28.416, None, "Derived", "現金流量表折舊 2026Q2 7.104 × 4（" + V22 + " capexModel.segDaRunRate）",
    "在役淨額 ≈ PP&E 321.2 − 在建 122.8＝198.4，÷ 28.4＝7.0 年 → 模型期內以直線法每年 28.4")
inp("cip_0", "在建工程（尚未在役，2026-06-30）", "US$B", None, fv("bs.ppe.notInService.2026q2") / 1000 if "bs.ppe.notInService.2026q2" in F else 122.814, None, "Interested-party", "10-Q 2026Q2 附註（事實總帳 bs.ppe.notInService）", "2026H2 起轉入在役；依 AI 資本支出結構分 IT／機房")
inp("cip_it", "在建工程中 IT（伺服器）占比", "%", 0.5, 0.6, 0.7, "Interested-party", fs("capex.mix.2026q2") if "capex.mix.2026q2" in F else "2026Q2 法說：約 60% 伺服器、40% 資料中心與網通（事實總帳 capex.mix）")
inp("capex_g", "2026 資本支出指引中點", "US$B", 195, 200, 205, "Interested-party", fs("guid.capex.2026.jul"), "首期資本支出下限：MAX(公式, 指引) − 上半年實際（C10）")
inp("capex_1h", "2026 上半年資本支出（購置不動產與設備）", "US$B", None, 80.598, None, "Interested-party", "財報新聞稿 Q1 35.674＋Q2 44.924")
inp("svc_sh", "自用 AI 增量成本中歸 Google Services 的比例（其餘＝Alphabet 層級：DeepMind 模型開發）", "%", 0.4, 0.6, 0.8, "Assumed",
    "10-K：技術基礎設施依使用量分攤到分部；2024 起 Google DeepMind 一般 AI 模型開發費用列 Alphabet-level activities", "服務端推論（搜尋 AI、YouTube、Gemini app）歸 Services；訓練歸 Alphabet 層級", judg=True)
inp("ai_mw_ttm", "近四季平均對外 AI MW（2025Q3–2026Q2）", "MW", None, 1447.9, None, "Derived", V22 + " legacyBiz cloudNonAi.aiMwTTM", "只用於拆 Cloud 其他的近四季年增（起始成長率）")
inp("ai_mw_pttm", "前四季平均對外 AI MW（2024Q3–2025Q2）", "MW", None, 657.3, None, "Derived", V22 + " legacyBiz cloudNonAi.aiMwPrevTTM")
for p, v in (("2H", 2.76), (27, 11.04), (28, 11.04), (29, 5.52)):
    inp(f"rent_{p}", f"SpaceX 算力租金（自用）{YH[p]}", "US$B", None, v, None, "Interested-party",
        "SpaceX FWP 2026-06-05：每月 9.2 億、2026-10～2029-06（事實總帳 rent.spacex；" + V22 + " leases.rentedCompute）", "列自用、計入營運成本（C8 d）；2026-12-31 後可 90 天通知解約" if p == "2H" else "")

sec("F. 營收分部（Seg 頁）：成長路徑與 Search 情境軸")
inp("conv_n", "成長率收斂年數（FY2026 起始 → 長期值）", "年", None, 4, None, "Assumed", "判斷：近四季年增在 FY2030 收斂到長期值（工作單：「近四季年增率 → 長期值」）")
inp("gl_search", "Google Search & other 長期成長（B 基準路徑）", "%", 0.05, 0.08, 0.11, "Assumed", V22 + " legacyBiz search.gLT", "近四季 +16.8%（Q 頁）；2015–2025 年均約 13%")
inp("gl_yt", "YouTube ads 長期成長", "%", 0.05, 0.08, 0.11, "Assumed", V22 + " legacyBiz youtube.gLT")
inp("gl_net", "Google Network 長期成長", "%", -0.03, 0.0, 0.03, "Assumed", V22 + " legacyBiz network.gLT", "近四季 −2.2%")
inp("gl_subs", "Google subscriptions, platforms, and devices 長期成長", "%", 0.07, 0.10, 0.13, "Assumed", V22 + " legacyBiz subs.gLT")
inp("gl_co", "Cloud 其他（非 AI＋Workspace＋Wiz）長期成長", "%", 0.05, 0.08, 0.12, "Assumed", V22 + " legacyBiz cloudNonAi.gLT")
inp("dA", "情境 A：查詢分流（負值＝AI 模式吸引查詢回流）FY2030 達到", "%", -0.10, -0.05, 0.0, "Assumed",
    "AI Overviews／AI Mode 擴大使用：公司稱 AI Overviews 帶動相關查詢增加（2025 法說，二手；本次未重新查核）", "線性：FY2026 為 0 → FY2030 全額，之後持平", judg=True)
inp("mA", "情境 A：每次查詢變現變動 FY2030 達到", "%", 0.05, 0.10, 0.15, "Assumed",
    "AI Mode／AI Overviews 廣告：公司稱 AI Overviews 變現率與傳統搜尋約略相同（2025Q1 法說，二手；本次未重新查核）；A 假設新廣告格式提高商業意圖查詢單價", "", judg=True)
inp("dC", "情境 C：查詢被 AI 助理分流 FY2030 達到", "%", 0.10, 0.20, 0.30, "Assumed",
    "Gartner 2024-02-19 預測：2026 年傳統搜尋引擎查詢量因 AI 聊天機器人減少 25%（第三方預測）；Apple Eddy Cue 2025-05-07 美國司法部案證詞：Safari 搜尋 2025-04 首次下滑（二手；本次未重新查核）",
    "反證：Google Search & other 近四季仍 +16.8%、2026Q2 年增 +16.8%（Q 頁）→ 分流至今未見於營收；C 是壓力情境", judg=True)
inp("mC", "情境 C：每次查詢變現變動 FY2030 達到", "%", -0.15, -0.10, -0.05, "Assumed",
    "AI 答案頁的廣告位少於十條藍色連結、點擊率下降（第三方點擊研究，二手；本次未重新查核）", "", judg=True)
inp("inc_m", "Search 情境差額的增量 EBITDA 率", "%", 0.6, 0.7, 0.8, "Assumed", "判斷：搜尋營收增減主要只連動 TAC（約占搜尋營收兩成，Q 頁 TAC ÷ 搜尋＋聯播網）→ 增量利潤率高於分部平均", "營收減少時利潤減少得更多（營運槓桿）", judg=True)
inp("tpu_2h", "TPU 系統銷售 2H26", "US$B", 0.0, 1.0, 2.0, "Assumed", V22 + " legacyBiz tpuSales（C8 b）；公司：2026 認列一小部分、大部分在 2027（事實總帳 guid.tpuRev）", "列 Google Cloud 子列（命題子列）")
inp("tpu_yr", "TPU 系統銷售 FY2027 起每年", "US$B", 6.5, 13.0, 19.5, "Interested-party", "Morgan Stanley：2027 約 50 萬顆、約 130 億（賣方；" + V22 + "）", "低／高 ±50% [Assumed]")
inp("tpu_m", "TPU 系統銷售 EBITDA 率", "%", 0.15, 0.25, 0.35, "Assumed", V22 + " legacyBiz tpuSales.m0")
for p, v in (("2H", 0.75), (27, 1.6), (28, 1.8), (29, 2.0), (30, 2.2), (31, 2.4), (32, 2.6)):
    inp(f"ob_rev_{p}", f"Other Bets 營收 {YH[p]}", "US$B", None, v, None, "Assumed", V22 + " legacyBiz otherBets.rev（FY2031–32 每年 +0.2 延伸）" if p == "2H" else "")
inp("ob_oi", "Other Bets 營業利益（每年；2H26 為一半）", "US$B", -11.0, -9.0, -7.0, "Assumed", "近四季 −8.942（Q 頁）；Waymo 擴張（2026-02 募資 16B，事實總帳 ob.waymo.funding）")
inp("abc_rate", "Alphabet 層級未分攤費用占營收（不含自用 AI 增量成本）", "%", 0.04, 0.0483, 0.055, "Derived", "2026Q2：5.789 ÷ 119.796（" + V22 + " legacyBiz.split.corpShareOfRevenue）")
inp("m_yt", "YouTube ads EBITDA 率（分部加總拆線用）", "%", 0.30, 0.40, 0.50, "Assumed", V22 + " legacyBiz.split.youtubeMargin", "Services 分部利益只揭露合計；Search＝分部殘差")
inp("m_net", "Google Network EBITDA 率", "%", 0.10, 0.15, 0.25, "Assumed", V22 + " legacyBiz.split.networkMargin")
inp("m_subs", "Subscriptions, platforms, and devices EBITDA 率", "%", 0.10, 0.20, 0.30, "Assumed", V22 + " legacyBiz.split.subsMargin")
inp("cx_svc", "非 AI 資本支出占營收：Google Services 各線", "%", 0.04, 0.06, 0.105, "Assumed", V22 + " legacyBiz.cxNote")
inp("cx_co", "非 AI 資本支出占營收：Cloud 其他", "%", 0.15, 0.20, 0.27, "Assumed", V22 + " legacyBiz.cxNote（比照 AWS 非 AI）")
inp("cx_ob", "非 AI 資本支出占營收：Other Bets", "%", None, 0.06, None, "Assumed", V22)

sec("G. 損益與現金流")
inp("tax", "有效稅率（預測期）", "%", 0.15, 0.168, 0.19, "Derived", "2025 有效稅率 16.8%（10-K）；2026H1 19.1% 含權益證券利益", V22 + " valuation.tax")
inp("tax_eq", "權益證券損益的稅率（EPS 不含權益證券損益用）", "%", None, 0.21, None, "Assumed", "美國聯邦稅率 21%（未實現利益之遞延所得稅）")
inp("r_cash", "現金收益率", "%", None, 0.038, None, "Derived", "2025 利息收入 ÷ 平均現金＋債權證券（Q 頁 4.3–4.6B／季 ÷ 約 110B）約 3.8–4.0%")
inp("r_debt", "債務利率", "%", None, 0.048, None, "Derived", "2026Q2 利息費用 1.278 × 4 ÷ 101.1＝5.1%；2026 新債加權票面 3.9%–5.3% → 4.8%")
inp("sbc_r", "股權報酬占營收", "%", None, 0.0631, None, "Derived", "近四季 SBC 28.147 ÷ 營收 445.866（" + V22 + " defaults.sbcRate）", "營運現金加回；DCF 不加回（視為真實成本）")
inp("dil", "股數年稀釋（股權報酬）", "%", 0.0, 0.01, 0.015, "Assumed", V22 + "（沿用模板每年 1%；代扣稅與 ATM 不建模）")
inp("wc_2h", "2H26 營運資金變動（流入）", "US$B", None, 8.49, None, "Derived", "2025 全年 +0.618 − 2025H1 −7.872（C14；" + V22 + " defaults.wcStub）")
inp("wc_k", "FY2027 起營運資金變動 ÷ 營收增量（流入）", "%", 0.0, 0.0117, 0.02, "Derived", "2025：+0.618 ÷ 營收增量 52.818（10-K）")
inp("dps", "普通股每股年股利（2026 起 0.22／季）", "US$", None, 0.88, None, "Interested-party", "事實總帳 div.dps（2026-04 起每季 0.22）")
inp("dps_g", "每股股利年成長", "%", None, 0.05, None, "Derived", "0.20（2024）→ 0.21（2025）→ 0.22（2026）（C19）")
inp("bb_t", "回購目標（FY2027 起每年，現金足夠時）", "US$B", 0, 17.003, 45.709, "Derived", "近四季實際 17.003（事實總帳 bb.paid.4q）；2026 暫停（C4）", "高＝2025 全年 45.709；現金不足時先減回購、再舉債")
inp("min_cash", "最低現金", "US$B", None, 50, None, "Assumed", V22 + " defaults.minCash")
for p, v in (("2H", 0.206), (27, 0.387), (28, 0.377), (29, 0.356), (30, 0.285), (31, 0.25), (32, 0.25)):
    inp(f"lease_{p}", f"融資租賃本金 {YH[p]}", "US$B", None, v, None, "Interested-party" if p not in (31, 32) else "Assumed",
        "10-Q 2026Q2 附註 4 到期表（" + V22 + " leases.financePayments）" if p == "2H" else ("FY2031 起 0.25／年" if p == 31 else ""))
for p, v in (("2H", 2.222), (27, 4.444), (28, 4.444), (29, 4.444), (30, 4.444), (31, 0.0), (32, 0.0)):
    inp(f"anth_{p}", f"Anthropic 條件式出資 {YH[p]}", "US$B", None, v, None, "Interested-party", "剩餘 20.0B 至 2030（" + V22 + " defaults.jvCommit）" if p == "2H" else "",
        "投資活動，不在自由現金流；換得持股（評價中性）" if p == "2H" else "")
for p, v in (("2H", 2.0), (27, 1.0), (28, 2.676), (29, 1.764), (30, 7.186), (31, 5.0), (32, 5.0)):
    inp(f"mat_{p}", f"債務到期 {YH[p]}", "US$B", None, v, None, "Interested-party" if p not in (31, 32) else "Assumed",
        "10-K 2025 附註 6 到期表＋信用額度（" + V22 + " debt.amortization）" if p == "2H" else ("FY2031 起 5.0／年（其後 86.5 未分年）" if p == 31 else ""))
inp("debt_new_2h", "2026-08 新發美元票據", "US$B", None, 25.0, None, "Interested-party", fs("debt.iss.202608"))
for p, v in (("2H", 0.535), (27, 1.203), (28, 1.203), (29, 0.602)):
    inp(f"pref_{p}", f"強制轉換特別股股利 {YH[p]}", "US$B", None, v, None, "Derived", "6.25% × 19.25B，每季 2/15、5/15、8/15、11/15 至 2029-05-15（" + V22 + " defaults.dividend.preferred）" if p == "2H" else "")

sec("H. 對照（Recon）")
inp("tp_v022", "v0.2.2 基準目標價（讀法 1）", "US$", None, 146.09, None, "Derived", "docs/reports/20261025_alphabet_v0.2.2.md；交接檔 v0.2.2", "讀法 2（自用 AI 價值中性）218.79；三情境皆賣出")
inp("tpw_v022", "v0.2.2 加權目標價（讀法 1，A 保守 175.08／B 146.09／C 積極 122.89 以 25／50／25 加權）", "US$", None, 147.54, None, "Derived", "交接檔 v0.2.2 第 1 節三情境目標價；本表以本版情境權重換算")
inp("px_v022", "v0.2.2 現價（2026-10-07）", "US$", None, 350.50, None, "Verified", fs("mkt.price.close.2026q2") if "mkt.price.close.2026q2" in F else "事實總帳 mkt.price.close（2026-10-07）")
CONS = {  # MarketScreener 共識（2026-10-26 頁面；S&P Capital IQ 系）
    "rev": (500.142, 626.817, 754.762), "ebit": (172.033, 220.967, 268.322), "eps": (20.60, 15.12, 18.21),
    "capex": (202.078, 308.024, 352.387), "fcf": (7.069, -37.207, 6.132), "ebitda": (233.956, 305.081, 393.759)}
for k, lab in (("rev", "營收"), ("ebit", "營業利益"), ("eps", "EPS（GAAP，含權益證券利益）"), ("capex", "資本支出"), ("fcf", "自由現金流"), ("ebitda", "EBITDA（共識口徑含 SBC 加回）")):
    for i, y in enumerate((26, 27, 28)):
        inp(f"c_{k}_{y}", f"共識 {lab} FY20{y}", "US$" if k == "eps" else "US$B", None, CONS[k][i], None, "Interested-party",
            "MarketScreener 共識（https://www.marketscreener.com/quote/stock/ALPHABET-INC-24203373/finances/，2026-10-26 頁面）" if i == 0 else "同上",
            "v0.2.2 時（2026-10-08）：營收 498.69／615.04／734.20" if (k == "rev" and i == 0) else "")
inp("c_eps26_sa", "共識 FY2026 EPS（non-GAAP，排除權益證券利益）", "US$", 11.38, 11.89, 12.43, "Interested-party", "StockAnalysis 預測頁（2026-10-26；52 位分析師）")

# 寫入 In
IN_ROW = {}
r = 4
for row in ROWS:
    if row[0] == "sec":
        IN.cell(r, 1, row[1]).font = BOLD; r += 1; continue
    key, item, unit, lo, base, hi, tag, src, note, judg = row
    if key in OVR: base = OVR[key]
    IN.cell(r, 1, key); IN.cell(r, 2, item); IN.cell(r, 3, unit)
    for j, v in ((4, lo), (5, base), (6, hi)):
        if v is not None:
            c = IN.cell(r, j, v)
            if not (isinstance(v, str) and v.startswith("=")): c.font = BLUE
            c.number_format = "0.0%" if unit == "%" else ("0.0000" if unit == "x" else "#,##0.000")
    IN.cell(r, 7, tag); IN.cell(r, 8, src); IN.cell(r, 9, note)
    if judg:
        for cc in range(1, 10): IN.cell(r, cc).fill = KEY
    dn(key, f"In!$E${r}")
    if lo is not None: dn(key + "_lo", f"In!$D${r}")
    if hi is not None: dn(key + "_hi", f"In!$F${r}")
    IN_ROW[key] = r
    r += 1
for col, w in zip("ABCDEFGHI", [14, 64, 12, 10, 11, 10, 16, 90, 80]): IN.column_dimensions[col].width = w
IN.freeze_panes = "D4"

# =====================================================================================================
# Q：實績（8 季＋FY2024、FY2025）
Q["A1"] = "Q — 實績（US$B；EPS 美元；股數十億股）。資料取自 data/alphabet_facts_20261008.json（MAG v0.1a 起的 SEC 一手資料；季度數＝財報新聞稿 [Interested-party，10-Q 核閱]，年度＝10-K [Verified]）"
Q["A2"] = "核對列（實績必須為 0）：產品線合計＝Google Services、分部合計＝合併營收、分部營業利益合計＝合併營業利益（Alphabet-level activities 含避險損益營收與未分攤費用）、營收 − 成本費用＝營業利益、FY2025 四季合計＝年報"
QCOLS = ["2024q3", "2024q4", "2025q1", "2025q2", "2025q3", "2025q4", "2026q1", "2026q2", "fy2024", "fy2025"]
QC = dict(zip(QCOLS, "BCDEFGHIJK"))
hdr(Q, 3, ["項目"] + [c.upper().replace("FY", "FY") for c in QCOLS] + ["1H26", "近四季 TTM", "前四季", "標記", "來源"])
QR = {}
SRCQ = ("季：Alphabet 財報新聞稿（8-K 附件 99.1，各季；事實總帳 is.*.2024q3–2026q2）；年：Form 10-K FY2024、FY2025（經查核）；"
        "例：2026Q2 https://www.sec.gov/Archives/edgar/data/1652044/000165204426000066/googexhibit991q22026.htm")
def qrow(r, key, label, fid=None, scale=1000, neg=False, formula=None, fmt="#,##0.000", h1="sum", ttm=True, tag="Verified（年）／Interested-party（季）", src=SRCQ, fy=True):
    Q.cell(r, 1, label); QR[key] = r
    for c in QCOLS:
        col = QC[c]
        if formula:
            Q[f"{col}{r}"] = "=" + formula(col)
        else:
            iid = f"{fid}.{c}"
            if iid not in F or (c.startswith("fy") and not fy):
                continue
            v = F[iid]["value"]
            Q[f"{col}{r}"] = (-v if neg else v) / scale
            Q[f"{col}{r}"].font = BLUE
        Q[f"{col}{r}"].number_format = fmt
    if h1 == "sum": Q[f"L{r}"] = f"=H{r}+I{r}"
    elif h1: Q[f"L{r}"] = h1
    if ttm:
        Q[f"M{r}"] = f"=SUM(F{r}:I{r})"; Q[f"N{r}"] = f"=SUM(B{r}:E{r})"
    for cc in "LMN": Q[f"{cc}{r}"].number_format = fmt
    Q.cell(r, 15, tag); Q.cell(r, 16, src)

r = 4
Q.cell(r, 1, "營收（官方分部與產品線）").font = BOLD; r += 1
qrow(r, "search", "Google Search & other", "is.rev.search"); r += 1
qrow(r, "yt", "YouTube ads", "is.rev.yt"); r += 1
qrow(r, "net", "Google Network", "is.rev.network"); r += 1
qrow(r, "subs", "Google subscriptions, platforms, and devices", "is.rev.subs"); r += 1
qrow(r, "gs", "Google Services（分部）", "is.rev.gs"); r += 1
qrow(r, "chk_gs", "檢查：產品線合計 − Google Services", formula=lambda c: f"{c}{QR['search']}+{c}{QR['yt']}+{c}{QR['net']}+{c}{QR['subs']}-{c}{QR['gs']}", fmt="0.000", tag="—", src="—"); r += 1
qrow(r, "cloud", "Google Cloud（分部）", "is.rev.cloud"); r += 1
qrow(r, "ob", "Other Bets（分部）", "is.rev.ob"); r += 1
qrow(r, "hedge", "Hedging gains (losses)（Alphabet-level）", "is.rev.hedge"); r += 1
qrow(r, "rev", "合併營收", "is.rev.total"); r += 1
qrow(r, "chk_rev", "檢查：分部合計＋避險 − 合併營收", formula=lambda c: f"{c}{QR['gs']}+{c}{QR['cloud']}+{c}{QR['ob']}+{c}{QR['hedge']}-{c}{QR['rev']}", fmt="0.000", tag="—", src="—"); r += 1
r += 1
Q.cell(r, 1, "分部營業利益").font = BOLD; r += 1
qrow(r, "oi_gs", "Google Services 營業利益", "is.ebit.gs"); r += 1
qrow(r, "oi_cloud", "Google Cloud 營業利益", "is.ebit.cloud"); r += 1
qrow(r, "oi_ob", "Other Bets 營業利益", "is.ebit.ob"); r += 1
qrow(r, "oi_abc", "Alphabet-level activities（避險營收 − 未分攤費用）", "is.ebit.abc"); r += 1
qrow(r, "oi", "合併營業利益", "is.ebit.total"); r += 1
qrow(r, "chk_oi", "檢查：分部營業利益合計 − 合併營業利益", formula=lambda c: f"{c}{QR['oi_gs']}+{c}{QR['oi_cloud']}+{c}{QR['oi_ob']}+{c}{QR['oi_abc']}-{c}{QR['oi']}", fmt="0.000", tag="—", src="—"); r += 1
r += 1
Q.cell(r, 1, "合併損益").font = BOLD; r += 1
qrow(r, "cogs", "營業成本", "is.cogs"); r += 1
qrow(r, "tac", "  其中：流量獲取成本 TAC", "is.tac", fy=False); r += 1
qrow(r, "rd", "研發費用", "is.rd"); r += 1
qrow(r, "sm", "行銷費用", "is.sm"); r += 1
qrow(r, "ga", "管理費用", "is.ga"); r += 1
qrow(r, "chk_cost", "檢查：營收 − 成本費用 − 營業利益", formula=lambda c: f"{c}{QR['rev']}-{c}{QR['cogs']}-{c}{QR['rd']}-{c}{QR['sm']}-{c}{QR['ga']}-{c}{QR['oi']}", fmt="0.000", tag="—", src="—"); r += 1
qrow(r, "oie", "其他收益淨額（含權益證券損益）", "is.oie"); r += 1
qrow(r, "eqg", "  其中：權益證券淨損益", "is.eqGain", fy=False); r += 1
qrow(r, "pt", "稅前利益", "is.pretax"); r += 1
qrow(r, "chk_pt", "檢查：營業利益＋其他收益 − 稅前利益", formula=lambda c: f"{c}{QR['oi']}+{c}{QR['oie']}-{c}{QR['pt']}", fmt="0.000", tag="—", src="—"); r += 1
qrow(r, "taxa", "所得稅", "is.tax"); r += 1
qrow(r, "ni", "淨利", "is.ni"); r += 1
qrow(r, "chk_ni", "檢查：稅前 − 所得稅 − 淨利", formula=lambda c: f"{c}{QR['pt']}-{c}{QR['taxa']}-{c}{QR['ni']}", fmt="0.000", tag="—", src="—"); r += 1
qrow(r, "sh", "稀釋加權股數（十億股）", "sh.dilW", h1="=AVERAGE(H{r}:I{r})".format(r=r), ttm=False); r += 1
qrow(r, "eps", "EPS（GAAP 稀釋，美元）", "is.epsDil", scale=1, fmt="0.00", ttm=True); r += 1
r += 1
Q.cell(r, 1, "現金流、資本支出與股東回饋").font = BOLD; r += 1
qrow(r, "da", "折舊（不動產與設備）", "cf.dep"); r += 1
qrow(r, "sbc", "股權報酬（SBC）", "cf.sbc"); r += 1
qrow(r, "cfo", "營運現金流", "cf.cfo"); r += 1
qrow(r, "capex", "資本支出（購置不動產與設備）", "cf.capex", neg=True); r += 1
qrow(r, "fcf", "自由現金流（公司定義：營運現金流 − 資本支出）", "cf.fcf"); r += 1
qrow(r, "chk_fcf", "檢查：營運現金流 − 資本支出 − 自由現金流", formula=lambda c: f"{c}{QR['cfo']}-{c}{QR['capex']}-{c}{QR['fcf']}", fmt="0.000", tag="—", src="—"); r += 1
qrow(r, "bb", "回購", "cf.buyback", neg=True); r += 1
qrow(r, "div", "普通股股利", "cf.div", neg=True); r += 1
r += 1
Q.cell(r, 1, "核對彙總（實績必須為 0）").font = BOLD; r += 1
QCHK = ["chk_gs", "chk_rev", "chk_oi", "chk_cost", "chk_pt", "chk_ni", "chk_fcf"]
Q.cell(r, 1, "檢查：上列核對列絕對值合計（8 季＋FY2024＋FY2025）"); QR["chk_all"] = r
Q[f"B{r}"] = "=" + "+".join(f"SUMPRODUCT(ABS(B{QR[k]}:K{QR[k]}))" for k in QCHK); Q[f"B{r}"].number_format = "0.000"; r += 1
FYK = ["gs", "cloud", "ob", "hedge", "rev", "oi_gs", "oi_cloud", "oi_ob", "oi_abc", "oi", "search", "yt", "net", "subs", "capex", "cfo", "da"]
Q.cell(r, 1, "檢查：FY2025 四季合計 − 年報（分部營收、產品線、分部營業利益、資本支出、營運現金、折舊）"); QR["chk_fy"] = r
Q[f"B{r}"] = "=" + "+".join(f"ABS(SUM(D{QR[k]}:G{QR[k]})-K{QR[k]})" for k in FYK); Q[f"B{r}"].number_format = "0.000"; r += 2

Q.cell(r, 1, "推導量（Seg、PL 頁的起始值；[Derived]）").font = BOLD; r += 1
hdr(Q, r, ["項目", "數值", "說明"]); r += 1
def qd(key, label, formula, fmt="0.0%", note=""):
    global r
    Q.cell(r, 1, label); Q[f"B{r}"] = formula; Q[f"B{r}"].number_format = fmt; Q.cell(r, 3, note)
    dn(key, f"Q!$B${r}"); QR[key] = r; r += 1
for k, lab in (("search", "Google Search & other"), ("yt", "YouTube ads"), ("net", "Google Network"), ("subs", "Google subscriptions, platforms, and devices"), ("cloud", "Google Cloud")):
    qd(f"g0_{k}", f"近四季年增：{lab}", f"=M{QR[k]}/N{QR[k]}-1", note="(2025Q3–2026Q2) ÷ (2024Q3–2025Q2) − 1；B 路徑 2H26 的年增，之後收斂到長期值")
qd("svc_ttm", "近四季 Google Services 營收", f"=M{QR['gs']}", "#,##0.000")
qd("rev3_ttm", "近四季三分部營收合計（不含避險）", f"=M{QR['gs']}+M{QR['cloud']}+M{QR['ob']}", "#,##0.000")
qd("sh_svc", "折舊分攤比例：Google Services（按近四季營收）", f"=M{QR['gs']}/rev3_ttm", note="公司不揭露分部折舊（技術基礎設施依使用量分攤，已含在分部營業利益）→ 按營收分攤（" + V22 + "；替代＝依使用量，雲端占比更高）")
qd("sh_cloud", "折舊分攤比例：Google Cloud", f"=M{QR['cloud']}/rev3_ttm")
qd("sh_ob", "折舊分攤比例：Other Bets", f"=M{QR['ob']}/rev3_ttm")
qd("da_ttm", "近四季折舊", f"=M{QR['da']}", "#,##0.000")
qd("m_svc_ttm", "Google Services EBITDA 率（近四季；營業利益＋分攤折舊）", f"=(M{QR['oi_gs']}+da_ttm*sh_svc)/M{QR['gs']}", note="v0.2.2 口徑 47.55%（mb_ttm＝1 時採用）")
qd("sh1_svc", "1H26 折舊分攤比例：Google Services", f"=L{QR['gs']}/(L{QR['gs']}+L{QR['cloud']}+L{QR['ob']})")
qd("sh1_cloud", "1H26 折舊分攤比例：Google Cloud", f"=L{QR['cloud']}/(L{QR['gs']}+L{QR['cloud']}+L{QR['ob']})")
qd("m_svc_1h", "Google Services EBITDA 率（1H26；營業利益＋分攤折舊）", f"=(L{QR['oi_gs']}+L{QR['da']}*sh1_svc)/L{QR['gs']}", note="與評價日機隊與 2026Q2 折舊年化（da_base）同一時點")
qd("m_svc", "Google Services EBITDA 率（預測期採用）", "=IF(mb_ttm=1,m_svc_ttm,m_svc_1h)", note="預設 1H26（最新已申報半年，利潤率仍在擴張；近四季口徑列 Sens）；自用 AI 增量成本另計")
qd("cloud_ebitda_ttm", "近四季 Google Cloud EBITDA（營業利益＋分攤折舊）", f"=M{QR['oi_cloud']}+da_ttm*sh_cloud", "#,##0.000")
qd("cloud_ttm", "近四季 Google Cloud 營收", f"=M{QR['cloud']}", "#,##0.000")
qd("cloud_pttm", "前四季 Google Cloud 營收", f"=N{QR['cloud']}", "#,##0.000")
qd("oi_ob_ttm", "近四季 Other Bets 營業利益", f"=M{QR['oi_ob']}", "#,##0.000")
qd("abc_ttm", "近四季 Alphabet-level activities", f"=M{QR['oi_abc']}", "#,##0.000")
qd("tac_sh", "TAC ÷ (Search＋Network) 近四季", f"=M{QR['tac']}/(M{QR['search']}+M{QR['net']})", note="增量 EBITDA 率 im 的參考")
for col in "BCDEFGHIJKLMN": Q.column_dimensions[col].width = 11
Q.column_dimensions["A"].width = 60; Q.column_dimensions["O"].width = 22; Q.column_dimensions["P"].width = 90
Q.freeze_panes = "B4"
QREF = lambda k, c: f"Q!{c}{QR[k]}"            # c：Q 頁欄
def QA(k, p):                                     # 實績欄對照：模型欄 → Q 頁欄
    return {"24": f"Q!J{QR[k]}", "25": f"Q!K{QR[k]}", "1H": f"Q!L{QR[k]}"}[p]

if os.environ.get("STAGE") == "1":
    for ws in (TH, SG, PL, CF, V, RC, SE, SO): wb.remove(ws)
    RD["A1"] = "Alphabet 個股模型 v0.3（段 1：Q 頁實績對帳）"
    wb.save(A.out); print("stage 1 saved", A.out); sys.exit(0)


# =====================================================================================================
# Thesis：命題模組（Cloud AI：MW × 每 MW 收入 × k；AI 資本支出 → 折舊 → 分部營業利益）
TH["A1"] = "Thesis — 命題模組：AI 容量（MW）× 每 MW 收入（Tokenomics 持有成本 × k）→ Google Cloud 的 Cloud AI 子列；同一套 MW 推出 AI 資本支出、AI 折舊與營運成本，接到分部營業利益（對外 → Google Cloud；自用 → Google Services／Alphabet 層級）"
TH["A2"] = "三情境共用（本版的情境軸是 Search，見 Seg 頁）；MW 速度、對外比例、k 的區間在 Sens 頁。沿用 v0.2.2：k 規格（kspec.py）、C10 基準速度、對外比例 30%、Tokenomics v21.4 快照。不以 Cloud 營收、RPO 反推 k 或每 MW 價格（工作單第 4 節）"
T = {}
r = 4
TH.cell(r, 1, "A. 價格與起始值（單一數值，C 欄）").font = BOLD; r += 1
hdr(TH, r, ["項目", "單位", "數值"] + [""] * 10 + ["說明"]); r += 1
def tscal(key, label, unit, formula, fmt="0.000", note=""):
    global r
    TH.cell(r, 1, label); TH.cell(r, 2, unit); TH[f"C{r}"] = formula; TH[f"C{r}"].number_format = fmt; TH.cell(r, NOTE_C, note)
    dn(key, f"Thesis!$C${r}"); T[key] = r; r += 1
tscal("k_long", "k_長約＝長約證據價格 ÷ IF_HoldEcon｜VR200", "x", "=kl_px/he_vr", note="kspec.py（v0.2.2）：證據價格不變、分母隨 Tokenomics 快照")
tscal("k_spot", "k_現貨＝H100 現貨指數 ÷ IF_GPUhrEcon｜H100", "x", "=ks_px/gpuhr_h", note="C1：MAG 三家一致")
tscal("k_blend", "混合 k＝長約占比 × k_長約＋(1 − 長約占比) × k_現貨", "x", "=long_sh*k_long+(1-long_sh)*k_spot", note="v0.2.2 為 1.201")
tscal("tpu_adj", "TPU 收入錨調整＝1 − TPU 占比 ×(1 − TPU 係數)", "x", "=1-tpu_sh*(1-tpu_f)", note="基準 1（C2：TPU × 1.0）")
tscal("tpu_adjc", "TPU IT 資本支出調整", "x", "=1-tpu_sh*(1-tpu_cf)")
tscal("permw0", "期初機隊每 MW 年收入（對外）", "US$m/MW-年", "=k_blend*tpu_adj*(mix0_h*he_h+mix0_b2*he_b2+mix0_b3*he_b3)", "0.00")
tscal("opex0", "期初機隊每 MW 年營運費用（Tokenomics IF_OpexGW 加權）", "US$m/MW-年", "=mix0_h*opx_h+mix0_b2*opx_b2+mix0_b3*opx_b3", "0.00")
tscal("ai_ttm", "近四季 Cloud AI 營收估計（近四季平均對外 MW × 期初每 MW 收入）", "US$B", "=ai_mw_ttm*permw0/1000", "#,##0.000", "只用於拆 Cloud 其他的起始成長率與利潤率")
tscal("ai_pttm", "前四季 Cloud AI 營收估計", "US$B", "=ai_mw_pttm*permw0/1000", "#,##0.000")
tscal("ai_ttm_e", "近四季 Cloud AI EBITDA 估計", "US$B", "=ai_mw_ttm*(permw0-opex0)/1000", "#,##0.000")
tscal("g0_co", "Cloud 其他近四季年增（Cloud − Cloud AI 估計）", "%", "=(cloud_ttm-ai_ttm)/(cloud_pttm-ai_pttm)-1", "0.0%", "[Derived]；Cloud 其他＝Workspace、非 AI 雲端、Wiz（2026-03 併入）")
tscal("m_co_ttm", "Cloud 其他 EBITDA 率（近四季：Cloud EBITDA − Cloud AI EBITDA 估計）÷ Cloud 其他營收", "%", "=(cloud_ebitda_ttm-ai_ttm_e)/(cloud_ttm-ai_ttm)", "0.0%", "[Derived]；mb_ttm＝1 時採用")
tscal("m_co_1h", "Cloud 其他 EBITDA 率（1H26）", "%", "=0", "0.0%", "[Derived]；(1H26 Cloud 營業利益＋分攤折舊 − Cloud AI 1H EBITDA) ÷ Cloud 其他 1H")
tscal("m_co", "Cloud 其他 EBITDA 率（預測期採用，固定）", "%", "=IF(mb_ttm=1,m_co_ttm,m_co_1h)", "0.0%", "預設 1H26（同 m_svc）")
r += 1
TH.cell(r, 1, "B. AI 容量（AI 總 IT MW；期末）— 依 Tokenomics 世代錨分桶").font = BOLD; r += 1
hdr(TH, r, MODEL_HDR); r += 1
BK = [("h", "Hopper 錨"), ("b2", "GB200 錨"), ("b3", "GB300 錨"), ("vr", "VR200 錨")]
MIX0 = {"h": "mix0_h", "b2": "mix0_b2", "b3": "mix0_b3", "vr": "0"}
def addmix(p, b):
    y = PY[p]
    if y == 2026: return "1" if b == "b3" else "0"
    if y == 2027: return {"b3": "am_27", "vr": "(1-am_27)"}.get(b, "0")
    return "1" if b == "vr" else "0"
def tput(key, label, unit, vals, fmt="#,##0.0", note=None, **kw):
    global r
    put(TH, r, label, unit, vals, fmt, note, **kw); T[key] = r; r += 1
for b, nm in BK:
    v = {"1H": f"=mw_tot0*{MIX0[b]}"}
    for p in FC:
        q = COL[prevp(p)]; ln = LEN[p]
        v[p] = f"={q}{r}+pace*{ln}*{addmix(p, b)}"
    v["26"] = f"=F{r}"
    tput(f"end_{b}", f"期末 MW：{nm}", "MW", v, note="期初（2026-06-30）＝AI 在役 MW × 期初機隊占比；新增依年度世代組合（2026 GB300 錨；2027 GB300／VR200 各半；2028 起 VR200 錨）" if b == "h" else None)
tput("end", "期末 AI 總 MW", "MW", {p: "=" + "+".join(f"{COL[p]}{T['end_' + b]}" for b, _ in BK) for p in ["1H"] + FC + ["26"]})
tput("end_ext", "期末對外 MW（＝總 MW × 對外比例）", "MW", {p: f"={COL[p]}{T['end']}*ext_sh" for p in ["1H"] + FC + ["26"]}, note="評價日 1,901（v0.2.2 defaults.billableOpen）")
for b, nm in BK:
    v = {"1H": f"=(mw_ye25+mw_tot0)/2*{MIX0[b]}"}
    for p in FC:
        v[p] = f"=({COL[prevp(p)]}{T['end_' + b]}+{COL[p]}{T['end_' + b]})/2"
    v["26"] = f"=(E{r}+F{r})/2"
    tput(f"avg_{b}", f"平均在役 MW：{nm}", "MW", v)
tput("avg", "平均在役 AI 總 MW", "MW", {p: "=" + "+".join(f"{COL[p]}{T['avg_' + b]}" for b, _ in BK) for p in ["1H"] + FC + ["26"]})
tput("avg_ext", "平均在役對外 MW", "MW", {p: f"={COL[p]}{T['avg']}*ext_sh" for p in ["1H"] + FC + ["26"]})
tput("adds", "AI 總 MW 新增", "MW", {**{p: f"=pace*{LEN[p]}" for p in FC}, "26": "=F{0}".format(r)}, note="三情境相同（預設，報告第 5 節）")
r += 1
TH.cell(r, 1, "C. Cloud AI（對外 AI 雲端；Google Cloud 的命題子列）").font = BOLD; r += 1
SUMHE = lambda c: "+".join(f"{c}{T['avg_' + b]}*he_{b}" for b, _ in BK)
SUMOP = lambda c: "+".join(f"{c}{T['avg_' + b]}*opx_{b}" for b, _ in BK)
SUMCAP = lambda c: "+".join(f"{c}{T['avg_' + b]}*cap_{b}" for b, _ in BK)
tput("ai_rev", "Cloud AI 營收＝對外比例 × Σ 平均 MW × IF_HoldEcon × k × TPU 調整", "US$B",
     {**{p: f"=ext_sh*k_blend*tpu_adj*({SUMHE(COL[p])})/1000*{LEN[p]}" for p in ["1H"] + FC}, "26": f"=E{r}+F{r}"}, "#,##0.00",
     note="收入端不乘 IF_Util（MW-year 計價）；1H26 為模型估計（公司未揭露 AI 營收），Cloud 其他＝實際 Cloud − 本列")
tput("permw", "每 MW 年收入（對外）", "US$m/MW-年", {p: f"={COL[p]}{T['ai_rev']}/({COL[p]}{T['avg_ext']}*{LEN[p]})*1000" for p in ["1H"] + FC}, "0.00")
tput("ai_opex", "Cloud AI 營運費用（Tokenomics IF_OpexGW，含電費）", "US$B", {**{p: f"=ext_sh*({SUMOP(COL[p])})/1000*{LEN[p]}" for p in ["1H"] + FC}, "26": f"=E{r}+F{r}"}, "#,##0.00")
tput("ai_ebitda", "Cloud AI EBITDA", "US$B", {p: f"={COL[p]}{T['ai_rev']}-{COL[p]}{T['ai_opex']}" for p in ["1H"] + FC + ["26"]}, "#,##0.00")
tput("ai_em", "  Cloud AI EBITDA 率", "%", {p: f"={COL[p]}{T['ai_ebitda']}/{COL[p]}{T['ai_rev']}" for p in ["1H"] + FC + ["26"]}, "0.0%")
tput("capchk", "上限檢查：每 MW 收入 ÷ 在役世代加權 IFW_RevGWFleet_Util（> 50% 示警；不得作預測）", "%",
     {p: f"={COL[p]}{T['permw']}/(({SUMCAP(COL[p])})/{COL[p]}{T['avg']})" for p in ["1H"] + FC}, "0.0%", note="v0.2.2 基準最高 38.0%")
TH[f"C{T['m_co_1h']}"] = f"=({QA('oi_cloud', '1H')}+{QA('da', '1H')}*sh1_cloud-E{T['ai_ebitda']})/({QA('cloud', '1H')}-E{T['ai_rev']})"
r += 1
TH.cell(r, 1, "D. AI 資本支出（新增 MW × Tokenomics 每 MW 成本；汰換只換 IT）").font = BOLD; r += 1
SUMADD = lambda p: "+".join(f"{addmix(p, b)}*cit_{b}" for b, _ in BK)
tput("cx_it", "AI IT 資本支出（成長）", "US$B", {**{p: f"={COL[p]}{T['adds']}*({SUMADD(p)})*tpu_adjc/1000" for p in FC}, "26": f"=F{r}"}, "#,##0.00")
tput("cx_fac", "AI 機房資本支出（自建比例 × IF_CapexFacility）", "US$B", {**{p: f"={COL[p]}{T['adds']}*selfb*cfac/1000" for p in FC}, "26": f"=F{r}"}, "#,##0.00")
tput("retire", "汰換 MW（6 年前上線的 MW，換成 VR200 錨 IT）", "MW", {"2H": 0, 27: 0, 28: 0, 29: "=mw_ye23", 30: "=mw_ye24-mw_ye23", 31: "=mw_ye25-mw_ye24",
     32: f"=mw_tot0-mw_ye25+F{T['adds']}", "26": f"=F{r}"}, note="IT 壽命 6 年：2023 年底在役 → 2029 汰換，以此類推；汰換後 MW 照常在役（收入不變）")
tput("cx_ref", "AI 汰換資本支出（只含 IT）", "US$B", {**{p: f"={COL[p]}{T['retire']}*cit_vr*tpu_adjc/1000" for p in FC}, "26": f"=F{r}"}, "#,##0.00")
tput("cx_ai_f", "AI 資本支出（公式）", "US$B", {p: f"={COL[p]}{T['cx_it']}+{COL[p]}{T['cx_fac']}+{COL[p]}{T['cx_ref']}" for p in FC + ["26"]}, "#,##0.00")
T["gap"] = r; r += 1          # 2H26 指引下限補足（Seg 寫完後填公式）
tput("cx_ai", "AI 資本支出（含指引下限補足）", "US$B", {**{p: f"={COL[p]}{T['cx_ai_f']}" for p in FC}, "2H": f"=F{T['cx_ai_f']}+F{T['gap']}", "26": f"=F{r}"}, "#,##0.00", bold=True)
tput("cx_rev", "  Cloud AI 營收 ÷ AI 資本支出", "x", {p: f"={COL[p]}{T['ai_rev']}/{COL[p]}{T['cx_ai']}" for p in FC}, "0.00")
r += 1
TH.cell(r, 1, "E. AI 折舊（在建工程與新增資本支出依年限直線攤提；評價日在役資產的折舊在 PL 頁「基礎折舊」）").font = BOLD; r += 1
TH.cell(r, 1, "每個資本支出年度：當期 × phi_0、之後每年 1、最後一年 1 − phi_0（年限 life_it／life_fac；IT 含汰換與指引補足；在建工程 2026H2 起轉入）")
r += 1
tput("cip_it", "在建工程轉入：IT（2026 年度批次）", "US$B", {"2H": "=cip_0*cip_it"}, "#,##0.00")
tput("cip_fac", "在建工程轉入：機房與網通（2026 年度批次）", "US$B", {"2H": "=cip_0*(1-cip_it)"}, "#,##0.00")
tput("v_it", "IT 批次金額（成長＋汰換＋指引補足）", "US$B", {p: (f"={COL[p]}{T['cx_it']}+{COL[p]}{T['cx_ref']}" + (f"+F{T['gap']}" if p == "2H" else "")) for p in FC}, "#,##0.00")
tput("v_fac", "機房批次金額", "US$B", {p: f"={COL[p]}{T['cx_fac']}" for p in FC}, "#,##0.00")
def depf(amt, age, life):
    return f"{amt}/{life}*IF({age}=0,phi_0,IF({age}<{life},1,IF({age}={life},1-phi_0,0)))"
def dep_row(srcs, life):
    """srcs：[(金額儲存格, 批次年度)]；回傳 {期間: 公式}"""
    out = {}
    for p in FC:
        terms = [depf(a, PY[p] - vy, life) for a, vy in srcs if PY[p] >= vy]
        out[p] = "=" + "+".join(terms)
    out["26"] = "=F{r}"
    return out
it_src = [(f"$F${T['cip_it']}", 2026)] + [(f"${COL[p]}${T['v_it']}", PY[p]) for p in FC]
fac_src = [(f"$F${T['cip_fac']}", 2026)] + [(f"${COL[p]}${T['v_fac']}", PY[p]) for p in FC]
d = dep_row(it_src, "life_it"); d["26"] = f"=F{r}"; tput("da_it", "新增 AI 折舊：IT", "US$B", d, "#,##0.00")
d = dep_row(fac_src, "life_fac"); d["26"] = f"=F{r}"; tput("da_fac", "新增 AI 折舊：機房", "US$B", d, "#,##0.00")
tput("da_ai", "新增 AI 折舊合計", "US$B", {p: f"={COL[p]}{T['da_it']}+{COL[p]}{T['da_fac']}" for p in FC + ["26"]}, "#,##0.00", bold=True)
tput("da_ext", "  其中：對外（→ Google Cloud）", "US$B", {p: f"={COL[p]}{T['da_ai']}*ext_sh" for p in FC + ["26"]}, "#,##0.00", note="按對外比例分攤（在建工程同）")
tput("da_int", "  其中：自用（→ Google Services／Alphabet 層級）", "US$B", {p: f"={COL[p]}{T['da_ai']}*(1-ext_sh)" for p in FC + ["26"]}, "#,##0.00")
r += 1
TH.cell(r, 1, "F. 自用 AI 增量成本（相對評價日機隊；評價日機隊的成本已在近四季分部利益率內）").font = BOLD; r += 1
tput("int_opex", "自用 AI 增量營運費用（Tokenomics IF_OpexGW）", "US$B",
     {**{p: f"=(1-ext_sh)*(({SUMOP(COL[p])})-mw_tot0*opex0)/1000*{LEN[p]}" for p in FC}, "26": f"=F{r}"}, "#,##0.00")
tput("rent", "SpaceX 算力租金（自用，營運成本）", "US$B", {"2H": "=rent_2H", 27: "=rent_27", 28: "=rent_28", 29: "=rent_29", 30: 0, 31: 0, 32: 0, "26": f"=F{r}"}, "#,##0.00")
tput("int_cost", "自用 AI 增量成本合計（營運＋租金＋自用折舊）", "US$B", {p: f"={COL[p]}{T['int_opex']}+{COL[p]}{T['rent']}+{COL[p]}{T['da_int']}" for p in FC + ["26"]}, "#,##0.00", bold=True)
tput("int_svc", "  歸 Google Services（× svc_sh）", "US$B", {p: f"={COL[p]}{T['int_cost']}*svc_sh" for p in FC + ["26"]}, "#,##0.00")
tput("int_abc", "  歸 Alphabet-level activities（× (1 − svc_sh)）", "US$B", {p: f"={COL[p]}{T['int_cost']}*(1-svc_sh)" for p in FC + ["26"]}, "#,##0.00")
widths(TH)
TH.freeze_panes = "C4"

# =====================================================================================================
# Seg：營收分部層（三情境各一完整區塊）
SG["A1"] = "Seg — 營收分部層：官方分部（Google Services、Google Cloud、Other Bets）＋Alphabet 層級避險；產品線為子列。情境差異只在 Google Search & other（A 擴大變現／B 近四季成長收斂／C 查詢分流、變現下降），其餘三情境相同"
SG["A2"] = "成長路徑：2H26 年增＝近四季年增（Q 頁），之後線性收斂到長期值（conv_n 年，FY2030 到達）。Search 情境：Search＝B 路徑 ×(1 − 分流) ×(1＋每次查詢變現變動)，分流與變現變動自 FY2026 的 0 線性到 FY2030 的 In 頁值。Cloud AI 由 Thesis 頁帶入"
SCN = {"A": ("dA", "mA", "情境 A（AI 搜尋擴大變現）"), "B": ("0", "0", "情境 B（近四季成長收斂；基準）"), "C": ("dC", "mC", "情境 C（查詢被 AI 助理分流、每次查詢變現下降）")}
SR = {s: {} for s in "ABC"}
r = 4
def gpath(g0, gl):
    return {**{"2H": f"={g0}"}, **{p: f"={g0}+({gl}-{g0})*MIN(1,({PY[p]}-2026)/conv_n)" for p in ANN}}
def sput(s, key, label, unit, vals, fmt="#,##0.00", note=None, **kw):
    global r
    put(SG, r, label, unit, vals, fmt, note, **kw); SR[s][key] = r; r += 1
for s in "ABC":
    dmax, mmax, title = SCN[s]
    SG.cell(r, 1, title).font = BOLD; SG.cell(r, 1).fill = SUB; r += 1
    hdr(SG, r, MODEL_HDR); r += 1
    R = SR[s]
    sput(s, "g_search", "  Search B 路徑年增", "%", gpath("g0_search", "gl_search"), "0.0%", "2H26 年增＝近四季 +16.8%（Q 頁），FY2030 收斂到 gl_search")
    gr = R["g_search"]
    v = {"24": "=" + QA("search", "24"), "25": "=" + QA("search", "25"), "1H": "=" + QA("search", "1H"),
         "2H": f"=(Q!F{QR['search']}+Q!G{QR['search']})*(1+F{gr})", "26": f"=E{r}+F{r}"}
    for p in ANN: v[p] = f"={COL[prevfy(p)]}{r}*(1+{COL[p]}{gr})"
    sput(s, "searchB", "Google Search & other：B 路徑", "US$B", v)
    sput(s, "ramp", "  情境爬坡（FY2026 0 → FY2030 1）", "x", {**{"2H": 0}, **{p: f"=MIN(1,({PY[p]}-2026)/conv_n)" for p in ANN}}, "0.00")
    sput(s, "div", "  查詢分流（相對 B 路徑）", "%", {p: f"={dmax}*{COL[p]}{R['ramp']}" for p in FC}, "0.0%",
         {"A": "In 頁 dA（負值＝AI 模式吸引查詢回流）", "B": "基準不設分流", "C": "In 頁 dC（壓力情境）"}[s])
    sput(s, "mon", "  每次查詢變現變動（相對 B 路徑）", "%", {p: f"={mmax}*{COL[p]}{R['ramp']}" for p in FC}, "0.0%")
    v = {p: f"={COL[p]}{R['searchB']}" for p in ("24", "25", "1H")}
    for p in FC: v[p] = f"={COL[p]}{R['searchB']}*(1-{COL[p]}{R['div']})*(1+{COL[p]}{R['mon']})"
    v["26"] = f"=E{r}+F{r}"
    sput(s, "search", "Google Search & other", "US$B", v, bold=True)
    for key, label, qk, g0, gl in (("yt", "YouTube ads", "yt", "g0_yt", "gl_yt"), ("net", "Google Network", "net", "g0_net", "gl_net"),
                                   ("subs", "Google subscriptions, platforms, and devices", "subs", "g0_subs", "gl_subs")):
        sput(s, "g_" + key, f"  {label} 年增", "%", gpath(g0, gl), "0.0%", "三情境相同")
        gr = R["g_" + key]
        v = {"24": "=" + QA(qk, "24"), "25": "=" + QA(qk, "25"), "1H": "=" + QA(qk, "1H"),
             "2H": f"=(Q!F{QR[qk]}+Q!G{QR[qk]})*(1+F{gr})", "26": f"=E{r}+F{r}"}
        for p in ANN: v[p] = f"={COL[prevfy(p)]}{r}*(1+{COL[p]}{gr})"
        sput(s, key, label, "US$B", v, bold=True)
    v = {p: "=" + QA("gs", p) for p in ("24", "25", "1H")}
    for p in FC: v[p] = f"={COL[p]}{R['search']}+{COL[p]}{R['yt']}+{COL[p]}{R['net']}+{COL[p]}{R['subs']}"
    v["26"] = f"=E{r}+F{r}"
    sput(s, "gs", "Google Services", "US$B", v, bold=True, fill=SUB)
    v = {p: f"=Thesis!{COL[p]}{T['ai_rev']}" for p in ["1H"] + FC + ["26"]}
    sput(s, "ai", "Cloud AI（命題子列）", "US$B", v, note="Thesis 頁：對外 MW × 每 MW 收入（1H26 為模型估計）")
    sput(s, "tpu", "TPU 系統銷售（命題子列）", "US$B", {"1H": 0, "2H": "=tpu_2h", "26": f"=E{r}+F{r}", **{p: "=tpu_yr" for p in ANN}}, note="C8 b：不走 MW × 每 MW；1H26 金額未揭露 → 0")
    sput(s, "g_co", "  Cloud 其他年增", "%", gpath("g0_co", "gl_co"), "0.0%", "起始＝Cloud 其他近四季年增（Thesis 頁 g0_co），FY2030 收斂到 gl_co")
    gr = R["g_co"]
    v = {"1H": f"={QA('cloud', '1H')}-E{R['ai']}-E{R['tpu']}", "2H": f"=E{r}*(1+F{gr})^0.5", "26": f"=E{r}+F{r}"}
    for p in ANN: v[p] = f"={COL[prevfy(p)]}{r}*(1+{COL[p]}{gr})"
    sput(s, "co", "Cloud 其他（命題子列）", "US$B", v, note="1H26＝Google Cloud 實績 − Cloud AI 估計 − TPU；2H26＝1H26 ×(1＋年增)^0.5")
    v = {p: "=" + QA("cloud", p) for p in ("24", "25", "1H")}
    for p in FC: v[p] = f"={COL[p]}{R['ai']}+{COL[p]}{R['tpu']}+{COL[p]}{R['co']}"
    v["26"] = f"=E{r}+F{r}"
    sput(s, "cloud", "Google Cloud", "US$B", v, bold=True, fill=SUB)
    v = {p: "=" + QA("ob", p) for p in ("24", "25", "1H")}
    for p in FC: v[p] = f"=ob_rev_{p}"
    v["26"] = f"=E{r}+F{r}"
    sput(s, "ob", "Other Bets", "US$B", v, bold=True, fill=SUB)
    v = {p: "=" + QA("hedge", p) for p in ("24", "25", "1H")}
    for p in FC: v[p] = 0
    v["26"] = f"=E{r}+F{r}"
    sput(s, "hedge", "Hedging gains (losses)", "US$B", v, note="Alphabet-level activities；預測期 0 [Assumed]")
    v = {p: "=" + QA("rev", p) for p in ("24", "25", "1H")}
    for p in FC: v[p] = f"={COL[p]}{R['gs']}+{COL[p]}{R['cloud']}+{COL[p]}{R['ob']}+{COL[p]}{R['hedge']}"
    v["26"] = f"=E{r}+F{r}"
    sput(s, "rev", "合併營收", "US$B", v, bold=True)
    sput(s, "chk1", "檢查：產品線合計 − Google Services", "US$B", {p: f"={COL[p]}{R['search']}+{COL[p]}{R['yt']}+{COL[p]}{R['net']}+{COL[p]}{R['subs']}-{COL[p]}{R['gs']}" for p in YK}, "0.000")
    sput(s, "chk2", "檢查：Cloud 子列合計 − Google Cloud", "US$B", {p: f"={COL[p]}{R['ai']}+{COL[p]}{R['tpu']}+{COL[p]}{R['co']}-{COL[p]}{R['cloud']}" for p in ["1H"] + FC + ["26"]}, "0.000")
    sput(s, "chk3", "檢查：分部合計＋避險 − 合併營收", "US$B", {p: f"={COL[p]}{R['gs']}+{COL[p]}{R['cloud']}+{COL[p]}{R['ob']}+{COL[p]}{R['hedge']}-{COL[p]}{R['rev']}" for p in YK}, "0.000")
    sput(s, "chk4", "檢查：實績欄合併營收 − Q 頁（FY2024、FY2025、1H26）", "US$B", {p: f"={COL[p]}{R['rev']}-{QA('rev', p)}" for p in ("24", "25", "1H")}, "0.000")
    sput(s, "cxs", "非 AI 資本支出：Google Services（營收 × cx_svc）", "US$B", {**{p: f"={COL[p]}{R['gs']}*cx_svc" for p in FC}, "26": f"=F{r}"})
    sput(s, "cxc", "非 AI 資本支出：Cloud 其他（營收 × cx_co）", "US$B", {**{p: f"={COL[p]}{R['co']}*cx_co" for p in FC}, "26": f"=F{r}"})
    sput(s, "cxo", "非 AI 資本支出：Other Bets（營收 × cx_ob）", "US$B", {**{p: f"={COL[p]}{R['ob']}*cx_ob" for p in FC}, "26": f"=F{r}"})
    sput(s, "cxn", "非 AI 資本支出合計", "US$B", {p: f"={COL[p]}{R['cxs']}+{COL[p]}{R['cxc']}+{COL[p]}{R['cxo']}" for p in FC + ["26"]}, bold=True)
    for key, src, lab in (("dns", "cxs", "Google Services"), ("dnc", "cxc", "Cloud 其他"), ("dno", "cxo", "Other Bets")):
        d = dep_row([(f"${COL[p]}${R[src]}", PY[p]) for p in FC], "life_na"); d["26"] = f"=F{r}"
        sput(s, key, f"新增非 AI 折舊：{lab}（年限 life_na）", "US$B", d)
    sput(s, "dnt", "新增非 AI 折舊合計", "US$B", {p: f"={COL[p]}{R['dns']}+{COL[p]}{R['dnc']}+{COL[p]}{R['dno']}" for p in FC + ["26"]}, bold=True)
    r += 1
widths(SG)
SG.freeze_panes = "C4"
# 2H26 指引下限補足（用情境 B 的非 AI 資本支出；2H26 三情境相同）
put(TH, T["gap"], "2H26 指引下限補足（MAX(0, 指引 − 上半年 − 公式)；歸 AI IT）", "US$B",
    {"2H": f"=MAX(0,capex_g-capex_1h-(F{T['cx_ai_f']}+Seg!F{SR['B']['cxn']}))", "26": f"=F{T['gap']}"}, "#,##0.00",
    "C10：首期資本支出＝MAX(公式, 公司指引)；2H26 三情境營收相同（情境爬坡自 FY2027 起），用情境 B 的非 AI 資本支出")

# =====================================================================================================
# PL：分部營收 → 分部 EBITDA／營業利益 → 合併 → EPS（三情境）
PL["A1"] = "PL — 分部營收 → 分部營業利益 → 合併營業利益 → 業外 → 稅 → 淨利 → EPS（三情境；US$B，EPS 美元）。實績欄（FY2024A、FY2025A、1H26A）取自 Q 頁；FY2026E＝1H 實績＋2H 模型"
PL["A2"] = ("分部營業利益＝分部 EBITDA − 分攤折舊。分部 EBITDA：Google Services＝B 路徑營收 × 近四季 EBITDA 率＋Search 情境差額 × 增量 EBITDA 率 − 自用 AI 增量營運成本 × svc_sh；"
            "Google Cloud＝Cloud AI EBITDA（Thesis）＋TPU × tpu_m＋Cloud 其他 × m_co。折舊＝評價日在役資產基礎折舊（按近四季營收分攤）＋新增 AI 折舊（對外→Cloud、自用→Services／Alphabet 層級）＋新增非 AI 折舊。持股評價損益預測期設 0")
PR_ = {s: {} for s in "ABC"}
r = 4
PL.cell(r, 1, "分部對帳彙總（三情境 × 全部年度取絕對值最大者；實績欄同列）").font = BOLD; r += 1
hdr(PL, r, MODEL_HDR); r += 1
SUMROW = r; r += 5   # 預留 4 列核對＋空列
def pput(s, key, label, unit, vals, fmt="#,##0.00", note=None, **kw):
    global r
    put(PL, r, label, unit, vals, fmt, note, **kw); PR_[s][key] = r; r += 1
for s in "ABC":
    R = SR[s]; P = PR_[s]
    PL.cell(r, 1, SCN[s][2]).font = BOLD; PL.cell(r, 1).fill = SUB; r += 1
    hdr(PL, r, MODEL_HDR); r += 1
    ref = lambda k: {p: f"=Seg!{COL[p]}{R[k]}" for p in YK if SG[f"{COL[p]}{R[k]}"].value is not None}
    pput(s, "search", "Google Search & other", "US$B", ref("search"))
    pput(s, "yt", "YouTube ads", "US$B", ref("yt"))
    pput(s, "net", "Google Network", "US$B", ref("net"))
    pput(s, "subs", "Google subscriptions, platforms, and devices", "US$B", ref("subs"))
    pput(s, "gs", "Google Services", "US$B", ref("gs"), bold=True, fill=SUB)
    pput(s, "ai", "Cloud AI（命題子列）", "US$B", ref("ai"))
    pput(s, "tpu", "TPU 系統銷售（命題子列）", "US$B", ref("tpu"))
    pput(s, "co", "Cloud 其他（命題子列）", "US$B", ref("co"))
    pput(s, "cloud", "Google Cloud", "US$B", ref("cloud"), bold=True, fill=SUB)
    pput(s, "ob", "Other Bets", "US$B", ref("ob"), bold=True, fill=SUB)
    pput(s, "hedge", "Hedging gains (losses)", "US$B", ref("hedge"))
    pput(s, "rev", "合併營收", "US$B", ref("rev"), bold=True)
    pput(s, "g", "  合併營收年增", "%", {p: f"={COL[p]}{P['rev']}/{COL[q]}{P['rev']}-1" for p, q in (("25", "24"), ("26", "25"), (27, "26"), (28, 27), (29, 28), (30, 29), (31, 30), (32, 31))}, "0.0%")
    pput(s, "chk1", "檢查：產品線合計 − Google Services", "US$B", {p: f"={COL[p]}{P['search']}+{COL[p]}{P['yt']}+{COL[p]}{P['net']}+{COL[p]}{P['subs']}-{COL[p]}{P['gs']}" for p in YK}, "0.000")
    pput(s, "chk2", "檢查：Cloud 子列合計 − Google Cloud", "US$B", {p: f"={COL[p]}{P['ai']}+{COL[p]}{P['tpu']}+{COL[p]}{P['co']}-{COL[p]}{P['cloud']}" for p in ["1H"] + FC + ["26"]}, "0.000")
    pput(s, "chk3", "檢查：分部合計＋避險 − 合併營收", "US$B", {p: f"={COL[p]}{P['gs']}+{COL[p]}{P['cloud']}+{COL[p]}{P['ob']}+{COL[p]}{P['hedge']}-{COL[p]}{P['rev']}" for p in YK}, "0.000")
    # 分部 EBITDA 與營業利益
    pput(s, "dsr", "  Search 情境差額（相對 B 路徑）", "US$B", {**{p: f"=Seg!{COL[p]}{R['search']}-Seg!{COL[p]}{R['searchB']}" for p in FC}, "26": f"=F{r}"})
    pput(s, "e_gs", "Google Services EBITDA", "US$B",
         {**{p: f"=m_svc*({COL[p]}{P['gs']}-{COL[p]}{P['dsr']})+inc_m*{COL[p]}{P['dsr']}-svc_sh*(Thesis!{COL[p]}{T['int_opex']}+Thesis!{COL[p]}{T['rent']})" for p in FC}},
         note="B 路徑營收 × m_svc（Q 頁近四季 47.55%）＋Search 差額 × inc_m − 自用 AI 增量營運與租金 × svc_sh")
    pput(s, "d_gs", "Google Services 折舊（基礎分攤＋新增非 AI＋自用 AI × svc_sh）", "US$B",
         {p: f"=da_base*{LEN[p]}*sh_svc+Seg!{COL[p]}{R['dns']}+Thesis!{COL[p]}{T['da_int']}*svc_sh" for p in FC})
    v = {p: "=" + QA("oi_gs", p) for p in ("24", "25", "1H")}
    for p in FC: v[p] = f"={COL[p]}{P['e_gs']}-{COL[p]}{P['d_gs']}"
    v["26"] = f"=E{r}+F{r}"
    pput(s, "oi_gs", "Google Services 營業利益", "US$B", v, bold=True)
    pput(s, "e_cl", "Google Cloud EBITDA", "US$B", {p: f"=Thesis!{COL[p]}{T['ai_ebitda']}+{COL[p]}{P['tpu']}*tpu_m+{COL[p]}{P['co']}*m_co" for p in FC},
         note="Cloud AI EBITDA（Tokenomics 營運成本）＋TPU × tpu_m＋Cloud 其他 × m_co")
    pput(s, "d_cl", "Google Cloud 折舊（基礎分攤＋對外 AI＋新增非 AI）", "US$B", {p: f"=da_base*{LEN[p]}*sh_cloud+Thesis!{COL[p]}{T['da_ext']}+Seg!{COL[p]}{R['dnc']}" for p in FC})
    v = {p: "=" + QA("oi_cloud", p) for p in ("24", "25", "1H")}
    for p in FC: v[p] = f"={COL[p]}{P['e_cl']}-{COL[p]}{P['d_cl']}"
    v["26"] = f"=E{r}+F{r}"
    pput(s, "oi_cl", "Google Cloud 營業利益", "US$B", v, bold=True)
    pput(s, "d_ob", "Other Bets 折舊（基礎分攤＋新增非 AI）", "US$B", {p: f"=da_base*{LEN[p]}*sh_ob+Seg!{COL[p]}{R['dno']}" for p in FC})
    v = {p: "=" + QA("oi_ob", p) for p in ("24", "25", "1H")}
    for p in FC: v[p] = f"=ob_oi*{LEN[p]}"
    v["26"] = f"=E{r}+F{r}"
    pput(s, "oi_ob", "Other Bets 營業利益", "US$B", v, bold=True)
    pput(s, "e_ob", "Other Bets EBITDA", "US$B", {p: f"={COL[p]}{P['oi_ob']}+{COL[p]}{P['d_ob']}" for p in FC})
    pput(s, "d_abc", "Alphabet-level 折舊（自用 AI × (1 − svc_sh)）", "US$B", {p: f"=Thesis!{COL[p]}{T['da_int']}*(1-svc_sh)" for p in FC})
    v = {p: "=" + QA("oi_abc", p) for p in ("24", "25", "1H")}
    for p in FC: v[p] = f"={COL[p]}{P['hedge']}-abc_rate*{COL[p]}{P['rev']}-Thesis!{COL[p]}{T['int_abc']}"
    v["26"] = f"=E{r}+F{r}"
    pput(s, "oi_abc", "Alphabet-level activities（避險營收 − 未分攤費用 − 自用 AI 增量成本 × (1 − svc_sh)）", "US$B", v, bold=True)
    pput(s, "e_abc", "Alphabet-level EBITDA", "US$B", {p: f"={COL[p]}{P['oi_abc']}+{COL[p]}{P['d_abc']}" for p in FC})
    pput(s, "oi_seg", "分部營業利益合計", "US$B", {p: f"={COL[p]}{P['oi_gs']}+{COL[p]}{P['oi_cl']}+{COL[p]}{P['oi_ob']}+{COL[p]}{P['oi_abc']}" for p in YK})
    pput(s, "ebitda", "合併 EBITDA（分部 EBITDA 合計）", "US$B", {**{p: f"={COL[p]}{P['e_gs']}+{COL[p]}{P['e_cl']}+{COL[p]}{P['e_ob']}+{COL[p]}{P['e_abc']}" for p in FC}})
    v = {p: "=" + QA("da", p) for p in ("24", "25", "1H")}
    for p in FC: v[p] = f"=da_base*{LEN[p]}+Thesis!{COL[p]}{T['da_ai']}+Seg!{COL[p]}{R['dnt']}"
    v["26"] = f"=E{r}+F{r}"
    pput(s, "da", "折舊合計（基礎＋新增 AI＋新增非 AI）", "US$B", v, note="1H26 與以前為現金流量表實績")
    pput(s, "da_ai", "  其中：新增 AI 折舊", "US$B", {**{p: f"=Thesis!{COL[p]}{T['da_ai']}" for p in FC}, "26": f"=F{r}"})
    v = {p: "=" + QA("oi", p) for p in ("24", "25", "1H")}
    for p in FC: v[p] = f"={COL[p]}{P['ebitda']}-{COL[p]}{P['da']}"
    v["26"] = f"=E{r}+F{r}"
    pput(s, "oi", "合併營業利益", "US$B", v, bold=True, fill=SUB, note="預測期＝合併 EBITDA − 折舊合計（與分部營業利益合計獨立計算，下一列核對折舊分攤）")
    pput(s, "chk4", "檢查：分部營業利益合計 − 合併營業利益", "US$B", {p: f"={COL[p]}{P['oi_seg']}-{COL[p]}{P['oi']}" for p in YK}, "0.000")
    pput(s, "om", "  營業利益率", "%", {p: f"={COL[p]}{P['oi']}/{COL[p]}{P['rev']}" for p in YK}, "0.0%")
    pput(s, "om_gs", "  Google Services 營業利益率", "%", {p: f"={COL[p]}{P['oi_gs']}/{COL[p]}{P['gs']}" for p in YK}, "0.0%")
    pput(s, "om_cl", "  Google Cloud 營業利益率", "%", {p: f"={COL[p]}{P['oi_cl']}/{COL[p]}{P['cloud']}" for p in YK}, "0.0%")
    P["ii"] = r; r += 1; P["ie"] = r; r += 1        # 利息收支（CF 頁寫完後填公式）
    v = {"25": f"=SUM(Q!D{QR['eqg']}:G{QR['eqg']})", "1H": "=" + QA("eqg", "1H")}
    for p in FC: v[p] = 0
    v["26"] = f"=E{r}+F{r}"
    pput(s, "eqg", "權益證券淨損益（預測期設 0）", "US$B", v, note="2026H1 135.9 主要為 SpaceX 上市與 Anthropic 重估（非營業）；預測期 0（工作單第 2 節 PL）")
    v = {p: "=" + QA("oie", p) for p in ("24", "25", "1H")}
    for p in FC: v[p] = f"={COL[p]}{P['ii']}-{COL[p]}{P['ie']}+{COL[p]}{P['eqg']}"
    v["26"] = f"=E{r}+F{r}"
    pput(s, "oie", "業外收益合計", "US$B", v)
    pput(s, "pt", "稅前利益", "US$B", {p: f"={COL[p]}{P['oi']}+{COL[p]}{P['oie']}" for p in YK})
    v = {p: "=" + QA("taxa", p) for p in ("24", "25", "1H")}
    for p in FC: v[p] = f"={COL[p]}{P['pt']}*tax"
    v["26"] = f"=E{r}+F{r}"
    pput(s, "tx", "所得稅", "US$B", v)
    pput(s, "ni", "淨利", "US$B", {p: f"={COL[p]}{P['pt']}-{COL[p]}{P['tx']}" for p in YK}, bold=True)
    pput(s, "chk5", "檢查：實績淨利 − Q 頁", "US$B", {p: f"={COL[p]}{P['ni']}-{QA('ni', p)}" for p in ("24", "25", "1H")}, "0.000")
    pput(s, "pref", "特別股股利", "US$B", {"1H": 0, "2H": "=pref_2H", "26": f"=E{r}+F{r}", 27: "=pref_27", 28: "=pref_28", 29: "=pref_29", 30: 0, 31: 0, 32: 0})
    pput(s, "nic", "歸屬普通股淨利", "US$B", {p: f"={COL[p]}{P['ni']}-{COL[p]}{P['pref']}" for p in ["1H"] + FC + ["26"]})
    P["she"] = r; r += 1; P["sh"] = r; r += 1        # 股數（CF 頁寫完後填公式）
    pput(s, "eps", "EPS（GAAP 稀釋，美元）", "US$", {**{p: f"={COL[p]}{P['ni']}/{COL[p]}{P['sh']}" for p in ("24", "25")},
                                                    **{p: f"={COL[p]}{P['nic']}/{COL[p]}{P['sh']}" for p in ["1H"] + FC + ["26"]}}, "0.00", bold=True)
    pput(s, "epsx", "EPS（不含權益證券損益，美元）", "US$", {p: f"=({COL[p]}{P['nic']}-{COL[p]}{P['eqg']}*(1-tax_eq))/{COL[p]}{P['sh']}" for p in ["1H"] + FC + ["26"]}, "0.00",
         note="權益證券損益以 tax_eq 21% 稅後扣除；可與共識 non-GAAP EPS 比較")
    v = {p: "=" + QA("sbc", p) for p in ("24", "25", "1H")}
    for p in FC: v[p] = f"={COL[p]}{P['rev']}*sbc_r"
    v["26"] = f"=E{r}+F{r}"
    pput(s, "sbc", "股權報酬（已含在分部營業利益內）", "US$B", v)
    r += 1
widths(PL)
PL.freeze_panes = "C4"
CHKS = [("chk1", "檢查：產品線合計 − Google Services"), ("chk2", "檢查：Cloud 子列合計 − Google Cloud"),
        ("chk3", "檢查：分部合計＋避險 − 合併營收"), ("chk4", "檢查：分部營業利益合計 − 合併營業利益")]
for i, (k, lab) in enumerate(CHKS):
    rr = SUMROW + i
    PL.cell(rr, 1, lab); PL.cell(rr, 2, "US$B")
    for p in YK:
        c = COL[p]
        if PL[f"{c}{PR_['A'][k]}"].value is None: continue
        PL[f"{c}{rr}"] = "=MAX(" + ",".join(f"ABS({c}{PR_[s][k]})" for s in "ABC") + ")"
        PL[f"{c}{rr}"].number_format = "0.000"
    PL.cell(rr, NOTE_C, "三情境各區塊同名核對列的最大絕對值；必須為 0（tools/segment_check 讀本列）")

# =====================================================================================================
# CF：營運現金 − 資本支出 − 租賃本金 → 自由現金流；股利、回購、融資；收支關卡
CF["A1"] = "CF — 現金流（US$B）。營運現金＝淨利＋折舊＋股權報酬＋營運資金；資本支出＝AI（Thesis）＋非 AI（Seg）；自由現金流＝營運現金 − 資本支出 − 融資租賃本金。融資順序：現金 → 減少回購 → 新債（最低現金 50）"
CF["A2"] = "2026 年：上半年發行普通股 30.5B（承銷＋Berkshire）與強制轉換特別股 19.1B、暫停回購（C4：現金牛最強，但 2026 已轉向外部融資）；2H26 回購 0；FY2027 起回購目標 17.0／年（現金不足時先減回購）"
CR = {s: {} for s in "ABC"}
r = 4
CF.cell(r, 1, "收支關卡（地圖讀這兩列；模型期間 FY2026–FY2030，FY2026 含上半年實績）").font = BOLD; r += 1
hdr(CF, r, ["項目", "情境 A", "情境 B", "情境 C", "說明"]); r += 1
GATE = r; r += 4
def cput(s, key, label, unit, vals, fmt="#,##0.00", note=None, **kw):
    global r
    put(CF, r, label, unit, vals, fmt, note, **kw); CR[s][key] = r; r += 1
for s in "ABC":
    R = SR[s]; P = PR_[s]; C = CR[s]
    CF.cell(r, 1, SCN[s][2]).font = BOLD; CF.cell(r, 1).fill = SUB; r += 1
    hdr(CF, r, MODEL_HDR); r += 1
    cput(s, "ni", "淨利", "US$B", {p: f"=PL!{COL[p]}{P['ni']}" for p in YK})
    cput(s, "da", "折舊", "US$B", {p: f"=PL!{COL[p]}{P['da']}" for p in YK})
    cput(s, "sbc", "股權報酬（非現金）", "US$B", {p: f"=PL!{COL[p]}{P['sbc']}" for p in YK})
    cput(s, "wc", "營運資金變動（流入＋）", "US$B", {"2H": "=wc_2h", **{p: f"=wc_k*(PL!{COL[p]}{P['rev']}-PL!{COL[prevfy(p)]}{P['rev']})" for p in ANN}, "26": f"=F{r}"},
         note="2H26＝上一年度同期實際（C14）；之後＝營收增量 × wc_k")
    v = {p: f"={QA('cfo', p)}-{COL[p]}{C['ni']}-{COL[p]}{C['da']}-{COL[p]}{C['sbc']}" for p in ("24", "25", "1H")}
    for p in FC: v[p] = 0
    v["26"] = f"=E{r}+F{r}"
    cput(s, "oth", "其他非現金調整（實績差額：權益證券利益迴轉、遞延稅等）", "US$B", v)
    cput(s, "cfo", "營運現金流", "US$B", {p: f"={COL[p]}{C['ni']}+{COL[p]}{C['da']}+{COL[p]}{C['sbc']}+{COL[p]}{C['wc']}+{COL[p]}{C['oth']}" for p in YK}, bold=True)
    cput(s, "cx_ai", "資本支出：AI（Thesis；首期含指引下限補足）", "US$B", {**{p: f"=Thesis!{COL[p]}{T['cx_ai']}" for p in FC}, "26": f"=F{r}"})
    cput(s, "cx_na", "資本支出：非 AI（Seg）", "US$B", {**{p: f"=Seg!{COL[p]}{R['cxn']}" for p in FC}, "26": f"=F{r}"})
    v = {p: "=" + QA("capex", p) for p in ("24", "25", "1H")}
    for p in FC: v[p] = f"={COL[p]}{C['cx_ai']}+{COL[p]}{C['cx_na']}"
    v["26"] = f"=E{r}+F{r}"
    cput(s, "cx", "資本支出合計", "US$B", v, bold=True)
    cput(s, "lease", "融資租賃本金", "US$B", {"1H": 0.84, **{p: f"=lease_{p}" for p in FC}, "26": f"=E{r}+F{r}"}, note="1H26：10-Q 附註 4 補充資訊 0.840 [Interested-party]；FY2024–25 未列（公司 FCF 定義不扣）")
    cput(s, "fcf", "自由現金流（營運現金 − 資本支出 − 融資租賃本金）", "US$B", {p: f"={COL[p]}{C['cfo']}-{COL[p]}{C['cx']}-N({COL[p]}{C['lease']})" for p in YK}, bold=True, fill=SUB)
    v = {p: "=" + QA("div", p) for p in ("24", "25", "1H")}
    v["2H"] = "=dps/2*sh_out"
    for p in ANN: v[p] = f"=dps*(1+dps_g)^({PY[p]}-2026)*PL!{COL[prevp(p)]}{P['she']}"
    v["26"] = f"=E{r}+F{r}"
    cput(s, "div", "普通股股利", "US$B", v, note="2H26＝0.22 × 2 × 流通股；之後每股 +5%／年 × 期初稀釋股數（避免與回購循環）")
    cput(s, "pref", "特別股股利", "US$B", {p: f"=PL!{COL[p]}{P['pref']}" for p in ["1H"] + FC + ["26"]})
    cput(s, "fcfd", "自由現金流扣股利", "US$B", {p: f"={COL[p]}{C['fcf']}-{COL[p]}{C['div']}-N({COL[p]}{C['pref']})" for p in YK}, bold=True)
    cput(s, "anth", "策略投資（Anthropic 出資；1H26 含 Wiz、Intersect 收購）", "US$B", {"1H": 55.44, **{p: f"=anth_{p}" for p in FC}, "26": f"=E{r}+F{r}"},
         note="1H26：非上市證券 22.1＋收購 33.7＋其他 − 出售（legacy_v0.2/company.json ytdActual.jv）[Derived]；投資活動，不在自由現金流")
    cput(s, "mat", "債務到期償還", "US$B", {"1H": 5.253, **{p: f"=mat_{p}" for p in FC}, "26": f"=E{r}+F{r}"}, note="1H26：10-Q 償還債務 [Interested-party]")
    cput(s, "newd", "已發行新債", "US$B", {"1H": 56.226, "2H": "=debt_new_2h", **{p: 0 for p in ANN}, "26": f"=E{r}+F{r}"}, note="1H26：發債淨額（10-Q）；2H26：2026-08 美元票據 25.0")
    cput(s, "eq", "股權發行（普通股＋強制轉換特別股，淨額）", "US$B", {"1H": 49.562, **{p: 0 for p in FC}, "26": f"=E{r}+F{r}"}, note="C4：2026-06 普通股 30.5（承銷 20.5＋Berkshire 10.0）＋特別股 19.1 [Interested-party]；ATM 40B 與代扣稅不建模")
    cput(s, "copen", "期初現金（現金＋債權類有價證券）", "US$B", {"1H": "=cash_ye25", **{p: f"={COL[prevp(p)]}{r + 5}" for p in FC}})
    C["oth2"] = r; r += 1
    cput(s, "pre", "回購前現金", "US$B", {p: f"={COL[p]}{C['copen']}+{COL[p]}{C['fcf']}-{COL[p]}{C['div']}-{COL[p]}{C['pref']}-{COL[p]}{C['anth']}-{COL[p]}{C['mat']}+{COL[p]}{C['newd']}+{COL[p]}{C['eq']}+{COL[p]}{C['oth2']}" for p in ["1H"] + FC})
    v = {"1H": "=" + QA("bb", "1H"), "2H": 0}
    for p in ANN: v[p] = f"=MAX(0,MIN(bb_t,{COL[p]}{C['pre']}-min_cash))"
    v["26"] = f"=E{r}+F{r}"
    cput(s, "bb", "回購", "US$B", v, note="2H26＝0（暫停，C4）；之後＝MIN(回購目標, 回購前現金 − 最低現金)")
    cput(s, "plug", "新增舉債（現金低於最低現金時）", "US$B", {"1H": 0, **{p: f"=MAX(0,min_cash-({COL[p]}{C['pre']}-{COL[p]}{C['bb']}))" for p in FC}, "26": f"=E{r}+F{r}"})
    cput(s, "cend", "期末現金（現金＋債權類有價證券）", "US$B", {p: f"={COL[p]}{C['pre']}-{COL[p]}{C['bb']}+{COL[p]}{C['plug']}" for p in ["1H"] + FC})
    assert C["cend"] == C["copen"] + 5
    put(CF, C["oth2"], "其他（1H26 實績差額：代扣稅、capped call、匯率；預測期 0）", "US$B",
        {"1H": f"=cash_0-(E{C['copen']}+E{C['fcf']}-E{C['div']}-E{C['pref']}-E{C['anth']}-E{C['mat']}+E{C['newd']}+E{C['eq']})", **{p: 0 for p in FC}})
    cput(s, "chk", "檢查：1H26 期末現金 − 10-Q（應為 0）", "US$B", {"1H": f"=E{C['cend']}-cash_0"}, "0.000")
    cput(s, "dopen", "期初債務", "US$B", {"1H": 49.085, **{p: f"={COL[prevp(p)]}{r + 1}" for p in FC}}, note="2025-12-31 面額 49.085（事實總帳 debt.face.fy2025）")
    cput(s, "dend", "期末債務", "US$B", {"1H": "=debt_0", **{p: f"={COL[p]}{C['dopen']}-{COL[p]}{C['mat']}+{COL[p]}{C['newd']}+{COL[p]}{C['plug']}" for p in FC}}, note="1H26＝10-Q 面額（含匯率變動）")
    cput(s, "nc", "淨現金（現金 − 債務）", "US$B", {p: f"={COL[p]}{C['cend']}-{COL[p]}{C['dend']}" for p in ["1H"] + FC})
    cput(s, "bpx", "回購均價（現價 × (1＋ke)^年數）", "US$", {**{p: f"=px*(1+ke)^({PY[p] + 0.5}-2026.82)" for p in ANN}, "2H": "=px"}, "0.00")
    r += 1
widths(CF)
CF.freeze_panes = "C4"
CF.cell(GATE, 1, "收支關卡：自由現金流扣股利（五年合計）")
CF.cell(GATE + 1, 1, "  其中：五年新增舉債")
CF.cell(GATE + 2, 1, "  其中：五年回購合計（FY2027–FY2030 目標 4 × bb_t）")
CF.cell(GATE + 3, 1, "是否需要舉債或停止回購")
for j, s in enumerate("ABC"):
    c = "BCD"[j]; C = CR[s]
    CF[f"{c}{GATE}"] = f"=SUM(G{C['fcfd']}:K{C['fcfd']})"; CF[f"{c}{GATE}"].number_format = "#,##0.0"
    CF[f"{c}{GATE + 1}"] = f"=SUM(G{C['plug']}:K{C['plug']})"; CF[f"{c}{GATE + 1}"].number_format = "#,##0.0"
    CF[f"{c}{GATE + 2}"] = f"=SUM(G{C['bb']}:K{C['bb']})"; CF[f"{c}{GATE + 2}"].number_format = "#,##0.0"
    CF[f"{c}{GATE + 3}"] = (f'=IF({c}{GATE + 1}>0.001,"需要舉債"&IF({c}{GATE + 2}<4*bb_t-0.001,"且減少回購",""),'
                            f'IF({c}{GATE + 2}<4*bb_t-0.001,"不需舉債，但須減少回購","不需舉債，回購照目標"))')
CF.cell(GATE, 5, "FY2026–FY2030 自由現金流（扣融資租賃本金）− 普通股與特別股股利；FY2026 含上半年實績。負值＝入不敷出")
CF.cell(GATE + 3, 5, "依現金瀑布：現金 → 減少回購 → 新債（最低現金 50）；2026 上半年已發行股權 49.6、2H26 起新債 25.0 為已發生事件")
# 回填 PL 的利息與股數
for s in "ABC":
    P = PR_[s]; C = CR[s]
    put(PL, P["ii"], "利息收入（期初現金 × r_cash）", "US$B", {**{p: f"=r_cash*CF!{COL[p]}{C['copen']}*{LEN[p]}" for p in FC}, "26": f"=F{P['ii']}"})
    put(PL, P["ie"], "利息費用（期初債務 × r_debt）", "US$B", {**{p: f"=r_debt*CF!{COL[p]}{C['dopen']}*{LEN[p]}" for p in FC}, "26": f"=F{P['ie']}"})
    v = {"1H": "=sh_dil_q2+(sh_out-sh_basic_q2)"}
    for p in FC:
        q = COL[prevp(p)]
        v[p] = f"={q}{P['she']}*(1+dil*{LEN[p]})-CF!{COL[p]}{C['bb']}/CF!{COL[p]}{C['bpx']}" + ("+conv_sh" if p == 29 else "")
    put(PL, P["she"], "期末稀釋股數（十億股）", "十億股", v, "0.000", "1H26 期末＝Q2 稀釋加權＋(7/15 流通股 − Q2 基本加權)；之後 ×(1＋dil) − 回購 ÷ 回購均價；2029 強制轉換特別股 +conv_sh")
    v = {"24": "=" + QA("sh", "24"), "25": "=" + QA("sh", "25"), "1H": "=" + QA("sh", "1H"), "26": f"=(E{P['sh']}+F{P['sh']})/2"}
    for p in FC: v[p] = f"=({COL[prevp(p)]}{P['she']}+{COL[p]}{P['she']})/2"
    v["2H"] = f"=E{P['she']}*(1+dil*0.25)"
    put(PL, P["sh"], "平均稀釋股數（十億股）", "十億股", v, "0.000")

# =====================================================================================================
# Val：0.1 DCF、0.2 分部加總（SOTP）、0.3 結論
V["A1"] = "Val — 目標價（約 12 個月後，2027-10-27）＝DCF 50%＋分部加總（SOTP）50%；情境權重 A 25%／B 50%／C 25%（In 頁，黃底＝判斷題）。分部加總即本模型的 Comparable（在分部層做）"
V["A2"] = "DCF：無槓桿自由現金流（NOPAT＋折舊 − 資本支出＋營運資金；股權報酬不加回）以 WACC 折現到評價日 2026-06-30（期中慣例），終值＝NOPAT(FY2033) ×(1 − g ÷ RONIC) ÷ (WACC − g)，再以 ke 推到目標時點。SOTP：目標時點 NTM EBITDA × 同業 NTM EV/EBITDA"
VR_ = {}
SENS = json.load(open(A.sens, encoding="utf-8")) if A.sens else {}
r = 4
def vput(key, label, vals, fmt="#,##0.00", note=None, bold=False, fill=None):
    global r
    V.cell(r, 1, label)
    if bold: V.cell(r, 1).font = BOLD
    for j, v in enumerate(vals):
        if v is None: continue
        c = V.cell(r, 2 + j, v); c.number_format = fmt[j] if isinstance(fmt, (list, tuple)) else fmt
    if note: V.cell(r, 5, note)
    if fill:
        for cc in range(1, 6): V.cell(r, cc).fill = fill
    if key: VR_[key] = r
    r += 1
V.cell(r, 1, "0.3 結論").font = BOLD; r += 1
hdr(V, r, ["項目", "數值", "對照", "對照", "說明"]); r += 1
CONC = r; r += 18
V.cell(r, 1, "0.0 折現率（CAPM；沿用 v0.2.2 輸入）").font = BOLD; r += 1
hdr(V, r, ["項目", "數值", "", "", "說明"]); r += 1
vput("emkt", "股權市值 E＝現價 × 流通股", ["=px*sh_out"], "#,##0.0")
vput("dmkt", "債務 D（面額）", ["=debt_0"], "#,##0.0")
vput("ke", "股東權益成本 ke＝rf＋β × ERP", ["=rf+beta*erp"], "0.00%")
dn("ke", f"Val!$B${VR_['ke']}")
vput("kdat", "稅後債務成本", ["=kd*(1-tax_w)"], "0.00%")
vput("wacc", "WACC", [f"=(B{VR_['emkt']}*ke+B{VR_['dmkt']}*B{VR_['kdat']})/(B{VR_['emkt']}+B{VR_['dmkt']})"], "0.00%", "v0.2.2 為 11.6%（現價 350.50）", bold=True)
dn("wacc", f"Val!$B${VR_['wacc']}")
r += 1
V.cell(r, 1, "0.1 DCF（無槓桿自由現金流；三情境各一區塊；欄 F＝2H26、H–M＝FY2027–FY2032）").font = BOLD; r += 1
DC = {s: {} for s in "ABC"}
VC = {"2H": "F", 27: "H", 28: "I", 29: "J", 30: "K", 31: "L", 32: "M"}
TT = {"2H": 0.25, 27: 1.0, 28: 2.0, 29: 3.0, 30: 4.0, 31: 5.0, 32: 6.0}
for s in "ABC":
    P = PR_[s]; C = CR[s]; D = DC[s]
    V.cell(r, 1, SCN[s][2]).font = BOLD; V.cell(r, 1).fill = SUB; r += 1
    hdr(V, r, ["項目", "", "", "", "說明", YH["2H"], "", YH[27], YH[28], YH[29], YH[30], YH[31], YH[32]]); r += 1
    def drow(key, label, f, fmt="#,##0.00", note=None):
        global r
        V.cell(r, 1, label)
        for p in FC:
            c = V[f"{VC[p]}{r}"]; c.value = f(p); c.number_format = fmt
        if note: V.cell(r, 5, note)
        D[key] = r; r += 1
    drow("ebit", "營業利益（EBIT）", lambda p: f"=PL!{COL[p]}{P['oi']}")
    drow("nopat", "NOPAT＝(EBIT＋讀法 2 加回的自用 AI 增量成本) ×(1 − 稅率)", lambda p: f"=({VC[p]}{D['ebit']}+rd2*Thesis!{COL[p]}{T['int_cost']})*(1-tax)",
         note="讀法 2（rd2＝1）：自用 AI 增量營運、租金與折舊不計入評價")
    drow("da", "＋折舊（讀法 2 扣除自用 AI 折舊）", lambda p: f"=PL!{COL[p]}{P['da']}-rd2*Thesis!{COL[p]}{T['da_int']}")
    drow("cx", "− 資本支出（讀法 2 扣除自用 AI 資本支出）", lambda p: f"=CF!{COL[p]}{C['cx']}-rd2*(1-ext_sh)*Thesis!{COL[p]}{T['cx_ai']}")
    drow("wc", "＋營運資金變動", lambda p: f"=CF!{COL[p]}{C['wc']}")
    drow("ufcf", "無槓桿自由現金流", lambda p: f"={VC[p]}{D['nopat']}+{VC[p]}{D['da']}-{VC[p]}{D['cx']}+{VC[p]}{D['wc']}", note="股權報酬不加回（視為真實成本；替代＝加回，見報告第 5 節）")
    drow("t", "折現年數（評價日 2026-06-30 起，期中）", lambda p: TT[p], "0.00")
    drow("pv", "現值", lambda p: f"={VC[p]}{D['ufcf']}/(1+wacc)^{VC[p]}{D['t']}")
    r += 1
V.cell(r, 1, "DCF 彙總").font = BOLD; r += 1
hdr(V, r, ["項目", "情境 A", "情境 B", "情境 C", "說明"]); r += 1
S3 = lambda f: [f(s) for s in "ABC"]
vput("pvx", "明確期現值合計（2H26–FY2032）", S3(lambda s: f"=SUM(F{DC[s]['pv']}:M{DC[s]['pv']})"))
vput("tvv", "終值（FY2032 年底）", S3(lambda s: f"=IF(tv_alt=1,M{DC[s]['ufcf']}*(1+g_tv)/(wacc-g_tv),M{DC[s]['nopat']}*(1+g_tv)*(1-g_tv/ronic)/(wacc-g_tv))"),
     note="NOPAT(FY2033) ×(1 − g ÷ RONIC) ÷ (WACC − g)；tv_alt＝1 時改用 FY2032 無槓桿自由現金流 ×(1＋g)")
vput("pvtv", "終值現值（6.5 年）", S3(lambda s: f"={'BCD'['ABC'.index(s)]}{VR_['tvv']}/(1+wacc)^6.5"))
vput("ev", "企業價值 EV", S3(lambda s: f"={'BCD'['ABC'.index(s)]}{VR_['pvx']}+{'BCD'['ABC'.index(s)]}{VR_['pvtv']}"), bold=True)
vput("tvsh", "  其中：終值占 EV", S3(lambda s: f"={'BCD'['ABC'.index(s)]}{VR_['pvtv']}/{'BCD'['ABC'.index(s)]}{VR_['ev']}"), "0.0%")
vput("nc0", "＋淨現金（評價日：現金 155.4 − 債務 101.1）", ["=cash_0-debt_0"] * 3)
vput("hold", "＋持股（Anthropic × 比例 ×(1 − 折價)＋SpaceX＋其他上市權益證券）", ["=ant_val*ant_stake*(1-illiq)+spacex+oth_eq"] * 3, note="與 SOTP 相同；Anthropic 剩餘出資換得持股，評價中性")
vput("dl", "− 類債（應計罰款＋非控制權益）", ["=fines+nci"] * 3)
vput("eqv", "股東權益價值", S3(lambda s: "={0}{1}+{0}{2}+{0}{3}-{0}{4}".format("BCD"["ABC".index(s)], VR_["ev"], VR_["nc0"], VR_["hold"], VR_["dl"])))
vput("shv", "股數（評價日期末稀釋＋強制轉換特別股）", S3(lambda s: f"=PL!E{PR_[s]['she']}+conv_sh"), "0.000")
vput("dps0", "每股價值（評價日）", S3(lambda s: "=MAX(0,{0}{1}/{0}{2})".format("BCD"["ABC".index(s)], VR_["eqv"], VR_["shv"])))
vput("dcf12", "DCF 情境價值推到 12 個月後（× (1＋ke)^t_tgt）", S3(lambda s: "={0}{1}*(1+ke)^t_tgt".format("BCD"["ABC".index(s)], VR_["dps0"])), bold=True, fill=KEY)
r += 1
V.cell(r, 1, "0.2 分部加總（SOTP；目標時點 NTM EBITDA × 同業 NTM EV/EBITDA；＝本模型的 Comparable）").font = BOLD; r += 1
hdr(V, r, ["同業組／倍數", "中位數", "家數", "", "說明"]); r += 1
for grp, lst in PEERL.items():
    keys = [f"pm_{grp}_{t.lower()}" for t, _, _ in lst]
    vput("mult_" + grp, f"{PEERS[grp][0]} NTM EV/EBITDA 中位數", ["=MEDIAN(" + ",".join(keys) + ")", len(lst)], ["0.00", "0"], note=PEERS[grp][1] + "：" + "、".join(t for t, _, _ in lst))
    dn("mult_" + grp, f"Val!$B${VR_['mult_' + grp]}")
vput("mult_adsx", "廣告線實際採用倍數", ["=IF(ads_own=1,mult_ads_own,mult_ads)"], "0.00", note="ads_own＝1 時改用 Alphabet 自身倍數 14.68（Sens）")
dn("mult_adsx", f"Val!$B${VR_['mult_adsx']}")
r += 1
hdr(V, r, ["分部／產品線（NTM EBITDA）", "情境 A", "情境 B", "情境 C", "倍數與說明", "EV：A", "EV：B", "EV：C"]); r += 1
NT = lambda s, k: f"(w_ntm27*PL!H{PR_[s][k]}+(1-w_ntm27)*PL!I{PR_[s][k]})"
LINES = [("yt", "YouTube ads（營收 × m_yt）", lambda s: f"={NT(s, 'yt')}*m_yt", "mult_adsx", "數位廣告"),
         ("net", "Google Network（營收 × m_net）", lambda s: f"={NT(s, 'net')}*m_net", "mult_adsx", "數位廣告"),
         ("subs", "Google subscriptions, platforms, and devices（營收 × m_subs）", lambda s: f"={NT(s, 'subs')}*m_subs", "mult_subs", "訂閱影音"),
         ("search", "Google Search & other（Services EBITDA − 其他三線）", lambda s: f"={NT(s, 'e_gs')}+rd2*svc_sh*(w_ntm27*(Thesis!H{T['int_opex']}+Thesis!H{T['rent']})+(1-w_ntm27)*(Thesis!I{T['int_opex']}+Thesis!I{T['rent']}))-{NT(s, 'yt')}*m_yt-{NT(s, 'net')}*m_net-{NT(s, 'subs')}*m_subs", "mult_adsx", "數位廣告"),
         ("ai", "Cloud AI（Thesis EBITDA）", lambda s: "=w_ntm27*Thesis!H{0}+(1-w_ntm27)*Thesis!I{0}".format(T["ai_ebitda"]), "mult_ai", "AI 基礎設施"),
         ("tpu", "TPU 系統銷售（營收 × tpu_m）", lambda s: f"={NT(s, 'tpu')}*tpu_m", "mult_ai", "AI 基礎設施"),
         ("co", "Cloud 其他（營收 × m_co）", lambda s: f"={NT(s, 'co')}*m_co", "mult_cloud", "非 AI 雲端"),
         ("ob", "Other Bets（EBITDA，負值）", lambda s: f"={NT(s, 'e_ob')}", "mult_ai", "沿用 AI 倍數（v0.2.2）"),
         ("abc", "Alphabet-level activities（EBITDA，負值）", lambda s: f"={NT(s, 'e_abc')}+rd2*(1-svc_sh)*(w_ntm27*(Thesis!H{T['int_opex']}+Thesis!H{T['rent']})+(1-w_ntm27)*(Thesis!I{T['int_opex']}+Thesis!I{T['rent']}))", "mult_adsx", "數位廣告（v0.2.2）")]
SOTP0 = r
for key, lab, f, mult, mlab in LINES:
    V.cell(r, 1, lab)
    for j, s in enumerate("ABC"):
        V.cell(r, 2 + j, f(s)).number_format = "#,##0.00"
        V.cell(r, 6 + j, f"={'BCD'[j]}{r}*{mult}").number_format = "#,##0.0"
    V.cell(r, 5, f"=\"{mlab} \"&TEXT({mult},\"0.00\")&\"×\"")
    VR_["s_" + key] = r; r += 1
SOTP1 = r - 1
vput("sev", "分部加總 EV", [f"=SUM(F{SOTP0}:F{SOTP1})", f"=SUM(G{SOTP0}:G{SOTP1})", f"=SUM(H{SOTP0}:H{SOTP1})"], "#,##0.0", bold=True)
vput("sebt", "  分部 NTM EBITDA 合計", [f"=SUM({c}{SOTP0}:{c}{SOTP1})" for c in "BCD"], "#,##0.0")
vput("snc", "＋淨現金（FY2027 年底，CF 頁）", S3(lambda s: f"=CF!H{CR[s]['nc']}"), note="目標時點附近的資產負債；含 2H26–FY2027 的自由現金流、股利與回購")
vput("shold", "＋持股（同 DCF）", [f"=B{VR_['hold']}"] * 3)
vput("sdl", "− 類債", [f"=B{VR_['dl']}"] * 3)
vput("seq", "股東權益價值", S3(lambda s: "={0}{1}+{0}{2}+{0}{3}-{0}{4}".format("BCD"["ABC".index(s)], VR_["sev"], VR_["snc"], VR_["shold"], VR_["sdl"])))
vput("ssh", "股數（FY2027 年底稀釋＋強制轉換特別股）", S3(lambda s: f"=PL!H{PR_[s]['she']}+conv_sh"), "0.000")
vput("sotp", "分部加總每股價值（目標時點）", S3(lambda s: "=MAX(0,{0}{1}/{0}{2})".format("BCD"["ABC".index(s)], VR_["seq"], VR_["ssh"])), bold=True, fill=KEY)
vput("simp", "  隱含 EV ÷ NTM EBITDA", S3(lambda s: "={0}{1}/{0}{2}".format("BCD"["ABC".index(s)], VR_["sev"], VR_["sebt"])), "0.00")
# 0.3 結論
r_end = r; r = CONC
W3 = lambda row: f"tp_wA*B{row}+tp_wB*C{row}+tp_wC*D{row}"
vput("tpdcf", "DCF 目標價（情境加權）", [f"={W3(VR_['dcf12'])}"], "#,##0.00", fill=KEY)
vput("tpcomp", "Comparable 目標價（市場主流同業法）", [f"={W3(VR_['sotp'])}"], "#,##0.00", note="＝分部加總（SOTP）情境加權；各分部用各自同業倍數")
vput("tp", "目標價（DCF 與 Comparable 加權）", ["=w_dcf*B{0}+(1-w_dcf)*B{1}".format(r - 2, r - 1)], "#,##0.00", bold=True, fill=KEY,
     note='="DCF 權重 "&TEXT(w_dcf,"0%")&"、分部加總 "&TEXT(1-w_dcf,"0%")&"（In 頁 w_dcf，Andy 可改）"')
vput("px", "現價（2026-10-26 收盤）", ["=px"], "#,##0.00", note="StockAnalysis 歷史股價（In 頁 px）")
vput("cons", "市場共識目標價", ["=cons_tp"], "#,##0.00", note="StockAnalysis 61 位分析師平均（2026-10-26）；只作對照")
vput("up", "上檔空間：加權目標價", [f"=B{VR_['tp']}/px-1"], "0.0%", bold=True, fill=KEY)
vput("upc", "上檔空間：Comparable 目標價", [f"=B{VR_['tpcomp']}/px-1"], "0.0%")
U = f"B{VR_['up']}"
vput("tier", "上檔空間級距", [f'=IF({U}>0.5,"大於 50%",IF({U}>0.25,"25%–50%",IF({U}>=0,"0%–25%",IF({U}>=-0.25,"0%～−25%",IF({U}>=-0.5,"−25%～−50%","小於 −50%")))))'], "@",
     note="以「上檔空間：加權目標價」分六級（Andy 2026-10-25 14:00）")
vput("rec", "投資建議", [f'=IF({U}>=0.5,"強力買進",IF({U}>=0.25,"買進",IF({U}>-0.25,"中立",IF({U}>-0.5,"放空","強力放空"))))'], "@", bold=True, fill=KEY,
     note="五級固定規則：≥ +50% 強力買進、+25%～+50% 買進、−25%～+25% 中立、−50%～−25% 放空、≤ −50% 強力放空")
vput("band", "價格帶（DCF 不如預期～順利）", [f"=D{VR_['dcf12']}", f"=B{VR_['dcf12']}"], "#,##0.00", note="B＝C 情境 DCF、C＝A 情境 DCF（皆推到 12 個月後）")
vput("pess", "悲觀情境價（C 情境 DCF）", [f"=D{VR_['dcf12']}"], "#,##0.00", note="地圖「悲觀情境空間」＝本列 ÷ 現價 − 1")
vput("scen", "情境目標價（A／B／C；DCF 與分部加總加權）", S3(lambda s: "=w_dcf*{0}{1}+(1-w_dcf)*{0}{2}".format("BCD"["ABC".index(s)], VR_["dcf12"], VR_["sotp"])), "#,##0.00")
vput("scenup", "情境上檔空間（A／B／C）", S3(lambda s: "={0}{1}/px-1".format("BCD"["ABC".index(s)], VR_["scen"])), "0.0%")
vput("tvsh2", "DCF 終值占 EV（A／B／C）", S3(lambda s: "={0}{1}".format("BCD"["ABC".index(s)], VR_["tvsh"])), "0.0%")
_r2 = SENS.get("rd2|高") if "SENS" in globals() else None
vput("rd2tp", "讀法 2（自用 AI 增量投資價值中性，C16）加權目標價與上檔空間", [round(_r2["tp"], 6) if _r2 else None, f"=B{r}/px-1" if _r2 else None], ["#,##0.00", "0.0%"],
     note="以 rd2＝1 整本重算的值（make_v03.sh；Sens 頁同列）；主值取讀法 1 或 2 待 Andy 決定（v0.2.2 主值＝讀法 1）")
vput("meth", "目標價方法", ["DCF（無槓桿 FCF、WACC、終值價值驅動）50%＋分部加總（各分部同業 NTM EV/EBITDA）50%；情境 25／50／25"], "@")
vput("wsum", "檢查：情境權重合計（應為 1）", ["=ROUND(tp_wA+tp_wB+tp_wC,6)"], "0.000")
assert r <= CONC + 18, r
V.column_dimensions["A"].width = 62
for c in "BCD": V.column_dimensions[c].width = 13
V.column_dimensions["E"].width = 60
for c in "FGHIJKLM": V.column_dimensions[c].width = 12

# =====================================================================================================
# Recon：口徑對照
RC["A1"] = "Recon — 口徑對照：合併層對市場共識（MarketScreener，2026-10-26）、逐分部對共識與公司指引、v0.3 對 v0.2.2。只作對照，不作校準（工作單第 4 節）"
r = 3
RC.cell(r, 1, "1. 合併層：本模型 vs 市場共識").font = BOLD; r += 1
hdr(RC, r, ["項目", "年度", "情境 A", "情境 B", "情境 C", "共識", "B ÷ 共識 − 1", "說明"]); r += 1
RCR = {}
for k, lab, src in (("rev", "營收", "rev"), ("ebit", "營業利益", "oi"), ("eps", "EPS（GAAP，含權益證券損益）", "eps"), ("capex", "資本支出", "cx"), ("fcf", "自由現金流", "fcf")):
    for y in (26, 27, 28):
        RC.cell(r, 1, lab); RC.cell(r, 2, f"FY20{y}")
        for j, s in enumerate("ABC"):
            col = COL["26" if y == 26 else y]
            ref = f"PL!{col}{PR_[s][src]}" if src in ("rev", "oi", "eps") else f"CF!{col}{CR[s][src]}"
            RC.cell(r, 3 + j, f"={ref}").number_format = "#,##0.00"
        RC.cell(r, 6, f"=c_{k}_{y}").number_format = "#,##0.00"
        RC.cell(r, 7, f"=D{r}/F{r}-1").number_format = "0.0%"
        if k == "fcf": RC.cell(r, 8, "本模型扣融資租賃本金；共識＝營運現金 − 資本支出")
        if k == "eps" and y == 26: RC.cell(r, 8, "含上半年權益證券利益 135.9；StockAnalysis 排除後的共識 11.89（In 頁 c_eps26_sa）")
        RCR[f"{k}{y}"] = r; r += 1
RC.cell(r, 1, "EPS（不含權益證券損益）FY2026 vs StockAnalysis 共識"); RC.cell(r, 2, "FY2026")
for j, s in enumerate("ABC"): RC.cell(r, 3 + j, f"=PL!G{PR_[s]['epsx']}").number_format = "0.00"
RC.cell(r, 6, "=c_eps26_sa").number_format = "0.00"; RC.cell(r, 7, f"=D{r}/F{r}-1").number_format = "0.0%"; r += 2
RC.cell(r, 1, "2. 逐分部、逐產品線：本模型（情境 B）vs 共識與公司指引").font = BOLD; r += 1
hdr(RC, r, ["分部／產品線", "FY2026E", "FY2027E", "FY2028E", "共識", "公司指引", "說明"]); r += 1
PB = PR_["B"]
for key, lab in (("search", "Google Search & other"), ("yt", "YouTube ads"), ("net", "Google Network"), ("subs", "Google subscriptions, platforms, and devices"),
                 ("gs", "Google Services"), ("cloud", "Google Cloud"), ("ai", "  Cloud AI（命題子列）"), ("tpu", "  TPU 系統銷售"), ("co", "  Cloud 其他"), ("ob", "Other Bets")):
    RC.cell(r, 1, lab)
    for j, p in enumerate(("26", 27, 28)): RC.cell(r, 2 + j, f"=PL!{COL[p]}{PB[key]}").number_format = "#,##0.00"
    RC.cell(r, 5, "未找到")
    RC.cell(r, 6, {"tpu": "2026 認列一小部分、大部分在 2027（2026Q2 法說，事實總帳 guid.tpuRev）", "cloud": "仍供不應求（2026Q2）；無營收指引"}.get(key, "無（公司不給營收指引）"))
    RC.cell(r, 7, "StockAnalysis、MarketScreener 免費頁只有合併層；分部共識需付費資料，本次網路搜尋額度用盡，未能補（資料缺口：未找到）" if key == "search" else "")
    r += 1
RC.cell(r, 1, "Google Cloud 積壓（RPO）約 50% 在 24 個月內認列 → 2026H2–2028H1 至少可認列"); RC.cell(r, 2, "=513.9*0.5").number_format = "#,##0.0"
RC.cell(r, 7, "10-Q 2026Q2：Cloud RPO 513.9、約 50% 於 24 個月內（事實總帳 rpo.cloud、rpo.share24m）[Interested-party]；只作驗證列，不反推 k 或 MW"); r += 1
RC.cell(r, 1, "本模型 Google Cloud 營收 2026H2–2028H1（情境 B）"); RC.cell(r, 2, f"=PL!F{PB['cloud']}+PL!H{PB['cloud']}+0.5*PL!I{PB['cloud']}").number_format = "#,##0.0"
RC.cell(r, 3, f"=B{r}/B{r - 1}").number_format = "0.00"; RC.cell(r, 7, "C 欄＝本模型 ÷ 積壓可認列（> 1 表示模型營收高於積壓可支撐的下限，差額需新簽約）"); r += 1
RC.cell(r, 1, "FY2026 資本支出：本模型 vs 公司指引 195–205"); RC.cell(r, 2, f"=CF!G{CR['B']['cx']}").number_format = "#,##0.0"; RC.cell(r, 3, "=capex_g").number_format = "#,##0.0"
RC.cell(r, 7, "C10：首期＝MAX(公式, 指引)"); r += 1
RC.cell(r, 1, "FY2027 資本支出：本模型 vs 公司「顯著增加」（無數字）vs 共識"); RC.cell(r, 2, f"=CF!H{CR['B']['cx']}").number_format = "#,##0.0"; RC.cell(r, 3, "=c_capex_27").number_format = "#,##0.0"
RC.cell(r, 7, "本模型 FY2027 由 MW 路徑推導（C10 基準速度 = 2026 指引隱含量延續）；低於共識是已知差異（v0.2.2 同）"); r += 2
RC.cell(r, 1, "3. v0.3 vs v0.2.2：目標價與主要差異來源").font = BOLD; r += 1
hdr(RC, r, ["項目", "數值", "說明"]); r += 1
RCV = {}
def rcv(key, lab, f, fmt="#,##0.00", note=""):
    global r
    RC.cell(r, 1, lab); RC.cell(r, 2, f).number_format = fmt; RC.cell(r, 3, note); RCV[key] = r; r += 1
rcv("v22", "v0.2.2 加權目標價（讀法 1；25／50／25）", "=tpw_v022", note="v0.2.2：保守 175.08／基準 146.09／積極 122.89（容量情境；讀法 2 為 218.79）")
rcv("v03", "v0.3 加權目標價", f"=Val!B{VR_['tp']}")
rcv("dif", "差額", f"=B{r - 1}-B{r - 2}")
RC_ATT = r          # 拆解列（make_v03.sh 以改輸入重算後寫入）
r += 8
RC.column_dimensions["A"].width = 70
for c in "BCDEFG": RC.column_dimensions[c].width = 13
RC.column_dimensions["H"].width = 70

# =====================================================================================================
# Sens：龍捲風（一次改一個輸入，整本重算後的加權目標價；make_v03.sh 產生）
SENS_CASES = [
    ("ext_sh", "對外比例（C23）", 0.10, 0.50, "10%／50%（±20pt；總 MW 不變）"),
    ("pace", "AI MW 新增速度", "lo", "hi", "v0.2.2 保守／積極速度"),
    ("kl_px", "k：長約價格（低＝判斷倍數 0.60）", "lo", None, "長約 0.75 → 0.60"),
    ("ks_px", "k：現貨價格（高＝GCP 牌價）", None, "hi", "2.82 → 4.858 US$/GPU-hr"),
    ("tpu_f", "TPU 係數（收入與資本支出雙邊 × 0.7）", "lo", None, "C2 敏感度；同時改 tpu_cf"),
    ("life_it", "AI IT 折舊年限", "lo", "hi", "5／7 年"),
    ("dC", "Search 情境 C 分流幅度", "lo", "hi", "10%／30%（只影響 C 情境）"),
    ("gl_search", "Search 長期成長", "lo", "hi", "5%／11%"),
    ("inc_m", "Search 差額增量 EBITDA 率", "lo", "hi", "60%／80%"),
    ("gl_co", "Cloud 其他長期成長", "lo", "hi", "5%／12%"),
    ("gl_yt", "YouTube 長期成長", "lo", "hi", "5%／11%"),
    ("beta", "β（WACC）", "lo", "hi", "1.21／1.367"),
    ("g_tv", "永續成長 g", "lo", "hi", "2%／4%"),
    ("ronic", "永續期新投資報酬（RONIC）", None, "hi", "12% → 20%"),
    ("tv_alt", "終值改用 FY2032 自由現金流", None, "hi", "價值驅動 → UFCF ×(1＋g)"),
    ("mult_ai", "Cloud AI／TPU 倍數", "lo", "hi", "4×／10×"),
    ("ads_own", "廣告線改用 Alphabet 自身倍數 14.68×", None, "hi", "9.11× → 14.68×"),
    ("illiq", "非上市持股折價", "lo", "hi", "0%／40%"),
    ("w_dcf", "DCF 權重", "lo", "hi", "30%／70%"),
    ("bb_t", "回購目標", "lo", "hi", "0／45.7（2025 全年）"),
    ("rd2", "讀法 2：自用 AI 增量投資價值中性（C16）", None, "hi", "讀法 1 → 讀法 2"),
    ("mb_ttm", "分部利益率起點改用近四季", None, "hi", "1H26 → 近四季（v0.2.2 口徑）"),
]
SE["A1"] = "Sens — 一次改一個輸入（In 頁低／高值），整本活頁簿以 LibreOffice 重算後的加權目標價（Val 頁「目標價（DCF 與 Comparable 加權）」）；由 alphabet/make_v03.sh 產生，建置時寫入（數值）"
SE["A2"] = "相對基準＝情境目標價 ÷ 基準加權目標價 − 1；龍捲風依絕對值排序（HTML）。基準與各案例的其他輸入相同"
hdr(SE, 3, ["驅動", "設定", "方向", "加權目標價", "相對基準", "鍵", "說明"])
r = 4
SE_ROWS = []
for key, lab, lo, hi, note in SENS_CASES:
    for side, val in (("低", lo), ("高", hi)):
        if val is None: continue
        SE.cell(r, 1, f"{lab}：{'低' if side == '低' else '高'}"); SE.cell(r, 2, note); SE.cell(r, 3, side); SE.cell(r, 6, key)
        res = SENS.get(f"{key}|{side}")
        if res is not None:
            c = SE.cell(r, 4, res["tp"]); c.number_format = "#,##0.00"; c.font = BLUE
            SE.cell(r, 5, f"=D{r}/Val!$B${VR_['tp']}-1").number_format = "+0.0%;-0.0%;0.0%"
        SE.cell(r, 7, "重算值（make_v03.sh）")
        SE_ROWS.append((key, side, val, r)); r += 1
SE_END = r - 1
r += 1
SE.cell(r, 1, "基準加權目標價（Val 頁）"); SE.cell(r, 4, f"=Val!B{VR_['tp']}").number_format = "#,##0.00"; SE_BASE = r; r += 1
SE.cell(r, 1, "敏感度重算時的基準值（應與上一列相同；不同表示敏感度已過期）")
if "base|base" in SENS:
    c = SE.cell(r, 4, SENS["base|base"]["tp"]); c.number_format = "#,##0.00"; c.font = BLUE
    SE.cell(r, 5, f"=D{r}-D{r - 1}").number_format = "0.000"
SE_STALE = r; r += 1
SE.column_dimensions["A"].width = 46; SE.column_dimensions["B"].width = 36; SE.column_dimensions["G"].width = 24
# Recon 拆解（以改輸入重算）
ATT = [("att_w", "w_dcf=0.45", "評價法權重：DCF 50% → v0.2.2 的 45%（其餘不變）"),
       ("att_px", "px=350.5", "現價回到 v0.2.2 的 350.50（只影響 WACC 的權益權重與回購價）")]
RC.cell(RC_ATT, 1, "拆解（以改輸入重算；make_v03.sh）").font = BOLD
for i, (k, setting, lab) in enumerate(ATT):
    rr = RC_ATT + 1 + i
    RC.cell(rr, 1, lab)
    res = SENS.get(f"{k}|att")
    if res is not None:
        c = RC.cell(rr, 2, f"=Val!B{VR_['tp']}-{res['tp']}"); c.number_format = "#,##0.00"
    RC.cell(rr, 3, f"v0.3 − 改成 {setting} 後的目標價（正值＝本版設定使目標價較高）")
rr = RC_ATT + 1 + len(ATT)
RC.cell(rr, 1, "其餘：營收結構（分部＋Search 情境軸）、AI 折舊與自用 AI 成本進分部損益、DCF 終值口徑（價值驅動）、SOTP 以各分部同業倍數取代 EV/EBITDA 錨定 2029、情境定義改為 Search")
RC.cell(rr, 2, f"=B{RCV['dif']}-" + "-".join(f"N(B{RC_ATT + 1 + i})" for i in range(len(ATT)))).number_format = "#,##0.00"
RC.cell(rr, 3, "兩版引擎不同，無法逐項以同一引擎重算；本列為殘差（報告第 6 節逐項說明方向）")
RC_ATT_END = rr

# =====================================================================================================
# Sources
SO["A1"] = "Sources — 來源（編號、文件、日期、網址、標記）"
hdr(SO, 3, ["編號", "文件", "日期", "網址", "標記", "用途"])
SRCS = [
    ("Alphabet Form 10-K FY2025（經查核）", "2026-02-05", "https://www.sec.gov/Archives/edgar/data/1652044/000165204426000018/goog-20251231.htm", "Verified", "FY2024、FY2025 分部與產品線、合併損益、現金流（經事實總帳）"),
    ("Alphabet 2026Q2 財報新聞稿（8-K 附件 99.1）", "2026-07-22", "https://www.sec.gov/Archives/edgar/data/1652044/000165204426000066/googexhibit991q22026.htm", "Interested-party", "2026Q2 與比較季"),
    ("Alphabet 2026Q1 財報新聞稿（8-K 附件 99.1）", "2026-04", "https://www.sec.gov/Archives/edgar/data/1652044/000165204426000043/googexhibit991q12026.htm", "Interested-party", "2026Q1"),
    ("Alphabet 2025Q3 財報新聞稿（8-K 附件 99.1）", "2025-10-29", "https://www.sec.gov/Archives/edgar/data/1652044/000165204425000087/googexhibit991q32025.htm", "Interested-party", "2024Q3 比較季"),
    ("Alphabet 2026Q2 10-Q", "2026-07-23", "https://www.sec.gov/cgi-bin/browse-edgar?action=getcompany&CIK=1652044&type=10-Q", "Interested-party", "資產負債、在建工程、租賃、債務、股數"),
    ("Alphabet 2026Q2 法說逐字稿（Benzinga 轉錄，二手）", "2026-07-22", "https://www.benzinga.com/news/26/09/61834150", "Interested-party", "2026 資本支出指引 195–205、TPU 營收時點、供不應求"),
    ("事實總帳 alphabet/data/alphabet_facts_20261008.json（MAG v0.1a）", "2026-10-08", "repo 內", "各筆標記", "Q 頁全部實績與多數輸入"),
    ("Tokenomics 快照 alphabet/data/tokenomics_snapshot_v21.4.json（20261024_Tokenomics_v21.4.xlsx，master 3ac70ac）", "2026-10-25", "repo 內（YinchenChang/Tokenomics）", "Derived", "IF_HoldEcon、IF_GPUhrEcon、IF_OpexGW、IF_CapexIT／Facility、上限檢查"),
    ("k 證據 alphabet/data/kevidence_alphabet_20261008.json；legacy_v0.2/kspec.py", "2026-10-22", "repo 內", "Verified／Assumed", "k_長約、k_現貨"),
    ("Epoch AI〈AI Chip Owners〉cumulative by designer", "2026-05-12", "https://epoch.ai/data/ai_chip_owners_cumulative_by_designer.csv", "Derived", "AI 在役 MW（期初、歷史）"),
    ("Silicon Data H100 Index", "2026-10-06", "https://www.silicondata.com/products/silicon-index/h100", "Verified", "k_現貨"),
    ("StockAnalysis GOOGL 歷史股價", "2026-10-26", "https://stockanalysis.com/stocks/googl/history/", "Verified", "現價 316.80（2026-10-26 收盤）"),
    ("StockAnalysis GOOGL 預測頁（S&P Global）", "2026-10-26", "https://stockanalysis.com/stocks/googl/forecast/", "Interested-party", "共識目標價 429.12、FY2026 non-GAAP EPS 11.89"),
    ("MarketScreener Alphabet 財務預估（S&P Capital IQ 系）", "2026-10-26", "https://www.marketscreener.com/quote/stock/ALPHABET-INC-24203373/finances/", "Interested-party", "FY2026–2028 合併共識"),
    ("共用同業檔 alphabet/data/peers_ads_20261009.json、peers_cloud_20261008.json", "2026-10-07", "repo 內（StockAnalysis EV ÷ MarketScreener NTM EBITDA）", "Derived", "SOTP 倍數"),
    ("共識檔 alphabet/data/consensus_googl_20261008.json", "2026-10-08", "repo 內", "Interested-party", "v0.2.2 時的共識（對照）"),
    ("MAG v0.2.2 成品與交接檔（legacy_v0.2/、dist/_待刪除/、docs/handoff/）", "2026-10-25", "repo 內", "Derived", "沿用參數與 v0.2.2 目標價"),
    ("MAG 命題與驅動對照表 r1（C1–C23）", "2026-10-08", "mag/docs/plan/20261008_MAG_命題與驅動對照表_r1.md", "—", "已決定的口徑"),
    ("Gartner 新聞稿：2026 年搜尋引擎查詢量因 AI 聊天機器人減少 25%（預測）", "2024-02-19", "https://www.gartner.com/en/newsroom（本次未重新查核）", "Interested-party", "情境 C 分流幅度的參考（第三方預測）"),
    ("Apple Eddy Cue 美國司法部 Google 案證詞（媒體轉述）", "2025-05-07", "（媒體轉述；本次未重新查核）", "Interested-party", "情境 C：Safari 搜尋 2025-04 首次下滑"),
]
for i, (doc, dt, url, tag, use) in enumerate(SRCS, 1):
    for j, v in enumerate((f"S{i:02d}", doc, dt, url, tag, use), 1): SO.cell(3 + i, j, v)
for col, w in zip("ABCDEF", [6, 80, 12, 70, 18, 60]): SO.column_dimensions[col].width = w

# =====================================================================================================
# README
RD["A1"] = "Alphabet（GOOGL）個股模型 v0.3（2026-10-27）— 命題模組＋官方營收分部 → 損益表"
RD["A3"] = "命題：AI 改變搜尋之後，Google Search 的變現能否守住，同時 Google Cloud（含 TPU）能否把翻倍的資本支出變成營收；資本支出與 2026 年股權發行後的每股價值是否高於現價。"
RD["A4"] = "核心問題：Search & other 在 C 情境（AI 助理分流、每次查詢變現下降）下 FY2028 營收與利益率會掉多少，Cloud 的 AI 營收能否抵銷；加計 AI 折舊與股權發行後的每股價值。"
VB = lambda k, c="B": f"Val!{c}{VR_[k]}"
PB_, PC_ = PR_["B"], PR_["C"]
RD["A5"] = f'="投資建議："&{VB("rec")}&"（加權目標價 US$"&TEXT({VB("tp")},"0.00")&"，上檔空間 "&TEXT({VB("up")},"+0.0%;-0.0%")&"；現價 "&TEXT(px,"0.00")&"，2026-10-26）"'
RD["A6"] = (f'="重點觀察 1：目標價 "&TEXT({VB("tp")},"0")&"＝DCF "&TEXT({VB("tpdcf")},"0")&"、分部加總 "&TEXT({VB("tpcomp")},"0")&" 各半；情境 A／B／C 為 "'
            f'&TEXT({VB("scen")},"0")&"／"&TEXT({VB("scen", "C")},"0")&"／"&TEXT({VB("scen", "D")},"0")&"，DCF 終值占 EV "&TEXT({VB("tvsh2", "C")},"0%")&"（B）。市場共識目標價 "&TEXT(cons_tp,"0")&"。"')
RD["A7"] = (f'="重點觀察 2：C 情境 FY2028 Search & other "&TEXT(PL!I{PC_["search"]},"0.0")&"（比 B 少 "&TEXT(1-PL!I{PC_["search"]}/PL!I{PB_["search"]},"0.0%")'
            f'&"），Google Services 營業利益率 "&TEXT(PL!I{PC_["om_gs"]},"0.0%")&"（B "&TEXT(PL!I{PB_["om_gs"]},"0.0%")&"）；同年 Cloud AI 營收 "&TEXT(PL!I{PB_["ai"]},"0.0")'
            f'&"、Google Cloud 營業利益 "&TEXT(PL!I{PB_["oi_cl"]},"0.0")&"，Cloud 營業利益比 FY2025 增加 "&TEXT(PL!I{PB_["oi_cl"]}-PL!D{PB_["oi_cl"]},"0.0")'
            f'&"，抵銷 C 情境 Services 營業利益損失 "&TEXT(PL!I{PB_["oi_gs"]}-PL!I{PC_["oi_gs"]},"0.0")&" 的 "&TEXT((PL!I{PB_["oi_cl"]}-PL!D{PB_["oi_cl"]})/(PL!I{PB_["oi_gs"]}-PL!I{PC_["oi_gs"]}),"0%")&"。"')
G_ = f"CF!C{GATE}"
RD["A8"] = (f'="重點觀察 3：AI 折舊進損益——情境 B 新增 AI 折舊 FY2028 "&TEXT(PL!I{PB_["da_ai"]},"0.0")&"，合併營業利益率由 FY2025 "&TEXT(PL!D{PB_["om"]},"0.0%")&" 變為 FY2028 "&TEXT(PL!I{PB_["om"]},"0.0%")'
            f'&"；收支關卡（FY2026–2030 自由現金流扣股利合計）B "&TEXT({G_},"0.0")&"，"&CF!C{GATE + 3}&"。"')
r = 10
RD.cell(r, 1, "驅動 → 推導對照表（每頁輸入、推導量、因果方向）").font = BOLD; r += 1
hdr(RD, r, ["頁", "輸入（In 頁鍵）", "推導量", "因果方向"]); r += 1
DRV = [("Thesis", "pace、ext_sh、mw_tot0、k（kl_px、ks_px、long_sh）、Tokenomics IF_HoldEcon／OpexGW／CapexIT", "Cloud AI 營收與 EBITDA、AI 資本支出、新增 AI 折舊、自用 AI 增量成本", "MW → 收入（× 每 MW 持有成本 × k）；MW → 資本支出 → 折舊 → 分部營業利益"),
       ("Seg", "g0（Q 頁近四季）、gl_*、conv_n、dA／mA／dC／mC、tpu_*、ob_rev_*", "四條產品線、三個分部、合併營收；非 AI 資本支出與折舊", "近四季年增 → 長期值；Search＝B 路徑 ×(1 − 分流)×(1＋變現變動)"),
       ("PL", "m_svc、inc_m、m_co、tpu_m、abc_rate、svc_sh、da_base、tax", "分部 EBITDA、分部營業利益、合併營業利益、淨利、EPS", "分部營收 → 分部 EBITDA − 折舊 → 營業利益 → 業外、稅 → EPS"),
       ("CF", "sbc_r、wc_*、dps、bb_t、min_cash、mat_*、anth_*", "自由現金流、股利、回購、新債、淨現金、收支關卡", "淨利＋折舊 − 資本支出 → 自由現金流 → 股利 → 回購／舉債"),
       ("Val", "rf、beta、erp、kd、g_tv、ronic、同業倍數、持股、tp_w*、w_dcf", "DCF、分部加總、目標價、上檔空間、投資建議", "無槓桿 FCF 折現＋分部 EBITDA × 倍數 → 情境加權 → 50／50"),
       ("Sens", "In 頁低／高值", "加權目標價變動", "一次改一個輸入、整本重算")]
for row in DRV:
    for j, v in enumerate(row, 1): RD.cell(r, j, v)
    r += 1
r += 1
RD.cell(r, 1, "與 v0.2.2 的差異").font = BOLD; r += 1
for t in ("營收主軸由「AI 雲端 MW × 每 MW 收入＋其餘成長率線」改為官方分部（Google Services 四條產品線、Google Cloud、Other Bets、Alphabet 層級）；Cloud AI 只是 Google Cloud 的子列。",
          "情境軸由「AI 容量（保守／基準／積極）」改為 Search（A 擴大變現／B 近四季成長收斂／C 查詢分流）；AI 容量三情境相同，改列 Sens。",
          "AI 資本支出接到折舊與分部營業利益（對外 → Cloud、自用 → Services 與 Alphabet 層級），不只接現金。",
          "評價由 DCF 45%＋EV/EBITDA 55%（錨定 2029）改為 DCF 50%＋分部加總 50%（各分部同業倍數）；評等改五級（上檔空間）。",
          "融資瀑布、站點、租賃承諾、RPO 等 CRWV 頁面移除（舊引擎保留在 legacy_v0.2/）。"):
    RD.cell(r, 1, "・" + t); r += 1
r += 1
RD.cell(r, 1, "工作表：README → In → Q → Thesis → Seg → PL → CF → Val → Recon → Sens → Sources。藍字＝輸入、黑字＝公式；黃底＝判斷題。建置：bash alphabet/make_v03.sh").font = BOLD; r += 1
RD.cell(r, 1, "數字標記：Verified（經查核）／Interested-party（公司或賣方自述）／Analogy／Assumed（給區間）／Derived（換算）。本模型是研究整理，不是投資建議以外的保證。")
RD.column_dimensions["A"].width = 120; RD.column_dimensions["B"].width = 60; RD.column_dimensions["C"].width = 60; RD.column_dimensions["D"].width = 60
for rr in range(3, 9): RD.cell(rr, 1).alignment = Alignment(wrap_text=True, vertical="top")

# 列號對照（make_v03.sh、html.json 撰寫用）
if os.environ.get("ROWMAP"):
    json.dump({"T": T, "SR": SR, "PR": PR_, "CR": CR, "VR": VR_, "DC": DC, "GATE": GATE, "SUMROW": SUMROW, "SE_ROWS": SE_ROWS, "SE_BASE": SE_BASE,
               "SE_STALE": SE_STALE, "IN_ROW": IN_ROW, "RCR": RCR, "RCV": RCV, "RC_ATT": RC_ATT, "QR": QR},
              open(os.environ["ROWMAP"], "w", encoding="utf-8"), ensure_ascii=False, indent=0)
wb.save(A.out)
print("saved", A.out)
