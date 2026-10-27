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
inp("cash0", "現金＋債權類有價證券（2026-06-30）", "US$B", None, 155.411, None, "Derived",
    "10-Q 2026Q2：現金 55.911＋有價證券 186.563 − 有價權益證券 87.063（事實總帳 bs.cash、bs.mktsec、bs.mktEq）", V22 + " defaults.cash（權益證券改列持股）")
inp("cash_ye25", "現金＋債權類有價證券（2025-12-31）", "US$B", None, 120.53, None, "Derived", "10-K 2025：30.708＋96.135 − 6.313（" + V22 + " ytdActual.cash1231）")
inp("debt0", "債務面額（2026-06-30）", "US$B", None, fv("debt.face.2026q2") / 1000 if "debt.face.2026q2" in F else 101.085, None, "Interested-party", "10-Q 2026Q2 附註 6（事實總帳 debt.face）")
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
inp("g", "永續成長率（FY2032 後）", "%", 0.02, 0.03, 0.04, "Assumed", "工作單第 2 節 Val 0.1（g 3%）", judg=True)
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

sec("D. 命題模組：AI 容量（沿用 v0.2.2；總 AI MW＝對外 MW ÷ 對外比例）")
inp("es", "對外比例（Google Cloud 客戶 MW ÷ AI 總 MW）", "%", 0.25, 0.30, 0.40, "Assumed", fs("mw.extShare.2026q2") if "mw.extShare.2026q2" in F else "事實總帳 mw.extShare（v0.1a 區間 25%–40%）",
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
inp("am27", "2027 新增中 GB300 錨占比（Ironwood＋GB300；其餘 VR200 錨）", "%", None, 0.5, None, "Assumed", V22 + "（2027 Ironwood 40%＋GB300 10%）", "2026 新增全為 GB300 錨（Ironwood 75%＋GB300 25%）；2028 起全為 VR200 錨（下一代 TPU＋VR200）")
inp("tpu_sh", "TPU 占 AI MW 比例", "%", None, 0.75, None, "Derived", "Epoch 2025 年新增功耗 TPU 約 77%（" + V22 + " pricing.mixNote）", "只在 TPU 係數不等於 1 時有作用")
inp("tpu_f", "TPU 每 MW 收入錨係數（相對同期 NVIDIA 世代）", "x", 0.7, 1.0, 1.0, "Assumed", "C2：基準 × 1.0、敏感度 × 0.7（" + V22 + " pricing.customFactor）", "Tokenomics v21.4 NonNV（IF_NonNV_Hold_TPU 0.8）只登錄、不切換（工作單：Tokenomics 本版維持 v21.4）")
inp("tpu_cf", "TPU 每 MW IT 資本支出係數", "x", 0.7, 1.0, 1.0, "Assumed", V22 + " pricing.customCapexFactor", "雙邊敏感度時與 tpu_f 同乘 0.7")

sec("E. 命題模組：每 MW 價格（k＝證據價格 ÷ Tokenomics 同世代成本；沿用 v0.2.2 kspec.py）")
inp("kl_px", "長約證據價格（判斷倍數 0.75 × v5.27 VR200 12.762）", "US$m/MW-年", 7.6572, 9.5715, None, "Assumed",
    "legacy_v0.2/company.json pricing.kEvidence long.base／long.low（data/kevidence_alphabet_20261008.json kspecInputs）", "不是觀察到的市場價格；分母＝VR200 IF_HoldEcon")
inp("ks_px", "現貨證據價格（Silicon Data H100 指數）", "US$/GPU-hr", None, 2.82, 4.858, "Verified",
    "Silicon Data H100 Index（2026-10-06）https://www.silicondata.com/products/silicon-index/h100；高＝GCP a3-highgpu-8g 3 年預留 4.858（Interested-party）", "分母＝IF_GPUhrEcon｜H100；C1 三家一致")
inp("ls", "長約占比", "%", 0.4, 0.6, 0.8, "Assumed", V22 + " pricing.longShare（Cloud 積壓以多年期大客戶為主）")
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
inp("phi0", "資本支出當期折舊比例（在役時點與半年慣例）", "x", 0.0, 0.25, 0.5, "Assumed", "判斷：支出平均於期中、約半年後上線 → 當期 25%；最後一年補足 75%", "AI 折舊進損益的時點（工作單第 2 節 Thesis）")
inp("da_base", "評價日在役資產年折舊（2026Q2 × 4）", "US$B/年", None, 28.416, None, "Derived", "現金流量表折舊 2026Q2 7.104 × 4（" + V22 + " capexModel.segDaRunRate）",
    "在役淨額 ≈ PP&E 321.2 − 在建 122.8＝198.4，÷ 28.4＝7.0 年 → 模型期內以直線法每年 28.4")
inp("cip0", "在建工程（尚未在役，2026-06-30）", "US$B", None, fv("bs.ppe.notInService.2026q2") / 1000 if "bs.ppe.notInService.2026q2" in F else 122.814, None, "Interested-party", "10-Q 2026Q2 附註（事實總帳 bs.ppe.notInService）", "2026H2 起轉入在役；依 AI 資本支出結構分 IT／機房")
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
inp("im", "Search 情境差額的增量 EBITDA 率", "%", 0.6, 0.7, 0.8, "Assumed", "判斷：搜尋營收增減主要只連動 TAC（約占搜尋營收兩成，Q 頁 TAC ÷ 搜尋＋聯播網）→ 增量利潤率高於分部平均", "營收減少時利潤減少得更多（營運槓桿）", judg=True)
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
inp("tp022", "v0.2.2 基準目標價（讀法 1）", "US$", None, 146.09, None, "Derived", "docs/reports/20261025_alphabet_v0.2.2.md；交接檔 v0.2.2", "讀法 2（自用 AI 價值中性）218.79；三情境皆賣出")
inp("tp022w", "v0.2.2 加權目標價（讀法 1，A 保守 175.08／B 146.09／C 積極 122.89 以 25／50／25 加權）", "US$", None, 147.54, None, "Derived", "交接檔 v0.2.2 第 1 節三情境目標價；本表以本版情境權重換算")
inp("px022", "v0.2.2 現價（2026-10-07）", "US$", None, 350.50, None, "Verified", fs("mkt.price.close.2026q2") if "mkt.price.close.2026q2" in F else "事實總帳 mkt.price.close（2026-10-07）")
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
qd("m_svc", "Google Services EBITDA 率（近四季；營業利益＋分攤折舊）", f"=(M{QR['oi_gs']}+da_ttm*sh_svc)/M{QR['gs']}", note="預測期 B 路徑的分部 EBITDA 率（自用 AI 增量成本另計）；v0.2.2 為 47.55%")
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

exec(open(os.path.join(HERE, "build_alphabet_model.py"), encoding="utf-8").read())
