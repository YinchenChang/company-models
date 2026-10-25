# 台積電模型 v0.2.1（AI 半導體 Project 步驟 4）：CoWoS 供需上限 × AI 營收由下而上 × 損益與估值
# 用法：python3 tsmc/build_tsmc.py raw.xlsx ；再以 LibreOffice headless 重算
import sys, os
import openpyxl
from openpyxl.styles import Font, PatternFill
from openpyxl.workbook.defined_name import DefinedName
ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, os.path.join(ROOT, "markets", "memory"))
import market as MEM
sys.path.insert(0, os.path.join(ROOT, "tools"))
import valuation_rule as VR

BLUE = Font(color="0000FF"); BOLD = Font(bold=True); HDR = PatternFill("solid", fgColor="DDEBF7")
wb = openpyxl.Workbook()
def hdr(ws, row, labels):
    for j, h in enumerate(labels, 1):
        c = ws.cell(row, j, h); c.font = BOLD; c.fill = HDR
YRS = (26, 27, 28, 29); SC = ("A", "B", "C")
PLAT = {"bw": {y: "GB" for y in YRS}, "rb": {y: "VR" for y in YRS}, "ru": {y: "RU" for y in YRS}, "hp": {y: "HP" for y in YRS},
        "amd": {26: "AMD8", 27: "AMDR", 28: "AMDR", 29: "AMDR"}, "tpu": {26: "TPU7", 27: "TPU8", 28: "TPU8", 29: "TPU8"},
        "trn": {y: "TRN" for y in YRS}, "oth": {y: "ASIC" for y in YRS}}
# 每片 CoWoS 晶圓可封裝的加速器數（良品），台積電內容（千美元／顆）
PP = {"GB": (14, 2.3, "Derived", "每片封裝數：36Kr 2026 NVIDIA 約 66 萬片 ÷ NVIDIA 約 7.6M 顆 ≈ 11.5（組合）；台積電內容：橋接表 v0.2（台積電內容與 v0.1 相同）（Silicon Analysts、Epoch：邏輯＋CoWoS 約 $1.9K，B300 封裝溢價）"),
      "VR": (8, 2.8, "Derived", "Morgan Stanley 2027 NVIDIA 122.2 萬片 ÷ 約 9.7M 顆 ≈ 7.9（組合）；內容：橋接表 v0.2（台積電內容與 v0.1 相同）"),
      "RU": (4, 6.0, "Assumed", "4 顆光罩尺寸晶粒＋16 疊 HBM4E"),
      "HP": (29, 1.1, "Analogy", "CoWoS-S 單晶粒；Silicon Analysts H100 邏輯 $300＋封裝 $750"),
      "AMD8": (8.5, 2.0, "Derived／Analogy", "36Kr 2026 AMD 約 9 萬片 ÷ 約 0.75M 顆"),
      "AMDR": (6, 3.0, "Assumed", "MI455X 12 疊 HBM4"),
      "TPU7": (28, 1.5, "Derived／Assumed", "36Kr：Broadcom 約 20 萬片、TPU 60–65% → 約 12.5 萬片 ÷ 約 3.85M 顆"),
      "TPU8": (24, 1.8, "Assumed", "TPU 8 HBM 288GB"),
      "TRN": (30, 1.2, "Derived／Assumed", "36Kr：世芯約 6 萬片（Trainium3）"),
      "ASIC": (20, 1.2, "Assumed", "")}

IN = wb.active; IN.title = "In"
IN["A1"] = "In — 輸入（藍字；基準欄 E）。金額：十億美元；台積電內容：千美元／顆；CoWoS：千片"
hdr(IN, 3, ["鍵", "項目", "單位", "低", "基準", "高", "標記", "來源／說明"])
rows = []; sec = lambda t: rows.append(("sec", t)); inp = lambda *a: rows.append(a)
sec("A. 加速器顆數（markets/memory，與 HBM、網通同一份）")
for k, nm, u, gb, tag, src in MEM.UNITS:
    for i, y in enumerate((26, 27, 28)):
        inp(f"u_{k}_{y}", f"{nm} 顆數 20{y}", "百萬顆", None, u[i], None, tag, ("markets/memory：" + src) if i == 0 else "同上")
    inp(f"g29_{k}", f"{nm} 2029 顆數成長", "%", None, {"bw": -0.5, "hp": 0.0}.get(k, 0.15), None, "Assumed", "")
for s, (v28, v29) in zip(SC, [MEM.SCEN_UNITS[s_] for s_ in SC]):   # 單一來源：markets/memory/market.py（HBM 需求位元用同一組）
    inp(f"um28_{s}", f"情境 {s}：2028 顆數 ÷ 基準", "x", None, v28, None, "Assumed", "A＝基準；B＝AI 資本支出消化；C＝明顯下修" if s == "A" else "")
    inp(f"um29_{s}", f"情境 {s}：2029 顆數 ÷ 基準", "x", None, v29, None, "Assumed", "")
sec("B. 每片 CoWoS 晶圓封裝數、每顆台積電內容")
for p, (pp, ct, tag, src) in PP.items():
    inp(f"pp_{p}", f"{p}：每片 CoWoS 封裝數", "顆／片", round(pp * 0.8, 1), pp, round(pp * 1.2, 1), tag, src)
    inp(f"ct_{p}", f"{p}：每顆台積電內容（邏輯晶圓＋CoWoS）", "$K／顆", round(ct * 0.8, 2), ct, round(ct * 1.2, 2), tag, "同上")
inp("cow_oth26", "其他 CoWoS 用途 2026（網通 ASIC、FPGA、中國 AI 晶片等）", "千片", None, 150, None, "Assumed", "MS 全球需求 139.4 萬片 − 本表加速器")
inp("cow_oth_g", "其他 CoWoS 用途年增", "%", None, 0.25, None, "Assumed", "")
sec("C. CoWoS 產能（台積電，月產能年底值，千片）")
for y, v, src in ((25, 77.5, "TrendForce（工商時報）2025-12：7.5–8 萬"), (26, 130, "TrendForce 12–14 萬；Morgan Stanley 約 13 萬"), (27, 200, "Morgan Stanley 20 萬；KGI 17 萬"), (28, 260, "Morgan Stanley 26 萬"), (29, 320, "Assumed")):
    inp(f"cap_end{y}", f"CoWoS 月產能 20{y} 年底", "千片／月", None, v, None, "Interested-party" if y < 29 else "Assumed", src)
inp("osat", "委外封測（日月光、艾克爾）補充產能 ÷ 台積電", "%", 0.05, 0.10, 0.20, "Assumed", "日月光 CoWoP 等")
inp("ms_dem26", "Morgan Stanley 全球 CoWoS 需求 2026", "千片", None, 1394, None, "Interested-party", "Morgan Stanley 2026-07")
inp("ms_dem27", "Morgan Stanley 全球 CoWoS 需求 2027", "千片", None, 2509, None, "Interested-party", "Morgan Stanley 2026-09（2.694M 為較早版本）")
inp("hbm_gap27", "HBM 2027 需求缺口 ÷ 供給（markets/memory）", "%", None, 0.21, None, "Derived", "markets/memory Mkt：需求 7.10 EB vs 供給 5.86 EB")
sec("D. 台積電營收")
inp("fx", "美元兌新台幣（2027 起）", "NT$/US$", 30, 32, 33, "Assumed", "v0.1.1 標 Verified（當時 Q3 指引匯率）；4Q26 指引匯率為 31.8，2027 起未獲公司指引，數值不變")
inp("rev25", "營收 2025", "$B", None, 122.43, None, "Verified", "四季加總")
PR3 = "公司：3Q26 新聞稿 pr.tsmc.com/english/news/3346（6-K a3q26e）｜媒體：見 Sources 頁"
inp("rev26_1q", "營收 1Q26", "$B", None, 35.90, None, "Verified", "公司：1Q26 新聞稿；NT$1,134.10B")
inp("rev26_2q", "營收 2Q26", "$B", None, 40.20, None, "Verified", "公司：2Q26 新聞稿；NT$1,270.38B")
inp("rev26_3q", "營收 3Q26", "$B", None, 46.72, None, "Verified", PR3 + "｜NT$1,494.24B；多數媒體 46.7（個別媒體 46.13、46.87 與公司數不符，不採）")
inp("rev26_4q", "營收 4Q26（指引低／中／高）", "$B", 52.3, 53.3, 54.3, "Interested-party", "公司 4Q26 指引 US$52.3–54.3B（匯率 31.8）｜媒體：BigGo、民視財經；中點為本模型取值，公司歷來常高於指引上緣")
inp("fx26_4q", "4Q26 指引匯率", "NT$/US$", None, 31.8, None, "Interested-party", "公司 4Q26 指引假設")
inp("ntd26_1q", "營收 1Q26（新台幣）", "NT$B", None, 1134.10, None, "Verified", "公司：1Q26 新聞稿")
inp("ntd26_2q", "營收 2Q26（新台幣）", "NT$B", None, 1270.38, None, "Verified", "公司：2Q26 新聞稿")
inp("ntd26_3q", "營收 3Q26（新台幣）", "NT$B", None, 1494.24, None, "Verified", PR3)
inp("ntd26_4q", "營收 4Q26（新台幣；指引美元 × 指引匯率）", "NT$B", "=rev26_4q_lo*fx26_4q", "=rev26_4q*fx26_4q", "=rev26_4q_hi*fx26_4q", "Derived", "")
inp("gm26_1q", "毛利率 1Q26", "%", None, 0.662, None, "Verified", "公司：1Q26 新聞稿")
inp("gm26_2q", "毛利率 2Q26", "%", None, 0.677, None, "Verified", "公司：2Q26 新聞稿")
inp("gm26_3q", "毛利率 3Q26", "%", None, 0.664, None, "Verified", PR3 + "｜2Q26 的 67.7% 下滑 1.3 個百分點，公司歸因 2 奈米初期爬坡稀釋")
inp("gm26_4q", "毛利率 4Q26（指引低／中／高）", "%", 0.65, 0.66, 0.67, "Interested-party", "公司 4Q26 指引 65–67%｜媒體：民視財經（中值 66%）")
inp("opm26_1q", "營業利益率 1Q26", "%", None, 0.581, None, "Verified", "公司：1Q26 新聞稿")
inp("opm26_2q", "營業利益率 2Q26", "%", None, 0.603, None, "Verified", "公司：2Q26 新聞稿")
inp("opm26_3q", "營業利益率 3Q26", "%", None, 0.587, None, "Verified", PR3)
inp("opm26_4q", "營業利益率 4Q26（指引低／中／高）", "%", 0.565, 0.575, 0.585, "Interested-party", "公司 4Q26 指引 56.5–58.5%｜媒體：BigGo")
inp("ni26_1q", "淨利 1Q26", "NT$B", None, 572.48, None, "Verified", "公司：1Q26 新聞稿（EPS NT$22.08）")
inp("ni26_2q", "淨利 2Q26", "NT$B", None, 706.56, None, "Verified", "公司：2Q26 新聞稿（EPS NT$27.25）")
inp("ni26_3q", "淨利 3Q26", "NT$B", None, 763.06, None, "Verified", PR3 + "（EPS NT$29.43）")
inp("eps26_1to3q", "EPS 1–3Q26 實績合計", "NT$", None, "=22.08+27.25+29.43", None, "Verified", "公司：各季新聞稿 22.08、27.25、29.43")
inp("hpc_3q", "HPC 占營收 3Q26（備忘，未進入公式）", "%", None, 0.68, None, "Verified", "公司：3Q26 簡報／逐字稿，2Q26 為 66%｜媒體：vestedfinance、Quest Review")
inp("g26_guid", "公司 2026 全年美元營收成長指引", "%", None, 0.44, None, "Interested-party", "公司：由「略高於 40%」上修至約 44%｜媒體：StockFeel、民視財經；僅作 Recon 對照，不驅動營收")
inp("rev26", "營收 2026（1–3Q 實績＋4Q 指引中點；低／高＝4Q 指引下／上緣）", "$B", "=rev26_1q+rev26_2q+rev26_3q+rev26_4q_lo", "=rev26_1q+rev26_2q+rev26_3q+rev26_4q", "=rev26_1q+rev26_2q+rev26_3q+rev26_4q_hi", "Derived", "v0.1.1 為 172（全年指引「略高於 40%」推算）")
inp("ntd26", "營收 2026（新台幣）", "NT$B", None, "=ntd26_1q+ntd26_2q+ntd26_3q+ntd26_4q", None, "Derived", "")
inp("fx_26", "2026 平均匯率（新台幣營收 ÷ 美元營收）", "NT$/US$", None, "=ntd26/rev26", None, "Derived", "僅用於 2026 EPS；2027 起用 fx")
inp("ai_oth", "其他 AI 相關 ÷ 加速器內容（網通 ASIC、HBM 基底晶粒、CPU）", "%", 0.0, 0.2, 0.4, "Assumed", "使 2026 AI 營收約等於台積電口徑推算（2025 十幾%後段 × CAGR 約 57%）")
for y, v in ((27, 0.12), (28, 0.08), (29, 0.05)):
    inp(f"px{y}", f"AI 內容價格調升 20{y}", "%", None, v, None, "Interested-party／Derived", "2027 先進製程 +5–10%（Nikkei、MS）、CoWoS 混合單價 +23%（JPMorgan）" if y == 27 else "Assumed")
for s, vals in zip(SC, ((0.12, 0.10, 0.08), (0.12, 0.05, 0.03), (0.12, 0.0, -0.03))):
    for y, v in zip((27, 28, 29), vals):
        inp(f"na{y}_{s}", f"非 AI 營收成長 20{y}（{s}）", "%", None, v, None, "Assumed", "公司 2024–29 全公司 CAGR 約 25%、AI 約 55–59% → 非 AI 約 10–15%" if (y == 27 and s == "A") else "")
sec("E. 利潤率、資本支出、股本、估值")
inp("gm26", "毛利率 2026（四季以新台幣營收加權）", "%", None, "=(ntd26_1q*gm26_1q+ntd26_2q*gm26_2q+ntd26_3q*gm26_3q+ntd26_4q*gm26_4q)/ntd26", None, "Derived", "v0.1.1 為 66%")
inp("opm26", "營業利益率 2026（四季以新台幣營收加權）", "%", None, "=(ntd26_1q*opm26_1q+ntd26_2q*opm26_2q+ntd26_3q*opm26_3q+ntd26_4q*opm26_4q)/ntd26", None, "Derived", "")
inp("opex_26", "營業費用 ÷ 營收 2026", "%", None, "=gm26-opm26", None, "Derived", "僅用於 2026；2027 起用 opex")
inp("ni_op_26", "淨利 ÷ 營業利益 2026（1–3Q 實績比，4Q 沿用）", "x", None, "=(ni26_1q+ni26_2q+ni26_3q)/(ntd26_1q*opm26_1q+ntd26_2q*opm26_2q+ntd26_3q*opm26_3q)", None, "Derived", "1Q 0.869、2Q 0.922、3Q 0.870；4Q 公司未指引淨利率，Assumed 沿用 1–3Q 比")
for s, vals in zip(SC, ((0.64, 0.63, 0.62), (0.64, 0.58, 0.56), (0.64, 0.54, 0.52))):
    for y, v in zip((27, 28, 29), vals):
        inp(f"gm{y}_{s}", f"毛利率 20{y}（{s}）", "%", None, v, None, "Derived／Assumed", "N2 稀釋 3–4 個百分點、海外廠 2–3 個百分點（公司）對上漲價；長期目標 56% 以上" if (y == 27 and s == "A") else "")
inp("opex", "營業費用 ÷ 營收", "%", None, 0.075, None, "Derived", "2Q26 毛利率 67.7% − 營業利益率 60.3%")
inp("ni_op", "淨利 ÷ 營業利益（2027 起；低／高＝區間）", "x", "=ni_op_underlying", "=(1+nonop_hist)*(1-etr_new)", "=ni_op_2023", "Derived", "基準＝(1＋歷史營業外比率)×(1－新稅率)，推導見 NI_OI 頁；低 0.864＝2026 年扣除一次性利益後的內含水準；高 0.909＝2023 年年度值（稅率上升前、2022–2025 年度最高）。v0.1.1 為 0.92（含一次性利益的 2Q26 單季值，不再當上限）")
inp("shares", "股數", "十億股", None, 25.93, None, "Derived", "市值 ÷ 股價；3Q26 淨利 NT$763.06B ÷ EPS 29.43 ＝ 25.93，一致")
inp("px", "股價（2026-10-07 收盤）", "NT$", None, 2585, None, "Interested-party", "Wantgoo")
for y, v, tag, src in ((26, 62, "Interested-party", "公司指引 US$60–64B 中點（7 月上修；10 月維持不變）｜媒體：民視財經、vestedfinance"), (27, 77, "Interested-party", "Citi；券商 73–81"), (28, 90, "Interested-party", "JPMorgan"), (29, 95, "Assumed", "")):
    inp(f"capex{y}", f"資本支出 20{y}", "$B", None, v, None, tag, src)
for i, (v, src) in enumerate(((15, "近谷底（2022 年約 11.9x）"), (20.5, "ADR 10 年中位數（GuruFocus）"), (32.7, "10 年最高（GuruFocus）")), 1):
    inp(f"pe_{i}", f"P/E {v}x", "x", None, v, None, "Interested-party", src)
inp("cons_rev27", "共識營收 2027", "$B", None, 229.0, None, "Interested-party", "MarketScreener NT$7,329B ÷ 32")
inp("cons_eps26", "共識 EPS 2026", "NT$", None, 107.6, None, "Interested-party", "MarketScreener")
inp("cons_eps27", "共識 EPS 2027", "NT$", None, 142.0, None, "Interested-party", "MarketScreener")
inp("ai25_share", "AI 加速器 ÷ 營收 2025", "%", None, 0.18, None, "Verified", "公司：十幾%後段（high-teens）")
inp("ai_cagr", "AI 加速器營收 CAGR（2024–29）", "%", None, 0.57, None, "Verified", "公司：50% 中後段，往高 50% 靠攏")
for s, w in (("A", 0.3), ("B", 0.5), ("C", 0.2)):
    inp(f"tp_w{s}", f"目標價情境權重 {s}（v0.2.2 新增；Val 頁第 0 節用）", "%", None, w, None, "Assumed", "Andy 可改（工作單 2026-10-25 預設 A 0.3／B 0.5／C 0.2）；三者合計應為 100%（Val 頁第 0 節列示）")
r = 4
for srow in rows:
    if srow[0] == "sec":
        IN.cell(r, 1, srow[1]).font = BOLD; r += 1; continue
    key, item, unit, lo, base, hi, tag, src = srow
    IN.cell(r, 1, key); IN.cell(r, 2, item); IN.cell(r, 3, unit)
    for j, v in ((4, lo), (5, base), (6, hi)):
        if v is not None:
            IN.cell(r, j, v)
            if not (isinstance(v, str) and v.startswith("=")): IN.cell(r, j).font = BLUE
    IN.cell(r, 7, tag); IN.cell(r, 8, src)
    wb.defined_names[key] = DefinedName(key, attr_text=f"In!$E${r}")
    if lo is not None: wb.defined_names[key + "_lo"] = DefinedName(key + "_lo", attr_text=f"In!$D${r}")
    if hi is not None: wb.defined_names[key + "_hi"] = DefinedName(key + "_hi", attr_text=f"In!$F${r}")
    r += 1
for col, w in zip("ABCDEFGH", [12, 50, 10, 8, 10, 8, 20, 110]): IN.column_dimensions[col].width = w
IN.freeze_panes = "D4"

def u(k, y, s="A"):
    base = f"u_{k}_{y}" if y < 29 else f"u_{k}_28*(1+g29_{k})"
    if y == 28: return f"{base}*um28_{s}"
    if y == 29: return f"{base}*um29_{s}"
    return base

# ---------------- CoWoS：供需 ----------------
CW = wb.create_sheet("CoWoS")
CW["A1"] = "CoWoS — 需求（加速器顆數 ÷ 每片封裝數＋其他用途）vs 產能（台積電月產能年均 × 12 × (1＋委外)）；千片。2028–29 用情境 A 顆數"
hdr(CW, 3, ["項目", "2026", "2027", "2028", "2029", "說明"])
YC = {26: "B", 27: "C", 28: "D", 29: "E"}; CR = {}; cr = [4]
def cput(label, fs, key=None, fmt="#,##0", note="", bold=False):
    rr = cr[0]; CW.cell(rr, 1, label).font = Font(bold=bold)
    for j, f in enumerate(fs):
        if f in (None, ""): continue
        CW.cell(rr, 2 + j, f).number_format = fmt
    CW.cell(rr, 6, note)
    if key: CR[key] = rr
    cr[0] += 1
for k, nm, *_ in MEM.UNITS:
    cput(f"  {nm}", [f"={u(k, y)}/pp_{PLAT[k][y]}*1000" for y in YRS], key="w_" + k)
cput("  其他 CoWoS 用途", ["=cow_oth26"] + [f"={YC[y-1]}{{r}}*(1+cow_oth_g)" for y in (27, 28, 29)], key="w_o")
for y in (27, 28, 29): CW.cell(CR["w_o"], 2 + YRS.index(y), f"={YC[y-1]}{CR['w_o']}*(1+cow_oth_g)").number_format = "#,##0"
cput("需求合計", [f"=SUM({YC[y]}4:{YC[y]}{CR['w_o']})" for y in YRS], key="dem", bold=True)
cput("  其中 NVIDIA", [f"={YC[y]}{CR['w_bw']}+{YC[y]}{CR['w_rb']}+{YC[y]}{CR['w_ru']}+{YC[y]}{CR['w_hp']}" for y in YRS], key="nv")
cput("台積電產能（月產能年均 × 12）", [f"=(cap_end{y-1}+cap_end{y})/2*12" for y in YRS], key="cap")
cput("可用產能（含委外）", [f"={YC[y]}{CR['cap']}*(1+osat)" for y in YRS], key="capt")
cput("需求 ÷ 可用產能", [f"={YC[y]}{CR['dem']}/{YC[y]}{CR['capt']}" for y in YRS], key="util", fmt="0%", note=">100%＝CoWoS 為瓶頸")
cput("對照：HBM 需求缺口 ÷ 供給（2027）", ["", "=hbm_gap27"], fmt="0%", note="markets/memory；與上一列比，數字大者為較緊的瓶頸")
cput("對照：Morgan Stanley 全球需求", ["=ms_dem26", "=ms_dem27"], key="ms")
cput("  本模型 ÷ MS", [f"=B{CR['dem']}/B{CR['ms']}-1", f"=C{CR['dem']}/C{CR['ms']}-1"], fmt="0%")
cput("  NVIDIA：本模型 vs 36Kr 66 萬片（2026）、MS 122.2 萬片（2027）", [f"=B{CR['nv']}/660-1", f"=C{CR['nv']}/1222-1"], fmt="0%")
CW.column_dimensions["A"].width = 52; CW.column_dimensions["F"].width = 40

# ---------------- P&L：三情境 ----------------
PL = wb.create_sheet("PL")
PL["A1"] = "PL — 台積電損益（十億美元；EPS 新台幣）。AI 營收＝Σ 顆數 × 每顆內容 × 價格調升 × (1＋其他 AI)；非 AI＝前一年 × (1＋成長)"
cols = [("B", "2026", None, 26)]
L = iter("CDEFGHIJK")
for s in SC:
    for y in (27, 28, 29):
        cols.append((next(L), f"20{y} {s}", s, y))
hdr(PL, 3, ["項目"] + [c[1] for c in cols])
PR = {k: 4 + i for i, k in enumerate(["ai_acc", "ai", "nonai", "rev", "aish", "gm", "gp", "op", "ni", "eps", "pe_px", "capex", "fcf"])}
LAB = {"ai_acc": "AI 加速器內容（Σ 顆數 × 每顆內容 × 價格）", "ai": "AI 相關營收（含其他 AI）", "nonai": "非 AI 營收", "rev": "營收",
       "aish": "AI ÷ 營收", "gm": "毛利率", "gp": "毛利", "op": "營業利益", "ni": "淨利", "eps": "EPS（新台幣）", "pe_px": "目前股價 ÷ EPS",
       "capex": "資本支出", "fcf": "淨利 − 資本支出（不含折舊，粗估）"}
for k, rr in PR.items(): PL.cell(rr, 1, LAB[k])
def pxf(y):
    return {26: "1", 27: "(1+px27)", 28: "(1+px27)*(1+px28)", 29: "(1+px27)*(1+px28)*(1+px29)"}[y]
for col, lab, s, y in cols:
    s_ = s or "A"
    acc = "(" + "+".join(f"{u(k, y, s_)}*ct_{PLAT[k][y]}" for k, *_ in MEM.UNITS) + f")*{pxf(y)}"
    PL[f"{col}{PR['ai_acc']}"] = "=" + acc
    PL[f"{col}{PR['ai']}"] = f"={col}{PR['ai_acc']}*(1+ai_oth)"
    if y == 26:
        PL[f"{col}{PR['rev']}"] = "=rev26"
        PL[f"{col}{PR['nonai']}"] = f"={col}{PR['rev']}-{col}{PR['ai']}"
        PL[f"{col}{PR['gm']}"] = "=gm26"
    else:
        prev = "B" if y == 27 else [c[0] for c in cols if c[2] == s and c[3] == y - 1][0]
        PL[f"{col}{PR['nonai']}"] = f"={prev}{PR['nonai']}*(1+na{y}_{s})"
        PL[f"{col}{PR['rev']}"] = f"={col}{PR['ai']}+{col}{PR['nonai']}"
        PL[f"{col}{PR['gm']}"] = f"=gm{y}_{s}"
    PL[f"{col}{PR['aish']}"] = f"={col}{PR['ai']}/{col}{PR['rev']}"
    PL[f"{col}{PR['gp']}"] = f"={col}{PR['rev']}*{col}{PR['gm']}"
    sx = "_26" if y == 26 else ""
    PL[f"{col}{PR['op']}"] = f"={col}{PR['rev']}*({col}{PR['gm']}-opex{sx})"
    PL[f"{col}{PR['ni']}"] = f"={col}{PR['op']}*ni_op{sx}"
    PL[f"{col}{PR['eps']}"] = f"={col}{PR['ni']}*fx{sx}/shares"
    PL[f"{col}{PR['pe_px']}"] = f"=px/{col}{PR['eps']}"
    PL[f"{col}{PR['capex']}"] = f"=capex{y}"
    PL[f"{col}{PR['fcf']}"] = f"={col}{PR['ni']}-{col}{PR['capex']}"
    for k, rr in PR.items():
        PL[f"{col}{rr}"].number_format = "0.0%" if k in ("aish", "gm") else ("0.0" if k != "pe_px" else "0.0\"x\"")
PL.column_dimensions["A"].width = 44
for c in "BCDEFGHIJK": PL.column_dimensions[c].width = 10
PL.cell(18, 1, "註：B 欄 2026 ＝1–3Q 實績＋4Q 指引中點（匯率、營業費用、淨利比用 2026 專屬參數）；非 AI 營收＝營收 − AI（由下而上）。2027 起兩者分開推。")

# ---------------- Val ----------------
V = wb.create_sheet("Val")
V["A1"] = "Val — P/E 估值（新台幣）。隱含股價 ÷ 目前股價 − 1"
hdr(V, 3, ["方法", "倍數", "EPS 年度", "情境", "EPS", "隱含股價", "相對目前"])
vr = 4
for col, lab, s, y in cols:
    if y not in (27, 28): continue
    for i in (1, 2, 3):
        V.cell(vr, 1, f"P/E"); V.cell(vr, 2, f"=pe_{i}"); V.cell(vr, 3, f"20{y}"); V.cell(vr, 4, s)
        V.cell(vr, 5, f"=PL!{col}{PR['eps']}").number_format = "0.0"
        V.cell(vr, 6, f"=B{vr}*E{vr}").number_format = "#,##0"
        V.cell(vr, 7, f"=F{vr}/px-1").number_format = "0%"
        vr += 1
V.cell(vr + 1, 1, "讀法：目前股價 ÷ 2027 EPS 即市場給的遠期 P/E；與 10 年中位 20.5x 比，看市場是否已提前反映 2027–28 的成長與情境 B／C 風險")
for c, w in zip("ABCDEFG", [8, 8, 10, 8, 10, 12, 10]): V.column_dimensions[c].width = w

# 估值規則區塊：十年 P/E 歷史與規則值（2027＝中段取中位數、區間 25–75 分位；另列 2028 情境 C × 75 分位）
# 歷史值原放在 markets/overview/extract.py（2026-10-23 移入）
VR.add_rule_block(V, vr + 4, [{"id": "P/E", "name": "P/E", "sources": [
    {"label": "FullRatio", "vals": [14.5, 17.71, 16.13, 25.62, 31.11, 29.18, 11.94, 19.46, 28.72, 29.03], "tag": "Interested-party"},
    {"label": "companiesmarketcap", "vals": [11.6, 14.8, 14.2, 23.5, 28.3, 27.2, 11.2, 19.3, 28.8, 28.3], "tag": "Verified"}],
    "cases": [
        {"case": "中段", "pos": "中段 → 中位數", "q": 50, "basis_label": "2027 EPS（A）", "basis": f"=PL!C{PR['eps']}"},
        {"case": "低", "pos": "25 分位（區間下緣）", "q": 25, "basis_label": "2027 EPS（A）", "basis": f"=PL!C{PR['eps']}"},
        {"case": "高", "pos": "75 分位（區間上緣）", "q": 75, "basis_label": "2027 EPS（A）", "basis": f"=PL!C{PR['eps']}"},
        {"case": "2028 C", "pos": "參考：2028 情境 C × 75 分位", "q": 75, "basis_label": "2028 EPS（C）", "basis": f"=PL!J{PR['eps']}"}]}],
    "0.0", "規則（Andy 2026-10-13）：合理倍數＝2016–2025 年底 P/E 依循環位置取分位；台積電 2027 為中段，取中位數，區間為 25 與 75 分位。兩個來源各自取分位再平均、四捨五入到兩位。倍數歷史為輸入（藍字）。")
V.column_dimensions["A"].width = 36

# ---------------- Recon ----------------
RC = wb.create_sheet("Recon")
RC["A1"] = "Recon — 本模型 vs 公司口徑、共識（只作事後對照）"
hdr(RC, 3, ["項目", "本模型", "對照值", "差距", "對照來源"])
rec = [("2026 AI 相關營收 vs 公司口徑推算", f"=PL!B{PR['ai']}", "=rev25*ai25_share*(1+ai_cagr)", "2025 AI 加速器十幾%後段 × CAGR 57%（公司）"),
       ("2026 AI 加速器內容（不含其他 AI）vs 同上", f"=PL!B{PR['ai_acc']}", "=rev25*ai25_share*(1+ai_cagr)", "其他 AI 相關（ai_oth）為 Assumed"),
       ("2026 EPS vs 共識", f"=PL!B{PR['eps']}", "=cons_eps26", "MarketScreener"),
       ("2026 營收 vs 公司全年成長指引（2025 × (1＋44%)）", f"=PL!B{PR['rev']}", "=rev25*(1+g26_guid)", "公司 3Q26 法說：全年美元成長約 44%"),
       ("2026 EPS vs 1–3Q 實績＋4Q 指引推算", f"=PL!B{PR['eps']}", "=eps26_1to3q+ntd26_4q*opm26_4q*ni_op_26/shares", "公司各季 EPS 22.08／27.25／29.43＋4Q 指引中點"),
       ("2027 營收（A）vs 共識", f"=PL!C{PR['rev']}", "=cons_rev27", "MarketScreener"),
       ("2027 EPS（A）vs 共識", f"=PL!C{PR['eps']}", "=cons_eps27", "MarketScreener"),
       ("2026 CoWoS 需求 vs Morgan Stanley", f"=CoWoS!B{CR['dem']}", "=ms_dem26", "千片"),
       ("2027 CoWoS 需求 vs Morgan Stanley", f"=CoWoS!C{CR['dem']}", "=ms_dem27", "千片")]
for i, (nm, a, b, src) in enumerate(rec, 4):
    RC.cell(i, 1, nm); RC.cell(i, 2, a).number_format = "#,##0.0"; RC.cell(i, 3, b).number_format = "#,##0.0"
    RC.cell(i, 4, f"=B{i}/C{i}-1").number_format = "0%"; RC.cell(i, 5, src)
RC.column_dimensions["A"].width = 44; RC.column_dimensions["E"].width = 50

# ---------------- Sens ----------------
SE = wb.create_sheet("Sens")
SE["A1"] = "Sens — 2027 情境 A：一次動一個輸入"
hdr(SE, 3, ["輸入", "2027 EPS", "相對基準", "2027 CoWoS 需求 ÷ 產能"])
base_eps = f"PL!C{PR['eps']}"; base_ai = f"PL!C{PR['ai']}"; base_rev = f"PL!C{PR['rev']}"
def eps_from(rev, gm="gm27_A"):
    return f"=({rev})*({gm}-opex)*ni_op*fx/shares"
items = [("基準", eps_from(base_rev), f"=CoWoS!C{CR['util']}"),
         ("Rubin 顆數 −30%", eps_from(f"{base_rev}-0.3*u_rb_27*ct_VR*(1+px27)*(1+ai_oth)"), f"=(CoWoS!C{CR['dem']}-0.3*CoWoS!C{CR['w_rb']})/CoWoS!C{CR['capt']}"),
         ("AI 內容價格 2027 +0%（不漲）", eps_from(f"{base_rev}-{base_ai}*px27/(1+px27)"), f"=CoWoS!C{CR['util']}"),
         ("毛利率 2027 60%", eps_from(base_rev, "0.60"), f"=CoWoS!C{CR['util']}"),
         ("毛利率 2027 67%", eps_from(base_rev, "0.67"), f"=CoWoS!C{CR['util']}"),
         ("非 AI 營收 2027 +0%", eps_from(f"{base_rev}-PL!B{PR['nonai']}*na27_A"), f"=CoWoS!C{CR['util']}"),
         ("CoWoS 2027 年底產能 17 萬片（KGI）", eps_from(base_rev), f"=CoWoS!C{CR['dem']}/((cap_end26+170)/2*12*(1+osat))")]
for i, (nm, e, ut) in enumerate(items, 4):
    SE.cell(i, 1, nm); SE.cell(i, 2, e).number_format = "0.0"; SE.cell(i, 3, f"=B{i}/B$4-1").number_format = "0%"; SE.cell(i, 4, ut).number_format = "0%"
SE.column_dimensions["A"].width = 36

SE.cell(13, 1, "淨利 ÷ 營業利益 對 EPS 的影響（情境 A）").font = BOLD
hdr(SE, 14, ["設定", "比值", "2027 EPS", "相對基準", "2028 EPS", "相對基準"])
for k, (nm, rr_) in enumerate((("基準（推導值）", "ni_op"), ("淨利 ÷ 營業利益 0.864（低）", "ni_op_lo"), ("淨利 ÷ 營業利益 0.909（高）", "ni_op_hi")), 15):
    SE.cell(k, 1, nm); SE.cell(k, 2, f"={rr_}").number_format = "0.000"
    SE.cell(k, 3, f"=PL!C{PR['op']}*B{k}*fx/shares").number_format = "0.0"
    SE.cell(k, 4, f"=C{k}/C$15-1").number_format = "0.0%"
    SE.cell(k, 5, f"=PL!D{PR['op']}*B{k}*fx/shares").number_format = "0.0"
    SE.cell(k, 6, f"=E{k}/E$15-1").number_format = "0.0%"
for c, w in zip("BCDEF", [10, 10, 10, 10, 10]): SE.column_dimensions[c].width = w



# ---------------- NI_OI：淨利 ÷ 營業利益拆解 ----------------
NO = wb.create_sheet("NI_OI")
NO["A1"] = "NI_OI — 淨利 ÷ 營業利益拆解（新台幣十億）。淨利÷營業利益 ＝ (1＋營業外÷營業利益) × (1－有效稅率) × 少數股權等調整。藍字＝輸入"
hdr(NO, 3, ["期間", "營業利益", "營業外收支", "稅前淨利", "所得稅", "淨利", "利息收入", "匯兌損益", "一次性項目", "有效稅率", "營業外÷營業利益", "少數股權等調整", "淨利÷營業利益", "獨立來源數", "標記", "來源／說明"])
AR22 = "https://investor.tsmc.com/static/annualReports/2022/english/index.html"
PER = [  # 期間, OI, nonop, tax, NI, interest, fx, oneoff, 來源數, 標記, 說明
 ("2022", 1121.28, 22.91, None, 1016.90, None, None, 0, 1, "Verified（單一來源）", "公司年報（經搜尋摘錄）；稅額＝稅前淨利－淨利（Derived）；媒體僅能核對淨利（歸屬母公司 1,016.53，股東會資料）；利息、匯兌明細未找到。待複核：另一搜尋結果的稅前淨利 1,140.7 疑為還原表欄位誤讀"),
 ("2023", 921.47, 57.71, 141.40, 837.77, 60.29, None, 0, 2, "Verified", "公司：2023 年報（6-K 2023-02）｜媒體：MoneyDJ 113 年損益表的前期欄；利息 60.29 僅 2025 20-F 一個來源；匯兌未找到"),
 ("2024", 1322.05, 83.79, 233.41, 1172.43, 87.21, 10.00, 0, 2, "Verified", "公司：2024 年報、2025 20-F｜媒體：MoneyDJ、中央社；匯兌 10.00 僅 20-F／財報一個來源"),
 ("2025", 1936.09, 105.56, 326.25, 1715.40, 105.74, 13.83, 0, 2, "Verified", "公司：2025 年報、2025 20-F｜媒體：工商時報（營業外約 1,056 億、稅前 2 兆 417 億、稅 3,263 億）；匯兌 13.83 僅 20-F／財報一個來源；其他利益及損失未拆"),
 ("1Q26", "=ntd26_1q*opm26_1q", 28.83, None, 572.48, None, None, 0, 1, "Derived", "營業利益＝新台幣營收×營益率（Derived）；營業外 28.83 取自 2Q26 簡報比較欄，欄位標籤未能直接確認；稅額＝稅前淨利－淨利（Derived）"),
 ("2Q26", "=ntd26_2q*opm26_2q", 95.83, 155.65, 706.56, None, None, 63.20, 2, "Verified／Derived", "公司：2Q26 簡報、3Q26 Management Report（稅 155.65、有效稅率 18.1%）；SEC 3Q26 6-K（含 2Q26 比較欄）核對營業利益、稅前淨利、淨利相符｜媒體：鉅亨（營業外 958.27 億）；營業利益 NT$ 欄為 Derived。一次性項目 63.20 的出處見下方「一次性利益出處」"),
 ("3Q26", 876.45, 34.59, 147.70, 763.06, None, None, 0, 2, "Verified", "公司：3Q26 Management Report（營業外 34.59、稅 147.70、有效稅率 16.2%；較 2Q26 降，主因無 VIS 處分與評價利益）｜SEC 3Q26 6-K a3q26e（營業利益、稅前淨利、淨利核對相符；使用者提供的核對）"),
]
r0 = 4
for i, (per, oi, no, tx, ni, it, fxg, one, nsrc, tag, note) in enumerate(PER):
    r = r0 + i
    NO.cell(r, 1, per)
    for j, v in ((2, oi), (3, no), (5, tx), (6, ni), (7, it), (8, fxg), (9, one)):
        if v is None: continue
        c = NO.cell(r, j, v)
        if not (isinstance(v, str) and v.startswith("=")): c.font = BLUE
    NO.cell(r, 4, f"=B{r}+C{r}")
    if tx is None: NO.cell(r, 5, f"=D{r}-F{r}")
    NO.cell(r, 14, nsrc); NO.cell(r, 15, tag); NO.cell(r, 16, note)
rT = r0 + len(PER)          # 1–3Q26 合計列
NO.cell(rT, 1, "1–3Q26 合計")
for j in "BCDEFI": NO[f"{j}{rT}"] = f"=SUM({j}{rT-3}:{j}{rT-1})"
NO.cell(rT, 15, "Derived"); NO.cell(rT, 16, "三季加總；一次性項目＝2Q26 VIS 處分利益（含股權評價利益未另拆）")
for r in range(r0, rT + 1):
    NO[f"J{r}"] = f"=E{r}/D{r}"; NO[f"K{r}"] = f"=C{r}/B{r}"; NO[f"L{r}"] = f"=F{r}/(D{r}-E{r})"; NO[f"M{r}"] = f"=F{r}/B{r}"
    for c in "BCDEFGHI": NO[f"{c}{r}"].number_format = "#,##0.0"
    for c in "JK": NO[f"{c}{r}"].number_format = "0.0%"
    NO[f"L{r}"].number_format = "0.000"; NO[f"M{r}"].number_format = "0.000"
p0 = rT + 2
NO.cell(p0, 1, "基準值推導（In!ni_op）").font = BOLD
rows_ = [
 ("新稅率＝2024 起各期所得稅合計 ÷ 稅前淨利合計（2024、2025、1–3Q26）", f"=(E{r0+2}+E{r0+3}+E{rT})/(D{r0+2}+D{r0+3}+D{rT})", "0.0%", "etr_new",
  "取 2024 起：2024 年起全球最低稅負生效，有效稅率由 2022 年約 11% 升到 14–18%（2024 16.6%、2025 16.0%、2Q26 18.1%、3Q26 16.2%）；視為長期性變化"),
 ("歷史營業外比率＝2022–2025 加 1–3Q26（扣除一次性項目）的營業外合計 ÷ 營業利益合計", f"=(SUM(C{r0}:C{r0+3})+C{rT}-I{rT})/(SUM(B{r0}:B{r0+3})+B{rT})", "0.0%", "nonop_hist",
  "營業外收支主體是利息收入（2024 87.2、2025 105.7，占營業外 100% 以上）；匯兌 2024 +10.0、2025 +13.8，量級小且為收益；2Q26 VIS 處分利益 63.2 為一次性，已扣除"),
 ("基準＝(1＋歷史營業外比率) × (1－新稅率)", f"=(1+B{p0+2})*(1-B{p0+1})", "0.000", "ni_op_calc", "少數股權等調整各期約 1.000，略去"),
 ("對照：2024、2025 實績比（前版低值 0.887）", f"=(F{r0+2}+F{r0+3})/(B{r0+2}+B{r0+3})", "0.000", None, ""),
 ("對照：1–3Q26 實績比", f"=M{rT}", "0.000", None, "含 2Q26 一次性利益"),
 ("區間低值＝1–3Q26 扣除一次性利益（按新稅率稅後）", f"=(F{rT}-I{rT}*(1-B{p0+1}))/B{rT}", "0.000", "ni_op_underlying", "2026 年的內含水準；1Q26、3Q26 單季分別 0.869、0.871"),
 ("區間高值＝2023 年年度值（稅率上升前、2022–2025 年度最高）", f"=M{r0+1}", "0.000", "ni_op_2023", "2022 0.907、2023 0.909、2024 0.887、2025 0.886"),
 ("敏感：63.2 若非一次性（不扣除）→ 歷史營業外比率", f"=(SUM(C{r0}:C{r0+3})+C{rT})/(SUM(B{r0}:B{r0+3})+B{rT})", "0.0%", "nonop_hist_alt", "分子不扣 63.2"),
 ("敏感：63.2 若非一次性 → 基準", f"=(1+B{p0+8})*(1-B{p0+1})", "0.000", "ni_op_alt", "相對基準 0.875 的變化見 README"),
]
for k, (lab, f, fmt, nm, note) in enumerate(rows_, 1):
    rr = p0 + k
    NO.cell(rr, 1, lab); NO.cell(rr, 2, f).number_format = fmt; NO.cell(rr, 16, note)
    if nm: wb.defined_names[nm] = DefinedName(nm, attr_text=f"NI_OI!$B${rr}")
NO.cell(p0 + 14, 1, "一次性利益出處（2Q26 VIS 處分與評價利益 63.20）：公司 2Q26 合併財務報告（SEC 6-K）載明處分利益 NT$63,202,285 千元，其中含對剩餘持股改按公允價值衡量的非現金重新衡量利益 56,398,783 千元（喪失重大影響力）；公司 2026-05-15 公告表示近期無進一步出售計畫；3Q26 Management Report 說明營業外由 95.83 降至 34.59 主因無該項處分與評價利益；媒體：中央社、東森財經、鉅亨。頁碼：搜尋摘錄未含頁碼，未取得（文件與段落見 Sources 頁）。SEC 3Q26 6-K 本身未標註一次性項目。")
NO.cell(p0 + 12, 1, "判定：差距兩個來源都有——(a) 稅率由 2022 年 11% 升至 2024 年起約 16.6%，屬長期；(b) 營業外比率 2022 年 2.0% → 2023–25 年 5.5–6.3%（利息收入墊高），2Q26 另有一次性處分利益使 0.92 偏高。依規則取「新稅率 × 歷史平均營業外比率」。")
NO.column_dimensions["A"].width = 60; NO.column_dimensions["P"].width = 120; NO.column_dimensions["O"].width = 22
for c in "BCDEFGHIJKLMN": NO.column_dimensions[c].width = 11

# ---------------- Sources、README ----------------
SO = wb.create_sheet("Sources"); SO["A1"] = "來源"; SO["A1"].font = BOLD
srcs = [("台積電 2Q26 新聞稿", "https://pr.tsmc.com/english/news/3326"),
 ("台積電 2Q26 法說逐字稿", "https://investor.tsmc.com/english/encrypt/files/encrypt_file/reports/2026-08/3e494f0c14dd0890f897aa044415e21d93486cc4/TSMC%202Q26%20Transcript.pdf"),
 ("台積電 4Q25 法說逐字稿", "https://investor.tsmc.com/japanese/encrypt/files/encrypt_file/reports/2026-01/51d09df96cd89ac19d65af39032b038dc2896a24/TSMC%204Q25%20Transcript.pdf"),
 ("台積電 2026-09 營收（6-K）", "https://www.sec.gov/Archives/edgar/data/0001046179/000104617926000680/tsm-revenue20261008.htm"),
 ("TrendForce CoWoS 2025-12-08", "https://www.trendforce.com/news/2025/12/08/news-tsmcs-cowos-l-s-reportedly-fully-booked-osat-partners-step-up-with-ases-cowop-in-focus"),
 ("TrendForce CoWoS 2026-06-15", "https://www.trendforce.com/news/2026/06/15/news-tsmc-cowos-supply-demand-gap-reportedly-seen-narrowing-from-20-to-10-by-end-2026-as-capacity-expands/"),
 ("經濟日報／Morgan Stanley 2026-09-13", "https://money.udn.com/money/story/5607/9752236"),
 ("36Kr CoWoS 分配 2026-01-09", "https://eu.36kr.com/en/p/3631956080985095"),
 ("經濟日報／JPMorgan 2026-09-29", "https://money.udn.com/money/amp/story/5607/9783752"),
 ("Tom's Hardware／Nikkei 2027 漲價", "https://www.tomshardware.com/tech-industry/semiconductors/tsmc-eyes-price-hikes-of-up-to-25-percent-on-chip-production-services-in-2027-report-claims-plans-to-raise-baseline-prices-by-5-percent-to-10-percent-on-advanced-nodes"),
 ("MarketScreener 共識", "https://in.marketscreener.com/quote/stock/TSMC-TAIWAN-SEMICONDUCTOR-6492349/valuation/"),
 ("GuruFocus P/E", "https://www.gurufocus.com/term/pettm/TSM"),
 ("【v0.2】台積電 3Q26 新聞稿（公司）", "https://pr.tsmc.com/english/news/3346"),
 ("【v0.2】台積電 3Q26 EarningsRelease（公司）", "https://investor.tsmc.com/english/encrypt/files/encrypt_file/reports/2026-10/526b65af7305df2dde6b9ece6812b1fdf1e49f9b/3Q26%20EarningsRelease.pdf"),
 ("【v0.2】台積電 3Q26 6-K（含指引）", "https://www.sec.gov/Archives/edgar/data/0001046179/000104617926000726/a3q26e_withguidancexfinal.htm"),
 ("【v0.2】台積電 3Q26 逐字稿（公司）", "https://investor.tsmc.com/english/encrypt/files/encrypt_file/reports/2026-10/eba4fac72873bf546b669bd2b5d973df5c93215d/TSMC%203Q26%20Transcript.pdf"),
 ("【v0.2】台積電 1Q26 新聞稿", "https://pr.tsmc.com/english/news/3297"),
 ("【v0.2】台積電 2Q26 6-K（含指引）", "https://www.sec.gov/Archives/edgar/data/0001046179/000104617926000451/a2q26e_withguidancexfinal.htm"),
 ("【v0.2】媒體：StockFeel 3Q26 法說重點", "https://www.stockfeel.com.tw/%E5%8F%B0%E7%A9%8D%E9%9B%BB-2330-2026q3-%E6%B3%95%E8%AA%AA%E6%9C%83/"),
 ("【v0.2】媒體：民視財經 3Q26 法說", "https://finance.ftvnews.com.tw/news/detail/2026A15W0572"),
 ("【v0.2】媒體：Futurum 3Q26", "https://futurumgroup.com/insights/tsmc-q3-fy-2026-earnings-beat-and-raise-on-ai-and-hpc-demand/"),
 ("【v0.2】媒體：BigGo Finance 3Q26", "https://finance.biggo.com/news/US_TSM_2026-10-15"),
 ("【v0.2】媒體：vestedfinance 3Q26", "https://vestedfinance.com/blog/us-stocks/tsmc-q3-earnings-2026/"),
 ("【v0.2】媒體：TrendForce 券商解讀 3Q26", "https://www.trendforce.com/news/2026/10/16/news-after-tsmcs-3q26-earnings-key-brokerage-takeaways-as-2028-capex-forecasts-reach-us105b/"),
 ("【v0.2】NI_OI：台積電 2024 年報", "https://investor.tsmc.com/static/annualReports/2024/english/index.html"),
 ("【v0.2】NI_OI：台積電 2025 年報", "https://investor.tsmc.com/static/annualReports/2025/english/index.html"),
 ("【v0.2】NI_OI：台積電 2025 20-F（利息、匯兌）", "https://investor.tsmc.com/sites/ir/sec-filings/2025_20F%20Report.pdf"),
 ("【v0.2】NI_OI：台積電 3Q26 Management Report", "https://investor.tsmc.com/schinese/encrypt/files/encrypt_file/qr/phase4_reports/2026-10/f53a64fd40fc0b453f94e077e5aba8cad506f27b/3Q26%20ManagementReport.pdf"),
 ("【v0.2】NI_OI：MoneyDJ 113 年（2024）損益表", "https://www.moneydj.com/kmdj/news/newsviewer.aspx?a=0c9c7a58-321f-4b5c-83c9-7f4837872969"),
 ("【v0.2】NI_OI：工商時報 2025 年報", "https://www.ctee.com.tw/news/20260416702066-430502"),
 ("【v0.2】NI_OI：台積電 2Q26 合併財務報告（SEC 6-K；VIS 處分利益 63,202,285 千元）", "https://www.sec.gov/Archives/edgar/data/0001046179/000104617926000541/a2026q2consolidatedreport-.htm"),
 ("【v0.2】NI_OI：台積電出售 VIS 股權公告（SEC 6-K，2026-05）", "https://www.sec.gov/Archives/edgar/data/0001046179/000104617926000280/tsmctosell81ofvanguardinte.htm"),
 ("【v0.2】NI_OI：台積電 3Q26 6-K（核對 2Q26、3Q26 營業利益、稅前淨利、淨利）", "https://www.sec.gov/Archives/edgar/data/0001046179/000104617926000726/a3q26e_withguidancexfinal.htm"),
 ("【v0.2】NI_OI：中央社（處分 VIS 1.52 億股）", "https://www.cna.com.tw/news/afe/202605180297.aspx"),
 ("【v0.2】NI_OI：東森財經（估獲利 632.02 億）", "https://fnc.ebc.net.tw/fncnews/stock/210975"),
 ("【v0.2】NI_OI：鏡報（2Q26 處分 VIS 利益）", "https://www.mirrordaily.news/story/79601"),
 ("【v0.2】NI_OI：鉅亨 2Q26 財報分析", "https://hao.cnyes.com/post/259466"),
 ("【v0.2】查證方式說明", "公司原文頁面在本次建置環境無法直接開啟（網路限制），3Q26 數字取自搜尋引擎對公司新聞稿的摘錄，並與至少兩則獨立媒體交叉；公司逐字稿原文待人工複核"),
 ("markets/chip_bridge v0.2、markets/memory（本 repo）", "")]
for i, (a, b) in enumerate(srcs, 2): SO.cell(i, 1, a); SO.cell(i, 2, b)
SO.column_dimensions["A"].width = 44; SO.column_dimensions["B"].width = 130
RD = wb.create_sheet("README", 0)
lines = [("台積電模型 v0.2.2（AI 半導體 Project 步驟 4）　2026-10-25　2026 年改為 1–3Q 實績＋4Q 指引中點；2027 起假設未變", True),
 ("模型命題：台積電在 AI 算力中的意義是供給上限而非分到的金額比例——AI 加速器只占其營收約兩成；2027 年加速器出貨受 HBM 而非 CoWoS 限制，台積電的獲利成長由「AI 顆數 × 每顆內容 × 漲價」與「非 AI 八成的穩定成長」共同決定。", False), ("", False),
 ("驅動因子 → 推導量", True),
 ("CoWoS：加速器顆數（markets/memory）÷ 每片封裝數 → CoWoS 需求；月產能年均 × 12 × (1＋委外) → 可用產能；需求 ÷ 產能 → 是否為瓶頸；與 HBM 缺口比較", False),
 ("PL：顆數 × 每顆台積電內容（橋接表）× 價格調升 × (1＋其他 AI) → AI 營收；非 AI 前一年 × (1＋成長)；× 毛利率 − 營業費用 → 營業利益 → 淨利 → EPS", False),
 ("情境：A 基準；B 2028 顆數 −20%、毛利率降至 58%；C 2028 顆數 −40%、毛利率 54%（AI 資本支出明顯下修）", False),
 ("Val：EPS × P/E（15x、10 年中位 20.5x、10 年最高 32.7x）", False), ("Recon、Sens：見各頁", False), ("", False),
 ("標記：Verified、Interested-party、Analogy、Assumed、Derived。SemiAnalysis 未使用。", False),
 ("v0.2.1 變更（2026-10-23）：Val 頁新增「十年倍數與規則值」區塊（十年 P/E 歷史、分位、規則值與隱含股價；原放在 markets/overview/extract.py），總覽與 HTML 展示層改讀此處；損益、EPS 與原有儲存格不變。", False),
 ("v0.2 變更（2026-10-22，3Q26 財報後）：2026 營收 172→176.1 $B（1–3Q 實績＋4Q 指引中點 53.3）；毛利率 66.0%→66.5%（四季加權）；EPS 114.2→112.1（新舊對照見 tsmc/README.md）。新增 2026 專屬參數 opex_26、ni_op_26、fx_26（來自實績）；2027 起淨利÷營業利益由 0.92 改為推導值（NI_OI 頁；區間 0.864–0.909）；毛利率、價格調升、非 AI 成長、CoWoS 產能全部未改（公司未給新的明確指引）。", False),
 ("公司 3Q26 說法未取得：AI 加速器營收占比、CoWoS 產能的新數字（搜尋未找到逐字稿原文），故 CoWoS 頁不動。", False),
 ("已知限制：(1) 台積電不揭露 CoWoS 片數與先進封裝營收，產能為券商估計；(2) 先進製程（N3／N2）晶圓產能未找到，未建晶圓瓶頸；(3) 其他 AI 相關營收比例為 Assumed；(4) 淨利 ÷ 營業利益 2027 起為單一推導值（見 NI_OI），2022 年與 1Q26 的明細為單一來源。", False)]
lines.append(("v0.2.2（2026-10-25）：Val 頁新增「0. 目標價與上檔空間」（放在 Val 既有內容之後，不位移既有儲存格）與 In 頁目標價情境權重 tp_wA／tp_wB／tp_wC（0.3／0.5／0.2，Assumed）。各情境隱含價＝十年 P/E 中位數（規則值：2027 中段）× 各情境 2027 EPS；其餘儲存格不變。", False))
for i, (t, b) in enumerate(lines, 1): RD.cell(i, 1, t).font = Font(bold=b)
RD.column_dimensions["A"].width = 170

# ---------------- Val 第 0 節：目標價與上檔空間（v0.2.2 新增；Andy 2026-10-25 要求，列標籤依工作單逐字） ----------------
# 為不位移既有儲存格（與舊版 dist 逐格相同），本節放在 Val 既有內容之後；HTML 估值章第一個區塊即為本節
def add_target_price(V, imp, note, fmt, method, px_note):
    t0 = V.max_row + 2
    V.cell(t0, 1, "0. 目標價與上檔空間").font = BOLD
    hdr(V, t0 + 1, ["項目", "數值", "", "", "說明"])
    V.cell(t0 + 2, 1, "目標價（12 個月，情境加權）"); V[f"B{t0 + 2}"] = f"=tp_wA*B{t0 + 4}+tp_wB*B{t0 + 5}+tp_wC*B{t0 + 6}"; V[f"B{t0 + 2}"].number_format = fmt
    V[f"E{t0 + 2}"] = "權重 × 各情境隱含價"
    V.cell(t0 + 3, 1, "上檔空間（目標價 ÷ 現價 − 1）"); V[f"B{t0 + 3}"] = f"=B{t0 + 2}/px-1"; V[f"B{t0 + 3}"].number_format = "0.0%"; V[f"E{t0 + 3}"] = px_note
    for i, (s, lab) in enumerate((("A", "情境 A 隱含價（樂觀）"), ("B", "情境 B 隱含價（基準）"), ("C", "情境 C 隱含價（下檔）"))):
        rr = t0 + 4 + i
        V.cell(rr, 1, lab); V[f"B{rr}"] = imp[s]; V[f"B{rr}"].number_format = fmt; V[f"E{rr}"] = note[s]
    rr = t0 + 7
    V.cell(rr, 1, "情境權重 A／B／C"); V[f"B{rr}"] = "=tp_wA"; V[f"C{rr}"] = "=tp_wB"; V[f"D{rr}"] = "=tp_wC"
    for c in "BCD": V[f"{c}{rr}"].number_format = "0%"
    V[f"E{rr}"] = f'="合計 "&TEXT(B{rr}+C{rr}+D{rr},"0%")&"；In 頁 tp_wA／tp_wB／tp_wC（Assumed，Andy 可改）"'
    V.cell(t0 + 8, 1, "目標價方法"); V[f"B{t0 + 8}"] = method
    return t0
MID = [rw for rw in range(1, V.max_row + 1) if V.cell(rw, 1).value == "規則值：中段（P/E）"]
assert len(MID) == 1, MID
EPS27 = {"A": "C", "B": "F", "C": "I"}   # PL 頁 2027 欄（表頭：2027 A／B／C）
assert all(PL[f"{c}3"].value == f"2027 {s}" for s, c in EPS27.items())
TP0 = add_target_price(V, {s: f"=$C${MID[0]}*PL!{EPS27[s]}{PR['eps']}" for s in SC},
    {s: f"情境 {s} 2027 EPS（PL 頁）× 十年 P/E 中位數（第 6 節規則值：中段）" for s in SC}, "#,##0",
    "P/E：各情境 2027 EPS × 十年（2016–2025 年底）P/E 中位數（FullRatio 與 companiesmarketcap 兩來源平均；規則值：2027 為中段，取中位數，第 6 節）。"
    "2027 三情境 EPS 相同（情境差從 2028 起），故三情境隱含價相同；十年倍數兼含 2022 年低谷（11–12x）與 2020、2024–2025 年（28–31x），中位數無明顯單向偏誤；"
    "隱含價約為 2027 年底的價格、未折現，當作 12 個月目標價時偏高；現價為 2026-10-07 收盤（本版未更新）。",
    "現價＝In 頁 px（2026-10-07 收盤）")
wb.save(sys.argv[1]); print("saved")
