# 台積電模型 v0.1（AI 半導體 Project 步驟 4）：CoWoS 供需上限 × AI 營收由下而上 × 損益與估值
# 用法：python3 tsmc/build_tsmc.py raw.xlsx ；再以 LibreOffice headless 重算
import sys, os
import openpyxl
from openpyxl.styles import Font, PatternFill
from openpyxl.workbook.defined_name import DefinedName
ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, os.path.join(ROOT, "markets", "memory"))
import market as MEM

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
for s, (v28, v29) in zip(SC, ((1.0, 1.0), (0.8, 0.85), (0.6, 0.6))):
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
inp("fx", "美元兌新台幣", "NT$/US$", 30, 32, 33, "Verified", "公司 Q3 指引匯率假設")
inp("rev25", "營收 2025", "$B", None, 122.43, None, "Verified", "四季加總")
inp("rev26", "營收 2026", "$B", 168, 172, 176, "Verified／Derived", "Q1 35.9、Q2 40.2、Q3 約 46.7（7–9 月營收 NT$1,494.2B ÷ 32）；全年指引「略高於 40%」")
inp("ai_oth", "其他 AI 相關 ÷ 加速器內容（網通 ASIC、HBM 基底晶粒、CPU）", "%", 0.0, 0.2, 0.4, "Assumed", "使 2026 AI 營收約等於台積電口徑推算（2025 十幾%後段 × CAGR 約 57%）")
for y, v in ((27, 0.12), (28, 0.08), (29, 0.05)):
    inp(f"px{y}", f"AI 內容價格調升 20{y}", "%", None, v, None, "Interested-party／Derived", "2027 先進製程 +5–10%（Nikkei、MS）、CoWoS 混合單價 +23%（JPMorgan）" if y == 27 else "Assumed")
for s, vals in zip(SC, ((0.12, 0.10, 0.08), (0.12, 0.05, 0.03), (0.12, 0.0, -0.03))):
    for y, v in zip((27, 28, 29), vals):
        inp(f"na{y}_{s}", f"非 AI 營收成長 20{y}（{s}）", "%", None, v, None, "Assumed", "公司 2024–29 全公司 CAGR 約 25%、AI 約 55–59% → 非 AI 約 10–15%" if (y == 27 and s == "A") else "")
sec("E. 利潤率、資本支出、股本、估值")
inp("gm26", "毛利率 2026", "%", None, 0.66, None, "Verified／Derived", "1Q 66.2%、2Q 67.7%、Q3 指引 65–67%")
for s, vals in zip(SC, ((0.64, 0.63, 0.62), (0.64, 0.58, 0.56), (0.64, 0.54, 0.52))):
    for y, v in zip((27, 28, 29), vals):
        inp(f"gm{y}_{s}", f"毛利率 20{y}（{s}）", "%", None, v, None, "Derived／Assumed", "N2 稀釋 3–4 個百分點、海外廠 2–3 個百分點（公司）對上漲價；長期目標 56% 以上" if (y == 27 and s == "A") else "")
inp("opex", "營業費用 ÷ 營收", "%", None, 0.075, None, "Derived", "2Q26 毛利率 67.7% − 營業利益率 60.3%")
inp("ni_op", "淨利 ÷ 營業利益", "x", None, 0.92, None, "Derived", "2Q26 淨利率 55.6% ÷ 營業利益率 60.3%")
inp("shares", "股數", "十億股", None, 25.93, None, "Derived", "市值 ÷ 股價")
inp("px", "股價（2026-10-07 收盤）", "NT$", None, 2585, None, "Interested-party", "Wantgoo")
for y, v, tag, src in ((26, 62, "Verified", "指引 US$60–64B（7 月上修）"), (27, 77, "Interested-party", "Citi；券商 73–81"), (28, 90, "Interested-party", "JPMorgan"), (29, 95, "Assumed", "")):
    inp(f"capex{y}", f"資本支出 20{y}", "$B", None, v, None, tag, src)
for i, (v, src) in enumerate(((15, "近谷底（2022 年約 11.9x）"), (20.5, "ADR 10 年中位數（GuruFocus）"), (32.7, "10 年最高（GuruFocus）")), 1):
    inp(f"pe_{i}", f"P/E {v}x", "x", None, v, None, "Interested-party", src)
inp("cons_rev27", "共識營收 2027", "$B", None, 229.0, None, "Interested-party", "MarketScreener NT$7,329B ÷ 32")
inp("cons_eps26", "共識 EPS 2026", "NT$", None, 107.6, None, "Interested-party", "MarketScreener")
inp("cons_eps27", "共識 EPS 2027", "NT$", None, 142.0, None, "Interested-party", "MarketScreener")
inp("ai25_share", "AI 加速器 ÷ 營收 2025", "%", None, 0.18, None, "Verified", "公司：十幾%後段（high-teens）")
inp("ai_cagr", "AI 加速器營收 CAGR（2024–29）", "%", None, 0.57, None, "Verified", "公司：50% 中後段，往高 50% 靠攏")
r = 4
for srow in rows:
    if srow[0] == "sec":
        IN.cell(r, 1, srow[1]).font = BOLD; r += 1; continue
    key, item, unit, lo, base, hi, tag, src = srow
    IN.cell(r, 1, key); IN.cell(r, 2, item); IN.cell(r, 3, unit)
    for j, v in ((4, lo), (5, base), (6, hi)):
        if v is not None: IN.cell(r, j, v).font = BLUE
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
    PL[f"{col}{PR['op']}"] = f"={col}{PR['rev']}*({col}{PR['gm']}-opex)"
    PL[f"{col}{PR['ni']}"] = f"={col}{PR['op']}*ni_op"
    PL[f"{col}{PR['eps']}"] = f"={col}{PR['ni']}*fx/shares"
    PL[f"{col}{PR['pe_px']}"] = f"=px/{col}{PR['eps']}"
    PL[f"{col}{PR['capex']}"] = f"=capex{y}"
    PL[f"{col}{PR['fcf']}"] = f"={col}{PR['ni']}-{col}{PR['capex']}"
    for k, rr in PR.items():
        PL[f"{col}{rr}"].number_format = "0.0%" if k in ("aish", "gm") else ("0.0" if k != "pe_px" else "0.0\"x\"")
PL.column_dimensions["A"].width = 44
for c in "BCDEFGHIJK": PL.column_dimensions[c].width = 10
PL.cell(18, 1, "註：B 欄 2026 營收取公司指引推算；非 AI 營收＝營收 − AI（由下而上）。2027 起兩者分開推。")

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

# ---------------- Recon ----------------
RC = wb.create_sheet("Recon")
RC["A1"] = "Recon — 本模型 vs 公司口徑、共識（只作事後對照）"
hdr(RC, 3, ["項目", "本模型", "對照值", "差距", "對照來源"])
rec = [("2026 AI 相關營收 vs 公司口徑推算", f"=PL!B{PR['ai']}", "=rev25*ai25_share*(1+ai_cagr)", "2025 AI 加速器十幾%後段 × CAGR 57%（公司）"),
       ("2026 AI 加速器內容（不含其他 AI）vs 同上", f"=PL!B{PR['ai_acc']}", "=rev25*ai25_share*(1+ai_cagr)", "其他 AI 相關（ai_oth）為 Assumed"),
       ("2026 EPS vs 共識", f"=PL!B{PR['eps']}", "=cons_eps26", "MarketScreener"),
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
 ("markets/chip_bridge v0.2、markets/memory（本 repo）", "")]
for i, (a, b) in enumerate(srcs, 2): SO.cell(i, 1, a); SO.cell(i, 2, b)
SO.column_dimensions["A"].width = 44; SO.column_dimensions["B"].width = 130
RD = wb.create_sheet("README", 0)
lines = [("台積電模型 v0.1.1（AI 半導體 Project 步驟 4）　2026-10-22（Tokenomics v19.0 同步：數字與 v0.1 相同）", True),
 ("模型命題：台積電在 AI 算力中的意義是供給上限而非分到的金額比例——AI 加速器只占其營收約兩成；2027 年加速器出貨受 HBM 而非 CoWoS 限制，台積電的獲利成長由「AI 顆數 × 每顆內容 × 漲價」與「非 AI 八成的穩定成長」共同決定。", False), ("", False),
 ("驅動因子 → 推導量", True),
 ("CoWoS：加速器顆數（markets/memory）÷ 每片封裝數 → CoWoS 需求；月產能年均 × 12 × (1＋委外) → 可用產能；需求 ÷ 產能 → 是否為瓶頸；與 HBM 缺口比較", False),
 ("PL：顆數 × 每顆台積電內容（橋接表）× 價格調升 × (1＋其他 AI) → AI 營收；非 AI 前一年 × (1＋成長)；× 毛利率 − 營業費用 → 營業利益 → 淨利 → EPS", False),
 ("情境：A 基準；B 2028 顆數 −20%、毛利率降至 58%；C 2028 顆數 −40%、毛利率 54%（AI 資本支出明顯下修）", False),
 ("Val：EPS × P/E（15x、10 年中位 20.5x、10 年最高 32.7x）", False), ("Recon、Sens：見各頁", False), ("", False),
 ("標記：Verified、Interested-party、Analogy、Assumed、Derived。SemiAnalysis 未使用。", False),
 ("已知限制：(1) 台積電不揭露 CoWoS 片數與先進封裝營收，產能為券商估計；(2) 先進製程（N3／N2）晶圓產能未找到，未建晶圓瓶頸；(3) 其他 AI 相關營收比例為 Assumed；(4) 淨利 ÷ 營業利益固定 0.92。", False)]
for i, (t, b) in enumerate(lines, 1): RD.cell(i, 1, t).font = Font(bold=b)
RD.column_dimensions["A"].width = 170
wb.save(sys.argv[1]); print("saved")
