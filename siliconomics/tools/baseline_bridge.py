# 三層基準協調 橋接表（步驟 3 量測；規則 docs/plan/三層基準協調規則.md v1.1：平滑法、乙前提、全市場分母）
# 方向：Wattonomics 上線 GW →（C2 ② 時差與庫存反推）實體出貨 GW →（本庫實體／成本等值比）成本等值出貨 GW，
#       與 Siliconomics 現行基準、原始顆數讀數（供給側）並排；差距拆 C2 三類；每年標 C3 狀態（門檻 50%、67%）。
# 用法：python3 siliconomics/tools/baseline_bridge.py <輸出 xlsx>   之後以 LibreOffice 重算。本工具不改任何模型。
import os, sys, datetime
import openpyxl
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter as CL

ROOT = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
sil_f = open(os.path.join(ROOT, "siliconomics/model/CURRENT")).read().strip()
wt_f = open(os.path.join(ROOT, "wattonomics/model/CURRENT")).read().strip()
S = openpyxl.load_workbook(os.path.join(ROOT, "siliconomics/model", sil_f), data_only=True)
W = openpyxl.load_workbook(os.path.join(ROOT, "wattonomics/model", wt_f), data_only=True)


def cells(wb, name):
    """具名範圍 → 值串列（單格也回串列）。"""
    sh, ref = list(wb.defined_names[name].destinations)[0]
    rng = wb[sh][ref.replace("$", "")]
    if not isinstance(rng, tuple):
        return [rng.value]
    return [c.value for row in rng for c in (row if isinstance(row, tuple) else (row,))]


def w_in(key):
    """Wattonomics In 頁 鍵 → (低, 基準)；高在下一欄。"""
    for row in W["In"].iter_rows(values_only=True):
        if row[0] == key:
            return row[7], row[8], row[9]
    raise KeyError(key)


YRS = [2026, 2027, 2028, 2029]
COL = {2025: 4, 2026: 5, 2027: 6, 2028: 7, 2029: 8}  # D..H
BOLD, BLUE = Font(bold=True), Font(color="0000FF")
HDR = PatternFill("solid", fgColor="DDEBF7")
KEY = PatternFill("solid", fgColor="FFF2CC")
wb = openpyxl.Workbook()

# ---------------- In ----------------
IN = wb.active; IN.title = "In"
IN["A1"] = f"In — 輸入（藍字）。來源：{sil_f}、{wt_f}、R4 橋接（siliconomics/tools/r4_bridge.py）。產生日 {datetime.date.today()}"
IN["A1"].font = BOLD
hdr = ["鍵", "項目", "單位", "年", "低", "基準", "高", "標記", "來源"]
for j, h in enumerate(hdr, 1):
    c = IN.cell(3, j, h); c.font = BOLD; c.fill = HDR
r = 4
N = {}


def inp(key, label, unit, yr, lo, base, hi, tag, src):
    global r
    for j, v in enumerate([key, label, unit, yr, lo, base, hi, tag, src], 1):
        c = IN.cell(r, j, v)
        if j in (5, 6, 7) and v is not None:
            c.font = BLUE
    N[key] = f"In!$F${r}"; N[key + "_lo"] = f"In!$E${r}"; N[key + "_hi"] = f"In!$G${r}"
    r += 1


lag = w_in("lag_on"); inv = w_in("inv_on")
inp("lag_on", "付款（出貨）到上線延遲", "月", None, lag[0], lag[1], lag[2], "Assumed", f"{wt_f} In lag_on（權威層＝Wattonomics，協調規則 J3）")
inp("inv_on", "未上線庫存 ÷ 年出貨", "比", None, inv[0], inv[1], inv[2], "Assumed", f"{wt_f} In inv_on（同上）")
inp("thr", "「有證據」門檻（基準 50%；高＝67% 對照）", "比", None, None, 0.5, 0.67, "Assumed", "協調規則 C3、J2；首次判定用；緩衝區 55%／45% 自第二次起適用（v1.1）")
inp("cov_done", "Wattonomics 覆蓋是否已補齊（1＝是）", "旗標", None, None, 0, None, "Derived", "v0.3 未含七家以外買方路徑（v0.3.2 W-1）；協調規則 v1.1：丙與上游有證據的乙須為 1 才改基準")
dep = cells(S, "SIF_DeployGW_Base"); coef = cells(S, "SIF_UnitsCoef_Base")[1:]
phys = cells(S, "SIF_DeployGWPhys_Base"); capon = cells(S, "SIF_CapOnlineGW_Base")
for i, y in enumerate(YRS):
    inp(f"wt_{y}", "Wattonomics 年新增上線 IT GW（主路徑，毛額）", "GW", y, None, cells(W, f"WIF_MktGW_Add_{y}")[0], None, "Derived", f"{wt_f} WIF_MktGW_Add_{y}")
    inp(f"wcov_{y}", "Wattonomics 證據比例（有站點證據的年新增 ÷ 主路徑年新增）", "比", y, None, cells(W, f"WIF_Coverage_{y}")[0], None, "Derived",
        f"{wt_f} WIF_Coverage_{y}；只計站點證據，未計七家資本支出實績（偏保守）")
    inp(f"sil_{y}", "Siliconomics 現行基準 部署 GW（成本等值、出貨年）", "GW", y, None, dep[i], None, "Derived", f"{sil_f} SIF_DeployGW_Base")
    inp(f"coef_{y}", "Siliconomics 基準顆數係數（決策 S33）", "x", y, None, coef[i], None, "Assumed", f"{sil_f} SIF_UnitsCoef_Base")
    inp(f"phys_{y}", "Siliconomics 實體部署 GW（出貨年）", "GW", y, None, phys[i], None, "Derived", f"{sil_f} SIF_DeployGWPhys_Base")
    inp(f"cap_{y}", "供給上限：可支撐新增上線 GW", "GW", y, None, capon[i], None, "Derived",
        f"{sil_f} SIF_CapOnlineGW_Base（注意：2026–29 等於本庫自身推得上線 GW，非獨立上限；R 線未解衝突第 13 項）")
inp("wt_2030", "Wattonomics 年新增上線 IT GW（主路徑，毛額）", "GW", 2030, None, cells(W, "WIF_MktGW_Add_2030")[0], None, "Derived", f"{wt_f} WIF_MktGW_Add_2030（2030 起為收斂假設 g_lt；只供平滑法 2029 用）")
inp("ship_2025", "實體出貨 GW 2025（反推的起點）", "GW", 2025, None, 16.137, None, "Derived",
    "R4 橋接 b3 2025（NVIDIA FY2026 資料中心營收 ÷ FY2027 錨點 × 世代每 GW 成本＋非 NVIDIA 約六成）；r4_bridge.py 以 v0.18／v0.3 重算")
# C2 ① 覆蓋缺口：Wattonomics 七家以外買方未入（④；W-2 實績下限讀法為主）＋ W2
for y, (lo, b, hi) in {2026: (0, 0, 0), 2027: (0.06, 3.03, 7.36), 2028: (0.06, 3.03, 7.36), 2029: (0.06, 3.03, 7.36)}.items():
    inp(f"cov4_{y}", "① Wattonomics 覆蓋缺口 ④（上線 GW）", "GW", y, lo, b, hi, "Assumed",
        "原 1.84／4.81／8.28（算力收支 11-03，全 Assumed）；W-2（11-05）(a) 七家多花由 1.84 改實績下限 0.06（0–0.92）→ 各值減 1.78（高值減 0.92）；2028–29 延續（11-05 回執）。v0.3.2 後由 W-1 實績取代")
for y, (lo, b, hi) in {2026: (0, 0, 0), 2027: (0, 0, 0), 2028: (0, 0, 0), 2029: (0, 1.12, 1.95)}.items():
    inp(f"cov2_{y}", "① W2 新雲、主權站點清單結束後年增歸零（上線 GW）", "GW", y, lo, b, hi, "Assumed", "算力收支 11-05 回執第 2 節（與 ④ 不重疊部分）")
# Siliconomics 證據比例（實績、已簽約、已確認站點支撐的比例；指引與預測不算）
for y, (lo, b, hi, src) in {
        2026: (0.44, 0.44, 0.65, "NVIDIA FY2027 資料中心 Q1＋Q2 實績 75.2＋89.0＝164.2 ÷ 錨點 375.46＝0.44（Verified ÷ 錨點）；高＝FY27Q3（11-19 公布）後約 0.65（以 Q3 指引估，未公布前不算）；非 NVIDIA 實績本庫未建，視同 NVIDIA 比例"),
        2027: (0, 0, 0.1, "無實績；CoWoS／HBM 產能是供給上限，不是需求證據；NVIDIA 在手訂單說法屬利害關係方（C5）"),
        2028: (0, 0, 0, "無實績、無已簽約訂單"),
        2029: (0, 0, 0, "同上")}.items():
    inp(f"scov_{y}", "Siliconomics 證據比例（基準值中由實績、已簽約支撐的比例）", "比", y, lo, b, hi, "Assumed", src)
for k, ref in N.items():
    wb.defined_names[k] = openpyxl.workbook.defined_name.DefinedName(k, attr_text=ref)
for j, w in enumerate([12, 52, 6, 6, 8, 9, 8, 10, 110], 1):
    IN.column_dimensions[CL(j)].width = w

# ---------------- Bridge ----------------
B = wb.create_sheet("Bridge")
B["A1"] = ("Bridge — Wattonomics 上線 GW 反推成 Siliconomics 口徑（成本等值、出貨年），與現行基準、原始顆數並排；差距拆 C2 三類。"
           "反推式（Wattonomics 自己的正向式的反函數）：上線ₜ＝(1−k)·出貨ₜ＋k·出貨ₜ₋₁，k＝lag_on÷12＋inv_on → 出貨ₜ＝(上線ₜ − k·出貨ₜ₋₁)÷(1−k)")
B["A1"].font = BOLD
for j, h in enumerate(["列", "項目", "說明", "2025"] + [str(y) for y in YRS], 1):
    c = B.cell(3, j, h); c.font = BOLD; c.fill = HDR
R = {}; br = 4
C = lambda y: CL(COL[y])


def row(key, label, note, f, bold=False, fmt="0.00", fill=None):
    global br
    B.cell(br, 1, key); B.cell(br, 2, label); B.cell(br, 3, note)
    for y in [2025] + YRS:
        v = f(y)
        if v is not None:
            c = B.cell(br, COL[y], v); c.number_format = fmt
            if fill: c.fill = fill
    if bold:
        for j in (1, 2):
            B.cell(br, j).font = BOLD
    R[key] = br; br += 1


def ship(key, label, note, online, kexpr="(lag_on/12+inv_on)", bold=False, fill=None):
    """反推實體出貨：遞迴，2025 起點 ship_2025。online(y)＝該年上線 GW 的公式文字。"""
    row(key, label, note, lambda y: "=ship_2025" if y == 2025 else
        f"=({online(y)}-{kexpr}*{CL(COL[y]-1)}{br})/(1-{kexpr})", bold, fill=fill)


row("k", "k＝lag_on÷12＋inv_on", "Wattonomics 權威（J3）", lambda y: None if y == 2025 else "=lag_on/12+inv_on", fmt="0.000")
row("w1", "Wattonomics 年新增上線 IT GW（現值）", "WIF_MktGW_Add", lambda y: None if y == 2025 else f"=wt_{y}", True)
row("w2", "① 覆蓋缺口（④＋W2；Assumed）", "七家以外買方未入；v0.3.2 後由實績取代", lambda y: None if y == 2025 else f"=cov4_{y}+cov2_{y}")
row("w3", "Wattonomics 上線（補覆蓋後）", "w1＋w2", lambda y: None if y == 2025 else f"={C(y)}{R['w1']}+{C(y)}{R['w2']}", True)
row("cap", "供給上限（上線 GW）", "SIF_CapOnlineGW_Base（非獨立，見 In）", lambda y: None if y == 2025 else f"=cap_{y}")
row("w4", "上線（補覆蓋、以供給為頂）", "MIN(w3, cap)", lambda y: None if y == 2025 else f"=MIN({C(y)}{R['w3']},{C(y)}{R['cap']})", True)
row("w4a", "上線（不補覆蓋、以供給為頂）", "MIN(w1, cap)", lambda y: None if y == 2025 else f"=MIN({C(y)}{R['w1']},{C(y)}{R['cap']})")
ship("s1", "反推實體出貨（補覆蓋）", "遞迴；2025＝ship_2025", lambda y: f"{C(y)}{R['w4']}", bold=True)
ship("s1a", "反推實體出貨（不補覆蓋）", "同上", lambda y: f"{C(y)}{R['w4a']}")
row("rt", "本庫 實體 ÷ 成本等值（出貨年）", "SIF_DeployGWPhys ÷ SIF_DeployGW；ASIC、AMD 每 GW 較便宜", lambda y: None if y == 2025 else f"=phys_{y}/sil_{y}", fmt="0.000")
row("x1", "換算值：成本等值出貨（補覆蓋）＝丙狀態基準候選", "s1 ÷ rt", lambda y: None if y == 2025 else f"={C(y)}{R['s1']}/{C(y)}{R['rt']}", True, fill=KEY)
row("x1a", "換算值（不補覆蓋）", "s1a ÷ rt", lambda y: None if y == 2025 else f"={C(y)}{R['s1a']}/{C(y)}{R['rt']}")
row("w4n", "　次年上線（補覆蓋、以供給為頂；2029 用 2030 值、覆蓋缺口沿用 2029）", "平滑法用",
    lambda y: None if y == 2025 else (f"=MIN(wt_2030+cov4_2029+cov2_2029,cap_2029)" if y == 2029 else f"={CL(COL[y]+1)}{R['w4']}"))
row("s2", "對照：平滑法實體出貨 ≈ 上線ₜ＋k ×（上線ₜ₊₁ − 上線ₜ）", "不遞迴、誤差不累積；r4_bridge 註記的「改讀設計」", lambda y: None if y == 2025 else f"={C(y)}{R['w4']}+{C(y)}{R['k']}*({C(y)}{R['w4n']}-{C(y)}{R['w4']})")
row("x2", "對照：平滑法換算值（補覆蓋）", "s2 ÷ rt；方法選擇待 Andy", lambda y: None if y == 2025 else f"={C(y)}{R['s2']}/{C(y)}{R['rt']}", True)
row("x0", "對照：上線 GW 直接換成本等值（不做時差反推）", "w4a ÷ rt", lambda y: None if y == 2025 else f"={C(y)}{R['w4a']}/{C(y)}{R['rt']}")
br += 1
row("s0", "Siliconomics 現行基準（成本等值、出貨年）", "SIF_DeployGW_Base＝原始 × S33 係數", lambda y: None if y == 2025 else f"=sil_{y}", True, fill=KEY)
row("s0r", "Siliconomics 原始顆數讀數（供給側讀數）", "現行基準 ÷ 係數", lambda y: None if y == 2025 else f"=sil_{y}/coef_{y}", True)
br += 1
B.cell(br, 1, "C2 差距拆解（現行基準 − 上線直接換算 ＝ ① ＋ ② ＋ ③）").font = BOLD; br += 1
row("g", "總差距", "s0 − x0", lambda y: None if y == 2025 else f"={C(y)}{R['s0']}-{C(y)}{R['x0']}", True)
row("g2", "② 時差與庫存（單年可、累計不可）", "x1a − x0", lambda y: None if y == 2025 else f"={C(y)}{R['x1a']}-{C(y)}{R['x0']}")
row("g2c", "　② 累計（2026 起）", "累計不得長期偏離 0", lambda y: None if y == 2025 else (f"={C(y)}{R['g2']}" if y == 2026 else f"={CL(COL[y]-1)}{br}+{C(y)}{R['g2']}"))
row("g1", "① 口徑與覆蓋（Assumed；非分歧）", "x1 − x1a", lambda y: None if y == 2025 else f"={C(y)}{R['x1']}-{C(y)}{R['x1a']}")
row("g3", "③ 殘差（分歧）", "s0 − x1", lambda y: None if y == 2025 else f"={C(y)}{R['s0']}-{C(y)}{R['x1']}", True, fill=KEY)
row("g3p", "　③ ÷ 現行基準", "容差 ±15%", lambda y: None if y == 2025 else f"={C(y)}{R['g3']}/{C(y)}{R['s0']}", fmt="0.0%")
row("g3t", "　容差內？", "", lambda y: None if y == 2025 else f'=IF(ABS({C(y)}{R["g3p"]})<=0.15,"一致","超出")', fmt="@")
row("chk", "　檢查：① ＋ ② ＋ ③ − 總差距（應為 0）", "", lambda y: None if y == 2025 else f"={C(y)}{R['g1']}+{C(y)}{R['g2']}+{C(y)}{R['g3']}-{C(y)}{R['g']}", fmt="0.000")
br += 1
B.cell(br, 1, "敏感度：換算值 x1（一次動一個輸入）").font = BOLD; br += 1
for key, lab, kexpr, covsuf in [("sl", "lag_on 低（3 月）", "(lag_on_lo/12+inv_on)", ""), ("sh", "lag_on 高（6 月）", "(lag_on_hi/12+inv_on)", ""),
                                ("il", "inv_on 低（5%）", "(lag_on/12+inv_on_lo)", ""), ("ih", "inv_on 高（20%）", "(lag_on/12+inv_on_hi)", ""),
                                ("cl", "覆蓋缺口 低", "(lag_on/12+inv_on)", "_lo"), ("ch", "覆蓋缺口 高", "(lag_on/12+inv_on)", "_hi")]:
    on = (lambda s: (lambda y: f"MIN(wt_{y}+cov4_{y}{s}+cov2_{y}{s},cap_{y})"))(covsuf)
    ship("ss_" + key, "　出貨：" + lab, "", on, kexpr)
    row("sx_" + key, "換算值：" + lab, "÷ rt", lambda y, kk=key: None if y == 2025 else f"={C(y)}{R['ss_'+kk]}/{C(y)}{R['rt']}")
B.column_dimensions["A"].width = 6; B.column_dimensions["B"].width = 46; B.column_dimensions["C"].width = 44
for y in [2025] + YRS:
    B.column_dimensions[C(y)].width = 10

# ---------------- Status ----------------
T = wb.create_sheet("Status")
T["A1"] = "Status — C3 證據狀態（規則只看狀態、不看數值高低）與依規則的基準候選；不改任何模型"
T["A1"].font = BOLD
for j, h in enumerate(["列", "項目", "說明"] + [str(y) for y in YRS], 1):
    c = T.cell(3, j, h); c.font = BOLD; c.fill = HDR
tr = 4; TR = {}
TC = {y: 4 + i for i, y in enumerate(YRS)}


def trow(key, label, note, f, fmt="0.00", fill=None):
    global tr
    T.cell(tr, 1, key); T.cell(tr, 2, label); T.cell(tr, 3, note)
    for y in YRS:
        c = T.cell(tr, TC[y], f(y)); c.number_format = fmt
        if fill: c.fill = fill
    TR[key] = tr; tr += 1


st = lambda w, s, t: f'=IF(AND({w}>={t},{s}>={t}),"甲",IF(OR({w}>={t},{s}>={t}),"乙","丙"))'
trow("ew0", "Wattonomics 證據比例（自身路徑分母；v1.0 量法）", "WIF_Coverage", lambda y: f"=wcov_{y}", "0%")
trow("ew", "Wattonomics 證據比例（全市場分母；v1.1）", "WIF_Coverage × 主路徑 ÷（主路徑＋覆蓋缺口）", lambda y: f"=wcov_{y}*wt_{y}/(wt_{y}+cov4_{y}+cov2_{y})", "0%")
trow("es", "Siliconomics 證據比例（基準）", "In scov", lambda y: f"=scov_{y}", "0%")
trow("esh", "Siliconomics 證據比例（高；NVIDIA Q3 公布後）", "", lambda y: f"=scov_{y}_hi", "0%")
EW = lambda y: f"{CL(TC[y])}{TR['ew']}"
trow("t50", "狀態（門檻 50%；首次判定）", "", lambda y: st(EW(y), f"scov_{y}", "thr"), "@", KEY)
trow("t67", "狀態（門檻 67%）", "", lambda y: st(EW(y), f"scov_{y}", "thr_hi"), "@")
trow("t50h", "狀態（門檻 50%；Siliconomics 取高值）", "", lambda y: st(EW(y), f"scov_{y}_hi", "thr"), "@")
trow("lead", "乙狀態時有證據的一層", "", lambda y: f'=IF({CL(TC[y])}{TR["t50"]}="乙",IF({EW(y)}>=thr,"Wattonomics","Siliconomics"),"—")', "@")


def up(tc, thrn):
    """v1.1：丙，或乙且有證據的是 Wattonomics → 由上游定基準。"""
    return lambda y: f'OR({CL(TC[y])}{tc}="丙",AND({CL(TC[y])}{tc}="乙",{EW(y)}>={thrn}))'


trow("b50", "依規則 v1.1 的基準（門檻 50%）", "上游定基準且覆蓋已補齊＝平滑法換算值 x2；上游定基準但覆蓋未補齊＝現行（標對帳中）；其餘＝現行",
     lambda y: f'=IF(AND({up(TR["t50"], "thr")(y)},cov_done=1),Bridge!{C(y)}{R["x2"]},sil_{y})', fill=KEY)
trow("tag", "　該年標示", "", lambda y: f'=IF({up(TR["t50"], "thr")(y)},IF(cov_done=1,"改用換算值","對帳中（待覆蓋補齊）"),IF({CL(TC[y])}{TR["t50"]}="甲",IF(ABS(Bridge!{C(y)}{R["g3p"]})<=0.15,"一致","對帳中（甲、待診斷）"),"維持"))', "@")
trow("b50s", "假設覆蓋已補齊時的基準（門檻 50%；平滑法）", "＝覆蓋補齊後預期值；v0.3.2 後以實績覆蓋重跑", lambda y: f'=IF({up(TR["t50"], "thr")(y)},Bridge!{C(y)}{R["x2"]},sil_{y})', fill=KEY)
trow("b67", "同上（門檻 67%）", "", lambda y: f'=IF({up(TR["t67"], "thr_hi")(y)},Bridge!{C(y)}{R["x2"]},sil_{y})')
trow("b50r", "對照：v1.0 遞迴反推換算（已停用）", "", lambda y: f'=IF({up(TR["t50"], "thr")(y)},Bridge!{C(y)}{R["x1"]},sil_{y})')
trow("cur", "現行基準", "", lambda y: f"=sil_{y}")
trow("raw", "原始顆數讀數（供給側；情境 A，J5）", "", lambda y: f"=sil_{y}/coef_{y}")
trow("chg", "覆蓋補齊後基準 ÷ 現行 − 1", "", lambda y: f"={CL(TC[y])}{TR['b50s']}/{CL(TC[y])}{TR['cur']}-1", "0.0%")
trow("gcur", "年增：現行基準", "斷崖檢查", lambda y: "" if y == 2026 else f"={CL(TC[y])}{TR['cur']}/{CL(TC[y]-1)}{TR['cur']}-1", "0.0%")
trow("g67", "年增：覆蓋補齊後基準（門檻 67%）", "", lambda y: "" if y == 2026 else f"={CL(TC[y])}{TR['b67']}/{CL(TC[y]-1)}{TR['b67']}-1", "0.0%")
trow("g50s", "年增：覆蓋補齊後基準（門檻 50%）", "", lambda y: "" if y == 2026 else f"={CL(TC[y])}{TR['b50s']}/{CL(TC[y]-1)}{TR['b50s']}-1", "0.0%")
trow("g50r", "年增：v1.0 遞迴（停用，對照）", "", lambda y: "" if y == 2026 else f"={CL(TC[y])}{TR['b50r']}/{CL(TC[y]-1)}{TR['b50r']}-1", "0.0%")
trow("graw", "年增：原始顆數", "", lambda y: "" if y == 2026 else f"={CL(TC[y])}{TR['raw']}/{CL(TC[y]-1)}{TR['raw']}-1", "0.0%")
T.column_dimensions["A"].width = 6; T.column_dimensions["B"].width = 44; T.column_dimensions["C"].width = 50
for y in YRS:
    T.column_dimensions[CL(TC[y])].width = 12

out = sys.argv[1]
wb.save(out)
print("wrote", out)
