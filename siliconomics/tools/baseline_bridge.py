# 三層基準協調 橋接表（步驟 3 量測；規則 docs/plan/三層基準協調規則.md v1.1：平滑法、乙前提、全市場分母）
# 方向：Wattonomics 上線 GW →（C2 ② 時差與庫存反推）實體出貨 GW →（本庫實體／成本等值比）成本等值出貨 GW，
#       與 Siliconomics 現行基準、原始顆數讀數（供給側）並排；差距拆 C2 三類；每年標 C3 狀態（門檻 50%、67%）。
# 用法：python3 siliconomics/tools/baseline_bridge.py <輸出 xlsx>   之後以 LibreOffice 重算。本工具不改任何模型。
# 2026-11-18 r2 收尾：p1–p4 改標「上線年 GW 交叉檢核」（查核確認口徑）；新增 q1–q4 付款年讀數；s0 標籤改「Siliconomics 基準（SIF_DeployGW_Base）」；不改既有數值。
# 2026-11-18（三層對帳第二波 W2-02／W2-03／W2-04A；不改任何既有數值）：Bridge 頁在 x0 與 s0 之間新增「付款年 GW 交叉檢核」p1–p4（只檢核，不被任何既有列引用）；
#   新增 Stock 頁（出貨到上線存量帳，A／B／C 三節，stk_2025 低／基準／高三組）與 Residual 頁（殘差帳骨架，歸因項讀 builder/seed/residual_items_v0.1.json）。
import os, sys, json, datetime
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
IN["A1"] = (f"In — 輸入（藍字）。來源：{sil_f}、{wt_f}、R4 橋接（siliconomics/tools/r4_bridge.py）。產生日 {datetime.date.today()}。"
            f"注意：本檔讀的 Siliconomics 是執行當時 model/CURRENT 指向的版本（{sil_f}，通常是「前一版」）；本檔換算值 x2 再被下一版的 In 頁 crd_x2 讀入，"
            f"所以來源版本落後一版是設計，不是標錯（三層對帳工作單 A8）")
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
inp("cov_done", "Wattonomics 覆蓋是否已補齊（1＝是）", "旗標", None, None, 1 if "WIF_Evidence_2026" in W.defined_names else 0, None, "Derived",
    "v0.3.2 起（W-1 七家以外買方入主路徑）＝1；v0.3 以前＝0。協調規則 v1.1：丙與上游有證據的乙須為 1 才改基準")
dep = cells(S, "SIF_DeployGW_Base"); coef = cells(S, "SIF_UnitsCoef_Base")[1:]
phys = cells(S, "SIF_DeployGWPhys_Base"); capon = cells(S, "SIF_CapOnlineGW_Base")
for i, y in enumerate(YRS):
    inp(f"wt_{y}", "Wattonomics 年新增上線 IT GW（主路徑，毛額）", "GW", y, None, cells(W, f"WIF_MktGW_Add_{y}")[0], None, "Derived", f"{wt_f} WIF_MktGW_Add_{y}")
    if "WIF_Evidence_2026" in W.defined_names:
        _e = cells(W, f"WIF_Evidence_{y}")[0]
        _lo = cells(W, f"WIF_Evidence_Lo_{y}")[0] if f"WIF_Evidence_Lo_{y}" in W.defined_names else _e   # v0.3.6 起有區間（規則 v1.3）
        _hi = cells(W, f"WIF_Evidence_Hi_{y}")[0] if f"WIF_Evidence_Hi_{y}" in W.defined_names else _e
        inp(f"wcov_{y}", "Wattonomics 證據比例（WIF_Evidence：全市場分母，含七家實績與站點；v0.3.6 起低／高＝WIF_Evidence_Lo_／Hi_，之前只有點值）", "比", y, _lo, _e, _hi, "Derived",
            f"{wt_f} WIF_Evidence_{y}、WIF_Evidence_Lo_{y}、WIF_Evidence_Hi_{y}（v0.3.2 起；分母已含七家以外買方，故 Status 頁 ew 不再另除覆蓋缺口）")
    else:
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
# v0.3.2 起（W-1：七家以外買方入主路徑、X6 新雲與主權 Assumed 年增入基準；有 WIF_Evidence 名稱）覆蓋已補齊 → 缺口＝0、cov_done＝1
W32 = "WIF_Evidence_2026" in W.defined_names
GAP4 = {2026: (0, 0, 0), 2027: (0.06, 3.03, 7.36), 2028: (0.06, 3.03, 7.36), 2029: (0.06, 3.03, 7.36)}
GAP2 = {2026: (0, 0, 0), 2027: (0, 0, 0), 2028: (0, 0, 0), 2029: (0, 1.12, 1.95)}
if W32:
    GAP4 = {y: (0, 0, 0) for y in YRS}; GAP2 = {y: (0, 0, 0) for y in YRS}
for y, (lo, b, hi) in GAP4.items():
    inp(f"cov4_{y}", "① Wattonomics 覆蓋缺口 ④（上線 GW）", "GW", y, lo, b, hi, "Assumed",
        "原 1.84／4.81／8.28（算力收支 11-03，全 Assumed）；W-2（11-05）(a) 七家多花由 1.84 改實績下限 0.06（0–0.92）→ 各值減 1.78（高值減 0.92）；2028–29 延續（11-05 回執）。v0.3.2 後由 W-1 實績取代")
for y, (lo, b, hi) in GAP2.items():
    inp(f"cov2_{y}", "① W2 新雲、主權站點清單結束後年增歸零（上線 GW）", "GW", y, lo, b, hi, "Assumed", "算力收支 11-05 回執第 2 節（與 ④ 不重疊部分）")
# Siliconomics 證據比例（實績、已簽約、已確認站點支撐的比例；指引與預測不算）
for y, (lo, b, hi, src) in {
        2026: (0.597, 0.789, 0.892, "siliconomics/docs/reports/20261114_供給側證據比例.xlsx（NVIDIA 實績＋10-Q 採購承諾、Broadcom／AMD 實績）"),
        2027: (0.303, 0.448, 0.749, "同上（NVIDIA 採購承諾＋AMD／OpenAI 首 1 GW）"),
        2028: (0.161, 0.236, 0.541, "同上（NVIDIA 採購承諾）"),
        2029: (0.030, 0.079, 0.181, "同上（NVIDIA FY29、FY30 採購承諾）")}.items():
    inp(f"scov_{y}", "Siliconomics 證據比例（基準值中由實績、已簽約支撐的比例）", "比", y, lo, b, hi, "Assumed", src)
# W2-02 上線年 GW 交叉檢核用輸入（只檢核；查核確認為上線年口徑，見 Bridge 頁 p 列說明）
_itpg = cells(S, "SIF_ITCapexPerGW")  # 2026–2029 四年
for i, y in enumerate(YRS[:2]):
    inp(f"wsp_{y}", "Wattonomics IT 支出合計（WIF_Spend_it_sum）", "十億美元", y, None, cells(W, f"WIF_Spend_it_sum_{y}")[0], None, "Derived",
        f"{wt_f} WIF_Spend_it_sum_{y}（上線年口徑、不含非 AI、非 NVIDIA 已用 spend_nnv_hold 換算；wattonomics/docs/reports/20261118_W2_B_非NVIDIA廠房持有比口徑.md 第 4 節）")
    inp(f"itpg_{y}", "Siliconomics 每 GW IT 資本支出", "十億美元／GW", y, None, _itpg[i], None, "Derived",
        f"{sil_f} SIF_ITCapexPerGW（四年值取第 {i + 1} 個；與 Siliconomics 成本等值 GW 同一分母）")
# W2-02 r2 付款年讀數用輸入：ShipGW（出貨年實體 GW；七家與 xAI 付款年、其餘鏡像反推）與成本等值換算係數 capfx（Compare 頁，無具名範圍，以鍵列＋年份欄讀取）
_CP = W["Compare"]
_hdr = next(r for r in range(1, _CP.max_row + 1) if _CP.cell(r, 2).value == "鍵" and _CP.cell(r, 4).value == "2023")
_ycol = {int(_CP.cell(_hdr, c).value): c for c in range(4, _CP.max_column + 1) if str(_CP.cell(_hdr, c).value).isdigit()}
_capfx_r = next(r for r in range(1, _CP.max_row + 1) if _CP.cell(r, 2).value == "capfx")
for y in YRS[:2]:
    inp(f"wship_{y}", "Wattonomics 出貨年實體 GW（WIF_Recon_NV_ShipGW；七家與 xAI 付款年，其餘由上線年鏡像反推）", "GW", y, None, cells(W, f"WIF_Recon_NV_ShipGW_{y}")[0], None, "Derived",
        f"{wt_f} WIF_Recon_NV_ShipGW_{y}（Compare 第 F 節）")
    inp(f"wcapfx_{y}", "Wattonomics 成本等值換算係數 capfx（每 IT GW IT 資本支出 ÷ SIF_ITCapexPerGW）", "x", y, None, _CP.cell(_capfx_r, _ycol[y]).value, None, "Derived",
        f"{wt_f} Compare!{CL(_ycol[y])}{_capfx_r}（鍵 capfx；無具名範圍）")
# W2-03 2025 年底已出貨未上線存量（實體 GW）。寫單者建議，chat 端 2026-11-18 採用：[Assumed]，基準＝inv_on × ship_2025，低／高用 inv_on 的 5%／20%
inp("stk_2025", "2025 年底已出貨未上線存量（實體 GW）", "GW", 2025, "=inv_on_lo*ship_2025", "=inv_on*ship_2025", "=inv_on_hi*ship_2025", "Assumed",
    "inv_on（低 5%／基準 10%／高 20%）× ship_2025；無外部實績可檢驗，三個值都要顯示、不得只用單點（W2-03；chat 端 2026-11-18 採寫單者建議）")
for k, ref in N.items():
    wb.defined_names[k] = openpyxl.workbook.defined_name.DefinedName(k, attr_text=ref)
for j, w in enumerate([12, 52, 6, 6, 8, 9, 8, 10, 110], 1):
    IN.column_dimensions[CL(j)].width = w

# ---------------- Bridge ----------------
B = wb.create_sheet("Bridge")
B["A1"] = ("Bridge — Wattonomics 上線 GW 換成 Siliconomics 口徑（成本等值、出貨年），與現行基準、原始顆數並排；差距拆 C2 三類。"
           "換算（規則 v1.1 平滑法）：實體出貨ₜ ≈ 上線ₜ＋k ×（上線ₜ₊₁ − 上線ₜ），k＝lag_on÷12＋inv_on；再 ÷（實體 ÷ 成本等值）。"
           "供給上限只作檢查（v1.1，上限非獨立）。遞迴精確反推列保留為停用對照")
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
row("w4", "上線（補覆蓋）", "＝w3；供給上限只作檢查（v1.1）", lambda y: None if y == 2025 else f"={C(y)}{R['w3']}", True)
row("capchk", "　檢查：超過供給上限？", "上限非獨立，只標示不封頂", lambda y: None if y == 2025 else f'=IF({C(y)}{R["w4"]}>{C(y)}{R["cap"]},"超過","未超過")', fmt="@")
row("w4a", "上線（不補覆蓋）", "＝w1", lambda y: None if y == 2025 else f"={C(y)}{R['w1']}")
ship("s1", "停用對照：遞迴反推實體出貨（補覆蓋）", "v1.0 量法；2025＝ship_2025", lambda y: f"{C(y)}{R['w4']}", bold=True)
ship("s1a", "停用對照：遞迴反推（不補覆蓋）", "同上", lambda y: f"{C(y)}{R['w4a']}")
row("rt", "本庫 實體 ÷ 成本等值（出貨年）", "SIF_DeployGWPhys ÷ SIF_DeployGW；ASIC、AMD 每 GW 較便宜", lambda y: None if y == 2025 else f"=phys_{y}/sil_{y}", fmt="0.000")
row("x1", "停用對照：遞迴換算值（補覆蓋）", "s1 ÷ rt", lambda y: None if y == 2025 else f"={C(y)}{R['s1']}/{C(y)}{R['rt']}")
row("x1a", "停用對照：遞迴換算值（不補覆蓋）", "s1a ÷ rt", lambda y: None if y == 2025 else f"={C(y)}{R['s1a']}/{C(y)}{R['rt']}")
row("w4n", "　次年上線（補覆蓋；2029 用 Wattonomics 2030 值＋2029 覆蓋缺口）", "平滑法用",
    lambda y: None if y == 2025 else ("=wt_2030+cov4_2029+cov2_2029" if y == 2029 else f"={CL(COL[y]+1)}{R['w4']}"))
row("w4an", "　次年上線（不補覆蓋）", "", lambda y: None if y == 2025 else ("=wt_2030" if y == 2029 else f"={CL(COL[y]+1)}{R['w4a']}"))
row("s2", "平滑法實體出貨（補覆蓋）≈ 上線ₜ＋k ×（上線ₜ₊₁ − 上線ₜ）", "規則 v1.1", lambda y: None if y == 2025 else f"={C(y)}{R['w4']}+{C(y)}{R['k']}*({C(y)}{R['w4n']}-{C(y)}{R['w4']})", True)
row("s2a", "平滑法實體出貨（不補覆蓋）", "", lambda y: None if y == 2025 else f"={C(y)}{R['w4a']}+{C(y)}{R['k']}*({C(y)}{R['w4an']}-{C(y)}{R['w4a']})")
row("x2", "換算值（補覆蓋）＝上游定基準時的基準", "s2 ÷ rt", lambda y: None if y == 2025 else f"={C(y)}{R['s2']}/{C(y)}{R['rt']}", True, fill=KEY)
row("x2a", "換算值（不補覆蓋）", "s2a ÷ rt", lambda y: None if y == 2025 else f"={C(y)}{R['s2a']}/{C(y)}{R['rt']}")
row("x0", "對照：上線 GW 直接換成本等值（不做時差反推）", "w4a ÷ rt", lambda y: None if y == 2025 else f"={C(y)}{R['w4a']}/{C(y)}{R['rt']}")
br += 1
B.cell(br, 1, "上線年 GW 交叉檢核（成本等值；原工作單稱付款年，查核確認為上線年口徑）（只檢核，不進 C2 拆解；r2 A5。WIF_Spend_it_sum ÷ Siliconomics 每 GW IT 資本支出＝上線年成本等值 GW）").font = BOLD; br += 1
_p0 = br
_OK = (2026, 2027)
_PN = "；查核結論（wattonomics/docs/reports/20261118_W2_B_非NVIDIA廠房持有比口徑.md 第 4 節）：WIF_Spend_it_sum 為上線年口徑、不含非 AI、非 NVIDIA 已以 spend_nnv_hold 換算，除以 SIF_ITCapexPerGW 不算重複換算。r2 A5 只要求 2026–2027，其餘年留空"
row("p1", "上線年 GW 交叉檢核（成本等值；原工作單稱付款年，查核確認為上線年口徑）＝WIF_Spend_it_sum ÷ SIF_ITCapexPerGW", "[Derived]，只供核對量級" + _PN,
    lambda y: f"=wsp_{y}/itpg_{y}" if y in _OK else None, True)
row("p2", "　上線年 GW 交叉檢核（成本等值）p1 對換算值 x2 的差異％", "p1 ÷ x2 − 1" + _PN, lambda y: f"={C(y)}{_p0}/{C(y)}{R['x2']}-1" if y in _OK else None, fmt="0.0%")
row("p3", "　上線年 GW 交叉檢核（成本等值）p1 對供給側讀數 s0r 的差異％", "p1 ÷ s0r − 1" + _PN, lambda y: f"={C(y)}{_p0}/{C(y)}{_p0 + 12}-1" if y in _OK else None, fmt="0.0%")
row("p4", "　上線年 GW 交叉檢核（成本等值）p1 對 Siliconomics 基準 s0 的差異％", "p1 ÷ s0 − 1" + _PN, lambda y: f"={C(y)}{_p0}/{C(y)}{_p0 + 11}-1" if y in _OK else None, fmt="0.0%")
br += 1
B.cell(br, 1, "付款年讀數（r2 收尾；出貨年實體 GW × capfx 再成本等值化；只填 2026、2027；只檢核，不被任何既有列引用）").font = BOLD; br += 1
_q0 = br
_QN = "；來源：WIF_Recon_NV_ShipGW × Compare capfx（＝每 IT GW IT 資本支出 ÷ SIF_ITCapexPerGW；查核報告第 4 節(1) 指定算法）。2026–2027 以外留空"
row("q1", "付款年 GW（成本等值）＝WIF_Recon_NV_ShipGW × capfx", "[Derived]" + _QN, lambda y: f"=wship_{y}*wcapfx_{y}" if y in _OK else None, True)
row("q2", "　付款年讀數 q1 對換算值 x2 的差異％", "q1 ÷ x2 − 1" + _QN, lambda y: f"={C(y)}{_q0}/{C(y)}{R['x2']}-1" if y in _OK else None, fmt="0.0%")
row("q3", "　付款年讀數 q1 對供給側讀數 s0r 的差異％", "q1 ÷ s0r − 1" + _QN, lambda y: f"={C(y)}{_q0}/{C(y)}{_q0 + 6}-1" if y in _OK else None, fmt="0.0%")
row("q4", "　付款年讀數 q1 對 Siliconomics 基準 s0 的差異％", "q1 ÷ s0 − 1" + _QN, lambda y: f"={C(y)}{_q0}/{C(y)}{_q0 + 5}-1" if y in _OK else None, fmt="0.0%")
br += 1
row("s0", "Siliconomics 基準（SIF_DeployGW_Base）（成本等值、出貨年）", "SIF_DeployGW_Base＝原始 × S33 係數", lambda y: None if y == 2025 else f"=sil_{y}", True, fill=KEY)
row("s0r", "Siliconomics 原始顆數讀數（供給側讀數）", "現行基準 ÷ 係數", lambda y: None if y == 2025 else f"=sil_{y}/coef_{y}", True)
assert R["s0"] == _q0 + 5 and R["s0r"] == _q0 + 6, "q3／q4 的前向參照列位置與 s0／s0r 不符"
assert R["q1"] == _q0 and R["p1"] == _p0 and R["s0"] == _p0 + 11, "p3／p4 前向參照列位置不符"
br += 1
B.cell(br, 1, "C2 差距拆解（現行基準 − 上線直接換算 ＝ ① ＋ ② ＋ ③）").font = BOLD; br += 1
row("g", "總差距", "s0 − x0", lambda y: None if y == 2025 else f"={C(y)}{R['s0']}-{C(y)}{R['x0']}", True)
row("g2", "② 時差與庫存（單年可、累計不可）", "x2a − x0", lambda y: None if y == 2025 else f"={C(y)}{R['x2a']}-{C(y)}{R['x0']}")
row("g2c", "　② 累計（2026 起）", "累計不得長期偏離 0", lambda y: None if y == 2025 else (f"={C(y)}{R['g2']}" if y == 2026 else f"={CL(COL[y]-1)}{br}+{C(y)}{R['g2']}"))
row("g2p", "　② 累計 ÷ 當年基準（檢查第 7 項，≤ 15%）", "", lambda y: None if y == 2025 else f"={C(y)}{R['g2c']}/{C(y)}{R['s0']}", fmt="0.0%")
row("g1", "① 口徑與覆蓋（Assumed；非分歧）", "x2 − x2a", lambda y: None if y == 2025 else f"={C(y)}{R['x2']}-{C(y)}{R['x2a']}")
row("g3", "③ 殘差（分歧）", "s0 − x2", lambda y: None if y == 2025 else f"={C(y)}{R['s0']}-{C(y)}{R['x2']}", True, fill=KEY)
row("g3p", "　③ ÷ 現行基準", "容差 ±15%", lambda y: None if y == 2025 else f"={C(y)}{R['g3']}/{C(y)}{R['s0']}", fmt="0.0%")
row("g3t", "　容差內？", "", lambda y: None if y == 2025 else f'=IF(ABS({C(y)}{R["g3p"]})<=0.15,"一致","超出")', fmt="@")
row("chk", "　檢查：① ＋ ② ＋ ③ − 總差距（應為 0）", "", lambda y: None if y == 2025 else f"={C(y)}{R['g1']}+{C(y)}{R['g2']}+{C(y)}{R['g3']}-{C(y)}{R['g']}", fmt="0.000")
br += 1
B.cell(br, 1, "敏感度：平滑法換算值 x2（一次動一個輸入）").font = BOLD; br += 1
for key, lab, kexpr, cs in [("sl", "lag_on 低（3 月）", "(lag_on_lo/12+inv_on)", ""), ("sh", "lag_on 高（6 月）", "(lag_on_hi/12+inv_on)", ""),
                            ("il", "inv_on 低（5%）", "(lag_on/12+inv_on_lo)", ""), ("ih", "inv_on 高（20%）", "(lag_on/12+inv_on_hi)", ""),
                            ("cl", "覆蓋缺口 低", "(lag_on/12+inv_on)", "_lo"), ("ch", "覆蓋缺口 高", "(lag_on/12+inv_on)", "_hi")]:
    def onx(y, cs=cs):
        return f"(wt_{y}+cov4_{y}{cs}+cov2_{y}{cs})"
    def nxt(y, cs=cs):
        return f"(wt_2030+cov4_2029{cs}+cov2_2029{cs})" if y == 2029 else onx(y + 1, cs)
    row("sx_" + key, "換算值：" + lab, "", lambda y, kx=kexpr, o=onx, n=nxt: None if y == 2025 else f"=({o(y)}+{kx}*({n(y)}-{o(y)}))/{C(y)}{R['rt']}")
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
trow("ew", "Wattonomics 證據比例（全市場分母；v1.1）", "v0.3 以前：WIF_Coverage × 主路徑 ÷（主路徑＋覆蓋缺口）；v0.3.2 起覆蓋缺口＝0，即 WIF_Evidence", lambda y: f"=wcov_{y}*wt_{y}/(wt_{y}+cov4_{y}+cov2_{y})", "0%")
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
trow("b50r", "停用對照：v1.0 遞迴反推換算", "", lambda y: f'=IF({up(TR["t50"], "thr")(y)},Bridge!{C(y)}{R["x1"]},sil_{y})')
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

# ---------------- Stock（W2-03：出貨到上線存量帳；單位＝實體 GW） ----------------
K = wb.create_sheet("Stock")
K["A1"] = ("Stock — 出貨到上線的存量帳（單位：實體 GW）。恆等式：期末存量＝期初存量＋出貨 − 上線；不假設時差或庫存比例推存量，不設上限、不截斷（可為負值，負值代表上線快於出貨、須檢討輸入）。"
           "2025 年底存量 stk_2025 為 [Assumed]（低 inv_on 5%／基準 10%／高 20% × ship_2025，無外部實績），三組都顯示。"
           "A 節＝需求側自洽（Bridge s2 平滑法實體出貨對 w4 上線）；B 節＝若 Siliconomics 的出貨都是真的（SIF_DeployGWPhys_Base 對 w4）；C 節＝現行單一係數 k 隱含的存量對照")
K["A1"].font = BOLD
for j, h in enumerate(["列", "項目", "說明", "2025"] + [str(y) for y in YRS], 1):
    c = K.cell(3, j, h); c.font = BOLD; c.fill = HDR
kr = 4; KR = {}
SC = [("lo", "低"), ("b", "基準"), ("hi", "高")]
SK = {"lo": "stk_2025_lo", "b": "stk_2025", "hi": "stk_2025_hi"}


def krow(key, label, note, f, bold=False, fmt="0.00", fill=None):
    global kr
    K.cell(kr, 1, key); K.cell(kr, 2, label); K.cell(kr, 3, note)
    for y in [2025] + YRS:
        v = f(y)
        if v is not None:
            c = K.cell(kr, COL[y], v); c.number_format = fmt
            if fill: c.fill = fill
    if bold:
        for j in (1, 2):
            K.cell(kr, j).font = BOLD
    KR[key] = kr; kr += 1


def ksec(title):
    global kr
    K.cell(kr, 1, title).font = BOLD; kr += 1


def stock_section(tag, title, ship_label, ship_note, ship_f):
    ksec(title)
    krow(tag + "_ship", ship_label, ship_note, lambda y: None if y == 2025 else ship_f(y), True)
    krow(tag + "_on", "上線（Bridge w4，補覆蓋；Wattonomics 年新增上線 IT GW）", "實體 GW", lambda y: None if y == 2025 else f"=Bridge!{C(y)}{R['w4']}", True)
    krow(tag + "_open", "期初存量（基準情境）", "＝上年期末（基準）", lambda y: None if y == 2025 else f"={CL(COL[y] - 1)}{kr + 2}")
    for sk, sn in SC:
        krow(f"{tag}_end_{sk}", f"期末存量（{sn}）＝期初＋出貨 − 上線", ("2025＝stk_2025 " + sn if sk == "b" else "同上，2025 起點用 stk_2025 " + sn),
             lambda y, sk=sk: f"={SK[sk]}" if y == 2025 else f"={CL(COL[y] - 1)}{kr}+{C(y)}{KR[tag + '_ship']}-{C(y)}{KR[tag + '_on']}",
             sk == "b", fill=KEY if sk == "b" else None)
    krow(tag + "_chk", "　守恆檢查（應為 0）：期末 −（stk_2025＋累計出貨 − 累計上線），取三情境絕對值最大者", "累計式與逐年遞推獨立計算",
         lambda y: None if y == 2025 else "=MAX(" + ",".join(
             f"ABS({C(y)}{KR[tag + '_end_' + sk]}-({SK[sk]}+SUM($E{KR[tag + '_ship']}:{C(y)}{KR[tag + '_ship']})-SUM($E{KR[tag + '_on']}:{C(y)}{KR[tag + '_on']})))"
             for sk, _ in SC) + ")", fmt="0.000000")
    krow(tag + "_mo", "　期末存量（基準）÷ 當年出貨", "存量相當於幾年的出貨（只顯示）", lambda y: None if y == 2025 else f"={C(y)}{KR[tag + '_end_b']}/{C(y)}{KR[tag + '_ship']}", fmt="0.0%")


stock_section("a", "A 節　需求側自洽（Bridge s2 平滑法實體出貨 對 w4 上線）", "實體出貨（Bridge s2，平滑法、補覆蓋）", "Wattonomics 上線反推，k＝lag_on÷12＋inv_on",
              lambda y: f"=Bridge!{C(y)}{R['s2']}")
kr += 1
stock_section("b", "B 節　供給側出貨對需求側上線（若 Siliconomics 的出貨都是真的，Wattonomics 的上線速度意味著累積多少未上線存量）", "實體出貨（SIF_DeployGWPhys_Base，In 頁 phys）", "Siliconomics 現行基準的實體部署 GW（出貨年）",
              lambda y: f"=phys_{y}")
kr += 1
ksec("C 節　單一係數對照（現行做法隱含的存量＝inv_on × 出貨 ＋ lag_on÷12 × 出貨＝k × 出貨）")
krow("c_ka", "係數隱含存量（A 節出貨）", "k × s2；k 用基準（Wattonomics 權威，J3）", lambda y: None if y == 2025 else f"=Bridge!{C(y)}{R['k']}*{C(y)}{KR['a_ship']}")
krow("c_kb", "係數隱含存量（B 節出貨）", "k × phys", lambda y: None if y == 2025 else f"=Bridge!{C(y)}{R['k']}*{C(y)}{KR['b_ship']}")
for tag, nm in (("a", "A"), ("b", "B")):
    for sk, sn in SC:
        krow(f"c_r{tag}_{sk}", f"恆等式存量 ÷ 係數隱含存量（{nm} 節；stk_2025 {sn}）", "期末存量（恆等式）÷ 係數隱含存量；1＝與單一係數一致",
             lambda y, tag=tag, sk=sk: None if y == 2025 else f"={C(y)}{KR[tag + '_end_' + sk]}/{C(y)}{KR['c_k' + tag]}", fmt="0.00", fill=KEY if sk == "b" else None)
K.column_dimensions["A"].width = 10; K.column_dimensions["B"].width = 70; K.column_dimensions["C"].width = 44
for y in [2025] + YRS:
    K.column_dimensions[C(y)].width = 10

# ---------------- Residual（W2-04A：殘差帳骨架；歸因項讀種子檔） ----------------
SEED = os.path.join(ROOT, "siliconomics/builder/seed/residual_items_v0.1.json")
seed = json.load(open(SEED, encoding="utf-8"))
items = seed.get("items", [])
GRADES = {"Verified", "Interested-party", "Analogy", "Assumed", "Derived"}
for it in items:  # 查不到來源者不得列入；缺欄位直接報錯
    miss = [k for k in ("name", "gw", "evidence_grade", "testable_actual", "source", "workorder") if not it.get(k)]
    if miss or it["evidence_grade"] not in GRADES:
        raise SystemExit(f"殘差歸因項不合格（缺 {miss} 或等級不在 {sorted(GRADES)}）：{it.get('name')}")
D = wb.create_sheet("Residual")
D["A1"] = (f"Residual — 殘差帳：Bridge ③ 殘差（g3）逐項歸因，最後一列「未歸因殘差」永遠顯示（不得設為 0 或隱藏）。種子檔 residual_items_v{seed.get('version')}.json 共 {len(items)} 項。"
           "不設任何結案百分比門檻（r2 不採）；歸因項只可因有來源與可檢驗實績而列入，不得為了讓殘差變小而列入。單位：成本等值 GW")
D["A1"].font = BOLD
for j, h in enumerate(["列", "項目", "說明"] + [str(y) for y in YRS] + ["證據等級", "可檢驗的實績", "來源", "對應工作單"], 1):
    c = D.cell(3, j, h); c.font = BOLD; c.fill = HDR
dr = 4; DR = {}
DC = {y: 4 + i for i, y in enumerate(YRS)}
for j, y in enumerate(YRS):
    pass
D.cell(dr, 1, "g3"); D.cell(dr, 2, "③ 殘差（Bridge g3＝現行基準 − 換算值）").font = BOLD; D.cell(dr, 3, "分歧部分；正值＝Siliconomics 高於 Wattonomics 換算值")
for y in YRS:
    c = D.cell(dr, DC[y], f"=Bridge!{C(y)}{R['g3']}"); c.number_format = "0.00"; c.fill = KEY
DR["g3"] = dr; dr += 1
first_item = dr
for n, it in enumerate(items, 1):
    D.cell(dr, 1, f"i{n}"); D.cell(dr, 2, it["name"]); D.cell(dr, 3, "歸因項（種子檔）")
    for y in YRS:
        v = it["gw"].get(str(y))
        if v is not None:
            c = D.cell(dr, DC[y], v); c.number_format = "0.00"; c.font = BLUE
    for j, k in zip((8, 9, 10, 11), ("evidence_grade", "testable_actual", "source", "workorder")):
        D.cell(dr, j, it[k])
    dr += 1
last_item = dr - 1
D.cell(dr, 1, "sum"); D.cell(dr, 2, "歸因項合計"); D.cell(dr, 3, "空清單時為 0")
for y in YRS:
    cl = CL(DC[y])
    c = D.cell(dr, DC[y], f"=SUM({cl}{first_item}:{cl}{last_item})" if items else 0); c.number_format = "0.00"
DR["sum"] = dr; dr += 1
D.cell(dr, 1, "un"); D.cell(dr, 2, "未歸因殘差＝g3 − 歸因項合計").font = BOLD; D.cell(dr, 3, "永遠顯示；沒有結案門檻")
for y in YRS:
    cl = CL(DC[y])
    c = D.cell(dr, DC[y], f"={cl}{DR['g3']}-{cl}{DR['sum']}"); c.number_format = "0.00"; c.fill = KEY; c.font = BOLD
DR["un"] = dr; dr += 1
D.column_dimensions["A"].width = 6; D.column_dimensions["B"].width = 46; D.column_dimensions["C"].width = 44
for y in YRS:
    D.column_dimensions[CL(DC[y])].width = 12
for col, w in zip("HIJK", (14, 40, 40, 14)):
    D.column_dimensions[col].width = w

out = sys.argv[1]
wb.save(out)
print("wrote", out)
