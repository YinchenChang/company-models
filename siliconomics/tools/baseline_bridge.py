# 三層基準協調 橋接表（步驟 3 量測；規則 docs/plan/三層基準協調規則.md v1.1：平滑法、乙前提、全市場分母）
# 方向：Wattonomics 上線 GW →（C2 ② 時差與庫存反推）實體出貨 GW →（本庫實體／成本等值比）成本等值出貨 GW，
#       與 Siliconomics 現行基準、原始顆數讀數（供給側）並排；差距拆 C2 三類；每年標 C3 狀態（門檻 50%、67%）。
# 用法：python3 siliconomics/tools/baseline_bridge.py <輸出 xlsx>   之後以 LibreOffice 重算。本工具不改任何模型。
# 2026-11-18 r2 收尾：p1–p4 改標「上線年 GW 交叉檢核」（查核確認口徑）；新增 q1–q4 付款年讀數；s0 標籤改「Siliconomics 基準（SIF_DeployGW_Base）」；不改既有數值。
# 2026-11-18（三層對帳第二波 W2-02／W2-03／W2-04A；不改任何既有數值）：Bridge 頁在 x0 與 s0 之間新增「付款年 GW 交叉檢核」p1–p4（只檢核，不被任何既有列引用）；
#   新增 Stock 頁（出貨到上線存量帳，A／B／C 三節，stk_2025 低／基準／高三組）與 Residual 頁（殘差帳骨架，歸因項讀 builder/seed/residual_items_v0.1.json）。
# 2026-11-18（三層對帳第三波 W-SF-01／04／05；不改任何既有數值）：Stock 頁新增 D–G 節（通路 S1、買方端 S2 兩段存量帳，內生停留時間，與單一係數並列，stk_2025 自洽版對照）；新增 Check 頁（通路容量、金額、改放買方端三個核對，只輸出放得下／放不下多少 GW，不設門檻）與 ChanObs、BuyerObs、OdmMonthly 三個觀測頁（讀 builder/seed/channel_obs_v0.1.json）。
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
# ---- W-SF-01／04／05（2026-11-18 第三波）輸入：種子檔與單值參數（不改任何既有輸入）----
CHS = os.path.join(ROOT, "siliconomics/builder/seed/channel_obs_v0.1.json")
_chs_txt = open(CHS, encoding="utf-8").read()
if "107.6" in _chs_txt:   # 共同規則第 1 條：NVIDIA FY2027 第三季實績不得進任何種子檔
    raise SystemExit("channel_obs 種子檔含 107.6（疑為 NVIDIA FY2027 第三季實績）：停止，請回報，不得自行刪改")
chs = json.loads(_chs_txt)
_G5 = {"Verified", "Interested-party", "Analogy", "Assumed", "Derived"}
for _r in chs["odm_quarterly"] + chs["buyer_quarterly"]:
    _miss = [k for k in ("company", "quarter", "evidence_grade", "source_ids") if not _r.get(k)]
    if _miss or _r["evidence_grade"] not in _G5 or any(s not in chs["sources"] for s in _r["source_ids"]):
        raise SystemExit(f"channel_obs 觀測列不合格（缺 {_miss}、等級不在五種或來源 ID 不在 sources）：{_r.get('company')} {_r.get('quarter')}")
if chs["monthly_revenue"]["evidence_grade"] not in _G5:
    raise SystemExit("channel_obs 月營收等級不合格")
_rc = S["Recon"]
assert str(_rc.cell(3, 5).value) == "2027", "Siliconomics Recon 頁欄位位置改變（E 欄應為 2027）"
_nvr = next(r_ for r_ in range(1, _rc.max_row + 1) if str(_rc.cell(r_, 1).value).startswith("隱含 NVIDIA 資料中心營收（層級"))
inp("wship_2025", "Wattonomics 出貨年實體 GW 2025（WIF_Recon_NV_ShipGW；階段 ② 起點）", "GW", 2025, None, cells(W, "WIF_Recon_NV_ShipGW_2025")[0], None, "Derived",
    f"{wt_f} WIF_Recon_NV_ShipGW_2025（只供 Stock D–G 節）")
inp("s1sh", "2025 年底存量中屬通路（S1）的比例", "比", 2025, 0.3, 0.5, 0.7, "Assumed",
    "無依據（設計書 W-SF-01）；只影響起點在兩段之間的分配，S1＋S2 合計不變；2027 年 S1、S2 各約差 ±0.3 GW 以內")
inp("ai_sh", "AI 在組裝廠存貨與營收中的占比（四家台廠與 Dell、Supermicro 共用；公司都沒拆）", "比", None, 0.35, "=2719.19/3998.81", 0.85, "Assumed",
    "基準＝廣達模型 FY2026 AI 伺服器營收 ÷ 營收（Derived 0.68，非公司揭露）套用到全部公司是假設；低＝公司說法中最低的 Dell（AI 伺服器營收 16.4B ÷ 46.97B＝35%，二手）；高 0.85 無依據（判斷上限）。"
    "同一組區間用於存貨換算（Check）與月營收換算（OdmMonthly，只顯示）")
inp("fxtw", "匯率 NT$／US$", "NT$/US$", None, cells(S, "fx_twd_lo")[0], cells(S, "fx_twd")[0], cells(S, "fx_twd_hi")[0], "Assumed", f"{sil_f} In fx_twd（低 30.5／基準 32／高 33）")
inp("qsh_26", "廣達 AI 伺服器組裝份額 2026（只供原型對照列）", "比", 2026, None, cells(S, "asm_qush_26")[0], None, "Derived", f"{sil_f} In asm_qush_26（ShareTbl；2026 公司模型校準）")
inp("asmsum_26", "四家台廠有效組裝份額合計 2026（SIF_AsmShareSum）", "比", 2026, None, cells(S, "SIF_AsmShareSum")[0], None, "Derived", f"{sil_f} SIF_AsmShareSum（鴻海、廣達、緯創、緯穎）")
_odm = cells(S, "SIF_ODMValPerGW")
inp("odm_26", "每 GW ODM 整櫃價值 2026", "十億美元／GW", 2026, None, _odm[0], None, "Derived", f"{sil_f} SIF_ODMValPerGW（Tokenomics 每 GW 架數 × ODM 每櫃價格）")
inp("odm_27", "每 GW ODM 整櫃價值 2027", "十億美元／GW", 2027, None, _odm[1], None, "Derived", f"{sil_f} SIF_ODMValPerGW")
inp("nv_sil27", "Siliconomics 隱含 NVIDIA 資料中心營收 CY2027（層級值，含指引；供顯示）", "十億美元", 2027, None, _rc.cell(_nvr, 5).value, None, "Derived",
    f"{sil_f} Recon!E{_nvr}（與 2027 年對帳單凍結紀錄同值；未使用 NVIDIA FY2027 第三季實績）")
inp("nv_wt27", "Wattonomics 隱含 NVIDIA 資料中心營收 CY2027", "十億美元", 2027, None, cells(W, "WIF_Recon_NV_ImpliedRev_2027")[0], None, "Derived", f"{wt_f} WIF_Recon_NV_ImpliedRev_2027")
inp("nvcpg27", "NVIDIA 每 GW 內容 2027（營收口徑）", "十億美元／GW", 2027, cells(S, "SIF_NVContentGW_VR200_lo")[0], cells(W, "WIF_Recon_NV_ContentPerGW_2027")[0], cells(S, "SIF_NVContentGW_VR200_hi")[0],
    "Derived", f"基準＝{wt_f} WIF_Recon_NV_ContentPerGW_2027；低／高＝{sil_f} SIF_NVContentGW_VR200_lo／_hi（VR200 內容區間）")
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
# ---- W-SF-01：Stock 頁新增 D–G 節（原有 A、B、C 節不動；單位：實體 GW）----
kr += 1
ksec("D 節　通路（晶片商出貨 → 買方收貨）S1：期末＝期初＋① 晶片商出貨 − ② 買方收貨；不設上限、不截斷；內生停留時間（月）＝期末存量 ÷ 當年流出 × 12")
krow("d_ship", "① 晶片商出貨（認列營收；供給側）", "In phys（Siliconomics 實體部署 GW）；2025＝ship_2025 [Derived]", lambda y: "=ship_2025" if y == 2025 else f"=phys_{y}", True)
krow("d_pay", "② 買方收貨（付款年；需求側）", "2025–27＝WIF_Recon_NV_ShipGW；2028–29 無名稱，取 Bridge s2（平滑法同式鏡像）[Derived]",
     lambda y: f"=wship_{y}" if y <= 2027 else f"=Bridge!{C(y)}{R['s2']}", True)
krow("d_on", "③ 上線（通電；Bridge w4，補覆蓋；同 A、B 節）", "實體 GW", lambda y: None if y == 2025 else f"=Bridge!{C(y)}{R['w4']}", True)
krow("d_open", "S1 期初存量（基準）", "＝上年期末（基準）", lambda y: None if y == 2025 else f"={CL(COL[y] - 1)}{kr + 2}")
for sk, sn in SC:
    krow(f"d_end_{sk}", f"S1 通路期末存量（{sn}）＝期初＋① − ②", "2025＝stk_2025 " + sn + " × s1sh（Assumed 比例）",
         lambda y, sk=sk: f"={SK[sk]}*s1sh" if y == 2025 else f"={CL(COL[y] - 1)}{kr}+{C(y)}{KR['d_ship']}-{C(y)}{KR['d_pay']}",
         sk == "b", fill=KEY if sk == "b" else None)
krow("d_chk", "　守恆檢查（應為 0）：S1 期末 −（起點 S1＋累計① − 累計②），取三情境絕對值最大者", "累計式與逐年遞推獨立計算",
     lambda y: None if y == 2025 else "=MAX(" + ",".join(
         f"ABS({C(y)}{KR['d_end_' + sk]}-({SK[sk]}*s1sh+SUM($E{KR['d_ship']}:{C(y)}{KR['d_ship']})-SUM($E{KR['d_pay']}:{C(y)}{KR['d_pay']})))" for sk, _ in SC) + ")", fmt="0.000000")
krow("d_tau", "　內生停留 τ1（月）＝S1 期末（基準）÷ ② × 12", "現行做法隱含 0（橋接表 B 節把出貨直接對上線）", lambda y: None if y == 2025 else f"={C(y)}{KR['d_end_b']}/{C(y)}{KR['d_pay']}*12", fmt="0.0", fill=KEY)
krow("d_tau_lo", "　　τ1（月；stk_2025 低）", "", lambda y: None if y == 2025 else f"={C(y)}{KR['d_end_lo']}/{C(y)}{KR['d_pay']}*12", fmt="0.0")
krow("d_tau_hi", "　　τ1（月；stk_2025 高）", "", lambda y: None if y == 2025 else f"={C(y)}{KR['d_end_hi']}/{C(y)}{KR['d_pay']}*12", fmt="0.0")
kr += 1
ksec("E 節　買方端（買方收貨 → 通電）S2：期末＝期初＋② 買方收貨 − ③ 上線；內生停留時間（月）＝期末存量 ÷ 當年流出 × 12")
krow("e_pay", "② 買方收貨（流入；同 D 節）", "", lambda y: None if y == 2025 else f"={C(y)}{KR['d_pay']}", True)
krow("e_on", "③ 上線（流出；同 D 節）", "", lambda y: None if y == 2025 else f"={C(y)}{KR['d_on']}", True)
krow("e_open", "S2 期初存量（基準）", "＝上年期末（基準）", lambda y: None if y == 2025 else f"={CL(COL[y] - 1)}{kr + 2}")
for sk, sn in SC:
    krow(f"e_end_{sk}", f"S2 買方端期末存量（{sn}）＝期初＋② − ③", "2025＝stk_2025 " + sn + " ×（1 − s1sh）",
         lambda y, sk=sk: f"={SK[sk]}*(1-s1sh)" if y == 2025 else f"={CL(COL[y] - 1)}{kr}+{C(y)}{KR['e_pay']}-{C(y)}{KR['e_on']}",
         sk == "b", fill=KEY if sk == "b" else None)
krow("e_chk", "　守恆檢查（應為 0）：S2 期末 −（起點 S2＋累計② − 累計③），取三情境絕對值最大者", "累計式與逐年遞推獨立計算",
     lambda y: None if y == 2025 else "=MAX(" + ",".join(
         f"ABS({C(y)}{KR['e_end_' + sk]}-({SK[sk]}*(1-s1sh)+SUM($E{KR['e_pay']}:{C(y)}{KR['e_pay']})-SUM($E{KR['e_on']}:{C(y)}{KR['e_on']})))" for sk, _ in SC) + ")", fmt="0.000000")
krow("e_tau", "　內生停留 τ2（月）＝S2 期末（基準）÷ ③ × 12", "現行單一係數 k 全部放在這一段（見 F 節）", lambda y: None if y == 2025 else f"={C(y)}{KR['e_end_b']}/{C(y)}{KR['e_on']}*12", fmt="0.0", fill=KEY)
krow("e_tau_lo", "　　τ2（月；stk_2025 低）", "", lambda y: None if y == 2025 else f"={C(y)}{KR['e_end_lo']}/{C(y)}{KR['e_on']}*12", fmt="0.0")
krow("e_tau_hi", "　　τ2（月；stk_2025 高）", "", lambda y: None if y == 2025 else f"={C(y)}{KR['e_end_hi']}/{C(y)}{KR['e_on']}*12", fmt="0.0")
kr += 1
ksec("F 節　兩段合計對 B 節守恆，並與現行單一係數（lag_on、inv_on、k）隱含值並列（只顯示）")
krow("f_tot", "S1＋S2 合計（基準）", "應等於 B 節期末存量（基準）", lambda y: f"={C(y)}{KR['d_end_b']}+{C(y)}{KR['e_end_b']}", True)
krow("f_chk", "　守恆檢查（應為 0）：（S1＋S2）− B 節期末存量，取三情境絕對值最大者", "兩段加總須回到既有 B 節",
     lambda y: "=MAX(" + ",".join(f"ABS({C(y)}{KR['d_end_' + sk]}+{C(y)}{KR['e_end_' + sk]}-{C(y)}{KR['b_end_' + sk]})" for sk, _ in SC) + ")", fmt="0.000000")
krow("f_lag", "現行 lag_on（月）", "Wattonomics 權威（J3）；現行只放在買方端", lambda y: None if y == 2025 else "=lag_on", fmt="0.0")
krow("f_inv", "現行 inv_on × 12（月）", "同上", lambda y: None if y == 2025 else "=inv_on*12", fmt="0.0")
krow("f_k", "現行 k（月）＝lag_on＋inv_on × 12", "全部放買方端；通路隱含 0", lambda y: None if y == 2025 else "=lag_on+inv_on*12", fmt="0.0")
krow("f_t1", "內生 τ1 通路（月）", "D 節", lambda y: None if y == 2025 else f"={C(y)}{KR['d_tau']}", fmt="0.0")
krow("f_t2", "內生 τ2 買方端（月）", "E 節", lambda y: None if y == 2025 else f"={C(y)}{KR['e_tau']}", fmt="0.0")
krow("f_tt", "內生 τ 合計＝（S1＋S2）÷ ③ × 12（月）", "若硬把兩段塞進單一係數 k 要多大", lambda y: None if y == 2025 else f"={C(y)}{KR['f_tot']}/{C(y)}{KR['d_on']}*12", fmt="0.0")
krow("f_kr", "　內生 τ2 ÷ 現行 k", "1＝買方端與單一係數一致", lambda y: None if y == 2025 else f"={C(y)}{KR['f_t2']}/{C(y)}{KR['f_k']}", fmt="0.00")
kr += 1
ksec("G 節　stk_2025 自洽版對照（含在途量；只當對照，不改 stk_2025）：2025 年底 S2＝k × ②₂₀₂₅（與單一係數自洽）、S1＝τ1obs ÷ 12 × ①₂₀₂₅（τ1obs＝ChanObs 最近季存貨天數，低／基準／高＝六家最小／廣達／六家最大）")
T1N = {"lo": "t1o_lo", "b": "t1o", "hi": "t1o_hi"}
for sk, sn in SC:
    krow(f"g_s1a_{sk}", f"S1 通路期末存量（自洽版；τ1obs {sn}）", "2025＝τ1obs ÷ 12 × ①₂₀₂₅；之後同 D 節遞推",
         lambda y, sk=sk: f"={T1N[sk]}/12*{C(y)}{KR['d_ship']}" if y == 2025 else f"={CL(COL[y] - 1)}{kr}+{C(y)}{KR['d_ship']}-{C(y)}{KR['d_pay']}",
         sk == "b", fill=KEY if sk == "b" else None)
krow("g_s2a", "S2 買方端期末存量（自洽版）", "2025＝（lag_on÷12＋inv_on）× ②₂₀₂₅；之後同 E 節遞推",
     lambda y: f"=(lag_on/12+inv_on)*{C(y)}{KR['d_pay']}" if y == 2025 else f"={CL(COL[y] - 1)}{kr}+{C(y)}{KR['e_pay']}-{C(y)}{KR['e_on']}", True)
for sk, sn in SC:
    krow(f"g_tot_{sk}", f"S1＋S2 合計（自洽版；τ1obs {sn}）", "", lambda y, sk=sk: f"={C(y)}{KR['g_s1a_' + sk]}+{C(y)}{KR['g_s2a']}", sk == "b")
krow("g_diff", "　自洽版合計（基準）− B 節期末存量（基準）", "＝補上在途量後多出的量；2025 欄＝自洽版起點 − stk_2025", lambda y: f"={C(y)}{KR['g_tot_b']}-{C(y)}{KR['b_end_b']}", fmt="0.00")
krow("g_chk", "　守恆檢查（應為 0）：自洽版合計 −（自洽版 2025 合計＋累計① − 累計③），取三情境絕對值最大者", "",
     lambda y: None if y == 2025 else "=MAX(" + ",".join(
         f"ABS({C(y)}{KR['g_tot_' + sk]}-($D{KR['g_tot_' + sk]}+SUM($E{KR['d_ship']}:{C(y)}{KR['d_ship']})-SUM($E{KR['d_on']}:{C(y)}{KR['d_on']})))" for sk, _ in SC) + ")", fmt="0.000000")
krow("g_t1", "　內生 τ1（月；自洽版，基準）", "", lambda y: f"={C(y)}{KR['g_s1a_b']}/{C(y)}{KR['d_pay']}*12", fmt="0.0")
krow("g_t2", "　內生 τ2（月；自洽版）", "≈ 現行 k（月）：需求側本身自洽", lambda y: None if y == 2025 else f"={C(y)}{KR['g_s2a']}/{C(y)}{KR['d_on']}*12", fmt="0.0")
krow("g_lagimp", "　反推隱含 lag_on（月）＝自洽版 τ2 − inv_on × 12", "inv_on 維持現值時，買方端存量隱含的 lag_on；兩者拆不開，只顯示", lambda y: None if y == 2025 else f"={C(y)}{KR['g_t2']}-inv_on*12", fmt="0.0")
krow("g_invimp", "　反推隱含 inv_on＝自洽版 S2 ÷ ② − lag_on ÷ 12", "lag_on 維持現值時；只顯示", lambda y: None if y == 2025 else f"={C(y)}{KR['g_s2a']}/{C(y)}{KR['e_pay']}-lag_on/12", fmt="0.000")
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

# ================= 新增頁：Check、ChanObs、BuyerObs、OdmMonthly（W-SF-04／05；只顯示、不被任何既有列引用；不設任何通過門檻）=================
NUM = "#,##0"
Z = wb.create_sheet("Check")      # 先佔位置，內容最後填
H = wb.create_sheet("ChanObs")
BY = wb.create_sheet("BuyerObs")
OM = wb.create_sheet("OdmMonthly")
SRC = chs["sources"]
_sc_txt = lambda ids: "；".join(f"{i}（{SRC[i]['desc'][:40]}）" for i in ids)

# ---------- ChanObs ----------
H["A1"] = (f"ChanObs — 通路（ODM／OEM）季報觀測（種子檔 channel_obs_v{chs['version']}.json 逐筆；藍字＝觀測值）。存貨天數＝期末存貨 ÷ 當季營業成本 × 當季天數；通路停留（月）＝期末存貨 ÷ 當季營業成本 × 3。"
           "口徑是全公司，含非 AI、含尚未上板的原料與零組件，所以當作 AI 通路停留的上界（偏寬）。AI 與非 AI 公司都沒拆。廣達、鴻海、緯創為二手（S&P Global 經 stockanalysis.com，公司原文未取得，標 Interested-party）；緯穎為公司財報 PDF（Verified）；"
           "Dell、Supermicro 為 SEC XBRL 申報值（Verified；由年度累計相減推得者標 Derived）。Dell FQ3 FY27 截至 2026-11-18 尚未公布，最近季為 2026Q2。未使用 NVIDIA FY2027 第三季實績")
H["A1"].font = BOLD
for j, h_ in enumerate(["鍵", "公司", "曆季", "財報期末", "單位", "營收", "營業成本", "期末存貨", "當季天數", "存貨天數（日）", "通路停留（月）", "等級", "來源類別", "來源ID", "備註（種子檔）"], 1):
    c = H.cell(3, j, h_); c.font = BOLD; c.fill = HDR
hr = 4
LAST = {}   # 公司 → 最近季列號
for n, o in enumerate(chs["odm_quarterly"], 1):
    H.cell(hr, 1, f"o{n}"); H.cell(hr, 2, o["company"]); H.cell(hr, 3, o["quarter"]); H.cell(hr, 4, o["period_end"]); H.cell(hr, 5, o["unit"])
    for j, k in ((6, "revenue"), (7, "cogs"), (8, "inventory"), (9, "days_in_quarter")):
        c = H.cell(hr, j, o[k]); c.font = BLUE; c.number_format = NUM
    c = H.cell(hr, 10, f"=H{hr}/G{hr}*I{hr}"); c.number_format = "0.0"
    c = H.cell(hr, 11, f"=H{hr}/G{hr}*3"); c.number_format = "0.00"
    H.cell(hr, 12, o["evidence_grade"]); H.cell(hr, 13, o["source_class"]); H.cell(hr, 14, "；".join(o["source_ids"])); H.cell(hr, 15, o["note"][:200])
    LAST[o["company"]] = hr
    hr += 1
hr += 1
H.cell(hr, 1, "最近季摘要（每家一列；停留月與存貨天數直接引用上表）").font = BOLD; hr += 1
for j, h_ in enumerate(["鍵", "公司", "最近季", "單位", "期末存貨（本幣百萬）", "期末存貨（US$B，匯率取基準）", "通路停留（月）", "存貨天數（日）", "等級"], 1):
    c = H.cell(hr, j, h_); c.font = BOLD; c.fill = HDR
hr += 1
FIRMS = ["廣達 2382", "鴻海 2317", "緯創 3231", "緯穎 6669", "Dell", "Supermicro"]
CS = {}   # 公司 → 摘要列號
for n, f_ in enumerate(FIRMS, 1):
    s_ = LAST[f_]
    H.cell(hr, 1, f"s{n}"); H.cell(hr, 2, f_); H.cell(hr, 3, f"=C{s_}"); H.cell(hr, 4, f"=E{s_}")
    c = H.cell(hr, 5, f"=H{s_}"); c.number_format = NUM
    c = H.cell(hr, 6, f"=E{hr}/1000/fxtw" if f_.split()[0] in ("廣達", "鴻海", "緯創", "緯穎") else f"=E{hr}/1000"); c.number_format = "0.0"
    c = H.cell(hr, 7, f"=K{s_}"); c.number_format = "0.00"
    c = H.cell(hr, 8, f"=J{s_}"); c.number_format = "0.0"
    H.cell(hr, 9, f"=L{s_}")
    CS[f_] = hr; hr += 1
_g = lambda: f"G{CS[FIRMS[0]]}:G{CS[FIRMS[-1]]}"
H.cell(hr, 1, "t_lo"); H.cell(hr, 2, "六家最近季通路停留 最小（τ1obs 低）"); c = H.cell(hr, 7, f"=MIN({_g()})"); c.number_format = "0.00"; c.fill = KEY; T_LO = hr; hr += 1
H.cell(hr, 1, "t_b"); H.cell(hr, 2, "廣達 3Q26（τ1obs 基準；沿用設計書原型，廣達是模型已校準的組裝廠；不是六家的中位數）"); c = H.cell(hr, 7, f"=G{CS['廣達 2382']}"); c.number_format = "0.00"; c.fill = KEY; T_B = hr; hr += 1
H.cell(hr, 1, "t_med"); H.cell(hr, 2, "六家最近季通路停留 中位數（對照，不用於基準）"); c = H.cell(hr, 7, f"=MEDIAN({_g()})"); c.number_format = "0.00"; hr += 1
H.cell(hr, 1, "t_hi"); H.cell(hr, 2, "六家最近季通路停留 最大（τ1obs 高；Supermicro 2026Q2 曾達約 4.2 個月，見上表）"); c = H.cell(hr, 7, f"=MAX({_g()})"); c.number_format = "0.00"; c.fill = KEY; T_HI = hr; hr += 1
for nm, rr in (("t1o_lo", T_LO), ("t1o", T_B), ("t1o_hi", T_HI)):
    wb.defined_names[nm] = openpyxl.workbook.defined_name.DefinedName(nm, attr_text=f"ChanObs!$G${rr}")
for col, w in zip("ABCDEFGHIJKLMNO", (6, 40, 10, 12, 10, 12, 12, 12, 10, 12, 12, 16, 40, 18, 60)):
    H.column_dimensions[col].width = w

# ---------- BuyerObs ----------
BY["A1"] = (f"BuyerObs — 買方在建工程與現金資本支出（種子檔 channel_obs_v{chs['version']}.json；SEC XBRL 申報與 10-Q 附註，Verified；US$百萬）。"
            "在建工程含建物，Meta、Oracle、CoreWeave 未拆建物與 IT；Alphabet 為『assets not yet in service』，含建物與 IT。所以由在建工程換算的 IT 量是上界。Microsoft、Amazon 沒有逐季在建工程揭露或本輪沒找到，未列入。"
            "Oracle 曆季對應財季：2026Q2＝FQ4 FY26（5/31）、2026Q3＝FQ1 FY27（8/31）；11 月與 2 月 10-Q 未列在建工程。通路停留（月）＝在建工程 ÷（最近季現金資本支出 × 4）× 12，僅供對照")
BY["A1"].font = BOLD
for j, h_ in enumerate(["鍵", "公司", "曆季", "財報期末", "現金資本支出（US$M）", "在建工程（US$M）", "在建工程口徑", "等級", "來源ID", "在建工程 ÷ 年化現金資本支出（月）", "備註（種子檔）"], 1):
    c = BY.cell(3, j, h_); c.font = BOLD; c.fill = HDR
br_ = 4
BROWS = {}
for n, o in enumerate(chs["buyer_quarterly"], 1):
    BY.cell(br_, 1, f"b{n}"); BY.cell(br_, 2, o["company"]); BY.cell(br_, 3, o["quarter"]); BY.cell(br_, 4, o["period_end"])
    c = BY.cell(br_, 5, o["cash_capex_usd_m"]); c.font = BLUE; c.number_format = NUM
    if o.get("cip_usd_m") is not None:
        c = BY.cell(br_, 6, o["cip_usd_m"]); c.font = BLUE; c.number_format = NUM
        c = BY.cell(br_, 10, f"=F{br_}/(E{br_}*4)*12"); c.number_format = "0.0"
    BY.cell(br_, 7, o.get("cip_scope")); BY.cell(br_, 8, o["evidence_grade"]); BY.cell(br_, 9, "；".join(o["source_ids"])); BY.cell(br_, 11, o["note"][:160])
    BROWS.setdefault(o["company"], []).append(br_)
    br_ += 1
br_ += 1
BY.cell(br_, 1, "最近季摘要（2026Q3）").font = BOLD; br_ += 1
for j, h_ in enumerate(["鍵", "公司", "最近季", "在建工程（US$B）", "最近季現金資本支出（US$B）", "近四季現金資本支出合計（US$B）", "停留（月；季年化分母）", "停留（月；近四季分母）", "口徑"], 1):
    c = BY.cell(br_, j, h_); c.font = BOLD; c.fill = HDR
br_ += 1
BS = {}
for n, nm in enumerate(["Meta", "Alphabet", "Oracle", "CoreWeave"], 1):
    rows_ = BROWS[nm]; last_ = rows_[-1]
    BY.cell(br_, 1, f"bs{n}"); BY.cell(br_, 2, nm); BY.cell(br_, 3, f"=C{last_}")
    c = BY.cell(br_, 4, f"=F{last_}/1000"); c.number_format = "0.0"
    c = BY.cell(br_, 5, f"=E{last_}/1000"); c.number_format = "0.0"
    c = BY.cell(br_, 6, f"=SUM(E{rows_[0]}:E{last_})/1000"); c.number_format = "0.0"
    c = BY.cell(br_, 7, f"=D{br_}/(E{br_}*4)*12"); c.number_format = "0.0"
    c = BY.cell(br_, 8, f"=D{br_}/F{br_}*12"); c.number_format = "0.0"
    BY.cell(br_, 9, "assets not yet in service（含建物與 IT）" if nm == "Alphabet" else "在建工程（含建物，未拆）")
    BS[nm] = br_; br_ += 1
for col, w in zip("ABCDEFGHIJK", (6, 12, 10, 14, 18, 18, 36, 14, 22, 20, 60)):
    BY.column_dimensions[col].width = w

# ---------- OdmMonthly（W-SF-05；只顯示）----------
OM["A1"] = ("OdmMonthly — 四家台廠月營收追蹤（NT$百萬；種子檔月營收，二手：Yahoo 奇摩股市與 HiStock 互核，轉載公司月營收公告，標 Interested-party；鴻海 2026-10 單一來源）。"
            "只列月營收、季加總（對季報營收）、年增；AI 占比換算列為 [Assumed] 區間，只顯示。本頁不與 2027 年對帳單檢驗點的任何門檻比較、不做任何判定（檢驗點的方法與判定在 adjudication/20261118_2027年對帳單_檢驗點事前寫定.md）；不被任何既有列引用")
OM["A1"].font = BOLD
for j, h_ in enumerate(["鍵", "月份", "廣達", "鴻海", "緯創", "緯穎", "四家合計", "廣達 年增", "鴻海 年增", "緯創 年增", "緯穎 年增", "四家合計 年增"], 1):
    c = OM.cell(3, j, h_); c.font = BOLD; c.fill = HDR
mm = chs["monthly_revenue"]["months"]
orow = 4
MR = {}
for n, m_ in enumerate(mm):
    OM.cell(orow, 1, f"m{n + 1}"); OM.cell(orow, 2, m_["month"])
    for j, k in ((3, "quanta"), (4, "hon_hai"), (5, "wistron"), (6, "wiwynn")):
        c = OM.cell(orow, j, m_[k]); c.font = BLUE; c.number_format = NUM
    c = OM.cell(orow, 7, f"=SUM(C{orow}:F{orow})"); c.number_format = NUM
    if n >= 12:
        for j in range(8, 13):
            c = OM.cell(orow, j, f"={CL(j - 5)}{orow}/{CL(j - 5)}{orow - 12}-1"); c.number_format = "0.0%"
    MR[m_["month"]] = orow
    orow += 1
m_first, m_last = 4, orow - 1
orow += 1
OM.cell(orow, 1, "季加總對季報營收（NT$百萬；季報營收取自種子檔 odm_quarterly；2026Q4 只有 10 月，標『不完整』，不計年增）").font = BOLD; orow += 1
for j, h_ in enumerate(["鍵", "曆季", "廣達 月加總", "鴻海 月加總", "緯創 月加總", "緯穎 月加總", "四家月加總", "廣達 季報", "鴻海 季報", "緯創 季報", "緯穎 季報", "四家季報",
                        "廣達 差異％", "鴻海 差異％", "緯創 差異％", "緯穎 差異％", "四家月加總 年增", "備註"], 1):
    c = OM.cell(orow, j, h_); c.font = BOLD; c.fill = HDR
orow += 1
QREV = {}
for o in chs["odm_quarterly"]:
    QREV[(o["company"], o["quarter"])] = o["revenue"]
QTRS = ["2025Q1", "2025Q2", "2025Q3", "2025Q4", "2026Q1", "2026Q2", "2026Q3", "2026Q4"]
QROW = {}
TW = ["廣達 2382", "鴻海 2317", "緯創 3231", "緯穎 6669"]
for qi, q in enumerate(QTRS):
    y_, qn = int(q[:4]), int(q[5])
    ms = [f"{y_}-{(qn - 1) * 3 + k:02d}" for k in (1, 2, 3)]
    rows_m = [MR[m] for m in ms if m in MR]
    OM.cell(orow, 1, f"q{qi + 1}"); OM.cell(orow, 2, q)
    for j in range(4):
        col = CL(3 + j)
        c = OM.cell(orow, 3 + j, f"=SUM({col}{rows_m[0]}:{col}{rows_m[-1]})"); c.number_format = NUM
    c = OM.cell(orow, 7, f"=SUM(C{orow}:F{orow})"); c.number_format = NUM
    if q != "2026Q4":
        for j, f_ in enumerate(TW):
            c = OM.cell(orow, 8 + j, QREV[(f_, q)]); c.font = BLUE; c.number_format = NUM
            c = OM.cell(orow, 13 + j, f"={CL(3 + j)}{orow}/{CL(8 + j)}{orow}-1"); c.number_format = "0.0%"
        c = OM.cell(orow, 12, f"=SUM(H{orow}:K{orow})"); c.number_format = NUM
        if qi >= 4:
            c = OM.cell(orow, 17, f"=G{orow}/G{QROW[QTRS[qi - 4]]}-1"); c.number_format = "0.0%"
    else:
        OM.cell(orow, 18, "不完整：只有 2026-10 一個月")
    QROW[q] = orow
    orow += 1
orow += 1
OM.cell(orow, 1, "AI 占比換算（只顯示；AI 占比 [Assumed] 低／基準／高＝ai_sh，四家共用，公司都沒拆；不得與任何門檻比較）：四家月營收合計 × ai_sh（NT$百萬）").font = BOLD; orow += 1
for j, h_ in enumerate(["鍵", "月份", "低", "基準", "高", "註"], 1):
    c = OM.cell(orow, j, h_); c.font = BOLD; c.fill = HDR
orow += 1
for n, m_ in enumerate(mm):
    r_ = MR[m_["month"]]
    OM.cell(orow, 1, f"a{n + 1}"); OM.cell(orow, 2, m_["month"])
    for j, nm in ((3, "ai_sh_lo"), (4, "ai_sh"), (5, "ai_sh_hi")):
        c = OM.cell(orow, j, f"=G{r_}*{nm}"); c.number_format = NUM
    if n == 0:
        OM.cell(orow, 6, "[Assumed]；只顯示；鴻海含大量非 AI 營收，同一占比套用到四家會高估")
    orow += 1
for col, w in zip("ABCDEFGHIJKLMNOPQR", (6, 10, 12, 12, 12, 12, 13, 12, 12, 12, 12, 13, 11, 11, 11, 11, 14, 30)):
    OM.column_dimensions[col].width = w

# ---------- Check ----------
Z["A1"] = ("Check — 把存量帳的兩段要放的量，對照可觀測的量：(i) 通路容量、(ii) 金額、(iii) 改放買方端。每個核對輸出『放得下／放不下多少 GW』（低／基準／高），"
           "【不做判定、不設任何通過門檻】；本頁不被任何既有列引用，也不進任何基準或殘差帳（不得用來縮小 Residual 頁的未歸因殘差）。"
           "列的『低／基準／高』指結果的低／基準／高（放不下的量由小到大），不是各輸入同步取低值。正值＝放不下，負值＝放得下。單位：實體 GW，金額 US$B")
Z["A1"].font = BOLD
for j, h_ in enumerate(["列", "項目", "說明", "低", "基準", "高", "等級"], 1):
    c = Z.cell(3, j, h_); c.font = BOLD; c.fill = HDR
zr = 4
ZR = {}
S_ = lambda key, yr=2027: f"Stock!{C(yr)}{KR[key]}"


def zrow(key, label, note, d=None, e=None, f=None, tag="Derived", fmt="0.00", fill=None, bold=False):
    global zr
    Z.cell(zr, 1, key); Z.cell(zr, 2, label); Z.cell(zr, 3, note)
    for j, v in zip((4, 5, 6), (d, e, f)):
        if v is not None:
            c = Z.cell(zr, j, v); c.number_format = fmt
            if fill: c.fill = fill
    Z.cell(zr, 7, tag)
    if bold:
        Z.cell(zr, 2).font = BOLD
    ZR[key] = zr; zr += 1


def zsec(t):
    global zr
    Z.cell(zr, 1, t).font = BOLD; zr += 1


def ztxt(key, label, src_key, note=""):
    r0 = ZR[src_key]
    zrow(key, label, note, *[f'=IF({c}{r0}>0,"放不下 "&TEXT({c}{r0},"0.0")&" GW","放得下，餘 "&TEXT(-{c}{r0},"0.0")&" GW")' for c in "DEF"], tag="Derived", fmt="@")


zsec("A. 2027 年底各段要放的量（Stock 頁 D、E、G 節；低／基準／高＝stk_2025 低／基準／高起點）")
zrow("a1", "S1 通路需要放的量", "Stock d_end；起點 stk_2025 × s1sh", f"={S_('d_end_lo')}", f"={S_('d_end_b')}", f"={S_('d_end_hi')}", "Derived（stk_2025 Assumed）", fill=KEY, bold=True)
zrow("a2", "S1 通路需要放的量（起點自洽版，含在途量）", "Stock g_s1a；低／基準／高＝τ1obs 低／基準／高（對照，不是基準）", f"={S_('g_s1a_lo')}", f"={S_('g_s1a_b')}", f"={S_('g_s1a_hi')}", "Derived（τ1obs Interested-party／Verified）")
zrow("a3", "S2 買方端需要放的量（需求側自洽）", "Stock e_end_b", None, f"={S_('e_end_b')}", None, "Derived")
zr += 1
zsec("B. (i) 通路容量核對——停留時間法：通路能放的量＝觀測停留時間 ÷ 12 × 2027 年出貨 ①（假設存貨隨出貨同比例放大）")
zrow("b1", "通路停留 τ1obs（月）", "低＝六家最近季最小；基準＝廣達 3Q26（沿用設計書原型）；高＝六家最大。全公司口徑含非 AI 與原料，偏寬（上界）", "=t1o_lo", "=t1o", "=t1o_hi", "Derived（存貨 Interested-party／Verified）", fmt="0.00")
zrow("b2", "通路容量（GW）", "τ1obs ÷ 12 × 2027 ①（phys_2027）", f"=D{ZR['b1']}/12*phys_2027", f"=E{ZR['b1']}/12*phys_2027", f"=F{ZR['b1']}/12*phys_2027", "Derived", bold=True)
zrow("b3", "放不下的量（只動容量；需要放＝a1 基準）", "低＝容量取高端；高＝容量取低端", f"=$E${ZR['a1']}-F{ZR['b2']}", f"=$E${ZR['a1']}-E{ZR['b2']}", f"=$E${ZR['a1']}-D{ZR['b2']}", "Derived", fill=KEY, bold=True)
ztxt("b3t", "　　文字", "b3")
zrow("b4", "放不下的量（容量與起點同時動；最寬範圍）", "低＝需要放取低且容量取高；高＝需要放取高且容量取低", f"=D{ZR['a1']}-F{ZR['b2']}", f"=E{ZR['a1']}-E{ZR['b2']}", f"=F{ZR['a1']}-D{ZR['b2']}", "Derived")
zrow("b5", "放不下的量（起點自洽版 a2；對照）", "補上在途量後需要放更多", f"=$E${ZR['a2']}-F{ZR['b2']}", f"=$E${ZR['a2']}-E{ZR['b2']}", f"=$E${ZR['a2']}-D{ZR['b2']}", "Derived")
zr += 1
zsec("C. (i) 通路容量核對——存貨水準法：各家最近季存貨 × AI 占比（[Assumed]）÷ 每 GW ODM 整櫃價值＝AI 存貨相當幾 GW；再按 2027 年出貨對 2026 年出貨放大")
CF = {}
for f_ in FIRMS:
    r_ = CS[f_]
    nt = f_.split()[0] in ("廣達", "鴻海", "緯創", "緯穎")
    d_ = f"=ChanObs!$E${r_}/1000/fxtw_hi*ai_sh_lo/odm_26" if nt else f"=ChanObs!$E${r_}/1000*ai_sh_lo/odm_26"
    e_ = f"=ChanObs!$E${r_}/1000/fxtw*ai_sh/odm_26" if nt else f"=ChanObs!$E${r_}/1000*ai_sh/odm_26"
    h_ = f"=ChanObs!$E${r_}/1000/fxtw_lo*ai_sh_hi/odm_26" if nt else f"=ChanObs!$E${r_}/1000*ai_sh_hi/odm_26"
    zrow("c_" + f_.split()[0], f"{f_} 最近季存貨換算 AI 存貨（GW）", "存貨 [Interested-party／Verified]；AI 占比與匯率取低／基準／高端", d_, e_, h_, "Derived（AI 占比 Assumed）", fmt="0.000")
    CF[f_] = ZR["c_" + f_.split()[0]]
tw0, tw1 = CF[FIRMS[0]], CF[FIRMS[3]]
zrow("c_tw4", "四家台廠合計（GW）", "鴻海存貨以全公司計、AI 占比同一組，會高估", *[f"=SUM({c}{tw0}:{c}{tw1})" for c in "DEF"], "Derived", fmt="0.000")
zrow("c_us2", "Dell＋Supermicro（GW；不併入下面的全市場外推，兩者是否已在四家台廠份額之外未驗證）", "Dell 最近季為 2026Q2", *[f"={c}{CF['Dell']}+{c}{CF['Supermicro']}" for c in "DEF"], "Derived", fmt="0.000")
zrow("c_ext", "全市場 AI 存貨（GW）＝四家台廠合計 ÷ 四家有效份額合計 2026", "SIF_AsmShareSum 2026；份額分母含其他組裝廠", *[f"={c}{ZR['c_tw4']}/asmsum_26" for c in "DEF"], "Derived", bold=True)
zrow("c_extusd", "　全市場 AI 存貨（US$B）", "× 2026 每 GW ODM 整櫃價值", *[f"={c}{ZR['c_ext']}*odm_26" for c in "DEF"], "Derived", fmt="0")
zrow("c_needusd", "　S1 通路需要放的存貨（US$B，a1 × 2027 每 GW ODM 價值）", "", *[f"={c}{ZR['a1']}*odm_27" for c in "DEF"], "Derived", fmt="0")
zrow("c_mult", "　需要放 ÷ 3Q26 全市場 AI 存貨（倍）", "低＝需要取低且觀測取高", f"=D{ZR['c_needusd']}/F{ZR['c_extusd']}", f"=E{ZR['c_needusd']}/E{ZR['c_extusd']}", f"=F{ZR['c_needusd']}/D{ZR['c_extusd']}", "Derived", fmt="0.0")
zrow("c_cap", "通路容量（GW）＝全市場 AI 存貨 × 2027 ① ÷ 2026 ①", "存貨隨年出貨同比例放大（與 B 節同一假設）", *[f"={c}{ZR['c_ext']}*phys_2027/phys_2026" for c in "DEF"], "Derived", bold=True)
zrow("c_exc", "放不下的量（需要放＝a1 基準）", "低＝容量取高端；高＝容量取低端", f"=$E${ZR['a1']}-F{ZR['c_cap']}", f"=$E${ZR['a1']}-E{ZR['c_cap']}", f"=$E${ZR['a1']}-D{ZR['c_cap']}", "Derived", fill=KEY, bold=True)
ztxt("c_exct", "　　文字", "c_exc")
q_ = CS["廣達 2382"]
zrow("c_q", "原型對照（只廣達）：3Q26 AI 存貨（GW）＝存貨 ÷ 匯率 × AI 占比 ÷ 廣達份額 ÷ 每 GW 價值", "設計書原型約 5.0 GW；基準欄", None, f"=ChanObs!$E${q_}/1000/fxtw*ai_sh/qsh_26/odm_26", None, "Derived", fmt="0.00")
zrow("c_q27", "　同上放大到 2027（GW）", "設計書原型約 8.8 GW", None, f"=E{ZR['c_q']}*phys_2027/phys_2026", None, "Derived", fmt="0.00")
zr += 1
zsec("D. (ii) 金額核對——兩側隱含的 2027 年 NVIDIA 資料中心營收差額，換算成 GW（只 NVIDIA、只算一年的流量差）")
zrow("d1", "Siliconomics 隱含 NVIDIA 資料中心營收 CY2027（US$B）", "層級值（含指引）；與 2027 年對帳單凍結紀錄同值", None, "=nv_sil27", None, "Derived", fmt="0.0")
zrow("d2", "Wattonomics 隱含 NVIDIA 資料中心營收 CY2027（US$B）", "", None, "=nv_wt27", None, "Derived", fmt="0.0")
zrow("d3", "兩側差額（US$B）", "d1 − d2", None, f"=E{ZR['d1']}-E{ZR['d2']}", None, "Derived", fmt="0.0")
zrow("d4", "差額換算 GW（金額面放不下的量）", "÷ 2027 NVIDIA 每 GW 內容；低＝內容取高端（VR200_hi）；高＝內容取低端（VR200_lo）", f"=E{ZR['d3']}/nvcpg27_hi", f"=E{ZR['d3']}/nvcpg27", f"=E{ZR['d3']}/nvcpg27_lo", "Derived（內容區間 Assumed）", fill=KEY, bold=True)
ztxt("d4t", "　　文字", "d4", "＝若 Wattonomics 的金額是對的，供給側的 2027 年出貨在金額上有這麼多 GW 沒有對應的買方錢")
zrow("d5", "　對照：通路放不下的量（B 節 b3）減金額面放不下的量（d4）", "兩者量級對照，只顯示；金額面是一年的流量差，通路是年底存量，不是同一個量", f"=D{ZR['b3']}-D{ZR['d4']}", f"=E{ZR['b3']}-E{ZR['d4']}", f"=F{ZR['b3']}-F{ZR['d4']}", "Derived")
zr += 1
zsec("E. (iii) 改放買方端的試算——若把通路放不下的量全部移到買方端（已收貨、未通電），對照買方在建工程")
zrow("e1", "買方端 2027 年底要放的量（情境 Y）＝S2（需求側）＋通路放不下的量（b3，負值取 0）", "低／基準／高＝b3 低／基準／高", *[f"={S_('e_end_b')}+MAX(0,{c}{ZR['b3']})" for c in "DEF"], "Derived", bold=True)
for nm in ["Meta", "Alphabet", "Oracle", "CoreWeave"]:
    zrow("e_" + nm, f"{nm} 最近季在建工程（US$B）", "含建物，上界；" + ("assets not yet in service（含 IT）" if nm == "Alphabet" else "未拆建物與 IT"), None, f"=BuyerObs!$D${BS[nm]}", None, "Verified", fmt="0.0")
zrow("e_cip", "四家在建工程合計（US$B）", "低＝只含標為在建工程的三家（Meta、Oracle、CoreWeave，不含 Alphabet 的『未投入使用』）；基準與高＝四家。缺 Microsoft、Amazon 與其他買方",
     f"=E{ZR['e_Meta']}+E{ZR['e_Oracle']}+E{ZR['e_CoreWeave']}", f"=SUM(E{ZR['e_Meta']}:E{ZR['e_CoreWeave']})", f"=SUM(E{ZR['e_Meta']}:E{ZR['e_CoreWeave']})", "Verified", fmt="0.0")
zrow("e_cipgw", "在建工程換算 GW 上限（若全是 IT）", "÷ 每 GW IT 資本支出；低與基準取 2027（37.6）、高取 2026（33.4）；含建物故為上界，但只含四家故也可能偏小——兩個偏誤方向相反，未量化",
     f"=D{ZR['e_cip']}/itpg_2027", f"=E{ZR['e_cip']}/itpg_2027", f"=F{ZR['e_cip']}/itpg_2026", "Derived", bold=True)
zrow("e_x1", "放不下的量（在建工程金額法）", "e1 − 在建工程換算 GW；低＝需要取低且容量取高", f"=D{ZR['e1']}-F{ZR['e_cipgw']}", f"=E{ZR['e1']}-E{ZR['e_cipgw']}", f"=F{ZR['e1']}-D{ZR['e_cipgw']}", "Derived", fill=KEY)
ztxt("e_x1t", "　　文字", "e_x1")
for nm in ["Meta", "Alphabet", "Oracle", "CoreWeave"]:
    zrow("e_t_" + nm, f"{nm} 在建工程停留（月；季年化／近四季）", "基準欄＝在建工程 ÷（最近季現金資本支出 × 4）；高欄＝在建工程 ÷ 近四季現金資本支出合計（BuyerObs）；含建物，上界。本列的欄位不是低／基準／高", None, f"=BuyerObs!$G${BS[nm]}", f"=BuyerObs!$H${BS[nm]}", "Derived", fmt="0.0")
_ta = ",".join(f"E{ZR['e_t_' + n_]}:F{ZR['e_t_' + n_]}" for n_ in ["Meta", "Alphabet", "Oracle", "CoreWeave"])
_tq = ",".join(f"E{ZR['e_t_' + n_]}" for n_ in ["Meta", "Alphabet", "Oracle", "CoreWeave"])
zrow("e_tau", "買方端停留 τ2obs（月）", "低＝八個值最小；基準＝四家季年化值的中位數；高＝八個值最大", f"=MIN({_ta})", f"=MEDIAN({_tq})", f"=MAX({_ta})", "Derived", fmt="0.0")
zrow("e_cap", "買方端容量（GW）＝τ2obs ÷ 12 × 2027 ③（Stock e_on）", "", *[f"={c}{ZR['e_tau']}/12*{S_('e_on')}" for c in "DEF"], "Derived", bold=True)
zrow("e_x2", "放不下的量（停留時間法）", "e1 − 容量；低＝需要取低且容量取高", f"=D{ZR['e1']}-F{ZR['e_cap']}", f"=E{ZR['e1']}-E{ZR['e_cap']}", f"=F{ZR['e1']}-D{ZR['e_cap']}", "Derived", fill=KEY)
ztxt("e_x2t", "　　文字", "e_x2")
zrow("e_tauy", "　對照：情境 Y 隱含的買方端停留（月）＝e1 ÷ 2027 ③ × 12", "", *[f"={c}{ZR['e1']}/{S_('e_on')}*12" for c in "DEF"], "Derived", fmt="0.0")
zr += 1
zsec("說明（不是判定）")
for t in ["本頁所有『放不下』只是算術差；不設通過門檻、不做結論，也不改任何基準；不得用來縮小 Residual 頁的未歸因殘差（規則：沒有可檢驗實績不得歸因）。",
          "通路觀測是全公司口徑（含非 AI、含尚未上板的原料），AI 與非 AI 公司都沒拆，AI 占比 [Assumed]（ai_sh）；停留時間法因此偏寬（容量偏大），水準法的鴻海因同一占比套用而可能高估。兩個偏誤都往『容量偏大』的方向。",
          "買方在建工程含建物（Alphabet 為含 IT 的未投入使用資產），換算 GW 是上界；但只含四家買方，缺 Microsoft、Amazon 與主權、新雲，實際可能更大。方向相反的兩個偏誤，未量化。",
          "2028–29 的 ② 買方收貨無 Wattonomics 名稱，取 Bridge s2 鏡像；本頁只用 2027 年底值。",
          "未使用 NVIDIA FY2027 第三季實績；本頁沒有任何數值取自 2026-08 之後才公布的 NVIDIA 資料。"]:
    Z.cell(zr, 2, t); zr += 1
for col, w in zip("ABCDEFG", (9, 66, 70, 12, 12, 12, 28)):
    Z.column_dimensions[col].width = w

out = sys.argv[1]
wb.save(out)
print("wrote", out)
