# Siliconomics 供給側證據比例（三層基準協調規則 v1.1 C3＋歸類補充；取代 In 頁 crd_es_26–29 的手填值）
# 用法：python3 siliconomics/tools/supply_evidence_ratio.py <輸出 xlsx>   之後以 LibreOffice headless 重算。
# 本工具不改任何模型：分母讀 Siliconomics 現行活頁簿（model/CURRENT）的原始顆數與每顆 IT capex；分子是本檔 In 頁的證據（藍字）。
# 每一步都是 Excel 公式；藍字＝輸入、黑字＝公式。三組參數：低／基準／高（「低」＝證據撐得最少的取法）。
import os, sys, datetime
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.workbook.defined_name import DefinedName

ROOT = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
sil_f = open(os.path.join(ROOT, "siliconomics/model/CURRENT")).read().strip()
S = openpyxl.load_workbook(os.path.join(ROOT, "siliconomics/model", sil_f), data_only=True)


def sv(name):
    sh, ref = list(S.defined_names[name].destinations)[0]
    return S[sh][ref.replace("$", "").split(":")[0]].value


def s_in(key):
    """Siliconomics In 頁鍵 → (低, 基準, 高, 列號)。"""
    ws = S["In"]
    for r in range(5, ws.max_row + 1):
        if ws.cell(r, 1).value == key:
            return ws.cell(r, 4).value, ws.cell(r, 5).value, ws.cell(r, 6).value, r
    raise KeyError(key)


YRS = [2026, 2027, 2028, 2029]
BOLD, BLUE = Font(bold=True), Font(color="0000FF")
HDR = PatternFill("solid", fgColor="DDEBF7")
KEY = PatternFill("solid", fgColor="FFF2CC")
WRAP = Alignment(wrap_text=True, vertical="top")
wb = openpyxl.Workbook()

# =============================== README ===============================
RD = wb.active; RD.title = "README"

# =============================== In ===============================
IN = wb.create_sheet("In")
IN["A1"] = (f"In — 輸入（藍字＝輸入；低＝證據撐得最少的取法，高＝最多）。分母輸入鏡像自 siliconomics/model/{sil_f}（標記沿用原格）；"
            f"證據輸入來自 /silev/supply_evidence.md 與原文核對（2026-11-14）。產生日 {datetime.date.today()}")
IN["A1"].font = BOLD
for j, h in enumerate(["鍵", "項目", "單位", "低", "基準", "高", "標記", "來源"], 1):
    c = IN.cell(3, j, h); c.font = BOLD; c.fill = HDR
N = {}
r = 4


def sec(title):
    global r
    r += 1
    IN.cell(r, 1, title).font = BOLD
    r += 1


def inp(key, label, unit, lo, base, hi, tag, src):
    global r
    if lo is None:
        lo = base
    if hi is None:
        hi = base
    for j, v in enumerate([key, label, unit, lo, base, hi, tag, src], 1):
        c = IN.cell(r, j, v)
        if j in (4, 5, 6):
            c.font = BLUE
    IN.cell(r, 8).alignment = WRAP
    N[key] = {"lo": f"In!$D${r}", "b": f"In!$E${r}", "hi": f"In!$F${r}"}
    wb.defined_names[f"ev_{key}"] = DefinedName(f"ev_{key}", attr_text=f"In!$E${r}")
    r += 1


# ---- A. 分母（Siliconomics 原始顆數 × 每顆 IT capex ÷ 每 GW IT capex；不含 S33 係数）
sec(f"A. 分母鏡像：Siliconomics {sil_f}（原始顆數＝Units 頁第 4–11 列，未乘 crd_k／S33；每顆 IT capex＝In 頁 M9；每 GW IT capex＝Deploy 頁第 6 列）")
UROW = {"bw": 4, "rb": 5, "ru": 6, "hp": 7, "amd": 8, "tpu": 9, "trn": 10, "oth": 11}
ULAB = {"bw": "NVIDIA Blackwell", "rb": "NVIDIA Rubin", "ru": "NVIDIA Rubin Ultra", "hp": "NVIDIA Hopper", "amd": "AMD Instinct",
        "tpu": "Google TPU", "trn": "AWS Trainium", "oth": "其他 ASIC（MTIA、Maia、昇騰、OpenAI 等）"}
for k, row in UROW.items():
    for j, y in enumerate(YRS):
        v = S["Units"].cell(row, 3 + j).value
        tag = "Interested-party／Assumed（沿用 Siliconomics）" if y < 2029 else "Derived（2028 × 平台成長率）"
        inp(f"u_{k}_{y % 100}", f"{ULAB[k]} 原始顆數 {y}", "百萬顆", None, v, None, tag, f"{sil_f} Units!{'CDEF'[j]}{row}")
for k in ["it_GB", "it_VR", "it_RU", "it_HP", "it_AMD8", "it_AMDR", "it_TPU", "it_TRN", "it_ASIC"]:
    lo, b, hi, rr = s_in(k)
    inp(k, f"每顆 IT capex：{S['In'].cell(rr, 2).value}", "$K/顆", None, b, None, S["In"].cell(rr, 7).value, f"{sil_f} In!E{rr}（基準值；分母只用基準）")
for j, y in enumerate(YRS):
    inp(f"cgw_{y % 100}", f"每 GW IT capex {y}（成本等值 GW 的換算基礎）", "$B/GW", None, S["Deploy"].cell(6, 3 + j).value, None, "Derived",
        f"{sil_f} Deploy!{'CDEF'[j]}6（2026＝GB300 與 VR200 依 vr26 混合；2027 起 VR200；Tokenomics IF_CapexIT）")
for j, y in enumerate(YRS):
    inp(f"sup_{y % 100}", f"對照：Siliconomics 供給側讀數 {y}（SIF_DeployGW_Supply）", "GW", None, sv(f"SIF_DeployGW_Supply_{y}"), None, "Derived",
        f"{sil_f} Interface SIF_DeployGW_Supply_{y}（本檔分母合計應與此相同）")

# ---- B. 換算參數
sec("B. 換算參數（營收／承諾金額 → 成本等值 GW）")
for k in ["nvr_GB300", "nvr_VR26", "nvr_VR27"]:
    lo, b, hi, rr = s_in(k)
    # 低證據＝每 GW 營收較高（同樣金額換到較少 GW），故「低」取原區間的高值
    inp(k, f"NVIDIA 營收 ÷ IT capex（{k[4:]}）；低欄＝原區間高值（每 GW 營收高 → 撐得少）", "%", hi, b, lo, "Derived",
        f"{sil_f} In!D{rr}:F{rr}（晶片內容橋接表 v0.2.3 營收口徑）；低／高對調")
lo, b, hi, rr = s_in("vr26")
inp("vr26", "2026 部署中 Rubin 世代占比", "%", None, b, None, "Interested-party", f"{sil_f} In!E{rr}")
inp("hold_amd", "AMD MI455X 持有成本比（每 GW 持有成本 ÷ VR200）", "x", None, sv("TK_NonNV_Hold_AMD"), None, "Assumed",
    f"{sil_f} TK_In（Tokenomics IF_NonNV_Hold_AMD）；實體 GW → 成本等值 GW＝× 此比")

# ---- C. NVIDIA 證據
sec("C. NVIDIA：已公布實績（等級 1）與 10-Q／CFO Commentary 的存貨、供應與產能採購承諾（等級 2，歸類補充 a）")
inp("nv_dc_fy26q4", "NVIDIA 資料中心營收 FY26 Q4（2025-11～2026-01）", "$B", None, 62.3, None, "Verified",
    "NVIDIA 8-K（siliconomics/tools/r4_bridge.py nv_dc_25 註：FY26 Q1–Q4 39.1／41.1／51.2／62.3）")
inp("jan_share_nv", "FY26 Q4 中落在日曆 2026 的比例（1 月＝3 個月中 1 個月）", "x", None, 1 / 3, None, "Derived", "會計年度對日曆年：按月數平均分攤")
inp("nv_dc_q1", "NVIDIA 資料中心營收 FY27 Q1（2026-02～04）", "$B", None, 75.246, None, "Verified",
    "NVIDIA Q2FY27 CFO Commentary 營收表（Data Center $75,246M）")
inp("nv_dc_q2", "NVIDIA 資料中心營收 FY27 Q2（2026-05～07）", "$B", None, 89.023, None, "Verified",
    "NVIDIA Q2FY27 CFO Commentary 營收表（Data Center $89,023M）；FY27 Q3（2026-08～10）11-19 公布，本檔未計")
inp("nv_rev_q2", "NVIDIA 總營收 FY27 Q2", "$B", None, 96.0, None, "Verified", "NVIDIA FY27 Q2 官方逐字稿（2026-08-26；supply_evidence D2）")
inp("nv_inv", "NVIDIA 存貨 2026-07-26（Q2 季底）", "$B", None, 31.6, None, "Verified",
    "Q2FY27 CFO Commentary：「Inventory was $31.6 billion, up from $25.8 billion sequentially」；季底未出貨，全數計入 2026 下半年")
for k, v, per in [("c_fy27r", 92, "FY27 剩餘（2026-08～2027-01，6 個月）"), ("c_fy28", 87, "FY28（2027-02～2028-01）"),
                  ("c_fy29", 88, "FY29（2028-02～2029-01）"), ("c_fy30", 6, "FY30（2029-02～2030-01）"), ("c_fy31", 5, "FY31（2030-02～2031-01）")]:
    inp(k, f"供應與產能承諾 {per}", "$B", None, v, None, "Verified",
        "NVIDIA Q2FY27 CFO Commentary 承諾表「Supply and capacity $92 $87 $88 $6 $5 $1 $279」；10-Q（SEC nvda-20260726）；"
        "增加主因為記憶體採購。另有雲端服務協議 290 億（買方性質，不計）")
inp("nv_gm", "NVIDIA 毛利率（承諾的銷貨成本 → 營收）；低＝毛利率低（同樣成本換到較少營收）", "%", 0.70, 0.75, 0.76, "Verified／Assumed",
    "FY27 Q1 74.9%、Q2 75.0%（GAAP）、Q3 指引 74.0%±0.5（CFO Commentary）；低 70%＝2027 記憶體漲價由 NVIDIA 吸收的假設（Assumed）")
inp("nv_dcsh", "承諾中屬資料中心產品的比例", "x", 0.85, None, 1.0, "Derived／Assumed",
    "基準＝FY27 Q2 資料中心 ÷ 總營收（89.023 ÷ 96.0）；低 0.85＝另扣部分長期產能預付（Assumed）；高＝全數")
IN.cell(r - 1, 5).value = "=E{0}/E{1}".format(int(N["nv_dc_q2"]["b"].split("$")[-1]), int(N["nv_rev_q2"]["b"].split("$")[-1]))
inp("nv_cov", "承諾須涵蓋的銷貨成本比例（1＝把每一元承諾當成成品晶片的一元成本；<1＝承諾只買記憶體）", "x", 1.0, 1.0, 0.65, "Assumed／Derived",
    "判斷類（待判斷 1）：基準＝1（承諾換算視同完整銷貨成本，保守）；高 0.65＝承諾全為記憶體時，記憶體占每 GW NVIDIA 銷貨成本的比例"
    "（晶片內容橋接表 v0.2.3：VR200 2027 HBM 3.025＋LPDDR5X 1.475 ＝4.50 $B/GW ÷（26.127 × 0.25）＝0.69，取 0.65）")
inp("lag", "採購到出貨的時滯（月；承諾月份＋時滯＝撐住的出貨月份）", "月", 3, 3, 3, "Assumed",
    "判斷類（待判斷 2）：存貨 31.6 ÷ 季銷貨成本約 22（96 × 25%）≈ 1.4 季；基準 3 個月；敏感度 0 與 6 個月見 Calc 頁第 5 節")
inp("lag_s0", "敏感度：時滯 0 個月", "月", None, 0, None, "Assumed", "只作敏感度")
inp("lag_s6", "敏感度：時滯 6 個月", "月", None, 6, None, "Assumed", "只作敏感度")

# ---- D. Broadcom（Google TPU、MTIA、OpenAI 等 XPU）
sec("D. Broadcom AI 半導體實績（等級 1；撐 Google TPU＋其他 ASIC 分母）")
inp("av_q1", "Broadcom AI 營收 FY26 Q1（2025-11-03～2026-02-01）", "$B", None, 8.4, None, "Verified",
    "Broadcom 8-K 新聞稿（SEC avgo-02012026x8kxex99）：「Q1 AI revenue of $8.4 billion」")
inp("jan_share_av", "FY26 Q1 中落在日曆 2026 的比例（1 月）", "x", None, 1 / 3, None, "Derived", "按月數平均分攤")
inp("av_q2", "Broadcom AI 營收 FY26 Q2（2026-02-02～05-03）", "$B", None, 10.8, None, "Verified",
    "Broadcom 8-K 新聞稿（SEC avgo-05032026x8kxex99）：「Q2 semiconductor revenue from AI of $10.8 billion」")
inp("av_q3", "Broadcom AI 營收 FY26 Q3（2026-05-04～08-02）", "$B", None, 16.7, None, "Verified",
    "Broadcom 新聞稿 2026-09-02（supply_evidence C4）；FY26 Q4（08-03～11-01）12 月公布，本檔未計")
inp("av_xpu", "AI 營收中 XPU（客製加速器）占比", "%", 0.65, 0.70, 0.73, "Interested-party／Assumed",
    "FY26 Q3 法說：XPU 占 AI 營收 73%（supply_evidence C5，B 級）；前兩季 XPU 占比較低，基準 70%、低 65%（Assumed）")
lo, b, hi, rr = s_in("core_tpu")
inp("core_tpu", "Broadcom 每顆 XPU 核心售價（不含 HBM）；低欄＝原區間高值", "$K/顆", hi, b, lo, "Assumed", f"{sil_f} In!D{rr}:F{rr}（core_tpu）；低／高對調")
lo, b, hi, rr = s_in("f_hbm")
inp("f_hbm", "HBM 由 Broadcom 代購並計入營收的比例；低欄＝原區間高值", "x", hi, b, lo, "Assumed", f"{sil_f} In!D{rr}:F{rr}（f_hbm）；低／高對調")
hp = S["Mkt"].cell(22, 4).value  # 2026 HBM 單價
gb_tpu = S["Mkt"].cell(10, 4).value / S["Units"].cell(9, 3).value  # PB ÷ 百萬顆 ＝ GB
inp("hbm_tpu", "每顆 TPU 的 HBM 金額 2026", "$K/顆", None, round(gb_tpu * hp / 1000, 4), None, "Derived",
    f"{sil_f} Mkt!D10 ÷ Units!C9＝{gb_tpu:.0f} GB × Mkt!D22 HBM 單價 {hp} $/GB")

# ---- E. AMD
sec("E. AMD（等級 1 實績；等級 2：與 OpenAI 6 GW 協議的首 1 GW，AMD 新聞稿與 8-K）")
inp("amd_q1", "AMD Data Center 部門營收 2026 Q1", "$B", None, 5.775, None, "Verified", "AMD Q2 2026 8-K 新聞稿（SEC q22026991）Selected Corporate Data：Q1 2026 $5,775M")
inp("amd_q2", "AMD Data Center 部門營收 2026 Q2", "$B", None, 6.718, None, "Verified", "同上：Q2 2026 $6,718M")
inp("amd_q3", "AMD Data Center 部門營收 2026 Q3", "$B", None, 8.546, None, "Verified", "AMD Q3 2026 8-K 新聞稿（SEC q32026991；supply_evidence C1）")
inp("amd_gpu", "Data Center 營收中 Instinct GPU 占比", "%", 0.30, 0.40, 0.50, "Assumed",
    "判斷類（待判斷 4）：AMD 未拆分；2025 年約四成（EPYC 約六成）；Q3 MI455X 開始量產，取 30–50%")
lo, b, hi, rr = s_in("pa_amd")
inp("pa_amd", "AMD Instinct 每顆售價；低欄＝原區間高值", "$K/顆", hi, b, lo, "Interested-party", f"{sil_f} In!D{rr}:F{rr}（pa_amd）；低／高對調")
inp("amd_oai27", "OpenAI 首 1 GW（MI450，2026 下半年開始）中落在 2027 年出貨的 GW", "GW", 0.5, 0.7, 1.0, "Assumed",
    "AMD IR 新聞稿 2025-10-06（supply_evidence C3，A 級，具法律約束的認股權證協議）；首 1 GW 的逐年分配未揭露（判斷類，待判斷 4）；"
    "2026 年部分已含在 Q3 實績內，為免重複只把剩餘部分放 2027；其餘 5 GW 無逐年分配，不計")
inp("amd_oai_it", "OpenAI 協議「GW」的口徑換成 IT GW 的係數（1＝IT 口徑；0.8＝設施口徑 ÷ PUE 1.25）", "x", 0.8, 1.0, 1.0, "Assumed", "新聞稿未說明口徑")

# ---- F. HBM 長約交叉檢核（不加總）
sec("F. HBM 長約交叉檢核（上游已簽約；與 NVIDIA 記憶體承諾是同一批 HBM，只作佐證、不加入分子）")
inp("skh_lta", "SK 海力士已簽 LTA 占 2027、2028 總產能比例（DRAM＋HBM 合計）", "%", 0.5, 0.5, 0.6, "Verified／Assumed",
    "SK 海力士 2026-10-27 新聞稿：「more than 50% of total production capacity for 2027 and 2028」（A 級，歸類補充 b 算等級 2）；"
    "假設 HBM 的已簽比例與總產能相同（Assumed）")
for y in (2027, 2028):
    lo, b, hi, rr = s_in(f"hbmsh_skh_{y % 100}")
    inp(f"skh_sh_{y % 100}", f"SK 海力士 HBM 份額 {y}", "%", None, b, None, "Interested-party／Assumed", f"{sil_f} In!E{rr}")
inp("mu_vm", "美光「vast majority of CY2027 HBM bit supply」已議約的數值化（只入高欄）", "%", 0, 0, 0.85, "Assumed",
    "美光 Q4 FY26 Prepared Remarks（IR 書面，非新聞稿或法定文件；待判斷 5）；基準不計，高欄取 85%")
lo, b, hi, rr = s_in("hbmsh_mu_27")
inp("mu_sh_27", "美光 HBM 份額 2027", "%", None, b, None, "Interested-party", f"{sil_f} In!E{rr}")
for j, y in enumerate((2027, 2028)):
    inp(f"hbm_sup_{y % 100}", f"HBM 位元供給 {y}", "EB", None, S["Mkt"].cell(17, 5 + j).value, None, "Derived", f"{sil_f} Mkt!{'EF'[j]}17")
    inp(f"hbm_dem_{y % 100}", f"HBM 需求（原始顆數 × 出貨係數）{y}", "EB", None, S["Mkt"].cell(14, 5 + j).value, None, "Derived", f"{sil_f} Mkt!{'EF'[j]}14")

# ---- G. Wattonomics 證據比例與門檻（狀態判定用）
sec("G. 狀態判定用（Siliconomics 試作 In 頁 M30 的鏡像）")
for y in YRS:
    lo, b, hi, rr = s_in(f"crd_ew_{y % 100}")
    inp(f"ew_{y % 100}", f"Wattonomics 證據比例 {y}（暫值）", "%", None, b, None, "Derived", f"{sil_f} In!E{rr}（crd_ew_{y % 100}）")
for k in ["crd_thr", "crd_enter"]:
    lo, b, hi, rr = s_in(k)
    inp(k, S["In"].cell(rr, 2).value, "%", lo, b, hi, "Assumed", f"{sil_f} In!D{rr}:F{rr}")
for y in YRS:
    lo, b, hi, rr = s_in(f"crd_es_{y % 100}")
    inp(f"old_es_{y % 100}", f"對照：試作原手填 Siliconomics 證據比例 {y}", "%", lo, b, hi, "Assumed", f"{sil_f} In!D{rr}:F{rr}（本檔要取代的值）")

IN.column_dimensions["A"].width = 16; IN.column_dimensions["B"].width = 62; IN.column_dimensions["H"].width = 110
for c in "DEF":
    IN.column_dimensions[c].width = 10

# =============================== NV_Month ===============================
NM = wb.create_sheet("NV_Month")
NM["A1"] = "NV_Month — NVIDIA 供應與產能承諾按月攤開（各會計年度金額 ÷ 該期月數；月序號＝年 × 12＋月）。Calc 頁以 SUMIFS 取「出貨年的月份 − 時滯」區間加總"
NM["A1"].font = BOLD
for j, h in enumerate(["月份", "月序號", "會計期", "該期承諾 $B", "該期月數", "當月承諾 $B"], 1):
    c = NM.cell(3, j, h); c.font = BOLD; c.fill = HDR
BUCK = [("FY27 剩餘", "c_fy27r", (2026, 8), 6), ("FY28", "c_fy28", (2027, 2), 12), ("FY29", "c_fy29", (2028, 2), 12),
        ("FY30", "c_fy30", (2029, 2), 12), ("FY31", "c_fy31", (2030, 2), 12)]
rr = 4
for lab, key, (y0, m0), n in BUCK:
    for i in range(n):
        y, m = y0 + (m0 - 1 + i) // 12, (m0 - 1 + i) % 12 + 1
        NM.cell(rr, 1, f"{y}-{m:02d}"); NM.cell(rr, 2, y * 12 + m); NM.cell(rr, 3, lab)
        NM.cell(rr, 4, f"={N[key]['b']}"); NM.cell(rr, 5, n); NM.cell(rr, 6, f"=D{rr}/E{rr}")
        rr += 1
NM_LAST = rr - 1
NM.cell(rr + 1, 1, "合計"); NM.cell(rr + 1, 6, f"=SUM(F4:F{NM_LAST})")
NM.column_dimensions["A"].width = 10

# =============================== Calc ===============================
CA = wb.create_sheet("Calc")
CA["A1"] = ("Calc — 每一步一列。三組：基準（C–F）、低（H–K）、高（M–P）；欄＝日曆年 2026–2029。分子按平台算、各自以該平台分母封頂後加總 ÷ 分母合計。"
            "同一批晶片在不同環節（承諾的記憶體、HBM 長約、CoWoS）的證據不相加：NVIDIA 只用 NVIDIA 本身的實績與承諾；HBM 長約只作第 7 節交叉檢核。")
CA["A1"].font = BOLD; CA["A1"].alignment = WRAP
CA.row_dimensions[1].height = 45
CASES = [("b", "基準", 3), ("lo", "低", 8), ("hi", "高", 13)]  # 起始欄 C、H、M
from openpyxl.utils import get_column_letter as CL
hdrs = ["步驟", "說明／單位"]
for j, h in enumerate(hdrs, 1):
    c = CA.cell(3, j, h); c.font = BOLD; c.fill = HDR
for case, lab, c0 in CASES:
    for j, y in enumerate(YRS):
        c = CA.cell(3, c0 + j, f"{lab} {y}"); c.font = BOLD; c.fill = HDR
CR = {}
r = 4


def P(key, case):
    return N[key][case]


def row(key, label, note, f, fmt="0.00", fill=None, bold=False, cases=CASES):
    """f(case, y, col) → 公式字串或值；None＝空白。"""
    global r
    CA.cell(r, 1, label); CA.cell(r, 2, note)
    if bold:
        CA.cell(r, 1).font = BOLD
    for case, _, c0 in cases:
        for j, y in enumerate(YRS):
            col = CL(c0 + j)
            v = f(case, y, col)
            if v is None:
                continue
            c = CA.cell(r, c0 + j, v); c.number_format = fmt
            if fill:
                c.fill = fill
    CR[key] = r
    r += 1


def R(key, col):
    return f"{col}{CR[key]}"


def head(t):
    global r
    r += 1
    CA.cell(r, 1, t).font = BOLD
    r += 1


yy = lambda y: y % 100
U = lambda k, y, case="b": P(f"u_{k}_{yy(y)}", "b")
it_amd = lambda y: "it_AMD8" if y == 2026 else "it_AMDR"   # 與 Siliconomics Deploy 頁相同：2026 MI355 8 卡、2027 起 Helios

head("1. 分母：Siliconomics 原始顆數讀數換成的成本等值 GW（規則 v1.1 分母；不含 S33 係數）＝ Σ 顆數 × 每顆 IT capex ÷ 每 GW IT capex")
row("d_nv", "NVIDIA", "GW；(Blackwell×it_GB＋Rubin×it_VR＋Rubin Ultra×it_RU＋Hopper×it_HP) ÷ 每 GW IT capex",
    lambda c, y, col: f"=({U('bw', y)}*{P('it_GB','b')}+{U('rb', y)}*{P('it_VR','b')}+{U('ru', y)}*{P('it_RU','b')}+{U('hp', y)}*{P('it_HP','b')})/{P(f'cgw_{yy(y)}','b')}")
row("d_amd", "AMD", "GW；顆數 × it_AMD8（2026）或 it_AMDR（2027 起）÷ 每 GW IT capex",
    lambda c, y, col: f"={U('amd', y)}*{P(it_amd(y),'b')}/{P(f'cgw_{yy(y)}','b')}")
row("d_xpu", "Google TPU＋其他 ASIC", "GW；Broadcom XPU 實績對應這一組（其他 ASIC 含非 Broadcom 的 Maia、昇騰，見待判斷 3）",
    lambda c, y, col: f"=({U('tpu', y)}*{P('it_TPU','b')}+{U('oth', y)}*{P('it_ASIC','b')})/{P(f'cgw_{yy(y)}','b')}")
row("d_trn", "AWS Trainium", "GW；無可量化證據", lambda c, y, col: f"={U('trn', y)}*{P('it_TRN','b')}/{P(f'cgw_{yy(y)}','b')}")
row("d_tot", "分母合計", "GW", lambda c, y, col: f"=SUM({col}{CR['d_nv']}:{col}{CR['d_trn']})", bold=True, fill=KEY)
row("d_chk", "核對：分母合計 − Siliconomics SIF_DeployGW_Supply", "應為 0", lambda c, y, col: f"={R('d_tot', col)}-{P(f'sup_{yy(y)}','b')}", "0.000")


def nvr_y(case, y):
    if y == 2026:
        return f"({P('nvr_GB300', case)}*(1-{P('vr26','b')})+{P('nvr_VR26', case)}*{P('vr26','b')})"
    return P("nvr_VR27", case)


head("2. NVIDIA 分子：實績（等級 1）＋採購承諾（等級 2）→ 成本等值 GW；以 NVIDIA 分母封頂")
row("nv_rpg", "每成本等值 GW 的 NVIDIA 營收", "$B/GW＝NVIDIA 營收 ÷ IT capex × 每 GW IT capex（2026 依 vr26 混合 GB300／VR200；2027 起 VR200）",
    lambda c, y, col: f"={nvr_y(c, y)}*{P(f'cgw_{yy(y)}','b')}")
row("nv_act", "實績：日曆年內已公布的資料中心營收", "$B；2026＝FY26 Q4 × 1/3（1 月）＋FY27 Q1＋Q2（到 2026-07）；Q3 未公布不算",
    lambda c, y, col: f"={P('nv_dc_fy26q4','b')}*{P('jan_share_nv','b')}+{P('nv_dc_q1','b')}+{P('nv_dc_q2','b')}" if y == 2026 else "=0")
row("nv_act_gw", "實績 → GW", "GW＝實績 ÷ 每 GW NVIDIA 營收", lambda c, y, col: f"={R('nv_act', col)}/{R('nv_rpg', col)}")
row("nv_inv", "存貨（2026-07 季底，2026 下半年出貨）", "$B 銷貨成本", lambda c, y, col: f"={P('nv_inv','b')}" if y == 2026 else "=0")


def cmt(y, lagref):
    s, e = y * 12 + 1, y * 12 + 12
    return (f"SUMIFS(NV_Month!$F$4:$F${NM_LAST},NV_Month!$B$4:$B${NM_LAST},\">=\"&({s}-{lagref}),"
            f"NV_Month!$B$4:$B${NM_LAST},\"<=\"&({e}-{lagref}))")


row("nv_cmt", "採購承諾：撐住該年出貨的部分", "$B 銷貨成本＝NV_Month 中「該年 1–12 月 − 時滯」月份的承諾加總",
    lambda c, y, col: f"={cmt(y, P('lag', c))}")
row("nv_cmt_rev", "存貨＋承諾 → 可支撐的資料中心營收", "$B＝(存貨＋承諾) × 資料中心比例 ÷ (1 − 毛利率) ÷ 承諾須涵蓋的成本比例",
    lambda c, y, col: f"=({R('nv_inv', col)}+{R('nv_cmt', col)})*{P('nv_dcsh', c)}/(1-{P('nv_gm', c)})/{P('nv_cov', c)}")
row("nv_cmt_gw", "存貨＋承諾 → GW", "GW", lambda c, y, col: f"={R('nv_cmt_rev', col)}/{R('nv_rpg', col)}")
row("nv_raw", "NVIDIA 證據合計（未封頂）", "GW", lambda c, y, col: f"={R('nv_act_gw', col)}+{R('nv_cmt_gw', col)}")
row("nv_ev", "NVIDIA 被撐住的量（封頂）", "GW＝MIN(證據, NVIDIA 分母)；超出部分不移到下一年（保守）",
    lambda c, y, col: f"=MIN({R('nv_raw', col)},{R('d_nv', col)})", fill=KEY, bold=True)
row("nv_sh", "NVIDIA 平台證據比例", "%", lambda c, y, col: f"={R('nv_ev', col)}/{R('d_nv', col)}", "0.0%")

head("3. Broadcom XPU 分子（等級 1；Google TPU＋其他 ASIC）")
row("av_act", "日曆年內已公布的 AI 營收", "$B；2026＝FY26 Q1 × 1/3（1 月）＋Q2＋Q3（到 2026-08-02）；FY26 Q4 未公布不算；2027 起無書面證據",
    lambda c, y, col: f"={P('av_q1','b')}*{P('jan_share_av','b')}+{P('av_q2','b')}+{P('av_q3','b')}" if y == 2026 else "=0")
row("av_ppc", "Broadcom 每顆 XPU 營收", "$K/顆＝核心售價＋HBM 代購比例 × 每顆 HBM 金額",
    lambda c, y, col: f"={P('core_tpu', c)}+{P('f_hbm', c)}*{P('hbm_tpu','b')}")
row("av_units", "XPU 顆數", "百萬顆＝AI 營收 × XPU 占比 ÷ 每顆營收", lambda c, y, col: f"={R('av_act', col)}*{P('av_xpu', c)}/{R('av_ppc', col)}")
row("av_gw", "XPU → GW", "GW＝顆數 × it_TPU ÷ 每 GW IT capex", lambda c, y, col: f"={R('av_units', col)}*{P('it_TPU','b')}/{P(f'cgw_{yy(y)}','b')}")
row("av_ev", "TPU＋其他 ASIC 被撐住的量（封頂）", "GW", lambda c, y, col: f"=MIN({R('av_gw', col)},{R('d_xpu', col)})", fill=KEY, bold=True)

head("4. AMD 分子（等級 1 實績；等級 2：OpenAI 首 1 GW）")
row("amd_act", "日曆年內已公布的 Data Center 營收", "$B；2026＝Q1＋Q2＋Q3；Q4 未公布不算",
    lambda c, y, col: f"={P('amd_q1','b')}+{P('amd_q2','b')}+{P('amd_q3','b')}" if y == 2026 else "=0")
row("amd_gw_act", "實績 → GW", "GW＝營收 × GPU 占比 ÷ 每顆售價 × 每顆 IT capex ÷ 每 GW IT capex",
    lambda c, y, col: f"={R('amd_act', col)}*{P('amd_gpu', c)}/{P('pa_amd', c)}*{P(it_amd(y),'b')}/{P(f'cgw_{yy(y)}','b')}")
row("amd_gw_oai", "OpenAI 首 1 GW 的 2027 部分 → 成本等值 GW", "GW＝實體 GW × 口徑係數 × AMD 持有成本比（Siliconomics Deploy 頁同法）",
    lambda c, y, col: f"={P('amd_oai27', c)}*{P('amd_oai_it', c)}*{P('hold_amd','b')}" if y == 2027 else "=0")
row("amd_ev", "AMD 被撐住的量（封頂）", "GW", lambda c, y, col: f"=MIN({R('amd_gw_act', col)}+{R('amd_gw_oai', col)},{R('d_amd', col)})", fill=KEY, bold=True)
row("trn_ev", "AWS Trainium 被撐住的量", "GW；無書面可量化證據（AWS 只有倍數說法、Marvell 客製營收屬作者估計）", lambda c, y, col: "=0", fill=KEY, bold=True)

head("5. 合計 → 證據比例（取代試作 In 頁 crd_es_26–29）")
row("ev_tot", "被撐住的量合計", "GW＝各平台封頂後加總", lambda c, y, col: f"={R('nv_ev', col)}+{R('av_ev', col)}+{R('amd_ev', col)}+{R('trn_ev', col)}", bold=True)
row("ev_sh", "Siliconomics 證據比例", "%＝被撐住的量 ÷ 分母合計", lambda c, y, col: f"={R('ev_tot', col)}/{R('d_tot', col)}", "0.0%", fill=KEY, bold=True)
row("ev_nv_part", "  其中 NVIDIA 貢獻", "%", lambda c, y, col: f"={R('nv_ev', col)}/{R('d_tot', col)}", "0.0%")
row("ev_nn_part", "  其中非 NVIDIA 貢獻", "%", lambda c, y, col: f"=({R('av_ev', col)}+{R('amd_ev', col)}+{R('trn_ev', col)})/{R('d_tot', col)}", "0.0%")
row("ev_old", "對照：試作原手填值", "%", lambda c, y, col: f"={P(f'old_es_{yy(y)}', c)}", "0.0%")
BASE = [("b", "基準", 3)]
for lk in ("lag_s0", "lag_s6"):
    m = N[lk]["b"]
    row(f"s_{lk}_cmt", f"敏感度（{lk[-1]} 個月時滯）：承諾撐住的 NVIDIA GW", "GW（其餘參數取基準）",
        lambda c, y, col, m=m: f"=({R('nv_inv', col)}+{cmt(y, m)})*{P('nv_dcsh', c)}/(1-{P('nv_gm', c)})/{P('nv_cov', c)}/{R('nv_rpg', col)}", cases=BASE)
    row(f"s_{lk}", f"敏感度（{lk[-1]} 個月時滯）：Siliconomics 證據比例", "%",
        lambda c, y, col, lk=lk: f"=(MIN({R('nv_act_gw', col)}+{R(f's_{lk}_cmt', col)},{R('d_nv', col)})+{R('av_ev', col)}+{R('amd_ev', col)}+{R('trn_ev', col)})/{R('d_tot', col)}",
        "0.0%", cases=BASE)
row("rng_lo", "區間下緣（低、基準、兩個時滯敏感度中的最小值）", "%",
    lambda c, y, col: f"=MIN({R('ev_sh', col)},{CL(8 + YRS.index(y))}{CR['ev_sh']},{R('s_lag_s0', col)},{R('s_lag_s6', col)})", "0.0%", cases=BASE, fill=KEY)
row("rng_hi", "區間上緣（高、基準、兩個時滯敏感度中的最大值）", "%",
    lambda c, y, col: f"=MAX({R('ev_sh', col)},{CL(13 + YRS.index(y))}{CR['ev_sh']},{R('s_lag_s0', col)},{R('s_lag_s6', col)})", "0.0%", cases=BASE, fill=KEY)

head("6. 狀態判定（規則 v1.1 C3；首次判定門檻 50%；並列 67%）")
row("ew", "Wattonomics 證據比例（暫值）", "%", lambda c, y, col: f"={P(f'ew_{yy(y)}','b')}", "0.0%")
row("hw", "Wattonomics 有證據（首次 50%）", "1／0", lambda c, y, col: f"=IF({R('ew', col)}>={P('crd_thr','b')},1,0)", "0")
row("hs", "Siliconomics 有證據（首次 50%）", "1／0", lambda c, y, col: f"=IF({R('ev_sh', col)}>={P('crd_thr','b')},1,0)", "0")
row("st", "狀態（50%）", "甲＝兩層都有；乙＝一層；丙＝都沒有",
    lambda c, y, col: f'=IF(AND({R("hw", col)}=1,{R("hs", col)}=1),"甲",IF(OR({R("hw", col)}=1,{R("hs", col)}=1),"乙","丙"))', "@", fill=KEY, bold=True)
row("st67", "狀態（67%）", "",
    lambda c, y, col: f'=IF(AND({R("ew", col)}>={P("crd_thr","hi")},{R("ev_sh", col)}>={P("crd_thr","hi")}),"甲",IF(OR({R("ew", col)}>={P("crd_thr","hi")},{R("ev_sh", col)}>={P("crd_thr","hi")}),"乙","丙"))', "@")
row("st_old", "對照：原手填值下的狀態（50%）", "",
    lambda c, y, col: f'=IF(AND({R("hw", col)}=1,{P(f"old_es_{yy(y)}", c)}>={P("crd_thr","b")}),"甲",IF(OR({R("hw", col)}=1,{P(f"old_es_{yy(y)}", c)}>={P("crd_thr","b")}),"乙","丙"))', "@")
row("st_chg", "狀態是否改變", "", lambda c, y, col: f'=IF({R("st", col)}={R("st_old", col)},"不變",{R("st_old", col)}&" → "&{R("st", col)})', "@", fill=KEY)

head("7. 交叉檢核：HBM 長約（上游已簽約；與 NVIDIA 記憶體承諾同一批 HBM，不加入分子）")
row("hbm_cov", "HBM 長約撐住的加速器比例（全平台）", "%＝(SK LTA 比例 × SK 份額＋美光已議約 × 美光份額〔只入高欄〕) × 供給 ÷ 原始需求；2026、2029 無長約",
    lambda c, y, col: (f"=({P('skh_lta', c)}*{P(f'skh_sh_{yy(y)}','b')}" + (f"+{P('mu_vm', c)}*{P('mu_sh_27','b')}" if y == 2027 else "")
                       + f")*{P(f'hbm_sup_{yy(y)}','b')}/{P(f'hbm_dem_{yy(y)}','b')}") if y in (2027, 2028) else None, "0.0%")
row("hbm_vs", "HBM 側 − 晶片側證據比例", "百分點；負值＝HBM 長約比晶片側證據少，取晶片側不會高估",
    lambda c, y, col: f"={R('hbm_cov', col)}-{R('ev_sh', col)}" if y in (2027, 2028) else None, "+0.0%;-0.0%")
CA.cell(r + 1, 1, "CoWoS：台積電沒有書面揭露的片數或客戶預訂（supply_evidence A11），證據＝0；只作供給上限參照。").font = Font(italic=True)
CA.column_dimensions["A"].width = 44; CA.column_dimensions["B"].width = 70
for j in range(3, 17):
    CA.column_dimensions[CL(j)].width = 10
CA.freeze_panes = "C4"
for k in ["ev_sh", "rng_lo", "rng_hi", "st", "st_chg", "d_tot", "ev_tot", "nv_sh"]:
    pass

# =============================== README 內容 ===============================
lines = [
    ("供給側證據比例 — Siliconomics 各年基準值中被證據撐住的比例（2026–2029）", BOLD),
    (f"產生：siliconomics/tools/supply_evidence_ratio.py（{datetime.date.today()}）；分母讀 siliconomics/model/{sil_f}；LibreOffice 重算後即為本檔。", None),
    ("規則：docs/plan/三層基準協調規則.md v1.1 C3＋歸類補充（Andy 2026-11-14 10:45）：法定文件的採購承諾算等級 2；法說口頭「售罄」只列參照；只有轉述者只列參照。", None),
    ("", None),
    ("怎麼讀：藍字＝輸入（In 頁，每格有標記與來源）；黑字＝公式。Calc 頁一列一步；三組參數：基準（C–F 欄）、低（H–K）、高（M–P）。「低」＝證據撐得最少的取法。", None),
    ("", None),
    ("方法（六步）", BOLD),
    ("1. 分母＝Siliconomics 原始顆數 × 每顆 IT capex ÷ 每 GW IT capex（成本等值 GW；不含 S33 係數），按平台分 NVIDIA、AMD、TPU＋其他 ASIC、Trainium。合計＝SIF_DeployGW_Supply。", None),
    ("2. NVIDIA：已公布資料中心營收（2026 年 1–7 月）÷ 每 GW NVIDIA 營收 → GW；另把 2026-07 存貨與 10-Q 採購承諾（FY27 剩餘 92、FY28 87、FY29 88、FY30 6，十億美元）按月攤開、加 3 個月時滯，"
     "以「÷（1 − 毛利率）× 資料中心比例」換成可支撐的營收，再 ÷ 每 GW NVIDIA 營收 → GW。基準把每一元承諾當成成品晶片的一元成本（保守）；高欄假設承諾只買記憶體（× 1/0.65）。", None),
    ("3. 非 NVIDIA：Broadcom AI 營收（2026 年 1–7 月）× XPU 占比 ÷ 每顆營收 → 顆數 → GW（撐 TPU＋其他 ASIC）；AMD Data Center 營收（2026 Q1–Q3）× GPU 占比 ÷ 售價 → GW，"
     "2027 加 OpenAI 首 1 GW 的剩餘部分；Trainium 無證據。", None),
    ("4. 各平台以自身分母封頂後加總 ÷ 分母合計。NVIDIA 承諾（主要是記憶體）與 SK 海力士 HBM 長約是同一批 HBM，長約只作第 7 節交叉檢核，不加入分子；CoWoS 無書面揭露＝0。", None),
    ("5. 會計年度對日曆年：NVIDIA（2–1 月）、Broadcom（11–10 月）按月分攤；AMD 會計年＝日曆年。", None),
    ("6. 區間＝低、高兩組參數與時滯 0／6 個月敏感度的最小、最大值。", None),
    ("", None),
    ("判斷類（預設值，待 Project 判斷）見報告 siliconomics/docs/reports/20261114_供給側證據比例.md 第 5 節。", None),
]
for i, (t, f) in enumerate(lines, 1):
    c = RD.cell(i, 1, t); c.alignment = Alignment(wrap_text=True, vertical="top")
    if f:
        c.font = f
RD.column_dimensions["A"].width = 100
RS = len(lines) + 2
RD.cell(RS, 1, "結果摘要（公式，讀 Calc 頁）").font = BOLD
for j, h in enumerate(["項目", "2026", "2027", "2028", "2029"]):
    c = RD.cell(RS + 1, 1 + j, h); c.font = BOLD; c.fill = HDR
for i, (k, lab, fm) in enumerate([("d_tot", "分母（GW）", "0.00"), ("ev_tot", "被撐住的量（GW，基準）", "0.00"), ("ev_sh", "證據比例（基準）", "0.0%"),
                                   ("rng_lo", "區間下緣", "0.0%"), ("rng_hi", "區間上緣", "0.0%"), ("ev_old", "原手填值", "0.0%"),
                                   ("st", "狀態（50%，基準）", "@"), ("st_chg", "狀態變化", "@")]):
    RD.cell(RS + 2 + i, 1, lab)
    for j in range(4):
        c = RD.cell(RS + 2 + i, 2 + j, f"=Calc!{CL(3 + j)}{CR[k]}"); c.number_format = fm
for j in range(2, 6):
    RD.column_dimensions[CL(j)].width = 14

out = sys.argv[1]
wb.save(out)
print("rows", {k: CR[k] for k in ["ev_sh", "rng_lo", "rng_hi", "st", "d_tot"]})
