# 標準版 DCF：無槓桿自由現金流（FCFF）＋ WACC — 個股模型共用模組 v2（2026-11-05；v1 為 NVIDIA v0.3 試作，v2 依工作單第二輪修正 Part 一）
# 依據：工作單「AI半導體 個股 DCF 改為標準版」（Andy 2026-11-05 20:42、20:45）、docs/plan/個股DCF標準規格.md v1.0；範本 spacex/build_spacex.py Val 頁、meta/build_meta.py β 寫法。
#
# 方法（每情境、每年；明確期到最後一個會計年度，通常 FY2032）
#   FCFF ＝ NOPAT ＋ 折舊攤銷 − 資本支出 − 營運資金增加；NOPAT ＝ EBIT ×（1 − 稅率）；股權報酬視為費用、不加回
#   折舊攤銷＝營收 × fc_da（conv_from 年起線性收斂到 fc_da_ss）；資本支出＝營收 ×（fc_cx 收斂到 fc_cx_ss ＋ fc_cx_d）
#   營運資金增加＝fc_wc ×（營收 − 前一年營收）
#   ke ＝ rf ＋ β ×（ERP ＋ CRP）；WACC ＝ (E × ke ＋ D × kd ×（1 − t）) ÷ (E ＋ D) ＋ wacc_adj；E＝px × sh_v，D＝debt_v ＋ lease_v
#   g ＝ MIN(g_cap, rf)（Andy 2026-11-05 20:45：3%，且不超過計價幣別 10 年期公債殖利率）；RONIC ＝ WACC ＋ ronic_sp
#   終值（最後一年年底）＝ NOPAT(最後一年) ×（1＋g）×（1 − g ÷ RONIC）÷（WACC − g）
#   期中慣例：第一年只計評價日後的 frac0（例 1/12），折現年數 frac0 ÷ 2；之後各年 frac0 ＋ (k − 1) ＋ 0.5；終值在 frac0 ＋ n − 1
#   每股價值＝MAX(0, (EV ＋ 淨現金與投資) ÷ sh_v)；推到目標時點 ×（1＋ke）^t_tgt − 期間每股股利 dps_tgt
#   淨現金與投資 dcf_nc ＝ cash_v ＋ inv_k × inv_v − debt_v − lease_v − mi_v
#
# 用法（四步；範例見 nvidia/build_nvidia.py「v0.3」段）
#   import fcff_dcf as FD
#   1. In 頁：rows = FD.in_rows(P)       → [("sec", 標題) 或 (鍵, 項目, 單位, 低, 基準, 高, 標記, 來源1, 來源2)]，接到公司自己的 In 列清單
#      （每列寫成定義名稱「鍵」，低／高欄寫成「鍵_lo」「鍵_hi」——各公司 In 迴圈的既有慣例）。P 的鍵見 in_rows 說明。
#   2. 明細頁：R = FD.write_detail(ws, r0, paths, years, frac0)   每情境一區塊、年度一欄；可對同一頁多次呼叫（例：另加 Sens 用路徑）
#      paths＝[{"name": "A", "label": "情境 A", "rev": {年: 儲存格}, "rev0": 前一年營收儲存格, "ebit": {年: 儲存格或公式}, "tax": {年: 儲存格或公式}}]
#   3. Val 摘要：S = FD.write_summary(wb, V, r0, R, names, weights)  寫折現率（ke、稅後 kd、權益占比、WACC、g、RONIC）緊接 DCF 彙總
#      （明確期現值、終值現值、EV、淨現金、股數、每股價值、推到 12 個月後、終值占比、隱含終值倍數、DCF 目標價）；共 SUMMARY_ROWS 列
#      定義名稱：ke、kd_at、w_e、wacc、ronic、dcf_nc。回傳 {鍵: 列號}；鍵 "dcf"＝DCF 目標價（情境加權），"v12"＝各情境推到 12 個月後
#   4. Sens：FD.write_sens(SE, r0, R, names, weights, comp_ref)   標準案例（WACC ±1pt、β、ERP、g ±1pt、RONIC＝WACC＋5pt、
#      資本支出 ÷ 營收 ±1pt、營運資金 ÷ 營收、淨現金不含權益投資、DCF 權重 0.3／0.7）對加權目標價；每格為閉式公式（與明細同式，
#      write_sens 另寫一列核對：閉式基準 − 明細基準＝0）。comp_ref 寫成「Val!$B$<列>」時，tools/relative_multiple.apply 會改成 tp_comp_adj。
#   5. 檢查列：FD.write_checks(ws, r0, nc_bs=…) 寫 WACC 由 In 頁輸入重算、淨現金與資產負債表核對（規格第 4 節）；明細每路徑另有
#      「FCFF 各列加總 − FCFF」與「終值占 EV 是否超過 80%」兩列。
#   FD.value_expr(R, name, **over) 回傳單一路徑每股價值的閉式公式（over：w、ke、g、ronic、cxd、wc、nc、tv="alt"），可用於其他對照路徑。
#
#
# ---------------- v2 新參數（2026-11-05 第二輪修正；未使用時輸出與 v1 相同，只有 Sens「β 低／高」列與 In 頁 M2 標題不同）----------------
#   (1) 資本支出、折舊逐年、逐情境讀模型自己的列（write_detail 的 paths 每條路徑可加）：
#       "da": {年: 儲存格或公式}、"cx": {年: 儲存格或公式}（金額，與營收同單位；例 {2027: "PL!D58", 2028: "PL!E58"}）。
#       須從第一年起連續；給到的年度：比例列＝該列 ÷ 營收（資本支出比例另加 fc_cx_d，Sens ±1pt 照舊）；
#       之後各年：由最後一個給定年度的比例線性收斂到 fc_da_ss／fc_cx_ss（最後一年等於穩態），且資本支出比例 ≥ 折舊比例 ×（1＋g_tv）
#       （收斂年度逐年取 MAX；write_detail(ss_floor=False) 可關閉）。只給 "da" 或只給 "cx" 亦可；未給的照 v1 占營收比例。
#       例：dict(name="B", …, da={y: f"PL!{C[y]}{r_da}" for y in (2027, 2028, 2029)}, cx={y: f"PL!{C[y]}{r_cx_B}" for y in (2027, 2028, 2029)})
#   (2) Stub 期（資產負債表日 ≠ 評價日 2026-12-31）：in_rows 的 P 加
#       stub=dict(bs_date="2026-07-26", val_date="2026-12-31",
#                 frac=(值或公式, 標記, 來源1, 來源2)       # 可省略：預設＝(評價日 − 資產負債表日) ÷ 365，寫成 =(DATE(…)-DATE(…))/365
#                 div=(值或公式, 標記, 來源1, 來源2))       # 該段已宣告股利／買回金額（金額單位同 debt_v）；可省略＝0
#       → In 列 bs_date（資產負債表日）、stub_frac（年）、stub_div（金額）；write_detail(…, stub=True)（或 R 由含 stub 的 in_rows 建好後傳 stub=True）。
#       淨現金滾到評價日＝dcf_nc（資產負債表）＋ stub FCFF − stub_div；stub FCFF 預設＝該路徑第一年（年度）FCFF × stub_frac，
#       路徑可加 "stub_fcff": "公式"（不含「=」；例 "0.25*PL!C80" 用模型季度數）。write_summary 在「＋淨現金與投資」之後多兩列
#       （「＋Stub 期 FCFF」「淨現金滾到評價日」；SUMMARY_ROWS＋2＝summary_rows(stub=True)）；明細頁不增列（每股價值公式內含）。
#       Sens 閉式同式（預設 stub FCFF 隨資本支出、營運資金案例重算；自訂 stub_fcff 視為常數）。write_checks 的 nc_bs 照舊對 dcf_nc（資產負債表）。
#   (3) β 敏感度＝β_u ± 0.3、各自重新加槓桿（write_sens 預設）：
#       write_sens(…, beta_u="beta_u", beta_du=0.3, relever="(1+(1-tax_w)*beta_de)")；beta_du 可給定義名稱（例 "beta_du"）。
#       In 頁須有 beta_u、beta_de（及 tax_w）；β 不是由 β_u 重新加槓桿的公司傳 beta_u=None，沿用 In 頁 beta_lo／beta_hi（v1 行為）。
#       建議同時把 In 頁 beta 的低／高欄改成同式（例 "=(beta_u-0.3)*(1+(1-tax_w)*beta_de)"），讓 In 與 Sens 一致。
#   (4) Sens「美元口徑折現」（非美國公司）：in_rows 的 P 加 rf_us=(0.05, 來源1, 來源2)（美國 10 年期公債殖利率）→ In 列 rf_us；
#       write_sens(…, usd=True) 加一列：ke＝rf_us ＋ β ×（ERP ＋ CRP）（CRP 照基準）、WACC 同式重算、g＝MIN(g_cap, rf_us)，其他不變
#       （kd、權重、現金流、匯率不變）；結果寫在「低」欄（單一案例）。美國公司此案例＝基準，省略（不傳 usd）。
#   (5) In 頁 M2 節標題不寫版號（「標準版 DCF」）；P 可加 m2_title 整句覆寫。
#   (6) ke 說明格的 β 說明由公司腳本提供：write_summary(…, beta_note="自算產業 β 重新加槓桿")；預設 None＝「Blume 調整」（v1 字樣）。
#
# 只用 openpyxl；不讀寫其他資料夾。各公司版面不同，本模組只管 FCFF／WACC 區塊；列標籤沿用下游（markets/invest_map、tools/target_summary）
# 讀取的「DCF 目標價（情境加權）」「DCF 情境價值推到 12 個月後」。
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.workbook.defined_name import DefinedName

BOLD = Font(bold=True); BLUE = Font(color="0000FF")
HDR = PatternFill("solid", fgColor="DDEBF7"); KEY = PatternFill("solid", fgColor="FFF2CC"); SUB = PatternFill("solid", fgColor="EDEDED")
SUMMARY_ROWS = 18          # write_summary 列數（無 stub）；有 stub 時多 2 列：summary_rows(stub=True)


def summary_rows(stub=False):
    return SUMMARY_ROWS + (2 if stub else 0)
IN_KEYS = ("rf", "erp", "crp", "beta", "blume_a", "kd", "tax_w", "wacc_adj", "g_cap", "g_tv", "ronic_sp", "t_tgt", "dps_tgt",
           "debt_v", "lease_v", "cash_v", "inv_v", "inv_k", "mi_v", "sh_v", "fc_da", "fc_da_ss", "fc_cx", "fc_cx_ss", "fc_cx_d", "fc_wc")
JUDG = ("erp", "crp", "beta", "g_cap", "ronic_sp", "fc_cx", "fc_cx_ss", "fc_wc", "inv_k")   # 黃底（Andy 可改）


def _dn(wb, name, ref):
    wb.defined_names[name] = DefinedName(name, attr_text=ref)


# ---------------- 1. In 頁參數 ----------------
def in_rows(P):
    """P（dict）：
      date：評價日字串；ccy：幣別（例 "US$"）；amt：金額單位（例 "US$B"）；shu：股數單位（例 "十億股"）
      rf：(值, 來源1, 來源2, 標籤)                 計價幣別 10 年期公債殖利率（Verified）
      betas：[(來源名, 原始 β, 來源說明), …]       1–3 個原始 β；基準＝各來源 Blume 調整（blume_a × β ＋ 1 − blume_a）的中位數
      beta_lohi：(低, 高, 說明)                    Sens 用（Analogy／Assumed）
      beta_fixed：(低, 基準, 高, 標記, 來源1, 來源2) 找不到原始 β 時用（同業中位數等），與 betas 擇一
      erp：(低, 基準, 高, 來源) 預設 4%／5%／6%；crp：(值, 標記, 來源1, 來源2) 預設 0
      kd：(值或公式, 標記, 來源1, 來源2)；tax_w：(值或公式, 標記, 來源1, 來源2)
      debt、lease、cash、inv、mi：(值, 標記, 來源1, 來源2)；inv＝現金以外的非營業投資（權益證券等）
      sh：(值或公式, 標記, 來源1, 來源2)          評價用稀釋股數
      fc_da、fc_cx、fc_wc：(低, 基準, 高, 標記, 來源1, 來源2)；fc_da_ss、fc_cx_ss（可省略，預設＝近期值）：(基準, 標記, 來源1, 來源2)
      g_cap（預設 0.03）、ronic_hi（預設 0.05）、t_tgt：(值, 說明)（預設 1 年）
      dps：(值或公式, 標記, 來源1, 來源2)  評價日到目標時點的每股股利合計（預設 0；規格第 3 節「扣除期間股利」）
    回傳 In 列清單。"""
    a, sh = P.get("amt", "US$B"), P.get("shu", "十億股")
    rows = [("sec", P.get("m2_title") or f"M2. 折現率與無槓桿自由現金流（標準版 DCF；tools/fcff_dcf.py；評價日 {P['date']}；黃底＝Andy 可改）")]
    add = rows.append
    rf = P["rf"]
    add(("rf", rf[3] if len(rf) > 3 else "無風險利率（計價幣別 10 年期公債殖利率）", "%", None, rf[0], None, "Verified", rf[1], rf[2]))
    if P.get("rf_us"):   # v2 (4)：Sens「美元口徑折現」用
        u = P["rf_us"]
        add(("rf_us", "美元口徑折現用：美國 10 年期公債殖利率（只用於 Sens「美元口徑折現」；基準用上一列當地利率＋固定匯率）", "%", None, u[0], None, "Verified", u[1], u[2]))
    if P.get("betas"):
        add(("blume_a", "Blume 調整係數（調整後 β＝a × 原始 β ＋ 1 − a）", "x", None, 0.67, None, "Derived", "Blume（1971）；同 Meta、Alphabet", "工作單 2026-11-05：β 用 Blume 調整"))
        ks = []
        for i, (nm, raw, src) in enumerate(P["betas"], 1):
            k = f"beta_{i}"; ks.append(k)
            add((k, f"原始 β：{nm}", "x", None, raw, None, "Interested-party", src, "資料商；未調整"))
        lo, hi, note = P["beta_lohi"]
        add(("beta", "β（各來源 Blume 調整後的中位數）", "x", lo, "=MEDIAN(" + ",".join(f"blume_a*{k}+1-blume_a" for k in ks) + ")", hi, "Derived",
             "＝MEDIAN(各來源 blume_a × 原始 β ＋ 1 − blume_a)", note))
    else:
        lo, b, hi, tag, s1, s2 = P["beta_fixed"]
        add(("beta", "β", "x", lo, b, hi, tag, s1, s2))
    e = P.get("erp", (0.04, 0.05, 0.06, "工作單 2026-11-05 預設 5%（同 SpaceX、Alphabet）"))
    add(("erp", "股權風險溢酬 ERP（成熟市場）", "%", e[0], e[1], e[2], "Assumed", e[3], "Sens：低／高"))
    c = P.get("crp", (0.0, "Assumed", "美國等成熟市場：0", "非美國市場：Damodaran 最新 Country Risk Premium 表（Andy 2026-11-05 20:45）"))
    add(("crp", "國家風險溢酬 CRP（ke＝rf ＋ β ×（ERP ＋ CRP）；不設股權成本下限）", "%", None, c[0], None, c[1], c[2], c[3]))
    k = P["kd"]; add(("kd", "稅前債務成本 kd", "%", None, k[0], None, k[1], k[2], k[3]))
    t = P["tax_w"]; add(("tax_w", "WACC 用稅率（債務稅盾）", "%", None, t[0], None, t[1], t[2], t[3]))
    add(("wacc_adj", "WACC 調整（Sens 用；基準 0）", "%", -0.01, 0.0, 0.01, "Assumed", "Sens：WACC −1pt／＋1pt", "—"))
    add(("g_cap", "永續成長率上限", "%", None, P.get("g_cap", 0.03), None, "Assumed", "Andy 2026-11-05 20:45：3%", "g＝MIN(上限, rf)"))
    add(("g_tv", "永續成長率 g（最後一年後；＝MIN(g_cap, rf)）", "%", "=g_tv-0.01", "=MIN(g_cap,rf)", "=MIN(g_tv+0.01,rf)", "Derived",
         "Andy 2026-11-05 20:45：3%，且不超過計價幣別 10 年期公債殖利率", "Sens：−1pt／＋1pt（上限 rf）"))
    add(("ronic_sp", "永續期新投資報酬率 RONIC 相對 WACC 的差（RONIC＝WACC＋本值）", "%", None, 0.0, P.get("ronic_hi", 0.05), "Assumed",
         "工作單 2026-11-05：RONIC＝WACC（同 SpaceX、Alphabet）", "Sens：WACC＋5pt"))
    tt = P.get("t_tgt", (1.0, "目標價為 12 個月後（與 v0.2 DCF 目標價同一時點）"))
    add(("t_tgt", "目標時點距評價日", "年", None, tt[0], None, "Derived", tt[1], "每股價值 ×（1＋ke）^t_tgt"))
    dp = P.get("dps", (0, "Assumed", "未配息或未設定", "—"))
    add(("dps_tgt", "評價日到目標時點的每股股利合計（推到目標時點後扣除）", P.get("ccy", "US$"), None, dp[0], None, dp[1], dp[2], dp[3]))
    for key, lab in (("debt", "債務本金（WACC 負債與淨負債）"), ("lease", "租賃負債（WACC 負債與淨負債）"), ("cash", "現金與有價債務證券"),
                     ("inv", "非營業投資（權益證券等；評價時加入）"), ("mi", "少數股權")):
        v = P[key]; add((f"{key}_v", lab, a, None, v[0], None, v[1], v[2], v[3]))
    if P.get("stub"):   # v2 (2)：資產負債表日到評價日的 stub 期
        st = P["stub"]; bd, vd = st["bs_date"], st.get("val_date", "2026-12-31")
        D_ = lambda d: "DATE({},{},{})".format(*[int(x) for x in d.split("-")])
        add(("bs_date", "資產負債表日（淨現金、債務等科目的日期）", "日期", None, bd, None, "Verified", st.get("bs_src", "上列資產負債科目的來源"), f"評價日 {vd}；淨現金以 stub 期滾到評價日"))
        fr_ = st.get("frac", (f"=({D_(vd)}-{D_(bd)})/365", "Derived", f"＝(評價日 {vd} − 資產負債表日 {bd}) ÷ 365", "stub FCFF＝各路徑第一年 FCFF × 本值（年度比例）"))
        add(("stub_frac", "Stub 期長度（資產負債表日 → 評價日；年）", "年", None, fr_[0], None, fr_[1], fr_[2], fr_[3]))
        dv = st.get("div", (0, "Assumed", "該段無已宣告股利／買回", "—"))
        add(("stub_div", "Stub 期已宣告股利／買回（資產負債表日 → 評價日；從淨現金扣除）", a, None, dv[0], None, dv[1], dv[2], dv[3]))
    add(("inv_k", "非營業投資計入評價的比例（1＝全部；Sens 低＝0）", "x", 0, 1, None, "Assumed", "帳面價值計入（公開權益證券為公允價值）", "Sens：不計"))
    s = P["sh"]; add(("sh_v", "評價用稀釋股數", sh, None, s[0], None, s[1], s[2], s[3]))
    for key, lab in (("fc_da", "折舊攤銷 ÷ 營收（明確期近期）"), ("fc_cx", "資本支出 ÷ 營收（明確期近期）"), ("fc_wc", "營運資金 ÷ 營收（營運資金增加＝本值 × 營收增量）")):
        v = P[key]; add((key, lab, "%", v[0], v[1], v[2], v[3], v[4], v[5]))
        if key != "fc_wc":
            ss = P.get(key + "_ss", (f"={key}", "Assumed", "預設＝近期比例（不收斂）", "—"))
            add((key + "_ss", lab.replace("明確期近期", "長期穩態（最後一年）"), "%", None, ss[0], None, ss[1], ss[2], ss[3]))
    add(("fc_cx_d", "資本支出 ÷ 營收的額外加減（Sens 用；基準 0）", "pt", -0.01, 0.0, 0.01, "Assumed", "Sens：±1 個百分點", "—"))
    return rows


def yellow(IN, keys=JUDG):
    for r in range(4, IN.max_row + 1):
        if IN.cell(r, 1).value in keys:
            for c in range(1, 10): IN.cell(r, c).fill = KEY


# ---------------- 2. 明細頁 ----------------
def write_detail(ws, r0, paths, years, frac0, conv_from=None, title=None, fmt="#,##0.00", note=None, stub=False, ss_floor=True):
    """ws：明細工作表；r0：起始列；years：明確期會計年度（遞增）；frac0：第一年在評價日後的比例；
    conv_from：折舊與資本支出比例開始線性收斂的年度（預設＝第二年；最後一年等於 *_ss；只用於未給模型列的路徑）。
    v2：路徑可加 "da"／"cx"（模型自己的逐年金額列；見檔頭 (1)）、"stub_fcff"；stub＝True 時每股價值含 stub 期（見檔頭 (2)）。
    ss_floor：模型列之後的收斂年度，資本支出比例 ≥ 折舊比例 ×（1＋g_tv）。回傳 R（列號對照）。"""
    n = len(years); col = {y: get_column_letter(3 + i) for i, y in enumerate(years)}
    last = col[years[-1]]
    conv_from = conv_from or years[1]
    r = r0
    ws.cell(r, 1, title or "FCFF 明細（金額單位同 PL 頁；年度一欄；期中慣例）").font = BOLD; r += 1
    if note: ws.cell(r, 1, note); r += 1
    hdr_r = r
    for j, h in enumerate(["項目", "說明"] + [f"FY{y}" for y in years], 1):
        cc = ws.cell(r, j, h); cc.font = BOLD; cc.fill = HDR
    r += 1
    fr, tr = r, r + 1
    ws.cell(fr, 1, "評價日後占該年度比例")
    ws.cell(tr, 1, "折現年數（期中，距評價日）")
    for i, y in enumerate(years):
        ws[f"{col[y]}{fr}"] = round(frac0, 6) if i == 0 else 1
        ws[f"{col[y]}{tr}"] = round(frac0 / 2, 6) if i == 0 else round(frac0 + i - 0.5, 6)
        ws[f"{col[y]}{fr}"].number_format = "0.000"; ws[f"{col[y]}{tr}"].number_format = "0.000"
    ws.cell(fr, 2, f"第一年只計評價日後 {frac0:.4f} 年"); ws.cell(tr, 2, "第一年＝frac0 ÷ 2；之後＝frac0 ＋ (k − 1) ＋ 0.5")
    tvt = tr + 1
    ws.cell(tvt, 1, "終值折現年數（最後一年年底）"); ws.cell(tvt, 3, round(frac0 + n - 1, 6)).number_format = "0.000"
    r = tvt + 2
    R = {"sheet": ws.title, "col": col, "years": list(years), "frac": fr, "t": tr, "tvt": f"{ws.title}!$C${tvt}", "hdr": hdr_r, "paths": {}, "stub": bool(stub)}
    k_ = lambda y: max(0, years.index(y) - years.index(conv_from) + 1)
    nconv = years.index(years[-1]) - years.index(conv_from) + 1
    for p in paths:
        D = {}
        ws.cell(r, 1, p["label"]).font = BOLD
        for cc in range(1, 3 + n): ws.cell(r, cc).fill = SUB
        r += 1
        def row(key, lab, f, fm=fmt, nt=""):
            nonlocal r
            ws.cell(r, 1, lab); ws.cell(r, 2, nt)
            for y in years:
                v = f(y); c = ws[f"{col[y]}{r}"]; c.value = v; c.number_format = fm
            D[key] = r; r += 1
        prev = lambda y: p["rev0"] if y == years[0] else f"{col[years[years.index(y) - 1]]}{D['rev']}"
        row("rev", "營收", lambda y: "=" + p["rev"][y])
        row("drev", "營收增量", lambda y: f"={col[y]}{D['rev']}-{prev(y)}")
        row("ebit", "營業利益 EBIT", lambda y: p["ebit"][y] if str(p["ebit"][y]).startswith("=") else "=" + p["ebit"][y], nt=p.get("ebit_note", ""))
        row("tax", "稅率", lambda y: p["tax"][y] if str(p["tax"][y]).startswith("=") else "=" + p["tax"][y], "0.0%")
        row("nopat", "NOPAT＝EBIT ×（1 − 稅率）", lambda y: f"={col[y]}{D['ebit']}*(1-{col[y]}{D['tax']})", nt="股權報酬視為費用、不加回")
        cv = lambda a, b: (lambda y: f"={a}+({b}-{a})*{min(k_(y), nconv)}/{nconv}" if k_(y) > 0 else f"={a}")
        # v2 (1)：模型自己的逐年金額列（da／cx）；給到的年度＝金額 ÷ 營收，之後由最後一個給定年度的比例線性收斂到穩態
        def given(key):
            g_ = p.get(key) or {}
            ys = [y for y in years if y in g_]
            assert ys == years[:len(ys)], (p["name"], key, "模型列須從第一年起連續", ys)
            return g_, ys
        def mdl(key, ss, base_of_last):
            g_, ys = given(key)
            yl = ys[-1]; nn = years.index(years[-1]) - years.index(yl)
            def f(y):
                if y in g_:
                    v = str(g_[y]); v = v[1:] if v.startswith("=") else v
                    return f"=({v})/{col[y]}{D['rev']}"
                k = years.index(y) - years.index(yl)
                a = base_of_last(col[yl])
                return f"={a}+({ss}-{a})*{k}/{nn}"
            return f, ys
        if p.get("da"):
            f_da, ys_da = mdl("da", "fc_da_ss", lambda c: f"{c}{r}")   # r＝正在寫的列（dar）
            row("dar", "折舊攤銷 ÷ 營收", f_da, "0.00%", f"FY{ys_da[0]}–FY{ys_da[-1]}＝模型列 ÷ 營收" + (f"；之後線性收斂到長期穩態" if ys_da[-1] != years[-1] else ""))
        else:
            row("dar", "折舊攤銷 ÷ 營收", cv("fc_da", "fc_da_ss"), "0.00%", f"FY{conv_from} 起線性收斂到長期穩態")
        if p.get("cx"):
            f_cx, ys_cx = mdl("cx", "fc_cx_ss", lambda c: f"({c}{r}-fc_cx_d)")   # r＝正在寫的列（cxr）
            def f_cx2(y):
                v = f_cx(y)
                if y in p["cx"]: return v + "+fc_cx_d"
                if ss_floor: return f"=MAX({v[1:]},{col[y]}{D['dar']}*(1+g_tv))+fc_cx_d"
                return v + "+fc_cx_d"
            row("cxr", "資本支出 ÷ 營收（含 fc_cx_d）", f_cx2, "0.00%", f"FY{ys_cx[0]}–FY{ys_cx[-1]}＝模型列 ÷ 營收" +
                (("；之後線性收斂到長期穩態" + ("，且 ≥ 折舊比例 ×（1＋g）" if ss_floor else "")) if ys_cx[-1] != years[-1] else ""))
        else:
            row("cxr", "資本支出 ÷ 營收（含 fc_cx_d）", lambda y: cv("fc_cx", "fc_cx_ss")(y) + "+fc_cx_d", "0.00%", f"FY{conv_from} 起線性收斂到長期穩態")
        row("da", "＋折舊攤銷", lambda y: f"={col[y]}{D['rev']}*{col[y]}{D['dar']}")
        row("cx", "− 資本支出", lambda y: f"={col[y]}{D['rev']}*{col[y]}{D['cxr']}")
        row("dwc", "− 營運資金增加（fc_wc × 營收增量）", lambda y: f"=fc_wc*{col[y]}{D['drev']}")
        row("fcff", "無槓桿自由現金流 FCFF", lambda y: f"={col[y]}{D['nopat']}+{col[y]}{D['da']}-{col[y]}{D['cx']}-{col[y]}{D['dwc']}")
        row("df", "折現因子 1 ÷（1＋WACC）^年數", lambda y: f"=1/(1+wacc)^{col[y]}${tr}", "0.0000")
        row("pv", "現值（FCFF × 評價日後比例 × 折現因子）", lambda y: f"={col[y]}{D['fcff']}*{col[y]}${fr}*{col[y]}{D['df']}")
        r += 1
        def one(key, lab, f, fm=fmt, nt=""):
            nonlocal r
            ws.cell(r, 1, lab); c = ws.cell(r, 3, f); c.number_format = fm; ws.cell(r, 2, nt); D[key] = r; r += 1
        one("pvx", "明確期現值合計", f"=SUM(C{D['pv']}:{last}{D['pv']})")
        one("tv", "終值（最後一年年底）＝NOPAT ×（1＋g）×（1 − g ÷ RONIC）÷（WACC − g）", f"={last}{D['nopat']}*(1+g_tv)*(1-g_tv/ronic)/(wacc-g_tv)")
        one("pvtv", "終值現值", f"=C{D['tv']}/(1+wacc)^{R['tvt']}")
        one("ev", "企業價值 EV", f"=C{D['pvx']}+C{D['pvtv']}")
        if stub:   # v2 (2)：淨現金滾到評價日＝dcf_nc ＋ stub FCFF − stub_div（明細不增列，寫在每股價值公式內）
            D["stub_x"] = p.get("stub_fcff") or f"{ws.title}!$C${D['fcff']}*stub_frac"
            one("val", "每股價值（評價日）＝MAX(0,(EV＋淨現金與投資＋stub 期 FCFF − stub 期已宣告股利) ÷ 股數)", f"=MAX(0,(C{D['ev']}+dcf_nc+{D['stub_x']}-stub_div)/sh_v)",
                nt="stub FCFF＝" + ("自訂（路徑 stub_fcff）" if p.get("stub_fcff") else "第一年 FCFF × stub_frac"))
        else:
            one("val", "每股價值（評價日）＝MAX(0,(EV＋淨現金與投資) ÷ 股數)", f"=MAX(0,(C{D['ev']}+dcf_nc)/sh_v)")
        one("v12", "每股價值推到目標時點 ×（1＋ke）^t_tgt − 期間股利", f"=C{D['val']}*(1+ke)^t_tgt-dps_tgt", nt="DCF 情境價值推到 12 個月後")
        one("tvsh", "終值現值占 EV", f"=C{D['pvtv']}/C{D['ev']}", "0.0%")
        one("mult", "隱含終值 EV／EBITDA（最後一年後一年預估）", f"=C{D['tv']}/(({last}{D['ebit']}+{last}{D['da']})*(1+g_tv))", "0.0", "EBITDA＝EBIT＋折舊攤銷")
        one("ssck", "一致性：最後一年資本支出 ≥ 折舊 ×（1＋g）", f'=IF({last}{D["cx"]}>={last}{D["da"]}*(1+g_tv),"通過","未通過")', "@")
        one("fsum", "明確期 FCFF 合計（× 評價日後比例，未折現）", f"=SUMPRODUCT(C{D['fcff']}:{last}{D['fcff']},C${fr}:{last}${fr})")
        one("ck_f", "檢查：FCFF 各列加總 − FCFF（各年合計，應為 0）", f"=SUM(C{D['nopat']}:{last}{D['nopat']})+SUM(C{D['da']}:{last}{D['da']})-SUM(C{D['cx']}:{last}{D['cx']})-SUM(C{D['dwc']}:{last}{D['dwc']})-SUM(C{D['fcff']}:{last}{D['fcff']})", "0.000000")
        one("ck_tv", "檢查：終值占 EV 超過 80%？（規格第 4 節；超過時在報告標示）", f'=IF(C{D["tvsh"]}>0.8,"超過 80%，需在報告標示","未超過")', "@")
        R["paths"][p["name"]] = D
        r += 1
    R["end"] = r
    return R


def _ref(R, row, a=None, b=None, abs_=True):
    s = R["sheet"]; c0, c1 = R["col"][R["years"][0]], R["col"][R["years"][-1]]
    return f"{s}!${c0}${row}:${c1}${row}"


def value_expr(R, name, w="wacc", ke="ke", g="g_tv", ronic=None, cxd="fc_cx_d", wc="fc_wc", nc="dcf_nc", push=True, tv="vd"):
    """單一路徑每股價值的閉式公式（不含開頭「=」）；與明細同式。push＝True 時乘（1＋ke）^t_tgt 再扣 dps_tgt。
    tv＝"vd"：價值驅動式（基準）；"alt"：終值改用最後一年 FCFF ×（1＋g）÷（WACC − g）（規格第 4 節敏感度）。"""
    D = R["paths"][name]; s = R["sheet"]; last = R["col"][R["years"][-1]]
    ronic = ronic or f"({w}+ronic_sp)"
    rg = lambda k: _ref(R, D[k])
    cx = f"({rg('cxr')}-fc_cx_d+{cxd})"
    fc = f"({rg('nopat')}+{rg('rev')}*({rg('dar')}-{cx})-{wc}*{rg('drev')})"
    pvx = f"SUMPRODUCT({fc}*{_ref(R, R['frac'])}/(1+{w})^{_ref(R, R['t'])})"
    L = lambda k: f"{s}!${last}${D[k]}"
    if tv == "alt":
        fl = f"({L('nopat')}+{L('rev')}*({L('dar')}-({L('cxr')}-fc_cx_d+{cxd}))-{wc}*{L('drev')})"
        tvx = f"{fl}*(1+{g})/({w}-{g})/(1+{w})^{R['tvt']}"
    else:
        tvx = f"{L('nopat')}*(1+{g})*(1-{g}/{ronic})/({w}-{g})/(1+{w})^{R['tvt']}"
    stub = ""
    if R.get("stub"):   # v2 (2)
        if D.get("stub_x", "").endswith("*stub_frac") and D["stub_x"].startswith(f"{s}!$C${D['fcff']}"):
            c0 = R["col"][R["years"][0]]; F = lambda k: f"{s}!${c0}${D[k]}"
            stub = f"+({F('nopat')}+{F('rev')}*({F('dar')}-({F('cxr')}-fc_cx_d+{cxd}))-{wc}*{F('drev')})*stub_frac-stub_div"
        else:
            stub = f"+({D['stub_x']})-stub_div"
    v = f"MAX(0,({pvx}+{tvx}+{nc}{stub})/sh_v)"
    return f"({v}*(1+{ke})^t_tgt-dps_tgt)" if push else v


# ---------------- 3. Val 摘要 ----------------
def write_summary(wb, V, r0, R, names, weights, cols="BCD", fmt="#,##0.00", title=None, px="px", beta_note=None):
    """names：三情境路徑名（對應 B／C／D 欄）；weights：情境權重名稱（例 ["tp_wA","tp_wB","tp_wC"]）。寫 summary_rows(R["stub"]) 列。
    beta_note：ke 說明格中 β 的說明（v2；例 "自算產業 β 重新加槓桿"）；None＝「Blume 調整」（v1 字樣）。"""
    r = r0; S = {}; sh = R["sheet"]
    def put(key, lab, vals, fm=fmt, note=None, bold=False, fill=None):
        nonlocal r
        V.cell(r, 1, lab)
        if bold: V.cell(r, 1).font = BOLD
        for j, v in enumerate(vals):
            if v is not None: V.cell(r, 2 + j, v).number_format = fm
        if note is not None: V.cell(r, 5, note)
        if fill:
            for cc in range(1, 6): V.cell(r, cc).fill = fill
        if key: S[key] = r
        r += 1
    put("title", title or "0.1 DCF（標準版：無槓桿自由現金流 FCFF＋WACC；v0.3；明細見 DCF 頁）", [], bold=True)
    for j, h in enumerate(["項目", "A", "B", "C", "說明"], 1):
        c = V.cell(r, j, h); c.font = BOLD; c.fill = HDR
    S["h"] = r; r += 1
    put("ke", "股權成本 ke＝rf ＋ β ×（ERP ＋ CRP）", ["=rf+beta*(erp+crp)"], "0.00%",
        note='="rf "&TEXT(rf,"0.00%")&"、β "&TEXT(beta,"0.000")&"（' + (beta_note or "Blume 調整") + '）、ERP "&TEXT(erp,"0.0%")&"、CRP "&TEXT(crp,"0.0%")')
    _dn(wb, "ke", f"{V.title}!$B${S['ke']}")
    put("kd_at", "稅後債務成本 kd ×（1 − t）", ["=kd*(1-tax_w)"], "0.00%", note='="kd "&TEXT(kd,"0.00%")&"、稅率 "&TEXT(tax_w,"0.0%")')
    _dn(wb, "kd_at", f"{V.title}!$B${S['kd_at']}")
    put("w_e", "權益占資本 E ÷（E＋D）（市值；D＝債務＋租賃）", [f"={px}*sh_v/({px}*sh_v+debt_v+lease_v)"], "0.0%",
        note=f'="E "&TEXT({px}*sh_v,"#,##0")&"、D "&TEXT(debt_v+lease_v,"#,##0.0")')
    _dn(wb, "w_e", f"{V.title}!$B${S['w_e']}")
    put("wacc", "WACC", [f"=w_e*ke+(1-w_e)*kd_at+wacc_adj"], "0.00%", bold=True, note="(E × ke ＋ D × kd ×（1 − t）) ÷ (E ＋ D)；wacc_adj 為 Sens 用（基準 0）")
    _dn(wb, "wacc", f"{V.title}!$B${S['wacc']}")
    put("g", "永續成長率 g＝MIN(g_cap, rf)", ["=g_tv"], "0.00%", note='="上限 "&TEXT(g_cap,"0.0%")&"、rf "&TEXT(rf,"0.00%")')
    put("ronic", "永續期 RONIC＝WACC＋ronic_sp", ["=wacc+ronic_sp"], "0.00%", note="終值再投資率＝g ÷ RONIC")
    _dn(wb, "ronic", f"{V.title}!$B${S['ronic']}")
    D = {n: R["paths"][n] for n in names}
    x = lambda k: [f"={sh}!$C${D[n][k]}" for n in names]
    put("pvx", "明確期 FCFF 現值", x("pvx"), note=f"FY{R['years'][0]}–FY{R['years'][-1]}；FCFF＝NOPAT＋折舊 − 資本支出 − 營運資金增加（DCF 頁）")
    put("pvtv", "終值現值", x("pvtv"), note="NOPAT ×（1＋g）×（1 − g ÷ RONIC）÷（WACC − g），折到評價日")
    put("ev", "企業價值 EV", x("ev"), bold=True)
    put("nc", "＋淨現金與投資（現金＋有價證券＋權益投資 − 債務 − 租賃 − 少數股權）", ["=cash_v+inv_k*inv_v-debt_v-lease_v-mi_v"] * len(names))
    _dn(wb, "dcf_nc", f"{V.title}!$B${S['nc']}")
    if R.get("stub"):   # v2 (2)
        put("stub", "＋Stub 期 FCFF（資產負債表日 → 評價日）", ["=" + D[n]["stub_x"] for n in names],
            note='="資產負債表日 "&bs_date&"；stub_frac "&TEXT(stub_frac,"0.000")&" 年（第一年 FCFF × stub_frac）"')
        put("ncr", "淨現金滾到評價日（＝淨現金與投資＋stub FCFF − stub 期已宣告股利／買回）", [f"={c}{S['nc']}+{c}{S['stub']}-stub_div" for c in cols[:len(names)]],
            note='="已宣告股利／買回 "&TEXT(stub_div,"#,##0.00")')
    put("sh", "評價用稀釋股數", ["=sh_v"] * len(names), "#,##0.000")
    put("dcfv", "DCF 情境每股價值（評價日）", x("val"))
    put("v12", "DCF 情境價值推到 12 個月後（×（1＋ke）^t_tgt − 期間股利）", x("v12"), fill=KEY)
    put("tvsh", "  其中：終值現值占 EV", x("tvsh"), "0.0%")
    put("mult", "  隱含終值 EV／EBITDA", x("mult"), "0.0", note="終值 ÷ 最後一年後一年 EBITDA（EBIT＋折舊）")
    put("dcf", "DCF 目標價（情境加權）", ["=" + "+".join(f"{wgt}*{c}{S['v12']}" for wgt, c in zip(weights, cols))], fill=KEY,
        note="Σ 情境權重 × 情境價值（已推到 12 個月後）")
    assert r - r0 == summary_rows(R.get("stub")), (r - r0, summary_rows(R.get("stub")))
    return S


# ---------------- 4. Sens ----------------
SENS_CASES = [
    # (名稱, 設定, 低, 高, 鍵)；鍵寫在 Sens 表「鍵」欄（tools/sens_coverage 用）
    ("WACC ±1pt", "WACC −1pt／＋1pt（RONIC 隨 WACC）", dict(w="(wacc+wacc_adj_lo)"), dict(w="(wacc+wacc_adj_hi)"), "wacc_adj"),
    ("β 低／高", '="β "&TEXT(beta_lo,"0.00")&"／"&TEXT(beta_hi,"0.00")&"（基準 "&TEXT(beta,"0.00")&"）"', "beta_lo", "beta_hi", "beta"),
    ("ERP 低／高", '="ERP "&TEXT(erp_lo,"0%")&"／"&TEXT(erp_hi,"0%")', "erp_lo", "erp_hi", "erp"),
    ("永續成長率 g ±1pt", '="g "&TEXT(g_tv_lo,"0.0%")&"／"&TEXT(g_tv_hi,"0.0%")', dict(g="g_tv_lo"), dict(g="g_tv_hi"), "g_tv"),
    ("RONIC＝WACC＋5pt", "RONIC＝WACC → WACC＋5pt（只有高）", None, dict(ronic="(wacc+ronic_sp_hi)"), "ronic_sp"),
    ("資本支出 ÷ 營收 ±1pt", "每年資本支出比例 −1pt／＋1pt", dict(cxd="fc_cx_d_lo"), dict(cxd="fc_cx_d_hi"), "fc_cx_d"),
    ("營運資金 ÷ 營收 低／高", '="營運資金比例 "&TEXT(fc_wc_lo,"0%")&"／"&TEXT(fc_wc_hi,"0%")&"（基準 "&TEXT(fc_wc,"0.0%")&"）"', dict(wc="fc_wc_lo"), dict(wc="fc_wc_hi"), "fc_wc"),
    ("淨現金不含權益投資", "權益證券等非營業投資不計（只有低）", dict(nc="(cash_v+inv_k_lo*inv_v-debt_v-lease_v-mi_v)"), None, "inv_k"),
    ("終值改用 FCFF ×（1＋g）", "終值＝最後一年 FCFF ×（1＋g）÷（WACC − g）（替代法；只有高）", None, dict(tv="alt"), ""),
    ("DCF 權重 0.3／0.7", "tp_w_dcf 0.3／0.7", "w03", "w07", "tp_w_dcf"),
]


def write_sens(SE, r0, R, names, weights, comp_ref, base_ref, w_dcf="tp_w_dcf", fmt="#,##0.00", title=None, cases=SENS_CASES, w_dcf_lohi=False,
               beta_u="beta_u", beta_du=0.3, relever="(1+(1-tax_w)*beta_de)", usd=False):
    """SE 第 r0 列起：標題、表頭、各案例（低／高加權目標價與相對基準）、閉式核對列。base_ref：基準加權目標價儲存格（例 "Val!$B$37"）。回傳 (案例首列, 案例末列, 下一列)。
    v2：β 案例＝(beta_u ∓ beta_du) × relever（各自重新加槓桿）；beta_u＝None 時沿用 In 頁 beta_lo／beta_hi。usd＝True 時在 DCF 權重案例前加「美元口徑折現」（需 In 頁 rf_us）。"""
    SE.cell(r0, 1, title or "目標價參數敏感度（v0.3 標準版 DCF）：加權目標價一次動一個參數；Comparable 不變（DCF 權重案例除外）").font = BOLD
    for j, h in enumerate(["參數", "設定（低／高）", "低：加權目標價", "低：相對基準", "高：加權目標價", "高：相對基準", "基準加權目標價", "鍵"], 1):
        c = SE.cell(r0 + 1, j, h); c.font = BOLD; c.fill = HDR

    def ke_of(beta="beta", erp="erp"):
        return f"(rf+{beta}*({erp}+crp))"

    def tp(over, wd=w_dcf):
        o = dict(over or {})
        ke = o.pop("ke", "ke")
        w = o.pop("w", "wacc")
        dcf = "+".join(f"{wgt}*{value_expr(R, n, w=w, ke=ke, **o)}" for wgt, n in zip(weights, names))
        return f"={wd}*({dcf})+(1-{wd})*{comp_ref}"

    bl = lambda sg: f"(({beta_u}{sg}{beta_du})*{relever})"
    def resolve(spec):
        if spec is None: return None
        if spec in ("beta_lo", "beta_hi"):
            b = (bl("-") if spec == "beta_lo" else bl("+")) if beta_u else spec
            k = ke_of(beta=b); return tp(dict(ke=k, w=f"(w_e*{k}+(1-w_e)*kd_at+wacc_adj)"))
        if spec == "usd":
            k = "(rf_us+beta*(erp+crp))"; return tp(dict(ke=k, w=f"(w_e*{k}+(1-w_e)*kd_at+wacc_adj)", g="MIN(g_cap,rf_us)"))
        if spec in ("erp_lo", "erp_hi"):
            k = ke_of(erp=spec); return tp(dict(ke=k, w=f"(w_e*{k}+(1-w_e)*kd_at+wacc_adj)"))
        if spec == "w03": return tp({}, w_dcf + "_lo" if w_dcf_lohi else "0.3")
        if spec == "w07": return tp({}, w_dcf + "_hi" if w_dcf_lohi else "0.7")
        return tp(spec)
    cases = list(cases)
    if beta_u:
        bd = beta_du if isinstance(beta_du, str) else repr(beta_du)
        bt = lambda sg: f"({beta_u}{sg}{bd})*{relever}"
        cases = [(nm, (f'="β_u "&TEXT({beta_u},"0.000")&" ∓ "&TEXT({bd},"0.0")&" 各自重新加槓桿：β "&TEXT({bt("-")},"0.00")&"／"&TEXT({bt("+")},"0.00")&"（基準 "&TEXT(beta,"0.00")&"）"'
                  if key == "beta" else st), lo, hi, key) for nm, st, lo, hi, key in cases]
    if usd:
        i_ = next((i for i, c in enumerate(cases) if c[4] == "tp_w_dcf"), len(cases))
        cases.insert(i_, ("美元口徑折現", '="rf＝美國 10 年期 "&TEXT(rf_us,"0.00%")&"（基準當地 "&TEXT(rf,"0.00%")&"）＋ β ×（ERP＋CRP）、g＝MIN("&TEXT(g_cap,"0%")&", "&TEXT(rf_us,"0.00%")&")；其他不變（單一案例，列在低欄）"',
                          "usd", None, "rf_us"))
    r = r0 + 2; first = r
    for nm, st, lo, hi, key in cases:
        SE.cell(r, 1, nm); SE.cell(r, 2, st); SE.cell(r, 8, key)
        for col, spec in ((3, lo), (5, hi)):
            f = resolve(spec)
            if f is None:
                SE.cell(r, col, "—"); SE.cell(r, col + 1, "—")
            else:
                SE.cell(r, col, f).number_format = fmt
                SE.cell(r, col + 1, f"={get_column_letter(col)}{r}/G{r}-1").number_format = "+0.0%;-0.0%;0.0%"
        SE.cell(r, 7, f"={base_ref}").number_format = fmt
        r += 1
    last = r - 1
    SE.cell(r, 1, "檢查：閉式公式的基準 − 基準加權目標價（應為 0；閉式與 DCF 頁明細同式）")
    SE.cell(r, 3, tp({}) + f"-{base_ref}").number_format = "0.000000"
    return first, last, r + 1


# ---------------- 5. 檢查列（規格第 4 節） ----------------
def write_checks(ws, r0, px="px", nc_bs=None, nc_note=""):
    """ws 第 r0 列起：WACC 由 In 頁輸入重算 − Val 的 WACC（應為 0）；nc_bs（資產負債表各科目的公式字串，不含「=」）給定時，
    另寫「淨現金與投資 − 資產負債表」（應為 0）。回傳下一列。"""
    ws.cell(r0, 1, "檢查列（docs/plan/個股DCF標準規格.md 第 4 節）").font = BOLD
    r = r0 + 1
    ws.cell(r, 1, "檢查：WACC 由 In 頁輸入重算 − Val 的 WACC（應為 0）")
    ws.cell(r, 3, f"=({px}*sh_v*(rf+beta*(erp+crp))+(debt_v+lease_v)*kd*(1-tax_w))/({px}*sh_v+debt_v+lease_v)+wacc_adj-wacc").number_format = "0.000000"
    r += 1
    if nc_bs:
        ws.cell(r, 1, "檢查：淨現金與投資 − 資產負債表科目（應為 0）"); ws.cell(r, 2, nc_note)
        ws.cell(r, 3, f"=dcf_nc-({nc_bs})").number_format = "0.000000"
        r += 1
    return r
