"""v0.6.1（工作單 20261022）：PxQ 頁 — P×Q 與「實際 m×L ÷ 所需 m×L」回本比值序列。

規格來源：docs/plan/20261022_OpenAI_PxQ回本比值_交接.md 第 2 節（Andy 2026-10-21～22 原話，權威）與工作單第 3 節。
- 每 GW 參數只取 TK_PxQ（Tokenomics v19.0 快照，具名範圍 TKQ_）；公司層級（GW、世代組合、營收）取本模型的 INP_／SRC_OAI_／CMP_／REV_。
- 公式只引用 TKQ_、INP_、SRC_OAI_、CMP_、REV_ 與本頁名稱／儲存格；公式內不含常數（E6；恆等式 1−x、年數含頭尾的 +1 除外）。
- Python 只產生結構；計算由 Excel（LibreOffice）與 engine 執行。比值＝實際 m×L ÷ 所需 m×L＝營收 ÷ 算力成本（Q 與牌價約掉）。
- 欄 D–H＝五個案例：2025 指數內插、2025 線性內插、2026 低／基準／高 GW。
"""
from __future__ import annotations

from openpyxl.styles import Alignment

from common import F_BOLD, F_CALC, F_IN, F_NOTE, FILL_HDR, FILL_SEC, header, put, title

CC = ["D", "E", "F", "G", "H"]
CASES = ["2025 指數內插", "2025 線性內插", "2026 低 GW", "2026 基準 GW", "2026 高 GW"]
YI = [0, 0, 1, 1, 1]                                   # 0＝2025、1＝2026
GENS = [("Hopper", "Hopper"), ("GB200", "GB200"), ("GB300", "GB300"), ("VR200", "VR200"), ("RubinUltra", "Rubin Ultra"), ("Custom", "自研／其他")]
PERC, NUM2, NUM3, NUM1 = "0.0%", "0.00", "0.000", "0.0"
SCI = "0.00E+00"


class PxQ:
    def __init__(self, wb, nm, R, S, I, snap):
        self.wb, self.nm = wb, nm
        self.ws = wb.create_sheet("PxQ")
        self.S, self.I, self.snap = S, I, snap
        self.rows = {}
        self.r = 16
        self.n = 0
        ws = self.ws
        title(ws, "PxQ — P×Q 與回本比值：實際 m×L ÷ 所需 m×L（＝營收 ÷ 算力成本）的年度序列（v0.6.1）",
              "每 GW 參數只取 TK_PxQ（Tokenomics v19.0 快照，TKQ_）；公司層級（平均上線 GW、世代組合、營收）取本模型。看的不是回本的絕對數字，而是實際 ÷ 所需的比值趨勢（Andy 2026-10-22）。",
              "m＝付費服務 token 的實收單價 ÷ 有效牌價（有效牌價＝表列 ×（1 − d），d＝TKQ_IF_Disc）；L 取壽命期價格係數（判斷題 1 預設）。比值中 Q 與牌價約掉，Q 只決定 m×L 的水準。"
              "欄 D–H＝2025 指數／線性內插、2026 低／基準／高 GW。$B＝十億美元；Q 單位 M tok（百萬 token）／年。")
        header(ws, 5, ["編號", "項目", "單位"] + CASES + ["", "說明"])
        for c, w in (("A", 7), ("B", 58), ("C", 14), ("D", 15), ("E", 15), ("F", 15), ("G", 15), ("H", 15), ("I", 12), ("J", 90)):
            ws.column_dimensions[c].width = w
        ws.freeze_panes = "D6"

    # ── 基本寫入 ─────────────────────────────────────────
    def section(self, label):
        put(self.ws, f"A{self.r}", label, F_BOLD, fill=FILL_SEC)
        self.r += 1

    def add(self, label, unit, fs, key, note="", fmt=None, name=None, font=None):
        """fs：長度 5 的清單（None＝空白）或 callable(i, col)。"""
        self.n += 1
        r = self.r
        put(self.ws, f"A{r}", f"X{self.n:02d}", F_CALC)
        put(self.ws, f"B{r}", label, F_CALC)
        put(self.ws, f"C{r}", unit, F_CALC)
        for i, c in enumerate(CC):
            f = fs(i, c) if callable(fs) else fs[i]
            if f is not None:
                put(self.ws, f"{c}{r}", f, font, fmt=fmt)
        put(self.ws, f"J{r}", note, F_NOTE)
        self.rows[key] = r
        if name:
            self.nm(self.wb, name, f"PxQ!$D${r}:$H${r}")
        self.r += 1
        return r

    def hdr(self, labels, key):
        r = self.r
        for i, lab in enumerate(labels):
            c = self.ws.cell(row=r, column=2 + i, value=lab)
            c.font, c.fill, c.alignment = F_BOLD, FILL_HDR, Alignment(wrap_text=True, vertical="top")
        self.rows[key] = r
        self.r += 1


def build_p6(wb, R, S, I, nm, snapq):
    X = PxQ(wb, nm, R, S, I, snapq)
    ws = X.ws
    # ── 具名輸入 ──────────────────────────────────────────
    sw25, gw_mid, gw_lo, gw_hi = I("2025 平均上線 GW 主值選擇"), I("2026 平均上線 GW：基準"), I("2026 平均上線 GW：低情境"), I("2026 平均上線 GW：高情境")
    mixsw, dsw, tol = I("世代組合來源"), I("生產折減讀法"), I("預覽值對帳容許相對誤差")
    hf = {k: I("交接檔 2026 世代組合：" + zh) for k, zh in GENS}
    custom, usd, mt = I("世代產能係數：自研／其他"), I("單位換算：$ → $B 的除數"), I("單位換算：T token → M token")
    y25, y26 = "INP_005", "INP_006"
    s_rev25, s_rev26 = S("營收：FY2025"), S("管理層營收目標：2026")
    s_ep16, s_bro50 = S("算力（雲端）支出：2025（媒體報導，經 Epoch 引用）"), S("算力支出：2026（Brockman 證詞）")
    s_serve50 = S("對外服務占算力比：2025（Epoch：研發與推論約各半）")

    # ═══ 三、案例定義（先建；頂部主輸出最後填）═══
    X.section("一、案例定義：年度、平均上線 GW（欄 D–H）、營收（基準）")
    X.add("年度", "年", lambda i, c: f"={y25}" if YI[i] == 0 else f"={y26}", "yr", "INP_005／INP_006（2025 為實際校準年）", name="PXQ_Year")
    X.add("年度序位（年度 − 起始年 + 1；供 INDEX 取 CMP_／REV_ 年度向量）", "序位", lambda i, c: f"={c}{X.rows['yr']}-{y25}+1", "idx", "年數含頭尾的 +1（恆等式）")
    X.add("平均上線 GW", "GW", lambda i, c: [
        "=(TKQ_SRC_DEM_009-TKQ_SRC_DEM_008)/LN(TKQ_SRC_DEM_009/TKQ_SRC_DEM_008)", "=AVERAGE(TKQ_SRC_DEM_008,TKQ_SRC_DEM_009)",
        f"={gw_lo}", f"={gw_mid}", f"={gw_hi}"][i], "gw",
          "2025：OpenAI 官方端點（2024 年底 0.6、2025 年底 1.9；Interested-party；電力口徑未明）內插——指數＝(端點差)÷ln(端點比)，線性＝端點平均（Derived）；"
          "2026：Inputs 三情境（Assumed；官方未公布）", fmt=NUM3, name="PXQ_GW")
    X.add("營收（基準）", "$B", lambda i, c: f"={s_rev25}" if YI[i] == 0 else f"={s_rev26}", "rev",
          "2025＝SRC_OAI（13.07，Checks 驗等於 TKQ_SRC_DEM_007）；2026＝管理層目標 36（FT 2026-09-18 七月投資人簡報；利害關係方的管理層目標，非市場傳聞）", fmt=NUM2, name="PXQ_Rev")

    # ═══ 二、世代參數表（TKQ）═══
    X.section("二、世代參數（每 GW；TK_PxQ 基準成本情境）：欄 D 年持有成本、E 每 GW 付費年產能、F 生產折減（v19.0 讀法）、G 生產折減（一律 IF_ProdDerate）、H 採用（Inputs 開關）")
    X.hdr(["世代", "TKQ_HdrGen 對照名稱", "持有成本 $B/GW/年", "付費年產能 M tok/GW/年", "折減（v19.0 讀法）", "折減（一律 ProdDerate）", "折減（採用）"], "g_hdr")
    r_cost = X.r
    put(ws, f"B{r_cost}", "成本情境（對應 TKQ_HdrCost；基準）", F_CALC)
    put(ws, f"D{r_cost}", snapq["hdr_cost"][1], F_CALC)
    X.rows["costkey"] = r_cost
    X.r += 1
    g = {}
    for k, (key, zh) in enumerate(GENS):
        r = X.r
        g[key] = r
        put(ws, f"A{r}", f"X{X.n + 1:02d}", F_CALC)
        X.n += 1
        put(ws, f"B{r}", zh, F_CALC)
        if key != "Custom":
            put(ws, f"C{r}", snapq["hdr_gen"][3 * k], F_CALC)
            put(ws, f"D{r}", f"=SUMIFS(TKQ_IF_HoldEcon,TKQ_HdrGen,$C{r},TKQ_HdrCost,$D${r_cost})", fmt=NUM3)
            put(ws, f"E{r}", f"=SUMIFS(TKQ_IF_TokGWPaidMix,TKQ_HdrGen,$C{r},TKQ_HdrCost,$D${r_cost})", fmt=SCI)
        else:
            put(ws, f"C{r}", "（VR200 × 世代產能係數）", F_CALC)
            put(ws, f"D{r}", f"=D{g['VR200']}*{custom}", fmt=NUM3)
            put(ws, f"E{r}", f"=E{g['VR200']}*{custom}", fmt=SCI)
        put(ws, f"F{r}", "=TKQ_IF_FleetProdRatio_VR200" if key == "VR200" else "=TKQ_IF_ProdDerate", fmt=NUM3)
        put(ws, f"G{r}", "=TKQ_IF_ProdDerate", fmt=NUM3)
        put(ws, f"H{r}", f"=(1-{dsw})*F{r}+{dsw}*G{r}", fmt=NUM3)
        put(ws, f"J{r}", {"Custom": "自研晶片：持有成本與產能＝VR200 × 世代產能係數（沿用 v0.6 P4／P3 做法；Analogy，係數區間 0.4–1.0）；折減用 IF_ProdDerate（非 VR200）",
                          "VR200": "v19.0 讀法：VR200 用 IF_FleetProdRatio_VR200（重解 SLO 後的比）；其他世代用 IF_ProdDerate",
                          "RubinUltra": "占比 >0 時才影響（基準 2025、2026 占比為 0）"}.get(key, ""), F_NOTE)
        X.r += 1
    g1, g6 = g["Hopper"], g["Custom"]
    nm(wb, "PXQ_HoldPerGW", f"PxQ!$D${g1}:$D${g6}")
    nm(wb, "PXQ_PaidCapPerGW", f"PxQ!$E${g1}:$E${g6}")

    # ═══ 通用案例區塊 ═══
    def sh_expr(mix, key, i, c):
        base = f"INDEX(CMP_Share_{key},{c}${X.rows['idx']})"
        if YI[i] == 0 or mix == "v06":
            return f"={base}"
        if mix == "handoff":
            return f"={hf[key]}"
        return f"=(1-{mixsw})*{base}+{mixsw}*{hf[key]}"

    def block(tag, title_, mix, dcol, serve, rev_fn=None, gw_fn=None, full=True, prefix=False):
        """回傳 rows 字典。mix：main／v06／handoff；dcol：折減欄（H 採用、F v19.0、G 一律）；serve：服務占比公式片段。"""
        X.section(title_)
        rr = {}
        gwr = X.add("平均上線 GW", "GW", lambda i, c: gw_fn(i, c) if gw_fn else f"={c}{X.rows['gw']}", f"{tag}_gw", "", fmt=NUM3)
        for key, zh in GENS:
            rr[key] = X.add(f"世代占比：{zh}", "比例", lambda i, c, key=key: sh_expr(mix, key, i, c), f"{tag}_sh_{key}", "", fmt=PERC)
        s1, s6 = rr["Hopper"], rr["Custom"]
        X.add("世代占比合計（應＝1）", "比例", lambda i, c: f"=SUM({c}{s1}:{c}{s6})", f"{tag}_shsum", "Checks 驗證", fmt=PERC)
        X.add("每 GW 年持有成本 h（世代加權）", "$B/GW/年", lambda i, c: f"=SUMPRODUCT({c}{s1}:{c}{s6},$D${g1}:$D${g6})", f"{tag}_h", "h＝Σ 世代占比 × TKQ_IF_HoldEcon（基準欄）", fmt=NUM3, name=f"PXQ_h" if tag == "m" else None)
        X.add("算力成本 C＝平均 GW × h", "$B", lambda i, c: f"={c}{X.rows[f'{tag}_gw']}*{c}{X.rows[f'{tag}_h']}", f"{tag}_C", "年度算力成本（機隊持有成本口徑；Andy：『從持有的機隊成本去推算』）", fmt=NUM2, name="PXQ_C" if tag == "m" else None)
        X.add("每 GW 付費年產能（世代加權，已乘生產折減）", "M tok/GW/年",
              lambda i, c: f"=SUMPRODUCT({c}{s1}:{c}{s6},$E${g1}:$E${g6},${dcol}${g1}:${dcol}${g6})", f"{tag}_cap", "Σ 世代占比 × TKQ_IF_TokGWPaidMix × 生產折減", fmt=SCI)
        X.add("Q（M tok／年）＝GW × 服務占比 ×（1 − 免費占比）× 每 GW 付費年產能 × 實際利用率", "M tok/年",
              lambda i, c: f"={c}{X.rows[f'{tag}_gw']}*{serve}*(1-TKQ_IF_FreeShare)*{c}{X.rows[f'{tag}_cap']}*TKQ_IF_UtilActual", f"{tag}_Q",
              "Andy：『Q 就是真正拿去用作服務的算力，去乘上每單位算力可以產生的 token 數』", fmt=SCI, name="PXQ_Q" if tag == "m" else None)
        X.add("Q（T token／年）", "T tok/年", lambda i, c: f"={c}{X.rows[f'{tag}_Q']}/{mt}", f"{tag}_QT", "顯示用（交接檔『約 6.4×10¹⁵ token』＝6,400 T）", fmt="#,##0")
        X.add("有效牌價 P_eff（付費層級組合；已含折扣 d）", "$/M tok", lambda i, c: "=TKQ_IF_PxEffMix25" if YI[i] == 0 else "=TKQ_IF_PxEffMix", f"{tag}_P",
              "2025＝IF_PxEffMix25（年度加權）；2026＝IF_PxEffMix（2026 快照）", fmt=NUM3, name="PXQ_Peff" if tag == "m" else None)
        if full:
            X.add("表列牌價＝P_eff ÷（1 − d）", "$/M tok", lambda i, c: f"={c}{X.rows[f'{tag}_P']}/(1-TKQ_IF_Disc)", f"{tag}_Plist", "d＝TKQ_IF_Disc，不寫死", fmt=NUM3)
        X.add("營收", "$B", lambda i, c: rev_fn(i, c) if rev_fn else f"={c}{X.rows['rev']}", f"{tag}_rev", "", fmt=NUM2)
        X.add("實際 m×L＝營收 ÷（Q × P_eff）", "倍", lambda i, c: f"={c}{X.rows[f'{tag}_rev']}*{usd}/({c}{X.rows[f'{tag}_Q']}*{c}{X.rows[f'{tag}_P']})", f"{tag}_act",
              "m 只看收費業務、跟營收有關跟成本無關", fmt=NUM3, name="PXQ_mL_Actual" if tag == "m" else None)
        X.add("所需 m×L＝算力成本 ÷（Q × P_eff）", "倍", lambda i, c: f"={c}{X.rows[f'{tag}_C']}*{usd}/({c}{X.rows[f'{tag}_Q']}*{c}{X.rows[f'{tag}_P']})", f"{tag}_req",
              "回本所需", fmt=NUM3, name="PXQ_mL_Required" if tag == "m" else None)
        X.add("比值＝實際 m×L ÷ 所需 m×L", "倍", lambda i, c: f"={c}{X.rows[f'{tag}_act']}/{c}{X.rows[f'{tag}_req']}", f"{tag}_ratio", "主輸出；＝營收 ÷ 算力成本", fmt=NUM3, name="PXQ_Ratio" if tag == "m" else None)
        if full:
            X.add("恆等式：比值 −（營收 ÷ 算力成本）（應＝0）", "倍", lambda i, c: f"={c}{X.rows[f'{tag}_ratio']}-{c}{X.rows[f'{tag}_rev']}/{c}{X.rows[f'{tag}_C']}", f"{tag}_ident", "Checks 驗證", fmt="0.0E+00", name="PXQ_Identity" if tag == "m" else None)
            X.add("隱含 m＝實際 m×L ÷ L（L＝2025 機隊，TKQ_IF_LTotalFleetVDisc）", "倍", lambda i, c: f"={c}{X.rows[f'{tag}_act']}/TKQ_IF_LTotalFleetVDisc", f"{tag}_mfleet",
                  "只作顯示（判斷題 1：L 讀法預設＝壽命期價格係數）", fmt=NUM3)
            X.add("隱含 m＝實際 m×L ÷ L（L＝新世代，TKQ_IF_LTotalDisc）", "倍", lambda i, c: f"={c}{X.rows[f'{tag}_act']}/TKQ_IF_LTotalDisc", f"{tag}_mnew", "只作顯示", fmt=NUM3)
        rg = f"{tag}_ratio"
        sel = f"((1-{sw25})*$D{X.rows[rg]}+{sw25}*$E{X.rows[rg]})"
        for lab, ref, key in (("持平所需 2026 平均 GW（以 2025 指數比值）", f"$D{X.rows[rg]}", "be_exp"), ("持平所需 2026 平均 GW（以 2025 線性比值）", f"$E{X.rows[rg]}", "be_lin"),
                              ("持平所需 2026 平均 GW（以 2025 主值比值）", sel, "be_sel")):
            if not full and key != "be_sel" and tag not in ("pv",):
                continue
            X.add(lab + "＝2026 營收 ÷（2026 h × 2025 比值）", "GW",
                  lambda i, c, ref=ref: None if YI[i] == 0 else f"={c}{X.rows[f'{tag}_rev']}/({c}{X.rows[f'{tag}_h']}*{ref})", f"{tag}_{key}",
                  "2026 平均上線 GW 低於此值，比值才高於 2025" if key == "be_sel" else "", fmt=NUM3,
                  name={"be_exp": "PXQ_BE_Exp", "be_lin": "PXQ_BE_Lin", "be_sel": "PXQ_BE_Sel"}[key] if tag == "m" else None)
        return rr

    serve_base = "TKQ_IF_ServeShare"
    X.section_marker = None
    block("m", "三、主序列（基準）：世代組合依 Inputs 開關（基準＝OpenAI v0.6 自有 V1 組合）、生產折減依 Inputs 開關（基準＝v19.0 讀法）、服務占比＝Tokenomics",
          "main", "H", serve_base)

    # ═══ 四、敏感度 ═══
    X.section("四、敏感度與對照基準（下列各區塊與主序列只差一個條件；比值不受服務占比影響）")
    block("pv", "4.1 預覽值基準：交接檔世代組合（2026）＋一律 IF_ProdDerate（對帳用；其餘同基準）", "handoff", "G", serve_base)
    block("sv", "4.2 服務占比改為 Epoch 對照值（SRC_OAI『研發與推論約各半』）：Q 與 m×L 改變，比值不變", "main", "H", s_serve50, full=False)
    block("rv", "4.3 2026 營收改用模型正向（REV_GrossCapped 2026；截頂後總額）", "main", "H", serve_base, full=False,
          rev_fn=lambda i, c: f"={s_rev25}" if YI[i] == 0 else f"=INDEX(REV_GrossCapped,{c}${X.rows['idx']})")
    block("mx", "4.4 世代組合改用交接檔（2026；v19.0 生產折減讀法不變）", "handoff", "H", serve_base, full=False)
    block("su", "4.5 2026 平均上線 GW 改用 v0.6 供給 GW（CMP_SupplyGW 2026；合約＋自建；是否年平均未明）", "main", "H", serve_base, full=False,
          gw_fn=lambda i, c: f"={c}{X.rows['gw']}" if YI[i] == 0 else f"=INDEX(CMP_SupplyGW,{c}${X.rows['idx']})")

    # ═══ 五、對照列 ═══
    X.section("五、對照列（外部報導的算力支出 ÷ 持有成本 GW × h；不進主序列）")
    X.add("外部報導算力支出（2025：雲端，Epoch 引媒體；2026：Brockman 證詞）", "$B", lambda i, c: f"={s_ep16}" if YI[i] == 0 else f"={s_bro50}", "x_ext",
          "SRC_OAI（新列；出處網址見該列）", fmt=NUM2, name="PXQ_ExtSpend")
    X.add("倍數＝外部支出 ÷（GW × h）", "倍", lambda i, c: f"={c}{X.rows['x_ext']}/{c}{X.rows['m_C']}", "x_mult", "固定倍數在比值趨勢中約掉，但須逐年用同一口徑", fmt=NUM2, name="PXQ_ExtMultiple")
    X.add("2025 推論＋訓練支出（TKQ_SRC_DEM_004＋006）", "$B", lambda i, c: "=TKQ_SRC_DEM_004+TKQ_SRC_DEM_006" if YI[i] == 0 else None, "x_it",
          "Andy：『兩百零四億，很明顯是高估了』——只作對照，不作算力成本", fmt=NUM2)
    X.add("倍數＝（推論＋訓練）÷（GW × h）", "倍", lambda i, c: f"={c}{X.rows['x_it']}/{c}{X.rows['m_C']}" if YI[i] == 0 else None, "x_it_mult", "", fmt=NUM2)

    # ═══ 六、預覽值對帳 ═══
    X.section("六、預覽值對帳（交接檔第 6 節；以 4.1 預覽值基準計；基準欄＝主序列）")
    X.hdr(["項目", "單位", "交接檔值", "模型（預覽基準）", "相對差", "判定", "模型（主序列）", "主序列相對差"], "rc_hdr")
    pv = lambda k, c: f"{c}{X.rows['pv_' + k]}"            # noqa: E731
    mm = lambda k, c: f"{c}{X.rows['m_' + k]}"             # noqa: E731
    items = [
        ("每 GW 年持有成本 2025（93.5 億）", "$B", 9.35, "h", "D"), ("每 GW 年持有成本 2026（97.3 億）", "$B", 9.73, "h", "G"),
        ("算力成本 2025 指數（106 億）", "$B", 10.6, "C", "D"), ("算力成本 2025 線性（117 億）", "$B", 11.7, "C", "E"),
        ("算力成本 2026 低 GW（243 億）", "$B", 24.3, "C", "F"), ("算力成本 2026 基準 GW（292 億）", "$B", 29.2, "C", "G"), ("算力成本 2026 高 GW（341 億）", "$B", 34.1, "C", "H"),
        ("比值 2025 指數", "倍", 1.24, "ratio", "D"), ("比值 2025 線性", "倍", 1.12, "ratio", "E"),
        ("比值 2026 低 GW", "倍", 1.48, "ratio", "F"), ("比值 2026 基準 GW", "倍", 1.23, "ratio", "G"), ("比值 2026 高 GW", "倍", 1.06, "ratio", "H"),
        ("持平所需 GW（2025 指數比值）", "GW", 2.99, "be_exp", "G"), ("持平所需 GW（2025 線性比值）", "GW", 3.31, "be_lin", "G"),
        ("Q 2025 線性 1.25 GW（約 6.4×10¹⁵ token）", "T tok", 6400, "QT", "E"),
    ]
    rc_rows = []
    for lab, unit, hv, k, c in items:
        r = X.r
        X.n += 1
        put(ws, f"A{r}", f"X{X.n:02d}", F_CALC)
        put(ws, f"B{r}", lab, F_CALC)
        put(ws, f"C{r}", unit, F_CALC)
        put(ws, f"D{r}", hv, F_NOTE)
        put(ws, f"E{r}", f"={pv(k, c)}", fmt=NUM3 if hv < 100 else "#,##0")
        put(ws, f"F{r}", f"=(E{r}-D{r})/D{r}", fmt="0.00%")
        put(ws, f"G{r}", f'=IF(ABS(F{r})<={tol},"OK","差異")')
        put(ws, f"H{r}", f"={mm(k, c)}", fmt=NUM3 if hv < 100 else "#,##0")
        put(ws, f"I{r}", f"=(H{r}-D{r})/D{r}", fmt="0.00%")
        rc_rows.append(r)
        X.r += 1
    X.rows["rc_first"], X.rows["rc_last"] = rc_rows[0], rc_rows[-1]
    put(ws, f"B{X.r}", "對帳超出容許誤差（預覽基準）的項數", F_BOLD)
    put(ws, f"E{X.r}", f'=COUNTIF(G{rc_rows[0]}:G{rc_rows[-1]},"差異")')
    X.rows["rc_bad"] = X.r
    nm(wb, "PXQ_ReconBad", f"PxQ!$E${X.r}")
    put(ws, f"J{X.r}", "交接檔值為 chat 端以 Tokenomics v18.0 計得的預覽（金額由『億』換算為 $B 的常數，只作對照）；本表模型值取 v19.0。差異逐項說明見報告", F_NOTE)
    X.r += 1

    # ═══ 頂部主輸出 ═══
    r = 7
    put(ws, f"A{r}", "主輸出：回本比值（實際 m×L ÷ 所需 m×L ＝ 營收 ÷ 算力成本）的年度序列", F_BOLD, fill=FILL_SEC)
    for i, c in enumerate(CC):
        put(ws, f"{c}8", CASES[i], F_BOLD, fill=FILL_HDR)
    put(ws, "B8", "案例", F_BOLD, fill=FILL_HDR)
    put(ws, "B9", "年度", F_CALC)
    put(ws, "B10", "平均上線 GW", F_CALC)
    put(ws, "B11", "比值（實際 ÷ 所需）", F_BOLD)
    put(ws, "B12", "實際 m×L", F_CALC)
    put(ws, "B13", "所需 m×L", F_CALC)
    put(ws, "B14", "2025 主值比值（Inputs 開關：指數／線性）", F_CALC)
    put(ws, "B15", "讀法", F_BOLD)
    for c in CC:
        put(ws, f"{c}9", f"={c}{X.rows['yr']}", fmt="0")
        put(ws, f"{c}10", f"={c}{X.rows['m_gw']}", fmt=NUM3)
        put(ws, f"{c}11", f"={c}{X.rows['m_ratio']}", F_BOLD, fmt=NUM2)
        put(ws, f"{c}12", f"={c}{X.rows['m_act']}", fmt=NUM3)
        put(ws, f"{c}13", f"={c}{X.rows['m_req']}", fmt=NUM3)
    put(ws, "D14", f"=(1-{sw25})*D11+{sw25}*E11", fmt=NUM2)
    nm(wb, "PXQ_Ratio2025Sel", "PxQ!$D$14")
    put(ws, "D15",
        '="2026 年平均上線 "&TEXT(G10,"0.0")&" GW 時，營收 ÷ 算力成本為 "&TEXT(G11,"0.00")&"，"&IF(G11>=D14,"不低於 ","低於 ")&"2025 年主值（"&TEXT(D14,"0.00")&"）；'
        '2026 平均上線 GW 低於約 "&TEXT(G' + str(X.rows["m_be_sel"]) + ',"0.00")&" GW 時比值才高於 2025。"')
    nm(wb, "PXQ_Reading", "PxQ!$D$15")
    put(ws, "J11", "看的並不是回本的絕對數字，而是實際除以所需這個比例的趨勢（Andy 2026-10-22）；比值＞1＝實收足以覆蓋機隊持有成本", F_NOTE)
    put(ws, "J15", "讀法句為公式文字（數字取自本頁儲存格，非手寫）", F_NOTE)
    return X
