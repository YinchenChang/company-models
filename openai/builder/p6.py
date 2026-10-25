"""v0.6.1（工作單 20261022）建立、v0.6.3（工作單 20261025）改寫：PxQ 頁 — P×Q 與「實際 m×L ÷ 所需 m×L」回本比值序列。

規格來源：docs/plan/20261022_OpenAI_PxQ回本比值_交接.md 第 2 節（Andy 2026-10-21～22 原話，權威）與 v0.6.1 工作單第 3 節；
v0.6.3 工作單 docs/workorders/20261025_openai_v0.6.3_PxQ比值2026三口徑.md（2026 改為三口徑、Fairwater 對照、GW 路線 token 帳）。
- 每 GW 參數只取 TK_Link（Tokenomics 單一快照，具名範圍 TK_）；公司層級（GW、世代組合、營收）取本模型的 INP_／SRC_OAI_／CMP_／REV_。
- 公式只引用 TK_、INP_、SRC_OAI_、CMP_、REV_ 與本頁名稱／儲存格；公式內不含常數（E6；恆等式 1−x、年數含頭尾的 +1 除外）。
- Python 只產生結構；計算由 Excel（LibreOffice）與 engine 執行。比值＝實際 m×L ÷ 所需 m×L＝營收 ÷ 算力成本（Q 與牌價約掉）。
- 欄 D–N＝11 個案例：2025 指數內插、2025 線性內插；口徑一（Q2 年化營收 ÷ Q2 季平均 GW）低／基準／高；
  口徑二（FY2026 營收 ÷ 全年平均 GW）低／基準／高；口徑三（9 月底 run-rate ÷ 年底 GW）低／基準／高。
"""
from __future__ import annotations

from openpyxl.styles import Alignment

from common import F_BOLD, F_CALC, F_NOTE, FILL_HDR, FILL_SEC, header, put, title

CC = list("DEFGHIJKLMN")
SCOPE = [0, 0, 1, 1, 1, 2, 2, 2, 3, 3, 3]              # 0＝2025；1＝口徑一；2＝口徑二；3＝口徑三
LEVEL = [0, 0, 0, 1, 2, 0, 1, 2, 0, 1, 2]              # 0 低／1 基準／2 高
YI = [0 if s == 0 else 1 for s in SCOPE]               # 0＝2025、1＝2026
CASES = ["2025 指數內插", "2025 線性內插"] + [f"口徑{z} {lv}" for z in "一二三" for lv in ("低", "基準", "高")]
SC_COLS = {1: ("F", "G", "H"), 2: ("I", "J", "K"), 3: ("L", "M", "N")}     # 各口徑 低／基準／高 的欄
SC_NAME = {1: "口徑一｜Q2 年化營收 ÷ Q2 季平均 GW", 2: "口徑二｜FY2026 營收 ÷ 全年平均 GW", 3: "口徑三｜9 月底 run-rate ÷ 年底 GW"}
GENS = [("Hopper", "Hopper"), ("GB200", "GB200"), ("GB300", "GB300"), ("VR200", "VR200"), ("RubinUltra", "Rubin Ultra"), ("Custom", "自研／其他")]
PERC, NUM2, NUM3 = "0.0%", "0.00", "0.000"
SCI = "0.00E+00"
OLD_COL = {"D": "D", "E": "E", "F": "I", "G": "J", "H": "K"}   # v18.0 預覽值的 2026 低／基準／高（2.5／3.0／3.5 GW）→ 口徑二的欄
MASK_Y25_S2 = [True, True, False, False, False, True, True, True, False, False, False]
MASK_Y25_S1 = [True, True, True, True, True, False, False, False, False, False, False]


class PxQ:
    def __init__(self, wb, nm, R, S, I, snap):
        self.wb, self.nm = wb, nm
        self.ws = wb.create_sheet("PxQ")
        self.S, self.I, self.snap = S, I, snap
        self.rows = {}
        self.r = 31
        self.n = 0
        self.mask = None
        self.deferred = []
        ws = self.ws
        title(ws, "PxQ — P×Q 與回本比值：實際 m×L ÷ 所需 m×L（＝營收 ÷ 算力成本）的年度序列（v0.6.3：2026 三口徑）",
              "每 GW 參數只取 TK_Link（Tokenomics 單一快照，TK_）；公司層級（平均上線 GW、世代組合、營收）取本模型。看的不是回本的絕對數字，而是實際 ÷ 所需的比值趨勢（Andy 2026-10-22）。",
              "m＝付費服務 token 的實收單價 ÷ 有效牌價（有效牌價＝表列 ×（1 − d），d＝TK_IF_Disc）；L 取壽命期價格係數（判斷題 1 預設）。比值中 Q 與牌價約掉，Q 只決定 m×L 的水準。"
              "欄 D–N＝2025 指數／線性內插 + 2026 三口徑 × 低／基準／高（口徑一：Q2 年化營收 ÷ Q2 季平均 GW；口徑二：FY2026 營收 ÷ 全年平均 GW；口徑三：9 月底 run-rate ÷ 年底 GW）。$B＝十億美元；Q 單位 M tok（百萬 token）／年。")
        header(ws, 5, ["編號", "項目", "單位"] + CASES + ["", "說明"])
        for c, w in (("A", 7), ("B", 58), ("C", 14)):
            ws.column_dimensions[c].width = w
        for c in CC:
            ws.column_dimensions[c].width = 15
        ws.column_dimensions["O"].width = 4
        ws.column_dimensions["P"].width = 90
        ws.freeze_panes = "D6"

    # ── 基本寫入 ─────────────────────────────────────────
    def section(self, label):
        put(self.ws, f"A{self.r}", label, F_BOLD, fill=FILL_SEC)
        self.r += 1

    def add(self, label, unit, fs, key, note="", fmt=None, name=None, font=None):
        """fs：長度 11 的清單（None＝空白）或 callable(i, col)。self.mask 為 False 的欄略過。"""
        self.n += 1
        r = self.r
        put(self.ws, f"A{r}", f"X{self.n:02d}", F_CALC)
        put(self.ws, f"B{r}", label, F_CALC)
        put(self.ws, f"C{r}", unit, F_CALC)
        for i, c in enumerate(CC):
            if self.mask is not None and not self.mask[i]:
                continue
            f = fs(i, c) if callable(fs) else fs[i]
            if f is not None:
                put(self.ws, f"{c}{r}", f, font, fmt=fmt)
        put(self.ws, f"P{r}", note, F_NOTE)
        self.rows[key] = r
        if name:
            self.nm(self.wb, name, f"PxQ!$D${r}:$N${r}")
        self.r += 1
        return r

    def hdr(self, labels, key):
        r = self.r
        for i, lab in enumerate(labels):
            c = self.ws.cell(row=r, column=2 + i, value=lab)
            c.font, c.fill, c.alignment = F_BOLD, FILL_HDR, Alignment(wrap_text=True, vertical="top")
        self.rows[key] = r
        self.r += 1

    def line(self, label, key, cells, note="", fmt_map=None, unit=""):
        """自由格式的一列：cells＝{欄字母: 值或公式}；公式含 {key} 者待 resolve() 時以列號取代（可參照後面才建立的列）。"""
        self.n += 1
        r = self.r
        put(self.ws, f"A{r}", f"X{self.n:02d}", F_CALC)
        put(self.ws, f"B{r}", label, F_CALC)
        put(self.ws, f"C{r}", unit, F_CALC)
        self.rows[key] = r
        for col, v in cells.items():
            fmt = (fmt_map or {}).get(col)
            if isinstance(v, str) and v.startswith("=") and "{" in v:
                self.deferred.append((f"{col}{r}", v, fmt))
            else:
                put(self.ws, f"{col}{r}", v, fmt=fmt)
        put(self.ws, f"P{r}", note, F_NOTE)
        self.r += 1
        return r

    def resolve(self):
        for ref, tpl, fmt in self.deferred:
            put(self.ws, ref, tpl.format_map(self.rows), fmt=fmt)
        self.deferred = []


def build_p6(wb, R, S, I, nm, snapq):
    X = PxQ(wb, nm, R, S, I, snapq)
    ws = X.ws
    # ── 具名輸入 ──────────────────────────────────────────
    sw25, gw_mid, gw_lo, gw_hi = I("2025 平均上線 GW 主值選擇"), I("口徑二：2026 全年平均上線 GW：基準"), I("口徑二：2026 全年平均上線 GW：低情境"), I("口徑二：2026 全年平均上線 GW：高情境")
    gwin = {1: [I("口徑一：2026 Q2 季平均上線 GW：" + k) for k in ("低情境", "基準", "高情境")],
            2: [gw_lo, gw_mid, gw_hi],
            3: [I("口徑三：2026 年底上線 GW：" + k) for k in ("低情境", "基準", "高情境")]}
    mixsw, dsw, tol = I("世代組合來源"), I("生產折減讀法"), I("預覽值對帳容許相對誤差")
    tol2 = I("證據對帳容許相對誤差")
    hf = {k: I("交接檔 2026 世代組合：" + zh) for k, zh in GENS}
    custom, usd, mt = I("世代產能係數：自研／其他"), I("單位換算：$ → $B 的除數"), I("單位換算：T token → M token")
    c_min, c_bt, c_day = I("每年分鐘數"), I("單位換算：B token → T token"), I("每年天數")
    y25, y26 = "INP_005", "INP_006"
    s_rev25 = S("營收：FY2025")
    s_q2, s_q2wsj = S("營收：2026 Q2 年化（不含廣告）"), S("營收：2026 Q2 年化（WSJ 季營收 6.7 × 4；含廣告；對照）")
    s_fy, s_rr, s_mgmt = S("營收：FY2026 預測（不含廣告）"), S("營收 run-rate：2026 年 9 月底（FT，OpenAI 淨額口徑）"), S("管理層營收目標：2026")
    rev_src = {0: s_rev25, 1: s_q2, 2: s_fy, 3: s_rr}
    s_ep16, s_bro50 = S("算力（雲端）支出：2025（媒體報導，經 Epoch 引用）"), S("算力支出：2026（Brockman 證詞）")
    s_serve50 = S("對外服務占算力比：2025（Epoch：研發與推論約各半）")
    s_api = S("API 每分鐘處理 token：2026-03-31（超過）")
    fwin = [I("J7 Q2 季平均上線 GW：" + k) for k in ("低占比＋晚日期", "低占比＋早日期", "中占比＋晚日期", "中占比＋早日期", "高占比＋晚日期", "高占比＋早日期")]
    y25in = [I("J7 2025 平均上線 GW（站點分布估）：低占比"), I("J7 2025 平均上線 GW（站點分布估）：高占比")]
    r25hi = I("2025 回本比值：上緣情境")
    d_msg, d_tok, d_free = I("ChatGPT 每日則數：2026 Q2"), I("ChatGPT 每則提示 token 數"), I("ChatGPT token 中免費用戶占比")

    # ═══ 一、案例定義 ═══
    X.section("一、案例定義：年度、平均上線 GW（欄 D–N）、營收（基準）")
    X.add("年度", "年", lambda i, c: f"={y25}" if YI[i] == 0 else f"={y26}", "yr", "INP_005／INP_006（2025 為實際校準年）", name="PXQ_Year")
    X.add("年度序位（年度 − 起始年 + 1；供 INDEX 取 CMP_／REV_ 年度向量）", "序位", lambda i, c: f"={c}{X.rows['yr']}-{y25}+1", "idx", "年數含頭尾的 +1（恆等式）")

    def gw_f(i, c):
        if SCOPE[i] == 0:
            return ["=(TK_SRC_DEM_009-TK_SRC_DEM_008)/LN(TK_SRC_DEM_009/TK_SRC_DEM_008)", "=AVERAGE(TK_SRC_DEM_008,TK_SRC_DEM_009)"][i]
        return f"={gwin[SCOPE[i]][LEVEL[i]]}"
    X.add("平均上線 GW（口徑一＝Q2 季平均；口徑二＝全年平均；口徑三＝年底時點）", "GW", gw_f, "gw",
          "2025：OpenAI 官方端點（2024 年底 0.6、2025 年底 1.9；Interested-party；電力口徑未明）內插——指數＝(端點差)÷ln(端點比)，線性＝端點平均（Derived）；"
          "口徑一：J7 站點由下而上（Derived）；口徑二：Inputs（Assumed，Andy『約 3 GW』）；口徑三：A 段年底三情境（Assumed）。官方未公布 2026 年上線 GW", fmt=NUM3, name="PXQ_GW")
    X.add("營收（基準）", "$B", lambda i, c: f"={rev_src[SCOPE[i]]}", "rev",
          "2025＝SRC_OAI（13.07，Checks 驗等於 TK_SRC_DEM_007）；口徑一＝Q2 年化 26.3（不含廣告）；口徑二＝FY2026 預測 35（不含廣告）；口徑三＝9 月底 run-rate 50（二手，非全年營收）。"
          "三者皆為 Interested-party 或二手報導，標記照 SRC_OAI", fmt=NUM2, name="PXQ_Rev")

    # ═══ 二、世代參數表 ═══
    X.section("二、世代參數（每 GW；TK_Link 基準成本情境）：欄 D 年持有成本、E 每 GW 付費年產能、F 生產折減（v19.0 讀法）、G 生產折減（一律 IF_ProdDerate）、H 採用（Inputs 開關）")
    X.hdr(["世代", "TK_HdrGen 對照名稱", "持有成本 $B/GW/年", "付費年產能 M tok/GW/年", "折減（v19.0 讀法）", "折減（一律 ProdDerate）", "折減（採用）"], "g_hdr")
    r_cost = X.r
    put(ws, f"B{r_cost}", "成本情境（對應 TK_HdrCost；基準）", F_CALC)
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
            put(ws, f"D{r}", f"=SUMIFS(TK_IF_HoldEcon,TK_HdrGen,$C{r},TK_HdrCost,$D${r_cost})", fmt=NUM3)
            put(ws, f"E{r}", f"=SUMIFS(TK_IF_TokGWPaidMix,TK_HdrGen,$C{r},TK_HdrCost,$D${r_cost})", fmt=SCI)
        else:
            put(ws, f"C{r}", "（VR200 × 世代產能係數）", F_CALC)
            put(ws, f"D{r}", f"=D{g['VR200']}*{custom}", fmt=NUM3)
            put(ws, f"E{r}", f"=E{g['VR200']}*{custom}", fmt=SCI)
        put(ws, f"F{r}", "=TK_IF_FleetProdRatio_VR200" if key == "VR200" else "=TK_IF_ProdDerate", fmt=NUM3)
        put(ws, f"G{r}", "=TK_IF_ProdDerate", fmt=NUM3)
        put(ws, f"H{r}", f"=(1-{dsw})*F{r}+{dsw}*G{r}", fmt=NUM3)
        put(ws, f"P{r}", {"Custom": "自研晶片：持有成本與產能＝VR200 × 世代產能係數（沿用 v0.6 P4／P3 做法；Analogy，係數區間 0.4–1.0）；折減用 IF_ProdDerate（非 VR200）",
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

    def block(tag, title_, mix, dcol, serve, rev_fn=None, gw_fn=None, full=True, mask=None):
        """mix：main／v06／handoff；dcol：折減欄（H 採用、F v19.0、G 一律）；serve：服務占比公式片段；mask：只建立為 True 的案例欄。"""
        X.section(title_)
        X.mask = mask
        rr = {}
        X.add("平均上線 GW", "GW", lambda i, c: gw_fn(i, c) if gw_fn else f"={c}{X.rows['gw']}", f"{tag}_gw", "", fmt=NUM3)
        for key, zh in GENS:
            rr[key] = X.add(f"世代占比：{zh}", "比例", lambda i, c, key=key: sh_expr(mix, key, i, c), f"{tag}_sh_{key}", "", fmt=PERC)
        s1, s6 = rr["Hopper"], rr["Custom"]
        X.add("世代占比合計（應＝1）", "比例", lambda i, c: f"=SUM({c}{s1}:{c}{s6})", f"{tag}_shsum", "Checks 驗證", fmt=PERC)
        X.add("每 GW 年持有成本 h（世代加權）", "$B/GW/年", lambda i, c: f"=SUMPRODUCT({c}{s1}:{c}{s6},$D${g1}:$D${g6})", f"{tag}_h", "h＝Σ 世代占比 × TK_IF_HoldEcon（基準欄）", fmt=NUM3, name="PXQ_h" if tag == "m" else None)
        X.add("算力成本 C＝平均 GW × h", "$B", lambda i, c: f"={c}{X.rows[f'{tag}_gw']}*{c}{X.rows[f'{tag}_h']}", f"{tag}_C", "年度算力成本（機隊持有成本口徑；Andy：『從持有的機隊成本去推算』）", fmt=NUM2, name="PXQ_C" if tag == "m" else None)
        X.add("每 GW 付費年產能（世代加權，已乘生產折減）", "M tok/GW/年",
              lambda i, c: f"=SUMPRODUCT({c}{s1}:{c}{s6},$E${g1}:$E${g6},${dcol}${g1}:${dcol}${g6})", f"{tag}_cap", "Σ 世代占比 × TK_IF_TokGWPaidMix × 生產折減", fmt=SCI)
        X.add("Q（M tok／年）＝GW × 服務占比 ×（1 − 免費占比）× 每 GW 付費年產能 × 實際利用率", "M tok/年",
              lambda i, c: f"={c}{X.rows[f'{tag}_gw']}*{serve}*(1-TK_IF_FreeShare)*{c}{X.rows[f'{tag}_cap']}*TK_IF_UtilActual", f"{tag}_Q",
              "Andy：『Q 就是真正拿去用作服務的算力，去乘上每單位算力可以產生的 token 數』", fmt=SCI, name="PXQ_Q" if tag == "m" else None)
        X.add("Q（T token／年）", "T tok/年", lambda i, c: f"={c}{X.rows[f'{tag}_Q']}/{mt}", f"{tag}_QT", "顯示用（交接檔『約 6.4×10¹⁵ token』＝6,400 T）", fmt="#,##0")
        X.add("有效牌價 P_eff（付費層級組合；已含折扣 d）", "$/M tok", lambda i, c: "=TK_IF_PxEffMix25" if YI[i] == 0 else "=TK_IF_PxEffMix", f"{tag}_P",
              "2025＝IF_PxEffMix25（年度加權）；2026＝IF_PxEffMix（2026 快照）", fmt=NUM3, name="PXQ_Peff" if tag == "m" else None)
        if full:
            X.add("表列牌價＝P_eff ÷（1 − d）", "$/M tok", lambda i, c: f"={c}{X.rows[f'{tag}_P']}/(1-TK_IF_Disc)", f"{tag}_Plist", "d＝TK_IF_Disc，不寫死", fmt=NUM3)
        X.add("營收", "$B", lambda i, c: rev_fn(i, c) if rev_fn else f"={c}{X.rows['rev']}", f"{tag}_rev", "", fmt=NUM2)
        X.add("實際 m×L＝營收 ÷（Q × P_eff）", "倍", lambda i, c: f"={c}{X.rows[f'{tag}_rev']}*{usd}/({c}{X.rows[f'{tag}_Q']}*{c}{X.rows[f'{tag}_P']})", f"{tag}_act",
              "m 只看收費業務、跟營收有關跟成本無關", fmt=NUM3, name="PXQ_mL_Actual" if tag == "m" else None)
        X.add("所需 m×L＝算力成本 ÷（Q × P_eff）", "倍", lambda i, c: f"={c}{X.rows[f'{tag}_C']}*{usd}/({c}{X.rows[f'{tag}_Q']}*{c}{X.rows[f'{tag}_P']})", f"{tag}_req",
              "回本所需", fmt=NUM3, name="PXQ_mL_Required" if tag == "m" else None)
        X.add("比值＝實際 m×L ÷ 所需 m×L", "倍", lambda i, c: f"={c}{X.rows[f'{tag}_act']}/{c}{X.rows[f'{tag}_req']}", f"{tag}_ratio", "主輸出；＝營收 ÷ 算力成本", fmt=NUM3, name="PXQ_Ratio" if tag == "m" else None)
        if full:
            X.add("恆等式：比值 −（營收 ÷ 算力成本）（應＝0）", "倍", lambda i, c: f"={c}{X.rows[f'{tag}_ratio']}-{c}{X.rows[f'{tag}_rev']}/{c}{X.rows[f'{tag}_C']}", f"{tag}_ident", "Checks 驗證", fmt="0.0E+00", name="PXQ_Identity" if tag == "m" else None)
            X.add("隱含 m＝實際 m×L ÷ L（L＝2025 機隊，TK_IF_LTotalFleetVDisc）", "倍", lambda i, c: f"={c}{X.rows[f'{tag}_act']}/TK_IF_LTotalFleetVDisc", f"{tag}_mfleet",
                  "只作顯示（判斷題 1：L 讀法預設＝壽命期價格係數）", fmt=NUM3)
            X.add("隱含 m＝實際 m×L ÷ L（L＝新世代，TK_IF_LTotalDisc）", "倍", lambda i, c: f"={c}{X.rows[f'{tag}_act']}/TK_IF_LTotalDisc", f"{tag}_mnew", "只作顯示", fmt=NUM3)
        rg = f"{tag}_ratio"
        sel = f"((1-{sw25})*$D{X.rows[rg]}+{sw25}*$E{X.rows[rg]})"
        for lab, ref, key in (("持平所需 2026 平均 GW（以 2025 指數比值）", f"$D{X.rows[rg]}", "be_exp"), ("持平所需 2026 平均 GW（以 2025 線性比值）", f"$E{X.rows[rg]}", "be_lin"),
                              ("持平所需 2026 平均 GW（以 2025 主值比值）", sel, "be_sel")):
            if not full and key != "be_sel" and tag not in ("pv",):
                continue
            X.add(lab + "＝2026 營收 ÷（2026 h × 2025 比值）", "GW",
                  lambda i, c, ref=ref: None if YI[i] == 0 else f"={c}{X.rows[f'{tag}_rev']}/({c}{X.rows[f'{tag}_h']}*{ref})", f"{tag}_{key}",
                  "該案例的 2026 GW 低於此值，比值才高於 2025" if key == "be_sel" else "", fmt=NUM3,
                  name={"be_exp": "PXQ_BE_Exp", "be_lin": "PXQ_BE_Lin", "be_sel": "PXQ_BE_Sel"}[key] if tag == "m" else None)
        X.mask = None
        return rr

    serve_base = "TK_IF_ServeShare"
    block("m", "三、主序列（基準）：世代組合依 Inputs 開關（基準＝OpenAI v0.6 自有 V1 組合；口徑三年底暫用 2026 年組合）、生產折減依 Inputs 開關（基準＝v19.0 讀法）、服務占比＝Tokenomics",
          "main", "H", serve_base)

    # ═══ 四、敏感度 ═══
    X.section("四、敏感度與對照基準（下列各區塊與主序列只差一個條件；比值不受服務占比影響；只建立適用的口徑欄）")
    block("pv", "4.1 預覽值基準：交接檔世代組合（2026）＋一律 IF_ProdDerate（對帳與第八節用；其餘同基準）", "handoff", "G", serve_base)
    block("sv", "4.2 服務占比改為 Epoch 對照值（SRC_OAI『研發與推論約各半』）：Q 與 m×L 改變，比值不變", "main", "H", s_serve50, full=False)
    block("rv", "4.3 口徑二營收改用模型正向（REV_GrossCapped 2026；截頂後總額）", "main", "H", serve_base, full=False, mask=MASK_Y25_S2,
          rev_fn=lambda i, c: f"={s_rev25}" if YI[i] == 0 else f"=INDEX(REV_GrossCapped,{c}${X.rows['idx']})")
    block("mx", "4.4 世代組合改用交接檔（2026；v19.0 生產折減讀法不變）", "handoff", "H", serve_base, full=False)
    block("su", "4.5 口徑二 GW 改用 v0.6 供給 GW（CMP_SupplyGW 2026；合約＋自建；是否年平均未明）", "main", "H", serve_base, full=False, mask=MASK_Y25_S2,
          gw_fn=lambda i, c: f"={c}{X.rows['gw']}" if YI[i] == 0 else f"=INDEX(CMP_SupplyGW,{c}${X.rows['idx']})")
    block("rm", "4.6 口徑二營收改用管理層目標 36（SRC_OAI_043；含廣告；原『市場傳聞 360 億』，v0.6.3 起僅作對照）", "main", "H", serve_base, full=False, mask=MASK_Y25_S2,
          rev_fn=lambda i, c: f"={s_rev25}" if YI[i] == 0 else f"={s_mgmt}")
    block("rw", "4.7 口徑一營收改用 WSJ 季營收 × 4＝26.8（含廣告；對照）", "main", "H", serve_base, full=False, mask=MASK_Y25_S1,
          rev_fn=lambda i, c: f"={s_rev25}" if YI[i] == 0 else f"={s_q2wsj}")

    # ═══ 五、對照列 ═══
    X.section("五、對照列（外部報導的算力支出 ÷ 持有成本 GW × h；不進主序列）")
    X.add("外部報導算力支出（2025：雲端，Epoch 引媒體；2026：Brockman 證詞）", "$B", lambda i, c: f"={s_ep16}" if YI[i] == 0 else f"={s_bro50}", "x_ext",
          "SRC_OAI（新列；出處網址見該列）", fmt=NUM2, name="PXQ_ExtSpend")
    X.add("倍數＝外部支出 ÷（GW × h）", "倍", lambda i, c: f"={c}{X.rows['x_ext']}/{c}{X.rows['m_C']}", "x_mult",
          "固定倍數在比值趨勢中約掉，但須逐年用同一口徑；2026 支出為全年數，與口徑二（全年平均 GW）可比，口徑一、三僅供參考", fmt=NUM2, name="PXQ_ExtMultiple")
    X.add("2025 推論＋訓練支出（TK_SRC_DEM_004＋006）", "$B", lambda i, c: "=TK_SRC_DEM_004+TK_SRC_DEM_006" if YI[i] == 0 else None, "x_it",
          "Andy：『兩百零四億，很明顯是高估了』——只作對照，不作算力成本", fmt=NUM2)
    X.add("倍數＝（推論＋訓練）÷（GW × h）", "倍", lambda i, c: f"={c}{X.rows['x_it']}/{c}{X.rows['m_C']}" if YI[i] == 0 else None, "x_it_mult", "", fmt=NUM2)

    # ═══ 六、預覽值對帳 ═══
    X.section("六、預覽值對帳。6A＝交接檔第 6 節（chat 端以 Tokenomics v18.0 計得）的歷史對照，v0.6.3 起不再計入『超出容許誤差項數』；6B＝A 段／J7 證據與工作單驗收值（容差見 Inputs『證據對帳容許相對誤差』）")
    X.hdr(["6A 歷史對照項目", "單位", "交接檔值（v18.0）", "模型（預覽基準）", "相對差", "判定（不計數）", "模型（主序列）", "主序列相對差"], "rc_hdr")
    pv = lambda k, c: f"{OLD_COL[c]}{X.rows['pv_' + k]}"            # noqa: E731
    mm = lambda k, c: f"{OLD_COL[c]}{X.rows['m_' + k]}"             # noqa: E731
    items = [
        ("每 GW 年持有成本 2025（93.5 億）", "$B", 9.35, "h", "D"), ("每 GW 年持有成本 2026（97.3 億）", "$B", 9.73, "h", "G"),
        ("算力成本 2025 指數（106 億）", "$B", 10.6, "C", "D"), ("算力成本 2025 線性（117 億）", "$B", 11.7, "C", "E"),
        ("算力成本 2026 2.5 GW（243 億；口徑二低）", "$B", 24.3, "C", "F"), ("算力成本 2026 3.0 GW（292 億；口徑二基準）", "$B", 29.2, "C", "G"),
        ("算力成本 2026 3.5 GW（341 億）；口徑二高端已改 3.8，不再可比", "$B", 34.1, "C", "H"),
        ("比值 2025 指數", "倍", 1.24, "ratio", "D"), ("比值 2025 線性", "倍", 1.12, "ratio", "E"),
        ("比值 2026 2.5 GW（口徑二低）", "倍", 1.48, "ratio", "F"), ("比值 2026 3.0 GW（口徑二基準）", "倍", 1.23, "ratio", "G"),
        ("比值 2026 3.5 GW；口徑二高端已改 3.8，不再可比", "倍", 1.06, "ratio", "H"),
        ("持平所需 GW（2025 指數比值；口徑二）", "GW", 2.99, "be_exp", "G"), ("持平所需 GW（2025 線性比值；口徑二）", "GW", 3.31, "be_lin", "G"),
        ("Q 2025 線性 1.25 GW（約 6.4×10¹⁵ token）", "T tok", 6400, "QT", "E"),
    ]
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
        X.r += 1
    put(ws, f"P{X.r - 1}", "交接檔值為 chat 端以 Tokenomics v18.0 計得的預覽（金額由『億』換算為 $B 的常數，只作對照）；v0.6.3 起 2026 欄對應口徑二（2.5／3.0 GW 與 v0.6.1 相同；高端 3.5→3.8）。歷史對照，不計入超出容許誤差項數", F_NOTE)
    X.r += 1
    X.hdr(["6B 證據對帳項目", "單位", "證據值", "模型值", "相對差", "判定（計數）", "出處"], "rc2_hdr")
    ev = []

    def rec(lab, unit, expect, tpl, src, fmt=NUM3):
        r = X.r
        X.n += 1
        put(ws, f"A{r}", f"X{X.n:02d}", F_CALC)
        put(ws, f"B{r}", lab, F_CALC)
        put(ws, f"C{r}", unit, F_CALC)
        put(ws, f"D{r}", expect, F_NOTE, fmt=fmt)
        X.deferred.append((f"E{r}", tpl, fmt))
        put(ws, f"F{r}", f"=(E{r}-D{r})/D{r}", fmt="0.00%")
        put(ws, f"G{r}", f'=IF(ABS(F{r})<={tol2},"OK","差異")')
        put(ws, f"H{r}", src, F_NOTE)
        X.r += 1
        ev.append(r)
    A, J = "A 段『比值』", "J7『結果』"
    rec("2025 比值：指數內插", "倍", 1.239, "=D{m_ratio}", A + "B4（任務給定 1.239）")
    rec("2025 比值：線性內插", "倍", 1.118, "=E{m_ratio}", A + "表 E B57（1.118）")
    rec("持平門檻 口徑一 Q2（2025 比值 1.239）", "GW", 2.1669, "=E{t_be1}", A + "B11（h 9.796）；J7 B4")
    rec("持平門檻 口徑二 FY（2025 比值 1.239）", "GW", 2.8837, "=E{t_be2}", A + "B13；J7 B5")
    rec("持平門檻 口徑三 run-rate（2025 比值 1.239）", "GW", 4.1196, "=E{t_be3}", A + "B52；J7 B6")
    rec("持平門檻 口徑一（2025 比值 1.118）", "GW", 2.4014, "=D{t_be1}", A + "B58")
    rec("持平門檻 口徑二（2025 比值 1.118）", "GW", 3.1958, "=D{t_be2}", A + "B59")
    rec("持平門檻 口徑三（2025 比值 1.118）", "GW", 4.5654, "=D{t_be3}", A + "B60")
    rec("持平門檻 口徑一（2025 比值 1.397）", "GW", 1.9218, "=F{t_be1}", A + "D58")
    rec("持平門檻 口徑二（2025 比值 1.397）", "GW", 2.5575, "=F{t_be2}", A + "D59")
    rec("持平門檻 口徑三（2025 比值 1.397）", "GW", 3.6536, "=F{t_be3}", A + "D60")
    rec("口徑一基準倍數（2.34 GW、263 億）≈0.93", "倍", 0.93, "=E{t_mu1}", "工作單 20261025 驗收 3（J7 中占比＋晚 0.926）", NUM2)
    rec("口徑二基準倍數（3.0 GW、350 億）≈0.96", "倍", 0.96, "=E{t_mu2}", "工作單 20261025 驗收 3", NUM2)
    rec("口徑三基準倍數（4.6 GW、500 億）≈0.90", "倍", 0.90, "=E{t_mu3}", "工作單 20261025 驗收 3（A 段 4.57 GW 為 0.901）", NUM2)
    for k, (lab, v) in enumerate((("低占比＋晚", 0.98249), ("低占比＋早", 0.90648), ("中占比＋晚", 0.92599), ("中占比＋早", 0.83742), ("高占比＋晚", 0.85707), ("高占比＋早", 0.76092))):
        rec(f"Fairwater 條件：Q2 倍數（{lab}）", "倍", v, "=F{fw%d}" % (k + 1), J + f"C{9 + k}", NUM3)
    rec("2025 站點估比值（低占比，GW 1.11）", "倍", 1.2541, "=E{y25_1}", J + "C19", NUM3)
    rec("2025 站點估比值（高占比，GW 1.06）", "倍", 1.3182, "=E{y25_2}", J + "C21", NUM3)
    X.line("對帳超出容許誤差（6B）的項數", "rc_bad", {"E": f'=COUNTIF(G{ev[0]}:G{ev[-1]},"差異")'},
           note="逐項說明見報告：超出者是否來自 Tokenomics 快照差異（A 段／J7 用 TK v21.2，本模型用 v21.4）或證據值取整")
    nm(wb, "PXQ_ReconBad", f"PxQ!$E${X.rows['rc_bad']}")

    # ═══ 七、Fairwater 條件（J7）與 2025 站點估對照（J6）═══
    X.section("七、J7 對照：Q2 季平均 GW 的六組情境（占比 × 上線日期；Epoch 站點由下而上）與 Fairwater 條件；7.2 為 2025 平均 GW 的站點分布估（J6）")
    X.hdr(["7.1 情境（占比／上線日期）", "單位", "Q2 季平均 GW", "比值＝Q2 年化營收 ÷（GW × h）", "倍數（對 2025 主值）", "倍數（營收 WSJ 26.8）"], "fw_hdr")
    fwlab = ["低占比＋晚日期（Epoch 日期）", "低占比＋早日期（證據下限）", "中占比＋晚日期（Epoch 日期）", "中占比＋早日期（證據下限）", "高占比＋晚日期（Epoch 日期）", "高占比＋早日期（證據下限）"]
    for k in range(6):
        X.line(fwlab[k], f"fw{k + 1}", {
            "D": f"={fwin[k]}",
            "E": "=$G${m_rev}/(D{fw%d}*$G${m_h})" % (k + 1),
            "F": "=E{fw%d}/PXQ_Ratio2025Sel" % (k + 1),
            "G": "=$G${rw_rev}/(D{fw%d}*$G${m_h})/PXQ_Ratio2025Sel" % (k + 1),
        }, note="Derived（J7 結果頁 B%d）；GW 為 Epoch 追蹤站點的下界；占比＝OpenAI 對 likely／speculative 站（含 Microsoft Fairwater）的使用占比假設" % (9 + k),
            fmt_map={"D": NUM2, "E": NUM3, "F": NUM3, "G": NUM3}, unit="GW")
    X.line("六組最小", "fw_min", {"D": "=MIN(D{fw1}:D{fw6})", "F": "=MIN(F{fw1}:F{fw6})", "G": "=MIN(G{fw1}:G{fw6})"}, fmt_map={"D": NUM2, "F": NUM2, "G": NUM2}, note="驗收：倍數範圍 0.76–0.98（J7）")
    X.line("六組最大", "fw_max", {"D": "=MAX(D{fw1}:D{fw6})", "F": "=MAX(F{fw1}:F{fw6})", "G": "=MAX(G{fw1}:G{fw6})"}, fmt_map={"D": NUM2, "F": NUM2, "G": NUM2})
    nm(wb, "PXQ_FW_GW", f"PxQ!$D${X.rows['fw1']}:$D${X.rows['fw6']}")
    nm(wb, "PXQ_FW_Mult", f"PxQ!$F${X.rows['fw1']}:$F${X.rows['fw6']}")
    put(ws, f"B{X.r}", "J7 結論（說明文字，非本模型計算）：『若 OpenAI 用到 Microsoft Fairwater 兩站四成以上，或新站實際上線早於 Epoch 快照日，Q2 比值低於 2025；兩個條件都不成立時，比值約持平到略升。』"
        "Fairwater Wisconsin／Atlanta 約占 Q2 增量 58%，OpenAI 占比未揭露（不可分離）；Epoch 未涵蓋的增量（Azure 其他、Oracle 其他、Google、AWS、AMD、Cerebras）只會使 GW 更高。", F_NOTE)
    X.r += 2
    X.hdr(["7.2 2025 站點分布估（J6）", "單位", "2025 平均 GW", "2025 比值＝營收 ÷（GW × h）", "相對 2025 指數內插比值"], "y25_hdr")
    for k, lab in enumerate(("低占比（J7 B19）", "高占比（J7 B21）")):
        X.line(f"2025 平均上線 GW 站點分布估：{lab}", f"y25_{k + 1}", {
            "D": f"={y25in[k]}", "E": "=$D${m_rev}/(D{y25_%d}*$D${m_h})" % (k + 1), "F": "=E{y25_%d}/$D${m_ratio}" % (k + 1)},
            note="Derived；對照列，不進主序列、不改 INP_276 開關（J6：1.128 偏高、基準 1.239 偏保守）", fmt_map={"D": NUM2, "E": NUM3, "F": NUM3}, unit="GW")
    X.line("2025 比值上緣（A 段表 E；Inputs）", "y25_hi", {"E": f"={r25hi}"}, fmt_map={"E": NUM3}, unit="倍",
           note="Derived（A 段『比值』D57＝1.397）；與 2025 平均 GW 約 1.0 相合（營收 13.07 ÷ 持有成本 9.353），起源待 Project 確認；只用於頂部持平門檻表")

    # ═══ 八、GW 路線 token 帳 ═══
    X.section("八、GW 路線 token 帳（2026 Q2，口徑一基準）＝Q（口徑一基準 GW）÷ 需求面付費 token（API 時點值年化 ＋ ChatGPT 付費 token）；＝1 表示完全閉合")
    X.hdr(["8.1 項目", "單位", "(A) 交接檔組合 × 折減 0.85（證據重現）", "(B) 交接檔組合 × 重解比 0.806（證據重現）", "(C) 本模型基準（V1 組合、INP_287 讀法）"], "w_hdr")
    X.line("口徑一基準 Q2 季平均 GW", "w_gw", {c: "=$G${m_gw}" for c in "DEF"}, fmt_map={c: NUM3 for c in "DEF"}, unit="GW")
    X.line("每 GW 付費年產能（世代加權，未乘生產折減）", "w_cap0", {
        "D": "=SUMPRODUCT($G${pv_sh_Hopper}:$G${pv_sh_Custom},$E$%d:$E$%d)" % (g1, g6),
        "E": "=SUMPRODUCT($G${pv_sh_Hopper}:$G${pv_sh_Custom},$E$%d:$E$%d)" % (g1, g6),
        "F": "=SUMPRODUCT($G${m_sh_Hopper}:$G${m_sh_Custom},$E$%d:$E$%d)" % (g1, g6)},
           fmt_map={c: SCI for c in "DEF"}, unit="M tok/GW/年", note="(A)(B)＝Tokenomics Cap_In 2026 服務世代組合（0.25／0.45／0.25／0.05，即交接檔組合 INP_281–286）；(C)＝本模型 V1 組合（INP_280 開關）")
    X.line("生產折減", "w_der", {"D": "=TK_IF_ProdDerate", "E": "=TK_IF_FleetProdRatio_VR200", "F": "=$G${m_cap}/F{w_cap0}"}, fmt_map={c: NUM3 for c in "DEF"}, unit="倍",
           note="(A) 一律 IF_ProdDerate 0.85；(B) 一律 IF_FleetProdRatio_VR200（Q 的重解比 0.806，證據『R1』）；(C) 為基準欄實際加權折減（隨 INP_287 與組合）")
    X.line("Q（M tok／年）＝GW × 服務占比 ×（1 − 免費占比）× 每 GW 付費年產能 × 折減 × 實際利用率", "w_Q", {
        c: "=%s{w_gw}*TK_IF_ServeShare*(1-TK_IF_FreeShare)*%s{w_cap0}*%s{w_der}*TK_IF_UtilActual" % (c, c, c) for c in "DEF"}, fmt_map={c: SCI for c in "DEF"}, unit="M tok/年",
           note="與 PxQ 主序列 Q 同式（GW 路線＝P×Q 的 Q 所用口徑）；(C) 應等於主序列口徑一基準欄的 Q")
    X.line("核對：（Q −（4.1 預覽基準 Q／主序列 Q））÷ 該 Q（應＝0；相對差）", "w_Qchk", {"D": "=(D{w_Q}-$G${pv_Q})/$G${pv_Q}", "F": "=(F{w_Q}-$G${m_Q})/$G${m_Q}"}, fmt_map={"D": "0.0E+00", "F": "0.0E+00"}, unit="比例",
           note="(A) 對 4.1 預覽基準區塊（交接檔組合＋一律 0.85）；(C) 對主序列；Checks 驗證")
    X.hdr(["8.2 需求面付費 token（2026 Q2，年化）", "單位", "(A)", "(B)", "(C)"], "w_hdr2")
    X.line("API 付費 token＝每分鐘處理 token（時點值）× 每年分鐘數 → M tok／年", "w_api", {
        c: f"={s_api}*{c_min}/{c_bt}*{mt}" for c in "DEF"}, fmt_map={c: SCI for c in "DEF"}, unit="M tok/年",
           note="SRC_OAI『API 每分鐘處理 token：2026-03-31』15 B tok/min（Interested-party）；Q2 當期值已搜尋，未找到，以 3 月底時點值年化（證據檔 D1 候選因子『Q2 季平均 ÷ 3 月底』1.0／1.15／1.35 未納入）")
    X.line("ChatGPT 付費 token＝每日則數 × 每則 token × 每年天數 ×（1 − 免費占比）→ M tok／年", "w_chat", {
        c: f"={d_msg}*{d_tok}*{c_day}/{c_bt}*{mt}*(1-{d_free})" for c in "DEF"}, fmt_map={c: SCI for c in "DEF"}, unit="M tok/年",
           note="Inputs：每日則數 3.6 B（Assumed；2.5 B×週活比）、每則 4,000 token、免費占比 0.6——照 Tokenomics Alloc_In／Cap_In 原紀錄抄錄")
    X.line("需求面付費 token 合計", "w_D", {c: "=%s{w_api}+%s{w_chat}" % (c, c) for c in "DEF"}, fmt_map={c: SCI for c in "DEF"}, unit="M tok/年")
    X.hdr(["8.3 結果與對帳", "單位", "(A)", "(B)", "(C)"], "w_hdr3")
    X.line("GW 路線 token 帳＝Q ÷ 需求面付費 token（1＝完全閉合）", "w_led", {c: "=%s{w_Q}/%s{w_D}" % (c, c) for c in "DEF"}, fmt_map={c: NUM3 for c in "DEF"}, unit="倍",
           note="(A)(B) 重現 token帳缺口 A 段『拆解』B7（2.10）與 R1（1.99）；(C) 為本模型基準口徑（V1 組合、INP_287 讀法），與 (A)(B) 的差來自世代組合與折減讀法")
    nm(wb, "PXQ_TokLedger", f"PxQ!$D${X.rows['w_led']}:$F${X.rows['w_led']}")
    X.line("證據值（token帳缺口 A 段）", "w_exp", {"D": 2.10, "E": 1.99}, fmt_map={"D": NUM2, "E": NUM2}, unit="倍",
           note="token帳缺口_A段.xlsx『拆解』B7＝2.0966（生產折減 0.85）、C17＝1.9876（0.806）；工作單驗收取 2.10／1.99，容差 1%")
    X.line("相對差", "w_rel", {"D": "=(D{w_led}-D{w_exp})/D{w_exp}", "E": "=(E{w_led}-E{w_exp})/E{w_exp}"}, fmt_map={"D": "0.00%", "E": "0.00%"})
    X.line("判定（容差＝Inputs『證據對帳容許相對誤差』）", "w_ok", {"D": f'=IF(ABS(D{{w_rel}})<={tol2},"OK","差異")', "E": f'=IF(ABS(E{{w_rel}})<={tol2},"OK","差異")'})
    nm(wb, "PXQ_TokLedgerRel", f"PxQ!$D${X.rows['w_rel']}:$E${X.rows['w_rel']}")
    X.line("GW 路線 token 帳（本模型基準口徑；口徑一 低／基準／高 GW）", "w_led3", {
        "D": "=$F${m_Q}/$F${w_D}", "E": "=$G${m_Q}/$F${w_D}", "F": "=$H${m_Q}/$F${w_D}"}, fmt_map={c: NUM3 for c in "DEF"}, unit="倍",
           note="欄 D／E／F＝口徑一 低／基準／高（主序列 Q ÷ 需求面付費 token (C)）")
    X.line("VR200 每 GW P×Q 營收的 GW 路線折算", "w_vr", {"D": "未建"},
           note="TK `IFW_`／`IF_` 名稱中沒有 VR200 P×Q 營收（`IF_RevGW*`／`IFW_RevGW*` 為理論營收上限，不含 m、服務占比、免費占比）；"
                "只有 `L1_RevGW_PQ_VR200`（本模型不得讀 `L1_`）。照工作單不建，列入報告『TK 缺 IF 名稱』")

    # ═══ 頂部主輸出 ═══
    rows = X.rows
    rows.update(t_hdr=8, t_r1=9, t_r2=10, t_r3=11, t_r25=12, t_mhdr=14, t_mu1=15, t_mu2=16, t_mu3=17, t_bhdr=19, t_bbase=20, t_be1=21, t_be2=22, t_be3=23)
    put(ws, "A7", "主輸出（一）：回本比值（實際 m×L ÷ 所需 m×L ＝ 營收 ÷ 算力成本）— 2026 三口徑 × 低／基準／高；2025 對照", F_BOLD, fill=FILL_SEC)
    for c, lab in zip("BCDEFGHIJ", ("口徑", "單位", "低", "基準", "高", "GW 低", "GW 基準", "GW 高", "營收 $B")):
        put(ws, f"{c}8", lab, F_BOLD, fill=FILL_HDR)
    for s in (1, 2, 3):
        r = 8 + s
        lo, ba, hi = SC_COLS[s]
        put(ws, f"B{r}", SC_NAME[s], F_BOLD)
        put(ws, f"C{r}", "倍", F_CALC)
        for tc, sc in zip("DEF", (lo, ba, hi)):
            put(ws, f"{tc}{r}", f"={sc}{rows['m_ratio']}", F_BOLD, fmt=NUM2)
        for tc, sc in zip("GHI", (lo, ba, hi)):
            put(ws, f"{tc}{r}", f"={sc}{rows['m_gw']}", fmt=NUM3)
        put(ws, f"J{r}", f"={ba}{rows['m_rev']}", fmt=NUM2)
    put(ws, "B12", "2025 對照（D＝指數內插、E＝線性內插、F＝主值〔Inputs 開關〕；G、H＝對應 GW；J＝營收）", F_CALC)
    put(ws, "C12", "倍", F_CALC)
    put(ws, "D12", f"=D{rows['m_ratio']}", F_BOLD, fmt=NUM2)
    put(ws, "E12", f"=E{rows['m_ratio']}", F_BOLD, fmt=NUM2)
    put(ws, "F12", f"=(1-{sw25})*D12+{sw25}*E12", F_BOLD, fmt=NUM2)
    put(ws, "G12", f"=D{rows['m_gw']}", fmt=NUM3)
    put(ws, "H12", f"=E{rows['m_gw']}", fmt=NUM3)
    put(ws, "J12", f"=D{rows['m_rev']}", fmt=NUM2)
    nm(wb, "PXQ_Ratio2025Sel", "PxQ!$F$12")
    nm(wb, "PXQ_Out_Ratio", "PxQ!$D$9:$F$11")
    put(ws, "A14", "主輸出（二）：倍數＝比值 ÷ 2025 主值比值", F_BOLD, fill=FILL_SEC)
    for s in (1, 2, 3):
        r = 14 + s
        put(ws, f"B{r}", SC_NAME[s], F_CALC)
        put(ws, f"C{r}", "倍", F_CALC)
        for tc in "DEF":
            put(ws, f"{tc}{r}", f"={tc}{8 + s}/PXQ_Ratio2025Sel", F_BOLD, fmt=NUM2)
    nm(wb, "PXQ_Out_Mult", "PxQ!$D$15:$F$17")
    put(ws, "A19", "主輸出（三）：持平門檻 GW（2026 比值＝2025 比值時的 GW；GW 高於此值，比值低於 2025）— 以 2025 比值的三個基準各算一次", F_BOLD, fill=FILL_SEC)
    for c, lab in zip("BCDEFG", ("口徑", "單位", "2025 線性內插", "2025 指數內插", "2025 上緣（A 段）", "2025 主值")):
        put(ws, f"{c}19", lab, F_BOLD, fill=FILL_HDR)
    put(ws, "B20", "2025 比值（基準值）", F_CALC)
    put(ws, "C20", "倍", F_CALC)
    put(ws, "D20", f"=E{rows['m_ratio']}", fmt=NUM3)
    put(ws, "E20", f"=D{rows['m_ratio']}", fmt=NUM3)
    put(ws, "F20", f"={r25hi}", fmt=NUM3)
    put(ws, "G20", "=PXQ_Ratio2025Sel", fmt=NUM3)
    for s in (1, 2, 3):
        r = 20 + s
        ba = SC_COLS[s][1]
        put(ws, f"B{r}", SC_NAME[s], F_CALC)
        put(ws, f"C{r}", "GW", F_CALC)
        for tc in "DEFG":
            put(ws, f"{tc}{r}", f"=${ba}${rows['m_rev']}/(${ba}${rows['m_h']}*{tc}$20)", F_BOLD, fmt=NUM2)
    nm(wb, "PXQ_Out_BE", "PxQ!$D$21:$G$23")
    put(ws, "A25", "讀法（公式文字；數字取自本頁儲存格）", F_BOLD, fill=FILL_SEC)
    for s in (1, 2, 3):
        r = 25 + s
        put(ws, f"B{r}", SC_NAME[s], F_CALC)
        ba = {1: "Q2 年化營收", 2: "FY2026 營收", 3: "9 月底 run-rate"}[s]
        gwn = {1: "Q2 季平均 GW", 2: "全年平均 GW", 3: "年底 GW"}[s]
        put(ws, f"D{r}",
            f'="{ba} "&TEXT(J{8 + s},"0.0")&" $B ÷ {gwn}：基準 "&TEXT(H{8 + s},"0.00")&" GW 時比值 "&TEXT(E{8 + s},"0.00")&"，為 2025 主值（"&TEXT(F12,"0.00")&"）的 "&TEXT(E{14 + s},"0.00")&" 倍，"'
            f'&IF(E{8 + s}>=F12,"不低於 ","低於 ")&"2025；{gwn}高於 "&TEXT(G{20 + s},"0.00")&" GW 時才低於 2025（以 2025 主值）；低至高情境倍數 "&TEXT(D{14 + s},"0.00")&"–"&TEXT(F{14 + s},"0.00")&"。"')
    nm(wb, "PXQ_Reading", "PxQ!$D$26:$D$28")
    put(ws, "P11", "看的並不是回本的絕對數字，而是實際除以所需這個比例的趨勢（Andy 2026-10-22）；比值＞1＝實收足以覆蓋機隊持有成本。三口徑不可混用：營收是流量，Q2 年化營收配 Q2 季平均 GW（口徑一為主）；run-rate 為時點值，只配年底 GW", F_NOTE)
    put(ws, "P26", "三口徑的 2025 比值基準本身有區間 1.118／1.239／1.397（A 段表 E）；J7 站點分布估 2025 比值約 1.25–1.32（第 7.2 節），1.128 偏高、基準 1.239 偏保守（J6）", F_NOTE)
    X.resolve()
    return X
