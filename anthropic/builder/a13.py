"""Anthropic v0.2.2 三層協調頁 Coord（三層基準協調規則 v1.4.2；工作單 20261114 第二批 13 家 第 2 節；模型商樣板，OpenAI 複製）。

本頁只列示，不被任何計算頁引用（Checks 除外）；新增本頁不改變任何既有工作表的數值。結構：
  A. 參數（判定門檻、季化倍數；月數取 Inputs 定義常數）
  B. 讀取層級（同源，不判定）：本模型讀哪一層、哪一版、哪些名稱；自層級讀入的值只列在此節
  C. 層級隱含值推導：Wattonomics 配置給 Anthropic 的 GW（T1，營運者→使用者）× Tokenomics 每 GW 營收
  D. C0（實績）：公司實績（S-1/A 季度損益、現金流量表）vs 層級隱含值；主讀數＝2026 前九月累計 vs 層級 CY2026 × 9/12，
     第二讀數＝2026 Q3 × 4 vs 層級 CY2026；兩者都超過 ±15% 且同號才「對帳中」，只一個超過為「時差邊界（參考）」
  E. 層級之間（參考）：Wattonomics 內部兩個 Anthropic 算力讀數；不計入 C0
  F. C0 結論與判定摘要
層級值來自快照檔（Tokenomics 走 TK_Link 具名範圍；Wattonomics 讀 wattonomics/data 的快照 JSON，值寫成藍字並附快照版本與雜湊），
公式只引用具名範圍與本頁儲存格，不含常數（判定門檻、季化倍數為本頁藍字參數）。
"""
import json
from pathlib import Path

from common import F_BOLD, F_CALC, F_IN, F_NOTE, FILL_SEC, header, nm, put, title

ROOT = Path(__file__).resolve().parent.parent
WATT_DIR = ROOT.parent / "wattonomics" / "data"


def read_watt():
    cur = (WATT_DIR / "CURRENT_SNAPSHOT").read_text(encoding="utf-8").strip()
    d = json.loads((WATT_DIR / cur).read_text(encoding="utf-8"))
    rows = {r["id"]: r for s in d["rows"] for r in ([s] if "id" in s else [])}
    return dict(file=d["source"]["file"], version=d["source"]["version"], sha=d["source"]["sha256"], rows=rows, names=d["names"], snap=cur)


def build(ctx):
    D, wb = ctx.D, ctx.wb
    S, I = D.S, D.I
    W = read_watt()
    R, N = W["rows"], W["names"]
    ws = wb.create_sheet("Coord")
    title(ws, "Coord — 三層協調（v0.2.2；規則 v1.4.2）：讀取層級、C0（實績）公司實績 vs 層級隱含值",
          "C0 比較基準＝公司實績或法定揭露（S-1/A 季度損益、現金流量表）；層級側＝Tokenomics 每 GW 營收 × Wattonomics 配置給 Anthropic 的 GW（兩者都不讀本模型的輸出）。"
          "差距超過 ±15% 且兩讀數同號標「對帳中」（開對帳單、不改上游）；只有一個讀數超過標「時差邊界（參考）」。自層級讀入的同源值只列在 B 節，不判定。",
          "本頁只列示，不被任何計算頁引用（Checks 除外）。單位：$B、GW。藍字＝快照值或參數，黑字＝公式。")
    P = {}
    r = [5]

    def sec(t):
        put(ws, f"A{r[0]}", t, F_BOLD, fill=FILL_SEC)
        r[0] += 1

    def row(key, label, vals, unit="", note="", fmt="#,##0.000", cols="CDEFGHIJK", fonts=None):
        n = r[0]
        put(ws, f"A{n}", label, F_CALC)
        put(ws, f"B{n}", unit, F_CALC)
        for i, (c, v) in enumerate(zip(cols, vals)):
            if v is None:
                continue
            is_f = isinstance(v, str) and v.startswith("=")
            put(ws, f"{c}{n}", v, (F_CALC if is_f else F_IN) if not (isinstance(v, str) and not is_f) else F_CALC,
                fmt=fmt if not (isinstance(v, str) and not is_f) else None)
        if note:
            put(ws, f"L{n}", note, F_NOTE)
        P[key] = n
        r[0] += 1
        return n

    # ── A. 參數 ──
    sec("A. 參數（Decision；規則 v1.4.2 C0 第 2 點與 tools/coord/README.md）")
    header(ws, r[0], ["項目", "單位", "值", "", "", "", "", "", "", "", "", "說明"])
    r[0] += 1
    row("th", "差距門檻（對帳中）", [0.15], "比例", "規則 v1.4.2：差距超過 ±15%，且兩個讀數同號，標「對帳中」", fmt="0%")
    row("qx", "季化倍數（最近一季 × 本值＝年化速率）", [4], "倍", "第二讀數＝2026 Q3 × 本值 vs 層級 CY2026", fmt="0")
    row("m9", "2026 已公布月數", [f"={I('m_9m')}"], "月", "Inputs 定義常數；S-1/A 期中財報涵蓋 2026-01～09", fmt="0")
    row("m12", "每年月數", [f"={I('months_per_year')}"], "月", "Inputs 定義常數", fmt="0")
    row("k2", "平均除數", [f"={I('arith_k2')}"], "—", "Inputs 定義常數；年底值平均＝（期初＋期末）÷ 本值", fmt="0")
    TH, QX, M9, M12, K2 = (f"$C${P[k]}" for k in ("th", "qx", "m9", "m12", "k2"))

    # ── B. 讀取層級 ──
    r[0] += 1
    sec("B. 讀取層級（同源，不判定）：本模型實際讀的層級、版本與名稱；自層級讀入的值只列在此節，不進 C0")
    header(ws, r[0], ["項目", "單位", "本模型讀入值", "層級現值（快照）", "差", "", "", "", "", "", "", "說明"])
    r[0] += 1
    row("tkv", "Tokenomics 版本（TK_Link 快照）", ["=TK_Version"], "", "本模型讀 Tokenomics 快照 v21.10（TK_Link；108 個具名範圍＋21 格讀表）；Tokenomics master 現行為 v22.0，本輪不換快照（工作單：模型仍讀舊快照者寫明版本、不在本輪換）。"
        "TK 值只作算力物理與對照，不回饋 TK", fmt="@", cols="C")
    row("tkc", "TK_Link 已取值名稱數", ['=COUNTIF(TK_Link!$F$10:$F$200,"OK")'], "個", "計算頁引用 TK_ 具名範圍（每 GW 產出、持有成本、單價、NonNV 產出比、PUE 等）；本頁 C 節另讀 L1_RevGW_PQ_VR200 三值（低／基準／高）", fmt="0")
    row("wtv", "Wattonomics 快照（本頁 C、E 節取值）", [f"{W['version']}｜{W['file']}"], "", f"快照檔 wattonomics/data/{W['snap']}；SHA-256 {W['sha']}", cols="C")
    row("trace", "Wattonomics Anthropic GW 與資本支出路徑追溯（C0 前置檢查）", ["已逐一追溯，層級 GW 與資本支出不讀本模型輸出"], "",
        "追溯：(1) T1 配置（al_amzn_ant、al_googl_ant、al_neo_ant）為 Wattonomics In 頁的輸入列，非公式，只進 Matrix（營運者×使用者）；AWS 列取 Epoch 衛星估計 New Carlisle（第三方觀測）、Google 列取公司公告「well over a gigawatt」2026 上線（公司揭露的規劃容量，Interested-party，非實績）、新雲列取 TeraWulf、Cipher、Hut 8 租約（公司揭露）；"
        "(2) 玩家路徑（Players、CapexGW）：Anthropic 自建年底 GW＝0（無證據），不讀本模型；(3) Gap 頁只匯入本模型的營業現金流與期初現金（Y9，影響資金缺口，不影響 GW、不影響本頁營收隱含值）；"
        "CapexAIModel 回寫值＝0；(4) 唯一讀本模型的 GW 相關值是 Y4 usr_ant 2025（1.4，取自本模型 Compute），只作 Matrix 對照列，本頁不用；"
        "(5) 未發現 w1_*、obs_*、Players 覆寫讀本模型。結論：營收列不是同源；Google 1 GW 為公司公告的規劃容量，屬公司揭露而非本模型輸出，已在 C 節說明。", cols="C")
    row("wtm", "本模型讀 Wattonomics 的版本", ["v0.1.7（Inputs 藍字，非即時連結）"], "",
        "own_gw_add_2025–2030（自建 IT GW 年新增）與 own_lag_bldg 在 v0.1.7 取 Wattonomics 快照值寫入 Inputs；之後未重讀。下列逐年與 v0.3.2 快照對照（同源，不判定）", cols="C")
    for y in range(2025, 2031):
        row(f"own{y}", f"自建 IT GW 年新增 {y}：Inputs 值 vs Wattonomics WIF_Gap_Player_ant_{y}_CapexAIGW",
            [f"={I(f'own_gw_add_{y}')}", N.get(f"WIF_Gap_Player_ant_{y}_CapexAIGW"), f"=C{r[0]}-D{r[0]}"], "GW", "同源（本模型取自 Wattonomics）；差為 0＝未落後" if y == 2025 else "")
    gocf_ref = {}
    for y in range(2026, 2031):
        idx = f"FND_FCF"
        i_ = y - 2025
        n = r[0]
        row(f"ocf{y}", f"營業現金流 {y}：本模型 FND_FCF 現值 vs Wattonomics Gap 匯入值（WIF_Gap_Player_ant_{y}_OCF）",
            [f"=INDEX(FND_FCF,1,{i_ + 1})", N.get(f"WIF_Gap_Player_ant_{y}_OCF"), f"=C{n}-D{n}"], "$B",
            ("Wattonomics 的 Anthropic 營業現金流由本模型匯入（交接檔：Funding F07＋F03；自建資本支出 F03＝0）；匯入值停在 v0.1.8，與本模型現值不同＝層級匯入落後，"
             "不是證據分歧；待 Wattonomics 下次重匯更新。同源，不判定、不計入對帳中" if y == 2026 else ""), cols="CDE")

    # ── C. 層級隱含值推導 ──
    r[0] += 1
    sec("C. 層級隱含值推導：Wattonomics 配置給 Anthropic 的年底 GW（T1 營運者→使用者）× Tokenomics 每 GW 營收")
    header(ws, r[0], ["項目", "單位", "2025 年底", "2026 年底", "", "", "", "", "", "", "", "說明"])
    r[0] += 1

    def t1(rid):
        v = R[rid]["values"]
        return v[3], v[4]            # K＝2025 年底、L＝2026 年底（空白＝層級未計入，以 0 計）
    z = lambda x: 0 if x is None else x  # noqa: E731
    for rid, lab in (("al_amzn_ant", "Amazon（AWS；Trainium＋NVIDIA）→ Anthropic"), ("al_googl_ant", "Alphabet（Google TPU）→ Anthropic"),
                     ("al_neo_ant", "其他新雲合計 → Anthropic")):
        a, b = t1(rid)
        row(rid, f"Wattonomics T1：{lab}", [z(a), z(b)], "GW",
            f"{rid}（{R[rid]['tag']}；{R[rid]['src']}）：{R[rid]['note'][:140]}；空白＝層級未計入，以 0 計", cols="CD")
    n0, n1 = P["al_amzn_ant"], P["al_neo_ant"]
    row("gw_ye", "配置給 Anthropic 的 GW 合計（年底，IT）", [f"=SUM(C{n0}:C{n1})", f"=SUM(D{n0}:D{n1})"], "GW", "Wattonomics 另有「Anthropic 自建」玩家，GW＝0（WIF_PlayerGW_YE_ant_2026）；Anthropic 的算力幾乎全為租用，故用 T1 配置 GW", cols="CD")
    row("gw_avg", "2026 年平均 GW（＝2025 年底與 2026 年底的簡單平均）", [None, f"=(C{P['gw_ye']}+D{P['gw_ye']})/{K2}"], "GW",
        "[Derived] 線性近似；年內上線時點未知，成長年的年平均可能偏低或偏高", cols="CD")
    GWAVG = f"$D${P['gw_avg']}"
    row("pq", "Tokenomics L1_RevGW_PQ_VR200：每 GW 年營收（第一年；低／基準／高）", ["=TK_L1_RevGW_PQ_VR200_Lo", "=TK_L1_RevGW_PQ_VR200", "=TK_L1_RevGW_PQ_VR200_Hi"], "$B/GW/年",
        "TK v21.10；VR200 口徑（Q 取實際利用率、P 取需求面有效單價、m 取 2026 Q2 年化）；Anthropic 機隊含 Trainium／TPU／NVIDIA 多世代，本表不做口徑換算（換算需本模型的產出比，會變成同源）；"
        "L1_M 取 OpenAI 當期營收與牌價校準（Tokenomics L1），非 Anthropic 值、非本模型輸出，故不同源。Tokenomics 自註：Q 是模型產能推算，若 token 帳誤差全在每 GW 產能側，實際營收約為本列 ÷ 2.0–2.1（OpenAI 2026 口徑；Anthropic 未必相同）——此時層級隱含值減半，主讀數差距轉正、第二讀數差距轉大正，判定會改變，故本列只作參考性檢驗", cols="CDE")
    row("cy", "層級隱含 CY2026 營收（每 GW 營收 × 2026 年平均 GW；低／基準／高）",
        [f"=C{P['pq']}*{GWAVG}", f"=D{P['pq']}*{GWAVG}", f"=E{P['pq']}*{GWAVG}"], "$B", "層級隱含值，連到上方儲存格；不含本模型任何輸出", cols="CDE")

    # ── D. C0（實績） ──
    r[0] += 1
    sec("D. C0（實績）：公司實績 vs 層級隱含值（規則 v1.4.2 C0 第 2 點；Andy 11-14 17:40 固定規格）")
    header(ws, r[0], ["項目", "單位", "公司實績（主：2026 前九月累計）", "層級隱含（主：CY2026 × 9/12）", "差距（主）", "公司實績（第二：Q3）", "層級隱含（第二：CY2026）",
                      "差距（第二；Q3 × 4 vs CY2026）", "判定", "出處與口徑", "等級", "說明"])
    r[0] += 1
    d1 = r[0]
    n = d1
    row("c0_rev", "營收（GAAP；總額，含雲端市集銷售）",
        [f"={S('pros_rev_9m26')}", f"=$D${P['cy']}*{M9}/{M12}", f"=C{n}/D{n}-1", f"={S('rev_s1a_2026q3')}", f"=$D${P['cy']}", f"=F{n}*{QX}/G{n}-1",
         f'=IF(AND(ISNUMBER(E{n}),ISNUMBER(H{n})),IF(AND(ABS(E{n})>{TH},ABS(H{n})>{TH},SIGN(E{n})=SIGN(H{n})),"對帳中",IF(OR(ABS(E{n})>{TH},ABS(H{n})>{TH}),"時差邊界（參考）","一致")),"無可比實績")',
         "Anthropic S-1/A（Amendment No. 1，2026-11-02）季度損益：2026 前九月 33.242、Q3 16.956（SRC_ANT_547、488）；GAAP 總額",
         "Verified（一手；期中未經查核）",
         "偏誤方向依期間中點：主讀數的實績期間（1–9 月）中點在 5 月中，早於全年中點（7 月 1 日），營收上升時 1–9 月累計占全年比例低於 9/12，主讀數會偏負；第二讀數 Q3 × 4 是 8 月中的年化速率，晚於全年中點，上升時會偏正。兩者方向相反是機械結果，不得讀成「夾住真值」。層級 GW 為年底值的簡單平均，上線時點未知。判定依規則：兩者都超過且同號才對帳中。實績不含公司指引與 run-rate 新聞值（只列參考，見 SRC_ANT_023）"],
        "$B", fmt="#,##0.000", cols="CDEFGHIJKL")
    for c in "EH":
        ws[f"{c}{n}"].number_format = "0.0%"
    n = r[0]
    row("c0_capex", "自建資本支出（購置不動產及設備；現金流量表）",
        [f"={S('capex_s1a_2026_9m')}", N.get("WIF_Gap_Player_ant_2026_CapexAI") or 0, None, None, None, None,
         f'=IF(D{n}=0,"無法判定（層級值為 0，無百分比差距）","一致")',
         "Anthropic S-1/A 現金流量表：2026 前九月購置不動產及設備 0.925（SRC_ANT_519）；層級＝Wattonomics WIF_Gap_Player_ant_2026_CapexAI（自建 AI 資本支出，快照值）",
         "Verified（一手；期中未經查核）",
         "Wattonomics 認定 Anthropic 無自建 GW（租用為主），自建 AI 資本支出為 0；公司實績 0.925 為辦公與小型設施等，與自建 AI 算力不同口徑，數量級遠小於每 GW 建置成本（TK IF_CapexTotal 約 31 以上），"
         "故只列示、不計入對帳中；層級值為 0 時差距百分比無定義"],
        "$B", cols="CDEFGHIJKL")
    n = r[0]
    row("c0_commit", "算力承諾（S-1/A 合約義務金額）vs Wattonomics 合約 GW",
        [None, None, None, None, None, None, "無可比實績（口徑：美元 vs GW）",
         "S-1/A 承諾以美元揭露（SRC_ANT_410–464、488–521 承諾三段）；Wattonomics T1 以 GW 揭露", "—",
         "缺：層級每 GW 合約價（Tokenomics 無合約價；以本模型通用合約價換算屬本模型輸出，會變成同源）。公司揭露的 GW 數（Google 1 GW 以上、AWS 容量）已是 Wattonomics T1 的來源（Interested-party），同源，不作實績"],
        "", cols="CDEFGHIJKL")
    d3 = r[0] - 1
    DR = f"$I${d1}:$I${d3}"
    nm(wb, "COORD_C0", f"Coord!$A${d1}:$K${d3}")

    # ── F. 判定摘要（命名範圍供 HTML 與 Checks 取值；放在 E 節之前以固定位置） ──
    r[0] += 1
    sec("判定摘要（只計 D 節）")
    n = r[0]
    put(ws, f"A{n}", "可判定項目數（有數值差距）", F_CALC)
    put(ws, f"C{n}", f"=COUNT(E{d1}:E{d3})", fmt="0")
    put(ws, f"A{n + 1}", "對帳中項目數", F_CALC)
    put(ws, f"C{n + 1}", f'=COUNTIF({DR},"對帳中")', fmt="0")
    put(ws, f"A{n + 2}", "時差邊界（參考）項目數", F_CALC)
    put(ws, f"C{n + 2}", f'=COUNTIF({DR},"時差邊界（參考）")', fmt="0")
    put(ws, f"A{n + 3}", "無可比實績／無法判定項目數", F_CALC)
    put(ws, f"C{n + 3}", f'=COUNTIF({DR},"無可比實績*")+COUNTIF({DR},"無法判定*")', fmt="0")
    nm(wb, "COORD_C0Sum", f"Coord!$C${n}:$C${n + 3}")
    P["sum"] = n
    r[0] = n + 4
    put(ws, f"A{r[0]}", "C0 結論", F_BOLD)
    put(ws, f"C{r[0]}", ("營收（唯一可判定的實績項）：主讀數偏負（1–9 月期間中點早於全年中點）、第二讀數接近零，兩者未同時超過門檻，判為「時差邊界（參考）」，不開對帳單、不改 Tokenomics 或 Wattonomics。"
                         "此結論有兩項限制：(1) 層級 GW 取年底值簡單平均；(2) Tokenomics 自註每 GW 營收可能高估約一倍（誤差全在產能側時），若成立則第二讀數會轉為明顯正差距，故本列是量級檢驗，不是精確對帳。"
                         "資本支出（層級值為 0）與算力承諾（美元 vs GW 口徑）無可比實績。往上回饋只有證據（C5）：待 2026 Q4 與 2027 年報的營收與算力揭露，Q4 實績出來後可重判。"), F_CALC)
    nm(wb, "COORD_C0Verdict", f"Coord!$C${r[0]}")
    P["verdict"] = r[0]
    r[0] += 1

    # ── E. 層級之間（參考） ──
    r[0] += 1
    sec("E. 層級之間（參考；規則 v1.4.2 C0：兩層之間的差距只列參考，不計入個股 C0）")
    header(ws, r[0], ["項目", "單位", "數值", "對照", "比值", "", "", "", "", "", "", "說明"])
    r[0] += 1
    n = r[0]
    u = R["usr_ant"]["values"]        # Y4：2023 起（2025＝索引 2、2026＝索引 3、2027＝索引 4）
    row("usr26", "Wattonomics 報導值：Anthropic 使用算力年底 GW（2026）vs 本頁 C 節配置 GW（T1，2026 年底）", [u[3], f"=D{P['gw_ye']}", f"=C{n}/D{n}"], "GW",
        f"Y4 usr_ant（{R['usr_ant']['tag']}；NYT 轉述向投資人說明的目標，非實績）；T1 為已簽約且有證據的營運者配置。兩者都在 Wattonomics 內，差距屬層級內部；不影響 C0 判定", cols="CDE")
    ws[f"E{n}"].number_format = "0.00"
    row("usr25", "Wattonomics 報導值：Anthropic 使用算力年底 GW（2025）vs 本頁 C 節配置 GW（T1，2025 年底）", [u[2], f"=C{P['gw_ye']}", f"=C{n + 1}/D{n + 1}"], "GW",
        "2025 值（1.4）取自 Anthropic 模型 Compute（Wattonomics 備註），同源，只列參考", cols="CDE")
    ws[f"E{n + 1}"].number_format = "0.00"
    P["lvl"] = n
    for c, w in zip("ABCDEFGHIJKL", (62, 9, 22, 22, 12, 20, 20, 18, 26, 60, 22, 90)):
        ws.column_dimensions[c].width = w
    ws.freeze_panes = "C5"
    ctx.P["a13"] = dict(P=P)


def checks(ctx):
    P = ctx.P["a13"]["P"]
    n = P["c0_rev"]
    return [
        ("coord_c0_num", "Coord：C0 營收主、第二讀數差距皆為數值", f"=--AND(ISNUMBER(Coord!$E${n}),ISNUMBER(Coord!$H${n}))", 1, "eq", "三層協調（v0.2.2）；只列示，不回饋計算"),
        ("coord_c0_nobad", "Coord：C0（實績）對帳中項目數（資訊）", "=INDEX(COORD_C0Sum,2)", None, "info", "對帳中＝兩讀數都超過 ±15% 且同號；開對帳單、不改上游"),
    ]
