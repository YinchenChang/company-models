"""智譜 v0.1.5：PL 頁（標準損益表；個股 HTML 版面規格 v1.3 D 節）與 Coord 頁（三層協調：讀取層級、C0 實績；規則 v1.4.2）。

兩頁只列示，不被任何既有計算頁引用；新增後既有工作表數值不變（tools/xlsx_diff.py 逐格驗證）。
- PL：FY2023–FY2025 取 SRC_ZP（招股章程、2025 年度業績公告；一手），FY2026–FY2030 取本模型基準情境（Cost、Funding、Revenue 具名範圍）。
  口徑＝非 IFRS（經調整）：營業費用扣股權報酬；稅後淨利歷史＝公司揭露經調整淨虧損。歷史的稅前與所得稅公司未分開揭露（年度公告摘要），標「未揭露」；
  預測期本模型不計稅，所得稅標「未建模」，稅後淨利＝稅前。
- Coord：讀取層級（Tokenomics 快照、Wattonomics 不讀、Siliconomics 不適用）；C0（實績）公司揭露 vs 層級隱含值。Wattonomics 玩家名單沒有智譜（中國只有合計列），
  各列層級側無值，判定「無可比實績（層級無此玩家）」。
層級值只來自 repo 內快照檔（wattonomics/data），寫成藍字並附版本與雜湊；公式只引用具名範圍與本頁儲存格。
"""
from __future__ import annotations

import json
from pathlib import Path

from common import F_BOLD, F_CALC, F_IN, F_NOTE, FILL_SEC, header, put, title
from openpyxl.workbook.defined_name import DefinedName

ROOT = Path(__file__).resolve().parent.parent
WATT_DIR = ROOT.parent / "wattonomics" / "data"
YEARS = list(range(2023, 2031))
COLS = list("CDEFGHIJ")                 # FY2023–FY2030
MCOL = dict(zip(range(2025, 2031), "DEFGHI"))   # 計算頁 D–I＝2025–2030
NOT_DISC, NOT_MODEL, NA = "未揭露", "未建模", "—"


def nm(wb, name, ref):
    wb.defined_names[name] = DefinedName(name, attr_text=ref)


def read_watt():
    cur = (WATT_DIR / "CURRENT_SNAPSHOT").read_text(encoding="utf-8").strip()
    d = json.loads((WATT_DIR / cur).read_text(encoding="utf-8"))
    players = sorted({k[len("WIF_PlayerGW_YE_"):].rsplit("_", 1)[0] for k in d["names"]
                      if k.startswith("WIF_PlayerGW_YE_") and "_SiteOnly_" not in k})
    blob = json.dumps(d, ensure_ascii=False)
    return dict(file=d["source"]["file"], version=d["source"]["version"], sha=d["source"]["sha256"], snap=cur, names=d["names"], players=players,
                hits=sum(blob.count(w) for w in ("Zhipu", "zhipu", "智譜", "智谱", "Z.ai", "GLM")))


# ───────────────────────── PL ─────────────────────────
def build_pl(wb, D, Z):
    S, I = D.S, D.I
    F = Z["F"].rows
    X = Z["X"].rows
    ws = wb.create_sheet("PL")
    title(ws, "PL — 標準損益表（非 IFRS 經調整；FY2023–FY2030；人民幣億元）",
          "歷史取 SRC_ZP（招股章程、2025 年度業績公告；一手）；預測取本模型基準情境（Cost、Funding）。公司未分開揭露者標「未揭露」，本模型未建者標「未建模」，不補編。",
          "只列示，不被任何計算頁引用。黑字＝公式。預測營業利益＝Cost 差額（含股權報酬）＋股權報酬；稅前＝營業利益＋其他收益 − 利息（＝Funding 自由現金流，核對列見下）。")
    header(ws, 5, ["項目", "單位"] + [f"FY{y}{'（預測）' if y >= 2026 else ''}" for y in YEARS] + ["說明"])
    R = {}
    r = [6]
    fmt = "#,##0.00"

    def row(key, label, vals, unit="RMB 億", note="", f=fmt, bold=False):
        n = r[0]
        put(ws, f"A{n}", label, F_BOLD if bold else F_CALC)
        put(ws, f"B{n}", unit, F_CALC)
        for c, v in zip(COLS, vals):
            if v is None:
                continue
            is_f = isinstance(v, str) and v.startswith("=")
            put(ws, f"{c}{n}", v, None, fmt=None if (isinstance(v, str) and not is_f) else f)
        if note:
            put(ws, f"K{n}", note, F_NOTE)
        R[key] = n
        r[0] += 1
        return n

    H = {2023: "fy23", 2024: "fy24", 2025: "fy25"}
    fc = list(range(2026, 2031))
    ix = lambda name, y: f"INDEX({name},1,{y - 2024})"          # noqa: E731  D–I＝2025–2030 → 索引 1–6
    fnd = lambda key, y: f"Funding!${MCOL[y]}${F[key]}"           # noqa: E731
    row("rev", "營收", [f"={S('rev_total_' + H[y])}" for y in H] + [f"={ix('REV_NetCapped', y)}" for y in fc],
        note="歷史：SRC_ZP_002／003／120（營收總額）；預測：營收淨額（截頂後；總額＝淨額，D5）", bold=True)
    n = R["rev"]
    row("yoy", "營收年增率（YoY）", [NA] + [f"={c}{n}/{p}{n}-1" for p, c in zip("CDEFGHI", "DEFGHIJ")], "比例", f="0.0%",
        note="FY2023 的前一年營收在 SRC 只有招股章程兩分部口徑，未登錄總額，留空")
    row("cogs", "營業成本（銷售成本）", [f"={S('cogs_' + H[y])}" for y in H] + [f"={ix('COST_InfCompute', y)}+{ix('COST_OnPremDelivery', y)}" for y in fc],
        note="歷史：SRC_ZP_017／018／123；預測：推論算力成本＋本地化部署交付成本（現金口徑）")
    c_ = R["cogs"]
    row("gp", "毛利", [f"={c}{n}-{c}{c_}" for c in COLS], bold=True)
    g_ = R["gp"]
    row("gpm", "毛利率", [f"={c}{g_}/{c}{n}" for c in COLS], "比例", f="0.0%")
    row("opex", "營業費用（不含股權報酬）",
        [f"={S('rd_' + H[y])}+{S('sm_' + H[y])}+{S('ga_' + H[y])}-{S('sbc_' + H[y])}" for y in H]
        + [f"={ix('COST_Compute', y)}-{ix('COST_InfCompute', y)}+{ix('COST_NonCompExSBC', y)}" for y in fc],
        note="歷史：研發＋銷售＋管理 − 股權報酬（SRC_ZP rd／sm／ga／sbc）；預測：研發算力＋閒置算力（＝算力成本 − 推論算力；Cost K07＋K08）＋非算力成本（不含股權報酬）")
    o_ = R["opex"]
    row("oi", "營業利益（非 IFRS）", [f"={c}{g_}-{c}{o_}" for c in COLS], bold=True, note="負＝虧損")
    i_ = R["oi"]
    row("oim", "營業利益率", [f"={c}{i_}/{c}{n}" for c in COLS], "比例", f="0.0%")
    row("oth", "其他收入（淨額）", [f"={S('other_inc_' + H[y])}-{S('fin_cost_' + H[y])}" for y in H] + [f"={fnd('oth', y)}-{fnd('int', y)}" for y in fc],
        note="歷史：其他收入 − 財務成本（SRC_ZP other_inc／fin_cost）；預測：Funding 其他收益 − 利息支出")
    t_ = R["oth"]
    row("pbt", "稅前利益（非 IFRS）", [NOT_DISC] * 3 + [f"={c}{i_}+{c}{t_}" for c in "FGHIJ"], bold=True,
        note="歷史：公司年度公告摘要未分開揭露稅前與所得稅（另有其他損益、減值等未拆項目，見下方未拆列）；預測：營業利益＋其他收入淨額")
    p_ = R["pbt"]
    row("tax", "所得稅", [NOT_DISC] * 3 + [NOT_MODEL] * 5, note="預測：本模型不計所得稅（虧損期）；稅後淨利＝稅前")
    row("ni", "稅後淨利（非 IFRS 經調整）", [f"=-{S('adj_loss_' + H[y])}" for y in H] + [f"={c}{p_}" for c in "FGHIJ"], bold=True,
        note="歷史：公司揭露經調整淨虧損取負號（SRC_ZP_082／083／146）；預測＝稅前（不計稅）")
    ni = R["ni"]
    row("nim", "淨利率", [f"={c}{ni}/{c}{n}" for c in COLS], "比例", f="0.0%")
    row("eps", "EPS（非 IFRS，基本；人民幣元）", [NOT_DISC] * 3 + [f"={c}{ni}*{I('div_yuan_yi')}/{S('shares_total_post_placing2')}" for c in "FGHIJ"], "元／股",
        note="歷史：非 IFRS 每股未揭露；預測＝稅後淨利 ×（億 → 元）÷ 2026-09-16 配售後總股本（SRC_ZP_563，含非上市股份）；本模型未建未來增發股數，未計外部資金需求的稀釋與可換股債券轉股", f="#,##0.00")
    row("niyoy", "淨利年增率", [NA] + [f"={c}{ni}/{p}{ni}-1" for p, c in zip("CDEFGHI", "DEFGHIJ")], "比例", f="0.0%",
        note="虧損期：正＝虧損擴大（分母為負）")
    r[0] += 1
    put(ws, f"A{r[0]}", "核對與未拆項目（只列示）", F_BOLD, fill=FILL_SEC)
    r[0] += 1
    row("plug", "未拆項目（歷史）＝經調整淨利 − 營業利益 − 其他收入淨額（含其他損益、減值、所得稅等摘要未拆者）",
        [f"={c}{ni}-{c}{i_}-{c}{t_}" for c in "CDE"], note="[Derived] 年度公告摘要未拆的項目合計；不是模型輸入")
    row("chk_fc", "核對：稅前（預測）− Funding 自由現金流（應為 0）", [None, None, None] + [f"={c}{p_}-{ix('FND_FCF', y)}" for c, y in zip("FGHIJ", fc)])
    row("chk_oi", "核對：營業利益（預測）− Cost 差額（含股權報酬）− 股權報酬（應為 0）",
        [None, None, None] + [f"=ROUND({c}{i_}-{ix('COST_GapCash', y)}-{ix('COST_SBC', y)},9)" for c, y in zip("FGHIJ", fc)])
    row("chk_25", "核對：FY2025 營收、營業成本、營業費用合計 − 本模型 2025 校準值（Cost 全成本 − 股權報酬）（應為 0）",
        [None, None, f"=E{c_}+E{o_}-({ix('COST_FullCash', 2025)}-{ix('COST_SBC', 2025)})"])
    for nm_, key in (("PL_Rev", "rev"), ("PL_RevYoY", "yoy"), ("PL_COGS", "cogs"), ("PL_GP", "gp"), ("PL_GPM", "gpm"), ("PL_Opex", "opex"), ("PL_OI", "oi"),
                     ("PL_OIM", "oim"), ("PL_Oth", "oth"), ("PL_PBT", "pbt"), ("PL_Tax", "tax"), ("PL_NI", "ni"), ("PL_NIM", "nim"), ("PL_EPS", "eps"), ("PL_NIYoY", "niyoy")):
        nm(wb, nm_, f"PL!$C${R[key]}:$J${R[key]}")
    nm(wb, "PL_Plug", f"PL!$C${R['plug']}:$E${R['plug']}")
    nm(wb, "PL_ChkFc", f"PL!$F${R['chk_fc']}:$J${R['chk_oi']}")
    nm(wb, "PL_Chk25", f"PL!$E${R['chk_25']}")
    ws.column_dimensions["A"].width = 64
    ws.column_dimensions["B"].width = 9
    for c in COLS:
        ws.column_dimensions[c].width = 13
    ws.column_dimensions["K"].width = 90
    ws.freeze_panes = "C6"
    return dict(R=R)


# ───────────────────────── Coord ─────────────────────────
def build_coord(wb, D, Z, snap):
    S = D.S
    W = read_watt()
    ws = wb.create_sheet("Coord")
    title(ws, "Coord — 三層協調（v0.1.5；規則 v1.4.2）：讀取層級、C0（實績）公司實績 vs 層級隱含值",
          "C0 比較基準＝公司法定揭露（港交所招股章程、2025 年度與 2026 中期業績公告）；層級側＝同口徑層級隱含值（算力買方：Tokenomics 每 GW 營收 × Wattonomics 該玩家 GW、"
          "Wattonomics 該玩家資本支出與上線 GW）。Wattonomics 玩家名單沒有智譜（中國只有合計列），各列層級側無值，判定「無可比實績（層級無此玩家）」。",
          "本頁只列示，不被任何計算頁引用。單位：人民幣億元、GW。藍字＝快照值或參數，黑字＝公式。")
    P = {}
    r = [5]

    def sec(t):
        put(ws, f"A{r[0]}", t, F_BOLD, fill=FILL_SEC)
        r[0] += 1

    def row(key, label, vals, unit="", note="", fmt="#,##0.00", cols="CDEFGHIJK"):
        n = r[0]
        put(ws, f"A{n}", label, F_CALC)
        put(ws, f"B{n}", unit, F_CALC)
        for c, v in zip(cols, vals):
            if v is None:
                continue
            is_f = isinstance(v, str) and v.startswith("=")
            is_txt = isinstance(v, str) and not is_f
            put(ws, f"{c}{n}", v, None if (is_f or is_txt) else F_IN, fmt=None if is_txt else fmt)
        if note:
            put(ws, f"L{n}", note, F_NOTE)
        P[key] = n
        r[0] += 1
        return n

    sec("A. 參數（Decision；規則 v1.4.2 C0 第 2 點與 tools/coord/README.md）")
    header(ws, r[0], ["項目", "單位", "值", "", "", "", "", "", "", "", "", "說明"])
    r[0] += 1
    row("th", "差距門檻（對帳中）", [0.15], "比例", "差距超過 ±15%，且兩個讀數同號，才標「對帳中」", fmt="0%")
    row("qx", "季化倍數（最近一季 × 本值＝年化速率）", [4], "倍", "第二讀數＝最近一季 × 本值 vs 層級 CY2026", fmt="0")
    row("m6", "2026 已公布月數", [6], "月", "2026 中期業績公告（2026-08-31）涵蓋 2026-01～06；公司未發布季度業績（季營收未揭露，第二讀數無法取得）", fmt="0")
    row("m12", "每年月數", [12], "月", "時差換算", fmt="0")

    r[0] += 1
    sec("B. 讀取層級（同源，不判定）：本模型實際讀的層級、版本與名稱；自層級讀入的值只列在此節，不進 C0")
    header(ws, r[0], ["項目", "單位", "內容", "", "", "", "", "", "", "", "", "說明"])
    r[0] += 1
    row("tkv", "Tokenomics（TK_Link 快照）", [f"{snap['version']}｜{snap['file']}｜提交 {snap['sha'][:7]}"], "",
        "TK_Link 頁 77 個具名範圍＋7 列讀表（README）。Tokenomics master 現行為 v22.0；本輪不換快照（工作單：模型仍讀舊快照者寫明版本、不在本輪換）。TK 值只作算力物理與對照，不回饋 Tokenomics", cols="C")
    row("wtv", "Wattonomics", [f"本模型不讀 Wattonomics；本頁對照用快照 {W['version']}｜{W['file']}"], "",
        f"快照檔 wattonomics/data/{W['snap']}（SHA-256 {W['sha']}）。Wattonomics 玩家名單（年底上線 GW 的玩家代號）：{'、'.join(W['players'])}；沒有智譜，中國只有合計列 cn（Rystad 容量 × AI 占比）", cols="C")
    row("wtr", "Wattonomics 是否讀本模型", [f"否（快照全文出現智譜／Zhipu／Z.ai／GLM 字樣 {W['hits']} 次）"], "", "同源檢查：層級沒有讀入本模型的任何值，本頁不存在同源校準列", cols="C")
    row("sif", "Siliconomics", ["不適用（算力買方，非硬體供應鏈個股）"], "", "算力買賣方讀 Tokenomics／Wattonomics", cols="C")

    r[0] += 1
    sec("C. C0（實績）：公司實績 vs 層級隱含值（規則 v1.4.2 C0 第 2 點；Andy 11-14 17:40 固定規格）")
    header(ws, r[0], ["項目", "單位", "公司實績（主）", "層級隱含（主）", "差距（主）", "公司實績（第二）", "層級隱含（第二）", "差距（第二）", "判定", "出處與口徑", "等級", "說明"])
    r[0] += 1
    d1 = r[0]
    NAV = "無可比實績（層級無此玩家）"
    row("c0_rev", "營收（營收總額；2026 已公布 1–6 月；第二讀數＝最近一季，未揭露）",
        [f"={S('rev_total_1h26')}", None, NA, None, None, NA, NAV,
         "SRC_ZP_177（1H2026 營收總額 9.5389；2026 中期業績公告 2026-08-31，港交所 2026083101539）；雲端部署（開放平台及 API）8.2518（SRC_ZP_179），本地化部署 1.2872（SRC_ZP_178）",
         "Verified（一手）",
         "層級隱含＝Tokenomics 每 GW 營收 × Wattonomics 智譜 GW；Wattonomics 沒有智譜玩家（只有中國合計），層級側無值。缺：Wattonomics 的智譜上線 GW（或中國合計拆到公司）。"
         "公司未發布季度營收，第二讀數（Q2 × 4）無法取得。2026-08 MaaS ARR 16–20 億美元、2026-09 全業務 ARR 18 億美元（SRC_ZP_417–419）屬 run-rate、業績會與媒體轉述，不入本列"],
        "RMB 億", cols="CDEFGHIJKL")
    row("c0_capex", "資本支出（1H2026；主要為收購北京紅鑽，非 AI 算力）",
        [f"={S('capex_1h26')}", None, NA, None, None, None, NAV,
         "SRC_ZP_209（1H2026 資本開支 3.906；2026 中期業績公告）；公司稱主因收購北京紅鑽（研究筆記 financials_notes 第 2 節）",
         "Verified（一手）",
         "層級對象＝Wattonomics 該玩家 AI 資本支出；無智譜玩家。智譜算力以租用（算力服務費）為主，資本開支不是 AI 算力口徑；只有半年一個讀數"],
        "RMB 億", cols="CDEFGHIJKL")
    row("c0_mw", "在役／簽約算力（GW）", [None, None, NA, None, None, None, "無可比實績（公司未揭露、層級無此玩家）",
                                    "公司未揭露自有或租用算力的電力口徑；只揭露推論用國產晶片「10 萬級」張數（SRC_ZP_452，中期公告，Interested-party）與「1 GW 國產 AI 算力資料中心」（SRC_ZP_458，口徑未明、媒體）",
                                    "—", "缺：公司算力 GW 揭露、Wattonomics 智譜玩家；本模型的 GW 由算力服務費 ÷ 每 GW 價格推得（Compute 頁），屬模型值，不當實績"], "GW", cols="CDEFGHIJKL")
    row("c0_fee", "算力成本（1H2026 期末應付算力服務費＋預付算力服務費；只列量級）",
        [f"={S('payable_compute_1h26')}+{S('prepay_compute_1h26')}", None, NA, None, None, None, NAV,
         "SRC_ZP_218（應付算力服務費 14.2742）、SRC_ZP_219（預付算力服務費及其他 12.4374；2026-06-30 資產負債表）",
         "Verified（一手）",
         "期末餘額不是期間支出（1H2026 算力服務費金額公告未揭露）；Wattonomics 無智譜承諾或支出列。只列參考"],
        "RMB 億", cols="CDEFGHIJKL")
    d4 = r[0] - 1
    DR = f"$I${d1}:$I${d4}"
    nm(wb, "COORD_C0", f"Coord!$A${d1}:$K${d4}")

    r[0] += 1
    sec("判定摘要（只計 C 節）")
    n = r[0]
    for k, (a, f) in enumerate([("可判定項目數（對帳中、時差邊界或一致）", f'=COUNTIF({DR},"對帳中")+COUNTIF({DR},"時差邊界*")+COUNTIF({DR},"一致")'),
                                ("對帳中項目數", f'=COUNTIF({DR},"對帳中")'), ("時差邊界（參考）項目數", f'=COUNTIF({DR},"時差邊界*")'),
                                ("參考（同源校準）項目數", f'=COUNTIF({DR},"參考（同源*")'), ("無可比實績項目數", f'=COUNTIF({DR},"無可比實績*")')]):
        put(ws, f"A{n + k}", a, F_CALC)
        put(ws, f"C{n + k}", f, fmt="0")
    nm(wb, "COORD_C0Sum", f"Coord!$C${n}:$C${n + 4}")
    r[0] = n + 5
    put(ws, f"A{r[0]}", "C0 結論", F_BOLD)
    put(ws, f"C{r[0]}", ("智譜的 C0 沒有可判定的項目，對帳中 0：Wattonomics 玩家名單沒有智譜（中國只有合計列），層級沒有智譜的 GW、資本支出或承諾，"
                         "營收、資本支出、算力成本餘額三列都是「無可比實績（層級無此玩家）」；算力 GW 公司亦未揭露。層級也沒有讀入本模型的任何值（無同源列）。"
                         "往上回饋只有證據（C5）：若 Wattonomics 將中國合計拆到公司，可用本頁公司實績（1H2026 營收、算力服務費應付與預付餘額）與公司說法（國產晶片十萬張級）作為錨點。"), F_CALC)
    nm(wb, "COORD_C0Verdict", f"Coord!$C${r[0]}")
    r[0] += 2

    sec("D. 參考（不判定）：模型值 vs 層級中國合計（兩邊都不是公司實績，只看量級）")
    header(ws, r[0], ["項目", "單位", "本模型 2026", "Wattonomics 中國合計 2026 年底", "占比", "", "", "", "", "", "", "說明"])
    r[0] += 1
    n = r[0]
    row("cn", "本模型 2026 供給 GW（Compute CMP_SupplyGW）vs Wattonomics WIF_PlayerGW_YE_cn_2026",
        ["=INDEX(CMP_SupplyGW,1,2)", W["names"].get("WIF_PlayerGW_YE_cn_2026"), f"=C{n}/D{n}"], "GW",
        "本模型 GW 是算力服務費 ÷ 每 GW 價格推得的模型值（非實績）；中國合計為 Rystad 容量 × AI 占比（Wattonomics S1c）。只示意量級，不判定", cols="CDE")
    ws[f"E{n}"].number_format = "0.0%"
    for c, w in zip("ABCDEFGHIJKL", (62, 9, 22, 20, 12, 18, 18, 14, 30, 60, 22, 90)):
        ws.column_dimensions[c].width = w
    ws.freeze_panes = "C5"
    return dict(P=P)
