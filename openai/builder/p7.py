"""OpenAI v0.6.5 三層協調頁 Coord（三層基準協調規則 v1.4.2；工作單 20261114 第二批 13 家 第 2 節；模型商，樣板 Anthropic v0.2.2）。

本頁只列示，不被任何計算頁引用（Checks 除外）；新增本頁不改變任何既有工作表的數值。結構：
  A. 參數（判定門檻、季化倍數、已公布月數；藍字）
  B. 讀取層級（同源，不判定）：本模型讀 Tokenomics v21.7 快照；Wattonomics 讀本模型的值（現金流、資本支出）與本模型現值對照；追溯結論
  C. 層級隱含值推導：Wattonomics T1 配置給 OpenAI 的年底 GW × Tokenomics 每 GW 營收（L1_RevGW_PQ_VR200）
  D. C0（實績）：公司揭露 vs 層級隱含。營收列「參考（同源校準）」（Tokenomics 的 m 由 OpenAI 營收校準）；
     在役算力 GW、資本支出、算力承諾依「有公司揭露且層級有獨立路徑」才判定，否則寫「無可比實績」
  E. 層級之間（參考）：Wattonomics 內部報導值 usr_oai 對 T1 配置；不計入 C0
判定摘要與 C0 結論另給具名範圍（COORD_C0、COORD_C0Sum、COORD_C0Verdict）供 HTML 與 Checks 取值。
層級值來自快照檔（Wattonomics 讀 wattonomics/data 的快照 JSON；Tokenomics 讀 --tk-dir 的 xlsx，值寫成藍字並附版本與雜湊），
公式只引用具名範圍與本頁儲存格，不含常數（判定門檻、季化倍數、月數為本頁藍字參數）。
"""
from __future__ import annotations

import hashlib
import json
from pathlib import Path

import openpyxl

from common import F_BOLD, F_CALC, F_IN, F_NOTE, FILL_SEC, header, put, title

ROOT = Path(__file__).resolve().parent.parent
WATT_DIR = ROOT.parent / "wattonomics" / "data"
TK_L1 = ("L1_RevGW_PQ_VR200_Lo", "L1_RevGW_PQ_VR200", "L1_RevGW_PQ_VR200_Hi", "L1_M")


def read_watt():
    cur = (WATT_DIR / "CURRENT_SNAPSHOT").read_text(encoding="utf-8").strip()
    d = json.loads((WATT_DIR / cur).read_text(encoding="utf-8"))
    rows = {r["id"]: r for r in d["rows"] if "id" in r}
    return dict(file=d["source"]["file"], version=d["source"]["version"], sha=d["source"]["sha256"], rows=rows, names=d["names"], snap=cur)


def read_tk_l1(tk_dir):
    """Tokenomics checkout 的 CURRENT xlsx 取 L1 每 GW 營收（只列示於 Coord；TK_Link 不變）。"""
    tk_dir = Path(tk_dir)
    cur = (tk_dir / "model" / "CURRENT").read_text(encoding="utf-8").strip()
    path = tk_dir / "model" / cur
    wb = openpyxl.load_workbook(path, data_only=True)
    out, cells = {}, {}
    for n in TK_L1:
        sh, c = next(iter(wb.defined_names[n].destinations))
        out[n] = wb[sh][c.replace("$", "")].value
        cells[n] = f"{sh}!{c}"
    return dict(file=cur, sha=hashlib.sha256(path.read_bytes()).hexdigest(), vals=out, cells=cells)


def build_p7(wb, R, S, I, nm, snap, tk_dir):
    W = read_watt()
    TKX = read_tk_l1(tk_dir)
    RW, N = W["rows"], W["names"]
    ws = wb.create_sheet("Coord")
    title(ws, "Coord — 三層協調（v0.6.5；規則 v1.4.2）：讀取層級、C0（實績）公司揭露 vs 層級隱含值",
          "C0 比較基準＝公司揭露（OpenAI 未上市，無財報；只有公司告知投資人、公司 CFO 部落格與合作方公告）；層級側＝Tokenomics 每 GW 營收 × Wattonomics T1 配置給 OpenAI 的 GW。"
          "營收列屬同源：Tokenomics 的 m 由 OpenAI 營收 ÷（付費 token × 牌價）校準，標「參考（同源校準）」；其餘列有公司揭露且層級有獨立路徑才判定，否則寫「無可比實績」。",
          "本頁只列示，不被任何計算頁引用（Checks 除外）。單位：$B、GW。藍字＝快照值或參數，黑字＝公式。")
    P = {}
    r = [5]

    def sec(t):
        put(ws, f"A{r[0]}", t, F_BOLD, fill=FILL_SEC)
        r[0] += 1

    def row(key, label, vals, unit="", note="", fmt="#,##0.000", cols="CDEFGHIJK"):
        n = r[0]
        put(ws, f"A{n}", label, F_CALC)
        put(ws, f"B{n}", unit, F_CALC)
        for c, v in zip(cols, vals):
            if v is None:
                continue
            is_f = isinstance(v, str) and v.startswith("=")
            is_txt = isinstance(v, str) and not is_f
            put(ws, f"{c}{n}", v, F_CALC if (is_f or is_txt) else F_IN, fmt=None if is_txt else fmt)
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
    row("qx", "季化倍數（最近一季 × 本值＝年化速率）", [4], "倍", "第二讀數＝2026 Q2 × 本值 vs 層級 CY2026", fmt="0")
    row("m6", "2026 已公布月數（Q1、Q2）", [6], "月", "OpenAI 告知投資人的 2026 Q1、Q2 營收（SRC_OAI_004、005）涵蓋 2026-01～06；Q3 尚無公司揭露（9 月底 run-rate 為媒體轉述，不是季營收）", fmt="0")
    row("m12", "每年月數", [12], "月", "年化與時差換算", fmt="0")
    row("k2", "平均除數", [2], "—", "年底值平均＝（期初＋期末）÷ 本值", fmt="0")
    TH, QX, M6, M12, K2 = (f"$C${P[k]}" for k in ("th", "qx", "m6", "m12", "k2"))

    # ── B. 讀取層級 ──
    r[0] += 1
    sec("B. 讀取層級（同源，不判定）：本模型實際讀的層級、版本與名稱；自層級讀入的值只列在此節，不進 C0")
    header(ws, r[0], ["項目", "單位", "本模型值", "層級現值（快照）", "差", "", "", "", "", "", "", "說明"])
    r[0] += 1
    row("tkv", "Tokenomics 版本（TK_Link 快照）", ["=TK_Version"], "",
        "本模型讀 Tokenomics 快照 v21.7（單一快照；TK_Link）；Tokenomics master 現行為 v22.0，本輪不換快照（工作單：模型仍讀舊快照者寫明版本、不在本輪換）。TK 值只作算力物理與對照，不回饋 TK",
        fmt="@", cols="C")
    row("tkc", "TK_Link 已取值名稱數", ['=COUNTIF(TK_Link!$F$10:$F$200,"OK")'], "個", "計算頁引用 TK_ 具名範圍；本頁 C 節另讀 L1 每 GW 營收三值（取自同一份 v21.7 xlsx，不進 TK_Link）", fmt="0", cols="C")
    row("tkl1", "Tokenomics 校準：L1_M（收入面 m；由 OpenAI 營收 ÷（付費 token × 牌價）校準）", [TKX["vals"]["L1_M"]], "倍",
        f"取自 {TKX['file']}（SHA-256 {TKX['sha'][:16]}…）{TKX['cells']['L1_M']}；m 以 OpenAI 營收校準＝同源，故 D 節營收列標「參考（同源校準）」", cols="C")
    row("wtv", "Wattonomics 快照（本頁 C、E 節取值）", [f"{W['version']}｜{W['file']}"], "", f"快照檔 wattonomics/data/{W['snap']}；SHA-256 {W['sha']}", cols="C")
    row("trace", "Wattonomics oai、usr_oai 路徑追溯（C0 前置檢查；tools/coord/trace_wattonomics.py＋人工確認）", ["已逐一追溯：層級 GW 不讀本模型的 GW，只讀本模型的現金流與資本支出回寫"], "",
        "追溯：(1) T1 配置 al_orcl_oai（2026 年底 0.824、2027 年底 0.991；來源＝Oracle 公告 Abilene，公司揭露）與 al_oai_self（OpenAI／Stargate 自建，比例 1）為 Wattonomics In 頁輸入列，只進 Matrix（營運者×使用者）；"
        "(2) 玩家路徑（Players、CapexGW）：OpenAI 自建年底 GW＝0（2025、2026；2027 起由 Ports Pike、Milam 等園區輸入列延伸，皆為公司公告的園區），自建 AI 資本支出 2026＝0；"
        "(3) CapexAIModel（In 頁 gcapai_oai，Wattonomics 稱「個股模型寫回」）讀本模型 Cost 自有資本支出，只進 Gap 頁對照欄；Interface 的 CapexAI 取 Gap!G476＝IF(G475<>\"\",G471,G460)，G475 是空字串常數，故取 G460（GW 法）＝0，不讀本模型"
        "（巢狀 IF 為工具已知限制，已人工確認）；(4) 營業現金流 gocf_oai、期初現金 gcash_oai 讀本模型（Gap 頁；影響資金缺口，不影響 GW）；"
        "(5) usr_oai（Y4 使用者報導 GW：2025 年底 1.9、2026 年底 5、2027 年底 11）只在 Matrix 對照列（Matrix!D21、D42、D63）被讀取，不進任何計算；它的 2025 值取自公司 CFO 部落格（SRC_OAI_090 同一來源），2026、2027 為第三方報導與 Assumed 中點；"
        "(6) 未發現任何 w1_*、obs_*、act26_* 或 Players 覆寫讀本模型的 GW。結論：T1 配置 GW 與自建 GW 不同源；營收的同源來自 Tokenomics 的 m，不來自 Wattonomics。", cols="C")
    row("wtm", "Wattonomics 讀本模型的版本", ["OpenAI 收支模型 v0.6.4（commit 4c027d8；In 頁 gcapai_oai、gocf_oai、gcash_oai，非即時連結）"], "",
        "下列自有資本支出、營業現金流、2025 年底現金：本模型現值 vs Wattonomics 匯入值（同源，不判定）；差為 0＝層級匯入未落後。本輪（v0.6.5）不改本模型數值，故差維持 0", cols="C")
    for k, y in enumerate(range(2026, 2031), start=2):
        n = r[0]
        row(f"cap{y}", f"自有資本支出 {y}：Cost 現值 vs Wattonomics WIF_Gap_Player_oai_{y}_CapexAIModel",
            [f"=INDEX(COST_OwnedCapex,1,{k})", N.get(f"WIF_Gap_Player_oai_{y}_CapexAIModel"), f"=C{n}-D{n}"], "$B",
            "同源（本模型寫入 Wattonomics 對照欄；Wattonomics 的 CapexAI 不讀它）；差為 0＝未落後" if y == 2026 else "", fmt="#,##0.00", cols="CDE")
    for k, y in enumerate(range(2026, 2031), start=2):
        n = r[0]
        row(f"ocf{y}", f"營業現金流 {y}：本模型（Funding 自由現金流 F06＋Cost 自有資本支出）vs Wattonomics WIF_Gap_Player_oai_{y}_OCF",
            [f"=INDEX(FND_FCF,1,{k})+INDEX(COST_OwnedCapex,1,{k})", N.get(f"WIF_Gap_Player_oai_{y}_OCF"), f"=C{n}-D{n}"], "$B",
            "Wattonomics 的營業現金流由本模型匯入（SRC_GL_319：Funding!13＋Cost!26）；同源，不判定、不計入對帳中" if y == 2026 else "", fmt="#,##0.00", cols="CDE")
    n = r[0]
    row("cash25", "2025 年底現金：Funding 現值 vs Wattonomics gcash_oai", ["=INDEX(FND_CashEnd,1,1)", RW["gcash_oai"]["values"][1], f"=C{n}-D{n}"], "$B",
        "同源；Wattonomics 期初現金 40 取自本模型 Funding!D30", fmt="#,##0.00", cols="CDE")

    # ── C. 層級隱含值推導 ──
    r[0] += 1
    sec("C. 層級隱含值推導：Wattonomics 配置給 OpenAI 的年底 GW（T1 營運者→使用者）× Tokenomics 每 GW 營收")
    header(ws, r[0], ["項目", "單位", "2025 年底", "2026 年底", "", "", "", "", "", "", "", "說明"])
    r[0] += 1

    def t1(rid):
        v = RW[rid]["values"]
        assert v[1] == "oai" and v[2] in ("GW", "比"), (rid, v[:3])      # 依鍵確認：使用者＝oai；形式 GW 或比（索引 3、4、5＝2025、2026、2027 年底）
        return v[2], v[3], v[4]
    z = lambda x: 0 if x is None else x  # noqa: E731
    f1, a1, b1 = t1("al_orcl_oai")
    f2, a2, b2 = t1("al_oai_self")
    assert f1 == "GW" and f2 == "比"
    row("al_orcl", "Wattonomics T1：Oracle（OCI；含承租 Stargate 站點）→ OpenAI", [z(a1), z(b1)], "GW",
        f"al_orcl_oai（{RW['al_orcl_oai']['tag']}；{RW['al_orcl_oai']['src']}）：{RW['al_orcl_oai']['note'][:120]}；空白＝層級未計入，以 0 計", cols="CD")
    row("al_self_r", "Wattonomics T1：OpenAI／Stargate 自建（SoftBank、SB Energy）→ OpenAI（占營運者年底 GW 的比例）", [z(a2), z(b2)], "比",
        f"al_oai_self（{RW['al_oai_self']['tag']}）：使用者為 OpenAI 的比例 1", fmt="0.00", cols="CD")
    row("oai_op", "Wattonomics 玩家 oai（OpenAI／Stargate 自建）年底上線 IT GW", [N["WIF_PlayerGW_YE_oai_2025"], N["WIF_PlayerGW_YE_oai_2026"]], "GW",
        "WIF_PlayerGW_YE_oai_2025、_2026（快照值）：自建站點 2027 年起才有上線量", cols="CD")
    n0, n1, n2 = P["al_orcl"], P["al_self_r"], P["oai_op"]
    row("gw_ye", "配置給 OpenAI 的 GW 合計（年底，IT）＝Oracle 配置＋自建 GW × 比例", [f"=C{n0}+C{n2}*C{n1}", f"=D{n0}+D{n2}*D{n1}"], "GW",
        "對應 Wattonomics Matrix 的 OpenAI 欄和（2025 年底 0、2026 年底 0.824）；T1 只計有證據的配置（空白＝未分配），因此是下界；其餘 OpenAI 使用的容量（Microsoft、CoreWeave、AWS 等）在各營運者的「未分配」中", cols="CD")
    row("gw_avg", "2026 年平均 GW（＝2025 年底與 2026 年底的簡單平均）", [None, f"=(C{P['gw_ye']}+D{P['gw_ye']})/{K2}"], "GW",
        "[Derived] 線性近似；年內上線時點未知", cols="CD")
    GWAVG = f"$D${P['gw_avg']}"
    v = TKX["vals"]
    row("pq", "Tokenomics L1_RevGW_PQ_VR200：每 GW 年營收（第一年；低／基準／高）", [v["L1_RevGW_PQ_VR200_Lo"], v["L1_RevGW_PQ_VR200"], v["L1_RevGW_PQ_VR200_Hi"]], "$B/GW/年",
        f"取自 {TKX['file']}（本模型 TK_Link 的同一快照 v21.7）{TKX['cells']['L1_RevGW_PQ_VR200']}；VR200 口徑；m 以 OpenAI 營收與牌價校準＝同源（B 節）。"
        "Tokenomics 自註：Q 是模型產能推算，若誤差全在產能側，實際約為本列 ÷ 2.0–2.1（OpenAI 2026 口徑）", cols="CDE")
    row("cy", "層級隱含 CY2026 營收（每 GW 營收 × 2026 年平均 GW；低／基準／高）", [f"=C{P['pq']}*{GWAVG}", f"=D{P['pq']}*{GWAVG}", f"=E{P['pq']}*{GWAVG}"], "$B",
        "層級隱含值；GW 取 T1 配置的下界，故只是量級參考", cols="CDE")

    # ── D. C0（實績） ──
    r[0] += 1
    sec("D. C0（實績）：公司揭露 vs 層級隱含值（規則 v1.4.2 C0 第 2 點；Andy 11-14 17:40 固定規格）")
    header(ws, r[0], ["項目", "單位", "公司揭露（主）", "層級隱含（主）", "差距（主）", "公司揭露（第二）", "層級隱含（第二）", "差距（第二）", "判定", "出處與口徑", "等級", "說明", "機械判定（若不計同源／口徑；參考）"])
    r[0] += 1
    d1 = r[0]
    mech = lambda n: (f'=IF(AND(ISNUMBER(E{n}),ISNUMBER(H{n})),IF(AND(ABS(E{n})>{TH},ABS(H{n})>{TH},SIGN(E{n})=SIGN(H{n})),"對帳中",'  # noqa: E731
                      f'IF(OR(ABS(E{n})>{TH},ABS(H{n})>{TH}),"時差邊界（參考）","一致")),"—")')
    n = r[0]
    row("c0_rev", "營收（OpenAI 淨額口徑，公司告知投資人）",
        [f"={S('營收：2026 Q1')}+{S('營收：2026 Q2')}", f"=$D${P['cy']}*{M6}/{M12}", f"=C{n}/D{n}-1", f"={S('營收：2026 Q2')}", f"=$D${P['cy']}", f"=F{n}*{QX}/G{n}-1",
         "參考（同源校準）",
         "SRC_OAI_004（2026 Q1 5.7；The Information 2026-06-16，股東文件）、SRC_OAI_005（2026 Q2 6.7；WSJ 2026-08-18，告知投資人）",
         "Interested-party（公司告知投資人；二手轉載，來源等級 2；非公開財報）",
         "同源：Tokenomics 的 m（L1_M）由 OpenAI 營收 ÷（付費 token × 牌價）校準，每 GW 營收 × GW 與 OpenAI 營收不是獨立比較——m 與牌價約掉後，本列實際只檢驗「OpenAI 營收 vs 校準值 × Wattonomics T1 配置 GW」，"
         "差距主要反映 T1 配置 GW 的覆蓋度，不是營收證據；判定標「參考（同源校準）」、不計入對帳中，結論＝無可比實績（層級值同源）。"
         "公司只揭露 Q1、Q2（無 Q3），已公布期間 1–6 月；時差偏誤：主讀數期間中點在 4 月初、早於全年中點，成長期累計占全年比例低於 6/12，偏低；第二讀數 Q2×4 是 5 月中的年化速率，仍早於全年中點，成長期亦偏低——兩者偏誤同向，不得讀成「夾住真值」。"
         "9 月底 run-rate（SRC_OAI_144，FT，二手、原文未讀）與 2025 年底 ARR 20（SRC_OAI_091）屬 run-rate，不算季營收，不入本列",
         mech(n)],
        "$B", fmt="#,##0.000", cols="CDEFGHIJKLM")
    for c in "EH":
        ws[f"{c}{n}"].number_format = "0.0%"
    n = r[0]
    row("c0_gw", "在役算力（2025 年底；單一時點，兩側同日，無時差）：公司 CFO 揭露 vs T1 配置 GW",
        [f"={S('算力規模：2025 年底')}", f"=$C${P['gw_ye']}", f'=IF(D{n}=0,"—",C{n}/D{n}-1)', None, None, None,
         "無可比實績（層級 T1 為下界）",
         "SRC_OAI_090（算力規模 2025 年底約 1.9 GW；CFO Friar 部落格 2026-01-18；口徑未明，Epoch 假設為 IT）；層級＝本頁 C 節 T1 配置合計（Wattonomics Matrix 的 OpenAI 欄和）",
         "Interested-party（公司 CFO 部落格；來源等級 3，查核狀態「待查核-報導」）；層級＝快照值",
         "層級 T1 配置給 OpenAI 的 2025 年底 GW 為 0：Wattonomics 只計有逐站證據的配置（Oracle Abilene 2025 年底上線量未揭露→未分配），空白不是層級對 OpenAI 算力為 0 的主張，而是下界；"
         "公司揭露 1.9 GW 大於下界 0，兩者相容，沒有可比的層級讀數（缺：Wattonomics 對 OpenAI 全部使用算力〔含 Microsoft、CoreWeave、AWS〕的配置）。單一時點只有一個讀數，不做兩讀數。"
         "層級值為 0，百分比差距無定義（M 欄顯示「—」），只反映配置覆蓋度；是否開對帳單（證據回饋：Wattonomics 補 2025 年底配置）留給 Wattonomics 端決定（見 E 節與結論）",
         f'=IF(ISNUMBER(E{n}),IF(ABS(E{n})>{TH},"超過門檻（單一讀數）","一致"),"—")'],
        "GW", fmt="#,##0.000", cols="CDEFGHIJKLM")
    ws[f"E{n}"].number_format = "0.0%"
    n = r[0]
    row("c0_capex", "資本支出（公司揭露已發生值）vs Wattonomics 自建 AI 資本支出（2026）",
        [None, N["WIF_Gap_Player_oai_2026_CapexAI"], None, None, None, None,
         "無可比實績（公司無已發生資本支出揭露）",
         "OpenAI 未上市，無現金流量表；公司只揭露計畫（SRC_OAI_139 Brockman 證詞：2026 算力支出約 50，是計畫不是實績；SRC_OAI_099 Camellia 20 為園區計畫例）；層級＝Wattonomics WIF_Gap_Player_oai_2026_CapexAI（快照值）",
         "—",
         "缺：公司已發生的資本支出揭露（OpenAI 租用為主，算力支出走雲端合約，不是資本支出）。層級值 0 是 GW 法（自建上線 GW 為 0）的結果；Wattonomics 的 CapexAIModel 20 取自本模型 Cost 自有資本支出（B 節，同源），不能當層級獨立值。不判定",
         None],
        "$B", cols="CDEFGHIJKLM")
    n = r[0]
    row("c0_commit", "算力承諾（Oracle 合約容量 GW）：公司揭露 vs Wattonomics ctr_orcl_oai",
        [f"={S('合約容量：Oracle')}", RW["ctr_orcl_oai"]["values"][1], f"=C{n}-D{n}", None, None, None,
         "無可比實績（層級讀同一份公司公告）",
         "SRC_OAI_061（Oracle 合約容量約 4.5 GW；WSJ 2025-09 轉述 Oracle 公告，電力口徑未明）；層級＝Wattonomics ctr_orcl_oai（SRC_GL_200，A#71 未抽查；同樣出自 Oracle 公告）",
         "Interested-party（Oracle 公告；來源等級 2）；層級＝快照值",
         "兩邊都讀同一份 Oracle 公告，差距恆為 0，沒有檢驗力（規則 v1.4.2：同源值不得當比較基準）；Wattonomics 另無其他承諾（AWS 2 GW Trainium、Cerebras 0.75 GW、Azure 250B 美元）的獨立 GW 路徑。"
         "承諾金額（美元）與層級 GW 口徑不同（缺層級每 GW 合約價）。不判定",
         None],
        "GW", fmt="#,##0.000", cols="CDEFGHIJKLM")
    d4 = r[0] - 1
    DR = f"$I${d1}:$I${d4}"
    nm(wb, "COORD_C0", f"Coord!$A${d1}:$K${d4}")

    # ── 判定摘要 ──
    r[0] += 1
    sec("判定摘要（只計 D 節）")
    n = r[0]
    put(ws, f"A{n}", "可判定項目數（判定為對帳中、時差邊界或一致）", F_CALC)
    put(ws, f"C{n}", f'=COUNTIF({DR},"對帳中")+COUNTIF({DR},"時差邊界*")+COUNTIF({DR},"一致")', fmt="0")
    put(ws, f"A{n + 1}", "對帳中項目數", F_CALC)
    put(ws, f"C{n + 1}", f'=COUNTIF({DR},"對帳中")', fmt="0")
    put(ws, f"A{n + 2}", "時差邊界（參考）項目數", F_CALC)
    put(ws, f"C{n + 2}", f'=COUNTIF({DR},"時差邊界*")', fmt="0")
    put(ws, f"A{n + 3}", "參考（同源校準）項目數", F_CALC)
    put(ws, f"C{n + 3}", f'=COUNTIF({DR},"參考（同源*")', fmt="0")
    put(ws, f"A{n + 4}", "無可比實績項目數", F_CALC)
    put(ws, f"C{n + 4}", f'=COUNTIF({DR},"無可比實績*")', fmt="0")
    nm(wb, "COORD_C0Sum", f"Coord!$C${n}:$C${n + 4}")
    P["sum"] = n
    r[0] = n + 5
    put(ws, f"A{r[0]}", "C0 結論", F_BOLD)
    put(ws, f"C{r[0]}", ("OpenAI 的 C0 沒有可判定的項目，對帳中 0：營收列與層級每 GW 營收同源（Tokenomics 的 m 以 OpenAI 營收校準），標「參考（同源校準）」，結論＝無可比實績（層級值同源）；"
                         "在役算力（2025 年底公司揭露約 1.9 GW）對層級 T1 配置 0 GW：T1 只計有逐站證據的配置、空白為未分配，層級值是下界，與公司揭露相容，無可比實績；"
                         "資本支出（公司無已發生值）與算力承諾（層級讀同一份 Oracle 公告）亦無可比實績。往上回饋只有證據（C5）：Wattonomics 的 OpenAI 配置覆蓋度偏低（2025 年底 0 對公司揭露 1.9），"
                         "建議 Wattonomics 端以公司揭露補 2025 年底配置證據；待 OpenAI 有 10-Q 等法定揭露（或 S-1 公開）再判定營收與資本支出。"), F_CALC)
    nm(wb, "COORD_C0Verdict", f"Coord!$C${r[0]}")
    P["verdict"] = r[0]
    r[0] += 1

    # ── E. 層級之間（參考） ──
    r[0] += 1
    sec("E. 層級之間（參考；規則 v1.4.2 C0：兩層之間的差距只列參考，不計入個股 C0）")
    header(ws, r[0], ["項目", "單位", "Wattonomics 報導值 usr_oai", "T1 配置合計（本頁 C 節）", "比值", "", "", "", "", "", "", "說明"])
    r[0] += 1
    u = RW["usr_oai"]["values"]        # Y4：2025＝索引 2、2026＝索引 3、2027＝索引 4
    n = r[0]
    row("usr25", "OpenAI 使用算力年底 GW（2025）：報導值 vs T1 配置", [u[2], f"=C{P['gw_ye']}", f'=IF(D{n}=0,"—",C{n}/D{n})'], "GW",
        f"Y4 usr_oai（{RW['usr_oai']['tag']}）：2025 取公司 CFO 部落格（即 D 節公司揭露）。T1 為有逐站證據的配置（下界）；兩者都在 Wattonomics 內，差距屬層級內部，不影響 C0", cols="CDE")
    n = r[0]
    row("usr26", "OpenAI 使用算力年底 GW（2026）：報導值 vs T1 配置", [u[3], f"=D{P['gw_ye']}", f'=IF(D{n}=0,"—",C{n}/D{n})'], "GW",
        "2026 約 5 為第三方報導（B#6 5–6 GW），非公司揭露、非實績；T1 只含 Oracle Abilene 0.824", cols="CDE")
    ws[f"E{n}"].number_format = "0.00"
    ws[f"E{n - 1}"].number_format = "0.00"
    for c, w in zip("ABCDEFGHIJKLM", (62, 9, 22, 22, 12, 20, 20, 18, 30, 60, 30, 90, 26)):
        ws.column_dimensions[c].width = w
    ws.freeze_panes = "C5"
    return dict(P=P)


def p7_checks(P7):
    P = P7["P"]
    n = P["c0_rev"]
    return [
        ("v0.6.5 Coord：C0 營收主、第二讀數差距皆為數值", f"=IF(AND(ISNUMBER(Coord!$E${n}),ISNUMBER(Coord!$H${n})),1,0)", 1, "eq", "三層協調（v0.6.5）；只列示，不回饋計算"),
        ("v0.6.5 Coord：C0（實績）對帳中項目數（資訊）", "=INDEX(COORD_C0Sum,2)", None, "info", "對帳中＝兩讀數都超過 ±15% 且同號；營收列為同源，不計入；開對帳單、不改上游"),
    ]
