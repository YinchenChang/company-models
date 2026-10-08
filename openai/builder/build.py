#!/usr/bin/env python3
"""產生 OpenAI 收支模型 Excel（v0.6-P1）。

用法：python3 builder/build.py --tk-dir <Tokenomics repo 路徑> --out model/20261004_OpenAI_v0.6.xlsx [--base <前一版>] [--date 2026-10-04]
  --base：前一版 xlsx；有給時依 Excel 優先規則保留 Excel 內已改過的藍字輸入（builder/preserve.py）。
輸出經 LibreOffice 重算後才寫到 --out（內含快取值；engine 仍會強制重算）。同目錄寫 restore_log.txt、leafmap.csv、build_summary.json。
"""
import argparse
import csv
import json
import sys
import tempfile
from pathlib import Path

import openpyxl
from openpyxl.workbook.defined_name import DefinedName

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))
import preserve  # noqa: E402
import source_rules  # noqa: E402
from common import F_BOLD, F_CALC, F_IN, F_NOTE, WRAP, header, lo_recalc, put, title  # noqa: E402
from tk_link import PENDING_NOTE, read_snapshot  # noqa: E402
from v05_map import REGISTRY_PATH, SEP, build_map, coverage_report  # noqa: E402
import p2  # noqa: E402
import p3  # noqa: E402
import p4  # noqa: E402
import p5  # noqa: E402

VERSION = "v0.6"
KIND_LABEL = {"SRC": "SRC_OAI", "INP": "Inputs", "FORMULA": "公式（後續工作包）", "DUP": "重複併入", "SKIP": "不遷入", "TK": "不遷入；改取 TK_Link"}
SRC_COLS = ["SRC_ID", "指標", "數值", "低", "高", "單位", "口徑", "適用對象", "日期", "出處（v0.5 原文）", "來源等級", "立場", "立場說明",
            "一手／二手", "狀態", "取代者", "v0.5 標記", "查核狀態（v0.5 chk）", "v0.5 路徑", "模型使用位置", "備註",
            "缺值（公式）", "缺出處（公式）", "區間順序異常（公式）", "四欄標註"]
INP_COLS = ["INP_ID", "參數", "索引", "單位", "值", "低", "高", "標記", "v0.5 原決策", "狀態", "v0.5 路徑", "備註", "區間順序異常（公式）"]
TK_FIRST, TK_VCOL0 = 10, 10       # TK_Link 表格從第 10 列；值從 J 欄（第 10 欄）起，最多 15 欄


def nm(wb, name, ref):
    wb.defined_names[name] = DefinedName(name, attr_text=ref)


def blank(v):
    return None if v == "" else v




def sheet_src(wb, R, final):
    ws = wb.create_sheet("SRC_OAI")
    title(ws, "SRC_OAI — 第 0 層 Source：OpenAI 公司財務原始數據（營收、合約、現金、融資輪、算力規模）",
          "只存已取得的原始訊息（非計算、非假設；V5）。藍字＝原始值，由本活頁簿擁有；格式比照 Tokenomics SRC_Demand。AI 技術與算力原始數據不在此登錄（取自 TK_Link）。",
          "來源等級（1 一手已讀原文／2 二手已讀或轉載／3 待查核）、立場、立場說明、一手／二手：依 builder/source_rules.py 的機械規則由 CC 初評（欄 Y 標『CC 初評』；V10）；規則無法判定者填『未評』。本包不做新查核。　圖例：『二手（含一手公告轉載）』＝媒體轉載公司公開公告；『二手（含一手轉載）』＝媒體轉載外流財報或內部文件（F2）；兩者不合併。")
    header(ws, 4, SRC_COLS)
    for i, r in enumerate(R.src_rows):
        n = 5 + i
        cl = source_rules.classify(r, R.src_rows[i - 1]["source"] if i else "")
        r["_cl"] = cl
        vals = [r["id"], r["metric"], blank(r["value"]), blank(r["lo"]), blank(r["hi"]), r["unit"], r["scope"], r["subject"], r["date"], r["source"],
                cl["grade"], cl["stance"], cl["why"], cl["hand"], r["status"], "—", r["tag"], r["chk"], r["v05"].replace(SEP, "."), r["use"], r["note"]]
        for c, v in enumerate(vals, start=1):
            cell = ws.cell(row=n, column=c, value=v)
            cell.font = F_IN if c in (3, 4, 5) and isinstance(v, (int, float)) else F_CALC
        for col in "CDE":
            k = ("SRC_OAI", r["id"], col)
            ws[f"{col}{n}"].value = final[k]
        put(ws, f"V{n}", f'=IF(AND($O{n}="Active",$C{n}="",$D{n}="",$E{n}="",$F{n}<>"文字"),1,0)')
        put(ws, f"W{n}", f'=IF(AND($O{n}="Active",OR($J{n}="",$J{n}="—")),1,0)')
        put(ws, f"X{n}", f'=IF(AND(ISNUMBER($C{n}),ISNUMBER($D{n}),ISNUMBER($E{n})),IF(AND($D{n}<=$C{n},$C{n}<=$E{n}),0,1),0)')
        put(ws, f"Y{n}", "CC 初評", F_NOTE)
        if r["value"] != "":
            nm(wb, r["id"], f"SRC_OAI!$C${n}")
    for col, w in zip("ABCDEFGHIJ", (13, 34, 10, 8, 8, 12, 40, 12, 11, 50)):
        ws.column_dimensions[col].width = w
    ws.freeze_panes = "C5"
    return ws


def sheet_inputs(wb, R, final):
    ws = wb.create_sheet("Inputs")
    title(ws, "Inputs — 假設輸入（藍字）；v0.6-P1：只建表頭與 v0.5 Assumed／Analogy／Decision 項目的遷入",
          "長表格式：一列一個（參數 × 索引）。標記沿用 v0.5；區間倍數（低／高情境倍數）另列。Analogy／Assumed 無區間者保留空白並於備註說明（不自行補值）。",
          "屬 Tokenomics 領域的 v0.5 項目（rack、classes、serving、hwCostPerM、tokenomicsRef、genFactor.hopper／nextGen）不在此頁，見 Map_v05：『不遷入；改取 TK_Link』。")
    header(ws, 4, INP_COLS)
    for i, r in enumerate(R.inp_rows):
        n = 5 + i
        vals = [r["id"], r["name"], r["index"], r["unit"], blank(r["value"]), blank(r["lo"]), blank(r["hi"]), r["tag"], r["decision"],
                r["status"], r["v05"].replace(SEP, "."), r["note"]]
        for c, v in enumerate(vals, start=1):
            cell = ws.cell(row=n, column=c, value=v)
            cell.font = F_IN if c in (5, 6, 7) and isinstance(v, (int, float)) else F_CALC
        for col in "EFG":
            ws[f"{col}{n}"].value = final[("Inputs", r["id"], col)]
        put(ws, f"M{n}", f'=IF(AND(ISNUMBER($E{n}),ISNUMBER($F{n}),ISNUMBER($G{n})),IF(AND($F{n}<=$E{n},$E{n}<=$G{n}),0,1),0)')
        if r["value"] != "":
            nm(wb, r["id"], f"Inputs!$E${n}")
    for col, w in zip("ABCDEFGHIJKL", (10, 40, 22, 14, 10, 8, 8, 18, 10, 30, 40, 60)):
        ws.column_dimensions[col].width = w
    ws.freeze_panes = "C5"
    return ws


def sheet_tk(wb, snap, date):
    ws = wb.create_sheet("TK_Link")
    title(ws, "TK_Link — Tokenomics 快照（第 0 層連結；不用 Excel 外部連結）",
          "每列一個 Tokenomics 名稱（SRC_／IF_／L1_）。值（藍字）由 builder 從 Tokenomics master 的 model/CURRENT 讀出寫入；本模型公式只引用 TK_ 具名範圍。更新快照＝本模型的修補版。",
          "Interface 向量名稱為 15 欄＝5 世代（Hopper、GB200、GB300、VR200、Rubin Ultra）× 3 成本情境（低、基準、高）；基準欄為每世代第 2 欄。")
    meta = [("Tokenomics 檔案", snap["file"], "TK_File"), ("Tokenomics 版本", snap["version"], "TK_Version"),
            ("Tokenomics 提交 SHA（master）", snap["sha"], "TK_Commit"), ("讀取日期", date, "TK_ReadDate")]
    for i, (lab, v, name) in enumerate(meta):
        put(ws, f"A{3 + i}", lab, F_BOLD)
        put(ws, f"B{3 + i}", v, F_CALC)
        nm(wb, name, f"TK_Link!$B${3 + i}")
    # 表頭（世代／成本情境）
    put(ws, "I7", "世代", F_BOLD)
    put(ws, "I8", "成本情境", F_BOLD)
    for k in range(15):
        put(ws, f"{openpyxl.utils.get_column_letter(TK_VCOL0 + k)}7", snap["hdr_gen"][k], F_CALC)
        put(ws, f"{openpyxl.utils.get_column_letter(TK_VCOL0 + k)}8", snap["hdr_cost"][k], F_CALC)
    nm(wb, "TK_HdrGen", f"TK_Link!$J$7:$X$7")
    nm(wb, "TK_HdrCost", f"TK_Link!$J$8:$X$8")
    header(ws, 9, ["Tokenomics 名稱", "本模型名稱", "說明", "單位", "欄數", "狀態", "Tokenomics 版本", "提交 SHA", "讀取日期"] + [f"值{k + 1}" for k in range(15)])
    r = TK_FIRST
    L = openpyxl.utils.get_column_letter
    for row in snap["rows"]:
        n = len(row["values"])
        vals = [row["name"], "TK_" + row["name"], row["label"], row["unit"], n, "OK", snap["version"], snap["sha"][:7], date]
        for c, v in enumerate(vals, start=1):
            put(ws, f"{L(c)}{r}", v, F_CALC)
        for k, v in enumerate(row["values"]):
            put(ws, f"{L(TK_VCOL0 + k)}{r}", v, F_IN)
        nm(wb, "TK_" + row["name"], f"TK_Link!$J${r}" + (f":${L(TK_VCOL0 + n - 1)}${r}" if n > 1 else ""))
        r += 1
    for row in snap["pending"]:
        vals = [row["name"], "（Tokenomics 現行版無此名稱；無具名範圍）", "Tokenomics 缺口：見報告", "", 0, row["status"], snap["version"], snap["sha"][:7], date]
        for c, v in enumerate(vals, start=1):
            put(ws, f"{L(c)}{r}", v, F_CALC)
        r += 1
    ws.column_dimensions["A"].width = 26
    ws.column_dimensions["B"].width = 30
    ws.column_dimensions["C"].width = 50
    ws.freeze_panes = "D10"
    return r - 1


def sheet_map(wb, R):
    ws = wb.create_sheet("Map_v05")
    title(ws, "Map_v05 — v0.5 JSON 每個葉節點的去處（表 1 的明細）",
          "類別：SRC_OAI／Inputs／公式（後續工作包）／重複併入／不遷入／不遷入；改取 TK_Link。『值』欄為 v0.5 原值（含不遷入者）。")
    header(ws, 4, ["v0.5 路徑", "v0.5 值", "去處類別", "去處", "理由／備註"])
    for i, (p, v) in enumerate(R.leaves.items()):
        k, ref, why = R.dest[p]
        for c, x in enumerate([p.replace(SEP, "."), json.dumps(v, ensure_ascii=False), KIND_LABEL[k], ref, why], start=1):
            ws.cell(row=5 + i, column=c, value=x).font = F_CALC
    ws.column_dimensions["A"].width = 60
    ws.column_dimensions["B"].width = 30
    ws.column_dimensions["D"].width = 40
    ws.freeze_panes = "A5"


def sid(R, metric):
    for r in R.src_rows:
        if r["metric"] == metric:
            return r["id"]
    raise KeyError(metric)


def inp_id(R, name_prefix, index=None):
    for r in R.inp_rows:
        if r["name"].startswith(name_prefix) and (index is None or r["index"].startswith(index)):
            return r["id"]
    raise KeyError((name_prefix, index))


def src_row(R, metric):
    return 5 + next(i for i, r in enumerate(R.src_rows) if r["metric"] == metric)


def inp_row(R, iid):
    return 5 + [r["id"] for r in R.inp_rows].index(iid)


def sheet_derived(wb, R):
    """Derived_V9：V9、V12、V15、E8 的公式（E6：公式內不含常數，換算係數與比例全在 Inputs；v0.5 原值並列對照）。"""
    ws = wb.create_sheet("Derived_V9")
    title(ws, "Derived_V9 — 跨年混合項目與 Derived 項目的公式（V9、V12、V15、E8）；v0.5 原值並列",
          "E6：公式不內含常數（算式中的 1－比例 為恆等式），換算係數與比例都在 Inputs（標記＋區間）。『v0.5 原值』欄為 builder 由 v0.5 JSON 寫入的對照常數，不參與計算。",
          "月營收單位 $M/月＝人數（M）× ARPU（$/月）；年營收＝月營收×月數，月數換算只在報告文字，不進公式。")
    header(ws, 4, ["編號", "項目", "值（公式）", "v0.5 原值（對照）", "差距＝值−v0.5（公式）", "單位", "說明"])
    S = lambda m: sid(R, m)  # noqa: E731
    g = R.get
    sf = "tokenRevenue/subscription/subsFYAvg"
    yrs = {y: i for i, y in enumerate((2025, 2026, 2027, 2028, 2029, 2030))}

    def others(y):
        return "(" + "+".join(inp_id(R, f"FY 平均用戶數：{x}", str(y)) for x in ("go", "plus", "seats")) + ")"

    def v05_others(y):
        return sum(g(f"{sf}/{x}/v/{yrs[y]}") for x in ("go", "plus", "seats"))

    ratio26 = inp_id(R, "Pro 占付費訂閱總數比例（2026）")
    ratio = {y: inp_id(R, "Pro 占付費訂閱總數比例（2027+", str(y)) for y in (2027, 2028, 2029, 2030)}
    kfree, kplus = inp_id(R, "換算係數：free"), inp_id(R, "換算係數：Plus")
    arpu25, arpu26 = inp_id(R, "ARPU（實收）：pro", "2025"), inp_id(R, "ARPU（實收）：pro", "2026")
    mult, wau, p35 = S("Pro 用戶倍數：2026 對 2025（內部預測：倍增）"), S("ChatGPT 週活躍用戶（WAU）：2026-02"), S("ChatGPT Plus＋Pro 付費訂閱：2025-07")
    up_row = src_row(R, "Pro 占付費訂閱總數：2026 上限（內部預測：<1%）")
    v05 = {y: g(f"{sf}/pro/v/{yrs[y]}") for y in yrs}
    v05_arpu25, v05_arpu26 = g("tokenRevenue/subscription/arpu/pro/v/0"), g("tokenRevenue/subscription/arpu/pro/v/1")
    rate, mins, div = inp_id(R, "API FY2025 平均每分鐘"), inp_id(R, "每年分鐘數"), inp_id(R, "單位換算：B token")
    rows = [  # (code, 項目, 公式, v0.5 值, 單位, 說明, 名稱)
        ("V01", "free FY2026 平均用戶（M）＝WAU×換算係數", f"={wau}*{kfree}", g(f"{sf}/free/v/1"), "M", "SRC_OAI WAU × Inputs 換算係數", "V9_Free2026"),
        ("V02", "Plus FY2025 平均用戶（M）＝2025-07 Plus＋Pro×換算係數", f"={p35}*{kplus}", g(f"{sf}/plus/v/0"), "M", "SRC_OAI 35M × Inputs 換算係數", "V9_Plus2025"),
        ("V03", "Pro FY2026 用戶（M）＝比例×付費總數（含席位）", f"={ratio26}/(1-{ratio26})*{others(2026)}", v05[2026], "M",
         "付費總數＝go＋plus＋seats＋pro；pro＝比例×總數 ⇒ pro＝比例/(1−比例)×(go＋plus＋seats)，無循環", "V9_Pro2026"),
        ("V04", "Pro FY2025 用戶（M）＝2026÷SRC 倍數", f"=C7/{mult}", v05[2025], "M", "SRC_OAI 倍增約束", "V9_Pro2025"),
        ("V05", "Pro 占付費總數 2026（含席位）", f"=C7/({others(2026)}+C7)", v05[2026] / (v05_others(2026) + v05[2026]), "比例",
         "v0.5 原值欄為 v0.5 pro 1.0 ÷ (go＋plus＋seats＋pro)", "V9_ProShare2026"),
        ("V06", "約束檢查：Pro 占比超出 SRC 上限的部分（0＝符合）", f"=C9-MIN(C9,SRC_OAI!$E${up_row})", None, "", "SRC_OAI『<1%』上限（僅 2026）", "V9_ProViolation"),
        ("V07", "v0.5 原值對約束的差距（百分點）＝v0.5 占比−SRC 上限", f"=(D9-SRC_OAI!$E${up_row})", None, "比例", "正值＝v0.5 原值超出內部預測所稱『<1%』", None),
        ("V08", "Pro 月訂閱營收 2025（$M/月）＝用戶×ARPU", f"=C8*{arpu25}", v05[2025] * v05_arpu25, "$M/月", "ARPU 為 Inputs（2025 欄）", "V9_ProRev2025"),
        ("V09", "Pro 月訂閱營收 2026（$M/月）＝用戶×ARPU", f"=C7*{arpu26}", v05[2026] * v05_arpu26, "$M/月", "ARPU 為 Inputs（2026 起欄）", "V9_ProRev2026"),
    ]
    for k, y in enumerate((2027, 2028, 2029, 2030)):      # V10–V13：pro 2027–2030（V15）
        rows.append((f"V{10 + k}", f"Pro FY{y} 用戶（M）＝比例_t/(1−比例_t)×(go＋plus＋seats)", f"={ratio[y]}/(1-{ratio[y]})*{others(y)}", v05[y], "M",
                     "V15：逐年比例（Inputs，0.5–2.0%）", f"V9_Pro{y}"))
    for k, y in enumerate((2027, 2028, 2029, 2030)):      # V14–V17：隱含占比（v0.5 原值的隱含占比並列）
        rows.append((f"V{14 + k}", f"Pro 占付費總數 {y}（含席位）", f"=C{14 + k}/({others(y)}+C{14 + k})",
                     v05[y] / (v05_others(y) + v05[y]), "比例", "v0.5 原值欄＝v0.5 pro ÷ (go＋plus＋seats＋pro)（約 1.48%→1.7%）", None))
    for k, y in enumerate((2027, 2028, 2029, 2030)):      # V18–V21：月營收
        rows.append((f"V{18 + k}", f"Pro 月訂閱營收 {y}（$M/月）＝用戶×ARPU", f"=C{14 + k}*{arpu26}", v05[y] * v05_arpu26, "$M/月", "ARPU 為 Inputs（2026 起欄）", f"V9_ProRev{y}"))
    rows += [
        ("V22", "API FY2025 token（T）＝每分鐘 token×每年分鐘數÷單位換算", f"={rate}*{mins}/{div}", g("tokenRevenue/actuals/apiTokensFY2025/v"), "T", "E8a：驅動值＝Inputs 5B/min", "V9_ApiTok"),
        ("V23", "API FY2025 token 低（T）", f"=Inputs!$F${inp_row(R, rate)}*{mins}/{div}", g("tokenRevenue/actuals/apiTokensFY2025/lo"), "T", "E8a：驅動值低端 3.8", None),
        ("V24", "API FY2025 token 高（T）", f"=Inputs!$G${inp_row(R, rate)}*{mins}/{div}", g("tokenRevenue/actuals/apiTokensFY2025/hi"), "T", "E8a：驅動值高端 6.1", None),
        ("V25", "Oracle 合約年額（$B/年）＝總額÷年數", f'=IF({S("合約期間迄：Oracle")}>={S("合約期間起：Oracle")},{S("合約總額：Oracle")}/({S("合約期間迄：Oracle")}-{S("合約期間起：Oracle")}+1),"期間無效")', g("spending/contracts/oracle/annual"), "$B", "E8f：v0.5 以數值 60 寫入，改公式", "V9_OracleAnnual"),
        ("V26", "AWS Trainium WSJ 口徑合計 GW＝推論＋訓練", f'={S("合約容量（WSJ 口徑）：AWS Trainium 推論")}+{S("合約容量（WSJ 口徑）：AWS Trainium 訓練（VR）")}', 5, "GW",
         "E8g：『5』改公式（v0.5 區間上緣 5）", "V9_AwsWsj"),
        ("V27", "2026-03 輪無條件部分＝四個組成加總", "=" + "+".join(S(f"股權融資 2026-03 輪：{x}") for x in ("Amazon 首筆", "SoftBank（三期）", "Nvidia", "其他")),
         g("funding/equityRound2026Mar/unconditional2026/v"), "$B", "E8j：87 拆為 SRC 組成", "V9_Uncond2026"),
    ]
    for i, (code, lab, f, v05v, unit, note, name) in enumerate(rows):
        n = 5 + i
        assert code == f"V{i + 1:02d}", (code, i)
        put(ws, f"A{n}", code, F_CALC)
        put(ws, f"B{n}", lab, F_CALC)
        put(ws, f"C{n}", f)
        if v05v is not None:
            put(ws, f"D{n}", v05v, F_CALC)
            put(ws, f"E{n}", f'=IF(ISNUMBER(C{n}),C{n}-D{n},"—")')
        put(ws, f"F{n}", unit, F_CALC)
        put(ws, f"G{n}", note, F_NOTE)
        if name:
            nm(wb, name, f"Derived_V9!$C${n}")
    ws.column_dimensions["B"].width = 60
    ws.column_dimensions["G"].width = 70


def sheet_checks(wb, R, n_src, n_inp, n_pending, P2, P3, P4, P5):
    ws = wb.create_sheet("Checks")
    title(ws, "Checks — P1–P5 檢查（公式；結果 ERR 的格數＝CHK_Errors）",
          "OK／ERR：比對期望值；INFO：只列示。編號（C##）由 builder/id_registry.json 固定：新檢查取下一個號碼，退役號碼不重用（F1）。期望值為 builder 寫入的常數（黑字）。")
    header(ws, 4, ["編號", "檢查項", "值（公式）", "期望", "結果", "說明"])
    S = lambda m: sid(R, m)  # noqa: E731
    tm = [(inp_id(R, "API token 層級占比：top 層", f"陣列[{i}]"), inp_id(R, "API token 層級占比：mid 層", f"陣列[{i}]")) for i in (0, 1)]
    rowof = lambda iid: inp_row(R, iid)  # noqa: E731

    def low_viol():
        parts = []
        for t, m in tm:
            rt, rm = rowof(t), rowof(m)
            for ct, cm in (("E", "E"), ("G", "G"), ("F", "F")):         # 基準、（top 高＋mid 高→low 最低）、（top 低＋mid 低→low 最高）
                L = f"(1-Inputs!${ct}${rt}-Inputs!${cm}${rm})"
                parts.append(f"({L}<0)+({L}>1)")
        return "=" + "+".join(parts)

    uncond = "+".join(S(f"股權融資 2026-03 輪：{x}") for x in ("Amazon 首筆", "SoftBank（三期）", "Nvidia", "其他"))
    aws_hi_row = rowof(inp_id(R, "合約容量（採用值）：AWS（Trainium）"))
    rows = [
        ("SRC_OAI 列數", "=SUMPRODUCT(--(LEN(SRC_OAI!$A$5:$A$400)>0))", n_src, "eq", "報告所列筆數"),
        ("SRC_OAI 缺值列數（Active）", "=SUM(SRC_OAI!$V$5:$V$400)", 0, "eq", "數值、低、高皆空且非文字列"),
        ("SRC_OAI 缺出處列數（Active）", "=SUM(SRC_OAI!$W$5:$W$400)", 0, "eq", ""),
        ("SRC_OAI 區間順序異常列數", "=SUM(SRC_OAI!$X$5:$X$400)", 0, "eq", "低 ≤ 值 ≤ 高"),
        ("Inputs 列數", "=SUMPRODUCT(--(LEN(Inputs!$A$5:$A$600)>0))", n_inp, "eq", ""),
        ("Inputs 區間順序異常列數", "=SUM(Inputs!$M$5:$M$600)", 0, "eq", "低 ≤ 值 ≤ 高（有三值者）"),
        ("TK_Link 已取值名稱數", '=COUNTIF(TK_Link!$F$10:$F$200,"OK")', None, "info", "取自 Tokenomics master CURRENT"),
        ("TK_Link 待 Tokenomics 提供名稱數", f'=COUNTIF(TK_Link!$F$10:$F$200,"{PENDING_NOTE}")', n_pending, "eq", "工作單要求、Tokenomics 現行版尚無的名稱（值留空；v5.24 快照後為 0）"),
        ("TK_Link 錯誤值格數", "=SUMPRODUCT(--ISERROR(TK_Link!$J$10:$X$200))", 0, "eq", ""),
        ("TK_IF_Util 介於 (0,1]（違反數）", "=IF(AND(TK_IF_Util>0,TK_IF_Util<=1),0,1)", 0, "eq", "Tokenomics 基準利用率"),
        ("2025 推論支出：付費＋非付費−合計（$B）",
         f"={S('推論成本：2025 付費用戶')}+{S('推論成本：2025 非付費用戶')}-{S('推論成本合計：2025')}", 0, "tol", "v0.5 三項同源（The Information 2026-02）"),
        ("2026-03 融資輪：無條件＋條件式−總額（$B）",
         f"={uncond}+{S('股權融資：條件式部分（Amazon）')}-{S('股權融資：2026-03 輪總額')}", 0, "tol", "122＝(15＋30＋30＋12)＋35（E8j：無條件部分由四個 SRC 組成加總）"),
        ("合約期間起>迄的合約數",
         f"=({S('合約期間起：Oracle')}>{S('合約期間迄：Oracle')})+({S('合約期間起：AWS（Nvidia 晶片）')}>{S('合約期間迄：AWS（Nvidia 晶片）')})+({S('合約期間起：AWS（Trainium）')}>{S('合約期間迄：AWS（Trainium）')})+({S('合約期間起：CoreWeave')}>{S('合約期間迄：CoreWeave')})", 0, "eq", "Oracle、AWS×2、CoreWeave"),
        ("預覽：Oracle 隱含每 GW 年合約價（$B/GW/年）", f'=IF({S("合約期間迄：Oracle")}-{S("合約期間起：Oracle")}+1>0,{S("合約總額：Oracle")}/({S("合約期間迄：Oracle")}-{S("合約期間起：Oracle")}+1)/{S("合約容量：Oracle")},"期間無效")', None, "info", "P3 V2 可比（Oracle 遠期合約隱含 13.3；V8）"),
        ("預覽：AWS Trainium 隱含每 GW 年價（$B/GW/年）", f'=IF({S("合約期間迄：AWS（Trainium）")}-{S("合約期間起：AWS（Trainium）")}+1>0,{S("合約總額：AWS（Trainium）")}/({S("合約期間迄：AWS（Trainium）")}-{S("合約期間起：AWS（Trainium）")}+1)/{S("合約容量：AWS（Trainium）")},"期間無效")', None, "info", "自研晶片，只並列（工作單 P3-3）"),
        ("預覽：FY2025 推論支出−營業成本（$B）", f"={S('推論成本合計：2025')}-{S('營業成本（cost of revenue）：FY2025')}", None, "info",
         "推論 8.4 高於營業成本 7.5：口徑差異，P3／P4 對帳時說明"),
        ("預覽：非算力研發 2025（$B）＝研發總額−訓練支出公式（V4）", f"={S('研發費用總額：2025')}-C{{TRN_B}}", None, "info", "工作單 V4：≈7.19（區間 5.68–8.59）"),
        ("V6 衍生值：訓練支出 2025 基準（$B）＝SRC 10.59＋其他雲端 1.41", f"={S('研發費用中付 Microsoft 部分：2025')}+Inputs!$E${{OC_ROW}}", 12, "tol", "v0.5 值 12.0（對照）"),
        ("V6 衍生值：訓練支出 2025 低（$B）", f"={S('研發費用中付 Microsoft 部分：2025')}+Inputs!$F${{OC_ROW}}", 10.59, "tol", "v0.5 低 10.6"),
        ("V6 衍生值：訓練支出 2025 高（$B）", f"={S('研發費用中付 Microsoft 部分：2025')}+Inputs!$G${{OC_ROW}}", 13.5, "tol", "v0.5 高 13.5"),
        ("V12：Pro 2026 占付費總數超出 SRC 上限的部分（0＝符合）", "=V9_ProViolation", 0, "eq", "Inputs 比例須使 pro 占比 ≤ 內部預測『<1%』"),
        ("V8：每 GW 年合約價參數列數（Inputs）", '=COUNTIF(Inputs!$B$5:$B$600,"*合約價*")', 1, "eq", "本模型只有一個合約價參數"),
        ("E8d：tokenMix 的 low＝1−top−mid 在基準／低／高情境皆 ∈[0,1]（違反數）", low_viol(), 0, "eq", "兩個時點 × 三個情境；low 不另設區間"),
        ("E8g：Inputs AWS 容量上限＝WSJ 推論＋訓練（GW）", f"=Inputs!$G${aws_hi_row}-V9_AwsWsj", 0, "tol", "『5』由 SRC 兩筆揭露相加"),
        ("E8j：無條件部分四個組成加總＝v0.5 值（$B）", "=V9_Uncond2026", R.get("funding/equityRound2026Mar/unconditional2026/v"), "tol", "15＋30＋30＋12"),
        ("E8f：Oracle 合約年額＝v0.5 值（$B/年）", "=V9_OracleAnnual", R.get("spending/contracts/oracle/annual"), "tol", "總額÷年數"),
    ]
    rows += p2_checks(R, P2, S, rowof)
    rows += p3_checks(R, P3, S, rowof)
    rows += p4_checks(R, P3, P4, S, rowof)
    rows += p5_checks(R, P2, P3, P4, P5, S, rowof)
    # F1：Checks 編號納入穩定 ID registry（label → C##）
    reg = json.loads(REGISTRY_PATH.read_text(encoding="utf-8"))
    known = reg.setdefault("CHK", {})
    nxt = max([int(v[1:]) for v in known.values()] + [0]) + 1
    ids = []
    for r_ in rows:
        if r_[0] not in known:
            known[r_[0]] = f"C{nxt:02d}"
            nxt += 1
        ids.append(known[r_[0]])
    assert len(set(ids)) == len(ids)
    for k in [k for k in known if k not in {r_[0] for r_ in rows}]:
        reg.setdefault("retired_CHK", []).append(known.pop(k))
    REGISTRY_PATH.write_text(json.dumps(reg, ensure_ascii=False, indent=0), encoding="utf-8")
    order = sorted(range(len(rows)), key=lambda i: int(ids[i][1:]))   # 數字排序（C100 起不得排在 C91 之前）
    rows = [rows[i] for i in order]
    ids = [ids[i] for i in order]
    oc_row = rowof(inp_id(R, "2025 其他雲端"))
    trn_b = 5 + next(i for i, r_ in enumerate(rows) if r_[0].startswith("V6 衍生值：訓練支出 2025 基準"))
    for i, (lab, f, exp, kind, note) in enumerate(rows):
        f = f.replace("{OC_ROW}", str(oc_row)).replace("C{TRN_B}", f"C{trn_b}")
        n = 5 + i
        put(ws, f"A{n}", ids[i], F_CALC)
        put(ws, f"B{n}", lab, F_CALC)
        put(ws, f"C{n}", f, F_CALC)
        if exp is not None:
            put(ws, f"D{n}", exp, F_CALC)
        if kind == "eq":
            put(ws, f"E{n}", f'=IF(C{n}=D{n},"OK","ERR")')
        elif kind == "tol":
            put(ws, f"E{n}", f'=IF(ISNUMBER(C{n}),IF(ABS(C{n}-D{n})<0.000000001,"OK","ERR"),"ERR")')
        elif kind == "warn":
            put(ws, f"E{n}", "WARN")
        else:
            put(ws, f"E{n}", "INFO")
        put(ws, f"F{n}", note, F_NOTE)
    last = 4 + len(rows)
    put(ws, f"A{last + 2}", "CHK_Errors", F_BOLD)
    put(ws, f"C{last + 2}", f'=COUNTIF($E$5:$E${last},"ERR")')
    nm(wb, "CHK_Errors", f"Checks!$C${last + 2}")
    ws.column_dimensions["B"].width = 52
    ws.column_dimensions["F"].width = 60


def p2_checks(R, P2, S, rowof):
    """v0.6-P2 新增的檢查（編號取 registry 下一號）。C27（E8c WARN）於 P2 撤除，編號退役不重用。"""
    D, V = P2["D"], P2["V"]
    dr, vr = D.rows, V.rows
    mix = "+".join(f'(ABS(Inputs!$E${rowof(inp_id(R, f"方案模型組合：{p}", "luna"))}+Inputs!$E${rowof(inp_id(R, f"方案模型組合：{p}", "sol"))}'
                   f'+Inputs!$E${rowof(inp_id(R, f"方案模型組合：{p}", "astra"))}-1)>0.000000001)' for p in p2.PLANS)
    ev = {k: V.ev_first + i for i, k in enumerate((2, 3, 4, 5))}
    order = "+".join(f"(Revenue!{c}{ev[3]}<Revenue!{c}{ev[4]})+(Revenue!{c}{ev[2]}<Revenue!{c}{ev[5]})" for c in "DEF")
    lpv = "+".join(f"(Revenue!G{r}>Revenue!D{r})+(Revenue!D{r}>Revenue!H{r})" for r in range(V.ev_last - 5, V.ev_last + 1))
    pairs = [("情境 base：價格彈性：任務數", "價格彈性：任務數"), ("情境 base：價格彈性：每任務 token", "價格彈性：每任務 token"),
             ("情境 base：每任務 token 年增率", "每任務 token 年增率"), ("情境 base：牌價年變動率 top 層", "牌價年變動率：top 層"),
             ("情境 base：牌價年變動率 mid 層", "牌價年變動率：mid 層"), ("情境 base：牌價年變動率 low 層", "牌價年變動率：low 層")]
    d11 = "+".join(f"({inp_id(R, a)}<>{inp_id(R, b)})" for a, b in pairs)
    tok_a, tok_b = dr["tok_all"], dr["tok_chk"]
    return [
        ("P2：2025 模型營收總額 − SRC 實際營收（$B；v0.5 校準恆等式）", "=REV_Gap2025", 0, "tol", "2025 API 為倒推項；不為 0 表示校準鏈斷裂"),
        ("P2：2025 API 計費比例 ∈ (0,1]（違反數）", "=IF(AND(DEM_ApiBilledRatio2025>0,DEM_ApiBilledRatio2025<=1),0,1)", 0, "eq",
         "計費 token（營收倒推）÷ 處理 token（Derived_V9 V22）；≤0 表示訂閱＋廣告已超過實際營收"),
        ("P2：方案模型組合 luna＋sol＋astra≠1 的方案數", "=" + mix, 0, "eq", "Inputs INP_155–179；token 層級拆分的前提"),
        ("P2：token 層級拆分合計 vs 拆分前總量（違反年數）",
         f"=SUMPRODUCT(--(ABS(Demand!$D${tok_a}:$I${tok_a}-Demand!$D${tok_b}:$I${tok_b})>0.000000001*Demand!$D${tok_b}:$I${tok_b}))", 0, "eq",
         "付費＋免費（層級）＝訂閱＋免費＋API 計費 token"),
        ("P2：消費端＋企業端＋其他 − 總額（各年絕對值合計，$B）", f"=SUMPRODUCT(ABS(Revenue!$D${vr['sumchk']}:$I${vr['sumchk']}))", 0, "tol", "v0.5 檢查列"),
        ("P2：Microsoft 分成累計超出上限的部分（$B）", f"=MAX(0,MAX(REV_MSCum)-{S('Microsoft 分成總額上限')})", 0, "tol", "上限 $38B（SRC_OAI）"),
        ("E8c：價格事件生效日順序違反數（基準／提前／延後三欄）", "=" + order, 0, "eq",
         "top 層：事件 3（促銷）須不晚於事件 4（Astra）；mid／low 層：事件 2 須不晚於事件 5"),
        ("E8c：2026 牌價區間 低 ≤ 基準 ≤ 高 的違反數（6 列）", "=" + lpv, 0, "eq", "listPriceFYAvg 2026 區間由事件時點 Inputs 傳遞（取代 C27 WARN）"),
        ("D11：Inputs 情境 base 列＝各參數基準值（違反數）", "=" + d11, 0, "eq", "本版未設情境選擇器；基準引用各參數 Inputs 值欄，情境列只作對照"),
        ("P2 預覽：2025 token 合計 ÷ Tokenomics IF_AllocDemand", "=DEM_TkRatio2025", None, "info", "口徑不同，只列差距（工作單 P2-3）"),
        ("P2 預覽：2025 未校準差距（$B）＝訂閱＋廣告＋其他＋處理 token×單價 − 實際", "=REV_GapUncal2025", None, "info", "自下而上口徑對實際"),
        ("P2 預覽：2030 淨營收（$B）", f"=Revenue!$I${vr['net']}", None, "info", "總額 − Microsoft 分成"),
    ]


def p3_checks(R, P3, S, rowof):
    """v0.6-P3 新增的檢查（編號取 registry 下一號）。"""
    C, V = P3["C"], P3["V"]
    cr, vr = C.rows, V.rows
    yr = lambda key: f"Compute!$D${cr[key]}:$I${cr[key]}"  # noqa: E731
    yr26 = lambda key: f"Compute!$E${cr[key]}:$I${cr[key]}"  # noqa: E731
    shares = "+".join(f"SUMPRODUCT(--({yr('s_' + g)}<0))" for g in p3.GEN_KEYS)
    g3r = rowof(inp_id(R, "GB300 占 Blackwell 比例", "2025")), rowof(inp_id(R, "GB300 占 Blackwell 比例", "2026 起"))
    g3 = "+".join(f"(Inputs!$E${r}<0)+(Inputs!$E${r}>1)" for r in g3r)
    labels = "+".join(f"(COUNTIF(TK_HdrGen,Compute!${c}${cr['g_lab']})=0)" for c in "DEFGH")
    sup_sum = ",".join(f"Compute!$D${cr['sup_' + k]}:$I${cr['sup_' + k]}" for k in P3["contracts"]) + f",Compute!$D${cr['own']}:$I${cr['own']}"   # S4：＋自建 GW
    return [
        ("P3：世代占比合計≠1 的年數＋任一世代占比 <0 的格數", f"=SUMPRODUCT(--(ABS({yr('s_sum')}-1)>0.000000001))+{shares}", 0, "eq",
         "六世代（Hopper、GB200、GB300、VR200、Rubin Ultra、自研／其他）；自研＝1−其餘，不得為負"),
        ("P3：GB300 占 Blackwell 比例 ∉ [0,1] 的列數", "=" + g3, 0, "eq", "Inputs（r6 P3-2）"),
        ("P3：Compute 世代名稱在 TK_HdrGen 找不到的數", "=" + labels, 0, "eq", "SUMIFS 取 TK_IF_TokGW_* 的鍵；Tokenomics 改名時轉 ERR"),
        ("P3：η（2025）≤0 或非數值（違反數）", "=IF(ISNUMBER(CMP_Eta2025),IF(CMP_Eta2025>0,0,1),1)", 0, "eq", "V2"),
        ("P3：V2 恆等式 η × 支出換算 GW − token 換算 GW（2025，GW）", f"=CMP_Eta2025*CMP_SpendGW2025-Compute!$D${cr['gw_tok']}", 0, "tol", "計算鏈可逐列追出"),
        ("P3：2025 有效推論 GW − 支出換算 GW（GW）", f"=Compute!$D${cr['eff']}-CMP_SpendGW2025", 0, "tol", "η 定義使 2025 兩者相等"),
        ("P3：推論（截頂後）＋研發＋閒置 − 供給（2026–2030 各年差的最大絕對值，GW）", f"=MAX(MAX({yr26('ident')}),-MIN({yr26('ident')}))", 0, "tol", "v0.5 算力MW 第 36 列（以 MAX／MIN 取最大絕對值：pycel 的 SUMPRODUCT(ABS()) 在全為 0 時回傳整數型別，ISNUMBER 判定與 Excel 不同）"),
        ("P3：容量上限係數 ∉ (0,1] 的年數", f"=SUMPRODUCT(--({yr('cap')}<=0))+SUMPRODUCT(--({yr('cap')}>1.000000001))", 0, "eq", "V1a"),
        ("P3：截頂後總額 > 未截頂總額 的年數", f"=SUMPRODUCT(--(Revenue!$D${vr['gross_c']}:$I${vr['gross_c']}>Revenue!$D${vr['gross']}:$I${vr['gross']}+0.000000001))", 0, "eq", ""),
        ("P3：研發 GW < 0 的年數", f"=SUMPRODUCT(--({yr('rd')}<0))", 0, "eq", "供給 ×（1−閒置）不足以容納截頂後推論時轉 ERR"),
        ("P3：逐合約加總 − 合約供給合計（絕對值合計，GW）", f"=ABS(SUM({sup_sum})-SUM({yr('sup')}))", 0, "tol", "揭露／未揭露小計＋自建 GW（S4 起計入供給）與合計一致"),
        ("P3：截頂後 Microsoft 分成累計超出上限的部分（$B）", f"=MAX(0,MAX(REV_MSCumCapped)-{S('Microsoft 分成總額上限')})", 0, "tol", "上限 $38B（SRC_OAI）"),
        ("P3 預覽：η（2025）", "=CMP_Eta2025", None, "info", "工作單參考值 0.196 為 Tokenomics v5.14 Block 6 需求口徑；本模型需求口徑不同（P2），以實算為準"),
        ("P3 並列：本模型 2025 推論 GW（token 換算）", f"=Compute!$D${cr['gw_tok']}", None, "info", "V11：本模型組合 × TK 世代產能"),
        ("P3 並列：TK IF_AllocServeGW（2025 服務 GW；Tokenomics Alloc 口徑）", "=TK_IF_AllocServeGW", None, "info", "V11：Alloc 對應值只在檢查頁並列"),
        ("P3 並列：TK Alloc 隱含每 GW 年產能＝IF_AllocDemand ÷ IF_AllocServeGW（M tok/GW/年）", "=TK_IF_AllocDemand/TK_IF_AllocServeGW", None, "info",
         "對照本模型 2025 付費＋免費加權產能（下一列）"),
        ("P3 並列：本模型 2025 每 GW 年產能（付費＋免費加權＝token 合計 ÷ token GW）", f"=(Compute!$D${cr['tok_paid']}+Compute!$D${cr['tok_free']})/Compute!$D${cr['gw_tok']}", None, "info", ""),
        ("P3 並列：本模型 2025 研發 GW（訓練支出 ÷ 合約價）", f"=Compute!$D${cr['rd']}", None, "info", "r6 P3-4"),
        ("P3 並列：TK IF_AllocRDGW（研發 GW·年，物理下限）", "=TK_IF_AllocRDGW", None, "info", "只並列（r6 P3-4）"),
        ("P3 並列：TK IF_AllocImpliedNk（隱含 N × k）", "=TK_IF_AllocImpliedNk", None, "info", "只並列（r6 P3-4）"),
        ("P3 三角對照：2025 隱含每 GW 年算力支出 − 合約價（$B/GW/年）", f"=Compute!$D${cr['t_gap']}", None, "info", "r6 P3-3：約 16.3 − 12；只列差距"),
        ("P3 對照：2025 總需求 GW − 年均算力 GW（0.6／1.9 平均）", "=CMP_DemGapAvg2025", None, "info", "r6 P3-5；只列差距"),
        ("P3 預覽：容量截頂年數（2026–2030）", "=SUM(CMP_CapFlag)", None, "info", "V1a 旗標"),
        ("P3 對照：計畫算力 2026–30 累計 ÷ 合約價（GW·年）", f"={S('計畫算力支出：2026–2030 累計')}/{inp_id(R, '每 GW 年合約價')}", None, "info",
         "SRC_OAI 856（600–856）；只作檢查（r6 P3-4）"),
        ("P3 對照：本模型 2026–30 合約供給 GW·年合計", f"=SUM(Compute!$E${cr['sup']}:$I${cr['sup']})", None, "info", ""),
        ("P3 對照：本模型 2026–30 總需求 GW·年合計", f"=SUM(Compute!$E${cr['dem']}:$I${cr['dem']})", None, "info", ""),
        ("P3 預覽：2030 有效推論 GW", f"=Compute!$I${cr['eff']}", None, "info", "η 依 Inputs 路徑開關（基準沿用）"),
    ]


def p4_checks(R, P3, P4, S, rowof):
    """v0.6-P4 新增的檢查（編號取 registry 下一號）。"""
    C, K = P3["C"], P4["K"]
    cr, kr = C.rows, K.rows
    yk = lambda key: f"Cost!$D${kr[key]}:$I${kr[key]}"  # noqa: E731
    yc = lambda key: f"Compute!$D${cr[key]}:$I${cr[key]}"  # noqa: E731
    mx = lambda key: f"=MAX(MAX({yk(key)}),-MIN({yk(key)}))"  # noqa: E731
    return [
        ("P4：逐合約實付加總 − 合約實付合計（各年差的最大絕對值，$B）", mx("ck_pay"), 0, "tol", "v0.5 支出第 11 列"),
        ("P4：未揭露 GW 合約（Azure、AWS Nvidia、CoreWeave）|實付 − 供給 GW × 合約價|（各年最大值，$B）", mx("ck_und"), 0, "tol",
         "S3：未揭露 GW＝付款 ÷ 合約價；成本與 GW 同一時程"),
        ("P4：算力成本（現金）−（合約實付＋自有資本支出扣合作方）（最大絕對值，$B）", mx("ck_cc"), 0, "tol", "V3"),
        ("P4：推論＋研發＋閒置算力 − 算力成本（Q3 口徑）（最大絕對值，$B）", mx("ck_split"), 0, "tol", "依 GW 拆分（v0.5 算力MW 第 47 列）；2025＝實際"),
        ("P4：非算力研發＋銷售＋管理（不含股權報酬）− 合計（最大絕對值，$B）", mx("ck_nc"), 0, "tol", "功能別占比加總＝1"),
        ("P4：2025 費用對帳：營收 −（營業成本＋研發＋銷售＋管理）＋營業損失（$B）",
         f"={S('營收：FY2025')}-({S('營業成本（cost of revenue）：FY2025')}+{S('研發費用總額：2025')}+{S('銷售費用（sales and marketing）：FY2025')}"
         f"+{S('管理費用（general and administrative）：FY2025')})+{S('營業損失（loss from operations）：FY2025')}", None, "info", "外流財報各列四捨五入；應接近 0"),
        ("P4：非算力營運費用（不含股權報酬）≤0 的年數", f"=SUMPRODUCT(--({yk('ncx')}<=0))", 0, "eq", "2025＝財報 − 股權報酬；股權報酬超過財報費用時轉 ERR"),
        ("P4：員工人數 ≤0 的年數＋自建 GW <0 的年數", f"=SUMPRODUCT(--({yk('hc')}<=0))+SUMPRODUCT(--({yc('own')}<0))", 0, "eq", ""),
        ("P4：Q3 ∉ [0,1] 的年數", f"=SUMPRODUCT(--({yk('q3')}<0))+SUMPRODUCT(--({yk('q3')}>1))", 0, "eq", "r6 P4-2"),
        ("P4 預覽：2025 算力成本（合約實付）− 實際算力支出（$B）", f"=Cost!$D${kr['gap25']}", None, "info", "只列差距"),
        ("P4 預覽：2025 非算力研發（財報口徑，$B）", "=COST_NonCompRD2025", None, "info", "r6：≈7.19（5.68–8.59）"),
        ("P4 預覽：2025 非算力營運費用（財報口徑）÷ 營收 vs v0.5 1.216", f"=Cost!$D${kr['v05pct']}", None, "info", "研發不含付 Microsoft＋銷售＋管理"),
        ("P4 預覽：2025 股權報酬（$B）÷ WSJ 46.2% × 營收", f"=Cost!$D${kr['sbc']}/Cost!$D${kr['wsj']}", None, "info", "人數 × 每人 vs 占營收比，兩個 WSJ 數字的一致性"),
        ("P4 預覽：2030 雲端毛利率（基準持有成本）", f"=Cost!$I${kr['gmp']}", None, "info", "V3"),
        ("P4 預覽：2025 Q3", f"=Cost!$D${kr['q3']}", None, "info", "r6 P4-2"),
        ("P4 並列：2025 研發算力 − TK L1_Ans3（$B/年）", f"=Cost!$D${kr['q_gap']}", None, "info", "只列差距"),
        ("P4 預覽：2030 每 VR 等值 GW 差額（含股權報酬；現金口徑，$B/GW/年）", f"=Cost!$I${kr['p_gap_vr']}", None, "info", "命題"),
        ("P4 預覽：2030 每 VR 等值 GW 差額（含股權報酬；經濟口徑，$B/GW/年）", f"=Cost!$I${kr['e_gap_vr']}", None, "info", "命題（經濟口徑）"),
        ("P4 預覽：2025–2030 累計差額（含股權報酬；現金口徑，$B）", f"=SUM({yk('p_gap_b')})", None, "info", "P5 融資需求的起點（不含既有現金與融資）"),
        ("P4 預覽：差額 <0 的年數（含股權報酬；現金口徑）", f"=SUMPRODUCT(--({yk('p_gap_vr')}<0))", None, "info", "6＝每一年營收都未覆蓋全成本"),
        ("P4 V16：2025 Azure 攤入額超過實際算力支出（旗標）", "=COST_V16Flag", None, "info", "S4 預設：只立旗標；基準起點維持 2025"),
        ("P4 V16：起點 2026 情境的 2025 攤入額差異（$B）", f"=Cost!$D${kr['v_diff']}", None, "info", "起點 2026 時 2025 攤入額為 0"),
        ("P4 並列：TK 每 VR200 GW 年全成本（下游預設）−（IF_FullCost）（$B/GW/年）", f"=Cost!$D${kr['tk_def']}-Cost!$D${kr['tk_full']}", None, "info",
         "S4 預設：下游預設為基準，IF_FullCost 並列"),
        ("P4 預覽：2030 自建 GW", f"=Compute!$I${cr['own']}", None, "info", "S4 預設：自建 GW 計入供給"),
    ]


def p5_checks(R, P2, P3, P4, P5, S, rowof):
    """v0.6-P5 新增的檢查（編號取 registry 下一號）：融資恆等式與缺口、2025 實際值對帳（r6 P5-2）、反向模式、TK 快照狀態。"""
    F, X, K = P5["F"], P5["X"], P4["K"]
    fr, xr, kr = F.rows, X.rows, K.rows
    cr, vr = P3["C"].rows, P2["V"].rows
    yf = lambda key, a="E": f"Funding!${a}${fr[key]}:$I${fr[key]}"  # noqa: E731
    yx = lambda key, a="E": f"Reverse!${a}${xr[key]}:$I${xr[key]}"  # noqa: E731
    mx = lambda rng: f"=MAX(MAX({rng}),-MIN({rng}))"  # noqa: E731
    fcf_id = "+".join(f"ABS(Funding!{c}{fr['rev']}-Funding!{c}{fr['cc']}-Funding!{c}{fr['ncx']}-Funding!{c}{fr['fcf']})" for c in "DEFGHI")
    cash_id = "+".join(f"ABS(Funding!{c}{fr['open']}+Funding!{c}{fr['fcf']}+Funding!{c}{fr['eq']}+Funding!{c}{fr['ext']}-Funding!{c}{fr['end']})" for c in "EFGHI")
    gap_sbc = "+".join(f"ABS(Funding!{c}{fr['fcf']}-Cost!{c}{kr['p_gap_b']}-Cost!{c}{kr['sbc']})" for c in "DEFGHI")
    tgt_up = "+".join(f"(Reverse!{c}{xr['tgt']}<=Reverse!{p}{xr['tgt']})" for p, c in zip("DEFGH", "EFGHI"))
    rev25 = f"Revenue!$D${vr['gross_c']}"
    amz = inp_id(R, "計入 Amazon 條件式 $35B")
    return [
        # ── 融資（OK／ERR）
        ("P5：自由現金流 − （營收淨額 − 算力成本 − 非算力成本）（各年絕對值合計，$B）", "=" + fcf_id, 0, "tol", "現金流恆等式"),
        ("P5：自由現金流 − （COST_GapCash ＋ COST_SBC）（各年絕對值合計，$B）", "=" + gap_sbc, 0, "tol", "S4 給 S5：股權報酬為非現金，加回"),
        ("P5：年底現金 −（期初＋自由現金流＋股權流入＋外部資金）（2026–2030 絕對值合計，$B）", "=" + cash_id, 0, "tol", "現金流量恆等式（來源順序 ①②③）"),
        ("P5：已到位融資＋Nvidia 非現金部分 − 2026-03 輪無條件部分（$B）",
         f"=SUM({yf('committed', 'D')})+SUM({yf('nv_nc', 'D')})-V9_Uncond2026", 0, "tol", "到位年須落在 2025–2030；組成＝SRC 四列（E8j）"),
        ("P5：最低現金 <0 或 2025 年底現金 <0（違反數）", f"=({S('現金：2025 年底')}<0)+(Funding!$D${fr['min']}<0)", 0, "eq", "S7"),
        ("P5：Amazon 條件式開關 ∉ {0,1}（違反數）", f"=IF(OR({S('股權融資：條件式部分（Amazon）')}<0,AND({amz}<>0,{amz}<>1)),1,0)", 0, "eq", "S5 預設：基準 0（不計入）"),
        ("P5：外部資金需求 <0 的年數＋年底現金低於最低現金的年數（2026–2030）",
         f"=SUMPRODUCT(--({yf('ext')}<0))+SUMPRODUCT(--({yf('end')}<{yf('min')}-0.000000001))", 0, "eq", "補足至最低現金的定義"),
        # ── 融資（INFO）
        ("P5 結果：2025–2030 累計外部資金需求（基準，$B）", f"=Funding!$I${fr['cum']}", None, "info", "r6 P5-1"),
        ("P5 結果：單年外部資金需求峰值（$B）", f"=MAX({yf('ext')})", None, "info", ""),
        ("P5 結果：峰值年度", f"=MIN({yf('peak_y')})", None, "info", ""),
        ("P5 結果：首次需要外部資金的年度（現金跌破最低現金）", f'=IF(COUNT({yf("flag_y")})=0,"無",MIN({yf("flag_y")}))', None, "info", "S7 缺口旗標"),
        ("P5 結果：只靠已到位融資，現金首次低於最低現金的年度", f'=IF(COUNT({yf("noext_y")})=0,"無",MIN({yf("noext_y")}))', None, "info",
         "已到位融資能撐到此年度的前一年"),
        ("P5 結果：缺口旗標年數（2026–2030）", f"=SUM({yf('flag')})", None, "info", ""),
        ("P5 結果：2030 年底現金（基準，$B）", f"=Funding!$I${fr['end']}", None, "info", "＝最低現金（有缺口時）"),
        ("P5 情境：計入 Amazon 條件式 $35B 的累計外部資金需求（$B）", f"=Funding!$I${fr['s_cum']}", None, "info", "S5 預設：不計入基準，列情境"),
        ("P5 情境：Amazon 條件式使累計外部資金需求減少（$B）", f"=Funding!$I${fr['cum']}-Funding!$I${fr['s_cum']}", None, "info", ""),
        ("P5 或有負債合計（只列示，$B）", "=FND_Contingent", None, "info", "S4b：Nvidia 擔保＋晶片融資洽談；不作資金來源"),
        ("P5 或有：合約剩餘承諾（2030 年後，$B）", "=FND_CommitAfter2030", None, "info", "v0.5 支出 J 欄"),
        # ── 2025 實際值對帳（r6 P5-2）
        ("P5 對帳 2025 營收：模型總額（截頂後）− SRC 實際（$B）", f"={rev25}-{S('營收：FY2025')}", 0, "tol", "v0.5 校準：2025 API 為倒推項"),
        ("P5 對帳 2025 推論支出：有效推論 GW × 合約價 − SRC 8.4（$B）", f"=Compute!$D${cr['eff']}*{inp_id(R, '每 GW 年合約價')}-{S('推論成本合計：2025')}", 0, "tol",
         "V2：η 以 2025 推論支出校準，故為 0"),
        ("P5 對帳 2025 推論支出 − 財報營業成本（$B）", f"={S('推論成本合計：2025')}-{S('營業成本（cost of revenue）：FY2025')}", None, "info",
         "推論 8.4 高於營業成本 7.5（口徑差異；營業成本可能不含部分推論或已扣折扣）"),
        ("P5 對帳 2025 訓練支出：研發 GW × 合約價 −（10.59＋其他雲端）（$B）",
         f"=Compute!$D${cr['rd']}*{inp_id(R, '每 GW 年合約價')}-({S('研發費用中付 Microsoft 部分：2025')}+{inp_id(R, '2025 其他雲端')})", 0, "tol", "r6 P3-4：2025 研發 GW＝訓練支出 ÷ 合約價"),
        ("P5 對帳 2025 GW：模型總需求 GW − SRC 2025 年底 1.9（GW）", f"=Compute!$D${cr['dem']}-{S('算力規模：2025 年底')}", None, "info", "只列差距（口徑未明：年底 vs 年均）"),
        ("P5 對帳 2025 GW：合約推得供給 GW − SRC 年均（(0.6＋1.9)÷2）（GW）", f"=Compute!$D${cr['sup']}-Compute!$D${cr['t_avg']}", None, "info", "S3：合約推得 1.23 對年均 1.25（C65 同源）"),
        ("P5 對帳 2025 算力成本：合約實付 − 實際算力支出 20.4（$B）", f"=Cost!$D${kr['cc']}-COST_Actual2025", None, "info", "S4 預設：基準用合約實付（r6）"),
        ("P5 對帳 2025 營業損益：模型（總額 − 實際算力 − 非算力 − 股權報酬）（$B）",
         f"={rev25}-COST_Actual2025-Cost!$D${kr['ncx']}-Cost!$D${kr['sbc']}", None, "info", "以實際算力支出 20.4 計"),
        ("P5 對帳 2025 營業損益：模型 − 財報（−營業損失）（$B）",
         f"=({rev25}-COST_Actual2025-Cost!$D${kr['ncx']}-Cost!$D${kr['sbc']})+{S('營業損失（loss from operations）：FY2025')}", None, "info",
         "差距主要＝推論 8.4 − 營業成本 7.5（下一列為扣除後殘差）"),
        ("P5 對帳 2025 營業損益殘差：差距 ＋（推論 − 營業成本）（$B）",
         f"=({rev25}-COST_Actual2025-Cost!$D${kr['ncx']}-Cost!$D${kr['sbc']})+{S('營業損失（loss from operations）：FY2025')}+{S('推論成本合計：2025')}-{S('營業成本（cost of revenue）：FY2025')}",
         None, "info", "≈0（財報各列四捨五入）"),
        ("P5 對帳 2025 Microsoft 分成（現金流已扣；財報位置未揭露，$B）", f"=Revenue!$D${vr['ms_c']}", None, "info", "若財報營業成本已含分成，現金流有重複扣除之虞（資料缺口）"),
        ("P5 對帳 2025 自由現金流（實際算力支出口徑，$B）", "=FND_FCF2025Act", None, "info", "只對照；現金流自 2026 起"),
        # ── 反向模式
        ("P5 反向：目標路徑 2026–2030 合計 − SRC Σ840（$B）", f"=Reverse!$D${xr['tsum_ck']}", 0, "tol", "2029＝殘差，恆為 0"),
        ("P5 反向：目標路徑非遞增的年數（2026–2030）", "=" + tgt_up, 0, "eq", "擬合值（Inputs）使 2029 殘差低於前一年時轉 ERR"),
        ("P5 反向：Microsoft 分成累計超出上限的部分（$B）", f"=MAX(0,MAX(RVS_MSCum)-{S('Microsoft 分成總額上限')})", 0, "tol", "同上限規則"),
        ("P5 反向 R2：2030 只靠 API 的倍數", f"=Reverse!$I${xr['m_api']}", None, "info", "v0.5 4.55"),
        ("P5 反向 R2：2030 只靠訂閱的倍數", f"=Reverse!$I${xr['m_sub']}", None, "info", "v0.5 5.66"),
        ("P5 反向 R2：2030 兩線等比例倍數", f"=Reverse!$I${xr['m_prop']}", None, "info", "v0.5 3.01"),
        ("P5 反向：2025–2030 累計外部資金需求（目標營收、同一支出，$B）", f"=Reverse!$I${xr['cum']}", None, "info", "v0.5 115.5"),
        ("P5 反向：Σ自由現金流 2026–30 − 簡報 −278（$B）", f"=Reverse!$D${xr['r_diff']}", None, "info", "v0.5 +118.3"),
        # ── TK 快照狀態（r6 P5-2）
        ("P5 TK 快照：Tokenomics 版本", "=TK_Version", None, "info", "check_tk_snapshot 結果見報告"),
        ("P5 TK 快照：Tokenomics master 提交", "=TK_Commit", None, "info", ""),
    ]


def sheet_readme(wb, snap, date, summary):
    ws = wb.active
    ws.title = "README"
    title(ws, f"OpenAI 收支模型 {VERSION}（{date}）",
          "命題（Andy 2026-09-30）：OpenAI 每 VR 等值 GW 的營收能否覆蓋每 GW 全成本；若不能，缺口由誰、以什麼條件融資。FY2025–FY2030，曆年制。")
    lines = [
        ("本版範圍", "P1：repo 骨架、SRC_OAI（公司財務原始數據）、TK_Link（Tokenomics 快照）、Inputs（v0.5 Assumed 項目遷入）。P2：Demand（需求與 token 量）、Revenue（營收、Microsoft 分成、淨額）。P3：Compute（算力需求與供給、η、容量上限、VR 等值）、Revenue 第八節（截頂後營收）。P4：Cost（算力成本、供應商持有成本與雲端毛利、非算力成本、Q3、命題輸出、V16 Azure 檢查）、Compute 第十三節（自建 GW 計入供給）。融資於 P5。"),
        ("Excel 為唯一計算引擎", "藍字＝輸入（Excel 擁有）；黑字＝公式；綠字＝跨頁連結。builder 只產生結構，重建時保留 Excel 內已改過的藍字。"),
        ("SRC_OAI", f"{summary['src']} 列；v0.5『已取得的原始訊息』逐筆遷入（Interested-party／Verified）。"),
        ("TK_Link", f"Tokenomics {snap['version']}（{snap['file']}），master 提交 {snap['sha'][:7]}，讀取日 {date}；{summary['tk_ok']} 個名稱有值、{summary['tk_pending']} 個待 Tokenomics 提供。"),
        ("Inputs", f"{summary['inp']} 列；v0.5 Assumed／Analogy／Decision 項目（長表）。"),
        ("Map_v05", f"v0.5 JSON {summary['leaves']} 個葉節點的去處。"),
        ("Demand", "P2 需求：訂閱各方案人數、每日任務數 × 每任務 token、API 任務數與計費 token；token 量（層級 × 付費／免費，具名範圍 DEM_Tok_*，供 P3）。"),
        ("Revenue", "P2 營收：API 單價路徑（2026 依價格事件時點天數加權）、訂閱、API、廣告、其他、總額、Microsoft 分成、淨額（具名範圍 REV_*）。"),
        ("Compute", "P3 算力：V1 世代組合（Blackwell 拆 GB200／GB300；自研＝VR200 × 係數）、付費／免費每 GW 年產能（TK IF_TokGW × IF_Util）、推論 GW（token 換算）、V2 η（2025 推論支出 ÷ 合約價校準）、有效推論 GW、三角對照、研發 GW、總需求 GW、供給 GW（逐合約）、容量上限（V1a）、VR 等值換算、敏感度（具名範圍 CMP_*）。"),
        ("Cost", "P4 成本：算力成本（V3：合約實付＋自有資本支出；現金口徑）、供應商持有成本（TK IF_HoldEcon 世代加權）與雲端毛利、經濟口徑、非算力成本（V4：2025 財報；2026 起人數 × 每人年成本，股權報酬單列）、Q3（對 TK L1_Ans3）、命題輸出（每 VR 等值 GW 的營收、成本、差額；現金與經濟口徑）、V16 Azure 攤入檢查（具名範圍 COST_*）。"),
        ("Funding", "P5 融資（S7）：自由現金流（營收淨額 − 算力成本 − 非算力成本；股權報酬加回）、已到位融資（2026-03 輪無條件部分）、最低現金、外部資金需求（年底現金低於最低現金的累計差額）、缺口旗標、Amazon 條件式情境、或有負債（S4b，只列示）（具名範圍 FND_*）。"),
        ("Reverse", "P5 反向模式（R1–R3）：管理層營收目標路徑、三組充分條件倍數、2030 可觀測量、橋接、反向資金、與簡報對帳；只作對照，不回饋基準（具名範圍 RVS_*）。"),
        ("Checks", "P1–P5 檢查（含 2025 實際值對帳、Tokenomics 快照狀態）；CHK_Errors 必須為 0。"),
        ("一手／二手圖例", "『二手（含一手公告轉載）』＝媒體轉載公司公開公告；『二手（含一手轉載）』＝F2，媒體轉載外流財報或內部文件；兩者不合併。（同文見 SRC_OAI 頁首說明）"),
        ("標記", "Verified／Interested-party／Analogy／Assumed／Derived（Analogy、Assumed 一律附區間）。"),
        ("分層", "公司財務原始數據→SRC_OAI；AI 技術與算力→TK_Link（取自 Tokenomics 名稱）；假設→Inputs。"),
    ]
    for i, (a, b) in enumerate(lines):
        put(ws, f"A{4 + i}", a, F_BOLD)
        put(ws, f"B{4 + i}", b, F_CALC, wrap=True)
    ws.column_dimensions["A"].width = 22
    ws.column_dimensions["B"].width = 120


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--tk-dir", required=True, type=Path)
    ap.add_argument("--out", required=True, type=Path)
    ap.add_argument("--base", type=Path)
    ap.add_argument("--date", default="2026-10-04")
    ap.add_argument("--no-lo", action="store_true", help="略過 LibreOffice 重算（輸出無快取值；僅供 LibreOffice 不可用時暫用，需在報告註明）")
    a = ap.parse_args()
    R = build_map(write_registry=True)
    miss = coverage_report(R)
    if miss:
        sys.exit(f"v0.5 葉節點未分類：{miss[:10]}")
    snap = read_snapshot(a.tk_dir)
    outdir = Path("build_out")            # 紀錄檔（restore_log、leafmap、build_summary）不放進 model/
    outdir.mkdir(parents=True, exist_ok=True)
    a.out.parent.mkdir(parents=True, exist_ok=True)

    defaults = {}
    for r in R.src_rows:
        for col, k in zip("CDE", ("value", "lo", "hi")):
            defaults[("SRC_OAI", r["id"], col)] = blank(r[k])
    for r in R.inp_rows:
        for col, k in zip("EFG", ("value", "lo", "hi")):
            defaults[("Inputs", r["id"], col)] = blank(r[k])
    if a.base:
        final, matched, kept, unmatched = preserve.merge(defaults, a.base, outdir / "restore_log.txt")
    else:
        final, matched, kept, unmatched = dict(defaults), 0, 0, []
        (outdir / "restore_log.txt").write_text("matched: 0; Excel value kept over code default: 0; unmatched: 0\n（首次建置，無底稿）\n", encoding="utf-8")
    print(f"restore: matched {matched}, Excel kept over code {kept}, unmatched {len(unmatched)}")

    wb = openpyxl.Workbook()
    summary = dict(src=len(R.src_rows), inp=len(R.inp_rows), leaves=len(R.leaves), tk_ok=len(snap["rows"]),
                   tk_pending=len(snap["pending"]))
    sheet_readme(wb, snap, a.date, summary)
    sheet_src(wb, R, final)
    sheet_tk(wb, snap, a.date)
    sheet_inputs(wb, R, final)
    sheet_derived(wb, R)
    P2 = p2.build_p2(wb, R, lambda m: sid(R, m), lambda n, i=None: inp_id(R, n, i), lambda iid: inp_row(R, iid))
    P3 = p3.build_p3(wb, R, lambda m: sid(R, m), lambda n, i=None: inp_id(R, n, i), lambda iid: inp_row(R, iid), P2, snap)
    P4 = p4.build_p4(wb, R, lambda m: sid(R, m), lambda n, i=None: inp_id(R, n, i), lambda iid: inp_row(R, iid), P2, P3, snap)
    P5 = p5.build_p5(wb, R, lambda m: sid(R, m), lambda n, i=None: inp_id(R, n, i), lambda iid: inp_row(R, iid), P2, P3, P4)
    sheet_checks(wb, R, len(R.src_rows), len(R.inp_rows), summary["tk_pending"], P2, P3, P4, P5)
    sheet_map(wb, R)
    preserve.write_defaults(wb, final)
    raw = Path(tempfile.mkdtemp()) / a.out.name
    wb.save(raw)
    if a.no_lo:
        raw.replace(a.out)
    else:
        lo_recalc(raw, Path(tempfile.mkdtemp())).replace(a.out)
    with open(outdir / "leafmap.csv", "w", newline="", encoding="utf-8-sig") as f:
        w = csv.writer(f)
        w.writerow(["v0.5 路徑", "v0.5 值", "去處類別", "去處", "理由／備註"])
        for p, v in R.leaves.items():
            k, ref, why = R.dest[p]
            w.writerow([p.replace(SEP, "."), json.dumps(v, ensure_ascii=False), KIND_LABEL[k], ref, why])
    json.dump(dict(summary, pending=R.pending, tk_file=snap["file"], tk_sha=snap["sha"]), open(outdir / "build_summary.json", "w", encoding="utf-8"),
              ensure_ascii=False, indent=1)
    print("saved", a.out, summary)


if __name__ == "__main__":
    main()
