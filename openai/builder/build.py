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
from tk_link import BLOCK6_EXTRA_V515, PENDING_NOTE, read_snapshot  # noqa: E402
from v05_map import REGISTRY_PATH, SEP, build_map, coverage_report  # noqa: E402

VERSION = "v0.6-P1"
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
          "每列一個 Tokenomics 名稱（SRC_／IF_）。值（藍字）由 builder 從 Tokenomics master 的 model/CURRENT 讀出寫入；本模型公式只引用 TK_ 具名範圍。更新快照＝本模型的修補版。",
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
        vals = [row["name"], "（待 v5.15 合併；無具名範圍）", f"v5.15 分支實際名稱：{row['v515']}", "", 0, row["status"], snap["version"], snap["sha"][:7], date]
        for c, v in enumerate(vals, start=1):
            put(ws, f"{L(c)}{r}", v, F_CALC)
        r += 1
    for row_name in BLOCK6_EXTRA_V515:
        for c, v in enumerate([row_name, "（待 v5.15 合併；無具名範圍）", "v5.15 新增、工作單未列", "", 0, PENDING_NOTE, snap["version"], snap["sha"][:7], date], start=1):
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


def sheet_checks(wb, R, n_src, n_inp, n_pending):
    ws = wb.create_sheet("Checks")
    title(ws, "Checks — P1 檢查（公式；結果 ERR 的格數＝CHK_Errors）",
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
        ("TK_Link 待 v5.15 合併名稱數", f'=COUNTIF(TK_Link!$F$10:$F$200,"{PENDING_NOTE}")', n_pending, "eq", "Block 6（IF_Alloc*）與 SRC_DEM_010–013；值留空"),
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
        ("E8c：listPriceFYAvg 2026 區間尚未傳遞（待 P2 事件時點 Inputs）", '="待 P2"', None, "warn", "WARN：不計入 CHK_Errors；P2 以價格事件時點 Inputs 傳遞 2026 區間後移除"),
    ]
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
    order = sorted(range(len(rows)), key=lambda i: ids[i])
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


def sheet_readme(wb, snap, date, summary):
    ws = wb.active
    ws.title = "README"
    title(ws, f"OpenAI 收支模型 {VERSION}（{date}）",
          "命題（Andy 2026-09-30）：OpenAI 每 VR 等值 GW 的營收能否覆蓋每 GW 全成本；若不能，缺口由誰、以什麼條件融資。FY2025–FY2030，曆年制。")
    lines = [
        ("本版範圍", "P1：repo 骨架、SRC_OAI（公司財務原始數據）、TK_Link（Tokenomics 快照）、Inputs（v0.5 Assumed 項目遷入）。需求、算力、成本、融資於 P2–P5。"),
        ("Excel 為唯一計算引擎", "藍字＝輸入（Excel 擁有）；黑字＝公式；綠字＝跨頁連結。builder 只產生結構，重建時保留 Excel 內已改過的藍字。"),
        ("SRC_OAI", f"{summary['src']} 列；v0.5『已取得的原始訊息』逐筆遷入（Interested-party／Verified）。"),
        ("TK_Link", f"Tokenomics {snap['version']}（{snap['file']}），master 提交 {snap['sha'][:7]}，讀取日 {date}；{summary['tk_ok']} 個名稱有值、{summary['tk_pending']} 個待 v5.15 合併。"),
        ("Inputs", f"{summary['inp']} 列；v0.5 Assumed／Analogy／Decision 項目（長表）。"),
        ("Map_v05", f"v0.5 JSON {summary['leaves']} 個葉節點的去處。"),
        ("Checks", "P1 檢查；CHK_Errors 必須為 0。"),
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
                   tk_pending=len(snap["pending"]) + len(BLOCK6_EXTRA_V515))
    sheet_readme(wb, snap, a.date, summary)
    sheet_src(wb, R, final)
    sheet_tk(wb, snap, a.date)
    sheet_inputs(wb, R, final)
    sheet_derived(wb, R)
    sheet_checks(wb, R, len(R.src_rows), len(R.inp_rows), summary["tk_pending"])
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
