#!/usr/bin/env python3
"""v0.1.7.1 對照 Excel 與報告數字（CLAUDE.md 第 5 節、第 8 節；工作單 20261102 v0.1.7.1）。

①新增或變動的列（SRC_ANT_472–487 新列、V1 改 Verified 的 46 列與 427–431 改值、Inputs INP_168 與 INP_218–220、Funding 改式列與第十三節尾／第十四節新列、Checks C198–C202）
②關鍵輸出 FY2025–2030：v0.1.7 vs V2 前（只關 V2）vs v0.1.7.1
③V2 瀑布（v0.1.7 → V2 前 → ＋各輪原文 → ＋Google G-2＝v0.1.7.1；另列只計 Google G-2）
④最低現金月數 3／6／12（v0.1.7 與 v0.1.7.1）
⑤已套用的預設（含替代讀法量化）
⑥S-1 替換對照（CLAUDE.md 第 8 節第 4 步）
⑦範圍外變動核對（逐格比對結果；由 tools/diff_xlsx.py 的 JSON 讀入）
情境一律以 engine（pycel）重算現行活頁簿（覆寫值取自 tools/html_scenarios.yaml 的群組定義，與 HTML 同一套）；本程式不做模型計算。
用法：python3 tools/make_report_v0171.py --out docs/reports/20261102_v0.1.7.1_對照.xlsx [--json build_out/report_v0171.json] [--diff-json build_out/diff_v017_off.json]
"""
import argparse
import json
import sys
from pathlib import Path

import openpyxl
import yaml
from openpyxl.styles import Font

REPO = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO))
sys.path.insert(0, str(REPO / "tools"))
from build_html import Locator, load_scenarios, resolve_overrides  # noqa: E402
from engine import Engine, current_model_path, resolve_key  # noqa: E402

YEARS = [2025, 2026, 2027, 2028, 2029, 2030]
KEY = [("FND_Committed", "已到位融資（$B）"), ("FND_FCF", "營運自由現金流（$B）"), ("FND_CashNoExt", "只靠已到位融資的年底現金（$B）"), ("FND_MinCash", "最低現金（$B）"),
       ("FND_ExtNeed", "當年外部資金需求（$B）"), ("FND_ExtNeedCum", "累計外部資金需求（$B）"), ("FND_CashEnd", "年底現金（$B）"), ("FND_Headroom", "現金餘裕（$B）"),
       ("COST_PropGap_VR", "每 VR 等值 GW 差額（$B/GW/年）"), ("COST_Coverage", "覆蓋率"), ("REV_NetCapped", "營收淨額（$B）"), ("COST_Compute", "算力成本（$B）"),
       ("COST_NonCompExSBC", "非算力成本（不含股權報酬，$B）"), ("FND_ExtNeedCumCond", "S2 條件式：累計外部資金需求（$B）"), ("FND_ExtNeedCumAll", "S3 IPO＋條件式：累計外部資金需求（$B）")]
SCAL = [("FND_CashNoExtTrough", "現金谷底（$B）"), ("FND_FirstGapYear", "首次缺口年"), ("COST_GapPosSustainYear", "命題 1 持續轉正年"), ("FND_FCFPosSustainYear", "營運自由現金流持續轉正年"),
        ("FND_ExtNeedPeak", "單年外部資金需求峰值（$B）"), ("FND_CommittedS1Gap", "已到位融資 − S-1 特別股收入（$B）"), ("FND_Cash0630Gap", "S-1 6-30 現金 − 模型近似（$B）"),
        ("FND_Cash0630Resid", "6-30 差額拆解其餘項（$B）"), ("FND_L1_ExtTotal", "第 1 層累計外部資金需求（$B）"), ("FND_L3_ExtTotal", "第 3 層累計外部資金需求（$B）"),
        ("CHK_Errors", "CHK_Errors"), ("CHK_Warnings", "CHK_Warnings")]
READS = [("v0.1.7", "v017"), ("V2 前（只關 V2；V1、V3、V4 已套用）", "v0171_v2only_off"), ("v0.1.7.1", None)]
STEPS = [("v0.1.7", "v017"), ("V2 前（V1、V3、V4 已套用）", "v0171_v2only_off"), ("＋各輪改原文 30.2／65.2", "v0171_round_only"),
         ("＋Google Series G-2 計入＝v0.1.7.1", None), ("說明：只計入 Google G-2（各輪仍 30／65）", "v0171_googl_only")]
ALTS = [("未歸屬餘額付款期 快 2027–2031", "s1u_fast"), ("未歸屬餘額付款期 慢 2028–2036", "s1u_slow"), ("SpaceXAI 以 S-1 上限 120.5 計", "spx_cap"),
        ("Google 額外 TPU 計入（中點）", "gaddl_on"), ("Broadcom 起始年 2026", "bcm_2026"), ("AMD＝S-1 下限 20", "amd_s1"), ("只關 W2（Fluidstack 舊法）", "only_w2_off"),
        ("只關 W4", "only_w4_off"), ("W4 兩項皆不用", "w4_neither"), ("W4 只動 Q3 口徑", "w4_q3only"), ("W4 只動股權報酬", "w4_sbconly"),
        ("承諾不進基準（W1、W2 皆關）", "commit_off"), ("只關 W1", "commit_off_w1only"), ("API 任務數成長取低", "api_task_lo"), ("價格彈性取弱端", "elast_weak"),
        ("每 GW 合約價 20", "price_hi"), ("任務數低＋彈性弱", "combo1")]
MC = [("3 個月", "v017_mc_lo", "mc_lo"), ("6 個月（基準）", "v017", None), ("12 個月", "v017_mc_hi", "mc_hi")]
DEFAULTS = [
    ("Google 2026 Q2 Series G-2 約 10", "計入基準（原文事實：Note 4、9、10 與 Series H 分開；googl_upfront_switch＝1）", "舊讀法 0（視為 Series H 的一部分；v0.1.7）", "v0171_round_only"),
    ("Series G／H 金額", "改原文 30.2／65.2（s1f_round_switch＝1；差 +0.4）", "新聞稿 30／65", "v0171_googl_only"),
    ("Google G-2 金額取法", "Google 股數 38.6 百萬 × 每股 259.1364 ÷ 1000＝10.003（原文 Note 4、10 的股數與價格）", "G-2 合計 15.0 − Amazon 10-Q 5.0＝10.0（差 −0.003）", None),
    ("Amazon Series G-2 5.0 的來源", "維持 Amazon 10-Q 5.0（SRC_ANT_319，Verified；S-1 股數 × 價格＝5.001，只對照）", "改用 S-1 股數 × 價格", None),
    ("受限現金 2.5（Broadcom 租約保證金）", "不另計（之後抵付租金）；只在 6-30 現金對照列示", "自 2026 已到位或年底現金扣除 2.5", None),
    ("條件式融資（V3）", "Google 2026 NPA 剩餘 20（IPO 認購選擇權 10 與之互抵、不相加）；Amazon 15；只動情境 S2／S3、第 ③ 層、回流對照 E 欄", "Google 30（v0.1.7）", "v0171_v3_off"),
    ("最低現金（V4）", "維持 D17：次年非算力成本 × 6 個月（Assumed 3–12）；契約下限 5.0 只對照；循環額度 15.0 只列或有", "3／12 個月（見 ④）", "mc_hi"),
    ("V1 改 Verified 的 party 欄", "工作單文字「Anthropic（正式 S-1；期中未經查核）」一律套用；年度報表附註（SRC_ANT_422、448）在備註註明經查核", "年度列另寫「年度經查核」", None),
    ("V1 的 source 欄", "改為正式 S-1（SEC EDGAR）；原二手來源與網址移入備註（不刪）", "保留二手來源文字", None),
    ("SRC_ANT_467（股權報酬 2.2）", "值不改（工作單 V1 只改 427–431 的精度）；備註寫原文 2.185", "改 2.185（只對照列，不影響基準）", None),
    ("N1 SpaceXAI、N2 期後短期 21、N3 有條件融資、既有六題", "照舊（工作單）", "—", None),
]
REPL = [("SRC_ANT_302", "SRC_ANT_473", "Funding 第二節 Series G（s1f_round_switch）"),
        ("SRC_ANT_298", "SRC_ANT_478", "Funding 第二節 Series H（s1f_round_switch）"),
        ("SRC_ANT_300", "SRC_ANT_480＋481", "Funding 第二節「其中：雲端業者」（只列示）"),
        ("SRC_ANT_328", "SRC_ANT_475 × 477", "Funding 第二節 Google G-2（googl_upfront_switch 0 → 1）"),
        ("SRC_ANT_329", "SRC_ANT_482", "Funding 第三節 Google 條件式（s1f_cond_switch；只動情境）"),
        ("SRC_ANT_322", "SRC_ANT_484", "Funding 第三節 Amazon 條件式（同值）"),
        ("SRC_ANT_306", "SRC_ANT_485", "Funding 第六節循環信貸（或有；同值）")]


def named_vals(eng, n):
    v = eng.get_name(n)
    return v if isinstance(v, list) else [v]


def run(eng, loc, cfg, sc=None):
    ov = resolve_overrides(sc, loc, cfg) if sc else {}
    orig = {k: eng.get(*resolve_key(loc.names, k)) for k in ov}
    try:
        for k, v in ov.items():
            eng.set_key(k, v)
        out = {n: named_vals(eng, n) for n, _ in KEY}
        out.update({n: eng.get_name(n) for n, _ in SCAL})
        out["F23_2027"] = eng.get("Funding", f"F{loc.rowid[('Funding', 'F23')]}")
        out["F17_2026"] = eng.get("Funding", f"E{loc.rowid[('Funding', 'F17')]}")
        return out
    finally:
        for k, v in orig.items():
            eng.set_key(k, v)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--out", required=True, type=Path)
    ap.add_argument("--json", type=Path, default=REPO / "build_out" / "report_v0171.json")
    ap.add_argument("--diff-json", type=Path, nargs="*", default=[])
    a = ap.parse_args()
    model = current_model_path()
    loc, cfg, eng = Locator(model), load_scenarios(), Engine(model)
    scs = {s for _l, s in READS + STEPS} | {x for _l, p, q in MC for x in (p, q)} | {d[3] for d in DEFAULTS if d[3]} | {"v0171_v3_off", "v016"} | {s for _l, s in ALTS} | {"v017_" + s for _l, s in ALTS if "v017_" + s in cfg["scenarios"]}
    R = {(s or "base"): run(eng, loc, cfg, s) for s in scs}

    wb = openpyxl.Workbook()
    B = Font(bold=True)

    def sheet(title, header):
        ws = wb.create_sheet(title)
        ws.append(header)
        for c in ws[1]:
            c.font = B
        return ws

    mv = openpyxl.load_workbook(model, data_only=True)
    old = openpyxl.load_workbook(REPO / "model" / "archive" / "20261031_Anthropic_v0.1.7.xlsx", data_only=True)
    # ① 新增與變動列
    ws = sheet("①新增與變動列", ["頁", "列 ID", "列名／項目", "值（D 或 E）", "低／F", "高／G", "v0.1.7 值", "標記", "來源／說明"])
    srcy = {r["id"]: r for r in yaml.safe_load((REPO / "data" / "anthropic_src.yaml").read_text(encoding="utf-8"))["rows"]}
    oldy = {}
    wo = old["SRC_ANT"]
    for r in range(5, wo.max_row + 1):
        if str(wo.cell(r, 1).value or "").startswith("SRC_ANT_"):
            oldy[wo.cell(r, 1).value] = (wo.cell(r, 5).value, wo.cell(r, 16).value)
    for rid, r in srcy.items():
        n = int(rid[8:])
        o = oldy.get(rid)
        if o is None or o[1] != r["tag"] or (isinstance(o[0], (int, float)) and isinstance(r["value"], (int, float)) and abs(o[0] - r["value"]) > 1e-12):
            ws.append(["SRC_ANT", rid, r["metric"], r["value"], r["lo"], r["hi"], (None if o is None else o[0]), r["tag"] + ("" if o is None else f"（v0.1.7：{o[1]}）"),
                       ("新增" if o is None else "V1 改標" + ("＋改值" if o[0] != r["value"] else "")) + "；" + str(r["source"])[:60]])
        elif n in (298, 300, 302, 306, 322, 328, 329, 421):
            ws.append(["SRC_ANT", rid, r["metric"], r["value"], r["lo"], r["hi"], o[0], r["tag"], "備註加註取代或原文核對（值不改）"])
    wi, woi = mv["Inputs"], old["Inputs"]
    for r in range(5, wi.max_row + 1):
        iid = wi.cell(r, 1).value
        if isinstance(iid, str) and iid.startswith("INP_") and (int(iid[4:]) >= 218 or iid == "INP_168"):
            ov_ = woi.cell(r, 5).value if iid == "INP_168" else None
            ws.append(["Inputs", iid, wi.cell(r, 2).value, wi.cell(r, 5).value, wi.cell(r, 6).value, wi.cell(r, 7).value, ov_, wi.cell(r, 8).value, wi.cell(r, 9).value])
    wf = mv["Funding"]
    chg = ("F12", "F13", "F14", "F15", "F17", "F18", "F20", "F22", "F23", "F81", "F88", "F89", "F49")
    for r in range(1, wf.max_row + 1):
        rid = wf.cell(r, 1).value
        if rid in chg:
            ws.append(["Funding", rid, "改式／改標籤：" + str(wf.cell(r, 2).value), wf.cell(r, 4).value, wf.cell(r, 5).value, wf.cell(r, 6).value,
                       old["Funding"].cell(r, 5).value, None, wf.cell(r, 11).value])
    st = next(r for r in range(1, wf.max_row + 1) if str(wf.cell(r, 2).value or "").startswith("其中：Broadcom 租約保證金"))
    for r in range(st, wf.max_row + 1):
        if wf.cell(r, 1).value is None and wf.cell(r, 2).value is None:
            continue
        ws.append(["Funding", wf.cell(r, 1).value, wf.cell(r, 2).value] + [wf.cell(r, c).value for c in (4, 5, 6)] + [None, None, wf.cell(r, 11).value])
    wc = mv["Checks"]
    for r in range(5, wc.max_row + 1):
        cid = wc.cell(r, 1).value
        if isinstance(cid, str) and cid.startswith("C") and cid[1:].isdigit() and int(cid[1:]) >= 198:
            ws.append(["Checks", cid, wc.cell(r, 2).value, wc.cell(r, 3).value, wc.cell(r, 4).value, wc.cell(r, 5).value, None, None, wc.cell(r, 6).value])

    # ② 關鍵輸出
    ws = sheet("②關鍵輸出", ["輸出", "讀法"] + [str(y) for y in YEARS])
    for n, lab in KEY:
        for rl, sc in READS:
            ws.append([lab, rl] + R[sc or "base"][n])
    for n, lab in SCAL:
        for rl, sc in READS:
            ws.append([lab, rl, R[sc or "base"][n]])

    # ③ V2 瀑布
    ws = sheet("③V2瀑布", ["輸出"] + [l for l, _s in STEPS])
    rows3 = [("2026 已到位融資", "FND_Committed", 1)] + [(f"每 VR 等值 GW 差額 {y}", "COST_PropGap_VR", YEARS.index(y)) for y in YEARS[1:]] + \
            [("覆蓋率 2030", "COST_Coverage", 5), ("累計外部資金需求 2026–2030", "FND_ExtNeedCum", 5), ("首次缺口年", "FND_FirstGapYear", None),
             ("現金谷底（只靠已到位融資）", "FND_CashNoExtTrough", None), ("2029 年底現金（只靠已到位）", "FND_CashNoExt", 4), ("2029 最低現金", "FND_MinCash", 4),
             ("命題 1 持續轉正年", "COST_GapPosSustainYear", None), ("營運自由現金流持續轉正年", "FND_FCFPosSustainYear", None),
             ("已到位融資 − S-1 特別股收入", "FND_CommittedS1Gap", None), ("CHK_Warnings", "CHK_Warnings", None)]
    for lab, n, i in rows3:
        ws.append([lab] + [(R[sc or "base"][n][i] if i is not None else R[sc or "base"][n]) for _l, sc in STEPS])

    # ④ 最低現金 3／6／12
    ws = sheet("④最低現金3_6_12", ["最低現金月數", "2029 最低現金", "v0.1.7 累計外部資金", "v0.1.7 首次缺口", "v0.1.7 谷底", "v0.1.7.1 累計外部資金", "v0.1.7.1 首次缺口", "v0.1.7.1 谷底", "v0.1.7.1 單年峰值"])
    for lab, p, q in MC:
        o, w = R[p], R[q or "base"]
        ws.append([lab, w["FND_MinCash"][4], o["FND_ExtNeedCum"][5], o["FND_FirstGapYear"], o["FND_CashNoExtTrough"], w["FND_ExtNeedCum"][5], w["FND_FirstGapYear"],
                   w["FND_CashNoExtTrough"], w["FND_ExtNeedPeak"]])

    # ⑤ 已套用的預設
    ws = sheet("⑤已套用的預設", ["問題", "採用的預設", "替代", "替代：累計外部資金", "替代：現金谷底", "替代：首次缺口", "替代：S2 條件式累計", "替代：2026 已到位"])
    for q, d, alt, sc in DEFAULTS:
        row = [q, d, alt]
        if sc:
            o = R[sc]
            row += [o["FND_ExtNeedCum"][5], o["FND_CashNoExtTrough"], o["FND_FirstGapYear"], o["FND_ExtNeedCumCond"][5], o["FND_Committed"][1]]
        ws.append(row)

    ws = sheet("⑤b替代讀法", ["讀法", "v0.1.7 累計外部資金", "v0.1.7 谷底", "v0.1.7.1 累計外部資金", "v0.1.7.1 首次缺口", "v0.1.7.1 谷底", "v0.1.7.1 2030 差額"])
    for lab, sc in ALTS:
        o = R.get("v017_" + sc)
        w = R[sc]
        ws.append([lab, o["FND_ExtNeedCum"][5] if o else None, o["FND_CashNoExtTrough"] if o else None, w["FND_ExtNeedCum"][5], w["FND_FirstGapYear"],
                   w["FND_CashNoExtTrough"], w["COST_PropGap_VR"][5]])

    # ⑥ S-1 替換對照
    ws = sheet("⑥S-1替換對照", ["舊列", "舊值", "新列", "新值", "差異比例", "受影響（本版改指新列）", "命題輸出前（v0.1.7）→後（v0.1.7.1）", "結論是否翻轉"])
    sv = {}
    wsrc = mv["SRC_ANT"]
    for r in range(5, wsrc.max_row + 1):
        if str(wsrc.cell(r, 1).value or "").startswith("SRC_ANT_"):
            sv[wsrc.cell(r, 1).value] = wsrc.cell(r, 5).value if wsrc.cell(r, 5).value is not None else (wsrc.cell(r, 6).value or wsrc.cell(r, 7).value)
    g0, g1 = R["v017"], R["base"]
    summ = (f"已到位 {g0['FND_Committed'][1]:.2f}→{g1['FND_Committed'][1]:.2f}；累計外部資金 {g0['FND_ExtNeedCum'][5]:.2f}→{g1['FND_ExtNeedCum'][5]:.2f}；"
            f"首次缺口 {g0['FND_FirstGapYear']}→{g1['FND_FirstGapYear']}；現金谷底 {g0['FND_CashNoExtTrough']:.2f}→{g1['FND_CashNoExtTrough']:.2f}；命題 1 不變")
    newv = {"SRC_ANT_480＋481": (sv.get("SRC_ANT_480") or 0) + (sv.get("SRC_ANT_481") or 0),
            "SRC_ANT_475 × 477": (sv.get("SRC_ANT_475") or 0) * (sv.get("SRC_ANT_477") or 0) / 1000}
    flip = {"SRC_ANT_328": "是：命題 2 累計外部資金需求 1.19 → 0（首次缺口 2029 → 無）；單獨只動本項（各輪仍 30／65）已足以翻轉",
            "SRC_ANT_302": "否（單獨 +0.2；與 298 合計 +0.4：累計外部資金 1.19 → 0.79）", "SRC_ANT_298": "否（同上）",
            "SRC_ANT_329": "否（只動情境 S2／S3；基準不變）"}
    for o_, n_, aff in REPL:
        o, nv = sv.get(o_), newv.get(n_, sv.get(n_))
        diff = (nv - o) / o if isinstance(o, (int, float)) and isinstance(nv, (int, float)) and o else None
        ws.append([o_, o, n_, nv, diff, aff, summ, flip.get(o_, "否（同值或只列示）")])

    # ⑦ 範圍外變動核對
    ws = sheet("⑦範圍外變動核對", ["比對", "比對格數", "其中數值格", "B 新增格", "不同格數", "不同格（頁!格：A → B）"])
    for p in a.diff_json:
        d = json.loads(Path(p).read_text(encoding="utf-8"))
        ws.append([Path(p).stem, d["compared"], d["numeric_compared"], d.get("added"), len(d["diff"]), "；".join(f"{x[0]}：{x[1]} → {x[2]}" for x in d["diff"][:40])])

    del wb["Sheet"]
    a.out.parent.mkdir(parents=True, exist_ok=True)
    wb.save(a.out)
    a.json.parent.mkdir(parents=True, exist_ok=True)
    json.dump(R, open(a.json, "w", encoding="utf-8"), ensure_ascii=False, indent=1, default=str)
    print("saved", a.out, a.json)


if __name__ == "__main__":
    main()
