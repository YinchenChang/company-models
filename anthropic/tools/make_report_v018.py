#!/usr/bin/env python3
"""v0.1.8 對照 Excel 與報告數字（CLAUDE.md 第 5 節、第 8 節；工作單 20261102 v0.1.8 R3–R7）。

①新增或變動的列（SRC_ANT_488–521 新列、被取代的舊列、Inputs INP_221–231、改式列、Cost 第二十六節、Funding 第十五節、Checks 新列）
②關鍵輸出 FY2025–2030：v0.1.7.1 vs v0.1.8（另列 R4、R5d 並列讀法）
③瀑布：逐項累加（v0.1.7.1 → ＋R5a → ＋R5b → ＋R5c）與「只動該項」（R5a、R5b、R5c 各自從 v0.1.7.1；R4、R5d 從 v0.1.8 基準）
④最低現金 3／6／12（v0.1.7.1、v0.1.8、v0.1.8＋R4）
⑤已套用的預設（含替代讀法量化）
⑥S-1/A 替換對照（CLAUDE.md 第 8 節第 4 步）
⑦範圍外變動核對（tools/diff_xlsx.py 的 JSON）
⑧R6 下游：Funding F07（自由現金流）／F03（自建資本支出）新舊值（Wattonomics import_gap_models 取用列）
情境一律以 engine（pycel）重算現行活頁簿（覆寫值取自 tools/html_scenarios.yaml，與 HTML 同一套）；本程式不做模型計算。
用法：python3 tools/make_report_v018.py --out docs/reports/20261103_v0.1.8_對照.xlsx [--diff-json build_out/diff_*.json ...]
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
KEY = [("REV_Gross", "營收總額（$B）"), ("REV_NetCapped", "營收淨額（$B）"), ("COST_Compute", "算力成本（$B）"), ("COST_Pay_S1Unalloc", "未歸屬餘額付款（$B）"),
       ("COST_Pay_Spacex", "SpaceX 付款（$B）"), ("COST_NonCompExSBC", "非算力成本（不含股權報酬，$B）"), ("COST_SBC", "股權報酬（$B）"),
       ("CMP_Supply_VReq", "供給 VR 等值 GW"), ("COST_PropGap_VR", "每 VR 等值 GW 差額（$B/GW/年）"), ("COST_Coverage", "覆蓋率"),
       ("FND_FCF", "營運自由現金流（$B）"), ("FND_Committed", "已到位融資（$B）"), ("FND_CashNoExt", "只靠已到位融資的年底現金（$B）"), ("FND_MinCash", "最低現金（$B）"),
       ("FND_ExtNeed", "當年外部資金需求（$B）"), ("FND_ExtNeedCum", "累計外部資金需求（$B）"), ("FND_Headroom", "現金餘裕（$B）"),
       ("FND_ExtNeedCumCond", "S2 條件式：累計外部資金需求（$B）"), ("FND_ExtNeedCumAll", "S3 IPO＋條件式：累計外部資金需求（$B）")]
SCAL = [("FND_CashNoExtTrough", "現金谷底（$B）"), ("FND_FirstGapYear", "首次缺口年"), ("COST_GapPosSustainYear", "命題 1 持續轉正年"), ("FND_FCFPosSustainYear", "營運自由現金流持續轉正年"),
        ("FND_ExtNeedPeak", "單年外部資金需求峰值（$B）"), ("COST_S1fCommit3", "S-1 承諾三段合計（$B）"), ("COST_S1Unalloc", "未歸屬餘額（$B）"),
        ("COST_S1fCommitModel", "模型逐家合約總額（不含 AMD，$B）"), ("FND_CommittedS1Gap", "已到位融資 − S-1 特別股收入（$B）"), ("FND_CommittedS1aGap", "已到位融資 − S-1/A 前九月特別股收入（$B）"),
        ("FND_Cash0930Gap", "S-1/A 9-30 現金 − 模型近似（$B）"), ("FND_Cash0930Resid", "9-30 差額拆解其餘項（$B）"), ("FND_L1_ExtTotal", "第 1 層累計外部資金需求（$B）"),
        ("CHK_Errors", "CHK_Errors"), ("CHK_Warnings", "CHK_Warnings")]
READS = [("v0.1.7.1", "v0171"), ("v0.1.8", None), ("v0.1.8＋R4 SpaceXAI S-1 路徑", "r4_on"), ("v0.1.8＋R5d 扣除 213.8", "r5d_on")]
STEPS = [("v0.1.7.1", "v0171"), ("＋R5a Q3 實績", "wf8_a"), ("＋R5b 承諾 871.4", "wf8_ab"), ("＋R5c 付款期＝v0.1.8", None)]
ONLY = [("v0.1.7.1", "v0171"), ("只動 R5a", "wf8_a"), ("只動 R5b", "wf8_b"), ("只動 R5c", "wf8_c"), ("v0.1.8（三項）", None),
        ("v0.1.8＋R4", "r4_on"), ("v0.1.8＋R5d", "r5d_on"), ("v0.1.8＋R4＋R5d", "r4_r5d"),
        ("v0.1.8 只關 R5a", "v018_no_r5a"), ("v0.1.8 只關 R5b", "v018_no_r5b"), ("v0.1.8 只關 R5c", "wf8_ab")]
MC = [("3 個月", "v0171_mc_lo", "mc_lo", "r4_mc_lo"), ("6 個月（基準）", "v0171", None, "r4_on"), ("12 個月", "v0171_mc_hi", "mc_hi", "r4_mc_hi")]
ALTS = [("R5c 舊讀法 2027–2033 平均", "wf8_ab"), ("舊讀法 快 2027–2031", "s1u_fast"), ("舊讀法 慢 2028–2036", "s1u_slow"), ("舊讀法 慢 2028–2037", "s1u_slow37"),
        ("S-1/A 結構，未起租起租年 2026", "lease_lo"), ("S-1/A 結構，未起租起租年 2028", "lease_hi"), ("R4 SpaceXAI S-1 路徑", "r4_on"), ("R5d 扣除附條件 213.8", "r5d_on"),
        ("R4＋R5d", "r4_r5d"), ("並列：殘差法（S-1/A 結構 − 逐家合約時程）", "resid_on"), ("殘差法＋R4", "r4_resid"), ("捐贈只計 Q3 一次", "chr_once"), ("SpaceXAI 以上限 120.5 等比例放大月費（v0.1.7 W1(a)）", "spx_cap"), ("Google 額外 TPU 計入（中點）", "gaddl_on"),
        ("Broadcom 起始年 2026", "bcm_2026"), ("AMD＝S-1 下限 20", "amd_s1"), ("只關 W2（Fluidstack 舊法）", "only_w2_off"), ("只關 W4", "only_w4_off"),
        ("W4 只動 Q3 口徑", "w4_q3only"), ("W4 只動股權報酬", "w4_sbconly"), ("W4 兩項皆不用", "w4_neither"), ("承諾不進基準（W1、W2 皆關）", "commit_off"),
        ("API 任務數成長取低", "api_task_lo"), ("價格彈性取弱端", "elast_weak"), ("每 GW 合約價 20", "price_hi"), ("任務數低＋彈性弱", "combo1")]
DEFAULTS = [
    ("N1 SpaceXAI（R4）", "基準維持既有月費 1.25 × 36＝45（s1f_spacex_fee_switch＝0）；S-1 路徑並列", "S-1 路徑：2026-05～11 線性爬升、12 月起每月 3.1、上限 120.5 至 2029", "r4_on"),
    ("R4 爬升形狀", "第 k 月＝滿載 × k ÷（爬升月數＋1），2026 年合計＝3.1 × 7 ÷ 2 ＋ 3.1（Assumed，線性；12 月為首個滿載月）", "k ÷ 爬升月數（11 月即滿載）：2026 多 3.1、上限更早觸及；總額同為 120.5", None),
    ("R5a Q3 實績", "營收、W4 下半年非算力、股權報酬代理皆改 S-1/A Q3 實際（Q4＝Q3）", "回到 v0.1.7.1（報導 16.8、Q3 隱含、Q2 代理）", "v018_no_r5a"),
    ("R5a 非現金捐贈 1.960 的 Q4 處理（判斷題 s1a_chr_q4_switch）", "2026 全成本照 Q4＝Q3 慣例 Q4 也計一次（保守）；捐贈不進每人股權報酬基準（核對後修正）", "只計 Q3 一次（2026 股權報酬 −1.96；現金不變）", "chr_once"),
    ("R5c 殘差法（並列；s1a_resid_switch）", "不採用（上限型合約 SpaceXAI、Nscale 的模型付款可能超過最低承諾，殘差會低估未歸屬付款；基準取保守）", "未歸屬餘額逐年＝S-1/A 三段逐年 − 模型逐家逐年", "resid_on"),
    ("R5b 三段口徑", "表內 320.0（原文 319.993）＋未起租 451.2＋9-30 後 100.2＝871.4；短期 17.4、期後短期 6.6 照 N2 不進基準", "回到 S-1 6-30 三段 832", "v018_no_r5b"),
    ("R5b Google 合約", "S-1/A ≥121.3（取代 114.5）", "114.5", None),
    ("R5c 未歸屬餘額付款期", "依 S-1/A 結構合成：表內逐年（之後 148.4 依 2031–2036 平均）＋未起租與 9-30 後依起租 2027、10 年平均，以三段金額加權", "舊 2027–2033 平均", "wf8_ab"),
    ("R5c 起租年（Assumed 2026–2028）", "2027", "2026／2028", "lease_lo"),
    ("R5c 2026Q4 表內付款", "納入（S-1/A 表 2026 欄 5.231 → 比例 0.6%，未歸屬餘額 2026 付 1.75）", "2026 不付（起於 2027）", None),
    ("R5d 附條件 213.8", "基準計入（S-1/A 列入總額）", "自三段合計扣除（s1a_cond_deduct_switch＝1）", "r5d_on"),
    ("R5e 9-30 現金對照段落位置", "新增 Funding 第十五節（第十三節之後已有第十四節；新列一律附加在頁尾）", "接在第十三節（會使第十四節列號移動）", None),
    ("R5e 受限現金 2.6", "只對照（不自現金扣除）", "自現金扣除", None),
    ("R5f 其他", "Q3 無新股權（已到位 110.40 不變）；Google 額外 TPU 125–150、AMD >20 不進基準；TK v21.7", "—", None),
    ("最低現金（R3）", "維持 6 個月（Andy 未決）；3／12 個月並列", "3／12 個月", "mc_hi"),
    ("N2 期後短期合約、N3 有條件融資額度、既有判斷題", "照舊（工作單預設表）", "—", None),
]
REPL = [("SRC_ANT_384", "SRC_ANT_488", "Revenue 營收總額基準 2026 的 Q3（s1a_q3_switch）"),
        ("SRC_ANT_385", "SRC_ANT_494", "調整後營業利益率（W4 下半年改 Q3 實際後，11% 只留在第十九節對照列）"),
        ("SRC_ANT_445", "SRC_ANT_504", "W1(b) 三段：未起租租賃（s1a_commit_switch）"),
        ("SRC_ANT_446", "SRC_ANT_496", "W1(b) 三段：6-30 採購承諾 → 9-30 表內義務"),
        ("SRC_ANT_447", "SRC_ANT_506", "W1(b) 三段：6-30 後新簽 → 9-30 後新簽"),
        ("SRC_ANT_450", "SRC_ANT_513", "Cost 第十一節 Google 合約總額與付款（s1a_commit_switch）")]
F_ROWS = ("F01", "F02", "F03", "F04", "F07")


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
        for rid in F_ROWS:
            r = loc.rowid[("Funding", rid)]
            out["Funding!" + rid] = [eng.get("Funding", f"{c}{r}") for c in "DEFGHI"]
        return out
    finally:
        for k, v in orig.items():
            eng.set_key(k, v)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--out", required=True, type=Path)
    ap.add_argument("--json", type=Path, default=REPO / "build_out" / "report_v018.json")
    ap.add_argument("--diff-json", type=Path, nargs="*", default=[])
    a = ap.parse_args()
    model = current_model_path()
    loc, cfg, eng = Locator(model), load_scenarios(), Engine(model)
    scs = ({s for _l, s in READS + STEPS + ONLY + ALTS} | {x for _l, p, q, r in MC for x in (p, q, r)} | {d[3] for d in DEFAULTS if d[3]})
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
    old = openpyxl.load_workbook(REPO / "model" / "archive" / "20261102_Anthropic_v0.1.7.1.xlsx", data_only=True)
    # ① 新增與變動列
    ws = sheet("①新增與變動列", ["頁", "列 ID", "列名／項目", "值（D 或 E）", "低／F", "高／G", "v0.1.7.1 值", "標記", "來源／說明"])
    srcy = {r["id"]: r for r in yaml.safe_load((REPO / "data" / "anthropic_src.yaml").read_text(encoding="utf-8"))["rows"]}
    for rid, r in srcy.items():
        n = int(rid[8:])
        if n >= 488:
            ws.append(["SRC_ANT", rid, r["metric"], r["value"], r["lo"], r["hi"], None, r["tag"], "新增（S-1/A No.1）；" + str(r["note"])[:120]])
        elif n in (384, 385, 445, 446, 447, 450, 452):
            ws.append(["SRC_ANT", rid, r["metric"], r["value"], r["lo"], r["hi"], r["value"], r["tag"], "備註加註取代（值不改）"])
    wi = mv["Inputs"]
    for r in range(5, wi.max_row + 1):
        iid = wi.cell(r, 1).value
        if isinstance(iid, str) and iid.startswith("INP_") and int(iid[4:]) >= 221:
            ws.append(["Inputs", iid, wi.cell(r, 2).value, wi.cell(r, 5).value, wi.cell(r, 6).value, wi.cell(r, 7).value, None, wi.cell(r, 8).value, wi.cell(r, 9).value])
    for sh, pref in (("Revenue", "營收總額基準"), ("Cost", "合約實付：SpaceX"), ("Cost", "合約總額（模型）：Google Cloud"), ("Cost", "合約總額（模型）：SpaceX"),
                     ("Cost", "合約實付：Google Cloud"), ("Cost", "W1(b) S-1 三段合計"), ("Cost", "W1(b) 差額"), ("Cost", "W1(b) S-1 承諾未歸屬餘額"),
                     ("Cost", "合約實付：S-1 承諾未歸屬餘額"), ("Cost", "W4 2026 下半年非算力"), ("Cost", "股權報酬代理：2026 全年")):
        w, wo = mv[sh], old[sh]
        for r in range(1, w.max_row + 1):
            if str(w.cell(r, 2).value or "").startswith(pref):
                ws.append([sh, w.cell(r, 1).value, "改式（開關）：" + str(w.cell(r, 2).value)[:80], w.cell(r, 4).value, w.cell(r, 5).value, w.cell(r, 6).value,
                           wo.cell(r, 4).value if w.cell(r, 4).value is not None else wo.cell(r, 5).value, None, str(w.cell(r, 11).value)[:120]])
                break
    for sh, start in (("Cost", "二十六、"), ("Funding", "十五、")):
        w = mv[sh]
        on = False
        for r in range(1, w.max_row + 1):
            if str(w.cell(r, 1).value or "").startswith(start):
                on = True
                continue
            if on and w.cell(r, 1).value:
                ws.append([sh, w.cell(r, 1).value, w.cell(r, 2).value] + [w.cell(r, c).value for c in (4, 5, 6)] + [None, None, str(w.cell(r, 11).value or "")[:120]])
    wc = mv["Checks"]
    reg = json.loads((REPO / "builder" / "id_registry.json").read_text(encoding="utf-8"))["CHK"]
    newc = {v for k, v in reg.items() if k.startswith("i1_")}
    for r in range(5, wc.max_row + 1):
        cid = wc.cell(r, 1).value
        if cid in newc:
            ws.append(["Checks", cid, wc.cell(r, 2).value, wc.cell(r, 3).value, wc.cell(r, 4).value, wc.cell(r, 5).value, None, None, wc.cell(r, 6).value])

    # ② 關鍵輸出
    ws = sheet("②關鍵輸出", ["輸出", "讀法"] + [str(y) for y in YEARS])
    for n, lab in KEY:
        for rl, sc in READS:
            ws.append([lab, rl] + R[sc or "base"][n])
    for n, lab in SCAL:
        for rl, sc in READS:
            ws.append([lab, rl, R[sc or "base"][n]])

    rows3 = [(f"每 VR 等值 GW 差額 {y}", "COST_PropGap_VR", YEARS.index(y)) for y in YEARS[1:]] + \
            [("覆蓋率 2026", "COST_Coverage", 1), ("覆蓋率 2030", "COST_Coverage", 5), ("命題 1 持續轉正年", "COST_GapPosSustainYear", None),
             ("累計外部資金需求 2026–2030", "FND_ExtNeedCum", 5), ("首次缺口年", "FND_FirstGapYear", None), ("單年峰值", "FND_ExtNeedPeak", None),
             ("現金谷底（只靠已到位融資）", "FND_CashNoExtTrough", None), ("2029 年底現金（只靠已到位）", "FND_CashNoExt", 4), ("2029 最低現金", "FND_MinCash", 4),
             ("營運自由現金流持續轉正年", "FND_FCFPosSustainYear", None), ("2026 營收總額", "REV_Gross", 1), ("2026 非算力（不含股權報酬）", "COST_NonCompExSBC", 1),
             ("2026 股權報酬", "COST_SBC", 1), ("S-1 承諾三段", "COST_S1fCommit3", None), ("未歸屬餘額", "COST_S1Unalloc", None),
             ("未歸屬餘額付款 2026", "COST_Pay_S1Unalloc", 1), ("未歸屬餘額付款 2028", "COST_Pay_S1Unalloc", 3), ("算力成本 2028", "COST_Compute", 3), ("CHK_Warnings", "CHK_Warnings", None)]
    for title, cols in (("③瀑布_逐項累加", STEPS), ("③b只動該項", ONLY)):
        ws = sheet(title, ["輸出"] + [l for l, _s in cols])
        for lab, n, i in rows3:
            ws.append([lab] + [(R[sc or "base"][n][i] if i is not None else R[sc or "base"][n]) for _l, sc in cols])

    # ④ 最低現金 3／6／12
    ws = sheet("④最低現金3_6_12", ["最低現金月數", "2029 最低現金（v0.1.8）", "v0.1.7.1 累計外部資金", "v0.1.7.1 首次缺口", "v0.1.7.1 谷底",
                                   "v0.1.8 累計外部資金", "v0.1.8 首次缺口", "v0.1.8 谷底", "v0.1.8＋R4 累計外部資金", "v0.1.8＋R4 首次缺口"])
    for lab, p, q_, r4 in MC:
        o, w = R[p], R[q_ or "base"]
        x = R[r4] if r4 else None
        ws.append([lab, w["FND_MinCash"][4], o["FND_ExtNeedCum"][5], o["FND_FirstGapYear"], o["FND_CashNoExtTrough"], w["FND_ExtNeedCum"][5], w["FND_FirstGapYear"],
                   w["FND_CashNoExtTrough"], x["FND_ExtNeedCum"][5] if x else None, x["FND_FirstGapYear"] if x else None])

    # ⑤ 已套用的預設
    ws = sheet("⑤已套用的預設", ["問題", "採用的預設", "替代", "替代：累計外部資金", "替代：現金谷底", "替代：首次缺口", "替代：2026 差額", "替代：2030 差額", "替代：持續轉正年"])
    for q2, d, alt, sc in DEFAULTS:
        row = [q2, d, alt]
        if sc:
            o = R[sc]
            row += [o["FND_ExtNeedCum"][5], o["FND_CashNoExtTrough"], o["FND_FirstGapYear"], o["COST_PropGap_VR"][1], o["COST_PropGap_VR"][5], o["COST_GapPosSustainYear"]]
        ws.append(row)
    ws = sheet("⑤b替代讀法", ["讀法", "累計外部資金", "首次缺口年", "現金谷底", "2026 差額", "2030 差額", "持續轉正年", "未歸屬餘額付款 2028"])
    for lab, sc in ALTS:
        w = R[sc]
        ws.append([lab, w["FND_ExtNeedCum"][5], w["FND_FirstGapYear"], w["FND_CashNoExtTrough"], w["COST_PropGap_VR"][1], w["COST_PropGap_VR"][5],
                   w["COST_GapPosSustainYear"], w["COST_Pay_S1Unalloc"][3]])

    # ⑥ S-1/A 替換對照
    ws = sheet("⑥S-1A替換對照", ["舊列", "舊值", "新列", "新值", "差異比例", "受影響（本版改指新列）", "命題輸出前（v0.1.7.1）→後（v0.1.8）", "結論是否翻轉"])
    sv = {}
    wsrc = mv["SRC_ANT"]
    for r in range(5, wsrc.max_row + 1):
        if str(wsrc.cell(r, 1).value or "").startswith("SRC_ANT_"):
            sv[wsrc.cell(r, 1).value] = wsrc.cell(r, 5).value
    g0, g1 = R["v0171"], R["base"]
    summ = (f"2026 差額 {g0['COST_PropGap_VR'][1]:+.2f}→{g1['COST_PropGap_VR'][1]:+.2f}；2030 差額 {g0['COST_PropGap_VR'][5]:+.2f}→{g1['COST_PropGap_VR'][5]:+.2f}；"
            f"累計外部資金 {g0['FND_ExtNeedCum'][5]:.2f}→{g1['FND_ExtNeedCum'][5]:.2f}；谷底 {g0['FND_CashNoExtTrough']:.2f}→{g1['FND_CashNoExtTrough']:.2f}")
    seg_old = sv["SRC_ANT_445"] + sv["SRC_ANT_446"] + sv["SRC_ANT_447"]
    seg_new = sv["SRC_ANT_496"] + sv["SRC_ANT_504"] + sv["SRC_ANT_506"]
    flip = {"SRC_ANT_384": "命題 1 的 2026 單年：與 R5a 其他兩項合計 +2.66 → +0.70（不翻轉）；三項合計後 −0.94（翻轉，見瀑布）",
            "SRC_ANT_446": f"三段 {seg_old:.1f} → {seg_new:.1f}（{seg_new / seg_old - 1:+.1%}）；R5b 單獨會使命題 2 翻為需要 {R['wf8_b']['FND_ExtNeedCum'][5]:.2f}（2029），與 R5c 合併後不翻轉"}
    for o_, n_, aff in REPL:
        o, nv = sv.get(o_), sv.get(n_)
        diff = (nv - o) / o if isinstance(o, (int, float)) and isinstance(nv, (int, float)) and o else None
        ws.append([o_, o, n_, nv, diff, aff, summ, flip.get(o_, "否（單獨不翻轉；見瀑布）")])

    # ⑦ 範圍外變動核對
    # ⑥b 核對後修正 A 的前後（每人股權報酬基準不含捐贈）
    pre = Path("/tmp/ant18/prefix_report.json")
    ws = sheet("⑥b核對後修正A", ["輸出", "修正前（同一活頁簿版本，捐贈帶進每人基準）", "修正後（現行）"])
    if pre.exists():
        P = json.loads(pre.read_text(encoding="utf-8"))
        for n_, lab in (("COST_SBC", "股權報酬"), ("COST_PropGap_VR", "每 VR 等值 GW 差額"), ("COST_Coverage", "覆蓋率")):
            for k, y in enumerate(YEARS):
                ws.append([f"{lab} {y}", P[n_][k], R["base"][n_][k]])
    ws = sheet("⑦範圍外變動核對", ["比對", "比對格數", "其中數值格", "B 新增格", "不同格數", "不同格（頁!格：A → B）"])
    for p in a.diff_json:
        d = json.loads(Path(p).read_text(encoding="utf-8"))
        ws.append([Path(p).stem, d["compared"], d["numeric_compared"], d.get("added"), len(d["diff"]), "；".join(f"{x[0]}：{x[1]} → {x[2]}" for x in d["diff"][:40])])

    # ⑧ R6 下游
    ws = sheet("⑧R6下游_Funding", ["列", "讀法"] + [str(y) for y in YEARS])
    for rid in F_ROWS:
        for rl, sc in (("v0.1.7.1", "v0171"), ("v0.1.8", None)):
            ws.append([f"Funding {rid}", rl] + R[sc or "base"]["Funding!" + rid])

    del wb["Sheet"]
    a.out.parent.mkdir(parents=True, exist_ok=True)
    wb.save(a.out)
    a.json.parent.mkdir(parents=True, exist_ok=True)
    json.dump(R, open(a.json, "w", encoding="utf-8"), ensure_ascii=False, indent=1, default=str)
    print("saved", a.out, a.json)


if __name__ == "__main__":
    main()
