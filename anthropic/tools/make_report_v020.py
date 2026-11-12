#!/usr/bin/env python3
"""v0.2.0 對照 Excel 與報告數字（CLAUDE.md 第 5 節；工作單 20261111 v0.2.0 初步招股說明書＋估值）。

①本版新增或變動的列（SRC_ANT_522–592、Inputs INP_234–276、Cost 第二十七節、Funding 第十六節、Val 頁具名範圍）
②關鍵輸出 FY2025–2030：v0.1.9（model/archive 的 v0.1.9 Excel 快取值）vs v0.2.0
③拆解：只動該項（B1、B2、C、D1、E 各自從 v0.1.9 加入）與從 v0.2.0 只關該項
④Val 摘要（具名範圍、Excel 位置）與 Sens 表
⑤已套用的預設
情境以 engine（pycel）重算現行活頁簿（覆寫值取自 tools/html_scenarios.yaml，與 HTML 同一套）；本程式不做模型計算。
用法：python3 tools/make_report_v020.py --out docs/reports/20261111_v0.2.0_對照.xlsx [--json out.json]
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
from build_html import Locator, resolve_overrides  # noqa: E402
from engine import Engine, current_model_path, resolve_key  # noqa: E402

V019 = REPO / "model" / "archive" / "20261104_Anthropic_v0.1.9.xlsx"
YEARS = [2025, 2026, 2027, 2028, 2029, 2030]
KEY = [("REV_GrossCapped", "營收總額（截頂後；$B）"), ("REV_NetCapped", "營收淨額（$B）"), ("COST_ContractPay", "合約付款（$B）"), ("COST_Compute", "算力成本（$B）"),
       ("COST_NonCompExSBC", "非算力成本（不含股權報酬）"), ("COST_SBC", "股權報酬（$B）"), ("FND_FCF", "營運自由現金流 F07（$B）"), ("FND_EquityIn", "計入基準的股權流入（$B）"),
       ("FND_CashNoExt", "只靠已到位融資的年底現金（$B）"), ("FND_MinCash", "最低現金（$B）"), ("FND_ExtNeedCum", "累計外部資金需求（$B）"),
       ("COST_PropRev_VR", "每 VR 等值 GW 營收淨額"), ("COST_PropFull_VR", "每 VR 等值 GW 全成本"), ("COST_PropGap_VR", "命題 1 差額（$B/GW/年）"), ("COST_Coverage", "覆蓋率")]
SCAL = [("COST_GapPosSustainYear", "命題 1 持續轉正年"), ("FND_CashNoExtTrough", "現金谷底（$B）"), ("FND_FirstGapYear", "首次缺口年"),
        ("FND_FCFPosSustainYear", "營運自由現金流持續轉正年"), ("COST_S1Unalloc", "未歸屬餘額（$B）"), ("COST_S1fCommit3", "承諾三段合計（$B）"),
        ("CHK_Errors", "CHK_Errors"), ("CHK_Warnings", "CHK_Warnings")]
DECOMP = [("v0.1.9（全關）", "v019"), ("只動 B1 IPO＋私募", "v020_only_b1"), ("只動 B2 條件式", "v020_only_b2"), ("只動 C1–C3 承諾", "v020_only_c"),
          ("只動 D1 Q4 營收", "v020_only_d1"), ("只動 E1＋E2 股權報酬", "v020_only_e"), ("v0.2.0（全部）", None), ("v0.2.0 只關 B1", "v020_no_b1"),
          ("v0.2.0 只關 D1", "v020_no_d1"), ("v0.2.0 只關 E", "v020_no_e"), ("C4 扣除 Lambda 34.9", "lam_on"), ("超額配售計入", "oa_on")]
DEC_OUT = [("REV_GrossCapped", 2026, "營收總額 2026"), ("REV_GrossCapped", 2027, "營收總額 2027"), ("REV_GrossCapped", 2030, "營收總額 2030")] + \
          [("FND_FCF", y, f"F07 {y}") for y in (2026, 2027, 2028, 2029, 2030)] + \
          [("FND_ExtNeedCum", 2030, "累計外部資金需求"), ("FND_CashNoExtTrough", None, "現金谷底"), ("FND_FCFPosSustainYear", None, "F07 持續轉正年"),
           ("COST_PropGap_VR", 2026, "差額 2026"), ("COST_PropGap_VR", 2027, "差額 2027"), ("COST_PropGap_VR", 2028, "差額 2028"), ("COST_PropGap_VR", 2029, "差額 2029"),
           ("COST_PropGap_VR", 2030, "差額 2030"), ("COST_Coverage", 2030, "覆蓋率 2030"), ("COST_GapPosSustainYear", None, "命題 1 持續轉正年"),
           ("COST_SBC", 2026, "股權報酬 2026"), ("COST_SBC", 2027, "股權報酬 2027"), ("COST_SBC", 2028, "股權報酬 2028"), ("COST_S1Unalloc", None, "未歸屬餘額"),
           ("VAL_TP_W", None, "加權目標價"), ("VAL_Rating", None, "評等")]
VALN = ["VAL_Px", "VAL_PxLo", "VAL_PxHi", "VAL_Rf", "VAL_Beta", "VAL_ERP", "VAL_Ke", "VAL_KdAT", "VAL_WACC", "VAL_g", "VAL_RONIC", "VAL_PVExp_B", "VAL_PVTV_B", "VAL_EV_B",
        "VAL_TVShare_B", "VAL_NetCash", "VAL_Shares", "VAL_EqV_B", "VAL_TP_DCF_A", "VAL_TP_DCF_B", "VAL_TP_DCF_C", "VAL_TP_DCF", "VAL_NTMRev", "VAL_PeerMult", "VAL_PeerFCFm",
        "VAL_AdjFactor", "VAL_TP_CompMain", "VAL_TP_CompAdj", "VAL_TP_W", "VAL_Up_W", "VAL_Up_CompMain", "VAL_Up_DCF", "VAL_Rating", "VAL_RatingNote", "VAL_Method", "VAL_SensBase"]
DEFAULTS = [
    ("J1 IPO＋私募進基準", "進（中點 74）：2026 股權流入加 淨募資 84.6 − RSU 代扣稅款 3.8＝80.8（p_ipo_base_switch＝1）", "不進（v0.1.9）", "現金谷底 40.2 → 142.2；命題 2 不變（0）"),
    ("J2 2026Q4 營收", "10-31 年化 81 ÷ 4＝20.25（p_q4_rr_switch＝1；中性近似）", "Q4＝Q3 16.956", "2026 營收總額 50.20 → 53.49；2027 起按校準 token 量同比例 +6.6%；F07 2029 轉正"),
    ("J3 RSU 一次性股權報酬 14.2", "計入 2026 全成本（2026Q4），不進每人基準、不外推", "不計", "2026 差額轉更負；非現金，F07 不變；Val EBIT 2026Q4 −14.2"),
    ("J4 Lambda 34.9 可終止", "照計（p_lambda_deduct_switch＝0）", "扣除（並列讀法 lam_on）", "見對照 Excel ③"),
    ("J5 估值方法", "DCF（FCFF）＋Comparable 50/50；收支模型加 Val 頁（M 線規格例外因掛牌失效）", "只做反向 DCF、不給評等", "—"),
    ("J6 Comparable", "高成長軟體 EV／NTM 營收中位數 26.32，× MIN(1, FY2030 自由現金流率 ÷ 同業 FCF 利益率中位數 29.8%)", "不調整（Sens）；EV／NTM 毛利（資料缺口，未算）", "不調整：加權目標價 103.91（買進）"),
    ("J7 情境", "A＝API 任務數年增率高端（2027 依 Q4 年化校準：基準 D2 比值 ≥ 1，下限不生效）；B＝基準；C＝低端＋Q4＝Q3；25／50／25", "只用基準", "DCF A／C 見 ④"),
    ("J8 現價", "IPO 區間中點 74（上市後改收盤價）", "區間上下緣（Sens）", "見 Sens"),
    ("F2 2026Q4 EBIT 的前九月代理（建置代理判斷）", "GAAP 營業損益 −1.260（與 Funding F06 同口徑：F06 已扣股權報酬與非現金捐贈）；val_q4_proxy_switch＝1",
     "工作單字面：營業損益＋股權報酬 3.3＋非現金捐贈＝調整後 4.081（會把股權報酬與捐贈重複扣在 2026Q4）", "字面讀法：2026Q4 EBIT 再減 5.34，加權目標價 −0.1（Sens 表最後一列）"),
    ("F5 淨現金對價格的連動", "（每 $1 淨募資 0.793 − 代扣稅款 0.052）＝0.741（pro forma 已扣代扣稅款，兩個敏感度同時套用）", "工作單字面 ±0.793／$", "74 時相同；70／78 時差 ±0.21"),
    ("F4 債務成本", "rf ＋ 利差 2%（Assumed 1–4%；無揭露）；債務只有融資租賃 1.95 → 對 WACC 影響 < 0.01pt", "kd＝ke", "可忽略"),
    ("F3 延伸期資本支出", "自建資本支出依 FY2030 占營收比例延伸（基準 0）", "—", "0"),
    ("F6 Comparable 時點", "NTM＝2026Q4 實值＋0.75 × FY2027（J9），每股值不再以 ke 推到目標時點", "以目標時點 NTM 或 × (1＋ke)", "不推進較保守"),
    ("中位數函數", "LARGE(範圍,(COUNT＋1)/2)（同業 5 家，奇數時等於中位數；pycel 不支援 MEDIAN）", "MEDIAN", "數值相同"),
    ("D 情境 A 的 2027 校準", "D2 比值（2027 營收總額 ÷ 2026Q4 年化）基準 1.099、A 1.412 皆 ≥ 1 → 不另加校準機制", "強制 2027＝Q4 年化（會使 A 低於基準，與「高端」矛盾）", "無"),
    ("DCF／Comparable 下限（獨立核對）", "淨現金每股（評價日；DCF 再以 ke 推到目標時點）", "0", "情境 C DCF 0 → 14.28"),
    ("估值 EBIT 的股權報酬（獨立核對）", "只扣未來新授：加回 2026Q4 一次性 14.2 與既有獎酬攤銷（以當年模型股權報酬為上限；這批 RSU 已計入評價股數）", "照扣全部（Sens）", "加權目標價 +0.74"),
    ("J9 NTM 營收（獨立核對）", "2026Q4 實值＋0.75 × FY2027", "0.25 × FY2026＋0.75 × FY2027（Sens）", "舊讀法 79.13"),
    ("融資租賃負債（獨立核對）", "現值 1.674（F-35）", "工作單 1.95", "淨現金 +0.28"),
    ("C5 2026Q4 合約付款", "全年 ÷ 季數（年內平均近似）", "—", "只對照"),
]


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--out", type=Path, required=True)
    ap.add_argument("--json", type=Path)
    a = ap.parse_args()
    model = current_model_path()
    eng = Engine(model)
    loc = Locator(model)

    def at(sc, name, year=None):
        if sc is None:
            v = eng.get_name(name)
        else:
            ov = resolve_overrides(sc, loc)
            orig = {k: eng.get(*resolve_key(loc.names, k)) for k in ov}
            try:
                for k, x in ov.items():
                    eng.set_key(k, x)
                v = eng.get_name(name)
            finally:
                for k, x in orig.items():
                    eng.set_key(k, x)
        if isinstance(v, list) and year is not None:
            return v[YEARS.index(year)]
        return v

    wb9 = openpyxl.load_workbook(V019, data_only=True)

    def old(name):
        d = wb9.defined_names[name]
        sh, ref = d.attr_text.split("!")
        c = wb9[sh][ref.replace("$", "")]
        return [x.value for row in c for x in row] if isinstance(c, tuple) else c.value

    out = openpyxl.Workbook()
    B = Font(bold=True)
    res = {}
    # ① 新列
    ws = out.active
    ws.title = "①新列"
    ws.append(["頁", "ID", "key／列名", "值", "低", "高", "單位", "標記", "來源／依據"])
    for c in ws[1]:
        c.font = B
    src = yaml.safe_load((REPO / "data" / "anthropic_src.yaml").read_text(encoding="utf-8"))["rows"]
    for r in src[521:]:
        ws.append(["SRC_ANT", r["id"], f"{r['key']}｜{r['metric']}", r["value"], r["lo"], r["hi"], r["unit"], r["tag"], f"{r['source']}；{r['url']}"])
    inp = yaml.safe_load((REPO / "data" / "anthropic_inputs.yaml").read_text(encoding="utf-8"))["rows"]
    reg = json.loads((REPO / "builder" / "id_registry.json").read_text(encoding="utf-8"))
    for r in inp:
        iid = reg["INP"][r["key"]]
        if int(iid[4:]) >= 234:
            ws.append(["Inputs", iid, f"{r['key']}｜{r['name']}", r["value"], r["lo"], r["hi"], r["unit"], r["tag"], r["basis"]])
    mwb = openpyxl.load_workbook(model)
    for sheet, first in (("Cost", "二十七、"), ("Funding", "十六、")):
        s = mwb[sheet]
        on = False
        for row in range(1, s.max_row + 1):
            a_ = s.cell(row, 1).value
            if isinstance(a_, str) and a_.startswith(first):
                on = True
                continue
            if on and s.cell(row, 2).value:
                ws.append([sheet, a_, s.cell(row, 2).value, None, None, None, s.cell(row, 3).value, "公式", s.cell(row, 11).value])
    for nm_ in VALN:
        ws.append(["Val", nm_, "具名範圍", at(None, nm_), None, None, None, "公式", mwb.defined_names[nm_].attr_text])
    # ② 關鍵輸出
    ws = out.create_sheet("②關鍵輸出")
    ws.append(["輸出", "版本"] + YEARS)
    for c in ws[1]:
        c.font = B
    res["key"] = {}
    for name, lab in KEY:
        o, nw = old(name), at(None, name)
        ws.append([lab, "v0.1.9"] + list(o))
        ws.append([name, "v0.2.0"] + list(nw))
        res["key"][name] = dict(old=o, new=nw)
    ws.append([])
    ws.append(["純量", "v0.1.9", "v0.2.0"])
    for name, lab in SCAL:
        o, nw = old(name), at(None, name)
        ws.append([f"{lab}（{name}）", o, nw])
        res["key"][name] = dict(old=o, new=nw)
    # ③ 拆解
    ws = out.create_sheet("③拆解")
    ws.append(["輸出"] + [l for l, _ in DECOMP])
    for c in ws[1]:
        c.font = B
    res["decomp"] = {}
    for name, y, lab in DEC_OUT:
        row = [at(sc, name, y) for _l, sc in DECOMP]
        ws.append([lab] + row)
        res["decomp"][lab] = row
    # ④ Val
    ws = out.create_sheet("④Val與Sens")
    ws.append(["具名範圍", "值", "Excel 位置"])
    for c in ws[1]:
        c.font = B
    res["val"] = {}
    for nm_ in VALN:
        v = at(None, nm_)
        ws.append([nm_, v, mwb.defined_names[nm_].attr_text])
        res["val"][nm_] = v
    for nm_ in ("VAL_Rev_B", "VAL_EBIT_B", "VAL_Tax_B", "VAL_FCFF_B", "VAL_DF_B", "VAL_Yr"):
        ws.append([nm_] + list(at(None, nm_)) + [mwb.defined_names[nm_].attr_text])
        res["val"][nm_] = at(None, nm_)
    ws.append([])
    sens = yaml.safe_load((REPO / "data" / "val_sens.yaml").read_text(encoding="utf-8"))
    ws.append(["Sens 案例", "低值情況加權目標價", "高值情況加權目標價", "低值設定", "高值設定", "評等（低）", "評等（高）"])
    for cs in sens["cases"]:
        ws.append([cs["name"], cs["tp_lo"], cs["tp_hi"], cs["lo_label"], cs["hi_label"], cs["rating_lo"], cs["rating_hi"]])
    ws.append(["基準", sens["base_tp"]])
    for g in sens.get("gap", []):
        ws.append(["資料缺口", g])
    res["sens"] = sens
    sn = yaml.safe_load((REPO / "data" / "val_scenarios.yaml").read_text(encoding="utf-8"))
    ws.append([])
    ws.append(["情境快照", "說明", "D2 比值", "營收淨額 FY2026–30", "EBIT FY2026–30", "自由現金流率 FY2026–30"])
    for s, d in sn["scenarios"].items():
        ws.append([s, d["desc"], d["d2_ratio"], json.dumps(d["values"]["rev"]), json.dumps(d["values"]["ebit"]), json.dumps(d["values"]["fcfm"])])
    # ⑤ 預設
    ws = out.create_sheet("⑤已套用的預設")
    ws.append(["問題", "採用的預設", "替代", "影響"])
    for c in ws[1]:
        c.font = B
    for r in DEFAULTS:
        ws.append(list(r))
    a.out.parent.mkdir(parents=True, exist_ok=True)
    out.save(a.out)
    if a.json:
        a.json.write_text(json.dumps(res, ensure_ascii=False, indent=1, default=str), encoding="utf-8")
    print("saved", a.out)


if __name__ == "__main__":
    main()
