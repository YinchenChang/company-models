#!/usr/bin/env python3
"""v0.1.9 對照 Excel 與報告數字（CLAUDE.md 第 5 節；工作單 20261104 v0.1.9 TK v21.10 同步）。

①TK 快照變動的名稱：新舊值與本模型使用位置
②關鍵輸出 FY2025–2030：v0.1.8（＝TK 16 個值回到 v21.7）vs v0.1.9
③拆解：逐項累加（v0.1.8 → ＋(a) TPU v7 產出比 → ＋(b) Trainium3 產出比 → ＋(c) 其他 DIFF＝v0.1.9）與「只動該項」
④敏感度（不進基準）：(i) TPU／Trainium3 產出比取 TK 低端（＝v21.9 原生精度）、(ii) 取 TK 高端、(iii) v21.7 舊值
⑤η 各版本小表（v0.1.1 v5.26 … v0.1.9 v21.10）
⑥逐格比較（tools/diff_xlsx.py 的 JSON）
⑦已套用的預設
情境一律以 engine（pycel）重算現行活頁簿（覆寫值取自 tools/html_scenarios.yaml，與 HTML 同一套）；本程式不做模型計算。
用法：python3 tools/make_report_v019.py --out docs/reports/20261104_v0.1.9_對照.xlsx [--tk-diff-json f.json] [--diff-json f.json ...]
"""
import argparse
import json
import sys
from pathlib import Path

import openpyxl
from openpyxl.styles import Font

REPO = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO))
sys.path.insert(0, str(REPO / "tools"))
from build_html import Locator, load_scenarios, resolve_overrides  # noqa: E402
from engine import Engine, current_model_path, resolve_key  # noqa: E402

YEARS = [2025, 2026, 2027, 2028, 2029, 2030]
KEY = [("CMP_Eta", "η（逐年）"), ("CMP_SupplyGW", "供給 GW（實體）"), ("CMP_Supply_VReq", "供給 VR 等值 GW"), ("CMP_InfGW_Eff", "有效推論 GW"), ("CMP_RDGW", "研發 GW"),
       ("COST_ContractPay", "合約付款（$B）"), ("COST_Compute", "算力成本（$B）"), ("REV_NetCapped", "營收淨額（$B）"),
       ("COST_PropRev_VR", "每 VR 等值 GW 營收淨額"), ("COST_PropGap_VR", "每 VR 等值 GW 差額（$B/GW/年）"), ("COST_Coverage", "覆蓋率"),
       ("FND_FCF", "營運自由現金流（$B）"), ("FND_CashNoExt", "只靠已到位融資的年底現金（$B）"), ("FND_MinCash", "最低現金（$B）"),
       ("FND_ExtNeedCum", "累計外部資金需求（$B）")]
SCAL = [("CMP_Eta2025", "η（2025 校準）"), ("COST_GapPosSustainYear", "命題 1 持續轉正年"), ("FND_CashNoExtTrough", "現金谷底（$B）"), ("FND_FirstGapYear", "首次缺口年"),
        ("FND_FCFPosSustainYear", "營運自由現金流持續轉正年"), ("ETA_Eta", "η（EtaDiag）"), ("ETA_fP", "f_P"), ("ETA_fU", "f_U"), ("ETA_fW", "f_W"),
        ("ETA_Resid", "未解釋殘差（不計生產折減）"), ("ETA_fU_Prod", "f_U'（計入生產折減）"), ("ETA_Resid2", "未解釋殘差（計入生產折減）"),
        ("CHK_Errors", "CHK_Errors"), ("CHK_Warnings", "CHK_Warnings")]
STEPS = [("v0.1.8（TK v21.7）", "v018"), ("＋(a) TPU v7 產出比", "wf9_a"), ("＋(b) Trainium3 產出比", "wf9_ab"), ("＋(c) 其他 DIFF＝v0.1.9", None)]
ONLY = [("v0.1.8", "v018"), ("只動 (a) TPU v7", "wf9_a"), ("只動 (b) Trainium3", "wf9_b"), ("只動 (c) 其他 DIFF", "wf9_c"), ("v0.1.9（三項）", None),
        ("v0.1.9 只關 (a)", "v019_no_tpu"), ("v0.1.9 只關 (b)", "v019_no_trn")]
SENS = [("v0.1.9 基準（FP8 對齊：TPU 0.66、Trn3 0.28）", None), ("(i) TK 低端＝v21.9 原生精度（0.35、0.14）", "sens9_lo"),
        ("(ii) TK 高端（0.84、0.50）", "sens9_hi"), ("(iii) v21.7 舊值（0.8、0.5；對照）", "wf9_c")]
ETA_VERS = [("v0.1.1（TK v5.26）", "v011"), ("v0.1.2（v18.0）", "v012"), ("v0.1.3（v19.0）", "v013"), ("v0.1.4（v19.1）", "v014"), ("v0.1.5–v0.1.6（v21.2）", "v016"),
            ("v0.1.7–v0.1.8（v21.7）", "v018"), ("v0.1.9（v21.10）", None)]
DEFAULTS = [
    ("敏感度 (i)(ii) 動哪些晶片", "只動 TPU v7、Trainium3 的產出比（本版 TK 變動者）；MI455X 維持基準", "沿用既有 nnv_lo／nnv_hi（三族含 MI455X 一起取低／高）", None),
    ("TPU v6e、Trainium2 的產出比", "沿用既有公式：Hopper × TPU（或 Trainium）產出比，v21.10 的 FP8 對齊比例也套到舊世代（公式不改）", "舊世代另設比例（Tokenomics 缺口：NonNV 表無 TPU v6e、Trainium2）", None),
    ("歷史版本在 HTML 的呈現", "①之七（v0.1.8）改以 v0.1.8 當時的輸出呈現（情境 v018、v018_*，TK 16 個值回到 v21.7）；①之六以前的歷史情境一併加上 tk18_back", "直接用現行基準", None),
    ("OpenAI 並排快照", "維持 OpenAI v0.6.4（其 TK 為 v21.7；並排時 TPU／Trainium 產出比口徑不同，只影響並排列）", "等 OpenAI 同步 v21.10 後再改", None),
    ("η 合理範圍 WARN（C54，0.5–2）", "門檻不改；v0.1.9 η 0.604 回到範圍內，C54 警示消失（CHK_Warnings 4 → 3）", "—", None),
]
TK_USE = {
    "TK_NNV_TPUv7_Out": "進計算：Compute C34（TPU v6e＝Hopper × 比例）、C38（TPU v7）→ 各族每 GW 產能（C110、C111、C120、C121）、VR 等值係數（C37、C41）→ η（2025 校準）、有效推論／研發 GW 分攤、供給 VR 等值 GW（命題 1 分母）；Checks C13（低 ≤ 基準 ≤ 高）",
    "TK_NNV_TPUv7_OutLo": "敏感度 C（Compute C200、C201、C205：產出比低端）；Checks C13",
    "TK_NNV_TPUv7_OutHi": "敏感度 C（Compute C207、C208、C212：產出比高端）；Checks C13",
    "TK_NNV_Trn3_Out": "進計算：Compute C42（Trainium2＝Hopper × 比例）、C46（Trainium3）→ 各族每 GW 產能（C112、C113、C122、C123）、VR 等值係數（C45、C49）→ η、推論／研發 GW、供給 VR 等值 GW；Checks C13",
    "TK_NNV_Trn3_OutLo": "敏感度 C（Compute C200、C201、C205）；Checks C13",
    "TK_NNV_Trn3_OutHi": "敏感度 C（Compute C207、C208、C212）；Checks C13",
    "TK_IF_HoldEcon": "只有高成本欄變動：Cost K05、K08 … K29（每 GW 年持有成本・高成本）→ K63 供應商持有成本（高成本）、K67 雲端毛利率（高成本）、EtaDiag E06 高欄（f_P 高）；基準欄不變，不進命題輸出",
    "TK_IF_FullCost_Sol": "Cost K112（TK 參考，並列）只讀基準欄；高成本欄變動不被引用",
    "TK_IF_FullCostDefault_Sol": "Cost K111（TK 參考，下游預設）只讀基準欄；高成本欄變動不被引用",
}
TK_ONLY_SNAP = "只快照（本模型無公式引用）"


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
        cash, mc = out["FND_CashNoExt"], out["FND_MinCash"]
        i = min(range(1, 6), key=lambda j: cash[j])
        out["TroughYear"], out["TroughMinCash"] = YEARS[i], mc[i]
        return out
    finally:
        for k, v in orig.items():
            eng.set_key(k, v)


def row5(o):
    return ([o["CMP_Eta2025"]] + o["COST_PropGap_VR"][1:] + [o["COST_Coverage"][5], o["COST_GapPosSustainYear"], o["FND_ExtNeedCum"][5], o["FND_FirstGapYear"],
            o["FND_CashNoExtTrough"], o["TroughYear"], o["TroughMinCash"], o["CMP_Supply_VReq"][5], o["CMP_InfGW_Eff"][5], o["CMP_RDGW"][5]])


H5 = ["η（2025）", "差額 2026", "差額 2027", "差額 2028", "差額 2029", "差額 2030", "覆蓋率 2030", "命題 1 持續轉正年", "累計外部資金 2026–2030", "首次缺口年",
      "現金谷底", "谷底年", "谷底年最低現金", "供給 VR 等值 GW 2030", "有效推論 GW 2030", "研發 GW 2030"]


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--out", required=True, type=Path)
    ap.add_argument("--json", type=Path, default=REPO / "build_out" / "report_v019.json")
    ap.add_argument("--tk-diff-json", type=Path)
    ap.add_argument("--diff-json", type=Path, nargs="*", default=[])
    a = ap.parse_args()
    model = current_model_path()
    loc, cfg, eng = Locator(model), load_scenarios(), Engine(model)
    scs = {s for _l, s in STEPS + ONLY + SENS + ETA_VERS}
    R = {(s or "base"): run(eng, loc, cfg, s) for s in scs}
    wb = openpyxl.Workbook()
    B = Font(bold=True)

    def sheet(title, header):
        ws = wb.create_sheet(title)
        ws.append(header)
        for c in ws[1]:
            c.font = B
        return ws

    # ① TK 名稱
    ws = sheet("①TK名稱新舊值與使用位置", ["名稱", "欄（世代／成本情境）", "v21.7（v0.1.8）", "v21.10（v0.1.9）", "變動率", "本模型使用位置"])
    old = openpyxl.load_workbook(REPO / "model" / "archive" / "20261103_Anthropic_v0.1.8.xlsx", data_only=True)["TK_Link"]
    new = openpyxl.load_workbook(model, data_only=True)["TK_Link"]
    for r in range(10, new.max_row + 1):
        nm = new.cell(r, 2).value
        if isinstance(nm, str) and nm.startswith("TK_NNV_") and old.cell(r, 10).value != new.cell(r, 10).value:
            o, nv = old.cell(r, 10).value, new.cell(r, 10).value
            ws.append([nm, new.cell(r, 1).value, o, nv, nv / o - 1, TK_USE.get(nm, TK_ONLY_SNAP)])
    if a.tk_diff_json and a.tk_diff_json.exists():
        d = json.loads(a.tk_diff_json.read_text(encoding="utf-8"))
        for it in d["changed"]:
            nm = "TK_" + it["name"]
            for c in it["cols"]:
                ws.append([nm, c["hdr"], c["old"], c["new"], c["rel"], TK_USE.get(nm, TK_ONLY_SNAP)])

    # ② 關鍵輸出
    ws = sheet("②關鍵輸出", ["輸出", "版本"] + [str(y) for y in YEARS])
    for n, lab in KEY:
        for rl, sc in (("v0.1.8", "v018"), ("v0.1.9", "base")):
            ws.append([lab, rl] + R[sc][n])
    for n, lab in SCAL:
        for rl, sc in (("v0.1.8", "v018"), ("v0.1.9", "base")):
            ws.append([lab, rl, R[sc][n]])

    # ③ 拆解、④ 敏感度
    for title, cols in (("③拆解_逐項累加", STEPS), ("③b只動該項", ONLY), ("④敏感度", SENS)):
        ws = sheet(title, ["讀法"] + H5)
        for lab, sc in cols:
            ws.append([lab] + row5(R[sc or "base"]))

    # ⑤ η 各版本
    ws = sheet("⑤η各版本", ["項目"] + [l for l, _s in ETA_VERS])
    for n in ("ETA_Eta", "ETA_fP", "ETA_fU", "ETA_fW", "ETA_Resid", "ETA_fU_Prod", "ETA_Resid2"):
        ws.append([dict(SCAL)[n]] + [R[sc or "base"][n] for _l, sc in ETA_VERS])

    # ⑥ 逐格比較
    ws = sheet("⑥逐格比較", ["比對", "比對格數", "其中數值格", "B 新增格", "不同格數", "不同格（頁!格：A → B；前 60）"])
    for p in a.diff_json:
        d = json.loads(Path(p).read_text(encoding="utf-8"))
        ws.append([Path(p).stem, d["compared"], d["numeric_compared"], d.get("added"), len(d["diff"]),
                   "；".join(f"{x[0]}：{x[1]} → {x[2]}" for x in d["diff"][:60])])

    # ⑦ 預設
    ws = sheet("⑦已套用的預設", ["問題", "採用的預設", "替代", "影響"])
    for q, d_, alt, _sc in DEFAULTS:
        ws.append([q, d_, alt, "命題輸出無影響（只動呈現或並排）"])

    del wb["Sheet"]
    a.out.parent.mkdir(parents=True, exist_ok=True)
    wb.save(a.out)
    a.json.parent.mkdir(parents=True, exist_ok=True)
    json.dump(R, open(a.json, "w", encoding="utf-8"), ensure_ascii=False, indent=1, default=str)
    print("saved", a.out, a.json)


if __name__ == "__main__":
    main()
