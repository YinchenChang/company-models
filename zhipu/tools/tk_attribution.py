#!/usr/bin/env python3
"""v0.1.1 TK 同步歸因（以 openai/tools/tk_attribution.py 改寫）：一次只換一組 Tokenomics 名稱，以 LibreOffice 重算整本活頁簿，取主要輸出。
不在 Python 另寫計算：變體檔＝新版活頁簿把 TK_Link 指定名稱的值改回 v5.26（取自 v0.1 活頁簿的 TK_Link），交給 LibreOffice 重算，從具名範圍讀輸出。

組別（工作單 v0.1.1 第 2 節 3）：
  1 IF_TokGW_*（每 GW 年產出）
  2 IF_HoldEcon＋IF_CapexTotal＋IF_OpexGW＋IF_DeprLifeIT（持有成本、資本支出、營運費、折舊年限）
  3 IF_CostPre／Cache／Dec_*（逐 token 類型單位成本）
  4 IF_AllocRDGW（研發算力；Compute 頁只作並列對照列）
  5 其餘對照類（兩版都有、值不同、不在 1–4 者，含讀表列）
情境：OLD（五組全 v5.26）、NEW（全 v19.0）、ONLY_g（只有 g 組為 v19.0）、LOO_g（只有 g 組為 v5.26）；
另 RNG_LO／RNG_HI：NEW 但 IF_AllocRDGW 取區間下／上端點（一次換端點）。
反證 --check-old：OLD 變體與 v0.1 在計算頁逐格相同（範圍外變動＝0）。
用法：python3 tools/tk_attribution.py --new model/<v0.1.1>.xlsx --old model/archive/<v0.1>.xlsx --out <json>
"""
import argparse
import json
import re
import sys
import tempfile
from pathlib import Path

import openpyxl

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "builder"))
from common import lo_recalc  # noqa: E402

OUTPUTS = {  # 名稱 → 說明（具名範圍；向量取其全部格，純量取一格）
    "CMP_Eta2025": "η 2025", "CMP_Eta1H26": "η 1H26", "CMP_Eta": "η（各年）",
    "CMP_TokGW": "token 換算 GW（÷η 前）", "CMP_InfGW_Eff": "有效推論 GW（未截頂）", "CMP_InfGW": "推論 GW（截頂後）",
    "CMP_RDGW": "研發 GW（殘差）", "CMP_RDShare": "研發占總需求比", "CMP_DemandGW": "總需求 GW", "CMP_CapFactor": "容量截頂係數",
    "CMP_SupplyGW": "供給 GW（實體）", "CMP_Supply_Owned": "自建 GW", "CMP_VReqFactor": "機隊 VR 等值係數", "CMP_Supply_VReq": "供給 VR 等值 GW",
    "COST_InfCompute": "推論算力成本 RMB 億", "COST_RDCompute": "研發算力成本 RMB 億", "COST_Compute": "算力成本合計 RMB 億",
    "COST_SupplierHold": "供應商持有成本 RMB 億", "COST_CloudGM": "雲端毛利 RMB 億", "COST_CloudGMPct": "雲端毛利 %",
    "COST_PropRev_VR": "每 VR 等值 GW 營收 RMB 億", "COST_PropFull_VR": "每 VR 等值 GW 全成本 RMB 億",
    "COST_PropGap_VR": "每 VR 等值 GW 差額 RMB 億（命題 1）", "COST_PropGap_VR_USD": "每 VR 等值 GW 差額 $B",
    "COST_PropGap_Phys": "每實體 GW 差額 RMB 億", "COST_PropGap_Phys_USD": "每實體 GW 差額 $B", "COST_Coverage": "營收覆蓋全成本比",
    "FND_ExtNeedCum": "累計外部資金需求 RMB 億（命題 2）", "FND_ExtNeedCum_USD": "累計外部資金需求 $B", "FND_ExtNeed": "外部資金需求（年）RMB 億",
    "FND_PeakExtNeed": "外部資金需求峰值 RMB 億", "FND_PeakYear": "外部資金需求峰值年", "FND_FirstGapYear": "首次缺口年",
    "FND_CashEnd": "年底現金 RMB 億", "CHK_Errors": "Checks 錯誤數",
}
CALC_SHEETS = ["Inputs", "Demand", "Revenue", "Compute", "Cost", "Funding", "Reverse"]


def values(wb, name):
    t = wb.defined_names[name].attr_text
    sh, ref = t.rsplit("!", 1)
    ws = wb[sh.strip("'")]
    cells = ws[ref.replace("$", "")]
    if not isinstance(cells, tuple):
        return [cells.value]
    return [c.value for row in cells for c in (row if isinstance(row, tuple) else (row,))]


def label_value(wb, sheet, startswith):
    """依 A／B 欄標籤開頭找列，回傳該列 D 欄值（Compute 對照列用，免新增具名範圍）。"""
    ws = wb[sheet]
    for r in range(1, ws.max_row + 1):
        for col in (1, 2):
            v = ws.cell(r, col).value
            if isinstance(v, str) and v.startswith(startswith):
                return ws.cell(r, 4).value
    return None


def tk_rows(ws):
    return {ws.cell(r, 1).value: r for r in range(10, 300) if ws.cell(r, 1).value}


def classify(n):
    if re.fullmatch(r"IF_TokGW_(Luna|Sol|Astra)", n):
        return 1
    if n in ("IF_HoldEcon", "IF_CapexTotal", "IF_OpexGW", "IF_DeprLifeIT"):
        return 2
    if re.fullmatch(r"IF_Cost(Pre|Cache|Dec)_(Luna|Sol|Astra)", n):
        return 3
    if n == "IF_AllocRDGW":
        return 4
    return 5


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--new", type=Path, required=True)
    ap.add_argument("--old", type=Path, required=True)
    ap.add_argument("--out", type=Path, required=True)
    ap.add_argument("--tk-xlsx", type=Path, help="Tokenomics v19.0 xlsx（RNG 端點變體用；省略則取新版 TK_Link 的 RNG_ 列）")
    a = ap.parse_args()
    new_ws = openpyxl.load_workbook(a.new, data_only=True)["TK_Link"]
    old_ws = openpyxl.load_workbook(a.old, data_only=True)["TK_Link"]
    nr, orr = tk_rows(new_ws), tk_rows(old_ws)
    diff, vals = [], {}
    for n, r in nr.items():
        if n in orr:
            nv = [new_ws.cell(r, c).value for c in range(10, 25)]
            ov = [old_ws.cell(orr[n], c).value for c in range(10, 25)]
            if nv != ov:
                diff.append(n)
                vals[n] = dict(old=ov, new=nv)
    groups = {g: [n for n in diff if classify(n) == g] for g in (1, 2, 3, 4, 5)}
    print({g: len(v) for g, v in groups.items()}, "異動名稱", len(diff), flush=True)

    def run(wb, tag):
        d = Path(tempfile.mkdtemp())
        f = d / f"{tag}.xlsx"
        wb.save(f)
        rc = lo_recalc(f, d / "out")
        res = openpyxl.load_workbook(rc, data_only=True)
        out = {k: values(res, k) for k in OUTPUTS}
        out["_RD_TK_cmp"] = [label_value(res, "Compute", "對照：TK IF_AllocRDGW")]
        return out, res

    def variant(old_groups, tag):
        wb = openpyxl.load_workbook(a.new)
        ws = wb["TK_Link"]
        for g in old_groups:
            for n in groups[g]:
                for c in range(10, 25):
                    ws.cell(nr[n], c).value = old_ws.cell(orr[n], c).value
        return run(wb, tag)

    scen = {}
    scen["NEW"], _ = variant([], "NEW")
    scen["OLD"], old_res = variant([1, 2, 3, 4, 5], "OLD")
    # 反證：OLD 變體 vs v0.1（archive）逐格比對計算頁
    v01 = openpyxl.load_workbook(a.old, data_only=True)
    nd, ncell = 0, 0
    firstdiff = []
    for sh in CALC_SHEETS:
        w1, w2 = old_res[sh], v01[sh]
        for row in w1.iter_rows():
            for c in row:
                v2 = w2[c.coordinate].value
                v1 = c.value
                ncell += 1
                same = (v1 == v2) or (isinstance(v1, (int, float)) and isinstance(v2, (int, float)) and abs(v1 - v2) <= 1e-9 * max(1, abs(v2)))
                if not same:
                    nd += 1
                    if len(firstdiff) < 20:
                        firstdiff.append((sh, c.coordinate, v1, v2))
    reverse_check = dict(sheets=CALC_SHEETS, cells=ncell, diffs=nd, first=firstdiff)
    print("反證：計算頁格數", ncell, "不同", nd, flush=True)
    for g in (1, 2, 3, 4, 5):
        scen[f"ONLY_{g}"], _ = variant([x for x in (1, 2, 3, 4, 5) if x != g], f"ONLY_{g}")
        scen[f"LOO_{g}"], _ = variant([g], f"LOO_{g}")
        print("done", g, flush=True)
    # RNG 端點（一次換端點）：IF_AllocRDGW 改取 RNG_Lo／_Hi，其餘照新版
    for tag, nm_ in (("RNG_LO", "RNG_IF_AllocRDGW_Lo"), ("RNG_HI", "RNG_IF_AllocRDGW_Hi")):
        wb = openpyxl.load_workbook(a.new)
        ws = wb["TK_Link"]
        ws.cell(nr["IF_AllocRDGW"], 10).value = new_ws.cell(nr[nm_], 10).value
        scen[tag], _ = run(wb, tag)
    a.out.write_text(json.dumps(dict(outputs=OUTPUTS, groups=groups, tk_values=vals, reverse_check=reverse_check, scen=scen,
                                     rng=dict(lo=new_ws.cell(nr["RNG_IF_AllocRDGW_Lo"], 10).value, hi=new_ws.cell(nr["RNG_IF_AllocRDGW_Hi"], 10).value,
                                              point=new_ws.cell(nr["IF_AllocRDGW"], 10).value, stale=new_ws.cell(nr["RNG_StaleFlag"], 10).value)),
                                ensure_ascii=False, indent=1, default=str), encoding="utf-8")


if __name__ == "__main__":
    main()
