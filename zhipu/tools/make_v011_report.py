#!/usr/bin/env python3
"""v0.1.1 TK 同步報告表：讀 tools/tk_attribution.py 的 JSON（LibreOffice 重算結果），輸出對照 Excel 與報告用 Markdown 表。
不另算模型數值；差額、交互項為對重算結果的減法。

用法：python3 tools/make_v011_report.py --attr <json> --defaults docs/reports/<報告>.md --out-xlsx docs/reports/<>_對照.xlsx --out-md <表.md>
"""
import argparse
import json
import re
from pathlib import Path

import openpyxl
from openpyxl.styles import Alignment, Font, PatternFill

YEARS = [2025, 2026, 2027, 2028, 2029, 2030]
BOLD = Font(bold=True)
HDR = PatternFill("solid", fgColor="FFEDEDED")
WRAP = Alignment(wrap_text=True, vertical="top")
GROUPS = {1: "①IF_TokGW_*", 2: "②HoldEcon＋Capex＋OpexGW", 3: "③IF_Cost*", 4: "④IF_AllocRDGW", 5: "⑤其餘對照類"}
# 報告主表的輸出（名稱, 顯示年份索引或 None＝全部六年, 小數位）
KEY = ["CMP_Eta2025", "CMP_Eta1H26", "CMP_TokGW", "CMP_InfGW_Eff", "CMP_RDGW", "CMP_RDShare", "CMP_DemandGW", "CMP_CapFactor", "CMP_SupplyGW",
       "CMP_VReqFactor", "CMP_Supply_VReq", "COST_InfCompute", "COST_RDCompute", "COST_Compute", "COST_SupplierHold", "COST_CloudGM", "COST_CloudGMPct",
       "COST_PropRev_VR", "COST_PropFull_VR", "COST_PropGap_VR", "COST_PropGap_VR_USD", "COST_PropGap_Phys", "COST_PropGap_Phys_USD", "COST_Coverage",
       "FND_ExtNeedCum", "FND_ExtNeedCum_USD", "FND_PeakExtNeed", "FND_PeakYear", "FND_FirstGapYear", "FND_CashEnd"]


def num(x):
    return isinstance(x, (int, float)) and not isinstance(x, bool)


def fmt(x, d=3):
    if x is None:
        return "—"
    if not num(x):
        return str(x)
    if float(x).is_integer() and 1990 <= x <= 2100:
        return str(int(x))
    return f"{x:,.{d}f}" if abs(x) < 1000 else f"{x:,.1f}"


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--attr", type=Path, required=True)
    ap.add_argument("--defaults", type=Path, required=True)
    ap.add_argument("--out-xlsx", type=Path, required=True)
    ap.add_argument("--out-md", type=Path, required=True)
    ap.add_argument("--tk-header", type=Path, help="新版活頁簿（讀 TK_Link 第 7、8 列欄標）")
    a = ap.parse_args()
    d = json.loads(a.attr.read_text(encoding="utf-8"))
    S, O = d["scen"], d["outputs"]
    gen = cost = [None] * 15
    if a.tk_header:
        ws = openpyxl.load_workbook(a.tk_header, data_only=True)["TK_Link"]
        gen = [ws.cell(7, 10 + k).value for k in range(15)]
        cost = [ws.cell(8, 10 + k).value for k in range(15)]

    wb = openpyxl.Workbook()
    # ① TK 名稱前後值
    w1 = wb.active
    w1.title = "①TK名稱前後值"
    w1.append(["Tokenomics 名稱", "歸因組", "世代", "成本情境", "v5.26（v0.1）", "v19.0（v0.1.1）", "變動 %"])
    n1 = 0
    for g, names in d["groups"].items():
        for n in names:
            v = d["tk_values"][n]
            for k in range(15):
                o, nw = v["old"][k], v["new"][k]
                if o is None and nw is None:
                    continue
                pct = (nw / o - 1) if num(o) and num(nw) and o not in (0, 0.0) else None
                w1.append([n, GROUPS[int(g)], gen[k], cost[k], o, nw, pct])
                n1 += 1
    # ② 主要輸出前後值與歸因
    w2 = wb.create_sheet("②主要輸出前後值與歸因")
    hdr = ["輸出", "代號", "年", "v0.1（全 v5.26）", "v0.1.1（全 v19.0）", "差額"] + [f"只換 {GROUPS[g]}" for g in (1, 2, 3, 4, 5)] + ["交互項（差額 − Σ只換）", "只還原一組之差（NEW−LOO）①", "②", "③", "④", "⑤"]
    w2.append(hdr)
    md = []
    md.append("| 輸出 | 年 | v0.1 | v0.1.1 | 差額 | ① TokGW | ② Hold/Capex/Opex | ③ Cost* | ④ AllocRDGW | ⑤ 其餘 | 交互 |")
    md.append("|---|---|---|---|---|---|---|---|---|---|---|")
    maxint = 0.0
    for key in O:
        if key == "CHK_Errors":
            continue
        vo, vn = S["OLD"][key], S["NEW"][key]
        for i in range(len(vn)):
            y = YEARS[i] if len(vn) == 6 else ""
            o, nw = vo[i], vn[i]
            if not (num(o) and num(nw)):
                w2.append([O[key], key, y, o, nw])
                continue
            dl = nw - o
            only = [S[f"ONLY_{g}"][key][i] - o if num(S[f"ONLY_{g}"][key][i]) else None for g in (1, 2, 3, 4, 5)]
            loo = [nw - S[f"LOO_{g}"][key][i] if num(S[f"LOO_{g}"][key][i]) else None for g in (1, 2, 3, 4, 5)]
            inter = dl - sum(x for x in only if x is not None)
            maxint = max(maxint, abs(inter) / max(1e-12, abs(dl)) if abs(dl) > 1e-9 else 0)
            w2.append([O[key], key, y, o, nw, dl] + only + [inter] + loo)
            if key in KEY and (len(vn) == 1 or y in (2025, 2027, 2030)):
                md.append(f"| {O[key]} | {y or '—'} | {fmt(o)} | {fmt(nw)} | {fmt(dl)} | " + " | ".join(fmt(x) if abs(x or 0) > 1e-9 else "0" for x in only) + f" | {fmt(inter) if abs(inter) > 1e-9 else '0'} |")
    # RNG 端點
    w2.append([])
    w2.append(["IF_AllocRDGW 區間端點（一次換端點）", "", "", "點值", "下端 Lo", "上端 Hi", "Compute 並列列值", "", "", "", "", "是否影響任何主輸出"])
    changed = [k for k in O if S["RNG_LO"][k] != S["NEW"][k] or S["RNG_HI"][k] != S["NEW"][k]]
    w2.append(["", "", "", d["rng"]["point"], d["rng"]["lo"], d["rng"]["hi"], f"{S['NEW']['_RD_TK_cmp'][0]:.4f}／{S['RNG_LO']['_RD_TK_cmp'][0]:.4f}／{S['RNG_HI']['_RD_TK_cmp'][0]:.4f}", "", "", "", "",
               "否（僅並列對照列）" if not changed else "是：" + "、".join(changed)])
    # ③ 已套用的預設（由報告 md 解析）
    w3 = wb.create_sheet("③已套用的預設")
    w3.append(["編號", "問題", "採用的預設", "替代選項", "對結果的影響方向"])
    txt = a.defaults.read_text(encoding="utf-8")
    m = re.search(r"^## \d+\. 已套用的預設.*?$(.*?)(?=^## |\Z)", txt, re.S | re.M)
    sec = m.group(1) if m else ""
    for line in sec.splitlines():
        if line.startswith("|") and not re.match(r"^\|[-| ]+\|$", line):
            cells = [c.strip() for c in line.strip("|").split("|")]
            if cells and cells[0] not in ("#", "") and len(cells) >= 5:
                w3.append(cells[:5])
    # ④ 反證與缺口
    w4 = wb.create_sheet("④反證與缺口")
    rc = d["reverse_check"]
    w4.append(["反證：全部 TK 名稱改回 v5.26，計算頁與 v0.1 逐格比對"])
    w4.append(["比對工作表", "、".join(rc["sheets"])])
    w4.append(["比對格數", rc["cells"]])
    w4.append(["不同格數（範圍外變動）", rc["diffs"]])
    for ws in wb.worksheets:
        for c in ws[1]:
            c.font, c.fill = BOLD, HDR
        for col in ws.columns:
            ws.column_dimensions[col[0].column_letter].width = 22
        for row in ws.iter_rows():
            for c in row:
                c.alignment = WRAP
    wb.save(a.out_xlsx)
    a.out_md.write_text("\n".join(md) + f"\n\n最大交互項占差額比：{maxint:.4%}；①表列數 {n1}\n", encoding="utf-8")
    print("saved", a.out_xlsx, a.out_md, "max interaction ratio", maxint)


if __name__ == "__main__":
    main()
