#!/usr/bin/env python3
"""Val 情境 A／C 快照（工作單 20261111 F7）：改 Inputs → LibreOffice 重算 → 讀 Val 第 3 節「情境 B」來源列（活公式）→ 寫 data/val_scenarios.yaml。

情境（J7 預設）：
- A：API 任務數年增率（api_gn_2027–2030）取 Inputs 高端；2027 營收依 2026Q4 年化校準（D2 讀法）——基準 D2 比值 ≥ 1（模型 2027 已高於 2026Q4 年化），
  高端任務數使 2027 更高，校準下限不生效，故 A 不另加機制（報告列出）。
- B：基準（快照只供過期檢查）。
- C：API 任務數年增率取低端＋2026Q4＝Q3（p_q4_rr_switch＝0）。
用法：python3 tools/val_snapshots.py model/<現行>.xlsx
"""
from __future__ import annotations

import datetime as dt
import sys
from pathlib import Path

import openpyxl
import yaml

sys.path.insert(0, str(Path(__file__).resolve().parent))
from val_recalc import ROOT, run_many  # noqa: E402

sys.path.insert(0, str(ROOT / "builder"))
from a11 import SRC_ZH, SRCK  # noqa: E402

GN = [f"api_gn_{y}" for y in (2027, 2028, 2029, 2030)]
SCN = {
    "A": dict(desc="API 任務數年增率取高端（2027–2030）；2027 營收依 2026Q4 年化校準（下限不生效）", ov={k: "hi" for k in GN}),
    "B": dict(desc="基準", ov={}),
    "C": dict(desc="API 任務數年增率取低端（2027–2030）＋2026Q4＝Q3（p_q4_rr_switch＝0）", ov={**{k: "lo" for k in GN}, "p_q4_rr_switch": 0}),
}


def reader(path):
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb["Val"]
    lab = {ws.cell(r, 1).value: r for r in range(1, ws.max_row + 1) if ws.cell(r, 1).value}
    vals = {}
    for k in SRCK:
        r = lab[f"{SRC_ZH[k]}（B）"]
        vals[k] = [float(ws.cell(r, c).value) for c in range(3, 8)]
    extra = {n: wb.defined_names[n] for n in ("REV_D2Ratio",)}
    sh, ref = extra["REV_D2Ratio"].attr_text.split("!")
    d2 = wb[sh][ref.replace("$", "")].value
    return dict(values=vals, d2=d2)


def main():
    model = Path(sys.argv[1]).resolve()
    res = run_many(model, {s: v["ov"] for s, v in SCN.items()}, reader, workers=3)
    out = dict(note="tools/val_snapshots.py 產生；值＝LibreOffice 重算後的 Val 第 3 節 B 列（各情境改 Inputs 後）；勿手改",
               model=model.name, date=dt.date.today().isoformat(), scenarios={})
    for s, v in SCN.items():
        out["scenarios"][s] = dict(id=f"VAL_SCN_{s}", desc=v["desc"], overrides={k: str(x) for k, x in v["ov"].items()}, model=model.name,
                                   d2_ratio=res[s]["d2"], values=res[s]["values"])
    p = ROOT / "data" / "val_scenarios.yaml"
    p.write_text(yaml.safe_dump(out, allow_unicode=True, sort_keys=False, width=200), encoding="utf-8")
    for s in SCN:
        print(s, "D2", round(res[s]["d2"], 4), {k: [round(x, 3) for x in v] for k, v in res[s]["values"].items()})


if __name__ == "__main__":
    main()
