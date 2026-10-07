"""Tokenomics 快照讀取（工作單 P1-3；補充 1 第 1 點）。

- 值取自 Tokenomics master 的 model/CURRENT（本包為 v5.14），逐名讀出寫入 TK_Link；不用 Excel 外部連結。
- Block 6（IF_Alloc*）與 SRC_DEM_010–013 在 v5.14 尚無：逐名列入、值留空、狀態「待 v5.15 合併」。
  名稱以 v5.15 分支（未合併）Interface F 節的實際名稱為準，只取名稱不取值。
"""
from __future__ import annotations

import re
import subprocess
from pathlib import Path

import openpyxl

# 工作單 P1-3 的名稱樣式（依序）
PATTERNS = [
    r"IF_TokGW_(Luna|Sol|Astra)", r"IF_Util", r"IF_HoldEcon", r"IF_FullCost_(Luna|Sol|Astra)",
    r"IF_Price(Fresh|Cached|Think|Out|Ref)_(Luna|Sol|Astra)", r"IF_FrontRef_(Luna|Sol|Astra)",
    r"IF_ProgGWyr_(Luna|Sol|Astra)", r"IF_RDMult", r"IF_TrainGenDefault",
]
SRC_DEM = [f"SRC_DEM_{i:03d}" for i in range(4, 14)]
# 工作單暫定名 → v5.15 分支 Interface F 節實際名稱（只對名稱；值待合併）
BLOCK6_WORKORDER = ["IF_AllocQ1", "IF_AllocQ2", "IF_AllocServeGW", "IF_AllocRDGW", "IF_AllocDemand", "IF_AllocImpliedNk"]
BLOCK6_V515 = {"IF_AllocQ1": "IF_AllocQ1", "IF_AllocQ2": "IF_AllocQ2", "IF_AllocServeGW": "IF_AllocServeGW", "IF_AllocRDGW": "IF_AllocRDGW",
               "IF_AllocDemand": "IF_AllocDemand", "IF_AllocImpliedNk": "IF_AllocImpliedNk"}
BLOCK6_EXTRA_V515 = ["IF_AllocQ1_R2"]          # v5.15 新增、工作單未列
PENDING_NOTE = "待 v5.15 合併"


def _split(attr: str):
    m = re.match(r"^(?:'([^']+)'|([^!]+))!\$?([A-Z]+)\$?(\d+)(?::\$?([A-Z]+)\$?(\d+))?$", attr)
    sheet = m.group(1) or m.group(2)
    return sheet, m.group(3), int(m.group(4)), m.group(5), (int(m.group(6)) if m.group(6) else None)


def read_snapshot(tk_dir: Path):
    tk_dir = Path(tk_dir)
    current = (tk_dir / "model" / "CURRENT").read_text(encoding="utf-8").strip()
    path = tk_dir / "model" / current
    sha = subprocess.run(["git", "-C", str(tk_dir), "rev-parse", "HEAD"], capture_output=True, text=True).stdout.strip()
    m = re.search(r"_v(\d+(?:\.\d+)?)\.xlsx$", current)
    wb = openpyxl.load_workbook(path, data_only=True)
    names = {k: v.attr_text for k, v in wb.defined_names.items()}
    chosen = []
    for pat in PATTERNS:
        chosen += sorted(n for n in names if re.fullmatch(pat, n))
    chosen += [n for n in SRC_DEM if n in names]
    rows, seen = [], set()
    for n in chosen:
        if n in seen:
            continue
        seen.add(n)
        sheet, c1, r1, c2, r2 = _split(names[n])
        ws = wb[sheet]
        if c2:
            from openpyxl.utils import column_index_from_string as ci
            vals = [ws.cell(r1, k).value for k in range(ci(c1), ci(c2) + 1)]
        else:
            vals = [ws[f"{c1}{r1}"].value]
        if sheet == "Interface":
            label = re.sub(r"[\s　]*\[IF_[^\]]*\]\s*$", "", str(ws.cell(r1, 1).value))
            unit = ws.cell(r1, 2).value or ""
        else:
            label = f"{ws.cell(r1, 2).value}（{ws.cell(r1, 7).value}）"
            unit = ws.cell(r1, 6).value or ""
        rows.append(dict(name=n, kind="SRC" if n.startswith("SRC_") else "IF", label=label, unit=unit, values=vals, status="OK"))
    present = {r["name"] for r in rows}
    pending = [dict(name=n, kind="SRC" if n.startswith("SRC_") else "IF", label="", unit="", values=[], status=PENDING_NOTE,
                    v515=BLOCK6_V515.get(n, n)) for n in SRC_DEM if n not in present]
    pending = [dict(name=n, kind="IF", label="Block 6（Alloc）", unit="", values=[], status=PENDING_NOTE, v515=BLOCK6_V515[n])
               for n in BLOCK6_WORKORDER if n not in names] + pending
    hdr_gen = [wb["Interface"].cell(4, k).value for k in range(3, 18)]
    hdr_cost = [wb["Interface"].cell(5, k).value for k in range(3, 18)]
    return dict(file=current, version=f"v{m.group(1)}" if m else "?", sha=sha, rows=rows, pending=pending,
                hdr_gen=hdr_gen, hdr_cost=hdr_cost)
