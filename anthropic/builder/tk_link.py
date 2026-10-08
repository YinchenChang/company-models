"""Tokenomics 快照讀取（沿用 OpenAI v0.6 的 63 名；Anthropic A2 加讀 NonNV 表，規格 D6）。

- 值取自 Tokenomics master 的 model/CURRENT，逐名讀出寫入 TK_Link；不用 Excel 外部連結。
- 列序：P1 原 37 名（列位不變）→ SRC_DEM_010–013（＋013 低／高）→ Block 6（IF_Alloc* 7 名）→ S1 新增（r6 下游需要）→ S4 新增（IF_CapexTotal）。
- 工作單要求、但 Tokenomics 現行版沒有的名稱：逐名列入、值留空、狀態 PENDING_NOTE（E1；check_tk_snapshot 報 MISSING）。
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
SRC_DEM = [f"SRC_DEM_{i:03d}" for i in range(4, 14)] + ["SRC_DEM_013_Lo", "SRC_DEM_013_Hi"]   # 013 只有低／高（無點值）
# Block 6（v5.15 起）：工作單暫定名與 Tokenomics 實際名稱相同；IF_AllocQ1_R2 為 r2 補列
BLOCK6 = ["IF_AllocQ1", "IF_AllocQ1_R2", "IF_AllocQ2", "IF_AllocServeGW", "IF_AllocRDGW", "IF_AllocDemand", "IF_AllocImpliedNk"]
# S1 新增（工程類；r6 下游會用到）：訓練成本（P3 研發檢查）、每 GW 理論營收（P4 命題對照）、
# 全成本下游預設（Tokenomics 標為「下游預設」；P4 對照）、L1 問 3 研發算力成本（P4 Q3，r6 P4-2）
EXTRA_PATTERNS = [
    r"IF_TrainCost_(Luna|Sol|Astra)", r"IF_RevGW_(Luna|Sol|Astra)", r"IF_FullCostDefault_(Luna|Sol|Astra)",
]
EXTRA_NAMES = ["L1_Ans3", "L1_Ans3_Lo", "L1_Ans3_Hi"]
# S4 新增（工程類；P4 自建 GW 需要每 GW 資本支出；列於最後，既有列位不變）
S4_NAMES = ["IF_CapexTotal"]
# Anthropic A2 新增（工程類；任務類別每任務 token 取自 Tokenomics，不在本模型自設）：任務名稱、任務長度、每次嘗試新鮮／快取／decode token
A2_NAMES = ["IF_HdrTask", "IF_TaskLen", "IF_TaskTokFresh", "IF_TaskTokCached", "IF_TaskTokDec"]
PENDING_NOTE = "待 Tokenomics 提供"


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
    chosen += [n for n in BLOCK6 if n in names]
    for pat in EXTRA_PATTERNS:
        chosen += sorted(n for n in names if re.fullmatch(pat, n))
    chosen += [n for n in EXTRA_NAMES if n in names]
    chosen += [n for n in S4_NAMES if n in names]
    chosen += [n for n in A2_NAMES if n in names]
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
        elif sheet == "L1":
            col = {"D": "基準", "E": "低", "F": "高"}.get(c1, c1)
            label = f"{ws.cell(r1, 2).value}（{col}）"
            unit = ws.cell(r1, 7).value or ""
        else:
            label = f"{ws.cell(r1, 2).value}（{ws.cell(r1, 7).value}）"
            unit = ws.cell(r1, 6).value or ""
        rows.append(dict(name=n, kind=n.split("_")[0], label=label, unit=unit, values=vals, status="OK"))
    present = {r["name"] for r in rows}
    # 工作單要求、但 Tokenomics 現行版沒有者：逐名列入、值留空（E1）
    wanted = SRC_DEM + BLOCK6 + EXTRA_NAMES + S4_NAMES + A2_NAMES
    pending = [dict(name=n, kind=n.split("_")[0], label="", unit="", values=[], status=PENDING_NOTE) for n in wanted if n not in present]
    hdr_gen = [wb["Interface"].cell(4, k).value for k in range(3, 18)]
    hdr_cost = [wb["Interface"].cell(5, k).value for k in range(3, 18)]
    return dict(file=current, version=f"v{m.group(1)}" if m else "?", sha=sha, rows=rows, pending=pending,
                hdr_gen=hdr_gen, hdr_cost=hdr_cost)


# ── Anthropic A2（規格 D6、D6 r1）：NonNV 表不是具名範圍，以列標籤＋欄標題讀表 ──
NONNV_SHEET = "NonNV"
NONNV_FAMILIES = [("TPUv7", "Google TPU v7 Ironwood"), ("Trn3", "AWS Trainium3"), ("MI455X", "AMD MI455X（Helios）")]
NONNV_COLS = [("Out", "產出比 基準"), ("OutLo", "產出比 低"), ("OutHi", "產出比 高"),
              ("Hold", "持有比 基準"), ("HoldLo", "持有比 低"), ("HoldHi", "持有比 高")]
NONNV_STATUS = "讀表（非具名）"


# 讀表的頁與標題列（A3 起加 Tokenomics Inputs 頁的 PUE 列：Tokenomics 未設具名範圍；工作單 A3「PUE：TK 若有則改引用」）
TABLE_HDR_ROW = {"NonNV": 4, "Inputs": 3}
TABLE_ROWS = [("Inputs", "PUE", [("TK_PUE", "基準"), ("TK_PUE_Lo", "低成本"), ("TK_PUE_Hi", "高成本")], "x（設施電力 ÷ IT 電力）")]


def table_lookup(wb, sheet: str, label: str, hdr: str):
    """在 sheet 頁找 A 欄＝label 的列、標題列（TABLE_HDR_ROW）＝hdr 的欄；回傳 (值, 儲存格位址)；找不到回傳 (None, None)。"""
    if sheet not in wb.sheetnames:
        return None, None
    ws = wb[sheet]
    h = TABLE_HDR_ROW.get(sheet, 4)
    col = next((c for c in range(1, ws.max_column + 1) if ws.cell(h, c).value == hdr), None)
    row = next((r for r in range(h + 1, ws.max_row + 1) if ws.cell(r, 1).value == label), None)
    if col is None or row is None:
        return None, None
    return ws.cell(row, col).value, ws.cell(row, col).coordinate


def nonnv_lookup(wb, label: str, hdr: str):
    """在 NonNV 頁找 A 欄＝label 的列、第 4 列標題＝hdr 的欄；回傳 (值, 儲存格位址)；找不到回傳 (None, None)。"""
    return table_lookup(wb, NONNV_SHEET, label, hdr)


def read_nonnv(tk_dir: Path):
    tk_dir = Path(tk_dir)
    current = (tk_dir / "model" / "CURRENT").read_text(encoding="utf-8").strip()
    wb = openpyxl.load_workbook(tk_dir / "model" / current, data_only=True)
    rows = []
    for fam, label in NONNV_FAMILIES:
        for suf, hdr in NONNV_COLS:
            v, addr = nonnv_lookup(wb, label, hdr)
            if v is None:
                raise RuntimeError(f"Tokenomics NonNV 找不到：{label}／{hdr}")
            rows.append(dict(name=f"{NONNV_SHEET}!{label}!{hdr}", our=f"TK_NNV_{fam}_{suf}", label=f"{label}：{hdr}（相對 VR200；NonNV!{addr}）",
                             unit="倍", values=[v], status=NONNV_STATUS))
    for sheet, label, cols, unit in TABLE_ROWS:          # A3：PUE（Tokenomics Inputs 頁，非具名）
        for our, hdr in cols:
            v, addr = table_lookup(wb, sheet, label, hdr)
            if v is None:
                raise RuntimeError(f"Tokenomics {sheet} 找不到：{label}／{hdr}")
            rows.append(dict(name=f"{sheet}!{label}!{hdr}", our=our, label=f"{label}：{hdr}（Tokenomics {sheet}!{addr}；非具名）",
                             unit=unit, values=[v], status=NONNV_STATUS))
    return rows
