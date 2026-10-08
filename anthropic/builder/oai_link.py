"""OAI_Link：OpenAI v0.6 命題輸出快照（規格 D20）。

- 讀 OpenAI 活頁簿（--oai-xlsx；建置時取分支 claude/openai-s6-release 或其合併後的 main 的 openai/model/CURRENT 所指檔）的快取值，
  逐名寫入 OAI_Link 頁，並記錄檔案 SHA-256 與（若可得）git 提交。
- 只被 Checks（與 A5 的 HTML）引用；任何計算頁不得引用（tests/parity/test_builder.py 掃描；Checks 另列 builder 掃描結果）。
"""
from __future__ import annotations

import hashlib
import subprocess
from pathlib import Path

import openpyxl

# 逐年（2025–2030，6 格）的命題與主要衍生量；名稱與 OpenAI 相同，本模型名稱加前綴 OAI_
NAMES = [
    ("REV_Gross", "營收總額", "$B"), ("REV_Net", "營收淨額（扣 Microsoft 分成）", "$B"), ("REV_NetCapped", "營收淨額（容量截頂後）", "$B"),
    ("REV_Sub", "訂閱營收", "$B"), ("REV_API", "API 營收", "$B"), ("DEM_Tok_Total", "token 合計（付費＋免費）", "T"),
    ("CMP_SupplyGW", "供給 GW", "GW"), ("CMP_Supply_VReq", "供給 VR 等值 GW", "GW"), ("COST_Compute", "算力成本（現金）", "$B"),
    ("COST_FullCash", "全成本（現金，含股權報酬）", "$B"), ("COST_PropRev_VR", "每 VR 等值 GW 營收淨額", "$B/GW/年"),
    ("COST_PropFull_VR", "每 VR 等值 GW 全成本", "$B/GW/年"), ("COST_PropGap_VR", "每 VR 等值 GW 差額（命題 1）", "$B/GW/年"),
    ("COST_Coverage", "覆蓋率（營收 ÷ 全成本）", "倍"), ("FND_ExtNeed", "外部資金需求（當年）", "$B"),
    ("FND_ExtNeedCum", "累計外部資金需求（命題 2）", "$B"), ("FND_CashEnd", "年底現金", "$B"),
]


def read_oai(xlsx: Path):
    xlsx = Path(xlsx)
    sha = hashlib.sha256(xlsx.read_bytes()).hexdigest()
    commit = subprocess.run(["git", "-C", str(xlsx.parent), "rev-parse", "HEAD"], capture_output=True, text=True).stdout.strip() or "（非 git 工作目錄）"
    wb = openpyxl.load_workbook(xlsx, data_only=True)
    names = {k: v.attr_text for k, v in wb.defined_names.items()}
    rows = []
    for n, label, unit in NAMES:
        if n not in names:
            raise RuntimeError(f"OpenAI 活頁簿缺具名範圍 {n}")
        sheet, rng = names[n].split("!")
        vals = [c.value for row in wb[sheet][rng.replace("$", "")] for c in row]
        if len(vals) != 6:
            raise RuntimeError(f"{n} 不是 6 年：{len(vals)}")
        rows.append(dict(name=n, label=label, unit=unit, values=vals, ref=names[n]))
    return dict(file=xlsx.name, sha256=sha, commit=commit, rows=rows)
