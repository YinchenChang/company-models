"""OAI_Link：OpenAI v0.6 命題輸出快照（規格 D26）。

讀 OpenAI v0.6 Excel（LibreOffice 重算後的快取值），逐名取 FY2025–FY2030 六個值寫入 OAI_Link 頁；記錄檔名與 SHA-256。
只被 Checks 與 HTML 引用（test_builder 掃描：其他頁的公式不得含 OAI_）。數值為美元（$B 或 $B/GW），不在本模型換算或計算。
"""
from __future__ import annotations

import hashlib
import re
from pathlib import Path

import openpyxl

# (OpenAI 具名範圍, 說明, 單位)
NAMES = [
    ("REV_Gross", "營收總額（未截頂）", "$B"),
    ("REV_NetCapped", "營收淨額（容量截頂後）", "$B"),
    ("DEM_Tok_Total", "token 合計（付費＋免費）", "T"),
    ("CMP_InfGW_Eff", "有效推論 GW", "GW"),
    ("CMP_Supply_VReq", "供給 VR 等值 GW", "GW"),
    ("COST_Compute", "算力成本（現金）", "$B"),
    ("COST_FullCash", "全成本（現金，含股權報酬）", "$B"),
    ("COST_PropRev_VR", "每 VR 等值 GW 營收淨額", "$B/GW"),
    ("COST_PropFull_VR", "每 VR 等值 GW 全成本", "$B/GW"),
    ("COST_PropGap_VR", "每 VR 等值 GW 差額（命題 1）", "$B/GW"),
    ("COST_Coverage", "覆蓋率（營收 ÷ 全成本）", "比例"),
    ("FND_FCF", "自由現金流", "$B"),
    ("FND_ExtNeed", "外部資金需求（單年）", "$B"),
    ("FND_ExtNeedCum", "累計外部資金需求（命題 2）", "$B"),
    ("FND_CashEnd", "年底現金", "$B"),
    ("FND_Committed", "已到位融資", "$B"),
    ("RVS_ExtNeedCum", "反向累計外部資金需求（管理層目標）", "$B"),
]


def _vals(wb, attr):
    m = re.match(r"^(?:'([^']+)'|([^!]+))!\$?([A-Z]+)\$?(\d+)(?::\$?([A-Z]+)\$?(\d+))?$", attr)
    ws = wb[m.group(1) or m.group(2)]
    from openpyxl.utils import column_index_from_string as ci
    c1, r1 = ci(m.group(3)), int(m.group(4))
    c2 = ci(m.group(5)) if m.group(5) else c1
    return [ws.cell(r1, c).value for c in range(c1, c2 + 1)]


def read_oai(path: Path, ref: str):
    path = Path(path)
    sha = hashlib.sha256(path.read_bytes()).hexdigest()
    wb = openpyxl.load_workbook(path, data_only=True)
    names = {k: v.attr_text for k, v in wb.defined_names.items()}
    rows = []
    for n, lab, unit in NAMES:
        vals = _vals(wb, names[n]) if n in names else []
        rows.append(dict(name=n, label=lab, unit=unit, values=vals, status="OK" if len(vals) == 6 else "MISSING"))
    return dict(file=path.name, sha=sha, ref=ref, rows=rows)
