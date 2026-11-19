#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""三層對帳總表產生程式（工作單 A6）。

只讀來源檔、不改 repo。所有數字都從檔案讀入，寫進「輸入」頁（藍字）；
其他頁的數字一律是連到「輸入」頁的公式（黑字）。

用法：
  python3 build_summary.py [--company-models DIR] [--tokenomics DIR] [--out DIR]
         [--date YYYYMMDD] [--ver v0.1] [--oos-q3 107.6]
各來源檔路徑也可個別覆寫（--tk-snapshot 等），見 --help。

硬性規則：NVIDIA FY2027 Q3（--oos-q3）只出現在「實績對帳」頁標示為「樣本外檢驗（未入模型）」的區塊，
不被任何其他格引用（建檔後由 verify 步驟再掃一次公式確認）。
"""
import argparse
import hashlib
import json
import os
import re
import subprocess
import sys

import openpyxl
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

# ----------------------------------------------------------------------------------------------
# 參數
# ----------------------------------------------------------------------------------------------
ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
ap.add_argument("--company-models", default="/home/claude/company-models")
ap.add_argument("--tokenomics", default="/home/claude/tokenomics")
ap.add_argument("--out", default="/home/claude/out/summary")
ap.add_argument("--date", default="20261118")
ap.add_argument("--ver", default="v0.1")
ap.add_argument("--tk-snapshot")
ap.add_argument("--watt-snapshot")
ap.add_argument("--sil-snapshot")
ap.add_argument("--status")
ap.add_argument("--recon")
ap.add_argument("--bridge")
ap.add_argument("--tk-xlsx")
ap.add_argument("--watt-xlsx")
ap.add_argument("--sil-xlsx")
ap.add_argument("--oos-q3", type=float, default=107.6,
                help="NVIDIA FY2027 Q3 資料中心營收（US$B）。只供「樣本外檢驗」區塊，不入模型")
ap.add_argument("--oos-source", default="Andy 2026-11-18 對帳單 r3（未核對原文）")
ap.add_argument("--skip-usage-scan", action="store_true", help="略過掃描兩本模型的 TK 公式引用（較快，但介面覆蓋頁引用次數為空）")
A = ap.parse_args()

CM, TKR = A.company_models, A.tokenomics


def _cur(path):
    with open(path, encoding="utf-8") as f:
        return f.read().strip()


def _latest(folder, pat):
    c = sorted(x for x in os.listdir(folder) if re.fullmatch(pat, x))
    return os.path.join(folder, c[-1])


_W_X = _cur(f"{CM}/wattonomics/model/CURRENT")
_W_V = re.search(r"_(v[\d.]+)\.xlsx$", _W_X).group(1)
P = dict(
    tk_snapshot=A.tk_snapshot or f"{CM}/wattonomics/data/tokenomics_snapshot_" + re.search(r"_(v[\d.]+)\.xlsx$", _cur(f"{TKR}/model/CURRENT")).group(1) + ".json",
    watt_snapshot=A.watt_snapshot or f"{CM}/wattonomics/data/wattonomics_snapshot_{_W_V}.json",
    sil_snapshot=A.sil_snapshot or f"{CM}/siliconomics/data/" + _cur(f"{CM}/siliconomics/data/CURRENT_SNAPSHOT"),
    status=A.status or f"{CM}/siliconomics/data/baseline_status.json",
    recon=A.recon or f"{CM}/siliconomics/data/" + _cur(f"{CM}/siliconomics/data/CURRENT_RECON"),
    bridge=A.bridge or _latest(f"{CM}/siliconomics/docs/reports", r"\d{8}_協調橋接_" + re.escape(_W_V) + r"(r\d+)?\.xlsx"),
    tk_xlsx=A.tk_xlsx or f"{TKR}/model/" + _cur(f"{TKR}/model/CURRENT"),
    watt_xlsx=A.watt_xlsx or f"{CM}/wattonomics/model/{_W_X}",
    sil_xlsx=A.sil_xlsx or f"{CM}/siliconomics/model/" + _cur(f"{CM}/siliconomics/model/CURRENT"),
)
YEARS = [2025, 2026, 2027, 2028, 2029]
GENS = ["Hopper H100", "GB200 NVL72", "GB300 NVL72", "VR200 NVL72", "Rubin Ultra（MGX NVL 單架，72 GPU）"]
GEN_SHORT = ["H100", "GB200", "GB300", "VR200", "RU"]
SCEN = ["低成本", "基準", "高成本"]


def sha256(path):
    h = hashlib.sha256()
    with open(path, "rb") as f:
        for b in iter(lambda: f.read(1 << 20), b""):
            h.update(b)
    return h.hexdigest()


def git_head(repo):
    try:
        h = subprocess.check_output(["git", "-C", repo, "rev-parse", "HEAD"], text=True, stderr=subprocess.DEVNULL).strip()
        dirty = subprocess.check_output(["git", "-C", repo, "status", "--porcelain"], text=True, stderr=subprocess.DEVNULL).strip()
        return h, ("工作樹有未提交變更" if dirty else "工作樹乾淨")
    except Exception:
        return "—", "非 git 目錄"


def jload(p):
    with open(p, encoding="utf-8") as f:
        return json.load(f)


# ----------------------------------------------------------------------------------------------
# 讀來源
# ----------------------------------------------------------------------------------------------
tk_snap = jload(P["tk_snapshot"])
w_snap = jload(P["watt_snapshot"])
s_snap = jload(P["sil_snapshot"])
status = jload(P["status"])
recon = jload(P["recon"])
WN = w_snap["names"]
SN = s_snap["names"]

SHA = {k: sha256(v) for k, v in P.items()}


def _ver(path):
    return re.search(r"_(v[\d.]+(?:r\d+)?)\.(?:xlsx|json)$", os.path.basename(path)).group(1)


TV, WV, SV, RV, BV = _ver(P["tk_xlsx"]), _ver(P["watt_xlsx"]), _ver(P["sil_xlsx"]), _ver(P["recon"]), _ver(P["bridge"])
GIT_CM = git_head(CM)
GIT_TK = git_head(TKR)

# 來源檔登記（代號 → 顯示用）
FILES = {
    "TKS": dict(file=os.path.basename(P["tk_snapshot"]), ver=tk_snap["source"]["version"],
                rec=None, act=SHA["tk_xlsx"], model="Tokenomics 模型檔"),
    "TKX": dict(file=os.path.basename(P["tk_xlsx"]), ver=TV, rec=None, act=SHA["tk_xlsx"], model="Tokenomics 模型檔"),
    "WAS": dict(file=os.path.basename(P["watt_snapshot"]), ver=w_snap["source"]["version"],
                rec=w_snap["source"]["sha256"], act=SHA["watt_xlsx"], model="Wattonomics 模型檔"),
    "WAX": dict(file=os.path.basename(P["watt_xlsx"]), ver=WV, rec=None, act=SHA["watt_xlsx"], model="Wattonomics 模型檔"),
    "SIS": dict(file=os.path.basename(P["sil_snapshot"]), ver=s_snap["source"]["version"],
                rec=s_snap["source"]["sha256"], act=SHA["sil_xlsx"], model="Siliconomics 模型檔"),
    "SIB": dict(file=os.path.basename(P["status"]), ver="（來源 " + status["source"] + "）", rec=None, act=SHA["sil_xlsx"], model="Siliconomics 模型檔"),
    "SIR": dict(file=os.path.basename(P["recon"]), ver=RV, rec=None, act=SHA["sil_xlsx"], model="Siliconomics 模型檔"),
    "BRG": dict(file=os.path.basename(P["bridge"]), ver=None, rec=None, act=SHA["bridge"], model="橋接表本檔"),
    "EXT": dict(file="（外部／對帳單）", ver="—", rec=None, act=None, model="—"),
}

# ---- Tokenomics 具名範圍（xlsx，data_only）----
tk_wb = openpyxl.load_workbook(P["tk_xlsx"], data_only=True)


def tk_named(name):
    t = tk_wb.defined_names[name].attr_text
    sh, rng = t.split("!")
    ws = tk_wb[sh.strip("'")]
    cells = ws[rng.replace("$", "")]
    vals = [c.value for c in cells[0]]
    label = ws.cell(cells[0][0].row, 1).value
    unit = ws.cell(cells[0][0].row, 2).value
    return t, vals, label, unit


TK_NAMES_XLSX = ["IF_CapexIT", "IF_CapexTotal", "IF_CapexFacility", "IF_GPUsPerGW", "IF_HoldEcon", "IF_HoldAcct", "IF_RevGWFleet"]
TKX = {}
for nm in TK_NAMES_XLSX:
    t, vals, label, unit = tk_named(nm)
    assert len(vals) == 15, (nm, len(vals))
    TKX[nm] = dict(ref=t, vals=vals, label=label, unit=unit)

# 快照內 23 個名稱（含 optional IF_FacilityGW）
TK_ITEMS = tk_snap["items"]
assert len(TK_ITEMS) == 23, len(TK_ITEMS)


def snap_base(nm, gen):
    it = TK_ITEMS.get(nm)
    if not it:
        return None
    v = it["values"]
    if isinstance(v, dict) and gen in v:
        return v[gen].get("基準") if isinstance(v[gen], dict) else v[gen]
    return None


# ---- Wattonomics 世代組合（Players!D5:M9；快照只有空白）----
w_wb = openpyxl.load_workbook(P["watt_xlsx"], data_only=True)
pl = w_wb["Players"]
MIX = {}
for r in range(5, 10):
    key = pl.cell(r, 2).value  # mix_H100 ...
    MIX[key] = {2023 + i: pl.cell(r, 4 + i).value for i in range(10)}
MIX_ORDER = ["mix_H100", "mix_GB200", "mix_GB300", "mix_VR200", "mix_RU"]
for k in MIX_ORDER:
    assert k in MIX, k

# ---- Wattonomics 名稱 ----


def wv(name, y):
    return WN.get(f"{name}_{y}")


# ---- Siliconomics 名稱 ----


def sarr(name):
    a = SN[name]
    if isinstance(a, list):
        n = len(a)
        return {2030 - n + i: a[i] for i in range(n)}
    return {}


def sv(name, y):
    v = SN.get(f"{name}_{y}")
    if v is None and isinstance(SN.get(name), list):
        v = sarr(name).get(y)
    return v


# ---- 橋接表 ----
br_wb = openpyxl.load_workbook(P["bridge"], data_only=True)
br_in, br_bridge = br_wb["In"], br_wb["Bridge"]
BRIN = {}
for r in range(4, br_in.max_row + 1):
    k = br_in.cell(r, 1).value
    if k:
        BRIN[k] = dict(low=br_in.cell(r, 5).value, base=br_in.cell(r, 6).value, high=br_in.cell(r, 7).value,
                       src=br_in.cell(r, 9).value, year=br_in.cell(r, 4).value)
BR = {}
for r in range(4, br_bridge.max_row + 1):
    k = br_bridge.cell(r, 1).value
    if k:
        BR[k] = {2025 + i: br_bridge.cell(r, 4 + i).value for i in range(5)}
BR_IN_NOTE = br_in.cell(1, 1).value
BSV = re.search(r"Siliconomics_(v[\d.]+)\.xlsx", BR_IN_NOTE).group(1)   # 橋接表 In 頁標示的 Siliconomics 輸入版本
BR_SAME = BSV == SV
FILES["BRG"]["ver"] = f"{BV}（In 頁標示 Siliconomics {BSV} 輸入）"
_br_s0_29, _sn_b29 = BR["s0"].get(2029), SN["SIF_DeployGW_Base"][-1] if isinstance(SN["SIF_DeployGW_Base"], list) else None

# ---- 對帳點 ----
MK = recon["markets"]
TOL = recon["tolerance"]
nv = MK["nv_dc"]
m = re.search(r"Q1\s*([\d.]+)、Q2\s*([\d.]+)", nv["ref_actual"])
NV_Q1, NV_Q2 = float(m.group(1)), float(m.group(2))
assert abs(NV_Q1 + NV_Q2 - nv["actual_cum"][0]) < 1e-9, "Q1+Q2 與 recon actual_cum 不符"
NV_MONTHS = nv["months"][0]

# ---- 掃描兩本模型對 TK 名稱的公式引用 ----


def scan_usage(path, snapshot_names, extra_names):
    wb = openpyxl.load_workbook(path)
    blobs = {}
    for ws in wb:
        if ws.title == "TK_In":
            continue
        parts = []
        for row in ws.iter_rows():
            for c in row:
                v = c.value
                if isinstance(v, str) and v.startswith("="):
                    parts.append(v)
        blobs[ws.title] = "\n".join(parts)
    mirror, local = set(), {}
    if "TK_In" in wb.sheetnames:
        for r in wb["TK_In"].iter_rows(values_only=True):
            if r and isinstance(r[0], str) and r[0].startswith("IF_"):
                mirror.add(r[0])
                if len(r) > 2 and r[2]:
                    local.setdefault(r[0], []).append(r[2])
    res = {}
    for nm in list(snapshot_names) + list(extra_names):
        names = local.get(nm, []) + [nm]
        cnt = 0
        for l in names:
            pat = re.compile(r"(?<![A-Za-z0-9_])" + re.escape(l) + r"(?![A-Za-z0-9_])")
            for t in blobs.values():
                cnt += len(pat.findall(t))
        res[nm] = dict(mirror=(nm in mirror), refs=cnt)
    return res


EXTRA_NAMES = ["IF_HoldEcon", "IF_HoldAcct", "IF_RevGWFleet", "IF_RevGW_Sol", "IF_RevGW_Luna", "IF_RevGW_Astra",
               "IF_RevGW_Terra", "IF_OpexGW", "IF_FullCost_Sol", "IF_GPUhrEcon"]
if A.skip_usage_scan:
    USE_W = USE_S = {n: dict(mirror=None, refs=None) for n in list(TK_ITEMS) + EXTRA_NAMES}
else:
    USE_W = scan_usage(P["watt_xlsx"], TK_ITEMS.keys(), EXTRA_NAMES)
    USE_S = scan_usage(P["sil_xlsx"], TK_ITEMS.keys(), EXTRA_NAMES)

# ----------------------------------------------------------------------------------------------
# 樣式
# ----------------------------------------------------------------------------------------------
FN = "Arial"
F_BASE = Font(name=FN, size=10)
F_BOLD = Font(name=FN, size=10, bold=True)
F_IN = Font(name=FN, size=10, color="0000FF")
F_INB = Font(name=FN, size=10, color="0000FF", bold=True)
F_HDR = Font(name=FN, size=10, bold=True, color="FFFFFF")
F_TITLE = Font(name=FN, size=14, bold=True)
F_NOTE = Font(name=FN, size=9, color="595959")
F_WARN = Font(name=FN, size=10, bold=True, color="C00000")
FILL_HDR = PatternFill("solid", fgColor="1F3864")
FILL_SEC = PatternFill("solid", fgColor="D9E1F2")
FILL_OOS = PatternFill("solid", fgColor="FCE4D6")
FILL_BAD = PatternFill("solid", fgColor="FFC7CE")
FILL_KEY = PatternFill("solid", fgColor="FFF2CC")
THIN = Side(style="thin", color="BFBFBF")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
AL_R = Alignment(horizontal="right", vertical="center")
AL_L = Alignment(horizontal="left", vertical="center", wrap_text=True)
AL_LT = Alignment(horizontal="left", vertical="top", wrap_text=True)
AL_C = Alignment(horizontal="center", vertical="center", wrap_text=True)
NF_GW = '#,##0.00'
NF_USD = '#,##0.0'
NF_PCT = '0.0%'
NF_X = '0.000'
NF_INT = '#,##0'
NF_RAW = '#,##0.0000'

wb = openpyxl.Workbook()
ws_sum = wb.active
ws_sum.title = "總表"
ws_act = wb.create_sheet("實績對帳")
ws_cov = wb.create_sheet("介面覆蓋")
ws_sem = wb.create_sheet("欄位語義")
ws_in = wb.create_sheet("輸入")
ws_note = wb.create_sheet("說明")


def put(ws, r, c, v, nf=None, font=None, fill=None, al=None, border=True):
    cell = ws.cell(r, c, v)
    cell.font = font or F_BASE
    if nf:
        cell.number_format = nf
    if fill:
        cell.fill = fill
    if al:
        cell.alignment = al
    elif isinstance(v, (int, float)) or (isinstance(v, str) and v.startswith("=")):
        cell.alignment = AL_R
    else:
        cell.alignment = AL_L
    if border:
        cell.border = BORDER
    return cell


def title(ws, r, text, font=None, merge_to=None, height=None):
    c = ws.cell(r, 1, text)
    c.font = font or F_TITLE
    c.alignment = Alignment(horizontal="left", vertical="top", wrap_text=bool(merge_to))
    if merge_to:
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=merge_to)
    if height:
        ws.row_dimensions[r].height = height


def hdr(ws, r, labels, c0=1):
    for i, t in enumerate(labels):
        put(ws, r, c0 + i, t, font=F_HDR, fill=FILL_HDR, al=AL_C)


def sec(ws, r, text, ncols):
    for c in range(1, ncols + 1):
        put(ws, r, c, text if c == 1 else None, font=F_BOLD, fill=FILL_SEC, al=Alignment(horizontal="left", vertical="center"))


def widths(ws, ws_widths):
    for i, w in enumerate(ws_widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w


def page(ws, landscape=True, fit_w=1, title_rows=None, area=None):
    ws.page_setup.orientation = "landscape" if landscape else "portrait"
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.page_setup.fitToWidth = fit_w
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr = openpyxl.worksheet.properties.PageSetupProperties(fitToPage=True)
    ws.print_options.horizontalCentered = True
    ws.page_margins.left = ws.page_margins.right = 0.4
    ws.page_margins.top = ws.page_margins.bottom = 0.5
    if title_rows:
        ws.print_title_rows = title_rows
    if area:
        ws.print_area = area
    ws.oddFooter.center.text = "&A　第 &P / &N 頁"


# ----------------------------------------------------------------------------------------------
# 「輸入」頁：所有讀入值（藍字）＋來源
# ----------------------------------------------------------------------------------------------
IN = "'輸入'"
YC = {2025: "D", 2026: "E", 2027: "F", 2028: "G", 2029: "H", 2030: "I"}
REG = {}          # key -> row
rowi = [1]


def nxt():
    r = rowi[0]
    rowi[0] += 1
    return r


ws = ws_in
title(ws, nxt(), "輸入：所有讀入的原始值（藍字＝直接取自來源檔；本頁除「比對」欄與核對列外不含公式）")
title(ws, nxt(), f"產生程式 build_summary.py 自動填寫，請勿手改。company-models HEAD {GIT_CM[0][:7]}（{GIT_CM[1]}）；tokenomics HEAD {GIT_TK[0][:7]}（{GIT_TK[1]}）。"
    "SHA 欄＝來源檔所依據之模型檔：「快照記錄」是快照 source.sha256，「實算」是本機 repo 內該模型檔的 SHA-256（前 12 碼；完整值見下方檔案清單）。", font=F_NOTE, merge_to=16, height=30)
nxt()

# --- 檔案清單 ---
sec(ws, nxt(), "0. 檔案清單（完整 SHA-256）", 16)
hr = nxt()
hdr(ws, hr, ["代號", "檔案（相對路徑）／角色", "版本", None, "檔案本身 SHA-256（實算）", None, None, None, None, None, "快照所記錄之模型檔 SHA-256", None, "repo 內模型檔 SHA-256（實算）", None, None, "比對"])
for a, b in ((3, 4), (5, 10), (11, 12), (13, 15)):
    ws.merge_cells(start_row=hr, start_column=a, end_row=hr, end_column=b)
file_rows = {}
F_SHA = Font(name=FN, size=8, color="0000FF")
LIST = [
    ("TKS", P["tk_snapshot"], "快照（下游實際讀的 23 個名稱）"),
    ("TKX", P["tk_xlsx"], "Tokenomics 模型檔（具名範圍直讀）"),
    ("WAS", P["watt_snapshot"], "Wattonomics 快照"),
    ("WAX", P["watt_xlsx"], "Wattonomics 模型檔（讀 Players 世代組合；公式引用掃描）"),
    ("SIS", P["sil_snapshot"], "Siliconomics 快照"),
    ("SIX", P["sil_xlsx"], "Siliconomics 模型檔（公式引用掃描）"),
    ("SIB", P["status"], "Siliconomics 逐年狀態檔"),
    ("SIR", P["recon"], "Siliconomics 對帳快照（13 點）"),
    ("BRG", P["bridge"], "協調橋接表"),
]
for code, path, role in LIST:
    r = nxt()
    file_rows[code] = r
    rel = os.path.relpath(path, CM) if path.startswith(CM) else (os.path.relpath(path, TKR) if path.startswith(TKR) else path)
    ver = SV if code == "SIX" else FILES.get(code, {}).get("ver", "")
    own = sha256(path)
    put(ws, r, 1, code, font=F_IN)
    put(ws, r, 2, rel + "\n" + role, font=F_IN)
    ws.row_dimensions[r].height = 40
    put(ws, r, 3, ver, font=F_IN)
    ws.merge_cells(start_row=r, start_column=3, end_row=r, end_column=4)
    put(ws, r, 5, own, font=F_SHA, al=AL_L)
    ws.merge_cells(start_row=r, start_column=5, end_row=r, end_column=10)
    rec = FILES.get(code, {}).get("rec")
    put(ws, r, 11, rec if rec else "（檔內未記錄）", font=F_SHA, al=AL_L)
    ws.merge_cells(start_row=r, start_column=11, end_row=r, end_column=12)
    if code in ("TKS", "WAS", "SIS"):
        modelfile = {"TKS": "tk_xlsx", "WAS": "watt_xlsx", "SIS": "sil_xlsx"}[code]
        put(ws, r, 13, SHA[modelfile], font=F_SHA, al=AL_L)
        ws.merge_cells(start_row=r, start_column=13, end_row=r, end_column=15)
        if rec:
            put(ws, r, 16, f'=IF(K{r}=M{r},"一致","不一致")', font=F_BOLD, al=AL_C)
        else:
            put(ws, r, 16, "快照未記錄", al=AL_C)
    else:
        put(ws, r, 13, "—", al=AL_C)
        ws.merge_cells(start_row=r, start_column=13, end_row=r, end_column=15)
        put(ws, r, 16, "—", al=AL_C)
nxt()
FILES_LAST = rowi[0] - 1


def prov(code):
    f = FILES[code]
    rec = f["rec"][:12] if f["rec"] else "—"
    act = f["act"][:12] if f["act"] else "—"
    return f["file"], f["ver"], rec, act


def prov_cells(ws, r, code, name_text, tag):
    file, ver, rec, act = prov(code)
    put(ws, r, 10, tag, font=F_IN, al=AL_C)
    put(ws, r, 11, file, font=F_IN)
    put(ws, r, 12, name_text, font=F_IN)
    put(ws, r, 13, ver, font=F_IN)
    put(ws, r, 14, rec, font=F_IN, al=AL_C)
    put(ws, r, 15, act, font=F_IN, al=AL_C)
    put(ws, r, 16, f'=IF(N{r}="—","未記錄",IF(N{r}=O{r},"一致","不一致"))', font=F_BOLD, al=AL_C)


IN_HDR_SERIES = ["鍵", "項目", "單位", "2025", "2026", "2027", "2028", "2029", "2030", "等級", "來源檔", "名稱／位置", "版本",
                 "SHA 快照記錄", "SHA 實算", "比對"]


def series(key, label, unit, vals, nf, code, name_text, tag):
    """vals: {year: value|None}"""
    r = nxt()
    REG[key] = r
    put(ws_in, r, 1, key, font=F_IN)
    put(ws_in, r, 2, label, font=F_IN)
    put(ws_in, r, 3, unit, font=F_IN)
    for y in YEARS + [2030]:
        v = vals.get(y)
        c = openpyxl.utils.column_index_from_string(YC[y])
        if v is None:
            put(ws_in, r, c, "—", font=F_IN, al=AL_R)
        else:
            put(ws_in, r, c, v, nf=nf, font=F_IN, al=(AL_R if isinstance(v, (int, float)) else AL_C))
    prov_cells(ws_in, r, code, name_text, tag)
    return r


def IR(key, y):
    return f"{IN}!${YC[y]}${REG[key]}"


# --- 1. Tokenomics 具名範圍（xlsx）---
sec(ws, nxt(), f"1. Tokenomics 每 GW 價目（直讀 {TV} 工作簿具名範圍；H 欄為快照基準值對照）", 16)
hdr(ws, nxt(), ["鍵", "名稱／世代", "單位", "低成本", "基準", "高成本", "快照基準值", "與快照差", "—", "等級", "來源檔", "名稱／位置", "版本", "SHA 快照記錄", "SHA 實算", "比對"])
TKROW = {}
TK_TAG = {"IF_CapexIT": "Derived（Rubin Ultra 含 Assumed 機架價倍數）", "IF_CapexTotal": "Derived（同上）", "IF_CapexFacility": "Assumed",
          "IF_GPUsPerGW": "Derived", "IF_HoldEcon": "Derived（含 Assumed 折現）", "IF_HoldAcct": "Derived", "IF_RevGWFleet": "Derived（理想上限，Assumed 單價）"}
for nm in TK_NAMES_XLSX:
    d = TKX[nm]
    for gi, gen in enumerate(GENS):
        r = nxt()
        TKROW[(nm, gi)] = r
        put(ws, r, 1, f"{nm}|{GEN_SHORT[gi]}", font=F_IN)
        put(ws, r, 2, f"{nm}（{d['label'][:18]}）— {gen[:12]}", font=F_IN)
        put(ws, r, 3, d["unit"], font=F_IN)
        for si in range(3):
            v = d["vals"][gi * 3 + si]
            put(ws, r, 4 + si, v, nf=(NF_INT if nm == "IF_GPUsPerGW" else NF_RAW), font=F_IN)
        sb = snap_base(nm, gen)
        if sb is not None:
            put(ws, r, 7, sb, nf=(NF_INT if nm == "IF_GPUsPerGW" else NF_RAW), font=F_IN)
            put(ws, r, 8, f"=E{r}-G{r}", nf=NF_RAW)
        else:
            put(ws, r, 7, "（不在快照）", font=F_IN, al=AL_C)
            put(ws, r, 8, "—", al=AL_C)
        put(ws, r, 9, "—", al=AL_C)
        prov_cells(ws_in, r, "TKX", d["ref"], TK_TAG[nm])
TK_SNAP_DIFF_ROWS = (min(TKROW.values()), max(TKROW.values()))
nxt()

# --- 2. 世代組合 ---
sec(ws, nxt(), "2. Wattonomics 新增 GW 世代組合（模型 Players!D5:M9，快照只有空白；NVIDIA 內部占比）", 16)
hdr(ws, nxt(), IN_HDR_SERIES)
MIXROW = {}
for gi, k in enumerate(MIX_ORDER):
    MIXROW[gi] = series(f"mix|{GEN_SHORT[gi]}", f"新增 GW 世代占比：{GEN_SHORT[gi]}", "比例",
                        {y: MIX[k][y] for y in YEARS}, NF_PCT, "WAX", f"Players!{get_column_letter(4)}{5 + gi}:M{5 + gi}（{k}）",
                        "Assumed／Derived（2025 Analogy；2026–29 由 Siliconomics 顆數推導）")
    # 2030 欄：模型沿用 2029，這裡不取
r = nxt()
REG["mix|sum"] = r
put(ws, r, 1, "mix|sum", font=F_IN)
put(ws, r, 2, "核對：世代占比合計（須＝1）", font=F_IN)
put(ws, r, 3, "比例", font=F_IN)
for y in YEARS:
    c = openpyxl.utils.column_index_from_string(YC[y])
    put(ws, r, c, f"=SUM({YC[y]}{MIXROW[0]}:{YC[y]}{MIXROW[4]})", nf=NF_PCT)
nxt()

# --- 3. Wattonomics 快照 ---
sec(ws, nxt(), f"3. Wattonomics 快照 {WV}（上線年；單位 IT GW 或 US$B）", 16)
hdr(ws, nxt(), IN_HDR_SERIES)
WAT_TAG = "Derived"


def wser(key, label, unit, nf, tag=WAT_TAG, name=None):
    nm = name or key
    return series(key, label, unit, {y: wv(nm, y) for y in YEARS + [2030]}, nf, "WAS", f"{nm}_<年>", tag)


wser("WIF_MktGW_Add", "年新增上線 IT GW（毛額）", "GW", NF_GW)
wser("WIF_MktGW_YE", "年底上線 IT GW", "GW", NF_GW)
wser("WIF_MktGW_Avg", "年平均上線 IT GW", "GW", NF_GW)
wser("WIF_Spend_it_sum", "IT 設備支出合計（上線年）", "US$B", NF_USD)
for it in ["compute", "foundry", "memory", "network", "power_thermal", "server", "storage", "substrate"]:
    wser(f"WIF_Spend_it_{it}", f"　IT 分項：{it}", "US$B", NF_USD)
wser("WIF_Spend_fac_sum", "廠房支出合計", "US$B", NF_USD)
wser("WIF_Spend_self_gen", "自備發電支出", "US$B", NF_USD)
wser("WIF_Spend_total", "總支出（IT＋廠房＋自備發電）", "US$B", NF_USD)
wser("WIF_Gap_Short_A_D0", "資金缺口（短缺；舊名 A_D0）", "US$B", NF_USD)
wser("WIF_Coverage", "覆蓋率（有證據支撐的年新增 ÷ 主路徑年新增）", "比例", NF_PCT)
wser("WIF_Evidence_Lo", "證據比例（低）", "比例", NF_PCT)
wser("WIF_Evidence", "證據比例（基準）", "比例", NF_PCT)
wser("WIF_Evidence_Hi", "證據比例（高）", "比例", NF_PCT)
for rn, lab, unit, nf in [("ActualRev", "NVIDIA 資料中心營收：公司值（2027 含指引）", "US$B", NF_USD), ("ImpliedRev", "NVIDIA 資料中心營收：Wattonomics 隱含值", "US$B", NF_USD),
                          ("NVGW", "NVIDIA 對應 GW（對帳用）", "GW", NF_GW), ("ShipGW", "NVIDIA 出貨 GW（對帳用）", "GW", NF_GW),
                          ("ContentPerGW", "每 GW NVIDIA 內容", "US$B/GW", NF_USD), ("Ratio", "隱含 ÷ 公司值", "比", NF_X)]:
    wser(f"WIF_Recon_NV_{rn}", lab, unit, nf)
r_chk = nxt()
REG["w_chk_it"] = r_chk
put(ws, r_chk, 1, "w_chk_it", font=F_IN)
put(ws, r_chk, 2, "核對：IT 八項分項合計 − WIF_Spend_it_sum（應為 0）", font=F_IN)
put(ws, r_chk, 3, "US$B", font=F_IN)
first_it = REG["WIF_Spend_it_compute"]
last_it = REG["WIF_Spend_it_substrate"]
for y in YEARS:
    c = openpyxl.utils.column_index_from_string(YC[y])
    put(ws, r_chk, c, f"=SUM({YC[y]}{first_it}:{YC[y]}{last_it})-{YC[y]}{REG['WIF_Spend_it_sum']}", nf='0.000')
r_chk2 = nxt()
REG["w_chk_tot"] = r_chk2
put(ws, r_chk2, 1, "w_chk_tot", font=F_IN)
put(ws, r_chk2, 2, "核對：IT＋廠房＋自備發電 − WIF_Spend_total（應為 0）", font=F_IN)
put(ws, r_chk2, 3, "US$B", font=F_IN)
for y in YEARS:
    c = openpyxl.utils.column_index_from_string(YC[y])
    put(ws, r_chk2, c, f"={YC[y]}{REG['WIF_Spend_it_sum']}+{YC[y]}{REG['WIF_Spend_fac_sum']}+{YC[y]}{REG['WIF_Spend_self_gen']}-{YC[y]}{REG['WIF_Spend_total']}", nf='0.000')
nxt()

# --- 4. Siliconomics 快照 ---
sec(ws, nxt(), f"4. Siliconomics 快照 {SV}（出貨年；GW＝成本等值 GW；金額 US$B）", 16)
hdr(ws, nxt(), IN_HDR_SERIES)


def sser_arr(key, label, unit, nf, tag="Derived"):
    return series(key, label, unit, sarr(key), nf, "SIS", f"{key}（陣列；長度 {len(SN[key])}，對齊至 2029）", tag)


def sser_yr(key, label, unit, nf, tag="Derived"):
    return series(key, label, unit, {y: sv(key, y) for y in YEARS}, nf, "SIS", f"{key}_<年>", tag)


sser_arr("SIF_DeployGW_Base", "部署 GW 基準（依規則）", "GW", NF_GW)
sser_yr("SIF_DeployGW_Supply", "供給側讀數：原始顆數（不乘係數）", "GW", NF_GW)
sser_yr("SIF_DeployGW_Demand", "需求側讀數：Wattonomics 換算值（平滑法）", "GW", NF_GW)
sser_yr("SIF_DeployGW_Alt", "另一側讀數（未被採為基準的一條）", "GW", NF_GW)
sser_yr("SIF_DeployGW_LegacyS33", "舊 S33 取法（原始 × 舊 S33 係數；v0.18 取法；v0.21 起取代 Current 別名）", "GW", NF_GW)
sser_yr("SIF_DeployGW_AltRatio", "另一側讀數 ÷ 基準", "x", NF_X)
sser_arr("SIF_DeployGWPhys_Base", "實體部署 GW（出貨年）", "GW", NF_GW)
sser_arr("SIF_AIITCapex_Base", "AI IT 資本支出（基準）", "US$B", NF_USD)
sser_arr("SIF_UnitsCoef_Base", "顆數係數（S33）", "x", NF_X, "Assumed")
MKT_NAMES = [("SIF_TestSoCMkt_Base", "SoC 測試機市場"), ("SIF_GpuAsicMkt_Base", "GPU／ASIC 晶片市場"), ("SIF_HBM_Mkt_Base", "HBM 市場"),
             ("SIF_AsmMkt_Base", "AI 伺服器組裝市場"), ("SIF_PwrMkt_Base", "AI 電源市場"), ("SIF_LiqMkt_Base", "AI 液冷市場"),
             ("SIF_PCBMkt_Base", "PCB 市場"), ("SIF_CCLMkt_Base", "CCL 市場"), ("SIF_GPUSysMkt_Base", "GPU 系統市場")]
for k, lab in MKT_NAMES:
    sser_arr(k, lab, "US$B", NF_USD)
nxt()

# --- 5. 逐年狀態 ---
sec(ws, nxt(), "5. Siliconomics 逐年狀態檔 baseline_status.json（2026–2029；規則 v1.4）", 16)
hdr(ws, nxt(), IN_HDR_SERIES)
ST = status["years"]


def stser(key, label, unit, fn, nf=None, tag="Derived", name=None):
    vals = {}
    for y in YEARS:
        if str(y) in ST:
            v = fn(ST[str(y)])
            vals[y] = v if v is not None else "—"
        else:
            vals[y] = None
    return series(key, label, unit, vals, nf, "SIB", name or key, tag)


stser("st_status", "狀態（甲／乙／丙／待定）", "文字", lambda d: d["status"], tag="Derived", name="years.<年>.status")
stser("st_label", "狀態標示", "文字", lambda d: d["label"], name="years.<年>.label")
stser("st_deadline", "裁定期限（對帳中）", "日期", lambda d: d.get("deadline"), name="years.<年>.deadline")
stser("st_c0_deadline", "C0 實績對帳期限", "日期", lambda d: d.get("c0_deadline"), name="years.<年>.c0_deadline")
stser("st_c0_fail", "C0 公司實績對帳是否超出 ±15%（1＝是）", "旗標", lambda d: (1 if d.get("c0_fail") else 0), NF_INT, name="years.<年>.c0_fail")
stser("st_use_conv", "是否改用換算值（1＝是）", "旗標", lambda d: (1 if d.get("use_converted") else 0), NF_INT, name="years.<年>.use_converted")
stser("st_continuity", "連續法（1＝是）", "旗標", lambda d: (1 if d.get("continuity") else 0), NF_INT, name="years.<年>.continuity")
stser("st_converted", "換算值（Demand）", "GW", lambda d: d["converted"], NF_GW, name="years.<年>.converted")
stser("st_supply", "供給側讀數", "GW", lambda d: d["supply_side"], NF_GW, name="years.<年>.supply_side")
stser("st_current", "舊 S33 取法（legacy_s33）", "GW", lambda d: d["legacy_s33"], NF_GW, name="years.<年>.legacy_s33")
stser("st_silbase", "Siliconomics 基準", "GW", lambda d: d["sil_base"], NF_GW, name="years.<年>.sil_base")
for lvl in ("lo", "base", "hi"):
    stser(f"st_ws_{lvl}", f"Wattonomics 證據比例（{ {'lo': '低', 'base': '基準', 'hi': '高'}[lvl] }）", "比例",
          lambda d, l=lvl: d["wattonomics"]["evidence_share"][l], NF_PCT, name=f"years.<年>.wattonomics.evidence_share.{lvl}")
for lvl in ("lo", "base", "hi"):
    stser(f"st_ss_{lvl}", f"Siliconomics 證據比例（{ {'lo': '低', 'base': '基準', 'hi': '高'}[lvl] }）", "比例",
          lambda d, l=lvl: d["siliconomics"]["evidence_share"][l], NF_PCT, "Assumed", name=f"years.<年>.siliconomics.evidence_share.{lvl}")
stser("st_c0_point", "C0 對帳點說明", "文字", lambda d: d.get("c0_point"), name="years.<年>.c0_point")
r = nxt()
REG["st_threshold"] = r
put(ws, r, 1, "st_threshold", font=F_IN)
put(ws, r, 2, "「有證據」門檻（規則 C3，首次判定）", font=F_IN)
put(ws, r, 3, "比例", font=F_IN)
put(ws, r, 5, status["threshold"], nf=NF_PCT, font=F_IN)
prov_cells(ws, r, "SIB", "threshold", "Assumed")
r = nxt()
REG["tol"] = r
put(ws, r, 1, "tol", font=F_IN)
put(ws, r, 2, "層間殘差容差（±）", font=F_IN)
put(ws, r, 3, "比例", font=F_IN)
put(ws, r, 5, TOL, nf=NF_PCT, font=F_IN)
prov_cells(ws, r, "SIR", "tolerance", "Assumed")
nxt()

# --- 6. 橋接表 ---
sec(ws, nxt(), f"6. 協調橋接表 {BV}（In 頁參數；Bridge 頁用於交叉核對，不用來算總表）", 16)
hdr(ws, nxt(), ["鍵", "項目", "單位", "低", "基準", "高", "—", "—", "—", "等級", "來源檔", "名稱／位置", "版本", "SHA 快照記錄", "SHA 實算", "比對"])
for k, lab, unit, tag in [("lag_on", "付款（出貨）到上線延遲", "月", "Assumed"), ("inv_on", "未上線庫存 ÷ 年出貨", "比", "Assumed")]:
    r = nxt()
    REG["br_" + k] = r
    put(ws, r, 1, "br_" + k, font=F_IN)
    put(ws, r, 2, lab, font=F_IN)
    put(ws, r, 3, unit, font=F_IN)
    for i, lv in enumerate(("low", "base", "high")):
        put(ws, r, 4 + i, BRIN[k][lv], nf=NF_GW if k == "lag_on" else NF_PCT, font=F_IN)
    for c in (7, 8, 9):
        put(ws, r, c, "—", al=AL_C)
    prov_cells(ws, r, "BRG", f"In!{k}", tag)
r = nxt()
put(ws, r, 1, "br_in_note", font=F_IN)
put(ws, r, 2, f"橋接表 In 頁 A1 原文：{BR_IN_NOTE}", font=F_IN)
ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=16)
nxt()
hdr(ws, nxt(), IN_HDR_SERIES)
BRROWS = [("g", "總差距（現行基準 − 上線直接換算）", "GW"), ("g2", "② 時差與庫存", "GW"), ("g3", "③ 殘差（分歧）", "GW"), ("g3p", "③ ÷ 現行基準", "比例"),
          ("g2p", "② 累計 ÷ 當年基準", "比例"), ("x0", "上線 GW 直接換成本等值（不做時差）", "GW"), ("x2", "換算值（平滑法；上游定基準時之基準）", "GW"),
          ("s2", "平滑法實體出貨", "GW"), ("rt", "本庫 實體 ÷ 成本等值", "x"), ("s0", f"Siliconomics 現行基準（橋接表內，{BSV} 輸入）", "GW"), ("k", "k＝lag_on÷12＋inv_on", "月")]
for k, lab, unit in BRROWS:
    nf = NF_PCT if unit == "比例" else NF_X if unit == "x" else NF_GW
    series("br_" + k, lab, unit, {y: BR[k].get(y) for y in YEARS if BR[k].get(y) is not None}, nf, "BRG", f"Bridge!{k}", "Derived")
nxt()

# --- 7. 對帳點 ---
sec(ws, nxt(), f"7. 實績與對帳點（{os.path.basename(P['recon'])}；NVIDIA 實績取到 {nv['actual_label']}）", 16)
hdr(ws, nxt(), IN_HDR_SERIES)
series("nv_sil_implied", "NVIDIA 資料中心營收：Siliconomics 隱含值（R4 橋接）", "US$B", {y: v for y, v in zip(nv["years"], nv["implied"]) if v is not None}, NF_USD, "SIR", "markets.nv_dc.implied", "Derived")
series("nv_company", "NVIDIA 資料中心營收：公司值（含指引，只列參考）", "US$B", {y: v for y, v in zip(nv["years"], nv["company"]) if v is not None}, NF_USD, "SIR", "markets.nv_dc.company", "Interested-party")
series("nv_ratio_main_rec", "對帳檔記錄：主讀數（H1 實績 ÷ 層級同期間）", "比", {2026: nv["ratio_main"][0]}, NF_X, "SIR", "markets.nv_dc.ratio_main", "Derived")
series("nv_ratio_second_rec", "對帳檔記錄：第二讀數（Q2×4 ÷ 層級全年）", "比", {2026: nv["ratio_second"][0]}, NF_X, "SIR", "markets.nv_dc.ratio_second", "Derived")
r = nxt()
REG["nv_q1"] = r
put(ws, r, 1, "nv_q1", font=F_IN)
put(ws, r, 2, "NVIDIA FY2027 Q1 資料中心營收（8-K）", font=F_IN)
put(ws, r, 3, "US$B", font=F_IN)
put(ws, r, 5, NV_Q1, nf=NF_USD, font=F_IN)
prov_cells(ws, r, "SIR", "markets.nv_dc.ref_actual（文字解析）", "Verified（依對帳檔引用 8-K；本建置未重核原文）")
r = nxt()
REG["nv_q2"] = r
put(ws, r, 1, "nv_q2", font=F_IN)
put(ws, r, 2, "NVIDIA FY2027 Q2 資料中心營收（8-K）", font=F_IN)
put(ws, r, 3, "US$B", font=F_IN)
put(ws, r, 5, NV_Q2, nf=NF_USD, font=F_IN)
prov_cells(ws, r, "SIR", "markets.nv_dc.ref_actual（文字解析）", "Verified（依對帳檔引用 8-K；本建置未重核原文）")
r = nxt()
REG["nv_months"] = r
put(ws, r, 1, "nv_months", font=F_IN)
put(ws, r, 2, "已實績月數（FY2027 上半年）", font=F_IN)
put(ws, r, 3, "月", font=F_IN)
put(ws, r, 5, NV_MONTHS, nf=NF_INT, font=F_IN)
prov_cells(ws, r, "SIR", "markets.nv_dc.months", "Derived")
r = nxt()
REG["nv_cum_rec"] = r
put(ws, r, 1, "nv_cum_rec", font=F_IN)
put(ws, r, 2, "對帳檔記錄之實績累計（Q1＋Q2）", font=F_IN)
put(ws, r, 3, "US$B", font=F_IN)
put(ws, r, 5, nv["actual_cum"][0], nf=NF_USD, font=F_IN)
prov_cells(ws, r, "SIR", "markets.nv_dc.actual_cum", "Verified（同上）")
r = nxt()
REG["nv_oos_q3"] = r
for c in range(1, 17):
    ws.cell(r, c).fill = FILL_OOS
put(ws, r, 1, "nv_oos_q3", font=F_INB, fill=FILL_OOS)
put(ws, r, 2, "【樣本外檢驗（未入模型）】NVIDIA FY2027 Q3 資料中心營收", font=F_INB, fill=FILL_OOS)
put(ws, r, 3, "US$B", font=F_IN, fill=FILL_OOS)
put(ws, r, 5, A.oos_q3, nf=NF_USD, font=F_INB, fill=FILL_OOS)
put(ws, r, 10, "Interested-party（公司公布數；未核對原文）", font=F_IN, fill=FILL_OOS, al=AL_C)
put(ws, r, 11, "（外部）", font=F_IN, fill=FILL_OOS)
put(ws, r, 12, A.oos_source, font=F_IN, fill=FILL_OOS)
put(ws, r, 13, "—", font=F_IN, fill=FILL_OOS, al=AL_C)
put(ws, r, 14, "—", font=F_IN, fill=FILL_OOS, al=AL_C)
put(ws, r, 15, "—", font=F_IN, fill=FILL_OOS, al=AL_C)
put(ws, r, 16, "只供實績對帳頁樣本外區塊", font=F_WARN, fill=FILL_OOS, al=AL_C)
nxt()
sec(ws, nxt(), "7b. 13 個對帳點明細見「介面覆蓋」頁（直接讀 recon 快照，該頁內藍字）", 16)

widths(ws_in, [20, 52, 9, 12, 12, 12, 12, 12, 12, 22, 34, 34, 22, 17, 17, 11])
ws_in.freeze_panes = "D4"
page(ws_in, True, 1)
# 不一致標紅（條件格式）
ws_in.conditional_formatting.add("P4:P400", FormulaRule(formula=['P4="不一致"'], fill=FILL_BAD, font=Font(name=FN, bold=True, color="9C0006")))
ws_in.conditional_formatting.add("P4:P400", FormulaRule(formula=['P4="一致"'], fill=PatternFill("solid", bgColor="C6EFCE"), font=Font(name=FN, bold=True, color="006100")))

# ----------------------------------------------------------------------------------------------
# 「總表」
# ----------------------------------------------------------------------------------------------
ws = ws_sum
SC = {2025: "D", 2026: "E", 2027: "F", 2028: "G", 2029: "H"}
Y26 = [2026, 2027, 2028, 2029]
rows = []   # list of (key, kind, label, unit, how, tag, src, fn, nf)


def add(key, label, unit, how, tag, src, fn, nf, years=YEARS, bold=False):
    rows.append(dict(kind="row", key=key, label=label, unit=unit, how=how, tag=tag, src=src, fn=fn, nf=nf, years=years, bold=bold))


def addsec(text):
    rows.append(dict(kind="sec", text=text))


mixw = lambda y: f"{IN}!${YC[y]}${MIXROW[0]}:${YC[y]}${MIXROW[4]}"
tkcol = lambda nm, col: f"{IN}!${col}${TKROW[(nm, 0)]}:${col}${TKROW[(nm, 4)]}"
S = lambda key, y: f"{SC[y]}{ROW[key]}"
ROW = {}

# 一眼看
addsec("一眼看：三層現在對到哪裡（每格都連到下方各列；色塊＝狀態）")
add("v_judge", "兩層殘差是否在容差內（±15%）", "判定", "殘差 ÷ Siliconomics 基準；容差取 recon 快照 B", "Derived", "recon.tolerance",
    lambda y: f"={S('g_judge', y)}", None, Y26, True)
add("v_rule", "協調狀態（規則 v1.4）", "狀態", "baseline_status 逐年標示；甲＝有證據、乙＝一層有證據、丙＝連續法、待定＝對帳中", "Derived", "baseline_status.years.<年>.label",
    lambda y: f"={S('r_label', y)}", None, Y26, True)
add("v_c0", "NVIDIA 實績檢驗（C0）", "狀態", "公司實績對層級同期間是否超出 ±15%", "Derived", "baseline_status.years.<年>.c0_fail",
    lambda y: f"={S('r_c0', y)}", None, Y26, True)
add("v_gap", "兩層差距（Siliconomics 基準 − Wattonomics 直接換算）", "GW", "見下方「兩層差距」", "Derived", "（連結）",
    lambda y: f"={S('gap_gw', y)}", NF_GW, Y26, True)
add("v_pct", "　差距 ÷ Siliconomics 基準", "%", "", "Derived", "（連結）", lambda y: f"={S('gap_pct', y)}", NF_PCT, Y26)
add("v_dl", "裁定期限", "日期", "規則期限；2026 為 C0 實績對帳期限", "Derived", "baseline_status", lambda y: f"={S('r_dl', y)}", None, Y26, True)

# Tokenomics
addsec("1. Tokenomics：每 GW 單位值（依當年新增 GW 的 NVIDIA 世代組合加權；GW＝IT 關鍵電力）")
add("tk_mixtxt", "當年世代組合（NVIDIA 內部占比）", "文字",
    "取法：占比取自 Wattonomics Players 頁（2025 為 CoreWeave 組合類比；2026 以 GB300／VR200；2027 起 VR200／Rubin Ultra），見輸入頁第 2 節", "Assumed／Derived",
    "Players!mix_*",
    lambda y: "=TRIM(" + "&".join(
        f'IF({IN}!{YC[y]}{MIXROW[g]}>0,"{GEN_SHORT[g]} "&TEXT({IN}!{YC[y]}{MIXROW[g]},"0%")&" ","")' for g in range(5)) + ")", None)
add("tk_it_lo", "每 GW IT 設備資本支出（低成本）", "US$B/GW", "世代占比 × IF_CapexIT 低成本欄（SUMPRODUCT）", "Derived", "IF_CapexIT（低）",
    lambda y: f"=SUMPRODUCT({mixw(y)},{tkcol('IF_CapexIT', 'D')})", NF_USD)
add("tk_it", "每 GW IT 設備資本支出（基準）", "US$B/GW", "世代占比 × IF_CapexIT 基準欄；與 Wattonomics CapexGW 第 7 列「NVIDIA 每 IT GW IT capex（世代加權）」同值", "Derived", "IF_CapexIT（基準）",
    lambda y: f"=SUMPRODUCT({mixw(y)},{tkcol('IF_CapexIT', 'E')})", NF_USD, YEARS, True)
add("tk_it_hi", "每 GW IT 設備資本支出（高成本）", "US$B/GW", "世代占比 × IF_CapexIT 高成本欄", "Derived", "IF_CapexIT（高）",
    lambda y: f"=SUMPRODUCT({mixw(y)},{tkcol('IF_CapexIT', 'F')})", NF_USD)
add("tk_tot", "每 GW 資本支出合計（IT＋廠房，基準）", "US$B/GW", "世代占比 × IF_CapexTotal 基準欄", "Derived", "IF_CapexTotal（基準）",
    lambda y: f"=SUMPRODUCT({mixw(y)},{tkcol('IF_CapexTotal', 'E')})", NF_USD)
add("tk_gpu", "每 GW GPU 數（基準）", "顆/GW", "世代占比 × IF_GPUsPerGW 基準欄", "Derived", "IF_GPUsPerGW（基準）",
    lambda y: f"=SUMPRODUCT({mixw(y)},{tkcol('IF_GPUsPerGW', 'E')})", NF_INT)
add("tk_hold", "每 GW 年持有成本（經濟口徑，基準）　※兩層均未使用", "US$B/年/GW", "世代占比 × IF_HoldEcon 基準欄；此介面不在下傳的 23 個名稱內", "Derived", "IF_HoldEcon（基準；工作簿直讀）",
    lambda y: f"=SUMPRODUCT({mixw(y)},{tkcol('IF_HoldEcon', 'E')})", NF_USD)

# Wattonomics
addsec("2. Wattonomics：全市場 GW 與支出（上線年；GW＝IT GW）")
add("w_add", "新增上線 GW（毛額）", "GW", "WIF_MktGW_Add", "Derived", "WIF_MktGW_Add", lambda y: f"={IR('WIF_MktGW_Add', y)}", NF_GW, YEARS, True)
add("w_avg", "平均上線 GW", "GW", "WIF_MktGW_Avg", "Derived", "WIF_MktGW_Avg", lambda y: f"={IR('WIF_MktGW_Avg', y)}", NF_GW)
add("w_ye", "年底上線 GW", "GW", "WIF_MktGW_YE", "Derived", "WIF_MktGW_YE", lambda y: f"={IR('WIF_MktGW_YE', y)}", NF_GW)
add("w_it", "IT 設備支出", "US$B", "WIF_Spend_it_sum（八項分項加總，核對見輸入頁）", "Derived", "WIF_Spend_it_sum", lambda y: f"={IR('WIF_Spend_it_sum', y)}", NF_USD, YEARS, True)
add("w_tot", "總支出（IT＋廠房＋自備發電）", "US$B", "WIF_Spend_total", "Derived", "WIF_Spend_total", lambda y: f"={IR('WIF_Spend_total', y)}", NF_USD)
add("w_impl", "　隱含每 GW IT 支出（IT 支出 ÷ 新增 GW）", "US$B/GW", "僅對照 Tokenomics 加權值；含非 NVIDIA、較便宜平台，故低於 NVIDIA 加權值", "Derived", "（計算）",
    lambda y: f"={S('w_it', y)}/{S('w_add', y)}", NF_USD)
add("w_short", "資金缺口（參考；非對帳項）", "US$B", "WIF_Gap_Short_A_D0（正＝短缺）", "Derived", "WIF_Gap_Short_A_D0", lambda y: f"={IR('WIF_Gap_Short_A_D0', y)}", NF_USD)

# 換算
addsec("3. Wattonomics 換算到出貨年（規則 v1.1 平滑法：出貨ₜ ≈ 上線ₜ ＋ k ×（上線ₜ₊₁ − 上線ₜ）；再 ÷ 實體／成本等值比）")
add("k", "k＝lag_on ÷ 12 ＋ inv_on（基準）", "年", "lag_on 4.5 月、inv_on 10%；皆 Assumed（低／高見輸入頁），Wattonomics 為權威層", "Assumed", "橋接表 In!lag_on、inv_on",
    lambda y: f"={IN}!$E${REG['br_lag_on']}/12+{IN}!$E${REG['br_inv_on']}", '0.000')
add("w_s2", "換算到出貨年：實體 GW", "GW（實體）", "平滑法；2029 用 Wattonomics 2030 值", "Derived", "（計算）",
    lambda y: f"={S('w_add', y)}+{S('k', y)}*({IR('WIF_MktGW_Add', y + 1)}-{S('w_add', y)})", NF_GW)
add("rt", "實體 GW ÷ 成本等值 GW（Siliconomics）", "x", "SIF_DeployGWPhys_Base ÷ SIF_DeployGW_Base", "Derived", "SIF_DeployGWPhys_Base、SIF_DeployGW_Base",
    lambda y: f"={IR('SIF_DeployGWPhys_Base', y)}/{IR('SIF_DeployGW_Base', y)}", NF_X, Y26)
add("w_x2", "換算到出貨年：成本等值 GW（含時差，＝需求側讀數）", "GW", "實體 GW ÷ 比值；＝SIF_DeployGW_Demand", "Derived", "（計算；對照 SIF_DeployGW_Demand）",
    lambda y: f"={S('w_s2', y)}/{S('rt', y)}", NF_GW, Y26, True)
add("w_x0", "上線 GW 直接換成本等值（不做時差）", "GW", "新增上線 GW ÷ 比值；橋接表「總差距」的比較基準", "Derived", "（計算）",
    lambda y: f"={S('w_add', y)}/{S('rt', y)}", NF_GW, Y26)

# Siliconomics
addsec("4. Siliconomics：供給與資本支出（出貨年；GW＝成本等值 GW）")
add("s_sup", "供給側原始讀數（顆數，不乘係數）", "GW", "SIF_DeployGW_Supply", "Derived", "SIF_DeployGW_Supply_<年>", lambda y: f"={IR('SIF_DeployGW_Supply', y)}", NF_GW, Y26)
add("s_base", "Siliconomics 基準（依規則）", "GW", "SIF_DeployGW_Base；2029 為連續法（前一年基準 × 換算值成長率）", "Derived", "SIF_DeployGW_Base", lambda y: f"={IR('SIF_DeployGW_Base', y)}", NF_GW, Y26, True)
add("s_coef", "　基準 ÷ 供給側原始讀數（隱含係數）", "x", "對照顆數係數 SIF_UnitsCoef_Base（" + "／".join(f"{x:.2f}" for x in SN["SIF_UnitsCoef_Base"][-4:]) + "；2028–29 為 Assumed）", "Derived", "（計算）",
    lambda y: f"={S('s_base', y)}/{S('s_sup', y)}", NF_X, Y26)
add("s_it", "AI IT 資本支出（基準）", "US$B", "SIF_AIITCapex_Base（基準 GW × 每 GW IT 單位值）", "Derived", "SIF_AIITCapex_Base", lambda y: f"={IR('SIF_AIITCapex_Base', y)}", NF_USD, Y26, True)

# 差距
addsec("5. 兩層差距（Siliconomics 基準 − Wattonomics 直接換算；與橋接表「總差距」同定義）")
add("gap_gw", "差距", "GW", "基準 − 上線直接換算", "Derived", "（計算）", lambda y: f"={S('s_base', y)}-{S('w_x0', y)}", NF_GW, Y26, True)
add("gap_pct", "差距 ÷ Siliconomics 基準", "%", "", "Derived", "（計算）", lambda y: f"={S('gap_gw', y)}/{S('s_base', y)}", NF_PCT, Y26)
add("gap_x", "Siliconomics 基準 ÷ Wattonomics 直接換算", "倍", "", "Derived", "（計算）", lambda y: f"={S('s_base', y)}/{S('w_x0', y)}", '0.00"x"', Y26)
add("gap_usd_gw", "差距折算金額（差距 GW × 當年 Tokenomics 每 GW IT 單位值）", "US$B", "口徑一致的金額量級；單位值為 NVIDIA 世代加權（非 NVIDIA 較便宜，故略高估）", "Derived", "（計算）",
    lambda y: f"={S('gap_gw', y)}*{S('tk_it', y)}", NF_USD, Y26)
add("gap_usd_dir", "IT 資本支出直接相減（Siliconomics − Wattonomics）　※僅供量級，不可當對帳差", "US$B", "年度口徑不同（出貨年 vs 上線年），逐年金額不可直接相減", "Derived", "（計算）",
    lambda y: f"={S('s_it', y)}-{S('w_it', y)}", NF_USD, Y26)

# 拆解
addsec("6. 差距拆解：時差（可解釋）＋ 殘差（分歧）")
add("g_time", "② 時差與庫存（換算含時差 − 直接換算）", "GW", "單年可、累計不可長期偏離 0", "Derived", "（計算）", lambda y: f"={S('w_x2', y)}-{S('w_x0', y)}", NF_GW, Y26)
add("g_resid", "③ 殘差（基準 − 含時差換算值）", "GW", "尚未歸因的分歧；R5：殘差列永遠顯示", "Derived", "（計算）", lambda y: f"={S('s_base', y)}-{S('w_x2', y)}", NF_GW, Y26, True)
add("g_resid_pct", "　殘差 ÷ Siliconomics 基準", "%", "容差 ±15%", "Derived", "（計算）", lambda y: f"={S('g_resid', y)}/{S('s_base', y)}", NF_PCT, Y26)
add("g_tol", "　容差", "%", "recon 快照 tolerance", "Assumed", "recon.tolerance", lambda y: f"={IN}!$E${REG['tol']}", NF_PCT, Y26)
add("g_judge", "　容差判定", "判定", "殘差比例絕對值 ≤ 容差＝一致，否則超出", "Derived", "（計算）",
    lambda y: f'=IF(ABS({S("g_resid_pct", y)})<={S("g_tol", y)},"一致","超出")', None, Y26, True)
add("g_share", "　殘差占總差距", "%", "", "Derived", "（計算）", lambda y: f"={S('g_resid', y)}/{S('gap_gw', y)}", NF_PCT, Y26)
add("g_cum", "　時差累計（2026 起）", "GW", "", "Derived", "（計算）",
    lambda y: (f"={S('g_time', y)}" if y == 2026 else f"={SC[y - 1]}{ROW['g_cum']}+{S('g_time', y)}"), NF_GW, Y26)
add("g_cum_pct", "　時差累計 ÷ 當年基準（檢查第 7 項，≤ 15%）", "%", "", "Derived", "（計算）", lambda y: f"={S('g_cum', y)}/{S('s_base', y)}", NF_PCT, Y26)
add("g_chk", "　檢查：時差＋殘差−差距（應為 0）", "GW", "", "Derived", "（計算）", lambda y: f"={S('g_time', y)}+{S('g_resid', y)}-{S('gap_gw', y)}", '0.000', Y26)

# 狀態
addsec("7. 協調狀態與裁定期限（來源：baseline_status.json）")
add("r_status", "狀態", "文字", "甲／乙／丙／待定", "Derived", "years.<年>.status", lambda y: f"={IN}!{YC[y]}{REG['st_status']}", None, Y26)
add("r_label", "狀態標示", "文字", "", "Derived", "years.<年>.label", lambda y: f"={IN}!{YC[y]}{REG['st_label']}", None, Y26, True)
add("r_use", "基準取法", "文字", "改用換算值／連續法／維持現行", "Derived", "use_converted、continuity",
    lambda y: (f'=IF({IN}!{YC[y]}{REG["st_continuity"]}=1,"連續法（前一年基準 × 換算值成長率）",IF({IN}!{YC[y]}{REG["st_use_conv"]}=1,"改用換算值","維持現行（對帳中）"))'), None, Y26)
add("r_dl", "裁定期限", "日期", "對帳中年份的期限；無期限者顯示 C0 期限", "Derived", "deadline、c0_deadline",
    lambda y: (f'=IF({IN}!{YC[y]}{REG["st_deadline"]}<>"—",{IN}!{YC[y]}{REG["st_deadline"]},'
               f'IF({IN}!{YC[y]}{REG["st_c0_deadline"]}<>"—","C0 "&{IN}!{YC[y]}{REG["st_c0_deadline"]},"—"))'), None, Y26)
add("r_c0", "NVIDIA 實績檢驗（C0）", "狀態", "未通過＝公司實績對層級超出 ±15%；其餘年度尚無實績", "Derived", "years.<年>.c0_fail",
    lambda y: (f'=IF({IN}!{YC[y]}{REG["st_c0_fail"]}=1,"未通過","無實績")'), None, Y26)

# 證據
addsec("8. 證據比例（基準值中由實績或已簽約支撐的比例；低／基準／高）")
for who, pre, tg in (("Wattonomics", "st_ws_", "Derived"), ("Siliconomics", "st_ss_", "Assumed")):
    for lv, nm in (("lo", "低"), ("base", "基準"), ("hi", "高")):
        add(f"e_{pre}{lv}", f"{who} 證據比例（{nm}）", "%", "", tg, f"{pre}{lv}（status 檔）",
            (lambda y, pre=pre, lv=lv: f"={IN}!{YC[y]}{REG[pre + lv]}"), NF_PCT, Y26, lv == "base")
add("e_thr", "「有證據」門檻（規則 C3 首次判定）", "%", "", "Assumed", "baseline_status.threshold", lambda y: f"={IN}!$E${REG['st_threshold']}", NF_PCT, Y26)
add("e_both", "兩層基準證據比例是否都過門檻（示意；規則判定以第 7 節為準）", "是／否", "", "Derived", "（計算）",
    lambda y: f'=IF(AND({S("e_st_ws_base", y)}>={S("e_thr", y)},{S("e_st_ss_base", y)}>={S("e_thr", y)}),"是","否")', None, Y26)

# 核對
addsec(f"9. 核對：本表重算值 vs 橋接表 {BV} 檔內值（橋接表 In 頁標示 Siliconomics {BSV} 輸入）")
for k, mine, lab in (("g", "gap_gw", "總差距"), ("g2", "g_time", "時差"), ("g3", "g_resid", "殘差")):
    add(f"b_{k}", f"橋接表{lab}", "GW", "橋接表 Bridge 頁", "Derived", f"Bridge!{k}", (lambda y, k=k: f"={IR('br_' + k, y)}"), NF_GW, Y26)
    add(f"d_{k}", f"　本表 − 橋接表（{lab}）", "GW", "", "Derived", "（計算）", (lambda y, k=k, mine=mine: f"={S(mine, y)}-{S('b_' + k, y)}"), '0.000;-0.000;0.000', Y26)
add("d_flag", "　核對結論", "判定", ("任一差異絕對值 > 0.05 GW 即標示；橋接表與本表讀同一版 Siliconomics（" + SV + "），差異應接近 0" if BR_SAME else f"任一差異絕對值 > 0.05 GW 即標示；差異來自橋接表用 {BSV} 輸入（2029 基準 {_br_s0_29:.3f} vs {SV} {_sn_b29:.3f}）"), "Derived", "（計算）",
    lambda y: f'=IF(MAX(ABS({S("d_g", y)}),ABS({S("d_g2", y)}),ABS({S("d_g3", y)}))>0.05,"有差異（>0.05 GW）","相符（≤0.05 GW）")', None, Y26)
add("d_w1", "核對：IT 分項加總 − IT 合計（Wattonomics）", "US$B", "應為 0", "Derived", "輸入頁核對列", lambda y: f"={IN}!{YC[y]}{REG['w_chk_it']}", '0.000')
add("d_w2", "核對：IT＋廠房＋自備發電 − 總支出（Wattonomics）", "US$B", "應為 0", "Derived", "輸入頁核對列", lambda y: f"={IN}!{YC[y]}{REG['w_chk_tot']}", '0.000')
add("d_w3", "核對：世代占比合計 − 1", "比例", "應為 0", "Derived", "輸入頁核對列", lambda y: f"={IN}!{YC[y]}{REG['mix|sum']}-1", '0.000')
add("d_dem", "核對：本表含時差換算 − SIF_DeployGW_Demand", "GW", "快照 Demand 只存 3 位小數，差異應 < 0.001", "Derived", "（計算）", lambda y: f"={S('w_x2', y)}-{IR('SIF_DeployGW_Demand', y)}", '0.0000', Y26)

# ---- 排版 ----
title(ws, 1, "三層對帳總表：Tokenomics × Wattonomics × Siliconomics")
put(ws, 2, 1,
    f"資料版本：Tokenomics {TV}（main {GIT_TK[0][:7]}）／Wattonomics {WV}／Siliconomics {SV}／協調橋接 {BV}；company-models main {GIT_CM[0][:7]}。"
    "單位：GW＝IT 關鍵電力（Siliconomics 為成本等值 GW）；金額 US$B。年度口徑：Wattonomics＝上線年、Siliconomics＝出貨年，逐年金額不可直接相減。藍字＝輸入、黑字＝公式。",
    font=F_NOTE, border=False)
ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=10)
ws.row_dimensions[2].height = 38
ws.cell(2, 1).alignment = AL_LT
hdr(ws, 4, ["項目", "單位", "取法／說明", "2025", "2026", "2027", "2028", "2029", "等級", "來源名稱"])
r = 5
for spec in rows:
    if spec["kind"] == "sec":
        sec(ws, r, spec["text"], 10)
    else:
        ROW[spec["key"]] = r
    r += 1
r = 5
for spec in rows:
    if spec["kind"] == "sec":
        r += 1
        continue
    bold = spec["bold"]
    put(ws, r, 1, spec["label"], font=F_BOLD if bold else F_BASE)
    put(ws, r, 2, spec["unit"], al=AL_C)
    put(ws, r, 3, spec["how"], font=F_NOTE, al=AL_L)
    for y in YEARS:
        c = openpyxl.utils.column_index_from_string(SC[y])
        if y in spec["years"]:
            cell = put(ws, r, c, spec["fn"](y), nf=spec["nf"], font=F_BOLD if bold else F_BASE,
                       al=(AL_C if spec["nf"] is None else AL_R))
        else:
            put(ws, r, c, "—", font=Font(name=FN, size=10, color="7F7F7F"), al=AL_R)
    put(ws, r, 9, spec["tag"], font=F_NOTE, al=AL_C)
    put(ws, r, 10, spec["src"], font=F_NOTE, al=AL_L)
    r += 1
LAST_SUM = r - 1
widths(ws, [52, 12, 52, 17, 17, 17, 17, 17, 18, 32])
ws.freeze_panes = "D5"
ws.sheet_view.zoomScale = 90
import math
for rr in range(5, LAST_SUM + 1):
    la = len(str(ws.cell(rr, 1).value or ""))
    lc = len(str(ws.cell(rr, 3).value or ""))
    lj = len(str(ws.cell(rr, 10).value or ""))
    lines = max(math.ceil(la / 46), math.ceil(lc / 58), math.ceil(lj / 34), 1)
    for cc in range(4, 9):
        vv = ws.cell(rr, cc).value
        if isinstance(vv, str) and not vv.startswith("=") and len(vv) > 16:
            lines = max(lines, math.ceil(len(vv) / 14))
    ws.row_dimensions[rr].height = 14 * lines + 4
for k in ("v_rule", "r_label", "tk_mixtxt", "r_use"):
    ws.row_dimensions[ROW[k]].height = 44
page(ws, True, 1, title_rows="4:4", area=f"A1:J{LAST_SUM}")

# 狀態色（條件格式）：一致＝綠、邊界／有條件＝黃、超出／未通過＝紅、待定／連續法＝灰
def status_cf(ws, rng, tl):
    G = PatternFill("solid", bgColor="C6EFCE")
    Yl = PatternFill("solid", bgColor="FFEB9C")
    Rd = PatternFill("solid", bgColor="FFC7CE")
    Gy = PatternFill("solid", bgColor="D9D9D9")
    ws.conditional_formatting.add(rng, FormulaRule(formula=[f'OR(ISNUMBER(SEARCH("有條件",{tl})),ISNUMBER(SEARCH("邊界",{tl})))'], fill=Yl, font=Font(name=FN, bold=True, color="7F6000"), stopIfTrue=True))
    ws.conditional_formatting.add(rng, FormulaRule(formula=[f'OR(ISNUMBER(SEARCH("超出",{tl})),ISNUMBER(SEARCH("未通過",{tl})),ISNUMBER(SEARCH("不一致",{tl})))'], fill=Rd, font=Font(name=FN, bold=True, color="9C0006"), stopIfTrue=True))
    ws.conditional_formatting.add(rng, FormulaRule(formula=[f'OR(ISNUMBER(SEARCH("待定",{tl})),ISNUMBER(SEARCH("連續法",{tl})),ISNUMBER(SEARCH("無實績",{tl})),ISNUMBER(SEARCH("對帳中",{tl})))'], fill=Gy, font=Font(name=FN, bold=True, color="404040"), stopIfTrue=True))
    ws.conditional_formatting.add(rng, FormulaRule(formula=[f'OR(ISNUMBER(SEARCH("一致",{tl})),ISNUMBER(SEARCH("相符",{tl})))'], fill=G, font=Font(name=FN, bold=True, color="006100"), stopIfTrue=True))


for k in ("v_judge", "v_rule", "v_c0", "g_judge", "r_label", "r_c0", "r_status", "d_flag"):
    status_cf(ws, f"E{ROW[k]}:H{ROW[k]}", f"E{ROW[k]}")

# ----------------------------------------------------------------------------------------------
# 「實績對帳」
# ----------------------------------------------------------------------------------------------
ws = ws_act
title(ws, 1, "實績對帳：NVIDIA 資料中心營收，兩層隱含值 vs 公司實績")
put(ws, 2, 1, "曆年 N 約對應 NVIDIA FY(N+1)（例：2026 ≈ FY2027，2025 ≈ FY2026）。「隱含值」＝該層 GW × 每 GW NVIDIA 內容，由來源檔提供，本頁不重算。"
    "模型內採用的公司實績到 FY2027 Q1＋Q2（75.2、89.0）。下方橘色區塊為樣本外檢驗，不回饋任何其他格。", font=F_NOTE, border=False)
ws.merge_cells("A2:H2")
ws.row_dimensions[2].height = 40
ws.cell(2, 1).alignment = AL_LT
r = 4
sec(ws, r, "A. 全年隱含值 vs 公司值（US$B）", 8)
r += 1
hdr(ws, r, ["曆年（NVIDIA 財年）", "Siliconomics 隱含", "Wattonomics 隱含", "公司值", "公司值性質", "Siliconomics ÷ 公司", "Wattonomics ÷ 公司", "備註"])
ACT = {}
spec_rows = [
    (2025, "2025（FY2026）", None, "WIF_Recon_NV_ImpliedRev", "WIF_Recon_NV_ActualRev", "全年實績", "Siliconomics 無 2025 隱含值"),
    (2026, "2026（FY2027）", "nv_sil_implied", "WIF_Recon_NV_ImpliedRev", "WIF_Recon_NV_ActualRev", "含指引（對帳檔標示；只列參考）", "公司值含指引，只作參考；主要對帳見下方 B 區"),
    (2027, "2027（FY2028）", "nv_sil_implied", "WIF_Recon_NV_ImpliedRev", "WIF_Recon_NV_ActualRev", "含指引（對帳檔標示；尚無全年實績）", "尚無實績；只列參考"),
]
for y, lab, silkey, wkey, ckey, nature, note in spec_rows:
    r += 1
    ACT[y] = r
    put(ws, r, 1, lab)
    put(ws, r, 2, (f"={IR(silkey, y)}" if silkey else "—"), nf=NF_USD, al=AL_R)
    put(ws, r, 3, f"={IR(wkey, y)}", nf=NF_USD)
    put(ws, r, 4, f"={IR(ckey, y)}", nf=NF_USD)
    put(ws, r, 5, nature, al=AL_L)
    put(ws, r, 6, (f"=B{r}/D{r}" if silkey else "—"), nf=NF_X, al=AL_R)
    put(ws, r, 7, f"=C{r}/D{r}", nf=NF_X)
    put(ws, r, 8, note, font=F_NOTE)
r += 2
sec(ws, r, "B. 模型內實績：FY2027 Q1＋Q2（已入模型）", 8)
r += 1
hdr(ws, r, ["項目", "Siliconomics", "Wattonomics", "實績／參數", "單位", "公式說明", "對帳檔記錄值", "與記錄值差"])
r += 1
rq1 = r
put(ws, r, 1, "FY2027 Q1 實績（8-K）")
put(ws, r, 2, "—", al=AL_R)
put(ws, r, 3, "—", al=AL_R)
put(ws, r, 4, f"={IN}!$E${REG['nv_q1']}", nf=NF_USD)
put(ws, r, 5, "US$B", al=AL_C)
r += 1
rq2 = r
put(ws, r, 1, "FY2027 Q2 實績（8-K）")
put(ws, r, 2, "—", al=AL_R)
put(ws, r, 3, "—", al=AL_R)
put(ws, r, 4, f"={IN}!$E${REG['nv_q2']}", nf=NF_USD)
put(ws, r, 5, "US$B", al=AL_C)
r += 1
rcum = r
put(ws, r, 1, "H1 實績累計（Q1＋Q2）", font=F_BOLD)
put(ws, r, 2, "—", al=AL_R)
put(ws, r, 3, "—", al=AL_R)
put(ws, r, 4, f"=D{rq1}+D{rq2}", nf=NF_USD, font=F_BOLD)
put(ws, r, 5, "US$B", al=AL_C)
put(ws, r, 6, "Q1＋Q2", font=F_NOTE)
put(ws, r, 7, f"={IN}!$E${REG['nv_cum_rec']}", nf=NF_USD)
put(ws, r, 8, f"=D{r}-G{r}", nf='0.000')
r += 1
rmo = r
put(ws, r, 1, "已實績月數")
put(ws, r, 2, "—", al=AL_R)
put(ws, r, 3, "—", al=AL_R)
put(ws, r, 4, f"={IN}!$E${REG['nv_months']}", nf=NF_INT)
put(ws, r, 5, "月", al=AL_C)
r += 1
rfy = r
put(ws, r, 1, "2026 全年隱含值（US$B）")
put(ws, r, 2, f"=B{ACT[2026]}", nf=NF_USD)
put(ws, r, 3, f"=C{ACT[2026]}", nf=NF_USD)
put(ws, r, 4, "—", al=AL_R)
put(ws, r, 5, "US$B", al=AL_C)
put(ws, r, 6, "取自 A 區", font=F_NOTE)
r += 1
rsame = r
put(ws, r, 1, "同期間隱含值（全年 × 已實績月數 ÷ 12）")
put(ws, r, 2, f"=B{rfy}*$D${rmo}/12", nf=NF_USD)
put(ws, r, 3, f"=C{rfy}*$D${rmo}/12", nf=NF_USD)
put(ws, r, 4, "—", al=AL_R)
put(ws, r, 5, "US$B", al=AL_C)
put(ws, r, 6, "按月數等比例；假設各月均勻，不含季節權重", font=F_NOTE)
r += 1
rm1 = r
put(ws, r, 1, "主讀數：H1 實績 ÷ 同期間隱含值", font=F_BOLD)
put(ws, r, 2, f"=$D${rcum}/B{rsame}", nf=NF_X, font=F_BOLD)
put(ws, r, 3, f"=$D${rcum}/C{rsame}", nf=NF_X, font=F_BOLD)
put(ws, r, 4, "—", al=AL_R)
put(ws, r, 5, "比", al=AL_C)
put(ws, r, 6, "＜1 表示層級隱含值高於公司實績", font=F_NOTE)
put(ws, r, 7, f"={IR('nv_ratio_main_rec', 2026)}", nf=NF_X)
put(ws, r, 8, f"=B{r}-G{r}", nf='0.000')
r += 1
rq4 = r
put(ws, r, 1, "Q2 年化（Q2 × 4）")
put(ws, r, 2, "—", al=AL_R)
put(ws, r, 3, "—", al=AL_R)
put(ws, r, 4, f"=D{rq2}*4", nf=NF_USD)
put(ws, r, 5, "US$B", al=AL_C)
r += 1
rm2 = r
put(ws, r, 1, "第二讀數：Q2 年化 ÷ 全年隱含值", font=F_BOLD)
put(ws, r, 2, f"=$D${rq4}/B{rfy}", nf=NF_X, font=F_BOLD)
put(ws, r, 3, f"=$D${rq4}/C{rfy}", nf=NF_X, font=F_BOLD)
put(ws, r, 4, "—", al=AL_R)
put(ws, r, 5, "比", al=AL_C)
put(ws, r, 7, f"={IR('nv_ratio_second_rec', 2026)}", nf=NF_X)
put(ws, r, 8, f"=B{r}-G{r}", nf='0.000')
r += 1
rpct = r
put(ws, r, 1, "主讀數偏離（讀數 − 1）")
put(ws, r, 2, f"=B{rm1}-1", nf=NF_PCT)
put(ws, r, 3, f"=C{rm1}-1", nf=NF_PCT)
put(ws, r, 6, f"對帳檔判定：Siliconomics 超出 ±{int(TOL * 100)}%（C0 未通過）；Wattonomics 對帳檔記錄在 2025 年為 {BR and ''}", font=F_NOTE)
ws.cell(r, 6).value = f"容差 ±{int(round(TOL * 100))}%；Siliconomics 2026 C0 標示未通過（baseline_status）"
r += 1
put(ws, r, 1, "對帳檔 C0 說明", font=F_NOTE)
put(ws, r, 2, f"={IN}!E{REG['st_c0_point']}", font=F_NOTE, al=AL_L)
ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=8)
ws.row_dimensions[r].height = 32
r += 2

# 樣本外
sec(ws, r, "C. 樣本外檢驗（未入模型）　NVIDIA FY2027 Q3", 8)
for c in range(1, 9):
    ws.cell(r, c).fill = FILL_OOS
    ws.cell(r, c).font = F_WARN
r += 1
put(ws, r, 1, "【警示】本區數字不進任何模型、快照、種子檔、係數或權重；沒有任何其他格引用本區。"
    f"Q3 數字來源：{A.oos_source}。只算與兩層讀數的比值，供 Q3 實績成為正式對帳來源前先看方向。僅一季，不足以裁定。",
    font=F_WARN, fill=FILL_OOS, al=AL_L)
ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=8)
ws.row_dimensions[r].height = 44
r += 1
hdr(ws, r, ["項目", "Siliconomics", "Wattonomics", "Q3 數字", "單位", "公式說明", "", ""])
for c in range(1, 9):
    ws.cell(r, c).fill = PatternFill("solid", fgColor="C55A11")
r += 1
oq3 = r
put(ws, r, 1, "FY2027 Q3 資料中心營收（樣本外）", font=F_INB, fill=FILL_OOS)
put(ws, r, 2, "—", al=AL_R, fill=FILL_OOS)
put(ws, r, 3, "—", al=AL_R, fill=FILL_OOS)
put(ws, r, 4, f"={IN}!$E${REG['nv_oos_q3']}", nf=NF_USD, font=F_BOLD, fill=FILL_OOS)
put(ws, r, 5, "US$B", al=AL_C, fill=FILL_OOS)
put(ws, r, 6, "輸入頁橘色列", font=F_NOTE, fill=FILL_OOS)
r += 1
oann = r
put(ws, r, 1, "Q3 年化（Q3 × 4）", fill=FILL_OOS)
put(ws, r, 2, "—", al=AL_R, fill=FILL_OOS)
put(ws, r, 3, "—", al=AL_R, fill=FILL_OOS)
put(ws, r, 4, f"=D{oq3}*4", nf=NF_USD, fill=FILL_OOS)
put(ws, r, 5, "US$B", al=AL_C, fill=FILL_OOS)
r += 1
oq = r
put(ws, r, 1, "兩層單季隱含值（全年 ÷ 4）", fill=FILL_OOS)
put(ws, r, 2, f"=B{rfy}/4", nf=NF_USD, fill=FILL_OOS)
put(ws, r, 3, f"=C{rfy}/4", nf=NF_USD, fill=FILL_OOS)
put(ws, r, 4, "—", al=AL_R, fill=FILL_OOS)
put(ws, r, 5, "US$B", al=AL_C, fill=FILL_OOS)
put(ws, r, 6, "假設四季均勻；不含季節權重", font=F_NOTE, fill=FILL_OOS)
r += 1
orat = r
put(ws, r, 1, "Q3 ÷ 兩層單季隱含值", font=F_INB, fill=FILL_OOS)
put(ws, r, 2, f"=$D${oq3}/B{oq}", nf=NF_X, font=F_BOLD, fill=FILL_OOS)
put(ws, r, 3, f"=$D${oq3}/C{oq}", nf=NF_X, font=F_BOLD, fill=FILL_OOS)
put(ws, r, 4, "—", al=AL_R, fill=FILL_OOS)
put(ws, r, 5, "比", al=AL_C, fill=FILL_OOS)
put(ws, r, 6, "＝Q3 年化 ÷ 全年隱含值", font=F_NOTE, fill=FILL_OOS)
r += 1
odev = r
put(ws, r, 1, "偏離（比值 − 1）", fill=FILL_OOS)
put(ws, r, 2, f"=B{orat}-1", nf=NF_PCT, fill=FILL_OOS)
put(ws, r, 3, f"=C{orat}-1", nf=NF_PCT, fill=FILL_OOS)
r += 1
oclo = r
put(ws, r, 1, "較接近 Q3 的一層（僅此一季）", fill=FILL_OOS)
put(ws, r, 2, f'=IF(ABS(B{odev})<ABS(C{odev}),"較接近","較遠")', al=AL_C, fill=FILL_OOS)
put(ws, r, 3, f'=IF(ABS(C{odev})<ABS(B{odev}),"較接近","較遠")', al=AL_C, fill=FILL_OOS)
put(ws, r, 6, "單季比較，且與 Q2 年化（B 區第二讀數）方向可能不同；不構成裁定", font=F_NOTE, fill=FILL_OOS)
r += 1
oqq = r
put(ws, r, 1, "對照：Q3 ÷ Q2（模型內實績）", fill=FILL_OOS)
put(ws, r, 2, "—", al=AL_R, fill=FILL_OOS)
put(ws, r, 3, "—", al=AL_R, fill=FILL_OOS)
put(ws, r, 4, f"=D{oq3}/D{rq2}", nf=NF_X, fill=FILL_OOS)
put(ws, r, 5, "比", al=AL_C, fill=FILL_OOS)
LAST_ACT = r
widths(ws, [44, 16, 16, 16, 30, 40, 16, 40])
page(ws, True, 1, area=f"A1:H{LAST_ACT}")
ws.freeze_panes = "B4"
OOS_CELLS = (oq3, LAST_ACT)

# ----------------------------------------------------------------------------------------------
# 「介面覆蓋」
# ----------------------------------------------------------------------------------------------
ws = ws_cov
title(ws, 1, "介面覆蓋：Tokenomics 往下傳了什麼、誰用了；Siliconomics 13 個對帳點")
put(ws, 2, 1, "「公式引用次數」＝掃描該模型工作簿所有公式（排除 TK_In 鏡像頁自身）中，TK 名稱或其本庫鏡像名稱出現的次數；"
    "以儲存格位址間接引用者掃不到，故 0 次＝未被直接引用，不等於一定沒用。", font=F_NOTE, border=False)
ws.merge_cells("A2:J2")
ws.row_dimensions[2].height = 30
ws.cell(2, 1).alignment = AL_LT
r = 4
sec(ws, r, f"A. Tokenomics 快照下傳的 23 個名稱（Tokenomics {TV} → Wattonomics {WV}／Siliconomics {SV}）", 10)
r += 1
hdr(ws, r, ["名稱", "內容", "單位", "快照類型", "Wattonomics 有鏡像", "Wattonomics 公式引用次數", "Siliconomics 有鏡像", "Siliconomics 公式引用次數", "判定", "備註"])
A_FIRST = r + 1
for nm, it in TK_ITEMS.items():
    r += 1
    lab = re.split(r"[（(]", it["label"])[0]
    uw, us = USE_W[nm], USE_S[nm]
    put(ws, r, 1, nm, font=F_IN, al=AL_L)
    put(ws, r, 2, lab, font=F_IN, al=AL_L)
    put(ws, r, 3, it["unit"], font=F_IN, al=AL_C)
    put(ws, r, 4, it["kind"], font=F_IN, al=AL_C)
    put(ws, r, 5, ("是" if uw["mirror"] else "否") if uw["mirror"] is not None else "—", font=F_IN, al=AL_C)
    put(ws, r, 6, uw["refs"] if uw["refs"] is not None else "—", nf=NF_INT, font=F_IN)
    put(ws, r, 7, ("是" if us["mirror"] else "否") if us["mirror"] is not None else "—", font=F_IN, al=AL_C)
    put(ws, r, 8, us["refs"] if us["refs"] is not None else "—", nf=NF_INT, font=F_IN)
    put(ws, r, 9, f'=IF(AND(ISNUMBER(F{r}),ISNUMBER(H{r})),IF(F{r}+H{r}=0,"僅鏡像、未被公式引用",IF(AND(F{r}>0,H{r}>0),"兩層都用",IF(F{r}>0,"只 Wattonomics 用","只 Siliconomics 用"))),"—")', al=AL_C)
    note = ""
    if nm == "IF_CapexTotal":
        note = "兩層可能改用 IT＋廠房相加，未直接引用"
    if nm == "IF_FacilityGW":
        note = "快照標為 optional；Wattonomics 用於設施電力換算"
    put(ws, r, 10, note, font=F_NOTE)
A_LAST = r
r += 1
put(ws, r, 1, "合計", font=F_BOLD)
put(ws, r, 2, f'=COUNTA(A{A_FIRST}:A{A_LAST})&" 個名稱；僅鏡像未被引用 "&COUNTIF(I{A_FIRST}:I{A_LAST},"僅鏡像*")&" 個；兩層都用 "&COUNTIF(I{A_FIRST}:I{A_LAST},"兩層都用")&" 個"', font=F_BOLD)
ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=9)
r += 2

sec(ws, r, "B. 重要介面沒有往下傳（不在 23 個名稱內，兩層公式亦未引用）", 10)
r += 1
hdr(ws, r, ["名稱", "內容", "單位", "VR200 基準值", "在 23 名稱快照內", "Wattonomics 公式引用次數", "Siliconomics 公式引用次數", "判定", "意義", ""])
B_ITEMS = [("IF_HoldEcon", "TKX", "每 GW 年持有成本（經濟口徑）", "三層協調鏈沒有持有成本出口；Tokenomics 與 Wattonomics 成本面同源，卻看不到持有成本"),
           ("IF_HoldAcct", "TKX", "每 GW 年持有成本（會計口徑）", "同上"),
           ("IF_RevGWFleet", "TKX", "付費服務營收合計（層級組合）", "理想上限（OpenAI 有效單價）；收入面沒有傳給兩層。依建議書 r2 V9，Tokenomics 收入介面有被個股模型讀取，沒讀的是三層協調鏈")]
for nm, code, lab, meaning in B_ITEMS:
    r += 1
    d = TKX[nm]
    put(ws, r, 1, nm, font=F_IN, al=AL_L)
    put(ws, r, 2, lab, font=F_IN, al=AL_L)
    put(ws, r, 3, d["unit"], font=F_IN, al=AL_C)
    put(ws, r, 4, f"={IN}!E{TKROW[(nm, 3)]}", nf=NF_USD)
    put(ws, r, 5, "否" if nm not in TK_ITEMS else "是", font=F_IN, al=AL_C)
    put(ws, r, 6, USE_W[nm]["refs"] if USE_W[nm]["refs"] is not None else "—", nf=NF_INT, font=F_IN)
    put(ws, r, 7, USE_S[nm]["refs"] if USE_S[nm]["refs"] is not None else "—", nf=NF_INT, font=F_IN)
    put(ws, r, 8, f'=IF(AND(ISNUMBER(F{r}),ISNUMBER(G{r})),IF(F{r}+G{r}=0,"未使用","有引用"),"—")', al=AL_C)
    put(ws, r, 9, meaning, font=F_NOTE)
    ws.merge_cells(start_row=r, start_column=9, end_row=r, end_column=10)
    ws.row_dimensions[r].height = 42
for nm, lab in [("IF_RevGW_Sol", "每 GW 理論營收（Sol 中層）"), ("IF_RevGW_Luna", "每 GW 理論營收（Luna 低層）"), ("IF_RevGW_Astra", "每 GW 理論營收（Astra 頂層）"),
                ("IF_RevGW_Terra", "每 GW 理論營收（Terra 中低層）"), ("IF_OpexGW", "每 GW 營運費用（不含折舊）"), ("IF_FullCost_Sol", "完全成本（Sol）"), ("IF_GPUhrEcon", "每 GPU 小時持有成本（經濟）")]:
    r += 1
    put(ws, r, 1, nm, font=F_IN, al=AL_L)
    put(ws, r, 2, lab, font=F_IN, al=AL_L)
    put(ws, r, 3, "—", al=AL_C)
    put(ws, r, 4, "—", al=AL_C)
    put(ws, r, 5, "否", font=F_IN, al=AL_C)
    put(ws, r, 6, USE_W[nm]["refs"] if USE_W[nm]["refs"] is not None else "—", nf=NF_INT, font=F_IN)
    put(ws, r, 7, USE_S[nm]["refs"] if USE_S[nm]["refs"] is not None else "—", nf=NF_INT, font=F_IN)
    put(ws, r, 8, f'=IF(AND(ISNUMBER(F{r}),ISNUMBER(G{r})),IF(F{r}+G{r}=0,"未使用","有引用"),"—")', al=AL_C)
    put(ws, r, 9, "（工作簿具名範圍存在；值未讀入本表）", font=F_NOTE)
    ws.merge_cells(start_row=r, start_column=9, end_row=r, end_column=10)
r += 2

sec(ws, r, f"C. Siliconomics 13 個對帳點：檢驗方式與狀態（{os.path.basename(P['recon'])}）", 10)
r += 1
hdr(ws, r, ["代號", "對帳點", "檢驗方式", "狀態", "2027 市場值（US$B／份額）", "2027 已建模加總", "加總 ÷ 市場值", "依據", "對應市場值名稱（快照）", "備註"])
POINT_ORDER = ["soc", "gpu_asic", "cowos", "wfe", "pcb_ccl", "optics", "asic_ds", "cpo", "abf", "power", "liquid", "asm", "hbm"]
MKT_MAP = {"soc": "SIF_TestSoCMkt_Base", "gpu_asic": "SIF_GpuAsicMkt_Base", "pcb_ccl": "SIF_PCBMkt_Base", "power": "SIF_PwrMkt_Base",
           "liquid": "SIF_LiqMkt_Base", "asm": "SIF_AsmMkt_Base", "hbm": "SIF_HBM_Mkt_Base"}
C_FIRST = r + 1
for i, k in enumerate(POINT_ORDER, start=1):
    r += 1
    p = MK[k]
    kind = p["kind"]
    method = {"market": ("個股實績加總" if "modeled_sum" in p else "未建"), "structural": "結構一致"}[kind]
    put(ws, r, 1, f"{i:02d} {k}", font=F_IN, al=AL_L)
    put(ws, r, 2, p["name"], font=F_IN, al=AL_L)
    put(ws, r, 3, method, font=F_IN, al=AL_C)
    put(ws, r, 4, p["status"], font=F_IN, al=AL_C)
    if "modeled_sum" in p:
        yi = p["years"].index(2027)
        put(ws, r, 5, p["market"][yi], nf=(NF_USD if p.get("unit") == "US$B" else NF_PCT), font=F_IN)
        put(ws, r, 6, p["modeled_sum"][yi], nf=(NF_USD if p.get("unit") == "US$B" else NF_PCT), font=F_IN)
        put(ws, r, 7, f"=F{r}/E{r}", nf=NF_PCT)
    else:
        for c in (5, 6, 7):
            put(ws, r, c, "—", al=AL_R)
    put(ws, r, 8, p.get("check") or (p["rows"][0]["ref"].split(" ")[0] if p.get("rows") else "—"), font=F_IN, al=AL_L)
    if k in MKT_MAP:
        put(ws, r, 9, f"={IR(MKT_MAP[k], 2027)}", nf=NF_USD)
    else:
        put(ws, r, 9, "—", al=AL_R)
    nt = ""
    if method == "未建" and k in MKT_MAP:
        nt = "市場值已有，但對帳未建（無「已建模加總＋殘差＝市場值」）"
    elif method == "未建":
        nt = "市場值與對帳均未建"
    elif method == "結構一致":
        nt = "只驗證結構（Checks 頁），非實績加總"
    elif k == "soc":
        nt = f"份額口徑；預期殘差 {p.get('expected_residual')}"
    put(ws, r, 10, nt, font=F_NOTE)
C_LAST = r
r += 1
for lab, key in (("個股實績加總", "個股實績加總"), ("結構一致", "結構一致"), ("未建", "未建")):
    put(ws, r, 1, lab, font=F_BOLD)
    put(ws, r, 2, f'=COUNTIF($C${C_FIRST}:$C${C_LAST},"{key}")', nf=NF_INT)
    r += 1
put(ws, r, 1, "合計", font=F_BOLD)
put(ws, r, 2, f"=COUNTA($A${C_FIRST}:$A${C_LAST})", nf=NF_INT, font=F_BOLD)
r += 2
sec(ws, r, "D. 另有兩個非個股點（不計入 13 個）", 10)
r += 1
hdr(ws, r, ["代號", "名稱", "種類", "狀態", "2026 讀數 1", "2026 讀數 2", "", "依據", "", "備註"])
for k, rr_ in (("r4", MK["r4"]), ("nv_dc", MK["nv_dc"])):
    r += 1
    put(ws, r, 1, k, font=F_IN, al=AL_L)
    put(ws, r, 2, rr_["name"], font=F_IN, al=AL_L)
    put(ws, r, 3, "層間對帳" if k == "r4" else "公司實績", font=F_IN, al=AL_C)
    put(ws, r, 4, "；".join(rr_["status_by_year"][:2]) if k == "r4" else rr_["status"], font=F_IN, al=AL_L)
    if k == "r4":
        put(ws, r, 5, rr_["sil_online"][0], nf=NF_GW, font=F_IN)
        put(ws, r, 6, rr_["wt_main"][0], nf=NF_GW, font=F_IN)
        put(ws, r, 10, "Siliconomics 可支撐上線 GW（非獨立上限）vs Wattonomics 主路徑上線 GW", font=F_NOTE)
    else:
        put(ws, r, 5, f"={IR('nv_sil_implied', 2026)}", nf=NF_USD)
        put(ws, r, 6, f"={IR('WIF_Recon_NV_ImpliedRev', 2026)}", nf=NF_USD)
        put(ws, r, 10, "兩層隱含營收；實績見「實績對帳」頁", font=F_NOTE)
    put(ws, r, 8, rr_["ref"][:60], font=F_IN, al=AL_L)
widths(ws, [24, 36, 14, 16, 18, 18, 18, 30, 22, 44])
page(ws, True, 1)
ws.freeze_panes = "A4"
status_cf(ws, f"D{C_FIRST}:D{C_LAST}", f"D{C_FIRST}")
ws.conditional_formatting.add(f"C{C_FIRST}:C{C_LAST}", FormulaRule(formula=[f'C{C_FIRST}="未建"'], fill=PatternFill("solid", bgColor="D9D9D9"), font=Font(name=FN, bold=True, color="404040")))

# ----------------------------------------------------------------------------------------------
# 「欄位語義」
# ----------------------------------------------------------------------------------------------
ws = ws_sem
title(ws, 1, "欄位語義：SIF_DeployGW 逐年 GW（成本等值、出貨年）")
put(ws, 2, 1, f"同一個 2029 年有幾個不同的數：基準 {sv('SIF_DeployGW_Base', 2029):.2f}、供給側 {sv('SIF_DeployGW_Supply', 2029):.2f}、需求側 {sv('SIF_DeployGW_Demand', 2029):.2f}、另一側 {sv('SIF_DeployGW_Alt', 2029):.2f}、舊 S33 取法 {sv('SIF_DeployGW_LegacyS33', 2029):.2f}。"
    "下游「現行基準」一詞只指 Base。LegacyS33（v0.20 以前名為 Current）是舊取法的對照欄，不是基準。", font=F_NOTE, border=False)
ws.merge_cells("A2:I2")
ws.row_dimensions[2].height = 30
ws.cell(2, 1).alignment = AL_LT
hdr(ws, 4, ["欄位", "2026", "2027", "2028", "2029", "一句話定義", "誰該讀", "與 Base 相同？", "備註"])
SEM = [
    ("Base", "SIF_DeployGW_Base", "依規則決定的基準：" + "、".join(f"{y} {ST[str(y)]['label']}" for y in Y26 if str(y) in ST) + "。", "硬體個股算基準時只讀這一欄",
     "唯一的基準欄"),
    ("Supply", "SIF_DeployGW_Supply", "供給側原始讀數：原始顆數部署 GW，不乘係數。", "只供並列顯示、對照", "Base 含顆數係數 " + "／".join(f"{x:.2f}" for x in SN["SIF_UnitsCoef_Base"][-2:]) + "（2028／2029，Assumed）"),
    ("Demand", "SIF_DeployGW_Demand", "需求側讀數：Wattonomics 上線 GW 以平滑法換到出貨年、成本等值。", "只供並列顯示、對照", "與 Alt 同值（另一側剛好是需求側）"),
    ("Alt", "SIF_DeployGW_Alt", "另一側讀數：未被採為基準的那一條；對帳中年份的個股並列兩讀數目標價用。", "對帳中年份的個股（C4 兩讀數）", "會隨規則採哪一側而換邊；目前 = Demand"),
    ("LegacyS33", "SIF_DeployGW_LegacyS33", f"舊 S33 取法：原始 × 舊 S33 係數，即 v0.18 取法；2029 為 {sv('SIF_DeployGW_LegacyS33', 2029):.3f}，不是基準。", "不該讀；僅對照",
     "v0.20 r1 新增此名；v0.21 起刪除舊名 SIF_DeployGW_Current（決策 S58）"),
]
SEMROW = {}
r = 4
for nm, key, defin, who, note in SEM:
    r += 1
    SEMROW[nm] = r
    put(ws, r, 1, nm, font=F_BOLD)
    for i, y in enumerate(Y26):
        put(ws, r, 2 + i, f"={IR(key, y)}", nf=NF_GW)
    put(ws, r, 6, defin, al=AL_L)
    put(ws, r, 7, who, al=AL_L)
    put(ws, r, 8, f'=IF(AND(ABS(B{r}-B${5})<0.0005,ABS(C{r}-C${5})<0.0005,ABS(D{r}-D${5})<0.0005,ABS(E{r}-E${5})<0.0005),"四年皆同",'
                  f'IF(COUNTIF(B{r + 100}:B{r + 100},"x")=0,"有 "&(SUMPRODUCT(--(ABS(B{r}:E{r}-B$5:E$5)>=0.0005)))&" 年不同",""))', al=AL_C)
    put(ws, r, 9, note, font=F_NOTE, al=AL_L)
    ws.row_dimensions[r].height = 46
r += 1
put(ws, r, 1, "Alt ÷ Base", font=F_NOTE)
for i, y in enumerate(Y26):
    put(ws, r, 2 + i, f"=B{SEMROW['Alt']}/B{SEMROW['Base']}".replace("B", "BCDE"[i]), nf=NF_X)
put(ws, r, 6, "＝快照 SIF_DeployGW_AltRatio_<年>（個股把產業量乘此比值得另一讀數）", font=F_NOTE, al=AL_L)
r += 1
put(ws, r, 1, "LegacyS33 − Base", font=F_BOLD)
for i, y in enumerate(Y26):
    c = "BCDE"[i]
    put(ws, r, 2 + i, f"={c}{SEMROW['LegacyS33']}-{c}{SEMROW['Base']}", nf=NF_GW, font=F_BOLD)
put(ws, r, 6, "若個股誤讀舊取法，產業量會差這麼多 GW", font=F_NOTE, al=AL_L)
r += 1
put(ws, r, 1, "Supply − Base", font=F_NOTE)
for i, y in enumerate(Y26):
    c = "BCDE"[i]
    put(ws, r, 2 + i, f"={c}{SEMROW['Supply']}-{c}{SEMROW['Base']}", nf=NF_GW)
put(ws, r, 6, "差距＝顆數係數與規則採側造成的部分", font=F_NOTE, al=AL_L)
widths(ws, [16, 12, 12, 12, 12, 58, 34, 16, 48])
page(ws, True, 1)
ws.freeze_panes = "B5"
# 修正「與 Base 相同？」欄為簡單公式（上面的複雜式只保留簡單分支）
for nm, key, *_ in SEM:
    rr_ = SEMROW[nm]
    ws.cell(rr_, 8).value = (f'=IF(SUMPRODUCT(--(ABS(B{rr_}:E{rr_}-B$5:E$5)>=0.0005))=0,"四年皆同","有 "&SUMPRODUCT(--(ABS(B{rr_}:E{rr_}-B$5:E$5)>=0.0005))&" 年不同")')

# ----------------------------------------------------------------------------------------------
# 「說明」
# ----------------------------------------------------------------------------------------------
ws = ws_note
title(ws, 1, "說明：單位、口徑、等級標籤、已知限制")
NOTES = [
    ("單位", [
        "GW＝IT 關鍵電力（Tokenomics Interface 第 1 列定義；設施電力＝IT × IF_FacilityGW，本表不用）。Siliconomics 的 GW 為「成本等值 GW」（IT 資本支出 ÷ 每 GW IT 資本支出），與實體 GW 差一個比值（總表第 3 節 rt）。",
        "金額一律 US$B（十億美元）；每 GW 單位值為 US$B／GW。NVIDIA 營收為資料中心營收，US$B。",
    ]),
    ("年度口徑", [
        "Wattonomics 為「上線年」（新增上線 GW、支出皆依上線年）；Siliconomics 為「出貨年」。同一曆年的兩個數字描述的是不同時間點的設備，逐年金額不可直接相減。",
        "總表第 3 節用規則 v1.1 平滑法把 Wattonomics 上線 GW 換到出貨年：出貨ₜ ≈ 上線ₜ ＋ k × (上線ₜ₊₁ − 上線ₜ)，k＝lag_on ÷ 12 ＋ inv_on（皆 Assumed，輸入頁有低／基準／高）。",
        "「差距」定義與橋接表相同：Siliconomics 基準 − Wattonomics 上線 GW 直接換成成本等值。拆成「時差」（含時差換算 − 直接換算）與「殘差」（基準 − 含時差換算）。殘差列永遠顯示，不設結案門檻（建議書 R5）。",
        "金額差距有兩種：折算金額（差距 GW × Tokenomics 每 GW IT 單位值，口徑一致）與直接相減（口徑不同，只看量級）。",
    ]),
    ("Tokenomics 每 GW 單位值怎麼取", [
        "每年用 Wattonomics Players 頁的「新增 GW 世代占比」（NVIDIA 內部占比）對 Tokenomics 工作簿 IF_CapexIT 等具名範圍的 5 個世代加權（SUMPRODUCT），低／基準／高成本分列。結果與 Wattonomics CapexGW 第 7 列「NVIDIA 每 IT GW IT capex（世代加權）」一致。",
        "世代占比不在快照內（快照為空白），所以輸入頁第 2 節直讀 Wattonomics 模型檔；非 NVIDIA 平台未納入此加權，Wattonomics 另以 IF_NonNV_Hold 比例處理，故「隱含每 GW IT 支出」低於 NVIDIA 加權值屬正常。",
        "IF_HoldEcon（持有成本）與 IF_RevGWFleet（收入）不在下傳的 23 個名稱內，總表只列持有成本做對照，標明兩層均未使用。",
    ]),
    ("等級標籤", [
        "[Verified]：可由一手來源直接核對。[Interested-party]：利害關係人自報（含公司指引）。[Analogy]：類比其他對象。[Assumed]：假設，須給低／基準／高。[Derived]：由其他輸入計算而得。",
        "等級欄依來源檔既有標示填寫；來源檔未標者依其性質歸類，並於輸入頁標明。NVIDIA Q1、Q2 實績依對帳檔引用 8-K 標為 Verified，本表建置時未重新核對 8-K 原文；Q3 107.6 標 Interested-party 且「未核對原文」。",
    ]),
    ("狀態色", [
        "綠＝一致；黃＝邊界（含「有條件一致」＝殘差在容差內但 NVIDIA 實績檢驗未過）；紅＝超出容差或檢驗未通過；灰＝待定（對帳中、連續法、尚無實績）。",
        "「容差判定」是本表依殘差 ÷ 基準與 ±15% 計算；「協調狀態」是 baseline_status.json 依規則 v1.4 判出的甲／乙／丙／待定，兩者來源不同，並列呈現。",
    ]),
    ("已知限制與需要留意之處", [
        (f"Siliconomics 快照 source.sha256 與 repo 內 {os.path.basename(P['sil_xlsx'])} 實算 SHA-256 不一致（輸入頁第 0 節與各列標紅）。本表依快照 JSON 的數值計算，數值未逐格比對全部，驗證紀錄只抽核部分名稱。" if s_snap["source"]["sha256"] != SHA["sil_xlsx"] else f"Siliconomics 快照 source.sha256 與 repo 內 {os.path.basename(P['sil_xlsx'])} 實算 SHA-256 一致。") + (" Wattonomics 兩者一致；" if w_snap["source"]["sha256"] == SHA["watt_xlsx"] else " Wattonomics 兩者不一致（標紅）；") + "Tokenomics 快照未記錄 SHA。",
        (f"協調橋接表 {BV} 的 In 頁標示 Siliconomics {BSV} 輸入，與本表讀的快照同版；第 9 節並列核對。" if BR_SAME else f"協調橋接表 {BV} 的 In 頁標示 Siliconomics {BSV} 輸入（2029 基準 {_br_s0_29:.3f}，{SV} 為 {_sn_b29:.3f}）。本表以 {SV} 快照重算，並於總表第 9 節與橋接表檔內值並列。"),
        "Siliconomics 快照陣列的年份對齊：長度 4 的陣列＝2026–2029；長度 5 的陣列（HBM、SoC 測試機、顆數係數）＝2025–2029。",
        "NVIDIA 公司值 2026 與 2027 含指引，只作參考；真正的公司實績只到 FY2027 Q1＋Q2。實績對帳 C 區（Q3＝107.6）為樣本外檢驗，未入模型，沒有任何其他格引用，也沒有用它擬合任何參數或季節權重；比值假設各季均勻。",
        "「公式引用次數」掃描只認具名引用；以儲存格位址間接引用的不計。IF_CapexTotal 在兩本模型都掃到 0 次，推測以 IT＋廠房相加取代，未逐格確認。",
        "Wattonomics 世代占比 2026–2029 為 Derived／Assumed 推導結果，直讀自模型檔 Players 頁；其中 2027 的 6% Rubin Ultra 等占比的來源為 Wattonomics 對 Siliconomics 顆數的推導，本表未重算。",
        "本表只讀、不改任何來源。所有數字由 build_summary.py 自檔案讀入；重跑只需更換路徑參數。",
    ]),
]
r = 2
for head, lines in NOTES:
    r += 1
    sec(ws, r, head, 1)
    for t in lines:
        r += 1
        put(ws, r, 1, t, al=AL_LT, border=False)
        ws.row_dimensions[r].height = 15 * max(1, (len(t) // 80) + 1)
widths(ws, [130])
page(ws, False, 1)

# 總表欄位保護：無
out = os.path.join(A.out, f"{A.date}_三層對帳總表_{A.ver}.xlsx")
os.makedirs(A.out, exist_ok=True)
wb.save(out)

# 供 verify 使用的位置登記
meta = dict(out=out, REG=REG, ROW=ROW, OOS_ROWS=OOS_CELLS, ACT=ACT, SHA=SHA, GIT_CM=GIT_CM, GIT_TK=GIT_TK,
            files={k: dict(file=v["file"], rec=v["rec"], act=v["act"]) for k, v in FILES.items()},
            q1=NV_Q1, q2=NV_Q2, usage_w=USE_W, usage_s=USE_S)
with open(os.path.join(A.out, ".build_meta.json"), "w", encoding="utf-8") as f:
    json.dump(meta, f, ensure_ascii=False, indent=1)
print("wrote", out)
