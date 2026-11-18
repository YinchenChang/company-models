#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""驗證三層對帳總表：重算後抽核格子與來源檔一致，掃錯誤值，掃 Q3 樣本外隔離。輸出 驗證紀錄.md。
所有「期望值」都直接重新讀來源檔（不沿用 build_summary.py 的讀法）。"""
import hashlib
import json
import os
import re
import subprocess
import sys

import openpyxl

OUT = "/home/claude/out/summary"
CM = "/home/claude/company-models"
TKR = "/home/claude/tokenomics"
XLSX = f"{OUT}/20261118_三層對帳總表_v0.1.xlsx"
meta = json.load(open(f"{OUT}/.build_meta.json", encoding="utf-8"))
ROW, REG = meta["ROW"], meta["REG"]

# 1. LibreOffice 重算（在副本上），讀快取值
import shutil, tempfile
tmp = tempfile.mkdtemp()
shutil.copy(XLSX, f"{tmp}/chk.xlsx")
subprocess.run(["soffice", "--headless", "--convert-to", "xlsx:Calc MS Excel 2007 XML", "--outdir", f"{tmp}/o", f"{tmp}/chk.xlsx"],
               capture_output=True, timeout=180)
wbv = openpyxl.load_workbook(f"{tmp}/o/chk.xlsx", data_only=True)
wbf = openpyxl.load_workbook(XLSX)  # 公式

res = []


def rec(name, ok, detail):
    res.append((name, bool(ok), detail))


# 2. 錯誤值掃描與公式計數
errs, nform = [], 0
for ws in wbv:
    for row in ws.iter_rows():
        for c in row:
            if isinstance(c.value, str) and c.value.startswith("#") or (isinstance(c.value, str) and c.value in ("Err:502", "Err:504", "Err:508")):
                errs.append(f"{ws.title}!{c.coordinate}={c.value}")
for ws in wbf:
    for row in ws.iter_rows():
        for c in row:
            if isinstance(c.value, str) and c.value.startswith("="):
                nform += 1
rec("重算後無錯誤值", not errs, f"公式 {nform} 個；錯誤值 {len(errs)} 個 {errs[:5]}")

S = wbv["總表"]
SC = {2025: "D", 2026: "E", 2027: "F", 2028: "G", 2029: "H"}


def sv_(key, y):
    return S[f"{SC[y]}{ROW[key]}"].value


# 3. 來源（直接重讀）
w = json.load(open(f"{CM}/wattonomics/data/wattonomics_snapshot_v0.3.6.json", encoding="utf-8"))["names"]
s = json.load(open(f"{CM}/siliconomics/data/siliconomics_snapshot_v0.20.json", encoding="utf-8"))["names"]
st = json.load(open(f"{CM}/siliconomics/data/baseline_status.json", encoding="utf-8"))
rc = json.load(open(f"{CM}/siliconomics/data/recon_snapshot_v0.20.json", encoding="utf-8"))
tk = json.load(open(f"{CM}/wattonomics/data/tokenomics_snapshot_v22.0.json", encoding="utf-8"))
br = openpyxl.load_workbook(f"{CM}/siliconomics/docs/reports/20261116_協調橋接_v0.3.6.xlsx", data_only=True)["Bridge"]
brow = {br.cell(r, 1).value: r for r in range(4, br.max_row + 1) if br.cell(r, 1).value}
tkx = openpyxl.load_workbook(f"{TKR}/model/20261105_Tokenomics_v22.0.xlsx", data_only=True)
wx = openpyxl.load_workbook(f"{CM}/wattonomics/model/20261116_Wattonomics_v0.3.6.xlsx", data_only=True)
sx = openpyxl.load_workbook(f"{CM}/siliconomics/model/20261116_Siliconomics_v0.20.xlsx", data_only=True)


def close(a, b, tol=1e-6):
    return isinstance(a, (int, float)) and isinstance(b, (int, float)) and abs(a - b) <= tol * max(1, abs(b))


checks = []   # (描述, 格, 實際, 期望)


def chk(desc, cell, actual, expected, tol=1e-6):
    checks.append((desc, cell, actual, expected, close(actual, expected, tol) if not isinstance(expected, str) else actual == expected))


for key, nm, y in [("w_add", "WIF_MktGW_Add", 2027), ("w_add", "WIF_MktGW_Add", 2025), ("w_avg", "WIF_MktGW_Avg", 2029), ("w_ye", "WIF_MktGW_YE", 2028),
                   ("w_it", "WIF_Spend_it_sum", 2029), ("w_tot", "WIF_Spend_total", 2026), ("w_short", "WIF_Gap_Short_A_D0", 2027)]:
    chk(f"Wattonomics {nm}_{y}", f"總表!{SC[y]}{ROW[key]}", sv_(key, y), w[f"{nm}_{y}"])
sb = s["SIF_DeployGW_Base"]
chk("Siliconomics SIF_DeployGW_Base 2029", f"總表!H{ROW['s_base']}", sv_("s_base", 2029), sb[3])
chk("Siliconomics SIF_DeployGW_Base 2027", f"總表!F{ROW['s_base']}", sv_("s_base", 2027), sb[1])
chk("Siliconomics SIF_DeployGW_Supply_2028", f"總表!G{ROW['s_sup']}", sv_("s_sup", 2028), s["SIF_DeployGW_Supply_2028"])
chk("Siliconomics SIF_AIITCapex_Base 2027", f"總表!F{ROW['s_it']}", sv_("s_it", 2027), s["SIF_AIITCapex_Base"][1])
# 欄位語義
SEM = wbv["欄位語義"]
chk("SIF_DeployGW_Current_2029（應為 51.118）", "欄位語義!E9", SEM["E9"].value, s["SIF_DeployGW_Current_2029"])
chk("SIF_DeployGW_Alt_2027", "欄位語義!C8", SEM["C8"].value, s["SIF_DeployGW_Alt_2027"])
chk("SIF_DeployGW_Demand_2026", "欄位語義!B7", SEM["B7"].value, s["SIF_DeployGW_Demand_2026"])
chk("Current − Base 2029", "欄位語義!E11", SEM["E11"].value, s["SIF_DeployGW_Current_2029"] - sb[3])
# 證據比例
chk("Wattonomics 證據比例（低）2027＝baseline_status（3 位小數）", f"總表!F{ROW['e_st_ws_lo']}", sv_("e_st_ws_lo", 2027), st["years"]["2027"]["wattonomics"]["evidence_share"]["lo"])
chk("　同上 vs Wattonomics 快照 WIF_Evidence_Lo_2027（狀態檔四捨五入，容許 5e-4）", f"總表!F{ROW['e_st_ws_lo']}", sv_("e_st_ws_lo", 2027), w["WIF_Evidence_Lo_2027"], 5e-4)
chk("Siliconomics 證據比例（高）2028", f"總表!G{ROW['e_st_ss_hi']}", sv_("e_st_ss_hi", 2028), st["years"]["2028"]["siliconomics"]["evidence_share"]["hi"])
# 獨立重算差距
for y in (2026, 2027, 2028, 2029):
    i = y - 2026
    phys = s["SIF_DeployGWPhys_Base"][i]
    base = sb[i]
    rt = phys / base
    add = w[f"WIF_MktGW_Add_{y}"]
    nxt = w[f"WIF_MktGW_Add_{y + 1}"]
    k = 4.5 / 12 + 0.1
    x0 = add / rt
    x2 = (add + k * (nxt - add)) / rt
    chk(f"差距（獨立重算）{y}", f"總表!{SC[y]}{ROW['gap_gw']}", sv_("gap_gw", y), base - x0)
    chk(f"殘差（獨立重算）{y}", f"總表!{SC[y]}{ROW['g_resid']}", sv_("g_resid", y), base - x2)
# 與橋接表值
for y, col in ((2027, "F"), (2028, "G")):
    chk(f"橋接表殘差 g3 {y} vs 本表", f"總表!{SC[y]}{ROW['g_resid']}", sv_("g_resid", y), br[f"{col}{brow['g3']}"].value, 1e-6)
chk("橋接表總差距 2026 vs 本表", f"總表!E{ROW['gap_gw']}", sv_("gap_gw", 2026), br[f"E{brow['g']}"].value, 1e-6)
chk("橋接表殘差 2029（容許 v0.19→v0.20 差 0.03）", f"總表!H{ROW['g_resid']}", sv_("g_resid", 2029), br[f"H{brow['g3']}"].value, 2.5e-3)
# Tokenomics 加權 vs Wattonomics CapexGW 第 7 列
cg = wx["CapexGW"]
for y, col in ((2026, "G"), (2027, "H"), (2029, "J")):
    chk(f"Tokenomics 加權 IT capex {y} vs Wattonomics CapexGW!{col}7", f"總表!{SC[y]}{ROW['tk_it']}", sv_("tk_it", y), cg[f"{col}7"].value, 1e-6)
# Tokenomics 具名範圍
dn = tkx.defined_names["IF_HoldEcon"].attr_text.split("!")
ws_ = tkx[dn[0]]
vr_hold = ws_["M13"].value
IN = wbv["輸入"]
found = None
for r in range(1, IN.max_row + 1):
    if IN.cell(r, 1).value == "IF_HoldEcon|VR200":
        found = IN.cell(r, 5).value
chk("IF_HoldEcon VR200 基準（工作簿 Interface!M13）", "輸入!IF_HoldEcon|VR200", found, vr_hold)
found = None
for r in range(1, IN.max_row + 1):
    if IN.cell(r, 1).value == "IF_CapexIT|RU":
        found = IN.cell(r, 6).value
chk("IF_CapexIT Rubin Ultra 高成本（Interface!Q9）", "輸入!IF_CapexIT|RU", found, ws_["Q9"].value if False else tkx["Interface"]["Q9"].value)
# 實績
ACT = wbv["實績對帳"]
nv = rc["markets"]["nv_dc"]
chk("Q1＋Q2 累計＝對帳檔 actual_cum", "實績對帳!D14", ACT["D14"].value, nv["actual_cum"][0])
chk("主讀數（Siliconomics）＝對帳檔 ratio_main", "實績對帳!B18", ACT["B18"].value, nv["ratio_main"][0], 1e-9)
chk("第二讀數（Siliconomics）＝對帳檔 ratio_second", "實績對帳!B20", ACT["B20"].value, nv["ratio_second"][0], 1e-9)
chk("Wattonomics 隱含營收 2026", "實績對帳!C7", ACT["C7"].value, w["WIF_Recon_NV_ImpliedRev_2026"])
chk("Wattonomics 2025 比值 vs WIF_Recon_NV_Ratio_2025", "實績對帳!G6", ACT["G6"].value, w["WIF_Recon_NV_Ratio_2025"], 1e-9)
# 13 個對帳點
COV = wbv["介面覆蓋"]
mk = rc["markets"]
pts = [k for k, v in mk.items() if v["kind"] in ("market", "structural")]
meth = {"個股實績加總": 0, "結構一致": 0, "未建": 0}
for r in range(1, COV.max_row + 1):
    v = COV.cell(r, 3).value
    if COV.cell(r, 1).value and re.match(r"^\d\d ", str(COV.cell(r, 1).value)):
        meth[v] += 1
chk("13 對帳點：個股實績加總 2／結構一致 7／未建 4", "介面覆蓋!C 欄", f"{meth['個股實績加總']}/{meth['結構一致']}/{meth['未建']}", "2/7/4")
chk("對帳點數量＝13", "介面覆蓋", len(pts), 13)

ok_n = sum(1 for c in checks if c[4])
for c in checks:
    rec(c[0], c[4], f"{c[1]}：實際 {c[2]}；來源 {c[3]}")

# 4. SHA
shas = {}
for r in range(1, IN.max_row + 1):
    code = IN.cell(r, 1).value
    if code in ("TKS", "TKX", "WAS", "WAX", "SIS", "SIX", "SIB", "SIR", "BRG") and IN.cell(r, 5).value and len(str(IN.cell(r, 5).value)) == 64:
        shas[code] = (IN.cell(r, 2).value.split("\n")[0], IN.cell(r, 5).value, IN.cell(r, 16).value, IN.cell(r, 11).value, IN.cell(r, 13).value)
bad = []
for code, (rel, sha, flag, recd, act) in shas.items():
    p = rel if rel.startswith("/") else (os.path.join(TKR, rel) if rel.startswith("model/2026") and "Tokenomics" in rel else os.path.join(CM, rel))
    h = hashlib.sha256(open(p, "rb").read()).hexdigest()
    if h != sha:
        bad.append(code)
rec("輸入頁檔案清單 SHA-256 與重算一致（9 個檔案）", not bad and len(shas) == 9, f"檢查 {len(shas)} 個；不符 {bad}")
sil = shas.get("SIS")
rec("Siliconomics 快照記錄 SHA 與 repo 實算不一致已如實顯示", sil and sil[2] == "不一致" and sil[3].startswith("cf29d2bb") and sil[4].startswith("ad8dfaaf"), f"快照記錄 {sil[3][:8]}、實算 {sil[4][:8]}、比對欄＝{sil[2]}")
wat = shas.get("WAS")
rec("Wattonomics 快照記錄 SHA 與實算一致", wat and wat[2] == "一致", f"比對欄＝{wat[2]}")

# 5. 快照 vs 工作簿（Siliconomics、Wattonomics 少數名稱；Tokenomics 全部 23 名稱基準）
ISI = sx["Interface"]
rows_s = {ISI.cell(r, 1).value: r for r in range(1, ISI.max_row + 1) if ISI.cell(r, 1).value}
mm = []
r_ = rows_s["SIF_DeployGW_Base"]
for i, col in enumerate("EFGH"):
    xv = ISI[f"{col}{r_}"].value
    xv = float(xv) if xv is not None else None
    mm.append(abs(xv - sb[i]) if xv is not None else None)
for nm in ("SIF_DeployGW_Supply_2028", "SIF_DeployGW_Current_2029", "SIF_DeployGW_Alt_2029", "SIF_DeployGW_Demand_2027"):
    xv = ISI[f"E{rows_s[nm]}"].value
    mm.append(abs(float(xv) - s[nm]))
rec("Siliconomics 快照數值 vs repo 內活頁簿 Interface 快取值（8 格；SHA 不一致下的數值抽核）", all(m is not None and m < 1e-9 for m in mm), f"最大差 {max(m for m in mm if m is not None):.2e}")
WI = wx["Interface"]
rows_w = {WI.cell(r, 1).value: r for r in range(1, WI.max_row + 1) if WI.cell(r, 1).value}
mmw = []
for nm, col in (("WIF_MktGW_Add", "G"), ("WIF_Spend_it_sum", "G"), ("WIF_Spend_total", "I")):
    if nm in rows_w:
        y = {"G": 2027, "I": 2029}[col]
        mmw.append(abs(float(WI[f"{col}{rows_w[nm]}"].value) - w[f"{nm}_{y}"]))
rec("Wattonomics 快照數值 vs 活頁簿 Interface 快取值", mmw and all(m < 1e-6 for m in mmw), f"檢查 {len(mmw)} 格；最大差 {max(mmw) if mmw else 'n/a'}")
# Tokenomics 快照 vs 工作簿
maxd = 0
n = 0
for nm, it in tk["items"].items():
    t = tkx.defined_names[nm].attr_text
    sh, rng = t.split("!")
    cells = tkx[sh.strip("'")][rng.replace("$", "")][0]
    xv = [c.value for c in cells]
    if isinstance(it["values"], dict):
        gens = list(it["values"].keys())
        for gi, g in enumerate(gens):
            for si, sc in enumerate(["低成本", "基準", "高成本"]):
                if g in it["values"] and isinstance(it["values"][g], dict) and sc in it["values"][g] and len(xv) >= 15:
                    a = it["values"][g][sc]
                    b = xv[gi * 3 + si]
                    if isinstance(a, (int, float)) and isinstance(b, (int, float)):
                        maxd = max(maxd, abs(a - b))
                        n += 1
rec("Tokenomics 快照（gen_cost 類）vs v22.0 工作簿具名範圍", n > 50 and maxd < 1e-9, f"比對 {n} 格；最大差 {maxd:.2e}")

# 6. 樣本外隔離
oos_cells = []
Q3 = 107.6
q3_row = REG["nv_oos_q3"]
oq_lo, oq_hi = meta["OOS_ROWS"]
viol = []
for wsn in wbf.sheetnames:
    ws = wbf[wsn]
    for row in ws.iter_rows():
        for c in row:
            v = c.value
            if isinstance(v, (int, float)) and abs(v - Q3) < 1e-9:
                oos_cells.append(f"{wsn}!{c.coordinate}")
            if isinstance(v, str) and v.startswith("="):
                # 引用輸入頁 Q3 列
                if re.search(rf"'輸入'!\$?E\$?{q3_row}\b", v) and not (wsn == "實績對帳" and oq_lo <= c.row <= oq_hi):
                    viol.append(f"{wsn}!{c.coordinate}")
                # 其他頁引用實績對帳 C 區
                if wsn != "實績對帳" and re.search(r"實績對帳", v):
                    viol.append(f"{wsn}!{c.coordinate}")
                # 實績對帳 B 區以前（Q3 區之前）引用 Q3 區儲存格
                if wsn == "實績對帳" and c.row < oq_lo:
                    v_local = re.sub(r"'[^']+'!\$?[A-Z]+\$?\d+", "", v)
                    for m_ in re.finditer(r"\$?[A-H]\$?(\d+)", v_local):
                        if int(m_.group(1)) >= oq_lo:
                            viol.append(f"{wsn}!{c.coordinate}")
# Q3 區內公式只可引用 B 區（rows<oq_lo）或自身區塊，此外不得被區外引用（上面已查）
rec("Q3＝107.6 只作為輸入出現一次（輸入頁橘色列）", oos_cells == [f"輸入!E{q3_row}"], f"出現位置 {oos_cells}")
rec("沒有任何區外公式引用 Q3 列或樣本外區塊", not viol, f"違規 {viol}")
# 模型內實績維持 Q1＋Q2
rec("模型內實績維持 75.2、89.0（Q1、Q2）", close(ACT["D12"].value, 75.2) and close(ACT["D13"].value, 89.0), f"D12={ACT['D12'].value}, D13={ACT['D13'].value}")
# 字型與顏色規則
blue_bad, black_bad = [], []
for wsn in wbf.sheetnames:
    ws = wbf[wsn]
    for row in ws.iter_rows():
        for c in row:
            v = c.value
            col = c.font.color.rgb if c.font and c.font.color and c.font.color.type == "rgb" else None
            if isinstance(v, str) and v.startswith("=") and col == "000000FF":
                blue_bad.append(f"{wsn}!{c.coordinate}")
            if isinstance(v, (int, float)) and wsn in ("總表", "實績對帳") and col != "000000FF":
                black_bad.append(f"{wsn}!{c.coordinate}")
rec("公式格無藍字；總表與實績對帳頁沒有手打常數數字", not blue_bad and not black_bad, f"公式藍字 {blue_bad[:3]}；常數數字 {black_bad[:5]}")
# 凍結窗格與列印
fz = {n: wbf[n].freeze_panes for n in wbf.sheetnames}
rec("凍結窗格（總表、實績對帳、介面覆蓋、欄位語義、輸入）", all(fz[n] for n in ("總表", "實績對帳", "介面覆蓋", "欄位語義", "輸入")), str(fz))
rec("每頁設定橫向／整頁寬列印", all(wbf[n].page_setup.fitToWidth == 1 for n in wbf.sheetnames), "fitToWidth=1 全部分頁")

# 輸出
passed = sum(1 for r in res if r[1])
lines = ["# 三層對帳總表 驗證紀錄", "", "- 檔案：`/home/claude/out/summary/20261118_三層對帳總表_v0.1.xlsx`",
         f"- 來源：company-models main {meta['GIT_CM'][0][:7]}（{meta['GIT_CM'][1]}）、tokenomics main {meta['GIT_TK'][0][:7]}（{meta['GIT_TK'][1]}）",
         f"- 重算：LibreOffice headless（在副本上）；公式 {nform} 個，錯誤值 {len(errs)} 個",
         f"- 結果：{passed} / {len(res)} 項通過；其中抽核格子 {ok_n} / {len(checks)} 格與來源檔一致（來源檔為重新直讀，不沿用建檔程式的讀法）", "",
         "| # | 檢查項 | 結果 | 明細 |", "|---|---|---|---|"]
for i, (n_, ok, d) in enumerate(res, 1):
    lines.append(f"| {i} | {n_} | {'通過' if ok else '**未通過**'} | {d.replace('|', '／')} |")
lines += ["", "## 沒有驗證的事", "",
          "- 沒有重核 NVIDIA 8-K 原文（Q1 75.2、Q2 89.0 取自對帳快照的引用文字）；Q3 107.6 為 Andy 提供、未核對原文，也不在任何來源檔內，故無法對檔驗證，只驗證它的隔離。",
          "- Siliconomics 活頁簿與快照 SHA 不一致，只抽核 8 個名稱的數值相同，沒有逐格比對全部名稱。",
          "- 「公式引用次數」只掃具名引用，未追蹤以儲存格位址的間接引用；掃描結果不等於「模型沒用到」。",
          "- Wattonomics 世代占比讀自模型檔 Players 頁（快照為空白），未重算其推導；僅以 Wattonomics CapexGW 第 7 列交叉核對加權結果。",
          "- 只在 LibreOffice 重算；未在 Excel 開啟。"]
open(f"{OUT}/驗證紀錄.md", "w", encoding="utf-8").write("\n".join(lines) + "\n")
print(f"{passed}/{len(res)} passed; cells {ok_n}/{len(checks)}")
for n_, ok, d in res:
    if not ok:
        print("FAIL:", n_, d)
sys.exit(0 if passed == len(res) else 1)
