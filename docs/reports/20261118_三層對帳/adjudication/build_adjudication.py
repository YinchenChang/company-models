# 2026 年對帳單期限裁定材料（期限 2026-11-30）的試算表產生程式。只讀輸入、不改任何模型。
# 用法：python3 build_adjudication.py <輸出.xlsx>
# 裁定方法照《三層基準協調規則 v1.5 草案》C9 第 6 點（事前寫死）：只比最高一級證據（已公布實績）；
# 期間對齊首選＝季節權重（歷史三個已完結會計年度；低＝最小、基準＝中位數、高＝最大），次佳＝等全年，第三＝線性。
# NVIDIA FY2027 第三季實績只在本檔作「樣本外檢驗，未入模型」的裁定證據。
import sys
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment

BLUE = Font(color="0000FF"); BOLD = Font(bold=True); OOS = PatternFill("solid", fgColor="FCE4D6")
wb = openpyxl.Workbook()
I = wb.active; I.title = "輸入"
I.append(["鍵", "項目", "值", "單位", "等級", "來源"])
rows = [
 ("S26", "供給側讀數：Siliconomics v0.20 隱含 NVIDIA 資料中心營收 2026", 448.3617, "US$B", "Derived", "siliconomics/data/recon_snapshot_v0.20.json nv_dc.implied[0]（Recon F 節）"),
 ("W26", "需求側讀數：Wattonomics v0.3.6 隱含 NVIDIA 資料中心營收 2026", 352.6705, "US$B", "Derived", "wattonomics_snapshot_v0.3.6.json WIF_Recon_NV_ImpliedRev_2026"),
 ("supGW", "供給側部署 GW 2026（＝現行基準）", 23.3995, "GW", "Derived", "siliconomics/data/baseline_status.json years.2026.supply_side"),
 ("convGW", "需求側換算值 2026", 20.972, "GW", "Derived", "baseline_status.json years.2026.converted"),
 ("q1", "NVIDIA 資料中心營收 FY27 Q1", 75.2, "US$B", "Verified", "NVIDIA 新聞稿；nvidia 個股模型 Q!G5"),
 ("q2", "NVIDIA 資料中心營收 FY27 Q2", 89.0, "US$B", "Verified", "NVIDIA 新聞稿 2026-08-26；nvidia 個股模型 Q!H5"),
 ("q3", "【樣本外檢驗，未入模型】NVIDIA 資料中心營收 FY27 Q3", 107.6, "US$B", "Verified", "NVIDIA 新聞稿 2026-11-17"),
 ("fy24", "NVIDIA 資料中心 FY2024 全年", 47.5, "US$B", "Verified", "FY2024 Q4 新聞稿"),
 ("fy24q3", "　FY24 Q3", 14.51, "US$B", "Verified", "同上"),
 ("fy24q4", "　FY24 Q4", 18.4, "US$B", "Verified", "同上"),
 ("fy25", "NVIDIA 資料中心 FY2025 全年", 115.2, "US$B", "Verified", "FY2025 Q4 新聞稿"),
 ("fy25q3", "　FY25 Q3", 30.8, "US$B", "Verified", "同上"),
 ("fy25q4", "　FY25 Q4", 35.6, "US$B", "Verified", "同上"),
 ("fy26q1", "NVIDIA 資料中心 FY26 Q1", 39.1, "US$B", "Verified", "nvidia 個股模型 Q!B5"),
 ("fy26q2", "　FY26 Q2", 41.1, "US$B", "Verified", "Q!C5"),
 ("fy26q3", "　FY26 Q3", 51.2, "US$B", "Verified", "Q!D5"),
 ("fy26q4", "　FY26 Q4", 62.3, "US$B", "Verified", "Q!E5"),
 ("tol", "容差（規則 C2）", 0.15, "比", "規則", "docs/plan/三層基準協調規則.md C2"),
]
ref = {}
for r in rows:
    I.append(list(r)); ref[r[0]] = f"輸入!$C${I.max_row}"; I.cell(I.max_row, 3).font = BLUE
    if r[0] == "q3":
        for c in range(1, 7): I.cell(I.max_row, c).fill = OOS
for c, w in zip("ABCDEF", (9, 62, 11, 8, 10, 70)): I.column_dimensions[c].width = w

W = wb.create_sheet("權重")
W.append(["歷史會計年度（皆已完結；FY2027 任何季度不進權重）", "上半年 ÷ 全年", "前三季 ÷ 全年"])
W.append(["FY2024", f"=({ref['fy24']}-{ref['fy24q3']}-{ref['fy24q4']})/{ref['fy24']}", f"=({ref['fy24']}-{ref['fy24q4']})/{ref['fy24']}"])
W.append(["FY2025", f"=({ref['fy25']}-{ref['fy25q3']}-{ref['fy25q4']})/{ref['fy25']}", f"=({ref['fy25']}-{ref['fy25q4']})/{ref['fy25']}"])
s26 = f"({ref['fy26q1']}+{ref['fy26q2']}+{ref['fy26q3']}+{ref['fy26q4']})"
W.append(["FY2026", f"=({ref['fy26q1']}+{ref['fy26q2']})/{s26}", f"=({ref['fy26q1']}+{ref['fy26q2']}+{ref['fy26q3']})/{s26}"])
W.append(["低（三年最小）", "=MIN(B2:B4)", "=MIN(C2:C4)"])
W.append(["基準（三年中位數）", "=MEDIAN(B2:B4)", "=MEDIAN(C2:C4)"])
W.append(["高（三年最大）", "=MAX(B2:B4)", "=MAX(C2:C4)"])
W.append(["線性（對照，第三序位）", 0.5, 0.75])
W.append(["等級 [Derived]：由 [Verified] 實績相除；「當年分布像過去」是延伸假設。取法（最小／中位數／最大）為機械規則，寫定於 K1 回測設計與 C9 第 6 點草案。"])
W.column_dimensions["A"].width = 46; W.column_dimensions["B"].width = 16; W.column_dimensions["C"].width = 16

A = wb.create_sheet("裁定")
A.append(["證據組", "已公布實績（US$B）", "對齊", "權重", "供給側同期值", "需求側同期值", "供給側偏差 e_S", "需求側偏差 e_W", "|e_S|", "|e_W|", "兩者都在 ±15% 內", "偏差絕對值較小者", "依草案打平定義的結果"])
ev = [("FY27 Q1＋Q2（樣本內）", f"={ref['q1']}+{ref['q2']}", "B", False), ("FY27 Q1–Q3（含樣本外 Q3；未入模型）", f"={ref['q1']}+{ref['q2']}+{ref['q3']}", "C", True)]
al = [("季節權重 基準【主讀數】", 6), ("季節權重 低", 5), ("季節權重 高", 7), ("線性（對照）", 8)]
for name, f, col, oos in ev:
    for an, wr in al:
        r = A.max_row + 1
        A.append([name, f, an, f"=權重!{col}{wr}", f"={ref['S26']}*D{r}", f"={ref['W26']}*D{r}", f"=E{r}/B{r}-1", f"=F{r}/B{r}-1", f"=ABS(G{r})", f"=ABS(H{r})",
                  f"=IF(AND(I{r}<={ref['tol']},J{r}<={ref['tol']}),\"是\",\"否\")", f"=IF(ROUND(I{r},3)=ROUND(J{r},3),\"相同\",IF(I{r}<J{r},\"供給側\",\"需求側\"))",
                  f"=IF(K{r}=\"是\",\"打平→維持現行\",L{r}&\"勝出\")"])
        if oos:
            for c in range(1, 14): A.cell(r, c).fill = OOS
        if "主讀數" in an:
            for c in range(1, 14): A.cell(r, c).font = BOLD
A.append([])
A.append(["e＝層級同期值 ÷ 公司實績 − 1（正＝層級高估）。主讀數＝季節權重基準、證據取到期時最完整的已公布期間（Q1–Q3）。橘底＝使用樣本外保留實績（只作裁定證據，未入模型）。"])
for c, w in zip("ABCDEFGHIJKLM", (36, 14, 24, 9, 13, 13, 13, 13, 8, 8, 14, 16, 22)): A.column_dimensions[c].width = w
for row in A.iter_rows(min_row=2, max_row=9):
    for c in row[3:4] + row[6:10]: c.number_format = "0.0%"
    for c in row[4:6]: c.number_format = "0.0"
for row in W.iter_rows(min_row=2, max_row=8):
    for c in row[1:3]: c.number_format = "0.0%"

P = wb.create_sheet("影響")
P.append(["情形", "2026 基準（GW）", "對現行差異（GW）", "差異％", "說明"])
P.append(["現行（供給側讀數；裁定「打平→維持」）", f"={ref['supGW']}", "=B2-$B$2", "=B2/$B$2-1", "不改任何值；回復方案不適用"])
P.append(["反事實：需求側勝出", f"={ref['convGW']}", "=B3-$B$2", "=B3/$B$2-1", "只有在「不採 ±15% 打平定義」且「只用 Q1＋Q2」時才會出現；62 家硬體個股 2026 年讀數同比例下修（個股目標價未實算）"])
P.column_dimensions["A"].width = 40; P.column_dimensions["E"].width = 90
for c in ("B", "C"): P.column_dimensions[c].width = 16
P["D2"].number_format = P["D3"].number_format = "0.0%"
wb.save(sys.argv[1])
