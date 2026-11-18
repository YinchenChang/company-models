# 產生影響分析 xlsx：情境逐年、個股逐家（每個跑過 cascade 的情境一頁）、評等變動、方法選項。輸入藍字、計算黑字公式。
import json, sys, os
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter as CL
R = "/home/claude/out/proposal/runs"; V = "/home/claude/out/verify"
OUT = sys.argv[1]
BLUE, BOLD = Font(color="0000FF"), Font(bold=True); HDR = PatternFill("solid", fgColor="DDEBF7")
wb = openpyxl.Workbook()
# ---------- 說明 ----------
ws = wb.active; ws.title = "說明"
for i, t in enumerate([
    "自建 NVIDIA 顆數入模：影響分析（proposal 代理，2026-11-18；唯讀試算，未提交、未改 repo）",
    "來源：company-models main 2b2a962c（Siliconomics v0.20 r2、Wattonomics v0.3.7）；#661 分支 claude/3layer-w2-sil-v1（build.py 修補套到 main）；各情境在 /tmp 複本重建 Siliconomics 兩遍（第二遍讀新 Recon 快照），LibreOffice 重算。",
    "藍字＝自試算活頁簿讀出的值（輸入）；黑字＝本檔公式。情境定義見 Scen 頁 B 欄。",
    "x2 固定＝In 頁 crd_x2（落後一版的設計值，合併時實際會用的值）；x2 收斂＝以該情境 Siliconomics 重跑 baseline_bridge.py（rt 隨顆數組合改變）的診斷值，未回寫。",
    "評等門檻（build_summary.py，Andy 2026-10-25）：> +50% 強力買進；+25%～+50% 買進；−25%～+25% 中立；−50%～−25% 放空；< −50% 強力放空。現價＝改前目標價 ÷（1＋改前上檔空間）。",
    "NVIDIA FY2027 第三季實績未用於任何情境、參數或檢驗。",
    "S1、S2 個股頁引用獨立核對（verify）代理的 cascade 結果（tp_662_r2、tp_comb_r2），本代理未重跑；S1、S2 的 GW 由本代理重建（與核對報告一致）。"], 1):
    ws.cell(i, 1, t)
ws.column_dimensions["A"].width = 160
# ---------- Scen ----------
SC = [("S0", "main 現狀（對照）"), ("S1", "只 #662（Rubin 7.0→6.7、重算截頂）"), ("S2", "#661＋#662"),
      ("S3a", "2027 NVIDIA＝613 萬（#662 組合等比例、截頂重算＝1）；2028 維持 0.8"), ("S3b", "同 S3a；2028 連續法（#661）"),
      ("S3m_a", "S3a 但用 main 組合（Rubin 7.0、含重複的 Rubin Ultra）"),
      ("S3split_a", "S3a 但世代拆分 1.40／4.09／0.10 比例；0.8"), ("S3split_b", "同上；連續法"),
      ("S4a", "S3a＋2026 NVIDIA 合計＝620 萬（含 Hopper，等比例；工作單字面）；0.8"), ("S4b", "同 S4a；連續法"),
      ("S4h_a", "S3a＋2026 Blackwell＋Rubin＝620 萬、Hopper 0.5 不動（與「2,000 萬晶粒」同口徑）；0.8"), ("S4h_b", "同 S4h_a；連續法"),
      ("S5L_a", "2027 NVIDIA＝400 萬（低端）；0.8"), ("S5L_b", "同上；連續法"), ("S5H_a", "2027 NVIDIA＝839 萬（高端）；0.8"), ("S5H_b", "同上；連續法")]
ws = wb.create_sheet("Scen")
ws.cell(1, 1, "情境逐年表（GW：成本等值、出貨年；藍字＝試算活頁簿讀值，黑字＝公式）").font = BOLD
ws.cell(2, 1, "公司指引推得 2027 資料中心營收（Recon E96；US$B）"); ws.cell(2, 5, 633.7).font = BLUE
ws.cell(3, 1, "自建報告 A1 推得 2027 曆年資料中心營收（US$B）"); ws.cell(3, 5, 626.1).font = BLUE
ws.cell(4, 1, "NVIDIA FY27 Q1＋Q2 資料中心實績（US$B）"); ws.cell(4, 5, 164.2).font = BLUE
ws.cell(5, 1, "每 IT GW NVIDIA 營收 n2 2027（R4 橋接；US$B／GW）"); ws.cell(5, 5, 26.127).font = BLUE
cols = []
def add(h, key, kind="in"): cols.append((h, key, kind))
for y in range(4): add(f"基準 {2026+y}", ("base", y))
for y in range(4): add(f"x2 固定 {2026+y}", ("dem", y))
for y in range(4): add(f"x2 收斂 {2026+y}", ("bridge.x2", y))
for y in range(4): add(f"差距（基準−x2 固定）{2026+y}", ("gapf", y), "f")
add("2027 差距（基準−x2 收斂）", ("gapc", 1), "f")
for y in range(4): add(f"基準對 S0 {2026+y}", ("d0", y), "f")
for y in range(4): add(f"實體 GW {2026+y}", ("phys", y))
for y in range(4): add(f"供給側讀數 {2026+y}", ("sup", y))
for y in range(4): add(f"crd_k {2026+y}", ("k", y))
for y in range(4): add(f"狀態 {2026+y}", ("st", y))
for y in range(4): add(f"標示 {2026+y}", ("lab", y))
add("NVIDIA 顆數 2026（百萬封裝）", ("units.2026.NV", None)); add("NVIDIA 顆數 2027（截頂後）", ("units.2027.NV", None))
add("2027 HBM 截頂係數", ("hbmcoef", None)); add("2027 HBM 需求÷供給", ("hbm27.util", None)); add("HBM 市場 2027（US$B）", ("hbmmkt27", None))
add("隱含 NV 資料中心營收 2026", ("reconF.implied_rev", 0)); add("隱含 NV 資料中心營收 2027", ("reconF.implied_rev", 1))
add("2027 隱含÷指引 633.7 − 1", ("r633", None), "f"); add("2027 隱含÷自建 626.1 − 1", ("r626", None), "f")
add("2027 所需每 GW 內容（指引÷NV 成本等值 GW）", ("need", None), "f")
add("2026 主讀數（實績÷層級同期間−1）", ("main26", None), "f"); add("2026 第二讀數（Q2×4÷全年−1）", ("reconF.second_reading", None))
add("2026 Recon F 狀態", ("reconF.status", None)); add("2026 季節讀數（基準權重）", ("reconF.seasonal", 1))
for y in range(4): add(f"Sil 證據比例重量測 {2026+y}（診斷）", ("ev.es", y))
add("2027 證據比例重量測 低", ("ev.lo", 1)); add("2027 證據比例重量測 高", ("ev.hi", 1))
add("SGOV_Errors（cascade 前；K1 自舉）", ("SGOV_Errors", None)); add("LibreOffice 錯誤格", ("err_cells", None))
H = 7
ws.cell(H, 1, "情境").font = BOLD; ws.cell(H, 2, "定義").font = BOLD
for j, (h, k, kind) in enumerate(cols, 3):
    c = ws.cell(H, j, h); c.font = BOLD; c.fill = HDR; c.alignment = Alignment(wrap_text=True)
CIX = {k: j for j, (h, k, kind) in enumerate(cols, 3)}
def get(m, key, idx):
    if key == "hbmcoef":
        u = m["units"]["2027"]; return 1.0 if m["hbm27"]["util"] < 0.9999 else min(1.0, 1/ (m["hbm27"]["demand_EB_after_units"]/m["hbm27"]["supply_EB"]) )
    if key == "hbmmkt27":
        v = m["SIF_HBM_Mkt_Base"]; return v[-3] if isinstance(v, list) else v
    cur = m
    for p in key.split("."):
        cur = cur[p] if not isinstance(cur, list) else cur[int(p)]
    return cur[idx] if idx is not None else cur
S0ROW = H + 1
for i, (n, desc) in enumerate(SC):
    r = H + 1 + i; m = json.load(open(f"{R}/m_{n}.json"))
    if n == "S0": coef = 0.9404
    ws.cell(r, 1, n).font = BOLD; ws.cell(r, 2, desc)
    for (h, k, kind) in cols:
        j = CIX[k]; key, idx = k
        if kind == "in":
            v = get(m, key, idx)
            if key == "hbmcoef":
                v = 0.940379 if n == "S0" else (0.953734 if n in ("S1", "S2") else 1.0)
            c = ws.cell(r, j, v); c.font = BLUE
        else:
            L = lambda kk: f"{CL(CIX[kk])}{r}"
            if key == "gapf": f = f"={L(('base', idx))}-{L(('dem', idx))}"
            elif key == "gapc": f = f"={L(('base', 1))}-{L(('bridge.x2', 1))}"
            elif key == "d0": f = f"={L(('base', idx))}/{CL(CIX[('base', idx)])}${S0ROW}-1"
            elif key == "r633": f = f"={L(('reconF.implied_rev', 1))}/$E$2-1"
            elif key == "r626": f = f"={L(('reconF.implied_rev', 1))}/$E$3-1"
            elif key == "need": f = f"=$E$2/({L(('reconF.implied_rev', 1))}/$E$5)"
            elif key == "main26": f = f"=$E$4/({L(('reconF.implied_rev', 0))}*6/12)-1"
            ws.cell(r, j, f)
for j in range(3, 3 + len(cols)): ws.column_dimensions[CL(j)].width = 11
ws.column_dimensions["B"].width = 48
# ---------- 個股 ----------
th = '=IF({u}>0.5,"強力買進",IF({u}>0.25,"買進",IF({u}>-0.25,"中立",IF({u}>-0.5,"放空","強力放空"))))'
def stock_sheet(title, before, after, note):
    a = json.load(open(before)); b = json.load(open(after))
    ws = wb.create_sheet(title)
    ws.cell(1, 1, note).font = BOLD
    hd = ["代碼", "公司", "改前目標價", "改前上檔", "改前建議", "改後目標價", "改後上檔（模型）", "改後建議（模型）",
          "現價（推得）", "目標價變動", "改後上檔（公式）", "改後建議（公式）", "建議改變", "離最近門檻（百分點）", "最近門檻"]
    for j, h in enumerate(hd, 1):
        c = ws.cell(3, j, h); c.font = BOLD; c.fill = HDR
    r = 4
    for k in a:
        x, y = a[k], b.get(k, {})
        if "tp" not in x or "tp" not in y: continue
        ws.cell(r, 1, k); ws.cell(r, 2, x["name"])
        for j, v in ((3, x["tp"]), (4, x["up"]), (5, x["rec"]), (6, y["tp"]), (7, y["up"]), (8, y["rec"])):
            ws.cell(r, j, v).font = BLUE
        ws.cell(r, 9, f"=C{r}/(1+D{r})"); ws.cell(r, 10, f"=F{r}/C{r}-1"); ws.cell(r, 11, f"=F{r}/I{r}-1")
        ws.cell(r, 12, th.format(u=f"K{r}")); ws.cell(r, 13, f'=IF(E{r}<>H{r},"是","")')
        ws.cell(r, 14, f"=MIN(ABS(K{r}-0.5),ABS(K{r}-0.25),ABS(K{r}+0.25),ABS(K{r}+0.5))*100")
        ws.cell(r, 15, f'=IF(N{r}=ABS(K{r}-0.5)*100,"+50%",IF(N{r}=ABS(K{r}-0.25)*100,"+25%",IF(N{r}=ABS(K{r}+0.25)*100,"−25%","−50%")))')
        for j in (4, 7, 10, 11): ws.cell(r, j).number_format = "+0.00%;-0.00%"
        ws.cell(r, 14).number_format = "0.00"
        r += 1
    n1 = r - 1
    ws.cell(r + 1, 1, "統計").font = BOLD
    ws.cell(r + 2, 1, "家數"); ws.cell(r + 2, 3, f"=COUNT(J4:J{n1})")
    ws.cell(r + 3, 1, "變動最小"); ws.cell(r + 3, 3, f"=MIN(J4:J{n1})")
    ws.cell(r + 4, 1, "變動最大"); ws.cell(r + 4, 3, f"=MAX(J4:J{n1})")
    ws.cell(r + 5, 1, "中位數"); ws.cell(r + 5, 3, f"=MEDIAN(J4:J{n1})")
    ws.cell(r + 6, 1, "平均"); ws.cell(r + 6, 3, f"=AVERAGE(J4:J{n1})")
    ws.cell(r + 7, 1, "建議改變家數"); ws.cell(r + 7, 3, f'=COUNTIF(M4:M{n1},"是")')
    ws.cell(r + 8, 1, "公式建議與模型建議不一致的家數（應為 0）"); ws.cell(r + 8, 3, f"=SUMPRODUCT((L4:L{n1}<>H4:H{n1})*1)")
    ws.cell(r + 9, 1, "變動 ≤ −5% 家數"); ws.cell(r + 9, 3, f'=COUNTIF(J4:J{n1},"<=-0.05")')
    ws.cell(r + 10, 1, "變動 ≥ +5% 家數"); ws.cell(r + 10, 3, f'=COUNTIF(J4:J{n1},">=0.05")')
    for q in range(3, 7): ws.cell(r + q, 3).number_format = "+0.00%;-0.00%"
    ws.column_dimensions["B"].width = 14
    return ws
STK = json.loads(sys.argv[2]) if len(sys.argv) > 2 else {}
for title, (before, after, note) in STK.items():
    stock_sheet(title, before, after, note)

# ---------- 評等變動 ----------
RAT = json.loads(sys.argv[3]) if len(sys.argv) > 3 else []
ws = wb.create_sheet("評等變動")
ws.cell(1, 1, "投資建議改變的每一家（穩健度：超過門檻 < 1 個百分點＝不穩健；1–3＝中等；> 3＝穩健。另列改前離門檻距離：改前就貼著門檻者，翻轉主要反映改前位置）").font = BOLD
hd = ["情境", "代碼", "公司", "改前建議", "改後建議", "改前上檔", "改後上檔", "越過的門檻", "改後超過門檻（百分點）", "改前離該門檻（百分點）", "目標價變動", "穩健度"]
for j, h in enumerate(hd, 1):
    c = ws.cell(3, j, h); c.font = BOLD; c.fill = HDR
r = 4
TH = [0.5, 0.25, -0.25, -0.5]
for scn, before, after in RAT:
    A = json.load(open(before)); B = json.load(open(after))
    for k in A:
        x, y = A[k], B.get(k, {})
        if "tp" not in x or "tp" not in y or x["rec"] == y["rec"]: continue
        px = x["tp"] / (1 + x["up"]); u1 = y["tp"] / px - 1
        t = [t for t in TH if (x["up"] - t) * (u1 - t) < 0][0]
        for j, v in enumerate([scn, k, x["name"], x["rec"], y["rec"], x["up"], u1, t], 1):
            c = ws.cell(r, j, v); c.font = BLUE if j >= 6 else Font()
        ws.cell(r, 9, f"=ABS(G{r}-H{r})*100"); ws.cell(r, 10, f"=ABS(F{r}-H{r})*100")
        ws.cell(r, 11, f"={y['tp']}/{x['tp']}-1")
        ws.cell(r, 12, f'=IF(I{r}<1,"不穩健",IF(I{r}<3,"中等","穩健"))')
        for j in (6, 7, 8, 11): ws.cell(r, j).number_format = "+0.00%;-0.00%"
        for j in (9, 10): ws.cell(r, j).number_format = "0.00"
        r += 1
# ---------- 方法選項 ----------
MO = json.load(open(sys.argv[4])) if len(sys.argv) > 4 else []
ws = wb.create_sheet("方法選項")
hd = ["序位", "選項", "改不改基準", "評等改變家數（實跑）", "落地要哪些 PR", "證據支持強度", "實施成本", "回復方案"]
for j, h in enumerate(hd, 1):
    c = ws.cell(1, j, h); c.font = BOLD; c.fill = HDR; ws.column_dimensions[CL(j)].width = 40
for i, row in enumerate(MO, 2):
    for j, v in enumerate(row, 1):
        c = ws.cell(i, j, v); c.alignment = Alignment(wrap_text=True, vertical="top")

wb.save(OUT)
print("saved", OUT)
