# xAI 2026 年起年新增 GW（由資本支出反推）對實績校準：試算表產生器（只量測，不讀寫任何模型檔）
# 用法：python3 build.py <輸出 xlsx>；之後以 LibreOffice 重算。數字來源見「輸入」頁的來源欄。
import sys
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter as CL
BLUE = Font(color="0000FF"); BOLD = Font(bold=True); KEY = PatternFill("solid", fgColor="FFF2CC")
wb = openpyxl.Workbook()
# ---------- 輸入 ----------
IN = wb.active; IN.title = "輸入"
IN.append(["鍵", "項目", "單位", "低", "基準", "高", "標記", "來源"]); [setattr(c, "font", BOLD) for c in IN[1]]
ROW = {}
def inp(k, lab, unit, lo, base, hi, tag, src):
    IN.append([k, lab, unit, lo, base, hi, tag, src]); r = IN.max_row; ROW[k] = r
    for c in (4, 5, 6):
        if IN.cell(r, c).value is not None: IN.cell(r, c).font = BLUE
    wb.defined_names[k] = openpyxl.workbook.defined_name.DefinedName(k, attr_text=f"輸入!$E${r}")
YEARS = [2024, 2025, 2026, 2027, 2028, 2029, 2030]
AI = {2024: 5.633, 2025: 12.727, 2026: 74.058119, 2027: 136.252477, 2028: 144.12072, 2029: 111.172527, 2030: 90.160889}
PIT = {2026: 33.414, 2027: 38.372, 2028: 41.333, 2029: 41.333, 2030: 41.333}
for y in YEARS:
    inp(f"ai_{y}", f"xAI AI 資本支出 {y}（現金）", "US$B", None, AI[y], None,
        "Verified" if y <= 2025 else "Derived",
        "SpaceX 424B4 分部資本支出表 AI" if y <= 2025 else "Wattonomics v0.3.9 In gcapai_xai（SpaceX 個股模型 v0.2.9 CF「資本支出：SI（現金）」；2026 含 9M 實績 39.995）")
for y in YEARS[2:]:
    inp(f"pit_{y}", f"每 IT GW 的 IT 設備 {y}", "US$B/GW", None, PIT[y], None, "Derived", "Wattonomics v0.3.9 CapexGW 頁 xai_pit")
inp("pfac_model", "每 IT GW 廠房：模型現行（全市場預設）", "US$B/GW", None, 12.67, None, "Assumed", "Wattonomics v0.3.9 CapexGW 頁 xai_pfac")
inp("pfac_act", "每 IT GW 廠房：xAI 帳面推得", "US$B/GW", 3.7, 5.5, 9.5, "Derived",
    "基準＝2025 年資料中心基礎設施增加 2.736 ÷ 新增 0.5 GW；低＝2026 前九個月 3.373 ÷ 0.9；高＝2025 年底（基礎設施 2.960＋在建工程 4.604）÷ 0.8（在建工程全算廠房，上限）")
inp("w1", "廠房提前入帳權重：前一年", "比", None, 0.5, None, "Assumed", "Wattonomics lag_bldg → w0／w1／w2＝0／0.5／0.5")
inp("w2", "廠房提前入帳權重：前兩年", "比", None, 0.5, None, "Assumed", "同上")
inp("lag_on", "付款到上線延遲", "年", 0.25, 0.375, 0.5, "Assumed", "Wattonomics In lag_on 3／4.5／6 個月")
inp("inv_on", "未上線庫存 ÷ 年出貨", "比", 0.05, 0.10, 0.20, "Assumed", "Wattonomics In inv_on")
inp("accr", "應計 ÷ 現金（以其他融資取得的設備）", "倍", 1.0, 1.10, 1.10, "Derived", "2026 前九個月以其他融資取得 3.921 ÷ SI 現金資本支出 39.995＝9.8%（10-Q；未分分部，歸給 SI 為上限）；未付款變動不計")
inp("gw_2025", "公司揭露算力 2025-12-31", "GW", None, 0.8, None, "Interested-party", "SpaceX 424B4 Nameplate compute draw（公司自訂指標、法定文件揭露）")
inp("add_2025", "2025 年新增", "GW", None, 0.5, None, "Interested-party", "0.8 − 0.3（424B4）")
for col, w in zip("ABCDEFGH", (14, 44, 10, 8, 12, 8, 16, 90)): IN.column_dimensions[col].width = w
# ---------- 逐季 ----------
Q = wb.create_sheet("逐季實績")
hdr = ["項目", "單位", "2025-12-31", "2026-03-31", "2026-06-30", "2026-09-30", "標記", "來源"]
Q.append(hdr); [setattr(c, "font", BOLD) for c in Q[1]]
rowsq = [
    ("公司揭露算力（期末）", "GW", 0.8, 1.0, 1.4, 1.7, "Interested-party", "424B4（2025 年底、2026Q1）；10-Q 2026Q2、2026Q3 Key Business Metrics"),
    ("伺服器與網通設備（帳面原值，全公司）", "US$B", 22.694, None, 34.771, 43.206, "Verified", "10-Q 2026Q2、2026Q3 PP&E 附註（2025 年底為比較欄）"),
    ("資料中心基礎設施（帳面原值）", "US$B", 2.960, None, 3.991, 6.333, "Verified", "同上"),
    ("在建工程（全公司）", "US$B", 4.604, None, 12.554, 23.745, "Verified", "同上；10-Q：主要是尚未啟用的伺服器、網通與資料中心設施"),
    ("SI 分部資本支出（現金；年初累計）", "US$B", None, 7.723, 23.551, 39.995, "Verified", "424B4（Q1）；10-Q 分部附註"),
    ("以其他融資取得的設備（非現金；年初累計；全公司）", "US$B", None, 2.684, 3.921, 3.921, "Verified", "S-1/A No.2 F 頁（Q1；經摘要工具讀取）；10-Q 現金流量表附表"),
    ("已入帳未付款的設備（期末；全公司）", "US$B", 7.088, 10.649, 5.513, 11.827, "Verified", "S-1/A No.2 F 頁（2025 年底、Q1；經摘要工具讀取）；10-Q 現金流量表附表"),
]
for x in rowsq: Q.append(list(x))
r0 = Q.max_row + 2
Q.cell(r0, 1, "推算（年初至該日；以 2025-12-31 為起點）").font = BOLD
calc = [
    ("新增算力", "GW", lambda c: f"={c}2-$C$2"),
    ("啟用金額＝伺服器與網通＋基礎設施的增加", "US$B", lambda c: f"=IF({c}3=\"\",\"\",{c}3-$C$3+{c}4-$C$4)"),
    ("　每 GW 啟用金額（下限：不含仍在在建工程的設備）", "US$B/GW", lambda c: f"=IF({c}3=\"\",\"\",({c}3-$C$3+{c}4-$C$4)/({c}2-$C$2))"),
    ("　其中伺服器與網通", "US$B/GW", lambda c: f"=IF({c}3=\"\",\"\",({c}3-$C$3)/({c}2-$C$2))"),
    ("　其中資料中心基礎設施", "US$B/GW", lambda c: f"=IF({c}4=\"\",\"\",({c}4-$C$4)/({c}2-$C$2))"),
    ("在建工程增加", "US$B", lambda c: f"=IF({c}5=\"\",\"\",{c}5-$C$5)"),
    ("　每 GW（上限：啟用＋在建工程增加全算當期）", "US$B/GW", lambda c: f"=IF({c}3=\"\",\"\",({c}3-$C$3+{c}4-$C$4+{c}5-$C$5)/({c}2-$C$2))"),
]
for i, (lab, unit, f) in enumerate(calc):
    r = r0 + 1 + i; Q.cell(r, 1, lab); Q.cell(r, 2, unit)
    for c in "DEF": Q[f"{c}{r}"] = f(c); Q[f"{c}{r}"].number_format = "0.00"
r = r0 + len(calc) + 2
Q.cell(r, 1, "2026 年底算力的三種讀法").font = BOLD
Q.cell(r + 1, 1, "照前九個月速度外推（0.9 ÷ 9 × 12）"); Q.cell(r + 1, 2, "GW"); Q.cell(r + 1, 6, "=C2+(F2-C2)/9*12")
Q.cell(r + 2, 1, "第四季＝第三季啟用金額占期初在建工程的比例 × 9 月底在建工程 ÷ 每 GW 啟用金額（九個月平均）"); Q.cell(r + 2, 2, "GW")
Q.cell(r + 2, 6, "=F2+((F3-E3+F4-E4)/E5)*F5/((F3-C3+F4-C4)/(F2-C2))")
Q.cell(r + 3, 1, "同上，每 GW 改用 2025 年啟用金額（(22.694−6.892＋2.960−0.224) ÷ 0.5）"); Q.cell(r + 3, 2, "GW")
Q.cell(r + 3, 6, "=F2+((F3-E3+F4-E4)/E5)*F5/((22.694-6.892+2.96-0.224)/0.5)")
Q.cell(r + 4, 1, "公司目標（2026 年底 >2 GW）；Wattonomics 站點讀法 mw_xai 2026＝2,000 MW"); Q.cell(r + 4, 2, "GW"); Q.cell(r + 4, 6, 2.0)
for rr in range(r + 1, r + 5): Q.cell(rr, 6).number_format = "0.00"; Q.cell(rr, 6).fill = KEY
QEV = r
for col, w in zip("ABCDEFGH", (62, 10, 12, 12, 12, 12, 16, 70)): Q.column_dimensions[col].width = w
# ---------- 情境 ----------
S = wb.create_sheet("情境")
S["A1"] = "年新增 GW 與 Wattonomics CapexGW 頁同式：資金＝IT 占比 × 當年＋廠房占比 ×（w1 × 前一年＋w2 × 前兩年）；付款年 GW＝資金 ÷（IT＋廠房）；上線年＝(1−lag)×當年＋lag×前一年 − inv×（當年−前一年）"
S["A1"].font = BOLD
SC = [("現行", "pfac_model", 0, "1"), ("A 只換廠房單價", "pfac_act", 0, "1"), ("C 只把 2024–25 實績放進付款鏈", "pfac_model", 1, "1"),
      ("D 只改應計口徑", "pfac_model", 0, "accr"), ("A＋C", "pfac_act", 1, "1"), ("A＋C＋D", "pfac_act", 1, "accr")]
YS = [2026, 2027, 2028, 2029, 2030]
r = 3; SUM = {}
for name, pf, hist, sc in SC:
    S.cell(r, 1, name).font = BOLD; S.cell(r, 2, "廠房單價"); S.cell(r, 3, f"={pf}"); S.cell(r, 4, "歷史實績入鏈"); S.cell(r, 5, hist); S.cell(r, 6, "口徑倍數"); S.cell(r, 7, f"={sc}")
    PF, H, K = f"$C${r}", f"$E${r}", f"$G${r}"
    S.cell(r + 1, 1, "年"); [S.cell(r + 1, 2 + i, y) for i, y in enumerate([2025] + YS)]
    labs = ["當年資本支出", "前一年（空白年代入當年）", "前兩年（空白年代入前一年）", "資金", "付款年 GW", "上線年 GW（年新增）", "年底 GW"]
    for i, l in enumerate(labs): S.cell(r + 2 + i, 1, "　" + l)
    R = {k: r + 2 + i for i, k in enumerate(("cur", "a1", "a2", "fund", "x", "on", "ye"))}
    S.cell(R["x"], 2, "=add_2025"); S.cell(R["on"], 2, "=add_2025"); S.cell(R["ye"], 2, "=gw_2025")
    for i, y in enumerate(YS):
        c = CL(3 + i); p = CL(2 + i)
        S[f"{c}{R['cur']}"] = f"=ai_{y}*{K}"
        if y == 2026:
            S[f"{c}{R['a1']}"] = f"=IF({H}=1,ai_2025,{c}{R['cur']})"; S[f"{c}{R['a2']}"] = f"=IF({H}=1,ai_2024,{c}{R['a1']})"
        elif y == 2027:
            S[f"{c}{R['a1']}"] = f"={p}{R['cur']}"; S[f"{c}{R['a2']}"] = f"=IF({H}=1,ai_2025,{c}{R['a1']})"
        else:
            S[f"{c}{R['a1']}"] = f"={p}{R['cur']}"; S[f"{c}{R['a2']}"] = f"={CL(1 + i)}{R['cur']}"
        S[f"{c}{R['fund']}"] = f"=pit_{y}/(pit_{y}+{PF})*{c}{R['cur']}+{PF}/(pit_{y}+{PF})*(w1*{c}{R['a1']}+w2*{c}{R['a2']})"
        S[f"{c}{R['x']}"] = f"={c}{R['fund']}/(pit_{y}+{PF})"
        sh = f"(1-lag_on)*{c}{R['x']}+lag_on*{p}{R['x']}-inv_on*({c}{R['x']}-{p}{R['x']})"
        S[f"{c}{R['on']}"] = (f"=IF({H}=1,{sh},{c}{R['x']})" if y == 2026 else f"={sh}")
        S[f"{c}{R['ye']}"] = f"={p}{R['ye']}+{c}{R['on']}"
        for k in R: S[f"{c}{R[k]}"].number_format = "0.000"
        S[f"{c}{R['on']}"].fill = KEY
    SUM[name] = R; r += 11
# 彙總
T = wb.create_sheet("彙總", 0)
T["A1"] = "xAI 年新增 GW：各情境對現行（公式連到「情境」頁）；全市場年新增取 Wattonomics v0.3.9 WIF_MktGW_Add"; T["A1"].font = BOLD
T.append([]); T.append(["情境", "年新增 2026", "2027", "2028", "2029", "2026 年底", "2029 年底", "對全市場年新增 2026", "2027", "2028", "2029"]); [setattr(c, "font", BOLD) for c in T[3]]
MK = {2026: 19.294, 2027: 25.678, 2028: 29.472, 2029: 29.569}
base = SUM["現行"]
for name, R in SUM.items():
    row = [name] + [f"=情境!{CL(3 + i)}{R['on']}" for i in range(4)] + [f"=情境!C{R['ye']}", f"=情境!F{R['ye']}"]
    row += [f"=(情境!{CL(3 + i)}{R['on']}-情境!{CL(3 + i)}{base['on']})/{MK[2026 + i]}" for i in range(4)]
    T.append(row); rr = T.max_row
    for c in range(2, 8): T.cell(rr, c).number_format = "0.00"
    for c in range(8, 12): T.cell(rr, c).number_format = "+0.0%;-0.0%;0.0%"
rr = T.max_row + 2
T.cell(rr, 1, "實體證據：2026 年底算力").font = BOLD
for i, lab in enumerate(["照前九個月速度外推", "在建工程轉啟用（每 GW 用九個月平均）", "在建工程轉啟用（每 GW 用 2025 年）", "公司目標／站點讀法"]):
    T.cell(rr + 1 + i, 1, lab); T.cell(rr + 1 + i, 6, f"=逐季實績!F{QEV + 1 + i}"); T.cell(rr + 1 + i, 6).number_format = "0.00"
T.column_dimensions["A"].width = 44
for c in "BCDEFGHIJK": T.column_dimensions[c].width = 13
wb.save(sys.argv[1])
