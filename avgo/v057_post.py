# Broadcom v0.5.7（2026-11-14）後處理：C0（實績）——公司實績 vs 層級隱含值（三層基準協調規則 v1.4.2 C0 第 2 點；
# 工作單 docs/workorders/20261114_C0實績對帳補列_七家.md；Andy 2026-11-14 17:40）。
# 只新增與改標：Coord 頁末 J 節（新）；G 表狀態欄改標「讀取層級（同源，不判定）」／「推估錨點（參考）」；對帳中項目數改由 J 節判定；README 版號與說明。
# 既有數值格不變；目標價、評等不變。build_avgo.py 在 v056_post.apply 之後呼叫 apply(wb, here, SIL)。環境變數 COORD_RUN／AVGO_SENS 設定時不執行。
import os
from openpyxl.styles import Font, PatternFill, Alignment
from v056_post import lab, F, T, dn, BOLD, BLUE, HDR

VER, BUILT = "0.5.7", "2026-11-14"
GREEN = Font(color="008000")
EDG = "https://www.sec.gov/Archives/edgar/data/1730168/"
SRC_Q = ("Q1：" + EDG + "000173016826000011/avgo-02012026x8kxex99.htm；Q2：" + EDG + "000173016826000051/avgo-05032026x8kxex99.htm；"
         "Q3：https://investors.broadcom.com/news-releases/news-release-details/broadcom-inc-announces-third-quarter-fiscal-year-2026-financial（均為 8-K 新聞稿；本模型 Q 頁第 7 列）")
BILL = "Drive 02-AI供應鏈/對帳/20261114_對帳單_Broadcom_AI網通FY2026_個股對市場層歸屬.md"


def apply(wb, here, SIL):
    if os.environ.get("COORD_RUN") == "1" or os.environ.get("AVGO_SENS") == "1":
        return "v057_post：略過（alt_run／sens_avgo 建置）"
    C, RD, Q = wb["Coord"], wb["README"], wb["Q"]
    CO = wb._v056["CO"]; RM = wb._v056["RM"]
    G0, G1 = CO["G0"], CO["G1"]
    r_xpu = lab(C, "FY2026 XPU 營收：公司實績推估", G0)
    r_net = lab(C, "FY2026 AI 網通營收：公司實績推估", G0)
    assert r_net == G1 and r_xpu == G1 - 1, (r_xpu, r_net, G1)
    rq = lab(Q, "AI 半導體營收", 1, exact=True)
    assert [Q.cell(3, c).value[:5] for c in (3, 4, 5)] == ["FY26 ", "FY26 ", "FY26 "]

    # ---- G 表改標：同源值＝讀取層級（不判定）；FY2026 XPU／網通＝推估錨點（參考，非公司揭露） ----
    g_title = lab(C, "G. C0 由下往上對帳")
    C.cell(g_title, 1, "G. 讀取層級（同源，不判定）與推估錨點參考（三層基準協調規則 v1.4.2 C0 第 2 點）：個股自層級讀入的值與層級值必然相同，只列讀取層級、不判定、不計對帳中；"
                       "FY2026 XPU 與 AI 網通為模型推估錨點（網通拆分為 Derived，非公司揭露），只列參考。C0 判定看本頁 J 節「C0（實績）」")
    for r in range(G0, G1 + 1):
        if r < r_xpu:
            if str(C.cell(r, 5).value).startswith("=IF(ABS(D"):
                C.cell(r, 5).value = "讀取層級（同源，不判定）"
            else:   # 份額恆等式：保留殘差明列
                C.cell(r, 5).value = f'="讀取層級（同源，不判定）；未建模殘差 "&TEXT(D{r},"0.0%")&" 明列"'
        elif r == r_xpu:
            C.cell(r, 5).value = f'=IF(ABS(D{r})>0.15,"推估錨點（參考，不判定）：差距超過 ±15%","推估錨點（參考，不判定）：差距在 ±15% 內")'
        else:
            C.cell(r, 5).value = (f'=IF(ABS(D{r})>0.15,"推估錨點（參考，不判定）：差距超過 ±15%；原判對帳中，已另開對帳單，判定改由 J 節實績列決定",'
                                  f'"推估錨點（參考，不判定）：差距在 ±15% 內")')
    C.cell(r_net, 9).value = (str(C.cell(r_net, 9).value) + "。本列保留為參考：網通拆分是 Derived（非公司揭露）；已另開對帳單（檔名見對照來源欄）；"
                              "判定改由 J 節以 AI 半導體營收實績（公司 8-K 揭露）對層級隱含值決定")
    C.cell(r_net, 7).value = str(C.cell(r_net, 7).value) + f"；對帳單：{BILL}"

    # ---- J 節：C0（實績） ----
    r = C.max_row + 2
    C.cell(r, 1, "J. C0（實績）：公司實績 vs 層級隱含值（規則 v1.4.2 C0 第 2 點；工作單 20261114_C0實績對帳補列_七家）").font = BOLD; j_title = r; r += 1
    C.cell(r, 1, "方法：個股側＝Broadcom 8-K 已公布的 AI 半導體營收（公司揭露，Verified；公司指引不算實績，只列參考）。層級側＝Siliconomics 現行版隱含的同口徑值＝XPU 由下而上"
                 "（顆數 × 份額 × 核心售價＋代購 HBM，未校準；G 表 FY2026 XPU 列的層級值）＋市場層 AI 網通 Broadcom 歸屬值（G 表 FY2026 AI 網通列的層級值）。"
                 "時差：層級值是日曆年 CY2026；Broadcom FY2026 Q1–Q3＝2025-11～2026-07（9 個月）。主讀數＝Q1–Q3 累計實績 vs 層級 CY2026 × 已公布月數 ÷ 12（線性時差近似，Derived）；"
                 "第二讀數＝最近一季（Q3）× 4 vs 層級 CY2026。成長年中主讀數偏低、第二讀數偏高，兩者夾住真值。判定：兩讀數差距都超過 ±15% 且同號 → 對帳中；只有一個超過 → 時差邊界（參考）；都在內 → 一致。"
                 "不用模型預測、同源值或賣方估計補位。")
    C.cell(r, 1).alignment = Alignment(wrap_text=True, vertical="top"); r += 2
    for j, h in enumerate(["輸入", "值", "單位", "期間／口徑", "出處", "等級"], 1):
        c = C.cell(r, j, h); c.font = BOLD; c.fill = HDR
    r += 1
    IN = {}

    def inp(key, label, val, unit, period, src, ev, font=None, fmt="#,##0.000"):
        nonlocal r
        C.cell(r, 1, label); C.cell(r, 2, val).number_format = fmt
        if font: C.cell(r, 2).font = font
        C.cell(r, 3, unit); C.cell(r, 4, period); C.cell(r, 5, src); C.cell(r, 6, ev); IN[key] = r; r += 1
    inp("q1", "AI 半導體營收：FY2026 Q1（公司 8-K）", f"=Q!C{rq}", "US$B", "2025-11-03～2026-02-01（Q1）", "Q 頁 C7；" + SRC_Q.split("；")[0], "Verified", GREEN)
    inp("q2", "AI 半導體營收：FY2026 Q2（公司 8-K）", f"=Q!D{rq}", "US$B", "2026-02-02～2026-05-03（Q2）", "Q 頁 D7；" + SRC_Q.split("；")[1], "Verified", GREEN)
    inp("q3", "AI 半導體營收：FY2026 Q3（公司 8-K）", f"=Q!E{rq}", "US$B", "2026-05-04～2026-08-02（Q3）", "Q 頁 E7；" + SRC_Q.split("；")[2], "Verified", GREEN)
    inp("mo", "已公布月數（FY2026 Q1–Q3）", 9, "月", "2025-11～2026-07", "會計年度 FY2026 起於 2025-11-03；三季＝9 個月", "Derived", BLUE, fmt="0")
    inp("x", "層級：FY2026 XPU 由下而上（顆數 × 份額 × 核心售價＋代購 HBM，未校準）", f"=C{r_xpu}", "US$B", "CY2026（日曆年）", f"G 表第 {r_xpu} 列層級值（Siliconomics {SIL.VERSION} 顆數、份額、核心售價）", "Assumed（層）", GREEN)
    inp("n", "層級：FY2026 AI 網通 Broadcom 歸屬值（市場層）", f"=C{r_net}", "US$B", "CY2026（日曆年）", f"G 表第 {r_net} 列層級值（markets/network net_26；廠商歸屬比例為 Assumed）", "Assumed（層）", GREEN)
    inp("L", "層級隱含：AI 半導體營收 CY2026（XPU＋網通）", f"=B{IN['x']}+B{IN['n']}", "US$B", "CY2026（日曆年）", "上兩列加總", "Assumed（層）／Derived（加總）")
    inp("g", "參考：公司 Q4 FY2026 指引（不判定）", f"=Q!F{rq}", "US$B", "FY2026 Q4（指引，2026-12 公布實績）", "Q 頁 F7（公司指引；Interested-party）", "Interested-party", GREEN)
    r += 1
    for j, h in enumerate(["項目", "公司實績", "層級隱含", "差距", "判定", "單位與口徑", "出處", "證據等級", "說明"], 1):
        c = C.cell(r, j, h); c.font = BOLD; c.fill = HDR
    r += 1
    J0 = r
    pct = '+0.0%;-0.0%'
    b = IN
    C.cell(r, 1, "主讀數：AI 半導體營收，FY2026 Q1–Q3 累計實績 vs 層級 CY2026 × 已公布月數 ÷ 12")
    C.cell(r, 2, f"=B{b['q1']}+B{b['q2']}+B{b['q3']}").number_format = "#,##0.0"
    C.cell(r, 3, f"=B{b['L']}*B{b['mo']}/12").number_format = "#,##0.0"
    C.cell(r, 4, f"=B{r}/C{r}-1").number_format = pct
    C.cell(r, 5, f'=IF(ABS(D{r})>0.15,"超過 ±15%","在 ±15% 內")')
    C.cell(r, 6, "US$B；公司實績＝三季 AI 半導體營收加總；層級＝CY2026 全年 × 9 ÷ 12")
    C.cell(r, 7, "公司：Broadcom 8-K（Q 頁第 7 列）；層級：Siliconomics " + SIL.VERSION + "（G 表 XPU 與網通列層級值）")
    C.cell(r, 8, "Verified（公司）／Assumed（層）／Derived（線性時差）")
    C.cell(r, 9, "成長年中，線性分攤的層級值偏高（後段季度較大），主讀數偏低"); rm = r; r += 1
    C.cell(r, 1, "第二讀數：AI 半導體營收，最近一季（FY2026 Q3）× 4 vs 層級 CY2026")
    C.cell(r, 2, f"=B{b['q3']}*4").number_format = "#,##0.0"
    C.cell(r, 3, f"=B{b['L']}").number_format = "#,##0.0"
    C.cell(r, 4, f"=B{r}/C{r}-1").number_format = pct
    C.cell(r, 5, f'=IF(ABS(D{r})>0.15,"超過 ±15%","在 ±15% 內")')
    C.cell(r, 6, "US$B；Q3 單季 × 4（年化）；層級＝CY2026 全年")
    C.cell(r, 7, "公司：Broadcom 8-K（Q 頁 E7）；層級同上")
    C.cell(r, 8, "Verified（公司）／Assumed（層）／Derived（年化）")
    C.cell(r, 9, "成長年中，最近一季年化偏高；且 Q3 為 2026-05～08，與 CY2026 全年平均相比位於後段"); rs = r; r += 1
    C.cell(r, 1, "C0（實績）判定：AI 半導體營收 FY2026").font = BOLD
    C.cell(r, 5, f'=IF(AND(ABS(D{rm})>0.15,ABS(D{rs})>0.15,SIGN(D{rm})=SIGN(D{rs})),"對帳中",IF(OR(ABS(D{rm})>0.15,ABS(D{rs})>0.15),"時差邊界（參考）","一致"))').font = BOLD
    C.cell(r, 6, "兩讀數差距都超過 ±15% 且同號＝對帳中；只有一個超過＝時差邊界（參考）；都在內＝一致")
    C.cell(r, 9, f'=IF(E{r}="對帳中","兩讀數同向偏離：開對帳單（Drive 02-AI供應鏈/對帳/），不改上游值",IF(E{r}="時差邊界（參考）","兩讀數一個在 ±15% 內、一個超過：層級值落在兩讀數之間，沒有同向偏離；不開改值對帳單，Q4 財報（2026-12）公布後以全年實績重判","兩讀數都在 ±15% 內：層級隱含與公司實績相容"))')
    rv = r; r += 1
    C.cell(r, 1, "參考：FY2026 全年（Q1–Q3 實績＋Q4 公司指引）vs 層級 CY2026（含指引，不判定）")
    C.cell(r, 2, f"=Q!G{rq}").number_format = "#,##0.0"; C.cell(r, 2).font = GREEN
    C.cell(r, 3, f"=B{b['L']}").number_format = "#,##0.0"
    C.cell(r, 4, f"=B{r}/C{r}-1").number_format = pct
    C.cell(r, 5, "參考（含指引，不判定）")
    C.cell(r, 6, "US$B；Q 頁 FY2026 錨點（Q1–Q3 實績＋Q4 指引）；會計年度與日曆年錯開約兩個月")
    C.cell(r, 7, "Q 頁 G7（Q4 為公司指引）"); C.cell(r, 8, "Verified／Interested-party（Q4 指引）")
    C.cell(r, 9, "指引不算實績；此列只讓讀者看到全年口徑下的差距，不進判定、不計對帳中"); rr = r; r += 1
    J1 = r - 1
    r += 1
    C.cell(r, 1, "對帳中的 C0 項目數（J 節 C0（實績）判定；G 表讀取層級與推估錨點列不計）"); C.cell(r, 2, f'=IF(E{rv}="對帳中",1,0)'); cnt2 = r; r += 1
    C.cell(r, 1, "說明：J 節為本公司唯一的 C0 判定；G 表同源列（讀取層級）與推估錨點列不判定；AI 網通 FY2026 原列的差距另見對帳單，原因為網通拆分屬 Derived。層與層之間的差距只列「層級之間（參考）」。")
    r += 1
    for rr_ in (j_title + 1,):
        C.cell(rr_, 1).alignment = Alignment(wrap_text=True, vertical="top")

    # ---- 既有對帳中計數格與說明改指向 J 節 ----
    C.cell(CO["cnt"], 1).value = "對帳中的 C0 項目數（見 J 節 C0（實績）判定；G 表讀取層級與推估錨點列不計）"
    C.cell(CO["cnt"], 2).value = f"=B{cnt2}"
    C.cell(CO["note"], 1).value = ("說明：C0 判定改由 J 節「C0（實績）」（公司 AI 半導體營收實績 vs 層級隱含值）決定；G 表只列讀取層級（同源，不判定）與推估錨點（參考）。"
                                   "AI 網通 FY2026 的推估錨點對市場層歸屬差距已另開對帳單；產業部署 GW 為層與層之間的差距（參考列，Siliconomics 對帳單 r3 處理），本模型並列兩讀數目標價（Coord 頁 A–E 節），不裁決，也不計入對帳中項目數。")
    # ---- README：coord 白話列、版本說明列、A1 ----
    RD.cell(RM["coord"], 1).value = F(
        T("三層協調：本模型的產業顆數、份額與對帳中旗標讀 Siliconomics "), f"Coord!C{CO['F1']}", T("（見 Coord 頁 F 節）。C0 改以公司實績對層級隱含值判定（Coord 頁 J 節）：Broadcom 8-K 的 AI 半導體營收 FY2026 Q1–Q3 累計 "),
        f"TEXT(Coord!B{rm},\"0.0\")", T(" 十億美元，對層級隱含（XPU 由下而上加網通歸屬）依已公布月數折算 "), f"TEXT(Coord!C{rm},\"0.0\")", T("，差 "), f"TEXT(Coord!D{rm},\"+0.0%;-0.0%\")",
        T("；最近一季年化 "), f"TEXT(Coord!B{rs},\"0.0\")", T("，差 "), f"TEXT(Coord!D{rs},\"+0.0%;-0.0%\")", T("；判定「"), f"Coord!E{rv}", T("」，對帳中項目 "), f"TEXT(Coord!B{CO['cnt']},\"0\")",
        T(" 項。讀自層級的同源值只列讀取層級、不判定；產業部署 GW 兩層差距另列參考、不計入；對帳中年份 CY2027–CY2029 依 C4 並列兩讀數目標價。"))
    RD.cell(RM["ver"], 1).value = (f"v{VER}（{BUILT}）：C0 實績對帳補列（三層基準協調規則 v1.4.2 C0 第 2 點；工作單 20261114_C0實績對帳補列_七家）——Coord 頁 J 節新增「C0（實績）」：Broadcom 8-K 的 AI 半導體營收（FY2026 Q1–Q3）對層級隱含值"
                                   "（XPU 由下而上＋網通歸屬）；G 表既有同源列改標讀取層級（不判定），FY2026 XPU／網通列改標推估錨點（參考）；對帳中項目數改由 J 節判定。"
                                   "其餘（版面 v1.3、目標價、評等、另一側讀數）不變；既有數值格不變（新增列與文字改標以外逐格比對 0 差異）")
    a1 = RD["A1"].value
    head = a1.split("；", 1)[0]
    assert "v0.5.6" in head, head[:120]
    RD["A1"].value = head.replace("v0.5.6", f"v{VER}", 1) + a1[len(head):] + (
        f"；v{VER}（{BUILT}）：C0 實績對帳補列（Coord 頁 J 節：AI 半導體營收實績 vs 層級隱含值；G 表改標讀取層級／推估錨點），既有數值不變")
    wb._v057 = dict(J0=J0, J1=J1, rm=rm, rs=rs, rv=rv, title=j_title, cnt2=cnt2)
    return f"v057_post：Coord J 節第 {j_title} 列起（主讀數第 {rm} 列、第二讀數第 {rs} 列、判定第 {rv} 列）"
