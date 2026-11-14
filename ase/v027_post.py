# 日月光投控 v0.2.7（2026-11-14）後處理：C0 實績對帳補列（三層基準協調規則 v1.4.2 C0 第 2 點；工作單 docs/workorders/20261114_C0實績對帳補列_七家.md）。
# 在 v026_post.apply 之後呼叫 apply(wb, here, SIL)。只新增列與改標：
#   Coord 頁 B 節：個股讀入層級值的同源列改標「讀取層級（同源，不判定）」，不計入對帳中項目數；標題、判定摘要、C0 結論改文字
#   Coord 頁末 E 節（新）：「C0（實績）：公司實績 vs 層級隱含值」表＋ATM 封裝／測試分項輸入（公司分部表）
#   README 頁末：v0.2.7 一列；A1、A2 版號文字
# 另一側讀數建置（COORD_RUN＝1）時不執行。
import os
from openpyxl.styles import Font, PatternFill

import v026_post as V26
from v026_post import lab, T, F

VER, BUILT = "0.2.7", "2026-11-14"
BOLD = Font(bold=True); BLUE = Font(color="0000FF"); HDR = PatternFill("solid", fgColor="DDEBF7")
SEC3Q = "https://www.sec.gov/Archives/edgar/data/0001122411/000095010326016323/dp255782_6k.htm"
SEC2Q = "https://www.sec.gov/Archives/edgar/data/0001122411/000095010326011353/dp250875_6k.htm"


def coord_actuals(wb, SIL):
    C = wb["Coord"]
    # ---- B 節改標：同源列不判定 ----
    rb = lab(C, "B. C0 由下往上對帳")
    C.cell(rb, 1, "B. 讀取層級（同源，不判定）：日月光讀入的 Siliconomics 值 vs 層級值（個股自層級讀入，差距恆為 0，無檢驗力；不計入對帳中項目數；規則 v1.4.2 C0 第 2 點）。"
                  "公司實績對層級隱含值的對帳見本頁 E 節「C0（實績）」")
    C.cell(rb + 1, 5, "判定")
    r_first = lab(C, "AI 加速器顆數（情境 B）CY2026")
    for r in range(r_first, r_first + 4):
        C.cell(r, 5, "讀取層級（同源，不判定）")
    r_nc = lab(C, "日月光 LEAP 營收（FY2026")
    C.cell(r_nc, 5, "不判定（無可比實績，見 E 節）")

    # ---- E 節：C0（實績） ----
    r = C.max_row + 3
    C.cell(r, 1, "E. C0（實績）：公司實績 vs 層級隱含值（規則 v1.4.2 C0 第 2 點；Andy 11-14 17:40 固定規格；差距超過 ±15% 且兩讀數同號標「對帳中」）").font = BOLD
    R = {"E": r}; r += 1
    C.cell(r, 1, "比較方法：個股側只用公司已公布實績（SEC 6-K 新聞稿與簡報，Verified）；層級側用 Siliconomics v0.19 同口徑隱含值。時差：層級值為日曆年 CY2026，日月光會計年度＝日曆年；"
                 "主讀數＝2026 年 1–3Q 累計實績 vs 層級 CY2026 × 9 ÷ 12（線性時差近似，Derived）；第二讀數＝3Q26 × 4 vs 層級 CY2026。成長年中主讀數偏低、第二讀數偏高，兩者夾住真值。"
                 "判定：兩讀數差距都超過 ±15% 且同號＝對帳中；只有一個超過＝時差邊界（參考）；都在內＝一致；層級側沒有同口徑隱含值＝無可比實績（不以模型預測或同源值補位）。"
                 "公司指引、賣方估計只列參考，不判定。本表「無可比實績」是指層級側沒有同口徑值（ATM、封裝、測試的公司實績都存在）；LEAP 則是公司側也沒有實績。"); R["note"] = r; r += 1
    for j, h in enumerate(["項目", "公司實績：2026 年 1–3Q 累計（主讀數）", "公司實績：3Q26 × 4（第二讀數）", "層級隱含：CY2026 全年", "判定（±15%）", "口徑、出處與證據等級", "主讀數差距（實績 ÷（層級 × 9÷12）− 1）", "第二讀數差距（實績 ÷ 層級 − 1）"], 1):
        c = C.cell(r, j, h); c.font = BOLD; c.fill = HDR
    R["hdr"] = r; r += 1
    R["first"] = r
    r_in = r + 7                       # 輸入區標題列（表 4 列＋判定摘要＋空列之後）
    r_pk, r_ts = r_in + 2, r_in + 3    # 封裝、測試分項輸入列

    def row(label, b, c, d, src, nb="0.000"):
        nonlocal r
        C.cell(r, 1, label)
        for col, v in ((2, b), (3, c)):
            x = C.cell(r, col, v)
            if not isinstance(v, str) or v.startswith("="): x.number_format = nb
        C.cell(r, 4, d)
        C.cell(r, 7, f'=IF(AND(ISNUMBER(B{r}),ISNUMBER(D{r})),B{r}/(D{r}*9/12)-1,"—")').number_format = "+0.0%;-0.0%"
        C.cell(r, 8, f'=IF(AND(ISNUMBER(C{r}),ISNUMBER(D{r})),C{r}/D{r}-1,"—")').number_format = "+0.0%;-0.0%"
        C.cell(r, 5, f'=IF(AND(ISNUMBER(G{r}),ISNUMBER(H{r})),IF(AND(ABS(G{r})>0.15,ABS(H{r})>0.15,SIGN(G{r})=SIGN(H{r})),"對帳中",IF(OR(ABS(G{r})>0.15,ABS(H{r})>0.15),"時差邊界（參考）","一致")),"無可比實績")')
        C.cell(r, 6, src)
        r += 1

    row("LEAP（AI 先進封裝）營收（US$B）", "無實績", "無實績", "無同口徑隱含值",
        "公司只揭露 FY2025 全年 US$1.6B（4Q25 法說）與 FY2026 指引（2Q26 簡報第 3 頁：LEAP 營收「tracking ahead of prior guidance of US$3.5 billion」；3Q26 法說：逾 US$4B），"
        "季度與 2026 年累計金額未揭露（2Q26 6-K 第 3、4 頁；3Q26 6-K 無 LEAP 金額）；本模型 Q 頁 1Q–3Q26 的 LEAP 為 BofA 占比推得（Interested-party／Derived），不是公司實績，不用。"
        "層級側 Siliconomics 沒有 AI 先進封裝的美元市場值：只有委外封測（日月光＋艾克爾合計）補充 CoWoS 月產能（2026 年底 130 千片／月為 Interested-party；委外占台積電產能 10% 為 Assumed），缺每片價格與日月光自己的份額，無法換算美元。"
        "兩側皆缺，寫無可比實績；改用最接近的已揭露分部（下列封測 ATM 與其封裝、測試分項）。橋接參考（LEAP 核心營收 ÷ 委外補充產能）見 B 節，不判定")
    row("封測（ATM）營收（NT$B；封裝、測試、材料）", "=SUM(Q!C5:E5)", "=Q!E5*4", "無同口徑隱含值",
        "公司實績：ATM 分部營收 1Q26 112.434、2Q26 126.148、3Q26 151.485（SEC 6-K 2026-07-30 第 7 頁、2026-10-29 第 5 頁，ATM 分部損益表；Q 頁 C5:E5；Verified；含集團內沖銷前口徑）。"
        "層級側：Siliconomics v0.19 沒有封測服務市場值或日月光份額（層級內與封測相關者只有測試機設備市場與委外補充 CoWoS 產能，單位與口徑皆不同）。無可比實績：缺層級的封測（含主流封裝）營收市場與日月光份額")
    row("封測（ATM）：封裝（Packaging）營收（NT$B）", f"=SUM(B{r_pk}:D{r_pk})", f"=D{r_pk}*4", "無同口徑隱含值",
        "公司實績：ATM 分部封裝營收（輸入區，公司分部損益表；Verified）。封裝含打線、覆晶、晶圓級與 SiP 等，LEAP 只是其中一部分，公司未拆。層級側無封裝市場值；無可比實績（缺層級的先進封裝美元市場與日月光份額）")
    row("封測（ATM）：測試（Testing）營收（NT$B）", f"=SUM(B{r_ts}:D{r_ts})", f"=D{r_ts}*4", "無同口徑隱含值",
        "公司實績：ATM 分部測試營收（輸入區，公司分部損益表；Verified）。層級側只有測試機設備市場（Advantest、泰瑞達的設備銷售，CY2026 SoC 測試機市場約 US$11.5B；Interested-party），是設備商營收，不是 OSAT 的測試服務營收，口徑不同，不可比；無可比實績")
    R["last"] = r - 1
    gw0 = SIL.NAMES["SIF_DeployGW_Base"][0]
    C.cell(r, 1, "參考（不判定）：外部估計 OSAT 先進封裝市場 ÷ 公司 LEAP 指引（US$B）")
    c = C.cell(r, 2, f"=0.5*{gw0}"); c.number_format = "0.0"; c.font = BLUE
    c = C.cell(r, 3, "=In!E35"); c.number_format = "0.0"
    C.cell(r, 4, "—")
    C.cell(r, 5, "參考（不判定）")
    C.cell(r, 6, f"B 欄＝外部估計 OSAT 先進封裝市場約 US$0.5B／GW（艾克爾投資人日，Interested-party）× Siliconomics v0.19 SIF_DeployGW_Base CY2026（{gw0:.2f} GW）；C 欄＝公司 FY2026 LEAP 指引（逾 US$4B，In 頁 leap26，Interested-party）。"
                 "C ÷ B 見 G 欄。這是外部估計對公司指引，不是公司實績對層級隱含值，不是 C0 證據、不判定。若要可判定，需 Siliconomics 補 OSAT 先進封裝美元市場（或委外 CoWoS 每片價格）與日月光份額。")
    C.cell(r, 7, f'=TEXT(C{r}/B{r},"0%")&"（指引 ÷ 外部估計）"')
    C.cell(r, 8, "—")
    R["ref"] = r; r += 1
    C.cell(r, 1, "判定摘要"); R["sum"] = r
    rg = f'E{R["first"]}:E{R["last"]}'
    C.cell(r, 2, F(T("C0（實績）共 "), f'COUNTA({rg})', T(" 項：可判定 "), f'COUNT(G{R["first"]}:G{R["last"]})', T(" 項、對帳中 "), f'COUNTIF({rg},"對帳中")', T(" 項、時差邊界 "), f'COUNTIF({rg},"時差邊界（參考）")',
                   T(" 項、無可比實績 "), f'COUNTIF({rg},"無可比實績")', T(" 項。原因：公司只公布 ATM 分部與其封裝、測試分項，未公布 AI 先進封裝（LEAP）金額；Siliconomics 沒有封測服務或先進封裝的美元市場與日月光份額。"
                   "本模型不以模型預測或同源值補位；待層級補上 OSAT 市場與份額，或公司揭露 LEAP 金額後再判定。")))
    R["sum_rg"] = rg
    r += 2
    assert r == r_in, (r, r_in)
    C.cell(r, 1, "輸入：ATM 封裝、測試分項營收（NT$B；公司 ATM 分部損益表，含集團內沖銷前口徑；Verified；藍字為輸入）").font = BOLD; r += 1
    for j, h in enumerate(["項目", "1Q26", "2Q26", "3Q26", "加總 ÷ 對照", "出處"], 1):
        c = C.cell(r, j, h); c.font = BOLD; c.fill = HDR
    r += 1
    for lab_, vals, chk in (("封裝（Packaging）", (89.673, 100.324, 122.945), "=B{r}+C{r}+D{r}"), ("測試（Testing）", (21.041, 23.665, 26.302), "=B{r}+C{r}+D{r}")):
        C.cell(r, 1, lab_)
        for j, v in enumerate(vals):
            c = C.cell(r, 2 + j, v); c.font = BLUE; c.number_format = "0.000"
        C.cell(r, 5, chk.format(r=r)).number_format = "0.000"
        C.cell(r, 6, f"1Q26、2Q26：公司 2Q26 新聞稿與簡報 ATM 分部損益表第 7 頁（{SEC2Q}）；3Q26：公司 3Q26 新聞稿與簡報第 5 頁（{SEC3Q}）；Verified。百萬元換算十億元（÷1000）")
        r += 1
    assert r - 2 == r_pk, (r, r_pk)
    C.cell(r, 1, "檢查：ATM 營收 − 封裝 − 測試 − 直接材料與其他（3Q26；應 ≈0；公司：直接材料 2.148、其他 0.090）"); R["chk"] = r
    C.cell(r, 2, f"=ROUND(Q!E5-D{r_pk}-D{r_ts}-2.148-0.09,6)").number_format = "0.000"
    r += 1
    R["end"] = r - 1

    # ---- B 節判定摘要與 C0 結論改文字（改指向 E 節計數） ----
    r_sum = lab(C, "判定摘要")                  # B 節那一列（E 節的在更後面，lab 取第一個）
    C.cell(r_sum, 2, F(T("本公司 C0 對帳中項目 "), f'COUNTIF({R["sum_rg"]},"對帳中")', T(" 項：公司實績對層級隱含值的 C0（實績）表見 E 節（可判定 "), f'COUNT(G{R["first"]}:G{R["last"]})',
                         T(" 項；其餘因層級側沒有同口徑隱含值，標無可比實績）。本節加速器顆數 CY2026–29 是個股讀入的層級值，與層級同源，改標「讀取層級（同源，不判定）」，不計入。"
                           "下方「層級之間」的產業部署 GW 差距只列參考，不計入。對帳單由主控與 Siliconomics 層處理，本模型不改上游值。")))
    r_c3 = lab(C, "C0 結論")
    C.cell(r_c3, 1, "C0 結論：本模型的產業量（顆數、封裝面積、CPO 滲透率、匯率）全部讀 Siliconomics 基準，與層同源，只列「讀取層級」、不判定；公司實績（ATM 分部營收與封裝、測試分項）已列入 E 節，"
                    "但 Siliconomics 沒有封測服務、先進封裝的美元市場與日月光份額，LEAP 金額公司也未揭露，故 C0（實績）全為無可比實績，不產生對帳中項目。往上回饋只有證據（C5）：待 4Q26 財報與 2027 法說的 LEAP、資本支出實績，"
                    "以及層級補上 OSAT 市場與份額後再評估是否送對帳單。")
    return R


def readme_note(wb, R):
    RD = wb["README"]
    r = RD.max_row + 1
    RD.cell(r, 1, f"v{VER}（{BUILT}）：C0 實績對帳補列（三層基準協調規則 v1.4.2 C0 第 2 點；Andy 11-14 17:40）——Coord 頁新增 E 節「C0（實績）：公司實績 vs 層級隱含值」（主讀數＝2026 年 1–3Q 累計、第二讀數＝3Q26 × 4）；"
                  "B 節個股讀入層級值的同源列改標「讀取層級（同源，不判定）」，不計入對帳中項目數。公司只公布 ATM 分部與其封裝、測試分項，LEAP 金額未揭露；Siliconomics 無封測服務或先進封裝美元市場與日月光份額，"
                  "故各列為無可比實績，對帳中項目 0。既有數值不變，目標價 264.40 與評等強力放空不變。")
    return r


def apply(wb, here, SIL):
    if os.environ.get("COORD_RUN") == "1" or SIL.COORD_ALT or "Coord" not in wb.sheetnames:
        return "v027_post：略過（另一側讀數建置或無 Coord 頁）"
    R = coord_actuals(wb, SIL)
    readme_note(wb, R)
    RD = wb["README"]
    a1 = RD["A1"].value
    assert f"個股模型 v{V26.VER}（AI 半導體" in a1, a1[:80]
    RD["A1"].value = a1.replace(f"個股模型 v{V26.VER}（AI 半導體", f"個股模型 v{VER}（AI 半導體", 1)
    a2 = str(RD["A2"].value)
    old = f"版本：v{V26.VER}（{V26.BUILT}；初版 v0.1，2026-10-31）。"
    assert a2.startswith(old), a2[:50]
    RD["A2"].value = a2.replace(old, f"版本：v{VER}（{BUILT}；初版 v0.1，2026-10-31）。", 1)
    return f"v027_post：Coord E 節 {R['first']}–{R['last']} 列（C0 實績）；B 節改標同源"
