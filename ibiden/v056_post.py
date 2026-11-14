# Ibiden v0.5.6（2026-11-14）後處理：C0 補「公司實績 vs 層級隱含值」對帳列（三層基準協調規則 v1.4.2 C0 第 2 點；工作單 docs/workorders/20261114_C0實績對帳補列_七家.md）。
# 只新增與改標，不改既有數值格：
#   Coord 頁「2. C0」節：既有同源列（個股讀入的層級值 vs 層級值，差距恆為 0 或同源推得）判定欄改標「讀取層級（同源，不判定）」，不計入對帳中項目數；
#     對帳中項目數改為計「4. C0（實績）」判定欄。
#   Coord 頁末：4. C0（實績）——電子事業營收（決算短信）vs Siliconomics SIF_abfrev_ib26 推得的同口徑電子事業營收；主讀數、第二讀數、判定（±15%）。
#   Sources 頁：新增 Q2 決算數字來源。README：A1 版號與本版說明。
import re
from openpyxl.styles import Font, PatternFill

VER, BUILT = "0.5.6", "2026-11-14"
BOLD = Font(bold=True); BLUE = Font(color="0000FF"); HDR = PatternFill("solid", fgColor="DDEBF7")
SAME = "讀取層級（同源，不判定）"
SRC_Q1 = ("公司：2027 年 3 月期 Q1 決算短信（TDnet 140120260731504718，2026-08-04；Sources 頁第 3 列）セグメント情報 電子事業 775.22 億日圓；"
          "與 note.com aimark 決算預習引述 775.22 億、In 頁 f_e 備註 Q1 實績 77.5 相符")
SRC_H1 = ("公司：2027 年 3 月期 第 2 四半期（中間期）決算短信（TDnet 140120261028557691，2026-10-29；Sources 頁）セグメント情報 電子事業對外部顧客營收 4～9 月 172,540 百萬日圓（年增 52.0%）；原文直接取用。"
          "交叉核對：二手 Investing.com 報 Q2 電子事業 950 億，與 172.540 − 77.522 ＝ 95.018 相符；全公司累計營收 2,625 億、營業利益 601 億與 edinetdb.jp 相符")
SRC_GUIDE = ("二手轉載：同 Investing.com 文——FY2026 全年指引上修（10-29）：營收 5,500 → 5,650 億、電子事業營收 3,750 → 3,850 億、電子事業營業利益 1,100 → 1,140 億日圓；"
             "路透經 Newsweek Japan 只報營業利益 1,270 → 1,330 億、淨利 840 → 890 億。公司指引不是實績，只作參考")


def apply(wb, here, SIL, CO):
    CW = wb["Coord"]
    V = SIL.V["abfrev_ib26"]
    first, last, cnt = CO["first"], CO["last"], CO["cnt"]
    # ---- 既有「2. C0」節：同源列改標，不判定 ----
    for r in range(first, last + 1):
        g = CW.cell(r, 7).value
        if isinstance(g, str) and g.startswith("=IF(ABS(F") and '"對帳中"' in g:
            CW.cell(r, 7).value = SAME
    h2 = CO["hdr2"]
    CW.cell(h2, 7).value = "判定（同源，不判定）"
    # 節標題與說明列：保留「2. C0」開頭（HTML 與工作單以此定位），改述為讀取層級核對
    for r in range(first - 4, first):
        v = CW.cell(r, 1).value
        if isinstance(v, str) and v.startswith("2. C0"):
            CW.cell(r, 1).value = ("2. C0 讀取層級核對（個股讀入的層級值 vs 層級值：同源，差距恆為 0 或由同一層級值推得，無檢驗力，故不判定、不計入對帳中；"
                                   "C0 判定見本頁「4. C0（實績）」，規則 v1.4.2 C0 第 2 點）")
    # ---- 4. C0（實績）----
    r = CW.max_row + 3
    CW.cell(r, 1, "4. C0（實績）：公司實績 vs 層級隱含值（v0.5.6；規則 v1.4.2 C0 第 2 點、工作單 20261114_C0實績對帳補列；差距超過 ±15% 且兩個讀數同號才標「對帳中」；不改上游值、不開改值 PR）").font = BOLD
    ROW = {"title": r}; r += 1
    CW.cell(r, 1, "方法：個股側＝公司已公布實績（決算短信 FY2026 Q1、中間期＝2026-04～2026-09，只用已公布期間，不用預測或指引；實績本身是公司原文、獨立）；層級側＝Siliconomics 現行基準的 Ibiden 2026 ABF 載板營收（US$B；鍵名見輸入表 I 欄）"
                  "換算成同口徑電子事業營收（× 匯率 ÷ ABF 占電子事業比率）。層級值是日曆年 CY2026：主讀數＝已公布期間累計實績 vs 層級 × 已公布月數 ÷ 12（線性時差近似，Derived）；"
                  "第二讀數＝最近一季 × 4 vs 層級；兩讀數的差距都超過 ±15% 且同號為「對帳中」，只有一個超過為「時差邊界（參考）」，都在內為「一致」。成長年中主讀數偏低、第二讀數偏高，兩者夾住真值。")
    ROW["method"] = r; r += 1
    CW.cell(r, 1, "公司實績 vs 層級隱含值的同口徑對象是電子事業營收（含 ABF 載板與其他 PCB 等）；IC 載板單獨營收：無可比實績（公司只揭露電子事業整體，未揭露載板或 ABF 分項）。")
    ROW["ic"] = r; r += 1
    for j, h in enumerate(["項目", "單位", "期間", "數值", "等級", "", "", "出處／口徑"], 1):
        c = CW.cell(r, j, h); c.font = BOLD; c.fill = HDR
    ROW["hdr_in"] = r; r += 1
    inp = {}
    KEYS = {k: f"{n}（D 欄為寫死值，讀自 Siliconomics {SIL.VERSION} 快照 {SIL.FILE}）" for k, n in (("abf", "SIF_abfrev_ib26"), ("abf_lo", "SIF_abfrev_ib26_lo"), ("abf_hi", "SIF_abfrev_ib26_hi"))}
    KEYS["sh"] = f"85%：寫死值，取自 Siliconomics {SIL.VERSION} abfrev_ib26 列原註"
    KEYS["fx"] = "In 頁 fx（連格）"
    CW.cell(ROW["hdr_in"], 9, "鍵名（追蹤用；HTML 不顯示）").font = BOLD

    def put(key, label, unit, per, val, grade, src, fmt):
        nonlocal r
        CW.cell(r, 1, label); CW.cell(r, 2, unit); CW.cell(r, 3, per)
        c = CW.cell(r, 4, val); c.number_format = fmt
        if not (isinstance(val, str) and val.startswith("=")): c.font = BLUE
        CW.cell(r, 5, grade); CW.cell(r, 8, src)
        if key in KEYS: CW.cell(r, 9, KEYS[key])
        inp[key] = r; r += 1
    put("h1", "電子事業營收：FY2026 Q1–Q2 累計（4～9 月）", "¥B", "2026-04～2026-09", 172.540, "Verified", SRC_H1, "#,##0.000")
    put("q1", "電子事業營收：FY2026 Q1", "¥B", "2026-04～2026-06", 77.522, "Verified", SRC_Q1, "#,##0.000")
    put("q2", "電子事業營收：FY2026 Q2（＝累計 − Q1）", "¥B", "2026-07～2026-09", f"=D{inp['h1']}-D{inp['q1']}", "Derived", "中間期累計（決算短信）減 Q1（決算短信）；與二手 Investing.com 的 950 億相符", "#,##0.000")
    put("mo", "已公布月數", "月", "2026-04～2026-09", 6, "Derived", "主讀數時差換算：層級 × 月數 ÷ 12", "0")
    put("abf", "Siliconomics 層級值：Ibiden 2026 ABF 載板營收（基準）", "US$B", "CY2026", V["base"], "Assumed（層級值）",
        f"Siliconomics {SIL.VERSION}（快照檔 {SIL.FILE}；鍵名見 I 欄）；該列標示 2026，實為由本模型 FY2026（2026-04～2027-03）電子事業營收預測推得，與日曆年 CY2026 有時差；該列原註：Ibiden 模型 FY2026 電子事業營收預測 ¥375B × ABF 占電子事業 85% ÷ ¥152", "0.000")
    put("abf_lo", "Siliconomics 層級值：Ibiden 2026 ABF 載板營收（下界）", "US$B", "CY2026", V["lo"], "Assumed（層級值）", "同上（下界；只作參考，不判定）", "0.000")
    put("abf_hi", "Siliconomics 層級值：Ibiden 2026 ABF 載板營收（上界）", "US$B", "CY2026", V["hi"], "Assumed（層級值）", "同上（上界；只作參考，不判定）", "0.000")
    put("fx", "匯率（In 頁 fx）", "¥/US$", "FY2026 前提", "=fx", "Interested-party", "公司 Q1 補足資料前提 ¥152；與 Siliconomics 該列換算同值", "0.0")
    put("sh", "ABF 占電子事業營收比率", "%", "FY2026", 0.85, "Assumed", "Siliconomics abfrev_ib26 列原註（範圍 75%–95%，公司未揭露）；用來把 ABF 營收換算回電子事業口徑", "0.0%")
    put("lay", "層級隱含電子事業營收（基準）", "¥B", "CY2026", f"=D{inp['abf']}*D{inp['fx']}/D{inp['sh']}", "Derived", "層級 ABF 營收 × 匯率 ÷ ABF 占比；連到 D 欄各輸入格", "#,##0.0")
    put("lay_lo", "層級隱含電子事業營收（下界）", "¥B", "CY2026", f"=D{inp['abf_lo']}*D{inp['fx']}/D{inp['sh']}", "Derived", "同上，用 _lo", "#,##0.0")
    put("lay_hi", "層級隱含電子事業營收（上界）", "¥B", "CY2026", f"=D{inp['abf_hi']}*D{inp['fx']}/D{inp['sh']}", "Derived", "同上，用 _hi", "#,##0.0")
    put("guide", "參考：公司 FY2026 全年電子事業營收指引（10-29 上修後）", "¥B", "FY2026", 385.0, "Interested-party（二手）", SRC_GUIDE, "#,##0.0")
    r += 1
    for j, h in enumerate(["讀數", "單位", "期間", "公司實績", "層級隱含", "差距（實績 ÷ 層級 − 1）", "判定（±15%）", "口徑與說明"], 1):
        c = CW.cell(r, j, h); c.font = BOLD; c.fill = HDR
    ROW["hdr_main"] = r; r += 1

    def line(label, per, f_a, f_l, note, judge="（讀數）"):
        nonlocal r
        CW.cell(r, 1, label); CW.cell(r, 2, "¥B"); CW.cell(r, 3, per)
        CW.cell(r, 4, f_a).number_format = "#,##0.0"; CW.cell(r, 5, f_l).number_format = "#,##0.0"
        CW.cell(r, 6, f"=D{r}/E{r}-1").number_format = "+0.0%;-0.0%"
        CW.cell(r, 7, judge); CW.cell(r, 8, note)
        r += 1
        return r - 1
    ra = line("主讀數：已公布期間累計電子事業營收 vs 層級 × 月數 ÷ 12", "2026-04～09 vs CY2026 × 6/12", f"=D{inp['h1']}", f"=D{inp['lay']}*D{inp['mo']}/12",
              "實績＝中間期決算短信累計（4～9 月）；層級＝Siliconomics Ibiden ABF 載板營收換算的電子事業營收 × 6 ÷ 12（線性時差近似）。成長年中，上半年占全年比重低於一半，此讀數偏低")
    rb = line("第二讀數：最近一季（Q2）× 4 vs 層級", "2026-07～09 × 4 vs CY2026", f"=D{inp['q2']}*4", f"=D{inp['lay']}",
              "最近一季年化；成長年中此讀數偏高。兩讀數夾住真值")
    CW.cell(r, 1, "C0（實績）結論")
    CW.cell(r, 7, "無可比實績（Siliconomics 層級值）")
    CW.cell(r, 8, "層級值由本公司預測推得（375 × 85% ÷ 152，85% 與匯率往返抵銷），故本表等於「實績 vs 公司上修前電子事業預測」，不檢驗 ABF 占比與 abfrev_ib26；缺 IC 載板單獨營收揭露。實績獨立，不獨立的是層級值。對帳中維持 0，不開對帳單")
    for k_ in (ra, rb): CW.cell(k_, 7, "參考（層級值同源，不檢驗 ABF 占比與 abfrev_ib26）")
    ROW["judge"] = r; r += 1
    CW.cell(r, 1, "參考：機械判定（兩讀數差距都超過 ±15% 且同號＝對帳中；只一個＝時差邊界；都在內＝一致）；因層級值同源，僅供參考")
    CW.cell(r, 7, f'="參考："&IF(AND(ABS(F{ra})>0.15,ABS(F{rb})>0.15,F{ra}*F{rb}>0),"對帳中",IF(OR(ABS(F{ra})>0.15,ABS(F{rb})>0.15),"時差邊界","一致"))')
    r += 1
    for key, nm in (("lay_lo", "下界"), ("lay_hi", "上界")):
        line(f"參考：主讀數（層級{nm}）", "2026-04～09 vs CY2026 × 6/12", f"=D{inp['h1']}", f"=D{inp[key]}*D{inp['mo']}/12", "層級值範圍（lo／hi）下的同一比較；不判定", "參考（不判定）")
        line(f"參考：第二讀數（層級{nm}）", "2026-07～09 × 4 vs CY2026", f"=D{inp['q2']}*4", f"=D{inp[key]}", "同上；不判定", "參考（不判定）")
    CW.cell(r, 1, "參考：已公布期間實績占公司全年電子事業營收指引（10-29 上修後；二手）")
    CW.cell(r, 2, "¥B"); CW.cell(r, 3, "2026-04～09 vs FY2026")
    CW.cell(r, 4, f"=D{inp['h1']}").number_format = "#,##0.0"; CW.cell(r, 5, f"=D{inp['guide']}").number_format = "#,##0.0"
    CW.cell(r, 6, f"=D{r}/E{r}").number_format = "0.0%"
    CW.cell(r, 7, "參考（不判定）"); CW.cell(r, 8, "本列的比值是進度（上半年 ÷ 全年指引），不是差距；公司指引不算實績。層級值本身由上修前指引 375 推得（見輸入列），故層級與公司指引同源，不構成獨立證據")
    r += 1
    CW.cell(r, 1, "讀法：公司實績是決算短信原文、本身獨立；層級值 abfrev_ib26 由本模型（公司上修前指引）FY2026 電子事業營收預測 375 × ABF 占比 85% ÷ 匯率 152 推得，本表換算回電子事業時 85% 與匯率往返抵銷，等於「實績 vs 公司上修前預測」。"
                  "所以這一列不檢驗 ABF 占比，也不檢驗 Siliconomics 的獨立估計，判定標參考；C0（實績）結論為無可比實績（層級值同源）。")
    ROW["read"] = r; r += 1
    CW.cell(r, 1, "對帳中項目數（C0 實績表）"); CW.cell(r, 7, f'=COUNTIF(G{ra}:G{r - 2},"對帳中")')
    ROW["cnt4"] = r; r += 1
    ROW["ra"], ROW["rb"], ROW["last"] = ra, rb, r - 1
    # 既有 2 節的計數改為 C0 實績表的判定（同源列不計）
    CW.cell(cnt, 1).value = "對帳中項目數（C0 對帳：同源列不判定、不計；以本頁「4. C0（實績）」判定為準）"
    CW.cell(cnt, 7).value = f'=COUNTIF(G{ROW["judge"]},"對帳中")'
    # ---- Sources 與 README ----
    SO = wb["Sources"]
    rr = SO.max_row + 1
    SO.cell(rr + 3, 1, "公司：2027 年 3 月期 第 2 四半期（中間期）決算短信（電子事業對外部顧客營收 4～9 月 172,540 百萬日圓；C0 實績用）"); SO.cell(rr + 3, 2, "https://finance.biggo.jp/news/jpx_tdnet_140120261028557691")
    SO.cell(rr, 1, "二手（交叉核對）：Investing.com 日本語 2026-10-29（Q2 電子事業營收 950 億、全公司 1,393 億、全年指引上修）"); SO.cell(rr, 2, "https://jp.investing.com/news/earnings/article-93CH-1732026")
    SO.cell(rr + 1, 1, "二手：edinetdb.jp FY2027 Q2（累計）全公司營收 2,625 億、營業利益 601 億（Q2 加總交叉核對用）"); SO.cell(rr + 1, 2, "https://edinetdb.jp/company/E00775/earnings/FY2027-Q2")
    SO.cell(rr + 2, 1, "二手：note.com aimark 決算預習（Q1 電子事業營收 775.22 億日圓，對決算短信）"); SO.cell(rr + 2, 2, "https://note.com/aimark/n/nfb8ab5b83937")
    RD = wb["README"]
    a1 = RD["A1"].value
    assert "模型 v0.5.5（AI 半導體 Project" in a1, a1[:80]
    RD["A1"].value = a1.replace("模型 v0.5.5（AI 半導體 Project", f"模型 v{VER}（AI 半導體 Project", 1) + (
        f"；v{VER}（{BUILT}）：三層協調 C0 補公司實績對帳（規則 v1.4.2）——Coord 頁新增「4. C0（實績）」電子事業營收實績 vs Siliconomics Ibiden ABF 載板營收隱含值；既有同源列改標「{SAME}」不計對帳中；既有數值不變")
    # README「本版要點」列（HTML 取用）：改為 v0.5.6 說明
    for row in range(RD.max_row, 0, -1):
        v = RD.cell(row, 1).value
        if isinstance(v, str) and v.startswith("v0.5.5（2026-11-14）：個股 HTML 版面規格 v1.3"):
            RD.cell(row, 1).value = (f"v{VER}（{BUILT}）：C0 補公司實績對帳（規則 v1.4.2）。Coord 頁新增「4. C0（實績）」：電子事業營收實績（決算短信 FY2026 Q1、中間期）對 Siliconomics Ibiden ABF 載板營收換算的同口徑隱含值，主讀數與第二讀數兩個讀數；層級值由本模型預測推得，判定標參考，結論「無可比實績（Siliconomics 層級值）」、"
                                     f"對帳中 0；既有同源列改標「{SAME}」，不計對帳中。承 v0.5.5 的個股 HTML 版面規格 v1.3（章節順序、標準損益表、反向 DCF、四個目標價、加權目標價相對基準敏感度）。既有數值格不變，加權目標價與評等不變")
            break
    else:
        raise AssertionError("README 找不到 v0.5.5 本版要點列")
    wb._v056 = ROW
    return f"v056_post：Coord C0（實績）{ROW['title']}–{ROW['last']}；判定列 {ROW['judge']}"
