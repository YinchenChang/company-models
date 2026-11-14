# 台達電 v0.5.7（2026-11-14）後處理：C0 實績對帳（規則 v1.4.2 C0 第 2 點；工作單 docs/workorders/20261114_C0實績對帳補列_七家.md）
# 只動 Coord 頁與 README 版本文字，不動任何計算格：
#   1. 既有「本模型 vs 層級值」四項（部署 GW、AI 電源市場、有效份額、每 kW 內容）同源，差距恆為 0：判定欄改標「讀取層級（同源，不判定）」，不計入對帳中；
#   2. Coord 頁最下方新增第 4 節「C0（實績）：公司實績 vs 層級隱含值」：月營收（2026-01～10）、3Q26 法說揭露的占比與分部營收，對 Siliconomics v0.19 AI 電源市場 × 台達電份額的隱含營收；
#   3. README A1 與 v1.3 白話列的版本文字升 v0.5.7。
# build_delta.py 在 v056_post.apply 之後、存檔前呼叫 apply(wb, P56)。alt_run.py 建置（無 Coord 頁）不呼叫。
from openpyxl.styles import Font, PatternFill

VER, BUILT = "0.5.7", "2026-11-14"
BOLD = Font(bold=True); BLUE = Font(color="0000FF"); HDR = PatternFill("solid", fgColor="DDEBF7")
SAME = "讀取層級（同源，不判定）"
M_OCT = "https://finance.technews.tw/2026/11/09/delta-2308-202610-financial-report/"
M_SEP = "https://finance.technews.tw/2026/10/12/delta-2308-202609-financial-report/"
M_JUL = "https://statementdog.com/news/18020"
TR3 = "https://filecenter.deltaww.com/IR/download/calendar/3Q26_Transcript.pdf"
SL3 = "https://filecenter.deltaww.com/ir/download/calendar/3Q26_Analyst Meeting.pdf"
TOL = 0.15


def lab(ws, prefix, start=1, col=1, end=None):
    for r in range(start, (end or ws.max_row) + 1):
        v = ws.cell(r, col).value
        if isinstance(v, str) and v.strip().startswith(prefix):
            return r
    raise KeyError((ws.title, prefix))


def lit(t):
    t = t.replace('"', '""')
    return "&".join('"' + t[i:i + 200] + '"' for i in range(0, len(t), 200)) or '""'


def coord_c0_actuals(wb, G):
    C = wb["Coord"]; Q = wb["Q"]; U = wb["Units"]
    # ---- 1. 同源四項：判定欄改標 ----
    for k, nm in (("gw", "部署 GW（IT 電力）"), ("mkt", "AI 電源市場（US$B）"), ("sh", "台達電有效份額"), ("pkw", "每 kW 電源內容")):
        rj = G[k][3]
        assert str(C.cell(rj, 1).value).endswith("：判定（±15%）"), C.cell(rj, 1).value
        C.cell(rj, 1).value = f"{nm}：判定（同源，不判定）"
        for col in range(2, 6):
            C.cell(rj, col).value = SAME
    r2 = lab(C, "2. C0 由下往上對帳")
    C.cell(r2, 1).value = ("2. 讀取層級（同源，不判定；v0.5.7 起）：本模型自己的產業量與 Siliconomics 層級值逐年對照。個股的量直接讀自 Siliconomics，差距恆為 0、無檢驗力，"
                           "不計入 C0 對帳中項目；C0 判定見下方第 4 節「C0（實績）」")
    s1 = C.cell(G["sum1"], 1).value
    old = "（±15% 內，一致）；已建模"
    assert old in s1, s1
    C.cell(G["sum1"], 1).value = s1.replace(old, "（同源，差距恆為 0，不判定）；已建模")

    # ---- 4. C0（實績） ----
    r = G["sum4"] + 2
    C.cell(r, 1, "4. C0（實績）：公司實績 vs Siliconomics 層級隱含值（規則 v1.4.2 C0 第 2 點；工作單 2026-11-14；差距＝實績 ÷ 層級隱含值 − 1）").font = BOLD
    R = {"title": r}; r += 1
    C.cell(r, 1, ("方法：個股側＝台達電已公布實績或法說揭露；層級側＝Siliconomics 現行版 AI 電源市場 × 台達電有效份額（連到本頁層級值列）換成新台幣。層級值是日曆年，"
                  "故主讀數＝2026-01～10 累計實績 vs 層級 CY2026 × 10 ÷ 12（線性時差近似，Derived）；第二讀數＝3Q26 × 4 vs 層級 CY2026。"
                  "判定：兩讀數差距都超過 ±15% 且同號 → 對帳中；只有一個超過 → 時差邊界（參考）；都在內 → 一致。"
                  "公司沒有揭露 AI 電源金額，只在 3Q26 法說口述占比（電源相關約 25%、AI 資料中心相關約 50%、液冷約 NT$25B 即 13%）：以占比 × 公司已公布營收換算，期間未明，套用於累計與當季屬 Assumed。"))
    R["method"] = r; r += 1
    C.cell(r, 1, ("限制：本模型的台達電有效份額（讀自 Siliconomics 的 2026 年份額）本來就由「FY2026 AI 相關占比 26% × 營收錨點」校準（Derived），因此第一列主要檢驗「營收錨點與占比讀法」，"
                  "不獨立檢驗 Siliconomics 市場大小（顆數 × 每 kW 內容）；公司若揭露 AI 電源金額，才能獨立檢驗市場大小。"))
    R["limit"] = r; r += 2

    def hdr(cols):
        nonlocal r
        for j, h in enumerate(cols, 1):
            c = C.cell(r, j, h); c.font = BOLD; c.fill = HDR
        r += 1

    def row(label, b=None, c=None, unit="", note="", fmt="#,##0.000", blue=False, cfmt=None, url=""):
        nonlocal r
        C.cell(r, 1, label)
        for col, v, f in ((2, b, fmt), (3, c, cfmt or fmt)):
            if v is None:
                continue
            cell = C.cell(r, col, v)
            cell.number_format = f
            if blue and not (isinstance(v, str) and v.startswith("=")):
                cell.font = BLUE
        if unit:
            C.cell(r, 6, unit)
        if note:
            C.cell(r, 7, note)
        if url:
            C.cell(r, 8, url)
        r += 1
        return r - 1

    # 4a 公司實績輸入
    C.cell(r, 1, "4a. 公司已公布實績與揭露（輸入；藍字＝來自原文的常數）").font = BOLD; r += 1
    hdr(["項目", "值", "", "", "", "單位", "口徑、來源、證據等級", "來源網址"])
    R["jul"] = row("月營收：2026-07", 67.073, None, "NT$B", "Verified：月營收公告 670.73 億元（statementdog 2026-11-10 讀取；標題轉述）", blue=True, url=M_JUL)
    R["sep"] = row("月營收：2026-09", 65.400, None, "NT$B", "Verified：月營收公告 654.00 億元，累計前 9 月 5,396.68 億元（TechNews 2026-10-12）", blue=True, url=M_SEP)
    R["oct"] = row("月營收：2026-10", 74.666, None, "NT$B", "Verified：月營收公告 746.66 億元，月增 14.2%、年增 30.1%（TechNews 2026-11-09）", blue=True, url=M_OCT)
    R["aug"] = row("月營收：2026-08（3Q26 營收 − 7 月 − 9 月）", f"=q3_rev-B{R['jul']}-B{R['sep']}", None, "NT$B", "Derived：In 頁 q3_rev（3Q26 月營收合計 197.059）扣 7 月、9 月；與 9 月月增 1.3% 相容")
    R["cum9"] = row("累計：2026-01～09", 539.668, None, "NT$B", "Verified：月營收公告 5,396.68 億元（TechNews 2026-10-12）；3Q26 法說逐字稿 YT3Q26 NT$539.7B", blue=True, url=M_SEP+" ；"+TR3)
    R["cum10"] = row("累計：2026-01～10（主讀數的實績基礎）", 614.334, None, "NT$B", "Verified：月營收公告 6,143.34 億元，年增 36.3%（TechNews 2026-11-09）", blue=True, url=M_OCT)
    R["chk1"] = row("檢查：累計 9 月 + 10 月 − 累計 10 月（應為 0）", f"=B{R['cum9']}+B{R['oct']}-B{R['cum10']}", None, "NT$B", "Derived")
    R["chk2"] = row("檢查：累計 9 月 − 3Q26 − Q 頁 1H26 營收（應為 0；四捨五入差 < 0.01）", f"=B{R['cum9']}-q3_rev-Q!C{lab(Q, '營收')}", None, "NT$B", "Derived：1H26 342.6（Q 頁，公司法說）")
    R["q3"] = row("3Q26 營收（最近一季；第二讀數的實績基礎）", "=q3_rev", None, "NT$B", "Verified：In 頁 q3_rev；3Q26 法說逐字稿 1.1 節「Revenue reached NT$197.1bn, up 8% QoQ and 31% YoY」", url=TR3)
    R["pe"] = row("部門營收：電源及零組件，3Q26", 105.930, None, "NT$B", "Verified：3Q26 法說簡報部門營收表 105,930 百萬元，年增 40%、季增 10%（經 WebFetch 抽取，未逐頁核對 PDF）", blue=True, url=SL3)
    R["inf"] = row("部門營收：基礎設施，3Q26", 65.880, None, "NT$B", f"Verified：同上 65,880 百萬元，年增 25%、季增 5%", blue=True)
    R["psh"] = row("公司口述：電源相關解決方案占營收", 0.25, None, "%", "Verified（公司口述、約數；期間未明）：3Q26 法說逐字稿 Q11 CFO「The previously mentioned 25% figure refers specifically to power-related solutions」；2Q26 法說稱「一定超過 25%」", fmt="0.0%", blue=True, url=TR3)
    R["ash"] = row("公司口述：AI 資料中心相關業務合計占營收（PSBG、FMBG、CP、ICTBG 四個事業單位）", 0.50, None, "%", "Verified（公司口述、約數；期間未明）：同上 Q11 Chairman「collectively accounting for approximately 50% of total revenue」（公司未定義「AI 相關」）", fmt="0.0%", blue=True, url=TR3)
    R["lq"] = row("公司口述：液冷營收，3Q26", 25.0, None, "NT$B", "Verified（公司口述、約數）：CFO「Liquid cooling revenue reached approximately NT$25bn in 3Q26, representing around 13% of total revenue」，全年預期類似占比（同上 Q11）", fmt="#,##0.0", blue=True)
    R["lqs"] = row("公司口述：液冷占營收，3Q26", 0.13, None, "%", "同上", fmt="0.0%", blue=True)
    R["months"] = row("已公布月數（2026-01～10）", 10, None, "月", "輸入：月營收至 2026-10（2026-11-09 公告）", fmt="0", blue=True)
    r += 1

    # 4b 層級隱含值
    C.cell(r, 1, "4b. 層級隱含值（Siliconomics 現行版；連到本頁層級值列）").font = BOLD; r += 1
    hdr(["項目", "CY2026", "", "", "", "單位", "口徑、來源、證據等級"])
    R["lmkt"] = row("AI 電源市場（層級值）", f"=B{G['mkt'][1]}", None, "US$B", "Derived：本頁第 2 節 Siliconomics SIF_PwrMkt_Base（CY2026）", fmt="#,##0.000")
    R["lsh"] = row("台達電有效份額（層級值）", f"=B{G['sh'][1]}", None, "%", "Interested-party／Assumed：本頁第 2 節層級值快照（pwr_dsh_26；由 FY2026 AI 相關占比 26% 校準）", fmt="0.0%")
    R["lpwr"] = row("層級隱含：台達電 AI 電源營收（市場 × 份額 × 匯率）", f"=B{R['lmkt']}*B{R['lsh']}*fx", None, "NT$B", "Derived：匯率 In 頁 fx", fmt="#,##0.0")
    R["lliq"] = row("本模型 Units：台達電液冷營收（CY2026；非 Siliconomics）", f"=Units!C{lab(U, '台達電液冷營收')}*fx", None, "NT$B",
                    "Derived：本模型 Units 頁（部署 GW × 每 GW 液冷營收，FY2026 以液冷占比 12.5% 校準）；Siliconomics M19 液冷市場不含台達電機架外 CDU 與 Liquid-to-Air，故不是層級值", fmt="#,##0.0")
    r += 1

    # 4c C0（實績）對帳表
    C.cell(r, 1, "C0（實績）：公司實績 vs 層級隱含值（HTML 用）").font = BOLD; R["tbl"] = r; r += 1
    hdr(["項目", "主讀數：2026-01～10 累計 vs 層級 CY2026 × 10 ÷ 12", "第二讀數：3Q26 × 4 vs 層級 CY2026", "", "", "單位", "口徑、來源、證據等級"])
    R["hdr"] = r - 1
    fr = f"B{R['months']}/12"
    pct = "+0.0%;-0.0%"

    def judge(rg):
        return (f'=IF(AND(ABS(B{rg})>{TOL},ABS(C{rg})>{TOL},B{rg}*C{rg}>0),"對帳中",'
                f'IF(OR(ABS(B{rg})>{TOL},ABS(C{rg})>{TOL}),"時差邊界（參考）","一致"))')

    # 列 1：AI 電源（判定）
    a = row("AI 電源（電源相關約 25%）：公司實績換算", f"=B{R['cum10']}*B{R['psh']}", f"=B{R['q3']}*B{R['psh']}*4", "NT$B",
            "Verified 營收 × 公司口述占比（Assumed 套用於累計與當季）；第二讀數＝3Q26 × 占比 × 4", fmt="#,##0.0")
    l = row("AI 電源（電源相關約 25%）：層級隱含", f"=B{R['lpwr']}*{fr}", f"=B{R['lpwr']}", "NT$B", "Derived：層級隱含 AI 電源營收（4b）；主讀數 × 10 ÷ 12", fmt="#,##0.0")
    g = row("AI 電源（電源相關約 25%）：差距（實績 ÷ 層級隱含 − 1）", f"=B{a}/B{l}-1", f"=C{a}/C{l}-1", fmt=pct)
    R["gap_pwr"] = g
    R["j_pwr"] = row("AI 電源（電源相關約 25%）：判定（±15%）", judge(g), None, "", "兩讀數都在內 → 一致", fmt="General")
    # 敏感度：28%
    s28 = row("AI 電源敏感度（參考）：占比取 28% 時的差距（BigGo 摘要稱朝 28% 目標，非實績）", f"=B{R['cum10']}*0.28/B{l}-1", f"=B{R['q3']}*0.28*4/C{l}-1", "", "Interested-party：占比 28% 只作敏感度，不用於判定", fmt=pct)
    R["s28"] = s28
    # 列 2：AI 資料中心合計（口徑不同，不判定）
    a2 = row("AI 資料中心合計（約 50%）：公司實績換算", f"=B{R['cum10']}*B{R['ash']}", f"=B{R['q3']}*B{R['ash']}*4", "NT$B", "Verified 營收 × 公司口述占比（Assumed）", fmt="#,##0.0")
    l2 = row("AI 資料中心合計：層級隱含（AI 電源 + 本模型液冷；層級無其他 AI 項目）", f"=(B{R['lpwr']}+B{R['lliq']})*{fr}", f"=B{R['lpwr']}+B{R['lliq']}", "NT$B", "Derived：4b 兩項相加；液冷為本模型 Units（非層級值）", fmt="#,##0.0")
    g2 = row("AI 資料中心合計：差距（實績 ÷ 層級隱含 − 1）", f"=B{a2}/B{l2}-1", f"=C{a2}/C{l2}-1", fmt=pct)
    R["gap_ai"] = g2
    R["j_ai"] = row("AI 資料中心合計：判定", "口徑不符，不判定（參考）", None, "",
                    "公司「AI 相關」含 UPS、網通、零組件等層級沒有市場的項目，且未拆分；差距由口徑差解釋，不是層級值有誤的證據；不計入對帳中", fmt="General")
    R["res"] = row("口徑落差：AI 資料中心占比 − 電源占比 − 液冷占比（層級未涵蓋的 AI 相關營收占比）", f"=B{R['ash']}-B{R['psh']}-B{R['lqs']}", None, "%", "Derived：公司口述三個占比相減；不得歸零，公司未揭露金額", fmt="0.0%")
    # 列 3：液冷（單一讀數）
    a3 = row("液冷：公司實績換算（3Q26 約 NT$25B × 4；累計液冷金額未揭露）", "無（1–10 月累計液冷金額未揭露）", f"=B{R['lq']}*4", "NT$B", "Verified（約數）：3Q26 液冷營收（4a）；主讀數缺揭露，不以模型值補位", fmt="#,##0.0")
    l3 = row("液冷：本模型 Units（非層級值）", None, f"=B{R['lliq']}", "NT$B", "Derived：4b", fmt="#,##0.0")
    g3 = row("液冷：差距（第二讀數；實績 ÷ 本模型 − 1）", None, f"=C{a3}/C{l3}-1", fmt=pct)
    R["gap_liq"] = g3
    R["j_liq"] = row("液冷：判定", "單一讀數，參考（層級側為本模型，非 Siliconomics；不計入對帳中）", None, "", "Siliconomics 液冷市場不含台達電；本列只檢驗 Units 校準", fmt="General")
    # 列 4：上限檢查
    a4 = row("上限檢查：電源及零組件部門營收（3Q26 × 4；Verified 部門數）", None, f"=B{R['pe']}*4", "NT$B", "Verified：3Q26 法說簡報部門營收表；部門含非 AI（一般伺服器電源、消費與工業電源、零組件）", fmt="#,##0.0")
    R["ub"] = row("上限檢查：層級隱含 AI 電源 ÷ 電源及零組件部門年化（應 ≤ 100%）", None, f"=B{R['lpwr']}/C{a4}", "%", "Derived：層級 AI 電源不得超過其所屬部門；通過＝合理，不是一致性判定", fmt="0.0%")
    R["j_ub"] = row("上限檢查：結果", None, f'=IF(C{R["ub"]}<=1,"通過","不通過")', "", "", fmt="General")
    r += 1
    # 本模型營收錨點對月營收（參考）
    C.cell(r, 1, "參考：本模型 FY2026 營收錨點對月營收實績（模型對實績，不是層級對帳；不計入 C0）").font = BOLD; r += 1
    R["anc"] = row("FY2026 營收錨點（Q 頁）", f"=Q!F{lab(Q, '營收')}", None, "NT$B", "Derived／Interested-party：1H26 實績 + 3Q26 實績 + 4Q 假設（季增 10%）", fmt="#,##0.0")
    R["imp"] = row("錨點隱含 11–12 月月均營收（錨點 − 1～10 月累計）÷ 2", f"=(B{R['anc']}-B{R['cum10']})/2", None, "NT$B", "Derived", fmt="#,##0.0")
    R["impg"] = row("隱含月均 ÷ 10 月實績 − 1", f"=B{R['imp']}/B{R['oct']}-1", None, "%", "Derived：負值＝錨點假設 11–12 月低於 10 月水準（10 月為單月歷史新高）；法說稱 4Q 營收一定高於 3Q", fmt=pct)
    r += 1
    # 結果摘要
    C.cell(r, 1, "C0（實績）結果摘要（HTML 用；公式帶數字）").font = BOLD; r += 1
    nopen = f'COUNTIF(B{R["j_pwr"]},"對帳中")'
    C.cell(r, 1, "=" + "&".join([
        lit("C0（實績）對帳中項目 "), f'{nopen}', lit(" 項。AI 電源（公司口述電源相關約 25% 換算）：主讀數差距 "), f'TEXT(B{R["gap_pwr"]},"+0.0%;-0.0%")',
        lit("、第二讀數差距 "), f'TEXT(C{R["gap_pwr"]},"+0.0%;-0.0%")', lit("，判定"), f'B{R["j_pwr"]}',
        lit("；占比取 28% 時為 "), f'TEXT(B{R["s28"]},"+0.0%;-0.0%")', lit("／"), f'TEXT(C{R["s28"]},"+0.0%;-0.0%")', f'IF(MAX(ABS(B{R["s28"]}),ABS(C{R["s28"]}))<={TOL},"（仍在範圍內）。","（超出範圍）。")']))
    R["sum1"] = r; r += 1
    C.cell(r, 1, "=" + "&".join([
        lit("口徑落差（參考，不判定）：公司口述 AI 資料中心相關約 50% 換成金額比層級的 AI 電源加液冷高 "), f'TEXT(B{R["gap_ai"]},"0%")', lit("／"), f'TEXT(C{R["gap_ai"]},"0%")',
        lit("，差在公司口徑含層級沒有市場的 AI 相關業務（占營收約 "), f'TEXT(B{R["res"]},"0%")', lit("，金額未揭露）；液冷 3Q26 年化對本模型 Units 差 "), f'TEXT(C{R["gap_liq"]},"+0.0%;-0.0%")',
        lit("（單一讀數）。電源及零組件部門年化 "), f'TEXT(C{a4},"#,##0")', lit(" 十億元，層級 AI 電源占其 "), f'TEXT(C{R["ub"]},"0%")', lit("，上限檢查"), f'C{R["j_ub"]}', lit("。")]))
    R["sum2"] = r; r += 1
    C.cell(r, 1, ("限制與待查：公司未揭露 AI 電源金額，占比為口述約數；本版層級份額由 26% 占比校準，故一致不等於市場大小已獨立驗證；"
                  "3Q26 法說「AI 相關約 50%」與 2Q26 的「超過 25%」口徑不同，是否需重校準份額屬 Siliconomics 對帳單範圍（C5：個股不改上游值）。"))
    R["sum3"] = r; r += 1
    R["end"] = r - 1
    return R


def version_text(wb, RM):
    RD = wb["README"]
    a1 = RD["A1"].value
    old = "個股模型 v0.5.6（"
    assert old in a1, a1[:120]
    a1 = a1.replace(old, f"個股模型 v{VER}（", 1)
    RD["A1"].value = a1 + (f"；v{VER}（{BUILT}）：C0 改以公司實績對帳（規則 v1.4.2；Coord 頁第 4 節 C0（實績））：月營收 2026-01～10 與 3Q26 法說占比、部門營收對 Siliconomics v0.19 AI 電源市場 × 台達電份額；"
                           "既有同源四項改標「讀取層級（同源，不判定）」；只動 Coord 頁與版本文字，數值與目標價不變")
    rv = RD.cell(RM["ver"], 1)
    t = rv.value
    assert isinstance(t, str) and t.startswith("v0.5.6"), t[:40]
    rv.value = (f"v{VER}（{BUILT}）：C0 改以公司實績對帳：Coord 頁新增「C0（實績）」（月營收 2026-01～10、3Q26 法說占比與部門營收對 Siliconomics AI 電源市場 × 份額），同源四項改標讀取層級、不判定；數值與目標價不變。"
                "前一版：" + t)


def apply(wb, P56):
    v = wb._v056
    R = coord_c0_actuals(wb, v["G"])
    version_text(wb, v["RM"])
    wb._v057 = dict(R=R)
    return f"v057_post：Coord 第 4 節 C0（實績）第 {R['title']}–{R['end']} 列；對帳表第 {R['hdr']} 列起"
