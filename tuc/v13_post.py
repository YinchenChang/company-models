# 台燿 v0.2.4（2026-11-15）後處理：個股 HTML 版面規格 v1.3＋三層協調兩讀數＋C0（實績）（工作單 docs/workorders/20261114_第二批13家_v1.3與三層協調.md；
# 規則 docs/plan/三層基準協調規則.md v1.4.2；樣板 NVIDIA v0.3.10、台達電、同次產業台光電 emc/v13_post.py）。共通步驟在 tuc/c13.py（高速 CCL／PCB 四家同一份），本檔只放台燿的設定。
# build_tuc.py 在 Coord 頁建好之後、存檔前呼叫 apply(wb, SIL, coord_rows)。alt_run.py 與 sens_tuc.py 的建置不呼叫。只新增，不改既有數值格。
import os, sys
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import c13
from c13 import T, F, lab, blockB, Sheet

VER, BUILT = "v0.2.4", "2026-11-15"
SA = "https://stockanalysis.com/quote/tpex/6274/financials/income-statement/"
M_OCT = "https://finance.technews.tw/2026/11/05/tuc-6274-202610-financial-report/"
M_SEP = "https://finance.technews.tw/2026/10/05/tuc-6274-202609-financial-report/"
AIR = "高速材料營收（ELL／SLL，M7 以上；AI 相關）＝ MIN（需求面、產能上限 − 其他等級）"
NAIR = "其他等級營收（VLL／LL、非低損耗、特殊應用）"
REVR = "營收（單一應報導部門：銅箔基板與膠片）"
NIR = "稅後淨利（歸屬母公司）"

C = dict(
    name="台燿", ver=VER, built=BUILT, rev_label=REVR, ni_label=NIR,
    hist_title=("Hist — 損益表歷史（v0.2.4 起；PL 頁「標準損益表」的歷史三年與年增率用）。金額：十億新台幣；EPS：新台幣；歸屬母公司。"
                f"來源：stockanalysis 年度損益表（百萬元，S&P Global，2026-11-15 讀取；{SA}），FY2025 與 Q 頁一致。營業外收益＝稅前 − 營業利益；營業費用＝毛利 − 營業利益；"
                "stockanalysis 淨利與歸屬母公司淨利相同（無非控制權益），非控制權益損益列 0"),
    hist=[
        ("rev", "營收", 18.472, 16.003, 23.070, 30.340, "Verified", f"{SA}：Revenue（百萬元）；FY2025 同 Q 頁 30.340"),
        ("gp", "毛利", 3.405, 3.159, 5.342, 6.898, "Verified", "同上：Gross Profit"),
        ("oi", "營業利益", 1.498, 1.424, 3.332, 4.344, "Verified", "同上：Operating Income"),
        ("pti", "稅前利益", 1.621, 1.550, 3.378, 4.519, "Verified", "同上：Pretax Income"),
        ("tax", "所得稅", 0.360, 0.726, 0.773, 1.109, "Verified", "同上：Income Tax Expense（359.54／726.47／773.27／1,109 百萬元）"),
        ("nci", "非控制權益損益", 0.0, 0.0, 0.0, 0.0, "Verified", "同上：Net Income 與 Net Income to Common 相同（無非控制權益）；核對列差額為百萬元四捨五入"),
        ("ni", "稅後淨利（歸屬母公司）", 1.261, 0.823, 2.604, 3.409, "Verified", "同上：Net Income to Common（FY2023 823.44 百萬元）；FY2025 同 Q 頁"),
        ("sh", "EPS 換算股數（十億股）", 0.269, 0.270, 0.272, 0.281, "Verified", "同上：Shares Outstanding (Basic)（百萬股取整）"),
        ("eps", "EPS（公司公告，基本，NT$）", 4.69, 3.05, 9.56, 12.13, "Verified", "同上：EPS (Basic)（稀釋 4.49／3.02／9.49／11.61，差異來自可轉債）；FY2025 同 Q 頁 12.13"),
    ],
    fy26_short="FY2026E 含 1H26 與 3Q26 實績",
    fy26_long="1H26 財報＋3Q26 財報（2026-10 月底公布）＋4Q26 假設（公司未給指引；營收季增 5%、毛利率 36%）",
    eps_note="歷史為公司公告基本 EPS；預測為歸屬母公司淨利 ÷ 股數（股數＝淨利 ÷ 公告基本 EPS 反推，同基本口徑；未計可轉債稀釋）",
    gw_label="部署 GW（IT 電力）",
    same_names="部署 GW、每 GW 高速 CCL 金額、AI 高速 CCL 市場三項",
    a1_tail=(f"{VER}（{BUILT}）：個股 HTML 版面規格 v1.3（標準損益表、反向 DCF 頁、四個目標價、加權目標價敏感度主圖）＋三層協調（Coord 頁兩讀數目標價、讀取層級、C0（實績））；"
             "A1 標題版號由舊的 v0.2.1 更正；既有數值不變"),
    ver_line=(f"{VER}（{BUILT}）：改為個股 HTML 版面規格 v1.3（樣板 NVIDIA、台達電、台光電）並補三層協調：Coord 頁並列兩讀數目標價（tools/coord）、讀取層級（同源，不判定）、"
              "C0（實績）（月營收 2026-01～10 對 Siliconomics AI 高速 CCL 市場 × 台燿份額；結論：無可比實績，份額由本模型錨點校準）；新增 RevDCF、Hist 頁、PL 標準損益表、Val 0.7、Sens 末表；"
              "既有數值、目標價與評等不變"),
)


def layer_rows(ver, CR):
    return [
        ("部署 GW CY2026／27（三情境相同；SIF_DeployGW_<情境>）", '=TEXT(gw_26,"0.0")&"／"&TEXT(gw_27,"0.0")', ver, "AI 高速 CCL 市場的部署量（In 頁 A 節直接讀入）", None),
        ("部署 GW CY2028／29：情境 A／B／C", '=TEXT(gw28_A,"0.0")&"／"&TEXT(gw28_B,"0.0")&"／"&TEXT(gw28_C,"0.0")&"；"&TEXT(gw29_A,"0.0")&"／"&TEXT(gw29_B,"0.0")&"／"&TEXT(gw29_C,"0.0")', ver, "三情境部署量（B＝Siliconomics 基準；A、C 依市場層情境係數）", None),
        ("每 GW 高速 CCL 金額 CY2026–29（US$M/GW；SIF_CCLperGW）", '=TEXT(pgw_26,"0")&"／"&TEXT(pgw_27,"0")&"／"&TEXT(pgw_28,"0")&"／"&TEXT(pgw_29,"0")', ver, "AI 高速 CCL 市場＝部署 GW × 每 GW 金額", None),
        ("匯率（NT$/US$；fx_twd）", "=fx", ver, "AI 高速 CCL 市場換算新台幣", "0.00"),
        ("台燿份額、售價指數、產能、組合指數（不在 Siliconomics）", "本模型", "個股模型（Siliconomics M16 註明：價格指數與各公司份額留在公司模型）", "FY2026 有效份額＝高速材料營收錨點 ÷ AI 高速 CCL 市場（校準）", None),
    ]


def same_blocks(SIL, pc):
    N = SIL.NAMES
    return [
        ("gw", "部署 GW（IT 電力）", "GW", "部署 GW（IT 電力）", list(N["SIF_DeployGW_Base"]),
         "讀自 Siliconomics（In 頁 gw_26、gw_27、gw28_B、gw29_B）", f"Derived：Siliconomics {SIL.VERSION} SIF_DeployGW_Base（依規則 v1.4 的基準）", "0.00"),
        ("pgw", "每 GW 高速 CCL 金額", "US$M/GW", "每 GW 高速 CCL 金額（不含價格指數）", list(N["SIF_CCLperGW"]), "讀自 Siliconomics（In 頁 pgw_26–29）", f"Derived：Siliconomics {SIL.VERSION} SIF_CCLperGW", "#,##0"),
        ("mkt", "AI 高速 CCL 市場（US$B）", "US$B", "AI 高速 CCL 市場（美元；不含價格指數）", list(N["SIF_CCLMkt_Base"]), "Derived：部署 GW × 每 GW 高速 CCL 金額（× mf，基準 1）", f"Derived：Siliconomics {SIL.VERSION} SIF_CCLMkt_Base", "0.00"),
    ]


def lines(wb):
    P = wb["PL"]; SE = wb["Sens"]
    rb = blockB(P)
    pb = lambda label, c: f"PL!{c}{lab(P, label, rb, exact=True, end=rb + 70)}"
    sh = lambda k: f"Sens!H{lab(SE, k)}"
    return dict(
        hero=lambda tp, up, rec, dcf, cadj: (T("加權目標價 "), f'TEXT({tp},"#,##0")', T(" 元，較現價"), f'IF({up}<0,"低 ","高 ")', f'TEXT(ABS({up}),"0.0%")', T("，評等"), rec,
                                             T("。現金流折現 "), f'TEXT({dcf},"#,##0")', T(" 元、同業比較法（相對倍數調整後）"), f'TEXT({cadj},"#,##0")',
                                             T(" 元各半；高速材料營收受產能上限與份額約束，目標價對售價、AI 高速 CCL 市場與擴產最敏感。")),
        premise="產業部署量與每 GW 高速 CCL 金額依 Siliconomics 基準（行業變數）、台燿份額相對值、售價指數與擴產計畫（假設）成立",
        o2=(T("高速材料是成長主力：情境 B 高速材料營收由 FY2026 約 "), f'TEXT({pb(AIR, "D")}*10,"#,##0")', T(" 億元增至 FY2028 約 "), f'TEXT({pb(AIR, "F")}*10,"#,##0")',
            T(" 億元，占營收由 "), f'TEXT({pb("高速材料 ÷ 營收", "D")},"0%")', T(" 升到 "), f'TEXT({pb("高速材料 ÷ 營收", "F")},"0%")', T("；同期台燿有效份額由 "),
            f'TEXT({pb("台燿有效份額", "D")},"0.0%")', T(" 升到 "), f'TEXT({pb("台燿有效份額", "F")},"0.0%")', T("（泰國、常熟擴產）。")),
        o3=(T("最大的變數是售價與組合（情境 A 為底，FY2028 EPS）：2028 年售價指數取低值 "), f'TEXT({sh("台燿 CCL 年均售價指數 2028")},"+0.0%;-0.0%")', T("，2028 年組合指數取低值 "),
            f'TEXT({sh("每張營收組合指數 2028")},"+0.0%;-0.0%")', T("，AI 高速 CCL 市場 ×0.8 為 "), f'TEXT({sh("AI 高速 CCL 市場係數")},"+0.0%;-0.0%")', T("；基準情境 B 的 FY2028 EPS 為 "),
            f'TEXT({pb("EPS", "F")},"0.0")', T(" 元。")),
        o5_tail="本模型的部署 GW、每 GW 高速 CCL 金額與市場同 Siliconomics 基準；C0（實績）無可比實績（公司未揭露高速材料營收金額，份額由本模型錨點校準）。",
        drv=(T("分部營收推導（情境 B）：高速材料營收＝MIN（部署 GW × 每 GW 高速 CCL 金額 × 台燿有效份額 × 售價指數，產能上限 − 其他等級）；其他等級依年增率；"
               "相加為合併營收。FY2028 高速材料 "), f'TEXT({pb(AIR, "F")},"#,##0.0")', T("、其他等級 "), f'TEXT({pb(NAIR, "F")},"#,##0.0")',
             T("，合計 "), f'TEXT({pb(REVR, "F")},"#,##0.0")', T(" 十億元。")),
    )


def c0_actuals(wb, G):
    CW = wb["Coord"]; Q = wb["Q"]; P = wb["PL"]
    rb = blockB(P)
    sh0 = lab(P, "台燿有效份額（FY2026 校準）", rb)
    S = Sheet(CW, G["end"] + 3)
    R = {}
    R["title"] = S.bold("4. C0（實績）：公司實績 vs Siliconomics 層級隱含值（規則 v1.4.2 C0 第 2 點；工作單 2026-11-14；差距＝實績 ÷ 層級隱含值 − 1）")
    R["method"] = S.text("方法：個股側＝台燿已公布實績（月營收、3Q26 季報）；層級側＝Siliconomics 現行版 AI 高速 CCL 市場（CY2026）× 台燿 FY2026 有效份額 × 匯率（市場連到本頁第 2 節層級值列）。"
                         "層級值是日曆年，故主讀數＝2026-01～10 累計實績 vs 層級 CY2026 × 10 ÷ 12（線性時差近似，Derived）；第二讀數＝3Q26 × 4 vs 層級 CY2026（約等於 3Q26 中點 8 月中的年化速率，"
                         "季中點晚於年中、成長年偏高）。判定：兩讀數差距都超過 ±15% 且同號 → 對帳中；只一個超過 → 時差邊界（參考）；都在內 → 一致。公司沒有揭露高速材料（ELL／SLL）營收金額"
                         "（法說只給占營收比例 1Q26 38%、2Q26 39%，3Q26 未揭露），故高速材料列無可比實績；以占比換算者只列參考。")
    R["limit"] = S.text("限制（同源校準）：Siliconomics 沒有台燿份額（M16 註明份額留在公司模型）；本模型 FY2026 有效份額＝高速材料營收錨點（高速材料占比 42% × FY2026 營收錨點）÷ AI 高速 CCL 市場，"
                        "而本模型的市場就是 Siliconomics 的市場（部署 GW 與每 GW 金額皆讀自 Siliconomics）。層級隱含高速材料營收＝市場 × 份額 × 匯率＝高速材料營收錨點，市場大小在比式中被約掉；"
                        "以 42% × 實績營收對照時，兩讀數代數上等於「實績營收 ÷ 營收錨點 − 1」，實際只檢驗本模型營收錨點，且 42% 為假設（1H26 公司占比 38–39% 往上推），不是公司揭露金額。"
                        "結論：無可比實績（層級值同源）；缺台燿高速材料營收金額揭露。")
    S.skip()
    S.bold("4a. 公司已公布實績與揭露（輸入；藍字＝來自原文的常數）")
    S.hdr(["項目", "值", "", "", "", "單位", "口徑、來源、證據等級", "來源網址"])
    R["sep"] = S.row("月營收：2026-09", 6.126, None, "NT$B", "Verified：月營收公告 61.26 億元，月增 3.6%、年增 114.8%，累計前 9 月 423.04 億元（TechNews 2026-10-05）", blue=True, url=M_SEP)
    R["oct"] = S.row("月營收：2026-10", 6.478, None, "NT$B", "Verified：月營收公告 64.78 億元，月增 5.8%、年增 127.0%（TechNews 2026-11-05）", blue=True, url=M_OCT)
    R["cum9"] = S.row("累計：2026-01～09", 42.304, None, "NT$B", "Verified：月營收公告 423.04 億元，年增 99.4%（TechNews 2026-10-05）", blue=True, url=M_SEP)
    R["cum10"] = S.row("累計：2026-01～10（主讀數的實績基礎）", 48.782, None, "NT$B", "Verified：月營收公告 487.82 億元，年增 102.7%（TechNews 2026-11-05）", blue=True, url=M_OCT)
    R["chk1"] = S.row("檢查：累計 9 月 + 10 月 − 累計 10 月（應為 0）", f"=B{R['cum9']}+B{R['oct']}-B{R['cum10']}", None, "NT$B", "Derived")
    R["q3"] = S.row("3Q26 營收（最近一季；第二讀數的實績基礎）", f"=Q!D{lab(Q, '營收')}", None, "NT$B", "Verified：Q 頁 3Q26 欄（季報 179.49 億元）")
    R["chk2"] = S.row("檢查：累計 9 月 − 3Q26 − Q 頁 1H26 營收（應接近 0；季報四捨五入）", f"=B{R['cum9']}-B{R['q3']}-Q!C{lab(Q, '營收')}", None, "NT$B", "Derived：1H26 243.55 億元（Q 頁，公司財報）")
    R["ai"] = S.row("公司揭露：高速材料（ELL／SLL）營收金額", "未揭露", None, "", "公司未揭露高速材料營收金額；法說只給占比（4Q25 36%、1Q26 38%、2Q26 39%；富果轉述，Interested-party）；In 頁 ai_26 42% 為假設", fmt="General")
    R["months"] = S.row("已公布月數（2026-01～10）", 10, None, "月", "輸入：月營收至 2026-10（2026-11-05 公告）", fmt="0", blue=True)
    S.skip()
    S.bold("4b. 層級隱含值（Siliconomics 現行版；市場連到本頁第 2 節層級值列）")
    S.hdr(["項目", "CY2026", "", "", "", "單位", "口徑、來源、證據等級"])
    R["lmkt"] = S.row("AI 高速 CCL 市場（層級值）", f"=B{G['mkt'][1]}", None, "US$B", "Derived：本頁第 2 節 Siliconomics SIF_CCLMkt_Base（CY2026）")
    R["lsh"] = S.row("台燿 FY2026 有效份額（本模型校準；Siliconomics 無此量）", f"=PL!D{sh0}", None, "%", "Derived（本模型校準）：高速材料營收錨點 ÷ AI 高速 CCL 市場（PL 頁情境 B）", fmt="0.0%")
    R["lai"] = S.row("層級隱含：台燿高速材料營收（市場 × 份額 × 匯率）", f"=B{R['lmkt']}*B{R['lsh']}*fx", None, "NT$B", "Derived：匯率 In 頁 fx", fmt="#,##0.0")
    S.skip()
    R["tbl"] = S.bold("C0（實績）：公司實績 vs 層級隱含值（HTML 用）")
    R["hdr"] = S.hdr(["項目", "主讀數：2026-01～10 累計 vs 層級 CY2026 × 10 ÷ 12", "第二讀數：3Q26 × 4 vs 層級 CY2026", "", "", "單位", "口徑、來源、證據等級"])
    fr = f"B{R['months']}/12"
    pct = "+0.0%;-0.0%"
    S.row("高速材料：公司實績", "無（未揭露金額）", "無（未揭露金額）", "NT$B", "公司未揭露高速材料營收金額；不以模型值或占比換算補位", fmt="General")
    l = S.row("高速材料：層級隱含", f"=B{R['lai']}*{fr}", f"=B{R['lai']}", "NT$B", "Derived：層級隱含高速材料營收（4b）；主讀數 × 10 ÷ 12", fmt="#,##0.0")
    R["j_ai"] = S.row("高速材料：判定", "無可比實績（層級值同源）", "無可比實績（層級值同源）", "", "缺高速材料營收揭露；份額由本模型校準，市場大小未受檢驗；不計入對帳中", fmt="General")
    a2 = S.row("高速材料（參考，同源校準）：占比 42% × 公司營收實績", f"=B{R['cum10']}*ai_26", f"=B{R['q3']}*ai_26*4", "NT$B", "Verified 營收 × 假設占比（In 頁 ai_26）；不是公司揭露金額", fmt="#,##0.0")
    g2 = S.row("高速材料（參考，同源校準）：差距（÷ 層級隱含 − 1）", f"=B{a2}/B{l}-1", f"=C{a2}/C{l}-1", fmt=pct)
    R["j_ref"] = S.row("高速材料（參考，同源校準）：判定", "參考（同源校準）：不判定", "參考（同源校準）：不判定", "", "兩讀數代數上等於實績營收 ÷ 營收錨點 − 1（占比與市場大小約掉），只檢驗本模型營收錨點", fmt="General")
    R["ub"] = S.row("上限檢查：層級隱含高速材料 ÷ 公司營收實績（主讀數期間；應 ≤ 100%）", f"=B{l}/B{R['cum10']}", None, "%", "Derived：層級高速材料不得超過公司全部營收；通過＝合理，不是一致性判定", fmt="0.0%")
    R["j_ub"] = S.row("上限檢查：結果", f'=IF(B{R["ub"]}<=1,"通過","不通過")', None, "", "", fmt="General")
    S.skip()
    S.bold("參考：本模型 FY2026 營收錨點對月營收實績（模型對實績，不是層級對帳；不計入 C0）")
    R["anc"] = S.row("FY2026 營收錨點（Q 頁）", f"=Q!F{lab(Q, '營收')}", None, "NT$B", "Derived／Assumed：1H26、3Q26 實績＋4Q26 假設季增 5%", fmt="#,##0.0")
    R["imp"] = S.row("錨點隱含 11–12 月月均營收（錨點 − 1～10 月累計）÷ 2", f"=(B{R['anc']}-B{R['cum10']})/2", None, "NT$B", "Derived", fmt="#,##0.00")
    R["impg"] = S.row("隱含月均 ÷ 10 月實績 − 1", f"=B{R['imp']}/B{R['oct']}-1", None, "%", "Derived：正值＝錨點假設 11–12 月高於 10 月水準", fmt=pct)
    S.skip()
    R["sum0"] = S.bold("C0（實績）結果摘要（HTML 用；公式帶數字）")
    nopen = f'COUNTIF(B{R["j_ai"]}:C{R["j_ref"]},"對帳中")'
    R["sum1"] = S.text(F(T("C0（實績）結論：無可比實績（層級值同源）：台燿未揭露高速材料營收金額；Siliconomics 沒有台燿份額，本模型份額以營收錨點對 Siliconomics 市場校準。對帳中項目 "), nopen,
                         T(" 項。參考：以假設占比換算，差距主讀數 "), f'TEXT(B{g2},"+0.0%;-0.0%")', T("、第二讀數 "), f'TEXT(C{g2},"+0.0%;-0.0%")',
                         T("，只反映營收錨點與實績的差異，不判定。")))
    R["sum2"] = S.text(F(T("上限檢查：層級隱含高速材料（主讀數期間）占公司營收實績 "), f'TEXT(B{R["ub"]},"0%")', T("，"), f'B{R["j_ub"]}',
                         T("；10 月營收 "), f'TEXT(B{R["oct"]}*10,"#,##0.0")', T(" 億元，營收錨點隱含 11–12 月月均 "), f'TEXT(B{R["imp"]}*10,"#,##0.0")', T(" 億元（較 10 月 "),
                         f'TEXT(B{R["impg"]},"+0%;-0%")', T("）。")))
    R["sum3"] = S.text("限制與待查：公司未揭露高速材料營收，Siliconomics 的 AI 高速 CCL 市場也沒有獨立的台燿份額，故市場大小未獲獨立驗證；4Q26 季報公布後於下一版更新第二讀數。")
    R["end"] = S.r - 1
    return R


def apply(wb, SIL, CR):
    C.update(layer_rows=layer_rows, same_blocks=same_blocks)
    return c13.apply(wb, SIL, CR, C, lines, c0_actuals)
