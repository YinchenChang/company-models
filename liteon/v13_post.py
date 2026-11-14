# 光寶科 v0.4.5（2026-11-15）後處理：個股 HTML 版面規格 v1.3＋三層協調兩讀數＋C0（實績）（工作單 docs/workorders/20261114_第二批13家_v1.3與三層協調.md；
# 規則 docs/plan/三層基準協調規則.md v1.4.2；樣板 台達電 delta/v056_post.py、v057_post.py，同次產業奇鋐、雙鴻）。共通步驟在 liteon/c13.py（電源散熱四家同一份），本檔只放光寶科的設定。
# build_liteon.py 在 Coord 頁建好之後、存檔前呼叫 apply(wb, SIL, coord_rows)。alt_run.py 建置（無 Coord 頁）不呼叫。只新增，不改既有數值格。
import os, sys
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import c13
from c13 import T, F, lab, blockB, Sheet

VER, BUILT = "v0.4.5", "2026-11-15"
SA = "https://stockanalysis.com/quote/tpe/2301/financials/income-statement/"
M_OCT = "https://finance.technews.tw/2026/11/09/ltc-2301-202610-financial-report/"
M_SEP = "https://finance.technews.tw/2026/10/08/ltc-2301-202609-financial-report/"
Q3 = "https://www.stockfeel.com.tw/?p=278632"

C = dict(
    name="光寶科", ver=VER, built=BUILT, nop_label="營業外淨額",
    ext_rows=lambda P: (lambda b: dict(rev=lab(P, "營收", b, exact=True, end=b + 16), ni=lab(P, "歸屬母公司淨利", b, exact=True, end=b + 16), sh=None,
                                         eps=lab(P, "EPS", b, exact=True, end=b + 16), hdr=3))(lab(P, "情境 B", exact=True)),   # 延伸期在 PL 頁上方摘要表 H–J 欄；股數＝淨利 ÷ EPS
    hist_title=("Hist — 損益表歷史（v0.4.5 起；PL 頁「標準損益表」的歷史三年與年增率用）。金額：十億新台幣；EPS：新台幣；歸屬母公司。"
                f"來源：stockanalysis 年度損益表（百萬元，S&P Global，2026-11-15 讀取；{SA}），FY2025 與 Q 頁一致。營業外收益＝稅前 − 營業利益；營業費用＝毛利 − 營業利益"),
    hist=[
        ("rev", "營收", 173.456, 148.333, 137.134, 166.085, "Verified", f"{SA}：Revenue（百萬元）；FY2025 同 Q 頁"),
        ("gp", "毛利", 33.262, 32.668, 29.630, 38.039, "Verified", "同上：Gross Profit"),
        ("oi", "營業利益", 15.073, 14.516, 12.934, 16.738, "Verified", "同上：Operating Income"),
        ("pti", "稅前利益", 17.846, 18.366, 15.635, 19.713, "Verified", "同上：Pretax Income（Q 頁 FY2025 19.714 為推導值）"),
        ("tax", "所得稅", 3.658, 3.767, 3.674, 4.601, "Verified", "同上：Income Tax Expense"),
        ("nci", "非控制權益損益", 0.036, 0.028, 0.019, -0.004, "Verified", "同上：Minority Interest in Earnings（FY2025 為少數股權虧損）"),
        ("ni", "稅後淨利（歸屬母公司）", 14.151, 14.571, 11.942, 15.116, "Verified", "同上：Net Income to Common；FY2025 同 Q 頁"),
        ("sh", "EPS 換算股數（十億股）", 2.285, 2.289, 2.292, 2.275, "Verified", "同上：Shares Outstanding (Basic)"),
        ("eps", "EPS（公司公告，基本，NT$）", 6.19, 6.36, 5.21, 6.64, "Verified", "同上：EPS (Basic)（稀釋 6.10／6.29／5.15／6.59）；FY2025 同 Q 頁 6.64"),
    ],
    fy26_short="FY2026E 含 1H26 與 3Q26 實績",
    fy26_long="1H26 財報＋3Q26 實績（2026-10-30 法說：營收 574.03 億元、歸屬母公司淨利 74.29 億元、EPS 3.27 元）＋4Q26 假設",
    eps_note="歷史為公司公告基本 EPS；預測為歸屬母公司淨利 ÷ 股數（同基本口徑）",
    gw_label="部署 GW（IT 電力）",
    same_names="部署 GW、每 kW 電源內容、AI 電源市場與光寶有效份額四項",
    resid=("shsum", "Flex、麥格米特、Vertiv 等；不得歸零（C0 第 3 點）"),
    a1_tail=(f"{VER}（{BUILT}）：個股 HTML 版面規格 v1.3（標準損益表、反向 DCF 頁、四個目標價、加權目標價敏感度主圖）＋三層協調（Coord 頁兩讀數目標價、讀取層級、C0（實績））；"
             "A1 標題版號由舊的 v0.2 更正；既有數值不變"),
    ver_line=(f"{VER}（{BUILT}）：改為個股 HTML 版面規格 v1.3（樣板 NVIDIA、台達電）並補三層協調：Coord 頁並列兩讀數目標價（tools/coord）、讀取層級（同源，不判定）、"
              "C0（實績）（月營收 2026-01～10、3Q26 法說部門營收對 Siliconomics AI 電源市場 × 光寶份額；結論：無可比實績，層級份額由本公司校準）；新增 RevDCF、Hist 頁、PL 標準損益表、Val 0.7、Sens 末表；"
              "既有數值、目標價與評等不變"),
)


def layer_rows(ver, CR):
    return [
        ("每 kW AI 電源內容 CY2026／27／28／29（US$/kW；pwr_pkw_26–29）", '=TEXT(pkw_26,"0")&"／"&TEXT(pkw_27,"0")&"／"&TEXT(pkw_28,"0")&"／"&TEXT(pkw_29,"0")', ver, "AI 電源市場＝部署 GW × 每 kW 內容", None),
        ("光寶份額相對值 CY2027／28／29（÷ CY2026；pwr_lrel_27–29）", '=TEXT(rel_27,"0.00")&"／"&TEXT(rel_28,"0.00")&"／"&TEXT(rel_29,"0.00")', ver, "有效份額路徑；CY2026 有效份額由本模型以公司 AI 相關占比 ×（1 − 非電源占比）× 營收錨點校準後寫入 Siliconomics（pwr_lsh_26）", None),
        ("匯率（NT$/US$；fx_twd）", "=fx", ver, "AI 電源市場換算新台幣", "0.00"),
        ("情境 B 顆數係數 CY2028／CY2029（÷ 原始顆數；SIF_UnitsCoef_Base）", '=TEXT(um28_B,"0.000")&"／"&TEXT(um29_B,"0.000")', f"{ver}，經 tsmc/dist 成品 In 頁 um28_B、um29_B 讀取", "基準情境 B 的 AI IT 資本支出係數；A（原始顆數）、C（0.6）為個股自訂情境", None),
        ("AI IT 資本支出 CY2026–29（US$B；markets/network）", '=TEXT(itc_26,"#,##0")&"／"&TEXT(itc_27,"#,##0")&"／"&TEXT(itc_28,"#,##0")&"／"&TEXT(itc_29,"#,##0")', "markets/network 成品（市場層；Tokenomics 經橋接表進入）", "部署 GW 的分子（原始顆數口徑）", None),
        ("每 GW IT 資本支出（US$B；GB300／VR200；markets/chip_bridge）", '=TEXT(itpg_gb,"0.0")&"／"&TEXT(itpg_vr,"0.0")', "markets/chip_bridge 成品（Tokenomics IF_CapexIT，經橋接表唯讀）", "部署 GW 的分母", None),
    ]


def same_blocks(SIL, pc):
    N, V = SIL.NAMES, SIL.V
    lsh, dsh = V["pwr_lsh_26"]["base"], V["pwr_dsh_26"]["base"]
    lrel = [1.0] + [V[f"pwr_lrel_{y}"]["base"] for y in (27, 28, 29)]
    drel = [1.0] + [V[f"pwr_drel_{y}"]["base"] for y in (27, 28, 29)]
    sh = pc("光寶有效份額")
    return [
        ("gw", "部署 GW（IT 電力）", "GW", "部署 GW（IT 電力）", list(N["SIF_DeployGW_Base"]),
         "Derived：AI IT 資本支出（markets/network）× 情境 B 顆數係數 ÷ 每 GW IT 資本支出（bridge）", f"Derived：Siliconomics {SIL.VERSION} SIF_DeployGW_Base（依規則 v1.4 的基準）", "0.00"),
        ("pkw", "每 kW 電源內容", "US$/kW", "AI 電源內容（台達電模型）", [V[f"pwr_pkw_{y}"]["base"] for y in (26, 27, 28, 29)],
         "讀自 Siliconomics（In 頁 pkw_26–29）", f"Interested-party（2026–27：凱基推導）／Assumed（2028 起）：Siliconomics {SIL.VERSION} pwr_pkw_26–29", "#,##0"),
        ("mkt", "AI 電源市場（US$B）", "US$B", "AI 電源市場（美元）", list(N["SIF_PwrMkt_Base"]), "Derived：部署 GW × 每 kW 電源內容", f"Derived：Siliconomics {SIL.VERSION} SIF_PwrMkt_Base", "0.00"),
        ("sh", "光寶有效份額", "%", "光寶有效份額", [f"={lsh}*{x}" for x in lrel],
         "Derived：FY2026 以公司 AI 相關占比 ×（1 − 非電源占比）× 營收錨點 ÷ 市場校準，之後 × 份額相對值（Assumed）", f"Derived（本公司校準值，Siliconomics 快照）：pwr_lsh_26 × pwr_lrel_27–29（CY2026 快照值 {lsh:.4f}）", "0.0%"),
        ("shsum", "已建模個股份額加總（光寶本模型＋台達電層級值）", "%", [f"{s}+{dsh}*{x}" for s, x in zip(sh, drel)], list(N["SIF_PwrShareSum_Base"]),
         "Derived：光寶取本模型；台達電取 Siliconomics pwr_dsh_26 × pwr_drel", f"Derived：Siliconomics {SIL.VERSION} SIF_PwrShareSum_Base", "0.0%"),
    ]


def lines(wb):
    P = wb["PL"]; SE = wb["Sens"]
    rb = blockB(P)
    pb = lambda label, c: f"PL!{c}{lab(P, label, rb, exact=True, end=rb + 70)}"
    sl = lambda k: f"Sens!L{lab(SE, k)}"
    return dict(
        hero=lambda tp, up, rec, dcf, cadj: (T("加權目標價 "), f'TEXT({tp},"#,##0.0")', T(" 元，較現價"), f'IF({up}<0,"低 ","高 ")', f'TEXT(ABS({up}),"0.0%")', T("，評等"), rec,
                                             T("。現金流折現 "), f'TEXT({dcf},"#,##0.0")', T(" 元、同業比較法（相對倍數調整後）"), f'TEXT({cadj},"#,##0.0")',
                                             T(" 元各半；AI 電源帶動雲端及物聯網成長，目標價對 AI 部署量與份額最敏感。")),
        premise="產業部署量與每 kW 電源內容依 Siliconomics 基準、光寶份額相對值（假設）成立",
        o2=(T("AI 電源是成長主力：情境 B AI 相關營收由 FY2026 約 "), f'TEXT({pb("AI 相關營收（＝AI 電源 ÷（1 − 非電源占比））", "D")}*10,"#,##0")', T(" 億元增至 FY2028 約 "),
            f'TEXT({pb("AI 相關營收（＝AI 電源 ÷（1 − 非電源占比））", "F")}*10,"#,##0")', T(" 億元，占營收由 "), f'TEXT({pb("AI 相關 ÷ 營收", "D")},"0%")', T(" 升到 "),
            f'TEXT({pb("AI 相關 ÷ 營收", "F")},"0%")', T("；光寶在 AI 電源市場的有效份額維持 "), f'TEXT({pb("光寶有效份額", "F")},"0.0%")', T("。")),
        o3=(T("最大的變數是 AI 部署量與份額：AI 部署 ×0.8 使情境 B 的 FY2028 EPS "), f'TEXT({sl("AI 部署 ×0.8")},"+0.0%;-0.0%")', T("，份額流失 "),
            f'TEXT({sl("份額流失（低）")},"+0.0%;-0.0%")', T("，電源內容取低值 "), f'TEXT({sl("電源內容 低")},"+0.0%;-0.0%")', T("；基準情境 B 的 FY2028 EPS 為 "),
            f'TEXT({pb("EPS", "F")},"0.00")', T(" 元。")),
        o5_tail="本模型的部署 GW 與 AI 電源市場同 Siliconomics 基準；C0（實績）無可比實績（公司未揭露 AI 電源金額，層級份額由本公司校準）。",
        drv=(T("分部營收推導（情境 B）：AI 電源營收＝部署 GW × 每 kW 電源內容 × 光寶有效份額；AI 相關＝AI 電源 ÷（1 − 非電源占比）；雲端及物聯網＝AI 相關＋非 AI；光電、資訊及消費性電子各依年增率；"
               "三大事業群相加為合併營收。FY2028 AI 電源 "), f'TEXT({pb("AI 電源營收", "F")},"#,##0")', T("、雲端及物聯網 "), f'TEXT({pb("雲端及物聯網", "F")},"#,##0")',
             T("、光電 "), f'TEXT({pb("光電", "F")},"#,##0")', T("、資訊及消費性電子 "), f'TEXT({pb("資訊及消費性電子", "F")},"#,##0")',
             T("，合計 "), f'TEXT({pb("營收", "F")},"#,##0")', T(" 十億元。")),
    )


def c0_actuals(wb, G):
    CW = wb["Coord"]; Q = wb["Q"]
    S = Sheet(CW, G["end"] + 3)
    R = {}
    R["title"] = S.bold("4. C0（實績）：公司實績 vs Siliconomics 層級隱含值（規則 v1.4.2 C0 第 2 點；工作單 2026-11-14；差距＝實績 ÷ 層級隱含值 − 1）")
    R["method"] = S.text("方法：個股側＝光寶科已公布實績（月營收、3Q26 財報與法說部門營收）；層級側＝Siliconomics 現行版 AI 電源市場（CY2026）× 光寶有效份額 × 匯率（連到本頁第 2 節層級值列）。"
                         "層級值是日曆年，故主讀數＝2026-01～10 累計實績 vs 層級 CY2026 × 10 ÷ 12（線性時差近似，Derived）；第二讀數＝3Q26 × 4 vs 層級 CY2026（約等於 3Q26 中點 8 月中的年化速率，"
                         "季中點晚於年中、成長年偏高）。判定：兩讀數差距都超過 ±15% 且同號 → 對帳中；只一個超過 → 時差邊界（參考）；都在內 → 一致。公司沒有揭露 AI 電源營收金額，"
                         "只口述「2026 年 AI 相關營收占比超過 30%」（含 BBU 以外的非電源 AI 產品），口述占比 × 營收屬 Assumed，只列參考；雲端及物聯網部門營收為 Verified，只作上限檢查。")
    R["limit"] = S.text("限制（同源校準）：Siliconomics 的光寶 CY2026 有效份額是本模型以「公司 AI 相關占比 31% ×（1 − 非電源占比 10%）× FY2026 營收錨點 ÷ 本庫 AI 電源市場」校準後寫入，"
                        "層級隱含 AI 電源營收＝31% × 90% × 營收錨點，Siliconomics 市場大小（部署 GW × 每 kW 內容）在比式中被約掉；以同一占比 × 實績營收對照時，兩讀數代數上等於「實績營收 ÷ 營收錨點 − 1」，"
                        "實際只檢驗本模型營收錨點。結論：無可比實績（層級值同源）；缺光寶 AI 電源營收金額揭露。公司若揭露 AI 電源金額，才能獨立檢驗市場大小。")
    S.skip()
    S.bold("4a. 公司已公布實績與揭露（輸入；藍字＝來自原文的常數）")
    S.hdr(["項目", "值", "", "", "", "單位", "口徑、來源、證據等級", "來源網址"])
    R["sep"] = S.row("月營收：2026-09", 18.848, None, "NT$B", "Verified：月營收公告 188.48 億元，月增 −3.8%、年增 22.1%，累計前 9 月 1,535.59 億元（TechNews 2026-10-08）", blue=True, url=M_SEP)
    R["oct"] = S.row("月營收：2026-10", 20.071, None, "NT$B", "Verified：月營收公告 200.71 億元，月增 6.5%、年增 32.0%（TechNews 2026-11-09）", blue=True, url=M_OCT)
    R["cum9"] = S.row("累計：2026-01～09", 153.559, None, "NT$B", "Verified：月營收公告 1,535.59 億元，年增 26.1%（TechNews 2026-10-08）", blue=True, url=M_SEP)
    R["cum10"] = S.row("累計：2026-01～10（主讀數的實績基礎）", 173.630, None, "NT$B", "Verified：月營收公告 1,736.30 億元，年增 26.8%（TechNews 2026-11-09）", blue=True, url=M_OCT)
    R["chk1"] = S.row("檢查：累計 9 月 + 10 月 − 累計 10 月（應為 0）", f"=B{R['cum9']}+B{R['oct']}-B{R['cum10']}", None, "NT$B", "Derived")
    R["q3"] = S.row("3Q26 營收（最近一季；第二讀數的實績基礎；財報）", 57.403, None, "NT$B", "Verified：3Q26 法說（2026-10-30）營收 574.03 億元，季增 9%、年增 28%；毛利率 25.9%、營業利益 86.55 億元、歸屬母公司淨利 74.29 億元、EPS 3.27 元（股感）", blue=True, url=Q3)
    R["chk2"] = S.row("檢查：3Q26 財報營收 − In 頁 q3_rev（月營收 1–9 月累計 − 1H26；財報與月營收口徑差）", f"=B{R['q3']}-q3_rev", None, "NT$B", "Derived：月營收合計 574.49 億元，財報 574.03 億元")
    R["cloud"] = S.row("部門營收：雲端及物聯網，3Q26", 33.753, None, "NT$B", "Verified：3Q26 法說部門營收 337.53 億元、占 59%、年增 57%；營業利益 71.13 億元（股感）", blue=True, url=Q3)
    R["ash"] = S.row("公司口述：AI 相關營收占比（2026 年）", 0.30, None, "%", "Verified（公司口述、下限「超過 30%」；期間為全年）：3Q26 法說「2026 年 AI 相關營收佔比超過 30%，明年預估提升到 40% 至 50%」（股感）；In 頁 ai_26 取 31%", fmt="0.0%", blue=True, url=Q3)
    R["months"] = S.row("已公布月數（2026-01～10）", 10, None, "月", "輸入：月營收至 2026-10（2026-11-09 公告）", fmt="0", blue=True)
    S.skip()
    S.bold("4b. 層級隱含值（Siliconomics 現行版；連到本頁第 2 節層級值列）")
    S.hdr(["項目", "CY2026", "", "", "", "單位", "口徑、來源、證據等級"])
    R["lmkt"] = S.row("AI 電源市場（層級值）", f"=B{G['mkt'][1]}", None, "US$B", "Derived：本頁第 2 節 Siliconomics SIF_PwrMkt_Base（CY2026）")
    R["lsh"] = S.row("光寶有效份額（層級值）", f"=B{G['sh'][1]}", None, "%", "Derived（本公司校準值，Siliconomics 快照）：pwr_lsh_26，由公司 AI 相關占比 × 營收錨點校準", fmt="0.0%")
    R["lpwr"] = S.row("層級隱含：光寶 AI 電源營收（市場 × 份額 × 匯率）", f"=B{R['lmkt']}*B{R['lsh']}*fx", None, "NT$B", "Derived：匯率 In 頁 fx", fmt="#,##0.0")
    S.skip()
    R["tbl"] = S.bold("C0（實績）：公司實績 vs 層級隱含值（HTML 用）")
    R["hdr"] = S.hdr(["項目", "主讀數：2026-01～10 累計 vs 層級 CY2026 × 10 ÷ 12", "第二讀數：3Q26 × 4 vs 層級 CY2026", "", "", "單位", "口徑、來源、證據等級"])
    fr = f"B{R['months']}/12"
    pct = "+0.0%;-0.0%"
    S.row("AI 電源：公司實績", "無（未揭露 AI 電源營收）", "無（未揭露 AI 電源營收）", "NT$B", "公司未揭露 AI 電源營收金額；不以模型值或口述占比補位", fmt="General")
    l = S.row("AI 電源：層級隱含", f"=B{R['lpwr']}*{fr}", f"=B{R['lpwr']}", "NT$B", "Derived：層級隱含 AI 電源營收（4b）；主讀數 × 10 ÷ 12", fmt="#,##0.0")
    R["j_pwr"] = S.row("AI 電源：判定", "無可比實績（層級值同源）", "無可比實績（層級值同源）", "", "缺 AI 電源營收揭露；層級份額由本公司校準，市場大小未受檢驗；不計入對帳中", fmt="General")
    a2 = S.row("AI 電源（參考，同源校準）：AI 相關占比 ×（1 − 非電源占比）× 公司營收實績", f"=B{R['cum10']}*ai_26*(1-x_np)", f"=B{R['q3']}*ai_26*(1-x_np)*4", "NT$B",
               "Verified 營收 × 公司口述占比 31%（In 頁 ai_26）×（1 − 非電源占比，In 頁 x_np，Assumed）；口述占比 × 營收屬 Assumed", fmt="#,##0.0")
    g2 = S.row("AI 電源（參考，同源校準）：差距（÷ 層級隱含 − 1）", f"=B{a2}/B{l}-1", f"=C{a2}/C{l}-1", fmt=pct)
    R["j_ref"] = S.row("AI 電源（參考，同源校準）：判定", "參考（同源校準）：不判定", "參考（同源校準）：不判定", "", "兩讀數代數上等於實績營收 ÷ 營收錨點 − 1（占比與市場大小約掉），只檢驗本模型營收錨點", fmt="General")
    s30 = S.row("AI 電源敏感度（參考）：AI 相關占比取口述下限 30% 時的差距", f"=B{R['cum10']}*B{R['ash']}*(1-x_np)/B{l}-1", f"=B{R['q3']}*B{R['ash']}*(1-x_np)*4/C{l}-1", "", "只作敏感度，不判定", fmt=pct)
    R["ub"] = S.row("上限檢查：層級隱含 AI 電源（CY2026 ÷ 4）÷ 雲端及物聯網部門 3Q26 營收（應 ≤ 100%）", None, f"=B{R['lpwr']}/4/B{R['cloud']}", "%", "Derived：AI 電源屬雲端及物聯網部門；通過＝合理，不是一致性判定", fmt="0.0%")
    R["j_ub"] = S.row("上限檢查：結果", None, f'=IF(C{R["ub"]}<=1,"通過","不通過")', "", "", fmt="General")
    S.skip()
    S.bold("參考：本模型 FY2026 營收錨點對月營收實績（模型對實績，不是層級對帳；不計入 C0）")
    R["anc"] = S.row("FY2026 營收錨點（Q 頁）", f"=Q!F{lab(Q, '營收')}", None, "NT$B", "Derived：1H26 實績＋3Q26 實績＋4Q 假設", fmt="#,##0.0")
    R["imp"] = S.row("錨點隱含 11–12 月月均營收（錨點 − 1～10 月累計）÷ 2", f"=(B{R['anc']}-B{R['cum10']})/2", None, "NT$B", "Derived", fmt="#,##0.00")
    R["impg"] = S.row("隱含月均 ÷ 10 月實績 − 1", f"=B{R['imp']}/B{R['oct']}-1", None, "%", "Derived：負值＝錨點假設 11–12 月低於 10 月水準（10 月為近 10 年新高）", fmt=pct)
    S.skip()
    R["sum0"] = S.bold("C0（實績）結果摘要（HTML 用；公式帶數字）")
    nopen = f'COUNTIF(B{R["j_pwr"]}:C{R["j_ref"]},"對帳中")'
    R["sum1"] = S.text(F(T("C0（實績）結論：無可比實績（層級值同源）：光寶科未揭露 AI 電源營收；Siliconomics 的光寶份額由本公司以 AI 相關占比校準。對帳中項目 "), nopen,
                         T(" 項。參考：以 AI 相關占比換算，差距主讀數 "), f'TEXT(B{g2},"+0.0%;-0.0%")', T("、第二讀數 "), f'TEXT(C{g2},"+0.0%;-0.0%")',
                         T("，只反映營收錨點與實績的差異，不判定；占比取口述下限時為 "), f'TEXT(B{s30},"+0.0%;-0.0%")', T("／"), f'TEXT(C{s30},"+0.0%;-0.0%")', T("。")))
    R["sum2"] = S.text(F(T("上限檢查：層級隱含 AI 電源每季約 "), f'TEXT(B{R["lpwr"]}/4,"#,##0.0")', T(" 十億元，占雲端及物聯網部門 3Q26 營收 "), f'TEXT(C{R["ub"]},"0%")', T("，"), f'C{R["j_ub"]}',
                         T("；10 月營收 "), f'TEXT(B{R["oct"]}*10,"#,##0")', T(" 億元，營收錨點隱含 11–12 月月均 "), f'TEXT(B{R["imp"]}*10,"#,##0")', T(" 億元（較 10 月 "),
                         f'TEXT(B{R["impg"]},"+0%;-0%")', T("）。")))
    R["sum3"] = S.text("限制與待查：公司未揭露 AI 電源金額，「AI 相關」含 BBU 以外的非電源 AI 產品且只有年度下限，Siliconomics 的光寶份額由本模型校準，故 AI 電源市場大小未獲獨立驗證（同台達電 v0.5.7 的結論）。")
    R["end"] = S.r - 1
    return R


def apply(wb, SIL, CR):
    C.update(layer_rows=layer_rows, same_blocks=same_blocks)
    return c13.apply(wb, SIL, CR, C, lines, c0_actuals)
