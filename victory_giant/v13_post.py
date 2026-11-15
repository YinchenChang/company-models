# 勝宏科技 v0.4.4（2026-11-15）後處理：個股 HTML 版面規格 v1.3＋三層協調兩讀數＋C0（實績）（工作單 docs/workorders/20261114_第二批13家_v1.3與三層協調.md；
# 規則 docs/plan/三層基準協調規則.md v1.4.2；樣板 NVIDIA v0.3.10、台達電、同次產業台光電 emc、金像電 gce/v13_post.py）。共通步驟在 victory_giant/c13.py（高速 CCL／PCB 四家同一份），本檔只放勝宏的設定。
# 勝宏以人民幣計價：c13 產生的新增格中「NT$」「新台幣」字樣在本檔最後改為「RMB」「人民幣」，標準損益表口徑說明改為中國企業會計準則（只改本版新增格）。
# 推導區塊直接延伸到 FY2032（無「延伸｜情境 B」表）：ext_rows 指向情境 B 區塊本身的營收、淨利、股數、EPS 列（H、I 欄＝FY2030、FY2031）。
# build_victory_giant.py 在 Coord 頁建好之後、存檔前呼叫 apply(wb, SIL, coord_rows)。alt_run.py 建置（無 Coord 頁）不呼叫。只新增，不改既有數值格。
import os, sys
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import c13
from c13 import T, F, lab, blockB, Sheet

VER, BUILT = "v0.4.4", "2026-11-15"
SA = "https://stockanalysis.com/quote/she/300476/financials/income-statement/"
Q3 = "https://www.cs.com.cn/ssgs/01/2026/10/30/detail_2026103010048485.html"
AIR = "AI 相關營收 ＝ MIN（需求面、產能上限 − 非 AI）"


def ext_rows(P):
    rb = blockB(P)
    g = lambda label: lab(P, label, rb, exact=True, end=rb + 70)
    return dict(rev=g("營收"), ni=g("歸屬母公司淨利"), sh=g("股數"), eps=g("EPS"), hdr=rb + 1)


C = dict(
    name="勝宏", ver=VER, built=BUILT, nop_label="營業外與財務淨額", ext_rows=ext_rows,
    hist_title=("Hist — 損益表歷史（v0.4.4 起；PL 頁「標準損益表」的歷史三年與年增率用）。金額：十億人民幣；EPS：人民幣；歸屬母公司。"
                f"來源：stockanalysis 年度損益表（百萬元人民幣，S&P Global，2026-11-15 讀取；{SA}；營收含其他業務收入，即公司營業總收入），FY2025 營收、營業利益、淨利與 Q 頁一致。"
                "營業外收益＝稅前 − 營業利益；營業費用＝毛利 − 營業利益；stockanalysis 淨利與歸屬母公司淨利相同，非控制權益損益列 0"),
    hist_chk2="FY2025：本表營收、毛利、營業利益、淨利 − Q 頁錨點（Q 頁毛利＝營收 × 年報毛利率 35.22%，stockanalysis 毛利以營業成本口徑計，差 0.04 十億元）",
    hist=[
        ("rev", "營收", 7.885, 7.931, 10.731, 19.292, "Verified", f"{SA}：Revenue（百萬元；含其他業務收入）；FY2025 同 Q 頁 192.92 億元"),
        ("gp", "毛利", 1.400, 1.623, 2.420, 6.757, "Verified", "同上：Gross Profit（stockanalysis 口徑；FY2025 Q 頁為 6.795）"),
        ("oi", "營業利益", 0.821, 0.794, 1.359, 5.148, "Verified", "同上：Operating Income；FY2025 同 Q 頁"),
        ("pti", "稅前利益", 0.897, 0.749, 1.312, 5.022, "Verified", "同上：Pretax Income"),
        ("tax", "所得稅", 0.107, 0.078, 0.157, 0.710, "Verified", "同上：Income Tax Expense（106.62／77.87／157.39／709.73 百萬元）"),
        ("nci", "非控制權益損益", 0.0, 0.0, 0.0, 0.0, "Verified", "同上：Net Income 與 Net Income to Common 相同；核對列差額為百萬元四捨五入"),
        ("ni", "稅後淨利（歸屬母公司）", 0.791, 0.671, 1.154, 4.312, "Verified", "同上：Net Income to Common（790.65／671.35 百萬元）；FY2025 同 Q 頁"),
        ("sh", "EPS 換算股數（十億股）", 0.859, 0.861, 0.862, 0.861, "Verified", "同上：Shares Outstanding (Basic)（百萬股）"),
        ("eps", "EPS（公司公告，基本，RMB）", 0.92, 0.78, 1.34, 5.01, "Verified", "同上：EPS (Basic)（稀釋 0.92／0.78／1.34／5.00）；FY2025 同 Q 頁 5.01"),
    ],
    fy26_short="FY2026E 含 1H26 與 3Q26 實績",
    fy26_long="1H26 實績＋3Q26 實績（2026-10-29 第三季度報告）＋4Q26 假設（營收季增 25%、毛利率 33%）",
    eps_note="歷史為公司公告基本 EPS；預測為歸屬母公司淨利 ÷ 加權股數（含 H 股發行，同基本口徑）",
    gw_label="部署 GW（IT 電力）",
    same_names="部署 GW、每 GW AI 高速 PCB 金額、AI 高速 PCB 市場三項",
    a1_tail=(f"{VER}（{BUILT}）：個股 HTML 版面規格 v1.3（標準損益表、反向 DCF 頁、四個目標價、加權目標價敏感度主圖）＋三層協調（Coord 頁兩讀數目標價、讀取層級、C0（實績））；"
             "A1 標題版號由舊的 v0.1.2 更正；既有數值不變"),
    ver_line=(f"{VER}（{BUILT}）：改為個股 HTML 版面規格 v1.3（樣板 NVIDIA、台達電、台光電）並補三層協調：Coord 頁並列兩讀數目標價（tools/coord）、讀取層級（同源，不判定）、"
              "C0（實績）（三季報 2026-01～09 營收對 Siliconomics AI 高速 PCB 市場 × 勝宏份額；結論：無可比實績，份額由本模型錨點校準）；新增 RevDCF、Hist 頁、PL 標準損益表、Val 0.7、Sens 末表；"
              "既有數值、目標價與評等不變"),
)


def layer_rows(ver, CR):
    return [
        ("每 GW PCB 2026：加速器卡／運算板、交換器板、中板／背板／其他（US$M/GW；pcb_cmp_26／pcb_sw_26／pcb_mid_26）", '=TEXT(cmp_26,"0")&"／"&TEXT(sw_26,"0")&"／"&TEXT(mid_26,"0")', ver, "每 GW AI 高速 PCB 金額的 2026 年元件金額", None),
        ("每 GW PCB 指數 CY2027／28／29：加速器卡／運算板（pcb_cmpx_27–29）", '=TEXT(cmpx_27,"0.00")&"／"&TEXT(cmpx_28,"0.00")&"／"&TEXT(cmpx_29,"0.00")', ver, "每 GW 金額（2026＝1 的倍數）", None),
        ("每 GW PCB 指數 CY2027／28／29：交換器板（pcb_swx_27–29）", '=TEXT(swx_27,"0.00")&"／"&TEXT(swx_28,"0.00")&"／"&TEXT(swx_29,"0.00")', ver, "同上", None),
        ("每 GW PCB 指數 CY2027／28／29：中板／背板／其他（pcb_midx_27–29）", '=TEXT(midx_27,"0.00")&"／"&TEXT(midx_28,"0.00")&"／"&TEXT(midx_29,"0.00")', ver, "同上", None),
        ("2026 年部署中 Rubin 世代占比（vr26）", "=vr26", ver, "混合 FY2026 每 GW IT 資本支出", "0%"),
        ("情境 B 顆數係數 CY2028／CY2029（÷ 原始顆數；SIF_UnitsCoef_Base）", '=TEXT(um28_B,"0.000")&"／"&TEXT(um29_B,"0.000")', f"{ver}，經 tsmc/dist 成品 In 頁 um28_B、um29_B 讀取", "基準情境 B 的 AI IT 資本支出係數；A（原始顆數）、C（0.6）為個股自訂情境", None),
        ("AI IT 資本支出 CY2026–29（US$B；markets/network）", '=TEXT(itc_26,"#,##0")&"／"&TEXT(itc_27,"#,##0")&"／"&TEXT(itc_28,"#,##0")&"／"&TEXT(itc_29,"#,##0")', "markets/network 成品（市場層；Tokenomics 經橋接表進入）", "部署 GW 的分子（原始顆數口徑）", None),
        ("每 GW IT 資本支出（US$B；GB300／VR200；markets/chip_bridge）", '=TEXT(itpg_gb,"0.0")&"／"&TEXT(itpg_vr,"0.0")', "markets/chip_bridge 成品（Tokenomics IF_CapexIT，經橋接表唯讀）", "部署 GW 的分母", None),
        ("匯率（人民幣 ÷ 美元；本模型 In 頁 fx）", "=fx", "個股模型（中國經濟網中間價；Siliconomics 只有新台幣匯率）", "市場換算人民幣", "0.00"),
        ("勝宏份額、價格指數、產能指數（不在 Siliconomics）", "本模型", "個股模型（Siliconomics M16 註明：價格指數與各公司份額留在公司模型）", "FY2026 有效份額＝AI 相關營收錨點 ÷ 本模型市場（校準）", None),
    ]


def same_blocks(SIL, pc):
    N = SIL.NAMES
    return [
        ("gw", "部署 GW（IT 電力）", "GW", "部署 GW（IT 電力）", list(N["SIF_DeployGW_Base"]),
         "Derived：AI IT 資本支出（markets/network）× 情境 B 顆數係數 ÷ 每 GW IT 資本支出（bridge）", f"Derived：Siliconomics {SIL.VERSION} SIF_DeployGW_Base（依規則 v1.4 的基準）", "0.00"),
        ("pgw", "每 GW AI 高速 PCB 金額", "US$M/GW", "每 GW AI 高速 PCB 金額（三項合計）", list(N["SIF_PCBperGW"]), "Derived：三項元件金額 × 指數（讀自 Siliconomics In 頁 pcb_*）", f"Derived：Siliconomics {SIL.VERSION} SIF_PCBperGW", "#,##0"),
        ("mkt", "AI 高速 PCB 市場（US$B）", "US$B", "AI 高速 PCB 市場（美元；不含價格指數）", list(N["SIF_PCBMkt_Base"]), "Derived：部署 GW × 每 GW AI 高速 PCB 金額", f"Derived：Siliconomics {SIL.VERSION} SIF_PCBMkt_Base", "0.00"),
    ]


def lines(wb):
    P = wb["PL"]; SE = wb["Sens"]
    rb = blockB(P)
    pb = lambda label, c: f"PL!{c}{lab(P, label, rb, exact=True, end=rb + 70)}"
    sl = lambda k: f"Sens!L{lab(SE, k)}"
    return dict(
        hero=lambda tp, up, rec, dcf, cadj: (T("加權目標價 "), f'TEXT({tp},"#,##0.0")', T(" 元人民幣，較現價"), f'IF({up}<0,"低 ","高 ")', f'TEXT(ABS({up}),"0.0%")', T("，評等"), rec,
                                             T("。現金流折現 "), f'TEXT({dcf},"#,##0.0")', T(" 元、同業比較法（相對倍數調整後）"), f'TEXT({cadj},"#,##0.0")',
                                             T(" 元各半；AI 加速卡與運算板 PCB 是主力，目標價對份額、每 GW PCB 金額與擴產最敏感。")),
        premise="產業部署量與每 GW AI 高速 PCB 金額依 Siliconomics 基準（行業變數）、勝宏份額相對值、價格指數與擴產（假設）成立",
        o2=(T("AI 高速 PCB 是成長主力：情境 B AI 相關營收由 FY2026 約 "), f'TEXT({pb(AIR, "D")}*10,"#,##0")', T(" 億元增至 FY2028 約 "), f'TEXT({pb(AIR, "F")}*10,"#,##0")',
            T(" 億元人民幣，占營收由 "), f'TEXT({pb("AI 相關 ÷ 營收", "D")},"0%")', T(" 升到 "), f'TEXT({pb("AI 相關 ÷ 營收", "F")},"0%")', T("；同期勝宏有效份額由 "),
            f'TEXT({pb("勝宏有效份額", "D")},"0.0%")', T(" 降到 "), f'TEXT({pb("勝宏有效份額", "F")},"0.0%")', T("。")),
        o3=(T("最大的變數是份額與每 GW PCB 金額（情境 A 為底，FY2028 EPS）：份額流失 "), f'TEXT({sl("份額流失（低）")},"+0.0%;-0.0%")', T("，每 GW 金額成長取低值 "),
            f'TEXT({sl("每 GW 金額成長 低")},"+0.0%;-0.0%")', T("，價格回落快 "), f'TEXT({sl("價格回落快")},"+0.0%;-0.0%")', T("，AI 部署 ×0.8 為 "),
            f'TEXT({sl("AI 部署 ×0.8")},"+0.0%;-0.0%")', T("；基準情境 B 的 FY2028 EPS 為 "), f'TEXT({pb("EPS", "F")},"0.00")', T(" 元人民幣。")),
        o5_tail="本模型的部署 GW、每 GW AI 高速 PCB 金額與市場同 Siliconomics 基準；C0（實績）無可比實績（公司未揭露 2026 年 AI 營收金額，份額由本模型錨點校準）。",
        drv=(T("分部營收推導（情境 B）：AI 相關營收＝MIN（部署 GW × 每 GW AI 高速 PCB 金額 × 勝宏有效份額 × 價格指數，產能上限 − 非 AI）；非 AI 依年增率；"
               "相加為合併營收。FY2028 AI 相關 "), f'TEXT({pb(AIR, "F")},"#,##0.0")', T("、非 AI "), f'TEXT({pb("非 AI 營收", "F")},"#,##0.0")',
             T("，合計 "), f'TEXT({pb("營收", "F")},"#,##0.0")', T(" 十億元人民幣。")),
    )


def c0_actuals(wb, G):
    CW = wb["Coord"]; Q = wb["Q"]; P = wb["PL"]
    rb = blockB(P)
    sh0 = lab(P, "勝宏有效份額（FY2026 校準）", rb)
    qr = lab(Q, "營收")
    S = Sheet(CW, G["end"] + 3)
    R = {}
    R["title"] = S.bold("4. C0（實績）：公司實績 vs Siliconomics 層級隱含值（規則 v1.4.2 C0 第 2 點；工作單 2026-11-14；差距＝實績 ÷ 層級隱含值 − 1）")
    R["method"] = S.text("方法：個股側＝勝宏已公布實績（A 股不發布月營收；用第三季度報告的前三季累計營收與 3Q26 單季營收）；層級側＝Siliconomics 現行版 AI 高速 PCB 市場（CY2026）× 勝宏 FY2026 有效份額 × 匯率（人民幣；市場連到本頁第 2 節層級值列）。"
                         "層級值是日曆年，故主讀數＝2026-01～09 累計實績 vs 層級 CY2026 × 9 ÷ 12（線性時差近似，Derived）；第二讀數＝3Q26 × 4 vs 層級 CY2026（約等於 3Q26 中點 8 月中的年化速率，"
                         "季中點晚於年中、成長年偏高）。判定：兩讀數差距都超過 ±15% 且同號 → 對帳中；只一個超過 → 時差邊界（參考）；都在內 → 一致。公司沒有揭露 2026 年 AI 相關營收金額"
                         "（H 股招股書只有 2024–2025 年 AI／HPC 占比，經券商轉述），故 AI 相關列無可比實績；以占比換算者只列參考。")
    R["limit"] = S.text("限制（同源校準）：Siliconomics 沒有勝宏份額（M16 註明份額留在公司模型；每 GW PCB 金額原在本模型，後移入 Siliconomics）；本模型 FY2026 有效份額＝AI 相關營收錨點（錨點營收 − 非 AI）÷ 本模型市場，"
                        "而本模型的市場就是 Siliconomics 的市場（部署 GW 與每 GW 金額皆讀自 Siliconomics）。層級隱含 AI 相關營收＝市場 × 份額 × 匯率＝AI 相關營收錨點，市場大小在比式中被約掉；"
                        "以錨點 AI 占比 × 實績營收對照時，兩讀數代數上等於「實績營收 ÷ 營收錨點 − 1」，實際只檢驗本模型營收錨點。結論：無可比實績（層級值同源）；缺勝宏 2026 年 AI 相關營收金額揭露。")
    S.skip()
    S.bold("4a. 公司已公布實績與揭露（輸入；藍字＝來自原文的常數）")
    S.hdr(["項目", "值", "", "", "", "單位", "口徑、來源、證據等級", "來源網址"])
    R["cum9"] = S.row("累計：2026-01～09 營收（主讀數的實績基礎；第三季度報告）", 18.461, None, "RMB B", "Verified：公司 2026 年第三季度報告：前三季營業收入 184.61 億元、年增 30.77%（中證網 2026-10-30 轉載）", blue=True, url=Q3)
    R["q3x"] = S.row("3Q26 營收（第三季度報告）", 6.831, None, "RMB B", "Verified：同上：3Q26 營業收入 68.31 億元、年增 34.30%、季增 11.81%", blue=True, url=Q3)
    R["q3"] = S.row("3Q26 營收（最近一季；第二讀數的實績基礎；Q 頁）", f"=Q!D{qr}", None, "RMB B", "Verified：Q 頁 3Q26 欄（1H26 × 季增 q3g，q3g 由季報反推）")
    R["chk1"] = S.row("檢查：3Q26 營收（Q 頁）− 季報（應接近 0）", f"=B{R['q3']}-B{R['q3x']}", None, "RMB B", "Derived")
    R["chk2"] = S.row("檢查：前三季累計 − 3Q26 − Q 頁 1H26 營收（應接近 0；季報四捨五入）", f"=B{R['cum9']}-B{R['q3']}-Q!C{qr}", None, "RMB B", "Derived：1H26 116.29 億元（Q 頁，公司半年報）")
    R["ai"] = S.row("公司揭露：2026 年 AI 相關營收金額", "未揭露", None, "", "公司未揭露 2026 年 AI 相關營收金額；H 股招股書 AI／HPC 占營收 2024 年 6.6%、2025 年 43.2%（摩根大通轉述，Interested-party）為 2025 年以前口徑", fmt="General")
    R["months"] = S.row("已公布月數（2026-01～09）", 9, None, "月", "輸入：第三季度報告（2026-10-29 公告）；A 股無月營收", fmt="0", blue=True)
    S.skip()
    S.bold("4b. 層級隱含值（Siliconomics 現行版；市場連到本頁第 2 節層級值列）")
    S.hdr(["項目", "CY2026", "", "", "", "單位", "口徑、來源、證據等級"])
    R["lmkt"] = S.row("AI 高速 PCB 市場（層級值）", f"=B{G['mkt'][1]}", None, "US$B", "Derived：本頁第 2 節 Siliconomics SIF_PCBMkt_Base（CY2026）")
    R["lsh"] = S.row("勝宏 FY2026 有效份額（本模型校準；Siliconomics 無此量）", f"=PL!D{sh0}", None, "%", "Derived（本模型校準）：AI 相關營收錨點 ÷ 本模型 AI 高速 PCB 市場（PL 頁情境 B）", fmt="0.0%")
    R["lai"] = S.row("層級隱含：勝宏 AI 相關營收（市場 × 份額 × 匯率）", f"=B{R['lmkt']}*B{R['lsh']}*fx", None, "RMB B", "Derived：匯率 In 頁 fx（人民幣 ÷ 美元）", fmt="#,##0.0")
    S.skip()
    R["tbl"] = S.bold("C0（實績）：公司實績 vs 層級隱含值（HTML 用）")
    R["hdr"] = S.hdr(["項目", "主讀數：2026-01～09 累計 vs 層級 CY2026 × 9 ÷ 12", "第二讀數：3Q26 × 4 vs 層級 CY2026", "", "", "單位", "口徑、來源、證據等級"])
    fr = f"B{R['months']}/12"
    pct = "+0.0%;-0.0%"
    S.row("AI 相關：公司實績", "無（未揭露 AI 營收）", "無（未揭露 AI 營收）", "RMB B", "公司未揭露 2026 年 AI 相關營收金額；不以模型值或估計占比補位", fmt="General")
    l = S.row("AI 相關：層級隱含", f"=B{R['lai']}*{fr}", f"=B{R['lai']}", "RMB B", "Derived：層級隱含 AI 相關營收（4b）；主讀數 × 9 ÷ 12", fmt="#,##0.0")
    R["j_ai"] = S.row("AI 相關：判定", "無可比實績（層級值同源）", "無可比實績（層級值同源）", "", "缺 AI 營收揭露；份額由本模型校準，市場大小未受檢驗；不計入對帳中", fmt="General")
    sh_ai = "(a26_ai/a26_rev)"
    a2 = S.row("AI 相關（參考，同源校準）：錨點 AI 占比 × 公司營收實績", f"=B{R['cum9']}*{sh_ai}", f"=B{R['q3']}*{sh_ai}*4", "RMB B", "Verified 營收 × 本模型 FY2026 錨點 AI 占比（a26_ai ÷ a26_rev）；不是公司揭露金額", fmt="#,##0.0")
    g2 = S.row("AI 相關（參考，同源校準）：差距（÷ 層級隱含 − 1）", f"=B{a2}/B{l}-1", f"=C{a2}/C{l}-1", fmt=pct)
    R["j_ref"] = S.row("AI 相關（參考，同源校準）：判定", "參考（同源校準）：不判定", "參考（同源校準）：不判定", "", "兩讀數代數上等於實績營收 ÷ 營收錨點 − 1（占比與市場大小約掉），只檢驗本模型營收錨點", fmt="General")
    R["ub"] = S.row("上限檢查：層級隱含 AI 相關 ÷ 公司營收實績（主讀數期間；應 ≤ 100%）", f"=B{l}/B{R['cum9']}", None, "%", "Derived：層級 AI 相關不得超過公司全部營收；通過＝合理，不是一致性判定", fmt="0.0%")
    R["j_ub"] = S.row("上限檢查：結果", f'=IF(B{R["ub"]}<=1,"通過","不通過")', None, "", "", fmt="General")
    S.skip()
    S.bold("參考：本模型 FY2026 營收錨點對季報實績（模型對實績，不是層級對帳；不計入 C0）")
    R["anc"] = S.row("FY2026 營收錨點（Q 頁）", f"=Q!F{qr}", None, "RMB B", "Derived／Assumed：1H26、3Q26 實績＋4Q26 假設季增 25%", fmt="#,##0.0")
    R["imp"] = S.row("錨點隱含 4Q26 營收（錨點 − 前三季累計）", f"=B{R['anc']}-B{R['cum9']}", None, "RMB B", "Derived", fmt="#,##0.00")
    R["impg"] = S.row("隱含 4Q26 ÷ 3Q26 實績 − 1", f"=B{R['imp']}/B{R['q3x']}-1", None, "%", "Derived：正值＝錨點假設 4Q26 高於 3Q26", fmt=pct)
    S.skip()
    R["sum0"] = S.bold("C0（實績）結果摘要（HTML 用；公式帶數字）")
    nopen = f'COUNTIF(B{R["j_ai"]}:C{R["j_ref"]},"對帳中")'
    R["sum1"] = S.text(F(T("C0（實績）結論：無可比實績（層級值同源）：勝宏未揭露 2026 年 AI 相關營收金額；Siliconomics 沒有勝宏份額，本模型份額以營收錨點對 Siliconomics 市場校準。對帳中項目 "), nopen,
                         T(" 項。參考：以錨點 AI 占比換算，差距主讀數 "), f'TEXT(B{g2},"+0.0%;-0.0%")', T("、第二讀數 "), f'TEXT(C{g2},"+0.0%;-0.0%")',
                         T("，只反映營收錨點與實績的差異，不判定。")))
    R["sum2"] = S.text(F(T("上限檢查：層級隱含 AI 相關（主讀數期間）占公司營收實績 "), f'TEXT(B{R["ub"]},"0%")', T("，"), f'B{R["j_ub"]}',
                         T("；前三季營收 "), f'TEXT(B{R["cum9"]}*10,"#,##0.0")', T(" 億元人民幣，營收錨點隱含 4Q26 "), f'TEXT(B{R["imp"]}*10,"#,##0.0")', T(" 億元（較 3Q26 "),
                         f'TEXT(B{R["impg"]},"+0%;-0%")', T("）。")))
    R["sum3"] = S.text("限制與待查：公司未揭露 2026 年 AI 相關營收，Siliconomics 的 AI 高速 PCB 市場也沒有獨立的勝宏份額，故市場大小未獲獨立驗證；A 股無月營收，主讀數只到 9 月，年報公布後於下一版更新。")
    R["end"] = S.r - 1
    return R


CAS_OLD = "口徑：GAAP（台灣採用的 IFRS 合併財報，歸屬母公司淨利）。勝宏不另發布 Non-GAAP，法說與市場共識 EPS 皆以此口徑，故本表不做 Non-GAAP 調整。"
CAS_NEW = ("口徑：中國企業會計準則（CAS）合併財報，歸屬母公司淨利。勝宏另揭露扣除非經常性損益後淨利（扣非），本表不採用；市場共識 EPS 以歸母淨利為準，故本表不做扣非調整。")


def rmb(wb, A):
    """c13 新增格的幣別字樣改為人民幣（只動本版新增的頁與列）。"""
    rep = (("NT$B", "RMB B"), ("NT$", "RMB"), ("新台幣", "人民幣"))
    def fix(ws, r0, r1=None):
        n = 0
        for row in ws.iter_rows(min_row=r0, max_row=r1 or ws.max_row):
            for c in row:
                v = c.value
                if isinstance(v, str):
                    w = v.replace(CAS_OLD, CAS_NEW)
                    for a, b in rep:
                        w = w.replace(a, b)
                    if w != v:
                        c.value = w; n += 1
        return n
    n = fix(wb["Hist"], 1) + fix(wb["RevDCF"], 1) + fix(wb["Coord"], A["G"]["c0"])
    n += fix(wb["PL"], A["PL"]["start"]) + fix(wb["Val"], A["VG"]) + fix(wb["Sens"], A["ST"]["t0"])
    n += fix(wb["README"], min(A["RM"].values()) - 1)
    return n


def apply(wb, SIL, CR):
    C.update(layer_rows=layer_rows, same_blocks=same_blocks)
    msg = c13.apply(wb, SIL, CR, C, lines, c0_actuals)
    return msg + f"；幣別字樣改人民幣 {rmb(wb, wb._v13)} 格"
