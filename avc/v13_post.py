# 奇鋐 v0.3.5（2026-11-15）後處理：個股 HTML 版面規格 v1.3＋三層協調兩讀數＋C0（實績）（工作單 docs/workorders/20261114_第二批13家_v1.3與三層協調.md；
# 規則 docs/plan/三層基準協調規則.md v1.4.2；樣板 台達電 delta/v056_post.py、v057_post.py）。共通步驟在 avc/c13.py（電源散熱四家同一份），本檔只放奇鋐的設定。
# build_avc.py 在 Coord 頁建好之後、存檔前呼叫 apply(wb, SIL, coord_rows)。alt_run.py 建置（無 Coord 頁）不呼叫。只新增，不改既有數值格。
import os, sys
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import c13
from c13 import T, F, lab, Sheet, judge, SAME

VER, BUILT = "v0.3.6", "2026-11-15"
SA = "https://stockanalysis.com/quote/tpe/3017/financials/income-statement/"
M_OCT = "https://finance.technews.tw/2026/11/08/avc-3017-202610-financial-report/"
M_SEP = "https://finance.technews.tw/2026/10/07/avc-3017-202609-financial-report/"
Q3A = "https://uanalyze.com.tw/articles/9857458877"
Q3B = "https://www.pocket.tw/school/articles/article-mwbrnuav/"
Q3C = "https://uanalyze.com.tw/articles/9896958909"
H1 = "https://uanalyze.com.tw/articles/5815653290"

C = dict(
    name="奇鋐", ver=VER, built=BUILT,
    hist_title=("Hist — 損益表歷史（v0.3.5 起；PL 頁「標準損益表」的歷史三年與年增率用）。金額：十億新台幣；EPS：新台幣；歸屬母公司。"
                f"來源：stockanalysis 年度損益表（百萬元，2026-11-15 讀取；{SA}），FY2025 與 Q 頁（公司年報）一致。營業外收益＝稅前 − 營業利益；營業費用＝毛利 − 營業利益"),
    hist=[
        ("rev", "營收", 56.017, 59.194, 71.761, 139.639, "Verified", f"{SA}：Revenue（百萬元）；FY2025 同 Q 頁"),
        ("gp", "毛利", 10.850, 12.388, 16.870, 36.034, "Verified", "同上：Gross Profit"),
        ("oi", "營業利益", 6.306, 7.424, 10.821, 27.555, "Verified", "同上：Operating Income（Q 頁 FY2025 27.551 為營益率 19.73% 推導值）"),
        ("pti", "稅前利益", 6.489, 8.025, 12.352, 28.786, "Verified", "同上：Pretax Income"),
        ("tax", "所得稅", 1.871, 2.211, 3.159, 7.836, "Verified", "同上：Income Tax Expense"),
        ("nci", "非控制權益損益", 0.456, 0.510, 1.020, 1.764, "Verified", "同上：Minority Interest in Earnings（富世達等）"),
        ("ni", "稅後淨利（歸屬母公司）", 4.162, 5.305, 8.172, 19.186, "Verified", "同上：Net Income to Common；FY2025 同 Q 頁"),
        ("sh", "EPS 換算股數（十億股）", 0.353, 0.376, 0.385, 0.390, "Verified", "同上：Shares Outstanding (Basic)"),
        ("eps", "EPS（公司公告，基本，NT$）", 11.78, 14.11, 21.21, 49.17, "Verified", "同上：EPS (Basic)（稀釋 11.67／13.71／20.62／48.24，差異來自可轉債）；FY2025 同 Q 頁 49.17"),
    ],
    fy26_short="FY2026E 含 1–3Q26 財報實績",
    fy26_long="1–3Q26 財報實績（公開資訊觀測站 2026-11-10）＋4Q26 假設",
    eps_note="歷史為公司公告基本 EPS；預測為歸屬母公司淨利 ÷ 股數（本模型不計可轉債稀釋，同基本口徑）",
    gw_label="部署 GW（IT 電力）",
    same_names="部署 GW、液冷滲透率、每 GW 液冷內容、AI 液冷市場與奇鋐有效份額五項",
    resid=("shsum", "雙鴻、酷碼、台達電、Boyd、CoolIT 等；不得歸零（C0 第 3 點）"),
    a1_tail=("v0.3.5（2026-11-15）：個股 HTML 版面規格 v1.3（標準損益表、反向 DCF 頁、四個目標價、加權目標價敏感度主圖）＋三層協調（Coord 頁兩讀數目標價、讀取層級、C0（實績））；"
             "A1 標題版號由舊的 v0.2.1 更正；既有數值不變；v0.3.6（2026-11-15）：滾入 3Q26 財報（公開資訊觀測站 2026-11-10）：FY2026 錨點改為 1–3Q26 財報實績＋4Q26 假設（4Q 假設不改）；DCF 資產負債表日 2026-06-30 → 09-30（借款含一年內到期長期負債、納入流動金融資產、股利已於 9-23 發放不再扣、stub 只剩 4Q26）；評價用股數改 9-30 股本 ÷ 面額＋預收股本約當；C0 第二讀數改用財報 3Q26 營收。"),
    ver_line=("v0.3.6（2026-11-15）：滾入 3Q26 財報（公開資訊觀測站 2026-11-10）：FY2026 錨點改為 1–3Q26 財報實績＋4Q26 假設（4Q 假設不改）；DCF 資產負債表日 2026-06-30 → 09-30（借款含一年內到期長期負債、納入流動金融資產、股利已於 9-23 發放不再扣、stub 只剩 4Q26）；評價用股數改 9-30 股本 ÷ 面額＋預收股本約當；C0 第二讀數改用財報 3Q26 營收。v0.3.5（2026-11-15）：改為個股 HTML 版面規格 v1.3（樣板 NVIDIA、台達電）並補三層協調：Coord 頁並列兩讀數目標價（tools/coord）、讀取層級（同源，不判定）、"
              "C0（實績）（月營收 2026-01～10、3Q26 財報對 Siliconomics 液冷市場 × 奇鋐份額；結論：無可比實績，層級份額由本公司校準）；新增 RevDCF、Hist 頁、PL 標準損益表、Val 0.7、Sens 末表；"
              "既有數值、目標價與評等不變"),
)


def layer_rows(ver, CR):
    return [
        ("液冷滲透率 CY2026／27／28／29（情境 B；lc_pen_26–29）", '=TEXT(lpen_26,"0%")&"／"&TEXT(lpen_27,"0%")&"／"&TEXT(lpen_28,"0%")&"／"&TEXT(lpen_29,"0%")', ver, "液冷市場＝部署 GW × 滲透率 × 每 GW 液冷內容（In 頁 B 節；情境 A 取高、C 取低）", None),
        ("每架液冷內容調整 CY2027／28／29（÷ 橋接表；lc_adj_27–29）", '=TEXT(lcadj_27,"0.00")&"／"&TEXT(lcadj_28,"0.00")&"／"&TEXT(lcadj_29,"0.00")', ver, "每 GW 液冷內容（Rubin 全液冷、ASIC 液冷）", None),
        ("分歧管每架支數（lc_mani）", "=mani", ver, "每架液冷內容含分歧管", "0"),
        ("奇鋐份額相對值 CY2027／28／29（÷ CY2026；lc_avrel_27–29）", '=TEXT(rel_27,"0.00")&"／"&TEXT(rel_28,"0.00")&"／"&TEXT(rel_29,"0.00")', ver, "有效份額路徑；CY2026 有效份額由本模型以券商液冷占比 × 營收錨點校準後寫入 Siliconomics（lc_avsh_26）", None),
        ("匯率（NT$/US$；fx_twd）", "=fx", ver, "液冷市場換算新台幣", "0.00"),
        ("情境 B 顆數係數 CY2028／CY2029（÷ 原始顆數；SIF_UnitsCoef_Base）", '=TEXT(um28_B,"0.000")&"／"&TEXT(um29_B,"0.000")', f"{ver}，經 tsmc/dist 成品 In 頁 um28_B、um29_B 讀取", "基準情境 B 的 AI IT 資本支出係數；A（原始顆數）、C（0.6）為個股自訂情境", None),
        ("AI IT 資本支出 CY2026–29（US$B；markets/network）", '=TEXT(itc_26,"#,##0")&"／"&TEXT(itc_27,"#,##0")&"／"&TEXT(itc_28,"#,##0")&"／"&TEXT(itc_29,"#,##0")', "markets/network 成品（市場層；Tokenomics 經橋接表進入）", "部署 GW 的分子（原始顆數口徑）", None),
        ("每 GW IT 資本支出（US$B；GB300／VR200；markets/chip_bridge）", '=TEXT(itpg_gb,"0.0")&"／"&TEXT(itpg_vr,"0.0")', "markets/chip_bridge 成品（Tokenomics IF_CapexIT，經橋接表唯讀）", "部署 GW 的分母", None),
    ]


def same_blocks(SIL, pc):
    N, V = SIL.NAMES, SIL.V
    ash, rel = V["lc_avsh_26"]["base"], [1.0] + [V[f"lc_avrel_{y}"]["base"] for y in (27, 28, 29)]
    aush = V["lc_aush_26"]["base"]
    sh = pc("奇鋐有效份額")
    return [
        ("gw", "部署 GW（IT 電力）", "GW", "部署 GW（IT 電力）", list(N["SIF_DeployGW_Base"]),
         "Derived：AI IT 資本支出（markets/network）× 情境 B 顆數係數 ÷ 每 GW IT 資本支出（bridge）", f"Derived：Siliconomics {SIL.VERSION} SIF_DeployGW_Base（依規則 v1.4 的基準）", "0.00"),
        ("pen", "液冷滲透率", "%", "液冷滲透率（部署 GW 中液冷占比）", list(N["SIF_LiqPen_Base"]), "讀自 Siliconomics（In 頁 lpen_26–29）", f"Assumed：Siliconomics {SIL.VERSION} SIF_LiqPen_Base", "0.0%"),
        ("pgw", "每 GW 液冷內容", "US$M/GW", "每 GW 液冷內容（機架內液冷＋分歧管）", list(N["SIF_LiqPerGW"]), "Derived：橋接表機架內液冷＋分歧管 × 每 GW 機架數 × 內容調整", f"Derived：Siliconomics {SIL.VERSION} SIF_LiqPerGW", "#,##0"),
        ("mkt", "AI 液冷市場（US$B）", "US$B", "AI 液冷市場（奇鋐可做部分；美元）", list(N["SIF_LiqMkt_Base"]), "Derived：部署 GW × 滲透率 × 每 GW 液冷內容", f"Derived：Siliconomics {SIL.VERSION} SIF_LiqMkt_Base", "0.00"),
        ("sh", "奇鋐有效份額", "%", "奇鋐有效份額", [f"={ash}*{x}" for x in rel],
         "Derived：FY2026 以券商液冷占比 × 營收錨點 ÷ 市場校準，之後 × 份額相對值（Assumed）", f"Derived（本公司校準值，Siliconomics 快照）：lc_avsh_26 × lc_avrel_27–29（CY2026 快照值 {ash:.4f}）", "0.0%"),
        ("shsum", "已建模個股份額加總（奇鋐本模型＋雙鴻層級值）", "%", [f"{s}+{aush}*{x}" for s, x in zip(sh, rel)], list(N["SIF_LiqShareSum_Base"]),
         "Derived：奇鋐取本模型；雙鴻取 Siliconomics lc_aush_26 × 份額相對值（同奇鋐一組值）", f"Derived：Siliconomics {SIL.VERSION} SIF_LiqShareSum_Base", "0.0%"),
    ]


def lines(wb):
    P = wb["PL"]; SE = wb["Sens"]
    rb = lab(P, "情境 B：")
    pb = lambda label, c: f"PL!{c}{lab(P, label, rb, exact=True, end=rb + 70)}"
    sl = lambda k: f"Sens!L{lab(SE, k)}"
    return dict(
        hero=lambda tp, up, rec, dcf, cadj: (T("加權目標價 "), f'TEXT({tp},"#,##0")', T(" 元，較現價"), f'IF({up}<0,"低 ","高 ")', f'TEXT(ABS({up}),"0.0%")', T("，評等"), rec,
                                             T("。現金流折現 "), f'TEXT({dcf},"#,##0")', T(" 元、同業比較法（相對倍數調整後）"), f'TEXT({cadj},"#,##0")',
                                             T(" 元各半；液冷營收續成長，目標價對每架液冷內容、AI 部署量與份額最敏感。")),
        premise="產業部署量依 Siliconomics 基準、液冷滲透率與每架液冷內容（Siliconomics 行業變數）與奇鋐份額相對值（假設）成立",
        o2=(T("液冷是成長主力：情境 B 液冷營收由 FY2026 約 "), f'TEXT({pb("液冷營收", "D")}*10,"#,##0")', T(" 億元增至 FY2028 約 "), f'TEXT({pb("液冷營收", "F")}*10,"#,##0")',
            T(" 億元，占營收由 "), f'TEXT({pb("液冷 ÷ 營收", "D")},"0%")', T(" 升到 "), f'TEXT({pb("液冷 ÷ 營收", "F")},"0%")', T("；同期奇鋐有效份額由 "),
            f'TEXT({pb("奇鋐有效份額", "D")},"0.0%")', T(" 降到 "), f'TEXT({pb("奇鋐有效份額", "F")},"0.0%")', T("。")),
        o3=(T("最大的變數是每架液冷內容與 AI 部署量：每架液冷內容取低值使情境 A 的 FY2028 EPS "), f'TEXT({sl("每架液冷內容 低")},"+0.0%;-0.0%")', T("，AI 部署 ×0.8 為 "),
            f'TEXT({sl("AI 部署 ×0.8")},"+0.0%;-0.0%")', T("，份額流失 "), f'TEXT({sl("份額流失（低）")},"+0.0%;-0.0%")', T("；基準情境 B 的 FY2028 EPS 為 "),
            f'TEXT({pb("EPS", "F")},"0.0")', T(" 元。")),
        o5_tail="本模型的部署 GW 與液冷市場同 Siliconomics 基準；C0（實績）無可比實績（公司未揭露液冷營收，層級份額由本公司校準）。",
        drv=(T("分部營收推導（情境 B）：液冷營收＝部署 GW × 液冷滲透率 × 每 GW 液冷內容 × 奇鋐有效份額；AI 氣冷與機構＝部署 GW × 每 GW 營收；一般伺服器與網通、非伺服器各依年增率；"
               "相加為合併營收。FY2028 液冷 "), f'TEXT({pb("液冷營收", "F")},"#,##0")', T("、AI 氣冷與機構 "), f'TEXT({pb("AI 氣冷與機構營收", "F")},"#,##0")',
             T("、一般伺服器與網通 "), f'TEXT({pb("一般伺服器與網通營收", "F")},"#,##0")', T("、非伺服器 "), f'TEXT({pb("非伺服器營收（3C、系統、富世達）", "F")},"#,##0")',
             T("，合計 "), f'TEXT({pb("營收", "F")},"#,##0")', T(" 十億元。")),
    )


def c0_actuals(wb, G):
    CW = wb["Coord"]; Q = wb["Q"]
    S = Sheet(CW, G["end"] + 3)
    R = {}
    R["title"] = S.bold("4. C0（實績）：公司實績 vs Siliconomics 層級隱含值（規則 v1.4.2 C0 第 2 點；工作單 2026-11-14；差距＝實績 ÷ 層級隱含值 − 1）")
    R["method"] = S.text("方法：個股側＝奇鋐已公布實績（月營收、3Q26 財報（公開資訊觀測站 2026-11-10）與法說）；層級側＝Siliconomics 現行版 AI 液冷市場（CY2026）× 奇鋐有效份額 × 匯率（連到本頁第 2 節層級值列）。"
                         "層級值是日曆年，故主讀數＝2026-01～10 累計實績 vs 層級 CY2026 × 10 ÷ 12（線性時差近似，Derived）；第二讀數＝3Q26 財報營收 × 4 vs 層級 CY2026（約等於 3Q26 中點 8 月中的年化速率，"
                         "季中點晚於年中、成長年偏高）。判定：兩讀數差距都超過 ±15% 且同號 → 對帳中；只一個超過 → 時差邊界（參考）；都在內 → 一致。公司沒有揭露液冷營收金額或占比（3Q26 法說只揭露伺服器與網通占比），"
                         "故液冷列無可比實績；以券商液冷占比換算者只列參考。")
    R["limit"] = S.text("限制（同源校準）：Siliconomics 的奇鋐 CY2026 有效份額是本模型以「券商液冷占比 32% × FY2026 營收錨點 ÷ 本庫液冷市場」校準後寫入，層級隱含液冷營收＝32% × 營收錨點，"
                        "Siliconomics 液冷市場大小（部署 GW × 滲透率 × 每 GW 內容）在比式中被約掉；以 32% × 實績營收對照時，兩讀數代數上等於「實績營收 ÷ 營收錨點 − 1」，實際只檢驗本模型營收錨點，"
                        "且 32% 為券商估計（Interested-party），不是公司揭露。結論：無可比實績（層級值同源）；缺奇鋐液冷營收金額揭露。公司若揭露液冷營收，才能獨立檢驗液冷市場大小。")
    S.skip()
    S.bold("4a. 公司已公布實績與揭露（輸入；藍字＝來自原文的常數）")
    S.hdr(["項目", "值", "", "", "", "單位", "口徑、來源、證據等級", "來源網址"])
    R["sep"] = S.row("月營收：2026-09", 21.018, None, "NT$B", "Verified：月營收公告 210.18 億元，月增 7.9%、年增 44.9%，累計前 9 月 1,572.48 億元（TechNews 2026-10-07）", blue=True, url=M_SEP)
    R["oct"] = S.row("月營收：2026-10", 22.103, None, "NT$B", "Verified：月營收公告 221.03 億元，月增 5.2%、年增 34.3%（TechNews 2026-11-08）", blue=True, url=M_OCT)
    R["cum9"] = S.row("累計：2026-01～09", 157.248, None, "NT$B", "Verified：月營收公告 1,572.48 億元，年增 71.2%（TechNews 2026-10-07）；3Q26 財報前三季營收 1,572.48 億元（優分析 2026-11-10）", blue=True, url=M_SEP)
    R["cum10"] = S.row("累計：2026-01～10（主讀數的實績基礎）", 179.351, None, "NT$B", "Verified：月營收公告 1,793.51 億元，年增 65.6%（TechNews 2026-11-08）", blue=True, url=M_OCT)
    R["chk1"] = S.row("檢查：累計 9 月 + 10 月 − 累計 10 月（應為 0）", f"=B{R['cum9']}+B{R['oct']}-B{R['cum10']}", None, "NT$B", "Derived")
    R["q3"] = S.row("3Q26 營收（最近一季；第二讀數的實績基礎；財報）", "=q3_rev", None, "NT$B", "Verified：In 頁 q3_rev＝公開資訊觀測站 t164sb04 3Q26 營業收入 59,089,348 千元；季增 20.3%、年增 51.8%（優分析 2026-11-10 同值）")
    R["chk2"] = S.row("檢查：累計 9 月（月營收）− 1–3Q26 財報營收（應接近 0）", f"=B{R['cum9']}-ytd_rev", None, "NT$B", "Derived：月營收前 9 月 1,572.48 億元 − 公開資訊觀測站 1–9 月營業收入 1,572.48 億元（差 0.0003，四捨五入）")
    R["chk3"] = S.row("檢查：3Q26 財報營收 − 7–9 月月營收合計（應接近 0）", f"=B{R['q3']}-(18.59+19.481+21.018)", None, "NT$B", "Derived：7 月 185.90、8 月 194.81、9 月 210.18 億元")
    R["gm3"] = S.row("3Q26 毛利率（參考；v0.3.6 已滾入 Q 頁）", "=q3_gp/q3_rev", None, "%", "Verified：公開資訊觀測站 3Q26 營業毛利 ÷ 營收＝33.86%，季增 1.29 個百分點（優分析 2026-11-10 同值）", fmt="0.00%")
    R["ni3"] = S.row("3Q26 歸屬母公司淨利（參考；v0.3.6 已滾入 Q 頁）", "=q3_ni", None, "NT$B", "Verified：公開資訊觀測站 3Q26 母公司業主淨利 11,563,760 千元、基本 EPS 29.32 元、營業利益 168.56 億元（口袋學堂、優分析同值）；v0.3.5 Q 頁 3Q26 模型 11.223")
    R["srv9"] = S.row("公司揭露：伺服器與網通占營收（2026 前三季）", 0.66, None, "%", "Verified（公司法說，經媒體轉述）：「前 3 季伺服器與網通合計營收年增 122%、占整體 66%」；3Q26 單季伺服器約 388 億元、占 65%（優分析 2026-11-11）", fmt="0.0%", blue=True, url=Q3C)
    R["lc"] = S.row("公司揭露：液冷營收或占比", "未揭露", None, "", "3Q26 法說與財報未揭露液冷（水冷）營收金額或占比（優分析、口袋學堂 2026-11-10 摘要皆未列）；In 頁 lc_26 32% 為券商估計（Interested-party）", fmt="General")
    R["months"] = S.row("已公布月數（2026-01～10）", 10, None, "月", "輸入：月營收至 2026-10（2026-11-08 公告）", fmt="0", blue=True)
    S.skip()
    S.bold("4b. 層級隱含值（Siliconomics 現行版；連到本頁第 2 節層級值列）")
    S.hdr(["項目", "CY2026", "", "", "", "單位", "口徑、來源、證據等級"])
    R["lmkt"] = S.row("AI 液冷市場（層級值）", f"=B{G['mkt'][1]}", None, "US$B", "Derived：本頁第 2 節 Siliconomics SIF_LiqMkt_Base（CY2026）")
    R["lsh"] = S.row("奇鋐有效份額（層級值）", f"=B{G['sh'][1]}", None, "%", "Derived（本公司校準值，Siliconomics 快照）：lc_avsh_26，由券商液冷占比 32% × 營收錨點校準", fmt="0.0%")
    R["lliq"] = S.row("層級隱含：奇鋐液冷營收（市場 × 份額 × 匯率）", f"=B{R['lmkt']}*B{R['lsh']}*fx", None, "NT$B", "Derived：匯率 In 頁 fx", fmt="#,##0.0")
    S.skip()
    R["tbl"] = S.bold("C0（實績）：公司實績 vs 層級隱含值（HTML 用）")
    R["hdr"] = S.hdr(["項目", "主讀數：2026-01～10 累計 vs 層級 CY2026 × 10 ÷ 12", "第二讀數：3Q26 × 4 vs 層級 CY2026", "", "", "單位", "口徑、來源、證據等級"])
    fr = f"B{R['months']}/12"
    pct = "+0.0%;-0.0%"
    a = S.row("液冷：公司實績", "無（未揭露液冷營收）", "無（未揭露液冷營收）", "NT$B", "公司未揭露液冷營收金額或占比；不以模型值或券商估計補位", fmt="General")
    l = S.row("液冷：層級隱含", f"=B{R['lliq']}*{fr}", f"=B{R['lliq']}", "NT$B", "Derived：層級隱含液冷營收（4b）；主讀數 × 10 ÷ 12", fmt="#,##0.0")
    R["j_lc"] = S.row("液冷：判定", "無可比實績（層級值同源）", "無可比實績（層級值同源）", "", "缺液冷營收揭露；層級份額由本公司校準，市場大小未受檢驗；不計入對帳中", fmt="General")
    a2 = S.row("液冷（參考，同源校準）：券商液冷占比 32% × 公司營收實績", f"=B{R['cum10']}*lc_26", f"=B{R['q3']}*lc_26*4", "NT$B", "Verified 營收 × 券商估計占比（In 頁 lc_26，Interested-party）；不是公司揭露", fmt="#,##0.0")
    g2 = S.row("液冷（參考，同源校準）：差距（÷ 層級隱含 − 1）", f"=B{a2}/B{l}-1", f"=C{a2}/C{l}-1", fmt=pct)
    R["gap_ref"] = g2
    R["j_ref"] = S.row("液冷（參考，同源校準）：判定", "參考（同源校準）：不判定", "參考（同源校準）：不判定", "", "兩讀數代數上等於實績營收 ÷ 營收錨點 − 1（32% 與市場大小約掉），只檢驗本模型營收錨點", fmt="General")
    a3 = S.row("上限檢查：公司揭露伺服器與網通營收（前三季＝占比 × 累計 9 月）", f"=B{R['srv9']}*B{R['cum9']}", None, "NT$B", "Verified（公司占比約數 × 月營收累計）：液冷屬伺服器與網通事業", fmt="#,##0.0")
    R["ub"] = S.row("上限檢查：層級隱含液冷（CY2026 × 9 ÷ 12）÷ 伺服器與網通（前三季；應 ≤ 100%）", f"=B{R['lliq']}*9/12/B{a3}", None, "%", "Derived：層級液冷不得超過其所屬事業；通過＝合理，不是一致性判定", fmt="0.0%")
    R["j_ub"] = S.row("上限檢查：結果", f'=IF(B{R["ub"]}<=1,"通過","不通過")', None, "", "", fmt="General")
    S.skip()
    S.bold("參考：本模型 FY2026 營收錨點對月營收實績（模型對實績，不是層級對帳；不計入 C0）")
    R["anc"] = S.row("FY2026 營收錨點（Q 頁）", f"=Q!F{lab(Q, '營收')}", None, "NT$B", "Derived／Interested-party：1–3Q26 財報實績＋4Q26 假設（季增 15%）", fmt="#,##0.0")
    R["imp"] = S.row("錨點隱含 11–12 月月均營收（錨點 − 1～10 月累計）÷ 2", f"=(B{R['anc']}-B{R['cum10']})/2", None, "NT$B", "Derived", fmt="#,##0.0")
    R["impg"] = S.row("隱含月均 ÷ 10 月實績 − 1", f"=B{R['imp']}/B{R['oct']}-1", None, "%", "Derived：正值＝錨點假設 11–12 月高於 10 月水準", fmt=pct)
    S.skip()
    R["sum0"] = S.bold("C0（實績）結果摘要（HTML 用；公式帶數字）")
    nopen = f'COUNTIF(B{R["j_lc"]}:C{R["j_ref"]},"對帳中")'
    R["sum1"] = S.text(F(T("C0（實績）結論：無可比實績（層級值同源）：奇鋐未揭露液冷營收；Siliconomics 的奇鋐份額由本公司以券商液冷占比校準。對帳中項目 "), nopen,
                         T(" 項。參考：以券商液冷占比換算，差距主讀數 "), f'TEXT(B{g2},"+0.0%;-0.0%")', T("、第二讀數 "), f'TEXT(C{g2},"+0.0%;-0.0%")',
                         T("，只反映營收錨點與實績的差異，不判定。")))
    R["sum2"] = S.text(F(T("上限檢查：層級隱含液冷（CY2026 前三季比例）占公司揭露伺服器與網通營收 "), f'TEXT(B{R["ub"]},"0%")', T("，"), f'B{R["j_ub"]}',
                         T("；10 月營收 "), f'TEXT(B{R["oct"]}*10,"#,##0")', T(" 億元，營收錨點隱含 11–12 月月均 "), f'TEXT(B{R["imp"]}*10,"#,##0")', T(" 億元（較 10 月 "),
                         f'TEXT(B{R["impg"]},"+0%;-0%")', T("）。")))
    R["sum3"] = S.text("限制與待查：公司未揭露液冷營收，Siliconomics 的奇鋐份額由本模型校準，故液冷市場大小未獲獨立驗證；3Q26 財報已滾入（v0.3.6；毛利率 33.86%），4Q26 毛利率假設 33% 未改；3Q26 法說簡報原檔未取得。")
    R["end"] = S.r - 1
    return R


def apply(wb, SIL, CR):
    C.update(layer_rows=layer_rows, same_blocks=same_blocks)
    out = c13.apply(wb, SIL, CR, C, lines, c0_actuals)
    # 獨立核對（2026-11-15）：v1.3 各段（標準損益表、RevDCF、Val 0.7、Sens 末表、README 展示層）的實際起始版為 v0.3.5，不隨 VER 改寫
    import re as _re
    _pat = _re.compile(r"(（|；)" + _re.escape(VER) + r" 起(；個股 HTML|，個股 HTML|，版面規格)")
    for _ws in wb:
        for _row in _ws.iter_rows():
            for _c in _row:
                if isinstance(_c.value, str) and VER + " 起" in _c.value:
                    _c.value = _pat.sub(lambda m: m.group(1) + "v0.3.5 起" + m.group(2), _c.value)
    return out
