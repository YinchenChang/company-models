# 健策 v0.2.5（2026-11-15）後處理：個股 HTML 版面規格 v1.3＋三層協調兩讀數＋C0（實績）（工作單 docs/workorders/20261114_第二批13家_v1.3與三層協調.md；
# 規則 docs/plan/三層基準協調規則.md v1.4.2；樣板 台達電 delta/v056_post.py、v057_post.py，同次產業奇鋐、雙鴻、光寶科）。共通步驟在 jentech/c13.py（電源散熱四家同一份），本檔只放健策的設定。
# build_jentech.py 在 Coord 頁建好之後、存檔前呼叫 apply(wb, SIL, coord_rows)。alt_run.py 建置（無 Coord 頁）不呼叫。只新增，不改既有數值格。
import os, sys
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import c13
from c13 import T, F, lab, blockB, Sheet

VER, BUILT = "v0.2.6", "2026-11-15"
AGM24 = "https://jentech.com.tw/_files/ugd/e8883b_d59bb0c172834f49962b5206aa2ce044.pdf"
SA = "https://stockanalysis.com/quote/tpe/3653/financials/"
HI = "https://histock.tw/stock/3653/每股盈餘"
M_OCT = "https://finance.technews.tw/2026/11/06/jentech-3653-202610-financial-report/"
M_SEP = "https://finance.technews.tw/2026/10/08/jentech-3653-202609-financial-report/"

C = dict(
    name="健策", ver=VER, built=BUILT, since="v0.2.5",   # v1.3 各段標題「起始版」（v0.2.5 起；c13 選用參數）
    ext_rows=lambda P: (lambda b: dict(rev=lab(P, "營收", b, exact=True, end=b + 50), ni=lab(P, "歸屬母公司淨利", b, exact=True, end=b + 50),
                                         sh=lab(P, "股數", b, exact=True, end=b + 50), eps=lab(P, "EPS", b, exact=True, end=b + 50), hdr=b + 1))(blockB(P)),   # 延伸期在推導區塊 H–J 欄
    hist_title=("Hist — 損益表歷史（v0.2.5 起；PL 頁「標準損益表」的歷史三年與年增率用）。金額：十億新台幣；EPS：新台幣；歸屬母公司。"
                f"來源：FY2022–FY2023＝公司 113 年股東常會議事錄附 112 年度合併財報（仟元；{AGM24}）；FY2024＝stockanalysis（百萬元；{SA}）、稅前利益由其稅前利益率推得、EPS 取 HiStock 各季加總（{HI}）；"
                "FY2025＝Q 頁（公司財報，經 HiStock、Goodinfo）。FY2024–FY2025 所得稅與非控制權益未拆分（所得稅列＝稅前 − 歸屬母公司淨利，非控制權益列記 0）"),
    hist=[
        ("rev", "營收", 12.032, 12.063, 14.278, 20.276, "Verified", f"FY2022–23：{AGM24}（營業收入 12,031,894／12,062,664 仟元）；FY2024：{SA}；FY2025 同 Q 頁"),
        ("gp", "毛利", 4.326, 4.042, 5.365, 8.434, "Verified", "同上（營業毛利 4,325,842／4,042,419 仟元）"),
        ("oi", "營業利益", 3.034, 2.677, 3.799, 6.498, "Verified", "同上（營業利益 3,034,060／2,676,962 仟元）"),
        ("pti", "稅前利益", 3.369, 2.786, 4.245, 6.607, "Verified／Derived", f"FY2022–23：議事錄（稅前淨利 3,368,719／2,785,545 仟元）；FY2024：stockanalysis 稅前利益率 29.73% × 營收（Derived）；FY2025：Q 頁（HiStock 各季）"),
        ("tax", "所得稅（FY2024–FY2025 含非控制權益）", 0.712, 0.534, 0.817, 1.330, "Verified／Derived", "FY2022–23：議事錄（所得稅費用 712,427／533,839 仟元）；FY2024–25：稅前 − 歸屬母公司淨利（含非控制權益，未拆分）"),
        ("nci", "非控制權益損益（FY2024–FY2025 未拆分，記 0）", 0.042, -0.046, 0, 0, "Verified／未拆分", "FY2022–23：議事錄（非控制權益 42,011／(46,488) 仟元）；FY2024–25：未找到拆分，記 0"),
        ("ni", "稅後淨利（歸屬母公司）", 2.614, 2.298, 3.428, 5.277, "Verified", "FY2022–23：議事錄（2,614,281／2,298,194 仟元）；FY2024：stockanalysis Net Income；FY2025 同 Q 頁"),
        ("sh", "EPS 換算股數（十億股）", 0.134, 0.139, 0.142, 0.144, "Derived", "＝歸屬母公司淨利 ÷ 基本 EPS（四捨五入到百萬股）"),
        ("eps", "EPS（公司公告，基本，NT$）", 19.54, 16.59, 24.15, 36.75, "Verified", f"FY2022–23：議事錄基本每股盈餘；FY2024：HiStock 各季 EPS 加總 24.15（{HI}）；FY2025：公司 36.75（Q 頁；HiStock 加總 36.76）"),
    ],
    hist_chk2="FY2025：本表營收、毛利、營業利益、淨利 − Q 頁錨點（應為 0）",
    fy26_short="FY2026E 含 1–3Q26 財報實績",
    fy26_long="1–3Q26 財報實績（公開資訊觀測站 t164sb04）＋4Q26 假設",
    eps_note="歷史為公司公告基本 EPS（FY2024 為各季加總）；預測為歸屬母公司淨利 ÷ 股數（同基本口徑）",
    gw_label=None,
    same_names="AI 液冷市場與健策液冷冷板元件份額兩項",
    a1_tail=("v0.2.5（2026-11-15）：個股 HTML 版面規格 v1.3（標準損益表、反向 DCF 頁、四個目標價、加權目標價敏感度主圖）＋三層協調（Coord 頁兩讀數目標價、讀取層級、C0（實績））；"
             "A1 標題版號由舊的 v0.1.2 更正；既有數值不變；v0.2.6（2026-11-15）：滾入 3Q26 財報（公開資訊觀測站）：FY2026 錨點改為 1–3Q26 財報實績＋4Q26 假設（4Q 假設不改）；DCF 資產負債表日 2026-06-30 → 09-30（已發放股利不再扣）；4Q26 與評價用股數改 9-30 股本 ÷ 面額"),
    ver_line=("v0.2.6（2026-11-15）：滾入 3Q26 財報（公開資訊觀測站）：FY2026 錨點改為 1–3Q26 財報實績＋4Q26 假設（4Q 假設不改）；DCF 資產負債表日 2026-06-30 → 09-30（已發放股利不再扣）；4Q26 與評價用股數改 9-30 股本 ÷ 面額；C0 第二讀數改用 3Q26 財報營收。v0.2.5（2026-11-15）：改為個股 HTML 版面規格 v1.3（樣板 NVIDIA、台達電、奇鋐）並補三層協調：Coord 頁並列兩讀數目標價（tools/coord）、讀取層級（同源，不判定）、"
              "C0（實績）（月營收 2026-01～10；結論：無可比實績，Siliconomics 沒有均熱片同口徑量、公司未揭露 AI 均熱片與液冷營收）；新增 RevDCF、Hist 頁、PL 標準損益表、Val 0.7、Sens 末表；"
              "既有數值、目標價與評等不變"),
)


def layer_rows(ver, CR):
    return [
        ("健策液冷冷板元件份額 CY2027／28／29（÷ 液冷市場；lc_jtsh_27–29）", '=TEXT(lcs_27,"0.0%")&"／"&TEXT(lcs_28,"0.0%")&"／"&TEXT(lcs_29,"0.0%")', ver, "液冷冷板元件營收＝液冷市場 × 份額（新產品，Assumed）", None),
        ("AI 液冷市場 CY2026（US$B；SIF_LiqMkt_Base，經奇鋐成品讀取並核對）", f"=PL!D{0}", ver, "液冷冷板元件的市場（建檔時核對與 Siliconomics 基準一致）", "0.00"),
        ("匯率（NT$/US$；fx_twd）", "=fx", ver, "市場換算新台幣", "0.00"),
        ("情境 B 顆數係數 CY2028／CY2029（÷ 原始顆數；SIF_UnitsCoef_Base）", '=TEXT(um28_B,"0.000")&"／"&TEXT(um29_B,"0.000")', f"{ver}，經 tsmc/dist 成品 In 頁 um28_B、um29_B 讀取", "加速器顆數的情境係數；A（原始顆數）、C（0.6）為個股自訂情境", None),
        ("加速器顆數（Blackwell、Rubin、AMD、ASIC）", "markets/network 成品", "markets/network（市場層；非 Siliconomics）", "均熱片市場＝顆數 × 每顆金額（每顆金額為本公司變數，Assumed／Interested-party）", None),
    ]


def same_blocks(SIL, pc):
    N, V = SIL.NAMES, SIL.V
    js = [0] + [V[f"lc_jtsh_{y}"]["base"] for y in (27, 28, 29)]
    return [
        ("mkt", "AI 液冷市場（US$B）", "US$B", "AI 液冷市場（奇鋐模型；美元）", list(N["SIF_LiqMkt_Base"]), "讀自奇鋐成品（建檔時核對＝Siliconomics 基準）", f"Derived：Siliconomics {SIL.VERSION} SIF_LiqMkt_Base", "0.00"),
        ("sh", "健策液冷冷板元件份額", "%", "健策液冷份額", js, "讀自 Siliconomics（In 頁 lcs_27–29；CY2026 無液冷營收）", f"Assumed：Siliconomics {SIL.VERSION} lc_jtsh_27–29（CY2026 無此值，記 0）", "0.0%"),
    ]


def lines(wb):
    P = wb["PL"]; SE = wb["Sens"]
    rb = blockB(P)
    pb = lambda label, c: f"PL!{c}{lab(P, label, rb, exact=True, end=rb + 50)}"
    sl = lambda k: f"Sens!L{lab(SE, k)}"
    return dict(
        hero=lambda tp, up, rec, dcf, cadj: (T("加權目標價 "), f'TEXT({tp},"#,##0")', T(" 元，較現價"), f'IF({up}<0,"低 ","高 ")', f'TEXT(ABS({up}),"0.0%")', T("，評等"), rec,
                                             T("。現金流折現 "), f'TEXT({dcf},"#,##0")', T(" 元、同業比較法（相對倍數調整後）"), f'TEXT({cadj},"#,##0")',
                                             T(" 元各半；AI 均熱片隨 Rubin 單價上升，目標價對 AI 部署量與可拆式上蓋單價最敏感。")),
        premise="加速器顆數依網通市場模型與 Siliconomics 顆數係數、每顆均熱片金額與健策份額（假設）成立",
        o2=(T("AI 均熱片是成長主力：情境 B AI 加速器均熱片營收由 FY2026 約 "), f'TEXT({pb("AI 加速器均熱片營收", "D")}*10,"#,##0")', T(" 億元增至 FY2028 約 "),
            f'TEXT({pb("AI 加速器均熱片營收", "F")}*10,"#,##0")', T(" 億元，占營收由 "), f'TEXT({pb("AI 加速器均熱片 ÷ 營收", "D")},"0%")', T(" 升到 "),
            f'TEXT({pb("AI 加速器均熱片 ÷ 營收", "F")},"0%")', T("；每顆加速器健策均熱片營收由 "), f'TEXT({pb("每顆加速器健策 AI 均熱片營收（全部加速器平均）", "D")},"0")',
            T(" 美元升到 "), f'TEXT({pb("每顆加速器健策 AI 均熱片營收（全部加速器平均）", "F")},"0")', T(" 美元。")),
        o3=(T("最大的變數是 AI 部署量與可拆式上蓋單價：AI 部署 ×0.8 使情境 B 的 FY2028 EPS "), f'TEXT({sl("AI 部署 ×0.8")},"+0.0%;-0.0%")', T("，可拆式上蓋慢 "),
            f'TEXT({sl("可拆式上蓋 慢")},"+0.0%;-0.0%")', T("，份額流失 "), f'TEXT({sl("份額流失（低）")},"+0.0%;-0.0%")', T("；基準情境 B 的 FY2028 EPS 為 "),
            f'TEXT({pb("EPS", "F")},"0.0")', T(" 元。")),
        o5_tail="本模型以加速器顆數推導，不讀部署 GW；C0（實績）無可比實績（Siliconomics 沒有均熱片同口徑量，公司未揭露 AI 均熱片與液冷營收）。",
        drv=(T("分部營收推導（情境 B）：AI 加速器均熱片營收＝各世代顆數 × 每顆金額 × 健策份額；液冷冷板元件＝液冷市場 × 健策份額；其他散熱、導線架與其他各依年增率；"
               "相加為合併營收。FY2028 AI 均熱片 "), f'TEXT({pb("AI 加速器均熱片營收", "F")},"#,##0.0")', T("、液冷冷板元件 "), f'TEXT({pb("液冷冷板元件營收（MCL、MLCP）", "F")},"#,##0.0")',
             T("、其他散熱 "), f'TEXT({pb("其他散熱營收（CPU 均熱片、交換器加強環、CPO、ILM、車用水冷）", "F")},"#,##0.0")', T("、導線架與其他 "),
             f'TEXT({pb("導線架、精密沖壓與其他營收", "F")},"#,##0.0")', T("，合計 "), f'TEXT({pb("營收", "F")},"#,##0.0")', T(" 十億元。")),
    )


def c0_actuals(wb, G):
    CW = wb["Coord"]; Q = wb["Q"]; P = wb["PL"]
    rb = blockB(P)
    S = Sheet(CW, G["end"] + 3)
    R = {}
    R["title"] = S.bold("4. C0（實績）：公司實績 vs Siliconomics 層級隱含值（規則 v1.4.2 C0 第 2 點；工作單 2026-11-14；差距＝實績 ÷ 層級隱含值 − 1）")
    R["method"] = S.text("方法：個股側＝健策已公布實績（月營收、3Q26 財報）；層級側＝Siliconomics 現行版同口徑量。主讀數＝2026-01～10 累計實績 vs 層級 CY2026 × 10 ÷ 12；第二讀數＝3Q26 財報營收 × 4 vs 層級 CY2026"
                         "（約等於 3Q26 中點 8 月中的年化速率）。判定：兩讀數差距都超過 ±15% 且同號 → 對帳中；只一個超過 → 時差邊界（參考）；都在內 → 一致。"
                         "健策主要營收為 AI 加速器均熱片（上蓋），Siliconomics 沒有均熱片市場、每顆金額或健策份額（顆數取自網通市場模型，每顆金額為本公司變數）；"
                         "液冷冷板元件份額在 Siliconomics 只有 CY2027 起的假設值（CY2026 無），公司也未揭露液冷營收。故兩列皆無可比實績。")
    R["limit"] = S.text("缺哪一項揭露：(1) 公司 AI 加速器均熱片（含可拆式上蓋）營收金額或占比（本模型 FY2026 40% 為假設）；(2) 液冷冷板元件（MCL、MLCP）營收；(3) Siliconomics 若新增均熱片市場或每顆金額的層級值，才有同口徑的層級側。"
                        "營收實績只對本模型營收錨點作參考，不是層級對帳。")
    S.skip()
    S.bold("4a. 公司已公布實績與揭露（輸入；藍字＝來自原文的常數）")
    S.hdr(["項目", "值", "", "", "", "單位", "口徑、來源、證據等級", "來源網址"])
    R["sep"] = S.row("月營收：2026-09", 3.054, None, "NT$B", "Verified：月營收公告 30.54 億元，月增 −5.3%、年增 80.0%，累計前 9 月 220.73 億元（TechNews 2026-10-08）", blue=True, url=M_SEP)
    R["oct"] = S.row("月營收：2026-10", 3.073, None, "NT$B", "Verified：月營收公告 30.73 億元，月增 0.6%、年增 80.7%（TechNews 2026-11-06）", blue=True, url=M_OCT)
    R["cum9"] = S.row("累計：2026-01～09", 22.073, None, "NT$B", "Verified：月營收公告 220.73 億元，年增 47.1%（TechNews 2026-10-08）", blue=True, url=M_SEP)
    R["cum10"] = S.row("累計：2026-01～10（主讀數的實績基礎）", 25.146, None, "NT$B", "Verified：月營收公告 251.46 億元，年增 50.5%（TechNews 2026-11-06）", blue=True, url=M_OCT)
    R["chk1"] = S.row("檢查：累計 9 月 + 10 月 − 累計 10 月（應為 0）", f"=B{R['cum9']}+B{R['oct']}-B{R['cum10']}", None, "NT$B", "Derived")
    R["q3"] = S.row("3Q26 營收（最近一季；財報）", "=q3_rev", None, "NT$B", "Verified：In 頁 q3_rev＝公開資訊觀測站 t164sb04 3Q26 營業收入 9,493,226 千元（7–9 月月營收合計 94.93 億元同值）")
    R["chk2"] = S.row("檢查：累計 9 月（月營收）− 1–3Q26 財報營收（應接近 0）", f"=B{R['cum9']}-ytd_rev", None, "NT$B", "Derived：月營收 220.73 億元 vs 公開資訊觀測站 1–9 月累計 220.73 億元（Q 頁 1H26＝累計 − 3Q26）")
    R["eps3"] = S.row("3Q26 EPS（財報；已滾入錨點）", 19.70, None, "NT$", "Verified：公開資訊觀測站 t164sb04 基本每股盈餘 19.70 元（HiStock 同；Q1 9.53、Q2 15.8）；v0.2.5 Q 頁 3Q26 模型 21.54", fmt="0.00", blue=True, url="https://mops.twse.com.tw/mops/#/web/t164sb04")
    R["dis"] = S.row("公司揭露：AI 均熱片、液冷營收或占比", "未揭露", None, "", "公司未揭露 AI 加速器均熱片或液冷冷板元件營收金額或占比（優分析 2026-09-21：公司未給散熱、AI 均熱片、液冷比重）", fmt="General")
    R["months"] = S.row("已公布月數（2026-01～10）", 10, None, "月", "輸入：月營收至 2026-10（2026-11-06 公告）", fmt="0", blue=True)
    S.skip()
    S.bold("4b. 層級隱含值（Siliconomics 現行版）")
    S.hdr(["項目", "CY2026", "", "", "", "單位", "口徑、來源、證據等級"])
    R["lhs"] = S.row("AI 加速器均熱片：Siliconomics 層級值", "無（Siliconomics 無均熱片市場、每顆金額與健策份額）", None, "", "本模型均熱片市場＝網通市場模型顆數 × 本公司每顆金額（Assumed／Interested-party），不是層級值", fmt="General")
    R["lmkt"] = S.row("AI 液冷市場（層級值）", f"=B{G['mkt'][1]}", None, "US$B", "Derived：本頁第 2 節 Siliconomics SIF_LiqMkt_Base（CY2026）")
    R["lsh"] = S.row("健策液冷冷板元件份額（層級值，CY2026）", f"=B{G['sh'][1]}", None, "%", "Assumed：lc_jtsh 只從 CY2027 起（CY2026 無，記 0）", fmt="0.0%")
    R["lliq"] = S.row("層級隱含：健策液冷冷板元件營收（CY2026）", f"=B{R['lmkt']}*B{R['lsh']}*fx", None, "NT$B", "Derived：市場 × 份額 × 匯率；CY2026 為 0", fmt="#,##0.0")
    S.skip()
    R["tbl"] = S.bold("C0（實績）：公司實績 vs 層級隱含值（HTML 用）")
    R["hdr"] = S.hdr(["項目", "主讀數：2026-01～10 累計 vs 層級 CY2026 × 10 ÷ 12", "第二讀數：3Q26 × 4 vs 層級 CY2026", "", "", "單位", "口徑、來源、證據等級"])
    S.row("AI 加速器均熱片：公司實績", "無（未揭露）", "無（未揭露）", "NT$B", "公司未揭露 AI 均熱片營收金額或占比", fmt="General")
    S.row("AI 加速器均熱片：層級隱含", "無（層級無此量）", "無（層級無此量）", "NT$B", "Siliconomics 無均熱片同口徑量", fmt="General")
    R["j_hs"] = S.row("AI 加速器均熱片：判定", "無可比實績", "無可比實績", "", "缺公司均熱片營收揭露，且 Siliconomics 沒有同口徑量；不計入對帳中", fmt="General")
    S.row("液冷冷板元件：公司實績", "無（未揭露）", "無（未揭露）", "NT$B", "公司未揭露液冷冷板元件營收", fmt="General")
    S.row("液冷冷板元件：層級隱含", f"=B{R['lliq']}*B{R['months']}/12", f"=B{R['lliq']}", "NT$B", "Derived：CY2026 層級份額為 0（Siliconomics 只有 CY2027 起的假設值）", fmt="#,##0.0")
    R["j_lc"] = S.row("液冷冷板元件：判定", "無可比實績", "無可比實績", "", "缺液冷營收揭露；CY2026 層級值為 0（新產品）；不計入對帳中", fmt="General")
    S.skip()
    S.bold("參考：本模型 FY2026 營收錨點對月營收實績（模型對實績，不是層級對帳；不計入 C0）")
    R["anc"] = S.row("FY2026 營收錨點（Q 頁）", f"=Q!F{lab(Q, '營收')}", None, "NT$B", "Derived／Interested-party：1–3Q26 財報實績＋4Q 假設（季增 10%）", fmt="#,##0.0")
    R["imp"] = S.row("錨點隱含 11–12 月月均營收（錨點 − 1～10 月累計）÷ 2", f"=(B{R['anc']}-B{R['cum10']})/2", None, "NT$B", "Derived", fmt="#,##0.00")
    R["impg"] = S.row("隱含月均 ÷ 10 月實績 − 1", f"=B{R['imp']}/B{R['oct']}-1", None, "%", "Derived：正值＝錨點假設 11–12 月高於 10 月水準", fmt="+0.0%;-0.0%")
    S.skip()
    R["sum0"] = S.bold("C0（實績）結果摘要（HTML 用；公式帶數字）")
    nopen = f'COUNTIF(B{R["j_hs"]}:C{R["j_lc"]},"對帳中")'
    R["sum1"] = S.text(F(T("C0（實績）結論：無可比實績：Siliconomics 沒有均熱片同口徑量（市場、每顆金額、健策份額），健策也未揭露 AI 均熱片與液冷冷板元件營收；液冷份額在層級只從 CY2027 起。對帳中項目 "), nopen, T(" 項。")))
    R["sum2"] = S.text(F(T("參考（模型對實績）：1～10 月營收 "), f'TEXT(B{R["cum10"]}*10,"#,##0")', T(" 億元，10 月 "), f'TEXT(B{R["oct"]}*10,"#,##0.0")', T(" 億元；營收錨點隱含 11–12 月月均 "),
                         f'TEXT(B{R["imp"]}*10,"#,##0.0")', T(" 億元（較 10 月 "), f'TEXT(B{R["impg"]},"+0%;-0%")', T("）；3Q26 EPS "), f'TEXT(B{R["eps3"]},"0.00")', T(" 元。")))
    R["sum3"] = S.text("限制與待查：均熱片屬硬體供應鏈個股的行業變數，若 Siliconomics 新增均熱片市場（顆數 × 每顆金額）與份額，才能以公司揭露的散熱或均熱片營收做 C0；3Q26 財報已滾入（v0.2.6；第二讀數改用財報營收），3Q26 法說未取得。")
    R["end"] = S.r - 1
    return R


def apply(wb, SIL, CR):
    P = wb["PL"]
    rb = blockB(P)
    r_mkt = lab(P, "AI 液冷市場（奇鋐模型；美元）", rb, exact=True, end=rb + 50)
    lr = layer_rows
    C.update(layer_rows=lambda ver, CR_: [(a, (f"=PL!D{r_mkt}" if b == "=PL!D0" else b), c, d, e) for a, b, c, d, e in lr(ver, CR_)], same_blocks=same_blocks)
    return c13.apply(wb, SIL, CR, C, lines, c0_actuals)
