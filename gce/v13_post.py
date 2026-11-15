# 金像電 v0.3.4（2026-11-15）後處理：個股 HTML 版面規格 v1.3＋三層協調兩讀數＋C0（實績）（工作單 docs/workorders/20261114_第二批13家_v1.3與三層協調.md；
# 規則 docs/plan/三層基準協調規則.md v1.4.2；樣板 NVIDIA v0.3.10、台達電、同次產業台光電 emc/v13_post.py）。共通步驟在 gce/c13.py（高速 CCL／PCB 四家同一份），本檔只放金像電的設定。
# build_gce.py 在 Coord 頁建好之後、存檔前呼叫 apply(wb, SIL, coord_rows)。alt_run.py 建置（無 Coord 頁）不呼叫。只新增，不改既有數值格。
# 金像電的推導區塊直接延伸到 FY2032（無「延伸｜情境 B」表）：ext_rows 指向情境 B 區塊本身的營收、淨利、股數、EPS 列（H、I 欄＝FY2030、FY2031）。
import os, sys
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import c13
from c13 import T, F, lab, blockB, Sheet

VER, BUILT = "v0.3.5", "2026-11-15"
SA = "https://stockanalysis.com/quote/tpe/2368/financials/income-statement/"
M_OCT = "https://finance.technews.tw/2026/11/06/gce-2368-202610-financial-report/"
M_SEP = "https://finance.technews.tw/2026/10/08/gce-2368-202609-financial-report/"
AIR = "AI 相關營收 ＝ MIN（需求面、產能上限 − 非 AI）"
VCMP, VSW, VMID = "每 GW PCB：加速器卡／運算板", "每 GW PCB：交換器板", "每 GW PCB：中板／背板／其他 × 金像電參與係數"


def ext_rows(P):
    rb = blockB(P)
    g = lambda label: lab(P, label, rb, exact=True, end=rb + 70)
    return dict(rev=g("營收"), ni=g("歸屬母公司淨利"), sh=g("股數"), eps=g("EPS"), hdr=rb + 1)


C = dict(
    name="金像電", ver=VER, built=BUILT, nop_label="營業外淨額", ext_rows=ext_rows,
    hist_title=("Hist — 損益表歷史（v0.3.4 起；PL 頁「標準損益表」的歷史三年與年增率用）。金額：十億新台幣；EPS：新台幣；歸屬母公司。"
                f"來源：stockanalysis 年度損益表（百萬元，S&P Global，2026-11-15 讀取；{SA}），FY2025 與 Q 頁一致。營業外收益＝稅前 − 營業利益；營業費用＝毛利 − 營業利益；"
                "stockanalysis 淨利與歸屬母公司淨利相同（無非控制權益），非控制權益損益列 0"),
    hist=[
        ("rev", "營收", 32.785, 30.044, 38.952, 60.004, "Verified", f"{SA}：Revenue（百萬元）；FY2025 同 Q 頁"),
        ("gp", "毛利", 8.728, 7.724, 11.396, 19.683, "Verified", "同上：Gross Profit"),
        ("oi", "營業利益", 6.037, 5.136, 8.067, 14.043, "Verified", "同上：Operating Income"),
        ("pti", "稅前利益", 6.388, 5.218, 8.490, 14.173, "Verified", "同上：Pretax Income"),
        ("tax", "所得稅", 1.820, 1.689, 2.875, 4.567, "Verified", "同上：Income Tax Expense"),
        ("nci", "非控制權益損益", 0.0, 0.0, 0.0, 0.0, "Verified", "同上：Net Income 與 Net Income to Common 相同（無非控制權益）；核對列差額為百萬元四捨五入"),
        ("ni", "稅後淨利（歸屬母公司）", 4.568, 3.529, 5.616, 9.607, "Verified", "同上：Net Income to Common；FY2025 同 Q 頁"),
        ("sh", "EPS 換算股數（十億股）", 0.516, 0.487, 0.487, 0.493, "Verified", "同上：Shares Outstanding (Basic)（百萬股取整）"),
        ("eps", "EPS（公司公告，基本，NT$）", 8.86, 7.25, 11.54, 19.47, "Verified", "同上：EPS (Basic)（稀釋 8.78／7.22／11.25／19.00）；FY2025 取公司公告 19.47（Q 頁；stockanalysis 19.48 為四捨五入差）"),
    ],
    fy26_short="FY2026E 含 1–3Q26 財報實績",
    fy26_long="1–3Q26 財報實績（公開資訊觀測站 2026-11-10）＋4Q26 假設",
    eps_note="歷史為公司公告基本 EPS；預測為歸屬母公司淨利 ÷ 加權股數（含 Amazon 私募；未計可轉債稀釋，同基本口徑）",
    gw_label="部署 GW（IT 電力）",
    same_names="部署 GW、每 GW AI 高速 PCB 金額（三項全額）、AI 高速 PCB 市場三項",
    a1_tail=("v0.3.4（2026-11-15）：個股 HTML 版面規格 v1.3（標準損益表、反向 DCF 頁、四個目標價、加權目標價敏感度主圖）＋三層協調（Coord 頁兩讀數目標價、讀取層級、C0（實績））；"
             "A1 標題版號由舊的 v0.1.1 更正；既有數值不變；v0.3.5（2026-11-15）：滾入 3Q26 財報（公開資訊觀測站 2026-11-10）：FY2026 錨點改為 1–3Q26 財報實績＋4Q26 假設（4Q 假設不改）；DCF 資產負債表日 2026-06-30 → 09-30（可轉債轉換完畢、已發放股利不再扣）；4Q26 與評價用股數改 9-30 股本 ÷ 面額；C0 第二讀數改用 3Q26 財報營收"),
    ver_line=("v0.3.5（2026-11-15）：滾入 3Q26 財報（公開資訊觀測站 2026-11-10）：FY2026 錨點改為 1–3Q26 財報實績＋4Q26 假設（4Q 假設不改）；DCF 資產負債表日 2026-06-30 → 09-30（可轉債轉換完畢、已發放股利不再扣）；4Q26 與評價用股數改 9-30 股本 ÷ 面額；C0 第二讀數改用 3Q26 財報營收。v0.3.4（2026-11-15）：改為個股 HTML 版面規格 v1.3（樣板 NVIDIA、台達電、台光電）並補三層協調：Coord 頁並列兩讀數目標價（tools/coord）、讀取層級（同源，不判定）、"
              "C0（實績）（月營收 2026-01～10 對 Siliconomics AI 高速 PCB 市場 × 參與係數 × 金像電份額；結論：無可比實績，份額由本模型錨點校準）；新增 RevDCF、Hist 頁、PL 標準損益表、Val 0.7、Sens 末表；"
              "既有數值、目標價與評等不變"),
)


def layer_rows(ver, CR):
    return [
        ("每 GW PCB 2026–29：加速器卡／運算板（US$M/GW；勝宏模型＝Siliconomics pcb_cmp × pcb_cmpx）", '=TEXT(vg_cmp26,"0")&"／"&TEXT(vg_cmp27,"0")&"／"&TEXT(vg_cmp28,"0")&"／"&TEXT(vg_cmp29,"0")',
         f"經 victory_giant/dist 成品 Units 頁讀入（勝宏模型讀 {ver}）", "每 GW 金像電可參與金額的第一項（本模型 In 頁 vg_cmp26–29）", None),
        ("每 GW PCB 2026–29：交換器板（pcb_sw × pcb_swx）", '=TEXT(vg_sw26,"0")&"／"&TEXT(vg_sw27,"0")&"／"&TEXT(vg_sw28,"0")&"／"&TEXT(vg_sw29,"0")', "同上", "第二項（vg_sw26–29）", None),
        ("每 GW PCB 2026–29：中板／背板／其他（pcb_mid × pcb_midx）", '=TEXT(vg_mid26,"0")&"／"&TEXT(vg_mid27,"0")&"／"&TEXT(vg_mid28,"0")&"／"&TEXT(vg_mid29,"0")', "同上", "第三項 × 金像電參與係數 w_mid（本模型假設）", None),
        ("2026 年部署中 Rubin 世代占比（vr26）", "=vr26", ver, "混合 FY2026 每 GW IT 資本支出", "0%"),
        ("匯率（NT$/US$；fx_twd）", "=fx", ver, "可參與市場換算新台幣", "0.00"),
        ("情境 B 顆數係數 CY2028／CY2029（÷ 原始顆數；SIF_UnitsCoef_Base）", '=TEXT(um28_B,"0.000")&"／"&TEXT(um29_B,"0.000")', f"{ver}，經 tsmc/dist 成品 In 頁 um28_B、um29_B 讀取", "基準情境 B 的 AI IT 資本支出係數；A（原始顆數）、C（0.6）為個股自訂情境", None),
        ("AI IT 資本支出 CY2026–29（US$B；markets/network）", '=TEXT(itc_26,"#,##0")&"／"&TEXT(itc_27,"#,##0")&"／"&TEXT(itc_28,"#,##0")&"／"&TEXT(itc_29,"#,##0")', "markets/network 成品（市場層；Tokenomics 經橋接表進入）", "部署 GW 的分子（原始顆數口徑）", None),
        ("每 GW IT 資本支出（US$B；GB300／VR200；markets/chip_bridge）", '=TEXT(itpg_gb,"0.0")&"／"&TEXT(itpg_vr,"0.0")', "markets/chip_bridge 成品（Tokenomics IF_CapexIT，經橋接表唯讀）", "部署 GW 的分母", None),
        ("金像電參與係數、份額、價格指數、產能指數（不在 Siliconomics）", "本模型", "個股模型（Siliconomics M16 註明：價格指數與各公司份額留在公司模型）", "FY2026 有效份額＝AI 相關營收錨點 ÷ 本模型可參與市場（校準）", None),
    ]


def same_blocks(SIL, pc):
    N = SIL.NAMES
    return [
        ("gw", "部署 GW（IT 電力）", "GW", "部署 GW（IT 電力）", list(N["SIF_DeployGW_Base"]),
         "Derived：AI IT 資本支出（markets/network）× 情境 B 顆數係數 ÷ 每 GW IT 資本支出（bridge）", f"Derived：Siliconomics {SIL.VERSION} SIF_DeployGW_Base（依規則 v1.4 的基準）", "0.00"),
        ("pgw", "每 GW AI 高速 PCB 金額（三項全額）", "US$M/GW", PGW, list(N["SIF_PCBperGW"]),
         "Derived：三項合計（中板項還原為全額＝÷ 參與係數；每 GW 金額成長倍率基準 1）", f"Derived：Siliconomics {SIL.VERSION} SIF_PCBperGW（經勝宏模型讀入）", "#,##0"),
        ("mkt", "AI 高速 PCB 市場（US$B；全額）", "US$B", MKT, list(N["SIF_PCBMkt_Base"]), "Derived：部署 GW × 每 GW 三項全額", f"Derived：Siliconomics {SIL.VERSION} SIF_PCBMkt_Base", "0.00"),
    ]


PGW, MKT = [], []


def _prep(wb):
    P = wb["PL"]; rb = blockB(P)
    g = lambda label: lab(P, label, rb, exact=True, end=rb + 70)
    a, b, m, gw = g(VCMP), g(VSW), g(VMID), g("部署 GW（IT 電力）")
    PGW[:] = [f"=PL!{c}{a}+PL!{c}{b}+PL!{c}{m}/w_mid" for c in "DEFG"]
    MKT[:] = [f"=PL!{c}{gw}*(PL!{c}{a}+PL!{c}{b}+PL!{c}{m}/w_mid)/1000" for c in "DEFG"]


def lines(wb):
    P = wb["PL"]; SE = wb["Sens"]
    rb = blockB(P)
    pb = lambda label, c: f"PL!{c}{lab(P, label, rb, exact=True, end=rb + 70)}"
    sl = lambda k: f"Sens!L{lab(SE, k)}"
    return dict(
        hero=lambda tp, up, rec, dcf, cadj: (T("加權目標價 "), f'TEXT({tp},"#,##0")', T(" 元，較現價"), f'IF({up}<0,"低 ","高 ")', f'TEXT(ABS({up}),"0.0%")', T("，評等"), rec,
                                             T("。現金流折現 "), f'TEXT({dcf},"#,##0")', T(" 元、同業比較法（相對倍數調整後）"), f'TEXT({cadj},"#,##0")',
                                             T(" 元各半；AI 伺服器與交換器板營收受產能約束，目標價對擴產進度、PCB 價格與份額最敏感。")),
        premise="產業部署量與每 GW AI 高速 PCB 金額依 Siliconomics 基準（行業變數，經勝宏模型讀入）、金像電參與係數、份額相對值、價格指數與擴產（假設）成立",
        o2=(T("AI 伺服器與交換器板是成長主力：情境 B AI 相關營收由 FY2026 約 "), f'TEXT({pb(AIR, "D")}*10,"#,##0")', T(" 億元增至 FY2028 約 "), f'TEXT({pb(AIR, "F")}*10,"#,##0")',
            T(" 億元，占營收由 "), f'TEXT({pb("AI 相關 ÷ 營收", "D")},"0%")', T(" 升到 "), f'TEXT({pb("AI 相關 ÷ 營收", "F")},"0%")', T("；同期金像電有效份額由 "),
            f'TEXT({pb("金像電有效份額", "D")},"0.0%")', T(" 降到 "), f'TEXT({pb("金像電有效份額", "F")},"0.0%")', T("，營收另受產值產能上限約束。")),
        o3=(T("最大的變數是擴產與價格（情境 B，FY2028 EPS）：擴產延遲 "), f'TEXT({sl("擴產延遲")},"+0.0%;-0.0%")', T("，價格回落快 "),
            f'TEXT({sl("價格回落快")},"+0.0%;-0.0%")', T("，份額流失 "), f'TEXT({sl("份額流失（低）")},"+0.0%;-0.0%")', T("，AI 部署 ×0.8 為 "),
            f'TEXT({sl("AI 部署 ×0.8")},"+0.0%;-0.0%")', T("；基準情境 B 的 FY2028 EPS 為 "), f'TEXT({pb("EPS", "F")},"0.0")', T(" 元。")),
        o5_tail="本模型的部署 GW 同 Siliconomics 基準、每 GW PCB 金額經勝宏模型讀入；C0（實績）無可比實績（公司未揭露 AI 營收金額，份額由本模型錨點校準）。",
        drv=(T("分部營收推導（情境 B）：AI 相關營收＝MIN（部署 GW × 每 GW 可參與 PCB 金額 × 金像電有效份額 × 價格指數，產能上限 − 非 AI）；非 AI 依年增率；"
               "相加為合併營收。FY2028 AI 相關 "), f'TEXT({pb(AIR, "F")},"#,##0.0")', T("、非 AI "), f'TEXT({pb("非 AI 營收", "F")},"#,##0.0")',
             T("，合計 "), f'TEXT({pb("營收", "F")},"#,##0.0")', T(" 十億元。")),
    )


def c0_actuals(wb, G):
    CW = wb["Coord"]; Q = wb["Q"]; P = wb["PL"]
    rb = blockB(P)
    g = lambda label: lab(P, label, rb, exact=True, end=rb + 70)
    sh0 = lab(P, "金像電有效份額（FY2026 校準）", rb)
    mku = g("可參與 AI 高速 PCB 市場（美元；不含價格指數）")
    S = Sheet(CW, G["end"] + 3)
    R = {}
    R["title"] = S.bold("4. C0（實績）：公司實績 vs Siliconomics 層級隱含值（規則 v1.4.2 C0 第 2 點；工作單 2026-11-14；差距＝實績 ÷ 層級隱含值 − 1）")
    R["method"] = S.text("方法：個股側＝金像電已公布實績（月營收；3Q26 財報，2026-11-10 公告）；層級側＝Siliconomics 現行版 AI 高速 PCB 市場（CY2026，全額）× 金像電可參與比例 × 金像電 FY2026 有效份額 × 匯率（市場連到本頁第 2 節層級值列）。"
                         "層級值是日曆年，故主讀數＝2026-01～10 累計實績 vs 層級 CY2026 × 10 ÷ 12（線性時差近似，Derived）；第二讀數＝3Q26 財報營收 × 4 vs 層級 CY2026（約等於 3Q26 中點 8 月中的年化速率，"
                         "季中點晚於年中、成長年偏高）。判定：兩讀數差距都超過 ±15% 且同號 → 對帳中；只一個超過 → 時差邊界（參考）；都在內 → 一致。公司沒有揭露 AI 相關營收金額"
                         "（AI 占比為券商與研調估計），故 AI 相關列無可比實績；以占比換算者只列參考。")
    R["limit"] = S.text("限制（同源校準）：Siliconomics 沒有金像電份額（M16 註明份額留在公司模型）；本模型 FY2026 有效份額＝AI 相關營收錨點（FY2026 營收錨點 − 非 AI）÷ 本模型可參與市場，"
                        "而可參與市場＝Siliconomics 部署 GW × 每 GW PCB 金額（經勝宏模型讀入）× 參與係數。層級隱含 AI 相關營收＝市場 × 可參與比例 × 份額 × 匯率＝AI 相關營收錨點，市場大小與參與係數在比式中被約掉；"
                        "以錨點 AI 占比 × 實績營收對照時，兩讀數代數上等於「實績營收 ÷ 營收錨點 − 1」，實際只檢驗本模型營收錨點。結論：無可比實績（層級值同源）；缺金像電 AI 相關營收金額揭露。")
    S.skip()
    S.bold("4a. 公司已公布實績與揭露（輸入；藍字＝來自原文的常數）")
    S.hdr(["項目", "值", "", "", "", "單位", "口徑、來源、證據等級", "來源網址"])
    R["sep"] = S.row("月營收：2026-09", 9.320, None, "NT$B", "Verified：月營收公告 93.20 億元，月增 −10.2%、年增 55.5%，累計前 9 月 733.79 億元（TechNews 2026-10-08）", blue=True, url=M_SEP)
    R["oct"] = S.row("月營收：2026-10", 9.913, None, "NT$B", "Verified：月營收公告 99.13 億元，月增 6.4%、年增 76.2%（TechNews 2026-11-06）", blue=True, url=M_OCT)
    R["cum9"] = S.row("累計：2026-01～09", 73.379, None, "NT$B", "Verified：月營收公告 733.79 億元，年增 69.1%（TechNews 2026-10-08）", blue=True, url=M_SEP)
    R["cum10"] = S.row("累計：2026-01～10（主讀數的實績基礎）", 83.292, None, "NT$B", "Verified：月營收公告 832.92 億元，年增 69.9%（TechNews 2026-11-06）", blue=True, url=M_OCT)
    R["chk1"] = S.row("檢查：累計 9 月 + 10 月 − 累計 10 月（應為 0）", f"=B{R['cum9']}+B{R['oct']}-B{R['cum10']}", None, "NT$B", "Derived")
    R["q3"] = S.row("3Q26 營收（最近一季；第二讀數的實績基礎；財報）", "=q3_rev", None, "NT$B", "Verified：In 頁 q3_rev＝公開資訊觀測站 t164sb04 3Q26 營業收入 29,567,267 千元（鉅亨 2026-11-11 同值）")
    R["chk2"] = S.row("檢查：累計 9 月（月營收）− 1–3Q26 財報營收（應接近 0）", f"=B{R['cum9']}-ytd_rev", None, "NT$B",
                      "Derived：月營收自結數與核閱後季報營收之差：1Q26 +0.29 億、2Q26 +0.39 億、3Q26 +1.52 億（7–9 月月營收 297.19 − 財報 295.67），合計 2.20 億（約 0.3%）；主讀數用月營收累計、第二讀數用財報 3Q26")
    R["ai"] = S.row("公司揭露：AI 相關營收金額", "未揭露", None, "", "公司未揭露 AI 相關營收金額；AI 占比為券商與研調估計（Interested-party）；FY2026 AI 相關錨點＝營收錨點 − 非 AI（非 AI 依 FY2025 × 成長假設）", fmt="General")
    R["months"] = S.row("已公布月數（2026-01～10）", 10, None, "月", "輸入：月營收至 2026-10（2026-11-06 公告）", fmt="0", blue=True)
    S.skip()
    S.bold("4b. 層級隱含值（Siliconomics 現行版；市場連到本頁第 2 節層級值列）")
    S.hdr(["項目", "CY2026", "", "", "", "單位", "口徑、來源、證據等級"])
    R["lmkt"] = S.row("AI 高速 PCB 市場（層級值，全額）", f"=B{G['mkt'][1]}", None, "US$B", "Derived：本頁第 2 節 Siliconomics SIF_PCBMkt_Base（CY2026）")
    R["lpart"] = S.row("金像電可參與比例（可參與市場 ÷ 全額市場；本模型參與係數）", f"=PL!D{mku}/B{R['lmkt']}", None, "%", "Derived（本模型假設）：中板／背板／其他 × 參與係數 w_mid", fmt="0.0%")
    R["lsh"] = S.row("金像電 FY2026 有效份額（本模型校準；Siliconomics 無此量）", f"=PL!D{sh0}", None, "%", "Derived（本模型校準）：AI 相關營收錨點 ÷ 可參與市場（PL 頁情境 B）", fmt="0.0%")
    R["lai"] = S.row("層級隱含：金像電 AI 相關營收（市場 × 可參與比例 × 份額 × 匯率）", f"=B{R['lmkt']}*B{R['lpart']}*B{R['lsh']}*fx", None, "NT$B", "Derived：匯率 In 頁 fx", fmt="#,##0.0")
    S.skip()
    R["tbl"] = S.bold("C0（實績）：公司實績 vs 層級隱含值（HTML 用）")
    R["hdr"] = S.hdr(["項目", "主讀數：2026-01～10 累計 vs 層級 CY2026 × 10 ÷ 12", "第二讀數：3Q26 財報 × 4 vs 層級 CY2026", "", "", "單位", "口徑、來源、證據等級"])
    fr = f"B{R['months']}/12"
    pct = "+0.0%;-0.0%"
    S.row("AI 相關：公司實績", "無（未揭露 AI 營收）", "無（未揭露 AI 營收）", "NT$B", "公司未揭露 AI 相關營收金額；不以模型值或估計占比補位", fmt="General")
    l = S.row("AI 相關：層級隱含", f"=B{R['lai']}*{fr}", f"=B{R['lai']}", "NT$B", "Derived：層級隱含 AI 相關營收（4b）；主讀數 × 10 ÷ 12", fmt="#,##0.0")
    R["j_ai"] = S.row("AI 相關：判定", "無可比實績（層級值同源）", "無可比實績（層級值同源）", "", "缺 AI 營收揭露；份額由本模型校準，市場大小未受檢驗；不計入對帳中", fmt="General")
    sh_ai = f"(a26_ai/a26_rev)"
    a2 = S.row("AI 相關（參考，同源校準）：錨點 AI 占比 × 公司營收實績", f"=B{R['cum10']}*{sh_ai}", f"=B{R['q3']}*{sh_ai}*4", "NT$B", "Verified 營收 × 本模型 FY2026 錨點 AI 占比（a26_ai ÷ a26_rev）；不是公司揭露金額", fmt="#,##0.0")
    g2 = S.row("AI 相關（參考，同源校準）：差距（÷ 層級隱含 − 1）", f"=B{a2}/B{l}-1", f"=C{a2}/C{l}-1", fmt=pct)
    R["j_ref"] = S.row("AI 相關（參考，同源校準）：判定", "參考（同源校準）：不判定", "參考（同源校準）：不判定", "", "兩讀數代數上等於實績營收 ÷ 營收錨點 − 1（占比與市場大小約掉），只檢驗本模型營收錨點", fmt="General")
    R["ub"] = S.row("上限檢查：層級隱含 AI 相關 ÷ 公司營收實績（主讀數期間；應 ≤ 100%）", f"=B{l}/B{R['cum10']}", None, "%", "Derived：層級 AI 相關不得超過公司全部營收；通過＝合理，不是一致性判定", fmt="0.0%")
    R["j_ub"] = S.row("上限檢查：結果", f'=IF(B{R["ub"]}<=1,"通過","不通過")', None, "", "", fmt="General")
    S.skip()
    S.bold("參考：本模型 FY2026 營收錨點對月營收實績（模型對實績，不是層級對帳；不計入 C0）")
    R["anc"] = S.row("FY2026 營收錨點（Q 頁）", f"=Q!F{lab(Q, '營收')}", None, "NT$B", "Derived／Assumed：1–3Q26 財報實績＋4Q26 假設季增 8%", fmt="#,##0.0")
    R["imp"] = S.row("錨點隱含 11–12 月月均營收（錨點 − 1～10 月累計）÷ 2", f"=(B{R['anc']}-B{R['cum10']})/2", None, "NT$B", "Derived", fmt="#,##0.00")
    R["impg"] = S.row("隱含月均 ÷ 10 月實績 − 1", f"=B{R['imp']}/B{R['oct']}-1", None, "%", "Derived：正值＝錨點假設 11–12 月高於 10 月水準", fmt=pct)
    S.skip()
    R["sum0"] = S.bold("C0（實績）結果摘要（HTML 用；公式帶數字）")
    nopen = f'COUNTIF(B{R["j_ai"]}:C{R["j_ref"]},"對帳中")'
    R["sum1"] = S.text(F(T("C0（實績）結論：無可比實績（層級值同源）：金像電未揭露 AI 相關營收金額；Siliconomics 沒有金像電份額，本模型份額以營收錨點對 Siliconomics 市場校準。對帳中項目 "), nopen,
                         T(" 項。參考：以錨點 AI 占比換算，差距主讀數 "), f'TEXT(B{g2},"+0.0%;-0.0%")', T("、第二讀數 "), f'TEXT(C{g2},"+0.0%;-0.0%")',
                         T("，只反映營收錨點與實績的差異，不判定。")))
    R["sum2"] = S.text(F(T("上限檢查：層級隱含 AI 相關（主讀數期間）占公司營收實績 "), f'TEXT(B{R["ub"]},"0%")', T("，"), f'B{R["j_ub"]}',
                         T("；10 月營收 "), f'TEXT(B{R["oct"]}*10,"#,##0.0")', T(" 億元，營收錨點隱含 11–12 月月均 "), f'TEXT(B{R["imp"]}*10,"#,##0.0")', T(" 億元（較 10 月 "),
                         f'TEXT(B{R["impg"]},"+0%;-0%")', T("）。")))
    R["sum3"] = S.text("限制與待查：公司未揭露 AI 相關營收，Siliconomics 的 AI 高速 PCB 市場也沒有獨立的金像電份額，故市場大小未獲獨立驗證；3Q26 財報已滾入（第二讀數改用財報營收），3Q26 法說未取得。")
    R["end"] = S.r - 1
    return R


def apply(wb, SIL, CR):
    _prep(wb)
    C.update(layer_rows=layer_rows, same_blocks=same_blocks)
    return c13.apply(wb, SIL, CR, C, lines, c0_actuals)
