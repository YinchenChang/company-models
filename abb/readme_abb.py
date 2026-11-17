# ABB README 頁（由 build_abb.py 呼叫；數字一律公式帶入，HTML 由本頁與各頁儲存格取數）。
from openpyxl.styles import Font, Alignment

BOLD = Font(bold=True)


def lit(s):
    s = s.replace('"', '""')
    return "&".join('"' + s[i:i + 90] + '"' for i in range(0, len(s), 90)) if s else '""'


def F(*parts):
    out = []
    for p in parts:
        if isinstance(p, str): out.append(lit(p))
        else: out.append(f'TEXT({p[0]},"{p[1]}")' if len(p) > 1 else p[0])
    return "=" + "&".join(out)


def build(wb, c):
    ws = wb.create_sheet("README", 0)
    T, S, PB, YC, q2, m0, rv = c["T"], c["S"], c["PB"], c["YC"], c["q2"], c["m0"], c["rv"]
    v = lambda k, col="B": f"Val!{col}{T[k]}"
    b = lambda key, y: f"PL!{YC[y]}{PB['B'][key]}"
    r = [1]
    EDGE = ("IF(%s>=-0.25,\"距「放空」邊界還有 \"&TEXT((%s+0.25)*100,\"0.0\")&\" 個百分點\",\"已低於「中立」邊界 \"&TEXT(-(%s+0.25)*100,\"0.0\")&\" 個百分點\")" % (v("up"), v("up"), v("up")),)

    def put(text, bold=False, plain=False):
        ws.cell(r[0], 1, text if plain or not str(text).startswith("=") else text)
        if bold: ws.cell(r[0], 1).font = BOLD
        ws.cell(r[0], 1).alignment = Alignment(wrap_text=True, vertical="top")
        r[0] += 1
        return r[0] - 1

    put(f"ABB（ABBN）個股模型 {c['VER']}　成品日 {c['BUILD_DATE']}；股價基準日 {c['PX_DATE']}（收盤）；build_id {c['BID']}。AI 供應鏈 Project：電源與散熱；新增五家（工作單 2026-11-17）。", True)
    put("用途：以 Excel 為事實基礎，檢驗 ABB 的 AI 資料中心暴露度與估值。金額：十億美元（報表幣別）；股價與目標價：瑞士法郎（SIX：ABBN）；匯率 CHF=X 0.83593（每美元）；會計年度＝日曆年。")
    put("分部：Electrification（電氣化）、Motion（馬達驅動）、Automation（自動化；2026 年起 ABB 新聞稿使用此名，原 Process Automation）。Robotics 已簽約出售給 SoftBank（企業價值約 54 億美元，預計 2026 下半年交割），原文為停業單位，本模型 FCFF 不含，出售淨款在 In 頁以非營業投資加入淨現金。")
    put("三層基準：ABB 的 AI 資料中心變數讀 Wattonomics（設施電力：機櫃外的配電與 UPS；機櫃內屬 IT），不讀 Siliconomics（設施供應商例外；見 Coord 頁）。快照 v0.3.6（檔名與 sha256 前 8 碼 " + c["WSHA"] + " 見 Wat 頁）。")
    put(F("估值：目標價＝DCF（標準版 FCFF、WACC ", (f"wacc", "0.0%"), "、永續成長 ", (f"g_tv", "0.0%"), "）與 Comparable（同業前瞻 P/E 中位數 ", (v("pmed"), "0.0"), " 倍 × FY2028 EPS）各半；DCF 情境價值以股權成本推到 12 個月後；評等依上檔空間固定規則（五級）。"))
    put("展示層", True)
    put(F("ABB 的 AI 資料中心只占營收約 ", (b("dcg", 2026), "0%"), "（FY2026E），其餘靠電網、工業電氣化、馬達與自動化；股價已以 AI 電力主題定價，本模型加權目標價 ",
          (v("tpw"), "0.0"), " 瑞士法郎，較現價 ", (v("up"), "+0.0%;-0.0%"), "，評等", (v("rec"),)))
    put(F("投資建議條件式：評等依上檔空間分五級（大於 +50% 強力買進、+25% 到 +50% 買進、−25% 到 +25% 中立、−50% 到 −25% 放空、小於 −50% 強力放空）；本模型上檔空間 ",
          (v("up"), "+0.0%;-0.0%"), "，" + "", EDGE, "；情境權重全給 A 時加權目標價約 ", (f"Sens!E{c['SJ_MAP']['wts']}*fx", "0.0"), " 瑞士法郎、全給 C 時約 ", (f"Sens!C{c['SJ_MAP']['wts']}*fx", "0.0"), "（Sens）。"))
    put(F("命題 1：ABB 的 AI 暴露度有限——資料中心只是 Electrification 的一部分，本模型估計 FY2026E 占集團營收約 ", (b("dcg", 2026), "0.0%"), "（ABB 不揭露，為假設，區間 ", ("dc26_pct_lo", "0%"), "–", ("dc26_pct_hi", "0%"),
          " 占電氣化），AI 資本支出放緩的衝擊因此有限；但 Wattonomics 的配電加 UPS 設施支出自 FY2031 起下降，這一塊在終值年轉為逆風，而不是主要成長來源。"))
    put(F("命題 2：股價已以 AI 電力主題定價——現價約為本模型 FY2027E EPS 的 ", (v("pe2027", "C").replace("Val!C", "Val!C"), "0.0"), " 倍，高於同業前瞻 P/E 中位數 ", (v("pmed"), "0.0"),
          " 倍；要支持現價，需要比本模型基準更高的利潤率或成長（見反向 DCF）。"))
    put(F("命題 3：非資料中心事業才決定 ABB 的價值——Electrification 在手訂單年增 83%、Motion 與 Automation 各增 14%、11%（2026-09-30，公司原文），本模型 FY2027E 非資料中心營收占集團 ",
          ("1-" + b("dcg", 2027), "0%"), "，Operational EBITA 率 ", (b("be_r", 2027), "0.0%"), "（FY2025 ", (b("be_r", 2025), "0.0%"), "）。"))
    put(F("重點 1：AI 暴露度有限——資料中心營收本模型估計 FY2026E 約 ", (b("dc", 2026), "0.0"), " 十億美元、占集團 ", (b("dcg", 2026), "0.0%"), "；其餘約 ", ("1-" + b("dcg", 2026), "0%"), " 來自電網、建築、工業電氣化、馬達與自動化。"))
    put(F("重點 2：資料中心營收隨 Wattonomics 配電加 UPS 設施支出：FY2026 ", (b("wat", 2026), "0"), "、FY2030 高點 ", (b("wat", 2030), "0"), "、FY2032 降到 ", (b("wat", 2032), "0"),
          " 十億美元；本模型 FY2032E 資料中心營收 ", (b("dc", 2032), "0.0"), " 十億美元、占集團 ", (b("dcg", 2032), "0.0%"), "。"))
    put(F("重點 3：FY2027E 營收 ", (b("rev", 2027), "0.0"), "、EPS ", (b("eps", 2027), "0.00"), " 美元，對市場共識 ", ("cons_rev27", "0.0"), "、", ("cons_eps27", "0.00"), " 各 ",
          (b("rev", 2027) + "/cons_rev27-1", "+0.0%;-0.0%"), "、", (b("eps", 2027) + "/cons_eps27-1", "+0.0%;-0.0%"), "；FY2026E 前三季為公司實績。"))
    put(F("重點 4：DCF 以 WACC ", ("wacc", "0.0%"), "、永續成長 ", ("g_tv", "0.0%"), " 得 ", (v("dcf"), "0.0"), " 瑞士法郎，遠低於現價 ", (v("px"), "0.0"), "；Comparable 以同業 ", (v("pmed"), "0.0"),
          " 倍乘 FY2028 EPS 得 ", (v("comp"), "0.0"), "；兩法各占一半。"))
    put(F("重點 5：評等「", (v("rec"),), "」，", EDGE, "；對目標價影響最大的是情境權重、DCF 權重、β 與同業倍數，資料中心份額與 Wattonomics 係數各約 ±", (f"Sens!F{c['SJ_MAP']['dc_sh']}", "0.0%"), "、±", (f"Sens!F{c['SJ_MAP']['wif']}", "0.0%"), "。"))
    put("命題回覆", True)
    put(F("回覆 1：資料中心占集團營收 FY2025 ", (b("dcg", 2025), "0.0%"), " → FY2026E ", (b("dcg", 2026), "0.0%"), " → FY2028E ", (b("dcg", 2028), "0.0%"), " → FY2032E ", (b("dcg", 2032), "0.0%"),
          "；集團營收 FY2026E–FY2032E 年均成長 ", (f"({b('rev', 2032)}/{b('rev', 2026)})^(1/6)-1", "0.0%"), "，成長主要來自非資料中心事業與 FY2026–FY2028 的資料中心放量。"))
    put(F("回覆 2：現價隱含 FY2027–FY2032 平均營業利益率 ", (f"RevDCF!B{rv['m1']}", "0.0%"), "（模型基準 ", (f"RevDCF!B{rv['m0']}", "0.0%"), "），或 FY2027 營收規模為模型的 ", (f"RevDCF!B{rv['s']}", "0.00"),
          " 倍（情境 B 路徑）；Comparable 隱含前瞻 P/E ", (f"RevDCF!B{rv['cm']}", "0.0"), " 倍對同業 ", (f"RevDCF!B{rv['cp']}", "0.0"), " 倍。"))
    put(F("回覆 3：非資料中心事業 FY2027E 營收 ", (f"{b('rev', 2027)}-{b('dc', 2027)}", "0.0"), " 十億美元，占集團 ", ("1-" + b("dcg", 2027), "0%"), "；若其成長與利潤率取 In 頁低值，加權目標價變動見 Sens（電氣化、馬達、自動化各案例）。"))
    put("", False)
    put("版本要點（v0.1.0）", True)
    for t in ["首版：新增五家工作單（2026-11-17）；公司原文為 Q4 2025 financial information（FY2025／FY2024，Robotics 重編後）與 Q3 2026 新聞稿（9M 2026、資產負債表、指引）。",
              "FY2026E＝前九個月實績＋Q4 模型（公司指引：Q4 可比營收成長與 Q3 相近、Operational EBITA 率季減幅小於歷史；Q4 另有約 2.4 億瑞士法郎房產出售收益）。",
              "三情境 A 順利／B 基準／C 不如預期，權重 0.3／0.5／0.2；FY2027–FY2032 為明確期；終值以 g＝MIN(3%, 美元 10 年期公債殖利率)、RONIC＝WACC。",
              "Coord／alt 讀數：不適用（電力與機房範圍定案 2026-11-17）：機房設施供應商只讀 Wattonomics 廠房拆分，該數字只有 Wattonomics 一個權威層、沒有 Siliconomics 側的另一讀數，三層協調 C4 兩讀數不適用；與 Wattonomics 的差異走對帳單。"]:
        put("• " + t)
    put("待判斷（區間與標記；判斷由 Andy 或未參與者複核）", True)
    for t in ["1. 資料中心營收拆分：ABB 不揭露。FY2025 占電氣化 10%–18%（基準 14%）、FY2026E 17%–27%（基準 22%）[Assumed]；Sens「資料中心校準」案例。",
              "2. ABB 占 Wattonomics 配電＋UPS 設施支出的份額：FY2027 6%–9%（基準 7.5%）、FY2032 5%–9%（基準 7%）；情境 A 加 0.5–1.5pt、C 減 1–2pt [Assumed]。FY2025–FY2026 隱含 7%–8%（Derived，PL 頁）。",
              "3. Wattonomics 係數 0.8–1.2（基準 1.0）[Assumed]：快照無區間，本係數只測 ABB 的敏感度；終值年（FY2032）落在 Wattonomics 設施支出下降段（FY2030 約 98.6 → FY2032 約 61.2 十億美元），資料中心營收在終值年偏保守。",
              "4. Robotics 出售淨款 44–54 億美元（基準 50 億）[Assumed]：企業價值約 54 億美元為公司公布；稅負與分離成本為假設；交割時點 2026 下半年，若延後或失敗需重估。",
              "5. Rotork 收購（2026-07 宣布）與三菱電機低壓配電 JV（51%，對價約 4.95 億美元、企業價值約 8.9 億美元、FY2025 營收約 6.4 億美元，公司原文）皆不計入基準，視為價值中性；Rotork 交易條款本版未取得，待補。",
              "6. β：Yahoo 單一來源 1.05，去槓桿後依 ABB 的 D／E 重新加槓桿，區間 β_u ± 0.3 [Analogy]；尚無自算的電氣設備同業 β。美元 rf 5.10%（FRED 2026-11-12）、CRP 0 [Assumed]。",
              "7. 非控制權益：帳面 0.559 十億美元（2025-12）；ABB India 市價下的少數股權價值可能高到約 3.0 十億美元 [Assumed；外部事實未檢核]，影響約 −2%。",
              "8. 淨退休金資產 0.619 十億美元（公司原文）保守不計入淨現金；計入約每股 +0.3 美元。",
              "9. Q4 房產出售收益 2.4 億瑞士法郎全額計入營業利益與 stub 期現金 [Assumed：現金等值、稅負]；區間 0–全額，影響每股約 ±0.2 美元。",
              "10. 股利：評價日到目標時點的股利未扣除（dps_tgt＝0）[Assumed]；ABB 年度股利約每股 0.9 瑞士法郎 [外部・未經本專案檢核]，若扣除目標價約 −1.3%。",
              "11. 買回（20 億美元計畫）視為價值中性，不調整現金與股數；EPS 另以每年 −1%（−1.5% 到 0%）的股數變動計。",
              "12. 評等邊界：上檔空間距「中立」邊界（−25%）不到 1 個百分點，評等對 In 頁微小改動敏感；請以目標價區間（價格帶）與 Sens 閱讀，不只看評等。",
              "13. 共識口徑：Yahoo 共識 FY2026 EPS 3.244 隱含 Q4 EPS 約 1.18（高於本模型含房產收益的 Q4），可能含 Robotics 出售收益或為調整後口徑，待核；Yahoo 的 ABB 前瞻 P/E 18.57 把瑞士法郎股價除以美元 EPS，幣別錯置，不採用。"]:
        put(t)
    put("公司原文與 Yahoo 的差異（詳見 Recon 頁）", True)
    for t in ["FY2024 營收 30,583（公司 Q4 2025 原文、Yahoo 皆同）：工作單註記的 32,850 為 2024 年報未重編數（含 Robotics 2,267），差異已查明。",
              "營業利益：Yahoo＝毛利−銷管−研發，不含「其他收入（淨額）」（FY2025 318、FY2024 −18）；公司損益表 Income from operations 含之。",
              "FY2022–FY2023 的 Yahoo 數為未重編（含 Robotics），不可與 FY2024 以後比較；公司原文的 FY2023 重編數未取得（PL 頁 FY2023A 僅供參考）。",
              "非控制權益 2025-12 公司 559、2026-09 Yahoo 483；Q3 2026 新聞稿未揭露。"]:
        put("• " + t)
    put("建議補進 Siliconomics／Wattonomics（本模型讀 Wattonomics，不改上游）", True)
    for t in ["Wattonomics：加入「機櫃外設施供應商份額」對照層（ABB、Schneider、Eaton、Vertiv、Siemens），讓 ABB 的份額假設有上游證據與區間。",
              "Wattonomics：配電＋UPS 設施支出在 FY2031–FY2032 的下降段（對應新增 GW 下滑）是否為情境或定案，影響所有設施供應商的終值。",
              "三層協調規則：寫入「設施供應商讀 Wattonomics」的例外（Coord 頁已註記），並讓 tools/coord/alt_run.py 支援讀 Wattonomics 的另一側讀數。",
              "同業 β：建立電氣設備／電力設備同業組的自算產業 β（現用 Yahoo 單一來源）。"]:
        put("• " + t)
    put("已知限制", True)
    for t in ["資料中心營收為本模型拆分（ABB 不揭露）；公司只說資料中心訂單 Q3 2026 起三位數增長。",
              "Rotork 與三菱電機 JV 未計入；Robotics 出售價款為假設。",
              "未做相對倍數調整（ABB 沒有 10 年前瞻倍數歷史）；Comparable 只有 4 檔同業且來源單一（Yahoo）。",
              "FY2023A 為未重編；公司原文頁碼待人工複核（表格以算術核對：分部合計、Operational EBITA 合計與集團一致）。",
              "目標時點為評價日後 12 個月；匯率固定在 2026-11-13（CHF=X 0.83593），不預測匯率。"]:
        put("• " + t)
    ws.column_dimensions["A"].width = 160
