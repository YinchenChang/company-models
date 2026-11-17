# build_vrt.py 的第三段：README、列位置輸出、存檔
RD = wb.create_sheet("README", 0)
tx = lambda ref, f: f'TEXT({ref},"{f}")'
TPW, DCFV, CMPA, UPW = f"Val!B{T['tpw']}", f"Val!B{T['dcftp']}", f"Val!B{T['compadj']}", f"Val!B{T['upw']}"
PLB = lambda key, y: f"PL!{FC[y]}{PB['B'][key]}"
R_ = {}
def put(key, text, bold=False):
    r = RD.max_row + 1 if RD.max_row > 1 or RD["A1"].value else 1
    RD.cell(r, 1, text)
    if bold: RD.cell(r, 1).font = BOLD
    R_[key] = r
    return r
put("head", f"個股模型 v{VER}（{BUILT}）　Vertiv（VRT）　AI 供應鏈：電源與散熱（機房設施端）　build_id {BUILD_ID}", True)
put("scope", "範圍：Vertiv 的機房電力（UPS、配電）、熱管理（含液冷 CDU）與服務。命題＝Vertiv 的資料中心設施營收是 Wattonomics 機房設施市場（機櫃外）乘以 Vertiv 份額；份額由訂單與出貨校準。")
put("basis", "會計年度＝日曆年（FY(t)＝CY(t)）；金額十億美元（US$B）；EPS 美元（GAAP 稀釋）；歸屬普通股淨利。FY2026＝1Q–3Q 實績＋4Q 假設（公司 4Q 指引中位）。FY2030 起為收斂法。")
put("tags", "標記：Verified（公司原文）、Interested-party（賣方、管理層指引、資料商；另寫來源誘因）、Analogy、Assumed、Derived。公司原文＝Vertiv 新聞稿與 10-Q；Yahoo 只作交叉核對（差異見 Recon）。")
put("mkt", "產業變數：機房設施市場讀 Wattonomics v0.3.6（WIF 配電、UPS、冷卻；機櫃內屬 IT 端，由台達電模型負責，不計入）。不讀 Siliconomics 的電源與散熱。Vertiv 份額為 Derived，區間進 Sens。")
put("sp1", "")
put("h_thesis", "命題", True)
put("p1", "命題一：Vertiv 的資料中心設施營收等於機櫃外設施市場（配電、UPS、冷卻）乘以 Vertiv 份額；市場來自 Wattonomics，份額由 backlog 與出貨校準。")
put("p2", "命題二：設施市場成長放緩後，營業費用的規模效應與毛利率能否撐住營業利益率，決定現價隱含的成長是否合理。")
put("p3", "命題三：Vertiv 的估值取決於市場成長在 FY2029 之後能否延續；Wattonomics 基準在該年後新增容量下降，是模型最大的不確定來源。")
put("q", "設施市場在 Wattonomics 基準下成長放緩後，Vertiv 的份額與利潤率能否撐住現價隱含的成長")
put("h_reply", "命題回覆", True)
put("r1", f'="現價合不合理？答："&IF({UPW}<0,"現價高於本模型的合理價值。","現價低於本模型的合理價值。")&"加權目標價約 "&{tx(TPW, "#,##0")}&" 美元（DCF "&{tx(DCFV, "#,##0")}&"、同業倍數法調整後 "&{tx(CMPA, "#,##0")}&"），比現價 "&{tx("px", "#,##0.0")}&" 美元"&IF({UPW}<0,"低 ","高 ")&{tx("ABS(" + UPW + ")", "0.0%")}&"，投資建議"&{REC}&"。"')
put("r2", f'="市場成長放緩後營業利益撐得住嗎？答：情境 B 設施市場 FY2027 成長 "&{tx("Units!E10", "0%")}&"、FY2029 僅 "&{tx("Units!G10", "0.0%")}&"；營收 FY2026–FY2029 年均成長 "&{tx(f"({PLB("rev", 2029)}/{PLB("rev", 2026)})^(1/3)-1", "0.0%")}&"，營業利益率由 "&{tx(PLB("opm", 2026), "0.0%")}&" 到 "&{tx(PLB("opm", 2029), "0.0%")}&"，FY2029 EPS "&{tx(PLB("eps", 2029), "0.00")}&" 美元。"')
put("r3", f'="現價隱含什麼？答：以情境 B 現金流估算，現價相當於模型營收規模的 "&{tx("RevDCF!B7", "0.00")}&" 倍（FY2027 營收約 "&{tx("RevDCF!B8", "0.0")}&" 十億美元，共識 "&{tx("cons_rev27", "0.0")}&" 十億美元）；在 WACC "&{tx("wacc", "0.0%")}&" 下現價已預期市場或份額遠高於 Wattonomics 基準，或折現率低於本模型。"')
put("sp2", "")
put("h_obs", "重點觀察", True)
put("o1", f'="目標價 "&{tx(TPW, "#,##0")}&" 美元（DCF "&{tx(DCFV, "#,##0")}&"、同業倍數法調整後 "&{tx(CMPA, "#,##0")}&" 各半），比現價 "&{tx("px", "#,##0.0")}&" 美元"&IF({UPW}<0,"低 ","高 ")&{tx("ABS(" + UPW + ")", "0.0%")}&"，級距 "&Val!B{T["grade"]}&"，投資建議"&{REC}&"。"')
put("o2", f'="差距主因在 DCF：WACC "&{tx("wacc", "0.0%")}&"、永續成長率 "&{tx("g_tv", "0.0%")}&" 下，終值隱含 EV／EBITDA 僅 "&{tx(f"Val!C{SDCF['mult']}", "0.0")}&" 倍，DCF 價值 "&{tx(DCFV, "#,##0")}&" 美元，約為現價的 "&{tx(DCFV + "/px", "0%")}&"；同業倍數法（調整後 "&{tx(CMPA, "#,##0")}&" 美元）與現價相當；設施市場 FY2029 年增僅 "&{tx("Units!G10", "0.0%")}&"。"')
put("o3", f'="份額：FY2025 Vertiv 資料中心設施產品營收約占可及市場 "&{tx("Units!C12", "0.0%")}&"，FY2026 降到 "&{tx("Units!D12", "0.0%")}&"（市場 FY2026 成長 "&{tx("Units!D10", "0%")}&"，營收成長 "&{tx("Coord!C15", "0%")}&"）；backlog 與營收認列落後使份額被低估，基準 FY2027 份額相對值 "&{tx("rel_27", "0.00")}&"。"')
put("o4", f'="EPS：情境 B FY2027 GAAP EPS "&{tx(PLB("eps", 2027), "0.00")}&" 美元，共識（調整後）"&{tx("cons_eps27", "0.00")}&" 美元，差 "&{tx("Val!B" + str(T["geps"]), "0.0%")}&"（含攤銷口徑差）；FY2026 錨點 GAAP EPS "&{tx("a26_eps", "0.00")}&" 美元，公司 GAAP 指引中位 6.29 美元。"')
put("o5", f'="敏感度：WACC ±1pt 使加權目標價 "&TEXT(Sens!F{TPS0 + 2},"+0.0%;-0.0%")&"／"&TEXT(Sens!D{TPS0 + 2},"+0.0%;-0.0%")&"；DCF 權重 0.3／0.7 使其 "&TEXT(Sens!D{TPS0 + 11},"+0.0%;-0.0%")&"／"&TEXT(Sens!F{TPS0 + 11},"+0.0%;-0.0%")&"；業務主因見 Sens 業務主因表。"')
put("rec", f'="投資建議："&{REC}&"（依加權目標價與現價的固定規則；上檔空間 "&{tx(UPW, "+0.0%;-0.0%")}&"）"', True)
put("banner", f'="Vertiv：加權目標價 "&{tx(TPW, "#,##0")}&" 美元，現價 "&{tx("px", "#,##0.0")}&" 美元，上檔空間 "&{tx(UPW, "+0.0%;-0.0%")}&"，評等"&{REC}&"。市場讀 Wattonomics 基準；本模型為新建初版。"')
put("valn", "估值：加權目標價＝標準版 DCF 與 Comparable（機房電力同業前瞻 P/E 中位數 × 相對倍數調整係數）各半；DCF 以 GAAP 營業利益為 EBIT，Non-GAAP 調整項目在 Q 頁橋接。")
put("sp3", "")
put("h_lim", "已知限制", True)
put("lim", "限制：公司未揭露產品線營收與資料中心占比（dcp 為 Assumed）；backlog 與 book-to-bill 只取得 FY2025 底（150 億美元、約 2.9），3Q26 的 backlog 未找到；Wattonomics 市場為裝機年度支出，與 Vertiv 認列時點有落差；FY2030–FY2032 為收斂法。")
put("ver", f"版本：v{VER}（{BUILT}）新建：Vertiv 個股模型初版（機房設施市場讀 Wattonomics v0.3.6；滾入 3Q26 實績與 4Q26 指引；DCF、Comparable、Sens、Coord、RevDCF、Recon；HTML 與地圖檔 profiles/VRT.md）。")
put("sp4", "")
put("h_todo", "待判斷（判斷題；已先決定並標範圍，Andy 可改 In 頁黃底或藍字輸入）", True)
TODO = ["資料中心設施產品占營收比例 dcp：FY2026 取 0.72（範圍 0.65–0.80；公司未揭露，由 10-Q 產品占 83% 推估）；範圍進 Sens。",
        "情境倍數：設施市場 2028／2029 倍數 A 1.10／1.15、B 1.00／1.00、C 0.80／0.65（個股自訂 Assumed；建議由 Wattonomics 提供官方上下情境）。",
        "份額路徑：FY2027／28／29 相對 FY2026 取 1.03／1.02／1.00（範圍 0.95–1.10），依 backlog 與訂單；FY2025 份額高於 FY2026 為認列落後，不當作趨勢。",
        "FY2030–FY2032 不讀 Wattonomics（新增容量下降）：營收年增由 FY2029 線性收斂到 3%（A／C 為 ±2pt；範圍 2%–4%），營業利益率收斂係數 0.85（範圍 0.75–0.95）；若改讀 Wattonomics 原值，營收將轉降。",
        "UtilityInnovation 收購（2026-10-26，現金 14.5 億美元）：基準視為價值中性，不扣淨現金也不加營收；Sens 另列「現金計為純流出」下緣（約為淨現金 14.5 億美元）。",
        "β：Yahoo 原始 2.11、Blume 調整 1.74，只有單一來源；Sens 區間 1.45–2.05（Analogy／Assumed）。",
        "同業：Eaton、Schneider Electric、台達電三家（機房電力與電氣；沿用 delta 模型同業前瞻 EPS）；Vertiv 自身 Bloomberg 前瞻 P/E 歷史有 7 年，剛好達相對倍數調整門檻，係數受夾值 0.5–2.0 限制。",
        "EPS 口徑：本模型 GAAP，共識與同業前瞻 EPS 為調整後或混合口徑；Comparable 因此略偏高（約加回攤銷的 0.7–0.9 美元），未另行調整。",
        "稀釋股數：評價用 3Q26 稀釋加權 392.7M（流通 385.1M），較保守。",
        "折現率：無風險利率 5.00%（同 NVIDIA 模型）、β 1.74、ERP 5% 使 WACC 約 13.3%（β 區間 1.45–2.05 對應 WACC 約 11.9%–14.8%）；終值隱含 EV／EBITDA 僅約 5 倍，是 DCF 遠低於同業倍數法的主因，不另行調整。",
        "Coord／alt 讀數：待主串加入 coord_meta（vrt 讀 Wattonomics 層，非 Siliconomics hardware_companies）與建檔目標後，才能跑 tools/coord/alt_run.py vrt；Coord 讀 Wattonomics，alt 讀數待主串。"]
for i, t in enumerate(TODO, 1): put(f"todo{i}", f"{i}. {t}")
put("sp5", "")
put("h_sug", "建議補進 Siliconomics／Wattonomics（本模型以 Assumed 暫代，不自行改上游）", True)
SUG = ["Wattonomics：機房設施市場（機櫃外配電、UPS、冷卻）的官方上下情境，取代本模型個股自訂的 2028／2029 倍數。",
       "Wattonomics：2030–2032 新增容量下降之外，補設施端更換與服務需求（Vertiv 服務與備品約占營收 17%–18%），避免市場層在 2030 年後系統性偏低。",
       "Siliconomics／Wattonomics：登錄機櫃內與機櫃外的邊界與每 GW 內容對照（本模型邊界：機櫃內歸 IT 端台達電；CDU 歸設施冷卻）。",
       "Wattonomics：揭露市場支出與廠商出貨認列的時差（backlog 轉營收約 2–4 季），便於各公司校準份額。"]
for i, t in enumerate(SUG, 1): put(f"sug{i}", f"{i}. {t}")
RD["A1"].value = RD["A1"].value  # build_id 已含於 A1
RD.column_dimensions["A"].width = 180
for rr_ in range(1, RD.max_row + 1): RD.cell(rr_, 1).alignment = openpyxl.styles.Alignment(wrap_text=True, vertical="top")

order = ["README", "In", "Q", "Hist", "Units", "PL", "Val", "Coord", "DCF", "RevDCF", "Recon", "Sens", "Sources"]
wb._sheets = [wb[n] for n in order]
if os.environ.get("VRT_ROWMAP"):
    json.dump({"T": {k: v for k, v in T.items() if isinstance(v, int)}, "RM": T["rm"], "SDCF": SDCF, "SUMR": {f"{s}|{k}": v for (s, k), v in SUMR.items()}, "PB": PB, "SRW": SRW, "README": R_,
               "TPS0": TPS0, "FD_FIRST": _tvs[0], "FD_LAST": SE_FD_LAST, "BIZR": BIZR, "BR0": BR0, "BIZ_LAST": BIZ_LAST, "BID_ROW": BID_ROW, "SCHK": SCHK, "SEGR": SEGR, "RD_": {k: v for k, v in RD_.items() if isinstance(v, (int, str))},
               "INLAST": INLAST, "SEC_LABELS": {k: [x[1] for x in BIZ if x[0] == k] for k, _, _ in BIZ}},
              open(os.environ["VRT_ROWMAP"], "w"), ensure_ascii=False, indent=1, default=str)
out = sys.argv[1]
wb.save(out)
print("saved", out, BUILD_ID)
