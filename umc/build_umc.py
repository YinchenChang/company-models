# 聯電（2303.TW；ADR UMC）個股模型 v0.1（AI 半導體 Project：代工層，成熟製程晶圓代工；樣板 guc/，營收結構參考 tsmc/）
# 用法：python3 umc/build_umc.py raw.xlsx ；再以 LibreOffice headless 重算
# 會計年度＝日曆年，FY(t)＝CY(t)。金額：十億新台幣（NT$B）；晶圓：千片 12 吋約當；售價：美元／片；EPS：新台幣（基本）。
# 營收＝成熟製程晶圓（產能 × 稼動率 × 平均售價 × 匯率）＋AI 相關（先進封裝矽中介層、矽光子；美元）＋Intel 12 奈米合作（美元）。
# 本檔上半部為公司設定（數字、文字、來源），下半部為與 vis/build_vis.py 相同的共用建置程式。
SA = "https://stockanalysis.com/quote/tpe/2303/"
SAF = "https://stockanalysis.com/quote/tpe/2303/forecast/"
SAI = "https://stockanalysis.com/quote/tpe/2303/financials/"
SAR = "https://stockanalysis.com/quote/tpe/2303/financials/ratios/"
SAH = "https://stockanalysis.com/quote/tpe/2303/history/"
K3 = "https://www.sec.gov/Archives/edgar/data/0001033767/000119312526434992/umc-ex99.htm"
K2 = "https://www.sec.gov/Archives/edgar/data/0001033767/000119312526322027/umc-ex99.htm"
K1 = "https://www.sec.gov/Archives/edgar/data/0001033767/000119312526188788/umc-ex99.htm"
K4 = "https://www.sec.gov/Archives/edgar/data/1033767/000119312526025557/umc-ex99_1.htm"
UA3 = "https://uanalyze.com.tw/articles/9109557881"
UA2 = "https://uanalyze.com.tw/articles/6783952429"
FS3 = "https://finsee.ai/earnings/umc/2026/q3/en/"
MS10 = "https://finance.technews.tw/2026/10/06/morgan-stanley-increases-its-support-for-the-two-leading-semiconductor-foundries/"
WR17 = "https://wantrich.chinatimes.com/news/20261017900003-420101"
CMC = "https://companiesmarketcap.com/united-microelectronics/pe-ratio/"
C = dict(kind="umc", folder="umc", name="聯電", wafer_unit="12 吋約當", wafer_name="成熟製程晶圓", wafer_rev_label="成熟製程晶圓營收（不含 AI 相關、Intel 12 奈米）")
C["urlname"] = {K3: "公司 6-K 3Q26 財報（2026-10-28）", K2: "公司 6-K 2Q26 財報（2026-07-29）", K1: "公司 6-K 1Q26 財報", K4: "公司 6-K 4Q25 財報（2026-01-28）",
                UA3: "優分析 2026-10-28（3Q26 法說轉述）", UA2: "優分析 2026-07-29（2Q26 法說轉述）", FS3: "Finsee 3Q26 財報摘要", MS10: "科技新報 2026-10-06（摩根士丹利）",
                WR17: "旺得富 2026-10-17", SA: "stockanalysis.com", SAF: "stockanalysis.com 共識頁", SAI: "stockanalysis.com 財報頁", SAR: "stockanalysis.com 比率頁", SAH: "stockanalysis.com 股價歷史", CMC: "companiesmarketcap"}
C["cap26_src"] = "Q 頁：各季出貨 ÷ 稼動率（1Q–3Q26 實績、4Q26 指引）"
C["cap26_note"] = "公司未揭露產能；以出貨與稼動率推得（Derived）"
C["capg"] = {27: (0.04, "公司：董事會核准新加坡 P4、台南新廠擴產預算約 1,487 億元（2026-07-29）與約 1,028 億元（2026-10-28；6-K）"), 28: (0.07, "同上；Fab12i P3 量產時程：未找到"), 29: (0.07, "同上")}
C["ut"] = {27: {"A": 0.95, "B": 0.90, "C": 0.82}, 28: {"A": 0.94, "B": 0.87, "C": 0.76}, 29: {"A": 0.93, "B": 0.88, "C": 0.80}}
C["ut_src"] = "公司：3Q26 稼動率 92%、4Q26 指引低 90% 區間（6-K）；1Q26 79%、2Q26 85%"
C["asp"] = {27: {"A": 0.10, "B": 0.05, "C": 0.0}, 28: {"A": 0.03, "B": 0.0, "C": -0.06}, 29: {"A": 0.02, "B": 0.0, "C": -0.03}}
C["asp_src"] = UA3 + "：公司估 2027 年平均售價年增約一成（22 奈米份額與產品組合）；摩根士丹利：8 吋代工 2027 年漲價"
C["asp26_f"] = "=(a26_rev-ai_26*fx_26)/a26_ship/fx_26*1000000"
C["asp26_src"] = "（FY2026 營收 − AI 相關）÷ 出貨 ÷ 匯率"
C["ai26"] = 0.3
C["ai26_src"] = MS10 + "：公司說 2026 年 AI 相關營收約 3 億美元（營收 3–5%）、2029 年 10 億美元"
C["aig27"] = {"A": 0.6, "B": 0.4, "C": 0.2}
C["ai_src2"] = FS3 + "：先進封裝與矽光子客戶接洽「強勁」；公司 2029 年 AI 相關 10 億美元目標（大摩轉述）"
C["aig_L"] = 0.45
C["int"] = {28: {"A": 0.25, "B": 0.12, "C": 0.0}, 29: {"A": 0.5, "B": 0.3, "C": 0.08}}
C["int_src"] = FS3 + "：設計套件 2026 年底、2027 年試產（先前說法）；旺得富 2026-10-17：驗證中、未來幾年貢獻營收"
C["da26"] = "=46.81/3*4"; C["da26_tag"] = "Derived"
C["da26_src"] = FS3 + "：9M26 折舊 468.1 億元（年增 12.7%）× 4／3"
C["da26_note"] = "第四季折舊：未找到，以前三季平均推估"
C["dag"] = {27: (0.12, "公司財務長：未來 1–2 年折舊與原物料成本上升（優分析 2026-10-28）"), 28: (0.10, "同上；P4、台南新廠"), 29: (0.08, "同上")}
C["opx_src"] = "FY2025 營業費用 250.6 億、9M26 約 204.8 億（Q 頁）"
C["nonop"] = (2.0, 4.0, 8.0)
C["nonop_src"] = "FY2025 約 58 億、1Q26 約 54 億、3Q26 約 13 億（Q 頁 Derived）"
C["nonop_note"] = "2Q26 業外約 302 億（投資與股利收益，優分析）為一次性，不外推"
C["tax_src"] = "FY2025 約 16%、3Q26 約 18%（Q 頁 Derived）"
C["sh_src"] = "3Q26 淨利 ÷ EPS 推得約 124.5 億股（Q 頁）"
C["sh_note"] = "海外可轉債約 18 億美元（旺得富 2026-10-17）的轉換稀釋未建模：轉換價未找到"
C["fx26"] = 31.6
C["fx26_src"] = "公司：3Q26 營收 742.4 億元＝23.2 億美元（6-K，約 32.0）；2Q26 淨利換算約 31.5"
C["q_inputs"] = [
    ("q4_ship", "4Q26 晶圓出貨季增", "%", 0.02, "Interested-party", "公司 4Q26 指引：出貨低個位數成長（6-K 2026-10-28）", "取 2%"),
    ("q4_asp", "4Q26 平均售價季變動（美元）", "%", 0.0, "Interested-party", "公司 4Q26 指引：美元平均售價持穩", "—"),
    ("q4_ut", "4Q26 稼動率", "%", 0.92, "Interested-party", "公司 4Q26 指引：低 90% 區間", "取 92%"),
    ("q4_gm", "4Q26 毛利率", "%", 0.34, "Interested-party", "公司 4Q26 指引：與上季相當（3Q26 34.0%）", "—"),
    ("q4_opx", "4Q26 營業費用（3Q26 × 1.03）", "NT$B", "=Q!E11*1.03", "Assumed", "3Q26 營業費用（Q 頁）", "—"),
    ("q4_nonop", "4Q26 營業外淨額", "NT$B", 1.5, "Assumed", "3Q26 約 12.7 億（Q 頁）", "—"),
    ("q4_tax", "4Q26 有效稅率", "%", 0.17, "Assumed", "3Q26 約 17.7%（Q 頁）", "—"),
    ("q_sh", "加權股數（3Q26）", "十億股", "=Q!E19", "Derived", "3Q26 淨利 ÷ EPS", "stockanalysis 流通股數 125.8 億股")]
C["cons"] = [
    ("cons_rev26", "共識營收 FY2026（只作對照，不進模型）", "NT$B", 283.2, SAF + "（21 位；2026-10-08 擷取，早於 3Q26 財報）", "區間 271.5–291.8；FY2027 起共識為付費資料：未找到"),
    ("cons_eps26", "共識 EPS FY2026（只作對照，不進模型）", "NT$", 7.02, SAF + "（2026-10-08 擷取）", "區間 4.55–8.15"),
    ("us_eps26", "美系外資 FY2026 EPS（中點）", "NT$", 7.57, WR17 + "：7.46–7.68", "目標價 130"),
    ("loc_eps27", "本土投顧 FY2027 EPS", "NT$", 9.2, WR17, "目標價 185（強力買進）；FY2027 共識：未找到"),
    ("co_asp27", "公司：FY2027 平均售價年增約一成", "%", 0.10, UA3, "—"),
    ("co_ai29", "公司：FY2029 AI 相關營收約 10 億美元", "US$B", 1.0, MS10, "大摩轉述")]
C["px"] = 125.5; C["px_date"] = "2026-10-22"
C["px_src"] = SAH + "：10-22 收盤 125.50"; C["px_src2"] = "10-23 以後收盤：未找到（stockanalysis 歷史頁只到 10-22；10-28 法說後股價未取得）"
C["eps_ttm"] = 6.81; C["eps_ttm_src"] = "4Q25 0.81＋1Q–3Q26 6.00（公司 6-K）；含 2Q26 一次性業外"
C["fpe_lab"] = "前瞻 P/E（stockanalysis，2026-10-22；股價 125.5）"; C["fpe"] = 20.22; C["fpe_src"] = SA
C["cfeps_lab"] = "共識前瞻 12 月 EPS（2026-10-22 股價 ÷ 前瞻 P/E）"; C["cfeps"] = "=125.5/fpe"; C["cfeps_src"] = "stockanalysis 前瞻 P/E 20.22（股價 125.5）"
C["cons_tp"] = 140.69; C["cons_tp_src"] = "stockanalysis：25 位分析師平均 140.69（2026-10-08 擷取；中位數 135、區間 41–258）"
C["cons_tp_src2"] = "旺得富 2026-10-17：法人目標價 130–185；FactSet 共識均值 84（口徑不同，未採用）"
C["nim5"] = "=AVERAGE(51.246/213.011,89.479/278.705,60.99/222.533,47.211/232.303,41.716/237.553)"
C["nim5_src"] = "stockanalysis 財報頁：FY2021–25 營收與淨利"
C["roe_f"] = "=AVERAGE(0.1801,0.1277,0.1096)"; C["roe_src"] = "stockanalysis 比率頁：FY2023 18.01%、FY2024 12.77%、FY2025 10.96%"
C["q_a1"] = "Q — 實績與 FY2026 錨點（十億新台幣；EPS 為新台幣、基本；股數為十億股；晶圓為千片 12 吋約當）。藍字＝輸入（公司 6-K、stockanalysis）；黑字＝公式"
C["q_a2"] = "FY2026 錨點＝1Q、2Q、3Q26 實績（3Q26 財報 2026-10-28 公布；季值＝累計 − 上半年）＋4Q26 公司指引（出貨低個位數成長、售價持穩、毛利率與上季相當、稼動率低 90%）"
C["q3_state"] = "實績"
C["q_s1"] = "公司 6-K：FY2025 與 4Q25（2026-01-28）、1Q26、2Q26（2026-07-29）、3Q26（2026-10-28）；FY2025 損益：stockanalysis 財報頁"
FX = {c: "fx_26" for c in "CDEF"}
C["qd"] = dict(
    rev={"B": 237.553, "C": 61.038, "D": "=129.771-C4", "E": "=204.010-129.771", "F": "=E4*(1+q4_ship)*(1+q4_asp)"}, rev_tag="Verified／Interested-party（4Q 指引）",
    rev_s2="1H26 1,297.71 億、9M26 2,040.10 億（6-K）；4Q26＝3Q26 ×（1＋出貨季增）×（1＋售價變動）",
    ship={"C": "=D5/(1+0.106)", "D": "=E5/(1+0.073)", "E": 1200, "F": "=E5*(1+q4_ship)"}, ship_tag="Interested-party／Derived",
    ship_s1="公司：3Q26 出貨超過 120 萬片 12 吋約當（創新高，季增 7.3%）；2Q26 季增 10.6%；1Q26 季增 2.7%", ship_s2="3Q26 取下限 1,200 千片；前兩季以季增率回推（Derived）",
    ut={"C": 0.79, "D": 0.85, "E": 0.92, "F": "=q4_ut"}, ut_tag="Verified", ut_s1="公司 6-K：1Q26 79%、2Q26 85%、3Q26 92%", ut_s2="FY2026 欄＝出貨合計 ÷ 產能合計",
    cap={c: f"={c}5/{c}6" for c in "CDEF"}, cap_s1="出貨 ÷ 稼動率", cap_s2="公司未揭露產能",
    fxq=FX, asp_s2="含 AI 相關營收；模型的成熟製程售價另扣 AI 相關（In 頁 asp_26）",
    gp={"B": 68.906, "C": 17.818, "D": "=40.141-C9", "E": "=65.346-40.141", "F": "=F4*q4_gm"}, gp_tag="Verified／Interested-party（4Q 指引）",
    gp_s2="1H26 毛利 401.41 億、9M26 653.46 億（6-K）", gm_s2="公司：1Q26 29.2%、2Q26 32.5%、3Q26 34.0%",
    opx={"B": "=B9-B12", "C": "=C9-C12", "D": "=D9-D12", "E": "=E9-E12", "F": "=q4_opx"},
    oi={"B": 43.850, "C": 11.276, "D": "=26.226-C12", "E": "=44.869-26.226", "F": "=F9-F11"}, oi_tag="Verified／Derived", oi_s2="1H26 262.26 億、9M26 448.69 億（6-K）",
    nonop={"B": "=B14-B12", "C": "=C14-C12", "D": "=D14-D12", "E": "=E14-E12", "F": "=q4_nonop"}, nonop_s2="稅前 − 營業利益；2Q26 含投資與股利收益約 302 億（優分析）",
    pti={"B": 49.655, "C": 16.644, "D": "=61.830-C14", "E": "=81.741-61.830", "F": "=F12+F13"}, pti_tag="Verified／Derived", pti_s2="FY2025 由 stockanalysis 稅前淨利率推得；1H26 618.30 億、9M26 817.41 億（6-K）",
    tax={"B": "=B14-B17-B18", "C": "=C14-C17-C18", "D": "=D14-D17-D18", "E": "=E14-E17-E18", "F": "=F14*q4_tax"},
    ni={"B": 41.716, "C": 16.171, "D": "=58.431-C17", "E": "=74.865-58.431", "F": "=F14-F15-F18"}, ni_tag="Verified／Derived", ni_s2="1H26 584.31 億、9M26 748.65 億（6-K）",
    mi={"B": 0, "C": "=16.117-C17", "D": "=(58.341-16.117)-D17", "E": "=(74.730-58.341)-E17", "F": 0}, mi_s2="本期淨利（含非控制權益）− 歸屬母公司；1H26 本期淨利 583.41 億、9M26 747.30 億；FY2025 未找到，取 0",
    sh={"B": "=41.716/3.34", "C": "=16.171/1.29", "D": "=42.260/3.39", "E": "=16.434/1.32", "F": "=q_sh"}, sh_s2="—",
    eps_rep={"B": 3.34, "C": 1.29, "D": 3.39, "E": 1.32}, eps_s1="公司 6-K：FY2025 3.34（stockanalysis 3.31）、1Q26 1.29、2Q26 3.39、3Q26 1.32", eps_s2="9M26 6.00")
C["q_checks"] = [("FY2026 營收年增", "=G4/B4-1", "0.0%"), ("3Q26 營收季增（公司 8.0%）", "=E4/D4-1", "0.0%"), ("FY2026 毛利率", "=G10", "0.0%"),
                 ("FY2026 營業利益率", "=G12/G4", "0.0%"), ("檢查：9M26 EPS 自算 − 公司 6.00（應約為 0）", "=SUM(C20:E20)-6.00", "0.000"),
                 ("檢查：9M26 營收 自算 − 公司 204.01（應約為 0）", "=SUM(C4:E4)-204.01", "0.000"), ("FY2026 EPS 不含 2Q26 一次性業外（業外以 3Q26 水準代入，稅率 17%；參考）", "=G20-(D13-E13)*(1-0.17)/G19", "0.00")]
C["seg"] = dict(title="官方報告分部與製程別（最近完整年度＋最近四季 4Q25–3Q26；分部原則第 3 條）",
    cols=["FY2025A", "4Q25", "1Q26", "2Q26", "3Q26", "最近四季合計"],
    rows=[("晶圓代工營收（單一主要部門）", {"B": 237.553, "C": 61.81, "D": "=C4", "E": "=D4", "F": "=E4", "G": "=SUM(C36:F36)"}, "#,##0.000", "Verified",
           "公司 6-K 各季合併營收；20-F 部門資訊以晶圓製造為主要部門", "其他部門金額：未找到（本模型併入晶圓代工）"),
          ("  22／28 奈米", {"C": "=C36*0.36", "D": "=D36*0.34", "E": "=E36*0.37", "F": "=F36*0.35", "G": "=SUM(C37:F37)"}, "#,##0.000", "Verified",
           "公司 6-K：22／28 奈米占營收 4Q25 36%、1Q26 34%、2Q26 37%、3Q26 35%", "FY2025 全年比重：未找到"),
          ("  其他製程（14 奈米與 40 奈米以上）", {c: f"={c}36-{c}37" for c in "CDEFG"}, "#,##0.000", "Derived", "差額", "40 奈米以上各製程比重：未找到"),
          ("檢查：分部合計 − 合併營收", {"B": "=B36-B4", "C": "=C36-61.81", "D": "=D36-C4", "E": "=E36-D4", "F": "=F36-E4"}, "#,##0.000", "Derived", "—", "應為 0"),
          ("檢查：子列合計 − 分部", {c: f"={c}37+{c}38-{c}36" for c in "CDEFG"}, "#,##0.000", "Derived", "—", "應為 0")])
C["scn_t"] = {"A": "情境 A（順利）：稼動率 95%／94%／93%；售價 +10%／+3%／+2%（公司估 2027 年約一成）；AI 相關 FY2027 +60%、之後 +45% × 顆數 A；Intel 12 奈米 FY2028 2.5 億、FY2029 5 億美元",
              "B": "情境 B（平平）：稼動率 90%／87%／88%；售價 +5%／0／0；AI 相關 +40%、之後 × 顆數 B（0.8／0.85）；Intel 12 奈米 1.2 億、3 億美元",
              "C": "情境 C（不如預期）：稼動率 82%／76%／80%；售價 0／−6%／−3%（中國成熟製程擴產的價格壓力）；AI 相關 +20%、之後 × 顆數 C（0.6）；Intel 12 奈米 0、0.8 億美元"}
C["units_a1"] = "Units — 聯電成熟製程晶圓（千片 12 吋約當）、稼動率、售價與新業務（三情境）。產能 FY2026 由出貨 ÷ 稼動率推得；AI 相關營收 FY2028–29 乘 Siliconomics 同字母顆數係數"
C["units_note"] = "讀法：聯電 EPS 主要由稼動率與美元售價決定；AI 相關（先進封裝矽中介層、矽光子）與 Intel 12 奈米在 FY2029 合計仍低於營收一成，影響在情境 A 的上緣。"
C["peers"] = [("台積電（TSMC）", "2330.TW", 16.81, "2026-10-22", "stockanalysis.com forward PE（分析師共識；價格日期如左）"),
              ("格芯（GlobalFoundries）", "GFS", 20.04, "2026-10-22", "stockanalysis.com forward PE（抄自台積電模型 Val 頁同業表）"),
              ("中芯國際（SMIC，港股）", "0981.HK", 37.14, "2026-10-16", "stockanalysis.com forward PE（抄自台積電模型 Val 頁同業表）"),
              ("高塔半導體（Tower）", "TSEM", 34.63, "2026-10-26", "stockanalysis.com forward PE（分析師共識；價格日期如左）"),
              ("世界先進（VIS）", "5347.TWO", 23.41, "2026-09-11", "stockanalysis.com forward PE（價格日期如左；10 月值未取得）")]
C["peer_note"] = "晶圓代工同業（台積電、格芯、中芯港股、高塔、世界先進）；不含 A 股（非 A 股公司，Andy 2026-10-25 11:53）；華虹（1347.HK）前瞻 P/E 133.5 倍且資料日 2026-08-07，未列"
C["tp_method"] = ("DCF：股東現金流＝EPS ×（1 − g／ROE），r＝10%、g＝4%、ROE 為近三年平均約 13.9%（再投資率約 29%），FY2027–32 明確期＋永續；聯電資本密集，但本模型只有折舊路徑、沒有資本支出推導，未改用淨利＋折舊－資本支出（若擴產資本支出持續高於折舊，DCF 偏高估）；情境價值負值取 0。"
                  "Comparable：台積電、格芯、中芯（港股）、高塔、世界先進前瞻 P/E 中位數，再乘相對倍數調整係數（第 0.5 節；聯電 10 年前瞻 P/E 長期低於同業，係數觸及下限 0.5）。"
                  "主要偏誤：稼動率與售價屬公司層級假設（待併入 Siliconomics）；AI 相關與 Intel 12 奈米為 Assumed 路徑；FY2030 起只用收斂法。情境對應：A＝順利（稼動率 95%、售價 +10%）、B＝平平、C＝不如預期（中國成熟製程擴產、稼動率 76–82%、售價下跌）")
C["sens_cases"] = [("fx", "匯率", "x", "0.0"), ("cap_adj", "產能係數", "x", "0.00"), ("ut_adj", "稼動率調整", "pt", "0.0%"), ("asp_adj", "售價水準調整", "%", "0.0%"),
                   ("dag_adj", "折舊年增調整", "pt", "0.0%"), ("cc_g", "每片現金成本年增", "%", "0.0%"), ("phi", "營業費用彈性", "x", "0.00"),
                   ("nonop_f", "營業外淨額", "NT$B", "0.0"), ("tax", "稅率", "%", "0%"), ("ai_adj", "AI 相關營收係數", "x", "0.00"), ("m_new", "新業務邊際貢獻率", "%", "0%")]
C["sens_read"] = "稼動率與售價是加權目標價最敏感的兩個輸入（固定成本高，稼動率每 3 個百分點、售價每 3% 都直接落到毛利）；AI 相關與 Intel 12 奈米的影響小於售價 1%。"
C["pe_cmc"] = [10.6, 11.8, 12.8, 14.1, 15.1, 10.9, 4.68, 9.38, 10.6, 14.9]
C["pe_sa"] = [None, None, None, None, None, 15.81, 5.69, 10.81, 11.45, 14.82]
C["pe_note"] = "stockanalysis 年度 P/E 只有 FY2021–25。十年中位數約 11 倍；2022 年 EPS 高峰時 P/E 只有 5–6 倍，典型的成熟製程循環股。"
C["recon"] = [("sec", "A. FY2026 錨點 vs 共識與券商"),
    ("FY2026 營收 vs 共識", "=cons_rev26", ("rev", 2026), "stockanalysis（21 位；2026-10-08）", "共識資料日早於 3Q26 財報", "#,##0.0", False),
    ("FY2026 EPS vs 共識", "=cons_eps26", ("eps", 2026), "stockanalysis（2026-10-08）", "共識區間 4.55–8.15；模型含 2Q26 一次性業外（約 2.0 元）", "#,##0.00", False),
    ("FY2026 EPS vs 美系外資中點", "=us_eps26", ("eps", 2026), "旺得富 2026-10-17", "7.46–7.68", "#,##0.00", False),
    ("4Q26 毛利率 vs 公司指引（百分點）", "=q4_gm", ("=Q!F10", 0), "公司 6-K（2026-10-28）", "三欄同值（FY2026 錨點）", "0.0%", True),
    ("sec", "B. FY2027–29 vs 公司說法與券商（校準未使用）"),
    ("FY2027 平均售價年增 vs 公司約一成（百分點）", "=co_asp27", ("aspg", 2027), "優分析 2026-10-28（法說）", "A＝公司說法；B、C 低於公司", "0.0%", True),
    ("FY2027 EPS vs 本土投顧", "=loc_eps27", ("eps", 2027), "旺得富 2026-10-17", "FY2027 共識：未找到", "#,##0.00", False),
    ("FY2029 AI 相關營收 vs 公司目標（US$B）", "=co_ai29", ("ai_usd", 2029), "科技新報 2026-10-06（大摩轉述）", "B、C 乘同字母顆數係數", "0.000", False),
    ("共識前瞻 12 月 EPS vs 模型（情境加權 2／12 × FY2027＋10／12 × FY2028）", "=cons_feps", ("feps", 0), "stockanalysis 前瞻 P/E 推得", "三欄同值（情境加權）", "#,##0.00", False)]
C["sources"] = [("公司：6-K 3Q26 財報（2026-10-28；營收、毛利率、營業利益、9M26 累計、稼動率 92%、出貨逾 120 萬片、4Q26 指引、資本預算約 1,028 億元）", K3),
    ("公司：6-K 2Q26 財報（2026-07-29；1H26 累計、稼動率 85%、22 奈米 17.5%、2026 資本支出 20 億美元、P4／台南擴產預算 1,487 億元）", K2),
    ("公司：6-K 1Q26 財報（稼動率 79%、22／28 奈米 34%）", K1), ("公司：6-K 4Q25 財報（2026-01-28；4Q25 營收 618.1 億、稼動率 78%、22／28 奈米 36%、FY2025 EPS 3.34）", K4),
    ("媒體：優分析 2026-10-28（3Q26 法說轉述：2027 年平均售價年增約一成、折舊與原物料成本上升）", UA3), ("媒體：優分析 2026-07-29（2Q26 投資與股利收益約 302 億）", UA2),
    ("資料商：Finsee 3Q26 財報摘要（9M26 折舊 468.1 億、Intel 12 奈米先前時程、AI 相關 3 億美元目標）", FS3),
    ("券商轉述：科技新報 2026-10-06（摩根士丹利：目標價 188、AI 相關 2026 年 3 億、2029 年 10 億美元；8 吋 2027 年漲價）", MS10),
    ("媒體：旺得富 2026-10-17（法人目標價 130–185、FY2026 EPS 7.46–7.68、FY2027 EPS 9.2、海外可轉債 18 億美元）", WR17),
    ("資料商：stockanalysis（股價、前瞻 P/E、共識目標價、FY2021–25 損益、ROE、年度 P/E）", SA + "；" + SAF + "；" + SAI + "；" + SAR + "；" + SAH),
    ("資料商：stockanalysis 同業前瞻 P/E（2330、GFS、0981、TSEM、5347）", "https://stockanalysis.com/quote/tpe/2330/；https://stockanalysis.com/stocks/gfs/；https://stockanalysis.com/quote/hkg/0981/；https://stockanalysis.com/stocks/tsem/；https://stockanalysis.com/quote/tpex/5347/"),
    ("資料商：companiesmarketcap 年底 P/E", CMC),
    (f"行業層：Siliconomics 快照（匯率 fx_twd、顆數係數 um28／um29；本 repo，唯讀）", "siliconomics/data/CURRENT_SNAPSHOT；siliconomics/reader.py"),
    ("待補（人工）：產能（公司未揭露）、各製程比重（40 奈米以上）、應用別與地區別營收、Fab12i P3 量產時程、Intel 12 奈米營收分配方式、AI 相關毛利率、4Q26 折舊、FY2027 起共識、2016–2020 淨利率（十年平均）、海外可轉債轉換價、10-23 以後收盤價", "未找到"),
    ("SemiAnalysis", "未使用")]

C["readme"] = [
 ("聯電（2303.TW；ADR UMC）個股模型 v0.1（AI 半導體 Project：代工層，成熟製程晶圓代工；樣板 guc/、營收結構參考 tsmc/）　2026-10-31", True),
 ("版本：v0.1（初版）。營收＝成熟製程晶圓（產能 × 稼動率 × 美元售價 × 匯率）＋AI 相關（先進封裝矽中介層、矽光子）＋Intel 12 奈米合作。滾入 3Q26 實績（2026-10-28 財報）與 4Q26 指引。本版未做先驗。", False),
 ("模型命題：聯電的獲利由成熟製程（22／28 奈米以上）的稼動率與價格循環決定。3Q26 稼動率回到 92%、毛利率 34.0%，公司估 2027 年平均售價年增約一成；AI 帶動的周邊（先進封裝矽中介層、矽光子，以及 OLED 驅動 IC、Wi-Fi 7、車用雷達等 22／28 奈米需求）與 Intel 12 奈米合作是增量，中國成熟製程擴產是價格的下行風險。", False),
 ("核心問題：稼動率與美元售價回升能走多遠（情境 A 稼動率 95%、售價 +10%），AI 相關與 Intel 合作在 FY2029 能補多少（模型情境 A 約 10 億＋5 億美元，不到營收一成）；現價約 20 倍前瞻 P/E，比聯電自己 10 年 Bloomberg 前瞻 P/E 中位數 15.2 倍高多少。", False),
 ("", False),
 ("會計年度＝日曆年，FY(t)＝CY(t)，不平移。FY2025A 實績、FY2026 錨點（1Q–3Q26 實績＋4Q26 指引）、FY2027–FY2029 驅動因子推導、FY2030–FY2032 收斂法。金額：十億新台幣；EPS：新台幣（基本）；口徑：歸屬母公司淨利（TIFRS）。晶圓：千片 12 吋約當。", False),
 ("", False),
 ("驅動因子 → 推導量（行業變數讀 Siliconomics；成熟製程稼動率與售價為公司層級，待併入 Siliconomics）", True),
 ("成熟製程晶圓：產能（FY2026 由各季出貨 ÷ 稼動率推得約 5,230 千片；FY2027–29 年增 4%／7%／7%，新加坡 P4 與台南新廠）× 稼動率（A／B／C：FY2027 95%／90%／82%、FY2028 94%／87%／76%、FY2029 93%／88%／80%）× 美元售價（FY2026 約 1,878 美元；FY2027 +10%／+5%／0、FY2028 +3%／0／−6%、FY2029 +2%／0／−3%）× 匯率 32（Siliconomics fx_twd）。", False),
 ("新業務：AI 相關 FY2026 3 億美元（公司）、FY2027 +60%／+40%／+20%、之後年增 45% 再乘 Siliconomics 同字母顆數係數（um28、um29）；Intel 12 奈米 FY2028 2.5／1.2／0 億美元、FY2029 5／3／0.8 億美元；新業務邊際貢獻率 45%（Assumed）。", False),
 ("成本：銷貨成本內折舊（FY2026 約 562 億元＝9M26 折舊年化 × 0.9；年增 12%／10%／8%）＋每片現金成本（FY2026 殘差，年增 2%）；營業費用彈性 0.5、年增下限 3%；營業外每年 40 億元（不外推 2Q26 一次性約 302 億元）、稅率 15%、股數 124.5 億股。", False),
 ("校準：FY2026 ＝ 1Q–3Q26 實績（公司 6-K）＋4Q26 公司指引（出貨 +2%、售價持穩、毛利率 34%、稼動率 92%）。成熟製程售價與每片現金成本以 FY2026 校準；不用券商數字（只在 Recon 對帳）。", False),
 ("估值：Val 第 0 節目標價＝0.5 × DCF 目標價（股東現金流＝EPS ×（1 − g／ROE），FY2027–32 明確期＋永續，r 10%、g 4%，三情境權重 0.3／0.5／0.2，情境價值負值取 0）＋0.5 × Comparable 目標價（台積電、格芯、中芯港股、高塔、世界先進前瞻 P/E 中位數 × 12 個月後前瞻 EPS，再乘相對倍數調整係數；不含 A 股）；另列 Comparable 單獨目標價、市場共識目標價與差距拆解。十年 P/E 規則值 × FY2027 EPS 保留為參考（不是目標價）。本版未做先驗，無 Andy 合理 P/E 對照。", False),
 ("預估期間：FY2027–FY2029 由驅動因子推導；FY2030 起為收斂法，非驅動因子推導（營收年增由 FY2029 線性收斂到 fade_g、淨利率收斂到 fade_m，FY2032 達到；PL 頁灰底）。十年平均淨利率未找到，B、C 依工作單以 FY2029 × 0.8 代替。", False),
 ("標記：Verified、Interested-party、Analogy、Assumed、Derived。SemiAnalysis 未使用。公司數字取自 6-K 財報；法說內容多取自媒體轉述；券商與資料商為 Interested-party。", False),
 ("", False),
 ("已知限制：(1) 公司不揭露產能、出貨片數（只說 3Q26 逾 120 萬片 12 吋約當）與售價絕對值，產能與售價為推得；(2) 稼動率與售價為公司層級假設，Siliconomics 尚無成熟製程行業層（待併入）；(3) AI 相關營收與 Intel 12 奈米的金額、毛利率、營收分配方式未揭露，皆 Assumed；(4) 2Q26 含約 302 億元投資與股利收益，FY2026 EPS 不能直接外推；(5) 海外可轉債約 18 億美元的轉換稀釋未建模；(6) 本模型只有折舊路徑、沒有資本支出推導，DCF 未扣擴產資本支出；(7) 現價為 2026-10-22 收盤（10-28 法說後股價未取得）。", False),
 ("", False),
 ("情境對照表（A／B／C 讀同字母的行業情境；公司特有變數 A 最有利、C 最不利）", True),
 ("行業：AI 相關營收 FY2028–29 乘 Siliconomics um28／um29（A 1.0／1.0、B 0.8／0.85、C 0.6／0.6）；匯率 fx_twd 32（三情境相同）。成熟製程稼動率與價格：Siliconomics 未建，公司層級（待併入 Siliconomics）。不讀 markets/equipment、markets/network（聯電無設備、網通營收）。", False),
 ("公司特有：稼動率 FY2027 95%／90%／82%、FY2028 94%／87%／76%、FY2029 93%／88%／80%；售價 FY2027 +10%／+5%／0、FY2028 +3%／0／−6%、FY2029 +2%／0／−3%（C＝中國成熟製程擴產的價格壓力）；AI 相關 FY2027 +60%／+40%／+20%；Intel 12 奈米 FY2028 2.5／1.2／0、FY2029 5／3／0.8 億美元。", False)]
C["obs"] = [
 "目標價 NT$76.4（DCF 82.5、Comparable 相對倍數調整後 70.3 各半），比現價 125.5（2026-10-22）低 39.1%，級距 −25%～−50%，投資建議放空；Comparable 單獨為 +12.1%。",
 "下檔主因有二：(1) 相對倍數調整係數 0.5——聯電 10 年 Bloomberg 前瞻 P/E 中位數 15.2 倍，只有台積電與中芯（港股）中位數的中位數 32.3 倍的 0.47（觸及下限 0.5；格芯、高塔、世界先進資料不足未納入），Comparable 因此由 140.7 減半為 70.3；若不調整，加權目標價約 111.6（−11%，中立）。(2) DCF 82.5 偏低：ROE 近三年平均 13.9%、再投資率 29%，情境 B 的 FY2027–29 EPS 只有 5.95／5.85／6.53（FY2026 的 7.36 含 2Q26 一次性投資收益約 2.0 元）。EPS 因素與倍數因素都只有 −3%，與市場共識的差距不在獲利預估，而在相對倍數調整與 DCF。",
 "稼動率與售價是承重變數：3Q26 稼動率 92%、4Q26 指引低 90% 區間，公司估 2027 年平均售價年增約一成（情境 A 照用、B 取 +5%、C 持平）；售價 ±3% 使加權目標價 ±10.2%、匯率 30.5 使 −16.5%，AI 相關營收 ±30% 只有 ±2.9%。下一個檢驗點是 2027-01 下旬的 4Q26 法說（2027 年售價、折舊與資本支出指引；日期未公告）與 Intel 12 奈米 2027 年試產進度。"]


# ======================= 共用建置程式（聯電、世界先進同一套結構；數字與文字在上方 C 設定） =======================
import sys, os, json
import openpyxl
from openpyxl.styles import Font, PatternFill
from openpyxl.workbook.defined_name import DefinedName
ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.dont_write_bytecode = True      # 不在上游資料夾寫入 __pycache__（上游唯讀）
sys.path.insert(0, os.path.join(ROOT, "tools"))
sys.path.insert(0, os.path.join(ROOT, "siliconomics"))
import reader as SIL   # Siliconomics 快照（行業變數單一來源）
import cycle_position as CP
import valuation_rule as VR

KIND = C["kind"]          # "umc"：成熟製程晶圓＋AI 相關（先進封裝、矽光子）＋Intel 12 奈米；"vis"：8 吋晶圓＋VSMC 12 吋
BLUE = Font(color="0000FF"); BOLD = Font(bold=True); HDR = PatternFill("solid", fgColor="DDEBF7")
KEY = PatternFill("solid", fgColor="FFF2CC")
wb = openpyxl.Workbook()
def hdr(ws, row, labels):
    for j, h in enumerate(labels, 1):
        c = ws.cell(row, j, h); c.font = BOLD; c.fill = HDR
SC = ("A", "B", "C")
SCN_NAME = {"A": "順利", "B": "平平", "C": "不如預期"}
SILV = f"Siliconomics {SIL.VERSION}（siliconomics/data/{SIL.FILE}）"
TBD = "待併入 Siliconomics"

# ---------------- In ----------------
IN = wb.active; IN.title = "In"
IN["A1"] = (f"In — 輸入（藍字；基準欄 E，低 D、高 F）。金額：十億新台幣（NT$B）；晶圓：千片（{C['wafer_unit']}）；售價：美元／片；匯率：新台幣／美元。"
            "公司財報與法說為 Interested-party（公司說法）或 Verified（財報數字）；券商、資料商、媒體為 Interested-party")
IN["A2"] = ("標記：Verified／Interested-party／Analogy／Assumed／Derived；Analogy、Assumed 以情境或 Sens 區間呈現，不當點估計。"
            f"行業變數讀 {SILV}；成熟製程稼動率與售價屬公司層級，說明欄標「{TBD}」。SemiAnalysis 未使用")
hdr(IN, 3, ["鍵", "項目", "單位", "低", "基準", "高", "標記", "來源1", "來源2／說明"])
rows = []; sec = lambda t: rows.append(("sec", t)); inp = lambda *a: rows.append(a)
URLNAME = C["urlname"]

sec(f"A. 行業變數（{SILV}；A／B／C 讀同字母上游，不自訂行業路徑）")
fl, fb, fh = SIL.vals("fx_twd")
inp("fx", "匯率（FY2027 起；美元售價換算）", "NT$/US$", fl, fb, fh, SIL.V["fx_twd"]["tag"], f"{SILV}｜fx_twd", "Siliconomics 共用匯率（2026–2029）；與台積電、聯發科等台股模型同值")
for s in (SC if C.get("use_um", True) else ()):
    for y in (28, 29):
        k = f"um{y}_{s}"
        inp(k, f"情境 {s}：CY20{y} AI 加速器顆數 ÷ 基準", "x", None, SIL.base(k), None, SIL.V[k]["tag"], f"{SILV}｜{k}",
            "AI 相關營收 CY2028–29 依同字母顆數係數縮放（行業情境；A＝基準）" if (s == "A" and y == 28) else "同上")

sec(f"B. 產能與稼動率（{C['wafer_unit']}；公司層級，{TBD}：Siliconomics 尚無成熟製程稼動率）")
inp("cap_26", f"FY2026 產能（{C['wafer_unit']}，千片／年）", "千片", None, "=Q!G7", None, "Derived", C["cap26_src"], C["cap26_note"])
for y, (v, s1) in C["capg"].items():
    inp(f"capg{y}", f"FY20{y} 產能年增", "%", None, v, None, "Assumed", s1, f"三情境相同（產能為公司已核准的擴產；{TBD}）")
inp("cap_adj", "產能係數（Sens 用；FY2027 起各年同乘）", "x", 0.95, 1.0, 1.05, "Assumed", "無", "±5%：擴產時程提前或延後")
for y in (27, 28, 29):
    for s in SC:
        inp(f"ut{y}_{s}", f"FY20{y} 平均稼動率（情境 {s}）", "%", None, C["ut"][y][s], None, "Assumed",
            C["ut_src"] if (s == "A" and y == 27) else "同上", (f"A 最有利、C 最不利（C＝中國成熟製程擴產的價格與份額壓力）；{TBD}" if s == "A" else TBD))
inp("ut_adj", "稼動率調整（Sens 用；FY2027 起各年加減，上限 98%）", "pt", -0.03, 0.0, 0.03, "Assumed", "無", "±3 個百分點")

sec(f"C. 平均售價（美元，每片 {C['wafer_unit']}；公司層級，{TBD}：Siliconomics 尚無成熟製程價格）")
inp("asp_26", f"FY2026 平均售價（{C['wafer_name']}；美元／片）", "US$/片", None, C["asp26_f"], None, "Derived", C["asp26_src"], "校準年；FY2027 起依情境年增")
for y in (27, 28, 29):
    for s in SC:
        inp(f"asp{y}_{s}", f"FY20{y} 平均售價年增（美元；情境 {s}）", "%", None, C["asp"][y][s], None, "Assumed",
            C["asp_src"] if (s == "A" and y == 27) else "同上", (f"A＝漲價延續；C＝中國成熟製程擴產造成價格壓力；{TBD}" if s == "A" else TBD))
inp("asp_adj", "售價水準調整（Sens 用；FY2027 起各年同乘 1＋x）", "%", -0.03, 0.0, 0.03, "Assumed", "無", "±3%")

if KIND == "umc":
    sec("D. 新業務：AI 相關（先進封裝矽中介層、矽光子）與 Intel 12 奈米合作（美元營收；毛利以邊際貢獻率計）")
    inp("ai_26", "FY2026 AI 相關營收（先進封裝、矽光子）", "US$B", None, C["ai26"], None, "Interested-party", C["ai26_src"], "公司未揭露金額與毛利率；本模型把它從晶圓營收拆出")
    for s in SC:
        inp(f"aig27_{s}", f"FY2027 AI 相關營收年增（情境 {s}）", "%", None, C["aig27"][s], None, "Assumed",
            C["ai_src2"] if s == "A" else "同上", "A 最有利、C 最不利" if s == "A" else "—")
    inp("aig_L", "FY2028–29 AI 相關營收年增（情境 A 路徑；B、C 再乘同字母顆數係數）", "%", None, C["aig_L"], None, "Assumed", C["ai_src2"], "情境 A 到 FY2029 約 US$1.0B（公司目標）")
    inp("ai_adj", "AI 相關營收係數（Sens 用；FY2027 起）", "x", 0.7, 1.0, 1.3, "Assumed", "無", "±30%")
    inp("m_new", "新業務（AI 相關、Intel 12 奈米）邊際貢獻率（毛利 ÷ 營收）", "%", 0.35, 0.45, 0.55, "Assumed", "公司未揭露；先進封裝中介層、矽光子屬特殊製程", "成熟製程晶圓的現金成本以 FY2026 校準，新業務另計")
    inp("int_27", "FY2027 Intel 12 奈米營收（試產）", "US$B", None, 0.0, None, "Assumed", C["int_src"], "三情境相同")
    for y in (28, 29):
        for s in SC:
            inp(f"int{y}_{s}", f"FY20{y} Intel 12 奈米營收（情境 {s}）", "US$B", None, C["int"][y][s], None, "Assumed",
                C["int_src"] if (s == "A" and y == 28) else "同上", "營收分配方式未找到；A 最有利、C 最不利" if s == "A" else "—")
else:
    sec("D. VSMC（新加坡 12 吋廠，與 NXP 合資，世界先進持股 60%、合併報表）：產能 × 稼動率 × 售價；毛利率路徑；非控制權益 40%")
    for y, (v, s1) in C["vcap"].items():
        inp(f"vcap{y}", f"FY20{y} VSMC 產能（12 吋，千片／年）", "千片", None, v, None, "Assumed", s1, "三情境相同（產能為公司規劃）")
    inp("vcap_adj", "VSMC 產能係數（Sens 用）", "x", 0.8, 1.0, 1.2, "Assumed", "無", "±20%：量產爬坡快慢")
    for s in SC:
        inp(f"vut_{s}", f"VSMC 稼動率（FY2027 起；情境 {s}）", "%", None, C["vut"][s], None, "Assumed",
            C["vut_src"] if s == "A" else "同上", "A 最有利、C 最不利" if s == "A" else "—")
    inp("vasp", "VSMC 平均售價（美元／12 吋片）", "US$", 1500, C["vasp"], 2100, "Assumed", C["vasp_src"], "130–40 奈米特殊製程 12 吋；公司未揭露")
    for y in (27, 28, 29):
        for s in SC:
            inp(f"vgm{y}_{s}", f"FY20{y} VSMC 毛利率（含折舊；情境 {s}）", "%", None, C["vgm"][y][s], None, "Assumed",
                C["vgm_src"] if (s == "A" and y == 27) else "同上", "量產初期折舊重；A 最有利、C 最不利" if s == "A" else "—")
    inp("vmi", "VSMC 非控制權益比例（NXP 持股）", "%", None, 0.4, None, "Verified", C["vmi_src"], "非控制權益＝VSMC 毛利 ×（1 − 稅率）× 40%（簡化：VSMC 營業費用未另計）")
    inp("ai_sh26", "FY2026 AI 相關 ÷ 營收（記憶列）", "%", None, C["ai_sh26"], None, "Interested-party", C["ai_sh_src"], "公司：全年雙位數；只作記憶列，不進營收加總")
    for s in SC:
        inp(f"ai_sh27_{s}", f"FY2027 起 AI 相關 ÷ 營收（記憶列；情境 {s}）", "%", None, C["ai_sh27"][s], None, "Assumed",
            C["ai_sh_src2"] if s == "A" else "同上", "AI 需求已反映在稼動率與售價情境；本列不進營收加總" if s == "A" else "—")

sec("E. 成本：折舊（銷貨成本內）、每片現金成本、營業費用、營業外、稅、股數")
inp("da_tot26", "FY2026 折舊與攤銷（全公司）", "NT$B", None, C["da26"], None, C["da26_tag"], C["da26_src"], C["da26_note"])
inp("da_sh", "折舊中計入銷貨成本的比例", "%", None, 0.9, None, "Assumed", "公司未揭露拆分；晶圓廠折舊多數計入銷貨成本", "—")
for y, (v, s1) in C["dag"].items():
    inp(f"dag_{y}", f"FY20{y} 銷貨成本內折舊年增", "%", None, v, None, "Assumed", s1, "三情境相同（擴產已核准）")
inp("dag_adj", "折舊年增調整（Sens 用；FY2027 起各年加減）", "pt", -0.05, 0.0, 0.05, "Assumed", "無", "±5 個百分點")
inp("cc_g", "每片現金成本年增（FY2027 起）", "%", 0.0, 0.02, 0.04, "Assumed", "原物料、電價、人力（Assumed）", "每片現金成本為 FY2026 殘差（Derived），之後依本列成長")
inp("phi", "營業費用成長對營收成長的彈性", "x", 0.3, 0.5, 0.8, "Assumed", C["opx_src"], "—")
inp("g_opx_min", "營業費用年增下限", "%", None, 0.03, None, "Assumed", "研發、新業務投入", "—")
inp("nonop_f", "營業外淨額（FY2027 起每年）", "NT$B", C["nonop"][0], C["nonop"][1], C["nonop"][2], "Assumed", C["nonop_src"], C["nonop_note"])
inp("tax", "有效稅率（FY2027 起）", "%", 0.12, 0.15, 0.20, "Assumed", C["tax_src"], "—")
inp("sh_g", "加權股數年變動", "%", None, 0.0, None, "Assumed", C["sh_src"], C["sh_note"])

sec("F. FY2026 錨點（1Q、2Q26 實績＋3Q26＋4Q26 指引或假設）")
inp("fx_26", "FY2026 平均匯率", "NT$/US$", None, C["fx26"], None, "Assumed", C["fx26_src"], "實際平均匯率：未找到")
for k, lab, unit, v, tag, s1, s2 in C["q_inputs"]:
    inp(k, lab, unit, None, v, None, tag, s1, s2)

sec("G. 公司說法、共識與券商（只作 Recon 對帳，不作校準）")
for k, lab, unit, v, s1, s2 in C["cons"]:
    inp(k, lab, unit, None, v, None, "Interested-party", s1, s2)

sec("H. 價格與市場倍數")
inp("px", f"股價（{C['px_date']} 收盤）", "NT$", None, C["px"], None, "Interested-party", C["px_src"], C["px_src2"])
inp("pe_ttm", f"近 12 月 P/E（現價 ÷ 近 12 月 EPS {C['eps_ttm']}）", "x", None, f"=px/{C['eps_ttm']}", None, "Derived", C["eps_ttm_src"], "—")
inp("fpe", C["fpe_lab"], "x", None, C["fpe"], None, "Interested-party", C["fpe_src"], "—")
inp("cons_feps", C["cfeps_lab"], "NT$", None, C["cfeps"], None, "Derived", C["cfeps_src"], "共識前瞻 EPS 的直接數字：未找到")
inp("cons_tp", "市場共識目標價（賣方平均）", "NT$", None, C["cons_tp"], None, "Interested-party", C["cons_tp_src"], C["cons_tp_src2"])
inp("nim_5", "近五年平均淨利率（FY2021–25；對照用）", "%", None, C["nim5"], None, "Derived", C["nim5_src"], "十年平均（2016–2025）：未找到（2016–2020 未取得）；收斂法依工作單改用最後一年 × 0.8")

sec("I. 目標價與收斂法參數（黃底；Assumed，Andy 可改）")
TP_KEYS = []
for s, w in (("A", 0.3), ("B", 0.5), ("C", 0.2)):
    inp(f"tp_w{s}", f"目標價情境權重 {s}（A 順利／B 平平／C 不如預期）", "%", None, w, None, "Assumed", "Andy 可改（工作單 2026-10-25 預設 A 0.3／B 0.5／C 0.2）", "三者合計應為 100%")
    TP_KEYS.append(f"tp_w{s}")
inp("tp_r", "折現率／要求報酬", "%", None, 0.10, None, "Assumed", "Andy 可改（工作單預設 10%）", "美、台、日、韓同一值")
inp("tp_g", "長期成長率", "%", None, 0.04, None, "Assumed", "Andy 可改（工作單預設 4%）", "永續成長與 FY2032 收斂目標（情境 B）")
inp("tp_roe", "ROE（近三年平均，FY2023–25）", "%", None, C["roe_f"], None, "Derived", C["roe_src"], "單一來源")
inp("tp_w_dcf", "目標價中 DCF 的權重；其餘為 Comparable", "%", None, 0.5, None, "Assumed", "Andy 可改（2026-10-25 11:53 定 DCF／Comparable 各半）", "—")
inp("fade_gA", "情境 A：FY2032 營收年增收斂目標（tp_g＋2pt）", "%", None, "=tp_g+0.02", None, "Assumed", "工作單 2026-10-25 11:20 收斂法", "FY2030 起由 FY2029 年增線性收斂")
inp("fade_gB", "情境 B：FY2032 營收年增收斂目標（＝tp_g）", "%", None, "=tp_g", None, "Assumed", "同上", "—")
inp("fade_gC", "情境 C：FY2032 營收年增收斂目標（tp_g−2pt）", "%", None, "=tp_g-0.02", None, "Assumed", "同上", "—")
inp("fade_mA", "情境 A：FY2032 淨利率收斂目標（＝情境 A FY2029 淨利率）", "%", None, 0, None, "Derived", "PL 頁情境 A FY2029（公式於建置時連結）", "—")
inp("fade_mB", "情境 B：FY2032 淨利率收斂目標（FY2029 與「FY2029 × 0.8」的中點）", "%", None, 0, None, "Assumed", "十年平均未找到，依工作單以最後一年 × 0.8 代替", "—")
inp("fade_mC", "情境 C：FY2032 淨利率收斂目標（FY2029 × 0.8）", "%", None, 0, None, "Assumed", "同上", "—")
TP_KEYS += ["tp_r", "tp_g", "tp_roe", "tp_w_dcf", "fade_gA", "fade_gB", "fade_gC", "fade_mA", "fade_mB", "fade_mC"]

r = 4; INROW = {}
for srow in rows:
    if srow[0] == "sec":
        IN.cell(r, 1, srow[1]).font = BOLD; r += 1; continue
    key, item, unit, lo, base, hi, tag, s1, s2 = srow
    assert key not in INROW, key
    IN.cell(r, 1, key); IN.cell(r, 2, item); IN.cell(r, 3, unit)
    for j, v in ((4, lo), (5, base), (6, hi)):
        if v is not None:
            IN.cell(r, j, v)
            if not (isinstance(v, str) and v.startswith("=")): IN.cell(r, j).font = BLUE
    for u_, nm_ in URLNAME.items():          # 網址只列在 Sources 頁；In 頁寫來源名稱
        s1 = s1.replace(u_, nm_); s2 = s2.replace(u_, nm_)
    IN.cell(r, 7, tag); IN.cell(r, 8, s1); IN.cell(r, 9, s2)
    wb.defined_names[key] = DefinedName(key, attr_text=f"In!$E${r}")
    if lo is not None: wb.defined_names[key + "_lo"] = DefinedName(key + "_lo", attr_text=f"In!$D${r}")
    if hi is not None: wb.defined_names[key + "_hi"] = DefinedName(key + "_hi", attr_text=f"In!$F${r}")
    INROW[key] = r; r += 1
for k in TP_KEYS:
    for c in range(1, 10): IN.cell(INROW[k], c).fill = KEY
for col, w in zip("ABCDEFGHI", [12, 58, 10, 9, 11, 9, 18, 70, 70]): IN.column_dimensions[col].width = w
IN.freeze_panes = "D4"
def defname(name, ref):
    wb.defined_names[name] = DefinedName(name, attr_text=ref)

# ---------------- Q：實績與 FY2026 錨點（NT$B） ----------------
Q = wb.create_sheet("Q")
Q["A1"] = C["q_a1"]
Q["A2"] = C["q_a2"]
hdr(Q, 3, ["項目", "FY2025A", "1Q26（實績）", "2Q26（實績）", "3Q26（" + C["q3_state"] + "）", "4Q26（模型）", "FY2026 錨點", "標記", "來源1", "來源2／說明"])
def qrow(r, label, vals, fmt="#,##0.000", tag="Verified", s1=C["q_s1"], s2="—", ann=None):
    Q.cell(r, 1, label)
    for col, v in vals.items():
        c = Q[f"{col}{r}"]; c.value = v; c.number_format = fmt
        if not (isinstance(v, str) and v.startswith("=")): c.font = BLUE
    if ann != "":
        Q[f"G{r}"] = ann or f"=SUM(C{r}:F{r})"; Q[f"G{r}"].number_format = fmt
    Q.cell(r, 8, tag); Q.cell(r, 9, s1); Q.cell(r, 10, s2)
QD = C["qd"]
qrow(4, "營收", QD["rev"], tag=QD["rev_tag"], s2=QD["rev_s2"])
qrow(5, f"晶圓出貨（千片，{C['wafer_unit']}）", QD["ship"], fmt="#,##0.0", tag=QD["ship_tag"], s1=QD["ship_s1"], s2=QD["ship_s2"])
qrow(6, "稼動率", QD["ut"], fmt="0.0%", ann="=G5/G7", tag=QD["ut_tag"], s1=QD["ut_s1"], s2=QD["ut_s2"])
qrow(7, f"產能（千片，{C['wafer_unit']}）", QD["cap"], fmt="#,##0.0", tag="Derived", s1=QD["cap_s1"], s2=QD["cap_s2"])
qrow(8, "平均售價（全部營收 ÷ 出貨；美元／片）", {c: f"={c}4/{c}5/{QD['fxq'][c]}*1000000" for c in "CDEF"}, fmt="#,##0", ann="=G4/G5/fx_26*1000000", tag="Derived", s1="營收 ÷ 出貨 ÷ 匯率", s2=QD["asp_s2"])
qrow(9, "毛利", QD["gp"], tag=QD["gp_tag"], s2=QD["gp_s2"])
qrow(10, "毛利率", {c: f"={c}9/{c}4" for c in "BCDEFG"}, fmt="0.00%", ann="", tag="Derived", s1="—", s2=QD["gm_s2"])
qrow(11, "營業費用", QD["opx"], tag="Derived／Assumed", s2="毛利 − 營業利益")
qrow(12, "營業利益", QD["oi"], tag=QD["oi_tag"], s2=QD["oi_s2"])
qrow(13, "營業外淨額", QD["nonop"], tag="Derived／Assumed", s2=QD["nonop_s2"])
qrow(14, "稅前利益", QD["pti"], tag=QD["pti_tag"], s2=QD["pti_s2"])
qrow(15, "所得稅", QD["tax"], tag="Derived／Assumed", s2="稅前 −（歸屬母公司淨利＋非控制權益）")
qrow(16, "  有效稅率", {c: f"={c}15/{c}14" for c in "BCDEFG"}, fmt="0.0%", ann="", tag="Derived", s1="—")
qrow(17, "歸屬母公司淨利", QD["ni"], tag=QD["ni_tag"], s2=QD["ni_s2"])
qrow(18, "  非控制權益", QD["mi"], tag="Derived／Assumed", s2=QD["mi_s2"])
qrow(19, "加權股數（十億股）", QD["sh"], fmt="0.00000", ann="=AVERAGE(C19:F19)", tag="Derived", s1="淨利 ÷ EPS（公司）", s2=QD["sh_s2"])
qrow(20, "EPS（基本）", {c: f"={c}17/{c}19" for c in "BCDEF"}, fmt="#,##0.00", ann="=G17/G19", tag="Derived", s1="淨利 ÷ 股數")
qrow(21, "  公司揭露 EPS", QD["eps_rep"], fmt="#,##0.00", ann="", tag="Verified", s1=QD["eps_s1"], s2=QD["eps_s2"])
Q.cell(23, 1, "FY2026 錨點的檢查（參考；校準未使用券商數字）").font = BOLD
for i, (lab, f, fm) in enumerate(C["q_checks"], 24):
    Q.cell(i, 1, lab); Q[f"B{i}"] = f; Q[f"B{i}"].number_format = fm
for nm, rr in (("rev", 4), ("gp", 9), ("opx", 11), ("oi", 12), ("nonop", 13), ("tax", 15), ("ni", 17), ("mi", 18), ("sh", 19), ("eps", 20)):
    defname(f"a25_{nm}", f"Q!$B${rr}"); defname(f"a26_{nm}", f"Q!$G${rr}")
defname("a26_ship", "Q!$G$5")
# 官方分部與子列（分部原則第 3 條；check_segments 讀本區）
SEG = C["seg"]
Q.cell(34, 1, SEG["title"]).font = BOLD
hdr(Q, 35, ["項目"] + SEG["cols"])
Q.cell(35, 8, "標記").font = BOLD; Q.cell(35, 9, "來源1").font = BOLD; Q.cell(35, 10, "來源2／說明").font = BOLD
for _r, (_lab, _vals, _fmt, _tag, _s1, _s2) in enumerate(SEG["rows"], 36):
    Q.cell(_r, 1, _lab)
    for col, v in _vals.items():
        Q[f"{col}{_r}"] = v; Q[f"{col}{_r}"].number_format = _fmt
        if not (isinstance(v, str) and v.startswith("=")): Q[f"{col}{_r}"].font = BLUE
    Q.cell(_r, 8, _tag); Q.cell(_r, 9, _s1); Q.cell(_r, 10, _s2)
Q.column_dimensions["A"].width = 56
for c in "BCDEFG": Q.column_dimensions[c].width = 14
Q.column_dimensions["H"].width = 30; Q.column_dimensions["I"].width = 50; Q.column_dimensions["J"].width = 70
Q.freeze_panes = "B4"

# ---------------- 引擎：一個情境的完整推導（PL 與 Sens 共用） ----------------
FYS = (2025, 2026, 2027, 2028, 2029); FC = {2025: "C", 2026: "D", 2027: "E", 2028: "F", 2029: "G", 2030: "H", 2031: "I", 2032: "J"}
FADE = (2030, 2031, 2032); FADEFILL = PatternFill("solid", fgColor="E7E6E6")
BU = (2026, 2027, 2028, 2029); F3 = (2027, 2028, 2029)
DEFP = dict(scn="A", fx="fx", cap_adj="cap_adj", ut_adj="ut_adj", asp_adj="asp_adj", dag_adj="dag_adj", cc_g="cc_g", phi="phi",
            nonop_f="nonop_f", tax="tax", ai_adj="ai_adj", m_new="m_new", vcap_adj="vcap_adj", vasp="vasp", fm_self=False)
FM_FACTOR = {"A": "1", "B": "0.9", "C": "0.8"}
def engine(ws, r0, title, **over):
    P = dict(DEFP); P.update(over); s = P["scn"]
    ws.cell(r0, 1, title).font = BOLD
    R = {}; r = r0 + 1
    hdr(ws, r, ["項目", "單位", "FY2025A", "FY2026（錨點）", "FY2027", "FY2028", "FY2029", "FY2030（收斂）", "FY2031（收斂）", "FY2032（收斂）", "說明"]); r += 1
    def row(key, label, unit, fs, fmt="#,##0.00", note=""):
        nonlocal r
        ws.cell(r, 1, label); ws.cell(r, 2, unit)
        for col, f in fs.items():
            ws[f"{col}{r}"] = f; ws[f"{col}{r}"].number_format = fmt
        if note: ws.cell(r, 11, note)
        R[key] = r; r += 1
    def sec_(t):
        nonlocal r
        ws.cell(r, 1, t).font = BOLD; r += 1
    yy = lambda y: y % 100
    sec_(f"A. {C['wafer_name']}：產能 × 稼動率 × 平均售價（D 欄＝FY2026 校準年）")
    row("cap", f"產能（{C['wafer_unit']}）", "千片", {"D": "=cap_26", "E": f"=D{r}*(1+capg27)*{P['cap_adj']}", "F": f"=E{r}*(1+capg28)", "G": f"=F{r}*(1+capg29)"}, "#,##0")
    row("ut", "平均稼動率", "%", {"D": "=a26_ship/cap_26", **{FC[y]: f"=MIN(0.98,ut{yy(y)}_{s}+{P['ut_adj']})" for y in F3}}, "0.0%")
    row("ship", f"晶圓出貨（{C['wafer_unit']}）", "千片", {FC[y]: f"={FC[y]}{R['cap']}*{FC[y]}{R['ut']}" for y in BU}, "#,##0")
    row("asp", f"平均售價（{C['wafer_name']}）", "US$/片", {"D": "=asp_26", "E": f"=D{r}*(1+asp27_{s})*(1+{P['asp_adj']})", "F": f"=E{r}*(1+asp28_{s})", "G": f"=F{r}*(1+asp29_{s})"}, "#,##0")
    row("aspg", "平均售價年增（美元）", "%", {FC[y]: f"={FC[y]}{R['asp']}/{FC[y - 1]}{R['asp']}-1" for y in F3}, "0.0%")
    row("fxr", "匯率", "NT$/US$", {"D": "=fx_26", **{FC[y]: f"={P['fx']}" for y in F3}})
    row("wrev", C["wafer_rev_label"], "NT$B", {FC[y]: f"={FC[y]}{R['ship']}*{FC[y]}{R['asp']}*{FC[y]}{R['fxr']}/1000000" for y in BU})
    if KIND == "umc":
        sec_("B. 新業務：AI 相關（先進封裝矽中介層、矽光子）與 Intel 12 奈米（美元營收）")
        row("ai_usd", "AI 相關營收（美元；FY2028 起 × 同字母顆數係數）", "US$B", {"D": "=ai_26", "E": f"=D{r}*(1+aig27_{s})*{P['ai_adj']}",
            "F": f"=E{r}*(1+aig_L)*um28_{s}", "G": f"=F{r}/um28_{s}*(1+aig_L)*um29_{s}"}, "0.000")
        row("ai_rev", "AI 相關營收", "NT$B", {FC[y]: f"={FC[y]}{R['ai_usd']}*{FC[y]}{R['fxr']}" for y in BU})
        row("int_usd", "Intel 12 奈米營收（美元）", "US$B", {"D": "=0", "E": "=int_27", "F": f"=int28_{s}", "G": f"=int29_{s}"}, "0.000")
        row("int_rev", "Intel 12 奈米營收", "NT$B", {FC[y]: f"={FC[y]}{R['int_usd']}*{FC[y]}{R['fxr']}" for y in BU})
        row("rev", "營收", "NT$B", {"C": "=a25_rev", **{FC[y]: f"={FC[y]}{R['wrev']}+{FC[y]}{R['ai_rev']}+{FC[y]}{R['int_rev']}" for y in BU}})
        row("aish", "AI 相關 ÷ 營收", "%", {FC[y]: f"={FC[y]}{R['ai_rev']}/{FC[y]}{R['rev']}" for y in BU}, "0.0%")
    else:
        sec_("B. VSMC（新加坡 12 吋廠；FY2027 第一季量產）")
        row("vcap", "VSMC 產能（12 吋）", "千片", {"D": "=0", **{FC[y]: f"=vcap{yy(y)}*{P['vcap_adj']}" for y in F3}}, "#,##0")
        row("vut", "VSMC 稼動率", "%", {FC[y]: f"=vut_{s}" for y in F3}, "0.0%")
        row("vship", "VSMC 出貨（12 吋）", "千片", {"D": "=0", **{FC[y]: f"={FC[y]}{R['vcap']}*{FC[y]}{R['vut']}" for y in F3}}, "#,##0")
        row("vrev", "VSMC 營收（12 吋）", "NT$B", {"D": "=0", **{FC[y]: f"={FC[y]}{R['vship']}*{P['vasp']}*{FC[y]}{R['fxr']}/1000000" for y in F3}})
        row("rev", "營收", "NT$B", {"C": "=a25_rev", **{FC[y]: f"={FC[y]}{R['wrev']}+{FC[y]}{R['vrev']}" for y in BU}})
        row("v12sh", "12 吋（VSMC）÷ 營收", "%", {FC[y]: f"={FC[y]}{R['vrev']}/{FC[y]}{R['rev']}" for y in BU}, "0.0%")
        row("aish", "AI 相關 ÷ 營收（記憶列，不進營收加總）", "%", {"D": "=ai_sh26", **{FC[y]: f"=ai_sh27_{s}" for y in F3}}, "0.0%")
    row("g", "營收年增", "%", {FC[y]: f"={FC[y]}{R['rev']}/{FC[y - 1]}{R['rev']}-1" for y in FYS[1:]}, "0.0%")
    sec_("C. 成本與損益")
    row("da", "銷貨成本內折舊", "NT$B", {"D": "=da_tot26*da_sh", **{FC[y]: f"={FC[y - 1]}{r}*(1+dag_{yy(y)}+{P['dag_adj']})" for y in F3}})
    if KIND == "umc":
        cal = f"=(D{R['rev']}-a26_gp-D{R['da']}-(1-{P['m_new']})*(D{R['ai_rev']}+D{R['int_rev']}))/D{R['ship']}*1000"
    else:
        cal = f"=(D{R['wrev']}-a26_gp-D{R['da']})/D{R['ship']}*1000"
    row("ccw", "每片現金成本（FY2026 殘差，之後年增 cc_g）", "NT$K/片", {"D": cal, **{FC[y]: f"={FC[y - 1]}{r}*(1+{P['cc_g']})" for y in F3}}, "#,##0.00",
        note="Derived：讓 FY2026 毛利等於錨點；含原物料、人力、電力等變動與固定現金成本")
    row("cogs", "現金成本合計（每片 × 出貨）", "NT$B", {FC[y]: f"={FC[y]}{R['ccw']}*{FC[y]}{R['ship']}/1000" for y in BU})
    if KIND == "umc":
        row("gpn", "新業務毛利（邊際貢獻率 × AI 相關與 Intel 營收）", "NT$B", {FC[y]: f"={P['m_new']}*({FC[y]}{R['ai_rev']}+{FC[y]}{R['int_rev']})" for y in BU})
    else:
        row("vgm", "VSMC 毛利率", "%", {FC[y]: f"=vgm{yy(y)}_{s}" for y in F3}, "0.0%")
        row("gpn", "VSMC 毛利", "NT$B", {"D": "=0", **{FC[y]: f"={FC[y]}{R['vrev']}*{FC[y]}{R['vgm']}" for y in F3}})
    row("gp", "毛利", "NT$B", {"C": "=a25_gp", **{FC[y]: f"={FC[y]}{R['wrev']}-{FC[y]}{R['da']}-{FC[y]}{R['cogs']}+{FC[y]}{R['gpn']}" for y in BU}})
    row("gm", "毛利率", "%", {FC[y]: f"={FC[y]}{R['gp']}/{FC[y]}{R['rev']}" for y in FYS}, "0.00%")
    row("opx", "營業費用", "NT$B", {"C": "=a25_opx", "D": "=a26_opx", **{FC[y]: f"={FC[y - 1]}{r}*(1+MAX(g_opx_min,{P['phi']}*{FC[y]}{R['g']}))" for y in F3}})
    row("oi", "營業利益", "NT$B", {FC[y]: f"={FC[y]}{R['gp']}-{FC[y]}{R['opx']}" for y in FYS})
    row("oim", "營業利益率", "%", {FC[y]: f"={FC[y]}{R['oi']}/{FC[y]}{R['rev']}" for y in FYS}, "0.00%")
    row("nonop", "營業外淨額", "NT$B", {"C": "=a25_nonop", "D": "=a26_nonop", **{FC[y]: f"={P['nonop_f']}" for y in F3}})
    row("pti", "稅前利益", "NT$B", {FC[y]: f"={FC[y]}{R['oi']}+{FC[y]}{R['nonop']}" for y in FYS})
    row("tax", "所得稅", "NT$B", {"C": "=a25_tax", "D": "=a26_tax", **{FC[y]: f"={FC[y]}{R['pti']}*{P['tax']}" for y in F3}})
    if KIND == "umc":
        row("mi", "非控制權益", "NT$B", {"C": "=a25_mi", "D": "=a26_mi", **{FC[y]: "=0" for y in F3}}, note="FY2027 起假設為 0（歷史上接近 0）")
    else:
        row("mi", "非控制權益（NXP 占 VSMC 40%）", "NT$B", {"C": "=a25_mi", "D": "=a26_mi", **{FC[y]: f"=vmi*{FC[y]}{R['gpn']}*(1-{P['tax']})" for y in F3}},
            note="VSMC 虧損時為負值（母公司淨利少扣）")
    row("ni", "歸屬母公司淨利", "NT$B", {FC[y]: f"={FC[y]}{R['pti']}-{FC[y]}{R['tax']}-{FC[y]}{R['mi']}" for y in FYS})
    row("shr", "加權股數", "十億股", {"C": "=a25_sh", "D": "=a26_sh", **{FC[y]: (f"=q_sh*(1+sh_g)" if y == 2027 else f"={FC[y - 1]}{r}*(1+sh_g)") for y in F3}}, "0.00000")
    row("eps", "EPS（基本）", "NT$", {FC[y]: f"={FC[y]}{R['ni']}/{FC[y]}{R['shr']}" for y in FYS})
    row("epsg", "EPS 年增", "%", {FC[y]: f"={FC[y]}{R['eps']}/{FC[y - 1]}{R['eps']}-1" for y in FYS[1:]}, "0.0%")
    row("nim", "淨利率（歸屬母公司淨利 ÷ 營收）", "%", {FC[y]: f"={FC[y]}{R['ni']}/{FC[y]}{R['rev']}" for y in FYS}, "0.00%")
    fg = f"fade_g{s}"
    fm = f"$G${R['nim']}*{FM_FACTOR[s]}" if P["fm_self"] else f"fade_m{s}"
    for y in FADE:      # FY2030–32：收斂法（上游只到 2029；非驅動因子推導）
        c, pc, k = FC[y], FC[y - 1], y - 2029
        ws[f"{c}{R['g']}"] = f"=$G{R['g']}+({fg}-$G{R['g']})*{k}/3"
        ws[f"{c}{R['rev']}"] = f"={pc}{R['rev']}*(1+{c}{R['g']})"
        ws[f"{c}{R['nim']}"] = f"=$G{R['nim']}+({fm}-$G{R['nim']})*{k}/3"
        ws[f"{c}{R['ni']}"] = f"={c}{R['rev']}*{c}{R['nim']}"
        ws[f"{c}{R['shr']}"] = f"={pc}{R['shr']}*(1+sh_g)"
        ws[f"{c}{R['eps']}"] = f"={c}{R['ni']}/{c}{R['shr']}"
        ws[f"{c}{R['epsg']}"] = f"={c}{R['eps']}/{pc}{R['eps']}-1"
        for key, fmt_ in (("g", "0.0%"), ("rev", "#,##0.00"), ("nim", "0.00%"), ("ni", "#,##0.00"), ("shr", "0.00000"), ("eps", "#,##0.00"), ("epsg", "0.0%")):
            ws[f"{c}{R[key]}"].number_format = fmt_; ws[f"{c}{R[key]}"].fill = FADEFILL
    ws.cell(R["rev"], 11, "FY2030 起為收斂法：營收年增由 FY2029 線性收斂到 fade_g、淨利率收斂到 fade_m（FY2032 達到），非驅動因子推導")
    R["end"] = r
    return R

# ---------------- PL ----------------
PL = wb.create_sheet("PL")
PL["A1"] = (f"PL — {C['name']}損益（十億新台幣；EPS 新台幣、基本；歸屬母公司）。FY2025A 實績、FY2026 錨點（三情境相同），FY2027–FY2029 由驅動因子推導，"
            "FY2030–FY2032 為收斂法。FY(t)＝CY(t)。三情境各一個完整推導區塊（本頁下方），頂端為摘要")
PB = {}; r0 = 60
for s in SC:
    PB[s] = engine(PL, r0, C["scn_t"][s], scn=s)
    r0 = PB[s]["end"] + 2
hdr(PL, 3, ["情境／項目", "單位", "FY2025A", "FY2026（錨點）", "FY2027", "FY2028", "FY2029", "FY2030（收斂）", "FY2031（收斂）", "FY2032（收斂）", "說明"])
sr = 4; SUMR = {}
SUMK = [("rev", "營收", "NT$B", "#,##0.0"), ("g", "  營收年增", "%", "0.0%"), ("wrev", "  " + C["wafer_rev_label"], "NT$B", "#,##0.0")]
SUMK += ([("ai_rev", "  AI 相關營收", "NT$B", "#,##0.0"), ("int_rev", "  Intel 12 奈米營收", "NT$B", "#,##0.0")] if KIND == "umc"
         else [("vrev", "  VSMC 營收（12 吋）", "NT$B", "#,##0.0"), ("aish", "  AI 相關 ÷ 營收（記憶列）", "%", "0.0%")])
SUMK += [("ut", "稼動率", "%", "0.0%"), ("asp", "平均售價（美元）", "US$/片", "#,##0"),
         ("gm", "毛利率", "%", "0.00%"), ("oim", "營業利益率", "%", "0.00%"), ("ni", "歸屬母公司淨利", "NT$B", "#,##0.0"), ("nim", "  淨利率", "%", "0.00%"),
         ("eps", "EPS", "NT$", "#,##0.00"), ("epsg", "  EPS 年增", "%", "0.0%")]
LONGK = ("rev", "g", "ni", "nim", "eps", "epsg")
NOC = {"ut", "asp", "wrev", "ai_rev", "int_rev", "vrev", "aish"}   # FY2025A 無驅動因子拆分
for s in SC:
    PL.cell(sr, 1, f"情境 {s}").font = BOLD; sr += 1
    for key, lab, unit, fmt in SUMK:
        PL.cell(sr, 1, lab); PL.cell(sr, 2, unit)
        for y in FYS + (FADE if key in LONGK else ()):
            if key in ("g", "epsg") and y == 2025: continue
            if key in NOC and y == 2025: continue
            PL[f"{FC[y]}{sr}"] = f"={FC[y]}{PB[s][key]}"; PL[f"{FC[y]}{sr}"].number_format = fmt
            if y in FADE: PL[f"{FC[y]}{sr}"].fill = FADEFILL
        SUMR[(s, key)] = sr; sr += 1
    sr += 1
IN.cell(INROW["fade_mA"], 5, f"=PL!G{SUMR[('A', 'nim')]}")
IN.cell(INROW["fade_mB"], 5, f"=PL!G{SUMR[('B', 'nim')]}*0.9")
IN.cell(INROW["fade_mC"], 5, f"=PL!G{SUMR[('C', 'nim')]}*0.8")
for k in ("fade_mA", "fade_mB", "fade_mC"): IN.cell(INROW[k], 5).font = Font(color="000000")
PL.cell(sr, 1, "核對（應為 0）：FY2026 模型 − 錨點").font = BOLD; sr += 1
CHK = []
for lab, f in (("  三情境：營收差（NT$B）", "+".join(f"ABS(D{PB[s]['rev']}-a26_rev)" for s in SC)),
               ("  三情境：毛利差（NT$B）", "+".join(f"ABS(D{PB[s]['gp']}-a26_gp)" for s in SC)),
               ("  三情境：EPS 差（NT$）", "+".join(f"ABS(D{PB[s]['eps']}-a26_eps)" for s in SC))):
    PL.cell(sr, 1, lab); PL[f"D{sr}"] = "=" + f; PL[f"D{sr}"].number_format = "0.0000"; CHK.append(sr); sr += 1
assert sr < 60, sr
PL.column_dimensions["A"].width = 62; PL.column_dimensions["B"].width = 10; PL.column_dimensions["K"].width = 70
for c in "CDEFGHIJ": PL.column_dimensions[c].width = 13
PL.cell(2, 1, "灰底＝FY2030–FY2032 收斂法（上游顆數只到 CY2029）：營收年增線性收斂到 fade_g、淨利率線性收斂到 fade_m（In 頁 I 節），非驅動因子推導")

# ---------------- Units ----------------
UN = wb.create_sheet("Units")
UN["A1"] = C["units_a1"]
cyc = lambda y: chr(ord("C") + y - 2026)
hdr(UN, 3, ["項目", "單位", "FY2026", "FY2027", "FY2028", "FY2029", "說明"])
ur = 4; UR = {}
UN.cell(ur, 1, f"1. {C['wafer_name']}：三情境稼動率、出貨與售價").font = BOLD; ur += 1
for key, lab, unit, fmt in (("ut", "稼動率", "%", "0.0%"), ("ship", "晶圓出貨", "千片", "#,##0"), ("asp", "平均售價（美元）", "US$/片", "#,##0"), ("wrev", "晶圓營收", "NT$B", "#,##0.0")):
    for s in SC:
        UN.cell(ur, 1, f"{lab}（情境 {s}）"); UN.cell(ur, 2, unit); UR[f"{key}_{s}"] = ur
        for y in BU:
            UN[f"{cyc(y)}{ur}"] = f"=PL!{FC[y]}{PB[s][key]}"; UN[f"{cyc(y)}{ur}"].number_format = fmt
        UN.cell(ur, 7, "PL 頁各情境推導區塊" if s == "A" else "同上"); ur += 1
ur += 1
UN.cell(ur, 1, "2. " + ("新業務：AI 相關與 Intel 12 奈米（三情境）" if KIND == "umc" else "VSMC 12 吋廠與 AI 相關（三情境）")).font = BOLD; ur += 1
NEWK = (("ai_usd", "AI 相關營收（美元）", "US$B", "0.000"), ("aish", "AI 相關 ÷ 營收", "%", "0.0%"), ("int_usd", "Intel 12 奈米營收（美元）", "US$B", "0.000")) if KIND == "umc" else \
       (("vship", "VSMC 出貨（12 吋）", "千片", "#,##0"), ("vrev", "VSMC 營收", "NT$B", "#,##0.0"), ("v12sh", "12 吋（VSMC）÷ 營收", "%", "0.0%"), ("aish", "AI 相關 ÷ 營收（記憶列）", "%", "0.0%"))
for key, lab, unit, fmt in NEWK:
    for s in SC:
        UN.cell(ur, 1, f"{lab}（情境 {s}）"); UN.cell(ur, 2, unit); UR[f"{key}_{s}"] = ur
        for y in BU:
            UN[f"{cyc(y)}{ur}"] = f"=PL!{FC[y]}{PB[s][key]}"; UN[f"{cyc(y)}{ur}"].number_format = fmt
        UN.cell(ur, 7, "PL 頁各情境推導區塊" if s == "A" else "同上"); ur += 1
ur += 1
UN.cell(ur, 1, "3. 校準值（只用 FY2026 已發生的實績與錨點；不用共識、不用 FY2027 起的目標）").font = BOLD; ur += 1
for key, lab, f, fmt, note in (("asp26", "FY2026 平均售價（美元／片）", f"=PL!D{PB['A']['asp']}", "#,##0", "晶圓營收 ÷ 出貨 ÷ 匯率"),
                               ("ccw", "FY2026 每片現金成本（NT$K）", f"=PL!D{PB['A']['ccw']}", "#,##0.00", "讓 FY2026 毛利等於錨點"),
                               ("da", "FY2026 銷貨成本內折舊（NT$B）", f"=PL!D{PB['A']['da']}", "#,##0.00", "全公司折舊 × da_sh")):
    UN.cell(ur, 1, lab); UN[f"C{ur}"] = f; UN[f"C{ur}"].number_format = fmt; UN.cell(ur, 7, note); UR[key] = ur; ur += 1
ur += 1
UN.cell(ur, 1, C["units_note"]).font = BOLD
UN_NOTE = ur
UN.column_dimensions["A"].width = 56; UN.column_dimensions["B"].width = 12
for c in "CDEF": UN.column_dimensions[c].width = 13
UN.column_dimensions["G"].width = 56

# ---------------- Val ----------------
V = wb.create_sheet("Val")
V["A1"] = "Val — P/E 估值（新台幣）。結論用「現價 ÷ FY2027、FY2028 EPS（模型）」；目標價＝DCF 與 Comparable 加權（第 0 節）；十年 P/E 規則值（第 3 節，年底落後 P/E）只作參考"
V["A2"] = f"循環位置：{C['name']}不是設備股，不套設備市場見頂年；以本模型營收最高的會計年度為見頂年（比照創意、聯發科；FY(t)＝CY(t)）。營收到 FY2029 仍在增加時，見頂年未知，判為中段"
O = 51
V.cell(O, 1, "1. 價格與市場隱含倍數").font = BOLD
hdr(V, O + 1, ["項目", "A", "B", "C", "說明"])
V.cell(O + 2, 1, f"現價（{C['px_date']} 收盤）"); V[f"B{O + 2}"] = "=px"; V[f"B{O + 2}"].number_format = "#,##0.00"
for k, y, lab in ((3, 2026, "現價 ÷ FY2026 EPS（錨點）"), (4, 2027, "現價 ÷ FY2027 EPS（模型）"), (5, 2028, "現價 ÷ FY2028 EPS（模型）"), (6, 2029, "現價 ÷ FY2029 EPS（模型）")):
    V.cell(O + k, 1, lab)
    for j, s in enumerate(SC):
        c = "BCD"[j]; V[f"{c}{O + k}"] = f"=px/PL!{FC[y]}{SUMR[(s, 'eps')]}"; V[f"{c}{O + k}"].number_format = "0.0"
V.cell(O + 7, 1, "未來 12 月 EPS（模型 B）≈ 0.2 × FY2026＋0.8 × FY2027"); V[f"B{O + 7}"] = f"=0.2*PL!D{SUMR[('B', 'eps')]}+0.8*PL!E{SUMR[('B', 'eps')]}"; V[f"B{O + 7}"].number_format = "0.00"
V.cell(O + 8, 1, "  現價 ÷ 未來 12 月 EPS（模型 B）"); V[f"B{O + 8}"] = f"=px/B{O + 7}"; V[f"B{O + 8}"].number_format = "0.0"
V.cell(O + 9, 1, "近 12 月 P/E（現價 ÷ 近 12 月 EPS）"); V[f"B{O + 9}"] = "=pe_ttm"; V[f"B{O + 9}"].number_format = "0.0"
V.cell(O + 10, 1, C["fpe_lab"]); V[f"B{O + 10}"] = "=fpe"; V[f"B{O + 10}"].number_format = "0.0"
V[f"E{O + 3}"] = "FY2026 三情境相同（錨點）"; V[f"E{O + 4}"] = "FY2027 三情境差在稼動率、售價與新業務；顆數情境係數從 CY2028 開始"
CY0 = O + 12
V.cell(CY0, 1, "2. 循環位置（本模型營收最高年；FY(t)＝CY(t)）").font = BOLD
hdr(V, CY0 + 1, ["情境", "營收最高年（FY2026–FY2029）", "FY2027 判定", "規則倍數（P/E）", "見頂年是否未知", "說明"])
CYR = {}
for i, s in enumerate(SC, CY0 + 2):
    rr = f"PL!D{PB[s]['rev']}:G{PB[s]['rev']}"
    V.cell(i, 1, s); CYR[s] = i
    V.cell(i, 2, f"=INDEX({{2026,2027,2028,2029}},MATCH(MAX({rr}),{rr},0))")
    V.cell(i, 3, '=IF(B{0}=2029,"中段",'.format(i) + CP.excel_rule("2027", f"B{i}") + ")")
    V.cell(i, 4, f'=IF(C{i}="見頂",pe_peak,IF(C{i}="谷底",pe_trough,pe_mid))').number_format = "0.0"
    V.cell(i, 5, f'=IF(B{i}=2029,"是：營收到觀測期末仍在增加","否")')
    V.cell(i, 6, "tools/cycle_position.py excel_rule；營收最高年取代設備市場見頂年" if s == "A" else "同上")
R30 = CY0 + 6
V.cell(R30, 1, "3. 規則值隱含股價（新台幣；約 FY2027 年底，未折現）＝ FY2027 EPS × 規則倍數").font = BOLD
hdr(V, R30 + 1, ["情境", "FY2027 EPS", "（同左）", "規則倍數", "隱含股價", "相對現價", "現價 ÷ FY2027 EPS"])
VAL = {}
for j, s in enumerate(SC):
    i = R30 + 2 + j
    V.cell(i, 1, s)
    V.cell(i, 2, f"=PL!E{SUMR[(s, 'eps')]}").number_format = "0.00"
    V.cell(i, 3, f"=B{i}").number_format = "0.00"
    V.cell(i, 4, f"=D{CYR[s]}").number_format = "0.0"
    V.cell(i, 5, f"=C{i}*D{i}").number_format = "#,##0.00"
    V.cell(i, 6, f"=E{i}/px-1").number_format = "0%"
    V.cell(i, 7, f"=px/C{i}").number_format = "0.0"
    for c in range(1, 8): V.cell(i, c).fill = KEY
    VAL[s] = i
W0 = R30 + 6
V.cell(W0, 1, "4. What must be true：FY2027 年底價 ＝ 現價 ×（1＋報酬）所需的倍數（FY2028 EPS；約 2027 年底的價格看下一年 EPS）").font = BOLD
hdr(V, W0 + 1, ["FY2027 年底價 ÷ 現價", "所需倍數（÷ FY2028 EPS，A）", "所需倍數（÷ FY2028 EPS，B）", "所需倍數（÷ FY2028 EPS，C）", "對照：十年中位數"])
vr = W0 + 2
for x in (1.0, 1.1, 1.21, 1.5, 2.0):
    V.cell(vr, 1, x).font = BLUE; V.cell(vr, 1).number_format = "0.00"
    for j, s in enumerate(SC):
        V.cell(vr, 2 + j, f"=px*A{vr}/PL!F{SUMR[(s, 'eps')]}").number_format = "0.0"
    V.cell(vr, 5, "=pe_mid").number_format = "0.0"
    vr += 1
V.cell(vr, 1, "讀法：1.00＝價格不變、1.21＝一年約 21%。所需倍數低於歷史中位數時，現價相對便宜；倍數是否合理，另看前瞻 P/E 歷史（Bloomberg）。本版未做先驗，無 Andy 合理 P/E 對照。")
WMT_NOTE = vr
T10 = vr + 2
V.cell(T10, 1, "5. 十年 P/E（2016–2025 年底；兩來源平均，分位排名用；無負值或 >200 需剔除）").font = BOLD
hdr(V, T10 + 1, ["年底", "兩來源平均 P/E", "companiesmarketcap", "stockanalysis"])
RB0 = T10 + 16
for i, y in enumerate(range(2016, 2026)):
    rr = T10 + 2 + i
    V.cell(rr, 1, str(y))
    V.cell(rr, 2, f'=IF(COUNT(C{rr}:D{rr})=0,"",AVERAGE(C{rr}:D{rr}))').number_format = "0.00"
    V.cell(rr, 3, f'=IF(B{RB0 + 5 + i}="","",B{RB0 + 5 + i})').number_format = "0.00"
    V.cell(rr, 4, f'=IF(C{RB0 + 5 + i}="","",C{RB0 + 5 + i})').number_format = "0.00"
PE10 = f"$B${T10 + 2}:$B${T10 + 11}"
V.cell(T10 + 12, 1, "年數"); V[f"B{T10 + 12}"] = f"=COUNT({PE10})"
block_end = VR.add_rule_block(V, RB0, [{"id": "P/E", "name": "P/E", "sources": [
    {"label": "companiesmarketcap（年底 P/E）", "vals": C["pe_cmc"], "tag": "Interested-party（資料商；2016–2025 共 10 年）"},
    {"label": "stockanalysis（年度 P/E）", "vals": C["pe_sa"], "tag": "Interested-party（資料商；只有 FY2021–25，2016–2020 未找到）"}],
    "cases": [
        {"case": "見頂", "pos": "25 分位", "q": 25, "basis_label": "FY2027 EPS（B）", "basis": f"=C{VAL['B']}"},
        {"case": "中段", "pos": "中位數", "q": 50, "basis_label": "FY2027 EPS（B）", "basis": f"=C{VAL['B']}"},
        {"case": "谷底", "pos": "75 分位", "q": 75, "basis_label": "FY2027 EPS（B）", "basis": f"=C{VAL['B']}"}]}],
    "0.00", "規則（Andy 2026-10-13）：合理倍數＝2016–2025 年底 P/E 依循環位置取分位，兩個來源分別取分位再平均。" + C["pe_note"])
assert V.cell(RB0 + 5, 1).value == "2016" and V.cell(RB0 + 14, 1).value == "2025", (V.cell(RB0 + 5, 1).value, V.cell(RB0 + 14, 1).value)
rule_rows = [rw for rw in range(RB0, block_end) if str(V.cell(rw, 1).value or "").startswith("規則值：")]
assert len(rule_rows) == 3, rule_rows
for nm, rw in zip(("pe_peak", "pe_mid", "pe_trough"), rule_rows):
    defname(nm, f"Val!$C${rw}")

# 0. 目標價（Andy 2026-10-25 11:28／11:53／12:22：DCF 與 Comparable 各半；DCF 負值取 0；列標籤依工作單逐字）
YRS6 = (2027, 2028, 2029, 2030, 2031, 2032)
V.cell(4, 1, "0. 目標價").font = BOLD
V.cell(5, 1, "0.1 DCF（股東權益現金流，FY2027–FY2032 明確期＋FY2032 後永續）").font = BOLD
hdr(V, 6, ["項目", "A（順利）", "B（平平）", "C（不如預期）", "說明"])
T = {}
def tset(r, key, label, fs, fmt, note=""):
    V.cell(r, 1, label); T[key] = r
    for c, f in zip("BCD", fs):
        if f is None: continue
        V[f"{c}{r}"] = f; V[f"{c}{r}"].number_format = fmt
    if note: V[f"E{r}"] = note
V.cell(7, 1, "情境 EPS（FY2027–FY2032）"); V["E7"] = "PL 頁各情境；FY2030 起為收斂法"; T["eps_h"] = 7
EPSR = {}
for i, y in enumerate(YRS6):
    tset(8 + i, f"eps{y}", f"  FY{y}", [f"=PL!{FC[y]}{SUMR[(s, 'eps')]}" for s in SC], "#,##0.00"); EPSR[y] = 8 + i
V.cell(14, 1, "股東現金流＝EPS ×（1 − tp_g ÷ tp_roe）"); T["cf_h"] = 14
tset(15, "reinv", "  再投資率（tp_g ÷ tp_roe，上限 0.8）", ["=MIN(tp_g/tp_roe,0.8)", None, None], "0.0%", "ROE 為近三年平均（In 頁 tp_roe）")
for i, y in enumerate(YRS6):
    tset(16 + i, f"cf{y}", f"  FY{y}", [f"={c}{EPSR[y]}*(1-$B$15)" for c in "BCD"], "#,##0.00")
def dcf_of(eps_refs, r="tp_r", g="tp_g", reinv="MIN(tp_g/tp_roe,0.8)"):
    """eps_refs：FY2027–32 六格 EPS 參照；回傳單一情境 DCF 每股價值（基準日 2026 年底；負值取 0）"""
    cf = lambda x: f"{x}*(1-{reinv})"
    pv = "+".join(f"{cf(x)}/(1+{r})^{t}" for t, x in enumerate(eps_refs, 1))
    return f"MAX(0,{pv}+{cf(eps_refs[-1])}*(1+{g})/({r}-{g})/(1+{r})^6)"
tset(22, "dcfv", "DCF 情境每股價值（折現率 tp_r）", [f"=MAX(0,SUMPRODUCT({c}16:{c}21,1/(1+tp_r)^{{1;2;3;4;5;6}})+{c}21*(1+tp_g)/(tp_r-tp_g)/(1+tp_r)^6)" for c in "BCD"], "#,##0.00",
     "Σ 明確期現金流 ÷ (1＋r)^t ＋ 永續價值 ÷ (1＋r)^6；永續價值＝FY2032 現金流 ×(1＋g)÷(r−g)；基準日 2026 年底；負值取 0")
tset(23, "dcf12", "  DCF 情境價值推到 12 個月後（×(1＋tp_r)）", [f"={c}22*(1+tp_r)" for c in "BCD"], "#,##0.00")
tset(24, "dcftp", "DCF 目標價（情境加權）", ["=tp_wA*B23+tp_wB*C23+tp_wC*D23", None, None], "#,##0.00", "Σ 權重 × 情境價值 ×(1＋r)；權重見 In 頁")
V.cell(26, 1, "0.2 Comparable（市場主流同業前瞻 P/E）").font = BOLD
hdr(V, 27, ["同業與前瞻 P/E", "代號", "前瞻 12 月 P/E", "日期", "來源", "標記"])
PEERS = C["peers"]
for i, (nm, tk, pe, dt, src) in enumerate(PEERS, 28):
    V.cell(i, 1, nm); V.cell(i, 2, tk); c = V.cell(i, 3, pe); c.font = BLUE; c.number_format = "0.0"
    V.cell(i, 4, dt); V.cell(i, 5, src); V.cell(i, 6, "Interested-party")
PEER_LAST = 27 + len(PEERS)
assert PEER_LAST <= 32
tset(33, "pmed", "同業前瞻 P/E 中位數", [f"=MEDIAN(C28:C{PEER_LAST})", None, None], "0.0", C["peer_note"])
tset(34, "feps", "12 個月後的前瞻 EPS（情境加權）", [f"=tp_wA*(2/12*B{EPSR[2027]}+10/12*B{EPSR[2028]})+tp_wB*(2/12*C{EPSR[2027]}+10/12*C{EPSR[2028]})+tp_wC*(2/12*D{EPSR[2027]}+10/12*D{EPSR[2028]})", None, None], "#,##0.00",
     "約 2027-10 時的未來 12 月 EPS ≈ 2／12 × FY2027＋10／12 × FY2028（會計年度＝日曆年）")
tset(35, "comptp", "Comparable 目標價", ["=B33*B34", None, None], "#,##0.00", "同業中位數 × 前瞻 EPS")
V.cell(37, 1, "0.3 結論").font = BOLD
hdr(V, 38, ["項目", "數值", "", "", "說明"])
tset(39, "tpw", "目標價（DCF 與 Comparable 加權）", ["=tp_w_dcf*B24+(1-tp_w_dcf)*B35", None, None], "#,##0.00", "tp_w_dcf × DCF＋（1 − tp_w_dcf）× Comparable（相對倍數調整後，見第 0.5 節）")
tset(40, "comp2", "Comparable 目標價（市場主流同業法）", ["=B35", None, None], "#,##0.00")
tset(41, "cons", "市場共識目標價", ["=cons_tp", None, None], "#,##0.00", C["cons_tp_src"])
tset(42, "upw", "上檔空間：加權目標價", ["=B39/px-1", None, None], "0.0%", f"現價＝In 頁 px（{C['px_date']} 收盤）")
tset(43, "upc", "上檔空間：Comparable 目標價", ["=B40/px-1", None, None], "0.0%")
tset(44, "geps", "與市場共識目標價的差距：EPS 因素", ["=B34/cons_feps-1", None, None], "0.0%", "本模型前瞻 EPS ÷ 共識前瞻 EPS − 1；共識前瞻 EPS 見 In 頁 cons_feps")
tset(45, "gmul", "與市場共識目標價的差距：倍數因素", ["=(cons_tp/cons_feps)/B33-1", None, None], "0.0%")
V["E45"] = '="共識目標價隱含 "&TEXT(cons_tp/cons_feps,"0.0")&"x，同業中位數 "&TEXT(B33,"0.0")&"x"'
tset(46, "band", "價格帶（DCF 不如預期～順利）", ['=TEXT(D23,"#,##0")&"～"&TEXT(B23,"#,##0")', None, None], "@", "C 情境價值～A 情境價值（推到 12 個月後）")
tset(47, "pr", "歷史倍數檢查：現價隱含 P/E 的十年分位（A／B／C）", [f"=IF(px/{c}{EPSR[2028]}>=MAX({PE10}),1,IF(px/{c}{EPSR[2028]}<=MIN({PE10}),0,PERCENTRANK({PE10},px/{c}{EPSR[2028]})))" for c in "BCD"], "0%",
     "現價 ÷ 各情境 FY2028 EPS 在 2016–2025 年底 P/E（第 5 節兩來源平均）的分位；高於十年最大值記為 100%、低於最小值記為 0%；只作檢查，不進目標價")
V.cell(48, 1, "目標價方法")
V["B48"] = C["tp_method"]
for r_ in (24, 35, 39):
    V[f"B{r_}"].fill = KEY
V.column_dimensions["A"].width = 58
for c in "BCDEFGH": V.column_dimensions[c].width = 16

# ---------------- Sens ----------------
SE = wb.create_sheet("Sens")
SE["A1"] = ("Sens — 一次動一個輸入（In 頁敏感度低／高值），三情境都在本頁下方逐案重算（FY2026 恆等於錨點，每片現金成本在每個案例內重算），"
            "表內為情境加權後的 DCF、Comparable（相對倍數調整後）與加權目標價；另列 FY2027、FY2028 情境 A 的 EPS")
SE["A2"] = "相對基準＝案例加權目標價 ÷ 基準加權目標價 − 1。目標價參數（tp_r、tp_g、tp_w_dcf）三案只改 DCF 參數，PL 收斂法的 fade_g 不隨 tp_g 變動"
hdr(SE, 3, ["案例", "設定", "鍵", "FY2027 EPS（A）", "FY2028 EPS（A）", "FY2028 EPS 相對基準（A）", "DCF 目標價", "Comparable（調整後）目標價", "加權目標價", "相對基準"])
cases = C["sens_cases"]
SENS_ROWS = {}
def put_case(i, name, setting, keys, blocks, pre, r="tp_r", g="tp_g", w="tp_w_dcf"):
    """blocks：{情境: 推導區塊列號 dict}；pre：區塊所在工作表前綴（PL! 或 空字串＝本頁）"""
    SE.cell(i, 1, name); SE.cell(i, 2, setting); SE.cell(i, 3, keys)
    SE[f"D{i}"] = f"={pre}E{blocks['A']['eps']}"; SE[f"E{i}"] = f"={pre}F{blocks['A']['eps']}"
    SE[f"F{i}"] = f"=E{i}/E$4-1"
    reinv = f"MIN({g}/tp_roe,0.8)"
    dc = "+".join(f"tp_w{s}*{dcf_of([f'{pre}{FC[y]}{blocks[s]["eps"]}' for y in YRS6], r, g, reinv)}" for s in SC)
    SE[f"G{i}"] = f"=({dc})*(1+{r})"
    fe = "+".join(f"tp_w{s}*(2/12*{pre}E{blocks[s]['eps']}+10/12*{pre}F{blocks[s]['eps']})" for s in SC)
    SE[f"H{i}"] = f"=Val!$B$33*({fe})*rm_factor"
    SE[f"I{i}"] = f"={w}*G{i}+(1-{w})*H{i}"
    SE[f"J{i}"] = f"=I{i}/I$4-1"
    for c, fm in (("D", "#,##0.00"), ("E", "#,##0.00"), ("F", "0.0%"), ("G", "#,##0.00"), ("H", "#,##0.00"), ("I", "#,##0.00"), ("J", "0.0%")):
        SE[f"{c}{i}"].number_format = fm
put_case(4, "基準（三情境加權）", "PL 頁三個推導區塊", "—", PB, "PL!")
n_tp = 6
STOP = 4 + 2 * len(cases) + n_tp + 4
base_row = STOP + 2; SB = {}
SE.cell(STOP, 1, "敏感度「逐案推導區塊」：以下每個案例三個情境各一個完整推導區塊（不是案例列）").font = BOLD
i = 5
for key, name, unit, fmtv in cases:
    for side in ("lo", "hi"):
        blocks = {}
        for s in SC:
            ov = {key: f"{key}_{side}", "scn": s, "fm_self": True}
            blocks[s] = engine(SE, base_row, f"案例區塊：{name}：{'低' if side == 'lo' else '高'}（情境 {s}）", **ov)
            base_row = blocks[s]["end"] + 2
        SB[(key, side)] = blocks
        setting = f'=TEXT({key}_{side},"{fmtv}")&"（基準 "&TEXT({key},"{fmtv}")&"）"'
        put_case(i, f"{name}：{'低' if side == 'lo' else '高'}", setting, key, blocks, "")
        SENS_ROWS[(key, side)] = i; i += 1
TPC = [("折現率 +1pt", "r 11%", "tp_r", dict(r="(tp_r+0.01)")), ("折現率 −1pt", "r 9%", "tp_r", dict(r="(tp_r-0.01)")),
       ("長期成長 +1pt", "g 5%", "tp_g", dict(g="(tp_g+0.01)")), ("長期成長 −1pt", "g 3%", "tp_g", dict(g="(tp_g-0.01)")),
       ("DCF 權重 0.3", "Comparable 0.7", "tp_w_dcf", dict(w="0.3")), ("DCF 權重 0.7", "Comparable 0.3", "tp_w_dcf", dict(w="0.7"))]
assert len(TPC) == n_tp
TP_ROWS = {}
for nm, st, k, kw in TPC:
    put_case(i, nm, st, k, PB, "PL!", **kw); TP_ROWS[nm] = i; i += 1
SENS_LAST = i - 1
SE.cell(i + 1, 1, "讀法：" + C["sens_read"])
SENS_NOTE = i + 1
assert SENS_NOTE < STOP, (SENS_NOTE, STOP)
SE.column_dimensions["A"].width = 34; SE.column_dimensions["B"].width = 26; SE.column_dimensions["C"].width = 12
for c in "DEFGHIJ": SE.column_dimensions[c].width = 15

# ---------------- Recon ----------------
RC = wb.create_sheet("Recon")
RC["A1"] = f"Recon — 本模型 vs 公司說法、共識、券商。只作事後對照，不為對上數字而調參數。券商與資料商數字為 Interested-party，不用於校準"
hdr(RC, 3, ["項目", "對照值", "模型 A", "差距 A", "模型 B", "差距 B", "模型 C", "差距 C", "對照來源", "說明"])
rr_ = 4
def rec(label, ref, getter, src, note="", fmt="#,##0.00", pt=False, gap=True):
    global rr_
    RC.cell(rr_, 1, label)
    if ref is not None: RC.cell(rr_, 2, ref).number_format = fmt
    for j, s in enumerate(SC):
        cv, cd = 3 + 2 * j, 4 + 2 * j
        RC.cell(rr_, cv, getter(s)).number_format = fmt
        if ref is not None and gap:
            L = chr(64 + cv)
            RC.cell(rr_, cd, f"={L}{rr_}-B{rr_}" if pt else f"={L}{rr_}/B{rr_}-1").number_format = "0.0%"
    RC.cell(rr_, 9, src); RC.cell(rr_, 10, note); rr_ += 1
pl = lambda key, y: (lambda s: f"=PL!{FC[y]}{PB[s][key]}")
for item in C["recon"]:
    if item[0] == "sec":
        RC.cell(rr_, 1, item[1]).font = BOLD; rr_ += 1; continue
    lab, ref, (key, y), src, note, fmt, pt = item
    if key == "feps": getter = lambda s: f"=Val!B{T['feps']}"
    elif key.startswith("="): getter = (lambda f_: (lambda s: f_))(key)
    else: getter = pl(key, y)
    rec(lab, ref, getter, src, note, fmt=fmt, pt=pt)
RC.cell(rr_, 1, "預留").font = BOLD; rr_ += 1
RC.cell(rr_, 1, "成熟製程稼動率與價格的行業層（Siliconomics）尚未建立：上線後以本模型稼動率、售價情境對照行業路徑"); RC.cell(rr_, 10, "預留列：本版無數值"); rr_ += 1
RC.column_dimensions["A"].width = 80
for c in "BCDEFGH": RC.column_dimensions[c].width = 11
RC.column_dimensions["I"].width = 36; RC.column_dimensions["J"].width = 60
RECON_LAST = rr_ - 1

# ---------------- Sources ----------------
SO = wb.create_sheet("Sources"); SO["A1"] = "來源（公司財報與法說內容部分為媒體轉述；券商、資料商、媒體皆為 Interested-party）"; SO["A1"].font = BOLD
for i, (a, b) in enumerate(C["sources"], 2): SO.cell(i, 1, a); SO.cell(i, 2, b)
SO.column_dimensions["A"].width = 100; SO.column_dimensions["B"].width = 130

# ---------------- README ----------------
RD = wb.create_sheet("README", 0)
for i, (t, b) in enumerate(C["readme"], 1): RD.cell(i, 1, t).font = Font(bold=b)
RD.column_dimensions["A"].width = 170

_up = T["upw"]
GRADE_ROW = V.max_row + 2; REC_ROW = GRADE_ROW + 1
V.cell(GRADE_ROW, 1, "上檔空間級距")
V.cell(GRADE_ROW, 2, f'=IF(B{_up}>0.5,"大於 50%",IF(B{_up}>0.25,"25%–50%",IF(B{_up}>=0,"0%–25%",IF(B{_up}>=-0.25,"0%～−25%",IF(B{_up}>=-0.5,"−25%～−50%","小於 −50%")))))')
V.cell(GRADE_ROW, 5, f"以「上檔空間：加權目標價」（B{_up}）分六級（Andy 2026-10-25 14:00）")
V.cell(REC_ROW, 1, "投資建議")
V.cell(REC_ROW, 2, f'=IF(B{_up}>0.5,"強力買進",IF(B{_up}>0.25,"買進",IF(B{_up}>=-0.25,"中立",IF(B{_up}>=-0.5,"放空","強力放空"))))')
V.cell(REC_ROW, 5, "依上檔空間固定規則（Andy 2026-10-25 14:17）：大於 +50% 強力買進、+25%～+50% 買進、−25%～+25% 中立、−50%～−25% 放空、小於 −50% 強力放空")
_rr = RD.max_row + 1
RD.cell(_rr, 1, "")
RD.cell(_rr + 1, 1, f'="投資建議："&Val!B{REC_ROW}&"（依上檔空間固定規則）"').font = Font(bold=True)
OBS_ROW0 = _rr + 2
for _j, _t in enumerate(C["obs"], 1): RD.cell(_rr + 1 + _j, 1, f"重點觀察 {_j}：" + _t)

order = ["README", "In", "Q", "Units", "PL", "Val", "Recon", "Sens", "Sources"]
wb._sheets = [wb[n] for n in order]
import relative_multiple as RMX  # noqa: E402
RMX.apply(wb, C["folder"])
RD.cell(RD.max_row, 1, "相對倍數調整（tools/relative_multiple.py；Val 第 0.5 節）已含在加權目標價：Comparable（相對倍數調整後）＝同業前瞻倍數中位數 × 本公司 10 年相對同業的倍數比（Bloomberg 前瞻倍數）。「Comparable 目標價（市場主流同業法）」維持不調整。重點觀察的數字以本版重算值撰寫（第 1 點第一句為公式）。")
sys.path.insert(0, os.path.join(ROOT, "tools", "in_layout")); import apply_in_layout  # noqa: E402
apply_in_layout.apply(wb, "v0.1", "2026-10-31")
wb.save(sys.argv[1])
if os.environ.get("ROWMAP"):   # 除錯用：輸出列號對照（供撰寫 html.json）；CI 不設此變數
    json.dump({"PB": PB, "SUMR": {f"{k[0]}|{k[1]}": v for k, v in SUMR.items()}, "UR": UR, "CHK": CHK, "VAL": VAL, "INROW": INROW,
               "SENS_LAST": SENS_LAST, "SENS_NOTE": SENS_NOTE, "STOP": STOP, "RECON_LAST": RECON_LAST, "RULE_ROWS": rule_rows, "RB0": RB0, "WMT_NOTE": WMT_NOTE,
               "UN_NOTE": UN_NOTE, "T": T, "T10": T10, "PEER_LAST": PEER_LAST, "EPSR": EPSR, "O": O, "CY0": CY0, "R30": R30, "W0": W0,
               "GRADE_ROW": GRADE_ROW, "REC_ROW": REC_ROW, "OBS_ROW0": OBS_ROW0, "RD_LAST": RD.max_row,
               "SENS_ROWS": {f"{k[0]}|{k[1]}": v for k, v in SENS_ROWS.items()}, "TP_ROWS": TP_ROWS, "SEG_N": len(SEG["rows"])},
              open(os.environ["ROWMAP"], "w"), ensure_ascii=False, indent=1)
