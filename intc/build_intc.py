# 英特爾（INTC）個股模型 v0.1.0（AI 半導體 Project：代工層，IDM＋Intel Foundry；樣板 tsmc/、umc/（代工），產品線參 amd/）
# 用法：python3 intc/build_intc.py raw.xlsx ；再以 LibreOffice headless 重算
# 會計年度＝日曆年（Intel 會計年度為 12 月最後一個週六，近似 FY(t)＝CY(t)）。金額：十億美元（US$B）；EPS：美元（稀釋）；股數：十億股。
# 營收＝CCPG（客戶端運算與實體 AI，原 CCG）＋DCAI（資料中心與 AI；拆 AI 相關與其他）＋All Other＋外部代工（Intel Foundry 對外營收；18A／14A、先進封裝）。
# Intel Foundry 營收含對 Intel Products 的內部銷售（合併時沖銷），故合併營收＝CCPG＋DCAI＋All Other＋（Foundry＋分部間沖銷）。
# 估值主軸：代工虧損收斂的路徑（Foundry 營業利益率）與資本支出強度；少數股權（Arizona／Ireland 合資、Mobileye 等）在 EV 到股權的橋接中以帳面扣除。
R3 = "https://www.intc.com/news-events/press-releases/detail/1786/intel-reports-third-quarter-2026-financial-results"
R2 = "https://www.intc.com/news-events/press-releases/detail/1776/intel-reports-second-quarter-2026-financial-results"
R1 = "https://www.intc.com/news-events/press-releases/detail/1767/intel-reports-first-quarter-2026-financial-results/"
R4 = "https://www.intc.com/news-events/press-releases/detail/1759/intel-reports-fourth-quarter-and-full-year-2025-financial-results"
YFW = "repo docs/workorders/20261117_新增五家_Yahoo基本面快照.md（Yahoo Finance 2026-11-17 擷取；第三方彙整）"
CMC = "https://companiesmarketcap.com/intel/pe-ratio/"
SAI = "https://stockanalysis.com/stocks/intc/statistics/"
TRS = "美國財政部每日殖利率曲線 10 Yr：2026-10-23 5.00%（https://home.treasury.gov/resource-center/data-chart-center/interest-rates/TextView?type=daily_treasury_yield_curve&field_tdr_date_value_month=202610）"
CSV13 = "repo markets/invest_map/data/closes_20261020_20261113.csv"
CSV13X = "repo markets/invest_map/data/closes_extra_20261113.csv"
C = dict(kind="intc", folder="intc", name="英特爾")
C["urlname"] = {R3: "公司新聞稿 3Q26 財報（2026-10-29）", R2: "公司新聞稿 2Q26 財報（2026-07-23）", R1: "公司新聞稿 1Q26 財報（2026-04-23）", R4: "公司新聞稿 4Q25 與 FY2025 財報（2026-01-22）",
                CMC: "companiesmarketcap", SAI: "stockanalysis.com 統計頁"}
C["q_inputs"] = [
    ("q4_rev", "4Q26 營收（公司指引中點）", "US$B", 15.5, "Interested-party", "公司 4Q26 指引：營收 150–160 億美元（新聞稿 3Q26）", "取區間中點"),
    ("q4_gm", "4Q26 GAAP 毛利率", "%", 0.371, "Interested-party", "公司 4Q26 指引：GAAP 毛利率 37.1%（新聞稿 3Q26）", "較 3Q26 43.0% 低約 6 點，公司未在新聞稿中拆因"),
    ("q4_opx", "4Q26 GAAP 營業費用", "US$B", 4.9, "Interested-party", "公司 4Q26 指引：GAAP 營業費用約 49 億美元（新聞稿 3Q26）", "FY2026 全年約 226 億美元（同新聞稿）"),
    ("q4_tax", "4Q26 GAAP 稅率", "%", 0.31, "Interested-party", "公司 4Q26 指引：GAAP 稅率 31%（non-GAAP 11%）", "—"),
    ("q4_ccpg", "4Q26 CCPG 營收", "US$B", 7.7, "Assumed", "3Q26 8.097；總營收指引中點 15.5 較 3Q26 低 1.0，分部拆分公司未揭露", "按比例分攤季減：CCPG −5%"),
    ("q4_dcai", "4Q26 DCAI 營收", "US$B", 6.85, "Assumed", "3Q26 7.318；同上", "DCAI −6%"),
    ("q4_oth", "4Q26 All Other 營收", "US$B", 0.6, "Assumed", "3Q26 0.692", "外部代工營收為餘額（營收指引中點扣上列三項）"),
    ("q4_ccpg_oi", "4Q26 CCPG 營業利益", "US$B", 1.7, "Assumed", "3Q26 2.285（28.2%）", "毛利率指引降 6 點，分部營業利益率下調到約 22%；Corporate 為餘額"),
    ("q4_dcai_oi", "4Q26 DCAI 營業利益", "US$B", 2.1, "Assumed", "3Q26 3.039（41.5%）", "約 31%"),
    ("q4_fnd_oi", "4Q26 Intel Foundry 營業利益", "US$B", -2.1, "Assumed", "3Q26 −1.853、2Q26 −2.089", "—"),
    ("q4_oth_oi", "4Q26 All Other 營業利益", "US$B", 0.1, "Assumed", "3Q26 0.179", "—"),
    ("q4_iota", "4Q26 內部銷售占 Intel Products 營收比", "%", 0.41, "Assumed", "1Q26 41.1%、2Q26 36.2%、3Q26 40.8%（分部間沖銷 ÷ Intel Products 營收）", "—"),
    ("q4_nci", "4Q26 非控制權益損益", "US$B", 0.2, "Assumed", "3Q26 0.200、2Q26 0.185", "—"),
    ("q4_int", "4Q26 利息及其他淨額（不含 Escrowed Shares 評價）", "US$B", 0.15, "Assumed", "3Q26 約 +0.125、2Q26 約 −0.047、1Q26 約 +0.352（扣除 Escrowed Shares 評價後；Q 頁 Derived）", "Escrowed Shares 評價為非現金、不可預測，4Q26 取 0"),
    ("q4_da", "4Q26 折舊", "US$B", "=Q!E36", "Derived", "3Q26 折舊（Q 頁；9M26 折舊 89.33 億減 1Q26 29.02 億後兩季平分）", "—"),
    ("q4_cx", "4Q26 總資本支出（投資活動＋融資活動的 PP&E 增加）", "US$B", 4.5, "Assumed", "1Q26 4.963、2Q26 2.652、3Q26 4.219（公司定義的 gross capex）", "公司全年資本支出指引：未找到；取 3Q26 略上調"),
    ("q_sh", "稀釋加權股數（3Q26）", "十億股", 5.307, "Verified", "公司新聞稿 3Q26：稀釋加權股數 5,307 百萬股；期末流通 5,310 百萬股", "—")]
C["cons"] = [
    ("cons_eps26", "共識 EPS FY2026（只作對照，不進模型）", "US$", 1.487, YFW + "：2026-12 期共識 EPS 1.487；營收 61.74B", "Yahoo 共識為 non-GAAP 口徑，與本模型 GAAP 不同"),
    ("cons_eps27", "共識 EPS FY2027（只作對照，不進模型）", "US$", 1.791, YFW + "：2027-12 期共識 EPS 1.791；營收 69.13B", "同上"),
    ("cons_rev27", "共識營收 FY2027（只作對照，不進模型）", "US$B", 69.13, YFW, "—"),
    ("cons_n", "共識目標價家數", "家", 43, YFW + "：目標價均值 90.65（43 家）", "stockanalysis 頁面同日顯示 49 家、90.65")]
C["px"] = 59.66; C["px_date"] = "2026-11-13"
C["px_src"] = "收盤價取自 " + CSV13 + "（INTC 2026-11-13 收盤 59.66 美元）"
C["px_src2"] = "價格只用 repo 兩個 CSV，不另查（工作單第 3 節）"
C["eps_ttm"] = -3.51
C["eps_ttm_src"] = "近四季 GAAP 稀釋 EPS：4Q25 −0.12、1Q26 −0.73、2Q26 −2.16、3Q26 0.60＝−2.41（公司新聞稿）；含 Escrowed Shares 非現金評價損失（2Q26 每股 −2.45）"
C["eps_ttm"] = -2.41
C["fpe_lab"] = "前瞻 P/E（Yahoo 2026-11-17；股價約 55.5）"; C["fpe"] = 30.98; C["fpe_src"] = YFW + "：Intel forwardPE 30.98"
C["cfeps_lab"] = "共識前瞻 12 月 EPS（2/12 × FY2026 共識＋10/12 × FY2027 共識）"; C["cfeps"] = "=2/12*cons_eps26+10/12*cons_eps27"; C["cfeps_src"] = YFW + "；non-GAAP 共識（偏高於 GAAP）"
C["cons_tp"] = 90.65; C["cons_tp_src"] = YFW + "：目標價均值 90.65（43 家）；stockanalysis 2026-11-13 頁面 49 家、均值 90.65（Hold）"
C["cons_tp_src2"] = "賣方有目標價誘因；只作對照，不進目標價"
C["nim5"] = "=AVERAGE(1689/54228,-18756/53101,-267/52853)"
C["nim5_src"] = YFW + "：FY2023–FY2025 歸屬普通股淨利 ÷ 營收（FY2023 1,689／54,228、FY2024 −18,756／53,101、FY2025 −267／52,853；Yahoo 數）"
C["pe_cmc"] = [13.6, 19.0, 8.84, 11.0, 9.04, 9.75, 12.9, None, None, None]   # 2023 年 127 倍（>200 規則外但淨利近零，扭曲）、2024 −4.58、2025 −461 依規則剔除
C["pe_sa"] = [None] * 10
C["pe_note"] = "companiesmarketcap 年底 P/E：2023 年 127 倍（淨利近零）、2024–2025 年為負值，2023 年 127 倍（淨利近零）與負值年依規則剔除後只剩 2016–2022 的 7 年；Intel 近三年獲利被重組與認股權／Escrowed Shares 評價扭曲，十年 P/E 規則值參考性低，只作參考，不進目標價。"
C["scn_t"] = {"A": "情境 A（樂觀）：18A／14A 外部客戶與先進封裝放量（外部代工營收 FY2027–29 25／55／90 億美元）、Foundry 營業利益率 −24%／−14%／−5%；DCAI AI 相關年增 +40% 再乘顆數係數；CCPG 穩定成長、營業利益率 27%；資本支出 FY2027–29 180／190／190 億美元",
              "B": "情境 B（基準）：外部代工營收 20／35／55 億美元、Foundry 營業利益率 −27%／−20%／−13%；DCAI AI 相關 +25% 再乘顆數係數（0.8／0.71）；CCPG +2%／+1%／+1%、營業利益率 25%；資本支出 160／160／150 億美元",
              "C": "情境 C（保守）：外部代工營收 15／20／25 億美元、Foundry 營業利益率 −32%／−28%／−24%（良率與客戶延遲）；DCAI AI 相關 +8% 再乘顆數係數（0.6）；CCPG −3%／−2%／−1%、營業利益率 22%／21%／20%；資本支出 140／130／120 億美元"}
C["tp_method"] = ("DCF（標準版）：無槓桿自由現金流 FCFF＝NOPAT＋折舊 − 資本支出 − 營運資金增加，以 WACC 折現到 2026 年底（期中慣例），終值＝NOPAT(FY2033) ×（1 − g ÷ RONIC）÷（WACC − g），g＝MIN(3%, 美國 10 年期公債殖利率)、RONIC＝WACC；"
                  "ke＝rf＋β ×（ERP 5%＋CRP 0），β 為自算 US 半導體產業無槓桿 β（Blume 調整 1.884）依市值 D/E 重新加槓桿；EV 加現金與短期投資、權益投資，扣債務與非控制權益（帳面），÷ 股數，以 ke 推到 12 個月後（DCF 頁、tools/fcff_dcf.py）。"
                  "資本支出與折舊逐年取 PL 頁各情境（總資本支出＝投資活動＋融資活動的 PP&E 增加，不計合資夥伴挹注與政府補助），FY2030 起線性收斂。"
                  "Comparable：台積電 ADR、AMD、聯電 ADR、格芯、高塔前瞻 P/E 中位數（2026-11-13 收盤 ÷ Yahoo 共識 EPS 時間加權）× 相對倍數調整係數（Intel 自身 10 年相對同業的倍數比）× 12 個月後前瞻 EPS。"
                  "主要偏誤：Foundry 營業利益率路徑、外部代工營收與 DCAI AI 相關為 Assumed；FY2030 起只用收斂法（淨利率往 FY2029 × 0.9 或 0.8 收斂，對正在轉虧為盈的公司偏保守）；高 β 使 WACC 偏高。")
C["pe_peers_note"] = "Comparable 同業：台積電（ADR）、AMD、聯電（ADR）、格芯、高塔；相對倍數調整係數的 Bloomberg 前瞻倍數只有台積電、AMD、聯電（2303.TW）滿 7 年（格芯、高塔資料不足未納入）。"
C["ccg"] = {27: {"A": 0.04, "B": 0.02, "C": -0.03}, 28: {"A": 0.04, "B": 0.01, "C": -0.02}, 29: {"A": 0.03, "B": 0.01, "C": -0.01}}
C["ccg_src"] = "公司：3Q26 CCPG 80.97 億（年減 5%）、4Q26 營收指引季減（新聞稿 3Q26）；FY2026 約 324 億，FY2025 322 億（Q 頁）"
C["ccm"] = {27: {"A": 0.27, "B": 0.25, "C": 0.22}, 28: {"A": 0.27, "B": 0.25, "C": 0.21}, 29: {"A": 0.27, "B": 0.25, "C": 0.20}}
C["ccm_src"] = "FY2025 28.9%、FY2026 約 27%（Q 頁）；4Q26 毛利率指引大降（37.1%）顯示成本壓力"
C["ai_src"] = "公司：DCAI 3Q26 營收 73.18 億，年增 78%（新聞稿 3Q26）；AI 相關占比與增速公司未揭露（Assumed）"
C["aig27"] = {"A": 0.40, "B": 0.25, "C": 0.08}
C["dng"] = {27: {"A": 0.06, "B": 0.03, "C": 0.0}, 28: {"A": 0.05, "B": 0.03, "C": 0.0}, 29: {"A": 0.05, "B": 0.03, "C": 0.0}}
C["dng_src"] = "DCAI 其他（傳統伺服器 CPU）：公司未揭露；FY2025 DCAI 169 億、FY2026 約 255 億（Q 頁）"
C["dcm"] = {27: {"A": 0.38, "B": 0.36, "C": 0.30}, 28: {"A": 0.39, "B": 0.36, "C": 0.28}, 29: {"A": 0.39, "B": 0.36, "C": 0.27}}
C["dcm_src"] = "FY2025 20.2%、FY2026 約 36%（Q 頁）；Granite Rapids／Clearwater Forest 產品組合：未找到"
C["ext"] = {27: {"A": 2.5, "B": 2.0, "C": 1.5}, 28: {"A": 5.5, "B": 3.5, "C": 2.0}, 29: {"A": 9.0, "B": 5.5, "C": 2.5}}
C["ext_src"] = "外部代工營收：FY2025 1.4 億、FY2026 約 12 億美元（Q 頁：Intel Foundry 營收＋分部間沖銷）；18A／14A 外部客戶與先進封裝（EMIB／Foveros）合約金額公司未揭露（Assumed）"
C["fm"] = {27: {"A": -0.24, "B": -0.27, "C": -0.32}, 28: {"A": -0.14, "B": -0.20, "C": -0.28}, 29: {"A": -0.05, "B": -0.13, "C": -0.24}}
C["fm_src"] = "Foundry 營業利益率（占 Foundry 營收）：FY2025 −57.9%、FY2026 約 −35%（Q 頁；3Q26 −27.7%）；18A 良率與 14A 客戶決定收斂速度（Assumed；公司未給目標損益兩平時點）"
C["nonop_src"] = "FY2026 前三季權益投資損益＋利息及其他（扣除 Escrowed Shares 評價）約 +6.5 億（Q 頁 Derived）"
C["nci_src"] = "公司新聞稿 3Q26：2026-09-26 非控制權益 16,642 百萬美元（2025-12-27 為 12,079）"
C["sh_src"] = "3Q26 期末流通 5,310 百萬股（新聞稿），較 2025-12 的 4,994 百萬股多 3.2 億股（美國政府、SoftBank、NVIDIA 入股與 Escrowed Shares 釋出；明細未取得）"
C["sh_note"] = "美國政府認股權證（Foundry 持股低於 51% 時行使）與託管股份餘額的未來稀釋未建模：金額未取得"
C["cx"] = {27: {"A": 18.0, "B": 16.0, "C": 14.0}, 28: {"A": 19.0, "B": 16.0, "C": 13.0}, 29: {"A": 19.0, "B": 15.0, "C": 12.0}}
C["cx_src"] = "FY2025 總資本支出約 176.7 億、FY2026 約 163 億（Q 頁）；公司 FY2027 資本支出指引：未找到。資本支出強度是估值主軸"
C["units_a1"] = "Units — Intel 分部、外部代工與資本支出（三情境）。DCAI 中 AI 相關營收 FY2028–29 乘 Siliconomics 同字母顆數係數（um28／um29）；其餘為公司層級假設（待併入 Siliconomics）"
C["units_note"] = "讀法：Intel 的 EPS 由 Foundry 虧損收斂速度與資本支出決定；DCAI 的 AI 相關成長是營收端變數，但 Foundry 年虧 50–90 億美元的規模使其影響被稀釋。"
_YH = "收盤 2026-11-13（repo closes CSV）÷ Yahoo 共識 EPS（2026-11-17；2/12 × 本年度＋10/12 × 次年度）"
C["peers"] = [("台積電（TSMC，ADR）", "TSM", "=398.73/(2/12*17.5+10/12*23.39)", "2026-11-13", f"TSM {_YH}：398.73 ÷（17.50、23.39）"),
              ("超微（AMD）", "AMD", "=534.06/(2/12*8.084+10/12*17.64)", "2026-11-13", f"AMD {_YH}：534.06 ÷（8.084、17.64）；Intel CPU 直接競爭者"),
              ("聯電（UMC，ADR）", "UMC", "=15.4/(2/12*1.135+10/12*0.9)", "2026-11-13", f"UMC {_YH}：15.40 ÷（1.135、0.900）；次年度 EPS 低於本年度（Yahoo 數，疑含一次性）"),
              ("格芯（GlobalFoundries）", "GFS", "=41.47/(2/12*1.988+10/12*2.6)", "2026-11-13", f"GFS {_YH}：41.47 ÷（1.988、2.600）"),
              ("高塔半導體（Tower）", "TSEM", "=146.89/(2/12*4.067+10/12*7.004)", "2026-11-13", f"TSEM {_YH}：146.89 ÷（4.067、7.004）")]
C["peer_note"] = "晶圓代工與 CPU 同業（台積電、AMD、聯電、格芯、高塔）；EPS 為第三方彙整（Interested-party），GAAP／non-GAAP 口徑不一；不含 A 股"
C["sens_cases"] = [("ai_adj", "DCAI 中 AI 相關營收係數", "x", "0.00"), ("ext_adj", "外部代工營收係數", "x", "0.00"), ("fm_adj", "Foundry 營業利益率調整", "pt", "0.0%"),
                   ("dcm_adj", "DCAI 營業利益率調整", "pt", "0.0%"), ("ccm_adj", "CCPG 營業利益率調整", "pt", "0.0%"), ("ccg_adj", "CCPG 營收年增調整", "pt", "0.0%"),
                   ("cx_adj", "總資本支出係數", "x", "0.00"), ("opx_adj", "營業費用係數", "x", "0.00"), ("nonop_f", "營業外淨額", "US$B", "0.0"),
                   ("tax", "稅率", "%", "0%"), ("nci_r", "非控制權益報酬率", "%", "0.0%")]
C["sens_read"] = "資本支出、Foundry 營業利益率與 CCPG／DCAI 營業利益率是加權目標價最敏感的輸入（Foundry 年虧規模大，利益率每 5 個百分點約 4–7 億美元／年；資本支出每 10% 約 16 億美元／年的自由現金流）；DCAI 中 AI 相關營收 ±30% 與外部代工營收 ±50% 的影響較小。"
C["recon"] = [("sec", "A. FY2026 錨點 vs 共識（non-GAAP；GAAP 與 non-GAAP 口徑不同，差距不代表錯誤）"),
    ("FY2026 營收 vs 共識（Yahoo 61.74B）", "=61.74", ("rev", 2026), "Yahoo 2026-11-17", "共識營收與 4Q26 指引中點接近", "#,##0.00", False),
    ("FY2026 EPS（GAAP）vs 共識（non-GAAP 1.487）", "=cons_eps26", ("eps", 2026), "Yahoo 2026-11-17", "本模型含 Escrowed Shares 非現金評價損失（每股約 −2.3），口徑不同；不含評價的參考值見 Q 頁", "#,##0.00", False),
    ("sec", "B. FY2027–29 vs 共識與行業"),
    ("FY2027 營收 vs 共識（Yahoo 69.13B）", "=cons_rev27", ("rev", 2027), "Yahoo 2026-11-17", "A／B／C 對共識的位置", "#,##0.00", False),
    ("FY2027 EPS（GAAP）vs 共識（non-GAAP 1.791）", "=cons_eps27", ("eps", 2027), "Yahoo 2026-11-17", "共識為 non-GAAP，排除股票薪酬與攤銷；GAAP 與 non-GAAP 差距約每股 0.9–1.0 美元（FY2026 營業費用 GAAP 226 億 vs non-GAAP 163 億）", "#,##0.00", False),
    ("FY2027 總資本支出年增 vs Siliconomics 代工與邏輯設備年增（百分點）", "=eqfd27", ("cxg", 2027), "Siliconomics 現行基準 eq_g_fd_27", "層級為全產業代工與邏輯設備成長，Intel 資本支出為公司自有假設", "0.0%", True),
    ("共識前瞻 12 月 EPS vs 模型（情境加權 2／12 × FY2027＋10／12 × FY2028）", "=cons_feps", ("feps", 0), "Yahoo 共識推得（non-GAAP）", "三欄同值（情境加權）", "#,##0.00", False)]
C["sources"] = [("公司：新聞稿 3Q26 財報（2026-10-29；3Q26 損益、分部、資產負債表、現金流量、4Q26 指引）", R3),
    ("公司：新聞稿 2Q26 財報（2026-07-23；2Q26 損益、分部、Escrowed Shares 評價 125.29 億、資產負債表、1H26 現金流）", R2),
    ("公司：新聞稿 1Q26 財報（2026-04-23；1Q26 損益、分部、資產負債表、資本支出、折舊）", R1),
    ("公司：新聞稿 4Q25 與 FY2025 財報（2026-01-22；FY2025 損益、分部、資產負債表、資本支出與折舊）", R4),
    ("第三方：Yahoo Finance 基本面快照（FY2022–25 年度與近五季、共識 EPS、目標價、同業前瞻 P/E；僅交叉檢查與同業 EPS；Interested-party）", YFW),
    ("資料商：companiesmarketcap 年底 P/E（2016–2025）", CMC), ("資料商：stockanalysis 統計頁（共識目標價、前瞻 P/E 對照）", SAI),
    ("價格：repo markets/invest_map/data/closes_20261020_20261113.csv 與 closes_extra_20261113.csv（2026-11-13 收盤；不另查）", CSV13 + "；" + CSV13X),
    ("無風險利率：美國財政部每日殖利率曲線 10 Yr（2026-10-23 5.00%）", TRS),
    ("行業層：Siliconomics 快照（顆數係數 um28／um29、Wattonomics 錨定 wta、代工與邏輯設備成長 eq_g_fd；本 repo，唯讀）", "siliconomics/data/CURRENT_SNAPSHOT；siliconomics/reader.py；markets/scen"),
    ("自算產業 β：US 半導體 1.884", "docs/plan/自算產業β_20261105.md"),
    ("待補（人工）：10-Q 租賃負債與權益投資明細、Escrowed Shares 餘額與美國政府認股權證條款、FY2023 與 FY2024 資產負債表（營運資金比）、公司資本支出指引、Foundry 損益兩平時點、稀釋 EPS 分子調整、2016–2020 淨利率", "未找到"),
    ("SemiAnalysis", "未使用")]
SABS = "公司新聞稿 3Q26（2026-10-29）資產負債表（2026-09-26）"
C["dcf_P"] = dict(
    date="2026-12-31（工作單評價日）", ccy="US$", amt="US$B", shu="十億股",
    rf=(0.0500, TRS + "；2026-11-05 讀取（repo AMD／NVIDIA 模型同源）", "FRED DGS10（同一序列）；利率日 2026-10-23，股價日 2026-11-13（利率未更新：未取得較新值）", "無風險利率（美國 10 年期公債殖利率，2026-10-23）"),
    beta_fixed=("=(beta_u-beta_du)*(1+(1-tax_w)*beta_de)", "=beta_u*(1+(1-tax_w)*beta_de)", "=(beta_u+beta_du)*(1+(1-tax_w)*beta_de)", "Derived",
                "自算產業 β：US 半導體 同業 13 家中位數（含 Intel；Yahoo Finance 月報酬 2021-11～2026-10，對 S&P 500 迴歸，Hamada 去槓桿、現金修正）；Blume 調整 0.67β＋0.33（Andy 2026-11-05 22:55）；以本公司市值 D/E 與稅率重新加槓桿",
                "Sens 低／高（Analogy）：β_u ∓ beta_du（0.3），各自同式重新加槓桿（Andy 2026-11-05 23:16 定案）；個股迴歸 β（Yahoo 2.43）未採用"),
    kd=("=0.0425", "Assumed", "債務平均票面利率：未找到（10-Q 債務附註）；取 4.25%（Assumed）", "債務占資本約 14%，對 WACC 影響小"),
    tax_w=("=tax", "Derived", "＝In 頁 tax（FY2027 起有效稅率 13%）", "NOPAT 用同一稅率"),
    debt=(50.134, "Verified", SABS + "：短期債務 3,817＋長期債務 46,317（US$M）", "不含租賃與合資夥伴的優先權益"),
    lease=(0.0, "Assumed", "租賃負債：未找到（10-Q 附註）；以 0 計（偏高估股權價值，金額相對 EV 小）", "待補"),
    cash=(54.005, "Verified", SABS + "：現金及約當現金 21,846＋短期投資 32,159（US$M）", "Yahoo：現金＋短投 54,005 同值"),
    inv=(8.481, "Verified", "公司新聞稿 1Q26 資產負債表：權益投資 8,481（US$M；2026-03-28；2026-09-26 數未取得）", "含公開與非公開權益投資；Sens 測不計"),
    mi=(16.642, "Verified", SABS + "：非控制權益 16,642（US$M；帳面）", "帳面值扣除；市價可能更高（Mobileye 上市部分、Arizona／Ireland 合資的優先報酬）"),
    stub=dict(bs_date="2026-09-26", val_date="2026-12-31", bs_src=SABS,
              div=(0, "Verified", "Intel 2025 年起未配發現金股利", "評價日前無已宣告股利或買回")),
    dps=(0.0, "Assumed", "Intel 目前不配息", "—"),
    sh=(5.31, "Verified", "公司新聞稿 3Q26：2026-09-26 期末流通 5,310 百萬股", "美國政府認股權證與託管股份餘額的未來稀釋未建模"),
    fc_da=(None, "=AVERAGE(10.757/52.853,11.9/61.7)", None, "Derived",
           "折舊 ÷ 營收：FY2025 10.757／52.853（新聞稿 4Q25）、FY2026 約 11.9／61.7（Q 頁）；平均約 19.8%", "DCF 頁 FY2027–29 取 PL 頁各情境折舊，本比例用於 stub 後的收斂與 Sens"),
    fc_cx=(None, "=AVERAGE(17.672/52.853,16.3/61.7)", None, "Derived",
           "總資本支出 ÷ 營收：FY2025 17.672／52.853、FY2026 約 16.3／61.7（Q 頁）；平均約 29.9%", "FY2027–29 取 PL 頁各情境資本支出；Sens 用 fc_cx_d ±1pt"),
    fc_cx_ss=("=MAX(fc_da_ss*(1+g_tv),(fc_cx+fc_da_ss*(1+g_tv))/2)", "Assumed", "FY2030 起線性收斂，FY2032 達到：近期比例與維持水準（折舊 ×（1＋g））的中點，下限為維持水準", "Andy 可改"),
    fc_wc=(0.05, "=(3.839+11.618-9.882)/52.853", 0.20, "Derived",
           "FY2025 年底營運資金 ÷ 營收：應收 3.839＋存貨 11.618 − 應付 9.882（公司新聞稿 4Q25 資產負債表）÷ 52.853＝10.5%；FY2023–24 資產負債表未取得", "Sens 低／高 5%／20%（Assumed）"))
C["dcf_ncbs"] = ("21.846+32.159+8.481-50.134-0+0-16.642", "2026-09-30 資產負債表（US$B）：現金 21.846＋短期投資 32.159＋權益投資 8.481（1Q26）− 債務 50.134 − 租賃 0（未找到）− 非控制權益 16.642")
C["readme"] = [
 ("英特爾（INTC）個股模型 v0.1.0（AI 半導體 Project：代工層，IDM＋Intel Foundry；樣板 tsmc/、umc/，產品線參 amd/）　2026-11-17", True),
 ("版本：v0.1.0（初版）。營收＝CCPG＋DCAI（拆 AI 相關與其他）＋All Other＋外部代工營收。滾入 3Q26 實績（2026-10-29 財報）與 4Q26 指引。本版未做先驗。", False),
 ("模型命題：Intel 的獲利由兩個變數決定——Intel Foundry 的虧損收斂速度（FY2026 年虧約 85 億美元、營業利益率約 −35%）與資本支出強度（總資本支出約占營收 26–30%）；Intel Products（CCPG、DCAI）是現金來源，DCAI 受 AI 伺服器 CPU 需求帶動。少數股權（Arizona Fab 與 Brookfield、Ireland 合資與 Apollo、Mobileye 等）持續上升，股數因美國政府、SoftBank、NVIDIA 入股增加。", False),
 ("核心問題：Foundry 能否在 FY2029 前把營業利益率由 −35% 收斂到接近損益兩平（情境 A −5%、B −13%、C −24%）；外部客戶（18A／14A、先進封裝）能把外部代工營收帶到多大（FY2029 A／B／C 90／55／25 億美元，Assumed）；現價 59.66 美元隱含的前瞻 P/E 約 34 倍（Yahoo 共識 EPS），模型要多少獲利才撐得住。", False),
 ("", False),
 ("會計年度近似日曆年（Intel 會計年度為 12 月最後一個週六），FY(t)≈CY(t)，不平移。FY2025A 實績、FY2026 錨點（1Q–3Q26 實績＋4Q26 指引）、FY2027–FY2029 驅動因子推導、FY2030–FY2032 收斂法。金額：十億美元；EPS：美元（稀釋）；口徑：GAAP、歸屬 Intel。", False),
 ("", False),
 ("驅動因子 → 推導量（行業變數讀 Siliconomics；其餘為公司層級，待併入 Siliconomics）", True),
 ("Intel Products：CCPG 營收年增（FY2027–29：A +4／+4／+3%、B +2／+1／+1%、C −3／−2／−1%）× 營業利益率（A 27%、B 25%、C 22／21／20%）；DCAI＝AI 相關（FY2026 占 35%，年增 A／B／C +40／+25／+8%，FY2028–29 起乘 Siliconomics 顆數係數 um28／um29）＋其他（年增 +6／+3／0%），營業利益率 A 38–39%、B 36%、C 30→27%。", False),
 ("Intel Foundry：營收＝內部銷售（Intel Products 營收 × 約 40%）＋外部代工營收（A／B／C FY2029 90／55／25 億美元）；營業利益率（占 Foundry 營收）FY2027–29 A −24／−14／−5%、B −27／−20／−13%、C −32／−28／−24%；Corporate Unallocated 每年 −40 億美元。", False),
 ("費用與現金：GAAP 營業費用 FY2027 208 億、之後年增 2%；營業外 +3 億；稅率 13%（稅前為負不計稅）；非控制權益損益＝期初餘額 × 5%（餘額 166 億，每年增 20 億夥伴挹注）；股數 FY2027 起年增 1%。總資本支出（投資活動＋融資活動的 PP&E 增加）FY2027–29 A 180／190／190 億、B 160／160／150 億、C 140／130／120 億；折舊年增 4%。", False),
 ("校準：FY2026 ＝ 1Q–3Q26 實績（公司新聞稿）＋4Q26 公司指引（營收中點 155 億、GAAP 毛利率 37.1%、營業費用 49 億、稅率 31%）；4Q26 分部拆分為假設（Corporate 為餘額）。不用券商數字（只在 Recon 對帳）。", False),
 ("估值：目標價＝0.5 × DCF（標準版 FCFF＋WACC，三情境權重 0.3／0.5／0.2，情境價值負值取 0）＋0.5 × Comparable（台積電、AMD、聯電、格芯、高塔前瞻 P/E 中位數 × 相對倍數調整係數 × 12 個月後前瞻 EPS）；另列 Comparable（市場主流同業法）目標價、市場共識目標價與差距拆解。EV 到股權：加現金與短期投資 540 億美元、權益投資 85 億，扣債務 501 億、非控制權益帳面 166 億（租賃負債未找到，以 0 計）。", False),
 ("預估期間：FY2027–FY2029 由驅動因子推導；FY2030 起為收斂法（營收年增收斂到 4%、淨利率往 FY2029 × 0.9（B）／ × 0.8（C）／持平（A）收斂；PL 頁灰底），對正在轉虧為盈的公司偏保守（待判斷）。", False),
 ("標記：Verified、Interested-party、Analogy、Assumed、Derived。SemiAnalysis 未使用。公司數字取自公司新聞稿（IR 頁，來源等級：公司原文）；Yahoo 為第三方彙整，只作交叉檢查與同業 EPS（Interested-party）。", False),
 ("", False),
 ("已知限制：(1) Escrowed Shares（美國商務部 CHIPS 協議託管股份）的 FY2026 評價損失 126 億美元為非現金，FY2027 起不預測；其餘額與美國政府認股權證的稀釋未建模；(2) 3Q26 公司稀釋 EPS 0.60 與淨利 ÷ 稀釋股數 0.63 不同（分子另有調整，未取得）；(3) 4Q26 分部拆分、Foundry 損益兩平時點、外部客戶金額公司未揭露；(4) 租賃負債、FY2023–24 資產負債表（營運資金比）、債務利率未取得；(5) 共識 EPS 為 non-GAAP，與 GAAP 差距每股約 0.9–1.0 美元；(6) 權益投資取 1Q26 帳面；(7) 利率為 2026-10-23（美國 10 年期）而股價為 2026-11-13。", False),
 ("", False),
 ("情境對照表（A／B／C 讀同字母的行業情境；公司特有變數 A 最有利、C 最不利）", True),
 ("行業：DCAI 中 AI 相關營收 FY2028–29 乘 Siliconomics um28／um29（A 1.0／1.0、B 0.8／0.71、C 0.6／0.6）；Siliconomics 代工與邏輯設備年增（eq_g_fd：2026 18.9%、2027 18.1%、2028 13.6%）只作 Intel 資本支出的對照（Recon）。Siliconomics 沒有伺服器 CPU、客戶端 PC、代工利益率、外部代工營收的變數，皆為公司層級假設（建議補進 Siliconomics，見 README 檔）。", False),
 ("公司特有：外部代工營收、Foundry 營業利益率路徑、CCPG／DCAI 營業利益率、資本支出（見上）。", False)]
C["obs"] = ["（建置後由公式取代）", "（建置後由公式取代）", "（建置後由公式取代）"]

# ======================= 共用建置程式（聯電、世界先進同一套結構；數字與文字在上方 C 設定） =======================
import sys, os, json, re
import openpyxl
from openpyxl.styles import Font, PatternFill
from openpyxl.workbook.defined_name import DefinedName
ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.dont_write_bytecode = True      # 不在上游資料夾寫入 __pycache__（上游唯讀）
sys.path.insert(0, os.path.join(ROOT, "tools"))
sys.path.insert(0, os.path.join(ROOT, "siliconomics"))
import reader as SIL   # Siliconomics 快照（行業變數單一來源）
sys.path.insert(0, os.path.join(ROOT, "markets", "scen")); import scen as SCN   # noqa: E402  v0.16：產業情境 A／C 參考值（市場層；B 讀 Siliconomics）
import cycle_position as CP
import fcff_dcf as FD   # v0.2：標準版 DCF（無槓桿自由現金流＋WACC）共用模組
import valuation_rule as VR

KIND = C["kind"]
BLUE = Font(color="0000FF"); BOLD = Font(bold=True); HDR = PatternFill("solid", fgColor="DDEBF7")
KEY = PatternFill("solid", fgColor="FFF2CC")
wb = openpyxl.Workbook()
def hdr(ws, row, labels):
    for j, h in enumerate(labels, 1):
        c = ws.cell(row, j, h); c.font = BOLD; c.fill = HDR
SC = ("A", "B", "C")
SCN_NAME = {"A": "樂觀", "B": "基準", "C": "保守"}
SILV = f"Siliconomics {SIL.VERSION}（siliconomics/data/{SIL.FILE}）"
TBD = "待併入 Siliconomics"

# ---------------- In ----------------
IN = wb.active; IN.title = "In"
IN["A1"] = ("In — 輸入（藍字；基準欄 E，低 D、高 F）。金額：十億美元（US$B）；股數：十億股；EPS：美元（稀釋）。"
            "公司財報數字為 Verified；公司指引為 Interested-party；券商、資料商為 Interested-party；其餘為 Assumed（區間＋Sens）或 Analogy")
IN["A2"] = ("標記：Verified／Interested-party／Analogy／Assumed／Derived；Analogy、Assumed 以情境或 Sens 區間呈現，不當點估計。"
            f"行業變數讀 {SILV}（AI 加速器顆數係數 um28／um29、Wattonomics 錨定 wta、代工與邏輯設備成長率 eq_g_fd 僅對照）；Siliconomics 無伺服器 CPU、客戶端 PC、代工利益率層，說明欄標「{TBD}」。SemiAnalysis 未使用")
hdr(IN, 3, ["鍵", "項目", "單位", "低", "基準", "高", "標記", "來源1", "來源2／說明"])
rows = []; sec = lambda t: rows.append(("sec", t)); inp = lambda *a: rows.append(a)
URLNAME = C["urlname"]

sec(f"A. 行業變數（{SILV}；A／B／C 讀同字母上游，不自訂行業路徑）")
for s in SC:
    for y in (28, 29):
        k = f"um{y}_{s}"
        inp(k, f"情境 {s}：CY20{y} AI 加速器顆數 ÷ 基準", "x", None, SCN.base(k), None, SCN.V[k]["tag"], (f"{SILV}｜{k}" if s == "B" else f"{SCN.SRC}｜{k}"),
            "DCAI 中 AI 相關營收 CY2028–29 依同字母顆數係數縮放（行業情境；A＝原始顆數）" if (s == "A" and y == 28) else "同上")
for y in (26, 27, 28):
    inp(f"eqfd{y}", f"對照：Siliconomics「代工與邏輯設備」CY20{y} 年增（SEMI；只用於 Recon 與 Coord，不驅動模型）", "%", None, SIL.base(f"eq_g_fd_{y}"), None, SIL.V[f"eq_g_fd_{y}"]["tag"], f"{SILV}｜eq_g_fd_{y}",
        "Intel 資本支出年增對照此行業成長；層級無 Intel 口徑的資本支出")

sec(f"B. Intel Products：CCPG 與 DCAI（{TBD}：Siliconomics 尚無客戶端 PC 與伺服器 CPU 層）")
for y in (27, 28, 29):
    for s in SC:
        inp(f"ccg{y}_{s}", f"FY20{y} CCPG 營收年增（情境 {s}）", "%", None, C["ccg"][y][s], None, "Assumed", C["ccg_src"] if (s == "A" and y == 27) else "同上",
            f"A 最有利、C 最不利（C＝記憶體成本上升與 PC 需求疲弱、AMD／Arm 搶份額）；{TBD}" if s == "A" else TBD)
for y in (27, 28, 29):
    for s in SC:
        inp(f"ccm{y}_{s}", f"FY20{y} CCPG 營業利益率（情境 {s}）", "%", None, C["ccm"][y][s], None, "Assumed", C["ccm_src"] if (s == "A" and y == 27) else "同上", TBD if s != "A" else f"FY2026 約 27%（Q 頁）；{TBD}")
inp("ccg_adj", "CCPG 營收年增調整（Sens 用；FY2027 起各年加減）", "pt", -0.03, 0.0, 0.03, "Assumed", "無", "±3 個百分點")
inp("ccm_adj", "CCPG 營業利益率調整（Sens 用；FY2027 起各年加減）", "pt", -0.03, 0.0, 0.03, "Assumed", "無", "±3 個百分點")
inp("ai_sh26", "FY2026 DCAI 中「AI 相關」占比（AI 加速器周邊的伺服器 CPU、先進封裝內部需求；本模型拆分用）", "%", None, 0.35, None, "Assumed", C["ai_src"], "公司未揭露 DCAI 內的 AI 相關占比；DCAI 3Q26 年增 78%（新聞稿），本值為拆分假設")
for s in SC:
    inp(f"aig27_{s}", f"FY2027 DCAI 中 AI 相關營收年增（情境 {s}）", "%", None, C["aig27"][s], None, "Assumed", C["ai_src"] if s == "A" else "同上", "A 最有利、C 最不利" if s == "A" else "—")
inp("aig_L", "FY2028–29 DCAI 中 AI 相關營收年增（情境 A 路徑；B、C 再乘同字母顆數係數）", "%", None, 0.40, None, "Assumed", C["ai_src"], "B、C 乘 um28／um29，隱含增速低於 A")
inp("ai_adj", "DCAI 中 AI 相關營收係數（Sens 用；FY2027 起）", "x", 0.7, 1.0, 1.3, "Assumed", "無", "±30%")
for y in (27, 28, 29):
    for s in SC:
        inp(f"dng{y}_{s}", f"FY20{y} DCAI 其他（非 AI 相關）營收年增（情境 {s}）", "%", None, C["dng"][y][s], None, "Assumed", C["dng_src"] if (s == "A" and y == 27) else "同上", TBD)
for y in (27, 28, 29):
    for s in SC:
        inp(f"dcm{y}_{s}", f"FY20{y} DCAI 營業利益率（情境 {s}）", "%", None, C["dcm"][y][s], None, "Assumed", C["dcm_src"] if (s == "A" and y == 27) else "同上", TBD if s != "A" else f"FY2026 約 36%；3Q26 41.5%、2Q26 39.5%、1Q26 30.5%（Q 頁）")
inp("dcm_adj", "DCAI 營業利益率調整（Sens 用；FY2027 起各年加減）", "pt", -0.03, 0.0, 0.03, "Assumed", "無", "±3 個百分點")
inp("oth_g", "All Other 營收年增（FY2027 起；Mobileye 等）", "%", None, 0.03, None, "Assumed", "FY2025 3.563、FY2026 約 2.62（Q 頁）；Altera 於 2025 年 9 月起不再合併（新聞稿 3Q25 說明）", "三情境相同")
inp("oth_m", "All Other 營業利益率", "%", None, 0.20, None, "Assumed", "FY2026 約 23%（Q 頁）", "三情境相同")

sec("C. Intel Foundry：營收（內部銷售＋外部）、營業利益率路徑（估值主軸）")
inp("iota", "內部銷售占 Intel Products 營收比（FY2027 起）", "%", None, "=a26_iota", None, "Derived", "FY2026 由 Q 頁（分部間沖銷 ÷ Intel Products 營收）推得；FY2025 36%、1Q–3Q26 41%／36%／41%", "三情境相同；影響 Foundry 營收規模與沖銷，不影響合併營收")
for y in (27, 28, 29):
    for s in SC:
        inp(f"ext{y}_{s}", f"FY20{y} 外部代工營收（情境 {s}；18A／14A 與先進封裝 EMIB／Foveros 對外）", "US$B", None, C["ext"][y][s], None, "Assumed", C["ext_src"] if (s == "A" and y == 27) else "同上", "FY2026 約 12 億美元（Q 頁）；客戶名單與合約金額公司未揭露" if (s == "A" and y == 27) else "—")
inp("ext_adj", "外部代工營收係數（Sens 用；FY2027 起）", "x", 0.5, 1.0, 1.5, "Assumed", "無", "±50%")
for y in (27, 28, 29):
    for s in SC:
        inp(f"fm{y}_{s}", f"FY20{y} Intel Foundry 營業利益率（占 Foundry 總營收；情境 {s}）", "%", None, C["fm"][y][s], None, "Assumed", C["fm_src"] if (s == "A" and y == 27) else "同上", "FY2026 約 −35%（Q 頁）；A 最有利、C 最不利" if (s == "A" and y == 27) else "—")
inp("fm_adj", "Foundry 營業利益率調整（Sens 用；FY2027 起各年加減）", "pt", -0.05, 0.0, 0.05, "Assumed", "無", "±5 個百分點")
inp("cu_f", "Corporate Unallocated（FY2027 起每年；重組、股票薪酬、併購攤銷等）", "US$B", -4.8, -4.0, -3.2, "Assumed", "FY2025 −5.518、2Q26 −1.416、3Q26 −1.058（Q 頁）；1Q26 −5.080 含重組與商譽減損", "三情境相同")

sec("D. 成本與損益：營業費用、營業外、稅、非控制權益、股數")
inp("opx27", "FY2027 GAAP 營業費用", "US$B", None, 20.8, None, "Assumed", "FY2026 約 226 億（公司指引，Q 頁 22.7）；公司 non-GAAP 約 163 億，差額為股票薪酬與併購攤銷", "重組費用下降；三情境相同")
inp("opx_g", "FY2028 起營業費用年增", "%", None, 0.02, None, "Assumed", "無", "—")
inp("opx_adj", "營業費用係數（Sens 用；FY2027 起）", "x", 0.95, 1.0, 1.05, "Assumed", "無", "±5%")
inp("nonop_f", "營業外淨額（FY2027 起每年；權益投資損益＋利息及其他，不含 Escrowed Shares 評價）", "US$B", -0.5, 0.3, 1.0, "Assumed", C["nonop_src"], "Escrowed Shares 評價為非現金、不可預測，FY2027 起取 0")
inp("tax", "有效稅率（FY2027 起；正稅前利益）", "%", 0.10, 0.13, 0.20, "Assumed", "公司 non-GAAP 稅率 11%、4Q26 GAAP 指引 31%（含一次性）；FY2025 GAAP 稅率 98%（稅前微幅正）", "稅前為負時不計稅負（不認遞延稅利益）")
inp("nci_bs", "非控制權益餘額（2026-09-26）", "US$B", None, 16.642, None, "Verified", C["nci_src"], "含 Arizona Fab（Brookfield）、Ireland 合資（Apollo）、Mobileye 少數股東等；12.079（2025-12）→ 16.642")
inp("nci_r", "非控制權益損益 ÷ 期初非控制權益餘額（FY2027 起）", "%", 0.03, 0.05, 0.07, "Assumed", "3Q26 0.200 ×4 ÷ 16.642＝4.8%；2Q26 0.185（Q 頁）", "合資夥伴挹注（財務活動）增加餘額，報酬率另計")
inp("nci_g", "非控制權益餘額年增（夥伴挹注，FY2027 起）", "US$B", None, 2.0, None, "Assumed", "FY2025 夥伴挹注 51.08 億（新聞稿 4Q25）；FY2026 至 9M 餘額增加 45.6 億", "不含 Mobileye 等因素")
inp("sh_g", "加權股數年變動（FY2027 起）", "%", None, 0.01, None, "Assumed", C["sh_src"], C["sh_note"])
inp("da_g", "折舊年增（FY2027 起）", "%", 0.02, 0.04, 0.06, "Assumed", "FY2025 107.57 億、FY2026 約 119 億（Q 頁）", "資本支出高於折舊使折舊上升")
for y in (27, 28, 29):
    for s in SC:
        inp(f"cx{y}_{s}", f"FY20{y} 總資本支出（投資活動＋融資活動的 PP&E 增加；情境 {s}）", "US$B", None, C["cx"][y][s], None, "Assumed", C["cx_src"] if (s == "A" and y == 27) else "同上", "A 最有利者資本支出最高（外部客戶需要產能）；FY2026 約 163 億（Q 頁）" if (s == "A" and y == 27) else "—")
inp("cx_adj", "總資本支出係數（Sens 用；FY2027 起）", "x", 0.9, 1.0, 1.1, "Assumed", "無", "±10%")

sec("E. FY2026 錨點（1Q–3Q26 實績＋4Q26 公司指引或假設）")
for k, lab, unit, v, tag, s1, s2 in C["q_inputs"]:
    inp(k, lab, unit, None, v, None, tag, s1, s2)

sec("F. 共識與券商（只作 Recon 對帳，不作校準）")
for k, lab, unit, v, s1, s2 in C["cons"]:
    inp(k, lab, unit, None, v, None, "Interested-party", s1, s2)

sec("G. 價格與市場倍數")
inp("px", f"股價（{C['px_date']} 收盤）", "US$", None, C["px"], None, "Interested-party", C["px_src"], C["px_src2"])
inp("pe_ttm", f"近 12 月 P/E（現價 ÷ 近四季 GAAP EPS {C['eps_ttm']}）", "x", None, f"=px/{C['eps_ttm']}", None, "Derived", C["eps_ttm_src"], "為負值（含 Escrowed Shares 非現金評價損失）；僅供參考")
inp("fpe", C["fpe_lab"], "x", None, C["fpe"], None, "Interested-party", C["fpe_src"], "—")
inp("cons_feps", C["cfeps_lab"], "US$", None, C["cfeps"], None, "Derived", C["cfeps_src"], "—")
inp("cons_tp", "市場共識目標價（賣方平均）", "US$", None, C["cons_tp"], None, "Interested-party", C["cons_tp_src"], C["cons_tp_src2"])
inp("nim_5", "近三年平均淨利率（FY2023–25；對照用）", "%", None, C["nim5"], None, "Derived", C["nim5_src"], "十年平均（2016–2025）：未找到；收斂法依工作單改用最後一年 × 0.8")

sec("H. 目標價與收斂法參數（黃底；Assumed，Andy 可改）")
TP_KEYS = []
for s, w in (("A", 0.3), ("B", 0.5), ("C", 0.2)):
    inp(f"tp_w{s}", f"目標價情境權重 {s}（A 樂觀／B 基準／C 保守）", "%", None, w, None, "Assumed", "Andy 可改（工作單 2026-10-25 預設 A 0.3／B 0.5／C 0.2）", "三者合計應為 100%")
    TP_KEYS.append(f"tp_w{s}")
inp("tp_g", "PL 延伸期營收年增的收斂目標（情境 B；FY2032；DCF 永續成長率見 g_tv）", "%", None, 0.04, None, "Assumed", "Andy 可改（工作單預設 4%）", "只決定 FY2030–FY2032 營收收斂路徑")
inp("tp_w_dcf", "目標價中 DCF 的權重；其餘為 Comparable", "%", None, 0.5, None, "Assumed", "Andy 可改（2026-10-25 11:53 定 DCF／Comparable 各半）", "—")
inp("fade_gA", "情境 A：FY2032 營收年增收斂目標（tp_g＋2pt）", "%", None, "=tp_g+0.02", None, "Assumed", "工作單 2026-10-25 11:20 收斂法", "FY2030 起由 FY2029 年增線性收斂")
inp("fade_gB", "情境 B：FY2032 營收年增收斂目標（＝tp_g）", "%", None, "=tp_g", None, "Assumed", "同上", "—")
inp("fade_gC", "情境 C：FY2032 營收年增收斂目標（tp_g−2pt）", "%", None, "=tp_g-0.02", None, "Assumed", "同上", "—")
inp("fade_mA", "情境 A：FY2032 淨利率收斂目標（＝情境 A FY2029 淨利率）", "%", None, 0, None, "Derived", "PL 頁情境 A FY2029（公式於建置時連結）", "—")
inp("fade_mB", "情境 B：FY2032 淨利率收斂目標（FY2029 × 0.9）", "%", None, 0, None, "Assumed", "十年平均未找到，依工作單以最後一年 × 0.8 代替（B 取中點 0.9）", "對轉虧為盈中的公司偏保守，見 README 待判斷")
inp("fade_mC", "情境 C：FY2032 淨利率收斂目標（FY2029 × 0.8）", "%", None, 0, None, "Assumed", "同上", "—")
for _row in FD.in_rows(C["dcf_P"]):
    rows.append(_row)
inp("beta_u", "產業無槓桿 β（採用）：自算 US 半導體 同業中位數 Blume 調整（0.67 × 中位數＋0.33）", "x", None, 1.884, None, "Analogy",
    "自算產業 β：US 半導體 同業 13 家（NVDA AMD AVGO MRVL QCOM INTC MU ARM CRDO ALAB RMBS MTSI GFS；含 Intel 本身）無槓桿 β 中位數 2.32（Yahoo Finance 月報酬 2021-11～2026-10，對 S&P 500 迴歸，Hamada 去槓桿、現金修正）→ Blume 調整 1.884（docs/plan/自算產業β_20261105.md）",
    "Blume 調整 0.67β＋0.33（Andy 2026-11-05 22:55）；個股 Yahoo beta 2.43（YFW）未採用")
inp("beta_du", "β 敏感度幅度（低／高＝β_u ∓ 本值，各自重新加槓桿）", "x", None, 0.3, None, "Assumed", "工作單第二輪修正（Andy 2026-11-05 23:16 定案）：β 敏感度改為 β_u ± 0.3", "—")
inp("beta_de", "市值負債權益比 D/E（重新加槓桿用）", "x", None, "=(debt_v+lease_v)/(px*sh_v)", None, "Derived", "＝(債務本金 debt_v ＋ 租賃 lease_v) ÷ (股價 px × 股數 sh_v)", "與 WACC 權重同一口徑")
TP_KEYS += ["tp_g", "tp_w_dcf", "fade_gA", "fade_gB", "fade_gC", "fade_mA", "fade_mB", "fade_mC"]

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
FD.yellow(IN, FD.JUDG + ("beta_u",))
def defname(name, ref):
    wb.defined_names[name] = DefinedName(name, attr_text=ref)

# ---------------- Q：實績與 FY2026 錨點（US$B） ----------------
Q = wb.create_sheet("Q")
Q["A1"] = "Q — 實績與 FY2026 錨點（十億美元；EPS 為美元；股數為十億股）。藍字＝輸入（公司新聞稿）；黑字＝公式。分部口徑：Intel Products（CCPG、DCAI）、Intel Foundry、All Other、Corporate Unallocated、分部間沖銷"
Q["A2"] = ("FY2026 錨點＝1Q26、2Q26、3Q26 實績（公司新聞稿；2Q26 起 CCG 更名 CCPG，歷史季已按新名稱列示）＋4Q26 公司指引（營收 150–160 億、GAAP 毛利率 37.1%、營業費用約 49 億、稅率 31%）。"
           "Escrowed Shares 評價（美國商務部 CHIPS 協議的託管股份，按公允價值評價，非現金）計入「利息及其他淨額」，2Q26 損失 125.29 億，FY2027 起不預測")
hdr(Q, 3, ["項目", "FY2025A", "1Q26（實績）", "2Q26（實績）", "3Q26（實績）", "4Q26（模型）", "FY2026 錨點", "標記", "來源1", "來源2／說明"])
S1 = "公司新聞稿：4Q25／FY2025（2026-01-22）、1Q26（2026-04-23）、2Q26（2026-07-23）、3Q26（2026-10-29）；網址見 Sources 頁"
def qrow(r, label, vals, fmt="#,##0.000", tag="Verified", s1=S1, s2="—", ann=None):
    Q.cell(r, 1, label)
    for col, v in vals.items():
        c = Q[f"{col}{r}"]; c.value = v; c.number_format = fmt
        if not (isinstance(v, str) and v.startswith("=")): c.font = BLUE
    if ann != "":
        Q[f"G{r}"] = ann or f"=SUM(C{r}:F{r})"; Q[f"G{r}"].number_format = fmt
    Q.cell(r, 8, tag); Q.cell(r, 9, s1); Q.cell(r, 10, s2)
ALLC = "BCDEFG"
qrow(4, "營收", dict(B=52.853, C=13.577, D=16.128, E=16.513, F="=q4_rev"), tag="Verified／Interested-party（4Q 指引）", s2="4Q26＝公司指引中點")
qrow(5, "CCPG（客戶端運算與實體 AI）", dict(B=32.228, C=7.727, D=8.877, E=8.097, F="=q4_ccpg"), tag="Verified／Assumed", s2="FY2025 為 CCG；4Q26 拆分為假設")
qrow(6, "DCAI（資料中心與 AI）", dict(B=16.919, C=5.052, D=6.262, E=7.318, F="=q4_dcai"), tag="Verified／Assumed")
qrow(7, "All Other", dict(B=3.563, C=0.628, D=0.701, E=0.692, F="=q4_oth"), tag="Verified／Assumed", s2="含 Mobileye 等；Altera 自 2025-09 起不再合併")
qrow(8, "Intel Foundry（含對內銷售）", dict(B=17.826, C=5.421, D=5.765, E=6.697, F="=q4_iota*(F5+F6)+F10"), tag="Verified／Derived", s2="4Q26＝內部銷售占比 × Intel Products＋外部代工")
qrow(9, "分部間沖銷", dict(B=-17.683, C=-5.251, D=-5.477, E=-6.291, F="=F10-F8"), tag="Verified／Derived")
qrow(10, "外部代工營收（Foundry＋沖銷）", dict(B="=B8+B9", C="=C8+C9", D="=D8+D9", E="=E8+E9", F="=F4-F5-F6-F7"), tag="Derived", s2="Intel Foundry 對 Intel 以外客戶的營收（含先進封裝）；4Q26 為指引中點扣其他分部的餘額")
qrow(11, "檢查：分部合計 − 合併營收", {c: f"={c}5+{c}6+{c}7+{c}8+{c}9-{c}4" for c in "BCDEF"}, tag="Derived", s2="應為 0")
qrow(12, "CCPG 營業利益", dict(B=9.317, C=2.516, D=2.343, E=2.285, F="=q4_ccpg_oi"), tag="Verified／Assumed")
qrow(13, "DCAI 營業利益", dict(B=3.422, C=1.542, D=2.474, E=3.039, F="=q4_dcai_oi"), tag="Verified／Assumed")
qrow(14, "Intel Foundry 營業利益（虧損）", dict(B=-10.318, C=-2.437, D=-2.089, E=-1.853, F="=q4_fnd_oi"), tag="Verified／Assumed")
qrow(15, "All Other 營業利益", dict(B=0.264, C=0.102, D=0.230, E=0.179, F="=q4_oth_oi"), tag="Verified／Assumed")
qrow(16, "Corporate Unallocated", dict(B=-5.518, C=-5.080, D=-1.416, E=-1.058, F="=F18-F12-F13-F14-F15-F17"), tag="Verified／Derived", s2="4Q26 為餘額（指引營業利益減各分部）；1Q26 含重組 40.70 億與商譽減損")
qrow(17, "分部間沖銷的營業利益", dict(B=0.619, C=0.221, D=0.254, E=-0.070, F=0), tag="Verified／Assumed")
qrow(18, "營業利益（合併 GAAP）", dict(B=-2.214, C=-3.136, D=1.796, E=2.522, F="=F20-F21"), tag="Verified／Derived", s2="Yahoo 3Q26 營業利益 2,576、FY2025 −23 與公司數不同，以公司為準（見 README 對照）")
qrow(19, "檢查：分部營業利益合計 − 合併營業利益", {c: f"={c}12+{c}13+{c}14+{c}15+{c}16+{c}17-{c}18" for c in "BCDEF"}, tag="Derived", s2="應為 0")
qrow(20, "毛利", dict(B=18.375, C=5.347, D=6.509, E=7.096, F="=F4*q4_gm"), tag="Verified／Interested-party（4Q 指引）")
qrow(21, "營業費用（含重組）", dict(B=20.589, C=8.483, D=4.713, E=4.574, F="=q4_opx"), tag="Verified／Interested-party（4Q 指引）", s2="研發＋行銷、一般及管理＋重組及其他")
qrow(22, "毛利率", {c: f"={c}20/{c}4" for c in ALLC}, fmt="0.0%", ann="", tag="Derived")
Q["G22"] = "=G20/G4"; Q["G22"].number_format = "0.0%"
qrow(23, "權益投資損益", dict(B=0.514, C=-0.072, D=-0.039, E=0.330, F=0), tag="Verified／Assumed")
qrow(24, "利息及其他淨額（含 Escrowed Shares 評價）", dict(B=3.257, C=-0.738, D=-12.576, E=1.130, F="=q4_int"), tag="Verified／Assumed", s2="FY2025 含 Altera 出售利得與 Escrowed Shares 評價等；2Q26 含評價損失 125.29 億")
qrow(25, "  其中：Escrowed Shares 評價損益（非現金）", dict(C=-1.090, D=-12.529, E=1.005, F=0), tag="Verified／Derived", s2="2Q26、3Q26 取新聞稿；1Q26＝9M26 損失 126.14 億 − 2Q26 125.29 億 + 3Q26 利得 10.05 億；FY2025 全年未取得（不填）")
qrow(26, "營業外淨額（權益投資損益＋利息及其他）", {c: f"={c}23+{c}24" for c in ALLC}, tag="Derived", ann="")
Q["G26"] = "=G23+G24"; Q["G26"].number_format = "#,##0.000"
qrow(27, "稅前利益", {c: f"={c}18+{c}26" for c in ALLC}, tag="Derived", s2="對公司揭露：FY2025 1.557、1Q26 −3.946、2Q26 −10.819、3Q26 3.982（見下方檢查）", ann="")
Q["G27"] = "=G18+G26"; Q["G27"].number_format = "#,##0.000"
qrow(28, "所得稅", dict(B=1.531, C=0.335, D=0.029, E=0.462, F="=MAX(0,F27)*q4_tax"), tag="Verified／Interested-party（4Q 指引）")
qrow(29, "  有效稅率", {c: f"=IF({c}27>0,{c}28/{c}27,\"n/a\")" for c in ALLC}, fmt="0.0%", ann="", tag="Derived")
Q["G29"] = '=IF(G27>0,G28/G27,"n/a")'; Q["G29"].number_format = "0.0%"
qrow(30, "本期淨利（含非控制權益）", {c: f"={c}27-{c}28" for c in ALLC}, tag="Derived", ann="")
Q["G30"] = "=G27-G28"; Q["G30"].number_format = "#,##0.000"
qrow(31, "非控制權益損益", dict(B=0.293, C=-0.553, D=0.185, E=0.200, F="=q4_nci"), tag="Verified／Assumed")
qrow(32, "歸屬 Intel 淨利", {c: f"={c}30-{c}31" for c in ALLC}, tag="Derived", ann="")
Q["G32"] = "=G30-G31"; Q["G32"].number_format = "#,##0.000"
qrow(33, "稀釋加權股數（十億股）", dict(B=4.530, C=5.083, D=5.104, E=5.307, F="=q_sh"), fmt="0.000", ann="=AVERAGE(C33:F33)", tag="Verified／Derived", s2="1Q26、2Q26 虧損期稀釋＝基本；4Q25 4.856")
qrow(34, "EPS（淨利 ÷ 稀釋股數）", {c: f"={c}32/{c}33" for c in "BCDEF"}, fmt="#,##0.00", ann="=G32/G33", tag="Derived", s2="3Q26 公司稀釋 EPS 0.60 與淨利 ÷ 稀釋股數 0.63 不同：公司稀釋分子另有調整（金額未取得，見 README 待判斷）")
qrow(35, "  公司揭露稀釋 EPS", dict(B=-0.06, C=-0.73, D=-2.16, E=0.60), fmt="#,##0.00", ann="", tag="Verified", s2="4Q25 −0.12；FY2026 全年 −2.19（四季加總，含 4Q26 指引 0.10）")
qrow(36, "折舊（現金流量表）", dict(B=10.757, C=2.902, D="=(8.933-2.902)/2", E="=(8.933-2.902)/2", F="=q4_da"), tag="Verified／Derived", s2="9M26 折舊 89.33 億、1Q26 29.02 億（新聞稿）；2Q26、3Q26 平分（兩季合計 60.31 億），Derived")
qrow(37, "總資本支出（投資活動＋融資活動的 PP&E 增加）", dict(B=17.672, C=4.963, D=2.652, E=4.219, F="=q4_cx"), tag="Verified／Assumed", s2="FY2025：投資活動 146.46 億＋融資活動 30.26 億；季值為公司揭露 gross capex；不計政府補助（FY2025 15.77 億）與夥伴挹注（FY2025 51.08 億）")
qrow(38, "內部銷售占 Intel Products 營收比", {c: f"=-{c}9/({c}5+{c}6)" for c in ALLC}, fmt="0.0%", ann="", tag="Derived")
Q["G38"] = "=-G9/(G5+G6)"; Q["G38"].number_format = "0.0%"
Q.cell(40, 1, "FY2026 錨點的檢查（參考）").font = BOLD
QCH = [("檢查：稅前利益 自算 − 公司（FY2025／1Q26／2Q26／3Q26 合計絕對差，應約為 0）", "=ABS(B27-1.557)+ABS(C27+3.946)+ABS(D27+10.819)+ABS(E27-3.982)", "0.000"),
       ("檢查：歸屬 Intel 淨利 自算 − 公司（FY2025 −0.267、1Q26 −3.728、2Q26 −11.033、3Q26 3.320；合計絕對差，應約為 0）", "=ABS(B32+0.267)+ABS(C32+3.728)+ABS(D32+11.033)+ABS(E32-3.320)", "0.000"),
       ("FY2026 營收年增", "=G4/B4-1", "0.0%"), ("3Q26 營收季增（公司：年增 21%）", "=E4/D4-1", "0.0%"), ("FY2026 毛利率", "=G22", "0.0%"),
       ("FY2026 營業利益（GAAP；含 1Q26 重組）", "=G18", "0.000"), ("FY2026 EPS（淨利 ÷ 平均稀釋股數；含 Escrowed Shares 評價損失）", "=G34", "0.00"),
       ("FY2026 EPS 不含 Escrowed Shares 評價（參考）", "=(G32-G25)/G33", "0.00"),
       ("FY2026 Foundry 營業利益率（占 Foundry 營收）", "=G14/G8", "0.0%"), ("FY2026 總資本支出占營收", "=G37/G4", "0.0%"),
       ("9M26 營收合計（1Q＋2Q＋3Q；公司 3Q26 累計＝46.218）", "=C4+D4+E4", "0.000")]
for i, (lab, f, fm) in enumerate(QCH, 41):
    Q.cell(i, 1, lab); Q[f"B{i}"] = f; Q[f"B{i}"].number_format = fm
QCH_ROWS = len(QCH)
for nm, rr in (("rev", 4), ("ccpg", 5), ("dcai", 6), ("oth", 7), ("fnd", 8), ("elim", 9), ("ext", 10), ("ccpg_oi", 12), ("dcai_oi", 13), ("fnd_oi", 14), ("oth_oi", 15), ("cu", 16),
               ("elimoi", 17), ("oi", 18), ("gp", 20), ("opx", 21), ("eqg", 23), ("nonop", 26), ("tax", 28), ("mi", 31), ("ni", 32), ("sh", 33), ("eps", 34), ("da", 36), ("cx", 37), ("iota", 38)):
    defname(f"a25_{nm}", f"Q!$B${rr}"); defname(f"a26_{nm}", f"Q!$G${rr}")
Q.column_dimensions["A"].width = 60
for c in "BCDEFG": Q.column_dimensions[c].width = 14
Q.column_dimensions["H"].width = 30; Q.column_dimensions["I"].width = 50; Q.column_dimensions["J"].width = 70
Q.freeze_panes = "B4"
# ---------------- 引擎：一個情境的完整推導（PL 與 Sens 共用） ----------------
FYS = (2025, 2026, 2027, 2028, 2029); FC = {2025: "C", 2026: "D", 2027: "E", 2028: "F", 2029: "G", 2030: "H", 2031: "I", 2032: "J"}
FADE = (2030, 2031, 2032); FADEFILL = PatternFill("solid", fgColor="E7E6E6")
BU = (2026, 2027, 2028, 2029); F3 = (2027, 2028, 2029)
DEFP = dict(scn="A", ai_adj="ai_adj", ext_adj="ext_adj", fm_adj="fm_adj", dcm_adj="dcm_adj", ccm_adj="ccm_adj", ccg_adj="ccg_adj", cx_adj="cx_adj", opx_adj="opx_adj",
            nonop_f="nonop_f", tax="tax", nci_r="nci_r", fm_self=False, wm=None)
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
    pv = lambda y: FC[y - 1]
    sec_("A. Intel Products：CCPG 與 DCAI（D 欄＝FY2026 錨點，由 Q 頁連入）")
    row("ccg", "CCPG 營收（客戶端運算與實體 AI）", "US$B", {"C": "=a25_ccpg", "D": "=a26_ccpg", **{FC[y]: f"={pv(y)}{r}*(1+ccg{yy(y)}_{s}+{P['ccg_adj']})" for y in F3}}, "#,##0.00")
    row("ccm", "CCPG 營業利益率", "%", {"C": "=a25_ccpg_oi/a25_ccpg", "D": "=a26_ccpg_oi/a26_ccpg", **{FC[y]: f"=ccm{yy(y)}_{s}+{P['ccm_adj']}" for y in F3}}, "0.0%")
    row("ccoi", "CCPG 營業利益", "US$B", {"C": "=a25_ccpg_oi", "D": "=a26_ccpg_oi", **{FC[y]: f"={FC[y]}{R['ccg']}*{FC[y]}{R['ccm']}" for y in F3}})
    if P["wm"]:   # v0.13：Wattonomics 錨定（DCAI AI 相關營收 FY2027 起 × wta；FY2028 起先除去前一年係數再乘當年，避免連乘）
        w_ = P["wm"]
        row("dai", "DCAI 中 AI 相關營收（FY2028 起 × 同字母顆數係數；FY2027 起 × Wattonomics 錨定係數）", "US$B", {"D": "=a26_dcai*ai_sh26", "E": f"=D{r}*(1+aig27_{s})*{P['ai_adj']}*{w_[2027]}",
            "F": f"=E{r}/{w_[2027]}*(1+aig_L)*um28_{s}*{w_[2028]}", "G": f"=F{r}/um28_{s}/{w_[2028]}*(1+aig_L)*um29_{s}*{w_[2029]}"})
    else:
        row("dai", "DCAI 中 AI 相關營收（FY2028 起 × 同字母顆數係數）", "US$B", {"D": "=a26_dcai*ai_sh26", "E": f"=D{r}*(1+aig27_{s})*{P['ai_adj']}",
            "F": f"=E{r}*(1+aig_L)*um28_{s}", "G": f"=F{r}/um28_{s}*(1+aig_L)*um29_{s}"})
    row("dnai", "DCAI 其他（非 AI 相關）營收", "US$B", {"D": f"=a26_dcai-D{R['dai']}", **{FC[y]: f"={pv(y)}{r}*(1+dng{yy(y)}_{s})" for y in F3}})
    row("dcr", "DCAI 營收（資料中心與 AI）", "US$B", {"C": "=a25_dcai", "D": "=a26_dcai", **{FC[y]: f"={FC[y]}{R['dai']}+{FC[y]}{R['dnai']}" for y in F3}})
    row("dcm", "DCAI 營業利益率", "%", {"C": "=a25_dcai_oi/a25_dcai", "D": "=a26_dcai_oi/a26_dcai", **{FC[y]: f"=dcm{yy(y)}_{s}+{P['dcm_adj']}" for y in F3}}, "0.0%")
    row("dcoi", "DCAI 營業利益", "US$B", {"C": "=a25_dcai_oi", "D": "=a26_dcai_oi", **{FC[y]: f"={FC[y]}{R['dcr']}*{FC[y]}{R['dcm']}" for y in F3}})
    row("prod", "Intel Products 營收（CCPG＋DCAI）", "US$B", {FC[y]: f"={FC[y]}{R['ccg']}+{FC[y]}{R['dcr']}" for y in FYS})
    row("oth", "All Other 營收", "US$B", {"C": "=a25_oth", "D": "=a26_oth", **{FC[y]: f"={pv(y)}{r}*(1+oth_g)" for y in F3}})
    row("othoi", "All Other 營業利益", "US$B", {"C": "=a25_oth_oi", "D": "=a26_oth_oi", **{FC[y]: f"={FC[y]}{R['oth']}*oth_m" for y in F3}})
    sec_("B. Intel Foundry：內部銷售、外部代工營收、營業利益率路徑")
    row("ext", "外部代工營收（18A／14A、先進封裝）", "US$B", {"C": "=a25_ext", "D": "=a26_ext", **{FC[y]: f"=ext{yy(y)}_{s}*{P['ext_adj']}" for y in F3}})
    row("fnd", "Intel Foundry 營收（含對內銷售）＝內部銷售占比 × Intel Products ＋ 外部", "US$B", {"C": "=a25_fnd", "D": "=a26_fnd", **{FC[y]: f"=iota*{FC[y]}{R['prod']}+{FC[y]}{R['ext']}" for y in F3}})
    row("elim", "分部間沖銷", "US$B", {"C": "=a25_elim", "D": "=a26_elim", **{FC[y]: f"={FC[y]}{R['ext']}-{FC[y]}{R['fnd']}" for y in F3}})
    row("fm", "Intel Foundry 營業利益率（占 Foundry 營收）", "%", {"C": "=a25_fnd_oi/a25_fnd", "D": "=a26_fnd_oi/a26_fnd", **{FC[y]: f"=fm{yy(y)}_{s}+{P['fm_adj']}" for y in F3}}, "0.0%")
    row("fndoi", "Intel Foundry 營業利益（虧損）", "US$B", {"C": "=a25_fnd_oi", "D": "=a26_fnd_oi", **{FC[y]: f"={FC[y]}{R['fnd']}*{FC[y]}{R['fm']}" for y in F3}})
    row("cu", "Corporate Unallocated", "US$B", {"C": "=a25_cu", "D": "=a26_cu", **{FC[y]: "=cu_f" for y in F3}})
    row("elimoi", "分部間沖銷的營業利益", "US$B", {"C": "=a25_elimoi", "D": "=a26_elimoi", **{FC[y]: "=0" for y in F3}})
    sec_("C. 合併損益")
    row("rev", "營收（CCPG＋DCAI＋All Other＋外部代工）", "US$B", {FC[y]: f"={FC[y]}{R['ccg']}+{FC[y]}{R['dcr']}+{FC[y]}{R['oth']}+{FC[y]}{R['ext']}" for y in FYS})
    row("g", "營收年增", "%", {FC[y]: f"={FC[y]}{R['rev']}/{FC[y - 1]}{R['rev']}-1" for y in FYS[1:]}, "0.0%")
    row("oi", "營業利益（GAAP；分部營業利益合計）", "US$B", {FC[y]: f"={FC[y]}{R['ccoi']}+{FC[y]}{R['dcoi']}+{FC[y]}{R['fndoi']}+{FC[y]}{R['othoi']}+{FC[y]}{R['cu']}+{FC[y]}{R['elimoi']}" for y in FYS})
    row("opx", "營業費用（含重組）", "US$B", {"C": "=a25_opx", "D": "=a26_opx", "E": f"=opx27*{P['opx_adj']}", **{FC[y]: f"={FC[y - 1]}{r}*(1+opx_g)" for y in (2028, 2029)}})
    row("gp", "毛利（＝營業利益＋營業費用）", "US$B", {FC[y]: f"={FC[y]}{R['oi']}+{FC[y]}{R['opx']}" for y in FYS}, note="毛利由分部營業利益加營業費用反推（公司不揭露分部毛利）")
    row("gm", "毛利率", "%", {FC[y]: f"={FC[y]}{R['gp']}/{FC[y]}{R['rev']}" for y in FYS}, "0.00%")
    row("oim", "營業利益率", "%", {FC[y]: f"={FC[y]}{R['oi']}/{FC[y]}{R['rev']}" for y in FYS}, "0.00%")
    row("nonop", "營業外淨額（權益投資＋利息及其他）", "US$B", {"C": "=a25_nonop", "D": "=a26_nonop", **{FC[y]: f"={P['nonop_f']}" for y in F3}}, note="FY2026 含 Escrowed Shares 評價損失（非現金）；FY2027 起不預測評價")
    row("pti", "稅前利益", "US$B", {FC[y]: f"={FC[y]}{R['oi']}+{FC[y]}{R['nonop']}" for y in FYS})
    row("tax", "所得稅", "US$B", {"C": "=a25_tax", "D": "=a26_tax", **{FC[y]: f"=MAX(0,{FC[y]}{R['pti']})*{P['tax']}" for y in F3}}, note="稅前為負時不計稅負")
    row("ncib", "非控制權益餘額（年底）", "US$B", {"D": "=nci_bs", **{FC[y]: f"={FC[y - 1]}{r}+nci_g" for y in F3}}, note="以 2026-09-26 帳面為 FY2026 年底；之後每年加夥伴挹注（In 頁 nci_g）")
    row("mi", "非控制權益損益", "US$B", {"C": "=a25_mi", "D": "=a26_mi", **{FC[y]: f"={P['nci_r']}*{FC[y - 1]}{R['ncib']}" for y in F3}})
    row("ni", "歸屬 Intel 淨利", "US$B", {"C": "=a25_ni", **{FC[y]: f"={FC[y]}{R['pti']}-{FC[y]}{R['tax']}-{FC[y]}{R['mi']}" for y in BU}})
    row("shr", "稀釋加權股數", "十億股", {"C": "=a25_sh", "D": "=a26_sh", **{FC[y]: (f"=q_sh*(1+sh_g)" if y == 2027 else f"={FC[y - 1]}{r}*(1+sh_g)") for y in F3}}, "0.00000")
    row("eps", "EPS（稀釋）", "US$", {FC[y]: f"={FC[y]}{R['ni']}/{FC[y]}{R['shr']}" for y in FYS})
    row("epsg", "EPS 年增", "%", {FC[y]: f"=IF({FC[y - 1]}{R['eps']}>0,{FC[y]}{R['eps']}/{FC[y - 1]}{R['eps']}-1,\"n/a\")" for y in FYS[1:]}, "0.0%")
    row("nim", "淨利率（歸屬 Intel 淨利 ÷ 營收）", "%", {FC[y]: f"={FC[y]}{R['ni']}/{FC[y]}{R['rev']}" for y in FYS}, "0.00%")
    sec_("D. 資本支出與折舊（DCF 逐年取用）")
    row("da", "折舊", "US$B", {"C": "=a25_da", "D": "=a26_da", **{FC[y]: f"={FC[y - 1]}{r}*(1+da_g)" for y in F3}})
    row("cx", "總資本支出（投資活動＋融資活動的 PP&E 增加）", "US$B", {"C": "=a25_cx", "D": "=a26_cx", **{FC[y]: f"=cx{yy(y)}_{s}*{P['cx_adj']}" for y in F3}}, note="不計合資夥伴挹注與政府補助（保守；夥伴挹注以非控制權益餘額與報酬率處理）")
    row("cxg", "總資本支出年增", "%", {FC[y]: f"={FC[y]}{R['cx']}/{FC[y - 1]}{R['cx']}-1" for y in BU}, "0.0%")
    row("cxs", "總資本支出 ÷ 營收", "%", {FC[y]: f"={FC[y]}{R['cx']}/{FC[y]}{R['rev']}" for y in FYS}, "0.0%")
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
        ws[f"{c}{R['epsg']}"] = f"=IF({pc}{R['eps']}>0,{c}{R['eps']}/{pc}{R['eps']}-1,\"n/a\")"
        for key, fmt_ in (("g", "0.0%"), ("rev", "#,##0.00"), ("nim", "0.00%"), ("ni", "#,##0.00"), ("shr", "0.00000"), ("eps", "#,##0.00"), ("epsg", "0.0%")):
            ws[f"{c}{R[key]}"].number_format = fmt_; ws[f"{c}{R[key]}"].fill = FADEFILL
    ws.cell(R["rev"], 11, "FY2030 起為收斂法：營收年增由 FY2029 線性收斂到 fade_g、淨利率收斂到 fade_m（FY2032 達到），非驅動因子推導")
    R["end"] = r
    return R

# ---------------- PL ----------------
PL = wb.create_sheet("PL")
PL["A1"] = (f"PL — {C['name']}損益（十億美元；EPS 美元、稀釋；歸屬 Intel）。FY2025A 實績、FY2026 錨點（三情境相同），FY2027–FY2029 由驅動因子推導，"
            "FY2030–FY2032 為收斂法。FY(t)≈CY(t)。三情境各一個完整推導區塊（本頁下方），頂端為摘要")
PB = {}; r0 = 90
for s in SC:
    # 三層協調 C4 另一側讀數建置（tools/coord/alt_run.py；SIL.COORD_ALT）：情境 B、C 於 CY2027 起 DCAI 中 AI 相關營收 × wta；情境 A 不縮放。正式建置不乘（wm=None）
    PB[s] = engine(PL, r0, C["scn_t"][s], scn=s, wm=({2027: "wta_27", 2028: "wta_28", 2029: "wta_29"} if (SIL.COORD_ALT and s in ("B", "C")) else None))
    r0 = PB[s]["end"] + 2
hdr(PL, 3, ["情境／項目", "單位", "FY2025A", "FY2026（錨點）", "FY2027", "FY2028", "FY2029", "FY2030（收斂）", "FY2031（收斂）", "FY2032（收斂）", "說明"])
sr = 4; SUMR = {}
SUMK = [("rev", "營收", "US$B", "#,##0.0"), ("g", "  營收年增", "%", "0.0%"), ("ccg", "  CCPG 營收", "US$B", "#,##0.0"), ("dcr", "  DCAI 營收", "US$B", "#,##0.0"),
        ("dai", "    其中 AI 相關（模型拆分）", "US$B", "#,##0.0"), ("oth", "  All Other 營收", "US$B", "#,##0.0"), ("ext", "  外部代工營收", "US$B", "#,##0.0"),
        ("fnd", "  Intel Foundry 營收（含對內）", "US$B", "#,##0.0"), ("fndoi", "  Intel Foundry 營業利益", "US$B", "#,##0.0"), ("fm", "  Foundry 營業利益率", "%", "0.0%"),
        ("gm", "毛利率", "%", "0.00%"), ("oim", "營業利益率", "%", "0.00%"), ("ni", "歸屬 Intel 淨利", "US$B", "#,##0.0"), ("nim", "  淨利率", "%", "0.00%"),
        ("eps", "EPS", "US$", "#,##0.00"), ("epsg", "  EPS 年增", "%", "0.0%"), ("cx", "總資本支出", "US$B", "#,##0.0"), ("da", "折舊", "US$B", "#,##0.0")]
LONGK = ("rev", "g", "ni", "nim", "eps", "epsg")
NOC = {"dai"}   # FY2025A 無 AI 相關拆分
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
for lab, f in (("  三情境：營收差（US$B）", "+".join(f"ABS(D{PB[s]['rev']}-a26_rev)" for s in SC)),
               ("  三情境：營業利益差（US$B）", "+".join(f"ABS(D{PB[s]['oi']}-a26_oi)" for s in SC)),
               ("  三情境：毛利差（US$B）", "+".join(f"ABS(D{PB[s]['gp']}-a26_gp)" for s in SC)),
               ("  三情境：EPS 差（US$）", "+".join(f"ABS(D{PB[s]['eps']}-a26_eps)" for s in SC))):
    PL.cell(sr, 1, lab); PL[f"D{sr}"] = "=" + f; PL[f"D{sr}"].number_format = "0.0000"; CHK.append(sr); sr += 1
assert sr < 90, sr
PL.column_dimensions["A"].width = 62; PL.column_dimensions["B"].width = 10; PL.column_dimensions["K"].width = 70
for c in "CDEFGHIJ": PL.column_dimensions[c].width = 13
PL.cell(2, 1, "灰底＝FY2030–FY2032 收斂法（上游顆數只到 CY2029）：營收年增線性收斂到 fade_g、淨利率線性收斂到 fade_m（In 頁 H 節），非驅動因子推導")

# ---------------- Units ----------------
UN = wb.create_sheet("Units")
UN["A1"] = C["units_a1"]
cyc = lambda y: chr(ord("C") + y - 2026)
hdr(UN, 3, ["項目", "單位", "FY2026", "FY2027", "FY2028", "FY2029", "說明"])
ur = 4; UR = {}
UN.cell(ur, 1, "1. 營收與利益率驅動（三情境）").font = BOLD; ur += 1
for key, lab, unit, fmt in (("dai", "DCAI 中 AI 相關營收", "US$B", "#,##0.00"), ("dcr", "DCAI 營收", "US$B", "#,##0.00"), ("ccg", "CCPG 營收", "US$B", "#,##0.00"),
                            ("ext", "外部代工營收", "US$B", "#,##0.00"), ("fm", "Foundry 營業利益率", "%", "0.0%"), ("fndoi", "Foundry 營業利益", "US$B", "#,##0.00")):
    for s in SC:
        UN.cell(ur, 1, f"{lab}（情境 {s}）"); UN.cell(ur, 2, unit); UR[f"{key}_{s}"] = ur
        for y in BU:
            UN[f"{cyc(y)}{ur}"] = f"=PL!{FC[y]}{PB[s][key]}"; UN[f"{cyc(y)}{ur}"].number_format = fmt
        UN.cell(ur, 7, "PL 頁各情境推導區塊" if s == "A" else "同上"); ur += 1
ur += 1
UN.cell(ur, 1, "2. 資本支出與折舊（三情境）").font = BOLD; ur += 1
for key, lab, unit, fmt in (("cx", "總資本支出", "US$B", "#,##0.00"), ("cxs", "總資本支出 ÷ 營收", "%", "0.0%"), ("da", "折舊", "US$B", "#,##0.00")):
    for s in SC:
        UN.cell(ur, 1, f"{lab}（情境 {s}）"); UN.cell(ur, 2, unit); UR[f"{key}_{s}"] = ur
        for y in BU:
            UN[f"{cyc(y)}{ur}"] = f"=PL!{FC[y]}{PB[s][key]}"; UN[f"{cyc(y)}{ur}"].number_format = fmt
        UN.cell(ur, 7, "PL 頁各情境推導區塊" if s == "A" else "同上"); ur += 1
ur += 1
UN.cell(ur, 1, "3. 校準值（只用 FY2026 已發生的實績與 4Q26 指引；不用共識、不用 FY2027 起的目標）").font = BOLD; ur += 1
for key, lab, f, fmt, note in (("fm26", "FY2026 Foundry 營業利益率", f"=PL!D{PB['A']['fm']}", "0.0%", "Q 頁：Foundry 營業利益 ÷ Foundry 營收"),
                               ("iota26", "FY2026 內部銷售占 Intel Products 營收比", "=a26_iota", "0.0%", "分部間沖銷 ÷ Intel Products 營收"),
                               ("ext26", "FY2026 外部代工營收", "=a26_ext", "#,##0.00", "Foundry 營收＋分部間沖銷")):
    UN.cell(ur, 1, lab); UN[f"C{ur}"] = f; UN[f"C{ur}"].number_format = fmt; UN.cell(ur, 7, note); UR[key] = ur; ur += 1
ur += 1
UN.cell(ur, 1, C["units_note"]).font = BOLD
UN_NOTE = ur
UN.column_dimensions["A"].width = 56; UN.column_dimensions["B"].width = 12
for c in "CDEF": UN.column_dimensions[c].width = 13
UN.column_dimensions["G"].width = 56
# ---------------- Val ----------------
V = wb.create_sheet("Val")
V["A1"] = "Val — P/E 估值（美元）。結論用「現價 ÷ FY2027、FY2028 EPS（模型）」；目標價＝DCF 與 Comparable 加權（第 0 節）；十年 P/E 規則值（第 3 節，年底落後 P/E）只作參考"
V["A2"] = f"循環位置：{C['name']}不是設備股，不套設備市場見頂年；以本模型營收最高的會計年度為見頂年（比照創意、聯發科；FY(t)≈CY(t)）。營收到 FY2029 仍在增加時，見頂年未知，判為中段"
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
V[f"E{O + 3}"] = "FY2026 三情境相同（錨點）"; V[f"E{O + 4}"] = "FY2027 三情境差在 Foundry 利益率、外部代工營收、DCAI 與資本支出；顆數情境係數從 CY2028 開始"
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
    V.cell(rr, 3, f'=IF(OR(B{RB0 + 5 + i}="",B{RB0 + 5 + i}<=0,B{RB0 + 5 + i}>200),"",B{RB0 + 5 + i})').number_format = "0.00"
    V.cell(rr, 4, f'=IF(C{RB0 + 5 + i}="","",C{RB0 + 5 + i})').number_format = "0.00"
PE10 = f"$B${T10 + 2}:$B${T10 + 11}"
V.cell(T10 + 12, 1, "年數"); V[f"B{T10 + 12}"] = f"=COUNT({PE10})"
block_end = VR.add_rule_block(V, RB0, [{"id": "P/E", "name": "P/E", "sources": [
    {"label": "companiesmarketcap（年底 P/E）", "vals": C["pe_cmc"], "tag": "Interested-party（資料商；2016–2025，2023 年起淨利近零或為負的年度剔除）"},
    ],
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

YRS6 = (2027, 2028, 2029, 2030, 2031, 2032)
V.cell(4, 1, "0. 目標價").font = BOLD
T = {}
def tset(r, key, label, fs, fmt, note=""):
    V.cell(r, 1, label); T[key] = r
    for c, f in zip("BCD", fs):
        if f is None: continue
        V[f"{c}{r}"] = f; V[f"{c}{r}"].number_format = fmt
    if note: V[f"E{r}"] = note
# 標準版 DCF（FCFF＋WACC；tools/fcff_dcf.py）。DCF 明細在 DCF 頁；EBIT：FY2027–29＝推導區塊營業利益；FY2030–32（收斂法只有淨利）由淨利還原；折舊與資本支出 FY2027–29 逐年取 PL 頁各情境
DCFW = wb.create_sheet("DCF")
DCFW["A1"] = ("DCF — 標準版（模組 v2）：無槓桿自由現金流 FCFF＝NOPAT＋折舊攤銷 − 資本支出 − 營運資金增加；WACC 折現到評價日 2026-12-31（期中慣例）；"
              "終值＝NOPAT ×（1＋g）×（1 − g ÷ RONIC）÷（WACC − g）。共用模組 tools/fcff_dcf.py；金額：十億美元")
DCFW["A2"] = ("EBIT：FY2027–FY2029 取各推導區塊「營業利益」；FY2030–FY2032（收斂法只有淨利）＝（歸屬 Intel 淨利＋FY2029 非控制權益損益）÷（1 − 稅率）− 營業外淨額。"
              "稅率＝In 頁 tax。營業外收支與 Escrowed Shares 評價不進 NOPAT，現金與投資另在評價時加入；非控制權益以帳面扣除。FY2027–29 折舊與資本支出取 PL 頁各情境逐年金額；"
              "之後資本支出比例線性收斂到 fc_cx_ss。營運資金：FY2025 年底占營收比。")
DCF_NEXT = [4]
def dpaths(blocks, pre, tax="tax", nonop="nonop_f", lab=None):
    out = []
    for s_ in SC:
        bl = blocks[s_]
        ebit = {y: (f"{pre}{FC[y]}{bl['oi']}" if y <= 2029 else
                    f"=({pre}{FC[y]}{bl['ni']}+{pre}$G${bl['mi']})/(1-{tax})-{nonop}") for y in YRS6}
        out.append(dict(name=s_, label=(lab or "情境") + f" {s_}（{SCN_NAME[s_]}）", rev={y: f"{pre}{FC[y]}{bl['rev']}" for y in YRS6}, rev0=f"{pre}D{bl['rev']}",
                        ebit=ebit, tax={y: tax for y in YRS6}, ebit_note="FY2030 起由淨利還原", stub_fcff="stub_fcff4q",
                        da={y: f"{pre}{FC[y]}{bl['da']}" for y in F3}, cx={y: f"{pre}{FC[y]}{bl['cx']}" for y in F3}))
    return out
def ddetail(blocks, pre, title, **kw):
    R_ = FD.write_detail(DCFW, DCF_NEXT[0], dpaths(blocks, pre, **kw), list(YRS6), 1.0, conv_from=2029, title=title, fmt="#,##0.00", stub=True)
    DCF_NEXT[0] = R_["end"] + 1
    return R_
def v12ref(R_):
    return {s_: f"DCF!$C${R_['paths'][s_]['v12']}" for s_ in SC}
RD_ = ddetail(PB, "PL!", "三情境 FCFF 明細（十億美元；會計年度≈日曆年；評價日 2026-12-31）")
DCF_NEXT[0] = FD.write_checks(DCFW, DCF_NEXT[0], nc_bs=C["dcf_ncbs"][0], nc_note=C["dcf_ncbs"][1]) + 1
_q0 = DCF_NEXT[0]
DCFW.cell(_q0, 1, "Stub 期 FCFF：Q 頁 4Q26（十億美元；2026-09-26 → 12-26 約一季；三情境相同）").font = BOLD
for _i, (_a, _f, _n) in enumerate((("營收 4Q26", "=Q!F4", "Q 頁（公司指引中點）"), ("營業利益 4Q26", "=Q!F18", "Q 頁（毛利率指引 − 營業費用指引）"),
                                    ("NOPAT＝營業利益 ×（1 − q4_tax）", f"=C{_q0 + 2}*(1-q4_tax)", "In 頁 q4_tax（公司 GAAP 稅率指引 31%）"),
                                    ("＋折舊 4Q26", "=Q!F36", "Q 頁"), ("− 總資本支出 4Q26", "=Q!F37", "Q 頁（假設）"),
                                    ("− 營運資金增加（fc_wc × 年化營收增量 4 ×（4Q26 − 3Q26））", f"=fc_wc*4*(C{_q0 + 1}-Q!E4)", "年化季營收變動；營收下降時為釋放"),
                                    ("4Q26 FCFF（＝stub 期 FCFF）", f"=C{_q0 + 3}+C{_q0 + 4}-C{_q0 + 5}-C{_q0 + 6}", "進淨現金滾到評價日")), 1):
    DCFW.cell(_q0 + _i, 1, _a); DCFW.cell(_q0 + _i, 2, _n); DCFW.cell(_q0 + _i, 3, _f).number_format = "#,##0.000"
wb.defined_names["stub_fcff4q"] = DefinedName("stub_fcff4q", attr_text=f"DCF!$C${_q0 + 7}")
DCF_NEXT[0] = _q0 + 9
DCFW.column_dimensions["A"].width = 58; DCFW.column_dimensions["B"].width = 34
for _c in "CDEFGH": DCFW.column_dimensions[_c].width = 12
SDCF = FD.write_summary(wb, V, 5, RD_, list(SC), ["tp_wA", "tp_wB", "tp_wC"], fmt="#,##0.00",
                        title="0.1 DCF（標準版：無槓桿自由現金流 FCFF＋WACC；明細見 DCF 頁）", beta_note="自算產業 β 重新加槓桿")
for _j, _h in enumerate(["項目", "A（樂觀）", "B（基準）", "C（保守）", "說明"], 1): V.cell(SDCF["h"], _j, _h)
T.update(dcfv=SDCF["dcfv"], dcf12=SDCF["v12"], dcftp=SDCF["dcf"])
V.cell(SDCF["stub"], 5, '="資產負債表日 "&bs_date&"；stub 期 FCFF＝Q 頁 4Q26 季度數（DCF 頁下方；三情境相同）"')
assert SDCF["dcf"] <= 24
_PE = lambda c, y: f"PL!{FC[y]}{SUMR[(dict(zip('BCD', SC))[c], 'eps')]}"
EPSR = {}
DCFW.column_dimensions["A"].width = 58; DCFW.column_dimensions["B"].width = 34
for _c in "CDEFGH": DCFW.column_dimensions[_c].width = 12
SDCF = FD.write_summary(wb, V, 5, RD_, list(SC), ["tp_wA", "tp_wB", "tp_wC"], fmt="#,##0.00",
                        title="0.1 DCF（標準版：無槓桿自由現金流 FCFF＋WACC；明細見 DCF 頁）", beta_note="自算產業 β 重新加槓桿")
for _j, _h in enumerate(["項目", "A（樂觀）", "B（基準）", "C（保守）", "說明"], 1): V.cell(SDCF["h"], _j, _h)
T.update(dcfv=SDCF["dcfv"], dcf12=SDCF["v12"], dcftp=SDCF["dcf"])
V.cell(SDCF["stub"], 5, '="資產負債表日 "&bs_date&"；stub 期 FCFF＝Q 頁 4Q26 季度數（DCF 頁下方；三情境相同）"')
assert SDCF["dcf"] <= 24
_PE = lambda c, y: f"PL!{FC[y]}{SUMR[(dict(zip('BCD', SC))[c], 'eps')]}"
EPSR = {}   # v0.2：舊 0.1 節 EPS 列已移除（ROWMAP 相容）
V.cell(26, 1, "0.2 Comparable（市場主流同業前瞻 P/E）").font = BOLD
hdr(V, 27, ["同業與前瞻 P/E", "代號", "前瞻 12 月 P/E", "日期", "來源", "標記"])
PEERS = C["peers"]
for i, (nm, tk, pe, dt, src) in enumerate(PEERS, 28):
    V.cell(i, 1, nm); V.cell(i, 2, tk); c = V.cell(i, 3, pe); c.font = BLUE; c.number_format = "0.0"
    V.cell(i, 4, dt); V.cell(i, 5, src); V.cell(i, 6, "Interested-party")
PEER_LAST = 27 + len(PEERS)
assert PEER_LAST <= 32
tset(33, "pmed", "同業前瞻 P/E 中位數", [f"=MEDIAN(C28:C{PEER_LAST})", None, None], "0.0", C["peer_note"])
tset(34, "feps", "12 個月後的前瞻 EPS（情境加權）", ["=" + "+".join(f"tp_w{dict(zip('BCD', SC))[c]}*(2/12*{_PE(c, 2027)}+10/12*{_PE(c, 2028)})" for c in "BCD"), None, None], "#,##0.00",
     "約 2027-10 時的未來 12 月 EPS ≈ 2／12 × FY2027＋10／12 × FY2028（會計年度＝日曆年）")
tset(35, "comptp", "Comparable 目標價", ["=B33*B34", None, None], "#,##0.00", "同業中位數 × 前瞻 EPS")
V.cell(37, 1, "0.3 結論").font = BOLD
hdr(V, 38, ["項目", "數值", "", "", "說明"])
tset(39, "tpw", "目標價（DCF 與 Comparable 加權）", [f"=tp_w_dcf*B{T['dcftp']}+(1-tp_w_dcf)*B35", None, None], "#,##0.00", "tp_w_dcf × DCF＋（1 − tp_w_dcf）× Comparable（相對倍數調整後，見第 0.5 節）")
tset(40, "comp2", "Comparable 目標價（市場主流同業法）", ["=B35", None, None], "#,##0.00")
tset(41, "cons", "市場共識目標價", ["=cons_tp", None, None], "#,##0.00", C["cons_tp_src"])
tset(42, "upw", "上檔空間：加權目標價", ["=B39/px-1", None, None], "0.0%", f"現價＝In 頁 px（{C['px_date']} 收盤）")
tset(43, "upc", "上檔空間：Comparable 目標價", ["=B40/px-1", None, None], "0.0%")
tset(44, "geps", "與市場共識目標價的差距：EPS 因素", ["=B34/cons_feps-1", None, None], "0.0%", "本模型前瞻 EPS ÷ 共識前瞻 EPS − 1；共識前瞻 EPS 見 In 頁 cons_feps")
tset(45, "gmul", "與市場共識目標價的差距：倍數因素", ["=(cons_tp/cons_feps)/B33-1", None, None], "0.0%")
V["E45"] = '="共識目標價隱含 "&TEXT(cons_tp/cons_feps,"0.0")&"x，同業中位數 "&TEXT(B33,"0.0")&"x"'
tset(46, "band", "價格帶（DCF 保守～樂觀）", [f'=TEXT(D{T["dcf12"]},"#,##0")&"～"&TEXT(B{T["dcf12"]},"#,##0")', None, None], "@", "C 情境價值～A 情境價值（推到 12 個月後）")
tset(47, "pr", "歷史倍數檢查：現價隱含 P/E 的十年分位（A／B／C）", [f"=IF(px/{_PE(c, 2028)}>=MAX({PE10}),1,IF(px/{_PE(c, 2028)}<=MIN({PE10}),0,PERCENTRANK({PE10},px/{_PE(c, 2028)})))" for c in "BCD"], "0%",
     "現價 ÷ 各情境 FY2028 EPS 在 2016–2025 年底 P/E（第 5 節兩來源平均）的分位；高於十年最大值記為 100%、低於最小值記為 0%；只作檢查，不進目標價")
V.cell(48, 1, "目標價方法")
V["B48"] = C["tp_method"]
for r_ in (T["dcftp"], 35, 39):
    V[f"B{r_}"].fill = KEY
V.column_dimensions["A"].width = 58
for c in "BCDEFGH": V.column_dimensions[c].width = 16

# ---------------- Sens ----------------
SE = wb.create_sheet("Sens")
SE["A1"] = ("Sens — 一次動一個輸入（In 頁敏感度低／高值），三情境都在本頁下方逐案重算（FY2026 恆等於錨點，每片現金成本在每個案例內重算），"
            "表內為情境加權後的 DCF、Comparable（相對倍數調整後）與加權目標價；另列 FY2027、FY2028 情境 A 的 EPS")
SE["A2"] = "相對基準＝案例加權目標價 ÷ 基準加權目標價 − 1。DCF 欄為標準版 DCF（FCFF＋WACC；每個案例三情境各在 DCF 頁另寫一個 FCFF 明細區塊）；DCF 參數（WACC、β、ERP、g、RONIC、資本支出、營運資金等）的敏感度見本頁最下方「目標價參數敏感度」表"
hdr(SE, 3, ["案例", "設定", "鍵", "FY2027 EPS（A）", "FY2028 EPS（A）", "FY2028 EPS 相對基準（A）", "DCF 目標價", "Comparable（調整後）目標價", "加權目標價", "相對基準"])
cases = C["sens_cases"]
SENS_ROWS = {}
def put_case(i, name, setting, keys, blocks, pre, dref, w="tp_w_dcf"):
    """blocks：{情境: 推導區塊列號 dict}；pre：區塊所在工作表前綴（PL! 或 空字串＝本頁）"""
    SE.cell(i, 1, name); SE.cell(i, 2, setting); SE.cell(i, 3, keys)
    SE[f"D{i}"] = f"={pre}E{blocks['A']['eps']}"; SE[f"E{i}"] = f"={pre}F{blocks['A']['eps']}"
    SE[f"F{i}"] = f"=E{i}/E$4-1"
    SE[f"G{i}"] = "=" + "+".join(f"tp_w{s}*{dref[s]}" for s in SC)
    fe = "+".join(f"tp_w{s}*(2/12*{pre}E{blocks[s]['eps']}+10/12*{pre}F{blocks[s]['eps']})" for s in SC)
    SE[f"H{i}"] = f"=Val!$B$33*({fe})*rm_factor"
    SE[f"I{i}"] = f"={w}*G{i}+(1-{w})*H{i}"
    SE[f"J{i}"] = f"=I{i}/I$4-1"
    for c, fm in (("D", "#,##0.00"), ("E", "#,##0.00"), ("F", "0.0%"), ("G", "#,##0.00"), ("H", "#,##0.00"), ("I", "#,##0.00"), ("J", "0.0%")):
        SE[f"{c}{i}"].number_format = fm
put_case(4, "基準（三情境加權）", "PL 頁三個推導區塊", "—", PB, "PL!", v12ref(RD_))
# v0.2：舊 tp_r／tp_g／tp_w_dcf 六案改為標準版 DCF 參數案例（WACC、β、ERP、g、RONIC、資本支出、營運資金、權益投資、終值替代法、DCF 權重；閉式公式與 DCF 頁明細同式，tools/fcff_dcf.value_expr）
_KE = lambda beta="beta", erp="erp": f"(rf+{beta}*({erp}+crp))"
_WK = lambda k: f"(w_e*{k}+(1-w_e)*kd_at+wacc_adj)"
FDC = [("WACC −1pt", "WACC −1pt（RONIC 隨 WACC）", "wacc_adj", dict(w="(wacc+wacc_adj_lo)"), None),
       ("WACC ＋1pt", "WACC ＋1pt（RONIC 隨 WACC）", "wacc_adj", dict(w="(wacc+wacc_adj_hi)"), None),
       ("β：低", '="β "&TEXT(beta_lo,"0.00")&"（基準 "&TEXT(beta,"0.00")&"；β_u "&TEXT(beta_u,"0.000")&" − "&TEXT(beta_du,"0.0")&" 重新加槓桿）"', "beta", dict(ke=_KE(beta="beta_lo"), w=_WK(_KE(beta="beta_lo"))), None),
       ("β：高", '="β "&TEXT(beta_hi,"0.00")&"（基準 "&TEXT(beta,"0.00")&"；β_u ＋ "&TEXT(beta_du,"0.0")&" 重新加槓桿）"', "beta", dict(ke=_KE(beta="beta_hi"), w=_WK(_KE(beta="beta_hi"))), None),
       ("ERP：低", '="ERP "&TEXT(erp_lo,"0%")&"（基準 "&TEXT(erp,"0%")&"）"', "erp", dict(ke=_KE(erp="erp_lo"), w=_WK(_KE(erp="erp_lo"))), None),
       ("ERP：高", '="ERP "&TEXT(erp_hi,"0%")&"（基準 "&TEXT(erp,"0%")&"）"', "erp", dict(ke=_KE(erp="erp_hi"), w=_WK(_KE(erp="erp_hi"))), None),
       ("永續成長率 g：低", '="g "&TEXT(g_tv_lo,"0.0%")&"（基準 "&TEXT(g_tv,"0.0%")&"）"', "g_tv", dict(g="g_tv_lo"), None),
       ("永續成長率 g：高", '="g "&TEXT(g_tv_hi,"0.0%")&"（上限 rf）"', "g_tv", dict(g="g_tv_hi"), None),
       ("RONIC＝WACC＋5pt", "永續期新投資報酬率 RONIC＝WACC → WACC＋5pt", "ronic_sp", dict(ronic="(wacc+ronic_sp_hi)"), None),
       ("資本支出 ÷ 營收 −1pt", "每年資本支出比例 −1 個百分點", "fc_cx_d", dict(cxd="fc_cx_d_lo"), None),
       ("資本支出 ÷ 營收 ＋1pt", "每年資本支出比例 ＋1 個百分點", "fc_cx_d", dict(cxd="fc_cx_d_hi"), None),
       ("營運資金 ÷ 營收：低", '="營運資金比例 "&TEXT(fc_wc_lo,"0%")&"（基準 "&TEXT(fc_wc,"0.0%")&"）"', "fc_wc", dict(wc="fc_wc_lo"), None),
       ("營運資金 ÷ 營收：高", '="營運資金比例 "&TEXT(fc_wc_hi,"0%")&"（基準 "&TEXT(fc_wc,"0.0%")&"）"', "fc_wc", dict(wc="fc_wc_hi"), None),
       ("淨現金不含權益投資", "權益投資（帳面 85 億美元）不計入評價", "inv_k", dict(nc="(cash_v+inv_k_lo*inv_v-debt_v-lease_v-mi_v)"), None),
       ("終值改用 FCFF ×（1＋g）", "終值＝FY2032 FCFF ×（1＋g）÷（WACC − g）（替代法）", "—", dict(tv="alt"), None),
       ("DCF 權重 0.3", "Comparable 0.7", "tp_w_dcf", {}, "0.3"), ("DCF 權重 0.7", "Comparable 0.3", "tp_w_dcf", {}, "0.7")]
n_tp = len(FDC) + 1
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
        _kw = dict(tax=f"tax_{side}") if key == "tax" else (dict(nonop=f"nonop_f_{side}") if key == "nonop_f" else {})
        _Rc = ddetail(blocks, "Sens!", f"Sens 案例：{name}：{'低' if side == 'lo' else '高'}（三情境 FCFF 明細）", lab="案例", **_kw)
        put_case(i, f"{name}：{'低' if side == 'lo' else '高'}", setting, key, blocks, "", v12ref(_Rc))
        SENS_ROWS[(key, side)] = i; i += 1
TP_ROWS = {}
for nm, st, k, over, wd in FDC:
    put_case(i, nm, st, k, PB, "PL!", {s_: FD.value_expr(RD_, s_, **over) for s_ in SC}, **({"w": wd} if wd else {})); TP_ROWS[nm] = i; i += 1
SE.cell(i, 1, "檢查：閉式公式的基準 DCF 目標價 − Val 第 0.1 節（應為 0；閉式與 DCF 頁明細同式）")
SE.cell(i, 7, "=" + "+".join(f"tp_w{s_}*{FD.value_expr(RD_, s_)}" for s_ in SC) + f"-Val!B{T['dcftp']}").number_format = "0.000000"; i += 1
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
RC.cell(rr_, 1, "伺服器 CPU、客戶端 PC、代工利益率與外部代工營收的行業層（Siliconomics）尚未建立：上線後以本模型情境對照行業路徑"); RC.cell(rr_, 10, "預留列：本版無數值"); rr_ += 1
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
V.cell(GRADE_ROW, 1, "上檔空間級距（六級，供分組）")
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

order = ["README", "In", "Q", "Units", "PL", "Val", "DCF", "Recon", "Sens", "Sources"]
wb._sheets = [wb[n] for n in order]
import relative_multiple as RMX  # noqa: E402
# Intel 不在 tools/target_summary 的 universe.csv／peers.csv（tools/ 不改）：在行程內補上本公司代號與同業（Bloomberg 前瞻倍數滿 7 年者：台積電、AMD、聯電 2303.TW；README 註明）
_U0, _P0 = RMX._universe, RMX._peers
RMX._universe = lambda: {**_U0(), "intc": {"folder": "intc", "公司": "英特爾", "own_ticker": "INTC"}}
RMX._peers = lambda folder: ([("台積電（TSMC）", "TSM"), ("超微（AMD）", "AMD"), ("聯電（UMC）", "2303.TW")] if folder == "intc" else _P0(folder))
RMX.apply(wb, C["folder"])
RD.cell(RD.max_row, 1, "相對倍數調整（tools/relative_multiple.py；Val 第 0.5 節）已含在加權目標價：Comparable（相對倍數調整後）＝同業前瞻倍數中位數 × 本公司 10 年相對同業的倍數比（Bloomberg 前瞻倍數；同業取台積電、AMD、聯電）。「Comparable 目標價（市場主流同業法）」維持不調整。tools/target_summary 未含 Intel，本檔在建置時於行程內補上（未改 tools/）。")
sys.path.insert(0, os.path.join(ROOT, "tools", "in_layout")); import apply_in_layout  # noqa: E402
apply_in_layout.apply(wb, "v0.1.0", "2026-11-17")
for _r in range(4, IN.max_row + 1):   # v0.2.2（第三輪修正第 15、23 點）：表頭統一順利／平平／不如預期；收斂法列名中性
    if str(IN.cell(_r, 1).value or "").startswith("情境驅動對照"):
        for _j, _h in zip((4, 5, 6), ("情境 A（樂觀）", "情境 B（基準）", "情境 C（保守）")): IN.cell(_r + 1, _j).value = _h
    if str(IN.cell(_r, 1).value or "").startswith("scn_fade"): IN.cell(_r, 2).value = re.sub(r"（[^）]*）$", "", str(IN.cell(_r, 2).value))
# ---------------- v0.13（2026-11-04）：Sens「Wattonomics 錨定」案例（三層校準原則；決策 S37）。In 頁最後三列、Sens 頁最下方小表與案例區塊；既有格不動 ----------------
WTA_NOTE = "Wattonomics 錨定：以情境 B 為底，產業量 CY2027 起 × wta（Siliconomics SIF_WTAnchor_Base＝R4 比值）；FY2026 維持錨點。三層校準原則（Andy 2026-11-04）；個股基準暫不改。"
WTA_MF = {2027: "wta_27", 2028: "wta_28", 2029: "wta_29"}
def wta_in_rows():
    r = IN.max_row + 2
    IN.cell(r, 1, f"Z. Wattonomics 錨定係數（敏感度用；Siliconomics {SIL.VERSION}）").font = BOLD; r += 1
    def _inp(key, item, unit, lo, base, hi, tag, s1, s2):
        nonlocal r
        for j, v in enumerate((key, item, unit, lo, base, hi, tag, s1, s2), 1):
            if v is not None: IN.cell(r, j, v)
        IN.cell(r, 5).number_format = "0.000"
        wb.defined_names[key] = DefinedName(key, attr_text=f"In!$E${r}")
        r += 1
    SIL.wta_rows(_inp)
def wta_sens():
    t0 = SE.max_row + 3; brow_ = t0 + 9; WB = {}
    for s in SC:
        WB[s] = engine(SE, brow_, f"案例區塊：Wattonomics 錨定（情境 {s}；FY2027 起 AI 相關營收 × wta_27／wta_28／wta_29）", scn=s, fm_self=True, wm=WTA_MF)
        brow_ = WB[s]["end"] + 2
    SE.cell(t0, 1, "Wattonomics 錨定（三層校準原則）").font = BOLD
    hdr(SE, t0 + 1, [SE.cell(3, c).value for c in range(1, 11)] + ["FY2027 EPS（B）", "FY2028 EPS（B）"])
    put_case(t0 + 2, "基準（三情境加權；同第 4 列）", "PL 頁三個推導區塊", "—", PB, "PL!", v12ref(RD_))
    _Rw = ddetail(WB, "Sens!", "Sens 案例：Wattonomics 錨定（三情境 FCFF 明細）", lab="錨定")
    put_case(t0 + 3, "Wattonomics 錨定", "FY2027 起 DCAI 中 AI 相關營收 × wta（三情境同乘；權重不變）", "wta_27／wta_28／wta_29", WB, "", v12ref(_Rw))
    for rr, bl, pre in ((t0 + 2, PB, "PL!"), (t0 + 3, WB, "")):
        SE[f"K{rr}"] = f"={pre}E{bl['B']['eps']}"; SE[f"L{rr}"] = f"={pre}F{bl['B']['eps']}"
        for c in "KL": SE[f"{c}{rr}"].number_format = "#,##0.00"
    SE.cell(t0 + 4, 1, "差異%：Wattonomics 錨定 ÷ 基準 − 1")
    for c in "DEGHIKL":
        SE[f"{c}{t0 + 4}"] = f'=IFERROR({c}{t0 + 3}/{c}{t0 + 2}-1,"")'; SE[f"{c}{t0 + 4}"].number_format = "0.0%"
    SE.cell(t0 + 5, 1, WTA_NOTE)
    SE.cell(t0 + 6, 1, "本模型 AI 量驅動只有 DCAI 中 AI 相關營收（FY2026 約占營收 15%，Assumed 拆分）；CCPG、Foundry 與其他不乘。FY(t)≈CY(t)；FY2030–FY2032 沿用收斂法。目標價為三情境皆 × wta 後依情境權重加權。")
    return WB
wta_in_rows(); WTA_R = wta_sens()   # v0.13：放在 In 版面套用之後，確保在 In 頁最後
_tvs = (TP_ROWS['WACC −1pt'], None)   # README 用：DCF 參數案例首列
# ---------------- v0.1.0（2026-11-17）：展示層（README）、三層協調（Coord）、版面 v1.3 後處理 ----------------
_HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, _HERE)
import show_intc  # noqa: E402  README 展示層與重點觀察（公式帶數字）
show_intc.apply(wb, dict(T=T, UR=UR, SUMR=SUMR, PB=PB, SENS_ROWS=SENS_ROWS, TP_ROWS=TP_ROWS, SDCF=SDCF, OBS_ROW0=OBS_ROW0, REC_ROW=REC_ROW, O=O, FC=FC, SIL=SIL))
sys.path.insert(0, os.path.join(ROOT, "tools", "coord"))
import coord as COORD  # noqa: E402
_CR = COORD.add_sheet(wb, _HERE, SIL)
if _CR:
    import v010_post  # noqa: E402
    print(v010_post.apply(wb, _HERE, SIL, _CR, dict(PB=PB, C=C, SENS_ROWS=SENS_ROWS, TP_ROWS=TP_ROWS)))
wb.save(sys.argv[1])
if os.environ.get("ROWMAP"):   # 除錯用：輸出列號對照；CI 不設此變數
    json.dump({"PB": PB, "SUMR": {f"{k[0]}|{k[1]}": v for k, v in SUMR.items()}, "UR": UR, "CHK": CHK, "VAL": VAL, "INROW": INROW,
               "SENS_LAST": SENS_LAST, "SENS_NOTE": SENS_NOTE, "STOP": STOP, "RECON_LAST": RECON_LAST, "RULE_ROWS": rule_rows, "RB0": RB0, "WMT_NOTE": WMT_NOTE,
               "UN_NOTE": UN_NOTE, "T": T, "T10": T10, "PEER_LAST": PEER_LAST, "EPSR": EPSR, "O": O, "CY0": CY0, "R30": R30, "W0": W0,
               "GRADE_ROW": GRADE_ROW, "REC_ROW": REC_ROW, "OBS_ROW0": OBS_ROW0, "RD_LAST": RD.max_row,
               "SENS_ROWS": {f"{k[0]}|{k[1]}": v for k, v in SENS_ROWS.items()}, "TP_ROWS": TP_ROWS, "QCH_ROWS": QCH_ROWS},
              open(os.environ["ROWMAP"], "w"), ensure_ascii=False, indent=1)
