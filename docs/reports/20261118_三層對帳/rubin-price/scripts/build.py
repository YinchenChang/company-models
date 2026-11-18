# -*- coding: utf-8 -*-
"""建 20261118_NVIDIA顆數自建_2027.xlsx（全公式；輸入頁藍字）。
用法：python3 build.py <輸出路徑>
NVIDIA FY2027 第三季實績未使用（只用 Q3 指引，屬 2026-08-26 已公開的 Interested-party）。"""
import sys
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

OUT = sys.argv[1]
wb = Workbook()
BLUE = Font(color="0000FF")
BOLD = Font(bold=True)
HDR = PatternFill("solid", fgColor="DDEBF7")
WRAP = Alignment(wrap_text=True, vertical="top")

# ---------------- 輸入 ----------------
# (名稱, 項目, 單位, 低, 基準, 高, 等級, 來源, URL, 日期, 說明)
U_NVQ2 = "https://s201.q4cdn.com/141608511/files/doc_financials/2027/Q227/Q2FY27-CFO-Commentary.pdf"
U_NVQ2T = "https://s201.q4cdn.com/141608511/files/content_files/TRANSCRIPT_-NVIDIA-Corp-NVDA-US-Q2-2027-Earnings-Call-26-August-2026-5_00-PM-ET.pdf"
U_NVQ1 = "https://s201.q4cdn.com/141608511/files/doc_financials/2027/Q127/Q1FY27-CFO-Commentary.pdf"
U_NVQ2FY26 = "https://s201.q4cdn.com/141608511/files/doc_financials/2026/Q226/Q2FY26-CFO-Commentary.pdf"
U_NVQ3FY26T = "https://www.fool.com/earnings/call-transcripts/2025/11/19/nvidia-nvda-q3-2026-earnings-call-transcript/"
U_GTCDC = "https://www.nextplatform.com/ai/2025/10/29/as-rackscale-gpu-systems-boom-big-green-is-emphatically-red-white-and-blue/1640768"
U_GTC25 = "https://media.iprsoftware.com/219/files/20252/nvidia-gtc-2025-keynote-deck.pdf"
U_1T = "https://www.letsdatascience.com/news/nvidia-projects-1-trillion-blackwell-vera-rubin-demand-b8d6cf74"
U_10K = "https://www.marketscreener.com/news/nvidia-annual-report-for-fiscal-year-ending-january-25-2026-form-10-k-ce7e5cd8d18af32d"
U_NV8K = "https://nvidianews.nvidia.com/news/nvidia-announces-financial-results-for-second-quarter-fiscal-2027"
U_TSMC4Q25 = "https://www.itiger.com/news/1163623376"
U_TSMC3Q26 = "https://www.webull.com/news/15751044230423552"
U_DIGI = "https://apps.digitimes.com/news/a20260202PD213/tsmc-revenue-growth-2026.html"
U_TOMS = "https://www.tomshardware.com/tech-industry/big-tech/microsoft-attributed-25-billion-of-its-record-ai-budget-to-memory-chip-costs"
U_UAO = "https://www.universalassetowners.com/newsroom/hyperscaler-q3-capex-debt-equity-leases/"
U_DCK = "https://datacenterknowledge.com/data-center-construction/hyperscalers-say-ai-race-has-entered-a-new-phase"
U_FOX = "https://www.dzrh.com.ph/post/foxconn-can-make-1000-ai-racks-a-week-to-increase-next-year-chairman-says"
U_HBM = "repo: docs/reports/20261118_三層對帳/hbm/20261118_HBM獨立供給_查證.md"
U_TK = "repo: Tokenomics v22.0 Interface（經 /home/claude/out/sens In 頁與 unit-basis 報告第三節）"
U_UB = "repo: docs/reports/20261118_三層對帳/unit-basis/20261118_顆數來源口徑查核.md"
U_BLOG = "https://developer.nvidia.com/blog/nvidia-vera-rubin-pod-seven-chips-five-rack-scale-systems-one-ai-supercomputer"
U_NVQ2FY26T = "https://s201.q4cdn.com/141608511/files/doc_financials/2026/q2/NVDA-Q2-2026-Earnings-Call-27-August-2025-5_00-PM-ET.pdf"
U_DSX = "https://developer.nvidia.com/blog/maximizing-ai-factory-performance-per-watt-with-nvidia-dsx-maxlps/"
U_NSC = "https://www.sec.gov/Archives/edgar/data/0002110365/000119312526458722/ck0002110365-20261112.htm"
U_460 = "https://theenergymag.com/news/market-news/anthropic-secures-460-mw-at-nscale-campus-in-45-billion-deal-report"
U_IREN = "https://seekingalpha.com/pr/20289163-iren-secures-9_7bn-ai-cloud-contract-with-microsoft"
U_JH45 = "https://www.tweaktown.com/news/107855/openai-and-nvidias-new-ai-project-requires-4-5-million-gpus-in-one-project-alone/index.html"
U_DELL1 = "https://www.businesswire.com/news/home/20260528449392/en/Dell-Technologies-Delivers-First-Quarter-Fiscal-2027-Financial-Results"
U_DELL2 = "https://www.01net.it/dell-technologies-delivers-second-quarter-fiscal-2027-financial-results/"
U_SMCI = "https://pulse2.com/supermicro-new-orders-top-60-billion-as-fiscal-2027-revenue-targets-up-to-72-billion/"
U_SMCIQ1 = "https://www.sec.gov/Archives/edgar/data/0001375365/000137536526000025/smcifirstquarterfy2027busi.htm"
U_HPE = "https://www.nextplatform.com/compute/2026/09/08/hpe-sees-no-customer-pushback-on-rising-datacenter-equipment-costs/5295102"
U_PREV = "/home/claude/out/nvunits/20261118_NVIDIA顆數自建_2027.md"
U_NEB = "https://www.sec.gov/Archives/edgar/data/0001513845/000110465926126008/tm2629571d1_ex99-2.htm"
U_BG2 = "https://www.tikr.com/blog/jensen-huang-two-gigawatts-to-us-is-like-80-billion-43-new-gigawatts-coming-in-2027"
U_BERN = "https://wccftech.com/bernstein-warns-nvidias-vera-rubin-racks-will-hit-9-1-million-as-hbm4-prices-triple-to-53-per-gigabyte/amp/"

INPUTS = [
    ("sec", "A. NVIDIA 公司實績與指引（十億美元；FY2027＝2026-02～2027-01）"),
    ("rev_q1", "FY27 Q1 營收", "US$B", None, 81.61, None, "Verified", "NVIDIA 8-K（個股模型 Q 頁 G4）", U_NV8K.replace("second", "first"), "2026-05", ""),
    ("rev_q2", "FY27 Q2 營收", "US$B", None, 96.22, None, "Verified", "NVIDIA 8-K（個股模型 Q 頁 H4）", U_NV8K, "2026-08-26", ""),
    ("rev_q3g", "FY27 Q3 營收指引（2026-08-26 公布；不是實績）", "US$B", 105.84, 108.0, 110.16, "Interested-party", "Q2 FY27 CFO commentary：108.0 ±2%", U_NVQ2, "2026-08-26", "低／高＝±2%"),
    ("g_q4", "FY27 Q4 營收 ÷ Q3 指引 − 1", "比", 0.0, 0.075, 0.15, "Assumed", "無公司指引（Q4 指引隨 Q3 財報公布，依規則不用）", "", "", "持平／+7.5%／+15%"),
    ("g_fy28", "FY2028 營收成長率（公司說法）", "比", 0.65, 0.70, 0.75, "Interested-party", "Q2 FY27 法說 CFO：「approximately 70%」", U_NVQ2T, "2026-08-26", "低／高＝『約』的寬度 ±5 個百分點（Assumed）；公司稱供給受限至 FY2028 底"),
    ("dc_sh28", "FY2028 資料中心占營收", "比", 0.92, 0.935, 0.95, "Assumed", "Q2 FY27 實績 89.0÷96.2＝92.5%（Derived）", U_NVQ2, "2026-08-26", ""),
    ("cal27", "曆年 2027 ÷ FY2028（月份錯開一個月）", "比", 0.96, 0.98, 1.00, "Assumed", "FY2028＝2027-02～2028-01；成長中月份錯開使曆年略低", "", "", ""),
    ("dc_q4fy25", "FY25 Q4 資料中心營收（2024-11～2025-01）", "US$B", None, 35.6, None, "Verified", "NVIDIA 8-K（FY2025 Q4）", "https://nvidianews.nvidia.com/news/nvidia-announces-financial-results-for-fourth-quarter-and-fiscal-2025", "2025-02-26", ""),
    ("dc_q1fy26", "FY26 Q1 資料中心營收", "US$B", None, 39.1, None, "Verified", "NVIDIA 個股模型 Q 頁（8-K）", "", "2025-05", ""),
    ("dc_q2fy26", "FY26 Q2 資料中心營收", "US$B", None, 41.096, None, "Verified", "Q2 FY27 CFO commentary 比較欄", U_NVQ2, "2026-08-26", ""),
    ("dc_q3fy26", "FY26 Q3 資料中心營收", "US$B", None, 51.2, None, "Verified", "NVIDIA 個股模型 Q 頁（8-K）", "", "2025-11-19", ""),
    ("dc_q4fy26", "FY26 Q4 資料中心營收", "US$B", None, 62.3, None, "Verified", "NVIDIA 個股模型 Q 頁（8-K）", "", "2026-02-25", ""),
    ("dc_q1", "FY27 Q1 資料中心營收", "US$B", None, 75.246, None, "Verified", "Q2 FY27 CFO commentary", U_NVQ2, "2026-08-26", ""),
    ("dc_q2", "FY27 Q2 資料中心營收", "US$B", None, 89.023, None, "Verified", "Q2 FY27 CFO commentary", U_NVQ2, "2026-08-26", ""),
    ("dc_sh_q3", "FY27 Q3 資料中心占營收（套在指引上）", "比", 0.92, 0.925, 0.93, "Assumed", "Q2 FY27 實績 92.5%", U_NVQ2, "", ""),
    ("hs_q2fy26", "FY26 Q2 Hyperscale 資料中心營收（重編後）", "US$B", None, 24.168, None, "Verified", "Q2 FY27 CFO commentary 客戶別表", U_NVQ2, "2026-08-26", ""),
    ("hs_q1", "FY27 Q1 Hyperscale 資料中心營收", "US$B", None, 43.050, None, "Verified", "同上", U_NVQ2, "2026-08-26", ""),
    ("hs_q2", "FY27 Q2 Hyperscale 資料中心營收", "US$B", None, 48.710, None, "Verified", "同上", U_NVQ2, "2026-08-26", ""),
    ("commit28", "供應與產能承諾：FY2028 到期部分（截至 2026-07-26）", "US$B", None, 87.0, None, "Verified", "Q2 FY27 CFO commentary；總額 279，FY27 剩餘 92、FY28 87、FY29 88；增加主因記憶體", U_NVQ2, "2026-08-26", "10-Q 原文未取得（sec.gov 被擋），用公司 CFO commentary"),
    ("gm28", "FY2028 毛利率（公司說法）", "比", 0.73, 0.725, 0.72, "Interested-party", "Q2 FY27 法說：FY2028 落在 72%–73%", U_NVQ2T, "2026-08-26", "低欄放 73%（使下限最低）"),

    ("sec", "B. NVIDIA 自己的能見度與單價說法（Interested-party：動機偏高）"),
    ("br_500", "Blackwell＋Rubin 營收能見度（2025 年初～2026 年底）", "US$B", None, 500.0, None, "Interested-party", "Q3 FY26 法說 CFO：「half a trillion dollars in Blackwell and Rubin revenue」（二手逐字稿）", U_NVQ3FY26T, "2025-11-19", ""),
    ("br_shipped", "上述能見度中已出貨（至 FY26 Q3，2025-10）", "US$B", None, 150.0, None, "Interested-party", "Q3 FY26 法說：分析師說 $150B，CFO 答「Yes. That's correct.」（二手）", U_NVQ3FY26T, "2025-11-19", ""),
    ("bw_dies_shipped", "已出貨 Blackwell（晶粒，GTC DC）", "百萬晶粒", None, 6.0, None, "Interested-party", "黃仁勳 GTC DC：「shipped 6 million of the Blackwells」；GTC 2025 簡報：2-GPUs per Chip", U_GTCDC, "2025-10-28", "÷2＝封裝；期間『first several quarters』未精確"),
    ("br_1t", "Blackwell＋Rubin 訂單（至 2027 年）", "US$B", None, 1000.0, None, "Interested-party（二手）", "GTC 2026；原文未取得，二手轉述『through 2027』；Q2 FY27 法說分析師稱『$1 trillion over the three years』", U_1T, "2026-03-16", "期間按 2025–2027 三年讀"),
    ("k_overrun", "實現 ÷ 公司能見度說法（由 $500B 校驗推得）", "比", 1.00, 1.03, 1.07, "Derived／Assumed", "見 A2 頁 2026 校驗列：2025–2026 實現（實績＋Q3 指引＋Q4 假設）÷ $500B＝0.99／1.03／1.07", "", "", "低＝說法即上限"),
    ("usd_gw_bw", "NVIDIA 每 GW 營收機會：Blackwell", "US$B/GW", None, 25.0, None, "Interested-party", "Q2 FY27 法說 CFO：Hopper 約 18、Blackwell 25、Vera Rubin 40", U_NVQ2T, "2026-08-26", "動機：放大總潛在市場"),
    ("usd_gw_rb", "NVIDIA 每 GW 營收機會：Vera Rubin", "US$B/GW", None, 40.0, None, "Interested-party", "Q2 FY27 法說原文（公司 IR PDF）Kress：「to $40 billion with Vera Rubin」；範圍含 Vera CPU、Rubin GPU、NVLink、InfiniBand/Ethernet、Groq LPU；GW 未定義 IT／設施", U_NVQ2T, "2026-08-26", "動機：放大總潛在市場；2026-09-17 蘇格蘭、2026-10 BG2 重複同數字"),

    ("sec", "H. 本輪新增：單價證據（rubinprice，2026-11-18）。等級見 G 欄；NVIDIA FY27 Q3 實績與 10-Q 未使用"),
    ("usd_gw_bw25", "NVIDIA 2025-08 說法：每 GW AI 工廠中 NVIDIA 內容", "US$B/GW", None, 35.0, None, "Interested-party", "Q2 FY26 法說（公司 IR PDF）Huang：「we represent about $35 billion」of 約 $50B／GW；世代未指明", U_NVQ2FY26T, "2025-08-27", "與 2026-08『Blackwell 25』不一致：只作對照，不入區間"),
    ("usd_gw_tot", "NVIDIA 2026-08 說法：每 GW 資料中心客戶總投資", "US$B/GW", None, 60.0, None, "Interested-party", "Q2 FY27 法說 Huang：「to now $60 billion today」", U_NVQ2T, "2026-08-26", "NVIDIA 內容 40 ÷ 60＝67%"),
    ("gpu_share40", "40 B$/GW 中屬 Rubin GPU 系統（含 NVLink、網通）的比例", "比", 0.85, 0.92, 1.00, "Assumed", "Kress 稱 Vera Rubin 範圍含 Groq LPU；CFO：FY2028 CPU 營收超過兩倍（含獨立 Vera CPU）", U_NVQ2T, "2026-08-26", "低欄＝單價低"),
    ("dens_ratio", "每 GW 封裝數：GB300 ÷ VR200（同一 GW 口徑）", "倍", 1.40, 1.55, 1.67, "Analogy／Derived", "Tokenomics 487,008÷291,744＝1.669（高欄）；GB300 機櫃約 136 kW（NVIDIA 技術部落格圖說）對 VR200 約 190–227 kW", U_TK, "2026-11", "低欄＝單價低（Rubin 機櫃功率較低時）"),
    ("nv_dens_grid", "NVIDIA 示意：每 100 MW 電網輸入可放 Rubin GPU（MaxLPS 開啟）", "千顆/100MW", None, 40.0, None, "Interested-party", "NVIDIA 技術部落格：「operators can fit 40K Rubin GPUs」；並稱「up to 40% more」", U_DSX, "2026-08-24", "示意圖；100 MW 為電網輸入：設施 −20、機架損耗 −10、運作 −10 → AI 負載 60 MW"),
    ("lps_gain", "MaxLPS 增加的 GPU 容量（還原無 MaxLPS 的基準）", "比", 1.25, 1.40, 1.40, "Interested-party／Assumed", "部落格：「up to 40% more」；低欄＝實際增益只有 25%（基準密度較高）", U_DSX, "2026-08-24", ""),
    ("it_of_grid", "NVIDIA 示意中 IT 電力占電網輸入", "比", None, 0.80, None, "Interested-party", "同上圖：設施負擔 20 MW／100 MW", U_DSX, "2026-08-24", ""),
    ("nsc_capex", "Nscale Monarch（安索普）GPU 與其他運算設備資本支出", "US$B", 21.0, 21.5, 22.0, "Verified（公司公開說明書，估計值）", "Nscale S-1/A：GPUs and other compute equipment「approximately $21 billion to $22 billion」；部署 Vera Rubin NVL72 GPUs、四批", U_NSC, "2026-11-12", "公司估計，非已發生"),
    ("nsc_mw", "Monarch 安索普案容量", "MW", 460, 460, 460, "Interested-party（二手：Bloomberg 經轉載）", "約 460 MW、六年 $45B；未寫 IT 或總電力。S-1：Building 1 DC 與電力 $10–10.5B、全園區 2 GW 總發電＝1.37 GW IT", U_460, "2026-08-26", "本報告取 IT（DC 與電力每 IT MW 約 $22M 偏向 IT；但 $45B÷6 年÷460 MW＝每 MW 年 $16M 低於同業，偏向總電力：證據分歧）；若為總電力，單價高 46%（價格頁敏感度）"),
    ("nsc_dens", "Monarch 每 IT GW Rubin 封裝數（換算 Nscale 金額用）", "封裝/GW", 357000, 320000, 291744, "Analogy", "低欄＝NVIDIA 示意 IT 口徑 28.6K/80 MW；高欄＝Tokenomics VR200", U_DSX, "2026-08-24", "低欄放高密度（單價低）"),
    ("nsc_nvsh", "NVIDIA 占運算設備資本支出", "比", 0.75, 0.80, 0.85, "Assumed", "其餘：Dell／ODM 組裝與毛利、儲存、第三方網通、軟體授權（IREN 例：$5.8B 含部署、伺服器、InfiniBand、線材、軟體授權）", U_IREN, "2025-11-03", ""),
    ("iren_usd", "IREN 向 Dell 採購 GB300 GPU 與附屬設備", "US$B", None, 5.8, None, "Verified（公司新聞稿）", "「purchase the GPUs and ancillary equipment for approximately $5.8 billion」", U_IREN, "2025-11-03", "含部署、伺服器、InfiniBand、線材、軟體授權"),
    ("iren_mw", "IREN Childress Horizon 1–4（Microsoft 案）", "MW IT", None, 200, None, "Verified（公司新聞稿）", "「200MW of critical IT load」", U_IREN, "2025-11-03", ""),
    ("gpu25_jh", "Huang：2025 年全年出貨 GPU（≈ 10 GW 專案）", "百萬", 4.0, 4.5, 5.0, "Interested-party（二手）", "CNBC 訪談：「roughly, 4 million or 5 million GPUs」≈「what we shipped all year this year」", U_JH45, "2025-09-22", "按封裝讀（若按晶粒，2025 全年應 >7 百萬）"),
    ("sinesgpu", "Nscale SINES（Microsoft）Vera Rubin GPU 數", "千顆", None, 66.0, None, "Verified（公開說明書）", "「over 66,000 NVIDIA Vera Rubin NVL72 GPUs starting in late 2027」；園區 200 MW", U_NSC, "2026-11-12", "只作密度檢查"),
    ("sines_mw", "SINES 園區容量", "MW", None, 200, None, "Verified（公開說明書）", "「a 200 MW owned hyperscale campus」；口徑未寫", U_NSC, "2026-11-12", ""),
    ("dell_fy27", "Dell FY27（2026-02～2027-01）AI 伺服器營收指引", "US$B", None, 74.0, None, "Interested-party（公司指引）", "Q2 FY27 新聞稿：上調至 $74.0B（原 $60B）", U_DELL2, "2026-09-01", ""),
    ("dell_h1", "Dell FY27 上半年 AI 伺服器營收（Q1 16.13＋Q2 16.4）", "US$B", None, 32.53, None, "Verified", "Q1 FY27 新聞稿 $16,132M；Q2 $16.4B", U_DELL1, "2026-05-28／09-01", ""),
    ("dell_bl", "Dell AI 積壓訂單（2026-07-31）", "US$B", None, 95.0, None, "Verified（公司揭露）", "「record $95 billion backlog」", U_DELL2, "2026-09-01", "含非 NVIDIA（AMD）"),
    ("smci_fy27", "SMCI FY27（2026-07～2027-06）營收指引", "US$B", 65.0, 68.5, 72.0, "Interested-party（二手）", "FY27 營收 $65–72B", U_SMCI, "2026-08-12", ""),
    ("smci_q1", "SMCI FY27 Q1（7–9 月）營收初估", "US$B", 18.5, 19.0, 19.5, "Interested-party（8-K 初估）", "$18.5–19.5B", U_SMCIQ1, "2026-10-20", ""),
    ("smci_ai", "SMCI 營收中 AI GPU 平台占比", "比", 0.75, 0.85, 0.90, "Assumed", "公司過去稱 AI GPU 平台占營收逾七成；本輪未取得 FY27 數字", "", "", ""),
    ("hpe_ai", "HPE AI 系統年化營收（Q3 FY26 約 1.6 ×4）", "US$B", None, 6.4, None, "Verified（二手）", "Q3 FY26 AI 營收約 $1.6B、積壓 $6.8B", U_HPE, "2026-09-08", ""),
    ("oem_nv", "OEM AI 伺服器營收中 NVIDIA 內容比例", "比", 0.65, 0.75, 0.85, "Assumed", "含 AMD 機櫃、x86 CPU、儲存、組裝與 OEM 毛利", "", "", "低欄＝NVIDIA 內容少"),
    ("smci_q2r", "SMCI FY27 Q2 ÷ Q1", "比", 0.9, 1.0, 1.1, "Assumed", "無公司 Q2 指引（本輪未取得）", "", "", ""),
    ("wat_gw_nv", "Wattonomics 每 GW NVIDIA 內容（chat 端提供）", "US$B/GW", None, 26.1, None, "Derived（他層）", "工作單所述『約 261 億』；本輪未追溯推導（若＝37.59×0.694，即假設 NVIDIA 占 IT 資本支出約 69%）", "", "2026-11-18", "對照標的，不是輸入"),
    ("wat_gw27", "Wattonomics 2027 NVIDIA 實體 GW（Rubin 14.41＋Ultra 0.93）", "GW", None, 15.34, None, "Derived（他層）", "unit-basis 報告", U_UB, "2026-11-18", "對照"),
    ("prev_lo", "上一輪 2027 重疊區：低", "百萬", None, 4.00, None, "Derived（上一輪）", "nvunits 報告", U_PREV, "2026-11-18", "對照"),
    ("prev_med", "上一輪 2027 最可能值", "百萬", None, 6.00, None, "Derived（上一輪）", "nvunits 報告", U_PREV, "2026-11-18", "對照"),
    ("prev_hi", "上一輪 2027 重疊區：高", "百萬", None, 8.81, None, "Derived（上一輪）", "nvunits 報告", U_PREV, "2026-11-18", "對照"),
    ("prev_rb_lo", "上一輪 Rubin 單價：低", "千美元", None, 90.0, None, "Derived（上一輪）", "nvunits 報告", U_PREV, "2026-11-18", "對照"),
    ("prev_rb_b", "上一輪 Rubin 單價：基準", "千美元", None, 120.0, None, "Derived（上一輪）", "nvunits 報告", U_PREV, "2026-11-18", "對照"),
    ("prev_rb_hi", "上一輪 Rubin 單價：高", "千美元", None, 137.0, None, "Derived（上一輪）", "nvunits 報告", U_PREV, "2026-11-18", "對照"),

    ("sec", "C. 每封裝單價（千美元，NVIDIA 全部內容含網通；Derived／Assumed）"),
    ("p_b25", "Blackwell 每封裝 NVIDIA 營收（2025 錨點）", "千美元", 43.0, 50.0, 55.0, "Derived（Interested-party 輸入）", "基準＝$150B ÷（600 萬晶粒÷2）；低＝晶粒數期間早於 $150B 截止（高估單價）；高＝Assumed", U_GTCDC, "2025-11", "交叉：25 B$/GW ÷ GB300 每 GW 487,008 封裝＝51.3"),
    ("p_b26", "Blackwell（Ultra）每封裝 NVIDIA 營收 2026", "千美元", 45.0, 50.0, 56.0, "Assumed", "以 2025 錨點為中心；記憶體轉嫁上調", "", "", ""),
    ("p_b27", "Blackwell（Ultra）每封裝 NVIDIA 營收 2027", "千美元", 45.0, 52.0, 60.0, "Assumed", "同上；CFO 稱漲價幅度超出預期", U_NVQ2T, "2026-08-26", ""),
    ("gpg_gb300", "每 GW 封裝數 GB300 NVL72（IT 關鍵電力）", "封裝/GW", None, 487008, None, "Derived", "Tokenomics v22.0 IF_GPUsPerGW", U_TK, "2026-11", ""),
    ("gpg_vr", "每 GW 封裝數 VR200 NVL72（227 kW/架）", "封裝/GW", None, 291744, None, "Derived", "Tokenomics v22.0 IF_GPUsPerGW", U_TK, "2026-11", ""),
    ("gpg_ru", "每 GW 封裝數 Rubin Ultra", "封裝/GW", None, 194760, None, "Derived", "Tokenomics v22.0 IF_GPUsPerGW", U_TK, "2026-11", ""),
    ("kw_vr", "VR200 每架功率（決定每 GW 封裝數）", "kW/架", 190, 227, 227, "Interested-party／Assumed", "Tokenomics 基準 227；低＝較低功率設定 190（Assumed）", U_TK, "", "低功率→每 GW 封裝多→每封裝單價低"),
    ("cit_vr", "VR200 每 GW IT 資本支出（買方，含網通與機架外）", "US$B/GW", None, 37.586, None, "Derived（機架價 Interested-party）", "Tokenomics v22.0 Interface", U_TK, "2026-11", ""),
    ("nv_take", "NVIDIA 占買方 IT 資本支出比例（Rubin 世代）", "比", 0.70, 0.80, 0.85, "Assumed", "其餘為 ODM、儲存、第三方網通、機架外", "", "", ""),
    ("s_r26", "2026 年 Rubin 占 NVIDIA 封裝數", "比", 0.03, 0.06, 0.10, "Assumed", "CFO：Q3 Rubin 約占資料中心營收 20%、8 月開始量產出貨", U_NVQ2T, "2026-08-26", ""),
    ("s_r27", "2027 年 Rubin（含 Ultra）占 NVIDIA 封裝數", "比", 0.65, 0.75, 0.85, "Assumed", "CFO：Rubin 為公司史上最快爬坡", U_NVQ2T, "2026-08-26", "低欄＝0.65（單價低→顆數多的組合在高欄另取）"),
    ("ru27", "2027 年 Rubin Ultra 封裝數", "百萬", 0.0, 0.1, 0.3, "Assumed", "路線圖 2027 下半年；無量的公司揭露", "", "", ""),
    ("br26", "2026 年資料中心中 Blackwell＋Rubin 系統占比（含網通）", "比", 0.94, 0.96, 0.98, "Assumed", "其餘：Hopper、中國 H200（<1%）、獨立 CPU、軟體服務", U_NVQ2, "", ""),
    ("br27", "2027 年資料中心中 Blackwell＋Rubin 系統占比（含網通）", "比", 0.88, 0.92, 0.95, "Assumed", "CFO：FY2028 CPU 營收預計超過兩倍；指引不含中國運算", U_NVQ2T, "2026-08-26", ""),
    ("hop25_rev", "2025 年 Hopper＋H20 營收", "US$B", 10.0, 14.0, 18.0, "Assumed（部分 Verified）", "H20 FY26 Q1 約 4.6（公司）；其餘 Hopper 未揭露", U_NVQ2FY26, "", ""),
    ("hop25_u", "2025 年 Hopper＋H20 封裝數", "百萬", 0.5, 0.7, 0.9, "Assumed", "H20 約 1.2 萬美元、H100/H200 約 2.5 萬美元換算", "", "", ""),

    ("sec", "D. 下游：四大買方資本支出（十億美元；公司說法，二手轉述）"),
    ("capex25", "四大（Microsoft、Alphabet、Amazon、Meta）2025 資本支出", "US$B", None, 410.0, None, "Verified（二手彙整）", "Tom's Hardware 引 FT 彙整：2025 約 4,100 億", U_TOMS, "2026-04-30", ""),
    ("cx_msft26", "Microsoft 曆年 2026 資本支出", "US$B", None, 170.0, None, "Interested-party（二手）", "Q3 法說：曆年 2026 約 1,700 億", U_UAO, "2026-10-29", ""),
    ("cx_goog26", "Alphabet 2026 資本支出指引", "US$B", 195.0, 200.0, 205.0, "Interested-party（二手）", "Q2 法說 195–205", U_DCK, "2026-07-30", ""),
    ("cx_amzn26", "Amazon 2026 資本支出", "US$B", None, 220.0, None, "Interested-party（二手）", "Q3：約 2,200 億", U_UAO, "2026-10-30", ""),
    ("cx_meta26", "Meta 2026 資本支出指引", "US$B", 135.0, 140.0, 145.0, "Interested-party（二手）", "Q3 收斂到 135–145", "https://letsdatascience.com/news/meta-narrows-2026-capital-spending-range-to-135-billion-145-46f12e07", "2026-10-29", ""),
    ("g_cx27", "四大 2027 資本支出 ÷ 2026", "比", 1.20, 1.35, 1.55, "Assumed", "公司只給方向：Alphabet『significantly increase』、Meta 建到 2027；無數字", U_TOMS, "2026-04/10", "含非 AI、非 NVIDIA"),
    ("hs27", "2027 年 Hyperscale 占 NVIDIA 資料中心營收", "比", 0.56, 0.52, 0.48, "Assumed", "趨勢：FY26 Q2 58.8% → FY27 Q1 57.2% → Q2 54.7%（ACIE 成長較快）", U_NVQ2, "2026-08-26", "低欄放 0.56（使顆數最低）"),
    ("cap_band", "截取比 2027 相對歷史的上下寬度", "比", 0.90, 1.00, 1.10, "Assumed", "2025、2026 兩點相差約 1–2 個百分點", "", "", ""),

    ("sec", "E. 上游：台積電 AI 加速器與封裝幾何"),
    ("g_tsmc", "台積電 AI 加速器營收 2027 ÷ 2025", "倍", 2.4, 2.6, 3.0, "Interested-party（Derived）", "台積電 2026-01：2024–2029 CAGR 『mid-to-high 50s』；2026-10 魏哲家：『probably will be higher』。1.55²／1.62²／1.73²", U_TSMC3Q26, "2026-10-15", "二手逐字稿"),
    ("g_tsmc26", "台積電 AI 加速器營收 2026 ÷ 2025", "倍", 1.55, 1.62, 1.73, "Interested-party（Derived）", "同上", U_TSMC4Q25, "2026-01-15", ""),
    ("sh_ratio27", "NVIDIA 占台積電 AI 加速器營收：2027 ÷ 2025", "比", 0.85, 0.93, 1.00, "Assumed", "TPU、Trainium、其他 ASIC 2026–27 擴量；無公司揭露", "", "", ""),
    ("sh_ratio26", "同上：2026 ÷ 2025", "比", 0.90, 0.96, 1.00, "Assumed", "同上", "", "", ""),
    ("pr_ratio27", "台積電每個 NVIDIA 封裝收入：2027 ÷ 2025", "比", 1.80, 1.55, 1.35, "Assumed／Analogy", "N4P→N3P 晶圓價、Rubin 加 I/O 小晶片、中介層 3.3x→4.5x 光罩、年度漲價；2027 仍含 Blackwell", "", "", "低欄放 1.80（使顆數最低）"),
    ("pr_ratio26", "同上：2026 ÷ 2025", "比", 1.20, 1.12, 1.05, "Assumed", "Blackwell Ultra 為主，少量 Rubin", "", "", "低欄放 1.20"),
    ("reticle", "光罩面積（26×33 mm）", "mm²", None, 858, None, "Verified", "微影標準", "", "", ""),
    ("rx_bw", "Blackwell 中介層（光罩倍數）", "倍", None, 3.3, None, "Analogy", "公開報導 3.3x；公司未揭露", "", "", ""),
    ("rx_rb", "Rubin 中介層（光罩倍數）", "倍", 5.0, 4.5, 4.0, "Analogy", "台積電 2026-10：「real-world production of 5x reticle size」", U_TSMC3Q26, "2026-10-15", "低欄放 5.0（每片最少）"),
    ("rx_ru", "Rubin Ultra 中介層（光罩倍數）", "倍", 9.0, 8.5, 8.0, "Analogy", "台積電：『working on the 9x reticle size』", U_TSMC3Q26, "2026-10-15", ""),
    ("eta", "矩形排片效率（300 mm、邊緣 3 mm）", "比", 0.61, 0.64, 0.67, "Derived", "scripts/pack.py 實際排片：3.3x 16 片、4x 12、4.5x 11、5x 10", "", "", ""),
    ("y_pkg", "CoWoS 封裝良率", "比", 0.80, 0.88, 0.95, "Assumed", "無公司揭露", "", "", ""),

    ("sec", "F. 記憶體路徑"),
    ("hbm27", "2027 HBM 位元供給（全產業）", "EB", 2.06, 3.82, 6.97, "Analogy", "HBM 獨立供給查證（三家公司未揭露位元產能）", U_HBM, "2026-11-18", ""),
    ("hbm26", "2026 HBM 位元供給（全產業）", "EB", 1.89, 3.16, 5.20, "Analogy", "同上", U_HBM, "2026-11-18", ""),
    ("nv_hbm_sh", "NVIDIA 占 HBM 位元", "比", 0.50, 0.60, 0.68, "Assumed", "NVIDIA 自行採購記憶體（CFO）；ASIC 擴量；無揭露", U_NVQ2T, "", ""),
    ("gb27", "2027 NVIDIA 每封裝平均 HBM", "GB", 320, 300, 288, "Derived", "Rubin、Blackwell Ultra 皆 288 GB；Rubin Ultra 1 TB 少量", U_BLOG, "2026-03-16", "低欄放 320"),
    ("gb26", "2026 NVIDIA 每封裝平均 HBM", "GB", 288, 270, 250, "Derived", "GB200 186、GB300 288；2025-11 起 GB300 占 Blackwell 約三分之二", U_NVQ3FY26T, "", ""),
    ("hbm_gross", "HBM 出貨 ÷ 裝載（良率、庫存、通路）", "比", 1.30, 1.15, 1.00, "Assumed", "Siliconomics In hbm_gross", "", "", "低欄放 1.30"),

    ("sec", "G. 對照標的（不是輸入）"),
    ("sil27", "Siliconomics v0.20 2027 NVIDIA 封裝", "百萬", None, 9.122, None, "Derived（券商來源）", "unit-basis 報告", U_UB, "2026-11-18", "Blackwell 2.257、Rubin 6.583、Ultra 0.282"),
    ("wat27", "Wattonomics v0.3.6 2027 NVIDIA 封裝", "百萬", None, 4.385, None, "Derived", "unit-basis 報告", U_UB, "2026-11-18", "Rubin 4.204、Ultra 0.181"),
    ("sil26", "Siliconomics 2026 NVIDIA 封裝", "百萬", None, 7.6, None, "Derived（券商來源）", "unit-basis 報告", U_UB, "", ""),
    ("fox_wk", "鴻海 AI 機櫃週產能（2025-11）", "櫃/週", None, 1000, None, "Interested-party（二手）", "劉揚偉：每週 1,000 櫃，明年提高", U_FOX, "2025-11-21", ""),
]

ws = wb.active
ws.title = "輸入"
cols = ["名稱", "項目", "單位", "低", "基準", "高", "等級", "來源", "URL", "日期", "說明"]
ws.append(["輸入：藍字＝輸入值。『低』『高』欄已排成『使 2027 顆數最低／最高』的那一端（說明欄註明反向者）。NVIDIA FY2027 第三季實績未使用。"])
ws.append(cols)
for c in range(1, 12):
    ws.cell(row=2, column=c).font = BOLD
    ws.cell(row=2, column=c).fill = HDR
REF = {}
r = 3
for it in INPUTS:
    if it[0] == "sec":
        ws.cell(row=r, column=1, value=it[1]).font = BOLD
        r += 1
        continue
    name, item, unit, lo, b, hi, grade, src, url, date, note = it
    if lo is None:
        lo = b
    if hi is None:
        hi = b
    vals = [name, item, unit, lo, b, hi, grade, src, url, date, note]
    for c, v in enumerate(vals, 1):
        cell = ws.cell(row=r, column=c, value=v)
        if c in (4, 5, 6):
            cell.font = BLUE
    REF[name] = r
    r += 1
for c, w in zip("ABCDEFGHIJK", [12, 42, 10, 10, 10, 10, 18, 50, 40, 11, 40]):
    ws.column_dimensions[c].width = w


def L(n):
    return f"輸入!$D${REF[n]}"


def B(n):
    return f"輸入!$E${REF[n]}"


def H(n):
    return f"輸入!$F${REF[n]}"


def P(n, k):  # k: 0 lo,1 base,2 hi
    return [L, B, H][k](n)


# helper sheet writer
def sheet(title, note):
    s = wb.create_sheet(title)
    s.append([note])
    s.append(["項目", "低（顆數最少）", "基準", "高（顆數最多）", "公式說明／等級"])
    for c in range(1, 6):
        s.cell(row=2, column=c).font = BOLD
        s.cell(row=2, column=c).fill = HDR
    s.column_dimensions["A"].width = 52
    for c in "BCD":
        s.column_dimensions[c].width = 15
    s.column_dimensions["E"].width = 70
    return s


def addrow(s, label, f3, note="", rowmap=None, key=None):
    """f3: function k->formula string (without '=') or list of 3."""
    rr = s.max_row + 1
    s.cell(row=rr, column=1, value=label)
    for k in range(3):
        f = f3(k) if callable(f3) else f3[k]
        s.cell(row=rr, column=2 + k, value=(("=" + f) if isinstance(f, str) else f))
    s.cell(row=rr, column=5, value=note)
    if rowmap is not None and key:
        rowmap[key] = rr
    return rr


def c(s, rr, k):
    return f"'{s.title}'!{'BCD'[k]}{rr}"


def cl(rr, k):
    return f"{'BCD'[k]}{rr}"


# ---------------- 價格 ----------------
pz = sheet("價格", "每封裝 NVIDIA 營收（千美元，含網通的全部內容）。本輪：Blackwell 四條證據線、Rubin 三條獨立證據線（R1 NVIDIA 每 GW 比值、R2 NVIDIA 密度示意、R3 Nscale 公開說明書）＋上一輪方法二（R5，Tokenomics 機架價，只作對照）。B 欄＝單價低、D 欄＝單價高。")
R = {}
addrow(pz, "【Blackwell B1】$150B ÷（600 萬晶粒÷2）（輸入 p_b25）", lambda k: P('p_b25', k), "Interested-party 推得（上一輪）", R, "b1")
BR25 = f"({B('dc_q4fy25')}/3+{B('dc_q1fy26')}+{B('dc_q2fy26')}+{B('dc_q3fy26')}+{B('dc_q4fy26')}*2/3-{B('hop25_rev')})"
addrow(pz, "【Blackwell B2】2025 年（B＋R 營收）÷（Huang『4–5 百萬 GPU』− Hopper／H20 顆數）", [f"{BR25}/({H('gpu25_jh')}-{B('hop25_u')})", f"{BR25}/({B('gpu25_jh')}-{B('hop25_u')})", f"{BR25}/({L('gpu25_jh')}-{B('hop25_u')})"], "Interested-party（二手）；GPU 按封裝讀；B＋R 營收＝曆年 2025 資料中心 − Hopper 基準", R, "b2")
addrow(pz, "【Blackwell B3】25 B$/GW ÷ GB300 每 IT GW 封裝 487,008", lambda k: f"{B('usd_gw_bw')}*1000000/{B('gpg_gb300')}", "若 NVIDIA 的 GW 是 IT 口徑", R, "b3")
addrow(pz, "  B3 若 GW 為電網口徑（÷0.80）", lambda k: f"{B('usd_gw_bw')}*1000000/({B('gpg_gb300')}*{B('it_of_grid')})", "對照：與 B1、B2、B4 不合→NVIDIA 的 GW 較接近 IT", R, "b3g")
addrow(pz, "【Blackwell B4】IREN：$5.8B ÷（200 MW IT × 487,008）× NVIDIA 占比", lambda k: f"{B('iren_usd')}*1000000/({B('iren_mw')}/1000*{B('gpg_gb300')})*{P('nsc_nvsh',k)}", "Verified 金額與 MW；NVIDIA 占比 Assumed", R, "b4")
addrow(pz, "  IREN 每 IT GW 全部設備（US$B/GW）", lambda k: f"{B('iren_usd')}/({B('iren_mw')}/1000)", "對照 NVIDIA『Blackwell 25』", R, "b4gw")
addrow(pz, "Blackwell 四線：基準最小／最大（互相印證程度）", [f"MIN(C{R['b1']},C{R['b2']},C{R['b3']},C{R['b4']})", f"MEDIAN(C{R['b1']},C{R['b2']},C{R['b3']},C{R['b4']})", f"MAX(C{R['b1']},C{R['b2']},C{R['b3']},C{R['b4']})"], "四線基準相差 < 15% → 沿用 p_b25 43／50／55", R, "bchk")
# Rubin
addrow(pz, "【Rubin R1】Blackwell 單價 ×（40÷25）× 密度比 × GPU 系統占比", lambda k: f"{P('p_b25',k)}*{B('usd_gw_rb')}/{B('usd_gw_bw')}*{P('dens_ratio',k)}*{P('gpu_share40',k)}", "同一場、同一人、同一 GW 口徑的兩個數字相除，GW 口徑抵銷；Interested-party", R, "r1")
addrow(pz, "  R2 用：Rubin 每 IT GW 封裝（NVIDIA 示意，無 MaxLPS）", lambda k: [f"{B('nv_dens_grid')}*1000/{L('lps_gain')}*10/{B('it_of_grid')}", f"{B('nv_dens_grid')}*1000/{B('lps_gain')}*10/{B('it_of_grid')}", f"{B('nv_dens_grid')}*1000/{B('lps_gain')}*10"][k], "低欄＝增益只有 25%（密度高）；高欄＝GW 為電網口徑（286K/GW）", R, "r2d")
addrow(pz, "【Rubin R2】40 B$/GW × GPU 系統占比 ÷ NVIDIA 示意密度", lambda k: f"{B('usd_gw_rb')}*{P('gpu_share40',k)}*1000000/{cl(R['r2d'],k)}", "Interested-party；與 R1 共用 40 B$ 說法", R, "r2")
addrow(pz, "  R3 用：Monarch 封裝數（千）", lambda k: f"{B('nsc_mw')}/1000*{P('nsc_dens',k)}/1000", "460 MW 按 IT", R, "r3u")
addrow(pz, "【Rubin R3】Nscale 運算設備資本支出 ÷ 封裝 × NVIDIA 占比", lambda k: f"{P('nsc_capex',k)}*1000000/({cl(R['r3u'],k)}*1000)*{P('nsc_nvsh',k)}", "Verified（公司估計）＋ MW 二手 ＋ 占比 Assumed；不用 NVIDIA 的 40 B$ 說法", R, "r3")
addrow(pz, "  R3 敏感度：若 460 MW 為總電力（IT＝×1.37/2.0）", lambda k: f"{cl(R['r3'],k)}*2/1.37", "與 R1、R2 都不重疊→本報告不採", R, "r3g")
addrow(pz, "  Monarch 每 IT GW 運算設備（US$B/GW）", lambda k: f"{P('nsc_capex',k)}/({B('nsc_mw')}/1000)", "對照 Tokenomics VR200 每 GW IT 資本支出 37.59", R, "r3gw")
addrow(pz, "  Monarch 每 IT GW 的 NVIDIA 內容（US$B/GW）", lambda k: f"{cl(R['r3gw'],k)}*{P('nsc_nvsh',k)}", "對照 NVIDIA 40、Wattonomics 26.1", R, "r3nv")
addrow(pz, "【Rubin R5，對照】Tokenomics 每 GW IT 資本支出 ÷ 封裝 × NVIDIA 比例（上一輪方法二）", lambda k: f"{B('cit_vr')}*1000000/{B('gpg_vr')}*{P('nv_take',k)}", "機架價為 Interested-party；與 Tokenomics 同源，不進重疊", R, "r5")
addrow(pz, "Rubin 每封裝 NVIDIA 營收（R1、R2、R3 重疊：低＝三者低值最大；高＝三者高值最小；基準＝三者基準中位數）",
       [f"MAX(B{R['r1']},B{R['r2']},B{R['r3']})", f"MEDIAN(C{R['r1']},C{R['r2']},C{R['r3']})", f"MIN(D{R['r1']},D{R['r2']},D{R['r3']})"],
       "注意：此列『低』是單價低（顆數多）。路徑頁取用時反向。", R, "p_rb")
addrow(pz, "  三線包絡（最寬）", [f"MIN(B{R['r1']},B{R['r2']},B{R['r3']})", f"AVERAGE(C{R['r1']},C{R['r2']},C{R['r3']})", f"MAX(D{R['r1']},D{R['r2']},D{R['r3']})"], "參考", R, "p_env")
addrow(pz, "  上一輪 Rubin 單價（對照）", lambda k: [B('prev_rb_lo'), B('prev_rb_b'), B('prev_rb_hi')][k], "", R, "p_prev")
addrow(pz, "  區間寬度：本輪 ÷ 上一輪", [f"(D{R['p_rb']}-B{R['p_rb']})/(D{R['p_prev']}-B{R['p_prev']})", None, None], "", R, "p_w")
addrow(pz, "2027 平均每封裝營收：顆數最少組合（Rubin 占比高、單價高）",
       lambda k: [f"{H('s_r27')}*D{R['p_rb']}+(1-{H('s_r27')})*{H('p_b27')}",
                  f"{B('s_r27')}*C{R['p_rb']}+(1-{B('s_r27')})*{B('p_b27')}",
                  f"{L('s_r27')}*B{R['p_rb']}+(1-{L('s_r27')})*{L('p_b27')}"][k],
       "低欄＝0.85×Rubin 高價＋0.15×Blackwell 高價；高欄＝0.65×Rubin 低價＋0.35×Blackwell 低價", R, "avg27")
addrow(pz, "2026 平均每封裝營收",
       lambda k: [f"{H('s_r26')}*D{R['p_rb']}+(1-{H('s_r26')})*{H('p_b26')}",
                  f"{B('s_r26')}*C{R['p_rb']}+(1-{B('s_r26')})*{B('p_b26')}",
                  f"{L('s_r26')}*B{R['p_rb']}+(1-{L('s_r26')})*{L('p_b26')}"][k],
       "同上（2026）", R, "avg26")
addrow(pz, "Rubin ÷ Blackwell 每封裝營收", [f"B{R['p_rb']}/{L('p_b27')}", f"C{R['p_rb']}/{B('p_b27')}", f"D{R['p_rb']}/{H('p_b27')}"], "參考", R, "ratio")
AVG27 = lambda k: f"'價格'!{'BCD'[k]}{R['avg27']}"
AVG26 = lambda k: f"'價格'!{'BCD'[k]}{R['avg26']}"

# common: DC CY2025, DC CY2026 composite (actual + guidance + assumption)
cm = sheet("共用", "共用推算：曆年資料中心營收。2026 年＝1 月（FY26 Q4 的 1/3）＋FY27 Q1、Q2 實績＋Q3 指引×資料中心占比＋Q4 假設的 2/3。2026 只有約一半是實績。")
CM = {}
addrow(cm, "曆年 2025 資料中心營收", lambda k: f"{B('dc_q4fy25')}/3+{B('dc_q1fy26')}+{B('dc_q2fy26')}+{B('dc_q3fy26')}+{B('dc_q4fy26')}*2/3", "Verified 加總（按月平均攤分）", CM, "dc25")
addrow(cm, "FY27 Q3 資料中心（指引 × 占比）", lambda k: f"{P('rev_q3g',k)}*{P('dc_sh_q3',k)}", "Interested-party", CM, "dcq3")
addrow(cm, "FY27 Q4 資料中心（假設）", lambda k: f"{cl(CM['dcq3'],k)}*(1+{P('g_q4',k)})", "Assumed", CM, "dcq4")
addrow(cm, "曆年 2026 資料中心營收（實績＋指引＋假設）", lambda k: f"{B('dc_q4fy26')}/3+{B('dc_q1')}+{B('dc_q2')}+{cl(CM['dcq3'],k)}+{cl(CM['dcq4'],k)}*2/3", "其中實績（1–7 月）", CM, "dc26")
addrow(cm, "  其中實績部分（2026-01～07）", lambda k: f"{B('dc_q4fy26')}/3+{B('dc_q1')}+{B('dc_q2')}", "Verified", CM, "dc26a")
addrow(cm, "2025 年 Blackwell＋Rubin 營收", lambda k: f"{cl(CM['dc25'],k)}-{P('hop25_rev',2-k)}", "低欄扣 Hopper 高值", CM, "br25")
addrow(cm, "2026 年 Blackwell＋Rubin 營收", lambda k: f"{cl(CM['dc26'],k)}*{P('br26',k)}", "", CM, "br26r")
addrow(cm, "2026 校驗參考顆數：NVIDIA『2,000 萬晶粒（含已出貨）至 2026 年底』推得的 2026 年封裝",
       lambda k: f"20/2-{B('bw_dies_shipped')}/2-{B('dc_q4fy26')}*2/3*{B('br26')}/{B('p_b25')}",
       "Interested-party；20M 晶粒為 Next Platform 依黃仁勳圖表所述（二手）；扣 2025-11～12 出貨。三欄同值", CM, "ref26")
addrow(cm, "2026 校驗參考顆數：實績＋指引營收 ÷ 平均單價", lambda k: f"{cl(CM['br26r'],k)}/{AVG26(k)}", "與各路徑共用單價；用來檢驗非價格步驟", CM, "ref26b")
CMc = lambda key, k: f"'共用'!{'BCD'[k]}{CM[key]}"

# ---------------- A1 ----------------
a1 = sheet("A1_營收指引", "[A] 公司自己的錢（一）：FY2028 營收指引 → 曆年 2027 資料中心 → Blackwell＋Rubin 系統營收 → ÷ 平均每封裝營收。時點＝NVIDIA 認列營收。")
A1 = {}
addrow(a1, "FY2027 營收（Q1、Q2 實績＋Q3 指引＋Q4 假設）", lambda k: f"{B('rev_q1')}+{B('rev_q2')}+{P('rev_q3g',k)}*(2+{P('g_q4',k)})", "", A1, "fy27")
addrow(a1, "FY2028 營收", lambda k: f"{cl(A1['fy27'],k)}*(1+{P('g_fy28',k)})", "Interested-party 指引", A1, "fy28")
addrow(a1, "曆年 2027 資料中心營收", lambda k: f"{cl(A1['fy28'],k)}*{P('dc_sh28',k)}*{P('cal27',k)}", "", A1, "dc27")
addrow(a1, "曆年 2027 Blackwell＋Rubin 系統營收（含網通）", lambda k: f"{cl(A1['dc27'],k)}*{P('br27',k)}", "", A1, "br27")
addrow(a1, "平均每封裝營收（千美元）", lambda k: AVG27(k), "價格頁", A1, "avg")
addrow(a1, "2027 NVIDIA 封裝（百萬）", lambda k: f"{cl(A1['br27'],k)}/{cl(A1['avg'],k)}", "＝營收（十億）÷ 單價（千）", A1, "u27")
addrow(a1, "  Blackwell（Ultra）", lambda k: f"{cl(A1['u27'],k)}*(1-{P('s_r27',2-k)})", "", A1, "ubw")
addrow(a1, "  Rubin", lambda k: f"{cl(A1['u27'],k)}*{P('s_r27',2-k)}-{P('ru27',k)}", "", A1, "urb")
addrow(a1, "  Rubin Ultra", lambda k: f"{P('ru27',k)}", "Assumed", A1, "uru")
addrow(a1, "隱含 IT GW（Tokenomics 每 GW 封裝）", lambda k: f"({cl(A1['ubw'],k)}/{B('gpg_gb300')}+{cl(A1['urb'],k)}/{B('gpg_vr')}+{cl(A1['uru'],k)}/{B('gpg_ru')})*1000000", "對照：Siliconomics 28.65、Wattonomics 15.34", A1, "gw")
addrow(a1, "【翻轉】達 912 萬所需平均單價（千美元）", lambda k: f"{cl(A1['br27'],k)}/{B('sil27')}", "對照基準平均單價", A1, "f1")
addrow(a1, "【翻轉】達 438 萬所需平均單價（千美元）", lambda k: f"{cl(A1['br27'],k)}/{B('wat27')}", "", A1, "f2")
addrow(a1, "【翻轉】基準單價下達 912 萬所需資料中心營收 ÷ 本路徑值", lambda k: f"{B('sil27')}*{cl(A1['avg'],k)}/{cl(A1['br27'],k)}", "", A1, "f3")
addrow(a1, "【翻轉】基準單價下達 438 萬所需資料中心營收 ÷ 本路徑值", lambda k: f"{B('wat27')}*{cl(A1['avg'],k)}/{cl(A1['br27'],k)}", "", A1, "f4")
addrow(a1, "【校驗 2026】同法：曆年 2026 封裝（實績＋指引營收 ÷ 單價）", lambda k: CMc('ref26b', k), "營收步驟用實績與指引（不需預測）；誤差只來自單價", A1, "c26")
addrow(a1, "【校驗 2026】對 NVIDIA『2,000 萬晶粒』推得值的誤差", lambda k: f"{cl(A1['c26'],k)}/{CMc('ref26',1)}-1", "兩者都用 NVIDIA 說法，非獨立實績", A1, "e26")
addrow(a1, "【校驗】指引可靠度：$500B 能見度 vs 實現（見 A2）", lambda k: f"'A2_訂單'!C{6}", "參考", A1, "e26b")

# ---------------- A2 ----------------
a2 = sheet("A2_訂單", "[A] 公司自己的錢（二）：GTC 2026『至 2027 年 1 兆美元』（讀為 2025–2027 三年）− 2025、2026 已實現 → 2027 ÷ 單價。")
A2 = {}
addrow(a2, "2025 年 Blackwell＋Rubin 營收", lambda k: CMc('br25', k), "共用頁", A2, "br25")
addrow(a2, "2026 年 Blackwell＋Rubin 營收（實績＋指引＋假設）", lambda k: CMc('br26r', k), "共用頁", A2, "br26")
addrow(a2, "2025–2026 合計", lambda k: f"{cl(A2['br25'],k)}+{cl(A2['br26'],k)}", "", A2, "cum")
addrow(a2, "【校驗】2025–2026 實現 ÷ $500B 說法", lambda k: f"{cl(A2['cum'],k)}/{B('br_500')}", "> 1 表示 NVIDIA 2025-11 的說法偏保守", A2, "k500")
addrow(a2, "2027 年 Blackwell＋Rubin 營收＝1 兆 × 實現係數 − 前兩年", lambda k: f"{B('br_1t')}*{P('k_overrun',k)}-{cl(A2['cum'],2-k)}", "低欄扣前兩年高值", A2, "br27")
addrow(a2, "平均每封裝營收", lambda k: AVG27(k), "", A2, "avg")
addrow(a2, "2027 NVIDIA 封裝（百萬）", lambda k: f"{cl(A2['br27'],k)}/{cl(A2['avg'],k)}", "", A2, "u27")
addrow(a2, "對照 A1 營收：本路徑 ÷ A1", lambda k: f"{cl(A2['br27'],k)}/'A1_營收指引'!{'BCD'[k]}{A1['br27']}", "1 兆說法（2026-03）早於 +70% 指引（2026-08）", A2, "vsA1")
addrow(a2, "【翻轉】基準單價下達 912 萬所需 1 兆的倍數", lambda k: f"({B('sil27')}*{cl(A2['avg'],k)}+{cl(A2['cum'],2-k)})/{B('br_1t')}", "", A2, "f1")
addrow(a2, "【翻轉】基準單價下達 438 萬所需 1 兆的倍數", lambda k: f"({B('wat27')}*{cl(A2['avg'],k)}+{cl(A2['cum'],2-k)})/{B('br_1t')}", "", A2, "f2")
addrow(a2, "【校驗 2026】$500B 說法 × 實現係數推 2026 營收 ÷ 單價", lambda k: f"({B('br_500')}*{P('k_overrun',k)}-{cl(A2['br25'],2-k)})/{AVG26(k)}", "", A2, "c26")
addrow(a2, "【校驗 2026】對 NVIDIA『2,000 萬晶粒』推得值的誤差", lambda k: f"{cl(A2['c26'],k)}/{CMc('ref26',1)}-1", "", A2, "e26")
# fix A1 reference to k500 row
a1.cell(row=A1['e26b'], column=2, value=f"='A2_訂單'!B{A2['k500']}")
a1.cell(row=A1['e26b'], column=3, value=f"='A2_訂單'!C{A2['k500']}")
a1.cell(row=A1['e26b'], column=4, value=f"='A2_訂單'!D{A2['k500']}")

# ---------------- A3 ----------------
a3 = sheet("A3_承諾下限", "[A] 公司自己的錢（三）：FY2028 到期的供應承諾（截至 2026-07）至少能覆蓋多少封裝的銷貨成本。只當下限；承諾之後還會增加，且含多年記憶體。")
A3 = {}
addrow(a3, "FY2028 到期承諾", lambda k: B('commit28'), "Verified", A3, "c")
addrow(a3, "其中屬 GPU 系統（Assumed 0.8／0.9／1.0）", [f"{B('commit28')}*0.8", f"{B('commit28')}*0.9", f"{B('commit28')}*1.0"], "Assumed", A3, "cg")
addrow(a3, "每封裝銷貨成本（千美元）＝平均單價 ×（1 − 毛利率）", lambda k: f"{AVG27(k)}*(1-{P('gm28',k)})", "", A3, "cogs")
addrow(a3, "已承諾可覆蓋的封裝（百萬）", lambda k: f"{cl(A3['cg'],k)}/{cl(A3['cogs'],k)}", "下限，非估計", A3, "u")

# ---------------- D ----------------
dd = sheet("D_買方資本支出", "[A／P] 下游承接：四大買方資本支出 × NVIDIA 截取比（以 NVIDIA 公布的 Hyperscale 營收校準）÷ Hyperscale 占比 → 資料中心營收 → ÷ 單價。不讀 Siliconomics 份額。")
D = {}
addrow(dd, "2025 Hyperscale 占資料中心（FY26 Q2）", lambda k: f"{B('hs_q2fy26')}/{B('dc_q2fy26')}", "Verified 推得", D, "hs25")
addrow(dd, "2025 NVIDIA Hyperscale 營收（估）", lambda k: f"{CMc('dc25',k)}*{cl(D['hs25'],k)}", "Derived", D, "hsr25")
addrow(dd, "截取比 2025＝Hyperscale 營收 ÷ 四大資本支出", lambda k: f"{cl(D['hsr25'],k)}/{B('capex25')}", "", D, "cap25")
addrow(dd, "2026 Hyperscale 占資料中心（FY27 Q1＋Q2）", lambda k: f"({B('hs_q1')}+{B('hs_q2')})/({B('dc_q1')}+{B('dc_q2')})", "Verified 推得", D, "hs26")
addrow(dd, "四大 2026 資本支出", lambda k: f"{B('cx_msft26')}+{P('cx_goog26',k)}+{B('cx_amzn26')}+{P('cx_meta26',k)}", "公司指引（二手）", D, "cx26")
addrow(dd, "截取比 2026", lambda k: f"{CMc('dc26',k)}*{cl(D['hs26'],k)}/{cl(D['cx26'],k)}", "2026 一半是指引與假設", D, "cap26")
addrow(dd, "【校驗 2026】用 2025 截取比 × 2026 資本支出 ÷ 2026 Hyperscale 占比 → 資料中心營收", lambda k: f"{cl(D['cap25'],k)}*{cl(D['cx26'],k)}/{cl(D['hs26'],k)}", "", D, "c26rev")
addrow(dd, "【校驗 2026】對 2026 資料中心（實績＋指引）的誤差", lambda k: f"{cl(D['c26rev'],k)}/{CMc('dc26',1)}-1", "營收步驟的校驗誤差", D, "e26rev")
addrow(dd, "【校驗 2026】同法封裝 ÷ NVIDIA『2,000 萬晶粒』推得值 − 1", lambda k: f"{cl(D['c26rev'],k)}*{P('br26',k)}/{AVG26(k)}/{CMc('ref26',1)}-1", "", D, "e26")
addrow(dd, "2027 截取比", lambda k: f"AVERAGE(C{D['cap25']},C{D['cap26']})*{P('cap_band',k)}", "Assumed 寬度", D, "cap27")
addrow(dd, "四大 2027 資本支出", lambda k: f"C{D['cx26']}*{P('g_cx27',k)}", "Assumed 成長", D, "cx27")
addrow(dd, "2027 NVIDIA Hyperscale 營收", lambda k: f"{cl(D['cx27'],k)}*{cl(D['cap27'],k)}", "", D, "hsr27")
addrow(dd, "2027 NVIDIA 資料中心營收", lambda k: f"{cl(D['hsr27'],k)}/{P('hs27',k)}", "", D, "dc27")
addrow(dd, "對照 A1 資料中心營收：本路徑 ÷ A1", lambda k: f"{cl(D['dc27'],k)}/'A1_營收指引'!{'BCD'[k]}{A1['dc27']}", "", D, "vsA1")
addrow(dd, "2027 NVIDIA 封裝（百萬）", lambda k: f"{cl(D['dc27'],k)}*{P('br27',k)}/{AVG27(k)}", "", D, "u27")
addrow(dd, "【翻轉】基準截取比與單價下，達 912 萬所需四大資本支出成長", lambda k: f"{B('sil27')}*{AVG27(1)}/{B('br27')}*{B('hs27')}/C{D['cap27']}/C{D['cx26']}", "", D, "f1")
addrow(dd, "【翻轉】達 438 萬所需四大資本支出成長", lambda k: f"{B('wat27')}*{AVG27(1)}/{B('br27')}*{B('hs27')}/C{D['cap27']}/C{D['cx26']}", "", D, "f2")

# ---------------- 幾何 ----------------
ge = sheet("幾何", "[P] CoWoS 中介層幾何：每片 300 mm 晶圓可切中介層數＝INT(π×147² ÷ 面積 × 排片效率)；好封裝＝× 良率。下半部：各情境所需 NVIDIA CoWoS 片數。台積電未揭露 CoWoS 片數，所以這頁只換算『需要多少片』，不估產能。")
G = {}
addrow(ge, "Blackwell 中介層面積 mm²", lambda k: f"{B('rx_bw')}*{B('reticle')}", "Analogy", G, "abw")
addrow(ge, "Rubin 中介層面積 mm²", lambda k: f"{P('rx_rb',k)}*{B('reticle')}", "Analogy；台積電稱 5x 已量產", G, "arb")
addrow(ge, "Rubin Ultra 中介層面積 mm²", lambda k: f"{P('rx_ru',k)}*{B('reticle')}", "Analogy", G, "aru")
addrow(ge, "Blackwell 每片毛中介層", lambda k: f"INT(PI()*147^2/{cl(G['abw'],k)}*{P('eta',k)})", "", G, "gbw")
addrow(ge, "Rubin 每片毛中介層", lambda k: f"INT(PI()*147^2/{cl(G['arb'],k)}*{P('eta',k)})", "", G, "grb")
addrow(ge, "Rubin Ultra 每片毛中介層", lambda k: f"MAX(1,INT(PI()*147^2/{cl(G['aru'],k)}*{P('eta',k)}))", "", G, "gru")
addrow(ge, "Blackwell 每片好封裝", lambda k: f"{cl(G['gbw'],k)}*{P('y_pkg',k)}", "", G, "ybw")
addrow(ge, "Rubin 每片好封裝", lambda k: f"{cl(G['grb'],k)}*{P('y_pkg',k)}", "", G, "yrb")
addrow(ge, "Rubin Ultra 每片好封裝", lambda k: f"{cl(G['gru'],k)}*{P('y_pkg',k)}", "", G, "yru")
addrow(ge, "Siliconomics 912 萬所需片數（千片／年；欄＝幾何低／基準／高）", lambda k: f"(2.257/{cl(G['ybw'],k)}+6.583/{cl(G['yrb'],k)}+0.282/{cl(G['yru'],k)})*1000", "低欄＝每片最少→片數最多", G, "wsil")
addrow(ge, "Wattonomics 438 萬所需片數（千片／年）", lambda k: f"(4.204/{cl(G['yrb'],k)}+0.181/{cl(G['yru'],k)})*1000", "", G, "wwat")
addrow(ge, "本報告基準（A1 基準拆分）所需片數（千片／年）", lambda k: f"('A1_營收指引'!C{A1['ubw']}/{cl(G['ybw'],k)}+'A1_營收指引'!C{A1['urb']}/{cl(G['yrb'],k)}+'A1_營收指引'!C{A1['uru']}/{cl(G['yru'],k)})*1000", "", G, "wours")
addrow(ge, "  換成月產能（千片／月）：912 萬", lambda k: f"{cl(G['wsil'],k)}/12", "", G, "msil")
addrow(ge, "  換成月產能（千片／月）：438 萬", lambda k: f"{cl(G['wwat'],k)}/12", "", G, "mwat")
addrow(ge, "  換成月產能（千片／月）：本報告基準", lambda k: f"{cl(G['wours'],k)}/12", "", G, "mours")

# ---------------- P1 ----------------
p1 = sheet("P1_台積電", "[P] 上游產能（以台積電公司揭露的 AI 加速器營收成長代替未揭露的 CoWoS 片數）：2025 NVIDIA 封裝 × 台積電 AI 加速器營收倍數 × NVIDIA 占比變化 ÷ 每封裝台積電收入變化。")
P1 = {}
addrow(p1, "2025 NVIDIA 封裝錨點（百萬）＝2025 B＋R 營收 ÷ 2025 單價 ＋ Hopper/H20", lambda k: f"{CMc('br25',k)}/{P('p_b25',2-k)}+{P('hop25_u',k)}", "營收 Verified；單價 Interested-party 推得", P1, "u25")
addrow(p1, "台積電 AI 加速器營收 2027 ÷ 2025", lambda k: P('g_tsmc', k), "", P1, "g")
addrow(p1, "NVIDIA 占比變化", lambda k: P('sh_ratio27', k), "", P1, "s")
addrow(p1, "每封裝台積電收入變化", lambda k: P('pr_ratio27', k), "", P1, "p")
addrow(p1, "2027 NVIDIA 封裝（百萬）", lambda k: f"{cl(P1['u25'],k)}*{cl(P1['g'],k)}*{cl(P1['s'],k)}/{cl(P1['p'],k)}", "時點＝晶圓與封裝產出（早於認列約 1–2 個月）", P1, "u27")
addrow(p1, "【翻轉】其他基準下達 912 萬所需台積電 AI 營收倍數（2025→27）", lambda k: f"{B('sil27')}*C{P1['p']}/C{P1['u25']}/C{P1['s']}", "對照公司 CAGR 隱含 2.4–3.0", P1, "f1")
addrow(p1, "【翻轉】達 438 萬所需倍數", lambda k: f"{B('wat27')}*C{P1['p']}/C{P1['u25']}/C{P1['s']}", "", P1, "f2")
addrow(p1, "【翻轉】達 912 萬所需每封裝收入變化（其他基準）", lambda k: f"C{P1['u25']}*C{P1['g']}*C{P1['s']}/{B('sil27')}", "", P1, "f3")
addrow(p1, "【校驗 2026】同法 2026 封裝", lambda k: f"{cl(P1['u25'],k)}*{P('g_tsmc26',k)}*{P('sh_ratio26',k)}/{P('pr_ratio26',k)}", "", P1, "c26")
addrow(p1, "【校驗 2026】對 NVIDIA『2,000 萬晶粒』推得值的誤差", lambda k: f"{cl(P1['c26'],k)}/{CMc('ref26',1)}-1", "", P1, "e26")
addrow(p1, "【校驗 2026】換成營收 ÷ 2026 資料中心（實績＋指引）− 1", lambda k: f"{cl(P1['c26'],k)}*{AVG26(1)}/{B('br26')}/{CMc('dc26',1)}-1", "以基準單價換算", P1, "e26r")

# ---------------- M ----------------
mm = sheet("M_HBM", "[P] 記憶體：HBM 全產業位元 × NVIDIA 占比 ÷（每封裝 GB × 出貨÷裝載）。位元供給沿用 HBM 查證報告（三家都未揭露位元產能，區間很寬）。")
M = {}
addrow(mm, "2027 HBM 位元（EB）", lambda k: P('hbm27', k), "Analogy", M, "eb")
addrow(mm, "NVIDIA 占比", lambda k: P('nv_hbm_sh', k), "Assumed", M, "sh")
addrow(mm, "每封裝 GB", lambda k: P('gb27', k), "Derived", M, "gb")
addrow(mm, "出貨 ÷ 裝載", lambda k: P('hbm_gross', k), "Assumed", M, "gr")
addrow(mm, "2027 NVIDIA 封裝（百萬）", lambda k: f"{cl(M['eb'],k)}*1000*{cl(M['sh'],k)}/({cl(M['gb'],k)}*{cl(M['gr'],k)})", "EB×10⁹ GB ÷ GB ÷ 10⁶；時點＝HBM 出貨（早於認列 1–2 季）", M, "u27")
addrow(mm, "【翻轉】達 912 萬所需全產業 HBM（EB；其他基準）", lambda k: f"{B('sil27')}*C{M['gb']}*C{M['gr']}/C{M['sh']}/1000", "對照獨立基準 3.82", M, "f1")
addrow(mm, "【翻轉】達 438 萬所需全產業 HBM（EB）", lambda k: f"{B('wat27')}*C{M['gb']}*C{M['gr']}/C{M['sh']}/1000", "", M, "f2")
addrow(mm, "【校驗 2026】同法 2026 封裝", lambda k: f"{P('hbm26',k)}*1000*{P('nv_hbm_sh',k)}/({P('gb26',k)}*{P('hbm_gross',k)})", "", M, "c26")
addrow(mm, "【校驗 2026】對 NVIDIA『2,000 萬晶粒』推得值的誤差", lambda k: f"{cl(M['c26'],k)}/{CMc('ref26',1)}-1", "", M, "e26")

# ---------------- O：OEM ----------------
oo = sheet("O_OEM", "[A] 第六條路徑（任務二）：Dell、SMCI、HPE 的 AI 伺服器營收與積壓 → 2027 年 OEM 通路的 NVIDIA 內容 → 封裝數下限。只能約束『下限』與『OEM 通路占比』；不能約束單價、也看不到 ODM 直供超大型雲端的量（鴻海、廣達、緯創、緯穎）。")
O = {}
addrow(oo, "Dell FY27 下半年（2026-08～2027-01）AI 營收＝全年指引 − 上半年", lambda k: f"{B('dell_fy27')}-{B('dell_h1')}", "Interested-party 指引", O, "dh2")
addrow(oo, "Dell 2026-07 積壓中，2027-02 以後才出貨的部分（未計 8 月後新訂單）", lambda k: f"{B('dell_bl')}-{cl(O['dh2'],k)}", "下限性質：只計已在手的訂單", O, "dcarry")
addrow(oo, "  其中 2027 年內出貨比例（Assumed 0.8／0.9／1.0）", [f"{cl(O['dcarry'],0)}*0.8", f"{cl(O['dcarry'],1)}*0.9", f"{cl(O['dcarry'],2)}*1.0"], "Assumed", O, "d27")
addrow(oo, "SMCI 2027-01～06 營收＝FY27 指引 − Q1 − Q2", lambda k: f"{P('smci_fy27',k)}-{P('smci_q1',k)}*(1+{P('smci_q2r',2-k)})", "低欄扣 Q2 高值", O, "s27")
addrow(oo, "  其中 AI GPU 平台", lambda k: f"{cl(O['s27'],k)}*{P('smci_ai',k)}", "Assumed 占比", O, "s27ai")
addrow(oo, "HPE AI 系統（年化，2027 按持平）", lambda k: B('hpe_ai'), "Verified（二手）", O, "h27")
addrow(oo, "OEM 2027 已可見 AI 伺服器營收（下限）", lambda k: f"{cl(O['d27'],k)}+{cl(O['s27ai'],k)}+{cl(O['h27'],k)}", "不含 Dell 8 月後新訂單、SMCI 2027 下半年、Lenovo", O, "oem27")
addrow(oo, "其中 NVIDIA 內容", lambda k: f"{cl(O['oem27'],k)}*{P('oem_nv',k)}", "Assumed 比例", O, "nv27")
addrow(oo, "OEM 通路 2027 NVIDIA 封裝下限（百萬）", lambda k: f"{cl(O['nv27'],k)}/{AVG27(k)}", "÷ 平均單價（價格頁）", O, "u27")
addrow(oo, "  占本報告 2027 最可能值的比例", [0, 0, 0], "（互驗頁建好後回填公式）", O, "shr")
addrow(oo, "【診斷】Dell FY27 AI 營收 × NVIDIA 內容 ÷ NVIDIA 曆年 2026 資料中心", lambda k: f"{B('dell_fy27')}*{P('oem_nv',k)}/{CMc('dc26',1)}", "Dell 一家約占 NVIDIA 資料中心營收的比例（月份錯開一個月）", O, "dsh")
addrow(oo, "【診斷】912 萬所需 Blackwell＋Rubin 營收 ÷ A1 指引推得值", lambda k: f"{B('sil27')}*{AVG27(1)}/'A1_營收指引'!C{A1['br27']}", "OEM 營收若在 2027-02 起季增明顯超過 NVIDIA 指引隱含速度，才支持此值", O, "need912")

# ---------------- 每 GW 檢查 ----------------
gw = sheet("每GW檢查", "任務三：依更新後的 Rubin 單價，檢查 Tokenomics 每 GW 封裝數（VR200 291,744）與 Wattonomics 每 GW NVIDIA 內容（約 26.1 B$）哪一個與證據不合。只指出問題，不是改值建議。")
W = {}
addrow(gw, "證據：Rubin 世代每 IT GW 的 NVIDIA 內容（US$B；低＝Monarch 低、基準＝NVIDIA 40×GPU 系統占比與 Monarch 基準平均、高＝NVIDIA 40 全額）",
       [f"'價格'!B{R['r3nv']}", f"AVERAGE({B('usd_gw_rb')}*{B('gpu_share40')},'價格'!C{R['r3nv']})", f"{B('usd_gw_rb')}"], "Interested-party＋Verified／Assumed", W, "cont")
addrow(gw, "Rubin 單價（千美元；價格頁重疊）", lambda k: f"'價格'!{'BCD'[k]}{R['p_rb']}", "", W, "pr")
addrow(gw, "證據隱含：每 IT GW Rubin 封裝數＝內容 ÷ 單價（低＝內容低÷單價高）", [f"B{W['cont']}*1000000/D{W['pr']}", f"C{W['cont']}*1000000/C{W['pr']}", f"D{W['cont']}*1000000/B{W['pr']}"], "", W, "dens")
addrow(gw, "Tokenomics VR200 每 IT GW 封裝", lambda k: B('gpg_vr'), "Derived", W, "tk")
addrow(gw, "  Tokenomics ÷ 證據隱含基準 − 1", [None, f"C{W['tk']}/C{W['dens']}-1", None], "", W, "tkgap")
addrow(gw, "  Tokenomics 是否在證據區間內", [None, f'IF(AND(C{W['tk']}>=B{W['dens']},C{W['tk']}<=D{W['dens']}),"是","否")', None], "", W, "tkin")
addrow(gw, "NVIDIA 示意密度（IT 口徑，無 MaxLPS）", lambda k: f"'價格'!C{R['r2d']}", "Interested-party", W, "nvd")
addrow(gw, "SINES：66,000 顆 ÷ 200 MW（若整個園區都給它）", lambda k: f"{B('sinesgpu')}*1000/({B('sines_mw')}/1000)", "下限性質的密度（園區 MW 口徑未寫）", W, "sines")
addrow(gw, "Wattonomics 每 GW NVIDIA 內容", lambda k: B('wat_gw_nv'), "對照標的", W, "wat")
addrow(gw, "證據：單價 × Tokenomics 密度（US$B/GW）", lambda k: f"'價格'!{'BCD'[k]}{R['p_rb']}*{B('gpg_vr')}/1000000", "", W, "pxd")
addrow(gw, "  Wattonomics ÷（單價 × 密度）基準 − 1", [None, f"C{W['wat']}/C{W['pxd']}-1", None], "", W, "watgap")
addrow(gw, "  Wattonomics ÷ 直接內容證據基準 − 1", [None, f"C{W['wat']}/C{W['cont']}-1", None], "", W, "watgap2")
addrow(gw, "  Wattonomics 是否在（單價 × 密度）區間內", [None, f'IF(AND(C{W['wat']}>=B{W['pxd']},C{W['wat']}<=D{W['pxd']}),"是","否")', None], "", W, "watin")
addrow(gw, "Wattonomics 2027 NVIDIA GW × 證據內容（US$B）", lambda k: f"{B('wat_gw27')}*{cl(W['cont'],k)}", "對照 A1 B＋R 系統營收", W, "watrev")
addrow(gw, "  A1 2027 Blackwell＋Rubin 系統營收（US$B）", lambda k: f"'A1_營收指引'!{'BCD'[k]}{A1['br27']}", "", W, "a1rev")
addrow(gw, "Wattonomics 2027 NVIDIA GW × 26.1（US$B）", lambda k: f"{B('wat_gw27')}*{B('wat_gw_nv')}", "", W, "watrev2")
addrow(gw, "Tokenomics VR200 每 GW IT 資本支出 對 Monarch 每 IT GW 運算設備 − 1", lambda k: f"{B('cit_vr')}/'價格'!{'BCD'[2-k]}{R['r3gw']}-1", "Monarch 只含運算設備（不含 DC），Tokenomics 含全部 IT", W, "citgap")

# ---------------- 互驗 ----------------
xv = wb.create_sheet("互驗")
xv.append(["互驗：2027 年 NVIDIA 加速器封裝（百萬；曆年；NVIDIA 認列營收時點，P1／M 為產出時點）。獨立性：A1／A2／D 營收來源不同但共用單價；P1、M 不用 NVIDIA 單價。"])
hdr = ["路徑", "2027 低", "2027 基準", "2027 高", "含 438 萬？", "含 912 萬？", "2026 同法基準", "2026 校驗誤差（對 NVIDIA 2,000 萬晶粒推得值）", "主要共用輸入", "進入重疊計算？"]
xv.append(hdr)
for i in range(1, 11):
    xv.cell(row=2, column=i).font = BOLD
    xv.cell(row=2, column=i).fill = HDR
rows = [
    ("A1 營收指引 ÷ 單價", "A1_營收指引", A1['u27'], A1['c26'], A1['e26'], "每封裝單價", "是"),
    ("A2 1 兆訂單 ÷ 單價", "A2_訂單", A2['u27'], A2['c26'], A2['e26'], "每封裝單價；與 A1 同為 NVIDIA 說法", "否（與 A1 同源，取 A1）"),
    ("D 四大資本支出 × 截取比 ÷ 單價", "D_買方資本支出", D['u27'], None, D['e26'], "每封裝單價", "是"),
    ("P1 台積電 AI 加速器營收", "P1_台積電", P1['u27'], P1['c26'], P1['e26'], "2025 錨點用 NVIDIA 單價", "是"),
    ("M HBM 位元", "M_HBM", M['u27'], M['c26'], M['e26'], "無", "是"),
]
XR = {}
for lab, sh, ur, c26, e26, shared, inc in rows:
    rr = xv.max_row + 1
    xv.cell(row=rr, column=1, value=lab)
    for k in range(3):
        xv.cell(row=rr, column=2 + k, value=f"='{sh}'!{'BCD'[k]}{ur}")
    xv.cell(row=rr, column=5, value=f'=IF(AND(B{rr}<={B("wat27")},D{rr}>={B("wat27")}),"是","否")')
    xv.cell(row=rr, column=6, value=f'=IF(AND(B{rr}<={B("sil27")},D{rr}>={B("sil27")}),"是","否")')
    if c26:
        xv.cell(row=rr, column=7, value=f"='{sh}'!C{c26}")
    else:
        xv.cell(row=rr, column=7, value=f"='{sh}'!C{D['c26rev']}*{B('br26')}/{AVG26(1)}")
    xv.cell(row=rr, column=8, value=f"='{sh}'!C{e26}")
    xv.cell(row=rr, column=9, value=shared)
    xv.cell(row=rr, column=10, value=inc)
    XR[sh] = rr
rr = xv.max_row + 1
xv.cell(row=rr, column=1, value="A3 承諾可覆蓋（下限）")
for k in range(3):
    xv.cell(row=rr, column=2 + k, value=f"='A3_承諾下限'!{'BCD'[k]}{A3['u']}")
xv.cell(row=rr, column=9, value="下限，不進重疊")
rr += 1
xv.cell(row=rr, column=1, value="O OEM 通路已可見營收（下限；任務二）")
for k in range(3):
    xv.cell(row=rr, column=2 + k, value=f"='O_OEM'!{'BCD'[k]}{O['u27']}")
xv.cell(row=rr, column=9, value="下限，不進重疊；看不到 ODM 直供")
rr += 1
xv.cell(row=rr, column=1, value="鴻海產能上限檢查（2025-11 週產能 × 50 週 × 72 ÷ 鴻海占比 0.5 ÷ NVL72 占封裝 0.85；2027 產能成長 1.6 倍 Assumed）")
xv.cell(row=rr, column=3, value=f"={B('fox_wk')}*1.6*50*72/0.5/0.85/1000000")
xv.cell(row=rr, column=9, value="上限，不進重疊；占比與成長皆 Assumed")
rr += 2
inc_rows = [XR['A1_營收指引'], XR['D_買方資本支出'], XR['P1_台積電'], XR['M_HBM']]
xv.cell(row=rr, column=1, value="重疊區（A1、D、P1、M）：下緣＝各低值最大者").font = BOLD
xv.cell(row=rr, column=2, value="=MAX(" + ",".join(f"B{x}" for x in inc_rows) + ")")
OV_LO = rr
rr += 1
xv.cell(row=rr, column=1, value="重疊區上緣＝各高值最小者").font = BOLD
xv.cell(row=rr, column=2, value="=MIN(" + ",".join(f"D{x}" for x in inc_rows) + ")")
OV_HI = rr
rr += 1
xv.cell(row=rr, column=1, value="四條路徑基準值中位數").font = BOLD
xv.cell(row=rr, column=2, value="=MEDIAN(" + ",".join(f"C{x}" for x in inc_rows) + ")")
MED = rr
rr += 1
xv.cell(row=rr, column=1, value="本報告看法：區間低／最可能／高（低高＝重疊區；最可能＝四條路徑基準中位數）").font = BOLD
xv.cell(row=rr, column=2, value=f"=B{OV_LO}")
xv.cell(row=rr, column=3, value=f"=B{MED}")
xv.cell(row=rr, column=4, value=f"=B{OV_HI}")
OURS = rr
for k in range(3):
    oo.cell(row=O['shr'], column=2 + k, value=f"={cl(O['u27'],k)}/'互驗'!C{OURS}")
rr += 1
xv.cell(row=rr, column=1, value="本報告看法換成 IT GW（以 A1 世代拆分比例換算）")
for k in range(3):
    xv.cell(row=rr, column=2 + k, value=f"={'BCD'[k]}{OURS}/'A1_營收指引'!C{A1['u27']}*'A1_營收指引'!C{A1['gw']}")
rr += 1
xv.cell(row=rr, column=1, value="上一輪（nvunits）：低／最可能／高")
xv.cell(row=rr, column=2, value=f"={B('prev_lo')}")
xv.cell(row=rr, column=3, value=f"={B('prev_med')}")
xv.cell(row=rr, column=4, value=f"={B('prev_hi')}")
PREV = rr
rr += 1
xv.cell(row=rr, column=1, value="區間寬度：本輪 ÷ 上一輪；最可能值變動")
xv.cell(row=rr, column=2, value=f"=(D{OURS}-B{OURS})/(D{PREV}-B{PREV})")
xv.cell(row=rr, column=3, value=f"=C{OURS}-C{PREV}")
rr += 1
xv.cell(row=rr, column=1, value="912 萬 ÷ 本報告最可能值")
xv.cell(row=rr, column=3, value=f"={B('sil27')}/C{OURS}")
rr += 1
xv.cell(row=rr, column=1, value="438 萬 ÷ 本報告最可能值")
xv.cell(row=rr, column=3, value=f"={B('wat27')}/C{OURS}")
rr += 1
xv.cell(row=rr, column=1, value="2026 年：Siliconomics 760 萬 ÷ NVIDIA『2,000 萬晶粒』推得值")
xv.cell(row=rr, column=3, value=f"={B('sil26')}/'共用'!C{CM['ref26']}")
rr += 1
xv.cell(row=rr, column=1, value="2026 年：NVIDIA『2,000 萬晶粒』推得值（百萬封裝）")
xv.cell(row=rr, column=3, value=f"='共用'!C{CM['ref26']}")
rr += 1
xv.cell(row=rr, column=1, value="2026 年：實績＋指引營收 ÷ 單價（百萬封裝）")
for k in range(3):
    xv.cell(row=rr, column=2 + k, value=f"='共用'!{'BCD'[k]}{CM['ref26b']}")
rr += 2
xv.cell(row=rr, column=1, value="區間圖數據（橫條：低→高；點：基準）").font = BOLD
rr += 1
xv.append(["路徑", "低", "基準", "高", "條長（高−低）"])
for lab, sh, *_ in rows:
    r2 = xv.max_row + 1
    xv.cell(row=r2, column=1, value=lab)
    x = XR[sh]
    xv.cell(row=r2, column=2, value=f"=B{x}")
    xv.cell(row=r2, column=3, value=f"=C{x}")
    xv.cell(row=r2, column=4, value=f"=D{x}")
    xv.cell(row=r2, column=5, value=f"=D{x}-B{x}")
r2 = xv.max_row + 1
xv.cell(row=r2, column=1, value="Siliconomics（點）")
xv.cell(row=r2, column=3, value=f"={B('sil27')}")
r2 += 1
xv.cell(row=r2, column=1, value="Wattonomics（點）")
xv.cell(row=r2, column=3, value=f"={B('wat27')}")
xv.column_dimensions["A"].width = 60
for cc in "BCDEFGH":
    xv.column_dimensions[cc].width = 14
xv.column_dimensions["I"].width = 36
xv.column_dimensions["J"].width = 22
wb.move_sheet("互驗", offset=-(len(wb.sheetnames) - 2))

# ---------------- 來源 ----------------
so = wb.create_sheet("來源")
so.append(["來源", "日期", "URL", "等級", "短引（≤15 字）或狀態"])
for i in range(1, 6):
    so.cell(row=1, column=i).font = BOLD
SRC = [
    ("NVIDIA Q2 FY27 CFO commentary（公司）", "2026-08-26", U_NVQ2, "Verified／Interested-party", "primarily related to the procurement of memory"),
    ("NVIDIA Q2 FY27 法說逐字稿（公司 IR 網站 PDF）", "2026-08-26", U_NVQ2T, "Interested-party", "grow revenue by approximately 70% in fiscal 2028"),
    ("NVIDIA Q1 FY27 CFO commentary", "2026-05", U_NVQ1, "Verified", "供應承諾 119、網通 14.8"),
    ("NVIDIA Q2 FY26 CFO commentary", "2025-08-27", U_NVQ2FY26, "Verified", "運算 33.8、網通 7.3"),
    ("NVIDIA Q3 FY26 法說逐字稿（Motley Fool，二手）", "2025-11-19", U_NVQ3FY26T, "Interested-party（二手）", "half a trillion dollars in Blackwell and Rubin revenue"),
    ("The Next Platform：GTC DC（二手）", "2025-10-29", U_GTCDC, "Interested-party（二手）", "shipped 6 million of the Blackwells"),
    ("NVIDIA GTC 2025 keynote deck", "2025-03-18", U_GTC25, "Interested-party", "2-GPUs per Chip（unit-basis 報告已讀）"),
    ("Let's Data Science：GTC 2026 1 兆（二手）", "2026-03-16", U_1T, "Interested-party（二手）", "原文未取得；轉述『through 2027』"),
    ("NVIDIA FY2026 10-K（MarketScreener 轉載）", "2026-02-25", U_10K, "Verified", "運算成長 59%、網通 142%；Note 12 承諾表未取得"),
    ("台積電 2025Q4 法說（老虎證券逐字稿，二手）", "2026-01-15", U_TSMC4Q25, "Interested-party（二手）", "AI 加速器 CAGR mid-to-high 50s"),
    ("台積電 2026Q3 法說（Webull 逐字稿，二手）", "2026-10-15", U_TSMC3Q26, "Interested-party（二手）", "real-world production of 5x reticle size"),
    ("DIGITIMES（付費牆，只讀到摘要）", "2026-02-02", U_DIGI, "二手", "2025 AI 加速器占營收『high double-digit』"),
    ("Tom's Hardware：四大 2026 資本支出（二手）", "2026-04-30", U_TOMS, "Interested-party（二手）", "四大 2026 約 7,250 億、2025 約 4,100 億"),
    ("Universal Asset Owners：Q3 資本支出（二手）", "2026-11", U_UAO, "Interested-party（二手）", "Microsoft 曆年 2026 約 1,700 億"),
    ("Data Center Knowledge：Q2 法說（二手）", "2026-07-30", U_DCK, "Interested-party（二手）", "Alphabet 2026 195–205"),
    ("DZRH：鴻海 AI 機櫃週產能（二手）", "2025-11-21", U_FOX, "Interested-party（二手）", "1,000 artificial intelligence racks per week"),
    ("HBM 獨立供給查證（本庫）", "2026-11-18", U_HBM, "Analogy", "三家公司未揭露位元產能"),
    ("顆數來源口徑查核（本庫）", "2026-11-18", U_UB, "Derived", "兩層都是封裝口徑"),
    ("NVIDIA 技術部落格 Vera Rubin", "2026-03-16", U_BLOG, "Interested-party", "72 NVIDIA Rubin GPUs（unit-basis 報告已讀）"),
    ("NVIDIA Q2 FY26 法說逐字稿（公司 IR PDF）", "2025-08-27", U_NVQ2FY26T, "Interested-party", "we represent about $35 billion"),
    ("NVIDIA 技術部落格 DSX MaxLPS", "2026-08-24", U_DSX, "Interested-party", "fit 40K Rubin GPUs"),
    ("Nscale S-1/A（SEC）", "2026-11-12", U_NSC, "Verified（公司估計）", "approximately $21 billion to $22 billion"),
    ("The Energy Mag 轉述 Bloomberg：安索普 460 MW", "2026-08-26", U_460, "二手", "approximately 460 megawatts"),
    ("IREN 新聞稿（Seeking Alpha 轉載）", "2025-11-03", U_IREN, "Verified", "200MW of critical IT load"),
    ("TweakTown 轉述 CNBC 訪談 Huang", "2025-09-24", U_JH45, "Interested-party（二手）", "4 million or 5 million GPUs"),
    ("Dell Q1 FY27 新聞稿", "2026-05-28", U_DELL1, "Verified", "AI 營收 $16,132M"),
    ("Dell Q2 FY27 新聞稿（01net 轉載）", "2026-09-01", U_DELL2, "Verified／Interested-party", "record $95 billion backlog"),
    ("SMCI FY27 指引（pulse2 轉述）", "2026-08-12", U_SMCI, "Interested-party（二手）", "FY27 $65–72B"),
    ("SMCI 8-K 初估 Q1 FY27", "2026-10-20", U_SMCIQ1, "Interested-party", "$18.5–19.5B"),
    ("Next Platform：HPE Q3 FY26", "2026-09-08", U_HPE, "Verified（二手）", "AI 積壓 $6.8B"),
    ("Nebius Q3 2026 股東信（6-K）", "2026-11", U_NEB, "Interested-party", "sized for more than 1 GW"),
    ("TIKR 轉述 BG2 Podcast Huang", "2026-10-19", U_BG2, "Interested-party（二手）", "two gigawatts ... $80 billion（循環，不獨立）"),
    ("wccftech：Bernstein／Morgan Stanley 機櫃價（券商，對照）", "2026-06-08", U_BERN, "券商（二手）", "未入計算"),
    ("未取得原文：NVIDIA Q2 FY27 10-Q（sec.gov 被擋）", "", "", "", "以 CFO commentary 代替"),
    ("未取得原文：Bloomberg 安索普—Nscale 報導（付費牆）", "", "", "", "用轉載"),
    ("未取得：NVIDIA FY27 Q3 法說逐字稿（樣本外，未搜到全文）", "", "", "", "未入估計"),
    ("未取得原文：GTC 2026 主題演講『1 兆』原話", "", "", "", "只有二手轉述"),
    ("未取得原文：台積電法說官方逐字稿", "", "", "", "用第三方逐字稿"),
    ("未取得原文：鴻海 2026 Q3 法說 AI 機櫃說法", "", "", "", "多個轉載網站抓取失敗"),
    ("公司沒揭露：台積電 CoWoS 片數、NVIDIA 出貨顆數、三家 HBM 位元產能", "", "", "", ""),
]
for s_ in SRC:
    so.append(list(s_))
so.column_dimensions["A"].width = 50
so.column_dimensions["C"].width = 60
so.column_dimensions["E"].width = 50

wb.save(OUT)
print("saved", OUT)
