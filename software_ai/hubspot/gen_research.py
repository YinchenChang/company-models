#!/usr/bin/env python3
"""由 facts_hubspot.py 的資料產生 data/research_hubspot.md（研究包；數字不手改，改 facts_hubspot.py 後重跑）。"""
import os
import facts_hubspot as F

H = F.HIST
Y = F.YRS


def row(name, fmt="{:,.1f}"):
    return "| " + name + " | " + " | ".join(fmt.format(v) for v in H[name]) + " |"


Q = F.Q
L = []
a = L.append
a("# HubSpot（HUBS）研究包（2026-11-16；AI 軟體應用 Project；原型 S：按席位計價）")
a("")
a("> 金額 US$M（原表千美元 ÷ 1,000）；會計年度＝曆年；逐列出處見 `facts_hubspot.csv`（9 欄：指標、數值、單位、期間、口徑、出處、擷取日、標記、來源誘因）。標記：Verified／Interested-party／Analogy／Assumed／Derived。本研究包未做先驗。")
a("")
a("## (1) 一句話公司描述")
a("HubSpot 是以「agentic customer platform」定位的 CRM 軟體公司，向 2–2,000 名員工的中型 B2B 企業銷售行銷、銷售、客服、內容與營運平台（Hubs）加 AI 代理（Breeze、Customer Agent、Prospecting Agent），營收約 98% 為訂閱，按席位與層級（Starter／Professional／Enterprise）計價，AI 代理自 2026-04-14 起改按結果計價（來源：10-K FY2025 Item 1；HubSpot 官網公告）。標記 Verified。")
a("")
a("## (2) 官方分部與固定產品線")
a("- 官方分部：**單一營運分部**（10-K FY2025 附註；主要營運決策者為執行長）。無分部營業利益揭露。")
a("- 官方營收線：Subscription（訂閱）、Professional services and other（專業服務與其他）。地理（Americas／Europe／Asia Pacific）為揭露備註，不是分部。")
a("- 產品線（無分部營收揭露）：Marketing Hub、Sales Hub、Service Hub、Content Hub、Operations Hub、Commerce Hub、Data Hub、Smart CRM、Breeze AI（Customer Agent、Prospecting Agent、Breeze Copilot、Breeze Intelligence）。")
a("")
a("## (3) 分部營收：5 個會計年度＋4 季（含對帳列）")
a("| 項目（US$M） | " + " | ".join(Y) + " |")
a("|---|" + "---|" * 5)
for n in ("訂閱營收", "專業服務與其他營收", "營收合計"):
    a(row(n))
a("| 對帳：訂閱＋專業服務 − 營收合計 | " + " | ".join(f"{H['訂閱營收'][i] + H['專業服務與其他營收'][i] - H['營收合計'][i]:.3f}" for i in range(5)) + " |")
a("")
a("| 項目（US$M） | Q4'25 | Q1'26 | Q2'26 | Q3'26 |")
a("|---|---|---|---|---|")
qs = ["Q4'25", "Q1'26", "Q2'26", "Q3'26"]
for lab, k in (("訂閱營收", "sub"), ("專業服務與其他營收", "ps"), ("營收合計", "rev")):
    a(f"| {lab} | " + " | ".join(f"{Q[q][k]:,.3f}" for q in qs) + " |")
a("| 對帳：訂閱＋專業服務 − 營收合計 | " + " | ".join(f"{Q[q]['sub'] + Q[q]['ps'] - Q[q]['rev']:.3f}" for q in qs) + " |")
a("")
a("地理（Americas／Europe／Asia Pacific，US$M）：FY2023 1,321.6／668.3／180.4；FY2024 1,588.9／825.6／213.1；FY2025 1,872.2／1,011.0／248.1；Q3'26 556.8／309.8／81.0（Verified；10-K 地理資訊附註與 10-Q）。")
a("FY2023 數字取 FY2025 10-K 重編後數（營業損失 −200.9；FY2023 10-K 原數 −208.1，訂閱營業成本 283.7 對 290.8），重編原因未找到。")
a("")
a("## (4) 原型 S（席位計價）的 KPI")
a("| KPI | " + " | ".join(Y) + " | 標記 |")
a("|---|" + "---|" * 6)
for name, vals in F.KPI.items():
    a(f"| {name} | " + " | ".join(f"{v:,}" for v in vals[0]) + " | Interested-party |")
a("| Net Revenue Retention（%；FY2021–22 舊口徑、FY2023–25 新口徑） | 115.2 | 110.3 | 103.0 | 101.8 | 103.5 | Interested-party |")
a("| 全職員工（期末） | 未找到 | 7,433 | 7,663 | 8,246 | 8,882 | Verified |")
a("")
a("近期（Interested-party，公司自訂；來源：Q3 2026 新聞稿與法說摘要）：Customers 288,706（Q4'25）→ 299,458 → 306,446 → 311,704（Q3'26）；ASRPC 11,683 → 11,722 → 11,800 → 12,038 美元；NRR 約 101%（Q3'26）；客戶淨增 Q2 7.0k、Q3 5.3k、Q4 預期低於 5k；客戶金額留存「高 80% 區間」。**席位數、Pro 以上客戶占比：未找到。**")
a("來源誘因：KPI 為公司自訂指標且定義可變更（Customers 自 2026Q3 起納入 YouSpot Pro 訂閱者；NRR 2025 年起改口徑）。")
a("")
a("## (5) AI 相關揭露")
for t in [
    "價格（Verified，HubSpot 官網）：Customer Agent 自 2026-04-14 起每次解決對話 0.50 美元（50 Credits；原為每次對話 1.00 美元）；Prospecting Agent 每筆推薦線索 1.00 美元（100 Credits）；每 Credit 約 0.01 美元；適用 Pro 與 Enterprise 客戶。",
    "用量（Interested-party，法說口頭說法，第二來源轉述）：Customer Agent Q3 2026 解決超過 200 萬件（以 0.50 美元計約 1 百萬美元／季，Derived）；Pro 以上客戶 20% 使用 agents（+11 個百分點）、64% 使用 AI 功能；agent 動作 5 倍、credits 消耗約 2.5 倍（基期未載）；自訂代理 5,600 個。",
    "定價模式（Interested-party）：seats and credits 新方案僅新客戶；北歐 2026-04、荷比盧 2026-06、EMEA 全部新客戶 Q4 2026；試點區初始 ASP 較低、credits 用量 2–3 倍；2027-02 公布 EMEA 結果。",
    "成本（Verified，10-Q）：訂閱營業成本含「AI inferencing costs」；Q3 訂閱與託管成本年增 23.2 百萬美元，部分因 AI 功能使用；非 GAAP 訂閱毛利率 85.4%（去年同期 86.7%）；公司預期毛利率隨時間小幅下降。",
    "管理層說法（Interested-party）：AI 改變客戶發現、評估與購買軟體的方式；Q2 逆風在 Q3 更明顯，預期延續到 2027。",
    "**未找到**：AI 專屬營收或 ARR、credits 總營收、每任務 token 量、席位被代理取代的量化證據。每任務 token 因此以 Tokenomics 工作負載為類比（Analogy），不標 Verified（工作單 G2.1）。",
]:
    a("- " + t)
a("")
a("## (6) 利潤率、人力與股票薪酬（US$M）")
a("| 項目 | " + " | ".join(Y) + " |")
a("|---|" + "---|" * 5)
for n in ("營業利益（GAAP）", "股票薪酬", "折舊與攤銷", "營業現金流"):
    a(row(n))
a("| 非 GAAP 營業利益（公司口徑） | 未找到 | 169.1 | 330.3 | 460.2 | 581.9 |")
a("| 非 GAAP 營業利益率 | 未找到 | 9.8% | 15.2% | 17.5% | 18.6% |")
a("| SBC 占營收 | " + " | ".join(f"{H['股票薪酬'][i] / H['營收合計'][i]:.1%}" for i in range(5)) + " |")
a("| GAAP 營業利益率 | " + " | ".join(f"{H['營業利益（GAAP）'][i] / H['營收合計'][i]:.1%}" for i in range(5)) + " |")
a("")
a("近四季（Q4'25＋9M26）：營收 3,587.0、非 GAAP 營業利益 747.3（20.8%）、SBC 503.7（14.0%）、GAAP 營業利益 191.4（5.3%）；Q3'26 非 GAAP 營業利益率 22.6%。員工 8,882（2025 年底）；2026-10-01 重組計畫約 7% 員工（約 660 人），費用 65–75 百萬美元，多數 Q4 認列（Verified，10-Q）。")
a("**歷史規模經濟（工作單 G3.1）**：非 GAAP 營業利益率逐年上升、SBC 占營收自 FY2023 起下降；模型以 FY2023→FY2025 增量利益率延伸，不凍結費用率，全期持平為 Sens 下緣。")
a("")
a("## (7) 市場共識（Interested-party；賣方有評等與投行業務關係）")
a("| 項目 | 數值 | 分析師數 | 日期 | 來源 |")
a("|---|---|---|---|---|")
a("| 平均目標價 | 243.18（低 190／高 300） | 49 | 2026-11-16 | TickFlow |")
a("| 平均目標價（對照） | 245.25 | 47 | 約 2026-11-13 | TickerGate |")
a("| FY2026 營收共識 | 3,690 | 32 | Q3 後（頁面無日期） | Twelve Data |")
a("| FY2027 營收共識 | 4,100（區間 3,980–4,220） | 32 | Q3 後（頁面無日期） | Twelve Data |")
a("| FY2027、FY2028 EPS 共識（Q3 後） | **未找到** | — | — | 取得頁面為財報前（FY2027 16.85）或欄位矛盾，不採用 |")
a("| FY2028 營收共識（Q3 後） | **未找到**（財報前 4,769） | — | — | MarketScreener 2026-10-12，不採用 |")
a("現價 217.25（2026-11-13 收盤，QuiverQuant 與 ChartRow 同值）；前瞻 P/E 13.2 倍；YTD 約 −46%。")
a("")
a("## (8) 多空各 3 點（含出處）")
a("**多方**")
for t in ["62% 的新 Pro 以上客戶使用一個以上 Hub，agent 採用率（Pro 以上 20%）與 credits 消耗（約 2.5 倍）快速上升（Q3 法說摘要，PPC Land；Interested-party）。",
          "非 GAAP 營業利益率九個月自 17.1% 升至 20.3%、Q3 22.6%；管理層目標 2027 年 23–24%、2030 年 30%；自由現金流約占營收 20%（Q3 新聞稿；法說）。",
          "前瞻 P/E 約 13 倍，低於同業中位數 15.4 倍；Q3 回購 230.8 百萬美元、授權剩餘約 964 百萬美元，淨現金 1,310.5 百萬美元無債務（Q3 10-Q）。"]:
    a("- " + t)
a("**空方**")
for t in ["管理層：「Q2 的逆風在 Q3 更明顯」，Q4 營收指引 945–947 低於 Q3 的 947.6，成長自 17% 降至 12%；AI 改變客戶發現與購買軟體的方式（Q3 新聞稿、法說摘要）。",
          "NRR 約 101%（FY2026 預期較 FY2025 降約 2 個百分點）、客戶淨增 Q4 低於 5k；seats and credits 新方案初始 ASP 較低，credits 可能替代席位支出（法說摘要；Interested-party）。",
          "AI 代理按任務計價單價已由 1.00 降至 0.50 美元；席位被取代的侵蝕在本模型情境 B 下遠大於按任務計價營收（模型 Thesis 頁；Derived）；賣方 35 位中 22 位 Hold（StockAnalysis 2026-10-23）。"]:
    a("- " + t)
a("")
a("## (9) 未找到清單")
for t in F.NOT_FOUND:
    a("- " + t)
a("")
a("## 附：工作單 G2／G3 落實對照（詳見 README.md）")
a("見 `README.md` 的「工作單 G2／G3 逐條對照」。")
open(os.path.join(F.HERE, "data", "research_hubspot.md"), "w", encoding="utf-8").write("\n".join(L) + "\n")
print("research_hubspot.md", len(L), "行")
