#!/usr/bin/env python3
"""AI 對軟體單位經濟 v0.1：產生 xlsx（openpyxl 寫公式），以 LibreOffice headless 重算後檢查。

用法：
  python3 software_ai/unit_economics/build_unit_economics.py            # 產生 dist/<日期>_AI對軟體單位經濟_v0.1.xlsx 並重算、檢查
  python3 software_ai/unit_economics/build_unit_economics.py --no-recalc

命題：每單位營收的 AI 淨效果＝AI 帶來的營收增量 − 定價侵蝕 − 推論成本（三者並列，逐年 2026–2032）。
推論成本只讀 Tokenomics 快照 data/tk_snapshot.json（tools/tokenomics/import_tokenomics.py 產生）；
本程式不寫任何 token 價格，只把快照值抄進 In 頁並註明名稱。其他參數全部在 In 頁（藍字黃底，低／基準／高＋標記）。
"""
import argparse, datetime, json, os, re, shutil, subprocess, sys
from pathlib import Path

import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

HERE = Path(__file__).resolve().parent
REPO = HERE.parents[1]
SNAP = HERE / 'data' / 'tk_snapshot.json'
TOOL = REPO / 'tools' / 'tokenomics' / 'import_tokenomics.py'
VERSION = 'v0.1'
YEARS = list(range(2026, 2033))
YC0 = 4  # 年度欄起點（D）
SENS_YEAR = 2030

BLUE = Font(color='0000FF')
BLACK = Font(color='000000')
BOLD = Font(bold=True)
HFONT = Font(bold=True, color='FFFFFF')
HFILL = PatternFill('solid', fgColor='1F4E78')
YFILL = PatternFill('solid', fgColor='FFF2CC')
SFILL = PatternFill('solid', fgColor='DDEBF7')
KFILL = PatternFill('solid', fgColor='E2EFDA')
WRAP = Alignment(wrap_text=True, vertical='top')

TIERS = ['Luna', 'Terra', 'Sol', 'Astra']          # 層級序號 1–4（單價由低到高）
TASKS = ['一般聊天', '推理聊天', '單代理（工具迴圈）', '多代理研究', 'Coding agent（長程）']  # Tokenomics Workload 任務 1–5（IF_HdrTask 文字，IF_Hdr* 不得連結，故抄錄）

# ---------------------------------------------------------------- 參數（In 頁）
# (key, 說明, 單位, 低, 基準, 高, 標記, 依據, 格式)；None 開頭＝節標題
P = []


def sec(title):
    P.append((None, title))


def par(k, lab, unit, lo, base, hi, flag, basis, fmt='0.00'):
    P.append((k, lab, unit, lo, base, hi, flag, basis, fmt))


PCT = '0.0%'
USD = '#,##0.00'
INT = '0'

sec('共用：推論成本路徑（Tokenomics 只給 2026 快照；以下三項為本模型的路徑假設）')
par('g_px', '模型 API 市場單價年變化（同一層級、同一 token 類型）｜待 AI供應鏈 登錄', '%/年', -0.50, -0.25, 0.00, 'Assumed',
    'Tokenomics v22.0 契約無「未來市場單價路徑」名稱，依工作單寫「待 AI供應鏈 登錄」並以區間作敏感度。下限參考：Epoch AI 量測同一能力水準的推論價格每年下降 9–900 倍（S01），但那是「同能力」口徑，同一層級名目單價的下降遠慢於此；上限＝層級牌價不變。', PCT)
par('g_tok', '每任務 token 數年變化（同一任務，代理迴圈加長、思考加深）', '%/年', 0.00, 0.20, 0.50, 'Assumed',
    '方向證據：Anthropic 公開的 Claude Code 每開發者每日平均成本估計由 6 美元上調為 13 美元（S04、S05），代表同一工作的 token 用量上升；幅度無直接量測。', PCT)
par('ovh', '實際工作負載 token ÷ Tokenomics 參考工作負載（重試、驗證、長上下文）', 'x', 1.0, 3.0, 8.0, 'Analogy',
    '推導：Tokenomics Coding agent（8 人時）Sol 層每成功任務市場價約 1.3 美元（In 頁 T 節）；企業部署 Claude Code 平均每開發者每活躍日約 13 美元、九成使用者低於 30 美元（S04）。牌價差約 1.9 倍（Sonnet 3/15 對 Sol 1.6/8 美元每百萬 token，S04 範例），一個開發日不等於 8 人時的代理完成量，故 token 倍數約 2–5，區間放寬為 1–8。', '0.0')
par('gap_vr', 'GB300→VR200 世代上市間隔 Δ（只用於全成本年變化）', '年', 1.11, 1.18, 1.24, 'Analogy',
    'Tokenomics 下游契約 v0.28 第 4 節文字（CAP_GenGap_VR200 基準／低／高 1.18／1.11／1.24；CAP_ 前綴禁止連結，依契約自條文抄錄）。', '0.00')

sec('L 按人力計價（IT 服務、BPO）；計算單位＝AI 前 1 名交付人員一年的工作範圍')
par('L_util', '利用率（計費工時 ÷ 可用工時）', '%', 0.72, 0.80, 0.88, 'Assumed', '產業慣用區間；樣板公司實績待研究包校準。', PCT)
par('L_hrs', '每人年可用工時', '小時/年', 1800, 1880, 1960, 'Assumed', '約 47 週 × 40 小時；區間反映各國假期差異。', '#,##0')
par('L_rate', '混合計費時薪（境外＋在地）', '$/小時', 30, 50, 90, 'Assumed', '境外交付比重高者近下限，在地顧問比重高者近上限。單位營收的絕對值只影響推論成本占比。', USD)
par('L_gm0', 'AI 前毛利率（交付人力成本以外視為 0）', '%', 0.28, 0.32, 0.36, 'Assumed', 'IT 服務毛利率常見約三成；樣板公司實績待研究包。', PCT)
par('L_a26', '任務自動化比例（占計費工時）— 2026', '%', 0.03, 0.08, 0.15, 'Assumed',
    '田野實驗：軟體開發者使用 AI 助手完成任務數 +26%（S06）、客服每小時解決件數 +14%（S07）；生產力 +x 等於節省 x/(1+x) 的工時，乘以可覆蓋工作比例。', PCT)
par('L_a32', '任務自動化比例（占計費工時）— 2032', '%', 0.15, 0.35, 0.60, 'Assumed', '代理可獨立完成的任務長度持續拉長（Tokenomics IF_TaskLen 最長 8 人時），上限假設六成工時可由代理承擔。', PCT)
par('L_r', '價格保留率（被自動化工時中，廠商仍向客戶收取的比例）', '%', 0.10, 0.35, 0.70, 'Assumed', '0＝生產力全數讓給客戶（按工時計價）；1＝完全改按成果計價。無直接量測，區間寬。', PCT)
par('L_n26', 'AI 導入新專案營收（占 AI 前單位營收）— 2026', '%', 0.00, 0.02, 0.05, 'Assumed', '資料整理、代理導入、治理等新工作；樣板公司揭露待研究包（不得以公司指引反推）。', PCT)
par('L_n32', 'AI 導入新專案營收（占 AI 前單位營收）— 2032', '%', 0.03, 0.10, 0.25, 'Assumed', '同上。', PCT)
par('L_tier', '代理使用的模型層級（1 Luna／2 Terra／3 Sol／4 Astra）', '序號', 2, 3, 4, 'Assumed', '交付程式碼與系統整合需中高層級；任務固定為 Tokenomics 任務 5「Coding agent（長程）」。', INT)

sec('S 按席位計價（CRM、設計與生產力、人資、工作流）；計算單位＝AI 前 1 個付費席位一年')
par('S_p', '每席位年費', '$/席位/年', 600, 1500, 3000, 'Assumed', '承重參數，決定代理營收相對席位營收的比例（每任務營收為絕對金額，席位年費愈低，代理營收占 R0 愈高）。區間取企業 SaaS 每使用者每月約 50–250 美元。', '#,##0')
par('S_gm0', 'AI 前毛利率', '%', 0.72, 0.78, 0.84, 'Assumed', 'SaaS 常見毛利率區間；樣板公司實績待研究包。', PCT)
par('S_h26', '客戶人力替代率（持有席位的職務被代理取代的比例）— 2026', '%', 0.00, 0.01, 0.03, 'Assumed', '與 L_a26 同一組田野實驗證據（S06、S07），但席位需人員離職或不補人才會減少，故低於自動化比例。', PCT)
par('S_h32', '客戶人力替代率 — 2032', '%', 0.05, 0.15, 0.35, 'Assumed', '同上。', PCT)
par('S_e', '席位彈性（人力每減 1%，付費席位減少 %）', 'x', 0.5, 0.8, 1.0, 'Assumed', '席位綁人頭時接近 1；合約最低席位數、全員授權方案使其低於 1。', '0.00')
par('S_V', '每席位每年可交給代理的任務數', '件/席位/年', 500, 1500, 3000, 'Assumed', '約每工作日 2–12 件（250 工作日），任務長度取 Tokenomics 任務 3「單代理」的 IF_TaskLen。', '#,##0')
par('S_ad26', '代理採用率（可交付任務中實際由代理處理的比例）— 2026', '%', 0.003, 0.01, 0.03, 'Assumed', '早期付費部署，多數客戶仍在試行；樣板公司付費交易數待研究包。', PCT)
par('S_ad32', '代理採用率 — 2032', '%', 0.10, 0.30, 0.60, 'Assumed', '同上。', PCT)
par('S_wage', '客戶人員每小時全負擔人力成本', '$/小時', 25, 40, 70, 'Analogy', '美國客服代表時薪中位數約 20 美元（BLS OEWS 43-4051，S08），全負擔約 1.3–1.6 倍；銷售、人資等職務較高。', USD)
par('S_cap', '廠商取價率（每任務價格 ÷ 該任務人力成本）', '%', 0.03, 0.10, 0.25, 'Assumed', '每任務價格＝時薪 × 任務長度 × 取價率，由價值面推導，不取公司牌價（牌價列入可證偽指標）。', PCT)
par('S_gpt', '每任務價格年變化', '%/年', -0.25, -0.10, 0.00, 'Assumed', '競爭與模型商直接提供同類代理會壓低每任務價格。', PCT)
par('S_tier', '代理使用的模型層級（1–4）', '序號', 1, 2, 3, 'Assumed', '客服、行政類工作多用中低層級；任務固定為 Tokenomics 任務 3「單代理（工具迴圈）」。', INT)

sec('A 按廣告計價（數位廣告）；計算單位＝AI 前每 1,000 次查詢')
par('A_load', '每次查詢廣告數（廣告負載）', '則/查詢', 0.6, 1.0, 1.6, 'Assumed', '搜尋結果頁頂部與底部廣告；入口網站與通訊平台的展示型廣告負載不同。', '0.00')
par('A_ctr', '廣告點擊率', '%', 0.015, 0.03, 0.05, 'Assumed', '搜尋廣告常見區間。', PCT)
par('A_cpc', '每次點擊價格', '$/點擊', 0.30, 0.80, 1.50, 'Assumed', '日本、台灣等市場低於美國；只影響推論成本占比。', USD)
par('A_gm0', 'AI 前毛利率（扣流量獲取成本後）', '%', 0.45, 0.58, 0.70, 'Assumed', '樣板公司實績待研究包。', PCT)
par('A_d26', '聊天介面分流比例（原查詢改在聊天介面完成）— 2026', '%', 0.03, 0.06, 0.12, 'Analogy',
    'Ahrefs 估 ChatGPT 的「傳統搜尋型」提示量約為 Google 搜尋量的 12%（嚴格口徑 4.4%–9.4%，2026-02，S03）；量的比例不等於被分流的比例，故基準取一半。', PCT)
par('A_d32', '聊天介面分流比例 — 2032', '%', 0.08, 0.20, 0.40, 'Assumed', '延伸 S03 的成長方向；無直接預測證據。', PCT)
par('A_u26', 'AI 帶來的轉換率提升（反映在每次點擊價格）— 2026', '%', 0.00, 0.02, 0.05, 'Assumed', '生成式廣告素材與出價自動化；樣板公司揭露待研究包。', PCT)
par('A_u32', 'AI 帶來的轉換率提升 — 2032', '%', 0.02, 0.08, 0.20, 'Assumed', '同上。', PCT)
par('A_q26', '保留查詢中附 AI 生成回答的比例 — 2026', '%', 0.15, 0.30, 0.60, 'Assumed', '搜尋引擎的 AI 摘要覆蓋率；各市場差異大。', PCT)
par('A_q32', '保留查詢中附 AI 生成回答的比例 — 2032', '%', 0.40, 0.70, 1.00, 'Assumed', '同上。', PCT)
par('A_tier', '生成回答使用的模型層級（1–4）', '序號', 1, 1, 2, 'Assumed', '每次查詢成本須極低，多用最低層級；任務固定為 Tokenomics 任務 1「一般聊天」。', INT)

sec('U 按用量計價（資料平台）；計算單位＝AI 前每 1 美元運算額度營收')
par('U_px', '每運算單位（credit）單價', '$/單位', 2.0, 3.0, 4.0, 'Assumed', '只用於恆等式「運算單位 × 單價＝1 美元」；結果不受此值影響。', USD)
par('U_gm0', 'AI 前產品毛利率（雲端基礎設施成本）', '%', 0.60, 0.70, 0.78, 'Assumed', '樣板公司實績待研究包。', PCT)
par('U_w26', '代理工作負載乘數（查詢量、資料量增加，占 AI 前用量）— 2026', '%', 0.01, 0.04, 0.08, 'Assumed', '代理自動產生的查詢；無直接量測。', PCT)
par('U_w32', '代理工作負載乘數 — 2032', '%', 0.10, 0.40, 0.90, 'Assumed', '同上。', PCT)
par('U_e26', '工作負載被替代（模型直接讀原始資料、開放格式外移）— 2026', '%', 0.00, 0.01, 0.03, 'Assumed', '無直接量測。', PCT)
par('U_e32', '工作負載被替代 — 2032', '%', 0.00, 0.08, 0.20, 'Assumed', '同上。', PCT)
par('U_j26', 'AI 功能呼叫數（向量檢索、內建模型；每 1 美元 AI 前用量）— 2026', '次/$', 0.1, 0.5, 1.0, 'Assumed', '任務固定為 Tokenomics 任務 2「推理聊天」。', '0.00')
par('U_j32', 'AI 功能呼叫數 — 2032', '次/$', 1.0, 4.0, 10.0, 'Assumed', '同上。', '0.00')
par('U_mk', 'AI 功能轉售加成（售價 ÷ 推論成本 − 1）', '%', 0.00, 0.30, 1.00, 'Assumed', '成本加成計價（credit 轉售）；0＝原價轉嫁。', PCT)
par('U_tier', 'AI 功能使用的模型層級（1–4）', '序號', 1, 2, 3, 'Assumed', '平台多提供多家模型，預設中低層級。', INT)

sec('P 按防護範圍計價（資安）；計算單位＝AI 前 1 個受保護身分或端點一年')
par('P_mod', '每受保護單位模組數', '個', 2, 4, 7, 'Assumed', '平台型資安廠商以模組數擴張；樣板公司模組採用分布待研究包。', '0.0')
par('P_px', '每模組每受保護單位年價', '$/年', 8, 20, 40, 'Assumed', '只影響推論成本占比的比例尺。', USD)
par('P_gm0', 'AI 前毛利率', '%', 0.70, 0.77, 0.82, 'Assumed', '樣板公司實績待研究包。', PCT)
par('P_n26', '受保護代理（非人類）身分數，每 1 個 AI 前單位 — 2026', '個', 0.02, 0.05, 0.15, 'Assumed', 'CyberArk 調查：機器身分對人類身分約 82:1（2025，S02），但絕大多數未被單獨計價；此處只計被納入計價的代理身分。', '0.00')
par('P_n32', '受保護代理身分數 — 2032', '個', 0.20, 0.80, 3.00, 'Assumed', '同上；代理數量隨企業部署成長。', '0.00')
par('P_rho', '代理身分單價 ÷ AI 前單位單價', 'x', 0.02, 0.10, 0.30, 'Assumed', '代理身分以批量或用量計價，單價遠低於人類端點。', '0.00')
par('P_m26', '攻擊面擴大帶來的模組加購（占 AI 前單位營收）— 2026', '%', 0.00, 0.02, 0.05, 'Assumed', '無直接量測。', PCT)
par('P_m32', '模組加購 — 2032', '%', 0.03, 0.10, 0.25, 'Assumed', '同上。', PCT)
par('P_e26', 'AI 原生競爭造成的價格侵蝕 — 2026', '%', 0.00, 0.01, 0.03, 'Assumed', '無直接量測。', PCT)
par('P_e32', 'AI 原生競爭造成的價格侵蝕 — 2032', '%', 0.00, 0.05, 0.15, 'Assumed', '同上。', PCT)
par('P_k26', 'AI 資安分析次數（每受保護單位每年）— 2026', '次/年', 1, 4, 10, 'Assumed', '告警分流、調查；任務固定為 Tokenomics 任務 3「單代理（工具迴圈）」。', '#,##0')
par('P_k32', 'AI 資安分析次數 — 2032', '次/年', 10, 40, 150, 'Assumed', '同上。', '#,##0')
par('P_tier', '資安分析使用的模型層級（1–4）', '序號', 1, 2, 3, 'Assumed', '大量告警用中低層級。', INT)

# ---------------------------------------------------------------- 次產業（營收池 × 滲透率 × 原型單位淨效果）
SI = [  # (代碼, 次產業, 原型, 營收池口徑)
    ('IT', 'IT 服務', 'L', 'Gartner「IT Services」終端支出'),
    ('CRM', 'CRM', 'S', 'CRM 應用軟體營收'),
    ('AD', '數位廣告', 'A', '純數位媒體廣告營收（WPP Media pure-play digital）'),
    ('DATA', '資料平台', 'U', '資料庫管理系統（DBMS）營收'),
    ('SEC', '資安', 'P', 'Gartner「Security Software」終端支出'),
]
sec('次產業：2025 全球營收池（$B）、基礎成長（不含 AI）、滲透率（營收池中適用該原型單位經濟的比例；線性 2026→2032）')
par('SI_IT_pool', 'IT 服務 2025 全球營收池', '$B', 1550, 1686, 1750, 'Interested-party',
    'Gartner 2025-07-15：2025 IT Services 支出 1,686,321 百萬美元、年增 4.4%（S10）。誘因：Gartner 為產業研究與顧問商，預測影響其客戶的預算判斷；低值反映各家定義較窄。', '#,##0')
par('SI_IT_g', 'IT 服務 基礎成長（不含 AI）', '%/年', 0.02, 0.04, 0.06, 'Assumed', 'S10 的 2024、2025 年增 4.8%、4.4% 已含部分 AI 效果，基礎成長取略低。', PCT)
par('SI_IT_pen26', 'IT 服務 滲透率 — 2026（按工時／人天計價的營收占比）', '%', 0.40, 0.60, 0.80, 'Assumed', '固定價格與託管服務合約不完全按工時計價，但定價仍以人力投入估算。', PCT)
par('SI_IT_pen32', 'IT 服務 滲透率 — 2032', '%', 0.50, 0.70, 0.90, 'Assumed', '同上。', PCT)
par('SI_CRM_pool', 'CRM 2025 全球營收池', '$B', 80, 85, 110, 'Derived',
    'Apps Run The World 2025-07-23：2024 年 CRM 應用營收 80,007 百萬美元（年增 10.5%），預估 2029 年 1,061 億（CAGR 5.8%）（S11）；2025＝80.0 × 1.058。Interested-party 來源（市場研究商）；高值反映較廣的 CRM 定義。', '#,##0')
par('SI_CRM_g', 'CRM 基礎成長（不含 AI）', '%/年', 0.04, 0.06, 0.09, 'Assumed', 'S11 的 2024–2029 CAGR 5.8%。', PCT)
par('SI_CRM_pen26', 'CRM 滲透率 — 2026（按席位計價的營收占比）', '%', 0.60, 0.75, 0.90, 'Assumed', '席位制仍為主要計價；專業服務與用量計價不在內。', PCT)
par('SI_CRM_pen32', 'CRM 滲透率 — 2032', '%', 0.60, 0.80, 0.95, 'Assumed', '同上。', PCT)
par('SI_AD_pool', '數位廣告 2025 全球營收池', '$B', 700, 791, 880, 'Derived',
    'WPP Media 2025 年中預測：全球廣告營收 1.08 兆美元、純數位占 73.2%（含數位延伸 81.6%）、搜尋廣告 2,262 億（年增 7.3%）（S12）；1,080 × 73.2%＝791。Interested-party（媒體代理商）。', '#,##0')
par('SI_AD_g', '數位廣告 基礎成長（不含 AI）', '%/年', 0.04, 0.07, 0.09, 'Assumed', 'S12 全球廣告 2025 年增 6.0%、搜尋 7.3%。', PCT)
par('SI_AD_pen26', '數位廣告 滲透率 — 2026（查詢型／意圖型廣告占比）', '%', 0.25, 0.35, 0.50, 'Assumed',
    '錨點（Derived）：搜尋廣告 2,262 ÷ 純數位 7,910＝29%（S12）；判斷加成（Assumed）：電商站內搜尋與入口網站查詢約再加 6 個百分點，基準 35%。分流主要影響查詢型廣告。', PCT)
par('SI_AD_pen32', '數位廣告 滲透率 — 2032', '%', 0.30, 0.45, 0.65, 'Assumed', '聊天介面分流擴及社群與內容推薦。', PCT)
par('SI_DATA_pool', '資料平台 2025 全球營收池', '$B', 120, 129, 150, 'Derived',
    'Gartner DBMS 2023 年營收 1,032 億美元（IT Market Strategy 轉述，S13）；以 SI_DATA_g 基準 12%／年推到 2025＝103.2 × 1.12²＝129（與基礎成長同口徑；2023–2025 實際成長含 AI 需求，可能更高，故高值 150≈103.2 × 1.20²）。', '#,##0')
par('SI_DATA_g', '資料平台 基礎成長（不含 AI）', '%/年', 0.08, 0.12, 0.16, 'Assumed', '雲端 DBMS 為主要成長來源。', PCT)
par('SI_DATA_pen26', '資料平台 滲透率 — 2026（按用量計價的營收占比）', '%', 0.40, 0.55, 0.70, 'Assumed', '雲端 DBMS 多為用量計價；授權與維護不在內。', PCT)
par('SI_DATA_pen32', '資料平台 滲透率 — 2032', '%', 0.55, 0.70, 0.85, 'Assumed', '同上。', PCT)
par('SI_SEC_pool', '資安 2025 全球營收池', '$B', 100, 106, 115, 'Interested-party',
    'Gartner 2025-07-29：2025 資安終端支出 2,130 億美元，其中 Security Software 1,059 億（2024 為 950 億）（S14）。誘因同 S10。', '#,##0')
par('SI_SEC_g', '資安 基礎成長（不含 AI）', '%/年', 0.07, 0.10, 0.13, 'Assumed', 'S14 Security Software 2025 年增約 11.6%（由表計算）。', PCT)
par('SI_SEC_pen26', '資安 滲透率 — 2026（按端點／身分計價的營收占比）', '%', 0.40, 0.60, 0.75, 'Assumed', '端點、身分、雲端工作負載以受保護數計價；網路設備軟體與服務不在內。', PCT)
par('SI_SEC_pen32', '資安 滲透率 — 2032', '%', 0.50, 0.70, 0.85, 'Assumed', '同上。', PCT)

PKEYS = [p[0] for p in P if p[0]]
PDEF = {p[0]: p for p in P if p[0]}

SOURCES = [
    ('S01', 'Epoch AI，LLM inference price trends（2025-03-12）', 'https://epoch.ai/data-insights/llm-inference-price-trends',
     '同一能力水準的推論價格每年下降 9–900 倍（例：GPQA 上 GPT-4 等級約每年 40 倍）', 'g_px 下限方向', 'Verified（第三方研究）'),
    ('S02', 'GovInfoSecurity 報導 CyberArk 2025 State of Machine Identity Security Report（2025-07-15）',
     'https://www.govinfosecurity.com/cyberark-rise-in-machine-identities-poses-new-risks-a-28967', '機器身分對人類身分約 82:1（調查 1,200+ 位資安主管）', 'P_n26／P_n32 量級', 'Interested-party（資安廠商調查；誘因：銷售身分安全產品）'),
    ('S03', 'Ahrefs，ChatGPT Has 12% of Google’s Search Volume（2026-02-11）', 'https://ahrefs.com/blog/chatgpt-has-12-percent-of-googles-search-volume/',
     'ChatGPT 傳統搜尋型提示約為 Google 搜尋量 12%；依 OpenAI／Harvard 分類為 4.4%–9.4%', 'A_d26', 'Interested-party（SEO 工具商；誘因：搜尋流量議題的關注度）'),
    ('S04', 'Claude Code 文件 Manage costs effectively（擷取 2026-11-16）', 'https://code.claude.com/docs/en/costs',
     '企業部署平均每開發者每活躍日約 13 美元、每月 150–250 美元；九成使用者每活躍日低於 30 美元', 'ovh、g_tok', 'Interested-party（模型商；誘因：影響客戶預算預期）'),
    ('S05', 'Yahoo Finance 報導：Anthropic 將 Claude Code 每開發者成本估計加倍', 'https://finance.yahoo.com/sectors/technology/articles/anthropic-quietly-doubles-estimate-much-220101627.html',
     '每日平均成本估計由 6 美元上調（與 S04 現值 13 美元一致）', 'g_tok 方向', 'Interested-party（媒體轉述模型商自身估計；誘因同 S04）'),
    ('S06', 'Cui, Demirer, Jaffe, Musolff, Peng, Salz（2024）The Effects of Generative AI on High-Skilled Work: Evidence from Three Field Experiments with Software Developers',
     'https://papers.ssrn.com/sol3/papers.cfm?abstract_id=4945566', '使用 AI 程式助手的開發者完成任務數增加約 26%', 'L_a26、S_h26', 'Verified（學術田野實驗）'),
    ('S07', 'Brynjolfsson, Li, Raymond（2023／2025）Generative AI at Work（NBER w31161）', 'https://www.nber.org/papers/w31161',
     '客服人員導入生成式 AI 助手後每小時解決件數平均增加約 14%', 'L_a26、S_h26', 'Verified（學術田野研究）'),
    ('S08', 'BLS OEWS 43-4051 Customer Service Representatives', 'https://www.bls.gov/oes/current/oes434051.htm',
     '美國客服代表時薪中位數約 20 美元（全負擔倍數另為 Assumed）', 'S_wage', 'Verified（官方統計）'),
    ('S09', 'Tokenomics 下游契約 v0.28（YinchenChang/Tokenomics docs/plan/Tokenomics_downstream_contract.md）', 'repo 內文件',
     '取數規則（只取 IF_／IFC_）、CAP_GenGap_VR200 區間文字、IFC_Use 用途限制', '全部推論成本', '契約'),
    ('S10', 'Gartner 新聞稿 Gartner Forecasts Worldwide IT Spending to Grow 7.9% in 2025（2025-07-15）',
     'https://www.gartner.com/en/newsroom/press-releases/2025-07-15-gartner-forecasts-worldwide-it-spending-to-grow-7-point-9-percent-in-2025',
     'IT Services 2025 支出 1,686,321 百萬美元（年增 4.4%）；2024 年增 4.8%', 'SI_IT_pool、SI_IT_g', 'Interested-party（研究顧問商；誘因：預測支撐其客戶服務）'),
    ('S11', 'Apps Run The World，Top 10 CRM Software Vendors and Market Forecast 2024-2029（2025-07-23）',
     'https://www.appsruntheworld.com/top-10-crm-software-vendors-and-market-forecast', '2024 CRM 應用營收 80,007 百萬美元（年增 10.5%）；2029 年 1,061 億（CAGR 5.8%）',
     'SI_CRM_pool、SI_CRM_g', 'Interested-party（市場研究商；誘因：銷售研究報告）'),
    ('S12', 'WPP Media This Year Next Year 2025 年中預測（LBB Online 轉述）',
     'https://lbbonline.com/news/Mid-Year-Global-Advertising-Forecast-Update-dollar108-Trillion-in-2025-Ad-Revenue-and-6percent-Growth',
     '2025 全球廣告營收 1.08 兆美元（年增 6.0%）；純數位 73.2%；搜尋 2,262 億（年增 7.3%）', 'SI_AD_pool、SI_AD_g、SI_AD_pen26', 'Interested-party（媒體代理商；誘因：廣告預算規劃）'),
    ('S13', 'IT Market Strategy，DBMS Revenues 2023 – Beyond the Top Ten（2024-06-05；轉述 Gartner）',
     'https://itmarketstrategy.com/2024/06/05/dbms-revenues-2023-beyond-the-top-ten/', '2023 年 DBMS 市場營收 1,032 億美元', 'SI_DATA_pool', 'Interested-party（產業分析師部落格轉述 Gartner）'),
    ('S14', 'Gartner 新聞稿 End-User Spending on Information Security to Total $213 Billion in 2025（2025-07-29）',
     'https://www.gartner.com/en/newsroom/press-releases/2025-07-29-gartner-forecasts-worldwide-end-user-spending-on-information-security-to-total-213-billion-us-dollars-in-2025',
     '2025 資安支出 213,025 百萬美元；Security Software 105,940 百萬美元（2024 為 94,960）', 'SI_SEC_pool、SI_SEC_g', 'Interested-party（研究顧問商）'),
]

# ---------------------------------------------------------------- 原型列定義
# 列＝(key, 說明, 單位, 公式樣板, 格式, 類別)；樣板代號：{p:x} 參數基準格；{r:x} 同欄的列；{t} 年序（2026＝0）；{x:NAME} 固定範圍


def lin(a, b):
    return f'({{p:{a}}}+({{p:{b}}}-{{p:{a}}})*{{t}}/6)'


def common_rows():
    return [
        ('t', '年序（2026＝0）', '年', None, '0', 'aux'),
        ('pxidx', '推論成本指數 — 市場價（＝(1+g_px)^t ×(1+g_tok)^t × ovh）', 'x', '(1+{p:g_px})^{t}*(1+{p:g_tok})^{t}*{p:ovh}', '0.000', 'aux'),
        ('fcidx', '推論成本指數 — 全成本（＝(1+g_cost)^t ×(1+g_tok)^t × ovh）', 'x', '(1+{p:g_cost})^{t}*(1+{p:g_tok})^{t}*{p:ovh}', '0.000', 'aux'),
    ]


def tail_rows(be_label, be_unit, be_expr, be_fmt):
    return [
        ('rev', '單位營收（＝R0＋增量−侵蝕）', '$', '{r:R0}+{r:incr}-{r:eros}', USD, 'out'),
        ('cogs', '單位營收成本（＝非推論成本＋推論成本）', '$', '{r:cogsx}+{r:inf}', USD, 'out'),
        ('gp', '單位毛利（＝單位營收−單位營收成本）', '$', '{r:rev}-{r:cogs}', USD, 'out'),
        ('gm', '毛利率', '%', 'IF({r:rev}=0,0,{r:gp}/{r:rev})', PCT, 'out'),
        ('infsh', '推論成本占營收成本比', '%', 'IF({r:cogs}=0,0,{r:inf}/{r:cogs})', PCT, 'out'),
        ('net', 'AI 淨效果（＝增量−侵蝕−推論成本）', '$', '{r:incr}-{r:eros}-{r:inf}', USD, 'key'),
        ('netr', 'AI 淨效果占單位營收比（÷ AI 前單位營收 R0）', '%', '{r:net}/{r:R0}', PCT, 'key'),
        ('incrr', '　增量 ÷ R0', '%', '{r:incr}/{r:R0}', PCT, 'out'),
        ('erosr', '　侵蝕 ÷ R0', '%', '{r:eros}/{r:R0}', PCT, 'out'),
        ('infr', '　推論成本（市場價）÷ R0', '%', '{r:inf}/{r:R0}', '0.00%', 'out'),
        ('inffcr', '　推論成本（全成本對照，模型商自做的成本下限）÷ R0', '%', '{r:inffc}/{r:R0}', '0.00%', 'out'),
        ('infsvr', '　推論成本（模型商服務成本對照，任務口徑、不含訓練攤提）÷ R0', '%', '{r:infsv}/{r:R0}', '0.00%', 'out'),
        ('pxfc', '　市場價 ÷ 全成本（混合口徑；參考請求單價 × 任務 token 有組合錯配，不宜引用）', 'x', 'IF({r:inffc}=0,0,{r:inf}/{r:inffc})', '0.0', 'out'),
        ('pxsv', '　市場價 ÷ 模型商服務成本（任務口徑）', 'x', 'IF({r:infsv}=0,0,{r:inf}/{r:infsv})', '0.0', 'out'),
        ('gpd', '單位毛利變化 ÷ R0（相對 AI 前毛利 R0 × 毛利率）', '%', '({r:gp}-{r:R0}*{r:gm0})/{r:R0}', PCT, 'out'),
        ('be', be_label, be_unit, be_expr, be_fmt, 'key'),
    ]


def mkt(tier, task):
    return f'INDEX({{x:MKT}},{{p:{tier}}},{task})'


def fcc(tier, task):
    return f'INDEX({{x:FC}},{{p:{tier}}},{task})'


def svc(tier, task):
    return f'INDEX({{x:SV}},{{p:{tier}}},{task})'


PROTO = {}
PROTO['L'] = dict(
    title='L 按人力計價（IT 服務、BPO）— 每 1 名交付人員年（AI 前工作範圍）',
    ident='R0＝人數(1) × 利用率 × 可用工時 × 時薪',
    calib=('Accenture／富士通', ['每員工營收與其年變化', '利用率', '毛利率與交付成本占比', 'GenAI 新訂單與營收（揭露期間與改口徑）', '員工數對營收成長的彈性']),
    falsify=['員工數下降而營收持平或上升（價格保留率高於基準）', '同工作範圍合約改按成果計價的比例', '每員工營收年增率顯著高於歷史（自動化比例上升）'],
    rows=common_rows() + [
        ('gm0', 'AI 前毛利率', '%', '{p:L_gm0}', PCT, 'aux'),
        ('R0', 'AI 前單位營收 R0（＝利用率 × 可用工時 × 時薪）', '$', '{p:L_util}*{p:L_hrs}*{p:L_rate}', USD, 'out'),
        ('a', '任務自動化比例（線性 2026→2032）', '%', lin('L_a26', 'L_a32'), PCT, 'drv'),
        ('n', 'AI 導入新專案營收 ÷ R0', '%', lin('L_n26', 'L_n32'), PCT, 'drv'),
        ('hauto', '被自動化工時', '小時', '{r:a}*{p:L_util}*{p:L_hrs}', '#,##0', 'drv'),
        ('tasks', '代理任務數（＝被自動化工時 ÷ 任務長度）', '件', '{r:hauto}/{x:TL5}', '#,##0.0', 'drv'),
        ('cps', '每成功任務推論成本 — 市場價', '$/件', mkt('L_tier', 5) + '*{r:pxidx}', '0.0000', 'drv'),
        ('cpsfc', '每成功任務推論成本 — 全成本對照', '$/件', fcc('L_tier', 5) + '*{r:fcidx}', '0.0000', 'drv'),
        ('cpssv', '每成功任務推論成本 — 模型商服務成本對照（任務口徑）', '$/件', svc('L_tier', 5) + '*{r:fcidx}', '0.0000', 'drv'),
        ('incr', 'AI 營收增量（新專案）', '$', '{r:n}*{r:R0}', USD, 'out'),
        ('eros', '定價侵蝕（被自動化工時 ×（1−價格保留率））', '$', '{r:R0}*{r:a}*(1-{p:L_r})', USD, 'out'),
        ('inf', '推論成本（市場價）', '$', '{r:tasks}*{r:cps}', USD, 'out'),
        ('inffc', '推論成本（全成本對照）', '$', '{r:tasks}*{r:cpsfc}', USD, 'out'),
        ('infsv', '推論成本（模型商服務成本對照）', '$', '{r:tasks}*{r:cpssv}', USD, 'out'),
        ('cogsx', '非推論營收成本（交付人力隨自動化減少；新專案同毛利率）', '$', '{r:R0}*(1-{r:gm0})*(1-{r:a})+{r:incr}*(1-{r:gm0})', USD, 'out'),
    ] + tail_rows('損益兩平所需的價格保留率（淨效果＝0；反解）', '%',
                  'IF({r:R0}*{r:a}=0,0,1-({r:incr}-{r:inf})/({r:R0}*{r:a}))', PCT) + [
        ('begp', '毛利不變的價格保留率 r*（單位毛利＝R0 × AI 前毛利率；反解，計入自動化省下的交付人力成本）', '%',
         'IF({r:R0}*{r:a}=0,0,{r:gm0}-({r:incr}*{r:gm0}-{r:inf})/({r:R0}*{r:a}))', PCT, 'key'),
    ],
    be_check='{r:incr}-{r:R0}*{r:a}*(1-{r:be})-{r:inf}',
    gp_check='{r:R0}*(1-{r:gm0})*{r:a}+{r:incr}*{r:gm0}-{r:R0}*{r:a}*(1-{r:begp})-{r:inf}',
    note='注意：AI 淨效果只計推論成本，不計自動化省下的交付人力成本；淨效果為負代表營收移轉給客戶，不等於廠商毛利受損，毛利看 gpd 與 begp 兩列。',
    r0_check="'In'!{L_util}*'In'!{L_hrs}*'In'!{L_rate}",
)
PROTO['S'] = dict(
    title='S 按席位計價（CRM、設計與生產力、人資、工作流）— 每 1 個 AI 前付費席位年',
    ident='R0＝席位(1) × 每席位年費',
    calib=('Salesforce', ['付費席位或使用者相關揭露', 'Agentforce 付費交易數與每次對話價格', 'Flex Credits／AELA 計價變更日期', '訂閱營收成長分解（量 vs 價）', '毛利率']),
    falsify=['客戶員工數下降時席位續約數同步下降（席位彈性）', '每任務／每次對話牌價相對 S_wage × 任務長度的比例（取價率）', '代理營收占訂閱營收比率'],
    rows=common_rows() + [
        ('gm0', 'AI 前毛利率', '%', '{p:S_gm0}', PCT, 'aux'),
        ('R0', 'AI 前單位營收 R0（＝1 席位 × 每席位年費）', '$', '1*{p:S_p}', USD, 'out'),
        ('h', '客戶人力替代率（線性）', '%', lin('S_h26', 'S_h32'), PCT, 'drv'),
        ('ad', '代理採用率（線性）', '%', lin('S_ad26', 'S_ad32'), PCT, 'drv'),
        ('tasks', '代理處理件數（＝採用率 × 每席位可交付任務）', '件', '{r:ad}*{p:S_V}', '#,##0.0', 'drv'),
        ('ptask', '每任務價格（＝時薪 × 任務長度 × 取價率 ×(1+年變化)^t）', '$/件', '{p:S_wage}*{x:TL3}*{p:S_cap}*(1+{p:S_gpt})^{t}', '0.000', 'drv'),
        ('cps', '每成功任務推論成本 — 市場價', '$/件', mkt('S_tier', 3) + '*{r:pxidx}', '0.0000', 'drv'),
        ('cpsfc', '每成功任務推論成本 — 全成本對照', '$/件', fcc('S_tier', 3) + '*{r:fcidx}', '0.0000', 'drv'),
        ('cpssv', '每成功任務推論成本 — 模型商服務成本對照（任務口徑）', '$/件', svc('S_tier', 3) + '*{r:fcidx}', '0.0000', 'drv'),
        ('incr', 'AI 營收增量（每任務計價）', '$', '{r:tasks}*{r:ptask}', USD, 'out'),
        ('eros', '定價侵蝕（席位縮減＝R0 × 替代率 × 席位彈性）', '$', '{r:R0}*{r:h}*{p:S_e}', USD, 'out'),
        ('inf', '推論成本（市場價）', '$', '{r:tasks}*{r:cps}', USD, 'out'),
        ('inffc', '推論成本（全成本對照）', '$', '{r:tasks}*{r:cpsfc}', USD, 'out'),
        ('infsv', '推論成本（模型商服務成本對照）', '$', '{r:tasks}*{r:cpssv}', USD, 'out'),
        ('cogsx', '非推論營收成本（席位與代理營收同毛利率）', '$', '({r:R0}-{r:eros}+{r:incr})*(1-{r:gm0})', USD, 'out'),
    ] + tail_rows('損益兩平所需的每任務價格（淨效果＝0；反解）', '$/件',
                  'IF({r:tasks}=0,0,({r:eros}+{r:inf})/{r:tasks})', '0.000'),
    be_check='{r:tasks}*{r:be}-{r:eros}-{r:inf}',
    r0_check="1*'In'!{S_p}",
)
PROTO['A'] = dict(
    title='A 按廣告計價（數位廣告）— 每 1,000 次 AI 前查詢',
    ident='R0＝查詢(1,000) × 廣告負載 × 點擊率 × 每次點擊價格',
    calib=('LY Corp', ['搜尋廣告營收與查詢相關指標', '帳號廣告與展示廣告拆分', 'AI 助理與生成式搜尋使用指標', '廣告營收成長分解（量 vs 價）', '毛利率']),
    falsify=['搜尋查詢量年變化（分流比例）', '每次點擊價格或每查詢營收年變化（轉換率提升）', '生成式回答覆蓋率與每查詢成本揭露'],
    rows=common_rows() + [
        ('gm0', 'AI 前毛利率', '%', '{p:A_gm0}', PCT, 'aux'),
        ('R0', 'AI 前單位營收 R0（＝1,000 × 廣告負載 × 點擊率 × CPC）', '$', '1000*{p:A_load}*{p:A_ctr}*{p:A_cpc}', USD, 'out'),
        ('d', '聊天介面分流比例（線性）', '%', lin('A_d26', 'A_d32'), PCT, 'drv'),
        ('u', 'AI 轉換率提升（線性；反映在 CPC）', '%', lin('A_u26', 'A_u32'), PCT, 'drv'),
        ('q', '附 AI 生成回答的比例（線性）', '%', lin('A_q26', 'A_q32'), PCT, 'drv'),
        ('tasks', '生成回答次數（＝保留查詢 × 覆蓋率）', '次', '1000*(1-{r:d})*{r:q}', '#,##0.0', 'drv'),
        ('cps', '每次成功回答推論成本 — 市場價', '$/次', mkt('A_tier', 1) + '*{r:pxidx}', '0.000000', 'drv'),
        ('cpsfc', '每次成功回答推論成本 — 全成本對照', '$/次', fcc('A_tier', 1) + '*{r:fcidx}', '0.000000', 'drv'),
        ('cpssv', '每次成功回答推論成本 — 模型商服務成本對照（任務口徑）', '$/次', svc('A_tier', 1) + '*{r:fcidx}', '0.000000', 'drv'),
        ('incr', 'AI 營收增量（保留查詢的 CPC 提升）', '$', '{r:R0}*(1-{r:d})*{r:u}', USD, 'out'),
        ('eros', '定價侵蝕（被分流的查詢）', '$', '{r:R0}*{r:d}', USD, 'out'),
        ('inf', '推論成本（市場價）', '$', '{r:tasks}*{r:cps}', USD, 'out'),
        ('inffc', '推論成本（全成本對照）', '$', '{r:tasks}*{r:cpsfc}', USD, 'out'),
        ('infsv', '推論成本（模型商服務成本對照）', '$', '{r:tasks}*{r:cpssv}', USD, 'out'),
        ('cogsx', '非推論營收成本（流量獲取等隨營收）', '$', '({r:R0}+{r:incr}-{r:eros})*(1-{r:gm0})', USD, 'out'),
    ] + tail_rows('損益兩平所需的轉換率提升（淨效果＝0；反解）', '%',
                  'IF({r:R0}*(1-{r:d})=0,0,({r:eros}+{r:inf})/({r:R0}*(1-{r:d})))', PCT),
    be_check='{r:R0}*(1-{r:d})*{r:be}-{r:eros}-{r:inf}',
    r0_check="1000*'In'!{A_load}*'In'!{A_ctr}*'In'!{A_cpc}",
)
PROTO['U'] = dict(
    title='U 按用量計價（資料平台）— 每 1 美元 AI 前運算額度營收',
    ident='R0＝運算單位(1 ÷ 單價) × 單價＝1',
    calib=('Snowflake', ['產品營收成長與淨營收留存率', 'Cortex／Snowflake Intelligence 使用客戶數', 'AI 影響的營收說法', '產品毛利率', '雲端供應商成本']),
    falsify=['產品營收成長扣除價格變動後的用量成長（工作負載乘數）', 'AI 功能營收或 credit 消耗占比', '產品毛利率是否因 AI 功能轉售而下降'],
    rows=common_rows() + [
        ('gm0', 'AI 前產品毛利率', '%', '{p:U_gm0}', PCT, 'aux'),
        ('units', '運算單位（＝1 ÷ 單價）', '單位', '1/{p:U_px}', '0.0000', 'aux'),
        ('R0', 'AI 前單位營收 R0（＝運算單位 × 單價）', '$', '{r:units}*{p:U_px}', USD, 'out'),
        ('w', '代理工作負載乘數（線性）', '%', lin('U_w26', 'U_w32'), PCT, 'drv'),
        ('e', '工作負載被替代（線性）', '%', lin('U_e26', 'U_e32'), PCT, 'drv'),
        ('tasks', 'AI 功能呼叫數（線性；每 1 美元 AI 前用量）', '次', lin('U_j26', 'U_j32') + '*{r:R0}', '0.00', 'drv'),
        ('cps', '每次成功呼叫推論成本 — 市場價', '$/次', mkt('U_tier', 2) + '*{r:pxidx}', '0.000000', 'drv'),
        ('cpsfc', '每次成功呼叫推論成本 — 全成本對照', '$/次', fcc('U_tier', 2) + '*{r:fcidx}', '0.000000', 'drv'),
        ('cpssv', '每次成功呼叫推論成本 — 模型商服務成本對照（任務口徑）', '$/次', svc('U_tier', 2) + '*{r:fcidx}', '0.000000', 'drv'),
        ('inf', '推論成本（市場價）', '$', '{r:tasks}*{r:cps}', '0.0000', 'out'),
        ('inffc', '推論成本（全成本對照）', '$', '{r:tasks}*{r:cpsfc}', '0.0000', 'out'),
        ('infsv', '推論成本（模型商服務成本對照）', '$', '{r:tasks}*{r:cpssv}', '0.0000', 'out'),
        ('airev', 'AI 功能轉售營收（＝推論成本 ×(1＋加成)）', '$', '{r:inf}*(1+{p:U_mk})', '0.0000', 'drv'),
        ('incr', 'AI 營收增量（工作負載乘數＋AI 功能轉售）', '$', '{r:w}*{r:R0}+{r:airev}', '0.0000', 'out'),
        ('eros', '定價侵蝕（被替代的工作負載）', '$', '{r:e}*{r:R0}', '0.0000', 'out'),
        ('cogsx', '非推論營收成本（基礎設施隨用量）', '$', '({r:R0}+{r:w}*{r:R0}-{r:eros})*(1-{r:gm0})', '0.0000', 'out'),
    ] + tail_rows('損益兩平所需的代理工作負載乘數（淨效果＝0；反解）', '%',
                  '{r:e}-{r:inf}*{p:U_mk}/{r:R0}', PCT),
    be_check='{r:be}*{r:R0}+{r:airev}-{r:eros}-{r:inf}',
    r0_check="(1/'In'!{U_px})*'In'!{U_px}",
)
PROTO['P'] = dict(
    title='P 按防護範圍計價（資安）— 每 1 個 AI 前受保護身分或端點年',
    ident='R0＝受保護數(1) × 模組數 × 每模組單價',
    calib=('CrowdStrike', ['ARR 與淨新增 ARR', '5／6／7／8 個以上模組客戶占比', 'Falcon Flex 合約', 'Charlotte AI 與非人類身分產品揭露', '訂閱毛利率']),
    falsify=['非人類身分產品營收或受保護數揭露（代理身分數 × 單價）', '多模組客戶占比年變化（模組加購）', '訂閱毛利率是否因 AI 分析成本下降'],
    rows=common_rows() + [
        ('gm0', 'AI 前毛利率', '%', '{p:P_gm0}', PCT, 'aux'),
        ('R0', 'AI 前單位營收 R0（＝1 × 模組數 × 每模組單價）', '$', '1*{p:P_mod}*{p:P_px}', USD, 'out'),
        ('n', '受保護代理身分數（線性）', '個', lin('P_n26', 'P_n32'), '0.00', 'drv'),
        ('m', '模組加購 ÷ R0（線性）', '%', lin('P_m26', 'P_m32'), PCT, 'drv'),
        ('e', 'AI 原生競爭價格侵蝕（線性）', '%', lin('P_e26', 'P_e32'), PCT, 'drv'),
        ('tasks', 'AI 資安分析次數（線性）', '次', lin('P_k26', 'P_k32'), '#,##0.0', 'drv'),
        ('cps', '每次成功分析推論成本 — 市場價', '$/次', mkt('P_tier', 3) + '*{r:pxidx}', '0.0000', 'drv'),
        ('cpsfc', '每次成功分析推論成本 — 全成本對照', '$/次', fcc('P_tier', 3) + '*{r:fcidx}', '0.0000', 'drv'),
        ('cpssv', '每次成功分析推論成本 — 模型商服務成本對照（任務口徑）', '$/次', svc('P_tier', 3) + '*{r:fcidx}', '0.0000', 'drv'),
        ('incr', 'AI 營收增量（代理身分＋模組加購）', '$', '{r:R0}*{r:n}*{p:P_rho}+{r:R0}*{r:m}', USD, 'out'),
        ('eros', '定價侵蝕', '$', '{r:R0}*{r:e}', USD, 'out'),
        ('inf', '推論成本（市場價）', '$', '{r:tasks}*{r:cps}', USD, 'out'),
        ('inffc', '推論成本（全成本對照）', '$', '{r:tasks}*{r:cpsfc}', USD, 'out'),
        ('infsv', '推論成本（模型商服務成本對照）', '$', '{r:tasks}*{r:cpssv}', USD, 'out'),
        ('cogsx', '非推論營收成本（隨營收）', '$', '({r:R0}+{r:incr}-{r:eros})*(1-{r:gm0})', USD, 'out'),
        ('beraw', '損益兩平代理身分單價比（未設下限；負值＝模組加購已足以覆蓋侵蝕與推論）', 'x',
         'IF({r:R0}*{r:n}=0,0,({r:eros}+{r:inf}-{r:R0}*{r:m})/({r:R0}*{r:n}))', '0.000', 'aux'),
    ] + tail_rows('損益兩平所需的代理身分單價 ÷ AI 前單價（下限 0；0＝不需要，模組加購已足以覆蓋）', 'x',
                  'MAX(0,{r:beraw})', '0.000'),
    be_check='{r:R0}*{r:n}*{r:beraw}+{r:R0}*{r:m}-{r:eros}-{r:inf}',
    r0_check="1*'In'!{P_mod}*'In'!{P_px}",
)
ORDER = ['L', 'S', 'A', 'U', 'P']

# ---------------------------------------------------------------- 工具


def hdr(ws, row, vals, widths=None):
    for i, v in enumerate(vals, 1):
        c = ws.cell(row, i, v)
        c.font = HFONT
        c.fill = HFILL
        c.alignment = WRAP
    if widths:
        for i, w in enumerate(widths, 1):
            ws.column_dimensions[get_column_letter(i)].width = w


def inp(c, fmt=None):
    c.font = BLUE
    c.fill = YFILL
    if fmt:
        c.number_format = fmt


def load_snapshot():
    s = json.load(open(SNAP, encoding='utf-8'))
    it = s['items']

    def v(n):
        return it[n]['values']

    return s, it, v


# ---------------------------------------------------------------- 建檔


def build(out):
    snap, items, v = load_snapshot()
    src = snap['source']
    wb = openpyxl.Workbook()
    ws0 = wb.active
    ws0.title = 'README'
    for name in ['In'] + ORDER + ['次產業', 'Sens', 'Checks', 'Sources']:
        wb.create_sheet(name)

    # ---------- In
    wi = wb['In']
    wi['A1'] = f'In — 全部輸入（藍字黃底）；AI 對軟體單位經濟 {VERSION}'
    wi['A1'].font = Font(bold=True, size=13)
    wi['A2'] = '每列輸入有 低／基準／高 三欄與標記；原型頁只讀「基準」欄，Sens 頁一次把一列換成低或高。標記：Verified／Interested-party／Analogy／Assumed／Derived。'
    hdr(wi, 4, ['名稱', '說明', '單位', '低', '基準', '高', '標記', '依據與推導'], [12, 58, 12, 11, 11, 11, 12, 110])
    prow = {}
    r = 5
    for p in P:
        if p[0] is None:
            c = wi.cell(r, 1, p[1])
            c.font = BOLD
            c.fill = SFILL
            for cc in range(2, 9):
                wi.cell(r, cc).fill = SFILL
            r += 1
            continue
        k, lab, unit, lo, base, hi, flag, basis, fmt = p
        wi.cell(r, 1, k)
        wi.cell(r, 2, lab).alignment = WRAP
        wi.cell(r, 3, unit)
        for j, x in enumerate((lo, base, hi)):
            inp(wi.cell(r, 4 + j, x), fmt)
        wi.cell(r, 7, flag)
        wi.cell(r, 8, basis).alignment = WRAP
        prow[k] = r
        r += 1
    p_first, p_last = 5, r - 1
    # g_cost（Derived；Tokenomics 全成本 GB300→VR200 的年化變化）— 列位置先保留，T 節建好後寫公式
    r += 1
    wi.cell(r, 1, '共用（Derived）：全成本年變化').font = BOLD
    for cc in range(1, 9):
        wi.cell(r, cc).fill = SFILL
    r += 1
    prow['g_cost'] = r
    wi.cell(r, 1, 'g_cost')
    wi.cell(r, 2, '模型商全成本年變化（Sol 層 IF_FullCost GB300→VR200 基準成本欄，按世代間隔年化）').alignment = WRAP
    wi.cell(r, 3, '%/年')
    wi.cell(r, 7, 'Derived')
    wi.cell(r, 8, '＝(IF_FullCost_Sol[VR200,基準] ÷ IF_FullCost_Sol[GB300,基準])^(1/Δ) − 1；低／基準／高分別用 Δ 的低／基準／高（間隔短＝年化下降快）。只用於「全成本對照」列，不進 AI 淨效果。').alignment = WRAP
    gcost_row = r
    r += 2

    # T 節：Tokenomics 快照
    T0 = r
    wi.cell(r, 1, f'T. Tokenomics 快照（只讀；{src["file"]}，{src["version"]}，commit {str(src["commit"])[:7]}，擷取 {src["extractedAt"]}；data/tk_snapshot.json）').font = BOLD
    for cc in range(1, 12):
        wi.cell(r, cc).fill = KFILL
    r += 1
    wi.cell(r, 1, '價格為「OpenAI 有效」單價（＝表列 ×（1 − IF_Disc）），代表軟體公司向模型商購買的市場價；全成本取「基準成本」情境欄，只作模型商自做的成本下限對照（IFC_Use：情境、相對比較）。').alignment = WRAP
    r += 1
    hdr_vals = ['層級', '序號', '名稱字尾', '新鮮輸入 $/M', '快取輸入 $/M', '思考 $/M', '可見輸出 $/M', '參考請求混合 $/M', '全成本 GB300 $/M', '全成本 VR200 $/M', '全成本 K6 預設 GB300 $/M']
    for i, h in enumerate(hdr_vals, 1):
        c = wi.cell(r, i, h)
        c.font = BOLD
        c.alignment = WRAP
    r += 1
    price_row = {}
    for i, tname in enumerate(TIERS, 1):
        wi.cell(r, 1, tname)
        wi.cell(r, 2, i)
        wi.cell(r, 3, f'_{tname}')
        vals = [v(f'IF_PriceFresh_{tname}'), v(f'IF_PriceCached_{tname}'), v(f'IF_PriceThink_{tname}'), v(f'IF_PriceOut_{tname}'),
                v(f'IF_PriceRef_{tname}'), v(f'IF_FullCost_{tname}')['GB300 NVL72']['基準'], v(f'IF_FullCost_{tname}')['VR200 NVL72']['基準'],
                v(f'IF_FullCostDefault_{tname}')['GB300 NVL72']['基準']]
        for j, x in enumerate(vals):
            inp(wi.cell(r, 4 + j, x), '0.0000')
        price_row[tname] = r
        r += 1
    r += 1
    for i, h in enumerate(['任務（Tokenomics Workload）', '', '名稱'] + TASKS, 1):
        c = wi.cell(r, i, h)
        c.font = BOLD
        c.alignment = WRAP
    r += 1
    task_rows = {}
    for key, lab, nm, fmt in [('len', '任務長度（人類完成時間，小時）', 'IF_TaskLen', '0.00'),
                              ('fr', '每次嘗試 新鮮輸入 token', 'IF_TaskTokFresh', '#,##0'),
                              ('ca', '每次嘗試 快取輸入 token', 'IF_TaskTokCached', '#,##0'),
                              ('de', '每次嘗試 decode token（思考＋可見）', 'IF_TaskTokDec', '#,##0'),
                              ('mult', '任務總 token ÷ 一般聊天總 token', 'IF_TaskTokMult', '0.000')] + \
            [(f'succ{t}', f'成功率 p — {t}', f'IF_TaskSucc_{t}', '0.000') for t in TIERS]:
        wi.cell(r, 1, lab)
        wi.cell(r, 3, nm)
        for j, x in enumerate(v(nm)):
            inp(wi.cell(r, 4 + j, x), fmt)
        task_rows[key] = r
        r += 1
    r += 1
    wi.cell(r, 1, '推導（公式）：每任務 token 與每成功任務推論成本').font = BOLD
    r += 1
    tot_row = r
    wi.cell(r, 1, '每次嘗試總 token（＝新鮮＋快取＋decode）')
    for j in range(5):
        cl = get_column_letter(4 + j)
        wi.cell(r, 4 + j, f'={cl}{task_rows["fr"]}+{cl}{task_rows["ca"]}+{cl}{task_rows["de"]}').number_format = '#,##0'
    r += 1
    mult_row = r
    wi.cell(r, 1, '總 token ÷ 一般聊天（對照 IF_TaskTokMult）')
    for j in range(5):
        cl = get_column_letter(4 + j)
        wi.cell(r, 4 + j, f'={cl}{tot_row}/$D${tot_row}').number_format = '0.000'
    r += 1
    att_rows = {}
    for t in TIERS:
        pr = price_row[t]
        wi.cell(r, 1, f'每次嘗試推論成本 — 市場價 — {t}（$；各 token 類型 × 對應有效單價）')
        for j in range(5):
            cl = get_column_letter(4 + j)
            wi.cell(r, 4 + j, f'=({cl}{task_rows["fr"]}*$D${pr}+{cl}{task_rows["ca"]}*$E${pr}+{cl}{task_rows["de"]}*$G${pr})/1000000').number_format = '0.000000'
        att_rows[t] = r
        r += 1
    mkt0 = r
    wi.cell(r - 0, 2, '')
    for t in TIERS:
        wi.cell(r, 1, f'每成功任務推論成本 — 市場價 — {t}（$；＝每次嘗試 ÷ 成功率）［MKT］')
        wi.cell(r, 2, TIERS.index(t) + 1)
        for j in range(5):
            cl = get_column_letter(4 + j)
            wi.cell(r, 4 + j, f'={cl}{att_rows[t]}/{cl}{task_rows["succ" + t]}').number_format = '0.000000'
        r += 1
    mkt1 = r - 1
    fc0 = r
    for t in TIERS:
        pr = price_row[t]
        wi.cell(r, 1, f'每成功任務推論成本 — 全成本對照 — {t}（$；＝總 token × IF_FullCost[GB300,基準] ÷ 成功率；混合口徑，快取比重高的任務偏高）［FC］')
        wi.cell(r, 2, TIERS.index(t) + 1)
        for j in range(5):
            cl = get_column_letter(4 + j)
            wi.cell(r, 4 + j, f'={cl}{tot_row}*$I${pr}/1000000/{cl}{task_rows["succ" + t]}').number_format = '0.000000'
        r += 1
    fc1 = r - 1
    sv0 = r
    for t in TIERS:
        wi.cell(r, 1, f'每成功任務模型商服務成本 — {t}（$；IF_CostSuccVR_{t}：VR200、經濟、基準成本與利用率，任務口徑、不含訓練攤提）［SV］')
        wi.cell(r, 2, TIERS.index(t) + 1)
        for j, x in enumerate(v(f'IF_CostSuccVR_{t}')):
            inp(wi.cell(r, 4 + j, x), '0.000000')
        r += 1
    sv1 = r - 1
    for rr in range(T0, r):
        wi.cell(rr, 1).alignment = WRAP
    # g_cost 公式
    sol = price_row['Sol']
    gap = prow['gap_vr']
    for j, col in enumerate('DEF'):
        c = wi.cell(gcost_row, 4 + j, f'=($J${sol}/$I${sol})^(1/{col}{gap})-1')
        c.number_format = PCT
    wi.freeze_panes = 'C5'

    FIX = {
        'MKT': f"'In'!$D${mkt0}:$H${mkt1}",
        'FC': f"'In'!$D${fc0}:$H${fc1}",
        'SV': f"'In'!$D${sv0}:$H${sv1}",
        'TL1': f"'In'!$D${task_rows['len']}", 'TL2': f"'In'!$E${task_rows['len']}", 'TL3': f"'In'!$F${task_rows['len']}",
        'TL4': f"'In'!$G${task_rows['len']}", 'TL5': f"'In'!$H${task_rows['len']}",
    }

    def pref(k, col='E'):
        return f"'In'!${col}${prow[k]}"

    # ---------- 原型頁
    rowmap = {}
    for code in ORDER:
        pr = PROTO[code]
        ws = wb[code]
        ws['A1'] = pr['title']
        ws['A1'].font = Font(bold=True, size=13)
        ws['A2'] = f'營收恆等式：{pr["ident"]}。AI 淨效果＝增量 − 侵蝕 − 推論成本（市場價）。全成本列只作對照，不進淨效果。' + pr.get('note', '')
        ws['A2'].alignment = Alignment(wrap_text=False)
        hdr(ws, 3, ['代碼', '項目', '單位'] + YEARS, [9, 62, 10] + [13] * len(YEARS))
        rm = {}
        rr = 4
        for row in pr['rows']:
            rm[row[0]] = rr
            rr += 1
        rowmap[code] = rm
        for row in pr['rows']:
            key, lab, unit, expr, fmt, kind = row
            r_ = rm[key]
            ws.cell(r_, 1, key)
            ws.cell(r_, 2, lab)
            ws.cell(r_, 3, unit)
            for j, y in enumerate(YEARS):
                col = get_column_letter(YC0 + j)
                if key == 't':
                    f = f'={col}$3-2026'
                else:
                    f = '=' + render(expr, col, rm, prow, FIX)
                c = ws.cell(r_, YC0 + j, f)
                c.number_format = fmt
                if kind == 'key':
                    c.font = BOLD
                    c.fill = KFILL
            if kind == 'key':
                ws.cell(r_, 2).font = BOLD
                ws.cell(r_, 2).fill = KFILL
        rr += 1
        ws.cell(rr, 1, '校準').font = BOLD
        ws.cell(rr, 2, f'樣板公司：{pr["calib"][0]}（實績對照；研究包完成前不得以公司指引反推參數）').font = BOLD
        rr += 1
        for ind in pr['calib'][1]:
            ws.cell(rr, 2, ind)
            ws.cell(rr, 3, '實績')
            ws.cell(rr, YC0, '待研究包')
            rr += 1
        rr += 1
        ws.cell(rr, 1, '可證偽').font = BOLD
        rr += 1
        for ind in pr['falsify']:
            ws.cell(rr, 2, ind).alignment = WRAP
            rr += 1
        ws.freeze_panes = 'D4'

    # ---------- 次產業
    wsi = wb['次產業']
    wsi['A1'] = '次產業 — 2025 全球營收池 × 基礎成長（不含 AI）× 滲透率 × 原型單位 AI 淨效果（$B）'
    wsi['A1'].font = Font(bold=True, size=13)
    wsi['A2'] = '本版未做先驗：營收池與成長率取第三方與 Interested-party 來源（見 Sources），AI 淨效果全由原型頁推得，未以任何主觀先驗調整。AI 移轉＝侵蝕（移給客戶或替代者）＋推論成本（移給模型層）。IT 服務的負值是營收移轉（L 原型不計自動化省下的交付人力成本），不等於廠商毛利受損；毛利見 L 頁 gpd、begp。'
    wsi['A2'].alignment = WRAP
    hdr(wsi, 3, ['代碼', '項目', '單位'] + YEARS, [12, 62, 8] + [13] * len(YEARS))
    sirow = {}
    rr = 4
    for code, nm, pc, scope in SI:
        wsi.cell(rr, 1, code).font = BOLD
        wsi.cell(rr, 2, f'{nm}（原型 {pc}；營收池口徑：{scope}）').font = BOLD
        for cc in range(1, 4 + len(YEARS)):
            wsi.cell(rr, cc).fill = SFILL
        rr += 1
        rm = {}
        defs = [
            ('base', '營收池（不含 AI）＝2025 營收池 ×(1+基礎成長)^(年−2025)', '$B', "'In'!$E${X_pool}*(1+'In'!$E${X_g})^({col}$3-2025)", '#,##0.0'),
            ('pen', '滲透率（線性 2026→2032）', '%', "('In'!$E${p26}+('In'!$E${p32}-'In'!$E${p26})*({col}$3-2026)/6)", PCT),
            ('netr', f'原型 {pc} 單位 AI 淨效果 ÷ R0（引用）', '%', "'{pc}'!{col}{X_netr}", PCT),
            ('incr', 'AI 營收增量', '$B', "{col}{base}*{col}{pen}*'{pc}'!{col}{X_incrr}", '#,##0.0'),
            ('eros', '定價侵蝕（移轉給客戶或替代者）', '$B', "{col}{base}*{col}{pen}*'{pc}'!{col}{X_erosr}", '#,##0.0'),
            ('inf', '推論成本（移轉給模型層）', '$B', "{col}{base}*{col}{pen}*'{pc}'!{col}{X_infr}", '#,##0.00'),
            ('net', 'AI 淨效果（＝增量−侵蝕−推論）', '$B', '{col}{incr}-{col}{eros}-{col}{inf}', '#,##0.0'),
            ('pool', '營收池（含 AI）＝不含 AI＋增量−侵蝕', '$B', '{col}{base}+{col}{incr}-{col}{eros}', '#,##0.0'),
            ('netsh', 'AI 淨效果 ÷ 營收池（不含 AI）', '%', 'IF({col}{base}=0,0,{col}{net}/{col}{base})', PCT),
            ('xfer', 'AI 移轉金額（＝侵蝕＋推論成本）', '$B', '{col}{eros}+{col}{inf}', '#,##0.0'),
        ]
        for key, _, _, _, _ in defs:
            rm[key] = rr
            rr += 1
        for key, lab, unit, f, fmt in defs:
            r_ = rm[key]
            wsi.cell(r_, 1, f'SI_{code}_{key}')
            wsi.cell(r_, 2, lab)
            wsi.cell(r_, 3, unit)
            for j, y in enumerate(YEARS):
                col = get_column_letter(YC0 + j)
                ff = f.format(col=col, pc=pc, X_pool=prow[f'SI_{code}_pool'], X_g=prow[f'SI_{code}_g'], p26=prow[f'SI_{code}_pen26'],
                              p32=prow[f'SI_{code}_pen32'], X_netr=rowmap[pc]['netr'], X_incrr=rowmap[pc]['incrr'], X_erosr=rowmap[pc]['erosr'],
                              X_infr=rowmap[pc]['infr'], **{k2: v2 for k2, v2 in rm.items()})
                c = wsi.cell(r_, YC0 + j, '=' + ff)
                c.number_format = fmt
                if key in ('net', 'xfer'):
                    c.font = BOLD
        sirow[code] = rm
        rr += 1
    wsi.cell(rr, 2, '合計（五個次產業）').font = BOLD
    rr += 1
    sitot = {}
    for key, lab in [('base', '營收池合計（不含 AI）'), ('net', 'AI 淨效果合計'), ('xfer', 'AI 移轉金額合計'), ('pool', '營收池合計（含 AI）')]:
        wsi.cell(rr, 1, f'SI_TOT_{key}')
        wsi.cell(rr, 2, lab)
        wsi.cell(rr, 3, '$B')
        for j in range(len(YEARS)):
            col = get_column_letter(YC0 + j)
            c = wsi.cell(rr, YC0 + j, '=' + '+'.join(f'{col}{sirow[c_[0]][key]}' for c_ in SI))
            c.number_format = '#,##0.0'
            c.font = BOLD
        sitot[key] = rr
        rr += 1
    wsi.freeze_panes = 'D4'

    # ---------- Sens
    wsn = wb['Sens']
    wsn['A1'] = f'Sens — 一次動一個輸入（In 頁每一列的低／高），看 {SENS_YEAR} 年「AI 淨效果占單位營收比」'
    wsn['A1'].font = Font(bold=True, size=13)
    wsn['A2'] = '每格是把原型頁的公式鏈整條展開、只把該輸入換成低值或高值後的結果（Excel 公式，非貼值）。「擺幅」＝五個原型中 |高案 − 低案| 的最大值。'
    heads = ['名稱', '說明', '標記', '低值', '高值']
    for code in ORDER:
        heads += [f'{code} 低案', f'{code} 高案']
    heads += ['次產業 AI 淨效果合計 低案 $B', '次產業 高案 $B', '擺幅（五原型最大）']
    hdr(wsn, 4, heads, [10, 46, 11, 9, 9] + [9] * 10 + [11, 11, 11])
    SIC = 6 + 2 * len(ORDER)  # 次產業欄起點

    def si_total(rr_, colcase, k):
        # 次產業 2030 AI 淨效果合計：Σ 營收池 ×(1+g)^5 × 滲透率(2030) × 該列該案的原型淨效果比
        parts = []
        for code, _, pc, _ in SI:
            def ref(name):
                col = colcase if k == f'SI_{code}_{name}' else 'E'
                return f"'In'!${col}${prow[f'SI_{code}_{name}']}"
            ci = ORDER.index(pc)
            if k.startswith('SI_'):
                netr = f'{get_column_letter(6 + 2 * ci)}${rb}'
            else:
                netr = f'{get_column_letter(6 + 2 * ci + (0 if colcase == "D" else 1))}{rr_}'
            parts.append(f"{ref('pool')}*(1+{ref('g')})^{SENS_YEAR - 2025}*({ref('pen26')}+({ref('pen32')}-{ref('pen26')})*{SENS_YEAR - 2026}/6)*{netr}")
        return '+'.join(parts)
    t_s = SENS_YEAR - 2026
    rb = 5
    wsn.cell(rb, 1, '基準')
    wsn.cell(rb, 2, '全部取基準欄')
    for i, code in enumerate(ORDER):
        f = '=' + inline(PROTO[code], 'netr', t_s, prow, FIX, None)
        for k in (0, 1):
            c = wsn.cell(rb, 6 + 2 * i + k, f)
            c.number_format = PCT
            c.font = BOLD
    for jj in (0, 1):
        c = wsn.cell(rb, SIC + jj, '=' + si_total(rb, 'E', ''))
        c.number_format = '#,##0.0'
        c.font = BOLD
    sens_rows = {}
    rr = rb + 1
    maxlen = 0
    for k in PKEYS + ['g_cost']:
        wsn.cell(rr, 1, k)
        lab = PDEF[k][1] if k in PDEF else '模型商全成本年變化（Derived）'
        wsn.cell(rr, 2, lab)
        wsn.cell(rr, 3, PDEF[k][6] if k in PDEF else 'Derived')
        wsn.cell(rr, 4, f"={pref(k, 'D')}").number_format = PDEF[k][8] if k in PDEF else PCT
        wsn.cell(rr, 5, f"={pref(k, 'F')}").number_format = PDEF[k][8] if k in PDEF else PCT
        for i, code in enumerate(ORDER):
            for jj, col in enumerate('DF'):
                f = '=' + inline(PROTO[code], 'netr', t_s, prow, FIX, (k, col))
                maxlen = max(maxlen, len(f))
                c = wsn.cell(rr, 6 + 2 * i + jj, f)
                c.number_format = PCT
        last = get_column_letter(6 + 2 * len(ORDER) - 1)
        parts = ','.join(f'ABS({get_column_letter(6 + 2 * i + 1)}{rr}-{get_column_letter(6 + 2 * i)}{rr})' for i in range(len(ORDER)))
        for jj, col in enumerate('DF'):
            wsn.cell(rr, SIC + jj, '=' + si_total(rr, col, k)).number_format = '#,##0.0'
        wsn.cell(rr, SIC + 2, f'=MAX({parts})').number_format = PCT
        sens_rows[k] = rr
        rr += 1
    assert maxlen < 8000, maxlen
    wsn.freeze_panes = 'D5'

    # ---------- Checks
    wc = wb['Checks']
    wc['A1'] = 'Checks — 核對（數值類全部應為 0）'
    wc['A1'].font = Font(bold=True, size=13)
    hdr(wc, 3, ['編號', '核對項目', '值', '容差', '結果', '說明'], [8, 70, 14, 10, 10, 70])
    rr = 4
    chk_rows = []
    yr = f'{get_column_letter(YC0)}{{r}}:{get_column_letter(YC0 + len(YEARS) - 1)}{{r}}'

    def rng(code, key):
        r_ = rowmap[code][key]
        return f"'{code}'!" + yr.format(r=r_)

    def add(cid, label, formula, note='', tol=1e-9, kind='num'):
        nonlocal rr
        wc.cell(rr, 1, cid)
        wc.cell(rr, 2, label).alignment = WRAP
        c = wc.cell(rr, 3, '=' + formula)
        c.number_format = '0.00E+00' if kind == 'num' else 'General'
        wc.cell(rr, 4, tol if kind == 'num' else None)
        if kind == 'num':
            wc.cell(rr, 5, f'=IF(ABS(C{rr})<=D{rr},"OK","FAIL")')
        else:
            wc.cell(rr, 5, f'=IF(C{rr}=0,"OK","FAIL")')
        wc.cell(rr, 6, note).alignment = WRAP
        chk_rows.append(rr)
        rr += 1

    for code in ORDER:
        pr = PROTO[code]
        r0c = pr['r0_check']
        for k in PKEYS:
            r0c = r0c.replace('{' + k + '}', f'$E${prow[k]}')
        add(f'{code}1', f'{code} 營收恆等式：R0 − 各項乘積（直接由 In 頁重算），各年絕對差合計', f"SUMPRODUCT(ABS({rng(code, 'R0')}-{r0c}))")
        add(f'{code}2', f'{code} 單位營收 −（R0＋增量−侵蝕）', f"SUMPRODUCT(ABS({rng(code, 'rev')}-({rng(code, 'R0')}+{rng(code, 'incr')}-{rng(code, 'eros')})))")
        add(f'{code}3', f'{code} AI 淨效果 −（增量−侵蝕−推論成本）', f"SUMPRODUCT(ABS({rng(code, 'net')}-({rng(code, 'incr')}-{rng(code, 'eros')}-{rng(code, 'inf')})))")
        add(f'{code}4', f'{code} 單位毛利 −（單位營收−營收成本）；營收成本 −（非推論＋推論）',
            f"SUMPRODUCT(ABS({rng(code, 'gp')}-({rng(code, 'rev')}-{rng(code, 'cogs')})))+SUMPRODUCT(ABS({rng(code, 'cogs')}-({rng(code, 'cogsx')}+{rng(code, 'inf')})))")
        bc = pr['be_check']
        bc_r = re.sub(r'\{r:(\w+)\}', lambda m: rng(code, m.group(1)), bc)
        bc_r = re.sub(r'\{p:(\w+)\}', lambda m: pref(m.group(1)), bc_r)
        add(f'{code}5', f'{code} 反解驗證：把損益兩平值代回，淨效果應為 0（各年絕對值合計）', f'SUMPRODUCT(ABS({bc_r}))', tol=1e-6)
        if pr.get('gp_check'):
            gc = re.sub(r'\{r:(\w+)\}', lambda m: rng(code, m.group(1)), pr['gp_check'])
            gp_id = (f"SUMPRODUCT(ABS({rng(code, 'gp')}-({rng(code, 'R0')}*{rng(code, 'gm0')}+{rng(code, 'R0')}*(1-{rng(code, 'gm0')})*{rng(code, 'a')}"
                     f"+{rng(code, 'incr')}*{rng(code, 'gm0')}-{rng(code, 'eros')}-{rng(code, 'inf')})))")
            add(f'{code}7', f'{code} 毛利反解驗證：單位毛利列與展開式一致，且把毛利不變的價格保留率代回後單位毛利 − R0 × AI 前毛利率＝0', f'SUMPRODUCT(ABS({gc}))+{gp_id}', tol=1e-6)
        add(f'{code}6', f'{code} Sens 基準列 − 原型頁 {SENS_YEAR} 年淨效果比（整條展開公式與逐列公式一致）',
            f"ABS(Sens!{get_column_letter(6 + 2 * ORDER.index(code))}{rb}-'{code}'!{get_column_letter(YC0 + t_s)}{rowmap[code]['netr']})")
    ys = f'{get_column_letter(YC0)}{{r}}:{get_column_letter(YC0 + len(YEARS) - 1)}{{r}}'

    def srng(code, key):
        return "'次產業'!" + ys.format(r=sirow[code][key])
    for code, nm, pc, _ in SI:
        add(f'SI_{code}', f'次產業 {nm}：淨效果 −（增量−侵蝕−推論）；含 AI 營收池 −（不含 AI＋淨效果＋推論）；淨效果 −（營收池 × 滲透率 × 原型淨效果比）',
            f"SUMPRODUCT(ABS({srng(code, 'net')}-({srng(code, 'incr')}-{srng(code, 'eros')}-{srng(code, 'inf')})))"
            f"+SUMPRODUCT(ABS({srng(code, 'pool')}-({srng(code, 'base')}+{srng(code, 'net')}+{srng(code, 'inf')})))"
            f"+SUMPRODUCT(ABS({srng(code, 'net')}-{srng(code, 'base')}*{srng(code, 'pen')}*{srng(code, 'netr')}))", tol=1e-6)
    add('SI_S', f'Sens 次產業基準 − 次產業頁 {SENS_YEAR} 年 AI 淨效果合計',
        f"ABS(Sens!{get_column_letter(SIC)}{rb}-'次產業'!{get_column_letter(YC0 + t_s)}{sitot['net']})", tol=1e-6)
    add('R1', '區間單調：低 ≤ 基準 ≤ 高 不成立的輸入列數',
        f"SUMPRODUCT((In!D{p_first}:D{p_last}>In!E{p_first}:E{p_last})+(In!E{p_first}:E{p_last}>In!F{p_first}:F{p_last}))+(In!D{gcost_row}>In!E{gcost_row})+(In!E{gcost_row}>In!F{gcost_row})", kind='cnt')
    add('R2', 'Sens 方向：每個輸入的低案與高案不得同時等於基準以外的值卻互換（低值案＝高值案且該輸入有用到時列出）；此處計數「低值 = 高值」的輸入列數',
        f"SUMPRODUCT(--(In!D{p_first}:D{p_last}=In!F{p_first}:F{p_last}),--ISNUMBER(In!E{p_first}:E{p_last}))", kind='cnt',
        note='每列 In 輸入都必須有區間（Analogy／Assumed 一律給區間）。')
    add('T1', 'Tokenomics 一致性：Σ 各 token 類型 ÷ 一般聊天 − IF_TaskTokMult（五個任務絕對差合計）',
        f"SUMPRODUCT(ABS(In!D{mult_row}:H{mult_row}-In!D{task_rows['mult']}:H{task_rows['mult']}))")
    add('T2', 'Tokenomics 一致性：IF_PriceThink − IF_PriceOut（思考依輸出計費；四個層級絕對差合計）',
        f"SUMPRODUCT(ABS(In!F{price_row['Luna']}:F{price_row['Astra']}-In!G{price_row['Luna']}:G{price_row['Astra']}))")
    errs = []
    for code in ORDER:
        last_r = max(rowmap[code].values())
        errs.append(f"SUMPRODUCT(--ISERROR('{code}'!D4:J{last_r}))")
    errs.append(f"SUMPRODUCT(--ISERROR(Sens!D5:R{rr + 200}))")
    errs.append(f"SUMPRODUCT(--ISERROR('次產業'!D4:J{max(sitot.values())}))")
    errs.append(f"SUMPRODUCT(--ISERROR(In!D5:K{sv1}))")
    add('E1', '錯誤值格數（In、五個原型頁、Sens）', '+'.join(errs), kind='cnt')
    # 快照版本與重現檢查（建檔時執行 --check，結果寫入）
    rc, msg = run_snapshot_check()
    wc.cell(rr, 1, 'V1')
    wc.cell(rr, 2, f'Tokenomics 快照版本：{src["file"]}（{src["version"]}；commit {src["commit"]}；擷取 {src["extractedAt"]}）')
    wc.cell(rr, 3, src['version'])
    wc.cell(rr, 5, '記錄')
    rr += 1
    wc.cell(rr, 1, 'V2')
    wc.cell(rr, 2, f'重現檢查 import_tokenomics.py --check（建檔時 {datetime.date.today().isoformat()} 執行；代碼 0＝通過）').alignment = WRAP
    wc.cell(rr, 3, rc)
    wc.cell(rr, 4, None)
    wc.cell(rr, 5, f'=IF(C{rr}=0,"OK","FAIL")')
    wc.cell(rr, 6, msg).alignment = WRAP
    chk_rows.append(rr)
    rr += 1
    use = items['IFC_Use']['values']
    bad = [n for n in items if not n.startswith('IFC') and isinstance(use.get(n), str) and '不得作預測' in use.get(n)]
    wc.cell(rr, 1, 'V3')
    wc.cell(rr, 2, '所取名稱中 IFC_Use 標「上限，不得作預測」者的個數（契約第 2 條第 2 項）')
    wc.cell(rr, 3, len(bad))
    wc.cell(rr, 5, f'=IF(C{rr}=0,"OK","FAIL")')
    wc.cell(rr, 6, '、'.join(bad) or '無')
    chk_rows.append(rr)
    rr += 2
    wc.cell(rr, 2, '全部核對').font = BOLD
    wc.cell(rr, 5, '=IF(COUNTIF(E4:E{0},"FAIL")=0,"全部 OK","有 FAIL")'.format(rr - 1)).font = BOLD
    total_row = rr

    # ---------- Sources
    wsr = wb['Sources']
    wsr['A1'] = 'Sources — 外部證據（只用於推導參數區間；樣板公司數字待研究包）'
    wsr['A1'].font = Font(bold=True, size=13)
    hdr(wsr, 3, ['編號', '文件', 'URL', '內容摘要', '用於', '標記與來源誘因', '擷取日'], [7, 50, 55, 60, 18, 36, 12])
    for i, s in enumerate(SOURCES):
        for j, x in enumerate(s):
            wsr.cell(4 + i, 1 + j, x).alignment = WRAP
        wsr.cell(4 + i, 7, '2026-11-16')
    r2 = 4 + len(SOURCES) + 1
    wsr.cell(r2, 1, 'Tokenomics 名稱（快照）').font = BOLD
    r2 += 1
    hdr(wsr, r2, ['名稱', '說明', '單位', 'IFC_Use', 'IFC_Conf', '位置', ''])
    r2 += 1
    conf = items['IFC_Conf']['values']
    for n, itv in items.items():
        if n.startswith('IFC'):
            continue
        wsr.cell(r2, 1, n)
        wsr.cell(r2, 2, itv.get('label')).alignment = WRAP
        wsr.cell(r2, 3, itv.get('unit'))
        wsr.cell(r2, 4, use.get(n))
        wsr.cell(r2, 5, conf.get(n))
        wsr.cell(r2, 6, itv.get('ref'))
        r2 += 1

    # ---------- README
    write_readme_sheet(ws0, src)

    wb.save(out)
    return dict(rowmap=rowmap, prow=prow, sens_rows=sens_rows, sirow=sirow, sitot=sitot, rb=rb, chk_rows=chk_rows, total_row=total_row, maxlen=maxlen)


def render(expr, col, rm, prow, FIX):
    def sub(m):
        kind, name = m.group(1), m.group(2)
        if kind == 'r':
            return f'{col}{rm[name]}'
        if kind == 'p':
            return f"'In'!$E${prow[name]}"
        if kind == 'x':
            return FIX[name]
        raise KeyError(kind)
    s = re.sub(r'\{(r|p|x):(\w+)\}', sub, expr)
    return s.replace('{t}', f'{col}$4')


def inline(pr, key, t, prow, FIX, swap):
    rows = {r[0]: r for r in pr['rows']}

    def ex(k):
        if k == 't':
            return str(t)
        e = rows[k][3]

        def sub(m):
            kind, name = m.group(1), m.group(2)
            if kind == 'r':
                return '(' + ex(name) + ')'
            if kind == 'p':
                col = swap[1] if swap and swap[0] == name else 'E'
                return f"'In'!${col}${prow[name]}"
            return FIX[name]
        s = re.sub(r'\{(r|p|x):(\w+)\}', sub, e)
        return s.replace('{t}', str(t))
    return ex(key)


def run_snapshot_check():
    try:
        p = subprocess.run([sys.executable, str(TOOL), '--check', str(SNAP), '--tokenomics', os.environ.get('TOKENOMICS_DIR', '/home/claude/yinchenchang/tokenomics')],
                           capture_output=True, text=True, timeout=900)
        return p.returncode, (p.stdout + p.stderr).strip()[:500]
    except Exception as e:  # noqa
        return 1, f'執行失敗：{e}'


def write_readme_sheet(ws, src):
    lines = [
        (f'AI 對軟體單位經濟 {VERSION}（{datetime.date.today().isoformat()}）', Font(bold=True, size=14)),
        ('命題：AI 進入一家軟體或服務公司後，每單位營收的變化＝AI 帶來的營收增量 − 定價侵蝕（席位、人天、查詢被替代）− 推論成本；三者並列，逐年 2026–2032。', BOLD),
        ('', None),
        ('工作表：In（全部輸入與 Tokenomics 快照）→ L／S／A／U／P（五個原型，逐年公式）→ 次產業（營收池 × 滲透率 × 原型淨效果）→ Sens（一次動一個輸入）→ Checks（恆等式、反解、區間、錯誤值、快照重現）→ Sources（外部證據與所取 Tokenomics 名稱）。', None),
        ('格式：藍字黃底＝輸入；黑字＝公式。原型頁只讀 In 頁「基準」欄；Sens 頁把公式鏈整條展開，一次把一個輸入換成低或高。', None),
        (f'推論成本只讀 Tokenomics：{src["file"]}（{src["version"]}，commit {str(src["commit"])[:7]}）。每成功任務成本＝Σ（任務各類 token × 該類有效單價）÷ 成功率 × 路徑指數；全成本（IF_FullCost，GB300 基準成本欄）只作模型商自做的成本下限對照。', None),
        ('未來 token 單價路徑：Tokenomics 無此名稱，寫「待 AI供應鏈 登錄」，以 g_px 區間作敏感度。', None),
        ('單位口徑：計算單位一律是「AI 前的 1 個單位」（同一工作範圍的世代追蹤），所以席位、人天、查詢的流失表現為侵蝕，而不是單位本身變少。', None),
        ('「AI 淨效果占單位營收比」＝AI 淨效果 ÷ AI 前單位營收 R0。', None),
        ('', None),
        ('驅動因子 → 推導量', BOLD),
        ('共用：g_px、g_tok、ovh → 推論成本指數；Tokenomics 單價 × 任務 token ÷ 成功率 → 每成功任務成本', None),
        ('L：自動化比例 a、價格保留率 r、新專案 n、層級 → 侵蝕＝R0·a·(1−r)；增量＝R0·n；推論＝(a·工時 ÷ 任務長度) × 每任務成本；反解 r*', None),
        ('S：人力替代 h、席位彈性 e、採用率 × 每席位任務、時薪 × 任務長度 × 取價率 → 侵蝕＝R0·h·e；增量＝件數 × 每件價格；反解每件價格 p*', None),
        ('A：分流 d、轉換提升 u、生成回答覆蓋 q → 侵蝕＝R0·d；增量＝R0·(1−d)·u；推論＝1000·(1−d)·q × 每次成本；反解 u*', None),
        ('U：工作負載乘數 w、被替代 e、AI 呼叫數 j、轉售加成 → 增量＝R0·w＋推論 ×(1＋加成)；侵蝕＝R0·e；反解 w*', None),
        ('P：代理身分 n × 單價比 ρ、模組加購 m、價格侵蝕 e、分析次數 k → 增量＝R0·(n·ρ＋m)；侵蝕＝R0·e；反解 ρ*', None),
        ('', None),
        ('限制：v0.1 的參數除 Tokenomics 快照與少數外部證據外多為 Assumed，區間寬；樣板公司校準欄「待研究包」。線性路徑（2026、2032 兩端）是簡化；各原型之間的相依（例如 S 的人力替代與 L 的自動化）未連動。', None),
    ]
    lines.insert(9, ('次產業頁：營收池（不含 AI）×滲透率×原型單位淨效果比＝AI 淨效果金額；AI 移轉＝侵蝕＋推論成本。本版未做先驗（營收池與成長取外部來源，未以主觀先驗調整）。', None))
    ws.column_dimensions['A'].width = 120
    ws.column_dimensions['B'].width = 70
    ws.column_dimensions['C'].width = 14
    for i, (t, f) in enumerate(lines, 1):
        c = ws.cell(i, 1, t)
        c.alignment = WRAP
        if f:
            c.font = f
    r = len(lines) + 2
    ws.cell(r, 1, '變數名稱表（穩定命名；個股模型的命題模組引用這些名稱）').font = BOLD
    r += 1
    hdr(ws, r, ['名稱（In 頁 A 欄＝輸入；原型頁 A 欄＝列代碼，引用時寫「頁!列代碼」）', '說明', '單位'])
    r += 1
    for p_ in P:
        if p_[0] is None:
            ws.cell(r, 1, p_[1]).font = BOLD
        else:
            ws.cell(r, 1, p_[0])
            ws.cell(r, 2, p_[1]).alignment = WRAP
            ws.cell(r, 3, p_[2])
        r += 1
    ws.cell(r, 1, 'g_cost').font = None
    ws.cell(r, 2, '模型商全成本年變化（Derived，In 頁）')
    r += 1
    for code in ORDER:
        ws.cell(r, 1, f'原型頁 {code} 的列代碼').font = BOLD
        r += 1
        for row in PROTO[code]['rows']:
            ws.cell(r, 1, f'{code}!{row[0]}')
            ws.cell(r, 2, row[1]).alignment = WRAP
            ws.cell(r, 3, row[2])
            r += 1
    ws.cell(r, 1, '次產業頁的列代碼：SI_<IT|CRM|AD|DATA|SEC>_<base|pen|netr|incr|eros|inf|net|pool|netsh|xfer>；合計 SI_TOT_<base|net|xfer|pool>').font = BOLD


def recalc(path):
    sys.path.insert(0, str(TOOL.parent))
    import import_tokenomics as it  # noqa
    tmp = it.recalc_copy(str(path))
    shutil.copy(tmp, path)


def verify(path, meta):
    wb = openpyxl.load_workbook(path, data_only=True)
    errs = []
    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for c in row:
                if isinstance(c.value, str) and c.value.startswith('#') and c.value.rstrip('!?0/').upper() in ('#REF', '#DIV', '#VALUE', '#NAME', '#N/A', '#NUM', '#NULL') or (isinstance(c.value, str) and c.value[:4] in ('#REF', '#DIV', '#VAL', '#NAM', '#N/A', '#NUM', '#NUL', 'Err:')):
                    errs.append(f'{ws.title}!{c.coordinate}={c.value}')
    wc = wb['Checks']
    fails = [(wc.cell(r, 1).value, wc.cell(r, 3).value, wc.cell(r, 5).value) for r in meta['chk_rows'] if wc.cell(r, 5).value != 'OK']
    return errs, fails, wc.cell(meta['total_row'], 5).value


def update_readme(path, meta):
    """把成品的數字讀出後組字，寫入 README.md 的 AUTO 區塊（不手寫數字）。"""
    wb = openpyxl.load_workbook(path, data_only=True)
    col = {y: YC0 + i for i, y in enumerate(YEARS)}

    def val(sheet, key, y, rm=None):
        rm = rm or meta['rowmap'][sheet]
        return wb[sheet].cell(rm[key], col[y]).value

    def pct(x, d=1):
        return f'{x * 100:+.{d}f}%'
    names = {'L': 'L 按人力', 'S': 'S 按席位', 'A': 'A 按廣告', 'U': 'U 按用量', 'P': 'P 按防護範圍'}
    out = ['| 原型 | 2026 | 2030 | 2032 | 單位毛利變化 ÷ R0（2030） | 推論成本（市場價）÷ R0（2032） |', '|---|---|---|---|---|---|']
    for c in ORDER:
        out.append(f'| {names[c]} | {pct(val(c, "netr", 2026))} | {pct(val(c, "netr", 2030))} | {pct(val(c, "netr", 2032))} | '
                   f'{pct(val(c, "gpd", 2030))} | {val(c, "infr", 2032) * 100:.2f}% |')
    out.append('')
    out.append(f'L 的反解（2030）：淨效果為 0 的價格保留率 {val("L", "be", 2030) * 100:.1f}%；毛利不變的價格保留率 r* {val("L", "begp", 2030) * 100:.1f}%'
               f'（基準 L_r＝{PDEF["L_r"][4] * 100:.0f}%）。L 的淨效果只計推論成本，不計自動化省下的交付人力成本，所以淨效果為負代表營收移轉給客戶，不等於廠商毛利受損。')
    p_raw = val('P', 'beraw', 2030)
    out.append(f'P 的反解（2030）：未設下限的值為 {p_raw:.3f}；'
               + ('負值表示模組加購已足以覆蓋侵蝕與推論成本，不需要代理身分營收，表上顯示下限 0。' if p_raw < 0 else '為正值。'))
    rmS = meta['sirow']
    tot = meta['sitot']
    si = ['', '| 次產業 | 2025 營收池（$B） | AI 淨效果 2030（$B） | AI 移轉 2030（$B） | AI 淨效果 2032（$B） |', '|---|---|---|---|---|']
    for code, nm, pc, _ in SI:
        si.append(f'| {nm} | {PDEF["SI_" + code + "_pool"][4]:,.0f} | {val("次產業", "net", 2030, rmS[code]):,.1f} | '
                  f'{val("次產業", "xfer", 2030, rmS[code]):,.1f} | {val("次產業", "net", 2032, rmS[code]):,.1f} |')
    ws = wb['次產業']
    si.append(f'| 合計 | | {ws.cell(tot["net"], col[2030]).value:,.1f} | {ws.cell(tot["xfer"], col[2030]).value:,.1f} | {ws.cell(tot["net"], col[2032]).value:,.1f} |')
    wc = wb['Checks']
    status = wc.cell(meta['total_row'], 5).value
    chk = [f'核對結果（{datetime.date.today().isoformat()} 重建）：{status}；核對項 {len(meta["chk_rows"])} 項。']
    blocks = {'results': '\n'.join(out), 'subind': '\n'.join(si), 'checks': '\n'.join(chk), 'vars': vars_table()}
    md = (HERE / 'README.md').read_text(encoding='utf-8')
    for k, body in blocks.items():
        a, b = f'<!-- AUTO:{k} -->', f'<!-- /AUTO:{k} -->'
        i, j = md.index(a), md.index(b)
        md = md[:i + len(a)] + '\n' + body + '\n' + md[j:]
    (HERE / 'README.md').write_text(md, encoding='utf-8')


def vars_table():
    o = ['| 名稱 | 說明 | 單位 | 低／基準／高 | 標記 |', '|---|---|---|---|---|']
    for p_ in P:
        if p_[0] is None:
            o.append(f'| **{p_[1]}** | | | | |')
            continue
        k, lab, u, lo, ba, hi, fl, _, _ = p_
        o.append(f'| `{k}` | {lab} | {u} | {lo}／{ba}／{hi} | {fl} |')
    o.append('| `g_cost` | 模型商全成本年變化（公式：Sol IF_FullCost GB300→VR200，按 gap_vr 年化） | %/年 | 公式 | Derived |')
    o += ['', '原型頁列代碼（引用寫 `<頁>!<代碼>`，例 `S!netr`）：', '', '| 頁 | 列代碼 |', '|---|---|']
    for c in ORDER:
        o.append(f'| {c} | ' + '、'.join(f'`{r[0]}`' for r in PROTO[c]['rows']) + ' |')
    return '\n'.join(o)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--out', default=str(HERE / 'dist' / f'{datetime.date.today().strftime("%Y%m%d")}_AI對軟體單位經濟_{VERSION}.xlsx'))
    ap.add_argument('--no-recalc', action='store_true')
    a = ap.parse_args()
    os.makedirs(os.path.dirname(a.out), exist_ok=True)
    meta = build(a.out)
    print(f'寫出 {a.out}（Sens 最長公式 {meta["maxlen"]} 字元）')
    if a.no_recalc:
        return 0
    recalc(a.out)
    errs, fails, total = verify(a.out, meta)
    print(f'LibreOffice 重算完成；錯誤值 {len(errs)} 格；核對 {total}；FAIL {len(fails)}')
    for e in errs[:20]:
        print('  錯誤', e)
    for f in fails:
        print('  FAIL', f)
    if not errs and not fails:
        update_readme(a.out, meta)
        print('README.md 的 AUTO 區塊已由成品數字更新')
    return 0 if not errs and not fails else 1


if __name__ == '__main__':
    sys.exit(main())
