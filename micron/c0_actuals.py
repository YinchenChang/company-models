# 美光 C0（實績）：公司實績 vs 層級隱含值（三層基準協調規則 v1.4.2 C0 第 2 點；v0.5.5，2026-11-15）。實作在 tools/memory_model/c0_actuals_common.py。
# 結論：公司未揭露 HBM 營收金額、占比、位元（FQ2–FQ4 FY2026）→ HBM 列「無可比實績」；已揭露的 DRAM（含 HBM）營收只做上限檢查（參考，不判定）。
import os, sys
sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "tools", "memory_model"))
import c0_actuals_common as _C

SPEC = dict(
    name="美光", cur="十億美元", mult="", conv="美元計價，無匯率換算", span="FQ2–FQ4（約 2026 年 1–9 月）", last="FQ4-26",
    qcols=["FQ2-26", "FQ3-26", "FQ4-26"], sumcol="FQ2–FQ4 FY2026 合計", den_key="dram", den_name="DRAM（含 HBM）營收", ledger_fmt={},
    fy_note=("會計年度約 9 月初結束：沿用模型的日曆對齊慣例（FY2026 約 2025Q4–2026Q3，見 company.py fy_desc），FQ2–FQ4 視為日曆 2026 年第一至第三季；實際期間為 2025-11-28 至 2026-09-03，比日曆 1–9 月約早一個月。"
             "偏誤方向：主讀數（累計）在成長年偏低——FQ2–FQ4 合併營收合計 119.5 十億美元，依日數換算的日曆 1–9 月約 125.6（FQ2 取 1 月 1 日以後部分、9 月 4–30 日以 FQ4 日均外推，Derived），實績累計偏低約 5%，方向向下；"
             "第二讀數（FQ4 × 4）：FQ4-26 為 14 週（較 13 週多約 7.7%，向上）與提早一個月（向下）互相抵銷，淨方向不定。"),
    same_src=("層級 HBM 市占 SIF_hbmsh_mu_26＝19%（Siliconomics M24；快照註記「原在 micron/company.py（sh26）逐字搬入」，來源欄為 Counterpoint 1Q26 21%、2Q26 18%，外部研調 Interested-party）；"
              "HBM 市場 SIF_HBM_Mkt_Base 2026 為 Siliconomics 自建（顆數 × 每顆 HBM 容量 × 單價；2025 年 350 億美元錨點取自美光自己的 TAM 說法，不是 2026 年）。"
              "結論：2026 年層級值不是本公司模型校準值（2028–29 市占為本公司模型假設，不在本次比較年）；不屬「同源校準」。本表無可比實績的原因是公司沒有揭露 HBM 營收，不是同源。"),
    miss_txt="HBM 營收、HBM 占 DRAM／總營收比、HBM 位元出貨（FQ2–FQ4 FY2026 皆未揭露）",
    missing=[("HBM 營收", "十億美元", "FQ4-26 書面說明（2026-09）未載 HBM 營收金額、run-rate 或市占，只稱 HBM 營收成長快於公司營收；FQ2-26 逐字稿：執行長表示 HBM 作為整體組合管理、不單獨揭露市占，亦未更新 HBM TAM；"
                         "執行長書面說明稱 HBM4 營收「已超過 10 億美元」為累計說法、無期間與口徑，不採用"),
             ("HBM 占 DRAM 營收比", "%", "公司未揭露；2026-03 逐字稿稱 2025 年 HBM 份額與 DRAM 份額相當，無數字"),
             ("HBM 位元出貨", "EB", "公司未揭露（FQ4-26 書面說明未載 HBM 位元）")],
    rows=[dict(key="rev", label="合併營收（公司實績；Q 頁實績輸入）", unit="十億美元", fmt="0.000", cells=["=Q!G4", "=Q!H4", "=Q!I4"],
               note=("Verified。FQ2-26 23.860、FQ3-26 41.456、FQ4-26 54.229 十億美元：美光 8-K 財報新聞稿與 CFO 書面說明（FQ4-26 Q4-FY26-Prepared-Remarks.pdf，網址見 H 欄；"
                     "本次核對 FQ4-26 營收 $54.2B、FQ2-26 $23.9B）；FQ3-26 為模型既有輸入（Q 頁來源）"), url="https://s25.q4cdn.com/621799436/files/doc_financials/2026/q4/Q4-FY26-Prepared-Remarks.pdf"),
          dict(key="dram", label="DRAM（含 HBM）營收（公司實績；Q 頁實績輸入）", unit="十億美元", fmt="0.000", cells=["=Q!G5", "=Q!H5", "=Q!I5"],
               note=("FQ2-26 18.768、FQ3-26 31.328：10-Q 分產品營收（Verified；FQ2 另與 2026-03-18 法說逐字稿 $18.8B（79%）相符）。FQ4-26：10-K 全年減 FQ1–FQ3 推算 DRAM 39.771、NAND 14.102（Derived，Verified 來源）；"
                     "模型 Q 頁沿用書面說明的四捨五入值 DRAM 39.8（73%）、NAND 14.1（26%），差 0.03 十億美元以內——該值進估值（FY2026 基準年），本版不改值，下次滾入 FQ1-27 時改用 10-K 推算值。公司不在 DRAM 內拆 HBM：此為 DRAM 加 HBM 合計；FY2026 全年 DRAM 營收書面說明稱「超過 1,000 億美元」，與四季合計約 100.7 相容"),
               url="https://s25.q4cdn.com/621799436/files/doc_financials/2026/q4/Q4-FY26-Prepared-Remarks.pdf ； https://www.fool.com/earnings/call-transcripts/2026/03/18/micron-mu-q2-2026-earnings-call-transcript")],
)


def add(wb, Co, L, F, T, hdr_row=None):
    return _C.add(SPEC, wb, Co, L, F, T)


def html_blocks(CW, lab):
    return _C.html_blocks(SPEC, CW, lab)
