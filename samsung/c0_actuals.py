# 三星 C0（實績）：公司實績 vs 層級隱含值（三層基準協調規則 v1.4.2 C0 第 2 點；v0.5.5，2026-11-15）。實作在 tools/memory_model/c0_actuals_common.py。
# 結論：公司未揭露 HBM 營收、占記憶體營收比、HBM 位元出貨（2026 年第一至第三季只有定性說明；3Q26 法說的 HBM4 季增、HBM4 占 HBM 比例列為參考）→ HBM 列「無可比實績」；已揭露的記憶體事業營收只做上限檢查（參考，不判定）。
import os, sys
sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "tools", "memory_model"))
import c0_actuals_common as _C

SPEC = dict(
    name="三星", cur="兆韓元", mult="*fx/1000", conv="韓元＝US$B × 匯率，匯率為 In 頁 fx，Derived", span="2026 年 1–9 月", last="2026Q3",
    qcols=["2026Q1", "2026Q2", "2026Q3"], sumcol="2026 年 1–9 月", den_key="mem", den_name="記憶體事業營收（含 HBM）", ledger_fmt={},
    fy_note="會計年度與日曆年相同，無須換月。",
    same_src=("層級 HBM 市占 SIF_hbmsh_ss_26＝29%（Siliconomics M24；快照註記「原在 samsung/company.py（sh26）逐字搬入」，來源欄為 Counterpoint 1Q26 21%、2Q26 33%，外部研調 Interested-party）"
              "（29% 含建模判斷：下半年 HBM4 放量，高於 Counterpoint 上半年平均 27%）；HBM 市場 SIF_HBM_Mkt_Base 2026 為 Siliconomics 由加速器顆數 × 每顆 HBM 容量 × 單價自建（M3、M3b），不由三星模型校準。"
              "結論：2026 年層級值不是本公司模型校準值（2028–29 市占為本公司模型假設，不在本次比較年）；不屬「同源校準」。本表無可比實績的原因是公司沒有揭露 HBM 營收，不是同源。"),
    miss_txt="HBM 營收、HBM 占記憶體營收比、HBM 位元出貨（2026 年第一至第三季皆未揭露）",
    missing=[("HBM 營收", "兆韓元", "公司未揭露 HBM 營收金額。2026Q3 財報新聞稿（2026-10-29）稱「大幅擴大 HBM4 銷售」「HBM 基底晶粒銷售成長」；同日 3Q26 法說揭露（韓國媒體 SBS Biz、MTN 報導，網址見 H 欄）："
                         "3Q HBM4 營收季增 3 倍以上、HBM4 占 HBM 營收約 60%、HBM 市占已與 DRAM 市占同等水準。"
                         "以上為成長率、HBM 內部組成與定性市占，沒有 HBM 營收金額或占記憶體營收比，無法換算成可比實績；列為參考、不判定（Interested-party，二手報導）。2025Q4 新聞稿亦無 HBM 金額",
              "https://biz.sbs.co.kr/amp/article/20000342946 ； https://news.mtn.co.kr/article/2026102911463093858"),
             ("HBM 占記憶體營收比", "%", "公司未揭露；3Q26 法說只給 HBM4 占 HBM 營收約 60%（HBM 內部組成，不是占記憶體營收比）"),
             ("HBM 位元出貨", "EB", "公司未揭露 HBM 位元出貨（2026Q3 新聞稿與 3Q26 法說報導皆無；法說「HBM 市占與 DRAM 市占同等水準」為定性說法，無數字）")],
    rows=[dict(key="mem", label="記憶體事業營收（公司實績）", unit="兆韓元", fmt="0.0", cells=["=Q!F4", "=Q!G4", "=Q!H4"],
               note=("2026Q1 74.8、2026Q2 120.8、2026Q3 140.7：Q 頁實績輸入，來源為三星 IR 季度簡報 1Q26、2Q26（Verified；沿用模型既有來源，本次未重新開啟 PDF：images.samsung.com 被擋）；"
                     "2026Q3 140.7 兆的出處：Blocks & Files 2026-10-30「Samsung memory virtually printing money」引述（占 DS 營收 147.1 兆的 95.7%；DS 營收取自三星 2026Q3 財報新聞稿 2026-10-29，Verified），記憶體金額為二手（Interested-party）"),
               url="https://images.samsung.com/is/content/samsung/assets/global/ir/docs/2026_2Q_conference_eng.pdf ； https://images.samsung.com/is/content/samsung/assets/global/ir/docs/2026_1Q_conference_eng.pdf ； https://www.blocksandfiles.com/flash/2026/10/30/samsung-memory-virtually-printing-money/5306688 ； https://news.samsung.com/global/samsung-electronics-announces-third-quarter-2026-results")],
)


def add(wb, Co, L, F, T, hdr_row=None):
    return _C.add(SPEC, wb, Co, L, F, T)


def html_blocks(CW, lab):
    return _C.html_blocks(SPEC, CW, lab)
