# 對帳單草稿：CoreWeave v5.0（J4 公司宣稱新約價格；Wattonomics 上線 GW 差距）｜2026-11-12

> 草稿，供 chat 端放入 Drive `02-AI供應鏈/對帳/`。個股不改 Tokenomics、Wattonomics 值（三層規則）。

## 一、J4：CoreWeave 宣稱的新約每 MW 年收入 vs Tokenomics 錨

| 欄位 | 公司（CoreWeave） | Tokenomics v22.0（本模型基準） |
|---|---|---|
| 主張 | 新簽 4–5 年合約每 MW 年收入 US$20–25m；短天期（約 3–6 個月）合約約 US$40m；7 月起各 SKU 漲價（chat 端轉述：25%＋10%） | 每 MW 年收入＝在役世代加權 `IF_HoldEcon` × k；k＝IREN–Microsoft 長約 9.70 ÷ GB300 持有成本（0.891）與 H100 現貨（1.828）以隨需 10% 加權＝0.985 |
| 數值 | 20–25（中點 22.5）；短約 40 | GB300 `IF_HoldEcon` 10.886；VR200 12.225；模型情境 B FY2027 10.28、FY2030 11.22 |
| 單位 | US$m／MW·年 | US$m／MW·年（＝US$B／GW·年） |
| 口徑 | **分母未說明（IT 或設施 MW 待查）**；短約為「年化收入 ÷ 服務該叢集所需電力」；是否含預付、服務範圍未說明 | IT 關鍵電力；經濟口徑持有成本（WACC 10% 下的打平租金，含設施） |
| 出處 | Q3 2026 法說（2026-11-10；chat 端提供，新聞稿未見）；2026-09-17 新聞稿（finviz 轉載） | `coreweave/data/tokenomics_snapshot_v22.0.json`（Tokenomics 4a13ba4）Interface!13 列 |
| 等級 | Interested-party | Derived（上游權威） |
| 立場 | 管理層有呈現定價力的誘因 | 中立 |
| 影響名稱 | `IF_HoldEcon`（各世代）、`IF_GPUhrEcon`、`L1_GPUhr_GB300_vsBE`；若證實，影響 k（個股公司因素）而非 Tokenomics 成本 | — |

**本模型處理：** 不改基準 k。Sens「新增 MW 改用公司宣稱中點 22.5」加權目標價 317.77（+450%，強力買進）；整本重算反解：支持現價 60.77 所需 k_adj＝1.015（FY2030 情境 B 每 MW 年收入 11.40，只比模型 11.22 高 1.5%，遠低於公司宣稱 22.5）。公司宣稱中點是 Tokenomics GB300 錨的 2.07 倍。

**待查：** (1) 分母 IT 或設施（若設施，IT 口徑 ×1.2＝27）；(2) 是否含預付、網路與儲存等服務；(3) 宣稱只適用新增合約的哪個比例。證據若成立，送 Tokenomics 的是「市場新約價格 ÷ 持有成本」的觀察值（k 證據），不是成本參數。

## 二、Wattonomics：CoreWeave 上線 GW（原則 54；差距 >15%）

| 欄位 | Wattonomics v0.3（mw_crwv） | CoreWeave v5.0 |
|---|---|---|
| 主張 | 2026 年底 1,850 MW；2030 年底 8,000 MW（公司目標最多 8 GW） | 2026 年底 1,850（三情境同）；2030 年底 A 8,000／B 5,600／C 4,600（已簽約量） |
| 差距 | — | 2030：B −30.0%、C −42.5%（A 0%） |
| 單位／口徑 | IT MW（mwBasis＝IT，Assumed） | 同 |
| 出處 | `wattonomics/data/wattonomics_snapshot_v0.3.json`（20261105_Wattonomics_v0.3.xlsx，sha256 c91aedf1…）列 mw_crwv | `coreweave/dist/20261112_CoreWeave個股模型_v5.0.xlsx` Thesis 在役 MW；In mwB_30、mwC_30 |
| 等級 | Interested-party（公司目標值） | Assumed（沿用 v4.13 容量路徑；J2） |
| 立場 | 以公司目標上緣為值 | 情境加權 25／50／25 |
| 影響名稱 | Wattonomics 新雲上線 GW（CoreWeave 列）；全市場 GW 帳 | — |

**建議（R 線判斷）：** Wattonomics 的 2030 值取公司目標上緣，與個股基準（5.6 GW）差 30%；建議 Wattonomics 改以情境中值或標「公司目標」口徑。個股不改。
