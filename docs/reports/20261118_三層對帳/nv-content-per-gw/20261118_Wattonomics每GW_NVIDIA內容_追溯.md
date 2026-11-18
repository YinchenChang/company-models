# Wattonomics 每 GW NVIDIA 內容：追溯與試算（watt261，2026-11-18）

唯讀。沒有改 repo、沒有 commit。基準＝main `2b2a962c` 的 Wattonomics v0.3.7（讀 Siliconomics v0.19 快照、Tokenomics v22.0 快照）。配套：同名 `.xlsx`（輸入頁藍字；試算全為公式；LibreOffice 重算錯誤格 0；12 個情境另以「改 SIL_In 後全模型重算」逐一核對，差 0）。產生程式在 `scripts/`。
**NVIDIA FY2027 第三季實績沒有用在任何地方**（2026 年只用 Q1＋Q2 實績 164.2，以及模型既有的 375.46 錨點）。

## 結論

1. **261 億是 `SIL_In!D32`（`SIL_NVContentGW_VR200`）＝26.1263 十億美元／IT GW。** Wattonomics 沒有自己的輸入，原封讀 Siliconomics 的 Interface 名稱（`Siliconomics Interface!E862`＝`nvr_VR27 × TK_CapexIT_VR200`＝0.6951 × 37.586352）。0.6951 來自晶片內容橋接表 v0.2.3 `Bridge!D46`（營收口徑基準）。等級 [Derived]。
2. **「Siliconomics 用 34.81、Wattonomics 用 26.13」不成立：兩層讀同一個名稱、值逐位相同。34.81 是 Rubin Ultra（`SIF_NVContentGW_RU`＝0.6951 × 50.0858），不是 VR200。** 在整個 Siliconomics v0.20、Wattonomics v0.3.7 活頁簿掃描 34.80–34.83，只出現在 RU 的三處。三層（同一世代、同一定義）的數值一致性見第 2 節：Wattonomics＝Siliconomics 逐位相同（差 0）；Tokenomics 沒有這個量，只有機架價與 IT 資本支出（定義不同，不是數值不同）。
3. **只動每 GW 內容，2025 年比值 0.826 可在中心情境下全部或大部分補上。** 進入 ±15% 帶只需內容 +2.9%（19.88→20.46）。中心情境：NVIDIA 自稱（Blackwell 250）1.009；證據基準 1.012；**只用獨立的 IREN 採購證據（不含 NVIDIA 自稱）0.968**。缺口的主體來自 **GB200（占 2025 年新增 70%，現值 18.53，是 Analogy）**，不是 Rubin；GB300 現值 24.88 與證據（24.4、IREN 21.8–24.6）相符。
4. 2026（對錨點 375.46）：現行 0.939；試算 0.82–1.11；**2027（指引推得，只作參照）**：現行 0.645；試算 0.49–1.00。
5. **這個參數只動對帳點，不動基準與換算值 x2。** LibreOffice 全模型逐格比對：變動只限 Compare F 節（含 Siliconomics 並列欄 F3）、Interface 的 `WIF_Recon_NV_` 中 6 類共 20 格、Checks 的 N-K0／N-K1 與 WGOV_Warn 計數。Players、Deploy、CapexGW、Spend、Gap、Funding 全部逐格不變；`baseline_bridge.py` 讀的五個名稱都不在變動集合內。
6. 修正要由誰決定、有哪些選項：第 5 節。

## 1. 完整公式鏈與各世代每 GW 內容

### 1.1 鏈（附位置與原文）

| 步 | 內容 | 位置 |
|---|---|---|
| a | chip_bridge：各零件 $B/GW；基準比例 `NVIDIA 營收 ÷ IT capex`＝（GPU 模組售價＋CPU＋SuperNIC／DPU＋NVLink 交換晶片＋交換匣其他＋匣級殘差）÷ IT capex | `markets/chip_bridge/dist/_待刪除/20261105_晶片內容橋接表_v0.2.3.xlsx` Bridge!B45:D47（v0.2.4 同值，`dist/20261118_晶片內容橋接表_v0.2.4.xlsx`） |
| b | Siliconomics 輸入 `nvr_*`（GB300、VR200 為 Derived；H100、GB200、RU 為 Analogy 代入） | `siliconomics/model/20261116_Siliconomics_v0.20.xlsx` In!E605:E610 |
| c | `SIF_NVContentGW_<世代> = nvr_<世代> * TK_CapexIT_<世代>`（原文公式） | 同檔 Interface!E858:E863；程式 `siliconomics/builder/build.py` 第 1667–1682 行 |
| d | Wattonomics 讀入（綠字，不得修改） | `wattonomics/model/20261118_Wattonomics_v0.3.7.xlsx` SIL_In!D20:D37；`builder/build.py` 第 641–650 行（`scal` 清單） |
| e | Compare F1：`=SIL_NVContentGW_<世代>`（VR200 另有 2026 專用名稱 `_VR200_26`） | 同檔 Compare!D28:D33（`build.py` 第 1562–1567 行） |
| f | (5) 世代加權內容：`=Players!F5*SIL_NVContentGW_H100+Players!F6*SIL_NVContentGW_GB200+…`（原文）；2025–2026 的 VR200 用 `_VR200_26`，2027 用 `SIL_NVContentGW_VR200` | Compare!F125:H125（`build.py` 第 1680–1683 行） |
| g | (6) 隱含營收＝(4)×(5)；(7) 實績；(8) 比值 | Compare!F128:H132（`build.py` 第 1687–1695 行）；對外 `WIF_Recon_NV_ImpliedRev_<年>` 等於 Interface 對應列（`build.py` 第 2913–2932 行） |

（Siliconomics Interface 第 861 列＝VR200_26、第 862 列＝VR200、第 863 列＝RU；值在 E 欄。）

### 1.2 各世代（$B／IT GW）

| 世代 | 現值 | 公式（nvr × TK IT capex） | 等級 | 定義 | 每封裝 NVIDIA 營收 | 各年權重（2025／2026／2027） |
|---|---|---|---|---|---|---|
| Hopper H100 | 19.33 | 0.7027 × 27.510 | Analogy（以 GB300 營收口徑低值代入） | 見下 | 26.9 千美元 | 0.10／0／0 |
| GB200 NVL72 | **18.53** | 0.7717 × 24.015 | Analogy（以 GB300 比例代入，橋接表未拆 GB200） | 見下 | 36.9 | **0.70**／0／0 |
| GB300 NVL72 | 24.88 | 0.7717 × 32.237 | Derived | 見下 | 51.1 | 0.20／0.78／0 |
| VR200 NVL72 | **26.13** | 0.6951 × 37.586 | Derived | 見下 | 89.6 | 0／0.22／0.9397 |
| Rubin Ultra | 34.81 | 0.6951 × 50.086 | Analogy（以 VR200 比例代入） | 見下 | 178.8 | 0／0／0.0603 |

- **定義（營收口徑，晶片內容橋接表基準）：** 含 GPU 模組（以售價計）、Vera／Grace CPU、SuperNIC＋DPU、NVLink 交換晶片、交換匣其他、**匣級殘差（全額歸 NVIDIA；VR200 為 5.61，占內容 21%）**。**不含**機架外 scale-out 網通（交換器、光模組、線纜；VR200 3.40）、機架內 LPDDR5X（1.47）、NAND（4.05）、非 NVIDIA 機架內容。低＝只含 GPU 模組、CPU、NIC／DPU、NVLink 晶片（VR200 19.67）；高＝基準＋LPDDR＋scale-out×NVIDIA 份額（VR200 29.30）。VR200 基準加總驗算：16.0459＋0.7294＋2.0422＋0.8509＋0.8509＋5.608＝26.127。
- **世代組合權重的來源：** Wattonomics 自己的輸入 `In!J159:L163`（`build.py` 第 453、694–702 行）。2025＝CoreWeave 2025 新增組合 GB200 70%／GB300 20%／H200 10%（[Analogy]；In 頁標 Assumed）；2026＝GB300 `1−SIL_vr26`、VR200 `SIL_vr26`（0.22，[Interested-party] TrendForce）；2027＝VR200 `1−RU 比`、RU 比＝Siliconomics Units ÷ TK 每 GW 顆數＝6.03%（**2027 Blackwell＝0 是 v0.1 工作單簡化、非證據**，見 `wattonomics/docs/reports/20261118_W2_A_2027不計Blackwell重審.md`）。

## 2. 三層並列（同一世代）

數字全在 xlsx「三層並列」頁（公式）。

| 世代 | Tokenomics 機架全價 $B/GW（架數×機架價；含非 NVIDIA 零件） | Tokenomics IT capex $B/GW | Siliconomics 每 GW NVIDIA 內容 | Wattonomics 讀值 | 差 | NVIDIA 自稱 | rubin-price 證據基準 |
|---|---|---|---|---|---|---|---|
| H100 | 24.68 | 27.51 | 19.33 | 19.33 | 0 | 18 | 無證據 |
| GB200 | 21.60 | 24.02 | 18.53 | 18.53 | 0 | 25（Blackwell） | 25.09 |
| GB300 | 29.09 | 32.24 | 24.88 | 24.88 | 0 | 25 | 24.35 |
| VR200 | 34.04 | 37.59 | 26.13 | 26.13 | 0 | 40 | 33.28（28.0–40.7；直接證據 37.1） |
| Rubin Ultra | 45.44 | 50.09 | 34.81 | 34.81 | 0 | 無 | 無證據 |

- **一致性：** Siliconomics＝Wattonomics 數值與定義完全一致（Wattonomics 無自有輸入）。**Tokenomics 沒有「NVIDIA 內容」名稱**：`IF_RackPrice`（買方取得價；VR200 含約 $1M NAND）、`IF_RacksPerGW`、`IF_GPUsPerGW`、`IF_CapexIT` 都含非 NVIDIA 零件（且 CapexIT 含機架外網通）。所以三層並列是「定義不同」，不是「同一定義下數值不同」。Siliconomics 內容是機架全價的 77%–86%。
- **34.81 vs 26.13 的差異原因：** nvr 相同（RU 以 VR200 代入），差全來自 Tokenomics RU 每 GW IT capex 是 VR200 的 1.333 倍（每封裝 IT capex 2 倍、每 GW 封裝數 0.667 倍）。Siliconomics 以「成本等值 GW」（分母 VR200 的 37.59）計時，RU 每成本等值 GW 內容也是 26.13，所以 Siliconomics 的 R4 橋接 `n2` 2027 年全用 26.127（不加 RU 的 34.81）。**兩層 2027 年的內容差只有 26.127 對 26.650（Wattonomics 含 6.03% RU）**。
- **Recon F 節（Siliconomics）：** 隱含 NVIDIA 資料中心營收＝`n1`（出貨年 NVIDIA 成本等值 GW）× `n2`（每 IT GW 內容；2026 為 GB300／VR200 依 vr26 混合 25.1528，2027 為 VR200 26.127）＝448.36（2026）、743.78（2027）（`siliconomics/docs/reports/r4_bridge/20261116_R4對帳橋接_Siliconomics側_r12.xlsx` Bridge!n3）。**與 Wattonomics 的 352.67／408.76 的差來自 NVIDIA GW（17.83／28.47 成本等值對 14.02／15.34 實體），不是內容。**
- **證據對各層的比（每封裝，基準）：** Blackwell 證據 50 千美元＝Siliconomics GB300 每封裝的 0.98 倍，**但＝GB200 的 1.35 倍**；Rubin 證據 114.1 千美元＝Tokenomics 機架價÷封裝（116.7）的 0.98 倍、IT capex÷封裝（128.8）的 0.89 倍、Siliconomics NVIDIA 每封裝（89.6）的 1.27 倍。意思是證據線把「整個機架價」近似當成 NVIDIA 內容（且 NVIDIA 400 億本就含機架外網通），而 Siliconomics 只認機架內約 69.5%。這個定義差距是 26.13 偏低的最直接來源（第 5 節決策 1）。
- **同一定義下數值是否一致？** 不能比：只有 Siliconomics 與 Wattonomics 同定義（一致）；NVIDIA 自稱與 rubin-price 證據是含網通的口徑，Siliconomics 低／基準／高的高值（VR200 29.30、GB300 27.83）才是最接近的定義。

## 3. 試算（只換每 GW 內容；NVIDIA 出貨 GW、世代組合、實績都不動）

### 3.1 情境定義（$B/GW；Rubin Ultra＝該情境 VR200 × 1.3325〔現行 RU÷VR200 結構比，Assumed，兩邊都無 RU 證據〕）

| 情境 | H100 | GB200 | GB300 | VR200 | RU | 依據 |
|---|---|---|---|---|---|---|
| B0 現行＝Siliconomics 基準 | 19.33 | 18.53 | 24.88 | 26.13 | 34.81 | — |
| (a) NVIDIA 自稱 | 18 | 25 | 25 | 40 | 53.30 | Hopper 180、Blackwell 250、Vera Rubin 400 億，[Interested-party]，GW 口徑未定義 |
| (b) 證據低 | 19.33（無證據，沿用） | 21.58 | 20.94 | 27.98 | 37.29 | Blackwell 4.3 萬、Rubin 9.59 萬美元／封裝 × TK 每 GW 封裝數 |
| (b) 證據基準（單價×密度） | 19.33 | 25.09 | 24.35 | 33.28 | 44.35 | 5.0 萬、11.41 萬 |
| (b) 證據基準（Rubin 直接內容） | 19.33 | 25.09 | 24.35 | 37.10 | 49.43 | rubin-price 每GW檢查!C3 |
| (b) 證據高 | 19.33 | 27.60 | 26.79 | 40.65 | 54.17 | 5.5 萬、13.93 萬 |
| (c) Siliconomics 低值 | 16.51 | 16.88 | 22.65 | 19.67 | 26.21 | SIF_NVContentGW_*_lo |
| (c) Siliconomics 高值 | 21.23 | 20.73 | 27.83 | 29.30 | 39.05 | SIF_NVContentGW_*_hi |
| (c′) 若 VR200 也用 34.81 | 19.33 | 18.53 | 24.88 | 34.81 | 34.81 | 假設性：只回答「若這樣」 |
| (d) 只動 Blackwell：IREN 獨立證據低／基準／高 | 現值 | 22.41／23.90／25.40 | 21.75／23.20／24.65 | 現值 | 現值 | IREN $5.8B÷200 MW，[Verified] 金額與 MW，NVIDIA 占比 0.75–0.85 [Assumed] |

(c) 其實就是現值與它的低／高：Siliconomics 的值本來就是 Wattonomics 的值。(d) 是為了避開循環而加：(a) 與 (b) 的 Blackwell 每封裝（B1、B2、B3）都部分源自 NVIDIA 自己的營收或每 GW 說法，用來對 NVIDIA 實績是循環的；只有 IREN（B4）獨立。

### 3.2 結果（隱含營收 $B；比值＝隱含 ÷ 實績，帶＝0.85–1.15）

| 情境 | 2025 隱含 | **2025 比值** | 2026 隱含 | **2026 比值（錨點 375.46）** | 2027 隱含 | **2027 比值（指引推得 633.7，只作參照）** |
|---|---|---|---|---|---|---|
| B0 現行 | 160.0 | **0.826（帶外）** | 352.7 | 0.939 | 408.8 | 0.645 |
| (a) NVIDIA 自稱 | 195.5 | 1.009 | 396.8 | 1.057 | 625.8 | 0.988 |
| (b) 低 | 170.8 | 0.882 | 315.3 | 0.840（帶外） | 437.8 | 0.691 |
| (b) 基準（單價×密度） | 196.1 | 1.012 | 369.0 | 0.983 | 520.7 | 0.822 |
| (b) 基準（Rubin 直接） | 196.1 | 1.012 | 380.7 | 1.014 | 580.4 | 0.916 |
| (b) 高 | 214.1 | 1.105 | 418.3 | 1.114 | 636.0 | 1.004 |
| (c) Siliconomics 低 | 144.8 | 0.747（帶外） | 308.4 | 0.821（帶外） | 307.7 | 0.486 |
| (c) Siliconomics 高 | 178.6 | 0.922 | 394.7 | 1.051 | 458.4 | 0.723 |
| (c′) VR200＝34.81 | 160.0 | 0.826 | 379.5 | 1.011 | 534.0 | 0.843 |
| (d) IREN 低 | 176.8 | 0.913 | 318.5 | 0.848（帶外） | 408.8 | 0.645 |
| (d) IREN 基準 | 187.5 | **0.968** | 334.3 | 0.890 | 408.8 | 0.645 |
| (d) IREN 高 | 198.3 | 1.024 | 350.2 | 0.933 | 408.8 | 0.645 |

- **2025 年是否進 ±15%：** 現行帶外（0.826）。(a)、(b) 低／基準／高、(c) 高、(d) 低／基準／高全部進帶；(c′) 不動 2025（仍 0.826，因 2025 沒有 Rubin）；(c) 低（內容更低）更差（0.747）。
- **2026 年「只用 Q1＋Q2 實績」的對齊（Siliconomics 現行讀法，164.2 對隱含 × 月數／季節權重）：**

| 情境 | 錨點（含 Q3 指引、Q4 假設） | H1 線性（隱含×6/12 ÷ 164.2） | H1 季節（隱含×0.414 ÷ 164.2；草案 C0 第 11 點，只顯示） |
|---|---|---|---|
| B0 現行 | 0.939 | 1.074 | 0.889 |
| (a) | 1.057 | 1.208（帶外） | 1.001 |
| (b) 低／基準／直接／高 | 0.840／0.983／1.014／1.114 | 0.960／1.124／1.159（帶外）／1.274（帶外） | 0.795（帶外）／0.930／0.960／1.055 |
| (c) 高 | 1.051 | 1.202（帶外） | 0.995 |
| (d) 基準 | 0.890 | 1.018 | 0.843（帶外） |

  （Siliconomics 的主讀數方向是實績 ÷ 隱含 H1，等於上表線性欄的倒數；現行 0.931。）H1 線性在內容上調時最先帶外，因為成長年中線性分攤高估層級同期間（Siliconomics 自述）。
- **2027（指引推得，未入判定）：** Rubin Ultra 取法對結果有 ±2 個百分點影響（RU 維持 34.81／無加成／結構比：(a) 0.961／0.968／0.988；(b) 基準 0.808／0.806／0.822；(b) 高 0.975／0.984／1.004）。2027 年權重 94% 在 VR200，結論取決於 VR200 值與 2027 NVIDIA GW 15.34 本身（另一個變數，本報告不動）。

### 3.3 2025 年比值 0.826 的缺口拆解

| 項目 | 數值 |
|---|---|
| 缺口 | 實績 193.7 − 隱含 160.0 ＝ 33.7 |
| 要比值＝1.00 所需 | 內容 24.07（+21.1%）**或** NVIDIA 出貨 GW 9.74（+21.1%）；乘法式，兩者可互換 |
| 要進帶（0.85）所需 | 內容 20.46（**+2.9%**）或 GW 8.28（+2.9%） |

| 情境 | 缺口填補比（Δ隱含÷缺口） | 若內容單獨解釋後，GW 還需 +（到 1.00） | 到進帶 GW 還需 |
|---|---|---|---|
| (a) NVIDIA 自稱 | 105% | −0.9%（GW 反而可低） | 0 |
| (b) 基準 | 107% | −1.2% | 0 |
| (b) 低 | 32% | +13.4% | 0（已進帶） |
| (b) 高 | 161% | −9.5% | 0 |
| (c) 高 | 55% | +8.4% | 0 |
| (d) IREN 低／基準／高 | 50%／82%／114% | +9.6%／**+3.3%**／−2.3% | 0 |

- **回答：** 以中心情境，**每 GW 內容偏低可解釋缺口的 82%（只用獨立證據）至 107%（含 NVIDIA 自稱與單價×密度）**；需要由 GW 本身偏低解釋的部分，中心情境為 0 至 3.3% 的 GW（獨立證據基準下到 1.00 需 GW +3.3%；進帶不需要）。保守端（證據低）內容只解釋 32%，GW 還需 +13%。
- **哪個世代：** 2025 年的 Δ幾乎全在 GB200（(a) +36.4、(b) 基準 +36.9、(b) 低 +17.1，單位 $B）；GB300 在 −6.3 至 +0.2 之間，Hopper −1.1 至 0。也就是說缺口對應的是 **GB200 這個 Analogy 值（以 GB300 的 nvr 套在 Tokenomics GB200 的較低 IT capex 24.0 上）**。
- **限制：** 這只是在「NVIDIA GW 8.05、世代組合不動」下的拆解；2025 年 GB200 占 70% 本身是 Analogy，GW 與組合是另一組不確定，本報告沒有檢驗。

## 4. 還有哪些輸出受影響（LibreOffice 全模型逐格比對；情境 (a)，其餘情境相同範圍）

| 工作表 | 受影響 | 備註 |
|---|---|---|
| Compare | D28:D33（`nvc_*`）；F125:H127（w5 含低／高）；F128:H128（w6）；F132:H139（w8、判定、低高比值）；G143、H145:H147（λ、內容成長、校準後比值）；F148:H159（備忘列）；G165:H187（F3 Siliconomics 並列欄 s5…） | 全部屬對帳點 |
| Interface | **`WIF_Recon_NV_ContentPerGW_2025–2027`、`_ImpliedRev_`、`_Ratio_`、`_Calib_2026`、`_ContentGrowth_2027`、`_CalibRatio_2027`、`_ContentMemoCB26_／CB27_`（2025–2026）、`_RatioMemoCB26_／CB27_`（2025–2026）**，共 20 格 | `ShipGW`、`NVGW`、`ActualRev`、`ReqContent` 不變 |
| Checks | N-K0（C168／E168／F168／H168：2025 帶外 WARN→改入帶則 OK）、N-K1（C169、E169、H169）、WGOV_Warn（C172：9→8） | 燈號與計數 |
| 不受影響（逐格相同） | Players（世代組合）、Deploy、CapexGW、Spend、Gap、Funding、Matrix、Delay；所有玩家資本支出拆分、Spend／Gap 名稱 | 世代組合是**輸入**，不是內容的函數 |

- **兩層基準協調用的換算值 x2：不受影響。** `siliconomics/tools/baseline_bridge.py` 讀 Wattonomics 的名稱只有 `WIF_MktGW_Add_*`、`WIF_Evidence_*（Lo／Hi）`、`WIF_Coverage_*`、`WIF_Spend_it_sum_*`、`WIF_Recon_NV_ShipGW_*`（第 76–126 行），都不在上表變動集合內；`x2`＝`s2 ÷ rt`（第 188 行），不含內容。**改這個參數只動對帳點，不動基準。**（Siliconomics 側：`export_baseline_status.py` 只讀 `WIF_Recon_NV_Ratio／ImpliedRev／ActualRev` 顯示於 c0 實績區，會跟著變，但不進基準判定。）
- **Siliconomics 側的連動未做重算驗證：** 程式搜尋顯示 `nvr_*` 只用在 `Recon D1`（第 1048–1049 行，平台金額＝顆數 × 每顆金額）與 Interface 的 `SIF_NVContentGW_*`（第 1667–1682 行）；`SIF_DeployGW_Base` 等基準名稱的公式沒有引用 `nvr_*`。但我沒有對 Siliconomics 做改值重算，這是以程式搜尋得到的判斷，標為「未驗證」。

## 5. 若要修正：誰決定什麼

**歸屬：** 這個值的作者是 Siliconomics（AI半導體；`nvr` 輸入與 chip_bridge）。Wattonomics 層沒有可改的參數（綠字不得修改）；工作單要寫給 Siliconomics／chip_bridge 層，Wattonomics 只在升版時吸收。依共同規則，會改值者單獨一個 PR、標【會改值】、draft，附預期變動表與回復方案。

**需 Andy 決定的題：**

| 編號 | 決定什麼 | 選項 |
|---|---|---|
| 1 | 對帳點的「NVIDIA 內容」要用哪個定義 | 機架內約 69.5%（現行）／含 LPDDR 與 scale-out×份額（chip_bridge 高值）／含網通的整機架 |
| 2 | GB200 要不要有自己的值（現為 Analogy，70% 權重，是 2025 年缺口主因） | 維持／以 IREN 獨立證據校準／以 NVIDIA 自稱 |
| 3 | Rubin（VR200）每 GW 用 nvr×IT capex，還是每封裝證據×Tokenomics 密度；是否接受 Tokenomics 291,744 的密度 | 見下 |
| 4 | 2025 年對帳單（已開）收在「內容」還是「GW／組合」 | 見下 |

**方法選項（證據強度相對實施成本；用序位與文字）：**

1. **首選：用獨立的買方採購證據（IREN）只校準 Blackwell（GB200／GB300），Rubin 與 RU 暫不動、另標區間。**
   - 證據強度：金額與 MW 為 [Verified]（IREN 新聞稿），NVIDIA 占比 [Assumed] 0.75–0.85；不依賴 NVIDIA 自稱，避開循環；只覆蓋有證據的世代。
   - 實施成本：低。Siliconomics 的 `nvr_GB200` 一個輸入（或新增一個 GB200 覆寫列），不動 Rubin；2025 年比值 0.913／0.968／1.024（低／基準／高）全進帶。
   - 代價：2026 年（GB300 78%）下修到 0.848–0.933，低檔略帶外，因為 IREN 基準 GB300 每封裝 4.76 萬低於 Siliconomics 的 5.11 萬。Rubin 的 2027 年不動，0.645 仍帶外（只作參照）。
2. **次佳：改採 Siliconomics 現成的 `nvr` 高值（定義放寬到含 LPDDR 與 scale-out×份額）。**
   - 證據強度：中低。定義方向與 NVIDIA 自稱（含網通）一致，但 `so_nv` 份額是 [Assumed]，GB200 高值（20.73）仍低於證據。
   - 實施成本：最低（值已存在，不加新輸入）。結果：2025 0.922、2026 1.051、2027 0.723。
3. **第三：整套改成每封裝證據 × Tokenomics 每 GW 封裝數（rubin-price 的 (b) 路徑）。**
   - 證據強度：廣度最高（涵蓋 Blackwell 與 Rubin），但 R1、R2 與 NVIDIA 說法循環、R3 的 460 MW 是二手；而且等於接受 Tokenomics 的密度為正確，並把定義改成含網通。
   - 實施成本：中高。要在 Siliconomics 新增每封裝輸入系統，並與 Tokenomics 協調密度；2027 年的 Rubin 與 RU 無證據，仍要 Assumed。
   - **不列入選項的做法：** 直接採 NVIDIA 的 250／400 億。證據最弱（行銷用語、GW 口徑未定義、拿 NVIDIA 說法對 NVIDIA 實績是循環），成本雖最低，但會讓 (a) 的 2025 年「完全補上」看似成立而實為循環。
   - **維持現狀＋把缺口歸給 GW／覆蓋：** 零成本；但 GB300 證據與現值相符、只有 GB200 明顯偏低，所以現有證據不支持完全歸給 GW。

## 6. 未驗證與限制

- **沒有開原文核對**：NVIDIA 逐字稿的 180／250／400 億、Nscale S-1/A、IREN 新聞稿皆取自 rubin-price 報告與其 xlsx 的輸入，我只核對了它們在本檔的使用，沒有重新讀原文。
- **Blackwell 證據的循環問題**：B1、B2、B3 部分源自 NVIDIA 自己的營收或每 GW 說法，所以 (a)、(b) 補上 2025 年缺口有「用 NVIDIA 說法對 NVIDIA 實績」的成分；只有 (d)（IREN）獨立。
- **GW 口徑**：NVIDIA 的 GW 為 IT 或設施口徑未定義（rubin-price 判斷較接近 IT，[Derived]、未驗證）。
- **2025 年的 NVIDIA 出貨 GW 8.05 與世代組合（GB200 70%）本身未檢驗**；本報告只動內容。
- **Siliconomics 與 Tokenomics 兩層沒有做改值重算**；Siliconomics 側影響僅憑程式搜尋判斷（第 4 節）。
- **世代組合（Players mix）也會影響**加權內容（例如 2027 加回 10% Blackwell 的效果見 W2_A 報告），本報告維持不動。
- RU 無任何證據，以現行結構比 1.3325（Assumed）處理；影響見 3.2。
- 2026 錨點 375.46 含 Q3 指引中點與 Q4 假設；我沒有用 Q3 實績。
- Wattonomics 讀的是 Siliconomics v0.19 快照（現行 Siliconomics 為 v0.20，`SIF_NVContentGW_*` 在 v0.18–v0.20 逐位相同）、Tokenomics v22.0（`tokenomics` repo 另有 v23.0 工作單草稿，未生效）。

## 7. 可寫進記憶的新事實（供主 session 判斷）

- 261 億＝Siliconomics `SIF_NVContentGW_VR200`（0.6951 × TK IF_CapexIT VR200 37.586），Wattonomics 無自有輸入；34.81 是 Rubin Ultra 值，不是 Siliconomics 的 VR200 值（來源：本次對活頁簿的掃描）。
- 改每 GW NVIDIA 內容只動對帳點，不動基準與換算值 x2（來源：LibreOffice 全模型逐格比對）。
