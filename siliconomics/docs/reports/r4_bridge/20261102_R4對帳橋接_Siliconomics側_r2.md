# R4 對帳橋接：Siliconomics 側（r2，2026-11-02；Siliconomics v0.11）

**一句結論：** 數值與 r1 相同（Siliconomics 部署 GW 與 Wattonomics v0.1.7 主路徑都沒變）。依工作方式 v1.7 第 8 節判定：2026、2027 為「時差」（已寄證據送件，待 Wattonomics v0.1.8 納入），2028、2029 為「待診斷」（兩邊都是外推，各自先用自己的數字）。

- 活頁簿：`20261102_R4對帳橋接_Siliconomics側_r2.xlsx`（`siliconomics/tools/r4_bridge.py` 產生、LibreOffice 重算）。讀取版本：Siliconomics v0.11、Wattonomics v0.1.7、Tokenomics v21.7。
- 本次起每次 Siliconomics 升版都跑；結果寫入 Recon 頁 B 節（經 `data/recon_snapshot_v0.11.json`）。
- 工具修改：Wattonomics 停用需求情景與延遲情景後，找不到「A、D0」欄時改讀同列名的基準列。

## Facts & Data

| | 2026 | 2027 | 2028 | 2029 |
|---|---|---|---|---|
| Siliconomics 推得淨新增上線 IT GW（情境 A） | 20.30 | 35.42 | 51.39 | 60.92 |
| Wattonomics 主路徑（A、D0） | 12.52 | 23.96 | 26.92 | 27.66 |
| 比值（Wattonomics ÷ Siliconomics） | 0.617 | 0.677 | 0.524 | 0.454 |
| 對照：Wattonomics 側歸因調整後比值（算力收支 10-31，尚未進入路徑） | 0.934 | 0.719 | 0.558 | 0.445 |
| 證據送件 | 10-31 R4 結果（觀察 2） | 11-02 NVIDIA FY2028 指引 | — | — |
| **狀態** | **時差** | **時差** | **待診斷** | **待診斷** |

## 2027 年新證據（11-02 送件）

- NVIDIA FY27 Q2 法說（2026-08-26）：FY2028（約 CY2027）營收預期年增約 70%，且為供給受限下的預估 [Interested-party：公司指引]。以 FY2027 共識約 3,960–4,100 億美元換算，FY2028 約 6,730–6,970 億；資料中心占比 92.5%（Q2 實績）→ 約 6,200–6,450 億 [Derived]。
- Siliconomics 2027 年 NVIDIA 顆數（Blackwell 2.4、Rubin 7.0、Rubin Ultra 0.3 百萬顆）× 2026 年每顆平均資料中心營收約 4.92 萬美元 × Rubin 世代溢價 1.2–1.5 倍 [Assumed] → 約 5,500–6,600 億，對公司指引比值 0.85–1.06。
- 敏感度：溢價取 1.2 時比值 0.85–0.89；取 1.5 時 1.02–1.06。整個區間都沒有低到 Wattonomics 2027 年水準所需的約 0.7。
- 判讀：2027 年 Siliconomics 的 NVIDIA 顆數受公司指引支持；R4 2027 年殘差較可能在 Wattonomics 側（七家資本支出或非七家買方的覆蓋），與算力收支 10-31 歸因第 4 節的兩種可能一致。是否成立，待 Wattonomics v0.1.8 納入後再對帳。

## 2028–29 與基準取向

- Siliconomics 2028 年起的 Rubin、Rubin Ultra 與 ASIC 顆數都是 Assumed（v0.11 已把 `u_rb_28`、`u_bw_28` 改標 Assumed 並給區間）；Wattonomics 為資本支出外推。兩邊都沒有硬證據，依 Andy 11-02 16:22「各層先用自己的數字、較上游優先更新」，兩邊基準都不動。
- 若 Siliconomics v0.12 只留基準且基準取 B（決策 S33，待 Andy），部署 GW 2028／2029 由 52.8／60.1 降為 42.2／51.1，殘差約縮小 10／9 GW，仍屬待診斷。
- 監看：AMD 3Q26 法說（11-03）、NVIDIA FY27 Q3（11-18）與 FY27 Q4 指引、四家雲端 2027 年資本支出指引（2027-01～02）。
