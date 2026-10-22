# 工作單 Anthropic v0.1.3 A 段｜U2 改法、資金三層、持續擴張情境、條件式措辭（2026-10-22，chat 端）

先讀 `20261022_anthropic_v0.1.3_共同.md`。開工：`git checkout -b claude/anthropic-v0.1.3 origin/main`（若分支已存在就 checkout 並 pull），push，開 draft PR（REST：`gh api repos/YinchenChang/company-models/pulls -f title='Anthropic v0.1.3（外部驗收採納）' -f head=claude/anthropic-v0.1.3 -f base=main -F draft=true -f body=...`）。進度檔：`anthropic/docs/reports/20261008_anthropic_進度.md` 新增「v0.1.3」節。

## 步驟（每步 commit＋push＋進度檔）
1. **TK 標記**：TK_Link 快照改讀 v19.0（值應與 v18.0 完全相同；`tools/check_tk_snapshot.py` 結果寫入報告；若有任何 DIFF，停下並在進度檔寫明）。
2. **U2 改法**（Andy 2026-10-22 同意預設）：Nscale、Lambda 的 GW 改用**揭露容量**（Nscale 460 MW、Lambda 350 MW，見 SRC_ANT），「最高 $86B」當付款上限：年付款＝MIN(揭露 GW × 合約價, 86 ÷ 期間 按兩約比例)；報告列前後值（v0.1.2 獨立核對推估：2030 差額約 +11.7、2028 約 −2.72）。
3. **資金三層**（Funding 頁尾端新增一節＋具名範圍）：①營運自由現金流自給（`FND_FCF` 各年、累計、首次轉正年、不靠任何融資時的現金路徑與谷底、缺口累計）；②加已到位股權（現行基準）；③再加條件式與談判中（IPO、Amazon、Google、NVIDIA、AMD）。三層各輸出：最低現金缺口年度、累計外部資金需求、現金谷底。
4. **持續擴張情境**（只作情境，Compute／Cost／Funding 尾端；Inputs 新增：每年追加金額 $B（基準 0；情境值 Assumed，例：2029、2030 各追加等於當年合約付款的 30%，區間 0–60%）、合約價與爬坡沿用既有、變現比例 3 檔：0（無需求）、0.5（部分）、1（完全；新增 GW 的每 GW 營收＝當年基準每 VR 等值 GW 營收淨額）。輸出：每 VR 等值 GW 差額、自由現金流、三層資金結果、轉正年度。與既有 U6（談判中 $150B）並列。
5. **條件式措辭**：`tools/build_html.py` 結論句與 ⑥「要成立必須」改為條件式（例：「在目前需求成長、價格彈性、已簽算力合約與 2026 年已到位融資的假設下，2029 起每 VR 等值 GW 轉正；按目前已到位融資與基準支出路徑，2026–2030 不出現最低現金缺口」），並在 ① 加一行「營運自由現金流本身要到 ____ 年才自給」（取自第 3 步①層）。
6. 重建 `model/20261022_Anthropic_v0.1.3.xlsx`（舊檔 `git mv` 到 archive、更新 CURRENT）；LibreOffice 重算錯誤 0、`CHK_Errors`＝0；parity 情境新增：U2 新法、擴張×3 檔、資金三層各一；pytest 分檔全過。HTML 與 dist 留到 B 段一起產（本段可先產草稿供自檢，不必放 dist）。
7. 進度檔寫清楚 B 段的起點；PR 留言 `[Anthropic 回報] v0.1.3 A｜完成｜日期`，附前後對照（每 VR 等值 GW 差額 2026–2030、覆蓋率、三層資金結果、擴張情境 3 檔）。
