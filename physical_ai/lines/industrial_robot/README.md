# 線 2a 工業機器人（含協作）單位經濟模型

AI實體應用 Project 底層五條線之一。依《線2a_工業機器人_框架_v0.1.md》（2026-11-16）建成 Excel，供上層個股模型（哈默納科、Nabtesco、Fanuc、安川、上銀等）以 ID 取數並互相對帳。

| 版本 | 日期 | 成品 |
|---|---|---|
| v0.1 | 2026-11-16 | `dist/20261116_工業機器人單位經濟_v0.1.xlsx` |

## 建置與核對

```
python3 physical_ai/lines/industrial_robot/build_industrial_robot.py
python3 physical_ai/lines/industrial_robot/check_industrial_robot.py   # LibreOffice 重算（覆寫為含快取值）、0 錯誤值、對帳、情境單調、下游 ID 唯一
```

## 頁面

說明｜輸入｜計算｜情境｜敏感度｜下游取數｜追蹤指標

- **輸入**：藍字＝輸入，黃底＝判斷值。每列有標記（[Verified]／[Interested-party]／[Analogy]／[Assumed]／[Derived]）、類別（[P]／[A]／[I]）、出處。Analogy、Assumed 一律給低／基準／高。
- **計算**：A 單台回本（國產 6 軸、外資 6 軸、國產協作；系統投資＝本體×(1＋整合倍數)，回本期＝系統投資÷年淨節省）；B 本體成本結構（減速機價值由上而下對由下而上）；C 每台減速機用量（混合口徑）；D 產業層對帳（模型 RV、諧波市場對 Nabtesco CMP、HDS 營收推得的上限；IFR 每台均價；國家統計局產量對 IFR 裝機）；E 循環層（JARA 訂單÷生產）。
- **情境**：2025＝IFR 實績；2026–2032 低／基準／高。成長、中國占比、協作占比、減速機均價依情境取值；機型結構與每台顆數三情境都用基準（不確定性放在敏感度）。
- **敏感度**：兩個輸出（2030 全球機器人減速機市場、國產 6 軸回本期），一次動一個變數，依擺幅排序；基準值與情境、計算頁自動對帳。
- **下游取數**：長表，以 ID 取數（不要用列號）。
- **追蹤指標**：JARA 季度（會員口徑）、最新讀數（含日期與出處）、框架第 5 節頻率表。

## 下游取數 ID（主要）

| ID 樣式 | 內容 | 單位 |
|---|---|---|
| `IR.inst.global.<年>.<lo/base/hi>` | 全球工業機器人裝機（含協作，IFR 口徑） | 台 |
| `IR.inst.china.<年>.<情境>` | 中國裝機 | 台 |
| `IR.inst.cobot.<年>.<情境>`、`IR.share.cobot.…` | 協作裝機、協作占比 | 台、% |
| `IR.inst.trad.<年>.<情境>` | 非協作裝機 | 台 |
| `IR.red.rv.units.…`、`IR.red.hd.units.…`、`IR.red.hd.cobot_units.…` | RV、諧波需求（僅工業機器人；含替換係數） | 顆 |
| `IR.red.rv.mkt.…`、`IR.red.hd.mkt.…`、`IR.red.total.mkt.…` | 機器人減速機市場（全球混合均價） | USD m |
| `IR.red.rv.units.china.…`、`IR.red.hd.units.china.…` | 中國 RV、諧波需求（假設同全球結構） | 顆 |
| `IR.mv.global.…` | 全球本體市場值 | USD bn |
| `IR.unit.<N_RV_ART／N_HD_ART_RV／N_HD_SMALL／N_HD_SCARA／N_HD_COBOT／RVTYPE_SH>.<情境>` | 每台物理用量 | 顆/台 |
| `IR.unit.U_RV_TRAD`、`IR.unit.U_HD_TRAD`、`IR.unit.U_RV_ALL25`、`IR.unit.U_HD_ALL25` | 混合每台用量 | 顆/台 |
| `IR.payback.cn6ax`／`for6ax`／`cncobot` | 回本期（基準） | 年 |
| `IR.cycle.jara.*` | JARA 最新季受注、YoY、生產、訂單÷生產、近四季、2025 全年、2026 協會預測 | 百萬日圓、%、倍 |
| `IR.cycle.cn_prod.ytd`、`.ytd_yoy` | 中國工業機器人產量 1–9 月 | 台、% |
| `IR.share.china_domestic.2025` | 中國裝機國產比 | % |

2025 年為實績（情境欄寫「實績」，只有 `.base`）。

## 範圍與口徑

- 減速機市場只含工業機器人（含協作）用途；不含半導體設備、人形（線 2b）、行星與擺線以外產品。
- 中國產量（國家統計局）約為 IFR 中國裝機 2 倍（含出口、計數口徑），只看增速。
- JARA 季度為會員口徑；JARA 全年預測含非會員。
- 公開產業報告的減速機市場規模互相矛盾，本模型改由下而上推算，並以 Nabtesco、HDS 營收推得的上限對帳。

## 未決

- 整合成本倍數、每班替代人數：無公開統計，回本期只能當區間讀。
- 減速機 2025 年均價：以 2022 年中國均價為錨的判斷值。
- 多關節中 RV 型占比、SCARA 每台諧波數：未取得公開統計。
- HDS 2026 年 7–9 月受注、Nabtesco 精密減速機分部訂單：未取得。
- 中國 2026 年 10 月產量：本輪未取得。
