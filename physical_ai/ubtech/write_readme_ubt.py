# 由重算後成品產生 physical_ai/ubtech/README.md（數字一律讀自 xlsx，不手寫；原則第 10 節第 13、51 條）。build_ubtech.py 呼叫。
import os
import openpyxl

HERE = os.path.dirname(os.path.abspath(__file__))


def write(path, meta, sens, checks):
    wb = openpyxl.load_workbook(path, data_only=True)
    V, P, RD, I = wb["Val"], wb["PL"], wb["README"], wb["In"]
    VR, PR = meta["VR"], meta["PR"]
    g = lambda k, c="B": V[f"{c}{VR[k]}"].value
    IN = {I.cell(r, 1).value: r for r in range(4, meta["LAST_IN"] + 1) if I.cell(r, 1).value}
    e = lambda k: I[f"E{IN[k]}"].value
    f2 = lambda x: f"{x:,.2f}"; pc = lambda x: f"{x * 100:+.1f}%"; p1 = lambda x: f"{x * 100:.1f}%"
    name = os.path.basename(path)
    YRS = list(range(2023, 2033)); COL = {y: "CDEFGHIJKL"[i] for i, y in enumerate(YRS)}
    def yr(k, s="B", fmt="{:,.2f}"):
        out = []
        for y in YRS:
            v = P[f"{COL[y]}{PR[s][k]}"].value
            out.append(fmt.format(v) if isinstance(v, (int, float)) else "—")
        return " | ".join(out)
    base = sens["base_tp"]
    scn = "\n".join(f"| 情境 {s} | {f2(g('dcfr', c))} | {f2(g('cu', c))} | {f2(g('tps', c))} |" for s, c in zip("ABC", "BCD"))
    top = sorted(sens["cases"], key=lambda c: -abs(c["tp"] / base - 1))[:8]
    stbl = "\n".join(f"| {c['drv']} | `{c['key']}` | {f2(c['tp'])} | {pc(c['tp'] / base - 1)} | {f2(c['dcf'])} | {c['rt']} |" for c in top)
    rts = sorted(set(c["rt"] for c in sens["cases"]))
    flips = [c for c in sens["cases"] if c["rt"] != g("rt")]
    ftbl = "\n".join(f"| {c['drv']} | `{c['key']}` | {f2(c['tp'])} | {pc(c['tp'] / base - 1)} | {c['rt']} |" for c in flips) or "| （無） | | | | |"
    rv = "\n".join(f"| {r['label']} | `{r['key']}` | {r['val'] if isinstance(r['val'], str) else p1(r['val'])} | {p1(r['base'])} | {r['note']} |" for r in sens["revs"])
    RCs = wb["Recon"]
    imp = [(RCs.cell(r, 1).value, RCs.cell(r, 2).value, RCs.cell(r, 4).value) for r in range(1, RCs.max_row + 1)
           if isinstance(RCs.cell(r, 1).value, str) and RCs.cell(r, 1).value.startswith("反解：") and isinstance(RCs.cell(r, 2).value, (int, float))]
    imptxt = "；".join(f"{a[3:]}：情境 B FY2027 全尺寸人形銷量約 {b:,.0f} 台（公司 2027 年產量目標 1 萬台的 {c:.2f} 倍）" for a, b, c in imp)
    ctbl = "\n".join(f"| `{n}` | {'通過' if rc == 0 else '未通過'}（結束碼 {rc}） | {s.replace('|', '／')} |" for n, s, rc in checks)
    txt = f"""# physical_ai/ubtech — 優必選 UBTech（9880.HK）個股模型 {VER(RD)}

{RD['A2'].value}｜build_id `{sens['build_id']}`｜成品 `dist/{name}`（LibreOffice 重算，錯誤值 0）

港交所主板上市（2023-12）。報表人民幣（中國企業會計準則，港交所公告同一口徑），股價與目標價港元（匯率取 csv 的 HKD=X、CNY=X）。所屬線：線 4 人形機器人（`physical_ai/lines/humanoid/`）。本 README 的數字由 `build_ubtech.py` 從重算後的成品讀出產生，不手寫。

**本模型沒有做的部分（證據邊界）：** 分類別毛利率公司未揭露，全尺寸人形毛利率以殘差法推得（Derived，區間 45–65%）；2026 年新訂單總額未找到（只有個別訂單公告），下半年出貨與 FY2027 起占比全為 Assumed；相對人工生產力沒有已驗證實例，模型不以生產力推需求，只用線模型全球出貨 × 占比；未建融資模組（收支缺口以 Val 頁「缺口 ÷ 淨現金」提示）；季度分部頁不做（公司只發半年報）。

## 命題
{RD['A3'].value}

**結論（條件式）：** {RD['A4'].value}

{RD['A5'].value}

**一句話：** 優必選是全尺寸人形整機的純標的：2026 上半年全尺寸人形營收 {f2(wb['Hist']['F4'].value)} 十億元、交付 {e('u_h126'):,.0f} 台（每台約 {e('asp_h1'):,.0f} 千元），已占營收約一半；但全公司仍虧損，營運資金與資本支出吃掉現金。在「出貨以已交付量與已公告訂單為錨」的基準下，模型營收（情境 B FY2027 {f2(P[f"G{PR['B']['rev']}"].value)} 十億元）明顯低於賣方共識（{f2(e('cons_rev27'))}），這是目標價低於現價的主因；倍數本身並不貴（現價對共識營收的 EV／營收約為系統與整機同業中位數的 {g('self'):.2f} 倍）。

## 估值（評價日 2026-12-31；目標時點約 12 個月；港元）
| 情境 | DCF（推到 12 個月後） | 分部 EV／營收加總 | 情境目標價 |
|---|---|---|---|
{scn}

- 加權目標價 **{f2(g('tp'))} 港元**（DCF {f2(g('dcf'))}、分部 EV／營收加總 {f2(g('cmp'))}），現價 {f2(e('px'))}（2026-11-20 收盤），上檔 {pc(g('up'))}，估值評等 **{g('rt')}**；收支關卡 {g('gj')}（情境 B FY2026–30 自由現金流扣股利 {f2(g('gate', 'C'))} 十億元，為 2026-06-30 淨現金的 {g('cov'):.2f} 倍）→ 投資建議 **{g('rec')}**。
- 方法：DCF 標準版 FCFF（WACC／ke／g／RONIC＝{g('wk')}；β 1.5 為 Assumed，資料商 5 年 β 3.18 因上市未滿 3 年不採用）50%＋分部 EV／營收加總 50%（全尺寸與非具身人形 × 系統與整機群組〔越疆、地平線〕中位數 {e('ms_sys'):.2f} 倍；其他類別 × 零組件與物流群組〔極智嘉、速騰聚創〕中位數 {e('ms_cmp'):.2f} 倍；FY2028 營收，加目標時點淨現金）。相對倍數係數 1.00（未調整：缺歷史相對倍數）。
- 情境 B 與 C 的 DCF 每股價值為零：明確期自由現金流全為負、終值不足以抵銷，股權價值為負，依有限責任取零；DCF 加權值全部來自情境 A。
- 隱含 EV／營收（加權目標價 ÷ 情境 B FY2028 營收）{g('evs'):.2f} 倍；三段歸因：DCF {pc(g('a1'))}、未調整 Comparable {pc(g('a2'))}、相對倍數調整 {pc(g('a3'))}。

## 營收分部（公司按產品類別揭露的營收；PL 頁三情境，情境 B 摘要，十億元人民幣）
| 列 | {" | ".join(f"FY{y}" for y in YRS)} |
|---|{"---|" * len(YRS)}
| 全尺寸人形銷量（台） | {yr('u', fmt='{:,.0f}')} |
| 每台營收（千元） | {yr('asp', fmt='{:,.0f}')} |
| 全尺寸具身智能人形 | {yr('fs')} |
| 非具身智能人形 | {yr('ne')} |
| 其他智能機器人 | {yr('or')} |
| 其他智能硬件 | {yr('hw')} |
| 園林機械、汽車及液壓零部件（鋒龍） | {yr('fl')} |
| 其他 | {yr('ot')} |
| 合併營收 | {yr('rev')} |
| 毛利率 | {yr('gm', fmt='{:.1%}')} |
| 營業利益 | {yr('op')} |
| 歸屬母公司淨利 | {yr('ni')} |
| EPS（元） | {yr('eps')} |

- FY2023–FY2025 實績（FY2023 類別拆分未取得）；FY2026＝上半年實績＋下半年假設；FY2027 起為模型。
- 命題模組：全尺寸人形銷量＝線模型全球人形出貨（情境 A 讀線高、B 讀基準、C 讀低）× 優必選占比；每台營收＝2026 上半年每台營收 × 線模型全尺寸售價指數（同欄）。
- 營業利益為本模型口徑＝稅前 − 業外（財務淨額＋其他收益）；營業費用含銷售、管理、研發、稅金及附加、投資收益與減值。

## 主要假設（In 頁；Assumed／Analogy 皆給區間）
- 2026 下半年全尺寸人形銷量（情境 A／B／C）：{e('u26h2_A'):,.0f}／{e('u26h2_B'):,.0f}／{e('u26h2_C'):,.0f} 台（上半年 {e('u_h126'):,.0f} 台）；每台營收相對上半年 {pc(e('asp_adj'))}（−15%～+10%）。
- FY2027 占全球人形出貨 {p1(e('shr27_A'))}／{p1(e('shr27_B'))}／{p1(e('shr27_C'))}，之後每年 {e('dshr_A')*100:+.2f}／{e('dshr_B')*100:+.2f}／{e('dshr_C')*100:+.2f} pt；全尺寸毛利率 FY2026 {p1(e('gm_fs26'))}（殘差法）→ FY2030 {p1(e('gm_fslt_A'))}／{p1(e('gm_fslt_B'))}／{p1(e('gm_fslt_C'))}。
- 其他類別毛利率 {p1(e('gm_leg'))}、鋒龍 {p1(e('gm_fl'))}；核心營業費用 FY2026 {f2(e('op26'))} 十億元、之後年增 {p1(e('gop_A'))}／{p1(e('gop_B'))}／{p1(e('gop_C'))}；減值 ÷ 營收 {p1(e('imp_r'))}。
- 現金流：折舊 {p1(e('fc_da'))}、資本支出 FY2026 {f2(e('cx26'))} 十億元後 {p1(e('fc_cx'))}、營運資金 {p1(e('fc_wc'))} × 營收增量；稅率 {p1(e('tax_w'))}（只對正稅前）。
- 折現：rf {e('rf')*100:.2f}%（中國 10 年期，跨模型一致）、ERP 5%、CRP 0.91%、g＝MIN(3%, rf)、RONIC＝WACC。評價日股數 {e('sh_v'):,.2f} 百萬股（RSU／選擇權未行使數量未取得）。

## 承重 Assumed 驅動與可證偽指標
| 驅動 | 公開可觀察指標 |
|---|---|
| 下半年全尺寸人形銷量（u26h2） | 2026 年報（2027-03）全尺寸人形銷量與營收；公司訂單公告累計 |
| 優必選占全球人形出貨（shr27、dshr） | 機構季度出貨統計（Omdia、TrendForce）；宇樹、智元、Tesla 出貨 |
| 全尺寸人形毛利率（gm_fs26、gm_fslt） | 半年報毛利率（全尺寸營收占比上升時）；每台營收趨勢 |
| 核心營業費用年增（gop） | 半年報銷售、管理、研發費用 |
| 營運資金率（fc_wc） | 應收帳款與存貨對營收比 |

失效條件：{RD['A6'].value.split('：', 1)[1]}。

## 敏感度（整本重算；Sens 頁；影響最大的 8 項）
| 驅動 | 鍵 | 加權目標價 | 相對基準 | DCF 目標價 | 估值評等 |
|---|---|---|---|---|---|
{stbl}

全部 {len(sens['cases'])} 案見 xlsx Sens 頁；各案例估值評等出現「{'、'.join(rts)}」。評等翻轉的案例：

| 驅動 | 鍵 | 加權目標價 | 相對基準 | 估值評等 |
|---|---|---|---|---|
{ftbl}

## 反解（整本重算，格點掃描＋線性內插；Recon 頁）
| 反解目標 | 鍵 | 反解值 | 基準 | 說明 |
|---|---|---|---|---|
{rv}

反解占比換算銷量（Recon 頁）：{imptxt}。反解值是「目標價＝現價」或評等邊界的點，不是評等門檻本身；其他輸入不變。

## 對帳（線層級變數；xlsx「對帳」頁）
讀自 `physical_ai/lines/humanoid/dist/20261116_人形機器人單位經濟_v0.1.xlsx`「下游取數」：全球人形出貨 `SHIP_2026`–`SHIP_2032`（C34:E40；情境 A／B／C 讀高／基準／低）、全尺寸售價 `HUM_PRICE_2026`（D10）、`HUM_PRICE_2028`／`2030`／`2035`（C77:E77、C79:E79、C81:E81；年度間幾何內插）；`HUM_BOM_2026`、`HUM_MFG_2026`、`HUM_PROD_REQ` 只作對照。與線模型不一致的自設值：全尺寸毛利率（FY2026 {p1(e('gm_fs26'))}、FY2030 起情境 B {p1(e('gm_fslt_B'))}）高於線模型定價毛利率 25%——線模型為出廠硬體定價假設，優必選營收含解決方案與服務，且以殘差法推得；2026 上半年每台營收換算美元為線模型 2026 出廠售價的 {wb['對帳'][f"D{[r for r in range(1, wb['對帳'].max_row + 1) if str(wb['對帳'].cell(r, 1).value).startswith('2026H1 每台營收')][0]}"].value:.2f} 倍（量級一致）。

## 限制與未取得資料
- **2026 年新訂單總額未找到**（2025 年訂單約 14 億元為媒體轉述公司說法；2026 年只有個別公告：Airbus、海外逾 5,000 萬元、U1 訂單逾 13,361 台）。
- 分類別毛利率、應報導分部的正式劃分、RSU／選擇權未行使數量、租賃負債、2026H1 現金流量表：未取得。FY2023 產品類別營收、折舊、資本支出：未取得。
- 同業只有 4 家（每群組 2 家），股數與淨現金取資料商 2026-09～10 頁面（未回原文核對），股價為 2026-11-20 收盤；不做迴歸。賣方共識（營收、EPS、目標價）資料日 2026-09-03，股價已自 77.40（2026-09-16）跌至 {f2(e('px'))}。
- 鋒龍股份（43%）以並表營收 × 零組件群組倍數計價、少數股東權益按帳面扣除；按 A 股市值計的持股價值未納入（Risk 頁）。
- 全公司 FY2026 所得稅、業外為 Assumed 定值；FY2027 起虧損不認列稅盾。
- 數字尚未經未參與製作的代理獨立核對（Project 規則第 10 條，待主代理安排）。

## 核對（建置腳本自動執行；結果摘要）
| 核對 | 結果 | 輸出摘要 |
|---|---|---|
{ctbl}

## 檔案
- `build_ubtech.py`：建置＋整本重算 Sens／反解＋產生 `sens.json`、`segments.json`、HTML 與本 README，並跑上表核對。`write_readme_ubt.py`：本 README 產生器。
- `dist/{name}`：成品（工作表 README、In、Hist、PL、DCF、Val、Recon、Sens、對帳、Risk、Sources）。
- `html.json`＋`gen_html_json.py`：HTML 展示層設定（tools/company_html，layout v1.3）；`dist/{name[:-5]}.html`。

## 版本紀錄
- v0.1（2026-11-23）：初版；現價 {f2(e('px'))} 港元（2026-11-20 收盤）；加權目標價 {f2(g('tp'))}、上檔 {pc(g('up'))}、估值評等 {g('rt')}、投資建議 {g('rec')}。
"""
    open(os.path.join(HERE, "README.md"), "w", encoding="utf-8").write(txt)


def VER(RD):
    a1 = RD["A1"].value
    return a1.split("個股模型 ")[1].split("｜")[0]
