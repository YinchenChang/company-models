#!/usr/bin/env python3
"""C3：Wattonomics Spend 頁 IT 分項與 Siliconomics 快照同年市場值並排（唯讀對照，不連結、不進任何計算）。
輸出 wattonomics/docs/reports/20261118_C3_支出分項對Siliconomics市場值.{xlsx,md}。不改任何活頁簿或快照。"""
import openpyxl, json, os, subprocess, shutil
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter as L
ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), '..', '..', '..'))
MODEL = os.path.join(ROOT, 'wattonomics', 'model', open(os.path.join(ROOT, 'wattonomics', 'model', 'CURRENT')).read().strip())
SNAPN = open(os.path.join(ROOT, 'siliconomics', 'data', 'CURRENT_SNAPSHOT')).read().strip()
SNAP = json.load(open(os.path.join(ROOT, 'siliconomics', 'data', SNAPN)))['names']
OUT = os.path.join(ROOT, 'wattonomics', 'docs', 'reports', '20261118_C3_支出分項對Siliconomics市場值.xlsx')
wv = openpyxl.load_workbook(MODEL, data_only=True)
sp = wv['Spend']
srow = {sp.cell(r, 2).value: r for r in range(1, sp.max_row + 1) if sp.cell(r, 2).value}
YEARS = [2026, 2027, 2028, 2029]
def spv(key, y): return sp.cell(srow[key], 4 + (y - 2023)).value
def sil(name, y):
    a = SNAP[name]
    i = y - 2025 if len(a) == 5 else y - 2026
    return a[i]
BLUE = Font(color='0000FF'); BOLD = Font(bold=True); HDR = PatternFill('solid', fgColor='DDEBF7'); GREY = Font(color='808080')
# (Spend 鍵, Spend 名, [Siliconomics 名稱...], 對應性, 口徑說明)
ITEMS = [
 ('it_compute', '運算', ['SIF_GpuAsicMkt_Base'], '部分對應', 'Spend＝IT capex × 固定比例 51.6%（GB300 對照，ASIC 近似），上線年；Siliconomics＝Σ 平台基準顆數 × 每顆金額（Recon D1，已建模個股加總＋殘差），出貨年，供給側顆數讀數（2027 為 39.2 GW 成本等值，Wattonomics 上線年為 25.5 GW）。兩者差距含 GW 基準差與比例假設。'),
 ('it_memory', '記憶體', ['SIF_HBM_Mkt_Base'], '部分對應（Siliconomics 只有 HBM）', 'Spend 記憶體＝固定比例 13.5%，含 HBM 以外的記憶體與儲存媒體；Siliconomics 只列 HBM 市場。HBM 單項已高於 Spend 整個記憶體，代表 Spend 比例（GB300 時期）低估記憶體價格上漲後的占比，或兩邊 GW 基準不同。'),
 ('it_network', '網通', [], '無對應市場', 'Siliconomics 目前沒有網通／光模組市場（r2 工作單 C2 規劃 HBM → PCB／CCL → CoWoS → 光模組 → 電源、液冷，光模組尚未建）。'),
 ('it_foundry', '代工封測', [], '無可比市場', 'Siliconomics 只有 SEMI 設備支出（SIF_SEMI_FoundryLogic）與 SoC 測試機市場（SIF_TestSoCMkt_Base），都是設備端，不是代工封測營收，不並排。'),
 ('it_substrate', '載板與 PCB', ['SIF_PCBMkt_Base'], '口徑差異大', 'Spend 比例 0.3% 為 [Assumed]（SegIn：GB300 對照）；Siliconomics AI 高速 PCB 市場（不含價格指數，另有上游 CCL 市場 SIF_CCLMkt_Base 為其原料，不加總）。Spend 項目名稱為「載板與 PCB」，但 SegIn 未說明 0.3% 的涵蓋範圍；可能只含 ABF 載板、未含 AI 板（主板、交換器板、OAM），後者若落在伺服器組裝殘差則兩邊口徑不同（待 AI半導體 確認，本表不下結論）。'),
 ('it_power_thermal', '電源散熱（機架內）', ['SIF_PwrMkt_Base', 'SIF_LiqMkt_Base'], '部分對應（電源＋液冷加總）', 'Spend＝固定比例 6.3%（含機構、整合）；Siliconomics 電源市場與液冷市場分列，此處加總對照。機架外（廠房）電源與冷卻在 Spend 的 Facility 項，不在此列。'),
 ('it_server', '伺服器組裝', ['SIF_AsmMkt_Base'], '口徑差異大', 'Spend 為 IT capex 的殘差（SegIn：未拆），不含已計入其他項的零組件；Siliconomics AI 伺服器組裝市場為 ODM 價（整機營收，含 GPU 等零組件價值），量級完全不同，只作參照。'),
 ('it_storage', '儲存與管理', [], '無對應市場', 'Siliconomics 無儲存市場。'),
]
wb = openpyxl.Workbook(); w0 = wb.active; w0.title = '說明'
ws = wb.create_sheet('C3_並排')
ws.cell(1, 1, 'C3 支出分項並排：Wattonomics Spend 頁 IT 分項 對 Siliconomics 快照市場值（US$B；唯讀對照欄）').font = Font(bold=True, size=12)
ws.cell(2, 1, f'Wattonomics：{os.path.basename(MODEL)} Spend 頁（上線年口徑）。Siliconomics：{SNAPN}（出貨年口徑）。藍字＝讀入值，黑字＝公式。灰字行＝無對應市場，只列 Spend 值。')
hdr = ['Spend 項目', '列別', '來源'] + [str(y) for y in YEARS] + ['對應性', '口徑說明']
r = 4
for j, h in enumerate(hdr, 1):
    c = ws.cell(3, j, h); c.font = BOLD; c.fill = HDR
YC = {y: 4 + i for i, y in enumerate(YEARS)}   # D..G
# 總量區塊
tot_rows = {}
def put_row(label, kind, src, vals=None, formula=None, fmt='#,##0.0', blue=True, grey=False):
    global r
    ws.cell(r, 1, label); ws.cell(r, 2, kind); ws.cell(r, 3, src)
    for y in YEARS:
        c = ws.cell(r, YC[y])
        if vals is not None:
            v = vals[y]; c.value = v
            if blue and v is not None: c.font = BLUE
        else: c.value = formula(y)
        c.number_format = fmt
        if grey: c.font = GREY
    r += 1; return r - 1
ws.cell(r, 1, '■ 總量與基準分解').font = BOLD; r += 1
rg_w = put_row('年新增 GW（Wattonomics：上線年主路徑）', 'GW', 'Spend!行4', {y: spv('', y) if False else sp.cell(4, 4 + y - 2023).value for y in YEARS}, fmt='#,##0.00')
rg_s = put_row('出貨年 GW（Siliconomics：成本等值）', 'GW', 'SIF_DeployGW_Base', {y: sil('SIF_DeployGW_Base', y) for y in YEARS}, fmt='#,##0.00')
rg_r = put_row('GW 比（Wattonomics ÷ Siliconomics）', '比', '公式', formula=lambda y: f'={L(YC[y])}{rg_w}/{L(YC[y])}{rg_s}', fmt='0.000')
pg_w = put_row('每 IT GW IT capex（Wattonomics）', 'US$B/GW', 'Spend!行5', {y: sp.cell(5, 4 + y - 2023).value for y in YEARS}, fmt='#,##0.00')
pg_s = put_row('每 GW IT capex（Siliconomics）', 'US$B/GW', 'SIF_ITCapexPerGW', {y: sil('SIF_ITCapexPerGW', y) for y in YEARS}, fmt='#,##0.00')
pg_r = put_row('單價比（Wattonomics ÷ Siliconomics）', '比', '公式', formula=lambda y: f'={L(YC[y])}{pg_w}/{L(YC[y])}{pg_s}', fmt='0.000')
it_w = put_row('IT 設備小計（Wattonomics Spend）', 'US$B', 'Spend!行16', {y: sp.cell(16, 4 + y - 2023).value for y in YEARS})
it_s = put_row('AI IT 資本支出（Siliconomics）', 'US$B', 'SIF_AIITCapex_Base', {y: sil('SIF_AIITCapex_Base', y) for y in YEARS})
it_r = put_row('IT 總量比（Wattonomics ÷ Siliconomics）＝GW 比 × 單價比', '比', '公式', formula=lambda y: f'={L(YC[y])}{it_w}/{L(YC[y])}{it_s}', fmt='0.000')
put_row('檢查：GW 比 × 單價比 − IT 總量比（須≈0）', '比', '公式', formula=lambda y: f'={L(YC[y])}{rg_r}*{L(YC[y])}{pg_r}-{L(YC[y])}{it_r}', fmt='0.0000')
r += 1
ws.cell(r, 1, '■ 分項並排').font = BOLD; r += 1
blocks = {}
for key, nm, names, rel, note in ITEMS:
    ws.cell(r, 1, nm).font = BOLD; ws.cell(r, 8, rel); ws.cell(r, 9, note); ws.cell(r, 9).alignment = Alignment(wrap_text=True, vertical='top')
    rw = put_row('  Wattonomics Spend', 'US$B', f'Spend!{key}', {y: spv(key, y) for y in YEARS}, fmt='#,##0.00', grey=not names)
    rsh = put_row('  占 Wattonomics IT 小計', '%', '公式', formula=lambda y, rw=rw: f'={L(YC[y])}{rw}/{L(YC[y])}{it_w}', fmt='0.0%', grey=not names)
    blocks[key] = {'w': rw, 'sh': rsh}
    if names:
        rs = []
        for n in names:
            rs.append(put_row(f'  Siliconomics {n}', 'US$B', n, {y: sil(n, y) for y in YEARS}, fmt='#,##0.0'))
        if len(rs) > 1:
            rsum = put_row('  Siliconomics 合計', 'US$B', '公式', formula=lambda y, rs=rs: '=' + '+'.join(f'{L(YC[y])}{x}' for x in rs), fmt='#,##0.0')
        else: rsum = rs[0]
        rdiff = put_row('  差距（Spend − Siliconomics）', 'US$B', '公式', formula=lambda y, rw=rw, rsum=rsum: f'={L(YC[y])}{rw}-{L(YC[y])}{rsum}', fmt='#,##0.0')
        rrat = put_row('  比值（Spend ÷ Siliconomics）', '比', '公式', formula=lambda y, rw=rw, rsum=rsum: f'={L(YC[y])}{rw}/{L(YC[y])}{rsum}', fmt='0.000')
        rsh2 = put_row('  Siliconomics 占 Siliconomics IT 資本支出', '%', '公式', formula=lambda y, rsum=rsum: f'={L(YC[y])}{rsum}/{L(YC[y])}{it_s}', fmt='0.0%')
        rsc = put_row('  Spend 以 Siliconomics IT 總量重新縮放（Spend 占比 × Siliconomics IT）', 'US$B', '公式（僅供對照）', formula=lambda y, rsh=rsh: f'={L(YC[y])}{rsh}*{L(YC[y])}{it_s}', fmt='#,##0.0')
        rrat2 = put_row('  縮放後比值（只剩占比假設的差異）', '比', '公式', formula=lambda y, rsc=rsc, rsum=rsum: f'={L(YC[y])}{rsc}/{L(YC[y])}{rsum}', fmt='0.000')
        blocks[key].update({'s': rsum, 'diff': rdiff, 'rat': rrat, 'ssh': rsh2, 'sc': rsc, 'rat2': rrat2})
    r += 1
for j, w in enumerate([46, 9, 24, 11, 11, 11, 11, 24, 100], 1): ws.column_dimensions[L(j)].width = w
ws.freeze_panes = 'D4'
lines = ['C3 支出分項並排（唯讀；不連結、不進任何計算）', '',
 '為什麼是獨立報告而不是 Compare／Spend 頁的對照欄：Wattonomics 的 Siliconomics 取數只限基準名稱（SIL_In；README 與規則 v1.4.3：上游不讀下游係數，唯一例外是 baseline_status.json 只作顯示）。Siliconomics 市場值（SIF_PCBMkt_Base 等）是下游的輸出，新增進 SIL_In 或活頁簿對照欄等於新增一類上游對下游的讀取，且要升版重建、牽動快照與 Checks。所以本張只讀快照 JSON 與 v0.3.6 活頁簿的已重算值，寫成報告，不動任何模型。',
 '年份對齊：Wattonomics Spend 為上線年；Siliconomics 市場值為出貨年。兩者同標籤年份不同時點（lag_on 4.5 個月、inv_on 10%）。', '快照陣列：4 個值者為 2026–2029；5 個值者（HBM 市場）為 2025–2029。',
 '「縮放後比值」列把 Spend 分項用 Siliconomics 的 IT 總量重新縮放，剩下的差異只來自占比假設，用來區分「總量不同」與「組成不同」。']
for i, t in enumerate(lines, 1): w0.cell(i, 1, t).alignment = Alignment(wrap_text=True)
w0.column_dimensions['A'].width = 150
tmp = OUT + '.raw.xlsx'; wb.save(tmp)
od = os.path.dirname(OUT)
subprocess.run(['soffice', '--headless', '--calc', '--convert-to', 'xlsx:Calc MS Excel 2007 XML', '--outdir', os.path.join(od, '_tmp_c3'), tmp], check=True, stdout=subprocess.DEVNULL)
shutil.move(os.path.join(od, '_tmp_c3', os.path.basename(tmp)), OUT); shutil.rmtree(os.path.join(od, '_tmp_c3')); os.remove(tmp)
# ---- Markdown ----
wr = openpyxl.load_workbook(OUT, data_only=True)['C3_並排']
def v(row, y): return wr.cell(row, YC[y]).value
def f(x, d=1):
    return f'{x:,.{d}f}' if isinstance(x, (int, float)) else ('' if x is None else str(x))
def pct(x): return f'{x*100:.1f}%' if isinstance(x, (int, float)) else ''
md = ['# C3 支出分項並排：Wattonomics Spend 頁 IT 分項 對 Siliconomics 市場值\n',
 f'**日期：** 2026-11-18　**範圍：** 只新增報告，不改任何活頁簿或快照。　**Wattonomics：** {os.path.basename(MODEL)}（Spend 頁，上線年）　**Siliconomics：** {SNAPN}（出貨年）\n',
 '**活頁簿：** `wattonomics/docs/reports/20261118_C3_支出分項對Siliconomics市場值.xlsx`（藍字＝讀入值，黑字＝公式；產生程式 `wattonomics/tools/reports_20261118/c3_spend_vs_sil.py`）。\n',
 '## 一、結論\n',
 f'1. **兩層的 IT 總量差一截：** 2027 年 Wattonomics IT 設備小計 {f(v(it_w,2027))} 對 Siliconomics AI IT 資本支出 {f(v(it_s,2027))}（比 {f(v(it_r,2027),3)}）。拆成 GW 與單價：GW 比 {f(v(rg_r,2027),3)}（上線年 {f(v(rg_w,2027),2)} 對出貨年成本等值 {f(v(rg_s,2027),2)}）、單價比 {f(v(pg_r,2027),3)}。**差距的大部分在 GW，不在單價**，與建議書 r2 的「14 GW 之謎」同一件事。',
 f'2. **分項占比與 Siliconomics 的組成不同：** Spend 的 IT 分項是固定比例（SegIn，GB300 對照）乘 IT 總量，每年占比相同；Siliconomics 的市場值各自獨立建模。例如載板與 PCB 2027 為 {f(v(blocks["it_substrate"]["w"],2027),2)}（占 IT {pct(v(blocks["it_substrate"]["sh"],2027))}），Siliconomics AI 高速 PCB 市場為 {f(v(blocks["it_substrate"]["s"],2027),1)}（占 {pct(v(blocks["it_substrate"]["ssh"],2027))}），即使把 Spend 用 Siliconomics 的 IT 總量重新縮放，比值仍只有 {f(v(blocks["it_substrate"]["rat2"],2027),2)}。主要是比例假設（0.3% 為 [Assumed]）與口徑（見第四節）的問題，不是總量問題。',
 f'3. **HBM 單項已高於 Spend 整個記憶體：** 2027 年 Siliconomics HBM 市場 {f(v(blocks["it_memory"]["s"],2027),1)} 對 Spend 記憶體 {f(v(blocks["it_memory"]["w"],2027),1)}（比 {f(v(blocks["it_memory"]["rat"],2027),2)}）。Spend 的 13.5% 是 GB300 對照的固定比例，可能沒反映記憶體價格上漲後的占比；也可能是兩邊 GW 基準不同（見第二節），本表分不開。',
 '4. **網通、代工封測、儲存沒有可並排的市場值**（Siliconomics 尚未建，或只有設備端口徑），本表只列 Spend 值。',
 '5. 只並排、不連結：Siliconomics 值沒有進任何 Wattonomics 計算；Spend 頁既有的 `SIF_AIITCapex_Base` 對照列（讀 v0.19 鏡像，2029 年 1,605.1，v0.20 為 1,604.2）不動。\n',
 '## 二、總量分解（US$B 或 GW）\n', '| 年 | Wattonomics 上線年 GW | Siliconomics 出貨年 GW | GW 比 | 每 GW IT capex（W） | 每 GW IT capex（S） | 單價比 | W IT 小計 | S AI IT 資本支出 | 總量比 |', '|---|---|---|---|---|---|---|---|---|---|']
for y in YEARS:
    md.append(f'| {y} | {f(v(rg_w,y),2)} | {f(v(rg_s,y),2)} | {f(v(rg_r,y),3)} | {f(v(pg_w,y),2)} | {f(v(pg_s,y),2)} | {f(v(pg_r,y),3)} | {f(v(it_w,y))} | {f(v(it_s,y))} | {f(v(it_r,y),3)} |')
md += ['\n## 三、分項並排（US$B）\n', '| Spend 項目 | 年 | Wattonomics Spend | 占 IT | Siliconomics 市場值 | 差距 | 比值 | Siliconomics 占 IT | 縮放後比值 | 對應性 |', '|---|---|---|---|---|---|---|---|---|---|']
for key, nm, names, rel, note in ITEMS:
    b = blocks[key]
    for y in YEARS:
        if names:
            md.append(f'| {nm} | {y} | {f(v(b["w"],y),2)} | {pct(v(b["sh"],y))} | {f(v(b["s"],y))} | {f(v(b["diff"],y))} | {f(v(b["rat"],y),3)} | {pct(v(b["ssh"],y))} | {f(v(b["rat2"],y),2)} | {rel} |')
        else:
            md.append(f'| {nm} | {y} | {f(v(b["w"],y),2)} | {pct(v(b["sh"],y))} | — | — | — | — | — | {rel} |')
md += ['\n## 四、口徑說明（逐項）\n']
for key, nm, names, rel, note in ITEMS:
    md.append(f'- **{nm}**（{rel}；對應 {"、".join(names) if names else "無"}）：{note}')
md += ['\n## 五、為什麼不放進活頁簿\n',
 'Wattonomics 的 Siliconomics 取數只限基準名稱（`SIL_In`；規則 v1.4.3：上游不讀下游係數，唯一例外是 `baseline_status.json` 只作顯示、不進計算或具名範圍）。Siliconomics 的市場值是下游輸出，把它們放進 Spend 或 Compare 的對照欄，等於讓上游多讀一類下游數字，並且需要升版（v0.3.7）、重建、重跑 Checks。工作單的處置是：若會違反取數規則，就改為 Markdown＋Excel 報告放 `wattonomics/docs/`。因此本張以報告交付；兩個檔案都不是模型的一部分，Wattonomics 的任何計算都不讀它們。\n',
 '## 六、限制\n',
 '- 兩邊年份標籤相同但時點不同（上線年對出貨年），GW 差距有一部分是 `lag_on`、`inv_on` 的時差（規則 C2 ②），不是分歧。',
 '- Spend 的 IT 分項比例（SegIn）是 2026 年 GB300 對照的固定值，未隨世代或 ASIC 比重變動；差距大小不等於哪一側「錯」，只指出組成假設需要由 AI半導體 重新檢視。',
 '- 本表不評斷誰對；兩層修正屬各自單層 PR（比例表的權威在 AI半導體 `segments_ai_semis`，不在 Wattonomics）。\n',
 '## 七、待 Andy 判斷\n',
 '1. SegIn 固定比例要不要改為逐年或讀 Siliconomics 市場值占比：選項 (a) 維持現狀（預設，Spend 只是示意性的市場金額，不進基準）；(b) 請 AI半導體 重做 SegIn（權威在該層），Wattonomics 下一版吸收。',
 '2. 是否在 Compare 新增並排區塊：違反取數規則，預設不做。\n',
 '## 八、通過條件對照\n', '| 條件 | 結果 |\n|---|---|\n| Spend 各 IT 分項與 Siliconomics 同年市場值並排，列差距與口徑說明 | 達成；8 個分項中 5 項有對應市場（運算、記憶體、載板與 PCB、電源散熱、伺服器組裝），3 項無（網通、代工封測、儲存）並寫明原因 |\n| 只並排、不連結 | 達成；Siliconomics 值只在報告，Wattonomics 活頁簿與快照未動 |\n| 會違反取數規則就改為報告放 wattonomics/docs/ | 達成（第五節） |']
open(OUT.replace('.xlsx', '.md'), 'w', encoding='utf-8').write('\n'.join(md) + '\n')
print('ok')
