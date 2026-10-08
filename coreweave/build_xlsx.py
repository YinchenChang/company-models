import json as _jrv
import os as _osrv
rv = _jrv.load(open(_osrv.path.join(_osrv.path.dirname(_osrv.path.abspath(__file__)), 'rv_snap.json'), encoding='utf-8'))
import sys as _sysv; _sysv.path.insert(0, _osrv.path.dirname(_osrv.path.abspath(__file__)))
from vlog import VLOG as _VL
VLOG_X = [list(r) for r in _VL]
import json as _jco
import re
import calendar_q as _calq  # v4.5：期間與日期由 company.json → calendar 推算（與 HTML 共用）
CO = _calq.load(_osrv.path.dirname(_osrv.path.abspath(__file__)))  # 公司資料單一來源（與 HTML 共用）
CAL = CO['cal']
PYEAR = [2000 + int(p[2:]) for p in CO['periods']]  # 5a：各期財年年份由日曆推算
MW_Y0 = PYEAR[0] - 1  # 首期前一財年（期初主動電力）
_n = lambda x: f'{x:g}'  # 5a：company.json 的數字寫入公式或說明（整數不帶小數）
YA = CO['ytdActual']
LQ_DEBT, CONV_PR = CO['latestQuarter']['debtPrincipal'], CO['debt']['convertible']['principal']  # 5a：評價日債務本金、評價日後新發可轉債
PREV_FY = f"FY{int(CO['periods'][0][2:]) - 1:02d}"  # 5a：首期的前一財年（營收 YoY 與比較基期）
PREV_FY_REV = next(h['revenue'] for h in CO['historicalPL'] if h['year'] == PREV_FY)  # v4.5：年初至今實際（原 actual1H）；數值與逐列備註都讀 company.json，滾動時隨資料更新
# v4.3：市場共識資料檔（只讀；路徑在 company.json → meta.consensusFile）。與 HTML 相同的一致性檢查見 build_html_portable.py
CONS = _jco.load(open(_osrv.path.join(_osrv.path.dirname(_osrv.path.abspath(__file__)), CO['meta']['consensusFile']), encoding='utf-8'))
assert (CONS['companyGuidance']['2026Q3']['revenueLow'], CONS['companyGuidance']['2026Q3']['revenueHigh']) == (CO['callFacts']['nextQRevLo'], CO['callFacts']['nextQRevHi']), 'Q3 營收指引：company.json 與共識檔不一致'
assert abs(CO['ytdActual']['adjEbitda'] - CO['ytdActual']['adjEbitdaMeta']['q1'] - CO['ytdActual']['adjEbitdaMeta']['q2']) < 1e-9, '1H 調整後 EBITDA ≠ Q1＋Q2'
D, V, M = CO['defaults'], CO['valuation'], CO['defaults']['m']
PCT_ = lambda xs: [x / 100 for x in xs]  # HTML 以百分點存、Excel 以比例存
  # 版本紀錄單一來源：vlog.py（HTML 端為 tail.js 的 VLOG）
# -*- coding: utf-8 -*-
"""把 CoreWeave 收支模型（HTML v1.3）轉成可重算的 Excel 活頁簿。
所有計算格都是公式，輸入格為藍字，跨頁連結為綠字。"""
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.comments import Comment

BLUE = Font(name="Arial", size=10, color="0000FF")
BLACK = Font(name="Arial", size=10)
GREEN = Font(name="Arial", size=10, color="008000")
BOLD = Font(name="Arial", size=10, bold=True)
HEAD = Font(name="Arial", size=11, bold=True, color="FFFFFF")
TITLE = Font(name="Arial", size=14, bold=True)
SMALL = Font(name="Arial", size=9, color="595959")
FILL_HEAD = PatternFill("solid", fgColor="1F3864")
FILL_SEC = PatternFill("solid", fgColor="D9E2F3")
FILL_KEY = PatternFill("solid", fgColor="FFFF00")
THIN = Side(style="thin", color="BFBFBF")
BOX = Border(top=THIN, bottom=THIN, left=THIN, right=THIN)

NUM = '#,##0.00;(#,##0.00);-'
NUM1 = '#,##0.0;(#,##0.0);-'
NUM0 = '#,##0;(#,##0);-'
PCT = '0.0%'
MULT = '0.0x'
USD = '$#,##0.00;($#,##0.00);-'

PERIODS = CO['periods']  # v4.5：由 calendar 推算
COLS = ["C", "D", "E", "F", "G"]

wb = Workbook()


OUTLINE = {}
FILL_SUB = PatternFill("solid", fgColor="EEF2F7")


def section(ws, row, text, span=8, level=1, collapsed=False):
    c = ws.cell(row=row, column=1, value=("▸ " if level == 2 else "") + text)
    c.font = BOLD if level == 1 else Font(name="Arial", size=10, bold=True, color="1F3864")
    for j in range(1, span + 1):
        ws.cell(row=row, column=j).fill = FILL_SEC if level == 1 else FILL_SUB
    OUTLINE.setdefault(ws.title, []).append((row, level, collapsed))
    return row + 1


def apply_outline(ws):
    heads = sorted(OUTLINE.get(ws.title, []))
    if not heads:
        return
    from openpyxl.worksheet.properties import Outline
    ws.sheet_properties.outlinePr = Outline(summaryBelow=False, summaryRight=True)
    last = ws.max_row
    for idx, (hr, lv, col) in enumerate(heads):
        end = last
        for hr2, lv2, _ in heads[idx + 1:]:
            if lv2 <= lv:
                end = hr2 - 1
                break
        # trim trailing blank rows
        while end > hr and all(ws.cell(row=end, column=k).value in (None, "") for k in range(1, 11)):
            end -= 1
        for rr in range(hr + 1, end + 1):
            d = ws.row_dimensions[rr]
            d.outline_level = max(d.outline_level or 0, lv)
            if col:
                d.hidden = True
        if col:
            ws.row_dimensions[hr].collapsed = True


def label(ws, row, text, note=None, col=1, indent=0):
    c = ws.cell(row=row, column=col, value=("  " * indent) + text)
    c.font = BLACK
    if note:
        n = ws.cell(row=row, column=9, value=note)
        n.font = SMALL
        n.alignment = Alignment(wrap_text=False)
    return row


def period_header(ws, row, first_label="項目", cols=None):
    cols = cols or PERIODS
    ws.cell(row=row, column=1, value=first_label).font = HEAD
    ws.cell(row=row, column=1).fill = FILL_HEAD
    ws.cell(row=row, column=2, value="單位").font = HEAD
    ws.cell(row=row, column=2).fill = FILL_HEAD
    for i, p in enumerate(cols):
        c = ws.cell(row=row, column=3 + i, value=p)
        c.font = HEAD
        c.fill = FILL_HEAD
        c.alignment = Alignment(horizontal="center")
    c = ws.cell(row=row, column=9, value="公式／來源說明")
    c.font = HEAD
    c.fill = FILL_HEAD
    return row + 1


def row_line(ws, row, name, unit, values, fmt=NUM, font=BLACK, note=None, indent=0):
    ws.cell(row=row, column=1, value=("    " * indent) + name).font = BOLD if indent == 0 and font is BOLD else BLACK
    ws.cell(row=row, column=2, value=unit).font = SMALL
    for i, v in enumerate(values):
        c = ws.cell(row=row, column=3 + i, value=v)
        c.font = font
        c.number_format = fmt
        c.border = BOX
    if note:
        n = ws.cell(row=row, column=9, value=note)
        n.font = SMALL
    return row + 1


# =====================================================================
# 1. 導覽
# =====================================================================
ws = wb.active
ws.title = "導覽"
ws.column_dimensions["A"].width = 118
r = 1
ws["A1"] = f"CoreWeave 收支與評價模型 — Excel 版 {VLOG_X[-1][0]}（更新 {CO['meta']['updateDate']}；與 HTML {VLOG_X[-1][0]} 同步；市價截至 {CO['meta']['priceDate']} 收盤）"  # v4.3：版本、日期改讀 vlog.py 與 company.json
ws["A1"].font = TITLE
r = 3
guide = [
    ("這個活頁簿在回答什麼", None),
    ("一句話：CoreWeave 手上 $103.7bn 的合約（RPO）在 2030 年前能變成多少現金，而同一段時間要付出多少現金——", None),
    ("差額就是必須向外籌措的金額。所有分頁用同一組數字，改任何一個輸入，後面全部會跟著動。", None),
    ("", None),
    ("分頁結構（依模組分組，分頁標籤顏色相同者為同一模組）", None),
    ("摘要            一頁摘要：結論、與市場共識的差異、現價隱含什麼、驗證點（v4.3；全部活公式）。", None),
    ("輸入與假設      所有可調參數（A 情境 → B 產能與收入 → C 利潤率 → D 支出 → E 融資 → F 評價 → G «YTD» 實際 → H 情境區間運算表 → I 市場共識 → J 季度追蹤輸入）。", None),
    ("各期收支        類現金流量：來源、用途、融資前缺口、期前融資瀑布。", None),
    ("季度追蹤        焦點季的模型／共識／指引／實際與差距（v4.4；由年度模型拆分，只用於追蹤；實際數填在輸入與假設 J 區）。", None),
    ("運營_產能與收入  合約排程 vs 機房產能 → 收入 → 現金；運營_站點：房東揭露與租金檢驗。", None),
    ("資產負債_既有債務／新債與新股／租賃承諾  類資產負債表：10-Q 債務明細、每期現金與總債務與股數、租賃承諾。", None),
    ("損益            類損益表（«FY0»）。", None),
    ("評價_DCF與目標價／評價_可比公司  DCF、EV/EBITDA、加權目標價；同業 EV/Sales 比較。", None),
    ("評價_反向DCF    現價隱含的收入與成本（目標搜尋＋HTML 快照）。", None),
    ("檢查_連動／檢查_版本紀錄  恆等式與口徑檢查；每一版的變更與目標價；來源：每個硬編碼數字的出處。", None),
    ("註：『各期收支』『資產與負債』『損益』並非會計準則的正式三表，而是對應公司實際運營的類三表呈現。", None),
    ("", None),
    ("顏色規則", None),
    ("藍字 = 手動輸入值（可改）｜黑字 = 本頁公式｜綠字 = 跨頁連結｜黃底 = 關鍵假設｜灰字 = 說明與來源", None),
    ("", None),
    ("■ 重點：容量上限與產能瓶頸是什麼關係", None),
    ("這是整個模型最常被誤解的地方，所以單獨說明。", None),
    ("", None),
    ("兩條線各自獨立產生：", None),
    ("  (A) 排程 RPO ＝ 已簽約、依 10-Q 揭露的時程「應該」在本期認列的收入。", None),
    ("      公式：(«VMD» RPO ＋ Q3 新增承諾) × 本期權重 ÷ 權重合計 × 五期認列比例", None),
    ("      這條線只看合約，完全不管機房蓋好了沒。", None),
    ("  (B) 容量上限 ＝ 本期實際能交付的最大收入 ＝ 平均在役 MW × 每 MW 年收入 × 利用率 × 期間係數。", None),
    ("      這條線只看機房，完全不管簽了多少約。", None),
    ("", None),
    ("兩條線相減，方向決定意義：", None),
    ("  排程 > 容量 → 差額＝『產能瓶頸』。合約在手但機房還沒好，這段收入本期收不到，", None),
    ("      也不遞延到下期（模型不做遞延，是保守處理）。本期收入 ＝ 排程 − 瓶頸 ＝ 容量上限。", None),
    ("  容量 > 排程 → 差額＝『未被期初 RPO 占用的產能』。機房好了但期初合約沒占滿，", None),
    ("      這段要靠新簽約來賣，乘上『新產能簽約率』後成為『新簽約收入』，賣不掉的是『未售產能』。", None),
    ("", None),
    ("所以本模型的收入有兩個來源，而且互斥：期初 RPO 轉換（受容量封頂）＋ 新簽約（用剩餘容量）。", None),
    ("基準路徑下，«P0» «STUBW»與 «P1» 是瓶頸期（合約多於機房），FY28 之後反轉為產能多於期初合約——", None),
    ("這就是為什麼 FY28 起的收入愈來愈依賴『還沒簽的約』，也是 8 GW 敘事的真正風險所在。", None),
    ("", None),
    ("■ 收入之後的三道折扣（產能與收入頁）", None),
    ("  收入 → 扣信用損失（違約率 × (1−回收率)）→ 收現 → 乘『EBITDAR 率』→ 可支應資本用途的現金。", None),
    ("  貢獻率是租金與利息『之前』的比率，因為租金與利息在支出頁單獨列示，不可在此重複扣除。", None),
    ("  Q2 反推上限：Adj. EBITDA 率 59%（已扣營業租賃 0.500＋變動租賃 0.150，約營收 25%）→ 租金前約 84%。", None),
    ("", None),
    ("■ 期別口徑：FY26 ＝ 1H 實際 ＋ 2H 模型", None),
    ("  支出與資金頁與損益頁的 FY26 欄，1H 部分直接採用 10-Q 現金流量表的實際值（CFO 3.663、現金購置 14.117、", None),
    ("  JV 0.688、借款 16.747、還款 5.219、私募股權 2.982、capped call 0.492），2H 部分才是模型推算。", None),
    ("  因此 FY26 的 CapEx 37.1 與營收 12.8 可以直接對照公司全年指引 35–39 與 12.4–13.2。", None),
    ("  產能與收入頁的首欄也是 «P0»，但標示「（模型期）」的收入列只含«STUBW»：RPO 桶自 «VD» 起算，", None),
    ("  1H 的收入已經實現、不再受產能約束，因此 1H 實際單列，兩者相加即該頁的「FY26 全年總營收」。DCF 只折現 6/30 之後的現金流。", None),
    ("  1H 的利息與租金已包含在實際 CFO 內，所以支出頁那兩列的 FY26 欄只有 2H 金額；全年數字見「FY26 全年備忘」區。", None),
    ("", None),
    ("■ 現金橋（支出與資金頁最下方）", None),
    ("  2025-12-31 現金 3.127 ＋ 營運缺口 ＋ 股權／可轉債 ＋ 1H 實際借款 − 排程還本 − capped call ＋ 調節 ＝ 期末現金", None),
    ("  營運缺口＝營運來源 −（用途 − 排程還本）。股權、可轉債、信用額度都不算營運來源，這是刻意的：", None),
    ("  它們回答的是「誰來補」，不是「本業能不能自給」。", None),
    ("", None),
    ("■ v1.5 變更（回應 Andy 反饋 1–5）", None),
    ("  1. 新增『站點租賃』頁：Helios／Polaris Forge 1／Core Scientific／Kenilworth 逐站列出契約 MW、合約值、每 MW 年租金，", None),
    ("     並與 10-Q 租賃承諾總額 79.6 對帳（具名 36.2、未具名殘差 43.4）；另以市場基準檢驗模型租金路徑。", None),
    ("  2. 新增『債務明細』頁：10-Q 16 筆工具逐筆列示（合計 35.551、加權有效利率 8.4%），存量利息改由到期表遞減本金推算。", None),
    ("  3. JV 拆為「已承諾餘額 1.15（10-Q：2026 年內履行）」＋「後續增資（假設）」兩列。", None),
    ("  4. 毛 CapEx 改由 Accepted MW 推導：(本期新增×(1−λ)＋次期新增×λ)×每 MW 成本；以 FY26 全年指引 35–39 校準。", None),
    ("  5. 刪除『受產能約束占 RPO』（Oracle 模板遺留，CoreWeave 恆為 100%）。", None),
    ("", None),
    ("■ v2.4：股權吸收能力可調＋融資前缺口", None),
    ("  股權每年最多募到現市值的 20%（可調；2026 年實際約 13%），超出部分以高息債 12% 補足（可調）。發行折價 10%（可調）。", None),
    ("  「支出與資金」頁新增『融資前累積現金』：若不做任何新融資，現金會落到多少——負值即外部資金需求。", None),
    ("  「EBITDAR 率」＝EBITDA 率＋租金÷營收：資金端使用的租金前利潤率，由 EBITDA 率推導、不是獨立輸入。", None),
    ("", None),
    ("■ v2.3：EBITDA 率單一來源＋融資校準", None),
    ("  EBITDA 率只在「輸入」頁兩格設定：起始 59%（Q2 實際）、穩態 65%（FY30），中間線性爬升；損益 EBITDA 與資金現金都引用同一列。", None),
    ("  資金端EBITDAR 率＝EBITDA 率＋租金÷營收（租金在支出端另列，故加回）；非算力服務現金＝服務營收×同一貢獻率。", None),
    ("  債務上限預設 1.0x（依 2026 年實際融資組合：債務約 77%、股權約 23%）。敏感度：把穩態改成 59%／70%、上限改成 0.4x 即可重算。", None),
    ("", None),
    ("■ v2.0：期前融資瀑布（取代「透支利息＋扣缺口」）", None),
    ("  缺口在需要前一期先融好，使每期末現金不低於最低現金 2.0bn。順序：未動用額度 10.0 → 資產層新債", None),
    ("  （總債務 ≤ 1.0 × 期末 backlog，依 2026 年實際融資組合校準；0.4x 為敏感度）→ 股權（9/21 收盤 $85.43 折價 10% 發行）。", None),
    ("  DCF 以融資後股數計每股，並加回新股募得現金的現值；EV/EBITDA 用 FY27 末淨負債與 FY27 末股數。", None),
    ("  舊版的透支利息（借款利息）與扣 NTM 缺口（借款本金）並非重複計費；v2.0 的差別是新債設上限、超出部分改以股權支應。瀑布位於「支出與資金」頁下方。", None),
]
for text, _ in guide:
    c = ws.cell(row=r, column=1, value=text)
    if text.startswith("■") or text in ("分頁結構（依模組分組，分頁標籤顏色相同者為同一模組）", "顏色規則", "這個活頁簿在回答什麼"):
        c.font = BOLD
    else:
        c.font = BLACK
    r += 1

# =====================================================================
# 2. 輸入（依類別排序；每一類可展開／收合）
# =====================================================================
ws = wb.create_sheet("輸入與假設")
ws.column_dimensions["A"].width = 44
ws.column_dimensions["B"].width = 12
for c in COLS:
    ws.column_dimensions[c].width = 12
ws.column_dimensions["I"].width = 92
ws.column_dimensions["J"].width = 60
ws["A1"] = "輸入（依類別排序；點左側 ＋／− 展開或收合。藍字可改；黃底為最關鍵假設）"
ws["A1"].font = TITLE
ws["A2"] = "A 情境 → B 產能與收入 → C 利潤率 → D 支出（CapEx／折舊／租賃／JV）→ E 融資 → F 評價 → G 參考：«YTDL» 實際 → H 情境區間運算表 → I 市場共識 → J 季度追蹤輸入。灰底『▸』為明細子區，預設收合"
ws["A2"].font = SMALL
IN = {}


def gi(row, name, unit, value, note, fmt=NUM, key=False, font=None):
    ws.cell(row=row, column=1, value=name).font = BLACK
    ws.cell(row=row, column=2, value=unit).font = SMALL
    c = ws.cell(row=row, column=3, value=value)
    c.font = font or (BLACK if isinstance(value, str) and value.startswith("=") else BLUE)
    c.number_format = fmt
    c.border = BOX
    if key:
        c.fill = FILL_KEY
    n = ws.cell(row=row, column=9, value=note)
    n.font = SMALL
    return f"'輸入與假設'!$C${row}"


def phdr(row, first="項目", last="來源／標籤"):
    for j, h in enumerate([first, "單位"] + PERIODS):
        c = ws.cell(row=row, column=1 + j, value=h); c.font = HEAD; c.fill = FILL_HEAD
        if j >= 2:
            c.alignment = Alignment(horizontal="center")
    c = ws.cell(row=row, column=9, value=last); c.font = HEAD; c.fill = FILL_HEAD
    return row + 1


def prow(row, name, unit, vals, fmt, note, font=BLUE, key=None):
    row_line(ws, row, name, unit, vals, fmt, font, note)
    IN[key or name] = row
    return row + 1


r = 4
# ---------------- A 情境 ----------------
r = section(ws, r, "A｜情境與規模（管理層擴張力道）")
ws.cell(row=r, column=1, value="情境選擇（1＝保守 4.2 GW、2＝基準 5.6 GW、3＝積極 8 GW）").font = BOLD
c = ws.cell(row=r, column=3, value=2); c.font = BLUE; c.number_format = NUM0; c.border = BOX; c.fill = FILL_KEY
ws.cell(row=r, column=4, value=f'=CHOOSE(C{r},"保守 4.2 GW","基準 5.6 GW","積極 8 GW")').font = BOLD
ws.cell(row=r, column=9, value="三情境只改變管理層的擴張力道；其餘假設相同。FY26–FY27 三者相同（年底 1,850／3,100 MW），FY28 起分歧").font = SMALL
SEL = f"'輸入與假設'!$C${r}"; r += 1
r = section(ws, r, "情境路徑明細（Accepted MW、爬坡比例、表外租金基準）", level=2)
for j, h in enumerate(["Accepted MW 路徑", "單位"] + PERIODS + ["", "FY31 新增 MW"]):
    if h:
        cc = ws.cell(row=r, column=1 + j, value=h); cc.font = HEAD; cc.fill = FILL_HEAD
r += 1
sc_rows = {}
for nm, vals, m31, note in [
    ("保守 4.2 GW", CO['scenarios']['accepted']['low'], CO['scenarios']['mw31']['low'], "只交付已簽約電力 4.2 GW（8/11 法說），不再新簽"),
    ("基準 5.6 GW", CO['scenarios']['accepted']['base'], CO['scenarios']['mw31']['base'], "2027 年後每年約 +800 MW"),
    ("積極 8 GW", CO['scenarios']['accepted']['high'], CO['scenarios']['mw31']['high'], "公司目標 2030 ≥8 GW"),
]:
    ws.cell(row=r, column=1, value=nm).font = BLACK
    ws.cell(row=r, column=2, value="MW").font = SMALL
    for i, v in enumerate(vals):
        c = ws.cell(row=r, column=3 + i, value=v); c.font = BLUE; c.number_format = NUM0; c.border = BOX
    c = ws.cell(row=r, column=9, value=m31); c.font = BLUE; c.number_format = NUM0; c.border = BOX
    ws.cell(row=r, column=10, value=note).font = SMALL
    sc_rows[nm] = r
    r += 1
row_line(ws, r, "Billable／Accepted 爬坡比例", "%", [b / a for b, a in zip(CO['scenarios']['billableRatio']['billable'], CO['scenarios']['billableRatio']['accepted'])], PCT, BLUE,
         "三情境共用")
BR_ROW = r; r += 1
row_line(ws, r, "表外現金租金（積極路徑）", "US$bn", CO['scenarios']['leaseHighPath'], NUM, BLUE,
         f"其他情境依「新增 MW（相對 «VMD» 的 {CO['scenarios']['leaseRampFloorMw']:,} MW）」等比例縮放")
LEASE_HI = r; r += 1
LEASE_FL = gi(r, "表外租金起算 MW", "MW", CO['scenarios']['leaseRampFloorMw'], "«VMD» 主動電力；表外租金依超過此值的新增 MW 等比例縮放（company.json → scenarios.leaseRampFloorMw）", NUM0); r += 1

# ---------------- B 產能與收入 ----------------
r = section(ws, r, "B｜產能與收入（RPO、MW、利用率、單價、簽約、信用）")
RPO0 = gi(r, "«VMD» RPO", "US$bn", D['rpoOpen'], "10-Q 103.7（backlog 104.2 含 0.5 其他）[Verified]", NUM, True); r += 1
RPOADD = gi(r, "Q3 初新增承諾", "US$bn", D['rpoPendingAdd'], "法說「>25」，尚未入 RPO [Interested-party]"); r += 1
RP = gi(r, "五期認列比例（至 2030 末）", "%", D['rp'] / 100, "10-Q 桶 41／39／20 線性分攤後五期合計 84% [Derived]", PCT, True); r += 1
WSUM = gi(r, "RPO 桶權重合計", "%", CO['rpo']['scheduledShare'], "下方權重列的合計，用於歸一化", PCT); r += 1
MWY = {}  # 5a：各年底主動電力（company.json → defaults.mwYearEnd，以年份為鍵；說明在 texts.mwYearEndNotes）
MWY[MW_Y0] = MW_YE25 = gi(r, f"{CAL['prevFYE']} 主動電力", "MW", D['mwYearEnd'][str(MW_Y0)], CO['texts']['mwYearEndNotes'][str(MW_Y0)], NUM0); r += 1
MW0 = gi(r, "«VMD» Billable MW", "MW", D['billableOpen'], "主動電力 1.5 GW 扣爬坡中 [Assumed]", NUM0); r += 1
AVGON = gi(r, "收入用平均在役 MW（1=是）", "", int(D['useAvgMw']), "0＝用期末存量全期化。CRWV 交付後載（6 月 +300MW），建議 1", NUM0); r += 1
REVSC = gi(r, "每 MW 年收入倍數（整體）", "%", D['revScale'], "反向 DCF 與壓力測試用；預設 100%。以『目標搜尋』調整此格即可反解市價隱含單價", PCT); r += 1
r = phdr(r)
r = prow(r, "模型期長度（年）", "", CO['periodYears'], NUM,
         "勿改。FY26 模型期為 2026-07-01～12-31（0.5 年）。只用於容量上限、電力、維護三處年率換算 [Derived]")
r = prow(r, "RPO 桶權重", "%", CO['rpo']['bucketWeights'], PCT, "10-Q 41%／24m、39%／25-48m、20%／49-78m 桶內線性分攤 [Derived]")
ws.cell(row=IN["RPO 桶權重"], column=8, value="=SUM(C{0}:G{0})".format(IN["RPO 桶權重"])).number_format = PCT
r = prow(r, "Accepted MW（期末主動電力）", "MW", [0] * 5, NUM0, "＝依 A 區情境選擇器", BLACK)
r = prow(r, "Billable MW", "MW", [0] * 5, NUM0, "＝Accepted × 爬坡比例（A 區）", BLACK)
_acc = IN["Accepted MW（期末主動電力）"]; _bil = IN["Billable MW"]
for i in range(5):
    L = COLS[i]
    ws.cell(row=_acc, column=3 + i, value=f"=CHOOSE({SEL},{L}{sc_rows['保守 4.2 GW']},{L}{sc_rows['基準 5.6 GW']},{L}{sc_rows['積極 8 GW']})")
    ws.cell(row=_bil, column=3 + i, value=f"=ROUND({L}{_acc}*{L}{BR_ROW},0)")
r = prow(r, "利用率", "%", PCT_(M['util']), PCT, "法說「近期產能實質售罄」，不擬合 100% [Assumed]")
r = prow(r, "每 MW 年收入", "US$bn/MW", M['revMW'], '0.0000', "期末 ARR 18.5–19.5÷1.85GW≈10.0–10.5M/MW；7 月漲價 25% 只影響新約 [Derived]")
r = prow(r, "新產能簽約率", "%", PCT_(M['fill']), PCT, "未被期初 RPO 占用的產能能賣掉的比例 [Assumed]")
r = prow(r, "客戶違約率", "%", PCT_(M['defaultP']), PCT, "前三大客戶 72% 營收的定價，非預測 [Assumed]")
r = prow(r, "回收率", "%", PCT_(M['recovery']), PCT, "[Assumed]")
r = prow(r, "非算力服務營收", "US$bn", D['services'], NUM, "軟體／儲存／CPU／managed inference；損益營收與資金現金共用 [Derived]")
r += 1
# ---------------- W2：世代組合（B 區之後的子區；公司專屬輸入＋Tokenomics 引用值，依世代） ----------------
# 方法開關 company.json → methodology.perMw（建置時決定公式）：全為 legacy／ebitdaPct 時，既有各列公式與 v4.5 完全相同（本子區與「每MW經濟性」頁只作對照）。
PMW = {'capex': 'legacy', 'cost': 'ebitdaPct', 'revenue': 'legacy', **(CO.get('methodology', {}).get('perMw') or {})}
FL = CO.get('fleet')
GEN = FL['generations'] if FL else []
MWBASIS = CO['meta'].get('mwBasis', 'IT')
_TKX = (_jco.load(open(_osrv.path.join(_osrv.path.dirname(_osrv.path.abspath(__file__)), CO['tokenomics']['snapshotFile']), encoding='utf-8'))
        if CO.get('tokenomics') else {'items': {}, 'source': {'generations': []}})
_TKGC = {g['name']: g['code'] for g in _TKX['source']['generations']}
tk_has = lambda n: n in _TKX['items'] and not _TKX['items'][n].get('missing')
TK_PH = ['IF_MaintIT', 'IF_StaffSW', 'IF_TaxIns', 'IF_DeprLifeIT']  # 可暫代的名稱（待 Tokenomics v5.26）
TK_MISS = [n for n in (['IF_DeprLifeIT'] if PMW['capex'] == 'tokenomics' else []) + (['IF_MaintIT', 'IF_StaffSW', 'IF_TaxIns'] if PMW['cost'] == 'bottomUp' else []) if not tk_has(n)]
if PMW['capex'] != 'legacy' or PMW['cost'] != 'ebitdaPct' or PMW['revenue'] != 'legacy':
    assert FL, 'methodology.perMw 使用新方法時 company.json 必須有 fleet（世代組合）'
    for _n_ in ['IF_CapexIT', 'IF_CapexTotal', 'IF_PowerCost', 'IF_GPUsPerGW', 'IF_FacilityGW', 'IF_HoldEcon', 'IF_GPUhrEcon']:
        assert tk_has(_n_), f'Tokenomics 快照缺少必要名稱 {_n_}（不可暫代）'
    assert set(GEN) <= set(_TKGC), f'fleet.generations 須與 Tokenomics 世代同名：{set(GEN) - set(_TKGC)}'
AMQ = (CO.get('pricing') or {}).get('anchorMultiple')  # W4：Tokenomics 錨的公司因素（定價倍數 k、隨需占比、證據表）
if PMW['revenue'] == 'tkAnchor':
    assert AMQ and all(k in AMQ for k in ('long', 'spot', 'onDemandShare', 'evidence')), 'revenue=tkAnchor：company.json 需要 pricing.anchorMultiple（long、spot、onDemandShare、evidence）'
    assert not [e for e in AMQ['evidence'] if e['use'] in ('long', 'spot', 'range') and not tk_has(e['tkName'])], 'revenue=tkAnchor：證據表引用的 Tokenomics 名稱不在快照'
if PMW['revenue'] == 'gpuHr':
    _P = (CO.get('pricing') or {}).get('gpuHr') or {}
    assert all(_P.get(g, {}).get('base') is not None for g in GEN if any(m.get(g) for m in FL['newMix']) or FL['openMix']['mix'].get(g)), \
        'revenue=gpuHr：pricing.gpuHr 缺少在役世代的 GPU 小時價格（資料不足時改用 revenue=legacy 備案）'
FLX = {}  # 世代組合各列的列號
if FL:
    r = section(ws, r, "世代組合（W2；公司專屬，世代與 Tokenomics 同名；每 MW 資本支出、成本、收入依此加權）", level=2, collapsed=True)
    ws.cell(row=r, column=1, value="每 MW 方法（資本支出／營運成本／收入）").font = BOLD
    ws.cell(row=r, column=3, value=f"{PMW['capex']}／{PMW['cost']}／{PMW['revenue']}").font = BLACK
    ws.cell(row=r, column=9, value="company.json → methodology.perMw（建置時決定公式）：tokenomics｜legacy、bottomUp｜ebitdaPct、gpuHr｜legacy；全為舊方法時與 v4.5 相同").font = SMALL
    r += 1
    MWB = gi(r, "MW 口徑（1＝IT 關鍵電力、2＝設施）", "", 1 if MWBASIS == 'IT' else 2,
             "company.json → meta.mwBasis。Tokenomics 每 GW＝IT 關鍵電力；設施口徑時 IT MW＝設施 MW ÷ IF_FacilityGW（各 Tokenomics 每 MW 值 × 換算係數）[Assumed，見 W1 第 6a 步]", NUM0); r += 1
    TKSEL = gi(r, "Tokenomics 成本情境（1＝低成本、2＝基準、3＝高成本）", "", 2, "模型主值取基準；低／高只作敏感度（每MW經濟性頁「敏感度」）", NUM0); r += 1
    MIXALT = gi(r, "世代組合敏感度（0＝基準、1＝Rubin Ultra 版）", "", 0, FL['newMixAlt']['label'] + "；只作敏感度", NUM0); r += 1
    PXSEL = gi(r, "GPU 小時價格情境（1＝低、2＝基準、3＝高）", "", 2, "只在 revenue=gpuHr 時影響收入；價格讀 company.json → pricing.gpuHr", NUM0); r += 1
    OPENMW = gi(r, f"期初在役 MW（{FL['openMix']['asOf']}，世代組合起點）", "MW", FL['openMix']['activeMW'], "company.json → fleet.openMix.activeMW（最新已申報季末主動電力）[Verified]", NUM0); r += 1
    FLX['open'] = {}
    for g in GEN:
        FLX['open'][g] = gi(r, f"期初在役占比｜{g}", "%", FL['openMix']['mix'].get(g, 0), FL['openMix']['tag'] + "；" + (FL['openMix']['source'] if g == GEN[0] else "同上"), PCT); r += 1
    r = phdr(r, "新增 MW 的世代占比")
    for kind, mixes, note in (("基準", FL['newMix'], FL['newMixNote']), ("Rubin Ultra 版", FL['newMixAlt']['mix'], FL['newMixAlt']['tag'])):
        FLX[kind] = r
        for g in GEN:
            r = prow(r, f"新增占比（{kind}）｜{g}", "%", [m.get(g, 0) for m in mixes], PCT, note if g == GEN[0] else None)
    FLX['adopt'] = r
    for j, g in enumerate(GEN):
        r = prow(r, f"新增占比（採用）｜{g}", "%", [f"=IF({MIXALT}=1,{COLS[i]}{FLX['Rubin Ultra 版'] + j},{COLS[i]}{FLX['基準'] + j})" for i in range(5)], PCT,
                 "＝依「世代組合敏感度」選擇基準或 Rubin Ultra 版" if j == 0 else None, BLACK)
    r = prow(r, "新增占比合計（須為 100%）", "%", [f"=SUM({COLS[i]}{FLX['adopt']}:{COLS[i]}{FLX['adopt'] + len(GEN) - 1})" for i in range(5)], PCT, "檢查頁核對", BLACK)
    # Tokenomics 引用值（列＝世代、欄＝項目；經 W1「Tokenomics_取數」具名範圍；OFFSET 依成本情境取低／基準／高欄；換成公司 MW 口徑）
    def _tkc(n, g, conv=True, sel=True):
        if not tk_has(n): return ""
        k = f"TK_{re.sub(r'^(IF|L1)_', '', n)}_{_TKGC[g]}"
        x = f"OFFSET({k},0,{TKSEL}-2)" if sel else k
        return f"={x}*{COLS[3]}{{conv}}" if conv else f"={x}"
    def _tkblock(title, cols, unit_note):
        global r
        for j, h in enumerate([title, "世代"] + [c[0] for c in cols]):
            c = ws.cell(row=r, column=1 + j, value=h); c.font = HEAD; c.fill = FILL_HEAD
        ws.cell(row=r, column=9, value=unit_note).font = SMALL
        r += 1; top = r
        for g in GEN:
            ws.cell(row=r, column=1, value=f"{title.split('（')[0]}｜{g}").font = BLACK
            ws.cell(row=r, column=2, value=_TKGC[g]).font = SMALL
            for i, (_, fn) in enumerate(cols):
                v = fn(g, r)
                c = ws.cell(row=r, column=3 + i, value=v if v != "" else None); c.font = GREEN; c.border = BOX
                c.number_format = '#,##0.000;(#,##0.000);-'
            r += 1
        return top
    FLX['tkCap'] = _tkblock("TK 資本與口徑（IT 設備、合計、設施／IT、MW 換算、折舊年限）", [
        ("IT 設備", lambda g, rr: _tkc('IF_CapexIT', g).replace('{conv}', str(rr))),
        ("合計", lambda g, rr: _tkc('IF_CapexTotal', g).replace('{conv}', str(rr))),
        ("設施／IT", lambda g, rr: _tkc('IF_FacilityGW', g, conv=False, sel=False)),
        ("MW 換算", lambda g, rr: f"=IF({MWB}=2,1/{COLS[2]}{rr},1)"),
        ("折舊年限", lambda g, rr: _tkc('IF_DeprLifeIT', g, conv=False, sel=False))],
        "US$m／公司 MW（IF_CapexIT、IF_CapexTotal 每 GW＝每 MW 百萬美元，× MW 換算）；設施／IT＝IF_FacilityGW（基準）；MW 換算＝IT 口徑 1、設施口徑 1÷設施／IT；折舊年限＝IF_DeprLifeIT（v5.26 前空白，暫代 D 區 GPU 經濟壽命）")
    FLX['tkOp'] = _tkblock("TK 營運成本（電費、IT 維護、人員軟體、稅險）", [
        ("電費", lambda g, rr: _tkc('IF_PowerCost', g).replace('{conv}', str(FLX['tkCap'] + GEN.index(g)))),
        ("IT 維護", lambda g, rr: _tkc('IF_MaintIT', g).replace('{conv}', str(FLX['tkCap'] + GEN.index(g)))),
        ("人員軟體", lambda g, rr: _tkc('IF_StaffSW', g).replace('{conv}', str(FLX['tkCap'] + GEN.index(g)))),
        ("稅險×IT占比", lambda g, rr: (_tkc('IF_TaxIns', g).replace('{conv}', str(FLX['tkCap'] + GEN.index(g))) + f"*C{FLX['tkCap'] + GEN.index(g)}/D{FLX['tkCap'] + GEN.index(g)}") if tk_has('IF_TaxIns') else "")],
        "US$m／公司 MW／年。電費＝IF_PowerCost（已含平均用電比與 PUE）；IT 維護 IF_MaintIT、人員軟體 IF_StaffSW、稅險 IF_TaxIns × IT 資本占總資本比（v5.26 前空白：IT 維護暫代 C 區『維護成本』、人員軟體與稅險暫代 0）")
    FLX['tkRev'] = _tkblock("TK 收入參考（GPU 數、GPU 小時持有成本、經濟持有成本、GPU 小時合約價）", [
        ("GPU／MW", lambda g, rr: _tkc('IF_GPUsPerGW', g, sel=False).replace('{conv}', str(FLX['tkCap'] + GEN.index(g))) + "/1000"),
        ("GPU 時成本", lambda g, rr: _tkc('IF_GPUhrEcon', g, conv=False)),
        ("持有成本", lambda g, rr: _tkc('IF_HoldEcon', g).replace('{conv}', str(FLX['tkCap'] + GEN.index(g)))),
        ("合約價", lambda g, rr: (lambda P: f"=CHOOSE({PXSEL},{P['low'] if P.get('low') is not None else P['base']},{P['base']},{P['high'] if P.get('high') is not None else P['base']})" if P and P.get('base') is not None else "")(((CO.get('pricing') or {}).get('gpuHr') or {}).get(g)))],
        "GPU／MW＝IF_GPUsPerGW ÷ 1000 × MW 換算（顆）；GPU 時成本＝IF_GPUhrEcon（US$/GPU-hr，100% 時數、不含利潤）；持有成本＝IF_HoldEcon（US$m／MW／年，含廠房資本回收）；合約價＝company.json → pricing.gpuHr（空白＝不適用）")
    AMX = {}  # W4：定價倍數 k（公司因素；只在 company.json 有 pricing.anchorMultiple 時建列；收入方法 tkAnchor 時進入模型）
    if AMQ:
        r = phdr(r, "定價倍數 k（W4；Tokenomics 錨的公司因素）")
        _kl, _ks, _od = AMQ['long'], AMQ['spot'], AMQ['onDemandShare']
        AMX['kL'] = gi(r, "定價倍數 k_長約（市場長約價 ÷ Tokenomics 同世代持有成本）", "倍", _kl['base'],
                       f"company.json → pricing.anchorMultiple.long（基準成本情境的值；其他成本情境以基準證據世代的成本比例重算）；區間 {_n(_kl['low'])}–{_n(_kl['high'])}、三筆長約中位數 {_n(_kl['sensMedian'])} {_kl['tag']}；證據見『每MW經濟性』k 證據表（不以公司營收、ARR、RPO 金額反推）", '0.00'); r += 1
        AMX['kS'] = gi(r, "定價倍數 k_現貨（市場現貨價 ÷ Tokenomics 同世代持有成本）", "倍", _ks['base'],
                       f"company.json → pricing.anchorMultiple.spot（基準成本情境的值）；區間 {_n(_ks['low'])}–{_n(_ks['high'])} {_ks['tag']}", '0.00'); r += 1
        AMX['od'] = gi(r, "隨需占比（占在役計費產能）", "%", _od['base'],
                       "k＝隨需占比 × k_現貨＋（1 − 隨需占比）× k_長約；CRWV 先簽多年期合約再建產能，RPO 未覆蓋的新增產能也按長約價（W4 r2）。公司未揭露隨需比例，敏感度 " + "／".join(f"{x * 100:g}%" for x in _od['sens']) + f" {_od['tag']}；不得以 Q2 隱含 k 反推", PCT); r += 1
        FLX['tkCapRev'] = _tkblock("TK 收入上限（客戶每 GW 付費 token 營收）", [
            ("token 營收", lambda g, rr: _tkc('IF_RevGWFleet', g).replace('{conv}', str(FLX['tkCap'] + GEN.index(g))))],
            "US$m／公司 MW／年＝IF_RevGWFleet（OpenAI 有效單價、層級組合的理想上限；× MW 換算）。收入上限檢查：CRWV 每 MW 計費收入 ÷ 本欄（快照沒有此名稱時空白＝不適用）")
    r += 1

# ---------------- C 利潤率 ----------------
r = section(ws, r, "C｜利潤率（EBITDA 率單一來源；損益與資金共用）")
EB0 = gi(r, "起始 EBITDA 率（FY26）", "%", D['ebStart'], "Q2 實際 Adj. EBITDA 率 58.6% [Verified]", PCT, True); r += 1
EBSS = gi(r, "穩態 EBITDA 率（FY30）", "%", D['ebSteady'], "由下而上：每 MW 現金成本約 $3.2–3.6m／年收入 $10.9m → 67–71%；Q2 59% 含爬坡成本。基準 65%，敏感度 59%／70% [Derived]", PCT, True); r += 1
r = phdr(r)
r = prow(r, "EBITDA 率", "%", [0] * 5, PCT, "＝起始＋(穩態−起始)×期數÷4（公式）", BLACK)
for i in range(5):
    ws.cell(row=IN["EBITDA 率"], column=3 + i, value=f"={EB0}+({EBSS}-{EB0})*{i}/4")
r = prow(r, "非算力服務現金", "US$bn", [0] * 5, NUM, "＝服務營收 × EBITDAR 率（公式）", BLACK)
r = section(ws, r, "電力／維護 overlay（預設關閉；EBITDA 已含電費）", level=2, collapsed=True)
OVERLAY = gi(r, "電力／維護 overlay（1=開）", "", int(D['overlay']), "EBITDA 率已含電費，開啟會重複扣除，僅供壓力測試", NUM0); r += 1
r = prow(r, "電價", "$/MWh", M['power'], NUM0, "overlay 用 [Assumed]")
r = prow(r, "PUE", "", M['pue'], '0.00', "overlay 用 [Assumed]")
r = prow(r, "維護成本", "US$m/MW", M['maint'], '0.00', "overlay 用 [Assumed]" + ("；W2 由下而上：IT 維護暫代值（待 Tokenomics v5.26 IF_MaintIT）" if PMW['cost'] == 'bottomUp' and not tk_has('IF_MaintIT') else ""))
CSX = {}  # W2：由下而上營運成本的公司輸入（最近一季 10-Q；company.json → latestQuarter、costs）
if FL:
    r = section(ws, r, "由下而上營運成本輸入（W2；最近一季 10-Q，公司專屬）", level=2, collapsed=True)
    _LQ = CO['latestQuarter']; _CS = CO.get('costs') or {}
    for _k, _lab, _note in (('revenue', '最近一季營收', '10-Q 損益表 [Verified]'), ('costRev', '最近一季營收成本', '10-Q 損益表 [Verified]'),
                            ('techInfra', '最近一季技術與基礎設施', '10-Q 損益表；D&A 主要在此列 [Verified]'),
                            ('sm', '最近一季銷售行銷', '10-Q 損益表 [Verified]'), ('smSbc', '　其中 SBC（銷售行銷）', '10-Q SBC 附註 [Verified]'),
                            ('ga', '最近一季一般管理', '10-Q 損益表 [Verified]'), ('gaSbc', '　其中 SBC（一般管理）', '10-Q SBC 附註 [Verified]'),
                            ('da', '最近一季 D&A 合計', '10-Q 現金流量表（未依列別揭露，視為全在營收成本與技術基礎設施）[Verified]'),
                            ('sbcCostTi', '　SBC（營收成本＋技術與基礎設施）', '10-Q SBC 附註：技術與基礎設施 0.060＋營收成本 0.009 [Verified]'),
                            ('opLeaseCost', '最近一季營業租賃成本', '10-Q 租賃附註 [Verified]'), ('varLeaseCost', '最近一季變動租賃成本', '10-Q 租賃附註 [Verified]')):
        CSX[_k] = gi(r, _lab, "US$bn", _LQ[_k], _note); r += 1
    CSX['basis'] = gi(r, "管銷率口徑（1＝扣 SBC、2＝GAAP 含 SBC）", "", 2 if _CS.get('sgaBasis') == 'gaap' else 1, "company.json → costs.sgaBasis；預設扣 SBC（W1 第 6d 步），GAAP 含 SBC 只作敏感度", NUM0); r += 1
    CSX['sga'] = gi(r, "管銷率（扣 D&A；各期持平）", "%", f"=IF({CSX['basis']}=2,({CSX['sm']}+{CSX['ga']})/{CSX['revenue']},({CSX['sm']}-{CSX['smSbc']}+{CSX['ga']}-{CSX['gaSbc']})/{CSX['revenue']})",
                    "＝（銷售行銷 − SBC ＋ 一般管理 − SBC）÷ 營收；銷售行銷與一般管理內的 D&A 視為 0 [Derived]", PCT); r += 1
    CSX['addQ'] = gi(r, "最近一季新增主動電力", "MW", CO['callFacts']['activeAddQ2'], "法說（Q2 新增近 500 MW）；Q2 平均在役 MW＝期末 − 新增 ÷ 2 [Interested-party]", NUM0); r += 1
r += 1

# ---------------- D 支出 ----------------
r = section(ws, r, "D｜支出（CapEx、車隊折舊、租賃、JV）")
LAMBDA = gi(r, "提前支出比例 λ", "%", D['lambda'], "次年才上線的 MW，其建置支出落在前一年的比例 [Assumed]", PCT); r += 1
MW31 = gi(r, "FY31 新增 MW（FY30 預建用）", "MW",
          f"=CHOOSE({SEL},I{sc_rows['保守 4.2 GW']},I{sc_rows['基準 5.6 GW']},I{sc_rows['積極 8 GW']})",
          "隨情境：保守 0／基準 500／積極 1,000 [Assumed]", NUM0); r += 1
FLOOR = gi(r, "FY26 CapEx 下限（已承諾）", "US$bn", D['capexFloorFY0'], "全年指引下緣：2026 年支出多已下單 [Interested-party]"); r += 1
if PMW['capex'] == 'tokenomics' and tk_has('IF_DeprLifeIT'):  # W2：IF_DeprLifeIT 依期初在役世代加權、取整數
    LIFE = gi(r, "GPU 經濟壽命（年）", "年", f"=ROUND(SUMPRODUCT($C${FLX['open'][GEN[0]].split('$')[-1]}:$C${FLX['open'][GEN[-1]].split('$')[-1]},$G${FLX['tkCap']}:$G${FLX['tkCap'] + len(GEN) - 1}),0)",
              "＝Tokenomics IF_DeprLifeIT 依期初在役世代加權、取整數；決定汰換時點與車隊折舊 [Derived]", NUM0); r += 1
else:
    LIFE = gi(r, "GPU 經濟壽命（年）", "年", D['gpuLife'], "公司對伺服器與網路設備的折舊年限；決定汰換時點與車隊折舊 [Verified]" + ("；W2：待 Tokenomics v5.26 IF_DeprLifeIT（暫代值）" if PMW['capex'] == 'tokenomics' else ""), NUM0); r += 1
for _y in sorted(int(y) for y in D['mwYearEnd'] if int(y) != MW_Y0):
    MWY[_y] = gi(r, f"{_y} 年底主動電力", "MW", D['mwYearEnd'][str(_y)], CO['texts']['mwYearEndNotes'][str(_y)], NUM0); r += 1
PPE0 = gi(r, "6/30 毛 PP&E", "US$bn", D['ppeOpen'], "10-Q（含 CIP 11.9）[Verified]"); r += 1
CAPSC = gi(r, "每 MW 建置成本倍數（整體）", "%", D['capexScale'], "同時影響 CapEx、汰換與車隊折舊；反向 DCF 用；預設 100%", PCT); r += 1
r = phdr(r)
if PMW['capex'] == 'tokenomics':  # W2：Σ 新增世代占比 × IF_CapexIT（不含廠房）
    r = prow(r, "每 MW 建置成本", "US$m/MW", [f"=SUMPRODUCT({COLS[i]}{FLX['adopt']}:{COLS[i]}{FLX['adopt'] + len(GEN) - 1},$C${FLX['tkCap']}:$C${FLX['tkCap'] + len(GEN) - 1})" for i in range(5)], NUM1,
             "＝Σ 新增世代占比（B 區後『世代組合』）× Tokenomics IF_CapexIT（基準；GPU＋網路＋機房內裝，不含廠房——CRWV 機房以租約取得，廠房成本在租金）[Derived]", BLACK)
    r = prow(r, "每 MW 建置成本（v4.5 舊值，對照）", "US$m/MW", CO['scenarios']['capexTemplate']['costMW'], NUM1, "舊方法：以 FY26 校準重現全年指引 35–39（以公司 CapEx 指引回推；W2 起只作對照，不入模型）[Derived]")
else:
    r = prow(r, "每 MW 建置成本", "US$m/MW", CO['scenarios']['capexTemplate']['costMW'], NUM1, "GPU＋網路＋機房內裝；以 FY26 校準重現全年指引 35–39 [Derived]")
r = prow(r, "客戶預付占毛 CapEx", "%", CO['scenarios']['capexTemplate']['customerFund'], PCT, "H1 遞延收入淨流入 1.365÷CapEx 16.139＝8.5% [Derived]")
r = prow(r, "在帳現金租金（10-Q 到期表）", "US$bn", CO['leases']['onBalanceCash'], NUM, f"營業＋融資租賃未折現付款；«LASTYR» 後尚有 {CO['leases']['afterFY30']} [Verified]")
r = prow(r, "表外現金租金（未起租）", "US$bn", [0] * 5, NUM, "＝積極路徑 × 新增 MW 比例（A 區）[Assumed]", BLACK)
for i in range(5):
    L = COLS[i]
    ws.cell(row=IN["表外現金租金（未起租）"], column=3 + i,
            value=f"={L}{LEASE_HI}*MAX(0,{L}{_acc}-{LEASE_FL})/({L}{sc_rows['積極 8 GW']}-{LEASE_FL})")
r = prow(r, "JV 已承諾餘額出資", "US$bn", D['jvCommit'], NUM, f"10-Q：兩個 JV 承諾最高 1.7，2026 年內履行；«YTDA» 已付 {YA['jvSplit']['jv']:.2f} [Verified／Derived]")
r = prow(r, "JV 後續增資＋策略投資", "US$bn", CO['scenarios']['capexTemplate']['div'], NUM, "入夥後須追加出資，金額未揭露 [Assumed]")
r = prow(r, "JV／策略投資出資", "US$bn",
         [f"=C{IN['JV 已承諾餘額出資']}+C{IN['JV 後續增資＋策略投資']}".replace("C", COLS[0], 0)] * 5, NUM, "＝已承諾餘額＋後續增資", BLACK)
for i in range(5):
    L = COLS[i]
    ws.cell(row=IN["JV／策略投資出資"], column=3 + i, value=f"={L}{IN['JV 已承諾餘額出資']}+{L}{IN['JV 後續增資＋策略投資']}")
r = section(ws, r, "CapEx 推導明細（由 Accepted MW 連動）", level=2, collapsed=True)
accr = _acc
r = prow(r, "期初主動電力", "MW", [f"={MW_YE25}"] + [f"={COLS[i-1]}{accr}" for i in range(1, 5)], NUM0, f"{PERIODS[0]} 期初＝{MW_Y0} 年底 {D['mwYearEnd'][str(MW_Y0)]} MW", BLACK)
r = prow(r, "本期新增 MW", "MW", [f"={COLS[i]}{accr}-{COLS[i]}{r-1}" for i in range(5)], NUM0, None, BLACK)
r = prow(r, "次期新增 MW", "MW", [f"={COLS[i+1]}{accr}-{COLS[i]}{accr}" for i in range(4)] + [f"={MW31}"], NUM0, None, BLACK)
r = prow(r, "全年毛 CapEx（公式）", "US$bn",
         [f"=({COLS[i]}{r-2}*(1-{LAMBDA})+{COLS[i]}{r-1}*{LAMBDA})*{COLS[i]}{IN['每 MW 建置成本']}*{CAPSC}/1000" for i in range(5)],
         NUM, "＝(本期新增×(1−λ)＋次期新增×λ)×每 MW 成本", BLACK)
r = prow(r, "成長型 CapEx（模型期）", "US$bn",
         [f"=MAX({COLS[0]}{r-1},{FLOOR})-§H_CAPEX§"] + [f"={COLS[i]}{r-1}" for i in range(1, 5)], NUM,
         "FY26＝MAX(全年公式, 下限) − 1H 已認列 16.139", BLACK)
def _vintage(i):
    ys, x = sorted(MWY), "0"
    for y in reversed(ys):
        x = f"IF({PYEAR[i]}-{LIFE}={y},{MWY[y]}{'-' + MWY[y - 1] if y - 1 in MWY else ''},{x})"
    return x
r = prow(r, "GPU 汰換 CapEx", "US$bn",
         [f"={_vintage(i)}"
          f"*{COLS[i]}{IN['每 MW 建置成本']}*{CAPSC}/1000" for i in range(5)], NUM,
         "＝(本年 − 壽命) 那年新增的 MW × 每 MW 成本；取代已折舊完的設備", BLACK)
r = prow(r, "期初毛 PP&E", "US$bn", [f"={PPE0}"] + ["=0"] * 4, NUM, "之後＝前期期初＋前期成長型 CapEx（汰換不增加基礎）", BLACK)
for i in range(1, 5):
    ws.cell(row=IN["期初毛 PP&E"], column=3 + i, value=f"={COLS[i-1]}{IN['期初毛 PP&E']}+{COLS[i-1]}{IN['成長型 CapEx（模型期）']}")
r = section(ws, r, "CapEx 與折舊（結果；下游引用）", level=2)
r = prow(r, "毛 CapEx（模型期，下游引用此列）", "US$bn",
         [f"={COLS[i]}{IN['成長型 CapEx（模型期）']}+{COLS[i]}{IN['GPU 汰換 CapEx']}" for i in range(5)], NUM, "＝成長型＋汰換", BLACK, key="毛 CapEx")
r = prow(r, "D&A（車隊折舊）", "US$bn",
         [f"=({COLS[i]}{IN['期初毛 PP&E']}+0.5*{COLS[i]}{IN['成長型 CapEx（模型期）']})/{LIFE}*{COLS[i]}{IN['模型期長度（年）']}" for i in range(5)],
         NUM, "＝(期初毛 PP&E＋本期成長型×½)÷壽命×期間長度", BLACK, key="D&A（車隊）")
for k in ("毛 CapEx", "D&A（車隊）"):
    for i in range(5):
        ws.cell(row=IN[k], column=3 + i).font = Font(name="Arial", size=10, bold=True)
r += 1

# ---------------- E 融資 ----------------
r = section(ws, r, "E｜融資（期初現金、既有債務、瀑布參數）")
CASH0 = gi(r, "期初現金（2026-06-30）", "US$bn", D['cash'], "10-Q：現金 5.524＋有價證券 0.015；受限現金 1.38 不計 [Verified]"); r += 1
ATM = gi(r, "股權／可轉債金額", "US$bn", D['atm'], "9/17 可轉債淨現金約 3.2＋ATM 上限 35m 股 × $85.43（9/21 收盤；v4.3 現價更新後維持）≈ 3.0 [Verified]"); r += 1
ATMON = gi(r, "計入股權／可轉債（1=是）", "", int(D['includeAtm']), "關閉則 FY26 少 6.2bn 來源", NUM0); r += 1
FAC = gi(r, "未動用信用額度", "US$bn", D['facility'], "10-Q 可用額度 10.014（RCF＋DDTL 未動用）[Verified]"); r += 1
FACON = gi(r, "瀑布可動用未動用額度（1=是）", "", int(D['useFacility']), "瀑布第一順位；已承諾額度，不受 債務／backlog 上限限制", NUM0); r += 1
DEBTON = gi(r, "債務排程攤還（1=開）", "", int(D['includeDebt']), "10-Q 本金到期表；關閉＝假設全額再融資", NUM0); r += 1
KBL = gi(r, "債務／backlog 上限", "x", D['debtBacklog'], "預設 1.0x：依 2026 年實際融資組合（債務約 77%、股權約 23%）校準；公司簡報當前約 0.4x（敏感度）[Derived]", '0.00', True); r += 1
TERM = gi(r, "新簽合約年期", "年", D['ctrTerm'], "Q4 法說：加權平均合約約 5 年；決定 backlog 補入量 [Interested-party]", NUM0); r += 1
MINC = gi(r, "最低現金", "US$bn", D['minCash'], "期前融資的現金底線 [Assumed]"); r += 1
EQPX = gi(r, "股權發行價", "US$", D['eqPx'], f"預設＝現價（{CO['meta']['priceDate']} 收盤）[Verified]", USD); r += 1
EQDISC = gi(r, "股權發行折價", "%", D['eqDisc'], "大額增資的折讓 [Assumed]", PCT); r += 1
EQCAP = gi(r, "每年股權吸收上限（占現市值）", "%", D['eqCapPct'], "2026 年實際約 13%；輸入 ≥900% 視為無上限 [Assumed]", PCT); r += 1
EQSH = gi(r, "股權上限的股數基礎", "bn", D['eqCapShares'], "現市值＝發行參考價 × 此股數（company.json → defaults.eqCapShares）", '0.0000'); r += 1
JRATE = gi(r, "高息債利率（股權上限溢出）", "%", D['junkRate'], "高於 2031 票據 9.75%；CDS 中價 720bps [Assumed]", PCT); r += 1
CDS = gi(r, "CDS 中價", "bps", D['cds'], "2026-09 中價 720、Bid/Ask 690／750、區間 680–855 [User-provided]", NUM0); r += 1
CDSON = gi(r, "CDS 傳入新債利率（1=開）", "", int(D['cdsLink']), f"開啟：新債利率＋MAX(0,CDS−{_n(D['cdsBaseBp'])})×{_n(D['cdsPassThrough'])}/100。預設關以免與 9% 重複", NUM0); r += 1
CDSB = gi(r, "CDS 傳入門檻", "bps", D['cdsBaseBp'], "CDS 超過此值的部分才傳入新債利率 [Assumed]", NUM0); r += 1
CDSP = gi(r, "CDS 傳入比例", "%", D['cdsPassThrough'], "每 1bp CDS 超額傳入新債利率的比例 [Assumed]", PCT); r += 1
r = phdr(r)
r = prow(r, "排程還本（10-Q 本金表）", "US$bn", CO['debt']['amortization'], NUM, "本金合計 35.551；其後 14.896 [Verified]")
r = prow(r, "新債利率", "%", PCT_(M['rate']), PCT, "瀑布新債成本：2031 票據 9.75%、2032 9.625%、DDTL5.0 SOFR+450 [Derived]")
r = prow(r, "存量債務利息（下游引用此列）", "US$bn", ["=0"] * 5, NUM, "＝債務明細頁：平均本金×加權有效利率×期間長度＋可轉債利息＋FY26 校準", GREEN, key="存量債務利息")
DEBT_INT_ROW = IN["存量債務利息"]
r += 1

# ---------------- F 評價 ----------------
r = section(ws, r, "F｜評價（價格、股數、折現、倍數、權重）")
PX = gi(r, f"現價（{CO['meta']['priceDate']} 收盤）", "US$", V['price'], f"{CONS['priceReference']['source']}（{CONS['priceReference']['url']}）；9/17 收盤 79.88 [Verified]", USD); r += 1
SH = gi(r, "股數（含 ATM 上限）", "bn", V['shares'], "7/31 流通 0.5515＋ATM 35m [Verified]", '0.000'); r += 1
ND = gi(r, "淨負債（6/30）", "US$bn", V['netDebt'], "本金 35.551 − 現金 5.524 [Verified]"); r += 1
WACC = gi(r, "WACC", "%", V['wacc'], "[Assumed]", PCT); r += 1
GG = gi(r, "永續成長 g", "%", V['g'], "[Assumed]", PCT); r += 1
MAINT = gi(r, "終值維持性 CapEx 占 D&A", "%", V['maintRatio'], "終值不讓成長性 CapEx 偽裝成永續 FCF [Assumed]", PCT); r += 1
TAX = gi(r, "稅率", "%", V['tax'], "美國聯邦名目稅率 [Verified]", PCT); r += 1
NOL0 = gi(r, "期初 NOL（虧損扣抵）", "US$bn", V['nol'], f"«VMD» 累積虧損約 {V['nol']:.1f}；抵扣上限為應稅所得 {_n(V['nolUsePct'] * 100)}% [Derived]"); r += 1
NOLU = gi(r, "NOL 每年可抵用比例", "%", V['nolUsePct'], "抵扣上限占應稅所得的比例（美國規定 80%）[Verified]", PCT); r += 1
WCP = gi(r, "營運資金占營收增量", "%", V['wcPctOfRevGrowth'], "營運資金變動＝營收增量 × 此比例 [Assumed]", PCT); r += 1
EVEBITDA = gi(r, "EV/EBITDA 倍數", "x", V['evEbitda'], "穩態合理倍數約 3.4–6.0x（轉換率 ÷ (WACC − g)），6x 為上緣 [Assumed]", MULT); r += 1
EVY = gi(r, "EV/EBITDA 錨定年度（1＝FY27、2＝FY28、3＝FY29、4＝FY30）", "", V['evYear'], "«EVDISC»（目標價時點）；v3.6 起預設 FY29：接近穩態利潤率，與 6x 穩態倍數一致 [Assumed]", NUM0, True); r += 1
from openpyxl.worksheet.datavalidation import DataValidation as _DV
_dv = _DV(type="whole", operator="between", formula1="1", formula2="4", allow_blank=False, showErrorMessage=True, error="請輸入 1–4", errorTitle="錨定年度")
ws.add_data_validation(_dv); _dv.add(EVY.split("!")[1].replace("$", ""))
WDCF = gi(r, "權重：DCF", "%", CO['methodology']['blendWeights']['dcf'], "", PCT); r += 1
WEVE = gi(r, "權重：EV/EBITDA", "%", CO['methodology']['blendWeights']['pe'], "", PCT); r += 1
RM_LO = gi(r, "方法區間倍數：下端", "x", min(CO['methodology']['rangeMultiples']), "v4.1 方法區間：目前輸入、EV/EBITDA 倍數換成兩端；6x 為穩態倍數上緣 [Assumed]", MULT); r += 1
RM_HI = gi(r, "方法區間倍數：上端", "x", max(CO['methodology']['rangeMultiples']), "", MULT); r += 1
SHG = gi(r, "股數年增（SBC 稀釋）", "%", 0.01, "SBC 約 0.66bn／年 ÷ $80 ≈ 1.4%／年，取 1% [Derived]", PCT); r += 1
SBCY = gi(r, "年度 SBC", "US$bn", V['sbc'], "Q2 0.165／季年化約 0.66，取 0.7 [Derived]"); r += 1
r = section(ws, r, "DCF 股權為負時的處理（0 截斷／選擇權）", level=2)
DMODE = gi(r, "DCF 下限方式（1＝0 截斷、2＝選擇權）", "", (1 if V['dcfMode'] == 'zero' else 2), "0 截斷：MAX(0, 股權價值)；選擇權：Merton，股權＝以企業價值為標的、淨負債為履約價的買權", NUM0, True); r += 1
SIGMA = gi(r, "企業價值波動率 σ", "%", V['sigma'], "選擇權法用 [Assumed]", PCT); r += 1
RF = gi(r, "無風險利率", "%", V['rf'], "選擇權法用 [Assumed]", PCT); r += 1
OPTT = gi(r, "選擇權期間（年）", "年", V['optT'], "至 FY30 末", NUM1); r += 1
# v4.5（5a-1）：期間與日期——由 company.json → calendar 推算（calendar_q.py，HTML 共用同一份結果）；折現年數以月計
r = section(ws, r, "期間與日期（company.json → calendar 推算；勿手改）", level=2)
_calc = lambda row, name, unit, v, note, fmt=NUM0: gi(row, name, unit, v, note, fmt, font=BLACK)
CAL_FYE = _calc(r, "財年結束月份", "月", CAL['fiscalYearEndMonth'], "calendar.fiscalYearEndMonth"); r += 1
CAL_Q = _calc(r, "最新已申報季度（年度首期依此滾動）", "", CAL['latestQuarterFiled'] or "無", "calendar.latestQuarterFiled；財報新聞稿只更新季度層，10-Q／10-K 申報後年度首期才滾動", "@"); r += 1
CAL_VD = _calc(r, "評價日（已申報季度的季末）", "", CAL['valuationDate'], "DCF 折現、淨負債與股數的基準日", "@"); r += 1
CAL_TD = _calc(r, "目標價時點", "", CAL['targetDate'],
               "首期之後第一個完整財年末" if CAL['targetHorizon'] == 'firstFullYearEnd' else "評價日＋12 個月", "@"); r += 1
CAL_TT = _calc(r, "目標價時點距評價日", "年", CAL['targetT'], "", NUM); r += 1
_calc(r, "模型期間", "", f"{PERIODS[0]}–{PERIODS[-1]}", "首期＋4 個完整財年（calendar.modelYears）", "@"); r += 1
_calc(r, "資料更新日", "", CO['meta']['updateDate'], "company.json → meta.updateDate", "@"); r += 1
_calc(r, "現價日（收盤）", "", CO['meta']['priceDate'], "company.json → meta.priceDate", "@"); r += 1
OFF = _calc(r, "錨定期折回年數調整（目標價時點 − 首期長度）", "年", CAL['evOffset'],
            "EV/EBITDA 腿：錨定期 k 的期末折回目標價時點的年數＝k − 此值", NUM); r += 1
r = phdr(r, first="期間與日期")
CAL_R = {}
for key, nm, fmt in [('periodEnd', "期末日", "@"), ('tEnd', "折現年數：評價日 → 期末", NUM), ('tStart', "折現年數：評價日 → 期初", NUM)]:
    r = prow(r, nm, "" if key == 'periodEnd' else "年", CAL[key], fmt, "月數 ÷ 12", font=BLACK, key='cal_' + key)
    CAL_R[key] = IN['cal_' + key]
assert len(PERIODS) == len(COLS), 'Excel 欄位固定為 5 期（首期＋4 個完整財年）'
# v4.2：評等門檻（company.json → methodology.rating）；評價頁的評等代碼、賣出門檻價、判斷句與「檢查」頁都引用這裡
RT = CO['methodology']['rating']
RT_PCT = lambda x: (lambda v: f"{v:.0f}%" if v == int(v) else f"{v:.1f}%")(round(abs(x) * 100, 4))
_MT = lambda x: f'IF({x}=INT({x}),TEXT({x},"0"),TEXT({x},"0.0"))'  # 倍數文字：整數不帶小數（與 HTML multTxt 同格式）
_PC = lambda x: _MT(f"ROUND(ABS({x})*100,4)") + '&"%"'  # 門檻百分比文字（絕對值）：與 HTML pctQ 同格式
r = section(ws, r, "評等門檻（空間＝加權目標價 ÷ 現價 − 1）", level=2)
RT_BUY = gi(r, "買進：空間下限", "%", RT['buyUpsideMin'], "空間 ≥ 此值、終值占比 < 上限，且股權募資未超標 → 買進 [Assumed]", PCT); r += 1
RT_TV = gi(r, "買進：終值占 EV 上限", "%", RT['buyTvShareMax'], "終值占比達此值以上不給買進 [Assumed]", PCT); r += 1
RT_SELL = gi(r, "賣出：空間上限", "%", RT['sellUpsideMax'], "空間 ≤ 此值 → 賣出；賣出門檻價＝現價 ×（1 ＋ 此值） [Assumed]", PCT); r += 1
RT_EQ = gi(r, "股權募資上限（÷ 現市值）", "x", RT['equityRaiseMaxMult'], "五期股權募資超過現市值的此倍數：市場吸收能力存疑，禁止買進 [Assumed]", MULT); r += 1
RT_SELL2 = gi(r, "股權募資超標時的賣出：空間上限", "%", RT['sellUpsideMaxIfEquityOver'], "股權募資超標且空間 ≤ 此值 → 賣出 [Assumed]", PCT); r += 1
RT_MG = gi(r, "判斷句：距賣出門檻揭露（占現價）", "%", RT['sellMarginAlert'], "任一情境與賣出門檻的距離小於現價的此比例時，判斷句揭露該距離 [Assumed]", PCT); r += 1
RT_TVW = gi(r, "終值占 EV 警示", "%", RT['tvShareWarn'], "超過此值表示結論由終值假設決定（「檢查」頁與 DCF 說明） [Assumed]", PCT); r += 1
r += 1

# ---------------- G 參考：1H26 實際 ----------------
r = section(ws, r, "G｜參考：«YTDL» 實際（10-Q，«FYSTART» 至 «VMMDD»；硬編碼，勿改）", collapsed=True)
h1_rows = [(lab, "US$" if k in ('eps', 'ngEps') else "US$bn", YA[k], YA['notes'][k]) for lab, k in [
    ("«PREVFYE» 現金", 'cash1231'), ("營運現金流 CFO", 'cfo'), ("現金購置 PP&E", 'cashCapex'), ("CapEx 認列（含 OEM 融資）", 'capex'),
    ("JV＋策略投資出資", 'jv'), ("借款", 'borrow'), ("還款", 'debtRepaid'), ("capped call 支出", 'cappedCall'), ("私募股權", 'equity'),
    ("利息費用", 'interest'), ("租賃現金支付", 'leasePaid'), ("營收", 'revenue'), ("GAAP 營業損益", 'opInc'), ("淨損", 'ni'),
    ("遞延收入淨流入", 'prepay'), ("D&A", 'da'), ("SBC", 'sbc'), ("GAAP EPS", 'eps'), ("EPS（加回 SBC）", 'ngEps')]] + [
    # v4.3：只用於「摘要」頁與市場共識的 FY26 調整後 EBITDA 比較；模型 GAAP 損益與評價不使用
    ("調整後 EBITDA（«YTDQS»）", "US$bn", YA['adjEbitda'],
     "＋".join(f"{k.upper()} {YA['adjEbitdaMeta'][k]:.3f}" for k in sorted(YA['adjEbitdaMeta']) if re.fullmatch(r'q\d', k)) + "；"
     + "；".join(f"{x['quarter']} {x['name']}" for x in YA['adjEbitdaMeta']['sources'])
     + f"；{YA['adjEbitdaMeta']['crossCheck']}。{YA['adjEbitdaMeta']['note']} [{YA['adjEbitdaMeta']['tag']}]"),
]
start_h1 = r
for nm, un, v, nt in h1_rows:
    gi(r, nm, un, v, nt)
    r += 1
(H_CASH1231, H_CFO, H_CCAPEX, H_CAPEX, H_JV, H_BORROW, H_REPAY, H_CAP, H_EQ,
 H_INT, H_LEASE, H_REV, H_OPINC, H_NI, H_PREPAY, H_DA, H_SBC, H_EPS, H_NGEPS, H_ADJEB) = [f"'輸入與假設'!$C${start_h1 + i}" for i in range(len(h1_rows))]
# resolve forward references
for row in ws.iter_rows():
    for cc in row:
        if isinstance(cc.value, str) and "§H_CAPEX§" in cc.value:
            cc.value = cc.value.replace("§H_CAPEX§", H_CAPEX)

IN_ref = {k: f"'輸入與假設'!{{col}}${v}" for k, v in IN.items()}


def inref(name, i):
    return f"'輸入與假設'!{COLS[i]}${IN[name]}"


# =====================================================================
# 3. 產能與收入
# =====================================================================
ws = wb.create_sheet("運營_產能與收入")
ws.column_dimensions["A"].width = 42
ws.column_dimensions["B"].width = 12
for c in COLS:
    ws.column_dimensions[c].width = 13
ws.column_dimensions["I"].width = 96
ws["A1"] = "產能與收入 — 合約排程 vs 機房產能"
ws["A1"].font = TITLE
ws["A2"] = "兩條線獨立產生後相減：排程>容量→瓶頸（收不到）；容量>排程→剩餘產能（要靠新簽約賣掉）"
ws["A2"].font = SMALL
ws["A3"] = "«P0» 欄口徑：MW 為年底存量；標示「（模型期）」的收入列只含 «STUBYR» «STUBW»——RPO 桶自 «VD» 起算，«YTD» 收入已實現、不再受產能約束。«YTD» 實際單列，兩者相加即 «P0» 全年（見本頁「FY26 全年」列）。"
ws["A3"].font = SMALL
r = 5
r = period_header(ws, r)
CR = {}

def crow(name, unit, formula_fn, fmt=NUM, font=BLACK, note=None):
    global r
    ws.cell(row=r, column=1, value=name).font = BLACK
    ws.cell(row=r, column=2, value=unit).font = SMALL
    for i in range(5):
        c = ws.cell(row=r, column=3 + i, value=formula_fn(i))
        c.font = font
        c.number_format = fmt
        c.border = BOX
    if note:
        ws.cell(row=r, column=9, value=note).font = SMALL
    CR[name] = r
    r += 1


crow("模型期間", "", lambda i: (f'="«VSTART»~«STUBEND»"' if i == 0 else f'="{CAL["periodEnd"][i][:4]} 全年"'), NUM, BLACK,
     "«P0» 的模型期只有«STUBW»；«YTD» 已實現的營收單列於下方")
crow("模型期長度（年）", "", lambda i: f"={inref('模型期長度（年）', i)}", NUM, GREEN,
     "只用於下方三處：容量上限、電力、維護——把年率換算成該期金額")
r = section(ws, r, "（A）機房：容量上限")
crow("Accepted MW（已驗收，單調不減）", "MW",
     lambda i: f"=MAX({inref('Accepted MW（期末主動電力）', i)},0)" if i == 0
     else f"=MAX({inref('Accepted MW（期末主動電力）', i)},{COLS[i-1]}{CR['模型期長度（年）']+2 if False else r-0})",
     NUM0, BLACK, "引擎規則：不得低於前期（產能不會消失）")
# fix the self-reference: rewrite Accepted row properly
acc_row = CR["Accepted MW（已驗收，單調不減）"]
for i in range(5):
    if i == 0:
        f = f"=MAX(0,{inref('Accepted MW（期末主動電力）', i)})"
    else:
        f = f"=MAX({inref('Accepted MW（期末主動電力）', i)},{COLS[i-1]}{acc_row})"
    ws.cell(row=acc_row, column=3 + i, value=f)

crow("Billable MW（上限為 Accepted）", "MW",
     lambda i: f"=MIN(MAX(0,{inref('Billable MW', i)}),{COLS[i]}{acc_row})", NUM0, BLACK,
     "引擎規則：可計費不得超過已驗收")
bil_row = CR["Billable MW（上限為 Accepted）"]
crow("期初在役 MW", "MW",
     lambda i: (f"={MW0}" if i == 0 else f"={COLS[i-1]}{bil_row}"), NUM0, BLACK, "前期期末＝本期期初")
beg_row = CR["期初在役 MW"]
crow("平均在役 MW（«P0» 欄為«STUBW»平均）", "MW",
     lambda i: f"=IF({AVGON}=1,({COLS[i]}{beg_row}+{COLS[i]}{bil_row})/2,{COLS[i]}{bil_row})", NUM0, BLACK,
     "開關在輸入頁：1＝(期初+期末)/2，0＝期末全期化")
avg_row = CR["平均在役 MW（«P0» 欄為«STUBW»平均）"]
crow("利用率", "%", lambda i: f"={inref('利用率', i)}", PCT, GREEN)
crow("每 MW 年收入", "US$bn/MW", lambda i: f"={inref('每 MW 年收入', i)}", '0.0000', GREEN)
crow("容量上限（模型期最多能交付的收入）", "US$bn",
     lambda i: f"={COLS[i]}{avg_row}*{inref('每 MW 年收入', i)}*{REVSC}*{inref('利用率', i)}*{COLS[i]}{CR['模型期長度（年）']}",
     NUM, BOLD, "＝平均在役 MW × 每 MW 年收入 × 利用率 × 期間係數。只看機房，不看合約")
cap_row = CR["容量上限（模型期最多能交付的收入）"]
for i in range(5):
    ws.cell(row=cap_row, column=3 + i).font = Font(name="Arial", size=10, bold=True)

r = section(ws, r, "（B）合約：排程 RPO")
crow("排程 RPO（模型期，依 10-Q 桶分攤）", "US$bn",
     lambda i: f"=({RPO0}+{RPOADD})*({inref('RPO 桶權重', i)}/{WSUM})*{RP}", NUM, BLACK,
     "＝(«VMD» RPO＋Q3 新增)×本期權重÷權重合計×五期認列比例。只看合約，不看機房")
sch_row = CR["排程 RPO（模型期，依 10-Q 桶分攤）"]
ai_row = sch_row

r = section(ws, r, "（C）兩條線相減")
crow("產能瓶頸（模型期：排程 − 容量）", "US$bn",
     lambda i: f"=MAX(0,{COLS[i]}{ai_row}-{COLS[i]}{cap_row})", NUM, BOLD,
     "合約有、機房沒有 → 本期收不到，且不遞延到下期（保守處理）")
bot_row = CR["產能瓶頸（模型期：排程 − 容量）"]
for i in range(5):
    ws.cell(row=bot_row, column=3 + i).font = Font(name="Arial", size=10, bold=True)
crow("期初 RPO 轉換收入（模型期）", "US$bn",
     lambda i: f"={COLS[i]}{sch_row}-{COLS[i]}{bot_row}", NUM, BLACK,
     "＝排程 − 瓶頸。當瓶頸>0 時，此值恆等於容量上限")
rev_row = CR["期初 RPO 轉換收入（模型期）"]
crow("未被期初 RPO 占用的產能", "US$bn",
     lambda i: f"=MAX(0,{COLS[i]}{cap_row}-{COLS[i]}{ai_row})", NUM, BLACK,
     "容量 − 排程。機房有、期初合約沒占滿 → 要靠新簽約")
free_row = CR["未被期初 RPO 占用的產能"]
crow("新產能簽約率", "%", lambda i: f"={inref('新產能簽約率', i)}", PCT, GREEN)
crow("新簽約收入", "US$bn",
     lambda i: f"={COLS[i]}{free_row}*{inref('新產能簽約率', i)}", NUM, BLACK,
     "尚未簽署的合約。FY28 起模型收入主要來自此列——8 GW 敘事的風險集中點")
newrev_row = CR["新簽約收入"]
crow("未售產能（浪費）", "US$bn",
     lambda i: f"=MAX(0,{COLS[i]}{cap_row}-{COLS[i]}{rev_row}-{COLS[i]}{newrev_row})", NUM, BLACK,
     "蓋好但賣不掉。Bernstein：簽約電力 74% 位於 Tier 3/4 市場")
crow("損益用算力收入（模型期＝RPO 轉換＋新簽約）", "US$bn",
     lambda i: f"={COLS[i]}{rev_row}+{COLS[i]}{newrev_row}", NUM, BOLD,
     "損益與評價頁的算力收入由此連結，確保現金與損益同一組數字")
isrev_row = CR["損益用算力收入（模型期＝RPO 轉換＋新簽約）"]

crow("«YTDL» 實際營收（10-Q，已實現）", "US$bn",
     lambda i: (f"={H_REV}" if i == 0 else "=0"), NUM, GREEN,
     "«FYSTART» 至 «VMMDD» 已認列營收。已實現，不受«STUBW»產能約束，故不進入上方的瓶頸計算")
h1rev_row = CR["«YTDL» 實際營收（10-Q，已實現）"]
crow("«P0» 全年總營收（«YTD» 實際＋模型期算力＋服務）", "US$bn",
     lambda i: (f"={COLS[i]}{h1rev_row}+{COLS[i]}{isrev_row}+{inref('非算力服務營收', i)}"), NUM, BOLD,
     "FY26 欄可直接對照公司全年指引 12.4–13.2；FY27 以後即為該年度總營收")
fy_rev_row = CR["«P0» 全年總營收（«YTD» 實際＋模型期算力＋服務）"]
for i in range(5):
    ws.cell(row=fy_rev_row, column=3 + i).font = Font(name="Arial", size=10, bold=True)
ws.cell(row=fy_rev_row, column=3).fill = FILL_KEY

r = section(ws, r, "（D）收入 → 現金：三道折扣")
crow("客戶違約率", "%", lambda i: f"={inref('客戶違約率', i)}", PCT, GREEN)
crow("回收率", "%", lambda i: f"={inref('回收率', i)}", PCT, GREEN)
crow("信用損失（期初 RPO 部分）", "US$bn",
     lambda i: f"={COLS[i]}{rev_row}*{inref('客戶違約率', i)}*(1-{inref('回收率', i)})", NUM, BLACK,
     "＝收入 × 違約率 × (1−回收率)")
loss_row = CR["信用損失（期初 RPO 部分）"]
crow("收現（期初 RPO 部分）", "US$bn", lambda i: f"={COLS[i]}{rev_row}-{COLS[i]}{loss_row}", NUM, BLACK)
coll_row = CR["收現（期初 RPO 部分）"]
crow("EBITDA 率（損益與資金共用）", "%", lambda i: f"={inref('EBITDA 率', i)}", PCT, GREEN, "起始→穩態線性爬升")
crow("模型期總營收（算力＋服務）", "US$bn", lambda i: f"={COLS[i]}{isrev_row}+{inref('非算力服務營收', i)}", NUM, BLACK)
totrev_row = CR["模型期總營收（算力＋服務）"]
crow("EBITDAR 率（EBITDA 率＋租金÷營收）", "%", lambda i: "=0", PCT, BOLD,
     "租金前的 EBITDA 率。EBITDA 已扣租金，而資金模型把租金列在支出端，所以要加回，避免重複扣除")
cm_row = CR["EBITDAR 率（EBITDA 率＋租金÷營收）"]
crow("RPO 現金（可支應資本用途）", "US$bn",
     lambda i: f"={COLS[i]}{coll_row}*{COLS[i]}{cm_row}", NUM, BOLD,
     "×EBITDAR 率。租金與利息在支出頁單獨列示")
rpocash_row = CR["RPO 現金（可支應資本用途）"]
crow("新簽約現金", "US$bn",
     lambda i: f"={COLS[i]}{newrev_row}*(1-{inref('客戶違約率', i)}*(1-{inref('回收率', i)}))*{COLS[i]}{cm_row}", NUM, BLACK,
     "同樣扣信用損失、乘 EBITDAR 率")
newcash_row = CR["新簽約現金"]

r += 1
ws.cell(row=r, column=1, value="讀法：«P0»（«STUB»）與 «P1» 瓶頸>0，合約多於機房；FY28 起反轉，收入愈來愈靠『新簽約收入』這一列。").font = BOLD
r += 1
ws.cell(row=r, column=1, value="把『新產能簽約率』調到 0，就能看到只靠期初 RPO 的資金缺口有多大——這是最重要的一個壓力測試。").font = SMALL
CAP = {"cm": cm_row, "totrev": totrev_row, "cap": cap_row, "sch": sch_row, "bot": bot_row, "rev": rev_row, "newrev": newrev_row,
       "isrev": isrev_row, "rpocash": rpocash_row, "newcash": newcash_row, "acc": acc_row,
       "loss": loss_row, "coll": coll_row, "avg": avg_row}

# =====================================================================
# 4. 支出與資金
# =====================================================================
ws = wb.create_sheet("各期收支")
ws.column_dimensions["A"].width = 42
ws.column_dimensions["B"].width = 12
for c in COLS:
    ws.column_dimensions[c].width = 13
ws.column_dimensions["I"].width = 96
ws["A1"] = "支出與資金 — 五層支出、四層來源、現金橋"
ws["A1"].font = TITLE
ws["A2"] = "«P0» 欄＝«YTD» 實際（10-Q 現金流量表口徑）＋«STUBW»模型；«P1» 以後為純模型。瀑布以「融資前現金」決定需求並直接由各支出列加總，因此沒有循環參照"
ws["A2"].font = SMALL
r = 4
r = period_header(ws, r)
FR = {}


def frow(name, unit, fn, fmt=NUM, font=BLACK, note=None, bold=False):
    global r
    ws.cell(row=r, column=1, value=name).font = BOLD if bold else BLACK
    ws.cell(row=r, column=2, value=unit).font = SMALL
    for i in range(5):
        c = ws.cell(row=r, column=3 + i, value=fn(i))
        c.font = Font(name="Arial", size=10, bold=True) if bold else font
        c.number_format = fmt
        c.border = BOX
    if note:
        ws.cell(row=r, column=9, value=note).font = SMALL
    FR[name] = r
    r += 1


r = section(ws, r, "支出（五層）")
frow("① 毛 CapEx（認列，備忘）", "US$bn",
     lambda i: (f"={H_CAPEX}+{inref('毛 CapEx', i)}" if i == 0 else f"={inref('毛 CapEx', i)}"), NUM, BLACK,
     f"«P0»＝«YTD» 實際認列 {CO['ytdActual']['capex']}＋«STUB» 模型；對照公司全年指引 {CO['callFacts']['capexLo']}–{CO['callFacts']['capexHi']}。備忘列：用途合計採現金口徑，不用這一列")
frow("　客戶預付率", "%", lambda i: f"={inref('客戶預付占毛 CapEx', i)}", PCT, GREEN)
frow("　客戶預付金額（抵減，«STUB» 起）", "US$bn",
     lambda i: f"={inref('毛 CapEx', i)}*{COLS[i]}{FR['　客戶預付率']}", NUM, BLACK,
     "«YTD» 的預付已含在實際 CFO 內，不重複計入")
frow("① CapEx（用途用：«YTD» 現金／«STUB» 毛額）", "US$bn",
     lambda i: (f"={H_CCAPEX}+{inref('毛 CapEx', i)}" if i == 0 else f"={inref('毛 CapEx', i)}"), NUM, BLACK,
     f"«YTD» 採現金流量表的現金購置 {CO['ytdActual']['cashCapex']}；«STUB» 起採毛額，客戶預付在來源端抵回（避免重複計入）")
frow("② 在帳現金租金（備忘，«YTD» 已含在 CFO）", "US$bn",
     lambda i: f"={inref('在帳現金租金（10-Q 到期表）', i)}", NUM, GREEN,
     f"見『租賃與承諾』頁；«LASTYR» 後尚有 {CO['leases']['afterFY30']}。«P0» 欄僅 «STUB»：«YTD» 租金 {CO['ytdActual']['leasePaid']} 已含在實際 CFO 內")
frow("② 表外現金租金（未起租）", "US$bn", lambda i: f"={inref('表外現金租金（未起租）', i)}", NUM, GREEN,
     "未起租 35.5＋單站上限 14.7＋355MW 造價計租的現金路徑")
frow("　租金合計", "US$bn",
     lambda i: f"={COLS[i]}{FR['② 在帳現金租金（備忘，«YTD» 已含在 CFO）']}+{COLS[i]}{FR['② 表外現金租金（未起租）']}", NUM, BLACK)
_wsc = wb["運營_產能與收入"]; _wsi = wb["輸入與假設"]
for i in range(5):
    _wsc.cell(row=CAP["cm"], column=3 + i,
              value=f"={inref('EBITDA 率', i)}+'各期收支'!{COLS[i]}{FR['　租金合計']}/MAX(0.01,{COLS[i]}{CAP['totrev']})")
    c = _wsi.cell(row=IN["非算力服務現金"], column=3 + i,
                  value=f"={inref('非算力服務營收', i)}*'運營_產能與收入'!{COLS[i]}{CAP['cm']}"); c.font = BLACK
frow("③ 存量債務利息（備忘，«YTD» 已含在 CFO）", "US$bn",
     lambda i: f"={inref('存量債務利息', i)}", NUM, GREEN,
     # v4.5：首期模型部分的利息由本列算出，說明句改為文字公式（滾動後隨模型值更新）
     f'="«VMD» 債務餘額的利息。«P0» 全年利息＝«YTD» 實際 {YA["interest"]}＋«STUB» "&TEXT(C{r},"0.0")&"＝"&TEXT({YA["interest"]}+C{r},"0.00")&"（見下方備忘）"')
frow("　期初累積現金", "US$bn", lambda i: f"={H_CASH1231}", NUM, BLACK,
     "FY26 自 2025-12-31 的 3.127 起算；之後＝前期期末累積現金")
beg_cash = FR["　期初累積現金"]
frow("　新債利率", "%",
     lambda i: f"=IF({CDSON}=1,{inref('新債利率', i)}+MAX(0,{CDS}-450)/100*0.004,{inref('新債利率', i)})",
     PCT, BLACK, f"CDS 開關：＋MAX(0,CDS−{_n(D['cdsBaseBp'])})×{_n(D['cdsPassThrough'])}bp/bp")
od_rate = FR["　新債利率"]
frow("③ 新債利息（瀑布）", "US$bn", lambda i: "=0", NUM, BLACK,
     "＝新債利率×期間長度×(期初新債＋本期舉借)；見下方「期前融資瀑布」")
newint_row = FR["③ 新債利息（瀑布）"]
frow("　利息合計", "US$bn",
     lambda i: f"={COLS[i]}{FR['③ 存量債務利息（備忘，«YTD» 已含在 CFO）']}+{COLS[i]}{newint_row}", NUM, BLACK)
int_row = FR["　利息合計"]
frow("④ JV／策略投資出資", "US$bn",
     lambda i: (f"={H_JV}+{inref('JV／策略投資出資', i)}" if i == 0 else f"={inref('JV／策略投資出資', i)}"), NUM, BLACK,
     f"«P0»＝«YTD» 實際 {YA['jv']}（JV {YA['jvSplit']['jv']:.3f}＋策略投資 {YA['jvSplit']['strategic']:.3f}）＋«STUB» 模型。10-Q 已承諾 1.7bn。{CO['meta']['ticker']} 不發股息")
frow("　電力（overlay）", "US$bn",
     lambda i: "=0" if PMW['cost'] == 'bottomUp' else (f"=IF({OVERLAY}=1,'運營_產能與收入'!{COLS[i]}{CAP['acc']}*8760*{inref('PUE', i)}*{inref('電價', i)}/1000000000*"
                f"'運營_產能與收入'!{COLS[i]}{CR['模型期長度（年）']},0)"), NUM, BLACK,
     "W2 由下而上：電費與維護已含在營運成本，overlay 自動停用（避免重複扣）" if PMW['cost'] == 'bottomUp' else "＝Accepted MW×8760×PUE×電價÷1e9×期間係數。CRWV 電費多經房東轉嫁，預設關閉")
frow("　維護（overlay）", "US$bn",
     lambda i: "=0" if PMW['cost'] == 'bottomUp' else (f"=IF({OVERLAY}=1,'運營_產能與收入'!{COLS[i]}{CAP['acc']}*{inref('維護成本', i)}/1000*"
                f"'運營_產能與收入'!{COLS[i]}{CR['模型期長度（年）']},0)"), NUM, BLACK)
frow("⑤ 排程還本（10-Q 本金表）", "US$bn",
     lambda i: (f"={H_REPAY}+IF({DEBTON}=1,{inref('排程還本（10-Q 本金表）', i)},0)" if i == 0
                else f"=IF({DEBTON}=1,{inref('排程還本（10-Q 本金表）', i)},0)"), NUM, BLACK,
     f"«P0»＝«YTD» 實際還款 {CO['ytdActual']['debtRepaid']}＋«STUB» 到期表 {CO['debt']['amortization'][0]}")
frow("⑥ capped call 支出（«YTD» 實際）", "US$bn",
     lambda i: (f"={H_CAP}" if i == 0 else "=0"), NUM, BLACK, "Q2 4.0bn 可轉債的 capped call 現金支出")
debt_row = FR["⑤ 排程還本（10-Q 本金表）"]
frow("用途合計（現金口徑）", "US$bn",
     lambda i: (f"={COLS[i]}{FR['① CapEx（用途用：«YTD» 現金／«STUB» 毛額）']}+{COLS[i]}{FR['　租金合計']}+{COLS[i]}{int_row}+"
                f"{COLS[i]}{FR['④ JV／策略投資出資']}+{COLS[i]}{FR['　電力（overlay）']}+"
                f"{COLS[i]}{FR['　維護（overlay）']}+{COLS[i]}{debt_row}+{COLS[i]}{FR['⑥ capped call 支出（«YTD» 實際）']}"),
     NUM, BLACK,
     "«P0» 的 «YTD» 部分採 10-Q 現金流量表口徑（現金購置、還款、JV、capped call）；«YTD» 的利息與租金已含在 CFO 內，不重複列",
     bold=True)
uses_row = FR["用途合計（現金口徑）"]

r = section(ws, r, "來源（四層）")
frow("Ⓐ0 «YTDL» 實際營運現金流（CFO）", "US$bn",
     lambda i: (f"={H_CFO}" if i == 0 else "=0"), NUM, GREEN,
     "10-Q 實際值，取代模型推算；已含 «YTD» 的利息、租金與客戶預付")
frow("Ⓐ RPO 現金（«STUB» 起）", "US$bn", lambda i: f"='運營_產能與收入'!{COLS[i]}{CAP['rpocash']}", NUM, GREEN)
frow("Ⓑ 新簽約現金", "US$bn", lambda i: f"='運營_產能與收入'!{COLS[i]}{CAP['newcash']}", NUM, GREEN)
frow("Ⓒ 非算力服務現金", "US$bn", lambda i: f"={inref('非算力服務現金', i)}", NUM, GREEN)
frow("Ⓓ 客戶預付（«STUB» 起）", "US$bn",
     lambda i: f"={COLS[i]}{FR['　客戶預付金額（抵減，«STUB» 起）']}", NUM, BLACK)
frow("營運來源合計", "US$bn",
     lambda i: (f"={COLS[i]}{FR['Ⓐ0 «YTDL» 實際營運現金流（CFO）']}+{COLS[i]}{FR['Ⓐ RPO 現金（«STUB» 起）']}+"
                f"{COLS[i]}{FR['Ⓑ 新簽約現金']}+{COLS[i]}{FR['Ⓒ 非算力服務現金']}+{COLS[i]}{FR['Ⓓ 客戶預付（«STUB» 起）']}"),
     NUM, BLACK, bold=True)
srcop_row = FR["營運來源合計"]
frow("Ⓔ 股權／可轉債（融資）", "US$bn",
     lambda i: (f"={H_EQ}+IF({ATMON}=1,{ATM},0)" if i == 0 else "=0"), NUM, BLACK,
     CO['texts']['fy0EquityNote'])
atm_row = FR["Ⓔ 股權／可轉債（融資）"]
frow("Ⓕ «YTDL» 實際借款（融資）", "US$bn",
     lambda i: (f"={H_BORROW}" if i == 0 else "=0"), NUM, GREEN,
     f"10-Q：«YTDA» 借款 {YA['borrow']}。«YTD» 的缺口實際上就是靠舉債填補的")
borrow_row = FR["Ⓕ «YTDL» 實際借款（融資）"]
frow("Ⓖ 瀑布：新債（額度＋資產層）", "US$bn", lambda i: "=0", NUM, BLACK, "見下方「期前融資瀑布」")
fac_row = FR["Ⓖ 瀑布：新債（額度＋資產層）"]
frow("Ⓗ 瀑布：股權募資", "US$bn", lambda i: "=0", NUM, BLACK, "債務用罄後的殘差；每年不超過股權吸收上限")
eqr_row = FR["Ⓗ 瀑布：股權募資"]
frow("Ⓘ 瀑布：高息債（股權上限溢出）", "US$bn", lambda i: "=0", NUM, BLACK, "股權超過每年上限的部分；不受 backlog 上限約束")
jr_row = FR["Ⓘ 瀑布：高息債（股權上限溢出）"]
frow("總來源（含融資）", "US$bn",
     lambda i: f"={COLS[i]}{srcop_row}+{COLS[i]}{atm_row}+{COLS[i]}{borrow_row}+{COLS[i]}{fac_row}+{COLS[i]}{eqr_row}+{COLS[i]}{jr_row}",
     NUM, BLACK, bold=True)
src_row = FR["總來源（含融資）"]

r = section(ws, r, "缺口與現金橋")
frow("　«YTD» 其他／受限現金調節", "US$bn",
     lambda i: (f"={CASH0}-({H_CASH1231}+{H_CFO}+{H_BORROW}+{H_EQ}-{H_CCAPEX}-{H_JV}-{H_REPAY}-{H_CAP})"
                if i == 0 else "=0"), NUM, BLACK,
     "使 «YTD» 實際流量接回 «VMD» 現金餘額；差額來自受限現金重分類與未逐項列出的項目")
plug_row = FR["　«YTD» 其他／受限現金調節"]
frow("營運缺口（不含融資、不含還本）", "US$bn",
     lambda i: f"={COLS[i]}{srcop_row}-({COLS[i]}{uses_row}-{COLS[i]}{debt_row}-{COLS[i]}{FR['⑥ capped call 支出（«YTD» 實際）']})",
     NUM, BLACK, "本業能不能自給：營運來源 −（用途 − 排程還本 − capped call）", bold=True)
opgap_row = FR["營運缺口（不含融資、不含還本）"]
frow("期間缺口（含融資）", "US$bn",
     lambda i: f"={COLS[i]}{src_row}-{COLS[i]}{uses_row}+{COLS[i]}{plug_row}", NUM, BLACK)
gap_row = FR["期間缺口（含融資）"]
frow("期末累積現金", "US$bn",
     lambda i: f"={COLS[i]}{beg_cash}+{COLS[i]}{gap_row}", NUM, BLACK,
     "期前融資：每期末不低於最低現金", bold=True)
cum_row = FR["期末累積現金"]
# 期初累積現金鏈結
for i in range(5):
    ws.cell(row=beg_cash, column=3 + i,
            value=(f"={H_CASH1231}" if i == 0 else f"={COLS[i-1]}{cum_row}"))

r += 1
r = section(ws, r, "期前融資瀑布（缺口在需要前一期先融好：額度 → 資產層新債 → 股權）")
LROW = f"'運營_產能與收入'!{{c}}{CR['模型期長度（年）']}"
frow("用途（不含新債利息）", "US$bn",
     lambda i: (f"={COLS[i]}{FR['① CapEx（用途用：«YTD» 現金／«STUB» 毛額）']}+{COLS[i]}{FR['　租金合計']}+"
                f"{COLS[i]}{FR['③ 存量債務利息（備忘，«YTD» 已含在 CFO）']}+{COLS[i]}{FR['④ JV／策略投資出資']}+"
                f"{COLS[i]}{FR['　電力（overlay）']}+{COLS[i]}{FR['　維護（overlay）']}+{COLS[i]}{debt_row}+"
                f"{COLS[i]}{FR['⑥ capped call 支出（«YTD» 實際）']}"), NUM, BLACK, "避免循環參照：直接由各支出列加總")
wu_row = FR["用途（不含新債利息）"]
frow("新債期初餘額", "US$bn", lambda i: "=0", NUM, BLACK)
dn_beg = FR["新債期初餘額"]
frow("新債利率 × 期間長度", "%",
     lambda i: f"={COLS[i]}{od_rate}*'運營_產能與收入'!{COLS[i]}{CR['模型期長度（年）']}", '0.00%', BLACK)
rl_row = FR["新債利率 × 期間長度"]
frow("高息債期初餘額", "US$bn", lambda i: "=0", NUM, BLACK)
jn_beg = FR["高息債期初餘額"]
frow("高息債利率 × 期間長度", "%",
     lambda i: f"=({JRATE}+IF({CDSON}=1,MAX(0,{CDS}-{CDSB})/10000*{CDSP},0))*'運營_產能與收入'!{COLS[i]}{CR['模型期長度（年）']}", '0.00%', BLACK)
rj_row = FR["高息債利率 × 期間長度"]
frow("融資前現金（扣既有新債利息）", "US$bn",
     lambda i: (f"={COLS[i]}{beg_cash}+{COLS[i]}{srcop_row}+{COLS[i]}{atm_row}+{COLS[i]}{borrow_row}+{COLS[i]}{plug_row}"
                f"-{COLS[i]}{wu_row}-{COLS[i]}{rl_row}*{COLS[i]}{dn_beg}-{COLS[i]}{rj_row}*{COLS[i]}{jn_beg}"), NUM, BLACK)
pre_row = FR["融資前現金（扣既有新債利息）"]
frow("融資需求（補足至最低現金）", "US$bn", lambda i: f"=MAX(0,{MINC}-{COLS[i]}{pre_row})", NUM, BLACK, bold=True)
need_row = FR["融資需求（補足至最低現金）"]
frow("期初 backlog", "US$bn", lambda i: (f"={RPO0}+{RPOADD}" if i == 0 else "=0"), NUM, BLACK,
     "FY26 期初＝6/30 RPO 103.7＋Q3 新增 25")
bl_beg = FR["期初 backlog"]
frow("　新簽合約增量（年化增量×年期）", "US$bn",
     lambda i: (f"=MAX(0,'運營_產能與收入'!{COLS[i]}{CAP['newrev']}/'運營_產能與收入'!{COLS[i]}{CR['模型期長度（年）']}-0)*{TERM}" if i == 0 else
                f"=MAX(0,'運營_產能與收入'!{COLS[i]}{CAP['newrev']}/'運營_產能與收入'!{COLS[i]}{CR['模型期長度（年）']}-'運營_產能與收入'!{COLS[i-1]}{CAP['newrev']}/'運營_產能與收入'!{COLS[i-1]}{CR['模型期長度（年）']})*{TERM}"),
     NUM, BLACK)
bl_add = FR["　新簽合約增量（年化增量×年期）"]
frow("期末 backlog", "US$bn",
     lambda i: f"={COLS[i]}{bl_beg}-'運營_產能與收入'!{COLS[i]}{CAP['sch']}-'運營_產能與收入'!{COLS[i]}{CAP['newrev']}+{COLS[i]}{bl_add}",
     NUM, BLACK, "＝期初 − 排程認列 − 新簽約認列 ＋ 新簽合約增量")
bl_end = FR["期末 backlog"]
for i in range(1, 5):
    ws.cell(row=bl_beg, column=3 + i, value=f"={COLS[i-1]}{bl_end}")
frow("既有債務＋9/17 可轉債（期末）", "US$bn",
     lambda i: "=0", NUM, BLACK, "＝(依到期表遞減的既有本金，若關閉攤還則維持 35.551)＋9/17 可轉債 3.7")
ex_row = FR["既有債務＋9/17 可轉債（期末）"]
frow("債務上限（債務／backlog × 期末 backlog）", "US$bn", lambda i: f"={KBL}*{COLS[i]}{bl_end}", NUM, BLACK)
cap_row = FR["債務上限（債務／backlog × 期末 backlog）"]
frow("未動用額度（期初）", "US$bn", lambda i: (f"=IF({FACON}=1,{FAC},0)" if i == 0 else "=0"), NUM, BLACK)
fr_beg = FR["未動用額度（期初）"]
frow("新債可借上限", "US$bn",
     lambda i: f"=MAX(0,{COLS[i]}{cap_row}-({COLS[i]}{ex_row}+{COLS[i]}{dn_beg}),{COLS[i]}{fr_beg})", NUM, BLACK,
     "＝MAX(上限 − 既有債務 − 期初新債, 未動用額度)")
capd_row = FR["新債可借上限"]
frow("新債舉借", "US$bn",
     lambda i: f"=MIN({COLS[i]}{capd_row},{COLS[i]}{need_row}/(1-{COLS[i]}{rl_row}))", NUM, BLACK,
     "期前融資：期初到位，當期全額計息，所以需求要除以 (1−利率×期間)", bold=True)
nd_row = FR["新債舉借"]
frow("債務後剩餘需求", "US$bn",
     lambda i: f"=MAX(0,{COLS[i]}{need_row}+{COLS[i]}{rl_row}*{COLS[i]}{nd_row}-{COLS[i]}{nd_row})", NUM, BLACK)
rem_row = FR["債務後剩餘需求"]
frow("每年股權吸收上限", "US$bn",
     lambda i: f"=IF({EQCAP}>=9,1E+99,{EQCAP}*{EQPX}*{EQSH}*'運營_產能與收入'!{COLS[i]}{CR['模型期長度（年）']})", NUM, BLACK,
     f"＝現市值（發行參考價 × {_n(D['eqCapShares'])}bn 股）× 上限 % × 期間長度")
eqc_row = FR["每年股權吸收上限"]
frow("股權募資", "US$bn", lambda i: f"=MIN({COLS[i]}{rem_row},{COLS[i]}{eqc_row})", NUM, BLACK, bold=True)
eq_row = FR["股權募資"]
frow("高息債舉借（溢出）", "US$bn", lambda i: f"=({COLS[i]}{rem_row}-{COLS[i]}{eq_row})/(1-{COLS[i]}{rj_row})", NUM, BLACK, bold=True)
jd_row = FR["高息債舉借（溢出）"]
frow("新債＋高息債利息（期初餘額＋本期舉借）", "US$bn",
     lambda i: f"={COLS[i]}{rl_row}*({COLS[i]}{dn_beg}+{COLS[i]}{nd_row})+{COLS[i]}{rj_row}*({COLS[i]}{jn_beg}+{COLS[i]}{jd_row})", NUM, BLACK)
ni_row = FR["新債＋高息債利息（期初餘額＋本期舉借）"]
frow("發行價（現價×(1−折價)）", "US$", lambda i: f"={EQPX}*(1-{EQDISC})", USD, BLACK)
px_row = FR["發行價（現價×(1−折價)）"]
frow("新發行股數", "bn", lambda i: f"={COLS[i]}{eq_row}/{COLS[i]}{px_row}", '0.000', BLACK)
ns_row = FR["新發行股數"]
frow("累計新股", "bn", lambda i: (f"={COLS[i]}{ns_row}" if i == 0 else f"={COLS[i-1]}{{r}}+{COLS[i]}{ns_row}"), '0.000', BLACK, bold=True)
cns_row = FR["累計新股"]
for i in range(1, 5):
    ws.cell(row=cns_row, column=3 + i, value=f"={COLS[i-1]}{cns_row}+{COLS[i]}{ns_row}")
frow("新債期末餘額", "US$bn", lambda i: f"={COLS[i]}{dn_beg}+{COLS[i]}{nd_row}", NUM, BLACK)
dn_end = FR["新債期末餘額"]
frow("未動用額度（期末）", "US$bn", lambda i: f"={COLS[i]}{fr_beg}-MIN({COLS[i]}{fr_beg},{COLS[i]}{nd_row})", NUM, BLACK)
fr_end = FR["未動用額度（期末）"]
frow("高息債期末餘額", "US$bn", lambda i: f"={COLS[i]}{jn_beg}+{COLS[i]}{jd_row}", NUM, BLACK)
jn_end = FR["高息債期末餘額"]
frow("期末總債務（既有＋可轉債＋新債＋高息債）", "US$bn", lambda i: f"={COLS[i]}{ex_row}+{COLS[i]}{dn_end}+{COLS[i]}{jn_end}", NUM, BLACK, bold=True)
td_row = FR["期末總債務（既有＋可轉債＋新債＋高息債）"]
frow("融資前累積現金", "US$bn",
     lambda i: (f"={COLS[i]}{beg_cash}+{COLS[i]}{srcop_row}+{COLS[i]}{atm_row}+{COLS[i]}{borrow_row}+{COLS[i]}{plug_row}-{COLS[i]}{wu_row}" if i == 0
                else f"={COLS[i-1]}{{r}}+{COLS[i]}{srcop_row}+{COLS[i]}{atm_row}-{COLS[i]}{wu_row}"), NUM, BLACK,
     "若不做任何新融資的累積現金；負值＝外部資金需求", bold=True)
pfc_row = FR["融資前累積現金"]
for i in range(1, 5):
    ws.cell(row=pfc_row, column=3 + i, value=f"={COLS[i-1]}{pfc_row}+{COLS[i]}{srcop_row}+{COLS[i]}{atm_row}-{COLS[i]}{wu_row}")
for i in range(1, 5):
    ws.cell(row=dn_beg, column=3 + i, value=f"={COLS[i-1]}{dn_end}")
    ws.cell(row=jn_beg, column=3 + i, value=f"={COLS[i-1]}{jn_end}")
    ws.cell(row=fr_beg, column=3 + i, value=f"={COLS[i-1]}{fr_end}")
for i in range(5):
    ws.cell(row=newint_row, column=3 + i, value=f"={COLS[i]}{ni_row}")
    ws.cell(row=fac_row, column=3 + i, value=f"={COLS[i]}{nd_row}")
    ws.cell(row=eqr_row, column=3 + i, value=f"={COLS[i]}{eq_row}")
    ws.cell(row=jr_row, column=3 + i, value=f"={COLS[i]}{jd_row}")

r += 1
r = section(ws, r, "四層橋（驗算）", collapsed=True)
ws.cell(row=r, column=1, value="期初現金（2025-12-31）").font = BLACK
ws.cell(row=r, column=3, value=f"={H_CASH1231}").number_format = NUM
b1 = r; r += 1
ws.cell(row=r, column=1, value="＋ 營運缺口合計").font = BLACK
ws.cell(row=r, column=3, value=f"=SUM(C{opgap_row}:G{opgap_row})").number_format = NUM
b2 = r; r += 1
ws.cell(row=r, column=1, value="＋ 股權／可轉債合計").font = BLACK
ws.cell(row=r, column=3, value=f"=SUM(C{atm_row}:G{atm_row})").number_format = NUM
b3 = r; r += 1
ws.cell(row=r, column=1, value="＋ 1H 實際借款").font = BLACK
ws.cell(row=r, column=3, value=f"=SUM(C{borrow_row}:G{borrow_row})").number_format = NUM
b3b = r; r += 1
ws.cell(row=r, column=1, value="− capped call（1H）").font = BLACK
ws.cell(row=r, column=3, value=f"=-SUM(C{FR['⑥ capped call 支出（«YTD» 實際）']}:G{FR['⑥ capped call 支出（«YTD» 實際）']})").number_format = NUM
b3c = r; r += 1
ws.cell(row=r, column=1, value="＋ «YTD» 其他／受限現金調節").font = BLACK
ws.cell(row=r, column=3, value=f"=SUM(C{plug_row}:G{plug_row})").number_format = NUM
b3d = r; r += 1
ws.cell(row=r, column=1, value="＋ 瀑布新債＋股權＋高息債合計").font = BLACK
ws.cell(row=r, column=3, value=f"=SUM(C{fac_row}:G{fac_row})+SUM(C{eqr_row}:G{eqr_row})+SUM(C{jr_row}:G{jr_row})").number_format = NUM
b4 = r; r += 1
ws.cell(row=r, column=1, value="− 排程還本合計").font = BLACK
ws.cell(row=r, column=3, value=f"=-SUM(C{debt_row}:G{debt_row})").number_format = NUM
b5 = r; r += 1
ws.cell(row=r, column=1, value="＝ 期末現金（橋算）").font = BOLD
ws.cell(row=r, column=3, value=f"=SUM(C{b1}:C{b5})").number_format = NUM
ws.cell(row=r, column=3).font = BOLD
b6 = r; r += 1
ws.cell(row=r, column=1, value="對照：期末累積現金（表算）").font = BLACK
ws.cell(row=r, column=3, value=f"=G{cum_row}").number_format = NUM
b7 = r; r += 1
ws.cell(row=r, column=1, value="差異（應為 0）").font = BOLD
ws.cell(row=r, column=3, value=f"=C{b6}-C{b7}").number_format = NUM
b8 = r
ws.cell(row=r, column=9, value="若不為 0，表示公式鏈被改壞").font = SMALL

r += 2
r = section(ws, r, "FY26 全年備忘（對照公司指引）", collapsed=True)
memo = [
    ("FY26 總營收（1H 實際＋2H 模型）", f"='運營_產能與收入'!C{fy_rev_row}", NUM, "公司指引 12.4–13.2"),
    ("FY26 CapEx 認列（1H＋2H）", f"=C{FR['① 毛 CapEx（認列，備忘）']}", NUM, "公司指引 35–39"),
    ("«P0» 利息（«YTD» 實際＋«STUB» 模型）", f"={H_INT}+{inref('存量債務利息', 0)}+C{newint_row}", NUM,
     "Q3 指引 0.86–0.94／季"),
    ("FY26 租賃現金（1H 實際＋2H 模型）", f"={H_LEASE}+C{FR['　租金合計']}", NUM, "1H 實際支付 0.748"),
    ("FY26 排程還本（1H 實際＋2H）", f"=C{debt_row}", NUM, "1H 實際還款 5.219＋2H 到期表 4.413"),
]
memo0 = r
for nm, f, fmt, nt in memo:
    ws.cell(row=r, column=1, value=nm).font = BLACK
    c = ws.cell(row=r, column=3, value=f)
    c.number_format = fmt
    c.font = BLACK
    ws.cell(row=r, column=9, value=nt).font = SMALL
    r += 1

FRR = dict(FR)
FRR.update({"cum": cum_row, "opgap": opgap_row, "uses": uses_row, "int": int_row,
            "debt": debt_row, "atm": atm_row, "fac": fac_row, "srcop": srcop_row,
            "gap": gap_row, "bridge_diff": b8, "cns": cns_row, "td": td_row, "eq": eq_row, "nd": nd_row, "jd": jd_row, "pfc": pfc_row, "memo_rev": memo0, "memo_capex": memo0 + 1,
            "memo_int": memo0 + 2})

# =====================================================================
# 5. 租賃與承諾
# =====================================================================
ws = wb.create_sheet("資產負債_租賃承諾")
ws.column_dimensions["A"].width = 46
ws.column_dimensions["B"].width = 14
for c in ["C", "D", "E", "F", "G", "H"]:
    ws.column_dimensions[c].width = 13
ws.column_dimensions["I"].width = 86
ws["A1"] = "租賃與承諾 — 10-Q（«VD»）原始揭露"
ws["A1"].font = TITLE
ws["A2"] = "本頁為硬編碼的申報數字（藍字），支出與資金頁的租金與還本列由此而來"
ws["A2"].font = SMALL
r = 4
ws.cell(row=r, column=1, value="在帳租賃到期表（未折現）").font = BOLD
r += 1
heads = ["項目", "單位", "2026 下半", "2027", "2028", "2029", "2030", "其後"]
for j, h in enumerate(heads):
    c = ws.cell(row=r, column=1 + j, value=h)
    c.font = HEAD
    c.fill = FILL_HEAD
r += 1
lease_tbl = [
    ("營業租賃付款", "US$bn", CO['leases']['operatingPayments']),
    ("融資租賃付款", "US$bn", CO['leases']['financePayments']),
]
lr0 = r
for nm, un, vals in lease_tbl:
    ws.cell(row=r, column=1, value=nm).font = BLACK
    ws.cell(row=r, column=2, value=un).font = SMALL
    for i, v in enumerate(vals):
        c = ws.cell(row=r, column=3 + i, value=v)
        c.font = BLUE
        c.number_format = NUM
        c.border = BOX
    r += 1
ws.cell(row=r, column=1, value="合計（＝支出頁「在帳現金租金」）").font = BOLD
for i in range(6):
    col = get_column_letter(3 + i)
    c = ws.cell(row=r, column=3 + i, value=f"=SUM({col}{lr0}:{col}{lr0+1})")
    c.font = BOLD
    c.number_format = NUM
lease_tot = r
ws.cell(row=r, column=9, value="營業租賃負債現值 16.319、融資租賃 0.221；加權剩餘租期 12 年、折現率 10% [Verified]").font = SMALL
r += 2

ws.cell(row=r, column=1, value="債務本金到期表").font = BOLD
r += 1
for j, h in enumerate(heads):
    c = ws.cell(row=r, column=1 + j, value=h)
    c.font = HEAD
    c.fill = FILL_HEAD
r += 1
ws.cell(row=r, column=1, value="本金（＝支出頁「排程還本」）").font = BLACK
ws.cell(row=r, column=2, value="US$bn").font = SMALL
for i, v in enumerate(CO['debt']['amortization'] + [CO['debt']['amortAfterFY30']]):
    c = ws.cell(row=r, column=3 + i, value=v)
    c.font = BLUE
    c.number_format = NUM
    c.border = BOX
debt_tbl = r
ws.cell(row=r, column=9, value="本金合計 35.551（DDTL 13.6、票據 16.6、OEM 5.1 等）[Verified]").font = SMALL
r += 1
ws.cell(row=r, column=1, value="合計").font = BOLD
ws.cell(row=r, column=3, value=f"=SUM(C{debt_tbl}:H{debt_tbl})").number_format = NUM
ws.cell(row=r, column=3).font = BOLD
r += 2

ws.cell(row=r, column=1, value="表外與契約性承諾（未入表）").font = BOLD
r += 1
for j, h in enumerate(["項目", "單位", "金額", "時程", "入表？", "模型位置"]):
    c = ws.cell(row=r, column=1 + j, value=h)
    c.font = HEAD
    c.fill = FILL_HEAD
r += 1
commit = [
    ("未起租租賃（未折現）", "US$bn", 35.5, "2026–2029 起租，7–16 年", "否", "表外現金租金"),
    ("單站按造價計租，上限", "US$bn", 14.7, "393 MW 未交付，16 年", "否", "表外現金租金"),
    ("按造價計租，355 MW", "US$bn", None, "2026–2028 分期，金額未定", "否", "表外現金租金"),
    ("承租人自備設備承諾", "US$bn", 1.2, "0.5–1.2，分期至 2028", "否", "併入毛 CapEx"),
    ("JV 出資承諾", "US$bn", 1.7, "已付 0.55", "否", "JV／策略投資出資"),
    ("VIE 最大損失暴露", "US$bn", 0.108, "—", "否", "不另列"),
    ("可用信用額度（RCF＋DDTL 未動用）", "US$bn", 10.014, "DDTL 4.0 可動用至 2027-06", "否（來源）", "選用開關"),
]
c0 = r
for nm, un, amt, sched, onbs, pos in commit:
    ws.cell(row=r, column=1, value=nm).font = BLACK
    ws.cell(row=r, column=2, value=un).font = SMALL
    c = ws.cell(row=r, column=3, value=amt if amt is not None else "未定")
    c.font = BLUE
    if amt is not None:
        c.number_format = NUM
    ws.cell(row=r, column=4, value=sched).font = SMALL
    ws.cell(row=r, column=5, value=onbs).font = SMALL
    ws.cell(row=r, column=6, value=pos).font = SMALL
    r += 1
ws.cell(row=r, column=1, value="已揭露金額之承諾合計（未起租＋單站上限）").font = BOLD
ws.cell(row=r, column=3, value=f"=C{c0}+C{c0+1}").number_format = NUM
ws.cell(row=r, column=3).font = BOLD
commit_tot = r
r += 1
ws.cell(row=r, column=1, value="模型表外租金五期合計").font = BLACK
ws.cell(row=r, column=3, value=f"=SUM('各期收支'!C{FR['② 表外現金租金（未起租）']}:G{FR['② 表外現金租金（未起租）']})").number_format = NUM
ws.cell(row=r, column=3).font = GREEN
off5 = r
r += 1
ws.cell(row=r, column=1, value="模型表外租金尾端（FY30×12 年）").font = BLACK
ws.cell(row=r, column=3, value=f"='各期收支'!G{FR['② 表外現金租金（未起租）']}*12").number_format = NUM
off_tail = r
r += 1
ws.cell(row=r, column=1, value="模型路徑合計 ÷ 已承諾（≥0.8 為合理）").font = BOLD
ws.cell(row=r, column=3, value=f"=(C{off5}+C{off_tail})/C{commit_tot}").number_format = MULT
ws.cell(row=r, column=3).font = BOLD
ws.cell(row=r, column=9, value="8 GW 路徑所需新租約多數尚未簽署，路徑高於已承諾屬假設而非錯誤").font = SMALL
lease_ratio_row = r

# =====================================================================
# 5b. 債務明細
# =====================================================================
ws = wb.create_sheet("資產負債_既有債務")
ws.column_dimensions["A"].width = 38
ws.column_dimensions["B"].width = 12
for c in ["C", "D", "E", "F", "G", "H"]:
    ws.column_dimensions[c].width = 13
ws.column_dimensions["I"].width = 90
ws["A1"] = "債務明細 — 10-Q Note 10（«VD»）逐筆＋存量利息推導"
ws["A1"].font = TITLE
ws["A2"] = "有效利率為 10-Q 揭露值（含折價攤銷，四捨五入至整數 %）。本頁算出的存量利息連回『輸入』頁，再進入支出與資金頁"
ws["A2"].font = SMALL
r = 4
for j, h in enumerate(["工具", "類別", "到期", "有效利率", "«VMD» 本金", "年化利息", "", "", "備註"]):
    if h:
        c = ws.cell(row=r, column=1 + j, value=h); c.font = HEAD; c.fill = FILL_HEAD
r += 1
debts = [
    ("DDTL 1.0", "追索", "2028-03", 0.15, 1.300, "資產層、隨客戶付款攤還"),
    ("DDTL 2.0", "追索", "2030-08", 0.11, 3.190, ""),
    ("DDTL 2.1", "追索", "2031-03", 0.09, 3.000, ""),
    ("DDTL 3.0", "追索", "2030-08", 0.09, 2.215, ""),
    ("DDTL 5.0", "追索", "2031-11", 0.09, 1.101, "SOFR+4.50%；額度 3.1"),
    ("2030 Senior Notes", "追索", "2030-06", 0.10, 2.000, "到期一次還本"),
    ("2031 9.00% Senior Notes", "追索", "2031-02", 0.10, 1.750, ""),
    ("2031 9.75% Senior Notes", "追索", "2031-10", 0.10, 2.750, "2026-04 發行"),
    ("2032 9.625% Senior Notes", "追索", "2032-07", 0.10, 1.250, "2026-06 發行"),
    ("2032 EUR Senior Notes（€2.0bn）", "追索", "2032-07", 0.09, 2.279, "票面 8.50%；已做換匯換利避險"),
    ("2031 可轉債", "追索", "2031-12", 0.02, 2.588, ""),
    ("2032 可轉債", "追索", "2032-10", 0.02, 4.000, "票面 1.75%；2026-04 發行"),
    ("OEM 與軟體授權融資", "追索", "2026-12～2030-07", 0.11, 4.220, "短期部分加權 9%"),
    ("Magnetar Loan", "追索", "2029-01", 0.12, 0.189, ""),
    ("DDTL 4.0", "非追索", "2032-03", 0.07, 2.837, "SOFR+2.25%／Treasury+2.00%；額度 8.5，可動用至 2027-06"),
    ("OEM 融資（非追索）", "非追索", "2026-08～2028-08", 0.09, 0.882, ""),
]
d0 = r
for nm, cat, mat, rate, amt, note in debts:
    ws.cell(row=r, column=1, value=nm).font = BLACK
    ws.cell(row=r, column=2, value=cat).font = SMALL
    ws.cell(row=r, column=3, value=mat).font = SMALL
    c = ws.cell(row=r, column=4, value=rate); c.font = BLUE; c.number_format = PCT; c.border = BOX
    c = ws.cell(row=r, column=5, value=amt); c.font = BLUE; c.number_format = NUM; c.border = BOX
    c = ws.cell(row=r, column=6, value=f"=D{r}*E{r}"); c.font = BLACK; c.number_format = NUM; c.border = BOX
    ws.cell(row=r, column=9, value=note + (" [Verified]" if note else "[Verified]")).font = SMALL
    r += 1
d1 = r - 1
ws.cell(row=r, column=1, value="合計（«VMD»）").font = BOLD
c = ws.cell(row=r, column=5, value=f"=SUM(E{d0}:E{d1})"); c.font = BOLD; c.number_format = NUM
c = ws.cell(row=r, column=6, value=f"=SUM(F{d0}:F{d1})"); c.font = BOLD; c.number_format = NUM
ws.cell(row=r, column=9, value="應等於 10-Q 本金合計 35.551").font = SMALL
tot_r = r; r += 1
ws.cell(row=r, column=1, value="加權有效利率").font = BOLD
c = ws.cell(row=r, column=4, value=f"=F{tot_r}/E{tot_r}"); c.font = BOLD; c.number_format = PCT; c.fill = FILL_KEY
ws.cell(row=r, column=9, value="公司 9/17 簡報：Q2 加權 8.3%、含新可轉債 pro forma 7.8%（該口徑用票面，本頁用有效利率）").font = SMALL
WRATE = f"'資產負債_既有債務'!$D${r}"; r += 2

ws.cell(row=r, column=1, value="9/17 期後：2033 可轉債").font = BLACK
c = ws.cell(row=r, column=4, value=0.02875); c.font = BLUE; c.number_format = '0.000%'; c.border = BOX
c = ws.cell(row=r, column=5, value=3.7); c.font = BLUE; c.number_format = NUM; c.border = BOX
c = ws.cell(row=r, column=6, value=f"=D{r}*E{r}"); c.font = BLACK; c.number_format = NUM
ws.cell(row=r, column=9, value="轉換價 $97.85；2033 到期，不在五期還本表內 [Verified]").font = SMALL
CONV_INT = f"'資產負債_既有債務'!$F${r}"; r += 2

r = section(ws, r, "存量利息推導（本金依 10-Q 到期表遞減）")
r = period_header(ws, r)
DR = {}
def drow(name, unit, fn, fmt=NUM, font=BLACK, note=None, bold=False):
    global r
    ws.cell(row=r, column=1, value=name).font = BOLD if bold else BLACK
    ws.cell(row=r, column=2, value=unit).font = SMALL
    for i in range(5):
        c = ws.cell(row=r, column=3 + i, value=fn(i))
        c.font = Font(name="Arial", size=10, bold=True) if bold else font
        c.number_format = fmt; c.border = BOX
    if note:
        ws.cell(row=r, column=9, value=note).font = SMALL
    DR[name] = r; r += 1
drow("期初本金", "US$bn", lambda i: f"=E{tot_r}" if i == 0 else "=0", NUM, BLACK, "«P0» 期初為 «VMD» 本金（模型期起點）")
drow("排程還本", "US$bn", lambda i: f"={inref('排程還本（10-Q 本金表）', i)}", NUM, GREEN, "10-Q 到期表")
drow("期末本金", "US$bn", lambda i: f"={COLS[i]}{DR['期初本金']}-{COLS[i]}{DR['排程還本']}", NUM, BLACK)
for i in range(1, 5):
    ws.cell(row=DR["期初本金"], column=3 + i, value=f"={COLS[i-1]}{DR['期末本金']}")
drow("平均本金", "US$bn", lambda i: f"=({COLS[i]}{DR['期初本金']}+{COLS[i]}{DR['期末本金']})/2", NUM, BLACK)
drow("存量債務利息", "US$bn",
     lambda i: f"={COLS[i]}{DR['平均本金']}*{WRATE}*{inref('模型期長度（年）', i)}", NUM, BLACK,
     "＝平均本金×加權有效利率×期間長度。限制：假設利率組合不變（實際上 15% 的 DDTL 1.0 先到期，後期利率會略降）")
drow("9/17 可轉債利息", "US$bn", lambda i: f"={CONV_INT}*{inref('模型期長度（年）', i)}", NUM, BLACK)
drow("FY26 新增借款利息校準", "US$bn", lambda i: (f"={_n(D['intCal'])}" if i == 0 else "=0"), NUM, BLUE,
     "Q3 指引 0.86–0.94／季含 Q3 新動用的 DDTL；存量推算約 0.73／季，差額×2 季≈0.35。FY27 起新增借款改由瀑布新債利息承接 [Derived]")
drow("存量利息合計（連回輸入頁）", "US$bn",
     lambda i: f"={COLS[i]}{DR['存量債務利息']}+{COLS[i]}{DR['9/17 可轉債利息']}+{COLS[i]}{DR['FY26 新增借款利息校準']}",
     NUM, BLACK, "v1.4 手動值為 1.8／2.7／2.3／2.0／1.8", bold=True)
DEBT_TOTAL_ROW = DR["存量利息合計（連回輸入頁）"]
_wsf = wb["各期收支"]
for i in range(5):
    _wsf.cell(row=ex_row, column=3 + i, value=f"=IF({DEBTON}=1,'資產負債_既有債務'!{COLS[i]}{DR['期末本金']},35.551)+3.7")
wsin = wb["輸入與假設"]
for i in range(5):
    wsin.cell(row=DEBT_INT_ROW, column=3 + i, value=f"='資產負債_既有債務'!{COLS[i]}{DEBT_TOTAL_ROW}")

# =====================================================================
# 5c. 站點租賃
# =====================================================================
ws = wb.create_sheet("運營_站點")
ws.column_dimensions["A"].width = 30
for c, w in zip("BCDEFGHIJK", [18, 10, 10, 12, 8, 12, 12, 18, 26, 60]):
    ws.column_dimensions[c].width = w
ws["A1"] = "站點租賃 — 房東揭露 vs 10-Q 總額（檢驗租金路徑用）"
ws["A1"].font = TITLE
ws["A2"] = "10-Q 不揭露逐站租金；下表為房東公告的合約值（Interested-party），用來檢驗『每 MW 年租金』與模型租金路徑是否一致"
ws["A2"].font = SMALL
r = 4
heads = ["站點", "房東", "契約 MW", "已起租 MW", "合約總值 $bn", "年期", "年租金 $bn", "每 MW 年租金 $m", "起租時程", "10-Q 對應", "來源"]
for j, h in enumerate(heads):
    c = ws.cell(row=r, column=1 + j, value=h); c.font = HEAD; c.fill = FILL_HEAD
r += 1
sites = [
    ("Helios（德州 Dickens）", "Galaxy (GLXY)", 526, 133, 15.0, 15,
     "Ph I 133 MW 2026Q2 起租；Ph II 260 MW 2027H1 起；Ph III 133 MW", "未交付 393 MW＝10-Q「單站上限 14.7／16 年」",
     "Galaxy 2026-07-06：526 MW、年均租金 >$1bn、15 年＋2×5 年延長 [Interested-party]"),
    ("Polaris Forge 1（北達科他 Ellendale）", "Applied Digital (APLD)", 400, 250, 11.0, 15,
     "100 MW 2025Q4；150 MW 2026 年中；150 MW 2027", "大致落在未起租 35.5 內",
     "APLD 2025-08-29／FQ2 法說：400 MW、約 $11bn、約 15 年 [Interested-party]"),
    ("Core Scientific 六站（Denton 等）", "Core Scientific (CORZ)", 590, 250, 10.2, 12,
     "2026 年約 250 MW；2027 年初全數 590 MW", "部分在帳、部分未起租；「按造價計租」355 MW 可能含部分",
     "CORZ 2025-02-26：約 590 MW、$10.2bn、12 年；年均約 $850m [Interested-party]"),
    ("Kenilworth（新澤西 JV）", "JV（CRWV 持股 35%）", None, 0, None, 15,
     "施工中；完工後起租", "10-Q：租金＝建造成本的一定比例，推定屬「按造價計租」355 MW",
     "10-Q Note 3：15 年租約、租金依建造成本計；MW 與金額未揭露 [Verified]"),
]
s0 = r
for nm, ll, mw, com, val, yrs, sched, map10q, src in sites:
    ws.cell(row=r, column=1, value=nm).font = BLACK
    ws.cell(row=r, column=2, value=ll).font = SMALL
    for col, v, fmt in [(3, mw, NUM0), (4, com, NUM0), (5, val, NUM1), (6, yrs, NUM0)]:
        c = ws.cell(row=r, column=col, value=v if v is not None else "未揭露")
        c.font = BLUE
        if v is not None:
            c.number_format = fmt
        c.border = BOX
    if val is not None:
        c = ws.cell(row=r, column=7, value=f"=E{r}/F{r}"); c.number_format = NUM; c.border = BOX
        c = ws.cell(row=r, column=8, value=f"=G{r}/C{r}*1000"); c.number_format = NUM; c.border = BOX
    ws.cell(row=r, column=9, value=sched).font = SMALL
    ws.cell(row=r, column=10, value=map10q).font = SMALL
    ws.cell(row=r, column=11, value=src).font = SMALL
    r += 1
s1 = r - 1
ws.cell(row=r, column=1, value="具名站點合計").font = BOLD
for col in (3, 4, 5, 7):
    L = get_column_letter(col)
    c = ws.cell(row=r, column=col, value=f"=SUM({L}{s0}:{L}{s1})"); c.font = BOLD
    c.number_format = NUM0 if col in (3, 4) else NUM
c = ws.cell(row=r, column=8, value=f"=G{r}/C{r}*1000"); c.font = BOLD; c.number_format = NUM; c.fill = FILL_KEY
ws.cell(row=r, column=11, value="加權每 MW 年租金＝具名站點的市場基準").font = SMALL
named_r = r
BENCH = f"'運營_站點'!$H${r}"
r += 2

r = section(ws, r, "對帳：具名站點 vs 10-Q 租賃承諾總額", span=11)
rec = [
    ("10-Q 在帳租賃（未折現，已起租）", 29.37, "營業 29.135＋融資 0.235 [Verified]"),
    ("10-Q 未起租租賃（未折現）", 35.5, "2026–2029 起租、7–16 年 [Verified]"),
    ("10-Q 單站按造價計租上限", 14.7, "即 Helios 未交付 393 MW [Verified]"),
]
rc0 = r
for nm, v, nt in rec:
    ws.cell(row=r, column=1, value=nm).font = BLACK
    c = ws.cell(row=r, column=5, value=v); c.font = BLUE; c.number_format = NUM; c.border = BOX
    ws.cell(row=r, column=11, value=nt).font = SMALL
    r += 1
ws.cell(row=r, column=1, value="10-Q 已揭露租賃承諾合計").font = BOLD
c = ws.cell(row=r, column=5, value=f"=SUM(E{rc0}:E{r-1})"); c.font = BOLD; c.number_format = NUM
ws.cell(row=r, column=11, value="另有 355 MW 按造價計租，金額未定，未含").font = SMALL
tot10q = r; r += 1
ws.cell(row=r, column=1, value="減：具名站點合約總值").font = BLACK
c = ws.cell(row=r, column=5, value=f"=-E{named_r}"); c.number_format = NUM
r += 1
ws.cell(row=r, column=1, value="＝ 未具名站點（殘差）").font = BOLD
c = ws.cell(row=r, column=5, value=f"=E{tot10q}+E{r-1}"); c.font = BOLD; c.number_format = NUM
ws.cell(row=r, column=11, value="43 座以上資料中心中未具名者（Equinix、Switch 等託管，以及其他房東）。殘差大是必然——公司不揭露逐站").font = SMALL
r += 1
ws.cell(row=r, column=1, value="殘差占 10-Q 總額").font = BLACK
c = ws.cell(row=r, column=5, value=f"=E{r-1}/E{tot10q}"); c.number_format = PCT
r += 2

r = section(ws, r, "檢驗：模型租金路徑 ÷ 在役 MW vs 市場基準", span=11)
for j, h in enumerate(["項目", "單位"] + PERIODS):
    c = ws.cell(row=r, column=1 + j, value=h); c.font = HEAD; c.fill = FILL_HEAD
r += 1
chk0 = r
def srow(name, unit, fn, fmt=NUM, note=None, bold=False):
    global r
    ws.cell(row=r, column=1, value=name).font = BOLD if bold else BLACK
    ws.cell(row=r, column=2, value=unit).font = SMALL
    for i in range(5):
        c = ws.cell(row=r, column=3 + i, value=fn(i)); c.number_format = fmt; c.border = BOX
        c.font = Font(name="Arial", size=10, bold=True) if bold else BLACK
    if note:
        ws.cell(row=r, column=11, value=note).font = SMALL
    r += 1
    return r - 1
rr_rent = srow("模型租金（在帳＋表外）", "US$bn", lambda i: f"='各期收支'!{COLS[i]}{FR['　租金合計']}")
rr_mw = srow("平均在役 MW（Accepted）", "MW",
             lambda i: (f"=('輸入與假設'!{COLS[i]}${IN['期初主動電力']}+'輸入與假設'!{COLS[i]}${IN['Accepted MW（期末主動電力）']})/2"), NUM0)
rr_len = srow("模型期長度（年）", "年", lambda i: f"={inref('模型期長度（年）', i)}")
rr_per = srow("模型每 MW 年租金", "US$m/MW", lambda i: f"={COLS[i]}{rr_rent}/{COLS[i]}{rr_len}/{COLS[i]}{rr_mw}*1000", NUM,
              "FY26 欄只含下半年租金、MW 取全年平均，略有低估", bold=True)
ws.cell(row=r, column=1, value="第三方租賃占比").font = BLACK
c = ws.cell(row=r, column=3, value=0.85); c.font = BLUE; c.number_format = PCT; c.border = BOX
ws.cell(row=r, column=11, value="其餘為自建／JV（Kenilworth、Lancaster），成本走 CapEx 或 JV 出資 [Assumed]").font = SMALL
SHARE = f"$C${r}"; r += 1
rr_bench = srow("基準租金（MW×市場基準×占比）", "US$bn",
                lambda i: f"={COLS[i]}{rr_mw}*{BENCH}/1000*{SHARE}*{COLS[i]}{rr_len}", NUM,
                "若 8 GW 多數比照 Helios／APLD／CORZ 的租金水準")
rr_gap = srow("差額（基準 − 模型）", "US$bn", lambda i: f"={COLS[i]}{rr_bench}-{COLS[i]}{rr_rent}", NUM,
              "正值＝模型租金可能低估的金額", bold=True)
ws.cell(row=r, column=1, value="五期差額合計").font = BOLD
c = ws.cell(row=r, column=3, value=f"=SUM(C{rr_gap}:G{rr_gap})"); c.font = BOLD; c.number_format = NUM; c.fill = FILL_KEY
LEASE_GAP = f"'運營_站點'!$C${r}"
RENT_PER = rr_per
r += 1

# =====================================================================
# 6. 損益與評價
# =====================================================================
ws = wb.create_sheet("損益")
ws.column_dimensions["A"].width = 42
ws.column_dimensions["B"].width = 12
for c in COLS:
    ws.column_dimensions[c].width = 13
ws.column_dimensions["I"].width = 92
ws["A1"] = "損益（類損益表）— 與資金模型同一組收入"
ws["A1"].font = TITLE
ws["A2"] = "算力收入、Cash CapEx、利息皆由前面分頁連結而來；改資金假設，此頁會跟著動"
ws["A2"].font = SMALL
r = 4
r = period_header(ws, r)
VR = {}


def vrow(name, unit, fn, fmt=NUM, font=BLACK, note=None, bold=False):
    global r
    ws.cell(row=r, column=1, value=name).font = BOLD if bold else BLACK
    ws.cell(row=r, column=2, value=unit).font = SMALL
    for i in range(5):
        c = ws.cell(row=r, column=3 + i, value=fn(i))
        c.font = Font(name="Arial", size=10, bold=True) if bold else font
        c.number_format = fmt
        c.border = BOX
    if note:
        ws.cell(row=r, column=9, value=note).font = SMALL
    VR[name] = r
    r += 1


r = section(ws, r, "損益表")
vrow("算力收入", "US$bn", lambda i: f"='運營_產能與收入'!{COLS[i]}{CAP['isrev']}", NUM, GREEN, "＝RPO 轉換＋新簽約")
vrow("非算力服務", "US$bn", lambda i: f"={inref('非算力服務營收', i)}", NUM, GREEN)
vrow("總營收", "US$bn", lambda i: f"={COLS[i]}{VR['算力收入']}+{COLS[i]}{VR['非算力服務']}", NUM, BLACK, bold=True)
rev_v = VR["總營收"]
vrow("EBITDA 率", "%", lambda i: f"={inref('EBITDA 率', i)}", PCT, GREEN)
vrow("營業利益（EBIT）", "US$bn", lambda i: f"={COLS[i]}{rev_v}*{inref('EBITDA 率', i)}-'輸入與假設'!{COLS[i]}${IN['D&A（車隊）']}", NUM, BLACK,
     "＝營收×EBITDA 率 − 車隊 D&A")
ebit_v = VR["營業利益（EBIT）"]
vrow("利息（含瀑布新債）", "US$bn", lambda i: f"='各期收支'!{COLS[i]}{FRR['int']}", NUM, GREEN)
int_v = VR["利息（含瀑布新債）"]
vrow("稅前損益", "US$bn", lambda i: f"={COLS[i]}{ebit_v}-{COLS[i]}{int_v}", NUM, BLACK)
pre_v = VR["稅前損益"]
vrow("NOL 期初餘額", "US$bn", lambda i: f"={NOL0}", NUM, BLACK, f"«VMD» 累積虧損約 {V['nol']:.1f}")
nol_b = VR["NOL 期初餘額"]
vrow("　本期動用 NOL", "US$bn",
     lambda i: f"=MIN({COLS[i]}{nol_b},MAX(0,{COLS[i]}{pre_v})*{NOLU})", NUM, BLACK, f"美國規定：抵扣上限為應稅所得 {_n(V['nolUsePct'] * 100)}%")
nol_u = VR["　本期動用 NOL"]
vrow("NOL 期末餘額", "US$bn",
     lambda i: f"={COLS[i]}{nol_b}-{COLS[i]}{nol_u}+MAX(0,-{COLS[i]}{pre_v})", NUM, BLACK)
nol_e = VR["NOL 期末餘額"]
for i in range(5):
    ws.cell(row=nol_b, column=3 + i, value=(f"={NOL0}" if i == 0 else f"={COLS[i-1]}{nol_e}"))
vrow("所得稅", "US$bn",
     lambda i: f"=MAX(0,{COLS[i]}{pre_v}-{COLS[i]}{nol_u})*{TAX}", NUM, BLACK,
     "虧損年不認列稅盾（與公司評價備抵一致）")
tax_v = VR["所得稅"]
vrow("淨利", "US$bn", lambda i: f"={COLS[i]}{pre_v}-{COLS[i]}{tax_v}", NUM, BLACK, bold=True)
ni_v = VR["淨利"]
vrow("D&A", "US$bn", lambda i: f"='輸入與假設'!{COLS[i]}${IN['D&A（車隊）']}", NUM, GREEN, "車隊折舊，見輸入頁")
da_v = VR["D&A"]
vrow("EBITDA", "US$bn", lambda i: f"={COLS[i]}{ebit_v}+{COLS[i]}{da_v}", NUM, BLACK, bold=True)
ebitda_v = VR["EBITDA"]
vrow("　基礎股數（每年 +SBC 稀釋）", "bn", lambda i: (f"={SH}" if i == 0 else f"={COLS[i-1]}{r}*(1+{SHG})"), '0.0000', BLACK,
     "FY26＝含 ATM 上限 0.587；之後每年 ×(1＋股數年增)")
shb_v = VR["　基礎股數（每年 +SBC 稀釋）"]
vrow("股數（含瀑布新股）", "bn", lambda i: f"={COLS[i]}{shb_v}+'各期收支'!{COLS[i]}{FRR['cns']}", '0.0000', BLACK,
     "＝基礎股數＋累計瀑布新股")
shr_v = VR["股數（含瀑布新股）"]
vrow("每股盈餘（EPS，模型期）", "US$", lambda i: f"={COLS[i]}{ni_v}/{COLS[i]}{shr_v}", USD, BLACK,
     "«P0» 欄只含«STUBW»；全年口徑見下方")

r += 1
r = section(ws, r, "損益表（全年口徑；«FY0»；與 HTML 損益簡表一致）")
vrow("«YTDL» 實際營收（已實現）", "US$bn", lambda i: (f"={H_REV}" if i == 0 else "=0"), NUM, GREEN)
fy_h1 = VR["«YTDL» 實際營收（已實現）"]
vrow("全年總營收", "US$bn", lambda i: f"={COLS[i]}{fy_h1}+{COLS[i]}{rev_v}", NUM, BLACK, "FY26 可對照指引 12.4–13.2", bold=True)
fy_rev = VR["全年總營收"]
vrow("營收 YoY", "%", lambda i: (f"={COLS[i]}{fy_rev}/{_n(PREV_FY_REV)}-1" if i == 0 else f"={COLS[i]}{fy_rev}/{COLS[i-1]}{fy_rev}-1"), PCT, BLACK, f"{PREV_FY} 實際 {_n(PREV_FY_REV)}")
vrow("全年 GAAP 營業利益", "US$bn", lambda i: (f"={H_OPINC}+{COLS[i]}{ebit_v}" if i == 0 else f"={COLS[i]}{ebit_v}"), NUM, BLACK)
fy_ebit = VR["全年 GAAP 營業利益"]
vrow("營利率", "%", lambda i: f"={COLS[i]}{fy_ebit}/{COLS[i]}{fy_rev}", PCT, BLACK)
vrow("全年 D&A", "US$bn", lambda i: (f"={H_DA}+{COLS[i]}{da_v}" if i == 0 else f"={COLS[i]}{da_v}"), NUM, BLACK, f"«YTD» 實際 {YA['da']:.3f}")
fy_da = VR["全年 D&A"]
vrow("全年 EBITDA", "US$bn", lambda i: f"={COLS[i]}{fy_ebit}+{COLS[i]}{fy_da}", NUM, BLACK, bold=True)
vrow("全年利息（含瀑布新債）", "US$bn", lambda i: (f"={H_INT}+{COLS[i]}{int_v}" if i == 0 else f"={COLS[i]}{int_v}"), NUM, BLACK,
     "FY27 起瀑布新債利息（9%）與新股同時影響 EPS")
vrow("全年淨利", "US$bn", lambda i: (f"={H_NI}+{COLS[i]}{ni_v}" if i == 0 else f"={COLS[i]}{ni_v}"), NUM, BLACK, bold=True)
fy_ni = VR["全年淨利"]
vrow("淨利率", "%", lambda i: f"={COLS[i]}{fy_ni}/{COLS[i]}{fy_rev}", PCT, BLACK)
vrow("全年 GAAP EPS", "US$", lambda i: (f"={H_EPS}+{COLS[i]}{ni_v}/{COLS[i]}{shr_v}" if i == 0 else f"={COLS[i]}{ni_v}/{COLS[i]}{shr_v}"),
     USD, BLACK, f"«P0»＝«YTD» 實際 EPS {CO['ytdActual']['eps']:.2f}＋«STUBW»淨利÷股數".replace("-", "−"), bold=True)
vrow("全年 EPS（加回 SBC）", "US$",
     lambda i: (f"={H_NGEPS}+({COLS[i]}{ni_v}+{SBCY}*{inref('模型期長度（年）', i)})/{COLS[i]}{shr_v}" if i == 0
                else f"=({COLS[i]}{ni_v}+{SBCY}*{inref('模型期長度（年）', i)})/{COLS[i]}{shr_v}"),
     USD, BLACK, "＝(淨利＋SBC)÷股數；NOL 狀態無現金稅，SBC 全額加回")
r += 1
PL = "'損益'!"
ws = wb.create_sheet("評價_DCF與目標價")
ws.column_dimensions["A"].width = 42
ws.column_dimensions["B"].width = 12
for c in COLS:
    ws.column_dimensions[c].width = 13
ws.column_dimensions["I"].width = 92
ws["A1"] = "評價 — DCF、EV/EBITDA、加權目標價"
ws["A1"].font = TITLE
ws["A2"] = "營收、營業利益、D&A、EBITDA、股數引用『損益』頁；Cash CapEx 與融資引用『各期收支』頁"
ws["A2"].font = SMALL
r = 4
r = period_header(ws, r)
r = section(ws, r, "DCF（無槓桿自由現金流；起點為 «VD»）")
vrow("現金 CapEx（«VD» 之後）", "US$bn",
     lambda i: f"={inref('毛 CapEx', i)}-'各期收支'!{COLS[i]}{FR['　客戶預付金額（抵減，«STUB» 起）']}", NUM, BLACK,
     f"DCF 起點為 «VMD»，故 «P0» 欄只含 «STUB» 的建置支出，不含 «YTD» 已發生的 {YA['cashCapex']}")
capex_v = VR["現金 CapEx（«VD» 之後）"]
vrow("比較基期營收", "US$bn",
     lambda i: (f"={_n(PREV_FY_REV)}" if i == 0 else (f"={_n(YA['revenue'])}+{PL}{COLS[0]}{rev_v}" if i == 1 else f"={PL}{COLS[i-1]}{rev_v}")),
     NUM, BLACK, f"{PREV_FY} 實際 {_n(PREV_FY_REV)}；«P1» 基期＝«YTD» 實際 {CO['ytdActual']['revenue']}＋«P0» 模型期營收")
base_v = VR["比較基期營收"]
vrow("營運資金變動", "US$bn",
     lambda i: f"={WCP}*MAX(0,{PL}{COLS[i]}{rev_v}-{COLS[i]}{base_v})", NUM, BLACK, f"營收增量的 {_n(V['wcPctOfRevGrowth'] * 100)}%")
wc_v = VR["營運資金變動"]
vrow("無槓桿 NOL 期初", "US$bn", lambda i: f"={NOL0}", NUM, BLACK,
     "DCF 用無槓桿稅（利息不可抵稅，否則稅盾被計兩次：一次在 UFCF、一次在 WACC）")
nolu_b = VR["無槓桿 NOL 期初"]
vrow("　本期動用（無槓桿）", "US$bn",
     lambda i: f"=MIN({COLS[i]}{nolu_b},MAX(0,{PL}{COLS[i]}{ebit_v})*{NOLU})", NUM, BLACK)
nolu_u = VR["　本期動用（無槓桿）"]
vrow("無槓桿 NOL 期末", "US$bn",
     lambda i: f"={COLS[i]}{nolu_b}-{COLS[i]}{nolu_u}+MAX(0,-{PL}{COLS[i]}{ebit_v})", NUM, BLACK)
nolu_e = VR["無槓桿 NOL 期末"]
for i in range(5):
    ws.cell(row=nolu_b, column=3 + i, value=(f"={NOL0}" if i == 0 else f"={COLS[i-1]}{nolu_e}"))
vrow("無槓桿所得稅（DCF 用）", "US$bn",
     lambda i: f"=MAX(0,{PL}{COLS[i]}{ebit_v}-{COLS[i]}{nolu_u})*{TAX}", NUM, BLACK,
     "與損益表的所得稅不同：此處不扣利息，數字較高，這是 DCF 的正確口徑")
utax_v = VR["無槓桿所得稅（DCF 用）"]
vrow("UFCF", "US$bn",
     lambda i: f"={PL}{COLS[i]}{ebit_v}-{COLS[i]}{utax_v}+{PL}{COLS[i]}{da_v}-{COLS[i]}{capex_v}-{COLS[i]}{wc_v}",
     NUM, BLACK, "＝EBIT−無槓桿稅＋D&A−現金 CapEx−營運資金。建置期全為負是正常的", bold=True)
ufcf_v = VR["UFCF"]
vrow("折現期數", "年", lambda i: f"='輸入與假設'!${COLS[i]}${CAL_R['tEnd']}", '0.0', BLACK,
     f"自 «VD» 起算；«P0» 模型期（«STUBW»）的現金流平均落在 {CAL['tEnd'][0]} 年")
t_v = VR["折現期數"]
vrow("折現因子", "", lambda i: f"=1/(1+{WACC})^{COLS[i]}{t_v}", '0.000', BLACK)
df_v = VR["折現因子"]
vrow("UFCF 現值", "US$bn", lambda i: f"={COLS[i]}{ufcf_v}*{COLS[i]}{df_v}", NUM, BLACK)
pv_v = VR["UFCF 現值"]

r += 1
dcf_lines = [
    ("五期 UFCF 現值合計", f"=SUM(C{pv_v}:G{pv_v})", NUM, "建置期現金流現值"),
    ("常態化 FCF（FY30）", f"={PL}G{ebit_v}*(1-{TAX})+{PL}G{da_v}-{PL}G{da_v}*{MAINT}", NUM,
     "＝EBIT×(1−稅)＋D&A−維持性 CapEx(D&A×比率)。不讓成長性 CapEx 偽裝成永續 FCF"),
    ("終值（Gordon）", None, NUM, "＝常態化 FCF×(1+g)÷(WACC−g)"),
    ("終值現值", None, NUM, None),
    ("企業價值 EV", None, NUM, None),
    ("減：淨負債", f"=-{ND}", NUM, f"«VMD» 本金 {CO['latestQuarter']['debtPrincipal']} − 現金 {CO['latestQuarter']['cash']}"),
    ("股權價值", None, NUM, None),
    ("DCF 每股", None, USD, None),
    ("終值占 EV 比重", None, PCT, f'="超過 "&{_PC(RT_TVW)}&" 表示結論由終值假設決定，不由現金流決定"'),
    ("新股募得現金（現值）", None, NUM, "期初到位，折現期數 0／0.5／1.5／2.5／3.5 年"),
    ("融資後股數（含瀑布新股）", None, '0.000', None),
    ("DCF 失效？（WACC ≤ g 或常態化 FCF ≤ 0）", None, NUM0, "1＝失效：DCF 權重歸零、EV/EBITDA 100%。股權為負不算失效"),
    ("DCF 每股：0 截斷", None, USD, None),
    ("DCF 每股：選擇權（Merton）", None, USD, "Black-Scholes：S＝企業價值＋新股現值、K＝淨負債"),
    ("　d1", None, '0.000', None),
    ("　d2", None, '0.000', None),
]
d0 = r
for nm, f, fmt, nt in dcf_lines:
    ws.cell(row=r, column=1, value=nm).font = BOLD if "每股" in nm else BLACK
    if f:
        c = ws.cell(row=r, column=3, value=f)
        c.number_format = fmt
        c.font = BLACK
    if nt:
        ws.cell(row=r, column=9, value=nt).font = SMALL
    r += 1
ws.cell(row=d0 + 11, column=3, value=f"=IF(OR({WACC}<={GG},C{d0+1}<=0),1,0)").number_format = NUM0
ws.cell(row=d0 + 2, column=3, value=f"=IF(C{d0+11}=1,0,C{d0+1}*(1+{GG})/({WACC}-{GG}))").number_format = NUM
ws.cell(row=d0 + 3, column=3, value=f"=C{d0+2}*G{df_v}").number_format = NUM
ws.cell(row=d0 + 4, column=3, value=f"=C{d0}+C{d0+3}").number_format = NUM
ws.cell(row=d0 + 6, column=3, value=f"=C{d0+4}+C{d0+5}").number_format = NUM
ws.cell(row=d0 + 9, column=3, value="=" + "+".join(
    f"'各期收支'!{COLS[i]}{FRR['eq']}/(1+{WACC})^'輸入與假設'!${COLS[i]}${CAL_R['tStart']}" for i in range(5))).number_format = NUM
ws.cell(row=d0 + 10, column=3, value=f"={SH}+'各期收支'!G{FRR['cns']}").number_format = '0.000'
ws.cell(row=d0 + 12, column=3, value=f"=MAX(0,(C{d0+6}+C{d0+9})/C{d0+10})").number_format = USD
_S = f"(C{d0+4}+C{d0+9})"
ws.cell(row=d0 + 14, column=3, value=f"=IF({_S}<=0,0,(LN({_S}/{ND})+({RF}+{SIGMA}^2/2)*{OPTT})/({SIGMA}*SQRT({OPTT})))").number_format = '0.000'
ws.cell(row=d0 + 15, column=3, value=f"=C{d0+14}-{SIGMA}*SQRT({OPTT})").number_format = '0.000'
ws.cell(row=d0 + 13, column=3, value=(f"=IF({_S}<=0,0,({_S}*NORMSDIST(C{d0+14})-{ND}*EXP(-{RF}*{OPTT})*NORMSDIST(C{d0+15}))/C{d0+10})")).number_format = USD
ws.cell(row=d0 + 7, column=3, value=f"=IF(C{d0+11}=1,0,IF({DMODE}=2,C{d0+13},C{d0+12}))").number_format = USD
ws.cell(row=d0 + 7, column=9, value="採用值：失效時為 0 且權重歸零；否則依「DCF 下限方式」取 0 截斷或選擇權").font = SMALL
ws.cell(row=d0 + 7, column=3).font = BOLD
ws.cell(row=d0 + 8, column=3, value=f"=C{d0+3}/C{d0+4}").number_format = PCT
DCF_PS = f"C{d0+7}"
DCF_BAD = f"C{d0+11}"

r += 1
r = section(ws, r, "倍數法（錨定年度，預設 FY29；«EVDISCS»）")
m0 = r
_EB = f"{PL}$D${ebitda_v}:$G${ebitda_v}"
_TD = f"'各期收支'!$D${FRR['td']}:$G${FRR['td']}"
_CU = f"'各期收支'!$D${FRR['cum']}:$G${FRR['cum']}"
_SHR = f"{PL}$D${shr_v}:$G${shr_v}"
mult_lines = [
    ("錨定年度", f'=CHOOSE({EVY},"FY27","FY28","FY29","FY30")', "@", "輸入與假設 F 區可改"),
    ("錨定年度 EBITDA", f"=INDEX({_EB},1,{EVY})", NUM, None),
    ("錨定年度企業價值（倍數 × EBITDA）", None, NUM, None),
    ("錨定年度末淨負債（總債務 − 現金）", f"=INDEX({_TD},1,{EVY})-INDEX({_CU},1,{EVY})", NUM, "含瀑布新債"),
    ("錨定年度末股數（含瀑布新股）", f"=INDEX({_SHR},1,{EVY})", '0.000', "損益股數：SBC 逐年稀釋＋瀑布新股"),
    ("折回 «TGT» 的折現因子", f"=1/(1+{WACC})^({EVY}-{OFF})", '0.000', "以 WACC 折現 [Assumed]"),
    ("EV/EBITDA 每股（融資後）", None, USD, "以 0 為下限"),
]
for nm, f, fmt, nt in mult_lines:
    ws.cell(row=r, column=1, value=nm).font = BLACK
    if f:
        c = ws.cell(row=r, column=3, value=f)
        c.number_format = fmt
        c.font = BLACK
    if nt:
        ws.cell(row=r, column=9, value=nt).font = SMALL
    r += 1
ws.cell(row=m0 + 2, column=3, value=f"=C{m0+1}*{EVEBITDA}").number_format = NUM
ws.cell(row=m0 + 6, column=3, value=f"=MAX(0,(C{m0+2}-C{m0+3})/C{m0+4})*C{m0+5}").number_format = USD
EVE_ADJ = f"C{m0+6}"

r += 1
r = section(ws, r, "加權目標價")
# v4.5：DCF 腿以 WACC 推到目標價時點（0 截斷與選擇權皆同），與 EV/EBITDA 腿同一時點；「DCF 每股」仍為評價日現值（反向 DCF 用）
ws.cell(row=r, column=1, value="DCF 每股（推到目標價時點）").font = BLACK
c = ws.cell(row=r, column=3, value=f"={DCF_PS}*(1+{WACC})^{CAL_TT}"); c.number_format = USD; c.font = BLACK
ws.cell(row=r, column=9, value="＝DCF 每股 ×（1＋WACC）^（評價日至目標價時點的年數）").font = SMALL
DCF_T = f"C{r}"; r += 1
t0 = r
tgt = [
    ("DCF 每股 × 有效權重", f"={DCF_T}*IF({DCF_BAD}=1,0,{WDCF})", USD),
    ("EV/EBITDA（融資後）× 有效權重", f"={EVE_ADJ}*IF({DCF_BAD}=1,1,1-{WDCF})", USD),
    ("加權目標價", None, USD),
    ("現價", f"={PX}", USD),
    ("潛在空間", None, PCT),
]
for nm, f, fmt in tgt:
    ws.cell(row=r, column=1, value=nm).font = BOLD if nm == "加權目標價" else BLACK
    if f:
        c = ws.cell(row=r, column=3, value=f)
        c.number_format = fmt
        c.font = BLACK
    r += 1
ws.cell(row=t0 + 2, column=3, value=f"=SUM(C{t0}:C{t0+1})").number_format = USD
ws.cell(row=t0 + 2, column=3).font = BOLD
ws.cell(row=t0 + 2, column=3).fill = FILL_KEY
ws.cell(row=t0 + 4, column=3, value=f"=C{t0+2}/C{t0+3}-1").number_format = PCT
ws.cell(row=t0 + 2, column=9,
        value="期前融資瀑布：缺口先由額度與資產層新債支應，殘差以股權募足；估值不再扣缺口本金").font = SMALL
TGT = f"C{t0+2}"
UPS = f"C{t0+4}"
NOTE_R = r  # v4.1：DCF 權重說明（公式在目標價區間段建好後寫入）
ws.cell(row=r, column=1, value="DCF 權重說明").font = BLACK
ws.cell(row=r, column=9, value="只在 DCF 被 0 截斷時顯示；DCF 失效時改為一句；DCF 為正時空白（與 HTML 同一字串）").font = SMALL
r += 1

# v4.1：目標價區間（點位＝目前輸入的加權目標價，評等仍依點位）
# 情境區間：保守與積極情境的加權目標價，取自『輸入與假設』底部的模擬運算表（輸入格＝情境選擇，保留其餘手動輸入）
# 方法區間：目前輸入，EV/EBITDA 倍數換成『輸入與假設』F 區的兩端倍數
r += 1
r = section(ws, r, "目標價區間（情境區間、方法區間；評等仍依點位）")
IN_WS = wb["輸入與假設"]
DT_H = IN_WS.max_row + 3          # 模擬運算表的公式列（其上一列為段落標題）
DT_SC = ["low", "base", "high"]
DT_NM = {k: CO['scenarios']['labels'][k] for k in DT_SC}
DT_SHORT = {k: DT_NM[k].split(" ")[0] for k in DT_SC}
_IW = "'輸入與假設'!"
ws.cell(row=r, column=9, value=(f'="評等規則不變：空間 ≥"&{_PC(RT_BUY)}&" 且終值占比 <"&{_PC(RT_TV)}&"、股權募資 ≤"&{_MT(RT_EQ)}&" 倍現市值 → 買進；'
                                f'空間 ≤−"&{_PC(RT_SELL)}&"，或股權募資 >"&{_MT(RT_EQ)}&" 倍且空間 ≤−"&{_PC(RT_SELL2)}&" → 賣出；其餘中立"')).font = SMALL
R0 = r
def _code(px):  # 評等代碼：1＝買進、0＝中立、−1＝賣出（與 HTML rateCall 同規則；門檻引用『輸入與假設』F 區評等門檻）
    c = f"({px}/MAX({PX},0.01)-1)"
    return (f"IF(AND({c}>={RT_BUY},$C${R0+1}<{RT_TV},$C${R0+2}<={RT_EQ}),1,"
            f"IF(OR({c}<={RT_SELL},AND($C${R0+2}>{RT_EQ},{c}<={RT_SELL2})),-1,0))")
_CALL = lambda cell: f'CHOOSE({cell}+2,"賣出","中立","買進")'
_T1 = lambda x: f'TEXT({x},"0.0")'
_EVM = lambda m: (f"{DCF_T}*IF({DCF_BAD}=1,0,{WDCF})+MAX(0,(C{m0+1}*{m}-C{m0+3})/C{m0+4})*C{m0+5}*IF({DCF_BAD}=1,1,1-{WDCF})")
rows_rg = [
    ("目標價區間｜點位（目前輸入的加權目標價）", [f"={TGT}"], USD),
    ("目標價區間｜終值占 EV（評等用）", [f"=C{d0+3}/MAX(ABS(C{d0+4}),1)*IF(C{d0+4}=0,1,SIGN(C{d0+4}))"], PCT),
    ("目標價區間｜股權募資 ÷ 現市值（評等用）", [f"=SUM('各期收支'!C{FRR['eq']}:G{FRR['eq']})/({PX}*{SH})"], MULT),
    ("目標價區間｜點位評等代碼（1＝買進、0＝中立、−1＝賣出）", [f"={_code(TGT)}", f"={_CALL(f'C{R0+3}')}"], NUM0),
    (f"目標價區間｜賣出門檻價（現價 × (1 − {RT_PCT(RT['sellUpsideMax'])})）", [f"={PX}*(1+{RT_SELL})"], USD),
]
for j, k in enumerate(DT_SC):
    rows_rg.append((f"目標價區間｜{DT_NM[k]}（加權目標價／評等代碼／距賣出門檻）",
                    [f"={_IW}D{DT_H+1+j}", f"={_IW}E{DT_H+1+j}", f"=C{R0+5+j}-$C${R0+4}"], USD))
rows_rg += [
    ("目標價區間｜情境區間（下緣／上緣）", [f"=MIN(C{R0+5},C{R0+7})", f"=MAX(C{R0+5},C{R0+7})"], USD),
    ("目標價區間｜方法區間倍數（下端／上端）", [f"=MIN({RM_LO},{RM_HI})", f"=MAX({RM_LO},{RM_HI})"], MULT),
    ("目標價區間｜方法區間（下緣／上緣）", [f"=MIN({_EVM(f'C{R0+9}')},{_EVM(f'D{R0+9}')})", f"=MAX({_EVM(f'C{R0+9}')},{_EVM(f'D{R0+9}')})"], USD),
    ("目標價區間｜100% EV/EBITDA 目標價（價位／評等代碼）", [f"={EVE_ADJ}", f"={_code(EVE_ADJ)}", f"={_CALL(f'D{R0+11}')}"], USD),
    ("目標價區間｜未截斷 DCF 每股", [f"=(C{d0+6}+C{d0+9})/C{d0+10}"], USD),
]
for nm, fs, fmt in rows_rg:
    ws.cell(row=r, column=1, value=nm).font = BLACK
    for j, f in enumerate(fs):
        c = ws.cell(row=r, column=3 + j, value=f); c.number_format = fmt; c.font = BLACK
    r += 1
for rr in (R0 + 5, R0 + 6, R0 + 7):
    ws.cell(row=rr, column=4).number_format = NUM0
ws.cell(row=R0 + 11, column=4).number_format = NUM0
ws.cell(row=R0 + 3, column=3).number_format = NUM0
ws.cell(row=R0 + 5, column=9, value="引用『輸入與假設』底部模擬運算表（情境選擇 1／2／3），保留其餘手動輸入").font = SMALL
ws.cell(row=R0 + 10, column=9, value="＝DCF 採用值 × 有效權重 ＋ EV/EBITDA 腿（倍數換成兩端）× 有效權重").font = SMALL
_A0, _A1, _B0, _B1 = f"C{R0+8}", f"D{R0+8}", f"C{R0+10}", f"D{R0+10}"
_TH = f"$C${R0+4}"
_mg = lambda j, k: (f'IF(ABS(E{R0+5+j})<{RT_MG}*{PX},"{DT_SHORT[k]}情境 $"&{_T1(f"C{R0+5+j}")}&"，"&IF(E{R0+5+j}<0,"距賣出門檻","高於賣出門檻")'
                    f'&" $"&{_T1(_TH)}&" 僅 $"&{_T1(f"ABS(E{R0+5+j})")}&"。","")')
_ns = '&" "&'.join(f'IF(D{R0+5+j}<>-1,"{DT_SHORT[k]}","")' for j, k in enumerate(DT_SC))
_nsc = "+".join(f"(D{R0+5+j}<>-1)" for j in range(3))
text_rg = [
    ("目標價區間｜點位摘要", f'="目標價 $"&{_T1(TGT)}&"（情境區間 $"&{_T1(_A0)}&"–$"&{_T1(_A1)}&"）"'),
    ("目標價區間｜方法區間標籤", f'="方法區間（目前情境，"&{_MT(f"C{R0+9}")}&"–"&{_MT(f"D{R0+9}")}&"x）"'),
    ("目標價區間｜點位位置", (f'=IF(ABS({EVEBITDA}-D{R0+9})<1E-9,"目前點位位於方法區間上緣（"&{_MT(f"D{R0+9}")}&"x 為穩態倍數上緣）",'
                        f'IF(ABS({EVEBITDA}-C{R0+9})<1E-9,"目前點位位於方法區間下緣（"&{_MT(f"C{R0+9}")}&"x）",'
                        f'"目前點位"&IF(AND({EVEBITDA}>C{R0+9},{EVEBITDA}<D{R0+9}),"位於方法區間內",IF({EVEBITDA}>D{R0+9},"高於方法區間上緣","低於方法區間下緣"))'
                        f'&"（目前倍數 "&{_MT(EVEBITDA)}&"x）"))')),
    ("目標價區間｜判斷句", (f'=IF({_A1}<{PX},"情境區間上緣 $"&{_T1(_A1)}&" 低於現價 $"&TEXT({PX},"0.00")'
                       f'&IF({_nsc}>0,"，但"&SUBSTITUTE(TRIM({_ns})," ","、")&"情境未達賣出門檻 $"&{_T1(_TH)},"：賣出在三個擴張情境下都成立"),'
                       f'"情境區間上緣 $"&{_T1(_A1)}&" 不低於現價 $"&TEXT({PX},"0.00")&"：賣出並非在三個擴張情境下都成立")&"。"'
                       + "".join("&" + _mg(j, k) for j, k in enumerate(DT_SC)))),
]
TRROW = {}  # v4.3：文字列的列號（「摘要」頁引用）
for nm, f in text_rg:
    TRROW[nm] = r
    ws.cell(row=r, column=1, value=nm).font = BOLD if nm.endswith("判斷句") else BLACK
    ws.cell(row=r, column=3, value=f).font = BLACK
    r += 1
_RAW = f"C{R0+12}"
ws.cell(row=NOTE_R, column=3, value=(
    f'=IF({DCF_BAD}=1,"DCF 失效，已以 EV/EBITDA 100% 計。",IF(AND({DMODE}<>2,{_RAW}<0),'
    f'"DCF 為 $0 仍計 "&TEXT({WDCF},"0%")&" 權重的原因：DCF 沒有失效，而是算出的股權價值為負（每股約 −$"&{_T1(f"ABS({_RAW})")}&"），'
    f'代表自由現金流折現後的企業價值不足以償還淨負債。EV/EBITDA 沒有扣除 GPU 汰換與擴張支出，保留 DCF 權重，是為了讓目標價反映「EBITDA 付不起資本成本」。'
    f'若不計 DCF，目標價為 $"&{_T1(EVE_ADJ)}&"（"&{_CALL(f"D{R0+11}")}&"）；"&TEXT({WDCF},"0%")&"／"&TEXT(1-{WDCF},"0%")&" 為判斷值。",""))')).font = BLACK
RG_PT_CODE = f"C{R0+3}"

# v3.5：錨定年度 × 倍數矩陣（全為活公式；與 HTML evGrid 同一算法）
from openpyxl.formatting.rule import CellIsRule as _CIR
r += 1
r = section(ws, r, "錨定年度 × 倍數矩陣（每股；«EVDISC»）")
ws.cell(row=r, column=9, value="EV/EBITDA 腿＝(錨定年 EBITDA × 倍數 − 錨定年末淨負債) ÷ 錨定年末股數 × 折現因子；加權＝DCF 採用值 × 有效權重 ＋ 腿 × 有效權重。綠底＝不低於現價").font = SMALL
EVG_M = [3.4, 4.5, 5.0, 6.0, 7.0]
for kind in ("EV/EBITDA 腿", "加權目標價"):
    for j, h in enumerate(["倍數 ＼ 錨定（" + kind + "）", "倍數", "FY27", "FY28", "FY29", "FY30"]):
        c = ws.cell(row=r, column=1 + j, value=h); c.font = HEAD; c.fill = FILL_HEAD
    r += 1
    g0 = r
    for m in EVG_M:
        ws.cell(row=r, column=1, value=f"錨定×倍數｜{kind}｜{m:.1f}x").font = BLACK
        cb = ws.cell(row=r, column=2, value=m); cb.number_format = MULT; cb.font = BLUE
        for k in range(1, 5):
            col = 2 + k
            if kind == "EV/EBITDA 腿":
                f = (f"=MAX(0,(INDEX({_EB},1,{k})*$B{r}-(INDEX({_TD},1,{k})-INDEX({_CU},1,{k})))"
                     f"/INDEX({_SHR},1,{k}))/(1+{WACC})^({k}-{OFF})")
            else:
                leg_row = LEG0 + (r - g0)
                f = f"={DCF_T}*IF({DCF_BAD}=1,0,{WDCF})+{get_column_letter(col)}{leg_row}*IF({DCF_BAD}=1,1,1-{WDCF})"
            c = ws.cell(row=r, column=col, value=f); c.number_format = USD; c.font = BLACK
        r += 1
    rng = f"C{g0}:F{r-1}"
    ws.conditional_formatting.add(rng, _CIR(operator="greaterThanOrEqual", formula=[PX], fill=PatternFill("solid", fgColor="E3F1E8")))
    if kind == "EV/EBITDA 腿":
        LEG0 = g0
    r += 1

# =====================================================================
# 6b. 可比公司
# =====================================================================
ws = wb.create_sheet("評價_可比公司")
for c, w in zip("ABCDEFGHIJKL", [26, 10, 10, 10, 10, 13, 13, 11, 11, 18, 70, 4]):
    ws.column_dimensions[c].width = w
PEERS = CO['peers']  # 同業 Comps：資料只在 company.json → peers（與 HTML 共用）
ws["A1"] = f"可比公司 — 以 EV/Sales 為主（市值：{PEERS['priceDate']} 收盤）"
ws["A1"].font = TITLE
ws["A2"] = PEERS['textXlsx']['subtitle']
ws["A2"].font = SMALL
r = 4
heads = ["公司", "市值", "淨負債", "EV", "TTM 營收", "EV/Sales（TTM）", "含租賃 EV/Sales", "TTM EBITDA", "EV/EBITDA", "資料期", "備註"]
for j, h in enumerate(heads):
    c = ws.cell(row=r, column=1 + j, value=h); c.font = HEAD; c.fill = FILL_HEAD
ws.cell(row=r, column=6).fill = PatternFill("solid", fgColor="0F5C61")
r += 1
peers = [(e["labelXlsx"], e["mkt"], e["netDebt"], e["rev"], e["opl"], e["opInc"], e["da"], e["asOf"], e["noteXlsx"]) for e in PEERS["list"]]
p0 = r
for nm, mkt, nd, rev, opl, opi, da, asof, note in peers:
    ws.cell(row=r, column=1, value=nm).font = BLACK
    for col, v in [(2, mkt), (3, nd), (5, rev)]:
        c = ws.cell(row=r, column=col, value=v); c.font = BLUE; c.number_format = NUM; c.border = BOX
    c = ws.cell(row=r, column=4, value=f"=B{r}+C{r}"); c.number_format = NUM; c.border = BOX
    c = ws.cell(row=r, column=6, value=f"=D{r}/E{r}"); c.number_format = MULT; c.border = BOX; c.font = BOLD
    c = ws.cell(row=r, column=7, value=f"=(D{r}+{opl})/E{r}"); c.number_format = MULT; c.border = BOX
    c = ws.cell(row=r, column=8, value=f"={opi}+{da}"); c.font = BLUE; c.number_format = NUM; c.border = BOX
    c = ws.cell(row=r, column=9, value=f'=IF(H{r}>0,D{r}/H{r},"n/m")'); c.number_format = MULT; c.border = BOX
    ws.cell(row=r, column=10, value=asof).font = SMALL
    ws.cell(row=r, column=11, value=f"營業租賃負債 {opl}；GAAP 營業損益 {opi}＋D&A {da}。{note}").font = SMALL
    r += 1
ws.cell(row=r, column=1, value="同業中位數").font = BOLD
c = ws.cell(row=r, column=6, value=f"=MEDIAN(F{p0}:F{r-1})"); c.number_format = MULT; c.font = BOLD
c = ws.cell(row=r, column=7, value=f"=MEDIAN(G{p0}:G{r-1})"); c.number_format = MULT
med_r = r
r += 1
ws.cell(row=r, column=1, value="CRWV（TTM 至 Q2）").font = BOLD
c = ws.cell(row=r, column=2, value=CO['callFacts']['mktCapLast']); c.font = BLUE; c.number_format = NUM
c = ws.cell(row=r, column=3, value=f"={ND}"); c.font = GREEN; c.number_format = NUM
c = ws.cell(row=r, column=4, value=f"=B{r}+C{r}"); c.number_format = NUM
c = ws.cell(row=r, column=5, value=CO['callFacts']['ttmRev']); c.font = BLUE; c.number_format = NUM
c = ws.cell(row=r, column=6, value=f"=D{r}/E{r}"); c.number_format = MULT; c.font = BOLD
c = ws.cell(row=r, column=7, value=f"=(D{r}+{CO['latestQuarter']['opLeaseLiab']})/E{r}"); c.number_format = MULT
c = ws.cell(row=r, column=8, value=f"={CO['callFacts']['ttmOpInc']}+{CO['callFacts']['ttmDa']}"); c.font = BLUE; c.number_format = NUM
c = ws.cell(row=r, column=9, value=f'=IF(H{r}>0,D{r}/H{r},"n/m")'); c.number_format = MULT
ws.cell(row=r, column=10, value="TTM 至 Q2").font = SMALL
ws.cell(row=r, column=11, value=f"EV 用 9/21 市值與 6/30 淨負債；含租賃再加營業租賃負債 {CO['latestQuarter']['opLeaseLiab']}").font = SMALL
crwv_r = r
r += 1
ws.cell(row=r, column=1, value=f"{CO['meta']['ticker']} 模型 «P1»E").font = BLACK
c = ws.cell(row=r, column=2, value=f"={PX}*{SH}"); c.number_format = NUM
c = ws.cell(row=r, column=3, value=f"={ND}"); c.number_format = NUM
c = ws.cell(row=r, column=4, value=f"=B{r}+C{r}"); c.number_format = NUM
c = ws.cell(row=r, column=5, value=f"='損益'!D{rev_v}"); c.font = GREEN; c.number_format = NUM
c = ws.cell(row=r, column=6, value=f"=D{r}/E{r}"); c.number_format = MULT
c = ws.cell(row=r, column=7, value=f"=(D{r}+{CO['latestQuarter']['opLeaseLiab']})/E{r}"); c.number_format = MULT
c = ws.cell(row=r, column=8, value=f"='損益'!D{ebitda_v}"); c.font = GREEN; c.number_format = NUM
c = ws.cell(row=r, column=9, value=f"=D{r}/H{r}"); c.number_format = MULT
ws.cell(row=r, column=11, value="市值用模型股數（含 ATM 上限）；分母為模型 FY27 營收").font = SMALL
r += 2
notes = PEERS['textXlsx']['notes']
for t in notes:
    ws.cell(row=r, column=1, value=t).font = SMALL
    r += 1

# =====================================================================
# 6c. 資產負債_新債與新股（連結各期收支與損益；類資產負債表）
# =====================================================================
ws = wb.create_sheet("資產負債_新債與新股")
ws.column_dimensions["A"].width = 44
ws.column_dimensions["B"].width = 10
for c in COLS:
    ws.column_dimensions[c].width = 12
ws.column_dimensions["I"].width = 80
ws["A1"] = "新債與新股（類資產負債表：現金、債務、股本、槓桿）"
ws["A1"].font = TITLE
ws["A2"] = "全部為連結公式。«P0» 欄的現金自 «VD» 起算（模型期），與 HTML『資產與負債 → 新債與新股』同口徑"
ws["A2"].font = SMALL
r = 4
r = period_header(ws, r)
F_ = "'各期收支'!"; P_ = "'損益'!"; I_ = "'輸入與假設'!"; Dd = "'資產負債_既有債務'!"
BR = {}
def brow(name, unit, fn, fmt=NUM, note=None, bold=False):
    global r
    ws.cell(row=r, column=1, value=name).font = BOLD if bold else BLACK
    ws.cell(row=r, column=2, value=unit).font = SMALL
    for i in range(5):
        c = ws.cell(row=r, column=3 + i, value=fn(i)); c.number_format = fmt; c.border = BOX
        c.font = Font(name="Arial", size=10, bold=True) if bold else BLACK
    if note:
        ws.cell(row=r, column=9, value=note).font = SMALL
    BR[name] = r; r += 1
r = section(ws, r, "現金")
brow("期初現金", "US$bn", lambda i: (f"={CASH0}" if i == 0 else f"={F_}{COLS[i-1]}{cum_row}"), NUM, "«P0»＝«VMD» 現金（模型期起點）")
brow("營運收支淨額（含 9/17 可轉債／ATM，不含瀑布）", "US$bn",
     lambda i: (f"={F_}{COLS[i]}{pfc_row}-{CASH0}" if i == 0 else f"={F_}{COLS[i]}{pfc_row}-{F_}{COLS[i-1]}{pfc_row}"), NUM,
     "＝融資前累積現金的當期變動；負值即當期外部資金需求")
brow("＋ 瀑布：新債（額度＋資產層）", "US$bn", lambda i: f"={F_}{COLS[i]}{nd_row}")
brow("＋ 瀑布：高息債", "US$bn", lambda i: f"={F_}{COLS[i]}{jd_row}")
brow("＋ 瀑布：股權募資", "US$bn", lambda i: f"={F_}{COLS[i]}{eq_row}")
brow("− 新融資利息（新債＋高息債）", "US$bn", lambda i: f"=-{F_}{COLS[i]}{ni_row}")
brow("期末現金", "US$bn", lambda i: f"=SUM({COLS[i]}{BR['期初現金']}:{COLS[i]}{BR['− 新融資利息（新債＋高息債）']})", NUM,
     "應等於『各期收支』期末累積現金（見下列驗算）", bold=True)
brow("　驗算：與各期收支期末現金差異", "US$bn", lambda i: f"={COLS[i]}{BR['期末現金']}-{F_}{COLS[i]}{cum_row}", NUM, "應為 0")
r = section(ws, r, "債務")
brow("既有債務期初（10-Q 本金）", "US$bn", lambda i: (f"={_n(LQ_DEBT)}" if i == 0 else f"={COLS[i-1]}{{r}}"), NUM)
brow("− 排程還本（模型期）", "US$bn", lambda i: f"=-IF({DEBTON}=1,{I_}{COLS[i]}${IN['排程還本（10-Q 本金表）']},0)")
brow("既有債務期末", "US$bn", lambda i: f"={F_}{COLS[i]}{ex_row}-{_n(CONV_PR)}", NUM, f"＝依到期表遞減（關閉攤還時維持 {_n(LQ_DEBT)}）")
for i in range(1, 5):
    ws.cell(row=BR["既有債務期初（10-Q 本金）"], column=3 + i, value=f"={COLS[i-1]}{BR['既有債務期末']}")
brow("＋ 9/17 可轉債（2033 到期）", "US$bn", lambda i: f"={_n(CONV_PR)}")
brow("＋ 瀑布新債餘額", "US$bn", lambda i: f"={F_}{COLS[i]}{dn_end}")
brow("＋ 高息債餘額", "US$bn", lambda i: f"={F_}{COLS[i]}{jn_end}")
brow("總債務", "US$bn", lambda i: f"=SUM({COLS[i]}{BR['既有債務期末']}:{COLS[i]}{BR['＋ 高息債餘額']})", NUM, "應等於『各期收支』期末總債務", bold=True)
brow("淨負債（總債務 − 期末現金）", "US$bn", lambda i: f"={COLS[i]}{BR['總債務']}-{COLS[i]}{BR['期末現金']}", NUM, None, bold=True)
r = section(ws, r, "股本")
brow("基礎股數（含 ATM 上限、SBC 稀釋）", "bn", lambda i: f"={P_}{COLS[i]}{shb_v}", '0.000')
brow("＋ 本期新股", "bn", lambda i: f"={F_}{COLS[i]}{ns_row}", '0.000')
brow("總股數（期末）", "bn", lambda i: f"={P_}{COLS[i]}{shr_v}", '0.000', "＝基礎股數＋累計新股", bold=True)
r = section(ws, r, "槓桿")
brow("期末 backlog", "US$bn", lambda i: f"={F_}{COLS[i]}{bl_end}")
brow("債務上限（債務／backlog × 期末 backlog）", "US$bn", lambda i: f"={F_}{COLS[i]}{cap_row}")
brow("總債務 ÷ 期末 backlog", "x", lambda i: f"={COLS[i]}{BR['總債務']}/MAX(0.01,{COLS[i]}{BR['期末 backlog']})", '0.00x', "高息債不受 backlog 上限約束")
brow("總債務 ÷ EBITDA（年化）", "x", lambda i: f"={COLS[i]}{BR['總債務']}/MAX(0.01,{P_}{COLS[i]}{ebitda_v}/{I_}{COLS[i]}${IN['模型期長度（年）']})", '0.0x', "FY26 模型期 EBITDA 以半年×2 年化")
NB = dict(BR)

# =====================================================================
# 6d. 評價_反向DCF（目標搜尋＋HTML 快照）
# =====================================================================
ws = wb.create_sheet("評價_反向DCF")
ws.column_dimensions["A"].width = 46
for c in "BCDEF":
    ws.column_dimensions[c].width = 16
ws.column_dimensions["G"].width = 70
ws["A1"] = "反向 DCF — 現價隱含什麼樣的收入與成本"
ws["A1"].font = TITLE
ws["A2"] = "不是問「值多少」，而是問「要值現價，營運要長什麼樣」。Excel 以『目標搜尋』反解；下方另附 HTML 即時計算結果的快照（v4.0 預設輸入）"
ws["A2"].font = SMALL
VQ2 = "'評價_DCF與目標價'!"
r = 4
r = section(ws, r, "即時：目前 DCF 與現價的差距", span=7)
ws.cell(row=r, column=1, value="現價").font = BLACK
c = ws.cell(row=r, column=2, value=f"={PX}"); c.number_format = USD; c.font = GREEN
r += 1
ws.cell(row=r, column=1, value="目前 DCF 每股（依下限方式）").font = BLACK
c = ws.cell(row=r, column=2, value=f"={VQ2}{DCF_PS}"); c.number_format = USD; c.font = GREEN
r += 1
ws.cell(row=r, column=1, value="差距（DCF − 現價）＝目標搜尋的『設定儲存格』").font = BOLD
c = ws.cell(row=r, column=2, value=f"=B{r-1}-B{r-2}"); c.number_format = USD; c.font = BOLD; c.fill = FILL_KEY
GAPCELL = f"B{r}"
r += 1
ws.cell(row=r, column=1, value="目前每 MW 年收入倍數／建置成本倍數").font = BLACK
c = ws.cell(row=r, column=2, value=f"={REVSC}"); c.number_format = PCT; c.font = GREEN
c = ws.cell(row=r, column=3, value=f"={CAPSC}"); c.number_format = PCT; c.font = GREEN
r += 2
r = section(ws, r, "操作：用『目標搜尋』反解（資料 → 模擬分析 → 目標搜尋）", span=7)
for t in [
    f"① 設定儲存格：本頁 {GAPCELL}；目標值：0",
    "② 變數儲存格（擇一）：輸入與假設 B 區『每 MW 年收入倍數』→ 市價隱含單價；或 D 區『每 MW 建置成本倍數』→ 市價隱含建置成本；或 C 區『穩態 EBITDA 率』",
    "③ 解完請記下結果，再把倍數改回 100%（或按 Ctrl+Z）。DCF 以 0 截斷時，負值區域斜率為 0，請從 100% 以上開始搜尋",
    "④ 組合分析：先把 D 區建置成本倍數與 C 區穩態 EBITDA 率設為想測的值，再對『每 MW 年收入倍數』做目標搜尋，即得下方矩陣中的一格",
]:
    ws.cell(row=r, column=1, value=t).font = SMALL; r += 1
r += 1
r = section(ws, r, "快照：單一槓桿（HTML v4.0 預設輸入，DCF 0 截斷）", span=7)
for j, h in enumerate(["單一槓桿（其他不變）", "目前", "市價隱含", "變動", "", "", "對照"]):
    if h:
        c = ws.cell(row=r, column=1 + j, value=h); c.font = HEAD; c.fill = FILL_HEAD
r += 1
_rows = [
    ("FY30 每 MW 年收入（US$m）", rv["rev30"], rv["rev30"] * rv["R"], rv["R"] - 1, "期末 ARR 指引隱含約 $10.0–10.5m；7 月新約漲價約 25%"),
    ("每 MW 建置成本（US$m）", rv["cost30"], rv["cost30"] * rv["C"], rv["C"] - 1, "FY26 指引隱含約 $32–37m"),
    ("穩態 EBITDA 率（FY30）", rv["eb30"], rv["Eb"] if rv["Eb"] is not None else "無解（>99%）", rv["Eb"] - rv["eb30"] if rv["Eb"] is not None else "—", "Q2 實際 59%；由下而上上緣約 67–71%（變動為百分點）；區間 30–99% 內無解時顯示「無解」"),
    ("（對照）加權目標價＝現價所需每 MW 年收入", rv["rev30"], rv["rev30"] * rv["Rt"], rv["Rt"] - 1, "含 EV/EBITDA 6x（FY29 錨定） 腿；非純反向 DCF"),
]
RV_R0 = r  # v4.3：單一槓桿快照第一列（「摘要」頁引用）
for k, (nm, a, b, d, note) in enumerate(_rows):
    ws.cell(row=r, column=1, value=nm).font = BLACK
    fm = PCT if k == 2 else NUM1
    for col, v in [(2, a), (3, b)]:
        c = ws.cell(row=r, column=col, value=round(v, 4) if isinstance(v, (int, float)) else v); c.number_format = fm; c.border = BOX
    c = ws.cell(row=r, column=4, value=round(d, 4) if isinstance(d, (int, float)) else d); c.number_format = PCT; c.border = BOX
    ws.cell(row=r, column=7, value=note).font = SMALL
    r += 1
r += 1
r = section(ws, r, "快照：收入與成本的綜合影響——要值現價，FY30 每 MW 年收入需要多少（US$m）", span=7)
c = ws.cell(row=r, column=1, value="建置成本 ＼ 穩態 EBITDA 率"); c.font = HEAD; c.fill = FILL_HEAD
for j, e in enumerate(rv["ebs"]):
    c = ws.cell(row=r, column=2 + j, value=e); c.font = HEAD; c.fill = FILL_HEAD; c.number_format = '0%'
r += 1
for i, cp in enumerate(rv["caps"]):
    c = ws.cell(row=r, column=1, value=f"${rv['cost30']*cp:.1f}m（" + ("目前" if cp == 1 else f"{(cp-1)*100:+.0f}%") + "）"); c.font = BLACK
    for j, x in enumerate(rv["grid"][i]):
        v = rv["rev30"] * x if x is not None else None
        c = ws.cell(row=r, column=2 + j, value=round(v, 2) if v is not None else "—"); c.number_format = NUM1; c.border = BOX
        if cp == 1 and abs(rv["ebs"][j] - rv["eb30"]) < 1e-6:
            c.fill = FILL_KEY
    r += 1
ws.cell(row=r, column=1, value="快照為靜態數值（HTML 以同一引擎即時計算）；改變輸入後請以目標搜尋重算，或參考 HTML『評價 → 反向 DCF』").font = SMALL

# =====================================================================
# 6e. 檢查_版本紀錄
# =====================================================================
ws = wb.create_sheet("檢查_版本紀錄")
for c, w in zip("ABCDE", [12, 10, 90, 14, 48]):
    ws.column_dimensions[c].width = w
ws["A1"] = "版本紀錄（基準情境目標價變化與原因）"
ws["A1"].font = TITLE
ws["A2"] = "v2.9 以前檔名日期統一沿用 20260918，此處日期為實際修改日。v1.7 以前基準為 8 GW、v1.7 為 5 GW、v1.8 起為 5.6 GW；情境定義改變的版本不可直接比較。目標價以當版預設輸入計算（DCF 0 截斷）"
ws["A2"].font = SMALL
r = 4
for j, h in enumerate(["版本", "日期", "主要變更", "基準目標價", "變動原因"]):
    c = ws.cell(row=r, column=1 + j, value=h); c.font = HEAD; c.fill = FILL_HEAD
r += 1
for v in reversed(VLOG_X):
    for j, x in enumerate(v):
        c = ws.cell(row=r, column=1 + j, value=x); c.font = BOLD if j == 3 else BLACK
        c.alignment = Alignment(wrap_text=j in (2, 4), vertical="top")
    r += 1

# =====================================================================
# 7. 連動檢查
# =====================================================================
ws = wb.create_sheet("檢查_連動")
ws.column_dimensions["A"].width = 44
ws.column_dimensions["B"].width = 16
ws.column_dimensions["C"].width = 14
ws.column_dimensions["D"].width = 12
ws.column_dimensions["E"].width = 100
ws["A1"] = "連動檢查 — 恆等式與口徑"
ws["A1"].font = TITLE
r = 3
for j, h in enumerate(["檢查項目", "實際值", "標準", "結果", "說明"]):
    c = ws.cell(row=r, column=1 + j, value=h)
    c.font = HEAD
    c.fill = FILL_HEAD
r += 1
CK, CFq = CO['methodology']['checks'], CO['callFacts']  # 5a：檢查門檻讀 company.json → methodology.checks；指引區間讀 callFacts
_CXLO, _CXHI, _RVLO, _RVHI = (_n(CFq[k]) for k in ('capexLo', 'capexHi', 'revLo', 'revHi'))
_INT_MIN = _n(round(YA['interest'] + CAL['stubMonths'] // 3 * CFq['nextQIntLo'], 1))  # 首期利息下限＝年初至今實際＋剩餘季數 × 下一季指引下緣
_JV0 = _n(round(sum(D['jvCommit']), 6))
_YR0 = CAL['periodEnd'][0][:4]
checks = [
    ("RPO 桶權重合計", f"=SUM('輸入與假設'!C{IN['RPO 桶權重']}:G{IN['RPO 桶權重']})", f"={WSUM}",
     "=IF(ABS(B{r}-C{r})<0.0001,\"通過\",\"不一致\")", PCT,
     "10-Q 桶 41／39／20 線性分攤後應為 84%"),
    ("現金橋差異", f"='各期收支'!C{FRR['bridge_diff']}", "=0",
     "=IF(ABS(B{r})<0.01,\"通過\",\"不一致\")", NUM,
     "期初＋營運缺口＋股權＋額度−還本 ＝ 期末"),
    ("Billable ≤ Accepted（FY30）", f"='運營_產能與收入'!G{CAP['acc']}-'運營_產能與收入'!G{bil_row}", "≥0",
     "=IF(B{r}>=0,\"通過\",\"不一致\")", NUM0,
     "引擎已以 MIN 截斷；此列驗證截斷有效"),
    ("«P0» CapEx 認列 vs 指引", f"='各期收支'!C{FRR['memo_capex']}", f"{_CXLO}–{_CXHI}",
     "=IF(AND(B{r}>=" + _CXLO + ",B{r}<=" + _CXHI + "),\"通過\",\"觀察\")", NUM,
     f"«YTD» 實際認列 {CO['ytdActual']['capex']}＋«STUB» 模型；公司 8/11 上修全年 {CO['callFacts']['capexLo']}–{CO['callFacts']['capexHi']}"),
    ("«P0» 營收 vs 指引", f"='各期收支'!C{FRR['memo_rev']}", f"{_RVLO}–{_RVHI}",
     "=IF(AND(B{r}>=" + _RVLO + ",B{r}<=" + _RVHI + "),\"通過\",\"觀察\")", NUM,
     f"«YTD» 實際 {CO['ytdActual']['revenue']}＋«STUB» 模型（算力＋非算力服務）"),
    ("CapEx 強度（模型期合計）", None, "$28–34m/MW",
     "=IF(AND(B{r}>=" + _n(CK['capexPerMwBand'][0]) + ",B{r}<=" + _n(CK['capexPerMwBand'][1]) + "),\"通過\",\"觀察\")", NUM0,
     "公司自身比率：2026 CapEx 35–39bn 對應約 +1.25 GW ＝ $28–34m/MW"),
    ("表外租金路徑 ÷ 已承諾", f"='資產負債_租賃承諾'!C{lease_ratio_row}", f"≥{_n(CK['leaseVsCommitMin'])}x",
     "=IF(B{r}>=" + _n(CK['leaseVsCommitMin']) + ",\"通過\",\"不一致\")", MULT,
     "8 GW 路徑所需新租約多數未簽，路徑高於承諾屬假設"),
    ("«P0» 利息（«YTD» 實際＋«STUB»）", f"='各期收支'!C{FRR['memo_int']}", f"≥{_INT_MIN}",
     "=IF(B{r}>=" + _INT_MIN + ",\"通過\",\"觀察\")", NUM,
     # v4.5：首期模型部分的存量利息讀「各期收支」③ 列（文字公式，滾動後隨模型值更新）
     f'="«YTD» 實際 {YA["interest"]}＋«STUB» 模型 "&TEXT(\'各期收支\'!C{FR["③ 存量債務利息（備忘，«YTD» 已含在 CFO）"]},"0.0")&"；Q3 指引 {CO["callFacts"]["nextQIntLo"]}–{CO["callFacts"]["nextQIntHi"]}／季"'),
    ("FY28 起收入是否依賴未簽約", None, f"<{_n(CK['unsignedRevShareMax'] * 100)}% 較穩健",
     "=IF(B{r}<" + _n(CK['unsignedRevShareMax']) + ",\"通過\",\"觀察\")", PCT,
     "FY28–30 新簽約收入 ÷ 同期總算力收入。高＝8 GW 敘事的風險集中點"),
    ("損益營收＝資金營收", None, "=0",
     "=IF(ABS(B{r})<0.001,\"通過\",\"不一致\")", NUM,
     "FY27 損益算力收入 − 產能頁 isRev。兩邊必須同一組數字"),
    ("終值占 EV 比重", f"='評價_DCF與目標價'!C{d0+8}", f'="<"&{_PC(RT_TVW)}',
     "=IF(B{r}<" + RT_TVW + ",\"通過\",\"觀察\")", PCT,
     "過高表示結論由終值假設決定"),
    ("EBITDA 單一來源（FY27 損益 EBITDA÷營收 − 輸入 EBITDA 率）", None, "=0",
     "=IF(ABS(B{r})<0.0001,\"通過\",\"不一致\")", '0.0000', "損益與資金引用同一列 EBITDA 率"),
    ("現金 EBITDA ＋ 信用調整 − 損益 EBITDA（FY27）", None, "=0",
     "=IF(ABS(B{r})<0.001,\"通過\",\"不一致\")", NUM, "現金 EBITDA＝RPO 現金＋新簽約現金＋服務現金−租金；信用調整＝信用損失×EBITDAR 率"),
    ("DCF 有效性（1＝失效）", None, "0",
     "=IF(B{r}=0,\"通過\",\"觀察\")", NUM0, "失效只定義為 WACC ≤ g 或常態化 FCF ≤ 0；股權為負不算失效"),
    ("期末現金 ≥ 最低現金（期前融資）", None, f"≥{D['minCash']:.1f}",
     "=IF(B{r}>=" + _n(D['minCash'] - 0.001) + ",\"通過\",\"不一致\")", NUM, "瀑布每期補足至最低現金；最小值應等於最低現金"),
    ("股權募資 ÷ 現市值", None, f'="≤"&{_MT(RT_EQ)}&"x"',
     "=IF(B{r}<=" + RT_EQ + ",\"通過\",\"觀察\")", MULT, f'="超過 "&{_MT(RT_EQ)}&" 倍：市場吸收能力存疑，結論不得為買進"'),
    (f"債務明細合計 = 10-Q 本金 {_n(LQ_DEBT)}", None, _n(LQ_DEBT),
     "=IF(ABS(B{r}-" + _n(LQ_DEBT) + ")<0.01,\"通過\",\"不一致\")", NUM, "16 筆工具逐筆加總"),
    (f"{PERIODS[0]} 全年 CapEx（MW 公式）", None, f"{_CXLO}–{_CXHI}",
     "=IF(AND(B{r}>=" + _CXLO + ",B{r}<=" + _CXHI + "),\"通過\",\"觀察\")", NUM, "由 YE25 850→YE26 1,850 MW 與預建 FY27 推得；應重現公司全年指引"),
    (f"JV 已承諾餘額於 {_YR0} 年內履行", None, f"={_JV0}",
     "=IF(ABS(B{r}-" + _JV0 + ")<0.01,\"通過\",\"不一致\")", NUM, f"10-Q：1.7 承諾預計 2026 年內履行，«YTDA» 已付 {YA['jvSplit']['jv']:.2f}"),
    ("模型每 MW 年租金（FY30）", None, "≥ 市場基準×占比",
     "=IF(B{r}>=0,\"通過\",\"觀察\")", NUM, "低於基準×占比的八成即標示：模型租金路徑可能低估"),
    ("站點租賃：五期租金可能低估", None, f"<{_n(CK['siteRentGapMax'])}",
     "=IF(B{r}<" + _n(CK['siteRentGapMax']) + ",\"通過\",\"觀察\")", NUM, "＝Σ(在役 MW×市場基準×第三方占比 − 模型租金)"),
]
if TK_MISS and (PMW['capex'] != 'legacy' or PMW['cost'] != 'ebitdaPct' or PMW['revenue'] != 'legacy'):  # W2：缺漏時顯示，Tokenomics 補齊名稱並重抓快照後自動消失（HTML 連動檢查同一則）
    checks.append((f"Tokenomics 名稱缺漏 {len(TK_MISS)} 項，成本為暫代值", f"={len(TK_MISS)}", "0 項",
                   "=IF(B{r}=0,\"通過\",\"警告\")", NUM0, f"{'、'.join(TK_MISS)} 尚未在 Tokenomics {CO['tokenomics']['version']} 提供（待 v5.26）；暫代：IT 維護＝C 區維護成本、人員軟體與稅險＝0、GPU 經濟壽命＝{D['gpuLife']} 年"))

REVCAP_CHECK = "收入上限：CRWV 每 MW 計費收入 ÷ 客戶付費 token 營收（各期最高）"  # W4：只在 revenue=tkAnchor 且快照有 IF_RevGWFleet 時建列（B 欄在『每MW經濟性』頁建好後回填）
if PMW['revenue'] == 'tkAnchor' and tk_has('IF_RevGWFleet'):
    checks.append((REVCAP_CHECK, None, f"≤{_n(CK['revCapShareMax'] * 100)}%",
                   "=IF(B{r}<=" + _n(CK['revCapShareMax']) + ",\"通過\",\"警示\")", PCT,
                   "neocloud 拿走客戶 token 營收的比例（Tokenomics IF_RevGWFleet，在役世代加權）；超過門檻＝CRWV 單價相對客戶端 token 經濟性偏高，明細見『每MW經濟性』"))

for nm, bf, std, rf, fmt, note in checks:
    ws.cell(row=r, column=1, value=nm).font = BLACK
    if bf:
        c = ws.cell(row=r, column=2, value=bf)
        c.number_format = fmt
        c.font = BLACK
    ws.cell(row=r, column=3, value=std).font = SMALL
    ws.cell(row=r, column=4, value=rf.format(r=r)).font = BOLD
    ws.cell(row=r, column=5, value=note).font = SMALL
    r += 1
# 補上需要跨頁計算的三列
def _find(label):
    for rr in range(4, 80):
        if ws.cell(row=rr, column=1).value == label:
            return rr
    raise KeyError(label)
_r = _find(f"債務明細合計 = 10-Q 本金 {_n(LQ_DEBT)}")
ws.cell(row=_r, column=2, value=f"='資產負債_既有債務'!E{tot_r}").number_format = NUM
_r = _find(f"{PERIODS[0]} 全年 CapEx（MW 公式）")
ws.cell(row=_r, column=2, value=f"='輸入與假設'!C{IN['全年毛 CapEx（公式）']}").number_format = NUM
_r = _find(f"JV 已承諾餘額於 {_YR0} 年內履行")
ws.cell(row=_r, column=2, value=f"=SUM('輸入與假設'!C{IN['JV 已承諾餘額出資']}:G{IN['JV 已承諾餘額出資']})").number_format = NUM
_r = _find("模型每 MW 年租金（FY30）")
ws.cell(row=_r, column=2, value=f"='運營_站點'!G{RENT_PER}").number_format = NUM
_share_x = SHARE.replace('$C$', "'運營_站點'!$C$")  # 另存變數：f-string 內重用引號需 Python 3.12+
ws.cell(row=_r, column=4, value=f"=IF(B{_r}>={BENCH}*{_share_x}*{_n(CK['rentVsBenchMin'])},\"通過\",\"觀察\")")
_r = _find("EBITDA 單一來源（FY27 損益 EBITDA÷營收 − 輸入 EBITDA 率）")
ws.cell(row=_r, column=2, value=f"='損益'!D{ebitda_v}/'損益'!D{rev_v}-'輸入與假設'!D{IN['EBITDA 率']}").number_format = '0.0000'
_r = _find("現金 EBITDA ＋ 信用調整 − 損益 EBITDA（FY27）")
ws.cell(row=_r, column=2, value=(f"=('運營_產能與收入'!D{CAP['rpocash']}+'運營_產能與收入'!D{CAP['newcash']}+'輸入與假設'!D{IN['非算力服務現金']}-'各期收支'!D{FR['　租金合計']})"
    f"+('運營_產能與收入'!D{CAP['loss']}+'運營_產能與收入'!D{CAP['newrev']}*'輸入與假設'!D{IN['客戶違約率']}*(1-'輸入與假設'!D{IN['回收率']}))*'運營_產能與收入'!D{CAP['cm']}"
    f"-'損益'!D{ebitda_v}")).number_format = NUM
_r = _find("DCF 有效性（1＝失效）")
ws.cell(row=_r, column=2, value=f"='評價_DCF與目標價'!{DCF_BAD}")
_r = _find("期末現金 ≥ 最低現金（期前融資）")
ws.cell(row=_r, column=2, value=f"=MIN('各期收支'!C{FRR['cum']}:G{FRR['cum']})").number_format = NUM
_r = _find("股權募資 ÷ 現市值")
ws.cell(row=_r, column=2, value=f"=SUM('各期收支'!C{FRR['eq']}:G{FRR['eq']})/({PX}*{SH})").number_format = MULT
_r = _find("站點租賃：五期租金可能低估")
ws.cell(row=_r, column=2, value=f"={LEASE_GAP}").number_format = NUM
ci = 4  # first check row
ws.cell(row=ci + 5, column=2,
        value=(f"=SUM('各期收支'!C{FR['① 毛 CapEx（認列，備忘）']}:G{FR['① 毛 CapEx（認列，備忘）']})/"
               f"(('運營_產能與收入'!G{CAP['acc']}-{MW0})/1000)/1000*1000")).number_format = NUM0
ws.cell(row=ci + 8, column=2,
        value=(f"=SUM('運營_產能與收入'!E{CAP['newrev']}:G{CAP['newrev']})/"
               f"SUM('運營_產能與收入'!E{CAP['isrev']}:G{CAP['isrev']})")).number_format = PCT
ws.cell(row=ci + 9, column=2,
        value=f"='損益'!D{VR['算力收入']}-'運營_產能與收入'!D{CAP['isrev']}").number_format = NUM

r += 1
ws.cell(row=r, column=1, value="關鍵輸出（快照；「模型期」＝2026-07-01～2030，與 HTML 頂部 KPI 同口徑）").font = BOLD
r += 1
snap = [
    ("模型期排程 RPO", f"=SUM('運營_產能與收入'!C{CAP['sch']}:G{CAP['sch']})", NUM),
    ("模型期產能瓶頸", f"=SUM('運營_產能與收入'!C{CAP['bot']}:G{CAP['bot']})", NUM),
    ("模型期新簽約收入", f"=SUM('運營_產能與收入'!C{CAP['newrev']}:G{CAP['newrev']})", NUM),
    ("模型期毛 CapEx（«VMD» 起）", f"=SUM('輸入與假設'!C{IN['毛 CapEx']}:G{IN['毛 CapEx']})", NUM),
    ("模型期租金（«VMD» 起）", f"=SUM('各期收支'!C{FR['　租金合計']}:G{FR['　租金合計']})", NUM),
    ("模型期利息（«VMD» 起，含瀑布新債）", f"=SUM('各期收支'!C{FRR['int']}:G{FRR['int']})", NUM),
    ("模型期排程還本（«VMD» 起）", f"=SUM('輸入與假設'!C{IN['排程還本（10-Q 本金表）']}:G{IN['排程還本（10-Q 本金表）']})*{DEBTON}", NUM),
    ("模型期營運缺口（«VMD» 起）", f"=SUM('各期收支'!C{FRR['opgap']}:G{FRR['opgap']})-({H_CFO}-{H_CCAPEX}-{H_JV})", NUM),
    ("«P0» 全年口徑營運缺口（含 «YTD» 實際）", f"=SUM('各期收支'!C{FRR['opgap']}:G{FRR['opgap']})", NUM),
    ("期末現金", f"='各期收支'!G{FRR['cum']}", NUM),
    ("DCF 每股", f"='評價_DCF與目標價'!{DCF_PS}", USD),
    ("加權目標價", f"='評價_DCF與目標價'!{TGT}", USD),
]
for nm, f, fmt in snap:
    ws.cell(row=r, column=1, value=nm).font = BLACK
    c = ws.cell(row=r, column=2, value=f)
    c.number_format = fmt
    c.font = BLACK
    r += 1

# =====================================================================
# 8. 來源
# =====================================================================
ws = wb.create_sheet("來源")
ws.column_dimensions["A"].width = 16
ws.column_dimensions["B"].width = 120
ws["A1"] = "來源與標籤"
ws["A1"].font = TITLE
r = 3
for j, h in enumerate(["標籤", "內容"]):
    c = ws.cell(row=r, column=1 + j, value=h)
    c.font = HEAD
    c.fill = FILL_HEAD
r += 1
# v4.3：同業市值讀 company.json → peers／callFacts；目標價讀共識資料檔。與 HTML consSourceTxtQ() 同一字串
_PE, _PT = CO['peers'], CONS['priceTarget']
_SRC_TXT = {
    "peers": (f"市值（{_PE['priceDate']} 收盤，{_PE['priceSource']}）：CRWV 現價 ${CO['callFacts']['priceLast']:,.2f}（{CO['callFacts']['priceDate']} 收盤）、"
              f"市值 {CO['callFacts']['mktCapLast']:,.2f}bn（{_PE['priceDate']}）、流通 {CO['latestQuarter']['sharesOut'] * 1e3:,.2f}m；"
              + "；".join(f"{p['ticker']} {p['mkt']:,.2f}" for p in _PE['list']) + "。淨負債取各公司最新申報："
              + "、".join(f"{p['ticker']} {p['netDebt']:,.2f}" for p in _PE['list']) + "。"),
    "targets": (f"賣方目標價（{_PT['source']}，擷取 {_PT['retrieved']}）：平均 ${_PT['mean']:,.2f}、中位數 ${_PT['median']:,.2f}、區間 ${_PT['low']:,.2f}–${_PT['high']:,.2f}（{_PT['analysts']} 家）；"
                f"共識評等 {_PT['consensusRating']}。對照 {_PT['crossCheck']['source']}：平均 ${_PT['crossCheck']['mean']:,.2f}（{_PT['crossCheck']['analysts']} 家，{_PT['crossCheck']['consensusRating']}）。"
                "數字為賣方意見；逐筆來源與日期見『輸入與假設』I 區。"),
}
src = [
    ("Verified", "Q2 2026 10-Q（期末 2026-06-30，2026-08-12 申報，Non-accelerated filer，Q3 期限 11/16）：營收 2.575bn +112%（H1 4.653）；"
     "GAAP 營損 −0.049；利息費用 0.640（H1 1.176，現金 0.806、資本化 0.176）；淨損 −0.626；Adj. EBITDA 1.510（59%）；"
     "SBC 0.165；D&A 1.393；營業租賃成本 0.500、變動租賃 0.150。"),
    ("Verified", "資產負債：現金 5.524、受限 1.380、有價證券 0.015、可用額度 10.014；PP&E 46.736（CIP 11.9）；營業 ROU 16.595；"
     "遞延收入 9.692；債務本金 35.551（追索淨 31.405／非追索 3.663）；股東權益 5.024。"),
    ("Verified", "RPO 103.7bn（backlog 104.2）：41% 於 2028-06 前、39% 於 25–48 月、20% 於 49–78 月。"
     "客戶集中：A 36%、B 26%、C 10%；OpenAI 最高 6.5bn 至 2031-05；Meta 最高 21.0bn；Jane Street 約 6.0bn。"),
    ("Verified", "租賃與承諾：在帳未折現 29.135（營業）＋0.235（融資），2030 後 19.023；未起租租賃 35.5；"
     "單站 393 MW 按造價計租上限 14.7／16 年；355 MW 按造價計租（未定）；自備設備 0.5–1.2；JV 承諾 1.7（已付 0.55）；VIE 0.108。"),
    ("Verified", "現金流（H1）：CFO 3.663；現金購置 PP&E 14.117；借款 16.747、還款 5.219；capped call 0.492；"
     "私募股權 2.982。CapEx 認列口徑 Q2 9.352、H1 16.139。"),
    ("Verified", "9/17 8-K：2033 可轉債 3.7bn、2.875%、轉換價 $97.85、capped call 上限 $199.70；股權分銷最多 35m 股；9/17 收盤 $79.88。"),
    ("Verified", _SRC_TXT["peers"]),
    (CONS['priceTarget']['tag'], _SRC_TXT["targets"]),
    ("Interested-party", "法說（8/11）：主動電力 1.5 GW（Q2 +近 500 MW，6 月 +300 MW）；簽約 4.2 GW；2030 ≥8 GW；"
     "近期產能實質售罄；7 月 SKU 漲價約 25%；全年營收 12.4–13.2、調整後營益 0.96–1.15、CapEx 35–39、期末 ARR 18.5–19.5；"
     "Q3 營收 3.45–3.6、利息 0.86–0.94、CapEx 11.5–13.5。"),
    ("Interested-party", "9/17 投資人簡報：Q2 pro forma 總債務 38.6bn；加權利率 FY23 14.9% → Q2 8.3% → pro forma 7.8%；"
     "總債務／backlog 0.4x；歷年 backlog 9.9／15.1／66.8／104.2。"),
    ("User-provided", "5Y CDS（2026-09）：中價 720bps、Bid/Ask 690／750、近期區間 680–855（對照 7/28 峰值 855、6 月低點 452）。"),
    ("Derived", "RPO 桶內線性分攤（41/39/20 → 五期 10.25／20.5／20／19.5／13.75%）；"
     "EBITDAR 率＝EBITDA 率＋租金÷營收（v2.3 起與損益連動）；客戶預付率 8.5%；rev/MW 錨 10.0–10.5M/MW。"),
    ("Assumed", "FY27–30 毛 CapEx 45／50／50／45；新產能簽約率 100→80%；表外現金租金 0.3→4.6；"
     "違約率 0.5→2.5%；瀑布新債利率 9%；債務／backlog 1.0x（敏感度 0.4x）；起始／穩態 EBITDA 率 59%／65%（敏感度 59%／70%）；最低現金 2.0；股權折價 10%；非算力服務營收 0.3→3.8；終值維持性 CapEx 占 D&A 80%。"),
    ("Interested-party", "房東揭露：Galaxy Helios 526 MW、年均 >$1bn、15 年（2026-07-06）；Applied Digital Polaris Forge 1 400 MW、約 $11bn、約 15 年（2025-08-29）；"
     "Core Scientific 約 590 MW、$10.2bn、12 年（2025-02-26）。Q4 2025 法說：年底主動電力 >850 MW、2025 CapEx 14.9bn。"),
    ("Verified", "10-Q Note 10 債務逐筆（有效利率）：DDTL 1.0 1.300@15%、2.0 3.190@11%、2.1 3.000@9%、3.0 2.215@9%、5.0 1.101@9%、4.0 2.837@7%（非追索）；"
     "票據 2030 2.000、2031 1.750＋2.750、2032 1.250@10%、EUR 2.279@9%；可轉債 2.588＋4.000@2%；OEM 4.220@11%＋0.882@9%；Magnetar 0.189@12%。"
     "Note 3：兩個 JV 承諾最高 1.7，預計 2026 年內履行，之後須追加出資；新澤西 JV 持股 35%，租金依建造成本計。"),
    ("方法", "v2.0 期前融資瀑布：額度→資產層新債（總債務 ≤ 債務/backlog 上限）→股權殘差；舊版假設缺口全額舉債且無上限（透支利息＝利息、扣缺口＝本金），v2.0 改為債務設上限、殘差以股權支應。"),
]
for tag, text in src:
    ws.cell(row=r, column=1, value=tag).font = BOLD
    c = ws.cell(row=r, column=2, value=text)
    c.font = BLACK
    c.alignment = Alignment(wrap_text=True, vertical="top")
    ws.row_dimensions[r].height = 46
    r += 1

# v4.1：情境區間的模擬運算表（Excel「模擬分析 → 運算列表」；輸入格＝情境選擇）。
# Excel 規定輸入格與運算表須在同一工作表，所以放在『輸入與假設』底部。LibreOffice 重算後會改寫成 TABLE()，
# 由 fix_datatable.py 還原為 OOXML 的 dataTable 公式（保留 LibreOffice 算出的值）。
ws = wb["輸入與假設"]
assert ws.max_row + 3 == DT_H, "輸入與假設在評價頁建立後又新增了列，模擬運算表位置需重算"
r = section(ws, DT_H - 1, "H｜情境區間運算表（模擬運算表；由 Excel 自動計算，勿手動輸入）")
ws.cell(row=DT_H - 1, column=9, value="保守與積極情境的加權目標價（情境區間）；其餘手動輸入保留。資料表公式：{=TABLE(,情境選擇)}").font = SMALL
ws.cell(row=DT_H, column=1, value="情境區間運算表｜公式列（加權目標價／評等代碼）").font = BLACK
for j, f in enumerate([f"='評價_DCF與目標價'!{TGT}", f"='評價_DCF與目標價'!{RG_PT_CODE}"]):
    c = ws.cell(row=DT_H, column=4 + j, value=f); c.font = BLACK; c.number_format = USD if j == 0 else NUM0
from openpyxl.worksheet.formula import DataTableFormula as _DTF
for j, k in enumerate(DT_SC):
    rr = DT_H + 1 + j
    ws.cell(row=rr, column=1, value=f"情境區間運算表｜{DT_NM[k]}").font = BLACK
    c = ws.cell(row=rr, column=3, value=j + 1); c.font = BLACK; c.number_format = NUM0
ws.cell(row=DT_H + 1, column=4, value=_DTF(ref=f"D{DT_H+1}:E{DT_H+3}", dt2D=False, dtr=False, r1=SEL.split("!")[1].replace("$", "")))
for rr in range(DT_H + 1, DT_H + 4):
    ws.cell(row=rr, column=4).number_format = USD; ws.cell(row=rr, column=5).number_format = NUM0

# =====================================================================
# v4.3：I｜市場共識（「輸入與假設」底部，接在 H 區模擬運算表之後，以免移動運算表）
# 逐筆與 HTML consensusItems() 同一組列（標籤相同；cmp31 以標籤比對數值、擷取日期、標記）。數字只讀共識資料檔，不增補、不修改。
# 欄位：A 項目｜B 單位｜C–E 數值（年度列＝FY26–FY28、季度列＝Q3／Q4、區間列＝低／高）｜F 擷取日期｜G 標記｜I 來源｜J 備註
# =====================================================================
def cons_items():
    C = CONS; it = []
    def add(sec, label, vals, unit, m, **ex):
        it.append(dict(sec=sec, label=f"共識｜{label}", vals=vals, unit=unit, src=m.get('source', ''), url=m.get('url', ''),
                       date=m.get('retrieved') or m.get('date') or '未列', tag=m.get('tag', ''), note=ex.get('note', ''), text=ex.get('text', '')))
    PR, PT, RA, AE, AL = C['priceReference'], C['priceTarget'], C['ratings'], C['annualEstimates'], C['annualEstimatesAlt']
    XC, QE, CG, IC, RM = PT['crossCheck'], C.get('quarterlyEstimates', {}), C['companyGuidance'], C['independentCrossCheck'], C['recentActionsMeta']
    S1 = "價格與目標價"
    add(S1, "現價參考（收盤）", [PR['close']], "US$", {**PR, 'retrieved': PR['date']}, note=f"收盤日 {PR['date']}")
    for a, x in [("平均", PT['mean']), ("中位數", PT['median']), ("最低", PT['low']), ("最高", PT['high'])]:
        add(S1, f"目標價｜{a}", [x], "US$", PT)
    add(S1, "目標價｜分析師家數", [PT['analysts']], "家", PT)
    add(S1, "目標價｜共識評等", [], "", PT, text=PT['consensusRating'])
    XM = {**XC, 'tag': PT['tag']}
    add(S1, "目標價｜MarketScreener 對照平均", [XC['mean']], "US$", XM)
    add(S1, "目標價｜MarketScreener 對照家數", [XC['analysts']], "家", XM)
    add(S1, "目標價｜MarketScreener 對照評等", [], "", XM, text=XC['consensusRating'])
    for a, k in [("強力買進", 'strongBuy'), ("買進", 'buy'), ("持有", 'hold'), ("賣出", 'sell'), ("強力賣出", 'strongSell'), ("合計", 'total')]:
        add(f"評等分布（{RA['month']}）", f"評等分布｜{a}", [RA[k]], "家", RA)
    SA, YR = "年度共識（FY26／FY27／FY28）", ["FY26", "FY27", "FY28"]
    for a, k, u in [("營收", 'revenue', "US$bn"), ("調整後 EBITDA", 'ebitda', "US$bn"), ("EBITDA 率", 'ebitdaMargin', "%"), ("EBIT", 'ebit', "US$bn"),
                    ("利息費用", 'interest', "US$bn"), ("淨利", 'netIncome', "US$bn"), ("GAAP EPS", 'epsGaap', "US$"), ("CapEx", 'capex', "US$bn"),
                    ("自由現金流", 'fcf', "US$bn"), ("淨負債", 'netDebt', "US$bn")]:
        add(SA, f"年度｜{a}", [AE[y][k] for y in YR], u, AE)
    add(SA, "年度｜口徑說明", [], "", AE, text=AE['definition'])
    SB = "S&P 對照（經 StockAnalysis.com；FY26／FY27）"
    add(SB, "S&P 對照｜營收", [AL['FY26']['revenue'], AL['FY27']['revenue']], "US$bn", AL)
    add(SB, "S&P 對照｜FY26 營收區間（低／高）", [AL['FY26']['revenueLow'], AL['FY26']['revenueHigh']], "US$bn", AL)
    add(SB, "S&P 對照｜調整後 EPS", [AL['FY26']['epsAdjusted'], AL['FY27']['epsAdjusted']], "US$", AL)
    add(SB, "S&P 對照｜FY26 調整後 EPS 區間（低／高）", [AL['FY26']['epsAdjustedLow'], AL['FY26']['epsAdjustedHigh']], "US$", AL)
    add(SB, "S&P 對照｜FY26 分析師家數", [AL['FY26']['analysts']], "家", AL, note=AL['note'])
    import re as _re_c
    QK = [k for k in (QE or {}) if _re_c.match(r'^\d{4}Q\d$', k)]  # v4.4：季度依共識檔實際列出的季別（沒有季度共識時不列）
    SQ = f"季度共識（{'／'.join(QK)}）"
    if QK:
        for a, k in [("營收", 'revenue'), ("EBITDA", 'ebitda'), ("EBIT", 'ebit'), ("淨利", 'netIncome')]:
            add(SQ, f"季度｜{a}", [QE[q][k] for q in QK], "US$bn", QE)
        add(SQ, "季度｜說明", [], "", QE, text=QE['note'])
    SG = "公司指引（管理層預估）"
    add(SG, "公司指引｜FY26 營收（低／高）", [CG['FY26']['revenueLow'], CG['FY26']['revenueHigh']], "US$bn", CG)
    add(SG, "公司指引｜FY26 調整後營業利益（低／高）", [CG['FY26']['adjOpIncomeLow'], CG['FY26']['adjOpIncomeHigh']], "US$bn", CG)
    add(SG, "公司指引｜FY26 CapEx（低／高）", [CG['FY26']['capexLow'], CG['FY26']['capexHigh']], "US$bn", CG)
    add(SG, "公司指引｜Q3 營收（低／高）", [CO['callFacts']['nextQRevLo'], CO['callFacts']['nextQRevHi']], "US$bn", CG, note=CG['tagNote'])  # 以 company.json 為準（建置時已檢查與共識檔一致）
    SL, LM = f"獨立對照（{IC['provider']}）", {**IC, 'source': f"{IC['provider']}（經 {IC['source']}）"}
    add(SL, f"{IC['provider']} 對照｜FY26 營收（Q2 前）", [IC['FY26RevenuePreQ2']], "US$bn", LM)
    add(SL, f"{IC['provider']} 對照｜Q3 營收（Q2 前）", [IC['2026Q3RevenuePreQ2']], "US$bn", LM, note=IC['note'])
    for x in C['recentActions']:
        add("最新分析師動作", f"分析師動作｜{x['date']} {x['firm']}", [] if x['target'] is None else [x['target']], "US$", RM,
            text="目標價未列" if x['target'] is None else "", note=x['rating'] + (f"；{x['note']}" if x.get('note') else ""))
    add("來源與限制", "來源獨立性", [], "", {'source': "資料檔說明", 'retrieved': C['asOf']}, text=C['sourceIndependence'])
    for i, t in enumerate(C['notFound']):
        add("來源與限制", f"未取得｜{i + 1}", [], "", {'source': "資料檔說明", 'retrieved': C['asOf']}, text=t)
    return it


ws = wb["輸入與假設"]
r = DT_H + 5
r = section(ws, r, f"I｜市場共識（{CO['meta']['consensusFile']}；只讀，不增補數字；下游只引用本區）")
ws.cell(row=r - 1, column=9, value=f"標記：Verified＝{CONS['tagLegend']['Verified']}；Interested-party＝{CONS['tagLegend']['Interested-party']}").font = SMALL
ws.column_dimensions["G"].width = 16
CI = {}
_sec = None
_FMT = {"US$": USD, "US$bn": '#,##0.000', "%": PCT, "家": NUM0, "": NUM}
for x in cons_items():
    if x['sec'] != _sec:
        _sec = x['sec']
        for j, h in enumerate([_sec, "單位", "數值", "", "", "擷取日期", "標記", "", "來源", "備註"]):
            c = ws.cell(row=r, column=1 + j, value=h or None); c.font = HEAD; c.fill = FILL_HEAD
        r += 1
    CI[x['label']] = r
    ws.cell(row=r, column=1, value=x['label']).font = BLACK
    ws.cell(row=r, column=2, value=x['unit'] or None).font = SMALL
    if x['vals']:
        for j, v in enumerate(x['vals']):
            c = ws.cell(row=r, column=3 + j, value=v); c.font = BLUE; c.number_format = _FMT[x['unit']]; c.border = BOX
    else:
        ws.cell(row=r, column=3, value=x['text']).font = BLACK
    ws.cell(row=r, column=6, value=x['date']).font = SMALL
    ws.cell(row=r, column=7, value=x['tag'] or None).font = SMALL
    ws.cell(row=r, column=9, value=x['src'] + (f"｜{x['url']}" if x['url'] else "")).font = SMALL
    ws.cell(row=r, column=10, value=x['note'] or None).font = SMALL
    r += 1
CIR = lambda label, j=0: f"'輸入與假設'!${'CDE'[j]}${CI['共識｜' + label]}"
CONS_TOL = gi(r, "判斷句門檻：與共識差距（營收、EBITDA、CapEx）", "%", CO['methodology']['consensusGapTol'],
              "任一項差距絕對值超過此比例即視為分歧（company.json → methodology.consensusGapTol）[Assumed]", PCT); r += 1
PM_TOL = gi(r, "差異原因門檻：利潤類利潤率差（百分點）", "pt", CO['methodology'].get('profitMarginTolPt'),
            "利潤類（調整後 EBITDA、調整後營業利益）利潤率差超過此值須附差異原因；營收、CapEx、淨負債、MW 用上一格（company.json → methodology.profitMarginTolPt）[Assumed]", PCT) if CO['methodology'].get('profitMarginTolPt') is not None else None
if PM_TOL:
    r += 1

# =====================================================================
# v4.4：J｜季度追蹤輸入（「輸入與假設」底部，接在 I 區之後，不移動 H 區模擬運算表）＋「季度追蹤」分頁
# 設定全部讀 company.json → quarterly（季度、所屬期間、拆分方法、指引、實際數、焦點季）與 varianceReasons；共識讀 I 區；年度數字引用模型各頁。
# 規則與 HTML quarterlyView（segB）相同，cmp31 逐格比對。缺資料以文字「不適用」表示；實際數空白＝待公布。
# =====================================================================
import re as _re_q
QC = CO.get('quarterly')
VRS = [x for x in CO.get('varianceReasons', {}).get('list', []) if all(PMW.get(k) == v for k, v in (x.get('perMw') or {}).items())]  # W2：perMw 條件只在方法相符時適用
QCOL = ["C", "D", "E", "F", "G", "H", "I", "J"]
_FX = lambda x, dp: f'IF({x}<0,"−","")&TEXT(ABS({x}),"#,##0{"." + "0" * dp if dp else ""}")'  # 與 HTML Y() 同格式（千分位、固定小數）；負號用 −
_GX = lambda x, pt=False: f'IF({x}<0,"−","+")&TEXT(ABS({x})*100,"#,##0.0")&"{"pt" if pt else "%"}"'  # 與 HTML gapTxtQ 同格式


def _VX(x, u, dp=2):  # 與 HTML valTxtQ 同格式（US$bn 預設 2 位小數）
    if u == "%":
        return f'IF({x}<0,"−","")&TEXT(ABS({x})*100,"#,##0.0")&"%"'
    if u == "MW":
        return f'IF({x}<0,"−","")&TEXT(ABS({x}),"#,##0")&" MW"'
    return f'IF({x}<0,"−","")&"$"&TEXT(ABS({x}),"#,##0.{"0" * dp}")&"bn"'


def _RX(lo, hi, u):  # 與 HTML rngTxtQ 同格式
    if u == "MW":
        return f'{_FX(lo, 0)}&"–"&{_FX(hi, 0)}&" MW"'
    if u == "%":
        return f'{_FX(lo + "*100", 2)}&"–"&{_FX(hi + "*100", 2)}&"%"'
    return f'"$"&{_FX(lo, 2)}&"–"&{_FX(hi, 2)}&"bn"'


_POS = lambda x, lo, hi: f'IF({x}>{hi},"高於上緣",IF({x}<{lo},"低於下緣",IF({x}>=({lo}+{hi})/2,"區間上半部","區間下半部")))'


def _TOKX(text, resolve):  # 原因文字 {路徑:格式} → Excel 文字公式（與 HTML fillTokQ 同格式；路徑可用加減，例如 c.ebitda-g.adjOpInc）
    parts, pos = [], 0
    for mt in _re_q.finditer(r'\{([\w.+-]+):([%s]?\w+)\}', text):
        if mt.start() > pos:
            parts.append('"' + text[pos:mt.start()].replace('"', '""') + '"')
        terms = [t for t in _re_q.split(r'(?=[+-])', mt.group(1)) if t]
        refs = [resolve(t.lstrip("+-")) for t in terms]
        x = "(" + "".join((t[0] if t[0] in "+-" else ("+" if i else "")) + str(v) for i, (t, v) in enumerate(zip(terms, refs))) + ")"
        f = mt.group(2)
        if any(v is None for v in refs):  # 缺值（例如該季沒有共識或指引）：與 HTML 同樣顯示「不適用」
            parts.append('"不適用"')
        elif f == 'gap':
            parts.append(_GX(x))
        elif f[0] == '%':
            dp = int(f[1:])
            parts.append(f'IF({x}<0,"−","")&TEXT(ABS({x})*100,"#,##0{"." + "0" * dp if dp else ""}")&"%"')
        elif f[0] == 's':
            dp = int(f[1:])
            parts.append(f'IF({x}<0,"−","+")&TEXT(ABS({x}),"#,##0{"." + "0" * dp if dp else ""}")')
        else:
            parts.append(_FX(x, int(f)))
        pos = mt.end()
    if pos < len(text):
        parts.append('"' + text[pos:].replace('"', '""') + '"')
    return "&".join(parts) if parts else '""'


def _find_rsn(scope, period, metric, vs):
    for x in VRS:
        if x['scope'] == scope and x['period'] in ('*', period) and x['metric'] == metric and x['vs'] in ('*', vs):
            return x
    return None


RSN_TYPES = ["觀點", "拆法", "已知限制", "未歸類"]
QT = {}
if QC:
    ws = wb["輸入與假設"]
    r += 1
    r = section(ws, r, "J｜季度追蹤：實際數、指引與拆分設定（company.json → quarterly；實際數預設空白，財報公布後填入藍字格）")
    QS = QC['quarters']; NQ = len(QS); QCL = QCOL[:NQ]
    PN = QC['periodNames']; PERS = sorted({q['period'] for q in QS})
    DRV = (QC.get('driver') or {}).get('type') == 'mw'
    QMET = QC['metrics']
    _UF = {"US$bn": '#,##0.000', "MW": NUM0, "%": PCT}
    for j, h in enumerate(["季度", "單位"] + [q['label'] for q in QS]):
        c = ws.cell(row=r, column=1 + j, value=h); c.font = HEAD; c.fill = FILL_HEAD
    c = ws.cell(row=r, column=3 + max(NQ, 6) + 1, value="說明"); c.font = HEAD; c.fill = FILL_HEAD
    _JN = 3 + max(NQ, 6) + 1  # 說明欄（6 季時為 J 欄）
    r += 1
    JQ = {}

    def jrow(name, unit, vals, fmt=NUM, note=None, font=BLUE, key=False):
        global r
        JQ[name] = r
        ws.cell(row=r, column=1, value=name).font = BLACK
        if unit:
            ws.cell(row=r, column=2, value=unit).font = SMALL
        for j, v in enumerate(vals):
            c = ws.cell(row=r, column=3 + j, value=v); c.font = font; c.number_format = fmt; c.border = BOX
            if key:
                c.fill = FILL_KEY
        if note:
            ws.cell(row=r, column=_JN, value=note).font = SMALL
        r += 1

    _PERL = "季度｜所屬期間（" + "、".join(f"{i + 1}＝{n}" for i, n in enumerate(PN)) + "）"
    jrow(_PERL, "", [q['period'] + 1 for q in QS], NUM0, "由 company.json 設定；拆分公式依此建置（改動需重建）", BLACK)
    _FI = next((i for i, q in enumerate(QS) if q['key'] == QC.get('focus')), 0)
    jrow(f"季度｜焦點季（1–{NQ}）", "", [_FI + 1], NUM0, "「季度追蹤」分頁第一屏與「摘要」驗證點顯示的季度；財報公布後改為下一季", BLUE, key=True)
    FOCUS = f"'輸入與假設'!$C${JQ[f'季度｜焦點季（1–{NQ}）']}"
    if QC.get('revenueAnchor'):
        A_ = QC['revenueAnchor']
        jrow(f"季度｜營收起點（{A_['label']}）", "US$bn", [A_['value']], '#,##0.000', f"營收拆分 anchor 用：從此值逐季線性爬升 [{A_.get('tag', '')}]")
    if DRV:
        jrow("季度｜MW 起點（第 1 期期初 Accepted）", "MW", [QC['driver']['endStart']], NUM0, QC['driver'].get('endStartNote'))
    for m in QMET:
        if m['key'] == 'ebitdaMargin':
            continue
        jrow(f"季度實際｜{m.get('actualLabel', m['label'])}", m['unit'], [QC['actuals'].get(q['key'], {}).get(m['key']) for q in QS], _UF[m['unit']],
             "實際數（公司公布）；空白＝待公布" if m is QMET[0] else None)
    for k, nm in (("source", "來源"), ("date", "日期"), ("tag", "標記")):
        jrow(f"季度實際｜{nm}", "", [QC['actuals'].get(q['key'], {}).get(k) for q in QS], '@')
    GM = QC.get('guidanceMeta', {})
    for m in QMET:
        g = [QC.get('guidance', {}).get(q['key'], {}).get(m['key']) for q in QS]
        if any(g):
            jrow(f"季度指引｜{m['label']}（低）", m['unit'], [x[0] if x else None for x in g], _UF[m['unit']], f"{GM.get('source', '')} [{GM.get('tag', '')}]")
            jrow(f"季度指引｜{m['label']}（高）", m['unit'], [x[1] if x else None for x in g], _UF[m['unit']])
    for m in QMET:  # 文字指引（例如「調整後營業利益率 low teens」）：只列不計差距
        t = [QC.get('guidanceText', {}).get(q['key'], {}).get(m['key']) for q in QS]
        if any(t):
            jrow(f"季度文字指引｜{m['label']}", "", [x['text'] if x else None for x in t], '@',
                 "；".join(f"{q['label']}：{x.get('source', '')} [{x.get('tag', '')}]" for q, x in zip(QS, t) if x))
    for P, pg in (QC.get('periodGuidance') or {}).items():
        for mk, g in pg.items():
            m = next(x for x in QMET if x['key'] == mk)
            jrow(f"期間指引｜{PN[PERS.index(int(P))]} {m['label']}（全年低／全年高／扣除 1H 實際）", m['unit'], [g['range'][0], g['range'][1], g.get('less', 0)],
                 _UF[m['unit']], QC.get('periodGuidanceNote'))
    JR = lambda name, j: f"'輸入與假設'!{QCL[j]}${JQ[name]}"

    # ---------------- 「季度追蹤」分頁 ----------------
    ws = wb.create_sheet("季度追蹤")
    ws.column_dimensions["A"].width = 44
    ws.column_dimensions["B"].width = 8
    for c in QCOL + ["K", "L"]:
        ws.column_dimensions[c].width = 14
    ws.column_dimensions["M"].width = 60
    ws["A1"] = "季度追蹤 — 焦點季的模型／共識／指引／實際與差距（由年度模型拆分；只用於追蹤，不影響年度數字與目標價）"
    ws["A1"].font = TITLE
    ws["A2"] = f"設定與實際數在『輸入與假設』J 區，共識在 I 區；缺資料顯示「不適用」，實際數空白＝待公布。差距超過門檻的項目附差異原因（拆法＝所屬期間合計差距在門檻內）。第 2 區為 6 季路徑與拆分明細（預設收合）。"
    ws["A2"].font = SMALL
    NM_ = len(QMET)
    r = 4 + 1 + 1 + NM_ + 1 + 2  # 第 1 區（標題、表頭、各指標、焦點季、差異原因）之後
    r = section(ws, r, "2｜6 季路徑、共識、指引、實際與差距（明細；第 1 區以焦點季引用本區）", span=12, collapsed=True)
    for j, h in enumerate(["項目", "單位"] + [q['label'] for q in QS]):
        c = ws.cell(row=r, column=1 + j, value=h); c.font = HEAD; c.fill = FILL_HEAD
    HDR = r; r += 1
    _NOTE = 3 + NQ + 1

    def qrow(name, unit, fn, fmt=NUM, note=None, font=BLACK):
        global r
        QT[name] = r
        ws.cell(row=r, column=1, value=name).font = BLACK
        if unit:
            ws.cell(row=r, column=2, value=unit).font = SMALL
        for j in range(NQ):
            c = ws.cell(row=r, column=3 + j, value=fn(j, QCL[j])); c.font = font; c.number_format = fmt; c.border = BOX
        if note:
            ws.cell(row=r, column=_NOTE, value=note).font = SMALL
        r += 1
        return QT[name]

    POS_ = []  # 每季：(期間 P、期內序號 j、期內季數 k、同期欄)
    for i, q in enumerate(QS):
        same = [x for x in range(NQ) if QS[x]['period'] == q['period']]
        POS_.append((q['period'], same.index(i), len(same), [QCL[x] for x in same]))
    IA = lambda name, P: f"'輸入與假設'!{COLS[P]}${IN[name]}"
    ACC = lambda P: f"'輸入與假設'!{COLS[P]}${_acc}"
    BIL = lambda P: f"'輸入與假設'!{COLS[P]}${_bil}"
    ANN = {"revenue": lambda P: f"{PL}{COLS[P]}{rev_v}", "adjEbitda": lambda P: f"{PL}{COLS[P]}{ebitda_v}",
           "adjOpInc": lambda P: f"{PL}{COLS[P]}{ebit_v}", "capex": lambda P: IA("毛 CapEx", P)}
    qrow("所屬期間", "", lambda j, c: f"={JR(_PERL, j)}", NUM0, "1＝" + PN[0] + (f"、2＝{PN[1]}" if len(PN) > 1 else ""), GREEN)
    qrow("季序", "", lambda j, c: j + 1, NUM0, "EBITDA 率直線的 x 軸")
    MST = f"'輸入與假設'!$C${JQ['季度｜MW 起點（第 1 期期初 Accepted）']}" if DRV else None
    a0 = lambda P: MST if P == PERS[0] else ACC(P - 1)
    b0 = lambda P: MW0 if P == PERS[0] else BIL(P - 1)
    if DRV:
        qrow("季末 Accepted MW（模型）", "MW", lambda j, c: f"={a0(POS_[j][0])}+({ACC(POS_[j][0])}-{a0(POS_[j][0])})*{POS_[j][1] + 1}/{POS_[j][2]}", NUM0,
             "期初與期末 Accepted 之間線性內插；第 1 期起點見 J 區")
        qrow("平均在役 MW（Billable，拆分依據）", "MW", lambda j, c: f"={b0(POS_[j][0])}+({BIL(POS_[j][0])}-{b0(POS_[j][0])})*{2 * POS_[j][1] + 1}/{2 * POS_[j][2]}", NUM0,
             "期初與期末 Billable 之間線性內插的季平均")
        qrow("季內新增 Accepted MW", "MW", lambda j, c: (f"={c}{QT['季末 Accepted MW（模型）']}-{a0(POS_[j][0])}" if POS_[j][1] == 0
                                                     else f"={c}{QT['季末 Accepted MW（模型）']}-{QCL[j - 1]}{QT['季末 Accepted MW（模型）']}"), NUM0, "CapEx 拆分依據")
    else:
        qrow("季末 Accepted MW（模型）", "MW", lambda j, c: "不適用", NUM0, "company.json 未設定 MW 驅動")
        qrow("平均在役 MW（Billable，拆分依據）", "MW", lambda j, c: "不適用", NUM0)
        qrow("季內新增 Accepted MW", "MW", lambda j, c: 1, NUM0, "無 MW 驅動：CapEx 期內平分")
    _AV, _AD = QT['平均在役 MW（Billable，拆分依據）'], QT['季內新增 Accepted MW']
    _AN = f"'輸入與假設'!$C${JQ[[k for k in JQ if k.startswith('季度｜營收起點')][0]]}" if QC.get('revenueAnchor') else None

    def _rev(j, c):
        P, jj, k, cols = POS_[j]
        R, how = ANN["revenue"](P), (QC.get('revenueSplit') or [])[PERS.index(P)] if PERS.index(P) < len(QC.get('revenueSplit') or []) else "equal"
        if how == "anchor" and _AN:
            return f"={_AN}+{jj + 1}*({R}-{k}*{_AN})/{k * (k + 1) / 2:g}"
        if how == "driverAvg" and DRV:
            return f"={R}*{c}{_AV}/(" + "+".join(f"{x}{_AV}" for x in cols) + ")"
        return f"={R}/{k}"
    qrow("營收（模型）", "US$bn", _rev, NUM, "拆分方法見『輸入與假設』J 區與 HTML「拆分方法」")
    _RV_ = QT["營收（模型）"]
    # EBITDA 率直線：兩個期間時同時符合兩期年度 EBITDA（a＋b × 季序）；一個期間時為常數
    for nm, fn in (("EBITDA 率直線｜各期營收合計 S", lambda P, cols: "=" + "+".join(f"{x}{_RV_}" for x in cols)),
                   ("EBITDA 率直線｜營收 × 季序合計 T", lambda P, cols: "=" + "+".join(f"{x}{_RV_}*{x}{QT['季序']}" for x in cols)),
                   ("EBITDA 率直線｜年度 EBITDA E", lambda P, cols: f"={ANN['adjEbitda'](P)}")):
        QT[nm] = r
        ws.cell(row=r, column=1, value=nm).font = BLACK
        for pi, P in enumerate(PERS):
            cols = [QCL[x] for x in range(NQ) if QS[x]['period'] == P]
            c = ws.cell(row=r, column=3 + pi, value=fn(P, cols)); c.number_format = NUM; c.border = BOX
        ws.cell(row=r, column=_NOTE, value="欄＝期間（" + "、".join(PN) + "）").font = SMALL
        r += 1
    _S, _T, _E = QT["EBITDA 率直線｜各期營收合計 S"], QT["EBITDA 率直線｜營收 × 季序合計 T"], QT["EBITDA 率直線｜年度 EBITDA E"]
    if len(PERS) == 2:
        QT["EBITDA 率直線｜截距 a／斜率 b"] = r
        ws.cell(row=r, column=1, value="EBITDA 率直線｜截距 a／斜率 b").font = BLACK
        ws.cell(row=r, column=3, value=f"=(C{_E}*D{_T}-C{_T}*D{_E})/(C{_S}*D{_T}-C{_T}*D{_S})").number_format = '0.0000%'
        ws.cell(row=r, column=4, value=f"=(C{_S}*D{_E}-D{_S}*C{_E})/(C{_S}*D{_T}-C{_T}*D{_S})").number_format = '0.0000%'
        ws.cell(row=r, column=_NOTE, value="解兩式：Σ營收×(a＋b×季序)＝各期年度 EBITDA").font = SMALL
        _AB = r; r += 1
        qrow("EBITDA 率（模型）", "%", lambda j, c: f"=$C${_AB}+$D${_AB}*{c}{QT['季序']}", PCT, "跨季一條直線，季度 EBITDA 加總＝年度")
    else:
        qrow("EBITDA 率（模型）", "%", lambda j, c: f"={'CDEFGH'[PERS.index(POS_[j][0])]}{_E}/{'CDEFGH'[PERS.index(POS_[j][0])]}{_S}", PCT, "期內常數")
    qrow("調整後 EBITDA（模型）", "US$bn", lambda j, c: f"={c}{_RV_}*{c}{QT['EBITDA 率（模型）']}", NUM)
    qrow("車隊折舊（模型）", "US$bn", lambda j, c: (f"=({IA('期初毛 PP&E', POS_[j][0])}+{IA('成長型 CapEx（模型期）', POS_[j][0])}*{2 * POS_[j][1] + 1}/{2 * POS_[j][2]})"
                                                  f"/{LIFE}*{IA('模型期長度（年）', POS_[j][0])}/{POS_[j][2]}"), NUM, "依期內平均 PP&E（與年度車隊折舊公式相同）")
    qrow("調整後營業利益（模型）", "US$bn", lambda j, c: f"={c}{QT['調整後 EBITDA（模型）']}-{c}{QT['車隊折舊（模型）']}", NUM, "＝調整後 EBITDA − 車隊折舊（未扣 SBC）")
    # 指引（J 區；數字區間與文字）——CapEx 指引錨定需先有指引列
    CROW, CSEC, GLO, GHI, AROW, GTX = {}, {}, {}, {}, {}, {}
    for m in QMET:
        lo, hi = f"季度指引｜{m['label']}（低）", f"季度指引｜{m['label']}（高）"
        if lo in JQ:
            GLO[m['key']] = qrow(f"指引｜{m['label']}（低）", m['unit'], lambda j, c, n=lo: f'=IF(ISNUMBER({JR(n, j)}),{JR(n, j)},"不適用")', NUM)
            GHI[m['key']] = qrow(f"指引｜{m['label']}（高）", m['unit'], lambda j, c, n=hi: f'=IF(ISNUMBER({JR(n, j)}),{JR(n, j)},"不適用")', NUM)
        tn = f"季度文字指引｜{m['label']}"
        if tn in JQ:
            GTX[m['key']] = qrow(f"指引（文字）｜{m['label']}", "", lambda j, c, n=tn: f'=IF(ISTEXT({JR(n, j)}),{JR(n, j)},"")', '@', "文字指引：只列不計差距")
    MIDX = lambda k, c: f"(({c}{GLO[k]}+{c}{GHI[k]})/2)"
    GOKX = lambda k, c: f"ISNUMBER({c}{GLO[k]}),ISNUMBER({c}{GHI[k]})"
    if QC.get('capexSplit') == 'guidanceAnchor' and 'capex' in GLO:
        qrow("CapEx 指引錨定值", "US$bn", lambda j, c: f'=IF(AND({GOKX("capex", c)}),{MIDX("capex", c)},"")', NUM, "有季度指引的季取指引中點（capexSplit＝guidanceAnchor）")
        _ANC = QT["CapEx 指引錨定值"]
        qrow("CapEx 未錨定季的新增 MW", "MW", lambda j, c: f'=IF(ISNUMBER({c}{_ANC}),0,{c}{_AD})', NUM0, "期間餘數依此分配")
        _FR = QT["CapEx 未錨定季的新增 MW"]

        def _cx(j, c):
            P, jj, k, cols = POS_[j]
            sa, sf = "+".join(f"{x}{_ANC}" for x in cols), "+".join(f"{x}{_FR}" for x in cols)
            rem = f"({ANN['capex'](P)}-SUM({','.join(f'{x}{_ANC}' for x in cols)}))"
            return f"=IF(ISNUMBER({c}{_ANC}),{c}{_ANC},IF(({sf})>0,{rem}*{c}{_FR}/({sf}),{rem}/({k}-COUNT({','.join(f'{x}{_ANC}' for x in cols)}))))"
        qrow("CapEx（毛額）（模型）", "US$bn", _cx, NUM, "有季度指引的季＝指引中點；其餘季＝期間餘數依新增 Accepted MW 分配")
    else:
        qrow("CapEx（毛額）（模型）", "US$bn", lambda j, c: (f"=IF((" + "+".join(f"{x}{_AD}" for x in POS_[j][3]) + f")>0,{ANN['capex'](POS_[j][0])}*{c}{_AD}/("
                                                            + "+".join(f"{x}{_AD}" for x in POS_[j][3]) + f"),{ANN['capex'](POS_[j][0])}/{POS_[j][2]})"), NUM, "依季內新增 Accepted MW")
    MROW = {"revenue": _RV_, "adjEbitda": QT["調整後 EBITDA（模型）"], "ebitdaMargin": QT["EBITDA 率（模型）"], "adjOpInc": QT["調整後營業利益（模型）"],
            "capex": QT["CapEx（毛額）（模型）"], "mw": QT["季末 Accepted MW（模型）"]}
    # 共識（I 區；只有共識檔列出的季度，其餘「不適用」）
    QEK = [k for k in CONS.get('quarterlyEstimates', {}) if _re_q.match(r'^\d{4}Q\d$', k)]
    _CL = {"revenue": "季度｜營收", "ebitda": "季度｜EBITDA", "ebit": "季度｜EBIT", "netIncome": "季度｜淨利"}
    ciq = lambda fld, qk: CIR(_CL[fld], QEK.index(qk)) if qk in QEK and fld in _CL else None
    for m in QMET:
        if m.get('consensus') == 'derived':
            def _cd(j, c):
                e, v = ciq('ebitda', QS[j]['key']), ciq('revenue', QS[j]['key'])
                return f'=IF(AND(ISNUMBER({e}),ISNUMBER({v})),{e}/{v},"不適用")' if e and v else "不適用"
            CROW[m['key']] = qrow(f"共識｜{m['label']} [Derived]", m['unit'], _cd, PCT, "＝共識 EBITDA ÷ 共識營收")
        elif m.get('consensus'):
            CROW[m['key']] = qrow(f"共識｜{m['label']}", m['unit'], lambda j, c, f=m['consensus']: (f'=IF(ISNUMBER({ciq(f, QS[j]["key"])}),{ciq(f, QS[j]["key"])},"不適用")'
                                                                                              if ciq(f, QS[j]['key']) else "不適用"), NUM, "I 區季度共識 [Interested-party]")
        if m.get('consensusSecondary'):
            CSEC[m['key']] = qrow(m.get('consensusSecondaryLabel') or f"共識｜{m['consensusSecondary']}", m['unit'],
                                  lambda j, c, f=m['consensusSecondary']: (f'=IF(ISNUMBER({ciq(f, QS[j]["key"])}),{ciq(f, QS[j]["key"])},"不適用")' if ciq(f, QS[j]['key']) else "不適用"), NUM,
                                  "口徑未揭露：只列不比較")
    for m in QMET:
        al = f"實際｜{m.get('actualLabel', m['label'])}"
        if m['key'] == 'ebitdaMargin':
            e_, v_ = "季度實際｜" + next(x for x in QMET if x['key'] == 'adjEbitda').get('actualLabel', "調整後 EBITDA"), "季度實際｜" + next(x for x in QMET if x['key'] == 'revenue').get('actualLabel', "營收")
            AROW[m['key']] = qrow(al, m['unit'], lambda j, c: f'=IF(AND(ISNUMBER({JR(e_, j)}),ISNUMBER({JR(v_, j)})),{JR(e_, j)}/{JR(v_, j)},"")', PCT, "＝實際 EBITDA ÷ 實際營收")
        else:
            n_ = f"季度實際｜{m.get('actualLabel', m['label'])}"
            AROW[m['key']] = qrow(al, m['unit'], lambda j, c, n=n_: f'=IF(ISNUMBER({JR(n, j)}),{JR(n, j)},"")', PCT if m['unit'] == "%" else NUM, None, GREEN)
    # 差距（v4.4 第 4 輪）：營收、CapEx＝比例；利潤類＝金額差＋利潤率百分點；EBITDA 率＝百分點；MW＝差額
    TOLX = CONS_TOL
    KIND = {m['key']: m.get('gap') or ('pt' if m['unit'] == "%" else 'ratio') for m in QMET}
    _GFMT = lambda k, u: PCT if KIND[k] in ('ratio', 'pt') else (NUM0 if u == "MW" else NUM)
    GAP, GPT = {}, {}
    for m in QMET:
        k, u, kd = m['key'], m['unit'], KIND[m['key']]
        rel = (lambda x, y: f"{x}/({y})-1") if kd == 'ratio' else (lambda x, y: f"{x}-({y})")
        nz = (lambda y: f",{y}<>0") if kd == 'ratio' else (lambda y: "")
        M_ = lambda c: f"{c}{MROW[k]}"
        C_ = (lambda c: f"{c}{CROW[k]}") if k in CROW else None
        MID = (lambda c: MIDX(k, c)) if k in GLO else None
        A_ = lambda c: f"{c}{AROW[k]}"
        un = "%" if kd == 'ratio' else ("pt" if kd == 'pt' else u)
        G = {}
        G['mc'] = qrow(f"差距｜{m['label']}｜模型 vs 共識", un, lambda j, c: (f'=IF(AND(ISNUMBER({M_(c)}),ISNUMBER({C_(c)}){nz(C_(c))}),{rel(M_(c), C_(c))},"不適用")' if C_ else "不適用"), _GFMT(k, u))
        G['mg'] = qrow(f"差距｜{m['label']}｜模型 vs 指引中點", un, lambda j, c: (f'=IF(AND(ISNUMBER({M_(c)}),{GOKX(k, c)}),{rel(M_(c), MID(c))},"不適用")' if MID else "不適用"), _GFMT(k, u))
        G['am'] = qrow(f"差距｜{m['label']}｜實際 vs 模型", un, lambda j, c: f'=IF(ISNUMBER({A_(c)}),IF(AND(ISNUMBER({M_(c)}){nz(M_(c))}),{rel(A_(c), M_(c))},"不適用"),"")', _GFMT(k, u))
        G['ac'] = qrow(f"差距｜{m['label']}｜實際 vs 共識", un, lambda j, c: (f'=IF(ISNUMBER({A_(c)}),IF(AND(ISNUMBER({C_(c)}){nz(C_(c))}),{rel(A_(c), C_(c))},"不適用"),"")' if C_ else f'=IF(ISNUMBER({A_(c)}),"不適用","")'), _GFMT(k, u))
        G['ag'] = qrow(f"差距｜{m['label']}｜實際 vs 指引中點", un, lambda j, c: (f'=IF(ISNUMBER({A_(c)}),IF(AND({GOKX(k, c)}),{rel(A_(c), MID(c))},"不適用"),"")' if MID else f'=IF(ISNUMBER({A_(c)}),"不適用","")'), _GFMT(k, u))
        GAP[k] = G
        if m.get('marginOf'):  # 利潤率百分點（利潤 ÷ 同來源營收）
            rk = m['marginOf']
            MM = lambda c: (f"ISNUMBER({c}{MROW[k]}),ISNUMBER({c}{MROW[rk]}),{c}{MROW[rk]}<>0", f"{c}{MROW[k]}/{c}{MROW[rk]}")
            CM = (lambda c: (f"ISNUMBER({c}{CROW[k]}),ISNUMBER({c}{CROW[rk]}),{c}{CROW[rk]}<>0", f"{c}{CROW[k]}/{c}{CROW[rk]}")) if k in CROW and rk in CROW else None
            GMg = (lambda c: (f"{GOKX(k, c)},{GOKX(rk, c)}", f"{MIDX(k, c)}/{MIDX(rk, c)}")) if k in GLO and rk in GLO else None
            AM = lambda c: (f"ISNUMBER({c}{AROW[k]}),ISNUMBER({c}{AROW[rk]}),{c}{AROW[rk]}<>0", f"{c}{AROW[k]}/{c}{AROW[rk]}")

            def _pt(a, b, act):
                def f(j, c):
                    if a is None or b is None:
                        return f'=IF(ISNUMBER({c}{AROW[k]}),"不適用","")' if act else "不適用"
                    (ca, va), (cb, vb) = a(c), b(c)
                    core = f'IF(AND({ca},{cb}),{va}-{vb},"不適用")'
                    return f'=IF(ISNUMBER({c}{AROW[k]}),{core},"")' if act else "=" + core
                return f
            P_ = {}
            P_['mc'] = qrow(f"差距｜{m['label']}｜模型 vs 共識（利潤率 pt）", "pt", _pt(MM, CM, False), PCT)
            P_['mg'] = qrow(f"差距｜{m['label']}｜模型 vs 指引中點（利潤率 pt）", "pt", _pt(MM, GMg, False), PCT)
            P_['am'] = qrow(f"差距｜{m['label']}｜實際 vs 模型（利潤率 pt）", "pt", _pt(AM, MM, True), PCT)
            P_['ac'] = qrow(f"差距｜{m['label']}｜實際 vs 共識（利潤率 pt）", "pt", _pt(AM, CM, True), PCT)
            P_['ag'] = qrow(f"差距｜{m['label']}｜實際 vs 指引中點（利潤率 pt）", "pt", _pt(AM, GMg, True), PCT)
            GPT[k] = P_

    def _DTX(x, u, dp=2):  # 金額差文字（與 HTML dTxtQ 同格式）
        return f'IF({x}<0,"−","+")&' + _VX(f"ABS({x})", u, dp)

    def _GTX(k, key, c, u):  # 差距文字（與 HTML gTxtQ 同格式）
        g = f"{c}{GAP[k][key]}"
        if KIND[k] == 'diff':
            s = _DTX(g, u)
            if k in GPT:
                p_ = f"{c}{GPT[k][key]}"
                s += f'&IF(ISNUMBER({p_}),"／"&{_GX(p_, True)},"")'
            return s
        return _GX(g, KIND[k] == 'pt')

    def _EXC(k, g, b):  # 超過門檻：比例／百分點＝|差|>門檻；金額差＝|差|>門檻 × |比較基準|
        if KIND[k] == 'diff':
            return f"OR(AND({b}<>0,ABS({g})>{TOLX}*ABS({b})),AND({b}=0,{g}<>0))"
        return f"ABS({g})>{TOLX}"
    # 差異原因（已決定事項 2）：拆法＝所屬期間合計差距在門檻內（指引另須在隱含區間內）；模型落在指引區間外也須附原因；其餘讀 varianceReasons；找不到為「未歸類」
    VSN = {"consensus": "vs 共識", "guidance": "vs 指引中點", "actual": "實際 vs 模型"}
    RSN, RTY = {}, {}
    _PG = QC.get('periodGuidance') or {}
    CNM = {"consensus": "共識合計", "guidance": "指引隱含"}
    _PCX = lambda x: f'TEXT(({x})*100,"#,##0.0")&"%"'

    def _NEED(k, gk, c, g, b):  # 需附原因：利潤類（marginOf）＝利潤率差 > PM_TOL（pt 無值時沿用比較基準門檻）；其餘＝比較基準門檻
        base = _EXC(k, g, b)
        if k in GPT and PM_TOL:
            p_ = f"{c}{GPT[k][gk]}"
            return f"IF(ISNUMBER({p_}),ABS({p_})>{PM_TOL},{base})"
        return base
    for m in QMET:
        k, u = m['key'], m['unit']
        if KIND[k] == 'pt':
            continue
        # 期間合計的比較基準與差距拆解（水準＝(模型期間 − 比較期間) × 比較季占比；分配＝(模型季占比 − 比較季占比) × 模型期間）
        DEC = {}
        for vs in ("consensus", "guidance"):
            if k not in ANN:
                continue
            if vs == "consensus" and k in CROW:
                bfn = lambda j, c: (f'=IF(COUNT(' + ",".join(f"{x}{CROW[k]}" for x in POS_[j][3]) + f')={POS_[j][2]},' + "+".join(f"{x}{CROW[k]}" for x in POS_[j][3]) + ',"不適用")')
                cq = lambda c: f"{c}{CROW[k]}"
            elif vs == "guidance" and k in GLO and any(k in v for v in _PG.values()):
                def bfn(j, c):
                    P = POS_[j][0]
                    if not _PG.get(str(P), {}).get(k):
                        return "不適用"
                    n = f"期間指引｜{PN[PERS.index(P)]} {m['label']}（全年低／全年高／扣除 1H 實際）"
                    return f"=('輸入與假設'!$C${JQ[n]}+'輸入與假設'!$D${JQ[n]})/2-'輸入與假設'!$E${JQ[n]}"
                cq = lambda c: MIDX(k, c)
            else:
                continue
            B = qrow(f"合計基準｜{m['label']}｜vs {'共識' if vs == 'consensus' else '指引'}", u, bfn, NUM,
                     "所屬期間的季度共識合計" if vs == "consensus" else "所屬期間的隱含指引中點（全年指引中點 − 1H 實際）")
            ok = lambda c, B=B, cq=cq: f"ISNUMBER({c}{B}),{c}{B}<>0,ISNUMBER({cq(c)}),ISNUMBER({c}{MROW[k]})"
            L = qrow(f"拆解｜{m['label']}｜vs {'共識' if vs == 'consensus' else '指引'}｜水準", u, lambda j, c, B=B, cq=cq, ok=ok: (
                f'=IF(AND({ok(c)}),({ANN[k](POS_[j][0])}-{c}{B})*({cq(c)})/{c}{B},"不適用")'), NUM, "(模型期間 − 比較期間) × 比較的季占比")
            A_ = qrow(f"拆解｜{m['label']}｜vs {'共識' if vs == 'consensus' else '指引'}｜分配", u, lambda j, c, B=B, cq=cq, ok=ok: (
                f'=IF(AND({ok(c)}),({c}{MROW[k]}/{ANN[k](POS_[j][0])}-({cq(c)})/{c}{B})*{ANN[k](POS_[j][0])},"不適用")'), NUM, "(模型季占比 − 比較季占比) × 模型期間；水準＋分配＝季度差距")
            DEC[vs] = (B, L, A_, cq)
        for vs, gk in (("consensus", "mc"), ("guidance", "mg"), ("actual", "am")):
            def _rs(j, c, typ_only=False, vs=vs, gk=gk, k=k, u=u, m=m, DEC=DEC):
                q = QS[j]; P = POS_[j][0]; pn = PN[PERS.index(P)]; G = f"{c}{GAP[k][gk]}"
                cfg = _find_rsn('quarter', q['key'], k, vs)
                head = f'"{m["label"]} {VSN[vs]} "&{_GTX(k, gk, c, u)}&"：'

                def ctx(path):  # 原因文字的 {路徑}：q.＝本季模型值、c.＝本季共識（I 區）、g.＝本季指引中點、lq.＝latestQuarter（常數）
                    if path == 'q.da':
                        return f"{c}{QT['車隊折舊（模型）']}"
                    if path == 'q.mwAvg':
                        return f"{c}{_AV}"
                    if path.startswith('q.') and path[2:] in MROW:
                        return f"{c}{MROW[path[2:]]}"
                    if path.startswith('c.'):
                        return ciq(path[2:], q['key'])  # 該季沒有共識＝None
                    if path.startswith('g.'):
                        return MIDX(path[2:], c) if path[2:] in GLO and QC.get('guidance', {}).get(q['key'], {}).get(path[2:]) else None
                    if path.startswith('lq.'):
                        return repr(float(CO['latestQuarter'][path[3:]]))
                    raise KeyError(f"varianceReasons 路徑無法對應 Excel 儲存格：{path}")
                base = {"consensus": f"{c}{CROW[k]}" if k in CROW else None, "guidance": MIDX(k, c) if k in GLO else None, "actual": f"{c}{MROW[k]}"}[vs]
                if base is None:
                    return ""
                trig = _NEED(k, gk, c, G, base)
                if vs == "guidance":
                    trig = f"OR({trig},{c}{MROW[k]}>{c}{GHI[k]},{c}{MROW[k]}<{c}{GLO[k]})"
                if vs == "actual":
                    if not cfg:
                        return ""
                    body = f'"{cfg["type"]}"' if typ_only else f'{head}{cfg["type"]}（"&{_TOKX(cfg["text"], ctx)}&"）"'
                    return f'=IF(ISNUMBER({G}),IF({trig},{body},""),"")'
                other = (f'"{cfg["type"]}"' if typ_only else f'{head}{cfg["type"]}（"&{_TOKX(cfg["text"], ctx)}&"）"') if cfg else ('"未歸類"' if typ_only else f'{head}未歸類（差異原因待補）"')
                inner = other
                if vs in DEC:
                    B, L, A_, cq = DEC[vs]
                    Lc, Ac, Bc, AP = f"{c}{L}", f"{c}{A_}", f"{c}{B}", ANN[k](P)
                    mS, cS = f"{c}{MROW[k]}/{AP}", f"({cq(c)})/{Bc}"
                    split = '"拆法"' if typ_only else (f'{head}拆法（分配 "&{_DTX(Ac, u, 3)}&"（{q["label"]} 占 {pn} "&{_PCX(mS)}&"，{CNM[vs]} "&{_PCX(cS)}&"）；水準 "&{_DTX(Lc, u, 3)}&"）"')
                    vtype = cfg["type"] if cfg else "觀點"
                    view = f'"{vtype}"' if typ_only else (f'{head}{vtype}（水準 "&{_DTX(Lc, u, 3)}&"（{pn} 模型 "&{_VX(AP, u, 3)}&"，{CNM[vs]} "&{_VX(Bc, u, 3)}&"）；分配 "&{_DTX(Ac, u, 3)}'
                                                           + (f'&"；"&{_TOKX(cfg["text"], ctx)}' if cfg else '') + '&"）"')
                    inner = f'IF(ISNUMBER({Lc}),IF(ABS({Ac})>ABS({Lc}),{split},{view}),{other})'
                return f'=IF(ISNUMBER({G}),IF({trig},{inner},""),"")'
            RSN[(k, vs)] = qrow(f"原因｜{m['label']}｜{VSN[vs]}", "", _rs, '@')
            RTY[(k, vs)] = qrow(f"原因類型｜{m['label']}｜{VSN[vs]}", "", lambda j, c, f=_rs: f(j, c, True), '@')
    _rl = [RSN[x] for m in QMET for x in [(m['key'], v) for v in ("consensus", "guidance", "actual")] if x in RSN]
    qrow("差異原因（一行）", "", lambda j, c: "=MID(" + "&".join(f'IF({c}{x}<>"","；"&{c}{x},"")' for x in _rl) + ",2,9999)", '@')
    TAG = {}
    for m in QMET:
        ts = [RTY[(m['key'], v)] for v in ("consensus", "guidance", "actual") if (m['key'], v) in RTY]
        if not ts:
            continue
        def _tag(j, c, ts=ts):
            t = [f"{c}{x}" for x in ts]
            s = t[0]
            for i in range(1, len(t)):
                prev = t[:i]
                s += f'&IF(AND({t[i]}<>"",' + ",".join(f"{t[i]}<>{p}" for p in prev) + f'),IF(({"&".join(prev)})<>"","、","")&{t[i]},"")'
            return f'=IF(({s})<>"","〔"&({s})&"〕","")'
        TAG[m['key']] = qrow(f"原因標籤｜{m['label']}", "", _tag, '@')
    # 驗證點句（與 HTML quarterlyView.line 同格式）
    LINE = {}
    for m in QMET:
        k, u = m['key'], m['unit']

        def _ln(j, c, k=k, u=u, m=m):
            M = f"{c}{MROW[k]}"
            gx = lambda key: f"{c}{GAP[k][key]}"
            gtx = lambda key: f'IF(ISNUMBER({gx(key)}),{_GTX(k, key, c, u)},"不適用")'
            s = f'="{m["label"]}：模型 "&IF(ISNUMBER({M}),{_VX(M, u)},"不適用")'
            if k in CROW:
                C_ = f"{c}{CROW[k]}"
                s += f'&IF(ISNUMBER({C_}),"｜共識 "&{_VX(C_, u)}&"（"&{gtx("mc")}&"）","")'
            txt = f'IF({c}{GTX[k]}<>"","｜指引（文字）"&{c}{GTX[k]},"")' if k in GTX else '""'
            if k in GLO:
                lo, hi = f"{c}{GLO[k]}", f"{c}{GHI[k]}"
                s += f'&IF(AND(ISNUMBER({lo}),ISNUMBER({hi})),"｜指引 "&{_RX(lo, hi, u)}&"（中點 "&{gtx("mg")}&"，"&{_POS(M, lo, hi)}&"）",{txt})'
            elif k in GTX:
                s += f'&{txt}'
            A = f"{c}{AROW[k]}"
            act = f'"｜實際 "&{_VX(A, u)}&"（較模型 "&{gtx("am")}'
            if k in CROW:
                C_ = f"{c}{CROW[k]}"
                act += f'&IF(ISNUMBER({C_}),"、較共識 "&{gtx("ac")},"")'
            if k in GLO:
                lo, hi = f"{c}{GLO[k]}", f"{c}{GHI[k]}"
                act += f'&IF(AND(ISNUMBER({lo}),ISNUMBER({hi})),"、"&{_POS(A, lo, hi)},"")'
            s += f'&IF(ISNUMBER({A}),{act}&"）","")'
            if k in TAG:
                s += f'&{c}{TAG[k]}'
            return s
        LINE[k] = qrow(f"句｜{m['label']}", "", _ln, '@')
    # 期間合計核對（季度加總＝年度模型；應全為 0）
    r += 1
    ws.cell(row=r, column=1, value="核對｜季度加總 − 年度模型（應為 0；欄＝期間 " + "、".join(PN) + "）").font = BOLD; r += 1
    for nm, row_, annf in (("營收", _RV_, ANN["revenue"]), ("調整後 EBITDA", QT["調整後 EBITDA（模型）"], ANN["adjEbitda"]),
                           ("調整後營業利益", QT["調整後營業利益（模型）"], ANN["adjOpInc"]), ("CapEx（毛額）", QT["CapEx（毛額）（模型）"], ANN["capex"]),
                           ("車隊折舊", QT["車隊折舊（模型）"], lambda P: IA("D&A（車隊）", P))):
        QT[f"核對｜{nm}"] = r
        ws.cell(row=r, column=1, value=f"核對｜{nm}").font = BLACK
        for pi, P in enumerate(PERS):
            cols = [QCL[x] for x in range(NQ) if QS[x]['period'] == P]
            c = ws.cell(row=r, column=3 + pi, value="=" + "+".join(f"{x}{row_}" for x in cols) + f"-{annf(P)}"); c.number_format = '0.000000'; c.border = BOX
        r += 1
    # ---------------- 第 1 區：焦點季（第一屏） ----------------
    r1 = 4
    section(ws, r1, "1｜焦點季（『輸入與假設』J 區選擇）：模型／共識／指引／實際與差距", span=17)
    FH = r1 + 1
    for j, h in enumerate(["指標", "單位", "模型", "共識", "模型 vs 共識", "指引（低）", "指引（高）", "模型 vs 指引中點", "實際", "實際 vs 模型", "實際 vs 共識", "實際 vs 指引中點",
                           "利潤率 pt：模型 vs 共識", "模型 vs 指引", "實際 vs 模型", "實際 vs 共識", "實際 vs 指引"]):
        c = ws.cell(row=FH, column=1 + j, value=h); c.font = HEAD; c.fill = FILL_HEAD
    IX = lambda row_: f"INDEX($C${row_}:${QCL[-1]}${row_},{FOCUS})"
    for i, m in enumerate(QMET):
        rr = FH + 1 + i; k, u = m['key'], m['unit']
        QT[f"焦點｜{m['label']}"] = rr
        ws.cell(row=rr, column=1, value=f"焦點｜{m['label']}").font = BLACK
        ws.cell(row=rr, column=2, value=u).font = SMALL
        fmt = PCT if u == "%" else (NUM0 if u == "MW" else NUM)
        _lo = (f'=IF(ISNUMBER({IX(GLO[k])}),{IX(GLO[k])},' + (f'IF({IX(GTX[k])}<>"",{IX(GTX[k])},"不適用"))' if k in GTX else '"不適用")')) if k in GLO else (
            f'=IF({IX(GTX[k])}<>"",{IX(GTX[k])},"不適用")' if k in GTX else "不適用")
        vals = [f"={IX(MROW[k])}", f"={IX(CROW[k])}" if k in CROW else "不適用", f"={IX(GAP[k]['mc'])}",
                _lo, f"={IX(GHI[k])}" if k in GLO else "不適用", f"={IX(GAP[k]['mg'])}",
                f'=IF(ISNUMBER({IX(AROW[k])}),{IX(AROW[k])},"待公布")', f"={IX(GAP[k]['am'])}", f"={IX(GAP[k]['ac'])}", f"={IX(GAP[k]['ag'])}"]
        vals += [f"={IX(GPT[k][x])}" if k in GPT else "" for x in ("mc", "mg", "am", "ac", "ag")]
        gf = _GFMT(k, u)
        for j, v in enumerate(vals):
            c = ws.cell(row=rr, column=3 + j, value=v); c.border = BOX
            c.number_format = gf if j in (2, 5, 7, 8, 9) else (PCT if j >= 10 else fmt)
    rr = FH + 1 + NM_
    QT["焦點｜季度"] = rr
    ws.cell(row=rr, column=1, value="焦點｜季度").font = BOLD
    ws.cell(row=rr, column=3, value=f"=INDEX($C${HDR}:${QCL[-1]}${HDR},{FOCUS})").font = BOLD
    QT["焦點｜差異原因"] = rr + 1
    ws.cell(row=rr + 1, column=1, value="焦點｜差異原因").font = BOLD
    ws.cell(row=rr + 1, column=3, value=f'=IF({IX(QT["差異原因（一行）"])}="","無（差距皆在 "&{_PC(CONS_TOL)}&" 以內）",{IX(QT["差異原因（一行）"])})')
    ws.freeze_panes = "C6"
    ws = wb["輸入與假設"]

# =====================================================================
# v4.3：摘要（一頁摘要；放在「導覽」之後）。全部活公式，與 HTML 總結頁第 1 頁同一組數字與字串（cmp31 逐項比對）
# 模型端為目前情境（預設＝基準）；FY26 為全年口徑：營收＝1H 實際＋2H 模型，調整後 EBITDA＝1H 實際調整後 EBITDA＋2H 模型，CapEx＝1H 實際毛額＋2H 模型毛額
# =====================================================================
ws = wb.create_sheet("摘要")
ws.column_dimensions["A"].width = 40
ws.column_dimensions["B"].width = 10
for c in "CDEFG":
    ws.column_dimensions[c].width = 13
ws.column_dimensions["I"].width = 70
ws["A1"] = "一頁摘要 — 結論、與市場的差異、現價隱含什麼、驗證點（全部依輸入連動）"
ws["A1"].font = TITLE
ws["A2"] = f'="模型端＝目前情境：" & CHOOSE({SEL},"{DT_NM["low"]}","{DT_NM["base"]}","{DT_NM["high"]}") & "（預設＝基準）；共識＝{CO["meta"]["consensusFile"]}（{CONS["asOf"]}）"'
ws["A2"].font = SMALL
VQ = "'評價_DCF與目標價'!"
SM = {}
def srow(name, unit, vals, fmt=NUM, note=None, bold=False):
    global r
    SM[name] = r
    ws.cell(row=r, column=1, value=name).font = BOLD if bold else BLACK
    if unit: ws.cell(row=r, column=2, value=unit).font = SMALL
    for j, v in enumerate(vals):
        c = ws.cell(row=r, column=3 + j, value=v); c.number_format = fmt; c.font = BLACK
    if note: ws.cell(row=r, column=9, value=note).font = SMALL
    r += 1
_G = lambda x, pt=False: f'IF({x}<0,"−","+")&TEXT(ABS({x})*100,"0.0")&"{"pt" if pt else "%"}"'  # 與 HTML gapTxtQ 同格式
r = 4
r = section(ws, r, "1｜結論")
srow("結論｜評等代碼（1＝買進、0＝中立、−1＝賣出）", "", [f"={VQ}{RG_PT_CODE}"], NUM0)
srow("結論｜評等", "", [f"=CHOOSE(C{SM['結論｜評等代碼（1＝買進、0＝中立、−1＝賣出）']}+2,\"賣出\",\"中立\",\"買進\")"])
srow("結論｜點位（加權目標價）", "US$", [f"={VQ}{TGT}"], USD, bold=True)
srow("結論｜現價", "US$", [f"={PX}"], USD)
srow("結論｜空間", "%", [f"=C{r-2}/C{r-1}-1"], PCT)
srow("結論｜情境區間（下緣／上緣）", "US$", [f"={VQ}{_A0}", f"={VQ}{_A1}"], USD)
srow("結論｜方法區間（下緣／上緣）", "US$", [f"={VQ}{_B0}", f"={VQ}{_B1}"], USD)
srow("結論｜賣出門檻價", "US$", [f"={VQ}{_TH}"], USD)
srow("結論｜點位 − 賣出門檻", "US$", [f"=C{SM['結論｜點位（加權目標價）']}-C{r-1}"], USD)
_s = lambda k: SM["結論｜" + k]
srow("結論｜結論句", "", [(f'=C{_s("評等")}&"：點位 $"&TEXT(C{_s("點位（加權目標價）")},"0.0")&"，較現價 $"&TEXT(C{_s("現價")},"0.00")&" "'
                         f'&IF(C{_s("空間")}>=0,"高","低")&" "&TEXT(ABS(C{_s("空間")})*100,"0")&"%；情境區間 $"&TEXT(C{_s("情境區間（下緣／上緣）")},"0.0")&"–$"&TEXT(D{_s("情境區間（下緣／上緣）")},"0.0")'
                         f'&"，方法區間 $"&TEXT(C{_s("方法區間（下緣／上緣）")},"0.0")&"–$"&TEXT(D{_s("方法區間（下緣／上緣）")},"0.0")'
                         f'&"；點位"&IF(C{_s("點位 − 賣出門檻")}<0,"低於","高於")&"賣出門檻 $"&TEXT(C{_s("賣出門檻價")},"0.0")&" 達 $"&TEXT(ABS(C{_s("點位 − 賣出門檻")}),"0.0")&"。"')],
     bold=True)
srow("結論｜情境判斷句", "", [f"={VQ}C{TRROW['目標價區間｜判斷句']}"])
# W3（v4.6）：每 MW 一句（首個完整財年；與 HTML 一頁摘要 pmLineQ 同一句）。只在每 MW 使用新方法時建列；公式在「每MW經濟性」頁建好後回填
_PMW_ON = PMW['capex'] != 'legacy' or PMW['cost'] != 'ebitdaPct' or PMW['revenue'] != 'legacy'
if FL and _PMW_ON:
    srow("結論｜每 MW 經濟性句", "", [None], NUM, "首個完整財年、目前情境；數字取自『每MW經濟性』彙總表與每 MW 收入對照（持有成本＝Tokenomics IF_HoldEcon 世代加權）")

r += 1
r = section(ws, r, "2｜與市場的差異（模型：目前情境 vs 共識；FY26–FY28）")
for j, h in enumerate(["項目", "單位", "FY26", "FY27", "FY28"]):
    c = ws.cell(row=r, column=1 + j, value=h); c.font = HEAD; c.fill = FILL_HEAD
r += 1
_NB = "'資產負債_新債與新股'!"
_MOD = {
    "營收": [f"={PL}{COLS[i]}{fy_rev}" for i in range(3)],
    "調整後 EBITDA": [f"={H_ADJEB}+{PL}C{ebitda_v}"] + [f"={PL}{COLS[i]}{VR['全年 EBITDA']}" for i in (1, 2)],
    "CapEx（毛額）": [f"={H_CAPEX}+'輸入與假設'!C${IN['毛 CapEx']}"] + [f"='輸入與假設'!{COLS[i]}${IN['毛 CapEx']}" for i in (1, 2)],
    "淨負債": [f"={_NB}{COLS[i]}{BR['淨負債（總債務 − 期末現金）']}" for i in range(3)],
}
_CON = {"營收": "年度｜營收", "調整後 EBITDA": "年度｜調整後 EBITDA", "CapEx（毛額）": "年度｜CapEx", "淨負債": "年度｜淨負債"}
for k in ["營收", "調整後 EBITDA", "EBITDA 率", "CapEx（毛額）", "淨負債"]:
    if k == "EBITDA 率":
        srow("差異｜EBITDA 率｜模型", "%", [f"={c}{SM['差異｜調整後 EBITDA｜模型']}/{c}{SM['差異｜營收｜模型']}" for c in "CDE"], PCT)
        srow("差異｜EBITDA 率｜共識", "%", [f"={CIR('年度｜EBITDA 率', j)}" for j in range(3)], PCT)
        srow("差異｜EBITDA 率｜差距（百分點）", "pt", [f"={c}{r-2}-{c}{r-1}" for c in "CDE"], PCT, "EBITDA 率以百分點列差距，不納入判斷")
        continue
    srow(f"差異｜{k}｜模型", "US$bn", _MOD[k], NUM)
    srow(f"差異｜{k}｜共識", "US$bn", [f"={CIR(_CON[k], j)}" for j in range(3)], NUM)
    srow(f"差異｜{k}｜差距", "%", [f"={c}{r-2}/{c}{r-1}-1" for c in "CDE"], PCT,
         "共識淨負債不含租賃、口徑未揭露：只列不判斷" if k == "淨負債" else ("共識 CapEx 口徑未揭露；«P0»＝«YTD» 實際毛額＋«STUB» 模型毛額" if k.startswith("CapEx") else
         ("«P0» 模型＝«YTD» 實際調整後 EBITDA（輸入與假設 G 區，報導轉述 [Verified]）＋«STUB» 模型" if k == "調整後 EBITDA" else None)))
    if k in ("調整後 EBITDA", "淨負債"):  # v4.4：利潤類與淨負債以金額差呈現（比例列保留作判斷門檻）
        srow(f"差異｜{k}｜差距（金額）", "US$bn", [f"={c}{r-3}-{c}{r-2}" for c in "CDE"], NUM, "與 HTML 一頁摘要、市場共識分頁同一呈現；判斷門檻仍用上一列比例")
_gr = {"營收": SM["差異｜營收｜差距"], "EBITDA": SM["差異｜調整後 EBITDA｜差距"], "CapEx": SM["差異｜CapEx（毛額）｜差距"]}
_DX = lambda x: f'IF({x}<0,"−","+")&"$"&TEXT(ABS({x}),"#,##0.00")&"bn"'  # 與 HTML dTxtQ 同格式
_GT = lambda n, c: (f'{_DX(c + str(SM["差異｜調整後 EBITDA｜差距（金額）"]))}&"／"&{_G(c + str(SM["差異｜EBITDA 率｜差距（百分點）"]), True)}'
                    if n == "EBITDA" else _G(f"{c}{_gr[n]}"))  # EBITDA：金額差／EBITDA 率百分點
srow("差異｜分歧項目（超過門檻者）", "", [("=MID(" + "&".join(f'IF(ABS({c}{_gr[n]})>{CONS_TOL},"、{n} "&{_GT(n, c)},"")' for n in ("營收", "EBITDA", "CapEx")) + ",2,999)") for c in "CDE"],
     NUM, "營收、EBITDA、CapEx 任一項差距絕對值超過門檻即列出（與 HTML 判斷句同一規則）")
_L = SM["差異｜分歧項目（超過門檻者）"]
srow("差異｜分歧起始年（1＝FY26…3＝FY28；0＝無）", "", [f'=IF(C{_L}<>"",1,IF(D{_L}<>"",2,IF(E{_L}<>"",3,0)))'], NUM0)
_F, _T = f"C{r-1}", _PC(CONS_TOL)
_nx = lambda c, y: f'"；{y} "&IF({c}{_L}="","回到 "&{_T}&" 以內",{c}{_L})'
srow("差異｜判斷句", "", [(f'=CHOOSE({_F}+1,"FY26–FY28 營收、EBITDA、CapEx 與共識差距皆在 "&{_T}&" 以內。",'
                          f'"分歧始於 FY26（"&C{_L}&"）"&{_nx("D", "FY27")}&{_nx("E", "FY28")}&"。",'
                          f'"FY26 營收、EBITDA、CapEx 與共識差距在 "&{_T}&" 以內；分歧始於 FY27（"&D{_L}&"）"&{_nx("E", "FY28")}&"。",'
                          f'"FY26–FY27 營收、EBITDA、CapEx 與共識差距在 "&{_T}&" 以內；分歧始於 FY28（"&E{_L}&"）。")')], bold=True)
# v4.4：差距超過門檻的項目附差異原因（已決定事項 2；原因讀 company.json → varianceReasons，與 HTML consensusView 同一組字串）
_ANK = [("營收", "rev", "營收"), ("EBITDA", "ebitda", "調整後 EBITDA"), ("CapEx", "capex", "CapEx（毛額）"), ("淨負債", "nd", "淨負債")]
_SMK = {"rev": "營收", "ebitda": "調整後 EBITDA", "ebM": "EBITDA 率", "capex": "CapEx（毛額）", "nd": "淨負債"}
_YRS = ["FY26", "FY27", "FY28"]
_RT = []
for _yi, _yr in enumerate(_YRS):
    _c = "CDE"[_yi]
    for _n, _k, _full in _ANK:
        _cfg = _find_rsn('annual', _yr, _k, 'consensus')
        if not _cfg:
            continue
        _g = f"{_c}{SM[f'差異｜{_full}｜差距']}"
        def _actx(path, _c=_c, _yi=_yi):  # 原因文字的 {路徑}：m／c／gap＝本頁差異列、y＝年度引擎列、in＝輸入格
            a, b = path.split('.', 1)
            if a in ('m', 'c'):
                return f"{_c}{SM['差異｜' + _SMK[b] + '｜' + ('模型' if a == 'm' else '共識')]}"
            if a == 'gap':
                return f"{_c}{SM['差異｜EBITDA 率｜差距（百分點）' if b == 'ebM' else '差異｜' + _SMK[b] + '｜差距']}"
            if a == 'y':
                return {"mwNew": f"'輸入與假設'!{COLS[_yi]}${IN['本期新增 MW']}", "accepted": f"'輸入與假設'!{COLS[_yi]}${_acc}",
                        "costMW": f"'輸入與假設'!{COLS[_yi]}${IN['每 MW 建置成本']}"}[b]
            if a == 'in':
                return {"ebStart": EB0, "ebSteady": EBSS}[b]
            raise KeyError(f"varianceReasons 路徑無法對應 Excel 儲存格：{path}")
        _cond = (f"ABS({_c}{SM['差異｜EBITDA 率｜差距（百分點）']})>{PM_TOL}" if _k == "ebitda" and PM_TOL else f"ABS({_g})>{CONS_TOL}")  # EBITDA：利潤率差門檻
        srow(f"差異原因｜{_yr}｜{_n}", "", [f'=IF({_cond},"{_cfg["type"]}","")', f'=IF({_cond},{_TOKX(_cfg["text"], _actx)},"")'], NUM,
             "C＝類型、D＝原因（需附原因時顯示；EBITDA 以利潤率差、其餘以比例門檻判斷）" if not _RT else None)
        _RT.append(SM[f"差異原因｜{_yr}｜{_n}"])
_GR4 = [SM[f"差異｜{_full}｜差距"] for _n, _k, _full in _ANK]
_GR4 = [SM["差異｜EBITDA 率｜差距（百分點）"] if (_k == "ebitda" and PM_TOL) else SM[f"差異｜{_full}｜差距"] for _n, _k, _full in _ANK]
_TL4 = [PM_TOL if (_k == "ebitda" and PM_TOL) else CONS_TOL for _n, _k, _full in _ANK]
srow("差異｜需附原因項數", "", ["=" + "+".join(f"SUMPRODUCT(--(ABS(C{g}:E{g})>{t}))" for g, t in zip(_GR4, _TL4))], NUM0, "營收、EBITDA（利潤率差）、CapEx、淨負債 × FY26–FY28")
_NX = f"C{SM['差異｜需附原因項數']}"
_cnt = lambda t: ("+".join(f'COUNTIF(C{x},"{t}")' for x in _RT) or "0")
_parts = "&".join(f'IF(({_cnt(t)})>0,"、{t} "&({_cnt(t)})&" 項","")' for t in RSN_TYPES[:-1])
_unx = f"({_NX}-(" + "+".join(f"({_cnt(t)})" for t in RSN_TYPES[:-1]) + "))"
srow("差異｜差異原因摘要", "", [f'=IF({_NX}=0,"","需附原因的 "&{_NX}&" 項差距："&MID({_parts}&IF({_unx}>0,"、未歸類 "&{_unx}&" 項",""),2,999)&"；原因見「損益與評價 → 市場共識」。")'],
     NUM, "與 HTML 一頁摘要同一句")

r += 1
r = section(ws, r, "3｜現價隱含什麼")
_E28, _N28, _PTM = CIR("年度｜調整後 EBITDA", 2), CIR("年度｜淨負債", 2), CIR("目標價｜平均")
srow("隱含｜共識平均目標價隱含 FY28 EV/EBITDA", "x", [f"=({_PTM}*{SH}+{_N28})/{_E28}"], MULT, "＝（共識平均目標價 × 股數＋共識 FY28 淨負債）÷ 共識 FY28 調整後 EBITDA")
srow("隱含｜現價隱含 FY28 EV/EBITDA", "x", [f"=({PX}*{SH}+{_N28})/{_E28}"], MULT, "同一共識 FY28 數字，價格換成現價")
srow("隱含｜模型方法區間上緣", "x", [f"=MAX({RM_LO},{RM_HI})"], MULT, "模型 EV/EBITDA 腿錨定 FY29 並折回 «TGT»，錨定年度與此不同")
_i1, _i2, _mh = f"C{SM['隱含｜共識平均目標價隱含 FY28 EV/EBITDA']}", f"C{SM['隱含｜現價隱含 FY28 EV/EBITDA']}", f"C{SM['隱含｜模型方法區間上緣']}"
srow("隱含｜隱含倍數句", "", [(f'="共識平均目標價 $"&TEXT({_PTM},"0.00")&" 隱含 FY28 EV/EBITDA "&TEXT({_i1},"0.0")&"x，"'
                            f'&IF(ROUND({_i1},1)>{_mh},"高於",IF(ROUND({_i1},1)<{_mh},"低於","等於"))&"模型方法區間上緣 "&{_MT(_mh)}&"x；現價 $"&TEXT({PX},"0.00")&" 隱含 "&TEXT({_i2},"0.0")&"x。"')], bold=True)
_RV = "'評價_反向DCF'!"
for k, nm in enumerate(["反向 DCF｜FY30 每 MW 年收入（目前／隱含／變動）", "反向 DCF｜每 MW 建置成本（目前／隱含／變動）",
                        "反向 DCF｜穩態 EBITDA 率（目前／隱含／變動 pt）", "反向 DCF｜加權目標價＝現價所需每 MW 年收入（目前／隱含／變動）"]):
    srow(nm, "", [f"={_RV}B{RV_R0 + k}", f"={_RV}C{RV_R0 + k}", f"={_RV}D{RV_R0 + k}"], PCT if k == 2 else NUM1,
         "快照（預設輸入）：見『評價_反向DCF』；改輸入後請以目標搜尋重算" if k == 0 else None)
    ws.cell(row=r - 1, column=5).number_format = PCT

r += 1
r = section(ws, r, "4｜驗證點與市場看法（賣方與管理層數字標記 Interested-party）")
if QC:  # v4.4：驗證點改為引用「季度追蹤」焦點季（company.json → quarterly.keyMetrics）
    _QTS = "'季度追蹤'!"
    srow("驗證｜焦點季", "", [f"={_QTS}C{QT['焦點｜季度']}"], NUM, "『輸入與假設』J 區選擇；季度層由年度模型拆分，只用於追蹤")
    for _k in QC.get('keyMetrics', []):
        _m = next(x for x in QMET if x['key'] == _k); _fr = QT[f"焦點｜{_m['label']}"]
        srow(f"驗證｜{_m['label']}（模型／共識／指引低／指引高／實際）", _m['unit'], [f"={_QTS}{c}{_fr}" for c in "CDFGI"],
             PCT if _m['unit'] == "%" else NUM, "共識、指引為 Interested-party")
        srow(f"驗證｜{_m['label']}｜句", "", [f"=INDEX({_QTS}$C${LINE[_k]}:${QCL[-1]}${LINE[_k]},{FOCUS})"], bold=True)
srow("驗證｜評等分布（強力買進／買進／持有／賣出／強力賣出）", "家", [f"={CIR('評等分布｜' + a)}" for a in ("強力買進", "買進", "持有", "賣出", "強力賣出")], NUM0,
     f"{CONS['ratings']['month']}，合計 {CONS['ratings']['total']} 家 [Interested-party]")
srow("驗證｜目標價（平均／中位數／最低／最高）", "US$", [f"={CIR('目標價｜' + a)}" for a in ("平均", "中位數", "最低", "最高")], USD,
     f"{CONS['priceTarget']['analysts']} 家；{CONS['priceTarget']['source']}，擷取 {CONS['priceTarget']['retrieved']} [Interested-party]")
ws.cell(row=r + 1, column=1, value="來源獨立性：" + CONS['sourceIndependence']).font = SMALL
ws.cell(row=r + 2, column=1, value=f"共識來源：{CONS['annualEstimates']['source']}、{CONS['priceTarget']['source']}；擷取 {CONS['annualEstimates']['retrieved']}。逐筆來源、日期與標記見『輸入與假設』I 區。").font = SMALL

# W2（2026-10-07）：「每MW經濟性」頁——世代結構推導、由下而上營運成本、每 MW 收入／資本支出對照、最近一季實際對照、彙總表、敏感度快照。
# 每 MW 值＝期間金額 ÷ 平均在役 MW（公司口徑，Accepted）÷ 期間長度 × 1000（US$m／MW／年）。HTML「資金模型 → 運營活動 → 每 MW 經濟性」同一邏輯（cmp31 逐列比對）。
PM = {}
if FL:
    ws = wb.create_sheet("每MW經濟性")
    ws.column_dimensions["A"].width = 52
    ws.column_dimensions["B"].width = 14
    for c in COLS: ws.column_dimensions[c].width = 13
    ws.column_dimensions["I"].width = 96
    ws["A1"] = "每 MW 經濟性（W2：Tokenomics 物理與產業數據 × 公司專屬輸入）"; ws["A1"].font = TITLE
    ws["A2"] = (f"方法：資本支出 {PMW['capex']}／營運成本 {PMW['cost']}／收入 {PMW['revenue']}（company.json → methodology.perMw）。MW 口徑："
                + ("IT 關鍵電力" if MWBASIS == 'IT' else "設施（Tokenomics 值已換成每設施 MW）") + "。每 MW 值＝期間金額 ÷ 平均在役 MW ÷ 期間長度（年化）。"
                + (f"Tokenomics 名稱缺漏 {len(TK_MISS)} 項（{'、'.join(TK_MISS)}），相關成本為暫代值（待 v5.26）。" if TK_MISS else ""))
    ws["A2"].font = SMALL
    _I = "'輸入與假設'!"; _C = "'運營_產能與收入'!"; _F = "'各期收支'!"
    _L = lambda i: f"{_I}{COLS[i]}${IN['模型期長度（年）']}"
    _NG = len(GEN)
    def mrow(name, unit, fn, fmt=NUM, font=BLACK, note=None, bold=False):
        global r
        ws.cell(row=r, column=1, value=name).font = BOLD if bold else BLACK
        ws.cell(row=r, column=2, value=unit).font = SMALL
        for i in range(5):
            v = fn(i); c = ws.cell(row=r, column=3 + i, value=v); c.font = font; c.number_format = fmt; c.border = BOX
        if note: ws.cell(row=r, column=9, value=note).font = SMALL
        PM[name] = r; r += 1
    r = 4
    r = period_header(ws, r)
    # 彙總表在最上方（列數固定，數值引用下方明細）：先保留列號
    SUMROWS = ["平均在役 MW（IT 關鍵電力）", "平均在役 MW（設施＝IT × IF_FacilityGW）", "每 MW 年收入（算力＋服務）", "　其中：算力收入", "電費", "IT 維護", "人員、軟體、水與耗材", "財產稅與保險",
               "公司管銷與其他", "租金", "現金成本合計（含租金）", "EBITDA", "D&A（模型車隊折舊）", "利息", "稅前", "每 MW 資本支出（新增 MW 的建置成本）", "每設施 MW 年收入", "每設施 MW EBITDA"]
    r = section(ws, r, "每 MW 經濟性彙總（每平均在役 MW、年化；US$m／MW／年；目前情境）")
    sum0 = r; r += len(SUMROWS) + 1
    # ---- 世代結構推導 ----
    r = section(ws, r, "世代結構推導（期初在役 → 汰換 → 新增 → 期末／平均在役；MW，公司口徑）", level=2, collapsed=True)
    mrow("期初在役 MW 合計", "MW", lambda i: f"={OPENMW}" if i == 0 else f"={COLS[i-1]}{{END_T}}", NUM0, BLACK, f"{PERIODS[0]} 期初＝{FL['openMix']['asOf']} 在役（世代組合起點）；之後＝前期期末")
    _st = r
    for j, g in enumerate(GEN):
        mrow(f"期初在役 MW｜{g}", "MW", lambda i, j=j, g=g: f"={OPENMW}*{FLX['open'][g]}" if i == 0 else f"={COLS[i-1]}{{END_{j}}}", NUM0)
    mrow("汰換 MW（壽命到期批次，最舊世代先出）", "MW", lambda i: f"=MIN({_vintage(i)},{COLS[i]}{PM['期初在役 MW 合計']})", NUM0, BLACK,
         "＝(本年 − GPU 經濟壽命) 那年新增的 MW（與『GPU 汰換 CapEx』同一批次），由世代清單中最舊的世代先出，以當期新增世代補回")
    _rt = PM["汰換 MW（壽命到期批次，最舊世代先出）"]; _ro = r
    for j, g in enumerate(GEN):
        mrow(f"汰換出｜{g}", "MW", lambda i, j=j: (f"=MIN({COLS[i]}{_st + j},{COLS[i]}{_rt})" if j == 0 else
                                                 f"=MIN({COLS[i]}{_st + j},MAX(0,{COLS[i]}{_rt}-SUM({COLS[i]}{_st}:{COLS[i]}{_st + j - 1})))"), NUM0)
    mrow("新增 MW（期末 Accepted − 期初，不為負）", "MW", lambda i: f"=MAX(0,{_C}{COLS[i]}{CAP['acc']}-{COLS[i]}{PM['期初在役 MW 合計']})", NUM0)
    _ad = PM["新增 MW（期末 Accepted − 期初，不為負）"]; _en = r
    for j, g in enumerate(GEN):
        mrow(f"期末在役 MW｜{g}", "MW", lambda i, j=j: f"={COLS[i]}{_st + j}-{COLS[i]}{_ro + j}+{_I}{COLS[i]}{FLX['adopt'] + j}*({COLS[i]}{_ad}+{COLS[i]}{_rt})", NUM0)
    mrow("期末在役 MW 合計", "MW", lambda i: f"=SUM({COLS[i]}{_en}:{COLS[i]}{_en + _NG - 1})", NUM0)
    _av = r
    for j, g in enumerate(GEN):
        mrow(f"平均在役 MW｜{g}", "MW", lambda i, j=j: f"=({COLS[i]}{_st + j}+{COLS[i]}{_en + j})/2", NUM0)
    mrow("平均在役 MW 合計", "MW", lambda i: f"=SUM({COLS[i]}{_av}:{COLS[i]}{_av + _NG - 1})", NUM0, BOLD, "公司 MW 口徑（Accepted）；每 MW 值的分母")
    _AT = PM["平均在役 MW 合計"]; _mx = r
    for j, g in enumerate(GEN):
        mrow(f"平均在役占比｜{g}", "%", lambda i, j=j: f"={COLS[i]}{_av + j}/MAX({COLS[i]}{_AT},1E-9)", PCT)
    for rr in range(_st - 1, r):  # 期初列引用前期期末（列號在上方建立後才知道）
        for i in range(5):
            c = ws.cell(row=rr, column=3 + i)
            if isinstance(c.value, str) and '{END_' in c.value:
                c.value = c.value.replace('{END_T}', str(PM['期末在役 MW 合計']))
                for j in range(_NG): c.value = c.value.replace(f'{{END_{j}}}', str(_en + j))
    MIXR = lambda i: f"{COLS[i]}{_mx}:{COLS[i]}{_mx + _NG - 1}"   # 平均在役占比（第 i 期）
    TKR = lambda blk, col: f"{_I}${col}${FLX[blk]}:${col}${FLX[blk] + _NG - 1}"
    # ---- 由下而上營運成本 ----
    r = section(ws, r, "由下而上營運成本（租金前；US$bn，模型期）" + ("——模型採用" if PMW['cost'] == 'bottomUp' else "——對照（模型採用 EBITDA 率路徑）"), level=2, collapsed=True)
    mrow("每 MW 電費（世代加權）", "US$m/MW", lambda i: f"=SUMPRODUCT({MIXR(i)},{TKR('tkOp', 'C')})", '0.000', BLACK, "Tokenomics IF_PowerCost（每 IT GW 年電費，已含平均用電比與 PUE）")
    mrow("每 MW IT 維護（世代加權）", "US$m/MW", (lambda i: f"=SUMPRODUCT({MIXR(i)},{TKR('tkOp', 'D')})") if tk_has('IF_MaintIT') else (lambda i: f"={_I}{COLS[i]}${IN['維護成本']}"), '0.000', BLACK,
         "Tokenomics IF_MaintIT" if tk_has('IF_MaintIT') else "暫代值：C 區『維護成本』（CRWV 現值；待 Tokenomics v5.26 IF_MaintIT）")
    mrow("每 MW 人員、軟體、水與耗材", "US$m/MW", (lambda i: f"=SUMPRODUCT({MIXR(i)},{TKR('tkOp', 'E')})") if tk_has('IF_StaffSW') else (lambda i: "=0"), '0.000', BLACK,
         "Tokenomics IF_StaffSW" if tk_has('IF_StaffSW') else "暫代值 0（待 Tokenomics v5.26 IF_StaffSW）")
    mrow("每 MW 財產稅與保險（只算 IT 部分）", "US$m/MW", (lambda i: f"=SUMPRODUCT({MIXR(i)},{TKR('tkOp', 'F')})") if tk_has('IF_TaxIns') else (lambda i: "=0"), '0.000', BLACK,
         "Tokenomics IF_TaxIns × IT 資本占總資本比（廠房屬房東）" if tk_has('IF_TaxIns') else "暫代值 0（待 Tokenomics v5.26 IF_TaxIns）")
    for nm, k in (("電費", "每 MW 電費（世代加權）"), ("IT 維護", "每 MW IT 維護（世代加權）"), ("人員、軟體、水與耗材", "每 MW 人員、軟體、水與耗材"), ("財產稅與保險", "每 MW 財產稅與保險（只算 IT 部分）")):
        mrow(f"{nm}（金額）", "US$bn", lambda i, k=k: f"={COLS[i]}{_AT}*{COLS[i]}{PM[k]}/1000*{_L(i)}", NUM, BLACK, "＝平均在役 MW × 每 MW 值 ÷ 1000 × 期間長度" if nm == "電費" else None)
    mrow("模型期總營收（算力＋服務）", "US$bn", lambda i: f"={_C}{COLS[i]}{CAP['totrev']}", NUM, GREEN)
    mrow("管銷率", "%", lambda i: f"={CSX['sga']}", PCT, GREEN, "C 區『由下而上營運成本輸入』：最近一季實際，各期持平 [Derived]")
    mrow("公司管銷與其他（金額）", "US$bn", lambda i: f"={COLS[i]}{PM['模型期總營收（算力＋服務）']}*{COLS[i]}{PM['管銷率']}", NUM)
    mrow("現金營運成本合計（租金前）", "US$bn", lambda i: "=" + "+".join(f"{COLS[i]}{PM[k]}" for k in ("電費（金額）", "IT 維護（金額）", "人員、軟體、水與耗材（金額）", "財產稅與保險（金額）", "公司管銷與其他（金額）")), NUM, BOLD,
         "廠房維護與廠房折舊不計（房東負擔，含在租金）")
    mrow("由下而上 EBITDAR 率", "%", lambda i: f"=1-{COLS[i]}{PM['現金營運成本合計（租金前）']}/MAX(0.01,{COLS[i]}{PM['模型期總營收（算力＋服務）']})", PCT, BLACK, "＝1 −（現金營運成本 ÷ 營收）")
    mrow("租金合計", "US$bn", lambda i: f"={_F}{COLS[i]}{FRR['　租金合計']}", NUM, GREEN, "在帳＋表外現金租金（只扣一次：EBITDA 率＝EBITDAR 率 − 租金 ÷ 營收）")
    mrow("由下而上 EBITDA 率", "%", lambda i: f"={COLS[i]}{PM['由下而上 EBITDAR 率']}-{COLS[i]}{PM['租金合計']}/MAX(0.01,{COLS[i]}{PM['模型期總營收（算力＋服務）']})", PCT, BOLD)
    mrow("EBITDA 率（模型採用）", "%", lambda i: f"={inref('EBITDA 率', i)}", PCT, GREEN,
         "＝由下而上 EBITDA 率＋（穩態輸入 − 由下而上 FY30）× 期數 ÷ 4（穩態輸入預設＝由下而上 FY30，差額 0）" if PMW['cost'] == 'bottomUp' else "舊方法：起始→穩態線性（由下而上只作對照）")
    BU_EB = PM["由下而上 EBITDA 率"]
    # ---- 每 MW 收入對照 ----
    r = section(ws, r, "每 MW 收入對照（收入方法：" + {'gpuHr': "GPU 小時價格", 'tkAnchor': "Tokenomics 錨 × 定價倍數 k（下一區）——以下為對照列"}.get(PMW['revenue'], "備案 legacy——每 MW 年收入為輸入，以下為對照列，不入損益") + "）", level=2, collapsed=True)
    mrow("每 MW 年收入（模型採用，100% 計費時數）", "US$m/MW", lambda i: f"={inref('每 MW 年收入', i)}*{REVSC}*1000", NUM, GREEN, "＝B 區『每 MW 年收入』× 倍數；收入端另乘利用率")
    mrow("每 MW 年收入（計費後＝× 利用率）", "US$m/MW", lambda i: f"={COLS[i]}{PM['每 MW 年收入（模型採用，100% 計費時數）']}*{inref('利用率', i)}", NUM)
    mrow("每 MW 年收入（v4.6 舊輸入 m.revMW，對照）" if PMW['revenue'] == 'tkAnchor' else "每 MW 年收入（v4.5 舊值，對照）", "US$m/MW", lambda i: M['revMW'][i] * 1000, NUM, BLUE, "company.json → defaults.m.revMW（舊方法以期末 ARR 指引回推，只作對照）")
    _OLDREV = "每 MW 年收入（v4.6 舊輸入 m.revMW，對照）" if PMW['revenue'] == 'tkAnchor' else "每 MW 年收入（v4.5 舊值，對照）"
    mrow("期末 ARR 指引 ÷ 年底主動電力（公司數字，只作對照）", "US$m/MW", lambda i: f"=({CO['callFacts']['arrLo']}+{CO['callFacts']['arrHi']})/2/{CO['callFacts']['yeActiveGw']}" if i == 0 else None, NUM, BLACK,
         f"法說：期末 ARR {CO['callFacts']['arrLo']}–{CO['callFacts']['arrHi']} ÷ 年底 {CO['callFacts']['yeActiveGw']} GW [Derived]；公司數字不作參數（共同規則第 4 節）")
    mrow("每 MW GPU 數（世代加權）", "顆/MW", lambda i: f"=SUMPRODUCT({MIXR(i)},{TKR('tkRev', 'C')})", NUM1, BLACK, "Tokenomics IF_GPUsPerGW ÷ 1000（× MW 換算）")
    _hasPx = all((((CO.get('pricing') or {}).get('gpuHr') or {}).get(g) or {}).get('base') is not None for g in GEN)
    mrow("GPU 小時價格路線：每 MW 年收入", "US$m/MW", (lambda i: f"=SUMPRODUCT({MIXR(i)},{TKR('tkRev', 'C')},{TKR('tkRev', 'F')})*8760/1000000") if _hasPx else (lambda i: "不適用"), NUM, BLACK,
         "＝Σ 平均在役占比 × 每 MW GPU 數 × GPU 小時合約價 × 8,760 ÷ 10⁶（100% 計費時數）" + ("" if _hasPx else "；company.json → pricing.gpuHr 未填（W1：長約價不足兩個獨立來源）→ 不適用"))
    if PMW['revenue'] == 'gpuHr':  # W2：B 區「每 MW 年收入」改為 GPU 小時路線（US$bn/MW；舊值見上方對照列）
        _wi = wb["輸入與假設"]
        for i in range(5):
            c = _wi.cell(row=IN["每 MW 年收入"], column=3 + i, value=f"='每MW經濟性'!{COLS[i]}{PM['GPU 小時價格路線：每 MW 年收入']}/1000"); c.font = GREEN
        _wi.cell(row=IN["每 MW 年收入"], column=9, value="W2 GPU 小時路線：＝『每MW經濟性』Σ 平均在役占比 × 每 MW GPU 數 × GPU 小時合約價 × 8,760（100% 計費時數；利用率在收入端另乘）").font = SMALL
    mrow("隱含每 GPU 小時價格（反算對照，不是輸入）", "US$/GPU-hr", lambda i: f"={COLS[i]}{PM['每 MW 年收入（模型採用，100% 計費時數）']}*1000000/MAX(1E-9,{COLS[i]}{PM['每 MW GPU 數（世代加權）']}*8760)", USD, BLACK,
         "＝每 MW 年收入 ÷（每 MW GPU 數 × 8,760）；利用率在收入端另乘，故不再除以利用率")
    mrow("每 GPU 小時經濟持有成本（GPU 數加權）", "US$/GPU-hr", lambda i: f"=SUMPRODUCT({MIXR(i)},{TKR('tkRev', 'C')},{TKR('tkRev', 'D')})/MAX(1E-9,{COLS[i]}{PM['每 MW GPU 數（世代加權）']})", USD, BLACK,
         "Tokenomics IF_GPUhrEcon（100% 時數、不含利潤）")
    mrow("隱含價格 ÷ GPU 小時持有成本", "倍", lambda i: f"={COLS[i]}{PM['隱含每 GPU 小時價格（反算對照，不是輸入）']}/MAX(1E-9,{COLS[i]}{PM['每 GPU 小時經濟持有成本（GPU 數加權）']})", MULT)
    mrow("每 MW 經濟持有成本（不賠錢下限）", "US$m/MW", lambda i: f"=SUMPRODUCT({MIXR(i)},{TKR('tkRev', 'E')})", NUM, BLACK,
         "Tokenomics IF_HoldEcon 依在役世代加權（100% 時數、含廠房資本回收；CRWV 廠房走租金，故此線高於 CRWV 的 IT 持有成本）")
    mrow("每 MW 年收入 ÷ 經濟持有成本", "倍", lambda i: f"={COLS[i]}{PM['每 MW 年收入（模型採用，100% 計費時數）']}/MAX(1E-9,{COLS[i]}{PM['每 MW 經濟持有成本（不賠錢下限）']})", MULT)
    for _pr in (CO.get('pricing') or {}).get('peerRevPerMw', []):
        mrow(f"同業｜{_pr['label']}", _pr['unit'], lambda i, v=_pr['value']: v, NUM, BLUE, f"{_pr['source']}（{_pr['date']}）{_pr['tag']}；{_pr.get('note', '')}")
    for _mr in (CO.get('pricing') or {}).get('marketRefs', []):
        mrow(f"市場價格｜{_mr['label']}", "US$/GPU-hr", lambda i, v=_mr['value']: v if i == 0 else None, USD, BLUE, f"{_mr['gen']}；{_mr['basis']}；{_mr['source']}（{_mr['date']}）{_mr['tag']}")
    # ---- W4：每 MW 收入＝Tokenomics 錨 × 定價倍數 k（company.json → pricing.anchorMultiple）----
    _gq = lambda x: f'{x:g}'
    if AMQ:
        _TKON = PMW['revenue'] == 'tkAnchor'
        r = section(ws, r, "每 MW 收入：Tokenomics 錨 × 定價倍數 k（W4）" + ("——模型採用" if _TKON else "——對照（模型採用其他收入方法）"), level=2, collapsed=True)
        _avgB = lambda i: f"{_C}{COLS[i]}{CAP['avg']}"
        mrow("錨｜每 MW 經濟持有成本（IF_HoldEcon，在役世代加權）", "US$m/MW", lambda i: f"={COLS[i]}{PM['每 MW 經濟持有成本（不賠錢下限）']}", NUM, BLACK,
             "＝Σ 平均在役占比 × IF_HoldEcon（目前 Tokenomics 成本情境；WACC 10% 下的打平租金，100% 時數，含廠房資本回收）")
        mrow("錨｜排程 RPO（模型期）", "US$bn", lambda i: f"={_C}{COLS[i]}{CAP['sch']}", NUM, GREEN, "『運營_產能與收入』排程 RPO（期初 RPO＋Q3 新增 × 桶權重 × 認列比例）；只用於 RPO 覆蓋率對照列，不用來反推 k")
        mrow("錨｜平均計費 MW", "MW", lambda i: f"={_avgB(i)}", NUM0, GREEN, "『運營_產能與收入』平均在役（Billable）MW；收入端的分母")
        mrow("錨｜利用率", "%", lambda i: f"={inref('利用率', i)}", PCT, GREEN)
        mrow("長約價產能收入（平均計費 MW × 錨 × k_長約 × 利用率 × 期間長度）", "US$bn",
             lambda i: f"={COLS[i]}{PM['錨｜平均計費 MW']}*{COLS[i]}{PM['錨｜每 MW 經濟持有成本（IF_HoldEcon，在役世代加權）']}/1000*{AMX['kL']}*{COLS[i]}{PM['錨｜利用率']}*{_L(i)}", NUM, BLACK,
             "在役計費產能全部以長約價計的收入（US$bn）")
        mrow("RPO 涵蓋的產能 ÷ 在役計費產能（對照，不驅動）", "%", lambda i: f"={COLS[i]}{PM['錨｜排程 RPO（模型期）']}/MAX(1E-9,{COLS[i]}{PM['長約價產能收入（平均計費 MW × 錨 × k_長約 × 利用率 × 期間長度）']})", PCT, BLACK,
             "＝排程 RPO ÷ 長約價產能收入：RPO 金額以長約價換成 MW 後占在役計費 MW 的比例 [Derived]；r2 起只作對照——未被 RPO 覆蓋的新增產能也按長約價出售（先簽約再建產能）")
        _ref = lambda side: next(e_ for e_ in AMQ['evidence'] if e_['label'] == AMQ[side]['refEvidence'])
        _rk = lambda e_: f"TK_{re.sub(r'^(IF|L1)_', '', e_['tkName'])}_{_TKGC[e_['gen']]}"
        _rL, _rS = _ref('long'), _ref('spot')
        mrow("k_長約（依目前成本情境重算）", "倍", lambda i: f"={AMX['kL']}*{_rk(_rL)}/OFFSET({_rk(_rL)},0,{TKSEL}-2)", '0.000', BLACK,
             f"＝輸入 k_長約 × 基準證據世代成本（{_rL['tkName']}｜{_rL['gen']}，基準）÷ 同名稱目前成本情境：價格是事實，k＝價格 ÷ 同情境持有成本")
        mrow("k_現貨（依目前成本情境重算）", "倍", lambda i: f"={AMX['kS']}*{_rk(_rS)}/OFFSET({_rk(_rS)},0,{TKSEL}-2)", '0.000', BLACK,
             f"＝輸入 k_現貨 × {_rS['tkName']}｜{_rS['gen']}（基準）÷ 目前成本情境")
        mrow("隨需占比（輸入）", "%", lambda i: f"={AMX['od']}", PCT, GREEN, "輸入頁 B 區後『定價倍數 k』")
        mrow("定價倍數 k（隨需占比 × k_現貨 ＋（1 − 隨需占比）× k_長約）", "倍",
             lambda i: f"={COLS[i]}{PM['隨需占比（輸入）']}*{COLS[i]}{PM['k_現貨（依目前成本情境重算）']}+(1-{COLS[i]}{PM['隨需占比（輸入）']})*{COLS[i]}{PM['k_長約（依目前成本情境重算）']}", '0.000', BOLD)
        mrow("Tokenomics 錨 × k：每 MW 年收入（100% 計費時數）", "US$m/MW",
             lambda i: f"={COLS[i]}{PM['錨｜每 MW 經濟持有成本（IF_HoldEcon，在役世代加權）']}*{COLS[i]}{PM['定價倍數 k（隨需占比 × k_現貨 ＋（1 − 隨需占比）× k_長約）']}", NUM, BOLD,
             ("模型採用：輸入頁 B 區『每 MW 年收入』＝本列 ÷ 1000（收入端另乘利用率與爬坡分母）" if _TKON else "對照（模型未採用）"))
        mrow("錨 × k 相對 v4.6 舊輸入 m.revMW", "%", lambda i: f"={COLS[i]}{PM['Tokenomics 錨 × k：每 MW 年收入（100% 計費時數）']}/MAX(1E-9,{COLS[i]}{PM[_OLDREV]})-1", PCT, BLACK)
        _hasRF = tk_has('IF_RevGWFleet')
        mrow("上限檢查｜客戶每 MW 付費 token 營收（IF_RevGWFleet，在役世代加權）", "US$m/MW",
             (lambda i: f"=SUMPRODUCT({MIXR(i)},{TKR('tkCapRev', 'C')})") if _hasRF else (lambda i: "不適用"), NUM, BLACK,
             "Tokenomics 客戶端（OpenAI 有效單價、層級組合）每 GW 付費 token 營收的理想上限，依在役世代加權" + ("" if _hasRF else "；快照無 IF_RevGWFleet → 不適用"))
        mrow("上限檢查｜CRWV 每 MW 計費收入 ÷ 客戶付費 token 營收", "%",
             (lambda i: f"={COLS[i]}{PM['每 MW 年收入（計費後＝× 利用率）']}/MAX(1E-9,{COLS[i]}{PM['上限檢查｜客戶每 MW 付費 token 營收（IF_RevGWFleet，在役世代加權）']})") if _hasRF else (lambda i: "不適用"), PCT, BOLD,
             f"＝neocloud 拿走客戶 token 營收的比例；> {_gq(CK['revCapShareMax'] * 100)}%（company.json → methodology.checks.revCapShareMax）時『檢查_連動』警示")
        mrow("上限檢查｜各期結果", "", (lambda i: f'=IF({COLS[i]}{PM["上限檢查｜CRWV 每 MW 計費收入 ÷ 客戶付費 token 營收"]}>{_gq(CK["revCapShareMax"])},"警示","通過")') if _hasRF else (lambda i: "不適用"), NUM, BOLD)
        # k 證據表（市場價格 ÷ Tokenomics 同世代持有成本；成本取基準、不隨成本情境）
        r = section(ws, r, "定價倍數 k 證據表（市場價格 ÷ Tokenomics 同世代持有成本；用途：long＝k_長約基準、spot＝k_現貨基準、range＝支持區間、list＝只列不用）", level=2, collapsed=True)
        for j, h in enumerate(["證據", "單位", "價格", "Tokenomics 同世代", "倍數", "用途", ""]):
            c = ws.cell(row=r, column=1 + j, value=h or None); c.font = HEAD; c.fill = FILL_HEAD
        ws.cell(row=r, column=9, value="Tokenomics＝IF_GPUhrEcon（US$/GPU-hr）或 IF_HoldEcon（US$m／IT MW／年），基準成本情境；世代｜合約型態｜期間｜來源（文件日期）標記｜說明").font = SMALL
        r += 1
        for e_ in AMQ['evidence']:
            ws.cell(row=r, column=1, value=f"k 證據｜{e_['label']}").font = BLACK
            ws.cell(row=r, column=2, value=e_['unit']).font = SMALL
            c = ws.cell(row=r, column=3, value=e_['price']); c.font = BLUE; c.number_format = '0.000'; c.border = BOX
            _k = f"TK_{re.sub(r'^(IF|L1)_', '', e_['tkName'])}_{_TKGC[e_['gen']]}"
            c = ws.cell(row=r, column=4, value=f"={_k}" if tk_has(e_['tkName']) else "不適用"); c.font = GREEN; c.number_format = '0.000'; c.border = BOX
            c = ws.cell(row=r, column=5, value=f"=C{r}/D{r}" if tk_has(e_['tkName']) else "不適用"); c.font = BOLD; c.number_format = '0.00'; c.border = BOX
            ws.cell(row=r, column=6, value=e_['use']).font = BLACK
            ws.cell(row=r, column=9, value=f"{e_['gen']}｜{e_['contract']}｜{e_['term']}｜{e_['source']}（{e_['date']}）{e_['tag']}｜{e_['note']}" + (f"｜{e_['url']}" if e_['url'] else "")).font = SMALL
            PM[f"k 證據｜{e_['label']}"] = r; r += 1
        for cm in AMQ.get('contractMix', []):
            ws.cell(row=r, column=1, value=f"合約組合｜{cm['label']}").font = BLACK
            ws.cell(row=r, column=2, value=cm['unit'] or None).font = SMALL
            c = ws.cell(row=r, column=3, value=cm['value'] if cm['value'] is not None else cm['tag']); c.font = BLUE; c.border = BOX
            ws.cell(row=r, column=9, value=f"{cm['source']}" + (f"（{cm['date']}）" if cm['date'] else "") + f" {cm['tag']}" + (f"｜{cm['url']}" if cm['url'] else "")).font = SMALL
            r += 1
        for j_, nf in enumerate(AMQ.get('notFound', [])):
            ws.cell(row=r, column=1, value=f"找不到｜第 {j_ + 1} 項").font = BLACK
            ws.cell(row=r, column=9, value=nf).font = SMALL; r += 1
        if _TKON and _hasRF:  # 檢查_連動：收入上限（各期最高）
            _wc = wb["檢查_連動"]
            for rr in range(4, _wc.max_row + 1):
                if _wc.cell(row=rr, column=1).value == REVCAP_CHECK:
                    _rc = PM['上限檢查｜CRWV 每 MW 計費收入 ÷ 客戶付費 token 營收']
                    c = _wc.cell(row=rr, column=2, value=f"=MAX('每MW經濟性'!C{_rc}:G{_rc})"); c.number_format = PCT; c.font = BLACK
        if _TKON:  # B 區「每 MW 年收入」改為 Tokenomics 錨 × k（US$bn/MW；舊輸入見上方對照列）
            _wi = wb["輸入與假設"]
            for i in range(5):
                c = _wi.cell(row=IN["每 MW 年收入"], column=3 + i, value=f"='每MW經濟性'!{COLS[i]}{PM['Tokenomics 錨 × k：每 MW 年收入（100% 計費時數）']}/1000"); c.font = GREEN
            _wi.cell(row=IN["每 MW 年收入"], column=9, value="W4 Tokenomics 錨：＝『每MW經濟性』Σ 平均在役占比 × IF_HoldEcon × 定價倍數 k（100% 計費時數；利用率在收入端另乘）").font = SMALL
    # ---- 每 MW 資本支出對照 ----
    r = section(ws, r, "每 MW 資本支出對照（US$m／MW；新增 MW 的世代加權）", level=2, collapsed=True)
    mrow("每 MW 建置成本（模型採用）", "US$m/MW", lambda i: f"={inref('每 MW 建置成本', i)}", NUM1, GREEN)
    mrow("Tokenomics IT 設備（IF_CapexIT）", "US$m/MW", lambda i: f"=SUMPRODUCT({_I}{COLS[i]}${FLX['adopt']}:{COLS[i]}${FLX['adopt'] + _NG - 1},{TKR('tkCap', 'C')})", NUM1)
    mrow("Tokenomics 含廠房合計（IF_CapexTotal；對照）", "US$m/MW", lambda i: f"=SUMPRODUCT({_I}{COLS[i]}${FLX['adopt']}:{COLS[i]}${FLX['adopt'] + _NG - 1},{TKR('tkCap', 'D')})", NUM1, BLACK,
         "CRWV 機房以租約取得、不含廠房；合計只作對照")
    mrow("v4.5 舊值（以 FY26 CapEx 指引校準）", "US$m/MW", lambda i: CO['scenarios']['capexTemplate']['costMW'][i], NUM1, BLUE, "舊方法，只作對照")
    mrow("每 MW 建置成本 ÷ 壽命（在役世代加權；D&A 參考）", "US$m/MW", lambda i: f"=SUMPRODUCT({MIXR(i)},{TKR('tkCap', 'C')})/{LIFE}", NUM)
    # ---- 最近一季實際對照 ----
    r = section(ws, r, "最近一季實際對照（10-Q；不強制平衡）", level=2, collapsed=True)
    mrow("最近一季平均在役 MW", "MW", lambda i: f"={OPENMW}-{CSX['addQ']}/2" if i == 0 else None, NUM0, BLACK, "＝季末主動電力 − 當季新增 ÷ 2")
    mrow("最近一季現金營運成本（租金前，不含管銷）", "US$bn", lambda i: f"={CSX['costRev']}+{CSX['techInfra']}-{CSX['da']}-{CSX['sbcCostTi']}-{CSX['opLeaseCost']}-{CSX['varLeaseCost']}" if i == 0 else None, NUM, BLACK,
         "＝營收成本＋技術與基礎設施 − D&A − SBC − 營業租賃成本 − 變動租賃成本")
    mrow("最近一季每 MW 現金營運成本（年化，不含管銷）", "US$m/MW", lambda i: f"=C{PM['最近一季現金營運成本（租金前，不含管銷）']}*4/C{PM['最近一季平均在役 MW']}*1000" if i == 0 else None, '0.000')
    mrow("模型由下而上每 MW 現金營運成本（不含管銷）", "US$m/MW", lambda i: "=" + "+".join(f"{COLS[i]}{PM[k]}" for k in ("每 MW 電費（世代加權）", "每 MW IT 維護（世代加權）", "每 MW 人員、軟體、水與耗材", "每 MW 財產稅與保險（只算 IT 部分）")), '0.000')
    mrow("差距（模型 «P0» − 實際）", "US$m/MW", lambda i: f"=C{PM['模型由下而上每 MW 現金營運成本（不含管銷）']}-C{PM['最近一季每 MW 現金營運成本（年化，不含管銷）']}" if i == 0 else None, '0.000', BOLD,
         "可能原因：爬坡期閒置產能（分母含尚未計費的 MW）、未揭露的成本項、D&A 未依列別揭露；v5.26 前人員軟體與稅險暫代 0")
    mrow("最近一季每 MW 年租金（營業＋變動，年化）", "US$m/MW", lambda i: f"=({CSX['opLeaseCost']}+{CSX['varLeaseCost']})*4/C{PM['最近一季平均在役 MW']}*1000" if i == 0 else None, '0.000')
    mrow("模型每 MW 年租金", "US$m/MW", lambda i: f"={COLS[i]}{PM['租金合計']}/{COLS[i]}{_AT}/{_L(i)}*1000", '0.000')
    if AMQ and PMW['revenue'] == 'tkAnchor':  # W4：Q2 收入對照（只作驗證，不校準 k；拆解見 W4 報告）
        mrow("最近一季每 MW 年收入（營收 × 4 ÷ 平均在役 MW）", "US$m/MW", lambda i: f"={CSX['revenue']}*4/C{PM['最近一季平均在役 MW']}*1000" if i == 0 else None, '0.000', BLACK, "10-Q 營收年化 ÷ 平均在役 MW（含爬坡中尚未計費的 MW）")
        mrow("最近一季計費比例（季末 Billable ÷ 季末在役 MW，假設）", "%", lambda i: f"={MW0}/{OPENMW}" if i == 0 else None, PCT, BLACK, "輸入頁 B 區季末 Billable MW ÷ 世代組合期初在役 MW [Assumed]；用於換算每計費 MW")
        mrow("最近一季每 MW 年收入（÷ 平均計費 MW）", "US$m/MW", lambda i: f"=C{PM['最近一季每 MW 年收入（營收 × 4 ÷ 平均在役 MW）']}/C{PM['最近一季計費比例（季末 Billable ÷ 季末在役 MW，假設）']}" if i == 0 else None, '0.000')
        mrow("最近一季錨（季末在役世代 × IF_HoldEcon）", "US$m/MW", lambda i: ("=" + "+".join(f"{FLX['open'][g]}*{_I}$E${FLX['tkRev'] + j}" for j, g in enumerate(GEN))) if i == 0 else None, '0.000', BLACK, "世代組合期初占比（季末）× Tokenomics IF_HoldEcon（目前成本情境）")
        mrow("最近一季隱含 k（年化營收 ÷ 平均在役 MW ÷ 錨，未調整）", "倍", lambda i: f"=C{PM['最近一季每 MW 年收入（營收 × 4 ÷ 平均在役 MW）']}/C{PM['最近一季錨（季末在役世代 × IF_HoldEcon）']}" if i == 0 else None, '0.000', BOLD,
             "只作驗證：不以此值校準 k（已決定事項 14）；含爬坡分母、服務收入、舊世代舊約價格")
    # ---- 敏感度（建置時快照；scripts/permw_sens.py 以 Excel 依各設定重算後寫入 permw_sens.json）----
    _SP = _osrv.path.join(_osrv.path.dirname(_osrv.path.abspath(__file__)), 'permw_sens.json')
    _SN = _jco.load(open(_SP, encoding='utf-8')) if _osrv.path.exists(_SP) else None
    r = section(ws, r, "敏感度（建置時快照：Tokenomics 成本情境、GPU 小時價格、世代組合、管銷率口徑 × 三情境；目標價 US$、融資缺口 US$bn）", level=2, collapsed=True)
    for j, h in enumerate(["敏感度｜欄位", "", "加權目標價", "融資缺口", "Δ目標價", "Δ融資缺口"]):
        c = ws.cell(row=r, column=1 + j, value=h or None); c.font = HEAD; c.fill = FILL_HEAD
    ws.cell(row=r, column=9, value="融資缺口＝MAX(0, −FY30 融資前累積現金)；Δ＝對同情境「基準（目前輸入）」；快照由 verify.sh 步驟 3c 產生，HTML 即時計算同一組設定").font = SMALL
    r += 1
    _SCN = [("low", "保守"), ("base", "基準"), ("high", "積極")]
    _SCN = [(k, CO['scenarios']['labels'][k].split(' ')[0]) for k, _ in _SCN]
    _snap_base = {}
    if _SN:
        for sk, sname in _SCN:
            row = _SN['scenarios'][sk]; b0 = row['base']
            for k, nm in _SN['cases']:
                ws.cell(row=r, column=1, value=f"敏感度｜{sname}｜{nm}").font = BOLD if k == 'base' else BLACK
                ws.cell(row=r, column=2, value="US$／US$bn").font = SMALL
                v = row.get(k)
                vals = ["不適用", None, None, None] if v is None else [v[0], v[1], v[0] - b0[0], v[1] - b0[1]]
                for j, x in enumerate(vals):
                    c = ws.cell(row=r, column=3 + j, value=x); c.border = BOX; c.font = BLACK if j else BOLD
                    c.number_format = USD if j in (0, 2) else NUM
                if k == 'base': _snap_base[sk] = r
                r += 1
        ws.cell(row=r, column=1, value="敏感度快照狀態（目前輸入的加權目標價 vs 快照）").font = BOLD
        ws.cell(row=r, column=3, value=f"='評價_DCF與目標價'!{TGT}").number_format = USD
        ws.cell(row=r, column=4, value=f"=CHOOSE({SEL},{','.join(f'C{_snap_base[k]}' for k, _ in _SCN)})").number_format = USD
        ws.cell(row=r, column=5, value=f'=IF(ABS(C{r}-D{r})<0.005,"與目前輸入一致","快照已過期")').font = BOLD
        ws.cell(row=r, column=9, value="改了任何輸入後快照不會自動更新：狀態顯示「快照已過期」時，以 python3 scripts/permw_sens.py 重新產生並重建").font = SMALL
        PM['敏感度快照狀態'] = r; r += 1
    else:
        ws.cell(row=r, column=1, value="敏感度快照：尚未產生（執行 scripts/verify.sh 步驟 3c 或 python3 scripts/permw_sens.py）").font = SMALL; r += 1
    # ---- 彙總表（填入保留列）----
    _pm = lambda x, i: f"{x}/MAX(1E-9,{COLS[i]}{_AT})/{_L(i)}*1000"
    _ITMW = lambda i: f"IF({MWB}=2,{COLS[i]}{_AT}/SUMPRODUCT({MIXR(i)},{TKR('tkCap', 'E')}),{COLS[i]}{_AT})"
    _FACW = lambda i: f"SUMPRODUCT({MIXR(i)},{TKR('tkCap', 'E')})"
    _rev = lambda i: f"{_C}{COLS[i]}{CAP['totrev']}"
    _ebd = lambda i: f"{_C}{COLS[i]}{CAP['totrev']}*{inref('EBITDA 率', i)}"
    SUMF = {
        "平均在役 MW（IT 關鍵電力）": (lambda i: f"={_ITMW(i)}", "MW", NUM0, "公司 MW 口徑＝IT 時同上；設施口徑時 ÷ IF_FacilityGW"),
        "平均在役 MW（設施＝IT × IF_FacilityGW）": (lambda i: f"={_ITMW(i)}*{_FACW(i)}", "MW", NUM0, "IF_FacilityGW 依在役世代加權（基準 1.2）"),
        "每 MW 年收入（算力＋服務）": (lambda i: "=" + _pm(_rev(i), i), "US$m/MW", NUM, "模型期總營收 ÷ 平均在役 MW ÷ 期間長度"),
        "　其中：算力收入": (lambda i: "=" + _pm(f"{_C}{COLS[i]}{CAP['isrev']}", i), "US$m/MW", NUM, None),
        "電費": (lambda i: "=" + _pm(f"{COLS[i]}{PM['電費（金額）']}", i), "US$m/MW", NUM, "由下而上明細（下方）"),
        "IT 維護": (lambda i: "=" + _pm(f"{COLS[i]}{PM['IT 維護（金額）']}", i), "US$m/MW", NUM, None if tk_has('IF_MaintIT') else "暫代值（待 v5.26）"),
        "人員、軟體、水與耗材": (lambda i: "=" + _pm(f"{COLS[i]}{PM['人員、軟體、水與耗材（金額）']}", i), "US$m/MW", NUM, None if tk_has('IF_StaffSW') else "暫代值 0（待 v5.26）"),
        "財產稅與保險": (lambda i: "=" + _pm(f"{COLS[i]}{PM['財產稅與保險（金額）']}", i), "US$m/MW", NUM, None if tk_has('IF_TaxIns') else "暫代值 0（待 v5.26）"),
        "公司管銷與其他": (lambda i: "=" + _pm(f"{COLS[i]}{PM['公司管銷與其他（金額）']}", i), "US$m/MW", NUM, None),
        "租金": (lambda i: "=" + _pm(f"{COLS[i]}{PM['租金合計']}", i), "US$m/MW", NUM, None),
        "現金成本合計（含租金）": (lambda i: "=" + _pm(f"({_rev(i)}-{_ebd(i)})", i), "US$m/MW", NUM, "＝收入 − EBITDA（模型採用的 EBITDA 率；由下而上時＝上列各項合計）"),
        "EBITDA": (lambda i: "=" + _pm(_ebd(i), i), "US$m/MW", NUM, "＝營收 × EBITDA 率（損益與資金共用）"),
        "D&A（模型車隊折舊）": (lambda i: "=" + _pm(f"{_I}{COLS[i]}${IN['D&A（車隊）']}", i), "US$m/MW", NUM, "模型車隊折舊（期初毛 PP&E＋本期成長型×½）÷ 壽命；在役世代加權的建置成本 ÷ 壽命見『資本支出對照』"),
        "利息": (lambda i: "=" + _pm(f"{_F}{COLS[i]}{FRR['int']}", i), "US$m/MW", NUM, "存量債務＋瀑布新債利息"),
        "稅前": (lambda i: f"={COLS[i]}{{EB}}-{COLS[i]}{{DA}}-{COLS[i]}{{IT}}", "US$m/MW", NUM, "＝EBITDA − D&A − 利息"),
        "每 MW 資本支出（新增 MW 的建置成本）": (lambda i: f"={inref('每 MW 建置成本', i)}*{CAPSC}", "US$m/MW", NUM1, "存量概念（每新增 MW 一次性），與上列年化數字不同"),
        "每設施 MW 年收入": (lambda i: f"={COLS[i]}{{RV}}*{COLS[i]}{{MI}}/{COLS[i]}{{MF}}", "US$m/MW", NUM, None),
        "每設施 MW EBITDA": (lambda i: f"={COLS[i]}{{EB}}*{COLS[i]}{{MI}}/{COLS[i]}{{MF}}", "US$m/MW", NUM, None),
    }
    for k, nm in enumerate(SUMROWS):
        rr = sum0 + k; fn, u, fmt, note = SUMF[nm]
        ws.cell(row=rr, column=1, value=nm).font = BOLD if nm in ("每 MW 年收入（算力＋服務）", "現金成本合計（含租金）", "EBITDA", "稅前") else BLACK
        ws.cell(row=rr, column=2, value=u).font = SMALL
        for i in range(5):
            v = fn(i).replace('{EB}', str(sum0 + SUMROWS.index("EBITDA"))).replace('{DA}', str(sum0 + SUMROWS.index("D&A（模型車隊折舊）"))).replace('{IT}', str(sum0 + SUMROWS.index("利息"))) \
                .replace('{RV}', str(sum0 + 2)).replace('{MI}', str(sum0)).replace('{MF}', str(sum0 + 1))
            c = ws.cell(row=rr, column=3 + i, value=v); c.font = BLACK; c.number_format = fmt; c.border = BOX
        if note: ws.cell(row=rr, column=9, value=note).font = SMALL
        PM[nm] = rr
    ws.cell(row=sum0 + len(SUMROWS), column=1, value="讀法：每 MW 收入 − 現金成本（含租金）＝EBITDA；再扣 D&A 與利息＝稅前。W3 前後對照報告直接讀本表。").font = SMALL
    if "結論｜每 MW 經濟性句" in SM:  # W3：回填一頁摘要的每 MW 句（欄 D＝首個完整財年）
        _X = lambda k: f"'每MW經濟性'!D{PM[k]}"
        wb["摘要"].cell(row=SM["結論｜每 MW 經濟性句"], column=3, value=(
            f'="每 MW（{PERIODS[1]}）：年收入 $"&TEXT({_X("每 MW 年收入（算力＋服務）")},"0.0")&"m、現金成本 $"&TEXT({_X("現金成本合計（含租金）")},"0.0")'
            f'&"m（含租金 $"&TEXT({_X("租金")},"0.0")&"m）、EBITDA $"&TEXT({_X("EBITDA")},"0.0")&"m；計費單價為 Tokenomics 經濟持有成本（含廠房資本回收的打平線）$"'
            f'&TEXT({_X("每 MW 經濟持有成本（不賠錢下限）")},"0.0")&"m 的 "&TEXT({_X("每 MW 年收入 ÷ 經濟持有成本")},"0.00")&" 倍"'
            f'&IF({_X("每 MW 年收入 ÷ 經濟持有成本")}<1,"，未回收全部持有成本","")&"。"'))
    # C 區 EBITDA 率：由下而上（穩態輸入預設＝由下而上 FY30，差額線性分攤）
    if PMW['cost'] == 'bottomUp':
        _wi = wb["輸入與假設"]
        for i in range(5):
            _wi.cell(row=IN["EBITDA 率"], column=3 + i, value=f"='每MW經濟性'!{COLS[i]}{BU_EB}+({EBSS}-'每MW經濟性'!$G${BU_EB})*{i}/4")
        _wi.cell(row=IN["EBITDA 率"], column=9, value="W2 由下而上：＝『每MW經濟性』由下而上 EBITDA 率＋（穩態 − 由下而上 FY30）× 期數 ÷ 4（公式）").font = SMALL
        _eb_r = int(EBSS.split('$')[-1]); _eb0 = int(EB0.split('$')[-1])
        _wi.cell(row=_eb_r, column=3, value=f"='每MW經濟性'!G{BU_EB}").font = BLACK
        _wi.cell(row=_eb_r, column=9, value="W2 由下而上：預設＝由下而上 FY30 EBITDA 率（公式）；改成數值即為 FY30 目標，差額線性分攤到各期（反向 DCF、敏感度用）").font = SMALL
        _wi.cell(row=_eb0, column=9, value="W2 由下而上模式下不使用（保留為舊方法對照）").font = SMALL

# W1（2026-10-07）：Tokenomics 取數分頁。值取自 company.json → tokenomics.snapshotFile 的版本固定快照（tools/tokenomics/import_tokenomics.py 產生），
# 藍字輸入格；每個名稱的「基準」值格另建具名範圍 TK_<名稱去掉 IF_／L1_>_<世代代碼>（單值名稱不加世代），供 W2 公式引用。W1 不得讓既有公式引用這些格。
from openpyxl.workbook.defined_name import DefinedName as _DN
TK = CO.get('tokenomics')
if TK:
    _TKS = _jco.load(open(_osrv.path.join(_osrv.path.dirname(_osrv.path.abspath(__file__)), TK['snapshotFile']), encoding='utf-8'))
    _src = _TKS['source']
    assert _src['commit'] == TK['commit'] and _src['version'] == TK['version'], 'company.json → tokenomics 的版本／commit 與快照檔不一致'
    assert list(_TKS['items']) == TK['names'], 'company.json → tokenomics.names 與快照檔名稱不一致'
    ws = wb.create_sheet("Tokenomics_取數")
    for _col, _w in zip("ABCDEFGHIJ", (26, 24, 44, 12, 30, 13, 13, 13, 24, 22)):
        ws.column_dimensions[_col].width = _w
    ws["A1"] = "Tokenomics 取數（算力相關的產業與物理層資料）"; ws["A1"].font = TITLE
    _GC = {g['name']: g['code'] for g in _src['generations']}
    ws["A2"] = (f"來源：{_src['repo']} {_src['file']}（{_src['version']}，commit {_src['commit'][:7]}，擷取 {_src['extractedAt']}）。"
                "只引用 IF_（不含 IF_Hdr*）與 L1_ 名稱；模型主值取「基準」，低／高只作敏感度；Tokenomics 每 GW＝IT 關鍵電力。")
    ws["A2"].font = SMALL
    ws["A3"] = "本分頁目前只顯示、尚未被任何公式引用（W1）；值為快照的藍字輸入格，更新方式見 repo 根目錄 README「共用工具：Tokenomics 取數」。"
    ws["A3"].font = SMALL
    r = 4
    for j, h in enumerate(["具名範圍（基準）", "名稱", "中文標籤", "單位", "世代", "低成本", "基準", "高成本", "Tokenomics 位置", "版本與 commit"]):
        c = ws.cell(row=r, column=1 + j, value=h); c.font = HEAD; c.fill = FILL_HEAD
    r += 1
    _ver = f"{_src['version']} · {_src['commit'][:7]}"
    _short = lambda n: re.sub(r'^(IF|L1)_', '', n)
    def _tkrow(key, name, lab, unit, gen, vals, loc, fmt):
        global r
        ws.cell(row=r, column=1, value=key).font = BLACK
        ws.cell(row=r, column=2, value=name).font = BLACK
        ws.cell(row=r, column=3, value=lab).font = BLACK
        ws.cell(row=r, column=4, value=unit).font = SMALL
        ws.cell(row=r, column=5, value=gen).font = BLACK
        for j, v in enumerate(vals):
            c = ws.cell(row=r, column=6 + j, value=v); c.border = BOX
            if isinstance(v, (int, float)): c.font = BLUE; c.number_format = fmt
            else: c.font = SMALL
        ws.cell(row=r, column=9, value=loc).font = SMALL
        ws.cell(row=r, column=10, value=_ver).font = SMALL
        r += 1
    def _fmt(unit, vals):
        nums = [abs(v) for v in vals if isinstance(v, (int, float))]
        if unit in ('%',): return '0.0%'
        if unit in ('顆', '架'): return NUM0
        if nums and max(nums) >= 1000: return NUM0
        return '#,##0.0000;(#,##0.0000);-'
    for _nm, _it in _TKS['items'].items():
        if _it.get('missing'):
            _tkrow(f"TK_{_short(_nm)}", _nm, "（Tokenomics 尚無此名稱）", "", "", ["", "", ""],
                   f"{_src['version']} 無此名稱（列為 optional；Tokenomics v5.25 預計新增）；未建具名範圍", NUM)
            continue
        if _it['kind'] == 'gen_cost':
            for _g, _v3 in _it['values'].items():
                _vals = [_v3['低成本'], _v3['基準'], _v3['高成本']]
                _key = f"TK_{_short(_nm)}_{_GC[_g]}"
                _tkrow(_key, _nm, _it['label'], _it['unit'], _g, _vals,
                       f"{_it['cells'][_g]['低成本']}:{_it['cells'][_g]['高成本'].split('!')[1]}", _fmt(_it['unit'], _vals))
                wb.defined_names[_key] = _DN(_key, attr_text=f"'Tokenomics_取數'!$G${r - 1}")
        elif _it['kind'] == 'single':
            _rg = _it.get('range') or {}
            _vals = [_rg.get('低', ''), _it['values'], _rg.get('高', '')] if _rg else ["", _it['values'], ""]
            _loc = _it['cell'] + (f"（低／高＝{_rg['cells']['低'].split('!')[1]}／{_rg['cells']['高'].split('!')[1]}；{_rg.get('def') or ''}）" if _rg else "")
            _key = f"TK_{_short(_nm)}"
            _tkrow(_key, _nm, _it['label'], _it['unit'], "—", _vals, _loc, _fmt(_it['unit'], _vals))
            wb.defined_names[_key] = _DN(_key, attr_text=f"'Tokenomics_取數'!$G${r - 1}")
        else:
            raise SystemExit(f'Tokenomics_取數：不支援的快照型態 {_it["kind"]}（{_nm}）')

for s in wb.worksheets:
    s.sheet_view.showGridLines = False
    if s.title not in ("導覽", "來源", "摘要"):
        s.freeze_panes = "C5"

_order = ["導覽", "摘要", "輸入與假設", "各期收支", "季度追蹤", "運營_產能與收入", "運營_站點"] + (["每MW經濟性"] if FL else []) + ["資產負債_既有債務", "資產負債_新債與新股",
          "資產負債_租賃承諾", "損益", "評價_DCF與目標價", "評價_反向DCF", "評價_可比公司", "檢查_連動", "檢查_版本紀錄", "來源"] + (["Tokenomics_取數"] if TK else [])
wb._sheets = [wb[n] for n in _order] + [w for w in wb.worksheets if w.title not in _order]
_tab = {"摘要": "C00000", "輸入與假設": "1F3864", "各期收支": "0F6B4C", "季度追蹤": "0F6B4C", "運營_產能與收入": "0F5C61", "運營_站點": "0F5C61",
        "資產負債_既有債務": "5B5778", "資產負債_新債與新股": "5B5778", "資產負債_租賃承諾": "5B5778",
        "損益": "C4A35A", "評價_DCF與目標價": "9F1239", "評價_可比公司": "9F1239", "評價_反向DCF": "9F1239", "檢查_連動": "808080", "檢查_版本紀錄": "808080", "來源": "808080", **({"Tokenomics_取數": "808080"} if TK else {}), **({"每MW經濟性": "0F5C61"} if FL else {})}
for _n, _c in _tab.items():
    wb[_n].sheet_properties.tabColor = _c
for _ws in wb.worksheets:
    apply_outline(_ws)
import json as _json
# 輸出位置相對於 repo 根目錄：預設 out/；可用第 1 個參數指定 xlsx 路徑（outline.json 一律寫在 out/）
_OUT_DIR = _osrv.path.join(_osrv.path.dirname(_osrv.path.abspath(__file__)), "out")
_osrv.makedirs(_OUT_DIR, exist_ok=True)
_xlsx = _sysv.argv[1] if len(_sysv.argv) > 1 else _osrv.path.join(_OUT_DIR, f"{CO['meta']['updateDate'].replace('-', '')}_CoreWeave收支模型_v{VLOG_X[-1][0][1:].replace('.', '_')}.xlsx")
_json.dump({k: v for k, v in OUTLINE.items()}, open(_osrv.path.join(_OUT_DIR, "outline.json"), "w"), ensure_ascii=False)
# v4.5（5a-1）：說明文字中的期間佔位符（«YTD»、«VMD»…）換成目前日曆的字樣（calendar_q.tokens；目前日曆下與 v4.4 文字相同）
for _ws in wb:
    for _row in _ws.iter_rows():
        for _c in _row:
            if isinstance(_c.value, str) and '«' in _c.value: _c.value = _calq.fill(_c.value, CAL['tokens'])
wb.save(_xlsx)
print("saved", _xlsx)
