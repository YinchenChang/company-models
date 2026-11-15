# MiniMax 收支模型 v0.2.2：PL 頁（版面 v1.3 標準損益表）與 Coord 頁（三層協調：讀取層級、C0 實績）。
# 兩頁都只列示：不被任何既有計算頁引用；新增後既有工作表的數值不變（tools/xlsx_diff.py 逐格驗證）。
# Python 只產生結構與公式；數值由 LibreOffice 重算。層級值取自 repo 內快照檔（wattonomics/data），寫成藍字並附版本與雜湊。
# 依據：docs/workorders/20261114_第二批13家_v1.3與三層協調.md、docs/plan/三層基準協調規則.md v1.4.2 C0 第 2 點、
#       docs/plan/個股HTML版面規格.md v1.3 D 節（標準損益表）。
import json
import os

from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.workbook.defined_name import DefinedName

BOLD = Font(bold=True); BLUE = Font(color='0000CC'); GREEN = Font(color='007A33'); NOTE = Font(color='5F605F', size=9)
HDR = PatternFill('solid', fgColor='DDE6F0'); SEC = PatternFill('solid', fgColor='F2F2F2')
WRAP = Alignment(wrap_text=True, vertical='top')
M = '#,##0.0;[Red]-#,##0.0'; P = '0.0%;[Red]-0.0%'
HERE = os.path.dirname(os.path.abspath(__file__))
WATT_DIR = os.path.join(os.path.dirname(HERE), 'wattonomics', 'data')
PL_YEARS = list(range(2023, 2031))            # C–J＝FY2023–FY2030
PL_COL = {y: 'CDEFGHIJ'[i] for i, y in enumerate(PL_YEARS)}


def _name(wb, nm, ref):
    wb.defined_names[nm] = DefinedName(nm, attr_text=ref)


def read_watt():
    cur = open(os.path.join(WATT_DIR, 'CURRENT_SNAPSHOT'), encoding='utf-8').read().strip()
    d = json.load(open(os.path.join(WATT_DIR, cur), encoding='utf-8'))
    players = sorted({k[len('WIF_PlayerGW_YE_'):].rsplit('_', 1)[0] for k in d['names']
                      if k.startswith('WIF_PlayerGW_YE_') and '_SiteOnly_' not in k})
    blob = json.dumps(d, ensure_ascii=False)
    return dict(file=d['source']['file'], version=d['source']['version'], sha=d['source']['sha256'], snap=cur,
                names=d['names'], players=players, hits=sum(blob.count(w) for w in ('MiniMax', 'minimax', '稀宇')))


# ───────────────────────── PL：標準損益表（非 IFRS 經調整口徑） ─────────────────────────
def build_pl(wb, REG, YCOL):
    S = lambda sid: f"SRC_MM!$E${REG[sid]['row']}"                     # noqa: E731
    Pn = lambda pid, y: f"PnL!${YCOL[y]}${REG[pid]['row']}"            # noqa: E731
    F = lambda fid, y: f"Funding!${YCOL[y]}${REG[fid]['row']}"         # noqa: E731
    ws = wb.create_sheet('PL')
    ws['A1'] = 'PL — 標準損益表（版面 v1.3 D 節；非 IFRS 經調整口徑；FY2023–FY2025 實績、FY2026 為 1H 實績＋2H 預測、FY2027–FY2030 基準情境）'; ws['A1'].font = Font(bold=True, size=12)
    ws['A2'] = ('口徑：非 IFRS（經調整）＝IFRS 扣除股份支付、優先股公允價值變動與上市開支；公司業績公告與市場（賣方、新聞）主要使用此口徑。'
                '歷史欄取 SRC_MM（招股書、FY2025 年度業績公告；一手）；預測欄取 PnL 頁（本頁只引用，不新增假設）。單位 $M（EPS 為美元）。')
    ws['A2'].alignment = WRAP
    for c, h in enumerate(['項目', '單位'] + [f'FY{y}' for y in PL_YEARS] + ['說明'], 1):
        x = ws.cell(4, c, h); x.font = BOLD; x.fill = HDR
    ws.cell(5, 1, '期間性質')
    for y in PL_YEARS:
        ws.cell(5, 3 + PL_YEARS.index(y), '實績' if y <= 2025 else ('1H 實績＋2H 預測' if y == 2026 else '預測（基準）'))
    hist = {2023: dict(rev=S('S114'), cogs=S('S115'), sd=S('S117'), adm=S('S119'), rd=S('S121'), oth=S('S123'), fin=S('S125'), imp=S('S128'),
                       fv=S('S131'), nl=S('S134'), sbc=S('S136'), lst=S('S138'), adj=S('S140')),
            2024: dict(rev=S('S001'), cogs=S('S116'), sd=S('S118'), adm=S('S120'), rd=S('S122'), oth=S('S124'), fin=S('S126'), imp=S('S129'),
                       fv=S('S132'), nl=S('S135'), sbc=S('S137'), lst=S('S139'), adj=S('S141')),
            2025: dict(rev=S('S002'), cogs=S('S019'), sd=S('S021'), adm=S('S022'), rd=S('S023'), oth=S('S032'), fin=S('S127'), imp=S('S130'),
                       fv=S('S133'), nl=S('S036'), sbc=S('S030'), lst=S('S037'), adj=S('S034'))}
    R = {}
    rows = [
        ('rev', '營收', '$M', 'PL_Rev', lambda y, c: f"={hist[y]['rev']}" if y <= 2025 else f"={Pn('P01', y)}", M,
         '歷史＝公司揭露總營收；預測＝PnL P01（企業端＋消費端）'),
        ('yoy', '營收年增率（YoY）', '比例', 'PL_RevYoY', lambda y, c: None if y == 2023 else f"={c}{{rev}}/{prev(c)}{{rev}}-1", P,
         'FY2023 的前一年（FY2022）營收為 0（招股書），年增率無定義，留空'),
        ('cogs', '營業成本（銷售成本）', '$M', 'PL_COGS', lambda y, c: f"={hist[y]['cogs']}" if y <= 2025 else f"={Pn('P05', y)}", M,
         '預測＝PnL P05（推論算力＋其他銷售成本）'),
        ('gp', '毛利', '$M', 'PL_GP', lambda y, c: f"={c}{{rev}}-{c}{{cogs}}", M, ''),
        ('gpm', '毛利率', '比例', 'PL_GPM', lambda y, c: f"={c}{{gp}}/{c}{{rev}}", P, ''),
        ('opex', '營業費用（不含股份支付與上市開支）', '$M', 'PL_Opex',
         lambda y, c: (f"={hist[y]['sd']}+{hist[y]['adm']}+{hist[y]['rd']}-{hist[y]['sbc']}-{hist[y]['lst']}" if y <= 2025
                       else f"={Pn('P15', y)}-{Pn('P13', y)}"), M,
         '歷史＝銷售及分銷＋行政＋研發 − 股份支付 − 上市開支（非 IFRS 調節項全數歸營業費用）；預測＝PnL P15 − P13（營業費用合計扣 SBC）'),
        ('oi', '營業利益（非 IFRS）', '$M', 'PL_OI', lambda y, c: f"={c}{{gp}}-{c}{{opex}}", M, '負＝虧損'),
        ('oim', '營業利益率', '比例', 'PL_OIM', lambda y, c: f"={c}{{oi}}/{c}{{rev}}", P, ''),
        ('oth', '其他收入（淨額）', '$M', 'PL_Oth',
         lambda y, c: f"={hist[y]['oth']}-{hist[y]['fin']}-{hist[y]['imp']}" if y <= 2025 else f"={Pn('P17', y)}", M,
         '歷史＝其他收入及收益淨額 − 財務成本 − 金融資產減值；預測＝PnL P17（利息與理財收益；本模型未建財務成本）'),
        ('pbt', '稅前利益（非 IFRS）', '$M', 'PL_PBT', lambda y, c: f"={c}{{oi}}+{c}{{oth}}", M, ''),
        ('tax', '所得稅', '$M', 'PL_Tax',
         lambda y, c: (f"=({hist[y]['rev']}-{hist[y]['cogs']}-{hist[y]['sd']}-{hist[y]['adm']}-{hist[y]['rd']}+{hist[y]['oth']}-{hist[y]['fv']}-{hist[y]['fin']}-{hist[y]['imp']})+{hist[y]['nl']}"
                       if y <= 2025 else f"={c}{{pbt}}-{Pn('P18', y)}"), M,
         '歷史＝IFRS 稅前（各損益項加總）− IFRS 淨利（淨虧損加回）；三年皆為 0（虧損期無所得稅開支）。預測＝稅前 − PnL P18（本模型經調整淨利不計稅，故為 0）'),
        ('ni', '稅後淨利（非 IFRS 經調整）', '$M', 'PL_NI', lambda y, c: f"=-{hist[y]['adj']}" if y <= 2025 else f"={Pn('P18', y)}", M,
         '歷史＝公司揭露經調整淨虧損（取負號）；預測＝PnL P18'),
        ('nim', '淨利率', '比例', 'PL_NIM', lambda y, c: f"={c}{{ni}}/{c}{{rev}}", P, ''),
        ('eps', 'EPS（非 IFRS，完全稀釋；美元）', '$/股', 'PL_EPS',
         lambda y, c: '未揭露' if y <= 2025 else f"={c}{{ni}}/({S('S051')}*(1+{F('F14', y)}))", '#,##0.00;[Red]-#,##0.00',
         '歷史：公司只揭露 IFRS 每股虧損（含優先股公允價值變動），非 IFRS 每股未揭露，不補編。預測＝經調整淨利 ÷（配售後股數 S051 ×（1＋累計新股稀釋 Funding F14））；可轉債基準以現金贖回（I23＝0），不計轉股'),
        ('niyoy', '淨利年增率', '比例', 'PL_NIYoY', lambda y, c: None if y == 2023 else f"={c}{{ni}}/{prev(c)}{{ni}}-1", P,
         '虧損期的年增率為「虧損變動」：正＝虧損擴大（分母為負）'),
    ]

    def prev(c):
        return chr(ord(c) - 1)
    r0 = 6
    for i, x in enumerate(rows):
        R[x[0]] = r0 + i
    for i, (key, lab, unit, nm, fn, fmt, note) in enumerate(rows):
        r = R[key]
        ws.cell(r, 1, lab); ws.cell(r, 2, unit); ws.cell(r, 11, note).alignment = WRAP
        for y in PL_YEARS:
            c = PL_COL[y]
            f = fn(y, c)
            if f is None:
                continue
            if isinstance(f, str) and f.startswith('='):
                f = f.format(**R) if '{' in f else f
                # 上一行 format 會把 {rev} 換成列號
            x = ws.cell(r, 3 + PL_YEARS.index(y), f)
            x.number_format = fmt
            if isinstance(f, str) and '!' in f:
                x.font = GREEN
        if key in ('rev', 'gp', 'oi', 'pbt', 'ni', 'eps'):
            ws.cell(r, 1).font = BOLD
        _name(wb, nm, f"PL!$C${r}:$J${r}")
    # 核對列（應為 0；只列示）
    r = r0 + len(rows) + 1
    ws.cell(r, 1, '核對（應為 0；只列示，不進 Checks 頁）').font = BOLD
    for c in range(1, 12):
        ws.cell(r, c).fill = SEC
    k1 = r + 1
    ws.cell(k1, 1, '歷史：稅前 − 所得稅 − 稅後淨利（以非 IFRS 各項重組 − 公司揭露經調整淨虧損；驗證歷史列）'); ws.cell(k1, 2, '$M')
    for y in (2023, 2024, 2025):
        c = PL_COL[y]
        ws.cell(k1, 3 + PL_YEARS.index(y), f"=ROUND({c}{R['pbt']}-{c}{R['tax']}-{c}{R['ni']},6)").number_format = '0.000'
    k2 = r + 2
    ws.cell(k2, 1, '預測：營業利益＋其他收入 − PnL 經調整淨利 P18（恆等式）'); ws.cell(k2, 2, '$M')
    for y in range(2026, 2031):
        c = PL_COL[y]
        ws.cell(k2, 3 + PL_YEARS.index(y), f"=ROUND({c}{R['oi']}+{c}{R['oth']}-{Pn('P18', y)},6)").number_format = '0.000'
    k3 = r + 3
    ws.cell(k3, 1, '預測：營業利益 − PnL 營業虧損 P16 − 股份支付 P13（恆等式）'); ws.cell(k3, 2, '$M')
    for y in range(2026, 2031):
        c = PL_COL[y]
        ws.cell(k3, 3 + PL_YEARS.index(y), f"=ROUND({c}{R['oi']}-{Pn('P16', y)}-{Pn('P13', y)},6)").number_format = '0.000'
    k4 = r + 4
    ws.cell(k4, 1, '核對合計（絕對值和）'); ws.cell(k4, 1).font = BOLD
    ws.cell(k4, 3, f"=SUMPRODUCT(ABS(C{k1}:J{k3}))").number_format = '0.000'
    _name(wb, 'PL_CheckSum', f"PL!$C${k4}")
    ws.cell(k4 + 2, 1, '對照（不入表）：FY2025 模型校準年（PnL D 欄）經調整淨利 − 公司揭露；差額＝財務成本與減值等本模型未建項目（同 PnL P19）')
    ws.cell(k4 + 2, 3, f"=PnL!$D${REG['P18']['row']}+{S('S034')}").number_format = M
    for col, w in zip('ABCDEFGHIJK', [44, 8, 11, 11, 11, 11, 11, 11, 11, 11, 90]):
        ws.column_dimensions[col].width = w
    ws.freeze_panes = 'C6'
    return R


# ───────────────────────── Coord：三層協調（讀取層級、C0 實績） ─────────────────────────
def build_coord(wb, REG, YCOL, snap_meta, n_tk):
    S = lambda sid: f"SRC_MM!$E${REG[sid]['row']}"                     # noqa: E731
    W = read_watt()
    ws = wb.create_sheet('Coord')
    ws['A1'] = 'Coord — 三層協調（v0.2.2；規則 v1.4.2）：讀取層級、C0（實績）公司實績 vs 層級隱含值'; ws['A1'].font = Font(bold=True, size=12)
    ws['A2'] = ('C0 比較基準＝公司法定揭露（港交所招股書、年度與中期業績公告）；層級側＝同口徑的層級隱含值（算力買方：Tokenomics 每 GW 營收 × Wattonomics 該玩家 GW、'
                'Wattonomics 該玩家資本支出與上線 GW）。Wattonomics 玩家名單沒有 MiniMax（中國只有合計列），因此各列層級側皆無值，判定「無可比實績（層級無此玩家）」。')
    ws['A3'] = '本頁只列示，不被任何計算頁引用。單位：$M、MW、GW。藍字＝快照值或參數，黑字＝公式。'
    for a in ('A2', 'A3'):
        ws[a].alignment = WRAP
    r = [5]
    P = {}

    def sec(t):
        x = ws.cell(r[0], 1, t); x.font = BOLD
        for c in range(1, 13):
            ws.cell(r[0], c).fill = SEC
        r[0] += 1

    def header(cols):
        for c, h in enumerate(cols, 1):
            x = ws.cell(r[0], c, h); x.font = BOLD; x.fill = HDR
        r[0] += 1

    def row(key, label, vals, unit='', note='', fmt='#,##0.0', start=3):
        n = r[0]
        ws.cell(n, 1, label); ws.cell(n, 2, unit)
        for k, v in enumerate(vals):
            if v is None:
                continue
            x = ws.cell(n, start + k, v)
            is_f = isinstance(v, str) and v.startswith('=')
            if isinstance(v, (int, float)):
                x.font = BLUE; x.number_format = fmt
            elif is_f:
                x.number_format = fmt
                if '!' in v:
                    x.font = GREEN
        if note:
            ws.cell(n, 12, note).alignment = WRAP
        P[key] = n
        r[0] += 1
        return n

    # A. 參數
    sec('A. 參數（Decision；規則 v1.4.2 C0 第 2 點與 tools/coord/README.md）')
    header(['項目', '單位', '值', '', '', '', '', '', '', '', '', '說明'])
    row('th', '差距門檻（對帳中）', [0.15], '比例', '差距超過 ±15%，且兩個讀數同號，才標「對帳中」', fmt='0%')
    row('qx', '季化倍數（最近一季 × 本值＝年化速率）', [4], '倍', '第二讀數＝最近一季 × 本值 vs 層級 CY2026', fmt='0')
    row('m6', '2026 已公布月數', [6], '月', '2026 中期業績公告（2026-08-26）涵蓋 2026-01～06；公司未發布季度業績', fmt='0')
    row('m12', '每年月數', [12], '月', '時差換算', fmt='0')

    # B. 讀取層級
    r[0] += 1
    sec('B. 讀取層級（同源，不判定）：本模型實際讀的層級、版本與名稱；自層級讀入的值只列在此節，不進 C0')
    header(['項目', '單位', '內容', '', '', '', '', '', '', '', '', '說明'])
    row('tkv', 'Tokenomics（TK_Link 快照）', [f"{snap_meta['version']}｜{snap_meta['file']}｜提交 {snap_meta['commit'][:7]}"], '',
        f"快照檔 minimax/data/{snap_meta['snap']}（SHA-256 {snap_meta['file_sha256']}；讀取日 {snap_meta['read_date']}）。Tokenomics master 現行為 v22.0；"
        '本輪不換快照（工作單：模型仍讀舊快照者寫明版本、不在本輪換）。TK 值只作算力物理（每 GW 產能、利用率、持有成本、全成本 $/M）與對照，不回饋 Tokenomics')
    row('tkn', 'Tokenomics 取值列數（TK_Link TK01–TK12；另列 IF_HoldEcon 全 15 欄）', [f'=COUNTA(TK_Link!$A$5:$A${4 + n_tk})'], '列',
        '讀入名稱：IF_TokGW_Luna、IF_TokGW_Sol、IF_Util、IF_HoldEcon（另列全 15 欄）、IF_FullCost_Luna、IF_FullCost_Sol；IF_PriceRef_Luna、IF_PriceRef_Sol 只用於 Inputs 說明文字（名稱檔 minimax/data/tokenomics_names.txt）', fmt='0')
    row('wtv', 'Wattonomics', [f"本模型不讀 Wattonomics；本頁對照用快照 {W['version']}｜{W['file']}"], '',
        f"快照檔 wattonomics/data/{W['snap']}（SHA-256 {W['sha']}）。Wattonomics 玩家名單（年底上線 GW 的玩家代號）：{'、'.join(W['players'])}；沒有 MiniMax，中國只有合計列 cn（Rystad 容量 × AI 占比）")
    row('wtr', 'Wattonomics 是否讀本模型', [f"否（快照全文出現 MiniMax／稀宇字樣 {W['hits']} 次）"], '',
        '同源檢查：層級沒有讀入本模型的任何值，本頁不存在同源校準列')
    row('sif', 'Siliconomics', ['不適用（算力買方，非硬體供應鏈個股）'], '', '算力買賣方讀 Tokenomics／Wattonomics（Project 規則第 3 條）')

    # C. C0（實績）
    r[0] += 1
    sec('C. C0（實績）：公司實績 vs 層級隱含值（規則 v1.4.2 C0 第 2 點；Andy 11-14 17:40 固定規格）')
    header(['項目', '單位', '公司實績（主）', '層級隱含（主）', '差距（主）', '公司實績（第二）', '層級隱含（第二）', '差距（第二）', '判定', '出處與口徑', '等級', '說明'])
    d1 = r[0]
    NA = '無可比實績（層級無此玩家）'
    row('c0_rev', '營收（IFRS 總營收；2026 已公布 1–6 月；第二讀數＝2026 Q2）',
        [f"={S('S004')}", None, '—', f"={S('S004')}*(1+{S('S076')})/(2+{S('S076')})", None, '—', NA,
         'SRC_MM S004（1H2026 營收 116.573；2026 中期業績公告 2026-08-26，港交所 2026082600680）；第二讀數 Q2＝1H26 ×（1＋季增）÷（2＋季增），季增 S076（81.8%，業績會轉述，InfoQ）',
         'Verified（主，一手）；第二讀數 Derived（季增為二手轉述，季營收未揭露）',
         '層級隱含＝Tokenomics 每 GW 營收 × Wattonomics MiniMax GW；Wattonomics 沒有 MiniMax 玩家（只有中國合計），無 GW 可乘，層級側無值。缺：Wattonomics 的 MiniMax 上線 GW（或中國合計拆到公司）。'
         '主讀數若有層級值應比 CY2026 × 6/12；第二讀數 Q2 × 4 是 5 月中的年化速率，成長期偏低於全年。2026-08 ARR 超過 8 億美元（S053）屬 run-rate、二手轉述，不入本列'],
        '$M', fmt='#,##0.0', start=3)
    row('c0_capex', '資本支出（1H2026 物業廠房設備＋預付款增加；第二讀數不適用）',
        [f"={S('S050')}-{S('S049')}", None, '—', None, None, None, NA,
         'SRC_MM S049、S050（2025-12-31 18.8、2026-06-30 341.1；FY2025 年度業績公告、2026 中期業績公告資產負債表）；口徑＝物業廠房設備＋預付款增額（公司未說明預付款用途；推測為算力採購）',
         'Verified（一手；用途 Assumed）',
         '層級對象＝Wattonomics 該玩家 AI 資本支出；無 MiniMax 玩家。期末餘額差只有半年一個讀數，不做第二讀數'],
        '$M', fmt='#,##0.0', start=3)
    row('c0_mw', '在役／簽約算力（MW）', [None, None, '—', None, None, None, '無可比實績（公司未揭露、層級無此玩家）',
                                     '公司未揭露自有或租用算力 MW（招股書、業績公告只揭露雲端服務支出與阿里雲採購上限）；本模型的 MW 由支出 ÷ 合約價推得（Compute 頁），屬模型值，不當實績',
                                     '—', '缺：公司算力 MW 揭露、Wattonomics MiniMax 玩家'], 'MW', start=3)
    row('c0_commit', '算力採購承諾（阿里雲關連交易 2026 年度上限）',
        [f"={S('S059')}", None, '—', None, None, None, NA,
         'SRC_MM S059（人民幣 21 億 ≈ 3 億美元；2026-08 關連交易上限修訂，媒體轉述）',
         'Interested-party（二手轉述公告）',
         '上限不是已發生支出；Wattonomics 無 MiniMax 承諾列。只列參考'],
        '$M', fmt='#,##0.0', start=3)
    d4 = r[0] - 1
    for n in range(d1, d4 + 1):
        for c in 'EH':
            ws[f'{c}{n}'].number_format = '0.0%'
    _name(wb, 'COORD_C0', f"Coord!$A${d1}:$K${d4}")
    DR = f"$I${d1}:$I${d4}"

    # 判定摘要
    r[0] += 1
    sec('判定摘要（只計 C 節）')
    n = r[0]
    labs = [('可判定項目數（對帳中、時差邊界或一致）', f'=COUNTIF({DR},"對帳中")+COUNTIF({DR},"時差邊界*")+COUNTIF({DR},"一致")'),
            ('對帳中項目數', f'=COUNTIF({DR},"對帳中")'),
            ('時差邊界（參考）項目數', f'=COUNTIF({DR},"時差邊界*")'),
            ('參考（同源校準）項目數', f'=COUNTIF({DR},"參考（同源*")'),
            ('無可比實績項目數', f'=COUNTIF({DR},"無可比實績*")')]
    for k, (a, f) in enumerate(labs):
        ws.cell(n + k, 1, a); ws.cell(n + k, 3, f).number_format = '0'
    _name(wb, 'COORD_C0Sum', f"Coord!$C${n}:$C${n + 4}")
    r[0] = n + 5
    ws.cell(r[0], 1, 'C0 結論').font = BOLD
    ws.cell(r[0], 3, ('MiniMax 的 C0 沒有可判定的項目，對帳中 0：Wattonomics 玩家名單沒有 MiniMax（中國只有合計列），層級沒有 MiniMax 的 GW、資本支出或承諾，'
                      '營收、資本支出、算力承諾三列都是「無可比實績（層級無此玩家）」；算力 MW 公司亦未揭露。層級也沒有讀入本模型的任何值（無同源列）。'
                      '往上回饋只有證據（C5）：若 Wattonomics 將中國合計拆到公司，可用本頁公司實績（1H2026 營收、物業廠房設備與預付款增額）作為錨點。'))
    ws.cell(r[0], 3).alignment = WRAP
    _name(wb, 'COORD_C0Verdict', f"Coord!$C${r[0]}")
    P['verdict'] = r[0]
    r[0] += 1

    # D. 參考（不判定）
    r[0] += 1
    sec('D. 參考（不判定）：模型值 vs 層級中國合計（兩邊都不是公司實績，只看量級）')
    header(['項目', '單位', '本模型 2026', 'Wattonomics 中國合計 2026 年底', '占比', '', '', '', '', '', '', '說明'])
    n = r[0]
    row('cn', '本模型 2026 年均總 MW（Compute C21）÷ 1,000 vs Wattonomics WIF_PlayerGW_YE_cn_2026',
        [f"=Compute!${YCOL[2026]}${REG['C21']['row']}/1000", W['names'].get('WIF_PlayerGW_YE_cn_2026'), f'=C{n}/D{n}'], 'GW',
        '本模型 MW 是支出 ÷ 合約價推得的模型值（非實績）；中國合計為 Rystad 容量 × AI 占比（Wattonomics S1c）。只示意 MiniMax 在中國 AI 算力的量級，不判定', fmt='#,##0.000')
    ws[f'E{n}'].number_format = '0.0%'
    for col, w in zip('ABCDEFGHIJKL', (58, 8, 20, 20, 12, 18, 18, 14, 28, 60, 30, 90)):
        ws.column_dimensions[col].width = w
    ws.freeze_panes = 'C5'
    return P
